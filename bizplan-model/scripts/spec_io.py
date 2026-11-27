"""Validate the bundled business-plan format and load normalized local CSVs."""
from __future__ import annotations

from calendar import monthrange
from collections import defaultdict
import copy
import csv
from datetime import date
from graphlib import CycleError, TopologicalSorter
import json
import math
from pathlib import Path
import re
import sys
import unicodedata
import warnings


from tax_policy import normalize_tax, shift_month, warn_zero_tax_payable


SCHEMA_PATH = Path(__file__).resolve().parents[1] / "assets" / "spec.schema.json"
OPERATING = ("매출", "매출원가", "판관비")
IS_SECTIONS = (*OPERATING, "영업외수익", "영업외비용", "법인세")
BS_SECTIONS = ("유동자산", "비유동자산", "유동부채", "비유동부채", "자본")
HEADERS = {
    "is_monthly": ("month", "account", "section", "amount"),
    "bs_monthly": ("month", "account", "section", "amount"),
    "debt": ("facility", "lender", "type", "balance", "rate", "maturity", "note"),
    "debt_schedule": ("facility", "month", "principal"),
    "journal": ("date", "account", "counterparty", "debit", "credit", "memo"),
    "revenue_items": ("item", "line", "start", "end", "amount", "probability", "note"),
    "headcount": ("role", "count", "monthly_cost", "start", "end"),
    "fixed_assets": ("asset_class", "cost", "accumulated_dep", "method", "rate_or_life"),
    "sales_volume": ("month", "line", "quantity", "price"),
    "metrics_monthly": ("month", "metric", "line", "value"),
}
OPTIONAL_HEADERS = {"debt": {"at_maturity"}, "headcount": {"raise_from"}, "sales_volume": {"activity"}}
UNIQUE_KEYS = {
    "is_monthly": ("month", "section", "account"),
    "bs_monthly": ("month", "section", "account"),
    "debt": ("facility",),
    "debt_schedule": ("facility", "month"),
    "revenue_items": ("line", "item"),
    "headcount": ("role", "start", "end"),
    "sales_volume": ("line", "month"),
    "metrics_monthly": ("metric", "line", "month"),
}
INTEGER_COLUMNS = {
    "amount", "balance", "principal", "debit", "credit", "monthly_cost",
    "cost", "accumulated_dep",
}
MONTH_PATTERN = r"(?!0000)[0-9]{4}-(?:0[1-9]|1[0-2])"


def fail(location, message):
    raise ValueError(f"{location}: {message}")


def read_json(path):
    def unique_object(pairs):
        result = {}
        for key, value in pairs:
            if key in result:
                fail(path, f"중복 JSON 키 {key!r}")
            result[key] = value
        return result

    try:
        return json.loads(path.read_text(encoding="utf-8-sig"),
                          object_pairs_hook=unique_object,
                          parse_constant=lambda token: fail(path, f"유효하지 않은 숫자 {token}"))
    except (OSError, UnicodeError, json.JSONDecodeError) as exc:
        raise ValueError(f"{path}: {exc}") from exc


def resolve(rule, definitions):
    if "$ref" in rule:
        return definitions[rule["$ref"].removeprefix("#/$defs/")]
    return rule


def validate(value, rule, definitions, location):
    """Apply defaults and the keywords used by our local schema (no network resolver)."""
    rule = resolve(rule, definitions)
    if "anyOf" in rule:
        errors = []
        for branch in rule["anyOf"]:
            try:
                value = validate(copy.deepcopy(value), branch, definitions, location)
                break
            except ValueError as exc:
                errors.append(str(exc))
        else:
            fail(location, "허용 형식과 일치하지 않습니다. " + " / ".join(errors))
    if "oneOf" in rule:
        # Each branch has a unique constant discriminator, either driver or method.
        if not isinstance(value, dict):
            fail(location, "라인은 JSON 객체여야 합니다")
        branches = rule["oneOf"]
        discriminator = next(k for k, v in branches[0]["properties"].items() if "const" in v)
        choices = {b["properties"][discriminator]["const"]: b for b in branches}
        selected = value.get(discriminator)
        if not isinstance(selected, str) or selected not in choices:
            fail(f"{location}.{discriminator}", f"{', '.join(choices)} 중 하나를 지정하세요")
        return validate(value, choices[selected], definitions, location)

    kind = rule.get("type")
    if kind == "object":
        if not isinstance(value, dict):
            fail(location, "JSON 객체가 필요합니다")
        if len(value) < rule.get("minProperties", 0):
            fail(location, "하나 이상의 필드를 지정하세요")
        properties = rule.get("properties", {})
        for key in rule.get("required", []):
            if key not in value:
                fail(location, f"필수 필드 {key!r} 누락")
        for key, child in properties.items():
            resolved = resolve(child, definitions)
            if key not in value and "default" in resolved:
                value[key] = copy.deepcopy(resolved["default"])
        for key, item in value.items():
            if "propertyNames" in rule:
                validate(key, rule["propertyNames"], definitions, f"{location}[{key!r}]")
            child = properties.get(key, rule.get("additionalProperties", True))
            if child is False:
                fail(location, f"정의되지 않은 필드 {key!r}")
            if isinstance(child, dict):
                value[key] = validate(item, child, definitions, f"{location}.{key}")
    elif kind == "array":
        if not isinstance(value, list):
            fail(location, "JSON 배열이 필요합니다")
        if len(value) < rule.get("minItems", 0) or len(value) > rule.get("maxItems", math.inf):
            fail(location, f"항목 수 {len(value)}개가 허용 범위를 벗어났습니다")
        for idx, item in enumerate(value):
            value[idx] = validate(item, rule["items"], definitions, f"{location}[{idx}]")
        if rule.get("uniqueItems") and len(set(value)) != len(value):
            fail(location, "중복 항목")
    elif kind == "string":
        if not isinstance(value, str) or not re.fullmatch(rule.get("pattern", r"[\s\S]*"), value):
            fail(location, f"문자열 형식 오류 {value!r}")
    elif kind == "boolean":
        if type(value) is not bool:
            fail(location, "true 또는 false가 필요합니다")
    elif kind in ("number", "integer"):
        if type(value) not in (float, int) or abs(value) > sys.float_info.max or not math.isfinite(value):
            fail(location, "유한한 JSON 숫자가 필요합니다")
        if kind == "integer":
            if int(value) != value:
                fail(location, "정수가 필요합니다")
            value = int(value)
        if value < rule.get("minimum", -math.inf) or value > rule.get("maximum", math.inf) or value <= rule.get("exclusiveMinimum", -math.inf):
            fail(location, f"허용 범위를 벗어난 숫자 {value!r}")
    if "enum" in rule and value not in rule["enum"]:
        fail(location, f"{rule['enum']} 중 하나를 지정하세요")
    if "const" in rule and value != rule["const"]:
        fail(location, f"{rule['const']!r} 값이 필요합니다")
    return value


def parse_cell(stem, column, raw, where):
    raw = raw.strip()
    if not raw and column not in ("note", "memo", "counterparty"):
        fail(where, "빈 필수 값")
    if column == "maturity" and re.fullmatch(r"[0-9]{4}-[0-9]{2}-[0-9]{2}",raw):
        try:
            date.fromisoformat(raw)
        except ValueError:
            fail(where, f"존재하지 않는 만기일 {raw!r}")
        raw=raw[:7]
    if column in ("month", "maturity", "start", "end", "raise_from"):
        if not re.fullmatch(MONTH_PATTERN, raw):
            fail(where, f"YYYY-MM 월 형식 오류 {raw!r}")
    elif column == "date":
        if not re.fullmatch(r"[0-9]{4}-[0-9]{2}-[0-9]{2}", raw):
            fail(where, f"YYYY-MM-DD 날짜 형식 오류 {raw!r}")
        try:
            date.fromisoformat(raw)
        except ValueError:
            fail(where, f"존재하지 않는 날짜 {raw!r}")
    elif column in INTEGER_COLUMNS:
        if not re.fullmatch(r"[+-]?[0-9]+", raw):
            fail(where, f"원 단위 금액은 정수로 적으세요 {raw!r}")
        value = int(raw)
        if abs(value) > sys.float_info.max:
            fail(where, "Excel 숫자 범위를 초과했습니다")
        if value < 0 and not (column == "amount" and stem in ("is_monthly", "bs_monthly")):
            fail(where, "0 이상이어야 합니다")
        return value
    elif column in ("rate", "probability", "rate_or_life", "count", "quantity", "price", "activity", "value"):
        if not re.fullmatch(r"[+-]?(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[eE][+-]?[0-9]+)?", raw):
            fail(where, f"십진수 형식 오류 {raw!r}")
        value = float(raw)
        if not math.isfinite(value) or value < 0 or (column in ("rate", "probability") and value > 1):
            fail(where, f"허용 범위를 벗어난 숫자 {raw!r}")
        return value
    elif column == "at_maturity" and raw not in ("repay", "rollover"):
        fail(where, "at_maturity는 repay 또는 rollover여야 합니다")
    return raw


def read_table(path):
    stem = path.stem
    columns = HEADERS[stem]
    optional = OPTIONAL_HEADERS.get(stem, set())
    rows = []
    seen = set()
    try:
        with path.open(encoding="utf-8-sig", newline="") as stream:
            reader = csv.reader(stream, strict=True)
            header = next(reader, [])
            names = set(header)
            if len(header) != len(names) or not set(columns) <= names <= set(columns) | optional:
                fail(path, f"열 이름을 확인하세요. 필수 {','.join(columns)}, 선택 {sorted(optional)}, 입력 {header}")
            for cells in reader:
                where = f"{path.name}:{reader.line_num}"
                if len(cells) != len(header):
                    fail(where, f"{len(header)}열 필요, {len(cells)}열 입력")
                row = {key: parse_cell(stem, key, raw, f"{where}.{key}")
                       for key, raw in zip(header, cells) if key not in optional or raw.strip()}
                if "start" in row and row["start"] > row["end"]:
                    fail(where, "start가 end보다 늦습니다")
                if stem in UNIQUE_KEYS:
                    key = tuple(row[c] for c in UNIQUE_KEYS[stem])
                    if key in seen:
                        fail(where, f"중복 행 키 {key}")
                    seen.add(key)
                if stem in ("is_monthly", "bs_monthly"):
                    allowed = IS_SECTIONS if stem == "is_monthly" else BS_SECTIONS
                    if row["section"] not in allowed:
                        fail(where, f"section은 {allowed} 중 하나여야 합니다")
                    if stem == "is_monthly" and row["section"] in ("매출원가", "판관비", "영업외비용", "법인세") and row["amount"] < 0:
                        fail(where, "비용은 양수 또는 0으로 적으세요")
                if stem == "debt" and row["type"] not in ("term", "revolver", "bond", "lease"):
                    fail(where, f"알 수 없는 약정 유형 {row['type']!r}")
                if stem == "fixed_assets":
                    if row["accumulated_dep"] > row["cost"]:
                        fail(where, "상각누계액이 취득원가보다 큽니다")
                    life = row["rate_or_life"]
                    if row["method"] == "straight_line":
                        if life < 1 or not life.is_integer():
                            fail(where, "정액법 내용연수는 양의 정수 개월입니다")
                        row["rate_or_life"] = int(life)
                    elif row["method"] != "declining" or life > 1:
                        fail(where, "method와 rate_or_life를 확인하세요 (declining은 연 0~1)")
                rows.append(row)
    except (OSError, UnicodeError, csv.Error) as exc:
        raise ValueError(f"{path}: {exc}") from exc
    return rows


def indexed(rows, keys, where):
    result = {}
    for index, row in enumerate(rows):
        key = tuple(row[k] for k in keys)
        if key in result:
            fail(f"{where}[{index}]", f"중복 키 {key!r}")
        result[key] = row
    return result


def next_month(month):
    year, number = map(int, month.split("-"))
    return f"{year + (number == 12):04d}-{number % 12 + 1:02d}"


def check_actuals(spec, data):
    first, last = spec["first_month"], spec["last_actual"]
    historical_months = set()
    month = first
    while month <= last:
        historical_months.add(month)
        month = next_month(month)
    for stem in ("is_monthly", "bs_monthly"):
        rows = data[stem]
        if stem == "bs_monthly" and not rows:
            continue
        observed = {row["month"] for row in rows}
        extra = observed - historical_months
        if extra:
            fail(f"{stem}.csv", f"실적 기간 밖의 월 {min(extra)}")
        missing = historical_months - observed
        if stem == "bs_monthly" and spec["statements"]["bs"] and last not in observed:
            fail("bs_monthly.csv", f"실적 마감 월 {last}의 월말 잔액 누락")
        if missing and stem == "is_monthly":
            fail(f"{stem}.csv", f"월 실적 누락 {min(missing)}")
    balances = {}
    for row in data["bs_monthly"]:
        sign = 1 if row["section"] in BS_SECTIONS[:2] else -1
        month = row["month"]
        balances[month] = balances.get(month, 0) + sign * row["amount"]
    for month, difference in balances.items():
        if difference:
            fail("bs_monthly.csv", f"{month} 자산과 부채·자본 차이 {difference:,}원")


def warn_working_capital(spec, data):
    """Rebuild cutoff balances with the same BS groups and activity bases as the forecast."""
    if not spec["statements"]["bs"]:
        return
    from finance_blocks import _bucket

    cutoff = spec["last_actual"]
    previous = shift_month(cutoff, -1)
    balances = defaultdict(int)
    observed = set()
    for row in data["bs_monthly"]:
        month = row["month"]
        if month in (previous, cutoff):
            observed.add(month)
            balances[month, _bucket(row["section"], row["account"])] += row["amount"]
    activity = defaultdict(int)
    accounts = defaultdict(int)
    for row in data["is_monthly"]:
        if row["month"] == cutoff:
            activity[row["section"]] += row["amount"]
            accounts[row["section"], row["account"]] += row["amount"]
    wc = spec["working_capital"]
    days = 30 if wc["day_basis"] == "month30" else monthrange(*map(int, cutoff.split("-")))[1]
    cogs = activity["매출원가"]
    # Match by account and its section: identically named COGS and SG&A
    # accounts can belong to different lines. An unreported month contributes 0.
    line_amounts = {line["id"]: sum(accounts[line["section"], account] for account in line["accounts"])
                    for line in spec["cost_lines"]}
    inventory_base = (sum(line_amounts[key] for key in wc["dio_base_lines"])
                      if "dio_base_lines" in wc else cogs)
    cash_ids = set(wc.get("cash_revenue_lines", []))
    cash_revenue = sum(accounts["매출", account] for line in spec["revenue_lines"]
                       if line["id"] in cash_ids for account in line["accounts"])
    bases = [("dso", 9, activity["매출"] - cash_revenue), ("dio", 10, inventory_base)]
    if "dpo_base_lines" in wc:
        bases.append(("dpo", 14, sum(line_amounts[key] for key in wc["dpo_base_lines"])))
    elif wc["ap_base"] == "cogs":
        bases.append(("dpo", 14, cogs))
    elif previous in observed:
        bases.append(("dpo", 14, cogs + balances[cutoff, 10] - balances[previous, 10]))
    else:
        warnings.warn(f"working_capital.dpo: {previous} 재고 원천이 없어 매입액 기준 검사를 생략합니다",
                      UserWarning, stacklevel=2)
    for key, group, base in bases:
        actual = balances[cutoff, group]
        estimated = max(0, base) * wc[key] / days
        if abs(estimated - actual) > abs(actual) * .05:
            warnings.warn(f"working_capital.{key}: {cutoff} 계산 잔액 {estimated:,.2f}원, "
                          f"실제 잔액 {actual:,}원 (5% 초과)", UserWarning, stacklevel=2)


def check_lines(spec, data):
    revenue = spec["revenue_lines"]
    lines = revenue + spec["cost_lines"]
    indexed(lines, ("id",), "revenue_lines / cost_lines")
    ids = {line["id"] for line in lines}
    revenue_ids = {line["id"] for line in revenue}
    cost_ids = {line["id"] for line in spec["cost_lines"]}
    invalid_cash_lines = set(spec["working_capital"].get("cash_revenue_lines", [])) - revenue_ids
    if invalid_cash_lines:
        fail("working_capital.cash_revenue_lines", f"정의된 매출 라인 ID가 필요합니다: {sorted(invalid_cash_lines)}")
    for key in ("dio_base_lines", "dpo_base_lines"):
        invalid = set(spec["working_capital"].get(key, [])) - cost_ids
        if invalid:
            fail(f"working_capital.{key}", f"정의된 비용 라인 ID가 필요합니다: {sorted(invalid)}")
    observed = {(r["section"], r["account"]) for r in data["is_monthly"] if r["section"] in OPERATING}
    owners = {}
    for line in lines:
        for account in line["accounts"]:
            key = (line.get("section", "매출"), account)
            if key in owners:
                fail(line["id"], f"계정 {key}를 {owners[key]}에도 배정했습니다")
            if key not in observed:
                fail(line["id"], f"is_monthly.csv에 배정 계정 {key}가 없습니다")
            owners[key] = line["id"]
        if line.get("method") == "pct_of_revenue" and "base_lines" in line["params"]:
            p = line["params"]
            if "revenue_ids" in p and set(p["base_lines"]) != set(p["revenue_ids"]):
                fail(line["id"], "base_lines와 revenue_ids의 대상이 다릅니다")
            p["revenue_ids"] = p["base_lines"]
        if line.get("method") in ("pct_of_revenue", "mixed"):
            unknown = set(line["params"]["revenue_ids"]) - revenue_ids
            if unknown:
                fail(line["id"], f"revenue_ids에 없는 매출 ID {sorted(unknown)}")
        check_forecast(spec, line["params"].get("monthly", {}), f"{line['id']}.params.monthly")
    dependencies = {}
    for line in lines:
        params = line["params"]
        if line.get("driver") == "contracts" and params.get("churn") and not params["new_monthly_recurring"]:
            fail(line["id"], "contracts.churn은 반복 청구하는 신규 계약에만 적용합니다")
        if line.get("method") == "pct_of_line" or line.get("driver") == "pct_of_line":
            targets = line["params"]["line_ids"]
            unknown = set(targets) - (revenue_ids if "driver" in line else ids)
            if unknown:
                fail(line["id"], f"line_ids에 정의되지 않은 라인 ID {sorted(unknown)}")
            dependencies[line["id"]] = targets
    try:
        TopologicalSorter(dependencies).prepare()
    except CycleError as exc:
        fail("pct_of_line", "순환 라인 참조 " + " → ".join(exc.args[1]))
    drivers = {line["id"]: line for line in revenue}
    for row in data["sales_volume"]:
        line = drivers.get(row["line"])
        month = row["month"]
        if line is None or line["driver"] not in ("price_volume", "activity"):
            fail("sales_volume.csv", f"{row['line']!r}은 물량 기반 매출 ID여야 합니다")
        if not spec["first_month"] <= month <= spec["plan_end"]:
            fail("sales_volume.csv", f"{month}은 모델 기간 밖입니다")
        for key, value in line["params"].get("monthly", {}).get(month, {}).items():
            if row[key] != value:
                fail("sales_volume.csv", f"{row['line']} {month} {key}가 params.monthly와 다릅니다")
    item_probabilities = {(r["line"], r["item"]): r["probability"] for r in data["revenue_items"]}
    allowed_controls = {"contracts": {"new_multiplier", "exclude_unconfirmed", "exclude_items", "renewal_probability"},
                        "cohort": {"new_multiplier", "exclude_unconfirmed"},
                        "price_volume": {"price_multiplier"},
                        "activity": {"activity_multiplier", "yield_multiplier", "price_multiplier"},
                        "consumption": {"efficiency_multiplier", "price_multiplier", "subsidy_multiplier"}}
    controls_by_id = {line["id"]: line.get("driver",line.get("method")) for line in lines}
    for scenario, settings in spec["scenarios"].items():
        for key, controls in settings["drivers"].items():
            where = f"scenarios.{scenario}.drivers.{key}"
            if key not in controls_by_id:
                fail(where, "정의된 라인 ID가 필요합니다")
            driver = controls_by_id[key]
            unsupported = controls.keys() - allowed_controls.get(driver, set())
            if unsupported:
                fail(where, f"{driver}에서 지원하지 않는 설정 {sorted(unsupported)}")
            if "renewal_probability" in controls and "renewal" not in drivers[key]["params"]:
                fail(where, "renewal_probability를 사용하려면 contracts.params.renewal을 지정하세요")
            for name in controls.get("exclude_items", []):
                if (key, name) not in item_probabilities:
                    fail(where, f"revenue_items.csv의 해당 라인에 없는 계약 {name!r}")
                if item_probabilities[key, name] == 1:
                    fail(where, f"확률 1인 확정 계약 {name!r}은 제외 항목으로 지정할 수 없습니다")
    unassigned = observed - owners.keys()
    if unassigned:
        fail("is_monthly.csv", f"미배정 영업 계정 {sorted(unassigned)}")
    contract_ids = {line["id"] for line in revenue if line["driver"] == "contracts"}
    for row in data["revenue_items"]:
        if row["line"] not in contract_ids:
            fail("revenue_items.csv", f"{row['item']!r}의 line {row['line']!r}은 contracts 매출 ID여야 합니다")
    for case in spec["cases"]:
        if case["label"] in spec["scenarios"]:
            fail("cases", f"기본 케이스와 이름 중복 {case['label']!r}")
    indexed(spec["cases"], ("label",), "cases")
    scenarios_and_cases = dict(spec["scenarios"])
    scenarios_and_cases.update((case["label"], case) for case in spec["cases"])
    for name, case in scenarios_and_cases.items():
        for field in ("multipliers", "additions"):
            unknown = case[field].keys() - ids
            if unknown:
                fail(f"{name}.{field}", f"정의되지 않은 라인 ID {sorted(unknown)}")
            for line in lines:
                case[field].setdefault(line["id"], 1 if field == "multipliers" else 0)
    for i, ext in enumerate(spec["extensions"]):
        if ext["line_id"] not in ids:
            fail(f"extensions[{i}].line_id", f"정의되지 않은 라인 {ext['line_id']!r}")
        check_forecast(spec, ext["monthly"], f"extensions[{i}].monthly")


def check_forecast(spec, monthly, where):
    for month in monthly:
        if not spec["last_actual"] < month <= spec["plan_end"]:
            fail(where, f"{month}은 추정 기간 밖입니다")


def _account_caption(account):
    return re.sub(r"[\W_]+", "", unicodedata.normalize("NFKC", account).casefold())


def is_interest_account(account):
    return _account_caption(account) in {"이자비용", "지급이자", "interestexpense", "interest"}


def check_depreciation(spec, data):
    """Each designated account belongs to a historical-only, dedicated cost line."""
    location = "investment.depreciation_accounts"
    designated = set(spec["investment"]["depreciation_accounts"])
    accounts = {r["account"] for r in data["is_monthly"] if r["section"] in OPERATING[1:]}
    if designated - accounts:
        fail(location, f"비용 실적에 없는 상각 계정 {sorted(designated - accounts)}")
    for account in sorted(accounts - designated):
        caption = _account_caption(account)
        if (caption.endswith(("감가상각비", "무형자산상각비"))
                or caption in {"depreciation", "amortization", "depreciationexpense",
                               "amortizationexpense", "depreciationandamortization"}):
            fail(location, f"실적 상각 계정 {account!r}을 지정하세요")
    for line in spec["cost_lines"]:
        owned = set(line["accounts"])
        if not owned & designated:
            continue
        key = line["id"]
        if line["noncash"]:
            fail(key, "지정된 상각 계정은 비현금 비용으로 다시 가산할 수 없습니다")
        if owned - designated:
            fail(key, "상각 계정과 다른 비용 계정은 별도 라인으로 배정하세요")
        if line["method"] != "manual" or any(line["params"]["monthly"].values()):
            fail(key, "지정된 상각 라인은 manual 방식으로 두고 월별 계획 금액을 비워 두세요")
        if any(ext["line_id"] == key for ext in spec["extensions"]):
            fail(key, "지정된 상각 라인에 확장 조정을 적용할 수 없습니다")
        for case in [*spec["scenarios"].values(), *spec["cases"]]:
            if case["multipliers"][key] != 1 or case["additions"][key] != 0:
                fail(key, "지정된 상각 라인의 케이스 배수는 1, 가산액은 0이어야 합니다")


def warn_missing_basis(spec):
    assumptions = [(line["id"], line) for line in spec["revenue_lines"] + spec["cost_lines"]]
    assumptions.extend((group, spec[group]) for group in
                       ("wages", "investment", "working_capital", "financing", "tax"))
    assumptions.extend((f"scenarios.{name}", case) for name, case in spec["scenarios"].items())
    assumptions.extend((f"cases.{case['label']}", case) for case in spec["cases"])
    if any(key in spec["non_operating"] for key in ("income_monthly", "other_expense_monthly")):
        assumptions.append(("non_operating", spec["non_operating"]))
    missing = [name for name, obj in assumptions if not obj["basis"].strip()]
    if missing:
        warnings.warn("추정 근거(basis) 미기재: " + ", ".join(missing), UserWarning, stacklevel=2)


def check_non_operating(spec, rows):
    """Include every observed month in the denominator, even if its account total is zero."""
    months = {r["month"] for r in rows}
    for field, section in (("income_monthly", "영업외수익"), ("other_expense_monthly", "영업외비용")):
        where = f"non_operating.{field}"
        assumptions = spec["non_operating"]
        if field in assumptions:
            if isinstance(assumptions[field], dict):
                check_forecast(spec, assumptions[field], where)
            continue
        is_expense = field == "other_expense_monthly"
        amount = sum(r["amount"] for r in rows if r["section"] == section
                     and not (is_expense and is_interest_account(r["account"])))
        if amount:
            caption = "이자 외 영업외비용" if is_expense else section
            warnings.warn(f"{where} 미지정: {caption} 실적 월평균 {amount / len(months):,.2f}원, 추정은 0원",
                          UserWarning, stacklevel=2)


def reconcile_financing(spec, data):
    financing = spec["financing"]
    configured = indexed(financing["facilities"], ("facility",), "financing.facilities")
    if data["debt"]:
        sources = indexed(data["debt"], ("facility",), "debt.csv")
        for key, overlay in configured.items():
            if key not in sources:
                fail("financing.facilities", f"debt.csv에 약정 {key[0]!r}이 없습니다")
            source = sources[key]
            for field, value in overlay.items():
                if field in source and value != source[field]:
                    fail("financing.facilities", f"{key[0]!r}의 {field} 값이 debt.csv와 다릅니다")
            source.update(overlay)
        financing["facilities"] = copy.deepcopy(data["debt"])
    else:
        for row in financing["facilities"]:
            missing = {"lender", "type", "maturity"} - row.keys()
            if missing:
                fail("financing.facilities", f"{row['facility']!r}: debt.csv 없이 필요한 항목 {sorted(missing)}")
            for key, value in (("balance", 0), ("rate", 0), ("note", "")):
                row.setdefault(key, value)
    keys = ("facility", "month")
    from_json = indexed(financing["schedule"], keys, "financing.schedule")
    from_csv = indexed(data["debt_schedule"], keys, "debt_schedule.csv")
    if from_json and from_csv and from_json != from_csv:
        fail("financing.schedule", "debt_schedule.csv와 항목 또는 값이 다릅니다")
    if from_csv:
        financing["schedule"] = copy.deepcopy(data["debt_schedule"])
    facilities = {row["facility"]: row for row in financing["facilities"]}
    future_principal = {}
    for row in financing["schedule"]:
        name = row["facility"]
        if name not in facilities:
            fail("financing.schedule", f"미등록 약정 {name!r}")
        if row["month"] > facilities[name]["maturity"]:
            fail("financing.schedule", f"{name!r} 만기 이후 상환 {row['month']}")
        if row["month"] > spec["last_actual"]:
            future_principal[name] = future_principal.get(name, 0) + row["principal"]
    for name, facility in facilities.items():
        if facility["balance"] and facility["maturity"] < spec["last_actual"]:
            fail("financing.facilities", f"{name!r} 실적 마감 이전 만기 약정에 잔액이 있습니다")
        if future_principal.get(name, 0) > facility["balance"]:
            fail("financing.schedule", f"{name!r} 추정 상환원금이 기초잔액을 초과합니다")


def reconcile_roles(spec, data, has_headcount):
    wages = spec["wages"]
    if has_headcount:
        for role in wages["roles"]:
            policy = {key: role[key] for key in ("raise_from", "annual_raise") if key in role}
            if not policy:
                continue
            fields = [key for key in ("role", "start", "end") if key in role]
            matches = [row for row in data["headcount"] if all(row[key] == role[key] for key in fields)]
            if len(matches) != 1:
                fail("wages.roles", f"{role['role']!r}의 임금 조건을 적용할 headcount.csv 행을 특정하세요")
            source = matches[0]
            for key, value in policy.items():
                if key in source and source[key] != value:
                    fail("wages.roles", f"{role['role']!r}의 {key}가 기존 직군 자료와 다릅니다")
                source[key] = value
        wages["roles"] = copy.deepcopy(data["headcount"])
    for role in wages["roles"]:
        role.setdefault("start", spec["first_month"])
        role.setdefault("end", spec["plan_end"])
        if role["end"] < role["start"]:
            fail("wages.roles", f"{role['role']!r}의 start가 end보다 늦습니다")
        if role["start"] > spec["last_actual"] and role.get("raise_from", role["start"]) < role["start"]:
            fail("wages.roles.raise_from", f"{role['role']!r}의 최초 인상은 입사 월부터 지정하세요")
        if (not wages.get("retro") and "annual_raise" not in role and role["start"] > spec["last_actual"] and "raise_from" in role
                and role["raise_from"][:4] <= role["start"][:4]):
            fail("wages.roles.raise_from",
                 f"{role['role']!r} 추정 입사 월 {role['start']}의 임금 인상은 입사 다음 연도부터 가능합니다")
    indexed(wages["roles"], ("role", "start", "end"), "wages.roles")
    role_names = {role["role"] for role in wages["roles"]}
    for line in spec["cost_lines"]:
        if line["method"] == "per_headcount" and "role" in line["params"]:
            if line["params"]["role"] not in role_names:
                fail(line["id"], f"인원 계획에 없는 직군 {line['params']['role']!r}")


def load_inputs(spec_path, data_path):
    """Return the validated spec and normalized CSV rows, with monetary amounts in won."""
    path, directory = Path(spec_path), Path(data_path)
    schema = read_json(SCHEMA_PATH)
    spec = validate(read_json(path), schema, schema["$defs"], str(path))
    if not spec["first_month"] <= spec["last_actual"] < spec["plan_end"]:
        fail(path, "first_month ≤ last_actual < plan_end 순서여야 합니다")
    if not directory.is_dir():
        fail(directory, "정규화 CSV 폴더가 없습니다")
    if not (directory / "is_monthly.csv").is_file():
        fail(directory, "필수 is_monthly.csv 누락")
    data = {stem: [] for stem in HEADERS}
    for table in sorted(directory.glob("*.csv")):
        if table.stem not in HEADERS:
            fail(table, "정의되지 않은 CSV 파일명")
        data[table.stem] = read_table(table)
        if table.stem in ("bs_monthly", "is_monthly") and not data[table.stem]:
            fail(table, "실적 데이터 행 누락")
    spec["statements"]["bs"] = bool(data["bs_monthly"]) and spec["statements"].get("bs", True)
    check_actuals(spec, data)
    from cash_items import validate_capital
    validate_capital(spec, data)
    normalize_tax(spec)
    from tax_cash import opening_accounts
    warn_zero_tax_payable(spec, data, opening_accounts(spec["tax"]))
    check_lines(spec, data)
    for line in spec["revenue_lines"]:
        if line["driver"] == "manual" and line["params"].get("nonnegative"):
            for month, amount in line["params"]["monthly"].items():
                if amount < 0:
                    fail(f"{line['id']}.monthly.{month}", "월 수취 계획에는 0 이상 금액을 지정하세요")
    from activity_inputs import validate_activity
    validate_activity(spec, data)
    warn_working_capital(spec, data)
    check_depreciation(spec, data)
    check_non_operating(spec, data["is_monthly"])
    warn_missing_basis(spec)
    check_forecast(spec, spec["investment"]["monthly_capex"], "investment.monthly_capex")
    reconcile_roles(spec, data, (directory / "headcount.csv").is_file())
    from investment_assets import validate_asset_classes
    from wage_followup import validate_wages
    validate_asset_classes(spec, data)
    validate_wages(spec)
    reconcile_financing(spec, data)
    return spec, data

