"""Korean current-tax policy and model defaults from references/tax_rules_kr.md."""
from copy import deepcopy
from decimal import Decimal
import re
import warnings


# Sections 2–5 of the supplied rulebook. Amounts are won, periods are in months.
FIRST_SUPPORTED_FY = "2023-01"
SMALL_FILING_START = "2025-01"
NEW_RATES_START = "2026-01"
NEW_LOSS_LIFE_START = "2020-01"
LOSS_LIFE_OLD = 10
LOSS_LIFE_NEW = 15
LOSS_LIMIT_SME = 1.0
LOSS_LIMIT_GENERAL = 0.8
YEAR_MONTHS = 12
INTERIM_PERIOD_MONTHS = 6
INTERIM_DUE_MONTH = 8
INTERIM_SME_EXEMPT_BELOW = 500_000
NATIONAL_DUE_MONTHS = 3
LOCAL_DUE_MONTHS = 4
FAITHFUL_EXTENSION_MONTHS = 1
INSTALLMENT_THRESHOLD = 10_000_000
INSTALLMENT_DELAY_GENERAL = 1
INSTALLMENT_DELAY_SME = 2
LOCAL_RATE_FACTOR = 0.1

_THRESHOLDS = (200_000_000, 20_000_000_000, 300_000_000_000)
_NATIONAL = {
    "pre2026": ((.09, .19, .21, .24), (.19, .19, .21, .24)),
    "from2026": ((.10, .20, .22, .25), (.20, .20, .22, .25)),
}


ANNUAL_AMOUNTS = ("additions", "deductions", "non_taxable", "income_deduction", "credits",
                  "withholding", "ad_hoc", "interim_paid", "installment_amount")


def default_rate_tables():
    """Fresh editable tables. Local marginal rates are one tenth of the applicable class."""
    local_factor = Decimal(str(LOCAL_RATE_FACTOR))
    tables = {}
    for period, (general, small) in _NATIONAL.items():
        tables[period] = {
            "thresholds": list(_THRESHOLDS),
            "national_general": list(general), "national_small_filing": list(small),
            "local_general": [float(Decimal(str(r)) * local_factor) for r in general],
            "local_small_filing": [float(Decimal(str(r)) * local_factor) for r in small],
        }
    return tables


def month_index(month):
    year, number = map(int, month.split("-"))
    return year * YEAR_MONTHS + number - 1


def shift_month(month, offset):
    year, index = divmod(month_index(month) + offset, YEAR_MONTHS)
    return f"{year:04d}-{index + 1:02d}"


def fiscal_period(tax, month):
    """Inclusive legal dates, independent of the data and forecast endpoints."""
    anchor = tax.get("fy_start_month", 1)
    year, number = map(int, month.split("-"))
    start = f"{year - (number < anchor):04d}-{anchor:02d}"
    end = shift_month(start, YEAR_MONTHS - 1)
    for stub_start, stub_end in tax.get("short_fys", {}).items():
        if stub_start <= month <= stub_end:
            return stub_start, stub_end
        if start <= stub_end < month:
            start = shift_month(stub_end, 1)
        if month < stub_start <= end:
            end = shift_month(stub_start, -1)
    founded = tax.get("founded_fy")
    if founded and start < founded <= month:
        start = founded
    return start, end


def loss_life_years(fy_start):
    return LOSS_LIFE_NEW if fy_start >= NEW_LOSS_LIFE_START else LOSS_LIFE_OLD


def _warn(message):
    warnings.warn("tax: " + message, UserWarning, stacklevel=3)


def warn_zero_tax_payable(spec, data, opening_accounts):
    """Compare the last completed legal FY's booked expense with cutoff liabilities."""
    tax = spec["tax"]
    cutoff = spec["last_actual"]
    forecast_start = fiscal_period(tax, shift_month(cutoff, 1))[0]
    prior_start, prior_end = fiscal_period(tax, shift_month(forecast_start, -1))
    expense = sum(r["amount"] for r in data["is_monthly"]
                  if r["section"] == "법인세" and prior_start <= r["month"] <= prior_end)
    if expense <= 0:
        return
    amounts = {}
    if spec["statements"]["bs"]:
        for key in ("national_payable", "local_payable"):
            amounts[key] = sum(r["amount"] for r in data["bs_monthly"]
                               if r["month"] == cutoff and r["section"] == "유동부채"
                               and re.sub(r"[\W_]", "", r["account"].casefold()) in opening_accounts[key])
    amounts.update(tax["opening"])
    # Without a closing balance sheet, a missing balance is unknown. Only
    # compare actual balances or explicit opening amounts; prepaid is an asset.
    if amounts.get("national_payable") == 0 and amounts.get("local_payable") == 0:
        _warn(f"{cutoff} 마감 미지급법인세 잔액은 0원이나 직전 사업연도({prior_start}~{prior_end}) "
              f"실적 법인세비용은 {expense:,}원입니다. 이미 납부했는지 장부에 미반영했는지 확인하세요.")


def credit_mode(tax):
    """Legacy models retain the existing annual-credit calculation and layout."""
    return any(key in tax for key in (
        "minimum_tax", "credit_vintages", "credit_balance_as_of", "new_credits"))


def credit_inputs(tax):
    """Read opening and new lots in user order. Spec amounts are in won."""
    lots = [dict(record, source="opening") for record in tax.get("credit_vintages", [])]
    for year, entry in tax.get("new_credits", {}).items():
        if isinstance(entry, (int, float)):
            records = [{"label": "당기 발생 세액공제", "amount": entry,
                        "minimum_tax_applies": True, "rural_tax": False}]
        else:
            records = entry if isinstance(entry, list) else [entry]
        for record in records:
            if "year" in record and record["year"] != int(year):
                raise ValueError(f"tax.new_credits.{year}: 개시 연도와 입력한 발생연도가 다릅니다")
            lots.append(dict(record, year=int(year), source="new"))
    return lots


def _credit_fy(tax, year):
    """An integer vintage identifies the FY by its start year, including non-December closes."""
    starts = set()
    for number in range(1, YEAR_MONTHS + 1):
        month = f"{year:04d}-{number:02d}"
        if month < tax.get("founded_fy", ""):
            continue
        start, end = fiscal_period(tax, month)
        if int(start[:4]) == year:
            starts.add((start, end))
    if len(starts) != 1:
        raise ValueError(f"tax: 발생연도 {year}에 개시하는 사업연도를 하나로 특정할 수 없습니다")
    return starts.pop()


def _normalize_credit_inputs(spec, first_forecast_fy):
    tax = spec["tax"]
    tax.setdefault("minimum_tax", {"apply": True, "rate": "auto"})
    if tax.get("credit_vintages") and "credit_balance_as_of" not in tax:
        raise ValueError("tax.credit_balance_as_of: 이월세액공제 기초 잔액의 사업연도 말 기준 월이 필요합니다")
    as_of = tax.setdefault("credit_balance_as_of", shift_month(first_forecast_fy, -1))
    before_founding = not tax.get("credit_vintages") and tax.get("founded_fy") == shift_month(as_of, 1)
    if (not before_founding and fiscal_period(tax, as_of)[1] != as_of) or as_of > spec["last_actual"]:
        raise ValueError("tax.credit_balance_as_of: 실적 마감일까지 종료한 사업연도의 마지막 월을 입력하세요")
    first_use = shift_month(as_of, 1)
    lots = credit_inputs(tax)
    if tax.get("credit_vintages") and first_use < spec["first_month"]:
        raise ValueError("tax.credit_balance_as_of: 기준 월 직후부터의 전체 사업연도 실적이 필요합니다")
    if tax.get("credit_vintages") and first_use < FIRST_SUPPORTED_FY:
        raise ValueError("tax.credit_balance_as_of: 잔액 사용기간이 제공된 세율표보다 이릅니다")
    for lot in lots:
        start, end = _credit_fy(tax, lot["year"])
        if lot["source"] == "opening":
            if end > as_of:
                raise ValueError(f"tax.credit_vintages: 기준 월까지 종료하지 않은 발생 사업연도 {start}")
            if as_of > shift_month(end, 10 * YEAR_MONTHS) and lot["amount"]:
                raise ValueError(f"tax.credit_vintages: 기준 월에 이미 만료한 이월세액공제 {start}")
        elif not max(first_forecast_fy, spec["first_month"]) <= start <= spec["plan_end"] or start <= as_of:
            raise ValueError(f"tax.new_credits: 추정 대상 사업연도의 개시 연도를 지정하세요 ({start})")


def normalize_tax(spec):
    """Validate tax-specific relationships after scalar/schema validation."""
    tax = spec["tax"]
    if "rate" in tax:
        del tax["rate"]
        _warn("단일 세율은 쓰지 않습니다. 법인세법 세율표로 계산합니다.")
    if "sme" in tax:
        if "is_sme" in tax and tax["is_sme"] != tax["sme"]:
            raise ValueError("tax.sme: tax.is_sme와 중소기업 여부가 다릅니다")
        tax["is_sme"] = tax.pop("sme")
    if "is_sme" not in tax:
        _warn("중소기업 여부 미확인. 중소기업 아님으로 계산합니다.")
        tax["is_sme"] = False
    stubs = sorted(tax["short_fys"].items())
    previous_end = ""
    for start, end in stubs:
        length = month_index(end) - month_index(start) + 1
        if not 1 <= length < YEAR_MONTHS or start <= previous_end:
            raise ValueError(f"tax.short_fys: 실제 단기사업연도의 기간 또는 중복 오류 {start} ~ {end}")
        previous_end = end
    founded = tax.get("founded_fy")
    if founded and (founded > spec["first_month"] or any(start < founded for start, _ in stubs)):
        raise ValueError("tax.founded_fy: 설립 전 기간에 실적 또는 단기사업연도가 있습니다")
    cutoff = spec["last_actual"]
    first_forecast_fy = fiscal_period(tax, shift_month(cutoff, 1))[0]
    if first_forecast_fy < FIRST_SUPPORTED_FY:
        raise ValueError("tax: 제공된 세율표의 적용 개시일보다 이른 추정 사업연도입니다")
    if credit_mode(tax):
        _normalize_credit_inputs(spec, first_forecast_fy)
    prior = tax["prior"]
    preceding_start, preceding_end = fiscal_period(tax, shift_month(first_forecast_fy, -1))
    preceding_months = month_index(preceding_end) - month_index(preceding_start) + 1
    if preceding_months != YEAR_MONTHS and prior.get("months", preceding_months) != preceding_months:
        raise ValueError("tax.prior.months: 지정한 직전 단기사업연도 월수와 다릅니다")
    prior.setdefault("months", preceding_months)
    prior.setdefault("is_sme", tax["is_sme"])
    if spec["first_month"] > first_forecast_fy:
        _warn(f"{first_forecast_fy} 개시 사업연도의 모델 시작 이전 실적이 없어 누적 세전이익 확인 필요")
    if "assessed_tax" not in prior and any(prior.get(k) for k in ("credits", "withholding", "ad_hoc")):
        _warn("직전 산출세액 미확인: 직전 공제·기납부 입력은 보관하고 첫 중간예납은 상반기 방식으로 계산합니다.")
    if founded == first_forecast_fy and "assessed_tax" in prior:
        raise ValueError("tax.prior: 신설 법인의 첫 사업연도에 직전 산출세액을 지정할 수 없습니다")
    if (not spec["statements"].get("bs") and first_forecast_fy <= cutoff
            and ("national_payable" not in tax["opening"] or "national_prepaid" not in tax["opening"])):
        _warn("BS 없는 사업연도 중간 마감: 미지정 국세 기초 잔액은 실적 누적 법인세비용과 "
              "확인된 기납부액으로 추정합니다. 실제 미지급·선납 잔액은 tax.opening에 입력하세요.")
    tables = default_rate_tables()
    for version, supplied in tax["rate_tables"].items():
        tables[version].update(deepcopy(supplied))
        thresholds = tables[version]["thresholds"]
        if thresholds[0] <= 0 or thresholds != sorted(set(thresholds)):
            raise ValueError(f"tax.rate_tables.{version}.thresholds: 양수 오름차순의 구간 경계가 필요합니다")
    tax["rate_tables"] = tables
    amounts = tax["loss_vintages"]
    legacy = tax.pop("loss_carryforward", 0)
    if legacy:
        if amounts:
            raise ValueError("tax.loss_carryforward: 총액과 loss_vintages를 중복 지정할 수 없습니다")
        amounts.append({"fy_start": "unknown", "amount": legacy})
    seen = set()
    for vintage in amounts:
        start = vintage["fy_start"]
        if start in seen:
            raise ValueError(f"tax.loss_vintages: 중복 발생 사업연도 {start}")
        seen.add(start)
        if start == "unknown":
            if vintage["amount"]:
                _warn("결손금 발생연도 미상. 만료일 계산을 생략합니다.")
            continue
        if start >= first_forecast_fy or (founded and start < founded):
            raise ValueError(f"tax.loss_vintages: 이미 종료한 사업연도의 기초 잔액만 지정하세요 ({start})")
        if fiscal_period(tax, start)[1] > cutoff:
            raise ValueError(f"tax.loss_vintages: 실적 마감까지 종료하지 않은 발생 사업연도 {start}")
        cutoff_start = fiscal_period(tax, cutoff)[0]
        if month_index(cutoff_start) - month_index(start) > YEAR_MONTHS * loss_life_years(start):
            raise ValueError(f"tax.loss_vintages: 실적 마감 사업연도 전에 이미 만료한 결손금 {start}")
    amounts.sort(key=lambda item: item["fy_start"])
    missing_adjustments = []
    active_start = first_forecast_fy
    while active_start <= spec["plan_end"]:
        active = tax["annual"].get(active_start, {})
        if "additions" not in active and "deductions" not in active:
            missing_adjustments.append(active_start)
        active_end = fiscal_period(tax, active_start)[1]
        if month_index(active_end) >= month_index(spec["plan_end"]):
            break
        active_start = shift_month(active_end, 1)
    if missing_adjustments:
        _warn("세무조정 자료 확인 필요. 미지정 가감액은 0원 (" + ", ".join(missing_adjustments) + ")")
    if tax["installments"] and not any(r.get("installment_amount", 0) for r in tax["annual"].values()):
        _warn("분납액 미지정: 분납할 금액이 있는 사업연도에 installment_amount를 입력하세요")
    for start, record in tax["annual"].items():
        for key in ANNUAL_AMOUNTS:
            record.setdefault(key, 0)
        if fiscal_period(tax, start)[0] != start or (founded and start < founded):
            raise ValueError(f"tax.annual: 실제 사업연도 개시 월을 지정하세요 ({start})")
        if record["credits"] > 0:
            if credit_mode(tax):
                _warn(f"{start} annual.credits는 당기 일반감면으로 계산합니다. 유형별 세액공제 입력과 중복 여부를 확인하세요.")
            else:
                _warn(f"{start} 세액공제·감면 입력: 최저한세 검토 필요")
        if record["installment_amount"] and not tax["installments"]:
            raise ValueError(f"tax.annual.{start}.installment_amount: 분납 선택이 필요합니다")
        if record["interim_paid"] and start > cutoff:
            raise ValueError(f"tax.annual.{start}.interim_paid: 실적 마감 이후의 기납부액입니다")
    return tax
