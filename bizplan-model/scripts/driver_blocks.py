"""Build operating drivers on the model's shared monthly axis. Input money is KRW."""

from collections import defaultdict
from datetime import date
from math import isfinite
import re

from openpyxl.utils.datetime import to_excel

from spec_io import is_interest_account


SECTIONS = ("매출", "매출원가", "판관비", "영업외수익", "영업외비용", "법인세")


def _number(value, field, minimum=None, maximum=None):
    try:
        result = float(value)
    except (TypeError, ValueError):
        raise ValueError(f"{field}: 숫자가 필요합니다 ({value!r}).") from None
    if not isfinite(result) or (minimum is not None and result < minimum) or (
        maximum is not None and result > maximum
    ):
        raise ValueError(f"{field}: 허용 범위를 벗어난 값입니다 ({value!r}).")
    return result


def _month(value, field):
    if not isinstance(value, str) or not re.fullmatch(r"(?!0000)\d{4}-(0[1-9]|1[0-2])", value):
        raise ValueError(f"{field}: YYYY-MM 월이 필요합니다 ({value!r}).")
    return value


def _sheet(name):
    return "'" + name.replace("'", "''") + "'"


def _sum(rows, col, sheet=None):
    """Compress adjacent same-sheet cells into ranges for large account mappings."""
    rows = sorted(rows)
    if not rows:
        return "0"
    prefix = _sheet(sheet) + "!" if sheet else ""
    spans = []
    start = end = rows[0]
    for row in rows[1:] + [None]:
        if row == end + 1:
            end = row
            continue
        first = f"{prefix}{col}{start}"
        spans.append(first if start == end else f"SUM({first}:{col}{end})")
        start = end = row
    return "(" + "+".join(spans) + ")"


def _formulas(m, sheet, row, label, expression, annual="sum"):
    m.row(sheet, row, label, ["=" + expression(m.col(i)) for i in range(len(m.months))],
          annual=annual)


def _monthly_input(m, sheet, row, label, monthly, group, description):
    values = [_number(monthly.get(month, 0), f"{label} {month}") for month in m.months]
    m.input(sheet, row, label, values, group, description)


def _annual_factor(col, rate):
    # Base inputs are at the cutoff year's level. Each following January applies one raise.
    years = f"MAX(0,YEAR({col}$5)-YEAR(LastActual))"
    return f"IF({years}=0,1,(1+{rate})^{years})"


def fixed_cost_basis(line, row, col, sheet=""):
    """Read the unraised monthly cost in display units, also used for retro pay."""
    if "activity_increment" in line["params"]:
        return f"{sheet}{col}{row+6}"
    standard = f"{sheet}$G${row+4}/Unit"
    if line["method"] != "fixed_inflation" or "monthly" not in line["params"]:
        return standard
    override = f"{sheet}{col}{row+6}"
    return f"IF(ISBLANK({override}),{standard},IF(AND(ISNUMBER({override}),{override}>=0),{override},NA()))"


def _cash_assumptions(m):
    finance = m.spec.get("financing", {})
    for field in ("opening_cash", "min_cash", "revolver_limit", "new_rate", "revolver_rate"):
        _number(finance.get(field, 0), f"financing.{field}", 0, 1 if field.endswith("_rate") else None)
    facilities = m.data.get("debt") or finance.get("facilities", [])
    schedule = m.data.get("debt_schedule") or finance.get("schedule", [])
    balances = {}
    maturities = {}
    drawn_revolver = 0
    for item in facilities:
        name = item["facility"]
        if not name or name in balances:
            raise ValueError(f"차입 약정 식별자가 비어 있거나 중복입니다 ({name!r}).")
        balances[name] = _number(item.get("balance", 0), f"{name} 잔액", 0)
        if item.get("type") == "revolver":
            drawn_revolver += balances[name]
        _number(item.get("rate", 0), f"{name} 연이율", 0, 1)
        maturity = item["maturity"]
        if isinstance(maturity, str) and re.fullmatch(r"\d{4}-\d{2}-\d{2}", maturity):
            try:
                date.fromisoformat(maturity)
            except ValueError:
                raise ValueError(f"{name} 만기: 유효한 날짜가 필요합니다 ({maturity}).") from None
            maturity = maturity[:7]
        maturities[name] = _month(maturity, f"{name} 만기")
    if drawn_revolver > finance.get("revolver_limit", 0):
        raise ValueError("실적 마감 한도대출 잔액이 약정 한도(revolver_limit)를 초과합니다.")
    repayments = defaultdict(float)
    for item in schedule:
        name = item["facility"]
        if name not in balances:
            raise ValueError(f"상환 스케줄에 등록되지 않은 약정이 있습니다 ({name}).")
        month = _month(item["month"], f"{name} 상환 월")
        principal = _number(item.get("principal", 0), f"{name} {month} 상환 원금", 0)
        if month > m.spec["last_actual"]:
            if principal and month > maturities[name]:
                raise ValueError(f"{name}: {month} 원금 상환이 약정 만기 이후입니다.")
            repayments[name] += principal
            if repayments[name] > balances[name] + 1e-6:
                raise ValueError(f"{name}: 계획 상환 원금 합계가 실적 마감 잔액을 초과합니다.")


def _actual_is(m, lines):
    """Key by both section and account. Net duplicate source records in each month."""
    owners = {}
    for line in lines:
        section = line.get("section", "매출")
        for account in line["accounts"]:
            key = (section, account)
            if key in owners:
                raise ValueError(f"{section} / {account}: {owners[key]} 및 {line['id']}에 중복 배정했습니다.")
            owners[key] = line["id"]

    by_account = defaultdict(lambda: defaultdict(float))
    observed = set()
    months = set(m.months)
    for record in m.data.get("is_monthly", []):
        section, account = record["section"], record["account"]
        month = _month(record["month"], "is_monthly.month")
        if section not in SECTIONS or not account:
            raise ValueError(f"is_monthly: 손익 구분 또는 계정 오류 ({section} / {account}).")
        if month not in months:
            continue
        if section in SECTIONS[:3] and (section, account) not in owners:
            raise ValueError(f"is_monthly: {section} / {account}의 모델 라인을 지정하세요.")
        by_account[section, account][month] += _number(record["amount"], f"{month} {account}")
        observed.add(month)
    if not any(month <= m.spec["last_actual"] for month in observed):
        raise ValueError("is_monthly: 모델의 실적 기간에 손익 자료가 없습니다.")
    missing = owners.keys() - by_account.keys()
    if missing:
        raise ValueError(f"is_monthly: 배정한 계정을 원천에서 찾을 수 없습니다 ({sorted(missing)}).")

    sheet = "실적_IS"
    mapped = defaultdict(list)
    totals = defaultdict(list)
    accounts = sorted(by_account, key=lambda k: (SECTIONS.index(k[0]), owners.get(k, ""), k[1]))
    for index, (section, account) in enumerate(accounts):
        row = 16 + index
        raw = row + len(accounts) + 1
        amounts = by_account[section, account]
        m.row(sheet, raw, f"{section} · {account} (원)", [amounts.get(mon, 0) for mon in m.months],
              kind="actual")
        _formulas(m, sheet, row, f"{section} · {account}", lambda c: f"{c}{raw}/Unit")
        totals[section].append(row)
        if (section, account) in owners:
            mapped[owners[section, account]].append(row)
    for row, section in enumerate(SECTIONS, 8):
        _formulas(m, sheet, row, section, lambda c: _sum(totals[section], c))
    m.row(sheet, 14, "월 손익 자료 있음", [int(mon in observed) for mon in m.months],
          kind="actual", annual="none")
    designated = set(m.spec["investment"]["depreciation_accounts"])
    dep, interest = [], []
    for i, (section, account) in enumerate(accounts, 16):
        if account in designated and section in SECTIONS[1:3]:
            dep.append(i)
        if section == "영업외비용" and is_interest_account(account):
            interest.append(i)
    m.actual_dep_row = 17 + 2 * len(accounts)
    m.actual_interest_row = m.actual_dep_row + 1
    _formulas(m, sheet, m.actual_dep_row, "지정 계정 감가상각비", lambda c: _sum(dep, c))
    _formulas(m, sheet, m.actual_interest_row, "이자비용", lambda c: _sum(interest, c))
    return mapped


def _extensions(m, ids):
    rows = defaultdict(list)
    for row, item in enumerate(m.spec.get("extensions", []), 8):
        key = item["line_id"]
        if key not in ids:
            raise ValueError(f"확장조정: 알 수 없는 라인입니다 ({key}).")
        _monthly_input(m, "확장조정", row, f"{key} · {item['label']} (원)", item.get("monthly", {}),
                       "확장조정", "해당 월의 매출 또는 비용에 더할 원 금액")
        rows[key].append(row)
    return rows


def _wages(m):
    sheet = "인건비"
    wages = m.spec.get("wages", {})
    roles = defaultdict(list)
    records = wages.get("roles", [])
    from wage_followup import check_fixed_cohorts, prepare_controls, salary_factor
    prepare_controls(m, records)
    retro = wages.get("retro")
    if records or any(line.get("follows_wage_raise") for line in m.spec["cost_lines"]):
        m.param(sheet, 10, "연 임금 인상률", wages.get("annual_raise", 0), sheet, "직군별 임금 및 연동 비용 고정분의 연 인상률")
        _formulas(m, sheet, 11, "기존 인원 기본 임금 배수",
                  lambda c: salary_factor(c, "$G$10", "WG_Effective") if retro else _annual_factor(c, "$G$10"), "last")
        m.format_row(sheet, 11, '0.00"x"')
    for i, record in enumerate(records):
        role = record["role"]
        row = 14 + 6 * i
        start = _month(record.get("start") or m.months[m.last + 1], f"{role} 시작 월")
        end = _month(record.get("end") or m.months[-1], f"{role} 종료 월")
        if end < start:
            raise ValueError(f"{role}: 종료 월이 시작 월보다 빠릅니다.")
        count = _number(record.get("count", 0), f"{role} 인원", 0)
        cost = _number(record.get("monthly_cost", 0), f"{role} 1인 월 비용", 0)
        first_year = max(int(m.months[m.last][:4]), int(start[:4])) + 1
        raise_from = _month(record.get("raise_from") or f"{first_year:04}-01", f"{role} 인상 시작 월")
        m.input(sheet, row, f"{role} 인원", [count if start <= mon <= end else 0 for mon in m.months],
                sheet, f"{start} ~ {end} 직군별 적용 인원", annual="last")
        m.param(sheet, row + 1, f"{role} 1인 월 비용 (원)", cost, sheet, "임금 인상 전 1인당 월 비용")
        first_raise = f"$G${row+3}"
        if retro and "raise_from" not in record:
            dependent = "WG_Effective"
            if start > m.months[m.last]:
                hire_date = to_excel(date.fromisoformat(start+"-01"))
                dependent = f"IF({hire_date}>=WG_Effective,DATE(YEAR({hire_date})+1,1,1),WG_Effective)"
            m.cell(sheet, row + 3, 2, f"{role} 인상 시작 월", "text")
            m.cell(sheet, row + 3, 7, "="+dependent)
            m.format_row(sheet, row + 3, "yyyy-mm")
        else:
            m.param(sheet, row + 3, f"{role} 인상 시작 월", to_excel(date.fromisoformat(raise_from + "-01")),
                    sheet, "첫 인상 월부터 매년 같은 달에 연 임금 인상률 적용")
            if retro or i in m.wage_rates or role in m.fixed_wage_roles:
                m.wage_date_rows.add(row + 3)
            if retro:
                first_raise = f"MAX({first_raise},WG_Effective)"
            if start > m.months[m.last]:
                hire_col = m.col(m.months.index(min(start, m.months[-1])))
                hire = f"${hire_col}$5"
                if retro or i in m.wage_rates:
                    first_raise = f"MAX({first_raise},{hire})"
                else:
                    # Preserve the entry-year treatment of existing headcount plans.
                    first_raise = f"MAX({first_raise},DATE(YEAR({hire})+1,1,1))"
        rate = m.wage_rates.get(i, "$G$10")
        m.wage_policy_by_row[row] = (first_raise, rate)
        _formulas(m, sheet, row + 4, f"{role} 누적 임금 배수",
                  lambda c: salary_factor(c, rate, first_raise), "last")
        _formulas(m, sheet, row + 5, f"{role} 인원 × 임금 배수",
                  lambda c: f"{c}{row}*{c}{row+4}", "last")
        _formulas(m, sheet, row + 2, f"{role} 인건비", lambda c: f"{c}{row+5}*$G${row+1}/Unit")
        m.format_row(sheet, row, "0.##")
        m.format_row(sheet, row + 4, '0.00"x"')
        m.format_row(sheet, row + 5, "0.##")
        roles[role].append(row)
    check_fixed_cohorts(m, roles)
    all_rows = [row for rows in roles.values() for row in rows]
    _formulas(m, sheet, 8, "총 인원", lambda c: _sum(all_rows, c), "last")
    m.format_row(sheet, 8, "0.##")
    _formulas(m, sheet, 9, "총 인건비", lambda c: _sum([r + 2 for r in all_rows], c))
    return roles


def _line_total(m, line, sheet, total, actual, base, adjust, ext_rows, month_gate=None, wage_timing=None):
    label = line["label"]
    if ext_rows:
        _formulas(m, sheet, adjust, f"{label} 수동 조정", lambda c: _sum(ext_rows, c, "확장조정"))
    else:
        _monthly_input(m, sheet, adjust, f"{label} 수동 조정 (원)", {}, "수동 조정", "추정 월에 가산할 원 금액")
    key = line["id"]
    factor = m.factor(key)
    addition = m.addition(key)
    if line.get("driver") in ("price_volume", "contracts", "activity"):
        factor = "1"  # Quantity or contract components already include the case multiplier.
    if line.get("driver") == "contracts":
        addition = "0"  # The case addition belongs to unidentified new contracts.
    def result(c):
        forecast = f"{c}{base}*{factor}+{addition}/Unit+{c}{adjust}"
        if month_gate is not None:
            forecast = f"IF({c}{month_gate},{forecast},0)"
        if wage_timing is not None:
            # Arrears are paid at settlement even if that month is outside a recurring expense gate.
            forecast += f"+'인건비'!{c}{wage_timing}"
        if m.activity_lines and line.get("method"):
            forecast += f"+'활동지표'!{c}{m.fixed_rows[key]}"
        return f"IF('INDEX'!{c}$6,{c}{actual},{forecast})"
    _formulas(m, sheet, total, label, result)


def _contracts(m, line, sheet, items):
    key = line["id"]
    params = line["params"]
    exclude = m.driver_switch(key, "exclude_unconfirmed")
    new_factor = m.driver_switch(key, "new_multiplier")
    m.param(sheet, 14, "미식별 신규 월 매출 (원)", params.get("new_monthly", 0),
            line["label"], line["basis"])
    # Contiguous expected-amount rows keep the sum compact even for large CSVs.
    expected_rows = range(16 + 2 * len(items), 16 + 3 * len(items))
    renewal = params.get("renewal")
    renewal_rows = []
    if renewal is not None:
        base_row = expected_rows.stop
        m.param(sheet, base_row, "기본 계약 갱신 확률", renewal["probability"], line["label"], line["basis"])
        m.param(sheet, base_row + 1, "계약 갱신 기간 (개월)", renewal["term_months"], line["label"], line["basis"])
        m.renewal_probability_refs[key] = f"{_sheet(sheet)}!$G${base_row}"
        renewal_probability = m.driver_switch(key, "renewal_probability")
        term = f"$G${base_row+1}"
    for i, item in enumerate(items):
        label = item["item"]
        row = 16 + i
        prob_row = row + len(items)
        start = _month(item["start"], f"{label} 계약 시작")
        end = _month(item["end"], f"{label} 계약 종료")
        if start > end:
            raise ValueError(f"{label}: 계약 종료 월이 시작 월보다 빠릅니다.")
        amount = _number(item["amount"], f"{label} 계약 월 금액")
        prob = _number(item.get("probability", 0), f"{label} 실현 확률", 0, 1)
        m.input(sheet, row, f"{label} 월 계약액 (원)", [amount if start <= mon <= end else 0 for mon in m.months],
                line["label"], f"{start} ~ {end} · {item.get('note', '')}")
        m.param(sheet, prob_row, f"{label} 실현 확률", prob, line["label"], "확정 계약은 1")
        item_switch = m.active_driver_switches.get((key, "exclude_item:" + label))
        exclusion = f"OR({exclude},{item_switch})" if item_switch else exclude
        _formulas(m, sheet, expected_rows[i], f"{label} 기대 매출", lambda c:
                  f"IF(AND({exclusion},$G${prob_row}<1),0,{c}{row}*$G${prob_row})")
        if renewal is not None and m.months[m.last] <= end < m.months[-1]:
            expiry = m.months.index(end)
            expiry_col = m.col(expiry)
            renewal_row = base_row + 2 + len(renewal_rows)
            renewal_rows.append(renewal_row)
            values = [None] * (expiry + 1)
            for j in range(expiry + 1, len(m.months)):
                c = m.col(j)
                cycle = f"(1+INT((COLUMN({c}$5)-COLUMN(${expiry_col}$5)-1)/{term}))"
                values.append(f"=IF(AND({exclusion},$G${prob_row}<1),0,"
                              f"${expiry_col}${row}*$G${prob_row}*{renewal_probability}^{cycle})")
            m.row(sheet, renewal_row, f"{label} 갱신 기대 매출", values)
    factor, addition = m.factor(line["id"]), m.addition(line["id"])
    contracts = [*expected_rows, *renewal_rows]
    _formulas(m, sheet, 12, "확인 계약 기대 매출", lambda c: f"{_sum(contracts,c)}*{factor}")
    new_intake = f"($G$14*{new_factor}*{factor}+{addition})/Unit"
    if params.get("new_monthly_recurring"):
        m.param(sheet, 15, "신규 계약 월 이탈률", params.get("churn", 0), line["label"], line["basis"])
        # Contracts already on the books are in the item rows. Start this
        # pool from zero at the cutoff and retain only forecast new intake.
        pool = [None] * (m.last + 1)
        for i in range(m.last + 1, len(m.months)):
            prior = f"{m.col(i-1)}13"
            pool.append(f"={prior}*(1-$G$15)+{new_intake}")
        m.row(sheet, 13, "미식별 신규 계약 반복 매출", pool)
    else:
        _formulas(m, sheet, 13, "미식별 신규 매출", lambda c: new_intake)
    _formulas(m, sheet, 10, "계약 및 신규 매출", lambda c: f"{c}12+{c}13")


def _price_volume(m, line, sheet, source_months):
    params = line["params"]
    label, basis = line["label"], line["basis"]
    for row, caption, key in ((14, "기준 월 수량", "quantity"), (15, "단가 (원)", "price"),
                              (16, "연 수량 성장률", "annual_growth")):
        m.param(sheet, row, caption, params.get(key, 0), label, basis)
    m.param(sheet, 19, "수량 계절성 적용 (1/0)", int(params.get("seasonal", False)), label, basis)
    # Overlapping CSV/spec fields are reconciled by spec_io. Missing fields stay
    # blank so removing an override in Excel immediately restores its fallback.
    explicit = params.get("monthly", {})
    for row, field, caption in ((20, "quantity", "월별 수량"), (21, "price", "월별 단가 (원)")):
        values = [explicit.get(mon, {}).get(field, source_months.get(mon, {}).get(field))
                  for mon in m.months]
        m.input(sheet, row, caption, values, label, "실적 관측값 또는 월별 계획. 빈 추정 월에는 기준 가정 적용",
                annual="sum" if field == "quantity" else "none")
    m.format_row(sheet, 20, "#,##0.##")
    _formulas(m, sheet, 22, "달력 월", lambda c: f"MONTH({c}$5)", "none")
    _formulas(m, sheet, 23, "실적 수량 관측 월", lambda c: f"'INDEX'!{c}$6*ISNUMBER({c}20)", "none")
    # Use only observed historical quantities. An observed zero is valid.
    end = m.col(m.last)
    qty, month, observed = [f"$J${r}:${end}${r}" for r in (20, 22, 23)]
    _formulas(m, sheet, 24, "수량 계절 지수", lambda c:
              f"IF($G$19,IFERROR(AVERAGEIFS({qty},{month},{c}22,{observed},1)"
              f"/AVERAGEIF({observed},1,{qty}),1),1)", "none")
    factor = m.factor(line["id"])
    price_factor = m.driver_switch(line["id"], "price_multiplier")
    _formulas(m, sheet, 17, "적용 수량", lambda c:
              f'IF(\'INDEX\'!{c}$6,IF(ISNUMBER({c}20),{c}20,""),'
              f"IF(ISNUMBER({c}20),{c}20,$G$14*{c}24*{_annual_factor(c,'$G$16')})*{factor})")
    _formulas(m, sheet, 18, "적용 단가 (원)", lambda c:
              f'IF(\'INDEX\'!{c}$6,IF(ISNUMBER({c}21),{c}21,""),'
              f"IF(ISNUMBER({c}21),{c}21,$G$15)*{price_factor})", "none")
    _formulas(m, sheet, 10, "수량 × 단가", lambda c:
              f'IF(COUNT({c}17:{c}18)=2,{c}17*{c}18/Unit,"")')
    m.format_row(sheet, 17, "#,##0.##")
    m.format_row(sheet, 24, '0.00"x"')


def _revenue(m, line, actual_rows, items, ext_rows, volume, rev_rows):
    sheet = m.details[line["id"]]
    p = line.get("params", {})
    driver = line["driver"]
    def param(row, label, key):
        m.param(sheet, row, label, p.get(key, 0), line["label"], line["basis"])

    _formulas(m, sheet, 9, "배정 계정 실적", lambda c: _sum(actual_rows, c, "실적_IS"))
    if driver == "pct_of_line":
        target_rows = [rev_rows[k] for k in p["line_ids"]]
        param(14, "연동 매출 대비 비율", "rate")
        _formulas(m, sheet, 10, "대상 매출 × 연동 비율", lambda c: f"{_sum(target_rows,c,'운영수입')}*$G$14")
    elif driver == "growth":
        param(14, "기준 월 매출 (원)", "base_monthly")
        param(15, "연 성장률", "annual_growth")
        m.param(sheet, 16, "계절성 적용 (1/0)", int(p.get("seasonal", False)), line["label"], line["basis"])
        last_col = m.col(len(m.months) - 1)
        _formulas(m, sheet, 17, "달력 월", lambda c: f"MONTH({c}$5)", "none")
        _formulas(m, sheet, 18, "계절성 관측 월", lambda c: f"'INDEX'!{c}$6*'실적_IS'!{c}14", "none")
        values, calendar, observed = [f"$J${r}:${last_col}${r}" for r in (9, 17, 18)]
        _formulas(m, sheet, 19, "계절 배수", lambda c:
                  f"IF($G$16,IFERROR(AVERAGEIFS({values},{calendar},{c}17,{observed},1)"
                  f"/AVERAGEIF({observed},1,{values}),1),1)", "none")
        _formulas(m, sheet, 10, "성장률 적용 매출", lambda c: f"$G$14/Unit*{c}19*{_annual_factor(c,'$G$15')}")
    elif driver == "activity":
        from activity_model import revenue_block
        revenue_block(m, line, sheet, volume)
    elif driver == "price_volume":
        _price_volume(m, line, sheet, volume)
    elif driver == "contracts":
        _contracts(m, line, sheet, items)
    elif driver == "cohort":
        param(14, "기초 고객 수", "opening_customers")
        param(15, "월 신규 고객 수", "new_customers")
        param(16, "월 이탈률", "churn")
        param(17, "월 객단가 (원)", "arpu")
        new_factor = m.driver_switch(line["id"], "new_multiplier")
        exclude = m.driver_switch(line["id"], "exclude_unconfirmed")
        m.param(sheet, 19, "신규 유입 실현 확률", p.get("new_probability", 1), line["label"], line["basis"])
        _formulas(m, sheet, 20, "추정 신규 유입 고객 수", lambda c:
                  f"IF(OR('INDEX'!{c}$6,AND({exclude},$G$19<1)),0,$G$15*{new_factor}*$G$19)")
        m.format_row(sheet, 20, "#,##0.##")
        customers = []
        for i in range(len(m.months)):
            c = m.col(i)
            prior = f"{m.col(i-1)}18" if i else "$G$14"
            customers.append(f"=IF('INDEX'!{c}$6,$G$14,{prior}*(1-$G$16)+{c}20)")
        m.row(sheet, 18, "기말 고객 수 (추정)", customers, annual="last")
        m.format_row(sheet, 18, "#,##0.##")
        _formulas(m, sheet, 10, "고객 수 × 객단가", lambda c: f"{c}18*$G$17/Unit")
    elif driver == "manual":
        _monthly_input(m, sheet, 14, "월 매출 계획 (원)", p.get("monthly", {}), line["label"], line["basis"])
        _formulas(m, sheet, 10, "월 매출 계획", lambda c: f"{c}14")
    else:
        raise ValueError(f"{line['id']}: 지원하지 않는 매출 동인입니다 ({driver}).")
    _line_total(m, line, sheet, 8, 9, 10, 11, ext_rows)


def _cost(m, line, row, actual_rows, rev_rows, cost_summaries, roles, ext_rows):
    sheet = "운영비용"
    p = line.get("params", {})
    method = line["method"]
    from wage_followup import fixed_factor, retro_adjustment
    def fixed_multiplier(c):
        return fixed_factor(m, line, roles, c) or _annual_factor(c, f"$G${row+5}")

    def param(offset, label, key):
        m.param(sheet, row + offset, f"{line['label']} {label}", p.get(key, 0), line["label"], line["basis"])

    if method in ("mixed", "fixed_inflation"):
        if line.get("follows_wage_raise"):
            m.cell(sheet, row + 5, 2, "연 임금 인상률 연동", "text")
            wage_rate = "$G$10"
            if p.get("role"):
                record_no = (roles[p["role"]][0]-14)//6
                wage_rate = m.wage_rates.get(record_no, wage_rate)
            m.cell(sheet, row + 5, 7, f"='인건비'!{wage_rate}")
            m.format_row(sheet, row + 5, "0.0%")
        else:
            param(5, "연 인상률", "annual_growth")
    _formulas(m, sheet, row + 1, "배정 계정 실적", lambda c: _sum(actual_rows, c, "실적_IS"))
    if method in ("pct_of_revenue", "mixed"):
        targets = p.get("revenue_ids", [])
        if not targets or set(targets) - rev_rows.keys() or len(set(targets)) != len(targets):
            raise ValueError(f"{line['id']}: 연동 매출 라인을 확인하세요 ({targets}).")
        source = [rev_rows[key] for key in targets]
        rate_offset = 6 if method == "mixed" else 4
        param(rate_offset, "매출 대비 비율", "rate")
        rate = f"$G${row+rate_offset}"
        if method == "mixed":
            param(4, "기준 월 고정비 (원)", "base_monthly")
            expr = lambda c: f"$G${row+4}/Unit*{fixed_multiplier(c)}+{_sum(source,c,'운영수입')}*{rate}"
        else:
            expr = lambda c: f"{_sum(source,c,'운영수입')}*{rate}"
    elif method in ("consumption", "per_metric", "per_activity"):
        from activity_model import metric_cost
        expr = metric_cost(m, line, row)
    elif method == "pct_of_line":
        targets = p["line_ids"]
        revenue = [rev_rows[k] for k in targets if k in rev_rows]
        expense = [cost_summaries[k] for k in targets if k not in rev_rows]
        param(4, "연동 라인 대비 비율", "rate")
        expr = lambda c: f"({_sum(revenue,c,'운영수입')}+{_sum(expense,c)})*$G${row+4}"
    elif method == "fixed_inflation":
        link = p.get("activity_increment")
        if link:
            from activity_model import NET_ACTIVITY, lookup
            source = m.details[link["line"]]
            m.cell(sheet, row+4, 2, "추가 활동당 비용 (원)", "text")
            m.cell(sheet, row+4, 7, link["unit_cost"], "actual")
            _formulas(m, sheet, row+6, f"{line['label']} · 기준 대비 활동 증감 × 회당 비용 (인상 전)",
                      lambda c: f"ROUND({lookup(source,c,NET_ACTIVITY)}*$G${row+4},0)/Unit")
        else:
            param(4, "기준 월 비용 (원)", "base_monthly")
        if "monthly" in p and not link:
            m.input(sheet, row+6, f"{line['label']} 월별 인상 전 비용 (원)",
                    [p["monthly"].get(t) for t in m.months], line["label"],
                    "추정 월에 입력한 기준비용에 인상률을 적용합니다. 빈칸이면 공통 기준 월 비용을 씁니다.")
        expr = lambda c: f"{fixed_cost_basis(line,row,c)}*{fixed_multiplier(c)}"
    elif method == "per_headcount":
        role = p.get("role")
        if role is not None and role not in roles:
            raise ValueError(f"{line['id']}: 인원 계획에 없는 직군입니다 ({role}).")
        counts = roles[role] if role is not None else [r for rows in roles.values() for r in rows]
        if "monthly_cost" in p:
            param(4, "1인 월 비용 (원)", "monthly_cost")
            weighted_counts = [r + 5 for r in counts]
            expr = lambda c: f"{_sum(weighted_counts,c,'인건비')}*$G${row+4}/Unit"
        else:
            payroll = [r + 2 for r in counts] if role is not None else [9]
            expr = lambda c: _sum(payroll, c, "인건비")
    elif method == "manual":
        _monthly_input(m, sheet, row + 4, f"{line['label']} 월 계획 (원)", p.get("monthly", {}),
                       line["label"], line["basis"])
        expr = lambda c: f"{c}{row+4}"
    else:
        raise ValueError(f"{line['id']}: 지원하지 않는 비용 방식입니다 ({method}).")
    month_gate = None
    if p.get("months"):
        month_gate = row + 7
        months = set(p["months"])
        flags = [int(int(month[-2:]) in months) if i > m.last else None
                 for i, month in enumerate(m.months)]
        m.input(sheet, month_gate, f"{line['label']} 적용 월 (1/0)", flags, line["label"],
                "추정 월에 비용을 반영하면 1, 미반영은 0", annual="none")
        _formulas(m, sheet, row + 2, "동인 기준 비용", lambda c: f"IF({c}{month_gate},{expr(c)},0)")
    else:
        _formulas(m, sheet, row + 2, "동인 기준 비용", expr)
    timing = retro_adjustment(m, line, row, roles)
    _line_total(m, line, sheet, row, row + 1, row + 2, row + 3, ext_rows, month_gate, timing)


def build_drivers(m):
    revenues = m.spec["revenue_lines"]
    costs = m.spec["cost_lines"]
    lines = revenues + costs
    ids = {line["id"] for line in lines}
    if len(ids) != len(lines):
        raise ValueError("매출·비용 라인의 id가 중복입니다.")
    if any(line["section"] not in SECTIONS[1:3] for line in costs):
        raise ValueError("비용 라인의 구분은 매출원가 또는 판관비여야 합니다.")
    _cash_assumptions(m)
    actuals = _actual_is(m, lines)
    extensions = _extensions(m, ids)
    roles = _wages(m)
    m.renewal_probability_refs = {}
    volumes = defaultdict(dict)
    for record in m.data.get("sales_volume", []):
        volumes[record["line"]][record["month"]] = record
    items = defaultdict(list)
    contract_ids = {line["id"] for line in revenues if line["driver"] == "contracts"}
    for item in m.data.get("revenue_items", []):
        if item["line"] not in contract_ids:
            raise ValueError(f"{item['item']}: 계약 방식의 매출 라인을 지정하세요 ({item['line']}).")
        items[item["line"]].append(item)

    revenue_sheet, cost_sheet = "운영수입", "운영비용"
    rev_rows = {line["id"]: row for row, line in enumerate(revenues, 12)}
    if m.activity_lines:
        from activity_model import setup_plans
        setup_plans(m)
    m.section(revenue_sheet, 7, "매출 요약")
    for row, line in enumerate(revenues, 12):
        key = line["id"]
        detail = _sheet(m.details[key])
        _revenue(m, line, actuals[key], items[key], extensions[key], volumes[key], rev_rows)
        _formulas(m, revenue_sheet, row, line["label"], lambda c: f"{detail}!{c}8")
        block = 14 + len(revenues) + 5 * (row - 12)
        m.section(revenue_sheet, block, line["label"])
        for offset, source, label in [(1, 9, "배정 계정 실적"), (2, 10, "동인 기준 매출"),
                                      (3, 11, "수동 조정"), (4, 8, "반영 매출")]:
            _formulas(m, revenue_sheet, block + offset, label, lambda c: f"{detail}!{c}{source}")
            m.level(revenue_sheet, block + offset, 4)
        m.total(revenue_sheet, block + 4)
    _formulas(m, revenue_sheet, 8, "매출 합계", lambda c: _sum(rev_rows.values(), c))
    m.total(revenue_sheet, 8)
    m.section(revenue_sheet, 11, "매출 라인별 금액")

    m.section(cost_sheet, 7, "운영비용 요약")
    cost_rows = defaultdict(list)
    dep_rows = defaultdict(list)
    designated = set(m.spec["investment"]["depreciation_accounts"])
    dep_lines = [line for line in costs if designated.intersection(line["accounts"])]
    first_actual = m.col(max(0, m.last - 11))
    last_actual = m.col(m.last)
    m.section(cost_sheet, 13, "비용 라인별 금액")
    dep_ids = {line["id"] for line in dep_lines}
    # Allocate every summary address up front. pct_of_line can depend on a
    # cost listed later in the spec; spec_io rejects dependency cycles.
    cost_summaries = {line["id"]: 14 + i for i, line in enumerate(costs)}
    m.cost_summary_rows = cost_summaries
    if m.activity_lines:
        from activity_model import build_metrics
        build_metrics(m)
    m.cost_detail_rows = {}
    m.noncash_cost_rows, m.noncash_cogs_rows = [], []
    row = 16 + len(costs)
    for i, line in enumerate(costs):
        key = line["id"]
        m.cost_detail_rows[key] = row
        summary = 14 + i
        m.section(cost_sheet, row - 1, line["label"])
        cost_rows[line["section"]].append(summary)
        _formulas(m, cost_sheet, summary, line["label"], lambda c: f"{c}{row}")
        if key in dep_ids:
            dep_rows[line["section"]].append(summary)
            _formulas(m, cost_sheet, row + 1, "배정 계정 실적",
                      lambda c: _sum(actuals[key], c, "실적_IS"))
            # Only historical cells enter the weights. Each forecast links one investment pool.
            share = (f"IFERROR(SUM(${first_actual}${row+1}:${last_actual}${row+1})"
                     f"/SUM('실적_IS'!${first_actual}${m.actual_dep_row}:"
                     f"${last_actual}${m.actual_dep_row}),1/{len(dep_lines)})")
            _formulas(m, cost_sheet, row, line["label"], lambda c:
                      f"IF('INDEX'!{c}$6,{c}{row+1},'투자비'!{c}9*{share})")
        else:
            _cost(m, line, row, actuals[key], rev_rows, cost_summaries, roles, extensions[key])
            if line.get("noncash"):
                m.noncash_cost_rows.append(summary)
                if line["section"] == "매출원가":
                    m.noncash_cogs_rows.append(summary)
        m.total(cost_sheet, row)
        row += 9 if line["params"].get("months") else 8
    from wage_followup import finish_payable
    finish_payable(m)
    if not dep_lines:
        row = 14 + len(costs)
        _formulas(m, cost_sheet, row, "투자자산 감가상각비",
                  lambda c: f"IF('INDEX'!{c}$6,0,'투자비'!{c}9)")
        cost_rows["매출원가"].append(row)
        dep_rows["매출원가"].append(row)
    for row, section in enumerate(SECTIONS[1:3], 8):
        _formulas(m, cost_sheet, row, section, lambda c: _sum(cost_rows[section], c))
    _formulas(m, cost_sheet, 10, "운영비용 합계", lambda c: f"{c}8+{c}9")
    m.total(cost_sheet, 10)
    _formulas(m, cost_sheet, 11, "매출원가 중 감가상각비", lambda c: _sum(dep_rows["매출원가"], c))
    _formulas(m, cost_sheet, 12, "감가상각비 합계", lambda c: f"{c}11+{_sum(dep_rows['판관비'], c)}")

    if m.activity_lines:
        from activity_model import finish_activity
        finish_activity(m)
    _formulas(m, "검증", 100, "매출·비용 라인 실적 배정 차이 (원)", lambda c:
              f"IF('INDEX'!{c}$6,ROUND((ABS('운영수입'!{c}8-'실적_IS'!{c}8)"
              f"+ABS('운영비용'!{c}8-'실적_IS'!{c}9)+ABS('운영비용'!{c}9-'실적_IS'!{c}10))*Unit,0),0)", "max")
    last = m.col(len(m.months) - 1)
    m.add_check("매출·비용 라인 실적 대사", f'=IF(MAX(검증!J100:{last}100)=0,"OK","NO")')
