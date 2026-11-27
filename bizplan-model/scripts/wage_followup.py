"""Monthly wage entitlements, agreement settlement and the related accrued liability."""
from collections import defaultdict
from datetime import date

from openpyxl.utils.datetime import to_excel

from driver_blocks import _formulas, _sum, fixed_cost_basis

WAGES = "인건비"
MODE_LABEL = "소급 인식 방식 (0 발생, 1 일시)"
EFFECTIVE_LABEL = "소급 적용 월"
SETTLED_LABEL = "소급 정산 월"


def validate_wages(spec):
    wages = spec["wages"]
    retro = wages.get("retro")
    if retro:
        start, settled = retro["effective_from"], retro["settled"]
        if not spec["last_actual"] < start <= spec["plan_end"] or settled < start:
            raise ValueError("wages.retro: 적용 월은 추정 구간, 정산 월은 적용 월 이후여야 합니다.")
    by_role = defaultdict(list)
    for record in wages["roles"]:
        name = record["role"]
        by_role[name].append(record)
        first = record.get("raise_from")
        if retro and first and (first < retro["effective_from"] or
                                record["start"] > spec["last_actual"] and first < record["start"]):
            raise ValueError(f"{name}: 인상 시작 월이 협약 적용 월 또는 추정 입사 월보다 빠릅니다.")
    for line in spec["cost_lines"]:
        name = line["params"].get("role")
        if not name or line["method"] not in ("fixed_inflation", "mixed"):
            continue
        if not line.get("follows_wage_raise") or name not in by_role:
            raise ValueError(f"{line['id']}: 등록 직군과 follows_wage_raise를 함께 지정하세요.")
        # A fixed account has a single policy even if the role has several hiring cohorts.
        policies = {(r.get("annual_raise", wages["annual_raise"]), first_raise_month(spec, r))
                    for r in by_role[name]}
        if len(policies) > 1:
            raise ValueError(f"{line['id']}: 고정비에 지정한 직군의 인상률·최초 인상 월이 서로 다릅니다.")


def first_raise_month(spec, record):
    """Default entry rates for new hires already in the agreement's pay regime."""
    if "raise_from" in record:
        return record["raise_from"]
    cutoff, start = spec["last_actual"], record["start"]
    retro = spec["wages"].get("retro")
    if retro and (start <= cutoff or start < retro["effective_from"]):
        return retro["effective_from"]
    return f"{max(int(cutoff[:4]), int(start[:4])) + 1:04}-01"


def salary_factor(col, rate, first):
    return (f"IF({col}$5<{first},1,(1+{rate})^(1+YEAR({col}$5)-YEAR({first})"
            f"-(MONTH({col}$5)<MONTH({first}))))")


def prepare_controls(m, records):
    top = 16 + 6 * len(records)
    m.fixed_wage_roles = dict.fromkeys(line["params"]["role"] for line in m.spec["cost_lines"]
                                      if line.get("follows_wage_raise") and "role" in line["params"])
    m.wage_rates = {}
    m.wage_policy_by_row = {}
    m.wage_date_rows = set()
    m.wage_retro_recognized_rows = {}
    m.wage_payable_rows = []
    m.wage_payable_row = None
    retro = m.spec["wages"].get("retro")
    if retro:
        for offset, key, label, name in ((0, "effective_from", EFFECTIVE_LABEL, "WG_Effective"),
                                         (1, "settled", SETTLED_LABEL, "WG_Settled")):
            r = top + offset
            value = to_excel(date.fromisoformat(retro[key] + "-01"))
            m.param(WAGES, r, label, value, WAGES, "임금 협약의 적용·지급 월")
            m.named(name, WAGES, f"$G${r}")
            m.wage_date_rows.add(r)
        m.param(WAGES, top + 2, MODE_LABEL, int(retro["recognition"] == "lump"), WAGES,
                "0 발생 기준, 1 정산 월 일시 인식. 두 방식 모두 정산 월에 소급분 지급")
        m.add_check("임금 소급 협약 입력 범위",
                    '=IF(AND(WG_Effective>LastActual,WG_Settled>=WG_Effective,'
                    'OR(WG_Recognition=0,WG_Recognition=1)),"OK","NO")')
        top += 3
    for i, record in enumerate(records):
        if "annual_raise" in record:
            m.param(WAGES, top, record["role"] + " 연 임금 인상률", record["annual_raise"],
                    WAGES, "해당 직군의 연 인상률. 공통 인상률에 우선함")
            m.wage_rates[i] = f"$G${top}"
            top += 1
    m.next_wage_row = top + 1


def check_fixed_cohorts(m, roles):
    """A single fixed account must continue to follow a consistent role policy after edits."""
    for name in m.fixed_wage_roles:
        cohorts = roles[name]
        if len(cohorts) < 2:
            continue
        first, rate = m.wage_policy_by_row[cohorts[0]]
        begin = m.next_wage_row
        for cohort in cohorts[1:]:
            row = m.next_wage_row
            m.next_wage_row += 1
            compare_first, compare_rate = m.wage_policy_by_row[cohort]
            m.cell(WAGES, row, 2, name + " 고정비 연동 임금 기준 일치", "text")
            m.cell(WAGES, row, 7, f"=--AND({first}={compare_first},{rate}={compare_rate})")
        bounds = f"'{WAGES}'!$G${begin}:$G${m.next_wage_row-1}"
        m.add_check(name + " 고정비 연동 직군 인상률·시행월 일치",
                    f'=IF(MIN({bounds})=1,"OK","NO")')


def fixed_factor(m, line, roles, col):
    """Return a wage schedule for the fixed component, or use the existing annual rule."""
    if not line.get("follows_wage_raise"):
        return None
    name = line["params"].get("role")
    if name:
        return f"'{WAGES}'!{col}{roles[name][0]+4}"
    if m.spec["wages"].get("retro"):
        return f"'{WAGES}'!{col}11"
    return None


def retro_adjustment(m, line, cost_row, roles):
    """Track amounts actually earned by the specified cost line before settlement."""
    if not m.spec["wages"].get("retro"):
        return None
    method = line["method"]
    is_fixed = method in ("fixed_inflation", "mixed") and line.get("follows_wage_raise")
    if method != "per_headcount" and not is_fixed:
        return None

    params = line["params"]
    name = params.get("role")
    cohorts = roles[name] if name else [r for rows in roles.values() for r in rows]
    cost = f"'운영비용'!$G${cost_row+4}/Unit"

    def premium(col):
        if is_fixed:
            factor = fixed_factor(m, line, roles, col)
            basis = fixed_cost_basis(line, cost_row, col, "'운영비용'!")
            return f"({factor}-1)*{basis}"
        terms = []
        for r in cohorts:
            amount = cost if "monthly_cost" in params else f"$G${r+1}/Unit"
            terms.append(f"{col}{r}*({col}{r+4}-1)*{amount}")
        return "(" + "+".join(terms or ["0"]) + ")"

    accrual = m.next_wage_row
    m.next_wage_row += 4
    outstanding, timing, recognized = accrual + 1, accrual + 2, accrual + 3
    m.wage_retro_recognized_rows[line["id"]] = recognized
    gate = f"*'운영비용'!{{c}}{cost_row+7}" if params.get("months") else ""
    factor = m.factor(line["id"])
    _formulas(m, WAGES, accrual, line["label"] + " 소급 대상 증감액",
              lambda c: f"IF(AND({c}$5>LastActual,{c}$5>=WG_Effective,{c}$5<WG_Settled),"
                        f"({premium(c)})*{factor}{gate.format(c=c)},0)")
    balances, expense_changes, recognized_premiums = [], [], []
    for i in range(len(m.months)):
        col = m.col(i)
        prev = f"{m.col(i-1)}{outstanding}" if i else "0"
        balances.append(f"=IF({col}$5<WG_Settled,{prev}+{col}{accrual},0)")
        expense_changes.append(f"=WG_Recognition*({prev}-{col}{outstanding})")
        # This expense has its own settlement schedule and is excluded from turnover estimates.
        recognized_premiums.append(f"=IF(WG_Recognition=0,{col}{accrual},IF({col}$5=WG_Settled,{prev},0))")
    m.row(WAGES, outstanding, line["label"] + " 정산 전 소급 누계", balances, annual="last")
    m.row(WAGES, timing, line["label"] + " 정산 월 인식 조정", expense_changes)
    m.row(WAGES, recognized, line["label"] + " 별도 정산 소급분 손익 인식액", recognized_premiums)
    if not line.get("noncash"):
        m.wage_payable_rows.append(outstanding)
    return timing


def finish_payable(m):
    if not m.spec["wages"].get("retro"):
        return
    row = m.next_wage_row
    m.wage_payable_row = row
    _formulas(m, WAGES, row, "임금 소급 미지급비용 기말잔액",
              lambda c: f"(1-WG_Recognition)*{_sum(m.wage_payable_rows,c)}", "last")
    m.total(WAGES, row)
    m.named("WG_Payable", WAGES, f"$J${row}:${m.col(len(m.months)-1)}${row}")
