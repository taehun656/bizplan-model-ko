"""Optional equity, nonoperating balance changes and deferred payroll cash."""
from collections import defaultdict

from spec_io import check_forecast
from tax_cash import opening_accounts


DIVIDEND = 73
CONTRIBUTION = 74
ASSET_CASH = 75
LIABILITY_CASH = 76
WAGE_CHANGE = 77
GROUP_ROWS = {11: 79, 13: 80, 16: 81, 18: 82}
RETAINED_NAMES = {
    "이익잉여금", "미처분이익잉여금", "결손금", "미처리결손금", "이익준비금",
    "임의적립금", "법정적립금", "retainedearnings", "accumulateddeficit"}
PAID_NAMES = {
    "납입자본", "자본금", "자본잉여금", "주식발행초과금", "보통주자본금", "우선주자본금",
    "sharecapital", "paidincapital", "additionalpaidincapital", "capitalstock"}


def account_schedules(spec, data):
    """Group inputs by exact BS identity and business-plan caption."""
    from finance_blocks import _bucket, _name
    sections = defaultdict(set)
    for entry in data.get("bs_monthly", []):
        sections[entry["account"]].add(entry["section"])
    schedules = defaultdict(lambda: defaultdict(lambda: defaultdict(int)))
    managed_tax = set().union(*opening_accounts(spec.get("tax", {})).values())
    for item in spec.get("other_items", []):
        name = item["bs_account"]
        choices = sections[name]
        section = item.get("bs_section")
        if section is None:
            if len(choices) != 1:
                raise ValueError(f"other_items {name}: 신규 또는 중복 계정은 bs_section을 지정하세요.")
            section = next(iter(choices))
        if choices and section not in choices:
            raise ValueError(f"other_items {name}: 지정한 구분 {section}이 실적 계정 구분과 다릅니다.")
        group = _bucket(section, name)
        if group not in GROUP_ROWS or _name(name) in managed_tax:
            raise ValueError(f"other_items {name}: 모델의 전용 계산 계정에는 기타 증감을 지정할 수 없습니다.")
        schedules[section, name][item["label"]][item["month"]] += item["amount"]
    return schedules


def _cutoff_equity(spec, data, names):
    """Return None when the source does not identify a category at cutoff."""
    from finance_blocks import _name
    amounts = [r["amount"] for r in data["bs_monthly"]
               if r["section"] == "자본" and r["month"] == spec["last_actual"]
               and _name(r["account"]) in names]
    return sum(amounts) if amounts else None


def validate_capital(spec, data):
    for key, schedule in spec.get("equity", {}).items():
        check_forecast(spec, schedule, "equity." + key)
    check_forecast(spec, (item["month"] for item in spec.get("other_items", [])), "other_items")
    capital = spec.get("equity", {}).get("contributions", {})
    if capital and spec["statements"]["bs"]:
        paid = _cutoff_equity(spec, data, PAID_NAMES)
        if paid is not None:
            for month, amount in sorted(capital.items()):
                paid += amount
                if paid < 0:
                    raise ValueError(f"equity.contributions {month}: 감자 후 납입자본 잔액이 음수입니다.")
    if not spec.get("other_items"):
        return
    cutoff = spec["last_actual"]
    opening = {(r["section"], r["account"]): r["amount"]
               for r in data["bs_monthly"] if r["month"] == cutoff}
    for account, labels in account_schedules(spec, data).items():
        balance = opening.get(account, 0)
        changes = defaultdict(int)
        for monthly in labels.values():
            for month, amount in monthly.items():
                changes[month] += amount
        # Contra balances are carried at their source sign. For other accounts
        # a recovery or repayment cannot use more than the opening balance.
        if balance >= 0:
            for month, amount in sorted(changes.items()):
                balance += amount
                if balance < 0:
                    raise ValueError(f"other_items {account[1]} {month}: 예정 회수·상환이 계정 잔액을 초과합니다.")


def add_terms(formula, column, rows, prefix=""):
    """Append active cash components, keeping legacy formulas identical."""
    return formula + "".join(f"+{prefix}{column}{r}" for r in rows)


def _monthly_input(m, index, label, schedule):
    forecast_cols = f"{m.col(m.last + 1)}{index}:{m.col(len(m.months) - 1)}{index}"
    values = [schedule.get(month, 0) if i > m.last else None for i, month in enumerate(m.months)]
    m.row("CF", index, label, values, kind="input")
    m.register("CF", forecast_cols, label, "자금", "추정 월 원 단위 직접 입력")


def _historical_sum(rows, column):
    terms = [f"'실적_BS'!{column}{r}" for r in rows]
    return "(" + "+".join(terms) + ")/Unit" if terms else "0"


def build_cash_items(m, row, forecast, observed):
    """Add CF detail before funding. Return the rows to include in each subtotal."""
    from finance_blocks import _bucket, _name
    equity = m.spec.get("equity", {})
    dividends = equity.get("dividends", {})
    contributions = equity.get("contributions", {})
    schedules = account_schedules(m.spec, m.data) if m.spec.get("other_items") else {}
    wages = getattr(m, "wage_payable_row", None)
    effects = {"operating": [], "investing": [], "financing": [], "bs": {}}
    m.dividend_cash_row = DIVIDEND if dividends else None
    if not any((dividends, contributions, schedules, wages)):
        return effects
    m.section("CF", 70, "계획 배당·자본·기타 현금흐름")
    for schedule, raw, signed, name, sign in (
            (dividends, 71, DIVIDEND, "배당 지급", "-"),
            (contributions, 72, CONTRIBUTION, "유상증자·감자", "")):
        if schedule:
            _monthly_input(m, raw, name + " 계획 (원)", schedule)
            forecast("CF", signed, name, lambda i,t,c,p: f"{sign}{c}{raw}/Unit")
            effects["financing"].append(signed)
    if dividends:
        m.add_check("계획 배당 지급액 비음수",
                    f'=IF(MIN(CF!{m.col(m.last+1)}71:{m.col(len(m.months)-1)}71)>=0,"OK","NO")')
    if wages is not None:
        forecast("CF", WAGE_CHANGE, "소급 임금 미지급비용 증감",
                 lambda i,t,c,p: f"'인건비'!{c}{wages}-'인건비'!{p}{wages}")
        effects["operating"].append(WAGE_CHANGE)
        if m.bs:
            forecast("BS명세", 27, "소급 임금 미지급비용", lambda i,t,c,p: f"'인건비'!{c}{wages}", "last")
    if m.bs and (dividends or contributions):
        profit = m.is_rows["net"]
        for r, label, names, delta in ((24, "이익잉여금", RETAINED_NAMES, DIVIDEND if dividends else None),
                                       (25, "납입자본", PAID_NAMES, CONTRIBUTION if contributions else None)):
            source = [n for (s, a), n in m.bs_source_rows.items() if s == "자본" and _name(a) in names]
            identified = _cutoff_equity(m.spec, m.data, names) is not None
            if not identified:
                label += " 계획 증감 누계 (기초 미구분)"
            elif r == 25 and contributions:
                first, end = m.col(m.last+1), m.col(len(m.months)-1)
                m.add_check("감자 후 납입자본 잔액", f'=IF(MIN(BS명세!{first}25:{end}25)>=0,"OK","NO")')
            def amount(i,t,c,p):
                if i <= m.last:
                    return _historical_sum(source, c) if identified and t in observed else None
                base = f"{p}{r}" + (f"+'IS'!{c}{profit}" if r == 24 else "")
                return base + (f"+'CF'!{c}{delta}" if delta is not None else "")
            row("BS명세", r, label, amount, "last")
        row("BS명세", 26, "기타 자본", lambda i,t,c,p:
            f"'실적_BS'!{c}19-{c}24-{c}25" if t in observed else (f"{p}26" if i>m.last else None), "last")
    if not schedules:
        return effects
    index = 85
    group_inputs = defaultdict(list)
    closing_rows = []
    for n, (identity, labels) in enumerate(sorted(schedules.items(), key=lambda v: (_bucket(*v[0]), v[0]))):
        section, account = identity
        group = _bucket(section, account)
        first = index
        for label, monthly in sorted(labels.items()):
            _monthly_input(m, index, f"{section} · {account} / {label} 증감 (원)", monthly)
            group_inputs[group].append(index)
            index += 1
        if m.bs:
            detail = 31 + n
            source = m.bs_source_rows.get(identity)
            def account_balance(i,t,c,p):
                if i <= m.last:
                    return _historical_sum([source] if source is not None else [], c) if t in observed else None
                return f"{p}{detail}+SUM('CF'!{c}{first}:{c}{index-1})/Unit"
            row("BS명세", detail, account + " 기말 잔액", account_balance, "last")
            closing_rows.append((detail, account))
    for group, indices in sorted(group_inputs.items()):
        r = GROUP_ROWS[group]
        forecast("CF", r, f"기타 {('유동자산','비유동자산','유동부채','비유동부채')[tuple(GROUP_ROWS).index(group)]} 계획 증감",
                 lambda i,t,c,p: f"SUM({c}{indices[0]}:{c}{indices[-1]})/Unit")
        effects["bs"][group] = r
    for groups, raw, kind, label, sign in (
            ((11,13), ASSET_CASH, "investing", "기타 비영업 자산 증감의 현금 효과", "-"),
            ((16,18), LIABILITY_CASH, "financing", "기타 비영업 부채 증감의 현금 효과", "")):
        selected = [GROUP_ROWS[g] for g in groups if g in group_inputs]
        if selected:
            forecast("CF", raw, label, lambda i,t,c,p: sign+"("+"+".join(c+str(r) for r in selected)+")")
            effects[kind].append(raw)
    if closing_rows:
        # Reconcile again after a monthly input is edited directly in Excel.
        for r, label in closing_rows:
            first, end = m.col(m.last+1), m.col(len(m.months)-1)
            cutoff = m.col(m.last)
            m.add_check(f"{label} 계획 잔액", f'=IF(OR(\'BS명세\'!{cutoff}{r}<0,'
                        f'MIN(\'BS명세\'!{first}{r}:{end}{r})>=-0.5/Unit),"OK","NO")')
    return effects

