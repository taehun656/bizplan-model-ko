"""Korean current tax payments and balance reconciliation (references/tax_rules_kr.md §5).

All worksheet amounts are in the workbook's display unit. Prepayments in excess
of accrued tax are carried as a receivable, never as a negative payable.
"""
import re



from tax_policy import (INTERIM_PERIOD_MONTHS as HALF_YEAR, INTERIM_DUE_MONTH,
                        NATIONAL_DUE_MONTHS as NATIONAL_LAG, LOCAL_DUE_MONTHS as LOCAL_LAG,
                        INTERIM_SME_EXEMPT_BELOW as SME_EXEMPT_WON,
                        INSTALLMENT_THRESHOLD as INSTALLMENT_THRESHOLD_WON, credit_mode)

ADVANCE_LAG = INTERIM_DUE_MONTH - HALF_YEAR
OPENING_ACCOUNTS = {
    "national_payable": {"미지급법인세", "당기법인세부채", "미지급법인세비용", "법인세미지급금",
                         "법인세부채", "미지급국세", "currenttaxliabilities", "incometaxpayable"},
    "local_payable": {"미지급법인지방소득세", "미지급지방소득세", "지방소득세미지급금"},
    "national_prepaid": {"선납세금", "선납법인세", "당기법인세자산", "법인세환급금", "prepaidincometax"},
    "local_prepaid": {"선납지방소득세", "선납법인지방소득세", "지방소득세환급금"},
}
RURAL_PAYABLE_ACCOUNTS = {"미지급농어촌특별세", "농어촌특별세미지급금", "농특세미지급금", "미지급농특세"}


def opening_accounts(tax):
    if not credit_mode(tax):
        return OPENING_ACCOUNTS
    return {**OPENING_ACCOUNTS, "national_payable": OPENING_ACCOUNTS["national_payable"] | RURAL_PAYABLE_ACCOUNTS}


def _bs_balance(m, component, column, aliases=None):
    allowed = OPENING_ACCOUNTS[component] if aliases is None else aliases
    section = "유동부채" if component.endswith("payable") else "유동자산"
    rows = []
    for (row, col), label in m.grid.get("실적_BS", {}).items():
        if col != 2 or not isinstance(label, str) or not label.lstrip("'").startswith(section + " /"):
            continue
        caption = label.rsplit(" / ", 1)[-1].removesuffix(" (원)")
        if re.sub(r"[\W_]", "", caption.casefold()) in allowed:
            rows.append(f"'실적_BS'!{column}{row}")
    return "(" + "+".join(rows) + ")/Unit" if rows else "0"


def build_tax_cash(m, row, forecast, check, isr):
    """Build statutory due events from the fiscal-year schedule on IS."""
    cash = "CF"
    last = m.last
    cutoff = m.col(last)
    inc = m.tax_rows
    refs = m.tax_refs
    has_rural = "rural_tax" in inc

    def iscell(c, key):
        return f"'IS'!{c}{inc[key]}"

    def scalar(key):
        return refs[key]

    def prior(key):
        return refs["prior", key]

    mixed = f"LastActual<{iscell(cutoff,'fy_end')}"
    booked = iscell(cutoff,'actual_ytd_tax')
    documented = "+".join(iscell(cutoff,k) for k in ("interim_paid","withholding","ad_hoc"))
    paid = {j:refs["paid_to_date",j] for j in ("national","local")}
    rural_source = _bs_balance(m, "national_payable", cutoff, RURAL_PAYABLE_ACCOUNTS) if has_rural and m.bs else "0"

    def reported(component, default):
        override = refs["opening",component]
        source = _bs_balance(m,component,cutoff) if m.bs else default
        if component == "national_payable" and has_rural and m.bs:
            source = f"({source})+({rural_source})"
        return f"IF(ISNUMBER({override}),{override},{source})"

    # The historical IS reports a combined tax expense. Identify the local
    # portion from jurisdiction-specific payments/balances, leaving the rest
    # national. A stock due for an earlier year is not current-year expense.
    local_stock = f"({reported('local_payable','0')}-{reported('local_prepaid','0')})"
    current_local_stock = f"IF(EDATE({iscell(cutoff,'fy_start')},{LOCAL_LAG}-1+{scalar('faithful_filing')})>LastActual,0,{local_stock})"
    national_stock = f"({reported('national_payable','0')}-{reported('national_prepaid','0')})"
    has_national_stock = ("TRUE" if m.bs else
                          f"OR(ISNUMBER({refs['opening','national_payable']}),ISNUMBER({refs['opening','national_prepaid']}))")
    national_current = f"EDATE({iscell(cutoff,'fy_start')},{NATIONAL_LAG}-1+{scalar('faithful_filing')})<=LastActual"
    # If only national remittances are documented, a reported current-year
    # national balance also identifies the local share of booked expense.
    local_split = (f"IF(AND(ISNUMBER({paid['national']}),NOT(ISNUMBER({paid['local']})),"
                   f"{has_national_stock},{national_current}),{booked}-{paid['national']}-{national_stock},"
                   f"N({paid['local']})+{current_local_stock})")
    local_booked = f"MIN(MAX(0,{booked}),MAX(0,{local_split}))"
    row(cash,63,"마감 사업연도 실적 중 지방세 인식분",lambda i,t,c,p:
        f"IF({mixed},{local_booked},0)" if i==last else None,"last")
    recognized = {"national":f"({booked}-${cutoff}$63)", "local":f"${cutoff}$63"}
    known = {"national":documented, "local":"0"}

    def opening(component):
        j,kind = component.split("_")
        sign = "-" if kind=="prepaid" else ""
        actual_paid = f"IF(ISNUMBER({paid[j]}),{paid[j]},{known[j]})"
        net = f"({recognized[j]}-({actual_paid}))"
        return reported(component,f"IF({mixed},MAX(0,{sign}{net}),0)")

    for r, key, label in (
        (40,"national_payable","마감 국세 미지급 잔액"),
        (41,"local_payable","마감 지방세 미지급 잔액"),
        (42,"national_prepaid","마감 국세 선납 잔액"),
        (43,"local_prepaid","마감 지방세 선납 잔액"),
    ):
        row(cash, r, label, lambda i,t,c,p: opening(key) if i == last else None, "last")

    m.section(cash, 30, "법인세 납부와 미지급·선납 잔액")
    faithful = scalar("faithful_filing")
    sme = scalar("is_sme")
    installments = scalar("installments")
    for r, lag, title in ((44,NATIONAL_LAG,"마감 국세의 다음 확정 납부 월"),
                          (45,LOCAL_LAG,"마감 지방세의 다음 확정 납부 월")):
        def first_due(i,t,c,p):
            if i != last:
                return None
            start, length = iscell(c,"fy_start"), iscell(c,"fy_length")
            previous_due = f"EDATE({start},{lag}-1+{faithful})"
            return f"IF({previous_due}>{c}$5,{previous_due},EDATE({start},{length}-1+{lag}+{faithful}))"
        row(cash,r,title,first_due,"none")
        m.format_row(cash,r,"yyyy-mm")

    for j, paid_row, old_row, liability, asset, due, lag in (
            ("national",61,64,40,42,44,NATIONAL_LAG), ("local",62,65,41,43,45,LOCAL_LAG)):
        net = f"(${cutoff}${liability}-${cutoff}${asset})"
        prior_due = f"EDATE(${cutoff}${due},-({lag}+{faithful}))<=LastActual"
        # If an earlier return is still open, allocate only the stock needed
        # for the current FY and settle the remainder on that earlier deadline.
        current_net = f"IF({prior_due},MAX(0,{net}),{net})"
        inferred = f"MAX({known[j]},{recognized[j]}-{current_net},0)"
        amount = f"IF({mixed},IF(ISNUMBER({paid[j]}),{paid[j]},{inferred}),0)"
        title = "국세" if j=="national" else "지방세"
        row(cash,paid_row,f"마감 사업연도 {title} 누적 기납부액",lambda i,t,c,p:
            amount if i==last else None,"last")
        prior_stock = f"{net}-IF({mixed},{recognized[j]}-${cutoff}${paid_row},0)"
        row(cash,old_row,f"마감 잔액 중 확정 신고 대상 {title}",lambda i,t,c,p:
            f"IF({prior_due},{prior_stock},0)" if i==last else None,"last")

    if has_rural:
        # Actuals carry a combined tax expense and may carry a combined liability.
        # Split only the part supported by the booked expense and opening stock.
        def current_rural_open(i,t,c,p):
            if i != last:
                return None
            current_stock = f"MAX(0,{c}40-MAX(0,{c}64))"
            booked_share = f"MIN(MAX(0,{booked}-{c}63),{iscell(c,'rural_tax')})"
            if rural_source == "0":
                share = f"MIN({booked_share},{current_stock})"
            else:
                earlier_return_open = f"EDATE({c}44,-({NATIONAL_LAG}+{faithful}))<=LastActual"
                # With one open FY a separately reported rural liability gives
                # the actual split, even when there is a national prepayment.
                share = (f"MIN(MAX(0,{rural_source}),{c}40,"
                         f"IF({earlier_return_open},{booked_share},{c}40))")
            return f"IF({mixed},{share},0)"
        row(cash,35,"마감 사업연도 실적 중 농어촌특별세 인식분",current_rural_open,"last")

        def old_rural_open(i,t,c,p):
            if i != last:
                return None
            closed = f"EDATE({c}44,-({NATIONAL_LAG}+{faithful}))"
            historical = (f"SUMIF($J$5:{c}$5,{closed},"
                          f"'IS'!$J${inc['rural_tax']}:{c}${inc['rural_tax']})")
            share = historical if rural_source == "0" else f"MAX(0,({rural_source})-{c}35)"
            return f"IF({closed}<=LastActual,MIN(MAX(0,{c}40-{c}35),{share}),0)"
        row(cash,36,"마감 잔액 중 확정 신고 대상 농어촌특별세",old_rural_open,"last")

    # The advance belongs to the FY of its six-month base. In a seven-month
    # short FY its due month can fall AFTER that FY's close.
    def advance_owner(i,t,c,p):
        if i < ADVANCE_LAG:
            # With just July actuals the preceding June can be outside the
            # visible axis. A supplied prior assessment still makes August due.
            start = iscell(c,"fy_start")
            prior_short = (f"IF(AND({iscell(p,'fy_month')}={HALF_YEAR+1},"
                           f"{iscell(p,'fy_length')}={HALF_YEAR+1},"
                           f"{iscell(p,'fy_start')}<>{scalar('founded_fy')}),{iscell(p,'fy_start')},0)")
            return (f"IF(AND({iscell(c,'fy_month')}={INTERIM_DUE_MONTH},"
                    f"{start}<>{scalar('founded_fy')}),{start},{prior_short})")
        h = m.col(i-ADVANCE_LAG)
        start, month, length = (iscell(h,k) for k in ("fy_start","fy_month","fy_length"))
        founded = scalar("founded_fy")
        return f"IF(AND({month}={HALF_YEAR},{length}>{HALF_YEAR},{start}<>{founded}),{start},0)"
    forecast(cash,46,"중간예납 대상 사업연도 개시 월",advance_owner,"none")
    m.format_row(cash,46,'yyyy-mm;;"-"')

    def previous_year_value(i, key):
        h = m.col(i - ADVANCE_LAG)
        target = f"EDATE({iscell(h,'fy_start')},-1)"
        data = f"'IS'!$J${inc[key]}:{h}${inc[key]}"
        dates = f"$J$5:{h}$5"
        return f"INDEX({data},1,MATCH({target},{dates},0))"

    def advance(i,t,c,p):
        h = m.col(max(0,i-ADVANCE_LAG))
        calculated = "0" if i < ADVANCE_LAG else (f"MAX(0,{previous_year_value(i,'national_net')}-{previous_year_value(i,'withholding')}"
                      f"-{previous_year_value(i,'ad_hoc')})*{HALF_YEAR}/{previous_year_value(i,'fy_length')}")
        prev_gross = prior("assessed_tax")
        supplied = (f"MAX(0,{prev_gross}-{prior('credits')}-{prior('withholding')}-{prior('ad_hoc')})"
                    f"*{HALF_YEAR}/{prior('months')}")
        first = iscell(m.col(min(last+1,len(m.months)-1)),"fy_start")
        interim_base = "half_year_net" if has_rural else "half_year_gross"
        fallback = iscell(h,interim_base) if i >= ADVANCE_LAG else "0"
        supplied_available = f"ISNUMBER({prev_gross})"
        if has_rural:
            supplied_available = f"AND({supplied_available},{prev_gross}>0)"
            if i >= ADVANCE_LAG:
                calculated = f"IF({previous_year_value(i,'national_gross')}>0,{calculated},{fallback})"
        raw = (f"IF(EDATE({c}46,-1)>LastActual,{calculated},"
               f"IF(AND({c}46={first},{supplied_available}),{supplied},{fallback}))")
        return f"IF({c}46=0,0,{raw})"
    forecast(cash,47,"중간예납 면제 판정 전 금액",advance)
    first_fy = iscell(m.col(min(last+1,len(m.months)-1)),"fy_start")
    def advance_cash(i,t,c,p):
        documented = (f"SUMIF({refs['annual_starts']},{c}46,{refs['annual','interim_paid']})")
        documented = (f"IF(AND({mixed},{c}46={iscell(cutoff,'fy_start')}),"
                      f"MIN({documented},${cutoff}$61),{documented})")
        return (f"IF(AND(IF({c}46={first_fy},{prior('is_sme')},{sme}),"
                f"{c}47<{SME_EXEMPT_WON}/Unit),0,MAX(0,{c}47-{documented}))")
    forecast(cash,48,"국세 중간예납",advance_cash)

    def at_deadline(i, lag, expression):
        choices = []
        for delay in (lag, lag + 1):
            end = i - delay
            if end <= last:
                choices.append("0")
                continue
            e = m.col(end)
            year_end = f"{iscell(e,'fy_month')}={iscell(e,'fy_length')}"
            choices.append(f"IF({year_end},{expression(e)},0)")
        return f"IF({faithful},{choices[1]},{choices[0]})"

    def initial_due(c, date_row, amount_row):
        return f"IF({c}$5=${cutoff}${date_row},${cutoff}${amount_row},0)"

    def national_settlement(i,t,c,p):
        def amount(e):
            owner = iscell(e,"fy_start")
            advance_paid = f"SUMIF($J$46:{p}46,{owner},$J$48:{p}48)"
            documented = "+".join(iscell(e,k) for k in ("interim_paid","withholding","ad_hoc"))
            credits = f"IF({owner}={iscell(cutoff,'fy_start')},${cutoff}$61,{documented})"
            return f"{iscell(e,'national_net')}-{advance_paid}-{credits}"
        result = f"{at_deadline(i,NATIONAL_LAG,amount)}+{initial_due(c,44,64)}"
        return result + (f"-{initial_due(c,44,36)}" if has_rural else "")
    forecast(cash,66,"국세 확정 신고 순납부액",national_settlement)
    forecast(cash,49,"국세 확정 기한의 납부할 세액",lambda i,t,c,p:f"MAX(0,{c}66)")
    forecast(cash,59,"국세 확정 신고 환급 유입",lambda i,t,c,p:f"MAX(0,-{c}66)")

    def deferred(i,t,c,p):
        elected = at_deadline(i,NATIONAL_LAG,lambda e: iscell(e,"installment_amount"))
        end = f"EDATE({c}$5,-({NATIONAL_LAG}+{faithful}))"
        start_range = f"'IS'!$J${inc['fy_start']}:{cutoff}${inc['fy_start']}"
        historical_start = (f"IFERROR(INDEX({start_range},1,MATCH({end},$J$5:{cutoff}$5,0)),"
                            f"EDATE({end},1-{prior('months')}))")
        historical = (f"IF(AND({c}$5=${cutoff}$44,{end}<=LastActual),"
                      f"SUMIF({refs['annual_starts']},{historical_start},{refs['annual','installment_amount']}),0)")
        return (f"IF(AND({installments},{c}49>{INSTALLMENT_THRESHOLD_WON}/Unit),"
                f"MIN({c}49,MAX({elected},{historical})),0)")
    forecast(cash,50,"국세 분납 이연액",deferred)
    forecast(cash,51,"국세 확정 납부",lambda i,t,c,p: f"{c}49-{c}50")
    forecast(cash,52,"국세 분납",lambda i,t,c,p:
             f"IF({sme},{m.col(i-2)}50,{p}50)" if i>=2 else f"IF({sme},0,{p}50)")

    def local_settlement(i,t,c,p):
        def amount(e):
            credit = f"IF({iscell(e,'fy_start')}={iscell(cutoff,'fy_start')},${cutoff}$62,0)"
            return f"{iscell(e,'local_tax')}-{credit}"
        return f"{at_deadline(i,LOCAL_LAG,amount)}+{initial_due(c,45,65)}"
    forecast(cash,67,"지방세 확정 신고 순납부액",local_settlement)
    forecast(cash,53,"법인지방소득세 납부",lambda i,t,c,p:f"MAX(0,{c}67)")
    forecast(cash,60,"지방세 확정 신고 환급 유입",lambda i,t,c,p:f"MAX(0,-{c}67)")

    if has_rural:
        # The rural tax is due in full at the national filing date. It neither
        # changes the interim base nor enters the national installment schedule.
        forecast(cash,69,"농어촌특별세 확정 납부",lambda i,t,c,p:
                 f"{at_deadline(i,NATIONAL_LAG,lambda e: iscell(e,'rural_tax'))}+{initial_due(c,44,36)}")
        forecast(cash,37,"당월 인식 농어촌특별세",lambda i,t,c,p:
                 f"{iscell(c,'rural_tax')}-IF({iscell(c,'fy_start')}={iscell(p,'fy_start')},"
                 f"IF({p}$5>LastActual,{iscell(p,'rural_tax')},{p}35),0)")
        row(cash,68,"농어촌특별세 미지급 잔액",lambda i,t,c,p:
            f"{c}35+{c}36" if i==last else (f"{p}68+{c}37-{c}69" if i>last else None),"last")

    # Undated remittances for an open historical FY are provided as amounts
    # already paid through cutoff. For entirely projected years show the annual
    # withholding/ad-hoc cash amounts at year-end when no monthly detail exists.
    forecast(cash,54,"국세 원천징수·수시부과 납부",lambda i,t,c,p:
             f"IF(AND({iscell(c,'fy_start')}>LastActual,{iscell(c,'fy_month')}={iscell(c,'fy_length')}),"
             f"{iscell(c,'withholding')}+{iscell(c,'ad_hoc')},0)")
    forecast(cash,33,"법인세 현금 납부 합계",lambda i,t,c,p:
             f"SUM({c}48,{c}51:{c}54)-SUM({c}59:{c}60)" + (f"+{c}69" if has_rural else ""))
    forecast(cash,34,"당월 인식 법인세비용",lambda i,t,c,p:f"'IS'!{c}{isr['tax']}")
    forecast(cash,55,"당월 인식 법인지방소득세",lambda i,t,c,p:
             f"{iscell(c,'local_tax')}-IF(AND({p}$5>LastActual,{iscell(p,'fy_start')}={iscell(c,'fy_start')}),"
             f"{iscell(p,'local_tax')},IF({iscell(c,'fy_start')}={iscell(cutoff,'fy_start')},${cutoff}$63,0))")
    forecast(cash,56,"당월 인식 국세",lambda i,t,c,p:f"{c}34-{c}55" + (f"-{c}37" if has_rural else ""))
    row(cash,57,"국세 미지급 순잔액",lambda i,t,c,p:
        f"{c}40-{c}42" + (f"-{c}68" if has_rural else "") if i==last else (
            f"{p}57+{c}56-SUM({c}48,{c}51:{c}52,{c}54)+{c}59" if i>last else None),"last")
    row(cash,58,"지방세 미지급 순잔액",lambda i,t,c,p:
        f"{c}41-{c}43" if i==last else (f"{p}58+{c}55-{c}53+{c}60" if i>last else None),"last")
    def closing_payable(i,t,c,p):
        if i == last:
            return f"{c}40+{c}41"
        if i > last:
            return f"MAX(0,{c}57)+MAX(0,{c}58)" + (f"+{c}68" if has_rural else "")
    row(cash,31,"기말 미지급법인세",closing_payable,"last")
    row(cash,32,"기말 선납·환급세액",lambda i,t,c,p:
        f"{c}42+{c}43" if i==last else (f"MAX(0,-{c}57)+MAX(0,-{c}58)" if i>last else None),"last")
    def balance_check(i,t,c,p):
        if i <= last:
            return None
        rural = f",-'CF'!{c}68" if has_rural else ""
        return (f"MAX(ABS('CF'!{p}31-'CF'!{p}32+'IS'!{c}{isr['tax']}-'CF'!{c}33"
                f"-('CF'!{c}31-'CF'!{c}32)),-'CF'!{c}31,-'CF'!{c}32{rural})")
    check(140,"미지급법인세·선납세액 연결 및 비음수",balance_check)
    m.total(cash,31)
    m.total(cash,33)
