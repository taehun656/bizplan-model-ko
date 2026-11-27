"""Workbook formulas for Korean current income tax (references/tax_rules_kr.md)."""


from tax_policy import (credit_mode, INTERIM_PERIOD_MONTHS, LOSS_LIFE_OLD, LOSS_LIFE_NEW, NEW_LOSS_LIFE_START,
                        SMALL_FILING_START, NEW_RATES_START, FIRST_SUPPORTED_FY, YEAR_MONTHS)
from tax_credit_ledger import CREDIT_LABELS, build_credit_ledger


def build_tax_accrual(m, row, check, isr, base):
    """Write statutory YTD on IS, preserving historical booked tax and monthly reversals."""
    sh = "IS"
    ref = m.tax_refs
    tax = m.spec["tax"]
    last = m.last
    old_era = int(FIRST_SUPPORTED_FY[:4])
    new_era = int(NEW_RATES_START[:4])
    small_start = int(SMALL_FILING_START[:4])
    new_loss_start = int(NEW_LOSS_LIFE_START[:4])
    labels = {
        "fy_start": "사업연도 개시월", "fy_end": "사업연도 종료월",
        "fy_month": "사업연도 경과 월수", "fy_length": "사업연도 전체 월수",
        "pretax_ytd": "사업연도 누적 세전이익", "actual_ytd_tax": "사업연도 누적 실적 법인세",
        "additions": "연간 익금산입·손금불산입", "deductions": "연간 손금산입·익금불산입",
        "non_taxable": "연간 비과세소득", "income_deduction": "연간 소득공제",
        "credits": "연간 세액공제·감면", "withholding": "연간 원천징수세액",
        "ad_hoc": "연간 수시부과세액", "interim_paid": "실적 중간예납액",
        "installment_amount": "분납 지정 금액", "income_ytd": "① 각 사업연도 소득 (누적)",
        "nol_open": "당월 기초 이월결손금", "nol_expired": "당월 만료 결손금",
        "nol_deduct_ytd": "② 누적 이월결손금 공제", "nol_used": "당월 결손금 공제 증감",
        "nol_generated": "당월 발생 결손금 (사업연도 말)", "nol_close": "기말 이월결손금",
        "deduction_limit": "사업연도 소득 기준 결손금 공제 한도",
        "tax_base": "② 누적 과세표준", "annualized_base": "12개월 환산 과세표준",
        "national_gross": "③ 법인세 산출세액", "credits_applied": "④ 누적 세액공제·감면 적용액",
        "national_net": "④ 감면·공제 후 법인세", "national_due_ytd": "④ 기납부 차감 후 국세 납부할 세액",
        "local_tax": "⑤ 법인지방소득세",
        "tax_ytd": "누적 당기법인세 부담액", "tax_base_change": "당월 과세표준 증감",
        "recognized_ytd": "사업연도 누적 인식 법인세", "fy_sequence": "추정 사업연도 순번",
        "half_year_gross": "상반기 실적 기준 환산 중간예납 산출세액",
    }
    tracked = credit_mode(tax)
    if tracked:
        labels.update(CREDIT_LABELS)
        labels["credits_applied"] = "④ 세액공제 사용액 (누적)"
        labels["national_net"] = "④ 결정세액 (국세)"
    r = m.tax_rows = {key: base+n for n,key in enumerate(labels)}
    r["tax_expense"] = isr["tax"]

    def put(key, fn, annual="last"):
        row(sh, r[key], labels[key], fn, annual)

    def at(c, key):
        return f"{c}{r[key]}"

    first_forecast = at("$" + m.col(last+1), "fy_start")
    def same(c, p):
        return f"{at(c,'fy_start')}={at(p,'fy_start')}"
    def carry(i, c, p, key):
        return f"IF({same(c,p)},{at(p,key)},0)" if i else "0"

    def period(month, end):
        start_month = ref["fy_start_month"]
        regular = f"DATE(YEAR({month})-(MONTH({month})<{start_month}),{start_month},1)"
        stubs = [(ref["short",n,"start"], ref["short",n,"end"])
                 for n in range(len(tax.get("short_fys", {})))]
        # Truncate adjacent normal FYs at a stub boundary. MIN/MAX use only
        # qualifying dates; wrapping the natural date in each term keeps this linear.
        natural = f"EDATE({regular},11)" if end else regular
        candidates = [natural]
        for start, stop in stubs:
            candidates.append(f"IF({month}<{start},EDATE({start},-1),{natural})" if end
                              else f"IF({month}>{stop},EDATE({stop},1),{natural})")
        expr = (("MIN" if end else "MAX")+"("+",".join(candidates)+")") if stubs else natural
        if not end:
            founded = ref["founded_fy"]
            expr = f"MAX({expr},IF(AND(ISNUMBER({founded}),{month}>={founded}),{founded},0))"
        for start, stop in reversed(stubs):
            expr = f"IF(AND({month}>={start},{month}<={stop}),{stop if end else start},{expr})"
        return expr
    put("fy_start", lambda i,t,c,p:period(f"{c}$5", False))
    put("fy_end", lambda i,t,c,p:period(f"{c}$5", True))
    for key in ("fy_start","fy_end"):
        m.format_row(sh,r[key],"yyyy-mm")
    def months_between(a,b):
        return f"(YEAR({b})-YEAR({a}))*12+MONTH({b})-MONTH({a})+1"
    put("fy_length",lambda i,t,c,p:months_between(at(c,"fy_start"),at(c,"fy_end")))
    put("fy_month",lambda i,t,c,p:months_between(at(c,"fy_start"),f"{c}$5"))
    put("fy_sequence",lambda i,t,c,p:
        f"IF({at(c,'fy_start')}<{first_forecast},0,{at(p,'fy_sequence')}+--({at(c,'fy_start')}<>{at(p,'fy_start')}))"
        if i else f"--({at(c,'fy_start')}>={first_forecast})")
    put("pretax_ytd",lambda i,t,c,p:f"{carry(i,c,p,'pretax_ytd')}+{c}{isr['pretax']}")
    put("actual_ytd_tax",lambda i,t,c,p:carry(i,c,p,'actual_ytd_tax')+
        (f"+'실적_IS'!{c}13" if i<=last else ""))
    for key in ("additions","deductions","non_taxable","income_deduction","credits","withholding",
                "ad_hoc","interim_paid","installment_amount"):
        put(key,lambda i,t,c,p:f"SUMIF({ref['annual_starts']},{at(c,'fy_start')},{ref['annual',key]})")
    def fraction(c):
        return f"{at(c,'fy_month')}/{at(c,'fy_length')}"
    put("income_ytd",lambda i,t,c,p:f"{at(c,'pretax_ytd')}+({at(c,'additions')}-{at(c,'deductions')})*{fraction(c)}")
    put("deduction_limit",lambda i,t,c,p:f"MAX(0,{at(c,'income_ytd')})*{ref['loss_limit']}")
    put("tax_base",lambda i,t,c,p:f"MAX(0,{at(c,'income_ytd')}-{at(c,'nol_deduct_ytd')}"
        f"-({at(c,'non_taxable')}+{at(c,'income_deduction')})*{fraction(c)})")
    put("tax_base_change",lambda i,t,c,p:f"{at(c,'tax_base')}-{carry(i,c,p,'tax_base')}","sum")
    put("annualized_base",lambda i,t,c,p:f"{at(c,'tax_base')}*12/{at(c,'fy_length')}")

    # Materialize the selected marginal rates and thresholds so every bracket
    # and both tax authorities remain auditable, including user overrides.
    band = base+len(labels)+2
    m.section(sh,band-1,"적용 구간 및 누진세액")
    for n in range(3):
        row(sh,band+n,f"과세표준 {n+1}구간 상한",lambda i,t,c,p:
            f"IF(YEAR({at(c,'fy_start')}) >= {new_era},{ref['threshold',new_era,n]},{ref['threshold',old_era,n]})","last")
    for j,source in enumerate(("national","local")):
        for n in range(4):
            rate_row = band+3+4*j+n
            def selected_rate(i,t,c,p):
                year = f"YEAR({at(c,'fy_start')})"
                small = f"AND({ref['corporation_type']}=1,{year}>={small_start})"
                def regime(era):
                    return f"IF({small},{ref['rate',source,era,'small',n]},{ref['rate',source,era,'general',n]})"
                return f"IF({year}>={new_era},{regime(new_era)},{regime(old_era)})"
            row(sh,rate_row,("국세" if j==0 else "지방세")+f" {n+1}구간 적용세율",selected_rate,"last")
            m.format_row(sh,rate_row,"0.0%")
            amount_row=band+11+j*4+n
            def bracket(i,t,c,p):
                value=at(c,"annualized_base")
                floor=f"{c}{band+n-1}" if n else "0"
                capped=f"MIN({value},{c}{band+n})" if n<3 else value
                return f"MAX(0,{capped}-{floor})*{c}{rate_row}*{at(c,'fy_length')}/12"
            row(sh,amount_row,("국세" if j==0 else "지방세")+f" {n+1}구간 누적세액",bracket,"last")
    put("national_gross",lambda i,t,c,p:f"SUM({c}{band+11}:{c}{band+14})")
    def half_year_gross(i,t,c,p):
        # Use the six-month assessment period independently of the legal FY length.
        value = f"({at(c,'tax_base')}*{YEAR_MONTHS}/{INTERIM_PERIOD_MONTHS})"
        terms = [f"{value}*{c}{band+3}"]
        for k in range(3):
            terms.append(f"MAX(0,{value}-{c}{band+k})*({c}{band+4+k}-{c}{band+3+k})")
        return (f"IF({at(c,'fy_month')}={INTERIM_PERIOD_MONTHS},"
                f"({'+'.join(terms)})*{INTERIM_PERIOD_MONTHS}/{YEAR_MONTHS},0)")
    put("half_year_gross",half_year_gross,"sum")
    if not tracked:
        put("credits_applied",lambda i,t,c,p:f"MIN({at(c,'national_gross')},{at(c,'credits')}*{fraction(c)})")
    put("national_net",lambda i,t,c,p:f"{at(c,'national_gross')}-{at(c,'credits_applied')}")
    def national_due_ytd(i,t,c,p):
        # The income statement displays statutory settlement separately from
        # tax expense. Advance payments are attributed to their originating FY.
        owner = at(c,"fy_start")
        advances = f"SUMIF('CF'!$J$46:{c}$46,{owner},'CF'!$J$48:{c}$48)"
        other_paid = f"({at(c,'withholding')}+{at(c,'ad_hoc')})"
        # Through-cutoff remittances have no historical month-by-month detail.
        # Projected undated withholding is paid on the FY closing month (CF54).
        credit = (f"IF({c}$5>=LastActual,{at(c,'interim_paid')},0)+"
                  f"IF(OR(AND({owner}<=LastActual,{c}$5>=LastActual),{c}$5={at(c,'fy_end')}),{other_paid},0)")
        cut = m.col(last)
        mixed = (f"AND({c}$5>=LastActual,{owner}={at(cut,'fy_start')},"
                 f"LastActual<{at(cut,'fy_end')})")
        return f"MAX(0,{at(c,'national_net')}-{advances}-IF({mixed},'CF'!${cut}$61,{credit}))"
    put("national_due_ytd",national_due_ytd)
    put("local_tax",lambda i,t,c,p:f"IF({ref['include_local']},SUM({c}{band+15}:{c}{band+18}),0)")
    put("tax_ytd",lambda i,t,c,p:f"{at(c,'national_net')}+{at(c,'local_tax')}"+
        (f"+{at(c,'rural_tax')}" if tracked else ""))
    row(sh,isr["tax"],"법인세",lambda i,t,c,p:f"'실적_IS'!{c}13" if i<=last else
        f"{at(c,'tax_ytd')}-{carry(i,c,p,'recognized_ytd')}")
    put("recognized_ytd",lambda i,t,c,p:f"{carry(i,c,p,'recognized_ytd')}+{c}{isr['tax']}")

    lots = getattr(m, "tax_vintage_inputs", tax.get("loss_vintages", []))
    seeds = len(lots)
    # Allow for any editable opening month plus the extra boundaries introduced by a stub.
    slots = len(set(t[:4] for t in m.months[last+1:]))+1+2*len(tax.get("short_fys",{}))
    count = seeds+slots
    start = band+21
    group = {k:start+j*(count+1)+1 for j,k in enumerate(
        ("date","priority","open","expired","eligible","used","new","close"))}
    titles=("결손금 발생 사업연도","공제 순서 (발생일 기준)","사업연도 기초 잔액","사업연도 만료액",
            "사업연도 공제 가능 잔액","사업연도 누적 공제액","당월 발생액","당월 기말 잔액")
    for pos,title in zip(group.values(),titles):
        m.section(sh,pos-1,title)
    def area(c,key):
        return f"{c}{group[key]}:{c}{group[key]+count-1}"
    for n in range(count):
        seed = n<seeds
        lot = lots[n]["fy_start"] if seed else f"추정 사업연도 {n-seeds+1}"
        label = "발생연도 미상" if lot=="unknown" else str(lot)
        loc = {k:v+n for k,v in group.items()}
        def cell(c,key):return f"{c}{loc[key]}"
        def date_value(i,t,c,p):
            if seed:
                return ref["vintage",n,"year"]
            return f"IF({at(c,'fy_sequence')}={n-seeds+1},{at(c,'fy_start')},{cell(p,'date') if i else '0'})"
        row(sh,loc["date"],label,date_value,"last")
        m.format_row(sh,loc["date"],'yyyy-mm;;"-"')
        row(sh,loc["priority"],label,lambda i,t,c,p:
            f"IF({cell(c,'date')}=0,{first_forecast},{cell(c,'date')})*{count}+{n}","last")
        def fy_open(i,t,c,p):
            init=ref["vintage",n,"amount"] if seed else "0"
            if not i:return f"IF({at(c,'fy_sequence')}=1,{init},0)"
            return (f"IF({at(c,'fy_sequence')}=0,0,IF({same(c,p)},{cell(p,'open')},"
                    f"IF({at(c,'fy_sequence')}=1,{init},{cell(p,'close')})))")
        row(sh,loc["open"],label,fy_open,"last")
        row(sh,loc["expired"],label,lambda i,t,c,p:
            f"IF(AND({cell(c,'date')}>0,{at(c,'fy_start')}>EDATE({cell(c,'date')},"
            f"{YEAR_MONTHS}*IF(YEAR({cell(c,'date')})>={new_loss_start},{LOSS_LIFE_NEW},{LOSS_LIFE_OLD}))),{cell(c,'open')},0)","last")
        row(sh,loc["eligible"],label,lambda i,t,c,p:f"{cell(c,'open')}-{cell(c,'expired')}","last")
        row(sh,loc["used"],label,lambda i,t,c,p:
            f'MIN({cell(c,"eligible")},MAX(0,{at(c,"deduction_limit")}-'
            f'SUMIF({area(c,"priority")},"<"&{cell(c,"priority")},{area(c,"eligible")})))',"last")
        row(sh,loc["new"],label,lambda i,t,c,p:
            f"IF(AND({at(c,'fy_sequence')}={n-seeds+1},{at(c,'fy_end')}={c}$5),MAX(0,-{at(c,'income_ytd')}),0)"
            if not seed and i>last else "0","sum")
        row(sh,loc["close"],label,lambda i,t,c,p:
            f"{cell(c,'eligible')}-{cell(c,'used')}+{cell(c,'new')}","last")
    def new_period(i,c,p):return f"NOT({same(c,p)})" if i else "TRUE"
    put("nol_open",lambda i,t,c,p:
        f"IF({new_period(i,c,p)},SUM({area(c,'open')}),{at(p,'nol_close') if i else '0'})")
    put("nol_expired",lambda i,t,c,p:f"IF({new_period(i,c,p)},SUM({area(c,'expired')}),0)","sum")
    put("nol_deduct_ytd",lambda i,t,c,p:f"SUM({area(c,'used')})")
    put("nol_used",lambda i,t,c,p:f"{at(c,'nol_deduct_ytd')}-{carry(i,c,p,'nol_deduct_ytd')}","sum")
    put("nol_generated",lambda i,t,c,p:f"SUM({area(c,'new')})","sum")
    put("nol_close",lambda i,t,c,p:f"SUM({area(c,'close')})")
    if tracked:
        build_credit_ledger(m, row, check, put, period, group["close"] + count + 1)
    m.total(sh,r["tax_base"])
    m.total(sh,r["tax_ytd"])
    # Cross-sheet controls use a different progressive identity: first rate on
    # all income, then only increases in marginal rate on each excess.
    def recalc_difference(i,t,c,p):
        if i<=last:return None
        axis=f"'IS'!{c}"
        x=f"{axis}{r['tax_base']}*12/{axis}{r['fy_length']}"
        checks=[]
        year = f"YEAR({axis}{r['fy_start']})"
        for j,target in enumerate(("national_gross","local_tax")):
            jurisdiction = "local" if j else "national"
            def rate(k):
                def from_table(era):
                    general = ref["rate",jurisdiction,era,"general",k]
                    small = ref["rate",jurisdiction,era,"small",k]
                    return f"IF(AND({year}>={small_start},{ref['corporation_type']}=1),{small},{general})"
                return f"IF({year}>={new_era},{from_table(new_era)},{from_table(old_era)})"
            terms = [f"({x})*{rate(0)}"]
            for k in range(3):
                threshold = f"IF({year}>={new_era},{ref['threshold',new_era,k]},{ref['threshold',old_era,k]})"
                terms.append(f"MAX(0,({x})-{threshold})*({rate(k+1)}-{rate(k)})")
            calc="("+"+".join(terms)+f")*{axis}{r['fy_length']}/{YEAR_MONTHS}"
            if j:calc=f"IF({ref['include_local']},{calc},0)"
            checks.append(f"ABS({axis}{r[target]}-({calc}))")
        return "+".join(checks)
    check(136,"국세·지방세 누진표 독립 재계산",recalc_difference)
    def monthly_difference(i,t,c,p):
        if i<=last:return None
        # Independent date-filtered sum includes actual booked months and forecast refunds.
        axis=f"'IS'!{c}"
        booked=(f"SUMIF('IS'!$J$5:{c}$5,\">=\"&{axis}{r['fy_start']},"
                f"'IS'!$J{isr['tax']}:{c}{isr['tax']})")
        rural = f"+{axis}{r['rural_tax']}" if tracked else ""
        return f"ABS({booked}-({axis}{r['national_net']}+{axis}{r['local_tax']}{rural}))"
    check(137,"월별 인식 합계와 사업연도 누적 세액",monthly_difference)
    def nol_difference(i,t,c,p):
        if i<=last:return None
        axis=f"'IS'!{c}"
        main=f"ABS({axis}{r['nol_open']}+{axis}{r['nol_generated']}-{axis}{r['nol_used']}-{axis}{r['nol_expired']}-{axis}{r['nol_close']})"
        # Check the ledger column itself as well as the reported monthly movements.
        return (f"{main}+SUMPRODUCT(ABS('IS'!{area(c,'open')}+'IS'!{area(c,'new')}"
                f"-'IS'!{area(c,'expired')}-'IS'!{area(c,'used')}-'IS'!{area(c,'close')}))")
    check(138,"발생 사업연도별 이월결손금 연결",nol_difference)
