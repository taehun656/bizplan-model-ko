"""Korean corporate credit vintages. All workbook amounts are in display units."""
from tax_policy import credit_inputs, INTERIM_PERIOD_MONTHS


CREDIT_LABELS = {
    "minimum_tax": "최저한세액",
    "credit_limit": "최저한세 대상 세액공제 사용 한도",
    "credit_restricted_used": "최저한세 대상 세액공제 사용액",
    "credit_used_change": "당월 세액공제 사용액 증감",
    "credit_close": "기말 이월세액공제",
    "rural_tax": "농어촌특별세 (누적)",
    "half_year_net": "상반기 실적 기준 공제 후 중간예납액",
}
RURAL_RATE = 0.20
SME_MIN_RATE = 0.07
FIRST_THRESHOLD = "10000000000/Unit"
SECOND_THRESHOLD = "100000000000/Unit"


def _minimum(refs, base, independent=False):
    first, second = FIRST_THRESHOLD, SECOND_THRESHOLD
    if independent:
        general = f"{base}*0.10+MAX(0,{base}-{first})*0.02+MAX(0,{base}-{second})*0.05"
    else:
        general = (f"MIN({base},{first})*0.10+MAX(0,MIN({base},{second})-{first})*0.12"
                   f"+MAX(0,{base}-{second})*0.17")
    rate = refs["min_tax_rate"]
    return (f"IF({refs['min_tax_apply']},IF(ISNUMBER({rate}),{base}*{rate},"
            f"IF({refs['is_sme']},{base}*{SME_MIN_RATE},{general})),0)")


def build_credit_ledger(m, row, check, put, period, first_row):
    lots = credit_inputs(m.spec["tax"])
    refs, r = m.tax_refs, m.tax_rows
    as_of = refs["credit_as_of"]
    count = len(lots)
    titles = {"open": "기초", "new": "발생 (누적)", "expired": "소멸", "eligible": "공제 가능액",
              "used": "사용 (누적)", "audit_used": "사용액 순서 독립 재계산", "close": "기말"}
    group = {key: first_row + 2 + n*(count+1) for n, key in enumerate(titles)}
    m.credit_ledger_rows = [{key: start+n for key, start in group.items()} for n in range(count)]
    metadata = group["close"] + count + 3
    if count:
        m.credit_metadata_bounds = (metadata, metadata+count-1)
        m.section("IS", metadata-2, "이월세액공제 · 유형 구분 및 공제 순서")
        m.table("IS", metadata-1, 10, [["발생 개시월", "공제 순서값", "최저한세 대상",
                                      "농특세 대상", "공제 기한월", "기준월 잔액"]])
        m.header("IS", metadata-1)
        m.section("IS", first_row, "이월세액공제 잔액표 · 월별 사업연도 누적")
        for key, caption in titles.items():
            m.section("IS", group[key]-1, "이월세액공제 · " + caption)

    def addr(c, key):
        return f"{c}{r[key]}"

    def area(c, key):
        return f"{c}{group[key]}:{c}{group[key]+count-1}"

    def meta(col):
        return f"${col}${metadata}:${col}${metadata+count-1}"

    def pool(c, flag):
        return f"SUMIF({meta('L')},{flag},{area(c,'eligible')})" if count else "0"

    def total(c, key):
        return f"SUM({area(c,key)})" if count else "0"

    def legacy(c):
        return f"{addr(c,'credits')}*{addr(c,'fy_month')}/{addr(c,'fy_length')}"

    for n, lot in enumerate(lots):
        loc = m.credit_ledger_rows[n]
        opening = lot["source"] == "opening"
        year, label, amount, restricted, rural = [
            refs["credit", n, field] for field in
            ("year", "label", "amount", "minimum_tax_applies", "rural_tax")]
        erow = loc["eligible"]
        mr = metadata+n
        origin, priority, applies, deadline = [f"${col}${mr}" for col in ("J","K","L","N")]
        # Keep sorting/flags outside the monthly grid; changes to assumptions
        # affect priority, expiry, rural tax and both annual/interim allocations.
        m.cell("IS", mr, 2, f'={year}&"년 · "&{label}')
        m.cell("IS", mr, 10,
               f"=IFERROR(INDEX({refs['annual_starts']},1,MATCH({year},{refs['annual_years']},0)),"
               f"DATE({year},{refs['fy_start_month']},1))")
        m.cell("IS", mr, 11, f"={origin}*{count}+{n}")
        m.cell("IS", mr, 12, f"=--{restricted}")
        m.cell("IS", mr, 13, f"=--{rural}")
        m.cell("IS", mr, 14, f"=EDATE({period(origin, True)},120)")
        m.cell("IS", mr, 15, "=" + (amount if opening else "0"))
        def opening_amount(i, t, c, p):
            fy = addr(c, "fy_start")
            seed = amount if opening else "0"
            if not i:
                return f"IF({fy}=EDATE({as_of},1),{seed},0)"
            return (f"IF({fy}<={as_of},0,IF({fy}={addr(p,'fy_start')},{p}{loc['open']},"
                    f"IF({fy}=EDATE({as_of},1),{seed},{p}{loc['close']})))")
        row("IS", loc["open"], lot["label"], opening_amount, "last")
        row("IS", loc["new"], lot["label"], lambda i,t,c,p:
            (f"IF(AND({addr(c,'fy_start')}>{as_of},{addr(c,'fy_start')}={origin}),"
             f"{amount}*{addr(c,'fy_month')}/{addr(c,'fy_length')},0)") if not opening else "0", "last")
        row("IS", loc["expired"], lot["label"], lambda i,t,c,p:
            f"IF({addr(c,'fy_end')}>{deadline},{c}{loc['open']},0)", "last")
        row("IS", erow, lot["label"], lambda i,t,c,p:
            f"{c}{loc['open']}+{c}{loc['new']}-{c}{loc['expired']}", "last")
        def used(i,t,c,p):
            higher = (f'SUMIFS({area(c,"eligible")},{meta("L")},{applies},'
                      f'{meta("K")},"<"&{priority})')
            budget = (f"IF({applies},{addr(c,'credit_limit')},"
                      f"{addr(c,'national_gross')}-{addr(c,'credit_restricted_used')})")
            return f"MIN({c}{erow},MAX(0,{budget}-{higher}))"
        row("IS", loc["used"], lot["label"], used, "last")
        def audit_used(i,t,c,p):
            # Independent FIFO identity: cumulative claim through this lot,
            # less the cumulative claim through the preceding lot.
            before = (f"SUMPRODUCT(--({meta('K')}<{priority}),--({meta('L')}={applies}),"
                      f"{area(c,'eligible')})")
            cap = (f"IF({applies},{addr(c,'credit_limit')},"
                   f"{addr(c,'national_gross')}-{addr(c,'credit_restricted_used')})")
            return f"MIN({cap},{before}+{c}{erow})-MIN({cap},{before})"
        row("IS", loc["audit_used"], lot["label"], audit_used, "last")
        row("IS", loc["close"], lot["label"], lambda i,t,c,p:
            f"IF({c}$5={as_of},{amount if opening else '0'},{c}{erow}-{c}{loc['used']})", "last")
        for data_row in loc.values():
            m.cell("IS", data_row, 2, f'={year}&"년 · "&{label}')

    put("minimum_tax", lambda i,t,c,p: _minimum(refs, addr(c,"tax_base")))
    put("credit_limit", lambda i,t,c,p: f"MAX(0,{addr(c,'national_gross')}-{addr(c,'minimum_tax')})")
    put("credit_restricted_used", lambda i,t,c,p: f"MIN({addr(c,'credit_limit')},{pool(c,1)})")
    put("credits_applied", lambda i,t,c,p:
        f"MIN({addr(c,'national_gross')},{total(c,'used')}+{legacy(c)})")
    def used_change(i,t,c,p):
        previous = f"IF({addr(c,'fy_start')}={addr(p,'fy_start')},{addr(p,'credits_applied')},0)" if i else "0"
        return f'IF({c}$5<={as_of},"",{addr(c,"credits_applied")}-{previous})'
    put("credit_used_change", used_change, "sum")
    put("credit_close", lambda i,t,c,p: f'IF({c}$5<{as_of},"",{total(c,"close")})')
    put("rural_tax", lambda i,t,c,p:
        f"SUMIF({meta('M')},1,{area(c,'used')})*{RURAL_RATE}" if count else "0")
    def interim_net(i,t,c,p):
        gross = addr(c, "half_year_gross")
        cap = f"MAX(0,{gross}-{addr(c,'minimum_tax')})"
        return (f"IF({addr(c,'fy_month')}={INTERIM_PERIOD_MONTHS},"
                f"MAX(0,{gross}-MIN({cap},{pool(c,1)})-{pool(c,0)}-{legacy(c)}),0)")
    put("half_year_net", interim_net, "sum")
    _credit_checks(m, check, group, count, metadata)
    if count:
        _fiscal_table(m, metadata + count + 1, lots)


def _credit_checks(m, check, group, count, metadata):
    refs, r = m.tax_refs, m.tax_rows
    def data(c, key):
        return f"'IS'!{c}{r[key]}"
    def area(c, key):
        return f"'IS'!{c}{group[key]}:{c}{group[key]+count-1}"
    def meta(col):
        return f"'IS'!${col}${metadata}:${col}${metadata+count-1}"

    opening = f"SUM({meta('O')})" if count else "0"
    def rollforward(i,t,c,p):
        if not count:
            return "0"
        movement = (f"SUMPRODUCT(ABS({area(c,'open')}+{area(c,'new')}-{area(c,'used')}"
                    f"-{area(c,'expired')}-{area(c,'close')}))"
                    f"+ABS({data(c,'credit_close')}-SUM({area(c,'close')}))")
        return (f"IF({c}$5<{refs['credit_as_of']},0,IF({c}$5={refs['credit_as_of']},"
                f"ABS({data(c,'credit_close')}-({opening})),{movement}))")
    check(141, "발생연도·유형별 세액공제 잔액표 연결", rollforward)

    def tax_difference(i,t,c,p):
        base, gross = data(c,"tax_base"), data(c,"national_gross")
        minimum, cap = data(c,"minimum_tax"), data(c,"credit_limit")
        restricted, used = data(c,"credit_restricted_used"), data(c,"credits_applied")
        rural = data(c,"rural_tax")
        subject, exempt, special, fifo_error = ("0",)*4
        if count:
            eligible = area(c,"eligible")
            subject = f"SUMPRODUCT({meta('L')},{eligible})"
            exempt = f"SUM({eligible})-({subject})"
            special = f"SUMPRODUCT({meta('M')},{area(c,'audit_used')})*{RURAL_RATE}"
            fifo_error = f"SUMPRODUCT(ABS({area(c,'used')}-{area(c,'audit_used')}))"
        legacy = f"{data(c,'credits')}*{data(c,'fy_month')}/{data(c,'fy_length')}"
        return (f"ABS({minimum}-({_minimum(refs,base,independent=True)}))"
                f"+ABS({cap}-MAX(0,{gross}-{minimum}))+ABS({restricted}-MIN({cap},{subject}))"
                f"+MAX(0,MIN({minimum},{gross})-({gross}-{restricted}))"
                f"+ABS({used}-MIN({gross},{restricted}+({exempt})+{legacy}))"
                f"+ABS({data(c,'national_net')}-({gross}-{used}))+ABS({rural}-({special}))+{fifo_error}"
                f"+ABS({data(c,'tax_ytd')}-({gross}-{used}+{data(c,'local_tax')}+{rural}))")
    check(142, "최저한세·공제순서·농어촌특별세 독립 재계산", tax_difference)


def _fiscal_table(m, header, lots):
    """Columns are legal FYs, including stubs; a partial final period is dated."""
    sh, r = "IS", m.tax_rows
    end = m.col(len(m.months)-1)
    month_axis = f"$J$5:${end}$5"
    fy_end_axis = f"$J${r['fy_end']}:${end}${r['fy_end']}"
    m.section(sh, header, "이월세액공제 잔액표 · 사업연도별")
    for offset, label in enumerate(("사업연도 개시월", "사업연도 종료월", "집계 기준월"), 1):
        m.cell(sh, header+offset, 2, label, "text")
        m.format_row(sh, header+offset, "yyyy-mm")
    m.header(sh, header+3)
    movements = {"open": "기초", "new": "발생", "used": "사용", "expired": "소멸", "close": "기말"}
    for n, lot in enumerate(lots):
        for index, caption in enumerate(movements.values()):
            dest = header+4+n*len(movements)+index
            year, name = (m.tax_refs["credit",n,key] for key in ("year","label"))
            m.cell(sh, dest, 2, f'={year}&"년 · "&{name}&" · {caption}"')
    # Reserve enough columns for an editable fiscal month and legal stubs.
    periods = len({month[:4] for month in m.months})+1+2*len(m.spec["tax"]["short_fys"])
    for idx in range(min(len(m.months), periods)):
        c, number = m.col(idx), 10+idx
        preceding = m.tax_refs["credit_as_of"] if not idx else f"{m.col(idx-1)}{header+2}"
        start, stop, at = (f"{c}{header+k}" for k in (1,2,3))
        opening = f"EDATE({preceding},1)"
        if not idx:
            opening = f"MAX({opening},$J${r['fy_start']})"
        m.cell(sh, header+1, number,
               f'=IF(AND(ISNUMBER({preceding}),{preceding}<${end}$5),{opening},"")')
        m.cell(sh, header+2, number,
               f'=IF(ISNUMBER({start}),INDEX({fy_end_axis},1,MATCH(MAX({start},$J$5),{month_axis},0)),"")')
        m.cell(sh, header+3, number, f'=IF(ISNUMBER({stop}),MIN({stop},${end}$5),"")')
        for n, loc in enumerate(m.credit_ledger_rows):
            for index, key in enumerate(movements):
                dest = header+4+n*len(movements)+index
                data = f"$J${loc[key]}:${end}${loc[key]}"
                m.cell(sh, dest, number,
                       f'=IF(ISNUMBER({at}),INDEX({data},1,MATCH({at},{month_axis},0)),"")')
