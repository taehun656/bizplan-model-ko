"""Excel controls for edits to the fiscal snapshot and credit lots."""

from tax_policy import FIRST_SUPPORTED_FY


def register_credit_input_checks(m, snapshot_row, first_lot, lots):
    sheet = "전제사항"
    ref = m.tax_refs
    as_of = ref["credit_as_of"]
    next_start = f"EDATE({as_of},1)"
    month = ref["fy_start_month"]
    founded = ref["founded_fy"]
    first_data = "'INDEX'!$J$5"
    last_data = f"'INDEX'!${m.col(len(m.months)-1)}$5"
    regular = [f"MONTH({next_start})={month}"]
    for i, _ in enumerate(m.spec["tax"].get("short_fys", {})):
        regular.append(f"OR({next_start}<{ref['short',i,'start']},{next_start}>{ref['short',i,'end']})")
    valid_start = f"OR(COUNTIF({ref['annual_starts']},{next_start})>0,AND({','.join(regular)}))"
    rate = ref["min_tax_rate"]
    switch = ref["min_tax_apply"]
    conditions = [f"ISNUMBER({as_of})", f"DAY({as_of})=1", f"{as_of}<=LastActual",
                  f"OR(COUNT({founded})=0,{next_start}>={founded})", valid_start,
                  f"ISNUMBER({switch})", f"OR({switch}=0,{switch}=1)",
                  f'OR({rate}="auto",AND(ISNUMBER({rate}),{rate}>=0,{rate}<=1))']
    if any(lot["source"] == "opening" for lot in lots):
        year, number = FIRST_SUPPORTED_FY.split("-")
        conditions.extend((f"{next_start}>={first_data}", f"{next_start}>=DATE({year},{int(number)},1)"))
    m.cell(sheet, snapshot_row, 16, f'=IFERROR(IF(AND({",".join(conditions)}),"OK","NO"),"NO")')
    m.add_check("세액공제 기준 월·최저한세 입력 유효성", f"='{sheet}'!$P${snapshot_row}")
    if not lots:
        return
    first_metadata = m.credit_metadata_bounds[0]
    forecast_fy = f"'IS'!${m.col(m.last+1)}${m.tax_rows['fy_start']}"
    for i, lot in enumerate(lots):
        row = first_lot + i
        origin = f"'IS'!$J${first_metadata+i}"
        expiry = f"'IS'!$N${first_metadata+i}"
        year = ref["credit", i, "year"]
        amount = ref["credit", i, "amount"]
        flags = [ref["credit", i, key] for key in ("minimum_tax_applies", "rural_tax")]
        label = ref["credit", i, "label"]
        conditions = [f"ISNUMBER({year})", f"INT({year})={year}", f"{year}>=1900", f"{year}<=9999",
                      f"ISTEXT({label})", f"LEN(TRIM({label}))>0", f"ISNUMBER({amount})", f"{amount}>=0",
                      f"COUNTIF({ref['annual_years']},{year})<=1",
                      f"OR(COUNT({founded})=0,{origin}>={founded})"]
        conditions.extend(f"AND(ISNUMBER({flag}),OR({flag}=0,{flag}=1))" for flag in flags)
        if lot["source"] == "opening":
            conditions.extend((f"EDATE({expiry},-120)<={as_of}", f"OR({amount}=0,{expiry}>={as_of})"))
        else:
            conditions.extend((f"{origin}>{as_of}", f"{origin}>=MAX({forecast_fy},{first_data})",
                               f"{origin}<={last_data}"))
        m.cell(sheet, row, 16, f'=IFERROR(IF(AND({",".join(conditions)}),"OK","NO"),"NO")')
    last_lot = first_lot + len(lots) - 1
    m.add_check("세액공제 발생연도·유형별 입력 유효성",
                f'=IF(COUNTIF(\'{sheet}\'!$P${first_lot}:$P${last_lot},"OK")={len(lots)},"OK","NO")')
