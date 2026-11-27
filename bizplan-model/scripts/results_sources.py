"""Formula-backed chart tables for the results worksheet."""
from openpyxl.utils import get_column_letter as letter

from assumption_layout import DISPLAY_UNIT
from dashboard import LIVE, RESULT


def _cell(m, row, col, expression):
    m.cell(RESULT, row, col, "=" + expression)


def _caption(m, row, col, text):
    _cell(m, row, col, '"' + text.replace('"', '""') + '"')


def _title(m, row, caption, col=2):
    _cell(m, row, col, f'"{caption} ("&{DISPLAY_UNIT[1:]}&")"')
    return f"{letter(col)}{row}"


def _table(header, count, category, columns, title):
    return dict(header=header, count=count, category=category,
                values=list(columns), title_cell=title)


def _monthly_tables(m, header, annual_rows):
    months = len(m.months)
    min_cash = next((sheet, addr) for _, label, sheet, addr, _ in m.inputs
                    if sheet == "전제사항" and label == "최소 보유현금 (표시 단위)")
    _caption(m, header, 11, "최소 보유현금")
    for col, key in enumerate(("revenue", "ebitda", "op"), 12):
        _cell(m, header, col, f"B{annual_rows[key]}")
    for i in range(months):
        row = header + i + 1
        _cell(m, row, 11, f"'{min_cash[0]}'!{min_cash[1]}")
        for col, key in enumerate(("revenue", "ebitda", "op"), 12):
            _cell(m, row, col, f"'IS'!{m.col(i)}{m.is_rows[key]}")
    title_row = header + months + 2
    monthly = _table(header, months, "F", "LMN",
                     _title(m, title_row, "월별 매출·EBITDA·영업이익"))
    cash = _table(header, months, "F", "GHK", _title(m, title_row + 1, "월별 기말 현금·차입금"))
    return monthly, cash, title_row + 3


def _annual_table(m, header, annual_rows):
    years = list(m.years.values())
    lines = len(m.spec["revenue_lines"])
    series = min(lines, 5)
    grouped = lines > series
    title = _title(m, header - 1, "연도별 매출")
    _caption(m, header, 2, "연도")
    total = 3 + series
    _caption(m, header, total, "합계")
    # Ranking all years by the plan-year size keeps the colors and legend
    # consistent when the historical mix changes. Ties retain spec order.
    raw = header + len(years) + 3
    plan_col = letter(3 + len(years))
    first, last = raw + 1, raw + lines
    ranks = f"C{first}:C{last}"
    _caption(m, raw, 2, "매출 라인")
    _caption(m, raw, 3, "계획 연도 순위")
    for n, indices in enumerate(years):
        year_col = m.col(len(m.months) + n)
        _cell(m, raw, n + 4, f"'INDEX'!{year_col}5")
        _cell(m, header + n + 1, 2, f"'INDEX'!{year_col}5")
        table_col = letter(3 + 2 * n)
        _cell(m, header + n + 1, total, f"{table_col}{annual_rows['revenue']}")
    for i in range(lines):
        row, is_row = first + i, 9 + i
        label_col = letter(m.levels["IS"][is_row])
        _cell(m, row, 2, f"'IS'!{label_col}{is_row}")
        for j in range(len(years)):
            year_col = m.col(len(m.months) + j)
            _cell(m, row, j + 4, f"'IS'!{year_col}{is_row}")
        _cell(m, row, 3, f"RANK({plan_col}{row},{plan_col}{first}:{plan_col}{last})"
              f"+COUNTIF({plan_col}{first}:{plan_col}{row},{plan_col}{row})-1")

    for i in range(series):
        col = 3 + i
        other = grouped and i == series - 1
        index = f"MATCH({i+1},{ranks},0)"
        if other:
            _caption(m, header, col, "기타")
        else:
            _cell(m, header, col, f"INDEX(B{first}:B{last},{index})")
        for j in range(len(years)):
            year_col, row = letter(j + 4), header + j + 1
            value = (f"{letter(total)}{row}-SUM(C{row}:{letter(col-1)}{row})" if other
                     else f"INDEX({year_col}{first}:{year_col}{last},{index})")
            _cell(m, row, col, value)
    m.header(RESULT, header)
    m.header(RESULT, raw)
    return _table(header, len(years), "B", [letter(c) for c in range(3, total + 1)], title), last + 3


def _bridge_table(m, header):
    years = list(m.years.items())
    if len(years) < 2:
        return None, header
    (prior_year, prior), (plan_year, plan) = years[-2:]
    prior_month = {m.months[i][5:]: i for i in prior}
    plan_month = {m.months[i][5:]: i for i in plan}
    overlap = sorted(prior_month.keys() & plan_month.keys())
    if not overlap:
        return None, header
    partial = len(overlap) < 12
    span = f"{int(overlap[0])}월" if len(overlap) == 1 else f"{int(overlap[0])}~{int(overlap[-1])}월"
    caption = f"매출 브리지 · 동월 {span}" if partial else "매출 브리지"
    title = _title(m, header - 1, caption)
    for c, name in enumerate(("항목", "받침", "매출", "증가", "감소", "누계", "증감",
                              "증감 표시 위치", "증감 표시"), 2):
        _caption(m, header, c, name)

    def amount(months, source_row):
        start = m.col(months[overlap[0]])
        end = m.col(months[overlap[-1]])
        return f"SUM('IS'!{start}{source_row}:{end}{source_row})"

    first = header + 1
    lines = len(m.spec["revenue_lines"])
    top = 5
    grouped = lines > top + 1
    count = (top + 1 if grouped else lines) + 2
    last = header + count
    next_header = last + 3
    if grouped:
        # Rank signed changes by magnitude. The running count breaks ties in
        # the original IS order, including zeros and equal opposite changes.
        raw = next_header
        detail_first, detail_last = raw + 1, raw + lines
        labels = f"B{detail_first}:B{detail_last}"
        changes = f"C{detail_first}:C{detail_last}"
        magnitudes = f"D{detail_first}:D{detail_last}"
        ranks = f"E{detail_first}:E{detail_last}"
        for col, name in enumerate(("매출 라인", "증감", "증감 절대값", "순위"), 2):
            _caption(m, raw, col, name)
        for i in range(lines):
            row, is_row = detail_first + i, 9 + i
            _cell(m, row, 2, f"'IS'!{letter(m.levels['IS'][is_row])}{is_row}")
            _cell(m, row, 3, f"{amount(plan_month, is_row)}-{amount(prior_month, is_row)}")
            _cell(m, row, 4, f"ABS(C{row})")
            _cell(m, row, 5, f"RANK(D{row},{magnitudes})+COUNTIF(D{detail_first}:D{row},D{row})-1")
        m.header(RESULT, raw)
        next_header = detail_last + 3
        _caption(m, header, 11, "차트 축 이름")
    for row, year, months in ((first, prior_year, prior_month), (last, plan_year, plan_month)):
        _caption(m, row, 2, year)
        _cell(m, row, 4, amount(months, m.is_rows["revenue"]))
        _cell(m, row, 7, f"D{row}")
        for col in (3, 5, 6, 8):
            _cell(m, row, col, "0")
    for i, row in enumerate(range(first + 1, last)):
        if grouped and i == top:
            _caption(m, row, 2, "기타")
            _cell(m, row, 8, f'SUMIF({ranks},">{top}",{changes})')
        elif grouped:
            index = f"MATCH({i+1},{ranks},0)"
            _cell(m, row, 2, f"INDEX({labels},{index})")
            _cell(m, row, 8, f"INDEX({changes},{index})")
        else:
            is_row = 9 + i
            label_col = letter(m.levels["IS"][is_row])
            _cell(m, row, 2, f"'IS'!{label_col}{is_row}")
            _cell(m, row, 8, f"{amount(plan_month, is_row)}-{amount(prior_month, is_row)}")
        _cell(m, row, 7, f"G{row-1}+H{row}")
        _cell(m, row, 3, f"MIN(G{row-1},G{row})")
        _cell(m, row, 4, "0")
        _cell(m, row, 5, f"MAX(H{row},0)")
        _cell(m, row, 6, f"MAX(-H{row},0)")
    for row in range(first, last + 1):
        _cell(m, row, 9, f"SUM(C{row}:F{row})")
        if grouped:
            # The feed retains the full label in B. Cap only the narrow
            # dashboard axis, keeping both the prefix and the distinguisher.
            _cell(m, row, 11, f'IF(LEN(B{row})>13,LEFT(B{row},6)&"…"&RIGHT(B{row},6),B{row})')
        if row in (first, last):
            _caption(m, row, 10, "")
        else:
            _cell(m, row, 10, f'TEXT(H{row},"+#,##0;(#,##0);0")')
    m.header(RESULT, header)
    bridge = _table(header, count, "K" if grouped else "B", "CDEF", title)
    bridge.update(label_height="I", label_text="J")
    if grouped:
        bridge["grouped"] = True
    return bridge, next_header


def _scenario_table(m, header):
    count = len(m.cases)
    dscr = next(i for i, (_, key) in enumerate(LIVE, 1) if key == "dscr")
    op = next(i for i, (_, key) in enumerate(LIVE, 1) if key == "op")
    title_row = header - 1
    mode = f"I{title_row}"
    _cell(m, title_row, 9, f"IF(ISBLANK(SN_LastRun),ISNUMBER(INDEX(SN_Live,{dscr},1)),"
          f"COUNT(INDEX(SN_Store,{dscr},0))>0)")
    _cell(m, title_row, 2, f'IF({mode},"시나리오별 DSCR (배) · 기준 1.0",'
          f'"시나리오별 영업이익 ("&{DISPLAY_UNIT[1:]}&")")'
          '&IF(ISBLANK(SN_LastRun)," · 민감도 갱신 필요","")')
    for col, name in enumerate(("케이스", "기타 케이스", "선택 케이스", "기준 1.0배",
                                "일반 값 표시", "선택 값 표시", "저장 지표"), 2):
        _caption(m, header, col, name)
    _cell(m, header, 5, f'IF({mode},"기준 1.0배","기준선 미적용")')
    for i in range(1, count + 1):
        row, case_col = header + i, letter(4 + i)
        source = f"INDEX(SN_Store,IF({mode},{dscr},{op}),{i})"
        _cell(m, row, 2, f"'Sensitivity'!{case_col}14")
        _cell(m, row, 8, f'IF(OR(ISBLANK(SN_LastRun),NOT(ISNUMBER({source}))),"",{source})')
        for col, label_col, test in ((3, 6, f"SN_Case<>{i}"), (4, 7, f"SN_Case={i}")):
            active = f"AND({test},ISNUMBER(H{row}))"
            _cell(m, row, col, f"IF({active},H{row},0)")
            _cell(m, row, label_col, f'IF({active},TEXT(H{row},IF({mode},"0.00","#,##0")),"")')
        _cell(m, row, 5, f'IF(AND({mode},NOT(ISBLANK(SN_LastRun))),1,"")')
    m.header(RESULT, header)
    scenario = _table(header, count, "B", "CDE", f"B{title_row}")
    scenario.update(labels=["F", "G"], mode_cell=mode)
    return scenario


def build_sources(m, annual_rows, source_header):
    monthly, cash, next_header = _monthly_tables(m, source_header, annual_rows)
    annual, next_header = _annual_table(m, next_header + 1, annual_rows)
    bridge, next_header = _bridge_table(m, next_header)
    scenario = _scenario_table(m, next_header)
    m.chart_sources = dict(annual=annual, bridge=bridge, scenario=scenario, monthly=monthly, cash=cash)

