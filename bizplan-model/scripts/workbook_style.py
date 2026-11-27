"""Excel presentation for the model grid. Values and formulas belong to the builders."""
from collections import defaultdict
from math import ceil
from unicodedata import east_asian_width

from openpyxl.utils import get_column_letter as letter, range_boundaries, quote_sheetname


from design_palette import BLACK, BLUE, DARK, FONT, NUMBER_FONT, ORANGE, RED, SILVER, TAB_COLORS, WHITE, YELLOW, rgb


NUMBER = '#,##0_-;[Red](#,##0);"-"_-'
SPECIAL = {"Sensitivity", "전제사항", "민감도및결과", "스타일 가이드"}
GROUPS = ("본표", "가정·상세", "실적·원천", "검증·관리")


def _group(m, name):
    if m.roles[name][0] == "actual":
        return 2
    if name in ("검증", "기준값갱신", "스타일 가이드", "모델정보"):
        return 3
    if name in m.details.values() or name in ("인건비", "확장조정", "제거조정", "영업계획", "활동지표"):
        return 1
    return 0


def setup_toc(m):
    """Populate four sheet directories in workbook order, then the reading guide."""
    s = "목차"
    row = 4
    for group, title in enumerate(GROUPS):
        m.section(s, row, title)
        m.cell(s, row + 1, 3, "시트", "text")
        m.cell(s, row + 1, 8, "내용", "text")
        m.cell(s, row + 1, 12, "바로가기", "text")
        m.header(s, row + 1)
        row += 2
        for name in m.sheets:
            if _group(m, name) == group:
                m.cell(s, row, 3, name, "text")
                m.cell(s, row, 8, m.roles[name][2], "text")
                m.cell(s, row, 12, "열기", "text")
                row += 1
        row += 1
    m.section(s, row, "모델 이용 안내")
    notes = (
        "민감도및결과에서 계획 연도의 손익, 현금 잔액과 조달 필요액을 확인합니다.",
        "전제사항에서 표시 단위와 실적 마감 월을 확인한 뒤 노란 칸에 가정을 입력합니다.",
        "월마다 달라지는 값은 전제사항에 표시된 상세 시트의 노란 칸에서 수정합니다.",
        "Sensitivity에서 케이스 번호를 선택합니다. 가정을 수정한 뒤에는 민감도 갱신 단추를 눌러 비교표를 갱신합니다.",
        "월별 본표는 실적 마감 월까지의 실적과 이후의 추정을 보여 줍니다. 연도별 합계와 기말 잔액은 월별 열 오른쪽에 있습니다.",
        "검증에서 대사 결과를 확인합니다. 계획을 확정한 뒤 기준값 갱신 단추를 누르면 비교에 쓸 기준값이 저장됩니다.",
    )
    for offset, note in enumerate(notes, 1):
        m.cell(s, row + offset, 3, note, "text")


def _plain(value):
    return value.removeprefix("'") if isinstance(value, str) else ""


def set_number_format(band, code):
    """Preserve numeric precision and color signed financial values across all model rows."""
    if code.startswith(("#", "0.")) and "[Red]" not in code:
        sections = code.split(";")
        if len(sections) == 1:
            sections.append(f"({code})")
        sections[1] = "[Red]" + sections[1]
        code = ";".join(sections)
    # Korean Excel can require its localized color keyword even via NumberFormat.
    try:
        band.NumberFormat = code
    except Exception:
        band.NumberFormatLocal = code.replace("[Red]", "[빨강]")


def set_numeric_list_format(band):
    # Excel's list validation compares displayed text. A whole-number mask
    # can turn an invalid 0.5 into the allowed 1. Preserve the original input.
    band.NumberFormatLocal = band.Application.GetInternational(26)  # xlGeneralFormatName


def _edge(band, index, color):
    edge = band.Borders(index)
    edge.LineStyle = 1
    edge.Weight = 2
    edge.Color = rgb(color)


def _batches(ws, addresses):
    """Batch noncontiguous range formatting below Excel's address length limit."""
    batch = []
    for address in addresses:
        if len(",".join([*batch, address])) > 215 and batch:
            yield ws.Range(",".join(batch))
            batch.clear()
        batch.append(address)
    if batch:
        yield ws.Range(",".join(batch))


def _wrapped_height(text, width):
    # Widths use digit units. Hangul and wide punctuation need about two units.
    lines = sum(max(1, ceil(sum(2 if east_asian_width(c) in "FWA" else 1
                                for c in line) / max(width - 2, 1)))
                for line in text.split("\n"))
    return min(409, 16 * lines + 4)


def _labels(ws, m, name, widths):
    """Leave short labels free to spill. Wrap long labels within empty indentation cells."""
    grid = m.grid[name]
    protected = set()
    for address in m.paint[name]:
        left, top, right, bottom = range_boundaries(address)
        for col in range(max(2, left), min(7, right) + 1):
            protected.update((r, col) for r in range(top, bottom + 1))
    for row in sorted({r for r, c in grid if r > 2 and c in (2, 3, 4)}):
        col = m.levels[name].get(row)
        if col is None:
            col = next((c for c in (2, 3, 4) if (row, c) in grid), None)
        value = grid.get((row, col))
        if not isinstance(value, str) or value.startswith("="):
            continue
        # Contract rows have simultaneous B:G fields and use the tabular widths.
        if name == "금융조건" and (row, 3) in grid and (row, 4) in grid:
            continue
        end = col
        while end < 7 and (row, end + 1) not in grid and (row, end + 1) not in protected:
            end += 1
        text = _plain(value)
        height = _wrapped_height(text, sum(widths[col:end + 1]))
        if height > 20:
            label = ws.Range(f"{letter(col)}{row}:{letter(end)}{row}")
            if end > col:
                label.Merge()
            label.WrapText = True
            label.RowHeight = height


def _toc_layout(ws, m):
    grid = m.grid["목차"]
    ws.Columns("H:K").ColumnWidth = 14
    for row in sorted({r for r, c in grid if r >= 4}):
        if row in m.sections["목차"]:
            band = ws.Range(f"B{row}:L{row}")
            band.Merge()
            title = _plain(grid[row, 2])
            if title in GROUPS:
                band.Interior.Color = rgb(DARK if title == GROUPS[2] else ORANGE)
                band.Font.Color = rgb(WHITE)
            continue
        label = _plain(grid.get((row, 3)))
        if (row, 12) not in grid:
            note = ws.Range(f"C{row}:L{row}")
            note.Merge()
            note.WrapText = True
            note.RowHeight = _wrapped_height(label, 100)
            continue
        sheet_cell = ws.Range(f"C{row}:G{row}")
        sheet_cell.Merge()
        sheet_cell.WrapText = True
        desc = ws.Range(f"H{row}:K{row}")
        desc.Merge()
        desc.WrapText = True
        desc.RowHeight = max(_wrapped_height(label, 34),
                             _wrapped_height(_plain(grid[row, 8]), 54))
        if row not in m.headers["목차"]:
            ws.Hyperlinks.Add(Anchor=ws.Range(f"L{row}"), Address="",
                              SubAddress=quote_sheetname(label) + "!A1", TextToDisplay="열기")
            # Excel applies the workbook's theme hyperlink font when it adds the link.
            link = ws.Range(f"L{row}")
            link.Font.Name = FONT
            link.Font.Color = rgb(BLACK)
            link.Font.Underline = 2


def style_sheet(ws, m):
    """Format a fully assigned Excel worksheet before sheet-specific overrides."""
    name = ws.Name
    grid = m.grid[name]
    bottom = max(r for r, c in grid)
    right = letter(max(10, max(c for r, c in grid)))
    block = ws.Range(f"A1:{right}{bottom}")
    block.Font.Name = FONT
    block.Font.Size = 10
    block.Font.Color = rgb(BLACK)
    block.VerticalAlignment = -4108  # xlCenter
    set_number_format(block, NUMBER)
    block.RowHeight = 19
    widths = [0, 2, 3, 3, 3, 3, 3, 24, 15, 2]
    if name == "금융조건":
        widths[2:8] = [24, 20, 12, 14, 14, 13]
    elif name == "기준값갱신":
        widths[2:8] = [40, 40, 15, 16, 16, 16]
    elif name == "모델정보":
        widths[1:6] = [33, 38, 28, 30, 60]
    for col in range(1, 10):
        ws.Columns(f"{letter(col)}:{letter(col)}").ColumnWidth = widths[col]
    ws.Range(f"J1:{right}1").EntireColumn.ColumnWidth = 14
    group = _group(m, name)
    ws.Tab.Color = rgb(TAB_COLORS[group])
    for row, color, size, height in ((1, DARK if group == 2 else ORANGE, 10, 21),
                                     (2, SILVER, 12, 24)):
        band = ws.Range(f"A{row}:{right}{row}")
        band.Interior.Color = rgb(color)
        band.Font.Bold = True
        band.Font.Size = size
        band.Font.Color = rgb(WHITE)
        band.RowHeight = height
        if not any((row, c) in grid for c in range(2, 9)):
            title = ws.Range(f"A{row}:H{row}")
            title.Merge()
            title.ShrinkToFit = True

    monthly = (3, 10) in grid and name not in SPECIAL
    if monthly:
        last_month = m.col(len(m.months) - 1)
        summary_header = ws.Range("H5")
        summary_header.Font.Bold = True
        summary_header.Interior.Color = rgb(SILVER)
        summary_header.HorizontalAlignment = -4108
        axis = ws.Range(f"J3:{last_month}5")
        axis.Font.Bold = True
        axis.HorizontalAlignment = -4108
        axis.Interior.Color = rgb(SILVER)
        if m.last >= 0:
            actual_header = ws.Range(f"J3:{m.col(m.last)}5")
            actual_header.Interior.Color = rgb(DARK)
            actual_header.Font.Color = rgb(WHITE)
        ws.Range(f"J3:{last_month}3").NumberFormat = "0"
        ws.Range(f"J5:{last_month}5").NumberFormat = "yyyy-mm"
        if m.years:
            first_year = m.col(len(m.months))
            last_year = m.col(len(m.months) + len(m.years) - 1)
            ws.Range(f"{first_year}3:{last_year}5").Interior.Color = rgb(SILVER)
            ws.Range(f"{first_year}3:{last_year}5").Font.Bold = True
            ws.Range(f"{first_year}3:{last_year}5").NumberFormat = "0"

    for row in sorted(m.headers[name]):
        band = ws.Range(f"B{row}:{right}{row}")
        band.Interior.Color = rgb(SILVER)
        band.Font.Bold = True
        _edge(band, 9, BLACK)  # xlEdgeBottom
    for row in sorted(m.sections[name]):
        band = ws.Range(f"B{row}:{right}{row}")
        band.Font.Bold = True
        band.RowHeight = 23
        _edge(band, 9, ORANGE)
    for row in sorted(m.totals[name]):
        band = ws.Range(f"B{row}:{right}{row}")
        band.Font.Bold = True
        _edge(band, 8, ORANGE if row - 1 in m.sections[name] else BLACK)  # xlEdgeTop
    formats = defaultdict(list)
    for row, fmt in m.formats[name].items():
        formats[fmt].append(f"G{row}:{right}{row}")
    for fmt, addresses in formats.items():
        for band in _batches(ws, addresses):
            set_number_format(band, fmt)
    if name == "IS" and hasattr(m, "credit_metadata_bounds"):
        first, last = m.credit_metadata_bounds
        for column in ("J", "N"):
            set_number_format(ws.Range(f"{column}{first}:{column}{last}"), "yyyy-mm")

    if name == "전제사항":
        # A loss-vintage date and its monetary balance share the same row.
        for _, label, sheet, address, _ in m.inputs:
            if sheet == name and label == "결손금 발생 월":
                ws.Range(address).NumberFormat = "yyyy-mm"

    if name == "목차":
        _toc_layout(ws, m)
    elif name == "모델정보":
        # Metadata tables start in column A; their named ranges include the headers.
        for key, (sheet, address) in m.names.items():
            if sheet == name and key.startswith("META_"):
                left, top, end, _ = range_boundaries(address)
                band = ws.Range(f"{letter(left)}{top}:{letter(end)}{top}")
                band.Interior.Color = rgb(SILVER)
                band.Font.Bold = True
        ws.Range(f"A3:{right}{bottom}").WrapText = True
        ws.Range(f"A3:{right}{bottom}").Rows.AutoFit()
    elif name == "기준값갱신":
        ws.Range("C4").NumberFormat = "yyyy-mm-dd hh:mm"
        for address in ("B7:D7", "E16:G16"):
            header = ws.Range(address)
            header.Font.Bold = True
            header.Interior.Color = rgb(SILVER)
    elif name not in SPECIAL:
        _labels(ws, m, name, widths)
        if name == "금융조건" and bottom >= 27:
            # Preserve facility addresses while bringing the table into the opening view.
            first_body_row = min(r for r, c in grid if r > 5)
            if first_body_row > 6:
                ws.Rows(f"6:{first_body_row - 1}").Hidden = True
            table = ws.Range(f"B27:G{bottom}")
            table.WrapText = True
            table.Rows.AutoFit()
            for row in range(28, bottom + 1):
                if all((row, c) in grid for c in (2, 3, 4, 7)):
                    set_number_format(ws.Range(f"G{row}"), "0.0%")
        if name == "검증":
            rule = ws.Range(f"G3:G{bottom}").FormatConditions.Add(1, 3, '="NO"')
            rule.Font.Color = rgb(RED)
            rule.Font.Bold = True

    for band in _batches(ws, m.paint[name]):
        band.Interior.Color = rgb(YELLOW)
        band.Font.Color = rgb(BLUE)

    ws.Application.Goto(ws.Range("A1"), True)
    window = ws.Application.ActiveWindow
    window.DisplayGridlines = False
    window.FreezePanes = False
    window.SplitColumn = 9 if monthly else 0
    window.SplitRow = 5 if monthly else 2
    window.FreezePanes = True


def _numeric_slots(m, name):
    """Include value slots which may display a placeholder or have no observation yet."""
    # Numeric manual inputs should keep their type when initially blank (tax
    # balances, monthly overrides, assumptions for future facilities, etc.).
    yield from m.paint[name]
    rows = sorted(m.monthly_rows[name])
    last_col = m.col(len(m.months) + len(m.years) - 1)
    for first, last in _consecutive(rows):
        yield f"H{first}:H{last}"
        yield f"J{first}:{last_col}{last}"
    if getattr(m, "combined", False):
        yield from m.numeric_slots.get(name, ())
        return
    if name == "Sensitivity":
        first, last = m.dashboard_sensitivity[:2]
        yield f"C{first}:{letter(4+len(m.cases))}{last}"
        yield "C6"
    elif name == "기준값갱신":
        yield f"E17:G{16+len(m.live_labels)}"
        yield "C4"
    elif name == "민감도및결과":
        from dashboard import ANNUAL, FIRST_METRIC
        yield "C6:M7"
        yield f"C{FIRST_METRIC}:{letter(2+2*len(m.years))}{FIRST_METRIC+len(ANNUAL)-1}"
        base, custom, _, _, _, header = m.dashboard_results
        yield f"C{base+2}:E{base+1+len(m.live_labels)}"
        if len(m.cases) > 3:
            yield f"C{custom+2}:{letter(len(m.cases)-1)}{custom+1+len(m.live_labels)}"
        yield f"F{header+1}:N{header+len(m.months)}"
        scenario = m.chart_sources["scenario"]
        first = scenario["header"] + 1
        last = first + scenario["count"] - 1
        # Columns F/G are TEXT label strings; only the numeric series/metric
        # columns stay Arial through unsaved and no-debt modes.
        yield f"C{first}:E{last}"
        yield f"H{first}:H{last}"


def _consecutive(rows):
    if not rows:
        return
    start = end = rows[0]
    for row in rows[1:]:
        if row != end + 1:
            yield start, end
            start = row
        end = row
    yield start, end


def style_typed_cells(wb, m):
    """Apply the numeric face after recalculation, in two typed selections per sheet."""
    import pywintypes
    for name in m.sheets:
        ws = wb.Worksheets(name)
        used = ws.UsedRange
        for kind in (2, -4123):  # xlCellTypeConstants, xlCellTypeFormulas
            try:
                numeric = used.SpecialCells(kind, 1)  # xlNumbers includes dates.
            except pywintypes.com_error as exc:
                if not exc.excepinfo or exc.excepinfo[5] != -2146827284:
                    raise
                continue  # No matching cells in this worksheet.
            numeric.Font.Name = NUMBER_FONT
        for band in _batches(ws, _numeric_slots(m, name)):
            band.Font.Name = NUMBER_FONT
