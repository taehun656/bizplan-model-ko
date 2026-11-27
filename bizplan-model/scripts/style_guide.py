"""Build the workbook's style reference from the same tokens as its worksheets and charts."""
import re

from openpyxl.utils import get_column_letter

from design_palette import (
    BLACK, BLUE, BRIDGE_COLORS, CASH_COLORS, DARK, FONT, MONTHLY_COLORS,
    NUMBER_FONT, ORANGE, RED, REVENUE_COLORS, SCENARIO_COLORS, SILVER, TAB_COLORS,
    WHITE, YELLOW, rgb,
)
from workbook_style import GROUPS, NUMBER, _wrapped_height
from dashboard import MONTH, MULTIPLE, PERCENT

SHEET = "스타일 가이드"
WIDTHS = (3, 24, 30, 16, 16, 55)


def _codes(colors):
    return " · ".join("#" + color for color in colors)


def _palettes():
    return (
        ("본표·가정·검증 띠", "첫 행, 선택 표시, 단추", ORANGE, WHITE, "구역 제목 아래 선에도 주황을 쓴다."),
        ("실적·원천 띠", "원천 시트 첫 행", DARK, WHITE, "보관된 실적과 계산 시트를 구분한다."),
        ("시트 제목", "두 번째 행", SILVER, WHITE, "시트 제목은 흰색 굵은 글씨로 쓴다."),
        ("표 머리글", "요약·관리 표와 KPI 카드", SILVER, BLACK, "선택 케이스 카드 머리는 주황으로 표시한다."),
        ("입력 셀", "직접 바꿀 수 있는 가정", YELLOW, BLUE, "입력 단위와 적용 기간을 함께 표시한다."),
        ("본문", "계산값, 연결 수식, 항목명", None, BLACK, "표 안에서 본문 배경은 비운다."),
        ("경고", "NO, 갱신 필요, 음수", None, RED, "판정 문구는 굵게 쓴다."),
        ("테두리", "표 외곽, 열 구분, 머리글 아래", BLACK, None, "검정 얇은 실선으로 마감한다."),
        ("가정·상세 탭", "시트 탭", TAB_COLORS[1], None, "시트 본문 채움에는 사용하지 않는다."),
        ("검증·관리 탭", "시트 탭", TAB_COLORS[3], None, "시트 본문 채움에는 사용하지 않는다."),
    )


def setup_guide(m):
    toc = m.grid["목차"]
    sections = sorted(r for r in m.sections["목차"] if toc[r, 2][1:] in GROUPS)
    scopes = ("핵심 입력, 재무제표와 결과 화면", "동인별 입력과 상세 계산",
              "실적값과 원천 자료", "검사, 저장 기준, 모델 문서")
    directory = []
    for i, start in enumerate(sections):
        rows = []
        r = start + 2
        while (r, 12) in toc and r not in m.headers["목차"]:
            rows.append(r)
            r += 1
        # Limit each cell's text so large workbooks remain readable in Excel.
        for offset in range(0, len(rows), 8):
            chunk = rows[offset:offset+8]
            references = [f"'목차'!C{n}" for n in chunk]
            formula = "=" + '&", "&'.join(references)
            directory.append((GROUPS[i], scopes[i], "#" + (DARK if i == 2 else ORANGE),
                              "#" + TAB_COLORS[i], formula))

    palette = _palettes()
    tables = (
        ("시트 구분과 탭 색", ("구분", "역할", "첫 행 띠", "시트 탭", "시트 목록 · 목차 연결"), directory),
        ("색상표", ("용도", "적용 위치", "색", "글씨", "사용 기준"),
         [(label, where, "#" + fill if fill else "채움 없음", "#" + ink if ink else "—", note)
          for label, where, fill, ink, note in palette]),
        ("시트 머리말", ("위치", "작성 기준"), (
            ("첫 행 · 모델명", "모델명을 왼쪽에 두고 시트 구분에 맞는 색으로 제목 띠를 채운다."),
            ("둘째 행 · 제목", "시트명을 회색 띠에 흰색 굵은 글씨로 표시한다."),
            ("단위와 시간축", "금액의 표시 단위를 밝힌다. 월별 표에는 연도·실적 또는 추정 구분·월을 함께 표시한다."),
            ("틀 고정", "월별 표에서는 시간축과 항목명이 스크롤해도 보이게 고정한다. 관리 표에서는 두 제목 행을 고정한다."),
        )),
        ("표와 본문", ("요소", "작성 기준"), (
            ("글꼴 · 셀", f"문자와 혼합 문구는 {FONT}, 숫자·날짜와 숫자 결과 수식은 {NUMBER_FONT}을 쓴다. 해당 없음이 될 수 있는 DSCR도 숫자 칸이다. 기본 글꼴은 {FONT}, 본문은 10pt로 한다."),
            ("글꼴 · 차트와 도형", f"라틴 문자·숫자는 {NUMBER_FONT}, 한글 등 동아시아 문자는 {FONT}으로 표시한다. 차트 본문 9pt, 제목 10.8pt를 사용하고 시트 눈금선은 숨긴다."),
            ("구역과 합계", "구역 제목에는 굵은 글씨와 주황색 밑줄을 쓰고 합계 행에는 굵은 글씨와 윗선을 쓴다."),
            ("요약·관리 표", "머리글은 회색 채움과 검정 굵은 글씨로 표시한다. 표 외곽과 열 경계는 얇은 검정 실선으로 구분한다."),
            ("입력과 선택", "직접 입력값에는 단위와 근거를 적는다. 선택된 케이스는 주황 채움과 흰 글씨로 구분한다."),
            ("경고", "검증 실패와 갱신 필요 상태는 빨간 굵은 글씨로 표시한다. 금융 지표의 경고에는 적용 기준을 함께 표시한다."),
        )),
        ("숫자 서식", ("값", "표시 형식", "적용 예"), (
            ("금액", NUMBER, "양수는 천 단위 구분, 음수는 괄호, 0은 대시"),
            ("비율", PERCENT, "이익률과 전년 대비 변화율"),
            ("배수", MULTIPLE, "DSCR와 시나리오 배수는 소수 둘째 자리까지"),
            ("계수", "0.000", "계산에 쓰는 계수는 필요한 자릿수까지"),
            ("날짜와 월", "yyyy-mm-dd  /  " + MONTH, "기준일과 시간축을 구분하고 차트의 월은 yy-mm로 표시"),
        )),
        ("차트 색과 표시", ("차트·요소", "계열별 색", "표시 기준"), (
            ("연도별 매출", _codes(REVENUE_COLORS), "계획 연도 매출 순위대로 적용한다. 기타는 마지막 색을 쓴다."),
            ("매출 브리지", "양 끝 · 증가 · 감소\n" + _codes(BRIDGE_COLORS), "받침은 채움과 선을 숨기고 증감액은 막대 위에 둔다."),
            ("시나리오", "일반 · 선택 · 기준선\n" + _codes(SCENARIO_COLORS), "기준 1.0배는 빨간 점선으로 표시한다. DSCR 적용 대상이 없으면 영업이익을 표시한다."),
            ("월별 손익", "매출 · EBITDA · 영업이익\n" + _codes(MONTHLY_COLORS), "손익 세 계열은 실선으로 표시한다."),
            ("월별 현금·차입", "기말 현금 · 차입금 · 최소 보유현금\n" + _codes(CASH_COLORS), "최소 보유현금은 빨간 점선으로 표시한다."),
            ("축과 배경", f"눈금선 #{SILVER} · 축과 글씨 #{BLACK}\n배경 #{WHITE}", "가로 눈금선만 가늘게 표시한다. 차트 바깥과 그림 영역의 테두리는 숨긴다."),
        )),
        ("새 시트 작성 순서", ("순서", "작성 기준"), (
            ("1 · 구분과 위치", "시트 역할을 정하고 해당 구역에 탭을 배치한다. 역할에 맞는 탭 색과 제목 띠를 적용한다."),
            ("2 · 머리말과 본문", "모델명과 시트 제목을 표시하고 단위, 시간축, 표 머리글, 입력 셀과 숫자 형식을 맞춘다."),
            ("3 · 등록", "목차에 시트명·설명·이동 링크를 등록하고 모델정보의 시트 목록에 역할과 계산 단계를 기록한다."),
            ("4 · 확인", "연결 수식과 차트 계열을 다시 계산한 뒤 결과 시트의 검증 건수와 핵심 지표를 대조한다."),
        )),
        ("변경 금지 사항", ("대상", "유지할 내용"), (
            ("계산·연결 수식", "서식을 수정할 때 수식과 실적값을 보존한다. 원천 자료를 교체할 때에는 대사를 다시 한다."),
            ("케이스와 저장값", "지정된 입력 셀에서 케이스를 고른다. 저장값과 기준값은 각 갱신 단추로 계산한다."),
            ("이름 정의와 매크로", "민감도·기준값·메타데이터의 범위 이름, 단추 이름과 매크로 연결을 유지한다."),
            ("참조 범위", "관리 표가 주소를 보관하는 영역의 구조를 유지한다. 구조가 바뀌면 참조 범위와 검증 항목을 함께 대조한다."),
        )),
    )
    layout = []
    row = 4
    for no, (heading, headers, records) in enumerate(tables, 1):
        # The two-column tables span C:F; three-column tables use C:E for their middle field.
        cols = (2, 3) if len(headers) == 2 else (2, 3, 6) if len(headers) == 3 else (2, 3, 4, 5, 6)
        m.section(SHEET, row, f"{no}. {heading}")
        m.header(SHEET, row + 1)
        items = [headers, *records]
        for delta, fields in enumerate(items, 1):
            for col, field in zip(cols, fields):
                kind = "formula" if no == 1 and delta > 1 and col == 6 else "text"
                m.cell(SHEET, row + delta, col, field, kind)
        layout.append((row, row + len(items), cols))
        row += len(items) + 2
    m.guide_layout = layout


def style_guide(ws, m):
    """Lay out the guide at a fixed readable width and keep examples within the palette."""
    for col, width in enumerate(WIDTHS, 1):
        ws.Columns(get_column_letter(col)).ColumnWidth = width
    bottom = m.guide_layout[-1][1]
    for row in (1, 2):
        ws.Range(f"A{row}:J{row}").UnMerge()
        ws.Range(f"A{row}:F{row}").Merge()
    ws.Range(f"G1:J{bottom}").ClearFormats()
    for section, end, cols in m.guide_layout:
        ws.Range(f"B{section}:F{section}").Merge()
        band = ws.Range(f"B{section+1}:F{end}")
        if len(cols) < 5:
            last = "F" if len(cols) == 2 else "E"
            ws.Range(f"C{section+1}:{last}{end}").Merge(True)
        for side in (7, 8, 9, 10, 11):  # outside edges and inner vertical separators
            edge = band.Borders(side)
            edge.LineStyle = 1
            edge.Weight = 2
            edge.Color = rgb(BLACK)
        # Adjacent table edges share one Excel pixel. Restore the section rule
        # after the table outline so its orange underline remains visible.
        rule = ws.Range(f"B{section}:F{section}").Borders(9)
        rule.LineStyle = 1
        rule.Weight = -4138  # xlMedium
        rule.Color = rgb(ORANGE)
        head = ws.Range(f"B{section+1}:F{section+1}")
        head.HorizontalAlignment = -4108
        ws.Range(f"B{section+2}:F{end}").WrapText = True
        for row in range(section + 2, end + 1):
            height = 21
            for j, col in enumerate(cols):
                raw = m.grid[SHEET][row, col]
                if raw.startswith("="):
                    # TOC formulas can be uncalculated while the book is assembled.
                    text = ", ".join(m.grid["목차"][int(n), 3][1:] for n in re.findall(r"!C(\d+)", raw))
                else:
                    text = raw.removeprefix("'")
                end_col = cols[j + 1] - 1 if j < len(cols) - 1 else 6
                width = sum(WIDTHS[col - 1:end_col])
                height = max(height, _wrapped_height(text, width))
            ws.Rows(row).RowHeight = height
        ws.Rows(end + 1).RowHeight = 7
    section, _end, _cols = m.guide_layout[1]
    for row, (_, _, fill, ink, _) in enumerate(_palettes(), section + 2):
        code = ws.Range(f"D{row}:E{row}")
        if ink and fill:
            code.Interior.Color = rgb(fill)
        if ink:
            code.Font.Color = rgb(ink)
        code.Font.Bold = True
