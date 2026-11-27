# -*- coding: utf-8 -*-
"""List files in an Input folder and guess which planning sources they are."""
import argparse
import csv
import re
import sys
from datetime import date, datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from common import say, write_text, wrote
from excel_retry import cleanup_excel, quit_excel, retry_excel

WORDS = (
    "월", "계정", "차변", "대변", "코드", "금액", "잔액", "만기", "시작", "종료",
    "수량", "단가", "고객", "약정", "month", "account", "debit", "credit", "자산", "부채",
)
TAGS = (
    ("분개장", ("차변", "대변", "전표", "debit", "credit", "분개")),
    ("손익", ("손익", "매출액", "판관비", "매출원가", "영업이익", "당기순이익")),
    ("재무상태표", ("재무상태", "자산총계", "부채와자본", "유동자산", "자본총계")),
    ("상환스케줄", ("상환스케줄", "분할상환", "납입월", "회차")),
    ("차입", ("차입", "약정", "금리", "만기", "이율", "lender")),
    ("계약", ("계약", "월이용료", "청구액", "수주")),
    ("판매계획", ("판매계획", "계획매출", "실적 아님")),
    ("판매수량", ("수량", "단가", "판매실적")),
    ("계정과목", ("계정과목", "계정코드", "차감")),
    ("인원", ("인원", "직군", "headcount", "급여명세")),
    ("고정자산", ("고정자산", "내용연수", "감가상각방법", "취득가")),
)
MONTH_EXACT = re.compile(r"^(20\d{2})\s*(?:[-./]\s*(0[1-9]|1[0-2])|년\s*(0?[1-9]|1[0-2])\s*월)$")
MONTH_HEADERS = {"월", "납입월", "귀속월", "기준월", "month"}
DATE_HEADERS = {"일자", "전표일자", "거래일자", "전기일자", "날짜", "거래일", "date"}
DATE_EXACT = re.compile(r"^(20\d{2})[-./](0[1-9]|1[0-2])(?:[-./](0[1-9]|[12]\d|3[01]))?$")
UNIT_TEXT = re.compile(r"단위\s*[:：]\s*(\S(?:.{0,118}\S)?)")


def text_of(value):
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%Y-%m")
    if isinstance(value, date):
        return value.strftime("%Y-%m")
    return str(value).strip()


def exact_month(value, serial=False):
    """A cell that is itself a month. A date inside a sentence is not a month column."""
    if isinstance(value, datetime):
        return f"{value.year:04d}-{value.month:02d}"
    if isinstance(value, date):
        return f"{value.year:04d}-{value.month:02d}"
    if serial and isinstance(value, (int, float)) and not isinstance(value, bool) and 30000 < float(value) < 60000:
        stamp = datetime(1899, 12, 30) + __import__("datetime").timedelta(days=float(value))
        if 1990 <= stamp.year <= 2100 and abs(float(value) - round(float(value))) < 1e-6:
            return f"{stamp.year:04d}-{stamp.month:02d}"
        return None
    if not isinstance(value, str):
        return None
    match = MONTH_EXACT.match(value.strip())
    if not match:
        return None
    month = next(group for group in match.groups()[1:] if group)
    return f"{int(match.group(1)):04d}-{int(month):02d}"


def header_score(values):
    texts = [text_of(v) for v in values if isinstance(v, str) and text_of(v)]
    numbers = sum(isinstance(v, (int, float)) and not isinstance(v, bool) for v in values)
    if len(texts) < 2:
        return 0
    hits = sum(any(word.lower() in item.lower() for word in WORDS) for item in texts)
    return hits * 10 + min(len(texts), 20) - numbers * 3


def classify(blob):
    found = []
    for tag, keys in TAGS:
        if any(key.lower() in blob.lower() for key in keys):
            found.append(tag)
    if "차입" in found and "상환스케줄" in found:
        pass
    return found or ["기타"]


def read_csv_rows(path):
    raw = path.read_bytes()
    text = None
    encoding = "utf-8"
    for enc in ("utf-8-sig", "utf-8", "cp949"):
        try:
            text = raw.decode(enc)
            encoding = enc
            break
        except UnicodeDecodeError:
            continue
    if text is None:
        return encoding, []
    rows = list(csv.reader(text.splitlines()))
    return encoding, rows


def unit_phrases(rows):
    found = []
    for row in rows[:15]:
        for cell in row:
            if cell is None or isinstance(cell, (datetime, date, int, float)):
                continue
            text = str(cell).replace("\n", " ").strip()
            if text.startswith("단위") or "단위:" in text or "단위：" in text:
                found.append(text)
    return found[:3]


def column_titles(row):
    titles = []
    for cell in row:
        text = text_of(cell)
        if text and len(text) <= 24:
            titles.append(text)
    return titles


def body_month(value):
    month = exact_month(value, serial=False)
    if month:
        return month
    if not isinstance(value, str):
        return None
    match = DATE_EXACT.match(value.strip())
    if not match:
        return None
    return f"{int(match.group(1)):04d}-{int(match.group(2)):02d}"


def month_key(month):
    year, number = map(int, month.split("-"))
    return year * 12 + number


def consecutive_months(months):
    ordered = sorted(set(months))
    if len(ordered) < 3:
        return []
    keys = [month_key(month) for month in ordered]
    if all(right - left == 1 for left, right in zip(keys, keys[1:])):
        return ordered
    return []


def row_blob(row):
    return " ".join(text_of(cell) for cell in row)


def header_period(rows):
    """A monthly table needs consecutive month labels in the header, not dates in the body."""
    best_i, best_score, best_row = None, 0, []
    header_months = []
    named_months = set()
    for index, row in enumerate(rows[:30], 1):
        score = header_score(row)
        if len(column_titles(row)) >= 3 and score > best_score:
            best_score, best_i, best_row = score, index, row
        labels = [text_of(cell).strip().lower() for cell in row]
        months_here = [item for item in (exact_month(cell, serial=True) for cell in row) if item]
        if score > 0 and len(months_here) >= 3:
            header_months.extend(months_here)
        if any(label in MONTH_HEADERS or label in DATE_HEADERS for label in labels):
            for col, label in enumerate(labels):
                if label not in MONTH_HEADERS and label not in DATE_HEADERS:
                    continue
                for body in rows[index:]:
                    if "기초" in row_blob(body):
                        continue
                    month = body_month(body[col]) if col < len(body) else None
                    if month:
                        named_months.add(month)
    run = consecutive_months(header_months)
    titles = column_titles(best_row)
    if run:
        layout, found = "월별 표", run
    else:
        layout = "표" if len(titles) >= 3 else "표가 아님"
        found = sorted(named_months)
    return (best_i if best_score else None), best_row, found, layout


def sheet_report(title, rows, merged=0, filename=""):
    header_i, header, found, layout = header_period(rows)
    blob = " ".join(text_of(cell) for row in rows[:8] for cell in row)
    nonempty = sum(1 for row in rows if any(text_of(cell) for cell in row))
    units = unit_phrases(rows)
    header_text = " ".join(text_of(c) for c in header if isinstance(c, str))
    return {
        "title": title,
        "rows": nonempty,
        "header": header_i,
        "header_preview": [text_of(c) for c in header if text_of(c)][:12],
        "months": (found[0], found[-1]) if found else None,
        "layout": layout,
        "tags": classify(" ".join((filename, title, header_text))),
        "merged": merged,
        "units": units[:3],
    }


def xlsx_sheets(path):
    from openpyxl import load_workbook
    book = load_workbook(path, read_only=False, data_only=False)
    reports = []
    try:
        for ws in book.worksheets:
            rows = [list(row) for row in ws.iter_rows(max_row=min(ws.max_row or 1, 400), max_col=min(ws.max_column or 1, 80), values_only=True)]
            merged = len(ws.merged_cells.ranges)
            reports.append(sheet_report(ws.title, rows, merged, path.name))
    finally:
        book.close()
    return reports


def xls_sheets(path):
    from win32com.client import DispatchEx
    reports = []
    excel = retry_excel(lambda: DispatchEx("Excel.Application"))
    with cleanup_excel(lambda: quit_excel(lambda: excel.Quit())):
        retry_excel(lambda: setattr(excel, "Visible", False))
        retry_excel(lambda: setattr(excel, "DisplayAlerts", False))
        retry_excel(lambda: setattr(excel, "AutomationSecurity", 3))
        book = retry_excel(lambda: excel.Workbooks.Open(str(path.resolve()), UpdateLinks=0, ReadOnly=True))
        with cleanup_excel(lambda: retry_excel(lambda: book.Close(False))):
            count = retry_excel(lambda: book.Worksheets.Count)
            for index in range(1, count + 1):
                ws = retry_excel(lambda: book.Worksheets.Item(index))
                values = retry_excel(lambda: ws.UsedRange.Value2)
                if values is None:
                    rows = []
                elif not isinstance(values, tuple):
                    rows = [[values]]
                elif values and not isinstance(values[0], tuple):
                    rows = [[value] for value in values]
                else:
                    rows = [list(row) for row in values]
                title = retry_excel(lambda: ws.Name)
                reports.append(sheet_report(str(title), rows[:400], 0, path.name))
    return reports


def pdf_pages(path):
    data = path.read_bytes()
    return len(re.findall(rb"/Type\s*/Page(?!s)", data))


def describe_file(path):
    suffix = path.suffix.lower()
    info = {"path": path.name, "suffix": suffix or "(none)", "bytes": path.stat().st_size, "sheets": [], "error": ""}
    try:
        if suffix in {".xlsx", ".xlsm", ".xltx"}:
            info["sheets"] = xlsx_sheets(path)
        elif suffix == ".xls":
            info["sheets"] = xls_sheets(path)
        elif suffix == ".csv":
            encoding, rows = read_csv_rows(path)
            report = sheet_report(path.stem, rows, 0, path.name)
            report["encoding"] = encoding
            info["sheets"] = [report]
        elif suffix == ".pdf":
            info["sheets"] = [{"title": path.name, "rows": pdf_pages(path), "pages": pdf_pages(path), "tags": ["PDF"], "header": None, "months": None, "units": [], "merged": 0, "header_preview": []}]
        else:
            info["sheets"] = [{"title": path.name, "rows": 0, "tags": ["기타"], "header": None, "months": None, "units": [], "merged": 0, "header_preview": []}]
    except Exception as exc:
        info["error"] = type(exc).__name__ + ": " + str(exc)[:180]
    return info


def coverage(files):
    tags = set()
    for item in files:
        for sheet in item["sheets"]:
            tags.update(sheet.get("tags") or [])
    required = "손익" in tags or "분개장" in tags
    lines = ["## 계약 자료와 대조", ""]
    lines.append(f"- 필수 (월별 손익 또는 분개장): {'있음' if required else '없음'}")
    lines.append(f"- 권장 월별 재무상태표: {'있음' if '재무상태표' in tags else '없음'}")
    lines.append(f"- 권장 차입 약정: {'있음' if '차입' in tags else '없음'}")
    lines.append(f"- 선택 상환 스케줄: {'있음' if '상환스케줄' in tags else '없음'}")
    lines.append(f"- 선택 계약·매출 목록: {'있음' if '계약' in tags else '없음'}")
    lines.append(f"- 선택 판매 수량·단가: {'있음' if '판매수량' in tags or '판매계획' in tags else '없음'}")
    lines.append(f"- 선택 인원: {'있음' if '인원' in tags else '없음'}")
    lines.append(f"- 선택 고정자산 대장: {'있음' if '고정자산' in tags else '없음'}")
    if not required:
        lines.append("")
        lines.append("필수 자료가 없으면 손익 모델을 만들 수 없다.")
    return "\n".join(lines)


def render(folder, files):
    lines = [f"# 입력 자료 목록", "", f"폴더: `{folder}`", ""]
    if not files:
        lines.append("파일이 없다.")
        lines.append("")
        lines.append(coverage(files))
        return "\n".join(lines) + "\n"
    lines.append("| 파일 | 형식 | 크기 | 추정 |")
    lines.append("|---|---|---:|---|")
    for item in files:
        tags = sorted({tag for sheet in item["sheets"] for tag in sheet.get("tags") or []})
        lines.append(f"| {item['path']} | {item['suffix']} | {item['bytes']} | {', '.join(tags) or item['error'] or '확인 못 함'} |")
    for item in files:
        lines.append("")
        lines.append(f"## {item['path']}")
        lines.append("")
        if item["error"]:
            lines.append(f"읽기 실패: {item['error']}")
            continue
        if item["suffix"] == ".pdf":
            pages = item["sheets"][0]["rows"] if item["sheets"] else 0
            lines.append(f"PDF. 쪽수 추정 {pages}. 본문은 읽지 않았다.")
            continue
        for sheet in item["sheets"]:
            lines.append(f"### {sheet['title']}")
            lines.append("")
            extra = f", 인코딩 {sheet['encoding']}" if sheet.get("encoding") else ""
            lines.append(f"- 행 수(내용 있는 행, 상위 400행까지): {sheet['rows']}{extra}")
            if sheet.get("merged"):
                lines.append(f"- 병합 범위: {sheet['merged']}개")
            lines.append(f"- 형식: {sheet.get('layout') or '확인 못 함'}")
            if sheet.get("header"):
                lines.append(f"- 머리글로 보이는 행: {sheet['header']}행 ({', '.join(sheet['header_preview'])})")
            else:
                lines.append("- 머리글: 확인 못 함")
            if sheet.get("layout") == "표가 아님":
                lines.append("- 월·기간 추정: 확인 못 함. 항목과 문장 나열이라 월 열을 세지 않는다.")
            elif sheet.get("months"):
                lines.append(f"- 월·기간 추정: {sheet['months'][0]} ~ {sheet['months'][1]}")
            else:
                lines.append("- 월·기간 추정: 확인 못 함")
            if sheet.get("units"):
                lines.append("- 단위 문구: " + "; ".join(sheet["units"]))
            lines.append("- 추정 종류: " + ", ".join(sheet["tags"]))
    lines.append("")
    lines.append(coverage(files))
    lines.append("")
    return "\n".join(lines) + "\n"


def main():
    parser = argparse.ArgumentParser(description="Inventory an Input folder.")
    parser.add_argument("input_dir", type=Path)
    parser.add_argument("--out", required=True, type=Path)
    args = parser.parse_args()
    folder = args.input_dir
    if not folder.is_dir():
        say("FAIL input folder missing")
        return 2
    files = []
    for path in sorted(folder.rglob("*")):
        if not path.is_file() or path.name.startswith("~$"):
            continue
        files.append(describe_file(path))
    write_text(args.out, render(folder, files))
    say(f"FILES {len(files)}")
    wrote(args.out, "inventory")
    say("RESULT PASS")
    return 0


if __name__ == "__main__":
    sys.exit(main())
