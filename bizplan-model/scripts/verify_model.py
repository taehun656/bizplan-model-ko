# -*- coding: utf-8 -*-
"""Recalculate a model in a separate Excel instance and check it."""
import argparse
import csv
import re
import sys
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from common import say, write_text, wrote
from excel_retry import BUSY_MESSAGE, ExcelShutdownBusy, busy, cleanup_after_error, cleanup_excel, quit_excel, retry_excel
from palette_audit import inspect_palette, warnings_for
from font_audit import inspect_fonts, font_warnings_for
from workbook_scan import constant_cells, covers, load_meta, model_is_stale, parse_box, scan_formulas

XL_CALC = -4105
ERROR_INTS = {2000, 2007, 2015, 2023, 2029, 2036, 2042}


def letter(n):
    out = ""
    while n:
        n, rem = divmod(n - 1, 26)
        out = chr(65 + rem) + out
    return out


def months_of(start, end):
    y, m = map(int, str(start).split("-"))
    stop = str(end)
    out = []
    while True:
        cur = f"{y:04d}-{m:02d}"
        out.append(cur)
        if cur >= stop:
            break
        m += 1
        if m == 13:
            y, m = y + 1, 1
    return out


def read_csv(path):
    with path.open(encoding="utf-8-sig", newline="") as handle:
        return list(csv.DictReader(handle))


def month_text(value):
    if isinstance(value, datetime):
        return f"{value.year:04d}-{value.month:02d}"
    if isinstance(value, (int, float)) and 30000 < float(value) < 80000:
        stamp = datetime(1899, 12, 30) + timedelta(days=float(value))
        return f"{stamp.year:04d}-{stamp.month:02d}"
    return str(value)[:7]


def num(value):
    return float(str(value).replace(",", ""))


def flat(value):
    if value is None:
        return []
    if not isinstance(value, tuple):
        return [value]
    if not value:
        return []
    if isinstance(value[0], tuple):
        if len(value) == 1:
            return list(value[0])
        if all(isinstance(row, tuple) and len(row) == 1 for row in value):
            return [row[0] for row in value]
    return list(value)


def is_error(value):
    if isinstance(value, int) and (value in ERROR_INTS or value < -2146820000):
        return True
    return False


def same(left, right):
    if isinstance(left, str) or isinstance(right, str):
        return str(left).strip() == str(right).strip()
    try:
        return abs(float(left) - float(right)) <= 1e-6
    except (TypeError, ValueError):
        return left == right


def sheet_labels(ws, last=400):
    """First text in columns B:F of each row. Layout may indent the caption."""
    raw = ws.Range(f"B1:F{last}").Value2
    if raw is None:
        return []
    if not isinstance(raw, tuple):
        raw = ((raw,),)
    labels = []
    for row in raw:
        if not isinstance(row, tuple):
            row = (row,)
        text = ""
        for cell in row:
            if isinstance(cell, str) and cell.strip():
                text = cell.strip()
                break
        labels.append(text)
    return labels


def caption(text):
    head = text.split("(")[0].strip()
    if "대사" in text or text.startswith("원천") or "률" in head:
        return ""
    return head


def find_labeled_row(labels, keys):
    cleaned = []
    for index, text in enumerate(labels, 1):
        head = caption(text) if text else ""
        if head:
            cleaned.append((index, head, text))
    for key in keys:
        for index, head, text in cleaned:
            if head == key:
                return index, text
    return None, None


def find_meta_input(meta, needle):
    for item in meta.get("inputs") or []:
        label = str(item.get("label") or "")
        if needle in label:
            return str(item.get("sheet")), str(item.get("cell")), label
    return None


def find_label_anywhere(book, keys):
    for ws in book.Worksheets:
        labels = sheet_labels(ws)
        row, text = find_labeled_row(labels, keys)
        if row:
            return ws, row, text
    return None, None, None


def value_as_won(value, label, unit_hint, model_unit):
    """Won inputs keep '(원)' or H='원'. Display-unit inputs are scaled by Unit."""
    if value is None:
        return None
    hint = str(unit_hint).strip() if unit_hint is not None else ""
    if "(원)" in str(label) or hint == "원":
        return float(value)
    if hint and ("백만" in hint or "천원" in hint or "표시" in hint):
        return float(value) * model_unit
    if abs(float(value)) >= model_unit:
        return float(value)
    return float(value) * model_unit


def read_labeled_amount(book, meta, needle, keys, model_unit):
    found = find_meta_input(meta, needle)
    if found:
        sheet, addr, label = found
        cell = book.Worksheets(sheet).Range(addr.replace("$", "").split(":")[0])
        hint = book.Worksheets(sheet).Cells(cell.Row, cell.Column + 1).Value2
        return value_as_won(cell.Value2, label, hint, model_unit), f"{sheet}!{addr}"
    ws, row, text = find_label_anywhere(book, keys)
    if ws is None:
        return None, ""
    value = ws.Cells(row, 7).Value2
    hint = ws.Cells(row, 8).Value2
    return value_as_won(value, text, hint, model_unit), f"{ws.Name}!G{row}"


def find_axis_row(ws, months, start_col, end_col):
    best_row, best_hits = None, 0
    for row in range(1, 16):
        vals = row_values(ws, row, start_col, end_col)
        hits = 0
        for index, month in enumerate(months):
            if index < len(vals) and month_text(vals[index]) == month:
                hits += 1
        if hits > best_hits:
            best_row, best_hits = row, hits
    if best_hits < max(3, len(months) // 2):
        return None
    return best_row


def shift_year(month, years):
    year, number = month.split("-")
    return f"{int(year) + years:04d}-{number}"


def numeric_at(vals, index):
    if index is None or index >= len(vals) or not isinstance(vals[index], (int, float)):
        return None
    return float(vals[index])


def gap_warning(name, label, forecast_sum, actual_sum, forecast_n, actual_n):
    if actual_sum is None or forecast_sum is None:
        return f"{name}: {label} 값이 비어 확인 못 함"
    if abs(actual_sum) < 1e-9:
        if abs(forecast_sum) >= 1e-9:
            return f"{name}: {label} 실적 합계가 0이고 추정 합계는 {forecast_sum:.4g}"
        return None
    gap = abs(forecast_sum / actual_sum - 1)
    if gap > 0.25:
        return (
            f"{name}: {label} 추정 {forecast_n}개월 합계 {forecast_sum:.4g}, "
            f"실적 {actual_n}개월 합계 {actual_sum:.4g}, 차이 {gap * 100:.1f}%"
        )
    return None


def continuity_warnings(months, last_actual, series):
    """Warn on a 12-month sum break and on the same three calendar months a year earlier."""
    index = {month: i for i, month in enumerate(months)}
    actual = [month for month in months if month <= last_actual]
    forecast = [month for month in months if month > last_actual]
    if not actual or not forecast:
        return ["실적과 추정을 나란히 볼 달이 부족해 확인 못 함"]
    window = min(12, len(forecast), len(actual))
    warnings = []
    for name, vals in series:
        trail = actual[-window:]
        head = forecast[:window]
        actual_vals = [numeric_at(vals, index[month]) for month in trail]
        forecast_vals = [numeric_at(vals, index[month]) for month in head]
        if any(value is None for value in actual_vals + forecast_vals):
            warnings.append(f"{name}: 12개월 합계 비교에 빈 값이 있어 확인 못 함")
        else:
            text = gap_warning(name, f"첫 {window}개월", sum(forecast_vals), sum(actual_vals), window, window)
            if text:
                warnings.append(text)
        pairs = []
        for month in forecast[:3]:
            prior = shift_year(month, -1)
            if prior not in index:
                continue
            left = numeric_at(vals, index[month])
            right = numeric_at(vals, index[prior])
            if left is None or right is None:
                continue
            pairs.append((left, right))
        if not pairs:
            warnings.append(f"{name}: 전년 같은 달을 찾지 못해 확인 못 함")
            continue
        text = gap_warning(
            name, "전년 같은 달",
            sum(item[0] for item in pairs), sum(item[1] for item in pairs), len(pairs), len(pairs),
        )
        if text:
            warnings.append(text)
    return warnings


def row_values(ws, row, c1, c2):
    raw = ws.Range(f"{letter(c1)}{row}:{letter(c2)}{row}").Value2
    return flat(raw)


def sheet_or_none(book, name):
    try:
        return book.Worksheets(name)
    except Exception as exc:
        if busy(exc):
            raise
        return None


def com_errors(book):
    hits = []
    for ws in book.Worksheets:
        try:
            used = ws.UsedRange
            values = used.Value2
        except Exception as exc:
            if busy(exc):
                raise
            continue
        rows = values
        if rows is None:
            continue
        if not isinstance(rows, tuple):
            rows = ((rows,),)
        elif rows and not isinstance(rows[0], tuple):
            rows = tuple((item,) for item in rows)
        origin_row = used.Row
        origin_col = used.Column
        for r, row in enumerate(rows):
            for c, value in enumerate(row):
                if is_error(value) and len(hits) < 40:
                    hits.append(f"{ws.Name}!{letter(origin_col + c)}{origin_row + r}")
    return hits


def close_excel(excel, book):
    with cleanup_excel(lambda: quit_excel(lambda: excel.Quit())):
        if book is not None:
            retry_excel(lambda: book.Close(False))


def open_excel(path):
    from win32com.client import DispatchEx
    excel = DispatchEx("Excel.Application")
    book = None
    try:
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False
        excel.EnableEvents = False
        excel.AutomationSecurity = 3
        book = excel.Workbooks.Open(str(Path(path).resolve()), UpdateLinks=0, ReadOnly=True, IgnoreReadOnlyRecommended=True)
        try:
            excel.Calculation = XL_CALC
        except Exception as exc:
            if busy(exc):
                raise
        excel.CalculateFull()
    except BaseException as primary:
        cleanup_after_error(lambda: close_excel(excel, book), primary)
        raise
    return excel, book


def check_cases(excel, book, failures):
    try:
        case = book.Names("SN_Case").RefersToRange
        live = book.Names("SN_Live").RefersToRange
        store = book.Names("SN_Store").RefersToRange
        numbers = flat(book.Names("SN_CaseNos").RefersToRange.Value2)
    except Exception as exc:
        if busy(exc):
            raise
        failures.append("민감도 이름 정의를 읽지 못했다: " + type(exc).__name__)
        return "확인 못 함"
    original = case.Value2
    worst = 0.0
    bad = 0
    def restore_case():
        with cleanup_excel(lambda: retry_excel(lambda: excel.Calculate())):
            retry_excel(lambda: setattr(case, "Value2", original))

    with cleanup_excel(restore_case):
        for index, number in enumerate(numbers, 1):
            case.Value2 = number
            excel.Calculate()
            left = flat(live.Value2)
            right = flat(store.Columns(index).Value2)
            if len(left) != len(right):
                failures.append(f"케이스 {number}: 현재값 {len(left)}칸, 저장값 {len(right)}칸")
                bad += 1
                continue
            for item, stored in zip(left, right):
                if same(item, stored):
                    continue
                bad += 1
                try:
                    worst = max(worst, abs(float(item) - float(stored)))
                except (TypeError, ValueError):
                    worst = worst
                if bad <= 8:
                    failures.append(f"케이스 {number} 저장값이 현재 계산과 다름: {item} / {stored}")
    if bad:
        return f"실패 {bad}칸, 최대 차이 {worst}"
    return f"통과 ({len(numbers)}개 케이스)"


IS_LINES = (
    ("매출", ("매출",)),
    ("매출원가", ("매출원가",)),
    ("판매관리비", ("판매관리비", "판관비")),
    ("감가상각비", ("감가상각비", "상각비")),
    ("영업이익", ("영업이익",)),
)


def tie_out(book, meta, data, failures):
    info = meta["info"]
    unit = float(info.get("unit") or 1)
    months = months_of(info["first_month"], info["plan_end"])
    start_col = int(float(info["first_month_col"]))
    end_col = int(float(info["last_month_col"]))
    last_actual = str(info["last_actual"])
    if len(months) != end_col - start_col + 1:
        failures.append("META_Info 월 개수와 열 범위가 다름")
        return [], []
    lines = []
    income = sheet_or_none(book, "IS")
    if income is None:
        failures.append("IS 시트가 없다")
        return lines, []
    labels = sheet_labels(income)
    located = {}
    for name, keys in IS_LINES:
        row, text = find_labeled_row(labels, keys)
        located[name] = (row, text)
    if not located["매출"][0] or not located["영업이익"][0]:
        failures.append("IS에서 매출 또는 영업이익 라벨을 찾지 못했다")
        return lines, []
    axis_row = find_axis_row(income, months, start_col, end_col)
    if axis_row is None:
        failures.append("IS 월축 행을 찾지 못했다")
        return lines, []
    axis = row_values(income, axis_row, start_col, end_col)
    csv_rows = read_csv(Path(data) / "is_monthly.csv")
    bucket = {}
    for row in csv_rows:
        bucket.setdefault(row["month"], {}).setdefault(row["section"], 0.0)
        bucket[row["month"]][row["section"]] += num(row["amount"])
    series = {}
    for name, (row, _text) in located.items():
        series[name] = row_values(income, row, start_col, end_col) if row else []
    rev_vals = series["매출"]
    op_vals = series["영업이익"]
    rev_bad = op_bad = 0
    for i, month in enumerate(months):
        if month > last_actual:
            continue
        header_month = month_text(axis[i]) if i < len(axis) else ""
        if header_month != month:
            failures.append(f"IS 월축 {header_month} 가 {month} 와 다름")
            rev_bad += 1
            continue
        source = bucket.get(month, {})
        revenue = source.get("매출", 0.0)
        profit = revenue - source.get("매출원가", 0.0) - source.get("판관비", 0.0)
        for label, model, target, counter in (
            ("매출", rev_vals[i] if i < len(rev_vals) else None, revenue, "rev"),
            ("영업이익", op_vals[i] if i < len(op_vals) else None, profit, "op"),
        ):
            if model is None:
                failures.append(f"{month} {label} 값이 비어 있다")
                if counter == "rev":
                    rev_bad += 1
                else:
                    op_bad += 1
                continue
            gap = abs(float(model) * unit - target)
            if gap > 1:
                if counter == "rev":
                    rev_bad += 1
                else:
                    op_bad += 1
                if rev_bad + op_bad <= 8:
                    failures.append(f"{month} {label} 차이 {gap:.2f}원")
    lines.append(f"- 실적 매출 대사 ({located['매출'][1]}): {'실패 ' + str(rev_bad) + '개월' if rev_bad else '통과'}")
    lines.append(f"- 실적 영업이익 대사 ({located['영업이익'][1]}): {'실패 ' + str(op_bad) + '개월' if op_bad else '통과'}")
    continuity = continuity_warnings(
        months, last_actual,
        [(name, series[name]) for name, _keys in IS_LINES],
    )
    bs = sheet_or_none(book, "BS")
    if bs is None:
        lines.append("- BS 대차: 시트 없음")
    else:
        bs_labels = sheet_labels(bs)
        asset, _asset_text = find_labeled_row(bs_labels, ("자산총계",))
        claim, _claim_text = find_labeled_row(bs_labels, ("부채와 자본총계",))
        if not asset or not claim:
            failures.append("BS 총계 라벨을 찾지 못했다")
        else:
            assets = row_values(bs, asset, start_col, end_col)
            claims = row_values(bs, claim, start_col, end_col)
            bad = 0
            for i, month in enumerate(months):
                if i >= len(assets) or i >= len(claims):
                    break
                if assets[i] is None and claims[i] is None:
                    continue
                gap = abs(float(assets[i] or 0) - float(claims[i] or 0)) * unit
                if gap > 1:
                    bad += 1
                    if bad <= 6:
                        failures.append(f"{month} BS 대차 차이 {gap:.2f}원")
            lines.append(f"- BS 대차: {'실패 ' + str(bad) + '개월' if bad else '통과'}")
    cash = sheet_or_none(book, "CF")
    funding = sheet_or_none(book, "재원조달")
    if cash is None or funding is None:
        failures.append("CF 또는 재원조달 시트가 없다")
        return lines, continuity
    cash_row, _cash_text = find_labeled_row(sheet_labels(cash), ("기말 현금",))
    add_row, _add_text = find_labeled_row(sheet_labels(funding), ("추가 한도 사용 기말 잔액", "한도대출 기말 잔액"))
    old_row, _old_text = find_labeled_row(sheet_labels(funding), ("기존 한도 약정 사용 잔액", "기존 한도 잔액"))
    minimum, min_where = read_labeled_amount(book, meta, "최소 보유현금", ("최소 보유현금",), unit)
    limit, limit_where = read_labeled_amount(book, meta, "한도대출 총 한도", ("한도대출 총 한도", "한도대출 한도"), unit)
    if not all((cash_row, add_row, old_row)) or minimum is None or limit is None:
        failures.append("최소 현금, 한도, 기말 현금, 한도 잔액 중 라벨을 찾지 못했다")
        return lines, continuity
    cash_vals = row_values(cash, cash_row, start_col, end_col)
    add_vals = row_values(funding, add_row, start_col, end_col)
    old_vals = row_values(funding, old_row, start_col, end_col)
    cash_bad = limit_bad = 0
    for i, month in enumerate(months):
        if month <= last_actual:
            continue
        if i < len(cash_vals) and cash_vals[i] is not None and float(cash_vals[i]) * unit + 1 < minimum:
            cash_bad += 1
            if cash_bad <= 6:
                failures.append(f"{month} 기말 현금이 최소 현금보다 작다")
        if i >= len(add_vals) or i >= len(old_vals):
            continue
        used = float(add_vals[i] or 0) + float(old_vals[i] or 0)
        if used * unit > limit + 1:
            limit_bad += 1
            if limit_bad <= 6:
                failures.append(f"{month} 한도대출 잔액이 한도를 넘는다")
    lines.append(f"- 추정 최소 현금 ({min_where}): {'실패 ' + str(cash_bad) + '개월' if cash_bad else '통과'}")
    lines.append(f"- 추정 한도대출 ({limit_where}): {'실패 ' + str(limit_bad) + '개월' if limit_bad else '통과'}")
    return lines, continuity


def names_for_box(book, sheet, box):
    found = []
    for name in book.defined_names:
        item = book.defined_names[name]
        text = item.attr_text or ""
        if "!" not in text:
            continue
        owner, addr = text.split("!", 1)
        owner = owner.strip("=").strip("'").replace("''", "'")
        if owner != sheet:
            continue
        try:
            other = parse_box(addr)
        except Exception:
            continue
        if not (other[3] < box[1] or other[1] > box[3] or other[2] < box[0] or other[0] > box[2]):
            found.append(name)
    return found


LOCAL_REF = re.compile(r"(?<![A-Za-z0-9_!])\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?", re.I)


def local_ref_hits(formula, box):
    from openpyxl.utils import column_index_from_string
    for match in LOCAL_REF.finditer(formula):
        c1 = column_index_from_string(match.group(1))
        r1 = int(match.group(2))
        if match.group(3):
            c2 = column_index_from_string(match.group(3))
            r2 = int(match.group(4))
            if not (max(c1, c2) < box[1] or min(c1, c2) > box[3] or max(r1, r2) < box[0] or min(r1, r2) > box[2]):
                return True
        elif covers([box], r1, c1):
            return True
    return False


def input_warnings(path, meta, scan_edges):
    del scan_edges
    warnings = []
    from openpyxl import load_workbook
    book = load_workbook(path, data_only=False, read_only=False)
    try:
        owned = []
        for ws in book.worksheets:
            for cell in ws._cells.values():
                text = cell.value.text if hasattr(cell.value, "text") else cell.value
                if isinstance(text, str) and text.startswith("="):
                    owned.append((ws.title, text))
        blob = "\n".join(text for _owner, text in owned)
        for item in meta.get("inputs") or []:
            sheet = str(item.get("sheet"))
            box = parse_box(item.get("cell"))
            hit = False
            for owner, text in owned:
                if sheet_ref_hits(text, sheet, box) or (owner == sheet and local_ref_hits(text, box)):
                    hit = True
                    break
            if not hit:
                for name in names_for_box(book, sheet, box):
                    if re.search(r"(?<![A-Za-z0-9_])" + re.escape(name) + r"(?![A-Za-z0-9_])", blob):
                        hit = True
                        break
            if not hit:
                warnings.append(f"{sheet}!{item.get('cell')} {item.get('label')}")
    finally:
        book.close()
    return warnings


def sheet_ref_hits(formula, sheet, box):
    from workbook_scan import REF
    from openpyxl.utils import column_index_from_string
    for ref in REF.finditer(formula):
        target = (ref.group("quoted") or ref.group("bare") or "").replace("''", "'")
        if target != sheet:
            continue
        c1 = column_index_from_string(ref.group("c1").replace("$", ""))
        r1 = int(ref.group("r1").replace("$", ""))
        if ref.group("c2"):
            c2 = column_index_from_string(ref.group("c2").replace("$", ""))
            r2 = int(ref.group("r2").replace("$", ""))
            if not (max(c1, c2) < box[1] or min(c1, c2) > box[3] or max(r1, r2) < box[0] or min(r1, r2) > box[2]):
                return True
        elif covers([box], r1, c1):
            return True
    return False


def constant_kind(sheet, value):
    number = float(value)
    whole = abs(number - round(number)) < 1e-9
    if sheet == "검증":
        return "source"
    if whole and 1990 <= round(number) <= 2100:
        return "year"
    if sheet == "Sensitivity" and whole and 1 <= round(number) <= 30:
        return "case"
    return "other"


def constant_report(constants):
    labels = (
        ("case", "케이스 번호"),
        ("year", "연도"),
        ("source", "검증 원천·기준값"),
    )
    buckets = {key: [] for key, _label in labels}
    buckets["other"] = []
    for sheet, coord, value in constants:
        buckets[constant_kind(sheet, value)].append((sheet, coord, value))
    lines = [
        "입력으로 등록되지 않은 숫자 상수다. 케이스 번호, 연도, 검증 시트의 원천·기준값은 모델이 일부러 둔 값이라 건수만 센다. 그 밖에 남은 숫자만 주소를 적는다.",
        "",
    ]
    for key, label in labels:
        lines.append(f"- {label}: {len(buckets[key])}칸")
    others = buckets["other"]
    if not others:
        lines.append("- 그 밖: 0칸. 통과")
        return lines
    lines.append(f"- 그 밖: {len(others)}칸")
    for sheet, coord, value in others[:40]:
        lines.append(f"  - {sheet}!{coord} = {value}")
    if len(others) > 40:
        lines.append(f"  - 나머지 {len(others) - 40}칸은 생략")
    return lines


# Marginal rates from tax_rules_kr.md, low bracket to high. Local small-company rates are one tenth of the national small-company rates.
TAX_RATES = {
    "2025": {
        "corp_general": (0.09, 0.19, 0.21, 0.24),
        "corp_small": (0.19, 0.19, 0.21, 0.24),
        "local_general": (0.009, 0.019, 0.021, 0.024),
        "local_small": (0.019, 0.019, 0.021, 0.024),
    },
    "2026": {
        "corp_general": (0.10, 0.20, 0.22, 0.25),
        "corp_small": (0.20, 0.20, 0.22, 0.25),
        "local_general": (0.01, 0.02, 0.022, 0.025),
        "local_small": (0.02, 0.02, 0.022, 0.025),
    },
}


def tax_row_kind(text):
    if "법인세" in text and "지방" not in text and "성실" not in text:
        return "corp_general"
    if "법인세" in text and "성실" in text:
        return "corp_small"
    if "지방" in text and "성실" not in text:
        return "local_general"
    if "지방" in text and "성실" in text:
        return "local_small"
    return None


def tax_table_lines(book):
    ws = sheet_or_none(book, "전제사항")
    if ws is None:
        return ["세율표 확인 못 함"], ["- 세율표 확인 못 함"]
    labels = sheet_labels(ws)
    warnings = []
    lines = ["세율표의 구간 세율을 기준서 기본값과 비교한다. 다르면 경고이며 실패로 세지 않는다."]
    found = False
    for row, text in enumerate(labels, 1):
        if "세율표" not in text:
            continue
        era = "2026" if "2026" in text else "2025" if "2025" in text else None
        if era is None:
            continue
        found = True
        for follow in range(row + 1, row + 8):
            if follow - 1 >= len(labels):
                break
            label = labels[follow - 1]
            if "세율표" in label:
                break
            kind = tax_row_kind(label)
            if kind is None:
                continue
            rates = []
            for value in row_values(ws, follow, 11, 16):
                if isinstance(value, (int, float)) and not isinstance(value, bool):
                    number = float(value)
                    if 1 < number <= 30:
                        number = number / 100
                    if 0 < number <= 0.30:
                        rates.append(number)
            expected = TAX_RATES[era][kind]
            if tuple(round(rate, 6) for rate in rates[:4]) != tuple(round(rate, 6) for rate in expected):
                warnings.append(f"세율표가 기본값과 다름 ({era} {label})")
                lines.append(f"- {era} {label}: 모델 {rates[:4]}, 기준서 {list(expected)}")
            else:
                lines.append(f"- {era} {label}: 기준서 기본 세율과 같다")
    if not found:
        return ["세율표 확인 못 함"], ["- 세율표 확인 못 함"]
    if not warnings:
        lines.append("- 읽은 세율은 기준서 기본값과 같다")
    return warnings, lines


def effective_tax_lines(book, meta):
    info = meta.get("info") or {}
    if "first_month" not in info or "plan_end" not in info:
        return ["- 실효세율 확인 못 함"]
    income = sheet_or_none(book, "IS")
    if income is None:
        return ["- 실효세율 확인 못 함"]
    labels = sheet_labels(income)
    tax_row, _tax = find_labeled_row(labels, ("법인세", "법인세비용"))
    pretax_row, _pretax = find_labeled_row(labels, ("세전이익",))
    if not tax_row or not pretax_row:
        return ["- 실효세율 확인 못 함"]
    months = months_of(info["first_month"], info["plan_end"])
    start_col = int(float(info["first_month_col"]))
    end_col = int(float(info["last_month_col"]))
    tax_vals = row_values(income, tax_row, start_col, end_col)
    pretax_vals = row_values(income, pretax_row, start_col, end_col)
    lines = ["| 연도 | 법인세 | 세전이익 | 실효세율 |", "|---|---:|---:|---:|"]
    years = []
    for month in months:
        if month[:4] not in years:
            years.append(month[:4])
    for year in years:
        tax = pretax = 0.0
        for index, month in enumerate(months):
            if not month.startswith(year):
                continue
            if index < len(tax_vals) and isinstance(tax_vals[index], (int, float)):
                tax += float(tax_vals[index])
            if index < len(pretax_vals) and isinstance(pretax_vals[index], (int, float)):
                pretax += float(pretax_vals[index])
        rate = f"{tax / pretax * 100:.1f}%" if pretax > 0 else "확인 못 함"
        lines.append(f"| {year} | {tax:,.1f} | {pretax:,.1f} | {rate} |")
    lines.append("")
    lines.append("금액은 모델 표시 단위다. 모델은 이 비율이 아니라 법인세법 세율표로 세액을 계산한다.")
    return lines


def balance_jump_lines(book, meta):
    info = meta.get("info") or {}
    if "last_actual" not in info or "first_month" not in info:
        return [], ["- 확인 못 함"]
    months = months_of(info["first_month"], info["plan_end"])
    if info["last_actual"] not in months:
        return [], ["- 실적 마감 월이 모델 구간에 없다"]
    last_index = months.index(str(info["last_actual"]))
    if last_index + 1 >= len(months):
        return [], ["- 추정 월이 없다"]
    start_col = int(float(info["first_month_col"]))
    end_col = int(float(info["last_month_col"]))
    sheet = None
    for name in ("BS명세", "CF", "BS"):
        candidate = sheet_or_none(book, name)
        if candidate is None:
            continue
        if find_labeled_row(sheet_labels(candidate), ("매출채권",))[0]:
            sheet = candidate
            break
    if sheet is None:
        sheet = sheet_or_none(book, "CF") or sheet_or_none(book, "BS")
    if sheet is None:
        return [], ["- BS와 CF가 없어 확인 못 함"]
    labels = sheet_labels(sheet)
    warnings = []
    lines = [f"실적 마감 {info['last_actual']}과 추정 첫 달 {months[last_index + 1]}을 비교한다. 20%를 넘으면 경고이며 실패로 세지 않는다."]
    for name in ("매출채권", "재고자산", "매입채무"):
        row, _text = find_labeled_row(labels, (name,))
        if not row:
            lines.append(f"- {name}: 행을 찾지 못했다")
            continue
        vals = row_values(sheet, row, start_col, end_col)
        actual = numeric_at(vals, last_index)
        forecast = numeric_at(vals, last_index + 1)
        if actual is None or forecast is None:
            lines.append(f"- {name}: 값이 비어 확인 못 함")
            continue
        if abs(actual) < 1e-9:
            lines.append(f"- {name}: 실적 마감 {actual:.4g}, 추정 첫 달 {forecast:.4g}")
            if abs(forecast) >= 1e-9:
                warnings.append(f"{name} 잔액 단절")
            continue
        gap = abs(forecast / actual - 1)
        lines.append(f"- {name}: 실적 마감 {actual:.4g}, 추정 첫 달 {forecast:.4g}, 차이 {gap * 100:.1f}%")
        if gap > 0.20:
            warnings.append(f"{name} 잔액이 20%를 넘음")
    return warnings, lines


def plan_year_column(result, plan_end):
    year = int(str(plan_end)[:4])
    section, _ = find_labeled_row(sheet_labels(result), ("연도별 손익 및 자금",))
    if section is None:
        return None
    last_col = max(4, result.UsedRange.Column + result.UsedRange.Columns.Count - 1)
    headers = row_values(result, section + 1, 3, last_col)
    kinds = row_values(result, section + 2, 3, last_col)
    for offset, value in enumerate(headers):
        if isinstance(value, bool) or not isinstance(value, (int, float)):
            continue
        if int(round(float(value))) != year:
            continue
        kind = kinds[offset] if offset < len(kinds) else ""
        if isinstance(kind, str) and "대비" in kind:
            continue
        return 3 + offset
    return None


def dscr_lines(book, meta):
    info = meta.get("info") or {}
    result = sheet_or_none(book, "민감도및결과")
    if result is None or "plan_end" not in info:
        return [], ["- DSCR 확인 못 함"]
    labels = sheet_labels(result)
    row, _text = find_labeled_row(labels, ("DSCR",))
    if not row:
        return [], ["- DSCR 행을 찾지 못했다"]
    column = plan_year_column(result, info["plan_end"])
    if column is None:
        return [], ["- DSCR 연도 열을 찾지 못했다"]
    value = result.Range(f"{letter(column)}{row}").Value2
    if not isinstance(value, (int, float)) or isinstance(value, bool):
        return [], [f"- 계획 구간 DSCR {value}"]
    plan = float(value)
    lines = [f"- 계획 구간 DSCR {plan:.2f} ({info['plan_end'][:4]}년 열)"]
    warnings = []
    if plan < 1:
        warnings.append("DSCR 1 미만")
        balloon = maturity_repayment(book, info)
        lines.append(f"- 원리금 중 만기 일시상환 합계(추정 월, 표시 단위) {balloon}")
    return warnings, lines


def maturity_repayment(book, info):
    fund = sheet_or_none(book, "재원조달")
    if fund is None or "first_month" not in info:
        return "확인 못 함"
    months = months_of(info["first_month"], info["plan_end"])
    start_col = int(float(info["first_month_col"]))
    end_col = int(float(info["last_month_col"]))
    total = 0.0
    found = False
    for row, text in enumerate(sheet_labels(fund), 1):
        if "만기" not in text or "상환" not in text:
            continue
        vals = row_values(fund, row, start_col, end_col)
        for index, month in enumerate(months):
            if month <= str(info["last_actual"]):
                continue
            value = numeric_at(vals, index)
            if value is not None:
                total += value
                found = True
    return f"{total:.4g}" if found else "행을 찾지 못했다"


def render(path, data, meta, formula, failures, notes, warnings, continuity, case_text, tax_lines, balance_lines, dscr_report):
    warn_items = list(continuity or []) + list(formula.get("extra_warnings") or []) + list(warnings or [])
    lines = ["# 모델 검증", "", f"파일: `{path}`", f"자료: `{data}`", ""]
    lines.append("결과: " + ("실패" if failures else "통과"))
    lines.append(f"실패 {len(failures)}건, 경고 {len(warn_items)}건")
    if warn_items:
        for item in warn_items:
            lines.append(f"- {item}")
    else:
        lines.append("- 경고 없음")
    lines.append("")
    lines.append("## META_Checks")
    lines.append("")
    for note in notes:
        lines.append(f"- {note}")
    lines.append("")
    lines.append("## 금지 수식·오류·순환")
    lines.append("")
    for key, label in (
        ("offset", "OFFSET"), ("indirect", "INDIRECT"), ("array", "배열 수식"),
        ("external", "외부 링크"), ("ref", "#REF!"),
    ):
        hits = formula["forbidden"][key]
        lines.append(f"- {label}: {len(hits)}건" + ("" if not hits else " (" + ", ".join(hits[:8]) + ")"))
    lines.append("")
    lines.append("## 실적·추정")
    lines.append("")
    lines.extend(formula.get("tie_lines") or [])
    lines.append("")
    lines.append("## 실적과 추정의 이어짐")
    lines.append("")
    lines.append("추정 첫 12개월이 있으면 그 합계를, 없으면 있는 만큼의 합계를 직전 같은 개월 수 합계와 비교한다. 추정 첫 3개월은 전년 같은 3개월과 비교한다. 25%를 넘으면 경고이며 실패로 세지 않는다.")
    lines.append("")
    if continuity:
        for item in continuity:
            lines.append(f"- {item}")
    else:
        lines.append("- 25%를 넘는 단절 없음")
    lines.append("")
    lines.append("## 세금")
    lines.append("")
    lines.extend(tax_lines or ["- 확인 못 함"])
    lines.append("")
    lines.append("## 운전자본 잔액")
    lines.append("")
    lines.extend(balance_lines or ["- 확인 못 함"])
    lines.append("")
    lines.append("## DSCR")
    lines.append("")
    lines.extend(dscr_report or ["- 확인 못 함"])
    lines.append("")
    lines.append("## 민감도 저장값")
    lines.append("")
    lines.append(f"- {case_text}")
    lines.append("")
    lines.append("## 계산 영역 하드코딩 숫자")
    lines.append("")
    lines.extend(constant_report(formula.get("constants") or []))
    lines.append("")
    lines.append("## 쓰이지 않는 입력")
    lines.append("")
    if warnings:
        lines.append(f"- 경고 {len(warnings)}건. 실패로 세지 않는다.")
        for item in warnings[:30]:
            lines.append(f"  - {item}")
    else:
        lines.append("- 없음")
    lines.append("")
    if failures:
        lines.append("## 실패")
        lines.append("")
        for item in failures:
            lines.append(f"- {item}")
        lines.append("")
    return "\n".join(lines) + "\n"


@dataclass
class ExcelChecks:
    failures: list = field(default_factory=list)
    notes: list = field(default_factory=list)
    tie_lines: list = field(default_factory=list)
    continuity: list = field(default_factory=list)
    extra_warnings: list = field(default_factory=list)
    tax_lines: list = field(default_factory=lambda: ["- 확인 못 함"])
    balance_lines: list = field(default_factory=lambda: ["- 확인 못 함"])
    dscr_lines: list = field(default_factory=lambda: ["- 확인 못 함"])
    case_text: str = "확인 못 함"


def check_excel_session(path, data, meta, result):
    excel, book = open_excel(path)
    with cleanup_excel(lambda: close_excel(excel, book)):
        if not meta.get("present"):
            result.failures.append("모델정보 또는 META_Info가 없다")
        else:
            table = book.Names("META_Checks").RefersToRange.Value2
            rows = table if isinstance(table[0], tuple) else (table,)
            for row in rows[1:]:
                name, sheet, cell, expected = row[:4]
                got = book.Worksheets(sheet).Range(cell).Value2
                if not (got == expected or same(got, expected)):
                    result.failures.append(f"META_Checks {name}: {got} (기대 {expected})")
                result.notes.append(f"{name}: {got}")
            try:
                circ = excel.CircularReference
            except Exception as exc:
                if busy(exc):
                    raise
                circ = None
            if circ is not None:
                result.failures.append("순환 참조 " + str(circ.Address))
                result.notes.append("순환 참조 있음")
            else:
                result.notes.append("순환 참조 없음")
            errors = com_errors(book)
            if errors:
                result.failures.append("오류 값 " + ", ".join(errors[:8]))
            else:
                result.notes.append("계산된 오류 값 없음")
            result.tie_lines, result.continuity = tie_out(book, meta, data, result.failures)
            tax_warns, result.tax_lines = tax_table_lines(book)
            result.tax_lines = result.tax_lines + [""] + effective_tax_lines(book, meta)
            balance_warns, result.balance_lines = balance_jump_lines(book, meta)
            dscr_warns, result.dscr_lines = dscr_lines(book, meta)
            result.extra_warnings = tax_warns + balance_warns + dscr_warns
            result.case_text = check_cases(excel, book, result.failures)
            if "ACT_Lines" in meta["defined"]:
                from activity_verify import verify_activity
                result.notes.extend(verify_activity(excel, book, meta, data, result.failures))


def verify_in_excel(path, data, meta):
    result = None

    def attempt():
        nonlocal result
        # A fresh read-only session also needs fresh diagnostics after an interrupted check.
        result = ExcelChecks()
        check_excel_session(path, data, meta, result)

    try:
        retry_excel(attempt, attempts=3)
    except Exception as exc:
        if busy(exc) or isinstance(exc, ExcelShutdownBusy):
            result.failures.append(BUSY_MESSAGE)
        else:
            import traceback
            tail = " | ".join(traceback.format_exc().strip().splitlines()[-4:])
            result.failures.append("Excel 재계산 실패: " + tail[-500:])
    return result


def main():
    parser = argparse.ArgumentParser(description="Verify a calculated business-plan workbook.")
    parser.add_argument("workbook", type=Path)
    parser.add_argument("--data", type=Path)
    parser.add_argument("--out", type=Path)
    parser.add_argument("--spec", type=Path)
    parser.add_argument("--sources", nargs="+", type=Path, help="합산 파일에 저장된 회사별 값을 원본 모델과 대조")
    args = parser.parse_args()
    if not args.workbook.exists():
        say("FAIL workbook missing")
        return 2
    meta = load_meta(args.workbook)
    if meta["info"].get("kind") == "combined":
        from combine_verify import verify_combined
        return verify_combined(args.workbook, meta, args.sources, args.out)
    if args.sources or args.data is None or args.out is None:
        parser.error("법인별 모델 검증에는 --data 및 --out이 필요합니다. --sources는 합산용입니다.")
    if not (args.data / "is_monthly.csv").exists():
        say("FAIL is_monthly.csv missing")
        return 2
    failures = []
    stale = model_is_stale(meta, args.spec)
    if stale:
        failures.append(stale)
    formula = scan_formulas(args.workbook)
    for key in ("offset", "indirect", "array", "external", "ref"):
        if formula["forbidden"][key]:
            failures.append(f"{key} {len(formula['forbidden'][key])}")
    formula["constants"] = constant_cells(args.workbook, meta) if meta.get("present") else []
    result = verify_in_excel(args.workbook, args.data, meta)
    failures.extend(result.failures)
    notes = result.notes
    formula["tie_lines"] = result.tie_lines
    continuity = result.continuity
    tax_lines = result.tax_lines
    balance_lines = result.balance_lines
    dscr_report = result.dscr_lines
    extra_warnings = result.extra_warnings
    case_text = result.case_text
    warnings = input_warnings(args.workbook, meta, formula["edges"]) if meta.get("present") else []
    palette = inspect_palette(args.workbook)
    palette_warnings = warnings_for(palette)
    font_warnings = font_warnings_for(inspect_fonts(args.workbook))
    extra_warnings.extend(palette_warnings + font_warnings)
    formula["extra_warnings"] = extra_warnings
    notes.append(f"통합문서 색상 검사: 경고 {len(palette_warnings)}건")
    notes.append(f"통합문서 글꼴 검사: 경고 {len(font_warnings)}건")
    write_text(args.out, render(
        args.workbook, args.data, meta, formula, failures, notes, warnings, continuity, case_text,
        tax_lines, balance_lines, dscr_report,
    ))
    say(f"FAILS {len(failures)}")
    say(f"WARNS {len(warnings) + len(continuity) + len(extra_warnings)}")
    wrote(args.out, "verify")
    say("RESULT " + ("FAIL" if failures else "PASS"))
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
