# -*- coding: utf-8 -*-
"""Read model metadata and cross-sheet formulas without Excel."""
import re
from datetime import datetime
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string, range_boundaries
from openpyxl.worksheet.formula import ArrayFormula

REF = re.compile(
    r"(?:'(?P<quoted>(?:[^']|'')+)'|(?P<bare>[A-Za-z0-9_\uac00-\ud7a3]+))!"
    r"(?P<c1>\$?[A-Z]{1,3})(?P<r1>\$?\d+)(?::(?P<c2>\$?[A-Z]{1,3})(?P<r2>\$?\d+))?",
    re.I,
)
FORBIDDEN = re.compile(r"(?<![A-Za-z0-9_])(OFFSET|INDIRECT)\s*\(", re.I)
EXTERNAL = re.compile(r"\[[^\]]+\]")


def formula_text(value):
    if isinstance(value, ArrayFormula):
        return value.text or "", True
    if isinstance(value, str) and value.startswith("="):
        return value, False
    return "", False


def parse_box(addr):
    addr = str(addr).replace("$", "")
    if "!" in addr:
        addr = addr.split("!")[-1]
    min_col, min_row, max_col, max_row = range_boundaries(addr)
    return min_row, min_col, max_row, max_col


def covers(boxes, row, col):
    return any(r1 <= row <= r2 and c1 <= col <= c2 for r1, c1, r2, c2 in boxes)


def load_meta(path):
    book = load_workbook(path, data_only=False, read_only=False)
    try:
        defined = {}
        for name in book.defined_names:
            item = book.defined_names[name]
            defined[name] = item.attr_text
        if "모델정보" not in book.sheetnames or "META_Info" not in defined:
            return {"defined": defined, "info": {}, "sheets": [], "inputs": [], "checks": [], "present": False}
        info = table(book, defined["META_Info"])
        return {
            "defined": defined,
            "present": True,
            "info": {row[0]: row[1] for row in info[1:] if row and row[0]},
            "sheets": records(table(book, defined["META_Sheets"])),
            "inputs": records(table(book, defined["META_Inputs"])),
            "checks": records(table(book, defined["META_Checks"])),
        }
    finally:
        book.close()


def table(book, refers):
    sheet, addr = refers.split("!")
    sheet = sheet.strip("=").strip("'").replace("''", "'")
    ws = book[sheet]
    min_col, min_row, max_col, max_row = range_boundaries(addr.replace("$", ""))
    rows = []
    for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col, values_only=True):
        rows.append(list(row))
    return rows


def records(rows):
    if not rows:
        return []
    header = [str(c).strip() if c is not None else "" for c in rows[0]]
    out = []
    for row in rows[1:]:
        item = {}
        for key, value in zip(header, row):
            item[key] = value
        if any(v not in (None, "") for v in item.values()):
            out.append(item)
    return out


def scan_formulas(path):
    book = load_workbook(path, data_only=False, read_only=False)
    forbidden = {"offset": [], "indirect": [], "array": [], "external": [], "ref": []}
    pair_cells = {}
    try:
        unit = unit_cell(book)
        for ws in book.worksheets:
            for row in ws.iter_rows():
                for cell in row:
                    text, is_array = formula_text(cell.value)
                    if not text:
                        continue
                    if is_array and len(forbidden["array"]) < 30:
                        forbidden["array"].append(f"{ws.title}!{cell.coordinate}")
                    if "#REF!" in text.upper() and len(forbidden["ref"]) < 30:
                        forbidden["ref"].append(f"{ws.title}!{cell.coordinate}")
                    if EXTERNAL.search(text) and len(forbidden["external"]) < 30:
                        forbidden["external"].append(f"{ws.title}!{cell.coordinate}")
                    match = FORBIDDEN.search(text)
                    if match:
                        key = match.group(1).lower()
                        if len(forbidden[key]) < 30:
                            forbidden[key].append(f"{ws.title}!{cell.coordinate}")
                    for ref in REF.finditer(text):
                        target = (ref.group("quoted") or ref.group("bare") or "").replace("''", "'")
                        if not target or target == ws.title or target.upper() == "INDEX":
                            continue
                        c1 = column_index_from_string(ref.group("c1").replace("$", ""))
                        r1 = int(ref.group("r1").replace("$", ""))
                        if ref.group("c2"):
                            c2 = column_index_from_string(ref.group("c2").replace("$", ""))
                            r2 = int(ref.group("r2").replace("$", ""))
                            wide = abs(c2 - c1) + abs(r2 - r1) > 0
                            cell_count = (abs(c2 - c1) + 1) * (abs(r2 - r1) + 1)
                        else:
                            wide = False
                            cell_count = 1
                            c2, r2 = c1, r1
                        if unit and target == unit[0] and not wide and (r1, c1) == (unit[1], unit[2]):
                            continue
                        bucket = pair_cells.setdefault((target, ws.title), {"cells": 0, "wide": False})
                        bucket["cells"] += cell_count
                        bucket["wide"] = bucket["wide"] or wide
    finally:
        book.close()
    edges = []
    for (src, dst), info in sorted(pair_cells.items()):
        if info["cells"] <= 1 and not info["wide"]:
            continue
        edges.append((src, dst, info["cells"]))
    return {"forbidden": forbidden, "edges": edges}


def unit_cell(book):
    if "Unit" not in book.defined_names:
        return None
    text = book.defined_names["Unit"].attr_text
    if "!" not in text:
        return None
    sheet, addr = text.split("!")
    sheet = sheet.strip("=").strip("'").replace("''", "'")
    addr = addr.replace("$", "")
    if ":" in addr:
        return None
    min_col, min_row, _max_col, _max_row = range_boundaries(addr)
    return sheet, min_row, min_col


def constant_cells(path, meta):
    boxes = {}
    for item in meta.get("inputs") or []:
        boxes.setdefault(str(item.get("sheet")), []).append(parse_box(item.get("cell")))
    book = load_workbook(path, data_only=False, read_only=False)
    found = []
    try:
        roles = {row["sheet"]: row.get("role") for row in meta.get("sheets") or []}
        # Raw activity observations are reconciled to their CSV by verify_activity.
        # ACT_Metrics and ACT_Effects store the row addresses used for that audit.
        stored_names = ["SN_Store", "SN_Case", "SN_LastRun", "BL_LastRun", "ACT_Metrics", "ACT_Effects"]
        stored_names.extend(name for name in book.defined_names
                            if name.startswith(("ACT_SalesSource_", "ACT_MetricSource_")))
        store_boxes = named_boxes(book, stored_names)
        for ws in book.worksheets:
            role = roles.get(ws.title)
            if role in ("actual", "admin", None) and meta.get("present"):
                if role in ("actual", "admin"):
                    continue
            if not meta.get("present") and ws.title.startswith("실적"):
                continue
            for cell in ws._cells.values():
                value = cell.value
                if isinstance(value, bool) or not isinstance(value, (int, float)):
                    continue
                if cell.row <= 6:
                    continue
                if covers(boxes.get(ws.title, []), cell.row, cell.column):
                    continue
                if covers(store_boxes.get(ws.title, []), cell.row, cell.column):
                    continue
                found.append((ws.title, cell.coordinate, value))
    finally:
        book.close()
    return found


def model_is_stale(meta, spec_path):
    """Return a note when META built_at is older than the spec file. None when they cannot be compared."""
    if spec_path is None:
        return None
    spec_path = Path(spec_path)
    if not spec_path.exists():
        return None
    built = (meta.get("info") or {}).get("built_at")
    if not built:
        return None
    text = str(built).strip()
    if text.endswith("Z"):
        text = text[:-1]
    try:
        stamp = datetime.fromisoformat(text[:19])
    except ValueError:
        return None
    spec_time = datetime.fromtimestamp(spec_path.stat().st_mtime).replace(microsecond=0)
    if stamp < spec_time:
        return (
            "낡은 모델 (built_at "
            + stamp.isoformat(sep=" ", timespec="seconds")
            + ", 설계서 "
            + spec_time.isoformat(sep=" ", timespec="seconds")
            + ")"
        )
    return None


def named_boxes(book, names):
    boxes = {}
    for name in names:
        if name not in book.defined_names:
            continue
        text = book.defined_names[name].attr_text
        if "!" not in text:
            continue
        sheet, addr = text.split("!")
        sheet = sheet.strip("=").strip("'").replace("''", "'")
        boxes.setdefault(sheet, []).append(parse_box(addr))
    return boxes
