# -*- coding: utf-8 -*-
"""Build a single-file model map from workbook metadata and formulas."""
import argparse
import json
import re
import sys
from collections import defaultdict
from datetime import date, datetime
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.formula import Tokenizer

sys.path.insert(0, str(Path(__file__).resolve().parent))
from common import say, write_text, wrote
from verify_model import LOCAL_REF, local_ref_hits
from workbook_scan import REF, load_meta, model_is_stale, parse_box, scan_formulas
from openpyxl.utils import column_index_from_string

TEMPLATE = Path(__file__).resolve().parent.parent / "assets" / "model_map_template.html"
COLS = [
    {"t": "입력", "s": "가정과 실적 원천", "c": "in"},
    {"t": "계산 ① 가정·기준", "s": "스위치와 월축", "c": "calc"},
    {"t": "계산 ② 동인별 상세", "s": "라인별 계산", "c": "calc"},
    {"t": "계산 ③ 본표 집계", "s": "수입, 비용, 투자, 조달", "c": "calc"},
    {"t": "산출", "s": "재무제표, 결과, 검증", "c": ""},
]
KNOWN = {
    "목차": (4, "admin"), "모델정보": (4, "admin"), "스타일 가이드": (4, "admin"), "기준값갱신": (4, "admin"),
    "Sensitivity": (1, "assumption"), "전제사항": (1, "assumption"), "가정상세": (1, "assumption"),
    "금융조건": (1, "assumption"), "INDEX": (1, "assumption"),
    "투자비": (3, "summary"), "재원조달": (3, "summary"), "운영수입": (3, "summary"), "운영비용": (3, "summary"),
    "BS명세": (3, "summary"), "인건비": (2, "driver"), "확장조정": (0, "input"),
    "IS": (4, "statement"), "BS": (4, "statement"), "CF": (4, "statement"),
    "민감도및결과": (4, "result"), "검증": (4, "check"),
}


def kind_of(role, stage):
    if role == "actual":
        return "act"
    if role in ("input",) or stage == 0:
        return "in"
    if stage == 4:
        return "out"
    return "calc"


def role_group(role):
    return {
        "input": "입력 시트", "assumption": "가정", "driver": "동인별 상세",
        "summary": "본표 집계", "statement": "재무제표", "result": "결과",
        "actual": "실적·원천", "check": "검증", "admin": "관리",
    }.get(role, "시트")


def show_value(value):
    if value is None or value == "":
        return ""
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, str):
        text = value.strip()
        if len(text) >= 19 and text[10:19] == " 00:00:00" and text[4] == "-" and text[7] == "-":
            return text[:10]
        return text
    if isinstance(value, float):
        return f"{value:,.4g}"
    return str(value)


SECTION_NAMES = {
    1: "기준", 2: "자금", 3: "운전자본", 4: "법인세", 5: "투자·상각",
    6: "인건비", 7: "매출 가정", 8: "비용 가정", 9: "시나리오",
}
SECTION_HEADER = re.compile(r"^(\d+)\.\s+")
MONTHLY_NOTE = re.compile(r"월별 입력:\s*(.+?)\s*시트\s*(\d+)\s*행")


def unit_label(unit):
    try:
        number = int(float(unit))
    except (TypeError, ValueError):
        return "단위 확인 못 함"
    names = {1: "원", 1000: "천원", 10000: "만원", 1000000: "백만원", 100000000: "억원"}
    if number in names:
        return "단위 " + names[number]
    return "단위 " + f"{number:,}" + "원"


def input_row(cell):
    head = str(cell or "").replace("$", "").split(":")[0]
    digits = "".join(ch for ch in head if ch.isdigit())
    return int(digits) if digits else None


def assumption_index(ws):
    mapping = {}
    current = None
    if ws is None:
        return mapping
    for row in range(1, (ws.max_row or 1) + 1):
        title = ws.cell(row, 2).value
        if isinstance(title, str):
            match = SECTION_HEADER.match(title.strip())
            if match:
                current = int(match.group(1))
        if not current:
            continue
        mapping[("전제사항", row)] = current
        note = ws.cell(row, 10).value
        if isinstance(note, str):
            found = MONTHLY_NOTE.search(note)
            if found:
                mapping[(found.group(1).strip(), int(found.group(2)))] = current
    return mapping


def section_number(item, mapping):
    sheet = str(item.get("sheet") or "")
    row = input_row(item.get("cell"))
    if row and (sheet, row) in mapping:
        return mapping[(sheet, row)]
    group = str(item.get("group") or "")
    if sheet == "Sensitivity" or group in ("시나리오", "민감도"):
        return 9
    if group == "단위":
        return 1
    if sheet == "금융조건":
        return 2
    return None


def input_text(item, cached):
    sheet = str(item.get("sheet") or "")
    cell = str(item.get("cell") or "")
    label = str(item.get("label") or "")
    ws = cached.get(sheet)
    if ws is None or not cell:
        return f"{label} = "
    parts = cell.replace("$", "").split(":")
    try:
        first = ws[parts[0]].value
        if len(parts) == 1:
            return f"{label} = {show_value(first)}"
        last = ws[parts[1]].value
        return f"{label} = {show_value(first)} … {show_value(last)}"
    except Exception:
        return f"{label} = "


ACTIVITY_GROUPS = ("사업부 계획", "활동 지표")


def group_inputs(meta, cached, layout):
    buckets = defaultdict(list)
    combined = meta.get("info", {}).get("kind") == "combined"
    for item in meta.get("inputs") or []:
        number = section_number(item, layout)
        name = item["group"] if combined or item["group"] in ACTIVITY_GROUPS else SECTION_NAMES.get(number, "기타 입력")
        buckets[name].append(item)
    order = [SECTION_NAMES[number] for number in range(1, 10) if SECTION_NAMES[number] in buckets]
    order.extend(name for name in ACTIVITY_GROUPS if name in buckets)
    if "기타 입력" in buckets:
        order.append("기타 입력")
    if combined:
        order = [name for name in ("회사별 저장값", "제거조정") if name in buckets]
    nodes = []
    for index, name in enumerate(order):
        items = buckets[name]
        shown = [input_text(item, cached) for item in items]
        nodes.append({
            "id": f"i{index}", "col": 0, "g": "입력 묶음", "k": "in",
            "t": name, "s": f"{len(items)}개 입력",
            "cells": "", "items": shown,
            "d": "이 구역의 칸을 수식에서 직접 참조하는 시트로 이어진다.",
            "boxes": [(str(item.get("sheet")), parse_box(item.get("cell"))) for item in items],
        })
    return nodes


def sheet_nodes(rows):
    nodes = []
    for i, row in enumerate(rows):
        if row.get("role") == "admin":
            continue
        stage = int(float(row.get("stage") or 0))
        stage = min(4, max(0, stage))
        name = str(row.get("sheet"))
        description = str(row.get("description") or "")
        nodes.append({
            "id": f"s{i}", "col": stage, "g": role_group(row.get("role")), "k": kind_of(row.get("role"), stage),
            "t": name, "s": description,
            "cells": "", "d": description, "sheet": name,
        })
    return nodes


def edges_for(nodes, formula_edges, formula_text_by_sheet):
    by_sheet = {node["sheet"]: node["id"] for node in nodes if node.get("sheet")}
    edges = []
    seen = set()
    for src, dst, _count in formula_edges:
        if src in by_sheet and dst in by_sheet:
            pair = (by_sheet[src], by_sheet[dst])
            if pair not in seen and pair[0] != pair[1]:
                seen.add(pair)
                edges.append([pair[0], pair[1]])
    for node in nodes:
        if "boxes" not in node:
            continue
        for sheet, formulas in formula_text_by_sheet.items():
            direct = any(ref_hits(text, source, box)
                         for text in formulas for source, box in node["boxes"])
            if not direct and node["t"] in ACTIVITY_GROUPS:
                local_boxes = [box for source, box in node["boxes"] if source == sheet]
                direct = bool(local_boxes) and any(
                    same_sheet_input(text, local_boxes) for text in formulas)
            if not direct:
                continue
            target = by_sheet.get(sheet)
            if target and target != node["id"]:
                pair = (node["id"], target)
                if pair not in seen:
                    seen.add(pair)
                    edges.append([node["id"], target])
    return edges


def same_sheet_input(formula, boxes):
    """Check unqualified A1 operands against inputs on the formula's own sheet."""
    for token in Tokenizer(formula).items:
        if token.type == "OPERAND" and token.subtype == "RANGE" and LOCAL_REF.fullmatch(token.value):
            if any(local_ref_hits(token.value, box) for box in boxes):
                return True
    return False


def ref_hits(formula, sheet, box):
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
        elif box[0] <= r1 <= box[2] and box[1] <= c1 <= box[3]:
            return True
    return False


def formula_texts(path):
    book = load_workbook(path, data_only=False, read_only=False)
    found = defaultdict(list)
    try:
        for ws in book.worksheets:
            for cell in ws._cells.values():
                text = cell.value.text if hasattr(cell.value, "text") else cell.value
                if isinstance(text, str) and text.startswith("="):
                    found[ws.title].append(text)
    finally:
        book.close()
    return found


def guess_rows(path, formula_edges):
    book = load_workbook(path, read_only=True, data_only=False)
    names = list(book.sheetnames)
    book.close()
    incoming = defaultdict(set)
    for src, dst, _count in formula_edges:
        incoming[dst].add(src)
    stage = {}
    role = {}
    for name in names:
        if name.startswith("실적"):
            stage[name], role[name] = 0, "actual"
        elif name.startswith("매출"):
            stage[name], role[name] = 2, "driver"
        elif name in KNOWN:
            stage[name], role[name] = KNOWN[name]
        else:
            stage[name], role[name] = None, "driver"
    changed = True
    while changed:
        changed = False
        for name in names:
            if stage[name] is not None:
                continue
            preds = [stage[src] for src in incoming.get(name, ()) if stage.get(src) is not None]
            if preds:
                stage[name] = min(4, max(preds) + 1)
                changed = True
    for name in names:
        if stage[name] is None:
            stage[name] = 2
    return [{"sheet": name, "role": role[name], "stage": stage[name], "description": "이름과 수식 방향으로 추정"} for name in names]


def payload_for(path, guess):
    meta = load_meta(path)
    scanned = scan_formulas(path)
    texts = formula_texts(path)
    if guess or not meta.get("present"):
        rows = guess_rows(path, scanned["edges"])
        inputs = []
        title = Path(path).stem + " 모델 구조도"
        subtitle = "모델정보 없이 시트 이름과 수식 방향으로 배치"
    else:
        rows = meta["sheets"]
        inputs = meta["inputs"]
        company = meta["info"].get("company") or Path(path).stem
        title = str(company) + " 사업계획 모델 구조도"
        subtitle = f"{Path(path).name} · {unit_label(meta['info'].get('unit'))} · 실적 마감 {meta['info'].get('last_actual')} · 값은 저장 캐시"
    book = load_workbook(path, data_only=True, read_only=False)
    try:
        sheets = {ws.title: ws for ws in book.worksheets}
        nodes = group_inputs({"inputs": inputs, "info": meta["info"]}, sheets, assumption_index(sheets.get("전제사항")))
    finally:
        book.close()
    nodes.extend(sheet_nodes(rows))
    edges = edges_for(nodes, scanned["edges"], texts)
    note_common_unit(nodes, edges)
    public = []
    for node in nodes:
        public.append({k: v for k, v in node.items() if k not in ("boxes", "sheet") and v not in ("", [], None)})
    return {"title": title, "subtitle": subtitle, "cols": COLS, "nodes": public, "edges": edges, "_nodes": nodes}


def note_common_unit(nodes, edges):
    sheets = {node["id"]: node.get("sheet") or "" for node in nodes}
    for node in nodes:
        if node.get("t") != "기준":
            continue
        targets = [sheets.get(edge[1], "") for edge in edges if edge[0] == node["id"]]
        if targets and set(targets) <= {"검증"}:
            node["d"] = "단위는 전 시트 공통이다. 단위 칸은 한 칸 참조라 연결에서 빠지고, 이 묶음에 남은 연결은 검증이다."


def main():
    parser = argparse.ArgumentParser(description="Write a single-file model map.")
    parser.add_argument("workbook", type=Path)
    parser.add_argument("--out", required=True, type=Path)
    parser.add_argument("--guess", action="store_true")
    parser.add_argument("--spec", type=Path)
    args = parser.parse_args()
    if not args.workbook.exists():
        say("FAIL workbook missing")
        return 2
    meta = load_meta(args.workbook)
    if not meta.get("present") and not args.guess:
        say("FAIL model info missing")
        return 2
    built = payload_for(args.workbook, args.guess)
    if model_is_stale(meta, args.spec):
        built["subtitle"] = built.get("subtitle", "") + " · 낡은 모델"
        say("WARN stale model")
    data = {k: v for k, v in built.items() if not k.startswith("_")}
    raw = json.dumps(data, ensure_ascii=False).replace("<", "\\u003c")
    html = TEMPLATE.read_text(encoding="utf-8").replace("__TITLE__", data["title"]).replace("__PAYLOAD__", raw)
    if "http://" in html or "https://" in html:
        say("FAIL external link in html")
        return 1
    write_text(args.out, html)
    say(f"NODES {len(data['nodes'])}")
    say(f"EDGES {len(data['edges'])}")
    wrote(args.out, "map")
    say("RESULT PASS")
    return 0


if __name__ == "__main__":
    sys.exit(main())
