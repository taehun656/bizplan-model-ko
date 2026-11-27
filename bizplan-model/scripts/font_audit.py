"""Report typefaces used by worksheet styles, rich text and linked Excel drawings."""
import re
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from openpyxl.utils import get_column_letter

from design_palette import ALLOWED_FONTS, FONT
from palette_audit import RID, child, elements, relationships, tag


def _themes(z):
    names = {}
    path = next((p for p in z.namelist() if p.startswith("xl/theme/theme") and p.endswith(".xml")), None)
    if path:
        root = ET.fromstring(z.read(path))
        for group in ("major", "minor"):
            for entry in elements(root, group + "Font"):
                for node in entry:
                    role = node.get("script", tag(node))
                    if node.get("typeface"):
                        names[group, role] = node.get("typeface")
    return names


def _theme_face(themes, scheme, script):
    if script == "ea":
        return themes.get((scheme, "Hang"), themes.get((scheme, "ea"), "테마 글꼴 확인 불가"))
    return themes.get((scheme, script), themes.get((scheme, "latin"), "테마 글꼴 확인 불가"))


def _cell_faces(font, value, themes):
    """Excel can preserve both a literal family and an effective themed typeface."""
    if font is None:
        return
    face = child(font, "name")
    if face is None:
        face = child(font, "rFont")
    if face is not None and face.get("val"):
        yield face.get("val"), "글꼴 이름"
    scheme = child(font, "scheme")
    if scheme is None or scheme.get("val") not in ("major", "minor"):
        return
    scripts = set()
    if value is None or re.search(r"[\x21-\x7E]", value):
        scripts.add("latin")
    if value is None or re.search(r"[\u1100-\u11FF\u3000-\u9FFF\uAC00-\uD7A3]", value):
        scripts.add("ea")
    for script in sorted(scripts):
        key = scheme.get("val")
        yield _theme_face(themes, key, script), f"테마 {key}/{script}"


def _style_fonts(styles):
    xfs = list(child(styles, "cellStyleXfs") or []) or [ET.Element("xf")]
    fonts = list(child(styles, "fonts") or []) or [ET.Element("font")]
    normal_base = next((x for x in child(styles, "cellStyles") or []
                        if x.get("builtinId") == "0"), None)
    normal_xf = xfs[int(normal_base.get("xfId", 0))] if normal_base is not None else xfs[0]
    styles_fonts = []
    for xf in child(styles, "cellXfs") or [ET.Element("xf")]:
        parent = xfs[int(xf.get("xfId", 0))]
        index = parent.get("fontId", "0") if xf.get("applyFont") == "0" else xf.get("fontId", parent.get("fontId","0"))
        styles_fonts.append(fonts[int(index)])
    return fonts[int(normal_xf.get("fontId",0))], styles_fonts


def _rich_runs(item):
    return [(child(run, "rPr"), child(run, "t").text or "")
            for run in item.findall("{*}r") if child(run, "t") is not None]


def inspect_fonts(path):
    """Return per-sheet font groups with role and location for every warning."""
    groups = {}
    def add(sheet, location, role, face, source):
        issue = "지정 글꼴 밖" if face not in ALLOWED_FONTS else ""
        if role == "Normal" and face != FONT:
            issue = "Normal 글꼴"
        key = sheet, role, face, source, issue
        if key not in groups:
            groups[key] = dict(sheet=sheet, purpose=role, font=face, source=source,
                               issue=issue, count=0, locations=[])
        row = groups[key]
        row["count"] += 1
        if issue or len(row["locations"]) < 3:
            row["locations"].append(location)
    with ZipFile(path) as z:
        themes = _themes(z)
        styles = ET.fromstring(z.read("xl/styles.xml")) if "xl/styles.xml" in z.namelist() else ET.Element("styles")
        normal, style_fonts = _style_fonts(styles)
        for face, source in _cell_faces(normal, None, themes):
            add("통합문서", "Normal", "Normal", face, source)
        diffs = list(child(styles,"dxfs") or [])
        strings = []
        if "xl/sharedStrings.xml" in z.namelist():
            for item in ET.fromstring(z.read("xl/sharedStrings.xml")):
                strings.append(("".join(e.text or "" for e in elements(item,"t")), _rich_runs(item)))
        wb = "xl/workbook.xml"
        sheet_paths = relationships(z,wb)
        for sheet in elements(ET.fromstring(z.read(wb)), "sheet"):
            name = sheet.attrib["name"]
            part = sheet_paths[sheet.attrib[RID]]
            root = ET.fromstring(z.read(part))
            links = relationships(z, part)
            def check(location, role, font, value=None):
                for face, source in _cell_faces(font, value, themes):
                    add(name, location, role, face, source)
            for col in elements(root, "col"):
                if col.get("style"):
                    loc = f"{get_column_letter(int(col.get('min')))}:{get_column_letter(int(col.get('max')))}"
                    check(loc, "열 서식", style_fonts[int(col.get("style"))])
            for row in elements(root,"row"):
                if row.get("s"):
                    check(row.get("r")+":"+row.get("r"), "행 서식", style_fonts[int(row.get("s"))])
                for cell in row.findall("{*}c"):
                    loc = cell.attrib["r"]
                    value_node = child(cell, "v")
                    value = value_node.text if value_node is not None else ""
                    runs = []
                    if cell.get("t") == "s":
                        value, runs = strings[int(value)]
                    elif cell.get("t") == "inlineStr" and child(cell,"is") is not None:
                        item = child(cell,"is")
                        value, runs = "".join(e.text or "" for e in elements(item,"t")), _rich_runs(item)
                    check(loc, "셀", style_fonts[int(cell.get("s","0"))], value)
                    for font, text in runs:
                        check(loc, "서식 문자열", font, text)
            for rule_block in elements(root, "conditionalFormatting"):
                sq = child(rule_block, "sqref")
                loc = rule_block.get("sqref",sq.text if sq is not None else "조건부 서식")
                for rule in rule_block.findall("{*}cfRule"):
                    if rule.get("dxfId") is not None:
                        check(loc, "조건부 서식", child(diffs[int(rule.get("dxfId"))],"font"))
                    for inline in elements(rule,"dxf"):
                        check(loc,"조건부 서식",child(inline,"font"))
            for ref in root.iter():
                if tag(ref) not in ("drawing","legacyDrawing","legacyDrawingHF") or ref.get(RID) not in links:
                    continue
                drawing = links[ref.get(RID)]
                dr = ET.fromstring(z.read(drawing))
                if drawing.endswith(".vml"):
                    for shape in elements(dr,"shape"):
                        for node in elements(shape,"font"):
                            if node.get("face"):
                                add(name, shape.get("id",drawing), "VML 도형", node.get("face"), "VML face")
                    continue
                draw_links = relationships(z,drawing)
                for anchor in dr.iter():
                    if tag(anchor) not in ("oneCellAnchor","twoCellAnchor","absoluteAnchor"):
                        continue
                    obj = next(iter(elements(anchor,"cNvPr")),None)
                    loc = obj.get("name",drawing) if obj is not None else drawing
                    parts = [("도형", anchor)]
                    for chart in elements(anchor, "chart"):
                        target = draw_links.get(chart.get(RID))
                        if target:
                            parts.append(("차트", ET.fromstring(z.read(target))))
                    for role, content in parts:
                        for node in content.iter():
                            if not node.get("typeface"):
                                continue
                            face = node.get("typeface")
                            theme = re.fullmatch(r"\+(mj|mn)-(lt|ea|cs)", face)
                            if theme:
                                scheme = "major" if theme[1]=="mj" else "minor"
                                script = "latin" if theme[2]=="lt" else theme[2]
                                face = _theme_face(themes,scheme,script)
                            add(name, loc, role+"/"+tag(node), face, node.get("typeface"))
    return list(groups.values())


def font_warnings_for(rows):
    return [f"글꼴 검사 {r['issue']} · {r['sheet']}!{r['locations'][0]} · "
            f"{r['purpose']} · {r['font']} ({r['source']}, {r['count']}곳)"
            for r in rows if r["issue"]]
