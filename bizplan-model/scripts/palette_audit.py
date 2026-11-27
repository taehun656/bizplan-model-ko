"""Audit colors actually referenced by the saved workbook's worksheets and drawings."""
import colorsys
import posixpath
import re
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from openpyxl.styles.colors import COLOR_INDEX
from openpyxl.styles.numbers import BUILTIN_FORMATS
from openpyxl.utils import get_column_letter

from design_palette import ALLOWED, BLACK, BLUE, GRAY, PEACH, RED, TAB_COLORS, WHITE, YELLOW


THEME_KEYS = ("lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3",
              "accent4", "accent5", "accent6", "hlink", "folHlink")
THEME_ALIAS = {"bg1": "lt1", "tx1": "dk1", "bg2": "lt2", "tx2": "dk2"}
# These are input encodings from the Excel file format, not workbook styling choices.
FORMAT_NAMES = {"black": BLACK, "blue": BLUE, "red": RED, "white": WHITE, "yellow": YELLOW,
                "cyan": "00FFFF", "green": "00FF00", "magenta": "FF00FF"}
DRAWING_COLORS = {"srgbClr", "schemeClr", "sysClr", "prstClr", "scrgbClr", "hslClr"}
RID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"


def tag(node):
    return node.tag.rsplit("}", 1)[-1]


def child(node, key):
    return node.find("{*}" + key) if node is not None else None


def elements(node, key):
    return node.findall(".//{*}" + key) if node is not None else []


def relationships(z, part):
    base, file = posixpath.split(part)
    key = posixpath.join(base, "_rels", file + ".rels")
    if key not in z.namelist():
        return {}
    rel = ET.fromstring(z.read(key))
    return {entry.attrib["Id"]: posixpath.normpath(posixpath.join(base, entry.attrib["Target"])).lstrip("/")
            for entry in rel if entry.get("TargetMode") != "External"}


def hex_rgb(channels):
    return "".join(f"{min(255, max(0, round(v * 255))):02X}" for v in channels)


def channels(rgb):
    return tuple(int(rgb[i:i+2], 16)/255 for i in (0, 2, 4))


def to_srgb(linear):
    return 12.92*linear if linear <= .0031308 else 1.055*linear**(1/2.4)-.055


def to_linear(srgb):
    return srgb/12.92 if srgb <= .04045 else ((srgb+.055)/1.055)**2.4


def theme_colors(z):
    name = next((p for p in z.namelist() if p.startswith("xl/theme/theme") and p.endswith(".xml")), None)
    if not name:
        return {}
    scheme = ET.fromstring(z.read(name)).find(".//{*}clrScheme")
    return {tag(entry): entry[0].get("lastClr", entry[0].get("val", "")).upper()
            for entry in scheme if len(entry)} if scheme is not None else {}


def sheet_color(node, theme, indexed):
    data = node.attrib
    if "rgb" in data:
        rgb, origin = data["rgb"][-6:].upper(), "RGB"
    elif "theme" in data:
        idx = int(data["theme"])
        key = THEME_KEYS[idx] if 0 <= idx < len(THEME_KEYS) else str(idx)
        rgb, origin = theme.get(key, "?"), "theme " + key
    elif "indexed" in data:
        idx = int(data["indexed"])
        rgb = indexed[idx] if idx < len(indexed) else {64: BLACK, 65: WHITE}.get(idx, "?")
        origin = "indexed " + str(idx)
    else:
        return BLACK, "automatic"
    if "tint" in data and re.fullmatch("[0-9A-F]{6}", rgb):
        tint = float(data["tint"])
        h, l, s = colorsys.rgb_to_hls(*channels(rgb))
        l = l*(1+tint) if tint < 0 else l*(1-tint)+tint
        rgb, origin = hex_rgb(colorsys.hls_to_rgb(h, l, s)), origin + " tint " + data["tint"]
    return rgb, origin


def drawing_color(node, theme):
    kind, val = tag(node), node.get("val", "")
    if kind == "srgbClr":
        rgb, origin = val.upper(), "RGB"
    elif kind == "sysClr":
        rgb, origin = node.get("lastClr", "?").upper(), "system " + val
    elif kind == "schemeClr":
        rgb, origin = theme.get(THEME_ALIAS.get(val, val), "?"), "theme " + val
    elif kind == "prstClr":
        rgb, origin = FORMAT_NAMES.get(val.lower(), val), "preset " + val
    elif kind == "scrgbClr":
        # scRGB stores linear light. Convert to the usual sRGB byte representation.
        rgb = hex_rgb(to_srgb(int(node.attrib[key])/100000) for key in ("r", "g", "b"))
        origin = "scRGB"
    else:
        rgb = hex_rgb(colorsys.hls_to_rgb(int(node.get("hue","0"))/21600000,
                      int(node.get("lum","0"))/100000, int(node.get("sat","0"))/100000))
        origin = "HSL"
    if len(node) and re.fullmatch("[0-9A-F]{6}", rgb):
        colors = channels(rgb)
        for transform in node:
            name = tag(transform)
            fraction = int(transform.get("val", "100000"))/100000
            origin += " " + name + "=" + transform.get("val", "")
            if name == "tint":
                colors = tuple(to_srgb(to_linear(v)*fraction+1-fraction) for v in colors)
            elif name == "shade":
                colors = tuple(to_srgb(to_linear(v)*fraction) for v in colors)
            elif name in ("lumMod","lumOff", "satMod", "satOff"):
                h,l,s = colorsys.rgb_to_hls(*colors)
                if name == "lumMod": l *= fraction
                if name == "lumOff": l += fraction
                if name == "satMod": s *= fraction
                if name == "satOff": s += fraction
                colors = colorsys.hls_to_rgb(h, min(1,max(0,l)), min(1,max(0,s)))
            elif name != "alpha" or fraction != 1:
                return rgb, origin
        rgb = hex_rgb(colors)
    return rgb, origin


def format_colors(code, indexed):
    # Ignore bracketed content inside literals or escaped characters.
    clean = re.sub(r'"[^"]*"|\\.', "", code or "")
    for bracket in re.findall(r"\[([^\]]+)\]", clean):
        token = bracket.lower()
        if token in FORMAT_NAMES:
            yield "숫자 서식", FORMAT_NAMES[token], bracket
        elif re.fullmatch(r"color\s*\d+", token):
            idx = int(re.search(r"\d+",token).group()) + 7
            yield "숫자 서식", indexed[idx] if idx < len(indexed) else "?", bracket


def style_tokens(node, role, theme, indexed):
    if node is None:
        return []
    if role == "글씨":
        colored = [(e, role) for e in elements(node, "color")]
    elif role == "테두리":
        colored = [(e, role + " " + tag(side)) for side in node
                   if side.get("style") for e in elements(side, "color")]
    else:
        colored = []
        for pattern in elements(node,"patternFill"):
            typ = pattern.get("patternType")
            for e in pattern:
                # A solid foreground ignores an indexed background placeholder.
                if typ == "none" or (typ == "solid" and tag(e) == "bgColor"):
                    continue
                colored.append((e, role))
        colored.extend((e, role+" 그라데이션") for e in elements(node, "color"))
    return [(purpose, *sheet_color(e, theme, indexed)) for e,purpose in colored]


def styles(z, theme):
    if "xl/styles.xml" not in z.namelist():
        return [v[-6:] for v in COLOR_INDEX], [[("글씨", BLACK, "automatic")]], []
    root = ET.fromstring(z.read("xl/styles.xml"))
    custom = elements(root,"indexedColors")
    indexed = [e.attrib["rgb"][-6:].upper() for e in custom[0]] if custom else [v[-6:] for v in COLOR_INDEX]
    num = dict(BUILTIN_FORMATS)
    num.update({int(x.attrib["numFmtId"]):x.attrib["formatCode"] for x in child(root,"numFmts") or []})
    styles_by_part = {}
    specs = [("fontId","font","글씨"),("fillId","fill","채움"),("borderId","border","테두리")]
    for key,part,role in specs:
        styles_by_part[key] = [style_tokens(node, role, theme, indexed) for node in child(root,part+"s") or []] or [[]]
    base = list(child(root,"cellStyleXfs") or []) or [ET.Element("xf")]
    tokens = []
    for xf in child(root,"cellXfs") or [ET.Element("xf")]:
        items = []
        parent = base[int(xf.get("xfId", 0))]
        for key, part, role in specs:
            field = "apply" + part.title()
            value = parent.get(key, "0") if xf.get(field) == "0" else xf.get(key, parent.get(key, "0"))
            items.extend(styles_by_part[key][int(value)])
        num_id = parent.get("numFmtId", "0") if xf.get("applyNumberFormat") == "0" else xf.get("numFmtId",parent.get("numFmtId","0"))
        items.extend(format_colors(num.get(int(num_id),""),indexed))
        tokens.append(items)
    diff = []
    for entry in child(root,"dxfs") or []:
        items = []
        for key,part,role in specs:
            items.extend(style_tokens(child(entry,part),role,theme,indexed))
        number = child(entry,"numFmt")
        if number is not None:
            items.extend(format_colors(number.get("formatCode"),indexed))
        diff.append(items)
    return indexed, tokens, diff


def drawing_tokens(node, path, theme):
    kind = tag(node)
    if kind in DRAWING_COLORS:
        yield path, *drawing_color(node, theme)
    if kind == "numFmt":
        yield from ((path+" "+name, rgb, origin) for name,rgb,origin in format_colors(node.get("formatCode"),[v[-6:] for v in COLOR_INDEX]))
    for index,item in enumerate(node):
        key = tag(item)
        suffix = f"[{index+1}]" if key in ("ser","dPt","dLbl","legendEntry","r") else ""
        yield from drawing_tokens(item, path+"/"+key+suffix, theme)


def inspect_palette(path):
    """Inventory active style/part references. Retain all exceptional locations by group."""
    rows = {}
    def record(sheet, loc, category, purpose, rgb, origin):
        issue = ""
        if rgb not in ALLOWED:
            issue = "색상표 밖"
        elif category in ("셀", "조건부 서식") and rgb in (GRAY, PEACH):
            issue = "셀 사용 제한"
        elif category == "탭" and rgb not in TAB_COLORS:
            issue = "탭 사용 제한"
        elif category == "차트" and not (origin == "RGB" or purpose.endswith("숫자 서식")):
            issue = "직접 RGB 미지정"
        group = sheet,category,purpose,rgb,origin,issue
        if group not in rows:
            rows[group] = dict(sheet=sheet, category=category, purpose=purpose, color=rgb,
                               source=origin, issue=issue, count=0, locations=[])
        row = rows[group]
        row["count"] += 1
        if issue or len(row["locations"]) < 3:
            row["locations"].append(loc)
    with ZipFile(path) as z:
        theme = theme_colors(z)
        indexed,tokens,diffs = styles(z, theme)
        strings = []
        if "xl/sharedStrings.xml" in z.namelist():
            for si in ET.fromstring(z.read("xl/sharedStrings.xml")):
                strings.append([("서식 문자열",*sheet_color(e,theme,indexed)) for e in elements(si,"color")])
        wb_part = "xl/workbook.xml"
        wb_rels = relationships(z, wb_part)
        sheets = elements(ET.fromstring(z.read(wb_part)), "sheet")
        for sheet in sheets:
            name = sheet.attrib["name"]
            part = wb_rels[sheet.attrib[RID]]
            xml = ET.fromstring(z.read(part))
            refs = relationships(z, part)
            def add(loc, category, uses):
                for use in uses:
                    record(name, loc, category, *use)
            for color in elements(xml,"tabColor"):
                record(name,"시트 탭","탭","탭 색",*sheet_color(color,theme,indexed))
            for column in elements(xml,"col"):
                if "style" in column.attrib:
                    label = f'{get_column_letter(int(column.attrib["min"]))}:{get_column_letter(int(column.attrib["max"]))}'
                    add(label,"셀",tokens[int(column.get("style"))])
            for row in elements(xml,"row"):
                if "s" in row.attrib:
                    add(row.get("r")+":"+row.get("r"),"셀",tokens[int(row.get("s"))])
                for cell in row.findall("{*}c"):
                    loc = cell.attrib["r"]
                    add(loc,"셀",tokens[int(cell.get("s",0))])
                    if cell.get("t") == "s":
                        add(loc,"셀",strings[int(child(cell,"v").text)])
                    elif cell.get("t") == "inlineStr":
                        add(loc,"셀",(("서식 문자열",*sheet_color(e,theme,indexed)) for e in elements(cell,"color")))
            for cf in elements(xml, "conditionalFormatting"):
                sq = child(cf,"sqref")
                loc = cf.get("sqref",sq.text if sq is not None else "조건부 서식")
                for rule in cf.findall("{*}cfRule"):
                    label = loc + " / 규칙 " + rule.get("priority", rule.get("type",""))
                    if "dxfId" in rule.attrib:
                        add(label, "조건부 서식", diffs[int(rule.attrib["dxfId"])])
                    for e in rule.iter():
                        if any(k in e.attrib for k in ("rgb","indexed","theme","auto")):
                            add(label,"조건부 서식",[(tag(e),*sheet_color(e,theme,indexed))])
            for ref in xml.iter():
                if tag(ref) not in ("drawing", "legacyDrawing", "legacyDrawingHF") or ref.get(RID) not in refs:
                    continue
                drawing = refs[ref.get(RID)]
                dr = ET.fromstring(z.read(drawing))
                if drawing.endswith(".vml"):
                    for shape in elements(dr,"shape"):
                        loc = shape.get("id",drawing)
                        for node in shape.iter():
                            for key in ("fillcolor","strokecolor","color","color2"):
                                value = node.get(key)
                                if not value:
                                    continue
                                rgb = FORMAT_NAMES.get(value.lower(),value.lstrip("#").upper())
                                if re.fullmatch(r"[0-9A-F]{3}",rgb):
                                    rgb = "".join(c*2 for c in rgb)
                                record(name,loc,"도형",f"VML {tag(node)} {key}",rgb,value)
                    continue
                dr_rels = relationships(z,drawing)
                for anchor in dr.iter():
                    if tag(anchor) not in ("twoCellAnchor","oneCellAnchor","absoluteAnchor"):
                        continue
                    obj = next(iter(elements(anchor,"cNvPr")),None)
                    loc = obj.get("name",drawing) if obj is not None else drawing
                    add(loc,"도형",drawing_tokens(anchor,"그리기",theme))
                    for chart in elements(anchor,"chart"):
                        part = dr_rels.get(chart.get(RID))
                        if part:
                            chart_xml = ET.fromstring(z.read(part))
                            add(loc,"차트",drawing_tokens(chart_xml,"차트",theme))
                            # A newly inserted/automatic series can inherit the Office accent
                            # without recording any <schemeClr> in the chart XML.
                            for group in chart_xml.iter():
                                kind = tag(group)
                                if kind not in ("barChart", "lineChart", "areaChart"):
                                    continue
                                for i,ser in enumerate(group.findall("{*}ser"),1):
                                    sp = child(ser,"spPr")
                                    paint = child(sp,"ln") if kind == "lineChart" else sp
                                    if paint is None or not any(tag(x) in ("solidFill","noFill","gradFill","pattFill","blipFill") for x in paint):
                                        record(name,loc,"차트",f"{kind}/계열 {i} 자동색","?", "자동 계열 색")
    return list(rows.values())


def warnings_for(rows):
    return [f"색상 검사 {row['issue']} · {row['sheet']}!{row['locations'][0]} · "
            f"{row['purpose']} · {row['color']} ({row['source']}, {row['count']}곳)"
            for row in rows if row["issue"]]
