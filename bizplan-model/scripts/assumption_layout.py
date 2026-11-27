"""Lay out the model's editable assumptions after driver and finance assembly."""

from bisect import bisect_right
from collections import defaultdict
import re
from datetime import date
from tax_policy import LOSS_LIMIT_GENERAL, LOSS_LIMIT_SME, YEAR_MONTHS, credit_inputs, credit_mode, shift_month
from openpyxl.utils.cell import absolute_coordinate
from openpyxl.utils.datetime import to_excel

from openpyxl.utils import get_column_letter as letter
from dashboard import scenario_controls
from tax_credit_input_checks import register_credit_input_checks


ASSUMPTIONS = "전제사항"
TERMS = "금융조건"
SECTIONS = ("기준", "자금", "운전자본", "법인세·영업외손익", "투자·상각", "인건비",
            "매출 라인별 핵심 가정", "비용 라인별 핵심 가정", "시나리오 배수 및 월 가산액")
SCENARIOS = ("Base", "Up", "Down")
MONEY = '#,##0.0#####;(#,##0.0#####);"-"'
RATE = "0.0%"
MULTIPLE = '0.00"x"'
DISPLAY_UNIT = ('=IF(Unit=1000000,"백만원",IF(Unit=100000000,"억원",'
                'IF(Unit=10000,"만원",IF(Unit=1000,"천원",IF(Unit=1,"원",TEXT(Unit,"#,##0")&"원")))))')
TERM_CELL = re.compile(r"(?:'금융조건'|금융조건)!\$?G\$?(\d+)(?!\d)", re.I)
SCENARIO_LOOKUP = re.compile(
    r"INDEX\((?:'전제사항'|전제사항)!\$?J\$?(\d+):\$?L\$?\1,\s*1,\s*([^(),]+)\)", re.I)



TAX_ANNUAL = {
    "additions": "익금산입·손금불산입", "deductions": "손금산입·익금불산입",
    "non_taxable": "비과세소득", "income_deduction": "소득공제",
    "credits": "세액공제·감면", "withholding": "원천징수세액", "ad_hoc": "수시부과세액",
    "interim_paid": "마감일까지 납부한 중간예납", "installment_amount": "확정세액 중 분납할 금액",
}
TAX_PRIOR = {
    "assessed_tax": "직전 사업연도 산출세액", "credits": "직전 사업연도 감면·공제",
    "withholding": "직전 사업연도 원천징수세액", "ad_hoc": "직전 사업연도 수시부과세액",
    "months": "직전 사업연도 월수", "is_sme": "직전 사업연도 중소기업 여부",
}
TAX_OPENING = {
    "national_payable": "기초 미지급 국세", "local_payable": "기초 미지급 지방세",
    "national_prepaid": "기초 선납·환급 국세", "local_prepaid": "기초 선납·환급 지방세",
}
TAX_PAID = {"national": "마감일까지 납부한 국세 총액", "local": "마감일까지 납부한 지방세 총액"}
TAX_ERAS = ((2023, "pre2026", "2025년까지 개시"), (2026, "from2026", "2026년 이후 개시"))


def initialize_tax_refs(m):
    """Reserve reference names before the IS and cash schedules are assembled."""
    tax = m.spec["tax"]
    m.tax_vintage_inputs = tax.get("loss_vintages") or [{"fy_start": "unknown", "amount": 0}]
    keys = list(("fy_start_month", "is_sme", "corporation_type", "include_local", "installments",
                 "faithful_filing", "loss_limit_override", "loss_limit", "founded_fy",
                 "annual_starts", "annual_years"))
    if credit_mode(tax):
        keys.extend(("credit_as_of", "min_tax_apply", "min_tax_rate"))
        keys.extend(("credit", i, field) for i, _ in enumerate(credit_inputs(tax))
                    for field in ("year", "label", "amount", "minimum_tax_applies", "rural_tax"))
    for group, fields in (("annual", TAX_ANNUAL), ("opening", TAX_OPENING), ("prior", TAX_PRIOR),
                          ("paid_to_date", TAX_PAID)):
        keys.extend((group, key) for key in fields)
    for era, _, _ in TAX_ERAS:
        keys.extend(("threshold", era, i) for i in range(3))
        keys.extend(("rate", kind, era, entity, i) for kind in ("national", "local")
                    for entity in ("general", "small") for i in range(4))
    for prefix, records, fields in (("short", tax.get("short_fys", {}), ("start", "end")),
                                    ("vintage", m.tax_vintage_inputs, ("year", "amount"))):
        keys.extend((prefix, i, field) for i in range(len(records)) for field in fields)
    m.tax_refs = {key: "KR_Tax_" + ("_".join(map(str, key)) if isinstance(key, tuple) else key)
                  for key in keys}


def _bind(m, key, address):
    m.named(m.tax_refs[key], ASSUMPTIONS, absolute_coordinate(address))


def _month_serial(month):
    if not month or month == "unknown":
        return None
    return to_excel(date.fromisoformat(month + "-01"))


def _tax_item(m, row, key, label, value, unit, basis, fmt, editable=True):
    _item(m, row, label, value, unit, basis, fmt)
    address = f"G{row}"
    _bind(m, key, address)
    if editable:
        m.register(ASSUMPTIONS, address, label, "법인세", basis)
    return row + 1


def _tax_rates(m, row, tax):
    for era, field, label in TAX_ERAS:
        schedule = tax["rate_tables"][field]
        m.cell(ASSUMPTIONS, row, 3, label + " 사업연도 세율표", "text")
        m.table(ASSUMPTIONS, row, 11, [["1구간", "2구간", "3구간", "4구간"]])
        m.header(ASSUMPTIONS, row)
        row += 1
        _item(m, row, "과세표준 상한", None, DISPLAY_UNIT, "각 구간의 상한. 마지막 구간은 상한이 없습니다.", MONEY)
        for i, amount in enumerate(schedule["thresholds"]):
            m.cell(ASSUMPTIONS, row, 11 + i, amount / m.spec["unit"], "input")
            _bind(m, ("threshold", era, i), f"{letter(11+i)}{row}")
        m.cell(ASSUMPTIONS, row, 14, "상한 없음", "text")
        m.register(ASSUMPTIONS, f"K{row}:M{row}", label + " 구간 상한", "법인세", "세율표 과세표준 경계")
        row += 1
        for kind, title in (("national", "법인세"), ("local", "법인지방소득세")):
            for entity, caption in (("general", "일반"), ("small", "성실신고확인대상 소규모")):
                column = kind + ("_general" if entity == "general" else "_small_filing")
                note = "세무 기준서 제3·4절. 구간별 초과 금액에 적용합니다."
                if entity == "small":
                    note += " 소규모 특례는 2025년 개시부터 적용하며 지방세 기본값은 해당 국세율의 10분의 1입니다."
                name = title + " · " + caption
                _item(m, row, name, None, "%", note, RATE)
                for i, v in enumerate(schedule[column]):
                    m.cell(ASSUMPTIONS, row, 11+i, v, "input")
                    _bind(m, ("rate", kind, era, entity, i), f"{letter(11+i)}{row}")
                m.register(ASSUMPTIONS, f"K{row}:N{row}", label + " " + name, "법인세", note)
                row += 1
    return row


def _tax_annual(m, row, tax, shorts):
    """Use dated FY starts, including a stub and the following resumed FY.

    Regular columns follow the editable fiscal month. Suppress a regular key
    that falls in a stub or coincides with a resumed period to avoid duplicates.
    """
    annual = tax.get("annual", {})
    regular_month = tax["fy_start_month"]
    earliest = min([int(m.months[0][:4])-1, *(int(t[:4]) for t in annual)])
    latest = max([int(m.months[-1][:4]), *(int(t[:4]) for t in annual)])
    if credit_mode(tax):
        # Include regular starts around historical stubs so an edited vintage
        # cannot identify two legal fiscal periods with one calendar year.
        boundaries = [month for pair in shorts for month in pair]
        if tax.get("founded_fy"):
            boundaries.append(tax["founded_fy"])
        earliest = min([earliest, *(int(month[:4]) for month in boundaries)])
        latest = max([latest, *(int(month[:4]) for month in boundaries)])
    fixed = {}
    for i, (start, end) in enumerate(shorts):
        fixed[start] = m.tax_refs["short", i, "start"]
        resumed = shift_month(end, 1)
        fixed.setdefault(resumed, f'EDATE({m.tax_refs["short",i,"end"]},1)')
    # Keep the short-period columns first so the regular-date formulas can
    # check them without referencing another regular-date formula.
    columns = [(start, "=" + formula) for start, formula in sorted(fixed.items())]
    founded = tax.get("founded_fy")
    founded_ref = m.tax_refs["founded_fy"]
    founded_formula = f'IF(ISNUMBER({founded_ref}),{founded_ref},"")'
    if fixed:
        initial = f"$K${row}:${letter(10+len(fixed))}${row}"
        founded_formula = f'IF(COUNTIF({initial},{founded_ref}),"",{founded_formula})'
    columns.append((founded if founded not in fixed else None, "=" + founded_formula))
    for year in range(earliest, latest + 1):
        candidate = f'DATE({year},{m.tax_refs["fy_start_month"]},1)'
        key = f"{year:04}-{regular_month:02}"
        fixed_range = f"$K${row}:${letter(11+len(fixed))}${row}"
        if shorts:
            stub_start = m.tax_refs["short", 0, "start"]
            stub_end = m.tax_refs["short", len(shorts)-1, "start"]
            end_first = m.tax_refs["short", 0, "end"]
            end_last = m.tax_refs["short", len(shorts)-1, "end"]
            occupied = (f'OR(COUNTIF({fixed_range},{candidate})>0,COUNTIFS('
                        f'{stub_start}:{stub_end},"<="&{candidate},{end_first}:{end_last},">="&{candidate})>0)')
            formula = f'=IF({occupied},"",{candidate})'
        else:
            formula = f'=IF(COUNTIF({fixed_range},{candidate}),"",{candidate})'
        if credit_mode(tax):
            formula = f'=IF(AND(ISNUMBER({founded_ref}),{candidate}<{founded_ref}),"",{formula[1:]})'
        columns.append((None if key in fixed or key == founded else key, formula))
    right = letter(10 + len(columns))
    for offset, (_, formula) in enumerate(columns, 11):
        m.cell(ASSUMPTIONS, row, offset, formula)
    _item(m, row, "사업연도 개시일", None, "월", "세무조정과 납부 자료는 해당 사업연도의 개시 월에 입력합니다.", "yyyy-mm")
    m.header(ASSUMPTIONS, row)
    _bind(m, "annual_starts", f"K{row}:{right}{row}")
    for offset in range(11, 11 + len(columns)):
        ref = f"{letter(offset)}{row}"
        m.cell(ASSUMPTIONS, row+1, offset, f'=IF(ISNUMBER({ref}),YEAR({ref}),"")')
    _item(m, row+1, "개시 연도", None, "년", "사업연도 개시일 기준", "0")
    _bind(m, "annual_years", f"K{row+1}:{right}{row+1}")
    row += 2
    for field, label in TAX_ANNUAL.items():
        note = "사업연도 전체 금액. 세무조정계산서와 납부 내역을 확인하세요."
        if field == "credits":
            note += (" 유형별 표와 중복 입력하지 마세요. 이 금액은 당해 연도 일반 감면으로 처리됩니다."
                     if credit_mode(tax) else " 공제·감면 입력 시 최저한세 검토가 필요합니다.")
        if field == "installment_amount":
            note = "분납을 선택한 경우 실제로 이연할 금액을 입력합니다."
        if field == "interim_paid":
            note = "실적 마감일까지 해당 사업연도에 이미 납부한 중간예납액입니다."
        _item(m, row, label, None, DISPLAY_UNIT, note, MONEY)
        for col, (start, _) in enumerate(columns, 11):
            m.cell(ASSUMPTIONS, row, col, annual.get(start, {}).get(field, 0) / m.spec["unit"], "input")
        address = f"K{row}:{right}{row}"
        m.register(ASSUMPTIONS, address, label, "법인세", note)
        _bind(m, ("annual", field), address)
        row += 1
    return row


def _tax_credit_inputs(m, row, tax):
    minimum = tax["minimum_tax"]
    row = _tax_item(m, row, "min_tax_apply", "최저한세 적용 여부", int(minimum["apply"]),
                    "0 / 1", "1이면 최저한세 대상 세액공제에 법정 한도를 적용합니다.", "0")
    row = _tax_item(m, row, "min_tax_rate", "최저한세율 (자동 / 직접)", minimum["rate"], "%",
                    "auto는 중소기업 여부와 과세표준에 따른 법정률입니다. 중소기업 졸업 시 해당 8% 또는 9%를 직접 입력합니다.", RATE)
    row = _tax_item(m, row, "credit_as_of", "이월세액공제 잔액 기준 월",
                    _month_serial(tax["credit_balance_as_of"]), "월",
                    "기초 잔액을 확인한 사업연도의 마지막 월. 다음 사업연도부터 사용액을 계산합니다.", "yyyy-mm")
    lots = credit_inputs(tax)
    register_credit_input_checks(m, row-1, row+1, lots)
    if not lots:
        return row
    columns = ("year", "label", "amount", "minimum_tax_applies", "rural_tax")
    titles = ("발생 사업연도", "공제 유형", "금액", "최저한세 대상", "농어촌특별세 대상")
    m.table(ASSUMPTIONS, row, 3, [["세액공제 발생연도·유형별 입력"]])
    m.table(ASSUMPTIONS, row, 11, [(*titles, "입력 확인")])
    m.header(ASSUMPTIONS, row)
    row += 1
    for i, lot in enumerate(lots):
        category = "기초 이월 잔액" if lot["source"] == "opening" else "당기 신규 발생"
        note = "발생 사업연도는 개시일의 연도. 최저한세·농어촌특별세 대상은 1, 비대상은 0."
        _item(m, row, category, None, DISPLAY_UNIT, note, MONEY)
        values = (lot["year"], lot["label"], lot["amount"] / m.spec["unit"],
                  int(lot["minimum_tax_applies"]), int(lot["rural_tax"]))
        for col, (field, title, value) in enumerate(zip(columns, titles, values), 11):
            address = f"{letter(col)}{row}"
            m.cell(ASSUMPTIONS, row, col, value, "input")
            _bind(m, ("credit", i, field), address)
            m.register(ASSUMPTIONS, address, f"{category} {i+1} · {title}", "법인세", lot["label"])
        row += 1
    return row


def _tax_inputs(m, row):
    tax = m.spec["tax"]
    unit = m.spec["unit"]
    settings = (
        ("fy_start_month", "사업연도 개시 월", tax["fy_start_month"], "월", "사업연도·원천 세율표·납부월의 기준", "0"),
        ("is_sme", "중소기업 여부", int(tax["is_sme"]), "0 / 1", "1 중소기업. 미확인 시 0으로 계산합니다.", "0"),
        ("corporation_type", "법인 구분", int(tax["corporation_type"] == "small_filing"), "0 / 1", "0 일반, 1 성실신고확인대상 소규모 법인", "0"),
        ("include_local", "법인지방소득세 포함", int(tax["include_local"]), "0 / 1", "1이면 국세와 지방세를 함께 반영합니다.", "0"),
        ("installments", "분납 여부", int(tax["installments"]), "0 / 1", "분납할 사업연도별 금액은 아래 납부 입력표에 적습니다.", "0"),
        ("faithful_filing", "성실신고확인서 제출 여부", int(tax["faithful_filing"]), "0 / 1", "확정 납부기한에 반영합니다.", "0"),
        ("founded_fy", "신설 법인의 첫 사업연도", _month_serial(tax.get("founded_fy")), "월", "해당 시 첫 사업연도 개시일을 입력합니다.", "yyyy-mm"),
        ("loss_limit_override", "결손금 공제 한도 별도 지정", tax.get("loss_limit_override"), "%", "빈칸이면 중소기업 여부에 따른 한도를 적용합니다.", RATE),
        ("loss_limit", "적용 결손금 공제 한도율", f'=IF(ISNUMBER({m.tax_refs["loss_limit_override"]}),{m.tax_refs["loss_limit_override"]},IF({m.tax_refs["is_sme"]},{LOSS_LIMIT_SME},{LOSS_LIMIT_GENERAL}))', "%", "세무 기준서 제2절", RATE),
    )
    for key, label, value, u, note, fmt in settings:
        row = _tax_item(m, row, key, label, value, u, note, fmt, key != "loss_limit")
    if credit_mode(tax):
        row = _tax_credit_inputs(m, row, tax)
    prior = tax.get("prior", {})
    for key, label in TAX_PRIOR.items():
        v = prior.get(key)
        if key == "months":
            v, u, fmt = prior.get(key, YEAR_MONTHS), "개월", "0"
        elif key == "is_sme":
            v, u, fmt = int(prior.get(key, tax["is_sme"])), "0 / 1", "0"
        else:
            v, u, fmt = (None if v is None and key == "assessed_tax" else (v or 0) / unit), DISPLAY_UNIT, MONEY
        note = "직전 사업연도 신고서 기준. 산출세액이 비어 있으면 상반기 실적 기준으로 중간예납을 계산합니다."
        row = _tax_item(m, row, ("prior", key), label, v, u, note, fmt)
    for key, label in TAX_OPENING.items():
        v = tax.get("opening", {}).get(key)
        note = ("빈칸은 실적 BS 잔액을 사용합니다. 지방세 구분이 없는 미지급 잔액은 국세로 분류합니다. "
                "BS 없는 혼합 사업연도는 실적 법인세비용에서 확인된 기납부액을 뺀 추정 잔액을 사용합니다. 0 입력은 잔액 없음입니다.")
        row = _tax_item(m, row, ("opening", key), label, None if v is None else v / unit, DISPLAY_UNIT, note, MONEY)
    for key, label in TAX_PAID.items():
        paid = tax.get("paid_to_date", {}).get(key)
        note = ("실적 마감 사업연도의 기납부 총액입니다. 확인된 중간예납·원천징수 등 기납부액을 포함합니다. "
                "빈칸이면 해당 세목의 마감 잔액을 기준으로 계산하며, 0은 납부액 없음입니다.")
        row = _tax_item(m, row, ("paid_to_date", key), label, None if paid is None else paid / unit,
                        DISPLAY_UNIT, note, '#,##0.0#####;(#,##0.0#####);0.0')
    row = _tax_rates(m, row, tax)
    vintages = m.tax_vintage_inputs
    if vintages:
        m.table(ASSUMPTIONS, row, 3, [["이월결손금 발생 내역"]])
        m.table(ASSUMPTIONS, row, 11, [["발생 사업연도 개시일", "마감 잔액"]])
        m.header(ASSUMPTIONS, row)
        row += 1
    for i, vintage in enumerate(vintages):
        start = _month_serial(vintage["fy_start"])
        _item(m, row, "발생연도별 이월결손금", None, DISPLAY_UNIT,
              "발생연도 미상이면 개시일을 비워 둡니다. 이 잔액의 만료일은 계산할 수 없습니다.", MONEY)
        for col, field, value in ((11, "year", start), (12, "amount", vintage["amount"] / unit)):
            addr = f"{letter(col)}{row}"
            m.cell(ASSUMPTIONS, row, col, value, "input")
            _bind(m, ("vintage", i, field), addr)
            m.register(ASSUMPTIONS, addr, "결손금 발생 월" if col == 11 else "이월결손금 잔액", "법인세", vintage["fy_start"])
        row += 1
    shorts = sorted(tax.get("short_fys", {}).items())
    if shorts:
        m.table(ASSUMPTIONS, row, 3, [["단기 사업연도"]])
        m.table(ASSUMPTIONS, row, 11, [["개시 월", "종료 월"]])
        m.header(ASSUMPTIONS, row)
        row += 1
    for i, (start, end) in enumerate(shorts):
        _item(m, row, "단기 사업연도", None, "월", "실제 단기 사업연도의 개시 월과 포함되는 마지막 월", "yyyy-mm")
        for col, field, value in ((11, "start", start), (12, "end", end)):
            addr = f"{letter(col)}{row}"
            m.cell(ASSUMPTIONS, row, col, _month_serial(value), "input")
            _bind(m, ("short", i, field), addr)
            m.register(ASSUMPTIONS, addr, "단기 사업연도 " + ("개시 월" if col == 11 else "종료 월"), "법인세", start)
        row += 1
    row = _tax_annual(m, row, tax, shorts)
    missing = set(m.tax_refs.values()).difference(m.names)
    if missing:
        raise ValueError(f"세무 수식에 연결되지 않은 입력 이름: {sorted(missing)}")
    return row


def _unit(label):
    if label.endswith("인상 시작 월") or label in ("소급 적용 월", "소급 정산 월"):
        return "월", "yyyy-mm"
    if label.endswith("1인 월 비용 (원)"):
        return "원/인·월", "#,##0"
    if label.endswith("객단가 (원)"):
        return "원/고객·월", "#,##0"
    if label.endswith("단가 (원)"):
        return "원/개", "#,##0.0###"
    if label.endswith("(원)"):
        return DISPLAY_UNIT, MONEY
    if label.endswith(("률", "율", "확률")):
        return "%", RATE
    if label.endswith("(일)"):
        return "일", "0.##"
    if label.endswith("(개월)"):
        return "개월", "0"
    if label.endswith(("고객 수", "인원")):
        return "명", "#,##0.0###"
    if label.endswith("수량"):
        return "개", "#,##0.##"
    return "0 / 1", "0"


def _caption(label):
    return re.sub(r"\s*\((?:원|일|개월)\)$", "", label)


def _header(m, row, number):
    m.cell(ASSUMPTIONS, row, 2, f"{number}. {SECTIONS[number - 1]}", "text")
    m.section(ASSUMPTIONS, row)
    m.table(ASSUMPTIONS, row + 1, 3, [["항목"]])
    m.table(ASSUMPTIONS, row + 1, 7, [["적용값" if number == 9 else "값", "단위", None, "근거 / 입력 위치"]])
    if number == 9:
        m.table(ASSUMPTIONS, row + 1, 11, [SCENARIOS])
    m.header(ASSUMPTIONS, row + 1)
    return row + 2


def _item(m, row, label, value, unit, basis, fmt, level=3):
    m.cell(ASSUMPTIONS, row, level, label, "text")
    m.level(ASSUMPTIONS, row, level)
    m.cell(ASSUMPTIONS, row, 7, value)
    m.cell(ASSUMPTIONS, row, 8, unit)
    m.cell(ASSUMPTIONS, row, 10, basis, "text")
    # Amounts expressed in the display unit keep their full input precision.
    m.format_row(ASSUMPTIONS, row, '#,##0.0;(#,##0.0);"-"' if fmt == MONEY else fmt)


def _collect(m):
    """Assign registered rows to their spec owner, independent of sheet ordering."""
    revenues = {m.details[line["id"]]: line for line in m.spec["revenue_lines"]}
    costs = m.spec["cost_lines"]
    # The finance driver emits one section for each line, in spec order, after
    # the summary headers. Labels can repeat across expense classifications.
    cost_sections = sorted(m.sections["운영비용"])[-len(costs):] if costs else []
    lines = {line["id"]: line for line in m.lines}
    rows = defaultdict(list)
    for monthly, inputs in ((False, m.params), (True, m.monthly_inputs)):
        for record in inputs:
            sheet, r, label, _, group, description = record
            if sheet == ASSUMPTIONS or (sheet == TERMS and not 12 <= r <= 25):
                continue
            line_id = None
            if sheet in revenues:
                section, owner = 7, revenues[sheet]
                line_id = owner["id"]
            elif sheet == "운영비용":
                section, owner = 8, costs[bisect_right(cost_sections, r) - 1]
                line_id = owner["id"]
            elif sheet == "확장조정":
                line_id = m.spec["extensions"][r - 8]["line_id"]
                owner = lines[line_id]
                section = 7 if line_id in m.details else 8
            else:
                if sheet == TERMS:
                    section, field = ((3, "working_capital") if 17 <= r <= 19 or r in (24, 25) else
                                      (4, "tax") if 20 <= r <= 21 else (2, "financing"))
                elif sheet == "투자비":
                    section, field = 5, "investment"
                elif sheet == "인건비":
                    section, field = 6, "wages"
                else:
                    raise ValueError(f"전제사항 구역을 확인하세요 ({sheet} {r}행 / {label}).")
                owner = m.spec.get(field, {})
            basis = owner.get("basis") or description
            rows[section, line_id].append((record, monthly, basis))
    return rows


def _remove_registrations(m, addresses):
    m.inputs[:] = [entry for entry in m.inputs if (entry[2], entry[3]) not in addresses]
    for sheet in {s for s, _ in addresses}:
        m.paint[sheet][:] = [a for a in m.paint[sheet] if (sheet, a) not in addresses]


def _write_input(m, row, record, monthly, basis, level, line_label, terms_refs):
    sheet, old, label, value, group, description = record
    unit, fmt = _unit(label)
    monetary = unit == DISPLAY_UNIT
    caption = _caption(label)
    if line_label:
        caption = caption.removeprefix(line_label + " ")
    if monthly:
        address = f"J{old}:{m.col(len(value)-1)}{old}"
        if monetary:
            for col, amount in enumerate(value, 10):
                if amount is not None:
                    m.cell(sheet, old, col, amount / m.spec["unit"], "input")
            label = _caption(label) + " (표시 단위)"
            m.cell(sheet, old, 2, label, "text")
            description = description.replace("원 금액", "표시 단위 금액")
        m.format_row(sheet, old, fmt)
        m.register(sheet, address, label, group, description)
        escaped = sheet.replace("'", "''").replace('"', '""')
        link = f'=HYPERLINK("#\'{escaped}\'!J{old}","월별 입력")'
        _item(m, row, caption, link, unit, f"{basis}\n월별 입력: {sheet} 시트 {old}행", fmt, level)
        return

    display_value = value / m.spec["unit"] if monetary else value
    _item(m, row, caption, display_value, unit, basis, fmt, level)
    new_label = _caption(label) + " (표시 단위)" if monetary else label
    m.register(ASSUMPTIONS, f"G{row}", new_label, group, description)
    m.scalar_input_cells[sheet, old] = f"G{row}"
    if sheet == "인건비" and label == "소급 인식 방식 (0 발생, 1 일시)":
        m.named("WG_Recognition", ASSUMPTIONS, f"$G${row}")
    target = f"'{ASSUMPTIONS}'!$G${row}"
    if sheet == "인건비" and old in m.wage_date_rows:
        target = f"DATE(YEAR({target}),MONTH({target}),1)"
    if monetary:
        target = f"({target}*Unit)"
    if sheet == TERMS:
        terms_refs[old] = target
        for col in (2, 7):
            m.grid[sheet].pop((old, col), None)
    else:
        # Existing detail formulas use the scalar's original unit, including KRW.
        m.cell(sheet, old, 7, "=" + target)
        m.level(sheet, old, 4)
        m.format_row(sheet, old, "#,##0" if monetary else fmt)


def _rewrite_finance(m, refs):
    def replace(formula):
        return TERM_CELL.sub(lambda hit: refs.get(int(hit[1]), hit[0]), formula)
    for grid in m.grid.values():
        for address, value in grid.items():
            if isinstance(value, str) and value.startswith("=") and TERMS in value:
                grid[address] = replace(value)
    for index, (label, formula) in enumerate(m.checks):
        if formula and TERMS in formula:
            m.checks[index] = (label, replace(formula))


def _scenarios(m, row, old_values):
    new_rows = {}
    scenario_basis = [m.spec["scenarios"][name].get("basis") for name in SCENARIOS]
    basis = "\n".join(f"{name} · {note}" for name, note in zip(SCENARIOS, scenario_basis) if note)
    for line in m.lines:
        for key, caption, _ in scenario_controls(line, m.spec["scenarios"],
                                                 m.spec["investment"]["depreciation_accounts"]):
            ident = line["id"], key
            old = m.scenario_refs[ident]
            addition = key == "additions"
            label = line["label"] + " " + caption
            values = old_values[old]
            if addition:
                values = [v / m.spec["unit"] for v in values]
            unit, fmt = (DISPLAY_UNIT, MONEY) if addition else ("배", MULTIPLE)
            note = basis or "시나리오별 " + caption
            inherited = None
            if key == "renewal_probability":
                inherited = m.renewal_probability_refs[line["id"]]
                unit, fmt = "%", RATE
                note = "빈칸이면 해당 매출 라인의 기본 계약 갱신 확률을 적용합니다. 0%는 갱신 없음입니다."
            if key == "exclude_unconfirmed" or key.startswith("exclude_item:"):
                unit, fmt = "0 / 1", "0"
                note = ("1이면 실현 확률 100% 미만 신규 유입을 제외합니다. 기존 고객에는 이탈률을 적용합니다."
                        if line.get("driver") == "cohort" else "1이면 실현 확률 100% 미만 계약을 제외합니다.")
                if key.startswith("exclude_item:"):
                    note = "1이면 이 계약의 실현 확률이 100% 미만일 때만 제외합니다."
                if basis:
                    note += "\n" + basis
            current = f"INDEX(K{row}:M{row},1,Scenario)"
            if inherited:
                current = f"IF(ISBLANK({current}),{inherited},{current})"
            _item(m, row, label, "=" + current, unit, note, fmt)
            for col, value in enumerate(values, 11):
                m.cell(ASSUMPTIONS, row, col, value, "input")
            m.register(ASSUMPTIONS, f"K{row}:M{row}", label, "시나리오", note)
            new_rows[old] = (row, addition, inherited)
            m.scenario_refs[ident] = row
            row += 1

    def lookup(hit):
        new, addition, inherited = new_rows[int(hit[1])]
        expr = f"INDEX('{ASSUMPTIONS}'!$K${new}:$M${new},1,{hit[2]})"
        if inherited:
            return f"IF(ISBLANK({expr}),{inherited},{expr})"
        return f"({expr}*Unit)" if addition else expr
    grid = m.grid["Sensitivity"]
    for address, value in grid.items():
        if isinstance(value, str) and value.startswith("=") and ASSUMPTIONS in value:
            grid[address] = SCENARIO_LOOKUP.sub(lookup, value)


def _facility_policies(m, row):
    entries = [e for e in m.inputs if e[2] == TERMS and re.fullmatch(r"I[0-9]+", e[3])]
    _remove_registrations(m, {(TERMS, e[3]) for e in entries})
    for group, label, _, address, description in entries:
        source_row = int(address[1:])
        policy = m.grid[TERMS][source_row, 9]
        _item(m, row, label, policy, "-1 / 0 / 1", description, "0", 4)
        m.register(ASSUMPTIONS, f"G{row}", label, group, description)
        m.cell(TERMS, source_row, 9, f"='{ASSUMPTIONS}'!$G${row}")
        row += 1
    return row


def _non_operating(m, row):
    spec = m.spec.get("non_operating", {})
    fields = (("income_monthly", "월 영업외수익"), ("other_expense_monthly", "월 기타 영업외비용 (이자 제외)"))
    forecast = range(m.last + 1, len(m.months))
    if any(isinstance(spec.get(key), dict) for key, _ in fields) and forecast:
        m.cell(ASSUMPTIONS, row, 3, "영업외 월별 입력", "text")
        for col, index in enumerate(forecast, 11):
            m.cell(ASSUMPTIONS, row, col, f"='INDEX'!{m.col(index)}5")
        m.header(ASSUMPTIONS, row)
        m.format_row(ASSUMPTIONS, row, "yyyy-mm")
        row += 1
    for key, label in fields:
        value = spec.get(key, 0)
        monthly = isinstance(value, dict)
        note = spec.get("basis") or "추정 기간의 월별 영업외손익"
        applied = "월별 입력 →" if monthly else value / m.spec["unit"]
        _item(m, row, label, applied, DISPLAY_UNIT, note, MONEY)
        address = f"G{row}"
        for col, index in enumerate(forecast, 11):
            source = "G"
            if monthly:
                source = letter(col)
                amount = value.get(m.months[index], 0)
                m.cell(ASSUMPTIONS, row, col, amount / m.spec["unit"], "input")
            m.cell("IS", m.non_operating_rows[key], 10 + index, f"='{ASSUMPTIONS}'!${source}${row}")
        if monthly and forecast:
            address = f"K{row}:{letter(10 + len(forecast))}{row}"
        if not monthly or forecast:
            m.register(ASSUMPTIONS, address, label, "영업외손익", note)
        row += 1
    return row


def _subsidy_link_inputs(m, row, line):
    from activity_model import subsidy_reference

    link = line["params"]["subsidy_link"]
    basis = link.get("basis") or line.get("basis") or "구매 단가에 연동하는 단위당 보조 조건"
    fields = [("threshold", "연동 기준 단가", False, link["threshold"]),
              ("rate", "기본 연동률", True, link["rate"]),
              ("cap", "단위당 연동 한도", False, link["cap"]),
              ("eligible_share", "지원 대상 비율", True, link["eligible_share"]),
              ("baseline_amount", "기준 연동 보조액", False, link["baseline_amount"])]
    fields.extend((f"rate_{year}", f"{year}년 연동률", True, value)
                  for year, value in sorted(link.get("annual_rates", {}).items()))
    start = row
    for key, label, is_rate, value in fields:
        _item(m, row, label, value, "%" if is_rate else "원/단위", basis,
              "0.00%" if is_rate else "#,##0.00##", level=4)
        m.register(ASSUMPTIONS, f"G{row}", f"{line['label']} · {label}", "단가 연동 보조", basis)
        m.named(subsidy_reference(m, line["id"], key), ASSUMPTIONS, f"$G${row}")
        row += 1
    inputs = f"'{ASSUMPTIONS}'!$G${start}:$G${row-1}"
    rates = f"'{ASSUMPTIONS}'!$G${start+1},'{ASSUMPTIONS}'!$G${start+3}"
    if link.get("annual_rates"):
        rates += f",'{ASSUMPTIONS}'!$G${start+5}:$G${row-1}"
    m.add_check(f"{line['label']} 단가 연동 보조 입력 범위",
                f'=IFERROR(IF(AND(COUNT({inputs})={len(fields)},MIN({inputs})>=0,MAX({rates})<=1),"OK","NO"),"NO")')
    return row


def finalize_assumptions(m):
    """Publish editable cells, redirect their consumers and refresh input metadata."""
    m.scalar_input_cells = {}
    grouped = _collect(m)
    grid = m.grid[ASSUMPTIONS]
    old_values = {r: [grid.get((r, c)) for c in (10, 11, 12)] for r in m.scenario_refs.values()}
    cutoff, scenario = grid[8, 7], grid[9, 7]
    addresses = {(ASSUMPTIONS, a) for a in m.paint[ASSUMPTIONS]}
    for records in grouped.values():
        for (sheet, r, _, value, _, _), monthly, _ in records:
            a = f"J{r}:{m.col(len(value)-1)}{r}" if monthly else f"G{r}"
            addresses.add((sheet, a))
    _remove_registrations(m, addresses)
    for address in [a for a in grid if a[0] > 2]:
        del grid[address]
    _header(m, 4, 1)
    _item(m, 6, "계획 기준 기간 시작", "='INDEX'!J5", "월", "실적 및 추정 월 축의 시작", "yyyy-mm")
    _item(m, 7, "표시 단위", m.spec["unit"], "원", "표시 단위 1당 원 금액. 표시 단위 입력액은 이 값으로 원 환산합니다.", "#,##0")
    m.register(ASSUMPTIONS, "G7", "표시 단위 (원)", "단위", "표시 단위 1당 원 금액")
    _item(m, 8, "실적 마감 월", cutoff, "월", "설계서의 실적 마감 월", "yyyy-mm")
    _item(m, 9, "적용 시나리오 번호", scenario, "번호", "Sensitivity에서 선택한 케이스의 시나리오", "0")
    last = m.col(len(m.months) - 1)
    _item(m, 10, "계획 종료 월", f"='INDEX'!{last}5", "월", "설계서의 계획 종료 월", "yyyy-mm")
    terms_refs = {}
    line_labels = {line["id"]: line["label"] for line in m.lines}
    row = 12
    for section in range(2, 9):
        row = _header(m, row, section)
        if section in (7, 8):
            lines = m.spec["revenue_lines" if section == 7 else "cost_lines"]
        else:
            lines = [None]
        for line in lines:
            records = grouped[section, line["id"] if line else None]
            if line:
                title, note = line["label"], line.get("basis") or ""
                method = line.get("method")
                if method in ("mixed", "pct_of_line"):
                    title += " · " + ("준변동비" if method == "mixed" else "라인 연동 비용")
                    ids = line["params"]["revenue_ids" if method == "mixed" else "line_ids"]
                    note += "\n연동 대상 · " + ", ".join(line_labels[k] for k in ids)
                m.cell(ASSUMPTIONS, row, 3, title, "text")
                m.cell(ASSUMPTIONS, row, 10, note.strip(), "text")
                m.level(ASSUMPTIONS, row, 3)
                row += 1
            for record, monthly, basis in records:
                if line and record[2] == "기본 계약 갱신 확률":
                    m.renewal_probability_refs[line["id"]] = f"'{ASSUMPTIONS}'!$G${row}"
                _write_input(m, row, record, monthly, basis, 4 if line else 3,
                             line["label"] if line else None, terms_refs)
                row += 1
            if line and line.get("method") == "consumption" and "subsidy_link" in line["params"]:
                row = _subsidy_link_inputs(m, row, line)
        if section == 2:
            row = _facility_policies(m, row)
        if section == 3:
            for key, caption in (("dio_base_lines", "재고 기준 비용"), ("dpo_base_lines", "매입채무 기준 비용")):
                ids = m.spec["working_capital"].get(key)
                if ids:
                    _item(m, row, caption, "지정 라인 합계", "", ", ".join(line_labels[k] for k in ids), "@")
                    row += 1
        if section == 4:
            row = _tax_inputs(m, row)
            row = _non_operating(m, row)
        row += 1
    _rewrite_finance(m, terms_refs)
    _scenarios(m, _header(m, row, 9), old_values)
    # Show a source sentence once per section or line. Keep row-specific
    # detail locations and preserve every original description in META_Inputs.
    seen = set()
    for r in sorted({r for r, _ in grid}):
        if r in m.sections[ASSUMPTIONS] or r in m.headers[ASSUMPTIONS] or (
                (r, 3) in grid and (r, 8) not in grid):
            seen.clear()
        note = grid.get((r, 10))
        if isinstance(note, str) and note.startswith("'"):
            lines = note[1:].splitlines()
            unique = [line for line in lines if line and line not in seen]
            seen.update(lines)
            if unique:
                m.cell(ASSUMPTIONS, r, 10, "\n".join(unique), "text")
            else:
                grid.pop((r, 10))


def style_assumptions(ws, m):
    """Keep long driver captions, provenance, units and input values readable."""
    ws.Columns("B:F").ColumnWidth = 3
    ws.Columns("G:G").ColumnWidth = 16
    ws.Columns("H:H").ColumnWidth = 14
    ws.Columns("I:I").ColumnWidth = 2
    ws.Columns("J:J").ColumnWidth = 62
    ws.Columns("K:N").ColumnWidth = 16
    if credit_mode(m.spec["tax"]):
        lots = credit_inputs(m.spec["tax"])
        rule = ws.Range(m.tax_refs["min_tax_apply"].rsplit("!", 1)[-1]).Validation
        rule.Delete()
        rule.Add(1, 1, 1, "0", "1")
        ref = m.tax_refs["min_tax_rate"].rsplit("!", 1)[-1]
        rule = ws.Range(ref).Validation
        rule.Delete()
        rule.Add(7, 1, 1, f'=OR({ref}="auto",AND(ISNUMBER({ref}),{ref}>=0,{ref}<=1))')
        if lots:
            def credit_column(field):
                first = m.tax_refs["credit", 0, field].rsplit("!", 1)[-1]
                last = m.tax_refs["credit", len(lots)-1, field].rsplit("!", 1)[-1]
                return ws.Range(f"{first}:{last}")
            ws.Columns("O:O").ColumnWidth = 17
            credit_column("label").WrapText = True
            for field, low, high in (("year", "1900", "9999"), ("minimum_tax_applies", "0", "1"),
                                     ("rural_tax", "0", "1")):
                cells = credit_column(field)
                cells.NumberFormat = "0"
                cells.Validation.Delete()
                cells.Validation.Add(1, 1, 1, low, high)
            cells = credit_column("amount")
            cells.Validation.Delete()
            cells.Validation.Add(2, 1, 7, "0")
    grid = m.grid[ASSUMPTIONS]
    bottom = max(r for r, c in grid)
    ws.Range(f"J4:J{bottom}").WrapText = True
    ws.Range(f"B4:D{bottom}").WrapText = True
    # Excel cannot AutoFit a merged caption. Measure each original caption
    # at the combined width of its indent band before creating the merges.
    for column, width in (("B", 41), ("C", 38), ("D", 35)):
        ws.Columns(f"{column}:{column}").ColumnWidth = width
    ws.Range(f"B4:M{bottom}").Rows.AutoFit()
    for row, col in sorted((r, c) for r, c in grid if r >= 4 and c in (2, 3, 4)):
        first = "BCD"[col - 2]
        band = ws.Range(f"{first}{row}:F{row}")
        height = max(19, band.RowHeight)
        band.Merge()
        band.WrapText = True
        band.ShrinkToFit = False
        band.RowHeight = height
    ws.Columns("B:D").ColumnWidth = 3
    ws.Columns("E:F").ColumnWidth = 16
    for key, low, high in (
        ("fy_start_month", 1, YEAR_MONTHS), (("prior", "months"), 1, YEAR_MONTHS),
        ("is_sme", 0, 1), (("prior", "is_sme"), 0, 1), ("include_local", 0, 1),
        ("installments", 0, 1), ("corporation_type", 0, 1), ("faithful_filing", 0, 1),
        ("loss_limit_override", 0, 1),
    ):
        address = m.tax_refs[key].rsplit("!", 1)[-1]
        rule = ws.Range(address).Validation
        rule.Delete()
        rule.Add(2 if key == "loss_limit_override" else 1, 1, 1, str(low), str(high))
    validation = ws.Range("G7").Validation
    validation.Delete()
    validation.Add(1, 1, 5, "0")  # xlValidateWholeNumber, xlGreater
    for (_, key), row in m.scenario_refs.items():
        if key in ("exclude_unconfirmed", "renewal_probability") or key.startswith("exclude_item:"):
            rule = ws.Range(f"K{row}:M{row}").Validation
            rule.Delete()
            rule.Add(2 if key == "renewal_probability" else 1, 1, 1, "0", "1")
    for group, label, sheet, address, _ in m.inputs:
        if sheet == ASSUMPTIONS and group == "단가 연동 보조":
            rule = ws.Range(address).Validation
            rule.Delete()
            if label.endswith(("연동률", "비율")):
                rule.Add(2, 1, 1, "0", "1")
            else:
                rule.Add(2, 1, 7, "0")
            rule.ShowError = True
        elif sheet == ASSUMPTIONS and label.endswith("확률") and group != "시나리오":
            rule = ws.Range(address).Validation
            rule.Delete()
            rule.Add(2, 1, 1, "0", "1")
        elif sheet == ASSUMPTIONS and label == "소급 인식 방식 (0 발생, 1 일시)":
            rule = ws.Range(address).Validation
            rule.Delete()
            rule.Add(3, 1, 1, "0,1")
            rule.ShowError = True
        elif sheet == ASSUMPTIONS and label == "계약 갱신 기간 (개월)":
            rule = ws.Range(address).Validation
            rule.Delete()
            rule.Add(1, 1, 7, "1")  # whole number >= 1
        elif sheet == ASSUMPTIONS and label in TAX_PAID.values():
            rule = ws.Range(address).Validation
            rule.Delete()
            rule.Add(2, 1, 7, "0")  # decimal >= 0, blank inherits cutoff balance
