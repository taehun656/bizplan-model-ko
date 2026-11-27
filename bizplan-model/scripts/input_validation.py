"""Keep editable numeric domains consistent between Excel entry and model checks."""
from collections import defaultdict
from dataclasses import dataclass, replace

from build_model import letter, ref
from workbook_style import set_numeric_list_format


@dataclass(frozen=True)
class Domain:
    minimum: int | None = None
    maximum: int | None = None
    strict: bool = False
    integer: bool = False
    blank: bool = False

    def optional(self):
        return replace(self, blank=True)

    def expression(self, cells):
        expected = f"COUNTA({cells})" if self.blank else f"ROWS({cells})*COLUMNS({cells})"
        tests = [f"COUNT({cells})={expected}"]
        if self.minimum is not None:
            operator = ">" if self.strict else ">="
            minimum = f"MIN({cells}){operator}{self.minimum}"
            if self.blank and self.strict:
                minimum = f"OR(COUNT({cells})=0,{minimum})"
            tests.append(minimum)
        if self.maximum is not None:
            tests.append(f"MAX({cells})<={self.maximum}")
        if self.integer:
            tests.append(f"SUMPRODUCT(--(MOD({cells},1)<>0))=0")
        return f'AND({",".join(tests)})'


NUMBER = Domain()
AMOUNT = Domain(0)
POSITIVE = Domain(0, strict=True)
RATE = Domain(0, 1)
GROWTH = Domain(-1)
FARE_CHANGE = Domain(-1, strict=True)
TOGGLE = Domain(0, 1, integer=True)
FACILITY = Domain(-1, 1, integer=True)
DATE = Domain(1, 2958465, integer=True)


def build_input_checks(m):
    """Watch the editable source cells, including unused cases and future months."""
    rules = {}
    relations = []
    m.purchase_price_conditions = {}
    first, last = m.col(m.last + 1), m.col(len(m.months) - 1)

    def forecast(row):
        return f"{first}{row}:{last}{row}"

    def bind(group, sheet, address, domain):
        key = sheet, address
        if key in rules and rules[key] != (group, domain):
            raise ValueError(f"중복 입력 범위 정의 {sheet}!{address}")
        rules[key] = group, domain

    for (sheet, row), address in m.scalar_input_cells.items():
        domain = None
        if sheet == "금융조건":
            domain = {12: AMOUNT, 13: AMOUNT, 14: AMOUNT, 15: RATE, 16: RATE,
                      17: AMOUNT, 18: AMOUNT, 19: AMOUNT, 22: TOGGLE, 23: TOGGLE,
                      24: TOGGLE, 25: TOGGLE}.get(row)
        elif sheet == "투자비":
            domain = {20: Domain(1, integer=True), 21: RATE, 22: TOGGLE,
                      23: AMOUNT, 24: AMOUNT, 26: TOGGLE}.get(row)
        if domain:
            bind(sheet, "전제사항", address, domain)

    bind("기준", "전제사항", "G7", Domain(1, integer=True))
    bind("투자비", "투자비", forecast(25), AMOUNT)
    bind("시나리오", "Sensitivity", "C4", Domain(1, len(m.cases), integer=True))
    for row in getattr(m, "plan_rows", ()):
        # A reserved initiative row can remain empty. A populated row needs an
        # explicit 0 or 1, even when the scenario currently excludes the plan.
        sheet = ref("영업계획")
        end = letter(13 + len(m.metric_cols))
        relation = (f'OR(COUNTA({sheet}!B{row}:{end}{row})=0,'
                    f'AND(ISNUMBER({sheet}!D{row}),OR({sheet}!D{row}=0,{sheet}!D{row}=1)))')
        relations.append(("사업부 계획", f"영업계획!D{row} · 계획 입력 행 적용 스위치", relation))
    for (_, key), row in m.scenario_refs.items():
        if key == "additions":
            policy = NUMBER
        elif key == "renewal_probability":
            policy = RATE.optional()
        elif key == "exclude_unconfirmed" or key.startswith("exclude_item:"):
            policy = TOGGLE
        elif key == "efficiency_multiplier":
            policy = POSITIVE
        else:
            policy = AMOUNT
        bind("시나리오", "전제사항", f"K{row}:M{row}", policy)

    for line in m.spec["revenue_lines"]:
        sheet = m.details[line["id"]]
        params = line["params"]
        if line["driver"] == "manual":
            policy = AMOUNT if params.get("nonnegative") else NUMBER
            bind("월 지정 금액", sheet, forecast(14), policy)
        elif line["driver"] == "activity":
            for row in (35, 36, 37):
                bind("활동 매출", sheet, forecast(row), AMOUNT.optional())
            bind("활동 매출", sheet, "G38", GROWTH)
            for row in range(41, 41 + len(params["price_steps"])):
                bind("단가 조정", sheet, f"G{row}", DATE)
                bind("단가 조정", sheet, f"H{row}", FARE_CHANGE)
                bind("단가 조정", sheet, f"I{row}", TOGGLE)
        elif line["driver"] == "price_volume":
            for row in (20, 21):
                bind("수량·단가", sheet, f"J{row}:{last}{row}", AMOUNT.optional())

    for index, asset in enumerate(m.spec["investment"]["asset_classes"]):
        base = 34 + 15 * index
        whole = asset.get("whole_units", False)
        qty, price = forecast(base + 4), forecast(base + 5)
        bind("투자비", "투자비", qty, Domain(0, integer=whole))
        bind("투자비", "투자비", price, AMOUNT)
        if whole:
            price_cell = f"{first}{base + 5}"
            qty_cell = f"{first}{base + 4}"
            m.purchase_price_conditions["투자비", price] = (
                f"=AND(ISNUMBER({price_cell}),{price_cell}>=0,OR({qty_cell}=0,{price_cell}>0))"
            )
            relations.append(("투자비", f"투자비!{price} · 취득 시 양수 단가",
                f"SUMPRODUCT(--('투자비'!{qty}>0),--('투자비'!{price}<=0))=0"))

    designated = set(m.spec["investment"]["depreciation_accounts"])
    for line in m.spec["cost_lines"]:
        if designated.intersection(line["accounts"]):
            continue
        start = m.cost_detail_rows[line["id"]]
        if line["method"] == "manual":
            bind("월 지정 금액", "운영비용", forecast(start + 4), NUMBER)
        if line["method"] == "fixed_inflation" and "monthly" in line["params"]:
            bind("월 지정 금액", "운영비용", forecast(start + 6), AMOUNT.optional())
        if line["params"].get("months"):
            bind("월 지정 금액", "운영비용", forecast(start + 7), TOGGLE)
    for address, positive in getattr(m, "cost_keyin_ranges", []):
        bind("활동 지표", "활동지표", address, POSITIVE if positive else AMOUNT)
    for group, label, sheet, address, _ in m.inputs:
        _registered_rule(bind, m, group, label, sheet, address)

    m.entry_domains = defaultdict(list)
    checks = defaultdict(list)
    for (sheet, address), (group, domain) in rules.items():
        m.entry_domains[sheet].append((address, domain))
        checks[group].append((f"{sheet}!{address}", domain.expression(f"{ref(sheet)}!{address}")))
    for group, label, expression in relations:
        checks[group].append((label, expression))
    _write_checks(m, checks)


def _registered_rule(bind, m, group, label, sheet, address):
    """Handle controls whose addresses depend on facilities, roles, metrics or extra cases."""
    domain = None
    if sheet == "금융조건" and label.endswith(" 연이율"):
        domain = RATE
    elif sheet == "전제사항" and group == "금융조건" and label.endswith(" 만기 처리"):
        domain = FACILITY
    elif group == "인건비":
        if label.endswith("연 임금 인상률"):
            domain = GROWTH
        elif label.endswith(("인상 시작 월", "소급 적용 월", "소급 정산 월")):
            domain = DATE
        elif label.startswith("소급 인식 방식"):
            domain = TOGGLE
        elif "1인 월 비용" in label:
            domain = AMOUNT
        elif sheet == "인건비" and label.endswith(" 인원"):
            domain = AMOUNT
    elif sheet == "CF":
        group = "월 지정 금액"
        domain = AMOUNT if label == "배당 지급 계획 (원)" else NUMBER
    elif sheet == "확장조정" or group == "수동 조정":
        group = "월 지정 금액"
        domain = NUMBER.optional()
    elif group == "영업외손익":
        domain = NUMBER
    elif sheet == "Sensitivity":
        if group == "사업부 계획":
            domain = TOGGLE
        elif label.endswith("케이스 구매 단가 배수"):
            domain = AMOUNT.optional()
        elif address.startswith("H") and group == "민감도":
            row = int(address.split(":")[0][1:])
            domain = AMOUNT if row in m.dashboard_sensitivity[-1] else NUMBER
    elif sheet == "활동지표":
        if label == "월별 비용 보정 (원)":
            domain = NUMBER.optional()
        elif label.endswith("월별 계수") or label == "관측 사용량":
            domain = AMOUNT.optional()
        elif label == "물량당 원가 연 증감률":
            domain = GROWTH
        elif address.startswith("G"):
            domain = POSITIVE if label.startswith("기준 효율") else AMOUNT
    if domain:
        bind(group, sheet, address, domain)


def _write_checks(m, groups):
    # Existing reconciliation rows use the monthly axis; keep their addresses stable.
    sheet = "검증"
    top = max(max(row for row, _ in m.grid[sheet]), 8 + len(m.checks) + len(groups) + 3) + 4
    m.section(sheet, top, "입력 검증 상세")
    m.table(sheet, top + 1, 2, [["입력 위치", None, None, None, None, "판정"]])
    m.header(sheet, top + 1)
    row = top + 2
    for category, entries in groups.items():
        first = row
        for label, expression in entries:
            m.cell(sheet, row, 2, label, "text")
            m.cell(sheet, row, 7, f'=IFERROR(IF({expression},"OK","NO"),"NO")')
            row += 1
        m.add_check("입력 검증 · " + category,
                    f'=IF(COUNTIF(G{first}:G{row-1},"OK")={len(entries)},"OK","NO")')


def style_input_validation(ws, m):
    for address, domain in m.entry_domains.get(ws.Name, ()):
        if ws.Name == "Sensitivity" and address == "C4":
            # style_dashboard already supplies the case selector's named list.
            continue
        validation = ws.Range(address).Validation
        validation.Delete()
        purchase_condition = m.purchase_price_conditions.get((ws.Name, address))
        if purchase_condition:
            validation.Add(7, 1, 1, purchase_condition)
        elif domain.minimum is None:
            first = address.split(":")[0]
            validation.Add(7, 1, 1, f"=ISNUMBER({first})")
        elif domain.maximum is None:
            validation.Add(1 if domain.integer else 2, 1, 5 if domain.strict else 7,
                           str(domain.minimum))
        elif domain == TOGGLE or domain == FACILITY:
            set_numeric_list_format(ws.Range(address))
            choices = "0,1" if domain == TOGGLE else "-1,0,1"
            validation.Add(3, 1, 1, choices)
        else:
            validation.Add(1 if domain.integer else 2, 1, 1, str(domain.minimum), str(domain.maximum))
        validation.IgnoreBlank = domain.blank
        validation.ShowError = True
