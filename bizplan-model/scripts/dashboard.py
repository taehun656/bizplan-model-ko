"""Sensitivity controls and management results for the plan workbook."""
from openpyxl.utils import get_column_letter as letter
from workbook_style import set_number_format, set_numeric_list_format
from design_palette import BLACK, BLUE, DARK, FONT, NUMBER_FONT, ORANGE, RED, SILVER, WHITE, YELLOW, rgb


SENS = "Sensitivity"
RESULT = "민감도및결과"
ASSUMPTIONS = "전제사항"
SCENARIOS = ("Base", "Up", "Down")
MULTIPLE = '0.00"x";[Red](0.00"x")'
PERCENT = "0.0%;[Red](0.0%)"
MONTH = "yyyy-mm"
MONEY = '#,##0.00####;[Red](#,##0.00####);"-"'
UNIT_LABEL = '=IF(Unit=100000000,"억원",IF(Unit=1000000,"백만원",IF(Unit=10000,"만원",IF(Unit=1000,"천원",TEXT(Unit,"#,##0")&"원"))))'
LIVE = (
    ("매출", "revenue"), ("EBITDA", "ebitda"), ("영업이익", "op"),
    ("순이익", "net"), ("기말 현금", "cash"), ("기말 차입금", "debt"),
    ("신규 조달 필요액 최고", "new_peak"), ("DSCR (계획 구간)", "dscr"),
)


def _section(m, sheet, row, title):
    m.cell(sheet, row, 2, title, "text")
    m.section(sheet, row)


def scenario_controls(line, scenarios=None, depreciation_accounts=()):
    """Yield editable Base/Up/Down controls in display order."""
    if line.get("method") and set(depreciation_accounts).intersection(line["accounts"]):
        return
    driver = line.get("driver")
    yield "multipliers", "수량 배수" if driver == "price_volume" else "배수", 1
    yield "additions", "월 가산액", 0
    if driver == "activity":
        yield "activity_multiplier", "활동량 배수", 1
        yield "yield_multiplier", "활동당 수량 배수", 1
    if line.get("method") == "consumption":
        yield "efficiency_multiplier", "효율 배수", 1
        yield "subsidy_multiplier", "단위 보조 배수", 1
    if driver in ("price_volume", "activity") or line.get("method") == "consumption":
        yield "price_multiplier", "단가 배수", 1
    if driver in ("cohort", "contracts"):
        yield "new_multiplier", "신규 고객 수 배수" if driver == "cohort" else "신규 월 매출 배수", 1
        yield "exclude_unconfirmed", "미확정 신규 고객 제외" if driver == "cohort" else "미확정 계약 제외", 0


    if driver == "contracts" and "renewal" in line.get("params", {}):
        yield "renewal_probability", "시나리오 갱신 확률", None
    if driver == "contracts" and scenarios:
        names = {name for scenario in scenarios.values()
                 for name in scenario.get("drivers", {}).get(line["id"], {}).get("exclude_items", [])}
        for name in sorted(names):
            yield "exclude_item:" + name, name + " 계약 제외", 0


def configure_sensitivity(m):
    """Allocate the named macro ranges and driver references before building flows."""
    end = letter(4 + len(m.cases))
    m.active_factors, m.active_additions, m.scenario_refs = {}, {}, {}
    from assumption_layout import initialize_tax_refs
    initialize_tax_refs(m)
    m.active_driver_switches = {}
    m.consumption_price_rows = {}
    m.live_labels = [label for label, key in LIVE]
    m.named("SN_Case", SENS, "$C$4")
    m.named("SN_LastRun", SENS, "$C$6")
    m.named("SN_CaseNos", SENS, f"$E$13:${end}$13")
    m.cell(SENS, 4, 2, "선택 케이스 번호", "text")
    m.cell(SENS, 4, 3, 1, "input")
    m.register(SENS, "C4", "케이스 선택", "민감도", "계산할 케이스 번호")
    m.cell(SENS, 4, 5, f"=INDEX(E14:{end}14,1,SN_Case)")
    m.cell(SENS, 5, 2, "선택 케이스 저장 상태", "text")
    m.cell(SENS, 6, 2, "최근 민감도 갱신", "text")
    # The template's two macro buttons occupy rows 7 and 8.
    m.table(SENS, 9, 2, [
        ["단추", "언제 누르나", "하는 일", "바뀌는 곳"],
        ["민감도 갱신", "가정을 수정한 후", "모든 케이스 재계산", "케이스별 저장값"],
        ["기준값 갱신", "계획을 확정할 때", "세 안의 기준값 저장", "기준값갱신 시트"],
    ])
    m.header(SENS, 9)
    _section(m, SENS, 12, "1. 케이스 설정")
    for row, label in ((13, "케이스 번호"), (14, "케이스명"), (15, "기준 시나리오 번호")):
        m.cell(SENS, row, 2, label, "text")
    m.header(SENS, 14)
    m.cell(SENS, 15, 3, f"=INDEX(E15:{end}15,1,SN_Case)")
    for col, case in enumerate(m.cases, 5):
        m.cell(SENS, 13, col, col - 4)
        m.cell(SENS, 14, col, case["label"], "text")
        m.cell(SENS, 15, col, SCENARIOS.index(case["scenario"]) + 1)
    for col, scenario in enumerate(SCENARIOS, 10):
        m.cell(ASSUMPTIONS, 12, col, scenario, "text")

    changed, unchanged = [], []
    driver_source = 14 + 2 * len(m.lines)
    for i, line in enumerate(m.lines):
        ident = line["id"]
        for offset, (key, caption, default) in enumerate(scenario_controls(
                line, m.spec["scenarios"], m.spec["investment"]["depreciation_accounts"])):
            legacy = key in ("multipliers", "additions")
            if legacy:
                source = 14 + 2 * i + offset
                base = [m.spec["scenarios"][s].get(key, {}).get(ident, default) for s in SCENARIOS]
            else:
                source = driver_source
                driver_source += 1
                base = [m.spec["scenarios"][s].get("drivers", {}).get(ident, {}).get(key, default)
                        for s in SCENARIOS]
                if key.startswith("exclude_item:"):
                    item = key.removeprefix("exclude_item:")
                    base = [int(item in m.spec["scenarios"][s].get("drivers", {}).get(ident, {}).get("exclude_items", []))
                            for s in SCENARIOS]
                elif key == "exclude_unconfirmed":
                    base = [int(v) for v in base]
            label = line["label"] + " " + caption
            m.cell(ASSUMPTIONS, source, 2, label, "text")
            for col, value in enumerate(base, 10):
                m.cell(ASSUMPTIONS, source, col, value, "input")
            m.register(ASSUMPTIONS, f"J{source}:L{source}", label, "시나리오", "세 시나리오의 기본 가정")
            m.scenario_refs[ident, key] = source
            # Custom cases inherit driver switches from their selected scenario.
            # Only line-wide multipliers and additions have custom adjustments.
            overrides = [case.get(key, {}).get(ident, default) for case in m.cases[3:]] if legacy else []
            effective = base.copy()
            if key == "renewal_probability":
                inherited = line["params"]["renewal"]["probability"]
                effective = [inherited if v is None else v for v in base]
            for case, adjustment in zip(m.cases[3:], overrides):
                scenario_value = base[SCENARIOS.index(case["scenario"])]
                effective.append(scenario_value + adjustment if key == "additions" else scenario_value * adjustment)
            group = changed if len(set(effective)) > 1 else unchanged
            group.append((ident, key, label, source, overrides))

    m.table(SENS, 16, 2, [["항목", "현재 적용값", "단위"]])
    m.header(SENS, 16)
    if not changed:
        m.cell(SENS, 17, 2, "케이스별로 다른 가정이 없습니다.", "text")
    result_title = 18 + max(len(changed), 1)
    first = result_title + 2
    last = first + len(LIVE) - 1
    _section(m, SENS, result_title, "2. 결과")
    m.table(SENS, first - 1, 2, [["계획 연도 지표", "현재 계산값", "현재 − 저장"]])
    m.header(SENS, first - 1)
    for col in range(5, 5 + len(m.cases)):
        m.cell(SENS, first - 1, col, f"={letter(col)}14")
    m.named("SN_Live", SENS, f"$C${first}:$C${last}")
    m.named("SN_Store", SENS, f"$E${first}:${end}${last}")
    _section(m, SENS, last + 2, "3. 사용 안내")
    for r, note in enumerate((
        "케이스 번호를 고르면 현재 적용값과 손익·자금 전망이 바뀝니다.",
        "세 안의 기본 가정은 전제사항에서, 추가 케이스의 조정값은 아래 전체 스위치에서 입력합니다.",
        "금액은 표시 단위입니다. 현재 − 저장은 선택한 케이스의 마지막 저장값과 비교합니다.",
    ), last + 3):
        m.cell(SENS, r, 2, note, "text")
    fold_title = last + 7
    _section(m, SENS, fold_title, "전체 스위치")
    hidden_start = fold_title + 1
    override_title = hidden_start + len(unchanged)
    if len(m.cases) > 3:
        m.cell(SENS, override_title, 2, "추가 케이스 조정값", "text")
        for j in range(8, 5 + len(m.cases)):
            m.cell(SENS, override_title, j, f"={letter(j)}14")
        m.header(SENS, override_title)
    hidden_last = override_title - 1
    consumption = {line["id"] for line in m.spec["cost_lines"] if line["method"] == "consumption"}
    factor_rows = []
    override_row = override_title + 1
    for i, (ident, key, label, source, overrides) in enumerate(changed + unchanged):
        row = 17 + i if i < len(changed) else hidden_start + i - len(changed)
        addition = key == "additions"
        factor = key == "multipliers" or key.endswith("_multiplier")
        unit = UNIT_LABEL if addition else "배" if factor else "%" if key == "renewal_probability" else "0 / 1"
        op = "*" if factor else "+"
        m.cell(SENS, row, 2, label, "text")
        m.cell(SENS, row, 4, unit)
        active = f"'Sensitivity'!$C${row}"
        if ident in consumption and key == "price_multiplier":
            m.consumption_price_rows[ident] = row
        if key == "multipliers":
            m.active_factors[ident] = active
        elif not addition:
            m.active_driver_switches[ident, key] = active
        # All effective case columns show final assumptions in display units.
        # The assumptions builder relocates the source INDEX and restores won.
        for col in range(5, 5 + len(m.cases)):
            c = letter(col)
            formula = f"INDEX(전제사항!J{source}:L{source},1,{c}$15)"
            if addition:
                formula += "/Unit"
            if col >= 8 and overrides:
                formula += f"{op}{c}{override_row}"
            m.cell(SENS, row, col, "=" + formula)
        m.cell(SENS, row, 3, f"=INDEX(E{row}:{end}{row},1,SN_Case)")
        if addition:
            # Drivers consume raw won. Keep the displayed active cell comparable
            # to the case columns and expose a separate won conversion.
            won_col = letter(5 + len(m.cases))
            m.cell(SENS, row, 5 + len(m.cases), f"=C{row}*Unit")
            m.active_additions[ident] = f"'Sensitivity'!${won_col}${row}"
        elif factor:
            factor_rows.append(row)
        if overrides:
            target = override_row
            override_row += 1
            m.cell(SENS, target, 2, label, "text")
            m.cell(SENS, target, 4, unit)
            for j, value in enumerate(overrides, 8):
                m.cell(SENS, target, j, value if factor else value / m.spec["unit"], "input")
            m.register(SENS, f"H{target}:{end}{target}", label, "민감도", "기준 시나리오에 곱할 배수 또는 더할 월 가산액")
            if factor:
                factor_rows.append(target)
            hidden_last = target
    if hidden_last < hidden_start:
        m.cell(SENS, hidden_start, 2, "모든 스위치가 위 케이스 설정에 표시되어 있습니다.", "text")
    m.dashboard_sensitivity = (first, last, hidden_start, hidden_last, factor_rows)


def add_consumption_price_inputs(m):
    """Allow every case to override its inherited purchasing price multiplier."""
    if not m.consumption_price_rows:
        return
    top = max(row for row, col in m.grid[SENS]) + 3
    last_col = letter(4 + len(m.cases))
    m.section(SENS, top, "구매 단가 배수 직접 입력")
    m.cell(SENS, top + 1, 2, "빈칸은 전제사항의 시나리오 배수를 따릅니다. 5% 인상은 1.05를 입력하세요.", "text")
    headers = ["비용", "현재 적용값", "단위", *[case["label"] for case in m.cases]]
    m.table(SENS, top + 2, 2, [headers])
    m.header(SENS, top + 2)
    labels = {line["id"]: line["label"] for line in m.spec["cost_lines"]}
    for index, (ident, switch) in enumerate(m.consumption_price_rows.items(), 1):
        row = top + 2 + index
        m.cell(SENS, row, 2, labels[ident], "text")
        m.cell(SENS, row, 3, f"=C{switch}")
        m.cell(SENS, row, 4, "배", "text")
        addr = f"E{row}:{last_col}{row}"
        m.register(SENS, addr, labels[ident] + " 케이스 구매 단가 배수", "민감도",
                   "빈칸은 시나리오 기본값, 숫자는 해당 케이스의 최종 구매 단가 배수")
        m.named(f"SN_ConsumptionPrice_{index}", SENS, f"$E${row}:${last_col}${row}")
        for col in range(5, 5 + len(m.cases)):
            cell = f"{letter(col)}{row}"
            inherited = m.grid[SENS][switch, col][1:]
            m.cell(SENS, switch, col, f"=IF(ISBLANK({cell}),{inherited},IF(AND(ISNUMBER({cell}),{cell}>=0),{cell},NA()))")
        m.format_row(SENS, row, MULTIPLE)


ANNUAL = (
    ("매출", "revenue"), ("매출원가", "cogs"), ("매출총이익", "gross"),
    ("판매관리비", "sga"), ("영업이익", "op"), ("영업외수익", "nonop_income"),
    ("영업외비용", "nonop_expense"), ("이자비용", "interest"),
    ("세전이익", "pretax"), ("과세표준", "tax_base"), ("법인세비용", "tax"),
    ("실효세율", "effective_tax_rate"), ("기말 이월결손금", "nol_close"), ("당기순이익", "net"),
    ("감가상각비", "dep"), ("EBITDA", "ebitda"), ("매출총이익률", "gross_margin"),
    ("영업이익률", "op_margin"), ("EBITDA율", "ebitda_margin"),
    ("기말 현금", "cash"), ("최저 월말 현금", "cash_min"),
    ("최저 현금 월", "cash_month"), ("기말 차입금", "debt"),
    ("한도대출 최고 잔액", "revolver"), ("신규 조달 필요액 최고", "new_peak"),
    ("신규 조달 필요액 최고 월", "new_month"), ("DSCR (계획 구간)", "dscr"),
)
MARGINS = {"gross_margin": "gross", "op_margin": "op", "ebitda_margin": "ebitda"}


def _annual_formula(m, key, indices, col, rows, source):
    start, end = m.col(indices[0]), m.col(indices[-1])
    first, last = source + indices[0], source + indices[-1]
    if key in MARGINS:
        return f'=IF({col}{rows["revenue"]}=0,"",{col}{rows[MARGINS[key]]}/{col}{rows["revenue"]})'
    if key == "dividends":
        region = f"'CF'!{start}{m.dividend_cash_row}:{end}{m.dividend_cash_row}"
        return f'=IF(COUNT({region}),-SUM({region}),"자료 없음")'
    if key == "effective_tax_rate":
        return f'=IF({col}{rows["pretax"]}=0,"",{col}{rows["tax"]}/{col}{rows["pretax"]})'
    if key in ("tax_base", "nol_close", "credit_used_change", "credit_close"):
        if key in ("tax_base", "credit_used_change"):
            r = m.tax_rows["tax_base_change" if key == "tax_base" else key]
            region = f"'IS'!{start}{r}:{end}{r}"
            count = f"COUNT({region})"
            if key == "credit_used_change":
                count += f"=COLUMNS({region})"
            return f'=IF({count},SUM({region}),"자료 없음")'
        cell = f"'IS'!{end}{m.tax_rows[key]}"
        return f'=IF(ISNUMBER({cell}),{cell},"자료 없음")'
    if key in m.is_rows:
        row = m.is_rows[key]
        return f"=SUM('IS'!{start}{row}:{end}{row})"
    if not m.bs and indices[-1] <= m.last and key in (
        "cash", "debt", "cash_min", "cash_month"
    ):
        return '="자료 없음"'
    if key in ("cash", "debt"):
        source_col = "G" if key == "cash" else "H"
        return f'=IF(ISNUMBER({source_col}{last}),{source_col}{last},"")'
    if key in ("cash_min", "revolver", "new_peak"):
        c, function = {"cash_min": ("G", "MIN"), "revolver": ("I", "MAX"),
                       "new_peak": ("J", "MAX")}[key]
        area = f"{c}{first}:{c}{last}"
        return f'=IF(COUNT({area}),{function}({area}),"")'
    if key in ("cash_month", "new_month"):
        c, target = ("G", "cash_min") if key == "cash_month" else ("J", "new_peak")
        value = f"{col}{rows[target]}"
        predicate = f"ISNUMBER({value})" if key == "cash_month" else f"AND(ISNUMBER({value}),{value}>0)"
        return f'=IF({predicate},INDEX(F{first}:F{last},MATCH({value},{c}{first}:{c}{last},0)),"")'
    # DSCR uses aggregate forecast cash flow and debt service within the year.
    projected = [i for i in indices if i > m.last]
    if not projected:
        return '="해당 없음"'
    start = m.col(projected[0])
    def summed(sheet, row):
        return f"SUM('{sheet}'!{start}{row}:{end}{row})"
    interest = summed("재원조달", 13)
    service = "+".join((summed("재원조달", 11), interest, summed("재원조달", 17)))
    return f'=IF(({service})>0,({summed("CF", 13)}+{interest})/({service}),"해당 없음")'


ANNUAL_TITLE = 24
YEAR_ROW = ANNUAL_TITLE + 1
KIND_ROW = ANNUAL_TITLE + 2
FIRST_METRIC = ANNUAL_TITLE + 3
CARDS = (
    ("B", "B", "시나리오", None), ("C", "D", "매출", "revenue"),
    ("E", "F", "EBITDA", "ebitda"), ("G", "H", "영업이익", "op"),
    ("I", "I", "기말 현금", "cash"), ("J", "K", "DSCR", "dscr"),
    ("L", "M", "신규 조달 필요액 최고", "new_peak"),
)


def build_cards(m, years, latest, annual_rows):
    year, months = years[-1]
    previous = int(year) - 1
    comparable = len(years) > 1 and (
        [m.months[i][5:] for i in years[-2][1]] == [m.months[i][5:] for i in months]
    )
    comparison = f"증감은 {previous}년 대비" if comparable else "연간 증감 비교 제외"
    summary = f'="{year}년 핵심 지표  ·  {comparison}  ·  단위 "&{UNIT_LABEL[1:]}'
    m.cell(RESULT, 3, 2, summary)
    count = len(m.checks)
    m.cell(RESULT, 3, 12, f'="모델 검증 "&COUNTIF(검증!G8:G{7+count},"OK")&"/{count} OK"')
    for left, _, label, key in CARDS:
        col = ord(left) - ord("A") + 1
        if key is None:
            m.cell(RESULT, 5, col, label, "text")
            m.cell(RESULT, 6, col, "=Sensitivity!E4")
            m.cell(RESULT, 7, col, "Sensitivity 시트에서 선택", "text")
            continue
        current = f"{latest}{annual_rows[key]}"
        title = label
        if key == "dscr":
            title = f'=IF(AND(ISNUMBER({current}),{current}<1),"DSCR · 기준 1.0배 미달","DSCR")'
        m.cell(RESULT, 5, col, title)
        m.cell(RESULT, 6, col, f'=IF({current}="","",{current})')
        if comparable:
            prior = f"{letter(1 + 2 * (len(years)-1))}{annual_rows[key]}"
            m.cell(RESULT, 7, col, f'=IF(COUNT({prior},{current})=2,{current}-{prior},"")')


def finish_dashboard(m):
    """Build results after financial rows and the complete check list are known."""
    years = list(m.years.items())
    plan_year = years[-1][0]
    latest = letter(3 + 2 * (len(years) - 1))
    metrics = ANNUAL + ((("배당 지급액", "dividends"),) if m.dividend_cash_row else ())
    if "credit_close" in m.tax_rows:
        metrics += (("세액공제 사용액", "credit_used_change"), ("기말 이월세액공제", "credit_close"))
    annual_rows = {key: row for row, (label, key) in enumerate(metrics, FIRST_METRIC)}
    first, last, *_ = m.dashboard_sensitivity
    base_title = FIRST_METRIC + len(metrics) + 1
    custom_title = base_title + 3 + len(LIVE)
    check_title = custom_title + (3 + len(LIVE) if len(m.cases) > 3 else 0)
    if m.activity_lines:
        from activity_model import results_table
        check_title = results_table(m, check_title)
    half = (len(m.checks) + 1) // 2
    chart_row = check_title + half + 3
    source_header = chart_row + 19
    source = source_header + 1
    m.dashboard_results = (base_title, custom_title, check_title, half, chart_row, source_header)
    build_cards(m, years, latest, annual_rows)
    _section(m, RESULT, ANNUAL_TITLE, "연도별 손익 및 자금")
    m.cell(RESULT, KIND_ROW, 2, "항목", "text")
    m.header(RESULT, KIND_ROW)
    for label, key in metrics:
        row = annual_rows[key]
        m.cell(RESULT, row, 2, label, "text")
        if key in ("gross", "op", "net", "ebitda"):
            m.total(RESULT, row)
    for j, (year, indices) in enumerate(years):
        number = 3 + 2 * j
        col = letter(number)
        classification = m.year_kind(indices)
        m.cell(RESULT, YEAR_ROW, number, int(year))
        m.cell(RESULT, KIND_ROW, number, classification, "text")
        m.cell(RESULT, KIND_ROW, number + 1, "전년 대비", "text")
        for label, key in metrics:
            row = annual_rows[key]
            m.cell(RESULT, row, number, _annual_formula(m, key, indices, col, annual_rows, source))
            if not j or key in ("cash_month", "new_month"):
                continue
            # Avoid comparing a partial year's flows against a full year's.
            prior_months = [m.months[i][5:] for i in years[j-1][1]]
            if prior_months != [m.months[i][5:] for i in indices]:
                continue
            p, v = f"{letter(number - 2)}{row}", f"{col}{row}"
            if key in MARGINS or key == "effective_tax_rate":
                delta = f'IF(COUNT({p},{v})=2,100*({v}-{p}),"")'
            else:
                delta = f'IF(AND(COUNT({p},{v})=2,{p}<>0),({v}-{p})/ABS({p}),"")'
            m.cell(RESULT, row, number + 1, "=" + delta)
    _section(m, RESULT, source_header - 1, "차트 원천 (수정하지 않음)")
    m.table(RESULT, source_header, 6, [["월", "기말 현금", "기말 차입금", "한도 사용액", "신규 조달 누적액"]])
    m.header(RESULT, source_header)
    for i in range(len(m.months)):
        col, row = m.col(i), source + i
        m.cell(RESULT, row, 6, f"='INDEX'!{col}5")
        # Historical cash/debt can be unavailable. Preserve gaps in charts and extrema.
        for target, sheet, r in ((7, "CF", 8), (8, "재원조달", 8), (10, "재원조달", 9)):
            if (r, 10 + i) in m.grid[sheet]:
                m.cell(RESULT, row, target, f"='{sheet}'!{col}{r}")
        if any((r, 10 + i) in m.grid["재원조달"] for r in (18, 21)):
            m.cell(RESULT, row, 9, f"=SUM('재원조달'!{col}18,'재원조달'!{col}21)")

    matches = []
    for index, (label, key) in enumerate(LIVE, 1):
        row = first + index - 1
        actual = f"C{row}"
        stored = f"INDEX(SN_Store,{index},SN_Case)"
        m.cell(SENS, row, 2, label, "text")
        m.cell(SENS, row, 3, f"='{RESULT}'!{latest}{annual_rows[key]}")
        for col in range(5, 5 + len(m.cases)):
            m.cell(SENS, row, col, 0)
        equal = f"EXACT({actual},{stored})"
        matches.append(f"IF(COUNT({actual},{stored})=2,ABS({actual}-{stored})<=1E-9,{equal})")
        m.cell(SENS, row, 4, f'=IF(ISBLANK(SN_LastRun),"",IF(COUNT({actual},{stored})=2,{actual}-{stored},IF({equal},"일치","불일치")))')
    m.cell(SENS, 5, 3, '=IF(ISBLANK(SN_LastRun),"미저장",IF(AND(' + ",".join(matches) + '),"일치","갱신 필요"))')

    tables = [(base_title, "세 안 비교", range(3))]
    if len(m.cases) > 3:
        tables.append((custom_title, "조건별 민감도", range(3, len(m.cases))))
    for title, heading, cases in tables:
        _section(m, RESULT, title, f"{plan_year}년 {heading}")
        m.cell(RESULT, title + 1, 2, "저장 지표", "text")
        m.header(RESULT, title + 1)
        for col, case_index in enumerate(cases, 3):
            m.cell(RESULT, title + 1, col, f"='Sensitivity'!{letter(5+case_index)}14")
            for index, label in enumerate(m.live_labels, 1):
                row = title + 1 + index
                m.cell(RESULT, row, 2, label, "text")
                m.cell(RESULT, row, col, f'=IF(ISBLANK(SN_LastRun),"",INDEX(SN_Store,{index},{case_index+1}))')
    _section(m, RESULT, check_title, "모델 검증")
    for i, (name, _) in enumerate(m.checks):
        row = check_title + 1 + i % half
        col = 2 if i < half else 5
        m.cell(RESULT, row, col, name, "text")
        m.cell(RESULT, row, col + 2, f"='검증'!G{8+i}")


    from results_sources import build_sources
    build_sources(m, annual_rows, source_header)


def _macro_buttons(ws):
    """Use macro-linked rectangles because legacy Form Control faces follow Windows colors."""
    for name, label, action, anchor in (
        ("btnRefreshSensitivity", "민감도 갱신", "RefreshSensitivity", "B7"),
        ("btnRefreshBaselines", "기준값 갱신", "RefreshBaselines", "C7"),
    ):
        ws.Shapes(name).Delete()
        cell = ws.Range(anchor)
        button = ws.Shapes.AddShape(1, cell.Left, cell.Top + 3, 100, 24)
        button.Name = name
        button.OnAction = action
        button.Placement = 2  # Move with the button row without stretching.
        button.Fill.Solid()
        button.Fill.ForeColor.RGB = rgb(ORANGE)
        button.Line.ForeColor.RGB = rgb(BLACK)
        button.Line.Visible = 0
        button.Shadow.Visible = 0
        frame = button.TextFrame2
        frame.VerticalAnchor = 3
        text = frame.TextRange
        text.Text = label
        text.ParagraphFormat.Alignment = 2
        text.Font.Name = NUMBER_FONT
        text.Font.NameAscii = NUMBER_FONT
        text.Font.NameFarEast = FONT
        text.Font.Size = 10
        text.Font.Bold = -1
        text.Font.Fill.ForeColor.RGB = rgb(WHITE)


def finalize_button_styles(path):
    """Set native button fallback colors and chart script fonts in the staged file."""
    from xml.dom import minidom
    from zipfile import ZipFile

    drawing = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
    art = "http://schemas.openxmlformats.org/drawingml/2006/main"
    fallback = {"lnRef": BLACK, "fillRef": ORANGE, "effectRef": BLACK, "fontRef": WHITE}
    output = path.with_suffix(".styled.xlsm")
    with ZipFile(path) as original, ZipFile(output, "w") as styled:
        for member in original.infolist():
            data = original.read(member)
            if member.filename.startswith("xl/drawings/drawing") and member.filename.endswith(".xml"):
                doc = minidom.parseString(data)
                changed = False
                for shape in doc.getElementsByTagNameNS(drawing, "sp"):
                    props = shape.getElementsByTagNameNS(drawing, "cNvPr")[0]
                    if props.getAttribute("name") not in ("btnRefreshSensitivity", "btnRefreshBaselines"):
                        continue
                    # Excel supplies a private fallback style even after formatting Fill,
                    # Line and TextFrame2. The COM API exposes no setter for these refs.
                    for theme in shape.getElementsByTagNameNS(art, "schemeClr"):
                        reference = theme.parentNode
                        if reference.localName in fallback:
                            explicit = doc.createElementNS(art, f"{theme.prefix}:srgbClr")
                            explicit.setAttribute("val", fallback[reference.localName])
                            reference.replaceChild(explicit, theme)
                            changed = True
                if changed:
                    data = doc.toxml(encoding="UTF-8")
                doc.unlink()
            elif member.filename.startswith("xl/charts/chart") and member.filename.endswith(".xml"):
                from results_charts import finalize_chart_fonts
                data = finalize_chart_fonts(data)
            styled.writestr(member, data)
    output.replace(path)


def _table_style(ws, area):
    borders = ws.Range(area).Borders
    borders.LineStyle = 1
    borders.Weight = 2
    borders.Color = rgb(BLACK)


def set_result_print_area(ws, year_count, last_annual_row):
    """Print the summary cards, charts and complete annual table on one page across."""
    last_col = letter(max(13, 2 + 2 * year_count))
    setup = ws.PageSetup
    setup.PrintArea = f"A1:{last_col}{last_annual_row}"
    setup.Orientation = 2  # xlLandscape
    setup.Zoom = False
    setup.FitToPagesWide = 1
    setup.FitToPagesTall = False


def style_dashboard(ws, m):
    """Apply the presentation and controls after the engine writes worksheet cells."""
    if ws.Name not in (SENS, RESULT):
        return
    ws.Columns("A:A").ColumnWidth = 3
    ws.Columns("B:B").ColumnWidth = 34
    ws.Rows("1:2").RowHeight = 20
    if ws.Name == SENS:
        first, last, start, stop, factors = m.dashboard_sensitivity
        end = letter(4 + len(m.cases))
        ws.Columns(f"C:{end}").ColumnWidth = 19
        ws.Columns("C:C").ColumnWidth = 22
        helper = letter(5 + len(m.cases))
        ws.Columns(f"{helper}:{helper}").Hidden = True
        ws.Range("C6").NumberFormat = "yyyy-mm-dd hh:mm"
        set_numeric_list_format(ws.Range("C4"))
        validation = ws.Range("C4").Validation
        validation.Delete()
        validation.Add(3, 1, 1, "=SN_CaseNos")
        validation.IgnoreBlank = False
        validation.ShowError = True
        ws.Range("C4").Interior.Color = rgb(YELLOW)
        ws.Range("C4").Font.Color = rgb(BLUE)
        for index in range(1, len(m.consumption_price_rows) + 1):
            _, address = m.names[f"SN_ConsumptionPrice_{index}"]
            band = ws.Range(address)
            validation = band.Validation
            validation.Delete()
            validation.Add(2, 1, 7, 0)  # decimal >= 0, blank means inherit
            validation.IgnoreBlank = True
            validation.ShowError = True
            set_number_format(ws.Range(f"C{band.Row}:{end}{band.Row}"), MULTIPLE)
        _macro_buttons(ws)
        status = ws.Range("C5").FormatConditions.Add(2, 1, '=OR($C$5="갱신 필요",$C$5="미저장")')
        status.Font.Color = rgb(RED)
        status.Font.Bold = True
        mismatch = ws.Range(f"D{first}:D{last}").FormatConditions.Add(1, 3, '="불일치"')
        mismatch.Font.Color = rgb(RED)
        mismatch.Font.Bold = True
        _table_style(ws, "B9:E11")
        _table_style(ws, f"B13:{end}{first - 4}")
        _table_style(ws, f"B{first-1}:{end}{last}")
        # The entire active case is visually consistent across settings and stores.
        for area in [f"E13:{end}{first-4}", f"E{first-1}:{end}{last}"]:
            rule = ws.Range(area).FormatConditions.Add(2, 1, "=E$13=$C$4")
            rule.Interior.Color = rgb(ORANGE)
            rule.Font.Color = rgb(WHITE)
            rule.Font.Bold = True
        if stop >= start:
            rule = ws.Range(f"E{start}:{end}{stop}").FormatConditions.Add(2, 1, "=E$13=$C$4")
            rule.Interior.Color = rgb(ORANGE)
            rule.Font.Color = rgb(WHITE)
            rule.Font.Bold = True
            ws.Outline.SummaryRow = 0
            rows = ws.Rows(f"{start}:{stop}")
            rows.Group()
            rows.Hidden = True
        for (row, col), value in m.grid[SENS].items():
            if col == 4 and value == "'0 / 1":
                ws.Range(f"C{row}:{end}{row}").NumberFormat = "0"
            if col == 4 and value == "'%":
                set_number_format(ws.Range(f"C{row}:{end}{row}"), PERCENT)
            if col == 4 and value == UNIT_LABEL:
                set_number_format(ws.Range(f"C{row}:{end}{row}"), MONEY)
        for row in factors:
            set_number_format(ws.Range(f"C{row}:{end}{row}"), MULTIPLE)
        set_number_format(ws.Range(f"C{last}:{end}{last}"), MULTIPLE)
        ws.Range(f"E14:{end}14").WrapText = True
        ws.Rows("14:14").AutoFit()
        return

    base, custom, checks, half, chart_row, header = m.dashboard_results
    set_result_print_area(ws, len(m.years), base - 2)
    end = letter(max(13, 2 + 2 * len(m.years), len(m.cases) - 1))
    ws.Columns(f"C:{end}").ColumnWidth = 16
    style_cards(ws, m)
    for j in range(len(m.years)):
        col, delta = letter(3 + 2*j), letter(4 + 2*j)
        ws.Range(f"{col}{YEAR_ROW}").NumberFormat = "0"
        set_number_format(ws.Range(f"{delta}{FIRST_METRIC}:{delta}{base-2}"), PERCENT)
        for row, (_, key) in enumerate(ANNUAL, FIRST_METRIC):
            if key in MARGINS or key == "effective_tax_rate":
                set_number_format(ws.Range(f"{col}{row}"), PERCENT)
                set_number_format(ws.Range(f"{delta}{row}"), '0.0"%p"')
            elif key in ("cash_month", "new_month"):
                ws.Range(f"{col}{row}").NumberFormat = MONTH
            elif key == "dscr":
                set_number_format(ws.Range(f"{col}{row}"), MULTIPLE)
    tables = [(base, 5)]
    if len(m.cases) > 3:
        tables.append((custom, len(m.cases) - 1))
    for title, last_col in tables:
        end_col, bottom = letter(last_col), title + 1 + len(LIVE)
        _table_style(ws, f"B{title+1}:{end_col}{bottom}")
        ws.Range(f"C{title+1}:{end_col}{title+1}").WrapText = True
        ws.Rows(f"{title+1}:{title+1}").AutoFit()
        set_number_format(ws.Range(f"C{bottom}:{end_col}{bottom}"), MULTIPLE)
    # Each check caption spans its two reserved columns, including long captions.
    for left, right in (("B", "C"), ("E", "F")):
        labels = ws.Range(f"{left}{checks+1}:{right}{checks+half}")
        labels.Merge(True)  # Across each row, preserving each distinct caption.
        labels.ShrinkToFit = True
    for col in ("D", "G"):
        if half:
            rule = ws.Range(f"{col}{checks+1}:{col}{checks+half}").FormatConditions.Add(1, 3, '="NO"')
            rule.Font.Color = rgb(RED)
            rule.Font.Bold = True
    ws.Range(f"F{header+1}:F{header+len(m.months)}").NumberFormat = MONTH
    ws.Rows(f"{chart_row}:{header-1}").RowHeight = 15


def style_cards(ws, m):
    area = ws.Range("A1:M24")
    area.Font.Name = FONT
    area.Font.Size = 9
    # Reserve the first row for the workbook name and the second for the sheet.
    for row, height, size, fill in (
            (1, 20, 10, ORANGE), (2, 26, 14, SILVER)):
        band = ws.Range(f"A{row}:M{row}")
        band.UnMerge()
        band.Merge()
        band.Interior.Color = rgb(fill)
        band.Font.Color = rgb(WHITE)
        band.Font.Size = size
        band.IndentLevel = 1
        band.RowHeight = height
    last_col = max(col for _, col in m.grid[RESULT])
    if last_col > 13:
        ws.Range(f"N1:{letter(last_col)}2").Interior.ColorIndex = -4142
    ws.Range("B3:K3").Merge()
    ws.Range("L3:M3").Merge()
    ws.Range("L3:M3").Interior.Color = rgb(SILVER)
    ws.Range("B3:M3").Font.Bold = True
    ws.Range("B3:M3").IndentLevel = 1
    ws.Range("L3").HorizontalAlignment = -4152
    fail = ws.Range("L3:M3").FormatConditions.Add(
        2, 1, f'=$L$3<>"모델 검증 {len(m.checks)}/{len(m.checks)} OK"')
    fail.Interior.Color = rgb(WHITE)
    fail.Font.Color = rgb(RED)
    fail.Font.Bold = True
    for left, right, _, key in CARDS:
        for row in (5, 6, 7):
            field = ws.Range(f"{left}{row}:{right}{row}")
            if left != right:
                field.Merge()
            field.HorizontalAlignment = -4108
            field.ShrinkToFit = True
        card = ws.Range(f"{left}5:{right}7")
        for border in (7, 8, 9, 10):  # xlEdgeLeft, Top, Bottom, Right
            edge = card.Borders(border)
            edge.LineStyle = 1
            edge.Weight = 2
            edge.Color = rgb(BLACK)
        heading = ws.Range(f"{left}5:{right}5")
        heading.Interior.Color = rgb(ORANGE if key is None else SILVER)
        heading.Font.Color = rgb(WHITE if key is None else BLACK)
        heading.Font.Bold = True
        value = ws.Range(f"{left}6")
        value.Font.Bold = True
        value.Font.Size = 15
        set_number_format(value, '0.00"배";[Red](0.00"배")' if key == "dscr" else '#,##0;[Red](#,##0);0')
        note = ws.Range(f"{left}7")
        note.Font.Color = rgb(DARK)
        set_number_format(note, '"+"0.00"배";[Red]"-"0.00"배";"0.00배"' if key == "dscr"
                             else '"+"#,##0;[Red]"-"#,##0;0')
    warn = ws.Range("J6:K6").FormatConditions.Add(2, 1, '=AND(ISNUMBER($J$6),$J$6<1)')
    warn.Font.Color = rgb(RED)
    warn.Font.Bold = True
    for rows, height in (("3:3", 20), ("4:4", 6), ("5:5", 18), ("6:6", 28),
                         ("7:7", 18), ("8:8", 8), ("9:22", 17), ("23:23", 8)):
        ws.Rows(rows).RowHeight = height
    ws.Application.ActiveWindow.Zoom = 100
