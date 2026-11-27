"""Class-level purchase and depreciation schedules on the existing investment sheet."""

SHEET = "투자비"


def validate_asset_classes(spec, data):
    """Reconcile class openings to the cutoff PPE pool and purchases to the forecast."""
    inv = spec["investment"]
    classes = inv.get("asset_classes", [])
    if not classes:
        return
    names = set()
    opening_net = 0
    for item in classes:
        name = item["asset_class"]
        if name in names:
            raise ValueError(f"investment.asset_classes: 중복 자산군 {name}")
        names.add(name)
        if item["method"] == "declining" and not item.get("annual_rate"):
            raise ValueError(f"{name}: 정률법의 연 상각률을 0 초과 1 이하로 지정하세요.")
        seen = set()
        for purchase in item["purchases"]:
            month = purchase["month"]
            if item.get("whole_units"):
                quantity = purchase["quantity"]
                if quantity % 1 or (quantity > 0 and purchase["unit_price"] <= 0):
                    raise ValueError(f"{name} {month}: 취득 수량은 정수, 취득 시 단가는 양수여야 합니다.")
            if month in seen or not spec["last_actual"] < month <= spec["plan_end"]:
                raise ValueError(f"{name} {month}: 자산군별 추정 월마다 취득 계획 한 건을 지정하세요.")
            seen.add(month)
        opening = item.get("opening")
        if opening:
            if opening["accumulated_dep"] > opening["cost"]:
                raise ValueError(f"{name}: 기존 자산의 감가상각누계액이 취득원가를 초과합니다.")
            if opening["remaining_life_months"] > item["life_months"]:
                raise ValueError(f"{name}: 기존 자산 잔여 월수가 내용연수를 초과합니다.")
            opening_net += opening["cost"] - opening["accumulated_dep"]
    if spec["statements"]["bs"]:
        from finance_blocks import _bucket
        total = sum(v["amount"] for v in data["bs_monthly"]
                    if v["month"] == spec["last_actual"] and _bucket(v["section"], v["account"]) == 12)
    else:
        # Detailed openings provide the total when there is no BS or separate pool.
        if not inv["opening_net_assets"]:
            inv["opening_net_assets"] = opening_net
        total = inv["opening_net_assets"]
    if opening_net > total + 0.5:
        raise ValueError("investment.asset_classes: 기존 자산군 순액 합계가 마감 유형자산 순액을 초과합니다.")


def _sum_rows(col, rows):
    return "+".join(f"{col}{r}" for r in rows) or "0"


def _new_dep_formula(col, prev, first, base, method):
    """Apply a finite life to every vintage, including additional Excel key-ins."""
    purchases = base + 6
    closing = base + 8
    span = f"${first}${purchases}:{col}{purchases}"
    elapsed = f"(COLUMNS({span})-$G$26)"
    life = f"$G${base+1}"
    available = f"({prev}{closing}+IF($G$26=0,{col}{purchases},0))"
    if method == "straight_line":
        # Include this vintage only over its depreciation months.
        start = f"INDEX({span},1,MAX(1,{elapsed}-{life}+1))"
        end = f"INDEX({span},1,{elapsed})"
        charge = f"IF({elapsed}>0,SUM({start}:{end})/{life},0)"
    else:
        rate = f"($G${base+2}/12)"
        # Remove residual book value when the vintage reaches its final month.
        maturity = f"IF({elapsed}>={life},INDEX({span},1,{elapsed}-{life}+1)*(1-{rate})^{life},0)"
        charge = f"{available}*{rate}+{maturity}"
    return f"MIN(MAX(0,{available}),{charge})"


def build_asset_classes(m, row, forecast, legacy_depreciation):
    """Combine itemized asset classes with any unallocated opening or new-asset pool."""
    classes = m.spec["investment"]["asset_classes"]
    first = m.col(m.last + 1)
    acquisition_rows, new_dep_rows, existing_dep_rows, opening_rows = [], [], [], []
    for index, item in enumerate(classes):
        base = 34 + 15 * index
        name = item["asset_class"]
        method = item["method"]
        method_label = "정액" if method == "straight_line" else "정률"
        m.section(SHEET, base, name + " · " + method_label)
        m.param(SHEET, base + 1, name + " 내용연수 (개월)", item["life_months"], SHEET,
                "취득 월별 상각 기간. 마지막 상각 월에 잔존 순액 전부를 반영합니다.")
        if method == "declining":
            m.param(SHEET, base + 2, name + " 연 상각률", item["annual_rate"], SHEET,
                    "월 정률은 연 상각률의 1/12입니다.")
        source = {p["month"]: p for p in item["purchases"]}
        for offset, field, caption, annual in ((4, "quantity", "취득 수량", "sum"),
                                               (5, "unit_price", "취득 단가 (원)", "none")):
            values = [source.get(month, {}).get(field, 0) if i > m.last else None
                      for i, month in enumerate(m.months)]
            m.input(SHEET, base + offset, name + " " + caption, values, SHEET,
                    "추정 월 취득 수량과 한 단위당 원 단가", annual=annual)
        forecast(SHEET, base + 6, name + " 취득액", lambda i,t,c,p: f"{c}{base+4}*{c}{base+5}/Unit")
        forecast(SHEET, base + 7, name + " 신규 상각", lambda i,t,c,p: _new_dep_formula(c,p,first,base,method))
        forecast(SHEET, base + 8, name + " 신규 기말 순액",
                 lambda i,t,c,p: f"MAX(0,{p}{base+8}+{c}{base+6}-{c}{base+7})", "last")
        acquisition_rows.append(base + 6)
        new_dep_rows.append(base + 7)
        opening = item.get("opening")
        if not opening:
            continue
        for offset, field, caption in ((9, "cost", "기존 취득원가 (원)"),
                                       (10, "accumulated_dep", "기존 상각누계액 (원)"),
                                       (11, "remaining_life_months", "기존 잔여기간 (개월)")):
            m.param(SHEET, base + offset, name + " " + caption, opening[field], SHEET,
                    "실적 마감 월말 자산대장 기준")
        end = base + 13
        def opening_charge(i,t,c,p):
            elapsed = f"COLUMNS(${first}${end}:{c}{end})"
            charge = (f"$G${base+9}/Unit/$G${base+1}" if method == "straight_line"
                      else f"{p}{end}*$G${base+2}/12")
            return f"MIN(MAX(0,{p}{end}),IF({elapsed}>=$G${base+11},{p}{end},{charge}))"
        forecast(SHEET, base + 12, name + " 기존 상각", opening_charge)
        row(SHEET, end, name + " 기존 기말 순액",
            lambda i,t,c,p: (f"($G${base+9}-$G${base+10})/Unit" if i == m.last else
                             f"MAX(0,{p}{end}-{c}{base+12})" if i > m.last else None), "last")
        existing_dep_rows.append(base + 12)
        opening_rows.append(end)

    # Keep the existing aggregate inputs available for distinct acquisitions
    # and for the part of the cutoff PPE balance without a class-level schedule.
    forecast(SHEET, 27, "금액 직접 입력분 신규 상각",
             lambda i,t,c,p: legacy_depreciation(i,t,c,p,25,28))
    forecast(SHEET, 28, "금액 직접 입력분 신규 기말 순액",
             lambda i,t,c,p: f"{p}28+{c}25-{c}27", "last")
    row(SHEET, 29, "자산군 미배정 기존 기말 순액",
        lambda i,t,c,p: (f"{c}10-({_sum_rows(c,opening_rows)})" if i == m.last else
                         f"{p}29-{c}30" if i > m.last else None), "last")
    forecast(SHEET, 30, "자산군 미배정 기존 상각", lambda i,t,c,p: f"MIN(MAX(0,{p}29),$G$23/Unit)")
    forecast(SHEET, 8, "투자 지출", lambda i,t,c,p: _sum_rows(c,[25,*acquisition_rows]))
    forecast(SHEET, 11, "기존 자산 상각", lambda i,t,c,p: _sum_rows(c,[30,*existing_dep_rows]))
    forecast(SHEET, 13, "신규 자산 상각", lambda i,t,c,p: _sum_rows(c,[27,*new_dep_rows]))
    cutoff = m.col(m.last)
    m.add_check("자산군 기초 순액 배정",
                f'=IF(\'{SHEET}\'!{cutoff}29>=-0.5/Unit,"OK","NO")')
