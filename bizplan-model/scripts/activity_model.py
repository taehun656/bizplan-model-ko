"""Activity-driven revenue, metric-linked expenses and ordered initiative attribution."""
from collections import defaultdict
from datetime import date

from openpyxl.utils import get_column_letter as letter
from openpyxl.utils.datetime import to_excel

from activity_inputs import sample_months
from build_model import ref
from dashboard import UNIT_LABEL
from design_palette import SILVER
from driver_blocks import _annual_factor
from workbook_style import _edge, set_number_format, set_numeric_list_format

PLAN = "영업계획"
METRIC = "활동지표"
RESULT = "민감도및결과"
SENS = "Sensitivity"
START = 7
ACT, YIELD, QTY, PRICE, REVENUE = range(14, 19)
BASE_ACT, BASE_YIELD, BASE_QTY, BASE_PRICE, BASE_REV = range(22, 27)
NET_ACTIVITY = 20
PLANNED_ACTIVITY, INCLUDED_ACTIVITY, REFERENCE_ACTIVITY = 27, 28, 29
INCREMENT_YIELD = 39


def serial(month):
    return to_excel(date.fromisoformat(month + "-01")) if month else None


def lit(text):
    return '"' + text.replace('"', '""') + '"'


def series(m, sheet, row, label, expression, annual="sum"):
    m.row(sheet, row, label, ["="+expression(i,m.col(i)) for i in range(len(m.months))], annual=annual)


def input_series(m, sheet, row, label, values, group):
    m.row(sheet,row,label,values,kind="input",annual="none")
    m.register(sheet,f"J{row}:{m.col(len(values)-1)}{row}",label,group,"월별 관측값 또는 계획 입력")


def scalar(m,sheet,row,label,value,group):
    m.cell(sheet,row,2,label,"text")
    m.cell(sheet,row,7,value,"input")
    m.register(sheet,f"G{row}",label,group,"모델 계산에 적용하는 입력값")
    return f"$G${row}"


def lookup(sheet,col,row):
    return f"{ref(sheet)}!{col}{row}"


def total(terms):
    terms=list(terms)
    return "("+ "+".join(terms)+")" if terms else "0"


def setup_plans(m):
    items=m.spec["initiatives"]
    m.plan_rows=list(range(START,START+len(items)+m.spec["initiative_blank_rows"]))
    m.metric_cols={metric["id"]:letter(i) for i,metric in enumerate(m.spec["metrics"],14)}
    headers=["계획 이름","매출 ID","적용 0/1","시작 월","종료 월","활동량 증감","해당 활동당 수량",
             "해당 단가 (원)","전체 단가 변화율","전체 활동당 수량 변화율","비용 ID","월 고정비 증감 (원)"]
    headers.extend(v["label"]+" / 추가 활동당" for v in m.spec["metrics"])
    m.cell(PLAN,4,2,"빈 줄에 계획 이름·매출 ID·적용 1·시작 월을 입력하세요. 종료 월을 비우면 계획 끝까지 적용합니다.","text")
    m.table(PLAN,START-1,2,[headers]);m.header(PLAN,START-1)
    end=letter(1+len(headers))
    if m.plan_rows:
        last=m.plan_rows[-1]
        m.register(PLAN,f"B{START}:{end}{last}","사업부 계획 입력표","사업부 계획","금액은 원, 변화율은 소수")
        m.named("ACT_Plans",PLAN,f"$B${START}:${end}${last}")
    # Append to the pre-existing switch layout without shifting its macro ranges.
    top=max(r for r,c in m.grid[SENS])+3
    m.plan_switches={}
    m.section(SENS,top,"사업부 계획 적용")
    m.cell(SENS,top+1,2,"계획","text");m.cell(SENS,top+1,3,"현재 적용","text")
    for j,case in enumerate(m.cases,5):
        m.cell(SENS,top+1,j,case["label"],"text")
    m.header(SENS,top+1)
    for j,r in enumerate(m.plan_rows):
        item=items[j] if j<len(items) else {}
        if item:
            values=[item["name"],item["line"],int(item["enabled"]),serial(item["start"]),serial(item.get("end")),
                    item["activity_delta"],item.get("yield"),item.get("price"),item["price_pct"],item["yield_pct"],
                    item.get("cost_line"),item["fixed_cost"]]
            values.extend(item.get("metric_factors",{}).get(k) for k in m.metric_cols)
            m.table(PLAN,r,2,[values])
        switch=top+2+j
        m.plan_switches[r]=f"'Sensitivity'!$C${switch}"
        m.cell(SENS,switch,2,f'=IF({ref(PLAN)}!B{r}="","",{ref(PLAN)}!B{r})')
        m.cell(SENS,switch,3,f"=INDEX(E{switch}:{letter(4+len(m.cases))}{switch},1,SN_Case)")
        for c,case in enumerate(m.cases,5):
            selected=case.get("initiatives",m.spec["scenarios"][case["scenario"]].get("initiatives"))
            enabled=not item or selected is None or item["id"] in selected
            m.cell(SENS,switch,c,int(enabled),"input")
        m.register(SENS,f"E{switch}:{letter(4+len(m.cases))}{switch}",f"계획 {j+1} 케이스 적용","사업부 계획","1 적용, 0 제외")
    from dashboard import add_consumption_price_inputs
    add_consumption_price_inputs(m)
    m.stages={}
    m.cost_rates={}
    m.cost_keyin_ranges = []
    m.metric_position=8


def allocate(m,size):
    row=m.metric_position
    m.metric_position+=size
    return row


def weighted(m,sheet,row,numerator,denominator,scale=1):
    for c in [8,*range(10+len(m.months),10+len(m.months)+len(m.years))]:
        col=letter(c)
        m.cell(sheet,row,c,f"=IFERROR({col}{numerator}*{scale}/{col}{denominator},0)")


def folded_steps(col,steps,test):
    factors=[f"IF(AND({col}$5>=$G${r},{test(r)}),1+$H${r},1)" for r in steps]
    return "*".join(factors) or "1"


def blended_coefficient(base_activity, increment, base_rate, increment_rate):
    volume = f"({base_activity}+{increment})"
    return (f"IF({increment}=0,{base_rate},IF({volume}=0,0,"
            f"({base_activity}*({base_rate})+{increment}*{increment_rate})/{volume}))")


def increment_schedule(m, sheet, increment):
    label = increment["label"]
    included = [increment["included_monthly"].get(t, 0) for t in m.months]
    planned = [None if i <= m.last else increment["monthly"].get(t, included[i])
               for i,t in enumerate(m.months)]
    m.row(sheet, PLANNED_ACTIVITY, label, planned, kind="input")
    address = f"{m.col(m.last+1)}{PLANNED_ACTIVITY}:{m.col(len(m.months)-1)}{PLANNED_ACTIVITY}"
    m.register(sheet, address, label, "사업부 계획", "월 계획 횟수를 입력하면 기준에 포함된 횟수를 뺀 차이가 수량·지표·연동 비용에 반영됩니다.")
    m.row(sheet, INCLUDED_ACTIVITY, "기준에 이미 포함된 활동량",
          [None if i<=m.last else v for i,v in enumerate(included)],kind="actual")
    m.cell(sheet, INCREMENT_YIELD, 2, "추가 활동당 수량", "text")
    m.cell(sheet, INCREMENT_YIELD, 7, increment["yield_per_activity"], "actual")
    def net(i,c):
        if i <= m.last:
            return "0"
        cell = f"{c}{PLANNED_ACTIVITY}"
        whole = f",{cell}=INT({cell})" if increment["whole"] else ""
        return f"IF(AND(ISNUMBER({cell}),{cell}>=0{whole}),{cell}-{c}{INCLUDED_ACTIVITY},NA())"
    series(m, sheet, NET_ACTIVITY, "기준 대비 활동량 증감", net)
    start, end = m.col(m.last+1), m.col(len(m.months)-1)
    span = lambda row: f"{ref(sheet)}!{start}{row}:{end}{row}"
    m.add_check(label + " 및 기준 활동량",
                f'=IFERROR(IF(AND(COUNT({span(NET_ACTIVITY)})={len(m.months)-m.last-1},'
                f'SUMPRODUCT(--(({span(REFERENCE_ACTIVITY)}+{span(NET_ACTIVITY)})<0))=0),"OK","NO"),"NO")')


def revenue_block(m,line,sheet,volume):
    ident=line["id"];p=line["params"]
    increment = p.get("increment")
    if increment:
        increment_schedule(m, sheet, increment)
    history=m.months[:m.last+1]
    for row,field,label in ((30,"activity","실적 활동량"),(31,"quantity","실적 수량"),(32,"price","실적 단가 (원)")):
        m.row(sheet,row,label,[volume[t][field] if i<=m.last else None for i,t in enumerate(m.months)],kind="actual",annual="none")
    source_id = m.activity_lines.index(line) + 1
    m.named(f"ACT_SalesSource_{source_id}", sheet, f"$J$30:${m.col(m.last)}$32")
    series(m,sheet,33,"실적 활동당 수량",lambda i,c:f"IFERROR({c}31/{c}30,0)","none")
    fields=[("activity",30,35),("yield",33,36),("price",32,37)]
    for field,_,row in fields:
        overrides=dict(p[field].get("monthly",{}))
        for t,v in volume.items():
            if t>history[-1]:
                value=v["quantity"]/v["activity"] if field=="yield" and v["activity"] else (0 if field=="yield" else v[field])
                if t in overrides and overrides[t]!=value:
                    raise ValueError(f"{ident}.{field} {t}: 설계서와 CSV의 월 계획 불일치")
                overrides[t]=value
        label = {"activity":"활동량 계획","yield":"활동당 수량 계획","price":"단가 계획 (원)"}[field]
        if increment and field in ("activity", "yield"):
            label = "기준 " + label + " (증감 전)"
        input_series(m,sheet,row,label,
                     [overrides.get(t) for t in m.months],"사업부 계획")
    growth=scalar(m,sheet,38,"활동당 수량 연 증감률",p["yield_growth"],"사업부 계획")
    steps=list(range(41,41+len(p["price_steps"])))
    if steps:
        m.table(sheet,40,7,[["시행 월","변화율","인상 전 기준 월만 (1/0)","근거"]]);m.header(sheet,40)
    for row,step in zip(steps,p["price_steps"]):
        m.table(sheet,row,7,[[serial(step["month"]),step["pct"],int(step["apply_to"]=="pre_step_months"),step["basis"]]])
        m.register(sheet,f"G{row}:I{row}","단가 조정","사업부 계획",step["basis"])
    # Historical sample prices are restated to the cutoff price regime only
    # for pre-step observations. Reported historical prices stay untouched.
    def comparable(i,c):
        fs=[f"IF(AND($I${r}=1,$G${r}<=LastActual,{c}$5<$G${r}),1+$H${r},1)" for r in steps]
        return f"{c}32"+ "".join("*"+v for v in fs) if i<=m.last else "0"
    series(m,sheet,34,"동일 가격 기준으로 환산한 실적 단가",comparable,"none")
    for field,raw,over in fields:
        rule=p[field]
        r={"activity":BASE_ACT,"yield":BASE_YIELD,"price":BASE_PRICE}[field]
        def reference(i,c):
            if i<=m.last:return f"{c}{raw}"
            target=m.months[i]
            samples=sample_months(rule,history,history[-1],target)
            source_row=34 if field=="price" else raw
            def avg(dates):
                cells=[f"{m.col(history.index(t))}{source_row}" for t in dates]
                return "AVERAGE("+",".join(cells)+")" if cells else "0"
            default=avg(samples)
            if rule["base"]=="base_month" and rule.get("seasonal"):
                base=rule["month"]
                peers=sample_months(rule,history,history[-1],base)
                default=f"{m.col(history.index(base))}{source_row}*IFERROR({default}/{avg(peers)},1)"
            expr=f"IF(ISNUMBER({c}{over}),{c}{over},{default})"
            return expr
        if increment and field == "activity":
            series(m,sheet,REFERENCE_ACTIVITY,"기준 활동량 (증감 전)",reference)
        def baseline(i,c):
            expr = reference(i,c)
            if i <= m.last:
                return expr
            if increment:
                if field == "activity":
                    expr = f"{c}{REFERENCE_ACTIVITY}+{c}{NET_ACTIVITY}"
                elif field == "yield":
                    expr = blended_coefficient(f"{c}{REFERENCE_ACTIVITY}", f"{c}{NET_ACTIVITY}", expr,
                                               f"$G${INCREMENT_YIELD}")
            if field=="activity":return f"({expr})*{m.factor(ident)}*{m.driver_switch(ident,'activity_multiplier')}"
            if field=="yield":return f"({expr})*{_annual_factor(c,growth)}*{m.driver_switch(ident,'yield_multiplier')}"
            factor=folded_steps(c,steps,lambda step:
                f"OR($I${step}=0,AND(NOT(ISNUMBER({c}{over})),$G${step}>LastActual))")
            return f"({expr})*{factor}*{m.driver_switch(ident,'price_multiplier')}"
        series(m,sheet,r,{"activity":"계획 조정 전 활동량","yield":"계획 조정 전 활동당 수량","price":"계획 조정 전 단가 (원)"}[field],baseline,"sum" if field=="activity" else "none")
    series(m,sheet,BASE_QTY,"계획 조정 전 수량",lambda i,c:f"{c}{BASE_ACT}*{c}{BASE_YIELD}")
    series(m,sheet,BASE_REV,"계획 조정 전 매출",lambda i,c:f"{c}{BASE_QTY}*{c}{BASE_PRICE}/Unit")
    prior=(BASE_ACT,BASE_QTY,BASE_REV)
    first=43+len(steps)
    for j,pr in enumerate(m.plan_rows):
        r=first+7*j
        m.stages[ident,pr]=r
        a,q,v=prior
        pl=lambda col: f"{ref(PLAN)}!${col}${pr}"
        def applies(i,c):
            return (f'--AND({c}$5>LastActual,{pl("B")}<>"",EXACT({pl("C")},{lit(ident)}),'
                    f'{pl("D")}=1,{m.plan_switches[pr]}=1,ISNUMBER({pl("E")}),{c}$5>={pl("E")},'
                    f'OR(ISBLANK({pl("F")}),{c}$5<={pl("F")}))')
        series(m,sheet,r,f"계획 {j+1} 적용",applies,"none")
        series(m,sheet,r+1,"활동량 증감",lambda i,c:f"{c}{r}*{pl('G')}")
        series(m,sheet,r+2,"적용 후 활동량",lambda i,c:f"{c}{a}+{c}{r+1}")
        def delta_qty(c):
            unit_yield=f"IF(ISNUMBER({pl('H')}),{pl('H')},IFERROR({c}{q}/{c}{a},{c}{BASE_YIELD}))"
            return f"{c}{r+1}*{unit_yield}"
        series(m,sheet,r+3,"적용 후 수량",lambda i,c:f"({c}{q}+{delta_qty(c)})*(1+{c}{r}*{pl('K')})")
        def after_revenue(i,c):
            price=f"IF(ISNUMBER({pl('I')}),{pl('I')},IFERROR({c}{v}*Unit/{c}{q},{c}{BASE_PRICE}))"
            return f"({c}{v}+{delta_qty(c)}*{price}/Unit)*(1+{c}{r}*{pl('K')})*(1+{c}{r}*{pl('J')})"
        series(m,sheet,r+4,"적용 후 매출",after_revenue)
        series(m,sheet,r+5,"이 계획의 매출 증감",lambda i,c:f"{c}{r+4}-{c}{v}")
        prior=(r+2,r+3,r+4)
    a,q,v=prior
    for row,label,r in ((ACT,"활동량",a),(QTY,"판매 수량",q),(REVENUE,"활동 기준 매출",v)):
        series(m,sheet,row,label,lambda i,c:f"{c}{r}")
    series(m,sheet,YIELD,"활동당 수량",lambda i,c:f"IFERROR({c}{QTY}/{c}{ACT},{c}{BASE_YIELD})","none")
    series(m,sheet,PRICE,"단가 (원)",lambda i,c:f"{c}32" if i<=m.last else f"IFERROR({c}{REVENUE}*Unit/{c}{QTY},{c}{BASE_PRICE})","none")
    weighted(m,sheet,YIELD,QTY,ACT)
    weighted(m,sheet,PRICE,REVENUE,QTY,"Unit")
    weighted(m,sheet,BASE_YIELD,BASE_QTY,BASE_ACT)
    weighted(m,sheet,BASE_PRICE,BASE_REV,BASE_QTY,"Unit")
    series(m,sheet,10,"활동량 × 활동당 수량 × 단가",lambda i,c:f"{c}{REVENUE}")
    series(m,sheet,19,"실적 매출 대사 차이 (원)",lambda i,c:f"IF({c}$5<=LastActual,ABS({c}{REVENUE}-{c}9)*Unit,0)","max")
    end=m.col(len(m.months)-1)
    m.add_check(line["label"]+" 활동 매출 실적 대사",f'=IF(MAX({ref(sheet)}!J19:{end}19)<=0.5,"OK","NO")')
    for row in (YIELD,PRICE,BASE_PRICE,BASE_YIELD,33,34,35,36,37):
        m.format_row(sheet,row,"#,##0.00##")
    m.format_row(sheet,38,"0.0%")
    m.section(sheet,13,"활동량과 판매 실적");m.section(sheet,21,"계획 조정 전 기준")
    m.cell(sheet,REVENUE,2,'="활동 기준 매출 ("&'+UNIT_LABEL[1:]+'&")"')


def build_metrics(m):
    m.metric_totals={}
    m.metric_deltas={}
    m.metric_observations=[]
    raw={(v["metric"],v["line"],v["month"]):v["value"] for v in m.data["metrics_monthly"]}
    history=m.months[:m.last+1]
    for metric in m.spec["metrics"]:
        key=metric["id"]
        head=allocate(m,2)
        m.section(METRIC,head,metric["label"]+" · 활동 지표")
        per_line=[]
        increments=defaultdict(list)
        for ident,coeff in metric["lines"].items():
            sh=m.details[ident]
            increment = next(line for line in m.activity_lines if line["id"] == ident)["params"].get("increment", {})
            r=allocate(m,6+len(m.plan_rows))
            per_line.append(r+4)
            observed=[raw.get((key,ident,t)) for t in m.months]
            m.row(METRIC,r,ident+" · 지표 원천",observed,kind="actual",annual="none")
            m.metric_observations.append((key,ident,r,r+4))
            source_id = len(m.metric_observations)
            m.named(f"ACT_MetricSource_{source_id}", METRIC, f"$J${r}:${m.col(m.last)}${r}")
            if isinstance(coeff,dict):
                input_series(m,METRIC,r+1,ident+" · 월별 계수", [coeff.get("monthly",{}).get(t) for t in m.months],"활동 지표")
                def coefficient(i,c):
                    if i<=m.last: return f"IFERROR({c}{r}/{lookup(sh,c,ACT)},0)"
                    dates=sample_months({"base":"last_12m"},history,history[-1],m.months[i])
                    base=f"{m.col(history.index(dates[-1]))}{r+2}" if dates else "0"
                    return f"IF(ISNUMBER({c}{r+1}),{c}{r+1},{base})"
            else:
                address=scalar(m,METRIC,r+1,ident+" · 활동당 계수",coeff,"활동 지표")
                def coefficient(i,c): return address
            increment_rate = increment.get("metric_factors", {}).get(key)
            applied = coefficient
            if increment_rate is not None:
                m.cell(METRIC,r+2,7,increment_rate,"actual")
                def applied(i,c):
                    base_rate = coefficient(i,c)
                    if i <= m.last:
                        return base_rate
                    return blended_coefficient(lookup(sh,c,REFERENCE_ACTIVITY), lookup(sh,c,NET_ACTIVITY),
                                               base_rate, f"$G${r+2}")
            label = ident+" · 적용 계수" + (" (G열은 추가 활동당 계수)" if increment_rate is not None else "")
            series(m,METRIC,r+2,label,applied,"none")
            series(m,METRIC,r+3,ident+" · 기준 지표",lambda i,c:f"{lookup(sh,c,BASE_ACT)}*{c}{r+2}")
            delta_rows=[]
            for j,pr in enumerate(m.plan_rows):
                dr=r+5+j
                delta_rows.append(dr);increments[pr].append(dr)
                override=f"{ref(PLAN)}!${m.metric_cols[key]}${pr}"
                series(m,METRIC,dr,f"{ident} · 계획 {j+1} 지표 증감",lambda i,c:
                       f"{lookup(sh,c,m.stages[ident,pr]+1)}*IF(ISNUMBER({override}),{override},{c}{r+2})")
            series(m,METRIC,r+4,ident+" · 적용 지표",lambda i,c:
                   f"IF(AND({c}$5<=LastActual,ISNUMBER({c}{r})),{c}{r},{c}{r+3}+{total(c+str(d) for d in delta_rows)})")
        m.metric_totals[key]=head+1
        series(m,METRIC,head+1,metric["label"]+" 합계",lambda i,c:total(c+str(r) for r in per_line))
        for pr in m.plan_rows:
            r=allocate(m,1);m.metric_deltas[key,pr]=r
            series(m,METRIC,r,"계획별 "+metric["label"]+" 증감",lambda i,c:total(c+str(v) for v in increments[pr]))
    m.plan_active={}
    for pr in m.plan_rows:
        r=allocate(m,1);m.plan_active[pr]=r
        series(m,METRIC,r,"계획 적용 여부",lambda i,c:
               total(lookup(m.details[line["id"]],c,m.stages[line["id"],pr]) for line in m.activity_lines),"max")
    m.fixed_rows = {}
    m.fixed_effects = {}
    for line in m.spec["cost_lines"]:
        ident = line["id"]
        top = allocate(m, 1 + len(m.plan_rows))
        m.fixed_rows[ident] = top
        for offset, plan_row in enumerate(m.plan_rows, 1):
            detail = top + offset
            m.fixed_effects[ident, plan_row] = detail
            target = f"{ref(PLAN)}!$L${plan_row}"
            amount = f"{ref(PLAN)}!$M${plan_row}"
            series(m, METRIC, detail, f"{line['label']} · 계획 {offset} 고정비",
                   lambda i, c: f"IF(EXACT({target},{lit(ident)}),{c}{m.plan_active[plan_row]}*{amount}/Unit,0)")
        series(m, METRIC, top, line["label"]+" · 고정비 증감",
               lambda i, c: f"SUM({c}{top+1}:{c}{top+len(m.plan_rows)})" if m.plan_rows else "0")


def forecast_cost_input(m, row, label, value, positive=False):
    """Populate each editable month from a master rate or a dated rate change."""
    start = m.last + 1
    address = f"{m.col(start)}{row}:{m.col(len(m.months) - 1)}{row}"
    values = [None] * start
    master = f"$G${row}"
    if isinstance(value, dict):
        rate = value[m.months[start]]
        for month in m.months[start:]:
            rate = value.get(month, rate)
            values.append(rate)
    else:
        if value != "actual":
            scalar(m, METRIC, row, label, value, "활동 지표")
        values.extend(["=" + master] * (len(m.months) - start))
    m.row(METRIC, row, label, values, annual="none")
    m.register(METRIC, address, label, "활동 지표", "월별 기준값 · 노란 칸에 직접 입력")
    m.format_row(METRIC, row, "#,##0.00##")
    m.cost_keyin_ranges.append((address, positive))


def checked_cost_input(col, row, *, positive=False):
    # Excel paste and Clear Contents can bypass validation on a yellow cell.
    cell = f"{col}{row}"
    operator = ">" if positive else ">="
    return f"IF(AND(ISNUMBER({cell}),{cell}{operator}0),{cell},NA())"


def subsidy_reference(m, ident, key):
    return f"CS_Link_{m.cost_detail_rows[ident]}_{key}"


def _consumption_cost(m, line, expense, row):
    p = line["params"]
    ident = line["id"]
    for offset, key, label in ((2, "efficiency", "기준 효율 (물량/사용량)"),
                               (3, "price", "구매 단가 (원/단위)"),
                               (4, "subsidy", "단위 보조 (원/단위)")):
        forecast_cost_input(m, row + offset, label, p[key], positive=key == "efficiency")

    observed = p.get("actual_usage", {})
    if observed:
        m.row(METRIC, row + 9, "관측 사용량", [observed.get(t) for t in m.months], annual="none")
        m.register(METRIC, f"J{row+9}:{m.col(m.last)}{row+9}", "관측 사용량", "활동 지표",
                   "해당 비용 라인에 배정한 실적 사용량. 미관측 월은 빈칸.")
        for i in range(m.last + 1):
            c = m.col(i)
            actual_cost = lookup("운영비용", c, expense + 1)
            m.cell(METRIC, row + 3, 10+i, f'=IF({c}{row+9}>0,{actual_cost}*Unit/{c}{row+9},"")')
    # Historical price cells show observed expense per unit; the historical model
    # rows retain their former scalar assumptions and never use a forecast case.
    scalar_history = all(not isinstance(p[k], dict) for k in ("efficiency", "price", "subsidy"))

    link = p.get("subsidy_link")
    if link:
        threshold, cap, share, baseline = [subsidy_reference(m, ident, key)
                                           for key in ("threshold", "cap", "eligible_share", "baseline_amount")]
        values = [None] * (m.last + 1)
        for i, month in enumerate(m.months[m.last + 1:], m.last + 1):
            c = m.col(i)
            year = month[:4]
            key = f"rate_{year}" if year in link.get("annual_rates", {}) else "rate"
            rate = subsidy_reference(m, ident, key)
            price = f"{checked_cost_input(c, row+3)}*{m.driver_switch(ident, 'price_multiplier')}"
            base = checked_cost_input(c, row+4)
            linked_amount = f"{share}*MIN(MAX({price}-{threshold},0)*{rate},{cap})"
            values.append(f"=({base}+{linked_amount}-{baseline})*{m.driver_switch(ident, 'subsidy_multiplier')}")
        m.row(METRIC, row+10, "적용 단위 보조 (원/단위)", values, annual="none")
        m.format_row(METRIC, row+10, "#,##0.00##")

    def usage(i, c):
        if i <= m.last:
            if isinstance(p["efficiency"], dict):
                return f'IF(ISNUMBER({c}{row+9}),{c}{row+9},"")'
            return f"{c}{row+1}/$G${row+2}"
        return f"{c}{row+1}/({checked_cost_input(c, row+2, positive=True)}*{m.driver_switch(ident, 'efficiency_multiplier')})"

    def net_price(i, c):
        if i <= m.last:
            return f"($G${row+3}-$G${row+4})"
        if link:
            return f"({checked_cost_input(c, row+3)}*{m.driver_switch(ident, 'price_multiplier')}-{c}{row+10})"
        return (f"({checked_cost_input(c, row+3)}*{m.driver_switch(ident, 'price_multiplier')}"
                f"-{checked_cost_input(c, row+4)}*{m.driver_switch(ident, 'subsidy_multiplier')})")

    def unit_rate(i, c):
        if i <= m.last:
            if scalar_history:
                return f"{net_price(i,c)}/$G${row+2}"
            return f"IFERROR({lookup('운영비용',c,expense+1)}*Unit/{c}{row+1},0)"
        return f"{net_price(i,c)}/({checked_cost_input(c, row+2, positive=True)}*{m.driver_switch(ident, 'efficiency_multiplier')})"

    def cost(i, c):
        if i <= m.last and not scalar_history:
            return lookup("운영비용", c, expense + 1)
        return f"({c}{row+5}*{net_price(i,c)}+{c}{row+7})/Unit"

    series(m, METRIC, row+5, "모형 사용량", usage)
    series(m, METRIC, row+6, "물량당 순원가 (원)", unit_rate, "none")
    input_series(m, METRIC, row+7, "월별 비용 보정 (원)",
                 [p.get("adjustments", {}).get(t) for t in m.months], "활동 지표")
    series(m, METRIC, row+8, "동인 기준 비용", cost)



def metric_cost(m, line, expense):
    p = line["params"]
    ident, method = line["id"], line["method"]
    r = allocate(m, 11)
    m.section(METRIC, r, line["label"] + " · 원가 동인")
    if method == "per_activity":
        def measure(i, c):
            return total(lookup(m.details[k], c, ACT) for k in p["base_lines"])
    else:
        def measure(i, c):
            return c + str(m.metric_totals[p["metric"]])
    series(m, METRIC, r+1, "적용 물량", measure)
    if method == "consumption":
        _consumption_cost(m, line, expense, r)
    else:
        value = p["unit_cost"]
        if value == "actual":
            first, last = m.col(max(0,m.last-11)), m.col(m.last)
            actual = f"SUM('운영비용'!{first}{expense+1}:{last}{expense+1})*Unit"
            denom = f"SUM({first}{r+1}:{last}{r+1})"
            m.cell(METRIC, r+2, 2, "실적 물량당 원가 (원)", "text")
            m.cell(METRIC, r+2, 7, f"=IF({denom}>0,{actual}/{denom},NA())")
        if method == "per_metric":
            forecast_cost_input(m, r+2, "물량당 기준 원가 (원)", value)
        elif value != "actual":
            scalar(m, METRIC, r+2, "물량당 원가 (원)", value, "활동 지표")
        growth = scalar(m, METRIC, r+3, "물량당 원가 연 증감률", p["annual_growth"], "활동 지표")
        m.format_row(METRIC, r+3, "0.0%")

        def applied_rate(i, c):
            if method == "per_metric" and i > m.last:
                rate = checked_cost_input(c, r+2)
            elif isinstance(value, dict):
                rate = f"IFERROR({lookup('운영비용',c,expense+1)}*Unit/{c}{r+1},0)"
            else:
                rate = f"$G${r+2}"
            return f"{rate}*{_annual_factor(c,growth)}"

        series(m, METRIC, r+6, "적용 물량당 원가 (원)", applied_rate, "none")
        series(m, METRIC, r+8, "동인 기준 비용", lambda i,c: f"{c}{r+1}*{c}{r+6}/Unit")
    m.cost_rates[ident] = r+6
    return lambda c: lookup(METRIC,c,r+8)



def finish_activity(m):
    lines=m.spec["revenue_lines"]+m.spec["cost_lines"]
    m.effect_rows=[]
    for pr in m.plan_rows:
        first=allocate(m,len(lines)+3)
        delta={v["id"]:first+j for j,v in enumerate(lines)}
        for item in lines:
            ident=item["id"];p=item["params"]
            kind=item.get("driver",item.get("method"))
            is_cost="method" in item
            def difference(i,c):
                if kind=="activity":
                    return lookup(m.details[ident],c,m.stages[ident,pr]+5)
                # Designated depreciation is driven by the investment module.
                if is_cost and set(item["accounts"]).intersection(m.spec["investment"]["depreciation_accounts"]):
                    return "0"
                variable="0"
                if kind in ("consumption","per_metric","per_activity"):
                    if kind=="per_activity":
                        qty=total(lookup(m.details[k],c,m.stages[k,pr]+1) for k in p["base_lines"])
                    else:
                        qty=c+str(m.metric_deltas[p["metric"],pr])
                    variable=f"{qty}*{c}{m.cost_rates[ident]}/Unit"
                elif kind in ("pct_of_revenue","mixed","pct_of_line"):
                    targets=p["line_ids"] if kind=="pct_of_line" else p["revenue_ids"]
                    if is_cost:
                        row=m.cost_detail_rows[ident]+(6 if kind=="mixed" else 4)
                        rate=lookup("운영비용","$G$",row)
                    else:
                        rate=lookup(m.details[ident],"$G$",14)
                    variable=total(c+str(delta[k]) for k in targets)+"*"+rate
                if p.get("months"):
                    variable+=f"*{lookup('운영비용',c,m.cost_detail_rows[ident]+7)}"
                result=f"({variable})*{m.factor(ident)}"
                if is_cost:
                    result += f"+{c}{m.fixed_effects[ident,pr]}"
                return result
            series(m,METRIC,delta[ident],f"계획 {pr-START+1} · {item['label']} 증감",difference)
        r=first+len(lines)
        for offset,group,label in ((0,m.spec["revenue_lines"],"매출 증감"),(1,m.spec["cost_lines"],"연동 원가 증감")):
            series(m,METRIC,r+offset,label,lambda i,c:total(c+str(delta[line["id"]]) for line in group))
        series(m,METRIC,r+2,"공헌이익 증감",lambda i,c:f"{c}{r}-{c}{r+1}")
        m.effect_rows.append(r)
    # Explicit addresses let verification use source IDs even with duplicate captions.
    def metadata(name,header,rows):
        r=allocate(m,len(rows)+2)
        m.table(METRIC,r,2,[header,*rows])
        m.header(METRIC,r)
        m.named(name,METRIC,f"$B${r}:${letter(1+len(header))}${r+len(rows)}")
    metadata("ACT_Lines",["라인 ID","시트"],[(v["id"],m.details[v["id"]]) for v in m.activity_lines])
    metadata("ACT_Metrics",["지표 ID","라인 ID","원천 행","적용 행"],m.metric_observations)
    metadata("ACT_Effects",["계획 행","효과 행"],list(zip(m.plan_rows,m.effect_rows)))
    targets=allocate(m,max(len(m.activity_lines),len(m.spec["cost_lines"]))+2)
    m.table(METRIC,targets,2,[["활동 매출 ID","고정비 대상 ID"]])
    for j,item in enumerate(m.activity_lines,1):
        m.cell(METRIC,targets+j,2,item["id"],"text")
    m.named("ACT_Targets",METRIC,f"$B${targets+1}:$B${targets+len(m.activity_lines)}")
    eligible=[v["id"] for v in m.spec["cost_lines"] if not set(v["accounts"]).intersection(m.spec["investment"]["depreciation_accounts"])]
    for j,key in enumerate(eligible,1):m.cell(METRIC,targets+j,3,key,"text")
    if eligible:m.named("ACT_CostTargets",METRIC,f"$C${targets+1}:$C${targets+len(eligible)}")
    last=m.col(len(m.months)-1)
    for line in m.activity_lines:
        sh=m.details[line["id"]]
        m.add_check(line["label"]+" 활동량·수량 음수 없음",
                    f'=IF(MIN({ref(sh)}!J14:{last}16)>=0,"OK","NO")')


def results_table(m,top):
    m.section(RESULT,top)
    m.cell(RESULT,top,2,f'="{m.months[-1][:4]}년 계획 조정 효과 · 단위 "&{UNIT_LABEL[1:]}')
    m.table(RESULT,top+1,2,[["계획","매출 증감","연동 원가 증감","공헌이익 증감"]])
    m.header(RESULT,top+1)
    year=m.col(len(m.months)+len(m.years)-1)
    for i,(pr,effect) in enumerate(zip(m.plan_rows,m.effect_rows),top+2):
        name=f"{ref(PLAN)}!B{pr}"
        m.cell(RESULT,i,2,f'=IF({name}="","",{name})')
        for k in range(3):
            m.cell(RESULT,i,3+k,f'=IF(B{i}="","",{lookup(METRIC,year,effect+k)})')
    bottom=top+2+len(m.effect_rows)
    m.cell(RESULT,bottom,2,"계획 효과 합계","text")
    for c in range(3,6):
        col=letter(c)
        m.cell(RESULT,bottom,c,f"=SUM({col}{top+2}:{col}{bottom-1})" if m.effect_rows else 0)
    m.total(RESULT,bottom)
    m.named("ACT_Result",RESULT,f"$B${top}:$E${bottom}")
    return bottom+2


def style_activity(ws,m):
    if ws.Name == METRIC:
        for address, positive in m.cost_keyin_ranges:
            validation = ws.Range(address).Validation
            validation.Delete()
            validation.Add(2, 1, 5 if positive else 7, 0)  # decimal, stop, > or >= zero
            validation.IgnoreBlank = False
            validation.ShowError = True
    if ws.Name==PLAN:
        for c,width in enumerate((24,16,11,13,13,15,16,17,15,21,16,20),2):
            ws.Columns(letter(c)+":"+letter(c)).ColumnWidth=width
        for col in m.metric_cols.values():ws.Columns(col+":"+col).ColumnWidth=18
        end=letter(13+len(m.metric_cols))
        ws.Range(f"B6:{end}6").WrapText=True
        ws.Rows("6:6").RowHeight=42
        if m.plan_rows:
            last=m.plan_rows[-1]
            ws.Rows(f"7:{last}").RowHeight=21
            entry_area=ws.Range(f"B7:{end}{last}")
            # Mark each spare input cell on the sheet's continuous yellow fill.
            _edge(entry_area, 11, SILVER)  # xlInsideVertical
            if last > START:
                _edge(entry_area, 12, SILVER)  # xlInsideHorizontal
            ws.Range(f"E7:F{last}").NumberFormat="yyyy-mm"
            set_number_format(ws.Range(f"J7:K{last}"),"0.0%")
            set_numeric_list_format(ws.Range(f"D7:D{last}"))
            for col,source in [("C","=ACT_Targets"),("D","0,1")]+([("L","=ACT_CostTargets")] if "ACT_CostTargets" in m.names else []):
                validation=ws.Range(f"{col}7:{col}{last}").Validation
                validation.Delete();validation.Add(3,1,1,source);validation.ShowError=True
    for line in m.activity_lines:
        increment = line["params"].get("increment")
        if ws.Name == m.details[line["id"]] and increment:
            address = f"{m.col(m.last+1)}{PLANNED_ACTIVITY}:{m.col(len(m.months)-1)}{PLANNED_ACTIVITY}"
            validation = ws.Range(address).Validation
            validation.Delete()
            validation.Add(1 if increment["whole"] else 2, 1, 7, 0)
            validation.IgnoreBlank = False
            validation.ShowError = True
        if ws.Name==m.details[line["id"]] and line["params"]["price_steps"]:
            end=40+len(line["params"]["price_steps"])
            ws.Range("I40").Columns.AutoFit()
            ws.Range(f"G41:G{end}").NumberFormat="yyyy-mm"
            set_number_format(ws.Range(f"H41:H{end}"),"0.0%")

