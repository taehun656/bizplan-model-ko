"""Read independent CSV observations and exercise all initiative switches in read-only Excel."""
from collections import defaultdict
from contextlib import ExitStack
from excel_retry import cleanup_excel, retry_excel
from build_model import method


def verify_activity(app,book,meta,folder,failures):
    from verify_model import months_of, read_csv, row_values
    info=meta["info"]
    months=months_of(info["first_month"],info["plan_end"])
    start=int(info["first_month_col"]);end=int(info["last_month_col"]);unit=info["unit"]
    actual={i for i,t in enumerate(months) if t<=info["last_actual"]}
    def named(n):return book.Names(n).RefersToRange
    def read(sheet,row):
        return row_values(book.Worksheets(sheet), row, start, end)
    tally=defaultdict(int);issues=0
    def check(got,want,label,where,scale=1,tolerance=1e-7):
        nonlocal issues
        tally[label]+=1
        if got is None or abs(got-want)*scale>max(tolerance,abs(want*scale)*1e-12):
            issues+=1
            if issues<=12:failures.append(f"{label} {where}: {got} / {want}")
    lines=named("ACT_Lines").Value2[1:]
    sales={(v["line"],v["month"]):v for v in read_csv(folder/"sales_volume.csv")}
    for key,sh in lines:
        a,y,q,p,rev=(read(sh,r) for r in (14,15,16,17,9))
        raw_a,raw_q,raw_p=(read(sh,r) for r in (30,31,32))
        for i in actual:
            v=sales.get((key,months[i]))
            if v is None or not str(v.get("activity","")).strip():
                failures.append(f"{key} {months[i]} 활동량 원천 누락");continue
            for got,want,label in [(a[i],float(v["activity"]),"실적 활동량"),(q[i],float(v["quantity"]),"실적 수량"),(p[i],float(v["price"]),"실적 단가")]:
                check(got,want,label,key+" "+months[i])
            for raw, field in ((raw_a, "activity"), (raw_q, "quantity"), (raw_p, "price")):
                check(raw[i],float(v[field]),"물량·단가 원천",key+" "+months[i])
            check(a[i]*y[i]*p[i],rev[i]*unit,"실적 매출",key+" "+months[i],tolerance=.5)
    metrics_file=folder/"metrics_monthly.csv"
    obs={(v["metric"],v["line"],v["month"]):float(v["value"]) for v in read_csv(metrics_file)} if metrics_file.exists() else {}
    seen=set()
    for met,line,raw,applied in named("ACT_Metrics").Value2[1:]:
        s,v=read("활동지표",int(raw)),read("활동지표",int(applied))
        for i in actual:
            key=met,line,months[i]
            if key in obs:
                seen.add(key)
                check(v[i],obs[key],"지표 실적"," / ".join(key))
                check(s[i],obs[key],"지표 원천"," / ".join(key))
            elif s[i] is not None:failures.append(f"지표 실적 CSV 누락: {key}")
    if seen!=set(obs):failures.append("원천 지표 CSV의 등록되지 않은 행을 확인하세요.")
    slots=named("ACT_Effects").Value2[1:]
    selector=named("SN_Case");original=selector.Value2
    plan=book.Worksheets("영업계획")
    switches=[(int(row),plan.Cells(int(row),4).Value2) for row,effect in slots]

    def restore_inputs():
        # Unwind every input, then the selected case, then recalculate even if a write fails.
        with ExitStack() as restore:
            restore.callback(method, app, "Calculate")
            restore.callback(retry_excel, lambda: setattr(selector, "Value2", original))
            for row, value in reversed(switches):
                restore.callback(retry_excel, lambda row=row, value=value:
                                 setattr(plan.Cells(row, 4), "Value2", value))

    with cleanup_excel(restore_inputs):
        for case in named("SN_CaseNos").Value2[0]:
            selector.Value2=case;method(app,"Calculate")
            on=[read("운영수입",8),read("운영비용",10)]
            effects=[[0.]*len(months),[0.]*len(months)]
            for row,first in slots:
                blank=not plan.Cells(int(row),2).Value2
                for j in range(2):
                    for i,val in enumerate(read("활동지표",int(first)+j)):
                        effects[j][i]+=val
                        if blank:check(val,0,"빈 계획 효과",f"{row} {months[i]}",unit)
            for row,_ in switches:plan.Cells(row,4).Value2=0
            method(app,"Calculate")
            for _,sh in lines:
                for r,base in ((14,22),(15,23),(16,24),(17,25),(18,26)):
                    for i,(got,want) in enumerate(zip(read(sh,r),read(sh,base))):
                        if i not in actual:
                            check(got,want,"전체 해제 기준",f"{sh} {months[i]} {r}행")
            for j,(sh,r) in enumerate((("운영수입",8),("운영비용",10))):
                for i,off in enumerate(read(sh,r)):
                    check(on[j][i]-off,effects[j][i],"계획 효과 대사",f"{sh} {months[i]}",unit,tolerance=.5)
            for row,v in switches:plan.Cells(row,4).Value2=v
    return ["활동 검증 · "+", ".join(f"{k} {v}건" for k,v in tally.items())+f" · 불일치 {issues}건"]

