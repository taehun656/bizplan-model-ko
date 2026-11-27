"""Profile physical activity and observed metric costs before selecting a forecast driver."""
from collections import defaultdict
from math import isfinite
from statistics import mean


def metric_quantities(rows, months, activity, spec):
    """Return a monthly denominator for each metric's complete set of lines."""
    observations = {}
    observed_lines = defaultdict(set)
    for row in rows:
        metric, line, month = row["metric"], row["line"], row["month"]
        key = metric, line, month
        if month not in months:
            raise ValueError(f"metrics_monthly.csv: 손익 실적 범위 밖 {key}")
        if key in observations:
            raise ValueError(f"metrics_monthly.csv: 중복 관측 {key}")
        value = float(row["value"])
        if not isfinite(value) or value < 0:
            raise ValueError(f"metrics_monthly.csv: 유효한 0 이상 물량을 입력하세요 {key}")
        observations[key] = value
        observed_lines[metric].add(line)

    if spec is None:
        # Without a spec only the lines observed somewhere in this file are known.
        return {metric: {
            month: sum(observations[metric, line, month] for line in lines)
            for month in months if all((metric, line, month) in observations for line in lines)
        } for metric, lines in observed_lines.items()}

    definitions = {metric["id"]: metric["lines"] for metric in spec.get("metrics", [])}
    for metric, line, month in observations:
        if line not in definitions.get(metric, {}):
            raise ValueError(f"metrics_monthly.csv: 설계서에 없는 지표 또는 라인 {metric}/{line}/{month}")
    if definitions:
        from spec_io import next_month
        month = spec["first_month"]
        while month <= spec["last_actual"]:
            if month not in months:
                raise ValueError(f"is_monthly.csv: 월 실적 누락 {month}")
            month = next_month(month)
        if any(t < spec["first_month"] or t > spec["last_actual"] for t in months):
            raise ValueError("is_monthly.csv: 설계서의 실적 범위 밖 월이 있습니다.")
    result = {}
    for metric, lines in definitions.items():
        totals = {}
        for month in sorted(months):
            total = 0.0
            for line, coefficient in lines.items():
                key = metric, line, month
                source = activity.get(line, {}).get(month)
                if source is None:
                    raise ValueError(f"sales_volume.csv: 지표 산출에 필요한 활동량 누락 {key}")
                observed = observations.get(key)
                if isinstance(coefficient, dict):
                    if observed is None:
                        raise ValueError(f"metrics_monthly.csv: 실적 계수 산출용 지표 누락 {key}")
                    if source[0] == 0 and observed != 0:
                        raise ValueError(f"metrics_monthly.csv: 활동량 0에 양의 관측 지표 {key}")
                total += observed if observed is not None else source[0] * coefficient
            totals[month] = total
        result[metric] = totals
    return result


def profile_activity(sales,metrics,ledger,spec=None):
    months={v["month"] for v in ledger}
    byline=defaultdict(dict)
    for v in sales:
        if v["month"] in months and str(v.get("activity","")).strip():
            values = tuple(float(v[k]) for k in ("activity","quantity","price"))
            if v["month"] in byline[v["line"]]:
                raise ValueError(f"sales_volume.csv: 중복 실적 {v['line']} {v['month']}")
            if not all(isfinite(n) and n >= 0 for n in values) or (values[0] == 0 and values[1] > 0):
                raise ValueError(f"sales_volume.csv: 활동량·수량·단가 실적 오류 {v['line']} {v['month']}")
            byline[v["line"]][v["month"]] = values
    quantities = metric_quantities(metrics, months, byline, spec)
    if not byline:return ""
    out=["","## 활동량 및 단가 추이",""]
    for ident,obs in sorted(byline.items()):
        out.extend([f"### {ident}","","| 연도 | 관측 월 | 활동량 | 활동당 수량 | 가중 단가 (원) |",
                    "|---|---:|---:|---:|---:|"])
        for yr in sorted({t[:4] for t in obs}):
            rows=[v for t,v in obs.items() if t.startswith(yr)]
            a=sum(v[0] for v in rows);q=sum(v[1] for v in rows)
            out.append(f"| {yr} | {len(rows)} | {a:,.2f} | {q/a if a else 0:.4f} | {sum(v[1]*v[2] for v in rows)/q if q else 0:,.2f} |")
        out.extend(["","| 계절지수 (월평균=1) | "+" | ".join(f"{n}월" for n in range(1,13))+" |",
                    "|---|"+"---:|"*12])
        for title,fn in [("활동량",lambda v:v[0]),("활동당 수량",lambda v:v[1]/v[0] if v[0] else 0),("단가",lambda v:v[2])]:
            avg=mean(fn(v) for v in obs.values())
            cells=[]
            for n in range(1,13):
                sample=[fn(v) for t,v in obs.items() if int(t[-2:])==n]
                cells.append(f"{mean(sample)/avg:.3f}" if sample and avg else "자료 없음")
            out.append("| "+title+" | "+" | ".join(cells)+" |")
        changes=[];prev=None
        for t,(_,_,p) in sorted(obs.items()):
            if prev and p!=prev:changes.append(f"{t} {p/prev-1:+.2%}")
            prev=p
        out.extend(["","관측 단가 변화 월 · "+(", ".join(changes) or "변화 없음"),""])
    if not quantities:return "\n".join(out)+"\n"
    accounts=defaultdict(lambda:defaultdict(float))
    for v in ledger:
        if v["section"] in ("매출원가","판관비"):
            accounts[v["section"],v["account"]][v["month"]]+=float(v["amount"])
    scope = ("설계서에 지정된 모든 라인의 관측 물량을 합산했다. 상수 계수 라인의 관측값이 없는 달은 활동량에 계수를 곱했다."
             if spec is not None else "파일에 나타난 모든 대상 라인의 관측이 있는 달만 합산했다. 비용 계정별 후보를 표시한다.")
    out.extend(["## 지표당 원가 추이","",
        "비용 합계(원)를 같은 관측 월의 지표 합계로 나눈 값이다. " + scope,"",
        "| 지표 | 비용 | 연도 | 관측 월 | 지표당 원가 (원) |","|---|---|---|---:|---:|"])
    for metric, totals in sorted(quantities.items()):
        groups = [(key[0]+" / "+key[1], [key]) for key in accounts]
        if spec is not None:
            groups = [(line["label"], [(line["section"],a) for a in line["accounts"]]) for line in spec["cost_lines"]
                      if line.get("params",{}).get("metric")==metric]
        for label,keys in groups:
            for yr in sorted({t[:4] for t in months}):
                observed = [t for t in totals if t.startswith(yr)]
                qty=sum(totals[t] for t in observed)
                cost=sum(accounts[key].get(t,0) for key in keys for t in observed)
                val=f"{cost/qty:,.4f}" if qty else "산출 불가"
                out.append(f"| {metric} | {label} | {yr} | {len(observed)} | {val} |")
    return "\n".join(out)+"\n"

