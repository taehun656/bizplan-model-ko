"""Validate activity inputs and select historical observations for forecast baselines."""
from collections import defaultdict

from tax_policy import shift_month


def sample_months(rule, history, cutoff, target):
    mode = rule["base"]
    if mode == "manual":
        return []
    if mode == "base_month" and not rule.get("seasonal"):
        return [rule["month"]]
    start = rule.get("start", shift_month(cutoff, -11) if mode == "last_12m" else min(history))
    end = rule.get("end", cutoff)
    if mode == "last_12m":
        start, end = shift_month(cutoff, -11), cutoff
    return sorted(t for t in history if start <= t <= end and t[5:] == target[5:])


def validate_activity(spec, data):
    from build_model import timeline
    lines = {line["id"]:line for line in spec["revenue_lines"] if line["driver"] == "activity"}
    metrics = {v["id"]:v for v in spec["metrics"]}
    plans = {v["id"]:v for v in spec["initiatives"]}
    costs = {v["id"]:v for v in spec["cost_lines"]}
    if len(metrics) != len(spec["metrics"]) or len(plans) != len(spec["initiatives"]):
        raise ValueError("metrics / initiatives: 중복 ID")
    if not lines and (metrics or plans):
        raise ValueError("계획과 지표에는 activity 매출이 필요합니다.")
    cutoff = spec["last_actual"]
    months = timeline(spec["first_month"], spec["plan_end"])
    history = {t for t in months if t <= cutoff}
    future = set(months) - history
    volumes = defaultdict(dict)
    for v in data["sales_volume"]:
        volumes[v["line"]][v["month"]] = v
    for ident, line in lines.items():
        rows = volumes[ident]
        for t in history:
            v = rows.get(t, {})
            if "activity" not in v or (v["activity"] == 0 and v["quantity"] > 0):
                raise ValueError(f"{ident} {t}: 실적 활동량이 누락되었거나 수량과 맞지 않습니다.")
        for t,v in rows.items():
            if "activity" not in v or (v["activity"] == 0 and v["quantity"]):
                raise ValueError(f"{ident} {t}: activity 열의 활동량을 확인하세요.")
        for field in ("activity","yield","price"):
            rule = line["params"][field]
            if rule["base"] == "base_month" and rule.get("month") not in history:
                raise ValueError(f"{ident}.{field}: 실적 기준 월을 지정하세요.")
            if rule.get("start",min(history)) > rule.get("end",cutoff) or any(
                    rule[k] not in history for k in ("start","end") if k in rule):
                raise ValueError(f"{ident}.{field}: 기준 기간이 실적 범위를 벗어납니다.")
            if set(rule.get("monthly",{})) - future:
                raise ValueError(f"{ident}.{field}: 직접 입력은 추정 월에 지정하세요.")
            for t in future:
                if t not in rows and t not in rule.get("monthly",{}) and not sample_months(rule,history,cutoff,t):
                    raise ValueError(f"{ident}.{field} {t}: 기준 관측값 또는 월별 계획을 지정하세요.")
        for step in line["params"]["price_steps"]:
            if step["month"] > spec["plan_end"]:
                raise ValueError(f"{ident}: 단가 조정 월이 계획 종료 후입니다.")
        increment = line["params"].get("increment")
        if increment:
            for field in ("monthly", "included_monthly"):
                amounts = increment[field]
                if amounts.keys() - future or (increment["whole"] and any(int(v) != v for v in amounts.values())):
                    raise ValueError(f"{ident}.increment.{field}: 계획 월과 활동량 단위를 확인하세요.")
            for key in increment["metric_factors"]:
                if key not in metrics or ident not in metrics[key]["lines"]:
                    raise ValueError(f"{ident}.increment: 매출에 연결된 활동 지표를 지정하세요 ({key}).")
    metric_actual = {(v["metric"],v["line"],v["month"]):v["value"] for v in data["metrics_monthly"]}
    for metric,line,month in metric_actual:
        if metric not in metrics or line not in metrics[metric]["lines"] or month not in history:
            raise ValueError(f"metrics_monthly.csv: 미등록 지표·라인 또는 실적 범위 밖 ({metric}/{line}/{month})")
    for ident,metric in metrics.items():
        if not metric["lines"] or metric["lines"].keys()-lines.keys():
            raise ValueError(f"{ident}: 지표 계수는 등록된 활동 매출 라인에 지정하세요.")
        for line,coeff in metric["lines"].items():
            if isinstance(coeff,dict):
                if set(coeff.get("monthly",{}))-future:
                    raise ValueError(f"{ident}/{line}: 월별 계수는 추정 월에 지정하세요.")
                for t in history:
                    if (ident,line,t) not in metric_actual:
                        raise ValueError(f"{ident}/{line} {t}: 실적 계수 산출용 지표가 필요합니다.")
                    if volumes[line][t]["activity"] == 0 and metric_actual[ident,line,t] != 0:
                        raise ValueError(f"{ident}/{line} {t}: 활동량 0의 지표값을 확인하세요.")
                for t in future:
                    if t not in coeff.get("monthly",{}) and not sample_months({"base":"last_12m"},history,cutoff,t):
                        raise ValueError(f"{ident}/{line} {t}: 동월 실적 또는 월별 계수가 필요합니다.")
    for line in costs.values():
        p = line["params"]
        if "activity_increment" in p:
            source = lines.get(p["activity_increment"]["line"])
            if not source or not source["params"].get("increment"):
                raise ValueError(f"{line['id']}: 월별 추가 활동량이 있는 매출 라인을 지정하세요.")
            if p.get("monthly") or p.get("base_monthly"):
                raise ValueError(f"{line['id']}: 활동 증감 연동과 독립적인 월 비용을 중복 지정할 수 없습니다.")
        if line["method"] in ("per_metric","consumption") and p["metric"] not in metrics:
            raise ValueError(f"{line['id']}: 미등록 지표 {p['metric']}")
        if line["method"] == "per_activity" and set(p["base_lines"])-lines.keys():
            raise ValueError(f"{line['id']}: per_activity는 활동 매출 ID를 사용합니다.")
        if line["method"] in ("consumption", "per_metric"):
            fields = ("price", "subsidy", "efficiency") if line["method"] == "consumption" else ("unit_cost",)
            for field in fields:
                values = p[field]
                if isinstance(values, dict) and (set(values) - future or min(future) not in values):
                    raise ValueError(
                        f"{line['id']}.{field}: 첫 추정 월 {min(future)}부터 추정 범위 안의 월을 지정하세요.")
            if set(p.get("actual_usage", {})) - history:
                raise ValueError(f"{line['id']}.actual_usage: 관측 사용량은 실적 월에만 지정하세요.")
        if set(p.get("adjustments",{}))-future:
            raise ValueError(f"{line['id']}: 월별 원가 보정은 추정 월에 지정하세요.")
    dep = set(spec["investment"]["depreciation_accounts"])
    for p in plans.values():
        if p["line"] not in lines or p["start"] not in future or p.get("end",max(future)) not in future or p.get("end",max(future)) < p["start"]:
            raise ValueError(f"{p['id']}: 계획 대상 라인 또는 적용 기간 오류")
        if p.get("fixed_cost") or "cost_line" in p:
            if p.get("cost_line") not in costs or dep.intersection(costs[p["cost_line"]]["accounts"]):
                raise ValueError(f"{p['id']}: 고정비 대상은 일반 비용 라인이어야 합니다.")
        for ident in p.get("metric_factors",{}):
            if ident not in metrics or p["line"] not in metrics[ident]["lines"]:
                raise ValueError(f"{p['id']}: 해당 활동 라인에 연결된 지표를 지정하세요.")
    for obj in [*spec["scenarios"].values(),*spec["cases"]]:
        if set(obj.get("initiatives",[]))-plans.keys():
            raise ValueError("시나리오에 미등록 계획 ID가 있습니다.")
        for controls in obj.get("drivers",{}).values():
            if controls.get("efficiency_multiplier",1) <= 0:
                raise ValueError("효율 배수는 양수여야 합니다.")
