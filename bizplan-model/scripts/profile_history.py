# -*- coding: utf-8 -*-
"""Describe monthly history. Facts only; no driver decision."""
import argparse
import calendar
import csv
import json
import math
import re
import sys
import unicodedata
from collections import defaultdict
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from common import say, write_text, wrote

SECTIONS = ("매출", "매출원가", "판관비", "영업외수익", "영업외비용", "법인세")


def read_csv(path):
    with path.open(encoding="utf-8-sig", newline="") as handle:
        return list(csv.DictReader(handle))


def num(value):
    if value is None or value == "":
        return 0.0
    return float(str(value).replace(",", ""))


def mean(values):
    return sum(values) / len(values) if values else 0.0


def stdev(values):
    if len(values) < 2:
        return 0.0
    mu = mean(values)
    return math.sqrt(sum((v - mu) ** 2 for v in values) / (len(values) - 1))


def linreg(xs, ys):
    n = len(xs)
    if n < 3:
        return None
    mx, my = mean(xs), mean(ys)
    sxx = sum((x - mx) ** 2 for x in xs)
    syy = sum((y - my) ** 2 for y in ys)
    sxy = sum((x - mx) * (y - my) for x, y in zip(xs, ys))
    if sxx == 0:
        return None
    slope = sxy / sxx
    intercept = my - slope * mx
    r2 = 0.0 if syy == 0 else (sxy * sxy) / (sxx * syy)
    return intercept, slope, r2


def money(value):
    return f"{value:,.0f}"


def pct(value):
    return f"{value * 100:.1f}%"


def growth_text(values, months):
    if len(values) >= 24 and sum(values[:12]) > 0:
        return f"최근 12개월 / 처음 12개월 - 1 = {pct(sum(values[-12:]) / sum(values[:12]) - 1)}"
    if len(values) >= 2 and values[0] > 0 and values[-1] > 0:
        years = (len(values) - 1) / 12
        rate = (values[-1] / values[0]) ** (1 / years) - 1 if years else 0
        return f"첫 달 {money(values[0])} → 마지막 달 {money(values[-1])}, 연환산 {pct(rate)}"
    return "확인 못 함"


def season_text(values, months):
    if len(set(m[5:7] for m in months)) < 12:
        return "확인 못 함 (달력이 12개월을 덮지 않음)"
    buckets = defaultdict(list)
    for month, value in zip(months, values):
        buckets[month[5:7]].append(value)
    overall = mean(values)
    if overall == 0:
        return "확인 못 함 (평균 0)"
    parts = []
    for key in sorted(buckets):
        parts.append(f"{key}:{mean(buckets[key]) / overall:.2f}")
    return " ".join(parts)


def outliers(months, values, label):
    if len(values) < 4:
        return [f"{label}: 확인 못 함 (개월 수가 적음)"]
    mu, sd = mean(values), stdev(values)
    if sd == 0:
        return [f"{label}: 편차 0"]
    hits = [f"{month} {money(value)} (평균 대비 {(value - mu) / sd:.1f} 표준편차)"
            for month, value in zip(months, values) if abs(value - mu) > 2 * sd]
    return hits or [f"{label}: 평균에서 2표준편차를 넘는 달 없음"]


def evidence_tag(amount_cv, intercept, slope, r2, mean_amount):
    if r2 is None:
        return "표본 부족"
    scale = abs(mean_amount) if mean_amount else 1
    small_intercept = abs(intercept) <= 0.15 * scale
    # A level shift can keep the monthly CV low while the recent year is fixed plus a rate.
    if abs(intercept) > 0.15 * scale and slope > 0 and r2 >= 0.3:
        return "준변동비에 가까운 수치 (절편과 기울기가 같이 있음)"
    if r2 >= 0.8 and slope > 0 and small_intercept:
        return "변동비에 가까운 수치 (결정계수 높고 절편 작음)"
    if amount_cv <= 0.15:
        return "고정비에 가까운 수치 (월 금액 변동이 작음)"
    if r2 < 0.3 and amount_cv <= 0.25:
        return "고정비에 가까운 수치 (매출과 상관이 약함)"
    return "한쪽으로 기울지 않음"


def load_is(folder):
    path = folder / "is_monthly.csv"
    if not path.exists():
        raise FileNotFoundError(path)
    rows = read_csv(path)
    need = {"month", "account", "section", "amount"}
    if not rows or need - set(rows[0]):
        raise ValueError("is_monthly.csv columns")
    return rows


def caption(name):
    return re.sub(r"[\s_().,·/\-]+", "", unicodedata.normalize("NFKC", str(name)).casefold())


def is_depreciation(name):
    key = caption(name)
    return key.endswith(("감가상각비", "무형자산상각비", "상각비")) or key in {
        "depreciation", "amortization", "depreciationexpense", "amortizationexpense",
    }


def is_interest(name):
    return caption(name) in {"이자비용", "지급이자", "interestexpense", "interest"}


def month_days(month):
    year, number = map(int, month.split("-"))
    return calendar.monthrange(year, number)[1]


def fit_against(xs, ys, amounts):
    fit = linreg(xs, ys)
    if not fit:
        return None
    intercept, slope, r2 = fit
    center = mean(amounts)
    amount_cv = (stdev(amounts) / center) if center else 0
    return intercept, slope, r2, evidence_tag(amount_cv, intercept, slope, r2, center)


def window_section(cost_accounts, by_account, section_of, months, revenue):
    lines = ["## 비용 회귀 — 연도별과 최근 12개월", ""]
    lines.append("회귀는 월 금액 = 절편 + 기울기 × 그 구간 매출 합계. 매출이 0인 달은 뺀다. 준변동이면 절편과 기울기를 같이 적는다.")
    lines.append("")
    lines.append("| 계정 | 구간 | 개월 | 절편 | 기울기 | 결정계수 | 수치 기준 |")
    lines.append("|---|---|---:|---:|---:|---:|---|")
    windows = [("최근 12개월", set(months[-12:]))]
    for year in sorted({month[:4] for month in months}):
        chosen = [month for month in months if month.startswith(year)]
        if len(chosen) >= 3:
            windows.append((f"{year}년", set(chosen)))
    for name in sorted(cost_accounts, key=lambda item: -abs(sum(by_account[item].values()))):
        for label, chosen in windows:
            pairs = [(month, rev) for month, rev in zip(months, revenue) if month in chosen and rev != 0]
            if len(pairs) < 3:
                lines.append(f"| {name} | {label} | {len(pairs)} | 확인 못 함 | 확인 못 함 | 확인 못 함 | 표본 부족 |")
                continue
            ys = [by_account[name][month] for month, _rev in pairs]
            result = fit_against([rev for _month, rev in pairs], ys, ys)
            if not result:
                lines.append(f"| {name} | {label} | {len(pairs)} | 확인 못 함 | 확인 못 함 | 확인 못 함 | 표본 부족 |")
                continue
            intercept, slope, r2, tag = result
            lines.append(f"| {name} | {label} | {len(pairs)} | {money(intercept)} | {slope:.4f} | {r2:.3f} | {tag} |")
    lines.append("")
    return lines


def account_fit_section(cost_accounts, revenue_accounts, by_account, months):
    lines = ["## 비용과 개별 매출 계정", ""]
    recent = months[-12:]
    lines.append(f"최근 {len(recent)}개월({recent[0]} ~ {recent[-1]})이다. 금액이 0 이하로만 있는 매출 계정은 대상에서 뺐다.")
    lines.append("")
    candidates = []
    for name in revenue_accounts:
        series = [by_account[name][month] for month in recent]
        if sum(series) > 0 and stdev(series) > 0:
            candidates.append(name)
    if not candidates:
        lines.append("비교할 매출 계정이 없다.")
        lines.append("")
        return lines
    lines.append("| 비용 | 가장 잘 맞는 대상 | 결정계수 | 절편 | 기울기 | 매출 합계 결정계수 |")
    lines.append("|---|---|---:|---:|---:|---:|")
    total = [sum(by_account[name][month] for name in revenue_accounts) for month in recent]
    for name in sorted(cost_accounts, key=lambda item: -abs(sum(by_account[item].values()))):
        ys = [by_account[name][month] for month in recent]
        total_fit = fit_against(total, ys, ys)
        best_name, best = "매출 합계", total_fit
        for account in candidates:
            xs = [by_account[account][month] for month in recent]
            result = fit_against(xs, ys, ys)
            if result and (best is None or result[2] > best[2]):
                best_name, best = account, result
        if not best:
            lines.append(f"| {name} | 확인 못 함 | 확인 못 함 | 확인 못 함 | 확인 못 함 | 확인 못 함 |")
            continue
        total_r2 = f"{total_fit[2]:.3f}" if total_fit else "확인 못 함"
        lines.append(f"| {name} | {best_name} | {best[2]:.3f} | {money(best[0])} | {best[1]:.4f} | {total_r2} |")
    lines.append("")
    return lines


def sales_section(sales_rows, months, revenue, by_account, revenue_accounts):
    lines = ["## 수량·단가", ""]
    if not sales_rows:
        lines.append("sales_volume.csv 없음.")
        lines.append("")
        return lines
    need = {"month", "line", "quantity", "price"}
    if need - set(sales_rows[0]):
        lines.append("sales_volume.csv 열이 부족하다. month, line, quantity, price가 필요하다.")
        lines.append("")
        return lines
    by_line = defaultdict(list)
    for row in sales_rows:
        by_line[row["line"].strip()].append(row)
    revenue_by_month = dict(zip(months, revenue))
    account_names = set(revenue_accounts)
    mismatches = []
    month_gap = []
    actual_months = set(months)
    for line, rows in by_line.items():
        actual_rows = [row for row in rows if row["month"] in actual_months]
        plan_rows = [row for row in rows if row["month"] not in actual_months]
        if actual_rows:
            quantities = [num(row["quantity"]) for row in actual_rows]
            line_months = [row["month"] for row in actual_rows]
            lines.append(f"- {line} 실적: {len(actual_rows)}개월, 수량 계절지수 {season_text(quantities, line_months)}")
            prices = [(row["month"], num(row["price"])) for row in sorted(actual_rows, key=lambda item: item["month"])]
            first_month, first_price = prices[0]
            last_month, last_price = prices[-1]
            if first_price:
                lines.append(f"  실적 단가 {first_month} {money(first_price)} → {last_month} {money(last_price)}, 차이 {pct(last_price / first_price - 1)}")
            else:
                lines.append(f"  실적 단가 {first_month} {money(first_price)} → {last_month} {money(last_price)}")
        if plan_rows:
            plan_months = sorted(row["month"] for row in plan_rows)
            lines.append(f"- {line} 계획: {plan_months[0]} ~ {plan_months[-1]}, {len(plan_rows)}개월. 실적 계절지수에는 넣지 않는다.")
        if line in account_names:
            for row in actual_rows:
                amount = num(row["quantity"]) * num(row["price"])
                actual = by_account[line][row["month"]]
                if abs(amount - actual) > 1:
                    mismatches.append((line, row["month"], amount, actual))
    by_month = defaultdict(float)
    for row in sales_rows:
        if row["month"] not in actual_months:
            continue
        by_month[row["month"]] += num(row["quantity"]) * num(row["price"])
    for month, amount in sorted(by_month.items()):
        if month in revenue_by_month:
            month_gap.append(abs(amount - revenue_by_month[month]))
    if mismatches:
        worst = max(mismatches, key=lambda item: abs(item[2] - item[3]))
        lines.append(f"- 라인 ID와 같은 손익 계정: {len(mismatches)}개월이 1원보다 어긋난다. 가장 큰 달 {worst[0]} {worst[1]}, 수량×단가 {money(worst[2])}, 손익 {money(worst[3])}.")
    elif any(line in account_names for line in by_line):
        lines.append("- 이름이 같은 손익 계정이 있는 라인은 실적 월의 수량 × 단가와 1원 이내로 같다.")
    if month_gap:
        lines.append(f"- 월별 수량 × 단가 합계와 손익 매출 합계의 차이: 최대 {money(max(month_gap))}, 평균 {money(mean(month_gap))}.")
    else:
        lines.append("- 수량 파일의 월이 손익 월과 겹치지 않아 합계 대사는 확인 못 함.")
    lines.append("")
    return lines


def contract_section(rows, first_month):
    lines = ["## 계약 목록", ""]
    if not rows:
        lines.append("revenue_items.csv 없음.")
        lines.append("")
        return lines
    lines.append(f"- 계약 {len(rows)}건")
    starts = [row.get("start", "") for row in rows if row.get("start")]
    ends = [row.get("end", "") for row in rows if row.get("end")]
    if starts and ends:
        lines.append(f"- 기간: {min(starts)} ~ {max(ends)}")
    by_line = defaultdict(list)
    for row in rows:
        by_line[row.get("line", "")].append(row)
    for line, items in sorted(by_line.items()):
        total = sum(num(item.get("amount")) for item in items)
        probs = sorted({item.get("probability", "") for item in items})
        lines.append(f"- 라인 {line or '(없음)'}: {len(items)}건, 월 금액 합계 {money(total)}, 확률 {', '.join(str(item) for item in probs)}")
    open_items = [row for row in rows if num(row.get("probability")) < 1]
    lines.append(f"- 확률 1 미만: {len(open_items)}건")
    if first_month:
        before = [row for row in rows if (row.get("start") or "") < first_month]
        during = [row for row in rows if (row.get("start") or "") >= first_month]
        counts = defaultdict(int)
        for row in during:
            counts[(row.get("start") or "")[:7]] += 1
        lines.append(f"- 실적 첫 달 {first_month}보다 시작이 이른 계약 {len(before)}건은 기존 계약이다.")
        lines.append(f"- 그 이후 시작한 계약 {len(during)}건, 시작 월 {len(counts)}개.")
    lines.extend(renewal_lines(rows))
    lines.append("")
    return lines


AR_KEYS = {"매출채권", "외상매출금", "받을어음", "매출채권대손충당금", "accountsreceivable", "tradereceivables", "ar"}
INV_KEYS = {"재고", "재고자산", "상품", "제품", "원재료", "재공품", "저장품", "inventory", "inventories"}
AP_KEYS = {"매입채무", "외상매입금", "지급어음", "accountspayable", "tradepayables", "ap"}
PPE_KEYS = {"유형자산", "유형자산순액", "토지", "건물", "구축물", "기계장치", "비품", "공구와기구", "차량운반구", "시설장치", "건설중인자산"}


def bs_balance(bs_rows, month, section, keys, accum=False):
    total = 0.0
    found = False
    for row in bs_rows:
        if row.get("month") != month or row.get("section", "").strip() != section:
            continue
        key = caption(row.get("account", ""))
        if key in keys or (accum and key.endswith("감가상각누계액")):
            total += num(row.get("amount"))
            found = True
    return total if found else None


def average_days(values):
    usable = [value for value in values if value is not None]
    if not usable:
        return "확인 못 함"
    return f"{mean(usable):.2f}일 ({len(usable)}개월)"


GROSS_KEYS = PPE_KEYS | {"소프트웨어", "무형자산", "개발비", "영업권"}


def account_kind(name):
    key = caption(name)
    if "상각누계" in key:
        return "accum"
    if key in AR_KEYS:
        return "ar"
    if key in INV_KEYS:
        return "inv"
    if key in AP_KEYS:
        return "ap"
    if key in GROSS_KEYS:
        return "gross"
    return None


def journal_snapshots(journal_rows):
    """Month-end balances. Assets rise with debits. Payables and accumulated depreciation rise with credits."""
    buckets = defaultdict(list)
    for row in journal_rows:
        month = (row.get("date") or "")[:7]
        if len(month) == 7:
            buckets[month].append(row)
    running = {"ar": 0.0, "inv": 0.0, "ap": 0.0, "gross": 0.0, "accum": 0.0}
    seen = {key: False for key in running}
    snaps = {}
    accounts = defaultdict(float)
    account_snaps = {}
    for month in sorted(buckets):
        for row in buckets[month]:
            name = (row.get("account") or "").strip()
            kind = account_kind(name)
            if not kind:
                continue
            debit, credit = num(row.get("debit")), num(row.get("credit"))
            if kind in ("ar", "inv", "gross"):
                running[kind] += debit - credit
                if kind == "gross":
                    accounts[name] += debit - credit
            else:
                running[kind] += credit - debit
                if kind == "accum":
                    accounts[name] += credit - debit
            seen[kind] = True
        snaps[month] = dict(running)
        account_snaps[month] = dict(accounts)
    notes = []
    if seen["accum"] and seen["gross"] and running["accum"] - running["gross"] > 1:
        notes.append(f"감가상각누계액 {money(running['accum'])}원이 취득가액 {money(running['gross'])}원을 넘는다.")
    return snaps, seen, notes, account_snaps


def carried_snapshots(snaps, months):
    carried = {}
    last = None
    for month in months:
        if month in snaps:
            last = snaps[month]
        if last is not None and month >= min(snaps):
            carried[month] = last
    return carried


def recent_mean(values):
    usable = [value for value in values[-12:] if value is not None]
    return mean(usable) if usable else None


def rebuilt(flow, days_value, day_count):
    if days_value is None or not day_count or flow is None:
        return None
    return flow * days_value / day_count


def fixed_parts(bs_rows, account_snaps, month):
    if bs_rows:
        gross = contra = 0.0
        found = False
        for row in bs_rows:
            if row.get("month") != month or row.get("section", "").strip() != "비유동자산":
                continue
            name = (row.get("account") or "").strip()
            amount = num(row.get("amount"))
            if "상각누계" in caption(name):
                contra += amount
                found = True
            elif caption(name) in GROSS_KEYS:
                gross += amount
                found = True
        if not found:
            return None
        return {"net": gross + contra, "gross": gross}
    chosen = None
    for key in sorted(account_snaps):
        if key <= month:
            chosen = account_snaps[key]
    if not chosen:
        return None
    gross = contra = 0.0
    for name, amount in chosen.items():
        if "상각누계" in caption(name):
            contra += amount
        else:
            gross += amount
    return {"net": gross - contra, "gross": gross}


def asset_breakdown(bs_rows, account_snaps, months, by_account, dep_names):
    lines = ["| 계정 | 금액 |", "|---|---:|"]
    for name in dep_names:
        lines.append(f"| {name} 상각 합계 | {money(sum(by_account[name][month] for month in months))} |")
    if bs_rows:
        present = [month for month in months if any(row.get("month") == month for row in bs_rows)]
    else:
        present = [month for month in months if any(key <= month for key in account_snaps)]
    if len(present) >= 2:
        start, end = present[0], present[-1]
        names = set()
        if bs_rows:
            for row in bs_rows:
                if row.get("month") in (start, end) and row.get("section", "").strip() == "비유동자산":
                    name = (row.get("account") or "").strip()
                    if caption(name) in GROSS_KEYS or "상각누계" in caption(name):
                        names.add(name)
            for name in sorted(names):
                opening = next((num(row.get("amount")) for row in bs_rows if row.get("month") == start and row.get("account") == name), 0)
                closing = next((num(row.get("amount")) for row in bs_rows if row.get("month") == end and row.get("account") == name), 0)
                lines.append(f"| {name} 증감 | {money(closing - opening)} |")
        else:
            start_snap = [account_snaps[key] for key in sorted(account_snaps) if key <= start][-1]
            end_snap = [account_snaps[key] for key in sorted(account_snaps) if key <= end][-1]
            for name in sorted(set(start_snap) | set(end_snap)):
                lines.append(f"| {name} 증감 | {money(end_snap.get(name, 0) - start_snap.get(name, 0))} |")
    return lines


def base_line_suggestions(months, by_account, section_of, observed):
    del months
    if len(observed) < 3:
        return []
    bundles = []
    cogs = [name for name, section in section_of.items() if section == "매출원가"]
    if cogs:
        bundles.append(("매출원가 전체", cogs))
    ex_dep = [name for name in cogs if not is_depreciation(name)]
    if ex_dep and len(ex_dep) != len(cogs):
        bundles.append((" + ".join(ex_dep), ex_dep))
    core = [name for name in cogs if any(word in name for word in ("원재료", "외주", "소모품"))]
    if len(core) >= 2:
        bundles.append((" + ".join(core), core))
    for name, section in section_of.items():
        if section in ("매출원가", "판관비") and sum(abs(by_account[name][row[0]]) for row in observed):
            bundles.append((name, [name]))

    def score(members, balance_at, add_inventory):
        samples = []
        prior = None
        for row in observed:
            month, _sales, _cost, _ex, _ar, inventory, _ap, _purchases, day = row
            flow = sum(by_account[name][month] for name in members)
            if add_inventory and prior is not None and inventory is not None:
                flow += inventory - prior
            prior = inventory
            balance = row[balance_at]
            if balance is None or not day or flow <= 0:
                continue
            samples.append((flow, balance, day))
        if len(samples) < 3:
            return None
        avg = mean([balance / flow * day for flow, balance, day in samples])
        error = mean([abs(max(flow, 0) * avg / day - balance) for flow, balance, day in samples])
        scale = mean([abs(balance) for _flow, balance, _day in samples])
        if scale <= 0:
            return None
        return avg, error / scale

    lines = ["- 비용 묶음 후보. 지정한 지표의 기준액은 해당 비용 라인 금액의 직접 합계다."]
    for title, index, add_inventory in (("재고", 5, False), ("매입채무", 6, False)):
        ranked = []
        for label, members in bundles:
            result = score(members, index, add_inventory)
            if result:
                ranked.append((result[1], result[0], label))
        if not ranked:
            continue
        ranked.sort()
        eligible = [item for item in ranked if 5 <= item[1] <= 120]
        pool = eligible or ranked
        shown = [pool[0]]

        def core_bundle(label):
            parts = [part.strip() for part in label.split("+")]
            return len(parts) >= 2 and all(any(word in part for word in ("원재료", "외주", "소모품")) for part in parts)

        for item in pool[1:]:
            if core_bundle(item[2]) and item[2] != shown[0][2]:
                shown.append(item)
                break

        def day_text(days):
            return f"{days:.0f}" if abs(days - round(days)) < 0.05 else f"{days:.1f}"

        for error, days, label in shown:
            lines.append(f"  - {title}: {label}의 {day_text(days)}일, 오차 {error * 100:.1f}%")
    return lines if len(lines) > 1 else []


def party_key(row):
    name = (row.get("item") or "").strip()
    return re.sub(r"_20\d{2}갱신$", "", name)


def renewal_lines(rows):
    by_end = defaultdict(list)
    for row in rows:
        end = (row.get("end") or "")[:7]
        if len(end) == 7:
            by_end[end].append(row)
    lines = []
    total = max(len(rows), 1)
    last_end = max(by_end) if by_end else ""
    for end, items in sorted(by_end.items()):
        if len(items) < 3 and len(items) / total < 0.3:
            continue
        year, number = map(int, end.split("-"))
        number += 1
        if number == 13:
            year, number = year + 1, 1
        nxt = f"{year:04d}-{number:02d}"
        historical = planned = 0
        for item in items:
            key = party_key(item)
            if not key:
                continue
            for other in rows:
                if other is item or party_key(other) != key:
                    continue
                start = (other.get("start") or "")[:7]
                if start not in (end, nxt):
                    continue
                note = other.get("note") or ""
                if "가정" in note or num(other.get("probability")) < 1:
                    planned += 1
                else:
                    historical += 1
                break
        churn = len(items) - historical - planned
        if end == last_end and historical == 0 and planned == 0:
            lines.append(f"- {end}에 {len(items)}건이 함께 만료된다. 자료에서 가장 늦은 만료월이라 그 뒤 갱신 행은 없다.")
        else:
            lines.append(
                f"- {end}에 {len(items)}건이 함께 만료된다. "
                f"같은 거래처로 이어진 실적 갱신 {historical}건, 계획 갱신 {planned}건, 이탈 {churn}건."
            )
    return lines


def gap_text(actual, rebuilt_value):
    if actual is None or rebuilt_value is None:
        return "확인 못 함"
    if abs(actual) < 1:
        return f"실제 {money(actual)}, 역산 {money(rebuilt_value)}"
    return f"실제 {money(actual)}, 역산 {money(rebuilt_value)}, 차이 {abs(rebuilt_value / actual - 1) * 100:.1f}%"


def working_capital_section(bs_rows, journal_rows, months, by_account, section_of, revenue):
    lines = ["## 운전자본 회전일수", ""]
    lines.append("월말 잔액과 같은 달의 흐름을 쓴다. 일수는 `working_capital.day_basis`에 따라 해당 달의 달력 일수 또는 30일이며 기초·기말 평균 잔액을 쓰지 않는다. 분자와 분모의 금액 단위를 맞추면 원이나 표시 단위 어느 쪽으로도 같은 일수가 나온다.")
    lines.append('기준 라인 목록이 없으면 재고의 기준액은 상각비·비현금 비용을 포함한 손익의 매출원가 전체다. 매입채무의 기본 기준액은 매출원가 전체에 재고 증감을 더한 매입액이다. `working_capital.ap_base="purchases"`(기본)는 매출원가 + 당월말 재고 − 전월말 재고, `"cogs"`는 매출원가를 쓴다.')
    lines.append("이 프로파일은 설계서의 기본값인 calendar, 곧 해당 달의 달력 일수로 계산한다.")
    lines.append("")
    snaps, seen, notes, _accounts = journal_snapshots(journal_rows)
    source = "bs" if bs_rows else "journal"
    if bs_rows:
        balance_months = {row.get("month") for row in bs_rows}
        lines.append("잔액은 bs_monthly.csv다.")
    elif any(seen.values()):
        balance_months = set(carried_snapshots(snaps, months))
        found = [name for name, flag in (("매출채권", seen["ar"]), ("재고", seen["inv"]), ("매입채무", seen["ap"])) if flag]
        lines.append("bs_monthly.csv가 없어 분개 잔액을 월말까지 누적했다. 잡은 계정: " + (", ".join(found) if found else "없음") + ".")
        lines.extend(f"- {note}" for note in notes)
    else:
        lines.append("bs_monthly.csv가 없고, 분개에서도 매출채권·재고·매입채무를 찾지 못했다.")
        lines.append("")
        return lines

    def point(month, kind):
        if source == "bs":
            section, keys = {
                "ar": ("유동자산", AR_KEYS),
                "inv": ("유동자산", INV_KEYS),
                "ap": ("유동부채", AP_KEYS),
            }[kind]
            return bs_balance(bs_rows, month, section, keys)
        row = carried_snapshots(snaps, months).get(month)
        if row is None or not seen[kind]:
            return None
        return row[kind]

    cogs_names = [name for name, section in section_of.items() if section == "매출원가"]
    dep_names = [name for name in cogs_names if is_depreciation(name)]
    dso, dio, dpo, dpo_cogs = [], [], [], []
    observed = []
    prior_inventory = None
    for month, sales in zip(months, revenue):
        if month not in balance_months:
            prior_inventory = None
            continue
        days = month_days(month)
        ar, inventory, ap = point(month, "ar"), point(month, "inv"), point(month, "ap")
        cost = sum(by_account[name][month] for name in cogs_names)
        cost_ex_dep = cost - sum(by_account[name][month] for name in dep_names)
        dso_m = (ar / sales * days) if ar is not None and sales > 0 else None
        dio_m = (inventory / cost * days) if inventory is not None and cost > 0 else None
        dso.append(dso_m)
        dio.append(dio_m)
        dpo_cogs.append((ap / cost * days) if ap is not None and cost > 0 else None)
        purchases = None
        if ap is not None and prior_inventory is not None and inventory is not None:
            purchases = cost + inventory - prior_inventory
            dpo.append((ap / purchases * days) if purchases > 0 else None)
        else:
            dpo.append(None)
        observed.append((month, sales, cost, cost_ex_dep, ar, inventory, ap, purchases, days))
        prior_inventory = inventory
    recent = slice(-12, None)
    lines.append(f"- 매출채권 회수일 평균 {average_days(dso)}. 최근 12개월 {average_days(dso[recent])}.")
    lines.append(f"- 재고 보유일 평균 {average_days(dio)}. 최근 12개월 {average_days(dio[recent])}.")
    lines.append(f"- 매입채무 지급일(매입액) 평균 {average_days(dpo)}. 최근 12개월 {average_days(dpo[recent])}.")
    lines.append(f"- 매입채무 지급일(원가) 평균 {average_days(dpo_cogs)}. 최근 12개월 {average_days(dpo_cogs[recent])}.")
    if observed:
        month, sales, cost, cost_ex_dep, ar, inventory, ap, purchases, days = observed[-1]
        lines.append(f"- 마감 {month}에 최근 12개월 평균 일수를 그 달 흐름에 다시 적용한 잔액:")
        lines.append(f"  - 매출채권 {gap_text(ar, rebuilt(sales, recent_mean(dso), days))}")
        lines.append(f"  - 재고 {gap_text(inventory, rebuilt(cost, recent_mean(dio), days))}")
        lines.append(f"  - 매입채무(매입액) {gap_text(ap, rebuilt(purchases, recent_mean(dpo), days))}")
    lines.extend(base_line_suggestions(months, by_account, section_of, observed))
    lines.append("")
    return lines


def capex_section(bs_rows, journal_rows, months, by_account, section_of):
    lines = ["## 과거 투자 지출", ""]
    snaps, seen, notes, account_snaps = journal_snapshots(journal_rows)
    if not bs_rows and not seen.get("gross"):
        lines.append("bs_monthly.csv가 없고, 분개에서도 고정자산 계정을 찾지 못했다.")
        lines.append("")
        return lines
    if not bs_rows:
        lines.append("bs_monthly.csv가 없어 분개의 취득가액과 감가상각누계액으로 순액을 만들었다.")
        lines.extend(f"- {note}" for note in notes)
    dep_names = [name for name, section in section_of.items() if section in ("매출원가", "판관비") and is_depreciation(name)]
    series = []
    gross_series = []
    prior_net = prior_gross = None
    if not bs_rows and snaps and months:
        earlier = [month for month in snaps if month < months[0]]
        if earlier:
            seed = snaps[max(earlier)]
            prior_net = seed["gross"] - seed["accum"]
            prior_gross = seed["gross"]
    bs_months = {row.get("month") for row in bs_rows} if bs_rows else set(carried_snapshots(snaps, months))
    for month in months:
        if month not in bs_months:
            prior_net = prior_gross = None
            continue
        parts = fixed_parts(bs_rows, account_snaps, month)
        net = None if parts is None else parts["net"]
        gross = None if parts is None else parts["gross"]
        dep = sum(by_account[name][month] for name in dep_names)
        if net is not None and prior_net is not None:
            series.append((month, net - prior_net + dep))
        if gross is not None and prior_gross is not None:
            gross_series.append(gross - prior_gross)
        prior_net, prior_gross = net, gross
    if not series:
        names = sorted({row.get("account", "") for row in bs_rows if row.get("section", "").strip() == "비유동자산"})
        lines.append("유형자산 순액을 두 달 이상 찾지 못했다. 비유동자산 계정: " + ", ".join(names))
        lines.append("")
        return lines
    values = [amount for _month, amount in series]
    lines.append("추정 투자 지출 = 유형·무형자산 순장부가 증가 + 그 기간의 상각. 순장부가는 취득가액에서 상각누계액을 뺀 값이다.")
    lines.extend(asset_breakdown(bs_rows, account_snaps, months, by_account, dep_names))
    lines.append(f"- 개월 {len(values)}, 월평균 {money(mean(values))}, 최소 {money(min(values))}, 최대 {money(max(values))}")
    if gross_series:
        lines.append(f"- 유형·무형자산 취득가액 증감 월평균 {money(mean(gross_series))}. 처분이 없으면 취득가액 증감이 투자 지출이다.")
    lines.append("")
    lines.append("| 월 | 추정 투자 지출 |")
    lines.append("|---|---:|")
    for month, amount in series:
        lines.append(f"| {month} | {money(amount)} |")
    lines.append("")
    return lines


def non_operating_section(months, by_account, section_of):
    lines = ["## 영업외", ""]
    income = [sum(by_account[name][month] for name, section in section_of.items() if section == "영업외수익") for month in months]
    other = [sum(by_account[name][month] for name, section in section_of.items() if section == "영업외비용" and not is_interest(name)) for month in months]
    interest = [sum(by_account[name][month] for name, section in section_of.items() if section == "영업외비용" and is_interest(name)) for month in months]
    lines.append(f"- 영업외수익 월평균 {money(mean(income))}. 최근 12개월 {money(mean(income[-12:]))}.")
    lines.append(f"- 이자 외 영업외비용 월평균 {money(mean(other))}. 최근 12개월 {money(mean(other[-12:]))}.")
    lines.append(f"- 이자비용 월평균 {money(mean(interest))}. 최근 12개월 {money(mean(interest[-12:]))}.")
    lines.append("")
    return lines


def tax_section(months, by_account, section_of):
    lines = ["## 실적 법인세 비율", ""]
    lines.append("모델은 이 비율이 아니라 법인세법 세율표로 계산한다 (tax_rules_kr.md).")
    lines.append("")

    def section_amount(month, section):
        return sum(by_account[name][month] for name, found in section_of.items() if found == section)

    for year in sorted({month[:4] for month in months}):
        chosen = [month for month in months if month.startswith(year)]
        pretax = 0.0
        tax = 0.0
        for month in chosen:
            pretax += (
                section_amount(month, "매출") - section_amount(month, "매출원가") - section_amount(month, "판관비")
                + section_amount(month, "영업외수익") - section_amount(month, "영업외비용")
            )
            tax += section_amount(month, "법인세")
        if pretax > 0:
            lines.append(f"- {year}: 법인세 {money(tax)} / 세전이익 {money(pretax)} = {pct(tax / pretax)}")
        else:
            lines.append(f"- {year}: 세전이익 {money(pretax)}, 법인세 {money(tax)}. 세전이익이 0 이하여서 비율은 확인 못 함.")
    lines.append("")
    return lines


def profile(rows, journal_rows, sales_rows, contract_rows, bs_rows):
    months = sorted({row["month"] for row in rows})
    by_account = defaultdict(lambda: defaultdict(float))
    section_of = {}
    for row in rows:
        section = row["section"].strip()
        account = row["account"].strip()
        by_account[account][row["month"]] += num(row["amount"])
        section_of[account] = section
    revenue_accounts = [name for name, section in section_of.items() if section == "매출"]
    cost_accounts = [name for name, section in section_of.items() if section in ("매출원가", "판관비", "영업외비용", "법인세")]
    revenue = [sum(by_account[name][month] for name in revenue_accounts) for month in months]
    cogs = [sum(by_account[name][month] for name in section_of if section_of[name] == "매출원가") for month in months]
    sga = [sum(by_account[name][month] for name in section_of if section_of[name] == "판관비") for month in months]
    op = [r - c - s for r, c, s in zip(revenue, cogs, sga)]
    total_revenue = sum(revenue)
    lines = ["# 실적 프로파일", "", "판단은 하지 않는다. 아래는 정규화 손익과 분개에서 계산한 수치다.", ""]
    lines.append(f"- 기간: {months[0]} ~ {months[-1]} ({len(months)}개월)")
    lines.append(f"- 매출 합계: {money(total_revenue)}")
    lines.append(f"- 영업이익 합계 (매출−매출원가−판관비): {money(sum(op))}")
    lines.append("")
    lines.append("## 매출 계정")
    lines.append("")
    lines.append("| 계정 | 합계 | 비중 | 월 변동계수 | 성장 |")
    lines.append("|---|---:|---:|---:|---|")
    for name in sorted(revenue_accounts, key=lambda n: -abs(sum(by_account[n].values()))):
        series = [by_account[name][month] for month in months]
        share = (sum(series) / total_revenue) if total_revenue else 0
        cv = (stdev(series) / mean(series)) if mean(series) else 0
        lines.append(f"| {name} | {money(sum(series))} | {pct(share)} | {cv:.2f} | {growth_text(series, months)} |")
    lines.append("")
    lines.append("매출 합계 계절지수 (달 평균 / 전체 평균): " + season_text(revenue, months))
    lines.append("")
    lines.append("매출 합계 성장: " + growth_text(revenue, months))
    lines.append("")
    lines.append("## 비용 계정")
    lines.append("")
    lines.append("회귀는 월 금액 = 절편 + 기울기 × 그 달 매출 합계. 매출이 0인 달은 회귀에서 뺀다.")
    lines.append("수치 기준: 결정계수 0.8 이상이고 절편이 평균 금액의 15% 이하면 변동비에 가깝다. 월 금액 변동계수 0.15 이하면 고정비에 가깝다. 절편과 기울기가 둘 다 있고 결정계수가 0.3 이상이면 준변동비에 가깝다.")
    lines.append("")
    lines.append("| 계정 | 구분 | 합계 | 매출 대비 | 비율 변동계수 | 매출 상관 | 절편 | 기울기 | 결정계수 | 수치 기준 |")
    lines.append("|---|---|---:|---:|---:|---:|---:|---:|---:|---|")
    usable = [(m, r) for m, r in zip(months, revenue) if r != 0]
    for name in sorted(cost_accounts, key=lambda n: -abs(sum(by_account[n].values()))):
        series = [by_account[name][month] for month in months]
        ratios = [by_account[name][month] / rev for month, rev in usable]
        xs = [rev for month, rev in usable]
        ys = [by_account[name][month] for month, rev in usable]
        fit = linreg(xs, ys)
        corr = 0.0
        if fit and stdev(xs) and stdev(ys):
            corr = sum((x - mean(xs)) * (y - mean(ys)) for x, y in zip(xs, ys)) / ((len(xs) - 1) * stdev(xs) * stdev(ys))
        amount_cv = (stdev(series) / mean(series)) if mean(series) else 0
        ratio_cv = (stdev(ratios) / mean(ratios)) if ratios and mean(ratios) else 0
        share = (sum(series) / total_revenue) if total_revenue else 0
        if fit:
            intercept, slope, r2 = fit
            tag = evidence_tag(amount_cv, intercept, slope, r2, mean(series))
            fit_text = (money(intercept), f"{slope:.4f}", f"{r2:.3f}")
        else:
            tag = "표본 부족"
            fit_text = ("확인 못 함", "확인 못 함", "확인 못 함")
        lines.append(
            f"| {name} | {section_of[name]} | {money(sum(series))} | {pct(share)} | {ratio_cv:.2f} | {corr:.2f} | {fit_text[0]} | {fit_text[1]} | {fit_text[2]} | {tag} |"
        )
    lines.append("")
    lines.append("## 이익률과 이상 월")
    lines.append("")
    margins = [(op_v / rev) if rev else None for op_v, rev in zip(op, revenue)]
    known = [m for m in margins if m is not None]
    if len(known) >= 2:
        half = len(known) // 2
        lines.append(f"- 영업이익률 평균 {pct(mean(known))}. 앞쪽 {pct(mean(known[:half]))}, 뒤쪽 {pct(mean(known[half:]))}.")
    else:
        lines.append("- 영업이익률: 확인 못 함")
    lines.append("- " + " ".join(outliers(months, revenue, "매출")))
    lines.append("- " + " ".join(outliers(months, op, "영업이익")))
    lines.append("")
    lines.extend(window_section(cost_accounts, by_account, section_of, months, revenue))
    lines.extend(account_fit_section(cost_accounts, revenue_accounts, by_account, months))
    lines.extend(sales_section(sales_rows, months, revenue, by_account, revenue_accounts))
    lines.extend(contract_section(contract_rows, months[0] if months else ""))
    lines.extend(working_capital_section(bs_rows, journal_rows, months, by_account, section_of, revenue))
    lines.extend(capex_section(bs_rows, journal_rows, months, by_account, section_of))
    lines.extend(non_operating_section(months, by_account, section_of))
    lines.extend(tax_section(months, by_account, section_of))
    lines.extend(journal_section(journal_rows, months, set(revenue_accounts), bool(contract_rows)))
    lines.extend(driver_notes(revenue, months, journal_rows, sales_rows, contract_rows))
    return "\n".join(lines) + "\n"


def journal_section(rows, months, revenue_accounts, use_contracts):
    lines = ["## 분개장", ""]
    if not rows:
        lines.append("journal.csv 없음.")
        lines.append("")
        return lines
    by_party = defaultdict(float)
    by_party_month = defaultdict(lambda: defaultdict(float))
    amounts = []
    first_seen = {}
    for row in rows:
        account = row.get("account", "").strip()
        if account not in revenue_accounts:
            continue
        credit, debit = num(row.get("credit")), num(row.get("debit"))
        amount = credit - debit
        if amount == 0:
            continue
        party = (row.get("counterparty") or "").strip() or "(거래처 없음)"
        month = (row.get("date") or "")[:7]
        by_party[party] += amount
        if month:
            by_party_month[party][month] += amount
            if party not in first_seen or month < first_seen[party]:
                first_seen[party] = month
        if credit > 0:
            amounts.append(credit)
    total = sum(by_party.values())
    ranked = sorted(by_party.items(), key=lambda item: -item[1])[:10]
    lines.append(f"- 매출 계정과 이름이 같은 분개의 순액 합계: {money(total)}")
    lines.append("")
    lines.append("| 순위 | 거래처 | 순액 | 비중 |")
    lines.append("|---:|---|---:|---:|")
    for i, (party, amount) in enumerate(ranked, 1):
        lines.append(f"| {i} | {party} | {money(amount)} | {pct(amount / total) if total else '확인 못 함'} |")
    recurring_parties = []
    recurring_amount = 0.0
    for party, month_map in by_party_month.items():
        series = list(month_map.values())
        if len(series) < 3 or mean(series) == 0:
            continue
        if stdev(series) / abs(mean(series)) <= 0.15:
            recurring_parties.append(party)
            recurring_amount += by_party[party]
    lines.append("")
    lines.append(f"- 3개월 이상 나오고 월 금액 변동계수가 0.15 이하인 거래처: {len(recurring_parties)}곳 / {len(by_party)}곳")
    lines.append(f"- 그 거래처의 매출 비중: {pct(recurring_amount / total) if total else '확인 못 함'}")
    if amounts:
        amounts.sort()
        def q(p):
            pos = min(len(amounts) - 1, max(0, int(round((len(amounts) - 1) * p))))
            return amounts[pos]
        lines.append(f"- 매출 대변 건당 금액: 건수 {len(amounts)}, 평균 {money(mean(amounts))}, 10% {money(q(0.1))}, 중앙값 {money(q(0.5))}, 90% {money(q(0.9))}")
    else:
        lines.append("- 매출 대변 건당 금액: 확인 못 함")
    parties = {party: month for party, month in first_seen.items() if party != "(거래처 없음)"}
    if use_contracts:
        lines.append("- 신규 거래처는 계약 목록의 시작 월로 세었다.")
    elif parties:
        opening_month = min(parties.values())
        opening = [party for party, month in parties.items() if month == opening_month]
        newcomers = {party: month for party, month in parties.items() if month != opening_month}
        later = [month for month in months if month > opening_month]
        span = len(later) or len(months) or 1
        lines.append(f"- 첫 관측 월 {opening_month}의 거래처 {len(opening)}곳은 기존 거래처로 보고 유입에서 뺐다.")
        lines.append(f"- 그 이후 처음 나타난 거래처: {len(newcomers)}곳. 이후 월평균 {len(newcomers) / span:.2f}곳.")
    else:
        lines.append("- 신규 거래처 유입: 확인 못 함")
    lines.append("")
    return lines


def driver_notes(revenue, months, journal_rows, sales_rows, contract_rows):
    cv = (stdev(revenue) / mean(revenue)) if mean(revenue) else 0
    lines = ["## 동인 유형별 근거 요약", ""]
    present = []
    if sales_rows:
        present.append("수량·단가 파일")
    if contract_rows:
        present.append("계약 목록")
    if journal_rows:
        present.append("분개")
    lines.append("이 입력에 있는 자료: " + (", ".join(present) if present else "손익만") + ".")
    lines.append("")
    lines.append("### growth")
    lines.append(f"- 지지하는 관찰: 월 매출 변동계수 {cv:.2f}.")
    lines.append("- 반대하는 관찰: 계절지수가 1에서 멀거나, 이상 월이 있으면 단일 성장률로 설명되지 않는다. 위의 계절지수와 이상 월을 같이 본다.")
    lines.append("")
    lines.append("### price_volume")
    if sales_rows:
        lines.append("- 지지하는 관찰: 월별 수량과 단가 파일이 있다. 수량 계절지수와 단가 추이는 위 수량·단가 절에 있다.")
        lines.append("- 반대하는 관찰: 수량 × 단가가 손익 매출과 어긋나면 에누리·다른 매출이 빠져 있거나 라인이 계정과 다르다.")
    else:
        lines.append("- 지지하는 관찰: 손익·분개만으로는 확인 못 함. 건당 금액이 작고 건수가 많으면 반복 판매와 어울린다.")
        lines.append("- 반대하는 관찰: 이 폴더의 정규화 파일에는 수량과 단가 열이 없다.")
    lines.append("")
    lines.append("### contracts")
    if contract_rows:
        lines.append(f"- 지지하는 관찰: 계약 목록 {len(contract_rows)}건. 시작·종료·월 금액·확률이 위 계약 절에 있다.")
        lines.append("- 반대하는 관찰: 같은 금액이 거의 모든 달에 반복되면 건별 착수·종료보다 구독에 가깝다.")
    elif journal_rows:
        lines.append("- 지지하는 관찰: 거래처 이름이 있고, 상위 비중과 처음 등장 달이 위에 있다. 일부 거래처만 특정 달에 크면 건별 계약과 어울린다.")
        lines.append("- 반대하는 관찰: 계약 목록 파일은 없다. 같은 금액이 거의 모든 달에 반복되면 건별 착수·종료보다 구독에 가깝다.")
    else:
        lines.append("- 지지하는 관찰: 확인 못 함 (분개장과 계약 목록 없음).")
        lines.append("- 반대하는 관찰: 거래처·계약 기간이 이 입력에 없다.")
    lines.append("")
    lines.append("### cohort")
    if journal_rows:
        lines.append("- 지지하는 관찰: 같은 거래처의 월 금액 변동계수가 0.15 이하로 여러 달 반복되면 고객 수 × 객단가와 어울린다.")
        lines.append("- 반대하는 관찰: 거래처가 적거나, 월 금액이 크게 들쭉날쭉하면 평균 이탈률로 뭉개기 어렵다.")
    else:
        lines.append("- 지지하는 관찰: 확인 못 함 (분개장 없음).")
        lines.append("- 반대하는 관찰: 고객별 반복 청구를 셀 수 없다.")
    lines.append("")
    return lines


def main():
    parser = argparse.ArgumentParser(description="Profile normalized monthly history.")
    parser.add_argument("--data", required=True, type=Path)
    parser.add_argument("--out", required=True, type=Path)
    parser.add_argument("--spec", type=Path, help="지표당 원가에 사용할 비용 계정 설계서")
    args = parser.parse_args()
    try:
        rows = load_is(args.data)
    except FileNotFoundError:
        say("FAIL is_monthly.csv missing")
        return 2
    except ValueError:
        say("FAIL is_monthly.csv columns")
        return 2
    def optional(name):
        path = args.data / name
        return read_csv(path) if path.exists() else []

    journal_rows = optional("journal.csv")
    sales_rows = optional("sales_volume.csv")
    contract_rows = optional("revenue_items.csv")
    bs_rows = optional("bs_monthly.csv")
    bad = [row["section"].strip() for row in rows if row["section"].strip() not in SECTIONS]
    if bad:
        say("FAIL unknown section")
        return 2
    from activity_profile import profile_activity
    spec = json.loads(args.spec.read_text(encoding="utf-8-sig")) if args.spec else None
    report = profile(rows, journal_rows, sales_rows, contract_rows, bs_rows)
    report += profile_activity(sales_rows, optional("metrics_monthly.csv"), rows, spec)
    write_text(args.out, report)
    say(f"MONTHS {len({row['month'] for row in rows})}")
    say(f"JOURNAL {len(journal_rows)}")
    wrote(args.out, "profile")
    say("RESULT PASS")
    return 0


if __name__ == "__main__":
    sys.exit(main())
