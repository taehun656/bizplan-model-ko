"""Financial statements and supporting schedules for the monthly model builder."""

import math
import re
import unicodedata
from collections import defaultdict
from datetime import date

from spec_io import is_interest_account
from tax_accrual import build_tax_accrual
from tax_cash import build_tax_cash, opening_accounts
from cash_items import add_terms, build_cash_items, WAGE_CHANGE


def _number(value, field, minimum=None):
    if value in (None, ""):
        value = 0
    try:
        result = float(str(value).replace(",", "").strip())
    except (TypeError, ValueError) as exc:
        raise ValueError(f"{field}: 숫자가 필요합니다 ({value!r}).") from exc
    if not math.isfinite(result) or (minimum is not None and result < minimum):
        raise ValueError(f"{field}: 허용 범위를 벗어난 값입니다 ({value!r}).")
    return result


def _month(value, field):
    s = str(value).strip()
    if not re.fullmatch(r"[0-9]{4}-(0[1-9]|1[0-2])", s):
        raise ValueError(f"{field}: YYYY-MM 형식이 필요합니다 ({s!r}).")
    return s


def _name(value):
    return re.sub(r"[\s_().,·/\-]+", "", unicodedata.normalize("NFKC", str(value)).casefold())


BS_SECTIONS = ("유동자산", "비유동자산", "유동부채", "비유동부채", "자본")
IS_SECTIONS = ("매출", "매출원가", "판관비", "영업외수익", "영업외비용", "법인세")
# Match the entire caption within its source section. Unrecognized captions
# retain their signed balance and source classification.
GROUPS = {
    8: ("유동자산", "현금", "현금 현금및현금성자산 현금성자산 보통예금 당좌예금 요구불예금 cash cashandcashequivalents"),
    9: ("유동자산", "매출채권", "매출채권 외상매출금 받을어음 매출채권대손충당금 accountsreceivable tradereceivables ar"),
    10: ("유동자산", "재고자산", "재고 재고자산 상품 제품 원재료 재공품 저장품 inventory inventories"),
    11: ("유동자산", "기타 유동자산", ""),
    12: ("비유동자산", "유형자산 순액", "유형자산 유형자산순액 토지 건물 구축물 기계장치 비품 공구와기구 차량운반구 시설장치 건설중인자산 감가상각누계액 유형자산감가상각누계액 ppe netppe propertyplantandequipment accumulateddepreciation"),
    13: ("비유동자산", "기타 비유동자산", ""),
    14: ("유동부채", "매입채무", "매입채무 외상매입금 지급어음 accountspayable tradepayables ap"),
    15: ("유동부채", "유동 차입금", "차입금 단기차입 단기차입금 단기사채 유동성장기부채 유동성장기차입금 유동성사채 유동리스부채 단기리스부채 shorttermdebt shorttermborrowings currentdebt currentportionoflongtermdebt"),
    16: ("유동부채", "기타 유동부채", ""),
    17: ("비유동부채", "비유동 차입금", "차입금 장기차입 장기차입금 장기사채 사채 리스부채 비유동리스부채 장기리스부채 장기부채 longtermdebt longtermborrowings bonds leaseliabilities"),
    18: ("비유동부채", "기타 비유동부채", ""),
    19: ("자본", "자본", ""),
}
FALLBACK = dict(zip(BS_SECTIONS, (11, 13, 16, 18, 19)))
ALIASES = {(section, _name(alias)): row for row, (section, _, names) in GROUPS.items()
           for alias in names.split()}


def _bucket(section, account):
    caption = _name(account)
    if section == "비유동자산" and caption.endswith("감가상각누계액"):
        return 12
    return ALIASES.get((section, caption), FALLBACK[section])


def _source_rows(m):
    """Keep BS account detail plus independently aggregated source controls."""
    months = m.months
    actual_months = set(months[:m.last + 1])
    ledger = defaultdict(lambda: defaultdict(float))
    totals = defaultdict(lambda: defaultdict(float))
    observed = set()
    for row in (m.data.get("bs_monthly", []) if m.bs else []):
        month = _month(row["month"], "bs_monthly.month")
        if month not in actual_months:
            continue
        section, account = row["section"].strip(), row["account"].strip()
        if section not in BS_SECTIONS or not account:
            raise ValueError(f"bs_monthly: 계정 또는 구분을 확인하세요 ({month}, {section}, {account}).")
        amount = _number(row["amount"], f"bs_monthly {month} {account}")
        ledger[section, account][month] += amount
        totals[section][month] += amount
        observed.add(month)
    if m.bs and months[m.last] not in observed:
        raise ValueError("BS 추정을 시작하려면 실적 마감 월의 bs_monthly 잔액이 필요합니다.")

    if m.bs:
        sh = "실적_BS"
        m.bs_source_rows = {}
        members = defaultdict(list)
        tax_members = defaultdict(list)
        accounts = opening_accounts(m.spec["tax"])
        for r, (section, account) in enumerate(sorted(ledger, key=lambda k: (_bucket(*k), k)), 34):
            amounts = ledger[section, account]
            group = _bucket(section, account)
            members[group].append(r)
            m.bs_source_rows[section, account] = r
            for component, aliases in accounts.items():
                expected = "유동부채" if component.endswith("payable") else "유동자산"
                if section == expected and _name(account) in aliases:
                    tax_members[component].append(r)
            m.row(sh, r, f"{section} / {GROUPS[group][1]} / {account} (원)",
                  [amounts.get(t, 0) if t in observed else None for t in months], kind="actual", annual="last")
        for r, (_, label, _) in GROUPS.items():
            values = []
            for i, t in enumerate(months):
                c = m.col(i)
                indices = members[r]
                expr = f"SUM({c}{indices[0]}:{c}{indices[-1]})" if indices else "0"
                values.append(f"=({expr})/Unit" if t in observed else None)
            m.row(sh, r, label, values, annual="last")
        for r, component, label in ((20,"national_payable","당기 국세 미지급액"),
                                     (22,"local_payable","당기 지방세 미지급액"),
                                     (23,"national_prepaid","당기 국세 선납액"),
                                     (24,"local_prepaid","당기 지방세 선납액")):
            m.row(sh, r, label, [
                "=(" + ("+".join(f"{m.col(i)}{n}" for n in tax_members[component]) or "0") + ")/Unit"
                if t in observed else None for i,t in enumerate(months)], annual="last")
        m.row(sh,21,"당기법인세 선납·환급자산",[
            f"={m.col(i)}23+{m.col(i)}24" if t in observed else None
            for i,t in enumerate(months)],annual="last")
        for r, section in enumerate(BS_SECTIONS, 109):
            m.row("검증", r, f"원천 합계 / {section} (원)",
                  [totals[section][t] if t in observed else None for t in months], kind="actual", annual="last")

    sums = defaultdict(lambda: defaultdict(float))
    dep_accounts = set(m.spec["investment"]["depreciation_accounts"])
    observed_is = set()
    for row in m.data.get("is_monthly", []):
        t = _month(row["month"], "is_monthly.month")
        if t not in actual_months:
            continue
        section = row["section"].strip()
        if section not in IS_SECTIONS:
            raise ValueError(f"is_monthly: 알 수 없는 손익 구분 ({section}).")
        observed_is.add(t)
        amount = _number(row["amount"], f"is_monthly {t} {row['account']}")
        sums[section][t] += amount
        if section in IS_SECTIONS[1:3] and row["account"] in dep_accounts:
            sums["상각비"][t] += amount
        if section == "영업외비용" and is_interest_account(row["account"]):
            sums["이자비용"][t] += amount
    for r, label in enumerate((*IS_SECTIONS, "상각비", "이자비용"), 101):
        m.row("검증", r, f"손익 원천 합계 / {label} (원)",
              [sums[label][t] if t in observed_is else None for t in months], kind="actual")
    return observed, observed_is


def _facilities(m, finance):
    sources = finance.get("facilities") or m.data.get("debt", [])
    installments = m.data.get("debt_schedule") or finance.get("schedule", [])
    facilities, repayments = {}, defaultdict(lambda: defaultdict(float))
    cutoff = m.months[m.last]
    for item in sources:
        key = str(item["facility"]).strip()
        typ = str(item.get("type", "")).strip()
        if not key or key in facilities or typ not in ("term", "bond", "lease", "revolver"):
            raise ValueError(f"차입 약정 이름 중복·누락 또는 유형 오류: {key!r}, {typ!r}")
        maturity = str(item.get("maturity") or "").strip()
        if maturity:
            # Normalizers may preserve a contractual day of maturity.
            if len(maturity) == 10:
                maturity = date.fromisoformat(maturity).isoformat()[:7]
            maturity = _month(maturity, f"{key} 만기")
        balance = _number(item.get("balance"), f"{key} 기초 잔액", 0)
        rate = _number(item.get("rate"), f"{key} 연이율", 0)
        if balance and maturity and maturity <= cutoff:
            raise ValueError(f"{key}: 실적 마감일 이전 만기의 미상환 잔액을 확인하세요.")
        facilities[key] = (typ, balance, rate, maturity, item.get("lender") or "", item.get("at_maturity"))
    for item in installments:
        key = str(item["facility"]).strip()
        if key not in facilities:
            raise ValueError(f"debt_schedule: 약정 목록에 없는 차입입니다 ({key}).")
        t = _month(item["month"], f"{key} 상환월")
        amount = _number(item.get("principal"), f"{key} {t} 상환액", 0)
        if amount and facilities[key][3] and t > facilities[key][3]:
            raise ValueError(f"{key}: 만기 후 상환 스케줄 ({t})")
        repayments[key][t] += amount
    for key, (_, balance, _, _, _, _) in facilities.items():
        if sum(amount for t, amount in repayments[key].items() if t > cutoff) > balance + .5:
            raise ValueError(f"{key}: 계획 상환 합계가 실적 마감 잔액을 초과합니다.")
    return facilities, repayments


def build_finance(m):
    """Populate finance sheets on the shared J-based time axis."""
    months, last = m.months, m.last
    if not 0 <= last < len(months):
        raise ValueError("실적 마감 월이 모델의 월 범위 안에 있어야 합니다.")
    inv = m.spec.get("investment", {})
    wc = m.spec.get("working_capital", {})
    fin = m.spec.get("financing", {})
    observed, observed_is = _source_rows(m)
    facilities, schedule = _facilities(m, fin)
    mode = inv.get("method", "straight_line")
    if mode not in ("straight_line", "declining"):
        raise ValueError("investment.method: straight_line 또는 declining이 필요합니다.")
    life = _number(inv.get("life_months", 60), "life_months", 1)
    if not life.is_integer():
        raise ValueError("life_months: 정수 월수가 필요합니다.")
    def param(sheet, r, label, value, description, upper=None):
        value = _number(value, label, 0)
        if upper is not None and value > upper:
            raise ValueError(f"{label}: {upper} 이하의 값이 필요합니다.")
        m.param(sheet, r, label, value, sheet, description)

    def row(sheet, r, label, formula, annual="sum"):
        values = []
        for i, t in enumerate(months):
            c = m.col(i)
            p = m.col(i - 1) if i else None
            expr = formula(i, t, c, p)
            values.append("=" + expr if expr is not None else None)
        m.row(sheet, r, label, values, annual=annual)

    def forecast(sheet, r, label, fn, annual="sum"):
        row(sheet, r, label, lambda i,t,c,p: fn(i,t,c,p) if i > last else None, annual)

    def check(r, label, formula, tolerance="0.5/Unit"):
        sheet = "검증"
        row(sheet, r, "대사 차이 / " + label, formula, "max")
        m.add_check(label, f'=IFERROR(IF(MAX(\'{sheet}\'!J{r}:{m.col(len(months)-1)}{r})<={tolerance},"OK","NO"),"NO")')

    invest, funding, terms, income, cash, detail = "투자비", "재원조달", "금융조건", "IS", "CF", "BS명세"
    param(invest, 20, "신규 자산 내용연수 (개월)", life, "상각 개시 월부터 내용연수 동안 정액상각")
    param(invest, 21, "신규 자산 연간 정률", inv.get("annual_rate"), "월 정률은 연율의 1/12", 1)
    param(invest, 22, "상각 방법 (정액 0, 정률 1)", int(mode == "declining"), "신규 취득자산에 적용", 1)
    param(invest, 23, "기존 자산 월 상각액 (원)", inv.get("existing_monthly_dep"), "설계서의 월 상각액을 기존 순장부가액까지 적용합니다. fixed_assets.csv는 설계서 작성 근거입니다.")
    param(invest, 24, "기초 자산 순액 가정 (원)", inv.get("opening_net_assets"), "BS 사용 시 마감 월 유형자산 순액을 사용합니다. 양수 가정만 원천과 대사합니다.")
    capex = inv.get("monthly_capex", {})
    for t, amount in capex.items():
        _month(t, "monthly_capex")
        _number(amount, f"monthly_capex {t}", 0)
    m.input(invest, 25, "신규 자산 취득액 (원)",
            [_number(capex.get(t), t, 0) if i > last else None for i,t in enumerate(months)],
            "투자비", "월 취득액은 현금 지출로 반영")
    forecast(invest, 8, "투자 지출", lambda i,t,c,p: f"{c}25")
    row(invest, 9, "감가상각비", lambda i,t,c,p: f"'실적_IS'!{c}{m.actual_dep_row}" if i <= last else f"{c}11+{c}13")
    row(invest, 10, "기존 자산 기말 순액", lambda i,t,c,p: (
        f"'실적_BS'!{c}12" if m.bs and t in observed else ("$G$24/Unit" if i == last else None)
    ) if i <= last else f"{p}10-{c}11", "last")
    forecast(invest, 11, "기존 자산 상각", lambda i,t,c,p: f"MIN(MAX(0,{p}10),$G$23/Unit)")
    forecast(invest, 12, "신규 자산 기말 순액", lambda i,t,c,p: f"{p}12+{c}8-{c}13", "last")
    param(invest, 26, "신규 자산 상각 개시 (0 취득월, 1 다음 월)", int(inv.get("dep_starts","next_month") == "next_month"),
          "신규 취득분에 적용하며 기존 자산 상각과 구분합니다.", 1)
    # Each eligible cohort enters the straight-line window once and leaves
    # after its useful life. Declining balance uses the eligible opening NBV.
    first = m.col(last+1)
    def new_depreciation(i,t,c,p,acquisitions=8,closing_balance=12):
        def straight(end):
            region = f"${first}{acquisitions}:{end}{acquisitions}"
            return f"SUM(INDEX({region},1,MAX(1,COLUMNS({region})-$G$20+1)):{end}{acquisitions})/$G$20"
        same = straight(c)
        delayed = straight(p) if i > last+1 else "0"
        eligible = f"({p}{closing_balance}+IF($G$26=0,{c}{acquisitions},0))"
        return f"MIN({eligible},IF($G$22=1,{eligible}*$G$21/12,IF($G$26=0,{same},{delayed})))"
    forecast(invest, 13, "신규 자산 상각", new_depreciation)
    if inv.get("asset_classes"):
        from investment_assets import build_asset_classes
        build_asset_classes(m, row, forecast, new_depreciation)
    row(invest, 14, "유형자산 기말 순액", lambda i,t,c,p: f"SUM({c}10,{c}12)" if t in observed or i >= last else None, "last")

    for r, key, label in [(12,"min_cash","최소 보유현금"), (13,"revolver_limit","한도대출 총 한도"),
                          (14,"opening_cash","실적 마감 월 현금 가정")]:
        param(terms, r, label+" (원)", fin.get(key), "BS 사용 시 기초 현금은 원천 기준이며 양수 가정만 대사합니다." if r==14 else label)
    param(terms, 15, "신규 조달 연이율", fin.get("new_rate"), "전월 말 신규 조달 잔액에 적용")
    param(terms, 16, "추가 한도 사용 연이율", fin.get("revolver_rate"), "전월 말 추가 한도 사용액에 적용")
    for r,key in enumerate(("dso","dio","dpo"), 17):
        param(terms, r, key.upper()+" (일)", wc.get(key), "계획 월의 실제 일수를 기준으로 계산")
    param(terms, 25, "운전자본 일수 기준 (0 달력, 1 월 30일)", int(wc.get("day_basis") == "month30"),
          "매출채권·재고·매입채무에 공통 적용", 1)

    param(terms, 22, "추가 조달 스위치 (1 사용, 0 중지)", int(fin["funding_enabled"]),
          "추가 한도 인출 및 신규 차입에 적용", 1)
    param(terms, 23, "만기 처리 (0 상환, 1 연장)", int(fin["maturity_treatment"] == "rollover"),
          "약정 원금은 예정대로 상환하고 만기 잔액에 적용", 1)
    if not wc.get("dpo_base_lines"):
        param(terms, 24, "매입채무 산정 기준 (0 매입액, 1 원가)", int(wc.get("ap_base") == "cogs"),
              "매입액은 보고 매출원가에 재고자산 증감을 더한 금액입니다.", 1)
    m.cell(terms, 27, 9, "만기 처리", "text")
    m.section(terms, 26, "차입 약정과 월별 상환 예정 원금")
    m.table(terms, 27, 2, [["약정", "대주", "유형", "마감 잔액", "만기", "연이율"]])
    m.header(terms, 27)

    # Each facility has its own opening balance, contractual schedule, maturity
    # balloon and interest. CSV takes precedence over embedded copies of it.
    principal, balloons, ends, interest, revolvers, due, deviations = [], [], [], [], [], [], []
    def next_year(t):
        year, month = map(int, t.split("-"))
        return f"{year + 1:04d}-{month:02d}"

    for n, (key, (typ, balance, rate, maturity, lender, policy)) in enumerate(facilities.items()):
        src, base, tr = 12+3*n, 42+7*n, 28+n
        sh = "실적_차입"
        label = f"{key} / {lender} / {typ} / 만기 {maturity or '미지정'}"
        m.row(sh, src, label+" / 마감 잔액 (원)", [balance if i == last else None for i in range(len(months))], kind="actual", annual="last")
        m.row(sh, src+1, label+" / 약정 상환 (원)", [schedule[key].get(t,0) for t in months], kind="actual")
        # Amount due in the next twelve months includes payments beyond the model horizon.
        m.row(sh, src+2, label+" / 향후 12개월 약정 상환 (원)",
              [sum(v for k,v in schedule[key].items() if t < k <= next_year(t)) for t in months], kind="actual", annual="last")
        m.table(terms, tr, 2, [[key, lender, typ]])
        m.cell(terms, tr, 5, f"='{sh}'!{m.col(last)}{src}/Unit")
        m.cell(terms, tr, 6, maturity, "text")
        m.cell(terms, tr, 7, rate, "input")
        m.register(terms, f"G{tr}", key + " 연이율", terms, "전월 말 잔액에 적용")
        m.cell(terms, tr, 9, {None: -1, "repay": 0, "rollover": 1}[policy], "input")
        m.register(terms, f"I{tr}", key + " 만기 처리", terms, "-1 전체 스위치, 0 상환, 1 연장")
        for i in range(len(months)):
            c = m.col(i)
            m.cell(terms, tr, 10 + i, f"='{sh}'!{c}{src+1}/Unit")
        for j, indices in enumerate(m.years.values()):
            first_col, last_col = m.col(indices[0]), m.col(indices[-1])
            m.cell(terms, tr, 10 + len(months) + j, f"=SUM({first_col}{tr}:{last_col}{tr})")
        m.cell(terms, tr, 8, f"={m.col(len(months) + len(m.years) - 1)}{tr}")
        forecast(funding, base, key+" / 기초", lambda i,t,c,p: f"'{sh}'!{p}{src}/Unit" if i==last+1 else f"{p}{base+3}", "last")
        forecast(funding, base+1, key+" / 약정 상환", lambda i,t,c,p: f"MIN({c}{base},'{terms}'!{c}{tr})")
        override = f"'{terms}'!$I${tr}"
        maturity_policy = f"IF({override}=-1,'{terms}'!$G$23,{override})"
        forecast(funding, base+2, key+" / 만기 잔액 상환", lambda i,t,c,p:
                 f"IF({maturity_policy}=0,{c}{base}-{c}{base+1},0)" if t==maturity else "0")
        row(funding, base+3, key+" / 기말", lambda i,t,c,p: f"'{sh}'!{c}{src}/Unit" if i==last else (
            f"{c}{base}-{c}{base+1}-{c}{base+2}" if i>last else None), "last")
        forecast(funding, base+4, key+" / 이자", lambda i,t,c,p: f"{c}{base}*'{terms}'!$G${tr}/12")
        forecast(funding, base+5, key+" / 유동성 분류", lambda i,t,c,p: f"{c}{base+3}" if typ=="revolver" or (maturity and maturity<=next_year(t))
                 else f"MIN({c}{base+3},'{sh}'!{c}{src+2}/Unit)", "last")
        deviation = 145 + n
        forecast("검증", deviation, key+" / 약정 상환 차이", lambda i,t,c,p:
                 f"ABS('재원조달'!{c}{base+1}-'{sh}'!{c}{src+1}/Unit)", "max")
        principal.append(base+1)
        balloons.append(base+2)
        ends.append(base+3)
        interest.append(base+4)
        due.append(base+5)
        deviations.append(deviation)
        if typ == "revolver":
            revolvers.append(base+3)
    def total(c, rows):
        return "+".join(f"{c}{r}" for r in rows) or "0"

    m.row("검증", 114, "약정 마감 잔액 원천 합계 (원)",
          [sum(f[1] for f in facilities.values()) if i == last else None for i in range(len(months))],
          kind="actual", annual="last")
    m.row("검증", 115, "약정 상환 스케줄 원천 합계 (원)",
          [sum(payments.get(t, 0) for payments in schedule.values()) for t in months], kind="actual")
    # Rows after revenue detail shift together for any number of business lines.
    revenues = m.spec["revenue_lines"]
    isr = m.is_rows = {"revenue": 8}
    cursor = 9 + len(revenues)
    for key in ("cogs", "gross", "sga", "op", "nonop_income", "nonop_expense",
                "interest", "pretax", "tax", "net"):
        isr[key] = cursor
        cursor += 1
    m.section(income, 7, "손익계산서")
    m.section(income, cursor + 1, "참고 지표")
    for offset, key in enumerate(("ebitda", "dep", "gross_margin", "op_margin", "ebitda_margin"), 2):
        isr[key] = cursor + offset
    m.section(income, cursor + 19, "사업연도 법인세 계산")
    m.section(income, cursor + 15, "영업외 손익 내역")
    nonop = m.non_operating_rows = {"income_monthly": cursor + 16, "other_expense_monthly": cursor + 17}
    # The assumptions finalizer links forecast columns to centrally editable scalar or monthly inputs.
    for field, label, source in (("income_monthly", "영업외수익 세부 내역", 11),
                                 ("other_expense_monthly", "기타 영업외비용 (이자 제외)", 12)):
        def historical_nonop(i, t, c, p):
            if i > last:
                return None
            value = f"'실적_IS'!{c}{source}"
            if field == "other_expense_monthly":
                value += f"-'실적_IS'!{c}{m.actual_interest_row}"
            return value
        row(income, nonop[field], label, historical_nonop)


    for i, line in enumerate(revenues):
        row(income, 9+i, line["label"], lambda j,t,c,p: f"'운영수입'!{c}{12+i}")
        m.level(income, 9+i, 4)
    for key, label, sheet, source in (
        ("revenue", "매출", "운영수입", 8), ("cogs", "매출원가", "운영비용", 8),
        ("sga", "판매관리비", "운영비용", 9), ("dep", "감가상각비", "투자비", 9)):
        row(income, isr[key], label, lambda i,t,c,p: f"'{sheet}'!{c}{source}")
    for key, label, lhs, rhs, operator in (
        ("gross", "매출총이익", "revenue", "cogs", "-"), ("op", "영업이익", "gross", "sga", "-"),
        ("net", "당기순이익", "pretax", "tax", "-"), ("ebitda", "EBITDA", "op", "dep", "+")):
        row(income, isr[key], label, lambda i,t,c,p: f"{c}{isr[lhs]}{operator}{c}{isr[rhs]}")
        m.total(income, isr[key])
    row(income, isr["nonop_income"], "영업외수익", lambda i,t,c,p: f"{c}{nonop['income_monthly']}")
    row(income, isr["interest"], "이자비용", lambda i,t,c,p:
        f"'실적_IS'!{c}{m.actual_interest_row}" if i<=last else f"'재원조달'!{c}13")
    m.level(income, isr["interest"], 4)
    row(income, isr["nonop_expense"], "영업외비용", lambda i,t,c,p:
        f"{c}{isr['interest']}+{c}{nonop['other_expense_monthly']}")
    row(income, isr["pretax"], "세전이익", lambda i,t,c,p:
        f"{c}{isr['op']}+{c}{isr['nonop_income']}-{c}{isr['nonop_expense']}")
    build_tax_accrual(m, row, check, isr, cursor+20)
    build_tax_cash(m, row, forecast, check, isr)
    if "rural_tax" in m.tax_rows:
        # Historical tax expense is reported as one line in the source P&L.
        m.tax_rows["rural_expense"] = cursor
        forecast(income, cursor, "법인세비용 중 당월 농어촌특별세 (추정)", lambda i,t,c,p: f"'CF'!{c}37")
    cash_items = build_cash_items(m, row, forecast, observed)
    forecast(funding, 11, "약정 상환 및 만기 상환", lambda i,t,c,p: total(c,principal+balloons))
    row(funding, 12, "기존 약정 기말 잔액", lambda i,t,c,p: total(c,ends) if i>=last else None, "last")
    forecast(funding, 13, "총 이자비용 (기초 잔액 기준)", lambda i,t,c,p:
             f"{total(c,interest)}+{p}18*'{terms}'!$G$16/12+{p}9*'{terms}'!$G$15/12")
    forecast(funding, 14, "추가 조달 전 현금", lambda i,t,c,p: add_terms(
             f"'{cash}'!{c}12+'{cash}'!{c}13+'{cash}'!{c}14-{c}11", c, cash_items["financing"], "'CF'!"), "last")
    forecast(funding, 15, "최소 현금 대비 부족액", lambda i,t,c,p: f"MAX(0,'{terms}'!$G$12/Unit-{c}14)", "max")
    forecast(funding, 16, "한도대출 추가 인출", lambda i,t,c,p: f"IF('{terms}'!$G$22=1,MIN({c}15,MAX(0,'{terms}'!$G$13/Unit-{c}21-{p}18)),0)")
    forecast(funding, 17, "추가 인출분 잉여현금 상환", lambda i,t,c,p:
             f"MIN(MAX(0,{p}18),MAX(0,{c}14-'{terms}'!$G$12/Unit))")
    # These balances track forecast draws only. Historical facilities, including
    # drawn revolvers, already enter opening debt through rows 12 and 21.
    row(funding, 18, "추가 한도 사용 기말 잔액", lambda i,t,c,p:
        "0" if i == last else (f"{p}18+{c}16-{c}17" if i > last else None), "last")
    forecast(funding, 19, "신규 조달 필요액 (당월 인출 가정)", lambda i,t,c,p: f"IF('{terms}'!$G$22=1,{c}15-{c}16,0)")
    row(funding, 9, "신규 조달 필요액 누적 잔액", lambda i,t,c,p:
        "0" if i == last else (f"{p}9+{c}19" if i > last else None), "last")
    row(funding, 21, "기존 한도 약정 사용 잔액", lambda i,t,c,p: total(c,revolvers) if i>=last else None, "last")
    # Optional debt file. Carry recognized opening BS debt at its historical
    # classification when there are no facility terms; no invented amortization.
    row(funding, 22, "상환 조건 미제공 유동 차입금", lambda i,t,c,p: (
        f"'실적_BS'!{c}15" if m.bs and not facilities and t in observed else (f"{p}22" if i>last else None)), "last")
    row(funding, 23, "상환 조건 미제공 비유동 차입금", lambda i,t,c,p: (
        f"'실적_BS'!{c}17" if m.bs and not facilities and t in observed else (f"{p}23" if i>last else None)), "last")
    row(funding, 8, "기말 차입금", lambda i,t,c,p: (f"'실적_BS'!{c}15+'실적_BS'!{c}17" if t in observed and m.bs
        else (f"{c}12" if i==last else None)) if i<=last else f"SUM({c}12,{c}18,{c}9,{c}22,{c}23)", "last")
    forecast(funding, 24, "유동 차입금 (신규 조달은 단기 가정)", lambda i,t,c,p:
             f"{total(c,due)}+{c}18+{c}9+{c}22", "last")
    forecast(funding, 10, "DSCR (세후 영업현금 + 이자) / 원리금", lambda i,t,c,p:
             f'IF(SUM({c}11,{c}13,{c}17)>0,(\'CF\'!{c}13+{c}13)/SUM({c}11,{c}13,{c}17),"")', "none")

    for col in [8, *range(10 + len(months), 10 + len(months) + len(m.years))]:
        c = "H" if col == 8 else m.col(col - 10)
        m.cell(funding, 10, col, f'=IFERROR((CF!{c}13+{c}13)/SUM({c}11,{c}13,{c}17),"")')

    for key, label, numerator in (("gross_margin", "매출총이익률", "gross"),
                                  ("op_margin", "영업이익률", "op"),
                                  ("ebitda_margin", "EBITDA율", "ebitda")):
        def margin(c):
            return f'IFERROR({c}{isr[numerator]}/{c}{isr["revenue"]},0)'
        r = isr[key]
        row(income, r, label, lambda i,t,c,p: margin(c), "none")
        for year_index in range(len(m.years)):
            col = len(months) + year_index
            m.cell(income, r, 10 + col, "=" + margin(m.col(col)))
        m.cell(income, r, 8, "=" + margin("H"))
        m.format_row(income, r, "0.0%")
    m.total(income, isr["pretax"])
    m.total(income, isr["revenue"])

    m.cell(cash, 28, 2, "실적 현금 증감은 연속된 BS 월말 잔액으로 확인합니다. 활동별 현금흐름은 계획 월부터 계산합니다.")
    m.cell(funding, 27, 2, "기존 한도 사용액도 총 한도를 사용합니다. 잉여현금은 추가 한도 인출분부터 상환합니다.")
    m.cell(funding, 28, 2, "추가 조달을 사용하면 한도 초과 부족액을 신규 단기 차입으로 충당합니다.")

    # CF holds core WC balances also when the balance sheet is omitted.
    dio_rows = [m.cost_summary_rows[k] for k in wc.get("dio_base_lines", [])]
    dpo_rows = [m.cost_summary_rows[k] for k in wc.get("dpo_base_lines", [])]
    cash_ids = set(wc.get("cash_revenue_lines", []))
    cash_revenue_rows = [r for r, line in enumerate(revenues, 9) if line["id"] in cash_ids]
    inventory_basis = 29 if dio_rows else 26
    for r, group, label, activity, days in ((20,9,"매출채권",isr["revenue"],17),(21,10,"재고자산",inventory_basis,18),(22,14,"매입채무",25,19)):
        prefix = "'IS'!" if r == 20 else ""
        def projected_balance(i, t, c, p):
            if m.bs and t in observed:
                return f"'실적_BS'!{c}{group}"
            if i > last or (i == last and not m.bs):
                base = f"{prefix}{c}{activity}"
                if r == 20:
                    base += "".join(f"-'IS'!{c}{cash_row}" for cash_row in cash_revenue_rows)
                return f"MAX(0,{base})*'{terms}'!$G${days}/IF('{terms}'!$G$25=1,30,DAY(EOMONTH({c}$5,0)))"
        row(cash, r, label, projected_balance, "last")
    def operating_cost_sum(c, rows):
        return "+".join(f"'운영비용'!{c}{r}" for r in rows) or "0"

    # The separately settled premium bypasses inventory/AP. Keep all other
    # costs in the selected turnover base, including wage increases earned
    # in and after the settlement month.
    retro_rows = getattr(m, "wage_retro_recognized_rows", {})
    def retro_in_base(ids):
        return [retro_rows[key] for key in ids if key in retro_rows]

    def turnover_base(i, c, base, premiums):
        if i <= last:
            return base
        return base + "".join(f"-'인건비'!{c}{r}" for r in premiums)

    cogs_premiums = retro_in_base(line["id"] for line in m.spec["cost_lines"]
                                 if line["section"] == "매출원가")
    dio_premiums = retro_in_base(wc.get("dio_base_lines", []))
    dpo_premiums = retro_in_base(wc.get("dpo_base_lines", []))

    for line in m.spec["cost_lines"]:
        if line.get("noncash") and set(line["accounts"]).intersection(inv["depreciation_accounts"]):
            raise ValueError(f"{line['id']}: 투자비 상각 계정은 비현금 비용으로 중복 가산할 수 없습니다.")
    forecast(cash, 24, "비현금 비용 가산 (상각 제외)", lambda i,t,c,p:
             operating_cost_sum(c, m.noncash_cost_rows))
    if dpo_rows:
        row(cash, 25, "매입채무 산정 기준액 (선택 비용 합계)", lambda i,t,c,p:
            turnover_base(i, c, operating_cost_sum(c, dpo_rows), dpo_premiums))
    else:
        row(cash, 25, "매입채무 산정 기준액", lambda i,t,c,p:
            f"IF('{terms}'!$G$24=1,{c}26,{c}27)")
    if dio_rows:
        row(cash, 29, "재고 산정 기준액 (선택 비용 합계)", lambda i,t,c,p:
            turnover_base(i, c, operating_cost_sum(c, dio_rows), dio_premiums))
    basis_label = "회전일수 기준 원가 (매출원가 중 소급 정산분 제외)" if cogs_premiums else "회전일수 기준 원가 (보고 매출원가)"
    row(cash,26,basis_label,lambda i,t,c,p: turnover_base(i, c, f"'IS'!{c}{isr['cogs']}", cogs_premiums))
    row(cash,27,"매입 추정액 (원가 + 재고 증감)",lambda i,t,c,p:f"{c}26"+(
        f"+{c}21-{p}21" if i>last or (i and t in observed and months[i-1] in observed) else ""))
    forecast(cash,23,"운전자본 증감의 현금 효과",lambda i,t,c,p:f"{p}20-{c}20+{p}21-{c}21+{c}22-{p}22")
    row(cash,8,"기말 현금",lambda i,t,c,p: (f"'실적_BS'!{c}8" if m.bs and t in observed else (
        f"'{terms}'!$G$14/Unit" if i==last else None)) if i<=last else f"SUM({c}12:{c}15)", "last")
    row(cash,12,"기초 현금",lambda i,t,c,p:f"{p}8" if i>last or (m.bs and i>0 and months[i-1] in observed and t in observed) else None,"last")
    forecast(cash,13,"영업현금흐름",lambda i,t,c,p:add_terms(
             f"'IS'!{c}{isr['net']}+'IS'!{c}{isr['dep']}+{c}23+{c}24+{c}34-{c}33", c, cash_items["operating"]))
    forecast(cash,14,"투자현금흐름",lambda i,t,c,p:add_terms(f"-'투자비'!{c}8", c, cash_items["investing"]))
    forecast(cash,15,"재무현금흐름",lambda i,t,c,p:add_terms(
             f"'재원조달'!{c}16-'재원조달'!{c}17+'재원조달'!{c}19-'재원조달'!{c}11", c, cash_items["financing"]))
    row(cash,16,"현금 증감 (실적은 BS 잔액 차이)",lambda i,t,c,p:f"SUM({c}13:{c}15)" if i>last else (
        f"{c}8-{p}8" if m.bs and i>0 and t in observed and months[i-1] in observed else None))

    if m.bs:
        plans = {8:("'CF'!",8),9:("'CF'!",20),10:("'CF'!",21),12:("'투자비'!",14),
                     14:("'CF'!",22),15:("'재원조달'!",24)}
        for r, (_,label,_) in GROUPS.items():
            def detail_formula(i,t,c,p):
                if i<=last:
                    return f"'실적_BS'!{c}{r}" if t in observed else None
                if r in plans:
                    sh, rr = plans[r]
                    return f"{sh}{c}{rr}"
                if r in (11,16):
                    balance = 32 if r==11 else 31
                    carried = f"{p}{r}+'CF'!{c}{balance}-'CF'!{p}{balance}"
                    if r == 16 and WAGE_CHANGE in cash_items["operating"]:
                        carried += f"+'CF'!{c}{WAGE_CHANGE}"
                    return add_terms(carried, c, [cash_items["bs"][r]] if r in cash_items["bs"] else [], "'CF'!")
                if r==17:
                    return f"'재원조달'!{c}8-{c}15"
                if r==18:
                    return add_terms(f"{p}18+'CF'!{c}24", c,
                                     [cash_items["bs"][r]] if r in cash_items["bs"] else [], "'CF'!")
                if r==19:
                    if any(m.spec.get("equity", {}).values()):
                        return f"SUM({c}24:{c}26)"
                    return f"{p}19+'IS'!{c}{isr['net']}"
                return add_terms(f"{p}{r}", c, [cash_items["bs"][r]] if r in cash_items["bs"] else [], "'CF'!")
            row(detail,r,label,detail_formula,"last")
        for r, cfrow, actualrow, label in ((21,31,None,"기타 유동부채 중 당기법인세부채"),
                                           (22,32,21,"기타 유동자산 중 선납·환급 법인세")):
            row(detail,r,label,lambda i,t,c,p:f"'CF'!{c}{cfrow}" if i>=last else (
                (f"'실적_BS'!{c}{actualrow}" if actualrow else f"'실적_BS'!{c}20+'실적_BS'!{c}22")
                if t in observed else None),"last")
        for r, (a,b), label in zip(range(12,17),((8,11),(12,13),(14,16),(17,18),(19,19)),BS_SECTIONS):
            row("BS",r,label,lambda i,t,c,p:f"SUM('BS명세'!{c}{a}:{c}{b})" if i>last or t in observed else None,"last")
        for r,label,fn in ((8,"자산총계",lambda c:f"{c}12+{c}13"),(9,"부채총계",lambda c:f"{c}14+{c}15"),
                           (10,"자본총계",lambda c:f"{c}16"),(11,"부채와 자본총계",lambda c:f"{c}9+{c}10")):
            row("BS",r,label,lambda i,t,c,p:fn(c) if i>last or t in observed else None,"last")
        check(120, "BS 대차 일치", lambda i,t,c,p:
              f"ABS('BS'!{c}8-'BS'!{c}11)" if i>last or t in observed else None)
        check(121, "실적 BS 5개 구분 원천 대사", lambda i,t,c,p: "+".join(
              f"ABS('BS'!{c}{12+k}-{c}{109+k}/Unit)" for k in range(5)) if t in observed else None)
        if facilities:
            check(122, "약정 기초 잔액과 BS 차입금 대사", lambda i,t,c,p:
                  f"ABS('재원조달'!{c}12-'재원조달'!{c}8)" if i==last else None)
        check(123, "기초 현금 가정과 BS 대사", lambda i,t,c,p:
              f"IF('{terms}'!$G$14=0,0,ABS('CF'!{c}8-'{terms}'!$G$14/Unit))" if i==last else None)
        check(124, "기초 유형자산 가정과 BS 대사", lambda i,t,c,p:
              f"IF('투자비'!$G$24=0,0,ABS('투자비'!{c}10-'투자비'!$G$24/Unit))" if i==last else None)
    check(125, "실적 손익 원천 대사", lambda i,t,c,p: "+".join(
          f"ABS('IS'!{c}{isr[key]}-{c}{101+k}/Unit)" for k,key in enumerate(
              ("revenue","cogs","sga","nonop_income","nonop_expense","tax","dep","interest"))) if t in observed_is else None)
    def actual_profit_difference(i, t, c, p):
        if t not in observed_is:
            return None
        operating = f"({c}101-{c}102-{c}103)"
        return (f"ABS('IS'!{c}{isr['op']}-{operating}/Unit)"
                f"+ABS('IS'!{c}{isr['ebitda']}-({operating}+{c}107)/Unit)"
                f"+ABS('IS'!{c}{isr['net']}-({operating}+{c}104-{c}105-{c}106)/Unit)")
    check(126, "실적 EBITDA·영업이익·순이익 원천 대사", actual_profit_difference)
    check(127, "실적 운영수입·비용 원천 대사", lambda i,t,c,p:
          f"ABS('운영수입'!{c}8-{c}101/Unit)+ABS('운영비용'!{c}8-{c}102/Unit)+ABS('운영비용'!{c}9-{c}103/Unit)"
          if t in observed_is else None)
    def cash_difference(i,t,c,p):
        if i <= last:
            return None
        finance_flow = add_terms(f"'재원조달'!{c}8-'재원조달'!{p}8", c, cash_items["financing"], "'CF'!")
        return f"ABS('CF'!{c}8-'CF'!{p}8-'CF'!{c}16)+ABS('CF'!{c}15-({finance_flow}))"
    check(128, "현금흐름 연결", cash_difference)
    check(129, "최소 보유현금 유지", lambda i,t,c,p:
          f"MAX(0,'{terms}'!$G$12/Unit-'CF'!{c}8)" if i>last else None)
    check(130, "한도대출 한도 준수", lambda i,t,c,p:
          f"MAX(0,'재원조달'!{c}18+'재원조달'!{c}21-'{terms}'!$G$13/Unit)" if i>=last else None)
    def schedule_difference(i, t, c, p):
        if i <= last:
            return None
        paid = "+".join(f"'재원조달'!{c}{r}" for r in principal) or "0"
        return f"{total(c,deviations)}+ABS(({paid})-{c}115/Unit)"
    check(131, "약정 상환 스케줄 대사", schedule_difference)
    check(132, "약정 기초 잔액 원천 대사", lambda i,t,c,p:
          f"ABS('재원조달'!{c}12-{c}114/Unit)" if i==last else None)
    check(133, "유형자산 증감 및 잔액", lambda i,t,c,p:
          f"ABS('투자비'!{c}14-'투자비'!{p}14-'투자비'!{c}8+'투자비'!{c}9)"
          f"+MAX(0,-'투자비'!{c}14,-'투자비'!{c}9)" if i>last else None)
    check(134, "손익 상각비와 투자비 대사", lambda i,t,c,p:
          f"ABS('운영비용'!{c}12-'투자비'!{c}9)")

    def funding_difference(i, t, c, p):
        if i < last:
            return None
        if i == last:
            return (f"IF(COUNT('재원조달'!{c}9,'재원조달'!{c}18)=2,"
                    f"ABS('재원조달'!{c}9)+ABS('재원조달'!{c}18),1/Unit)")
        balances = ",".join(f"-'재원조달'!{c}{r}" for r in (8, 9, 12, 13, 16, 17, 18, 19, 21, 24))
        return f"MAX(0,{balances},'재원조달'!{c}17-'재원조달'!{p}18)"
    check(135, "차입 잔액·이자 비음수 및 추가 인출분 상환 상한", funding_difference)

    for sheet, title in ((funding, "차입금과 상환"), (invest, "투자와 감가상각"),
                         (cash, "현금흐름"), ("검증", "월별 대사 내역")):
        m.section(sheet, 99 if sheet == "검증" else 7, title)
    m.section(cash, 19, "운전자본")
    if facilities:
        m.section(funding, 40, "약정별 차입금")
    m.total(funding, 8)
    m.total(cash, 8)
    m.total(invest, 14)
    m.format_row(funding, 10, '0.00"x"')
    if m.bs:
        m.section("BS", 7, "재무상태표")
        m.section(detail, 7, "재무상태표 계정 내역")
        for r in (8, 9, 10, 11):
            m.total("BS", r)
