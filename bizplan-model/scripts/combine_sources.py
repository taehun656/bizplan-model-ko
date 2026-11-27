"""Read the individual workbook contract and capture labelled rows in every case."""
from contextlib import contextmanager
from collections import Counter
from dataclasses import dataclass, field
from datetime import datetime
from hashlib import sha256
from pathlib import Path
from tempfile import TemporaryDirectory

from openpyxl import load_workbook
from openpyxl.utils.datetime import from_excel

from build_model import com_call, method, private_excel, timeline
from workbook_scan import load_meta
from excel_retry import cleanup_excel

IS = ("매출", "매출원가", "매출총이익", "판매관리비", "영업이익", "영업외수익", "영업외비용",
      "이자비용", "세전이익", "법인세", "당기순이익", "감가상각비", "EBITDA")
BS = ("유동자산", "비유동자산", "자산총계", "유동부채", "비유동부채", "부채총계", "자본", "자본총계", "부채와 자본총계")
CF = ("기초 현금", "영업현금흐름", "투자현금흐름", "재무현금흐름", "현금 증감 (실적은 BS 잔액 차이)",
      "기말 현금", "매출채권", "재고자산", "매입채무", "운전자본 증감의 현금 효과", "비현금 비용 가산 (상각 제외)",
      "기말 미지급법인세", "기말 선납·환급세액", "법인세 현금 납부 합계", "당월 인식 법인세비용")
DEBT = "기말 차입금"
OCF = "영업현금흐름"
SERVICE = ("약정 상환 및 만기 상환", "총 이자비용 (기초 잔액 기준)", "추가 인출분 잉여현금 상환")
DRAWN = ("추가 한도 사용 기말 잔액", "기존 한도 약정 사용 잔액")
NEW = "신규 조달 필요액 (당월 인출 가정)"
NEW_BAL = "신규 조달 필요액 누적 잔액"
SOURCE_DSCR = "DSCR (세후 영업현금 + 이자) / 원리금"
FUND = (DEBT, *DRAWN, NEW, NEW_BAL, *SERVICE)
INV = ("투자 지출", "감가상각비", "기존 자산 기말 순액", "기존 자산 상각", "신규 자산 기말 순액",
       "신규 자산 상각", "유형자산 기말 순액")
TAX = ("② 누적 과세표준", "당월 과세표준 증감", "③ 법인세 산출세액", "⑤ 법인지방소득세", "기말 이월결손금")
CHECK_KEYS = (("IS", "매출"), ("IS", "영업이익"), ("IS", "당기순이익"), ("CF", "기말 현금"), ("재원조달", DEBT))


def digest(path):
    with Path(path).open("rb") as f:
        h = sha256()
        for block in iter(lambda: f.read(1024 * 1024), b""):
            h.update(block)
    return h.hexdigest()


@contextmanager
def readonly_excel():
    with TemporaryDirectory(prefix="bizplan-read-", ignore_cleanup_errors=True) as temp:
        with private_excel(temp) as app:
            for name,value in (("Visible",False),("DisplayAlerts",False),("EnableEvents",False),
                               ("ScreenUpdating",False),("AutomationSecurity",3)):
                com_call(lambda: setattr(app,name,value))
            yield app


@contextmanager
def opened(app, path):
    book = com_call(lambda: app.Workbooks.Open(str(Path(path).resolve()), UpdateLinks=0, ReadOnly=True,
                        IgnoreReadOnlyRecommended=True, AddToMru=False))
    with cleanup_excel(lambda: com_call(lambda: book.Close(SaveChanges=False))):
        # Workbooks.Open can return before Excel accepts the first range read.
        com_call(lambda: book.Worksheets.Count)
        yield book


@contextmanager
def preserved_case(app, book):
    """Restore the original selector and its calculations before closing a read-only model."""
    switch = com_call(lambda: book.Names("SN_Case").RefersToRange)
    original = com_call(lambda: switch.Value2)
    def restore():
        com_call(lambda: setattr(switch, "Value2", original))
        method(app, "Calculate")
    with cleanup_excel(restore):
        yield switch


def excel_errors(wb):
    """Ask Excel for typed error cells; ordinary amounts can equal an error code."""
    import pywintypes
    hits=[]
    for ws in wb.Worksheets:
        for kind in (-4123,2):  # formulas, constants
            try:
                cells=ws.UsedRange.SpecialCells(kind,16)
            except pywintypes.com_error as exc:
                if exc.excepinfo and exc.excepinfo[5]==-2146827284:
                    continue  # xlErrNoCellsFound
                raise
            if cells is not None: hits.append(f"{ws.Name}!{cells.Address}")
    return hits


@dataclass
class Source:
    path: Path
    meta: dict
    cases: tuple
    rows: dict
    hash: str
    values: dict = field(default_factory=dict)
    checks: dict = field(default_factory=dict)
    read_at: str = ""

    @property
    def info(self):
        return self.meta["info"]

    @property
    def company(self):
        return self.info["company"]


def labelled(ws, title):
    # Summary labels use the indentation columns B:F. Prefer the shallowest
    # level when a revenue line has the same caption as a statement total.
    matches = [(c.column, c.row) for row in ws.iter_rows(min_col=2, max_col=6)
               for c in row if c.value == title]
    if matches:
        level = min(c for c, r in matches)
        rows = [r for c, r in matches if c == level]
        if len(rows) == 1:
            return rows[0]
    raise ValueError(f"{ws.title}에서 '{title}' 행을 하나로 식별할 수 없습니다.")


def operating_lines(ws, title):
    """Keep distinct source lines even when the spec uses the same display caption twice."""
    items=[]
    for line in ws.iter_rows(min_row=labelled(ws,title)+1,min_col=2,max_col=6):
        caption=next((cell.value for cell in line if cell.value is not None),None)
        if caption is None: break
        items.append((str(caption),line[0].row))
    counts=Counter(label for label,_ in items)
    taken=set(counts)
    occurrences=Counter()
    for label,row in items:
        key=label
        if counts[label]>1:
            while key in taken:
                occurrences[label]+=1
                key=f"{label} ({occurrences[label]})"
            taken.add(key)
        yield key,row


def inspect_source(path):
    path = Path(path).resolve()
    fingerprint = digest(path)
    meta = load_meta(path)
    info = meta["info"]
    required = ("company", "unit", "first_month", "last_actual", "plan_end", "built_at", "version",
                "first_month_col", "last_month_col")
    if not meta["present"] or any(k not in info for k in required) or info.get("kind") == "combined":
        raise ValueError(f"{path.name}: META_Info가 있는 법인별 사업계획 모델이 필요합니다.")
    wb = load_workbook(path, data_only=True)
    try:
        sens, addr = next(wb.defined_names["SN_CaseNos"].destinations)
        identifiers = wb[sens][addr]
        label_row = labelled(wb[sens], "케이스명")
        cases = tuple((c.value, wb[sens].cell(label_row, c.column).value) for c in identifiers[0])
        if (cases[:3] != ((1,"Base"),(2,"Up"),(3,"Down")) or
                any(n != i or not isinstance(label,str) or not label for i,(n,label) in enumerate(cases,1))):
            raise ValueError(f"{path.name}: Base, Up, Down부터 시작하는 케이스 번호·이름을 확인하세요.")
        # Resolve month positions from the actual worksheet axis; an inserted
        # worksheet row must not shift the values assigned to a financial line.
        start, end = int(info["first_month_col"]), int(info["last_month_col"])
        months = timeline(info["first_month"], info["plan_end"])
        if len(months) != end-start+1:
            raise ValueError("META_Info 월 열 수와 월 범위가 다릅니다.")
        rows = {}
        groups = {"IS": IS + TAX, "CF": CF, "재원조달": FUND + (SOURCE_DSCR,), "투자비": INV}
        if "BS" in wb.sheetnames:
            groups["BS"] = BS
        for sheet, captions in groups.items():
            for caption in captions:
                rows[sheet,caption] = labelled(wb[sheet], caption)
        for sheet, title in (("운영수입","매출 라인별 금액"),("운영비용","비용 라인별 금액")):
            for label,row in operating_lines(wb[sheet],title):
                rows[sheet,label] = row
        for sheet in {s for s, _ in rows}:
            axis = next((row for row in wb[sheet].iter_rows(min_col=start,max_col=end)
                         if isinstance(row[0].value,datetime)), ())
            if [c.value.strftime("%Y-%m") if isinstance(c.value,datetime) else None for c in axis] != months:
                raise ValueError(f"{sheet}: 월 축이 META_Info와 다릅니다.")
    except (KeyError, ValueError) as exc:
        raise ValueError(f"{path.name}: {exc}") from exc
    finally:
        wb.close()
    return Source(path,meta,cases,rows,fingerprint)


def align_sources(sources, mode):
    if len(sources)<2 or len({s.company.casefold() for s in sources}) != len(sources):
        raise ValueError("서로 다른 회사의 원본 모델을 2개 이상 지정하세요.")
    fields = [("표시 단위", lambda s:s.info["unit"]), ("실적 마감 월",lambda s:s.info["last_actual"]),
              ("케이스 번호·이름",lambda s:s.cases)]
    if mode != "intersect":
        fields.append(("월 범위 (--align intersect로 교집합 선택)",lambda s:(s.info["first_month"],s.info["plan_end"])))
    errors = [label+" 불일치 | "+" / ".join(f"{s.company} {get(s)}" for s in sources)
              for label,get in fields if any(get(s)!=get(sources[0]) for s in sources[1:])]
    if errors:
        raise ValueError("\n".join(errors))
    months = timeline(max(s.info["first_month"] for s in sources), min(s.info["plan_end"] for s in sources))
    if not months or sources[0].info["last_actual"] not in months:
        raise ValueError("실적 마감 월을 포함하는 공통 월 범위가 없습니다.")
    return months


def capture(app, src, months):
    start = int(src.info["first_month_col"]) + timeline(src.info["first_month"],months[0]).index(months[0])
    with opened(app,src.path) as wb:
        unit = com_call(lambda: wb.Names("Unit").RefersToRange.Value2)
        cutoff = from_excel(com_call(lambda: wb.Names("LastActual").RefersToRange.Value2)).strftime("%Y-%m")
        if unit != src.info["unit"] or cutoff != src.info["last_actual"]:
            raise ValueError(f"{src.company}: Unit/LastActual 현재값과 META_Info가 다릅니다. 원본을 재빌드하세요.")
        with preserved_case(app, wb) as switch:
            for n,_ in src.cases:
                com_call(lambda: setattr(switch,"Value2",n))
                method(app,"CalculateFull")
                values = {}
                for sheet in dict.fromkeys(sh for sh,label in src.rows):
                    by_row = [(key,r) for key,r in src.rows.items() if key[0]==sheet]
                    ws = com_call(lambda: wb.Worksheets(sheet))
                    block = com_call(lambda: ws.Range(ws.Cells(1,start), ws.Cells(max(r for _,r in by_row),start+len(months)-1)).Value2)
                    for key,r in by_row:
                        series = tuple(None if v=="" else v for v in block[r-1])
                        if any(v is not None and (not isinstance(v,(int,float)) or
                               (isinstance(v,int) and (v & 0xffff0000)==0x800a0000)) for v in series):
                            raise ValueError(f"{src.company} 케이스 {n} {sheet} {key[1]}: 월별 금액 오류")
                        values[key]=series
                src.values[n]=values
                src.checks[n]={c["name"]:com_call(lambda: wb.Worksheets(c["sheet"]).Range(c["cell"]).Value2)
                               for c in src.meta["checks"]}
    src.read_at=datetime.now().isoformat(timespec="seconds")
    if digest(src.path)!=src.hash:
        raise ValueError(f"{src.path.name}: 읽는 동안 원본 파일이 변경되었습니다.")
