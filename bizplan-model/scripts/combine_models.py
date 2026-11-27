"""Build a self-contained, formula-driven simple aggregation of individual models."""
import argparse
import csv
from collections import defaultdict
from datetime import datetime
import math
import os
from pathlib import Path
import re
import sys
import time
import traceback

sys.path.insert(0,str(Path(__file__).resolve().parent))
sys.dont_write_bytecode = True
from openpyxl.utils import get_column_letter as letter
from openpyxl.utils.datetime import to_excel

from build_model import (INFO, Model, build_log, busy, method, output_lock, ref, write_prepared_model)
from combine_sources import (BS, CF, CHECK_KEYS, DEBT, DRAWN, FUND, INV, IS, NEW_BAL, OCF,
                             SERVICE, SOURCE_DSCR, TAX, align_sources, capture, digest, excel_errors, inspect_source, readonly_excel)
from dashboard import MULTIPLE
from style_guide import setup_guide
from workbook_style import setup_toc

ELIM = "제거조정"
FUNDING = "재원조달"
TAX_NOTE = "법인세는 개별 신고 기준으로 각 회사가 계산한 세액을 합산합니다."
# Each entry is an incremental adjustment to its named line. Roll it through
# statement subtotals so the post-elimination P&L and balance sheet still foot.
COMPONENTS = {
    ("IS","매출총이익"): (("매출",1),("매출원가",-1)),
    ("IS","영업이익"): (("매출총이익",1),("판매관리비",-1)),
    ("IS","영업외비용"): (("이자비용",1),),
    ("IS","세전이익"): (("영업이익",1),("영업외수익",1),("영업외비용",-1)),
    ("IS","당기순이익"): (("세전이익",1),),
    ("IS","EBITDA"): (("영업이익",1),("감가상각비",1)),
    ("BS","유동자산"): (("기말 현금",1),), ("BS","유동부채"): (("유동차입금",1),),
    ("BS","비유동부채"): (("비유동차입금",1),), ("BS","자산총계"): (("유동자산",1),("비유동자산",1)),
    ("BS","부채총계"): (("유동부채",1),("비유동부채",1)), ("BS","자본총계"): (("자본",1),),
    ("BS","부채와 자본총계"): (("부채총계",1),("자본총계",1)),
    ("BS",DEBT): (("유동차입금",1),("비유동차입금",1)),
}
ADJUSTABLE = {"IS": ("매출","매출원가","판매관리비","영업외수익","영업외비용","이자비용","감가상각비"),
              "BS": ("유동자산","비유동자산","유동부채","비유동부채","자본","기말 현금","유동차입금","비유동차입금")}
HEADER = ("month","statement","line","amount","description")
HEADER_KO = ("월","구분","줄","금액","설명")


def read_adjustments(path,months,has_bs):
    if not path:
        return []
    with path.open(encoding="utf-8-sig",newline="") as f:
        reader=csv.reader(f)
        if tuple(next(reader,())) not in (HEADER,HEADER_KO):
            raise ValueError("제거 CSV 머리글: "+",".join(HEADER))
        rows=[]
        for n,row in enumerate(reader,2):
            if not row: continue
            if len(row)!=5: raise ValueError(f"제거 CSV {n}행: 5개 열이 필요합니다.")
            month,sh,label,amount,description=(v.strip() for v in row)
            sh={"손익":"IS","재무상태":"BS"}.get(sh,sh)
            if month not in months or label not in ADJUSTABLE.get(sh,()):
                raise ValueError(f"제거 CSV {n}행: 월 또는 조정 대상 줄을 확인하세요 ({month}, {sh}, {label}).")
            if sh=="BS" and not has_bs and label not in ("기말 현금","유동차입금","비유동차입금"):
                raise ValueError(f"제거 CSV {n}행: 일부 원본에 BS가 없어 {label}을 조정할 수 없습니다.")
            try: number=float(amount)
            except ValueError as exc: raise ValueError(f"제거 CSV {n}행: 금액 오류") from exc
            if not math.isfinite(number): raise ValueError(f"제거 CSV {n}행: 유한한 금액이 필요합니다.")
            rows.append((month,sh,label,number,description))
        return rows


def serial(month):
    return to_excel(datetime.strptime(month,"%Y-%m"))


def complete_sum(addresses,adjustment="0"):
    args=",".join(addresses)
    return f'IF(COUNT({args})={len(addresses)},SUM({args})-({adjustment}),"")'


class Combined(Model):
    combined=True

    def __init__(self,name,sources,months,adjustments):
        self.sources,self.months,self.adjustments=sources,months,adjustments
        self.spec={"company":name,"unit":sources[0].info["unit"]}
        self.last=months.index(sources[0].info["last_actual"])
        self.years=defaultdict(list)
        for i,month in enumerate(months): self.years[month[:4]].append(i)
        self.cases=sources[0].cases
        self.bs=all(("BS","자산총계") in src.rows for src in sources)
        self.sheets=["목차","Sensitivity","민감도및결과","IS"]+(["BS"] if self.bs else [])+["CF",FUNDING,ELIM]
        self.roles={s:INFO[s] for s in self.sheets if s in INFO}
        self.roles[ELIM]=("input",1,"월별 회사 간 거래 제거와 재무제표 조정")
        self.roles["Sensitivity"]=("assumption",1,"회사 저장값의 케이스 선택")
        self.stores=[]
        for i,src in enumerate(sources,1):
            stem=re.sub(r"[\[\]:*?/\\'\x00-\x1f]","_",src.company).strip()
            sheet=stem[:28]+"_저장"
            existing={s.casefold() for s in self.stores}
            suffix_number=i
            while sheet.casefold() in existing:
                suffix=f"_{suffix_number}_저장"
                sheet=stem[:31-len(suffix)]+suffix
                suffix_number+=1
            self.stores.append(sheet)
            self.roles[sheet]=("actual",0,src.company+" · 모든 케이스의 월별 저장값")
        self.sheets+=self.stores+["검증","스타일 가이드","모델정보"]
        self.roles.update({s:INFO[s] for s in self.sheets if s not in self.roles})
        self.details={}
        self.initialize_grid()
        self.numeric_slots=defaultdict(list)
        self.outlines=defaultdict(list)
        self.store_rows=[]
        self.lines={}
        self.adj_rows={}
        self.check_table=[]
        self.built_at=datetime.now().isoformat(timespec="seconds")

    def series(self,sh,r,label,values,mode="sum"):
        self.row(sh,r,label,values,annual="none")
        periods=list(self.years.values())
        for c,indices in [(8,periods[-1]),*((10+len(self.months)+j,v) for j,v in enumerate(periods))]:
            a,b=self.col(indices[0]),self.col(indices[-1])
            area=f"{a}{r}:{b}{r}"
            expr=(f'IF(ISNUMBER({b}{r}),{b}{r},"")' if mode=="last" else
                  f'IF(COUNT({area}),{mode.upper()}({area}),"")')
            self.cell(sh,r,c,"="+expr)

    def selected(self,i,key,col,case=None):
        first=self.store_rows[i][key]
        sheet=ref(self.stores[i])
        if case is not None:
            return f"{sheet}!{col}{first+case-1}"
        return f"INDEX({sheet}!{col}${first}:{col}${first+len(self.cases)-1},MATCH(SN_Case,SN_CaseNos,0))"

    def adjustment(self,key,col):
        sh,label=key
        if key==("CF","기말 현금") or key==(FUNDING,DEBT): sh="BS"
        row=self.adj_rows.get((sh,label))
        return f"{ref(ELIM)}!{col}{row}" if row else "0"

    def direct_total(self,key,col,case=None):
        return complete_sum([self.selected(i,key,col,case) for i in range(len(self.sources))],
                            self.adjustment(key,col))

    def add_metric(self,key,r,mode):
        sh,label=key
        detail=r+1
        self.lines[key]=(r,detail)
        self.total(sh,r)
        cols=[self.col(i) for i in range(len(self.months))]
        self.series(sh,r,label,["="+complete_sum([f"{col}{detail+j}" for j in range(len(self.sources))],
                                                  self.adjustment(key,col)) for col in cols],mode)
        for j,src in enumerate(self.sources):
            selected=[self.selected(j,key,col) for col in cols]
            self.series(sh,detail+j,src.company,[f'=IF(ISNUMBER({v}),{v},"")' for v in selected],mode)
            self.level(sh,detail+j,4)
        self.outlines[sh].append((detail,detail+len(self.sources)-1))
        return detail+len(self.sources)


def headings(m):
    monthly={"IS","CF","BS",FUNDING,"검증",ELIM,*m.stores}
    for s in m.sheets:
        m.cell(s,1,1,m.spec["company"]+" 합산 사업계획","text")
        m.cell(s,2,1,s,"text")
        if s not in monthly: continue
        m.cell(s,5,8,"Sum or Max","text")
        for i,month in enumerate(m.months):
            col=m.col(i)
            m.cell(s,3,10+i,f"=YEAR({col}5)")
            m.cell(s,4,10+i,f'=IF({col}5<=LastActual,"실적","추정")')
            m.cell(s,5,10+i,serial(month))
        for j,(year,idx) in enumerate(m.years.items(),10+len(m.months)):
            m.cell(s,4,j,m.year_kind(idx),"text")
            m.cell(s,5,j,int(year))
    setup_toc(m)
    toc=m.grid["목차"]
    guide=max(m.sections["목차"])
    for cell in list(toc):
        if cell[0]>guide: del toc[cell]
    notes=["Sensitivity에서 케이스 번호를 선택하면 합산 재무제표와 결과가 바뀝니다.",TAX_NOTE,
           "차입·현금 부족은 법인별 계산 결과를 더합니다. 원본을 변경하면 합산 명령을 다시 실행하세요.",
           "월별 금액 중 한 회사라도 자료가 없으면 합계는 빈칸입니다.",
           ("모든 원본의 BS를 합산했습니다." if m.bs else "BS 미작성 · 재무상태표가 없는 회사: "+
            ", ".join(s.company for s in m.sources if ("BS","자산총계") not in s.rows))]
    notes += [f"{s.company}  |  {s.path.name}  |  원본 생성 {s.info['built_at']}  |  읽은 시각 {s.read_at}" for s in m.sources]
    for r,note in enumerate(notes,guide+1): m.cell("목차",r,3,note,"text")
    setup_guide(m)
    s="Sensitivity"
    for row,label,value in [(4,"선택 케이스",1),(6,"원본을 읽은 시각",m.built_at),
                            (8,"표시 단위 (원)",m.spec["unit"]),(9,"실적 마감 월",serial(m.months[m.last]))]:
        m.cell(s,row,2,label,"text");m.cell(s,row,3,value,"text")
    m.cell(s,4,5,"=INDEX(SN_CaseNames,1,MATCH(SN_Case,SN_CaseNos,0))")
    m.cell(s,11,2,"원본을 수정한 뒤 합산 명령을 다시 실행하세요.","text")
    for n,label in m.cases:
        m.cell(s,13,4+n,n);m.cell(s,14,4+n,label,"text")
    for r,label in [(13,"케이스 번호"),(14,"케이스명")]: m.cell(s,r,2,label,"text")
    last=letter(4+len(m.cases))
    for name,addr in (("Unit","$C$8"),("LastActual","$C$9"),("Scenario","$C$4"),("SN_Case","$C$4"),
                      ("SN_CaseNos",f"$E$13:${last}$13"),("SN_CaseNames",f"$E$14:${last}$14"),
                      ("SN_LastRun","$C$6")): m.named(name,s,addr)
    m.paint[s].append("C4")


def stores(m):
    for src,sheet in zip(m.sources,m.stores):
        for r,text in enumerate([src.path.name,"원본 생성 "+str(src.info["built_at"]),"읽은 시각 "+src.read_at],6):
            m.cell(sheet,r,2,text,"text")
        rows={}
        r=10
        for group in dict.fromkeys(sh for sh,label in src.rows):
            m.section(sheet,r,group);r+=1
            for key in (key for key in src.rows if key[0]==group):
                rows[key]=r
                for n,label in m.cases:
                    m.cell(sheet,r,3,key[1],"text");m.cell(sheet,r,7,label,"text")
                    for j,v in enumerate(src.values[n][key],10): m.cell(sheet,r,j,v,"text")
                    m.monthly_rows[sheet].add(r)
                    r+=1
        m.store_rows.append(rows)
        m.inputs.append(["회사별 저장값",src.company,sheet,f"J10:{m.col(len(m.months)-1)}{r-1}","원본의 케이스별 월 금액"])


def adjustments(m):
    s=ELIM
    m.table(s,7,2,[HEADER_KO]);m.header(s,7)
    if not m.adjustments: m.cell(s,6,2,"내부거래 없음","text")
    for r,(month,sh,label,amount,note) in enumerate(m.adjustments,8):
        m.table(s,r,2,[(serial(month),sh,label,amount,note)])
    last=max(8,7+len(m.adjustments))
    m.register(s,f"B8:F{last}","제거 명세","제거조정","모든 케이스에 동일하게 적용하는 표시 단위 금액")
    m.named("COMB_Eliminations",s,f"$B$7:$F${last}")
    r=last+3
    for sh,labels in ADJUSTABLE.items():
        m.section(s,r,sh+" 제거액");r+=1
        for label in labels:
            m.adj_rows[sh,label]=r
            m.cell(s,r,2,label,"text")
            for i in range(len(m.months)):
                m.cell(s,r,10+i,f'=SUMIFS($E$8:$E${last},$B$8:$B${last},{m.col(i)}$5,$C$8:$C${last},"{sh}",$D$8:$D${last},$B{r})')
            m.monthly_rows[s].add(r)
            r+=1
    for (sh,label),parts in COMPONENTS.items():
        prev=m.adj_rows.get((sh,label))
        m.cell(s,r,2,label+" · 소계 반영","text")
        for i in range(len(m.months)):
            col=m.col(i)
            expr=f"{col}{prev}" if prev else "0"
            for part,sign in parts: expr+=("+" if sign==1 else "-")+col+str(m.adj_rows[sh,part])
            m.cell(s,r,10+i,"="+expr)
        m.adj_rows[sh,label]=r;m.monthly_rows[s].add(r);r+=1


def statements(m):
    for sh,labels in [("IS",IS),("BS",BS),("CF",CF),(FUNDING,FUND)]:
        if sh not in m.sheets: continue
        r=8
        for label in labels:
            balance=(sh=="BS" or label in (DEBT,NEW_BAL,*DRAWN,"기초 현금","기말 현금","매출채권",
                                           "재고자산","매입채무","기말 미지급법인세","기말 선납·환급세액"))
            if (sh,label)==("IS","법인세"): m.cell(sh,r,7,"개별 신고 세액 합계","text")
            r=m.add_metric((sh,label),r,"last" if balance else "sum")
        if sh!=FUNDING: continue
        for title,parts in [("한도 사용액",DRAWN),("원리금",SERVICE)]:
            m.lines[sh,title]=(r,r+1)
            for j in range(len(m.sources)+1):
                vals=[]
                for i in range(len(m.months)):
                    col=m.col(i)
                    vals.append("="+complete_sum([col+str(m.lines[sh,k][0]+j) for k in parts]))
                m.series(sh,r+j,title if not j else m.sources[j-1].company,vals,"last" if title=="한도 사용액" else "sum")
            m.total(sh,r);m.outlines[sh].append((r+1,r+len(m.sources)));r+=1+len(m.sources)
        m.lines[sh,"DSCR"]=(r,r+1)
        for j in range(len(m.sources)+1):
            cash=m.lines["CF",OCF][0]+j;service=m.lines[sh,"원리금"][0]+j
            m.cell(sh,r+j,2,"DSCR · 영업현금흐름 / 원리금" if j==0 else m.sources[j-1].company+" · 원본 DSCR","text")
            for c in [8,*range(10,10+len(m.months)+len(m.years))]:
                col=letter(c)
                flow=f"CF!{col}{cash}"
                if j and 10<=c<10+len(m.months):
                    source=m.selected(j-1,(sh,SOURCE_DSCR),col)
                    formula=f'=IF(ISNUMBER({source}),{source},"")'
                else:
                    interest=m.lines[sh,SERVICE[1]][0]+j
                    numerator=f"({flow}+{col}{interest})" if j else flow
                    formula=f'=IF(AND(COUNT({flow},{col}{service})=2,{col}{service}>0),{numerator}/{col}{service},"")'
                m.cell(sh,r+j,c,formula)
            m.format_row(sh,r+j,MULTIPLE)
        m.total(sh,r);m.outlines[sh].append((r+1,r+len(m.sources)))


def checks(m):
    s="검증"
    count=len(CHECK_KEYS)+int(m.bs)+len(m.sources)*len(m.cases)
    first=10+count
    def add(label,expression):
        r=8+len(m.check_table)
        m.cell(s,r,2,label,"text");m.cell(s,r,7,expression)
        m.check_table.append((label,s,"G"+str(r),"OK"));m.checks.append((label,expression))
    for sh,label in CHECK_KEYS:
        r=first;first+=1
        values=[]
        for i in range(len(m.months)):
            col=m.col(i)
            actual=f"{ref(sh)}!{col}{m.lines[sh,label][0]}"
            expected=m.direct_total((sh,label),col)
            values.append(f'=IF(COUNT({actual},{expected})=2,ABS({actual}-({expected})),IF(COUNT({actual},{expected})=0,0,1))')
        m.series(s,r,label+" 월별 차이",values,"max")
        add(label+" 회사 합계 − 제거",f'=IF(MAX(J{r}:{m.col(len(m.months)-1)}{r})<=0.5/Unit,"OK","NO")')
    if m.bs:
        r=first;first+=1
        a,b=[m.lines["BS",k][0] for k in ("자산총계","부채와 자본총계")]
        m.series(s,r,"대차 차이",[f'=ABS(N(BS!{m.col(i)}{a})-N(BS!{m.col(i)}{b}))' for i in range(len(m.months))],"max")
        add("합산 대차 일치",f'=IF(MAX(J{r}:{m.col(len(m.months)-1)}{r})<=0.5/Unit,"OK","NO")')
    for src in m.sources:
        for n,case in src.cases:
            start=first
            for item in src.meta["checks"]:
                m.table(s,first,2,[[src.company,case,item["name"],item["expected"],src.checks[n][item["name"]]]])
                m.cell(s,first,7,f'=IF(E{first}=F{first},"OK","NO")');first+=1
            if first==start: raise ValueError(f"{src.company}: META_Checks에 검증 항목이 없습니다.")
            add(src.company+" "+case+" 원본 검증",f'=IF(COUNTIF(G{start}:G{first-1},"OK")=ROWS(G{start}:G{first-1}),"OK","NO")')


def metadata(m):
    info=[("kind","combined"),("company",m.spec["company"]),("unit",m.spec["unit"]),
          ("first_month",m.months[0]),("last_actual",m.months[m.last]),("plan_end",m.months[-1]),
          ("first_month_col",10),("last_month_col",9+len(m.months)),("version","1.1"),("built_at",m.built_at)]
    info += [(f"source_{i}",s.path.name) for i,s in enumerate(m.sources,1)]
    tables=[
        ("META_Info",["key","value"],info),
        ("META_Sheets",["sheet","role","stage","description"],[(s,*m.roles[s]) for s in m.sheets]),
        ("META_Inputs",["group","label","sheet","cell","description"],m.inputs),
        ("META_Checks",["name","sheet","cell","expected"],m.check_table),
        ("META_Sources",["company","filename","built_at","read_at","sha256","store"],
            [(src.company,src.path.name,src.info["built_at"],src.read_at,src.hash,s) for src,s in zip(m.sources,m.stores)]),
        ("META_StoredLines",["store","source_sheet","label","first_case_row"],
            [(sheet,*key,r) for sheet,rows in zip(m.stores,m.store_rows) for key,r in rows.items()]),
        ("META_CombinedLines",["sheet","label","total_row"],[(*key,rows[0]) for key,rows in m.lines.items()])]
    r=5
    for name,header,rows in tables:
        m.table("모델정보",r,1,[header,*rows]);m.named(name,"모델정보",f"$A${r}:${letter(len(header))}${r+len(rows)}")
        r+=len(rows)+3


def calculate_cases(app,wb,m):
    switch=wb.Names("SN_Case").RefersToRange
    for n,label in m.cases:
        switch.Value2=n;method(app,"CalculateFull")
        errors=excel_errors(wb)
        if errors: raise ValueError(f"{label}: 계산 오류 셀 {', '.join(errors[:8])}")
        for name,sh,addr,expected in m.check_table:
            if wb.Worksheets(sh).Range(addr).Value2!=expected:
                raise ValueError(f"합산 검증 실패: {label} / {name}")
    switch.Value2=1


def run(args,log):
    from combine_dashboard import build_results, style_combined
    from results_charts import build_charts
    sources=[inspect_source(path) for path in args.models]
    months=align_sources(sources,args.align)
    rows=read_adjustments(args.eliminations,months,all(("BS","자산총계") in s.rows for s in sources))
    with readonly_excel() as app:
        for s in sources:
            capture(app,s,months)
            print(f"- {s.company} | {s.path.name} | {len(s.cases)}개 케이스 | SHA-256 {s.hash}",file=log)
    m=Combined(args.name,sources,months,rows)
    headings(m);stores(m);adjustments(m);statements(m);checks(m);build_results(m);metadata(m)
    # Move only company detail captions, leaving elimination criteria labels in B.
    for sh,bands in m.outlines.items():
        for first,last in bands:
            for r in range(first,last+1):
                m.grid[sh][r,4]=m.grid[sh].pop((r,2));m.level(sh,r,4)
    for s in sources:
        if digest(s.path)!=s.hash: raise ValueError(f"{s.path.name}: 합산 도중 원본이 변경되었습니다.")
    write_prepared_model(m,args.out,calculate_cases,build_charts,style_combined)
    print(f"- 월 {months[0]} ~ {months[-1]}, 제거 {len(rows)}건, BS {'포함' if m.bs else '생략'}",file=log)


def main():
    parser=argparse.ArgumentParser(description="법인별 사업계획 모델을 케이스별로 단순 합산합니다.")
    parser.add_argument("--models",nargs="+",type=Path,required=True)
    parser.add_argument("--name",required=True)
    parser.add_argument("--out",type=Path,required=True)
    parser.add_argument("--eliminations",type=Path)
    parser.add_argument("--align",choices=("strict","intersect"),default="strict")
    parser.add_argument("--log",type=Path)
    args=parser.parse_args()
    args.out=args.out.resolve()
    backup=args.out.with_suffix(args.out.suffix+".prev")
    log=(args.log or args.out.with_suffix(".build_log.md")).resolve()
    template=Path(__file__).resolve().parents[1]/"assets"/"template.xlsm"
    inputs={p.resolve() for p in [*args.models,template]+([args.eliminations] if args.eliminations else [])}
    try:
        if args.out.suffix.lower()!=".xlsm" or not args.name.strip() or len(args.models)<2:
            raise ValueError("모델 두 개 이상, 합산 이름 및 .xlsm 산출 경로가 필요합니다.")
        if len({args.out,backup,log})!=3 or inputs.intersection({args.out,backup,log}):
            raise ValueError("원본·제거 CSV·템플릿·산출·백업·로그 경로는 서로 달라야 합니다.")
        args.out.parent.mkdir(parents=True,exist_ok=True);log.parent.mkdir(parents=True,exist_ok=True)
        with output_lock(args.out),build_log(log) as stream:
            print("# 합산 빌드\n",args.name,file=stream)
            if args.out.exists(): os.replace(args.out,backup)
            try:
                for attempt in range(3):
                    try: run(args,stream);break
                    except Exception as exc:
                        if not busy(exc) or attempt==2: raise
                        print(f"Excel 호출 재시도 {attempt+1}",file=stream);time.sleep(1)
            except Exception:
                traceback.print_exc(file=stream);raise
            print("\n결과: 성공",file=stream)
        print(f"합산 완료: {args.out}");return 0
    except Exception as exc:
        print(f"합산 실패: {exc}",file=sys.stderr);return 1


if __name__=="__main__":
    raise SystemExit(main())
