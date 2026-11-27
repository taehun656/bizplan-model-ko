"""Independently reconcile all monthly combined statements and optional live sources."""
from collections import defaultdict
from datetime import date, datetime
from math import isclose
from pathlib import Path
import time
from zipfile import ZipFile

from openpyxl import load_workbook
from openpyxl.utils.datetime import from_excel

from build_model import busy, com_call, method, timeline
from combine_models import ADJUSTABLE
from combine_sources import DEBT, DRAWN, OCF, SERVICE, SOURCE_DSCR, align_sources, capture, excel_errors, inspect_source, opened, preserved_case, readonly_excel
from common import write_text
from excel_retry import BUSY_MESSAGE, ExcelShutdownBusy
from font_audit import font_warnings_for, inspect_fonts
from palette_audit import inspect_palette, warnings_for
from workbook_scan import records, scan_formulas, table


def equal(a,b):
    if a in (None,"") or b in (None,""):
        return a in (None,"") and b in (None,"")
    if isinstance(a,(float,int)) and isinstance(b,(float,int)):
        return isclose(a,b,rel_tol=1e-13,abs_tol=1e-9)
    return a==b


def adjustments_by_month(rows,months):
    direct=defaultdict(float)
    for month,sh,label,amount,note in rows:
        if all(x in ("",None) for x in (month,sh,label,amount,note)): continue
        if isinstance(month,(int,float)): month=from_excel(month)
        if isinstance(month,date): month=month.strftime("%Y-%m")
        if month not in months or label not in ADJUSTABLE.get(sh,()) or not isinstance(amount,(float,int)):
            raise ValueError(f"제거조정 명세 오류: {month} / {sh} / {label} / {amount}")
        direct[month,sh,label]+=amount
    result={}
    # Expand the accounting identities separately from the Excel formula generator.
    for month in months:
        inc={k:direct[month,"IS",k] for k in ADJUSTABLE["IS"]}
        inc["매출총이익"]=inc["매출"]-inc["매출원가"]
        inc["영업이익"]=inc["매출총이익"]-inc["판매관리비"]
        inc["영업외비용"]+=inc["이자비용"]
        inc["세전이익"]=inc["영업이익"]+inc["영업외수익"]-inc["영업외비용"]
        inc["당기순이익"]=inc["세전이익"]
        inc["EBITDA"]=inc["영업이익"]+inc["감가상각비"]
        bal={k:direct[month,"BS",k] for k in ADJUSTABLE["BS"]}
        bal["유동자산"]+=bal["기말 현금"];bal["유동부채"]+=bal["유동차입금"];bal["비유동부채"]+=bal["비유동차입금"]
        bal["자산총계"]=bal["유동자산"]+bal["비유동자산"]
        bal["부채총계"]=bal["유동부채"]+bal["비유동부채"];bal["자본총계"]=bal["자본"]
        bal["부채와 자본총계"]=bal["부채총계"]+bal["자본"]
        for sh,items in (("IS",inc),("BS",bal)):
            result.update({(month,sh,k):v for k,v in items.items()})
        result[month,"CF","기말 현금"]=bal["기말 현금"]
        result[month,"재원조달",DEBT]=bal["유동차입금"]+bal["비유동차입금"]
    return result


def read_snapshot(path,meta,months):
    w=load_workbook(path,data_only=True)
    try:
        named=lambda key:table(w,w.defined_names[key].attr_text)
        manifest=records(named("META_Sources"))
        lines=records(named("META_CombinedLines"))
        rows=records(named("META_StoredLines"))
        adjustments=adjustments_by_month(named("COMB_Eliminations")[1:],months)
        nums,labels=named("SN_CaseNos")[0],named("SN_CaseNames")[0]
        cases=tuple(zip(nums,labels))
        if cases[:3]!=((1,"Base"),(2,"Up"),(3,"Down")) or list(nums)!=list(range(1,len(nums)+1)):
            raise ValueError("합산 케이스 표의 번호 또는 이름 오류")
        values=defaultdict(dict)
        for row in rows:
            store=row["store"];key=(row["source_sheet"],row["label"])
            ws=w[store];first=int(row["first_case_row"])
            for n,label in cases:
                r=first+n-1
                if ws.cell(r,3).value!=key[1] or ws.cell(r,7).value!=label:
                    raise ValueError(f"{store} 저장 행 이름·케이스 불일치: {r}")
                values[store,n][key]=tuple(ws.cell(r,10+i).value for i in range(len(months)))
        return manifest,lines,adjustments,cases,values
    finally: w.close()


def verify_combined(path,meta,source_paths,out):
    for attempt in range(3):
        try:
            return verify_once(path,meta,source_paths,out,retry_busy=attempt<2)
        except Exception as exc:
            if not busy(exc): raise
            time.sleep(attempt + 1)


def verify_once(path,meta,source_paths,out,retry_busy):
    failures=[];notes=[];counts=defaultdict(int);maximum=defaultdict(float)
    def compare(a,b,tag,loc):
        counts[tag]+=1
        if isinstance(a,(int,float)) and isinstance(b,(int,float)):
            maximum[tag]=max(maximum[tag],abs(a-b))
        if not equal(a,b):
            counts["불일치"]+=1
            if counts["불일치"]<=20: failures.append(f"{tag} {loc}: {a!r} / 기대 {b!r}")
    if out and out.resolve() in {Path(p).resolve() for p in [path,*(source_paths or [])]}:
        print("검증 보고서 경로가 모델과 같습니다.");return 2
    try:
        info=meta["info"];months=timeline(info["first_month"],info["plan_end"])
        manifest,lines,adj,cases,values=read_snapshot(path,meta,months)
        with ZipFile(path) as z:
            if "xl/vbaProject.bin" not in z.namelist(): failures.append("매크로 프로젝트 누락")
            if any(n.startswith("xl/externalLinks/") for n in z.namelist()):failures.append("외부 통합문서 연결이 있습니다.")
        scan=scan_formulas(path)["forbidden"]
        failures.extend(f"{k} 수식: {', '.join(v[:5])}" for k,v in scan.items() if v)
        colors=warnings_for(inspect_palette(path));fonts=font_warnings_for(inspect_fonts(path))
        failures.extend(colors+fonts);notes.append(f"색 경고 {len(colors)}건 · 글꼴 경고 {len(fonts)}건")
        stores=[s["store"] for s in manifest]
        with readonly_excel() as app:
            if source_paths:
                sources=[inspect_source(p) for p in source_paths]
                overlap=align_sources(sources,"intersect")
                source_manifest={s["company"]:s for s in manifest}
                if set(source_manifest)!={s.company for s in sources}:
                    raise ValueError("합산 대상 회사와 --sources 회사 목록이 다릅니다.")
                if not all(t in overlap for t in months): raise ValueError("원본에서 합산의 월 범위를 찾을 수 없습니다.")
                for src in sources:
                    item=source_manifest[src.company]
                    if datetime.fromisoformat(str(src.info["built_at"]))>datetime.fromisoformat(str(info["built_at"])):
                        failures.append(f"낡은 합산: {src.company} 원본 생성 {src.info['built_at']}, 합산 생성 {info['built_at']}")
                    if src.hash!=item["sha256"]: notes.append(f"{src.company}: 원본 해시 변경, 현재 셀 값 재대조")
                    if src.cases!=cases or src.info["unit"]!=info["unit"] or src.info["last_actual"]!=info["last_actual"]:
                        raise ValueError(f"{src.company}: 케이스·표시 단위·마감 월이 합산 명세와 다릅니다.")
                    capture(app,src,months)
                    for n,label in cases:
                        stored=values[item["store"],n]
                        if stored.keys()!=src.values[n].keys(): raise ValueError(f"{src.company}: 저장 행 구성 변경")
                        for key,series in stored.items():
                            for i,(got,want) in enumerate(zip(series,src.values[n][key])):
                                compare(got,want,"원본 저장값",f"{src.company} {label} {key} {months[i]}")
                        for check in src.meta["checks"]:
                            compare(src.checks[n][check["name"]],check["expected"],"원본 검증",f"{src.company} {label} {check['name']}")
            else: notes.append("원본 재대조 생략 (--sources 미지정)")
            with opened(app,path) as wb, preserved_case(app, wb) as switch:
                for n,label in cases:
                    com_call(lambda: setattr(switch,"Value2",n));method(app,"CalculateFull")
                    for ws in wb.Worksheets:
                        circular=ws.CircularReference
                        if circular is not None: failures.append(f"{label}: {ws.Name}!{circular.Address} 순환 참조")
                    errors=excel_errors(wb)
                    if errors: failures.append(f"{label}: 오류 셀 "+", ".join(errors[:6]))
                    for check in meta["checks"]:
                        compare(wb.Worksheets(check["sheet"]).Range(check["cell"]).Value2,check["expected"],
                                "합산 검증",label+" "+check["name"])
                    live=wb.Names("SN_Live").RefersToRange.Value2
                    saved=wb.Names("SN_Store").RefersToRange.Value2
                    for i,(row,store) in enumerate(zip(live,saved)):
                        compare(row[0],store[n-1],"케이스 비교표",f"{label} 지표 {i+1}")
                    for sh in dict.fromkeys(line["sheet"] for line in lines):
                        ws=wb.Worksheets(sh)
                        selected=[line for line in lines if line["sheet"]==sh]
                        maxrow=max(int(line["total_row"]) for line in selected)+len(stores)
                        matrix=ws.Range(ws.Cells(1,10),ws.Cells(maxrow,9+len(months))).Value2
                        for i,month in enumerate(months):
                            def raw(store,key):return values[store,n][key][i]
                            def sum_parts(store,keys):
                                nums=[raw(store,key) for key in keys]
                                return sum(nums) if all(x is not None for x in nums) else None
                            for line in selected:
                                key=(sh,line["label"]);offset=int(line["total_row"])-1
                                if key[1] in ("원리금","한도 사용액","DSCR") and sh=="재원조달":
                                    parts=DRAWN if key[1]=="한도 사용액" else SERVICE
                                    amounts=[sum_parts(store,[(sh,k) for k in parts]) for store in stores]
                                else: amounts=[raw(store,key) for store in stores]
                                want=sum(amounts)-adj.get((month,*key),0) if None not in amounts else None
                                if key==("재원조달","DSCR"):
                                    flows=[raw(store,("CF",OCF)) for store in stores]
                                    want=sum(flows)/want if want and want>0 and None not in flows else None
                                    amounts=[raw(store,("재원조달",SOURCE_DSCR)) for store in stores]
                                compare(matrix[offset][i],want,"월별 합계",f"{label} {month} {key}")
                                for j,(store,val) in enumerate(zip(stores,amounts),1):
                                    compare(matrix[offset+j][i],val,"회사별 소계",f"{store} {label} {month} {key}")
    except Exception as exc:
        if retry_busy and busy(exc): raise
        failures.append(f"검증 실행 실패: {BUSY_MESSAGE if busy(exc) or isinstance(exc, ExcelShutdownBusy) else exc}")
    details=[f"- {k} {v:,}건, 최대 절대차 {maximum[k]:.12g}" for k,v in counts.items() if k!="불일치"]
    status="FAIL" if failures else "PASS"
    report="\n".join([f"# {Path(path).name} 합산 검증","",f"결과 {status}","",*details,"",*("- "+n for n in notes),
                      *("- "+f for f in failures),""])
    if out: write_text(out,report)
    print(report)
    return int(bool(failures))
