"""Management results, company contributions and native charts for a combined workbook."""
import re
from openpyxl.utils import get_column_letter as letter

from build_model import ref
from combine_sources import DEBT, NEW_BAL, OCF, SERVICE
from dashboard import (LIVE, MONTH, MULTIPLE, PERCENT, RESULT, UNIT_LABEL, build_cards,
                       set_result_print_area, style_cards, _table_style)
from design_palette import BLUE, ORANGE, WHITE, YELLOW, rgb
from results_sources import _scenario_table, _table
from workbook_style import set_number_format, set_numeric_list_format

METRICS = (("매출","revenue","IS","매출"),("매출원가","cogs","IS","매출원가"),
           ("매출총이익","gross","IS","매출총이익"),("판매관리비","sga","IS","판매관리비"),
           ("영업이익","op","IS","영업이익"),("EBITDA","ebitda","IS","EBITDA"),
           ("감가상각비","dep","IS","감가상각비"),("세전이익","pretax","IS","세전이익"),
           ("법인세","tax","IS","법인세"),("당기순이익","net","IS","당기순이익"),
           ("영업이익률","op_margin",None,None),("EBITDA율","ebitda_margin",None,None),
           ("기말 현금","cash","CF","기말 현금"),("기말 차입금","debt","재원조달",DEBT),
           ("한도 사용액","drawn","재원조달","한도 사용액"),("신규 조달 필요액 최고","new_peak","재원조달",NEW_BAL),
           ("DSCR (계획 구간)","dscr","재원조달","DSCR"))
KEYS={key:(sheet,caption) for _,key,sheet,caption in METRICS}


def ratio(cash,service):
    return f'IF(AND(COUNT({cash},{service})=2,{service}>0),{cash}/{service},"")'


def annual_area(m,idx,sheet,row):
    return f"{ref(sheet)}!{m.col(idx[0])}{row}:{m.col(idx[-1])}{row}"


def case_comparison(m,annual):
    s="Sensitivity"
    last_year=list(m.years.values())[-1]
    first,last=last_year[0],last_year[-1]
    for n,label in m.cases:
        keys=[KEYS[k] for _,k in LIVE[:-1]]+[("CF",OCF)]+[("재원조달",key) for key in SERVICE]
        start=32+(n-1)*(len(keys)+2)
        m.section(s,start,label+" · 계획 연도 월별 근거")
        for j,month in enumerate(m.months[first:last+1],10): m.cell(s,start,j,month,"text")
        for j,key in enumerate(keys,1):
            r=start+j
            m.cell(s,r,2,key[1],"text")
            for i in last_year:
                m.cell(s,r,10+i-first,"="+m.direct_total(key,m.col(i),n))
        for j,(label,key) in enumerate(LIVE,1):
            r=start+j
            end=m.col(last-first)
            area=f"J{r}:{end}{r}"
            if key in ("cash","debt"):
                f=f'IF(ISNUMBER({end}{r}),{end}{r},"")'
            elif key=="new_peak": f=f'IF(COUNT({area}),MAX({area}),"")'
            elif key=="dscr":
                projected=[i for i in last_year if i>m.last]
                if projected:
                    a=m.col(projected[0]-first)
                    service=f"SUM({a}{r+1}:{end}{r+len(SERVICE)})"
                    f=ratio(f"SUM({a}{r}:{end}{r})",service)
                else: f='""'
            else: f=f'IF(COUNT({area}),SUM({area}),"")'
            m.cell(s,17+j,4+n,"="+f)
            m.cell(s,17+j,2,label,"text")
            latest=letter(3+2*(len(m.years)-1))
            m.cell(s,17+j,3,f"='{RESULT}'!{latest}{annual[key]}")
    end=letter(4+len(m.cases))
    m.named("SN_Store",s,f"$E$18:${end}$25")
    m.named("SN_Live",s,"$C$18:$C$25")
    m.numeric_slots[s].append(f"C18:{end}25")
    m.cell(s,17,2,"계획 연도 지표","text");m.cell(s,17,3,"선택 케이스","text")
    for n,_ in m.cases: m.cell(s,17,4+n,f"={letter(4+n)}14")
    m.header(s,17)


def build_results(m):
    s=RESULT
    years=list(m.years.items());last=letter(3+2*(len(years)-1))
    annual={key:r for r,(_,key,_,_) in enumerate(METRICS,27)}
    m.section(s,24,"연도별 손익 및 자금");m.header(s,26);m.cell(s,26,2,"항목","text")
    for j,(year,idx) in enumerate(years):
        c=3+2*j;col=letter(c);yc=m.col(len(m.months)+j)
        m.cell(s,25,c,int(year));m.cell(s,26,c,m.year_kind(idx),"text");m.cell(s,26,c+1,"전년 대비","text")
        for caption,key,sh,label in METRICS:
            r=annual[key]
            m.cell(s,r,2,caption,"text")
            if key.endswith("_margin"):
                f=ratio(col+str(annual[key.split("_")[0]]),col+str(annual["revenue"]))
            elif key=="new_peak":
                area=annual_area(m,idx,sh,m.lines[sh,label][0]);f=f'IF(COUNT({area}),MAX({area}),"")'
            else:
                cell=f"{ref(sh)}!{yc}{m.lines[sh,label][0]}";f=f'IF(ISNUMBER({cell}),{cell},"")'
            m.cell(s,r,c,"="+f)
            if j and [m.months[i][5:] for i in idx]==[m.months[i][5:] for i in years[j-1][1]]:
                prior=letter(c-2)+str(r);current=col+str(r)
                delta=(f"{current}-{prior}" if key.endswith("_margin") else f"({current}-{prior})/ABS({prior})")
                m.cell(s,r,c+1,f'=IFERROR(IF(COUNT({current},{prior})=2,{delta},""),"")')
    build_cards(m,years,last,annual)
    case_comparison(m,annual)
    base=max(annual.values())+2
    m.section(s,base,years[-1][0]+"년 세 안 비교"+(" 및 추가 케이스" if len(m.cases)>3 else ""))
    m.header(s,base+1)
    for n,label in m.cases:
        m.cell(s,base+1,2+n,label,"text")
        for j,(metric,key) in enumerate(LIVE,1):
            m.cell(s,base+1+j,2,metric,"text")
            v=f"INDEX(SN_Store,{j},{n})";m.cell(s,base+1+j,2+n,f'=IF(ISNUMBER({v}),{v},"")')
    m.compare_box=f"B{base+1}:{letter(2+len(m.cases))}{base+1+len(LIVE)}"
    r=base+len(LIVE)+4
    m.section(s,r)
    m.cell(s,r,2,'="'+years[-1][0]+'년 회사별 기여 · 단위 "&'+UNIT_LABEL[1:])
    m.header(s,r+1);m.cell(s,r+1,2,"회사","text")
    keys=("revenue","ebitda","op","net","cash","debt")
    end_col=m.col(len(m.months)+len(years)-1)
    for j,key in enumerate(keys,3):
        sh,label=KEYS[key];total,detail=m.lines[sh,label]
        m.cell(s,r+1,j,label,"text")
        for i,src in enumerate(m.sources):
            row=r+2+i;cell=f"{ref(sh)}!{end_col}{detail+i}"
            m.cell(s,row,2,src.company,"text");m.cell(s,row,j,f'=IF(ISNUMBER({cell}),{cell},"")')
        adjust=r+2+len(m.sources)
        col=letter(j);summary=f"{last}{annual[key]}"
        m.cell(s,adjust,2,"제거 조정 (차감)","text")
        idx=years[-1][1]
        adjustment=m.adjustment((sh,label),m.col(idx[-1]))
        if key not in ("cash","debt") and adjustment!="0":
            adjustment=f"SUM({m.adjustment((sh,label),m.col(idx[0]))}:{adjustment.split('!')[-1]})"
        m.cell(s,adjust,j,f'=IF(ISNUMBER({summary}),{adjustment},"")')
        m.cell(s,adjust+1,2,"합계","text");m.cell(s,adjust+1,j,f'=IF(ISNUMBER({summary}),{summary},"")')
    m.total(s,adjust+1);m.contribution=(r,adjust+1)
    m.named("COMB_Contribution",s,f"$B${r}:$H${adjust+1}")
    header=adjust+5
    m.section(s,header-2,"차트 원천")
    revenue=m.lines["IS","매출"][0]
    columns=[(src.company,revenue+i+1) for i,src in enumerate(m.sources)]
    has_elim=any(sh=="IS" and line=="매출" for month,sh,line,amount,note in m.adjustments)
    # Include an adjustment series even for an initially blank register: changing
    # the amount or target of an existing entry must also update the chart.
    if has_elim or m.adjustments: columns.append(("제거 조정",None))
    title=f"B{header-1}";m.cell(s,header-1,2,'="연도별 매출 ("&'+UNIT_LABEL[1:]+'&")"')
    for c,(label,_) in enumerate(columns,3): m.cell(s,header,c,label,"text")
    total_col=3+len(columns);m.cell(s,header,total_col,"합계","text")
    for j,(year,idx) in enumerate(years):
        r=header+j+1;col=m.col(len(m.months)+j)
        m.cell(s,r,2,int(year));m.cell(s,r,total_col,f"='IS'!{col}{revenue}")
        for c,(_,row) in enumerate(columns,3):
            m.cell(s,r,c,f"='IS'!{col}{row}" if row else f"={letter(total_col)}{r}-SUM(C{r}:{letter(c-1)}{r})")
    m.chart_sources={"annual":_table(header,len(years),"B",[letter(i) for i in range(3,total_col+1)],title)}
    header+=len(years)+4
    # Use matching months for a year bridge; for a single year show the
    # change from its first to its last month instead.
    if len(years)>1:
        pyear,pidx=years[-2];cyear,cidx=years[-1]
        overlap={m.months[i][5:] for i in pidx}&{m.months[i][5:] for i in cidx}
        prev=[i for i in pidx if m.months[i][5:] in overlap];curr=[i for i in cidx if m.months[i][5:] in overlap]
    else: prev=curr=[]
    if not prev:
        prev,curr=[0],[len(m.months)-1];pyear,cyear=m.months[0],m.months[-1]
    m.cell(s,header-1,2,"매출 브리지 · 회사별 증감","text")
    m.table(s,header,2,[["항목","받침","매출","증가","감소","누계","증감","표시 위치","표시값"]])
    first=header+1;end=first+len(columns)+1
    def amount(idx,row):return "SUM("+annual_area(m,idx,"IS",row)+")"
    for r,year,idx in ((first,pyear,prev),(end,cyear,curr)):
        m.cell(s,r,2,year,"text");m.cell(s,r,4,"="+amount(idx,revenue));m.cell(s,r,7,f"=D{r}")
    for j,(name,row) in enumerate(columns,first+1):
        m.cell(s,j,2,name,"text")
        delta=(amount(curr,row)+"-"+amount(prev,row) if row else f"D{end}-G{j-1}")
        for c,expr in ((8,delta),(7,f"G{j-1}+H{j}"),(3,f"MIN(G{j-1},G{j})"),
                       (5,f"MAX(H{j},0)"),(6,f"MAX(-H{j},0)"),(10,f'TEXT(H{j},"+#,##0;(#,##0);0")')):
            m.cell(s,j,c,"="+expr)
    for r in range(first,end+1):m.cell(s,r,9,f"=SUM(C{r}:F{r})")
    bridge=_table(header,end-header,"B","CDEF",f"B{header-1}")
    bridge.update(label_height="I",label_text="J");m.chart_sources["bridge"]=bridge
    header=end+4
    m.chart_sources["scenario"]=_scenario_table(m,header)
    header+=len(m.cases)+4
    m.cell(s,header-1,2,"월별 매출·EBITDA·영업이익","text")
    m.table(s,header,2,[["월","매출","EBITDA","영업이익"]])
    for i in range(len(m.months)):
        r=header+i+1;col=m.col(i)
        m.cell(s,r,2,f"='IS'!{col}5")
        for c,k in enumerate(("매출","EBITDA","영업이익"),3):
            m.cell(s,r,c,f"='IS'!{col}{m.lines['IS',k][0]}")
    m.chart_sources["monthly"]=_table(header,len(m.months),"B","CDE",f"B{header-1}")
    m.annual=annual


def style_combined(ws,m):
    s=ws.Name
    if s in m.outlines:
        ws.Outline.SummaryRow=0
        for a,b in m.outlines[s]:
            ws.Rows(f"{a}:{b}").Group()
        ws.Outline.ShowLevels(RowLevels=1)
    if s=="Sensitivity":
        while ws.Shapes.Count: ws.Shapes.Item(1).Delete()
        ws.Columns("B:B").ColumnWidth=37;ws.Columns("C:C").ColumnWidth=23
        ws.Range("C9").NumberFormat=MONTH
        set_numeric_list_format(ws.Range("C4"))
        validation=ws.Range("C4").Validation
        validation.Delete();validation.Add(3,1,1,"=SN_CaseNos");validation.ShowError=True
        ws.Range("C4").Interior.Color=rgb(YELLOW);ws.Range("C4").Font.Color=rgb(BLUE)
        last=letter(4+len(m.cases))
        ws.Columns(f"E:{last}").ColumnWidth = 19
        ws.Range(f"E13:{last}14").HorizontalAlignment = -4108  # xlCenter
        ws.Range(f"E17:{last}17").HorizontalAlignment = -4108
        for row in (14, 17):
            ws.Range(f"E{row}:{last}{row}").WrapText = True
            ws.Rows(f"{row}:{row}").AutoFit()
        set_number_format(ws.Range(f"C25:{last}25"),MULTIPLE)
        rule=ws.Range(f"E13:{last}14").FormatConditions.Add(2,1,"=E$13=$C$4")
        rule.Interior.Color=rgb(ORANGE);rule.Font.Color=rgb(WHITE)
    elif s=="제거조정":
        for col,width in zip("BCDEF",(17,14,28,16,42)):ws.Columns(col+":"+col).ColumnWidth=width
        ws.Range(f"B8:B{max(8,7+len(m.adjustments))}").NumberFormat=MONTH
    elif s=="검증":
        for col,width in zip("BCDEFG",(32,15,43,12,12,15)):ws.Columns(col+":"+col).ColumnWidth=width
    elif s==RESULT:
        set_result_print_area(ws, len(m.years), max(m.annual.values()))
        ws.Columns("B:B").ColumnWidth=27;ws.Columns("C:M").ColumnWidth=15
        style_cards(ws,m)
        for j in range(len(m.years)):
            c=letter(3+2*j);d=letter(4+2*j)
            ws.Range(f"{c}25").NumberFormat="0"
            set_number_format(ws.Range(f"{d}27:{d}{max(m.annual.values())}"),PERCENT)
            for k in ("op_margin","ebitda_margin"):set_number_format(ws.Range(c+str(m.annual[k])),PERCENT)
            set_number_format(ws.Range(c+str(m.annual["dscr"])),MULTIPLE)
        _table_style(ws,m.compare_box)
        bottom=int(re.search(r"(\d+)$",m.compare_box).group(1))
        set_number_format(ws.Range(f"C{bottom}:{letter(2+len(m.cases))}{bottom}"),MULTIPLE)
        a,b=m.contribution;_table_style(ws,f"B{a+1}:H{b}")
        ws.Range(f"B{a}:H{a}").Merge()
        h=m.chart_sources["monthly"]["header"]
        ws.Range(f"B{h+1}:B{h+len(m.months)}").NumberFormat=MONTH
