"""Build a formula-driven plan in a private Excel instance from a local model specification."""
from __future__ import annotations

import argparse
from collections import defaultdict
from contextlib import contextmanager
from datetime import datetime
import hashlib
from pathlib import Path
import os
import re
import sys
import shutil
import tempfile
import time
import traceback
import warnings

# Python isolated mode omits the script directory. Only import bundled project modules.
sys.path.insert(0, str(Path(__file__).resolve().parent))
sys.dont_write_bytecode = True
from openpyxl.utils import get_column_letter as letter
from openpyxl.utils.datetime import to_excel
from excel_retry import busy

MAIN = ['목차','Sensitivity','전제사항','민감도및결과','금융조건','INDEX','투자비','재원조달','운영수입','운영비용','IS','BS','CF','BS명세']
INFO = {
    '목차':('admin',4,'시트별 바로가기'), 'Sensitivity':('assumption',1,'케이스 선택, 비교값 갱신'),
    '전제사항':('assumption',1,'단위, 실적 마감, 시나리오별 가정'), '민감도및결과':('result',4,'연도별 전망과 케이스 비교'),
    '금융조건':('assumption',1,'약정과 원금 상환 조건'), 'INDEX':('assumption',1,'월 축과 실적 구간'),
    '투자비':('summary',3,'취득과 감가상각'), '재원조달':('summary',3,'기존 차입과 부족 자금 조달'),
    '운영수입':('summary',3,'매출 라인별 합계'), '운영비용':('summary',3,'비용 라인별 합계'),
    'IS':('statement',4,'월별 손익'), 'BS':('statement',4,'월말 재무상태'),
    'CF':('statement',4,'현금흐름과 기말 현금'), 'BS명세':('summary',3,'운전자본과 재무상태 명세'),
    '인건비':('driver',2,'직군별 인원과 급여'), '확장조정':('input',0,'동인별 추가 조정'),
    '검증':('check',4,'원천 대사와 계산 검증'), '기준값갱신':('admin',4,'케이스별 검증 기준값'),
    '스타일 가이드':('admin',4,'시트·셀·차트의 공통 서식 기준'),
    '모델정보':('admin',4,'검증·구조도용 모델 명세')}
def ref(sheet):
    return "'" + sheet.replace("'", "''") + "'"
def abscell(sheet, addr):
    return ref(sheet)+'!'+addr
def timeline(start, end):
    months=[]; y,m=map(int,start.split('-'))
    while f'{y:04}-{m:02}' <= end:
        months.append(f'{y:04}-{m:02}'); y,m=y+(m==12),m%12+1
    return months
def aggregate(sheet,row,first,last,mode):
    r=ref(sheet)+'!'; region=f'{r}{first}{row}:{last}{row}'
    if mode=='last': return f'={r}{last}{row}'
    return f'={mode.upper()}({region})'

class Model:
    def __init__(self,spec,data):
        self.spec,self.data=spec,data
        self.months=timeline(spec['first_month'],spec['plan_end'])
        self.last=self.months.index(spec['last_actual'])
        self.bs=bool(spec.get('statements',{}).get('bs',True) and data.get('bs_monthly'))
        self.details={}
        self.extensions=bool(spec.get('extensions'))
        self.lines=spec['revenue_lines']+spec['cost_lines']
        self.ids=[x['id'] for x in self.lines]
        self.years={}
        for i,month in enumerate(self.months):
            indices=self.years.setdefault(month[:4],[]);indices.append(i)
        self.sheets=[s for s in MAIN if self.bs or s not in ('BS','BS명세')]
        for i,line in enumerate(spec['revenue_lines'],1):
            safe=re.sub(r"[\[\]:*?/\'\x00-\x1f]",'_',line['label'])
            name=f'매출{i:02}_{safe}'[:31]
            self.details[line['id']]=name;self.sheets.append(name)
        self.activity_lines = [line for line in spec['revenue_lines'] if line['driver'] == 'activity']
        if self.activity_lines:
            self.sheets.extend(['영업계획','활동지표'])
        if self.extensions:self.sheets.append('확장조정')
        self.sheets+=['인건비','실적_IS']+(['실적_BS'] if self.bs else [])+['실적_차입','검증','기준값갱신','스타일 가이드','모델정보']
        self.initialize_grid()
        self.roles={s:INFO.get(s,('actual',0,'정규화한 원자료') if s.startswith('실적_') else ('driver',2,'매출 동인별 상세')) for s in self.sheets}
        if self.activity_lines:
            self.roles['영업계획'] = ('input',0,'사업부 계획 및 월 적용 조건')
            self.roles['활동지표'] = ('driver',2,'활동량 기반 지표와 비용 계산')
        self.active_driver_switches = {}
        self.cases=[{'label':n,'scenario':n} for n in ('Base','Up','Down')]+spec.get('cases',[])

    def initialize_grid(self):
        self.grid={s:{} for s in self.sheets}
        self.names={}; self.inputs=[]; self.checks=[]; self.paint=defaultdict(list)
        self.params, self.monthly_inputs = [], []
        self.sections, self.headers, self.totals = (defaultdict(set) for _ in range(3))
        self.levels, self.formats = (defaultdict(dict) for _ in range(2))
        self.monthly_rows = defaultdict(set)

    def col(self,i):return letter(10+i)
    def cell(self,s,r,c,value,kind='formula'):
        if value is None:return
        if not (1<=r<=1048576 and 1<=c<=16384):raise ValueError(f'Excel 범위 초과 {s} ({r},{c})')
        if isinstance(value,str) and (kind!='formula' or not value.startswith('=')):
            value="'"+value  # never interpret source text as an Excel expression
        self.grid[s][r,c]=value
    def row(self,s,row,label,values,kind='formula',annual='sum'):
        if len(values)!=len(self.months):raise ValueError(f'{s} {row}행 월 개수가 다릅니다.')
        self.cell(s,row,2,label,'text')
        for i,v in enumerate(values):self.cell(s,row,10+i,v,kind)
        self.monthly_rows[s].add(row)
        if annual=='none':return
        planned=list(self.years.values())[-1]
        first=0 if s=='검증' and annual=='max' else planned[0]
        self.cell(s,row,8,aggregate(s,row,self.col(first),self.col(planned[-1]),annual))
        for j,indices in enumerate(self.years.values()):
            self.cell(s,row,10+len(values)+j,aggregate(s,row,self.col(indices[0]),self.col(indices[-1]),annual))
    def register(self,s,addr,label,group,description):
        self.inputs.append([group,label,s,addr,description]);self.paint[s].append(addr)
    def input(self,s,row,label,values,group,description,annual='sum'):
        self.monthly_inputs.append((s,row,label,values,group,description))
        self.row(s,row,label,values,'input',annual)
        self.register(s,f'J{row}:{self.col(len(values)-1)}{row}',label,group,description)
    def param(self,s,row,label,value,group,description):
        self.params.append((s,row,label,value,group,description))
        self.cell(s,row,2,label,'text');self.cell(s,row,7,value,'input')
        self.register(s,'G'+str(row),label,group,description)
    def factor(self,id):return self.active_factors[id]
    def addition(self,id):return self.active_additions[id]
    def driver_switch(self,line_id,key):
        if (line_id,key) in self.active_driver_switches:
            return self.active_driver_switches[line_id,key]
        if key == 'renewal_probability':
            return self.renewal_probability_refs[line_id]
        defaults = {'new_multiplier':'1','price_multiplier':'1','exclude_unconfirmed':'0'}
        neutral = '0' if key.startswith('exclude_item:') else defaults[key]
        return neutral

    def year_kind(self,indices):
        if indices[-1] <= self.last:
            return '실적'
        if indices[0] > self.last:
            return '계획'
        return '예상'

    def section(self,s,row,title=None):
        self.sections[s].add(row)
        if title is not None:self.cell(s,row,2,title,'text')
    def header(self,s,row):self.headers[s].add(row)
    def total(self,s,row):self.totals[s].add(row)
    def level(self,s,row,col):self.levels[s][row]=col
    def format_row(self,s,row,fmt):self.formats[s][row]=fmt
    def add_check(self,name,formula):self.checks.append((name,formula))
    def table(self,s,r,c,rows):
        for i,rr in enumerate(rows):
            for j,v in enumerate(rr):self.cell(s,r+i,c+j,v,'text')
    def named(self,n,s,addr):self.names[n]=(s,addr)

def setup(m):
    spec=m.spec;axis='INDEX';ass='전제사항';sens='Sensitivity'
    monthly=set(m.sheets)-{'목차',sens,ass,'민감도및결과','기준값갱신','스타일 가이드','모델정보','영업계획'}
    for s in m.sheets:
        m.cell(s,1,1,spec['company']+' 사업계획','text');m.cell(s,2,1,s,'text')
        if s not in monthly:continue
        m.cell(s,5,8,'Sum or Max','text')
        for i,month in enumerate(m.months):
            value=to_excel(datetime.strptime(month,'%Y-%m')) if s==axis else '='+abscell(axis,m.col(i)+'5')
            m.cell(s,5,10+i,value)
            m.cell(s,3,10+i,f'=YEAR({m.col(i)}5)')
            m.cell(s,4,10+i,f'=IF({m.col(i)}5<=LastActual,"실적","추정")')
        for i,(year,indices) in enumerate(m.years.items()):
            annual_col=10+len(m.months)+i
            m.cell(s,4,annual_col,m.year_kind(indices),'text')
            m.cell(s,5,annual_col,int(year))
    m.row(axis,6,'실적 구간',[f'=--({m.col(i)}$5<=LastActual)' for i in range(len(m.months))],annual='none')
    m.param(ass,7,'표시 단위 (원)',spec['unit'],'단위','금액을 표시할 원 단위')
    m.cell(ass,8,2,'실적 마감 월','text');m.cell(ass,8,7,to_excel(datetime.strptime(spec['last_actual'],'%Y-%m')))
    m.cell(ass,9,2,'시나리오 번호','text');m.cell(ass,9,7,'=Sensitivity!C15')
    for name,k in [('Unit','$G$7'),('LastActual','$G$8'),('Scenario','$G$9')]:m.named(name,ass,k)
    from dashboard import configure_sensitivity
    from workbook_style import setup_toc
    configure_sensitivity(m)
    setup_toc(m)
    from style_guide import setup_guide
    setup_guide(m)


def add_control_checks(m):
    counts=[]
    for sh,grid in m.grid.items():
        if sh in {'Sensitivity','기준값갱신','검증','민감도및결과','스타일 가이드','모델정보','목차'} or not grid:continue
        bounds=f'A1:{letter(max(c for r,c in grid))}{max(r for r,c in grid)}'
        counts.append(f'SUMPRODUCT(--ISERROR({ref(sh)}!{bounds}))')
    m.add_check('계산 영역 오류 값 없음','=IF(SUM('+','.join(counts)+')=0,"OK","NO")')
    m.add_check('선택 시나리오 기준값 대사',None)


def finalize_controls(m):
    bl,check,meta='기준값갱신','검증','모델정보'
    m.cell(bl,4,2,'기준값 갱신 시각','text');m.named('BL_LastRun',bl,'$C$4')
    m.table(bl,7,2,[['현재값 범위','저장할 기준값 범위','케이스 번호']])
    baseline_last=16+len(m.live_labels)
    sh,addr=m.names['SN_Live']
    for j,label in enumerate(('Base','Up','Down'),5):
        c=letter(j)
        m.table(bl,j+3,2,[[sh+'!'+addr.replace('$',''),f'{bl}!{c}17:{c}{baseline_last}',j-4]])
        m.cell(bl,16,j,label,'text')
        for i,name in enumerate(m.live_labels,17):
            m.cell(bl,i,2,name,'text');m.cell(bl,i,j,0)
    m.named('BL_Jobs',bl,'$B$8:$D$10')
    matches=[]
    for i in range(1,len(m.live_labels)+1):
        current=f'INDEX(SN_Live,{i},1)'
        stored=f'INDEX(기준값갱신!$E$17:$G${baseline_last},{i},MIN(3,SN_Case))'
        matches.append(f'IF(ISNUMBER({current}),IF(ISNUMBER({stored}),ABS({current}-{stored})<=1E-9,FALSE),EXACT({current},{stored}))')
    baseline_formula='=IF(SN_Case>3,"OK",IF(AND('+','.join(matches)+'),"OK","NO"))'
    m.checks[-1]=(m.checks[-1][0],baseline_formula)
    checks=[]
    first=max(30,baseline_last+3)
    for i,(name,formula) in enumerate(m.checks,8):
        m.cell(check,i,2,name,'text');m.cell(check,i,7,formula)
        checks.append([name,check,f'G{i}','OK'])
        m.table(bl,first+i-8,2,[[f'검증!G{i}','OK']])
    m.named('BL_Checks',bl,f'$B${first}:$C${first+len(checks)-1}')
    metadata=[('META_Info',['key','value'],[
        ['company',m.spec['company']],['unit',m.spec['unit']],['first_month',m.months[0]],
        ['last_actual',m.months[m.last]],['plan_end',m.months[-1]],['first_month_col',10],
        ['last_month_col',9+len(m.months)],['version','1.5'],['built_at',datetime.now().isoformat(timespec='seconds')]]),
        ('META_Sheets',['sheet','role','stage','description'],[[sh,*m.roles[sh]] for sh in m.sheets]),
        ('META_Inputs',['group','label','sheet','cell','description'],m.inputs),
        ('META_Checks',['name','sheet','cell','expected'],checks)]
    r=5
    for name,header,rows in metadata:
        m.table(meta,r,1,[header,*rows]);m.named(name,meta,f'$A${r}:${letter(len(header))}${r+len(rows)}')
        r+=len(rows)+3

def com_call(action):
    for attempt in range(40):
        try:return action()
        except Exception as exc:
            if not busy(exc) or attempt==39:raise
            time.sleep(.25)

def method(obj,name):
    # Excel type libraries expose some zero-argument methods as both property
    # and method. Invoke as a method to avoid calling the Boolean return value.
    dispatch=obj._oleobj_
    return com_call(lambda:dispatch.Invoke(dispatch.GetIDsOfNames(name),0,1,0))

@contextmanager
def private_excel(temp):
    """Create a distinct server and use its registered type library for method signatures."""
    from excel_retry import cleanup_excel, quit_excel
    import pythoncom
    import win32com
    import win32com.gen_py
    previous_path = win32com.__gen_path__
    previous_search = win32com.gen_py.__path__
    cache = str(Path(temp) / 'gen_py')
    win32com.__gen_path__ = cache
    win32com.gen_py.__path__ = [cache]
    raw = None
    def quit_owned_server():
        if raw is not None:
            quit_excel(lambda: raw.Invoke(raw.GetIDsOfNames('Quit'), 0, 1, 0), retry=com_call)

    try:
        with cleanup_excel(quit_owned_server):
            from win32com.client import Dispatch, gencache
            raw = com_call(lambda: pythoncom.CoCreateInstanceEx(
                'Excel.Application', None, pythoncom.CLSCTX_LOCAL_SERVER, None,
                (pythoncom.IID_IDispatch,))[0])
            info = com_call(raw.GetTypeInfo)
            typelib, _ = info.GetContainingTypeLib()
            guid, lcid, _, major, minor, _ = typelib.GetLibAttr()
            gencache.EnsureModule(guid, lcid, major, minor, bForDemand=False)
            app = Dispatch(raw, resultCLSID=info.GetTypeAttr()[0])
            if not hasattr(type(app), 'CLSID'):
                raise RuntimeError('Excel 형식 라이브러리 바인딩을 생성하지 못했습니다.')
            yield app
    finally:
        win32com.__gen_path__ = previous_path
        win32com.gen_py.__path__ = previous_search

def write_sheet(ws,m, extra_style=None):
    name=ws.Name;grid=m.grid[name]
    lastrow=max(r for r,c in grid);lastcol=max(c for r,c in grid)
    block=ws.Range(f'A1:{letter(lastcol)}{lastrow}')
    values=tuple(tuple(grid.get((r,c)) for c in range(1,lastcol+1)) for r in range(1,lastrow+1))
    block.Formula=values
    from workbook_style import style_sheet
    from assumption_layout import style_assumptions
    from dashboard import style_dashboard
    style_sheet(ws,m)
    if name=='스타일 가이드':
        from style_guide import style_guide
        style_guide(ws,m)
    if extra_style is not None:
        extra_style(ws,m)
        return
    if getattr(m, 'activity_lines', None):
        from activity_model import style_activity
        style_activity(ws,m)
    if name=='전제사항':style_assumptions(ws,m)
    if name in ('Sensitivity','민감도및결과'):style_dashboard(ws,m)
    from input_validation import style_input_validation
    style_input_validation(ws,m)


def finalize_hierarchy(m):
    for sheet in m.sheets:
        if sheet in ('금융조건','전제사항','목차','민감도및결과','Sensitivity','기준값갱신','모델정보'):continue
        grid=m.grid[sheet]
        rows=set(m.monthly_rows[sheet])|set(m.levels[sheet])
        if sheet=='검증':rows.update(range(8,8+len(m.checks)))
        for row in rows:
            col=2 if row in m.sections[sheet] else m.levels[sheet].get(row,3)
            if col!=2 and (row,2) in grid and (row,col) not in grid:
                grid[row,col]=grid.pop((row,2))


def resolve_tax_references(m):
    """Bind assembly-time tax references without expanding the workbook contract."""
    addresses = {token: abscell(*m.names.pop(token)) for token in m.tax_refs.values()}
    token_pattern = re.compile(r"\bKR_Tax_[A-Za-z0-9_]+\b")
    def replace(formula):
        return token_pattern.sub(lambda match: addresses[match.group()], formula)
    for grid in m.grid.values():
        for cell, value in grid.items():
            if isinstance(value, str) and value.startswith("=") and "KR_Tax_" in value:
                grid[cell] = replace(value)
    m.checks[:] = [(label, replace(formula) if formula else formula) for label, formula in m.checks]
    m.tax_refs = {key: addresses[token] for key, token in m.tax_refs.items()}


def prepare_model(m):
    from finance_blocks import build_finance
    from driver_blocks import build_drivers
    from assumption_layout import finalize_assumptions
    from dashboard import finish_dashboard
    setup(m)
    build_drivers(m)
    build_finance(m)
    finalize_assumptions(m)
    from input_validation import build_input_checks
    build_input_checks(m)
    add_control_checks(m)
    finish_dashboard(m)
    finalize_controls(m)
    finalize_hierarchy(m)
    resolve_tax_references(m)


def populate_individual_cases(excel,wb,m):
    def named(n):return wb.Names(n).RefersToRange
    switch=named('SN_Case');live=named('SN_Live');store=named('SN_Store')
    for n in range(1,len(m.cases)+1):
        switch.Value2=n;method(excel,'Calculate')
        store.Columns(n).Value2=live.Value2
        if n<=3:
            destination=named('BL_Jobs').Cells(n,2).Value2
            sheet,address=destination.rsplit('!',1)
            wb.Worksheets(sheet).Range(address).Value2=live.Value2
    switch.Value2=1
    stamp=datetime.now().strftime('%Y-%m-%d %H:%M')
    for n in ['SN_LastRun','BL_LastRun']:named(n).Value2=stamp

def to_excel_file(m,out):
    prepare_model(m)
    from results_charts import build_charts
    write_prepared_model(m,out,populate_individual_cases,build_charts)

def write_prepared_model(m,out,populate_cases,draw_charts,extra_style=None):
    template=Path(__file__).resolve().parent.parent/'assets'/'template.xlsm'
    if not template.exists():raise ValueError('assets/template.xlsm을 찾을 수 없습니다.')
    out=out.resolve()
    if out.suffix.lower()!='.xlsm' or out==template:raise ValueError('--out은 템플릿과 다른 .xlsm 파일이어야 합니다.')
    out.parent.mkdir(parents=True,exist_ok=True)
    with tempfile.TemporaryDirectory(prefix='bizplan-',ignore_cleanup_errors=True) as temp:
        staged=Path(temp)/'result'/out.name
        staged.parent.mkdir()
        local_template=Path(temp)/'template.xlsm'
        shutil.copyfile(template,local_template)
        with private_excel(temp) as excel:
            excel.Visible=False;excel.DisplayAlerts=False;excel.AutomationSecurity=3;excel.EnableEvents=False;excel.ScreenUpdating=False
            wb=com_call(lambda: excel.Workbooks.Open(str(local_template),UpdateLinks=0,ReadOnly=True,IgnoreReadOnlyRecommended=True))
            from design_palette import FONT
            default_font = wb.Styles.Item('Normal').Font
            default_font.ThemeFont = 0
            default_font.Name = FONT
            default_font = None
            excel.Calculation=-4135
            if [s.Name for s in wb.Worksheets]!=['Sensitivity']:raise ValueError('템플릿은 Sensitivity 시트 한 장이어야 합니다.')
            for s in reversed(m.sheets[2:]):wb.Worksheets.Add().Name=s
            wb.Worksheets('Sensitivity').Move(wb.Worksheets.Item(1))
            method(wb.Worksheets('Sensitivity'),'Activate')
            wb.Worksheets.Add().Name='목차'
            order=[wb.Worksheets.Item(i).Name for i in range(1,wb.Worksheets.Count+1)]
            if order!=m.sheets:raise ValueError('시트 순서가 설계와 다릅니다.')
            for n,(sh,addr) in m.names.items():wb.Names.Add(Name=n,RefersTo='='+abscell(sh,addr))
            for title in m.sheets:
                s=wb.Worksheets.Item(title)
                try:write_sheet(s,m,extra_style)
                except Exception as exc:raise ValueError(f'{s.Name} 시트 작성 오류: {exc}') from exc
            populate_cases(excel,wb,m)
            def named(n):return wb.Names(n).RefersToRange
            excel.Calculation=-4105;method(excel,'CalculateFull')
            for name,s,addr,expected in named('META_Checks').Value2[1:]:
                got=wb.Worksheets(s).Range(addr).Value2
                if got!=expected:raise ValueError(f'검증 실패: {name} ({s}!{addr} = {got}, 기대 {expected})')
            from workbook_style import style_typed_cells
            style_typed_cells(wb, m)
            results = wb.Worksheets('민감도및결과')
            method(results,'Activate')
            draw_charts(results,m)
            excel.Goto(results.Range('A1'), True)
            wb.SaveAs(str(staged),FileFormat=52)
            # Drop worksheet and range proxies before their private server exits.
            s = results = None
            com_call(lambda: wb.Close(SaveChanges=False))
            wb = excel = None
        from dashboard import finalize_button_styles
        finalize_button_styles(staged)
        # Excel lock files stay local. Publish atomically even on a mounted drive.
        publish=None
        try:
            with tempfile.NamedTemporaryFile(dir=out.parent,prefix='.publish-',suffix='.tmp',delete=False) as dest:
                publish=Path(dest.name)
                with staged.open('rb') as src:shutil.copyfileobj(src,dest)
            os.replace(publish,out)
        finally:
            if publish is not None:publish.unlink(missing_ok=True)

@contextmanager
def output_lock(out):
    """Serialize two writers targeting the same published file on this workstation."""
    import win32api
    import win32event
    key = hashlib.sha256(os.path.normcase(str(out)).encode('utf-8')).hexdigest()
    mutex = win32event.CreateMutex(None, False, 'Local'+chr(92)+'BizPlanBuild-'+key)
    try:
        result = win32event.WaitForSingleObject(mutex, win32event.INFINITE)
        if result not in (win32event.WAIT_OBJECT_0, win32event.WAIT_ABANDONED):
            raise OSError('빌드 결과 파일 잠금을 얻지 못했습니다.')
        try:
            yield
        finally:
            win32event.ReleaseMutex(mutex)
    finally:
        win32api.CloseHandle(mutex)


@contextmanager
def build_log(path):
    """Publish one complete UTF-8 log if different outputs share a log path."""
    with tempfile.NamedTemporaryFile(mode='w',encoding='utf-8',dir=path.parent,
                                     prefix='.build-log-',suffix='.tmp',delete=False) as stream:
        staged=Path(stream.name)
        try:
            yield stream
        finally:
            stream.close()
            try:
                os.replace(staged,path)
            finally:
                staged.unlink(missing_ok=True)


def main():
    from spec_io import load_inputs
    parser=argparse.ArgumentParser(description=__doc__)
    for flag in ('spec','data','out'):
        parser.add_argument('--'+flag,required=True,type=Path)
    parser.add_argument('--log',type=Path,help='UTF-8 build log (default: build_log.md beside --out)')
    args=parser.parse_args()
    out=args.out.resolve()
    backup=out.with_name(out.name+'.prev')
    log=(args.log or out.with_name('build_log.md')).resolve()
    template=Path(__file__).resolve().parents[1]/'assets/template.xlsm'
    inputs={args.spec.resolve(),template.resolve()}
    try:
        if out.suffix.lower()!='.xlsm' or out in inputs or log in inputs|{out,backup} or log.is_relative_to(args.data.resolve()):
            raise ValueError('입력·템플릿·출력·로그 경로를 서로 다른 파일로 지정하세요.')
        out.parent.mkdir(parents=True,exist_ok=True)
        log.parent.mkdir(parents=True,exist_ok=True)
        with output_lock(out), build_log(log) as stream:
            print('# 모델 빌드', f'설계서: {args.spec}', f'자료: {args.data}', f'산출: {out}', sep='\n\n', file=stream)
            errors=None
            with warnings.catch_warnings(record=True) as caught:
                warnings.simplefilter('always')
                try:
                    if out.exists():
                        os.replace(out,backup)
                        print(f'\n이전 산출물 보관: {backup}',file=stream)
                    spec,data=load_inputs(args.spec,args.data)
                    for attempt in range(3):
                        try:
                            to_excel_file(Model(spec,data),out)
                            break
                        except Exception as exc:
                            if not busy(exc) or attempt==2:raise
                            print(f'\nExcel 일시 호출 거절로 재시도 {attempt+1}: {exc}',file=stream)
                            time.sleep(1)
                except Exception:
                    errors=traceback.format_exc()
            for warning in caught:
                print(f'\n경고: {warning.message}',file=stream)
            if errors:
                print('\n결과: 실패',errors,sep='\n\n',file=stream)
                print(f'FAIL build; warnings={len(caught)}; see build log')
                return 1
            size=out.stat().st_size
            print(f'\n결과: 성공 ({size:,} bytes)',file=stream)
            print(f'WROTE model ({size} bytes); warnings={len(caught)}')
    except (OSError,ValueError) as exc:
        print(f'FAIL build setup ({type(exc).__name__}); check output and log paths')
        return 1
    return 0

if __name__=='__main__':raise SystemExit(main())
