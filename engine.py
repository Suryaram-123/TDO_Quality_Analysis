from __future__ import annotations
from collections import defaultdict
from dataclasses import dataclass
from pathlib import Path
import re, unicodedata
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.views import Selection

RED='FFFF0000'; ORANGE='FFFFA500'; HEADER='FF123B7A'; WHITE='FFFFFFFF'; BLACK='FF111827'

@dataclass
class Result:
    output_path: str
    reference_sheet: str|None
    analysis_sheet: str
    red_rows: int
    orange_rows: int
    pair_groups: int
    row_count_before: int
    row_count_after: int


def raw(v):
    if v is None: return ''
    if isinstance(v,float) and v.is_integer(): return str(int(v))
    return str(v).strip()

def norm(v):
    s=unicodedata.normalize('NFKD',raw(v)).encode('ascii','ignore').decode().upper()
    s=re.sub(r'[^A-Z0-9]+',' ',s)
    return ' '.join(s.split())

def norm_street(v):
    repl={'STREET':'ST','ROAD':'RD','AVENUE':'AVE','BOULEVARD':'BLVD','DRIVE':'DR','HIGHWAY':'HWY','ROUTE':'RTE','NORTH':'N','SOUTH':'S','EAST':'E','WEST':'W','SUITE':'STE'}
    return ' '.join(repl.get(t,t) for t in norm(v).split())

def addr_parts(v):
    s=norm_street(v); m=re.match(r'^(\d+)\s+(.+)$',s)
    return (int(m.group(1)),m.group(2)) if m else (None,s)

def hmap(ws):
    d={}
    for c in range(1,ws.max_column+1):
        v=ws.cell(1,c).value
        if v is not None and str(v).strip().casefold() not in d: d[str(v).strip().casefold()]=c
    return d

def find_col(h,*names):
    for n in names:
        if n.casefold() in h: return h[n.casefold()]
    return None

def detect_sheets(wb):
    reference=None; analysis=None
    for ws in wb.worksheets:
        h=hmap(ws)
        if all(find_col(h,x) for x in ('TDLINX_CODE','USER_ID','PROFILE_NAME')): reference=ws
        if find_col(h,'TDLinx Store Code','TDLINX_CODE') and find_col(h,'Store Name') and find_col(h,'Street Address'): analysis=ws
    if analysis is None:
        if len(wb.worksheets)==1: analysis=wb.worksheets[0]
        else: raise ValueError('Could not identify analysis sheet. Required: TDLinx Store Code, Store Name, Street Address.')
    return reference,analysis

def process_workbook(input_file,output_file):
    wb=load_workbook(input_file)
    ref,csv=detect_sheets(wb)
    if 'Summary' in wb.sheetnames: del wb['Summary']
    hc=hmap(csv)
    tdc=find_col(hc,'TDLinx Store Code','TDLINX_CODE'); namec=find_col(hc,'Store Name'); addrc=find_col(hc,'Street Address')
    numc=find_col(hc,'Store Number'); placec=find_col(hc,'Place Name'); cityc=find_col(hc,'City'); statec=find_col(hc,'State'); zipc=find_col(hc,'Postal Code','ZIP'); modc=find_col(hc,'Modified Date','DATE_MODIFIED')
    required={'TDLinx Store Code':tdc,'Store Name':namec,'Street Address':addrc}
    missing=[k for k,v in required.items() if not v]
    if missing: raise ValueError('Missing columns: '+', '.join(missing))
    before=csv.max_row-1

    owners=defaultdict(list)
    if ref:
        hr=hmap(ref); rt=find_col(hr,'TDLINX_CODE','TDLinx Store Code'); ru=find_col(hr,'USER_ID'); rp=find_col(hr,'PROFILE_NAME')
        for r in range(2,ref.max_row+1):
            td=norm(ref.cell(r,rt).value)
            if td: owners[td].append((raw(ref.cell(r,ru).value),raw(ref.cell(r,rp).value)))

    # remove previous helper columns
    for title in ['MATCH_TYPE','MATCH_RULE','PARENT_RED_TDLINX']:
        hc=hmap(csv); c=hc.get(title.casefold())
        if c: csv.delete_cols(c)
    base=csv.max_column; helpers={'MATCH_TYPE':base+1,'MATCH_RULE':base+2,'PARENT_RED_TDLINX':base+3}
    for k,c in helpers.items(): csv.cell(1,c,k)

    redfill=PatternFill('solid',fgColor=RED); orangefill=PatternFill('solid',fgColor=ORANGE); headfill=PatternFill('solid',fgColor=HEADER)
    white=Font(color=WHITE); black=Font(color=BLACK)
    for row in csv.iter_rows(min_row=2):
        for cell in row: cell.fill=PatternFill(fill_type=None); cell.font=black

    data={}; red_rows=[]
    for r in range(2,csv.max_row+1):
        v={'td':norm(csv.cell(r,tdc).value),'name':norm(csv.cell(r,namec).value),'addr':norm_street(csv.cell(r,addrc).value),
           'num':norm(csv.cell(r,numc).value) if numc else '', 'place':norm(csv.cell(r,placec).value) if placec else '',
           'city':norm(csv.cell(r,cityc).value) if cityc else '', 'state':norm(csv.cell(r,statec).value) if statec else '',
           'zip':norm(csv.cell(r,zipc).value) if zipc else ''}
        v['house'],v['street']=addr_parts(csv.cell(r,addrc).value); data[r]=v
        if ref and v['td'] in owners:
            red_rows.append(r); csv.cell(r,helpers['MATCH_TYPE'],'RED'); csv.cell(r,helpers['MATCH_RULE'],'RED_TDLINX_MATCH')
            for cell in csv[r]: cell.fill=redfill; cell.font=white
    # single-sheet mode: every row can be a parent, but no red highlighting
    parents=red_rows if ref else list(range(2,csv.max_row+1)); redset=set(red_rows)

    idx_num_name=defaultdict(set); idx_addr_zip=defaultdict(set); idx_identity=defaultdict(set); idx_city_street=defaultdict(set)
    for r,v in data.items():
        if v['num'] and v['name']: idx_num_name[(v['num'],v['name'])].add(r)
        if v['addr'] and v['zip']: idx_addr_zip[(v['addr'],v['zip'])].add(r)
        if all(v[k] for k in ('name','addr','city','state')): idx_identity[(v['name'],v['addr'],v['city'],v['state'])].add(r)
        if v['city'] and v['street']: idx_city_street[(v['city'],v['street'])].add(r)

    pair_records=[]; orange=defaultdict(lambda:{'rules':set(),'parents':set()})
    for pr in parents:
        p=data[pr]; candidates=set()
        if p['num'] and p['name']: candidates|=idx_num_name[(p['num'],p['name'])]
        if p['addr'] and p['zip']: candidates|=idx_addr_zip[(p['addr'],p['zip'])]
        if all(p[k] for k in ('name','addr','city','state')): candidates|=idx_identity[(p['name'],p['addr'],p['city'],p['state'])]
        if p['city'] and p['street']: candidates|=idx_city_street[(p['city'],p['street'])]
        for r in sorted(candidates):
            if r==pr or (ref and r in redset): continue
            if not ref and r<pr: continue
            d=data[r]; rules=[]
            if p['num'] and p['name'] and p['num']==d['num'] and p['name']==d['name']: rules.append('R1_STORE_NUMBER_AND_NAME')
            if p['addr'] and p['zip'] and p['addr']==d['addr'] and p['zip']==d['zip']: rules.append('R2_ADDRESS_AND_POSTAL')
            if all(p[k] and p[k]==d[k] for k in ('name','addr','city','state')): rules.append('R3_STORE_IDENTITY')
            if p['td'] and d['td'] and p['td']!=d['td'] and all(p[k] and p[k]==d[k] for k in ('name','addr','zip')): rules.append('R4_DIFFERENT_TDLINX_SAME_STORE')
            if p['name'] and d['name'] and p['name']!=d['name'] and p['addr'] and p['zip'] and p['addr']==d['addr'] and p['zip']==d['zip']: rules.append('R5_DIFFERENT_NAME_SAME_LOCATION')
            variant=bool(p['name'] and d['name'] and p['name']!=d['name'] and (p['name'] in d['name'] or d['name'] in p['name']))
            if variant and p['city'] and p['city']==d['city'] and p['street'] and p['street']==d['street'] and p['house'] is not None and d['house'] is not None and abs(p['house']-d['house'])<=2:
                rules.append('R6_NAME_VARIANT_NEARBY_ADDRESS')
            if rules:
                rules=sorted(set(rules)); pair_records.append((pr,r,rules)); orange[r]['rules'].update(rules); orange[r]['parents'].add(p['td'] or f'ROW_{pr}')

    for r,info in orange.items():
        csv.cell(r,helpers['MATCH_TYPE'],'ORANGE'); csv.cell(r,helpers['MATCH_RULE'],','.join(sorted(info['rules']))); csv.cell(r,helpers['PARENT_RED_TDLINX'],','.join(sorted(info['parents'])))
        for cell in csv[r]: cell.fill=orangefill; cell.font=black

    s=wb.create_sheet('Summary')
    heads=['PAIR_ID','TYPE','USER_ID','PROFILE_NAME','PARENT_RED_TDLINX','RECORD_TDLINX','STORE_NAME','STORE_NUMBER','STREET_ADDRESS','PLACE_NAME','CITY','STATE','POSTAL_CODE','MODIFIED_DATE','MATCH_RULE','CSV_ROW']
    for c,h in enumerate(heads,1):
        cell=s.cell(1,c,h); cell.fill=headfill; cell.font=Font(color=WHITE,bold=True); cell.alignment=Alignment(horizontal='center',wrap_text=True)
    grouped=defaultdict(list)
    for pr,r,rules in pair_records: grouped[pr].append((r,rules))
    outrow=2; pair_id=0
    for pr in parents:
        matches=grouped.get(pr,[])
        if not matches: continue
        pair_id+=1; p=data[pr]; owner_rows=owners.get(p['td'],[('','')]) if ref else [('','')]
        for user,profile in owner_rows:
            def value(row,col): return raw(csv.cell(row,col).value) if col else ''
            parent=[pair_id,'RED' if ref else 'PARENT',user,profile,p['td'] or f'ROW_{pr}',value(pr,tdc),value(pr,namec),value(pr,numc),value(pr,addrc),value(pr,placec),value(pr,cityc),value(pr,statec),value(pr,zipc),csv.cell(pr,modc).value if modc else '','RED_TDLINX_MATCH' if ref else 'PARENT_RECORD',pr]
            for c,v in enumerate(parent,1): s.cell(outrow,c,v).fill=redfill if ref else headfill; s.cell(outrow,c).font=white
            outrow+=1
            seen=set()
            for r,rules in matches:
                if r in seen: continue
                seen.add(r)
                vals=[pair_id,'ORANGE',user,profile,p['td'] or f'ROW_{pr}',value(r,tdc),value(r,namec),value(r,numc),value(r,addrc),value(r,placec),value(r,cityc),value(r,statec),value(r,zipc),csv.cell(r,modc).value if modc else '',','.join(rules),r]
                for c,v in enumerate(vals,1): s.cell(outrow,c,v).fill=orangefill; s.cell(outrow,c).font=black
                outrow+=1

    for ws in (csv,s):
        ws.freeze_panes='A2'; ws.sheet_view.topLeftCell='A1'; ws.sheet_view.selection=[Selection(activeCell='A1',sqref='A1')]; ws.sheet_view.showGridLines=False
        for r in range(1,ws.max_row+1): ws.row_dimensions[r].hidden=False
        for c in range(1,ws.max_column+1): ws.column_dimensions[get_column_letter(c)].hidden=False
    csv.auto_filter.ref=f'A1:{get_column_letter(csv.max_column)}{csv.max_row}'
    for c in range(1,csv.max_column+1):
        cell=csv.cell(1,c); cell.fill=headfill; cell.font=Font(color=WHITE,bold=True); cell.alignment=Alignment(horizontal='center',wrap_text=True)
        csv.column_dimensions[get_column_letter(c)].width=38 if c in helpers.values() else 16
    s.auto_filter.ref=f'A1:P{s.max_row}'
    for c,w in enumerate([10,12,18,22,21,18,30,16,32,22,20,18,14,20,42,10],1): s.column_dimensions[get_column_letter(c)].width=w
    wb.active=wb.sheetnames.index(csv.title)
    Path(output_file).parent.mkdir(parents=True,exist_ok=True); wb.save(output_file)
    return Result(str(output_file),ref.title if ref else None,csv.title,len(red_rows),len(orange),pair_id,before,csv.max_row-1)
