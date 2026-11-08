"""Copy independently recalculated values into the Excel-preserving package."""
import math, re, sys, zipfile
from pathlib import Path
from lxml import etree as E
from openpyxl import load_workbook
from openpyxl.utils.cell import range_boundaries
from dashboard_package import dashboard_chart_data

X='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
C='http://schemas.openxmlformats.org/drawingml/2006/chart'
R='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
def x(n): return f'{{{X}}}{n}'
def c(n): return f'{{{C}}}{n}'
def pack(tree): return E.tostring(tree,encoding='UTF-8',xml_declaration=True,standalone=True)

template,checked,final=map(Path,sys.argv[1:4])
final.parent.mkdir(parents=True,exist_ok=True)
calc=load_workbook(checked,data_only=True)
count=calc['Inputs']['B70'].value
assert isinstance(count,int) and count>0
risk_count=count if calc['Summary']['B25'].value=='YES' else 1
updates={}
with zipfile.ZipFile(template) as src:
    wb=E.fromstring(src.read('xl/workbook.xml')); rr=E.fromstring(src.read('xl/_rels/workbook.xml.rels'))
    paths={r.get('Id'):r.get('Target') for r in rr}
    for s in wb.find(x('sheets')):
        name=s.get('name'); path=paths[s.get(f'{{{R}}}id')]; path=path.lstrip('/') if path.startswith('/') else 'xl/'+path
        root=E.fromstring(src.read(path)); ws=calc[name]
        for el in root.iter(x('c')):
            if el.find(x('f')) is None: continue
            saved=ws[el.get('r')]
            assert saved.data_type!='e',(name,el.get('r'),saved.value)
            v=saved.value; old=el.find(x('v'))
            if old is not None: el.remove(old)
            el.attrib.pop('t',None)
            if isinstance(v,bool): el.set('t','b'); E.SubElement(el,x('v')).text='1' if v else '0'
            elif isinstance(v,(int,float)): assert math.isfinite(v); E.SubElement(el,x('v')).text=str(v)
            elif isinstance(v,str) or v is None: el.set('t','str'); E.SubElement(el,x('v')).text=v or ''
            else: raise ValueError((name,el.get('r'),v))
        updates[path]=pack(root)
    for path in [p for p in src.namelist() if re.fullmatch(r'xl/charts/chart\d+\.xml',p)]:
        root=E.fromstring(src.read(path))
        for ref in root.iter():
            if ref.tag not in {c('numRef'),c('strRef')}: continue
            fm=ref.findtext(c('f')) or ''
            cache_tag=c('numCache' if ref.tag==c('numRef') else 'strCache')
            cache=ref.find(cache_tag)
            match=re.fullmatch(r"'retirement-model\.xlsx'!RM(Plan|Trials)([A-Z]+)",fm)
            data=dashboard_chart_data(fm,calc)
            if data is not None:
                if cache is None: cache=E.SubElement(ref,cache_tag)
            elif cache is None:
                continue
            elif match:
                key,col=match.groups(); sheet,first=('Projection',5) if key=='Plan' else ('Monte_Carlo',25)
                data=[calc[sheet][f'{col}{r}'].value for r in range(first,first+(count if key=='Plan' else risk_count))]
            elif '!' in fm:
                sheet,a=fm.split('!',1); sheet=sheet.strip("'")
                if sheet not in calc.sheetnames: continue
                c1,r1,c2,r2=range_boundaries(a)
                data=[calc[sheet].cell(r,col).value for r in range(r1,r2+1) for col in range(c1,c2+1)]
            else: continue
            for pt in list(cache.findall(c('pt'))): cache.remove(pt)
            pt_count=cache.find(c('ptCount'))
            if pt_count is None: pt_count=E.SubElement(cache,c('ptCount'))
            pt_count.set('val',str(len(data)))
            for i,value in enumerate(data):
                if value is None or value=='': continue
                if ref.tag==c('numRef') and not isinstance(value,(int,float)): continue
                E.SubElement(E.SubElement(cache,c('pt'),idx=str(i)),c('v')).text=str(value)
        updates[path]=pack(root)
    with zipfile.ZipFile(final,'w',zipfile.ZIP_DEFLATED,compresslevel=5) as out:
        for p in src.namelist(): out.writestr(p,updates[p] if p in updates else src.read(p))

out=load_workbook(final,data_only=True)
assert all(c.data_type!='e' for ws in out for c in ws._cells.values())
assert out['Summary']['B29'].value==calc['Summary']['B29'].value
if 'Monte_Carlo' in out.sheetnames:
    assert out['Monte_Carlo']['B15'].value==calc['Monte_Carlo']['B15'].value
for name,first,last in [('Projection',5,130),('Monte_Carlo',25,150),('Sequence_Risk',15,140),('Cash_Flow',5,130)]:
    if name not in out.sheetnames: continue
    col='B' if name=='Cash_Flow' else 'C'
    if first+count<=last:
        assert all(row[0].value in ('',None) for row in out[name][f'{col}{first+count}:{col}{last}']),name
with zipfile.ZipFile(final) as z:
    for path in [p for p in z.namelist() if re.fullmatch(r'xl/charts/chart\d+\.xml',p)]:
        chart=E.fromstring(z.read(path))
        for ref in chart.iter(c('numRef'),c('strRef')):
            data=dashboard_chart_data(ref.findtext(c('f')) or '',out)
            if data is None: continue
            pt_count=next(ref.iter(c('ptCount')),None)
            assert pt_count is not None and int(pt_count.get('val'))==len(data),path
print(final.resolve())
