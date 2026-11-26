"""Preserve the baseline's OOXML objects; transplant authored cell content and share probe formulas."""
import json, re, sys, zipfile
from copy import deepcopy
from pathlib import Path
from lxml import etree as E
from dashboard import SHEET as REPORT_DATA
from dashboard_package import install_dashboard
import risk_cleanup
import risk_presentation
import risk_workbook
import report_layout
from plan_report import APPENDIX, CONVERSION_DETAIL, INDEX
from targeting import FIRST_TRIAL, LAST_TRIAL
import workbook_layout
import workbook_metadata
import input_validation
from workbook_protection import WorksheetProtection
import native_risk
import native_layout
import decision_layout
import native_package
from state_policy import WORKBOOK_JURISDICTION_ROW
from openpyxl.formula.translate import Translator
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string

X='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
R='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
C='http://schemas.openxmlformats.org/drawingml/2006/chart'
A='http://schemas.openxmlformats.org/drawingml/2006/main'
DRAW='http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
PKG='http://schemas.openxmlformats.org/package/2006/relationships'
CT='http://schemas.openxmlformats.org/package/2006/content-types'
def q(n,t): return f'{{{n}}}{t}'
def x(t): return q(X,t)
def read(z,p): return E.fromstring(z.read(p))
def pack(tree): return E.tostring(tree,encoding='UTF-8',xml_declaration=True,standalone=True)

def shared_strings(archive):
    part = 'xl/sharedStrings.xml'
    return [''.join(item.itertext()) for item in read(archive, part)] if part in archive.namelist() else []


def paths(z):
    ws=read(z,'xl/workbook.xml'); rel=read(z,'xl/_rels/workbook.xml.rels'); by_id={r.get('Id'):r.get('Target') for r in rel}
    return {s.get('name'):(lambda p:p[1:] if p.startswith('/') else 'xl/'+p)(by_id[s.get(q(R,'id'))]) for s in ws.find(x('sheets'))}

def main():
    base, plan_path, stage, target=map(Path,sys.argv[1:5]); plan=json.loads(plan_path.read_text()); changes={}
    target.parent.mkdir(parents=True,exist_ok=True)
    with zipfile.ZipFile(base) as src, zipfile.ZipFile(stage) as authored:
        sp,ap=paths(src),paths(authored)
        wb=read(src,'xl/workbook.xml'); sheets=wb.find(x('sheets')); rel=read(src,'xl/_rels/workbook.xml.rels'); ct=read(src,'[Content_Types].xml')
        next_sheet=max(int(s.get('sheetId')) for s in sheets)+1
        next_rel=max(int(s.get('Id')[3:]) for s in rel if re.fullmatch(r'rId\d+',s.get('Id')))+1
        next_part=max(int(re.search(r'sheet(\d+).xml',p)[1]) for p in sp.values())+1
        skip = risk_cleanup.retire_legacy(src, wb, rel, ct)
        sst = shared_strings(authored)
        baseline_strings = shared_strings(src)
        orig={n:{c.get('r'):c for c in read(src,sp[n]).iter(x('c'))} for n in ('Inputs','Summary','Projection')}
        styles={'title':('Summary','A1'),'band':('Inputs','A73'),'header':('Inputs','A74'),'note':('Inputs','G53'),
                'label':('Inputs','A53'),'money':('Projection','AW6'),'whole':('Projection','C6'),
                'percent_active':('Inputs','B14'),'percent_input':('Inputs','D14'),
                'input_active':('Inputs','B53'),'input':('Inputs','D53'),'summary_money':('Summary','B27')}
        style_ids={k:orig[n][a].get('s') for k,(n,a) in styles.items()}
        style_root=read(src,'xl/styles.xml'); xfs=style_root.find(x('cellXfs'))
        numfmts=style_root.find(x('numFmts'))
        cent_fmt=max([163]+[int(item.get('numFmtId')) for item in numfmts])+1
        E.SubElement(numfmts,x('numFmt'),numFmtId=str(cent_fmt),formatCode='#,##0.00;(#,##0.00);0.00')
        access_fmt = cent_fmt + 1
        E.SubElement(numfmts, x('numFmt'), numFmtId=str(access_fmt),
                     formatCode='[=126]"Access unconfirmed";0')
        decimal_fmt = access_fmt + 1
        E.SubElement(numfmts, x('numFmt'), numFmtId=str(decimal_fmt), formatCode='0.0000000000')
        numfmts.set('count',str(len(numfmts)))
        for name,source,fmt in [('input','input',1),('input_active','input_active',1),('whole','whole',1),('cents','money',cent_fmt),('published_percent','percent_active',10),
                                ('access_input','input',access_fmt),('access_active','input_active',access_fmt),
                                ('decimal','money',decimal_fmt),('decimal_input','input',cent_fmt),
                                ('report_date','money',14)]:
            xf=deepcopy(xfs[int(style_ids[source] or 0)]); xf.set('numFmtId',str(fmt)); xf.set('applyNumberFormat','1')
            style_ids[name]=str(len(xfs)); xfs.append(xf)
        active_fill, number_formats = report_layout.install_styles(style_root, style_ids)
        protection = WorksheetProtection(style_root)
        overlay = {}
        revised_sst = []
        if len(sys.argv)>5:
            with zipfile.ZipFile(sys.argv[5]) as revised:
                revised_sst = shared_strings(revised)
                for n,p in paths(revised).items():
                    overlay[n]={c.get('r'):c for c in read(revised,p).iter(x('c'))}
        with_state_policy = f'B{WORKBOOK_JURISDICTION_ROW}' in plan['cells'].get('Inputs', {})
        for n,edits in plan['cells'].items():
            new=n not in sp
            if new:
                part=f'xl/worksheets/sheet{next_part}.xml'; rid=f'rId{next_rel}'
                next_part+=1; next_rel+=1
                sa={'name':n,'sheetId':str(next_sheet),q(R,'id'):rid}; next_sheet+=1
                E.SubElement(sheets,x('sheet'),attrib=sa)
                sp[n] = part
                E.SubElement(rel,q(PKG,'Relationship'),Id=rid,Type=R+'/worksheet',Target=part[3:])
                E.SubElement(ct,q(CT,'Override'),PartName='/'+part,ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml')
                root=E.Element(x('worksheet'),nsmap={None:X,'r':R}); E.SubElement(root,x('dimension'),ref='A1')
                views=E.SubElement(root,x('sheetViews')); view=E.SubElement(views,x('sheetView'),workbookViewId='0',showGridLines='0',zoomScale='90')
                fr='1' if n in {'Funding_Calc','Targeting_Calc',REPORT_DATA} else '5' if n == APPENDIX else '4'
                E.SubElement(view,x('pane'),ySplit=fr,topLeftCell='A'+str(int(fr)+1),activePane='bottomLeft',state='frozen')
                E.SubElement(view,x('selection'),pane='bottomLeft',activeCell='A'+str(int(fr)+1),sqref='A'+str(int(fr)+1))
                E.SubElement(root,x('sheetFormatPr'),defaultRowHeight='16')
                cols=E.SubElement(root,x('cols'))
                for col in range(1,max(column_index_from_string(coordinate_from_string(a)[0]) for a in edits)+1):
                    w=10 if col<=2 and n not in {'Funding_Calc','Targeting_Calc'} else (27 if n=='Checks' and col==9 or n=='Cash_Flow' and col==12 or n=='Healthcare' and col==20 else 19)
                    if n == 'Risk_Settings' and col <= 3: w = (35, 16, 67)[col - 1]
                    E.SubElement(cols,x('col'),min=str(col),max=str(col),width=str(w),customWidth='1')
                    if n == 'Cash_Flow' and 30 <= col <= 39:
                        cols[-1].set('hidden', '1')
                E.SubElement(root,x('sheetData'))
            else: part=sp[n]; root=read(src,part)
            if n == 'Summary':
                report_layout.prepare_summary(root)
            data=root.find(x('sheetData'))
            if n == 'Income_Targeting':
                # The template's former sweep is now authored on Conversion Detail.
                for row in list(data):
                    if int(row.get('r')) >= 41:
                        data.remove(row)
            sr=read(authored,ap[n]); staged={c.get('r'):c for c in sr.iter(x('c'))}
            revised_cells = overlay.get(n, {})
            staged.update(revised_cells)
            old={c.get('r'):c for c in data.iter(x('c'))}; rows={int(r.get('r')):r for r in data}; sequence=0
            def obtain(a):
                if a in old: return old[a]
                col,r=coordinate_from_string(a)
                if r not in rows:
                    rows[r]=E.SubElement(data,x('row'),r=str(r))
                    if (new and r in (1,2,4)) or (n=='Inputs' and r in (85,98)):
                        rows[r].set('ht','32' if r in (4,98) or n=='Funding_Calc' else '22'); rows[r].set('customHeight','1')
                cell=E.SubElement(rows[r],x('c'),r=a); old[a]=cell; return cell
            for a,val in edits.items():
                dst=obtain(a); sc=staged.get(a); col,r=coordinate_from_string(a)
                style=plan['styles'].get(n,{}).get(a)
                if new and style=='money' and (col in ('A','B') or n=='Roth_Access' and col in ('C','N','O')): style='whole'
                if n=='Checks' and r>=5 and col in 'CDEFGHKLM': style='cents'
                if n=='Cash_Flow' and r>=5 and col in ('U','Z'): style='cents'
                if n=='Tax_Tables' and col in ('C','D') and (204<=r<=209 or 221<=r<=226): style='published_percent'
                if style and style_ids[style] is not None: dst.set('s',style_ids[style])
                for ch in list(dst):
                    if ch.tag in {x('f'),x('v'),x('is')}: dst.remove(ch)
                dst.attrib.pop('t',None)
                if val=='': continue
                if sc is None: raise ValueError((n,a,'Missing authored cell',val))
                if isinstance(val,str) and val.startswith('='):
                    ff=sc.findtext(x('f'))
                    if ff is None or ff.strip()!=val[1:].strip(): raise ValueError((n,a,ff,val))
                    formula=E.SubElement(dst,x('f')); formula.text=ff.strip()
                    if a in plan['shared'].get(n,{}):
                        ref=plan['shared'][n][a]; end=coordinate_from_string(ref.split(':')[1])[1]
                        # Excel for Mac can discard shared groups that refer to other sheets.
                        if '!' in formula.text:
                            translator = Translator('='+formula.text, origin=a)
                            for i in range(r+1,end+1):
                                address = f'{col}{i}'
                                E.SubElement(obtain(address),x('f')).text = translator.translate_formula(dest=address)[1:]
                        else:
                            formula.set('t','shared'); formula.set('si',str(sequence)); formula.set('ref',ref)
                            for i in range(r+1,end+1): E.SubElement(obtain(f'{col}{i}'),x('f'),t='shared',si=str(sequence))
                            sequence+=1
                elif isinstance(val,str):
                    t=sc.get('t'); value=sc.findtext(x('v')); inline=sc.find(x('is'))
                    strings = revised_sst if a in revised_cells else sst
                    actual=strings[int(value)] if t=='s' else ''.join(inline.itertext()) if inline is not None else value
                    if actual!=val: raise ValueError((n,a,actual,val))
                    dst.set('t','inlineStr'); E.SubElement(E.SubElement(dst,x('is')),x('t')).text=actual
                else:
                    value=sc.findtext(x('v')); assert float(value)==val,(n,a,value,val)
                    E.SubElement(dst,x('v')).text=value
            for r,row in rows.items():
                row[:]=sorted(row,key=lambda c:column_index_from_string(coordinate_from_string(c.get('r'))[0])); row.attrib.pop('spans',None)
            data[:]=[rows[r] for r in sorted(rows)]
            if new and n not in {'Funding_Calc','Targeting_Calc',REPORT_DATA,APPENDIX,INDEX,CONVERSION_DETAIL,risk_workbook.SHEET,risk_presentation.SHEET}:
                merged=E.SubElement(root,x('mergeCells'),count='1')
                E.SubElement(merged,x('mergeCell'),ref='A2:C2' if n == 'Risk_Settings' else 'A2:I2')
            if n=='Projection':
                cols=root.find(x('cols'))
                if cols is not None:
                    E.SubElement(cols,x('col'),min='72',max='84',width='19',customWidth='1')
            max_col=max((coordinate_from_string(a)[0] for a in old),key=column_index_from_string)
            root.find(x('dimension')).set('ref',f'A1:{max_col}{max(rows)}')
            layout_name = native_risk.BASE_LAYOUT.get(n, n)
            workbook_layout.apply_sheet_layout(layout_name, root)
            report_layout.apply(layout_name, root, style_ids, number_formats, active_fill,
                                with_state_policy=with_state_policy)
            native_layout.apply(n, root, style_ids, number_formats, active_fill)
            decision_layout.apply(n, root, style_ids, number_formats, baseline_strings)
            input_validation.apply(n, root, with_state_policy=n == 'Inputs' and with_state_policy)
            protection.apply(n, root)
            workbook_metadata.clear_revision_ids(root)
            changes[part]=pack(root); print(n,len(edits),flush=True)
        changes['xl/styles.xml'] = pack(style_root)
        app = read(src, 'docProps/app.xml')
        workbook_layout.apply_tab_order(wb, app)
        report_layout.install_print_names(wb)
        install_dashboard(src, authored, wb, changes, sp['Summary'], ap['Summary'])
        native_package.install(wb,changes,plan.get('native_names',{}))
        # Keep the native conversion charts below the comparison controls.
        part = 'xl/drawings/drawing2.xml'
        drawing = read(src, part)
        decision_layout.conversion_charts(drawing)
        changes[part] = pack(drawing)
        wb.find(x('calcPr')).set('fullCalcOnLoad','1')
        wb.find(x('calcPr')).set('calcMode','auto')
        skip.update(workbook_metadata.clean_metadata(src, wb, app, ct, rel, changes))
        workbook_metadata.compact_shared_strings(src, (part for part in sp.values() if part not in skip), changes)
        changes['docProps/app.xml'] = pack(app)
        changes.update({'xl/workbook.xml':pack(wb),'xl/_rels/workbook.xml.rels':pack(rel),'[Content_Types].xml':pack(ct)})
        for index, title in [(2,'Incremental tax and healthcare cost per dollar converted'),(3,'Current-year tax and healthcare costs')]:
            part=f'xl/charts/chart{index}.xml'; chart=read(src,part)
            for formula in chart.iter(q(C, 'f')):
                sheet, sep, address = (formula.text or '').partition('!')
                if sep and sheet.strip("'") == 'Income_Targeting':
                    rows = {41: FIRST_TRIAL - 1, 42: FIRST_TRIAL, 92: LAST_TRIAL}
                    address = re.sub(r'\$(41|42|92)\b', lambda match: f'${rows[int(match[1])]}', address)
                    formula.text = f"'{CONVERSION_DETAIL}'!{address}"
            title_node=chart.find(f'.//{q(C,"chart")}/{q(C,"title")}')
            if title_node is not None:
                runs=list(title_node.iter(q(A,'t')))
                if runs:
                    runs[0].text=title
                    for run in runs[1:]: run.text=''
            decision_layout.conversion_chart_style(chart)
            changes[part]=pack(chart)
        with zipfile.ZipFile(target,'w',zipfile.ZIP_DEFLATED,compresslevel=5) as out:
            for p in src.namelist():
                if p not in skip: out.writestr(p,changes.pop(p) if p in changes else src.read(p))
            for p,content in changes.items(): out.writestr(p,content)
    print(target)

if __name__=='__main__': main()
