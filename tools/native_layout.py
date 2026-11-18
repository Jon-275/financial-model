"""Use the workbook's shared design for native risk calculation details."""
from lxml import etree as E
from openpyxl.utils.cell import get_column_letter as L,column_index_from_string as C
from report_layout import X,_columns,_rows,_merges,_stripe,_style_table,_replace,_link,_collapse_empty_return_rows
from native_risk import MODEL_SHEETS,TRIAL,MARKET,RANDOM
from native_capture import COLUMN_STYLES,LAST,SIGNATURE_COLUMN
import native_results
import native_funding_paths


def _column_styles(root,styles,overrides):
    for column in root.find(X+'cols'):
        number=int(column.get('min'))
        column.set('style',styles[overrides.get(number,'report_money')])


def apply(name, root, styles, numbers, active_fill):
    if name==native_results.SHEET:
        widths=[20]*C('BN')
        for address,width in {'A':11,'B':10,'K':29,'Z':28,'AA':38,'AB':30,'AD':4,'AE':4,
                              'AF':14,'BA':34,'BB':32,'BC':4,'BD':4,'BE':4,'BF':4,'BG':4,'BH':4}.items():
            widths[C(address)-1]=width
        for number in range(C('S'),C('Y')+1):widths[number-1]=4
        _columns(root,widths)
        _merges(root,['A1:I1','A2:I2','A3:C3','D3:E3','F3:G3','H3:I3',
                      'K1:Q1','K20:R20','AF1:AZ1','AA4:AD4'])
        for block in root.findall(X+'conditionalFormatting'):root.remove(block)
        for area,key in (('A5:I130','$A5'),('K5:Q17','$K5'),('K23:R29','$K23'),
                         ('AF5:AZ7','$AF5'),('BI5:BN130','$A5')):
            _stripe(root,area,key,active_fill+2)
        _rows(root,4,21,{1:33,2:30,3:24,4:70,20:30,22:60})
        # The annual source tables use conditional formatting, so rows after
        # the selected end age remain blank and unshaded.
        links=root.find(X+'hyperlinks')
        for link in links:
            if link.get('ref')=='F3':link.set('location',TRIAL+'!A1')
        return
    if name=='Risk_Settings':
        merged=root.find(X+'mergeCells')
        areas=[item.get('ref') for item in merged] if merged is not None else []
        for row in (28,38,39,40,42,43,44):
            area=f'A{row}:C{row}'
            if area not in areas:areas.append(area)
        _merges(root,areas)
        _rows(root,44,24,{1:33,2:36,4:35,13:36,28:29,38:29,39:36,40:36,42:29,43:29,44:55})
        return
    if name==native_funding_paths.SHEET:
        _columns(root,[15,12,12,10,18,*([21]*(native_funding_paths.FIELD_COUNT-3)),*([4]*(C('AS')-native_funding_paths.FIELD_COUNT-3)),27])
        _merges(root,['A1:O1','A2:O2','A3:O3'])
        _rows(root,5,22,{1:33,2:30,3:34,5:70})
        _stripe(root,f'A6:{L(native_funding_paths.FIELD_COUNT+2)}{5+3*126}','$A6',active_fill+2)
        _link(root,{'AS1':'Summary!B3','AS2':TRIAL+'!A1'})
        view=root.find(f'{X}sheetViews/{X}sheetView')
        for item in view.findall(X+'pane')+view.findall(X+'selection'):view.remove(item)
        E.SubElement(view,X+'pane',xSplit='4',ySplit='5',topLeftCell='E6',activePane='bottomRight',state='frozen')
        E.SubElement(view,X+'selection',pane='bottomRight',activeCell='E6',sqref='E6')
        return
    if name==native_funding_paths.DATA_SHEET:
        _columns(root,[12,28,14,*([20]*(native_funding_paths.SIGNATURE_COLUMN-3))])
        _column_styles(root,styles,{1:'whole',2:'report_caption',3:'whole',
                                   native_funding_paths.SIGNATURE_COLUMN:'report_caption'})
        for elapsed in range(126):
            for offset,(_,_,unit) in enumerate(native_funding_paths.FIELDS):
                if unit=='money':continue
                root.find(X+'cols')[3+elapsed*native_funding_paths.FIELD_COUNT+offset].set(
                    'style',styles['decimal' if unit=='index' else 'whole'])
        _merges(root,['A1:L1','A2:L2'])
        _rows(root,10,22,{1:33,2:36,6:85,7:26})
        _stripe(root,f'A8:{native_funding_paths.SIGNATURE}10','$A8',active_fill+2)
        view=root.find(f'{X}sheetViews/{X}sheetView')
        for item in view.findall(X+'pane')+view.findall(X+'selection'):view.remove(item)
        E.SubElement(view,X+'pane',ySplit='7',topLeftCell='A8',activePane='bottomLeft',state='frozen')
        E.SubElement(view,X+'selection',pane='bottomLeft',activeCell='A8',sqref='A8')
        return
    if name not in {*MODEL_SHEETS.values(),TRIAL,MARKET,RANDOM}:return
    header = 5 if name==MARKET else 4 if name==RANDOM else None
    if header:
        for cell in root.findall(f"{X}sheetData/{X}row[@r='{header}']/{X}c"):
            cell.set('s',styles['report_table_head'])
    view=root.find(f'{X}sheetViews/{X}sheetView')
    for item in view.findall(X+'pane')+view.findall(X+'selection'):view.remove(item)
    frozen = 1 if name=='Trial_Funding_Calc' else 5 if name==MARKET else 4
    E.SubElement(view,X+'pane',ySplit=str(frozen),topLeftCell=f'A{frozen+1}',activePane='bottomLeft',state='frozen')
    E.SubElement(view,X+'selection',pane='bottomLeft',activeCell=f'A{frozen+1}',sqref=f'A{frozen+1}')
    links=root.find(X+'hyperlinks')
    if links is not None:
        for link in links:
            location=link.get('location','')
            source, separator, address=location.partition('!')
            if separator and source.strip("'") in MODEL_SHEETS:
                link.set('location',MODEL_SHEETS[source.strip("'")]+separator+address)
    if name=='Trial_Inputs':
        _columns(root,(49,24,24))
        _merges(root,['A1:C1','A2:C2','A3:C3'])
        _rows(root,133,21,{1:32,2:30,3:26})
        _collapse_empty_return_rows(root)
        _stripe(root,'A4:B74','$A4',active_fill+2)
        _stripe(root,'A75:C82','$B75',active_fill+2)
        _stripe(root,'A83:B133','$A83',active_fill+2)
    elif name==MARKET:
        _columns(root,[13,11,11,15,*([16]*28),*([4]*8),*([19]*11)])
        _merges(root,['A1:J1','A2:J2','AO6:AS6','AO15:AS15'])
        _rows(root,131,20,{1:32,2:30,3:25,5:44,17:38})
        _stripe(root,'A6:AG131','$A6',active_fill+2)
        context_end=max(int(row.get('r')) for row in root.find(X+'sheetData'))
        _rows(root,context_end,20,{1:32,2:30,3:25,4:44,5:44,17:38})
        columns=root.find(X+'cols')
        for column,width in ((53,39),(54,28),(55,28),(56,28)):
            E.SubElement(columns,X+'col',min=str(column),max=str(column),width=str(width),customWidth='1')
        _merges(root,['A1:J1','A2:J2','AO6:AS6','AO15:AS15','BA1:BD1','BA2:BD2'])
        _stripe(root,f'BA5:BD{context_end}','$BA5',active_fill+2)
    elif name==RANDOM:
        _columns(root,[14]*44)
        _merges(root,['A1:P1','A2:P2','T1:AR1','T2:AR2'])
        _rows(root,1004,20,{1:32,2:30,3:25,4:44})
        _stripe(root,'A5:P754','$A5',active_fill+2)
        _stripe(root,'T5:AR1004','$T5',active_fill+2)
    elif name==TRIAL:
        _merges(root,['A1:L1','A2:L2'])
        _columns(root,[12,28,*([20]*SIGNATURE_COLUMN)])
        _rows(root,7,22,{1:32,2:36,4:32,6:85,7:30})
        _merges(root,['A1:L1','A2:L2','A4:L4'])
        _stripe(root,f'A8:{L(SIGNATURE_COLUMN)}{LAST}',
                'IF($A8>0,$A8,"")',active_fill+2)
        _column_styles(root,styles,COLUMN_STYLES)
        view.find(X+'pane').set('ySplit','7')
        view.find(X+'pane').set('topLeftCell','A8')
        view.find(X+'selection').set('activeCell','A8')
        view.find(X+'selection').set('sqref','A8')
    else:
        # All annual guides are readable above their detail tables.
        if name!='Trial_Funding_Calc':
            existing=root.find(X+'mergeCells')
            ranges=[item.get('ref') for item in existing] if existing is not None else []
            for merged in ('A1:I1','A2:I2'):
                if merged not in ranges:ranges.append(merged)
            _merges(root,ranges)
        if name=='Trial_Projection':
            _stripe(root,'A5:CJ130','$C5',active_fill+2)
