"""Capture and present three complete funding paths from the annual engine."""
from openpyxl.utils.cell import get_column_letter as L
from native_capture import CONTEXT,ANNUAL_CAPACITY

TRIAL='Risk_Trials'
SHEET='Funding_Paths'
DATA_SHEET='Funding_Path_Data'
TITLE_ROW=1
HEADER_ROW=6
CURRENT_ROW=7
FIRST_ROW=8
LAST_ROW=FIRST_ROW+2

# Unit distinguishes a calculation index from money that follows Summary's
# display setting. These outputs are links, not another financial engine.
FIELDS=(
    ('Year','Trial_Projection!B{row}','year'),
    ('Age','Trial_Projection!C{row}','age'),
    ('Price index','Trial_Projection!G{row}','index'),
    ('Income available','Trial_Cash_Flow!C{row}','money'),
    ('Income used','MIN(Trial_Cash_Flow!C{row},Trial_Cash_Flow!N{row}+Trial_Cash_Flow!P{row})','money'),
    ('Living costs','Trial_Projection!AW{row}','money'),
    ('One-time costs','Trial_Projection!AX{row}','money'),
    ('Gross healthcare','Trial_Projection!BA{row}','money'),
    ('Spending due, including prior bills','Trial_Cash_Flow!D{row}','money'),
    ('Tax due, including prior tax','Trial_Cash_Flow!F{row}+Trial_Cash_Flow!M{row}','money'),
    ('HSA-eligible healthcare','Trial_Healthcare!Q{row}','money'),
    ('HSA used','Trial_Cash_Flow!E{row}','money'),
    ('Tax paid','Trial_Cash_Flow!N{row}','money'),
    ('Spending paid','Trial_Cash_Flow!P{row}','money'),
    ('Brokerage available','Trial_Cash_Flow!H{row}','money'),
    ('Traditional available under plan rules','Trial_Cash_Flow!I{row}','money'),
    ('Roth available under plan rules','Trial_Cash_Flow!J{row}','money'),
    ('Brokerage withdrawn','Trial_Projection!V{row}','money'),
    ('Traditional withdrawn','Trial_Projection!W{row}','money'),
    ('Roth withdrawn','Trial_Projection!Y{row}','money'),
    ('Roth conversion','Trial_Projection!AA{row}','money'),
    ('Roth IRA savings','Trial_Projection!M{row}','money'),
    ('Taxable deposit','Trial_Projection!BB{row}','money'),
    ('Brokerage closing balance','Trial_Projection!BC{row}','money'),
    ('Traditional closing balance','Trial_Projection!BE{row}','money'),
    ('Roth closing balance','Trial_Projection!BF{row}','money'),
    ('HSA closing balance','Trial_Projection!BG{row}','money'),
    ('Total closing portfolio','Trial_Projection!BH{row}','money'),
    ('Outstanding tax','Trial_Cash_Flow!O{row}','money'),
    ('Outstanding spending','Trial_Cash_Flow!Q{row}','money'),
    ('Newly unpaid spending','Trial_Cash_Flow!AU{row}','money'),
    ('Prior spending repaid','Trial_Cash_Flow!AS{row}','money'),
    ('Cumulative newly unpaid spending','Trial_Cash_Flow!AV{row}','money'),
    ('Temporary living-cost cut','Market_Paths!AE{market_row}','money'),
    ('Social Security','Trial_Projection!N{row}','money'),
    ('Pay before tax and savings','Trial_Projection!I{row}','money'),
    ('Cash residual','Trial_Cash_Flow!U{row}','money'),
    ('Current-year tax','Trial_Cash_Flow!M{row}','money'),
)
FIELD_COUNT=len(FIELDS)
LAST_COLUMN=3+ANNUAL_CAPACITY*FIELD_COUNT
SIGNATURE_COLUMN=LAST_COLUMN+1
SIGNATURE=L(SIGNATURE_COLUMN)


def configure(m):
    put=lambda a,v,s='money':m.put(DATA_SHEET,a,v,s)
    put(f'A{TITLE_ROW}','Complete low, median and high funding paths','detail_title')
    put(f'A{TITLE_ROW+1}','The runner captures the full annual ledger for three ranked fixed-spending paths. Read Funding_Paths for a compact table.','report_caption')
    put(f'A{HEADER_ROW}','Trial','header')
    put(f'B{HEADER_ROW}','Assumptions used','header')
    put(f'C{HEADER_ROW}','Calculation valid','header')
    put(f'B{CURRENT_ROW}','='+CONTEXT,'report_caption')
    put(f'C{CURRENT_ROW}',f'={TRIAL}!$C$7','whole')
    for elapsed in range(ANNUAL_CAPACITY):
        row=elapsed+5
        for offset,(label,expression,unit) in enumerate(FIELDS):
            column=L(4+elapsed*FIELD_COUNT+offset)
            put(f'{column}{HEADER_ROW}',f'Year {elapsed}: {label}','header')
            formula=expression.format(row=row,market_row=row+1)
            put(f'{column}{CURRENT_ROW}',f'=IF({elapsed}<Trial_Inputs!$B$70,{formula},"")',
                'whole' if unit in ('year','age') else 'decimal' if unit=='index' else 'money')
    put(f'{SIGNATURE}{HEADER_ROW}','Assumptions after capture','header')
    put(f'{SIGNATURE}{CURRENT_ROW}','='+CONTEXT,'report_caption')
    m.native_names[TRIAL].update({
        'RiskPathCurrent':f"'{DATA_SHEET}'!$B${CURRENT_ROW}:${SIGNATURE}${CURRENT_ROW}",
        'RiskPathResults':f"'{DATA_SHEET}'!$B${FIRST_ROW}:${SIGNATURE}${LAST_ROW}",
        'RiskPathTrials':f"'{DATA_SHEET}'!$A${FIRST_ROW}:$A${LAST_ROW}",
    })
    _presentation(m)
    m.MODULES.extend((SHEET,DATA_SHEET))


def _presentation(m):
    put=lambda a,v,s='report_money':m.put(SHEET,a,v,s)
    put('A1','Funding in three complete market paths','detail_title')
    put('A2','=IF(Risk_Data!$Z$1="current",IF(Summary!$K$18="Future dollars","Future dollars","Today’s dollars")&" from Summary. Each path keeps its own annual prices.","Run market risk on Summary to update this detail.")','report_caption')
    put('A3','Low, median and high are actual fixed-spending paths ranked by ending balance. Opening balances have no annual income or costs.','report_caption')
    put('AS1','Back to Summary','report_link')
    put('AS2','Inspect the trial calculation','report_link')
    for column,label in enumerate(('Path','Trial',*(label for label,_,_ in FIELDS)),1):
        put(f'{L(column)}5',label,'header')
    source=f'{DATA_SHEET}!$D${FIRST_ROW}:${L(LAST_COLUMN)}${LAST_ROW}'
    n='Inputs!$B$70'
    for row in range(6,6+3*ANNUAL_CAPACITY):
        offset=row-6
        case=f'1+INT({offset}/{n})'
        period=f'MOD({offset},{n})'
        on=f'AND(Risk_Data!$Z$1="current",{offset}<3*{n})'
        put(f'A{row}',f'=IF({on},CHOOSE({case},"Low","Median","High"),"")','label')
        put(f'B{row}',f'=IF({on},INDEX({DATA_SHEET}!$A${FIRST_ROW}:$A${LAST_ROW},{case}),"")','whole')
        price=f'INDEX({source},{case},({period})*{FIELD_COUNT}+3)'
        for field,(_,_,unit) in enumerate(FIELDS,1):
            value=f'INDEX({source},{case},({period})*{FIELD_COUNT}+{field})'
            if unit=='money':value=f'({value})/IF(Summary!$K$18="Future dollars",1,{price})'
            put(f'{L(field+2)}{row}',f'=IF({on},{value},"")',
                'whole' if unit in ('year','age') else 'decimal' if unit=='index' else 'report_money')
