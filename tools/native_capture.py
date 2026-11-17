"""Declare visible output ranges read by the workbook's on-demand runner.

The runner changes only Risk_Trials!B3. The current output row links to the
second instance of the annual financial compiler. The main plan stays fixed.
"""
from openpyxl.utils.cell import get_column_letter as L

from risk_engine.scenario import SCENARIO_INPUT_ROWS, ASSET_ROWS, TAX_TABLE_LAST_ROW

SHEET='Risk_Trials'
MARKET='Market_Paths'
FIRST=8
CAPACITY=1000
LAST=FIRST+2*CAPACITY+33-1
ANNUAL_FIRST=34
ANNUAL_CAPACITY=126
SIGNATURE_COLUMN=ANNUAL_FIRST+2*ANNUAL_CAPACITY
SIGNATURE=L(SIGNATURE_COLUMN)
RANK=L(SIGNATURE_COLUMN+1)
FAILURE_RANK=L(SIGNATURE_COLUMN+2)
CONTEXT=f'{MARKET}!$BB$3'

FIELDS=(
    'Assumptions used','Calculation valid','Paid every year’s costs','First shortfall age',
    'Largest unpaid balance, today’s dollars','Ending portfolio, today’s dollars',
    'Ending portfolio, future dollars','Years with spending cuts','Total cuts, today’s dollars',
    'First shortfall year','Price index at first shortfall','First unpaid balance, today’s dollars',
    'First unpaid spending, today’s dollars','First unpaid tax, today’s dollars',
    'Spending due at first shortfall, today’s dollars','Tax due at first shortfall, today’s dollars',
    'Income available at first shortfall, today’s dollars','Income used at first shortfall, today’s dollars',
    'HSA used at first shortfall, today’s dollars','HSA-eligible costs at first shortfall, today’s dollars',
    'Brokerage available at first shortfall, today’s dollars','Brokerage used at first shortfall, today’s dollars',
    'Traditional available at first shortfall, today’s dollars','Traditional used at first shortfall, today’s dollars',
    'Roth available at first shortfall, today’s dollars','Roth used at first shortfall, today’s dollars',
    'Refresh date','Ending unpaid tax, today’s dollars','Ending unpaid spending, today’s dollars',
    'Ending price index','First unpaid balance, future dollars','Selected end age',
)


def _context(m,baseline):
    """Compare all financial drivers, including numeric tax data, without a hash."""
    refs=[('Active scenario','Inputs!B4')]
    refs.extend((f'Selected input {row}',f'INDEX(Inputs!D{row}:F{row},1,Inputs!B4)')
                for row in SCENARIO_INPUT_ROWS if row!=7)
    refs.extend((f'Effective input {row}',f'Inputs!B{row}')
                for row in (*range(17,25),*range(63,67),130,131,132,134,135))
    refs.extend((f'Expense {col}{row}',f'Inputs!{col}{row}')
                for row in range(75,84) for col in 'BC')
    refs.extend((f'Asset {col}{row}',f'Asset_Allocation!{col}{row}')
                for row in ASSET_ROWS for col in 'BCDE')
    refs.extend((f'Correlation {col}{row}',f'Asset_Allocation!{col}{row}')
                for row in range(16,21) for col in 'BCDEF')
    refs.extend((f'Mix {stage} {col}{asset}',
                 f'INDEX(Asset_Allocation!{col}26:{col}48,{asset}+9*(Inputs!B{stage}-1))')
                for stage in (15,16) for col in 'BCDE' for asset in range(1,6))
    refs.extend((f'Risk input {row}',f'Risk_Settings!B{row}')
                for row in (*range(5,10),*range(14,27),*range(29,37)))
    for row in range(1,TAX_TABLE_LAST_ROW+1):
        for col in 'ABCDEFGH':
            address=f'{col}{row}'
            value=m.cells['Tax_Tables'].get(address,baseline['Tax_Tables'][address].value)
            if type(value) in (int,float) or isinstance(value,str) and value.startswith('='):
                refs.append((f'Tax data {address}',f'Tax_Tables!{address}'))
    put=lambda a,v,s='report_caption':m.put(MARKET,a,v,s)
    put('BA1','Assumptions used by the simulation','detail_title')
    put('BA2','Current typed values are compared with the completed run and its captured trials.','report_caption')
    for col,label in zip(('BA','BB','BC','BD'),('Financial driver','Current value','Typed value','Context segment')):
        put(f'{col}4',label,'header')
    for row,(label,ref) in enumerate(refs,5):
        put(f'BA{row}',label)
        put(f'BB{row}',f'=IF(ISBLANK({ref}),"",{ref})')
        # Scientific notation keeps units and small rates. Type prefixes distinguish
        # a missing input, a valid zero, and text that happens to look numeric.
        put(f'BC{row}',f'=IF(ISNUMBER(BB{row}),"n"&TEXT(BB{row},"0.00000000000000E+00"),IF(BB{row}="","b","t"&BB{row}))&"|"')
    chunks=[]
    for row,start in enumerate(range(5,5+len(refs),16),5):
        stop=min(start+16,5+len(refs))
        put(f'BD{row}','='+'&'.join(f'BC{i}' for i in range(start,stop)))
        chunks.append(f'BD{row}')
    put('BA3','Current assumption record','label')
    put('BB3','="native-risk-v1|"&'+'&'.join(chunks))
    return len(refs)


def _outputs(m):
    put=lambda a,v,s='money':m.put(SHEET,a,v,s)
    n='Trial_Inputs!$B$70'
    p=lambda c:f'Trial_Projection!${c}$5:${c}$130'
    cash=lambda c:f'Trial_Cash_Flow!${c}$5:${c}$130'
    end=lambda rng:f'INDEX({rng},{n})'
    has=f'{n}>1'
    first='IF($E$7="",1,$E$7-Trial_Inputs!$B$8+1)'
    at=lambda rng:f'INDEX({rng},{first})'
    price=at(p('G'))
    gap=lambda expression:f'IF($E$7<>"",({expression})/{price},"")'
    valid=(f'--AND(Market_Paths!$J$3="OK",COUNTIF(Trial_Cash_Flow!$AM$5:$AM$130,"OK")={n},'
           'COUNTIF(Trial_Healthcare!$T$6:$T$130,"Missing*")+COUNTIF(Trial_Healthcare!$T$6:$T$130,"Confirm*")=0)')
    expr=[CONTEXT,valid,f'--(MAX({p("BR")})<0.01)',
          f'IF(MIN({p("CI")})=999,"",MIN({p("CI")}))',
          f'MAX({p("CJ")})',end(p('BN')),end(p('BH')),
          'COUNTIF(Market_Paths!$AE$7:$AE$131,">0")','SUM(Market_Paths!$AF$7:$AF$131)',
          f'IF($E$7<>"",{at(p("B"))},"")',f'IF($E$7<>"",{price},"")',
          gap(at(p('BR'))),gap(at(cash('Q'))),gap(at(cash('O'))),gap(at(cash('D'))),
          gap(f'{at(cash("F"))}+{at(cash("M"))}'),gap(at(cash('C'))),
          gap(f'MIN({at(cash("C"))},{at(cash("N"))}+{at(cash("P"))})'),
          gap(at(cash('E'))),gap(at('Trial_Healthcare!$Q$5:$Q$130')),
          *(gap(at(cash(col))) for col in ('H',)),gap(at(p('V'))),
          gap(at(cash('I'))),gap(at(p('W'))),gap(at(cash('J'))),gap(at(p('Y'))),
          'TODAY()',f'{end(cash("O"))}/{end(p("G"))}',
          f'{end(cash("Q"))}/{end(p("G"))}',end(p('G')),
          f'IF($E$7<>"",{at(p("BR"))},"")','Trial_Inputs!B53']
    assert len(expr)==len(FIELDS)==32
    put('A6','Trial','header')
    for column,(label,formula) in enumerate(zip(FIELDS,expr,strict=True),2):
        address=f'{L(column)}7'
        put(f'{L(column)}6',label,'header')
        if column in (2,28):wrapped=formula
        else:wrapped=f'IF({has},{formula},{0 if column==3 else chr(34)+chr(34)})'
        put(address,'='+wrapped)
    for offset in range(ANNUAL_CAPACITY):
        for nominal in (False,True):
            column=ANNUAL_FIRST+offset+int(nominal)*ANNUAL_CAPACITY
            put(f'{L(column)}6',f'Year {offset} balance, {"future" if nominal else "today’s"} dollars','header')
            put(f'{L(column)}7',f'=IF({offset}<{n},Trial_Projection!{"BH" if nominal else "BN"}{offset+5},"")')
    put(f'{SIGNATURE}6','Assumptions after capture','header')
    put(f'{SIGNATURE}7','='+CONTEXT,'report_caption')
    # Per-year flags simplify first-failure retrieval without array formulas.
    m.put('Trial_Projection','CI4','Shortfall age, or 999','header')
    m.put('Trial_Projection','CJ4','Unpaid balance, today’s dollars','header')
    for row in range(5,131):
        m.put('Trial_Projection',f'CI{row}',f'=IF({row}-4<={n},IF(BR{row}>=0.01,C{row},999),"")','whole')
        m.put('Trial_Projection',f'CJ{row}',f'=IF({row}-4<={n},BR{row}/G{row},"")','money')


def configure(m,baseline):
    _context(m,baseline)
    _outputs(m)
    m.put(SHEET,'A7','=B3','whole')
    # The runner writes a contiguous batch: fixed, flexible, early-loss and
    # required stresses. Unused capacity remains genuinely blank.
    m.native_names={SHEET:{
        'RiskTrialCurrent':f'$B$7:${SIGNATURE}$7',
        'RiskTrialResults':f'$B${FIRST}:${SIGNATURE}${LAST}',
        'RiskTrialIDs':f'$A${FIRST}:$A${LAST}',
    }}
