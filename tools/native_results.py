"""Derive risk statistics from captured Excel trials, with no external report."""
from openpyxl.utils.cell import get_column_letter as L
from native_capture import FIRST,CAPACITY,LAST,ANNUAL_FIRST,ANNUAL_CAPACITY,SIGNATURE,RANK,FAILURE_RANK,CONTEXT
import native_funding_paths

SHEET='Risk_Data'
TRIAL='Risk_Trials'
SCHEMA='retirement-risk-native-v1'
ENGINE='Excel 1.0'
STRESSES=('Inflation 4%','Inflation 5%','Plan through age 100','Plan through age 105',
          'Inflation 5%, weaker returns, age 105','Medical budget $18,000','Medical budget $25,000')


def _trial_range(column,flexible=False):
    data=f'{TRIAL}!${column}${FIRST}:${column}${LAST}'
    start='1+Risk_Settings!$B$29' if flexible else '1'
    stop='2*Risk_Settings!$B$29' if flexible else 'Risk_Settings!$B$29'
    return f'INDEX({data},{start}):INDEX({data},{stop})'


def _trial_at(column,position):
    return f'INDEX({TRIAL}!${column}${FIRST}:${column}${LAST},{position})'


def _batch_id_checks():
    """Require each completed result row to retain its documented trial ID."""
    ids=f'{TRIAL}!$A${FIRST}:$A${LAST}'
    n='Risk_Settings!$B$29'
    segments=(('1',n,n,0),(f'1+{n}',f'2*{n}',n,1000),
              (f'1+2*{n}',f'13+2*{n}','13',2000),
              (f'14+2*{n}',f'26+2*{n}','13',2100),
              (f'27+2*{n}',f'33+2*{n}','7',2200))
    checks=[f'COUNTA({ids})=2*{n}+33']
    for start,stop,length,base in segments:
        captured=f'INDEX({ids},{start}):INDEX({ids},{stop})'
        sequence=f'ROW($A$1:INDEX($A$1:$A${CAPACITY},{length}))+{base}'
        checks.append(f'SUMPRODUCT(--({captured}=({sequence})))={length}')
    return checks


def _rank_position(percentile):
    ranks=f'{TRIAL}!${RANK}${FIRST}:${RANK}${FIRST+CAPACITY-1}'
    return f'MATCH(1+INT({percentile}*(Risk_Settings!B29-1)+0.5),{ranks},0)'


def _ranks(m):
    end=FIRST+CAPACITY-1
    p=lambda c:f'${c}${FIRST}:INDEX(${c}${FIRST}:${c}${end},Risk_Settings!$B$29)'
    m.put(TRIAL,f'{RANK}6','Ending-balance rank','header')
    m.put(TRIAL,f'{FAILURE_RANK}6','Failure age for tied ranks','header')
    for row in range(FIRST,end+1):
        m.put(TRIAL,f'{FAILURE_RANK}{row}',f'=IF({row-FIRST+1}<=Risk_Settings!$B$29,IF(ISNUMBER(E{row}),E{row},1000000000),"")','whole')
        # Numeric comparisons retain the full stored value. COUNTIFS string
        # criteria can round a balance and count the current trial against itself.
        balance,age,gap,trial_id = (p(col) for col in ('G',FAILURE_RANK,'F','A'))
        later_tiebreak=(f'({gap}>F{row})+({gap}=F{row})*({trial_id}<A{row})')
        tied_balance=(f'({age}<{FAILURE_RANK}{row})+({age}={FAILURE_RANK}{row})*({later_tiebreak})')
        precedes=f'({balance}<G{row})+({balance}=G{row})*({tied_balance})'
        count=f'SUMPRODUCT(--({p("C")}=1),{precedes})'
        m.put(TRIAL,f'{RANK}{row}',f'=IF(AND({row-FIRST+1}<=Risk_Settings!$B$29,C{row}=1),1+{count},"")','whole')


def _status(put, *, stored_context=CONTEXT):
    valid_count='(2*Risk_Settings!B29+33)'
    contexts=_batch_id_checks()
    for column in ('B',SIGNATURE):
        span=f'{TRIAL}!${column}${FIRST}:INDEX({TRIAL}!${column}${FIRST}:${column}${LAST},{valid_count})'
        contexts.append(f'SUMPRODUCT(--EXACT({span},{stored_context}))={valid_count}')
    put('Z5',f'=IFERROR(IF(AND(COUNT({TRIAL}!C{FIRST}:C{LAST})={valid_count},'
        f'AB3={valid_count},{",".join(contexts)}),"current","incomplete"),"incomplete")','report_caption')
    secondary=[]
    for column in ('B',native_funding_paths.SIGNATURE):
        span=f'{native_funding_paths.DATA_SHEET}!${column}${native_funding_paths.FIRST_ROW}:${column}${native_funding_paths.LAST_ROW}'
        secondary.append(f'SUMPRODUCT(--EXACT({span},{stored_context}))=3')
    secondary.append(f'COUNTIF({native_funding_paths.DATA_SHEET}!$C${native_funding_paths.FIRST_ROW}:$C${native_funding_paths.LAST_ROW},1)=3')
    for row,percentile in enumerate((.1,.5,.9),native_funding_paths.FIRST_ROW):
        ids=f'{TRIAL}!$A${FIRST}:$A${FIRST+CAPACITY-1}'
        secondary.append(f'{native_funding_paths.DATA_SHEET}!$A${row}=INDEX({ids},{_rank_position(percentile)})')
    complete=f'AND(Z5="current",{",".join(secondary)})'
    put('Z1',f'=IF(OR(Z6="running",Z6="cancelled",Z6="failed"),Z6,'
        f'IF(Market_Paths!AL3<>"OK","invalid",IF(Z6<>"complete","missing",'
        f'IF(OR(Z8<>{CONTEXT},Z10<>Risk_Settings!B29,Z11<>Risk_Settings!B30),"stale",'
        f'IFERROR(IF({complete},"current","incomplete"),"incomplete")))))','report_caption')


def configure(m):
    _ranks(m)
    m.cells[SHEET].clear();m.styles[SHEET].clear()
    put=lambda a,v,s='report_money':m.put(SHEET,a,v,s)
    for a,v,s in (
        ('A1','Market balances by age','detail_title'),
        ('A2','Real and nominal percentiles are calculated separately from the captured market paths.','report_caption'),
        ('A3','Back to risk results','report_link'),('D3','Open early-loss tests','report_link'),
        ('F3','Inspect market trials','report_link'),('H3','Open funding examples','report_link'),
        ('K1','Early retirement stock tests','detail_title'),('AF1','Three complete market paths','detail_title'),
        ('AA1','Risk refresh checks','detail_title'),('AA2','Assumption check','report_caption'),
        ('AB2','=Market_Paths!AL3','report_caption'),('AA3','Valid captured paths and stresses','report_caption'),
        ('AB3',f'=COUNTIF({TRIAL}!C{FIRST}:C{LAST},1)','whole'),
        ('AA4','Use Run market risk on Summary to refresh the captured trials.','report_caption'),
        ('BA1','Run field','header'),('BB1','Value','header'),
        ('Z2',SCHEMA,'report_caption'),('Z4',ENGINE,'report_caption'),
        ('Z6','','report_caption'),('Z7','','report_caption'),('Z8','','report_caption'),
        ('Z9','','report_caption'),('Z10','','whole'),('Z11','','whole'),
    ):put(a,v,s)
    _status(put)
    current='$Z$1="current"'
    labels=('Run date','Scenario','End age','Trials per spending policy','Seed','Engine version',
            'Fixed success rate','Flexible success rate','Paths with cuts','Median years with cuts',
            'Spending cut trigger','Cut of flexible budget','Flexible share','Opening year',
            'Market variation multiplier','Early-loss years','Median first shortfall age',
            'Inflation mode','Assumption record','Inflation variation','Inflation persistence',
            'Lowest inflation','Highest inflation')
    meta=('TEXT(Z7,"yyyy-mm-dd hh:mm")','Inputs!B7','Inputs!B53',
          'Risk_Settings!B29','Risk_Settings!B30',f'"{ENGINE}"',
          f'SUM({_trial_range("D")})/Risk_Settings!B29',
          f'SUM({_trial_range("D",True)})/Risk_Settings!B29',
          f'COUNTIF({_trial_range("I",True)},">0")',f'MEDIAN({_trial_range("I",True)})',
          'Risk_Settings!B6','Risk_Settings!B7','Risk_Settings!B8','Inputs!B130','Risk_Settings!B31',
          'MIN(Risk_Settings!B5,MAX(0,Inputs!B53-MAX(Inputs!B8+1,Inputs!B10)+1))',
          f'IF(COUNT({_trial_range("E")})>0,MEDIAN({_trial_range("E")}),"")',
          'IF(Risk_Settings!B32=1,"Variable annual inflation","Expected inflation")',
          CONTEXT,'Risk_Settings!B33','Risk_Settings!B34','Risk_Settings!B35','Risk_Settings!B36')
    for row,(label,value) in enumerate(zip(labels,meta,strict=True),2):
        put(f'BA{row}',label,'report_caption')
        put(f'BB{row}',f'=IF({current},{value},"")','report_caption' if row in (2,3,7,19,20) else 'report_money')
    for col,label in zip('ABCDEFGHI',('Year','Age','Base balance, real','Fixed low, real','Fixed median, real',
                                      'Fixed high, real','Flexible low, real','Flexible median, real','Flexible high, real')):
        put(f'{col}4',label,'header')
    for col,label in zip(('BI','BJ','BK','BL','BM','BN'),('Fixed low, nominal','Fixed median, nominal','Fixed high, nominal',
                                                        'Flexible low, nominal','Flexible median, nominal','Flexible high, nominal')):
        put(f'{col}4',label,'header')
    for row in range(5,131):
        on=f'AND({current},{row-4}<=Inputs!$B$70)'
        for col,expression in (('A',f'Projection!B{row}'),('B',f'Projection!C{row}'),('C',f'Projection!BN{row}')):
            put(f'{col}{row}',f'=IF({on},{expression},"")','whole' if col!='C' else 'report_money')
        for flexible in (False,True):
            for offset,percentile in enumerate((.1,.5,.9)):
                for nominal in (False,True):
                    source=L(ANNUAL_FIRST+row-5+int(nominal)*ANNUAL_CAPACITY)
                    target=L((61 if nominal else 4)+int(flexible)*3+offset)
                    put(f'{target}{row}',f'=IF({on},PERCENTILE({_trial_range(source,flexible)},{percentile}),"")')
    for col,label in zip('KLMNOPQ',('Stock return','Fixed shortfall age','Fixed unpaid, real','Flexible shortfall age',
                                  'Flexible unpaid, real','Fixed shortfall price index','Flexible shortfall price index')):
        put(f'{col}4',label,'header')
    for offset in range(13):
        row=offset+5
        fixed=f'2*Risk_Settings!B29+{offset+1}'
        flex=f'2*Risk_Settings!B29+{offset+14}'
        values=[f'Risk_Settings!B{offset+14}']
        for source,position in (('E',fixed),('M',fixed),('E',flex),('M',flex),('L',fixed),('L',flex)):
            value=_trial_at(source,position)
            values.append(f'IF({value}="","",{value})')
        for col,value in zip('KLMNOPQ',values):put(f'{col}{row}',f'=IF({current},{value},"")')
    for col,label in zip(('AF','AG','AH','AI','AJ','AK','AL','AM','AN','AO','AP','AQ','AR','AS','AT','AU','AV','AW','AX','AY','AZ'),
        ('Path','Ending balance, real','First shortfall age','Unpaid, real','Spending unpaid, real','Tax unpaid, real',
         'Spending due, real','Tax due, real','Income available, real','Income used, real','HSA used, real',
         'HSA-eligible costs, real','Brokerage available, real','Brokerage used, real','Traditional available, real',
         'Traditional used, real','Roth available, real','Roth used, real','Trial','First shortfall price index','Ending price index')):
        put(f'{col}4',label,'header')
    for row,(label,percentile) in enumerate(zip(('Low','Median','High'),(.1,.5,.9)),5):
        position=_rank_position(percentile)
        for col,source in zip(('AG','AH','AI','AJ','AK','AL','AM','AN','AO','AP','AQ','AR','AS','AT','AU','AV','AW','AX','AY','AZ'),
                              ('G','E','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z','AA','A','L','AE')):
            value=f'INDEX({TRIAL}!${source}${FIRST}:${source}${FIRST+CAPACITY-1},{position})'
            put(f'{col}{row}',f'=IF({current},IF({value}="","",{value}),"")')
        put(f'AF{row}',f'=IF({current},"{label}","")','label')
    put('K20','Inflation, longevity and medical tests','detail_title')
    for col,label in zip('KLMNOPQR',('Test','Paid all costs','End age','First shortfall age','Unpaid, real',
                                    'First shortfall price index','Ending balance, real','Ending price index')):
        put(f'{col}22',label,'header')
    for row,label in enumerate(STRESSES,23):
        put(f'K{row}',label,'label')
        for col,source in zip('LMNOPQR',('D','AG','E','M','L','G','AE')):
            value=_trial_at(source,f'2*Risk_Settings!B29+{row+4}')
            put(f'{col}{row}',f'=IF({current},IF({value}="","",{value}),"")')
