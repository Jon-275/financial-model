"""Build a native Excel trial model from the same annual financial compiler."""
import investment_results
from copy import deepcopy
import re
from openpyxl.formula.tokenizer import Tokenizer
from openpyxl.utils.cell import get_column_letter as L
from risk_engine.scenario import ONE_TIME_EXPENSE_ROWS

import native_random
import native_capture
import native_results
import native_funding_paths

TRIAL = 'Risk_Trials'
MARKET = 'Market_Paths'
RANDOM = 'Random_Numbers'
CONTROL = f'{TRIAL}!$B$3'
NATIVE_SHEETS = (TRIAL, MARKET, RANDOM)
MODEL_SHEETS = {
    'Inputs': 'Trial_Inputs', 'Projection': 'Trial_Projection',
    'Cash_Flow': 'Trial_Cash_Flow', 'Roth_Access': 'Trial_Roth_Access',
    'Healthcare': 'Trial_Healthcare', 'Funding_Calc': 'Trial_Funding_Calc',
    'Checks': 'Trial_Checks',
}
BASE_LAYOUT = {trial: base for base,trial in MODEL_SHEETS.items() if base!='Inputs'}


def _rename(value):
    if not isinstance(value, str):
        return value
    if not value.startswith('='):
        for old, new in MODEL_SHEETS.items():
            value=re.sub(r'\b'+re.escape(old)+r'\b',new,value)
        return value
    tokens = Tokenizer(value)
    for token in tokens.items:
        if token.type == 'OPERAND' and token.subtype == 'RANGE' and '!' in token.value:
            sheet, address = token.value.rsplit('!', 1)
            name = sheet.strip("'").replace("''", "'")
            if name in MODEL_SHEETS:
                token.value = MODEL_SHEETS[name] + '!' + address
    return tokens.render()


def _settings(m):
    m.put('Risk_Settings', 'A28', 'Market simulation', 'detail_band')
    settings = (
        (29, 'Paired market paths', 1000, 'Enter 1 to 1,000. Both spending policies use the same paths.', 'input'),
        (30, 'Random seed', 2026, 'A whole number from 0 to 4,294,967,295. Keep it fixed to compare plans.', 'input'),
        (31, 'Market variation multiplier', 1, '1 uses the asset volatility assumptions. 0 repeats the expected returns.', 'decimal_input'),
        (32, 'Vary annual inflation?', 0, '0 = expected inflation only. 1 = annual variation around that expectation.', 'input'),
        (33, 'Long-run inflation variation', .015, 'Illustrative annual standard deviation: 1.5 percentage points.', 'percent_input'),
        (34, 'Inflation persistence', .5, 'The share of the prior year’s inflation difference carried into the next year.', 'percent_input'),
        (35, 'Lowest annual inflation', -.02, 'Lower bound for the variable-inflation option.', 'percent_input'),
        (36, 'Highest annual inflation', .10, 'Upper bound for the variable-inflation option.', 'percent_input'),
    )
    for row,label,value,note,style in settings:
        for column, item, kind in (('A',label,'label'),('B',value,style),('C',note,'report_caption')):
            m.put('Risk_Settings',f'{column}{row}',item,kind)
    m.put('Risk_Settings','A38','Refresh market risk','report_section')
    m.put('Risk_Settings','A39','Use Run market risk on Summary after changing the plan or risk settings.','report_body')
    m.put('Risk_Settings','A40','The success rate depends on these assumptions. It is not a forecast or guarantee.','report_caption')
    m.put('Risk_Settings','A42','Last run status','report_section')
    m.put('Risk_Settings','A43','=IF(Risk_Data!Z6="","No completed workbook run",Risk_Data!Z6)','report_body')
    m.put('Risk_Settings','A44','=IF(Market_Paths!AL3<>"OK",Market_Paths!AL3,IF(Risk_Data!Z9="","",Risk_Data!Z9))','report_caption')


def _clone_annual(m, baseline):
    """Instantiate the compiled financial formulas; do not maintain a second policy."""
    for source, target in MODEL_SHEETS.items():
        if source == 'Inputs':
            for row in range(4,134):
                if row in investment_results.RETIRED_ROWS:
                    continue
                m.put(target,f'A{row}',f'=Inputs!A{row}','label')
                m.put(target,f'B{row}',f'=IF(ISBLANK(Inputs!B{row}),"",Inputs!B{row})','money')
            for column in 'ABC':
                m.put(target,f'{column}74',f'=Inputs!{column}74','header')
            for row in ONE_TIME_EXPENSE_ROWS:
                for column, style in (('A','label'),('B','whole'),('C','report_money')):
                    address = f'{column}{row}'
                    m.put(target, address, f'=IF(ISBLANK(Inputs!{address}),"",Inputs!{address})', style)
            continue
        inherited = {}
        if source in baseline.sheetnames:
            inherited = {c.coordinate:c.value for row in baseline[source].iter_rows()
                         for c in row if c.row<=130 and c.column<=84 and c.value is not None}
        values = {**inherited, **deepcopy(m.cells[source])}
        for address, value in values.items():
            row = int(''.join(c for c in address if c.isdigit()))
            style = m.styles[source].get(address,
                'header' if row==4 else 'report_caption' if row in (2,3) else 'money')
            m.put(target,address,_rename(value),style)
        m.shared[target] = deepcopy(m.shared[source])
    trial_inputs = MODEL_SHEETS['Inputs']
    m.put(trial_inputs,'A1','Market trial assumptions','detail_title')
    m.put(trial_inputs,'A2','Values link to Inputs. Trial overrides appear in rows 14, 53 and 70.','report_caption')
    m.put(trial_inputs,'A3','Choose a trial on Risk_Trials to inspect its annual ledger.','report_caption')
    m.put(trial_inputs,'B14',f'=IF(Inputs!B130=2026,{MARKET}!F7,Inputs!B14)','percent_active')
    m.put(trial_inputs,'A14','2026–2027 price growth for published tax limits','label')
    m.put(trial_inputs,'A53','End age used by this trial','label')
    m.put(trial_inputs,'A70','Rows used by the trial model','label')
    m.put(trial_inputs,'B53',f'={MARKET}!F3','whole')
    m.put(trial_inputs,'B70',f'=IF(AND({CONTROL}>0,{MARKET}!J3="OK",Inputs!B70>1),{MARKET}!D3,0)','whole')


def _calendar_and_returns(m):
    put = lambda a,v,s=None:m.put(MARKET,a,v,s)
    put('A1','Market path and annual prices','detail_title')
    put('A2','Inspect the selected trial. Fixed and flexible policies share its returns and inflation.','report_caption')
    for address,label,value in (
        ('A3','Trial',f'={CONTROL}'),('C3','Modeled years',f'=IF(OR({CONTROL}<=0,Inputs!B70=0),0,F3-Inputs!B8+1)'),
        ('E3','End age',f'=MAX(Inputs!B53,IF({CONTROL}=2203,100,IF(OR({CONTROL}=2204,{CONTROL}=2205),105,0)))'),
        ('G3','Flexible spending',f'=OR(AND({CONTROL}>=1001,{CONTROL}<=2000),AND({CONTROL}>=2101,{CONTROL}<=2113))'),
        ('I3','Input check','=IF(IF(ISNUMBER(B3),AND(B3=INT(B3),OR(B3=0,B3=-1,AND(B3>=1,B3<=Risk_Settings!B29),AND(B3>=1001,B3<=1000+Risk_Settings!B29),AND(B3>=2001,B3<=2013),AND(B3>=2101,B3<=2113),AND(B3>=2201,B3<=2207))),FALSE()),AL3,"Enter a listed trial number")'),
    ):
        col = ord(address[0])-64
        put(address,label,'label');put(f'{L(col+1)}3',value,'input_active' if address=='A3' else 'whole')
    put('AK3','Risk assumptions','label');put('AL3',_validation(),'report_caption')
    headers = ('Year','Primary age','Spouse age','Stage','Inflation shock','Annual inflation','Price index',
               'Medical growth','Medical price index','Tax index',
               'Independent normal 1','Independent normal 2','Independent normal 3',
               'Independent normal 4','Independent normal 5','Independent normal 6',
               'Correlated normal 1','Correlated normal 2','Correlated normal 3','Correlated normal 4','Correlated normal 5',
               'US stock return','International stock return','Bond return','Cash return','Property return',
               'Taxable account return','Traditional account return','Roth account return','HSA return',
               'Living-cost cut','Cut in today’s dollars','Benefit price index')
    for index,title in enumerate(headers,1):put(f'{L(index)}5',title,'header')
    _correlation(m)
    for row in range(6,132):
        period = row-6
        on = f'AND({period}<$D$3,$J$3="OK")'
        def emit(column,expression,style='money'):
            put(f'{column}{row}',f'=IF({on},{expression},"")',style)
        emit('A',f'Inputs!$B$130+{period}','whole')
        emit('B',f'Inputs!$B$8+{period}','whole')
        emit('C',f'Inputs!$B$9+{period}','whole')
        emit('D',f'IF(B{row}<Inputs!$B$10,"Working","Retirement")','label')
        random = f'AND($B$3>=1,$B$3<=2000)'
        mean = 'Inputs!$B$14'
        stress_mean = f'IF($B$3=2201,0.04,IF(OR($B$3=2202,$B$3=2205),0.05,{mean}))'
        if period==0:
            for col,expr in (('E','0'),('F',mean),('G','1'),('H','IF(ISNUMBER(Inputs!$B$132),Inputs!$B$132,Inputs!$B$14)'),('I','1'),('J','(1+Inputs!$B$14)^(Inputs!$B$130-2026)')):
                emit(col,expr,'percent_active' if col in ('F','H') else 'decimal')
        else:
            emit('E',f'IF(AND({random},Risk_Settings!$B$32=1),P{row},0)','decimal')
            variable = (f'MIN(Risk_Settings!$B$36,MAX(Risk_Settings!$B$35,{mean}+'
                        f'Risk_Settings!$B$34*(F{row-1}-{mean})+'
                        f'Risk_Settings!$B$33*SQRT(1-Risk_Settings!$B$34^2)*E{row}))')
            emit('F',f'IF(AND({random},Risk_Settings!$B$32=1),{variable},{stress_mean})','percent_active')
            variable_prices=f'AND({random},Risk_Settings!$B$32=1)'
            emit('G',f'IF({variable_prices},G{row-1}*(1+F{row}),(1+F{row})^{period})','decimal')
            medical_mean='IF(ISNUMBER(Inputs!$B$132),Inputs!$B$132,Inputs!$B$14)'
            emit('H',f'IF(F{row}=Inputs!$B$14,{medical_mean},(1+{medical_mean})*(1+F{row})/(1+Inputs!$B$14)-1)','percent_active')
            emit('I',f'IF({variable_prices},I{row-1}*(1+H{row}),(1+H{row})^{period})','decimal')
            emit('J',f'IF(AND(NOT({variable_prices}),F{row}=Inputs!$B$14),(1+Inputs!$B$14)^(A{row}-2026),G{row}*(1+Inputs!$B$14)^(Inputs!$B$130-2026))','decimal')
        for number in range(6):
            if period==0:expr='0'
            else:
                pair=number//2
                first=5+(period-1)*6+pair*2
                radius=f'SQRT(-2*LN({RANDOM}!$I${first}))'
                angle=f'2*PI()*{RANDOM}!$I${first+1}'
                expr=f'{radius}*{"COS" if number%2==0 else "SIN"}({angle})'
            emit(L(11+number),expr,'decimal')
        for asset in range(5):
            terms='+'.join(f'${L(41+j)}${8+asset}*{L(11+j)}{row}' for j in range(asset+1))
            emit(L(17+asset),terms,'decimal')
            ar=asset+6
            base=f'Asset_Allocation!$B${ar}'
            sampled=f'(1+{base})*EXP($AS${18+asset}+$AR${18+asset}*{L(17+asset)}{row})-1'
            first_age='MAX(Inputs!$B$8+1,Inputs!$B$10)'
            early=f'AND(OR(AND($B$3>=2001,$B$3<=2013),AND($B$3>=2101,$B$3<=2113)),B{row}>={first_age},B{row}<{first_age}+Risk_Settings!$B$5)'
            loss=f'INDEX(Risk_Settings!$B$14:$B$26,MOD($B$3-1,100)+1)'
            expected=f'MAX(-1,{base}-IF($B$3=2205,0.02,0))'
            expr=f'IF({random},{sampled},IF({early},{loss},{expected}))' if asset<2 else f'IF({random},{sampled},{expected})'
            emit(L(22+asset),expr,'percent_active')
        for account,(working,retired) in enumerate(((17,18),(19,20),(21,22),(134,135))):
            mixcol=L(account+2)
            terms=[]; expected=[]
            for asset in range(5):
                profile=f'IF(D{row}="Working",Inputs!$B$15,Inputs!$B$16)'
                weight=f'INDEX(Asset_Allocation!${mixcol}$26:${mixcol}$48,{asset+1}+9*({profile}-1))'
                terms.append(f'({weight})*(1+{L(22+asset)}{row})')
                expected.append(f'({weight})*(1+Asset_Allocation!$B${asset+6})')
            nominal=f'IF(D{row}="Working",{investment_results.reference(working)},{investment_results.reference(retired)})'
            sample=f'(({"+".join(terms)})/({"+".join(expected)}))*(1+{nominal})-1'
            emit(L(27+account),f'IF(AND({random},Risk_Settings!$B$31=0),{nominal},{sample})','percent_active')
        # The cut formula is bound after the trial engine has been instantiated.
        emit('AE','0','money');emit('AF',f'AE{row}/G{row}','money')
        claim_start=f'MAX(1,MIN({period+1},Inputs!$B$45-Inputs!$B$8+1))'
        emit('AG',f'MAX(INDEX($G$6:G{row},{claim_start}):G{row})','decimal')


def _correlation(m):
    put = lambda a,v,s=None:m.put(MARKET,a,v,s)
    put('AO6','Correlation calculation','detail_title')
    put('AO15','Annual asset-return parameters','detail_title')
    for i in range(5):
        for j in range(5):
            a_source=f'Asset_Allocation!{L(2+j)}{16+i}'
            b_source=f'Asset_Allocation!{L(2+i)}{16+j}'
            comparison=(f'IF(AND(ISNUMBER({a_source}),ISNUMBER({b_source})),'
                        f'--AND({a_source}>=-1,{a_source}<=1,ABS({a_source}-{b_source})<=0.000000000001'
                        f'{","+a_source+"=1" if i==j else ""}),0)')
            put(f'{L(47+j)}{8+i}','='+comparison,'whole')
            a=f'{L(41+j)}{8+i}'
            if j>i:put(a,0,'decimal');continue
            source=f'Asset_Allocation!{L(2+j)}{16+i}'
            if j:
                left=f'$AO${8+i}:${L(40+j)}${8+i}'
                right=f'$AO${8+j}:${L(40+j)}${8+j}'
                residual=f'{source}-SUMPRODUCT({left},{right})'
                valid=f'AND(COUNT({left})={j},COUNT({right})={j})'
            else:residual=source;valid='TRUE()'
            if i==j:formula=f'IF({valid},IF(({residual})>=-0.000000000001,SQRT(MAX(0,{residual})),"Check correlation"),"Check correlation")'
            else:
                diagonal=f'${L(41+j)}${8+j}'
                formula=f'IF(AND({valid},ISNUMBER({diagonal})),IF({diagonal}>0,({residual})/{diagonal},IF(ABS({residual})<=0.000000000001,0,"Check correlation")),"Check correlation")'
            put(a,'='+formula,'decimal')
        row=18+i;mean=f'Asset_Allocation!$B${6+i}';sd=f'Asset_Allocation!$C${6+i}'
        variance=f'({sd}/(1+{mean})*Risk_Settings!$B$31)^2'
        logvar=f'LN(1+{variance}/IF(Risk_Settings!$B$9=1,0.5+SQRT(0.25+{variance}),1))'
        for col,label in (('AO','Asset'),('AP','Expected return'),('AQ','Annual volatility'),('AR','Log scale'),('AS','Log shift')):put(f'{col}17',label,'header')
        put(f'AO{row}',f'=Asset_Allocation!A{6+i}','label')
        put(f'AP{row}','='+mean,'percent_active');put(f'AQ{row}','='+sd,'percent_active')
        valid=f'IF(AND(ISNUMBER({mean}),ISNUMBER({sd}),ISNUMBER(Risk_Settings!$B$31)),AND({mean}>-1,{sd}>=0,Risk_Settings!$B$31>=0),FALSE())'
        put(f'AR{row}',f'=IF({valid},SQRT({logvar}),"Review asset returns")','decimal')
        put(f'AS{row}',f'=IF({valid},IF(Risk_Settings!$B$9=1,0,-({logvar})/2),"Review asset returns")','decimal')


def _validation():
    number = lambda row,condition:f'IF(ISNUMBER(Risk_Settings!B{row}),{condition},FALSE())'
    checks = [
        number(5,'AND(Risk_Settings!B5=INT(Risk_Settings!B5),Risk_Settings!B5>=0,Risk_Settings!B5<=125)'),
        *(number(row,f'AND(Risk_Settings!B{row}>=0,Risk_Settings!B{row}<=1)') for row in (6,7,8)),
        number(29,'AND(Risk_Settings!B29=INT(Risk_Settings!B29),Risk_Settings!B29>=1,Risk_Settings!B29<=1000)'),
        number(30,f'AND(Risk_Settings!B30=INT(Risk_Settings!B30),Risk_Settings!B30>=0,Risk_Settings!B30<={native_random.MASK})'),
        number(31,'Risk_Settings!B31>=0'),
        'OR(Risk_Settings!B9=1,Risk_Settings!B9=2)',
        'OR(Risk_Settings!B32=0,Risk_Settings!B32=1)',
        'COUNT($AO$8:$AS$12)=25','SUM($AU$8:$AY$12)=25','COUNT($AR$18:$AS$22)=10',
        'AND(COUNT(Risk_Settings!B14:B26)=13,MIN(Risk_Settings!B14:B26)>=-1)',
        'IF(Risk_Settings!B32=1,AND(COUNT(Risk_Settings!B33:B36)=4,Risk_Settings!B33>=0,'
        'Risk_Settings!B34>=0,Risk_Settings!B34<1,Risk_Settings!B35>-1,'
        'Risk_Settings!B35<=Inputs!B14,Risk_Settings!B36>=Inputs!B14),TRUE())',
    ]
    retirement_in_tests = 'MAX(Inputs!B53,105)>=MAX(Inputs!B10,Inputs!B8+1)'
    missing_budget = 'COUNT(INDEX(Inputs!D37:F37,1,Inputs!B4))<>1'
    return ('=IF(Inputs!B70<2,"Extend or correct the plan",'
            f'IFERROR(IF(AND({retirement_in_tests},{missing_budget}),'
            '"Enter retirement spending on Inputs for the longer market tests.",'
            f'IF(AND({",".join(checks)}),"OK","Review risk assumptions")),"Review risk assumptions"))')


def _bind_annual(m):
    projection = MODEL_SHEETS['Projection']
    healthcare = MODEL_SHEETS['Healthcare']
    for row in range(5,131):
        market_row=row+1
        active=f'{row}-4<=Trial_Inputs!$B$70'
        m.put(projection,f'G{row}',f'=IF({active},{MARKET}!G{market_row},"")','decimal')
        m.put(projection,f'H{row}',f'=IF({active},{MARKET}!J{market_row},"")','decimal')
        m.put(projection,f'CG{row}',f'=IF({active},{MARKET}!AG{market_row},"")','decimal')
        if row==5:continue
        replacements={}
        for working, retired, market_col in ((17,18,'AA'),(19,20,'AB'),(21,22,'AC'),(134,135,'AD')):
            for phase in (f'E{row}',f'{projection}!$E${row}'):
                original=f'IF({phase}="Working",{investment_results.reference(working)},{investment_results.reference(retired)})'
                replacements[original]=f'{MARKET}!${market_col}${market_row}'
        medical=(f'(1+IF(ISNUMBER(Trial_Inputs!$B$132),Trial_Inputs!$B$132,'
                 f'Trial_Inputs!$B$14))^{projection}!$A${row}')
        replacements[medical]=f'{MARKET}!$I${market_row}'
        for target in MODEL_SHEETS.values():
            for address, formula in tuple(m.cells[target].items()):
                if not address.endswith(str(row)) or not isinstance(formula,str) or not formula.startswith('='):
                    continue
                # Exact expressions are compiler-owned. References in other
                # annual rows are left intact rather than rewritten by substring.
                if int(''.join(c for c in address if c.isdigit())) != row:continue
                for before,after in replacements.items():formula=formula.replace(before,after)
                m.cells[target][address]=formula
        original=m.cells[projection][f'AW{row}']
        m.put(projection,f'CH{row}',original,'money')
        m.put(projection,f'AW{row}',f'=IF({active},CH{row}-{MARKET}!AE{market_row},"")','money')
        fixed=(f'Trial_Inputs!$B$38*IF(B{row}<Trial_Inputs!$B$39,1,'
               f'IF(B{row}=Trial_Inputs!$B$39,Trial_Inputs!$B$40/12,0))*'
               f'IF(F{row}="Single",Trial_Inputs!$B$56,1)')
        cut=(f'IF($H$3,IF(AND(D{market_row}<>"Working",{projection}!BH{row-1}<'
             f'Risk_Settings!$B$6*Projection!BH{row-1}),'
             f'MAX(0,{projection}!CH{row}-({fixed.replace("B"+str(row),projection+"!B"+str(row)).replace("F"+str(row),projection+"!F"+str(row))}))'
             f'*Risk_Settings!$B$7*Risk_Settings!$B$8,0),0)')
        m.put(MARKET,f'AE{market_row}',f'=IF(AND({market_row}-6<$D$3,$J$3="OK"),{cut},"")','money')
        # Additional medical reserves are deliberately outside HSA eligibility.
        m.put(healthcare,f'AG{row}',m.cells[healthcare][f'P{row}'],'money')
        reserve=(f'IF(AND(OR({CONTROL}=2206,{CONTROL}=2207),{projection}!E{row}<>"Working"),'
                 f'MAX(0,IF({CONTROL}=2206,18000,25000)*{MARKET}!I{market_row}*'
                 f'IF({projection}!F{row}="Single",Trial_Inputs!B57,1)-(AG{row}-K{row})),0)')
        m.put(healthcare,f'AH{row}',f'=IF({active},{reserve},"")','money')
        m.put(healthcare,f'P{row}',f'=IF({active},AG{row}+AH{row},"")','money')
        tokens=Tokenizer(m.cells[healthcare][f'Q{row}'])
        for token in tokens.items:
            if token.type=='OPERAND' and token.subtype=='RANGE' and token.value==f'P{row}':token.value=f'AG{row}'
        m.cells[healthcare][f'Q{row}']=tokens.render()
    m.put(projection,'CH4','Living costs before temporary cuts','header')
    m.put(healthcare,'AG4','Gross cost before stress reserve','header')
    m.put(healthcare,'AH4','Additional medical cash reserve','header')
    for source,target in MODEL_SHEETS.items():
        if source=='Inputs':continue
        if source=='Funding_Calc':
            m.put(target,'AR1','Market trial withdrawal calculation','detail_title')
        else:
            m.put(target,'A1',f'Market trial — {source.replace("_"," ").lower()}','detail_title')
            m.put(target,'A2','Choose a trial on Risk_Trials. Amounts use that trial’s future dollars.','report_caption')


def configure(m, baseline):
    _settings(m)
    _clone_annual(m,baseline)
    _calendar_and_returns(m)
    _bind_annual(m)
    native_random.configure(m.put)
    m.put(TRIAL,'A1','Market trial results','detail_title')
    m.put(TRIAL,'A2','Select a trial number to inspect its cash flow and account balances. Use 0 for the main plan.','report_caption')
    m.put(TRIAL,'A3','Trial to inspect','label');m.put(TRIAL,'B3',0,'input')
    m.put(TRIAL,'A4','Row 7 shows the selected trial. Completed runs start at row 8: fixed spending, flexible spending, then stress tests.','report_caption')
    native_capture.configure(m,baseline)
    native_funding_paths.configure(m)
    native_results.configure(m)
    m.MODULES.extend((*NATIVE_SHEETS,*MODEL_SHEETS.values()))
