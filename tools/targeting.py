"""Selected-year Roth conversion comparison using the shared annual funding engine."""

from openpyxl.utils.cell import get_column_letter as col_letter
from funding_search import detail_column

VIEW = 'Income_Targeting'
CALC = 'Targeting_Calc'
DETAIL = 'Conversion Detail'
FIRST_TRIAL = 5
LAST_TRIAL = 55
INTERVALS = LAST_TRIAL - FIRST_TRIAL
BASE_INDEX = '$B$6'
HEALTH_READY = 'OR($F$5="OK",$F$5="Legacy budget")'


def _index(sheet, col, prior=False, selector=BASE_INDEX):
    """Select the current or opening-year observation without volatile lookups."""
    position = f'MAX(1,{selector}-1)' if prior else f'MAX(1,{selector})'
    return f'INDEX({sheet}!${col}$5:${col}$130,{position})'


def _context(m):
    """Cache the common selected-year inputs beside the trial schedule."""
    position = f'{VIEW}!$B$6'
    refs = {}

    def ref(sheet, col, prior=False):
        key = (sheet, col, prior)
        if key not in refs:
            rr = 3 + len(refs)
            address = f'$AV${rr}'  # The solver adds three leading audit columns.
            refs[key] = address
            m.put(CALC, f'AX{rr}', f'{sheet} {col}' + (' (opening)' if prior else ''), 'report_caption')
            m.put(CALC, f'AY{rr}', f'=IF({position}=0,"",{_index(sheet,col,prior,position)})')
        return refs[key]

    m.put(CALC, 'AX1', 'Selected-year inputs', 'report_section')
    m.put(CALC, 'AX2', 'Source', 'header')
    m.put(CALC, 'AY2', 'Value', 'header')
    p = lambda col: ref('Projection', col)
    prior = lambda col: ref('Projection', col, True)
    return {
        'p': p,
        'prev': prior,
        'cash': lambda col: ref('Cash_Flow', 'AB' if col == 'C' else col),
        'health': lambda col: ref('Healthcare', col),
        'active': f'AND({position}>0,OR({VIEW}!$F$5="OK",{VIEW}!$F$5="Legacy budget"))',
        'fed_loss': prior('BW'),
        'ca_loss': prior('BY'),
    }


def _trial(m, start, context, requested):
    active = f'AND({context["active"]},ISNUMBER({requested}),{requested}>=0)'
    final, columns = m.emit_funding_group(CALC, start, {**context, 'active': active}, requested)

    def result(key):
        return f'{CALC}!${detail_column(columns[key])}${final}'

    return result


def configure(m, wb):
    """Replace the frozen sweep, preserving its chart ranges and worksheet."""
    del wb  # All selected-year data is drawn from the compiled workbook.
    m.MODULES.append(DETAIL)

    def put(address, value, style=None):
        m.put(VIEW, address, value, style)

    def label(row, title, help_text=None):
        put(f'A{row}', title, 'label')
        if help_text is not None:
            put(f'C{row}', help_text, 'note')

    def selected(address, expr, style=None):
        put(address, f'=IF($B$6=0,"",{expr})', style)

    p = lambda c: _index('Projection', c)
    p0 = lambda c: _index('Projection', c, True)
    cash = lambda c: _index('Cash_Flow', c)
    health = lambda c: _index('Healthcare', c)

    put('A1', 'Roth conversion comparison')
    put('A2', 'Compares one retirement year. Each trial recalculates draws, tax and ACA credit.')
    put('A4', 'Comparison settings')
    label(5, 'Retirement age to compare', 'Enter a retirement age in the plan. Enter 0 for the first retirement year.')
    put('B5', 0)  # The legacy age 54 was outside the current retirement window.
    label(6, 'Selected row in Projection')
    start = 'MAX(Inputs!$B$8+1,Inputs!$B$10)'
    age = f'IF($B$5=0,{start},$B$5)'
    put('B6', f'=IFERROR(IF(AND(Inputs!$B$70>0,ISNUMBER($B$5),$B$5=INT($B$5),{age}>={start},{age}<=Inputs!$B$53),{age}-Inputs!$B$8+1,0),0)')
    label(7, 'Filing status', 'Uses the active scenario’s filing status for this year.')
    put('B7', f'=IF(Inputs!$B$70=0,"Check Inputs",IF($B$6=0,"Check retirement age",{p("F")}))')
    put('E5', 'Healthcare checks', 'label')
    selected('F5', f'IF(Inputs!$B$125<>"OK","Check healthcare inputs",{health("T")})', 'label')
    rows = {
        8: ('Calendar year', p('B'), 'whole'),
        9: ('Selected age', p('C'), 'whole'),
        10: ('Opening traditional balance', p0('BE'), 'money'),
        11: ('Forced traditional distributions', p('T'), 'money'),
        12: ('Cash income before portfolio draws', cash('AB'), 'money'),
        13: ('Required spending before ACA credit', cash('D'), 'money'),
        14: ('Tax payable from prior year', cash('F'), 'money'),
        15: ('HSA payment for eligible costs', cash('E'), 'money'),
        16: ('Conversion in the active plan', p('AA'), 'money'),
        17: ('Current-year tax in the active plan', p('AV'), 'money'),
        18: ('Eligible Marketplace benchmark premium', health('N'), 'money'),
        19: ('Eligible actual Marketplace premium', health('M'), 'money'),
        20: ('Poverty guideline for the tax household', health('O'), 'money'),
    }
    notes = {
        9:'Primary person’s age. The spouse’s age follows Projection.',
        10:'The opening balance is fixed across all trials.',
        11:'Required distributions are paid before any conversion.',
        12:'Spendable income before the ACA credit.',
        13:'Includes health costs before the ACA credit and HSA payment.',
        14:'The same prior-year amount is due in every trial.',
        15:'The HSA pays eligible expenses once.',
        16:'The comparison does not change the active plan.',
        17:'Excludes tax carried in from the prior year.',
        18:'Uses the eligible quote on Healthcare.',
        19:'Only eligible Marketplace premiums can generate a credit.',
        20:'Zero if no Marketplace policy has a modeled credit.',
    }
    for rr,(title,expr,style) in rows.items():
        label(rr,title,notes.get(rr,''))
        selected(f'B{rr}',expr,style)
    label(21, 'Conversion to test', 'Enter an amount in selected-year dollars, including zero.')
    put('B21', 0, 'input')

    put('A22', 'Active plan and requested conversion')
    for col,title in zip('ABCD',('Selected-year result','Active plan','Requested scenario','Change')):
        put(f'{col}23',title,'header')
    m.funding_headers(CALC)
    context = _context(m)
    custom = _trial(m, 2, context, f'{VIEW}!$B$21')
    put('A31', 'Funding result', 'label')
    selected('B31', f'IF({HEALTH_READY},{cash("L")},"Check healthcare")', 'label')
    put('C31', f'=IF($B$6=0,"",IF(NOT({HEALTH_READY}),"Check healthcare",IF(OR(NOT(ISNUMBER($B$21)),$B$21<0),"Enter a nonnegative amount",{_status(custom,"$B$21")})))', 'label')

    top_rows = {
        24:('Modeled Roth conversion',p('AA'),custom('AA')),
        25:('Current-year tax',p('AV'),custom('AV')),
        26:('ACA premium credit',health('R'),custom('PremiumCredit')),
        27:('Tax and net healthcare',f'{p("AV")}+{health("S")}',f'{custom("AV")}+$B$32-{custom("PremiumCredit")}'),
        28:('Required cash shortfall',p('Z'),f'MAX(0,-{custom("FundingMargin")})'),
        29:('Total portfolio withdrawals',f'SUM({p("V")},{p("W")},{p("Y")})',f'SUM({custom("V")},{custom("W")},{custom("Y")})'),
    }
    for rr,(title,current,trial) in top_rows.items():
        put(f'A{rr}',title,'label')
        selected(f'B{rr}',f'IF({HEALTH_READY},{current},"")','money')
        put(f'C{rr}',f'=IF(OR($B$6=0,NOT({HEALTH_READY}),$C$31="Enter a nonnegative amount",$C$31="Check funding calculation"),"",{trial})','money')
        put(f'D{rr}',f'=IF(AND($B$31="Funded",$C$31="Funded"),C{rr}-B{rr},"")','money')
    put('A30', 'Changes require two fully funded results.')
    put('C30', 'Trials can exceed the plan’s conversion age limit. Account and cash limits still apply.', 'note')
    label(32, 'Gross healthcare before ACA credit', 'Fixed across trials. Future Medicare surcharges are excluded.')
    selected('B32', health('P'), 'money')

    put('A33', 'Conversion range')
    label(34, 'Upper bound before optional draws', 'After mandatory distributions and any bracket limit. Cash needs may lower it.')
    selected('B34', f'IF({p("R")}>0,0,MAX(0,MIN($B$10-$B$11,IF(Inputs!$B$48=2,{p("U")},$B$10))))', 'money')
    label(35, 'Largest requested amount shown', 'The sweep covers the lower of the upper bound and the display limit.')
    put('B35', f'=IF(OR($B$6=0,NOT({HEALTH_READY}),NOT(ISNUMBER($B$36)),$B$36<0),"",ROUNDDOWN(MIN($B$34,$B$36),2))','money')
    label(36, 'Display limit for the sweep', 'Only limits the grid. Enter an amount in selected-year dollars.')
    put('B36', 250000, 'input')
    label(37, 'Approximate spacing between requests', 'The grid includes the largest requested amount.')
    put('B37', f'=IF(NOT(ISNUMBER($B$35)),"",IF($B$35<0.50,$B$35,$B$35/{INTERVALS}))','money')
    label(38, 'Rows funding the requested amount')
    label(39, 'Requests the plan did not fully fund')
    label(40, 'Baseline used for cost changes', 'Changes use a funded $0 request. Rates compare adjacent funded rows.')
    detail_ref = f"'{DETAIL}'!"
    put('B38', f'=IF(NOT(ISNUMBER($B$35)),"",COUNTIF({detail_ref}$C$5:$C$55,"Funded"))', 'whole')
    put('B39', f'=IF(NOT(ISNUMBER($B$35)),"",COUNT({detail_ref}$A$5:$A$55)-B38)', 'whole')
    put('B40', f'=IF(NOT(ISNUMBER($B$35)),"",IF({detail_ref}$C$5="Funded",{detail_ref}$S$5,"Unavailable"))', 'money')
    put('A41', 'Open conversion detail', 'report_link')

    def detail(address, value, style=None):
        m.put(DETAIL, address, value, style)

    detail('A1', 'Roth conversion detail', 'detail_title')
    detail('A2', 'Costs and withdrawals for each request in the selected retirement year. Amounts use that year’s dollars.', 'report_caption')
    detail('A3', 'Back to Roth conversion comparison', 'report_link')
    detail('D3', f'=IF({VIEW}!$B$6=0,"Choose a retirement age on Income_Targeting","Age "&{VIEW}!$B$9&" in "&{VIEW}!$B$8)', 'report_caption')
    headers = [
        'Requested conversion','Modeled conversion','Funding result','Federal AGI','ACA MAGI',
        'Income as % of poverty guideline','Taxable Social Security','Taxable account draw',
        'Traditional draw','Roth draw','ACA premium credit','Net healthcare','Federal regular tax',
        'Alternative minimum tax','Investment income tax','California income tax','Current-year tax',
        'Required cash shortfall','Tax and net healthcare','Cost change vs no conversion',
        'Incremental cost per dollar','Traditional balance before growth',
    ]
    for i,title in enumerate(headers,1): detail(f'{col_letter(i)}4',title,'header')

    p = lambda column: _index('Projection', column, selector=f'{VIEW}!{BASE_INDEX}')
    for rr in range(FIRST_TRIAL,LAST_TRIAL+1):
        step = rr-FIRST_TRIAL
        if step == 0:
            detail(f'A{rr}', f'=IF(AND({VIEW}!$B$6>0,ISNUMBER({VIEW}!$B$35)),0,"")', 'money')
        else:
            small = f'{VIEW}!$B$35' if step == 1 else '""'
            detail(f'A{rr}',f'=IF(OR({VIEW}!$B$6=0,NOT(ISNUMBER({VIEW}!$B$35)),{VIEW}!$B$35=0),"",IF({VIEW}!$B$35<0.50,{small},ROUND({step}*{VIEW}!$B$35/{INTERVALS},2)))','money')
        requested = f'{detail_ref}$A${rr}'
        get = _trial(m,2+(step+1)*m.FUNDING_ROWS,context,requested)
        detail(f'C{rr}',f'=IF(ISNUMBER(A{rr}),{_status(get,f"A{rr}")},"")','label')
        social_security = p('N')
        credit = get('PremiumCredit')
        magi = f'{get("AK")}+{p("AG")}+MAX(0,{social_security}-{get("AI")})'
        net_health = f'{VIEW}!$B$32-{credit}'
        values = {
            'B':get('AA'),'D':get('AK'),'E':magi,'F':f'IF({VIEW}!$B$20>0,E{rr}/{VIEW}!$B$20,"")','G':get('AI'),
            'H':get('V'),'I':get('W'),'J':get('Y'),'K':credit,'L':net_health,
            'M':f'{get("AN")}+{get("AO")}','N':get('AQ'),'O':get('AP'),'P':get('AT'),'Q':get('AV'),
            'R':f'MAX(0,-{get("FundingMargin")})','S':f'Q{rr}+L{rr}',
            'V':f'MAX(0,{VIEW}!$B$10-I{rr}-B{rr})',
        }
        for col,expr in values.items():
            guard = f'C{rr}="Funded"' if col in ('L','S') else f'AND(ISNUMBER(A{rr}),C{rr}<>"Check funding calculation")'
            detail(f'{col}{rr}', f'=IF({guard},{expr},"")', 'percent_active' if col=='F' else 'money')
        detail(f'T{rr}',f'=IF(AND(C{rr}="Funded",$C$5="Funded"),S{rr}-$S$5,"")','money')
        rate = '""' if step==0 else f'IF(AND(C{rr}="Funded",C{rr-1}="Funded"),IF(B{rr}>B{rr-1},(S{rr}-S{rr-1})/(B{rr}-B{rr-1}),""),"")'
        detail(f'U{rr}',f'={rate}','percent_active')


def _status(result, requested):
    code = result('FundingStatus')
    funded = result('AA')
    return (f'IF({code}="Unfunded","Cash shortfall",'
            f'IF({code}<>"Funded","Check funding calculation",'
            f'IF(ABS({funded}-{requested})>0.005,"Request not fully funded","Funded")))')
