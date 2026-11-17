"""A bounded, discontinuity-aware annual withdrawal solver for Excel.

Probe both conversion policies, choose the earliest feasible cash interval,
then refine to a cent. Every evaluated draw has its own tax and ACA credit.
"""
import re

from openpyxl.formula.tokenizer import Tokenizer
from openpyxl.utils.cell import column_index_from_string, coordinate_from_string, get_column_letter as L

POINTS = 9
CLIFF_PROBES = 27
GRID = POINTS + CLIFF_PROBES
ITERATIONS = 49
ROWS = 2 * GRID + 1 + ITERATIONS + 1

BASE_COLUMNS = {'V':'E','W':'F','Y':'G','AA':'H','AE':'I','AD':'J','AH':'K','AI':'L','AJ':'M','AK':'N',
                'AM':'O','AN':'P','AO':'Q','AP':'R','AQ':'S','AS':'T','AT':'U','AV':'V',
                'FundingMargin':'W','AcaMagi':'X','PremiumCredit':'Y','FundingStatus':'Z'}
BASE_HEADERS = ['Requested draw','Lower draw','Upper draw','Allow conversion','Taxable draw','Traditional draw',
                'Roth draw','Conversion','Realized gain or loss','Ordinary before Social Security','Provisional income',
                'Taxable Social Security','Ordinary income','Adjusted gross income','Ordinary taxable','Federal ordinary',
                'Federal gains','Investment income tax','AMT','California taxable','California tax','Current tax',
                'Funding margin','ACA MAGI','Premium credit','Funding status']

def columns(tax):
    return {**BASE_COLUMNS,**{key:L(i) for i,key in enumerate(tax.EXTRA_HEADERS,27)}}


def detail_column(column):
    return L(column_index_from_string(column) + 3)


def detail_address(address):
    column, row = coordinate_from_string(address)
    return f'{detail_column(column)}{row}'


def _detail_formula(formula):
    """Shift local references past the worksheet’s audit columns."""
    tokens = Tokenizer('=' + formula)
    for token in tokens.items:
        if token.subtype == 'RANGE' and '!' not in token.value:
            token.value = re.sub(r'([A-Z]+)(?=\$?\d+)', lambda match: detail_column(match[1]), token.value)
    return tokens.render()


def headers(m,sheet,tax):
    m.row(sheet, 1, ['Year', 'Age', 'Calculation step', *BASE_HEADERS, *tax.EXTRA_HEADERS.values(),
                    'Candidate point or lower bound', 'Best funding margin', 'Best available draw'], 'header')

def emit(m,sheet,first,c,request,tax,health):
    col=columns(tax)
    p,prev,cash,h=(c[k] for k in ('p','prev','cash','health'))
    enabled=c['active']; cap=cash('K'); mandatory=p('T'); trad_limit=cash('I'); broker_cap=cash('H')
    room=f'MAX(0,{trad_limit}-{mandatory})'
    pref=f'IF(Inputs!$B$48=2,MIN({room},{p("U")}),0)'
    # Use the exact traditional balance when the preferred draw exhausts it.
    # Re-adding a mandatory draw to the remaining room can change its last bit.
    preferred_limit=f'IF({pref}>={room},{trad_limit},{mandatory}+{pref})'
    aca=f'AND({h("M")}>0,{h("N")}>0,{h("O")}>0)'
    exact_conversion=request is not None
    req=request if exact_conversion else f'Inputs!$B$47*{p("G")}'
    allow=(f'AND({p("E")}<>"Working",{p("R")}=0)' if exact_conversion else
           f'AND({p("E")}<>"Working",{p("R")}=0,IF(Inputs!$B$97=1,{p("C")},{p("D")})<=Inputs!$B$50)')
    control=first+2*GRID
    search=control+1; final=first+ROWS-1
    scratch=L(27+len(tax.EXTRA_HEADERS))
    best=L(28+len(tax.EXTRA_HEADERS))
    scratch_cell=lambda t:f'${scratch}${t}'
    # The first grid always has these source values, including during a
    # no-conversion retry. Both policies use the same withdrawal boundaries.
    broker_lo = scratch_cell(first+1)
    broker_hi = scratch_cell(first+2)
    trad_hi = scratch_cell(first+3)
    def put(addr,value):
        m.put(sheet, detail_address(addr), m.source_cache.localize(sheet, _detail_formula(value)))
    def on(guard,formula): return f'IF({guard},{formula},"")'
    def at(letter,t): return f'${letter}${t}'
    def shared(letter,lo,hi,formula):
        put(f'{letter}{lo}',formula)
        if hi>lo:
            column = detail_column(letter)
            m.shared[sheet][f'{column}{lo}'] = f'{column}{lo}:{column}{hi}'

    for column, value in zip('AB', (p('B'), p('C'))):
        m.put(sheet, f'{column}{first}', _detail_formula(on(enabled, value)), 'whole')
        following = f'{column}{first + 1}'
        m.put(sheet, following, f'=IF(ISNUMBER($A${first}),${column}${first},"")', 'whole')
        m.shared[sheet][following] = f'{following}:{column}{final}'

    grid_end = first + GRID - 1
    for start, stop, label in (
        (first, grid_end, 'Test with conversion'),
        (grid_end + 1, control - 1, 'Test without conversion'),
        (control, control, 'Choose search range'),
        (search, final - 1, 'Refine draw'),
        (final, final, 'Selected draw'),
    ):
        row = f'C{start}'
        m.put(sheet, row, f'=IF(ISNUMBER($A{start}),IF(ISNUMBER($D{start}),"{label}","Skipped"),"")', 'label')
        if stop > start:
            m.shared[sheet][row] = f'{row}:C{stop}'

    if request is not None:
        m.put(sheet, 'AP1', 'Conversion request', 'header')
        m.put(sheet, f'AP{first}', f'=IF(ISNUMBER($A{first}),{request},"")', 'cents')
        following = f'AP{first + 1}'
        m.put(sheet, following, f'=IF(ISNUMBER($A${first}),$AP${first},"")', 'cents')
        m.shared[sheet][following] = f'{following}:AP{final}'

    # Structural nodes partition gross draw into traditional, brokerage and
    # Roth sections, and partition the capped conversion and loss formulas.
    def draw_for_trad(amount):
        w=f'MAX({mandatory},MIN({trad_limit},{amount}))'
        return f'IF({w}<={broker_lo},{w},{w}+{broker_cap})'
    base_points=['0',preferred_limit,f'{broker_lo}+{broker_cap}',f'{trad_limit}+{broker_cap}',cap,
                 draw_for_trad(f'{prev("BE")}-{req}'),draw_for_trad(f'{mandatory}+{p("U")}-{req}'),
                 draw_for_trad(f'{mandatory}+{p("U")}'),
                 f'IF({p("AB")}=1,{broker_lo},{broker_lo}+MAX(0,MIN({broker_cap},({c["fed_loss"]}-Tax_Tables!$B$190)/(1-{p("AB")}))))']

    def min_feasible(lo,hi):
        options=','.join(f'IF(AND(ISNUMBER(W{t}),W{t}>=0),A{t},{cap}+1)' for t in range(lo,hi+1))
        return f'MIN({options})'

    grid_1_end=first+GRID-1
    mode1_ok=f'{at("A",control)}<={cap}'
    for mode,g in ((1,first),(0,first+GRID)):
        guard=enabled if mode==1 else f'AND({enabled},NOT({mode1_ok}))'
        if exact_conversion and mode==0:
            # A request is tested as a policy, not silently replaced with zero.
            guard='FALSE()'
        # Stable sort nodes; default legacy coverage evaluates just zero and capacity.
        for j,node in enumerate(base_points): put(f'{scratch}{g+j}',on(guard,node))
        for j in range(POINTS):
            run=guard if j in (0,POINTS-1) else f'AND({guard},{aca})'
            put(f'A{g+j}',on(run,f'SMALL({scratch_cell(g)}:{scratch_cell(g+POINTS-1)},{j+1})'))
        phase_ranges=[('0',broker_lo),(broker_lo,broker_hi),(broker_hi,cap)]
        t=g+POINTS
        for low,high in phase_ranges:
            for ratio in ('1','1.33','4'):
                cutoff=f'({h("O")}*{ratio})'; terms=[]
                for j in range(POINTS-1):
                    a,b=at('A',g+j),at('A',g+j+1); ml,mh=at('X',g+j),at('X',g+j+1)
                    crossed=(f'OR(AND({ml}<={cutoff},{mh}>{cutoff}),AND({mh}<={cutoff},{ml}>{cutoff}))' if ratio=='4'
                             else f'OR(AND({ml}<{cutoff},{mh}>={cutoff}),AND({mh}<{cutoff},{ml}>={cutoff}))')
                    terms.append(f'IF(AND({b}>{a},{a}>={low},{b}<={high},{crossed}),{a}+({cutoff}-{ml})*({b}-{a})/({mh}-{ml}),{cap}+1)')
                put(f'{scratch}{t}',on(f'AND({guard},{aca})',f'MIN({",".join(terms)})'))
                cross=scratch_cell(t)
                eps=f'MAX(0.00001,ABS({cross})*0.000000000000002)'
                for k,x in enumerate((f'MAX({low},{cross}-{eps})',cross,f'MIN({high},{cross}+{eps})')):
                    put(f'A{t+k}',on(f'AND(ISNUMBER({cross}),{cross}<={cap})',x))
                t+=3
        assert t==g+GRID
        shared('D',g,g+GRID-1,on('ISNUMBER($A'+str(g)+')',str(mode)))

    put(f'A{control}',on(enabled,min_feasible(first,grid_1_end)))
    put(f'B{control}',on(enabled,min_feasible(first+GRID,control-1)))
    if exact_conversion:
        put(f'D{control}',on(enabled,'1'))
        chosen=f'A{control}'
    else:
        put(f'D{control}',on(enabled,f'IF({mode1_ok},1,0)'))
        chosen=f'MIN(A{control},B{control})'
    put(f'C{control}',on(enabled,chosen))
    feasible=f'{at("C",control)}<={cap}'
    initial_lo=','.join(f'IF(AND(ISNUMBER(A{t}),D{t}={at("D",control)},A{t}<{at("C",control)}),A{t},0)' for t in range(first,control))
    put(f'{scratch}{control}',on(enabled,f'IF({feasible},MAX({initial_lo}),0)'))
    best_margin=(f'IF({at("D",control)}=1,MAX(W{first}:W{grid_1_end}),'
                 f'MAX(W{first+GRID}:W{control-1}))')
    put(f'{best}{control}',on(f'AND({enabled},NOT({feasible}))',best_margin))
    best_draw=','.join(f'IF(AND(ISNUMBER(W{t}),D{t}={at("D",control)},W{t}>={at(best,control)}-0.0000001),A{t},{cap}+1)' for t in range(first,control))
    short_draw=L(29+len(tax.EXTRA_HEADERS))
    put(f'{short_draw}{control}',on(f'AND({enabled},NOT({feasible}))',f'MIN({best_draw})'))
    lo0=scratch_cell(control); hi0=at('C',control)
    search_needed=f'AND({enabled},{feasible},{hi0}>{lo0})'
    put(f'B{search}',on(search_needed,lo0)); put(f'C{search}',on(search_needed,hi0))
    t=search+1
    shared('B',t,final-1,on(search_needed,f'IF(W{t-1}<0,A{t-1},B{t-1})'))
    shared('C',t,final-1,on(search_needed,f'IF(W{t-1}>=0,A{t-1},C{t-1})'))
    shared('A',search,final-1,on(search_needed,f'(B{search}+C{search})/2'))
    shared('D',search,final-1,on(f'ISNUMBER(A{search})',at('D',control)))
    put(f'B{final}',on(enabled,f'IF({search_needed},IF(W{final-1}<0,A{final-1},B{final-1}),0)'))
    put(f'C{final}',on(enabled,f'IF({search_needed},IF(W{final-1}>=0,A{final-1},C{final-1}),IF({feasible},{hi0},{at(short_draw,control)}))'))
    put(f'A{final}',on(enabled,f'C{final}'))
    put(f'D{final}',on(enabled,at('D',control)))

    # One shared candidate block. The control row has a numeric A, so exclude
    # it explicitly to prevent its result cells from entering either grid.
    t=first; a=f'A{t}'
    # Each exact breakpoint returns the account's stated capacity. Between
    # breakpoints, only the current account takes a partial withdrawal.
    taxable=f'IF({a}>={broker_hi},{broker_cap},MIN({broker_cap},MAX(0,{a}-{broker_lo})))'
    trad=f'IF({a}>={trad_hi},{trad_limit},IF({a}<={broker_hi},MIN(MAX({a},{mandatory}),{broker_lo}),MIN({trad_limit},MAX({broker_lo},{a}-{broker_cap}))))'
    roth=f'IF({a}>={cap},{cash("J")},MIN({cash("J")},MAX(0,{a}-{trad_hi})))'
    # The comparison varies the request, using the same capped conversion
    # policy and bracket as the projection.
    conversion=f'IF(AND(D{t}=1,{allow}),MIN(MAX(0,{prev("BE")}-F{t}),{req},IF(Inputs!$B$48=2,MAX(0,{p("U")}-MAX(0,F{t}-{mandatory})),{req})),0)'
    def ref(k):
        if k=='OpeningFedLoss': return c['fed_loss']
        if k=='OpeningCaLoss': return c['ca_loss']
        if k in col: return f'{col[k]}{t}'
        return p(k)
    old_ordinary=f'{p("I")}-{p("J")}-{p("K")}+F{t}+H{t}+{p("O")}+{p("AC")}'
    formulas={'V':taxable,'W':trad,'Y':roth,'AA':conversion,'AD':old_ordinary,**tax.candidate_formulas(ref)}
    formulas['AcaMagi']=f'N{t}+{p("N")}-L{t}+{p("AG")}'
    formulas['PremiumCredit']=health.credit_from(f'X{t}',h('M'),h('N'),h('O'),p('B'))
    formulas['FundingMargin']=f'{cash("C")}+SUM(E{t}:G{t})-{cash("D")}+{cash("E")}-{cash("F")}-V{t}+Y{t}'
    probe=f'AND(ISNUMBER(A{t}),ROW()<>{control})'
    for key,formula in formulas.items():
        assert key in col,(key,formula)
        shared(col[key],first,final,on(probe,formula.lstrip('=')))
    put(f'Z{final}',on(enabled,f'IF(AND({feasible},OR(C{final}-B{final}>0.005,W{final}<-0.005),{search_needed}),"Check funding convergence",IF(W{final}<-0.009999,"Unfunded","Funded"))'))
    return final,col
