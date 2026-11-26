"""Shared tax expressions for annual funding and conversion comparisons.

The compiler assigns result columns; this policy only consumes named references.
It models ordinary W-2 wages, proportional long-term sales, qualified dividends,
the standard deduction, and MFJ/single filing. Unsupported tax features are not
implicitly added to a scenario. See the accompanying tax policy notes.
"""

from collections.abc import Callable

from openpyxl.utils.cell import absolute_coordinate

import published_tax
from state_policy import (
    CALIFORNIA_POLICY, WORKBOOK_CA_POLICY_VERSION_CELL, WORKBOOK_JURISDICTION_ERROR,
    WORKBOOK_JURISDICTION_LABEL, WORKBOOK_JURISDICTION_ROW, WORKBOOK_STATE_POLICY_SCHEMA,
    WORKBOOK_STATE_POLICY_SCHEMA_CELL,
)

EXTRA_HEADERS = {
    "FedCapitalIncome": "Federal capital income in AGI",
    "ClosingFedLoss": "Federal capital loss carryforward",
    "CaCapitalIncome": "California capital income in AGI",
    "ClosingCaLoss": "California capital loss carryforward",
    "AmtTaxable": "AMT taxable after exemption",
    "AmtPreferential": "AMT income eligible for gain rates",
    "AmtTentative": "Tentative minimum tax",
    "CaAgi": "California adjusted gross income",
    "CaExemptionCredit": "California exemption credit",
}

PROJECTION_EXTRA = dict(zip(EXTRA_HEADERS, ('BV','BW','BX','BY','BZ','CA','CB','CC','CD')))

LOSS_LIMIT = 'Tax_Tables!$B$190'
MEDICARE_SINGLE = 'Tax_Tables!$B$191'
CA_CREDIT_MFJ_THRESHOLD = 'Tax_Tables!$B$192'
CA_CREDIT_SINGLE_THRESHOLD = 'Tax_Tables!$B$193'
CA_CREDIT_STEP = 'Tax_Tables!$B$194'
CA_CREDIT_REDUCTION = 'Tax_Tables!$B$195'

_TAX_INPUT_CHECK = ('IFERROR(IF(AND(COUNT(INDEX(D126:F128,0,B4))=3,B126>=0,B126<=1,'
                    'MIN(B127:B128)>=0,OR(B11=1,B11=2),ISNUMBER(B51),B51>=0,OR(B25>0,B51=0)),'
                    '"OK","Check tax inputs"),"Check tax inputs")')


def opening_loss_refs(r):
    """Prior-year unused losses, anchored for a shared candidate formula."""
    if r < 6:
        raise ValueError('The opening position has no tax candidate.')
    return {
        'OpeningFedLoss': f'Projection!$BW${r-1}',
        'OpeningCaLoss': f'Projection!$BY${r-1}',
    }


def capital_agi_expr(realized, opening_loss):
    """Signed capital line in AGI: net gain or capped net capital loss."""
    return f'MAX(-{LOSS_LIMIT},({realized})-({opening_loss}))'


def aca_magi_expr(ordinary_before_ss, qualified_dividends, gross_ss, tax_exempt, realized, opening_fed_loss):
    """Taxable SS plus its ACA add-back equals total gross SS."""
    capital = capital_agi_expr(realized, opening_fed_loss)
    return f'SUM(({ordinary_before_ss}),({qualified_dividends}),({gross_ss}),({tax_exempt}),{capital})'


def _tier_tax(taxable, index, lo, hi, year):
    base = 2026 if lo in (30, 160) else 2027
    slopes = f'Tax_Tables!$D${lo}:$D${hi}'

    def tax_on_floors(column, factor):
        floors = f'(Tax_Tables!${column}${lo}:${column}${hi}*({factor}))'
        return f'SUMPRODUCT(--(({taxable})>{floors}),(({taxable})-{floors}),{slopes})'

    published_index = f'({index})/(1+Inputs!$B$14)' if base == 2027 else index
    # Excel's legacy array evaluation can reduce an IF inside SUMPRODUCT
    # to a scalar. Select the tax total after each SUMPRODUCT instead.
    return f'IF({published_tax.use(year, base)},{tax_on_floors("E", published_index)},{tax_on_floors("C", index)})'


def _bracket_tax(taxable, single, index, single_rows, joint_rows, year):
    return f'IF({single},{_tier_tax(taxable,index,*single_rows,year)},{_tier_tax(taxable,index,*joint_rows,year)})'


def _gain_tax(ordinary, eligible, single, index, year):
    stacked = _bracket_tax(f'{ordinary}+{eligible}',single,index,(154,156),(17,19),year)
    ordinary_band = _bracket_tax(ordinary,single,index,(154,156),(17,19),year)
    return f'({stacked}-{ordinary_band})'


def _amt_rates(amount, index, year):
    upper = published_tax.indexed('Tax_Tables!$B$85','Tax_Tables!$B$239',year,index)
    return f'(Tax_Tables!$B$83*MIN({amount},{upper})+Tax_Tables!$B$84*MAX(0,{amount}-{upper}))'


def _closing_loss(realized, opening, capital_income, agi, deduction):
    """Schedule D carryover worksheet: don't consume loss that sheltering from the standard deduction made unused."""
    # Loss claimed is -capital_income when negative. The carryover worksheet
    # adds that claim back to signed taxable income before capping usage.
    claim = f'MAX(0,-{capital_income})'
    used = f'MIN({claim},MAX(0,{agi}-{deduction}+{claim}))'
    return f'MAX(0,{opening}-{realized}-{used})'


def candidate_formulas(ref):
    """Return formulas without '='. Caller resolves every key to a cell/expression.

    Existing Projection codes identify inputs and outputs. The extra keys are
    listed in EXTRA_HEADERS; OpeningFedLoss and OpeningCaLoss are input refs.
    AD is ordinary income before Social Security and before any capital loss.
    """
    def p(key):
        return f'({ref(key)})'

    single = f'{p("F")}="Single"'
    idx = p('H')
    year = p('B')
    regular_base = p('AM')
    fed_std = p('AL')
    ca_std = published_tax.indexed(f'IF({single},Tax_Tables!$B$178,Tax_Tables!$B$42)',f'IF({single},Tax_Tables!$B$242,Tax_Tables!$B$241)',year,idx,2026)
    fed_cap = p('FedCapitalIncome')
    ca_cap = p('CaCapitalIncome')
    amount = p('AE')
    agi = p('AK')
    social = p('N')
    ordinary_tax = p('AN')
    ordinary_before_ss = p('AD')
    dividends = p('AF')
    low_ss = f'IF({single},Tax_Tables!$B$176,Tax_Tables!$B$135)'
    high_ss = f'IF({single},Tax_Tables!$B$177,Tax_Tables!$B$136)'
    taxable = f'MAX(0,{agi}-{fed_std})'
    regular_gain_part = f'MAX(0,{agi}-{fed_std}-{regular_base})'
    regular_cap_tax = _gain_tax(regular_base,regular_gain_part,single,idx,year)
    full_ordinary_tax = _bracket_tax(taxable,single,idx,(144,150),(6,12),year)
    amt_base = p('AmtTaxable')
    amt_preferred = p('AmtPreferential')
    amt_normal_part = f'({amt_base}-{amt_preferred})'
    amt_split_tax = f'{_amt_rates(amt_normal_part,idx,year)}+{_gain_tax(regular_base,amt_preferred,single,idx,year)}'
    amt_exemption = published_tax.indexed(f'IF({single},Tax_Tables!$B$174,Tax_Tables!$B$80)',f'IF({single},Tax_Tables!$B$236,Tax_Tables!$B$235)',year,idx)
    amt_phaseout = published_tax.indexed(f'IF({single},Tax_Tables!$B$175,Tax_Tables!$B$81)',f'IF({single},Tax_Tables!$B$238,Tax_Tables!$B$237)',year,idx)
    ca_taxable = p('AS')
    ca_senior_count = f'(IF({p("C")}>=65,1,0)+IF(AND(NOT({single}),{p("D")}>=65),1,0))'
    ca_count = f'(IF({single},1,2)+{ca_senior_count})'
    ca_old=f'IF({single},{CA_CREDIT_SINGLE_THRESHOLD},{CA_CREDIT_MFJ_THRESHOLD})'
    ca_phaseout_threshold = published_tax.indexed(ca_old,f'ROUND({ca_old}*1.034,0)',year,idx,2026)
    ca_personal=published_tax.indexed(f'IF({single},Tax_Tables!$B$179,Tax_Tables!$B$43)',f'IF({single},Tax_Tables!$B$244,Tax_Tables!$B$243)',year,idx,2026)
    ca_senior=published_tax.indexed('Tax_Tables!$B$179','Tax_Tables!$B$244',year,idx,2026)
    # Form 540 uses upward rounding to each $2,500 increment. Interpolate the
    # increments continuously for a forward estimate used by the cash solver.
    ca_phaseout = f'{CA_CREDIT_REDUCTION}*MAX(0,{agi}-{ca_phaseout_threshold})/{CA_CREDIT_STEP}'
    net_investment = f'MAX(0,SUM({p("AC")},{dividends},{fed_cap}))'

    return {
        'AE': f'{p("V")}*(1-{p("AB")})',
        'AH': f'SUM({ordinary_before_ss},{fed_cap},{dividends},{p("AG")},0.5*{social})',
        'AI': f'MIN(Tax_Tables!$B$138*{social},MIN(Tax_Tables!$B$137*MIN(MAX(0,{p("AH")}-{low_ss}),{high_ss}-{low_ss}),Tax_Tables!$B$137*{social})+Tax_Tables!$B$138*MAX(0,{p("AH")}-{high_ss}))',
        'AJ': f'{ordinary_before_ss}+{p("AI")}+MIN(0,{fed_cap})',
        'AK': f'SUM({p("AJ")},MAX(0,{fed_cap}),{dividends})',
        'AM': f'MAX(0,{p("AJ")}-{fed_std})',
        'AN': _bracket_tax(regular_base,single,idx,(144,150),(6,12),year),
        'AO': f'MAX(0,MIN({regular_cap_tax},{full_ordinary_tax}-{ordinary_tax}))',
        'AP': f'Tax_Tables!$B$74*MIN({net_investment},MAX(0,{agi}-IF({single},Tax_Tables!$B$173,Tax_Tables!$B$75)))',
        'AQ': f'MAX(0,{p("AmtTentative")}-{ordinary_tax}-{p("AO")})',
        'AS': f'MAX(0,{p("CaAgi")}-{ca_std})',
        'AT': f'MAX(0,{_bracket_tax(ca_taxable,single,idx,(160,168),(30,38),year)}+MAX(0,{ca_taxable}-Tax_Tables!$B$45)*Tax_Tables!$B$44-{p("CaExemptionCredit")})',
        'AV': f'SUM({ordinary_tax},{p("AO")},{p("AP")},{p("AQ")},{p("AR")},{p("AT")},{p("AU")})',
        'FedCapitalIncome': capital_agi_expr(amount,p('OpeningFedLoss')),
        'ClosingFedLoss': _closing_loss(amount,p('OpeningFedLoss'),fed_cap,agi,fed_std),
        'CaCapitalIncome': capital_agi_expr(amount,p('OpeningCaLoss')),
        'ClosingCaLoss': _closing_loss(amount,p('OpeningCaLoss'),ca_cap,p('CaAgi'),ca_std),
        'AmtTaxable': f'MAX(0,{agi}-MAX(0,{amt_exemption}-Tax_Tables!$B$82*MAX(0,{agi}-{amt_phaseout})))',
        'AmtPreferential': f'MIN({amt_base},MAX(0,{fed_cap})+{dividends})',
        'AmtTentative': f'MIN({_amt_rates(amt_base,idx,year)},{amt_split_tax})',
        'CaAgi': f'SUM({ordinary_before_ss},{p("K")},{ca_cap},{dividends})',
        'CaExemptionCredit': f'MAX(0,{ca_personal}+{ca_senior}*{ca_senior_count}-{ca_count}*{ca_phaseout})',
    }


def configure_workbook_state_policy(put: Callable[[str, str, str, str], None]) -> None:
    """Author the candidate California selector and gate before annual calculations run."""
    row = WORKBOOK_JURISDICTION_ROW
    code = CALIFORNIA_POLICY.jurisdiction
    selected = f'INDEX($D${row}:$F${row},1,$B$4)'
    put('Inputs', f'A{row}', WORKBOOK_JURISDICTION_LABEL, 'label')
    put('Inputs', f'B{row}', f'=IFERROR(IF(ISBLANK({selected}),"",{selected}),"")', 'input_active')
    for column in 'DEF':
        put('Inputs', f'{column}{row}', code, 'input')
    put('Inputs', f'G{row}', 'Use CA for a full-year California household.', 'note')

    schema_ref = f'Tax_Tables!{absolute_coordinate(WORKBOOK_STATE_POLICY_SCHEMA_CELL)}'
    version_ref = f'Tax_Tables!{absolute_coordinate(WORKBOOK_CA_POLICY_VERSION_CELL)}'
    put('Tax_Tables', WORKBOOK_STATE_POLICY_SCHEMA_CELL, WORKBOOK_STATE_POLICY_SCHEMA, 'note')
    put('Tax_Tables', WORKBOOK_CA_POLICY_VERSION_CELL, CALIFORNIA_POLICY.version, 'note')
    put('Tax_Tables', f'A{WORKBOOK_STATE_POLICY_SCHEMA_CELL[1:]}', 'State tax workbook schema', 'label')
    put('Tax_Tables', f'A{WORKBOOK_CA_POLICY_VERSION_CELL[1:]}', 'California state tax policy version', 'label')

    scenario_formulas = ','.join(f'_xlfn.ISFORMULA(${column}${row})' for column in 'DEF')
    checks = (f'EXACT($A${row},"{WORKBOOK_JURISDICTION_LABEL}"),_xlfn.ISFORMULA($B${row}),'
              f'EXACT($B${row},"{code}"),EXACT({selected},"{code}"),'
              f'NOT(CHOOSE($B$4,{scenario_formulas})),'
              f'NOT(_xlfn.ISFORMULA({schema_ref})),EXACT({schema_ref},"{WORKBOOK_STATE_POLICY_SCHEMA}"),'
              f'NOT(_xlfn.ISFORMULA({version_ref})),EXACT({version_ref},"{CALIFORNIA_POLICY.version}")')
    error = f'"{WORKBOOK_JURISDICTION_ERROR}"'
    put('Inputs', 'B129', f'=IFERROR(IF(AND({checks}),{_TAX_INPUT_CHECK},{error}),{error})', 'note')


def configure(m, wb):
    """Write Phase 2 tax inputs, sources, non-candidate payroll and basis ledger."""
    s = 'Inputs'
    m.put(s,'A125','TAX AND TAXABLE ACCOUNT','band')
    for r,label,default,note in (
        (126,'Primary share of combined wages',0.5,'Used in joint years with two earners. One earner or a single filer uses all modeled wages.'),
        (127,'Opening federal capital loss carryforward',0,'Enter the unused loss from the prior federal return as a positive amount.'),
        (128,'Opening California capital loss carryforward',0,'Enter the unused loss from the prior California return as a positive amount.'),
    ):
        m.put(s,f'A{r}',label,'label')
        m.put(s,f'B{r}',f'=INDEX($D{r}:$F{r},1,$B$4)','percent_active' if r==126 else 'input_active')
        for col in 'DEF':
            m.put(s,f'{col}{r}',default,'percent_input' if r==126 else 'input')
        m.put(s,f'G{r}',note,'note')
        m.put(s,f'H{r}',f'=IF(AND(D{r}=E{r},E{r}=F{r}),"","VARIES")','note')
    m.put(s,'A129','Tax inputs','label')
    m.put(s,'B129','='+_TAX_INPUT_CHECK,'note')
    m.put(s,'G129','A positive opening cost basis requires a positive opening taxable balance.','note')
    m.put(s,'G51','Enter adjusted basis. It may exceed market value. Sales release basis in proportion to the amount sold.','note')

    tax = 'Tax_Tables'
    for r,label,value,note in (
        (190,'Annual net capital loss deduction limit (MFJ or single)',3000,'IRS Topic 409; California Schedule D (540), line 9. Nominal limit.'),
        (191,'Additional Medicare tax threshold (Single)',200000,'IRS Topic 560. Nominal threshold. Household liability uses filing status.'),
        (192,'California exemption credit phaseout start (MFJ, 2025)',504411,'FTB 2025 Form 540, AGI Limitation Worksheet. Forecast indexes this threshold.'),
        (193,'California exemption credit phaseout start (Single, 2025)',252203,'FTB 2025 Form 540, AGI Limitation Worksheet. Forecast indexes this threshold.'),
        (194,'California exemption credit AGI phaseout increment',2500,'FTB 2025 Form 540. Forecast interpolates instead of rounding up.'),
        (195,'California exemption credit reduction per exemption and increment',6,'FTB 2025 Form 540. Forecast interpolates; filing calculations round up.'),
    ):
        m.put(tax,f'A{r}',label,'label')
        m.put(tax,f'B{r}',value,'money')
        m.put(tax,f'D{r}',note,'note')
    m.put(tax,'A20','Source: IRS Rev. Proc. 2025-32. Gains and qualified dividends stack above ordinary taxable income. The model includes NIIT.')
    m.put(tax,'A26','AMT exemption (MFJ) — see section 10 for modeled rules')
    m.put(tax,'D51','SSA contribution and benefit base; capped per earner. Future bases use the tax index proxy.')
    m.put(tax,'D80','IRS Rev. Proc. 2025-32; Form 6251 ordinary and preferential tax computations.')

    s='Projection'
    for col,label in {'AB':'Basis per taxable dollar','AE':'Realized gain or loss','AQ':'Additional federal AMT',
                      'AR':'Payroll tax by earner','BT':'Primary wages','BU':'Spouse wages',
                      **{col:EXTRA_HEADERS[key] for key,col in PROJECTION_EXTRA.items()},
                      'CE':'Taxable basis removed','CF':'Basis added to retained investments'}.items():
        m.put(s,f'{col}4',label,'header')
    for r in range(5,m.END+1):
        prev=r-1
        raw = {
            'AB': ('IF(Inputs!$B$25>0,Inputs!$B$51/Inputs!$B$25,0)' if r==5
                   else f'IF(BC{prev}>0,BD{prev}/BC{prev},0)'),
            'BT': f'IF(OR(F{r}="Single",Inputs!$B$11=1),I{r},I{r}*Inputs!$B$126)',
            'BU': f'I{r}-BT{r}',
            'CE': '0' if r==5 else f'IF(V{r}>=BC{prev},BD{prev},MIN(BD{prev},V{r}*AB{r}))',
            'CF': '0' if r==5 else f'IF(BC{r}>0,BB{r}+IF(BC{prev}>0,(AF{r}+AC{r})*(MAX(0,BC{prev}-V{r})/BC{prev}),0),0)',
        }
        if r>5:
            year=f'B{r}'; idx=f'H{r}'
            wage_cap=published_tax.indexed('Tax_Tables!$B$51','Tax_Tables!$B$240',year,idx)
            limit=f'IF(F{r}="Single",{MEDICARE_SINGLE},Tax_Tables!$B$54)'
            raw['AR'] = f'IF(E{r}="Working",Tax_Tables!$B$50*(MIN(BT{r},{wage_cap})+MIN(BU{r},{wage_cap}))+Tax_Tables!$B$52*I{r}+Tax_Tables!$B$53*MAX(0,I{r}-{limit}),0)'
            sdi_rate = f'IF({published_tax.use(year, 2027)},Tax_Tables!$B$251,Tax_Tables!$B$46)'
            raw['AU'] = f'IF(E{r}="Working",I{r}*{sdi_rate},0)'
            std_joint=published_tax.indexed('Tax_Tables!$B$23','Tax_Tables!$B$231',year,idx)
            std_single=published_tax.indexed('Tax_Tables!$B$171','Tax_Tables!$B$232',year,idx)
            add_joint=published_tax.indexed('Tax_Tables!$B$24','Tax_Tables!$B$233',year,idx)
            add_single=published_tax.indexed('Tax_Tables!$B$172','Tax_Tables!$B$234',year,idx)
            raw['AL']=f'IF(F{r}="Single",{std_single}+IF(C{r}>=65,{add_single},0),{std_joint}+(--(C{r}>=65)+--(D{r}>=65))*{add_joint})'
            # Preserve basis through a price decline. Only selling releases
            # it; after complete liquidation no stale basis remains.
            raw['BD']=f'IF(BC{r}>0,MAX(0,BD{prev}-CE{r})+CF{r},0)'
        for col,formula in raw.items():
            m.put(s,f'{col}{r}',m.active(r,formula),'money' if col in ('BT','BU','CE','CF') else None)


def bind_projection(m, candidate_final_ref):
    """Expose accepted candidate tax measures; bind after the funding pass.

    candidate_final_ref(year_row, semantic_key) returns the selected candidate
    cell address (without leading '=').
    """
    for r in range(5,m.END+1):
        for key,col in PROJECTION_EXTRA.items():
            value = ({'ClosingFedLoss':'Inputs!$B$127','ClosingCaLoss':'Inputs!$B$128'}.get(key,'0')
                     if r==5 else candidate_final_ref(r,key))
            m.put('Projection',f'{col}{r}',m.active(r,value),'money')
