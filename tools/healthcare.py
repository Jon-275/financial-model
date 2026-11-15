"""Annual coverage, HSA reimbursement, and ACA premium credits.

The household budget switches each person’s share to the Medicare budget at 65.
Detailed coverage is an opt-in, annual US-contiguous-states model.
The funding solver evaluates the same credit formula at each candidate MAGI.
"""


SHEET = 'Healthcare'
INPUT_STATUS = 'Inputs!$B$125'

CONTROL_ROWS = (
    ('Detailed healthcare model?', 0, 'Use 0 for a household budget or 1 for costs by person.', 'integer'),
    ('Primary pre-Medicare coverage', 0, '0 None, 1 Marketplace, 2 Other, 3 COBRA. See the Model guide for 4.', 'integer'),
    ('Spouse pre-Medicare coverage', 0, '0 None, 1 Marketplace, 2 Other, 3 COBRA. See the Model guide for 4.', 'integer'),
    ('Primary Medicare starts at age', 65, 'Enter the confirmed start age. Use 126 for no enrollment.', 'integer'),
    ('Spouse Medicare starts at age', 65, 'Enter the confirmed start age. Use 126 for no enrollment.', 'integer'),
    ('Primary pre-Medicare gross premium', 0, 'Enter this person’s yearly premium before credits.', 'dollar'),
    ('Spouse pre-Medicare gross premium', 0, 'Enter this person’s yearly premium before credits.', 'dollar'),
    ('Primary Medicare basic premium', 0, 'Enter yearly basic Medicare premiums. See the Model guide.', 'dollar'),
    ('Spouse Medicare basic premium', 0, 'Enter yearly basic Medicare premiums. See the Model guide.', 'dollar'),
    ('Primary Medigap and other premiums', 0, 'Enter yearly Medigap or other supplemental premiums.', 'dollar'),
    ('Spouse Medigap and other premiums', 0, 'Enter yearly Medigap or other supplemental premiums.', 'dollar'),
    ('Primary qualified out-of-pocket costs', 0, 'Enter yearly medical bills your HSA can cover after insurance.', 'dollar'),
    ('Spouse qualified out-of-pocket costs', 0, 'Enter yearly medical bills your HSA can cover after insurance.', 'dollar'),
    ('Primary other out-of-pocket costs', 0, 'Enter yearly out-of-pocket costs your HSA can’t cover.', 'dollar'),
    ('Spouse other out-of-pocket costs', 0, 'Enter yearly out-of-pocket costs your HSA can’t cover.', 'dollar'),
    ('Primary eligible for Marketplace credit?', 0, 'Use 1 if eligible or 0 otherwise. See the Model guide.', 'integer'),
    ('Spouse eligible for Marketplace credit?', 0, 'Use 1 if eligible or 0 otherwise. See the Model guide.', 'integer'),
    ('Marketplace benchmark: both eligible', 0, 'Enter the yearly benchmark quote for both eligible people.', 'dollar'),
    ('Marketplace benchmark: primary only', 0, 'Enter the yearly benchmark quote for you alone.', 'dollar'),
    ('Marketplace benchmark: spouse only', 0, 'Enter the yearly benchmark quote for your spouse alone.', 'dollar'),
    ('HSA account owner', 1, 'Use 1 if you own the HSA or 2 if your spouse owns it.', 'integer'),
    ('Medicare surcharge income, prior year', -1, 'Enter income for the year before row 130. Use -1 if unknown.', 'dollar'),
    ('Medicare surcharge income, opening year', -1, 'Enter income for the year in row 130. Use -1 if unknown.', 'dollar'),
    ('HSA-qualified share of long-term care', 0, 'Enter the share of care costs your HSA can cover. Use 0% if unsure.', 'percent'),
)

HEADERS = (
    'Year', 'Primary age', 'Spouse age', 'Primary coverage', 'Spouse coverage',
    'Primary premiums', 'Spouse premiums', 'Primary qualified expenses', 'Spouse qualified expenses',
    'Other expenses', 'Long-term care', 'Medicare income surcharge', 'Eligible Marketplace premiums',
    'Eligible Marketplace benchmark', 'Applicable poverty guideline', 'Gross healthcare cost',
    'HSA-qualified healthcare', 'Marketplace tax credit', 'Net healthcare cost', 'Healthcare status',
    'Primary credit eligible', 'Spouse credit eligible', 'Primary Medicare', 'Spouse Medicare',
    'IRMAA lookback MAGI', 'IRMAA lookback filing status', 'IRMAA per beneficiary',
    'Qualified insurance premiums', 'Marketplace household MAGI', 'Marketplace MAGI / FPL',
    'HSA reimbursement', 'Healthcare paid from cash',
)


def _hc(col, row):
    return f'{SHEET}!${col}${row}'


def magi(agi, ss, taxable_ss, exempt):
    """ACA MAGI adds back tax-exempt interest and nontaxable Social Security."""
    return f'(({agi})+MAX(0,({ss})-({taxable_ss}))+({exempt}))'


def applicable_rate(ratio, first):
    """Unrounded interpolation within the IRS annual contribution bands."""
    expr = f'Tax_Tables!$C${first+5}'
    for row in range(first+4, first-1, -1):
        start = f'Tax_Tables!$A${row}'
        end = f'Tax_Tables!$B${row}'
        low = f'Tax_Tables!$C${row}'
        high = f'Tax_Tables!$D${row}'
        rate = low if row == first else f'({low}+(({ratio})-{start})*({high}-{low})/({end}-{start}))'
        expr = f'IF(({ratio})<{end},{rate},{expr})'
    return expr


def credit_from(magi, actual, benchmark, fpl, year):
    """Estimate annual federal PTC, bounded by eligible enrollment premium.

    The 100% and 400% endpoints are inclusive under the ordinary rules. The
    inputs to this function already exclude months/persons not eligible for PTC.
    """
    ratio = f'(({magi})/({fpl}))'
    rate = f'IF(({year})>=2027,{applicable_rate(ratio,221)},{applicable_rate(ratio,204)})'
    return (f'IF(OR(({actual})<=0,({benchmark})<=0,({fpl})<=0),0,'
            f'IF(OR(({magi})<({fpl}),({magi})>4*({fpl})),0,'
            f'MIN(({actual}),MAX(0,({benchmark})-({magi})*{rate}))))')


credit_expr = credit_from


def credit(row, magi):
    """Candidate PTC for a fixed projection year."""
    return credit_from(magi, _hc('M', row), _hc('N', row), _hc('O', row), _hc('A',row))


def _tables(m):
    s = 'Tax_Tables'
    m.put(s, 'A202', 'MARKETPLACE PREMIUM TAX CREDIT (2026 RULES)', 'band')
    m.row(s, 203, ('Band starts at', 'Band ends at', 'Initial contribution rate', 'Final contribution rate'), 'header')
    bands = ((1, 1.33, .0210, .0210), (1.33, 1.5, .0314, .0419), (1.5, 2, .0419, .0660),
             (2, 2.5, .0660, .0844), (2.5, 3, .0844, .0996), (3, 4, .0996, .0996))
    for row, band in enumerate(bands, 204):
        m.row(s, row, band, 'percent_active')
    m.put(s, 'A210', 'Source: IRS Rev. Proc. 2025-25. Applies to 2026 only.', 'note')
    m.put(s, 'A211', 'https://www.irs.gov/pub/irs-drop/rp-25-25.pdf', 'note')
    m.put(s, 'A212', 'POVERTY GUIDELINES USED FOR MARKETPLACE COVERAGE: CONTIGUOUS STATES AND DC', 'band')
    m.row(s, 213, ('Marketplace coverage year', 'One-person tax household', 'Two-person tax household'), 'header')
    m.put(s, 'A214', 2026, 'whole')
    m.put(s, 'B214', 15650, 'money')
    m.put(s, 'C214', 21150, 'money')
    m.put(s, 'A215', 2027, 'whole')
    m.put(s, 'B215', 15960, 'money')
    m.put(s, 'C215', 21640, 'money')
    m.put(s, 'A216', 'Source: HHS 2025 and 2026 poverty guidelines. The 2027 coverage amounts use the 2026 HHS release and are increased with CPI after 2027.', 'note')
    m.put(s, 'A217', 'https://aspe.hhs.gov/topics/poverty-economic-mobility/poverty-guidelines/prior-hhs-poverty-guidelines-federal-register-references', 'note')
    m.row(s,220,('2027 FPL band starts at','Band ends at','Initial contribution rate','Final contribution rate'),'header')
    rates2027=((.0215,.0215),(.0323,.0430),(.0430,.0678),(.0678,.0866),(.0866,.1022),(.1022,.1022))
    for row, ((left,right,_,_),rates) in enumerate(zip(bands,rates2027),221):
        m.row(s,row,(left,right,*rates),'percent_active')
    m.put(s, 'A227', 'Source: IRS Rev. Proc. 2026-26. The 2027 contribution percentages are held constant for later years.', 'note')
    m.put(s, 'A228', 'https://www.irs.gov/irb/2026-31_irb', 'note')


def _inputs(m):
    s = 'Inputs'
    m.put(s, 'A100', 'HEALTHCARE BY PERSON (OPTIONAL)', 'band')
    for row, (label, default, note, kind) in enumerate(CONTROL_ROWS, 101):
        m.put(s, f'A{row}', label, 'label')
        active_style = 'percent_active' if kind == 'percent' else 'input_active'
        editable_style = 'percent_input' if kind == 'percent' else 'input'
        m.put(s, f'B{row}', f'=INDEX($D{row}:$F{row},1,$B$4)', active_style)
        for col in 'DEF':
            m.put(s, f'{col}{row}', default, editable_style)
        m.put(s, f'G{row}', note, 'note')
        m.put(s, f'H{row}', f'=IF(AND(D{row}=E{row},E{row}=F{row}),"","VARIES")', 'note')

    ints = tuple((row, low, high) for row, low, high in (
        (102, 0, 4), (103, 0, 4), (104, 0, 126), (105, 0, 126),
        (116, 0, 1), (117, 0, 1), (121, 1, 2),
    ))
    checks = ','.join((
        'COUNT(INDEX(D101:F124,0,B4))=24',
        *(f'B{row}=INT(B{row}),B{row}>={low},B{row}<={high}' for row, low, high in ints),
        'MIN(B106:B115,B118:B120)>=0', 'MIN(B122:B123)>=-1',
        'B124>=0', 'B124<=1',
        'OR(B116=0,B102=1,B102=4)', 'OR(B117=0,B103=1,B103=4)',
    ))
    status = (f'=IFERROR(IF(COUNT(INDEX(D101:F101,1,B4))<>1,"Check healthcare inputs",'
              f'IF(B101=0,"OK",IF(AND(B101=1,{checks}),"OK","Check healthcare inputs"))),"Check healthcare inputs")')
    m.put(s, 'A125', 'Healthcare input check', 'label')
    m.put(s, 'B125', status, 'note')
    m.put(s, 'G125', 'Check Healthcare for missing quotes or Medicare income.', 'note')


def _irmaa_per_beneficiary(mag, filing, factor, year):
    """Annual Medicare Part B plus Part D surcharge for one enrollee.

    Ordinary boundaries are exceeded strictly; the last tier starts at its
    stated cutoff. The top threshold is unchanged through 2027. Below that
    tier, the model estimates future thresholds and premiums with CPI.
    """
    def threshold(single_row, factor, joint_ratio):
        base = f'ROUND(Tax_Tables!$B${single_row}*({factor})/1000,0)*1000'
        return f'({base})*IF({filing}="Single",1,{joint_ratio})'

    others = '+'.join(f'--(({mag})>{threshold(row, factor, 2)})' for row in range(183, 187))
    top_factor = f'IF(({year})<=2027,1,MAX(1,({factor})/(1+Inputs!$B$14)))'
    top = f'--(({mag})>={threshold(187, top_factor, 1.5)})'
    return f'INDEX(Tax_Tables!$E$89:$E$94,1+{others}+{top})*({factor})'


def _healthcare_rows(m):
    s = SHEET
    m.put(s, 'A1', 'Healthcare by person', 'title')
    m.put(s, 'A2', 'Costs are annual. Detailed mode separates Marketplace credits, Medicare surcharges and HSA reimbursements.', 'note')
    m.row(s, 4, HEADERS, 'header')
    for row in range(5, m.END + 1):
        p = lambda col: m.pj(col, row)
        if row == 5:
            vals = [p('B'), p('C'), p('D'), '"Opening balances"', '"Opening balances"',
                    *('0',) * 14, '"Opening balances"', *('0',) * 5, p('F'), *('0',) * 6]
        else:
            year, age, spouse, phase, filing = (p(k) for k in ('B', 'C', 'D', 'E', 'F'))
            detailed = 'Inputs!$B$101=1'
            legacy = 'Inputs!$B$101=0'
            single = f'{filing}="Single"'
            retired = f'AND({detailed},{phase}<>"Working")'
            living = f'{filing}<>"Single"'
            owner_alive = f'OR(Inputs!$B$121=1,{living})'
            hsa_owner_age = f'IF(Inputs!$B$121=1,{age},{spouse})'
            premium_qualified = f'{hsa_owner_age}>=65'
            coverage = lambda med, pre: f'IF({med}=1,"Medicare",CHOOSE({pre}+1,"None","Marketplace","Other","COBRA","Marketplace"))'
            lookback = 'Inputs!$B$122' if row == 6 else 'Inputs!$B$123' if row == 7 else f'{m.pj("AK",row-2)}+{m.pj("AG",row-2)}'
            prior_filing = (f'IF(Inputs!$B$8-{7-row}>Inputs!$B$54,"Single","MFJ")'
                            if row < 8 else m.pj('F', row-2))
            missing_irmaa = f'AND(W{row}+X{row}>0,Y{row}=-1)' if row < 8 else 'FALSE()'
            legacy_year = max(5, row - 2)
            legacy_magi = f'({m.pj("AK", legacy_year)}+{m.pj("AG", legacy_year)})/{p("H")}'
            legacy_tier = (f'IF({single},MATCH({legacy_magi},Tax_Tables!$B$182:$B$187,1),'
                           f'MATCH({legacy_magi},Tax_Tables!$B$89:$B$94,1))')
            med_index = f'(1+IF(ISNUMBER(Inputs!$B$132),Inputs!$B$132,Inputs!$B$14))^{p("A")}'
            legacy_count = f'--({age}>=65)+--AND(NOT({single}),{spouse}>=65)'
            legacy_irmaa = f'IF({phase}="Working",0,INDEX(Tax_Tables!$E$89:$E$94,{legacy_tier})*{p("H")}*({legacy_count}))'
            primary_cost = f'IF({age}<65,Inputs!$B$42,Inputs!$B$43)'
            spouse_cost = f'IF({spouse}<65,Inputs!$B$42,Inputs!$B$43)'
            budget = f'IF({single},{primary_cost}*Inputs!$B$57,({primary_cost}+{spouse_cost})/2)'
            legacy_cost = f'IF({phase}="Working",0,({budget})*({med_index})+L{row})+K{row}'
            # Annual premiums use inflation relative to the opening date. FPL
            # starts from the published 2027 coverage guideline when available.
            fpl = (f'IF(AND(M{row}>0,N{row}>0),IF(A{row}<=Tax_Tables!$A$214,'
                   f'IF({single},Tax_Tables!$B$214,Tax_Tables!$C$214),'
                   f'IF({single},Tax_Tables!$B$215,Tax_Tables!$C$215)*{p("H")}/(1+Inputs!$B$14)),0)')
            actual_magi = magi(p('AK'), p('N'), p('AI'), p('AG'))
            cols = {
                'A': year, 'B': age, 'C': spouse,
                'D': f'IF({legacy},"Legacy budget",IF({phase}="Working","Working budget",{coverage(f"W{row}","Inputs!$B$102")}))',
                'E': f'IF({single},"Not covered",IF({legacy},"Legacy budget",IF({phase}="Working","Working budget",{coverage(f"X{row}","Inputs!$B$103")})))',
                'F': f'IF({retired},IF(W{row}=1,SUM(Inputs!$B$108,Inputs!$B$110),IF(Inputs!$B$102=0,0,Inputs!$B$106))*({med_index}),0)',
                'G': f'IF(AND({retired},{living}),IF(X{row}=1,SUM(Inputs!$B$109,Inputs!$B$111),IF(Inputs!$B$103=0,0,Inputs!$B$107))*({med_index}),0)',
                'H': f'IF({retired},Inputs!$B$112*({med_index}),0)',
                'I': f'IF(AND({retired},{living}),Inputs!$B$113*({med_index}),0)',
                'J': f'IF({retired},(Inputs!$B$114+IF({single},0,Inputs!$B$115))*({med_index}),0)',
                'K': p('AY'),
                'L': f'IF({legacy},{legacy_irmaa},IF({retired},AA{row}*(W{row}+X{row}),0))',
                'M': f'IF(U{row}=1,F{row},0)+IF(V{row}=1,G{row},0)',
                'N': f'IF(U{row}+V{row}=2,Inputs!$B$118,IF(U{row}=1,Inputs!$B$119,IF(V{row}=1,Inputs!$B$120,0)))*({med_index})',
                'O': fpl,
                'P': f'IF({legacy},{legacy_cost},SUM(F{row}:L{row}))',
                'Q': (f'IF({legacy},IF({phase}="Working",0,P{row}),'
                      f'IF(AND({retired},{owner_alive}),MIN(P{row},SUM(H{row}:I{row})+AB{row}+K{row}*Inputs!$B$124),0))'),
                'R': credit(row, f'AC{row}'),
                'S': f'MAX(0,P{row}-R{row})',
                'T': (f'IF({legacy},"Legacy budget",IF(OR(AND(D{row}="Marketplace",F{row}<=0),AND(E{row}="Marketplace",G{row}<=0),'
                      f'AND(U{row}+V{row}>0,N{row}<=0)),"Missing ACA quote",'
                      f'IF({missing_irmaa},"Missing IRMAA MAGI",'
                      f'IF(OR(AND(W{row}=1,F{row}<=0),AND(X{row}=1,G{row}<=0)),"Missing Medicare cost",'
                      f'IF(AND(NOT({owner_alive}),{m.pj("BG",row-1)}>0),"Confirm HSA ownership","OK")))))'),
                'U': f'--AND(D{row}="Marketplace",Inputs!$B$116=1,OR({age}<65,Inputs!$B$102=4))',
                'V': f'--AND(E{row}="Marketplace",Inputs!$B$117=1,OR({spouse}<65,Inputs!$B$103=4))',
                'W': f'--AND({retired},Inputs!$B$104<126,{age}>=Inputs!$B$104)',
                'X': f'--AND({retired},{living},Inputs!$B$105<126,{spouse}>=Inputs!$B$105)',
                'Y': f'IF(W{row}+X{row}=0,0,{lookback})',
                'Z': f'IF(W{row}+X{row}=0,"",{prior_filing})',
                'AA': f'IF(OR(W{row}+X{row}=0,Y{row}<0),0,{_irmaa_per_beneficiary(f"Y{row}",f"Z{row}",p("H"),year)})',
                'AB': (f'IF(AND({retired},{owner_alive}),IF(D{row}="COBRA",F{row},0)+IF(E{row}="COBRA",G{row},0)'
                       f'+IF({premium_qualified},(W{row}*Inputs!$B$108+X{row}*Inputs!$B$109)*({med_index})+L{row},0),0)'),
                'AC': actual_magi,
                'AD': f'IF(O{row}>0,AC{row}/O{row},0)',
                'AE': p('P'),
                'AF': f'MAX(0,S{row}-AE{row})',
            }
            vals = [cols[m.L(i)] for i in range(1, len(HEADERS)+1)]

            # Projection is the book of record for gross spending and HSA draws.
            m.put('Projection', f'AY{row}', m.active(row,
                f'IF(AND(C{row}>=Inputs!$B$58,C{row}<Inputs!$B$58+Inputs!$B$60),Inputs!$B$59*({med_index}),0)'))
            m.put('Projection', f'AZ{row}', m.active(row, _hc('L', row)))
            m.put('Projection', f'BA{row}', m.active(row, _hc('P', row)))
            m.put('Projection', f'P{row}', m.active(row, f'MIN(BG{row-1},{_hc("Q",row)})'))
            # Employee HSA payroll contribution is separate from reimbursement.
            # In detailed mode, no contribution is made for an owner already
            # enrolled in Medicare or for the deceased spouse.
            owner_age = f'IF(Inputs!$B$121=1,C{row},D{row})'
            start = 'IF(Inputs!$B$121=1,Inputs!$B$104,Inputs!$B$105)'
            prior_eligible = f'OR(Inputs!$B$121=1,F{row}<>"Single")'
            eligibility = f'OR({legacy},AND({prior_eligible},OR({start}=126,{owner_age}<{start})))'
            contrib = f'IF(AND(E{row}="Working",{eligibility}),MIN(MAX(0,I{row}-J{row}),Inputs!$B$34*G{row}),0)'
            m.put('Projection', f'K{row}', m.active(row, contrib))

        if len(vals) != len(HEADERS):
            raise ValueError(f'Healthcare row {row}: got {len(vals)} values, expected {len(HEADERS)}')
        for i, expr in enumerate(vals, 1):
            col = m.L(i)
            style = ('whole' if col in ('A', 'B', 'C', 'U', 'V', 'W', 'X') else
                     'label' if col in ('D', 'E', 'T', 'Z') else
                     'percent_active' if col == 'AD' else 'money')
            m.put(s, f'{col}{row}', m.active(row, expr), style)


def configure(m, wb):
    """Append healthcare controls, tax inputs and annual schedules to a plan.

    Run after base Projection formulas so healthcare-owned cells are the final
    writes. Cash/Funding owns application of the PTC to the annual ledger.
    """
    if SHEET not in m.MODULES:
        m.MODULES.append(SHEET)
    _tables(m)
    _inputs(m)
    _healthcare_rows(m)
    m.put('Projection', 'BA4', 'Gross healthcare before Marketplace credit')
    m.put('Projection', 'P4', 'HSA reimbursement for eligible healthcare')
