"""Published federal, California, and payroll values used before forecast indexing."""

from state_policy import CALIFORNIA_POLICY

FEDERAL = {
    6: (0, 25600, 104050, 218250, 416650, 529100, 793650),
    17: (0, 102100, 633600),
    144: (0, 12800, 52025, 109125, 208325, 264550, 661375),
    154: (0, 51050, 563200),
}
CALIFORNIA = {
    30: (0, 22912, 54314, 85722, 118996, 150394, 768218, 921854, 1536426),
    160: (0, 11456, 27157, 42861, 59498, 75197, 384109, 460927, 768213),
}
SCALARS = {
    231: ('2027 federal standard deduction — joint', 33200),
    232: ('2027 federal standard deduction — single', 16600),
    233: ('2027 additional standard deduction per eligible joint filer', 1700),
    234: ('2027 additional standard deduction — eligible single filer', 2100),
    235: ('2027 federal AMT exemption — joint', 144700),
    236: ('2027 federal AMT exemption — single', 93000),
    237: ('2027 federal AMT exemption phaseout starts — joint', 1031900),
    238: ('2027 federal AMT exemption phaseout starts — single', 515950),
    239: ('2027 federal AMT 28% threshold', 252500),
    240: ('2027 Social Security wage base per earner', 192300),
    241: ('2026 California standard deduction — joint', 11800),
    242: ('2026 California standard deduction — single', 5900),
    243: ('2026 California joint personal exemption credit', 316),
    244: ('2026 California individual personal or senior credit', 158),
}

def use(year, base):
    return f'AND(Tax_Tables!$B$219=1,({year})>={base})'

def indexed(old, new, year, index, base=2027):
    factor=f'({index})/(1+Inputs!$B$14)' if base==2027 else f'({index})'
    return f'IF({use(year,base)},{new}*({factor}),{old}*({index}))'

def configure(m):
    s='Tax_Tables'
    m.put(s,'A219','Use published tax-year values', 'label')
    m.put(s,'B219',1,'input')
    m.put(s,'D219','1 = published values followed by modeled inflation. 0 = original baseline for controlled comparisons.','note')
    for source,base in [(FEDERAL,2027),(CALIFORNIA,CALIFORNIA_POLICY.published_income_year)]:
        for start,values in source.items():
            m.put(s,f'E{start-1}',f'{base} bracket floor','header')
            for r,value in enumerate(values,start): m.put(s,f'E{r}',value,'money')
    m.put(s,'A230','PUBLISHED TAX-YEAR AMOUNTS','band')
    for r,(label,value) in SCALARS.items():
        m.put(s,f'A{r}',label,'label'); m.put(s,f'B{r}',value,'money')
    m.put(s,'A245','Federal: IRS Rev. Proc. 2026-35','note')
    m.put(s,'A246','https://www.irs.gov/pub/irs-drop/rp-26-35.pdf','note')
    m.put(s,'A247','Social Security: SSA Contribution and Benefit Base','note')
    m.put(s,'A248','https://www.ssa.gov/oact/cola/cbb.html','note')
    m.put(s,'A249','California: FTB October 2026 indexed rates and credits','note')
    m.put(s,'A250',CALIFORNIA_POLICY.income_source,'note')
    m.put(s,'A251',f'{CALIFORNIA_POLICY.published_sdi_year} California disability insurance rate','label')
    m.put(s,'B251',0.014,'percent_active')
    m.put(s,'D251','Future projections use this rate until a new rate is entered.','note')
    m.put(s,'A252',CALIFORNIA_POLICY.sdi_source,'note')
