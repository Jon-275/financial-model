"""Build the reader-facing plan from the existing annual schedules."""

import risk_presentation as risk
from targeting import DETAIL as CONVERSION_DETAIL

SUMMARY = 'Summary'
METRICS = 'Planning Metrics'
METRICS_REF = f"'{METRICS}'!"
METRIC_SECTIONS = {
    7: 'Savings target',
    14: 'At retirement',
    22: 'Funding and balances',
    37: 'Taxes and conversions',
    44: 'Healthcare at retirement (future dollars)',
    49: 'Scenario comparison',
}
APPENDIX = 'Annual Detail'
INDEX = 'Appendix'
DOLLAR_VIEW_CELL = 'K18'
TODAY_DOLLARS = 'Today’s dollars'
FUTURE_DOLLARS = 'Future dollars'
FUTURE_VIEW = f'{SUMMARY}!$K$18="{FUTURE_DOLLARS}"'
DECISION_TABS = (METRICS, 'Asset_Allocation', 'Income_Targeting', 'Risk_Settings', risk.SHEET, 'Tax_Tables')
DETAIL_GROUPS = (
    ('Annual schedules', (
        ('Cash_Flow', 'Annual cash flow', 'Income, tax, spending, withdrawals and unpaid costs.'),
        ('Projection', 'Full projection', 'Year-by-year income, assets and portfolio balances.'),
        ('Roth_Access', 'Roth account access', 'Balances and available withdrawals by owner.'),
        ('Healthcare', 'Healthcare', 'Coverage, premiums and credits by person.'),
    )),
    ('Calculation checks', (
        ('Checks', 'Model checks', 'Reconcile cash, taxes, accounts and healthcare.'),
    )),
    ('Supporting data and calculations', (
        (CONVERSION_DETAIL, 'Conversion detail', 'Tax, healthcare costs and withdrawals for each tested conversion.'),
        ('Report_Data', 'Report data', 'Source values for the plan and its chart.'),
        ('Funding_Calc', 'Withdrawal calculation', 'The funding trials for each projected year.'),
        ('Targeting_Calc', 'Conversion calculation', 'The funding trials for the conversion comparison.'),
        (risk.DATA_SHEET, 'Market risk data', 'Saved simulation results and their input comparison.'),
    )),
)
FIRST_TRANSITION_ROW = 41
LAST_TRANSITION_ROW = 53
FIRST_APPENDIX_ROW = 6
LAST_APPENDIX_ROW = 131
INDEX_LINK_ROW = 30
CASH_HEADERS = ('Year', 'Age', 'Income used', 'Savings used', 'Other spending',
                'Healthcare', 'Tax due', 'Unpaid')


def display_dollars(real_amount, price_index):
    return f'({real_amount})*IF({FUTURE_VIEW},{price_index},1)'


def price_index_at(position):
    return f'INDEX(Projection!$G$5:$G$130,{position})'


def move_metrics(template, updates, styles, put):
    """Keep the supporting calculations on their own worksheet."""
    for row in template.iter_rows(max_col=3):
        for cell in row:
            updates.setdefault(cell.coordinate, cell.value)
    metrics = list(updates.items())
    updates.clear()
    styles.clear()
    currency = {5, 6, 8, 9, 15, 16, 17, 18, 19, 24, 27, 28, 29, 30, 31,
                32, 34, 38, 40, 42, 45, 46, 47}
    percent = {12, 39, 41}
    for address, value in metrics:
        if value is None:
            continue
        column, row = address[0], int(address[1:])
        style = {'A': 'report_body', 'B': 'report_row', 'C': 'report_caption_gutter'}[column]
        if column == 'B':
            if row == 3:
                style = 'report_active'
            elif row == 25:
                style = 'report_body'
            elif row in currency:
                style = 'report_money'
            elif row in percent:
                style = 'report_percent'
        put(METRICS, address, value, style)
    put(METRICS, 'A1', 'Planning metrics', 'report_title')
    put(METRICS, 'A2', '=IF(Inputs!$B$70=0,"Check Inputs","Scenario "&Inputs!$B$7&"   •   Retire at age "&Inputs!$B$10&"   •   Plan through age "&Inputs!$B$53)', 'report_caption')
    put(METRICS, 'A3', 'Does the plan cover its costs?', 'report_section')
    for row, label in {
        6: 'Workplace Roth unavailable at end (today’s dollars)',
        9: 'Spending after the mortgage ends (today’s dollars)',
        16: 'Portfolio at retirement (today’s dollars)',
        19: 'Savings available at retirement (future dollars)',
        25: 'Funding calculation status',
        27: 'Largest annual shortfall (future dollars)',
        28: 'Portfolio at age 65 (today’s dollars)',
        34: 'First-year Roth savings below target (future dollars)',
        38: 'Total tax, final working year (future dollars)',
        40: 'Total tax, first retirement year (future dollars)',
        42: 'Total Roth converted (future dollars)',
    }.items():
        put(METRICS, f'A{row}', label, 'report_body')
    for row, label in METRIC_SECTIONS.items():
        put(METRICS, f'A{row}', label, 'detail_band')
    for address, label, style in (
        ('E1', 'Related views', 'report_section'),
        ('E2', 'Back to Summary', 'report_link'),
        ('E3', 'Open Annual Detail', 'report_link'),
        ('E4', 'Open Cash_Flow', 'report_link'),
    ):
        put(METRICS, address, label, style)


def _annual_amounts(row):
    cash = lambda col: f'Cash_Flow!{col}{row}'
    proj = lambda col: f'Projection!{col}{row}'
    paid = f'({cash("N")}+{cash("P")})'
    real_amounts = (
        cash('A'), cash('B'),
        f'MIN({cash("C")},{paid})/{proj("G")}',
        f'(MAX(0,{paid}-{cash("C")})+{cash("E")})/{proj("G")}',
        f'({cash("D")}-{proj("BA")})/{proj("G")}',
        f'{proj("BA")}/{proj("G")}',
        f'({cash("F")}+{cash("M")})/{proj("G")}',
        f'{cash("V")}/{proj("G")}',
        proj('BN'),
    )
    return (*real_amounts[:2], *(display_dollars(amount, proj('G')) for amount in real_amounts[2:]))


def configure(put):
    def summary(address, value, style='report_body'):
        put(SUMMARY, address, value, style)

    def appendix(address, value, style='report_body'):
        put(APPENDIX, address, value, style)

    ok = 'Report_Data!$Q$2>0'
    opening_only = 'Report_Data!$Q$2=1'
    end_age = 'Inputs!$B$53'
    retirement = 'Report_Data!$Q$5'
    gap = 'Report_Data!$Q$6'
    first_age = f'INDEX(Projection!$C$5:$C$130,{gap})'
    unpaid_tax = f'INDEX(Cash_Flow!$O$5:$O$130,{gap})'
    unpaid_spending = f'INDEX(Cash_Flow!$Q$5:$Q$130,{gap})'
    cpi = price_index_at(gap)
    context = '=IFERROR("Scenario: "&Inputs!$B$7&"   •   Opening year: "&TEXT(Inputs!$B$130,"0"),"Choose a scenario on Inputs")'
    dollars = '"$#,##0"'
    first_gap = display_dollars(f'({unpaid_tax}+{unpaid_spending})/{cpi}', cpi)
    end_cpi = price_index_at('Report_Data!$Q$2')
    cents = '"$#,##0.00"'

    summary('B2', '1 / Retirement outlook', 'report_kicker')
    summary('B3', 'Your retirement plan', 'report_title')
    summary('B5', context, 'report_caption')
    summary('B6', f'=IF({ok},"Retire at age "&Inputs!$B$10&"   •   Plan through age "&{end_age}&"   •   Prepared "&TEXT(TODAY(),"mmm d, yyyy"),"Review the item below, then return to Inputs.")', 'report_caption')
    summary('B8', f'=IF({ok},IF({gap}=0,"Projected spending and taxes are covered through age "&{end_age}&".","First shortfall at age "&{first_age}&" ("&{METRICS_REF}$B$26&")."),"Review the plan inputs or calculation")', 'report_result')
    summary('B9', f'=IF({ok},IF({gap}=0,"Based on the assumptions in this scenario.",IF({first_gap}<0.01,"Less than $0.01",TEXT({first_gap},{cents}))&" in "&IF({unpaid_tax}>0,IF({unpaid_spending}>0,"spending and taxes","taxes"),"spending")&" goes unpaid in that year."),IF({METRICS_REF}$B$25="Check funding calculation","Review the year flagged on Checks.",{METRICS_REF}$B$25))', 'report_body')
    summary('B11', f'=IF({ok},"Estimated portfolio at age "&{end_age},"")', 'report_caption')
    summary('B12', f'=IF({ok},{display_dollars(f"{METRICS_REF}$B$29", end_cpi)},"")', 'report_amount')
    summary('F11', f'=IF({ok},IF({risk.CURRENT},"Success rate with fixed spending","Market risk"),"")', 'report_caption')
    summary('F12', risk.display(risk.result('AF8')), 'report_percent_amount')
    summary('F13', risk.display(
        f'TEXT({risk.success_count("AF8")},"#,##0")&" of "&TEXT({risk.result("AF5")},"#,##0")&" paths paid all costs"',
        f'IF({opening_only},"Extend the plan to run market risk",'
        f'IF({ok},IF({risk.STATUS}="stale","Refresh market risk","Run market risk"),""))',
    ), 'report_caption')
    restricted = display_dollars(f'{METRICS_REF}$B$6', end_cpi)
    summary('B13', f'=IF(N({METRICS_REF}$B$6)>0,TEXT({restricted},{dollars})&" in unavailable workplace Roth","")', 'report_caption')
    summary('B29', f'=IF({FUTURE_VIEW},"Future dollars use each year’s prices.","Today’s dollars use opening-year prices.")&" Balances are before estimated tax on traditional savings."', 'report_caption')
    summary('B30', risk.display(
        '"The shaded range shows the middle 80% of simulated balances at each age."'
    ), 'report_caption')
    summary('B31', risk.display(
        f'"Market risk run "&{risk.result("AF2")}&IF({risk.result("AF16")}=1,""," • "&{risk.volatility_notice()})'
    ), 'report_caption')

    summary('B33', '="2 / Retirement cash flow   •   "&Inputs!$B$7', 'report_kicker')
    summary('B34', 'How you’ll pay for retirement', 'report_title')
    summary('B36', f'=IF(AND({ok},SUM(Projection!$N$6:$N$130)>0),"Social Security income starts at age "&MAX(Inputs!$B$45,Inputs!$B$8+1)&" in this projection.","")', 'report_caption')
    summary('B38', f'=IF(Report_Data!$Q$4>0,"Income and savings used to pay bills around retirement.",IF({ok},"This projection ends before the first modeled retirement year.","Cash flow is available after you resolve the item on page 1."))')
    summary('B39', f'=IF({FUTURE_VIEW},"Future dollars show what you expect to pay or receive in each year.","Today’s dollars show opening-year purchasing power. The forecast includes inflation.")', 'report_caption')
    for column, label in zip('BCDEFGHI', CASH_HEADERS):
        summary(f'{column}40', label, 'report_table_head')
    for row in range(FIRST_TRANSITION_ROW, LAST_TRANSITION_ROW + 1):
        offset = row - FIRST_TRANSITION_ROW
        position = f'(Report_Data!$Q$3+{offset})'
        for index, (column, source) in enumerate(zip('BCDEFGHI', ('E', 'A', 'F', 'K', 'L', 'M', 'N', 'I'))):
            style = 'report_row_alt' if index < 2 else 'report_money_alt'
            if row % 2 == 0:
                style = 'report_row' if index < 2 else 'report_money'
            amount = f'INDEX(Report_Data!${source}$2:${source}$127,{position})'
            if index >= 2:
                amount = display_dollars(amount, price_index_at(position))
            summary(f'{column}{row}', f'=IF({offset}<Report_Data!$Q$4,{amount},"")', style)
    summary('B55', 'Income used can include pay after pre-tax savings, Social Security, other cash income and Marketplace credits.', 'report_caption')
    summary('B56', 'Savings used includes withdrawals and HSA payments. Tax due includes any carryover.', 'report_caption')
    bridge = f'AND({ok},ISNUMBER({METRICS_REF}$B$19))'
    summary('B57', f'=IF({bridge},"Early retirement access","")', 'report_section')
    bridge_cpi = price_index_at(retirement)
    bridge_amount = display_dollars(f'{METRICS_REF}$B$19/{bridge_cpi}', bridge_cpi)
    summary('B58', f'=IF({bridge},"Available in the first retirement year: "&TEXT({bridge_amount},{dollars})&" in brokerage and Roth funds this plan can draw on.","")')
    summary('B59', f'=IF({bridge},"Use Roth Access to check when additional funds become available.","")', 'report_caption')

    summary('B65', '="3 / Risk and decisions   •   "&Inputs!$B$7', 'report_kicker')
    summary('B66', 'What could change the outcome', 'report_title')
    summary('B68', risk.display('"Funding in three market paths"', '"Key plan decisions"'), 'report_section')
    summary('B69', risk.display(
        '"Actual simulations ranked by ending balance, with fixed spending."',
        f'IF({opening_only},"Set an end age after the opening age to run market risk.",'
        f'IF({ok},{risk.status_message()},"Resolve the plan item on page 1 to see funding results."))',
    ), 'report_caption')
    for column, label, plan_label in zip(
        'BEGH',
        ('Market outcome', 'Ending portfolio', 'First shortfall age', 'Unpaid that year'),
        ('Plan decision', 'Your setting', 'Unit', 'When it applies'),
    ):
        summary(f'{column}70', risk.display(f'"{label}"', f'"{plan_label}"'), 'report_table_head')
    choices = (
        ('Retirement timing', 'TEXT(Inputs!$B$10,"0")', 'Age', 'Retirement begins'),
        ('Retirement spending', 'Inputs!$B$37', 'Per year', 'Before healthcare'),
        ('Roth IRA savings goal', 'Inputs!$B$35', 'Per year', 'While working'),
    )
    for row, ((label, case_row), (choice, setting, unit, timing)) in enumerate(zip(risk.CASES, choices), 71):
        amount = risk.result(f'AI{case_row}')
        age = risk.result(f'AH{case_row}')
        path_balance = display_dollars(risk.result(f'AG{case_row}'), end_cpi)
        path_amount = display_dollars(amount, price_index_at(f'{age}-Inputs!$B$8+1'))
        style = 'report_row_alt' if row % 2 else 'report_row'
        money = 'report_money_alt' if row % 2 else 'report_money'
        summary(f'B{row}', risk.display(f'"{label}"', f'"{choice}"'), style)
        summary(f'E{row}', risk.display(path_balance, f'IF({ok},{setting},"")'), money)
        summary(f'G{row}', risk.display(risk.age_or_none(age), f'"{unit}"'), style)
        summary(f'H{row}', risk.display(f'IF(ISNUMBER({amount}),{risk.shortfall(path_amount)},"")', f'"{timing}"'), style)
    failures = f'({risk.result("AF5")}-{risk.success_count("AF8")})'
    median_age = risk.result('AF18')
    summary('B74', risk.display(
        f'IF(ISNUMBER({median_age}),TEXT({failures},"#,##0")&IF({failures}=1," path"," paths")'
        f'&" had unpaid costs. Median first shortfall age: "'
        f'&IF(MOD({median_age},1)=0,TEXT({median_age},"0"),TEXT({median_age},"0.0"))&".",'
        '"All fixed-spending paths paid their required costs.")'
    ), 'report_caption')
    summary('B75', risk.display(
        f'IF({FUTURE_VIEW},"Future dollars.","Today’s dollars.")&" See Risk Results for the income and accounts behind each shortfall."',
        '"Spending and savings inputs are in today’s dollars."',
    ), 'report_caption')
    summary('B76', 'Would temporary spending cuts help?', 'report_section')
    rate = lambda cell: f'TEXT({risk.result(cell)},{risk.PERCENT})'
    fixed_rate, flexible_rate = (f'TEXT({risk.result(cell)},{risk.SUCCESS_PERCENT})' for cell in ('AF8', 'AF9'))
    summary('B77', risk.display(
        f'"Paid all costs: "&{fixed_rate}&" with fixed spending and "&{flexible_rate}&" with flexible spending."',
        f'IF({ok},"Decide how much recurring retirement spending you could cut in a downturn.","")',
    ))
    cuts = risk.result('AF10')
    summary('B78', risk.display(
        f'IF({cuts}>0,TEXT({cuts},"#,##0")&" of "&TEXT({risk.result("AF5")},"#,##0")'
        f'&" paths used cuts. Median number of years with cuts: "&TEXT({risk.result("AF11")},"0.0")&".",'
        '"No paths used a spending cut.")'
    ), 'report_caption')
    summary('B79', risk.display(
        f'"The rule cuts "&{rate("AF13")}&" of flexible living costs when the opening balance falls below "'
        f'&{rate("AF12")}&" of plan."'
    ), 'report_caption')
    summary('B81', 'Early stock losses', 'report_section')
    stress_rate, years = risk.result('AF28'), risk.result('AF17')
    fixed_age, flex_age = risk.result('AG28'), risk.result('AI28')
    summary('B82', risk.display(
        f'IF(AND({years}>0,ISNUMBER({stress_rate}),{stress_rate}<0),"Lowest tested stock return: "&'
        f'TEXT({stress_rate},{risk.PERCENT})&" a year for the first "&{years}&IF({years}=1," retirement year"," retirement years")'
        f'&". First shortfall: "&'
        f'IF(ISNUMBER({fixed_age}),"age "&{fixed_age},"none")&" with fixed spending and "&'
        f'IF(ISNUMBER({flex_age}),"age "&{flex_age},"none")&" with flexible spending.",'
        '"No early retirement stock losses were tested for this scenario.")',
        f'IF({ok},"Risk Results shows how the plan handles stock losses early in retirement.","")',
    ))
    target = f'AND(ISNUMBER({METRICS_REF}$B$16),ISNUMBER({METRICS_REF}$B$8))'
    benchmark_cpi = price_index_at('MAX(1,Inputs!$B$10-Inputs!$B$8)')
    projected = display_dollars(f'{METRICS_REF}$B$16', benchmark_cpi)
    benchmark = display_dollars(f'{METRICS_REF}$B$8', benchmark_cpi)
    summary('B84', f'=IF(AND({ok},{target}),"Savings at retirement ("&IF({FUTURE_VIEW},"future","today’s")&" dollars): "&TEXT({projected},{dollars})&" projected and "&TEXT({benchmark},{dollars})&" targeted.","")', 'report_caption')
    summary('B86', 'Next decision', 'report_section')
    summary('B87', f'=IF({ok},IF({gap}>0,"Try a later retirement age, lower spending or higher savings. Change one input and compare the first shortfall.",IF({risk.CURRENT},"Choose one change in spending, retirement age or savings. Refresh market risk and compare the funding results.",IF({opening_only},"Extend the plan on Inputs before testing market risk.","Run market risk, then test a change in spending, retirement age or savings."))),"Resolve the item on page 1 before comparing scenarios.")')
    summary('B90', 'Change to test', 'report_caption')
    for column in 'DEFGHI':
        summary(f'{column}90', '', 'report_write_in')

    summary('K3', 'Plan navigation', 'report_section')
    for address, label in {
        'K5': 'Edit assumptions', 'K6': 'Retirement cash flow',
        'K7': 'Risk and decisions', 'K8': 'How to print',
        'K10': 'Annual detail', 'K11': 'Appendix and model details',
        'K12': 'Market risk results', 'K13': 'Annual cash flow', 'K14': 'Account access',
        'K15': 'Market risk settings', 'K30': 'Input decisions', 'K31': 'Spending decisions',
        'K32': 'Investment mix', 'K33': 'Roth conversion comparison',
        'K34': 'Planning metrics', 'K35': 'Tax rates and limits',
    }.items():
        summary(address, label, 'report_link')
    summary('K17', 'Dollar view', 'report_section')
    summary(DOLLAR_VIEW_CELL, TODAY_DOLLARS, 'report_active')
    summary('K19', 'Select a view before printing.', 'report_caption')
    summary('K20', 'Today’s dollars show what', 'report_caption')
    summary('K21', 'money can buy today.', 'report_caption')
    summary('K22', 'Future dollars use prices', 'report_caption')
    summary('K23', 'in each projected year.', 'report_caption')
    summary('K25', 'Print the plan', 'report_section')
    summary('K26', 'Choose File → Print.', 'report_caption')
    summary('K27', 'Use Active Sheets with only', 'report_caption')
    summary('K28', 'Summary selected.', 'report_caption')

    appendix('B2', 'Annual projection', 'report_title')
    appendix('L2', 'Back to Summary', 'report_link')
    appendix('B3', context, 'report_caption')
    appendix('B4', f'=IF({FUTURE_VIEW},"Future dollars","Today’s dollars")&" from the view on Summary. The first row gives opening balances. Tax due includes carryover."', 'report_caption')
    for column, label in zip('BCDEFGHIJ', (*CASH_HEADERS, 'Portfolio')):
        appendix(f'{column}5', label, 'report_table_head')
    for row in range(FIRST_APPENDIX_ROW, LAST_APPENDIX_ROW + 1):
        source = row - FIRST_APPENDIX_ROW + 5
        for index, (column, amount) in enumerate(zip('BCDEFGHIJ', _annual_amounts(source))):
            style = 'report_row' if index < 2 else 'report_money'
            appendix(f'{column}{row}', f'=IF({ok},IF(Report_Data!$Q$2>={source - 4},{amount},""),"")', style)


def index_rows():
    row = 8
    for title, schedules in DETAIL_GROUPS:
        yield row, title, schedules
        row += len(schedules) + 2


def configure_index(put):
    put(INDEX, 'B3', 'Appendix', 'report_title')
    put(INDEX, 'B5', 'Inspect the schedules, source data and calculations behind your plan.', 'report_body')
    put(INDEX, 'B6', 'Print Summary for the five-page report. Print Annual Detail separately for the full timeline.', 'report_caption')
    for row, title, schedules in index_rows():
        put(INDEX, f'B{row}', title, 'report_section')
        for offset, (_, label, description) in enumerate(schedules, 1):
            put(INDEX, f'B{row + offset}', label, 'report_link')
            put(INDEX, f'E{row + offset}', description, 'report_caption')
    put(INDEX, f'B{INDEX_LINK_ROW}', 'Back to Summary', 'report_link')
    put(INDEX, f'E{INDEX_LINK_ROW}', 'Open Annual Detail', 'report_link')
