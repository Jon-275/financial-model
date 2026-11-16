"""Build the reader-facing plan from the existing annual schedules."""

import re

from openpyxl.formula.tokenizer import Tokenizer

SUMMARY = 'Summary'
APPENDIX = 'Annual Detail'
INDEX = 'Appendix'
DETAIL_GROUPS = (
    ('Annual schedules', (
        ('Cash_Flow', 'Annual cash flow', 'Income, tax, spending, withdrawals and unpaid costs.'),
        ('Projection', 'Full projection', 'Year-by-year income, assets and portfolio balances.'),
        ('Roth_Access', 'Roth account access', 'Balances and available withdrawals by owner.'),
        ('Healthcare', 'Healthcare', 'Coverage, premiums and credits by person.'),
    )),
    ('Planning tools', (
        ('Asset_Allocation', 'Investment mix', 'Working and retirement investment assumptions.'),
        ('Income_Targeting', 'Roth conversion comparison', 'Compare conversion amounts for a selected year.'),
        ('Risk_Settings', 'Market risk settings', 'Inputs for the separate market-risk report.'),
    )),
    ('Rates and calculation checks', (
        ('Tax_Tables', 'Tax tables', 'Published amounts and the years they apply to.'),
        ('Checks', 'Model checks', 'Reconcile cash, taxes, accounts and healthcare.'),
        ('Report_Data', 'Report data', 'Source values for the plan and its chart.'),
        ('Funding_Calc', 'Withdrawal calculation', 'The funding trials for each projected year.'),
        ('Targeting_Calc', 'Conversion calculation', 'The funding trials for the conversion comparison.'),
    )),
)
DETAIL = 'Report_Data'
DETAIL_COLUMNS = dict(zip('ABC', 'MNO'))
FIRST_TRANSITION_ROW = 41
LAST_TRANSITION_ROW = 53
FIRST_APPENDIX_ROW = 6
LAST_APPENDIX_ROW = 131
CASH_HEADERS = ('Year', 'Age', 'Income used', 'Savings used', 'Other spending',
                'Healthcare', 'Tax due', 'Unpaid')


def _detail_address(address):
    return DETAIL_COLUMNS[address[0]] + address[1:]


def move_details(template, updates, styles, put):
    """Retain the detailed results on the visible schedule that feeds the report."""
    for row in template.iter_rows(max_col=3):
        for cell in row:
            updates.setdefault(cell.coordinate, cell.value)
    for address, value in updates.items():
        if value is None:
            continue
        if isinstance(value, str) and value.startswith('='):
            tokens = Tokenizer(value)
            for token in tokens.items:
                if token.type == 'OPERAND' and token.subtype == 'RANGE':
                    token.value = re.sub(
                        r'(?<![!\w])(\$?)([ABC])(\$?\d+)',
                        lambda match: match[1] + DETAIL_COLUMNS[match[2]] + match[3],
                        token.value,
                    ) if '!' not in token.value else token.value
            value = tokens.render()
        style = styles.get(address) or ('label' if address[0] == 'A'
                                         else 'note' if address[0] == 'C' else None)
        if address in ('B8', 'B9', 'B15', 'B16', 'B17', 'B18', 'B19', 'B24', 'B27',
                       'B28', 'B29', 'B30', 'B31', 'B32', 'B38', 'B40', 'B42'):
            style = 'summary_money'
        elif address in ('B12', 'B39', 'B41'):
            style = 'percent_active'
        put(DETAIL, _detail_address(address), value, style)
    updates.clear()
    styles.clear()


def _annual_amounts(row):
    cash = lambda col: f'Cash_Flow!{col}{row}'
    proj = lambda col: f'Projection!{col}{row}'
    paid = f'({cash("N")}+{cash("P")})'
    return (
        cash('A'), cash('B'),
        f'MIN({cash("C")},{paid})/{proj("G")}',
        f'(MAX(0,{paid}-{cash("C")})+{cash("E")})/{proj("G")}',
        f'({cash("D")}-{proj("BA")})/{proj("G")}',
        f'{proj("BA")}/{proj("G")}',
        f'({cash("F")}+{cash("M")})/{proj("G")}',
        f'{cash("V")}/{proj("G")}',
        proj('BN'),
    )


def configure(put):
    def summary(address, value, style='report_body'):
        put(SUMMARY, address, value, style)

    def appendix(address, value, style='report_body'):
        put(APPENDIX, address, value, style)

    ok = 'Report_Data!$K$2>0'
    end_age = 'Inputs!$B$53'
    retirement = 'Report_Data!$K$5'
    gap = 'Report_Data!$K$6'
    first_age = f'INDEX(Projection!$C$5:$C$130,{gap})'
    unpaid_tax = f'INDEX(Cash_Flow!$O$5:$O$130,{gap})'
    unpaid_spending = f'INDEX(Cash_Flow!$Q$5:$Q$130,{gap})'
    cpi = f'INDEX(Projection!$G$5:$G$130,{gap})'
    context = '=IFERROR("Scenario: "&Inputs!$B$7&"   •   Opening year: "&TEXT(Inputs!$B$130,"0"),"Choose a scenario on Inputs")'
    dollars = '"$#,##0"'
    first_gap = f'({unpaid_tax}+{unpaid_spending})/{cpi}'
    cents = '"$#,##0.00"'

    summary('B2', '1 / Retirement outlook', 'report_kicker')
    summary('B3', 'Your retirement plan', 'report_title')
    summary('B5', context, 'report_caption')
    summary('B6', f'=IF({ok},"Retire at age "&Inputs!$B$10&"   •   Plan through age "&{end_age}&"   •   Prepared "&TEXT(TODAY(),"mmm d, yyyy"),"Review the item below, then return to Inputs.")', 'report_caption')
    summary('B8', f'=IF({ok},IF({gap}=0,"Projected spending and taxes are covered through age "&{end_age}&".","First shortfall at age "&{first_age}&" ("&Report_Data!$N$26&")."),"Review the plan inputs or calculation")', 'report_result')
    summary('B9', f'=IF({ok},IF({gap}=0,"Based on the assumptions in this scenario.",IF({first_gap}<0.01,"Less than $0.01",TEXT({first_gap},{cents}))&" in "&IF({unpaid_tax}>0,IF({unpaid_spending}>0,"spending and taxes","taxes"),"spending")&" goes unpaid in that year."),IF(Report_Data!$N$25="Check funding calculation","Review the year flagged on Checks.",Report_Data!$N$25))', 'report_body')
    summary('B11', f'=IF({ok},"Estimated portfolio at age "&{end_age},"")', 'report_caption')
    summary('B12', f'=IF({ok},Report_Data!$N$29,"")', 'report_amount')
    summary('F11', '=IF(N(Report_Data!$N$6)>0,"Includes workplace Roth unavailable at the end age","")', 'report_caption')
    summary('F12', '=IF(N(Report_Data!$N$6)>0,Report_Data!$N$6,"")', 'report_amount')
    summary('B29', 'Portfolio values are in today’s dollars, before estimated tax on traditional savings.', 'report_caption')
    summary('B30', f'=IF({ok},IF({retirement}>Report_Data!$K$2,"Retirement falls after the end of this projection.","Guide lines show retirement and the first shortfall when these occur in the plan."),"The chart will populate when the plan can be calculated.")', 'report_caption')

    summary('B33', '="2 / Retirement cash flow   •   "&Inputs!$B$7', 'report_kicker')
    summary('B34', 'How you’ll pay for retirement', 'report_title')
    summary('B38', f'=IF(Report_Data!$K$4>0,"Income and savings used to pay bills around retirement.",IF({ok},"This projection ends before the first modeled retirement year.","Cash flow is available after you resolve the item on page 1."))')
    summary('B39', 'Annual amounts in today’s dollars (opening-year prices)', 'report_caption')
    for column, label in zip('BCDEFGHI', CASH_HEADERS):
        summary(f'{column}40', label, 'report_table_head')
    for row in range(FIRST_TRANSITION_ROW, LAST_TRANSITION_ROW + 1):
        data_row = row - FIRST_TRANSITION_ROW + 2
        for index, (column, source) in enumerate(zip('BCDEFGHI', ('U', 'E', 'F', 'V', 'W', 'X', 'Y', 'I'))):
            style = 'report_row_alt' if index < 2 else 'report_money_alt'
            if row % 2 == 0:
                style = 'report_row' if index < 2 else 'report_money'
            summary(f'{column}{row}', f'=IF(Report_Data!$J{data_row}="","",Report_Data!{source}{data_row})', style)
    summary('B55', 'Income includes any Marketplace credit. Savings includes withdrawals and health savings. Tax due includes any carryover.', 'report_caption')
    summary('B56', 'Working-year healthcare is part of other spending.', 'report_caption')
    bridge = f'AND({ok},ISNUMBER(Report_Data!$N$19))'
    summary('B57', f'=IF({bridge},"Early retirement access","")', 'report_section')
    summary('B58', f'=IF({bridge},"Available in the first retirement year: "&TEXT(Report_Data!$N$19/INDEX(Projection!$G$5:$G$130,{retirement}),{dollars})&" in brokerage and Roth funds this plan can draw on.","")')
    summary('B59', f'=IF({bridge},"Use Roth Access to check when additional funds become available.","")', 'report_caption')

    summary('B65', '="3 / Assumptions and decisions   •   "&Inputs!$B$7', 'report_kicker')
    summary('B66', 'What drives this plan', 'report_title')
    summary('B70', 'Key assumptions', 'report_section')
    assumptions = (
        ('Annual pay', 'Inputs!$B$12', 'report_money', 'Annual pay increase', 'Inputs!$B$13', 'report_percent'),
        ('Annual spending now', 'Inputs!$B$36', 'report_money', 'Inflation', 'Inputs!$B$14', 'report_percent'),
        ('Retirement spending¹', 'Inputs!$B$37', 'report_money', 'Working years return²', 'Inputs!$B$23', 'report_percent'),
        ('Annual Roth IRA savings goal', 'Inputs!$B$35', 'report_money', 'Retirement return²', 'Inputs!$B$24', 'report_percent'),
        ('Workplace and health savings³', 'SUM(Inputs!$B$32:$B$34)', 'report_money', 'Social Security claim age', 'Inputs!$B$45', 'report_row'),
    )
    for row, (left, left_value, left_style, right, right_value, right_style) in enumerate(assumptions, 72):
        summary(f'B{row}', left)
        summary(f'E{row}', f'=IF({ok},{left_value},"")', left_style)
        summary(f'F{row}', right, 'report_body_gutter')
        summary(f'I{row}', f'=IF({ok},{right_value},"")', right_style)
    summary('B77', '¹ Before healthcare.  ² Blended estimates at today’s account mix.  ³ Includes employer contributions.', 'report_caption')
    summary('B79', 'Savings benchmark', 'report_section')
    target = 'AND(ISNUMBER(Report_Data!$N$16),ISNUMBER(Report_Data!$N$8))'
    summary('B80', f'=IF({ok},IF({target},"At retirement: "&TEXT(Report_Data!$N$16,{dollars})&" projected savings and a "&TEXT(Report_Data!$N$8,{dollars})&" savings target.","A retirement balance is unavailable for this timeline."),"")')
    summary('B81', 'The target is a benchmark based on spending, healthcare and the withdrawal rate in Inputs.', 'report_caption')
    summary('B83', 'Next decision', 'report_section')
    summary('B84', f'=IF({ok},IF({gap}>0,"Try a later retirement age, lower spending or higher savings. Change one input and compare the first shortfall.","Test how a change in spending, retirement age or market returns affects the result."),"Resolve the item on page 1 before comparing scenarios.")')
    for row, label in ((86, 'Scenario chosen'), (88, 'Change to test'), (90, 'Next review date')):
        summary(f'B{row}', label, 'report_caption')
        for column in 'DEFGHI':
            summary(f'{column}{row}', '', 'report_write_in')

    summary('B94', 'Select Inputs, hold Shift and select Summary. Choose File → Print → Active Sheets to print the full plan.', 'report_caption')
    summary('K3', 'Plan navigation', 'report_section')
    for address, label in {
        'K5': 'Edit assumptions', 'K6': 'Retirement cash flow',
        'K7': 'Assumptions and decisions', 'K8': 'How to print',
        'K10': 'Annual detail', 'K11': 'Appendix and model details',
        'K12': 'Market risk settings', 'K13': 'Annual cash flow', 'K14': 'Account access',
    }.items():
        summary(address, label, 'report_link')

    appendix('B2', 'Annual projection', 'report_title')
    appendix('L2', 'Back to Summary', 'report_link')
    appendix('B3', context, 'report_caption')
    appendix('B4', 'Opening-year dollars. The first row gives opening balances. Tax due includes any carryover.', 'report_caption')
    for column, label in zip('BCDEFGHIJ', (*CASH_HEADERS, 'Portfolio')):
        appendix(f'{column}5', label, 'report_table_head')
    for row in range(FIRST_APPENDIX_ROW, LAST_APPENDIX_ROW + 1):
        source = row - FIRST_APPENDIX_ROW + 5
        for index, (column, amount) in enumerate(zip('BCDEFGHIJ', _annual_amounts(source))):
            style = 'report_row_alt' if index < 2 else 'report_money_alt'
            if row % 2 == 0:
                style = 'report_row' if index < 2 else 'report_money'
            appendix(f'{column}{row}', f'=IF({ok},IF(Report_Data!$K$2>={source - 4},{amount},""),"")', style)


def index_rows():
    row = 8
    for title, schedules in DETAIL_GROUPS:
        yield row, title, schedules
        row += len(schedules) + 2


def configure_index(put):
    put(INDEX, 'B3', 'Appendix', 'report_title')
    put(INDEX, 'B5', 'Inspect the schedules behind your plan or open an advanced planning tool.', 'report_body')
    put(INDEX, 'B6', 'Print Inputs and Summary for the plan. Add Annual Detail for the full timeline.', 'report_caption')
    for row, title, schedules in index_rows():
        put(INDEX, f'B{row}', title, 'report_section')
        for offset, (_, label, description) in enumerate(schedules, 1):
            put(INDEX, f'B{row + offset}', label, 'report_link')
            put(INDEX, f'E{row + offset}', description, 'report_caption')
    put(INDEX, 'B28', 'Back to Summary', 'report_link')
    put(INDEX, 'E28', 'Open Annual Detail', 'report_link')


def add_transition_values(put, report_row, cash, cpi):
    """Attach the uses of cash to the existing retirement funding schedule."""
    position = f'$J{report_row}'
    healthcare = f'INDEX(Projection!$BA$5:$BA$130,{position})'
    for column, expression in {
        'U': cash('A'),
        'V': f'G{report_row}+H{report_row}',
        'W': f'({cash("D")}-{healthcare})/{cpi}',
        'X': f'{healthcare}/{cpi}',
        'Y': f'({cash("F")}+{cash("M")})/{cpi}',
    }.items():
        put(DETAIL, f'{column}{report_row}', f'=IF({position}="","",{expression})',
            'whole' if column == 'U' else 'money')
