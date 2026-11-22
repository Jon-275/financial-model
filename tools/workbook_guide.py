"""Add guidance to the decision tools and the schedules they use."""

from risk_engine.scenario import ASSET_ROWS, asset_profile_start


TABLE_GUIDES = {
    'Report_Data': ('AK', 'Report data', (
        'These values feed the Summary chart and cash-flow table.',
        'The source values use opening-year dollars. Chart values follow the Dollar view.',
        'Each modeled year has one row, including the opening year.',
    ), 'Summary!B3'),
    'Funding_Calc': ('AR', 'Withdrawal calculation', (
        'The model tries different withdrawal amounts each year.',
        'A trial includes the tax and Marketplace credit it would produce.',
        'Skipped means the model did not need that trial. See Selected draw for the result.',
        'Cash_Flow shows the chosen result for each year.',
        'Linked sources begin in BD. Clear the trial filter to inspect every source row.',
    ), 'Cash_Flow!A1'),
    'Targeting_Calc': ('AR', 'Conversion calculation', (
        'These are the withdrawal and tax trials for one retirement year.',
        'Each run uses the same starting balances and tests a different conversion.',
        'Skipped means the model did not need that trial. The last column identifies the conversion request.',
        'Linked sources begin in BD. Clear the trial filter to inspect every source row.',
    ), 'Income_Targeting!A1'),
}

TAX_LINKS = {
    'J3': ('Federal income tax', 'A4'),
    'J4': ('California income tax', 'A28'),
    'J5': ('Payroll and contribution limits', 'A49'),
    'J6': ('Marketplace premium credits', 'A66'),
    'J7': ('Medicare surcharges', 'A87'),
    'J8': ('Single-filer tables', 'A141'),
    'J9': ('Published tax-year updates', 'A230'),
}


def configure(put):
    for sheet, (column, title, notes, _) in TABLE_GUIDES.items():
        put(sheet, f'{column}1', title, 'detail_title')
        for row, note in enumerate(notes, 2):
            put(sheet, f'{column}{row}', note, 'report_caption')
        put(sheet, f'{column}7', 'See results', 'report_link')
        put(sheet, f'{column}8', 'View calculation table', 'report_link')

    put('Projection', 'A1', 'Year-by-year projection', 'detail_title')
    put('Projection', 'A2', 'The annual balances feed Cash_Flow, Annual Detail and Summary. Most amounts use future dollars.', 'report_caption')
    put('Healthcare', 'A3', 'Follow coverage and gross cost, then the Marketplace credit, HSA payment and remaining cash cost.', 'report_caption')

    put('Asset_Allocation', 'A1', 'Investment mix', 'detail_title')
    put('Asset_Allocation', 'A2', 'Review the assumptions and choose a profile for working years and retirement on Inputs.', 'report_caption')
    put('Asset_Allocation', 'A3', 'Use the blue cells for edits. Each account’s mix must total 100%.', 'report_caption')
    put('Asset_Allocation', 'A4', 'Expected returns and income', 'detail_band')
    put('Asset_Allocation', 'A14', 'How asset returns move together', 'detail_band')
    put('Asset_Allocation', 'A23', 'Investment profiles by account', 'detail_band')
    put('Asset_Allocation', 'A52', 'Returns and income for each profile', 'detail_band')
    put('Asset_Allocation', 'A68', 'Portfolio mix for the selected retirement profile', 'detail_band')
    put('Asset_Allocation', 'A93', 'Selected profile checks', 'detail_band')
    for row, input_row, label in ((94, 15, 'Working-years profile'), (95, 16, 'Retirement profile')):
        profile = f'Inputs!B{input_row}'
        checks = []
        for number in range(1, 4):
            start = asset_profile_start(number)
            end = start + len(ASSET_ROWS) - 1
            ranges = [f'{column}{start}:{column}{end}' for column in 'BCDE']
            columns = [f'AND(COUNT({cells})={len(ASSET_ROWS)},MIN({cells})>=0,ABS(SUM({cells})-1)<0.000000001)'
                       for cells in ranges]
            checks.append(f'AND({",".join(columns)})')
        valid = f'AND(ISNUMBER({profile}),{profile}=INT({profile}),CHOOSE({profile},{",".join(checks)}))'
        put('Asset_Allocation', f'A{row}', label, 'report_caption')
        put('Asset_Allocation', f'B{row}', f'=IFERROR(IF({valid},"OK","Check mix"),"Check mix")', 'report_active')

    put('Income_Targeting', 'A1', 'Roth conversion comparison', 'detail_title')
    put('Income_Targeting', 'A2', 'Choose a retirement age and enter a conversion in that year’s dollars.', 'report_caption')
    put('Income_Targeting', 'A3', 'Start with zero, then compare tax, healthcare and the cash needed for a larger conversion.', 'report_caption')
    put('Income_Targeting', 'A4', 'Comparison settings', 'detail_band')
    put('Income_Targeting', 'A22', 'Active plan and requested conversion', 'detail_band')
    put('Income_Targeting', 'A33', 'Conversion range', 'detail_band')

    put('Tax_Tables', 'A1', 'Tax rates and limits', 'detail_title')
    put('Tax_Tables', 'A2', 'Federal and California tables use separate schedules. Follow the links to review their source years.', 'report_caption')
    put('Tax_Tables', 'J1', 'Find a rate or limit', 'report_section')
    for address, (label, _) in TAX_LINKS.items():
        put('Tax_Tables', address, label, 'report_link')
    for row, heading in {
        4: 'Federal income tax · Joint · 2026',
        15: 'Federal long-term gains and qualified dividends · Joint · 2026',
        22: 'Federal deductions · Joint · 2026',
        28: 'California income tax · Joint · 2025',
        41: 'California deductions, credits and additional taxes',
        49: 'Federal payroll taxes · 2026',
        141: 'Single-filer tax tables',
    }.items():
        put('Tax_Tables', f'A{row}', heading, 'detail_band')
    import decision_layout
    decision_layout.configure(put)
