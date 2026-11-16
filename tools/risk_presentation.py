"""Lay out the market-risk results and provide formulas for the printed plan."""

SHEET = 'Risk Results'
STATUS = f"'{SHEET}'!$Z$1"
CURRENT = f'{STATUS}="current"'
LAST_BALANCE_ROW = 183
CASES = (('Low', 21), ('Median', 22), ('High', 23))
PERCENT = '"0.0%"'
SUCCESS_PERCENT = '"0.00%"'
CENTS = '"$#,##0.00"'


def result(address):
    return f"'{SHEET}'!{address}"


def display(value, fallback='""'):
    return f'=IF({CURRENT},{value},{fallback})'


def status_message():
    return (
        f'IF({STATUS}="stale","Refresh market risk after changing assumptions.",'
        '"Run market risk to see results for this scenario.")'
    )


def shortfall(amount):
    return f'IF({amount}<0.01,"Less than $0.01",TEXT({amount},{CENTS}))'


def age_or_none(age):
    return f'IF(ISNUMBER({age}),{age},"None")'


def success_count(rate):
    return f'ROUND({result(rate)}*{result("AF5")},0)'


def volatility_notice():
    setting = result('AF16')
    return (f'IF({setting}=1,"",IF({setting}=0,"No market variation in this run.",'
            f'"Market variation is "&{setting}&" times the workbook setting."))')


def configure(put):
    """Show the evidence behind the headline on a visible workbook tab."""
    def cell(address, value, style='report_body'):
        put(SHEET, address, value, style)

    cell('B2', 'Market risk results', 'report_title')
    cell('B3', display(
        '"Scenario "&AF3&" • Through age "&AF4&" • Run "&AF2',
        status_message(),
    ), 'report_caption')
    cell('B4', display(
        '"Balances and costs are in "&AF15&" dollars. Portfolio balances are before estimated tax on traditional savings."',
        '"Open Summary for the current plan. Save your changes before refreshing the simulation."',
    ), 'report_caption')
    cell('B5', display(volatility_notice()), 'report_caption')
    cell('K2', 'Back to Summary', 'report_link')
    cell('K3', 'Market risk settings', 'report_link')

    for col, name, rate in (('B', 'Fixed spending', 'AF8'), ('F', 'Flexible spending', 'AF9')):
        cell(f'{col}6', display(f'"{name}"'), 'report_caption')
        cell(f'{col}7', display(rate), 'report_percent_amount')
        cell(f'{col}8', display(
            f'TEXT({success_count(rate)},"#,##0")&" of "&TEXT(AF5,"#,##0")&" paths paid all costs"'
        ), 'report_caption')

    cell('B10', 'Funding in three market paths', 'report_section')
    cell('B11', 'Actual fixed-spending paths ranked by the balance at the selected end age.', 'report_caption')
    cell('B12', 'Funding measure', 'report_table_head')
    for col, (label, _) in zip('FHJ', CASES):
        cell(f'{col}12', label, 'report_table_head')
    measures = (
        ('Ending portfolio', 'AG'),
        ('First shortfall age', 'AH'),
        ('Costs unpaid in that year', 'AI'),
        ('Spending unpaid', 'AJ'),
        ('Taxes unpaid', 'AK'),
        ('Spending due', 'AL'),
        ('Taxes due', 'AM'),
        ('Income available', 'AN'),
        ('Income used', 'AO'),
        ('Health savings used', 'AP'),
        ('Health costs eligible for health savings', 'AQ'),
        ('Brokerage available', 'AR'),
        ('Brokerage used', 'AS'),
        ('Traditional account available under plan rules', 'AT'),
        ('Traditional account used', 'AU'),
        ('Roth available under plan rules', 'AV'),
        ('Roth used', 'AW'),
    )
    for row, (label, field) in enumerate(measures, 13):
        cell(f'B{row}', label, 'report_row_alt' if row % 2 else 'report_row')
        for col, (_, case) in zip('FHJ', CASES):
            source = f'{field}{case}'
            if field == 'AG':
                value = source
            elif field == 'AH':
                value = age_or_none(source)
            else:
                amount = shortfall(source) if field == 'AI' else source
                value = f'IF(ISNUMBER({source}),{amount},"")'
            style = 'report_row' if field == 'AH' else 'report_cost' if field in ('AJ', 'AK') else 'report_money'
            cell(f'{col}{row}', display(value), f'{style}_alt' if row % 2 else style)
    cell('B30', '"None" means the path paid every year’s required spending and taxes.', 'report_caption')

    cell('B32', 'Spending flexibility', 'report_section')
    cell('B33', display(
        'IF(AF10>0,TEXT(AF10,"#,##0")&" of "&TEXT(AF5,"#,##0")'
        '&" paths used cuts. Median number of years with cuts: "&TEXT(AF11,"0.0")&".",'
        '"No paths used a spending cut.")'
    ))
    cell('B34', display(
        f'"When the opening balance falls below "&TEXT(AF12,{PERCENT})&" of plan, cut "&TEXT(AF13,{PERCENT})'
        f'&" from the flexible part. That part is "&TEXT(AF14,{PERCENT})&" of recurring retirement living costs."'
    ))

    cell('B36', 'Early retirement stock tests', 'report_section')
    cell('B37', display(
        'IF(AF17>0,"The stock return applies for "&AF17&IF(AF17=1," year"," years")'
        '&" from retirement. Other assets use the plan’s assumed returns.",'
        '"There are no early retirement stress years in this result.")'
    ), 'report_caption')
    for col, label in zip(('B', 'D', 'F', 'H', 'J'), (
        'Annual stock return', 'Fixed first shortfall age', 'Fixed costs unpaid',
        'Flexible first shortfall age', 'Flexible costs unpaid',
    )):
        cell(f'{col}39', label, 'report_table_head')
    for row, source_row in enumerate(range(28, 41), 40):
        valid = f'ISNUMBER(AF{source_row})'
        values = (
            f'AF{source_row}', age_or_none(f'AG{source_row}'),
            f'IF(ISNUMBER(AH{source_row}),{shortfall(f"AH{source_row}")},"")',
            age_or_none(f'AI{source_row}'),
            f'IF(ISNUMBER(AJ{source_row}),{shortfall(f"AJ{source_row}")},"")',
        )
        for col, value in zip(('B', 'D', 'F', 'H', 'J'), values):
            style = 'report_percent' if col == 'B' else 'report_row_alt' if row % 2 else 'report_row'
            cell(f'{col}{row}', display(f'IF({valid},{value},"")'), style)

    cell('B54', 'Fixed-spending balances by age', 'report_section')
    cell('B55', 'The low and high values mark the middle 80% of simulated balances at each age.', 'report_caption')
    for col, label in zip(('B', 'D', 'F', 'H', 'J'), ('Age', 'Base plan', 'Low', 'Median', 'High')):
        cell(f'{col}57', label, 'report_table_head')
    for row, source_row in enumerate(range(46, 172), 58):
        for col, source in zip(('B', 'D', 'F', 'H', 'J'), ('AG', 'AH', 'AI', 'AJ', 'AK')):
            style = 'report_row' if col == 'B' else 'report_money'
            if row % 2:
                style += '_alt'
            cell(f'{col}{row}', display(f'IF(ISNUMBER(AG{source_row}),{source}{source_row},"")'), style)
