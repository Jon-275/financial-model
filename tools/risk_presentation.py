"""Lay out the market-risk results and provide formulas for the printed plan."""

import re

SHEET = 'Risk Results'
DATA_SHEET = 'Risk_Data'
STATUS = f'{DATA_SHEET}!$Z$1'
CURRENT = f'{STATUS}="current"'
NATIVE = f'LEFT({DATA_SHEET}!$Z$2,LEN("retirement-risk-native"))="retirement-risk-native"'
FUTURE_VIEW = 'Summary!$K$18="Future dollars"'
LAST_BALANCE_ROW = 183
CASES = (('Low', 21), ('Median', 22), ('High', 23))
PERCENT = '"0.0%"'
SUCCESS_PERCENT = '"0.00%"'
CENTS = '"$#,##0.00"'


def snapshot_address(address):
    """Place saved annual, stress and representative path results in contiguous tables."""
    match = re.fullmatch(r'(A[F-X])(\d+)', address)
    if match is None:
        return address
    column, row = match[1][1], int(match[2])
    if row >= 46:
        return f'{chr(ord(column) - ord("F") + ord("A"))}{row - 41}'
    if row >= 28:
        return f'{chr(ord(column) - ord("F") + ord("K"))}{row - 23}'
    if row >= 21:
        return f'{match[1]}{row - 16}'
    return f'BB{row}'


def result(address):
    return f'{DATA_SHEET}!{snapshot_address(address)}'


def display(value, fallback='""'):
    return f'=IF({CURRENT},{value},{fallback})'


def status_message():
    return (
        f'IF({STATUS}="running","Market risk is running.",'
        f'IF({STATUS}="stale","Choose Run market risk to update the results.",'
        f'IF({STATUS}="invalid","Review the plan and risk assumptions.",'
        f'IF({STATUS}="cancelled","Run cancelled. Choose Run market risk to try again.",'
        f'IF({STATUS}="failed","Run failed. Review the message on Risk Settings.",'
        f'IF({STATUS}="incomplete","The risk run is incomplete. Run it again.",'
        '"Choose Run market risk to see this scenario’s results."))))))'
    )


def short_status():
    labels=(('stale','Refresh market risk'),('running','Running market risk'),
            ('invalid','Review risk settings'),('cancelled','Run cancelled'),
            ('failed','Run failed'),('incomplete','Run incomplete'))
    expression='"Run market risk"'
    for state,label in reversed(labels):expression=f'IF({STATUS}="{state}","{label}",{expression})'
    return expression


def case_price(case, *, ending=False):
    """A representative path uses its own prices at its event or end age."""
    row=case-16
    column='AZ' if ending else 'AY'
    age=result('AF4') if ending else result(f'AH{case}')
    fallback=f'INDEX(Projection!$G$5:$G$130,{age}-Inputs!$B$8+1)'
    return f'IF({NATIVE},{DATA_SHEET}!${column}${row},{fallback})'


def case_dollars(real_amount,case,*,ending=False):
    return f'({real_amount})*IF({FUTURE_VIEW},{case_price(case,ending=ending)},1)'


def annual_dollars(real_column,nominal_column,row):
    """Choose separately calculated real or nominal percentiles for the chart."""
    real=f'{DATA_SHEET}!{real_column}{row}'
    nominal=f'{DATA_SHEET}!{nominal_column}{row}'
    fallback=f'({real})*INDEX(Projection!$G$5:$G$130,{row-4})'
    return f'IF({FUTURE_VIEW},IF({NATIVE},{nominal},{fallback}),{real})'


def early_dollars(real_amount,age,data_row,*,flexible=False):
    column='Q' if flexible else 'P'
    fallback=f'INDEX(Projection!$G$5:$G$130,{age}-Inputs!$B$8+1)'
    cpi=f'IF({NATIVE},{DATA_SHEET}!{column}{data_row},{fallback})'
    return f'({real_amount})*IF({FUTURE_VIEW},{cpi},1)'


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


def configure(put, display_dollars, price_index_at, future_view):
    """Show the evidence behind the headline on a visible workbook tab."""
    def cell(address, value, style='report_body'):
        put(SHEET, address, value, style)

    def dollars_at(amount, age):
        return display_dollars(amount, price_index_at(f'{age}-Inputs!$B$8+1'))

    cell('B2', 'Market risk details', 'report_title')
    cell('B3', display(
        f'"Scenario "&{result("AF3")}&" • Through age "&{result("AF4")}&" • Run "&{result("AF2")}',
        status_message(),
    ), 'report_caption')
    cell('B4', display(
        f'IF({future_view},"Future dollars","Today’s dollars")&" from Summary. Balances are before estimated tax on traditional savings."',
        '"Choose Run market risk on Summary after reviewing your inputs."',
    ), 'report_caption')
    cell('B5', display(volatility_notice()), 'report_caption')
    cell('K2', 'Back to Summary', 'report_link')
    cell('K3', 'Market risk settings', 'report_link')
    cell('K4', 'Open saved risk data', 'report_link')
    cell('M2', 'Simulation settings', 'report_section')
    for row, label, reference, style in (
        (3, 'Scenario', 'AF3', 'report_body'),
        (4, 'End age', 'AF4', 'report_case_age'),
        (5, 'Run date', 'AF2', 'report_body'),
        (6, 'Simulated paths', 'AF5', 'report_row'),
        (7, 'Seed', 'AF6', 'report_body'),
        (8, 'Engine version', 'AF7', 'report_body'),
        (9, 'Balance trigger', 'AF12', 'report_percent'),
        (10, 'Cut to flexible costs', 'AF13', 'report_percent'),
        (11, 'Flexible share', 'AF14', 'report_percent'),
    ):
        cell(f'M{row}', label, 'report_caption')
        cell(f'N{row}', display(result(reference)), style)

    for col, name, rate in (('B', 'Fixed spending', 'AF8'), ('F', 'Flexible spending', 'AF9')):
        cell(f'{col}6', display(f'"{name}"'), 'report_caption')
        cell(f'{col}7', display(result(rate)), 'report_percent_amount')
        cell(f'{col}8', display(
            f'TEXT({success_count(rate)},"#,##0")&" of "&TEXT({result("AF5")},"#,##0")&" paths paid all costs"'
        ), 'report_caption')

    cell('B10', 'Funding in three market paths', 'report_section')
    cell('B11','Actual fixed-spending paths near the 10th, 50th and 90th ending ranks in today’s dollars.','report_caption')
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
            source = result(f'{field}{case}')
            if field == 'AG':
                value = case_dollars(source,case,ending=True)
            elif field == 'AH':
                value = age_or_none(source)
            else:
                amount = case_dollars(source,case)
                if field == 'AI':
                    amount = shortfall(amount)
                value = f'IF(ISNUMBER({source}),{amount},"")'
            style = 'report_row' if field == 'AH' else 'report_cost' if field in ('AJ', 'AK') else 'report_money'
            cell(f'{col}{row}', display(value), f'{style}_alt' if row % 2 else style)
    cell('B30', '"None" means the path paid every year’s required spending and taxes.', 'report_caption')

    cell('B32', 'Spending flexibility', 'report_section')
    cuts, trials, median = (result(address) for address in ('AF10', 'AF5', 'AF11'))
    cell('B33', display(
        f'IF({cuts}>0,TEXT({cuts},"#,##0")&" of "&TEXT({trials},"#,##0")'
        f'&" paths used cuts. Median across all flexible paths: "&TEXT({median},"0.0")&" years with cuts.",'
        '"No paths used a spending cut.")'
    ))
    cell('B34', display(
        f'"Below "&TEXT({result("AF12")},{PERCENT})&" of the planned opening balance, cut "&TEXT({result("AF13")},{PERCENT})'
        f'&" of the flexible budget. That reduces non-mortgage living costs by "&TEXT({result("AF13")}*{result("AF14")},"0.0%")&"."'
    ))

    cell('B36', 'Early retirement stock tests', 'report_section')
    stress_years = result('AF17')
    cell('B37', display(
        f'IF({stress_years}>0,"The stock return applies for "&{stress_years}&IF({stress_years}=1," year"," years")'
        '&" from retirement. Other assets use the plan’s assumed returns.",'
        '"There are no early retirement stress years in this result.")'
    ), 'report_caption')
    for col, label in zip(('B', 'D', 'F', 'H', 'J'), (
        'Annual stock return', 'Fixed first shortfall age', 'Fixed costs unpaid',
        'Flexible first shortfall age', 'Flexible costs unpaid',
    )):
        cell(f'{col}39', label, 'report_table_head')
    for row, source_row in enumerate(range(28, 41), 40):
        stock, fixed_age, fixed_gap, flexible_age, flexible_gap = (
            result(f'{column}{source_row}') for column in ('AF', 'AG', 'AH', 'AI', 'AJ')
        )
        valid = f'ISNUMBER({stock})'
        values = (
            stock, age_or_none(fixed_age),
            f'IF(ISNUMBER({fixed_gap}),{shortfall(early_dollars(fixed_gap,fixed_age,row-35))},"")',
            age_or_none(flexible_age),
            f'IF(ISNUMBER({flexible_gap}),{shortfall(early_dollars(flexible_gap,flexible_age,row-35,flexible=True))},"")',
        )
        for col, value in zip(('B', 'D', 'F', 'H', 'J'), values):
            style = 'report_percent' if col == 'B' else 'report_row_alt' if row % 2 else 'report_row'
            cell(f'{col}{row}', display(f'IF({valid},{value},"")'), style)

    cell('B54', 'Fixed-spending balances by age', 'report_section')
    cell('B55', 'The low and high values mark the middle 80% of simulated balances at each age.', 'report_caption')
    for col, label in zip(('B', 'D', 'F', 'H', 'J'), ('Age', 'Base plan', 'Low', 'Median', 'High')):
        cell(f'{col}57', label, 'report_table_head')
    for row, source_row in enumerate(range(46, 172), 58):
        age = result(f'AG{source_row}')
        for col, source in zip(('B', 'D', 'F', 'H', 'J'), ('AG', 'AH', 'AI', 'AJ', 'AK')):
            style = 'report_row' if col == 'B' else 'report_money'
            value = result(f'{source}{source_row}')
            if col in ('F','H','J'):
                value=annual_dollars({'F':'D','H':'E','J':'F'}[col],{'F':'BI','H':'BJ','J':'BK'}[col],source_row-41)
            elif col != 'B':
                value = dollars_at(value, age)
            cell(f'{col}{row}', display(f'IF(ISNUMBER({age}),{value},"")'), style)
