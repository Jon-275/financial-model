"""Reconcile Summary headlines and chart values with the annual schedules."""

from math import isclose


def validate_dashboard(wb, tolerance):
    """Check the plotted time spans, retirement marker and sources of payments."""
    inputs, summary, projection, cash, data = (
        wb[name] for name in ('Inputs', 'Summary', 'Projection', 'Cash_Flow', 'Report_Data')
    )
    status = summary['B25'].value
    has_result = status in ('YES', 'NO - cash shortfall')
    horizon = inputs['B70'].value if has_result else 0
    retirement = max(2, inputs['B10'].value - inputs['B8'].value + 1) if horizon else 2
    cash_start = max(2, retirement - 2)
    cash_stop = min(horizon, retirement + 10)
    cash_count = cash_stop - cash_start + 1 if retirement <= horizon else 0

    def equal(cell, expected):
        actual = data[cell].value
        if isinstance(expected, (int, float)):
            ok = isinstance(actual, (int, float)) and isclose(actual, expected, rel_tol=1e-13, abs_tol=tolerance)
        elif expected is None:
            ok = actual in (None, '')
        else:
            ok = actual == expected
        if not ok:
            raise AssertionError(f'Report_Data!{cell} differs from the annual schedule.')

    expected_headline = 'YES' if status == 'YES' else 'NO' if has_result else 'CHECK PLAN'
    assert summary['B3'].value == expected_headline, 'Summary headline differs from the funding result.'
    if has_result:
        assert summary['B4'].value == summary['B26'].value, 'Summary first unpaid year differs.'
        assert summary['B5'].value == summary['B29'].value, 'Summary ending balance differs.'
    else:
        assert summary['B4'].value in (None, '') and summary['B5'].value in (None, '')
    equal('K2', horizon)
    equal('K4', cash_count)

    for row in range(2, 128):
        position = row - 1
        if position > horizon:
            for column in 'ABC':
                equal(f'{column}{row}', None)
            continue
        source = position + 4
        value = projection[f'BN{source}'].value
        equal(f'A{row}', projection[f'C{source}'].value)
        equal(f'B{row}', value)
        equal(f'C{row}', value if position == retirement else None)

    for row in range(2, 15):
        if row - 1 > cash_count:
            for column in 'EFGHIJ':
                equal(f'{column}{row}', None)
            continue
        position = cash_start + row - 2
        source = position + 4
        cpi = projection[f'G{source}'].value
        amounts = {column: cash[f'{column}{source}'].value for column in ('C', 'D', 'E', 'F', 'M', 'N', 'P', 'V')}
        paid = amounts['N'] + amounts['P']
        equal(f'E{row}', projection[f'C{source}'].value)
        equal(f'F{row}', min(paid, amounts['C']) / cpi)
        equal(f'G{row}', max(0, paid - amounts['C']) / cpi)
        equal(f'H{row}', amounts['E'] / cpi)
        equal(f'I{row}', amounts['V'] / cpi)
        equal(f'J{row}', position)
        total = sum(data[f'{column}{row}'].value for column in 'FGHI')
        due = (amounts['D'] + amounts['F'] + amounts['M']) / cpi
        assert isclose(total, due, rel_tol=1e-13, abs_tol=tolerance), f'Chart bills do not reconcile in row {row}.'
