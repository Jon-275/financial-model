"""Reconcile Summary headlines and chart values with the annual schedules."""

from math import isclose


def validate_dashboard(wb, tolerance):
    """Check the plotted time spans, retirement marker and sources of payments."""
    inputs, summary, projection, cash, data = (
        wb[name] for name in ('Inputs', 'Summary', 'Projection', 'Cash_Flow', 'Report_Data')
    )
    report = 'Annual Detail' in wb.parts
    status = (data['N25'] if report else summary['B25']).value
    has_result = status in ('YES', 'NO - cash shortfall')
    horizon = inputs['B70'].value if has_result else 0
    checks = wb['Checks']
    assert all(checks[f'J{row}'].value == 'OK' for row in range(5, 5 + int(horizon))), (
        'Summary gives a funding answer before every active row reconciles.'
    )
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

    if report:
        _validate_report(wb, data, summary, status, horizon, retirement, cash_count, tolerance)
    else:
        expected_headline = 'YES' if status == 'YES' else 'NO' if has_result else 'CHECK PLAN'
        assert summary['B3'].value == expected_headline, 'Summary headline differs from the funding result.'
        if has_result:
            assert summary['B4'].value == summary['B26'].value, 'Summary first unpaid year differs.'
            assert summary['B5'].value == summary['B29'].value, 'Summary ending balance differs.'
        else:
            assert summary['B4'].value in (None, '') and summary['B5'].value in (None, '')
    equal('K2', horizon)
    equal('K4', cash_count)
    if report:
        gap = data['K6'].value
        expected_gap = data['N26'].value - projection['B5'].value + 1 if status == 'NO - cash shortfall' else 0
        equal('K6', expected_gap)
        upper = max([1, *(projection[f'BN{5 + row}'].value for row in range(horizon))])
        for age_col, balance_col, position in (('AA', 'AB', retirement), ('AC', 'AD', gap)):
            age = projection[f'C{position + 4}'].value if 0 < position <= horizon else None
            for row, height in ((2, 0), (3, upper)):
                equal(f'{age_col}{row}', age)
                equal(f'{balance_col}{row}', height if age is not None else None)

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
        if report:
            equal(f'D{row}', value if position == gap else None)

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



def _validate_report(wb, data, summary, status, horizon, retirement, cash_count, tolerance):
    inputs, projection, cash, appendix = (wb[name] for name in ('Inputs', 'Projection', 'Cash_Flow', 'Annual Detail'))
    selected = inputs['B4'].value
    if 'Appendix' in wb.parts and selected in (1, 2, 3):
        label = inputs['L139'].value
        assert label.startswith(f"Scenario: {inputs['B7'].value}")
        for source, reported in (('B10', 'O146'), ('B53', 'O147'), ('B13', 'O150'),
                                 ('B27', 'S146'), ('B36', 'O171')):
            editable = f'{"DEF"[int(selected) - 1]}{source[1:]}'
            expected = inputs[source].value if inputs[editable].value not in (None, '') else None
            value = inputs[reported].value
            assert value == expected or expected is None and value in (None, ''), (source, reported)
    headline = summary['B8'].value
    if status == 'YES':
        assert headline == f"Projected spending and taxes are covered through age {wb['Inputs']['B53'].value}."
    elif status == 'NO - cash shortfall':
        row = data['N26'].value - projection['B5'].value + 5
        assert headline == f"First shortfall at age {cash[f'B{row}'].value} ({cash[f'A{row}'].value})."
        tax, spending = (cash[f'{col}{row}'].value for col in ('O', 'Q'))
        kind = 'spending and taxes' if tax > 0 and spending > 0 else 'taxes' if tax > 0 else 'spending'
        amount = cash[f'V{row}'].value / projection[f'G{row}'].value
        label = 'Less than $0.01' if amount < .01 else f'${amount:,.2f}'
        assert f"{label} in {kind}" in summary['B9'].value
    else:
        assert headline == 'Review the plan inputs or calculation'
        assert summary['B9'].value not in (None, '')
    if horizon:
        assert summary['B12'].value == projection[f'BN{4 + horizon}'].value
    else:
        assert summary['B12'].value in (None, '')

    def annual_row(row):
        def cf(column):
            return cash[f'{column}{row}'].value
        healthcare = projection[f'BA{row}'].value
        cpi = projection[f'G{row}'].value
        paid = cf('N') + cf('P')
        return (cf('A'), cf('B'), min(cf('C'), paid) / cpi,
                (max(0, paid - cf('C')) + cf('E')) / cpi,
                (cf('D') - healthcare) / cpi, healthcare / cpi,
                (cf('F') + cf('M')) / cpi, cf('V') / cpi,
                projection[f'BN{row}'].value)

    for sheet, first, count, source, columns in (
        (summary, 41, cash_count, max(2, retirement - 2) + 4, 'BCDEFGHI'),
        (appendix, 6, horizon, 5, 'BCDEFGHIJ'),
    ):
        for offset in range(13 if sheet is summary else 126):
            row = first + offset
            values = [sheet[f'{col}{row}'].value for col in columns]
            if offset >= count:
                assert all(value in (None, '') for value in values), (sheet.title, row)
                continue
            for actual, expected in zip(values, annual_row(source + offset)):
                assert isinstance(actual, (int, float)) and isclose(actual, expected, abs_tol=tolerance), (
                    sheet.title, row, actual, expected
                )
            assert isclose(sum(values[2:4]) + values[7], sum(values[4:7]), abs_tol=tolerance), (
                sheet.title, row, 'cash sources and costs'
            )
