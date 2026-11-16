"""Reconcile Summary headlines and chart values with the annual schedules."""

from math import isclose


def planning_metrics(wb):
    if 'Planning Metrics' in wb.parts:
        return wb['Planning Metrics'], 'B'
    if 'Annual Detail' in wb.parts:
        return wb['Summary' if 'Conversion Detail' in wb.parts else 'Report_Data'], 'N'
    return wb['Summary'], 'B'


def validate_dashboard(wb, tolerance):
    """Check the plotted time spans, retirement marker and sources of payments."""
    inputs, summary, projection, cash, data = (
        wb[name] for name in ('Inputs', 'Summary', 'Projection', 'Cash_Flow', 'Report_Data')
    )
    report = 'Annual Detail' in wb.parts
    arranged = 'Conversion Detail' in wb.parts
    metrics, metric_col = planning_metrics(wb)
    status = metrics[f'{metric_col}25'].value
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
        if 'Risk Results' in wb.parts:
            risk = wb['Risk Results']
            snapshot = wb['Risk_Data'] if 'Risk_Data' in wb.parts else risk
            _validate_risk_report(risk, snapshot, inputs, projection, data, summary, horizon, tolerance, arranged)
    else:
        expected_headline = 'YES' if status == 'YES' else 'NO' if has_result else 'CHECK PLAN'
        assert summary['B3'].value == expected_headline, 'Summary headline differs from the funding result.'
        if has_result:
            assert summary['B4'].value == summary['B26'].value, 'Summary first unpaid year differs.'
            assert summary['B5'].value == summary['B29'].value, 'Summary ending balance differs.'
        else:
            assert summary['B4'].value in (None, '') and summary['B5'].value in (None, '')
    controls = 'Q' if arranged else 'K'
    equal(f'{controls}2', horizon)
    equal(f'{controls}4', cash_count)
    if report:
        gap = data[f'{controls}6'].value
        expected_gap = metrics[f'{metric_col}26'].value - projection['B5'].value + 1 if status == 'NO - cash shortfall' else 0
        equal(f'{controls}6', expected_gap)
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

    for row in range(2, 128 if arranged else 15):
        if row - 1 > (horizon if arranged else cash_count):
            for column in ('EFGHIJKLMN' if arranged else 'EFGHIJ'):
                equal(f'{column}{row}', None)
            continue
        position = row - 1 if arranged else cash_start + row - 2
        source = position + 4
        cpi = projection[f'G{source}'].value
        amounts = {column: cash[f'{column}{source}'].value for column in ('C', 'D', 'E', 'F', 'M', 'N', 'P', 'V')}
        paid = amounts['N'] + amounts['P']
        equal(f'E{row}', projection[f'{"B" if arranged else "C"}{source}'].value)
        equal(f'F{row}', min(paid, amounts['C']) / cpi)
        equal(f'G{row}', max(0, paid - amounts['C']) / cpi)
        equal(f'H{row}', amounts['E'] / cpi)
        equal(f'I{row}', amounts['V'] / cpi)
        equal(f'J{row}', position)
        if arranged:
            equal(f'K{row}', (max(0, paid - amounts['C']) + amounts['E']) / cpi)
            healthcare = projection[f'BA{source}'].value
            equal(f'L{row}', (amounts['D'] - healthcare) / cpi)
            equal(f'M{row}', healthcare / cpi)
            equal(f'N{row}', (amounts['F'] + amounts['M']) / cpi)
        total = sum(data[f'{column}{row}'].value for column in 'FGHI')
        due = (amounts['D'] + amounts['F'] + amounts['M']) / cpi
        assert isclose(total, due, rel_tol=1e-13, abs_tol=tolerance), f'Chart bills do not reconcile in row {row}.'



def _validate_risk_report(results, risk, inputs, projection, data, summary, horizon, tolerance, arranged):
    status = risk['Z1'].value
    assert status in ('current', 'stale', 'missing'), 'Unknown market risk result status.'
    future_view = summary['K18'].value == 'Future dollars'

    def factor(offset):
        return projection[f'G{offset + 5}'].value if future_view else 1

    for offset in range(horizon):
        row = offset + 2
        assert isclose(data[f'AE{row}'].value, data[f'B{row}'].value * factor(offset),
                       rel_tol=1e-13, abs_tol=tolerance)

    if status != 'current':
        assert summary['F12'].value in (None, '')
        assert summary['B68'].value == 'Key plan decisions'
        assert [summary[f'B{row}'].value for row in range(71, 74)] == [
            'Retirement timing', 'Retirement spending', 'Roth IRA savings goal'
        ]
        result_cells = ('B7', 'F7', 'B58', *(f'{col}{row}' for row in (13, 14, 15) for col in 'FHJ'))
        assert all(results[address].value in (None, '') for address in result_cells)
        assert all(data[f'{col}{row}'].value in (None, '')
                   for row in range(2, 128) for col in ('AF', 'AG', 'AH'))
        if horizon:
            message = 'Run market risk' if status == 'missing' else 'Refresh market risk'
            if horizon == 1:
                message = 'Extend the plan to run market risk'
            assert summary['F13'].value == message
            assert summary['E71'].value == str(int(inputs['B10'].value))
            assert summary['E72'].value == inputs['B37'].value
            assert summary['E73'].value == inputs['B35'].value
        else:
            assert all(summary[f'E{row}'].value in (None, '') for row in range(71, 74))
        return

    fixed, flexible = risk['BB8' if arranged else 'AF8'].value, risk['BB9' if arranged else 'AF9'].value
    trials = risk['BB5' if arranged else 'AF5'].value
    assert horizon >= 2 and 0 <= fixed <= 1 and 0 <= flexible <= 1 and trials >= 1
    assert isclose(summary['F12'].value, fixed, abs_tol=1e-14)
    assert isclose(results['B7'].value, fixed, abs_tol=1e-14)
    assert isclose(results['F7'].value, flexible, abs_tol=1e-14)
    assert summary['F13'].value == f'{round(fixed * trials):,} of {trials:,} paths paid all costs'
    for printed, source in zip(range(71, 74), range(21, 24)):
        assert isclose(summary[f'E{printed}'].value, risk[f'AG{source - 16 if arranged else source}'].value * factor(horizon - 1),
                       rel_tol=1e-13, abs_tol=tolerance)
        first_age = risk[f'AH{source - 16 if arranged else source}'].value
        if first_age is None:
            assert summary[f'G{printed}'].value == 'None'
            assert summary[f'H{printed}'].value in (None, '')
        else:
            assert summary[f'G{printed}'].value == first_age
            amount = risk[f'AI{source - 16 if arranged else source}'].value * factor(first_age - inputs['B8'].value)
            expected = 'Less than $0.01' if amount < 0.01 else f'${amount:,.2f}'
            assert summary[f'H{printed}'].value == expected
    for offset in range(126):
        chart_row, table_row, source_row = offset + 2, offset + 58, offset + 46
        if offset >= horizon:
            assert all(data[f'{col}{chart_row}'].value in (None, '') for col in ('AF', 'AG', 'AH'))
            assert results[f'B{table_row}'].value in (None, '')
            continue
        balance_cols = ('D', 'E', 'F') if arranged else ('AI', 'AJ', 'AK')
        balance_row = source_row - 41 if arranged else source_row
        low, median, high = (risk[f'{column}{balance_row}'].value for column in balance_cols)
        assert 0 <= low <= median <= high, 'Risk interval points are out of order.'
        assert results[f'B{table_row}'].value == risk[f'B{source_row - 41}' if arranged else f'AG{source_row}'].value == data[f'A{chart_row}'].value
        assert isclose(data[f'AF{chart_row}'].value, low * factor(offset), abs_tol=tolerance)
        assert isclose(data[f'AG{chart_row}'].value, (high - low) * factor(offset), abs_tol=tolerance)
        assert isclose(data[f'AH{chart_row}'].value, median * factor(offset), abs_tol=tolerance)
        if results is not risk:
            for col, amount in zip('FHJ', (low, median, high)):
                assert isclose(results[f'{col}{table_row}'].value, amount * factor(offset), abs_tol=tolerance)


def _validate_report(wb, data, summary, status, horizon, retirement, cash_count, tolerance):
    inputs, projection, cash, appendix = (wb[name] for name in ('Inputs', 'Projection', 'Cash_Flow', 'Annual Detail'))
    selected = inputs['B4'].value
    future_view = summary['K18'].value == 'Future dollars'
    if 'Appendix' in wb.parts and selected in (1, 2, 3):
        if 'Risk_Data' in wb.parts:
            record = summary
            assert record['B98'].value == 'Input decisions'
            assert record['B130'].value == 'Spending and account decisions'
            links = (('B10', 'E107'), ('B53', 'E108'), ('B13', 'E111'), ('B27', 'I107'), ('B36', 'E137'))
        else:
            record = inputs
            assert record['L139'].value.startswith(f"Scenario: {inputs['B7'].value}")
            links = (('B10', 'O146'), ('B53', 'O147'), ('B13', 'O150'), ('B27', 'S146'), ('B36', 'O171'))
        for source, reported in links:
            editable = f'{"DEF"[int(selected) - 1]}{source[1:]}'
            expected = inputs[source].value if inputs[editable].value not in (None, '') else None
            value = record[reported].value
            assert value == expected or expected is None and value in (None, ''), (source, reported)
    headline = summary['B8'].value
    if status == 'YES':
        assert headline == f"Projected spending and taxes are covered through age {wb['Inputs']['B53'].value}."
    elif status == 'NO - cash shortfall':
        metrics, metric_col = planning_metrics(wb)
        row = metrics[f'{metric_col}26'].value - projection['B5'].value + 5
        assert headline == f"First shortfall at age {cash[f'B{row}'].value} ({cash[f'A{row}'].value})."
        tax, spending = (cash[f'{col}{row}'].value for col in ('O', 'Q'))
        kind = 'spending and taxes' if tax > 0 and spending > 0 else 'taxes' if tax > 0 else 'spending'
        amount = cash[f'V{row}'].value if future_view else cash[f'V{row}'].value / projection[f'G{row}'].value
        label = 'Less than $0.01' if amount < .01 else f'${amount:,.2f}'
        assert f"{label} in {kind}" in summary['B9'].value
    else:
        assert headline == 'Review the plan inputs or calculation'
        assert summary['B9'].value not in (None, '')
    if horizon:
        balance_col = 'BH' if future_view else 'BN'
        assert isclose(summary['B12'].value, projection[f'{balance_col}{4 + horizon}'].value,
                       rel_tol=1e-13, abs_tol=tolerance)
    else:
        assert summary['B12'].value in (None, '')

    if 'Risk_Data' in wb.parts:
        assert isclose(wb['Risk_Settings']['B10'].value,
                       wb['Risk_Settings']['B7'].value * wb['Risk_Settings']['B8'].value,
                       abs_tol=1e-13)
        for row in range(6, 5 + horizon):
            p = lambda col: projection[f'{col}{row}'].value
            flow = lambda col: cash[f'{col}{row}'].value
            expected = (p('I') - p('J') - p('K'), p('N'), p('O') * inputs['B86'].value,
                        p('AG') * inputs['B89'].value)
            for column, amount in zip(('AN', 'AO', 'AP', 'AQ'), expected):
                assert isclose(flow(column), amount, rel_tol=1e-13, abs_tol=tolerance)
            assert isclose(flow('AB'), sum(expected), rel_tol=1e-13, abs_tol=tolerance)
            assert isclose(flow('C'), flow('AB') + flow('AC'), rel_tol=1e-13, abs_tol=tolerance)

    def annual_row(row, nominal):
        def cf(column):
            return cash[f'{column}{row}'].value
        healthcare = projection[f'BA{row}'].value
        cpi = 1 if nominal else projection[f'G{row}'].value
        paid = cf('N') + cf('P')
        return (cf('A'), cf('B'), min(cf('C'), paid) / cpi,
                (max(0, paid - cf('C')) + cf('E')) / cpi,
                (cf('D') - healthcare) / cpi, healthcare / cpi,
                (cf('F') + cf('M')) / cpi, cf('V') / cpi,
                projection[f'{"BH" if nominal else "BN"}{row}'].value)

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
            nominal = future_view and (sheet is summary or 'Risk_Data' in wb.parts)
            for actual, expected in zip(values, annual_row(source + offset, nominal)):
                assert isinstance(actual, (int, float)) and isclose(actual, expected, rel_tol=1e-13, abs_tol=tolerance), (
                    sheet.title, row, actual, expected
                )
            assert isclose(sum(values[2:4]) + values[7], sum(values[4:7]), abs_tol=tolerance), (
                sheet.title, row, 'cash sources and costs'
            )
