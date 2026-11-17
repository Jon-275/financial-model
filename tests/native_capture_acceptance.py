"""Compare populated release examples with the visible worksheet trial formulas."""
from __future__ import annotations

import argparse
from concurrent.futures import ThreadPoolExecutor, as_completed
from math import isclose, isfinite
from pathlib import Path
import sys
from tempfile import TemporaryDirectory
from zipfile import ZIP_DEFLATED, ZipFile

from lxml import etree

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'tools'))
import native_capture as trial
import native_funding_paths as paths
from workbook_recalc import clear_formula_results, recalculate
from xlsx_read import Workbook, col_letter
from xlsx_values import X, sheet_parts, write_values


def same(actual, expected):
    if expected in (None, ''):
        return actual in (None, '')
    if isinstance(expected, (float, int)) and type(actual) in (float, int):
        return isclose(actual, expected, rel_tol=2e-14, abs_tol=.01)
    return actual == expected


def compare_rows(sheet, saved_sheet, live, saved, last, code, *, first=1, refresh_date=False):
    count = 0
    for column in range(first, last + 1):
        label = col_letter(column)
        calculated = sheet[f'{label}{live}']
        stored = saved_sheet[f'{label}{saved}']
        # A release example can be checked on a different day.
        if refresh_date and label == 'AB':
            if type(stored.value) not in (int, float) or stored.value <= 0:
                raise AssertionError('The captured trial is missing its refresh date.')
        elif not same(stored.value, calculated.value):
            raise AssertionError(f'Trial {code}, {sheet.title}!{label}{saved} differs from its worksheet result.')
        count += 1
    return count


def _number(value, location, maximum=None):
    if type(value) not in (int, float) or not isfinite(value) or value < 0 or (maximum is not None and value > maximum):
        raise AssertionError(f'{location} must contain a finite number in the expected range.')
    return value


def expected_cut(fixed_openings, plan_years, threshold, fraction, share):
    """Before the first cut, a flexible path has the fixed path's opening balances."""
    _number(threshold, 'Risk_Settings!B6', 1)
    _number(fraction, 'Risk_Settings!B7', 1)
    _number(share, 'Risk_Settings!B8', 1)
    cut = False
    for fixed, (base, has_flexible_spending) in zip(fixed_openings, plan_years, strict=True):
        _number(fixed, 'Fixed opening balance')
        _number(base, 'Base opening balance')
        cut |= bool(fraction and share and has_flexible_spending and fixed < threshold * base)
    return cut


def validate_cut_decisions(book, data, positions):
    inputs, projection, settings = (book[name] for name in ('Inputs', 'Projection', 'Risk_Settings'))
    horizon = _number(inputs['B70'].value, 'Inputs!B70', trial.ANNUAL_CAPACITY)
    if horizon != int(horizon) or horizon < 2:
        raise AssertionError('The paired trials require a valid plan horizon.')
    budget = _number(inputs['B37'].value, 'Inputs!B37')
    mortgage = _number(inputs['B38'].value, 'Inputs!B38')
    survivor = _number(inputs['B56'].value, 'Inputs!B56')
    threshold, fraction, share = (settings[f'B{row}'].value for row in (6, 7, 8))
    plan_years = []
    for row in range(6, 5 + int(horizon)):
        stage, filing = projection[f'E{row}'].value, projection[f'F{row}'].value
        if stage not in ('Working', 'Bridge', 'Pre-Medicare', 'Medicare') or filing not in ('MFJ', 'Single'):
            raise AssertionError(f'Projection row {row} has an unknown stage or filing status.')
        # Only the nonmortgage part of retirement living costs can be cut.
        # Its sign is unchanged by the positive simulated price index.
        eligible = stage != 'Working' and budget > mortgage and (filing != 'Single' or survivor > 0)
        plan_years.append((projection[f'BH{row-1}'].value, eligible))
    nominal_start = trial.ANNUAL_FIRST + trial.ANNUAL_CAPACITY
    openings = [col_letter(nominal_start + offset) for offset in range(len(plan_years))]
    pairs = 0
    for code, row in positions.items():
        if 1 <= code <= 1000:
            partner = code + 1000
        elif 2001 <= code <= 2013:
            partner = code + 100
        else:
            continue
        fixed = (data[f'{col}{row}'].value for col in openings)
        actual = data[f'I{positions[partner]}'].value
        if expected_cut(fixed, plan_years, threshold, fraction, share) != (actual > 0):
            raise AssertionError(f'Flexible trial {partner} cut decision disagrees with its fixed path.')
        pairs += 1
    print(f'PASS independent first-cut decisions for {pairs} fixed and flexible pairs.')


def validate_uncut_pairs(data, positions):
    """Compare complete paired financial histories and their capture metadata."""
    compared = 0
    for code, row in positions.items():
        if 1001 <= code <= 2000:
            fixed = code - 1000
        elif 2101 <= code <= 2113:
            fixed = code - 100
        else:
            continue
        fixed_row = positions[fixed]
        signature = col_letter(trial.SIGNATURE_COLUMN)
        context = data[f'B{fixed_row}'].value
        if not isinstance(context, str) or not context:
            raise AssertionError(f'Fixed trial {fixed} has no captured assumption record.')
        for number, saved in ((code, row), (fixed, fixed_row)):
            for col in ('B', signature):
                if data[f'{col}{saved}'].value != context:
                    raise AssertionError(f'Trial {number} has a different captured assumption record in {col}.')
            date = _number(data[f'AB{saved}'].value, f'Trial {number} refresh date')
            if date <= 0 or date != int(date):
                raise AssertionError(f'Trial {number} has an invalid refresh date.')
        years = _number(data[f'I{row}'].value, f'Flexible trial {code} cut years', trial.ANNUAL_CAPACITY - 1)
        amount = _number(data[f'J{row}'].value, f'Flexible trial {code} spending reduction')
        if years != int(years) or bool(years) != bool(amount):
            raise AssertionError(f'Flexible trial {code} has inconsistent spending cuts.')
        if data[f'I{fixed_row}'].value != 0 or data[f'J{fixed_row}'].value != 0:
            raise AssertionError(f'Fixed trial {fixed} recorded a spending cut.')
        if years:
            continue
        for index in range(3, trial.SIGNATURE_COLUMN):
            if index == 28:  # The two trials can run on opposite sides of midnight.
                continue
            column = col_letter(index)
            if data[f'{column}{row}'].value != data[f'{column}{fixed_row}'].value:
                raise AssertionError(f'No-cut trial {code} differs from fixed trial {fixed} in {column}{row}.')
        compared += 1
    print(f'PASS {compared} no-cut flexible paths exactly match their fixed paths.')


def prepare_trial_copy(workbook, destination):
    """Run the live worksheet formulas with empty capture areas, as on first use.

    Compare the results to the published file independently. Loading thousands
    of completed paths into Calc for each trial would calculate the report
    repeatedly and compare its saved data with a resaved version of itself.
    """
    with ZipFile(workbook) as source, ZipFile(destination, 'x', ZIP_DEFLATED, compresslevel=3) as output:
        parts = sheet_parts(source)
        ranges = {parts[trial.SHEET]: (trial.FIRST, trial.LAST),
                  parts[paths.DATA_SHEET]: (paths.FIRST_ROW, paths.LAST_ROW)}
        for part in source.namelist():
            payload = source.read(part)
            if part in ranges:
                root = etree.fromstring(payload)
                first, last = ranges[part]
                for row in root.find(X + 'sheetData'):
                    if first <= int(row.get('r')) <= last:
                        # The primary table also has ranking formulas beyond
                        # the captured columns. Keep every calculation intact.
                        row[:] = [c for c in row if c.find(X + 'f') is not None]
                payload = etree.tostring(root, encoding='UTF-8', xml_declaration=True, standalone=True)
            output.writestr(part, payload)


def validate(workbook, requested, office, jobs=1):
    with Workbook(workbook) as book:
        if book['Risk_Data']['Z1'].value != 'current':
            raise AssertionError('The workbook requires a completed run.')
        n = int(book['Risk_Settings']['B29'].value)
        codes = [*range(1, n + 1), *range(1001, 1001 + n), *range(2001, 2014),
                 *range(2101, 2114), *range(2201, 2208)]
        positions = {code: trial.FIRST + index for index, code in enumerate(codes)}
        data = book[trial.SHEET]
        for code, row in positions.items():
            if data[f'A{row}'].value != code or data[f'C{row}'].value != 1:
                raise AssertionError(f'Trial {code} is missing or invalid.')
        for row in range(trial.FIRST + len(codes), trial.LAST + 1):
            if data[f'A{row}'].value is not None or data[f'C{row}'].value is not None:
                raise AssertionError('Unused trial capacity contains results.')
        validate_uncut_pairs(data, positions)
        validate_cut_decisions(book, data, positions)
        if requested == 'pairs':
            return
        funding = book[paths.DATA_SHEET]
        representatives = {r: int(funding[f'A{r}'].value) for r in range(paths.FIRST_ROW, paths.LAST_ROW + 1)}
    checked = tuple(codes if requested == 'all' else map(int, requested.split(',')))
    if any(code not in positions for code in checked):
        raise ValueError('A requested code does not belong to this completed batch.')
    with TemporaryDirectory() as temp:
        root = Path(temp)
        cleared = root / 'empty-trials.xlsx'
        clean = root / 'clean.xlsx'
        prepare_trial_copy(workbook, cleared)
        clear_formula_results(cleared, clean)
        def verify(code):
            with TemporaryDirectory(dir=root) as run:
                folder = Path(run)
                source = folder / 'trial.xlsx'
                write_values(clean, source, {trial.SHEET: {'B3': code}, 'Risk_Data': {'Z6': 'running'}})
                out = folder / 'out'
                out.mkdir()
                calculated = recalculate(source, out, folder / 'profile', office, 600)
                with Workbook(calculated) as book:
                    if book['Market_Paths']['J3'].value != 'OK' or book[trial.SHEET]['C7'].value != 1:
                        raise AssertionError(f'Trial {code} failed a worksheet check.')
                    total = compare_rows(book[trial.SHEET], data, 7, positions[code], trial.SIGNATURE_COLUMN,
                                         code, refresh_date=True)
                    for row, representative in representatives.items():
                        if code == representative:
                            total += compare_rows(book[paths.DATA_SHEET], funding, paths.CURRENT_ROW, row,
                                                  paths.SIGNATURE_COLUMN, code, first=2)
                return code, total
        comparisons = 0
        pool = ThreadPoolExecutor(max_workers=jobs)
        try:
            for result in as_completed([pool.submit(verify, code) for code in checked]):
                code, count = result.result()
                print(f'PASS worksheet trial {code}: {count:,} captured cells', flush=True)
                comparisons += count
        finally:
            pool.shutdown(wait=True, cancel_futures=True)
        print(f'PASS {len(checked)} direct worksheet trials, {comparisons:,} captured cells.')


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('workbook', type=Path)
    parser.add_argument('--codes', default='all', help='Comma-separated trial codes, pairs for saved no-cut paths, or all.')
    parser.add_argument('--soffice', default='soffice')
    parser.add_argument('--jobs', type=int, default=1)
    options = parser.parse_args()
    validate(options.workbook, options.codes, options.soffice, options.jobs)
