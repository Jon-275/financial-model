"""Compare populated release examples with the visible worksheet trial formulas."""
from __future__ import annotations

import argparse
from concurrent.futures import ThreadPoolExecutor, as_completed
from math import isclose
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


def validate_uncut_pairs(data, positions):
    """Without a spending cut, the flexible policy must reproduce its fixed path."""
    compared = 0
    for code, row in positions.items():
        if 1001 <= code <= 2000:
            fixed = code - 1000
        elif 2101 <= code <= 2113:
            fixed = code - 100
        else:
            continue
        if data[f'I{row}'].value != 0:
            continue
        fixed_row = positions[fixed]
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
