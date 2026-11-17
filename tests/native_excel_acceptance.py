"""Reconcile a saved desktop Excel run with an independent example using the same assumptions."""

import argparse
from collections import Counter
from math import isclose
from pathlib import Path
import sys

from native_capture_acceptance import same
from xlsx_read import Workbook, col_letter

import native_capture as trials
import native_funding_paths as funding


def value(sheet, row, column):
    cell = sheet.rows.get(row, {}).get(column)
    return cell.value if cell else None


def assumptions(book):
    market = book['Market_Paths']
    return {value(market, row, 53): value(market, row, 54) for row in market.rows
            if row >= 5 and value(market, row, 53) not in (None, 'Risk input 29')}


def positions(book):
    data = book['Risk_Data']
    n = book['Risk_Settings']['B29'].value
    codes = [*range(1, n + 1), *range(1001, 1001 + n), *range(2001, 2014),
             *range(2101, 2114), *range(2201, 2208)]
    metadata = [data[key].value for key in ('Z1', 'Z6', 'Z10', 'Z11', 'AB3')]
    if metadata != ['current', 'complete', n, book['Risk_Settings']['B30'].value, len(codes)]:
        raise AssertionError('A workbook has an incomplete or stale risk result.')
    if type(data['Z7'].value) not in (int, float) or data['Z7'].value <= 0:
        raise AssertionError('A workbook has no valid run date.')
    table = book[trials.SHEET]
    rows = {code: trials.FIRST + i for i, code in enumerate(codes)}
    for code, row in rows.items():
        if value(table, row, 1) != code or value(table, row, 3) != 1:
            raise AssertionError(f'Missing or invalid capture for trial {code}.')
    return rows


def compare(reference, desktop):
    with Workbook(reference) as expected, Workbook(desktop) as actual:
        ref_rows, test_rows = positions(expected), positions(actual)
        if not test_rows.keys() <= ref_rows.keys():
            raise AssertionError('The reference must contain all the desktop trials.')
        if actual['Risk_Data']['Z7'].value == expected['Risk_Data']['Z7'].value:
            raise AssertionError('The two saved run times are identical. Save the workbook after running Excel.')
        ref_inputs, test_inputs = assumptions(expected), assumptions(actual)
        if not ref_inputs or ref_inputs.keys() != test_inputs.keys():
            raise AssertionError('The workbooks use different sets of assumptions.')
        for name, left in ref_inputs.items():
            right = test_inputs[name]
            if type(left) in (int, float) and type(right) in (int, float):
                matches = isclose(left, right, rel_tol=1e-13, abs_tol=1e-13)
            else:
                matches = left == right
            if not matches:
                raise AssertionError(f'The assumption "{name}" differs.')

        differences, compared = Counter(), Counter()
        samples = []

        def check(label, code, column, observed, target):
            compared[label] += 1
            if not same(observed, target):
                differences[(label, code)] += 1
                if len(samples) < 10:
                    samples.append(f'{label} {code}, column {col_letter(column)}')

        ref_table, test_table = expected[trials.SHEET], actual[trials.SHEET]
        # The two assumption signatures include the run size. AB is the local refresh date.
        for code, row in test_rows.items():
            for column in range(3, trials.SIGNATURE_COLUMN):
                if column != 28:
                    check('trial', code, column, value(test_table, row, column),
                          value(ref_table, ref_rows[code], column))

        ref_paths, test_paths = expected[funding.DATA_SHEET], actual[funding.DATA_SHEET]
        ranked_rows = range(funding.FIRST_ROW, funding.LAST_ROW + 1)
        same_batch = len(ref_rows) == len(test_rows)
        total_offset = next(i for i, field in enumerate(funding.FIELDS) if field[0] == 'Total closing portfolio')
        for row in ranked_rows:
            code = value(test_paths, row, 1)
            if code not in test_rows or code >= 1001 or value(test_paths, row, 3) != 1:
                raise AssertionError('A ranked funding path is missing or invalid.')
            for year in range(trials.ANNUAL_CAPACITY):
                column = 4 + year * funding.FIELD_COUNT + total_offset
                check('funding balance', code, column, value(test_paths, row, column),
                      value(ref_table, ref_rows[code], trials.ANNUAL_FIRST + trials.ANNUAL_CAPACITY + year))
            if same_batch:
                check('ranked path ID', code, 1, code, value(ref_paths, row, 1))
            match = next((r for r in ranked_rows if value(ref_paths, r, 1) == code), None)
            if match is not None:
                for column in range(3, funding.SIGNATURE_COLUMN):
                    check('funding ledger', code, column, value(test_paths, row, column),
                          value(ref_paths, match, column))
        print(f'{len(test_rows):,} outcomes. ' + ', '.join(f'{n:,} {key} fields' for key, n in compared.items()) + '.')
        if differences:
            print(f'{sum(differences.values()):,} differences in {len(differences)} comparisons by trial or path. '
                  + 'Examples: ' + ' / '.join(samples), file=sys.stderr)
            return False
        return True


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('reference', type=Path, help='Independently populated fictional example')
    parser.add_argument('desktop', type=Path, help='Disposable copy saved after running Excel')
    args = parser.parse_args()
    sys.exit(0 if compare(args.reference, args.desktop) else 1)
