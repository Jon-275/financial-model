"""Check the saved Excel state after restoring a synthetic trial inspection."""
from pathlib import Path
from math import inf, nextafter
from shutil import copyfile
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch
from zipfile import ZipFile

from lxml import etree
from openpyxl import Workbook, load_workbook

from native_excel_acceptance import FUNDING_ROWS_PER_YEAR, compare_restoration, same_saved_value
import native_capture as trials
import native_funding_paths as funding
from xlsx_values import X, sheet_parts, write_values


def make_reference(path):
    book = Workbook()
    table = book.active
    table.title = trials.SHEET
    table['B3'] = 1
    codes = [1, 1001, *range(2001, 2014), *range(2101, 2114), *range(2201, 2208)]
    for row, code in enumerate([1, *codes], trials.FIRST - 1):
        table.cell(row, 1, code)
        table.cell(row, 3, 1)
        table.cell(row, 7, 250)
        table.cell(row, 28, 46000)
    book.create_sheet('Trial_Inputs')['B70'] = 2
    settings = book.create_sheet('Risk_Settings')
    settings['B29'], settings['B30'] = 1, 42
    data = book.create_sheet('Risk_Data')
    for address, val in {'Z1': 'current', 'Z6': 'complete', 'Z7': 46000, 'Z10': 1,
                         'Z11': 42, 'AB3': len(codes)}.items():
        data[address] = val
    book.create_sheet(funding.DATA_SHEET)['A8'] = 1
    solver = book.create_sheet('Trial_Funding_Calc')
    solver['A1'], solver['B1'] = 'Year', 'Best available draw'
    for row in range(2, 2 + 2 * FUNDING_ROWS_PER_YEAR):
        solver.cell(row, 1, '=IF(1=1,"",1)')
    book.save(path)


class SavedInspectionTests(unittest.TestCase):
    def setUp(self):
        folder = TemporaryDirectory()
        self.addCleanup(folder.cleanup)
        self.reference = Path(folder.name) / 'before.xlsx'
        self.after = Path(folder.name) / 'after.xlsx'
        make_reference(self.reference)
        # The small fixture has an opening and two possible withdrawal years.
        maximum = patch.object(trials, 'ANNUAL_CAPACITY', 3)
        maximum.start()
        self.addCleanup(maximum.stop)

    def test_matching_active_inspection_passes(self):
        copyfile(self.reference, self.after)
        self.assertTrue(compare_restoration(self.reference, self.after))

    def test_only_numeric_serialization_differences_can_pass(self):
        self.assertTrue(same_saved_value(1e6, nextafter(1e6, inf)))
        self.assertFalse(same_saved_value(0, 1e-13))
        self.assertFalse(same_saved_value(100, 100.001))
        self.assertFalse(same_saved_value(False, 0))

    def test_interruption_replaces_only_the_run_status_and_message(self):
        for status in ('cancelled', 'failed'):
            with self.subTest(status=status):
                self.after.unlink(missing_ok=True)
                write_values(self.reference, self.after, {'Risk_Data': {'Z1': status, 'Z6': status, 'Z9': 'Run interrupted.'}})
                self.assertTrue(compare_restoration(self.reference, self.after))

    def test_incorrectly_restored_selector_fails_before_the_financial_comparison(self):
        write_values(self.reference, self.after, {trials.SHEET: {'B3': 2}})
        with self.assertRaisesRegex(AssertionError, 'Risk_Trials!B3 was not restored'):
            compare_restoration(self.reference, self.after)

    def test_correct_selector_with_the_wrong_active_financials_fails(self):
        write_values(self.reference, self.after, {trials.SHEET: {'G7': 251}})
        with self.assertRaisesRegex(AssertionError, 'Trial 1'):
            compare_restoration(self.reference, self.after)

    def test_restored_trial_requires_its_own_refresh_date(self):
        write_values(self.reference, self.after, {trials.SHEET: {'AB7': None}})
        with self.assertRaisesRegex(AssertionError, 'Risk_Trials!AB7 needs a valid refresh date'):
            compare_restoration(self.reference, self.after)

    def test_recovery_can_preserve_the_bank_even_if_the_desktop_exercise_failed(self):
        saved = load_workbook(self.reference)
        saved.create_sheet('Exercise')['A2'] = 'FAIL'
        saved.save(self.after)
        saved.close()
        with self.assertRaisesRegex(AssertionError, 'did not pass at Exercise!A2'):
            compare_restoration(self.reference, self.after, 'Exercise!A2')

    def test_saved_ranked_results_must_also_be_restored(self):
        write_values(self.reference, self.after, {funding.DATA_SHEET: {'C10': 0}})
        with self.assertRaisesRegex(AssertionError, 'Funding_Path_Data!C10'):
            compare_restoration(self.reference, self.after)

    def test_erasing_a_formula_cannot_make_an_inactive_year_pass(self):
        row = 2 + FUNDING_ROWS_PER_YEAR
        saved = load_workbook(self.reference)
        saved['Trial_Funding_Calc'][f'A{row}'].value = None
        saved.save(self.after)
        saved.close()
        with self.assertRaisesRegex(AssertionError, f'Trial_Funding_Calc row {row} is missing'):
            compare_restoration(self.reference, self.after)

    def test_inactive_formula_cannot_keep_a_cached_amount_including_zero(self):
        address = 'A' + str(2 + FUNDING_ROWS_PER_YEAR)
        for amount in (None, '0', '7.25'):
            with self.subTest(amount=amount):
                with ZipFile(self.reference) as source, ZipFile(self.after, 'w') as dest:
                    target = sheet_parts(source)['Trial_Funding_Calc']
                    for info in source.infolist():
                        raw = source.read(info)
                        if info.filename == target:
                            xml = etree.fromstring(raw)
                            cell = xml.find(f'.//{X}c[@r="{address}"]')
                            self.assertIsNotNone(cell.find(X + 'f'))
                            for cached in cell.findall(X + 'v'):
                                cell.remove(cached)
                            if amount is not None:
                                etree.SubElement(cell, X + 'v').text = amount
                            raw = etree.tostring(xml)
                        dest.writestr(info, raw)
                if amount is None:
                    self.assertTrue(compare_restoration(self.reference, self.after))
                else:
                    with self.assertRaisesRegex(AssertionError, f'Inactive funding cell Trial_Funding_Calc!{address}'):
                        compare_restoration(self.reference, self.after)


if __name__ == '__main__':
    unittest.main()
