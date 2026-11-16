"""Protect the risk refresh file boundary and workbook payload."""

from datetime import date
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from zipfile import ZipFile

from openpyxl import Workbook, load_workbook

import risk_workbook
from risk_engine.report import run_analysis
from risk_engine.scenario import load_scenario
from run_risk import _check_closed, _publish


class RiskWorkbookTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        workbook = Path(__file__).resolve().parents[1] / 'workbook' / 'retirement-model.xlsx'
        cls.scenario = load_scenario(workbook)
        cls.report = run_analysis(cls.scenario, trials=3, seed=2**64 - 1, volatility=1,
                                  source_sha256=cls.scenario.workbook_sha256,
                                  workbook_calculation='Test')

    def test_snapshot_preserves_other_parts_and_clears_prior_horizon(self):
        with TemporaryDirectory() as root:
            source, calculated, refreshed = (Path(root) / name for name in
                                            ('source.xlsx', 'calculated.xlsx', 'refreshed.xlsx'))
            book = Workbook()
            book.active.title = 'Unchanged'
            book.active['A1'] = '=2+2'
            sheet = book.create_sheet(risk_workbook.SHEET)
            risk_workbook.configure_freshness(
                lambda name, address, value, *_: book[name].__setitem__(address, value))
            sheet['AF171'] = 9999
            book.save(source)
            risk_workbook.validate_source(source)

            for row in range(2, sheet.max_row + 1):
                sheet[f'AB{row}'].value = None
            sheet['AB2'] = 0
            sheet['AB3'] = '=literal, not a formula'
            sheet['AB3'].data_type = 's'
            book.save(calculated)

            risk_workbook.write_snapshot(source, refreshed, calculated, self.scenario, self.report,
                                         run_date=date(2026, 1, 1))
            with ZipFile(source) as before, ZipFile(refreshed) as after:
                self.assertEqual(before.namelist(), after.namelist())
                differences = [part for part in before.namelist() if before.read(part) != after.read(part)]
                self.assertEqual(differences, ['xl/worksheets/sheet2.xml'])

            written = load_workbook(refreshed, read_only=False)
            try:
                sheet = written[risk_workbook.SHEET]
                self.assertEqual(sheet['AC2'].value, 0)
                self.assertEqual(sheet['AC3'].value, '=literal, not a formula')
                self.assertEqual(sheet['AC3'].data_type, 's')
                self.assertIsNone(sheet['AC4'].value)
                self.assertEqual(sheet['AF6'].value, str(2**64 - 1))
                self.assertEqual(sheet['AF2'].value, '2026-01-01')
                self.assertEqual(sheet['AF19'].value, self.report['provenance']['workbook_sha256'])
                self.assertEqual(sheet['AG46'].value, self.report['timeline']['opening_age'])
                self.assertEqual(sheet['AI46'].value, self.report['baseline']['opening_balance_real'])
                self.assertEqual(sheet['AG47'].value, self.report['annual'][0]['age'])
                self.assertIsNone(sheet['AF171'].value)
            finally:
                written.close()

    def test_publish_does_not_replace_output(self):
        with TemporaryDirectory() as root:
            output = Path(root) / 'report.xlsx'
            _publish(b'first', output)
            with self.assertRaises(FileExistsError):
                _publish(b'second', output)
            self.assertEqual(output.read_bytes(), b'first')

    def test_refresh_rejects_an_office_owner_lock(self):
        with TemporaryDirectory() as root:
            source = Path(root) / 'plan.xlsx'
            _check_closed(source)
            for name in ('~$plan.xlsx', '.~lock.plan.xlsx#'):
                with self.subTest(name=name):
                    lock = source.with_name(name)
                    lock.touch()
                    with self.assertRaisesRegex(ValueError, 'Save and close'):
                        _check_closed(source)
                    lock.unlink()
