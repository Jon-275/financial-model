"""Presentation changes must leave the financial inputs and formulas intact."""
from collections import defaultdict
import os
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from lxml import etree as E
from openpyxl.utils.cell import range_boundaries, get_column_letter

import decision_layout
from xlsx_read import Workbook


class DecisionLayoutTests(unittest.TestCase):
    def test_saved_example_can_prepare_a_scenario_calculation_copy(self):
        from prepare_native_cache import prepare_cache_copy
        path = Path(os.environ.get('RISK_TEST_WORKBOOK',
                                   Path(__file__).resolve().parents[1] / 'workbook' / 'retirement-model.xlsm'))
        with Workbook(path) as book:
            count = int(book['Risk_Settings']['B29'].value)
            context = book['Risk_Data']['Z8'].value
        with TemporaryDirectory() as directory:
            target = Path(directory) / 'scenario.xlsx'
            checked = prepare_cache_copy(path, target, count, context)
            self.assertGreater(checked, 0)
            self.assertTrue(target.is_file())

    def test_guidance_cannot_replace_calculations_or_numeric_inputs(self):
        changes = defaultdict(dict)
        decision_layout.configure(lambda sheet, address, value, style=None:
                                  changes[sheet].__setitem__(address, value))
        path = Path(__file__).resolve().parents[1] / 'workbook' / 'retirement-model.xlsm'
        with Workbook(path) as book:
            for sheet, cells in changes.items():
                for address, value in cells.items():
                    previous = book[sheet][address]
                    if previous.formula is not None:
                        self.assertEqual(value.lstrip('= '), previous.formula.lstrip('= '), (sheet, address))
                    elif isinstance(previous.value, (int, float)):
                        self.assertEqual(value, previous.value, (sheet, address))
            for address in ('A230', 'A253'):
                self.assertEqual(changes['Tax_Tables'].get(address, book['Tax_Tables'][address].value),
                                 book['Tax_Tables'][address].value)
            # The scenario loader treats tax-sheet formulas as numeric assumptions.
            self.assertFalse(any(isinstance(value, str) and value.startswith('=')
                                 for value in changes['Tax_Tables'].values()))

    def test_working_tab_merges_preserve_every_populated_cell(self):
        from zipfile import ZipFile
        from xlsx_read import sheet_parts
        path = os.environ.get('RISK_TEST_WORKBOOK',
                              Path(__file__).resolve().parents[1] / 'workbook' / 'retirement-model.xlsm')
        with ZipFile(path) as archive:
            parts = sheet_parts(archive)
            for name in (*decision_layout.NAVIGATION, 'Inputs', 'Planning Metrics'):
                root = E.fromstring(archive.read(parts[name]))
                cells = {cell.get('r'): cell for cell in root.iter(decision_layout.X + 'c')}
                seen = set()
                for merge in root.iter(decision_layout.X + 'mergeCell'):
                    area = merge.get('ref')
                    left, top, right, bottom = range_boundaries(area)
                    first = f'{get_column_letter(left)}{top}'
                    for row in range(top, bottom + 1):
                        for column in range(left, right + 1):
                            address = f'{get_column_letter(column)}{row}'
                            self.assertNotIn(address, seen, (name, area))
                            seen.add(address)
                            if address != first and address in cells:
                                self.assertFalse(decision_layout._content(cells[address]), (name, area, address))


if __name__ == '__main__':
    unittest.main()
