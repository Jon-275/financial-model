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


def layout_sheet():
    root = E.Element(decision_layout.X + 'worksheet')
    E.SubElement(root, decision_layout.X + 'dimension', ref='A1:H265')
    views = E.SubElement(root, decision_layout.X + 'sheetViews')
    E.SubElement(views, decision_layout.X + 'sheetView', workbookViewId='0')
    E.SubElement(root, decision_layout.X + 'sheetData')
    return root


class LayoutStyles(dict):
    def __init__(self, numbers):
        super().__init__()
        self.numbers = numbers

    def __missing__(self, key):
        value = str(len(self.numbers))
        self.numbers.append(key[1] if isinstance(key, tuple) else '0')
        self[key] = value
        return value


class DecisionLayoutTests(unittest.TestCase):
    def test_tax_navigation_caption_has_room_in_the_builder_and_saved_workbook(self):
        from zipfile import ZipFile
        from xlsx_read import sheet_parts
        root = layout_sheet()
        cell = decision_layout._cell(root, 'G3')
        cell.set('t', 'inlineStr')
        E.SubElement(E.SubElement(cell, decision_layout.X + 'is'), decision_layout.X + 't').text = 'Shared by all scenarios.'
        numbers = ['0']
        decision_layout.apply('Tax_Tables', root, LayoutStyles(numbers), numbers)
        row_path = f"{decision_layout.X}sheetData/{decision_layout.X}row[@r='3']"
        self.assertGreaterEqual(float(root.find(row_path).get('ht')), 20)
        path = os.environ.get('RISK_TEST_WORKBOOK',
                              Path(__file__).resolve().parents[1] / 'workbook' / 'retirement-model.xlsm')
        with ZipFile(path) as archive:
            root = E.fromstring(archive.read(sheet_parts(archive)['Tax_Tables']))
            self.assertTrue(decision_layout._content(decision_layout._cell(root, 'G3')))
            self.assertGreaterEqual(float(root.find(row_path).get('ht')), 20)

    def test_empty_inputs_keep_percent_and_currency_formats(self):
        for name, address, code in (('Asset_Allocation', 'B6', '10'),
                                    ('Income_Targeting', 'B36', '165')):
            with self.subTest(sheet=name, cell=address):
                root = layout_sheet()
                cell = decision_layout._cell(root, address)
                cell.set('s', '1')
                numbers = ['0', code]
                styles = LayoutStyles(numbers)
                decision_layout.apply(name, root, styles, numbers)
                self.assertEqual(numbers[int(cell.get('s'))], code)
                self.assertEqual(cell.get('s'), styles['input_value', code])

    def test_navigation_keeps_the_conversion_detail_and_external_links(self):
        from zipfile import ZipFile
        from xlsx_read import sheet_parts
        root = layout_sheet()
        links = E.SubElement(root, decision_layout.X + 'hyperlinks')
        E.SubElement(links, decision_layout.X + 'hyperlink', ref='A41', location="'Conversion Detail'!A1")
        relationship = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'
        E.SubElement(links, decision_layout.X + 'hyperlink', ref='C40', attrib={relationship: 'rId7'})
        numbers = ['0']
        decision_layout.apply('Income_Targeting', root, LayoutStyles(numbers), numbers)
        actual = {node.get('ref'): node for node in root.iter(decision_layout.X + 'hyperlink')}
        self.assertEqual(actual['A41'].get('location'), "'Conversion Detail'!A1")
        self.assertEqual(actual['C40'].get(relationship), 'rId7')
        self.assertEqual(actual['H2'].get('location'), 'Summary!B3')
        path = os.environ.get('RISK_TEST_WORKBOOK',
                              Path(__file__).resolve().parents[1] / 'workbook' / 'retirement-model.xlsm')
        with ZipFile(path) as archive:
            root = E.fromstring(archive.read(sheet_parts(archive)['Income_Targeting']))
            link = root.find(f"{decision_layout.X}hyperlinks/{decision_layout.X}hyperlink[@ref='A41']")
            self.assertIsNotNone(link, 'The distributed workbook must retain its original detail link.')
            self.assertEqual(link.get('location'), "'Conversion Detail'!A1")

    def test_tax_row_with_an_empty_label_keeps_its_value_column(self):
        root = layout_sheet()
        decision_layout._cell(root, 'A10').set('s', '0')
        E.SubElement(decision_layout._cell(root, 'B10'), decision_layout.X + 'v').text = '123'
        numbers = ['0']
        decision_layout.apply('Tax_Tables', root, LayoutStyles(numbers), numbers)
        self.assertEqual(root.find(f".//{decision_layout.X}c[@r='B10']/{decision_layout.X}v").text, '123')
        self.assertNotIn('A10:E10', [node.get('ref') for node in root.iter(decision_layout.X + 'mergeCell')])

    def test_tax_shared_text_gets_the_same_heading_and_note_layout_as_inline_text(self):
        from copy import deepcopy
        strings = ['ADDITIONAL LIMITS', 'Published source details. ' * 8,
                   'This explanatory note describes the tax assumption. ' * 5]
        shared = layout_sheet()
        for address, index in (('A20', 0), ('D21', 1), ('A22', 2)):
            cell = decision_layout._cell(shared, address)
            cell.set('t', 's')
            E.SubElement(cell, decision_layout.X + 'v').text = str(index)
        cell = decision_layout._cell(shared, 'A21')
        cell.set('t', 'inlineStr')
        E.SubElement(E.SubElement(cell, decision_layout.X + 'is'), decision_layout.X + 't').text = 'Deduction'
        E.SubElement(decision_layout._cell(shared, 'B21'), decision_layout.X + 'v').text = '100'
        inline = deepcopy(shared)
        for cell in inline.iter(decision_layout.X + 'c'):
            if cell.get('t') == 's':
                value = cell.find(decision_layout.X + 'v')
                text = strings[int(value.text)]
                cell.remove(value)
                cell.set('t', 'inlineStr')
                E.SubElement(E.SubElement(cell, decision_layout.X + 'is'), decision_layout.X + 't').text = text
        numbers = ['0']
        styles = LayoutStyles(numbers)
        decision_layout.apply('Tax_Tables', shared, styles, numbers, strings)
        decision_layout.apply('Tax_Tables', inline, styles, numbers)
        for row in (20, 21, 22):
            left = shared.find(f"{decision_layout.X}sheetData/{decision_layout.X}row[@r='{row}']")
            right = inline.find(f"{decision_layout.X}sheetData/{decision_layout.X}row[@r='{row}']")
            self.assertEqual(left.get('ht'), right.get('ht'))
        self.assertEqual(decision_layout._cell(shared, 'A20').get('s'), styles['detail_band'])
        self.assertGreaterEqual(float(shared.find(f"{decision_layout.X}sheetData/{decision_layout.X}row[@r='21']").get('ht')), 42)
        self.assertGreaterEqual(float(shared.find(f"{decision_layout.X}sheetData/{decision_layout.X}row[@r='22']").get('ht')), 38)

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
        path = Path(os.environ.get('RISK_TEST_WORKBOOK',
                                   Path(__file__).resolve().parents[1] / 'workbook' / 'retirement-model.xlsm'))
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
