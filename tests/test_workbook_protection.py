"""Check formula visibility, editable ranges and validation addresses."""
from collections import Counter
import unittest

from lxml import etree as ET
from openpyxl.utils.cell import get_column_letter, range_boundaries

import input_validation
from example_inputs import example_cells
from workbook_protection import WorksheetProtection, X, _entry_cells


def addresses(ref):
    left, top, right, bottom = range_boundaries(ref)
    return {f'{get_column_letter(col)}{row}' for row in range(top, bottom + 1)
            for col in range(left, right + 1)}


class ProtectionTests(unittest.TestCase):
    def test_only_selected_literals_are_unlocked_and_formulas_remain_visible(self):
        styles = ET.fromstring(f'<styleSheet xmlns="{X[1:-1]}"><cellStyleXfs><xf/></cellStyleXfs>'
                               '<cellXfs count="1"><xf><protection hidden="1" locked="0"/>'
                               '<extLst/></xf></cellXfs></styleSheet>')
        sheet = ET.fromstring(f'<worksheet xmlns="{X[1:-1]}"><sheetData><row r="18"><c r="K18"/></row>'
                              '<row r="90"><c r="B90"><f>1+1</f></c><c r="D90"/></row></sheetData>'
                              '<sheetCalcPr/><mergeCells><mergeCell ref="D90:I90"/></mergeCells></worksheet>')
        guard = WorksheetProtection(styles)
        guard.apply('Summary', sheet)
        flags = sheet.find(X + 'sheetProtection')
        self.assertEqual(flags.get('sheet'), '1')
        self.assertIsNone(flags.get('password'))
        self.assertEqual([flags.get(name) for name in ('selectLockedCells','selectUnlockedCells','autoFilter')],
                         ['0','0','0'])
        self.assertEqual(flags.getprevious().tag, X + 'sheetCalcPr')
        self.assertEqual(flags.getnext().tag, X + 'mergeCells')
        self.assertEqual([(xf.find(X + 'protection').get('locked'), xf.find(X + 'protection').get('hidden'))
                          for xf in guard.xfs], [('1','0'), ('0','0')])
        self.assertEqual({cell.get('r') for cell in sheet.iter(X + 'c') if cell.get('s') == '1'},
                         {'K18'} | addresses('D90:I90'))
        self.assertIsNone(sheet.find(f'.//{X}c[@r="B90"]').get('s'))
        self.assertEqual(guard.xfs.get('count'), '2')

    def test_formula_cannot_be_unlocked_by_a_stale_address(self):
        styles = ET.fromstring(f'<styleSheet xmlns="{X[1:-1]}"><cellStyleXfs/><cellXfs><xf/></cellXfs></styleSheet>')
        sheet = ET.fromstring(f'<worksheet xmlns="{X[1:-1]}"><sheetData><row r="3"><c r="B3">'
                              '<f>1+1</f></c></row></sheetData></worksheet>')
        with self.assertRaisesRegex(ValueError, 'Risk_Trials!B3'):
            WorksheetProtection(styles).apply('Risk_Trials', sheet)

    def test_entry_contract_matches_the_scenarios_and_each_numeric_field_has_one_rule(self):
        root = ET.Element(X + 'worksheet')
        self.assertEqual(_entry_cells('Inputs', root), set(example_cells()) - {'B83','C83'})
        for name, rules in input_validation.RULES.items():
            with self.subTest(sheet=name):
                checked = Counter(cell for refs, *_ in rules for ref in refs.split() for cell in addresses(ref))
                self.assertEqual({r for r,v in checked.items() if v != 1}, set())
                text_fields = {'Inputs': addresses('D7:F7') | addresses('A75:A82'),
                               'Asset_Allocation': {'A24','A33','A42'}}.get(name, set())
                self.assertEqual(set(checked), _entry_cells(name, root) - text_fields)

    def test_trial_sheets_and_output_cells_cannot_be_edited(self):
        root = ET.Element(X + 'worksheet')
        for name in ('Start Here','README','Risk_Data','Funding_Path_Data','Trial_Inputs','Projection','Checks','Annual Detail'):
            self.assertEqual(_entry_cells(name, root), set(), name)
        self.assertEqual(_entry_cells('Risk_Trials', root), {'B3'})

    def test_tax_literals_are_editable_and_derived_tax_rates_are_protected(self):
        root = ET.fromstring(f'<worksheet xmlns="{X[1:-1]}"><sheetData><row r="6"><c r="B6"><v>0.1</v></c>'
                             '<c r="D6"><f>B6</f><v>0.1</v></c><c r="G6" t="s"><v>1</v></c>'
                             '</row></sheetData></worksheet>')
        self.assertEqual(_entry_cells('Tax_Tables', root), {'B6'})
