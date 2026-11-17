"""Protect report sequencing and prevent content from being merged away."""
from collections import defaultdict
import unittest

from openpyxl.formula.tokenizer import Tokenizer
from openpyxl.utils.cell import get_column_letter, range_boundaries

import input_record
import plan_report
import report_layout
import risk_presentation


class ReportStructureTests(unittest.TestCase):
    def setUp(self):
        self.cells = defaultdict(dict)

        def put(sheet, address, value, style=None):
            self.cells[sheet][address] = value

        plan_report.configure(put)
        plan_report.configure_index(put)
        input_record.configure(put)
        risk_presentation.configure(
            put, plan_report.display_dollars, plan_report.price_index_at, plan_report.FUTURE_VIEW
        )

    def test_six_pages_keep_all_input_and_spending_records(self):
        summary=self.cells['Summary']
        self.assertEqual(summary['B98'],'Test a more demanding retirement')
        self.assertEqual(summary['B130'],'Input decisions')
        self.assertEqual(summary['B162'],'Spending and account decisions')
        self.assertEqual(plan_report.REPORT_PAGE_BREAKS,(32,64,96,128,160))
        self.assertEqual(input_record.LAST_ROW,189)
        self.assertIn('D10:F10',summary['E139'])
        self.assertIn('D36:F36',summary['E169'])
        self.assertIn('Access unconfirmed',summary['E156'])
        self.assertIn('Access unconfirmed',summary['I156'])

    def test_merges_never_overlap_or_discard_populated_report_cells(self):
        covered={}
        for area in report_layout.summary_merges():
            left,top,right,bottom=range_boundaries(area)
            first=f'{get_column_letter(left)}{top}'
            for row in range(top,bottom+1):
                for column in range(left,right+1):
                    address=f'{get_column_letter(column)}{row}'
                    self.assertNotIn(address,covered,f'{area} overlaps {covered.get(address)}')
                    covered[address]=area
                    if address!=first:
                        self.assertIn(self.cells['Summary'].get(address),(None,''),
                                      f'{area} would discard {address}')

    def test_seven_stress_rows_and_dollar_units_are_complete(self):
        summary=self.cells['Summary']
        for row in (*range(104,109),113,114):
            self.assertTrue(summary[f'B{row}'])
            for column in 'EFGH':self.assertTrue(summary[f'{column}{row}'].startswith('='))
        self.assertIn('Risk_Data!P23',summary['H104'])
        self.assertIn('Risk_Data!P29',summary['H114'])
        self.assertIn('Risk_Data!$AZ$5',summary['E71'])
        self.assertIn('Risk_Data!$AY$5',summary['H71'])
        self.assertEqual(summary['K20'],'')
        self.assertNotIn('K21',summary)

    def test_authored_formulas_fit_excel_limits_and_tokenize(self):
        for sheet,values in self.cells.items():
            for address,value in values.items():
                if isinstance(value,str) and value.startswith('='):
                    self.assertLessEqual(len(value),8192,(sheet,address))
                    Tokenizer(value)

    def test_appendix_links_each_funding_path_view_once(self):
        items = [
            (target, self.cells['Appendix'][f'B{row + offset}'])
            for row, _, schedules in plan_report.index_rows()
            for offset, (target, _, _) in enumerate(schedules, 1)
        ]
        self.assertEqual(len(items), len({target for target, _ in items}))
        self.assertIn(('Funding_Paths', 'Funding paths'), items)
        self.assertIn(('Funding_Path_Data', 'Funding path source data'), items)


if __name__=='__main__':unittest.main()
