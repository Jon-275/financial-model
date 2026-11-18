"""Protect report sequencing and prevent content from being merged away."""
from collections import defaultdict
from datetime import datetime
import os
from pathlib import Path
import subprocess
import sys
from tempfile import TemporaryDirectory
from types import SimpleNamespace
import unittest
from zipfile import ZipFile

from lxml import etree
from openpyxl import Workbook, load_workbook
from openpyxl.formula.tokenizer import Tokenizer
from openpyxl.utils.cell import get_column_letter, range_boundaries
from openpyxl.utils.datetime import CALENDAR_MAC_1904, CALENDAR_WINDOWS_1900

import input_record
import native_package
import native_results
import native_funding_paths
import plan_report
import report_layout
import risk_presentation
from xlsx_values import X, sheet_parts


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

    def test_annual_percentiles_follow_the_same_age_in_both_dollar_views(self):
        report = self.cells['Risk Results']
        for data_row, row in enumerate(range(58, 184), 5):
            for column, real, nominal in (('F','D','BI'), ('H','E','BJ'), ('J','F','BK')):
                formula = report[f'{column}{row}']
                for reference in (f'Risk_Data!B{data_row}', f'Risk_Data!{real}{data_row}',
                                  f'Risk_Data!{nominal}{data_row}', f'INDEX(Projection!$G$5:$G$130,{data_row-4})'):
                    self.assertIn(reference, formula)

    def test_risk_source_fields_show_their_units(self):
        styles = defaultdict(dict)
        def put(sheet, address, value, style):
            styles[sheet][address] = style
        native_results.configure(SimpleNamespace(cells=defaultdict(dict), styles=styles, put=put))
        table = styles['Risk_Data']
        groups = {
            'report_percent': ('BB8','BB9','BB12','BB13','BB14','BB21','BB22','BB23','BB24','K5','K17'),
            'report_row': ('BB4','BB5','BB6','BB10','BB15','BB17','L5','N17','AH7','AX5','L23','M29','N29'),
            'report_decimal': ('BB11','BB16','BB18','P5','Q17','AY5','AZ7','P23','R29'),
            'report_money': ('M5','O17','AG7','AW5','O23','Q29'),
        }
        for style, addresses in groups.items():
            for address in addresses:
                self.assertEqual(table[address], style, address)

    def test_live_and_saved_trial_columns_share_display_units(self):
        default = Path(__file__).resolve().parents[1] / 'workbook/retirement-model.xlsm'
        with ZipFile(os.environ.get('RISK_TEST_WORKBOOK', default)) as archive:
            parts = sheet_parts(archive)
            formats = etree.fromstring(archive.read('xl/styles.xml')).find(X+'cellXfs')
            def format_id(style):
                return int(formats[int(style)].get('numFmtId','0'))
            # Only read the column definitions and current row from the large trial ledger.
            with archive.open(parts['Risk_Trials']) as sheet:
                columns = {}
                for _, node in etree.iterparse(sheet, events=('end',), tag=(X+'col', X+'row')):
                    if node.tag == X+'col':
                        for col in range(int(node.get('min')), int(node.get('max'))+1):
                            columns[get_column_letter(col)] = node.get('style')
                    elif node.get('r') == '7':
                        for col in ('B','C','D','E','F','I','K','L','AB','AE','AG'):
                            current = node.find(f"{X}c[@r='{col}7']")
                            self.assertEqual(format_id(current.get('s')), format_id(columns[col]), col)
                        self.assertEqual(format_id(columns['AB']), 14)
                        break
                else:
                    self.fail('Missing current trial row')
            with archive.open(parts[native_funding_paths.DATA_SHEET]) as sheet:
                for _, cols in etree.iterparse(sheet, events=('end',), tag=X+'cols'):
                    source = list(cols)
                    self.assertEqual(source[1].get('style'), source[native_funding_paths.SIGNATURE_COLUMN-1].get('style'))
                    break
                else:
                    self.fail('Missing funding path column formats')

    def test_date_formula_cache_uses_destination_workbook_epoch(self):
        expected = datetime(2026, 11, 17, 9, 30)
        cache = Path(__file__).resolve().parents[1] / 'tools/cache_model.py'
        with TemporaryDirectory() as folder:
            template, checked, output = [Path(folder) / name for name in ('template.xlsx','values.xlsx','output.xlsx')]
            for source_epoch, destination_epoch in ((CALENDAR_MAC_1904, CALENDAR_WINDOWS_1900),
                                                    (CALENDAR_WINDOWS_1900, CALENDAR_MAC_1904)):
                with self.subTest(destination_epoch=destination_epoch):
                    book = Workbook()
                    book.epoch = destination_epoch
                    book.active.title = 'Inputs'
                    book.active['B70'] = 1
                    summary = book.create_sheet('Summary')
                    summary['B12'] = '=NOW()'
                    summary['B12'].number_format = 'yyyy-mm-dd hh:mm'
                    book.save(template)
                    book.epoch = source_epoch
                    summary['B12'] = expected
                    book.save(checked)
                    subprocess.run([sys.executable, str(cache), str(template), str(checked), str(output)],
                                   check=True, capture_output=True)
                    result = load_workbook(output, data_only=True, read_only=True)
                    try:
                        self.assertEqual(result['Summary']['B12'].value, expected)
                    finally:
                        result.close()

    def test_new_named_ranges_precede_later_workbook_elements(self):
        x = native_package.X
        for trailing in ((), ('calcPr',), ('pivotCaches','extLst')):
            book = etree.Element(x+'workbook')
            sheets = etree.SubElement(book, x+'sheets')
            etree.SubElement(sheets, x+'sheet', name='Risk_Trials')
            etree.SubElement(book, x+'externalReferences')
            for tag in trailing:
                etree.SubElement(book, x+tag)
            for _ in range(2):
                native_package.install_names(book, {'Risk_Trials': {'Current':'$B$7'}})
            self.assertEqual([etree.QName(node).localname for node in book],
                             ['sheets','externalReferences','definedNames',*trailing])
            self.assertEqual([el.text for el in book.find(x+'definedNames')], ["'Risk_Trials'!$B$7"])

    def test_run_and_cancel_remain_distinct_accessible_nonprinting_controls(self):
        d = native_package.D
        drawing = etree.Element(d + 'wsDr')
        native_package.install_buttons(drawing)
        native_package.install_buttons(drawing)
        self.assertEqual(len(drawing), 2)
        right_edge = 0
        ids = set()
        for anchor, name, procedure in zip(drawing, ('Run market risk', 'Cancel'),
                                           ('RunMarketRisk', 'CancelMarketRisk'), strict=True):
            shape = anchor.find(d + 'sp')
            body = shape.find(d + 'txBody').find(native_package.A + 'bodyPr')
            self.assertEqual(body.get('wrap'), 'none')
            self.assertEqual((body.get('lIns'), body.get('rIns')), ('12700', '12700'))
            self.assertEqual(shape.find('.//' + native_package.A + 'rPr').get('sz'), '1000')
            self.assertEqual(shape.get('macro'), '[0]!' + procedure)
            properties = shape.find('.//' + d + 'cNvPr')
            self.assertEqual(properties.get('name'), name)
            self.assertTrue(properties.get('descr'))
            ids.add(properties.get('id'))
            self.assertEqual(anchor.findtext(d + 'from/' + d + 'col'), '10')
            self.assertEqual(anchor.findtext(d + 'from/' + d + 'row'), '19')
            left = int(anchor.findtext(d + 'from/' + d + 'colOff'))
            self.assertGreaterEqual(left, right_edge)
            right_edge = left + int(anchor.find(d + 'ext').get('cx'))
            self.assertEqual(anchor.find(d + 'clientData').get('fPrintsWithSheet'), '0')
        self.assertEqual(len(ids), 2)

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
