"""Check that temporary report recalculation preserves exact assumption identity."""

from pathlib import Path
import sys
from tempfile import TemporaryDirectory
import unittest
from zipfile import ZipFile

from lxml import etree

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'tools'))
import native_capture as trial
import native_funding_paths as funding
from native_results import _status
from prepare_native_cache import prepare_cache_copy
from xlsx_values import R, X


def fixture(path, context, *, changed=False):
    bookkeeping = etree.Element(X + 'workbook')
    listing = etree.SubElement(bookkeeping, X + 'sheets')
    rels = etree.Element('Relationships')
    strings = etree.Element(X + 'sst', count='77', uniqueCount='2')
    for value in (context, context.swapcase()):
        etree.SubElement(etree.SubElement(strings, X + 'si'), X + 't').text = value
    with ZipFile(path, 'x') as archive:
        for index, (name, end, suffix) in enumerate(((trial.SHEET, 42, trial.SIGNATURE),
                                                     (funding.DATA_SHEET, 10, funding.SIGNATURE),
                                                     ('Risk_Data', 0, '')), 1):
            sheet = etree.SubElement(listing, X + 'sheet', name=name)
            sheet.set(R + 'id', str(index))
            part = f'worksheets/sheet{index}.xml'
            etree.SubElement(rels, 'Relationship', Id=str(index), Target=part)
            root = etree.Element(X + 'worksheet')
            data = etree.SubElement(root, X + 'sheetData')
            if name == 'Risk_Data':
                def formula(address, value, _style):
                    node = etree.SubElement(data, X + 'c', r=address)
                    etree.SubElement(node, X + 'f').text = value[1:]
                _status(formula)
                cell = etree.SubElement(data, X + 'c', r='Z8', t='s')
                etree.SubElement(cell, X + 'v').text = '0'
            else:
                for row in range(8, end + 1):
                    parent = etree.SubElement(data, X + 'row', r=str(row))
                    for col in ('B', suffix):
                        cell = etree.SubElement(parent, X + 'c', r=f'{col}{row}', t='s')
                        incorrect = changed and name == trial.SHEET and row == end and col == suffix
                        etree.SubElement(cell, X + 'v').text = str(int(incorrect))
            archive.writestr('xl/' + part, etree.tostring(root))
        archive.writestr('xl/workbook.xml', etree.tostring(bookkeeping))
        archive.writestr('xl/_rels/workbook.xml.rels', etree.tostring(rels))
        archive.writestr('xl/sharedStrings.xml', etree.tostring(strings))


class NativeCacheTests(unittest.TestCase):
    def test_only_compared_capture_fields_are_encoded(self):
        with TemporaryDirectory() as directory:
            source, encoded = (Path(directory) / name for name in ('original.xlsx', 'encoded.xlsx'))
            context = 'native-risk-v1|' + '1200|' * 3000
            fixture(source, context)
            self.assertEqual(prepare_cache_copy(source, encoded, 1, context), 76)
            with ZipFile(encoded) as archive:
                strings = etree.fromstring(archive.read('xl/sharedStrings.xml'))
                self.assertEqual(strings[0][0].text, context)
                self.assertEqual(strings[2][0].text, 'verified-completed-assumptions')
                risk = etree.fromstring(archive.read('xl/worksheets/sheet3.xml'))
                self.assertEqual(risk.findtext(f".//{X}c[@r='Z8']/{X}v"), '0')
                self.assertIn('EXACT($Z$8,Market_Paths!$BB$3)', risk.findtext(f".//{X}c[@r='Z5']/{X}f"))
                for part in ('sheet1', 'sheet2'):
                    data = etree.fromstring(archive.read(f'xl/worksheets/{part}.xml'))
                    self.assertEqual({node.text for node in data.iter(X + 'v')}, {'2'})

    def test_case_change_in_last_stress_prevents_publication(self):
        with TemporaryDirectory() as directory:
            source, encoded = (Path(directory) / name for name in ('original.xlsx', 'encoded.xlsx'))
            fixture(source, 'sensitive-to-Case', changed=True)
            with self.assertRaisesRegex(ValueError, 'different captured assumption record'):
                prepare_cache_copy(source, encoded, 1, 'sensitive-to-Case')
            self.assertFalse(encoded.exists())


if __name__ == '__main__':
    unittest.main()
