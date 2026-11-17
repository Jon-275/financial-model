"""Preserve Excel namespace bindings while preparing disposable test inputs."""
from hashlib import sha256
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from zipfile import ZipFile

from lxml import etree as E

from xlsx_acceptance import patch_xml, prepare_xlsx
from xlsx_read import q, sheet_parts


MC = '{http://schemas.openxmlformats.org/markup-compatibility/2006}'
SHEET = b'''<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
 xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision"
 xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2"
 mc:Ignorable="xr xr2" xr:uid="fixture">
 <sheetData><row r="1"><c r="A1"><v>2</v></c>
 <c r="B1"><f>A1+1</f><v>3</v></c></row></sheetData></worksheet>'''


class FixturePackageTests(unittest.TestCase):
    def test_ignored_namespace_and_formula_survive_literal_edit(self):
        original = E.fromstring(SHEET)
        changed = E.fromstring(patch_xml(SHEET, {'A1': 7}))
        self.assertEqual(changed.nsmap, original.nsmap)
        self.assertEqual(changed.get(MC + 'Ignorable'), 'xr xr2')
        self.assertEqual(changed.findtext(f'.//{q("c")}[@r="A1"]/{q("v")}'), '7')
        formula = changed.find(f'.//{q("c")}[@r="B1"]')
        self.assertEqual(formula.findtext(q('f')), 'A1+1')
        self.assertIsNone(formula.find(q('v')))

    def test_formula_override_is_rejected(self):
        with self.assertRaisesRegex(ValueError, 'overwrite a formula'):
            patch_xml(SHEET, {'B1': 7})

    def test_release_roundtrip_preserves_namespaces_formulas_and_other_parts(self):
        source = Path(__file__).resolve().parents[1] / 'workbook/retirement-model.xlsx'
        original_hash = sha256(source.read_bytes()).digest()
        with TemporaryDirectory() as directory:
            destination = Path(directory) / 'case.xlsx'
            prepare_xlsx(source, destination, {'Inputs': {'D53': 100}})
            with ZipFile(source) as before, ZipFile(destination) as after:
                self.assertEqual(before.namelist(), after.namelist())
                parts = sheet_parts(before)
                xml_parts = set(parts.values()) | {'xl/workbook.xml'}
                for part in before.namelist():
                    with self.subTest(part=part):
                        if part not in xml_parts:
                            self.assertEqual(before.read(part), after.read(part))
                            continue
                        original = E.fromstring(before.read(part))
                        changed = E.fromstring(after.read(part))
                        self.assertEqual(changed.nsmap, original.nsmap)
                        for node in changed.iter():
                            for prefix in node.get(MC + 'Ignorable', '').split():
                                self.assertIn(prefix, node.nsmap)
                        self.assertEqual(
                            [(f.text, dict(f.attrib)) for f in original.iter(q('f'))],
                            [(f.text, dict(f.attrib)) for f in changed.iter(q('f'))],
                        )
                        if part in parts.values():
                            self.assertFalse(changed.findall(f'.//{q("c")}[{q("f")}]/'+q('v')))
                inputs = E.fromstring(after.read(parts['Inputs']))
                self.assertEqual(inputs.findtext(f'.//{q("c")}[@r="D53"]/{q("v")}'), '100')
            self.assertEqual(sha256(source.read_bytes()).digest(), original_hash)

    def test_fixture_cannot_replace_its_source(self):
        with TemporaryDirectory() as directory:
            source = Path(directory) / 'case.xlsx'
            source.write_bytes(b'original')
            alias = Path(directory) / 'linked-case.xlsx'
            alias.hardlink_to(source)
            for destination in (source, alias):
                with self.subTest(destination=destination.name), self.assertRaisesRegex(ValueError, 'separate destination'):
                    prepare_xlsx(source, destination, {})
            self.assertEqual(source.read_bytes(), b'original')


if __name__ == '__main__':
    unittest.main()
