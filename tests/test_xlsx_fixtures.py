"""Preserve Excel namespace bindings while preparing disposable test inputs."""
from hashlib import sha256
import os
from pathlib import Path
from tempfile import TemporaryDirectory
from types import SimpleNamespace
from unittest.mock import patch
import unittest
from zipfile import ZipFile

from lxml import etree as E

from xlsx_acceptance import patch_xml, prepare_xlsx
from xlsx_read import q, sheet_parts
from xlsx_values import update_sheet, write_values


MC = '{http://schemas.openxmlformats.org/markup-compatibility/2006}'
SHEET = b'''<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
 xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision"
 xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2"
 mc:Ignorable="xr xr2" xr:uid="fixture">
 <sheetData><row r="1"><c r="A1"><v>2</v></c>
 <c r="B1"><f>A1+1</f><v>3</v></c></row></sheetData></worksheet>'''


class FixturePackageTests(unittest.TestCase):
    def test_writing_an_empty_sheet_with_and_without_dimensions(self):
        for has_dimension in (False, True):
            with self.subTest(has_dimension=has_dimension):
                root = E.Element(q('worksheet'))
                if has_dimension:
                    E.SubElement(root, q('dimension'), ref='A1')
                E.SubElement(root, q('sheetData'))
                self.assertEqual(update_sheet(root, {'C7': None}), 0)
                self.assertEqual(list(root.iter(q('row'))), [])
                dim = root.find(q('dimension'))
                self.assertEqual(dim.get('ref') if dim is not None else None, 'A1:A1' if has_dimension else None)
                update_sheet(root, {'C7': 42})
                self.assertEqual(root.find(f".//{q('c')}[@r='C7']").findtext(q('v')), '42')
                dim = root.find(q('dimension'))
                self.assertEqual(dim.get('ref') if dim is not None else None, 'A1:C7' if has_dimension else None)

    def test_shared_string_count_follows_replacements_without_new_shared_text(self):
        with TemporaryDirectory() as directory:
            source, output = (Path(directory) / name for name in ('source.xlsx','output.xlsx'))
            ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
            r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
            with ZipFile(source,'w') as z:
                z.writestr('xl/workbook.xml', f'<workbook xmlns="{ns}" xmlns:r="{r}"><sheets>'
                            '<sheet name="Example" r:id="r1"/></sheets></workbook>')
                z.writestr('xl/_rels/workbook.xml.rels','<Relationships><Relationship Id="r1" '
                            'Target="worksheets/sheet1.xml"/></Relationships>')
                cells = ''.join(f'<c r="{col}1" t="s"><v>0</v></c>' for col in 'ABCD')
                z.writestr('xl/worksheets/sheet1.xml',f'<worksheet xmlns="{ns}"><dimension ref="A1:D1"/>'
                            f'<sheetData><row r="1">{cells}</row></sheetData></worksheet>')
                z.writestr('xl/sharedStrings.xml',f'<sst xmlns="{ns}" count="4" uniqueCount="1">'
                            '<si><t>Example</t></si></sst>')
            write_values(source,output,{'Example':{'A1':None,'B1':42,'C1':'Inline'}})
            with ZipFile(output) as z:
                strings = E.fromstring(z.read('xl/sharedStrings.xml'))
                sheet = E.fromstring(z.read('xl/worksheets/sheet1.xml'))
                shared = [c for c in sheet.iter(q('c')) if c.get('t') == 's']
                self.assertEqual([c.get('r') for c in shared], ['D1'])
                self.assertEqual(strings.get('count'), str(len(shared)))
                self.assertEqual(strings.get('uniqueCount'), '1')

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
        default = Path(__file__).resolve().parents[1] / 'workbook/retirement-model.xlsm'
        source = Path(os.environ.get('RISK_TEST_WORKBOOK', default))
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

    def test_macro_enabled_input_uses_the_xlsx_exported_by_calc(self):
        from workbook_recalc import recalculate
        with TemporaryDirectory() as directory:
            base = Path(directory)
            source = base / 'test.xlsm'
            source.touch()
            output = base / 'result'
            output.mkdir()
            exported = output / 'test.xlsx'
            def run_calc(*args, **kwargs):
                with ZipFile(exported, 'w') as result:
                    result.writestr('[Content_Types].xml', '<Types/>')
                return SimpleNamespace(returncode=0)
            with patch('workbook_recalc.subprocess.run', side_effect=run_calc):
                self.assertEqual(recalculate(source, output, base / 'profile', 'soffice', 30), exported)

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
