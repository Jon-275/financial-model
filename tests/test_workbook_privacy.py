"""Reject private content even when a workbook's checksum is approved."""

from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from xml.etree import ElementTree as E
from zipfile import ZipFile

from workbook_privacy import audit_fixture, audit_package
from check_workbook_commit import check_sample_workbook
from xlsx_read import Workbook, sheet_parts


ROOT = Path(__file__).resolve().parents[1]
TEMPLATE = ROOT / 'workbook/build-template.xlsx'
X = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'


class WorkbookPrivacyTests(unittest.TestCase):
    def copy_with(self, directory, changes, source_path=TEMPLATE):
        destination = Path(directory) / 'changed.xlsx'
        with ZipFile(source_path) as source, ZipFile(destination, 'w') as output:
            for entry in source.infolist():
                output.writestr(entry, changes.pop(entry.filename, source.read(entry)))
            for part, payload in changes.items():
                output.writestr(part, payload)
        return destination

    def change_cell(self, sheet, address, value, source_path=TEMPLATE):
        with ZipFile(source_path) as source:
            part = sheet_parts(source)[sheet]
            root = E.fromstring(source.read(part))
        cell = root.find(f".//{X}c[@r='{address}']")
        self.assertIsNotNone(cell)
        for node in list(cell):
            cell.remove(node)
        if isinstance(value, str):
            cell.set('t', 'inlineStr')
            E.SubElement(E.SubElement(cell, X + 'is'), X + 't').text = value
        else:
            cell.attrib.pop('t', None)
            E.SubElement(cell, X + 'v').text = str(value)
        return {part: E.tostring(root)}

    def test_clean_template_is_synthetic(self):
        audit_package(TEMPLATE, template=True)
        with Workbook(TEMPLATE) as workbook:
            audit_fixture(workbook)

    def test_personal_profile_text_is_rejected(self):
        with TemporaryDirectory() as directory:
            path = self.copy_with(directory, self.change_cell('Asset_Allocation', 'A33', 'Your plan: private example'))
            with self.assertRaisesRegex(ValueError, 'private profile'):
                audit_package(path, template=True)

    def test_changed_allocation_is_rejected(self):
        with TemporaryDirectory() as directory:
            path = self.copy_with(directory, self.change_cell('Asset_Allocation', 'B38', 0.123))
            with Workbook(path) as workbook, self.assertRaisesRegex(ValueError, 'synthetic'):
                audit_fixture(workbook)

    def test_private_scenario_value_is_rejected(self):
        with TemporaryDirectory() as directory:
            path = self.copy_with(directory, self.change_cell('Inputs', 'D25', 123456))
            with Workbook(path) as workbook, self.assertRaisesRegex(ValueError, 'synthetic'):
                audit_fixture(workbook)

    def test_changed_starter_salary_is_rejected(self):
        release = ROOT / 'workbook/retirement-model.xlsm'
        with TemporaryDirectory() as directory:
            changes = self.change_cell('Start Here', 'E27', 123456, release)
            path = self.copy_with(directory, changes, release)
            with Workbook(path) as workbook, self.assertRaisesRegex(ValueError, 'Start Here!E27.*synthetic'):
                audit_fixture(workbook)

    def test_unused_shared_string_is_rejected(self):
        with ZipFile(TEMPLATE) as source:
            root = E.fromstring(source.read('xl/sharedStrings.xml'))
        E.SubElement(E.SubElement(root, X + 'si'), X + 't').text = 'Stored only in the package'
        with TemporaryDirectory() as directory:
            path = self.copy_with(directory, {'xl/sharedStrings.xml': E.tostring(root)})
            with self.assertRaisesRegex(ValueError, 'unused or invalid'):
                audit_package(path, template=True)

    def test_inherited_metadata_is_rejected(self):
        with TemporaryDirectory() as directory:
            path = self.copy_with(directory, {'docProps/custom.xml': b'<Properties/>'})
            with self.assertRaisesRegex(ValueError, 'metadata'):
                audit_package(path, template=True)

    def test_template_formula_cache_is_rejected(self):
        with ZipFile(TEMPLATE) as source:
            part = sheet_parts(source)['Summary']
            root = E.fromstring(source.read(part))
        cell = next(cell for cell in root.iter(X + 'c') if cell.find(X + 'f') is not None)
        E.SubElement(cell, X + 'v').text = '123456'
        with TemporaryDirectory() as directory:
            path = self.copy_with(directory, {part: E.tostring(root)})
            with self.assertRaisesRegex(ValueError, 'cached financial'):
                audit_package(path, template=True)

    def test_extended_personal_properties_are_rejected(self):
        for name in ('Company', 'Manager'):
            with self.subTest(property=name), TemporaryDirectory() as directory:
                with ZipFile(TEMPLATE) as source:
                    root = E.fromstring(source.read('docProps/app.xml'))
                namespace = '{http://schemas.openxmlformats.org/officeDocument/2006/extended-properties}'
                E.SubElement(root, namespace + name).text = 'Example personal property'
                path = self.copy_with(directory, {'docProps/app.xml': E.tostring(root)})
                with self.assertRaisesRegex(ValueError, 'extended document'):
                    audit_package(path, template=True)

    def test_added_notes_outside_known_ranges_are_rejected(self):
        release = ROOT / 'workbook/retirement-model.xlsm'
        with ZipFile(release) as source:
            part = sheet_parts(source)['Summary']
            root = E.fromstring(source.read(part))
        row = root.find(X + 'sheetData').find(X + 'row')
        cell = E.SubElement(row, X + 'c', r='M1', t='inlineStr')
        E.SubElement(E.SubElement(cell, X + 'is'), X + 't').text = 'Additional private note'
        with TemporaryDirectory() as directory:
            path = self.copy_with(directory, {part: E.tostring(root)}, source_path=release)
            with self.assertRaisesRegex(ValueError, 'approved synthetic artifact'):
                check_sample_workbook(path)


if __name__ == '__main__':
    unittest.main()
