"""Reject private content even when a workbook's checksum is approved."""

from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch
from contextlib import redirect_stderr
from io import StringIO
from xml.etree import ElementTree as E
from zipfile import ZipFile

from workbook_privacy import audit_fixture, audit_package, audit_start_here
from check_workbook_commit import check_native_package, check_sample_workbook, main as check_index
from start_here import CONTENT
from xlsx_read import Cell, Sheet, Workbook, address_parts, sheet_parts


ROOT = Path(__file__).resolve().parents[1]
TEMPLATE = ROOT / 'workbook/build-template.xlsx'
X = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'


class WorkbookPrivacyTests(unittest.TestCase):
    def test_tracked_audit_checks_unchanged_files_and_requires_both_artifacts(self):
        required = b'workbook/build-template.xlsx\0workbook/retirement-model.xlsm\0'
        for listing, message in ((b'README.md\0', 'both the shared workbook'),
                                 (b'personal.xlsx\0' + required, 'Keep extra workbook copies')):
            with self.subTest(listing=listing), patch('check_workbook_commit.git', return_value=listing) as git:
                errors = StringIO()
                with redirect_stderr(errors):
                    self.assertEqual(check_index(tracked=True), 1)
                git.assert_called_once_with('ls-files', '--cached', '-z')
                self.assertIn(message, errors.getvalue())

    def copy_with(self, directory, changes, source_path=TEMPLATE):
        destination = Path(directory) / 'changed.xlsx'
        with ZipFile(source_path) as source, ZipFile(destination, 'w') as output:
            for entry in source.infolist():
                output.writestr(entry, changes.pop(entry.filename, source.read(entry)))
            for part, payload in changes.items():
                output.writestr(part, payload)
        return destination

    def test_saved_summary_buttons_call_the_macro_on_their_label(self):
        release = ROOT / 'workbook/retirement-model.xlsm'
        check_native_package(release)
        with ZipFile(release) as archive:
            summary = sheet_parts(archive)['Summary']
            sheet = E.fromstring(archive.read(summary))
            drawing_part, = [part for part in archive.namelist()
                             if part.startswith('xl/drawings/') and part.endswith('.xml')
                             and b'[0]!RunMarketRisk' in archive.read(part)]
            drawing = E.fromstring(archive.read(drawing_part))
        sheet.remove(sheet.find(X + 'drawing'))
        controls = [shape for shape in drawing.findall('.//{*}sp') if shape.get('macro')]
        self.assertEqual(len(controls), 2)
        run, cancel = controls
        run_action, cancel_action = run.get('macro'), cancel.get('macro')
        run.set('macro', cancel_action)
        cancel.set('macro', run_action)
        cases = [('detached', {summary: E.tostring(sheet)}, 'button drawing'),
                 ('swapped', {drawing_part: E.tostring(drawing)}, 'macro named by its label')]
        run.set('macro', run_action)
        cancel.set('macro', cancel_action)
        cancel.find('{*}nvSpPr/{*}cNvPr').set('hidden', '1')
        cases.append(('hidden', {drawing_part: E.tostring(drawing)}, 'show both'))
        for name, changes, diagnostic in cases:
            with self.subTest(name=name), TemporaryDirectory() as directory:
                changed = self.copy_with(directory, changes, release)
                with self.assertRaisesRegex(ValueError, diagnostic):
                    check_native_package(changed)

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

    def test_introduction_rejects_private_data_or_formulas_anywhere(self):
        def put(sheet, address, value, formula=None):
            row, col = address_parts(address)
            sheet.rows.setdefault(row, {})[col] = Cell(address, value, 's', formula)

        for address, value, formula in (('E27', 123456, None), ('G44', 'Household note', None),
                                        ('E30', None, '=Inputs!B12'), ('B3', 'Household name', None)):
            with self.subTest(cell=address):
                sheet = Sheet('Start Here')
                for location, (text, _) in CONTENT.items():
                    put(sheet, location, text)
                audit_start_here(sheet)
                put(sheet, address, value, formula)
                with self.assertRaisesRegex(ValueError, f'Start Here!{address}'):
                    audit_start_here(sheet)

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
