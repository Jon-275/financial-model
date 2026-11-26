"""Verify the disposable proof bank's formulas and workbook package boundary."""

from hashlib import sha256
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch
from zipfile import ZIP_STORED, ZipFile, ZipInfo

from openpyxl import Workbook

from native_funding_proof_package import (
    bank_formulas, package_content_sha256, prepare_package, shift_formula, verify_package,
)
from xlsx_values import X, _read_xml_part, sheet_parts


def fixture(path: Path, *, overlap: bool = False) -> str:
    """Create small synthetic worksheet parts with sentinel data outside the bank."""
    workbook = Workbook()
    checks = workbook.active
    checks.title = 'Trial_Checks'
    checks['A1'] = 'Unrelated checks'
    checks['BA5'] = 'Before annual rows'
    checks['CQ6'] = 'Beyond the proof bank'
    for row in range(6, 131):
        checks[f'Y{row}'] = '=1+1'
    if overlap:
        checks.merge_cells('BA126:BB126')
    workbook.create_sheet('Other')['A1'] = 'Existing sheet'
    workbook.save(path)
    with ZipFile(path, 'a') as archive:
        archive.writestr('xl/vbaProject.bin', b'unchanged VBA sentinel')
    return package_content_sha256(path)


class NativeFundingProofPackageTests(unittest.TestCase):
    def test_complete_bank_shape_and_live_claim_mapping(self) -> None:
        """Bind the first, canary and last year to the actual selected solver rows."""
        formulas = bank_formulas()
        self.assertEqual(len(formulas), 125 * 39)
        for row, selected in ((6, 124), (126, 14884), (130, 15376)):
            with self.subTest(row=row):
                self.assertIn(f'Trial_Funding_Calc!$AC${selected}', formulas[f'BA{row}'])
                self.assertIn(f'Trial_Funding_Calc!$D${selected}', formulas[f'BB{row}'])
                self.assertNotIn('Trial_Funding_Calc!', formulas[f'BC{row}'])
                self.assertIn('Trial_Inputs!$B$47=0', formulas[f'BC{row}'])
                self.assertIn(f'Trial_Projection!$BE${row - 1}=0', formulas[f'BC{row}'])
                self.assertIn(',0,""),"")', formulas[f'BC{row}'])
                self.assertIn(f'SUM(BH{row}:BJ{row})', formulas[f'BZ{row}'])
                self.assertIn(f'Trial_Cash_Flow!$AB{row}', formulas[f'BZ{row}'])
                self.assertIn(f'$BZ{row}>0.005', formulas[f'CP{row}'])
                self.assertNotIn(f'BE{row}', formulas)
                self.assertNotIn(f'BF{row}', formulas)
                self.assertNotIn(f'CC{row}', formulas)
        self.assertFalse(any('Trial_Funding_Calc!' in formula for address, formula in formulas.items()
                             if not address.startswith(('BA', 'BB'))))

    def test_shift_preserves_qualified_sources_and_text(self) -> None:
        """Shift local addresses exactly once while retaining absolute references."""
        self.assertEqual(shift_formula('=IF($A6="A6",SUM(B6:$D$6)+Trial_Projection!$BE125,0)'),
                         '=IF($BA6="A6",SUM(BB6:$BD$6)+Trial_Projection!$BE125,0)')
        self.assertEqual(shift_formula("=SUM('Tax_Tables'!$D$6:$D$12)+$AM6"),
                         "=SUM('Tax_Tables'!$D$6:$D$12)+$CM6")
        with self.assertRaisesRegex(ValueError, 'unqualified'):
            shift_formula('=CachedQuote')

    def test_package_changes_only_trial_checks_and_preserves_original_cells(self) -> None:
        """Compare a complete package, including untouched VBA and worksheet bytes."""
        with TemporaryDirectory() as folder:
            source, destination = (Path(folder) / name for name in ('source.xlsm', 'proof.xlsm'))
            pinned = fixture(source)
            source_file = sha256(source.read_bytes()).hexdigest()
            with patch('native_funding_proof_package.PINNED_INPUT_CONTENT_SHA256', pinned):
                self.assertEqual(prepare_package(source, destination), destination)
                self.assertEqual(verify_package(source, destination), destination)
                with self.assertRaises(FileExistsError):
                    prepare_package(source, destination)
                with ZipFile(source) as before, ZipFile(destination) as after:
                    part = sheet_parts(before)['Trial_Checks']
                    self.assertEqual(before.namelist(), after.namelist())
                    self.assertEqual([key for key in before.namelist() if before.read(key) != after.read(key)], [part])
                    sheet = _read_xml_part(after, part, worksheet=True)
                    self.assertEqual(sheet.find(X + 'dimension').get('ref'), 'A1:CQ130')
                    formulas = {cell.get('r'): cell.findtext(X + 'f') for cell in sheet.iter(X + 'c')}
                    self.assertEqual(formulas['Y126'], '1+1')
                    self.assertEqual(formulas['BZ126'], bank_formulas()['BZ126'][1:])
                    self.assertIsNone(sheet.find(f'.//{X}c[@r="BZ126"]/{X}v'))
            self.assertEqual(sha256(source.read_bytes()).hexdigest(), source_file)

    def test_unpinned_source_and_overlaps_leave_no_output(self) -> None:
        """Refuse a different workbook and reserved cells that already overlap."""
        with TemporaryDirectory() as folder:
            source, destination = (Path(folder) / name for name in ('source.xlsm', 'proof.xlsm'))
            pinned = fixture(source, overlap=True)
            with self.assertRaisesRegex(ValueError, 'pinned fictional'):
                prepare_package(source, destination)
            with patch('native_funding_proof_package.PINNED_INPUT_CONTENT_SHA256', pinned):
                with self.assertRaisesRegex(ValueError, 'overlap'):
                    prepare_package(source, destination)
            self.assertFalse(destination.exists())

    def test_verifier_detects_changes_after_bank_packaging(self) -> None:
        """Require independent byte identity for non-worksheet package parts."""
        with TemporaryDirectory() as folder:
            source, destination, modified = (Path(folder) / name for name in ('source.xlsm', 'proof.xlsm', 'edited.xlsm'))
            pinned = fixture(source)
            with patch('native_funding_proof_package.PINNED_INPUT_CONTENT_SHA256', pinned):
                prepare_package(source, destination)
                with ZipFile(destination) as before, ZipFile(modified, 'x') as after:
                    for info in before.infolist():
                        data = b'different VBA' if info.filename == 'xl/vbaProject.bin' else before.read(info)
                        after.writestr(info, data)
                with self.assertRaisesRegex(ValueError, 'unrelated part: xl/vbaProject.bin'):
                    verify_package(source, modified)

    def test_fresh_zip_metadata_cannot_change_the_workbook_identity(self) -> None:
        """Accept the same OOXML and VBA with fresh timestamps, ordering and compression."""
        with TemporaryDirectory() as folder:
            source, fresh, destination = (Path(folder) / name for name in ('old.xlsm', 'fresh.xlsm', 'proof.xlsm'))
            pinned = fixture(source)
            with ZipFile(source) as before, ZipFile(fresh, 'x') as rewritten:
                for member in reversed(before.namelist()):
                    fresh_member = ZipInfo(member, date_time=(2026, 11, 26, 8, 0, 0))
                    rewritten.writestr(fresh_member, before.read(member), compress_type=ZIP_STORED)
            self.assertNotEqual(sha256(fresh.read_bytes()).digest(), sha256(source.read_bytes()).digest())
            self.assertEqual(package_content_sha256(fresh), pinned)
            with patch('native_funding_proof_package.PINNED_INPUT_CONTENT_SHA256', pinned):
                self.assertEqual(prepare_package(fresh, destination), destination)

    def test_pin_rejects_changed_payloads_names_and_duplicate_members(self) -> None:
        """Bind the original VBA and every worksheet to an unambiguous name."""
        with TemporaryDirectory() as folder:
            directory = Path(folder)
            source = directory / 'source.xlsm'
            pinned = fixture(source)
            with ZipFile(source) as original:
                checks = sheet_parts(original)['Trial_Checks']
                vba = 'xl/vbaProject.bin'
                for member, updated_name, payload in (
                    (vba, vba, b'different VBA'), (checks, checks, b'different worksheet'),
                    (vba, 'xl/renamedVba.bin', original.read(vba)),
                ):
                    with self.subTest(member=member, name=updated_name):
                        modified = directory / 'modified.xlsm'
                        with ZipFile(modified, 'w') as changed:
                            for part in original.namelist():
                                if part == member:
                                    changed.writestr(updated_name, payload)
                                else:
                                    changed.writestr(part, original.read(part))
                        with (patch('native_funding_proof_package.PINNED_INPUT_CONTENT_SHA256', pinned),
                              self.assertRaisesRegex(ValueError, 'pinned fictional')):
                            prepare_package(modified, directory / 'bank.xlsm')
            duplicate = directory / 'duplicate.xlsm'
            with ZipFile(duplicate, 'x') as ambiguous:
                ambiguous.writestr(vba, b'original')
                with self.assertWarns(UserWarning):
                    ambiguous.writestr(vba, b'other')
            with self.assertRaisesRegex(ValueError, 'duplicate package members'):
                package_content_sha256(duplicate)
            self.assertFalse((directory / 'bank.xlsm').exists())
