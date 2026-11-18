"""Check staged Excel workbooks before a repository commit."""

import argparse
from hashlib import sha256
import json
from pathlib import Path
import subprocess
import sys
from tempfile import TemporaryDirectory
from xml.etree import ElementTree
from zipfile import BadZipFile, ZipFile

from workbook_privacy import audit_fixture, audit_package


ROOT = Path(__file__).resolve().parents[1]
RELEASE = 'workbook/retirement-model.xlsm'
RELEASE_SHA256 = 'de63530a99f2eaa3effe6eb517fe013d808cf31980755ae54a09369a4fc71438'
TEMPLATE = 'workbook/build-template.xlsx'
TEMPLATE_SHA256 = 'fdf966bd8556733b946ce6c1ee36b812c8a1b1abe901c17fc94bb76b90e45221'
EXAMPLE_RISK_SHA256 = '81751c7bc45baf2f8bccfa793789b33cc75a3a192522d349e12c4224b69e1536'
SPREADSHEET_EXTENSIONS = {'.xlsx', '.xls', '.xlsb', '.xlsm', '.ods'}
sys.path.insert(0, str(ROOT / 'tests'))

from model_acceptance import check_release_package
from xlsx_read import Workbook


def git(*args):
    return subprocess.check_output(['git', *args], cwd=ROOT)


def risk_snapshot_digest(workbook):
    """Fingerprint the saved report, every trial and the three full funding paths."""
    digest = sha256()
    for name in ('Risk_Data', 'Risk_Trials', 'Funding_Path_Data'):
        digest.update(name.encode('ascii') + b'\0')
        for number, row in sorted(workbook[name].rows.items()):
            for column, cell in sorted(row.items()):
                if name != 'Risk_Data' and (number < 8 or cell.formula is not None):
                    continue
                record = [number, column, cell.data_type, cell.formula, cell.value]
                digest.update(json.dumps(record, separators=(',', ':'), allow_nan=False).encode('utf-8'))
                digest.update(b'\n')
    return digest.hexdigest()


def check_native_package(path):
    """Verify the file type, VBA relationship and the two on-demand Summary actions."""
    with ZipFile(path) as archive:
        types = ElementTree.fromstring(archive.read('[Content_Types].xml'))
        main = types.find("{*}Override[@PartName='/xl/workbook.xml']")
        if main is None or main.get('ContentType') != 'application/vnd.ms-excel.sheet.macroEnabled.main+xml':
            raise ValueError('The shared workbook must use the macro-enabled Excel file format.')
        rels = ElementTree.fromstring(archive.read('xl/_rels/workbook.xml.rels'))
        projects = [item for item in rels if item.get('Type') == 'http://schemas.microsoft.com/office/2006/relationships/vbaProject']
        if len(projects) != 1 or projects[0].get('Target') != 'vbaProject.bin' or 'xl/vbaProject.bin' not in archive.namelist():
            raise ValueError('The workbook is missing its embedded VBA project.')
        drawing = ElementTree.fromstring(archive.read('xl/drawings/drawing1.xml'))
        actions = sorted(shape.get('macro') for shape in drawing.findall('.//{*}sp') if shape.get('macro'))
        if actions != ['[0]!CancelMarketRisk', '[0]!RunMarketRisk']:
            raise ValueError('Expected one Run market risk button and one Cancel button on Summary.')


def audit_sample_workbook(path):
    """Check package and example content independently of the approved checksums."""
    check_release_package(path)
    check_native_package(path)
    audit_package(path)
    with Workbook(path) as workbook:
        audit_fixture(workbook)
        for address, expected in {'B5': 0, 'B21': 0, 'B36': 250000}.items():
            cell = workbook['Income_Targeting'][address]
            if cell.formula is not None or type(cell.value) is bool or cell.value != expected:
                raise ValueError(f'Income_Targeting!{address} differs from the example settings.')
        settings = workbook['Risk_Settings']
        for row, expected in enumerate((1000, 2026, 1, 0, .015, .5, -.02, .10), 29):
            cell = settings[f'B{row}']
            if cell.formula is not None or type(cell.value) not in (int, float) or cell.value != expected:
                raise ValueError(f'Risk_Settings!B{row} differs from the fictional run.')
        results = workbook['Risk_Data']
        expected_run = {'Z1': 'current', 'Z2': 'retirement-risk-native-v1', 'Z4': 'Excel 1.0', 'Z5': 'current',
                        'Z6': 'complete', 'Z10': 1000, 'Z11': 2026, 'AB3': 2033}
        if any(results[address].value != value for address, value in expected_run.items()):
            raise ValueError('The fictional example requires a current, complete 1,000-pair run.')
        return risk_snapshot_digest(workbook)


def check_sample_workbook(path):
    if audit_sample_workbook(path) != EXAMPLE_RISK_SHA256:
        raise ValueError('The saved trial and funding results differ from the verified fictional example.')
    if sha256(path.read_bytes()).hexdigest() != RELEASE_SHA256:
        raise ValueError('The release workbook differs from the approved synthetic artifact.')


def main(*, tracked=False):
    data = (git('ls-files', '--cached', '-z') if tracked else
            git('diff', '--cached', '--name-only', '--diff-filter=AMT', '--no-renames', '-z'))
    names = set(filter(None, data.split(b'\0')))
    if tracked and not {RELEASE.encode(), TEMPLATE.encode()} <= names:
        print('The index must contain both the shared workbook and its build template.', file=sys.stderr)
        return 1
    checked = 0
    for raw_path in sorted(names):
        name = raw_path.decode('utf-8', errors='surrogateescape')
        if Path(name).suffix.lower() not in SPREADSHEET_EXTENSIONS:
            continue
        try:
            if name not in (RELEASE, TEMPLATE):
                raise ValueError('Keep extra workbook copies under the ignored private/ directory.')
            files = git('ls-files', '--stage', '--', name).splitlines()
            if len(files) != 1 or files[0].split(b' ', 1)[0] != b'100644':
                raise ValueError('Expected a regular workbook file.')
            content = git('show', f':{name}')
            with TemporaryDirectory(prefix='retirement-workbook-check-') as directory:
                staged = Path(directory) / Path(name).name
                staged.write_bytes(content)
                if name == TEMPLATE:
                    check_release_package(staged)
                    audit_package(staged, template=True)
                    with Workbook(staged) as workbook:
                        audit_fixture(workbook)
                    if sha256(content).hexdigest() != TEMPLATE_SHA256:
                        raise ValueError('The audited build template differs from its approved checksum.')
                else:
                    check_sample_workbook(staged)
            checked += 1
        except (AssertionError, BadZipFile, ElementTree.ParseError, KeyError, ValueError) as exc:
            print(f'{name}: {exc}', file=sys.stderr)
            return 1
    if tracked:
        print(f'Verified {checked} tracked workbook files.')
    return 0


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--tracked', action='store_true',
                        help='Audit every workbook in the Git index, including unchanged files.')
    raise SystemExit(main(tracked=parser.parse_args().tracked))
