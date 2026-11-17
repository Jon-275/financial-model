"""Check staged Excel workbooks before a repository commit."""

from hashlib import sha256
import json
from pathlib import Path
import subprocess
import sys
from tempfile import TemporaryDirectory
from xml.etree import ElementTree
from zipfile import BadZipFile

from example_inputs import example_cells
from workbook_privacy import audit_fixture, audit_package


ROOT = Path(__file__).resolve().parents[1]
RELEASE = 'workbook/retirement-model.xlsx'
RELEASE_SHA256 = 'fe692dbc84776c373b0513383b4dc4dfc8508130a6a483f80c17940c1fa6a9da'
TEMPLATE = 'workbook/build-template.xlsx'
TEMPLATE_SHA256 = 'fdf966bd8556733b946ce6c1ee36b812c8a1b1abe901c17fc94bb76b90e45221'
EXAMPLE_RISK_SHA256 = '1a2295fa1b277e78f6cf97cb354cd977ae91962b87eff15a886d55c6953712a9'
SPREADSHEET_EXTENSIONS = {'.xlsx', '.xls', '.xlsb', '.xlsm', '.ods'}
sys.path.insert(0, str(ROOT / 'tests'))

from model_acceptance import check_release_package
from xlsx_read import Workbook, address_parts


def git(*args):
    return subprocess.check_output(['git', *args], cwd=ROOT)


def risk_snapshot_digest(workbook):
    """Fingerprint every populated snapshot cell, including saved input values."""
    entries = []
    for cells in workbook['Risk_Data'].iter_rows():
        for cell in cells:
            row, column = address_parts(cell.coordinate)
            if (cell.coordinate == 'Z3' or (row > 1 and column in (29, 54))
                    or (row >= 5 and (1 <= column <= 9 or 11 <= column <= 15 or 32 <= column <= 50))):
                if cell.formula is not None or cell.value not in (None, ''):
                    entries.append((cell.coordinate, cell.formula, cell.value))
    return sha256(json.dumps(sorted(entries), separators=(',', ':')).encode('utf-8')).hexdigest()


def check_sample_workbook(path):
    check_release_package(path)
    audit_package(path)
    with Workbook(path) as workbook:
        audit_fixture(workbook)
        inputs = workbook['Inputs']
        for address, expected in example_cells().items():
            cell = inputs[address]
            actual = '' if cell.value is None else cell.value
            if cell.formula is not None or type(actual) is bool or actual != expected:
                raise ValueError(f'Inputs!{address} differs from the fictional scenario.')
        for address, expected in {'B5': 0, 'B21': 0, 'B36': 250000}.items():
            cell = workbook['Income_Targeting'][address]
            if cell.formula is not None or type(cell.value) is bool or cell.value != expected:
                raise ValueError(f'Income_Targeting!{address} differs from the example settings.')
        if workbook['Risk_Data']['Z1'].value != 'current' or risk_snapshot_digest(workbook) != EXAMPLE_RISK_SHA256:
            raise ValueError('Risk_Data differs from the verified fictional example.')
    if sha256(path.read_bytes()).hexdigest() != RELEASE_SHA256:
        raise ValueError('The release workbook differs from the approved synthetic artifact.')


def main():
    data = git('diff', '--cached', '--name-only', '--diff-filter=AMT', '--no-renames', '-z')
    for raw_path in filter(None, data.split(b'\0')):
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
        except (AssertionError, BadZipFile, ElementTree.ParseError, KeyError, ValueError) as exc:
            print(f'{name}: {exc}', file=sys.stderr)
            return 1
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
