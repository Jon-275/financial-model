"""Audit shared workbook content independently of approved file checksums."""

import re
from xml.etree import ElementTree as E
from zipfile import ZipFile

from start_here import CONTENT as START_HERE_CONTENT
from example_inputs import example_asset_cells, example_cells, example_risk_cells


X = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
DISCARDED_PARTS = frozenset({'docMetadata/LabelInfo.xml', 'docProps/custom.xml', 'xl/calcChain.xml'})
_CURRENT_START_HERE = {address: text for address, (text, _) in START_HERE_CONTENT.items()}
# Keep this historical package pin independent of the current release checksum.
_V1_0_0_WORKBOOK_SHA256 = 'f5faea8318773c911800462d43488c5063dc20b9363f1121151454d5e8007af4'
_V1_0_0_START_HERE = {
    'B3': 'Retirement planner',
    'B5': 'See whether income and savings can cover your spending and taxes through retirement.',
    'B7': 'Save a private copy',
    'B8': ('Save a private copy of this fictional example. To run market tests, open it in desktop Excel '
           'with macros enabled.'),
    'B10': 'Set up your plan',
    'B11': 'Choose a scenario on Inputs. Enter your ages, income, savings and spending in its shaded cells.',
    'B12': 'Open Inputs',
    'B14': 'Run market risk and print your plan',
    'B15': ('In Risk_Settings, use 100 paths while editing and 1,000 for the final report. Select Run market risk '
            'on Summary. Print when it says “Risk results are current.”'),
    'B16': 'Open Summary',
    'B18': 'Inspect the calculations',
    'B19': 'The Appendix lists the schedules behind the projections. Select a formula to see how it works.',
    'B20': 'Open Appendix',
}
_SHARED_REFERENCE = re.compile(
    rb'<(?:[\w]+:)?c\b[^>]*\bt=["\x27]s["\x27][^>]*>'
    rb'\s*<(?:[\w]+:)?v>(\d+)</(?:[\w]+:)?v>'
)
_EMAIL = re.compile(rb'[\w.+-]+@[\w.-]+\.[A-Za-z]{2,}')
_PRIVATE_TEXT = (
    b'your plan:', b'cash buffer instead', b'sv_query_list',
    b'oracle', b'/users/', b'macintosh hd',
)


def audit_package(path, *, template=False):
    """Inspect stored text, package features and unused string storage."""
    with ZipFile(path) as archive:
        names = archive.namelist()
        app = E.fromstring(archive.read('docProps/app.xml'))
        properties = '{http://schemas.openxmlformats.org/officeDocument/2006/extended-properties}'
        allowed = {properties + 'HeadingPairs', properties + 'TitlesOfParts'}
        if any(node.tag not in allowed for node in app):
            raise ValueError('Workbook contains extended document or personal properties.')
        if set(names) & DISCARDED_PARTS:
            raise ValueError('Workbook contains inherited metadata or calculation history.')
        if any(name.startswith(('xl/externalLinks/', 'xl/embeddings/', 'xl/pivotCache/'))
               or name == 'xl/connections.xml' for name in names):
            raise ValueError('Workbook contains an external or embedded data source.')
        used = set()
        for part in names:
            if not part.endswith(('.xml', '.rels')):
                continue
            data = archive.read(part)
            lower = data.lower()
            if any(text in lower for text in _PRIVATE_TEXT):
                raise ValueError(f'Workbook contains private profile or machine text in {part}.')
            if b'@' in data and _EMAIL.search(data):
                raise ValueError(f'Workbook contains an email address in {part}.')
            if part.startswith('xl/worksheets/'):
                used.update(int(match[1]) for match in _SHARED_REFERENCE.finditer(data))
                if template:
                    for cell in E.fromstring(data).iter(X + 'c'):
                        if cell.find(X + 'f') is not None and cell.find(X + 'v') is not None:
                            raise ValueError('The build template contains cached financial results.')
            if template and part.startswith('xl/charts/'):
                if b'numCache' in data or b'strCache' in data:
                    raise ValueError('The build template contains cached chart results.')
        if 'xl/sharedStrings.xml' in names:
            strings = E.fromstring(archive.read('xl/sharedStrings.xml'))
            if used != set(range(len(strings))):
                raise ValueError('Workbook contains unused or invalid shared strings.')


def audit_start_here(sheet, expected=_CURRENT_START_HERE):
    """Only the approved navigation text belongs on the distributed introduction."""
    for row in sheet.rows.values():
        for cell in row.values():
            if cell.formula is not None or (cell.value not in (None, '') and cell.coordinate not in expected):
                raise ValueError(f'Start Here!{cell.coordinate} contains unexpected data.')
    for address, text in expected.items():
        if sheet[address].value != text:
            raise ValueError(f'Start Here!{address} differs from the approved introduction.')


def audit_fixture(workbook):
    """Require synthetic inputs and allocations in every published source."""
    fixtures = [
        ('Inputs', example_cells()), ('Asset_Allocation', example_asset_cells()),
    ]
    if 'Start Here' in workbook.parts:
        expected = _CURRENT_START_HERE
        if workbook.package_sha256() == _V1_0_0_WORKBOOK_SHA256:
            expected = _V1_0_0_START_HERE
        audit_start_here(workbook['Start Here'], expected)
    if 'Risk_Settings' in workbook.parts:
        fixtures.append(('Risk_Settings', example_risk_cells()))
    for sheet_name, expected in fixtures:
        sheet = workbook[sheet_name]
        for address, value in expected.items():
            cell = sheet[address]
            actual = '' if cell.value is None else cell.value
            if cell.formula is not None or isinstance(actual, bool) or actual != value:
                raise ValueError(f'{sheet_name}!{address} differs from the synthetic example.')
