"""Audit shared workbook content independently of approved file checksums."""

import re
from xml.etree import ElementTree as E
from zipfile import ZipFile

from example_inputs import STARTER_EXAMPLE, example_asset_cells, example_cells, example_risk_cells


X = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
DISCARDED_PARTS = frozenset({'docMetadata/LabelInfo.xml', 'docProps/custom.xml', 'xl/calcChain.xml'})
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


def audit_fixture(workbook):
    """Require synthetic inputs and allocations in every published source."""
    fixtures = [
        ('Inputs', example_cells()), ('Asset_Allocation', example_asset_cells()),
    ]
    if 'Start Here' in workbook.parts:
        fixtures.append(('Start Here', STARTER_EXAMPLE))
    if 'Risk_Settings' in workbook.parts:
        fixtures.append(('Risk_Settings', example_risk_cells()))
    for sheet_name, expected in fixtures:
        sheet = workbook[sheet_name]
        for address, value in expected.items():
            cell = sheet[address]
            actual = '' if cell.value is None else cell.value
            if cell.formula is not None or isinstance(actual, bool) or actual != value:
                raise ValueError(f'{sheet_name}!{address} differs from the synthetic example.')
