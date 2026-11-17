"""Build the desktop workbook's VBA project from its reviewed text source."""

from argparse import ArgumentParser
from os import link
from pathlib import Path
from tempfile import TemporaryDirectory
from zipfile import ZIP_DEFLATED, ZipFile

from lxml import etree
from pyopenvba import ExcelFile
from pyopenvba.cfb import CFB
from pyopenvba.vba import decompress


SOURCE = Path(__file__).resolve().parent / 'vba' / 'RunMarketRisk.bas'
MODULE = 'MarketRiskRunner'
WORKBOOK_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'
MACRO_TYPE = 'application/vnd.ms-excel.sheet.macroEnabled.main+xml'


def audit_project_metadata(payload, module_names):
    """Check the build's project layout and local-profile text in its binary records."""
    project = CFB.from_bytes(payload)
    if (set(project.list_storages_at()) != {'VBA'}
            or set(project.list_streams_at()) != {'PROJECT', 'PROJECTwm'}
            or project.list_storages_at(('VBA',))
            or set(project.list_streams_at(('VBA',))) != module_names | {'dir', '_VBA_PROJECT'}):
        raise ValueError('The source-built VBA project contains unexpected records.')
    # The directory is compressed. Excel also stores library paths as UTF-16.
    directory = decompress(project.get_stream_at(('VBA',), 'dir'))
    for record in (payload, directory):
        for location in ('/users/', '/home/', '\\users\\', 'macintosh hd:users:'):
            for encoding in ('ascii', 'utf-16le', 'utf-16be'):
                if location.encode(encoding) in record.lower():
                    raise ValueError('The VBA project contains a local profile path.')


def verify_project(path):
    """Require the exact published runner and empty workbook and sheet modules."""
    with ZipFile(path) as archive:
        book = etree.fromstring(archive.read('xl/workbook.xml'))
        sheets = book.find('{*}sheets')
        expected = {MODULE, 'ThisWorkbook', *(f'Sheet{i}' for i in range(1, len(sheets) + 1))}
        audit_project_metadata(archive.read('xl/vbaProject.bin'), expected)
    with ExcelFile(path) as workbook:
        modules = workbook.vba_modules()
        if set(modules) != expected or workbook.forms():
            raise ValueError('The VBA project contains unexpected modules or forms.')
        if modules.pop(MODULE).splitlines() != SOURCE.read_text(encoding='ascii').splitlines():
            raise ValueError('The embedded runner differs from its published source.')
        for source in modules.values():
            if any(line.strip() and not line.startswith('Attribute VB_') for line in source.splitlines()):
                raise ValueError('Unexpected code in a workbook or worksheet module.')
        refs = workbook.references()
        if {ref.name for ref in refs} != {'Office', 'stdole'} or len(refs) != 2:
            raise ValueError('Unexpected library reference in the VBA project.')
        for ref in refs:
            location = (ref.libid + ref.libid_secondary).replace('\\', '/').lower()
            if '/users/' in location or '/home/' in location or '@' in location:
                raise ValueError('A VBA reference contains a local profile path.')


def package_macro(source, destination):
    source, destination = Path(source), Path(destination)
    if source.suffix.lower() != '.xlsx' or destination.suffix.lower() != '.xlsm':
        raise ValueError('Use an .xlsx calculation workbook and a new .xlsm destination.')
    if destination.exists():
        raise FileExistsError(destination)
    destination.parent.mkdir(parents=True, exist_ok=True)
    with TemporaryDirectory(dir=destination.parent) as directory:
        build = Path(directory) / 'retirement-model.xlsm'
        with ZipFile(source) as original, ZipFile(build, 'x', ZIP_DEFLATED, compresslevel=5) as output:
            if 'xl/vbaProject.bin' in original.namelist():
                raise ValueError('Expected a calculation workbook without an embedded project.')
            for part in original.namelist():
                content = original.read(part)
                if part == '[Content_Types].xml':
                    root = etree.fromstring(content)
                    main = root.find("{*}Override[@PartName='/xl/workbook.xml']")
                    if main is None or main.get('ContentType') != WORKBOOK_TYPE:
                        raise ValueError('Expected an ordinary Excel workbook.')
                    main.set('ContentType', MACRO_TYPE)
                    content = etree.tostring(root, encoding='UTF-8', xml_declaration=True, standalone=True)
                output.writestr(part, content)
        with ExcelFile(build) as workbook:
            code = SOURCE.read_text(encoding='ascii').replace('\n', '\r\n')
            workbook.add_vba_project().add_module(MODULE, code)
            workbook.save()
        verify_project(build)
        # One filesystem operation publishes the complete file and refuses to overwrite another run.
        link(build, destination)


if __name__ == '__main__':
    parser = ArgumentParser(description=__doc__)
    parser.add_argument('source', type=Path)
    parser.add_argument('destination', type=Path)
    args = parser.parse_args()
    package_macro(args.source, args.destination)
