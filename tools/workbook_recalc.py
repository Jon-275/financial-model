"""Prepare a disposable Excel package for a clean LibreOffice recalculation."""

from __future__ import annotations

import argparse
from pathlib import Path
import subprocess
from zipfile import BadZipFile, ZIP_DEFLATED, ZipFile

from lxml import etree


_MAIN = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
_CONTENT_TYPES = '{http://schemas.openxmlformats.org/package/2006/content-types}'
_WORKSHEET = 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'


class RecalculationError(RuntimeError):
    """Recalculation failed. Messages exclude workbook values and process output."""


def clear_formula_results(source: Path, destination: Path) -> None:
    """Preserve workbook parts while removing saved results from formula cells."""
    if destination.exists() or source.resolve() == destination.resolve():
        raise RecalculationError('Use a new destination for the recalculation copy.')
    parser = etree.XMLParser(resolve_entities=False, no_network=True, huge_tree=True)
    created = False
    try:
        with ZipFile(source) as original:
            manifest = etree.fromstring(original.read('[Content_Types].xml'), parser)
            worksheets = {part.get('PartName').lstrip('/')
                          for part in manifest.iter(_CONTENT_TYPES + 'Override')
                          if part.get('ContentType') == _WORKSHEET and part.get('PartName')}
            if not worksheets or not worksheets.issubset(original.namelist()):
                raise RecalculationError('The workbook worksheet manifest is incomplete.')
            with ZipFile(destination, 'x', compression=ZIP_DEFLATED, compresslevel=5) as clean:
                created = True
                for entry in original.infolist():
                    part = entry.filename
                    data = original.read(entry)
                    if part in worksheets:
                        sheet = etree.fromstring(data, parser)
                        for cell in sheet.iter(_MAIN + 'c'):
                            if cell.find(_MAIN + 'f') is not None:
                                result = cell.find(_MAIN + 'v')
                                if result is not None:
                                    cell.remove(result)
                                cell.attrib.pop('t', None)
                        data = etree.tostring(sheet, encoding='UTF-8', xml_declaration=True)
                    elif part == 'xl/workbook.xml':
                        book = etree.fromstring(data, parser)
                        calc = book.find(_MAIN + 'calcPr')
                        if calc is None:
                            raise RecalculationError('The workbook is missing its calculation settings.')
                        calc.attrib.update({'calcMode': 'auto', 'fullCalcOnLoad': '1', 'forceFullCalc': '1'})
                        data = etree.tostring(book, encoding='UTF-8', xml_declaration=True)
                    clean.writestr(part, data)
    except (BadZipFile, OSError, KeyError, etree.XMLSyntaxError, RecalculationError) as exc:
        if created:
            destination.unlink(missing_ok=True)
        if isinstance(exc, RecalculationError):
            raise
        raise RecalculationError('Could not prepare an Excel copy for recalculation.') from None


def recalculate(source: Path, target_dir: Path, profile_dir: Path, executable: str, timeout: int) -> Path:
    """Calculate the disposable input with an isolated LibreOffice profile."""
    calculated = target_dir / source.name
    if calculated.exists() or calculated.resolve() == source.resolve():
        raise RecalculationError('Use a new LibreOffice output location.')
    try:
        result = subprocess.run(
            [executable, f'-env:UserInstallation={profile_dir.resolve().as_uri()}', '--headless',
             '--convert-to', 'xlsx', '--outdir', str(target_dir.resolve()), str(source.resolve())],
            capture_output=True, timeout=timeout, check=False,
        )
    except subprocess.TimeoutExpired:
        raise RecalculationError('LibreOffice recalculation timed out.') from None
    except OSError:
        raise RecalculationError('Could not run LibreOffice for recalculation.') from None
    if result.returncode or not calculated.is_file():
        raise RecalculationError('LibreOffice did not produce a recalculated workbook.')
    return calculated


def _main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('source', type=Path)
    parser.add_argument('destination', type=Path)
    args = parser.parse_args()
    try:
        clear_formula_results(args.source, args.destination)
    except RecalculationError as exc:
        parser.exit(1, f'{exc}\n')


if __name__ == '__main__':
    _main()
