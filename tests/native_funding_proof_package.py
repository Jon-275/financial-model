"""Install a test-only funding witness bank in a pinned fictional workbook."""

from argparse import ArgumentParser
from collections import defaultdict
from functools import cache
from hashlib import file_digest
from os import link
from pathlib import Path
import re
from tempfile import TemporaryDirectory
from zipfile import ZipFile

from lxml import etree
from openpyxl.formula.tokenizer import Tokenizer
from openpyxl.utils.cell import column_index_from_string, get_column_letter, range_boundaries

from funding_search import ROWS
from native_funding_witness_formulas import FIRST_YEAR, LAST_YEAR, witness_formulas
from xlsx_values import X, _read_xml_part, sheet_parts


PINNED_INPUT_SHA256 = '2626763f5abdfdbdcb2f72b79c9439955e4e78ac02dce21ba2f72c26515788d1'
SHEET = 'Trial_Checks'
OFFSET = 52
BANK = (53, FIRST_YEAR, 94, LAST_YEAR)
CELL = re.compile(r'(\$?)([A-Z]{1,3})(\$?[1-9][0-9]*)')


def _shift_address(address: str) -> str:
    """Move one local A1 address while maintaining its absolute markers."""
    match = CELL.fullmatch(address)
    if match is None:
        raise ValueError(f'Unsupported unqualified proof reference {address!r}.')
    anchor, column, row = match.groups()
    return f'{anchor}{get_column_letter(column_index_from_string(column) + OFFSET)}{row}'


def shift_formula(formula: str) -> str:
    """Move local proof cells to BA:CP without changing live qualified references."""
    tokens = Tokenizer(formula)
    for token in tokens.items:
        if token.type == 'OPERAND' and token.subtype == 'RANGE' and '!' not in token.value:
            token.value = ':'.join(_shift_address(address) for address in token.value.split(':'))
    return tokens.render()


def _no_conversion(row: int) -> str:
    """Certify zero conversion from year inputs, independent of the selected solver mode."""
    inputs, projection = 'Trial_Inputs!', 'Trial_Projection!'
    request, owner, last_age, order = (f'{inputs}$B${number}' for number in (47, 97, 50, 48))
    phase, sepp, headroom = (f'{projection}${column}${row}' for column in 'ERU')
    opening_traditional = f'{projection}$BE${row - 1}'
    age = f'IF({owner}=1,{projection}$C${row},{projection}$D${row})'
    return (
        f'OR(AND(ISNUMBER({request}),{request}=0),EXACT({phase},"Working"),'
        f'AND(ISNUMBER({sepp}),{sepp}>0),'
        f'AND(ISNUMBER({owner}),OR({owner}=1,{owner}=2),ISNUMBER({age}),ISNUMBER({last_age}),{age}>{last_age}),'
        f'AND(ISNUMBER({opening_traditional}),{opening_traditional}=0),'
        f'AND({order}=2,ISNUMBER({headroom}),{headroom}=0))'
    )


@cache
def bank_formulas() -> dict[str, str]:
    """Bind generated quote rows to the selected draw in the real native solver."""
    formulas = {_shift_address(address): shift_formula(formula) for address, formula in witness_formulas().items()}
    for row in range(FIRST_YEAR, LAST_YEAR + 1):
        final = 1 + (row - FIRST_YEAR + 1) * ROWS
        active = f'AND(ISNUMBER(Trial_Inputs!$B$70),{row}-4<=Trial_Inputs!$B$70)'
        formulas[f'BA{row}'] = f'=IF({active},Trial_Funding_Calc!$AC${final},"")'
        formulas[f'BB{row}'] = f'=IF({active},Trial_Funding_Calc!$D${final},"")'
        formulas[f'BC{row}'] = f'=IFERROR(IF(AND({active},{_no_conversion(row)}),0,""),"")'
    return formulas


def _overlaps(ref: str) -> bool:
    """Recognize any pre-existing range that would intersect the proof bank."""
    left, top, right, bottom = range_boundaries(ref)
    bank_left, bank_top, bank_right, bank_bottom = BANK
    return left <= bank_right and right >= bank_left and top <= bank_bottom and bottom >= bank_top


def _install(root: etree._Element) -> bytes:
    """Add formulas with empty caches and preserve the existing check worksheet."""
    rows = {int(row.get('r')): row for row in root.find(X + 'sheetData').findall(X + 'row')}
    for element in (*root.iter(X + 'mergeCell'), *root.iter(X + 'f')):
        if element.get('ref') and _overlaps(element.get('ref')):
            raise ValueError('Existing worksheet content overlaps the funding proof bank.')
    additions = defaultdict(list)
    for address, formula in bank_formulas().items():
        column, row, _, _ = range_boundaries(address)
        additions[row].append((column, address, formula))
    for number, entries in additions.items():
        if number not in rows:
            raise ValueError(f'Missing trial check row {number}.')
        row = rows[number]
        existing = {range_boundaries(cell.get('r'))[0]: cell for cell in row.findall(X + 'c')}
        if len(existing) != len(row):
            raise ValueError(f'Unexpected existing cell layout in trial check row {number}.')
        if any(BANK[0] <= column <= BANK[2] for column in existing):
            raise ValueError(f'Existing cells overlap the funding proof bank in row {number}.')
        for column, address, formula in entries:
            cell = etree.Element(X + 'c', r=address)
            etree.SubElement(cell, X + 'f').text = formula[1:]
            existing[column] = cell
        row[:] = [existing[column] for column in sorted(existing)]
        row.attrib.pop('spans', None)
    dimension = root.find(X + 'dimension')
    left, top, right, bottom = range_boundaries(dimension.get('ref'))
    dimension.set('ref', f'{get_column_letter(min(left, BANK[0]))}{min(top, BANK[1])}:'
                  f'{get_column_letter(max(right, BANK[2]))}{max(bottom, BANK[3])}')
    return etree.tostring(root, encoding='UTF-8', xml_declaration=True, standalone=True)


def _require_source(source: Path) -> None:
    """Limit this disposable overlay to the reviewed 100-pair input bytes."""
    if source.is_symlink() or not source.is_file() or source.suffix.lower() != '.xlsm':
        raise ValueError('Use the pinned fictional 100-pair full-range workbook.')
    with source.open('rb') as stream:
        if file_digest(stream, 'sha256').hexdigest() != PINNED_INPUT_SHA256:
            raise ValueError('Use the pinned fictional 100-pair full-range workbook.')


def verify_package(source: Path, candidate: Path) -> Path:
    """Require the exact worksheet overlay and untouched bytes in every other part."""
    _require_source(source)
    if candidate.is_symlink() or not candidate.is_file() or source.samefile(candidate):
        raise ValueError('Use a separate regular proof workbook.')
    with ZipFile(source) as original, ZipFile(candidate) as modified:
        part = sheet_parts(original)[SHEET]
        if original.namelist() != modified.namelist() or original.comment != modified.comment:
            raise ValueError('The proof package changed the workbook part inventory.')
        expected = _install(_read_xml_part(original, part, worksheet=True))
        if modified.read(part) != expected:
            raise ValueError('The proof worksheet differs from the generated bank.')
        for name in original.namelist():
            if name != part and original.read(name) != modified.read(name):
                raise ValueError(f'The proof package changed an unrelated part: {name}.')
    return candidate


def prepare_package(source: Path, destination: Path) -> Path:
    """Write a verified bank-only workbook to a new output path."""
    if destination.exists() or destination.is_symlink():
        raise FileExistsError(destination)
    if destination.suffix.lower() != '.xlsm':
        raise ValueError('Use an .xlsm destination for the proof workbook.')
    _require_source(source)
    destination.parent.mkdir(parents=True, exist_ok=True)
    with TemporaryDirectory(dir=destination.parent) as directory:
        staged = Path(directory) / 'proof.xlsm'
        with ZipFile(source) as original, ZipFile(staged, 'x') as modified:
            part = sheet_parts(original)[SHEET]
            payload = _install(_read_xml_part(original, part, worksheet=True))
            modified.comment = original.comment
            for info in original.infolist():
                modified.writestr(info, payload if info.filename == part else original.read(info))
        verify_package(source, staged)
        link(staged, destination)
    return destination


if __name__ == '__main__':
    parser = ArgumentParser(description=__doc__)
    parser.add_argument('source', type=Path)
    parser.add_argument('destination', type=Path)
    parser.add_argument('--check', action='store_true', help='Verify an existing bank-only workbook.')
    options = parser.parse_args()
    action = verify_package if options.check else prepare_package
    print(action(options.source, options.destination))
