"""Verify the packaged Start Here page and recalculate its published formulas."""
import argparse
from math import isclose
from pathlib import Path
import subprocess
import sys
import tempfile
from zipfile import ZipFile

from lxml import etree as ET
from openpyxl import load_workbook
from openpyxl.utils.cell import get_column_letter, range_boundaries

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'tools'))
from author_model import author
from starter_illustration import EDITABLE_CELLS
from xlsx_read import q, sheet_parts


def read_formulas(source):
    with ZipFile(source) as package:
        page = ET.fromstring(package.read(sheet_parts(package)['Start Here']))
    controls = {rule.get('sqref') for rule in page.iter(q('dataValidation'))}
    assert controls == set(EDITABLE_CELLS), controls
    links = {link.get('ref'): link.get('location') for link in page.iter(q('hyperlink'))}
    assert links['B18'] == "'Start Here'!B22" and links['B46'] == 'Inputs!A1'
    chart = page.find(q('conditionalFormatting'))
    assert chart.get('sqref') == 'E41:E42'
    bounds = [(item.get('type'), item.get('val')) for item in chart.iter(q('cfvo'))]
    assert bounds == [('num', '0'), ('max', None)], bounds
    formulas = {cell.get('r'): '=' + cell.findtext(q('f')) for cell in page.iter(q('c'))
                if cell.find(q('f')) is not None}
    occupied = set()
    for merge in page.iter(q('mergeCell')):
        left, top, right, bottom = range_boundaries(merge.get('ref'))
        first = f'{get_column_letter(left)}{top}'
        for col in range(left, right + 1):
            for row in range(top, bottom + 1):
                address = f'{get_column_letter(col)}{row}'
                assert address not in occupied, f'Overlapping merges at {address}'
                occupied.add(address)
                if address != first:
                    assert address not in formulas and address not in controls, address
    return formulas


def expected(inputs):
    age, pay, saving_rate, annual_return, retire = inputs
    deposit = pay * saving_rate
    amounts = {'E32': deposit}
    for column, wait, chart_row in (('F', 0, 41), ('H', 5, 42)):
        years = max(0, retire - age - wait)
        balance = 0
        for _ in range(years):
            balance = balance * (1 + annual_return) + deposit
        amounts.update({f'{column}35': years, f'{column}36': deposit * years,
                        f'{column}37': balance - deposit * years, f'{column}38': balance,
                        f'E{chart_row}': balance})
    return amounts


def verify(source, executable='soffice'):
    formulas = read_formulas(source)
    base = (30, 50000, .1, .04, 65)
    good = {
        'example': base,
        'zero return': (30, 50000, .1, 0, 65),
        'negative return': (30, 50000, .1, -.05, 65),
        'annual wipeout': (30, 50000, .1, -1, 65),
        'same age': (65, 50000, .1, .04, 65),
        'short delay': (62, 50000, .1, .04, 65),
        'zero pay': (30, 0, 1, .04, 65),
        'no savings': (0, 50000, 0, .04, 125),
    }
    bad = {
        'missing': ((None, *base[1:]), 'number'),
        'text paste': ((30, 'pay', *base[2:]), 'number'),
        'fractional age': ((30.5, *base[1:]), 'current age'),
        'negative pay': ((30, -1, *base[2:]), 'annual pay'),
        'above 100 percent saved': ((30, 50000, 1.01, .04, 65), 'saving percentage'),
        'return below minus 100': ((30, 50000, .1, -1.01, 65), 'annual return'),
        'early end': ((*base[:4], 25), 'retirement age'),
        'over age limit': ((*base[:4], 126), 'retirement age'),
        'overflow': ((0, 50000, 1, 1e50, 125), 'range'),
        'deposit total overflow': ((0, 1e308, 1, -1, 125), 'range'),
    }
    cases = {**good, **{name: inputs for name, (inputs, _) in bad.items()}}
    with tempfile.TemporaryDirectory(prefix='retirement-starter-') as temporary:
        work = Path(temporary)
        book = work / 'illustration.xlsx'
        output = work / 'recalculated'
        output.mkdir()
        cells = {name: {**formulas, **dict(zip(EDITABLE_CELLS, values, strict=True))}
                 for name, values in cases.items()}
        author({'order': list(cases), 'cells': cells}, book)
        subprocess.run([executable, f'-env:UserInstallation={(work / "profile").as_uri()}',
                        '--headless', '--convert-to', 'xlsx', '--outdir', str(output), str(book)],
                       check=True, capture_output=True, timeout=90)
        recalculated = load_workbook(output / book.name, data_only=True, read_only=True)
        try:
            for name, inputs in good.items():
                sheet = recalculated[name]
                assert sheet['B31'].value in (None, ''), (name, sheet['B31'].value)
                for address, value in expected(inputs).items():
                    actual = sheet[address].value
                    assert isinstance(actual, (int, float)) and isclose(actual, value, abs_tol=.000001), (name, address, actual, value)
            for name, (_, message) in bad.items():
                sheet = recalculated[name]
                assert message in str(sheet['B31'].value), (name, sheet['B31'].value)
                for address in expected(base):
                    assert sheet[address].value in (None, ''), (name, address, sheet[address].value)
        finally:
            recalculated.close()
    print(f'Start Here layout and {len(cases)} saving illustrations verified.')


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('workbook', type=Path)
    parser.add_argument('--soffice', default='soffice')
    options = parser.parse_args()
    verify(options.workbook, options.soffice)
