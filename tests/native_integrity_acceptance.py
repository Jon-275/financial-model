"""Recalculate completion checks after editing individual stored trial IDs."""
import argparse
from pathlib import Path
import random
import subprocess
import sys
from tempfile import TemporaryDirectory

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'tools'))
from openpyxl import Workbook, load_workbook
import native_capture as trial
import native_funding_paths as funding
from native_results import _ranks, _status


def fixture(path, n, edit=None):
    wb = Workbook()
    result = wb.active
    result.title = 'Risk_Data'
    primary = wb.create_sheet(trial.SHEET)
    secondary = wb.create_sheet(funding.DATA_SHEET)
    settings = wb.create_sheet('Risk_Settings')
    market = wb.create_sheet('Market_Paths')
    settings['B29'], settings['B30'] = n, 2026
    market['AL3'], market['BB3'] = 'OK', 'synthetic-context'
    for address, value in {'Z6': 'complete', 'Z8': 'synthetic-context',
                           'Z10': n, 'Z11': 2026, 'AB3': f'=COUNTIF(Risk_Trials!C8:C{trial.LAST},1)'}.items():
        result[address] = value
    def put(address, value, _style):
        result[address] = value
    _status(put)
    codes = [*range(1, n+1), *range(1001, 1001+n), *range(2001, 2014),
             *range(2101, 2114), *range(2201, 2208)]
    rng = random.Random(2026)
    keys = []
    for offset, code in enumerate(codes):
        r = trial.FIRST + offset
        primary[f'A{r}'] = code
        primary[f'B{r}'] = primary[f'{trial.SIGNATURE}{r}'] = 'synthetic-context'
        primary[f'C{r}'] = 1
        if offset < n:
            balance = (0 if offset % 4 == 0 else rng.random()*900_000) if n == 1000 else n-offset
            age = 65+offset%12 if n == 1000 and offset%4 == 0 else None
            gap = rng.random()*20_000 if age else 0
            for col, value in (('G',balance),('E',age),('F',gap)):
                primary[f'{col}{r}'] = value
            keys.append((balance, age if age is not None else float('inf'), -gap, code))
    class Writer:
        def put(self, sheet, address, value, _style):
            wb[sheet][address] = value
    _ranks(Writer())
    ordered = [entry[-1] for entry in sorted(keys)]
    for row, pct in enumerate((10, 50, 90), 8):
        rank = 1+(pct*(n-1)+50)//100
        secondary[f'A{row}'] = ordered[rank-1]
        secondary[f'B{row}'] = secondary[f'{funding.SIGNATURE}{row}'] = 'synthetic-context'
        secondary[f'C{row}'] = 1
    if edit:
        sheet, address, value = edit
        wb[sheet][address] = value
    wb.save(path)
    return ordered


def verify(executable):
    cases = [('minimum',1,None), ('regular',3,None), ('capacity',1000,None),
             ('duplicate-fixed',3,(trial.SHEET, 'A8', 2)),
             ('missing-flexible',3,(trial.SHEET, 'A11', None)),
             ('misplaced-early',3,(trial.SHEET, 'A14', 2101)),
             ('missing-last-stress',3,(trial.SHEET, 'A46', None)),
             ('extra-id',3,(trial.SHEET, 'A47', 42)),
             ('wrong-funding-path',3,(funding.DATA_SHEET, 'A8', 1)),
             ('missing-funding-path',3,(funding.DATA_SHEET, 'A9', None))]
    with TemporaryDirectory() as temp:
        root = Path(temp)
        out = root/'calculated'
        out.mkdir()
        ranks = {name: fixture(root/f'{name}.xlsx', n, edit) for name, n, edit in cases}
        subprocess.run([executable, f'-env:UserInstallation={(root/"profile").as_uri()}', '--headless',
                        '--convert-to', 'xlsx', '--outdir', str(out), *(str(p) for p in root.glob('*.xlsx'))],
                       capture_output=True, check=True, timeout=180)
        for name, _, edit in cases:
            wb = load_workbook(out/f'{name}.xlsx', read_only=True, data_only=True)
            try:
                if name == 'capacity':
                    column = 1 + trial.SIGNATURE_COLUMN
                    rows = wb[trial.SHEET].iter_rows(min_row=8, max_row=1007, min_col=column,
                                                     max_col=column, values_only=True)
                    expected_ranks = {tid: rank for rank, tid in enumerate(ranks[name], 1)}
                    assert all(row[0] == expected_ranks[tid] for tid, row in enumerate(rows, 1))
                expected = 'incomplete' if edit else 'current'
                assert wb['Risk_Data']['Z1'].value == expected, (name, wb['Risk_Data']['Z1'].value)
                expected_primary = 'incomplete' if edit and edit[0] == trial.SHEET else 'current'
                assert wb['Risk_Data']['Z5'].value == expected_primary, name
            finally:
                wb.close()
    print(f'PASS {len(cases)} completed-batch and funding-path integrity cases.')


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--soffice', default='soffice')
    verify(parser.parse_args().soffice)
