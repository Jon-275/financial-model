"""Exercise the workbook risk refresh and live assumptions in LibreOffice."""

import argparse
from hashlib import sha256
import os
from pathlib import Path
import re
import subprocess
import sys
from tempfile import TemporaryDirectory
from xml.etree import ElementTree as E
from zipfile import ZipFile

from dashboard_acceptance import validate_dashboard
from xlsx_acceptance import prepare_xlsx, recalculate
from xlsx_read import Workbook, sheet_parts


ROOT = Path(__file__).resolve().parents[1]
RISK = 'Risk_Data'
CHART = '{http://schemas.openxmlformats.org/drawingml/2006/chart}'


def check_state(workbook, expected):
    with Workbook(workbook) as book:
        assert book[RISK]['Z1'].value == expected, (workbook.name, expected, book[RISK]['Z1'].value)
        validate_dashboard(book, tolerance=1e-7)


def check_refresh(workbook, office, timeout):
    check_state(workbook, 'current')
    with Workbook(workbook) as book:
        inputs = book['Inputs']
        case = inputs['B4'].value
        active = 'DEF'[case - 1]
        inactive = 'DEF'[case % 3]
        inactive_age = inputs[f'{inactive}8'].value
        basis = book['Risk_Settings']['B9'].value
        disability_rate = book['Tax_Tables']['B251'].value
        profile_start = 26 + 9 * (inputs['B15'].value - 1)
        profile_weight = book['Asset_Allocation'][f'B{profile_start}'].value
        assert book[RISK]['BB3'].value == inputs['B7'].value
        assert book[RISK]['BB5'].value == 1000
        assert book['Summary']['F12'].value == book[RISK]['BB8'].value

    with TemporaryDirectory(prefix='retirement-risk-acceptance-') as temp:
        work = Path(temp)
        source = work / 'source.xlsx'
        different_hsa_mix = {f'E{profile_start + offset}': int(offset == 2) for offset in range(5)}
        prepare_xlsx(workbook, source, {'Inputs': {f'{active}31': 10000, f'{active}53': 125},
                                       'Asset_Allocation': different_hsa_mix})
        original = source.read_bytes()

        refreshed = work / "Family's plan.xlsx"
        command = [sys.executable, str(ROOT / 'tools/run_risk.py'), str(source),
                   '--output', str(refreshed), '--trials', '3', '--seed', '2026',
                   '--soffice', office, '--timeout', str(timeout)]
        subprocess.run(command, check=True, timeout=3 * timeout)
        assert source.read_bytes() == original, 'Risk refresh wrote to the source workbook'
        check_state(refreshed, 'current')
        with Workbook(refreshed) as book:
            risk, projection = book[RISK], book['Projection']
            assert risk['BB4'].value == 125
            assert risk['BB5'].value == 3
            assert risk['BB19'].value == sha256(original).hexdigest()
            assert risk['B5'].value == projection['C5'].value
            end_row = 5 + 125 - risk['B5'].value
            assert risk[f'B{end_row}'].value == 125
            assert risk[f'B{end_row + 1}'].value is None
            for row in (5, 6, 7):
                assert type(risk[f'AG{row}'].value) in (int, float)
        with ZipFile(source) as before, ZipFile(refreshed) as after:
            assert set(before.namelist()) == set(after.namelist())
            worksheets = set(sheet_parts(before).values())
            for part in before.namelist():
                if part not in worksheets and not re.fullmatch(r'xl/charts/chart\d+\.xml', part):
                    assert before.read(part) == after.read(part), ('Changed workbook package part', part)
            named_series = [
                formula.text for part in after.namelist() if re.fullmatch(r'xl/charts/chart\d+\.xml', part)
                for formula in E.fromstring(after.read(part)).iter(CHART + 'f') if '!RMChart' in (formula.text or '')
            ]
            assert named_series and all(name.startswith("'Report_Data'!") for name in named_series)

        published = refreshed.read_bytes()
        collision = subprocess.run(command, capture_output=True, timeout=15, check=False)
        assert collision.returncode and 'must name a new' in collision.stderr.decode()
        assert refreshed.read_bytes() == published, 'Refresh replaced an existing report'

        output = work / 'calculated'
        output.mkdir()
        variants = (
            ('active', {'Inputs': {f'{active}53': 124}}, 'stale'),
            ('inactive', {'Inputs': {f'{inactive}8': inactive_age + 1}}, 'current'),
            ('future-dollars', {'Summary': {'K18': 'Future dollars'}}, 'current'),
            ('asset-mix', {'Asset_Allocation': {f'B{profile_start}': profile_weight + 0.01}}, 'stale'),
            ('returns', {'Risk_Settings': {'B9': 3 - basis}}, 'stale'),
            ('tax', {'Tax_Tables': {'B251': disability_rate + 0.00000006}}, 'stale'),
            ('scenario', {'Inputs': {'B4': case % 3 + 1}}, 'stale'),
        )
        for name, changes, state in variants:
            changed = work / f'{name}.xlsx'
            prepare_xlsx(refreshed, changed, changes)
            candidate = recalculate(changed, output, work / 'profile', office, timeout)
            check_state(candidate, state)
            if name == 'asset-mix':
                with Workbook(candidate) as book:
                    assert book['Inputs']['B71'].value == 'Check the selected mixes on Asset_Allocation'
    print('PASS workbook risk refresh, stale results, scenario switches, and source preservation')


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('workbook', type=Path)
    parser.add_argument('--soffice', default=os.environ.get('SOFFICE_BIN', 'soffice'))
    parser.add_argument('--timeout', type=int, default=240, help='Seconds allowed for each recalculation')
    args = parser.parse_args()
    check_refresh(args.workbook.resolve(), args.soffice, args.timeout)


if __name__ == '__main__':
    main()
