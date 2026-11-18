"""Recalculate Marketplace coverage years under controlled, unequal inflation rates."""
import argparse
from pathlib import Path
import sys
from tempfile import TemporaryDirectory
from zipfile import ZIP_DEFLATED, ZipFile

from lxml import etree

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'tools'))
from fixtures import COMMON
from native_capture_acceptance import prepare_trial_copy
from risk_engine.native_annual import NativeAnnualModel, price_path
from risk_engine.scenario import load_scenario
from workbook_recalc import recalculate
from xlsx_acceptance import prepare_xlsx
from xlsx_read import Workbook, q, sheet_parts

RATES = (0.10, -0.05, 0.08)
# Opening year, age at first death, filing status, trial guidelines and base-plan guidelines.
CASES = (
    (2026, 999, ('MFJ', 'MFJ', 'MFJ'), (21640, 23804, 22613.8), (21640, 22181, 22735.525)),
    (2026, 55, ('MFJ', 'Single', 'Single'), (21640, 17556, 16678.2), (21640, 16359, 16767.975)),
    (2027, 999, ('MFJ', 'MFJ', 'MFJ'), (22181, 24399.1, 23179.145),
     (22181, 22735.525, 23303.913125)),
    (2030, 54, ('Single', 'Single', 'Single'),
     (17616.853734375, 19378.5391078125, 18409.612152421875),
     (17616.853734375, 18057.275077734375, 18508.706954677734375)),
)


def equal(label, actual, expected):
    if isinstance(expected, str):
        matches = actual == expected
    else:
        matches = type(actual) in (int, float) and abs(actual - expected) <= 0.0000001
    if not matches:
        raise AssertionError(f'{label}: expected {expected!r}, found {actual!r}')


def set_test_inflation(source, destination):
    """Replace three exogenous rates in the disposable workbook, retaining all dependent formulas."""
    with ZipFile(source) as original, ZipFile(destination, 'x', ZIP_DEFLATED) as output:
        part = sheet_parts(original)['Market_Paths']
        for entry in original.infolist():
            payload = original.read(entry.filename)
            if entry.filename == part:
                xml = etree.fromstring(payload)
                for row, rate in enumerate(RATES, 7):
                    cell = xml.find(f".//{q('c')}[@r='F{row}']")
                    formula = cell.find(q('f')) if cell is not None else None
                    if formula is None:
                        raise AssertionError(f'Expected an annual inflation formula in Market_Paths!F{row}')
                    cell.remove(formula)
                    cached = cell.find(q('v'))
                    if cached is None:
                        cached = etree.SubElement(cell, q('v'))
                    cached.text = str(rate)
                    cell.set('t', 'n')
                payload = etree.tostring(xml)
            output.writestr(entry, payload)


def check_case(source, folder, case, office):
    opening, death, filing, guidelines, base_guidelines = case
    inputs = {**COMMON['Inputs'], 'B130': opening, 'B132': .025, 'D14': .025, 'D37': 100000,
              'D52': 85000, 'D53': 57, 'D54': death, 'D101': 1, 'D102': 1, 'D103': 0,
              'D106': 25000, 'D116': 1, 'D117': 0, 'D119': 24000}
    changes = {'Inputs': inputs, 'Risk_Settings': {'B29': 1, 'B32': 1},
               'Risk_Trials': {'B3': 1}, 'Risk_Data': {'Z6': 'running'}}
    scenario, trial = (folder / (name + source.suffix) for name in ('scenario', 'trial'))
    prepare_xlsx(source, scenario, changes, compact=True)
    set_test_inflation(scenario, trial)
    output = folder / 'calculated'
    output.mkdir()
    result = recalculate(trial, output, folder / 'profile', office, 600)
    model = NativeAnnualModel(load_scenario(result))
    path = price_path(model, RATES, end_age=57)
    # Changing a coverage year's shock must leave that same year's guideline unchanged.
    for index, rates in enumerate(((-.02, RATES[1], RATES[2]), (RATES[0], .08, RATES[2]))):
        modified = price_path(model, rates, end_age=57)
        equal('Python same-year shock', modified.periods[index].aca[2], guidelines[index])
    with Workbook(result) as book:
        health = book['Trial_Healthcare']
        projection = book['Trial_Projection']
        prior_prices = 1.025 ** (opening - 2026)
        equal('Opening balance year', health['A5'].value, opening)
        equal('Trial status', book['Risk_Trials']['C7'].value, 1)
        for sheet in (book['Healthcare'], health):
            for column in ('M', 'N', 'O', 'R'):
                cell = sheet[f'{column}5']
                if type(cell.value) not in (int, float) or cell.value != 0:
                    raise AssertionError(f'{sheet.title}!{cell.coordinate} must be numeric zero in the opening row.')
        for index, (filers, fpl, base) in enumerate(zip(filing, guidelines, base_guidelines, strict=True)):
            row, year = 6 + index, opening + 1 + index
            equal('Coverage year', health[f'A{row}'].value, year)
            equal('Year of the guideline price index', projection[f'B{row-1}'].value, year - 1)
            equal('Prior year price index', projection[f'H{row-1}'].value, prior_prices)
            prior_prices *= 1 + RATES[index]
            equal('Trial filing', book['Trial_Projection'][f'F{row}'].value, filers)
            equal('Annual inflation', book['Market_Paths'][f'F{row+1}'].value, RATES[index])
            equal('Base guideline', book['Healthcare'][f'O{row}'].value, base)
            equal('Trial guideline', health[f'O{row}'].value, fpl)
            period = path.periods[index]
            equal('Python guideline', period.aca[2], fpl)
            premium, benchmark, magi = (health[f'{col}{row}'].value for col in ('M', 'N', 'AC'))
            if min(premium, benchmark) <= 0 or magi < 3 * fpl:
                raise AssertionError('Fixture must have Marketplace premiums and income in or above the 300% band.')
            # At 300%–400% the 2027 contribution rate is 10.22%. Above 400% credit is zero.
            credit = min(premium, max(0, benchmark - .1022 * magi)) if magi <= 4 * fpl else 0
            equal('Marketplace credit', health[f'R{row}'].value, credit)
            equal('Credit funding spending', book['Trial_Cash_Flow'][f'AC{row}'].value, credit)
            equal('Python credit', period.year_tax.premium_credit(magi, *period.aca), credit)
            if opening == 2026 and death == 999 and year == 2028:
                equal('Credit lost under the former current-year index', credit, 16002.085)
                if magi <= 4 * 21640 * .95:
                    raise AssertionError('Fixture did not distinguish the two sides of the 400% boundary.')
        print(f'PASS opening year {opening}, filing {" / ".join(filing)}, three coverage years.', flush=True)


def verify(workbook, office):
    with TemporaryDirectory() as temp:
        root = Path(temp)
        cleared = root / ('template' + workbook.suffix)
        prepare_trial_copy(workbook, cleared)
        for index, case in enumerate(CASES):
            folder = root / str(index)
            folder.mkdir()
            check_case(cleared, folder, case, office)


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('workbook', type=Path)
    parser.add_argument('--soffice', default='soffice')
    args = parser.parse_args()
    verify(args.workbook, args.soffice)
