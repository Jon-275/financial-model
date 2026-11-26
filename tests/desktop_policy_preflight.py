"""Package the state policy preflight on a disposable fictional workbook."""

from argparse import ArgumentParser
from pathlib import Path

from openpyxl import load_workbook

from native_capture import CONTEXT_PREFIX
from native_results import ENGINE, SCHEMA
from package_macro import SOURCE, package_macro
from risk_engine.scenario import workbook_has_explicit_state_policy
from state_policy import (
    CALIFORNIA_POLICY, LEGACY_CA_POLICY_HEADERS, WORKBOOK_CA_POLICY_VERSION_CELL, WORKBOOK_JURISDICTION_LABEL,
    WORKBOOK_JURISDICTION_NOTE,
    WORKBOOK_JURISDICTION_ROW, WORKBOOK_STATE_POLICY_NATIVE_CONTEXT, WORKBOOK_STATE_POLICY_NATIVE_SCHEMA,
    WORKBOOK_STATE_POLICY_NATIVE_SUFFIX, WORKBOOK_STATE_POLICY_SCHEMA, WORKBOOK_STATE_POLICY_SCHEMA_CELL,
    WORKBOOK_STATE_POLICY_TABLE_LABELS,
)
from vba_declarations import audit_vba_declarations


FRAGMENT = SOURCE.parent / 'tests' / 'RunnerStatePreflight.fragment'
BEGIN = (
    '    runBegan = True\n'
    '    results.Range("Z6").Value2 = "running"\n'
    '    results.Range("Z9").Value2 = "Running market risk. Use Cancel or press Esc to stop."\n'
)


def declarations() -> str:
    """Generate the VBA literals from the same contract as the workbook author."""
    (schema_cell, schema_label), (version_cell, version_label) = WORKBOOK_STATE_POLICY_TABLE_LABELS
    (input_sheet, input_cell, input_label), (ledger_sheet, ledger_cell, ledger_label) = LEGACY_CA_POLICY_HEADERS
    constants = {
        'NATIVE_SCHEMA': WORKBOOK_STATE_POLICY_NATIVE_SCHEMA,
        'NATIVE_CONTEXT': WORKBOOK_STATE_POLICY_NATIVE_CONTEXT,
        'LEGACY_CONTEXT': CONTEXT_PREFIX,
        'CONTEXT_SUFFIX': WORKBOOK_STATE_POLICY_NATIVE_SUFFIX,
        'SCHEMA_CELL': WORKBOOK_STATE_POLICY_SCHEMA_CELL,
        'SCHEMA': WORKBOOK_STATE_POLICY_SCHEMA,
        'SCHEMA_LABEL_CELL': schema_cell,
        'SCHEMA_LABEL': schema_label,
        'VERSION_CELL': WORKBOOK_CA_POLICY_VERSION_CELL,
        'VERSION': CALIFORNIA_POLICY.version,
        'VERSION_LABEL_CELL': version_cell,
        'VERSION_LABEL': version_label,
        'JURISDICTION': CALIFORNIA_POLICY.jurisdiction,
        'INPUT_LABEL': WORKBOOK_JURISDICTION_LABEL,
        'INPUT_NOTE': WORKBOOK_JURISDICTION_NOTE,
        'LEGACY_INPUT_SHEET': input_sheet,
        'LEGACY_INPUT_CELL': input_cell,
        'LEGACY_INPUT_HEADER': input_label,
        'LEGACY_LEDGER_SHEET': ledger_sheet,
        'LEGACY_LEDGER_CELL': ledger_cell,
        'LEGACY_LEDGER_HEADER': ledger_label,
    }
    lines = [f'Private Const STATE_INPUT_ROW As Long = {WORKBOOK_JURISDICTION_ROW}']
    for key, value in constants.items():
        literal = value.replace('"', '""')
        lines.append(f'Private Const STATE_{key} As String = "{literal}"')
    return '\n'.join(lines) + '\n'


def overlay(source: str) -> str:
    """Gate a test run before publishing status and retain existing rollback."""
    source = source.replace('\r\n', '\n').rstrip() + '\n'
    engine = ('    RequireText results.Range("Z4").Value2, EXPECTED_ENGINE, '
              '"The workbook\'s risk engine version does not match the runner."\n')
    schema = ('    RequireText results.Range("Z2").Value2, EXPECTED_SCHEMA, '
              '"Use the workbook and runner from the same release."\n')
    calculation = '    Application.Calculation = xlCalculationManual\n'
    context = '    context = AssumptionRecord(book)\n'
    cleanup = '    selectorRestored = True\n'
    changes = (
        ('Option Explicit\n', 'Option Explicit\n\n' + declarations()),
        ('    Dim seed As Double, context As String, code As Long, index As Long\n',
         '    Dim seed As Double, context As String, code As Long, index As Long, jurisdiction As String\n'),
        (schema + engine, engine + '    jurisdiction = ReadWorkbookStatePolicy(book)\n'),
        (calculation, calculation + '    RecalculateStateSelection book, jurisdiction\n'),
        (BEGIN, ''),
        (context, context + '    RequireStatePolicyContext book, jurisdiction, context\n'
         '    CheckCancelled\n' + BEGIN + '    results.Range("Z1").Calculate\n'),
        (cleanup, cleanup + '    RequireText ReadWorkbookStatePolicy(book), jurisdiction, '
         '"The state tax policy changed during the run."\n'
         '    RequireStatePolicyContext book, jurisdiction, AssumptionRecord(book)\n'),
    )
    for before, after in changes:
        if source.count(before) != 1:
            raise ValueError(f'Expected one state preflight anchor for {before.strip().splitlines()[0]!r}.')
        source = source.replace(before, after, 1)
    result = source + '\n' + FRAGMENT.read_text(encoding='ascii')
    audit_vba_declarations(result)
    return result


def verify_pairing(draft: Path) -> None:
    """Check that the raw format and assumption record are supported by the overlay."""
    workbook = load_workbook(draft, read_only=True, keep_links=False)
    try:
        required = {'Inputs', 'Tax_Tables', 'Projection', 'Risk_Data', 'Market_Paths'}
        if not required.issubset(workbook.sheetnames) or workbook['Risk_Data']['Z4'].value != ENGINE:
            raise ValueError('Use a native workbook draft that matches the test runner engine.')
        schema = workbook['Risk_Data']['Z2'].value
        if schema == WORKBOOK_STATE_POLICY_NATIVE_SCHEMA:
            tables = workbook['Tax_Tables']
            expected = (*WORKBOOK_STATE_POLICY_TABLE_LABELS,
                        (WORKBOOK_STATE_POLICY_SCHEMA_CELL, WORKBOOK_STATE_POLICY_SCHEMA),
                        (WORKBOOK_CA_POLICY_VERSION_CELL, CALIFORNIA_POLICY.version))
            if any(tables[address].value != value for address, value in expected):
                raise ValueError('The v2 draft needs complete state tax schema and policy markers.')
            prefix = WORKBOOK_STATE_POLICY_NATIVE_CONTEXT
        elif schema == SCHEMA and not workbook_has_explicit_state_policy(workbook):
            if any(workbook[sheet][address].value != value for sheet, address, value in LEGACY_CA_POLICY_HEADERS):
                raise ValueError('Cannot identify the legacy California tax layout.')
            prefix = CONTEXT_PREFIX
        else:
            raise ValueError('The test overlay requires a supported native schema for this workbook layout.')
        formula = workbook['Market_Paths']['BB3']
        if formula.data_type != 'f' or not formula.value.startswith(f'="{prefix}'):
            raise ValueError('The draft assumption record formula does not match the native risk schema.')
    finally:
        workbook.close()


def prepare(draft: Path, destination: Path) -> Path:
    """Package a new scratch workbook without changing the published runner."""
    verify_pairing(draft)
    package_macro(draft, destination, test_runner_source=overlay(SOURCE.read_text(encoding='ascii')))
    return destination


if __name__ == '__main__':
    parser = ArgumentParser(description=__doc__)
    parser.add_argument('draft', type=Path)
    parser.add_argument('destination', type=Path)
    args = parser.parse_args()
    print(prepare(args.draft, args.destination))
