# Test and release guide

Run commands from the repository root. Use Python 3.11 or later, Node.js, and LibreOffice Calc with `soffice` on your path. For a different LibreOffice path, add `--soffice /path/to/soffice` to the scenario test command.

```sh
python3 -m pip install -r tools/requirements.txt
```

## Run the checks locally

```sh
python3 tests/model_acceptance.py --release workbook/retirement-model.xlsx
PYTHONPATH=tools python3 -m unittest discover -s tests -p 'test_risk_*.py'
node --test tests/gemini_review.test.mjs
```

The release check compares saved cash, taxes, withdrawals and account balances. It also checks the Summary, annual appendix, chart, worksheet visibility and workbook metadata. The risk tests check the simulator against the workbook. The Node tests cover when Gemini should pass or fail a pull request.

For a newly built workbook, use its path in the release and scenario commands. Set the same path for the risk tests.

```sh
candidate=/absolute/path/to/retirement-model.xlsx
RISK_TEST_WORKBOOK="$candidate" PYTHONPATH=tools python3 -m unittest discover -s tests -p 'test_risk_*.py'
```

## Recalculate workbook scenarios

```sh
python3 tests/xlsx_acceptance.py workbook/retirement-model.xlsx \
  tests/phase2-fixtures.json --compact --ledger
python3 tests/xlsx_acceptance.py workbook/retirement-model.xlsx \
  tests/phase3-fixtures.json --compact --ledger
```

Each case runs on a temporary copy in LibreOffice. The runner applies test inputs, clears saved formula results, and checks the recalculated values. Cases use fictional inputs. The Phase 2 file includes the earlier tests for the current workbook. The Phase 3 case checks that required spending can go unpaid in a year even when all taxes are paid.

Use `--case NAME` to run one case, `--jobs 1` if the machine runs low on memory, or `--keep DIRECTORY` to inspect the test copies. The `--compact` option trims inactive calculation rows in temporary copies. For a release, also run representative cases on a full-size copy by omitting `--compact`.

When changing the Phase 2 cases, generate a candidate file in scratch and compare it before updating the tracked copy.

```sh
mkdir -p scratch
python3 tests/phase2_fixtures.py > scratch/phase2-fixtures.json
diff -u tests/phase2-fixtures.json scratch/phase2-fixtures.json
```

## Pull request checks

GitHub requires **Workbook checks**, **Greptile Review** and **Gemini Review** before merging into `main`. Workbook checks run on PRs and after changes merge. They include the release and privacy checks, the risk and review tests, and seven representative recalculation cases. [Workbook workflow](../.github/workflows/workbook-checks.yml) · [Branch rules](https://github.com/Jon-275/financial-model/rules/27295551)

Gemini posts its findings on ready PRs. Critical or high-severity findings and incomplete reviews block merging. Gemini checks source code and workflows. It records a pass when neither changed. A maintainer can retry a review from **Actions → Gemini 4 Hydrogen PR Review → Run workflow**. Select `main` and enter the PR number. [Gemini workflow](../.github/workflows/gemini-review.yml)

## Microsoft Excel release check

Use the candidate workbook from the [build instructions](../README.md#rebuild-the-workbook). Set its path below. The commands create a copy with cleared formula results, using the filename the charts expect.

```sh
candidate=/absolute/path/to/retirement-model.xlsx
mkdir -p scratch
excel_check=$(mktemp -d "$PWD/scratch/excel-check.XXXXXXXX")
python3 tools/workbook_recalc.py "$candidate" "$excel_check/retirement-model.xlsx"
```

1. Open the copy in desktop Excel and check for a repair prompt. In Excel for Mac, press **Control+Option+Shift+F9** to rebuild dependencies and recalculate. Save the copy. [Microsoft Excel keyboard shortcuts](https://support.microsoft.com/en-us/accessibility/excel/keyboard-shortcuts-in-excel)
2. Verify the saved copy with the command below. On a second copy, change the active scenario’s end age and check that the projection and balance chart reach the new final year.
3. Preview **Inputs** and **Summary** with **Active Sheets** selected. Check Letter and A4 for two pages of input decisions and three pages of results. Check the annual appendix for repeated headers and the correct last age. Save the plan as PDF and inspect the chart, table, figures and page breaks.
4. Record the Excel version, checks run and results in the release record.

```sh
python3 tests/model_acceptance.py "$excel_check/retirement-model.xlsx"
```
