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

Check that a market risk refresh preserves the workbook and that changed assumptions hide saved results. This check uses fictional inputs and recalculates disposable copies.

```sh
python3 tests/risk_workbook_acceptance.py workbook/retirement-model.xlsx
```

When changing the Phase 2 cases, generate a candidate file in scratch and compare it before updating the tracked copy.

```sh
mkdir -p scratch
python3 tests/phase2_fixtures.py > scratch/phase2-fixtures.json
diff -u tests/phase2-fixtures.json scratch/phase2-fixtures.json
```

## Pull request checks

GitHub requires **Workbook checks**, **Greptile Review** and **Gemini Review** before merging into `main`. It also requires all review threads to be resolved. Workbook checks run on PRs and after changes merge. They include release and privacy checks, simulator and review tests, representative recalculation cases, and a workbook risk refresh. [Workbook workflow](../.github/workflows/workbook-checks.yml) · [Branch rules](https://github.com/Jon-275/financial-model/rules/27295551)

Gemini posts each finding on the relevant line of a ready PR. Fix the issue or explain why the finding does not apply, then resolve the thread. Critical or high-severity findings and incomplete reviews fail the required check. Address every open thread before merging, including findings from a review that passed. Gemini reviews source and workflow changes and records a pass when the PR contains neither.

**Run Gemini reviewer** is the GitHub Actions job. **Gemini Review** is the result it posts on the PR commit and the check GitHub requires. A title or description edit can create a skipped job entry. Changing the PR base to `main` starts a review.

A maintainer can retry a review from **Actions → Gemini 4 Hydrogen PR Review → Run workflow**. Select `main` and enter the PR number. [Gemini workflow](../.github/workflows/gemini-review.yml)

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
3. Verify that the fictional example opens with the success rate, shaded range, and funding examples already visible on **Summary**. On a test copy, change an active assumption and check that the figures and chart range clear. Restore the assumption and check that they return. Refresh the test copy and confirm the results stay in its **Summary**. [Risk refresh instructions](../docs/risk-analysis.md#refresh-the-workbook)
4. Preview **Summary** in Excel for Mac and Windows. With **Active Sheets** selected, verify that Letter and A4 each show three results pages followed by two decision pages. Preview **Annual Detail** separately. Check the selected dollar view, repeating headers and the last age. Save a PDF and inspect the figures and page breaks. [Print setup](../docs/model-guide.md#print-your-plan)
5. Record the Excel versions, checks run, results and any pending platform checks in the release record.

```sh
python3 tests/model_acceptance.py "$excel_check/retirement-model.xlsx"
```

Before publishing, confirm that the candidate came from the build template and that **Risk_Data** contains only the seeded fictional run. After reviewing the candidate, use the command below to update `EXAMPLE_RISK_SHA256` in `tools/check_workbook_commit.py`. The commit hook compares the staged results with that checksum. The build template has its own checksum, `TEMPLATE_SHA256`. Audit its inputs, saved results and document properties before changing that checksum.

```sh
PYTHONPATH=tools python3 - "$candidate" <<'PY'
import sys
from check_workbook_commit import risk_snapshot_digest
from xlsx_read import Workbook

with Workbook(sys.argv[1]) as workbook:
    print(risk_snapshot_digest(workbook))
PY
```
