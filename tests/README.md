# Test and release guide

Run commands from the repository root. Use Python 3.11 or later, Node.js, and LibreOffice Calc with `soffice` on your path. For a different LibreOffice path, add `--soffice /path/to/soffice` to the scenario test command.

```sh
python3 -m pip install -r tools/requirements.txt
```

## Run the checks locally

```sh
python3 tests/model_acceptance.py --release workbook/retirement-model.xlsx
PYTHONPATH=tools python3 -m unittest discover -s tests -p 'test_workbook_privacy.py'
PYTHONPATH=tools python3 -m unittest discover -s tests -p 'test_risk_*.py'
node --test tests/gemini_review.test.mjs
```

The release check compares saved cash, taxes, withdrawals and account balances. It also checks the print areas, charts, tab order, worksheet visibility, row shading, solver filters and workbook metadata. The risk tests check the simulator against the workbook. The Node tests cover when Gemini should pass or fail a pull request.

For a newly built workbook, use its path in the release, scenario, and risk refresh commands. Set `RISK_TEST_WORKBOOK` to run the risk engine’s scenario and annual calculation tests against the candidate.

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

## Where to update the workbook

- **Report layout and navigation.** The report generator defines the decision tabs and appendix index. The layout code sets the print areas, links, row shading and solver filters. [Report generation](../tools/plan_report.py) · [Report layout](../tools/report_layout.py) · [Workbook tab order](../tools/workbook_layout.py)
- **Annual figures and conversion calculations.** Report data supplies the chart and annual cash figures. The same withdrawal solver serves the base projection and conversion comparison. The comparison writes the detailed results to **Conversion Detail**, and workbook packaging points its charts to that table. [Report data](../tools/dashboard.py) · [Withdrawal solver](../tools/funding_search.py) · [Conversion comparison](../tools/targeting.py) · [Workbook packaging](../tools/package_model.py)
- **Saved risk results.** The risk runner writes the results to **Risk_Data**. The workbook layout for this release uses `retirement-risk-workbook-v4` and engine 1.4.0, including carried spending obligations and working-year Roth access. Keep the snapshot writer, display references and privacy check in sync. Use the runner from the same release as the workbook. [Snapshot writer](../tools/risk_workbook.py) · [Display references](../tools/risk_presentation.py) · [Privacy check](../tools/check_workbook_commit.py)

## Microsoft Excel release check

Use the candidate workbook from the [build instructions](../README.md#rebuild-the-workbook). Set its path below. The commands create a copy with cleared formula results, using the filename the charts expect.

```sh
candidate=/absolute/path/to/retirement-model.xlsx
mkdir -p scratch
excel_check=$(mktemp -d "$PWD/scratch/excel-check.XXXXXXXX")
python3 tools/workbook_recalc.py "$candidate" "$excel_check/retirement-model.xlsx"
```

1. Open the copy in desktop Excel and check for a repair prompt. In Excel for Mac, press **Control+Option+Shift+F9** to rebuild dependencies and recalculate. Save the copy. [Microsoft Excel keyboard shortcuts](https://support.microsoft.com/en-us/accessibility/excel/keyboard-shortcuts-in-excel)
2. Verify the saved copy with the command below. On a second copy, change the end age. Check that the projection, balance chart, and shaded rows stop at that age. Refresh market risk for the new end age and check the last row on **Risk Results**.
3. Verify that the fictional example opens with the success rate, shaded range, and funding examples already visible on **Summary**. On a test copy, change an active assumption and check that the figures and chart range clear. Restore the assumption and check that they return. Refresh the test copy and confirm the results stay in its **Summary**. [Risk refresh instructions](../docs/risk-analysis.md#refresh-the-workbook)
4. Open **Income_Targeting** and change the comparison age or conversion amount. Verify the selected-year results and the link to **Conversion Detail**. Check that the charts and detail table update with the selected year. On **Funding_Calc** and **Targeting_Calc**, filter **Calculation step** to **Selected draw**. Compare the annual withdrawal by account with **Cash_Flow** and a tested conversion with its row on **Conversion Detail**. Clear the filters before using the copy for other checks.
5. Open **Planning Metrics** from **Summary**. Check the section headings, dollar units and links back to the report and annual detail. Preview **Summary** in Excel for Mac and Windows. With **Active Sheets** selected, verify that Letter and A4 each show three results pages followed by two decision pages. Preview **Annual Detail** separately. Check the selected dollar view, repeating headers and the last age. Save a PDF and inspect the figures and page breaks. [Print setup](../docs/model-guide.md#print-your-plan)
6. Record the Excel versions, checks run, results and any pending platform checks. [Current release record](../docs/release-notes/workbook-layout.md)

```sh
python3 tests/model_acceptance.py "$excel_check/retirement-model.xlsx"
```

Before publishing, confirm that the candidate came from the synthetic build template and that **Risk_Data** contains only the seeded fictional run. The content audit checks the template and release inputs, investment profiles, stored text and metadata independently of their checksums. The template also requires empty formula and chart caches.

After reviewing the candidate, use the command below to update `EXAMPLE_RISK_SHA256` and `RELEASE_SHA256` in `tools/check_workbook_commit.py`. The full-file checksum also catches added sheets, notes and other changes outside the input and risk ranges. The build template has its own checksum, `TEMPLATE_SHA256`. Audit each artifact before updating its checksum.

To refresh the synthetic template, write fixture cells with the public author and preserve the source workbook's native parts. Use new output paths for each step.

```sh
python3 tools/template_fixture.py workbook/build-template.xlsx scratch/template-plan.json
python3 tools/author_model.py scratch/template-plan.json scratch/template-cells.xlsx
python3 tools/template_fixture.py workbook/build-template.xlsx scratch/clean-template.xlsx \
  --authored scratch/template-cells.xlsx
```

```sh
PYTHONPATH=tools python3 - "$candidate" <<'PY'
import sys
from hashlib import sha256
from pathlib import Path
from check_workbook_commit import risk_snapshot_digest
from xlsx_read import Workbook

with Workbook(sys.argv[1]) as workbook:
    print(risk_snapshot_digest(workbook))
print(sha256(Path(sys.argv[1]).read_bytes()).hexdigest())
PY
```
