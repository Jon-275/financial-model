# Test and release guide

Run commands from the repository root. Use Python 3.11 or later, Node.js, and LibreOffice Calc with `soffice` on your path. For a different LibreOffice path, add `--soffice /path/to/soffice` to the scenario test command.

```sh
python3 -m pip install -r tools/requirements.txt
```

## Run the checks locally

```sh
python3 tests/model_acceptance.py --release workbook/retirement-model.xlsm
PYTHONPATH=tools python3 -m unittest discover -s tests -p 'test_workbook_privacy.py'
PYTHONPATH=tools python3 -m unittest discover -s tests -p 'test_risk_*.py'
PYTHONPATH=tools python3 -m unittest discover -s tests -p 'test_report_structure.py'
PYTHONPATH=tools python3 -m unittest discover -s tests -p 'test_native_*.py'
PYTHONPATH=tools python3 -m unittest discover -s tests -p 'test_macro_runner_contract.py'
PYTHONPATH=tools python3 -m unittest discover -s tests -p 'test_workbook_protection.py'
python3 tests/input_validation_acceptance.py
python3 tests/phase_readiness_acceptance.py workbook/retirement-model.xlsm
python3 tests/workbook_protection_acceptance.py workbook/retirement-model.xlsm
python3 tests/native_integrity_acceptance.py
python3 tests/report_formula_acceptance.py
python3 -m unittest discover -s tests -p 'test_xlsx_fixtures.py'
node --test tests/gemini_review.test.mjs
```

The release check compares saved cash, taxes, withdrawals and account balances. It also checks the print areas, charts, tab order, worksheet visibility, row shading, solver filters and workbook metadata. The risk tests check the simulator against the workbook. Report tests cover page order, account-access decisions, merged cells and the links to funding details. They recalculate 28 amounts in both dollar views and eight statements about Social Security and survivor timing. The Start Here checks verify its instructions and links against the published guide. Protection tests check which inputs can be edited and whether formulas remain visible. The Node tests cover when Gemini should pass or fail a pull request.

These instructions cover the macro-enabled edition in development. Run the release, scenario and capture checks against your build until the new workbook is published. Set `RISK_TEST_WORKBOOK` to check a specific candidate.

```sh
candidate=/absolute/path/to/retirement-model.xlsm
RISK_TEST_WORKBOOK="$candidate" PYTHONPATH=tools python3 -m unittest discover -s tests -p 'test_risk_*.py'
```

## Recalculate workbook scenarios

```sh
python3 tests/xlsx_acceptance.py workbook/retirement-model.xlsm \
  tests/phase2-fixtures.json --compact --ledger
python3 tests/xlsx_acceptance.py workbook/retirement-model.xlsm \
  tests/phase3-fixtures.json --compact --ledger
```

Each case runs on a temporary copy in LibreOffice. For a completed native run, the runner first checks every saved assumption record against the full value. It uses one short key for those repeated comparisons in the temporary copies. All financial values, formula checks and stale-input comparisons remain in the tests. The runner applies test inputs, clears saved formula results, and checks the recalculated values. Cases use fictional inputs. The Phase 2 file includes the earlier tests for the current workbook. The Phase 3 case checks that required spending can go unpaid in a year even when all taxes are paid.

Use `--case NAME` to run one case, `--jobs 1` if the machine runs low on memory, or `--keep DIRECTORY` to inspect the test copies. The `--compact` option trims inactive calculation rows in temporary copies. For a release, also run representative cases on a full-size copy by omitting `--compact`.

Run `python3 tests/aca_inflation_acceptance.py workbook/retirement-model.xlsm` to check Marketplace coverage when inflation changes each year. It compares the workbook and the Python calculation with expected one- and two-person poverty guidelines and checks a case near the credit cutoff.

Run `python3 tests/survivor_floor_acceptance.py workbook/retirement-model.xlsm` to check the medical reserve after one spouse dies. It verifies the survivor’s medical costs with and without Social Security income.

The phase-readiness test checks which pay and spending fields a plan needs. It also checks when the longer market tests require a retirement budget beyond the selected end age.

Independently recalculate sampled workbook trials and compare the saved results. The command below covers a fixed path, its flexible-spending partner and the combined stress. For release testing, run every code on a three-pair example with `--codes all`. Test interactive refresh and cancellation in desktop Excel.

```sh
python3 tests/native_flexible_policy_acceptance.py workbook/retirement-model.xlsm
python3 tests/native_capture_acceptance.py workbook/retirement-model.xlsm --codes pairs
python3 tests/native_capture_acceptance.py workbook/retirement-model.xlsm --codes 1,1001,2205
```

When changing the Phase 2 cases, generate a candidate file in scratch and compare it before updating the tracked copy.

```sh
mkdir -p scratch
python3 tests/phase2_fixtures.py > scratch/phase2-fixtures.json
diff -u tests/phase2-fixtures.json scratch/phase2-fixtures.json
```

## Pull request checks

GitHub requires **Workbook checks**, **Greptile Review** and **Gemini Review** before merging into `main`. It also requires all review threads to be resolved. Workbook checks run on PRs and after changes merge. They include release and privacy checks, simulator and review tests, representative recalculation cases, and direct comparisons between saved market results and the visible trial formulas. [Workbook workflow](../.github/workflows/workbook-checks.yml) · [Branch rules](https://github.com/Jon-275/financial-model/rules/27295551)

Gemini posts each finding on the relevant line of a ready PR. Fix the issue or explain why the finding does not apply, then resolve the thread. Critical or high-severity findings and incomplete reviews fail the required check. Address every open thread before merging, including findings from a review that passed. Gemini checks changed source, including Excel VBA modules and test fragments, hooks, dependency pins and workflows. A change limited to workbooks, PDFs, Markdown or data fixtures sends no content to Gemini. Workbook checks cover the Excel files and their privacy requirements. Source code and any literals committed to it are sent to the reviewer. Keep personal data and credentials out of source files.

Small changes are reviewed together. On larger PRs, Gemini reviews every source hunk in batches of up to 120,000 characters. A final review compares notes across batches for problems between files. The PR receives one required status for its current commit. Missing patches, incomplete responses and failed review passes prevent that status from passing.

**Run Gemini reviewer** is the GitHub Actions job. **Gemini Review** is the result it posts on the PR commit and the check GitHub requires. A title or description edit can create a skipped job entry. Changing the PR base to `main` starts a review.

A maintainer can retry a review from **Actions → Gemini 4 Hydrogen PR Review → Run workflow**. Select `main` and enter the PR number. [Gemini workflow](../.github/workflows/gemini-review.yml)

## Where to update the workbook

- **Report layout and navigation.** The report generator defines the decision tabs and appendix index. The layout code sets the print areas, links, row shading and solver filters. [Report generation](../tools/plan_report.py) · [Report layout](../tools/report_layout.py) · [Workbook tab order](../tools/workbook_layout.py)
- **Annual figures and conversion calculations.** Report data supplies the chart and annual cash figures. The same withdrawal solver serves the base projection and conversion comparison. The comparison writes the detailed results to **Conversion Detail**, and workbook packaging points its charts to that table. [Report data](../tools/dashboard.py) · [Withdrawal solver](../tools/funding_search.py) · [Conversion comparison](../tools/targeting.py) · [Workbook packaging](../tools/package_model.py)
- **Market runs.** The VBA module chooses each trial and saves its results. The visible annual sheets calculate its cash flows and tax. `Risk_Trials` stores the market outcomes and stress cases, `Funding_Path_Data` holds three complete funding histories, and `Risk_Data` feeds the report. The Python builder independently prepares the fictional example. [Runner source](../tools/vba/RunMarketRisk.bas) · [Runner and workbook interface](../docs/macro-runner.md) · [Public build](../tools/build_model.sh)

## Microsoft Excel release check

Use the candidate workbook from the [build instructions](../README.md#build-the-next-edition). Set its path below. The commands create a copy with cleared formula results, leaving the candidate available for publication.

```sh
candidate=/absolute/path/to/retirement-model.xlsm
mkdir -p scratch
excel_check=$(mktemp -d "$PWD/scratch/excel-check.XXXXXXXX")
python3 tools/workbook_recalc.py "$candidate" "$excel_check/retirement-model.xlsm"
```

1. Open the copy in desktop Excel and check for a repair prompt. In Excel for Mac, press **Control+Option+Shift+F9** to rebuild dependencies and recalculate. Save the copy. [Microsoft Excel keyboard shortcuts](https://support.microsoft.com/en-us/accessibility/excel/keyboard-shortcuts-in-excel)
2. Verify the saved copy with the command below. On a second copy, change the end age. Check that the projection, balance chart, and shaded rows stop at that age. Refresh market risk for the new end age and check the last row on **Risk Results**.
3. Verify that the fictional example opens with the success rate, shaded range, and funding examples already visible on **Summary**. On a test copy, change an active assumption and check that the figures and chart range clear. Restore the assumption and check that they return. Run market risk and confirm the results stay in the same workbook. Check progress, cancellation and a successful run after cancellation. Verify that failed or incomplete runs remain unavailable to the report. [Risk refresh instructions](../docs/risk-analysis.md#run-the-analysis)
4. Open **Income_Targeting** and change the comparison age or conversion amount. Verify the selected-year results and the link to **Conversion Detail**. Check that the charts and detail table update with the selected year. On **Funding_Calc** and **Targeting_Calc**, filter **Calculation step** to **Selected draw**. Compare the annual withdrawal by account with **Cash_Flow** and a tested conversion with its row on **Conversion Detail**. Clear the filters before using the copy for other checks.
5. Open **Planning Metrics** from **Summary**. Check the section headings, dollar units and links back to the report and annual detail. Preview **Summary** in Excel for Mac and Windows. With **Active Sheets** selected, verify that Letter and A4 each show four results pages followed by two decision pages. Page 4 must include all seven inflation, longevity and medical stresses. Preview **Annual Detail** separately. Check both dollar views, repeating headers and the last age. With variable inflation enabled, compare representative amounts with that path’s prices and chart bands with the separately calculated nominal percentiles. Save a PDF and inspect all six pages. [Print setup](../docs/model-guide.md#print-your-plan)
6. Record the Excel versions, checks run, results and any pending platform checks. [Current release record](../docs/release-notes/workbook-layout.md)

```sh
python3 tests/model_acceptance.py "$excel_check/retirement-model.xlsm"
```

Compare an Excel-saved run with the independently prepared example, using the same assumptions and seed. If the Excel run uses fewer pairs, the check ranks that smaller group and recalculates its low, middle and high funding histories. It checks every annual funding field in all three examples and reports mismatches by trial and field without printing household amounts.

```sh
python3 tests/native_excel_acceptance.py "$candidate" "$excel_check/retirement-model.xlsm"
```

For this fictional example and seed 2026, compare a fresh desktop run with the saved example using `--expect-reused 46` for 100 pairs or `--expect-reused 515` for 1,000 pairs. This also catches a change that preserves the financial results but quietly recalculates every flexible path. Use fresh, separate copies when timing two versions. Record the path count, Excel version, elapsed run time and the completion message on Risk_Settings.

Before publishing, confirm that the candidate came from the synthetic build template, contains the exact published VBA source, and holds 1,000 fixed paths, 1,000 paired flexible paths and all 33 stresses for the fictional plan. The content audit checks household inputs, the fixed text on Start Here, investment profiles, stored text and file metadata. These checks run before the file checksum is approved. The template also requires empty formula and chart caches.

After reviewing the candidate, use the command below to update `EXAMPLE_RISK_SHA256` and `RELEASE_SHA256` in `tools/check_workbook_commit.py`. The trial checksum covers the saved report, the individual trials and the three full funding histories. The full-file checksum includes notes, charts, VBA and every other part of the file. The build template has its own checksum, `TEMPLATE_SHA256`. Audit each artifact before updating its checksum.

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
from check_workbook_commit import audit_sample_workbook
from package_macro import verify_project

verify_project(sys.argv[1])
print(audit_sample_workbook(sys.argv[1]))
print(sha256(Path(sys.argv[1]).read_bytes()).hexdigest())
PY
```

## Check the fictional sample PDF

Use the workbook that produced the sample. The check compares the six printed pages with the saved Summary, checks the portfolio chart against its source values and inspects the PDF for personal information and active content. It runs without Excel. The chart must retain its vector lines and shaded band for the comparison.

```sh
python3 -m pip install pypdf==6.21.4 pdfplumber==0.11.9
python3 tests/example_pdf_acceptance.py workbook/retirement-model.xlsm docs/examples/retirement-plan-example.pdf
```

To preview a new edition in Excel, select **Summary**, confirm the fictional scenario and completed market run, and print the active sheet on US Letter paper. Confirm six pages in print preview. The public PDF includes a dated fictional-example footer, six bookmarks and a text description of the chart. Review the pages at print size before sharing.
