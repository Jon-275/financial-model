# Test and release guide

Households can enter a plan, run market risk and print the report in desktop Excel. Maintainers use Python 3.11 or later, Node.js and LibreOffice Calc to build and test the workbook. The desktop release check uses Excel for Mac. Run commands from the repository root with `soffice` on your path. [User guide](../docs/model-guide.md) · [Build instructions](../README.md)

## Protect personal files

Keep household workbooks outside the checkout or under the ignored `private/` folder. Contribute from a clone of the current repository. If an older checkout contains private files or commits, save your work privately, start with a fresh clone and copy across only the source changes you have reviewed.

In GitHub **Settings → Emails**, copy your noreply address. Set it locally with `git config --local user.email "ADDRESS"`, using the address you copied. Enable the privacy hook, check the configured address and confirm that both tracked workbooks pass the audit.

```sh
git config --local core.hooksPath .githooks
git config --local user.email
python3 -S -B tools/check_workbook_commit.py --tracked
```

## Run the checks locally

Set `candidate` to the fictional workbook you want to test. The tracked file is the default below. The privacy check above always reads the Git index.

```sh
candidate=workbook/retirement-model.xlsm
python3 -m pip install -r tools/requirements.txt
python3 tests/model_acceptance.py --release "$candidate"
RISK_TEST_WORKBOOK="$candidate" PYTHONPATH=tools python3 -m unittest discover -s tests -p 'test_*.py'
node --test tests/gemini_review.test.mjs
python3 tests/workbook_protection_acceptance.py "$candidate"
python3 tests/native_order_acceptance.py "$candidate"
```

These checks cover cash, tax, withdrawals, account balances, printed pages, charts, input protection and the calculation order used to recover a market trial. GitHub also checks the public VBA against its embedded copy and runs the standalone test scripts without a custom Python import path. [Workbook checks workflow](../.github/workflows/workbook-checks.yml)

## Recalculate workbook scenarios

Run the following in LibreOffice Calc. The scenario tests make temporary copies, enter fictional inputs, clear cached results and recalculate the formulas.

```sh
python3 tests/input_validation_acceptance.py
python3 tests/phase_readiness_acceptance.py "$candidate"
python3 tests/aca_inflation_acceptance.py "$candidate"
python3 tests/survivor_floor_acceptance.py "$candidate"
python3 tests/one_time_expense_acceptance.py "$candidate"
python3 tests/xlsx_acceptance.py "$candidate" tests/phase2-fixtures.json --compact --ledger
python3 tests/xlsx_acceptance.py "$candidate" tests/phase3-fixtures.json --compact --ledger
python3 tests/report_formula_acceptance.py
python3 tests/native_integrity_acceptance.py
python3 tests/native_closing_acceptance.py "$candidate"
python3 tests/native_flexible_policy_acceptance.py "$candidate"
python3 tests/native_capture_acceptance.py "$candidate" --codes pairs
python3 tests/native_capture_acceptance.py "$candidate" --codes 1,1001,2205
```

The financial cases cover spending, payroll, account access, Marketplace credits under variable inflation, survivor costs and scheduled bills. The market checks compare saved outcomes with worksheet trials and test the annual account balances, spending cuts and handling of stale results. On a three-pair test workbook with a completed run, use `--codes all` to compare every trial.

Use `xlsx_acceptance.py --case NAME` to select a case, `--jobs 1` to limit memory use, `--keep DIRECTORY` to inspect the test copies, or `--soffice /path/to/soffice` to select Calc. The compact tests trim inactive rows in those copies. For a release, also run representative cases without `--compact` to check the full workbook.

If you change the financial fixtures, compare the generated cases before updating the tracked file.

```sh
mkdir -p scratch
python3 tests/phase2_fixtures.py > scratch/phase2-fixtures.json
diff -u tests/phase2-fixtures.json scratch/phase2-fixtures.json
```

## Pull request checks

Before merging to `main`, **Workbook checks**, **Greptile Review** and **Gemini Review** must pass on the latest commit. Resolve every review thread, including findings from a passing review. Fix the issue or explain why the finding does not apply. The workbook checks also run after a merge. [Branch rules](https://github.com/Jon-275/financial-model/rules/27295551)

**Run Gemini reviewer** is the Actions job that posts the required **Gemini Review** status. A title or description edit can leave a skipped job entry. Check the status on the current commit. Gemini reviews changed source files, which can include VBA, hooks and workflows. Documentation, workbook, PDF and data-only changes send no content to Gemini. Keep personal data and credentials out of source. If a review needs a retry, open **Actions → Gemini 4 Hydrogen PR Review → Run workflow**, select `main` and enter the PR number. [Gemini workflow](../.github/workflows/gemini-review.yml)

## Where to update the workbook

- **Report and navigation.** Edit the page content, print areas, sheet order and links in the report generators. [Report pages](../tools/plan_report.py) · [Layout](../tools/report_layout.py) · [Tab order](../tools/workbook_layout.py)
- **Annual balances and Roth conversions.** The projection and the conversion comparison share a withdrawal solver. The comparison writes its detailed results to **Conversion Detail**. [Report data](../tools/dashboard.py) · [Withdrawal solver](../tools/funding_search.py) · [Conversion comparison](../tools/targeting.py) · [Workbook packaging](../tools/package_model.py)
- **Market risk.** The VBA runner selects the trials. The visible worksheets calculate them. **Risk_Trials** stores the outcomes, **Funding_Path_Data** stores three funding histories and **Risk_Data** feeds the report. The builder independently calculates the fictional example. [VBA runner](../tools/vba/RunMarketRisk.bas) · [Runner guide](../docs/macro-runner.md) · [Build script](../tools/build_model.sh)

## Microsoft Excel release check

Use disposable copies in desktop Excel for Mac. Keep the unedited fictional candidate for publication. Create a copy with cleared formula results for the first two steps.

```sh
mkdir -p scratch
excel_check=$(mktemp -d "$PWD/scratch/excel-check.XXXXXXXX")
python3 tools/workbook_recalc.py "$candidate" "$excel_check/retirement-model.xlsm"
```

1. Open the copy and confirm it opens without a repair prompt. Press **Control+Option+Shift+F9** to rebuild dependencies and recalculate. Save the copy. [Excel keyboard shortcuts](https://support.microsoft.com/en-us/accessibility/excel/keyboard-shortcuts-in-excel)
2. Run `python3 tests/model_acceptance.py "$excel_check/retirement-model.xlsm"`. In a separate copy, change the end age. Confirm that the projection, balance chart and shading stop at the selected age. Run market risk and check the last year on **Risk Results**.
3. On a fresh copy, confirm **Summary** opens with the success rate, balance range and funding examples. Change an active assumption and check that stale figures clear. Restore the assumption and check that they return. Use **Run market risk**, wait for the results to be current and save. Check progress, cancellation, a successful retry and the report status after an incomplete run. Record whether cancellation was tested with the actual button or through its procedure. [Risk instructions](../docs/risk-analysis.md#run-the-analysis)
4. On **Income_Targeting**, change the age or conversion amount. Trace the updated chart to **Conversion Detail**. On **Funding_Calc** and **Targeting_Calc**, filter **Calculation step** to **Selected draw**. Check the withdrawal against **Cash_Flow** and the conversion against its detail row, then clear the filters.
5. Open **Planning Metrics** from **Summary** and follow its links to the report and annual detail. Switch between today’s and future dollars. Check **Annual Detail** through the last age, the funding examples against their own price paths and the balance band against the nominal percentiles in the future-dollar view.
6. Print **Summary** with **Active Sheets** on Letter and A4. Verify four results pages followed by two decision pages. Page four should show all seven inflation, longevity and medical stresses. Check legends, units and page breaks on both sizes. Preview **Annual Detail** separately and check that headings repeat. Record the candidate revision, Excel version, results and any outstanding checks in the release PR. [Print instructions](../docs/model-guide.md#print-your-plan)

### Compare the saved market results

Run the fictional example in Excel with seed 2026. Save a copy after each run. Compare it with the independent reference using the same financial assumptions. The check compares each outcome and all three ranked funding histories. For a smaller run, it finds the reference histories for that run’s own low, median and high paths.

```sh
python3 tests/native_excel_acceptance.py "$candidate" scratch/example-100-mac.xlsm --expect-reused 46
python3 tests/native_excel_acceptance.py "$candidate" scratch/example-1000-mac.xlsm --expect-reused 515
```

The expected counts check that Excel reused the fixed result when the flexible policy made no change. Record the path count, Excel version, elapsed time and completion message. Use a fresh copy for each timing.

To check cancellation or recovery, select a trial on **Risk_Trials** in a workbook with a completed run and save a baseline copy. Run the exercise on a second copy and save it. The check compares the selected trial, active figures, run date and all previous results. It also checks that the withdrawal formulas remain and that unused years are blank.

```sh
python3 tests/native_excel_acceptance.py scratch/before-recovery.xlsm scratch/after-recovery.xlsm --check-restoration
```

For a test that records its verdict on a worksheet, add `--expect-pass 'Recovery!A2'` using the actual sheet and cell.

### Reproduce the age-125 comparison

Use the fictional example, seed 2026 and 1,000 paired paths. In the Excel test copy, set both starting ages to 0 and the end age to 125. Set the four Roth opening years listed below to 2026. This fills all 126 annual positions, including the opening year. Run market risk and save the Excel copy.

Build a separate reference from the repository template. Run these commands from the repository root with the usual build dependencies and `soffice` available.

```sh
mkdir -p scratch
capacity=$(mktemp -d "$PWD/scratch/native-capacity.XXXXXXXX")
python3 tools/model_plan.py workbook/build-template.xlsx "$capacity/plan.json"
python3 tools/author_model.py "$capacity/plan.json" "$capacity/authored.xlsx"
python3 tools/package_model.py workbook/build-template.xlsx "$capacity/plan.json" "$capacity/authored.xlsx" "$capacity/model.xlsx"
PYTHONPATH=tools python3 - "$capacity" <<'PYTHON'
from pathlib import Path
import sys
from xlsx_values import write_values

folder = Path(sys.argv[1])
inputs = {"D8": 0, "D9": 0, "D53": 125, "D29": 2026, "D92": 2026, "D93": 2026, "D94": 2026}
write_values(folder / "model.xlsx", folder / "capacity.xlsx", {"Inputs": inputs})
PYTHON
python3 tools/native_bootstrap.py "$capacity/capacity.xlsx" "$capacity/reference.xlsx"
python3 tests/native_excel_acceptance.py "$capacity/reference.xlsx" /path/to/age-125-mac-run.xlsm --expect-reused 501
```

The comparison covers all 2,033 outcomes, the three path rankings, their annual balances and their complete funding histories. Record the Excel runtime separately from the reference build time.

### Prepare the release file

Audit the synthetic inputs, investment profiles, metadata and stored results before updating any approved checksum. The workbook must contain the reviewed VBA source, 1,000 fixed paths, their 1,000 flexible partners, 33 stress cases and three complete funding histories. The synthetic build template must have empty formula and chart caches.

If the template changes, regenerate it in scratch. Use new output paths so the previous template stays available for comparison.

```sh
python3 tools/template_fixture.py workbook/build-template.xlsx scratch/template-plan.json
python3 tools/author_model.py scratch/template-plan.json scratch/template-cells.xlsx
python3 tools/template_fixture.py workbook/build-template.xlsx scratch/clean-template.xlsx \
  --authored scratch/template-cells.xlsx
```

Verify the VBA and example content, then use the reported hashes for the saved results and full workbook in `tools/check_workbook_commit.py`. If the template changes, update its separately audited `TEMPLATE_SHA256`. Stage the intended fictional files and run `python3 -S -B tools/check_workbook_commit.py --tracked` to check the approved contents in the Git index.

```sh
PYTHONPATH=tools python3 - "$candidate" <<'PY'
import sys
from hashlib import sha256
from pathlib import Path
from check_workbook_commit import audit_sample_workbook
from package_macro import verify_project

verify_project(sys.argv[1])
print('EXAMPLE_RISK_SHA256', audit_sample_workbook(sys.argv[1]))
print('RELEASE_SHA256', sha256(Path(sys.argv[1]).read_bytes()).hexdigest())
PY
```

To restore a release, open a PR that restores the verified workbook, matching source, fixtures, checksum values and sample PDF together. Repeat the release and privacy checks against that revision.

## Check the fictional sample PDF

Print **Summary** from the exact fictional candidate with current market results. Use US Letter for the shareable PDF and check all six pages at print size. Keep the dated fictional-example footer, six bookmarks, chart description, vector lines and shaded band. The checker compares the pages and chart with the saved workbook and checks for personal information and active content.

```sh
python3 -m pip install pypdf==6.21.4 pdfplumber==0.11.9
python3 tests/example_pdf_acceptance.py "$candidate" docs/examples/retirement-plan-example.pdf
```
