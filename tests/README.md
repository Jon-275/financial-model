# Workbook checks for maintainers

Use this guide when changing the planner or preparing a new workbook to share. To enter household figures and run a plan, start with the [workbook user guide](../docs/model-guide.md#run-your-first-plan).

## Set up a working copy

Run commands from the repository root with Python 3.11 or later. Install the project dependencies and enable the workbook privacy check once per clone. Use your GitHub noreply address for commits and merges. Keep household workbooks outside the repository or in the ignored `private/` folder.

```sh
python3 -m pip install -r tools/requirements.txt
git config --local core.hooksPath .githooks
```

## Build the workbook

Rebuild after changing the workbook source or template. The builder uses the fictional `workbook/build-template.xlsx` and LibreOffice Calc to create a new macro-enabled planner. Put `soffice` on your path or set `SOFFICE_BIN` to its executable.

```sh
mkdir -p scratch
build_dir=$(mktemp -d scratch/build.XXXXXXXX)
bash tools/build_model.sh workbook/build-template.xlsx "$build_dir/retirement-model.xlsm"
```

## Check a change

The commands below check the workbook in the repository. To check a new build, set `workbook` to its path instead.

```sh
workbook=workbook/retirement-model.xlsm
python3 tests/model_acceptance.py --release "$workbook"
RISK_TEST_WORKBOOK="$workbook" PYTHONPATH=tools python3 -m unittest discover -s tests -p 'test_*.py'
```

GitHub also uses LibreOffice Calc to recalculate selected scenarios and compare them with expected results. If a check fails, the workflow lists the exact command and case. Tax changes have a separate annual review. [Workbook checks workflow](../.github/workflows/workbook-checks.yml) · [Annual tax review](../docs/annual-tax-review.md)

Open a pull request into `main`. **Workbook checks**, **Greptile Review** and **Gemini Review** must pass on the latest commit. Address every open review thread before merging. [Branch rules](https://github.com/Jon-275/financial-model/rules/27295551)

## Verify a changed workbook in Excel

1. Open a disposable copy in desktop Excel for Mac. Confirm it opens without a repair prompt. Change an input and the plan end age. Check that **Summary** updates, **Annual Detail** stops at the selected age and the old market results are marked out of date.
2. On **Summary**, choose **Run market risk**. Test **Cancel**, then run again and save when the status says **Risk results are current**. For changes to the market calculations, compare the saved Excel results with an independent reference that uses the same fictional assumptions and seed. [Market result comparison](native_excel_acceptance.py)
3. Switch between today’s and future dollars. Preview **Summary** on Letter and A4. Check all six pages, including the two pages of input and spending decisions. Record the Excel version, what you tested and any checks still open in the pull request. [Print instructions](../docs/model-guide.md#print-your-plan)

## Share an updated workbook

Use the unedited fictional workbook for distribution. Check its inputs, saved results, embedded VBA and file properties before updating the approved checksums. Stage the intended files, then run `python3 -S -B tools/check_workbook_commit.py --tracked` to verify the staged workbook and template. [Workbook privacy check](../tools/check_workbook_commit.py) · [VBA verification](../tools/package_macro.py)

Print the sample PDF from that same workbook and compare it with the saved Summary. Include a copy of the license when sharing the workbook. [Apache License 2.0](../LICENSE)

```sh
python3 -m pip install pypdf==6.21.4 pdfplumber==0.11.9
python3 tests/example_pdf_acceptance.py "$workbook" docs/examples/retirement-plan-example.pdf
```
