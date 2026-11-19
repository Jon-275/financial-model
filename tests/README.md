# Workbook checks for maintainers

Use these checks when changing the planner or preparing a workbook to share. For help entering household figures, see the [first-use guide](../docs/model-guide.md#run-your-first-plan).

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

1. Open a test copy in desktop Excel for Mac. It should open without a repair prompt. Change an input and the plan end age. Check that **Summary** updates, **Annual Detail** ends at the selected age and the workbook marks the old market results out of date.
2. On **Summary**, choose **Run market risk**. Test **Cancel**, then run again and save when the status says **Risk results are current**. For changes to the market calculations, compare the saved Excel results with an independent reference that uses the same fictional assumptions and seed. [Market result comparison](native_excel_acceptance.py)
3. Switch between today’s and future dollars. Preview **Summary** on Letter and A4. Check all six pages, including the two pages of input and spending decisions. Record the Excel version, what you tested and any checks still open in the pull request. [Print instructions](../docs/model-guide.md#print-your-plan)

## Share an updated workbook

Share a fresh workbook with fictional inputs. Set `workbook` to that file. Check its inputs and file properties, then run the command below to verify the VBA code and calculate the two release checksums.

```sh
PYTHONPATH=tools python3 - "$workbook" <<'PY'
import sys
from hashlib import sha256
from pathlib import Path
from check_workbook_commit import audit_sample_workbook
from package_macro import verify_project

book = Path(sys.argv[1])
verify_project(book)
print("EXAMPLE_RISK_SHA256", audit_sample_workbook(book))
print("RELEASE_SHA256", sha256(book.read_bytes()).hexdigest())
PY
```

After reviewing the results, update the matching values in the privacy checker. Stage the intended files and run `python3 -S -B tools/check_workbook_commit.py --tracked` to verify the staged workbook and template. [Workbook privacy check](../tools/check_workbook_commit.py) · [VBA verification](../tools/package_macro.py)

Print the sample PDF from that same workbook and compare it with the saved Summary. Include a copy of the license when sharing the workbook. [Apache License 2.0](../LICENSE)

```sh
python3 -m pip install pypdf==6.21.4 pdfplumber==0.11.9
python3 tests/example_pdf_acceptance.py "$workbook" docs/examples/retirement-plan-example.pdf
```
