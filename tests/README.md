# Workbook checks for maintainers

Use this guide to build, test and release the workbook. For household inputs, follow the [first-use guide](../docs/getting-started.md).

## Set up a working copy

Run commands from the repository root with **Python 3.11 or later**. Install the dependencies and enable the privacy hook once per clone.

Use your GitHub noreply address for commits and merges. Keep personal workbooks outside the repo or in its ignored `private/` folder.

```sh
python3 -m pip install -r tools/requirements.txt
git config --local core.hooksPath .githooks
```

## Build the workbook

Rebuild after changing the source or template. The builder uses the fictional `workbook/build-template.xlsx` and LibreOffice Calc. Put `soffice` on your path or set `SOFFICE_BIN` to its executable.

```sh
mkdir -p scratch
build_dir=$(mktemp -d scratch/build.XXXXXXXX)
bash tools/build_model.sh workbook/build-template.xlsx "$build_dir/retirement-model.xlsm"
```

## Check a change

Set `workbook` to the tracked file or the new build you want to check.

```sh
workbook=workbook/retirement-model.xlsm
python3 tests/model_acceptance.py --release "$workbook"
RISK_TEST_WORKBOOK="$workbook" PYTHONPATH=tools python3 -m unittest discover -s tests -p 'test_*.py'
```

GitHub recalculates sample plans in LibreOffice and compares them with expected results. A failure lists the command and case. Follow the separate annual review for tax changes. [Workbook checks workflow](../.github/workflows/workbook-checks.yml) · [Annual tax review](../docs/annual-tax-review.md)

Open a pull request into `main`. **Workbook checks**, **Greptile Review** and **Gemini Review** must pass on the latest commit. Address every open review thread before merging. [Branch rules](https://github.com/Jon-275/financial-model/rules/27295551)

## Verify a changed workbook in Excel

1. Open a test copy in **desktop Excel for Mac** without a repair prompt. Go from **Start Here** to **Inputs**. Change an input and the end age. Check that **Summary** updates, **Annual Detail** ends at that age and old market results become out of date.
2. Review **Asset_Allocation** and **Risk_Settings**, then choose **Run market risk** on **Summary**. Test **Cancel**, rerun and save when the status says **Risk results are current**. If market calculations changed, compare saved results with an independent reference using the same fictional inputs and seed. [Market result comparison](native_excel_acceptance.py)
3. Switch dollar views on **Summary**. Preview all six pages on Letter and A4, including both decision records. Record the Excel version, checks performed and checks still open in the pull request. [Print instructions](../docs/model-guide.md#print-your-plan)

## Share an updated workbook

Use a fresh workbook with fictional inputs. Set `workbook` to that file and inspect its inputs and file properties. Verify the VBA code and calculate both release checksums:

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

Review the results, then update the matching values in the privacy checker. Stage the files and run `python3 -S -B tools/check_workbook_commit.py --tracked` to verify the staged workbook and template. [Workbook privacy check](../tools/check_workbook_commit.py) · [VBA verification](../tools/package_macro.py)

Print the sample PDF from the same workbook and compare it with the saved Summary. Include the license with shared workbooks. [Apache License 2.0](../LICENSE)

```sh
python3 -m pip install pypdf==6.21.4 pdfplumber==0.11.9
python3 tests/example_pdf_acceptance.py "$workbook" docs/examples/retirement-plan-example.pdf
```
