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

## Use dev-loop during development

Install the pinned dev-loop release with `uv`. It uses Python 3.12 or later.

```sh
uv tool install --python 3.12 'git+https://github.com/Jon-275/dev-loop@1e8fc4c86fab7bc87fe6fb2575276deee8cd342f'
devloop doctor --check
devloop changelog check
```

1. Work on one change per branch. For changes to financial calculations or saved market results, describe the expected behavior and how to verify it in the pull request.
2. For code, workbook, or workflow changes, add a short release note in `changes/<slug>.<section>.md`. The supported sections are `breaking`, `added`, `changed`, `fixed` and `security`. Run `devloop changelog check` and the relevant workbook checks before review.
3. After committing, run `devloop size --base origin/main`. The pilot uses an 800-line budget. A workbook or PDF is binary and counts as zero lines, so describe the changes and link the validation evidence in the pull request.
4. Wait for reviewers on the current commit. Reply to findings with `accepted:` or `rejected:` and explain the fix or decision. Batch fixes into one push per review round.

The [pilot workflow](../.github/workflows/devloop.yml) reports instruction drift, PR size and changelog validation. Treat its results as advisory during the pilot. Use `devloop sync` after editing `devloop.toml` or a project skill, then commit the generated updates with the source change. Keep `.githooks` enabled for the workbook privacy check. Before preparing a release, use `devloop changelog build --dry-run` to preview the notes.

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
