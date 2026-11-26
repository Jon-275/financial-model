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

Install the pinned dev-loop revision with `uv`. It uses Python 3.12 or later.

```sh
uv tool install --python 3.12 'git+https://github.com/Jon-275/dev-loop@017bc7c1e3bee10ee79461345dd20d30f9495f30'
devloop doctor --check
devloop changelog check
```

1. Work on one change per branch. For changes to financial calculations or saved market results, describe the expected behavior and how to verify it in the pull request.
2. For code, workbook, or workflow changes, add a short release note in `changes/<slug>.<section>.md`. The supported sections are `breaking`, `added`, `changed`, `fixed` and `security`. Run `devloop changelog check` and the relevant workbook checks before review.
3. After committing, run `devloop size --base origin/main`, or use the predecessor ref for a stacked pull request. The pilot uses an 800-line budget. A workbook or PDF is binary and counts as zero lines, so describe the changes and link the validation evidence in the pull request.
4. Wait for reviewers on the current commit. Reply to findings with `accepted:` or `rejected:` and explain the fix or decision. Batch fixes into one push per review round.

The [pilot workflow](../.github/workflows/devloop.yml) reports instruction drift, PR size and changelog validation. Treat its results as advisory during the pilot. Use `devloop sync` after editing `devloop.toml` or a project skill, then commit the generated updates with the source change. Keep `.githooks` enabled for the workbook privacy check. Before preparing a release, use `devloop changelog build --dry-run` to preview the notes.

Use the normal repository identity for commits. Keep coding-tool attribution out of commit metadata, messages and pull request descriptions. The repository setting disables automatic commit trailers.

## Verify a changed workbook in Excel

1. Open a test copy in **desktop Excel for Mac** without a repair prompt. Go from **Start Here** to **Inputs**. Change an input and the end age. Check that **Summary** updates, **Annual Detail** ends at that age and old market results become out of date.
2. Review **Asset_Allocation** and **Risk_Settings**, then choose **Run market risk** on **Summary**. Test **Cancel**, rerun and save when the status says **Risk results are current**. If market calculations changed, compare saved results with an independent reference using the same fictional inputs and seed. [Market result comparison](native_excel_acceptance.py)
3. Switch dollar views on **Summary**. Preview all six pages on Letter and A4, including both decision records. Record the Excel version, checks performed and checks still open in the pull request. [Print instructions](../docs/model-guide.md#print-your-plan)

### Verify the California v2 report candidate

Use a disposable explicit-policy workbook with a completed, independently reconciled run. Check the opening year, `CA`, `ca-individual-v1` and published California income and disability insurance table years on Summary, Annual Detail and Risk Results. Switch to baseline tables and confirm the caption shows `Baseline tax tables` and the previous run’s caption disappears from Risk Results. Rerun and verify the baseline caption there. Use a separate invalid fixture to check that changing the jurisdiction or policy version clears the valid scope and prompts a review of the inputs.

On each supported desktop platform, preview Summary in both dollar views on Letter and A4. Confirm the six pages and existing breaks, that the complete scope is readable on page one and that the separately printed Annual Detail shows its scope without clipping. Compare each export’s printed values and chart with the corresponding saved workbook. Distribution also requires the sample PDF acceptance command below and the matching fictional workbook and release checks.

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
