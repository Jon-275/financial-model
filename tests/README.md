# Workbook acceptance tests

Run from this directory using Python 3.11 or newer. The scripts read `.xlsx` packages with Python's standard library. The scenario runner also needs LibreOffice's `soffice` command, or its path supplied with `--soffice`.

```sh
python3 model_acceptance.py --release /path/to/retirement-model.xlsx
python3 phase2_fixtures.py > phase2-fixtures.json
python3 xlsx_acceptance.py /path/to/retirement-model.xlsx phase2-fixtures.json --ledger
```

`model_acceptance.py` independently reconciles saved cash, spending, tax payables, Roth balances by owner, and portfolio cash flow. It checks the Summary's first shortfall, the selected end age and the four retained chart objects. It reads saved values. The release flag also checks that every sheet is visible and that inherited document metadata is absent. Use the Excel release check below to verify recalculation and chart rendering.

`xlsx_acceptance.py` makes temporary copies, applies the fixture inputs, clears stale formula caches, recalculates each copy in a separate LibreOffice profile, and verifies both the explicit expected values and the annual reconciliations. It never overwrites the source. Use `--case NAME` to run selected cases, `--jobs 1` to reduce memory use, or `--keep DIRECTORY` to keep recalculated copies for investigation. Fixture settings use a synthetic 20% tax and zero investment returns to make the expected results calculable by hand. These assumptions are for testing only.

For quicker development feedback, use `--compact` to omit inactive solver rows in disposable copies. Cases focused on conversion comparison still include the full selected-year sweep. Run an unmodified workbook and representative full-size cases before release. The combined fixture file contains the Phase 1 regressions and independent checks for Marketplace thresholds, Medicare lookbacks, HSA eligibility, tax tables, capital losses and one-year conversions.

Use the combined fixtures with the current workbook. The historical `phase1-fixtures.json` still covers the old risk worksheets in the Phase 1 release.

## Portfolio risk

From the repository root, install `tools/requirements.txt` and run:

```sh
PYTHONPATH=tools python3 -m unittest discover -s tests -p 'test_risk_*.py'
```

The risk tests check return paths, account access and funding behavior, saved formula validation, and annual parity with the release workbook. The synthetic SEPP case verifies that spending can remain unpaid even when the year’s taxes are fully paid:

```sh
python3 tests/xlsx_acceptance.py workbook/retirement-model.xlsx \
  tests/phase3-fixtures.json --compact --ledger
```

## Pull request checks

GitHub runs **Workbook checks** on each pull request into `main` and after changes land on `main`. It checks tracked workbooks for private inputs and metadata, validates the saved release workbook, runs the risk tests, and recalculates six scenarios in LibreOffice. See the [workflow](../.github/workflows/workbook-checks.yml) for the commands and the scenario list.

**Gemini Review** checks source and workflow changes on ready pull requests, including forks. It blocks merging for critical or high-severity findings or if the review cannot finish. Suggestions appear in the PR comment. PRs without source or workflow changes pass the Gemini check, and **Workbook checks** still validates the workbook.

Gemini reads GitHub's text patches and uses the `GEMINI_API_KEY` repository secret. To retry a review, use **Actions → Gemini 4 Hydrogen PR Review → Run workflow**, select `main`, and enter the PR number. Run the reviewer tests locally with `node --test tests/gemini_review.test.mjs` from the repository root.

## Microsoft Excel release check

Create a test directory and use `tools/workbook_recalc.py` to copy the release without its saved formula values. Keep the filename `retirement-model.xlsx` so the workbook chart references resolve. Open the copy in desktop Excel and check that it opens without a repair prompt. In Excel for Mac, use Control+Option+Shift+F9 to rebuild dependencies and recalculate. Save the copy and run `model_acceptance.py` against it. [Microsoft Excel keyboard shortcuts](https://support.microsoft.com/en-us/accessibility/excel/keyboard-shortcuts-in-excel)

On a separate test copy, change the active scenario's end age and confirm the Projection and portfolio chart both reach the new end year. Record the Excel version and results in the release notes.
