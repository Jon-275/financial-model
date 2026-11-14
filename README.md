# Retirement model

The Excel workbook projects household cash, taxes and account access through the end age you choose, up to age 125. Open **Start Here**, choose a scenario on **Inputs**, then read the funding answer at the top of **Summary**. [Open the workbook](workbook/retirement-model.xlsx)

**Summary** charts the portfolio balance and shows how the plan pays its bills around retirement. Use **Cash_Flow** to see what went unpaid and **Roth_Access** to see which savings were available.

The separate risk report recalculates spending, taxes and withdrawals under different market returns. Adjust its inputs on **Risk_Settings** and rerun the report after changing the workbook. [Portfolio risk analysis](docs/risk-analysis.md)

- [Model guide](docs/model-guide.md): how to enter assumptions and interpret results.
- [Implementation plan](PLAN.md): phase scope, release gates and progress.
- [Release records](docs/release-notes/phase-3.md) for Phase 3, [Phase 2](docs/release-notes/phase-2.md) and [Phase 1](docs/release-notes/phase-1.md).
- [Portfolio management](docs/portfolio-management.md): tax harvesting guidance separate from the forecast.

The workbook uses fictional inputs. Save a private copy before entering your own figures. All 15 worksheet tabs are visible, including **Funding_Calc** for withdrawals, **Targeting_Calc** for Roth conversions and **Report_Data** for the Summary charts.

## Rebuild the workbook

The [build template](workbook/build-template.xlsx) provides the starting layout and legacy formulas. Use the [finished workbook](workbook/retirement-model.xlsx) to run a plan.

Install Python 3.11 or later, Node.js and LibreOffice. The build also requires the Codex spreadsheet runtime's `@oai/artifact-tool` package. Point `ARTIFACT_NODE_MODULES` to the runtime's `node_modules` directory and `NODE_BIN` to its Node.js executable if needed. The finished workbook opens in desktop Excel independently of the build tools.

```sh
python3 -m pip install -r tools/requirements.txt
mkdir -p scratch
ARTIFACT_NODE_MODULES=/path/to/runtime/node_modules \
  bash tools/build_model.sh workbook/build-template.xlsx scratch/retirement-model.xlsx
python3 tests/model_acceptance.py --release scratch/retirement-model.xlsx
```

The build needs a new output path. Keep the file name `retirement-model.xlsx` so the Excel chart references resolve. See the [test guide](tests/README.md) for scenario checks and Excel verification.
