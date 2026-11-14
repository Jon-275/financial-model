# Portfolio risk analysis

Use this report to see how market returns affect the ability to fund spending and taxes. It reads the active scenario and the **Risk_Settings** tab, then recalculates cash, taxes and account access for each path. The report reflects the workbook settings at the time it was generated. It expects the Tax_Tables layout produced by the build and stops if that layout has changed.

## Run the report

Use Python 3.11 or later with the packages in `tools/requirements.txt` and LibreOffice’s `soffice` on your path. From the repository root:

```sh
python3 tools/run_risk.py workbook/retirement-model.xlsx \
  --trials 1000 --seed 2026 --output private/portfolio-risk.json
```

The command clears saved formula results in a temporary copy before recalculating. It checks the annual results against a path with no market shocks before writing a report. Use a new output filename for each run. Use `--cached` only with a workbook that has already been recalculated and saved. The same comparison still runs.

The JSON includes annual balance ranges, first shortfall ages, early retirement stress results and the workbook checksum. To create the Excel report, run `tools/render_risk.mjs` with Node.js and the bundled `@oai/artifact-tool` package available to the script. If the package is provided in a separate runtime, copy the script to a temporary directory next to a link to the runtime’s `node_modules`.

```sh
node tools/render_risk.mjs private/portfolio-risk.json private/portfolio-risk.xlsx
```

The report is a snapshot. Rerun it after changing any workbook inputs. Keep the workbook and reports on approved local storage.

## How the analysis works

- **Returns.** The engine uses the active working and retirement portfolios from **Asset_Allocation**. It draws one correlated set of asset returns per year and applies each account’s mix to those returns. The model treats years as independent. If you entered an investment fee, the workbook and market paths both apply it to account growth. The workbook’s return basis controls whether the starting assumption anchors compound growth or the average annual return. The HSA follows the Roth return in the deterministic workbook, while its own asset mix determines its exposure to market shocks. `--volatility 0` reproduces the baseline.
- **Annual funding.** Each path uses its own account balances and brokerage basis. It recalculates taxes, eligible withdrawals, conversions, healthcare credits and income-based Medicare surcharges. Then it pays tax and spending, funds any optional savings, and applies the year’s returns. A path has a shortfall if at least one cent of required spending or taxes is unpaid in any year. A missed savings target is tracked separately from required obligations.
- **Flexible spending.** Both policies use identical market returns. The report cuts spending when the opening portfolio falls below the share of the base plan set on **Risk_Settings**. Only flexible recurring retirement costs are eligible for cuts. Healthcare, taxes, one-time expenses and fixed mortgage payments keep their planned amounts.
- **Early losses.** The report applies each stock return on **Risk_Settings** to the first retirement years. Other assets and later years use the plan’s investment assumptions.
- **Reading the results.** Balances and shortfalls are in opening-year dollars. Balance ranges include all paths, including paths with a shortfall. The 95% simulation interval measures sampling variation within the chosen return assumptions. Wages, inflation, longevity and tax assumptions follow the workbook. The same seed, inputs and engine version reproduce the paths.

The tax and coverage logic uses the workbook’s planning estimates. Phase 2 documents the federal and California tax assumptions, Medicare lookback and Marketplace eligibility limits. See the [Phase 2 release record](release-notes/phase-2.md), [IRS Marketplace credit guidance](https://www.irs.gov/affordable-care-act/individuals-and-families/questions-and-answers-on-the-premium-tax-credit), and [SSA Medicare surcharge guidance](https://www.ssa.gov/OP_Home/handbook/handbook.25/handbook-2504.html).
