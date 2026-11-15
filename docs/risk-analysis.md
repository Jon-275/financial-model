# Portfolio risk analysis

The risk report shows how often your plan pays its bills and taxes under different market returns. It uses the active scenario and the settings on **Risk_Settings**. Save the workbook after making changes and run a new report for that version.

## Run the report

Use Python 3.11 or later and LibreOffice Calc. Put `soffice` on your path or set `SOFFICE_BIN` to the executable. Run these commands from the repository root.

```sh
python3 -m pip install -r tools/requirements.txt
mkdir -p private
report_dir=$(mktemp -d "$PWD/private/risk.XXXXXXXX")
python3 tools/run_risk.py workbook/retirement-model.xlsx \
  --trials 1000 --seed 2026 --output "$report_dir/portfolio-risk.json"
```

Replace the workbook path with your private copy to use your figures. The runner recalculates a temporary copy in LibreOffice and checks that the plan’s baseline results match before running the report. Use `--cached` only after the workbook has been recalculated and saved. The baseline check still runs.

The JSON report contains balance ranges, when a shortfall first occurs, and the results of early retirement losses. Each run uses a new output folder under `private/`.

## Create the Excel report

Use the same terminal and output folder. Set `ARTIFACT_NODE_MODULES` to the Codex spreadsheet runtime’s absolute `node_modules` path. Set `NODE_BIN` if you use a separate Node executable. The [build instructions](../README.md#rebuild-the-workbook) explain these settings.

```sh
(
  set -eu
  render_dir=$(mktemp -d)
  trap 'rm -rf -- "$render_dir"' EXIT
  cp tools/render_risk.mjs "$render_dir/"
  ln -s "${ARTIFACT_NODE_MODULES:?Set the runtime node_modules path}" "$render_dir/node_modules"
  "${NODE_BIN:-node}" "$render_dir/render_risk.mjs" \
    "$report_dir/portfolio-risk.json" "$report_dir/portfolio-risk.xlsx"
)
```

Open `portfolio-risk.xlsx` in the report folder. Keep copies with personal figures in `private/` or outside the checkout.

## Read the results

- **Plans that cover their costs.** A run counts as a shortfall if at least one cent of required spending or tax goes unpaid in any year. A missed savings goal is tracked separately.
- **Fixed and flexible spending.** Both policies use the same market returns. The flexible policy cuts part of recurring retirement living costs when the opening balance falls below the share of the planned balance set on **Risk_Settings**. Each run recalculates tax, withdrawals, healthcare credits and Medicare surcharges. The spending rule leaves fixed costs such as mortgage payments in place.
- **Early losses.** The report also applies the stock returns listed on **Risk_Settings** to the first retirement years. Other assets and later years use the plan’s assumed returns.
- **Balance ranges.** Dollar amounts use the opening year’s prices. The ranges include runs with shortfalls. The 95% interval shows how much uncertainty comes from running a limited number of simulations.

The report uses each account’s investment mix and applies the fee entered on **Inputs**. The base plan uses the Roth return for the health savings account. Market simulations use the health savings account’s own investments to estimate how much its return varies.

The return setting on **Inputs** determines whether the assumed return represents compound growth or an average for a single year. The model treats each year’s market returns as independent. Wages, inflation and the end age follow the workbook. The same workbook, seed and engine version reproduce the simulations. Use `--volatility 0` to run the market simulations at the plan’s assumed returns. The early-loss tests still run separately.

The report uses the tax and healthcare rules in the workbook. Use a copy built with the current **Tax_Tables** layout. See the [model guide](model-guide.md#tax-rates-and-healthcare-credits) for the tax and coverage assumptions.
