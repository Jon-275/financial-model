# Portfolio risk analysis

Market risk shows how often your plan pays its bills and taxes when returns vary from year to year. It uses the selected scenario, investment mix, end age, and **Risk_Settings**. Save and close your private workbook before refreshing the results.

## Refresh the workbook

Use Python 3.11 or later and LibreOffice Calc. Put `soffice` on your path, or set `SOFFICE_BIN` to the executable. Use the runner from the same release as your workbook. From the repository root, install the requirements once and run the following command with paths to your private workbook and a new output file.

```sh
python3 -m pip install -r tools/requirements.txt
mkdir -p private
python3 tools/run_risk.py private/my-plan.xlsx \
  --output private/my-plan-with-risk.xlsx
```

Open the new workbook in Excel and review **Summary**. The runner checks that its baseline matches the workbook before simulating returns. By default it uses 1,000 market paths and seed `2026`. The files and calculations stay on your computer.

Keep Excel’s calculation setting on **Automatic** so changes update the base plan and risk status before you print. When **Summary** says **Refresh market risk**, save and close your plan, then run the command again with a new output name. The risk figures return when you refresh or restore the assumptions used in the saved run.

## Read the Summary

- **Success rate.** The percentage and count of runs that paid every year’s required spending and taxes through the end age. A run with at least one cent unpaid in any modeled year counts as a shortfall. The savings benchmark is a separate figure.
- **Simulated balance range.** The line shows the median balance at each age. The shaded band shows the middle 80% of balances under fixed spending. The chart uses the **Dollar view** selected on Summary. Balances are before estimated tax on traditional savings.
- **Low, median, and high funding examples.** These are three complete fixed-spending runs ranked by ending balance. For a path with a shortfall, **Risk Results** shows taxes and spending due in the first year with unpaid costs, what went unpaid, and the income and withdrawals available and used that year.
- **Fixed and flexible spending.** These comparisons use the same market paths. The flexible rule reduces part of recurring retirement living costs when the opening portfolio falls below the share of the planned balance set on **Risk_Settings**. Review the size of the cut and how often the rule used it. Fixed costs such as mortgage payments and healthcare stay outside the cut.
- **Early retirement losses.** The stock returns listed on **Risk_Settings** are applied to the first retirement years. Other assets and later years follow the plan’s assumed returns. **Risk Results** shows the early-loss outcomes for both spending policies.

Print **Inputs** and **Summary** for the five-page plan. Switching the dollar view changes the amounts shown on Summary. The success rate and shortfall ages carry across views. Open **Risk Results** from **Appendix** for details in today’s dollars. [Print setup](model-guide.md#print-your-plan)

## Return assumptions

The simulation uses each account’s investment mix and the fee on **Inputs**. The base plan uses the Roth return for the health savings account. The simulation uses the health savings account’s mix to estimate how much its return varies.

The **Investment return convention** on **Risk_Settings** (cell B9) determines whether the assumed return represents compound growth or a simple annual average. The simulation draws independent market returns each year. Wages, inflation, and the end age follow the workbook. Use `--volatility 0` to replay the simulation at the plan’s assumed returns. Early-loss tests still run.

Use the same workbook, engine version, seed, trial count, and volatility setting to reproduce the results. The success rate depends on the plan and the return assumptions used. See the [model guide](model-guide.md#tax-rates-and-healthcare-credits) for tax and coverage assumptions.

## Export results for review

Use a new `.json` output name for a machine-readable report. The JSON includes the run settings, sampling intervals, annual percentiles, and the first-shortfall details for the representative paths. Use `--cached` with JSON only after the workbook has been recalculated and saved. The baseline check still runs.

```sh
python3 tools/run_risk.py private/my-plan.xlsx \
  --trials 1000 --seed 2026 --output private/portfolio-risk.json
```

You can also turn that JSON into a standalone Excel report. Set `ARTIFACT_NODE_MODULES` to the bundled `node_modules` directory and set `NODE_BIN` if you use a separate Node executable. [Build instructions](../README.md#rebuild-the-workbook)

```sh
(
  set -eu
  render_dir=$(mktemp -d)
  trap 'rm -rf -- "$render_dir"' EXIT
  cp tools/render_risk.mjs "$render_dir/"
  ln -s "${ARTIFACT_NODE_MODULES:?Set the runtime node_modules path}" "$render_dir/node_modules"
  "${NODE_BIN:-node}" "$render_dir/render_risk.mjs" \
    private/portfolio-risk.json private/portfolio-risk.xlsx
)
```
