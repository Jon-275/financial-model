# Portfolio risk analysis

Market risk shows how often your plan pays its bills and taxes when returns vary from year to year. The downloaded workbook opens with results for the fictional example on **Summary**. Excel checks whether the assumptions still match and clears the risk figures when they change. The local Python runner performs new simulations and saves them in a refreshed copy of your Excel plan.

## Refresh the workbook

Use Python 3.11 or later and LibreOffice Calc. Put `soffice` on your path, or set `SOFFICE_BIN` to the executable. Use the runner from the same release as your workbook. From the repository root, install the requirements once and run the following command with paths to your private workbook and a new output file.

```sh
python3 -m pip install -r tools/requirements.txt
mkdir -p private
python3 tools/run_risk.py private/my-plan.xlsx \
  --output private/my-plan-with-risk.xlsx
```

Open the refreshed plan in Excel and review **Summary**. Print the plan from that same tab. **Risk Results** shows the outcomes and settings before **Appendix**. **Risk_Data** holds the saved balances by year, early-loss tests, funding examples and comparisons with the current assumptions. Before simulating returns, the runner checks that its baseline matches the workbook. By default it uses 1,000 market paths and seed `2026`. The files and calculations stay on your computer.

Keep Excel’s calculation setting on **Automatic** so changes update the base plan and risk status before you print. When **Summary** says **Refresh market risk**, save and close your plan, then run the command again with a new output name. The risk figures return when you refresh or restore the assumptions used in the saved run.

## Read the Summary

- **Success rate.** The percentage and count of runs that paid every year’s required spending and taxes through the end age. A run with at least one cent unpaid in any modeled year counts as a shortfall. The savings benchmark is a separate figure.
- **Simulated balance range.** The line shows the median balance at each age. The shaded band shows the middle 80% of balances under fixed spending. The chart uses the **Dollar view** selected on Summary. Balances are before estimated tax on traditional savings.
- **Low, median, and high funding examples.** These are three complete fixed-spending runs ranked by ending balance. For a path with a shortfall, **Risk Results** shows taxes and spending due in the first year with unpaid costs, what went unpaid, and the income and withdrawals available and used that year.
- **Fixed and flexible spending.** Both use the same market paths. Set the flexible share, the cut to that share, and the balance that triggers the cut on **Risk_Settings**. In the example, a 15% cut on a 30% flexible share cuts total living costs by 4.5%. The first retirement years with stock losses can show when that cut would matter most. Review how many runs used a cut and how many years it lasted when used.
- **Early retirement losses.** The stock returns listed on **Risk_Settings** are applied to the first retirement years. Other assets and later years follow the plan’s assumed returns. **Risk Results** shows the early-loss outcomes for both spending policies.

Print **Summary** for the five-page plan. The dollar view also applies to **Risk Results** and **Annual Detail**. The success rate and shortfall ages carry across views. [Print setup](model-guide.md#print-your-plan)

## Return assumptions

The base plan and the simulation use the investment mix for each account, including the health savings account (HSA). Both include the fee on **Inputs**.

The **Investment return convention** on **Risk_Settings** (cell B9) determines whether the assumed return represents compound growth or a simple annual average. The simulation draws independent market returns each year. Wages, inflation, and the end age follow the workbook. Use `--volatility 0` to replay the simulation at the plan’s assumed returns. Early-loss tests still run.

Use the same workbook, engine version, seed, trial count, and volatility setting to reproduce the results. The success rate depends on the plan and the return assumptions used. See the [model guide](model-guide.md#tax-rates-and-healthcare-credits) for tax and coverage assumptions.
