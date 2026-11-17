# Portfolio risk analysis

Market risk tests whether income and accessible savings pay the plan’s spending and taxes when investment returns vary. The macro-enabled workbook runs the analysis in desktop Excel and displays the findings in **Summary**. Its formulas, market assumptions and trial results remain visible for inspection.

## Run the analysis

1. Save a private copy of the macro-enabled `.xlsm` workbook before entering personal figures. Use desktop Excel and follow your organization’s macro rules. Excel for the web cannot run VBA. The workbook does not require a global change to macro security. [Microsoft compatibility guidance](https://support.microsoft.com/en-us/excel/work-with-vba-macros-in-excel-for-the-web)
2. Complete **Inputs** and review the base funding result on **Summary**. Resolve missing inputs and calculation messages first.
3. Review **Risk_Settings** and **Asset_Allocation**. Choose the number of paired paths, seed, spending policy and inflation mode. Use fewer paths for a quick comparison and 1,000 for the full run.
4. Choose **Run market risk** beside Summary’s navigation. The run shows progress and supports cancellation. Wait for the results to be marked current.
5. Read Summary’s market paths, spending flexibility and inflation, longevity and medical tests. Open **Risk Results** for more detail or **Funding_Paths** for the annual cash flows of the three selected market paths.
6. Choose the dollar view, then print **Summary** or save it as a PDF.

Keep Excel’s calculation setting on **Automatic** while editing. Changes to the active financial assumptions mark the previous risk results out of date. Choose **Run market risk** again to replace them. A cancelled, failed or incomplete run does not supply current report results. Switching the dollar view uses the same completed run.

The macro selects each trial, recalculates the visible annual engine and captures the results. It updates the report after a complete batch passes its checks. It does not run automatically when the file opens, saves or prints. The calculations stay in the workbook and do not need a separate risk report, Python installation or external service.

Earlier `.xlsx` releases used the local Python runner. Use the documentation from the matching release for those files.

## Read the Summary

- **Success rate.** The percentage and count of simulated paths that paid every year’s required spending and taxes through the selected end age. At least one cent unpaid in any modeled year makes a path unsuccessful, including an amount repaid later. The rate depends on the entered assumptions and the sampled paths.
- **Middle 80% of simulated balances.** Each age has its own 10th and 90th percentile. Those values can belong to different paths. The line shows the median at each age. Balances are before estimated tax on traditional savings.
- **Low, median and high funding examples.** Three complete fixed-spending paths near the 10th, 50th and 90th percentiles of ending balances in today’s dollars. Their identities stay the same when the dollar view changes. A positive ending portfolio can coexist with an earlier shortfall when money was unavailable to pay that year’s costs.
- **First shortfall.** The age and outstanding amount in the first year with unpaid costs. **Risk Results** separates unpaid spending and tax and shows the income, HSA payments and account withdrawals available and used that year. Outstanding bills carry forward with zero financing interest.
- **Fixed and flexible spending.** Both policies use the same returns and inflation for each paired path. Compare their success rates with the number of paths and years requiring cuts.
- **Required stress tests.** Summary page 4 shows the inflation, longer-retirement and medical-budget comparisons. They are separate planning tests, each starting from the base plan.

Print **Summary** for all six pages: outlook, retirement cash flow, market risk and decisions, inflation and longevity, input decisions, and spending and account decisions. The selected dollar view also applies to **Risk Results**, **Annual Detail** and **Funding_Paths**. Input decisions retain their opening-year amounts. [Print setup](model-guide.md#print-your-plan)

## Spending flexibility

On **Risk_Settings**, enter the share of non-mortgage living costs you could reduce, the cut to that share, and the balance that would trigger it. The trigger compares each year’s opening portfolio with that year’s opening balance in the base plan.

The example cuts 15% of a 30% flexible share, reducing non-mortgage living costs by 4.5% in a year when the rule applies. Mortgage payments, healthcare and one-time bills remain in required costs. Tax follows the revised withdrawals. The rule is tested each year, so cuts can recur. The reported median number of years with cuts includes every flexible-spending path, including paths with no cuts.

Use a reduction you could sustain. Review the total budget effect before relying on the flexible-spending result.

## Inflation, longevity and medical tests

| Test | Change from the base plan |
| --- | --- |
| Inflation at 4% | Annual consumer inflation of 4%, with the model’s linked price adjustments |
| Inflation at 5% | Annual consumer inflation of 5%, with the model’s linked price adjustments |
| Plan through age 100 | Extend the selected end age to at least 100 |
| Plan through age 105 | Extend the selected end age to at least 105 |
| Higher inflation and weaker returns | Use 5% inflation, asset returns two percentage points lower and an end age of at least 105 |
| Medical budget of $18,000 | Add a reserve when annual retirement medical costs fall below $18,000 for a couple at opening-year prices |
| Medical budget of $25,000 | Add a reserve when annual retirement medical costs fall below $25,000 for a couple at opening-year prices |

The medical tests follow the entered survivor share. Their reserve excludes long-term care from the budget floor and is not assumed eligible for HSA reimbursement. These are cash-budget comparisons; actual premiums and out-of-pocket costs still need review.

The early stock tests apply the listed stock return for the selected number of years from retirement. Other assets and later years retain the plan’s assumptions. **Risk Results** shows all early-loss outcomes for both spending policies.

These tests do not assign probabilities to inflation, medical costs or survival. A selected end age is a planning horizon rather than a mortality estimate.

## Returns and annual prices

The base plan and trial engine use each account’s investment mix, including the HSA, and apply the fee on **Inputs**. Expected investment returns are nominal rates. Their purchasing power depends on inflation. The real return is `(1 + nominal return) / (1 + inflation) - 1`, after the selected fee treatment. Wage growth remains the total annual raise entered on Inputs, inclusive of inflation.

The **Investment return convention** on **Risk_Settings** selects compound growth or a simple annual average. The trial method uses correlated asset shocks and a lognormal return distribution. It assumes the same asset means, volatility and correlations throughout the plan. It does not model market regimes or historical crisis sequences beyond the explicit stresses.

The optional **Vary annual inflation?** setting adds year-to-year changes around the expected inflation rate. Its variation, persistence and bounds are illustrative controls. Inflation shocks are independent of asset shocks in this model. Review that limitation when comparing plans exposed to inflation and market losses at the same time.

Each path updates living costs, medical prices, indexed tax parameters, Social Security and purchasing-power measures from its own price sequence. Fixed nominal costs and statutory fixed thresholds retain their stated treatment. The Social Security calculation uses an annual price approximation and holds the benefit at its previous high after benefits begin. It does not reproduce the actual quarterly CPI-W timing of benefit announcements.

In the future-dollar view, the balance band uses separately calculated nominal percentiles. A representative path’s amounts use that path’s prices at the relevant age. The base plan continues to use its expected inflation rate.

## Repeat and inspect a run

Keep the same model version, financial assumptions, seed, path count and risk settings to reproduce a comparison. A fixed seed makes the sampled returns repeatable. A different seed can produce a different success rate or balance range. More paths reduce sampling variation but do not correct unrealistic assumptions.

**Risk_Trials** records outcomes and annual balances. **Market_Paths** shows returns, inflation and cuts for a selected trial. **Random_Numbers** shows the seed and random-number calculation. The **Trial_** schedules use the same annual compiler as the base plan for taxes, healthcare, account access, withdrawals and unpaid costs. Choose a trial on **Risk_Trials** to inspect that ledger.

Open **Funding_Paths** to follow the low, median and high paths year by year. **Funding_Path_Data** holds their saved income, withdrawals, account balances and unpaid costs in each year’s dollars. **Risk_Data** calculates the report statistics and checks that the current assumptions match the completed run.
