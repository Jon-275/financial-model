# Portfolio risk analysis

Use market risk to see how often your plan can pay its bills when investment returns change. Run the analysis in desktop Excel and read the results on **Summary**. You can inspect the assumptions, formulas and individual trials in the workbook.

## Run the analysis

1. Save a private copy of the macro-enabled `.xlsm` workbook before entering personal figures. Desktop Excel runs the macro. Excel for the web can display saved results. Follow your organization’s macro settings. [Microsoft compatibility guidance](https://support.microsoft.com/en-us/excel/work-with-vba-macros-in-excel-for-the-web)
2. Complete **Inputs** and review the base funding result on **Summary**. Resolve missing inputs and calculation messages first.
3. Review the portfolio mix on **Asset_Allocation**. On **Risk_Settings**, enter how much spending you could cut and review the inflation assumptions. Choose 1 to 1,000 pairs. Each pair tests the same market conditions twice, once with fixed spending and once with flexible spending. Use 100 pairs for a quick comparison or 1,000 for the final report. With the 33 extra stress tests, these settings give 233 or 2,033 outcomes. Keep the random seed the same to compare plans under the same market conditions.
4. Choose **Run market risk** beside Summary’s navigation. Excel shows progress in the status bar. Wait for **Risk results are current** before printing. If the run fails, read the message on **Risk_Settings** in A44, correct the issue and run again.
5. Read Summary’s market paths, spending flexibility and inflation, longevity and medical tests. Open **Risk Results** for more detail or **Funding_Paths** for the annual cash flows of the three selected market paths.
6. Choose the dollar view, then print **Summary** or save it as a PDF.

Keep Excel’s calculation setting on **Automatic** while editing. Changes to the active financial assumptions mark the previous risk results out of date. Choose **Run market risk** again to replace them. A cancelled, failed or incomplete run does not supply current report results. Switching the dollar view uses the same completed run.

Allow time for the final run. In desktop Excel for Mac, 1,000 pairs took about 8 minutes for the example plan (ages 35–95) and 33 minutes for the full age range (0–125). Your run time will depend on your Mac and your plan.

## Read the Summary

- **Success rate.** This is the percentage of randomly generated, fixed-spending paths that paid every bill and tax on time through your end age. Even one cent unpaid counts as a shortfall, including a bill repaid later. The rate depends on your assumptions and the simulated returns. Flexible spending and stress tests are reported separately.
- **Middle 80% of simulated balances.** At each age, about 10% of balances fall below the band and 10% above it. The line shows the middle balance. Different paths can set the edges from year to year. Balances are before estimated tax on traditional savings.
- **Low, median and high funding examples.** Three complete fixed-spending paths near the 10th, 50th and 90th percentiles of ending balances in today’s dollars. Their identities stay the same when the dollar view changes. A positive ending portfolio can coexist with an earlier shortfall when money was unavailable to pay that year’s costs.
- **First shortfall.** The age and outstanding amount in the first year with unpaid costs. **Risk Results** separates unpaid spending and tax and shows the income, HSA payments and account withdrawals available and used that year. Unpaid bills carry forward without interest.
- **Fixed and flexible spending.** Both policies use the same returns and inflation for each paired path. Compare their success rates with the number of paths and years requiring cuts.
- **Required stress tests.** Summary page 4 shows the inflation, longer-retirement and medical-budget comparisons. They are separate planning tests, each starting from the base plan.

The six-page **Summary** begins with the outlook and retirement cash flow, then covers market risk, inflation and longevity. The final two pages record the input, spending and account decisions. The selected dollar view also applies to **Risk Results**, **Annual Detail** and **Funding_Paths**. Input decisions retain their opening-year amounts. [Print setup](model-guide.md#print-your-plan)

## Spending flexibility

On **Risk_Settings**, enter the share of non-mortgage living costs you could reduce, the cut to that share, and the balance that would trigger it. The trigger compares each year’s opening portfolio with that year’s opening balance in the base plan.

The longer-retirement tests keep spending fixed. In **Market_Paths**, the first extra year can still use the base plan’s final balance for comparison. In later years the comparison cells are blank, and the actual spending cut stays at zero.

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

The medical tests follow the entered survivor share. Long-term care sits outside the minimum medical budget. The model assumes the extra reserve cannot be reimbursed from a health savings account (HSA). Check actual premiums and out-of-pocket costs separately.

The early stock tests apply the listed stock return for the selected number of years from retirement. Other assets and later years retain the plan’s assumptions. **Risk Results** shows all early-loss outcomes for both spending policies.

These tests do not assign probabilities to inflation, medical costs or survival. A selected end age is a planning horizon rather than a mortality estimate.

## Returns and annual prices

The base plan and market trials use each account’s investment mix, including the HSA, and the fee on **Inputs**. Enter expected returns before inflation. The change in buying power is `(1 + return after fees) / (1 + inflation) - 1`. The annual raise on **Inputs** already includes inflation.

The **Investment return convention** on **Risk_Settings** selects compound growth or a simple annual average. The model lets returns on different assets move together. It uses a lognormal distribution and keeps expected returns, volatility and correlations fixed throughout the plan. Only the listed stress tests impose a particular sequence of losses. The model does not replay historical crises.

Choose **Vary annual inflation?** to let inflation change around your expected rate. You can set how much it varies, how long a change tends to last and its upper and lower limits. The model varies inflation independently of investment returns. A market loss does not automatically trigger higher inflation.

Each path updates living costs, medical prices, indexed tax parameters, Social Security and purchasing-power measures from its own price sequence. Fixed-dollar costs and tax thresholds that do not rise with inflation stay fixed. After Social Security starts, the model uses annual price changes and keeps benefits at their previous high if prices fall. Actual benefit adjustments use quarterly CPI-W figures.

In the future-dollar view, the balance band uses separately calculated nominal percentiles. A representative path’s amounts use that path’s prices at the relevant age. The base plan continues to use its expected inflation rate.

## Repeat and inspect a run

Excel calculates annual cash flows in the visible worksheets. When the workbook verifies that flexible spending would leave every year unchanged, both spending policies use the same amounts. Each policy still has its own result. The report updates when the whole run passes its checks.

To compare spending policies, enter a fixed trial number in **Risk_Trials** cell B3, then enter its flexible partner. For random paths, add 1,000 to the fixed number. For early-loss tests 2001–2013, add 100. Both records are saved even when their outcomes match. Columns I and J show how many years spending was cut and the total reduction in today’s dollars. The seven inflation, longevity and medical tests use fixed spending. Enter 0 in B3 to return to the base plan.

Keep the same model version, financial assumptions, seed, path count and risk settings to reproduce a comparison. A fixed seed makes the sampled returns repeatable. A different seed can produce a different success rate or balance range. More paths reduce sampling variation but do not correct unrealistic assumptions.

**Risk_Trials** records outcomes and annual balances. **Market_Paths** shows returns, inflation and cuts for a selected trial. **Random_Numbers** shows the seed and random-number calculation. Sheets beginning with **Trial_** use the same annual formulas as the base plan for taxes, healthcare, account access, withdrawals and unpaid costs. Select a trial on **Risk_Trials** to trace its figures.

Open **Funding_Paths** to follow the low, median and high paths year by year. **Funding_Path_Data** holds their saved income, withdrawals, account balances and unpaid costs in each year’s dollars. **Risk_Data** calculates the report statistics and checks that the current assumptions match the completed run.
