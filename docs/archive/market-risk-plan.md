# Market risk and funding decisions

Archived design for the five-page, macro-free workbook. For instructions for the macro-enabled edition, see the [market risk guide](../risk-analysis.md).

The printed plan shows whether the base case covers its costs, how often it does so when markets vary, and which spending choices affect the result. **Risk Results** shows the first shortfall age when one occurs, the bills that went unpaid, and the income and savings the plan could use.

## Printed plan

In that edition, all five pages printed from **Summary**. **Risk Results** appeared before **Appendix**. **Risk_Data** held the saved run.

| Page | Content |
| --- | --- |
| 1. Retirement outlook | Base projection, fixed-spending success rate and count, and the portfolio balance range by age. The dashed line shows the base projection. The solid line and shaded band show the median and middle 80% of simulated balances. |
| 2. Retirement cash flow | Cash sources and payments around retirement, including savings available before age 60 where applicable. |
| 3. Risk and decisions | Three simulated paths with low, median, and high ending balances. Each example shows the first unpaid costs when a shortfall occurs. The page compares fixed and flexible spending, includes an early-loss example, and leaves room to record a decision. |
| 4. Input decisions | The scenario and assumptions used for the plan. |
| 5. Spending and account decisions | The spending, healthcare, and account choices used for the plan. |

The report shows the refresh step where the risk figure would normally appear if the workbook needs a new run.

**Dollar view** next to the **Summary** navigation defaults to today’s dollars. It changes the outlook, cash flow, and funding examples, as well as the dollar amounts on **Annual Detail** and **Risk Results**. Future dollars use each amount’s year and the plan’s inflation assumption. The simulation stores opening-year dollars, so both views use the same run.

## Calculation and refresh

1. Use the saved scenario, selected end age, asset mix, and risk settings from the workbook. Reconcile the simulation engine to the workbook before applying any market variation. Use the existing annual tax, healthcare, funding, and account access calculations for every run.
2. Use the same seeded market path for fixed and flexible spending. A path succeeds when every modeled year pays required spending and taxes. Show the number of successes alongside the percentage. Report the spending cut and show how often it applies.
3. Rank actual fixed-spending runs by ending balance. Select runs near the 10th, 50th, and 90th percentiles, then use those exact return sequences for the funding examples. For the first shortfall, record required spending, tax due, each unpaid amount, and income and account withdrawals. Preserve the same path for the flexible-spending comparison.
4. Calculate the chart range separately at each age from all fixed-spending runs. Use a single age axis and a single dollar axis. The center line and shaded range describe balances across runs at each age. The three funding examples each follow one complete sequence.
5. Save the results to a new local copy of the workbook. Record the scenario, assumptions, run date, seed, trial count, and engine version. **Summary** and **Risk Results** show figures while the active assumptions match the saved run. Edits to an inactive scenario leave the selected scenario’s results available. The public workbook opens with the verified fictional run already saved.

## Release gates

- Reconcile the base case and zero-volatility paths. Test the default run, reproducibility, one-cent shortfalls, restricted balances, flexible-spending triggers, and age 125.
- Check ranking ties, small trial counts, all-funded and all-shortfall cases, correct year-by-year percentiles, and the early-loss comparison. Verify the three funding examples against the engine audit for their exact return paths.
- Check workbook output with no results, current results, changed assumptions, scenario switches, restored assumptions, and invalid input. Verify both dollar views against the annual cash-flow schedules and check that the success rate and shortfall ages agree. Confirm that a failed refresh preserves the original file and never creates a partial deliverable.
- Measure refresh time and peak memory on the standard scenario and the longest supported horizon. Confirm that ordinary workbook recalculation requires only the light display and freshness formulas.
- Review the workbook in desktop Excel and check US Letter and A4. Print **Summary** as three results pages and two decision pages. Check the balance range, funding table, and navigation to **Risk Results**. Record the platforms and versions tested.
- Update the [model guide](../model-guide.md), [risk guide](../risk-analysis.md), build and release checks, and workbook instructions. Keep maintainer steps separate from the reader’s refresh instructions.

The first integrated market risk release passed 80 workbook scenarios, 34 risk tests, and the workbook refresh test. [Market risk release](https://github.com/Jon-275/financial-model/pull/11)

The release record describes its five-page **Summary**, appendix layout and Excel checks. [Five-page release record](../release-notes/workbook-layout.md)
