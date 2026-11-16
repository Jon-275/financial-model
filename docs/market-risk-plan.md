# Market risk and funding decisions

The printed plan should show whether the base case covers its costs, how often it does so when markets vary, and which spending choices make a difference. The reader should be able to trace a shortfall to the year, the bills that went unpaid, and the income and savings the plan could use.

## Printed plan

Keep the two-page input record and three-page Summary. Use the existing type sizes, margins, and page breaks. Keep the first page focused on the funding answer and the third page focused on decisions. Put the full calculations and risk results behind **Appendix**, with all sheets visible.

| Page | Content |
| --- | --- |
| Retirement outlook | Base projection, the percentage and count of fixed-spending runs that paid all required costs, and portfolio balances by age. Use a thin dashed line for the base projection, a solid line for the simulated median, and a light band for the middle 80% of simulated balances. Include the selected end age and dollar view. |
| Retirement cash flow | The annual funding table around retirement, cash available before age 60, and the sources the plan draws on. |
| Funding decisions | Three complete simulated paths with low, median, and high ending balances. Show the first age with unpaid costs and the unpaid amount. Compare fixed and flexible spending on the same market paths, including how often the cut is used. Show an early-loss example and leave room to record a decision. |

Use direct labels and one restrained accent color for the range. Give the chart and the funding table room to breathe. Verify legibility on screen and in a grayscale printout. A report with missing or older results should make the refresh step visible in the same place as the risk number.

Put **Dollar view** next to the Summary navigation. Default to today’s dollars and label the units on every printed Summary page. For future dollars, use each amount’s year and the plan’s inflation assumption. The switch applies to the outlook, cash flow, and funding examples. The simulation stores opening-year dollars so the choice of view does not require a refresh.

## Calculation and refresh

1. Use the saved scenario, selected end age, asset mix, and risk settings from the workbook. Reconcile the simulation engine to the workbook before applying any market variation. Use the existing annual tax, healthcare, funding, and account access calculations for every run.
2. Use the same seeded market path for fixed and flexible spending. A path succeeds when every modeled year pays required spending and taxes. Show the number of successes alongside the percentage. Report the spending cut and show how often it applies.
3. Rank actual fixed-spending runs by ending balance. Select runs near the 10th, 50th, and 90th percentiles, then use those exact return sequences for the funding examples. For the first shortfall, record required spending, tax due, each unpaid amount, and income and account withdrawals. Preserve the same path for the flexible-spending comparison.
4. Calculate the chart range separately at each age from all fixed-spending runs. Use a single age axis and a single dollar axis. The center line and shaded range describe balances across runs at each age. The three funding examples each follow one complete sequence.
5. Save the results to a new local copy of the workbook. Record the scenario, assumptions, run date, seed, trial count, and engine version. Have the workbook check the current effective inputs and publish the figures in Summary only while they match the run. Changes to an inactive scenario should leave the selected scenario’s results available. The public workbook should contain fictional inputs and an empty risk results record.

## Release gates

- Reconcile the base case and zero-volatility paths. Test the default run, reproducibility, one-cent shortfalls, restricted balances, flexible-spending triggers, and age 125.
- Check ranking ties, small trial counts, all-funded and all-shortfall cases, correct year-by-year percentiles, and the early-loss comparison. Verify the three funding examples against the engine audit for their exact return paths.
- Check workbook output with no results, current results, changed assumptions, scenario switches, restored assumptions, and invalid input. Verify both dollar views against the annual cash-flow schedules and check that the success rate and shortfall ages agree. Confirm that a failed refresh preserves the original file and never creates a partial deliverable.
- Measure refresh time and peak memory on the standard scenario and the longest supported horizon. Confirm that ordinary workbook recalculation requires only the light display and freshness formulas.
- Review the workbook in desktop Excel and check US Letter and A4. Verify two input pages, three Summary pages, clean rendering of the range and funding table, and navigation to **Risk Results**. Record the platforms and versions tested.
- Update the [model guide](model-guide.md), [risk guide](risk-analysis.md), build and release checks, and workbook instructions. Keep maintainer steps separate from the reader’s refresh instructions.

## Verification

The rebuilt workbook passed 80 scenario checks, 34 risk tests, and the workbook refresh test. The refresh test covered age 125, both dollar views, active and inactive scenario edits, a change to the disability insurance rate, and protection against overwriting the source. The baseline matched the previous release across 26,680 cells in the projection, cash flow, Roth access, healthcare, and checks schedules.

Excel for Mac opened the refreshed workbook and switched the cash-flow table, charts, and funding examples to future dollars while retaining the success rate. The five-page package fit US Letter and A4. Excel for Windows print verification remains open in the [project plan](../PLAN.md#remaining-maintenance-work).
