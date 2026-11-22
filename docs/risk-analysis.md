# Market risk

Use market risk to test whether the plan pays its bills when returns change. Complete **Inputs** and review the base result on **Summary** first.

[Workbook guide](model-guide.md) · [Calculation assumptions](model-methods.md#market-returns-and-inflation)

<a id="returns-and-annual-prices"></a>

## Risk_Settings

Review the investment mix on **Asset_Allocation**, then set the following on **Risk_Settings**:

| Setting | What it does |
| --- | --- |
| Spending-cut trigger, B6 | Compares the year's opening portfolio with the base plan's opening balance for that year. |
| Cut, B7, and flexible share, B8 | Set how much of non-mortgage living costs can be cut. B10 shows the combined effect. |
| Paired market paths, B29 | Use **100** for comparisons or **1,000** for a final report. Each pair uses the same markets for fixed and flexible spending. |
| Seed, B30 | Keep it unchanged when comparing plans under the same sampled markets. |
| Return and inflation settings | Review the return convention and whether inflation varies by year. [How these assumptions work](model-methods.md#market-returns-and-inflation) |

You can choose 1–1,000 pairs. Each run also includes 33 stress outcomes: 100 pairs give 233 outcomes; 1,000 give 2,033.

### Spending flexibility

A 15% cut to a 30% flexible share reduces eligible living costs by **4.5%** when the rule applies. Mortgage payments, healthcare and one-time bills remain due. The rule is checked each year, so cuts can recur.

Use a cut you could sustain. Read the number of paths and years needing cuts alongside the success rate. The median number of years includes paths with no cuts.

## Run the analysis

1. Resolve input and calculation messages. Keep Excel calculation set to **Automatic**.
2. On **Summary**, choose **Run market risk** with macros permitted in desktop Excel. Follow your organization's macro policy.
3. Watch the status bar. Wait for **Risk results are current** before reading or printing the results.

The documented 1,000-pair example took about **8 minutes**; the full age range took **33 minutes**. Runtime depends on your Mac and plan. Use **Cancel** to stop a run and wait for confirmation.

If a run fails, read **Risk_Settings A44**, correct the issue and rerun. Editing plan assumptions makes old results out of date. A failed, cancelled or incomplete run does not provide current results. Changing the dollar view does not require a new run.

## Read the Summary

Read pages 1, 3 and 4 together:

| Result | Meaning |
| --- | --- |
| Success rate | The share of random fixed-spending paths that paid every bill and tax on time through the end age. Even one cent unpaid counts as a shortfall, including a bill repaid later. Flexible spending and stress tests are separate. |
| Middle 80% of balances | At each age, about 10% of balances fall below the band and 10% above it. Different paths can set its edges at different ages. The line is the median. Balances are before estimated tax on traditional savings. |
| Low, median and high paths | Three complete fixed-spending paths near the 10th, 50th and 90th percentiles of ending balances in today's dollars. Their identities stay the same in either dollar view. |
| First shortfall | The age and unpaid amount in the first year bills or taxes were not paid on time. A positive ending portfolio can still have an earlier shortfall. |
| Fixed and flexible spending | Results under the same markets, with and without the spending-cut rule. Check how often cuts were needed and how long they lasted. |

Results depend on the model's assumptions. More paths reduce sampling variation; they do not correct unrealistic assumptions. Keep the workbook version, inputs, settings, path count and seed the same to repeat a comparison.

## Inflation, longevity and medical tests

Page 4 shows separate tests, each starting from the base plan:

| Test | Change |
| --- | --- |
| Inflation at 4% or 5% | Use that annual inflation rate and the model's linked price adjustments. |
| Plan through age 100 or 105 | Extend the selected end age to at least that age. These tests keep spending fixed. |
| Higher inflation and weaker returns | Use 5% inflation, returns two percentage points lower and an end age of at least 105. |
| Medical budget of $18,000 or $25,000 | Add a reserve when the couple's annual retirement medical costs fall below that amount at opening-year prices. |

Medical tests use the entered survivor share. Long-term care is outside the medical minimum. The extra reserve cannot be paid from the HSA in the model.

Early-loss tests apply the listed stock return for the selected years from retirement. Other assets and later years keep the plan's assumptions. These tests do not assign probabilities to inflation, medical costs or lifespan. The end age is a planning horizon.

## Risk Results

This is the second tab after **Appendix**. Use it for funding outcomes, early stock losses and balances by age. For a shortfall, check income, HSA payments, available withdrawals and unpaid tax and spending in that year.

<a id="repeat-and-inspect-a-run"></a>

## Funding_Paths

Follow the low, median and high paths year by year. **Funding_Path_Data** holds their saved source values in each year's dollars.

Summary's dollar view also applies to **Risk Results**, **Annual Detail** and **Funding_Paths**. Input decision records keep opening-year amounts. [Print the report](model-guide.md#print-your-plan)

To inspect an individual trial or check saved data, use the [trial-sheet guide](workbook-schedules.md#market-trial-calculations) and [runner reference](macro-runner.md#inspect-a-trial).
