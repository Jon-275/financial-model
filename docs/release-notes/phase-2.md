# Phase 2 release record

This release adds an optional healthcare schedule by person, improves tax calculations, and replaces the one-year Roth conversion illustration with a comparison that uses the main funding calculation. It keeps the selected scenario, spendable other income, existing household healthcare budget and age-125 limit.

| Area | What changed |
| --- | --- |
| Healthcare | Detailed mode records separate Medicare and Marketplace coverage, gross premiums, qualified expenses and HSA payments. Marketplace credits are booked once. Medicare surcharges use MAGI and filing status from the modeled lookback year. Missing coverage inputs prevent a funded headline. |
| Annual funding | The draw search evaluates the cash result on both sides of the Marketplace income floor, the 133% contribution-rate change and the 400% eligibility limit. If no evaluated draw funds the year, the search uses the draw with the smallest observed gap and reports the shortfall. Optional conversions are skipped in the main plan if they cannot be funded. |
| Tax | Social Security payroll tax applies the wage ceiling to each earner. Federal AMT compares tentative tax with both regular income tax components. Ordinary taxable sales release basis proportionally and may realize a loss. Federal and California loss carryforwards remain separate. |
| Conversion comparison | Each trial uses the selected year's opening balances, filing status, spending, taxes and Marketplace eligibility. The view shows actual versus requested conversion and withholds comparative costs for incomplete requests. It does not change the active plan. |

The workbook uses the published 2027 federal tax and Social Security figures and the published 2026 California brackets. After those years, it estimates bracket changes with modeled inflation. The 2027 Marketplace contribution percentages are held constant after 2027, while the poverty guideline is indexed. Marketplace eligibility is entered by the user. This annual estimate does not model monthly changes, return reconciliation, Alaska or Hawaii, dependents or exceptions to the standard eligibility range. Medicare surcharges use the available 2026 schedule as the forecast base. Detailed HSA reimbursements exclude ordinary Marketplace and Medigap premiums. Confirm HSA contribution eligibility separately.

The tax calculation covers ordinary W-2 and retirement income, qualifying gains and dividends, and standard deductions for joint or single filing. California's personal exemption phaseout is interpolated for planning. The model does not allocate inherited capital losses, simulate tax lots, or implement discretionary harvesting. Those strategies remain in the [portfolio guidance](../portfolio-management.md).

## Sources

- [IRS 2027 federal tax adjustments](https://www.irs.gov/pub/irs-drop/rp-26-35.pdf), [SSA wage base](https://www.ssa.gov/oact/cola/cbb.html) and [FTB 2026 California brackets](https://www.ftb.ca.gov/about-ftb/newsroom/tax-news/2026/10.html).
- [IRS 2027 Marketplace percentages](https://www.irs.gov/irb/2026-31_irb), [HHS poverty guidelines](https://aspe.hhs.gov/topics/poverty-economic-mobility/poverty-guidelines/prior-hhs-poverty-guidelines-federal-register-references), [HSA rules](https://www.irs.gov/publications/p969) and [SSA Medicare lookback](https://secure.ssa.gov/POMS.nsf/lnx/0601101010).
- [IRS AMT instructions](https://www.irs.gov/instructions/i6251), [federal capital loss instructions](https://www.irs.gov/instructions/i1040sd) and [California capital loss instructions](https://www.ftb.ca.gov/forms/2025/2025-540-d-instructions.html).

## Build and verify

Use the [current build instructions](../../README.md#build-the-next-edition) and [test guide](../../tests/README.md).

On November 6, 2026, all 65 regression and synthetic scenarios passed in LibreOffice. Three representative cases also passed on full-size copies, including a plan funded only below the Marketplace limit and a conversion that is better funded by preserving the credit. The scripted build passed end to end, and the packaged workbook passed independent cash, account and healthcare reconciliation in all 73 active years. Existing scenario cells were preserved, and all five chart objects remain in the file. Phase 2 verification used LibreOffice. See [Phase 3](phase-3.md#microsoft-excel-verification) for later Excel verification.
