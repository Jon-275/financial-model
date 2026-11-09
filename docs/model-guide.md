# Retirement model guide

1. **Pick a scenario.** On **Inputs**, enter 1 for the base plan, 2 to retire at 60 or 3 to test 10% lower retirement spending. Edit the blue-text cells in the matching column, D, E or F.
2. **Check the basics.** Review your ages, wages, savings and spending. Enter the last age you want to see in row 53, up to 125. Use today’s dollars unless a note says otherwise. Use the **+** signs beside the rows to open optional sections.
3. **Read the answer.** The top of **Summary** shows whether the plan covers its costs, the first year it falls short and the ending balance in today’s dollars. Open **Cash_Flow** to see what went unpaid in a particular year.

The current-age row shows the money you start with. The forecast begins the following year.

## Confirm cash and Roth inputs

| Where | What to enter |
| --- | --- |
| `Inputs` rows 12–14 | Enter household wages separately from other income. Enter the full annual wage increase in row 13, including inflation. A 5% raise takes $80,000 to $84,000 in the first projected year. |
| `Inputs` row 52 and rows 86–88 | Enter annual other taxable income in today’s dollars. Set the share that produces spendable cash (0 to 100%) and the primary person’s first and last ages for receiving it. The full amount is taxable in active years. The age bounds are inclusive. A start age of 0 applies from the first modeled year; the same start and end age gives a single year. |
| `Inputs` rows 46 and 89 | Set the cash share of external tax-exempt interest. Avoid counting the same interest in brokerage returns and as external cash. |
| `Inputs` rows 27–30, 90 and 92 | Classify the opening balance as Roth IRA or workplace Roth. Split it between the two owners. For an IRA, enter only opening principal already available, and confirm the first Roth IRA tax year for each owner. The owner split applies to both opening balance and opening IRA principal. An opening balance with both account types cannot be represented by the single opening type field. |
| `Inputs` rows 35, 91 and 98 | Set the combined annual Roth IRA savings target and the planned owner split. The model inflates the target and uses remaining cash after obligations. It applies modeled per-owner IRA limits and a wages-based funding cap. Choose 1 to assume the savings route is eligible and does not incur extra tax; choose 0 to turn it off. Verify contribution eligibility, account ownership and any IRA pro-rata tax separately. |
| `Inputs` rows 93–96 | For workplace Roth, enter each owner’s first designated Roth tax year and the age when that plan permits distributions. Leave the first year at 0 if unknown or distribution age at 126 to exclude those funds. The model also requires the account clock and an owner age of at least 60. |
| `Inputs` row 97 | Identify the owner of the modeled traditional account. The model uses that owner’s age for elective access, required distributions, the 72(t) schedule and the conversion cutoff. Conversions go to that owner’s Roth IRA. The workbook uses one traditional balance and does not model separately inherited accounts. |

The added fields initially treat other taxable income and external tax-exempt interest as fully spendable, split opening Roth and planned IRA savings equally, and disable workplace Roth access. Confirm those settings for each scenario before relying on a result.

## Read the charts and results

1. **Start with Summary.** The balance chart shows your savings by age in today’s dollars and marks when retirement starts. The spending chart shows how income and withdrawals pay bills around retirement and highlights unpaid costs. A shortfall means at least one year leaves required spending or tax unpaid. The gap to the savings target compares the portfolio with a rule of thumb. If Summary shows **CHECK PLAN**, follow the note beside it. Open **Checks** when the note asks you to review the calculation.
2. **Trace the year in Cash_Flow.** The model estimates tax on that year’s withdrawals and pays tax first, including any prior unpaid tax. Remaining cash covers required spending, then optional Roth IRA savings. Unpaid tax carries forward and reduces after-tax net worth; unpaid spending is reported for that year. A bounded search chooses the needed eligible draw. The model skips a planned conversion if it cannot fund the associated obligations with eligible assets.
3. **Review access in Roth_Access.** The model separates owner balances and regular IRA principal, conversion principal and earnings. IRA conversion principal generally becomes available in the model after five tax years or when the owner reaches age 60. IRA earnings require both age 60 and an established five-year account clock. Whole-year projections use age 60 for age 59½. Workplace Roth requires verified plan access and a qualified distribution. The model does not assume rollovers, an age-55 separation exception, or nonqualified workplace withdrawals. See the [IRS Roth comparison](https://www.irs.gov/retirement-plans/ten-differences-between-a-roth-ira-and-a-designated-roth-account) and [IRA distribution guidance](https://www.irs.gov/publications/p590b).
4. **Check the bridge year by year.** “Available bridge assets at retirement” is opening brokerage plus eligible Roth in the first bridge year before the primary person reaches age 60. It is a balance at one point in time. For annual funding, also review income, tax, other eligible withdrawals and changes in access as each owner ages.
5. **Check the reconciliations.** **Checks** compares annual cash, tax and Roth balances and flags excess Roth draws, negative balances and funding-calculation issues. A year can reconcile and still be unfunded. **Funding_Calc** is hidden because it contains the draw search, not user inputs.

## Compare risks and Roth conversions

**Risk_Settings** holds the inputs for the separate risk report. The report recalculates cash, taxes and account access for each market path and compares fixed and flexible spending under the same returns. Rerun it after changing plan inputs or risk settings. [Portfolio risk analysis](risk-analysis.md)

Use **Income_Targeting** to compare conversion amounts for one year. Record deliberate tax harvesting in [Portfolio management](portfolio-management.md).

## Healthcare and taxes

Keep **Detailed healthcare model?** at `0` to use the household budget. To estimate coverage by person, set it to `1` and use the **+** sign beside row 101 to open rows 102–124. Working-year premiums are part of the working budget. Enter annual quotes in today’s dollars before credits, including the Marketplace benchmark where applicable. Enter eligible medical expenses separately. If Medicare needs income from before the projection, enter that year’s modified adjusted gross income. Resolve any missing inputs flagged on **Healthcare** before reading Summary or the risk report.

**Healthcare** shows each person's coverage, gross costs, the Marketplace credit, and the HSA payment. **Cash_Flow** books the credit once and pays HSA-eligible costs only once. The tax model distinguishes the two earners and tracks federal and California capital losses separately. Check the wage split and any prior-year loss carryforwards on Inputs rows 126–128.

On **Income_Targeting**, select a retirement age (or leave `0` for the first retirement year), then enter a conversion amount in that year's dollars. The table shows the amount the plan can actually convert, the withdrawals needed to pay the bill, and the change in tax and net healthcare. The comparison applies to one year. To see the effect on later balances, Roth access or Medicare surcharges, change the main scenario and rerun it.

See the [Phase 2 release record](release-notes/phase-2.md) for assumptions, sources and verification.
