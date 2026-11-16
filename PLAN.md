# Retirement model implementation plan

The model tests whether household income and available savings can pay spending and taxes each year. The plan runs through the age selected on **Inputs**, up to 125.

## Completed

| Phase | Changes |
| --- | --- |
| [1. Cash flow and account access](docs/release-notes/phase-1.md) | The annual projection pays taxes first, then required spending. It reports the first year with unpaid costs and tracks each owner’s Roth access. |
| [2. Taxes, healthcare and conversions](docs/release-notes/phase-2.md) | The model adds healthcare by person, tests tax and Marketplace credit changes as withdrawals rise, and compares Roth conversion amounts for one year. |
| [3. Risk analysis and workbook usability](docs/release-notes/phase-3.md) | The risk report recalculates cash, taxes and account access for each set of market returns. **Start Here**, **Inputs** and **Summary** guide readers from setup to a five-page printable plan. |

Keep tax loss and gain harvesting in the [portfolio guide](docs/portfolio-management.md). The model includes the tax effect of investment sales needed to pay bills.

## Release checks

Use the [test and release guide](tests/README.md) for commands and Excel steps. Use fictional inputs for all shared workbooks and test results.

- Compare the candidate with the prior release. Record changes to inputs, defaults, formulas or charts.
- Check each scenario and confirm the projection stops at the selected age. Include age 125, an invalid age and a shortfall late in the plan.
- Exercise account access before and after eligibility dates, changes in tax when money is withdrawn, and a year when available savings run out.
- Reconcile cash and account balances for all active years. Run the risk report against the candidate and verify it matches the workbook when market returns follow the plan.
- In desktop Excel, check for repair prompts, recalculate from cleared formula results, and change the end age on a test copy. Verify the charts update and all sheets remain visible. Check the five-page print package and the optional annual detail on Letter and A4.
- Record the changes, sources, checks and remaining limits in the release record. Git keeps the prior version of the workbook.

## Remaining maintenance work

- Verify the Letter and A4 print layouts in Excel for Windows. The [release record](docs/release-notes/phase-3.md#published-workbook) records the completed Mac checks.
- Follow the [annual tax data review](docs/annual-tax-review.md) at the start of each tax year and when an agency publishes new figures used by the model. Record what was checked, even when the workbook needs no update.
- Have a second maintainer rehearse building, verifying and restoring a release using fictional data. Record the result and update any steps that need clarification.
