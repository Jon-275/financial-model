# Retirement model implementation plan

The model tests whether household income and available savings can pay spending and taxes each year. The plan runs through the age selected on **Inputs**, up to 125.

## Next planned release

The [model improvement plan](docs/model-improvement-plan.md) addresses the November 17 review. It starts with the sensitive build template and repository history, then corrects expense funding, simplifies the calculation, brings market risk calculation into Excel and completes the design and publication checks. Implementation is open.

Phase 1 includes deleting the 12 completed topic branches, enabling automatic branch deletion after future merges and verifying the remaining history and local clones.

## Completed

| Phase | Changes |
| --- | --- |
| [1. Cash flow and account access](docs/release-notes/phase-1.md) | The annual projection pays taxes first, then required spending. It reports the first year with unpaid costs and tracks each owner’s Roth access. |
| [2. Taxes, healthcare and conversions](docs/release-notes/phase-2.md) | The model adds healthcare by person, tests tax and Marketplace credit changes as withdrawals rise, and compares Roth conversion amounts for one year. |
| [3. Risk analysis and workbook usability](docs/release-notes/phase-3.md) | The risk report recalculates cash, taxes and account access for each set of market returns. **Start Here**, **Inputs** and **Summary** guide readers from setup to a five-page printable plan. |
| [4. Market risk and funding decisions](docs/market-risk-plan.md) | **Summary** includes market risk, funding examples, and a choice of today’s or future dollars. The five-page plan prints with the selected view. |
| [5. Workbook layout and calculation review](docs/release-notes/workbook-layout.md) | All five pages print from **Summary**. Supporting figures sit on **Planning Metrics**. **Annual Detail** follows the dollar view and selected end age. The Roth conversion summary and tax tables come before **Appendix**, and the detailed comparison and withdrawal trials sit behind it. The HSA projection uses its own investment mix. |

Keep tax loss and gain harvesting in the [portfolio guide](docs/portfolio-management.md). The model includes the tax effect of investment sales needed to pay bills.

## Release checks

Use the [test and release guide](tests/README.md) for commands and Excel steps. Use fictional inputs for all shared workbooks and test results.

- Compare the candidate with the prior release. Record changes to inputs, defaults, formulas or charts.
- Check each scenario and confirm the projection stops at the selected age. Include age 125, an invalid age and a shortfall late in the plan.
- Exercise account access before and after eligibility dates, changes in tax when money is withdrawn, and a year when available savings run out.
- Reconcile cash and account balances for all active years. Run the risk report against the candidate and verify it matches the workbook when market returns follow the plan.
- In desktop Excel, check for repair prompts, recalculate from cleared formula results, and change the end age on a test copy. Check that the chart and row shading follow the end age. Confirm the conversion charts use **Conversion Detail** and that the solver filters show the selected draws. Preview all five pages from **Summary** on Letter and A4. Check **Annual Detail** in both dollar views.
- Record the changes, sources, checks and remaining limits in the release record. Git keeps the prior version of the workbook.

## Remaining maintenance work

- Verify the Letter and A4 print layouts in Excel for Windows. The [current release record](docs/release-notes/workbook-layout.md#verification) records the Mac checks.
- Follow the [annual tax data review](docs/annual-tax-review.md) at the start of each tax year and when an agency publishes new figures used by the model. Record what was checked, even when the workbook needs no update.
- Have a second maintainer rehearse building, verifying and restoring a release using fictional data. Record the result and update any steps that need clarification.
