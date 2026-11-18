# Retirement model implementation plan

The model tests whether household income and available savings can pay spending and taxes each year. The plan runs through the age selected on **Inputs**, up to 125.

## Next planned release

The release candidate calculates market risk in Excel for Mac. Its six-page **Summary** covers the results and the decisions behind them. The improvement plan tracks the remaining release checks, including privacy and performance. [Model improvement plan](docs/model-improvement-plan.md)

The owner will work through the guide on a separate copy, enter a plan, refresh market risk and preview the report. After the walkthrough, colleagues will receive a fresh copy of the fictional example and can enter their own figures privately. Licensing remains pending. [First-use guide](docs/model-guide.md#run-your-first-plan)

## Completed

| Phase | Changes |
| --- | --- |
| [1. Cash flow and account access](docs/release-notes/phase-1.md) | The annual projection pays taxes first, then required spending. It reports the first year with unpaid costs and tracks each owner’s Roth access. |
| [2. Taxes, healthcare and conversions](docs/release-notes/phase-2.md) | The model adds healthcare by person, tests tax and Marketplace credit changes as withdrawals rise, and compares Roth conversion amounts for one year. |
| [3. Risk analysis and workbook usability](docs/release-notes/phase-3.md) | The risk analysis added cash, tax and account-access checks for each set of market returns. **Start Here**, **Inputs** and **Summary** linked setup with the printable plan. |
| [4. Market risk and funding decisions](docs/market-risk-plan.md) | **Summary** added market risk, funding examples and a choice of today’s or future dollars for the printed plan. |
| [5. Workbook layout and calculation review](docs/release-notes/workbook-layout.md) | The printable pages moved to **Summary**, with supporting figures on **Planning Metrics**. **Annual Detail** follows the dollar view and selected end age. The Roth conversion summary and tax tables come before **Appendix**, and the detailed comparison and withdrawal trials sit behind it. The HSA projection uses its own investment mix. |

Keep tax loss and gain harvesting in the [portfolio guide](docs/portfolio-management.md). The model includes the tax effect of investment sales needed to pay bills.

## Release checks

Use fictional inputs for release testing. Run the workbook checks in Excel for Mac. [Test and release guide](tests/README.md)

- Compare the candidate with the prior release. Record changes to inputs, defaults, formulas or charts.
- Check each scenario and confirm the projection stops at the selected age. Include age 125, an invalid age and a shortfall late in the plan.
- Exercise account access before and after eligibility dates, changes in tax when money is withdrawn, and a year when available savings run out.
- Reconcile cash and account balances for all active years. Run market risk in the workbook and compare a no-shock path with the base projection.
- In Excel for Mac, check for repair prompts, recalculate from cleared formula results, and change the end age on a test copy. Check that the chart and row shading follow the end age. Confirm the conversion charts use **Conversion Detail** and that the solver filters show the selected draws. Preview all six pages from **Summary** on Letter and A4. Check **Annual Detail** in both dollar views.
- Record the changes, sources, checks and remaining limits in the release record. Git keeps the prior version of the workbook.

## Remaining maintenance work

- Follow the [annual tax data review](docs/annual-tax-review.md) at the start of each tax year and when an agency publishes new figures used by the model. Record what was checked, even when the workbook needs no update.
- Have a second maintainer rehearse building, verifying and restoring a release using fictional data. Record the result and update any steps that need clarification.
