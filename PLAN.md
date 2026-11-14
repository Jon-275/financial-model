# Retirement model implementation plan

The first release makes the workbook’s funding result answer a specific question: can the household pay each year’s obligations from cash and accounts it can draw from? The annual projection, funding checks and retirement bridge share the same access rules. The model supports projections through age 125 and stops at the scenario’s selected end age.

## Phase 1: Reconcile cash, tax and account access

1. **Validate the scenario.** `Inputs` supplies the timeline, account ownership and the ages and cash shares for outside income. Existing wages, spending and opening balances remain as entered. Users confirm the new fields in rows 86–98.
2. **Build annual obligations and assess access.** `Cash_Flow` combines spendable income and required spending. `Roth_Access` keeps each owner’s Roth IRA and workplace Roth separate and tracks IRA principal and conversion aging. `Projection` supplies opening balances and other eligible sources. The model uses one traditional account owner and requires verified, qualified workplace Roth access.
3. **Fund draws and tax together.** `Funding_Calc` uses a bounded search over eligible draws. It evaluates the tax on each candidate and skips an optional conversion when the year cannot support it. Actual tax is paid before required spending. Unpaid tax is carried forward and deducted from after-tax net worth. Actual Roth IRA savings use cash left after obligations.
4. **Roll forward and reconcile.** The selected draws, actual savings, transfers and growth produce the next year’s opening balances. `Checks` independently flags reconciliation errors, excess Roth draws and calculation issues. Required shortfalls are reported even if later balances recover.
5. **Publish results.** `Summary` shows the first shortfall and the Roth funds included in the bridge. Risk headlines require a funded base plan. Where available, the risk views vary returns while holding the base cash-flow, tax and access schedule fixed.

## Release gates

| Gate | Evidence to retain |
| --- | --- |
| Input migration | Compare pre- and post-release active scenario values. Record any fields with changed meaning or defaults. Verify switching to each scenario recalculates from its own inputs. |
| Annual cash and account checks | For every active year, cash sources and shortfall reconcile to cash uses. Account opening balances, transfers, withdrawals and growth reconcile to closing balances. Optional contributions do not create cash. |
| Roth and bridge checks | Cover a Roth IRA contribution, a conversion before and after its eligibility date, and an employer Roth balance with and without plan access. Confirm the bridge and actual withdrawals use the same assumptions. |
| Tax funding checks | Verify the first and last retirement years; include a case in which a taxable withdrawal raises the tax that must be funded, and a case in which eligible funds are exhausted. |
| Horizon and regression | Exercise a short plan, the active plan, age 125 and invalid ages. Check for hidden dependencies on old age limits. Verify a late funding failure reaches Summary. Compare unchanged formulas, sheets, charts and scenario values against the baseline. |
| Workbook review | Open the release in Microsoft Excel and check for repair prompts. Recalculate a copy from blank caches and run the accounting checks. Change the end age on a test copy and confirm the charts refresh. Confirm every worksheet is visible and the workbook contains only fictional inputs. Record the Excel version and results. |
| Release record | Keep a versioned change note, assumption and source references, tests run and any deferred work. Retain the prior workbook in Git. Use synthetic examples in shared test fixtures. |

## Progress

| Work item | State | Evidence |
| --- | --- | --- |
| Unify the projection horizon | Implemented before Phase 1 | Workbook supports age 125 and stops at the selected age. The workbook and portfolio chart were checked through age 125 in Excel. |
| Annual cash and income policy | Implemented | `Inputs` rows 86–89 and annual `Cash_Flow` ledger. |
| Current-year tax funding | Implemented | Bounded draw search and explicit tax liability roll-forward. |
| Roth account and bridge logic | Implemented | Owner-specific Roth schedules, eligible draw limits and bridge links. |
| Reconciliation and release verification | See release record | [Phase 1 release record](docs/release-notes/phase-1.md) records verification and engine limits. |
| Healthcare by person, HSA and Marketplace credits | Implemented in Phase 2 | Detailed coverage is optional. The annual funding search tests both sides of credit thresholds. |
| Payroll, AMT and taxable brokerage sales | Implemented in Phase 2 | Wages are split by earner. Taxable sales release basis and carry federal and California losses separately. |
| Roth conversion comparison | Implemented in Phase 2 | The comparison uses the selected year's filing status and reruns the same funding calculation. |
| Phase 2 validation and source data | See release record | [Phase 2 release record](docs/release-notes/phase-2.md) records source years, checks and remaining limits. |
| Per-path portfolio risk | Implemented in Phase 3 | Seeded market paths use separate working and retirement allocations. Each path reruns funding, taxes and account access. The spending rule protects fixed costs. |
| Workbook layout and reports | Implemented in Phase 3 | Start Here, grouped Inputs and a funding-first Summary. Risk_Settings supplies the companion risk report. |
| Phase 3 verification | See release record | [Phase 3 release record](docs/release-notes/phase-3.md) records workbook parity, seed checks and the calculation cache fix. |

## Next phase

| Phase | Scope | Exit test |
| --- | --- | --- |
| 4. Release operation | Establish annual tax data review, change ownership, independent release review and rollback rehearsal. Phase 1 already provides a scripted build and test cases. | A second operator can build, validate, review and roll back a release using synthetic cases without changing private assumptions. |

Keep deliberate tax harvesting in [Portfolio management](docs/portfolio-management.md). The model should still handle the tax effect of sales required for spending.
