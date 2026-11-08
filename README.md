# Retirement model

The Excel workbook projects household cash, taxes and account access through the end age you choose, up to age 125. Open **Start Here**, choose a scenario on **Inputs**, then read the funding answer at the top of **Summary**. [Open the workbook](workbook/retirement-model.xlsx)

**Summary** charts the portfolio balance and shows how the plan pays its bills around retirement. Use **Cash_Flow** to see what went unpaid and **Roth_Access** to see which savings were available.

The separate risk report recalculates spending, taxes and withdrawals under different market returns. Adjust its inputs on **Risk_Settings** and rerun the report after changing the workbook. [Portfolio risk analysis](docs/risk-analysis.md)

- [Model guide](docs/model-guide.md): how to enter assumptions and interpret results.
- [Implementation plan](PLAN.md): phase scope, release gates and progress.
- [Release records](docs/release-notes/phase-3.md) for Phase 3, [Phase 2](docs/release-notes/phase-2.md) and [Phase 1](docs/release-notes/phase-1.md).
- [Portfolio management](docs/portfolio-management.md): tax harvesting guidance separate from the forecast.

The workbook build is scripted from the pre-Phase 1 horizon workbook. Use the [release record](docs/release-notes/phase-1.md) to identify the build baseline and review the checks run. Keep personal inputs, statements and workbook copies on approved local storage. Use synthetic inputs for shared tests.
