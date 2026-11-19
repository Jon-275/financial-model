# Retirement model improvement plan

Updated November 18, 2026

The published download is the macro-free workbook. The next edition is a macro-enabled Excel workbook with a six-page Summary. It uses fictional inputs and is available for testing on Mac. [Workbook and guides](../README.md)

## Workbook ownership and refresh

Keep the financial calculations in visible worksheet formulas. With Excel set to Automatic, changing a plan input updates the base projection. After making edits, choose **Run market risk** on **Summary** to refresh the market scenarios. Use 100 paired paths for a quick comparison or 1,000 for the final report. Both include 33 stress cases.

In Excel for Mac, the worksheet reports took 7 min 52 sec for ages 35–95 and 33 min 10 sec for ages 0–125. The in-memory prototype took 33 min 52 sec and 60 min 12 sec for the same reports. The financial results and saved Summary charts matched. The owner chose the worksheet version. [Full report comparison](native-risk-implementation.md#full-report-comparison)

## Work before publication

| Area | What remains |
| --- | --- |
| Calculation time | The targets for 1,000 pairs are two minutes for ages 35–95 and five minutes through age 125. Measure where the worksheet run spends time before choosing further changes. |
| First use | The owner will complete the walkthrough below. Record any instructions, controls or report figures that need clarification before sharing the fictional example with colleagues. |
| Privacy | The owner confirmed that GitHub removed the affected workbook history. One retained commit still has an earlier author email. Resolve that item privately. [Cleanup record](history-cleanup.md) |
| Licensing | The owner’s license decision is pending. Confirm rights to the code, workbook and documentation and review third-party notices before applying a license. |
| Release | Verify the final revision in desktop Excel for Mac, complete the release checklist, match the sample PDF to the workbook and update the published download. Automated cancellation passed. The owner deferred the manual Cancel-button check. [Test and release guide](../tests/README.md#microsoft-excel-release-check) |

## First-use walkthrough

1. Save a private copy of the fictional workbook as an Excel Macro-Enabled Workbook (`.xlsm`). On **Start Here**, choose **Open Inputs**, select a scenario and enter the household figures.
2. Open **Summary** and check whether the base plan covers its costs. Review any shortfall, the sources of retirement income and the decisions recorded for the plan.
3. Review the investment mix and risk assumptions. Choose **Run market risk**, wait for **Risk results are current**, and compare the funding examples and stress tests.
4. Switch between today’s and future dollars. Preview the six pages from **Summary**, including the two decision pages. Print **Annual Detail** separately if needed.
5. Note what was hard to find, enter or interpret. After the walkthrough, share a fresh fictional copy so each colleague can enter their own figures privately. [First-use guide](model-guide.md#run-your-first-plan)

## Report and workbook standards

Keep the six printable pages on **Summary**. Show the outlook and retirement cash flow first, then market results and stresses, followed by the two decision pages. Keep **Planning Metrics** separate and detailed calculations after **Appendix**. All sheets and formulas should be visible for inspection. Label dollar units, follow the Summary dollar view on report tabs, and check that chart labels and shaded ranges can be read in grayscale. [Workbook schedules](workbook-schedules.md) · [Print checks](../tests/README.md#microsoft-excel-release-check)

## Completed work

- The annual funding model carries unpaid bills and taxes forward, tracks access to Roth accounts and accounts for brokerage cost basis. All 93 financial scenarios passed. [Funding corrections](release-notes/funding-and-solver.md) · [Account and tax checks](release-notes/financial-boundaries.md)
- Both 1,000-pair worksheet runs produced 2,033 outcomes and three complete ranked funding histories. Saved-file tests checked balances, the selected trial and recovery after interruption. [Calculation evidence](native-risk-implementation.md)
- Excel for Mac printed the six-page Summary on US Letter and A4. Annual Detail followed the selected dollar view. All worksheets remain visible for inspection. [Print instructions](model-guide.md#print-your-plan)
- The build packages the reviewed VBA source and fictional data. The workbook, Greptile and Gemini checks passed for the worksheet version. [PR 23](https://github.com/Jon-275/financial-model/pull/23)

## Ongoing maintenance

Review tax data annually and when a source agency changes a figure used by the model. Have a second maintainer rehearse the build, verification and release recovery with fictional data. [Annual tax review](annual-tax-review.md) · [Build and release checks](../tests/README.md)

The original design standards, review findings and phase-by-phase acceptance criteria are in the [plan archive](archive/README.md).
