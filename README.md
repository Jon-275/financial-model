# Retirement planner

Test whether income and savings can cover spending and taxes through retirement. Compare retirement ages, budgets and market returns.

The model estimates US federal and California taxes for couples, with an option to model a surviving spouse. It is for educational use and includes fictional examples. [Scope and limits](docs/model-guide.md#items-to-estimate-separately)

**[Download the workbook](https://github.com/Jon-275/financial-model/releases/latest/download/retirement-model.xlsm)** · **[View the website](https://jon-275.github.io/financial-model/)** · **[Preview the report](docs/examples/retirement-plan-example.pdf)**

<a href="docs/examples/retirement-plan-example.pdf"><img src="docs/images/summary.png" alt="Fictional retirement report showing the outlook and simulated portfolio balances" width="680"></a>

*Page 1 of the six-page example report. These figures are not an expected outcome for your household.*

## Before you start

- Use **desktop Excel for Mac** and keep the `.xlsm` file format. Windows and other spreadsheet apps are not verified release targets.
- **Run market risk** requires macros. Follow your organization's macro policy. Excel for the web can show saved results but cannot run the analysis.
- Keep calculation set to **Automatic**. The base projection updates when you edit inputs.
- Use **100 paired paths** while trying changes or **1,000** for a final report. The documented final example took about 8 minutes; the full age range took 33 minutes. Your Mac and plan affect runtime.

You do not need Python or a copy of the source code. Save a private copy before entering personal figures.

## Test different scenarios

Start with the [fictional example](docs/getting-started.md). On **Inputs**, select later retirement or lower spending, then compare **Summary**. Save a report before changing a plan. Run market risk again before comparing market results.

## Follow the workbook

Follow the tabs from left to right:

| Tab | What to do |
| --- | --- |
| Start Here | Save a copy and open Inputs. |
| Inputs | Select a scenario and enter your figures. |
| Summary | Read the six-page report and save it as a PDF. |
| Planning Metrics | Review the supporting figures. |
| Asset_Allocation | Review investment mixes and returns. |
| Income_Targeting | Compare Roth conversions. |
| Risk_Settings | Set up market tests, then run them from Summary. |
| Tax_Tables | Review tax values and their effective years. |
| Appendix | Open the annual schedules and calculation checks. |

[Workbook guide](docs/model-guide.md) · [Input reference](docs/input-reference.md) · [Market risk guide](docs/risk-analysis.md) · [All documentation](docs/README.md)

## Help

[Ask a question or report a problem](https://github.com/Jon-275/financial-model/issues/new/choose). Include your Excel version, steps and exact message. Use fictional figures; do not attach a personal workbook or report.

## About the project

Created by [Jon](https://github.com/Jon-275). Python builds the workbook, Excel formulas calculate the plan and a VBA macro runs the market tests. [How I built it](docs/case-study.md)

For development, see the [build and test guide](tests/README.md). [Tax data review](docs/annual-tax-review.md) · [Website and release maintenance](docs/publishing.md) · [Releases](https://github.com/Jon-275/financial-model/releases)

The original workbook, source code and documentation use the [Apache 2.0 license](LICENSE).
