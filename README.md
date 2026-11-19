# Retirement planner

This model is for educational purposes. Use it to explore whether your income and the savings you can withdraw will cover spending and taxes through the age you choose, up to 125. The examples use fictional figures. The model estimates US federal and California taxes for a couple filing jointly and, when selected, a surviving spouse filing as single. [Scope and limits](docs/model-guide.md#items-to-estimate-separately)

The first release is a macro-enabled workbook for desktop Excel for Mac. [Download the retirement planner (.xlsm)](workbook/retirement-model.xlsm?raw=true)

## Use the macro-enabled workbook

1. Open the `.xlsm` file in desktop Excel for Mac. Save a private copy in the same format. On **Start Here**, choose **Open Inputs**, select a scenario and enter your household figures.
2. Open **Summary** to see whether the base plan covers its costs. With Excel set to Automatic, the projection updates as you edit. If bills go unpaid, the report shows the first age and year this happens.
3. Review the investment mix and spending assumptions. When you’re ready, choose **Run market risk** on **Summary** with macros enabled. You can make all your edits before starting this calculation.
4. When Summary says **Risk results are current**, choose today’s or future dollars and print **Summary**. The six pages include your input and spending decisions. [First-use guide](docs/model-guide.md#run-your-first-plan) · [How to read the market results](docs/risk-analysis.md#read-the-summary)

## Understand the results

- **Cash flow and account access.** Follow income, spending, withdrawals and taxes through the annual schedules. All sheets are visible. [Workbook schedules](docs/workbook-schedules.md)
- **Market risk.** Read the success rate, balance range, spending cuts and funding examples together. [Market risk guide](docs/risk-analysis.md)
- **Portfolio management.** Review tax gain and loss harvesting as part of managing a taxable account. [Portfolio guide](docs/portfolio-management.md)

## For maintainers

To change the workbook, follow the [build and testing instructions](tests/README.md). Review tax figures annually and whenever a source agency updates them. [Annual tax review](docs/annual-tax-review.md)

## License

The project’s original workbook, source code and documentation are licensed under the [Apache License, Version 2.0](LICENSE).
