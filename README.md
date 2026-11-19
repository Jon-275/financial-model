# Retirement planner

Use Excel to see whether household income and accessible savings can pay your spending and taxes through the age you choose, up to 125. The examples use fictional figures. The model estimates US federal and California taxes for a couple filing jointly and, when selected, a surviving spouse filing as single. [Scope and limits](docs/model-guide.md#items-to-estimate-separately)

The retirement planner is available for testing in desktop Excel for Mac. [Download the workbook (.xlsm)](workbook/retirement-model.xlsm?raw=true) · [Preview the six-page example report (PDF)](docs/examples/retirement-plan-example.pdf)

## Use the macro-enabled workbook

1. Open the `.xlsm` file in desktop Excel for Mac. Save a private copy in the same format. On **Start Here**, choose **Open Inputs**, select a scenario and enter your household figures.
2. Open **Summary** to see whether the base plan covers its costs. With Excel set to Automatic, the projection updates as you edit. If bills go unpaid, the report shows the first age and year this happens.
3. Review the investment mix and spending assumptions. When you’re ready, choose **Run market risk** on **Summary** with macros enabled. You can make all your edits before starting this calculation.
4. When Summary says **Risk results are current**, choose today’s or future dollars and print **Summary**. The six pages include your input and spending decisions. [First-use guide](docs/model-guide.md#run-your-first-plan) · [How to read the market results](docs/risk-analysis.md#read-the-summary)

## Understand the results

- **Cash flow and account access.** Follow income, spending, withdrawals and taxes through the annual schedules. All sheets are visible. [Workbook schedules](docs/workbook-schedules.md)
- **Market risk.** Read the success rate, balance range, spending cuts and funding examples together. [Market risk guide](docs/risk-analysis.md)
- **Portfolio management.** Review tax gain and loss harvesting as part of managing a taxable account. [Portfolio guide](docs/portfolio-management.md)

## Build the workbook

Maintainers use `workbook/build-template.xlsx` to build the planner. Use Python 3.11 or later and LibreOffice Calc. Put `soffice` on your path or set `SOFFICE_BIN` to its executable. From the repository root, run the following commands.

```sh
python3 -m pip install -r tools/requirements.txt
mkdir -p scratch
build_dir=$(mktemp -d scratch/build.XXXXXXXX)
bash tools/build_model.sh workbook/build-template.xlsx "$build_dir/retirement-model.xlsm"
```

The builder enters fictional inputs, calculates 1,000 paired market paths and 33 stress cases, and packages the reviewed VBA source into the workbook. Users run their own scenarios in desktop Excel. [Build and release checks](tests/README.md)

## Contribute

Open a pull request into `main`. **Workbook checks**, **Greptile Review** and **Gemini Review** must pass, and all review threads must be resolved. [Test and release guide](tests/README.md) · [Branch rules](https://github.com/Jon-275/financial-model/rules/27295551)

Enable the workbook privacy check once per clone.

```sh
git config --local core.hooksPath .githooks
```

Keep personal workbooks outside the checkout or in the ignored `private/` folder. Use fictional inputs in shared files and a GitHub noreply address for commits and merges. [Protect personal files](tests/README.md#protect-personal-files)

## Maintenance

- [Annual tax data review](docs/annual-tax-review.md)
- [Release readiness and open work](docs/model-improvement-plan.md)
