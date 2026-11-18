# Retirement planner

Use Excel to see whether your household can pay its spending and taxes through the age you choose, up to 125. The published workbook includes a fictional household. Save a private copy before entering your own figures. [Download the current workbook](https://raw.githubusercontent.com/Jon-275/financial-model/b40af21ddaa65129bfe7ce180eb336243f5246f9/workbook/retirement-model.xlsx)

1. Open **Start Here** and save a private copy. Choose **Open Inputs**, select a scenario and enter your ages, income, savings and spending.
2. Open **Summary**. The headline tells you whether the base plan covers its costs. If it falls short, the report gives the first age and year that bills go unpaid.
3. Review the investment mix, spending policy and market results. Use the dollar selector on **Summary** to switch between today's purchasing power and projected amounts. Print **Summary** to keep a record of the report and your choices. [Model guide](docs/model-guide.md)

The next edition adds **Run market risk** in desktop Excel and a six-page Summary with stress tests and funding examples. The current download uses the earlier risk refresh process. Release testing uses Excel for Mac. [Download the workbook for testing](workbook/retirement-model.xlsm) · [Preview the six-page example report (PDF)](docs/examples/retirement-plan-example.pdf) · [Implementation plan](docs/model-improvement-plan.md)

## Understand the results

- **Cash flow and account access.** Follow income, spending, withdrawals and taxes through the annual schedules. All sheets are visible. [Workbook schedules](docs/workbook-schedules.md)
- **Market risk.** Read the success rate, balance range, spending cuts and funding examples together. [Market risk guide](docs/risk-analysis.md)
- **Portfolio management.** Review tax gain and loss harvesting as part of managing a taxable account. [Portfolio guide](docs/portfolio-management.md)

## Build the next edition

Use Python 3.11 or later and LibreOffice Calc. Put `soffice` on your path or set `SOFFICE_BIN` to its executable. From the repository root, run the following commands.

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

Store personal files outside the checkout or in the ignored `private/` folder. The template and shared workbook use fictional inputs and investment profiles. Use your GitHub private email for local commits and hosted merges. Replace clones made before the November 17, 2026 history cleanup. [History cleanup and clone replacement](docs/history-cleanup.md)

## Maintenance

- [Annual tax data review](docs/annual-tax-review.md)
- [Project plan](docs/model-improvement-plan.md)
- [Published financial corrections](docs/release-notes/financial-boundaries.md) and [workbook layout history](docs/release-notes/workbook-layout.md)
- Earlier releases cover [Phase 3](docs/release-notes/phase-3.md), [Phase 2](docs/release-notes/phase-2.md) and [Phase 1](docs/release-notes/phase-1.md).
