# Financial Planning

Build a retirement plan in Excel and see whether your household can pay its bills and taxes through the age you choose, up to 125. [Download the workbook](https://raw.githubusercontent.com/Jon-275/financial-model/main/workbook/retirement-model.xlsx)

1. Save a private copy before entering personal figures. The workbook contains a fictional example.
2. Open the copy in desktop Excel. Follow **Start Here** and choose a scenario on **Inputs**.
3. Open **Summary** to see whether the plan covers its costs and how it fares in different markets. The workbook opens with market risk results for the fictional example. [Print your plan with the input decisions](docs/model-guide.md#print-your-plan)

**Appendix** links to the supporting schedules and calculations. The decision tools sit before it. All workbook sheets are visible.

- Enter your assumptions and print the five-page plan from Summary. [Model guide](docs/model-guide.md)
- Follow the income, account and tax calculations. [Workbook schedules](docs/workbook-schedules.md)
- Refresh the market risk results after changing your assumptions. Open the refreshed Excel plan to see the success rate and funding examples on **Summary**. [Market risk guide](docs/risk-analysis.md)
- Review tax loss and gain harvesting. [Portfolio guide](docs/portfolio-management.md)

## Contribute

Open a pull request into `main`. GitHub requires **Workbook checks**, **Greptile Review** and **Gemini Review** to pass. Resolve every review thread before merging. Run the local checks before submitting a change and use desktop Excel to verify a workbook release. [Test and release guide](tests/README.md) · [Branch rules](https://github.com/Jon-275/financial-model/rules/27295551)

### Keep private workbook copies out of Git

Run this command from the repository root once per clone to enable the workbook check before each commit.

```sh
git config --local core.hooksPath .githooks
```

Keep personal workbooks and reports in the ignored `private/` folder or outside the checkout. The published workbook uses the fictional scenario and its saved market risk results. The commit check accepts the reviewed example snapshot by checksum. [Review a rebuilt workbook before publishing](tests/README.md#microsoft-excel-release-check)

## Rebuild the workbook

The build starts from `workbook/build-template.xlsx` and writes a new workbook. Run the commands from the repository root with Python 3.11 or later, Node.js, LibreOffice Calc, and the Codex spreadsheet runtime containing `@oai/artifact-tool`.

Put `soffice` on your path or set `SOFFICE_BIN` to its executable. Set `ARTIFACT_NODE_MODULES` to the runtime’s absolute `node_modules` path. If the runtime has a separate Node executable, set `NODE_BIN` to that executable.

```sh
python3 -m pip install -r tools/requirements.txt
export ARTIFACT_NODE_MODULES=/absolute/path/to/runtime/node_modules
mkdir -p scratch
build_dir=$(mktemp -d scratch/build.XXXXXXXX)
bash tools/build_model.sh workbook/build-template.xlsx "$build_dir/retirement-model.xlsx"
```

The script runs 1,000 market paths for the example, saves the results on the workbook, and runs the release check. Keep the file name `retirement-model.xlsx` so the Excel charts resolve their references. The [test and release guide](tests/README.md) shows how to verify a candidate before publication.

## Maintenance and history

- [Annual tax data review](docs/annual-tax-review.md)
- [Implementation plan and remaining work](PLAN.md)
- [Release records for Phase 3](docs/release-notes/phase-3.md), [Phase 2](docs/release-notes/phase-2.md) and [Phase 1](docs/release-notes/phase-1.md)
