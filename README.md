# Financial Planning

Build a retirement plan in Excel and see whether your household can pay its bills and taxes through the age you choose, up to 125. [Download the workbook](https://raw.githubusercontent.com/Jon-275/financial-model/main/workbook/retirement-model.xlsx)

1. Save a private copy before entering personal figures. The workbook contains a fictional example.
2. Open the copy in desktop Excel. Follow **Start Here** and choose a scenario on **Inputs**.
3. Open **Summary** to see whether the plan covers its costs. If it falls short, use **Cash_Flow** to see what went unpaid that year.

All 15 workbook sheets are visible so you can inspect the calculations.

- Enter your assumptions and read the results. [Model guide](docs/model-guide.md)
- See how the plan handles different market returns. [Risk guide](docs/risk-analysis.md)
- Review tax loss and gain harvesting. [Portfolio guide](docs/portfolio-management.md)

## Contribute

Open a pull request into `main`. GitHub requires **Workbook checks**, **Greptile Review** and **Gemini Review** to pass. Run the local checks before submitting a change and use desktop Excel to verify a workbook release. [Test and release guide](tests/README.md) · [Branch rules](https://github.com/Jon-275/financial-model/rules/27295551)

### Keep private workbook copies out of Git

Run this command from the repository root once per clone to enable the workbook check before each commit.

```sh
git config --local core.hooksPath .githooks
```

Keep personal workbooks and reports in the ignored `private/` folder or outside the checkout. The published workbook uses the fictional scenario. Audit any change to the build template for private inputs, saved results and document properties before updating its checksum in `tools/check_workbook_commit.py`.

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

The script runs the release check and prints the result. Keep the file name `retirement-model.xlsx` so the Excel charts resolve their references. The [test and release guide](tests/README.md) shows how to verify a candidate before publication.

## Maintenance and history

- [Annual tax data review](docs/annual-tax-review.md)
- [Implementation plan and remaining work](PLAN.md)
- [Release records for Phase 3](docs/release-notes/phase-3.md), [Phase 2](docs/release-notes/phase-2.md) and [Phase 1](docs/release-notes/phase-1.md)
