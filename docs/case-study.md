# Building an inspectable retirement planner

A project by [Jon](https://github.com/Jon-275). [Try the planner](getting-started.md) · [See a fictional report](examples/retirement-plan-example.pdf) · [Browse the source](https://github.com/Jon-275/financial-model)

## The problem

A retirement forecast needs to explain more than a final account balance. Income can arrive after bills are due, some savings may be inaccessible and changing withdrawals can change taxes and healthcare costs. A useful model needs to expose those interactions and make its assumptions reviewable.

## The implementation

This repository contains the workbook-generation pipeline, Excel schedules, VBA market-risk runner, documentation and acceptance checks. The deliverable is a macro-enabled workbook for desktop Excel for Mac, with a six-page Summary and annual calculation schedules.

The project demonstrates financial-model implementation, Python tooling, VBA orchestration, report design and automated validation. The source and commit history show the implementation work; this case study does not claim independently measured user adoption or business impact.

## Decisions and tradeoffs

| Decision | Why it matters | Evidence |
| --- | --- | --- |
| Keep calculations in visible Excel schedules | Users can follow income, tax, account access and unpaid costs instead of trusting a single score. | [Schedule guide](workbook-schedules.md), [formula sources](../tools/formula_sources.py) |
| Generate the workbook with Python | Layout, formulas and packaging are maintained as source alongside the distributable workbook. | [Builder](../tools/author_model.py), [build workflow](../tools/build_model.sh) |
| Orchestrate trials with VBA | The market analysis uses the workbook's annual calculations in desktop Excel. The tradeoff is macro permissions and runtime. | [Runner design](macro-runner.md), [VBA source](../tools/vba/RunMarketRisk.bas) |
| Compare spending policies under paired conditions | Fixed and flexible policies face the same sampled returns and inflation. | [Risk guide](risk-analysis.md), [scenario code](../tools/risk_engine/scenario.py) |
| Distinguish current results from stale or incomplete runs | Editing assumptions, cancelling or failing a run must not present previous figures as current evidence. | [Runner checks](../tests/test_macro_runner_contract.py), [report checks](../tests/test_report_structure.py) |
| Publish fictional inputs and inspect workbook contents | Household information belongs in private working files. | [Privacy check](../tools/check_workbook_commit.py), [privacy tests](../tests/test_workbook_privacy.py) |

## How validation is structured

The repository includes model acceptance checks, scenario fixtures, financial and report tests, LibreOffice recalculation checks and a documented native Excel verification process. Native Excel checks cover opening without repair, recalculation, run cancellation and recovery, and comparison with an independent reference. These are different forms of evidence: a CI pass does not by itself establish native Excel behavior.

See the [validation instructions](../tests/README.md) and [CI runs](https://github.com/Jon-275/financial-model/actions/workflows/workbook-checks.yml) for the procedures and actual run results. No new financial validation is claimed by a documentation or website update.

## What the experience delivers

Users can compare three input scenarios, inspect annual cash flows, run market and stress tests, and save a report that records inputs alongside results. The report below comes from the repository's fictional example.

![Fictional retirement outlook with projected portfolio ranges](images/summary.png)

## Constraints and next improvements

The verified release target is desktop Excel for Mac. The model's tax scope is US federal and California, with joint-filer and selected survivor assumptions. Simulations use specified distributions and assumptions rather than replaying historical crises. Long runs can take minutes. [Full limits](model-guide.md#items-to-estimate-separately)

A useful next evaluation would observe first-time users opening the example, changing one assumption and explaining the resulting report. Completion rate, misunderstandings and time to first useful comparison would provide evidence for further improvements. These outcomes have not yet been measured.
