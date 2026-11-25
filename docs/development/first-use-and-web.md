# First-use study, workbook experience and web decision

Use a fictional release and retain the workbook's **Start Here → Inputs → Summary** route. The first result is the deterministic funding answer and its first-shortfall explanation. Market risk is a separate step that depends on a verified run. [First-use guide](../getting-started.md) · [Risk guide](../risk-analysis.md)

## D2. Run the study

Recruit five people who have never used the planner. Include at least two occasional spreadsheet users and an assistive-technology user when available. The project owner can nominate participants. Track the count invited, unavailable and unable to access eligible Excel even when they cannot join the five sessions. Use supported Mac desktop Excel for the initial study and repeat the critical flow on Windows when available. Use a moderator, a note taker and a staff dry run outside the sample. Obtain consent for coded notes before the session. Keep names and contact details separate from coded observations. Use fictional figures and leave recording off by default. Store raw notes in the approved research location, delete them after 30 days unless organizational policy requires a different retention period, and publish aggregate timing and anonymized friction in the PR.

**Invitation draft.** Would you be willing to spend 20 minutes trying a fictional retirement workbook? We’re testing how easy it is to compare scenarios in Excel. You’ll use made-up figures, and we won’t ask for personal financial information. Please tell us whether you can use desktop Excel and whether you’d like accessibility accommodations. Participation is optional.

Prepare a checksum-verified copy of the [fictional release workbook](https://github.com/Jon-275/financial-model/releases/tag/v1.0.0) for each session. Leave **Example plan** active. Generate a separate shortfall workbook from the matching repository checkout. Confirm both workbooks and the answer key in desktop Excel before the sessions. Use the shortfall workbook for the deterministic interpretation probe and the release `.xlsm` for the main and risk tasks. [Fixture runner](../../tests/xlsx_acceptance.py)

```sh
python3 tests/xlsx_acceptance.py workbook/retirement-model.xlsm tests/phase2-fixtures.json \
  --case later_cash_repays_prior_spending_before_deposits --compact --jobs 1 \
  --keep scratch/first-use
```

Moderator introduction. “We’re testing this workbook. It uses fictional figures. Tell me what you’re looking for as you go. Use any help you find in the workbook or its guides. Please don’t enter any of your own information.”

Start an acquisition timer on the product page and a separate primary timer when Excel finishes opening the file. Read the prompts as written and log any hints. Score the unaided outcome at five minutes. Continue to eight minutes for diagnosis when useful. Record security prompts as observed friction and preserve the participant's macro settings. Distinguish actions they performed from steps they described.

| Exercise | Prompt | Verified v1.0.0 answer key |
| --- | --- | --- |
| Primary 1 | Find the age the person plans to retire, the age the plan covers, and whether spending and taxes are paid through then. | Example plan retires at 60 and covers costs through age 95. |
| Primary 2 | Compare with retirement at 65. Does either plan fall short before the end, and where did you find the answer? | The later plan retires at 65 and covers costs to age 95. Both answers come from the base projection. Saved risk for the original scenario becomes stale after switching. |
| Interpretation | In this second workbook, what first goes unpaid and when? Would the ending portfolio alone tell you? | $100 of spending first goes unpaid at age 55 in 2027. Later cash repays it. The $1,300 ending portfolio and the first shortfall describe different points in the plan. |
| Risk and print | Show how you would make a quick market comparison and prepare a more thorough report. How would you know the results are current? | Use 100 and 1,000 pairs respectively, preserve seed for comparison, find current status and saved count, choose dollar view and preview. Stop before a lengthy run if needed. |

The first answer comes from the tagged workbook. The later-retirement and shortfall answers were checked using the repository's Calc fixture path. Repeat the answer-key check in desktop Excel during the dry run and after each change to a workbook or fixture.

| Code | Excel/OS build and input mode | Acquisition seconds | First correct answer seconds | Both answers seconds | Correct with workbook evidence | Hints | Shortfall and stale-risk understanding | Report mode and dollar view | Blocker and observation |
| --- | --- | ---: | ---: | ---: | --- | --- | --- | --- | --- |
| P01 | | | | | | | | | |
| P02 | | | | | | | | | |
| P03 | | | | | | | | | |
| P04 | | | | | | | | | |
| P05 | | | | | | | | | |

Record Excel experience and any assistive input in the session notes. The primary target is four of five answering both parts correctly within five minutes with no moderator hints. Self-serve help is allowed. Report completion counts, median among finishers, non-completions and repeated blockers. Treat a material misreading of stale or missing risk or the first shortfall as a release blocker. Five sessions support qualitative decisions. Retest on a fresh cohort using the same tasks and record a discrete accessibility audit.

## U1. Update the workbook and its guidance

1. **Find and compare.** Change the highest-impact observed friction first. Make **Try the example** and return to **Summary** clear beside the scenario picker if those steps slow people down. Preserve input validation, `Inputs!B4`, existing scenario values, editable and protected regions, freeze panes and the macro-free route. Work through `start_here.py`, `workbook_layout.py`, `decision_layout.py`, and the decision-layout tests. [Layout tests](../../tests/test_decision_layout.py)
2. **Run and recover.** Keep `Risk_Settings!B29/B30` the authoritative count and seed, with Explore (100), Final (1,000) and Custom. Use one validated setting for every entry point. Prefer accessible native controls with a manual cell route. If cross-platform preset wiring fails, label the current mode from the same setting. Launch only on user action. Give current, missing, stale, invalid, incomplete, running, cancelled and failed results distinct messages and appropriate next actions. Show elapsed time, counts and Cancel at safe yield points. Derive duration estimates from verified timings. [Risk settings and native execution](model-contracts.md#market-paths-execution-and-publication)
3. **Read and share.** Use a clear visual hierarchy for base funding, first shortfall, and simulated success. Show conditional meaning, actual completed count and seed, date, scope and model version. Label Explore in the saved report. Keep the annual bands and representative paths distinct and provide a readable table for chart values. Today's dollars means opening-year purchasing power. Future dollars use each year's prices. Keep the same representative paths across dollar views. An unrun or stale risk section can print with a clear status while the current base projection remains legible. Preserve the six pages and both decision records. Change `plan_report.py`, `risk_presentation.py`, `input_record.py`, `report_layout.py`, and `dashboard_package.py` only where needed. Verify formulas, charts, PDF data and protection. [Report structure tests](../../tests/test_report_structure.py) · [Fictional PDF checks](../../tests/example_pdf_acceptance.py)
4. **Test accessibility and deployment.** Exercise keyboard-only navigation, edit controls, screen readers on supported platforms, visible focus, high zoom, contrast and label/order consistency. Run Excel Accessibility Checker and record exceptions. Verify protected sheets expose meaningful labels and chart alternatives. Inspect both paper formats and dollar views. The UX lead resolves barriers that prevent completion of the tested tasks. Offer an accessible HTML companion for the fictional report if native PDF export cannot meet the accessibility target, regardless of the W1 decision. [Microsoft Excel accessibility guidance](https://support.microsoft.com/en-us/accessibility/excel/accessibility-best-practices-with-excel-spreadsheets)

Update README compatibility, the short walkthrough, model and risk interpretation, the input reference, the maintainer release notes and the download site in the applicable release. Use actual UI labels. Publish the verified workbook, synthetic example report, screenshots, checksums and compatibility/source notes together. Give desktop IT a short deployment note covering managed macro policy, verified build/channel, failure recovery and rollback. Keep maintainer implementation detail outside the general-audience workflow. [README](../../README.md) · [Publishing](../publishing.md)

## W1 and W2. Decide what the web should solve

Track Excel acquisition and access among all invitees, including people screened out. Score managed macro policy against the separate risk task. Track accessibility barriers on the exact task affected. Consider a web pilot when at least two invitees are blocked in the same relevant workflow and support or IT evidence corroborates it. Weigh the simplest remedy, ongoing maintenance and accessibility. Document the decision, denominators, target users, affected task and funding. Keep the existing download site when there is insufficient evidence.

**W1. Fictional read-only companion.** Generate accessible results and comparison tables from versioned, reconciled fictional fixtures and approved release assets. Serve them through the static site. Include no personal inputs or plan storage. Target WCAG 2.2 AA through manual and automated checks, including keyboard, status announcements, chart equivalents and reflow at 320 CSS pixels. Match the downloaded release and keep an asset rollback to the current static site. [Current site](https://jon-275.github.io/financial-model/) · [WCAG 2.2](https://www.w3.org/TR/WCAG22/)

**W2. Private offline calculator.** If W1 and further research justify editing, put calculation in a browser worker that consumes the [versioned model contracts](model-contracts.md). Start with fictional plans. The UI sends an immutable validated run spec with a request ID, receives bounded progress and a verified result, and ignores late responses after cancellation or replacement. Provide keyboard and screen-reader equivalents and use local explicit export/import only after testing size, schema, unsupported values and version migration. Reuse one approved computation interface and policy dataset. Verify the base ledger, RNG vectors, both spending policies, state/year and scores against Excel and the independent reference before exposing real inputs. Measure time to first projection, native comparison, and cancel response on the slowest supported test device. Prove no input network transmission or default persistence, including error logging, offline behavior and any service worker updates.

For an enterprise-hosted variant, treat identity, access control and retention as a separate approval before enabling employee data. SSO can govern access to the app while the worker keeps plan inputs on the device. Server computation or retained plans require an explicit data inventory, tenant isolation, encryption, retention and deletion policy, safe logs, quotas, incident/rollback controls and release evidence. Keep any service status, latency or error telemetry free of household assumptions and user plan identifiers. Desktop Excel and current workbook release checks stay supported when a web pilot exists.
