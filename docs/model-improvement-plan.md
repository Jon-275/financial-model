# Retirement model improvement plan

November 17, 2026

The published workbook includes the funding corrections from PRs 16 and 17. The macro-enabled workbook is a release candidate. It calculates market risk in Excel and presents the results in the six-page Summary. [Published financial fixes](https://github.com/Jon-275/financial-model/pull/17)

## Current status

| Phase | Completed | Still open |
| --- | --- | --- |
| 1. Privacy and history | Replaced the publication sources with synthetic data, cleaned the working clones and enabled GitHub email privacy. PR 17 uses the verified noreply address. | GitHub Support’s response, replacement of the colleague’s clone and a prior merge’s author metadata. The private recovery record identifies the affected commit. |
| 2. Annual funding and account access | PRs 16 and 17 corrected account depletion, Medicare surcharges and the first shortfall shown in the printed report. All 88 financial scenarios and four Mac Excel release cases passed for PR 17. The draft corrects cost basis after brokerage sales. Its independent ledger, three Mac Excel cases and all 93 financial scenarios passed. The protected calculation sheets passed all 93 cases and the final workbook passed three direct trials. | Publish the brokerage correction with the reviewed workbook. |
| 3. Calculation size and performance | Ordinary edits at maximum capacity took 1.69 seconds. The complete in-memory withdrawal search matched the worksheet across seven years and 231 proposed withdrawals in Mac Excel. Two blocks of 140 annual searches took 0.53–0.54 seconds in memory and 2.63–2.65 seconds on the worksheets. | Connect the annual search to the full multi-year cash flow, verify both spending policies and measure a complete market run. |
| 4. Workbook risk calculation | With year-by-year recovery on failed checks, the Mac completed 1,000-pair runs in 10 min 2 sec for the example and 27 min 30 sec through age 125. All 2,033 outcomes and three funding histories matched the independent calculation in each run. Separate Mac tests confirmed that scheduled bills enter each trial and can change the funding answer. | Reduce full-run times to two and five minutes. Verify restored inspection rows, the cancellation control and the released workbook on desktop Excel. |
| 5. Design and first use | The fictional workbook printed in six pages on Letter and A4 in Mac Excel. Each export included all 335 printable Summary cells. Switching dollar views updated Annual Detail and preserved the current risk results. An editorial walkthrough traced all 57 navigation links and led to clearer instructions for refreshing, saving a comparison report and entering the survivor benefit. | Observe a first-time reader completing their own plan. |
| 6. Release and enablement | The public build creates a macro-enabled workbook directly from the reviewed VBA source. The dated fictional PDF passed its financial, page and metadata checks. The original Mac-authored test project contained a local profile path and was excluded. | Final release and checks on its exact revision, including desktop Excel for Windows. Licensing remains pending at the owner’s request. |

The published file remains the macro-free workbook. The release candidate should replace it after the outstanding workbook and release checks pass. [Runner evidence](native-risk-implementation.md) · [Release checks](../tests/README.md)

## Purpose and release standard

Give a first-time user a clear answer to three questions.

1. Can my income and accessible savings pay the costs I entered?
2. How does the answer change with weaker markets, higher prices and a longer retirement?
3. Which assumption or decision should I review next?

The shared workbook must calculate its own projections, funding analysis and market risk in desktop Excel. Its report must explain the results in plain language and print directly from Summary. Every supporting sheet and formula must remain available for inspection.

Retain the existing GitHub repository. Address its sensitive template and history before promoting another release. A clean current download and a clean repository history require separate evidence.

## Evidence and scope

The source review is *Retirement Model v2: Review*, dated November 17, 2026, pages 1–8. Its findings N1–N12 are tracked below. Its recommendations are proposals evaluated against the model and the user's requirements.

The local and remote main matched the reviewed baseline during the initial investigation. The workbook checksum also matched the reviewed file. Only the PDF was supplied. Its listed scripts and stress-test workbooks remain to be reconstructed or obtained. The table below records the starting findings; implementation evidence follows at the end of this plan.

| Area inspected | Finding |
| --- | --- |
| Template and published workbook | The template retains the profile wording identified in N1. The distributed workbook has generic wording, but retains the same associated allocation weights. Both files have empty author properties and lack the inherited sensitivity-label and custom-property parts checked here. |
| Remote history | All 13 branch tips reference the same legacy template. Fetching those branches exposed 45 distinct commits and 12 distinct workbook blobs. The profile-wording scan flagged the template. These counts cover the fetched branch histories. Pull-request references, caches and copies held elsewhere need their own cleanup. |
| Completed branches | All 12 topic branches were merged into main. Each current tip matches its merged PR's final head, and there are no open PRs. Retire these branches during Phase 1. |
| Publication controls | The template check accepts a pinned file checksum. It does not independently validate the template's cell content. GitHub reports a public repository, zero forks and no license. Zero forks does not establish how many downloads or clones exist. |
| Savings benchmark | Planning Metrics B8 uses spending and healthcare divided by the withdrawal rate. It excludes Social Security, bridge years and the annual tax calculation. B18 compares savings with that benchmark. |
| Working-year funding | Cash_Flow J excludes Roth access while working. The Python annual engine has the same restriction. Unpaid spending has no roll-forward balance. Unpaid tax does carry forward. |
| Risk ownership | Python calculates the simulated paths. Excel displays the saved results and detects changed assumptions. A user currently needs Python and LibreOffice to refresh risk. |
| Calculation size | Funding_Calc allocates 125 annual solver groups. Targeting_Calc allocates 51 conversion trials for one selected year. These need different performance improvements. |
| Baseline verification | The current workbook passed the release acceptance check and all 35 risk tests during this investigation. Targeted review cases and their verification are recorded in the final checklist below. |

Read the current calculation explanations alongside this plan. They describe the starting implementation. [Workbook schedules](workbook-schedules.md) · [Market risk guide](risk-analysis.md) · [Release record](release-notes/workbook-layout.md)

## Priorities and ownership

| Phase | Priority | Accountable disciplines | Deliverable and exit condition |
| --- | --- | --- | --- |
| 1. Clean the publication sources and history | Immediate | Software engineering and release ownership | Synthetic source data, privacy checks of file content, completed branches deleted, automatic branch deletion enabled, rewritten retained history and a verified clean clone. |
| 2. Correct funding and clarify the benchmark | Required for the next shared model | Financial planning and financial engineering | Working-year Roth access, carried spending obligations and a consistent funding answer. |
| 3. Simplify the annual calculation | Required before workbook risk | Financial engineering and software engineering | A compact annual engine that preserves taxes, account access, discontinuities and the full age range. |
| 4. Calculate risk inside Excel | Required for the requested workbook | Financial engineering, software engineering and financial planning | Reproducible Excel-only risk refresh with inflation and longevity stresses, visible calculation detail and verified performance. |
| 5. Finish the experience and example | Required before public promotion | Product design, financial planning and editorial | A simple first-use route, coherent report, credible fictional example and accessible supporting sheets. |
| 6. Verify and release | Required before public promotion | Independent financial review, engineering and design review | Cross-platform evidence, reproducible public build, current guidance and all required PR checks. |

Phases 2 and 3 feed Phase 4. Report wireframes can proceed during the calculation work. Final report figures and screenshots follow the verified engine. Keep each phase's changes and evidence in a focused PR.

## 1. Clean the publication sources and history

### Replace the source of the exposure

1. Inventory the current files, every historical workbook path, branch and tag, open and closed pull requests, releases, Actions artifacts, caches and known local clones. Include the colleague's review clone. Keep raw personal strings and sensitive object identifiers in a private cleanup record.
2. Inspect both workbook packages beyond their visible cells. Cover shared strings, unused strings, cached formula results, defined names, comments, hidden rows and columns, links, connections, drawings and document properties. Review cached financial outputs as well as input cells.
3. Replace the template's household-derived labels, inputs and allocations with a documented synthetic fixture. Replace the corresponding allocations in the distributed example. Changing the profile names alone leaves the original numerical pattern.
4. Review the template's formulas and tax tables as functional source material. Preserve their verified behavior while replacing household data. Longer term, generate the clean template from explicit source tables and layout definitions.
5. Add a template content audit alongside the file checksum. Apply the same policy to the release workbook, generated examples, exported reports and anything uploaded by CI. Update checksums only after the content audit passes.
6. Add negative privacy tests. An altered profile label, injected unused shared string, personal scenario input or inherited metadata part must fail. Synthetic fixture amounts that happen to match a tax threshold must remain valid.

The current guard and metadata cleaner provide the starting points for this work. [Commit guard](../tools/check_workbook_commit.py) · [Metadata handling](../tools/workbook_metadata.py) · [Fictional inputs](../tools/example_inputs.py)

### Retire completed branches

1. After the inventory and restricted recovery copy are ready, record the names, current commit identifiers and merged PRs for the 12 completed topic branches.
2. Immediately before deletion, verify that each branch still matches its merged PR's final head. Preserve and review any work added since that check.
3. Delete the 12 confirmed completed branches from GitHub before rewriting the retained history. Keep main as the permanent branch. Retain future topic branches while their work is active.
4. Enable the repository setting `delete_branch_on_merge` so GitHub automatically deletes topic branches after future PRs merge.
5. Verify the remote branch list and setting. After the history cleanup, replace affected local clones and prune stale branch references. Check automatic deletion with the next merged PR.

The merged changes remain in main, and the PRs retain their review discussions. Continue with the history and cached-reference cleanup below to address the exposed template.

### Clean the existing repository

Use the four steps already discussed with the owner — rewrite locally, update GitHub, replace other clones and prevent reintroduction. Prepare reviewed clean artifacts outside the rewrite checkout. Preserve a restricted recovery copy under the owner's control and record how it will be retired.

Use `git-filter-repo` in a disposable mirror to remove affected workbook paths from retained history. Cover renamed paths and derived workbook snapshots identified by the audit. Reintroduce the reviewed synthetic template and release workbook. Rewrite main and any other branches or tags retained after the completed-branch cleanup. A squash merge of a cleanup commit leaves the earlier objects reachable. [GitHub history-removal guidance](https://docs.github.com/en/authentication/keeping-your-account-and-data-secure/removing-sensitive-data-from-a-repository)

Rehearse the rewrite, compare the intended source tree, record affected refs and coordinate open PRs. Arrange the narrowly scoped force update through the repository owner. Restore branch rules and the required Workbook, Greptile and Gemini checks afterward. Treat the force update as a separate controlled maintenance action.

Inspect pull-request refs and cached commit views after updating writable refs. Request GitHub Support assistance where the exposure qualifies. Coordinate replacement of known clones, including this project's earlier working copies and the colleague's clone. Record any remaining copies or cached references that cannot be verified. Previously downloaded copies remain outside the maintainer's control. [GitHub removal limits](https://docs.github.com/en/authentication/keeping-your-account-and-data-secure/removing-sensitive-data-from-a-repository#fully-removing-the-data-from-github)

### Evidence required to close Phase 1

- A new clone of the existing repository passes the content audit on current artifacts and every workbook reachable from retained branches and tags.
- The 12 completed branches are deleted. Any remaining topic branch has active work recorded in the cleanup inventory.
- Automatic branch deletion is enabled and its behavior is verified with the next merged PR. The protected main branch and required checks are verified.
- Historical paths and derived allocations have been reviewed. The inventory also accounts for pull-request refs, published assets and Actions artifacts.
- The clean source builds a clean example. The privacy test detects a deliberately contaminated copy.
- Known clones are replaced or quarantined against pushes. Required branch rules and reviews are restored and tested.
- The private incident record states the scope checked and any residual exposure. Public release notes describe the remediation without repeating personal values.

## 2. Correct funding and clarify the benchmark

### Give the funding result one meaning

Define success as paying all required costs and taxes in every modeled year through the selected end age. A shortfall of at least one cent makes the path unsuccessful, including a shortfall that is repaid later.

Show the first shortfall year, the amount unpaid and the account or cash limitation that caused it. Reconciliation establishes that money was accounted for. Funding establishes that bills were paid.

### Explain the savings benchmark

Keep the existing spending-divided-by-withdrawal-rate calculation as a labeled reference on Planning Metrics. Call it **Spending benchmark before income**. Explain which spending and healthcare budget it uses and that the annual funding result includes income and taxes.

Remove its use as an implied verdict on the plan. Summary should emphasize actual unpaid costs and the years needing portfolio withdrawals. Show the income source and start age beside the retirement funding table.

Add a compact bridge explanation covering retirement to Social Security, Social Security onward and a survivor period when selected. Follow the annual schedule for benefit timing, tax, medical costs and inflation. Subtracting the full Social Security amount from spending at retirement would miss the years before benefits begin.

Keep the comparison with the benchmark available on Planning Metrics with a plain explanation. Its presentation must let a reader understand how a plan can meet its modeled costs while sitting below that reference.

If a later release adds a required-capital estimate, solve it against the complete annual ledger with stated account balances and access assumptions. Include the bridge, benefit taxation and healthcare credits.

### Allow eligible Roth IRA money to fund working-year costs

Remove employment status as a blanket restriction on accessible Roth IRA funds. Use each owner's remaining regular contributions, conversion history, account age and distribution eligibility. Apply the Roth IRA ordering rules and reduce both the balance and the corresponding basis after a withdrawal. Workplace Roth requires its own plan-access rules. [IRS Roth IRA distribution guidance](https://www.irs.gov/publications/p590b)

Fund required costs before making elective IRA contributions or taxable deposits. Define the within-year timing once. Count each contribution once and apply the chosen timing convention to its availability. Reflect the change in Cash_Flow, Roth_Access, Projection, the conversion comparison and the independent Python calculation.

### Carry unpaid spending explicitly

Add an annual **Outstanding spending obligations** balance for required costs that remain unpaid. Its roll-forward is opening obligations plus newly unpaid required costs minus repayments. Repay prior tax first, then outstanding required spending and current required costs, before elective saving. Document the exact payment priority.

Use zero financing interest as an explicit initial assumption. Show outstanding tax and spending obligations beside assets and deduct them when reporting net resources. Keep outstanding balances, newly unpaid costs and cumulative newly unpaid costs distinct. A single unpaid bill carried across three years must remain one bill.

Record approved flexible spending reductions as lower required costs before funding. The first correction uses eligible assets and a visible obligation ledger. A future borrowing feature would require actual loan terms, payments and account rules.

Treat the reviewer's optional early traditional withdrawals as a later extension. Define an explicit policy for IRA and employer-plan access, income tax, additional taxes and exceptions before including them as a funding source.

### Account-access defaults

Replace the unexplained age-126 experience with a visible **Access unconfirmed** state for workplace Roth. Ask for the owner's plan distribution age and first account tax year when that account type is selected. Once confirmed, apply both plan access and tax qualification.

Continue to disclose the whole-year approximation that uses age 60 for age 59½. Test years on each side of that boundary. Add date-level treatment only with a specified annual timing convention and verified plan-specific rules.

### Evidence required to close Phase 2

- The review's working-year expense uses eligible IRA basis and identifies any remaining unpaid amount. It preserves owner-specific limits and later account balances.
- Required spending and tax obligations carry forward, reconcile and reduce net resources without double counting.
- The benchmark and funding result have distinct, understandable labels. Changing the benefit start age changes the bridge funding.
- Existing tests for conversion tax, Marketplace credit discontinuities, payroll, survivorship and account access continue to pass.

Primary implementation areas are the annual compiler, Roth rules, withdrawal solver, report formulas and independent risk calculation. [Annual model](../tools/model_plan.py) · [Withdrawal solver](../tools/funding_search.py) · [Report](../tools/plan_report.py) · [Python annual calculation](../tools/risk_engine/annual.py)

## 3. Simplify the annual calculation

The large calculation tables are the main performance dependency for native workbook risk. Address repeated work before adding simulation paths.

Separate the annual engine into visible steps that read from one set of active assumptions.

| Module | Inputs and responsibility | Output |
| --- | --- | --- |
| Calendar and prices | Opening year, both ages, end age, inflation and retirement dates | Active years, price levels and life stages |
| Income and costs | Wages, benefit timing, living costs, medical coverage and one-time costs | Cash income and required costs |
| Account access | Owner, account type, contribution basis, conversion dates and plan rules | Legally and operationally available amounts |
| Tax and funding | Cash needs, account access, gains, credits and withdrawal policy | Draws, conversions, taxes and unpaid costs |
| Account roll-forward | Deposits, draws, investment income, returns and obligations | Closing balances, basis and next-year state |
| Report measures | Completed annual results and trial results | Funding answer, cash-flow charts and risk measures |

Use labeled cells and bounded references for the calculation steps. Keep source assumptions in one place. Checks independently reconcile the completed schedules. Business calculations should derive their own required status from the responsible module.

Refactor the existing funding solver using its documented breakpoints and feasible intervals. Preserve testing around tax changes and Marketplace credit cliffs. Reuse year-specific inputs and evaluate tax-independent values once. Keep the full withdrawal search until an alternative reproduces the funding decision and the balances carried into later years. Even a small change near a tax or health-insurance threshold can alter a later decision. Keep convergence failures visible.

Treat Targeting_Calc separately. It compares conversion amounts in one chosen year. Reduce repeated context and unnecessary refinement across its 51 trials.

Retain interactive capacity through age 125. Changing the end age from 95 to 125 must work inside Excel. Rows, charts and print areas stop at the selected end age. Permanently building only the current horizon would break that requirement.

Record baseline and candidate formula counts, compressed and expanded size, memory, open time and recalculation time on named Mac and Windows test machines. Initial design targets are a workbook below 15 MB, ordinary edits below 2 seconds and a full deterministic recalculation below 10 seconds. These are proposed targets to validate with the prototype.

Phase 3 closes when the compact engine matches the verified annual cash, tax, access and balance results, handles discontinuities and changes the horizon through 125 without regeneration.

## 4. Calculate risk inside Excel

### Workbook ownership and refresh

The target deliverable is one self-contained workbook for supported desktop Excel on Mac and Windows. The recommended run control is an on-demand macro, which requires the `.xlsm` format. Its risk inputs, random-number method, annual calculation, trial results, macro source and report formulas must all be inspectable.

Use a seeded trial selector that reruns the same compact annual engine. A trial identifies its sequence of returns, inflation settings and spending policy. The native Data Table prototype captured ten complete paths correctly but its larger-capacity run did not meet the performance requirement. The recommended runner is a small VBA procedure that runs only the requested trials, collects their outputs and refreshes the report after a complete batch. Financial rules remain in visible worksheet formulas. [Runner design and evidence](native-risk-implementation.md)

Use a documented deterministic random-number method with stable trial and year identifiers. Avoid volatile random changes when the user opens the file or switches dollar views. Check the generator's distribution, independence and cross-asset correlation. Fixed and flexible spending must use the same market paths.

Test 1, 10, 100 and 1,000 paired paths through both a typical horizon and age 125. Target a 1,000-pair refresh within two minutes for the example and five minutes for the longest horizon on the documented reference machines. If the prototype misses the targets, revise the annual engine and trial calculation before release. Publishing a saved Python result would leave the workbook-only requirement open.

The next performance experiment is the calculation for one proposed withdrawal. Generate a small in-memory routine from the existing worksheet formulas, with the visible annual sheets as the reference. Replay the same proposed withdrawals for capped accounts, losses, changes in tax status and Marketplace credit thresholds. Compare funding decisions as well as amounts. Extend the experiment to a full market path only if the single-withdrawal results match and native Excel timing shows a material gain.

Make the refresh action visible in Excel and test the exact steps on both platforms. Record the full set of assumptions used for the results. Under the supported calculation setting, changed assumptions must mark saved results out of date before the user relies on them. Refresh must update the results, input record and run settings consistently.

Require automatic calculation for ordinary model edits. Make **Run market risk** the explicit simulation action, with progress and cancellation. Restore Excel's prior settings after success, cancellation and errors. Print the run context and assumptions used so a reader can assess a saved report. Test calculation settings, interrupted refreshes, save and reopen, and restoration of earlier assumptions. Document that VBA requires desktop Excel and allowed macros; Excel for the web cannot run it. [Microsoft compatibility guidance](https://support.microsoft.com/en-us/excel/work-with-vba-macros-in-excel-for-the-web)

The Python engine remains useful as an independent comparison during development. Runtime workbook users must be able to recalculate their own analysis offline in Excel.

### Inflation, returns and longevity

Keep nominal wage growth inclusive of inflation, as previously requested. A 10% annual raise applied to 80,000 produces 88,000 before tax in the next full year. Apply the entered total growth rate once.

Keep the current nominal investment-return convention explicit. Show the derived real return as `(1 + nominal return) / (1 + inflation) - 1`, after the selected fee treatment. Higher inflation changes purchasing power under a fixed nominal-return assumption. Automatically increasing expected asset returns when inflation rises would change that assumption.

Add required workbook stresses for 4% and 5% inflation and end ages 100 and 105. Include a combined higher-inflation, weaker-return case. Distinguish changes to the base plan from independently labeled stress cases.

Add a stochastic inflation option after the deterministic stresses reconcile. Document its source, long-run level, variation, persistence, bounds and assumed relationship with asset returns. Update spending, medical costs, benefit increases, indexed tax parameters and real-dollar reporting from the relevant path. Preserve statutory thresholds that are fixed in nominal dollars and document any benefit-indexing approximation.

Use fixed end ages for longevity comparisons. A mortality model would need its own demographic sources and household assumptions. The success rate must state the end age used.

Continue the correlated return model and early-loss tests with a concise statement of their assumptions. Cover severe return sequences and combined stresses explicitly. A calibrated regime or heavy-tail model is a later extension, after independent review of the data and validation method. That limitation belongs beside the risk interpretation.

### Report measures and audit trail

- Show the success percentage with its count and denominator. Explain that it is conditional on the selected assumptions.
- Label the chart shading **Middle 80% of simulated balances**. It is a pointwise range across paths.
- Show low, median and high funding examples as complete individual paths with their ranks defined. Keep their year-by-year cash flows together.
- Show unpaid spending, unpaid tax, outstanding obligations and first shortfall age. A later recovery must preserve the earlier failure.
- Explain when spending cuts occur, the reduction to the total budget and the years cuts last. Show the effect on healthcare and other protected costs.
- Record seed, trial count, model version, return convention, inflation mode, end age and the assumption comparison. Explain simulation sampling uncertainty in the supporting method note.

Phase 4 closes only when every path uses the verified tax, healthcare, account-access and obligation rules, and an Excel-only user can refresh the results and read them on Summary.

## 5. Finish the experience and example

### The report sequence

Summary contains the complete printable plan. A separate stress page gives each test room to show what changed, whether costs were covered and the size of any gap.

| Page | Reader's question | Content |
| --- | --- | --- |
| 1. Retirement outlook | Does this plan cover its costs? | Scenario, horizon, dollar view, funding result, conditional risk result, balance range and the most important caution |
| 2. Retirement cash flow | Where does the money come from? | Wages, Social Security and other income, savings used, taxes, healthcare, bridge years and outstanding costs |
| 3. Risk and next decisions | What could change the answer? | Low, median and high funding examples, spending flexibility and the next decisions to test |
| 4. Stress tests | What happens if conditions change? | Higher inflation, a longer retirement, combined pressures, medical budgets and the first unpaid costs |
| 5. Input decisions | What assumptions did I choose? | Ages, income, saving, return and inflation conventions, investment profiles and access settings |
| 6. Spending and account decisions | What commitments and restrictions did I enter? | Living costs, medical costs, one-time expenses, coverage, account rules and unresolved inputs |

Summary contains the findings, stress tests and decision record. Put the full annual timeline and market outcome tables immediately after Appendix. Keep the underlying annual schedules, trial calculations and source data with them.

Planning Metrics remains its own supporting tab. Sequence and format every added block within that tab or the report's established page structure. Reserve the report's side navigation for links and controls.

### Workbook design standard

- Use the established Arial typography, dark navy headings, muted teal accents, pale row shading and generous spacing consistently. Set a shared style specification in the generator.
- Use one title, a short context line and clear sections on each reader-facing tab. Put the outcome before the explanation.
- Keep body text at a readable 10–11 points or larger. Resolve crowding through layout and content changes.
- Align comparable amounts. Use consistent currency, percentages, ages, years, negative values and missing-value treatments. Label the opening year for today's dollars.
- Apply alternating shading to populated detail rows. Finish borders and formatting at the selected end age. Keep section gaps deliberate.
- Name income sources and account types in full on first use. Explain abbreviations beside the relevant decision.
- Use chart titles, axis units, restrained gridlines and direct legends. Color must reinforce labels so grayscale printouts remain understandable.
- Keep all worksheets visible. Protect formula cells against accidental edits while allowing formula inspection, selection, navigation and filters. Make protection easy to remove and leave formula-hiding disabled. [Microsoft worksheet protection guidance](https://support.microsoft.com/en-us/excel/protect-a-worksheet)
- Validate editable ages, amounts, rates, account types, coverage choices and allocation weights. Use plain messages for missing inputs and pasted invalid values. Preserve a valid zero and identify a missing required value.
- Follow the Summary dollar view throughout reader-facing financial outputs. Technical raw-data tables can retain their calculation units when the heading states them clearly.
- Put short instructions beside editable controls. A new user should see what to enter, where to find it and what changes after editing it.

### A short first-use route

Use Start Here for a short introduction. Readers enter household figures on Inputs, read and print the plan on Summary, and use the Appendix to inspect the calculations.

### A credible fictional example

Replace the household-derived profiles with documented synthetic, diversified working and retirement profiles. Explain the purpose of each account's mix and the treatment of any short-term reserve. Review weights, yields, returns, volatility and fees together.

Review pre-Medicare costs as gross premiums, expected credits and out-of-pocket costs. Include explicit tests at 18,000 and 25,000 per year for a couple, expressed in opening-year dollars. Use sourced coverage-year assumptions for a detailed fictional healthcare example and identify any illustrative budget values.

Confirm benefit amounts, RMD settings, filing status, survivor assumptions and account access for the example. Recompute its report and risk results after the assumptions change. Publish the new result with the assumptions that produced it.

## 6. Verify and release

### Public scope, licensing and build

State the scope near the start of the README, Start Here and the report footer.

> Educational planning estimate. US federal and California taxes. Joint filing and modeled survivor-single status. Review the assumptions and obtain qualified financial and tax advice before acting.

Explain the supported account types and material omissions beside that scope. Define compatible Excel versions based on actual Mac and Windows verification.

Licensing is pending at the owner's request. Apache 2.0 is under consideration for its explicit contributor patent terms. Do not apply a license until the owner selects it and confirms the rights to license the original work. State how workbook and documentation reuse is covered. Inventory dependency licenses and externally sourced tables. Add the selected license before describing the project as open source. [GitHub licensing guidance](https://docs.github.com/en/repositories/managing-your-repositorys-settings-and-features/customizing-your-repository/licensing-a-repository)

Replace the private authoring-runtime requirement with a documented public build path. Audit the existing OOXML packaging and use publicly installable, pinned dependencies to create the same formulas, charts, styles, validation and native Excel features. Build from the audited synthetic source in a clean environment. Verify source-to-workbook equivalence in CI.

Provide one canonical workbook and a clearly dated fictional sample PDF for preview. A personalized report can be exported from Summary. A lighter editable workbook is a later option once measured improvements justify a second maintained artifact.

Update the README, model guide, schedule explanations, risk guide, annual tax review, test guide, example screenshots and release notes together. State the supported tax years, data-review date and any estimates carried into later years.

### Financial and engineering checks

| Test area | Required evidence |
| --- | --- |
| Privacy | Current artifacts, historical reachable workbook objects, unused strings, properties, derived allocations and clean-clone verification |
| Expense funding | Working-year expense with zero, partial and sufficient IRA basis, multiple owners, unavailable workplace Roth and a remaining obligation |
| Obligation ledger | Repayment priority, multiple unpaid years, later recovery, no duplicate bill and no elective saving ahead of required payments |
| Benefit timing | Retirement before, at and after benefit start, survivor change, zero benefits and an inflation change |
| Tax and access | Existing annual fixtures plus credit cliffs, gain realization, payroll by earner, conversion seasoning, required distributions and owner eligibility |
| Dollar views | Identical underlying cash flows, correct conversion to opening-year dollars and consistent report and annual-detail displays |
| Risk | Exact seeded repeatability, zero-volatility comparison, every-year accounting, fixed/flexible pairing, valid correlations, inflation stresses and horizon changes through 125 |
| Risk freshness | Every active assumption group, calculation settings, interrupted refresh, save and reopen, restored assumptions and changed trial count |
| Native workbook | Offline Excel-only refresh, no repair prompt, no external calculation dependency and visible trial detail |
| Performance | Measured open, edit, full recalculation and risk-refresh times, memory, formula count and file sizes on both reference platforms |
| Public build | Clean installation, documented commands, deterministic financial outputs and preservation of workbook features |
| Existing safeguards | Misplaced one-time expense entries, invalid ages, high-earner Roth phase-out, fee treatment and independent HSA allocation continue to work |
| Input validation | Required blanks, valid zero values, pasted invalid amounts, unsupported choices and allocation totals produce clear local messages |

Recreate the colleague's stress cases from the supplied descriptions and record the exact input changes. Keep the executable cases with the project's tests. Preserve both positive and negative cases.

Use cent-level reconciliation and boundary fixtures for the annual ledger. Compare the same trial draws between Excel and the independent Python model. Statistical tests need defined tolerances for means, variation and correlation.

### Editorial, design and platform checks

Ask a first-time reader to enter a fictional plan, explain the funding answer, identify the source of income at benefit start, change an assumption, refresh risk and print the result. Record where the reader hesitates and revise those steps.

Preview every report page on US Letter and A4 in desktop Excel for Mac and Windows. Inspect both dollar views, long scenario names, shortfalls, missing inputs, age 95 and age 125. Check chart legends, row cutoffs, page breaks, grayscale legibility and formula protection.

Use Workbook checks, Greptile Review and Gemini Review on each relevant PR. Resolve actionable review findings and record the reason for any disputed finding. A release needs completed financial, privacy and design evidence as well as passing automated checks.

## Review coverage and completion record

| Review point | Planned response | Completion evidence |
| --- | --- | --- |
| N1. Template and history | Phase 1 replaces household-derived source data, strengthens the guard and cleans retained history in the existing repository | Clean clone, historical-object audit and residual-exposure record |
| N2. Savings target | Phase 2 labels the gross benchmark and uses timed income and annual costs for the funding answer | Benefit-timing fixtures and a reader explanation |
| N3. Risk limits | Phase 4 adds inflation and longevity stresses, variable inflation and clear interpretation | Native Excel results and documented model assumptions |
| N4. Working-year expenses | Phase 2 uses eligible IRA funds and carries required spending obligations | Annual funding, owner access and obligation tests |
| N5. Returns and inflation | Phase 4 explains nominal and real returns and preserves inclusive wage growth | Inflation and wage fixtures with no double counting |
| N6. Investment example | Phases 1 and 5 replace derived profiles with documented diversified synthetic profiles | Allocation audit and recomputed example |
| N7. Healthcare example | Phase 5 adds sourced coverage assumptions and higher-budget stresses | Gross premium, credit and out-of-pocket reconciliation |
| N8. External runtime | Phases 4 and 6 deliver Excel risk calculation and a public build path | Fresh-machine build and offline workbook refresh |
| N9. Public scope | Phase 6 adds visible tax and filing scope, limitations and the selected license | Consistent workbook, README and license review |
| N10. Access defaults | Phase 2 makes workplace access confirmation explicit and documents the whole-year convention | Unknown-access and age-boundary tests |
| N11. Workbook size | Phase 3 addresses annual solver duplication and selected-year trial cost while retaining capacity to 125 | Before/after measurements and interactive horizon test |
| N12. Protection and Windows QA | Phases 5 and 6 protect editable boundaries and finish platform checks | Formula visibility and Mac/Windows print evidence |
| First use | Start Here links to Inputs, Summary and the Appendix | First-time-reader test |
| Lighter sharing copy | Phase 6 supplies a sample PDF and measures the canonical workbook before adding another editable edition | Preview matches the released workbook |

### Plan preparation completed

- [x] Read and visually inspect all eight review pages.
- [x] Match the reviewed commit and workbook checksum to the current repository.
- [x] Inspect the current template, publication guard, formulas, risk engine and build process.
- [x] Inventory all remote branch tips and scan fetched workbook history for the reported profile wording.
- [x] Verify the current release check and 35 risk tests.
- [x] Recalculate three targeted cases in LibreOffice using temporary copies with inactive solver rows trimmed. Confirm first shortfall in 2057 at 5% inflation, zero working-year Roth access for the one-time expense, and a 91-row projection ending at age 125. Explicit assertions and annual reconciliation passed.
- [x] Validate all 91 local documentation links, coverage of N1–N12, phase dependencies and wording. Keep raw personal values and the review PDF outside the repository.

### Implementation tracking

- [x] Delete the 12 confirmed completed remote branches and record the result.
- [x] Enable automatic branch deletion after merges and verify the setting.
- [x] Verify automatic branch deletion with the publication cleanup PR.
- [x] Rewrite retained history, restore required branch rules and verify a fresh clone with only the two approved workbook blobs.
- [x] Replace the owner's checkout and disable normal pushes from 12 earlier task clones. Preserve the owner's earlier files and planning edits in restricted recovery storage.
- [x] Submit the authorized request to GitHub Support for cached views and old pull-request references. Keep its contents in restricted recovery storage.
- [ ] Obtain GitHub Support's disposition of cached PR references and confirmation that the colleague replaced the review clone.
- [ ] Phase 1 complete with privacy and history evidence.
- [x] Phase 2 complete with financial and access evidence. PR #16 passed Workbook, Greptile and Gemini checks and merged.
- [ ] Phase 3 complete with equivalent results and performance measurements.
- [ ] Phase 4 complete with Excel-only risk refresh and integrated report.
- [ ] Phase 5 complete with example, design and first-use evidence.
- [ ] Phase 6 complete with public build, documentation and platform evidence.

### Implementation evidence

The publication cleanup passed nine privacy tests, all three required PR checks, 81 original financial scenarios and 35 risk tests before merging. Retained branch history was then rewritten in the existing repository. The current source tree was identical before and after the rewrite. A new clone passed the content audit. [Cleanup record](history-cleanup.md)

The funding update adds working-year access to eligible Roth funds, carried spending obligations and visible repayment detail. It labels the spending benchmark before income and explains benefit timing on Summary. All 85 financial scenarios passed on the compact engine, including five new funding cases. The SEPP case was updated to carry bills and test repayment after age 60. Representative scenarios were also recalculated without trimming the calculation sheets. [Funding and solver evidence](release-notes/funding-and-solver.md)

An earlier prototype reduced the calculation workbook from about 24 MB to 6.1 MB by replacing repeated source formulas with local links. Excel for Mac recalculated the full model through age 125. The current 13.8 MB draft also includes the on-demand runner, the market results and three complete funding examples. Windows performance and print verification remain open.

The source-built VBA runner completed a desktop Excel for Mac run and the report printed in six pages. The full-size release candidate holds 2,033 calculated outcomes. The current status table records desktop timing and the remaining release checks. [Runner design and verification](native-risk-implementation.md)
