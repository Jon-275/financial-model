# Desktop Excel verification and performance

Use the [enhancement ledger](enhancement-plan.md) to track the implementation and evidence. A Calc or Python pass supplies model evidence. Record repair, execution, cancellation, save/reopen and pagination from desktop Excel on each supported platform.

## D1. Build fixtures and check evidence

The desktop benchmark builder creates a disposable `.xlsm` from a workbook that matches both its published SHA-256 and the tracked workbook at the specified Git revision. It writes the declared inputs and records the resulting input hash in `fixture.json`. The released fictional snapshot is the seed-2026 reference for the standard example, including a 100-pair subset with independently ranked funding histories. For an edited case, the builder independently authors and packages a fresh pre-VBA `.xlsx` from the pinned source and template. It applies the same overrides and gives the empty result batch to `native_bootstrap`. [Benchmark builder](../../tests/desktop_benchmark.py) · [Bootstrap](../../tools/native_bootstrap.py)

| Fixture ID | Inputs or action | Evidence |
| --- | --- | --- |
| `example-95` | Released Example plan, opening age 35, end age 95, seed 2026. Run 100 and 1,000 pairs. | 233 and 2,033 outcomes respectively, including the 33 stresses. Independent reference, reports and timings. |
| `extended-125` | Same released inputs with `Inputs!D53=125`. Run 100 and 1,000 pairs with seed 2026. | Ending age, report and independently rebuilt reference for the extended scenario. |
| `full-0-125` | Explicit synthetic maximum-capacity case. Set fictional `Inputs!D8=0`, `D9=0` and `D53=125` on the release fixture and preserve the rest of the declared inputs. | Maximum model capacity, closure checks and an independently rebuilt reference. Preserve this workload apart from `extended-125`. |
| `stale-and-invalid` | Change an active input, swap scenarios, change dollar view, enter invalid age, restore the original input, and edit an inactive scenario. | Correct readiness and current/stale behavior. Inactive edits and the dollar view leave matching results eligible. |
| `recover-and-reopen` | Start from a complete batch. Exercise visible Cancel at a computation yield point. Inject a controlled error during staged capture, guarded writes or post-write verification, then save, reopen and retry. | Application settings, protection, prior batch and selection restored. Only a complete matching retry appears current. |
| `format-and-boundary` | Managed download and local save, macro denied, non-target workbook, supported locale, marked 1904 date-system fixture, Letter and A4. | Actual download trust behavior, base projection usable, comparable report values, six pages, no repair or unrelated workbook edits. |

Planning preflight recalculated `extended-125` and `full-0-125` in Calc using seed 2026 and a 100-pair setting. Both accepted the intended horizons and the risk settings. The independent deterministic ledger matched 8,640 and 12,000 annual values respectively within the existing one-cent tolerance. Native Excel execution and timing remain D1 and D3 evidence.

### Prepare and record a benchmark

Run from the repository root with the [maintainer dependencies](../../tests/README.md). Use the `.xlsm` and `SHA256SUMS.txt` from the same [published release](https://github.com/Jon-275/financial-model/releases/tag/v1.0.0), and fetch its Git tag if your clone lacks it. The fixture command accepts all three benchmark cases and either pair count. The independent reference for an edited horizon also requires LibreOffice Calc.

```sh
release_sha256=$(awk '$2=="retirement-model.xlsm" {print $1}' scratch/SHA256SUMS.txt)
fixture_dir=scratch/benchmarks/example-95-100
python3 tests/desktop_benchmark.py prepare scratch/retirement-model.xlsm "$fixture_dir" \
  --revision v1.0.0 --sha256 "$release_sha256" --fixture example-95 --pairs 100
python3 tests/desktop_benchmark.py reference "$fixture_dir"
cp "$fixture_dir/desktop-input.xlsm" "$fixture_dir/mac-01.xlsm"
```

Open only the run copy in desktop Excel and run **Run market risk**. Measure the run separately from file open and save. After **Risk results are current**, save the copy in `.xlsm` format. Keep the original `desktop-input.xlsm` in the directory. Record the observed timing and environment in a separate JSON file. Use `cold` for the first run in a fresh Excel process, or `warm` after the process has already run the workbook. Count repeats within each phase.

```json
{
  "platform": "mac",
  "phase": "warm",
  "repeat": 1,
  "seconds": 120.0,
  "excel_build": "16.115.26110837",
  "excel_channel": "not observed",
  "excel_bitness": 64,
  "os_build": "macOS 26.7.3 (25G342)",
  "architecture": "arm64",
  "machine": "MacBook Pro, Apple M5 Pro, 24 GB",
  "distribution": "local copy of fictional release",
  "macro_policy": "per-file prompt",
  "timing_method": "launch to visible completion"
}
```

The JSON shows the field format with an illustrative duration. Record the actual measurement and method. Run both commands below to save and replay the evidence.

```sh
python3 tests/desktop_evidence.py record "$fixture_dir" mac-01.xlsm scratch/timing.json \
  --output evidence-mac-01.json
python3 tests/desktop_evidence.py check "$fixture_dir/evidence-mac-01.json"
```

`DesktopEvidence/v1` binds the source, fixture, desktop input, reference and Excel capture by hash. The checker requires a fresh completed batch, replays the independent comparison, and rejects missing checks, mismatched inputs or a reused capture supplied as another repeat. Timing and environment fields describe the observed run. Keep evidence bundles local because Excel can add personal document metadata on save. Share sanitized summaries and hashes in PRs. The D3 release record also requires the hands-on compatibility, recovery and print checks below. [Evidence checker](../../tests/desktop_evidence.py)

### November 24, 2026 baseline

| Platform | Fixture | Observation | Reconciliation |
| --- | --- | --- | --- |
| MacBook Pro, Apple M5 Pro, 24 GB, macOS 26.7.3 (25G342), Excel 16.115.26110837 | `v1.0.0`, `example-95`, seed 2026, 100 pairs | One timed warm repeat, **135 seconds** from `RunMarketRisk` invocation to the workbook's saved batch timestamp. Excel showed current results and saved the `.xlsm`. | **233 outcomes**, with 80,454 compared fields and zero differences. The runner reported 46 verified fixed-result reuses. |
| Windows | Desktop Excel | Unverified | No Windows capture |

An earlier untimed run on the same Mac also reconciled. The timed value covers computation through the captured timestamp. Final verification and file save follow that timestamp. This is a single-device baseline observation. The repeated workload matrix, native cancellation and report exercises remain in R1 and D3.

Run the existing comparison from the repository root against saved captures. Replace the paths below with the fixture artifacts from the manifest. Use `--expect-reused` only when the independent case supplies that expected count.

```sh
python3 tests/native_excel_acceptance.py "$reference" "$excel_saved_copy"
python3 tests/native_excel_acceptance.py "$before" "$after" --check-restoration
```

## D3. Execute on supported desktop builds

Record Microsoft 365 desktop Excel on Apple Silicon and 64-bit Windows first. Name the supported update channels and versions before release. Add other architectures or Windows/macOS channels only when their own test evidence exists. Exercise the public download under normal policies. In E1b, repeat acceptance on E1a's signed file under its target distribution policy. Preserve application security settings in either lane. The released v1.0.0 workbook has prior Mac compatibility evidence. **Windows remains unverified.** [Maintainer walkthrough](../../tests/README.md#verify-a-changed-workbook-in-excel) · [Windows macro behavior](https://learn.microsoft.com/en-us/microsoft-365-apps/security/internet-macros-blocked)

1. Open a disposable copy through the supported user route, record any security or repair prompts, and verify the normal path to **Inputs** and **Summary**. Check protected formulas and editable inputs. Edit the end age and an assumption from an active scenario. Verify that the base Summary and Annual Detail update and saved risk becomes stale. Undo, switch scenarios, change only dollar view and verify status after each step.
2. Run on the worksheet button and verify progress, visible Cancel, failure status, retry, last-audit fields, save/reopen and the inspection selector. In a separate exercise with a second disposable workbook, record application-wide recalculation and verify that the runner makes no unintended edits or settings changes. Confirm the base projection works when macros are blocked. Repeat essential button and print checks interactively if app automation is used. Dispatch only reviewed commits with synthetic inputs to a licensed managed test device or private interactive host.
3. Compare the saved results with the independent workbook output. Check the fixed and flexible counts, annual ledgers, ranked examples, preserved report choices and all stress cases. Inspect raw desktop exports in both dollar views on Letter and A4, recording page count, layouts, extracted values and both decision pages. Use `example_pdf_acceptance.py` for the prepared public fictional release PDF with its metadata and structure requirements. [PDF comparison](../../tests/example_pdf_acceptance.py)

## R1 and R2. Measure and remove the dominant cost

Time the native macro independently of open, fixture preparation and save. For `example-95`, `extended-125` and `full-0-125`, collect 100- and 1,000-pair timings with one cold run and three warm repeats per supported platform on fixed hardware and settings. Close other workbooks during benchmark runs. Record the individual values and warm median. Report percentiles only with a sample that supports them. Capture peak memory where safely measurable and record progress and cancel latency separately. Budget for the existing long-running fixtures as an integration lane, with full grids on candidates touching execution and smaller smoke coverage on unrelated changes.

Instrument preflight/snapshot, trial preparation, recalculation, verified capture, fallback, ranking, publish and restore. Count executed and reused trials, fallback reasons, and all 33 stresses. Store timing and counters in test evidence outside plan identity and financial outputs. Measure the primary full calculation in `FinishTrialCalculation`, the year-ordered retry and workbook ranking separately. Inspect formula dependencies with `native_order_acceptance.py` before trying to narrow the calculation scope. [Native runner](../../tools/vba/RunMarketRisk.bas) · [Order acceptance](../../tests/native_order_acceptance.py) · [Excel calculation performance](https://learn.microsoft.com/en-us/office/vba/excel/concepts/excel-performance/excel-improving-calculation-performance)

Take one optimization per PR. If recalculation dominates, evaluate scoped calculation using the proven dependency order. Keep primary full calculation for the baseline and retain it as a recovery path only after proving the scoped path, with the existing ordered retry available. If marshaling dominates, batch reads and writes at the same verified boundary. If ranking dominates, preserve exact tie breakers and revalidate the three complete funding ledgers. Compare the baseline and candidate on each platform where the change will be enabled. A recurring fallback has an investigated cause and a regression case. Narrowed calculation may become the default only after the regular and fallback paths independently reconcile, including cancel and failure recovery. Retain the baseline runner strategy for rollback while the candidate is under evaluation.

Set release latency budgets after collecting raw timings. Initial comparison targets are at least 25% better warm median for full-range 1,000 pairs, no more than 10% regression in any other fixture, and visible cancel acknowledgement within five seconds at normal yield points. Track median 100-pair time as the exploration outcome and require an explicit budget on each supported device before claiming the new fast path. Record the prior 8- and 33-minute figures as historical context. Investigate and correct the cause of every new reconciliation or recovery failure. Keep the candidate Mac-only if Windows evidence is absent and label the missing tests in the release decision.
