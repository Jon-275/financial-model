# Desktop Excel verification and performance

Use the [enhancement ledger](enhancement-plan.md) to track the implementation and evidence. A Calc or Python pass supplies model evidence. Record repair, execution, cancellation, save/reopen and pagination from desktop Excel on each supported platform.

## D1. Build fixtures and check evidence

The desktop benchmark builder creates a disposable `.xlsm` from a workbook that matches both its published SHA-256 and the tracked workbook at the specified Git revision. It writes the declared inputs and records the resulting input hash and resolved commit in `fixture.json`. Replay uses that commit even if the original branch name moves. The released fictional snapshot is the seed-2026 reference for the standard example, including a 100-pair subset with independently ranked funding histories. For an edited case, the builder independently authors and packages a fresh pre-VBA `.xlsx` from the pinned source and template. It applies the same overrides and gives the empty result batch to `native_bootstrap`. [Benchmark builder](../../tests/desktop_benchmark.py) · [Bootstrap](../../tools/native_bootstrap.py)

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

`DesktopEvidence/v1` binds the source, fixture, desktop input, reference and Excel capture by hash. The checker requires a fresh batch and verifies the saved workbook's VBA source and library IDs against the pinned input. It replays the independent comparison and rejects missing checks, mismatched inputs, and duplicate saved batch timestamps for the same candidate, fixture and machine, even when the two save files have different hashes. Timing and environment fields describe the observed run. Keep evidence bundles local because Excel can add personal document metadata on save. Share sanitized summaries and hashes in PRs. The D3 release record also requires the hands-on compatibility, recovery and print checks below. [Evidence checker](../../tests/desktop_evidence.py)

### November 24, 2026 baseline

| Platform | Fixture | Observation | Reconciliation |
| --- | --- | --- | --- |
| MacBook Pro, Apple M5 Pro, 24 GB, macOS 26.7.3 (25G342), Excel 16.115.26110837 | `v1.0.0`, `example-95`, seed 2026, 100 pairs | One timed warm repeat, **135 seconds** from `RunMarketRisk` invocation to the workbook's saved batch timestamp. Excel showed current results and saved the `.xlsm`. | **233 outcomes**, with 80,454 compared fields and zero differences. The runner reported 46 verified fixed-result reuses. |
| Windows | Desktop Excel | Unverified | No Windows capture |

An earlier untimed run on the same Mac also reconciled. The timed value covers computation through the captured timestamp. Final verification and file save follow that timestamp. This is a single-device baseline observation. The repeated workload matrix, native cancellation and report exercises remain in R1 and D3.

### November 25, 2026 Mac matrix

The Mac workload matrix has one cold and three warm runs for each scenario and pair count. Each cold run starts in a fresh Excel process. The warm repeats use the same process and a prior run of the same workload. The first standard-case warm repeat opened a fresh copy of the fixture in that process. Subsequent warm repeats keep the workbook open with the previous published batch. Inputs and the runner are fixed, and each run starts with only the benchmark workbook open. Two extended-horizon 100-pair attempts were discarded when unrelated workbook prompts interrupted capture. Cold retries start in a fresh Excel process.

Timing uses a monotonic clock from starting the synchronous Excel command for `RunMarketRisk` through its return, including final verification and restoration. Workbook opening, saving and independent replay take place outside that interval. The November 24 observation ends at the batch timestamp inside the macro and stays separate. Peak RSS is the maximum of two-second process samples where collected. Raw captures stay local.

**Mac matrix coverage is 16 of 24 independently reconciled runs.** Windows remains unverified.

The fictional `v1.0.0` workbook at `6573e1e4bd86eb99bbce6678c6fdf9f575c2b681` uses seed 2026. The test Mac was a MacBook Pro with Apple M5 Pro, 24 GB RAM, macOS 26.7.3 (25G342), and 64-bit Excel 16.115.26110837. The Excel update channel was unobserved.

| Fixture | Pairs | Cold (s) | Warm 1 (s) | Warm 2 (s) | Warm 3 (s) | Warm median (s) |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| `example-95` | 100 | 143.965 | 130.199 | 161.823 | 135.364 | 135.364 |
| `example-95` | 1,000 | — | — | — | — | — |
| `extended-125` | 100 | 201.229 | 204.682 | 188.917 | 180.185 | 188.917 |
| `extended-125` | 1,000 | — | — | — | — | — |
| `full-0-125` | 100 | 274.019 | 298.783 | 271.708 | 259.099 | 271.708 |
| `full-0-125` | 1,000 | 2,063.465 | 2,415.751 | 2,508.750 | 2,549.974 | 2,508.750 |

Each capture below passed freshness, pinned runner identity and independent reconciliation. The field counts include trial, balance, ranked path ID and funding ledger values.

| Fixture and pairs | Repeat | Outcomes | Fields checked | Peak sampled RSS (MiB) | Capture SHA-256 |
| --- | --- | ---: | ---: | ---: | --- |
| `example-95`, 100 | cold 1 | 233 | 80,454 | — | `6bfdf0872eb40dc75bfb89c31ff33f7883b282645f1f018ff5fa74c839a1952e` |
| `example-95`, 100 | warm 1 | 233 | 80,454 | 993 | `24d3e6bc620116e1486c32f8cbfb894c9ad686eb45dc9c47650277bc559dd916` |
| `example-95`, 100 | warm 2 | 233 | 80,454 | 1,213 | `e3f3e9145429d54e36d9e66049645ff2d2185b9a4f8d50f07608a3847a25718f` |
| `example-95`, 100 | warm 3 | 233 | 80,454 | 970 | `568380daa0aa0821641bbe1702ae0af1418339c738620e3ca7fb7817a76b19cb` |
| `extended-125`, 100 | cold 1 | 233 | 80,454 | 1,237 | `1c31e9b0daba118a17267a9f6007a2829571fcac763e7ffa6bf2156726b7bc80` |
| `extended-125`, 100 | warm 1 | 233 | 80,454 | 1,558 | `3499c2cd62e27e863d7f8f2eab70cd2d477bc6b2117a71fbe0f49c5e4eb6db68` |
| `extended-125`, 100 | warm 2 | 233 | 80,454 | 1,502 | `8d69ea56a4ba028219a9a4b73b8bbcae222db167aaa350728ac0539eb053c853` |
| `extended-125`, 100 | warm 3 | 233 | 80,454 | 1,608 | `84555000918c69878d8e3ad38d0c6c6199e748b47cb7c30cedaa2abab3c0d8f7` |
| `full-0-125`, 100 | cold 1 | 233 | 80,454 | 1,446 | `3187571f73ac3d36bb8da2534a3a98b48f9da526b0d655092757b7cbdfbd5c8f` |
| `full-0-125`, 100 | warm 1 | 233 | 80,454 | 1,677 | `821e6f93786d585abca2613214485786f54ce63e769a706e975274c35ad6b097` |
| `full-0-125`, 100 | warm 2 | 233 | 80,454 | 1,697 | `60dade3311bccbf995bcb97b97002471a8e21e17cc846a790e4e0671bb68dce4` |
| `full-0-125`, 100 | warm 3 | 233 | 80,454 | 1,719 | `c1749b4d45121ab4fa0b3e77c314df5a4b9eea2322612c6132ad2ca740aac2a3` |
| `full-0-125`, 1,000 | cold 1 | 2,033 | 588,054 | 3,112 | `92c8739b1b2555966e035eb634062e4479465a10604ca0bd7a9143dc467163aa` |
| `full-0-125`, 1,000 | warm 1 | 2,033 | 588,054 | 3,571 | `0630091ddeaea2facfb911410ac0a4ea420a55d460a8691501a6a843951e70ca` |
| `full-0-125`, 1,000 | warm 2 | 2,033 | 588,054 | 3,763 | `3d992a6210495f71a841bb715198d5f4a43209a891cb2e798bf1c716de03bc52` |
| `full-0-125`, 1,000 | warm 3 | 2,033 | 588,054 | 3,496 | `ce2110230e8f83cd9e4769e9d5a938c789157b709173bf11e69ae9821a2906a3` |

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
