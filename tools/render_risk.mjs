import fs from "node:fs/promises";
import { dirname, join } from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const [input, output, previewDir] = process.argv.slice(2);
if (!input || !output) throw new Error("Usage: node render_risk.mjs REPORT.json NEW_REPORT.xlsx [PREVIEW_DIR]");
const data = JSON.parse(await fs.readFile(input, "utf8"));
const wb = Workbook.create();
const navy = "#173149";
const teal = "#00776E";
const ink = "#253A4E";
const slate = "#4D6377";
const mist = "#EDF4F8";
const mint = "#E1F4EF";
const white = "#FFFFFF";
const money = '"$"#,##0;("$"#,##0);"$"0';
const percent = "0.0%";
const count = "#,##0";

function sheet(name, widths) {
  const page = wb.worksheets.add(name);
  page.showGridLines = false;
  page.tabColor = name === "Overview" ? teal : navy;
  for (const [column, width] of Object.entries(widths)) {
    page.getRange(`${column}:${column}`).format.columnWidthPx = width;
  }
  return page;
}

function merged(page, range, value) {
  const cell = page.getRange(range);
  cell.merge();
  page.getRange(range.split(":")[0]).values = [[value]];
  return cell;
}

function title(page, last, text, note) {
  const heading = merged(page, `A1:${last}1`, text);
  heading.format.fill = navy;
  heading.format.font = { name: "Arial", size: 21, bold: true, color: white };
  heading.format.rowHeightPx = 43;
  heading.format.verticalAlignment = "center";
  const caption = merged(page, `A2:${last}2`, note);
  caption.format.font = { name: "Arial", size: 11, color: slate };
  caption.format.rowHeightPx = 34;
  caption.format.wrapText = true;
  caption.format.verticalAlignment = "center";
}

function band(page, address, text) {
  const row = merged(page, address, text);
  row.format.fill = navy;
  row.format.font = { name: "Arial", size: 11, bold: true, color: white };
  row.format.rowHeightPx = 25;
  row.format.verticalAlignment = "center";
}

function paragraph(page, address, text, height = 28) {
  const row = merged(page, address, text);
  row.format.font = { name: "Arial", size: 11, color: slate };
  row.format.wrapText = true;
  row.format.rowHeightPx = height;
  row.format.verticalAlignment = "center";
}

function grid(page, start, end, last, headings, rows) {
  const header = page.getRange(`A${start}:${last}${start}`);
  header.values = [headings];
  header.format.fill = navy;
  header.format.font = { name: "Arial", size: 11, bold: true, color: white };
  header.format.wrapText = true;
  header.format.rowHeightPx = 40;
  header.format.verticalAlignment = "center";
  if (!rows.length) return;
  const body = page.getRange(`A${start + 1}:${last}${end}`);
  body.values = rows;
  body.format.font = { name: "Arial", size: 11, color: ink };
  body.format.rowHeightPx = 22;
  body.format.verticalAlignment = "center";
  for (let row = start + 2; row <= end; row += 2) {
    page.getRange(`A${row}:${last}${row}`).format.fill = mist;
  }
}

function numeric(page, address, format) {
  page.getRange(address).setNumberFormat(format);
}

const overview = sheet("Overview", { A: 285, B: 130, C: 130, D: 130, E: 130 });
const paths = sheet("Market Paths", {
  A: 80, B: 65, C: 130, D: 130, E: 130, F: 130, G: 130, H: 130, I: 130, J: 26,
  K: 110, L: 110, M: 110, N: 110, O: 110, P: 110,
});
const stress = sheet("Early Losses", { A: 125, B: 135, C: 135, D: 150, E: 150, F: 130, G: 145 });
const assumptions = sheet("Assumptions", {
  A: 185, B: 125, C: 90, D: 104, E: 104, F: 112, G: 112, H: 112, I: 112, J: 112,
});
const { provenance: run, timeline, baseline: base, fixed, flexible, controls, reconciliation: check } = data;
const ageLabel = (age) => age ?? "None";
const intervals = (s) => s.success_interval_95.map((x) => `${(x * 100).toFixed(1)}%`).join(" to ");

title(overview, "E", "Retirement portfolio risk",
  `Scenario ${run.scenario}  •  Retire at ${timeline.retirement_age}  •  Plan through age ${timeline.end_age}  •  ${run.trials.toLocaleString("en-US")} paired market paths`);
grid(overview, 4, 11, "E",
  ["Funding result", "Baseline", "Fixed spending", "Flexible spending", ""],
  [
    ["Paths that paid all required costs", null, fixed.success_rate, flexible.success_rate, null],
    ["95% simulation interval", null, intervals(fixed), intervals(flexible), null],
    ["Paths with a shortfall", base.first_shortfall_age == null ? 0 : 1, fixed.failed_trials, flexible.failed_trials, null],
    ["First shortfall age¹", ageLabel(base.first_shortfall_age), ageLabel(fixed.median_failure_age), ageLabel(flexible.median_failure_age), null],
    ["Ending balance²", base.ending_balance_real, fixed.ending_balance_p50, flexible.ending_balance_p50, null],
    ["Worst annual cash shortfall²", base.max_cash_shortfall_real, fixed.max_cash_shortfall, flexible.max_cash_shortfall, null],
    ["Paths that cut spending", null, 0, data.spending_cuts.trials_with_cuts, null],
  ]);
numeric(overview, "C5:D5", percent);
numeric(overview, "B7:D7", count);
numeric(overview, "B9:D10", money);
numeric(overview, "C11:D11", count);
overview.getRange("C5:D5").format.fill = mint;
overview.getRange("C5:D5").format.font = { name: "Arial", size: 17, bold: true, color: teal };
overview.getRange("A5:E5").format.rowHeightPx = 35;
paragraph(overview, "A13:E13", "¹ For simulated paths that had a shortfall, this is the median age of the first one.");
paragraph(overview, "A14:E14", `² Values are in ${timeline.opening_year} dollars. Simulated ending balances show the median.`);
band(overview, "A16:E16", "How to read this report");
paragraph(overview, "A17:E17", "A path has a shortfall when it leaves at least one cent of required spending or tax unpaid in any year. A later recovery still counts as a shortfall.", 36);
paragraph(overview, "A18:E18", "The two spending policies use the same market paths. The flexible policy can cut part of the recurring retirement budget. Mortgage payments, healthcare, taxes and one-time expenses stay outside the cut.", 38);
paragraph(overview, "A19:E19", "The simulation interval reflects sampling variation under the chosen return assumptions. Different market assumptions can change the result.", 36);
band(overview, "A21:E21", "Reconciliation");
paragraph(overview, "A22:E22", `With no market shocks, all ${check.values_compared.toLocaleString("en-US")} cash, tax and account values matched the workbook to within one cent across ${check.years} years.`);
paragraph(overview, "A23:E23", "This report is a snapshot. Rerun the risk analysis after changing workbook inputs.");

title(paths, "I", "Market paths by age",
  `Portfolio balances in ${timeline.opening_year} dollars, before any estimated tax on remaining balances. P10 means 10% of simulated balances are at or below the value.`);
const marketRows = data.annual.map((year) => [year.year, year.age, year.base, ...year.fixed_p10_p50_p90, ...year.flexible_p10_p50_p90]);
grid(paths, 4, marketRows.length + 4, "I",
  ["Year", "Age", "Workbook baseline", "Fixed P10", "Fixed P50", "Fixed P90", "Flexible P10", "Flexible P50", "Flexible P90"],
  marketRows);
numeric(paths, `C5:I${marketRows.length + 4}`, money);
paths.freezePanes.freezeRows(4);
const chart = paths.charts.add("line", paths.getRange(`B4:F${marketRows.length + 4}`));
chart.setPosition("K4", "Q23");
chart.title = "Portfolio balances with fixed spending";
chart.titleTextStyle.typeface = "Arial";
chart.legend = { position: "bottom", textStyle: { typeface: "Arial" } };
chart.xAxis = { axisType: "textAxis", tickLabelInterval: 10, textStyle: { typeface: "Arial", fontSize: 10 } };
chart.xAxis.title.text = "Age";
chart.yAxis = { numberFormatCode: '"$"0,,"M"', numberFormatSourceLinked: false, textStyle: { typeface: "Arial" } };

const failureAges = [...new Set([
  ...Object.keys(fixed.failure_age_counts), ...Object.keys(flexible.failure_age_counts),
])].map(Number).sort((a, b) => a - b);
const failureHeader = paths.getRange("K25:M25");
failureHeader.values = [["First shortfall age", "Fixed paths", "Flexible paths"]];
failureHeader.format.fill = navy;
failureHeader.format.font = { name: "Arial", size: 11, bold: true, color: white };
failureHeader.format.wrapText = true;
failureHeader.format.rowHeightPx = 40;
if (failureAges.length) {
  paths.getRange(`K26:M${25 + failureAges.length}`).values = failureAges.map((age) => [
    age, fixed.failure_age_counts[age] ?? 0, flexible.failure_age_counts[age] ?? 0,
  ]);
} else {
  paragraph(paths, "K26:P26", "No simulated path had a funding shortfall.");
}

const lossNote = controls.stress_years_applied
  ? `The listed annual equity return applies for the first ${controls.stress_years_applied} modeled retirement years. Other assets and later years use the workbook assumptions.`
  : "The selected horizon and shock length contain no early retirement stress years.";
title(stress, "G", "Early retirement losses", lossNote);
const stressRows = data.sequence.map((row) => [
  row.equity_return, ageLabel(row.fixed.first_shortfall_age), ageLabel(row.flexible.first_shortfall_age),
  row.fixed.ending_balance_real, row.flexible.ending_balance_real, row.flexible.years_with_cuts,
  row.flexible.total_cuts_real,
]);
grid(stress, 4, 4 + stressRows.length, "G",
  ["Equity return", "First shortfall (fixed)", "First shortfall (flexible)", "Ending balance (fixed)",
    "Ending balance (flexible)", "Years with cuts", "Total cuts"],
  stressRows);
if (stressRows.length) {
  stress.getRange(`B5:C${4 + stressRows.length}`).format.horizontalAlignment = "center";
  numeric(stress, `A5:A${4 + stressRows.length}`, percent);
  numeric(stress, `D5:E${4 + stressRows.length}`, money);
  numeric(stress, `G5:G${4 + stressRows.length}`, money);
}
const stressFoot = 6 + stressRows.length;
paragraph(stress, `A${stressFoot}:G${stressFoot}`, `Balances and total spending cuts are in ${timeline.opening_year} dollars.`);
stress.freezePanes.freezeRows(4);

title(assumptions, "J", "Assumptions and source",
  "Returns and account allocations come from the active workbook scenario. The same portfolio can appear in both stages if it was selected for both.");
const allocationRows = data.allocations.map((row) => [
  row.stage, row.account, row.profile, row.base_return, row.return_volatility, ...row.weights,
]);
const allocationLast = String.fromCharCode(69 + data.asset_names.length);
grid(assumptions, 4, 4 + allocationRows.length, allocationLast,
  ["Stage", "Account", "Profile", "Base return", "Volatility", ...data.asset_names], allocationRows);
numeric(assumptions, `D5:${allocationLast}${4 + allocationRows.length}`, percent);
band(assumptions, "A14:J14", "Spending and return settings");
const runSettings = [
  ["Spending cut trigger", controls.trigger, "The opening portfolio falls below this share of the workbook baseline."],
  ["Flexible spending share", controls.flexible_share, "The eligible share of recurring retirement spending after mortgage payments."],
  ["Cut to flexible spending", controls.cut, "The cut applies only in years the trigger is met."],
  ["Volatility multiplier", run.volatility_multiplier, "Scales the asset-level volatility assumptions."],
  ["Return basis", run.return_basis === 1 ? "Compound growth" : "Arithmetic mean", "The workbook sets how the assumed asset return anchors the simulation."],
  ["Scenario and seed", `${run.scenario} / ${run.seed}`, "Rerun the same inputs and seed in the same engine version to reproduce the paths."],
];
for (const [index, [name, value, note]] of runSettings.entries()) {
  const row = 15 + index;
  assumptions.getRange(`A${row}:B${row}`).values = [[name, value]];
  assumptions.getRange(`A${row}:B${row}`).format.font = { name: "Arial", size: 11, color: ink };
  paragraph(assumptions, `D${row}:J${row}`, note, 24);
}
numeric(assumptions, "B15:B17", percent);
band(assumptions, "A22:J22", "Source and calculation");
paragraph(assumptions, "A23:J23", `Workbook SHA-256  ${run.workbook_sha256}`);
paragraph(assumptions, "A24:J24", `Active assumptions SHA-256  ${run.inputs_sha256}`);
paragraph(assumptions, "A25:J25", `Risk engine ${run.engine_version}  •  Python ${run.python}  •  ${run.workbook_calculation}`);
paragraph(assumptions, "A26:J26", "Returns vary by asset and year. The model uses the workbook’s tax and healthcare estimates. Other inputs, including wages, inflation and longevity, follow the selected scenario.");
paragraph(assumptions, "A27:J27", "The workbook uses its Roth return for deterministic HSA growth. The HSA allocation determines its exposure to market shocks.");

wb.recalculate();
if (previewDir) {
  for (const [name, range] of [["Overview", "A1:E23"], ["Market Paths", "A1:Q25"], ["Early Losses", "A1:G19"], ["Assumptions", "A1:J27"]]) {
    const preview = await wb.render({ sheetName: name, range, scale: 1.15, format: "png" });
    await fs.writeFile(`${previewDir}/${name.toLowerCase().replaceAll(" ", "-")}.png`,
      new Uint8Array(await preview.arrayBuffer()));
  }
}

const temporary = await fs.mkdtemp(join(dirname(output), ".portfolio-risk-"));
try {
  const stage = join(temporary, "report.xlsx");
  const blob = await SpreadsheetFile.exportXlsx(wb);
  await blob.save(stage);
  await fs.link(stage, output);
} finally {
  await fs.rm(temporary, { recursive: true, force: true });
}
