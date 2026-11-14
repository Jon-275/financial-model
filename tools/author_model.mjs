import fs from 'node:fs/promises';
import { Workbook, SpreadsheetFile } from '@oai/artifact-tool';
import { addDashboardCharts } from './dashboard_charts.mjs';

const [planPath, out] = process.argv.slice(2);
const plan = JSON.parse(await fs.readFile(planPath, 'utf8'));
const wb = Workbook.create();
for (const name of plan.order) wb.worksheets.add(name);
const rank = (s) => [...s].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0);
for (const [name, cells] of Object.entries(plan.cells)) {
  const sheet = wb.worksheets.getItem(name);
  const rows = new Map();
  for (const [address, value] of Object.entries(cells)) {
    const [, col, r] = /^([A-Z]+)(\d+)$/.exec(address);
    if (!rows.has(r)) rows.set(r, []);
    rows.get(r).push({ address, col: rank(col), value, formula: typeof value === 'string' && value.startsWith('=') });
  }
  for (const list of rows.values()) list.sort((a, b) => a.col - b.col);
  // INDEX can temporarily spill over empty cells while its scenario inputs are absent.
  for (const formulas of [false, true]) {
    for (const list of rows.values()) {
      const entries = list.filter((cell) => cell.formula === formulas);
      for (let first = 0; first < entries.length;) {
        let end = first + 1;
        while (end < entries.length && entries[end].col === entries[end - 1].col + 1) end++;
        const range = sheet.getRange(`${entries[first].address}:${entries[end - 1].address}`);
        const values = [entries.slice(first, end).map((cell) => cell.value)];
        if (formulas) range.formulas = values;
        else range.values = values;
        first = end;
      }
    }
  }
}
if (plan.order.includes('Report_Data')) addDashboardCharts(wb);
await (await SpreadsheetFile.exportXlsx(wb)).save(out);
