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
  for (const list of rows.values()) {
    list.sort((a, b) => a.col - b.col);
    for (let i = 0; i < list.length;) {
      let j = i + 1;
      while (j < list.length && list[j].col === list[j - 1].col + 1 && list[j].formula === list[i].formula) j++;
      const range = sheet.getRange(`${list[i].address}:${list[j-1].address}`);
      const values = [list.slice(i,j).map((v) => v.value)];
      if (list[i].formula) range.formulas = values;
      else range.values = values;
      i = j;
    }
  }
}
if (plan.order.includes('Report_Data')) addDashboardCharts(wb);
await (await SpreadsheetFile.exportXlsx(wb)).save(out);
