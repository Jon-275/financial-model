/** Author the editable portfolio chart from the active projection. */
const font = 'Arial';

export function addDashboardCharts(workbook) {
  const summary = workbook.worksheets.getItem('Summary');
  const chart = summary.charts.add('scatter', {
    from: { row: 13, col: 1 },
    extent: { widthPx: 610, heightPx: 300 },
  });
  chart.title = 'Projected portfolio by age';
  chart.setPosition('B14', 'I28');
  chart.titleTextStyle.typeface = font;
  chart.titleTextStyle.fontSize = 16;
  chart.legend = { position: 'bottom', textStyle: { typeface: font, fontSize: 11 } };
  chart.displayBlanksAs = 'gap';
  chart.scatterOptions.style = 'lineWithMarkers';
  chart.xAxis = {
    numberFormatCode: '0', numberFormatSourceLinked: false,
    textStyle: { typeface: font, fontSize: 11 },
  };
  chart.xAxis.title.text = 'Age';
  chart.yAxis = {
    min: 0,
    numberFormatCode: '"$"#,##0',
    numberFormatSourceLinked: false,
    textStyle: { typeface: font, fontSize: 11 },
    majorGridlines: { fill: '#D9E4EB', width: 1, style: 'solid' },
  };

  const balance = chart.series.add('Portfolio');
  balance.xValuesSource = 'Report_Data!$A$2:$A$127';
  balance.valuesSource = 'Report_Data!$B$2:$B$127';
  chart.applySeriesStyle(0, {
    stroke: { fill: '#182C42', width: 2.5, style: 'solid' },
    marker: { symbol: 'none' },
  });

  const milestones = [
    ['Retirement', 'AA', 'AB', '#126761'],
    ['First shortfall', 'AC', 'AD', '#9E3534'],
  ];
  for (const [index, [label, ages, balances, color]] of milestones.entries()) {
    const guide = chart.series.add(label);
    guide.xValuesSource = `Report_Data!$${ages}$2:$${ages}$3`;
    guide.valuesSource = `Report_Data!$${balances}$2:$${balances}$3`;
    chart.applySeriesStyle(index + 1, {
      stroke: { fill: color, width: 1.5, style: 'dashed' },
      marker: { symbol: 'none' },
    });
  }
}
