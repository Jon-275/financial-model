/** Author native chart parts that are combined on one age and dollar axis. */
const font = 'Arial';

function series(chart, title, values) {
  const item = chart.series.add(title);
  item.categoryFormula = 'Report_Data!$A$2:$A$127';
  item.valuesSource = `Report_Data!$${values}$2:$${values}$127`;
  return item;
}

export function addDashboardCharts(workbook) {
  const summary = workbook.worksheets.getItem('Summary');
  const chart = summary.charts.add('line', {
    from: { row: 13, col: 1 },
    extent: { widthPx: 610, heightPx: 300 },
  });
  chart.setPosition('B14', 'I28');
  chart.title = 'Portfolio balance by age';
  chart.titleTextStyle.typeface = font;
  chart.titleTextStyle.fontSize = 15;
  chart.legend = { position: 'bottom', textStyle: { typeface: font, fontSize: 10 } };
  chart.displayBlanksAs = 'gap';
  chart.xAxis = {
    axisType: 'textAxis', numberFormatCode: '0', numberFormatSourceLinked: false,
    textStyle: { typeface: font, fontSize: 10 },
  };
  chart.xAxis.title.text = 'Age';
  chart.yAxis = {
    min: 0, numberFormatCode: '"$"#,##0', numberFormatSourceLinked: false,
    textStyle: { typeface: font, fontSize: 10 },
    majorGridlines: { fill: '#D9E4EB', width: 1, style: 'solid' },
  };

  series(chart, 'Base projection', 'AE');
  chart.applySeriesStyle(0, {
    stroke: { fill: '#344D64', width: 2, style: 'dashed' },
    marker: { symbol: 'none' },
  });
  series(chart, 'Simulated median', 'AH');
  chart.applySeriesStyle(1, {
    stroke: { fill: '#126761', width: 2.5, style: 'solid' },
    marker: { symbol: 'none' },
  });

  const range = summary.charts.add('area', {
    from: { row: 13, col: 1 },
    extent: { widthPx: 610, heightPx: 300 },
  });
  series(range, '10th percentile', 'AF');
  series(range, 'Middle 80% of balances', 'AG').fill = '#C7E6DF';
}
