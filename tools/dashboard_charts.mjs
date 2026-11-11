/** Author the two editable dashboard charts. Their data lives on Report_Data. */
const font = "Arial";
const dollars = '"$"#,##0';

function style(chart, title, first, last) {
  chart.title = title;
  chart.setPosition(first, last);
  chart.titleTextStyle.typeface = font;
  chart.titleTextStyle.fontSize = 15;
  chart.legend = { position: "bottom", textStyle: { typeface: font, fontSize: 10 } };
  chart.displayBlanksAs = "gap";
  chart.yAxis = {
    min: 0,
    numberFormatCode: dollars,
    numberFormatSourceLinked: false,
    textStyle: { typeface: font, fontSize: 10 },
    majorGridlines: { fill: "#E2E8EF", width: 1, style: "solid" },
  };
}

export function addDashboardCharts(workbook) {
  const summary = workbook.worksheets.getItem("Summary");
  const portfolio = summary.charts.add("scatter", {
    from: { row: 2, col: 4 },
    extent: { widthPx: 640, heightPx: 290 },
  });
  style(portfolio, "Portfolio balance by age (today’s dollars)", "E3", "P19");
  portfolio.scatterOptions.style = "lineWithMarkers";
  portfolio.xAxis = {
    numberFormatCode: "0", numberFormatSourceLinked: false,
    textStyle: { typeface: font, fontSize: 10 },
  };
  portfolio.xAxis.title.text = "Age";

  const balance = portfolio.series.add("Portfolio balance");
  balance.xValuesSource = "Report_Data!$A$2:$A$127";
  balance.valuesSource = "Report_Data!$B$2:$B$127";
  portfolio.applySeriesStyle(0, {
    stroke: { fill: "#173D62", width: 2.5, style: "solid" },
    marker: { symbol: "none" },
  });

  const retirement = portfolio.series.add("Retirement");
  retirement.xValuesSource = "Report_Data!$A$2";
  retirement.valuesSource = "Report_Data!$B$2";
  portfolio.applySeriesStyle(1, {
    stroke: { fill: "#BF751B", width: 2, style: "solid" },
    marker: { symbol: "diamond", size: 9, fill: "#BF751B" },
  });

  const cash = summary.charts.add("bar", {
    from: { row: 20, col: 4 },
    extent: { widthPx: 640, heightPx: 290 },
  });
  style(cash, "How the plan pays its bills (today’s dollars)", "E21", "P37");
  cash.format.apply({
    barOptions: { direction: "column", grouping: "stacked", gapWidth: 50, overlap: 100 },
  });
  cash.xAxis = {
    axisType: "textAxis", tickLabelInterval: 1,
    textStyle: { typeface: font, fontSize: 10 },
  };
  cash.xAxis.title.text = "Age around retirement";
  for (const [column, label, color] of [
    ["F", "Income used", "#00776E"],
    ["G", "Withdrawals used", "#4677A6"],
    ["H", "Health savings used", "#735997"],
    ["I", "Unpaid", "#B5433D"],
  ]) {
    const series = cash.series.add(label);
    series.categoriesSource = "Report_Data!$E$2:$E$14";
    series.valuesSource = `Report_Data!$${column}$2:$${column}$14`;
    series.fill = color;
  }
}
