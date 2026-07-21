function refreshDashboard_() {
  const dash = getRequiredSheet_(AD.SHEETS.DASHBOARD);
  ensureDashboardLayout_(dash);
  const cat = getRequiredSheet_(AD.SHEETS.CATALOG);
  const input = getInputRows_();
  const rows = cat.getLastRow() >= 2
    ? cat.getRange(2, 1, cat.getLastRow() - 1, AD.CATALOG_HEADERS.length).getValues()
    : [];

  const ratings = rows.map(r => number_(r[16])).filter(x => x > 0);
  const hist = rows.map(r => number_(r[11])).filter(x => x > 0);
  const current = rows.map(r => number_(r[10])).filter(x => x > 0);
  const lastUpdates = rows.map(r => r[18]).filter(isValidDate_).map(d => new Date(d));
  const latestRankUpdate = lastUpdates.length
    ? new Date(Math.max(...lastUpdates.map(d => d.getTime())))
    : '';

  let topBook = '';
  let topRank = null;
  rows.forEach(r => {
    const rank = number_(r[10]);
    if (rank > 0 && (topRank == null || rank < topRank)) {
      topRank = rank;
      topBook = clean_(r[1]);
    }
  });

  const seriesByFormat = getOverallRankSeriesByFormat_();
  const series = getOverallRankSeries_();
  let trend = '';
  if (series.length >= 2) {
    const prev = series[series.length - 2].rank;
    const latest = series[series.length - 1].rank;
    if (latest < prev) trend = 'Improving (lower rank)';
    else if (latest > prev) trend = 'Worsening (higher rank)';
    else trend = 'Unchanged';
  } else {
    trend = 'Need more history';
  }

  const metrics = [
    rows.length,
    rows.filter(r => clean_(r[4]).toLowerCase() === 'published').length,
    rows.filter(r => !['', 'published', 'paused', 'cancelled'].includes(clean_(r[4]).toLowerCase())).length,
    input.filter(r => clean_(r[1])).length,
    input.filter(r => clean_(r[13]).toLowerCase() === 'live').length,
    rows.reduce((s, r) => s + number_(r[5]), 0),
    rows.reduce((s, r) => s + number_(r[12]), 0),
    rows.reduce((s, r) => s + number_(r[13]), 0),
    rows.reduce((s, r) => s + number_(r[14]), 0),
    rows.reduce((s, r) => s + number_(r[15]), 0),
    ratings.length ? average_(ratings) : '',
    hist.length ? Math.min(...hist) : '',
    current.length ? Math.min(...current) : '',
    latestRankUpdate,
    topBook ? (topBook + (topRank != null ? ' (#' + topRank.toLocaleString('en-US') + ')' : '')) : '',
    trend
  ];

  dash.getRange(4, 2, metrics.length, 1).setValues(metrics.map(x => [x]));

  // Catalog Performance block (D4 header, data from D5). Clear a tall range so growth doesn't leave stale rows.
  clearBlockUnmerged_(dash, 5, 4, 200, 5);
  const perf = rows.map(r => [r[1], r[4], r[12], r[14], r[11]]).sort((a, b) => number_(b[3]) - number_(a[3]));
  if (perf.length) dash.getRange(5, 4, perf.length, 5).setValues(perf);

  dash.getRange('B4:B11').setNumberFormat('#,##0');
  dash.getRange('B12').setNumberFormat('$#,##0.00');
  dash.getRange('B13').setNumberFormat('#,##0');
  dash.getRange('B14').setNumberFormat('0.0');
  dash.getRange('B15:B16').setNumberFormat('#,##0');
  dash.getRange('B17').setNumberFormat('m/d/yyyy');
  dash.getRange('F5:F').setNumberFormat('#,##0');
  dash.getRange('G5:G').setNumberFormat('$#,##0.00');
  dash.getRange('H5:H').setNumberFormat('#,##0');

  // Layout:
  //  A-B metrics | D-H portfolio | J+ chart (right side)
  //  Category ranks start BELOW whichever of metrics/portfolio ends lower.
  const metricsEndRow = 3 + metrics.length; // labels begin at row 4
  const catalogEndRow = perf.length ? (4 + perf.length) : 4;
  const categoryStartRow = Math.max(metricsEndRow, catalogEndRow) + 2;

  // Wipe prior category blocks (old fixed start at row 22 and any later dynamic starts).
  clearBlockUnmerged_(dash, 20, 1, 250, 6);

  refreshCategoryRankTable_(dash, categoryStartRow);
  refreshRankTrendChart_(dash, seriesByFormat);
}

function ensureDashboardLayout_(sheet) {
  const labels = [
    'Total Books',
    'Published Books',
    'Books in Progress',
    'Total Store Listings',
    'Live Store Listings',
    'Total Published Words',
    'Lifetime Unit Sales',
    'Lifetime KU Pages',
    'Lifetime Royalties',
    'Total Reviews',
    'Average Rating',
    'Best Rank Ever',
    'Current Best Rank',
    'Latest Rank Update',
    'Top-Ranked Book',
    'Rank Trend (lower is better)'
  ];
  const current = sheet.getRange(4, 1, labels.length, 1).getValues().map(r => clean_(r[0]));
  const needs = labels.some((label, i) => normalizeKey_(current[i] || '') !== normalizeKey_(label));
  if (needs) {
    sheet.getRange(4, 1, labels.length, 1).setValues(labels.map(x => [x]));
  }
  if (!clean_(sheet.getRange('A1').getValue())) {
    sheet.getRange('A1:H1').merge().setValue('Author Portfolio Dashboard')
      .setFontSize(22).setFontWeight('bold').setHorizontalAlignment('center')
      .setBackground('#1f4e78').setFontColor('#ffffff');
  }
  // Leave room on the right for the chart.
  [225, 160, 30, 270, 120, 100, 110, 110].forEach((w, i) => {
    if (sheet.getColumnWidth(i + 1) < w) sheet.setColumnWidth(i + 1, w);
  });
}

function refreshCategoryRankTable_(dash, startRow) {
  const row0 = startRow || 22;
  const cols = 4;

  const byFormat = getCategoryRankSummaryByFormat_();
  const sections = [
    { title: 'eBook Category Ranks (lower number = better)', rows: byFormat.ebook },
    { title: 'Paperback Category Ranks (lower number = better)', rows: byFormat.paperback },
    { title: 'Hardcover Category Ranks (lower number = better)', rows: byFormat.hardcover }
  ];

  let row = row0;
  let wroteAny = false;

  sections.forEach(section => {
    mergeRowSafe_(dash, row, 1, cols)
      .setValue(section.title)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBackground('#1f4e78')
      .setFontColor('#ffffff');
    row++;

    dash.getRange(row, 1, 1, cols).setValues([[
      'Category',
      'Best Rank Ever',
      'Current Best Rank',
      'Last Seen'
    ]]);
    styleHeader_(dash.getRange(row, 1, 1, cols));
    row++;

    if (!section.rows.length) {
      dash.getRange(row, 1).setValue('No category ranks yet for this format.');
      row += 2;
      return;
    }

    wroteAny = true;
    const values = section.rows.map(r => [
      r.category,
      r.bestEver,
      r.currentBest,
      r.lastSeen
    ]);
    dash.getRange(row, 1, values.length, cols).setValues(values);
    dash.getRange(row, 2, values.length, 2).setNumberFormat('#,##0');
    dash.getRange(row, 4, values.length, 1).setNumberFormat('m/d/yyyy');
    row += values.length + 1;
  });

  if (!wroteAny) {
    dash.getRange(row0 + 2, 1).setValue('No category ranks yet. Run Update Amazon Rankings Now.');
  }

  [280, 120, 130, 110].forEach((w, i) => dash.setColumnWidth(i + 1, Math.max(dash.getColumnWidth(i + 1) || 0, w)));
}

function refreshRankTrendChart_(dash, seriesByFormat) {
  // Chart floats on the right (column J). Source data lives farther right so it doesn't clutter the view.
  const chartAnchorRow = 3;
  const chartAnchorCol = 10; // J
  const dataStartRow = 3;
  const dataStartCol = 20; // column T — chart source data
  const data = seriesByFormat || { dates: [], series: { ebook: [], paperback: [], hardcover: [] } };

  clearBlockUnmerged_(dash, dataStartRow, dataStartCol, 200, 4);
  dash.getRange(dataStartRow, dataStartCol, 1, 4).setValues([[
    'Snapshot Date',
    'eBook Overall',
    'Paperback Overall',
    'Hardcover Overall'
  ]]);

  if (data.dates && data.dates.length) {
    const values = data.dates.map((d, i) => [
      d,
      data.series.ebook[i],
      data.series.paperback[i],
      data.series.hardcover[i]
    ]);
    dash.getRange(dataStartRow + 1, dataStartCol, values.length, 4).setValues(values);
    dash.getRange(dataStartRow + 1, dataStartCol, values.length, 1).setNumberFormat('m/d/yyyy');
    dash.getRange(dataStartRow + 1, dataStartCol + 1, values.length, 3).setNumberFormat('#,##0');
  }

  // Hide chart source columns so the sheet stays clean.
  try { dash.hideColumns(dataStartCol, 4); } catch (e) {}

  dash.getCharts().forEach(c => {
    try {
      const title = String((c.getOptions() && c.getOptions().get('title')) || '');
      if (/rank/i.test(title) || /overall/i.test(title)) dash.removeChart(c);
    } catch (e) {}
  });

  if (!data.dates || data.dates.length < 2) return;

  const dataRange = dash.getRange(dataStartRow, dataStartCol, data.dates.length + 1, 4);
  const chart = dash.newChart()
    .asLineChart()
    .addRange(dataRange)
    .setTitle('Overall Amazon Rank Over Time by Format (lower is better)')
    .setXAxisTitle('Snapshot Date')
    .setYAxisTitle('Overall Rank')
    .setNumHeaders(1)
    .setLegendPosition(Charts.Position.BOTTOM)
    .setOption('curveType', 'function')
    .setOption('pointSize', 5)
    .setOption('width', 720)
    .setOption('height', 360)
    .setPosition(chartAnchorRow, chartAnchorCol, 0, 0)
    .build();
  dash.insertChart(chart);
}
