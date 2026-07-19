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
  dash.getRange('D5:H1000').clearContent();
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

  refreshRankTrendChart_(dash, series);
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
}

function refreshRankTrendChart_(dash, series) {
  const startRow = 3;
  const startCol = 10; // column J
  dash.getRange(startRow, startCol, 200, 2).clearContent();
  dash.getRange(startRow, startCol, 1, 2).setValues([['Snapshot Date', 'Best Overall Rank (lower is better)']]);

  if (series && series.length) {
    const values = series.map(p => [p.date, p.rank]);
    dash.getRange(startRow + 1, startCol, values.length, 2).setValues(values);
    dash.getRange(startRow + 1, startCol, values.length, 1).setNumberFormat('m/d/yyyy');
    dash.getRange(startRow + 1, startCol + 1, values.length, 1).setNumberFormat('#,##0');
  }

  dash.getCharts().forEach(c => {
    try {
      const title = String((c.getOptions() && c.getOptions().get('title')) || '');
      if (/rank/i.test(title)) dash.removeChart(c);
    } catch (e) {}
  });

  if (!series || series.length < 2) return;

  const dataRange = dash.getRange(startRow, startCol, series.length + 1, 2);
  const chart = dash.newChart()
    .asLineChart()
    .addRange(dataRange)
    .setTitle('Best Overall Amazon Rank Over Time (lower is better)')
    .setXAxisTitle('Snapshot Date')
    .setYAxisTitle('Best Overall Rank')
    .setNumHeaders(1)
    .setPosition(20, 1, 0, 0)
    .build();
  dash.insertChart(chart);
}
