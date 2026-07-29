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

  // Catalog Performance starts at column F (leave D–E as spacer).
  clearBlockUnmerged_(dash, 3, 4, 200, 10);
  mergeRowSafe_(dash, 3, 6, 7)
    .setValue('Catalog Performance')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBackground('#1f4e78')
    .setFontColor('#ffffff');
  dash.getRange(4, 6, 1, 7).setValues([[
    'Book', 'Stage', 'Units', 'KENP Read', 'Royalties (USD)', 'Comments', 'Best Rank'
  ]]);
  styleHeader_(dash.getRange(4, 6, 1, 7));

  // Catalog cols: 1 title, 4 stage, 12 units, 13 ku, 14 roy, 15 reviews, 11 best rank ever
  const perf = rows
    .map(r => [r[1], r[4], r[12], r[13], r[14], r[15], r[11]])
    .sort((a, b) => number_(b[4]) - number_(a[4]));
  if (perf.length) dash.getRange(5, 6, perf.length, 7).setValues(perf);

  dash.getRange('B4:B11').setNumberFormat('#,##0');
  dash.getRange('B12').setNumberFormat('$#,##0.00');
  dash.getRange('B13').setNumberFormat('#,##0');
  dash.getRange('B14').setNumberFormat('0.0');
  dash.getRange('B15:B16').setNumberFormat('#,##0');
  dash.getRange('B17').setNumberFormat('m/d/yyyy');
  dash.getRange('H5:I').setNumberFormat('#,##0');
  dash.getRange('J5:J').setNumberFormat('$#,##0.00');
  dash.getRange('K5:K').setNumberFormat('#,##0');
  dash.getRange('L5:L').setNumberFormat('#,##0');
  if (dash.getColumnWidth(11) < 90) dash.setColumnWidth(11, 90);
  if (dash.getColumnWidth(12) < 100) dash.setColumnWidth(12, 100);

  const metricsEndRow = 3 + metrics.length;
  const catalogEndRow = perf.length ? (4 + perf.length) : 4;
  const categoryStartRow = Math.max(metricsEndRow, catalogEndRow) + 2;

  // Clear space below metrics (drops old Meta/KDP rows that moved to Statistics).
  clearBlockUnmerged_(dash, metricsEndRow + 1, 1, 250, 5);

  // Remove any leftover charts from the numbers Dashboard (charts live on Visual Dashboard).
  dash.getCharts().forEach(c => {
    try { dash.removeChart(c); } catch (e) {}
  });

  refreshCategoryRankTable_(dash, categoryStartRow);
  refreshStatistics_();
  refreshVisualDashboard_();
  hideDiagnosticSheets_();
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
    'Lifetime Royalties (USD, est.)',
    'Total Reviews',
    'Average Rating',
    'Best Rank Ever',
    'Current Rank',
    'Latest Rank Update',
    'Top-Ranked Book',
    'Rank Trend (lower is better)'
  ];
  const current = sheet.getRange(4, 1, labels.length, 1).getValues().map(r => clean_(r[0]));
  const needs = labels.some((label, i) => normalizeKey_(current[i] || '') !== normalizeKey_(label));
  if (needs) {
    sheet.getRange(4, 1, labels.length, 1).setValues(labels.map(x => [x]));
  }
  try {
    sheet.getRange('A1:L1').getMergedRanges().forEach(m => {
      try { m.breakApart(); } catch (e) {}
    });
  } catch (e) {}
  sheet.getRange('A1:L1').merge()
    .setValue('Author Portfolio Dashboard')
    .setFontSize(22).setFontWeight('bold').setHorizontalAlignment('center')
    .setBackground('#1f4e78').setFontColor('#ffffff');
  sheet.setRowHeight(1, 46);
  [225, 220, 30, 40, 40, 220, 110, 90, 100, 110, 90, 100].forEach((w, i) => {
    if (sheet.getColumnWidth(i + 1) < w) sheet.setColumnWidth(i + 1, w);
  });
}

/** Create Statistics sheet if missing (locked automatic sheet for sync/health metrics). */
function ensureStatisticsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(AD.SHEETS.STATISTICS);
  if (!sh) {
    sh = ss.insertSheet(AD.SHEETS.STATISTICS);
    buildStatisticsSheet_(sh);
  }
  return sh;
}

function buildStatisticsSheet_(sheet) {
  sheet.clear();
  sheet.getRange('A1:B1').merge()
    .setValue('Statistics')
    .setFontSize(22).setFontWeight('bold').setHorizontalAlignment('center')
    .setBackground('#1f4e78').setFontColor('#ffffff');
  sheet.setRowHeight(1, 46);
  sheet.getRange('A3:B3').setValues([['Metric', 'Current Value']]);
  styleHeader_(sheet.getRange('A3:B3'));
  const labels = [
    'Last Run',
    'Daily Jobs',
    'Last Meta Sync',
    'Meta Sync Status',
    'KDP Months on File',
    'KDP Month Gaps'
  ];
  sheet.getRange(4, 1, labels.length, 2).setValues(labels.map(x => [x, '']));
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 420);
}

function setLastDashboardRun_(source) {
  PropertiesService.getDocumentProperties().setProperty('AD_LAST_RUN', JSON.stringify({
    at: new Date().toISOString(),
    source: source || 'refresh'
  }));
}

function getLastDashboardRun_() {
  try {
    const raw = PropertiesService.getDocumentProperties().getProperty('AD_LAST_RUN');
    return raw ? JSON.parse(raw) : null;
  } catch (e) {
    return null;
  }
}

/** Sync/health metrics — separate from portfolio Dashboard; protected with other auto sheets. */
function refreshStatistics_() {
  const sh = ensureStatisticsSheet_();
  ensureStatisticsLayout_(sh);

  const lastRun = getLastDashboardRun_();
  const metrics = [];
  metrics.push(lastRun && lastRun.at ? new Date(lastRun.at) : '');
  try {
    metrics.push(getDailyJobsStatus_());
  } catch (e) {
    metrics.push('Could not read triggers');
  }

  const metaHealth = getMetaSyncHealth_();
  if (metaHealth && metaHealth.at) {
    metrics.push(new Date(metaHealth.at));
    const stale = (Date.now() - new Date(metaHealth.at).getTime()) > (3 * 86400000);
    metrics.push(
      (metaHealth.status || 'OK') +
        (stale ? ' — STALE (>3 days)' : '') +
        (metaHealth.dateMax ? ' (data through ' + metaHealth.dateMax + ')' : '')
    );
  } else {
    metrics.push('');
    metrics.push(
      hasMetaApiCredentials_()
        ? 'No Meta sync yet — Refresh Everything to pull insights'
        : 'No Meta sync — run configureMetaApiCredentials() once, then Refresh Everything'
    );
  }

  const kdpMonths = getKdpMonthGapStatus_();
  metrics.push(kdpMonths.monthsOnFile || '(none)');
  metrics.push(kdpMonths.gapMessage || '');

  sh.getRange(4, 2, metrics.length, 1).setValues(metrics.map(x => [x]));
  sh.getRange('B4').setNumberFormat('m/d/yyyy hh:mm:ss');
  sh.getRange('B6').setNumberFormat('m/d/yyyy hh:mm:ss');

  const gapRow = 3 + metrics.length;
  try {
    const gapCell = sh.getRange(gapRow, 1, 1, 2);
    if (kdpMonths.hasGap) {
      gapCell.setBackground('#FFF3CD').setFontColor('#7A5B00');
    } else {
      gapCell.setBackground(null).setFontColor(null);
    }
  } catch (e) {}
}

function ensureStatisticsLayout_(sheet) {
  const labels = [
    'Last Run',
    'Daily Jobs',
    'Last Meta Sync',
    'Meta Sync Status',
    'KDP Months on File',
    'KDP Month Gaps'
  ];
  const current = sheet.getRange(4, 1, labels.length, 1).getValues().map(r => clean_(r[0]));
  const needs = labels.some((label, i) => normalizeKey_(current[i] || '') !== normalizeKey_(label));
  if (needs) {
    sheet.getRange(4, 1, labels.length, 1).setValues(labels.map(x => [x]));
  }
  try {
    sheet.getRange('A1:B1').getMergedRanges().forEach(m => {
      try { m.breakApart(); } catch (e) {}
    });
  } catch (e) {}
  sheet.getRange('A1:B1').merge()
    .setValue('Statistics')
    .setFontSize(22).setFontWeight('bold').setHorizontalAlignment('center')
    .setBackground('#1f4e78').setFontColor('#ffffff');
  sheet.setRowHeight(1, 46);
  sheet.getRange('A3:B3').setValues([['Metric', 'Current Value']]);
  styleHeader_(sheet.getRange('A3:B3'));
  if (sheet.getColumnWidth(1) < 200) sheet.setColumnWidth(1, 200);
  if (sheet.getColumnWidth(2) < 420) sheet.setColumnWidth(2, 420);
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
      'Current Rank',
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

function ensureVisualDashboard_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(AD.SHEETS.VISUAL);
  if (!sh) sh = ss.insertSheet(AD.SHEETS.VISUAL);
  return sh;
}

/** Hidden sheet that holds chart source ranges (never shown on Visual Dashboard). */
function ensureVisualDataSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(AD.SHEETS.VISUAL_DATA);
  if (!sh) sh = ss.insertSheet(AD.SHEETS.VISUAL_DATA);
  try { sh.hideSheet(); } catch (e) {}
  return sh;
}

/**
 * Charts-only sheet. Source data lives on a hidden sheet so nothing leaks into view.
 * Layout: 2 charts per row with spaced anchors.
 */
function refreshVisualDashboard_() {
  const sh = ensureVisualDashboard_();
  const dataSh = ensureVisualDataSheet_();
  prepareSheetForRebuild_(sh);
  prepareSheetForRebuild_(dataSh);
  sh.clear();
  dataSh.clear();
  sh.getCharts().forEach(c => {
    try { sh.removeChart(c); } catch (e) {}
  });

  // Light canvas — paint the full grid so scrolling right isn’t white.
  const bg = '#f4f6f8';
  const ink = '#1f4e78';
  const maxC = sh.getMaxColumns();
  const maxR = sh.getMaxRows();
  sh.setHiddenGridlines(true);
  try { sh.setTabColor('#1f4e78'); } catch (e) {}
  for (let c = 1; c <= Math.min(maxC, 40); c++) sh.setColumnWidth(c, 72);
  sh.setRowHeight(1, 48);
  sh.getRange(1, 1, maxR, maxC)
    .clearContent()
    .clearNote()
    .setBackground(bg)
    .setFontColor(ink);
  sh.getRange(1, 1, 1, maxC)
    .setBackground('#1f4e78')
    .setFontColor('#ffffff');
  sh.getRange(1, 1)
    .setValue('Visual Dashboard')
    .setFontWeight('bold')
    .setFontSize(22)
    .setFontColor('#ffffff')
    .setBackground('#1f4e78')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');

  const layout = {
    chartW: 540,
    chartH: 330,
    leftCol: 1,
    rightCol: 12,
    row1: 3,
    row2: 24,
    dataRow: 1,
    rankScoreCol: 1,
    ordersCol: 1,
    kenpCol: 1,
    bg: bg,
    ink: ink
  };

  const seriesByFormat = getOverallRankSeriesByFormat_();
  writeVisualRankChart_(sh, dataSh, seriesByFormat, layout);
  writeVisualSalesCharts_(sh, dataSh, layout);
  try { dataSh.hideSheet(); } catch (e) {}
}

/** Light chart theme — matches the rest of the workbook. */
function applyVisualChartTheme_(builder, layout) {
  const axis = {
    titleTextStyle: { color: '#5a6a7a', fontSize: 11 },
    textStyle: { color: '#3d4a57', fontSize: 10 },
    gridlines: { color: '#e2e8ee' },
    baselineColor: '#c5ced6'
  };
  return builder
    .setOption('width', layout.chartW)
    .setOption('height', layout.chartH)
    .setOption('backgroundColor', { fill: '#ffffff', stroke: '#d8dee6', strokeWidth: 1 })
    .setOption('chartArea', { left: '12%', top: '16%', width: '78%', height: '62%', backgroundColor: '#ffffff' })
    .setOption('titleTextStyle', { color: layout.ink || '#1f4e78', fontSize: 14, bold: true })
    .setOption('legend', { position: 'bottom', textStyle: { color: '#3d4a57', fontSize: 11 } })
    .setOption('hAxis', axis)
    .setOption('vAxis', axis);
}

/** Rank-score series on hidden data sheet + chart (top-left). */
function writeVisualRankChart_(sh, dataSh, data, layout) {
  const scoreCol = 1;
  const dataStartRow = 1;

  let maxRank = 0;
  (data.dates || []).forEach((_, i) => {
    ['ebook', 'paperback', 'hardcover'].forEach(fmt => {
      const n = number_(data.series[fmt][i]);
      if (n > maxRank) maxRank = n;
    });
  });
  const ceiling = maxRank > 0 ? maxRank : 1000000;
  const toScore = v => {
    const n = number_(v);
    return n > 0 ? (ceiling - n) : '';
  };

  dataSh.getRange(dataStartRow, scoreCol, 1, 4).setValues([[
    'Snapshot Date',
    'eBook Score',
    'Paperback Score',
    'Hardcover Score'
  ]]);

  if (!data.dates || !data.dates.length) return;

  const scores = data.dates.map((d, i) => [
    d,
    toScore(data.series.ebook[i]),
    toScore(data.series.paperback[i]),
    toScore(data.series.hardcover[i])
  ]);
  dataSh.getRange(dataStartRow + 1, scoreCol, scores.length, 4).setValues(scores);
  dataSh.getRange(dataStartRow + 1, scoreCol, scores.length, 1).setNumberFormat('m/d/yyyy');
  dataSh.getRange(dataStartRow + 1, scoreCol + 1, scores.length, 3).setNumberFormat('#,##0');

  if (data.dates.length < 2) return;

  const dataRange = dataSh.getRange(dataStartRow, scoreCol, data.dates.length + 1, 4);
  let builder = sh.newChart()
    .asLineChart()
    .addRange(dataRange)
    .setTitle('Overall rank trend')
    .setXAxisTitle('Snapshot Date')
    .setYAxisTitle('Rank score')
    .setNumHeaders(1)
    .setLegendPosition(Charts.Position.BOTTOM)
    .setOption('curveType', 'function')
    .setOption('pointSize', 5)
    .setOption('vAxes', { 0: { title: 'Rank score', format: '#,##0' } })
    .setPosition(layout.row1, layout.leftCol, 0, 0);
  builder = applyVisualChartTheme_(builder, layout);
  sh.insertChart(builder.build());
}

/** Orders (top-right) + KENP (bottom-left); source data on hidden sheet. */
function writeVisualSalesCharts_(sh, dataSh, layout) {
  const year = Number(Utilities.formatDate(getSpreadsheetToday_(), AD.TZ, 'yyyy'));
  const orders = buildSalesPivotForYear_(year, 'periodUnits');
  const kenpYears = getSalesYearsFromHistory_();
  const years = kenpYears.length ? kenpYears : [year];
  const kenpByDate = buildKenpByWeekEnding_(years);

  // Rank block used cols A–D on data sheet; orders at col 6; KENP at col 20.
  const ordersCol = 6;
  const kenpCol = 20;
  const dataRow = 1;

  const orderHeaders = ['Week Ending'].concat(orders.bookTitles);
  const orderWidth = Math.max(2, orderHeaders.length);
  dataSh.getRange(dataRow, ordersCol, 1, orderWidth).setValues([
    orderHeaders.length > 1 ? orderHeaders : ['Week Ending', 'Orders']
  ]);

  let orderRows = 0;
  if (orders.weeks.length && orders.bookTitles.length) {
    const body = orders.weeks.map(wk => {
      const row = [orders.weekDates[wk]];
      orders.bookTitles.forEach(t => row.push((orders.matrix[wk] && orders.matrix[wk][t]) || 0));
      return row;
    });
    dataSh.getRange(dataRow + 1, ordersCol, body.length, orderWidth).setValues(body);
    dataSh.getRange(dataRow + 1, ordersCol, body.length, 1).setNumberFormat('m/d/yyyy');
    if (orderWidth > 1) {
      dataSh.getRange(dataRow + 1, ordersCol + 1, body.length, orderWidth - 1).setNumberFormat('#,##0');
    }
    orderRows = body.length;
  }

  const kenpHeaders = ['Week Ending'].concat(years.map(y => String(y) + ' KENP'));
  const kenpWidth = Math.max(2, kenpHeaders.length);
  dataSh.getRange(dataRow, kenpCol, 1, kenpWidth).setValues([kenpHeaders]);

  let kenpRows = 0;
  if (kenpByDate.weeks.length) {
    const body = kenpByDate.weeks.map(wk => {
      const row = [kenpByDate.weekDates[wk]];
      years.forEach(y => {
        const v = kenpByDate.matrix[wk] && kenpByDate.matrix[wk][y];
        row.push(v == null || v === '' ? '' : v);
      });
      return row;
    });
    dataSh.getRange(dataRow + 1, kenpCol, body.length, kenpWidth).setValues(body);
    dataSh.getRange(dataRow + 1, kenpCol, body.length, 1).setNumberFormat('m/d/yyyy');
    if (kenpWidth > 1) {
      dataSh.getRange(dataRow + 1, kenpCol + 1, body.length, kenpWidth - 1).setNumberFormat('#,##0');
    }
    kenpRows = body.length;
  }

  if (orderRows >= 1 && orders.bookTitles.length) {
    const ordersRange = dataSh.getRange(dataRow, ordersCol, orderRows + 1, orderWidth);
    let builder = sh.newChart()
      .asColumnChart()
      .addRange(ordersRange)
      .setStacked()
      .setTitle('Orders by week — ' + year)
      .setXAxisTitle('Week Ending')
      .setYAxisTitle('Orders (units)')
      .setNumHeaders(1)
      .setLegendPosition(Charts.Position.BOTTOM)
      .setPosition(layout.row1, layout.rightCol, 0, 0);
    builder = applyVisualChartTheme_(builder, layout);
    sh.insertChart(builder.build());
  }

  if (kenpRows >= 1) {
    const kenpRange = dataSh.getRange(dataRow, kenpCol, kenpRows + 1, kenpWidth);
    let builder = sh.newChart()
      .asLineChart()
      .addRange(kenpRange)
      .setTitle('KENP read by week ending' + (years.length > 1 ? ' (by year)' : ''))
      .setXAxisTitle('Week Ending')
      .setYAxisTitle('KENP read')
      .setNumHeaders(1)
      .setLegendPosition(Charts.Position.BOTTOM)
      .setOption('curveType', 'function')
      .setOption('pointSize', 5)
      .setPosition(layout.row2, layout.leftCol, 0, 0);
    builder = applyVisualChartTheme_(builder, layout);
    sh.insertChart(builder.build());
  }
}
