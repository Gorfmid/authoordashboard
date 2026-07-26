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
    metrics.push('No Meta sync yet — run meta/sync_meta_insights.py then Upload Meta Insights CSV');
  }

  dash.getRange(4, 2, metrics.length, 1).setValues(metrics.map(x => [x]));

  // Catalog Performance starts at column F (leave D–E as spacer).
  // Columns: Book | Stage | Units | KENP | Royalties | Best Rank
  clearBlockUnmerged_(dash, 3, 4, 200, 10);
  mergeRowSafe_(dash, 3, 6, 6)
    .setValue('Catalog Performance')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBackground('#1f4e78')
    .setFontColor('#ffffff');
  dash.getRange(4, 6, 1, 6).setValues([[
    'Book', 'Stage', 'Units', 'KENP Read', 'Royalties (USD)', 'Best Rank'
  ]]);
  styleHeader_(dash.getRange(4, 6, 1, 6));

  // Catalog cols: 1 title, 4 stage, 12 units, 13 ku, 14 roy, 11 best rank ever
  const perf = rows
    .map(r => [r[1], r[4], r[12], r[13], r[14], r[11]])
    .sort((a, b) => number_(b[4]) - number_(a[4]));
  if (perf.length) dash.getRange(5, 6, perf.length, 6).setValues(perf);

  dash.getRange('B4:B11').setNumberFormat('#,##0');
  dash.getRange('B12').setNumberFormat('$#,##0.00');
  dash.getRange('B13').setNumberFormat('#,##0');
  dash.getRange('B14').setNumberFormat('0.0');
  dash.getRange('B15:B16').setNumberFormat('#,##0');
  dash.getRange('B17').setNumberFormat('m/d/yyyy');
  dash.getRange('B20').setNumberFormat('m/d/yyyy hh:mm:ss');
  dash.getRange('H5:I').setNumberFormat('#,##0');
  dash.getRange('J5:J').setNumberFormat('$#,##0.00');
  dash.getRange('K5:K').setNumberFormat('#,##0');

  const metricsEndRow = 3 + metrics.length;
  const kuEndRow = writeKuEstimatesBlock_(dash, metricsEndRow + 2);
  const catalogEndRow = perf.length ? (4 + perf.length) : 4;
  const categoryStartRow = Math.max(kuEndRow, catalogEndRow) + 2;

  clearBlockUnmerged_(dash, Math.max(kuEndRow + 1, 22), 1, 250, 5);

  // Remove any leftover charts from the numbers Dashboard (charts live on Visual Dashboard).
  dash.getCharts().forEach(c => {
    try { dash.removeChart(c); } catch (e) {}
  });

  refreshCategoryRankTable_(dash, categoryStartRow);
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
    'Rank Trend (lower is better)',
    'Last Meta Sync',
    'Meta Sync Status'
  ];
  const current = sheet.getRange(4, 1, labels.length, 1).getValues().map(r => clean_(r[0]));
  const needs = labels.some((label, i) => normalizeKey_(current[i] || '') !== normalizeKey_(label));
  if (needs) {
    sheet.getRange(4, 1, labels.length, 1).setValues(labels.map(x => [x]));
  }
  // Title spans through Catalog Performance (column K).
  try {
    sheet.getRange('A1:K1').getMergedRanges().forEach(m => {
      try { m.breakApart(); } catch (e) {}
    });
  } catch (e) {}
  sheet.getRange('A1:K1').merge()
    .setValue('Author Portfolio Dashboard')
    .setFontSize(22).setFontWeight('bold').setHorizontalAlignment('center')
    .setBackground('#1f4e78').setFontColor('#ffffff');
  sheet.setRowHeight(1, 46);
  [225, 220, 30, 40, 40, 220, 110, 90, 100, 110, 100].forEach((w, i) => {
    if (sheet.getColumnWidth(i + 1) < w) sheet.setColumnWidth(i + 1, w);
  });
}

/**
 * KU Estimates block from latest Royalty Periods row + portfolio KENPC rules.
 * Returns last row written.
 */
function writeKuEstimatesBlock_(dash, startRow) {
  const period = getLatestRoyaltyPeriod_();
  const kenpcInfo = getPortfolioKenpcForCalc_();
  const note =
    'Calculated from estimated KENP royalties divided by KENP read for the selected reporting period. ' +
    'The final KDP Select payout rate may change when Amazon finalizes the month.';

  mergeRowSafe_(dash, startRow, 1, 3)
    .setValue('KU Estimates')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBackground('#1f4e78')
    .setFontColor('#ffffff');

  let rate = null;
  let totalKenp = 0;
  let kenpRoy = 0;
  let ebook = 0;
  let print = 0;
  let total = 0;
  let periodLabel = '';
  if (period) {
    totalKenp = number_(period.totalKenp);
    kenpRoy = number_(period.kenpRoyalties);
    ebook = number_(period.ebookRoyalties);
    print = number_(period.printRoyalties);
    total = number_(period.totalRoyalties);
    if (period.ratePerKenp !== '' && period.ratePerKenp != null) rate = number_(period.ratePerKenp);
    else if (totalKenp > 0 && kenpRoy > 0) rate = kenpRoy / totalKenp;
    const s = isValidDate_(period.periodStart) ? dateKey_(startOfDay_(new Date(period.periodStart))) : '';
    const e = isValidDate_(period.periodEnd) ? dateKey_(startOfDay_(new Date(period.periodEnd))) : '';
    periodLabel = (s && e) ? (s + ' → ' + e) : (e || s || '');
  }

  const useKenpc = kenpcInfo.useForPortfolio;
  const kenpc = useKenpc ? kenpcInfo.kenpc : null;
  const fullRead = (useKenpc && rate != null) ? kenpc * rate : null;
  const equiv = (useKenpc && totalKenp > 0) ? totalKenp / kenpc : null;
  const mixE = total > 0 ? ebook / total : null;
  const mixP = total > 0 ? print / total : null;
  const mixK = total > 0 ? kenpRoy / total : null;
  const cents = rate != null ? rate * 100 : null;

  const rows = [
    ['Reporting Period', periodLabel || '(upload a KDP report)'],
    ['Estimated KENP Royalty Rate', rate != null ? rate : ''],
    ['(cents per KENP)', cents != null ? cents : ''],
    ['KU Royalty per 1,000 Pages', rate != null ? rate * 1000 : ''],
    ['Estimated Full KU Read Royalty', fullRead != null ? fullRead : ''],
    ['KU Equivalent Reads', equiv != null ? equiv : ''],
    ['Royalty Mix — eBook', mixE != null ? mixE : ''],
    ['Royalty Mix — Print', mixP != null ? mixP : ''],
    ['Royalty Mix — KENP', mixK != null ? mixK : ''],
    ['Period eBook / Print / KENP $', total > 0
      ? ('$' + ebook.toFixed(2) + ' / $' + print.toFixed(2) + ' / $' + kenpRoy.toFixed(2))
      : ''],
    ['Royalty Estimate Status', period ? (period.status || 'estimated') : '']
  ];

  const bodyStart = startRow + 1;
  dash.getRange(bodyStart, 1, rows.length, 2).setValues(rows);
  dash.getRange(bodyStart, 1, rows.length, 1).setFontWeight('bold');

  dash.getRange(bodyStart + 1, 2).setNumberFormat('$#,##0.00000');
  dash.getRange(bodyStart + 2, 2).setNumberFormat('0.000" cents"');
  dash.getRange(bodyStart + 3, 2).setNumberFormat('$#,##0.00');
  dash.getRange(bodyStart + 4, 2).setNumberFormat('$#,##0.00');
  dash.getRange(bodyStart + 5, 2).setNumberFormat('0.00');
  dash.getRange(bodyStart + 6, 2, bodyStart + 8, 2).setNumberFormat('0.00%');

  const noteRow = bodyStart + rows.length;
  dash.getRange(noteRow, 1, 1, 3).merge()
    .setValue(note)
    .setFontSize(9)
    .setFontColor('#555555')
    .setWrap(true)
    .setBackground('#FFF8E7');
  dash.setRowHeight(noteRow, 48);
  dash.getRange(bodyStart + 1, 1).setNote(note);

  if (!useKenpc) {
    dash.getRange(bodyStart + 4, 2).setNote(
      'Enter KENPC on Manual Entry for exactly one published book to populate full-read and equivalent-read portfolio metrics.'
    );
  }

  return noteRow;
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

/**
 * Charts-only sheet, tab order: immediately after Dashboard.
 * Rank chart uses a readable "rank score" (higher = better) plus a real-rank table.
 */
function refreshVisualDashboard_() {
  const sh = ensureVisualDashboard_();
  prepareSheetForRebuild_(sh);
  sh.clear();
  sh.getCharts().forEach(c => {
    try { sh.removeChart(c); } catch (e) {}
  });

  styleReportTitleRow_(sh, 'Visual Dashboard — ranks, orders, KENP', 8);
  setBannerRow_(
    sh,
    2,
    8,
    'Charts only. Numbers / Catalog Performance live on Dashboard. ' +
      'Rank chart uses Rank Score = (max rank in view − Amazon rank). Higher score = better. ' +
      'Real Amazon ranks are in the table on the right.',
    { background: '#FFF8E7', fontSize: 10, rowHeight: 48 }
  );

  const seriesByFormat = getOverallRankSeriesByFormat_();
  writeVisualRankSection_(sh, seriesByFormat);
  writeVisualSalesCharts_(sh);
}

function writeVisualRankSection_(sh, data) {
  const tableCol = 1; // A — real ranks (readable)
  const scoreCol = 10; // J — score data for chart (hidden)
  const dataStartRow = 4;

  sh.getRange(dataStartRow, tableCol, 1, 4).setValues([[
    'Snapshot Date',
    'eBook Rank',
    'Paperback Rank',
    'Hardcover Rank'
  ]]);
  styleHeader_(sh.getRange(dataStartRow, tableCol, 1, 4));

  // Build score: higher = better. ceiling = max observed rank across series.
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

  sh.getRange(dataStartRow, scoreCol, 1, 4).setValues([[
    'Snapshot Date',
    'eBook Score',
    'Paperback Score',
    'Hardcover Score'
  ]]);

  if (data.dates && data.dates.length) {
    const actual = data.dates.map((d, i) => [
      d,
      data.series.ebook[i] === '' || data.series.ebook[i] == null ? '' : number_(data.series.ebook[i]),
      data.series.paperback[i] === '' || data.series.paperback[i] == null ? '' : number_(data.series.paperback[i]),
      data.series.hardcover[i] === '' || data.series.hardcover[i] == null ? '' : number_(data.series.hardcover[i])
    ]);
    const scores = data.dates.map((d, i) => [
      d,
      toScore(data.series.ebook[i]),
      toScore(data.series.paperback[i]),
      toScore(data.series.hardcover[i])
    ]);
    sh.getRange(dataStartRow + 1, tableCol, actual.length, 4).setValues(actual);
    sh.getRange(dataStartRow + 1, tableCol, actual.length, 1).setNumberFormat('m/d/yyyy');
    sh.getRange(dataStartRow + 1, tableCol + 1, actual.length, 3).setNumberFormat('#,##0');
    sh.getRange(dataStartRow + 1, scoreCol, scores.length, 4).setValues(scores);
    sh.getRange(dataStartRow + 1, scoreCol, scores.length, 1).setNumberFormat('m/d/yyyy');
    sh.getRange(dataStartRow + 1, scoreCol + 1, scores.length, 3).setNumberFormat('#,##0');
  }

  try { sh.hideColumns(scoreCol, 4); } catch (e) {}

  sh.getRange(dataStartRow, 6).setValue('← Real Amazon ranks (lower # = better)')
    .setFontStyle('italic').setFontColor('#555555');

  if (!data.dates || data.dates.length < 2) return;

  const dataRange = sh.getRange(dataStartRow, scoreCol, data.dates.length + 1, 4);
  const chart = sh.newChart()
    .asLineChart()
    .addRange(dataRange)
    .setTitle('Overall rank trend — higher score = better (improving moves up)')
    .setXAxisTitle('Snapshot Date')
    .setYAxisTitle('Rank score (higher = better)')
    .setNumHeaders(1)
    .setLegendPosition(Charts.Position.BOTTOM)
    .setOption('curveType', 'function')
    .setOption('pointSize', 5)
    .setOption('width', 780)
    .setOption('height', 340)
    .setOption('vAxes', { 0: { title: 'Rank score (higher = better)', format: '#,##0' } })
    .setPosition(4, 6, 0, 0)
    .build();
  sh.insertChart(chart);
}

function writeVisualSalesCharts_(sh) {
  const year = Number(Utilities.formatDate(getSpreadsheetToday_(), AD.TZ, 'yyyy'));
  const orders = buildSalesPivotForYear_(year, 'periodUnits');
  const kenpYears = getSalesYearsFromHistory_();
  const years = kenpYears.length ? kenpYears : [year];
  const kenpOverlay = buildKenpByYearOverlay_(years);

  const ordersCol = 20;
  const kenpCol = 34;
  const ordersDataRow = 4;
  const kenpDataRow = 4;

  clearBlockUnmerged_(sh, ordersDataRow, ordersCol, 120, 12);
  clearBlockUnmerged_(sh, kenpDataRow, kenpCol, 80, 8);

  const orderHeaders = ['Week Ending'].concat(orders.bookTitles);
  const orderWidth = Math.max(2, orderHeaders.length);
  sh.getRange(ordersDataRow, ordersCol, 1, orderWidth).setValues([
    orderHeaders.length > 1 ? orderHeaders : ['Week Ending', 'Orders']
  ]);

  let orderRows = 0;
  if (orders.weeks.length && orders.bookTitles.length) {
    const body = orders.weeks.map(wk => {
      const row = [orders.weekDates[wk]];
      orders.bookTitles.forEach(t => row.push((orders.matrix[wk] && orders.matrix[wk][t]) || 0));
      return row;
    });
    sh.getRange(ordersDataRow + 1, ordersCol, body.length, orderWidth).setValues(body);
    sh.getRange(ordersDataRow + 1, ordersCol, body.length, 1).setNumberFormat('m/d/yyyy');
    if (orderWidth > 1) {
      sh.getRange(ordersDataRow + 1, ordersCol + 1, body.length, orderWidth - 1).setNumberFormat('#,##0');
    }
    orderRows = body.length;
  }

  const kenpHeaders = ['Week of Year'].concat(years.map(y => String(y) + ' KENP'));
  const kenpWidth = Math.max(2, kenpHeaders.length);
  sh.getRange(kenpDataRow, kenpCol, 1, kenpWidth).setValues([kenpHeaders]);

  let kenpRows = 0;
  if (kenpOverlay.weeks.length) {
    const body = kenpOverlay.weeks.map(wk => {
      const row = [wk];
      years.forEach(y => row.push((kenpOverlay.matrix[wk] && kenpOverlay.matrix[wk][y]) || 0));
      return row;
    });
    sh.getRange(kenpDataRow + 1, kenpCol, body.length, kenpWidth).setValues(body);
    if (kenpWidth > 1) {
      sh.getRange(kenpDataRow + 1, kenpCol + 1, body.length, kenpWidth - 1).setNumberFormat('#,##0');
    }
    kenpRows = body.length;
  }

  try { sh.hideColumns(ordersCol, 12); } catch (e) {}
  try { sh.hideColumns(kenpCol, 8); } catch (e) {}

  if (orderRows >= 1 && orders.bookTitles.length) {
    const ordersRange = sh.getRange(ordersDataRow, ordersCol, orderRows + 1, orderWidth);
    sh.insertChart(
      sh.newChart()
        .asColumnChart()
        .addRange(ordersRange)
        .setStacked()
        .setTitle('Orders by week — ' + year)
        .setXAxisTitle('Week Ending')
        .setYAxisTitle('Orders (units)')
        .setNumHeaders(1)
        .setLegendPosition(Charts.Position.BOTTOM)
        .setOption('width', 780)
        .setOption('height', 300)
        .setPosition(24, 1, 0, 0)
        .build()
    );
  }

  if (kenpRows >= 1) {
    const kenpRange = sh.getRange(kenpDataRow, kenpCol, kenpRows + 1, kenpWidth);
    sh.insertChart(
      sh.newChart()
        .asLineChart()
        .addRange(kenpRange)
        .setTitle('KENP read by week of year' + (years.length > 1 ? ' (line per year)' : ''))
        .setXAxisTitle('Week of year')
        .setYAxisTitle('KENP read')
        .setNumHeaders(1)
        .setLegendPosition(Charts.Position.BOTTOM)
        .setOption('curveType', 'function')
        .setOption('pointSize', 4)
        .setOption('width', 780)
        .setOption('height', 300)
        .setPosition(42, 1, 0, 0)
        .build()
    );
  }
}
