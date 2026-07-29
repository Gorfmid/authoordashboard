/** Sales History shading + yearly sales report sheets + Year over Year. */

function refreshSalesReports_() {
  ensureSalesHistorySchema_();
  formatSalesHistoryShading_();
  rebuildAllSalesYearSheets_();
  rebuildYearOverYearSheet_();
}

function formatSalesHistoryShading_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(AD.SHEETS.SALES);
  if (!sh || sh.getLastRow() < 2) return;

  const lastRow = sh.getLastRow();
  const cols = AD.SALES_HEADERS.length;
  const values = sh.getRange(2, 1, lastRow - 1, cols).getValues();
  const backgrounds = [];
  const palette = AD.BOOK_SHADES;
  let colorIdx = 0;
  let prevBook = null;
  let prevColor = null;

  values.forEach(r => {
    // Prefer stable Book ID for shading groups.
    const bookKey = clean_(r[2]) || clean_(r[4]) || '__blank__';
    if (bookKey !== prevBook) {
      let next = colorIdx % palette.length;
      if (palette[next] === prevColor) next = (next + 1) % palette.length;
      colorIdx = next + 1;
      prevColor = palette[next];
      prevBook = bookKey;
    }
    backgrounds.push(new Array(cols).fill(prevColor));
  });

  sh.getRange(2, 1, backgrounds.length, cols).setBackgrounds(backgrounds);
}

function rebuildAllSalesYearSheets_() {
  const years = getSalesYearsFromHistory_();
  if (!years.length) {
    const y = Number(Utilities.formatDate(getSpreadsheetToday_(), AD.TZ, 'yyyy'));
    ensureSalesYearSheet_(y);
    rebuildSalesYearSheet_(y);
    return;
  }
  years.forEach(y => {
    ensureSalesYearSheet_(y);
    rebuildSalesYearSheet_(y);
  });
}

function getSalesYearsFromHistory_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(AD.SHEETS.SALES);
  if (!sh || sh.getLastRow() < 2) return [];
  const years = new Set();
  sh.getRange(2, 2, sh.getLastRow() - 1, 1).getValues().forEach(r => {
    if (!isValidDate_(r[0])) return;
    years.add(Number(Utilities.formatDate(new Date(r[0]), AD.TZ, 'yyyy')));
  });
  return [...years].sort((a, b) => a - b);
}

function ensureSalesYearSheet_(year) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const name = AD.salesYearSheetName(year);
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

/**
 * Sales YYYY layout (Phase 0):
 * 1) Legend — lifetime vs period vs snapshot change
 * 2) Units since previous snapshot (period) + stacked chart
 * 3) Lifetime cumulative units at each week
 * 4) KENP since previous snapshot + chart
 * 5) Lifetime cumulative KENP
 */
function rebuildSalesYearSheet_(year) {
  const sh = ensureSalesYearSheet_(year);
  const periodUnits = buildSalesPivotForYear_(year, 'periodUnits');
  const lifeUnits = buildSalesPivotForYear_(year, 'lifetimeUnits');
  const periodKenp = buildSalesPivotForYear_(year, 'periodKenp');
  const lifeKenp = buildSalesPivotForYear_(year, 'lifetimeKenp');
  const books = periodUnits.bookTitles.length
    ? periodUnits.bookTitles
    : (lifeUnits.bookTitles.length ? lifeUnits.bookTitles : getBookTitleOrder_());
  const weeks = uniqueSortedWeekKeys_([
    periodUnits.weeks,
    lifeUnits.weeks,
    periodKenp.weeks,
    lifeKenp.weeks
  ]);
  const weekDates = Object.assign(
    {},
    lifeKenp.weekDates,
    periodKenp.weekDates,
    lifeUnits.weekDates,
    periodUnits.weekDates
  );
  const numCols = Math.max(2, books.length + 1);

  prepareSheetForRebuild_(sh);
  sh.clear();
  sh.getCharts().forEach(c => {
    try { sh.removeChart(c); } catch (e) {}
  });

  styleReportTitleRow_(sh, 'Sales Report ' + year + ' — period vs lifetime (units & KENP)', numCols);

  setBannerRow_(
    sh,
    2,
    numCols,
    'LEGEND — Lifetime = cumulative total at snapshot. ' +
      '“Since Prev Snapshot” = change since the previous Sales History row for that listing (not a KDP calendar week). ' +
      'Week Ending uses the report/snapshot week date, not “today” alone. ' +
      'First snapshot period change = 0 so lifetime is not mistaken for weekly sales.',
    { background: '#FFF8E7', fontColor: '#333333', fontSize: 10, rowHeight: 54 }
  );

  // Tables only here — Orders / KENP charts live on the Dashboard.
  let row = 4;
  row = writePivotSection_(sh, row, numCols, books, weeks, weekDates, periodUnits.matrix,
    'A) Orders (units) since previous week’s lifetime total — NOT lifetime',
    '',
    false);
  row += 2;
  row = writePivotSection_(sh, row, numCols, books, weeks, weekDates, lifeUnits.matrix,
    'B) Lifetime cumulative units (at each week-ending snapshot)',
    '',
    false);
  row += 2;
  row = writePivotSection_(sh, row, numCols, books, weeks, weekDates, periodKenp.matrix,
    'C) KENP read since previous week’s lifetime total — NOT lifetime',
    '',
    false);
  row += 2;
  writePivotSection_(sh, row, numCols, books, weeks, weekDates, lifeKenp.matrix,
    'D) Lifetime cumulative KENP (at each week-ending snapshot)',
    '',
    false);

  for (let c = 1; c <= numCols; c++) sh.autoResizeColumn(c);
}

function uniqueSortedWeekKeys_(lists) {
  const set = new Set();
  lists.forEach(arr => (arr || []).forEach(k => set.add(k)));
  return [...set].sort();
}

/**
 * Write a labeled pivot block. Returns next free row index.
 * @param {boolean} withChart
 */
function writePivotSection_(sh, startRow, numCols, books, weeks, weekDates, matrix, title, chartTitle, withChart) {
  const headers = ['Week Ending'].concat(books);
  setBannerRow_(sh, startRow, numCols, title, {
    background: '#D6E3F0',
    fontColor: '#1f4e78',
    bold: true,
    fontSize: 11
  });
  const headerRow = startRow + 1;
  sh.getRange(headerRow, 1, 1, headers.length).setValues([headers]);
  styleHeader_(sh.getRange(headerRow, 1, 1, headers.length));

  if (!weeks.length || !books.length) {
    sh.getRange(headerRow + 1, 1).setValue('No snapshot data for this section yet.');
    return headerRow + 2;
  }

  const body = weeks.map(weekKey => {
    const row = [weekDates[weekKey] || weekKey];
    books.forEach(title => {
      const m = matrix[weekKey] || {};
      const v = m[title];
      row.push(v == null ? 0 : v);
    });
    return row;
  });

  sh.getRange(headerRow + 1, 1, body.length, headers.length).setValues(body);
  sh.getRange(headerRow + 1, 1, body.length, 1).setNumberFormat('m/d/yyyy');
  if (headers.length > 1) {
    sh.getRange(headerRow + 1, 2, body.length, headers.length - 1).setNumberFormat('#,##0');
  }

  if (withChart && chartTitle) {
    const dataRange = sh.getRange(headerRow, 1, body.length + 1, headers.length);
    const chart = sh.newChart()
      .asColumnChart()
      .addRange(dataRange)
      .setStacked()
      .setTitle(chartTitle)
      .setXAxisTitle('Week Ending')
      .setYAxisTitle('Amount')
      .setNumHeaders(1)
      .setLegendPosition(Charts.Position.BOTTOM)
      .setOption('width', 900)
      .setOption('height', 380)
      .setPosition(headerRow, Math.max(headers.length + 2, 6), 0, 0)
      .build();
    sh.insertChart(chart);
  }

  return headerRow + 1 + body.length;
}

/**
 * Pivot builder.
 * kind: periodUnits | lifetimeUnits | periodKenp | lifetimeKenp
 *
 * Lifetime = sum of listing lifetimes at each week (latest snapshot per listing/week).
 * Period (Orders / KENP read) = week-over-week change in that book lifetime total
 * (avoids double-counting duplicate same-week rows and inflated first-snapshot period columns).
 */
function buildSalesPivotForYear_(year, kind) {
  const bookOrder = getBookTitleOrder_();
  const life = buildBookLifetimeByWeek_(kind === 'periodKenp' || kind === 'lifetimeKenp' ? 'kenp' : 'units');
  if (!life.weeks.length) {
    return { weeks: [], weekDates: {}, bookTitles: bookOrder.slice(), matrix: {} };
  }

  const yearWeeks = life.weeks.filter(wk => {
    const d = life.weekDates[wk];
    return d && Number(Utilities.formatDate(d, AD.TZ, 'yyyy')) === year;
  });

  const seenBooks = new Set();
  yearWeeks.forEach(wk => {
    Object.keys(life.matrix[wk] || {}).forEach(t => seenBooks.add(normalizeKey_(t)));
  });
  const ordered = orderBookTitles_(
    bookOrder,
    seenBooks,
    [...seenBooks].map(k => {
      // Recover original title casing from matrix
      for (let i = 0; i < life.weeks.length; i++) {
        const titles = Object.keys(life.matrix[life.weeks[i]] || {});
        const hit = titles.find(t => normalizeKey_(t) === k);
        if (hit) return hit;
      }
      return k;
    })
  );

  const wantPeriod = kind === 'periodUnits' || kind === 'periodKenp';
  const matrix = {};
  yearWeeks.forEach(wk => {
    matrix[wk] = {};
    ordered.forEach(title => {
      const cur = (life.matrix[wk] && life.matrix[wk][title]) || 0;
      if (!wantPeriod) {
        matrix[wk][title] = cur;
        return;
      }
      const idx = life.weeks.indexOf(wk);
      const prevKey = idx > 0 ? life.weeks[idx - 1] : null;
      const prev = prevKey && life.matrix[prevKey] ? (life.matrix[prevKey][title] || 0) : 0;
      matrix[wk][title] = Math.max(0, cur - prev);
    });
  });

  return {
    weeks: yearWeeks,
    weekDates: life.weekDates,
    bookTitles: ordered,
    matrix: matrix
  };
}

/**
 * Latest snapshot per listing per week-ending, then sum lifetime by Book ID
 * (display labels use title). Title is never the relationship key.
 * metric: 'units' | 'kenp'
 */
function buildBookLifetimeByWeek_(metric) {
  const sales = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(AD.SHEETS.SALES);
  const weekDates = {};
  const matrix = {}; // weekKey -> displayTitle -> amount
  const idToTitle = {};
  if (!sales || sales.getLastRow() < 2) {
    return { weeks: [], weekDates: {}, matrix: {} };
  }

  const lifeCol = metric === 'kenp' ? 10 : 8;
  // listing|weekKey -> { snap, bookId, title, amount, week, weekKey }
  const best = new Map();

  sales.getRange(2, 1, sales.getLastRow() - 1, AD.SALES_HEADERS.length).getValues().forEach(r => {
    if (!isValidDate_(r[1]) || !isValidDate_(r[0])) return;
    const listing = clean_(r[3]);
    if (!listing) return;
    const week = startOfDay_(new Date(r[1]));
    const snap = startOfDay_(new Date(r[0]));
    const weekKey = dateKey_(week);
    const bookId = clean_(r[2]);
    const title = clean_(r[4]) || bookId || 'Unknown';
    const groupKey = bookId || ('TITLE:' + normalizeKey_(title));
    const amount = number_(r[lifeCol]);
    const key = listing + '|' + weekKey;
    const prev = best.get(key);
    if (!prev || snap > prev.snap) {
      best.set(key, {
        snap: snap,
        groupKey: groupKey,
        title: title,
        amount: amount,
        week: week,
        weekKey: weekKey
      });
    }
    if (bookId) idToTitle[bookId] = title;
  });

  // Aggregate by book id, but expose matrix keys as display titles for charts.
  const byWeekId = {}; // weekKey -> groupKey -> amount
  best.forEach(item => {
    if (!byWeekId[item.weekKey]) byWeekId[item.weekKey] = {};
    byWeekId[item.weekKey][item.groupKey] =
      (byWeekId[item.weekKey][item.groupKey] || 0) + item.amount;
    weekDates[item.weekKey] = item.week;
    if (!idToTitle[item.groupKey]) idToTitle[item.groupKey] = item.title;
  });

  Object.keys(byWeekId).forEach(weekKey => {
    matrix[weekKey] = {};
    Object.keys(byWeekId[weekKey]).forEach(gid => {
      const label = idToTitle[gid] || gid;
      matrix[weekKey][label] = (matrix[weekKey][label] || 0) + byWeekId[weekKey][gid];
    });
  });

  const weeks = Object.keys(matrix).sort();
  return { weeks: weeks, weekDates: weekDates, matrix: matrix };
}

function orderBookTitles_(bookOrder, seenBooks, bookTitles) {
  const ordered = [];
  const used = new Set();
  bookOrder.forEach(t => {
    if (seenBooks.has(normalizeKey_(t))) {
      ordered.push(t);
      used.add(normalizeKey_(t));
    }
  });
  bookTitles.forEach(t => {
    if (!used.has(normalizeKey_(t))) ordered.push(t);
  });
  return ordered;
}

function getBookTitleOrder_() {
  const rows = getInputRows_();
  const map = new Map();
  rows.forEach(r => {
    const id = clean_(r[AD.COL.BOOK_ID]);
    const title = clean_(r[AD.COL.TITLE]);
    if (!title) return;
    const key = id || normalizeKey_(title);
    if (!map.has(key)) {
      map.set(key, {
        title: title,
        series: clean_(r[3]),
        num: number_(r[4])
      });
    }
  });
  return [...map.values()]
    .sort((a, b) => a.series.localeCompare(b.series) || a.num - b.num || a.title.localeCompare(b.title))
    .map(b => b.title);
}

function ensureYearOverYearSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(AD.SHEETS.YOY);
  if (!sh) sh = ss.insertSheet(AD.SHEETS.YOY);
  return sh;
}

/** Title styling without merge — avoids freeze-column errors on merged cells. */
function styleReportTitleRow_(sh, title, numCols) {
  const cols = Math.max(2, numCols || 2);
  try {
    sh.getRange(1, 1, 1, cols).getMergedRanges().forEach(m => {
      try { m.breakApart(); } catch (e) {}
    });
  } catch (e) {}
  sh.getRange(1, 1, 1, cols)
    .setBackground('#1f4e78')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(16)
    .clearContent();
  sh.getRange(1, 1)
    .setValue(title)
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');
  sh.setRowHeight(1, 36);
}

function rebuildYearOverYearSheet_() {
  const sh = ensureYearOverYearSheet_();
  const years = getSalesYearsFromHistory_();
  const totals = buildYearOverYearTotals_(years);
  const kenpOverlay = buildKenpByYearOverlay_(years);
  const width = Math.max(2, 1 + years.length * 2);

  prepareSheetForRebuild_(sh);
  sh.clear();
  sh.getCharts().forEach(c => {
    try { sh.removeChart(c); } catch (e) {}
  });

  styleReportTitleRow_(
    sh,
    'Year over Year — period units / royalties + KENP by year',
    width
  );

  setBannerRow_(
    sh,
    2,
    width,
    'Period Units / Royalties for a year = change in lifetime from the last snapshot before that year ' +
      'to the latest snapshot in that year (per listing, then summed by book). ' +
      'This includes sales already present on the first snapshot (which stores period change = 0). ' +
      'KENP overlay below still uses week-of-year for multi-year lines.',
    { background: '#FFF8E7', fontSize: 10, rowHeight: 54 }
  );

  if (!years.length) {
    sh.getRange(4, 1).setValue('No sales history yet.');
    return;
  }

  const headers = ['Book Title'];
  years.forEach(y => {
    headers.push(y + ' Period Units');
    headers.push(y + ' Period Royalties');
  });
  sh.getRange(4, 1, 1, headers.length).setValues([headers]);
  styleHeader_(sh.getRange(4, 1, 1, headers.length));
  // Freeze header row only — never freeze columns while banners span multiple columns.
  sh.setFrozenRows(4);

  const titles = totals.titles;
  if (!titles.length) {
    sh.getRange(5, 1).setValue('No book sales totals yet.');
  } else {
    const body = titles.map(title => {
      const row = [title];
      years.forEach(y => {
        const cell = (totals.byBook[title] && totals.byBook[title][y]) || { units: 0, royalties: 0 };
        row.push(cell.units);
        row.push(cell.royalties);
      });
      return row;
    });

    sh.getRange(5, 1, body.length, headers.length).setValues(body);
    years.forEach((_, i) => {
      const unitsCol = 2 + i * 2;
      const royCol = 3 + i * 2;
      sh.getRange(5, unitsCol, body.length, 1).setNumberFormat('#,##0');
      sh.getRange(5, royCol, body.length, 1).setNumberFormat('$#,##0.00');
    });
  }

  // KENP-by-year data table (chart is on the Dashboard).
  const kenpStart = 5 + Math.max(titles.length, 1) + 3;
  setBannerRow_(
    sh,
    kenpStart,
    Math.max(2, years.length + 1),
    'KENP read data by week of year (chart on Dashboard — one line per calendar year)',
    { background: '#D6E3F0', fontColor: '#1f4e78', bold: true, fontSize: 11 }
  );

  const kenpHeaders = ['Week of Year'].concat(years.map(y => String(y) + ' KENP (period)'));
  sh.getRange(kenpStart + 1, 1, 1, kenpHeaders.length).setValues([kenpHeaders]);
  styleHeader_(sh.getRange(kenpStart + 1, 1, 1, kenpHeaders.length));

  if (kenpOverlay.weeks.length && years.length) {
    const kenpBody = kenpOverlay.weeks.map(wk => {
      const row = [wk];
      years.forEach(y => {
        row.push((kenpOverlay.matrix[wk] && kenpOverlay.matrix[wk][y]) || 0);
      });
      return row;
    });
    sh.getRange(kenpStart + 2, 1, kenpBody.length, kenpHeaders.length).setValues(kenpBody);
    if (kenpHeaders.length > 1) {
      sh.getRange(kenpStart + 2, 2, kenpBody.length, kenpHeaders.length - 1).setNumberFormat('#,##0');
    }
  } else {
    sh.getRange(kenpStart + 2, 1).setValue('No KENP period data yet.');
  }

  headers.forEach((_, i) => sh.autoResizeColumn(i + 1));
}

/**
 * Week-ending date × calendar year matrix of period KENP (for Visual Dashboard).
 * X-axis is the actual Week Ending date (not week-of-year number).
 */
function buildKenpByWeekEnding_(years) {
  const matrix = {};
  const weekDates = {};
  if (!years || !years.length) return { weeks: [], weekDates: weekDates, matrix: matrix };

  years.forEach(y => {
    const pivot = buildSalesPivotForYear_(y, 'periodKenp');
    pivot.weeks.forEach(weekKey => {
      const d = pivot.weekDates[weekKey];
      if (!d) return;
      weekDates[weekKey] = d;
      if (!matrix[weekKey]) matrix[weekKey] = {};
      let sum = 0;
      pivot.bookTitles.forEach(t => {
        sum += (pivot.matrix[weekKey] && pivot.matrix[weekKey][t]) || 0;
      });
      matrix[weekKey][y] = sum;
    });
  });

  const weeks = Object.keys(weekDates).sort();
  return { weeks: weeks, weekDates: weekDates, matrix: matrix };
}

/**
 * Week-of-year × calendar year matrix of period KENP (book-level week-over-week).
 * Enables 2026 vs 2027 as separate chart lines when both years exist (YoY sheet).
 */
function buildKenpByYearOverlay_(years) {
  const matrix = {};
  const weekSet = new Set();
  if (!years || !years.length) return { weeks: [], matrix: matrix };

  years.forEach(y => {
    const pivot = buildSalesPivotForYear_(y, 'periodKenp');
    pivot.weeks.forEach(weekKey => {
      const d = pivot.weekDates[weekKey];
      if (!d) return;
      const weekOfYear = Number(Utilities.formatDate(d, AD.TZ, 'w'));
      weekSet.add(weekOfYear);
      if (!matrix[weekOfYear]) matrix[weekOfYear] = {};
      let sum = 0;
      pivot.bookTitles.forEach(t => {
        sum += (pivot.matrix[weekKey] && pivot.matrix[weekKey][t]) || 0;
      });
      matrix[weekOfYear][y] = (matrix[weekOfYear][y] || 0) + sum;
    });
  });

  const weeks = [...weekSet].sort((a, b) => a - b);
  weeks.forEach(wk => {
    years.forEach(y => {
      if (matrix[wk][y] == null) matrix[wk][y] = 0;
    });
  });
  return { weeks: weeks, matrix: matrix };
}

/**
 * YoY period totals from lifetime endpoints (not sum of “Since Prev Snapshot”).
 * First snapshots store period change = 0, so summing period columns undercounts
 * lifetime already present on the first row (e.g. 32 units → showed 24).
 *
 * Per listing, per year Y:
 *   baseline = lifetime on latest snapshot with Week Ending year < Y (else 0)
 *   end = lifetime on latest snapshot with Week Ending year === Y
 *   period = max(0, end − baseline)
 * Then sum listings by book title.
 */
function buildYearOverYearTotals_(years) {
  const byBook = {};
  const titleSet = new Set();
  if (!years || !years.length) {
    return { titles: getBookTitleOrder_(), byBook: byBook };
  }

  const byListing = getSalesLifetimeSeriesByListingForYoy_();
  byListing.forEach(series => {
    if (!series.length) return;
    const title = series[0].title;
    titleSet.add(title);
    if (!byBook[title]) byBook[title] = {};

    years.forEach(y => {
      let baselineU = 0;
      let baselineR = 0;
      let endU = null;
      let endR = null;
      series.forEach(s => {
        if (s.year < y) {
          baselineU = s.units;
          baselineR = s.royalties;
        } else if (s.year === y) {
          endU = s.units;
          endR = s.royalties;
        }
      });
      if (endU == null && endR == null) return;
      if (!byBook[title][y]) byBook[title][y] = { units: 0, royalties: 0 };
      byBook[title][y].units += Math.max(0, number_(endU) - number_(baselineU));
      byBook[title][y].royalties += Math.max(0, number_(endR) - number_(baselineR));
    });
  });

  // Include catalog books with no history yet
  getBookTitleOrder_().forEach(t => titleSet.add(t));

  const ordered = [];
  const used = new Set();
  getBookTitleOrder_().forEach(t => {
    if (titleSet.has(t) || byBook[t]) {
      ordered.push(t);
      used.add(normalizeKey_(t));
    }
  });
  [...titleSet].sort().forEach(t => {
    if (!used.has(normalizeKey_(t))) ordered.push(t);
  });

  return { titles: ordered, byBook: byBook };
}

/**
 * One series per listing: latest snapshot per Week Ending, sorted by week.
 * Points: { week, year, units, royalties, title }
 */
function getSalesLifetimeSeriesByListingForYoy_() {
  const map = new Map(); // listing -> Map(weekKey -> point)
  const sales = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(AD.SHEETS.SALES);
  if (!sales || sales.getLastRow() < 2) return [];

  sales.getRange(2, 1, sales.getLastRow() - 1, AD.SALES_HEADERS.length).getValues().forEach(r => {
    if (!isValidDate_(r[1]) || !clean_(r[3])) return;
    const listing = clean_(r[3]);
    const week = startOfDay_(new Date(r[1]));
    const weekKey = dateKey_(week);
    const snap = isValidDate_(r[0]) ? new Date(r[0]).getTime() : 0;
    const point = {
      week: week,
      year: Number(Utilities.formatDate(week, AD.TZ, 'yyyy')),
      units: number_(r[8]),
      royalties: number_(r[12]),
      title: clean_(r[4]) || clean_(r[2]) || 'Unknown',
      snap: snap
    };
    if (!map.has(listing)) map.set(listing, new Map());
    const weeks = map.get(listing);
    const prev = weeks.get(weekKey);
    if (!prev || snap >= prev.snap) weeks.set(weekKey, point);
  });

  return [...map.values()].map(weekMap =>
    [...weekMap.values()].sort((a, b) => a.week - b.week)
  );
}

function getAutomaticSheetNames_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // Events stays editable (user logs releases/ads). Manual Entry stays editable.
  const names = Object.values(AD.SHEETS).filter(n =>
    n !== AD.SHEETS.INPUT && n !== AD.SHEETS.EVENTS
  );
  ss.getSheets().forEach(sh => {
    const n = sh.getName();
    if (AD.isSalesYearSheetName(n) && names.indexOf(n) === -1) names.push(n);
  });
  return names;
}

function orderReportSheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const base = [
    AD.SHEETS.INPUT,
    AD.SHEETS.DASHBOARD,
    AD.SHEETS.STATISTICS,
    AD.SHEETS.VISUAL,
    AD.SHEETS.CATALOG,
    AD.SHEETS.SALES,
    AD.SHEETS.RANKS,
    AD.SHEETS.MARKETING
  ];
  const years = ss.getSheets()
    .map(s => s.getName())
    .filter(n => AD.isSalesYearSheetName(n))
    .sort();
  // Visible tabs only — setActiveSheet/moveActiveSheet unhides sheets.
  const visibleOrder = base.concat(years).concat([
    AD.SHEETS.YOY,
    AD.SHEETS.EVENTS,
    AD.SHEETS.META_DAILY
  ]);
  visibleOrder.forEach((name, i) => {
    const sh = ss.getSheetByName(name);
    if (sh) {
      ss.setActiveSheet(sh);
      ss.moveActiveSheet(i + 1);
    }
  });

  // Park diagnostics at the end, then force-hide (reorder briefly unhides).
  const diagnostics = getDiagnosticSheetNames_();
  diagnostics.forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) return;
    try {
      ss.setActiveSheet(sh);
      ss.moveActiveSheet(ss.getNumSheets());
    } catch (e) {}
    try { sh.hideSheet(); } catch (e) {}
  });
  hideDiagnosticSheets_();
}
