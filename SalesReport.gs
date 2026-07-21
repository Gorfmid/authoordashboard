/** Sales History shading + yearly sales report sheets + Year over Year. */

function refreshSalesReports_() {
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

function rebuildSalesYearSheet_(year) {
  const sh = ensureSalesYearSheet_(year);
  const pivot = buildWeeklyUnitsPivotForYear_(year);

  sh.clear();
  sh.getCharts().forEach(c => {
    try { sh.removeChart(c); } catch (e) {}
  });

  // No merged title row — merging across columns breaks setFrozenColumns later.
  styleReportTitleRow_(
    sh,
    'Sales Report ' + year + ' — weekly units by book',
    Math.max(2, pivot.bookTitles.length + 1)
  );

  const headers = ['Week Ending'].concat(pivot.bookTitles);
  sh.getRange(3, 1, 1, headers.length).setValues([headers]);
  styleHeader_(sh.getRange(3, 1, 1, headers.length));
  sh.setFrozenRows(3);

  if (!pivot.weeks.length || !pivot.bookTitles.length) {
    sh.getRange(4, 1).setValue('No sales snapshot data for ' + year + ' yet.');
    return;
  }

  const body = pivot.weeks.map(weekKey => {
    const row = [pivot.weekDates[weekKey]];
    pivot.bookTitles.forEach(title => {
      const v = pivot.matrix[weekKey][title];
      row.push(v == null ? 0 : v);
    });
    return row;
  });

  sh.getRange(4, 1, body.length, headers.length).setValues(body);
  sh.getRange(4, 1, body.length, 1).setNumberFormat('m/d/yyyy');
  if (headers.length > 1) {
    sh.getRange(4, 2, body.length, headers.length - 1).setNumberFormat('#,##0');
  }

  const dataRange = sh.getRange(3, 1, body.length + 1, headers.length);
  const chart = sh.newChart()
    .asColumnChart()
    .addRange(dataRange)
    .setStacked()
    .setTitle('Weekly Units by Book — ' + year)
    .setXAxisTitle('Week Ending')
    .setYAxisTitle('Units')
    .setNumHeaders(1)
    .setLegendPosition(Charts.Position.BOTTOM)
    .setOption('width', 900)
    .setOption('height', 420)
    .setPosition(4, Math.max(headers.length + 2, 6), 0, 0)
    .build();
  sh.insertChart(chart);

  headers.forEach((_, i) => sh.autoResizeColumn(i + 1));
}

/**
 * Pivot: week ending -> book title -> sum of Weekly Unit Change (all formats).
 * Book order follows Manual Entry series # / title when available.
 */
function buildWeeklyUnitsPivotForYear_(year) {
  const sales = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(AD.SHEETS.SALES);
  const bookOrder = getBookTitleOrder_();
  const bookTitles = [];
  const seenBooks = new Set();
  const weekDates = {};
  const matrix = {};

  if (!sales || sales.getLastRow() < 2) {
    return { weeks: [], weekDates: {}, bookTitles: bookOrder.slice(), matrix: {} };
  }

  sales.getRange(2, 1, sales.getLastRow() - 1, AD.SALES_HEADERS.length).getValues().forEach(r => {
    if (!isValidDate_(r[1])) return;
    const week = startOfDay_(new Date(r[1]));
    const y = Number(Utilities.formatDate(week, AD.TZ, 'yyyy'));
    if (y !== year) return;
    const title = clean_(r[4]) || clean_(r[2]) || 'Unknown';
    const units = number_(r[9]); // Weekly Unit Change
    const weekKey = dateKey_(week);

    if (!seenBooks.has(normalizeKey_(title))) {
      seenBooks.add(normalizeKey_(title));
      bookTitles.push(title);
    }
    if (!matrix[weekKey]) {
      matrix[weekKey] = {};
      weekDates[weekKey] = week;
    }
    matrix[weekKey][title] = (matrix[weekKey][title] || 0) + units;
  });

  // Stable book column order: Manual Entry order first, then any extras.
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

  const weeks = Object.keys(matrix).sort();
  weeks.forEach(wk => {
    ordered.forEach(t => {
      if (matrix[wk][t] == null) matrix[wk][t] = 0;
    });
  });

  return { weeks: weeks, weekDates: weekDates, bookTitles: ordered, matrix: matrix };
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

  sh.clear();
  styleReportTitleRow_(
    sh,
    'Year over Year — units and royalties by book',
    Math.max(2, 1 + years.length * 2)
  );

  if (!years.length) {
    sh.getRange(3, 1).setValue('No sales history yet.');
    return;
  }

  const headers = ['Book Title'];
  years.forEach(y => {
    headers.push(y + ' Units');
    headers.push(y + ' Royalties');
  });
  sh.getRange(3, 1, 1, headers.length).setValues([headers]);
  styleHeader_(sh.getRange(3, 1, 1, headers.length));
  sh.setFrozenRows(3);
  sh.setFrozenColumns(1);

  const titles = totals.titles;
  if (!titles.length) {
    sh.getRange(4, 1).setValue('No book sales totals yet.');
    return;
  }

  const body = titles.map(title => {
    const row = [title];
    years.forEach(y => {
      const cell = (totals.byBook[title] && totals.byBook[title][y]) || { units: 0, royalties: 0 };
      row.push(cell.units);
      row.push(cell.royalties);
    });
    return row;
  });

  sh.getRange(4, 1, body.length, headers.length).setValues(body);
  years.forEach((_, i) => {
    const unitsCol = 2 + i * 2;
    const royCol = 3 + i * 2;
    sh.getRange(4, unitsCol, body.length, 1).setNumberFormat('#,##0');
    sh.getRange(4, royCol, body.length, 1).setNumberFormat('$#,##0.00');
  });
  headers.forEach((_, i) => sh.autoResizeColumn(i + 1));
}

function buildYearOverYearTotals_(years) {
  const sales = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(AD.SHEETS.SALES);
  const byBook = {};
  const titleSet = new Set();
  if (!sales || sales.getLastRow() < 2) {
    return { titles: getBookTitleOrder_(), byBook: byBook };
  }

  sales.getRange(2, 1, sales.getLastRow() - 1, AD.SALES_HEADERS.length).getValues().forEach(r => {
    if (!isValidDate_(r[1])) return;
    const week = new Date(r[1]);
    const y = Number(Utilities.formatDate(week, AD.TZ, 'yyyy'));
    if (years.indexOf(y) === -1) return;
    const title = clean_(r[4]) || clean_(r[2]) || 'Unknown';
    titleSet.add(title);
    if (!byBook[title]) byBook[title] = {};
    if (!byBook[title][y]) byBook[title][y] = { units: 0, royalties: 0 };
    byBook[title][y].units += number_(r[9]);
    byBook[title][y].royalties += number_(r[13]);
  });

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

function getAutomaticSheetNames_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const names = Object.values(AD.SHEETS).filter(n => n !== AD.SHEETS.INPUT);
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
    AD.SHEETS.CATALOG,
    AD.SHEETS.SALES,
    AD.SHEETS.RANKS,
    AD.SHEETS.MARKETING
  ];
  const years = ss.getSheets()
    .map(s => s.getName())
    .filter(n => AD.isSalesYearSheetName(n))
    .sort();
  const order = base.concat(years).concat([AD.SHEETS.YOY]);
  order.forEach((name, i) => {
    const sh = ss.getSheetByName(name);
    if (sh) {
      ss.setActiveSheet(sh);
      ss.moveActiveSheet(i + 1);
    }
  });
}
