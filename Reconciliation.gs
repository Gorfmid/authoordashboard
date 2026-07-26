/**
 * Phase 0 reconciliation — flag discrepancies; never silently adjust data.
 */

function rebuildReconciliationSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(AD.SHEETS.RECONCILE);
  if (!sh) sh = ss.insertSheet(AD.SHEETS.RECONCILE);

  prepareSheetForRebuild_(sh);
  sh.clear();
  sh.getCharts().forEach(c => {
    try { sh.removeChart(c); } catch (e) {}
  });

  styleReportTitleRow_(sh, 'Reconciliation & data health (Phase 0)', 6);

  // Version block
  sh.getRange(2, 1, 1, 2).setValues([['Workbook / script version', AD.VERSION]]);
  sh.getRange(3, 1, 1, 2).setValues([['Schema version', AD.SCHEMA_VERSION]]);
  sh.getRange(4, 1, 1, 2).setValues([['Script version', AD.SCRIPT_VERSION]]);
  sh.getRange(5, 1, 1, 2).setValues([['Last migration date', AD.LAST_MIGRATION]]);
  sh.getRange(6, 1, 1, 2).setValues([['Migration notes', AD.MIGRATION_NOTES]]);
  sh.getRange(2, 1, 5, 1).setFontWeight('bold');
  sh.getRange(6, 2).setWrap(true);
  sh.setRowHeight(6, 48);

  const report = runSalesReconciliation_();
  sh.getRange(8, 1).setValue('Checks run: ' + Utilities.formatDate(new Date(), AD.TZ, 'yyyy-MM-dd HH:mm:ss'))
    .setFontWeight('bold');

  const headers = [
    'Check',
    'Listing / Book',
    'Expected / Left',
    'Actual / Right',
    'Delta',
    'Status'
  ];
  sh.getRange(10, 1, 1, headers.length).setValues([headers]);
  styleHeader_(sh.getRange(10, 1, 1, headers.length));

  const body = report.rows.length
    ? report.rows
    : [['(no listings to check)', '', '', '', '', 'OK']];

  sh.getRange(11, 1, body.length, headers.length).setValues(body);
  body.forEach((r, i) => {
    if (String(r[5]).indexOf('FLAG') === 0) {
      sh.getRange(11 + i, 1, 1, headers.length).setBackground('#F8E0E0');
    } else if (String(r[5]) === 'OK') {
      sh.getRange(11 + i, 1, 1, headers.length).setBackground('#E8F5E9');
    }
  });

  const summaryRow = 11 + body.length + 2;
  sh.getRange(summaryRow, 1).setValue(
    report.flagCount
      ? ('FLAGS: ' + report.flagCount + ' — review Sales History Lifetime vs Manual Entry and period sums. Data was not auto-corrected.')
      : 'All checks OK (or no comparable data yet).'
  ).setFontWeight('bold').setFontColor(report.flagCount ? '#a12622' : '#0b6b3a');

  setBannerRow_(
    sh,
    summaryRow + 2,
    6,
    'Definitions: Lifetime = cumulative total. Period sum = sum of “Since Prev Snapshot” for a listing. ' +
      'Expected period sum ≈ latest lifetime − first lifetime snapshot. ' +
      'Manual Entry lifetime should match the latest Sales History lifetime for that listing after a snapshot.',
    { background: '#FFF8E7', fontSize: 10, rowHeight: 48 }
  );

  headers.forEach((_, i) => sh.autoResizeColumn(i + 1));

  hideDiagnosticSheets_();
  return report;
}

function runSalesReconciliationMenu() {
  ensureSalesHistorySchema_();
  const report = rebuildReconciliationSheet_();
  lockAutomaticSheets();
  SpreadsheetApp.getUi().alert(
    'Reconciliation complete.\n\nFlags: ' + report.flagCount + '\nRows checked: ' + report.rows.length +
      '\n\nOpen the Reconciliation sheet for details.'
  );
}

/**
 * Compare:
 * - Manual Entry lifetime vs latest Sales History lifetime
 * - Sum of period changes vs (latest lifetime − first lifetime)
 * Per listing for units, KENP, royalties.
 */
function runSalesReconciliation_() {
  const rows = [];
  let flagCount = 0;
  const input = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(AD.SHEETS.INPUT);
  const sales = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(AD.SHEETS.SALES);
  if (!input || !sales || input.getLastRow() < 2) {
    return { rows: rows, flagCount: 0 };
  }

  const series = getSalesSeriesByListing_();
  const inputRows = input.getRange(2, 1, input.getLastRow() - 1, AD.INPUT_HEADERS.length).getValues();

  inputRows.forEach(r => {
    const listing = clean_(r[AD.COL.LISTING_ID]);
    if (!listing) return;
    const label = (clean_(r[AD.COL.TITLE]) || listing) + ' / ' + (clean_(r[AD.COL.FORMAT]) || '');
    const hist = series.get(listing);

    const meUnits = number_(r[AD.COL.UNITS]);
    const meKenp = number_(r[AD.COL.KU]);
    const meRoy = number_(r[AD.COL.ROYALTIES]);

    if (!hist || !hist.length) {
      if (meUnits || meKenp || meRoy) {
        rows.push(['Manual Entry has lifetime but no Sales History', label, meUnits, '(none)', '', 'FLAG']);
        flagCount++;
      }
      return;
    }

    const first = hist[0];
    const last = hist[hist.length - 1];
    const sumPeriodU = hist.reduce((s, x) => s + x.du, 0);
    const sumPeriodK = hist.reduce((s, x) => s + x.dk, 0);
    const sumPeriodR = hist.reduce((s, x) => s + x.dr, 0);
    const expectedU = Math.max(0, last.units - first.units);
    const expectedK = Math.max(0, last.ku - first.ku);
    const expectedR = Math.max(0, last.roy - first.roy);

    flagCount += pushReconcileRow_(rows, 'Latest Sales History Units vs Manual Entry', label, last.units, meUnits, 0);
    flagCount += pushReconcileRow_(rows, 'Latest Sales History KENP vs Manual Entry', label, last.ku, meKenp, 0);
    flagCount += pushReconcileRow_(rows, 'Latest Sales History Royalties vs Manual Entry', label, last.roy, meRoy, 0.02);

    flagCount += pushReconcileRow_(rows, 'Sum period units vs (latest − first lifetime)', label, expectedU, sumPeriodU, 0);
    flagCount += pushReconcileRow_(rows, 'Sum period KENP vs (latest − first lifetime)', label, expectedK, sumPeriodK, 0);
    flagCount += pushReconcileRow_(rows, 'Sum period royalties vs (latest − first lifetime)', label, expectedR, sumPeriodR, 0.02);

    const meEbook = number_(r[AD.COL.ROYALTY_EBOOK]);
    const mePrint = number_(r[AD.COL.ROYALTY_PRINT]);
    const meKenpR = number_(r[AD.COL.ROYALTY_KENP]);
    if (meEbook || mePrint || meKenpR || meRoy) {
      flagCount += pushReconcileRow_(
        rows,
        'eBook + Print + KENP royalties vs Lifetime Royalties',
        label,
        meEbook + mePrint + meKenpR,
        meRoy,
        0.01
      );
    }
  });

  const period = getLatestRoyaltyPeriod_();
  if (period) {
    const pe = number_(period.ebookRoyalties);
    const pp = number_(period.printRoyalties);
    const pk = number_(period.kenpRoyalties);
    const pt = number_(period.totalRoyalties);
    const periodLabel = 'Royalty Periods latest' +
      (period.periodEnd ? ' (end ' + (isValidDate_(period.periodEnd) ? dateKey_(startOfDay_(new Date(period.periodEnd))) : '') + ')' : '');
    flagCount += pushReconcileRow_(
      rows,
      'Period eBook + Print + KENP vs Total Royalties',
      periodLabel,
      pe + pp + pk,
      pt,
      0.01
    );
  }

  return { rows: rows, flagCount: flagCount };
}

function pushReconcileRow_(rows, check, label, left, right, tol) {
  const a = Number(left) || 0;
  const b = Number(right) || 0;
  const delta = Math.round((a - b) * 1000) / 1000;
  const ok = Math.abs(a - b) <= (tol || 0);
  rows.push([check, label, a, b, delta, ok ? 'OK' : 'FLAG']);
  return ok ? 0 : 1;
}

function getSalesSeriesByListing_() {
  const map = new Map();
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(AD.SHEETS.SALES);
  if (!sh || sh.getLastRow() < 2) return map;
  const values = sh.getRange(2, 1, sh.getLastRow() - 1, AD.SALES_HEADERS.length).getValues();
  values.forEach(r => {
    const listing = clean_(r[3]);
    if (!listing || !isValidDate_(r[0])) return;
    if (!map.has(listing)) map.set(listing, []);
    map.get(listing).push({
      date: new Date(r[0]),
      units: number_(r[8]),
      du: number_(r[9]),
      ku: number_(r[10]),
      dk: number_(r[11]),
      roy: number_(r[12]),
      dr: number_(r[13])
    });
  });
  map.forEach(list => list.sort((a, b) => a.date - b.date));
  return map;
}
