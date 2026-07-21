function recordCurrentSnapshot() {
  assignInternalIds_();
  createStoreUrls_();
  ensureRankHistorySchema_();
  const rows = getInputRows_();
  const today = getSpreadsheetToday_();
  const week = getWeekEndingDate_(today);
  recordRankSnapshot_(rows, today, week);
  recordSalesSnapshot_(rows, today, week);
  refreshSalesReports_();
  rebuildCatalogSummary_();
  refreshDashboard_();
  lockAutomaticSheets();
  SpreadsheetApp.getUi().alert('Current sales and rank snapshot recorded.');
}

function recordRankSnapshot_(rows, date, week) {
  ensureRankHistorySchema_();
  const sh = getRequiredSheet_(AD.SHEETS.RANKS);
  const keys = getRankHistoryDuplicateKeys_();
  const out = [];
  rows.forEach(r => {
    const listing = clean_(r[AD.COL.LISTING_ID]);
    const rank = number_(r[AD.COL.RANK]);
    if (!listing || !rank) return;
    const category = 'Manual Snapshot';
    const key = rankHistoryKey_(date, listing, 'Overall', category);
    if (keys.has(key)) return;
    out.push([
      date,
      week,
      clean_(r[AD.COL.BOOK_ID]),
      listing,
      clean_(r[AD.COL.TITLE]),
      clean_(r[AD.COL.STORE]),
      clean_(r[AD.COL.FORMAT]),
      clean_(r[AD.COL.IDENTIFIER]),
      'Overall',
      category,
      rank,
      '',
      'MANUAL'
    ]);
    keys.add(key);
  });
  appendRows_(sh, out);
  formatRankHistorySheet_(sh);
}

function recordSalesSnapshot_(rows, date, week) {
  const sh = getRequiredSheet_(AD.SHEETS.SALES);
  const prev = getLatestSalesByListing_();
  const keys = getExistingSnapshotKeys_(sh, 1, 4, null);
  const out = [];
  rows.forEach(r => {
    const listing = clean_(r[AD.COL.LISTING_ID]);
    if (!listing) return;
    const key = [dateKey_(date), listing].join('|');
    if (keys.has(key)) return;
    const u = number_(r[AD.COL.UNITS]);
    const k = number_(r[AD.COL.KU]);
    const roy = number_(r[AD.COL.ROYALTIES]);
    const p = prev.get(listing);
    out.push([
      date,
      week,
      clean_(r[AD.COL.BOOK_ID]),
      listing,
      clean_(r[AD.COL.TITLE]),
      clean_(r[AD.COL.STORE]),
      clean_(r[AD.COL.FORMAT]),
      clean_(r[AD.COL.IDENTIFIER]),
      u,
      p ? Math.max(0, u - p.units) : u,
      k,
      p ? Math.max(0, k - p.ku) : k,
      roy,
      p ? Math.max(0, roy - p.roy) : roy
    ]);
    keys.add(key);
  });
  appendRows_(sh, out);
  sh.getRange('A:B').setNumberFormat('m/d/yyyy');
  sh.getRange('I:L').setNumberFormat('#,##0');
  sh.getRange('M:N').setNumberFormat('$#,##0.00');
  formatSalesHistoryShading_();
}

function processMarketingEntries() { processMarketingEntries_(true); }

function processMarketingEntries_(showAlert) {
  const input = getRequiredSheet_(AD.SHEETS.INPUT);
  const outSheet = getRequiredSheet_(AD.SHEETS.MARKETING);
  if (input.getLastRow() < 2) return;
  const rows = getInputRows_();
  const out = [];
  const clear = [];
  rows.forEach((r, i) => {
    if (!isValidDate_(r[23]) || !clean_(r[24]) || !clean_(r[25])) return;
    out.push([
      new Date(r[23]),
      clean_(r[0]),
      clean_(r[1]),
      clean_(r[2]),
      clean_(r[8]),
      clean_(r[9]),
      clean_(r[24]),
      clean_(r[25]),
      number_(r[26]),
      clean_(r[27]),
      clean_(r[28])
    ]);
    clear.push(i + 2);
  });
  appendRows_(outSheet, out);
  clear.forEach(row => {
    input.getRange(row, 24, 1, 6).clearContent();
    input.getRange(row, 30).setValue('Marketing entry recorded');
  });
  if (showAlert) {
    SpreadsheetApp.getUi().alert(
      out.length
        ? out.length + ' marketing entr' + (out.length === 1 ? 'y' : 'ies') + ' recorded.'
        : 'No complete marketing entries found.'
    );
  }
}

function getRankHistoryMap_() {
  return getBestOverallRanksByBook_();
}

function getLatestSalesByListing_() {
  const sh = getRequiredSheet_(AD.SHEETS.SALES);
  const map = new Map();
  if (sh.getLastRow() < 2) return map;
  sh.getRange(2, 1, sh.getLastRow() - 1, AD.SALES_HEADERS.length).getValues().forEach(r => {
    const d = r[0];
    const id = clean_(r[3]);
    if (!id || !isValidDate_(d)) return;
    const cur = map.get(id);
    if (!cur || new Date(d) > cur.date) {
      map.set(id, { date: new Date(d), units: number_(r[8]), ku: number_(r[10]), roy: number_(r[12]) });
    }
  });
  return map;
}
