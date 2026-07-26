function recordCurrentSnapshot() {
  assignInternalIds_();
  createStoreUrls_();
  ensureRankHistorySchema_();
  ensureSalesHistorySchema_();
  const rows = getInputRows_();
  const today = getSpreadsheetToday_();
  const week = getWeekEndingDate_(today);
  // Sales only — rank history comes from Update Amazon Rankings / weekly rank trigger
  // (not a MANUAL copy of Manual Entry ranks).
  recordSalesSnapshot_(rows, today, week);
  refreshSalesReports_();
  rebuildCatalogSummary_();
  refreshDashboard_();
  rebuildReconciliationSheet_();
  lockAutomaticSheets();
  SpreadsheetApp.getUi().alert('Current sales snapshot recorded. (Ranks are written only by Amazon rank updates.)');
}

/**
 * Phase 0 migration (safe, incremental):
 * - Rename Sales History period column headers
 * - Recompute period deltas from lifetime series (fixes inflated first-snapshot "weekly" values)
 * - Rebuild year / YoY / Reconciliation sheets
 */
function migrateSalesHistoryPhase0() {
  const ui = SpreadsheetApp.getUi();
  const answer = ui.alert(
    'Migrate Sales History (Phase 0)?',
    'This renames period columns and recomputes Units/KENP/Royalties Since Prev Snapshot from Lifetime columns. Lifetime values are not changed. Continue?',
    ui.ButtonSet.YES_NO
  );
  if (answer !== ui.Button.YES) return;

  ensureSalesHistorySchema_();
  const removedDupWeeks = consolidateSalesHistoryByWeekEnding_();
  const recomputed = recomputeSalesPeriodChangesFromLifetime_();
  const removedRanks = cleanupManualRankSnapshots_();
  refreshSalesReports_();
  rebuildCatalogSummary_();
  refreshDashboard_();
  rebuildReconciliationSheet_();
  lockAutomaticSheets();
  ui.alert(
    'Phase 0 sales migration complete.\n\n' +
      'Duplicate week rows removed: ' + removedDupWeeks + '\n' +
      'Rows recomputed: ' + recomputed.rowsUpdated + '\n' +
      'Listings touched: ' + recomputed.listingsTouched + '\n' +
      'Manual Snapshot rank rows removed: ' + removedRanks + '\n\n' +
      'Reconciliation sheet is hidden (still exists for diagnostics).'
  );
}

/** Ensure Sales History headers match AD.SALES_HEADERS (alias rename only). */
function ensureSalesHistorySchema_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(AD.SHEETS.SALES);
  if (!sh || sh.getLastRow() < 1) return false;
  const cols = AD.SALES_HEADERS.length;
  const headers = sh.getRange(1, 1, 1, cols).getValues()[0].map(clean_);
  let changed = false;
  AD.SALES_HEADERS.forEach((wanted, i) => {
    const cur = headers[i] || '';
    const aliasTarget = AD.SALES_HEADER_ALIASES[cur];
    if (cur !== wanted && (aliasTarget === wanted || !cur)) {
      sh.getRange(1, i + 1).setValue(wanted);
      changed = true;
    }
  });
  if (changed) styleHeader_(sh.getRange(1, 1, 1, cols));
  // Notes clarifying the three concepts.
  sh.getRange(1, 9).setNote('Lifetime cumulative units at snapshot time (imported or entered).');
  sh.getRange(1, 10).setNote('Change since previous snapshot for this listing. NOT inferred from import day. First snapshot = 0.');
  sh.getRange(1, 11).setNote('Lifetime cumulative KENP pages at snapshot time.');
  sh.getRange(1, 12).setNote('KENP change since previous snapshot. First snapshot = 0.');
  sh.getRange(1, 13).setNote('Lifetime cumulative royalties (USD) at snapshot time.');
  sh.getRange(1, 14).setNote('Royalty change since previous snapshot. First snapshot = 0.');
  return changed;
}

/**
 * Recompute period columns from sorted lifetime series per listing.
 * First snapshot for a listing → period deltas = 0 (lifetime stays as recorded).
 */
function recomputeSalesPeriodChangesFromLifetime_() {
  const sh = getRequiredSheet_(AD.SHEETS.SALES);
  const result = { rowsUpdated: 0, listingsTouched: 0 };
  if (sh.getLastRow() < 2) return result;

  const lastRow = sh.getLastRow();
  const cols = AD.SALES_HEADERS.length;
  const values = sh.getRange(2, 1, lastRow - 1, cols).getValues();
  const byListing = new Map();

  values.forEach((r, idx) => {
    const listing = clean_(r[3]);
    if (!listing || !isValidDate_(r[0])) return;
    if (!byListing.has(listing)) byListing.set(listing, []);
    byListing.get(listing).push({ idx: idx, date: new Date(r[0]), row: r });
  });

  byListing.forEach(list => {
    list.sort((a, b) => a.date - b.date || a.idx - b.idx);
    let prev = null;
    list.forEach(item => {
      const u = number_(item.row[8]);
      const k = number_(item.row[10]);
      const roy = number_(item.row[12]);
      const du = prev ? Math.max(0, u - prev.units) : 0;
      const dk = prev ? Math.max(0, k - prev.ku) : 0;
      const dr = prev ? Math.max(0, roy - prev.roy) : 0;
      if (number_(item.row[9]) !== du || number_(item.row[11]) !== dk || number_(item.row[13]) !== dr) {
        values[item.idx][9] = du;
        values[item.idx][11] = dk;
        values[item.idx][13] = dr;
        result.rowsUpdated++;
      }
      prev = { units: u, ku: k, roy: roy };
    });
    result.listingsTouched++;
  });

  sh.getRange(2, 1, values.length, cols).setValues(values);
  return result;
}

/**
 * Legacy helper — no longer called from sales snapshot / weekly sales trigger.
 * Rank History rows should come from Amazon fetch (Fetch Status OK), not MANUAL copies.
 * Kept for rollback / rare manual scripting only.
 */
function recordRankSnapshot_(rows, date, week) {
  ensureRankHistorySchema_();
  const sh = getRequiredSheet_(AD.SHEETS.RANKS);
  const keys = getRankHistoryDuplicateKeys_();
  const out = [];
  rows.forEach(r => {
    const listing = clean_(r[AD.COL.LISTING_ID]);
    const rank = number_(r[AD.COL.RANK]);
    if (!listing || !rank) return;
    const category = 'Entry Snapshot';
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
      'SNAPSHOT'
    ]);
    keys.add(key);
  });
  appendRows_(sh, out);
  formatRankHistorySheet_(sh);
}

/**
 * Remove Manual Snapshot / MANUAL rank rows ONLY when a real Amazon Overall row
 * already exists for the same date + listing. Never delete the sole data point
 * for a format/day (that caused missing paperback ranks on 7/25).
 */
function cleanupManualRankSnapshots_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(AD.SHEETS.RANKS);
  if (!sh || sh.getLastRow() < 2) return 0;
  const cols = AD.RANK_HEADERS.length;
  const values = sh.getRange(2, 1, sh.getLastRow() - 1, cols).getValues();

  const amazonOverall = new Set();
  values.forEach(r => {
    if (!isValidDate_(r[0])) return;
    if (!/^overall$/i.test(clean_(r[8]))) return;
    const status = clean_(r[12]).toUpperCase();
    const category = normalizeKey_(r[9]);
    if (status === 'MANUAL' || category === 'manual snapshot') return;
    const listing = clean_(r[3]);
    if (!listing) return;
    amazonOverall.add(dateKey_(new Date(r[0])) + '|' + listing);
  });

  let removed = 0;
  for (let i = values.length - 1; i >= 0; i--) {
    if (!/^overall$/i.test(clean_(values[i][8]))) continue;
    const category = normalizeKey_(values[i][9]);
    const status = clean_(values[i][12]).toUpperCase();
    const isManual = category === 'manual snapshot' || status === 'MANUAL';
    if (!isManual) continue;
    const listing = clean_(values[i][3]);
    if (!listing || !isValidDate_(values[i][0])) continue;
    const key = dateKey_(new Date(values[i][0])) + '|' + listing;
    // Only remove when a real Amazon Overall already covers this listing/day.
    if (amazonOverall.has(key)) {
      sh.deleteRow(i + 2);
      removed++;
    }
  }
  return removed;
}

/**
 * Restore Overall ranks for 2026-07-25 when a format is missing but we still know
 * the values (from the former Manual Snapshot rows) or Manual Entry.
 * Looks up current Listing/Book IDs by ASIN so SOL migration stays intact.
 */
function repairMissingJuly25OverallRanks_() {
  const targets = [
    { asin: 'B0H97B5M1L', formatHint: 'kindle', rank: 45424 },
    { asin: 'B0H97M1ZMK', formatHint: 'paper', rank: 548832 },
    { asin: 'B0H8XM27K6', formatHint: 'hard', rank: 1109541 }
  ];
  const snap = startOfDay_(new Date(2026, 6, 25)); // month 0-based → July
  const week = getWeekEndingDate_(snap);
  const sh = getRequiredSheet_(AD.SHEETS.RANKS);
  ensureRankHistorySchema_();
  const keys = getRankHistoryDuplicateKeys_();
  const input = getInputRows_();
  const out = [];

  targets.forEach(t => {
    const row = input.find(r =>
      normalizeAsin_(r[AD.COL.IDENTIFIER]) === t.asin ||
      normalizeKey_(r[AD.COL.IDENTIFIER]) === normalizeKey_(t.asin)
    );
    if (!row) return;
    const listing = clean_(row[AD.COL.LISTING_ID]);
    if (!listing) return;

    // Already have any Overall for this listing on 7/25?
    const has = [...keys].some(k => k.indexOf(dateKey_(snap) + '|' + listing + '|Overall|') === 0);
    if (has) return;

    const category = 'Restored Snapshot';
    const key = rankHistoryKey_(snap, listing, 'Overall', category);
    if (keys.has(key)) return;

    // Use the known 7/25 values (not today's Manual Entry rank).
    out.push([
      snap,
      week,
      clean_(row[AD.COL.BOOK_ID]),
      listing,
      clean_(row[AD.COL.TITLE]),
      clean_(row[AD.COL.STORE]) || 'Amazon',
      clean_(row[AD.COL.FORMAT]),
      t.asin,
      'Overall',
      category,
      t.rank,
      '',
      'RESTORED'
    ]);
    keys.add(key);
  });

  if (out.length) {
    appendRows_(sh, out);
    formatRankHistorySheet_(sh);
  }
  return out.length;
}

/**
 * Snapshot Manual Entry lifetime totals into Sales History.
 * Period columns = change since previous snapshot for that listing.
 * First snapshot for a listing writes 0 for period columns (lifetime is still stored).
 *
 * options.upsertByWeek — for KDP imports: update existing Week Ending + Listing row
 * instead of appending a mid-week duplicate (fixes KENP Sunday → Week Ending next Sat).
 */
function recordSalesSnapshot_(rows, date, week, reportEndDate, options) {
  options = options || {};
  ensureSalesHistorySchema_();
  const sh = getRequiredSheet_(AD.SHEETS.SALES);
  const snapDate = startOfDay_(date);
  const weekEnd = reportEndDate && isValidDate_(reportEndDate)
    ? getWeekEndingDate_(reportEndDate)
    : week;
  const upsertByWeek = !!options.upsertByWeek;
  const dateKeys = getExistingSnapshotKeys_(sh, 1, 4, null);
  const weekRowMap = getSalesWeekListingRowMap_();
  const out = [];

  rows.forEach(r => {
    const listing = clean_(r[AD.COL.LISTING_ID]);
    if (!listing) return;
    const u = number_(r[AD.COL.UNITS]);
    const k = number_(r[AD.COL.KU]);
    const roy = number_(r[AD.COL.ROYALTIES]);
    const p = getLatestSalesByListingBeforeWeek_(listing, weekEnd);
    const du = p ? Math.max(0, u - p.units) : 0;
    const dk = p ? Math.max(0, k - p.ku) : 0;
    const dr = p ? Math.max(0, roy - p.roy) : 0;
    const rowVals = [
      snapDate,
      weekEnd,
      clean_(r[AD.COL.BOOK_ID]),
      listing,
      clean_(r[AD.COL.TITLE]),
      clean_(r[AD.COL.STORE]),
      clean_(r[AD.COL.FORMAT]),
      clean_(r[AD.COL.IDENTIFIER]),
      u,
      du,
      k,
      dk,
      roy,
      dr
    ];

    const weekKey = dateKey_(weekEnd) + '|' + listing;
    if (upsertByWeek && weekRowMap[weekKey]) {
      sh.getRange(weekRowMap[weekKey], 1, 1, AD.SALES_HEADERS.length).setValues([rowVals]);
      return;
    }

    const dateKey = dateKey_(snapDate) + '|' + listing;
    if (dateKeys.has(dateKey)) return;
    out.push(rowVals);
    dateKeys.add(dateKey);
    weekRowMap[weekKey] = sh.getLastRow() + 1 + out.length - 1;
  });

  appendRows_(sh, out);
  sh.getRange('A:B').setNumberFormat('m/d/yyyy');
  sh.getRange('I:L').setNumberFormat('#,##0');
  sh.getRange('M:N').setNumberFormat('$#,##0.00');
  formatSalesHistoryShading_();
}

/** Map weekEnding|listingId -> sheet row number (1-based). Keeps latest snapshot date per week. */
function getSalesWeekListingRowMap_() {
  const map = {};
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(AD.SHEETS.SALES);
  if (!sh || sh.getLastRow() < 2) return map;
  const values = sh.getRange(2, 1, sh.getLastRow() - 1, AD.SALES_HEADERS.length).getValues();
  const best = {};
  values.forEach((r, i) => {
    if (!isValidDate_(r[1]) || !clean_(r[3])) return;
    const weekKey = dateKey_(startOfDay_(new Date(r[1]))) + '|' + clean_(r[3]);
    const snap = isValidDate_(r[0]) ? new Date(r[0]).getTime() : 0;
    if (!best[weekKey] || snap >= best[weekKey].snap) {
      best[weekKey] = { snap: snap, row: i + 2 };
    }
  });
  Object.keys(best).forEach(k => { map[k] = best[k].row; });
  return map;
}

/** Latest Sales History lifetime for listing with Week Ending strictly before weekEnd. */
function getLatestSalesByListingBeforeWeek_(listing, weekEnd) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(AD.SHEETS.SALES);
  if (!sh || sh.getLastRow() < 2 || !listing) return null;
  const weekEndT = startOfDay_(weekEnd).getTime();
  let best = null;
  sh.getRange(2, 1, sh.getLastRow() - 1, AD.SALES_HEADERS.length).getValues().forEach(r => {
    if (clean_(r[3]) !== listing || !isValidDate_(r[1])) return;
    const wt = startOfDay_(new Date(r[1])).getTime();
    if (wt >= weekEndT) return;
    const snap = isValidDate_(r[0]) ? new Date(r[0]).getTime() : wt;
    if (!best || snap > best.snap) {
      best = {
        snap: snap,
        units: number_(r[8]),
        ku: number_(r[10]),
        roy: number_(r[12])
      };
    }
  });
  return best;
}

/**
 * Remove duplicate Week Ending + Listing rows (keep latest Snapshot Date), then
 * recompute period columns. Cleans mid-week KDP re-imports that created extra buckets.
 */
function consolidateSalesHistoryByWeekEnding_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(AD.SHEETS.SALES);
  if (!sh || sh.getLastRow() < 2) return 0;
  const values = sh.getRange(2, 1, sh.getLastRow() - 1, AD.SALES_HEADERS.length).getValues();
  const bestIdx = {};
  const bestSnap = {};
  values.forEach((r, i) => {
    if (!isValidDate_(r[1]) || !clean_(r[3])) return;
    const weekKey = dateKey_(startOfDay_(new Date(r[1]))) + '|' + clean_(r[3]);
    const snap = isValidDate_(r[0]) ? new Date(r[0]).getTime() : 0;
    if (bestIdx[weekKey] == null || snap >= bestSnap[weekKey]) {
      bestIdx[weekKey] = i;
      bestSnap[weekKey] = snap;
    }
  });
  const keep = new Set(Object.keys(bestIdx).map(k => bestIdx[k]));
  let removed = 0;
  for (let i = values.length - 1; i >= 0; i--) {
    if (keep.has(i)) continue;
    if (!isValidDate_(values[i][1]) || !clean_(values[i][3])) continue;
    sh.deleteRow(i + 2);
    removed++;
  }
  return removed;
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
