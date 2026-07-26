/**
 * Stable internal IDs (Phase 1a)
 * Book ID: SOL-001, SOL-002, … (permanent; not title-based)
 * Listing ID: SOL-001-AMAZ-KIND-XXXX (book + store + format + suffix)
 *
 * Title / ASIN / marketplace / format stay as separate fields.
 */

function migrateStableBookIdsMenu() {
  const ui = SpreadsheetApp.getUi();
  const answer = ui.alert(
    'Migrate to stable SOL book IDs?',
    'Assigns permanent Book IDs (SOL-001, SOL-002, …) and updates Listing IDs plus Sales / Rank / Marketing history. Titles are not used as keys. Continue?',
    ui.ButtonSet.YES_NO
  );
  if (answer !== ui.Button.YES) return;
  const result = migrateStableBookIds_();
  refreshEverything();
  ui.alert(
    'Stable ID migration complete.\n\n' +
      'Books mapped: ' + result.booksMapped + '\n' +
      'Manual Entry rows updated: ' + result.inputRowsUpdated + '\n' +
      'History rows updated: ' + result.historyRowsUpdated + '\n\n' +
      formatBookIdMapForAlert_(result.orderedBooks)
  );
}

function migrateStableBookIds_() {
  ensureSalesHistorySchema_();
  const plan = buildStableBookIdPlan_();
  const applied = applyStableIdsToManualEntry_(plan);
  const historyRowsUpdated =
    rewriteHistoryIds_(AD.SHEETS.SALES, 2, 3, plan.bookIdByOld, applied.listingMap) +
    rewriteHistoryIds_(AD.SHEETS.RANKS, 2, 3, plan.bookIdByOld, applied.listingMap) +
    rewriteHistoryIds_(AD.SHEETS.MARKETING, 1, 2, plan.bookIdByOld, applied.listingMap);

  return {
    booksMapped: plan.orderedBooks.length,
    inputRowsUpdated: applied.inputRowsUpdated,
    historyRowsUpdated: historyRowsUpdated,
    map: plan.bookIdByOld,
    orderedBooks: plan.orderedBooks
  };
}

function formatBookIdMapForAlert_(orderedBooks) {
  const lines = ['Book ID map:'];
  (orderedBooks || []).forEach(b => {
    lines.push(
      b.bookId + ' = ' + b.title +
        (b.oldBookId && b.oldBookId !== b.bookId ? ' (was ' + b.oldBookId + ')' : '')
    );
  });
  return lines.join('\n');
}

/**
 * Preferred order for known Solmare titles, then series #, then title.
 */
function buildStableBookIdPlan_() {
  const rows = getInputRows_();
  const byBookKey = new Map(); // oldBookId or titleKey -> meta

  rows.forEach(r => {
    const title = clean_(r[AD.COL.TITLE]);
    if (!title) return;
    const oldBookId = clean_(r[AD.COL.BOOK_ID]);
    const key = oldBookId || ('TITLE:' + normalizeKey_(title));
    if (!byBookKey.has(key)) {
      byBookKey.set(key, {
        oldBookId: oldBookId,
        title: title,
        series: clean_(r[3]),
        seriesNum: number_(r[4]),
        titleKey: normalizeKey_(title)
      });
    }
  });

  const books = [...byBookKey.values()];
  const preferred = AD.STABLE_BOOK_ORDER.map(normalizeKey_);

  books.sort((a, b) => {
    const ai = preferred.indexOf(a.titleKey);
    const bi = preferred.indexOf(b.titleKey);
    const ap = ai === -1 ? 999 : ai;
    const bp = bi === -1 ? 999 : bi;
    if (ap !== bp) return ap - bp;
    if (a.seriesNum !== b.seriesNum) return a.seriesNum - b.seriesNum;
    return a.title.localeCompare(b.title);
  });

  const bookIdByOld = {}; // oldBookId or TITLE:key -> SOL-xxx
  const titleToSol = {};
  const orderedBooks = [];
  let next = 1;

  // Preserve existing SOL-### numbers when already assigned.
  books.forEach(b => {
    const m = String(b.oldBookId || '').match(/^SOL-(\d+)$/i);
    if (m) next = Math.max(next, Number(m[1]) + 1);
  });

  books.forEach(b => {
    let sol = '';
    if (/^SOL-\d+$/i.test(b.oldBookId || '')) {
      sol = String(b.oldBookId).toUpperCase();
    } else {
      sol = formatSolBookId_(next++);
    }
    orderedBooks.push({ title: b.title, oldBookId: b.oldBookId, bookId: sol });
    if (b.oldBookId) bookIdByOld[b.oldBookId] = sol;
    bookIdByOld['TITLE:' + b.titleKey] = sol;
    titleToSol[b.titleKey] = sol;
  });

  return { orderedBooks: orderedBooks, bookIdByOld: bookIdByOld, titleToSol: titleToSol };
}

function formatSolBookId_(n) {
  const num = Math.max(1, Math.floor(Number(n) || 1));
  return AD.BOOK_ID_PREFIX + '-' + ('000' + num).slice(-3);
}

function applyStableIdsToManualEntry_(plan) {
  const sh = getRequiredSheet_(AD.SHEETS.INPUT);
  const listingMap = {}; // oldListingId -> newListingId
  let inputRowsUpdated = 0;
  if (sh.getLastRow() < 2) {
    return { listingMap: listingMap, inputRowsUpdated: 0 };
  }

  const values = sh.getRange(2, 1, sh.getLastRow() - 1, AD.INPUT_HEADERS.length).getValues();
  values.forEach((r, i) => {
    const title = clean_(r[AD.COL.TITLE]);
    if (!title) return;
    const oldBookId = clean_(r[AD.COL.BOOK_ID]);
    const oldListingId = clean_(r[AD.COL.LISTING_ID]);
    const store = clean_(r[AD.COL.STORE]);
    const format = clean_(r[AD.COL.FORMAT]);
    const newBookId =
      (oldBookId && plan.bookIdByOld[oldBookId]) ||
      plan.titleToSol[normalizeKey_(title)] ||
      '';
    if (!newBookId) return;

    let newListingId = oldListingId;
    if (oldListingId && oldBookId && oldListingId.indexOf(oldBookId) === 0) {
      newListingId = newBookId + oldListingId.substring(oldBookId.length);
    } else if (!oldListingId && store && format) {
      newListingId = generateListingId_(newBookId, store, format);
    } else if (oldListingId && oldBookId && oldBookId !== newBookId) {
      // Listing did not start with book id — rebuild keeping a stable suffix when possible.
      const parts = oldListingId.split('-');
      const suffix = parts.length ? parts[parts.length - 1] : '';
      newListingId = generateListingId_(newBookId, store || 'Amazon', format || 'Other', suffix);
    }

    const row = i + 2;
    let changed = false;
    if (oldBookId !== newBookId) {
      sh.getRange(row, AD.COL.BOOK_ID + 1).setValue(newBookId);
      changed = true;
    }
    if (newListingId && oldListingId !== newListingId) {
      sh.getRange(row, AD.COL.LISTING_ID + 1).setValue(newListingId);
      if (oldListingId) listingMap[oldListingId] = newListingId;
      changed = true;
    }
    if (changed) inputRowsUpdated++;
  });

  return { listingMap: listingMap, inputRowsUpdated: inputRowsUpdated };
}

/**
 * Rewrite Book ID + Listing ID columns on a history sheet.
 * bookCol / listingCol are 0-based indexes in the row array.
 */
function rewriteHistoryIds_(sheetName, bookCol, listingCol, bookIdByOld, listingMap) {
  listingMap = listingMap || {};
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sh || sh.getLastRow() < 2) return 0;
  const lastCol = sh.getLastColumn();
  const values = sh.getRange(2, 1, sh.getLastRow() - 1, lastCol).getValues();
  let updated = 0;
  values.forEach((r, i) => {
    let changed = false;
    const oldBook = clean_(r[bookCol]);
    const oldListing = clean_(r[listingCol]);
    if (oldBook && bookIdByOld[oldBook] && bookIdByOld[oldBook] !== oldBook) {
      r[bookCol] = bookIdByOld[oldBook];
      changed = true;
    }
    if (oldListing && listingMap[oldListing] && listingMap[oldListing] !== oldListing) {
      r[listingCol] = listingMap[oldListing];
      changed = true;
    } else if (oldListing && oldBook && bookIdByOld[oldBook]) {
      const neuBook = bookIdByOld[oldBook];
      if (oldListing.indexOf(oldBook) === 0 && neuBook !== oldBook) {
        r[listingCol] = neuBook + oldListing.substring(oldBook.length);
        changed = true;
      }
    }
    if (changed) {
      values[i] = r;
      updated++;
    }
  });
  if (updated) sh.getRange(2, 1, values.length, lastCol).setValues(values);
  return updated;
}

function isStableBookId_(id) {
  return /^SOL-\d{3,}$/i.test(clean_(id));
}

function getMaxSolBookNumber_() {
  let max = 0;
  getInputRows_().forEach(r => {
    const m = String(clean_(r[AD.COL.BOOK_ID]) || '').match(/^SOL-(\d+)$/i);
    if (m) max = Math.max(max, Number(m[1]));
  });
  return max;
}
