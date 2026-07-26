/** Rename Lifetime Royalties → Lifetime Royalties (USD) without shifting columns. */
function ensureInputRoyaltyHeader_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(AD.SHEETS.INPUT);
  if (!sh) return;
  const cur = clean_(sh.getRange(1, AD.COL.ROYALTIES + 1).getValue());
  const wanted = AD.INPUT_HEADERS[AD.COL.ROYALTIES];
  if (cur !== wanted && (/^lifetime royalties$/i.test(cur) || !cur)) {
    sh.getRange(1, AD.COL.ROYALTIES + 1).setValue(wanted);
  }
}

/**
 * Insert eBook/Print/KENP royalty + KENPC columns before Process Status when missing.
 * Does not shift columns mid-schema beyond that insert.
 */
function ensureInputKuRoyaltySchema_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(AD.SHEETS.INPUT);
  if (!sh) return;
  ensureInputRoyaltyHeader_();
  const lastCol = Math.max(sh.getLastColumn(), 1);
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(clean_);
  const hasKenpc = headers.some(h => normalizeKey_(h) === 'kenpc');
  if (!hasKenpc) {
    let psIdx = headers.findIndex(h => normalizeKey_(h) === 'processstatus');
    if (psIdx < 0) psIdx = headers.length;
    sh.insertColumnsBefore(psIdx + 1, 4);
    sh.getRange(1, psIdx + 1, 1, 4).setValues([[
      'Lifetime eBook Royalties (USD)',
      'Lifetime Print Royalties (USD)',
      'Lifetime KENP Royalties (USD)',
      'KENPC'
    ]]);
  }
  sh.getRange(1, 1, 1, AD.INPUT_HEADERS.length).setValues([AD.INPUT_HEADERS]);
  styleHeader_(sh.getRange(1, 1, 1, AD.INPUT_HEADERS.length));
  applyInputFormats_(sh);
  applyManualEntryColumnVisibility_(sh);
}

function assignInternalIds_() {
  const sh = getRequiredSheet_(AD.SHEETS.INPUT);
  if (sh.getLastRow() < 2) return;
  const values = sh.getRange(2, 1, sh.getLastRow() - 1, AD.INPUT_HEADERS.length).getValues();
  const titleMap = new Map();
  let maxSol = getMaxSolBookNumber_();

  values.forEach(r => {
    if (clean_(r[0]) && clean_(r[2])) titleMap.set(normalizeKey_(r[2]), clean_(r[0]));
  });

  values.forEach((r, i) => {
    const row = i + 2;
    const title = clean_(r[2]);
    const store = clean_(r[8]);
    const format = clean_(r[9]);
    if (!title) return;
    let bookId = clean_(r[0]);
    if (!bookId) {
      bookId = titleMap.get(normalizeKey_(title));
      if (!bookId) {
        maxSol += 1;
        bookId = formatSolBookId_(maxSol);
      }
      sh.getRange(row, 1).setValue(bookId);
      titleMap.set(normalizeKey_(title), bookId);
    } else if (
      !isStableBookId_(bookId) &&
      titleMap.get(normalizeKey_(title)) &&
      isStableBookId_(titleMap.get(normalizeKey_(title)))
    ) {
      bookId = titleMap.get(normalizeKey_(title));
      sh.getRange(row, 1).setValue(bookId);
    } else {
      titleMap.set(normalizeKey_(title), bookId);
    }
    if (!clean_(r[1]) && store && format) {
      sh.getRange(row, 2).setValue(generateListingId_(bookId, store, format));
    }
  });
  syncListingReleaseDates_();
}

/** If Listing Release Date is blank, copy Original Release Date (same book/format launch in most cases). */
function syncListingReleaseDates_() {
  const sh = getRequiredSheet_(AD.SHEETS.INPUT);
  if (sh.getLastRow() < 2) return;
  const values = sh.getRange(2, 1, sh.getLastRow() - 1, AD.INPUT_HEADERS.length).getValues();
  values.forEach((r, i) => {
    if (!clean_(r[AD.COL.TITLE])) return;
    if (isValidDate_(r[AD.COL.LISTING_RELEASE])) return;
    if (!isValidDate_(r[AD.COL.ORIGINAL_RELEASE])) return;
    sh.getRange(i + 2, AD.COL.LISTING_RELEASE + 1).setValue(startOfDay_(new Date(r[AD.COL.ORIGINAL_RELEASE])));
  });
}

function createStoreUrls_() {
  const sh = getRequiredSheet_(AD.SHEETS.INPUT);
  if (sh.getLastRow() < 2) return;
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, AD.INPUT_HEADERS.length).getValues();
  rows.forEach((r, i) => {
    const store = clean_(r[8]);
    const type = clean_(r[11]).toUpperCase();
    const id = normalizeAsin_(r[12]);
    if (store.toLowerCase() === 'amazon' && type === 'ASIN' && id) {
      sh.getRange(i + 2, 16).setFormula('=HYPERLINK("https://www.amazon.com/dp/' + id + '","Open Amazon Listing")');
    }
  });
}

function addStandardAmazonFormats() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt('Add Amazon formats', 'Enter the exact book title:', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const title = res.getResponseText().trim();
  if (!title) return;
  const sh = getRequiredSheet_(AD.SHEETS.INPUT);
  const rows = getInputRows_();
  const base = rows.find(r => normalizeKey_(r[2]) === normalizeKey_(title));
  const existing = new Set(
    rows
      .filter(r => normalizeKey_(r[2]) === normalizeKey_(title) && normalizeKey_(r[8]) === 'amazon')
      .map(r => normalizeKey_(r[9]))
  );
  const newRows = [];
  ['Kindle eBook', 'Paperback', 'Hardcover'].forEach(format => {
    if (existing.has(normalizeKey_(format))) return;
    const r = new Array(AD.INPUT_HEADERS.length).fill('');
    r[2] = title;
    r[3] = base ? base[3] : '';
    r[4] = base ? base[4] : '';
    r[5] = base ? base[5] : 'Published';
    r[6] = base ? base[6] : '';
    r[7] = base ? base[7] : '';
    r[8] = 'Amazon';
    r[9] = format;
    r[10] = 'First Edition';
    r[11] = 'ASIN';
    r[13] = 'Live';
    r[14] = base && isValidDate_(base[14]) ? base[14] : (base && isValidDate_(base[7]) ? base[7] : '');
    newRows.push(r);
  });
  if (newRows.length) sh.getRange(sh.getLastRow() + 1, 1, newRows.length, AD.INPUT_HEADERS.length).setValues(newRows);
  refreshEverything();
}
