function updateAmazonRanks() {
  const summary = updateAmazonRanks_(true);
  SpreadsheetApp.getUi().alert(formatRankUpdateSummary_(summary));
}

function updateAmazonRanks_(showUi, asOfDate) {
  const lock = LockService.getDocumentLock();
  const gotLock = lock.tryLock(30000);
  if (!gotLock) {
    const busy = {
      listingsChecked: 0,
      successfulUpdates: 0,
      robotChecks: 0,
      missingRanks: 0,
      httpFailures: 0,
      otherFailures: 0,
      historyRowsAdded: 0,
      message: 'Could not acquire document lock. Try again in a moment.'
    };
    console.log(JSON.stringify(busy));
    return busy;
  }

  const summary = {
    listingsChecked: 0,
    successfulUpdates: 0,
    robotChecks: 0,
    missingRanks: 0,
    httpFailures: 0,
    otherFailures: 0,
    historyRowsAdded: 0
  };

  try {
    ensureRankHistorySchema_();
    assignInternalIds_();
    createStoreUrls_();

    const sh = getRequiredSheet_(AD.SHEETS.INPUT);
    if (sh.getLastRow() < 2) return summary;

    const range = sh.getRange(2, 1, sh.getLastRow() - 1, AD.INPUT_HEADERS.length);
    const rows = range.getValues();
    const today = asOfDate ? startOfDay_(asOfDate) : getSpreadsheetToday_();
    const week = getWeekEndingDate_(today);
    const historyKeys = getRankHistoryDuplicateKeys_();
    const historyOut = [];

    rows.forEach((r, i) => {
      const rowNum = i + 2;
      const store = normalizeKey_(r[AD.COL.STORE]);
      const idType = clean_(r[AD.COL.ID_TYPE]).toUpperCase();
      const status = normalizeKey_(r[AD.COL.STATUS]);
      const asin = normalizeAsin_(r[AD.COL.IDENTIFIER]);
      const listingId = clean_(r[AD.COL.LISTING_ID]);

      if (store !== 'amazon' || idType !== 'ASIN') return;
      if (AD.ACTIVE_LISTING_STATUSES.indexOf(status) === -1) return;

      summary.listingsChecked++;

      if (!asin) {
        summary.otherFailures++;
        sh.getRange(rowNum, AD.COL.PROCESS_STATUS + 1).setValue(AD.STATUS.MISSING_ASIN);
        return;
      }

      if (summary.listingsChecked > 1) Utilities.sleep(AD.AMAZON_FETCH_DELAY_MS);

      const result = fetchAmazonListingData_(asin);
      console.log('Amazon fetch ' + asin + ': ' + JSON.stringify({
        success: result.success,
        status: result.status,
        overallRank: result.overallRank,
        categories: (result.categoryRanks || []).length
      }));

      if (!result.success) {
        tallyAmazonFailure_(summary, result.status);
        sh.getRange(rowNum, AD.COL.PROCESS_STATUS + 1).setValue(statusMessageForAmazon_(result));
        return;
      }

      if (result.overallRank && result.overallRank > 0) {
        sh.getRange(rowNum, AD.COL.RANK + 1).setValue(result.overallRank);
      }
      if (result.rating != null && number_(result.rating) > 0) {
        sh.getRange(rowNum, AD.COL.RATING + 1).setValue(number_(result.rating));
      }
      if (result.reviewCount != null && number_(result.reviewCount) > 0) {
        sh.getRange(rowNum, AD.COL.REVIEWS + 1).setValue(Math.round(number_(result.reviewCount)));
      }
      applyReleaseDatesFromAmazon_(sh, rowNum, r, result.publicationDate);
      sh.getRange(rowNum, AD.COL.LAST_DATA_DATE + 1).setValue(today);
      sh.getRange(rowNum, AD.COL.PROCESS_STATUS + 1).setValue(AD.STATUS.UPDATED);
      summary.successfulUpdates++;

      const base = [
        today,
        week,
        clean_(r[AD.COL.BOOK_ID]),
        listingId,
        clean_(r[AD.COL.TITLE]),
        clean_(r[AD.COL.STORE]) || 'Amazon',
        clean_(r[AD.COL.FORMAT]),
        asin
      ];

      if (result.overallRank && result.overallRank > 0) {
        const category = result.overallCategory || 'Kindle Store';
        const key = rankHistoryKey_(today, listingId, 'Overall', category);
        if (!historyKeys.has(key)) {
          historyOut.push(base.concat(['Overall', category, result.overallRank, result.url, 'OK']));
          historyKeys.add(key);
        }
      }

      (result.categoryRanks || []).forEach(cr => {
        if (!cr || !cr.rank || cr.rank <= 0 || !cr.category) return;
        const key = rankHistoryKey_(today, listingId, 'Category', cr.category);
        if (historyKeys.has(key)) return;
        historyOut.push(base.concat(['Category', cr.category, cr.rank, result.url, 'OK']));
        historyKeys.add(key);
      });
    });

    if (historyOut.length) {
      const rankSheet = getRequiredSheet_(AD.SHEETS.RANKS);
      appendRows_(rankSheet, historyOut);
      formatRankHistorySheet_(rankSheet);
      summary.historyRowsAdded = historyOut.length;
    }

    rebuildCatalogSummary_();
    refreshDashboard_();
    lockAutomaticSheets();
  } catch (err) {
    console.error('updateAmazonRanks_ failed: ' + err);
    summary.otherFailures++;
    summary.message = String(err && err.message ? err.message : err);
  } finally {
    lock.releaseLock();
  }

  console.log('Amazon rank update summary: ' + JSON.stringify(summary));
  return summary;
}

/** Fill blank Original / Listing release dates from Amazon publication date, or copy original → listing. */
function applyReleaseDatesFromAmazon_(sh, rowNum, rowValues, publicationDate) {
  const pub = publicationDate && isValidDate_(publicationDate) ? startOfDay_(new Date(publicationDate)) : null;
  const hasOriginal = isValidDate_(rowValues[AD.COL.ORIGINAL_RELEASE]);
  const hasListing = isValidDate_(rowValues[AD.COL.LISTING_RELEASE]);

  if (!hasOriginal && pub) {
    sh.getRange(rowNum, AD.COL.ORIGINAL_RELEASE + 1).setValue(pub);
    rowValues[AD.COL.ORIGINAL_RELEASE] = pub;
  }
  if (!hasListing) {
    const source = pub || (isValidDate_(rowValues[AD.COL.ORIGINAL_RELEASE])
      ? startOfDay_(new Date(rowValues[AD.COL.ORIGINAL_RELEASE]))
      : null);
    if (source) sh.getRange(rowNum, AD.COL.LISTING_RELEASE + 1).setValue(source);
  }
}

function tallyAmazonFailure_(summary, status) {
  const s = String(status || '');
  if (s === 'ROBOT_CHECK') summary.robotChecks++;
  else if (s === 'MISSING_RANK') summary.missingRanks++;
  else if (/^HTTP_/.test(s)) summary.httpFailures++;
  else summary.otherFailures++;
}

function statusMessageForAmazon_(result) {
  const status = String(result && result.status || '');
  if (status === 'ROBOT_CHECK') return AD.STATUS.ROBOT;
  if (status === 'MISSING_RANK') return AD.STATUS.MISSING_RANK;
  if (status === 'INVALID_ASIN') return AD.STATUS.MISSING_ASIN;
  if (status === 'UNAVAILABLE') return AD.STATUS.UNAVAILABLE;
  if (status === 'PARSER_ERROR') return AD.STATUS.PARSER;
  if (/^HTTP_/.test(status)) return AD.STATUS.http(status.replace('HTTP_', ''));
  return result && result.message ? result.message : AD.STATUS.PARSER;
}

function formatRankUpdateSummary_(summary) {
  return [
    'Amazon rank update complete.',
    '',
    'Listings checked: ' + summary.listingsChecked,
    'Successful updates: ' + summary.successfulUpdates,
    'Robot checks: ' + summary.robotChecks,
    'Missing ranks: ' + summary.missingRanks,
    'HTTP failures: ' + summary.httpFailures,
    'Other failures: ' + summary.otherFailures,
    'History rows added: ' + summary.historyRowsAdded,
    summary.message ? ('\n' + summary.message) : ''
  ].join('\n');
}

function ensureRankHistorySchema_() {
  const sh = getRequiredSheet_(AD.SHEETS.RANKS);
  const lastCol = Math.max(sh.getLastColumn(), 1);
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(clean_);
  const expected = AD.RANK_HEADERS;

  const alreadyGood = expected.every((h, i) => normalizeKey_(headers[i] || '') === normalizeKey_(h));
  if (alreadyGood && headers.length >= expected.length) {
    formatRankHistorySheet_(sh);
    return false;
  }

  const oldRows = sh.getLastRow() >= 2
    ? sh.getRange(2, 1, sh.getLastRow() - 1, lastCol).getValues()
    : [];

  const mapped = oldRows.map(r => mapLegacyRankRow_(r, headers));
  sh.clear();
  sh.getRange(1, 1, 1, expected.length).setValues([expected]);
  styleHeader_(sh.getRange(1, 1, 1, expected.length));
  sh.setFrozenRows(1);
  if (mapped.length) sh.getRange(2, 1, mapped.length, expected.length).setValues(mapped);
  addFilter_(sh, expected.length);
  formatRankHistorySheet_(sh);
  return true;
}

function mapLegacyRankRow_(row, headers) {
  const idx = name => {
    const i = headers.findIndex(h => normalizeKey_(h) === normalizeKey_(name));
    return i;
  };
  const get = (name, fallbackIndex) => {
    const i = idx(name);
    if (i >= 0) return row[i];
    return fallbackIndex != null ? row[fallbackIndex] : '';
  };

  let rankType = clean_(get('Rank Type', 8));
  let category = clean_(get('Category', -1));
  let rank = number_(get('Rank', 9));

  if (/^overall(\s*rank)?$/i.test(rankType)) {
    rankType = 'Overall';
    if (!category) category = 'Kindle Store';
  } else if (/^category$/i.test(rankType)) {
    rankType = 'Category';
  } else if (rankType && !category) {
    category = rankType;
    rankType = 'Category';
  }

  return [
    get('Snapshot Date', 0),
    get('Week Ending', 1),
    get('Book ID', 2),
    get('Listing ID', 3),
    get('Book Title', 4),
    get('Store', 5),
    get('Format', 6),
    get('Identifier', 7),
    rankType,
    category,
    rank > 0 ? rank : get('Rank', 9),
    get('Source URL', -1) || '',
    get('Fetch Status', -1) || ''
  ];
}

function formatRankHistorySheet_(sh) {
  const cols = AD.RANK_HEADERS.length;
  if (sh.getLastRow() >= 2) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 2).setNumberFormat('m/d/yyyy');
    sh.getRange(2, 11, sh.getLastRow() - 1, 1).setNumberFormat('#,##0');
  }
  [110, 110, 115, 175, 220, 90, 120, 120, 90, 180, 90, 220, 110].forEach((w, i) => {
    if (i < cols) sh.setColumnWidth(i + 1, w);
  });
}

function getRankHistoryDuplicateKeys_() {
  const sh = getRequiredSheet_(AD.SHEETS.RANKS);
  const set = new Set();
  if (sh.getLastRow() < 2) return set;
  sh.getRange(2, 1, sh.getLastRow() - 1, AD.RANK_HEADERS.length).getValues().forEach(r => {
    if (!isValidDate_(r[0]) || !clean_(r[3])) return;
    set.add(rankHistoryKey_(new Date(r[0]), clean_(r[3]), clean_(r[8]), clean_(r[9])));
  });
  return set;
}

function rankHistoryKey_(date, listingId, rankType, category) {
  return [dateKey_(new Date(date)), clean_(listingId), clean_(rankType), clean_(category)].join('|');
}

function getSpreadsheetToday_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let tz = AD.TZ;
  try { tz = ss.getSpreadsheetTimeZone() || AD.TZ; } catch (e) {}
  const now = new Date();
  const iso = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
  const parts = iso.split('-').map(Number);
  return new Date(parts[0], parts[1] - 1, parts[2]);
}

function getBestOverallRanksByBook_() {
  const sh = getRequiredSheet_(AD.SHEETS.RANKS);
  const map = new Map();
  if (sh.getLastRow() < 2) return map;
  sh.getRange(2, 1, sh.getLastRow() - 1, AD.RANK_HEADERS.length).getValues().forEach(r => {
    const bookId = clean_(r[2]);
    const rankType = clean_(r[8]);
    const rank = number_(r[10]);
    if (!bookId || rank <= 0) return;
    if (!/^overall$/i.test(rankType) && !/^overall\s*rank$/i.test(rankType)) return;
    if (!map.has(bookId) || rank < map.get(bookId)) map.set(bookId, rank);
  });
  return map;
}

function getLastRankUpdateByBook_() {
  const sh = getRequiredSheet_(AD.SHEETS.RANKS);
  const map = new Map();
  if (sh.getLastRow() < 2) return map;
  sh.getRange(2, 1, sh.getLastRow() - 1, AD.RANK_HEADERS.length).getValues().forEach(r => {
    const bookId = clean_(r[2]);
    const rankType = clean_(r[8]);
    if (!bookId || !isValidDate_(r[0])) return;
    if (!/^overall$/i.test(rankType) && !/^overall\s*rank$/i.test(rankType) && !/^category$/i.test(rankType)) return;
    const d = new Date(r[0]);
    if (!map.has(bookId) || d > map.get(bookId)) map.set(bookId, d);
  });
  return map;
}

function getOverallRankSeries_() {
  const sh = getRequiredSheet_(AD.SHEETS.RANKS);
  if (sh.getLastRow() < 2) return [];
  const byDate = new Map();
  sh.getRange(2, 1, sh.getLastRow() - 1, AD.RANK_HEADERS.length).getValues().forEach(r => {
    if (!isValidDate_(r[0])) return;
    const rankType = clean_(r[8]);
    const rank = number_(r[10]);
    if (rank <= 0) return;
    if (!/^overall$/i.test(rankType) && !/^overall\s*rank$/i.test(rankType)) return;
    const key = dateKey_(new Date(r[0]));
    if (!byDate.has(key) || rank < byDate.get(key).rank) {
      byDate.set(key, { date: startOfDay_(new Date(r[0])), rank: rank, title: clean_(r[4]) });
    }
  });
  return [...byDate.values()].sort((a, b) => a.date - b.date);
}
