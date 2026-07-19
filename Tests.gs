function testAmazonRankFetch() {
  const rows = getInputRows_();
  const hit = rows.find(r =>
    normalizeKey_(r[AD.COL.STORE]) === 'amazon' &&
    clean_(r[AD.COL.ID_TYPE]).toUpperCase() === 'ASIN' &&
    normalizeAsin_(r[AD.COL.IDENTIFIER])
  );

  if (!hit) {
    SpreadsheetApp.getUi().alert('No Amazon ASIN found in Manual Entry.');
    return;
  }

  const asin = normalizeAsin_(hit[AD.COL.IDENTIFIER]);
  const result = fetchAmazonListingData_(asin);
  console.log('testAmazonRankFetch result: ' + JSON.stringify(result, null, 2));

  SpreadsheetApp.getUi().alert([
    'Amazon rank fetch test',
    '',
    'ASIN: ' + asin,
    'Overall rank: ' + (result.overallRank != null ? result.overallRank : '(none)'),
    'Category ranks: ' + ((result.categoryRanks || []).length),
    'Rating: ' + (result.rating != null ? result.rating : '(none)'),
    'Reviews: ' + (result.reviewCount != null ? result.reviewCount : '(none)'),
    'Status: ' + (result.status || '(unknown)')
  ].join('\n'));
}

function testRankHistoryAppend() {
  ensureRankHistorySchema_();
  const sh = getRequiredSheet_(AD.SHEETS.RANKS);
  const today = getSpreadsheetToday_();
  const week = getWeekEndingDate_(today);
  const listingId = '__TEST_LISTING__';
  const category = '__TEST_CATEGORY__';
  const key = rankHistoryKey_(today, listingId, 'Overall', category);

  const before = getRankHistoryDuplicateKeys_();
  const row = [
    today,
    week,
    'BK-TEST',
    listingId,
    'TEST ROW — SAFE TO DELETE',
    'Amazon',
    'Kindle eBook',
    'B000000000',
    'Overall',
    category,
    999999,
    'https://www.amazon.com/dp/B000000000',
    'TEST'
  ];

  if (!before.has(key)) {
    appendRows_(sh, [row]);
  }

  const mid = getRankHistoryDuplicateKeys_();
  if (!mid.has(key)) {
    SpreadsheetApp.getUi().alert('Duplicate-key test failed: key was not recorded.');
    return;
  }

  // Attempt a duplicate append path without writing again.
  const wouldDuplicate = mid.has(key);
  removeTestRankHistoryRows_();

  SpreadsheetApp.getUi().alert([
    'Rank history duplicate prevention test complete.',
    '',
    'Duplicate key blocked: ' + (wouldDuplicate ? 'YES' : 'NO'),
    'Test rows removed: YES'
  ].join('\n'));
}

function removeTestRankHistoryRows_() {
  const sh = getRequiredSheet_(AD.SHEETS.RANKS);
  if (sh.getLastRow() < 2) return;
  const values = sh.getRange(2, 1, sh.getLastRow() - 1, AD.RANK_HEADERS.length).getValues();
  for (let i = values.length - 1; i >= 0; i--) {
    if (clean_(values[i][3]) === '__TEST_LISTING__' || clean_(values[i][12]) === 'TEST') {
      sh.deleteRow(i + 2);
    }
  }
}
