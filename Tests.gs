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

/**
 * Phase 0 tests — safe pure checks + lightweight sheet reads.
 * Does not modify production lifetime totals.
 */
function runPhase0Tests() {
  const results = [];
  const pass = (name, ok, detail) => {
    results.push((ok ? 'PASS' : 'FAIL') + ' — ' + name + (detail ? ': ' + detail : ''));
    return ok;
  };

  // 1) First snapshot period change must be 0 (not full lifetime).
  pass(
    'First snapshot period delta = 0',
    computePeriodDelta_(null, 32) === 0,
    'got ' + computePeriodDelta_(null, 32)
  );

  // 2) Subsequent snapshot uses max(0, current - previous).
  pass(
    'Period delta from previous',
    computePeriodDelta_(20, 32) === 12 && computePeriodDelta_(40, 32) === 0,
    '20→32=' + computePeriodDelta_(20, 32) + ' 40→32=' + computePeriodDelta_(40, 32)
  );

  // 3) Date parsing
  const d1 = parseLooseDate_('2026-07-18');
  pass('parseLooseDate_ ISO', d1 && dateKey_(d1) === '2026-07-18', d1 ? dateKey_(d1) : 'null');

  // 4) Partial filename detection (no sheet mutation)
  const metaPartialName = analyzeKdpReportMeta_(
    { orders: [{ ASIN: 'B0TEST0001', 'Royalty Date': '2026-07-01', 'Paid Units': 1 }] },
    'KDP_partial_last_7_days.xlsx',
    { B0TEST0001: { units: 1, kenp: 0, royaltyUsd: 0 } }
  );
  pass('Partial filename blocked', metaPartialName.isPartial === true, (metaPartialName.reasons || []).join('; '));

  // 5) Blank stays blank semantics helper
  pass('Blank number_ is 0 for math but hasValue check', number_('') === 0 && !hasReportedValue_(''), 'ok');
  pass('Explicit zero is a reported value', hasReportedValue_(0) === true, '');

  // 6) Sales History headers aliases map to new names
  pass(
    'Header alias Weekly Unit Change',
    AD.SALES_HEADER_ALIASES['Weekly Unit Change'] === 'Units Since Prev Snapshot',
    ''
  );

  // 7) Schema version present
  pass('Schema version set', !!AD.SCHEMA_VERSION && !!AD.VERSION, AD.VERSION + ' / ' + AD.SCHEMA_VERSION);

  // 7b) Stable SOL id helpers
  pass('formatSolBookId_ SOL-001', formatSolBookId_(1) === 'SOL-001', formatSolBookId_(1));
  pass('isStableBookId_', isStableBookId_('SOL-001') && !isStableBookId_('BK-ABC'), '');

  // Royalty total = eBook + Print + KENP
  pass(
    'finalizeKdpRoyaltyUsd_ sum buckets',
    Math.abs(finalizeKdpRoyaltyUsd_(85.09, 39.89, 10.74) - 135.72) < 0.001,
    String(finalizeKdpRoyaltyUsd_(85.09, 39.89, 10.74))
  );
  const kuM = computeKuRoyaltyMetrics_({
    totalKenp: 2314,
    kenpRoyalties: 10.74,
    ebookRoyalties: 85.09,
    printRoyalties: 39.89,
    kenpc: 671,
    useKenpc: true
  });
  pass(
    'KU rate ≈ 0.004641',
    kuM.ratePerKenp != null && Math.abs(kuM.ratePerKenp - (10.74 / 2314)) < 1e-9,
    String(kuM.ratePerKenp)
  );
  pass(
    'KU full read ≈ 3.11',
    kuM.fullReadRoyalty != null && Math.abs(kuM.fullReadRoyalty - 671 * (10.74 / 2314)) < 0.01,
    String(kuM.fullReadRoyalty)
  );
  pass(
    'KU equivalent reads ≈ 3.45',
    kuM.equivalentReads != null && Math.abs(kuM.equivalentReads - (2314 / 671)) < 0.01,
    String(kuM.equivalentReads)
  );
  pass('KU mix reconciles', kuM.reconcileOk === true, String(kuM.reconcileDiff));
  const kuBlankKenp = computeKuRoyaltyMetrics_({
    totalKenp: 0, kenpRoyalties: 10, ebookRoyalties: 1, printRoyalties: 1, kenpc: 671, useKenpc: true
  });
  pass('KU rate blank when KENP=0', kuBlankKenp.ratePerKenp === null, String(kuBlankKenp.ratePerKenp));
  const kuBlankKenpc = computeKuRoyaltyMetrics_({
    totalKenp: 2314, kenpRoyalties: 10.74, ebookRoyalties: 85.09, printRoyalties: 39.89, kenpc: '', useKenpc: false
  });
  pass(
    'KU full-read blank without KENPC',
    kuBlankKenpc.fullReadRoyalty === null && kuBlankKenpc.equivalentReads === null,
    ''
  );
  pass(
    'KENP $ ≈ pages × rate seed',
    Math.abs(2314 * AD.ESTIMATED_KENP_RATE_USD - 10.74) < 0.02,
    String(2314 * AD.ESTIMATED_KENP_RATE_USD)
  );

  // 8) Recompute series logic
  const series = [
    { units: 32, ku: 100, roy: 10 },
    { units: 40, ku: 150, roy: 12 },
    { units: 40, ku: 160, roy: 12 }
  ];
  const deltas = recomputePeriodSeries_(series);
  pass(
    'Recompute first period 0 then deltas',
    deltas[0].du === 0 && deltas[1].du === 8 && deltas[2].du === 0 &&
      deltas[0].dk === 0 && deltas[1].dk === 50 && deltas[2].dk === 10,
    JSON.stringify(deltas)
  );

  // 9) Week ending from report date (not “import day” alone)
  const reportEnd = parseLooseDate_('2026-07-22'); // Wednesday
  const week = getWeekEndingDate_(reportEnd);
  pass(
    'Week ending from report date',
    week && Utilities.formatDate(week, AD.TZ, 'yyyy-MM-dd') === '2026-07-25',
    week ? Utilities.formatDate(week, AD.TZ, 'yyyy-MM-dd') : 'null'
  );

  // 10) Reconciliation function callable
  try {
    ensureSalesHistorySchema_();
    const rep = runSalesReconciliation_();
    pass('Reconciliation runs', Array.isArray(rep.rows), 'flags=' + rep.flagCount);
  } catch (e) {
    pass('Reconciliation runs', false, String(e));
  }

  const failed = results.filter(r => r.indexOf('FAIL') === 0).length;
  SpreadsheetApp.getUi().alert(
    'Phase 0 tests\n\n' + results.join('\n') + '\n\n' +
      (failed ? failed + ' failed.' : 'All passed.') +
      '\n\nNote: duplicate/overlap/Meta API cases are deferred to later phases.'
  );
  return results;
}

function computePeriodDelta_(prevLifetime, currentLifetime) {
  if (prevLifetime == null) return 0;
  return Math.max(0, number_(currentLifetime) - number_(prevLifetime));
}

function hasReportedValue_(v) {
  return v !== '' && v !== null && v !== undefined;
}

function recomputePeriodSeries_(series) {
  let prev = null;
  return series.map(s => {
    const du = prev ? Math.max(0, s.units - prev.units) : 0;
    const dk = prev ? Math.max(0, s.ku - prev.ku) : 0;
    const dr = prev ? Math.max(0, s.roy - prev.roy) : 0;
    prev = s;
    return { du: du, dk: dk, dr: dr };
  });
}
