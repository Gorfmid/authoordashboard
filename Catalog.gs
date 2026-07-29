function rebuildCatalogSummary_() {
  const catalog = getRequiredSheet_(AD.SHEETS.CATALOG);
  ensureCatalogSchema_();
  ensureRankHistorySchema_();
  clearDataRows_(catalog);
  const books = new Map();
  getInputRows_().forEach(r => {
    const id = clean_(r[AD.COL.BOOK_ID]);
    const title = clean_(r[AD.COL.TITLE]);
    if (!id || !title) return;
    if (!books.has(id)) {
      books.set(id, {
        id: id,
        title: title,
        series: clean_(r[3]),
        num: r[4],
        stage: clean_(r[5]),
        words: number_(r[6]),
        date: r[7],
        formats: new Set(),
        listings: 0,
        active: 0,
        ranks: [],
        units: 0,
        ku: 0,
        roy: 0,
        royEbook: 0,
        royPrint: 0,
        royKenp: 0,
        kenpc: '',
        reviews: 0,
        ratings: [],
        dates: []
      });
    }
    const b = books.get(id);
    if (clean_(r[AD.COL.FORMAT])) b.formats.add(clean_(r[AD.COL.FORMAT]));
    b.listings++;
    if (normalizeKey_(r[AD.COL.STATUS]) === 'live') b.active++;
    if (number_(r[AD.COL.RANK]) > 0) b.ranks.push(number_(r[AD.COL.RANK]));
    b.units += number_(r[AD.COL.UNITS]);
    b.ku += number_(r[AD.COL.KU]);
    b.roy += number_(r[AD.COL.ROYALTIES]);
    b.royEbook += number_(r[AD.COL.ROYALTY_EBOOK]);
    b.royPrint += number_(r[AD.COL.ROYALTY_PRINT]);
    b.royKenp += number_(r[AD.COL.ROYALTY_KENP]);
    const kenpcVal = r[AD.COL.KENPC];
    if ((b.kenpc === '' || b.kenpc == null) && kenpcVal !== '' && kenpcVal != null && number_(kenpcVal) > 0) {
      b.kenpc = number_(kenpcVal);
    }
    // Reviews/comments: take MAX across ASINs (formats often share the same Amazon pool).
    const rev = number_(r[AD.COL.REVIEWS]);
    if (rev > b.reviews) b.reviews = rev;
    if (number_(r[AD.COL.RATING]) > 0) b.ratings.push(number_(r[AD.COL.RATING]));
    if (isValidDate_(r[AD.COL.LAST_DATA_DATE])) b.dates.push(new Date(r[AD.COL.LAST_DATA_DATE]));
  });

  // If KENP $ was never written to Manual Entry, estimate from pages × rate.
  // Prefer splits sum over summing Lifetime Royalties (avoids double-count when
  // each format row wrongly held the full book total).
  const kenpRate = getEstimatedKenpRoyaltyRate_();
  books.forEach(b => {
    if (b.ku > 0 && kenpRate > 0 && !(b.royKenp > 0)) {
      b.royKenp = Math.round(b.ku * kenpRate * 100) / 100;
    }
    const fromSplits = b.royEbook + b.royPrint + b.royKenp;
    if (fromSplits > 0.009) {
      b.roy = Math.round(fromSplits * 100) / 100;
    }
  });

  const best = getBestOverallRanksByBook_();
  const lastRank = getLastRankUpdateByBook_();
  const out = [...books.values()]
    .sort((a, b) => a.series.localeCompare(b.series) || number_(a.num) - number_(b.num))
    .map(b => [
      b.id,
      b.title,
      b.series,
      b.num,
      b.stage,
      b.words,
      b.date,
      b.formats.size,
      b.listings,
      b.active,
      b.ranks.length ? Math.min(...b.ranks) : '',
      best.get(b.id) || '',
      b.units,
      b.ku,
      b.roy,
      b.reviews,
      b.ratings.length ? average_(b.ratings) : '',
      b.dates.length ? new Date(Math.max(...b.dates.map(d => d.getTime()))) : '',
      lastRank.get(b.id) || '',
      b.royEbook,
      b.royPrint,
      b.royKenp,
      b.kenpc === '' || b.kenpc == null ? '' : b.kenpc
    ]);

  if (out.length) catalog.getRange(2, 1, out.length, AD.CATALOG_HEADERS.length).setValues(out);
  catalog.getRange('F2:F').setNumberFormat('#,##0');
  catalog.getRange('G2:G').setNumberFormat('m/d/yyyy');
  catalog.getRange('H2:N').setNumberFormat('#,##0');
  catalog.getRange('O2:O').setNumberFormat('$#,##0.00');
  catalog.getRange('P2:P').setNumberFormat('#,##0');
  catalog.getRange('Q2:Q').setNumberFormat('0.0');
  catalog.getRange('R2:S').setNumberFormat('m/d/yyyy');
  catalog.getRange('T2:V').setNumberFormat('$#,##0.00');
  catalog.getRange('W2:W').setNumberFormat('#,##0');
}

function ensureCatalogSchema_() {
  const sh = getRequiredSheet_(AD.SHEETS.CATALOG);
  const lastCol = Math.max(sh.getLastColumn(), 1);
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(clean_);
  const expected = AD.CATALOG_HEADERS;
  const good = expected.every((h, i) => normalizeKey_(headers[i] || '') === normalizeKey_(h));
  if (good && headers.length >= expected.length) return;

  const oldRows = sh.getLastRow() >= 2
    ? sh.getRange(2, 1, sh.getLastRow() - 1, lastCol).getValues()
    : [];
  sh.getRange(1, 1, 1, expected.length).setValues([expected]);
  styleHeader_(sh.getRange(1, 1, 1, expected.length));
  if (oldRows.length) {
    const mapped = oldRows.map(r => {
      const row = new Array(expected.length).fill('');
      for (let i = 0; i < Math.min(r.length, expected.length); i++) row[i] = r[i];
      return row;
    });
    clearDataRows_(sh);
    sh.getRange(2, 1, mapped.length, expected.length).setValues(mapped);
  }
  addFilter_(sh, expected.length);
}
