/**
 * Royalty Periods — one row per KDP report window.
 * Stores period-matched KENP pages + KENP $ for estimated rate (not Global Fund final).
 */

function ensureRoyaltyPeriodsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(AD.SHEETS.ROYALTY_PERIODS);
  if (!sh) {
    sh = ss.insertSheet(AD.SHEETS.ROYALTY_PERIODS);
  }
  const headers = AD.ROYALTY_PERIOD_HEADERS;
  const lastCol = Math.max(sh.getLastColumn(), 1);
  const cur = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(clean_);
  const good = headers.every((h, i) => normalizeKey_(cur[i] || '') === normalizeKey_(h));
  if (!good || cur.length < headers.length) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    styleHeader_(sh.getRange(1, 1, 1, headers.length));
    addFilter_(sh, headers.length);
  }
  try { sh.hideSheet(); } catch (e) {}
  return sh;
}

/**
 * Working Estimated KENP Royalty Rate (USD per page).
 * Prefers a rate learned from real KENP $ ÷ pages; else Config seed.
 * Ignores absurd rates (corruption from lifetime $ ÷ period pages).
 */
function getEstimatedKenpRoyaltyRate_() {
  const sane = r => number_(r) > 0 && number_(r) <= 0.05;

  try {
    const stored = PropertiesService.getDocumentProperties().getProperty('estimatedKenpRoyaltyRate');
    if (sane(stored)) return number_(stored);
    if (stored && !sane(stored)) {
      PropertiesService.getDocumentProperties().deleteProperty('estimatedKenpRoyaltyRate');
    }
  } catch (e) {}

  const period = getLatestRoyaltyPeriod_();
  if (period && sane(period.ratePerKenp)) return number_(period.ratePerKenp);

  let pages = 0;
  let roy = 0;
  getInputRows_().forEach(r => {
    pages += number_(r[AD.COL.KU]);
    roy += number_(r[AD.COL.ROYALTY_KENP]);
  });
  if (pages > 0 && roy > 0 && sane(roy / pages)) return roy / pages;

  if (AD.ESTIMATED_KENP_RATE_USD > 0) return number_(AD.ESTIMATED_KENP_RATE_USD);
  return null;
}

function storeEstimatedKenpRoyaltyRate_(rate) {
  if (!(number_(rate) > 0) || number_(rate) > 0.05) return;
  try {
    PropertiesService.getDocumentProperties().setProperty(
      'estimatedKenpRoyaltyRate',
      String(number_(rate))
    );
  } catch (e) {}
}

/**
 * When the KDP export has KENP pages but no Royalty column:
 * estimated KENP $ = pages × estimated rate (full precision on the object).
 */
function applyEstimatedKenpRoyaltiesToTotals_(totals) {
  const rate = getEstimatedKenpRoyaltyRate_();
  if (!(rate > 0) || !totals) return { applied: false, rate: null, kenpRoyalties: 0 };
  const seen = new Set();
  let kenpRoyalties = 0;
  Object.keys(totals).forEach(key => {
    const t = totals[key];
    if (!t || seen.has(t)) return;
    seen.add(t);
    const pages = number_(t.kenp);
    if (!(pages > 0)) return;
    t.royaltyKenp = pages * rate;
    t.royaltyUsd = number_(t.royaltyEbook) + number_(t.royaltyPrint) + t.royaltyKenp;
    kenpRoyalties += t.royaltyKenp;
  });
  return { applied: kenpRoyalties > 0, rate: rate, kenpRoyalties: kenpRoyalties };
}

/**
 * Estimated KENP $ = KU pages × rate. Writes Lifetime KENP Royalties and
 * Lifetime Royalties (unit splits + KENP, or prior unit total + KENP).
 */
function syncEstimatedKenpRoyaltiesFromRate_() {
  ensureInputKuRoyaltySchema_();
  const rate = getEstimatedKenpRoyaltyRate_();
  if (!(rate > 0)) return false;
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(AD.SHEETS.INPUT);
  if (!sh || sh.getLastRow() < 2) return false;
  const values = sh.getRange(2, 1, sh.getLastRow() - 1, AD.INPUT_HEADERS.length).getValues();
  let wrote = false;
  values.forEach((r, i) => {
    if (normalizeKey_(r[AD.COL.STORE]) !== 'amazon') return;
    const pages = number_(r[AD.COL.KU]);
    if (!(pages > 0)) return;
    const nextK = Math.round(pages * rate * 100) / 100;
    const e = number_(r[AD.COL.ROYALTY_EBOOK]);
    const p = number_(r[AD.COL.ROYALTY_PRINT]);
    const prevK = number_(r[AD.COL.ROYALTY_KENP]);
    let unitPart = e + p;
    if (unitPart <= 0) {
      // Splits never filled — keep unit $ already in Lifetime Royalties, strip old KENP estimate
      unitPart = Math.max(0, number_(r[AD.COL.ROYALTIES]) - prevK);
    }
    const nextTotal = Math.round((unitPart + nextK) * 100) / 100;
    const rowNum = i + 2;
    if (Math.abs(nextK - prevK) > 0.009 || Math.abs(nextTotal - number_(r[AD.COL.ROYALTIES])) > 0.009) {
      sh.getRange(rowNum, AD.COL.ROYALTY_KENP + 1).setValue(nextK);
      sh.getRange(rowNum, AD.COL.ROYALTIES + 1).setValue(nextTotal);
      wrote = true;
    }
  });
  return wrote;
}

/**
 * Aggregate unique listing totals from a KDP import and upsert one period row.
 * When the xlsx has no KENP $, KENP royalties = Total KENP × estimated rate (same period pages).
 */
function upsertRoyaltyPeriodFromKdp_(totals, meta, fileName) {
  const sh = ensureRoyaltyPeriodsSheet_();
  const agg = aggregateKdpRoyaltyBuckets_(totals);
  const kenpcInfo = getPortfolioKenpcForCalc_();
  let kenpRoyalties = agg.royaltyKenp;
  let status = 'estimated';
  const rate = getEstimatedKenpRoyaltyRate_();
  if (!(kenpRoyalties > 0) && !(meta && meta.kenpRoyaltiesInReport) && agg.kenp > 0 && rate > 0) {
    kenpRoyalties = agg.kenp * rate;
    status = 'estimated (KENP $ = pages × rate)';
  }
  const metrics = computeKuRoyaltyMetrics_({
    totalKenp: agg.kenp,
    kenpRoyalties: kenpRoyalties,
    ebookRoyalties: agg.royaltyEbook,
    printRoyalties: agg.royaltyPrint,
    kenpc: kenpcInfo.kenpc,
    useKenpc: kenpcInfo.useForPortfolio,
    summaryRoyaltyUsd: meta && meta.summaryRoyaltyUsd != null ? meta.summaryRoyaltyUsd : null
  });
  if (metrics.ratePerKenp != null && (meta && meta.kenpRoyaltiesInReport)) {
    storeEstimatedKenpRoyaltyRate_(metrics.ratePerKenp);
  }

  const periodStart = meta && meta.minDate ? startOfDay_(meta.minDate) : '';
  const periodEnd = meta && (meta.maxDate || meta.maxKenpDate || meta.maxSalesDate)
    ? startOfDay_(meta.maxDate || meta.maxKenpDate || meta.maxSalesDate)
    : '';
  const startKey = periodStart ? dateKey_(periodStart) : '';
  const endKey = periodEnd ? dateKey_(periodEnd) : '';

  const row = [
    periodStart || '',
    periodEnd || '',
    agg.kenp,
    kenpRoyalties,
    metrics.ratePerKenp === null ? '' : metrics.ratePerKenp,
    kenpcInfo.useForPortfolio ? kenpcInfo.kenpc : '',
    metrics.equivalentReads === null ? '' : metrics.equivalentReads,
    metrics.fullReadRoyalty === null ? '' : metrics.fullReadRoyalty,
    agg.royaltyEbook,
    agg.royaltyPrint,
    metrics.totalRoyalties,
    metrics.mixEbook === null ? '' : metrics.mixEbook,
    metrics.mixPrint === null ? '' : metrics.mixPrint,
    metrics.mixKenp === null ? '' : metrics.mixKenp,
    status,
    metrics.reconcileDiff,
    metrics.reconcileOk ? 'YES' : 'FLAG',
    new Date(),
    fileName || ''
  ];

  let updateRow = null;
  if (sh.getLastRow() >= 2 && startKey && endKey) {
    const values = sh.getRange(2, 1, sh.getLastRow() - 1, AD.ROYALTY_PERIOD_HEADERS.length).getValues();
    for (let i = 0; i < values.length; i++) {
      const s = isValidDate_(values[i][0]) ? dateKey_(startOfDay_(new Date(values[i][0]))) : '';
      const e = isValidDate_(values[i][1]) ? dateKey_(startOfDay_(new Date(values[i][1]))) : '';
      if (s === startKey && e === endKey) {
        updateRow = i + 2;
        break;
      }
    }
  }

  if (updateRow) {
    sh.getRange(updateRow, 1, 1, row.length).setValues([row]);
  } else {
    sh.appendRow(row);
  }

  formatRoyaltyPeriodsSheet_(sh);
  return metrics;
}

function formatRoyaltyPeriodsSheet_(sh) {
  if (sh.getLastRow() < 2) return;
  sh.getRange('A2:B').setNumberFormat('m/d/yyyy');
  sh.getRange('C2:C').setNumberFormat('#,##0');
  sh.getRange('D2:D').setNumberFormat('$#,##0.00');
  sh.getRange('E2:E').setNumberFormat('$#,##0.00000');
  sh.getRange('F2:F').setNumberFormat('#,##0');
  sh.getRange('G2:G').setNumberFormat('0.00');
  sh.getRange('H2:K').setNumberFormat('$#,##0.00');
  sh.getRange('L2:N').setNumberFormat('0.00%');
  sh.getRange('P2:P').setNumberFormat('$#,##0.00');
  sh.getRange('R2:R').setNumberFormat('m/d/yyyy hh:mm:ss');
}

function sumManualKenpRoyalties_() {
  return getInputRows_().reduce((s, r) => s + number_(r[AD.COL.ROYALTY_KENP]), 0);
}

/**
 * Do not copy lifetime Manual Entry KENP $ into a single Royalty Periods row.
 * That mixed lifetime $ with period pages and inflated $/KENP (feedback into estimates).
 * Period rows are written only from KDP uploads via upsertRoyaltyPeriodFromKdp_.
 */
function syncLatestRoyaltyPeriodKenpFromManual_() {
  return false;
}

/** Deduped sum of listing royalty buckets from buildKdpTotalsFromRows_. */
function aggregateKdpRoyaltyBuckets_(totals) {
  const seen = new Set();
  let kenp = 0;
  let royaltyEbook = 0;
  let royaltyPrint = 0;
  let royaltyKenp = 0;
  Object.keys(totals || {}).forEach(key => {
    const t = totals[key];
    if (!t || seen.has(t)) return;
    seen.add(t);
    kenp += number_(t.kenp);
    royaltyEbook += number_(t.royaltyEbook);
    royaltyPrint += number_(t.royaltyPrint);
    royaltyKenp += number_(t.royaltyKenp);
  });
  return { kenp: kenp, royaltyEbook: royaltyEbook, royaltyPrint: royaltyPrint, royaltyKenp: royaltyKenp };
}

/**
 * Portfolio KENPC: only when exactly one published book has a non-blank KENPC.
 * Returns { kenpc, useForPortfolio, bookCountWithKenpc }.
 */
function getPortfolioKenpcForCalc_() {
  const rows = getInputRows_();
  const byBook = new Map();
  rows.forEach(r => {
    const bookId = clean_(r[AD.COL.BOOK_ID]);
    if (!bookId) return;
    const kenpc = r[AD.COL.KENPC];
    const stage = normalizeKey_(r[5]);
    if (kenpc === '' || kenpc === null || kenpc === undefined) return;
    const n = number_(kenpc);
    if (!(n > 0)) return;
    if (!byBook.has(bookId)) {
      byBook.set(bookId, { kenpc: n, published: stage === 'published' });
    } else if (stage === 'published') {
      byBook.get(bookId).published = true;
      byBook.get(bookId).kenpc = n;
    }
  });
  const publishedWith = [...byBook.values()].filter(b => b.published);
  const withKenpc = publishedWith.length ? publishedWith : [...byBook.values()];
  if (withKenpc.length === 1) {
    return { kenpc: withKenpc[0].kenpc, useForPortfolio: true, bookCountWithKenpc: 1 };
  }
  return {
    kenpc: withKenpc.length === 1 ? withKenpc[0].kenpc : '',
    useForPortfolio: false,
    bookCountWithKenpc: withKenpc.length
  };
}

/**
 * Pure metrics from period-matched inputs. Blanks (null) for divide-by-zero / missing KENPC.
 * Preserves full float precision in ratePerKenp.
 */
function computeKuRoyaltyMetrics_(input) {
  const totalKenp = number_(input.totalKenp);
  const kenpRoyalties = number_(input.kenpRoyalties);
  const ebook = number_(input.ebookRoyalties);
  const print = number_(input.printRoyalties);
  const total = ebook + print + kenpRoyalties;
  const kenpcRaw = input.kenpc;
  const hasKenpc = input.useKenpc !== false &&
    kenpcRaw !== '' && kenpcRaw !== null && kenpcRaw !== undefined && number_(kenpcRaw) > 0;
  const kenpc = hasKenpc ? number_(kenpcRaw) : 0;

  // Spec: if KENP pages blank/zero OR no KENP $, leave rate blank (not zero).
  const ratePerKenp = (totalKenp > 0 && kenpRoyalties > 0) ? (kenpRoyalties / totalKenp) : null;

  let equivalentReads = null;
  let fullReadRoyalty = null;
  if (hasKenpc && totalKenp > 0) equivalentReads = totalKenp / kenpc;
  if (hasKenpc && ratePerKenp != null) fullReadRoyalty = kenpc * ratePerKenp;

  const mixEbook = total > 0 ? ebook / total : null;
  const mixPrint = total > 0 ? print / total : null;
  const mixKenp = total > 0 ? kenpRoyalties / total : null;

  // eBook + Print + KENP must equal Total (±$0.01)
  const componentSum = ebook + print + kenpRoyalties;
  const reconcileDiff = Math.round((componentSum - total) * 100) / 100;
  const reconcileOk = Math.abs(reconcileDiff) <= 0.01;

  return {
    ratePerKenp: ratePerKenp,
    ratePer1000: ratePerKenp != null ? ratePerKenp * 1000 : null,
    centsPerKenp: ratePerKenp != null ? ratePerKenp * 100 : null,
    equivalentReads: equivalentReads,
    fullReadRoyalty: fullReadRoyalty,
    totalRoyalties: total,
    mixEbook: mixEbook,
    mixPrint: mixPrint,
    mixKenp: mixKenp,
    reconcileDiff: reconcileDiff,
    reconcileOk: reconcileOk,
    ebookRoyalties: ebook,
    printRoyalties: print,
    kenpRoyalties: kenpRoyalties,
    totalKenp: totalKenp
  };
}

/** Latest Royalty Periods row as object, or null. */
function getLatestRoyaltyPeriod_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(AD.SHEETS.ROYALTY_PERIODS);
  if (!sh || sh.getLastRow() < 2) return null;
  const values = sh.getRange(2, 1, sh.getLastRow() - 1, AD.ROYALTY_PERIOD_HEADERS.length).getValues();
  let best = null;
  let bestTime = -1;
  values.forEach(r => {
    const imported = isValidDate_(r[17]) ? new Date(r[17]).getTime() : 0;
    const end = isValidDate_(r[1]) ? new Date(r[1]).getTime() : 0;
    const score = imported || end;
    if (score >= bestTime) {
      bestTime = score;
      best = r;
    }
  });
  if (!best) return null;
  return {
    periodStart: best[0],
    periodEnd: best[1],
    totalKenp: best[2],
    kenpRoyalties: best[3],
    ratePerKenp: best[4],
    kenpc: best[5],
    equivalentReads: best[6],
    fullReadRoyalty: best[7],
    ebookRoyalties: best[8],
    printRoyalties: best[9],
    totalRoyalties: best[10],
    mixEbook: best[11],
    mixPrint: best[12],
    mixKenp: best[13],
    status: best[14],
    reconcileDiff: best[15],
    reconcileOk: best[16],
    dateImported: best[17],
    sourceFile: best[18]
  };
}
