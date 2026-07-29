/**
 * One-click KDP sales upload.
 * Parses the .xlsx in a dialog (no clutter sheets), then updates Manual Entry,
 * records a sales snapshot, and refreshes Catalog + Dashboard.
 *
 * Phase 0: never let a partial date-range report REPLACE lifetime with lower totals;
 * short-range exports ADD to lifetime instead. Use report dates for Week Ending.
 */

function openKdpReportsPage() {
  const url = AD.KDP_REPORTS_URL;
  const html = HtmlService.createHtmlOutput(
    '<div style="font-family:Arial,sans-serif;padding:12px;">' +
      '<p>Opening KDP Reports…</p>' +
      '<p>If it does not open, use this link:</p>' +
      '<p><a href="' + url + '" target="_blank">' + url + '</a></p>' +
      '<p style="margin-top:16px;">QR (phone):</p>' +
      '<img alt="KDP Reports QR" width="160" height="160" src="https://api.qrserver.com/v1/create-qr-code/?size=160x160&data=' +
      encodeURIComponent(url) + '"/>' +
      '<script>window.open(' + JSON.stringify(url) + ');</script>' +
    '</div>'
  ).setWidth(380).setHeight(360);
  SpreadsheetApp.getUi().showModalDialog(html, 'KDP Reports');
}

function uploadKdpSalesReport() {
  const html = HtmlService.createHtmlOutputFromFile('KdpUpload')
    .setWidth(520)
    .setHeight(560);
  SpreadsheetApp.getUi().showModalDialog(html, 'Upload Amazon / KDP sales data');
}

/**
 * Called from KdpUpload.html after client-side .xlsx parse.
 * payload = { fileName, sheets: { 'Combined Sales': [rowObjects...], ... } }
 */
function processKdpSalesUpload(payload) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    throw new Error('Workbook is busy. Try again in a moment.');
  }

  try {
    if (!payload || !payload.sheets) throw new Error('No report data received.');

    const found = {
      combined: payload.sheets[AD.KDP_SHEETS.COMBINED] || null,
      orders: payload.sheets[AD.KDP_SHEETS.ORDERS] || null,
      kenp: payload.sheets[AD.KDP_SHEETS.KENP] || null,
      ebook: payload.sheets[AD.KDP_SHEETS.EBOOK] || null,
      paperback: payload.sheets[AD.KDP_SHEETS.PAPERBACK] || null,
      hardcover: payload.sheets[AD.KDP_SHEETS.HARDCOVER] || null,
      summary: payload.sheets[AD.KDP_SHEETS.SUMMARY] || null
    };

    const hasAny = Object.keys(found).some(k => k !== 'summary' && found[k] && found[k].length);
    if (!hasAny) throw new Error('Report contained no usable KDP sales rows.');

    const totals = buildKdpTotalsFromRows_(found);
    const meta = analyzeKdpReportMeta_(found, payload.fileName || '', totals);
    meta.kenpRoyaltiesInReport = kenpSheetHasRoyaltyColumn_(found.kenp);

    // xlsx often has KENP pages but no Royalty $ — estimate: pages × working rate.
    let kenpRateApply = { applied: false, rate: null };
    if (!meta.kenpRoyaltiesInReport) {
      kenpRateApply = applyEstimatedKenpRoyaltiesToTotals_(totals);
    } else {
      const agg = aggregateKdpRoyaltyBuckets_(totals);
      if (agg.kenp > 0 && agg.royaltyKenp > 0) {
        storeEstimatedKenpRoyaltyRate_(agg.royaltyKenp / agg.kenp);
      }
    }

    ensureInputKuRoyaltySchema_();
    assignInternalIds_();
    maybeResetInflatedKdpMonthStore_(meta);
    repairDuplicateFormatRoyalties_();
    const summary = applyKdpTotalsToManualEntry_(totals, meta);
    recomputeLifetimeRoyaltiesFromSplits_();
    summary.fileName = payload.fileName || '';
    summary.reportMeta = {
      minDate: meta.minDate ? dateKey_(meta.minDate) : '',
      maxDate: meta.maxDate ? dateKey_(meta.maxDate) : '',
      maxSalesDate: meta.maxSalesDate ? dateKey_(meta.maxSalesDate) : '',
      maxKenpDate: meta.maxKenpDate ? dateKey_(meta.maxKenpDate) : '',
      summaryRoyaltyUsd: meta.summaryRoyaltyUsd,
      kenpRoyaltiesInReport: !!meta.kenpRoyaltiesInReport,
      kenpRateUsed: kenpRateApply.rate,
      spanDays: meta.spanDays,
      isMonthReport: !!meta.isMonthReport,
      isPartial: meta.isPartial,
      reasons: meta.reasons,
      warnings: meta.warnings || []
    };
    if (!meta.kenpRoyaltiesInReport) {
      if (kenpRateApply.applied) {
        summary.reportMeta.warnings = (summary.reportMeta.warnings || []).concat([
          'KENP Read has pages but no Royalty column. Estimated KENP $ = KENP pages × $' +
            Number(kenpRateApply.rate).toFixed(7) +
            ' per KENP (estimated rate — not the finalized Global Fund payout).'
        ]);
      } else {
        summary.reportMeta.warnings = (summary.reportMeta.warnings || []).concat([
          'KENP Read has pages but no Royalty column, and no estimated KENP rate is available yet.'
        ]);
      }
    }
    summary.royaltiesUsdBook = sumBookRoyaltyUsd_(totals);

    ensureRankHistorySchema_();
    ensureSalesHistorySchema_();
    ensureRoyaltyPeriodsSheet_();
    const periodMetrics = upsertRoyaltyPeriodFromKdp_(totals, meta, payload.fileName || '');
    summary.periodMetrics = periodMetrics;

    // Snapshot when Manual Entry lifetime changed (full replace or partial add).
    if (summary.listingsUpdated > 0) {
      const rows = getInputRows_();
      const today = getSpreadsheetToday_();
      // Week Ending / snapshot anchor from ORDER/ROYALTY dates only — not KENP.
      // KENP can post on Sunday (e.g. 7/26) and wrongly push Week Ending to next Saturday (8/1).
      const salesAnchor = meta.maxSalesDate || today;
      const snapDate = salesAnchor;
      const week = getWeekEndingDate_(salesAnchor);
      recordSalesSnapshot_(rows, snapDate, week, salesAnchor, { upsertByWeek: true });
      repairKenpOnlySundayWeekEndings_();
      consolidateSalesHistoryByWeekEnding_();
      recomputeSalesPeriodChangesFromLifetime_();
      summary.salesSnapshot = true;
      summary.snapshotWeekEnding = dateKey_(week);
      summary.snapshotDate = dateKey_(snapDate);
      if (meta.maxKenpDate && meta.maxSalesDate && meta.maxKenpDate > meta.maxSalesDate) {
        summary.reportMeta.warnings = (summary.reportMeta.warnings || []).concat([
          'KENP has newer dates (' + dateKey_(meta.maxKenpDate) +
            ') than orders/royalties (' + dateKey_(meta.maxSalesDate) +
            '). Sales History week uses the sales/royalty date so KENP-only days do not create a new week bucket.'
        ]);
      }
    } else {
      summary.salesSnapshot = false;
    }

    refreshSalesReports_();
    rebuildCatalogSummary_();
    refreshDashboard_();
    rebuildReconciliationSheet_();
    lockAutomaticSheets();

    console.log('KDP upload summary: ' + JSON.stringify(summary));
    return formatKdpImportSummary_(summary);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Detect report date range and whether the file looks like a partial period report.
 * Separates order/royalty dates from KENP dates (KENP often posts a day later).
 */
function analyzeKdpReportMeta_(found, fileName, totals) {
  const salesDates = collectKdpDatesFromSheets_(found, ['orders', 'combined', 'ebook', 'paperback', 'hardcover']);
  const kenpDates = collectKdpDatesFromSheets_(found, ['kenp']);
  const dates = salesDates.concat(kenpDates);

  let minDate = null;
  let maxDate = null;
  let maxSalesDate = null;
  let maxKenpDate = null;
  dates.forEach(d => {
    if (!minDate || d < minDate) minDate = d;
    if (!maxDate || d > maxDate) maxDate = d;
  });
  salesDates.forEach(d => {
    if (!maxSalesDate || d > maxSalesDate) maxSalesDate = d;
  });
  kenpDates.forEach(d => {
    if (!maxKenpDate || d > maxKenpDate) maxKenpDate = d;
  });

  const spanDays = minDate && maxDate
    ? Math.round((startOfDay_(maxDate) - startOfDay_(minDate)) / 86400000) + 1
    : null;

  const reasons = [];
  const warnings = [];
  const name = String(fileName || '').toLowerCase();
  // KDP Dashboard typical options: Month / Today / Yesterday (no All-time).
  const monthStyle = looksLikeKdpMonthReport_(minDate, maxDate, spanDays, name);
  const dayStyle =
    !monthStyle &&
    spanDays != null &&
    spanDays > 0 &&
    spanDays <= 2;

  const lowerCount = countReportLowerThanLifetime_(totals || {});
  const mostlyLower =
    lowerCount.compared > 0 &&
    lowerCount.lower >= Math.max(1, Math.ceil(lowerCount.compared * 0.5));

  if (monthStyle) {
    warnings.push(
      'Month report — Manual Entry updates to these totals (existing higher lifetime values are kept).'
    );
  } else if (dayStyle || /today|yesterday|last\s*\d+\s*day|partial|custom.?range/i.test(name)) {
    if (mostlyLower || dayStyle) {
      reasons.push(
        'Today/Yesterday-style report (short range). Same-week uploads replace that week’s contribution.'
      );
    }
    if (mostlyLower) {
      reasons.push(
        'Report totals are lower than current Manual Entry for ' +
          lowerCount.lower + ' of ' + lowerCount.compared +
          ' matched listing metric(s).'
      );
    }
  } else if (mostlyLower && spanDays != null && spanDays < AD.KDP_PARTIAL_MAX_SPAN_DAYS) {
    reasons.push(
      'Report date span is ' + spanDays +
        ' day(s) (' + dateKey_(minDate) + ' → ' + dateKey_(maxDate) +
        ') and totals are below current lifetime — treating as short-range.'
    );
  }

  if (!salesDates.length) {
    warnings.push(
      'No order/royalty dates found. Week Ending will not use KENP-only dates.'
    );
  }

  const summaryRoyaltyUsd = readSummaryRoyaltyUsd_(found.summary);

  return {
    minDate: minDate,
    maxDate: maxDate,
    maxSalesDate: maxSalesDate,
    maxKenpDate: maxKenpDate,
    summaryRoyaltyUsd: summaryRoyaltyUsd,
    spanDays: spanDays,
    isMonthReport: monthStyle,
    isPartial: !monthStyle && reasons.length > 0,
    reasons: reasons,
    warnings: warnings,
    lowerCount: lowerCount
  };
}

/** Month / month-to-date is the normal KDP Dashboard export (no All-time option). */
function looksLikeKdpMonthReport_(minDate, maxDate, spanDays, fileName) {
  const name = String(fileName || '').toLowerCase();
  if (/today|yesterday/.test(name) && !/month/.test(name)) return false;
  if (/month|month.?to.?date|mtd/i.test(name)) return true;
  if (!minDate || !maxDate || spanDays == null || spanDays < 8) return false;
  try {
    const a = new Date(minDate);
    const b = new Date(maxDate);
    if (a.getFullYear() !== b.getFullYear() || a.getMonth() !== b.getMonth()) return false;
    // Same calendar month and at least ~a week (month-to-date mid-month is fine).
    return spanDays >= 8;
  } catch (e) {
    return false;
  }
}

function collectKdpDates_(found) {
  return collectKdpDatesFromSheets_(found, Object.keys(found || {}));
}

function collectKdpDatesFromSheets_(found, keys) {
  const out = [];
  const asRows = v => Array.isArray(v) ? v : [];
  (keys || []).forEach(key => {
    asRows(found[key]).forEach(row => {
      if (!row || typeof row !== 'object') return;
      Object.keys(row).forEach(col => {
        if (!/date|start|end|period/i.test(col)) return;
        const d = parseLooseDate_(row[col]);
        if (d) out.push(d);
      });
    });
  });
  return out;
}

function readSummaryRoyaltyUsd_(summaryRows) {
  const rows = Array.isArray(summaryRows) ? summaryRows : [];
  if (!rows.length) return null;
  const row = rows[0];
  const v = row['Royalty (USD)'] != null ? row['Royalty (USD)'] : row['Royalty(USD)'];
  if (v === '' || v == null) return null;
  return number_(v);
}

function sumBookRoyaltyUsd_(totals) {
  const seen = new Set();
  let sum = 0;
  Object.keys(totals || {}).forEach(k => {
    const t = totals[k];
    if (!t || seen.has(t)) return;
    seen.add(t);
    sum += number_(t.royaltyUsd);
  });
  return sum;
}

/**
 * If Manual Entry lifetime royalties are far above KDP Summary, stored month
 * slices were almost certainly double-counted or cumulative. Clear them so the
 * next month uploads rebuild lifetime cleanly.
 */
function maybeResetInflatedKdpMonthStore_(meta) {
  const summaryUsd = meta && number_(meta.summaryRoyaltyUsd);
  if (!(summaryUsd > 0)) return false;
  let manualUsd = 0;
  getInputRows_().forEach(r => {
    if (normalizeKey_(r[AD.COL.STORE]) !== 'amazon') return;
    manualUsd += number_(r[AD.COL.ROYALTIES]);
  });
  if (!(manualUsd > summaryUsd * 1.35)) return false;
  try {
    PropertiesService.getDocumentProperties().deleteProperty('AD_KDP_MONTH_CONTRIBS');
    PropertiesService.getDocumentProperties().deleteProperty('AD_KDP_LAST_PARTIAL_FP');
    PropertiesService.getDocumentProperties().deleteProperty('AD_KDP_WEEK_PARTIAL_ADDS');
    if (meta.warnings) {
      meta.warnings.push(
        'Cleared stored KDP month contributions — Manual Entry royalties ($' +
          manualUsd.toFixed(2) + ') were far above KDP Summary ($' +
          summaryUsd.toFixed(2) +
          '). Re-upload each prior month’s report to rebuild lifetime.'
      );
    }
    return true;
  } catch (e) {
    return false;
  }
}

/**
 * When Kindle + print rows each hold the same full-book royalty splits (old bug),
 * keep format-scoped fields only so Catalog no longer triples KDP Summary.
 */
function repairDuplicateFormatRoyalties_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(AD.SHEETS.INPUT);
  if (!sh || sh.getLastRow() < 2) return 0;
  const values = sh.getRange(2, 1, sh.getLastRow() - 1, AD.INPUT_HEADERS.length).getValues();
  const byBook = new Map();
  values.forEach((r, i) => {
    if (normalizeKey_(r[AD.COL.STORE]) !== 'amazon') return;
    const id = clean_(r[AD.COL.BOOK_ID]);
    if (!id) return;
    if (!byBook.has(id)) byBook.set(id, []);
    byBook.get(id).push({ r: r, i: i });
  });

  let fixed = 0;
  byBook.forEach(list => {
    if (list.length < 2) return;
    const totals = list.map(x => number_(x.r[AD.COL.ROYALTIES]));
    const nonzero = totals.filter(t => t > 0.009);
    if (nonzero.length < 2) return;
    const first = nonzero[0];
    const allSame = nonzero.every(t => Math.abs(t - first) < 0.02);
    if (!allSame) return;

    // Same total on multiple formats → classic double-count.
    const ebook = list.find(x => listingFormatKind_(x.r[AD.COL.FORMAT]) === 'ebook');
    const prints = list.filter(x => {
      const k = listingFormatKind_(x.r[AD.COL.FORMAT]);
      return k === 'paperback' || k === 'hardcover';
    });
    if (!ebook || !prints.length) return;

    const hasSplits = number_(ebook.r[AD.COL.ROYALTY_EBOOK]) +
      number_(ebook.r[AD.COL.ROYALTY_PRINT]) +
      number_(ebook.r[AD.COL.ROYALTY_KENP]) > 0.009;
    if (!hasSplits) return;

    const srcE = number_(ebook.r[AD.COL.ROYALTY_EBOOK]);
    const srcP = number_(ebook.r[AD.COL.ROYALTY_PRINT]);
    const srcK = number_(ebook.r[AD.COL.ROYALTY_KENP]);
    const srcU = number_(ebook.r[AD.COL.UNITS]);
    const srcKenp = number_(ebook.r[AD.COL.KU]);

    const eRow = ebook.i + 2;
    sh.getRange(eRow, AD.COL.ROYALTY_EBOOK + 1).setValue(srcE);
    sh.getRange(eRow, AD.COL.ROYALTY_PRINT + 1).setValue(0);
    sh.getRange(eRow, AD.COL.ROYALTY_KENP + 1).setValue(srcK);
    sh.getRange(eRow, AD.COL.ROYALTIES + 1).setValue(Math.round((srcE + srcK) * 100) / 100);
    sh.getRange(eRow, AD.COL.UNITS + 1).setValue(srcU);
    sh.getRange(eRow, AD.COL.KU + 1).setValue(srcKenp);

    prints.forEach((x, idx) => {
      const rowNum = x.i + 2;
      const printAmt = idx === 0 ? srcP : 0;
      sh.getRange(rowNum, AD.COL.ROYALTY_EBOOK + 1).setValue(0);
      sh.getRange(rowNum, AD.COL.ROYALTY_PRINT + 1).setValue(printAmt);
      sh.getRange(rowNum, AD.COL.ROYALTY_KENP + 1).setValue(0);
      sh.getRange(rowNum, AD.COL.ROYALTIES + 1).setValue(printAmt);
      sh.getRange(rowNum, AD.COL.UNITS + 1).setValue(0);
      sh.getRange(rowNum, AD.COL.KU + 1).setValue(0);
    });
    fixed++;
  });
  return fixed;
}

function parseLooseDate_(v) {
  if (v === null || v === undefined || v === '') return null;
  if (Object.prototype.toString.call(v) === '[object Date]' && isValidDate_(v)) {
    return startOfDay_(v);
  }
  if (typeof v === 'number' && isFinite(v)) {
    // Excel serial (SheetJS sometimes leaves numbers if raw)
    if (v > 20000 && v < 80000) {
      const epoch = new Date(Date.UTC(1899, 11, 30));
      epoch.setUTCDate(epoch.getUTCDate() + Math.floor(v));
      return startOfDay_(epoch);
    }
  }
  const s = clean_(v);
  if (!s) return null;
  // Prefer yyyy-mm-dd / mm/dd/yyyy
  const iso = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (iso) {
    const d = new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3]));
    return isValidDate_(d) ? startOfDay_(d) : null;
  }
  const d = new Date(s);
  return isValidDate_(d) ? startOfDay_(d) : null;
}

function countReportLowerThanLifetime_(totals) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(AD.SHEETS.INPUT);
  const result = { compared: 0, lower: 0 };
  if (!sh || sh.getLastRow() < 2) return result;
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, AD.INPUT_HEADERS.length).getValues();
  rows.forEach(r => {
    if (normalizeKey_(r[AD.COL.STORE]) !== 'amazon') return;
    const keys = listingIdentifierKeys_(r);
    let hit = null;
    for (let k = 0; k < keys.length; k++) {
      if (totals[keys[k]]) {
        hit = totals[keys[k]];
        break;
      }
    }
    if (!hit) return;
    const curU = number_(r[AD.COL.UNITS]);
    const curK = number_(r[AD.COL.KU]);
    const curR = number_(r[AD.COL.ROYALTIES]);
    if (curU > 0) {
      result.compared++;
      if (hit.units < curU) result.lower++;
    }
    if (curK > 0) {
      result.compared++;
      if (hit.kenp < curK) result.lower++;
    }
    if (curR > 0) {
      result.compared++;
      if (hit.royaltyUsd < curR - 0.009) result.lower++;
    }
  });
  return result;
}

/**
 * Apply report totals to lifetime columns.
 * - Month reports: store that month’s slice; lifetime = sum of all uploaded months
 * - Today/Yesterday: REPLACE this week’s contribution only
 * - Other: SET lifetime (never decrease)
 */
function applyKdpTotalsToManualEntry_(totals, meta) {
  const sh = getRequiredSheet_(AD.SHEETS.INPUT);
  const summary = {
    identifiersInReport: Object.keys(totals).length,
    listingsUpdated: 0,
    listingsUnmatched: 0,
    unitsSet: 0,
    kuSet: 0,
    royaltiesSet: 0,
    ebookRoyaltiesSet: 0,
    printRoyaltiesSet: 0,
    kenpRoyaltiesSet: 0,
    unmatchedIds: [],
    blockedPartial: false,
    appliedAsAddition: false,
    appliedAsWeekReplace: false,
    appliedAsMonthContrib: false,
    monthKey: '',
    skippedDuplicatePartial: false,
    skippedLowerUnits: 0,
    skippedLowerKenp: 0,
    skippedLowerRoyalties: 0,
    fieldsWritten: 0,
    partialReasons: (meta && meta.reasons) || []
  };

  if (meta && meta.isMonthReport) {
    return applyKdpMonthContribToManualEntry_(totals, meta, summary);
  }
  if (meta && meta.isPartial) {
    return applyKdpWeekReplaceToManualEntry_(totals, meta, summary);
  }

  if (sh.getLastRow() < 2) return summary;

  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, AD.INPUT_HEADERS.length).getValues();
  const matchedKeys = new Set();
  const today = getSpreadsheetToday_();
  const claimMap = new Map();

  // Apply ebook listings before print so units/KENP claim correctly when ASIN↔ISBN aliased.
  const order = rows
    .map((r, i) => ({ r: r, i: i }))
    .filter(x => normalizeKey_(x.r[AD.COL.STORE]) === 'amazon')
    .sort((a, b) => {
      const ka = listingFormatKind_(a.r[AD.COL.FORMAT]);
      const kb = listingFormatKind_(b.r[AD.COL.FORMAT]);
      const rank = k => (k === 'ebook' ? 0 : k === 'paperback' ? 1 : k === 'hardcover' ? 2 : 3);
      return rank(ka) - rank(kb);
    });

  order.forEach(({ r, i }) => {
    const keys = listingIdentifierKeys_(r);
    if (!keys.length) return;

    let hit = null;
    let hitKey = '';
    for (let k = 0; k < keys.length; k++) {
      if (totals[keys[k]]) {
        hit = totals[keys[k]];
        hitKey = keys[k];
        break;
      }
    }
    if (!hit) return;

    matchedKeys.add(hitKey);
    const rowNum = i + 2;
    const curU = r[AD.COL.UNITS];
    const curK = r[AD.COL.KU];
    const curR = r[AD.COL.ROYALTIES];
    let wrote = false;
    const notes = [];
    const split = splitKdpHitForListing_(hit, r[AD.COL.FORMAT], claimBucketForHit_(claimMap, hit));

    // Units
    if (split.units !== '') {
      const next = number_(split.units);
      const hasCur = curU !== '' && curU !== null && curU !== undefined;
      if (hasCur && next < number_(curU)) {
        summary.skippedLowerUnits++;
        notes.push('units not lowered');
      } else if (!hasCur || next >= number_(curU)) {
        if (!hasCur || next !== number_(curU)) {
          sh.getRange(rowNum, AD.COL.UNITS + 1).setValue(next);
          summary.unitsSet += next;
          wrote = true;
        }
      }
    }

    // KENP pages
    if (split.kenp !== '') {
      const next = number_(split.kenp);
      const hasCur = curK !== '' && curK !== null && curK !== undefined;
      if (hasCur && next < number_(curK)) {
        summary.skippedLowerKenp++;
        notes.push('KENP not lowered');
      } else if (!hasCur || next >= number_(curK)) {
        if (!hasCur || next !== number_(curK)) {
          sh.getRange(rowNum, AD.COL.KU + 1).setValue(next);
          summary.kuSet += next;
          wrote = true;
        }
      }
    }

    // Format-scoped royalty splits (always replace on full apply — fixes prior double-counts).
    const nextEbook = number_(split.royaltyEbook);
    const nextPrint = number_(split.royaltyPrint);
    const nextKenpR = number_(split.royaltyKenp);
    const nextTotal = number_(split.royaltyUsd);
    if (nextEbook || nextPrint || nextKenpR || nextTotal || curR !== '' && curR != null) {
      if (
        Math.abs(nextTotal - number_(curR)) > 0.009 ||
        Math.abs(nextEbook - number_(r[AD.COL.ROYALTY_EBOOK])) > 0.009 ||
        Math.abs(nextPrint - number_(r[AD.COL.ROYALTY_PRINT])) > 0.009 ||
        Math.abs(nextKenpR - number_(r[AD.COL.ROYALTY_KENP])) > 0.009
      ) {
        sh.getRange(rowNum, AD.COL.ROYALTY_EBOOK + 1).setValue(nextEbook);
        sh.getRange(rowNum, AD.COL.ROYALTY_PRINT + 1).setValue(nextPrint);
        sh.getRange(rowNum, AD.COL.ROYALTY_KENP + 1).setValue(nextKenpR);
        sh.getRange(rowNum, AD.COL.ROYALTIES + 1).setValue(nextTotal);
        summary.royaltiesSet += nextTotal;
        summary.ebookRoyaltiesSet += nextEbook;
        summary.printRoyaltiesSet += nextPrint;
        summary.kenpRoyaltiesSet += nextKenpR;
        wrote = true;
      }
    }

    if (wrote) {
      summary.listingsUpdated++;
      summary.fieldsWritten++;
      sh.getRange(rowNum, AD.COL.LAST_DATA_DATE + 1).setValue(
        meta && meta.maxDate ? meta.maxDate : today
      );
      const rangeNote = meta && meta.minDate && meta.maxDate
        ? ' KDP report ' + dateKey_(meta.minDate) + ' → ' + dateKey_(meta.maxDate) + '.'
        : '';
      sh.getRange(rowNum, AD.COL.PROCESS_STATUS + 1).setValue(
        'KDP lifetime update applied.' + rangeNote + (notes.length ? ' (' + notes.join('; ') + ')' : '')
      );
    } else if (notes.length) {
      sh.getRange(rowNum, AD.COL.PROCESS_STATUS + 1).setValue(
        'KDP report matched but lifetime not overwritten (' + notes.join('; ') + ').'
      );
    }
  });

  Object.keys(totals).forEach(id => {
    if (!matchedKeys.has(id)) {
      const already = [...matchedKeys].some(k => totals[k] === totals[id]);
      if (already) return;
      summary.listingsUnmatched++;
      summary.unmatchedIds.push(id);
    }
  });

  // Non-month full apply — drop short-range week tracking only.
  if (summary.listingsUpdated > 0) {
    try {
      PropertiesService.getDocumentProperties().deleteProperty('AD_KDP_WEEK_PARTIAL_ADDS');
      PropertiesService.getDocumentProperties().deleteProperty('AD_KDP_LAST_PARTIAL_FP');
    } catch (e) {}
  }

  return summary;
}

/**
 * Month download workflow (KDP has no All-time):
 * Each file replaces that calendar month’s contribution; lifetime = sum of months.
 * - Upload July now, again in November → only July slice is replaced
 * - Miss August → August missing until you upload it (lifetime undercounts until then)
 */
function applyKdpMonthContribToManualEntry_(totals, meta, summary) {
  summary = summary || {};
  summary.appliedAsMonthContrib = true;
  summary.blockedPartial = false;

  const anchor = (meta && (meta.maxDate || meta.minDate)) || getSpreadsheetToday_();
  const monthKey = Utilities.formatDate(startOfDay_(anchor), AD.TZ, 'yyyy-MM');
  summary.monthKey = monthKey;

  const fingerprint = 'M|' + monthKey + '|' + kdpPartialFingerprint_(totals, meta);
  const props = PropertiesService.getDocumentProperties();
  if (fingerprint && fingerprint === String(props.getProperty('AD_KDP_LAST_PARTIAL_FP') || '')) {
    summary.skippedDuplicatePartial = true;
    summary.partialReasons = (meta.reasons || []).concat([
      'Identical ' + monthKey + ' month report already applied — skipped.'
    ]);
    return summary;
  }

  const sh = getRequiredSheet_(AD.SHEETS.INPUT);
  if (sh.getLastRow() < 2) return summary;

  const store = readKdpMonthContribs_();
  const monthSlice = {};
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, AD.INPUT_HEADERS.length).getValues();
  const matchedKeys = new Set();
  const listingByKey = {};
  const claimMap = new Map();

  const order = rows
    .map((r, i) => ({ r: r, i: i }))
    .filter(x => normalizeKey_(x.r[AD.COL.STORE]) === 'amazon' && clean_(x.r[AD.COL.LISTING_ID]))
    .sort((a, b) => {
      const ka = listingFormatKind_(a.r[AD.COL.FORMAT]);
      const kb = listingFormatKind_(b.r[AD.COL.FORMAT]);
      const rank = k => (k === 'ebook' ? 0 : k === 'paperback' ? 1 : k === 'hardcover' ? 2 : 3);
      return rank(ka) - rank(kb);
    });

  order.forEach(({ r, i }) => {
    const listing = clean_(r[AD.COL.LISTING_ID]);
    listingByKey[listing] = { row: r, rowNum: i + 2 };
    const keys = listingIdentifierKeys_(r);
    let hit = null;
    let hitKey = '';
    for (let k = 0; k < keys.length; k++) {
      if (totals[keys[k]]) {
        hit = totals[keys[k]];
        hitKey = keys[k];
        break;
      }
    }
    if (!hit) return;
    matchedKeys.add(hitKey);
    const split = splitKdpHitForListing_(hit, r[AD.COL.FORMAT], claimBucketForHit_(claimMap, hit));
    monthSlice[listing] = {
      u: number_(split.units === '' ? 0 : split.units),
      k: number_(split.kenp === '' ? 0 : split.kenp),
      e: number_(split.royaltyEbook),
      p: number_(split.royaltyPrint),
      kr: number_(split.royaltyKenp)
    };
  });

  store[monthKey] = monthSlice;

  // Lifetime per listing = sum of every stored month slice.
  const sums = {};
  Object.keys(store).forEach(mk => {
    const slice = store[mk] || {};
    Object.keys(slice).forEach(listing => {
      if (!sums[listing]) sums[listing] = { u: 0, k: 0, e: 0, p: 0, kr: 0 };
      const s = slice[listing];
      sums[listing].u += number_(s.u);
      sums[listing].k += number_(s.k);
      sums[listing].e += number_(s.e);
      sums[listing].p += number_(s.p);
      sums[listing].kr += number_(s.kr);
    });
  });

  const today = getSpreadsheetToday_();
  Object.keys(sums).forEach(listing => {
    const info = listingByKey[listing];
    if (!info) return;
    const sum = sums[listing];
    const nextEbook = Math.round(sum.e * 100) / 100;
    const nextPrint = Math.round(sum.p * 100) / 100;
    const nextKenpR = Math.round(sum.kr * 100) / 100;
    const nextTotal = Math.round((nextEbook + nextPrint + nextKenpR) * 100) / 100;
    const rowNum = info.rowNum;
    const r = info.row;

    const changed =
      number_(r[AD.COL.UNITS]) !== sum.u ||
      number_(r[AD.COL.KU]) !== sum.k ||
      Math.abs(number_(r[AD.COL.ROYALTY_EBOOK]) - nextEbook) > 0.009 ||
      Math.abs(number_(r[AD.COL.ROYALTY_PRINT]) - nextPrint) > 0.009 ||
      Math.abs(number_(r[AD.COL.ROYALTY_KENP]) - nextKenpR) > 0.009;

    sh.getRange(rowNum, AD.COL.UNITS + 1).setValue(sum.u);
    sh.getRange(rowNum, AD.COL.KU + 1).setValue(sum.k);
    sh.getRange(rowNum, AD.COL.ROYALTY_EBOOK + 1).setValue(nextEbook);
    sh.getRange(rowNum, AD.COL.ROYALTY_PRINT + 1).setValue(nextPrint);
    sh.getRange(rowNum, AD.COL.ROYALTY_KENP + 1).setValue(nextKenpR);
    sh.getRange(rowNum, AD.COL.ROYALTIES + 1).setValue(nextTotal);

    if (changed || monthSlice[listing]) {
      summary.listingsUpdated++;
      summary.fieldsWritten++;
      const slice = monthSlice[listing] || { u: 0, k: 0, e: 0, p: 0, kr: 0 };
      summary.unitsSet += number_(slice.u);
      summary.kuSet += number_(slice.k);
      summary.royaltiesSet += Math.round((number_(slice.e) + number_(slice.p) + number_(slice.kr)) * 100) / 100;
      summary.ebookRoyaltiesSet += number_(slice.e);
      summary.printRoyaltiesSet += number_(slice.p);
      summary.kenpRoyaltiesSet += number_(slice.kr);
      sh.getRange(rowNum, AD.COL.LAST_DATA_DATE + 1).setValue(
        meta && meta.maxDate ? meta.maxDate : today
      );
      sh.getRange(rowNum, AD.COL.PROCESS_STATUS + 1).setValue(
        'KDP month ' + monthKey + ' stored. Lifetime = sum of months (' +
          Object.keys(store).sort().join(', ') + ').'
      );
    }
  });

  Object.keys(totals).forEach(id => {
    if (!matchedKeys.has(id)) {
      const already = [...matchedKeys].some(k => totals[k] === totals[id]);
      if (already) return;
      summary.listingsUnmatched++;
      summary.unmatchedIds.push(id);
    }
  });

  if (summary.listingsUpdated > 0 || Object.keys(monthSlice).length) {
    // Count update even when values unchanged so Sales History can refresh.
    if (!summary.listingsUpdated && Object.keys(monthSlice).length) {
      summary.listingsUpdated = Object.keys(monthSlice).length;
    }
    writeKdpMonthContribs_(store);
    props.setProperty('AD_KDP_LAST_PARTIAL_FP', fingerprint);
    summary.monthsStored = Object.keys(store).sort();
  }
  return summary;
}

function readKdpMonthContribs_() {
  try {
    const raw = PropertiesService.getDocumentProperties().getProperty('AD_KDP_MONTH_CONTRIBS');
    if (!raw) return {};
    const obj = JSON.parse(raw);
    return obj && typeof obj === 'object' ? obj : {};
  } catch (e) {
    return {};
  }
}

function writeKdpMonthContribs_(store) {
  PropertiesService.getDocumentProperties().setProperty(
    'AD_KDP_MONTH_CONTRIBS',
    JSON.stringify(store || {})
  );
}

/**
 * Dashboard helper: which KDP month slices are stored, and any gaps through the current month.
 * @returns {{ monthsOnFile: string, gapMessage: string, hasGap: boolean, gaps: string[] }}
 */
function getKdpMonthGapStatus_() {
  const store = readKdpMonthContribs_();
  const keys = Object.keys(store || {})
    .filter(k => /^\d{4}-\d{2}$/.test(k))
    .sort();
  const today = getSpreadsheetToday_();
  const currentKey = Utilities.formatDate(startOfDay_(today), AD.TZ, 'yyyy-MM');

  if (!keys.length) {
    return {
      monthsOnFile: '(none)',
      gapMessage: 'No Month uploads yet — Author Dashboard → Upload KDP Sales Report (Month)',
      hasGap: true,
      gaps: []
    };
  }

  const startKey = keys[0];
  const expected = enumerateYearMonths_(startKey, currentKey);
  const have = new Set(keys);
  const gaps = expected.filter(m => !have.has(m));
  const pastGaps = gaps.filter(m => m < currentKey);
  const missingCurrent = gaps.indexOf(currentKey) >= 0;

  const parts = [];
  if (pastGaps.length) {
    parts.push('Missing: ' + pastGaps.map(formatYearMonthLabel_).join(', '));
  }
  if (missingCurrent) {
    parts.push('Current month (' + formatYearMonthLabel_(currentKey) + ') not uploaded yet');
  }

  return {
    monthsOnFile: keys.map(formatYearMonthLabel_).join(', '),
    gapMessage: parts.length ? parts.join(' · ') : 'OK — no gaps through ' + formatYearMonthLabel_(currentKey),
    hasGap: parts.length > 0,
    gaps: gaps
  };
}

function enumerateYearMonths_(startKey, endKey) {
  const out = [];
  const sm = String(startKey || '').match(/^(\d{4})-(\d{2})$/);
  const em = String(endKey || '').match(/^(\d{4})-(\d{2})$/);
  if (!sm || !em) return out;
  let y = Number(sm[1]);
  let m = Number(sm[2]);
  const ey = Number(em[1]);
  const emon = Number(em[2]);
  let guard = 0;
  while ((y < ey || (y === ey && m <= emon)) && guard < 240) {
    out.push(y + '-' + (m < 10 ? '0' : '') + m);
    m++;
    if (m > 12) {
      m = 1;
      y++;
    }
    guard++;
  }
  return out;
}

function formatYearMonthLabel_(key) {
  const m = String(key || '').match(/^(\d{4})-(\d{2})$/);
  if (!m) return String(key || '');
  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const idx = Number(m[2]) - 1;
  return (months[idx] || m[2]) + ' ' + m[1];
}

/**
 * Date-range KDP exports for the current week:
 * undo the previous same-week upload contribution, then apply this file’s amounts.
 * So multiple uploads in one week replace (wipe prior week delta), they do not stack.
 */
function applyKdpWeekReplaceToManualEntry_(totals, meta, summary) {
  summary = summary || {};
  summary.appliedAsWeekReplace = true;
  summary.appliedAsAddition = true; // keep summary flag for older UI text paths
  summary.blockedPartial = false;

  const anchor = (meta && (meta.maxSalesDate || meta.maxDate)) || getSpreadsheetToday_();
  const weekKey = dateKey_(getWeekEndingDate_(anchor));
  const fingerprint = kdpPartialFingerprint_(totals, meta);
  const props = PropertiesService.getDocumentProperties();
  const lastFp = String(props.getProperty('AD_KDP_LAST_PARTIAL_FP') || '');
  if (fingerprint && fingerprint === lastFp) {
    summary.skippedDuplicatePartial = true;
    summary.partialReasons = (meta.reasons || []).concat([
      'Identical report already applied for this week — skipped.'
    ]);
    return summary;
  }

  const sh = getRequiredSheet_(AD.SHEETS.INPUT);
  if (sh.getLastRow() < 2) return summary;

  const weekAdds = readKdpWeekPartialAdds_();
  const prevWeek = weekAdds[weekKey] || {};
  const nextWeek = {};
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, AD.INPUT_HEADERS.length).getValues();
  const matchedKeys = new Set();
  const today = getSpreadsheetToday_();
  const claimMap = new Map();

  const order = rows
    .map((r, i) => ({ r: r, i: i }))
    .filter(x => normalizeKey_(x.r[AD.COL.STORE]) === 'amazon' && clean_(x.r[AD.COL.LISTING_ID]))
    .sort((a, b) => {
      const ka = listingFormatKind_(a.r[AD.COL.FORMAT]);
      const kb = listingFormatKind_(b.r[AD.COL.FORMAT]);
      const rank = k => (k === 'ebook' ? 0 : k === 'paperback' ? 1 : k === 'hardcover' ? 2 : 3);
      return rank(ka) - rank(kb);
    });

  order.forEach(({ r, i }) => {
    const listing = clean_(r[AD.COL.LISTING_ID]);
    const keys = listingIdentifierKeys_(r);
    if (!keys.length || !listing) return;

    let hit = null;
    let hitKey = '';
    for (let k = 0; k < keys.length; k++) {
      if (totals[keys[k]]) {
        hit = totals[keys[k]];
        hitKey = keys[k];
        break;
      }
    }
    if (!hit) return;
    matchedKeys.add(hitKey);

    const rowNum = i + 2;
    const split = splitKdpHitForListing_(hit, r[AD.COL.FORMAT], claimBucketForHit_(claimMap, hit));
    const newU = number_(split.units === '' ? 0 : split.units);
    const newK = number_(split.kenp === '' ? 0 : split.kenp);
    const newEbook = number_(split.royaltyEbook);
    const newPrint = number_(split.royaltyPrint);
    const newKenpR = number_(split.royaltyKenp);
    const prev = prevWeek[listing] || { u: 0, k: 0, e: 0, p: 0, kr: 0 };

    // lifetime - previous same-week contribution + this report’s contribution
    const nextU = Math.max(0, number_(r[AD.COL.UNITS]) - number_(prev.u) + newU);
    const nextK = Math.max(0, number_(r[AD.COL.KU]) - number_(prev.k) + newK);
    const nextEbook = Math.round((number_(r[AD.COL.ROYALTY_EBOOK]) - number_(prev.e) + newEbook) * 100) / 100;
    const nextPrint = Math.round((number_(r[AD.COL.ROYALTY_PRINT]) - number_(prev.p) + newPrint) * 100) / 100;
    const nextKenpR = Math.round((number_(r[AD.COL.ROYALTY_KENP]) - number_(prev.kr) + newKenpR) * 100) / 100;
    const nextTotal = Math.round((Math.max(0, nextEbook) + Math.max(0, nextPrint) + Math.max(0, nextKenpR)) * 100) / 100;

    sh.getRange(rowNum, AD.COL.UNITS + 1).setValue(nextU);
    sh.getRange(rowNum, AD.COL.KU + 1).setValue(nextK);
    sh.getRange(rowNum, AD.COL.ROYALTY_EBOOK + 1).setValue(Math.max(0, nextEbook));
    sh.getRange(rowNum, AD.COL.ROYALTY_PRINT + 1).setValue(Math.max(0, nextPrint));
    sh.getRange(rowNum, AD.COL.ROYALTY_KENP + 1).setValue(Math.max(0, nextKenpR));
    sh.getRange(rowNum, AD.COL.ROYALTIES + 1).setValue(nextTotal);

    nextWeek[listing] = { u: newU, k: newK, e: newEbook, p: newPrint, kr: newKenpR };

    summary.listingsUpdated++;
    summary.fieldsWritten++;
    summary.unitsSet += newU;
    summary.kuSet += newK;
    summary.royaltiesSet += Math.round((newEbook + newPrint + newKenpR) * 100) / 100;
    summary.ebookRoyaltiesSet += newEbook;
    summary.printRoyaltiesSet += newPrint;
    summary.kenpRoyaltiesSet += newKenpR;

    sh.getRange(rowNum, AD.COL.LAST_DATA_DATE + 1).setValue(
      meta && meta.maxDate ? meta.maxDate : today
    );
    const rangeNote = meta && meta.minDate && meta.maxDate
      ? ' KDP period ' + dateKey_(meta.minDate) + ' → ' + dateKey_(meta.maxDate) + '.'
      : '';
    sh.getRange(rowNum, AD.COL.PROCESS_STATUS + 1).setValue(
      'KDP week REPLACE applied (week ending ' + weekKey + ').' + rangeNote +
        ' This week: ' + newU + ' units, ' + newK + ' KENP, $' +
        (Math.round((newEbook + newPrint + newKenpR) * 100) / 100).toFixed(2) + '.'
    );
  });

  Object.keys(totals).forEach(id => {
    if (!matchedKeys.has(id)) {
      const already = [...matchedKeys].some(k => totals[k] === totals[id]);
      if (already) return;
      summary.listingsUnmatched++;
      summary.unmatchedIds.push(id);
    }
  });

  if (summary.listingsUpdated > 0) {
    weekAdds[weekKey] = nextWeek;
    props.setProperty('AD_KDP_WEEK_PARTIAL_ADDS', JSON.stringify(weekAdds));
    if (fingerprint) props.setProperty('AD_KDP_LAST_PARTIAL_FP', fingerprint);
    summary.snapshotWeekEnding = weekKey;
  }
  return summary;
}

function readKdpWeekPartialAdds_() {
  try {
    const raw = PropertiesService.getDocumentProperties().getProperty('AD_KDP_WEEK_PARTIAL_ADDS');
    if (!raw) return {};
    const obj = JSON.parse(raw);
    return obj && typeof obj === 'object' ? obj : {};
  } catch (e) {
    return {};
  }
}

function kdpPartialFingerprint_(totals, meta) {
  let units = 0;
  let kenp = 0;
  let roy = 0;
  const seen = new Set();
  Object.keys(totals || {}).forEach(k => {
    const t = totals[k];
    if (!t || seen.has(t)) return;
    seen.add(t);
    units += number_(t.units);
    kenp += number_(t.kenp);
    roy += number_(t.royaltyEbook) + number_(t.royaltyPrint) + number_(t.royaltyKenp);
  });
  return [
    meta && meta.minDate ? dateKey_(meta.minDate) : '',
    meta && meta.maxDate ? dateKey_(meta.maxDate) : '',
    units,
    kenp,
    Math.round(roy * 100)
  ].join('|');
}

function buildKdpTotalsFromRows_(found) {
  const map = {};
  const asRows = v => Array.isArray(v) ? v : [];

  const ensure = id => {
    const key = normalizeKdpId_(id);
    if (!key) return null;
    if (!map[key]) {
      map[key] = {
        units: 0,
        kenp: 0,
        royaltyEbook: 0,
        royaltyPrint: 0,
        royaltyKenp: 0,
        royaltyUsd: 0,
        aliases: new Set([key])
      };
    }
    return map[key];
  };

  const alias = (a, b) => {
    const ka = normalizeKdpId_(a);
    const kb = normalizeKdpId_(b);
    if (!ka || !kb || ka === kb) return;
    const ea = ensure(ka);
    const eb = ensure(kb);
    if (!ea || !eb || ea === eb) return;
    ea.units += eb.units;
    ea.kenp += eb.kenp;
    ea.royaltyEbook += eb.royaltyEbook;
    ea.royaltyPrint += eb.royaltyPrint;
    ea.royaltyKenp += eb.royaltyKenp;
    eb.aliases.forEach(x => ea.aliases.add(x));
    ea.aliases.add(kb);
    map[kb] = ea;
  };

  const hasOrders = asRows(found.orders).length > 0;
  const hasCombined = asRows(found.combined).length > 0;
  const royaltyCombined = {};
  const royaltyEbook = {};
  const royaltyPrint = {};
  const royaltyKenp = {};

  const addRoyalty_ = (bucket, id, row) => {
    const key = normalizeKdpId_(id);
    if (!key || !isUsdRoyaltyRow_(row)) return;
    const amt = royaltyAmount_(row);
    if (!amt) return;
    bucket[key] = (bucket[key] || 0) + amt;
  };

  asRows(found.orders).forEach(row => {
    const t = ensure(row.ASIN || row['ASIN/ISBN']);
    if (!t) return;
    t.units += number_(row['Paid Units'] != null && row['Paid Units'] !== '' ? row['Paid Units'] : row['Net Units Sold']);
  });

  asRows(found.combined).forEach(row => {
    const id = row['ASIN/ISBN'] || row.ASIN || row.ISBN;
    const t = ensure(id);
    if (!t) return;
    if (!hasOrders) t.units += number_(row['Net Units Sold']);
    addRoyalty_(royaltyCombined, id, row);
  });

  asRows(found.ebook).forEach(row => {
    const asin = row.ASIN;
    const isbn = row.ISBN || row['ASIN/ISBN'];
    if (asin && isbn) alias(asin, isbn);
    const id = asin || isbn;
    const t = ensure(id);
    if (!t) return;
    addRoyalty_(royaltyEbook, id, row);
    if (!hasOrders && !hasCombined) t.units += number_(row['Net Units Sold']);
  });

  [found.paperback, found.hardcover].forEach(rows => {
    asRows(rows).forEach(row => {
      const asin = row.ASIN;
      const isbn = row.ISBN || row['ASIN/ISBN'];
      if (asin && isbn) alias(asin, isbn);
      const id = asin || isbn;
      const t = ensure(id);
      if (!t) return;
      addRoyalty_(royaltyPrint, id, row);
      if (!hasOrders && !hasCombined) t.units += number_(row['Net Units Sold']);
    });
  });

  asRows(found.kenp).forEach(row => {
    const id = row.ASIN || row['ASIN/ISBN'];
    const t = ensure(id);
    if (!t) return;
    t.kenp += number_(
      row['Kindle Edition Normalized Page (KENP) Read'] != null && row['Kindle Edition Normalized Page (KENP) Read'] !== ''
        ? row['Kindle Edition Normalized Page (KENP) Read']
        : row.KENP
    );
    addRoyalty_(royaltyKenp, id, row);
  });

  // Finalize: eBook + Print + KENP. If format sheets empty, fall back Combined → eBook.
  const seen = new Set();
  Object.keys(map).forEach(key => {
    const obj = map[key];
    if (seen.has(obj)) return;
    seen.add(obj);
    let c = 0;
    let e = 0;
    let p = 0;
    let k = 0;
    obj.aliases.forEach(a => {
      c = Math.max(c, royaltyCombined[a] || 0);
      e = Math.max(e, royaltyEbook[a] || 0);
      p = Math.max(p, royaltyPrint[a] || 0);
      k = Math.max(k, royaltyKenp[a] || 0);
    });
    if (e + p <= 0 && c > 0) e = c;
    obj.royaltyEbook = e;
    obj.royaltyPrint = p;
    obj.royaltyKenp = k;
    obj.royaltyUsd = finalizeKdpRoyaltyUsd_(e, p, k);
  });

  return map;
}

/**
 * Estimated total USD royalties = eBook + Print + KENP (same reporting period buckets).
 */
function finalizeKdpRoyaltyUsd_(ebook, print, kenp) {
  return number_(ebook) + number_(print) + number_(kenp);
}

/** True when KENP sheet rows include a usable Royalty / earnings field. */
function kenpSheetHasRoyaltyColumn_(kenpRows) {
  const rows = Array.isArray(kenpRows) ? kenpRows : [];
  if (!rows.length) return false;
  const keys = [
    'Royalty', 'royalty', 'Royalty (USD)', 'Estimated Royalty',
    'Earnings', 'Royalty USD', 'Net Royalty'
  ];
  return rows.some(row => keys.some(k => row[k] != null && row[k] !== ''));
}

/**
 * Lifetime Royalties (USD) = eBook + Print + KENP when split columns are populated.
 * Does not replace a unit-only total with KENP-alone (0+0+k).
 */
function recomputeLifetimeRoyaltiesFromSplits_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(AD.SHEETS.INPUT);
  if (!sh || sh.getLastRow() < 2) return;
  const values = sh.getRange(2, 1, sh.getLastRow() - 1, AD.INPUT_HEADERS.length).getValues();
  values.forEach((r, i) => {
    if (!clean_(r[AD.COL.TITLE])) return;
    const e = number_(r[AD.COL.ROYALTY_EBOOK]);
    const p = number_(r[AD.COL.ROYALTY_PRINT]);
    const k = number_(r[AD.COL.ROYALTY_KENP]);
    if (!(e || p)) return; // syncEstimatedKenpRoyaltiesFromRate_ handles KENP-only updates
    const total = Math.round((e + p + k) * 100) / 100;
    if (Math.abs(total - number_(r[AD.COL.ROYALTIES])) > 0.009) {
      sh.getRange(i + 2, AD.COL.ROYALTIES + 1).setValue(total);
    }
  });
}

function royaltyAmount_(row) {
  if (!row) return 0;
  const keys = [
    'Royalty', 'royalty', 'Royalty (USD)', 'Estimated Royalty',
    'Earnings', 'Royalty USD', 'Net Royalty'
  ];
  for (let i = 0; i < keys.length; i++) {
    if (row[keys[i]] != null && row[keys[i]] !== '') return number_(row[keys[i]]);
  }
  return 0;
}

function isUsdRoyaltyRow_(row) {
  const c = normalizeKey_(row && (row.Currency || row.currency) || '');
  // Include blank currency (common on US reports) and explicit USD.
  if (!c || c === 'usd' || c === 'us dollar' || c === '$' || c === 'usa') return true;
  return false;
}

function normalizeKdpId_(v) {
  const asin = normalizeAsin_(v);
  if (asin) return asin;
  const digits = clean_(v).replace(/[^0-9Xx]/g, '').toUpperCase();
  if (digits.length === 10 || digits.length === 13) return digits;
  return '';
}

function listingIdentifierKeys_(row) {
  const keys = [];
  const id = normalizeKdpId_(row[AD.COL.IDENTIFIER]);
  if (id) keys.push(id);
  const asin = normalizeAsin_(row[AD.COL.IDENTIFIER]);
  if (asin && keys.indexOf(asin) === -1) keys.push(asin);
  return keys;
}

/** ebook | paperback | hardcover | other */
function listingFormatKind_(format) {
  const n = normalizeKey_(format);
  if (/kindle|ebook|e-book/.test(n)) return 'ebook';
  if (/paper/.test(n)) return 'paperback';
  if (/hard/.test(n)) return 'hardcover';
  return 'other';
}

/**
 * Map a KDP hit onto one Manual Entry listing without double-counting.
 * When ASIN/ISBN alias into one totals object, ebook + print listings used to
 * each get the FULL book royalties → Catalog summed 2–3× KDP Summary.
 * claims is per hit-object: { units, kenpPages, printRoy }.
 */
function splitKdpHitForListing_(hit, format, claims) {
  const kind = listingFormatKind_(format);
  const c = claims || {};
  const out = {
    units: '',
    kenp: '',
    royaltyEbook: 0,
    royaltyPrint: 0,
    royaltyKenp: 0,
    royaltyUsd: 0
  };

  if (kind === 'ebook') {
    out.royaltyEbook = number_(hit.royaltyEbook);
    out.royaltyKenp = Math.round(number_(hit.royaltyKenp) * 100) / 100;
    if (!c.kenpPages) {
      out.kenp = number_(hit.kenp);
      c.kenpPages = true;
    } else {
      out.kenp = 0;
    }
    if (!c.units) {
      out.units = number_(hit.units);
      c.units = true;
    } else {
      out.units = 0;
    }
  } else if (kind === 'paperback' || kind === 'hardcover') {
    if (!c.printRoy) {
      out.royaltyPrint = number_(hit.royaltyPrint);
      c.printRoy = true;
    }
    if (!c.units) {
      out.units = number_(hit.units);
      c.units = true;
    } else {
      out.units = 0;
    }
    out.kenp = 0;
  } else {
    // Unknown format — take whatever has not been claimed yet (legacy rows).
    if (!c.units) {
      out.units = number_(hit.units);
      c.units = true;
    } else out.units = 0;
    if (!c.kenpPages) {
      out.kenp = number_(hit.kenp);
      c.kenpPages = true;
    } else out.kenp = 0;
    out.royaltyEbook = number_(hit.royaltyEbook);
    if (!c.printRoy) {
      out.royaltyPrint = number_(hit.royaltyPrint);
      c.printRoy = true;
    }
    out.royaltyKenp = Math.round(number_(hit.royaltyKenp) * 100) / 100;
  }

  out.royaltyUsd = Math.round(
    (out.royaltyEbook + out.royaltyPrint + out.royaltyKenp) * 100
  ) / 100;
  return out;
}

/** Weak identity map for hit objects during one apply pass. */
function claimBucketForHit_(claimMap, hit) {
  if (!claimMap.has(hit)) claimMap.set(hit, {});
  return claimMap.get(hit);
}

function formatKdpImportSummary_(summary) {
  const lines = [];
  lines.push('KDP sales upload' + (summary.fileName ? ' (' + summary.fileName + ')' : '') + '.');
  lines.push('');

  if (summary.skippedDuplicatePartial) {
    lines.push('⚠ Identical short-range report already applied — skipped.');
    (summary.partialReasons || []).forEach(r => lines.push('• ' + r));
    lines.push('');
    lines.push('Download Month again when you want a full refresh.');
    return lines.join('\n');
  }

  const meta = summary.reportMeta || {};
  if (summary.appliedAsMonthContrib || meta.isMonthReport) {
    lines.push(
      'Month report (' + (summary.monthKey || '?') +
        ') — that month’s slice was stored; lifetime = sum of all uploaded months.'
    );
    if (summary.monthsStored && summary.monthsStored.length) {
      lines.push('Months on file: ' + summary.monthsStored.join(', '));
    }
    lines.push(
      'Re-upload an older month anytime to replace only that month. ' +
        'If you skip a month, upload it when you can — until then lifetime omits it.'
    );
    lines.push('');
  } else if (summary.appliedAsWeekReplace || summary.appliedAsAddition) {
    lines.push(
      'Today/Yesterday report — this week’s contribution was REPLACED with this file.'
    );
    (summary.partialReasons || []).forEach(r => lines.push('• ' + r));
    lines.push('Tip: for your usual habit, download Month — it just updates lifetime + Sales History.');
    lines.push('');
  }
  if (meta.minDate || meta.maxDate) {
    lines.push('Report dates (all): ' + (meta.minDate || '?') + ' → ' + (meta.maxDate || '?'));
  }
  if (meta.maxSalesDate) {
    lines.push('Orders/royalty through: ' + meta.maxSalesDate + ' (Sales History week uses this)');
  }
  if (meta.maxKenpDate) lines.push('KENP through: ' + meta.maxKenpDate);
  if (meta.spanDays != null) lines.push('Span: ' + meta.spanDays + ' day(s)');
  lines.push('Listings updated: ' + summary.listingsUpdated);
  lines.push('Units written (sum of applied values): ' + summary.unitsSet);
  lines.push('KENP pages written (sum): ' + summary.kuSet);
  const pm = summary.periodMetrics || {};
  if (pm.ebookRoyalties != null || summary.ebookRoyaltiesSet) {
    lines.push(
      'Royalty mix (period): eBook $' + Number(pm.ebookRoyalties != null ? pm.ebookRoyalties : summary.ebookRoyaltiesSet || 0).toFixed(2) +
        ' | Print $' + Number(pm.printRoyalties != null ? pm.printRoyalties : summary.printRoyaltiesSet || 0).toFixed(2) +
        ' | KENP $' + Number(pm.kenpRoyalties != null ? pm.kenpRoyalties : summary.kenpRoyaltiesSet || 0).toFixed(2) +
        ' | Total $' + Number(pm.totalRoyalties != null ? pm.totalRoyalties : summary.royaltiesSet || 0).toFixed(2)
    );
  } else {
    lines.push('USD royalties written (sum of listings): $' + Number(summary.royaltiesSet || 0).toFixed(2));
  }
  if (pm.ratePerKenp != null) {
    lines.push(
      'Estimated KENP Royalty Rate: $' + Number(pm.ratePerKenp).toFixed(5) +
        ' per KENP (' + Number(pm.centsPerKenp).toFixed(3) + ' cents) — not the finalized Global Fund rate'
    );
  }
  if (meta.summaryRoyaltyUsd != null) {
    lines.push('KDP Summary Royalty (USD): $' + Number(meta.summaryRoyaltyUsd).toFixed(2) + ' (may exclude estimated KENP $)');
  }
  lines.push('');
  lines.push('Royalties are USD only (CAD/GBP excluded). KENP $ come from the KENP Read royalty column when present.');
  lines.push('Enter KENPC on Manual Entry (Kindle row) for full-read and equivalent-read metrics.');
  if (summary.skippedLowerUnits || summary.skippedLowerKenp || summary.skippedLowerRoyalties) {
    lines.push(
      'Skipped lower-than-current fields — units: ' + (summary.skippedLowerUnits || 0) +
        ', KENP: ' + (summary.skippedLowerKenp || 0) +
        ', royalties: ' + (summary.skippedLowerRoyalties || 0)
    );
  }
  const unmatched = (summary.unmatchedIds || []).slice(0, 8).join(', ');
  lines.push('Unmatched report IDs: ' + summary.listingsUnmatched + (unmatched ? ' (' + unmatched + ')' : ''));
  if (summary.salesSnapshot) {
    lines.push(
      'Sales History upserted for week ending ' + (summary.snapshotWeekEnding || '') +
        ' (snapshot date ' + (summary.snapshotDate || '') + ').'
    );
  } else {
    lines.push('No Sales History snapshot (no lifetime fields written).');
  }
  lines.push('Royalty Periods row upserted for this reporting window.');
  const warns = (summary.reportMeta && summary.reportMeta.warnings) || [];
  warns.forEach(w => lines.push('Note: ' + w));
  return lines.join('\n');
}
