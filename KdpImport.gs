/**
 * One-click KDP sales upload.
 * Parses the .xlsx in a dialog (no clutter sheets), then updates Manual Entry,
 * records a sales snapshot, and refreshes Catalog + Dashboard.
 *
 * Phase 0: never let a partial date-range report overwrite lifetime totals;
 * use report date range for Week Ending when available.
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

    // Only snapshot when lifetime writes occurred (partial blocks skip lifetime).
    if (summary.listingsUpdated > 0 && !summary.blockedPartial) {
      const rows = getInputRows_();
      const today = getSpreadsheetToday_();
      // Week Ending / snapshot anchor from ORDER/ROYALTY dates only — not KENP.
      // KENP can post on Sunday (e.g. 7/26) and wrongly push Week Ending to next Saturday (8/1).
      const salesAnchor = meta.maxSalesDate || today;
      const snapDate = salesAnchor;
      const week = getWeekEndingDate_(salesAnchor);
      recordSalesSnapshot_(rows, snapDate, week, salesAnchor, { upsertByWeek: true });
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
  const nameLooksPartial = /partial|last\s*\d+\s*day|custom.?range|\d+\s*days?/i.test(name);
  if (nameLooksPartial) {
    reasons.push('Filename suggests a date-range / partial export.');
  }

  const lowerCount = countReportLowerThanLifetime_(totals || {});
  const mostlyLower =
    lowerCount.compared > 0 &&
    lowerCount.lower >= Math.max(1, Math.ceil(lowerCount.compared * 0.5));
  if (mostlyLower) {
    reasons.push(
      'Report totals are lower than current Manual Entry lifetime for ' +
        lowerCount.lower + ' of ' + lowerCount.compared +
        ' matched listing metric(s). Partial reports must not replace lifetime totals.'
    );
  }

  if (
    spanDays != null &&
    spanDays > 0 &&
    spanDays < AD.KDP_PARTIAL_MAX_SPAN_DAYS &&
    mostlyLower
  ) {
    reasons.push(
      'Report date span is only ' + spanDays +
        ' day(s) (' + dateKey_(minDate) + ' → ' + dateKey_(maxDate) +
        ') and totals are below current lifetime — treating as partial.'
    );
  } else if (spanDays != null && spanDays > 0 && spanDays < AD.KDP_PARTIAL_MAX_SPAN_DAYS) {
    warnings.push(
      'Report date span is ' + spanDays +
        ' day(s). If this was a custom range (not All time), do not use it for lifetime totals.'
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
    isPartial: reasons.length > 0,
    reasons: reasons,
    warnings: warnings,
    lowerCount: lowerCount
  };
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
 * - Partial reports: block all lifetime writes
 * - Non-partial: never decrease an existing lifetime value (blank stays blank unless report has a value)
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
    skippedLowerUnits: 0,
    skippedLowerKenp: 0,
    skippedLowerRoyalties: 0,
    fieldsWritten: 0
  };

  if (meta && meta.isPartial) {
    summary.blockedPartial = true;
    summary.partialReasons = meta.reasons || [];
    return summary;
  }

  if (sh.getLastRow() < 2) return summary;

  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, AD.INPUT_HEADERS.length).getValues();
  const matchedKeys = new Set();
  const today = getSpreadsheetToday_();

  rows.forEach((r, i) => {
    if (normalizeKey_(r[AD.COL.STORE]) !== 'amazon') return;
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

    // Units
    if (hit.units != null && hit.units !== '') {
      const next = number_(hit.units);
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

    // KENP
    if (hit.kenp != null && hit.kenp !== '') {
      const next = number_(hit.kenp);
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

    // Royalty splits (eBook / Print / KENP) + total = sum.
    // If xlsx lacks KENP $, hit.royaltyKenp may already be pages × estimated rate.
    const nextEbook = number_(hit.royaltyEbook);
    const nextPrint = number_(hit.royaltyPrint);
    const nextKenpR = Math.round(number_(hit.royaltyKenp) * 100) / 100;
    const nextTotal = Math.round((nextEbook + nextPrint + nextKenpR) * 100) / 100;
    if ((hit.royaltyUsd != null && hit.royaltyUsd !== '') || nextEbook || nextPrint || nextKenpR) {
      const hasCur = curR !== '' && curR !== null && curR !== undefined;
      if (hasCur && nextTotal < number_(curR) - 0.009 && meta && meta.kenpRoyaltiesInReport) {
        summary.skippedLowerRoyalties++;
        notes.push('royalties not lowered');
      } else if (!hasCur || nextTotal >= number_(curR) - 0.009 || !(meta && meta.kenpRoyaltiesInReport)) {
        if (!hasCur || Math.abs(nextTotal - number_(curR)) > 0.009 ||
            Math.abs(nextEbook - number_(r[AD.COL.ROYALTY_EBOOK])) > 0.009 ||
            Math.abs(nextPrint - number_(r[AD.COL.ROYALTY_PRINT])) > 0.009 ||
            Math.abs(nextKenpR - number_(r[AD.COL.ROYALTY_KENP])) > 0.009) {
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

  return summary;
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

/** Lifetime Royalties (USD) = eBook + Print + KENP for every listing row. */
function recomputeLifetimeRoyaltiesFromSplits_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(AD.SHEETS.INPUT);
  if (!sh || sh.getLastRow() < 2) return;
  const values = sh.getRange(2, 1, sh.getLastRow() - 1, AD.INPUT_HEADERS.length).getValues();
  values.forEach((r, i) => {
    if (!clean_(r[AD.COL.TITLE])) return;
    const e = number_(r[AD.COL.ROYALTY_EBOOK]);
    const p = number_(r[AD.COL.ROYALTY_PRINT]);
    const k = number_(r[AD.COL.ROYALTY_KENP]);
    if (!(e || p || k || number_(r[AD.COL.ROYALTIES]))) return;
    const total = e + p + k;
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

function formatKdpImportSummary_(summary) {
  const lines = [];
  lines.push('KDP sales upload' + (summary.fileName ? ' (' + summary.fileName + ')' : '') + '.');
  lines.push('');

  if (summary.blockedPartial) {
    lines.push('⚠ BLOCKED — report looks PARTIAL. Lifetime totals were NOT overwritten.');
    (summary.partialReasons || []).forEach(r => lines.push('• ' + r));
    lines.push('');
    lines.push('Download a KDP Dashboard report covering the full life of each book (All time / lifetime), then upload again.');
    lines.push('Partial-period reports must not replace Lifetime Units / KENP / Royalties.');
    return lines.join('\n');
  }

  const meta = summary.reportMeta || {};
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
