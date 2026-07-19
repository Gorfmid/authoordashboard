/**
 * One-click KDP sales upload.
 * Parses the .xlsx in a dialog (no clutter sheets), then updates Manual Entry,
 * records a sales snapshot, and refreshes Catalog + Dashboard.
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
    .setWidth(460)
    .setHeight(420);
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
      hardcover: payload.sheets[AD.KDP_SHEETS.HARDCOVER] || null
    };

    const hasAny = Object.keys(found).some(k => found[k] && found[k].length);
    if (!hasAny) throw new Error('Report contained no usable KDP sales rows.');

    assignInternalIds_();
    const summary = applyKdpTotalsToManualEntry_(buildKdpTotalsFromRows_(found));

    // Magically refresh derived sheets without leaving import clutter.
    ensureRankHistorySchema_();
    const rows = getInputRows_();
    const today = getSpreadsheetToday_();
    const week = getWeekEndingDate_(today);
    recordSalesSnapshot_(rows, today, week);
    rebuildCatalogSummary_();
    refreshDashboard_();
    lockAutomaticSheets();

    summary.fileName = payload.fileName || '';
    summary.salesSnapshot = true;
    console.log('KDP upload summary: ' + JSON.stringify(summary));
    return formatKdpImportSummary_(summary);
  } finally {
    lock.releaseLock();
  }
}

function applyKdpTotalsToManualEntry_(totals) {
  const sh = getRequiredSheet_(AD.SHEETS.INPUT);
  const summary = {
    identifiersInReport: Object.keys(totals).length,
    listingsUpdated: 0,
    listingsUnmatched: 0,
    unitsSet: 0,
    kuSet: 0,
    royaltiesSet: 0,
    unmatchedIds: []
  };

  if (sh.getLastRow() < 2) return summary;

  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, AD.INPUT_HEADERS.length).getValues();
  const matchedKeys = new Set();

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
    sh.getRange(rowNum, AD.COL.UNITS + 1).setValue(hit.units);
    sh.getRange(rowNum, AD.COL.KU + 1).setValue(hit.kenp);
    sh.getRange(rowNum, AD.COL.ROYALTIES + 1).setValue(hit.royaltyUsd);
    sh.getRange(rowNum, AD.COL.LAST_DATA_DATE + 1).setValue(getSpreadsheetToday_());
    sh.getRange(rowNum, AD.COL.PROCESS_STATUS + 1).setValue('KDP sales upload applied');
    summary.listingsUpdated++;
    summary.unitsSet += hit.units;
    summary.kuSet += hit.kenp;
    summary.royaltiesSet += hit.royaltyUsd;
  });

  Object.keys(totals).forEach(id => {
    if (!matchedKeys.has(id)) {
      // Avoid counting alias duplicates pointing at the same object.
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
    if (!map[key]) map[key] = { units: 0, kenp: 0, royaltyUsd: 0, aliases: new Set([key]) };
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
    ea.royaltyUsd += eb.royaltyUsd;
    eb.aliases.forEach(x => ea.aliases.add(x));
    ea.aliases.add(kb);
    map[kb] = ea;
  };

  const hasOrders = asRows(found.orders).length > 0;
  const hasCombined = asRows(found.combined).length > 0;

  asRows(found.orders).forEach(row => {
    const t = ensure(row.ASIN || row['ASIN/ISBN']);
    if (!t) return;
    t.units += number_(row['Paid Units'] != null && row['Paid Units'] !== '' ? row['Paid Units'] : row['Net Units Sold']);
  });

  asRows(found.combined).forEach(row => {
    const t = ensure(row['ASIN/ISBN'] || row.ASIN || row.ISBN);
    if (!t) return;
    if (!hasOrders) t.units += number_(row['Net Units Sold']);
    if (normalizeKey_(row.Currency) === 'usd' || !clean_(row.Currency)) {
      t.royaltyUsd += number_(row.Royalty);
    }
  });

  [found.ebook, found.paperback, found.hardcover].forEach(rows => {
    asRows(rows).forEach(row => {
      const asin = row.ASIN;
      const isbn = row.ISBN || row['ASIN/ISBN'];
      if (asin && isbn) alias(asin, isbn);
      const t = ensure(asin || isbn);
      if (!t) return;
      if (!hasCombined && (normalizeKey_(row.Currency) === 'usd' || !clean_(row.Currency))) {
        t.royaltyUsd += number_(row.Royalty);
      }
      if (!hasOrders && !hasCombined) t.units += number_(row['Net Units Sold']);
    });
  });

  asRows(found.kenp).forEach(row => {
    const t = ensure(row.ASIN || row['ASIN/ISBN']);
    if (!t) return;
    t.kenp += number_(
      row['Kindle Edition Normalized Page (KENP) Read'] != null && row['Kindle Edition Normalized Page (KENP) Read'] !== ''
        ? row['Kindle Edition Normalized Page (KENP) Read']
        : row.KENP
    );
  });

  return map;
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
  const unmatched = (summary.unmatchedIds || []).slice(0, 8).join(', ');
  return [
    'KDP sales upload complete' + (summary.fileName ? ' (' + summary.fileName + ')' : '') + '.',
    '',
    'Listings updated: ' + summary.listingsUpdated,
    'Units: ' + summary.unitsSet,
    'KENP: ' + summary.kuSet,
    'USD royalties: $' + Number(summary.royaltiesSet || 0).toFixed(2),
    'Unmatched report IDs: ' + summary.listingsUnmatched + (unmatched ? ' (' + unmatched + ')' : ''),
    summary.salesSnapshot ? 'Sales History snapshot recorded.' : ''
  ].filter(Boolean).join('\n');
}
