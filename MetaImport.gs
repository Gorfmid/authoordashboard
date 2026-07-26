/**
 * Meta Marketing insights — import only (no secrets in the workbook).
 * Pull data locally with meta/sync_meta_insights.py + .env, then upload the CSV here.
 *
 * Do NOT label Amazon orders as Meta conversions. Spend/clicks are ad metrics only.
 */

function uploadMetaInsightsCsv() {
  const html = HtmlService.createHtmlOutputFromFile('MetaUpload')
    .setWidth(520)
    .setHeight(480);
  SpreadsheetApp.getUi().showModalDialog(html, 'Upload Meta insights CSV');
}

/**
 * payload = { fileName, rows: [ {Date, Campaign ID, ...}, ... ] }
 */
function processMetaInsightsUpload(payload) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) throw new Error('Workbook is busy. Try again in a moment.');

  try {
    if (!payload || !payload.rows || !payload.rows.length) {
      throw new Error('No Meta insight rows received.');
    }
    ensureMetaSheets_();
    const batchId = 'META-' + Utilities.formatDate(new Date(), AD.TZ, 'yyyyMMdd-HHmmss');
    const importedAt = new Date();
    const summary = upsertMetaDailyRows_(payload.rows, {
      batchId: batchId,
      fileName: payload.fileName || '',
      importedAt: importedAt
    });
    appendMetaSyncLog_(summary);
    setMetaSyncHealth_(summary);
    syncMetaCampaignMarketingRows_();
    syncAutoEvents_();
    refreshDashboard_();
    lockAutomaticSheets();
    return formatMetaImportSummary_(summary);
  } finally {
    lock.releaseLock();
  }
}

function ensureMetaSheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let daily = ss.getSheetByName(AD.SHEETS.META_DAILY);
  if (!daily) {
    daily = ss.insertSheet(AD.SHEETS.META_DAILY);
    buildMetaHistorySheet_(daily, AD.META_DAILY_HEADERS);
  } else {
    ensureHeaderRow_(daily, AD.META_DAILY_HEADERS);
  }
  let sync = ss.getSheetByName(AD.SHEETS.META_SYNC);
  if (!sync) {
    sync = ss.insertSheet(AD.SHEETS.META_SYNC);
    buildMetaHistorySheet_(sync, AD.META_SYNC_HEADERS);
  } else {
    ensureHeaderRow_(sync, AD.META_SYNC_HEADERS);
  }
  daily.getRange(1, 1).setNote(
    'Daily Meta ad metrics from the Marketing API. NOT the same as Marketing History ' +
      '(manual one-off notes). Do not treat clicks/spend as Amazon conversions.'
  );
  hideDiagnosticSheets_();
  return { daily: daily, sync: sync };
}

function buildMetaHistorySheet_(sheet, headers) {
  sheet.clear();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeader_(sheet.getRange(1, 1, 1, headers.length));
  sheet.setFrozenRows(1);
  addFilter_(sheet, headers.length);
  if (headers[0] === 'Date' || headers[0] === 'Batch ID') {
    sheet.getRange('A:A').setNumberFormat('@');
  }
}

function ensureHeaderRow_(sh, headers) {
  const cols = headers.length;
  const cur = sh.getRange(1, 1, 1, cols).getValues()[0].map(clean_);
  let changed = false;
  headers.forEach((h, i) => {
    if (cur[i] !== h) {
      sh.getRange(1, i + 1).setValue(h);
      changed = true;
    }
  });
  if (changed) styleHeader_(sh.getRange(1, 1, 1, cols));
}

function upsertMetaDailyRows_(rows, meta) {
  const sh = getRequiredSheet_(AD.SHEETS.META_DAILY);
  const existing = getMetaDailyKeyMap_();
  const summary = {
    batchId: meta.batchId,
    fileName: meta.fileName,
    importedAt: meta.importedAt,
    rowsUpserted: 0,
    rowsSkipped: 0,
    dateMin: null,
    dateMax: null,
    spendSum: 0,
    status: 'OK',
    message: ''
  };

  const toAppend = [];
  rows.forEach(raw => {
    const norm = normalizeMetaInsightRow_(raw, meta);
    if (!norm) {
      summary.rowsSkipped++;
      return;
    }
    const key = metaDailyKey_(norm.dateKey, norm.campaignId, norm.adSetId, norm.adId);
    summary.spendSum += norm.spend;
    if (!summary.dateMin || norm.dateKey < summary.dateMin) summary.dateMin = norm.dateKey;
    if (!summary.dateMax || norm.dateKey > summary.dateMax) summary.dateMax = norm.dateKey;

    if (existing[key] != null) {
      const rowNum = existing[key];
      const vals = norm.values.slice();
      const prevImported = sh.getRange(rowNum, 24).getValue(); // Imported At
      if (prevImported) vals[23] = prevImported;
      vals[24] = meta.importedAt; // Last Updated
      sh.getRange(rowNum, 1, 1, AD.META_DAILY_HEADERS.length).setValues([vals]);
      summary.rowsUpserted++;
    } else {
      toAppend.push(norm.values);
      summary.rowsUpserted++;
    }
  });

  if (toAppend.length) appendRows_(sh, toAppend);
  sh.getRange('A:A').setNumberFormat('m/d/yyyy');
  sh.getRange('H:H').setNumberFormat('$#,##0.00');
  sh.getRange('I:L').setNumberFormat('#,##0');
  sh.getRange('N:P').setNumberFormat('0.0000');
  sh.getRange('X:Y').setNumberFormat('m/d/yyyy hh:mm:ss');
  // Force ID columns as text display
  ['B', 'D', 'F'].forEach(col => {
    try { sh.getRange(col + '2:' + col).setNumberFormat('@'); } catch (e) {}
  });
  return summary;
}

function getMetaDailyKeyMap_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(AD.SHEETS.META_DAILY);
  const map = {};
  if (!sh || sh.getLastRow() < 2) return map;
  const values = sh.getRange(2, 1, sh.getLastRow() - 1, 7).getValues();
  values.forEach((r, i) => {
    if (!isValidDate_(r[0])) return;
    const key = metaDailyKey_(dateKey_(new Date(r[0])), clean_(r[1]), clean_(r[3]), clean_(r[5]));
    map[key] = i + 2;
  });
  return map;
}

function metaDailyKey_(dateKey, campaignId, adSetId, adId) {
  return [dateKey, campaignId, adSetId, adId].join('|');
}

function normalizeMetaInsightRow_(raw, meta) {
  if (!raw || typeof raw !== 'object') return null;
  const dateRaw = raw.Date || raw.date || raw.date_start || raw['Date Start'];
  const d = parseLooseDate_(dateRaw);
  if (!d) return null;

  const campaignId = asTextId_(raw['Campaign ID'] || raw.campaign_id || '');
  const adSetId = asTextId_(raw['Ad Set ID'] || raw.adset_id || '');
  const adId = asTextId_(raw['Ad ID'] || raw.ad_id || '');
  if (!campaignId && !adSetId && !adId) return null;

  const spend = number_(raw.Spend != null ? raw.Spend : raw.spend);
  const impressions = number_(raw.Impressions != null ? raw.Impressions : raw.impressions);
  const clicksAll = number_(raw['Clicks (all)'] != null ? raw['Clicks (all)'] : (raw.clicks != null ? raw.clicks : raw.Clicks));
  const linkClicks = number_(raw['Link Clicks'] != null ? raw['Link Clicks'] : raw.link_clicks);
  const lpv = number_(raw['Landing Page Views'] != null ? raw['Landing Page Views'] : raw.landing_page_views);
  const reach = number_(raw.Reach != null ? raw.Reach : raw.reach);
  const ctr = number_(raw.CTR != null ? raw.CTR : raw.ctr);
  const cpc = number_(raw.CPC != null ? raw.CPC : raw.cpc);
  const cpm = number_(raw.CPM != null ? raw.CPM : raw.cpm);
  const actionsJson = clean_(raw['Action Types (JSON)'] || raw.actions_json || '');
  const bookId = clean_(raw['Book ID'] || raw.book_id || '');

  const values = [
    d,
    campaignId,
    clean_(raw['Campaign Name'] || raw.campaign_name || ''),
    adSetId,
    clean_(raw['Ad Set Name'] || raw.adset_name || ''),
    adId,
    clean_(raw['Ad Name'] || raw.ad_name || ''),
    spend,
    impressions,
    clicksAll,
    linkClicks,
    lpv,
    reach,
    ctr,
    cpc,
    cpm,
    actionsJson,
    bookId,
    'active',
    'imported',
    'Meta Marketing API',
    meta.fileName || '',
    meta.batchId,
    meta.importedAt,
    meta.importedAt
  ];

  return {
    dateKey: dateKey_(d),
    campaignId: campaignId,
    adSetId: adSetId,
    adId: adId,
    spend: spend,
    values: values
  };
}

/** Keep Meta IDs as text (avoid scientific notation). */
function asTextId_(v) {
  if (v === null || v === undefined || v === '') return '';
  if (typeof v === 'number') return String(Math.round(v));
  return clean_(v).replace(/^['\t]+/, '').replace(/\.0+$/, '');
}

function appendMetaSyncLog_(summary) {
  const sh = getRequiredSheet_(AD.SHEETS.META_SYNC);
  appendRows_(sh, [[
    summary.batchId,
    summary.importedAt,
    summary.fileName,
    summary.rowsUpserted,
    summary.rowsSkipped,
    summary.dateMin || '',
    summary.dateMax || '',
    summary.spendSum,
    summary.status,
    summary.message || ''
  ]]);
  sh.getRange('B:B').setNumberFormat('m/d/yyyy hh:mm:ss');
  sh.getRange('H:H').setNumberFormat('$#,##0.00');
}

function setMetaSyncHealth_(summary) {
  const props = PropertiesService.getDocumentProperties();
  props.setProperty('AD_META_LAST_SYNC', JSON.stringify({
    at: summary.importedAt ? summary.importedAt.toISOString() : new Date().toISOString(),
    batchId: summary.batchId,
    status: summary.status,
    rows: summary.rowsUpserted,
    spendSum: summary.spendSum,
    dateMax: summary.dateMax || '',
    message: summary.message || ''
  }));
}

function getMetaSyncHealth_() {
  try {
    const raw = PropertiesService.getDocumentProperties().getProperty('AD_META_LAST_SYNC');
    return raw ? JSON.parse(raw) : null;
  } catch (e) {
    return null;
  }
}

function formatMetaImportSummary_(summary) {
  return [
    'Meta insights import complete.',
    '',
    'Batch: ' + summary.batchId,
    'Rows upserted: ' + summary.rowsUpserted,
    'Rows skipped: ' + summary.rowsSkipped,
    'Date range: ' + (summary.dateMin || '?') + ' → ' + (summary.dateMax || '?'),
    'Spend in file (sum of daily rows): $' + Number(summary.spendSum || 0).toFixed(2),
    '',
    'These are ad metrics only — not Amazon conversions.',
    'One Marketing History row per campaign was upserted; Events auto-synced campaign starts.'
  ].join('\n');
}

/**
 * Roll Meta Daily into one Marketing History row per campaign (upsert on refresh).
 * Marker in Notes: AUTO_META|{campaignId}|...
 */
function syncMetaCampaignMarketingRows_() {
  ensureMetaSheets_();
  const campaigns = summarizeMetaCampaigns_();
  if (!campaigns.length) return 0;

  const sh = getRequiredSheet_(AD.SHEETS.MARKETING);
  const existing = {}; // campaignId -> row number
  if (sh.getLastRow() >= 2) {
    const notes = sh.getRange(2, 11, sh.getLastRow() - 1, 1).getValues();
    notes.forEach((r, i) => {
      const m = String(r[0] || '').match(/^AUTO_META\|([^|]+)/);
      if (m) existing[m[1]] = i + 2;
    });
  }

  let n = 0;
  campaigns.forEach(c => {
    const start = parseLooseDate_(c.dateMin) || getSpreadsheetToday_();
    const notes =
      'AUTO_META|' + c.campaignId + '|' + (c.dateMin || '') + '→' + (c.dateMax || '') +
      '|daily rows in Meta Daily; not Amazon conversions';
    const row = [
      start,
      c.bookId || '',
      'META-' + c.campaignId,
      c.bookTitle || '',
      'Meta',
      '',
      'Meta',
      'Campaign: ' + (c.campaignName || c.campaignId),
      c.spend,
      '',
      notes
    ];
    if (existing[c.campaignId]) {
      sh.getRange(existing[c.campaignId], 1, 1, AD.MARKETING_HEADERS.length).setValues([row]);
    } else {
      appendRows_(sh, [row]);
    }
    n++;
  });
  sh.getRange('A:A').setNumberFormat('m/d/yyyy');
  sh.getRange('I:I').setNumberFormat('$#,##0.00');
  return n;
}

/** Aggregate Meta Daily by campaign id. */
function summarizeMetaCampaigns_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(AD.SHEETS.META_DAILY);
  if (!sh || sh.getLastRow() < 2) return [];
  const values = sh.getRange(2, 1, sh.getLastRow() - 1, AD.META_DAILY_HEADERS.length).getValues();
  const byId = new Map();

  values.forEach(r => {
    const campaignId = asTextId_(r[1]);
    if (!campaignId || !isValidDate_(r[0])) return;
    const dk = dateKey_(new Date(r[0]));
    if (!byId.has(campaignId)) {
      byId.set(campaignId, {
        campaignId: campaignId,
        campaignName: clean_(r[2]),
        bookId: clean_(r[17]),
        bookTitle: '',
        spend: 0,
        dateMin: dk,
        dateMax: dk
      });
    }
    const c = byId.get(campaignId);
    c.spend += number_(r[7]);
    if (clean_(r[2])) c.campaignName = clean_(r[2]);
    if (clean_(r[17])) c.bookId = clean_(r[17]);
    if (dk < c.dateMin) c.dateMin = dk;
    if (dk > c.dateMax) c.dateMax = dk;
  });

  // Resolve book titles from Manual Entry
  const books = new Map();
  getInputRows_().forEach(r => {
    const id = clean_(r[AD.COL.BOOK_ID]);
    if (id && !books.has(id)) books.set(id, clean_(r[AD.COL.TITLE]));
  });
  const list = [...byId.values()];
  list.forEach(c => {
    if (c.bookId && books.has(c.bookId)) c.bookTitle = books.get(c.bookId);
  });
  return list.sort((a, b) => a.campaignName.localeCompare(b.campaignName));
}
