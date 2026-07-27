/**
 * Meta Marketing insights.
 * Preferred: Refresh Everything pulls from the Meta Graph API when credentials
 * are stored in Document Properties (not sheet cells). CSV upload remains a fallback.
 *
 * Do NOT label Amazon orders as Meta conversions. Spend/clicks are ad metrics only.
 */

var AD_META_PROP_TOKEN = 'AD_META_ACCESS_TOKEN';
var AD_META_PROP_ACCOUNT = 'AD_META_AD_ACCOUNT_ID';
var AD_META_PROP_VERSION = 'AD_META_API_VERSION';

function uploadMetaInsightsCsv() {
  const html = HtmlService.createHtmlOutputFromFile('MetaUpload')
    .setWidth(520)
    .setHeight(480);
  SpreadsheetApp.getUi().showModalDialog(html, 'Upload Meta insights CSV');
}

/**
 * One-time setup: store Meta token + ad account in Document Properties, then sync.
 * Menu: Author Dashboard → Connect Meta Ads…
 * Same values as local .env META_ACCESS_TOKEN / META_AD_ACCOUNT_ID.
 */
function configureMetaApiCredentials() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getDocumentProperties();
  const existing = getMetaApiConfig_();
  if (existing.token && existing.accountId) {
    const again = ui.alert(
      'Meta already connected',
      'Account: ' + existing.accountId + '\n\nReplace credentials?',
      ui.ButtonSet.YES_NO
    );
    if (again !== ui.Button.YES) {
      // Re-sync with existing credentials.
      try {
        const summary = syncMetaInsightsFromApi_({ quiet: false, refreshDashboard: true, lockSheets: true });
        ui.alert(formatMetaImportSummary_(summary));
      } catch (e) {
        ui.alert('Meta sync failed', e && e.message ? e.message : String(e), ui.ButtonSet.OK);
      }
      return;
    }
  }

  const tokenRes = ui.prompt(
    'Meta access token',
    'Paste your long-lived Marketing API token (same as META_ACCESS_TOKEN in .env).\nStored in Document Properties — not written to sheet cells.',
    ui.ButtonSet.OK_CANCEL
  );
  if (tokenRes.getSelectedButton() !== ui.Button.OK) return;
  const token = String(tokenRes.getResponseText() || '').trim();
  if (!token) {
    ui.alert('Token required.');
    return;
  }

  const acctRes = ui.prompt(
    'Meta ad account ID',
    'Ad account id (same as META_AD_ACCOUNT_ID, e.g. act_1234567890):',
    ui.ButtonSet.OK_CANCEL
  );
  if (acctRes.getSelectedButton() !== ui.Button.OK) return;
  let account = String(acctRes.getResponseText() || '').trim();
  if (!account) {
    ui.alert('Ad account ID required.');
    return;
  }
  if (!/^act_/i.test(account)) account = 'act_' + account;

  props.setProperty(AD_META_PROP_TOKEN, token);
  props.setProperty(AD_META_PROP_ACCOUNT, account);
  if (!props.getProperty(AD_META_PROP_VERSION)) {
    props.setProperty(AD_META_PROP_VERSION, 'v21.0');
  }

  try {
    const summary = syncMetaInsightsFromApi_({ quiet: false, refreshDashboard: true, lockSheets: true });
    ui.alert(
      'Meta connected and synced.\n\n' + formatMetaImportSummary_(summary) +
        '\n\nRefresh Everything will keep Meta Daily updated.'
    );
  } catch (e) {
    ui.alert(
      'Credentials saved, but sync failed',
      (e && e.message ? e.message : String(e)) +
        '\n\nCheck the token/account, then use Connect Meta Ads… again (or Refresh Everything).',
      ui.ButtonSet.OK
    );
  }
}

function getMetaApiConfig_() {
  const props = PropertiesService.getDocumentProperties();
  return {
    token: String(props.getProperty(AD_META_PROP_TOKEN) || '').trim(),
    accountId: String(props.getProperty(AD_META_PROP_ACCOUNT) || '').trim(),
    apiVersion: String(props.getProperty(AD_META_PROP_VERSION) || 'v21.0').trim() || 'v21.0'
  };
}

function hasMetaApiCredentials_() {
  const cfg = getMetaApiConfig_();
  return !!(cfg.token && cfg.accountId);
}

/**
 * Pull Meta insights via Graph API and upsert Meta Daily.
 * @param {{quiet?: boolean}=} opts quiet=true skips when credentials missing (for Refresh Everything).
 * @returns {Object} summary or { skipped: true, reason }
 */
function syncMetaInsightsFromApi_(opts) {
  opts = opts || {};
  const cfg = getMetaApiConfig_();
  if (!cfg.token || !cfg.accountId) {
    if (opts.quiet) return { skipped: true, reason: 'no credentials' };
    throw new Error(
      'Meta API credentials not set. Run configureMetaApiCredentials() from Apps Script once.'
    );
  }

  const end = getSpreadsheetToday_();
  const start = new Date(end.getTime());
  start.setDate(start.getDate() - 30);
  const dateStart = Utilities.formatDate(start, AD.TZ, 'yyyy-MM-dd');
  const dateEnd = Utilities.formatDate(end, AD.TZ, 'yyyy-MM-dd');

  const insightRows = fetchMetaInsightsFromApi_(cfg, dateStart, dateEnd);
  const csvRows = metaInsightsToImportRows_(insightRows);
  return applyMetaInsightRows_(csvRows, {
    fileName: 'Meta API ' + dateStart + '→' + dateEnd,
    refreshDashboard: opts.refreshDashboard !== false,
    lockSheets: opts.lockSheets !== false
  });
}

/** Menu / script entry: sync Meta now (shows alert). */
function syncMetaInsightsFromApi() {
  const summary = syncMetaInsightsFromApi_({ quiet: false, refreshDashboard: true, lockSheets: true });
  SpreadsheetApp.getUi().alert(formatMetaImportSummary_(summary));
}

/**
 * payload = { fileName, rows: [ {Date, Campaign ID, ...}, ... ] }
 */
function processMetaInsightsUpload(payload) {
  if (!payload || !payload.rows || !payload.rows.length) {
    throw new Error('No Meta insight rows received.');
  }
  const summary = applyMetaInsightRows_(payload.rows, {
    fileName: payload.fileName || '',
    refreshDashboard: true,
    lockSheets: true
  });
  return formatMetaImportSummary_(summary);
}

function applyMetaInsightRows_(rows, opts) {
  opts = opts || {};
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) throw new Error('Workbook is busy. Try again in a moment.');

  try {
    ensureMetaSheets_();
    const batchId = 'META-' + Utilities.formatDate(new Date(), AD.TZ, 'yyyyMMdd-HHmmss');
    const importedAt = new Date();
    const summary = upsertMetaDailyRows_(rows, {
      batchId: batchId,
      fileName: opts.fileName || '',
      importedAt: importedAt
    });
    appendMetaSyncLog_(summary);
    setMetaSyncHealth_(summary);
    syncMetaCampaignMarketingRows_();
    syncAutoEvents_();
    if (opts.refreshDashboard !== false) refreshDashboard_();
    if (opts.lockSheets !== false) lockAutomaticSheets();
    return summary;
  } finally {
    lock.releaseLock();
  }
}

function fetchMetaInsightsFromApi_(cfg, dateStart, dateEnd) {
  const account = /^act_/i.test(cfg.accountId) ? cfg.accountId : 'act_' + cfg.accountId;
  const fields = [
    'campaign_id', 'campaign_name',
    'adset_id', 'adset_name',
    'ad_id', 'ad_name',
    'spend', 'impressions', 'clicks', 'reach', 'ctr', 'cpc', 'cpm',
    'actions', 'date_start', 'date_stop'
  ].join(',');

  const params = {
    access_token: cfg.token,
    level: 'ad',
    time_increment: '1',
    limit: '100',
    fields: fields,
    time_range: JSON.stringify({ since: dateStart, until: dateEnd })
  };
  const qs = Object.keys(params)
    .map(k => encodeURIComponent(k) + '=' + encodeURIComponent(params[k]))
    .join('&');
  let nextUrl = 'https://graph.facebook.com/' + cfg.apiVersion + '/' + account + '/insights?' + qs;

  const rows = [];
  let pages = 0;
  while (nextUrl && pages < 40) {
    pages++;
    const data = metaApiGetJson_(nextUrl);
    const chunk = data.data || [];
    if (Array.isArray(chunk)) {
      chunk.forEach(r => rows.push(r));
    }
    nextUrl = (data.paging && data.paging.next) || null;
    if (nextUrl) Utilities.sleep(250);
  }
  return rows;
}

function metaApiGetJson_(url) {
  let delayMs = 2000;
  let lastErr = null;
  for (let attempt = 1; attempt <= 5; attempt++) {
    const resp = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      followRedirects: true
    });
    const code = resp.getResponseCode();
    let data = {};
    try {
      data = JSON.parse(resp.getContentText() || '{}');
    } catch (e) {
      data = {};
    }
    if (code === 200 && !data.error) return data;

    const err = data.error || {};
    const errCode = err.code;
    const msg = err.message || ('HTTP ' + code);

    if (errCode === 190 || errCode === 102 || code === 401) {
      throw new Error('Meta token error (' + errCode + '): ' + msg + '. Run configureMetaApiCredentials().');
    }
    if (code === 429 || code >= 500 || errCode === 4 || errCode === 17 || errCode === 32 || errCode === 613) {
      lastErr = new Error(msg);
      Utilities.sleep(delayMs);
      delayMs = Math.min(delayMs * 2, 30000);
      continue;
    }
    throw new Error('Meta API error: ' + msg);
  }
  throw new Error('Meta API failed after retries: ' + (lastErr && lastErr.message ? lastErr.message : 'unknown'));
}

function metaActionValue_(actions, actionType) {
  if (!Array.isArray(actions)) return 0;
  for (let i = 0; i < actions.length; i++) {
    const item = actions[i];
    if (item && item.action_type === actionType) return number_(item.value);
  }
  return 0;
}

function metaInsightsToImportRows_(insights) {
  return (insights || []).map(row => {
    const actions = row.actions || [];
    let linkClicks = metaActionValue_(actions, 'link_click');
    if (linkClicks <= 0) linkClicks = metaActionValue_(actions, 'outbound_click');
    const lpv = metaActionValue_(actions, 'landing_page_view');
    return {
      Date: row.date_start || '',
      'Campaign ID': asTextId_(row.campaign_id),
      'Campaign Name': String(row.campaign_name || ''),
      'Ad Set ID': asTextId_(row.adset_id),
      'Ad Set Name': String(row.adset_name || ''),
      'Ad ID': asTextId_(row.ad_id),
      'Ad Name': String(row.ad_name || ''),
      Spend: row.spend != null ? row.spend : 0,
      Impressions: row.impressions != null ? row.impressions : 0,
      'Clicks (all)': row.clicks != null ? row.clicks : 0,
      'Link Clicks': linkClicks,
      'Landing Page Views': lpv,
      Reach: row.reach != null ? row.reach : 0,
      CTR: row.ctr != null ? row.ctr : '',
      CPC: row.cpc != null ? row.cpc : '',
      CPM: row.cpm != null ? row.cpm : '',
      'Action Types (JSON)': JSON.stringify(actions),
      'Book ID': ''
    };
  });
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
