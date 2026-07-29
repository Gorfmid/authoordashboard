/** Menu: install both daily sales + rank jobs at once. */
function installDailyJobs() {
  installDailyTrigger_(false);
  installDailyRankTrigger_(false);
  uiAlert_(
    'Daily jobs installed for ' + AD.DAILY_TRIGGER_HOUR + ':00 AM ' + AD.TZ +
      '.\n\n• Sales / dashboard refresh\n• Amazon rank update\n\n' +
      'Apps Script runs sometime during that hour (not exactly on the minute).\n' +
      'Check Statistics → Last Run / Daily Jobs after tomorrow morning.'
  );
}

/** Menu: remove both daily jobs. */
function removeDailyJobs() {
  removeDailyTrigger();
  removeDailyRankTrigger();
  uiAlert_('Daily sales and rank jobs removed.');
}

function installDailyTrigger() {
  installDailyTrigger_(true);
}

function installDailyTrigger_(showAlert) {
  removeDailyTrigger();
  ScriptApp.newTrigger('scheduledDailyUpdate')
    .timeBased()
    .everyDays(1)
    .atHour(AD.DAILY_TRIGGER_HOUR)
    .inTimezone(AD.TZ)
    .create();
  if (showAlert !== false) {
    uiAlert_(
      'Daily update installed for ' + AD.DAILY_TRIGGER_HOUR + ':00 AM ' + AD.TZ +
        '.\n\nSales History upserts the current week-ending Saturday row.'
    );
  }
}

function removeDailyTrigger() {
  const handlers = [
    'scheduledDailyUpdate',
    'scheduledWeeklyUpdate' // legacy weekly handler
  ];
  ScriptApp.getProjectTriggers().forEach(t => {
    if (handlers.indexOf(t.getHandlerFunction()) !== -1) ScriptApp.deleteTrigger(t);
  });
}

/**
 * Create missing daily triggers (no UI). Safe to call from onOpen / refresh.
 * Does not recreate existing triggers.
 */
function ensureDailyJobsInstalled_() {
  const handlers = ScriptApp.getProjectTriggers().map(t => t.getHandlerFunction());
  const hasSales = handlers.indexOf('scheduledDailyUpdate') !== -1;
  const hasRank = handlers.indexOf('scheduledDailyRankUpdate') !== -1;
  if (!hasSales) installDailyTrigger_(false);
  if (!hasRank) installDailyRankTrigger_(false);
  return {
    sales: hasSales || true,
    ranks: hasRank || true,
    createdSales: !hasSales,
    createdRanks: !hasRank
  };
}

function getDailyJobsStatus_() {
  const handlers = ScriptApp.getProjectTriggers().map(t => t.getHandlerFunction());
  const sales = handlers.indexOf('scheduledDailyUpdate') !== -1;
  const ranks = handlers.indexOf('scheduledDailyRankUpdate') !== -1;
  if (sales && ranks) {
    return 'Installed — ' + AD.DAILY_TRIGGER_HOUR + ':00 AM ' + AD.TZ + ' (sales + ranks)';
  }
  if (sales) return 'Partial — sales only (' + AD.DAILY_TRIGGER_HOUR + ':00 AM ' + AD.TZ + ')';
  if (ranks) return 'Partial — ranks only (' + AD.DAILY_TRIGGER_HOUR + ':00 AM ' + AD.TZ + ')';
  return 'Not installed — Author Dashboard → Install Daily Jobs (5 AM)';
}

function scheduledDailyUpdate() {
  // Time-driven triggers have no UI session — never call SpreadsheetApp.getUi().
  try {
    assignInternalIds_();
    createStoreUrls_();
    recordCurrentSnapshotSilent_();
    try {
      syncMetaInsightsFromApi_({ quiet: true, refreshDashboard: false, lockSheets: false });
    } catch (metaErr) {
      console.error('Meta API sync (daily job): ' + (metaErr && metaErr.message ? metaErr.message : metaErr));
    }
    syncMetaCampaignMarketingRows_();
    syncAutoEvents_();
    processMarketingEntries_(false);
    refreshSalesReports_();
    rebuildCatalogSummary_();
    setLastDashboardRun_('daily');
    refreshDashboard_();
    rebuildReconciliationSheet_();
    hideDiagnosticSheets_();
    lockAutomaticSheets();
    console.log('scheduledDailyUpdate OK at ' + new Date().toISOString());
  } catch (err) {
    console.error('scheduledDailyUpdate failed: ' + (err && err.message ? err.message : err));
    throw err;
  }
}

/** Upsert Sales History for the current week ending (no daily duplicate rows). */
function recordCurrentSnapshotSilent_() {
  ensureRankHistorySchema_();
  ensureSalesHistorySchema_();
  const rows = getInputRows_();
  const today = getSpreadsheetToday_();
  const week = getWeekEndingDate_(today);
  recordSalesSnapshot_(rows, today, week, today, { upsertByWeek: true });
  consolidateSalesHistoryByWeekEnding_();
  recomputeSalesPeriodChangesFromLifetime_();
}

function installDailyRankTrigger() {
  installDailyRankTrigger_(true);
}

function installDailyRankTrigger_(showAlert) {
  removeDailyRankTrigger();
  ScriptApp.newTrigger('scheduledDailyRankUpdate')
    .timeBased()
    .everyDays(1)
    .atHour(AD.DAILY_TRIGGER_HOUR)
    .inTimezone(AD.TZ)
    .create();
  if (showAlert !== false) {
    uiAlert_(
      'Daily Amazon rank update installed for ' + AD.DAILY_TRIGGER_HOUR + ':00 AM ' + AD.TZ + '.'
    );
  }
}

function removeDailyRankTrigger() {
  const handlers = [
    'scheduledDailyRankUpdate',
    'scheduledWeeklyRankUpdate' // legacy
  ];
  ScriptApp.getProjectTriggers().forEach(t => {
    if (handlers.indexOf(t.getHandlerFunction()) !== -1) ScriptApp.deleteTrigger(t);
  });
}

function removeDailyRankTriggerUi() {
  removeDailyRankTrigger();
  uiAlert_('Daily Amazon rank update removed.');
}

function scheduledDailyRankUpdate() {
  // Time-driven triggers have no UI session — never call SpreadsheetApp.getUi().
  try {
    updateAmazonRanks_(false, getSpreadsheetToday_());
    console.log('scheduledDailyRankUpdate OK at ' + new Date().toISOString());
  } catch (err) {
    console.error('scheduledDailyRankUpdate failed: ' + (err && err.message ? err.message : err));
    throw err;
  }
}

// --- Legacy menu entry points (redirect to daily) ---
function installWeeklyTrigger() { installDailyTrigger(); }
function removeWeeklyTrigger() {
  removeDailyTrigger();
  uiAlert_('Daily / weekly sales update trigger removed.');
}
function installWeeklyRankTrigger() { installDailyRankTrigger(); }
function removeWeeklyRankTrigger() { removeDailyRankTrigger(); }
function removeWeeklyRankTriggerUi() { removeDailyRankTriggerUi(); }
function scheduledWeeklyUpdate() { scheduledDailyUpdate(); }
function scheduledWeeklyRankUpdate() { scheduledDailyRankUpdate(); }

function getEffectiveTimezone_() {
  try {
    return SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone() || AD.TZ;
  } catch (e) {
    return AD.TZ;
  }
}
