/** Menu: install both daily sales + rank jobs at once. */
function installDailyJobs() {
  installDailyTrigger_(false);
  installDailyRankTrigger_(false);
  SpreadsheetApp.getUi().alert(
    'Daily jobs installed for ' + AD.DAILY_TRIGGER_HOUR + ':00 AM ' + AD.TZ +
      '.\n\n• Sales / dashboard refresh\n• Amazon rank update'
  );
}

/** Menu: remove both daily jobs. */
function removeDailyJobs() {
  removeDailyTrigger();
  removeDailyRankTrigger();
  SpreadsheetApp.getUi().alert('Daily sales and rank jobs removed.');
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
    SpreadsheetApp.getUi().alert(
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

function scheduledDailyUpdate() {
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
  refreshDashboard_();
  rebuildReconciliationSheet_();
  hideDiagnosticSheets_();
  lockAutomaticSheets();
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
    SpreadsheetApp.getUi().alert(
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
  SpreadsheetApp.getUi().alert('Daily Amazon rank update removed.');
}

function scheduledDailyRankUpdate() {
  updateAmazonRanks_(false, getSpreadsheetToday_());
}

// --- Legacy menu entry points (redirect to daily) ---
function installWeeklyTrigger() { installDailyTrigger(); }
function removeWeeklyTrigger() {
  removeDailyTrigger();
  try {
    SpreadsheetApp.getUi().alert('Daily / weekly sales update trigger removed.');
  } catch (e) {}
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
