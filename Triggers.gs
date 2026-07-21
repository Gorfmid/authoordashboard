function installWeeklyTrigger() {
  removeWeeklyTrigger();
  ScriptApp.newTrigger('scheduledWeeklyUpdate')
    .timeBased()
    .onWeekDay(AD.WEEKLY_TRIGGER_WEEKDAY)
    .atHour(AD.WEEKLY_TRIGGER_HOUR)
    .inTimezone(AD.TZ)
    .create();
  SpreadsheetApp.getUi().alert(
    'Weekly sales snapshot installed for Saturday night midnight Mountain time ' +
      '(Sunday 12:00 AM ' + AD.TZ + '). Week ending is Saturday.'
  );
}

function removeWeeklyTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'scheduledWeeklyUpdate') ScriptApp.deleteTrigger(t);
  });
}

function scheduledWeeklyUpdate() {
  assignInternalIds_();
  createStoreUrls_();
  recordCurrentSnapshotSilent_();
  processMarketingEntries_(false);
  refreshSalesReports_();
  rebuildCatalogSummary_();
  refreshDashboard_();
  lockAutomaticSheets();
}

function recordCurrentSnapshotSilent_() {
  ensureRankHistorySchema_();
  const rows = getInputRows_();
  const saturday = getMostRecentWeekEndingSaturday_();
  recordRankSnapshot_(rows, saturday, saturday);
  recordSalesSnapshot_(rows, saturday, saturday);
}

function installWeeklyRankTrigger() {
  removeWeeklyRankTrigger();
  ScriptApp.newTrigger('scheduledWeeklyRankUpdate')
    .timeBased()
    .onWeekDay(AD.WEEKLY_TRIGGER_WEEKDAY)
    .atHour(AD.WEEKLY_TRIGGER_HOUR)
    .inTimezone(AD.TZ)
    .create();
  SpreadsheetApp.getUi().alert(
    'Weekly Amazon rank update installed for Saturday night midnight Mountain time ' +
      '(Sunday 12:00 AM ' + AD.TZ + '). Week ending is Saturday.'
  );
}

function removeWeeklyRankTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'scheduledWeeklyRankUpdate') ScriptApp.deleteTrigger(t);
  });
}

function removeWeeklyRankTriggerUi() {
  removeWeeklyRankTrigger();
  SpreadsheetApp.getUi().alert('Weekly Amazon rank update removed.');
}

function scheduledWeeklyRankUpdate() {
  updateAmazonRanks_(false, getMostRecentWeekEndingSaturday_());
}

function getEffectiveTimezone_() {
  try {
    return SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone() || AD.TZ;
  } catch (e) {
    return AD.TZ;
  }
}
