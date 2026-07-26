function onOpen() {
  SpreadsheetApp.getUi().createMenu('Author Dashboard')
    .addItem('Initialize — ERASE AND REBUILD', 'initializeDashboard')
    .addSeparator()
    .addItem('Refresh Everything', 'refreshEverything')
    .addItem('Record Current Snapshot', 'recordCurrentSnapshot')
    .addItem('Update Amazon Rankings Now', 'updateAmazonRanks')
    .addItem('Upload KDP Sales Report', 'uploadKdpSalesReport')
    .addItem('Open KDP Reports Page', 'openKdpReportsPage')
    .addItem('Upload Meta Insights CSV', 'uploadMetaInsightsCsv')
    .addItem('Process Marketing Entries', 'processMarketingEntries')
    .addItem('Add Event', 'addEventDialog')
    .addItem('Add Standard Amazon Formats', 'addStandardAmazonFormats')
    .addSeparator()
    .addItem('Migrate Sales History (Phase 0)', 'migrateSalesHistoryPhase0')
    .addItem('Migrate Stable Book IDs (SOL-001…)', 'migrateStableBookIdsMenu')
    .addItem('Run Reconciliation', 'runSalesReconciliationMenu')
    .addItem('Run Phase 0 Tests', 'runPhase0Tests')
    .addSeparator()
    .addItem('Install Weekly Saturday Night Update', 'installWeeklyTrigger')
    .addItem('Remove Weekly Update', 'removeWeeklyTrigger')
    .addItem('Install Weekly Saturday Night Rank Update', 'installWeeklyRankTrigger')
    .addItem('Remove Weekly Rank Update', 'removeWeeklyRankTriggerUi')
    .addSeparator()
    .addItem('Unlock Automatic Sheets', 'unlockAutomaticSheets')
    .addItem('Relock Automatic Sheets', 'lockAutomaticSheets')
    .addToUi();
}

// Compatibility alias for earlier versions.
function initializeAuthorDashboard() {
  initializeDashboard();
}

function initializeDashboard() {
  const ui = SpreadsheetApp.getUi();
  const answer = ui.alert(
    'Erase and rebuild workbook?',
    'This permanently deletes every existing sheet and creates the Author Dashboard from scratch.',
    ui.ButtonSet.YES_NO
  );
  if (answer !== ui.Button.YES) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  try { ss.setSpreadsheetTimeZone(AD.TZ); } catch (e) {}
  removeWeeklyTrigger();
  removeWeeklyRankTrigger();
  removeAllProtections_();

  const temp = ss.insertSheet('__TEMP__' + Date.now());
  ss.getSheets().forEach(s => {
    if (s.getSheetId() !== temp.getSheetId()) ss.deleteSheet(s);
  });

  temp.setName(AD.SHEETS.INPUT);
  const dashboard = ss.insertSheet(AD.SHEETS.DASHBOARD);
  ensureVisualDashboard_();
  const catalog = ss.insertSheet(AD.SHEETS.CATALOG);
  const sales = ss.insertSheet(AD.SHEETS.SALES);
  const ranks = ss.insertSheet(AD.SHEETS.RANKS);
  const marketing = ss.insertSheet(AD.SHEETS.MARKETING);

  buildInputSheet_(temp);
  buildDashboardSheet_(dashboard);
  buildCatalogSheet_(catalog);
  buildHistorySheet_(sales, AD.SALES_HEADERS);
  buildHistorySheet_(ranks, AD.RANK_HEADERS);
  buildHistorySheet_(marketing, AD.MARKETING_HEADERS);
  ensureYearOverYearSheet_();
  ensureEventsSheet_();
  ensureMetaSheets_();
  ensureRoyaltyPeriodsSheet_();
  rebuildReconciliationSheet_();
  seedFirstBook_(temp);
  orderSheets_();
  orderReportSheets_();
  refreshEverything();
  lockAutomaticSheets();
  ss.setActiveSheet(temp);
  ss.moveActiveSheet(1);
  ui.alert('Author Dashboard created. Manual Entry is the only sheet you need to edit.');
}

function refreshEverything() {
  assignInternalIds_();
  createStoreUrls_();
  applyManualEntryColumnVisibility_();
  ensureRankHistorySchema_();
  ensureSalesHistorySchema_();
  ensureCatalogSchema_();
  ensureEventsSheet_();
  ensureMetaSheets_();
  cleanupManualRankSnapshots_();
  repairMissingJuly25OverallRanks_();
  consolidateSalesHistoryByWeekEnding_();
  recomputeSalesPeriodChangesFromLifetime_();
  ensureInputKuRoyaltySchema_();
  ensureRoyaltyPeriodsSheet_();
  syncEstimatedKenpRoyaltiesFromRate_();
  recomputeLifetimeRoyaltiesFromSplits_();
  syncLatestRoyaltyPeriodKenpFromManual_();
  syncMetaCampaignMarketingRows_();
  syncAutoEvents_();
  refreshSalesReports_();
  rebuildCatalogSummary_();
  refreshDashboard_();
  rebuildReconciliationSheet_();
  processMarketingEntries_(false);
  orderReportSheets_();
  hideDiagnosticSheets_();
  lockAutomaticSheets();
}

function onEdit(e) {
  if (!e || !e.range) return;
  const sh = e.range.getSheet();
  if (sh.getName() !== AD.SHEETS.INPUT || e.range.getRow() < 2) return;
  if ([1, 2, 16, AD.COL.PROCESS_STATUS + 1].includes(e.range.getColumn())) return;
  try {
    assignInternalIds_();
    createStoreUrls_();
    const col = e.range.getColumn();
    if ([AD.COL.ROYALTY_EBOOK + 1, AD.COL.ROYALTY_PRINT + 1, AD.COL.ROYALTY_KENP + 1].includes(col)) {
      recomputeLifetimeRoyaltiesFromSplits_();
    }
  } catch (err) {
    console.error(err);
  }
}
