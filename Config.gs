const AD = {
  VERSION: '1.9.0',
  SCHEMA_VERSION: '1.9.0',
  SCRIPT_VERSION: '1.9.0',
  LAST_MIGRATION: '2026-07-26',
  MIGRATION_NOTES: 'Auto KENP $ = pages × estimated rate; Royalty Periods + KU Estimates.',
  /**
   * Working estimated USD per KENP when the KDP xlsx has pages but no KENP Royalty column.
   * Overridden by DocumentProperties once a period computes rate from real KENP $ ÷ pages.
   * Not written as a hard-coded dashboard cell — used only as the formula input.
   */
  ESTIMATED_KENP_RATE_USD: 0.0046413,
  BOOK_ID_PREFIX: 'SOL',
  // Preferred Book ID order when migrating (title match, case-insensitive).
  STABLE_BOOK_ORDER: [
    'The Kestrel Veil Incident',
    'The Helion Accord',
    'Book 3'
  ],
  TZ: 'America/Boise',
  AMAZON_FETCH_DELAY_MS: 2000,
  AMAZON_USER_AGENT: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36',
  // Used only when Amazon blocks UrlFetchApp (robot check / empty rank HTML).
  AMAZON_READER_FALLBACK_PREFIX: 'https://r.jina.ai/https://www.amazon.com/dp/',
  ACTIVE_LISTING_STATUSES: ['live', 'in review'],
  KDP_REPORTS_URL: 'https://kdp.amazon.com/en_US/reports-new',
  // Daily automated refresh (Mountain / America/Boise).
  DAILY_TRIGGER_HOUR: 5,
  // Legacy names kept so older menu/scripts still resolve.
  WEEKLY_TRIGGER_WEEKDAY: ScriptApp.WeekDay.SUNDAY,
  WEEKLY_TRIGGER_HOUR: 5,
  KDP_SHEETS: {
    COMBINED: 'Combined Sales',
    EBOOK: 'eBook Royalty',
    PAPERBACK: 'Paperback Royalty',
    HARDCOVER: 'Hardcover Royalty',
    ORDERS: 'Orders Processed',
    KENP: 'KENP Read',
    SUMMARY: 'Summary',
    DEFINITIONS: 'Report Definitions',
    PLACED: 'eBook Orders Placed'
  },
  SHEETS: {
    INPUT: 'Manual Entry',
    DASHBOARD: 'Dashboard',
    STATISTICS: 'Statistics',
    VISUAL: 'Visual Dashboard',
    VISUAL_DATA: '_Visual Chart Data',
    CATALOG: 'Catalog Summary',
    SALES: 'Sales History',
    RANKS: 'Rank History',
    MARKETING: 'Marketing History',
    YOY: 'Year over Year',
    RECONCILE: 'Reconciliation',
    EVENTS: 'Events',
    META_DAILY: 'Meta Daily',
    META_SYNC: 'Meta Sync Log',
    ROYALTY_PERIODS: 'Royalty Periods'
  },
  ROYALTY_PERIOD_HEADERS: [
    'Reporting Period Start',
    'Reporting Period End',
    'Total KENP',
    'KENP Royalties',
    'Royalty per KENP',
    'KENPC (calc)',
    'Equivalent Full Reads',
    'Estimated Full-Read Royalty',
    'eBook Royalties',
    'Print Royalties',
    'Total Royalties',
    'Royalty Mix eBook %',
    'Royalty Mix Print %',
    'Royalty Mix KENP %',
    'Royalty Estimate Status',
    'Reconcile Diff',
    'Reconcile OK',
    'Date Imported',
    'Source File'
  ],
  META_DAILY_HEADERS: [
    'Date',
    'Campaign ID',
    'Campaign Name',
    'Ad Set ID',
    'Ad Set Name',
    'Ad ID',
    'Ad Name',
    'Spend',
    'Impressions',
    'Clicks (all)',
    'Link Clicks',
    'Landing Page Views',
    'Reach',
    'CTR',
    'CPC',
    'CPM',
    'Action Types (JSON)',
    'Book ID',
    'Record Status',
    'Value Kind',
    'Source System',
    'Source File',
    'Import Batch ID',
    'Imported At',
    'Last Updated'
  ],
  META_SYNC_HEADERS: [
    'Batch ID',
    'Imported At',
    'Source File',
    'Rows Upserted',
    'Rows Skipped',
    'Date Min',
    'Date Max',
    'Spend Sum',
    'Status',
    'Message'
  ],
  EVENT_TYPES: [
    'Book release',
    'Meta campaign started',
    'Meta campaign ended',
    'Ad creative changed',
    'Price changed',
    'Free promotion',
    'Newsletter',
    'Major social post',
    'Other'
  ],
  EVENT_HEADERS: [
    'Event ID',
    'Date',
    'End Date',
    'Book ID',
    'Event Type',
    'Event Name',
    'Description'
  ],
  // Partial KDP reports shorter than this many days must not overwrite lifetime totals.
  KDP_PARTIAL_MAX_SPAN_DAYS: 350,
  // Light fills for Sales History rows (rotate; adjacent books never share a color).
  BOOK_SHADES: [
    '#E8F1F8',
    '#EAF6EE',
    '#F8F0E6',
    '#F3EAF8',
    '#E8F6F6',
    '#F8ECEC',
    '#EEF0F8',
    '#F4F6E8',
    '#F8F3EA',
    '#EAF2F8'
  ],
  salesYearSheetName: function(year){ return 'Sales ' + year; },
  isSalesYearSheetName: function(name){ return /^Sales\s+\d{4}$/.test(String(name || '')); },
  salesYearFromSheetName: function(name){
    const m = String(name || '').match(/^Sales\s+(\d{4})$/);
    return m ? Number(m[1]) : null;
  },
  COL: {
    BOOK_ID: 0,
    LISTING_ID: 1,
    TITLE: 2,
    ORIGINAL_RELEASE: 7,
    STORE: 8,
    FORMAT: 9,
    ID_TYPE: 11,
    IDENTIFIER: 12,
    STATUS: 13,
    LISTING_RELEASE: 14,
    UNITS: 16,
    KU: 17,
    ROYALTIES: 18,
    RANK: 19,
    RATING: 20,
    REVIEWS: 21,
    LAST_DATA_DATE: 22,
    ROYALTY_EBOOK: 29,
    ROYALTY_PRINT: 30,
    ROYALTY_KENP: 31,
    KENPC: 32,
    PROCESS_STATUS: 33
  },
  INPUT_HEADERS: [
    'Book ID','Listing ID','Book Title','Series','Series #','Book Stage','Word Count',
    'Original Release Date','Store / Platform','Format','Edition Version','Identifier Type',
    'Identifier / ASIN / ISBN','Listing Status','Listing Release Date','Store URL',
    'Lifetime Units','Lifetime KU Pages','Lifetime Royalties (USD)','Current Overall Rank',
    'Current Rating','Current Reviews','Last Data Date','Marketing Date',
    'Marketing Platform','Marketing Activity','Marketing Cost','Marketing Link',
    'Marketing Notes',
    'Lifetime eBook Royalties (USD)','Lifetime Print Royalties (USD)','Lifetime KENP Royalties (USD)','KENPC',
    'Process Status'
  ],
  CATALOG_HEADERS: [
    'Book ID','Book Title','Series','Series #','Stage','Word Count','Original Release Date',
    'Formats','Store Listings','Active Listings','Current Rank','Best Rank Ever',
    'Total Units','Total KU Pages','Total Royalties (USD)','Total Reviews','Average Rating',
    'Last Data Date','Last Rank Update',
    'Total eBook Royalties (USD)','Total Print Royalties (USD)','Total KENP Royalties (USD)','KENPC'
  ],
  CATALOG_COL: {
    UNITS: 12,
    KU: 13,
    ROYALTIES: 14,
    REVIEWS: 15,
    RATING: 16,
    LAST_DATA: 17,
    LAST_RANK: 18,
    ROYALTY_EBOOK: 19,
    ROYALTY_PRINT: 20,
    ROYALTY_KENP: 21,
    KENPC: 22
  },
  // Period columns = change since previous Sales History snapshot (not a KDP report week).
  SALES_HEADERS: [
    'Snapshot Date','Week Ending','Book ID','Listing ID','Book Title','Store','Format','Identifier',
    'Lifetime Units','Units Since Prev Snapshot','Lifetime KU Pages','KENP Since Prev Snapshot',
    'Lifetime Royalties','Royalties Since Prev Snapshot'
  ],
  SALES_HEADER_ALIASES: {
    'Weekly Unit Change': 'Units Since Prev Snapshot',
    'Weekly KU Change': 'KENP Since Prev Snapshot',
    'Weekly Royalty Change': 'Royalties Since Prev Snapshot'
  },
  RANK_HEADERS: [
    'Snapshot Date','Week Ending','Book ID','Listing ID','Book Title','Store','Format','Identifier',
    'Rank Type','Category','Rank','Source URL','Fetch Status'
  ],
  MARKETING_HEADERS: [
    'Entry Date','Book ID','Listing ID','Book Title','Store','Format','Marketing Platform','Activity','Cost','Link','Notes'
  ],
  STATUS: {
    UPDATED: 'Amazon rank updated',
    ROBOT: 'Amazon page returned robot check',
    MISSING_ASIN: 'ASIN missing or invalid',
    MISSING_RANK: 'Amazon rank not found',
    UNAVAILABLE: 'Amazon listing unavailable',
    PARSER: 'Amazon parser error',
    http: function(code){ return 'Amazon HTTP ' + code; }
  }
};
