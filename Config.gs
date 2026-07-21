const AD = {
  VERSION: '1.3.0',
  TZ: 'America/Boise',
  AMAZON_FETCH_DELAY_MS: 2000,
  AMAZON_USER_AGENT: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36',
  // Used only when Amazon blocks UrlFetchApp (robot check / empty rank HTML).
  AMAZON_READER_FALLBACK_PREFIX: 'https://r.jina.ai/https://www.amazon.com/dp/',
  ACTIVE_LISTING_STATUSES: ['live', 'in review'],
  KDP_REPORTS_URL: 'https://kdp.amazon.com/en_US/reports-new',
  // Saturday night close = Sunday 00:00 Mountain (America/Boise).
  WEEKLY_TRIGGER_WEEKDAY: ScriptApp.WeekDay.SUNDAY,
  WEEKLY_TRIGGER_HOUR: 0,
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
    CATALOG: 'Catalog Summary',
    SALES: 'Sales History',
    RANKS: 'Rank History',
    MARKETING: 'Marketing History',
    YOY: 'Year over Year'
  },
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
    PROCESS_STATUS: 29
  },
  INPUT_HEADERS: [
    'Book ID','Listing ID','Book Title','Series','Series #','Book Stage','Word Count',
    'Original Release Date','Store / Platform','Format','Edition Version','Identifier Type',
    'Identifier / ASIN / ISBN','Listing Status','Listing Release Date','Store URL',
    'Lifetime Units','Lifetime KU Pages','Lifetime Royalties','Current Overall Rank',
    'Current Rating','Current Reviews','Last Data Date','Marketing Date',
    'Marketing Platform','Marketing Activity','Marketing Cost','Marketing Link',
    'Marketing Notes','Process Status'
  ],
  CATALOG_HEADERS: [
    'Book ID','Book Title','Series','Series #','Stage','Word Count','Original Release Date',
    'Formats','Store Listings','Active Listings','Current Best Rank','Best Rank Ever',
    'Total Units','Total KU Pages','Total Royalties','Total Reviews','Average Rating',
    'Last Data Date','Last Rank Update'
  ],
  SALES_HEADERS: [
    'Snapshot Date','Week Ending','Book ID','Listing ID','Book Title','Store','Format','Identifier',
    'Lifetime Units','Weekly Unit Change','Lifetime KU Pages','Weekly KU Change',
    'Lifetime Royalties','Weekly Royalty Change'
  ],
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
