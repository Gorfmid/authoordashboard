function buildInputSheet_(sheet) {
  sheet.clear();
  sheet.getRange(1,1,1,AD.INPUT_HEADERS.length).setValues([AD.INPUT_HEADERS]);
  styleHeader_(sheet.getRange(1,1,1,AD.INPUT_HEADERS.length));
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(3);
  const widths = [115,175,260,190,75,145,105,125,135,145,130,130,170,120,125,230,115,125,130,135,100,105,125,120,140,270,105,220,280,220];
  widths.forEach((w,i)=>sheet.setColumnWidth(i+1,w));
  applyInputValidation_(sheet);
  applyInputFormats_(sheet);
  addFilter_(sheet, AD.INPUT_HEADERS.length);
  applyManualEntryColumnVisibility_(sheet);
}

/**
 * Manual Entry shows only hand-entered fields.
 * Auto-managed columns stay in the sheet (for scripts) but are greyed and hidden.
 * Current ranks / URLs / dates appear on Catalog, Rank History, and Dashboard.
 */
function applyManualEntryColumnVisibility_(sheet) {
  const sh = sheet || getRequiredSheet_(AD.SHEETS.INPUT);
  const autoRanges = ['A:B', 'O:P', 'T:W', 'AD:AD'];
  autoRanges.forEach(a1 => {
    sh.getRange(a1).setBackground('#eeeeee').setFontColor('#666666');
  });

  // 1-based column index, count
  const hideGroups = [
    [1, 2],  // Book ID, Listing ID
    [15, 2], // Listing Release Date, Store URL
    [20, 4], // Current Overall Rank, Rating, Reviews, Last Data Date
    [30, 1]  // Process Status
  ];
  hideGroups.forEach(pair => {
    try { sh.hideColumns(pair[0], pair[1]); } catch (e) {}
  });
}

function hideInternalIdColumns_(sheet) {
  applyManualEntryColumnVisibility_(sheet);
}

function buildDashboardSheet_(sheet) {
  sheet.clear();
  sheet.getRange('A1:H1').merge().setValue('Author Portfolio Dashboard')
    .setFontSize(22).setFontWeight('bold').setHorizontalAlignment('center')
    .setBackground('#1f4e78').setFontColor('#ffffff');
  sheet.setRowHeight(1,46);
  sheet.getRange('A3:B3').setValues([['Portfolio Metric','Current Value']]);
  styleHeader_(sheet.getRange('A3:B3'));
  const labels = [
    'Total Books','Published Books','Books in Progress','Total Store Listings','Live Store Listings',
    'Total Published Words','Lifetime Unit Sales','Lifetime KU Pages','Lifetime Royalties','Total Reviews',
    'Average Rating','Best Rank Ever','Current Best Rank','Latest Rank Update','Top-Ranked Book',
    'Rank Trend (lower is better)'
  ];
  sheet.getRange(4,1,labels.length,2).setValues(labels.map(x=>[x,'']));
  sheet.getRange('D3:H3').merge().setValue('Catalog Performance').setFontWeight('bold')
    .setHorizontalAlignment('center').setBackground('#1f4e78').setFontColor('#ffffff');
  sheet.getRange('D4:H4').setValues([['Book','Stage','Units','Royalties','Best Rank']]);
  styleHeader_(sheet.getRange('D4:H4'));
  [225,145,30,270,135,105,125,115,110,200].forEach((w,i)=>sheet.setColumnWidth(i+1,w));
}

function buildCatalogSheet_(sheet) {
  sheet.clear();
  sheet.getRange(1,1,1,AD.CATALOG_HEADERS.length).setValues([AD.CATALOG_HEADERS]);
  styleHeader_(sheet.getRange(1,1,1,AD.CATALOG_HEADERS.length));
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(2);
  [115,260,190,75,145,105,125,100,110,110,125,120,115,125,130,110,105,130,130].forEach((w,i)=>sheet.setColumnWidth(i+1,w));
  addFilter_(sheet, AD.CATALOG_HEADERS.length);
}

function buildHistorySheet_(sheet, headers) {
  sheet.clear();
  sheet.getRange(1,1,1,headers.length).setValues([headers]);
  styleHeader_(sheet.getRange(1,1,1,headers.length));
  sheet.setFrozenRows(1);
  addFilter_(sheet, headers.length);
  sheet.autoResizeColumns(1, headers.length);
}

function seedFirstBook_(sheet) {
  const rows = ['Kindle eBook','Paperback','Hardcover'].map(format => {
    const r = new Array(AD.INPUT_HEADERS.length).fill('');
    r[2] = 'The Kestrel Veil Incident';
    r[3] = 'The Solmare Cycle';
    r[4] = 1;
    r[5] = 'Published';
    r[6] = 91114;
    r[8] = 'Amazon';
    r[9] = format;
    r[10] = 'First Edition';
    r[11] = 'ASIN';
    r[13] = 'Live';
    return r;
  });
  sheet.getRange(2,1,rows.length,AD.INPUT_HEADERS.length).setValues(rows);
}

function orderSheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  [AD.SHEETS.INPUT,AD.SHEETS.DASHBOARD,AD.SHEETS.CATALOG,AD.SHEETS.SALES,AD.SHEETS.RANKS,AD.SHEETS.MARKETING].forEach((name,i)=>{
    const sh=ss.getSheetByName(name); if(sh){ss.setActiveSheet(sh);ss.moveActiveSheet(i+1);}
  });
}
