function getRequiredSheet_(name){const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);if(!sh)throw new Error('Required sheet "'+name+'" not found. Run initializeDashboard().');return sh;}
function getInputRows_(){const sh=getRequiredSheet_(AD.SHEETS.INPUT);return sh.getLastRow()<2?[]:sh.getRange(2,1,sh.getLastRow()-1,AD.INPUT_HEADERS.length).getValues();}
function styleHeader_(r){r.setFontWeight('bold').setBackground('#1f4e78').setFontColor('#ffffff').setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);}
function addFilter_(sh,cols){if(sh.getFilter())sh.getFilter().remove();sh.getRange(1,1,Math.max(sh.getMaxRows(),2),cols).createFilter();}
function clearDataRows_(sh){if(sh.getLastRow()>=2)sh.getRange(2,1,sh.getLastRow()-1,sh.getMaxColumns()).clearContent();}
function appendRows_(sh,rows){if(rows.length)sh.getRange(sh.getLastRow()+1,1,rows.length,rows[0].length).setValues(rows);}
/** Next permanent book id: SOL-001, SOL-002, … */
function generateBookId_(){
  return formatSolBookId_(getMaxSolBookNumber_()+1);
}
/**
 * Listing id: {BookId}-{STORE4}-{FORM4}-{SUFFIX}
 * Optional suffix keeps listing identity stable across book-id migrations.
 */
function generateListingId_(bookId,store,format,suffix){
  const suf=clean_(suffix).replace(/[^A-Za-z0-9]/g,'').toUpperCase().substring(0,8)
    || Utilities.getUuid().replace(/-/g,'').substring(0,4).toUpperCase();
  return [bookId,code_(store).substring(0,4),code_(format).substring(0,4),suf].join('-');
}
function code_(v){return String(v||'').replace(/[^A-Za-z0-9]/g,'').toUpperCase();}
function normalizeAsin_(v){const t=clean_(v).toUpperCase(),m=t.match(/\/(?:DP|GP\/PRODUCT)\/([A-Z0-9]{10})/i)||t.match(/\b([A-Z0-9]{10})\b/);return m?m[1]:'';}
function clean_(v){return String(v===null||v===undefined?'':v).trim();}
function normalizeKey_(v){return clean_(v).toLowerCase();}
function number_(v){if(typeof v==='number')return isFinite(v)?v:0;const n=Number(clean_(v).replace(/[^0-9.-]/g,''));return isFinite(n)?n:0;}
function average_(a){return a.length?a.reduce((s,x)=>s+x,0)/a.length:0;}
function isValidDate_(v){return !isNaN(new Date(v).getTime());}
function startOfDay_(d){const x=new Date(d);x.setHours(0,0,0,0);return x;}
/** Upcoming or current Saturday (week ends Saturday night). */
function getWeekEndingDate_(d){const x=startOfDay_(d);x.setDate(x.getDate()+((6-x.getDay()+7)%7));return x;}
/** Most recently completed week-ending Saturday (for Saturday-night / Sunday-midnight runs). */
function getMostRecentWeekEndingSaturday_(d){
  const x=startOfDay_(d||getSpreadsheetToday_());
  const daysBack=(x.getDay()+1)%7; // Sat=0, Sun=1, Mon=2, ...
  x.setDate(x.getDate()-daysBack);
  return x;
}
function dateKey_(d){let tz=AD.TZ;try{tz=SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone()||AD.TZ;}catch(e){}return Utilities.formatDate(d,tz,'yyyy-MM-dd');}
function getExistingSnapshotKeys_(sh,dateCol,listingCol,typeCol){const set=new Set();if(sh.getLastRow()<2)return set;sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).getValues().forEach(r=>{if(!isValidDate_(r[dateCol-1])||!clean_(r[listingCol-1]))return;const p=[dateKey_(new Date(r[dateCol-1])),clean_(r[listingCol-1])];if(typeCol)p.push(clean_(r[typeCol-1]));set.add(p.join('|'));});return set;}
/** Unmerge any merged ranges that overlap a block, then clear contents. Avoids merge/clear errors. */
function clearBlockUnmerged_(sheet,startRow,startCol,numRows,numCols){
  const endRow=startRow+numRows-1,endCol=startCol+numCols-1;
  try{
    const merges=sheet.getRange(1,1,Math.max(sheet.getMaxRows(),endRow),Math.max(sheet.getMaxColumns(),endCol)).getMergedRanges();
    merges.forEach(m=>{
      const r1=m.getRow(),c1=m.getColumn(),r2=r1+m.getNumRows()-1,c2=c1+m.getNumColumns()-1;
      if(r1<=endRow&&r2>=startRow&&c1<=endCol&&c2>=startCol){try{m.breakApart();}catch(e){}}
    });
  }catch(e){}
  sheet.getRange(startRow,startCol,numRows,numCols).clearContent().clearFormat();
}
function mergeRowSafe_(sheet,row,startCol,numCols){
  // Never merge across a freeze boundary — unfreeze first.
  try{sheet.setFrozenRows(0);}catch(e){}
  try{sheet.setFrozenColumns(0);}catch(e){}
  const range=sheet.getRange(row,startCol,1,numCols);
  try{
    const merges=range.getMergedRanges();
    merges.forEach(m=>{try{m.breakApart();}catch(e){}});
  }catch(e){}
  return range.merge();
}

/** Clear freeze + merges so rebuilds never hit "can't merge frozen and non-frozen columns". */
function prepareSheetForRebuild_(sh){
  try{sh.setFrozenRows(0);}catch(e){}
  try{sh.setFrozenColumns(0);}catch(e){}
  try{
    sh.getMergedRanges().forEach(m=>{try{m.breakApart();}catch(e){}});
  }catch(e){
    try{
      const mr=Math.min(sh.getMaxRows(),200);
      const mc=Math.min(sh.getMaxColumns(),40);
      sh.getRange(1,1,mr,mc).getMergedRanges().forEach(m=>{try{m.breakApart();}catch(e2){}});
    }catch(e3){}
  }
}

/**
 * Banner/legend row without merge (safe with frozen columns).
 * Styles a block and writes text only in the first cell.
 */
/** Sheets kept for diagnostics / history but not shown in the tab bar. */
function getDiagnosticSheetNames_(){
  return [
    AD.SHEETS.RECONCILE,
    AD.SHEETS.META_SYNC,
    AD.SHEETS.ROYALTY_PERIODS,
    AD.SHEETS.VISUAL_DATA
  ];
}

/** Hide diagnostic sheets. Call after create, rebuild, or tab reordering (reorder unhides). */
function hideDiagnosticSheets_(){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  getDiagnosticSheetNames_().forEach(name=>{
    const sh=ss.getSheetByName(name);
    if(!sh)return;
    try{sh.hideSheet();}catch(e){}
  });
}

function setBannerRow_(sh,row,numCols,text,opts){
  opts=opts||{};
  const cols=Math.max(1,numCols||1);
  const range=sh.getRange(row,1,1,cols);
  try{range.getMergedRanges().forEach(m=>{try{m.breakApart();}catch(e){}});}catch(e){}
  range.clearContent().clearFormat()
    .setBackground(opts.background||'#FFF8E7')
    .setFontColor(opts.fontColor||'#333333')
    .setFontWeight(opts.bold?'bold':'normal')
    .setFontSize(opts.fontSize||10)
    .setWrap(true);
  sh.getRange(row,1).setValue(text||'').setHorizontalAlignment('left').setVerticalAlignment('middle');
  if(opts.rowHeight)sh.setRowHeight(row,opts.rowHeight);
}
