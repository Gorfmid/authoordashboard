function assignInternalIds_() {
  const sh = getRequiredSheet_(AD.SHEETS.INPUT);
  if (sh.getLastRow() < 2) return;
  const values = sh.getRange(2,1,sh.getLastRow()-1,AD.INPUT_HEADERS.length).getValues();
  const titleMap = new Map();
  values.forEach(r=>{ if(clean_(r[0]) && clean_(r[2])) titleMap.set(normalizeKey_(r[2]), clean_(r[0])); });
  values.forEach((r,i)=>{
    const row=i+2, title=clean_(r[2]), store=clean_(r[8]), format=clean_(r[9]);
    if(!title) return;
    let bookId=clean_(r[0]);
    if(!bookId){ bookId=titleMap.get(normalizeKey_(title)) || generateBookId_(); sh.getRange(row,1).setValue(bookId); titleMap.set(normalizeKey_(title),bookId); }
    if(!clean_(r[1]) && store && format) sh.getRange(row,2).setValue(generateListingId_(bookId,store,format));
  });
  syncListingReleaseDates_();
}

/** If Listing Release Date is blank, copy Original Release Date (same book/format launch in most cases). */
function syncListingReleaseDates_() {
  const sh = getRequiredSheet_(AD.SHEETS.INPUT);
  if (sh.getLastRow() < 2) return;
  const values = sh.getRange(2, 1, sh.getLastRow() - 1, AD.INPUT_HEADERS.length).getValues();
  values.forEach((r, i) => {
    if (!clean_(r[AD.COL.TITLE])) return;
    if (isValidDate_(r[AD.COL.LISTING_RELEASE])) return;
    if (!isValidDate_(r[AD.COL.ORIGINAL_RELEASE])) return;
    sh.getRange(i + 2, AD.COL.LISTING_RELEASE + 1).setValue(startOfDay_(new Date(r[AD.COL.ORIGINAL_RELEASE])));
  });
}

function createStoreUrls_() {
  const sh=getRequiredSheet_(AD.SHEETS.INPUT);
  if(sh.getLastRow()<2)return;
  const rows=sh.getRange(2,1,sh.getLastRow()-1,AD.INPUT_HEADERS.length).getValues();
  rows.forEach((r,i)=>{
    const store=clean_(r[8]), type=clean_(r[11]).toUpperCase(), id=normalizeAsin_(r[12]);
    if(store.toLowerCase()==='amazon' && type==='ASIN' && id){
      sh.getRange(i+2,16).setFormula('=HYPERLINK("https://www.amazon.com/dp/'+id+'","Open Amazon Listing")');
    }
  });
}

function addStandardAmazonFormats() {
  const ui=SpreadsheetApp.getUi();
  const res=ui.prompt('Add Amazon formats','Enter the exact book title:',ui.ButtonSet.OK_CANCEL);
  if(res.getSelectedButton()!==ui.Button.OK)return;
  const title=res.getResponseText().trim(); if(!title)return;
  const sh=getRequiredSheet_(AD.SHEETS.INPUT), rows=getInputRows_();
  const base=rows.find(r=>normalizeKey_(r[2])===normalizeKey_(title));
  const existing=new Set(rows.filter(r=>normalizeKey_(r[2])===normalizeKey_(title)&&normalizeKey_(r[8])==='amazon').map(r=>normalizeKey_(r[9])));
  const newRows=[];
  ['Kindle eBook','Paperback','Hardcover'].forEach(format=>{
    if(existing.has(normalizeKey_(format)))return;
    const r=new Array(AD.INPUT_HEADERS.length).fill('');
    r[2]=title; r[3]=base?base[3]:''; r[4]=base?base[4]:''; r[5]=base?base[5]:'Published'; r[6]=base?base[6]:''; r[7]=base?base[7]:'';
    r[8]='Amazon'; r[9]=format; r[10]='First Edition'; r[11]='ASIN'; r[13]='Live';
    r[14]=base && isValidDate_(base[14]) ? base[14] : (base && isValidDate_(base[7]) ? base[7] : '');
    newRows.push(r);
  });
  if(newRows.length) sh.getRange(sh.getLastRow()+1,1,newRows.length,AD.INPUT_HEADERS.length).setValues(newRows);
  refreshEverything();
}
