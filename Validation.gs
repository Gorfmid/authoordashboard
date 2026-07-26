function applyInputValidation_(sh){
  const list=(items)=>SpreadsheetApp.newDataValidation().requireValueInList(items,true).setAllowInvalid(true).build();
  sh.getRange('F2:F').setDataValidation(list(['Idea','Outline','Rough Draft','Story Revision','Writing Pass','Final Candidate','Final Lock','Published','Paused','Cancelled']));
  sh.getRange('I2:I').setDataValidation(list(['Amazon','Apple Books','Barnes & Noble','Kobo','Google Play Books','IngramSpark','Audible','Spotify','Direct Sales','Other']));
  sh.getRange('J2:J').setDataValidation(list(['Kindle eBook','EPUB eBook','Paperback','Hardcover','Audiobook','Large Print','Box Set','Special Edition','Other']));
  sh.getRange('L2:L').setDataValidation(list(['ASIN','ISBN-10','ISBN-13','Apple ID','Kobo ID','Google Books ID','SKU','URL','Other']));
  sh.getRange('N2:N').setDataValidation(list(['Planned','In Setup','In Review','Live','Paused','Unpublished']));
}
function applyInputFormats_(sh){
  sh.getRange('E2:E').setNumberFormat('0'); sh.getRange('G2:G').setNumberFormat('#,##0'); sh.getRange('H2:H').setNumberFormat('m/d/yyyy'); sh.getRange('O2:O').setNumberFormat('m/d/yyyy');
  sh.getRange('Q2:R').setNumberFormat('#,##0'); sh.getRange('S2:S').setNumberFormat('$#,##0.00'); sh.getRange('T2:T').setNumberFormat('#,##0'); sh.getRange('U2:U').setNumberFormat('0.0'); sh.getRange('V2:V').setNumberFormat('#,##0'); sh.getRange('W2:X').setNumberFormat('m/d/yyyy'); sh.getRange('AA2:AA').setNumberFormat('$#,##0.00');
  sh.getRange('AD2:AF').setNumberFormat('$#,##0.00');
  sh.getRange('AG2:AG').setNumberFormat('#,##0');
  sh.getRange('A1').setNote('Permanent Book ID (SOL-001, SOL-002, …). Do not use title as the key. Run Migrate Stable Book IDs once if you still see BK-… values.');
  sh.getRange('B1').setNote('Permanent Listing ID (BookId-STORE-FORMAT-suffix). ASIN/ISBN stay in Identifier. Store and Format are separate columns.');
  sh.getRange('M1').setNote('Enter the ASIN or ISBN here. Amazon rank updates require ASIN.');
  sh.getRange('Q1').setNote('LIFETIME cumulative units for this listing (All-time). Never paste a single-week or custom-range KDP total here.');
  sh.getRange('R1').setNote('LIFETIME cumulative KENP pages (All-time). Period/week changes are calculated on Sales History snapshots.');
  sh.getRange('S1').setNote('LIFETIME estimated USD royalties = eBook + Print + KENP $. CAD/GBP excluded. Updated by KDP upload.');
  sh.getRange('AD1').setNote('Imported from KDP eBook Royalty sheet (USD).');
  sh.getRange('AE1').setNote('Imported from KDP Paperback + Hardcover Royalty sheets (USD).');
  sh.getRange('AF1').setNote(
    'Estimated KENP / KU royalties (USD) = Lifetime KU Pages × Estimated KENP Royalty Rate. Auto-filled on KDP import / Refresh when the xlsx has pages but no Royalty column. Rate is estimated (not finalized Global Fund).'
  );
  sh.getRange('AG1').setNote('KENPC = Kindle Edition Normalized Page Count for this book (from KDP book details). Enter manually once on the Kindle listing. Used for Estimated Full KU Read Royalty and KU Equivalent Reads.');
  sh.getRange('H1').setNote('Book publication date. Enter manually, or leave blank for Amazon fill when available.');
  sh.getRange('O1').setNote('Auto-filled from Original Release Date (or Amazon publication date) when blank.');
  sh.getRange('T1').setNote('Filled by Update Amazon Rankings Now when available.');
  sh.getRange('U1').setNote('Filled by Amazon rank update when available.');
  sh.getRange('V1').setNote('Filled by Amazon rank update when available.');
}
