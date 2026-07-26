/**
 * Events table — marketing / release / promo timeline.
 * Designed so events can later appear as markers on Dashboard charts.
 */

function ensureEventsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(AD.SHEETS.EVENTS);
  if (!sh) {
    sh = ss.insertSheet(AD.SHEETS.EVENTS);
    buildEventsSheet_(sh);
  } else {
    ensureEventsSchema_(sh);
  }
  return sh;
}

function buildEventsSheet_(sheet) {
  sheet.clear();
  sheet.getRange(1, 1, 1, AD.EVENT_HEADERS.length).setValues([AD.EVENT_HEADERS]);
  styleHeader_(sheet.getRange(1, 1, 1, AD.EVENT_HEADERS.length));
  sheet.setFrozenRows(1);
  [120, 100, 100, 100, 160, 220, 360].forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  sheet.getRange('B:C').setNumberFormat('m/d/yyyy');
  applyEventsValidation_(sheet);
  addFilter_(sheet, AD.EVENT_HEADERS.length);
  sheet.getRange(1, 1).setNote('Permanent Event ID (EVT-001…). Optional Book ID links to SOL-###.');
  sheet.getRange(1, 4).setNote('Optional. Use SOL-001 style Book ID, not title.');
  sheet.getRange(1, 5).setNote('Event type — used later for chart markers.');
  sheet.getRange(1, 7).setNote(
    'Refresh auto-adds Book release + Meta campaign started (description starts with AUTO:…). ' +
      'Manual events are never overwritten. Add more via Author Dashboard → Add Event.'
  );
}

function ensureEventsSchema_(sh) {
  const cols = AD.EVENT_HEADERS.length;
  const headers = sh.getRange(1, 1, 1, cols).getValues()[0].map(clean_);
  let changed = false;
  AD.EVENT_HEADERS.forEach((wanted, i) => {
    if (headers[i] !== wanted) {
      sh.getRange(1, i + 1).setValue(wanted);
      changed = true;
    }
  });
  if (changed) styleHeader_(sh.getRange(1, 1, 1, cols));
  applyEventsValidation_(sh);
}

function applyEventsValidation_(sh) {
  const list = SpreadsheetApp.newDataValidation()
    .requireValueInList(AD.EVENT_TYPES, true)
    .setAllowInvalid(true)
    .build();
  sh.getRange('E2:E').setDataValidation(list);
}

function addEventDialog() {
  ensureEventsSheet_();
  const ui = SpreadsheetApp.getUi();

  const typeRes = ui.prompt(
    'Add event — type',
    'Event type:\n' + AD.EVENT_TYPES.join('\n'),
    ui.ButtonSet.OK_CANCEL
  );
  if (typeRes.getSelectedButton() !== ui.Button.OK) return;
  let eventType = clean_(typeRes.getResponseText());
  if (!eventType) eventType = 'Other';
  const known = AD.EVENT_TYPES.some(t => normalizeKey_(t) === normalizeKey_(eventType));
  if (!known) {
    const match = AD.EVENT_TYPES.find(t => normalizeKey_(t).indexOf(normalizeKey_(eventType)) !== -1);
    eventType = match || eventType;
  }

  const nameRes = ui.prompt('Add event — name', 'Short event name:', ui.ButtonSet.OK_CANCEL);
  if (nameRes.getSelectedButton() !== ui.Button.OK) return;
  const eventName = clean_(nameRes.getResponseText());
  if (!eventName) {
    ui.alert('Event name is required.');
    return;
  }

  const dateRes = ui.prompt(
    'Add event — date',
    'Start date (YYYY-MM-DD). Leave blank for today:',
    ui.ButtonSet.OK_CANCEL
  );
  if (dateRes.getSelectedButton() !== ui.Button.OK) return;
  const dateText = clean_(dateRes.getResponseText());
  const startDate = dateText ? parseLooseDate_(dateText) : getSpreadsheetToday_();
  if (!startDate) {
    ui.alert('Could not parse date.');
    return;
  }

  const endRes = ui.prompt(
    'Add event — end date (optional)',
    'End date (YYYY-MM-DD) or leave blank:',
    ui.ButtonSet.OK_CANCEL
  );
  if (endRes.getSelectedButton() !== ui.Button.OK) return;
  const endText = clean_(endRes.getResponseText());
  const endDate = endText ? parseLooseDate_(endText) : '';

  const bookRes = ui.prompt(
    'Add event — book (optional)',
    'Book ID (SOL-001), book title, or leave blank:',
    ui.ButtonSet.OK_CANCEL
  );
  if (bookRes.getSelectedButton() !== ui.Button.OK) return;
  const bookRaw = clean_(bookRes.getResponseText());
  let bookId = '';
  if (bookRaw) {
    if (isStableBookId_(bookRaw) || /^SOL-\d+/i.test(bookRaw)) {
      bookId = bookRaw.toUpperCase();
    } else {
      const hit = getInputRows_().find(r => normalizeKey_(r[AD.COL.TITLE]) === normalizeKey_(bookRaw));
      bookId = hit ? clean_(hit[AD.COL.BOOK_ID]) : '';
      if (!bookId) {
        ui.alert('No Book ID match for "' + bookRaw + '". Event will be saved without a book link.');
      }
    }
  }

  const descRes = ui.prompt(
    'Add event — description (optional)',
    'Notes:',
    ui.ButtonSet.OK_CANCEL
  );
  if (descRes.getSelectedButton() !== ui.Button.OK) return;
  const description = clean_(descRes.getResponseText());

  const eventId = appendEvent_({
    date: startDate,
    endDate: endDate || '',
    bookId: bookId || '',
    eventType: eventType,
    eventName: eventName,
    description: description
  });

  ui.alert('Event recorded: ' + eventId + '\n\nOpen the Events sheet to edit or add more.');
}

function appendEvent_(evt) {
  const sh = ensureEventsSheet_();
  const eventId = generateEventId_();
  appendRows_(sh, [[
    eventId,
    startOfDay_(evt.date),
    evt.endDate ? startOfDay_(evt.endDate) : '',
    clean_(evt.bookId),
    clean_(evt.eventType),
    clean_(evt.eventName),
    clean_(evt.description)
  ]]);
  sh.getRange('B:C').setNumberFormat('m/d/yyyy');
  return eventId;
}

function generateEventId_() {
  let max = 0;
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(AD.SHEETS.EVENTS);
  if (sh && sh.getLastRow() >= 2) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().forEach(r => {
      const m = String(r[0] || '').match(/^EVT-(\d+)$/i);
      if (m) max = Math.max(max, Number(m[1]));
    });
  }
  return 'EVT-' + ('000' + (max + 1)).slice(-3);
}

/**
 * Auto-fill safe Events (does not overwrite manual rows).
 * - Book release from Manual Entry Original / Listing release dates
 * - Meta campaign started from first spend day in Meta Daily
 */
function syncAutoEvents_() {
  ensureEventsSheet_();
  const existing = getEventAutoKeys_();
  let added = 0;

  // Book releases
  const seenBooks = new Set();
  getInputRows_().forEach(r => {
    const bookId = clean_(r[AD.COL.BOOK_ID]);
    const title = clean_(r[AD.COL.TITLE]);
    if (!bookId || !title || seenBooks.has(bookId)) return;
    seenBooks.add(bookId);
    const release = isValidDate_(r[AD.COL.ORIGINAL_RELEASE])
      ? startOfDay_(new Date(r[AD.COL.ORIGINAL_RELEASE]))
      : (isValidDate_(r[AD.COL.LISTING_RELEASE]) ? startOfDay_(new Date(r[AD.COL.LISTING_RELEASE])) : null);
    if (!release) return;
    const key = 'AUTO:RELEASE:' + bookId;
    if (existing.has(key)) return;
    appendEvent_({
      date: release,
      endDate: '',
      bookId: bookId,
      eventType: 'Book release',
      eventName: title + ' released',
      description: key
    });
    existing.add(key);
    added++;
  });

  // Meta campaigns
  const campaigns = summarizeMetaCampaigns_();
  campaigns.forEach(c => {
    const key = 'AUTO:META_START:' + c.campaignId;
    if (existing.has(key) || !c.dateMin) return;
    appendEvent_({
      date: parseLooseDate_(c.dateMin) || getSpreadsheetToday_(),
      endDate: c.dateMax ? (parseLooseDate_(c.dateMax) || '') : '',
      bookId: c.bookId || '',
      eventType: 'Meta campaign started',
      eventName: c.campaignName || ('Meta campaign ' + c.campaignId),
      description: key + ' | spend $' + c.spend.toFixed(2) + ' (auto from Meta Daily)'
    });
    existing.add(key);
    added++;
  });

  return added;
}

function getEventAutoKeys_() {
  const set = new Set();
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(AD.SHEETS.EVENTS);
  if (!sh || sh.getLastRow() < 2) return set;
  sh.getRange(2, 1, sh.getLastRow() - 1, AD.EVENT_HEADERS.length).getValues().forEach(r => {
    const desc = clean_(r[6]);
    if (/^AUTO:RELEASE:/i.test(desc) || /^AUTO:META_START:/i.test(desc)) {
      set.add(desc.split(/\s|\|/)[0]);
    }
  });
  return set;
}

/** Events for a date range — later used as chart markers. */
function getEventsInRange_(startDate, endDate) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(AD.SHEETS.EVENTS);
  if (!sh || sh.getLastRow() < 2) return [];
  const start = startDate ? startOfDay_(startDate).getTime() : null;
  const end = endDate ? startOfDay_(endDate).getTime() : null;
  return sh.getRange(2, 1, sh.getLastRow() - 1, AD.EVENT_HEADERS.length).getValues()
    .filter(r => isValidDate_(r[1]))
    .map(r => ({
      eventId: clean_(r[0]),
      date: startOfDay_(new Date(r[1])),
      endDate: isValidDate_(r[2]) ? startOfDay_(new Date(r[2])) : null,
      bookId: clean_(r[3]),
      eventType: clean_(r[4]),
      eventName: clean_(r[5]),
      description: clean_(r[6])
    }))
    .filter(e => {
      const t = e.date.getTime();
      if (start != null && t < start) return false;
      if (end != null && t > end) return false;
      return true;
    });
}
