/**
 * Smile Well IV membership — Google Apps Script backend
 *
 * SETUP (run once in the Google Sheet):
 * 1. Create a Google Sheet.
 * 2. Extensions → Apps Script, paste this file as Code.gs.
 * 3. Select setup in the function dropdown and click Run. Approve permissions.
 * 4. Deploy → New deployment → Type: Web app
 *      Execute as: Me
 *      Who has access: Anyone
 * 5. Copy the Web app URL into form.js as WEBHOOK_URL, then redeploy the site.
 */

var SHEET_NAME = 'Memberships';
var ID_PREFIX = 'IV';
var LOCK_TIMEOUT_MS = 10000;
var MAX_TEXT = 2000;
var MIN_AGE = 18;
var HEADERS = [
  'Submission_ID',
  'Timestamp',
  'Full_Name',
  'Date_of_Birth',
  'Phone',
  'Email',
  'Start_Date',
  'Expiry_Date',
  'Signature',
  'Agreement_Accepted',
  'Card_On_File_Ack',
  'Contact_Consent',
  'Health_Notes',
  'Lead_Status',
  'UTM_Source',
  'UTM_Campaign',
  'Internal_Notes'
];

function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  var firstHeader = String(sheet.getRange(1, 1).getValue());
  if (firstHeader !== 'Submission_ID') {
    if (sheet.getLastRow() > 0) {
      sheet.insertRowBefore(1);
    }
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  } else if (sheet.getLastColumn() < HEADERS.length) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  }

  sheet.getRange(1, 1, 1, HEADERS.length)
    .setFontWeight('bold')
    .setBackground('#1d3557')
    .setFontColor('#ffffff');
  sheet.setFrozenRows(1);

  var widths = [140, 170, 180, 130, 150, 220, 120, 120, 180, 150, 140, 140, 260, 120, 140, 160, 240];
  for (var i = 0; i < widths.length; i++) {
    sheet.setColumnWidth(i + 1, widths[i]);
  }

  var statusCol = col_('Lead_Status');
  var statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['New', 'Contacted', 'Active', 'Expired', 'Closed'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, statusCol, Math.max(sheet.getMaxRows() - 1, 1), 1).setDataValidation(statusRule);
}

function doGet() {
  return jsonResponse_({ ok: true, service: 'smile-well-iv-membership' });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(LOCK_TIMEOUT_MS)) {
      return jsonResponse_({ ok: false, error: 'Busy, please try again.' });
    }

    var data = parseBody_(e);
    if (isSpam_(data)) {
      return jsonResponse_({ ok: true, id: 'IV-OK' });
    }

    var error = validate_(data);
    if (error) {
      return jsonResponse_({ ok: false, error: error });
    }

    var startDate = sanitize_(data.startDate, 20);
    var expiryDate = addYearsISO_(startDate, 1);
    var sheet = getSheet_();
    var id = nextId_(sheet);
    var timezone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone() || 'America/Vancouver';
    var timestamp = Utilities.formatDate(new Date(), timezone, 'yyyy-MM-dd HH:mm:ss');
    var utmSource = compact_([sanitize_(data.utm_source, 120), sanitize_(data.utm_medium, 120)], ' / ');

    sheet.appendRow([
      id,
      timestamp,
      sanitize_(data.fullName, 120),
      sanitize_(data.dob, 20),
      sanitize_(data.phone, 40),
      sanitize_(String(data.email || '').toLowerCase(), 120),
      startDate,
      expiryDate,
      sanitize_(data.signature, 120),
      'Yes',
      'Yes',
      'Yes',
      sanitize_(data.healthNotes, MAX_TEXT),
      'New',
      utmSource,
      sanitize_(data.utm_campaign, 120),
      ''
    ]);

    return jsonResponse_({ ok: true, id: id, expiryDate: expiryDate });
  } catch (err) {
    return jsonResponse_({ ok: false, error: 'Server error' });
  } finally {
    try {
      lock.releaseLock();
    } catch (releaseErr) {
      // lock may not have been acquired
    }
  }
}

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    setup();
    sheet = ss.getSheetByName(SHEET_NAME);
  }
  return sheet;
}

function parseBody_(e) {
  if (!e || !e.postData || !e.postData.contents) {
    throw new Error('Empty body');
  }
  var parsed = JSON.parse(e.postData.contents);
  if (!parsed || typeof parsed !== 'object') {
    throw new Error('Invalid JSON');
  }
  return parsed;
}

function isSpam_(data) {
  return Boolean(sanitize_(data.honeypot || data.company_website, 200));
}

function validate_(data) {
  var name = String(data.fullName || '').trim();
  if (name.length < 2) return 'Invalid name';

  var dob = String(data.dob || '');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(dob)) return 'Invalid date of birth';
  if (ageOn_(dob, todayISO_()) < MIN_AGE) return 'Must be 18 or older';

  var digits = String(data.phone || '').replace(/\D/g, '');
  if (digits.length < 10 || digits.length > 15) return 'Invalid phone';
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(data.email || ''))) return 'Invalid email';

  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(data.startDate || ''))) return 'Invalid start date';

  var signature = String(data.signature || '').trim();
  if (!signature) return 'Signature required';
  if (signature.toLowerCase().replace(/\s+/g, ' ') !== name.toLowerCase().replace(/\s+/g, ' ')) {
    return 'Signature must match full name';
  }

  if (data.agreement !== true && data.agreement !== 'true') return 'Agreement required';
  if (data.cardAck !== true && data.cardAck !== 'true') return 'Card acknowledgement required';
  if (data.consent !== true && data.consent !== 'true') return 'Consent required';
  return '';
}

function nextId_(sheet) {
  var year = new Date().getFullYear();
  var lastRow = sheet.getLastRow();
  var seq = 1;
  if (lastRow > 1) {
    var lastId = String(sheet.getRange(lastRow, 1).getValue());
    var match = lastId.match(/^IV-(\d{4})-(\d+)$/);
    if (match && Number(match[1]) === year) {
      seq = Number(match[2]) + 1;
    }
  }
  return ID_PREFIX + '-' + year + '-' + pad_(seq, 4);
}

function todayISO_() {
  return Utilities.formatDate(new Date(), SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone() || 'America/Vancouver', 'yyyy-MM-dd');
}

function ageOn_(dobISO, onISO) {
  var dob = dobISO.split('-').map(Number);
  var on = onISO.split('-').map(Number);
  var age = on[0] - dob[0];
  if (on[1] < dob[1] || (on[1] === dob[1] && on[2] < dob[2])) age -= 1;
  return age;
}

function addYearsISO_(iso, years) {
  var parts = iso.split('-').map(Number);
  var end = new Date(parts[0] + years, parts[1] - 1, parts[2]);
  if (end.getMonth() !== parts[1] - 1) end.setDate(0);
  return end.getFullYear() + '-' + pad_(end.getMonth() + 1, 2) + '-' + pad_(end.getDate(), 2);
}

function col_(name) {
  return HEADERS.indexOf(name) + 1;
}

function sanitize_(value, maxLen) {
  var text = value == null ? '' : String(value);
  text = text.replace(/[\u0000-\u001F\u007F]/g, ' ').replace(/\s+/g, ' ').trim();
  if (text.length > maxLen) {
    text = text.slice(0, maxLen);
  }
  return text;
}

function compact_(parts, joiner) {
  var kept = [];
  for (var i = 0; i < parts.length; i++) {
    if (parts[i]) kept.push(parts[i]);
  }
  return kept.join(joiner);
}

function pad_(num, width) {
  var s = String(num);
  while (s.length < width) s = '0' + s;
  return s;
}

function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
