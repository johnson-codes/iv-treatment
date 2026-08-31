/**
 * Smile Well IV membership — Google Apps Script backend
 *
 * AFTER PASTING: Deploy → Manage deployments → Edit (pencil) → Version: New version → Deploy
 * Keep the same Web app URL so the website does not need a new link.
 *
 * Writes to the sheet tab whose header row is:
 * Submission_ID | Timestamp | Full_Name | Phone | Email | Selected_Package |
 * Preferred_Date | Health_Notes | Lead_Status | UTM_Source | UTM_Campaign | Internal_Notes
 */

var LOCK_TIMEOUT_MS = 10000;
var MAX_TEXT = 2000;
var HEADERS = [
  'Submission_ID',
  'Timestamp',
  'Full_Name',
  'Phone',
  'Email',
  'Selected_Package',
  'Preferred_Date',
  'Health_Notes',
  'Lead_Status',
  'UTM_Source',
  'UTM_Campaign',
  'Internal_Notes'
];

function setup() {
  var sheet = getSheet_();
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  sheet.getRange(1, 1, 1, HEADERS.length)
    .setFontWeight('bold')
    .setBackground('#1d3557')
    .setFontColor('#ffffff');
  sheet.setFrozenRows(1);

  var widths = [140, 170, 180, 150, 220, 200, 200, 260, 120, 140, 160, 280];
  for (var i = 0; i < widths.length; i++) {
    sheet.setColumnWidth(i + 1, widths[i]);
  }

  var statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['New', 'Contacted', 'Active', 'Expired', 'Closed'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('I2:I').setDataValidation(statusRule);
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

    var sheet = getSheet_();
    var id = nextId_(sheet);
    var timezone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone() || 'America/Vancouver';
    var timestamp = Utilities.formatDate(new Date(), timezone, 'yyyy-MM-dd HH:mm:ss');
    var startDate = sanitize_(data.startDate || data.preferredDate, 20);
    var expiryDate = sanitize_(data.expiryDate, 20) || (startDate ? addYearsISO_(startDate, 1) : '');
    var preferred = compact_([startDate, expiryDate ? 'to ' + expiryDate : ''], ' ');
    if (data.timeWindow && data.timeWindow !== 'Membership term') {
      preferred = compact_([sanitize_(data.preferredDate, 20), '(' + sanitize_(data.timeWindow, 20) + ')'], ' ');
    }

    var pkg = sanitize_(data.package, 80) || 'IV Treatment Membership';
    var notes = compact_([
      sanitize_(data.dob, 20) ? 'DOB: ' + sanitize_(data.dob, 20) : '',
      sanitize_(data.signature, 120) ? 'Signed: ' + sanitize_(data.signature, 120) : '',
      data.agreement ? 'Agreement: Yes' : '',
      data.cardAck ? 'Card on file: Yes' : ''
    ], ' | ');
    var utmSource = compact_([sanitize_(data.utm_source, 120), sanitize_(data.utm_medium, 120)], ' / ');

    sheet.appendRow([
      id,
      timestamp,
      sanitize_(data.fullName, 120),
      sanitize_(data.phone, 40),
      sanitize_(String(data.email || '').toLowerCase(), 120),
      pkg,
      preferred,
      sanitize_(data.healthNotes, MAX_TEXT),
      'New',
      utmSource,
      sanitize_(data.utm_campaign, 120),
      notes
    ]);

    return jsonResponse_({ ok: true, id: id, expiryDate: expiryDate });
  } catch (err) {
    return jsonResponse_({ ok: false, error: 'Server error' });
  } finally {
    try {
      lock.releaseLock();
    } catch (releaseErr) {}
  }
}

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (String(sheets[i].getRange(1, 1).getValue()) === 'Submission_ID') {
      return sheets[i];
    }
  }
  return ss.getActiveSheet();
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
  if (!data.fullName || String(data.fullName).trim().length < 2) return 'Invalid name';
  var digits = String(data.phone || '').replace(/\D/g, '');
  if (digits.length < 10 || digits.length > 15) return 'Invalid phone';
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(data.email || ''))) return 'Invalid email';
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
  return 'IV-' + year + '-' + pad_(seq, 4);
}

function addYearsISO_(iso, years) {
  var parts = iso.split('-').map(Number);
  if (parts.length !== 3 || !parts[0]) return '';
  var end = new Date(parts[0] + years, parts[1] - 1, parts[2]);
  if (end.getMonth() !== parts[1] - 1) end.setDate(0);
  return end.getFullYear() + '-' + pad_(end.getMonth() + 1, 2) + '-' + pad_(end.getDate(), 2);
}

function sanitize_(value, maxLen) {
  var text = value == null ? '' : String(value);
  text = text.replace(/[\u0000-\u001F\u007F]/g, ' ').replace(/\s+/g, ' ').trim();
  if (text.length > maxLen) text = text.slice(0, maxLen);
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
