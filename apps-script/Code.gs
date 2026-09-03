/**
 * Smile Well IV membership — Google Apps Script backend
 *
 * AFTER PASTING: Deploy → Manage deployments → Edit (pencil) → Version: New version → Deploy
 * Keep the same Web app URL so the website does not need a new link.
 * Run setup() once after pasting so Total_Amount is added if the sheet is missing it.
 *
 * Writes to the sheet tab whose header row is:
 * Submission_ID | Timestamp | Full_Name | Phone | Email | Selected_Package |
 * Preferred_Date (or I_Date) | Health_Notes | Lead_Status | UTM_Source | UTM_Campaign |
 * Internal_Notes | Total_Amount
 *
 * Desk POSTs (list, updateTotal) require a valid Firebase ID token.
 * Membership doPost (no action / signup) stays public.
 * Row lookup uses getDataRange(), which includes rows hidden via hideRows().
 */

var LOCK_TIMEOUT_MS = 10000;
var MAX_TEXT = 2000;
var FIREBASE_WEB_API_KEY = 'AIzaSyDgV0ZN5h1MWzQWNwNqe-ZJmy2aBWL8diI';
var FIREBASE_LOOKUP_URL =
  'https://identitytoolkit.googleapis.com/v1/accounts:lookup?key=' + FIREBASE_WEB_API_KEY;
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
  'Internal_Notes',
  'Total_Amount'
];

var HEADER_WIDTHS = {
  Submission_ID: 140,
  Timestamp: 170,
  Full_Name: 180,
  Phone: 150,
  Email: 220,
  Selected_Package: 200,
  Preferred_Date: 200,
  I_Date: 200,
  Health_Notes: 260,
  Lead_Status: 120,
  UTM_Source: 140,
  UTM_Campaign: 160,
  Internal_Notes: 280,
  Total_Amount: 130
};

function setup() {
  var sheet = getSheet_();
  ensureHeaders_(sheet);

  var lastCol = Math.max(sheet.getLastColumn(), HEADERS.length);
  sheet.getRange(1, 1, 1, lastCol)
    .setFontWeight('bold')
    .setBackground('#1d3557')
    .setFontColor('#ffffff');
  sheet.setFrozenRows(1);

  var headers = headerNames_(sheet);
  for (var i = 0; i < headers.length; i++) {
    var width = HEADER_WIDTHS[headers[i]];
    if (width) sheet.setColumnWidth(i + 1, width);
  }

  var statusCol = headerIndex_(sheet, 'Lead_Status') + 1;
  if (statusCol > 0) {
    var statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['New', 'Contacted', 'Active', 'Expired', 'Closed'], true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(2, statusCol, sheet.getMaxRows() - 1, 1).setDataValidation(statusRule);
  }

  var totalCol = headerIndex_(sheet, 'Total_Amount') + 1;
  if (totalCol > 0) {
    sheet.getRange(2, totalCol, sheet.getMaxRows() - 1, 1).setNumberFormat('$#,##0.00');
  }
}

function doGet(e) {
  var action = requestAction_(e, null);
  if (action === 'updateTotal') {
    return jsonResponse_({ ok: false, error: 'POST required' });
  }
  if (action !== 'list') {
    return jsonResponse_({ ok: true, service: 'smile-well-iv-membership' });
  }

  var denied = requireDeskAuth_(e, null);
  if (denied) return denied;

  var lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(LOCK_TIMEOUT_MS)) {
      return jsonResponse_({ ok: false, error: 'Busy, please try again.' });
    }
    return jsonResponse_({ ok: true, clients: listClients_() });
  } catch (err) {
    return jsonResponse_({ ok: false, error: 'Server error' });
  } finally {
    try {
      lock.releaseLock();
    } catch (releaseErr) {}
  }
}

function doPost(e) {
  var data = {};
  var hasBody = Boolean(e && e.postData && e.postData.contents);
  var parseFailed = false;
  if (hasBody) {
    try {
      data = parseBody_(e) || {};
    } catch (parseErr) {
      parseFailed = true;
      data = {};
    }
  }
  var action = requestAction_(e, data);

  if (action === 'list' || action === 'updateTotal') {
    var denied = requireDeskAuth_(e, data);
    if (denied) return denied;
  }

  var lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(LOCK_TIMEOUT_MS)) {
      return jsonResponse_({ ok: false, error: 'Busy, please try again.' });
    }

    // Desk actions first — never fall through to membership validate_().
    if (action === 'list') {
      return jsonResponse_({ ok: true, clients: listClients_() });
    }
    if (action === 'updateTotal') {
      return updateTotal_(data, e);
    }

    if (parseFailed || !hasBody) {
      return jsonResponse_({ ok: false, error: 'Server error' });
    }

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
    var hearAbout = sanitize_(data.hearAbout, 40);
    var hearAboutOther = sanitize_(data.hearAboutOther, 120);
    var heardFrom = '';
    if (hearAbout) {
      heardFrom = 'Heard from: ' + hearAbout;
      if (hearAbout === 'Other' && hearAboutOther) {
        heardFrom += ' (' + hearAboutOther + ')';
      }
    }

    var notes = compact_([
      sanitize_(data.dob, 20) ? 'DOB: ' + sanitize_(data.dob, 20) : '',
      sanitize_(data.signature, 120) ? 'Signed: ' + sanitize_(data.signature, 120) : '',
      data.agreement ? 'Agreement: Yes' : '',
      data.cardAck ? 'Card on file: Yes' : '',
      heardFrom
    ], ' | ');

    var utmSourceRaw = sanitize_(data.utm_source, 120);
    if (!utmSourceRaw && hearAbout) {
      utmSourceRaw = hearAbout;
    }
    var utmSource = compact_([utmSourceRaw, sanitize_(data.utm_medium, 120)], ' / ');

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
      notes,
      ''
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

function listClients_() {
  var sheet = getSheet_();
  // getDataRange() includes rows hidden via hideRows(); do not skip them.
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];

  var idx = headerMap_(values[0]);
  var termCol = firstCol_(idx, ['Preferred_Date', 'I_Date']);
  var totalCol = firstCol_(idx, ['Total_Amount']);

  var clients = [];
  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    var id = String(cellString_(row, idx.Submission_ID) || '').trim();
    if (!id) continue;
    clients.push({
      id: id,
      timestamp: formatCell_(idx.Timestamp != null ? row[idx.Timestamp] : ''),
      fullName: cellString_(row, idx.Full_Name),
      phone: cellString_(row, idx.Phone),
      email: cellString_(row, idx.Email),
      package: cellString_(row, idx.Selected_Package),
      term: cellString_(row, termCol),
      healthNotes: cellString_(row, idx.Health_Notes),
      leadStatus: cellString_(row, idx.Lead_Status),
      utmSource: cellString_(row, idx.UTM_Source),
      utmCampaign: cellString_(row, idx.UTM_Campaign),
      internalNotes: cellString_(row, idx.Internal_Notes),
      totalAmount: formatAmountCell_(totalCol != null ? row[totalCol] : '')
    });
  }
  clients.reverse();
  return clients;
}

function updateTotal_(data, e) {
  data = data || {};
  var param = (e && e.parameter) || {};
  var idRaw = data.id != null && String(data.id) !== '' ? data.id : param.id;
  var amountRaw = data.totalAmount != null && String(data.totalAmount) !== ''
    ? data.totalAmount
    : param.totalAmount;

  var id = sanitize_(idRaw, 40);
  if (!/^IV-[A-Za-z0-9-]{1,32}$/.test(id)) {
    return jsonResponse_({ ok: false, error: 'Invalid id' });
  }

  var amount = sanitizeAmount_(amountRaw);
  if (amount === null) {
    return jsonResponse_({ ok: false, error: 'Invalid amount' });
  }

  var sheet = getSheet_();
  var col = headerIndex_(sheet, 'Total_Amount') + 1;
  if (col < 1) {
    return jsonResponse_({ ok: false, error: 'Total_Amount column missing. Run setup().' });
  }

  var row = findRowById_(sheet, id);
  if (row < 0) {
    return jsonResponse_({ ok: false, error: 'Member not found' });
  }

  sheet.getRange(row, col).setValue(amount === '' ? '' : Number(amount));
  return jsonResponse_({ ok: true, id: id });
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

function ensureHeaders_(sheet) {
  var lastCol = Math.max(sheet.getLastColumn(), 1);
  var existing = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) {
    return String(h || '').trim();
  });

  var nextCol = 1;
  for (var i = 0; i < existing.length; i++) {
    if (existing[i]) nextCol = i + 2;
  }

  if (nextCol === 1 && !existing[0]) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    return;
  }

  for (var j = 0; j < HEADERS.length; j++) {
    if (!headerPresent_(existing, HEADERS[j])) {
      sheet.getRange(1, nextCol).setValue(HEADERS[j]);
      existing.push(HEADERS[j]);
      nextCol++;
    }
  }
}

function headerNames_(sheet) {
  var lastCol = Math.max(sheet.getLastColumn(), 1);
  return sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) {
    return String(h || '').trim();
  });
}

function headerMap_(headers) {
  var idx = {};
  for (var c = 0; c < headers.length; c++) {
    var name = String(headers[c] || '').trim();
    if (name && idx[name] == null) idx[name] = c;
  }
  return idx;
}

function firstCol_(idx, names) {
  for (var i = 0; i < names.length; i++) {
    if (idx[names[i]] != null) return idx[names[i]];
  }
  return null;
}

function headerPresent_(existing, name) {
  if (existing.indexOf(name) !== -1) return true;
  if (name === 'Preferred_Date' && existing.indexOf('I_Date') !== -1) return true;
  return false;
}

function headerIndex_(sheet, name) {
  var headers = headerNames_(sheet);
  var target = String(name || '').trim();
  for (var i = 0; i < headers.length; i++) {
    if (headers[i] === target) return i;
  }
  return -1;
}

function findRowById_(sheet, id) {
  // SpreadsheetApp getDataRange()/getValues() include hidden rows. Do not skip them.
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return -1;
  var idCol = headerMap_(values[0]).Submission_ID;
  if (idCol == null) return -1;
  for (var r = 1; r < values.length; r++) {
    if (String(values[r][idCol] || '').trim() === id) return r + 1;
  }
  return -1;
}

function parseBody_(e) {
  if (!e || !e.postData || !e.postData.contents) {
    throw new Error('Empty body');
  }
  var raw = String(e.postData.contents).replace(/^\uFEFF/, '').trim();
  var parsed = JSON.parse(raw);
  if (!parsed || typeof parsed !== 'object') {
    throw new Error('Invalid JSON');
  }
  return parsed;
}

function requireDeskAuth_(e, data) {
  var token = readIdToken_(e, data);
  if (!firebaseIdTokenEmail_(token)) {
    return jsonResponse_({ ok: false, error: 'Sign in required' });
  }
  return null;
}

function readIdToken_(e, data) {
  if (data && data.idToken != null && String(data.idToken).trim()) {
    return String(data.idToken).trim();
  }

  var header = headerValue_(e, 'Authorization') || headerValue_(e, 'authorization');
  if (header) return bearerToken_(header);

  var param = (e && e.parameter) || {};
  if (param.idToken != null && String(param.idToken).trim()) {
    return String(param.idToken).trim();
  }
  if (param.Authorization) return bearerToken_(param.Authorization);
  if (param.authorization) return bearerToken_(param.authorization);
  return '';
}

function headerValue_(e, name) {
  if (!e) return '';
  if (e.headers && e.headers[name] != null) return String(e.headers[name]);
  if (e.header && e.header[name] != null) return String(e.header[name]);
  return '';
}

function bearerToken_(value) {
  var text = String(value || '').trim();
  var match = text.match(/^Bearer\s+(\S+)/i);
  return match ? match[1] : text;
}

function firebaseIdTokenEmail_(token) {
  if (!token) return '';
  try {
    var response = UrlFetchApp.fetch(FIREBASE_LOOKUP_URL, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ idToken: token }),
      muteHttpExceptions: true
    });
    if (response.getResponseCode() !== 200) return '';
    var body = JSON.parse(response.getContentText() || '{}');
    if (!body || !body.users || !body.users.length) return '';
    var email = body.users[0] && body.users[0].email;
    return email ? String(email).trim() : '';
  } catch (err) {
    return '';
  }
}

function requestAction_(e, data) {
  var fromBody = data && data.action != null ? String(data.action).trim() : '';
  if (fromBody) return fromBody;

  var fromQuery = '';
  if (e && e.parameter && e.parameter.action != null) {
    fromQuery = String(e.parameter.action).trim();
  }
  if (!fromQuery && e && e.parameters && e.parameters.action && e.parameters.action.length) {
    fromQuery = String(e.parameters.action[0] || '').trim();
  }
  if (!fromQuery && e && e.queryString) {
    var match = String(e.queryString).match(/(?:^|&)action=([^&]*)/);
    if (match) {
      try {
        fromQuery = decodeURIComponent(match[1].replace(/\+/g, ' ')).trim();
      } catch (decodeErr) {
        fromQuery = String(match[1] || '').trim();
      }
    }
  }
  return fromQuery;
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

function sanitizeAmount_(value) {
  var text = sanitize_(value, 24);
  if (!text) return '';
  text = text.replace(/CAD/gi, '').replace(/[$,\s]/g, '');
  if (!/^\d+(\.\d{1,2})?$/.test(text)) return null;
  var n = Number(text);
  if (!isFinite(n) || n < 0 || n > 9999999.99) return null;
  return n.toFixed(2);
}

function formatAmountCell_(value) {
  if (value == null || value === '') return '';
  if (typeof value === 'number' && isFinite(value)) return value.toFixed(2);
  var parsed = sanitizeAmount_(value);
  return parsed == null ? sanitize_(value, 24) : parsed;
}

function formatCell_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    var tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone() || 'America/Vancouver';
    return Utilities.formatDate(value, tz, 'yyyy-MM-dd HH:mm:ss');
  }
  return value == null ? '' : String(value);
}

function cellString_(row, index) {
  if (index == null) return '';
  var value = row[index];
  return value == null ? '' : String(value);
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
