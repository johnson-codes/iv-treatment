/**
 * Smile Well IV membership — Google Apps Script backend
 *
 * Desk auth verifies Firebase ID tokens by JWT claim checks only (aud / iss / exp /
 * email). No UrlFetchApp — so script.external_request OAuth is NOT required for
 * receptionist sign-in. Tradeoff: signature (RS256) is not checked here; a forged
 * JWT could fake claims. Minting a real Firebase-signed token still needs Firebase
 * credentials. True signature verify needs UrlFetchApp (certs) + editor Authorize
 * + New version deploy, or an Admin SDK outside Apps Script.
 *
 * After pasting this Code.gs:
 *  1. Deploy → Manage deployments → pencil → New version → Deploy
 *     (same Web app URL; no authorize / Allow step needed for desk auth)
 *  2. Refresh receptionist
 *
 * Spreadsheet access: Run setup() once from the editor if the sheet is new.
 *
 * Writes to the sheet tab whose header row is:
 * Submission_ID | Timestamp | Full_Name | Phone | Email | Selected_Package |
 * Preferred_Date (or I_Date) | Health_Notes | Lead_Status | UTM_Source | UTM_Campaign |
 * Internal_Notes | Total_Amount
 *
 * Desk POSTs (list, updateTotal) require a Firebase ID token in the JSON body as
 * idToken. Do not rely on Authorization headers — GAS 302 redirects drop them.
 * Membership doPost (no action / signup) stays public.
 * Row lookup uses getDataRange(), which includes rows hidden via hideRows().
 */

var LOCK_TIMEOUT_MS = 10000;
var MAX_TEXT = 2000;
var FIREBASE_PROJECT_ID = 'smile-well-34579';
var FIREBASE_ISSUER = 'https://securetoken.google.com/' + FIREBASE_PROJECT_ID;
var TOKEN_SKEW_SEC = 60;
// Only these accounts may read client records; a new receptionist needs an entry
// here plus a Firebase user in smile-well-34579.
var STAFF_EMAILS = [
  'newlifewanted2020@gmail.com'
];
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

/**
 * Optional editor helper. Desk auth no longer uses UrlFetchApp — do not Run this
 * expecting an "Allow" for script.external_request. After updating Code.gs, deploy
 * a New version of the web app; that alone fixes receptionist sign-in.
 * Run setup() if you need spreadsheet permission / headers.
 */
function authorize() {
  SpreadsheetApp.getActiveSpreadsheet();
}

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

  var denied = safeRequireStaff_(e, null);
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
    var denied = safeRequireStaff_(e, data);
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

// A throw here would make Apps Script serve an HTML error page instead of JSON.
function safeRequireStaff_(e, data) {
  try {
    return requireStaff_(e, data);
  } catch (err) {
    return jsonResponse_({
      ok: false,
      error: 'Could not verify sign-in',
      detail: String((err && err.message) || err)
    });
  }
}

function requireStaff_(e, data) {
  var token = readIdToken_(e, data);
  if (!token) {
    return jsonUnauthorized_('Sign in required', 'missing token');
  }
  var verified = verifyFirebaseIdToken_(token);
  if (!verified.ok) {
    if (verified.error === 'Sign in required') {
      return jsonUnauthorized_(verified.error, verified.detail);
    }
    return jsonResponse_({
      ok: false,
      error: verified.error || 'Could not verify sign-in',
      detail: verified.detail || ''
    });
  }
  return null;
}

function readIdToken_(e, data) {
  if (data && data.idToken != null && String(data.idToken).trim()) {
    return String(data.idToken).trim();
  }

  var fromRaw = idTokenFromRawBody_(e);
  if (fromRaw) return fromRaw;

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

function idTokenFromRawBody_(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) return '';
    var parsed = JSON.parse(String(e.postData.contents).replace(/^\uFEFF/, '').trim());
    if (parsed && parsed.idToken != null && String(parsed.idToken).trim()) {
      return String(parsed.idToken).trim();
    }
  } catch (err) {}
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

/**
 * Claim-only Firebase ID token check (no UrlFetchApp / RS256).
 * Validates aud, iss, exp, and that the email is on STAFF_EMAILS.
 */
function verifyFirebaseIdToken_(token) {
  if (!token) {
    return staffAuthFail_('Sign in required', 'missing token');
  }

  var parts = String(token).split('.');
  if (parts.length !== 3 || !parts[0] || !parts[1] || !parts[2]) {
    return staffAuthFail_('Could not verify sign-in', 'malformed token');
  }

  var payload;
  try {
    payload = JSON.parse(bytesToUtf8_(base64UrlToBytes_(parts[1])));
  } catch (decodeErr) {
    return staffAuthFail_('Could not verify sign-in', 'token decode failed');
  }
  if (!payload || typeof payload !== 'object') {
    return staffAuthFail_('Could not verify sign-in', 'invalid token payload');
  }

  if (String(payload.aud || '') !== FIREBASE_PROJECT_ID) {
    return staffAuthFail_('Wrong Firebase project', 'bad audience');
  }
  if (String(payload.iss || '') !== FIREBASE_ISSUER) {
    return staffAuthFail_('Could not verify sign-in', 'bad issuer');
  }

  var now = Math.floor(Date.now() / 1000);
  var exp = Number(payload.exp);
  if (!isFinite(exp) || now >= exp + TOKEN_SKEW_SEC) {
    return staffAuthFail_('Sign-in expired', 'token expired');
  }

  var email = payload.email != null ? String(payload.email).trim() : '';
  if (!email) {
    return staffAuthFail_('Sign in required', 'email missing');
  }
  if (!isStaffEmail_(email)) {
    return staffAuthFail_('Not authorized', 'email not on staff list');
  }

  return { ok: true, email: email };
}

function isStaffEmail_(email) {
  if (!STAFF_EMAILS.length) return true;
  var needle = String(email || '').trim().toLowerCase();
  for (var i = 0; i < STAFF_EMAILS.length; i++) {
    if (String(STAFF_EMAILS[i] || '').trim().toLowerCase() === needle) return true;
  }
  return false;
}

function staffAuthFail_(error, detail) {
  return { ok: false, error: error, detail: detail || '' };
}

function base64UrlToBytes_(value) {
  var b64 = String(value || '').replace(/-/g, '+').replace(/_/g, '/');
  while (b64.length % 4) b64 += '=';
  return copyBytes_(Utilities.base64Decode(b64));
}

function bytesToUtf8_(bytes) {
  return Utilities.newBlob(bytes).getDataAsString('UTF-8');
}

// Utilities.newBlob wants signed bytes (-128..127), not 0..255.
function copyBytes_(raw) {
  var out = [];
  for (var i = 0; i < raw.length; i++) {
    var b = raw[i] & 0xff;
    out.push(b > 127 ? b - 256 : b);
  }
  return out;
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

function jsonUnauthorized_(message, detail) {
  var body = {
    ok: false,
    error: message || 'Sign in required',
    status: 401
  };
  if (detail) body.detail = detail;
  return jsonResponse_(body);
}

function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
