/**
 * Code.gs
 * ---------------------------------------------------------------------------
 * Web app entry point and shared server-side helpers.
 *
 * SECURITY NOTE: every top-level function whose name does NOT end with an
 * underscore can be called from the browser through google.script.run.
 * All internal helpers therefore end with "_". Public functions always
 * validate the session token before touching any data.
 * ---------------------------------------------------------------------------
 */

/* ======================================================================== */
/* Web app                                                                  */
/* ======================================================================== */

/**
 * Serves the applicant portal (default) or the admin dashboard (?page=admin).
 */
function doGet(e) {
  var page = (e && e.parameter && e.parameter.page) ? String(e.parameter.page) : '';
  var isAdmin = page === 'admin';
  var template = HtmlService.createTemplateFromFile(isAdmin ? 'AdminPage' : 'Index');
  template.publicConfigJson = safeJsonForHtml_(getPublicConfig_(isAdmin));
  template.portalTitle = CONFIG.PORTAL_TITLE;
  template.orgName = CONFIG.ORGANISATION_NAME;
  return template.evaluate()
    .setTitle(isAdmin ? CONFIG.PORTAL_TITLE + ' – Administration' : CONFIG.PORTAL_TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, shrink-to-fit=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT); // the page cannot be framed by other sites
}

/**
 * Includes an HTML partial inside a template. Only known partials are allowed.
 */
function include(filename) {
  var allowed = ['Styles', 'JavaScript', 'Login', 'ApplicationForm', 'Preview', 'AdminJavaScript'];
  if (allowed.indexOf(filename) === -1) {
    throw new Error('Unknown include: ' + filename);
  }
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * The only configuration values the browser ever receives. No IDs, no
 * credentials, no folder names.
 */
function getPublicConfig_(isAdmin) {
  return {
    portalTitle: CONFIG.PORTAL_TITLE,
    organisation: CONFIG.ORGANISATION_NAME,
    supportText: CONFIG.SUPPORT_TEXT,
    portalOpen: CONFIG.PORTAL_OPEN === true,
    portalClosedMessage: CONFIG.PORTAL_CLOSED_MESSAGE,
    idleTimeoutSeconds: (isAdmin ? CONFIG.SESSION.ADMIN_IDLE_TIMEOUT_MINUTES : CONFIG.SESSION.IDLE_TIMEOUT_MINUTES) * 60,
    maxFileSizeMb: CONFIG.UPLOAD.MAX_FILE_SIZE_MB,
    allowedExtensions: CONFIG.UPLOAD.ALLOWED_EXTENSIONS,
    mobileRegex: CONFIG.MOBILE_REGEX,
    requireApproval: CONFIG.REQUIRE_APPROVAL === true,
    maxTextLength: CONFIG.MAX_TEXT_LENGTH,
    maxTextareaLength: CONFIG.MAX_TEXTAREA_LENGTH
  };
}

/* ======================================================================== */
/* Response wrapper and errors                                              */
/* ======================================================================== */

/**
 * Creates an error whose message is safe to show to the user.
 * @param {string} code    machine-readable code (SESSION_EXPIRED, VALIDATION, ...)
 * @param {string} message human-readable message
 * @param {Object=} details optional extra data (e.g. per-field errors)
 */
function portalError_(code, message, details) {
  var err = new Error(message);
  err.code = code;
  err.isPortalError = true;
  err.details = details || null;
  return err;
}

/**
 * Runs fn and converts the result into {ok, data} / {ok:false, code, message}.
 * Unexpected errors are logged with a reference number; internals are never
 * sent to the browser.
 */
function respond_(fn) {
  try {
    return { ok: true, data: fn() };
  } catch (err) {
    if (err && err.isPortalError) {
      return { ok: false, code: err.code, message: err.message, details: err.details || null };
    }
    var ref = Utilities.formatDate(new Date(), 'UTC', 'yyMMddHHmmss') + '-' + Math.floor(Math.random() * 9000 + 1000);
    console.error('[' + ref + '] ' + (err && err.stack ? err.stack : err));
    return {
      ok: false,
      code: 'SERVER_ERROR',
      message: 'An unexpected error occurred while processing your request. Please try again. If the problem continues, quote reference ' + ref + ' to the helpdesk.'
    };
  }
}

/**
 * Blocks functions that must only be run by the project owner from the
 * Apps Script editor (setup, admin account management, sample data).
 * Anonymous web-app visitors have no active user and are rejected.
 */
function requireEditorContext_() {
  var active = '';
  var effective = '';
  try { active = Session.getActiveUser().getEmail(); } catch (e) { active = ''; }
  try { effective = Session.getEffectiveUser().getEmail(); } catch (e) { effective = ''; }
  if (!active || !effective || active.toLowerCase() !== effective.toLowerCase()) {
    throw new Error('This function can only be run by the project owner from the Apps Script editor.');
  }
}

/* ======================================================================== */
/* Spreadsheet access                                                       */
/* ======================================================================== */

var RUNTIME_ = { spreadsheet: null, logSpreadsheet: null, timeZone: null, folders: {} };

function getSpreadsheetId_() {
  var id = extractSpreadsheetId_(CONFIG.SPREADSHEET_ID);
  if (id) return id;
  var prop = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  return extractSpreadsheetId_(prop);
}

/** Accepts a bare ID or a full spreadsheet URL (…/spreadsheets/d/<ID>/edit…). */
function extractSpreadsheetId_(value) {
  var s = String(value == null ? '' : value).trim();
  if (!s || s === 'PASTE_SPREADSHEET_ID_HERE') return '';
  var m = s.match(/\/d\/([A-Za-z0-9_-]{20,})/);
  return m ? m[1] : s;
}

function getSpreadsheet_() {
  if (RUNTIME_.spreadsheet) return RUNTIME_.spreadsheet;
  var id = getSpreadsheetId_();
  var ss = null;
  if (id) {
    ss = SpreadsheetApp.openById(id);
  } else {
    ss = SpreadsheetApp.getActiveSpreadsheet(); // container-bound script
  }
  if (!ss) {
    throw portalError_('CONFIG', 'The portal is not configured yet (spreadsheet not set). Please contact the administrator.');
  }
  RUNTIME_.spreadsheet = ss;
  return ss;
}

function getLogSpreadsheet_() {
  if (RUNTIME_.logSpreadsheet) return RUNTIME_.logSpreadsheet;
  RUNTIME_.logSpreadsheet = extractSpreadsheetId_(CONFIG.LOG_SPREADSHEET_ID) ? SpreadsheetApp.openById(extractSpreadsheetId_(CONFIG.LOG_SPREADSHEET_ID)) : getSpreadsheet_();
  return RUNTIME_.logSpreadsheet;
}

function getTimeZone_() {
  if (!RUNTIME_.timeZone) {
    RUNTIME_.timeZone = getSpreadsheet_().getSpreadsheetTimeZone() || Session.getScriptTimeZone() || 'Asia/Kolkata';
  }
  return RUNTIME_.timeZone;
}

function getApplicationSheet_() {
  var sheet = getSpreadsheet_().getSheetByName(CONFIG.APPLICATION_SHEET_NAME);
  if (!sheet) {
    throw portalError_('CONFIG', 'Application sheet "' + CONFIG.APPLICATION_SHEET_NAME + '" was not found. Please contact the administrator.');
  }
  return sheet;
}

function getLogSheet_() {
  var sheet = getLogSpreadsheet_().getSheetByName(CONFIG.LOG_SHEET_NAME);
  if (!sheet) {
    throw portalError_('CONFIG', 'Correction log sheet is missing. The administrator must run setupPortal().');
  }
  return sheet;
}

/** Reads the header row (one API call). */
function readHeaders_(sheet) {
  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) throw portalError_('CONFIG', 'The application sheet has no header row.');
  return sheet.getRange(CONFIG.HEADER_ROW, 1, 1, lastCol).getDisplayValues()[0].map(function (h) {
    return String(h).trim();
  });
}

function normalizeHeader_(header) {
  return String(header == null ? '' : header).trim().toLowerCase().replace(/\s+/g, ' ');
}

/** Returns the 0-based index of the first alias found in headers, or -1. */
function findColumn_(headers, aliases) {
  var list = Array.isArray(aliases) ? aliases : [aliases];
  var normalized = headers.map(normalizeHeader_);
  for (var i = 0; i < list.length; i++) {
    var idx = normalized.indexOf(normalizeHeader_(list[i]));
    if (idx !== -1) return idx;
  }
  return -1;
}

/**
 * Resolves the indexes of key and system columns.
 * @return {{appNo:number,email:number,mobile:number,dob:number,name:number,
 *           lastModified:number,count:number,status:number,txn:number}}
 */
function resolveKeyColumns_(headers, strict) {
  var k = CONFIG.KEY_COLUMNS;
  var s = CONFIG.SYSTEM_COLUMNS;
  var cols = {
    appNo: findColumn_(headers, k.APPLICATION_NUMBER),
    email: findColumn_(headers, k.EMAIL),
    mobile: findColumn_(headers, k.MOBILE),
    dob: findColumn_(headers, k.DOB),
    name: findColumn_(headers, k.NAME),
    lastModified: findColumn_(headers, s.LAST_MODIFIED),
    count: findColumn_(headers, s.CORRECTION_COUNT),
    status: findColumn_(headers, s.CORRECTION_STATUS),
    txn: findColumn_(headers, s.LAST_TRANSACTION_ID)
  };
  if (strict !== false) {
    var missing = [];
    if (cols.appNo === -1) missing.push('Application Number');
    if (cols.dob === -1) missing.push('Date of Birth');
    if (cols.email === -1 && cols.mobile === -1) missing.push('Email or Mobile Number');
    if (cols.lastModified === -1 || cols.count === -1 || cols.status === -1 || cols.txn === -1) missing.push('portal system columns (run setupPortal)');
    if (missing.length) {
      throw portalError_('CONFIG', 'Portal configuration error – missing column(s): ' + missing.join(', ') + '. Please contact the administrator.');
    }
  }
  return cols;
}

/**
 * Loads everything most requests need: sheet, fresh headers, key columns and
 * field configuration. Headers are read fresh on every request (one cheap
 * call) so a column inserted by an administrator can never cause a write to
 * the wrong cell.
 */
function getSheetContext_() {
  var sheet = getApplicationSheet_();
  var headers = readHeaders_(sheet);
  var cols = resolveKeyColumns_(headers, true);
  var fieldConfig = buildFieldConfig_(headers, cols);
  var lastRow = sheet.getLastRow();
  return {
    sheet: sheet,
    headers: headers,
    cols: cols,
    fields: fieldConfig.fields,
    byHeader: fieldConfig.byHeader,
    lastRow: lastRow,
    dataRows: Math.max(0, lastRow - CONFIG.HEADER_ROW)
  };
}

/** Reads a single full column of data (excluding header). */
function readColumn_(ctx, colIndex, display) {
  if (colIndex < 0 || ctx.dataRows === 0) return [];
  var range = ctx.sheet.getRange(CONFIG.HEADER_ROW + 1, colIndex + 1, ctx.dataRows, 1);
  var values = display ? range.getDisplayValues() : range.getValues();
  return values.map(function (r) { return r[0]; });
}

/**
 * Finds the sheet row (1-based) of an application number using TextFinder
 * (server-side search, no bulk read). Throws if missing or duplicated.
 */
function findRowByAppNo_(ctx, appNo) {
  if (ctx.dataRows === 0) throw portalError_('NOT_FOUND', 'Application not found.');
  var range = ctx.sheet.getRange(CONFIG.HEADER_ROW + 1, ctx.cols.appNo + 1, ctx.dataRows, 1);
  var matches = range.createTextFinder(String(appNo)).matchEntireCell(true).matchCase(false).findAll();
  if (!matches || matches.length === 0) {
    throw portalError_('NOT_FOUND', 'Application ' + appNo + ' was not found.');
  }
  if (matches.length > 1) {
    console.error('Duplicate application number in sheet: ' + appNo);
    throw portalError_('DATA_INTEGRITY', 'Application number ' + appNo + ' appears more than once in the records. Corrections are blocked until the administrator resolves this.');
  }
  return matches[0].getRow();
}

/** Reads one complete row: values, display values and formulas. */
function readRow_(ctx, rowIndex) {
  var range = ctx.sheet.getRange(rowIndex, 1, 1, ctx.headers.length);
  return {
    rowIndex: rowIndex,
    values: range.getValues()[0],
    display: range.getDisplayValues()[0],
    formulas: range.getFormulas()[0]
  };
}

/* ======================================================================== */
/* Normalisation, sanitisation and formatting                               */
/* ======================================================================== */

function cleanAppNo_(value) {
  var s = String(value == null ? '' : value).trim();
  if (!s || s.length > 60 || /[\u0000-\u001F<>"'`]/.test(s)) {
    throw portalError_('VALIDATION', 'Invalid application number.');
  }
  return s;
}

function normalizeEmail_(value) {
  return String(value == null ? '' : value).trim().toLowerCase();
}

function normalizeMobile_(value) {
  if (value === null || value === undefined) return '';
  var s = (typeof value === 'number') ? value.toFixed(0) : String(value);
  var digits = s.replace(/\D/g, '');
  if (digits.length > 10) digits = digits.slice(-10); // strip +91 / 0 prefixes
  return digits;
}

function isValidEmail_(value) {
  return /^[^\s@<>()[\]\\,;:"]+@[^\s@<>()[\]\\,;:"]+\.[A-Za-z]{2,}$/.test(String(value)) && String(value).length <= 254;
}

function isValidMobile_(value) {
  return new RegExp(CONFIG.MOBILE_REGEX).test(String(value));
}

function isDateObject_(v) {
  return Object.prototype.toString.call(v) === '[object Date]';
}

var MONTHS_ = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, sept: 9, oct: 10, nov: 11, dec: 12 };

function buildIsoDate_(y, m, d) {
  y = Number(y); m = Number(m); d = Number(d);
  if (!y || !m || !d || m < 1 || m > 12 || d < 1 || d > 31 || y < 1800 || y > 2200) return '';
  var probe = new Date(Date.UTC(y, m - 1, d));
  if (probe.getUTCFullYear() !== y || probe.getUTCMonth() !== m - 1 || probe.getUTCDate() !== d) return '';
  return y + '-' + (m < 10 ? '0' : '') + m + '-' + (d < 10 ? '0' : '') + d;
}

/**
 * Converts any reasonable date representation found in a sheet (Date object,
 * serial number, dd/mm/yyyy, yyyy-mm-dd, 15-Aug-2000 ...) into yyyy-MM-dd.
 * Returns '' when the value is not a recognisable date.
 */
function toIsoDate_(value) {
  if (value === null || value === undefined || value === '') return '';
  if (isDateObject_(value)) {
    if (isNaN(value.getTime())) return '';
    return Utilities.formatDate(value, getTimeZone_(), 'yyyy-MM-dd');
  }
  if (typeof value === 'number') {
    if (value > 1000 && value < 100000) { // spreadsheet serial date
      var ms = Date.UTC(1899, 11, 30) + Math.round(value) * 86400000;
      var dt = new Date(ms);
      return buildIsoDate_(dt.getUTCFullYear(), dt.getUTCMonth() + 1, dt.getUTCDate());
    }
    return '';
  }
  var s = String(value).trim();
  var m = s.match(/^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})(?:[T\s].*)?$/);
  if (m) return buildIsoDate_(m[1], m[2], m[3]);
  m = s.match(/^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{4})(?:\s.*)?$/);
  if (m) {
    return CONFIG.DATE.TEXT_DATE_ORDER === 'MDY' ? buildIsoDate_(m[3], m[1], m[2]) : buildIsoDate_(m[3], m[2], m[1]);
  }
  m = s.match(/^(\d{1,2})[-\s\/.]+([A-Za-z]{3,9})[-\s\/.,]+(\d{4})$/);
  if (m && MONTHS_[m[2].slice(0, 4).toLowerCase()] !== undefined) return buildIsoDate_(m[3], MONTHS_[m[2].slice(0, 4).toLowerCase()], m[1]);
  if (m && MONTHS_[m[2].slice(0, 3).toLowerCase()] !== undefined) return buildIsoDate_(m[3], MONTHS_[m[2].slice(0, 3).toLowerCase()], m[1]);
  m = s.match(/^([A-Za-z]{3,9})\s+(\d{1,2}),?\s+(\d{4})$/);
  if (m && MONTHS_[m[1].slice(0, 3).toLowerCase()] !== undefined) return buildIsoDate_(m[3], MONTHS_[m[1].slice(0, 3).toLowerCase()], m[2]);
  return '';
}

/** yyyy-MM-dd -> display format (dd/MM/yyyy by default). */
function formatIsoForDisplay_(iso) {
  var m = String(iso || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return String(iso || '');
  return CONFIG.DATE.DISPLAY_FORMAT.replace('yyyy', m[1]).replace('MM', m[2]).replace('dd', m[3]);
}

/** yyyy-MM-dd -> Date at midnight in the spreadsheet's time zone. */
function isoToSheetDate_(iso) {
  return Utilities.parseDate(iso, getTimeZone_(), 'yyyy-MM-dd');
}

function formatTimestamp_(date) {
  return Utilities.formatDate(date || new Date(), getTimeZone_(), 'dd/MM/yyyy HH:mm:ss');
}

/**
 * Removes control characters and angle brackets, trims and limits length.
 */
function sanitizeText_(value, multiline, maxLength) {
  var s = String(value == null ? '' : value);
  if (multiline) {
    s = s.replace(/\r\n?/g, '\n').replace(/[\u0000-\u0009\u000B-\u001F\u007F]/g, '');
    s = s.replace(/[ \t]+/g, ' ').replace(/\n{3,}/g, '\n\n');
  } else {
    s = s.replace(/[\u0000-\u001F\u007F]+/g, ' ').replace(/\s+/g, ' ');
  }
  s = s.replace(/[<>]/g, '').trim();
  if (maxLength && s.length > maxLength) s = s.slice(0, maxLength);
  return s;
}

/**
 * Prevents spreadsheet formula injection: text that starts with = + @ (or a
 * dash not followed by a digit) is stored with a leading apostrophe so
 * Google Sheets treats it as plain text.
 */
function neutralizeFormula_(s) {
  var str = String(s == null ? '' : s);
  if (/^[=+@]/.test(str) || /^-[^\d.]/.test(str)) return "'" + str;
  return str;
}

/** Masks an email or mobile for logs and messages: ra****@gmail.com / ******3210 */
function maskIdentifier_(value) {
  var s = String(value || '');
  if (s.indexOf('@') > 0) {
    var parts = s.split('@');
    var user = parts[0];
    return user.slice(0, Math.min(2, user.length)) + '****@' + parts[1];
  }
  return s.length > 4 ? new Array(s.length - 3).join('*') + s.slice(-4) : '****';
}

/** HTML-safe JSON for embedding inside a <script> element. */
function safeJsonForHtml_(obj) {
  return JSON.stringify(obj)
    .replace(/</g, '\\u003c')
    .replace(/>/g, '\\u003e')
    .replace(/&/g, '\\u0026')
    .replace(/\u2028/g, '\\u2028')
    .replace(/\u2029/g, '\\u2029');
}

/* ======================================================================== */
/* Crypto helpers                                                           */
/* ======================================================================== */

function bytesToHex_(bytes) {
  var out = [];
  for (var i = 0; i < bytes.length; i++) {
    var v = bytes[i] & 0xff;
    out.push((v < 16 ? '0' : '') + v.toString(16));
  }
  return out.join('');
}

function sha256Hex_(text) {
  return bytesToHex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(text), Utilities.Charset.UTF_8));
}

/** 64 hex chars of randomness (two v4 UUIDs). */
function randomToken_() {
  return (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '').toLowerCase();
}

/** Constant-time comparison of two hex digests. */
function constantTimeEquals_(a, b) {
  a = String(a); b = String(b);
  if (a.length !== b.length) return false;
  var diff = 0;
  for (var i = 0; i < a.length; i++) diff |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return diff === 0;
}

/** Unique, human-readable correction transaction ID, e.g. CRN-20261103-142233-7F3A9C */
function generateTransactionId_() {
  var stamp = Utilities.formatDate(new Date(), getTimeZone_(), 'yyyyMMdd-HHmmss');
  return 'CRN-' + stamp + '-' + Utilities.getUuid().replace(/-/g, '').slice(0, 6).toUpperCase();
}
