/**
 * Application.gs
 * ---------------------------------------------------------------------------
 * Dynamic field mapping, application retrieval and correction submission.
 *
 * Field mapping
 * -------------
 * Fields are generated from the sheet's header row. The optional
 * Field_Config tab (created by setupPortal) overrides the label, section,
 * type, required/editable flags, dropdown options, regex and max length for
 * any column. Columns that are not listed in Field_Config still appear, using
 * auto-detected settings – nothing is hard-coded per column.
 * ---------------------------------------------------------------------------
 */

var FIELD_CONFIG_CACHE_KEY_ = 'FIELD_CONFIG_ROWS_V1';

/* ======================================================================== */
/* Public API                                                               */
/* ======================================================================== */

/**
 * Returns one application (only if it belongs to the session) ready to be
 * rendered as an editable form.
 */
function loadApplication(token, appNo) {
  return respond_(function () {
    var session = getSession_(token, 'applicant');
    appNo = cleanAppNo_(appNo);
    requireApplicationAccess_(session, appNo);
    var ctx = getSheetContext_();
    var row = readRow_(ctx, findRowByAppNo_(ctx, appNo));
    return buildApplicationView_(ctx, row, token, appNo, session);
  });
}

/**
 * Final submission.
 * @param {string} token
 * @param {string} appNo
 * @param {{version:string, changes:Object<string,string>, replacedDocuments:string[], declaration:boolean}} payload
 */
function submitCorrection(token, appNo, payload) {
  return respond_(function () {
    var session = getSession_(token, 'applicant');
    appNo = cleanAppNo_(appNo);
    requireApplicationAccess_(session, appNo);
    if (CONFIG.PORTAL_OPEN !== true) throw portalError_('PORTAL_CLOSED', CONFIG.PORTAL_CLOSED_MESSAGE);
    if (!payload || typeof payload !== 'object') throw portalError_('VALIDATION', 'Invalid submission.');
    if (payload.declaration !== true) throw portalError_('VALIDATION', 'Please accept the declaration before submitting.');

    var changes = (payload.changes && typeof payload.changes === 'object' && !Array.isArray(payload.changes)) ? payload.changes : {};
    var replaced = Array.isArray(payload.replacedDocuments) ? payload.replacedDocuments.map(String) : [];
    if (Object.keys(changes).length > 500 || replaced.length > 100) throw portalError_('VALIDATION', 'Invalid submission.');

    var lock = LockService.getScriptLock();
    if (!lock.tryLock(CONFIG.LOCK_TIMEOUT_MS)) {
      throw portalError_('BUSY', 'The server is busy processing other submissions. Please wait a moment and click Submit again.');
    }
    try {
      return applyCorrection_(session, token, appNo, String(payload.version || ''), changes, replaced);
    } finally {
      lock.releaseLock();
    }
  });
}

/* ======================================================================== */
/* Field configuration                                                      */
/* ======================================================================== */

/** Reads and caches the Field_Config tab (non-sensitive configuration). */
function readFieldConfigSheet_() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get(FIELD_CONFIG_CACHE_KEY_);
  if (cached) {
    try { return JSON.parse(cached); } catch (e) { /* rebuild */ }
  }
  var rows = {};
  var sheet = getSpreadsheet_().getSheetByName(CONFIG.FIELD_CONFIG_SHEET_NAME);
  if (sheet && sheet.getLastRow() > 1) {
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, FIELD_CONFIG_HEADERS.length).getDisplayValues();
    data.forEach(function (r) {
      var key = normalizeHeader_(r[0]);
      if (!key) return;
      rows[key] = {
        label: String(r[1] || '').trim(),
        section: String(r[2] || '').trim(),
        type: String(r[3] || '').trim().toLowerCase(),
        required: String(r[4] || '').trim(),
        editable: String(r[5] || '').trim(),
        options: String(r[6] || '').trim(),
        pattern: String(r[7] || '').trim(),
        maxLength: String(r[8] || '').trim(),
        help: String(r[9] || '').trim(),
        order: String(r[10] || '').trim()
      };
    });
  }
  try { cache.put(FIELD_CONFIG_CACHE_KEY_, JSON.stringify(rows), CONFIG.FIELD_CONFIG_CACHE_SECONDS); } catch (e) { /* >100KB – skip caching */ }
  return rows;
}

function clearFieldConfigCache_() {
  CacheService.getScriptCache().remove(FIELD_CONFIG_CACHE_KEY_);
}

function isYes_(value, fallback) {
  var s = String(value == null ? '' : value).trim().toLowerCase();
  if (!s) return fallback;
  return s === 'yes' || s === 'y' || s === 'true' || s === '1';
}

function parseOptions_(text) {
  var s = String(text || '').trim();
  if (!s) return [];
  var parts = s.indexOf('|') !== -1 ? s.split('|') : s.split(',');
  var seen = {};
  return parts.map(function (p) { return p.trim(); }).filter(function (p) {
    if (!p || seen[p.toLowerCase()]) return false;
    seen[p.toLowerCase()] = true;
    return true;
  });
}

/**
 * Merges header row + Field_Config overrides + auto-detection into the final
 * field list. Security-relevant rules are enforced here regardless of what
 * Field_Config says (application number and system columns are never
 * editable).
 */
function buildFieldConfig_(headers, cols) {
  var overrides = readFieldConfigSheet_();
  var systemIdx = [cols.lastModified, cols.count, cols.status, cols.txn];
  var usedKeys = {};
  var fields = [];
  var byKey = {};

  headers.forEach(function (header, i) {
    if (!header) return; // ignore columns without a header
    var key = header;
    if (usedKeys[key.toLowerCase()]) key = header + ' (column ' + (i + 1) + ')';
    usedKeys[key.toLowerCase()] = true;

    var auto = detectField_(header, [], cols, i);
    var o = overrides[normalizeHeader_(header)] || null;
    var f = {
      key: key,
      header: header,
      colIndex: i,
      role: auto.role,
      label: auto.label,
      section: auto.section,
      type: auto.type,
      required: auto.required,
      editable: auto.editable,
      options: auto.options,
      pattern: '',
      maxLength: 0,
      help: '',
      order: i + 1
    };
    if (o) {
      if (o.label) f.label = o.label;
      if (o.section) f.section = o.section;
      if (o.type && FIELD_TYPES.indexOf(o.type) !== -1) f.type = o.type;
      f.required = isYes_(o.required, f.required);
      f.editable = isYes_(o.editable, f.editable);
      if (o.options) f.options = parseOptions_(o.options);
      if (o.pattern) f.pattern = o.pattern;
      if (o.maxLength && !isNaN(parseInt(o.maxLength, 10))) f.maxLength = parseInt(o.maxLength, 10);
      if (o.help) f.help = o.help;
      if (o.order && !isNaN(parseFloat(o.order))) f.order = parseFloat(o.order);
    }

    // Enforced rules.
    if (i === cols.appNo) { f.type = 'readonly'; f.editable = false; f.required = false; }
    if (systemIdx.indexOf(i) !== -1) { f.type = 'hidden'; f.editable = false; f.required = false; f.role = 'system'; f.section = CONFIG.DEFAULT_SECTION; }
    if (f.type === 'readonly' || f.type === 'hidden') { f.editable = false; }
    if (f.type === 'document') f.section = CONFIG.DOCUMENT_SECTION;
    if (f.type === 'select' && !f.options.length) f.type = 'text';
    if (f.pattern) {
      try { new RegExp(f.pattern); } catch (e) { console.warn('Invalid regex ignored for "' + header + '": ' + f.pattern); f.pattern = ''; }
    }
    if (!f.maxLength) f.maxLength = f.type === 'textarea' ? CONFIG.MAX_TEXTAREA_LENGTH : CONFIG.MAX_TEXT_LENGTH;
    if (!f.section) f.section = CONFIG.DEFAULT_SECTION;

    fields.push(f);
    byKey[key] = f;
  });
  return { fields: fields, byHeader: byKey };
}

var DOC_KEYWORDS_ = /(upload|document|certificate|photo|photograph|signature|marksheet|mark sheet|proof|scan|attachment|attach|file|copy of|pdf|image)/i;
var SELECT_KEYWORDS_ = /(gender|\bsex\b|category|caste|marital|religion|nationality|whether|are you|do you|have you|pwd|divyang|disab|\bews\b|domicile|\bboard\b|medium|stream|blood group|minority|\bbpl\b|ex.?service|yes\/no)/i;

/**
 * Infers role, type, section, flags and (for selects) options for a column.
 * @param {string} header
 * @param {Array} samples sample values from the column (used by setupPortal)
 * @param {Object} cols key column indexes
 * @param {number} index column index
 */
function detectField_(header, samples, cols, index) {
  var h = String(header).toLowerCase();
  var label = String(header).replace(/\s*\*+\s*$/, '').trim();
  var role = '';
  if (index === cols.appNo) role = 'appNo';
  else if (index === cols.email) role = 'email';
  else if (index === cols.mobile) role = 'mobile';
  else if (index === cols.dob) role = 'dob';
  else if (index === cols.name) role = 'name';
  else if ([cols.lastModified, cols.count, cols.status, cols.txn].indexOf(index) !== -1) role = 'system';

  samples = (samples || []).filter(function (v) { return v !== '' && v !== null && v !== undefined; });
  var driveLinks = samples.filter(function (v) { return /drive\.google\.com|docs\.google\.com/i.test(String(v)); }).length;
  var looksLikeDocs = samples.length ? (driveLinks / samples.length >= 0.5) : DOC_KEYWORDS_.test(h);
  var dateSamples = samples.filter(function (v) { return isDateObject_(v); }).length;

  var type = 'text';
  if (role === 'appNo') type = 'readonly';
  else if (role === 'system') type = 'hidden';
  else if (role === 'email') type = 'email';
  else if (role === 'mobile') type = 'mobile';
  else if (role === 'dob') type = 'date';
  else if (h === 'timestamp' || h === 'submission date' || h === 'submitted on') type = 'readonly';
  else if (looksLikeDocs && (driveLinks > 0 || DOC_KEYWORDS_.test(h))) type = 'document';
  else if (/e-?mail/.test(h)) type = 'email';
  else if (/(mobile|phone|contact no|contact number|whatsapp)/.test(h)) type = 'mobile';
  else if (/(date|\bdob\b|birth)/.test(h) || (samples.length && dateSamples / samples.length > 0.8)) type = 'date';
  else if (/(pin ?code|postal code|zip)/.test(h)) type = 'pincode';
  else if (/(address|remark|description|details of|reason)/.test(h)) type = 'textarea';
  else if (/(marks|percentage|\bage\b|year|score|cgpa|sgpa|experience|income)/.test(h) && samples.length &&
           samples.every(function (v) { return typeof v === 'number' || /^-?\d+(\.\d+)?$/.test(String(v).trim()); })) type = 'number';

  var options = [];
  if (type === 'text' && SELECT_KEYWORDS_.test(h) && samples.length >= 5) {
    var distinct = {};
    samples.forEach(function (v) { var s = String(v).trim(); if (s) distinct[s] = true; });
    var list = Object.keys(distinct);
    if (list.length > 0 && list.length <= 15) { type = 'select'; options = list.sort(); }
  }

  var section = CONFIG.DEFAULT_SECTION;
  if (type === 'document') section = CONFIG.DOCUMENT_SECTION;
  else if (role === 'email' || role === 'mobile' || /(e-?mail|mobile|phone|contact|whatsapp)/.test(h)) section = 'Contact Information';
  else if (/(categor|caste|reservation|\bews\b|pwd|divyang|disab|domicile|ex.?service|minority|quota|\bbpl\b|creamy)/.test(h)) section = 'Category and Reservation Details';
  else if (/(qualification|education|board|universit|marks|percentage|passing|school|college|degree|10th|12th|class|division|roll no|subject|stream|cgpa|institute|diploma|graduat|matric)/.test(h)) section = 'Educational Qualifications';
  else if (/(address|village|city|town|district|state|pin ?code|block|tehsil|post office|police station|ward|panchayat|house|street|landmark|locality|mohalla)/.test(h)) section = 'Address Details';
  else if (role === 'name' || role === 'dob' || /(name|father|mother|husband|guardian|spouse|gender|\bsex\b|birth|\bdob\b|\bage\b|marital|nationality|religion|aadhaar|aadhar|identification|blood)/.test(h)) section = 'Personal Information';

  var required = role === 'name' || role === 'dob' || role === 'email' || role === 'mobile';
  var editable = !(role === 'appNo' || role === 'system' || type === 'readonly' || type === 'hidden');

  return { role: role, label: label, type: type, section: section, required: required, editable: editable, options: options };
}

/* ======================================================================== */
/* Building the view sent to the browser                                    */
/* ======================================================================== */

/** Converts a cell into the string the form control should hold. */
function cellToFormValue_(field, raw, display) {
  if (field.type === 'date') {
    return toIsoDate_(raw) || toIsoDate_(display) || '';
  }
  if (field.type === 'mobile') {
    return typeof raw === 'number' ? raw.toFixed(0) : String(display == null ? '' : display).trim();
  }
  return String(display == null ? '' : display);
}

/** Human-friendly version of a cell (used for original values in preview and log). */
function cellToDisplayValue_(field, raw, display) {
  if (field.type === 'date') {
    var iso = toIsoDate_(raw) || toIsoDate_(display);
    return iso ? formatIsoForDisplay_(iso) : String(display || '');
  }
  if (field.type === 'mobile' && typeof raw === 'number') return raw.toFixed(0);
  return String(display == null ? '' : display);
}

/** Optimistic-concurrency token: changes whenever a correction is saved. */
function computeVersion_(ctx, row) {
  return String(row.display[ctx.cols.txn] || '') + '|' + String(row.display[ctx.cols.count] || '0');
}

function getSectionOrder_() {
  return CONFIG.SECTIONS.map(function (s) { return s.name; });
}

function sectionIcon_(name) {
  for (var i = 0; i < CONFIG.SECTIONS.length; i++) {
    if (CONFIG.SECTIONS[i].name === name) return CONFIG.SECTIONS[i].icon;
  }
  return 'fa-list';
}

function submissionBlockReason_(ctx, row) {
  if (CONFIG.PORTAL_OPEN !== true) return CONFIG.PORTAL_CLOSED_MESSAGE;
  var count = Number(row.values[ctx.cols.count]) || 0;
  if (CONFIG.MAX_CORRECTIONS_PER_APPLICATION > 0 && count >= CONFIG.MAX_CORRECTIONS_PER_APPLICATION) {
    return 'The maximum number of corrections (' + CONFIG.MAX_CORRECTIONS_PER_APPLICATION + ') has already been made for this application.';
  }
  return '';
}

function buildApplicationView_(ctx, row, token, appNo, session) {
  var sectionsByName = {};
  var documents = [];
  var docFields = ctx.fields.filter(function (f) { return f.type === 'document'; });
  var stagedMap = getStagedUploadsFor_(token, appNo, docFields);

  ctx.fields.forEach(function (f) {
    if (f.type === 'hidden') return;
    if (f.type === 'document') {
      documents.push(buildDocumentView_(f, row, stagedMap[f.key] || null));
      return;
    }
    var raw = row.values[f.colIndex];
    var disp = row.display[f.colIndex];
    var item = {
      key: f.key,
      label: f.label,
      type: f.type,
      role: f.role,
      required: f.required,
      editable: f.editable,
      options: f.options,
      pattern: f.pattern,
      maxLength: f.maxLength,
      help: f.help,
      order: f.order,
      col: f.colIndex,
      value: cellToFormValue_(f, raw, disp),
      displayValue: cellToDisplayValue_(f, raw, disp)
    };
    if (f.type === 'date' && !item.value && String(disp || '').trim()) {
      item.unparsedValue = String(disp); // existing text that is not a recognisable date
    }
    if (!sectionsByName[f.section]) sectionsByName[f.section] = [];
    sectionsByName[f.section].push(item);
  });

  var order = getSectionOrder_();
  Object.keys(sectionsByName).forEach(function (name) { if (order.indexOf(name) === -1) order.push(name); });
  var sections = order.filter(function (name) { return sectionsByName[name] && sectionsByName[name].length; }).map(function (name) {
    var list = sectionsByName[name].sort(function (a, b) { return (a.order - b.order) || (a.col - b.col); });
    list.forEach(function (x) { delete x.order; delete x.col; });
    return { name: name, icon: sectionIcon_(name), fields: list };
  });

  var dupInfo = (session.dup && session.dup[appNo]) || null;
  var blockReason = submissionBlockReason_(ctx, row);
  return {
    applicationNumber: appNo,
    applicantName: ctx.cols.name !== -1 ? String(row.display[ctx.cols.name] || '') : '',
    status: String(row.display[ctx.cols.status] || '') || 'Not corrected',
    correctionCount: Number(row.values[ctx.cols.count]) || 0,
    lastModified: String(row.display[ctx.cols.lastModified] || ''),
    version: computeVersion_(ctx, row),
    canSubmit: !blockReason,
    blockedReason: blockReason,
    duplicate: dupInfo ? {
      email: dupInfo.email > 1 ? dupInfo.email : 0,
      mobile: dupInfo.mobile > 1 ? dupInfo.mobile : 0,
      record: dupInfo.record > 1
    } : null,
    sections: sections,
    documentSectionName: CONFIG.DOCUMENT_SECTION,
    documents: documents.sort(function (a, b) { return (a.order - b.order); }).map(function (d) { delete d.order; return d; }),
    upload: {
      maxFileSizeMb: CONFIG.UPLOAD.MAX_FILE_SIZE_MB,
      extensions: CONFIG.UPLOAD.ALLOWED_EXTENSIONS
    }
  };
}

/* ======================================================================== */
/* Validation                                                               */
/* ======================================================================== */

/**
 * Sanitises raw input for a field according to its type.
 * @return {string}
 */
function sanitizeFieldInput_(field, raw) {
  if (raw === null || raw === undefined) raw = '';
  if (typeof raw !== 'string' && typeof raw !== 'number') {
    throw portalError_('VALIDATION', 'Invalid value for "' + field.label + '".');
  }
  var s = String(raw);
  if (s.length > 10000) throw portalError_('VALIDATION', 'Value too long for "' + field.label + '".');
  s = sanitizeText_(s, field.type === 'textarea', 0);
  switch (field.type) {
    case 'email': return s.toLowerCase().replace(/\s+/g, '');
    case 'mobile': return s ? (/[A-Za-z]/.test(s) ? s : normalizeMobile_(s)) : '';
    case 'pincode': return s.replace(/\s+/g, '');
    case 'number': return s.replace(/,/g, '').trim();
    case 'date': return s.trim();
    case 'select':
      // Match options case-insensitively and return the canonical option text.
      for (var i = 0; i < field.options.length; i++) {
        if (field.options[i].toLowerCase() === s.toLowerCase()) return field.options[i];
      }
      return s;
    default: return s;
  }
}

/**
 * Validates a sanitised value. Returns an error message or ''.
 * @param {Object} field
 * @param {string} value
 * @param {string} originalValue current value in the sheet (form representation)
 */
function validateFieldValue_(field, value, originalValue) {
  if (!value) return field.required ? 'This field is required.' : '';
  if (value.length > field.maxLength) return 'Maximum ' + field.maxLength + ' characters allowed.';
  switch (field.type) {
    case 'email':
      if (!isValidEmail_(value)) return 'Enter a valid email address.';
      break;
    case 'mobile':
      if (!isValidMobile_(value)) return 'Enter a valid 10-digit mobile number.';
      break;
    case 'pincode':
      if (!/^[1-9][0-9]{5}$/.test(value)) return 'Enter a valid 6-digit PIN code.';
      break;
    case 'number':
      if (!/^-?\d+(\.\d+)?$/.test(value)) return 'Enter a valid number.';
      break;
    case 'date':
      if (!/^\d{4}-\d{2}-\d{2}$/.test(value) || !buildIsoDate_(value.slice(0, 4), value.slice(5, 7), value.slice(8, 10))) return 'Enter a valid date.';
      if (value < '1900-01-01' || value > '2100-12-31') return 'Enter a valid date.';
      if (field.role === 'dob' || /birth|\bdob\b/i.test(field.label)) {
        var today = Utilities.formatDate(new Date(), getTimeZone_(), 'yyyy-MM-dd');
        if (value >= today) return 'Date of birth must be in the past.';
      }
      break;
    case 'select':
      if (field.options.length && field.options.indexOf(value) === -1 && value !== originalValue) return 'Select a valid option.';
      break;
  }
  if (field.pattern) {
    try {
      if (!new RegExp(field.pattern).test(value)) return 'The value is not in the required format.';
    } catch (e) { /* invalid pattern already reported in config */ }
  }
  return '';
}

/** Canonical form used to decide whether a value really changed. */
function comparable_(field, value) {
  var s = String(value == null ? '' : value);
  switch (field.type) {
    case 'email': return s.trim().toLowerCase();
    case 'mobile': return normalizeMobile_(s) || s.trim();
    case 'date': return toIsoDate_(s) || s.trim();
    case 'textarea': return s.replace(/\r\n?/g, '\n').replace(/[ \t]+/g, ' ').trim();
    default: return s.replace(/\s+/g, ' ').trim();
  }
}

/* ======================================================================== */
/* Submission                                                               */
/* ======================================================================== */

/** Must be called while holding the script lock. */
function applyCorrection_(session, token, appNo, clientVersion, changes, replacedKeys) {
  var ctx = getSheetContext_();
  var rowIndex = findRowByAppNo_(ctx, appNo);
  var row = readRow_(ctx, rowIndex);

  if (clientVersion !== computeVersion_(ctx, row)) {
    throw portalError_('CONFLICT', 'This application was updated after you opened it (possibly from another device or tab). Please reload the application and make your corrections again.');
  }
  var blockReason = submissionBlockReason_(ctx, row);
  if (blockReason) throw portalError_('NOT_ALLOWED', blockReason);

  var fieldErrors = {};
  var diffs = [];
  var newValues = {};

  // 1. Text / date / select fields.
  Object.keys(changes).forEach(function (key) {
    var f = ctx.byHeader[key];
    if (!f || f.type === 'hidden' || f.type === 'document') {
      fieldErrors[key] = 'Unknown or non-editable field.';
      return;
    }
    var raw = row.values[f.colIndex];
    var disp = row.display[f.colIndex];
    var original = cellToFormValue_(f, raw, disp);
    var value = sanitizeFieldInput_(f, changes[key]);
    if (comparable_(f, value) === comparable_(f, original)) return; // not actually changed
    if (!f.editable) {
      fieldErrors[key] = '"' + f.label + '" cannot be changed.';
      return;
    }
    var err = validateFieldValue_(f, value, original);
    if (err) { fieldErrors[key] = err; return; }
    newValues[key] = value;
    diffs.push({
      field: f,
      previousDisplay: cellToDisplayValue_(f, raw, disp),
      newValue: value,
      newDisplay: f.type === 'date' && value ? formatIsoForDisplay_(value) : value,
      oldRaw: raw
    });
  });

  // 2. Mandatory fields that are empty in the final record must be filled.
  ctx.fields.forEach(function (f) {
    if (!f.required || !f.editable || f.type === 'document' || f.type === 'hidden' || fieldErrors[f.key]) return;
    var finalValue = newValues.hasOwnProperty(f.key) ? newValues[f.key] : cellToFormValue_(f, row.values[f.colIndex], row.display[f.colIndex]);
    if (!String(finalValue || '').trim()) fieldErrors[f.key] = 'This field is required.';
  });

  // 3. Document replacements (must have been uploaded to staging in this session).
  var docJobs = [];
  var stagedForApp = getStagedUploadsFor_(token, appNo, ctx.fields.filter(function (f) { return f.type === 'document'; }));
  replacedKeys.forEach(function (key) {
    var f = ctx.byHeader[key];
    if (!f || f.type !== 'document' || !f.editable) { fieldErrors[key] = 'This document cannot be replaced.'; return; }
    var staged = stagedForApp[key];
    if (!staged) { fieldErrors[key] = 'The replacement upload has expired. Please upload the document again.'; return; }
    docJobs.push({ field: f, staged: staged });
  });
  ctx.fields.forEach(function (f) {
    if (f.type !== 'document' || !f.required || !f.editable || fieldErrors[f.key]) return;
    var hasExisting = extractDriveFileIds_(row.values[f.colIndex], row.formulas[f.colIndex], row.display[f.colIndex]).length > 0;
    var isReplacing = docJobs.some(function (j) { return j.field.key === f.key; });
    if (!hasExisting && !isReplacing) fieldErrors[f.key] = 'This document is mandatory. Please upload it.';
  });

  if (Object.keys(fieldErrors).length) {
    throw portalError_('VALIDATION', 'Please correct the highlighted fields and submit again.', { fieldErrors: fieldErrors });
  }
  if (!diffs.length && !docJobs.length) {
    throw portalError_('NO_CHANGES', 'No changes were found. Modify at least one field or document before submitting.');
  }

  var txnId = generateTransactionId_();
  var now = new Date();
  var status = CONFIG.REQUIRE_APPROVAL ? CONFIG.STATUS.PENDING : CONFIG.STATUS.MODIFIED;
  var changedBy = 'Applicant (' + (session.who || 'login') + ')';

  // 4. Drive work first (archive old, move new). Nothing in the sheet changes
  //    if archiving fails.
  var docResults = processDocumentReplacements_(appNo, txnId, row, docJobs);

  // 5. Write only the cells that changed (keeps formatting/formulas of all
  //    other cells untouched), then the system columns.
  var sheet = ctx.sheet;
  diffs.forEach(function (d) {
    writeCell_(sheet, rowIndex, d.field, d.newValue, d.oldRaw);
  });
  docResults.forEach(function (r) {
    sheet.getRange(rowIndex, r.field.colIndex + 1).setValue(r.newUrl);
  });
  var newCount = (Number(row.values[ctx.cols.count]) || 0) + 1;
  sheet.getRange(rowIndex, ctx.cols.lastModified + 1).setValue(now);
  sheet.getRange(rowIndex, ctx.cols.count + 1).setValue(newCount);
  sheet.getRange(rowIndex, ctx.cols.status + 1).setValue(status);
  sheet.getRange(rowIndex, ctx.cols.txn + 1).setValue(txnId);

  // 6. Audit log (one batch write).
  var logRows = diffs.map(function (d) {
    return [txnId, appNo, d.field.header, d.previousDisplay, d.newDisplay, now, '', status, changedBy];
  });
  docResults.forEach(function (r) {
    logRows.push([txnId, appNo, r.field.header, r.previousValue, r.newUrl, now, r.details, status, changedBy]);
  });
  appendLogRows_(logRows);
  SpreadsheetApp.flush();

  // 7. Clear staged references for this application.
  clearStagedUploads_(token, appNo, docJobs.map(function (j) { return j.field; }));

  return {
    transactionId: txnId,
    applicationNumber: appNo,
    submittedAt: formatTimestamp_(now),
    fieldsChanged: diffs.map(function (d) { return d.field.label; }),
    documentsReplaced: docResults.map(function (r) { return r.field.label; }),
    status: status,
    correctionCount: newCount
  };
}

/**
 * Writes a single cell, choosing the right data type so Sheets does not
 * re-interpret the value (leading zeros, long numbers, dd/mm dates ...).
 */
function writeCell_(sheet, rowIndex, field, value, oldRaw) {
  var cell = sheet.getRange(rowIndex, field.colIndex + 1);
  if (value === '') { cell.setValue(''); return; }

  if (field.type === 'date') {
    if (oldRaw === '' || oldRaw === null || isDateObject_(oldRaw) || typeof oldRaw === 'number') {
      cell.setValue(isoToSheetDate_(value));
    } else {
      cell.setNumberFormat('@');
      cell.setValue(Utilities.formatDate(isoToSheetDate_(value), getTimeZone_(), CONFIG.DATE.STORAGE_FORMAT));
    }
    return;
  }
  if (field.type === 'number') {
    if (typeof oldRaw === 'string' && oldRaw !== '') cell.setNumberFormat('@');
    cell.setValue(typeof oldRaw === 'string' && oldRaw !== '' ? value : Number(value));
    return;
  }
  if (field.type === 'mobile' && typeof oldRaw === 'number') {
    cell.setValue(Number(value));
    return;
  }
  // Text-like values: keep them as text when they could be mis-parsed.
  var risky = /^0\d/.test(value) || /^\d{11,}$/.test(value) || /^\d{1,4}[\/\-.]\d{1,2}[\/\-.]\d{1,4}$/.test(value) || /^[\d.]+%$/.test(value) ||
              (typeof oldRaw === 'string' && oldRaw !== '' && /^[\d\s.\/\-+%]+$/.test(value));
  if (risky) cell.setNumberFormat('@');
  cell.setValue(neutralizeFormula_(value));
}
