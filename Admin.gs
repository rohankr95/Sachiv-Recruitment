/**
 * Admin.gs
 * ---------------------------------------------------------------------------
 * Administrative dashboard (open the web app URL with ?page=admin).
 *
 * Admin accounts are stored in Script Properties as salted, iterated SHA-256
 * hashes – never in code and never sent to the browser. To create one:
 *   1. Project Settings > Script Properties > add
 *        NEW_ADMIN_USERNAME = <username>
 *        NEW_ADMIN_PASSWORD = <password, min 10 chars>
 *   2. Run setupAdminAccount() from the editor. The plain-text properties are
 *      deleted automatically after hashing.
 * ---------------------------------------------------------------------------
 */

var ADMIN_ACCOUNTS_PROPERTY_ = 'ADMIN_ACCOUNTS';
var PASSWORD_ITERATIONS_ = 500;

/* ======================================================================== */
/* Authentication                                                           */
/* ======================================================================== */

function adminLogin(username, password) {
  return respond_(function () {
    var user = String(username == null ? '' : username).trim().toLowerCase();
    var pass = String(password == null ? '' : password);
    if (!user || !pass || user.length > 64 || pass.length > 200) {
      throw portalError_('VALIDATION', 'Enter your username and password.');
    }
    var limiter = checkRateLimit_('admin', user, CONFIG.RATE_LIMIT.ADMIN_LOCKOUT_MINUTES);
    var accounts = getAdminAccounts_();
    var account = accounts[user];
    // Always compute a hash so response time does not reveal valid usernames.
    var hash = hashPassword_(pass, account ? account.salt : 'no-such-user-salt', account ? account.iterations : PASSWORD_ITERATIONS_);
    if (!account || !constantTimeEquals_(hash, account.hash)) {
      var remaining = registerFailure_(limiter, CONFIG.RATE_LIMIT.ADMIN_MAX_FAILED_ATTEMPTS, CONFIG.RATE_LIMIT.ADMIN_LOCKOUT_MINUTES);
      Utilities.sleep(CONFIG.RATE_LIMIT.FAILURE_DELAY_MS);
      console.warn('Failed admin login for "' + user + '"');
      throw portalError_('AUTH_FAILED', remaining > 0 ? 'Invalid username or password.' :
        'Too many failed attempts. Admin login is locked for ' + CONFIG.RATE_LIMIT.ADMIN_LOCKOUT_MINUTES + ' minutes.');
    }
    clearRateLimit_(limiter);
    var token = createSession_('admin', { user: user }, CONFIG.SESSION.ADMIN_IDLE_TIMEOUT_MINUTES);
    console.log('Admin login: ' + user);
    return { token: token, username: user, idleTimeoutSeconds: CONFIG.SESSION.ADMIN_IDLE_TIMEOUT_MINUTES * 60, requireApproval: CONFIG.REQUIRE_APPROVAL === true };
  });
}

function adminLogout(token) {
  return respond_(function () { destroySession_(token); return { loggedOut: true }; });
}

function adminKeepAlive(token) {
  return respond_(function () {
    getSession_(token, 'admin');
    return { idleTimeoutSeconds: CONFIG.SESSION.ADMIN_IDLE_TIMEOUT_MINUTES * 60 };
  });
}

/* ======================================================================== */
/* Dashboard                                                                */
/* ======================================================================== */

function adminGetDashboard(token) {
  return respond_(function () {
    getSession_(token, 'admin');
    var ctx = getSheetContext_();
    var appNos = readColumn_(ctx, ctx.cols.appNo, true);
    var counts = readColumn_(ctx, ctx.cols.count, false);
    var statuses = readColumn_(ctx, ctx.cols.status, true);
    var emails = ctx.cols.email !== -1 ? readColumn_(ctx, ctx.cols.email, false) : [];
    var dobRaw = readColumn_(ctx, ctx.cols.dob, false);
    var dobDisp = readColumn_(ctx, ctx.cols.dob, true);

    var total = 0, excluded = 0, corrected = 0, pending = 0, approved = 0, rejected = 0;
    for (var i = 0; i < appNos.length; i++) {
      if (!String(appNos[i]).trim()) continue;
      if (isExcludedRow_(dobRaw[i], dobDisp[i])) { excluded++; continue; }
      total++;
      if ((Number(counts[i]) || 0) > 0) corrected++;
      var st = String(statuses[i] || '');
      if (st === CONFIG.STATUS.PENDING) pending++;
      else if (st === CONFIG.STATUS.APPROVED) approved++;
      else if (st === CONFIG.STATUS.REJECTED) rejected++;
    }
    var emailCounts = countOccurrences_(emails, normalizeEmail_);
    var dupGroups = 0, dupRecords = 0;
    Object.keys(emailCounts).forEach(function (k) {
      if (emailCounts[k] > 1) { dupGroups++; dupRecords += emailCounts[k]; }
    });
    var log = getLogSummary_(20);
    return {
      stats: {
        totalApplications: total,
        excludedIncomplete: excluded,
        applicationsCorrected: corrected,
        pendingCorrections: pending,
        approved: approved,
        rejected: rejected,
        documentsReplaced: log.documentsReplaced,
        totalTransactions: log.transactions,
        fieldChanges: log.fieldChanges,
        duplicateEmailGroups: dupGroups,
        duplicateEmailRecords: dupRecords
      },
      approvalEnabled: CONFIG.REQUIRE_APPROVAL === true,
      portalOpen: CONFIG.PORTAL_OPEN === true,
      recent: log.recent
    };
  });
}

/**
 * Search by application number (exact), email or mobile number.
 */
function adminSearch(token, query) {
  return respond_(function () {
    getSession_(token, 'admin');
    var q = String(query == null ? '' : query).trim();
    if (!q || q.length > 254) throw portalError_('VALIDATION', 'Enter an application number, email or mobile number.');
    var ctx = getSheetContext_();
    var appNos = readColumn_(ctx, ctx.cols.appNo, true);
    var idx = [];
    var qLower = q.toLowerCase();
    appNos.forEach(function (a, i) { if (String(a).trim().toLowerCase() === qLower) idx.push(i); });
    var matchedBy = 'Application Number';
    if (!idx.length && q.indexOf('@') !== -1 && ctx.cols.email !== -1) {
      readColumn_(ctx, ctx.cols.email, false).forEach(function (e, i) { if (normalizeEmail_(e) === qLower) idx.push(i); });
      matchedBy = 'Email';
    } else if (!idx.length && /^[+\d\s-]{10,16}$/.test(q) && ctx.cols.mobile !== -1) {
      var m = normalizeMobile_(q);
      readColumn_(ctx, ctx.cols.mobile, false).forEach(function (v, i) { if (normalizeMobile_(v) === m) idx.push(i); });
      matchedBy = 'Mobile Number';
    }
    if (!idx.length) return { matchedBy: matchedBy, results: [] };
    var names = ctx.cols.name !== -1 ? readColumn_(ctx, ctx.cols.name, true) : [];
    var statuses = readColumn_(ctx, ctx.cols.status, true);
    var counts = readColumn_(ctx, ctx.cols.count, false);
    return {
      matchedBy: matchedBy,
      results: idx.slice(0, 100).map(function (i) {
        return {
          applicationNumber: String(appNos[i]),
          applicantName: names.length ? String(names[i]) : '',
          status: String(statuses[i] || '') || 'Not corrected',
          correctionCount: Number(counts[i]) || 0
        };
      })
    };
  });
}

/**
 * Full application with modified-field markers, documents (with Drive links –
 * admins are signed in as the owner), correction history and duplicates.
 */
function adminGetApplication(token, appNo) {
  return respond_(function () {
    getSession_(token, 'admin');
    appNo = cleanAppNo_(appNo);
    var ctx = getSheetContext_();
    var row = readRow_(ctx, findRowByAppNo_(ctx, appNo));
    var history = getCorrectionHistory_(appNo);

    var modified = {};
    history.forEach(function (t) { t.changes.forEach(function (c) { modified[c.field] = (modified[c.field] || 0) + 1; }); });

    var sectionsByName = {};
    var documents = [];
    ctx.fields.forEach(function (f) {
      var raw = row.values[f.colIndex];
      var disp = row.display[f.colIndex];
      if (f.type === 'document') {
        var ids = extractDriveFileIds_(raw, row.formulas[f.colIndex], disp);
        documents.push({
          label: f.label,
          header: f.header,
          timesReplaced: modified[f.header] || 0,
          files: ids.map(function (id) {
            var info = getDriveFileInfo_(id);
            return { id: id, name: info.name, available: info.available, url: 'https://drive.google.com/file/d/' + id + '/view' };
          }),
          rawValue: ids.length ? '' : String(disp || '')
        });
        return;
      }
      var section = f.type === 'hidden' ? 'Portal System Fields' : f.section;
      if (!sectionsByName[section]) sectionsByName[section] = [];
      sectionsByName[section].push({
        label: f.label,
        header: f.header,
        type: f.type,
        editable: f.editable,
        value: cellToDisplayValue_(f, raw, disp),
        timesModified: modified[f.header] || 0,
        order: f.order,
        col: f.colIndex
      });
    });
    var order = getSectionOrder_();
    Object.keys(sectionsByName).forEach(function (n) { if (order.indexOf(n) === -1) order.push(n); });
    var sections = order.filter(function (n) { return sectionsByName[n]; }).map(function (n) {
      return { name: n, icon: sectionIcon_(n), fields: sectionsByName[n].sort(function (a, b) { return (a.order - b.order) || (a.col - b.col); }) };
    });

    // Duplicate information.
    var dup = { email: [], mobile: [] };
    var appNos = readColumn_(ctx, ctx.cols.appNo, true);
    if (ctx.cols.email !== -1 && row.values[ctx.cols.email]) {
      var em = normalizeEmail_(row.values[ctx.cols.email]);
      readColumn_(ctx, ctx.cols.email, false).forEach(function (v, i) {
        if (normalizeEmail_(v) === em && String(appNos[i]) !== appNo) dup.email.push(String(appNos[i]));
      });
    }
    if (ctx.cols.mobile !== -1 && row.values[ctx.cols.mobile]) {
      var mo = normalizeMobile_(row.values[ctx.cols.mobile]);
      readColumn_(ctx, ctx.cols.mobile, false).forEach(function (v, i) {
        if (mo && normalizeMobile_(v) === mo && String(appNos[i]) !== appNo) dup.mobile.push(String(appNos[i]));
      });
    }

    var status = String(row.display[ctx.cols.status] || '');
    var lastTxn = String(row.display[ctx.cols.txn] || '');
    return {
      applicationNumber: appNo,
      applicantName: ctx.cols.name !== -1 ? String(row.display[ctx.cols.name] || '') : '',
      status: status || 'Not corrected',
      correctionCount: Number(row.values[ctx.cols.count]) || 0,
      lastModified: String(row.display[ctx.cols.lastModified] || ''),
      lastTransactionId: lastTxn,
      excluded: isExcludedRow_(row.values[ctx.cols.dob], row.display[ctx.cols.dob]),
      canDecide: CONFIG.REQUIRE_APPROVAL === true && status === CONFIG.STATUS.PENDING && !!lastTxn,
      sections: sections,
      documents: documents,
      modifiedFields: Object.keys(modified),
      history: history,
      duplicates: dup
    };
  });
}

/** Streams a document for the admin viewer – only files linked to this application. */
function adminViewDocument(token, appNo, fileId) {
  return respond_(function () {
    getSession_(token, 'admin');
    appNo = cleanAppNo_(appNo);
    fileId = String(fileId || '');
    if (!/^[A-Za-z0-9_-]{20,}$/.test(fileId)) throw portalError_('VALIDATION', 'Invalid document reference.');
    var ctx = getSheetContext_();
    var row = readRow_(ctx, findRowByAppNo_(ctx, appNo));
    var allowed = false;
    ctx.fields.forEach(function (f) {
      if (f.type !== 'document' || allowed) return;
      if (extractDriveFileIds_(row.values[f.colIndex], row.formulas[f.colIndex], row.display[f.colIndex]).indexOf(fileId) !== -1) allowed = true;
    });
    if (!allowed) {
      allowed = getCorrectionHistory_(appNo).some(function (t) {
        return t.changes.some(function (c) {
          return (c.previousValue + ' ' + c.updatedValue + ' ' + c.documentDetails).indexOf(fileId) !== -1;
        });
      });
    }
    if (!allowed) throw portalError_('FORBIDDEN', 'This document is not linked to the selected application.');
    return readFileForViewer_(fileId);
  });
}

/**
 * Approve or reject the latest correction (only when REQUIRE_APPROVAL is on).
 * Rejecting restores the previous values of that correction.
 * @param {string} decision 'APPROVE' | 'REJECT'
 */
function adminSetStatus(token, appNo, txnId, decision, remark) {
  return respond_(function () {
    var session = getSession_(token, 'admin');
    if (CONFIG.REQUIRE_APPROVAL !== true) throw portalError_('NOT_ALLOWED', 'Approval workflow is disabled.');
    appNo = cleanAppNo_(appNo);
    txnId = String(txnId || '');
    decision = String(decision || '').toUpperCase();
    if (decision !== 'APPROVE' && decision !== 'REJECT') throw portalError_('VALIDATION', 'Invalid decision.');
    var note = sanitizeText_(remark, false, 300);

    var lock = LockService.getScriptLock();
    if (!lock.tryLock(CONFIG.LOCK_TIMEOUT_MS)) throw portalError_('BUSY', 'The server is busy. Please try again.');
    try {
      var ctx = getSheetContext_();
      var rowIndex = findRowByAppNo_(ctx, appNo);
      var row = readRow_(ctx, rowIndex);
      if (String(row.display[ctx.cols.txn]) !== txnId) throw portalError_('CONFLICT', 'Only the latest correction of an application can be approved or rejected.');
      if (String(row.display[ctx.cols.status]) !== CONFIG.STATUS.PENDING) throw portalError_('CONFLICT', 'This correction has already been decided.');

      var newStatus = decision === 'APPROVE' ? CONFIG.STATUS.APPROVED : CONFIG.STATUS.REJECTED;
      var reverted = [], skipped = [];
      if (decision === 'REJECT') {
        getTransactionRows_(txnId).forEach(function (r) {
          if (String(r[2]) === LOG_STATUS_FIELD_) return;
          var field = null;
          ctx.fields.forEach(function (f) { if (!field && f.header === String(r[2])) field = f; });
          if (!field) { skipped.push(String(r[2]) + ' (column no longer exists)'); return; }
          var raw = row.values[field.colIndex];
          var disp = row.display[field.colIndex];
          var current = field.type === 'document' ? (row.formulas[field.colIndex] || String(raw)) : cellToDisplayValue_(field, raw, disp);
          if (String(current) !== String(r[4])) { skipped.push(field.label + ' (changed again since)'); return; }
          var previous = String(r[3]);
          var cell = ctx.sheet.getRange(rowIndex, field.colIndex + 1);
          if (field.type === 'document') {
            if (previous.charAt(0) === '=') cell.setFormula(previous); else cell.setValue(previous);
          } else if (field.type === 'date') {
            writeCell_(ctx.sheet, rowIndex, field, displayDateToIso_(previous) || '', raw);
          } else {
            writeCell_(ctx.sheet, rowIndex, field, previous, raw);
          }
          reverted.push(field.label);
        });
      }
      ctx.sheet.getRange(rowIndex, ctx.cols.status + 1).setValue(newStatus);
      ctx.sheet.getRange(rowIndex, ctx.cols.lastModified + 1).setValue(new Date());
      var details = (decision === 'REJECT' ? 'Rejected; reverted: ' + (reverted.join(', ') || 'none') + (skipped.length ? '; not reverted: ' + skipped.join(', ') : '') : 'Approved') +
        (note ? '; remark: ' + note : '');
      appendLogRows_([[txnId, appNo, LOG_STATUS_FIELD_, CONFIG.STATUS.PENDING, newStatus, new Date(), details, newStatus, 'Admin (' + session.user + ')']]);
      SpreadsheetApp.flush();
      return { status: newStatus, reverted: reverted, skipped: skipped };
    } finally {
      lock.releaseLock();
    }
  });
}

/** Groups of applications that share an email address (and mobile number). */
function adminGetDuplicates(token) {
  return respond_(function () {
    getSession_(token, 'admin');
    var ctx = getSheetContext_();
    var appNos = readColumn_(ctx, ctx.cols.appNo, true);
    var names = ctx.cols.name !== -1 ? readColumn_(ctx, ctx.cols.name, true) : [];
    var dobs = readColumn_(ctx, ctx.cols.dob, true);
    var build = function (col, normalizer, label) {
      if (col === -1) return [];
      var values = readColumn_(ctx, col, false);
      var groups = {};
      values.forEach(function (v, i) {
        var k = normalizer(v);
        if (!k || !String(appNos[i]).trim()) return;
        (groups[k] = groups[k] || []).push({ applicationNumber: String(appNos[i]), applicantName: names.length ? String(names[i]) : '', dob: String(dobs[i] || '') });
      });
      return Object.keys(groups).filter(function (k) { return groups[k].length > 1; })
        .sort(function (a, b) { return groups[b].length - groups[a].length; })
        .slice(0, 300)
        .map(function (k) {
          var distinctDob = {};
          groups[k].forEach(function (a) { distinctDob[a.dob] = true; });
          return { type: label, value: k, count: groups[k].length, differentPersons: Object.keys(distinctDob).length > 1, applications: groups[k] };
        });
    };
    var appNoGroups = {};
    appNos.forEach(function (a) { var k = String(a).trim(); if (k) appNoGroups[k] = (appNoGroups[k] || 0) + 1; });
    return {
      email: build(ctx.cols.email, normalizeEmail_, 'Email'),
      mobile: build(ctx.cols.mobile, normalizeMobile_, 'Mobile'),
      duplicateApplicationNumbers: Object.keys(appNoGroups).filter(function (k) { return appNoGroups[k] > 1; })
    };
  });
}

/** Clears cached Field_Config so edits to that tab take effect immediately. */
function adminRefreshConfig(token) {
  return respond_(function () {
    getSession_(token, 'admin');
    clearFieldConfigCache_();
    var ctx = getSheetContext_();
    return {
      fields: ctx.fields.length,
      documents: ctx.fields.filter(function (f) { return f.type === 'document'; }).length,
      editable: ctx.fields.filter(function (f) { return f.editable; }).length
    };
  });
}

/* ======================================================================== */
/* Account management (run from the Apps Script editor only)                */
/* ======================================================================== */

/**
 * Creates or updates an admin account from the Script Properties
 * NEW_ADMIN_USERNAME / NEW_ADMIN_PASSWORD, then deletes those properties.
 */
function setupAdminAccount() {
  requireEditorContext_();
  var props = PropertiesService.getScriptProperties();
  var user = String(props.getProperty('NEW_ADMIN_USERNAME') || '').trim().toLowerCase();
  var pass = String(props.getProperty('NEW_ADMIN_PASSWORD') || '');
  if (!user || !/^[a-z0-9._-]{3,64}$/.test(user)) throw new Error('Set Script Property NEW_ADMIN_USERNAME (3-64 chars: letters, digits, . _ -).');
  if (pass.length < 10) throw new Error('Set Script Property NEW_ADMIN_PASSWORD (at least 10 characters).');
  var accounts = getAdminAccounts_();
  var salt = randomToken_().slice(0, 32);
  accounts[user] = { salt: salt, hash: hashPassword_(pass, salt, PASSWORD_ITERATIONS_), iterations: PASSWORD_ITERATIONS_, updated: new Date().toISOString() };
  props.setProperty(ADMIN_ACCOUNTS_PROPERTY_, JSON.stringify(accounts));
  props.deleteProperty('NEW_ADMIN_PASSWORD');
  props.deleteProperty('NEW_ADMIN_USERNAME');
  console.log('Admin account "' + user + '" saved. Plain-text properties removed. Accounts: ' + Object.keys(accounts).join(', '));
}

/** Removes the admin named in Script Property REMOVE_ADMIN_USERNAME. */
function removeAdminAccount() {
  requireEditorContext_();
  var props = PropertiesService.getScriptProperties();
  var user = String(props.getProperty('REMOVE_ADMIN_USERNAME') || '').trim().toLowerCase();
  var accounts = getAdminAccounts_();
  if (!user || !accounts[user]) throw new Error('Set Script Property REMOVE_ADMIN_USERNAME to an existing admin username.');
  delete accounts[user];
  props.setProperty(ADMIN_ACCOUNTS_PROPERTY_, JSON.stringify(accounts));
  props.deleteProperty('REMOVE_ADMIN_USERNAME');
  console.log('Admin account "' + user + '" removed. Remaining: ' + (Object.keys(accounts).join(', ') || 'none'));
}

function getAdminAccounts_() {
  var raw = PropertiesService.getScriptProperties().getProperty(ADMIN_ACCOUNTS_PROPERTY_);
  if (!raw) return {};
  try { return JSON.parse(raw) || {}; } catch (e) { return {}; }
}

function hashPassword_(password, salt, iterations) {
  var h = sha256Hex_(salt + ':' + password);
  for (var i = 1; i < (iterations || PASSWORD_ITERATIONS_); i++) h = sha256Hex_(h + ':' + salt);
  return h;
}

/** Converts a value formatted with CONFIG.DATE.DISPLAY_FORMAT back to yyyy-MM-dd. */
function displayDateToIso_(text) {
  var s = String(text || '').trim();
  if (!s) return '';
  var m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})$/);
  if (m) {
    return CONFIG.DATE.DISPLAY_FORMAT.indexOf('MM') === 0 ? buildIsoDate_(m[3], m[1], m[2]) : buildIsoDate_(m[3], m[2], m[1]);
  }
  return toIsoDate_(s);
}
