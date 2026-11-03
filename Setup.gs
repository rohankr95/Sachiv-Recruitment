/**
 * Setup.gs
 * ---------------------------------------------------------------------------
 * One-time (and re-runnable) setup. Run setupPortal() from the Apps Script
 * editor after filling in Config.gs. It is idempotent: running it again only
 * adds what is missing.
 *
 * What it does
 *  1. Verifies the key columns (Application Number, DOB, Email/Mobile) exist.
 *  2. Appends the 4 portal system columns to the application sheet if absent
 *     (Last Modified, Correction Count, Correction Status, Last Correction ID).
 *  3. Creates the Field_Config tab with auto-detected settings for every
 *     column (document columns are detected from Drive links in the data).
 *  4. Creates the protected, append-only Correction_Log tab.
 *  5. Creates private Drive folders for uploads / archive / staging.
 *  6. Installs a daily trigger that clears abandoned staged uploads.
 *  7. Reports sharing problems (public spreadsheet / folders).
 * ---------------------------------------------------------------------------
 */

function setupPortal() {
  requireEditorContext_();
  var report = [];
  var ss = getSpreadsheet_();
  var sheet = getApplicationSheet_();
  var headers = readHeaders_(sheet);

  // 1. Header sanity.
  var seen = {};
  headers.forEach(function (h, i) {
    var k = normalizeHeader_(h);
    if (!k) return;
    if (seen[k]) report.push('WARNING: duplicate header "' + h + '" in columns ' + seen[k] + ' and ' + (i + 1) + '. Rename one of them.');
    seen[k] = i + 1;
  });
  var cols = resolveKeyColumns_(headers, false);
  var missing = [];
  if (cols.appNo === -1) missing.push('Application Number (aliases: ' + CONFIG.KEY_COLUMNS.APPLICATION_NUMBER.join(', ') + ')');
  if (cols.dob === -1) missing.push('Date of Birth (aliases: ' + CONFIG.KEY_COLUMNS.DOB.join(', ') + ')');
  if (cols.email === -1 && cols.mobile === -1) missing.push('Email or Mobile Number');
  if (missing.length) {
    throw new Error('Key column(s) not found in "' + CONFIG.APPLICATION_SHEET_NAME + '": ' + missing.join('; ') +
      '. Add your header text to CONFIG.KEY_COLUMNS in Config.gs.');
  }
  if (cols.email === -1) report.push('NOTE: no email column found – applicants can log in with mobile number only.');
  if (cols.mobile === -1) report.push('NOTE: no mobile column found – applicants can log in with email only.');
  if (cols.name === -1) report.push('NOTE: no applicant name column found – names will not be shown in lists.');

  // 2. System columns (the only structural change to the application sheet).
  var toAdd = [];
  Object.keys(CONFIG.SYSTEM_COLUMNS).forEach(function (k) {
    if (findColumn_(headers, CONFIG.SYSTEM_COLUMNS[k]) === -1) toAdd.push(CONFIG.SYSTEM_COLUMNS[k]);
  });
  if (toAdd.length) {
    var startCol = headers.length + 1;
    var needed = startCol + toAdd.length - 1 - sheet.getMaxColumns();
    if (needed > 0) sheet.insertColumnsAfter(sheet.getMaxColumns(), needed);
    sheet.getRange(CONFIG.HEADER_ROW, startCol, 1, toAdd.length).setValues([toAdd]).setFontWeight('bold');
    report.push('Added system column(s): ' + toAdd.join(', '));
    headers = readHeaders_(sheet);
  }
  cols = resolveKeyColumns_(headers, true);
  var lastModCol = sheet.getRange(CONFIG.HEADER_ROW + 1, cols.lastModified + 1, Math.max(1, sheet.getMaxRows() - CONFIG.HEADER_ROW), 1);
  lastModCol.setNumberFormat('dd/MM/yyyy HH:mm:ss');

  // Data integrity: duplicate application numbers block corrections for those rows.
  var lastRow = sheet.getLastRow();
  if (lastRow > CONFIG.HEADER_ROW) {
    var appNos = sheet.getRange(CONFIG.HEADER_ROW + 1, cols.appNo + 1, lastRow - CONFIG.HEADER_ROW, 1).getDisplayValues();
    var counts = {};
    var blanks = 0;
    appNos.forEach(function (r) { var a = String(r[0]).trim(); if (!a) { blanks++; return; } counts[a] = (counts[a] || 0) + 1; });
    var dups = Object.keys(counts).filter(function (a) { return counts[a] > 1; });
    if (dups.length) report.push('WARNING: duplicate application numbers (corrections blocked for these): ' + dups.slice(0, 20).join(', ') + (dups.length > 20 ? ' …' : ''));
    if (blanks) report.push('NOTE: ' + blanks + ' row(s) have no application number and cannot be corrected.');
    report.push('Records found: ' + (lastRow - CONFIG.HEADER_ROW));
  }

  // 3-6.
  setupFieldConfigSheet_(ss, sheet, headers, cols, report);
  setupLogSheet_(report);
  setupFolders_(report);
  ensureCleanupTrigger_(report);
  clearFieldConfigCache_();

  // 7. Sharing checks.
  checkPrivate_(ss.getId(), 'Spreadsheet', report, false);
  ['UPLOAD', 'ARCHIVE', 'STAGING'].forEach(function (k) { checkPrivate_(getFolder_(k).getId(), k + ' folder', report, true); });

  var text = 'Setup complete.\n' + report.map(function (r) { return ' • ' + r; }).join('\n');
  console.log(text);
  return text;
}

/** Creates Field_Config (or appends rows for new columns). */
function setupFieldConfigSheet_(ss, appSheet, headers, cols, report) {
  var cfgSheet = ss.getSheetByName(CONFIG.FIELD_CONFIG_SHEET_NAME);
  var existing = {};
  if (!cfgSheet) {
    cfgSheet = ss.insertSheet(CONFIG.FIELD_CONFIG_SHEET_NAME);
    cfgSheet.getRange(1, 1, 1, FIELD_CONFIG_HEADERS.length).setValues([FIELD_CONFIG_HEADERS]).setFontWeight('bold').setBackground('#e8eef7');
    cfgSheet.setFrozenRows(1);
    report.push('Created "' + CONFIG.FIELD_CONFIG_SHEET_NAME + '" tab.');
  } else if (cfgSheet.getLastRow() > 1) {
    cfgSheet.getRange(2, 1, cfgSheet.getLastRow() - 1, 1).getDisplayValues().forEach(function (r) { existing[normalizeHeader_(r[0])] = true; });
  }

  // Read up to 5,000 rows in one call to detect document / date / select columns and dropdown options.
  var dataRows = appSheet.getLastRow() - CONFIG.HEADER_ROW;
  var sample = dataRows > 0 ? appSheet.getRange(CONFIG.HEADER_ROW + 1, 1, Math.min(5000, dataRows), headers.length).getValues() : [];
  var appNoIsFormula = dataRows > 0 && appSheet.getRange(CONFIG.HEADER_ROW + 1, cols.appNo + 1, Math.min(50, dataRows), 1)
    .getFormulas().some(function (r) { return !!r[0]; });
  if (appNoIsFormula && cols.email !== -1) {
    report.push('NOTE: the Application Number column contains formulas. Email is set to non-editable in Field_Config ' +
      'because changing it could change the application number.');
  }

  var rows = [];
  headers.forEach(function (h, i) {
    if (!h || existing[normalizeHeader_(h)]) return;
    var samples = sample.map(function (r) { return r[i]; });
    var d = detectField_(h, samples, cols, i);
    var opts = d.options.some(function (o) { return o.indexOf(',') !== -1; }) ? d.options.join(' | ') : d.options.join(', ');
    var editable = d.editable;
    var help = d.help || '';
    if (appNoIsFormula && i === cols.email) {
      // The application number is calculated from the email, so changing the
      // email would change (or break) the application number.
      editable = false;
      help = 'Email cannot be changed here because the application number is linked to it.';
    }
    rows.push([h, d.label, d.section, d.type, d.required ? 'Yes' : 'No', editable ? 'Yes' : 'No', opts, d.pattern || '', '', help, i + 1]);
  });
  if (rows.length) {
    var start = Math.max(2, cfgSheet.getLastRow() + 1);
    cfgSheet.getRange(start, 1, rows.length, FIELD_CONFIG_HEADERS.length).setValues(rows);
    report.push('Field_Config: added ' + rows.length + ' column(s). Detected documents: ' +
      (rows.filter(function (r) { return r[3] === 'document'; }).map(function (r) { return r[0]; }).join(', ') || 'none') + '.');
  } else {
    report.push('Field_Config: no new columns.');
  }

  var n = Math.max(1, cfgSheet.getMaxRows() - 1);
  var sectionNames = CONFIG.SECTIONS.map(function (s) { return s.name; });
  cfgSheet.getRange(2, 3, n, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(sectionNames, true).setAllowInvalid(true).build());
  cfgSheet.getRange(2, 4, n, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(FIELD_TYPES, true).setAllowInvalid(false).build());
  cfgSheet.getRange(2, 5, n, 2).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Yes', 'No'], true).setAllowInvalid(false).build());
  cfgSheet.getRange(1, 1).setNote('Edit this tab to control how each column appears in the portal. ' +
    'Field Type: ' + FIELD_TYPES.join(', ') + '. "hidden" columns are never shown to applicants; "readonly" are shown but locked. ' +
    'Options: comma separated (use | if an option contains a comma). Changes apply within ' + Math.round(CONFIG.FIELD_CONFIG_CACHE_SECONDS / 60) +
    ' minutes, or immediately via "Refresh field configuration" on the admin dashboard.');
  try { cfgSheet.autoResizeColumns(1, 6); } catch (e) { /* cosmetic */ }
}

/** Creates the protected Correction_Log tab. */
function setupLogSheet_(report) {
  var ss = getLogSpreadsheet_();
  var log = ss.getSheetByName(CONFIG.LOG_SHEET_NAME);
  if (!log) {
    log = ss.insertSheet(CONFIG.LOG_SHEET_NAME);
    log.getRange(1, 1, 1, LOG_HEADERS.length).setValues([LOG_HEADERS]).setFontWeight('bold').setBackground('#fdf2e9');
    log.setFrozenRows(1);
    if (log.getMaxRows() < 1000) log.insertRowsAfter(log.getMaxRows(), 1000 - log.getMaxRows());
    report.push('Created "' + CONFIG.LOG_SHEET_NAME + '" tab.');
  }
  var rows = Math.max(1, log.getMaxRows() - 1);
  // Plain-text format so previous/updated values are stored exactly as typed.
  log.getRange(2, 1, rows, 5).setNumberFormat('@');
  log.getRange(2, 6, rows, 1).setNumberFormat('dd/MM/yyyy HH:mm:ss');
  log.getRange(2, 7, rows, 3).setNumberFormat('@');

  var protections = log.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  if (!protections.length) {
    var p = log.protect().setDescription('Correction log – append-only, written by the portal');
    var me = Session.getEffectiveUser();
    p.addEditor(me);
    p.removeEditors(p.getEditors().filter(function (u) { return u.getEmail() !== me.getEmail(); }));
    if (p.canDomainEdit()) p.setDomainEdit(false);
    report.push('Protected "' + CONFIG.LOG_SHEET_NAME + '" (only the owner can edit manually).');
  }
}

/** Creates (or validates) the Drive folders and stores their IDs. */
function setupFolders_(report) {
  var props = PropertiesService.getScriptProperties();
  var names = { UPLOAD: 'Replacement Uploads', ARCHIVE: 'Archived Documents', STAGING: 'Staging (temporary uploads)' };
  var root = null;
  Object.keys(names).forEach(function (kind) {
    var configured = CONFIG.DRIVE[kind + '_FOLDER_ID'];
    if (configured) {
      DriveApp.getFolderById(configured); // throws if inaccessible
      report.push(kind + ' folder: using ID from Config.gs.');
      return;
    }
    var stored = props.getProperty('FOLDER_' + kind);
    if (stored) {
      try { DriveApp.getFolderById(stored); report.push(kind + ' folder: ' + stored); return; } catch (e) { /* recreate */ }
    }
    if (!root) {
      var rootId = props.getProperty('FOLDER_ROOT');
      try { root = rootId ? DriveApp.getFolderById(rootId) : null; } catch (e) { root = null; }
      if (!root) {
        root = DriveApp.createFolder(CONFIG.PORTAL_TITLE + ' – Documents');
        props.setProperty('FOLDER_ROOT', root.getId());
      }
    }
    var folder = root.createFolder(names[kind]);
    props.setProperty('FOLDER_' + kind, folder.getId());
    report.push('Created ' + kind + ' folder "' + names[kind] + '" (ID ' + folder.getId() + ').');
  });
  RUNTIME_.folders = {};
}

function ensureCleanupTrigger_(report) {
  var exists = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'cleanupStagingFolder'; });
  if (!exists) {
    ScriptApp.newTrigger('cleanupStagingFolder').timeBased().everyDays(1).atHour(2).create();
    report.push('Installed daily trigger for cleanupStagingFolder (02:00).');
  }
}

function checkPrivate_(id, label, report, isFolder) {
  try {
    var item = isFolder ? DriveApp.getFolderById(id) : DriveApp.getFileById(id);
    var access = item.getSharingAccess();
    if (access === DriveApp.Access.ANYONE || access === DriveApp.Access.ANYONE_WITH_LINK) {
      report.push('SECURITY WARNING: ' + label + ' is shared with "' + access + '". Change General access to "Restricted".');
    } else {
      report.push(label + ' sharing: ' + access + ' (OK).');
    }
  } catch (e) {
    report.push('Could not verify sharing for ' + label + ': ' + e.message);
  }
}

/**
 * Optional helper: prints a configuration summary to the execution log
 * (useful after editing Field_Config).
 */
function showFieldMapping() {
  requireEditorContext_();
  clearFieldConfigCache_();
  var ctx = getSheetContext_();
  ctx.fields.forEach(function (f) {
    console.log((f.colIndex + 1) + '. ' + f.header + ' -> ' + f.section + ' / ' + f.type +
      (f.required ? ' / required' : '') + (f.editable ? ' / editable' : ' / locked') +
      (f.options.length ? ' / options: ' + f.options.join(', ') : ''));
  });
}
