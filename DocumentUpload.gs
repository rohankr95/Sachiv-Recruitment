/**
 * DocumentUpload.gs
 * ---------------------------------------------------------------------------
 * Document viewing, staging of replacement uploads and the final
 * replace-and-archive step.
 *
 * Flow
 *  1. Applicant chooses a file  -> uploadReplacement() validates type, size
 *     and file signature, then saves it in the private STAGING folder.
 *     The sheet is NOT changed yet.
 *  2. Applicant submits         -> processDocumentReplacements_() (called by
 *     submitCorrection under LockService) archives the previous document(s)
 *     into the ARCHIVE folder, moves the staged file into the UPLOAD folder
 *     and returns the new URL that is written into the sheet.
 *  3. Abandoned staged files are trashed by cleanupStagingFolder()
 *     (daily trigger). Original documents are never deleted.
 *
 * No file is ever shared publicly: documents are streamed to the browser
 * through viewDocument() after the session has been verified.
 * ---------------------------------------------------------------------------
 */

/* ======================================================================== */
/* Public API                                                               */
/* ======================================================================== */

/**
 * Stages a replacement document.
 * @param {string} token
 * @param {string} appNo
 * @param {string} fieldKey  document column key
 * @param {{fileName:string, mimeType:string, data:string}} upload base64 data (no data: prefix)
 */
function uploadReplacement(token, appNo, fieldKey, upload) {
  return respond_(function () {
    var session = getSession_(token, 'applicant');
    appNo = cleanAppNo_(appNo);
    requireApplicationAccess_(session, appNo);
    if (CONFIG.PORTAL_OPEN !== true) throw portalError_('PORTAL_CLOSED', CONFIG.PORTAL_CLOSED_MESSAGE);
    if (!upload || typeof upload !== 'object') throw portalError_('VALIDATION', 'No file received.');

    var ctx = getSheetContext_();
    var field = ctx.byHeader[String(fieldKey)];
    if (!field || field.type !== 'document' || !field.editable) {
      throw portalError_('FORBIDDEN', 'This document cannot be replaced.');
    }
    if ((session.uploads || 0) >= CONFIG.UPLOAD.MAX_UPLOADS_PER_SESSION) {
      throw portalError_('LIMIT', 'Upload limit for this session reached. Please log out and log in again.');
    }

    var originalName = sanitizeFileName_(upload.fileName);
    var ext = (originalName.split('.').pop() || '').toLowerCase();
    if (CONFIG.UPLOAD.ALLOWED_EXTENSIONS.indexOf(ext) === -1) {
      throw portalError_('VALIDATION', 'Only ' + CONFIG.UPLOAD.ALLOWED_EXTENSIONS.join(', ').toUpperCase() + ' files are allowed.');
    }
    var declaredMime = String(upload.mimeType || '').toLowerCase();
    if (declaredMime && CONFIG.UPLOAD.ALLOWED_MIME_TYPES.indexOf(declaredMime) === -1) {
      throw portalError_('VALIDATION', 'Only PDF and JPEG files are allowed.');
    }

    var maxBytes = Math.round(CONFIG.UPLOAD.MAX_FILE_SIZE_MB * 1024 * 1024);
    var data = String(upload.data || '');
    if (!data) throw portalError_('VALIDATION', 'The selected file is empty.');
    if (data.length > Math.ceil(maxBytes * 4 / 3) + 16) {
      throw portalError_('VALIDATION', 'File is larger than ' + CONFIG.UPLOAD.MAX_FILE_SIZE_MB + ' MB.');
    }
    var bytes;
    try { bytes = Utilities.base64Decode(data); } catch (e) { throw portalError_('VALIDATION', 'The file could not be read. Please try again.'); }
    if (!bytes || bytes.length === 0) throw portalError_('VALIDATION', 'The selected file is empty.');
    if (bytes.length > maxBytes) throw portalError_('VALIDATION', 'File is larger than ' + CONFIG.UPLOAD.MAX_FILE_SIZE_MB + ' MB.');

    var detected = detectFileType_(bytes);
    if (!detected) throw portalError_('VALIDATION', 'The file content is not a valid PDF or JPEG.');
    if ((ext === 'pdf') !== (detected === 'pdf')) throw portalError_('VALIDATION', 'The file extension does not match its content.');
    var mime = detected === 'pdf' ? 'application/pdf' : 'image/jpeg';

    var stamp = Utilities.formatDate(new Date(), getTimeZone_(), 'yyyyMMdd-HHmmss');
    var stagedName = 'STAGED_' + safeNamePart_(appNo) + '_' + safeNamePart_(field.label).slice(0, 40) + '_' + stamp + '.' + (detected === 'pdf' ? 'pdf' : 'jpg');
    var file = getFolder_('STAGING').createFile(Utilities.newBlob(bytes, mime, stagedName));
    file.setDescription('Staged replacement for application ' + appNo + ', column "' + field.header + '". Original file name: ' + originalName);

    var cache = CacheService.getScriptCache();
    var key = stagedKey_(token, appNo, field.key);
    var previous = cache.get(key);
    if (previous) trashStagedFile_(previous);
    cache.put(key, JSON.stringify({ fileId: file.getId(), name: originalName, size: bytes.length, mimeType: mime, at: Date.now() }), stagedTtl_());

    session.uploads = (session.uploads || 0) + 1;
    session.staged = session.staged || [];
    if (session.staged.indexOf(key) === -1) session.staged.push(key);
    putSession_(token, session);

    return { fieldKey: field.key, name: originalName, size: bytes.length, mimeType: mime };
  });
}

/** Cancels a staged replacement (the existing document stays). */
function removeStagedUpload(token, appNo, fieldKey) {
  return respond_(function () {
    var session = getSession_(token, 'applicant');
    appNo = cleanAppNo_(appNo);
    requireApplicationAccess_(session, appNo);
    var key = stagedKey_(token, appNo, String(fieldKey));
    var cache = CacheService.getScriptCache();
    var raw = cache.get(key);
    if (raw) {
      trashStagedFile_(raw);
      cache.remove(key);
    }
    return { removed: !!raw };
  });
}

/**
 * Streams a document to the browser (base64) after verifying the session
 * owns the application. Drive files are never made public.
 * @param {string} source 'current' (document in the sheet) | 'staged' (new upload)
 * @param {number} index  which file when a cell holds several links
 */
function viewDocument(token, appNo, fieldKey, source, index) {
  return respond_(function () {
    var session = getSession_(token, 'applicant');
    appNo = cleanAppNo_(appNo);
    requireApplicationAccess_(session, appNo);
    var ctx = getSheetContext_();
    var field = ctx.byHeader[String(fieldKey)];
    if (!field || field.type !== 'document') throw portalError_('NOT_FOUND', 'Document not found.');

    var fileId = '';
    if (source === 'staged') {
      var staged = getStagedUploadsFor_(token, appNo, [field])[field.key];
      if (!staged) throw portalError_('NOT_FOUND', 'The uploaded file has expired. Please upload it again.');
      fileId = staged.fileId;
    } else {
      var rowIndex = findRowByAppNo_(ctx, appNo);
      var cell = ctx.sheet.getRange(rowIndex, field.colIndex + 1);
      var ids = extractDriveFileIds_(cell.getValue(), cell.getFormula(), cell.getDisplayValue());
      var i = Math.max(0, parseInt(index, 10) || 0);
      fileId = ids[i];
      if (!fileId) throw portalError_('NOT_FOUND', 'No document is attached for this field.');
    }
    return readFileForViewer_(fileId);
  });
}

/* ======================================================================== */
/* Helpers used by Application.gs / Admin.gs                                */
/* ======================================================================== */

/**
 * Extracts Google Drive file IDs from a cell (plain URL, several
 * comma-separated URLs as produced by Google Forms, HYPERLINK formula or a
 * bare file ID).
 */
function extractDriveFileIds_(raw, formula, display) {
  var ids = [];
  var add = function (id) { if (id && ids.indexOf(id) === -1) ids.push(id); };
  [raw, formula, display].forEach(function (v) {
    var s = String(v == null ? '' : v);
    if (!s) return;
    var re = /(?:\/d\/|[?&]id=)([A-Za-z0-9_-]{20,})/g;
    var m;
    while ((m = re.exec(s)) !== null) add(m[1]);
  });
  var bare = String(raw == null ? '' : raw).trim();
  if (!ids.length && /^[A-Za-z0-9_-]{25,}$/.test(bare)) add(bare);
  return ids;
}

/** Name / type / size of a Drive file, without failing on missing files. */
function getDriveFileInfo_(fileId) {
  try {
    var f = DriveApp.getFileById(fileId);
    return { available: true, name: f.getName(), mimeType: f.getMimeType(), size: f.getSize() };
  } catch (e) {
    return { available: false, name: 'Document not accessible', mimeType: '', size: 0 };
  }
}

/** Document block for the applicant form (no file IDs or URLs are sent). */
function buildDocumentView_(field, row, staged) {
  var raw = row.values[field.colIndex];
  var ids = extractDriveFileIds_(raw, row.formulas[field.colIndex], row.display[field.colIndex]);
  var files = ids.map(function (id, i) {
    var info = getDriveFileInfo_(id);
    return { index: i, name: info.name, mimeType: info.mimeType, size: info.size, available: info.available };
  });
  return {
    key: field.key,
    label: field.label,
    required: field.required,
    editable: field.editable,
    help: field.help,
    order: field.order,
    files: files,
    unrecognised: !ids.length && String(row.display[field.colIndex] || '').trim() !== '',
    staged: staged ? { name: staged.name, size: staged.size, mimeType: staged.mimeType } : null
  };
}

function stagedKey_(token, appNo, fieldKey) {
  return STAGED_PREFIX_ + sha256Hex_(token + '|' + appNo + '|' + fieldKey).slice(0, 48);
}

function stagedTtl_() {
  return Math.min(21600, CONFIG.SESSION.ABSOLUTE_TIMEOUT_MINUTES * 60);
}

/** Map of fieldKey -> staged upload info for this session + application. */
function getStagedUploadsFor_(token, appNo, docFields) {
  if (!docFields.length) return {};
  var keyToField = {};
  docFields.forEach(function (f) { keyToField[stagedKey_(token, appNo, f.key)] = f.key; });
  var found = CacheService.getScriptCache().getAll(Object.keys(keyToField));
  var out = {};
  Object.keys(found).forEach(function (k) {
    try { out[keyToField[k]] = JSON.parse(found[k]); } catch (e) { /* ignore */ }
  });
  return out;
}

/** Removes staged references after a successful submission (files were moved). */
function clearStagedUploads_(token, appNo, fields) {
  var cache = CacheService.getScriptCache();
  fields.forEach(function (f) { cache.remove(stagedKey_(token, appNo, f.key)); });
}

/** Trashes staged (never submitted) files, e.g. on logout. */
function discardStagedUploads_(keys) {
  var cache = CacheService.getScriptCache();
  var found = cache.getAll(keys);
  Object.keys(found).forEach(function (k) { trashStagedFile_(found[k]); });
  cache.removeAll(keys);
}

/** Trashes a file ONLY if it is still inside the staging folder. */
function trashStagedFile_(rawJson) {
  try {
    var info = JSON.parse(rawJson);
    var file = DriveApp.getFileById(info.fileId);
    var stagingId = getFolder_('STAGING').getId();
    var parents = file.getParents();
    while (parents.hasNext()) {
      if (parents.next().getId() === stagingId) { file.setTrashed(true); return; }
    }
  } catch (e) {
    console.warn('Could not discard staged file: ' + e);
  }
}

/**
 * Archives previous documents and moves staged files into the upload folder.
 * Called inside the submission lock. Throws (before the sheet is touched)
 * if an existing document cannot be archived.
 * @return {Array<{field:Object,newUrl:string,newId:string,previousValue:string,details:string}>}
 */
function processDocumentReplacements_(appNo, txnId, row, jobs) {
  if (!jobs.length) return [];
  var uploadFolder = getFolder_('UPLOAD');
  var archiveFolder = getFolder_('ARCHIVE');

  // Phase 1: archive every document that is about to be replaced.
  var prepared = jobs.map(function (job) {
    var f = job.field;
    var raw = row.values[f.colIndex];
    var formula = row.formulas[f.colIndex];
    var oldIds = extractDriveFileIds_(raw, formula, row.display[f.colIndex]);
    var archived = [];
    oldIds.forEach(function (id) {
      var file;
      try { file = DriveApp.getFileById(id); } catch (e) {
        archived.push('previous file ' + id + ' not accessible (link kept in this log)');
        return;
      }
      try {
        var name = file.getName();
        if (CONFIG.ARCHIVE_MODE === 'MOVE') {
          file.moveTo(archiveFolder);
          file.setDescription('Archived by correction ' + txnId + ' (application ' + appNo + ').');
          archived.push('previous "' + name + '" (ID ' + id + ') moved to archive');
        } else {
          var copy = file.makeCopy('ARCHIVED_' + safeNamePart_(appNo) + '_' + txnId + '_' + name, archiveFolder);
          copy.setDescription('Archive copy of ' + id + ' made by correction ' + txnId + ' (application ' + appNo + ').');
          archived.push('previous "' + name + '" (ID ' + id + ') archived as ID ' + copy.getId());
        }
      } catch (e2) {
        console.error('Archive failed for ' + id + ': ' + e2);
        throw portalError_('ARCHIVE_FAILED', 'The existing "' + f.label + '" document could not be archived, so nothing was changed. Please try again later or contact the helpdesk.');
      }
    });
    if (!oldIds.length) archived.push('no previous document');
    return { job: job, archived: archived, previousValue: formula || String(raw == null ? '' : raw) };
  });

  // Phase 2: move each staged file to the upload folder with a final name.
  return prepared.map(function (p) {
    var f = p.job.field;
    var staged = p.job.staged;
    var file;
    try { file = DriveApp.getFileById(staged.fileId); } catch (e) {
      throw portalError_('UPLOAD_MISSING', 'The replacement for "' + f.label + '" is no longer available. Please upload it again.');
    }
    var ext = staged.mimeType === 'application/pdf' ? 'pdf' : 'jpg';
    var finalName = safeNamePart_(appNo) + '_' + safeNamePart_(f.label).slice(0, 60) + '_' + txnId + '.' + ext;
    file.setName(finalName);
    file.moveTo(uploadFolder);
    file.setDescription('Replacement for application ' + appNo + ', column "' + f.header + '", correction ' + txnId + '. Uploaded as "' + staged.name + '".');
    var details = 'New file "' + finalName + '" (ID ' + file.getId() + ', ' + Math.round(staged.size / 1024) + ' KB, uploaded as "' +
      staged.name + '"); ' + p.archived.join('; ');
    return { field: f, newUrl: file.getUrl(), newId: file.getId(), previousValue: p.previousValue, details: details };
  });
}

/** Returns {name, mimeType, data(base64)} for the in-portal viewer. */
function readFileForViewer_(fileId) {
  var file;
  try { file = DriveApp.getFileById(fileId); } catch (e) {
    throw portalError_('NOT_FOUND', 'This document is not accessible. Please contact the helpdesk or upload a replacement.');
  }
  var mime = file.getMimeType();
  var blob;
  if (/^application\/vnd\.google-apps\./.test(mime)) {
    blob = file.getAs('application/pdf');
  } else {
    if (file.getSize() > CONFIG.UPLOAD.MAX_VIEW_SIZE_MB * 1024 * 1024) {
      throw portalError_('TOO_LARGE', 'This document is too large to preview here (over ' + CONFIG.UPLOAD.MAX_VIEW_SIZE_MB + ' MB).');
    }
    blob = file.getBlob();
  }
  var outMime = blob.getContentType();
  var viewable = ['application/pdf', 'image/jpeg', 'image/png', 'image/gif', 'image/webp'];
  if (viewable.indexOf(outMime) === -1) {
    try { blob = blob.getAs('application/pdf'); outMime = 'application/pdf'; } catch (e) {
      throw portalError_('UNSUPPORTED', 'Preview is not available for this file type.');
    }
  }
  return { name: file.getName(), mimeType: outMime, data: Utilities.base64Encode(blob.getBytes()) };
}

/** Detects PDF / JPEG from the file signature ("magic bytes"). */
function detectFileType_(bytes) {
  if (bytes.length >= 4 && (bytes[0] & 0xff) === 0x25 && (bytes[1] & 0xff) === 0x50 && (bytes[2] & 0xff) === 0x44 && (bytes[3] & 0xff) === 0x46) return 'pdf';
  if (bytes.length >= 3 && (bytes[0] & 0xff) === 0xff && (bytes[1] & 0xff) === 0xd8 && (bytes[2] & 0xff) === 0xff) return 'jpeg';
  return '';
}

function sanitizeFileName_(name) {
  var s = String(name == null ? '' : name).split(/[\\/]/).pop();
  s = s.replace(/[^A-Za-z0-9 ._()\-]/g, '_').replace(/_+/g, '_').trim();
  if (s.length > 100) {
    var dot = s.lastIndexOf('.');
    s = dot > 0 ? s.slice(0, 90) + s.slice(dot) : s.slice(0, 100);
  }
  return s || 'document';
}

function safeNamePart_(s) {
  return String(s == null ? '' : s).replace(/[^A-Za-z0-9\-]+/g, '_').replace(/^_+|_+$/g, '') || 'x';
}

/**
 * Returns a Drive folder by logical name: 'UPLOAD' | 'ARCHIVE' | 'STAGING'.
 * IDs come from Config.gs or, if blank, from Script Properties set by setupPortal().
 */
function getFolder_(kind) {
  if (RUNTIME_.folders[kind]) return RUNTIME_.folders[kind];
  var id = CONFIG.DRIVE[kind + '_FOLDER_ID'] || PropertiesService.getScriptProperties().getProperty('FOLDER_' + kind);
  if (!id) throw portalError_('CONFIG', 'Document storage is not configured. The administrator must run setupPortal().');
  var folder;
  try { folder = DriveApp.getFolderById(id); } catch (e) {
    throw portalError_('CONFIG', 'Document storage folder is not accessible. Please contact the administrator.');
  }
  RUNTIME_.folders[kind] = folder;
  return folder;
}

/**
 * Daily housekeeping (installed as a time-driven trigger by setupPortal):
 * trashes files left in the STAGING folder by applicants who never
 * submitted. Only the staging folder is touched.
 */
function cleanupStagingFolder() {
  var props = PropertiesService.getScriptProperties();
  var last = Number(props.getProperty('LAST_STAGING_CLEANUP') || 0);
  if (Date.now() - last < 30 * 60000) return 'Skipped – ran recently.'; // also limits abuse via google.script.run
  props.setProperty('LAST_STAGING_CLEANUP', String(Date.now()));
  var cutoff = Date.now() - CONFIG.UPLOAD.STAGING_RETENTION_HOURS * 3600000;
  var files = getFolder_('STAGING').getFiles();
  var trashed = 0;
  var started = Date.now();
  while (files.hasNext() && Date.now() - started < 240000) {
    var f = files.next();
    if (f.getDateCreated().getTime() < cutoff) {
      f.setTrashed(true);
      trashed++;
    }
  }
  console.log('Staging cleanup: ' + trashed + ' file(s) trashed.');
  return 'Trashed ' + trashed + ' file(s).';
}
