/**
 * AuditLog.gs
 * ---------------------------------------------------------------------------
 * Correction_Log: an append-only history of every change.
 *
 * Columns (see LOG_HEADERS in Config.gs):
 *   Correction Transaction ID | Application Number | Field Name |
 *   Previous Value | Updated Value | Modification Date and Time |
 *   Document Replacement Details | Correction Status | Changed By
 *
 * One row is written per changed field. Admin status decisions are logged
 * with Field Name "[STATUS]". Rows are only ever appended – nothing in the
 * portal updates or deletes log rows, and setupPortal() protects the tab so
 * only the owner can edit it manually.
 * ---------------------------------------------------------------------------
 */

var LOG_STATUS_FIELD_ = '[STATUS]';

/**
 * Appends rows in a single batch write. Must be called while holding the
 * script lock (submitCorrection / admin actions do this).
 * @param {Array<Array>} rows each row follows LOG_HEADERS
 */
function appendLogRows_(rows) {
  if (!rows || !rows.length) return;
  var sheet = getLogSheet_();
  var safe = rows.map(function (r) {
    return LOG_HEADERS.map(function (_, i) {
      var v = r[i];
      if (isDateObject_(v)) return v;
      return neutralizeFormula_(v == null ? '' : String(v).slice(0, 49000));
    });
  });
  var start = sheet.getLastRow() + 1;
  var overflow = start + safe.length - 1 - sheet.getMaxRows();
  if (overflow > 0) {
    // New rows inherit the plain-text formatting of the row above.
    sheet.insertRowsAfter(sheet.getMaxRows(), overflow + 500);
  }
  sheet.getRange(start, 1, safe.length, LOG_HEADERS.length).setValues(safe);
}

/**
 * Reads every log row for one application, grouped by transaction (newest
 * first). Uses a single TextFinder search plus one bulk read.
 */
function getCorrectionHistory_(appNo) {
  var sheet = getLogSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var appCol = sheet.getRange(2, 2, lastRow - 1, 1);
  var hits = appCol.createTextFinder(String(appNo)).matchEntireCell(true).matchCase(false).findAll();
  if (!hits.length) return [];
  var minRow = hits[0].getRow();
  var maxRow = hits[hits.length - 1].getRow();
  var block = sheet.getRange(minRow, 1, maxRow - minRow + 1, LOG_HEADERS.length).getValues();
  var tz = getTimeZone_();
  var wanted = String(appNo).toLowerCase();
  var byTxn = {};
  var order = [];
  block.forEach(function (r) {
    if (String(r[1]).trim().toLowerCase() !== wanted) return;
    var txn = String(r[0]);
    if (!byTxn[txn]) {
      byTxn[txn] = { transactionId: txn, time: '', timeMs: 0, status: '', changedBy: '', changes: [], statusEvents: [] };
      order.push(txn);
    }
    var t = byTxn[txn];
    var when = isDateObject_(r[5]) ? r[5] : null;
    var whenText = when ? Utilities.formatDate(when, tz, 'dd/MM/yyyy HH:mm:ss') : String(r[5]);
    if (String(r[2]) === LOG_STATUS_FIELD_) {
      t.statusEvents.push({ time: whenText, status: String(r[7]), note: String(r[6]), by: String(r[8]) });
      t.status = String(r[7]);
      return;
    }
    if (!t.time) { t.time = whenText; t.timeMs = when ? when.getTime() : 0; }
    if (!t.statusEvents.length) t.status = String(r[7]);
    t.changedBy = String(r[8]);
    t.changes.push({
      field: String(r[2]),
      previousValue: String(r[3]),
      updatedValue: String(r[4]),
      documentDetails: String(r[6])
    });
  });
  return order.map(function (k) { return byTxn[k]; }).reverse();
}

/**
 * Reads the whole log once and returns aggregate figures and the most recent
 * transactions (admin dashboard).
 */
function getLogSummary_(recentLimit) {
  var sheet = getLogSheet_();
  var lastRow = sheet.getLastRow();
  var result = { transactions: 0, fieldChanges: 0, documentsReplaced: 0, recent: [] };
  if (lastRow < 2) return result;
  var data = sheet.getRange(2, 1, lastRow - 1, LOG_HEADERS.length).getValues();
  var tz = getTimeZone_();
  var seen = {};
  var recent = [];
  for (var i = data.length - 1; i >= 0; i--) {
    var r = data[i];
    var txn = String(r[0]);
    if (!txn) continue;
    if (String(r[2]) === LOG_STATUS_FIELD_) continue;
    result.fieldChanges++;
    if (String(r[6]).indexOf('New file') === 0) result.documentsReplaced++;
    if (!seen[txn]) {
      seen[txn] = { transactionId: txn, applicationNumber: String(r[1]), time: isDateObject_(r[5]) ? Utilities.formatDate(r[5], tz, 'dd/MM/yyyy HH:mm') : String(r[5]), status: String(r[7]), fields: 0 };
      result.transactions++;
      if (recent.length < (recentLimit || 20)) recent.push(seen[txn]);
    }
    seen[txn].fields++;
  }
  result.recent = recent;
  return result;
}

/** Returns all log rows belonging to one transaction (with sheet row numbers). */
function getTransactionRows_(txnId) {
  var sheet = getLogSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var hits = sheet.getRange(2, 1, lastRow - 1, 1).createTextFinder(String(txnId)).matchEntireCell(true).findAll();
  if (!hits.length) return [];
  var minRow = hits[0].getRow();
  var maxRow = hits[hits.length - 1].getRow();
  var block = sheet.getRange(minRow, 1, maxRow - minRow + 1, LOG_HEADERS.length).getValues();
  return block.filter(function (r) { return String(r[0]) === String(txnId); });
}
