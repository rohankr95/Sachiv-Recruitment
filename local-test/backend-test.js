/**
 * Runs the real .gs code against the in-memory mock and checks the main
 * functional, security and performance behaviours.
 *   node local-test/backend-test.js
 */
'use strict';
process.env.TZ = 'Asia/Kolkata';
const assert = require('assert');
const { createContext } = require('./gas-mock');

let passed = 0;
const results = [];
function test(name, fn) {
  try { fn(); passed++; results.push('  ✓ ' + name); }
  catch (e) { results.push('  ✗ ' + name + '\n      ' + (e && e.stack ? e.stack.split('\n').slice(0, 3).join('\n      ') : e)); process.exitCode = 1; }
}
const same = (a, b, m) => assert.strictEqual(JSON.stringify(a), JSON.stringify(b), m);
const ok = r => { if (!r.ok) throw new Error('Expected ok, got ' + JSON.stringify(r)); return r.data; };
const fail = (r, code) => { assert.strictEqual(r.ok, false, 'expected failure, got ok'); if (code) assert.strictEqual(r.code, code, r.message); return r; };

function bootstrap() {
  const g = createContext({ verbose: false });
  const t0 = Date.now();
  g.createSampleData();
  const tSample = Date.now() - t0;
  g.setupPortal();
  return { g, tSample };
}

const { g, tSample } = bootstrap();
const ss = Object.values(g.__state.spreadsheets)[0];
const sheet = ss.getSheetByName('Form Responses 1');
const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
const col = h => headers.indexOf(h) + 1;
const cell = (row, h) => sheet.getRange(row, col(h)).getValue();
const disp = (row, h) => sheet.getRange(row, col(h)).getDisplayValue();
const dobIso = row => g.toIsoDate_(cell(row, 'Date of Birth'));
const rowOf = appNo => g.findRowByAppNo_(g.getSheetContext_(), appNo);

test('sample data has 2500 rows x 50 columns plus 4 system columns', () => {
  assert.strictEqual(sheet.getLastRow(), 2501);
  assert.strictEqual(headers.length, 54);
  ['Last Modified', 'Correction Count', 'Correction Status', 'Last Correction ID'].forEach(h => assert.ok(headers.includes(h), h));
});

test('Field_Config auto-detects document, date, select, contact columns', () => {
  const ctx = g.getSheetContext_();
  const t = k => ctx.byHeader[k].type;
  assert.strictEqual(t('Photograph'), 'document');
  assert.strictEqual(t('Caste Certificate'), 'document');
  assert.strictEqual(t('Aadhaar Card Copy'), 'document');
  assert.strictEqual(t('Date of Birth'), 'date');
  assert.strictEqual(t('Gender'), 'select');
  assert.strictEqual(t('Category'), 'select');
  assert.strictEqual(t('Email Address'), 'email');
  assert.strictEqual(t('Mobile Number'), 'mobile');
  assert.strictEqual(t('Pin Code'), 'pincode');
  assert.strictEqual(t('Application Number'), 'readonly');
  assert.strictEqual(ctx.byHeader['Application Number'].editable, false);
  assert.strictEqual(t('Correction Count'), 'hidden');
  assert.strictEqual(ctx.byHeader['Block'].section, 'Address Details');
  assert.strictEqual(ctx.byHeader['10th Board'].section, 'Educational Qualifications');
  assert.strictEqual(ctx.fields.filter(f => f.type === 'document').length, 7);
});

test('setup functions refuse to run for anonymous web visitors', () => {
  g.__state.activeUser = '';
  assert.throws(() => g.setupPortal(), /project owner/);
  assert.throws(() => g.setupAdminAccount(), /project owner/);
  assert.throws(() => g.createSampleData(), /project owner/);
  g.__state.activeUser = 'owner@example.com';
});

// --- Authentication -------------------------------------------------------
const r2 = 2; // APP202600001
const email1 = cell(r2, 'Email Address');
const mobile1 = String(cell(r2, 'Mobile Number'));
let token;

test('login with email + DOB succeeds and returns only own application', () => {
  const d = ok(g.login(email1.toUpperCase(), dobIso(r2)));
  assert.match(d.token, /^[a-f0-9]{64}$/);
  same(d.applications.map(a => a.applicationNumber), ['APP202600001']);
  assert.strictEqual(d.applications[0].duplicateEmail, false);
  token = d.token;
});

test('login with mobile (+91 / spaces) + DOB succeeds', () => {
  const d = ok(g.login('+91 ' + mobile1.slice(0, 5) + ' ' + mobile1.slice(5), dobIso(r2)));
  same(d.applications.map(a => a.applicationNumber), ['APP202600001']);
});

test('row 9 stored "+91 …" mobile still matches 10-digit login', () => {
  const m = g.normalizeMobile_(cell(9, 'Mobile Number'));
  ok(g.login(m, dobIso(9)));
});

test('wrong DOB fails with generic message; input validation errors', () => {
  const r = fail(g.login(email1, '1970-01-01'), 'AUTH_FAILED');
  assert.match(r.message, /do not match/);
  fail(g.login('not-an-email@', '2000-01-01'), 'VALIDATION');
  fail(g.login('12345', '2000-01-01'), 'VALIDATION');
  fail(g.login(email1, '2999-01-01'), 'VALIDATION');
  fail(g.login(email1, ''), 'VALIDATION');
});

test('rate limiting locks an identifier after 5 failures, even with correct DOB', () => {
  const e = cell(5, 'Email Address');
  for (let i = 0; i < 5; i++) fail(g.login(e, '1970-01-01'), i < 4 ? 'AUTH_FAILED' : 'AUTH_FAILED');
  fail(g.login(e, dobIso(5)), 'RATE_LIMITED');
  // other identifiers unaffected
  ok(g.login(cell(6, 'Email Address'), dobIso(6)));
});

test('duplicate email + same DOB: both applications listed and flagged duplicate', () => {
  // rows for APP202600025 / APP202600026 share email, mobile and DOB
  const row = rowOf('APP202600026');
  const d = ok(g.login(cell(row, 'Email Address'), dobIso(row)));
  same(d.applications.map(a => a.applicationNumber).sort(), ['APP202600025', 'APP202600026']);
  assert.ok(d.applications.every(a => a.duplicateEmail && a.emailShareCount === 2));
});

test('duplicate email + different DOB (MATCH_DOB): only own application, still flagged duplicate', () => {
  const row = rowOf('APP202600041');
  const d = ok(g.login(cell(row, 'Email Address'), dobIso(row)));
  same(d.applications.map(a => a.applicationNumber), ['APP202600041']);
  assert.strictEqual(d.applications[0].duplicateEmail, true);
  assert.strictEqual(d.applications[0].emailShareCount, 2);
});

test('SHOW_ALL policy lists every application on the email once DOB matches one', () => {
  g.CONFIG.DUPLICATE_IDENTIFIER_POLICY = 'SHOW_ALL';
  const row = rowOf('APP202600041');
  const d = ok(g.login(cell(row, 'Email Address'), dobIso(row)));
  same(d.applications.map(a => a.applicationNumber).sort(), ['APP202600040', 'APP202600041']);
  g.CONFIG.DUPLICATE_IDENTIFIER_POLICY = 'MATCH_DOB';
});

// --- Access control -------------------------------------------------------
test('session cannot load another applicant\'s record; bad tokens rejected', () => {
  fail(g.loadApplication(token, 'APP202600002'), 'FORBIDDEN');
  fail(g.loadApplication('x'.repeat(64), 'APP202600001'), 'SESSION_EXPIRED');
  fail(g.loadApplication(null, 'APP202600001'), 'SESSION_EXPIRED');
  fail(g.viewDocument(token, 'APP202600002', 'Photograph', 'current', 0), 'FORBIDDEN');
  fail(g.uploadReplacement(token, 'APP202600002', 'Photograph', {}), 'FORBIDDEN');
  fail(g.submitCorrection(token, 'APP202600002', { declaration: true }), 'FORBIDDEN');
});

test('admin functions reject applicant tokens', () => {
  fail(g.adminGetDashboard(token), 'SESSION_EXPIRED');
  fail(g.adminGetApplication(token, 'APP202600001'), 'SESSION_EXPIRED');
});

test('idle session expiry', () => {
  const d = ok(g.login(email1, dobIso(r2)));
  const key = 'SESS_' + d.token;
  const s = JSON.parse(g.__state.cache.get(key));
  s.last = Date.now() - 16 * 60000;
  g.__state.cache.set(key, JSON.stringify(s));
  fail(g.keepAlive(d.token), 'SESSION_EXPIRED');
  assert.strictEqual(g.__state.cache.has(key), false);
});

// --- Retrieval ------------------------------------------------------------
let app;
test('loadApplication returns sections, prefilled values and documents', () => {
  app = ok(g.loadApplication(token, 'APP202600001'));
  assert.strictEqual(app.applicationNumber, 'APP202600001');
  const names = app.sections.map(s => s.name);
  same(names.slice(0, 3), ['Personal Information', 'Address Details', 'Educational Qualifications']);
  const all = [].concat(...app.sections.map(s => s.fields));
  const dob = all.find(f => f.key === 'Date of Birth');
  assert.strictEqual(dob.value, dobIso(r2));
  assert.ok(!all.some(f => f.key === 'Correction Count'), 'hidden fields must not be sent');
  assert.strictEqual(app.documents.length, 7);
  assert.strictEqual(app.documents[0].files.length, 1);
  assert.ok(!JSON.stringify(app).includes('drive.google.com'), 'no Drive URLs sent to applicant');
  assert.strictEqual(app.version, '|0');
});

test('viewDocument streams the file as base64 for the owner', () => {
  const doc = ok(g.viewDocument(token, 'APP202600001', 'Photograph', 'current', 0));
  assert.strictEqual(doc.mimeType, 'image/jpeg');
  assert.ok(Buffer.from(doc.data, 'base64').length > 1000);
});

// --- Uploads ------------------------------------------------------------------
const PDF_B64 = Buffer.from('%PDF-1.4\n%test\n1 0 obj<<>>endobj\ntrailer<<>>\n%%EOF').toString('base64');
test('upload validation: extension, signature, size', () => {
  fail(g.uploadReplacement(token, 'APP202600001', 'Caste Certificate', { fileName: 'x.png', mimeType: 'image/png', data: PDF_B64 }), 'VALIDATION');
  fail(g.uploadReplacement(token, 'APP202600001', 'Caste Certificate', { fileName: 'x.pdf', mimeType: 'application/pdf', data: Buffer.from('MZ fake exe').toString('base64') }), 'VALIDATION');
  fail(g.uploadReplacement(token, 'APP202600001', 'Caste Certificate', { fileName: 'x.jpg', mimeType: 'image/jpeg', data: PDF_B64 }), 'VALIDATION');
  const big = Buffer.concat([Buffer.from('%PDF-'), Buffer.alloc(3 * 1024 * 1024)]).toString('base64');
  fail(g.uploadReplacement(token, 'APP202600001', 'Caste Certificate', { fileName: 'x.pdf', mimeType: 'application/pdf', data: big }), 'VALIDATION');
  fail(g.uploadReplacement(token, 'APP202600001', 'Applicant Name', { fileName: 'x.pdf', mimeType: 'application/pdf', data: PDF_B64 }), 'FORBIDDEN');
});

let stagedId;
test('valid PDF is staged (sheet unchanged until submit)', () => {
  const before = cell(r2, 'Caste Certificate');
  const d = ok(g.uploadReplacement(token, 'APP202600001', 'Caste Certificate', { fileName: '../../caste<cert>.pdf', mimeType: 'application/pdf', data: PDF_B64 }));
  assert.strictEqual(d.name, 'caste_cert_.pdf');
  assert.strictEqual(cell(r2, 'Caste Certificate'), before);
  const staged = g.getStagedUploadsFor_(token, 'APP202600001', [g.getSheetContext_().byHeader['Caste Certificate']]);
  stagedId = staged['Caste Certificate'].fileId;
  assert.ok(stagedId);
  ok(g.viewDocument(token, 'APP202600001', 'Caste Certificate', 'staged', 0));
  const again = ok(g.loadApplication(token, 'APP202600001'));
  assert.strictEqual(again.documents.find(x => x.key === 'Caste Certificate').staged.name, 'caste_cert_.pdf');
});

// --- Submission ---------------------------------------------------------------
test('server-side validation rejects bad values and non-editable fields', () => {
  const r = fail(g.submitCorrection(token, 'APP202600001', {
    version: app.version, declaration: true,
    changes: { 'Mobile Number': '12345', 'Email Address': 'bad', 'Pin Code': '0123', 'Application Number': 'HACK', 'Gender': 'Robot', 'Correction Count': '99', 'Date of Birth': '2999-01-01' }
  }), 'VALIDATION');
  const fe = r.details.fieldErrors;
  ['Mobile Number', 'Email Address', 'Pin Code', 'Application Number', 'Gender', 'Correction Count', 'Date of Birth'].forEach(k => assert.ok(fe[k], k + ' should error'));
  fail(g.submitCorrection(token, 'APP202600001', { version: app.version, declaration: false, changes: { 'Applicant Name': 'X' } }), 'VALIDATION');
  fail(g.submitCorrection(token, 'APP202600001', { version: app.version, declaration: true, changes: { 'Applicant Name': cell(r2, 'Applicant Name') } }), 'NO_CHANGES');
});

test('required field cannot be blanked', () => {
  const r = fail(g.submitCorrection(token, 'APP202600001', { version: app.version, declaration: true, changes: { 'Applicant Name': '   ' } }), 'VALIDATION');
  assert.ok(r.details.fieldErrors['Applicant Name']);
});

let txn;
test('successful submission updates only changed cells, system columns, log, documents', () => {
  const untouched = cell(r2, "Father's Name");
  const oldCaste = cell(r2, 'Caste Certificate');
  const oldFileId = g.extractDriveFileIds_(oldCaste)[0];
  const d = ok(g.submitCorrection(token, 'APP202600001', {
    version: app.version, declaration: true,
    changes: { 'Applicant Name': 'Corrected  Name ', 'Date of Birth': '2001-02-03', 'Identification Mark': '=HYPERLINK("http://evil")', 'Pin Code': '497 001', 'Aadhaar Number': '012345678901' },
    replacedDocuments: ['Caste Certificate']
  }));
  txn = d.transactionId;
  assert.match(txn, /^CRN-\d{8}-\d{6}-[0-9A-F]{6}$/);
  assert.strictEqual(cell(r2, 'Applicant Name'), 'Corrected Name');
  assert.strictEqual(g.toIsoDate_(cell(r2, 'Date of Birth')), '2001-02-03');
  assert.strictEqual(cell(r2, 'Identification Mark'), '=HYPERLINK("http://evil")', 'formula stored as text, not executed');
  assert.strictEqual(sheet.formulas[r2 - 1][col('Identification Mark') - 1], '');
  assert.strictEqual(String(cell(r2, 'Pin Code')), '497001');
  assert.strictEqual(cell(r2, 'Aadhaar Number'), '012345678901', 'leading zero preserved');
  assert.strictEqual(cell(r2, "Father's Name"), untouched);
  assert.strictEqual(cell(r2, 'Correction Count'), 1);
  assert.strictEqual(cell(r2, 'Correction Status'), 'Modified');
  assert.strictEqual(cell(r2, 'Last Correction ID'), txn);
  assert.ok(Object.prototype.toString.call(cell(r2, 'Last Modified')) === '[object Date]');
  // documents
  const newUrl = cell(r2, 'Caste Certificate');
  assert.ok(newUrl.includes(stagedId), 'new link points at the uploaded file');
  const files = g.__state.files;
  assert.strictEqual(files[stagedId].parents[0].getId(), g.getFolder_('UPLOAD').getId());
  assert.ok(files[oldFileId] && !files[oldFileId].trashed, 'original document not deleted');
  const archived = Object.values(files).filter(f => f.parents && f.parents[0] && f.parents[0].getId && f.parents[0].getId() === g.getFolder_('ARCHIVE').getId());
  assert.strictEqual(archived.length, 1);
  // log
  const log = ss.getSheetByName('Correction_Log');
  const rows = log.getRange(2, 1, log.getLastRow() - 1, 9).getValues();
  const mine = rows.filter(r => r[0] === txn);
  assert.strictEqual(mine.length, 6);
  const docRow = mine.find(r => r[2] === 'Caste Certificate');
  assert.ok(/^New file/.test(docRow[6]) && docRow[3] === oldCaste && docRow[4] === newUrl);
  assert.ok(mine.every(r => r[1] === 'APP202600001' && r[7] === 'Modified'));
  const dobRow = mine.find(r => r[2] === 'Date of Birth');
  assert.strictEqual(dobRow[4], '03/02/2001');
  // staged reference cleared
  same(g.getStagedUploadsFor_(token, 'APP202600001', [g.getSheetContext_().byHeader['Caste Certificate']]), {});
});

test('stale version is rejected (optimistic locking) and no duplicate rows are created', () => {
  const rowsBefore = sheet.getLastRow();
  fail(g.submitCorrection(token, 'APP202600001', { version: app.version, declaration: true, changes: { 'Applicant Name': 'Another' } }), 'CONFLICT');
  const fresh = ok(g.loadApplication(token, 'APP202600001'));
  assert.strictEqual(fresh.version, txn + '|1');
  ok(g.submitCorrection(token, 'APP202600001', { version: fresh.version, declaration: true, changes: { 'Block': 'Odgi' } }));
  assert.strictEqual(cell(r2, 'Correction Count'), 2);
  assert.strictEqual(sheet.getLastRow(), rowsBefore);
});

test('MAX_CORRECTIONS_PER_APPLICATION and PORTAL_OPEN are enforced', () => {
  g.CONFIG.MAX_CORRECTIONS_PER_APPLICATION = 2;
  const fresh = ok(g.loadApplication(token, 'APP202600001'));
  assert.strictEqual(fresh.canSubmit, false);
  fail(g.submitCorrection(token, 'APP202600001', { version: fresh.version, declaration: true, changes: { 'Block': 'Surajpur' } }), 'NOT_ALLOWED');
  g.CONFIG.MAX_CORRECTIONS_PER_APPLICATION = 0;
  g.CONFIG.PORTAL_OPEN = false;
  fail(g.submitCorrection(token, 'APP202600001', { version: fresh.version, declaration: true, changes: { 'Block': 'Surajpur' } }), 'PORTAL_CLOSED');
  g.CONFIG.PORTAL_OPEN = true;
});

test('login still works with corrected DOB; old DOB no longer works', () => {
  ok(g.login(email1, '2001-02-03'));
});

test('logout destroys the session and trashes unsubmitted uploads', () => {
  const d = ok(g.login(cell(3, 'Email Address'), dobIso(3)));
  ok(g.uploadReplacement(d.token, 'APP202600002', 'Domicile Certificate', { fileName: 'd.pdf', mimeType: 'application/pdf', data: PDF_B64 }));
  const staged = Object.values(g.__state.files).filter(f => f.parents && f.parents[0] && f.parents[0].getId && f.parents[0].getId() === g.getFolder_('STAGING').getId() && !f.trashed);
  assert.strictEqual(staged.length, 1);
  ok(g.logout(d.token));
  assert.strictEqual(staged[0].trashed, true);
  fail(g.loadApplication(d.token, 'APP202600002'), 'SESSION_EXPIRED');
});

// --- Admin -------------------------------------------------------------------
let admin;
test('admin account setup hashes password and removes plain text', () => {
  g.__state.props.set('NEW_ADMIN_USERNAME', 'Collector');
  g.__state.props.set('NEW_ADMIN_PASSWORD', 'S3cure-Passw0rd!');
  g.setupAdminAccount();
  assert.strictEqual(g.__state.props.has('NEW_ADMIN_PASSWORD'), false);
  assert.ok(!g.__state.props.get('ADMIN_ACCOUNTS').includes('S3cure'));
  fail(g.adminLogin('collector', 'wrong-password'), 'AUTH_FAILED');
  admin = ok(g.adminLogin('collector', 'S3cure-Passw0rd!')).token;
});

test('admin dashboard statistics', () => {
  const d = ok(g.adminGetDashboard(admin));
  assert.strictEqual(d.stats.totalApplications, 2500);
  assert.strictEqual(d.stats.applicationsCorrected, 1);
  assert.strictEqual(d.stats.documentsReplaced, 1);
  assert.strictEqual(d.stats.totalTransactions, 2);
  assert.ok(d.stats.duplicateEmailGroups >= 100);
  assert.strictEqual(d.recent[0].applicationNumber, 'APP202600001');
});

test('admin search, details, history, duplicates and document access', () => {
  assert.strictEqual(ok(g.adminSearch(admin, 'app202600001')).results.length, 1);
  const byEmail = ok(g.adminSearch(admin, cell(rowOf('APP202600026'), 'Email Address')));
  assert.strictEqual(byEmail.results.length, 2);
  const d = ok(g.adminGetApplication(admin, 'APP202600001'));
  assert.strictEqual(d.history.length, 2);
  assert.ok(d.modifiedFields.includes('Caste Certificate'));
  assert.ok(d.modifiedFields.includes('Applicant Name'));
  const caste = d.documents.find(x => x.header === 'Caste Certificate');
  assert.strictEqual(caste.timesReplaced, 1);
  ok(g.adminViewDocument(admin, 'APP202600001', caste.files[0].id));
  // previous (archived) document referenced in history is viewable
  const prevId = g.extractDriveFileIds_(d.history[1].changes.find(c => c.field === 'Caste Certificate').previousValue)[0];
  ok(g.adminViewDocument(admin, 'APP202600001', prevId));
  // unrelated file is refused
  fail(g.adminViewDocument(admin, 'APP202600001', 'file_' + 'a'.repeat(30)), 'FORBIDDEN');
  const dups = ok(g.adminGetDuplicates(admin));
  assert.ok(dups.email.some(gp => gp.applications.some(a => a.applicationNumber === 'APP202600041') && gp.differentPersons));
});

test('approval workflow: pending -> reject reverts values', () => {
  g.CONFIG.REQUIRE_APPROVAL = true;
  const row = rowOf('APP202600003');
  const t = ok(g.login(cell(row, 'Email Address'), dobIso(row))).token;
  const a = ok(g.loadApplication(t, 'APP202600003'));
  const oldName = cell(row, 'Applicant Name');
  const oldDob = disp(row, 'Date of Birth');
  const res = ok(g.submitCorrection(t, 'APP202600003', { version: a.version, declaration: true, changes: { 'Applicant Name': 'Temp Name', 'Date of Birth': '1999-09-09' } }));
  assert.strictEqual(res.status, 'Pending Approval');
  assert.strictEqual(ok(g.adminGetDashboard(admin)).stats.pendingCorrections, 1);
  const r = ok(g.adminSetStatus(admin, 'APP202600003', res.transactionId, 'REJECT', 'Not supported by documents'));
  same(r.reverted.sort(), ['Applicant Name', 'Date of Birth']);
  assert.strictEqual(cell(row, 'Applicant Name'), oldName);
  assert.strictEqual(disp(row, 'Date of Birth'), oldDob);
  assert.strictEqual(cell(row, 'Correction Status'), 'Rejected');
  fail(g.adminSetStatus(admin, 'APP202600003', res.transactionId, 'APPROVE', ''), 'CONFLICT');
  const hist = ok(g.adminGetApplication(admin, 'APP202600003')).history;
  assert.strictEqual(hist[0].status, 'Rejected');
  g.CONFIG.REQUIRE_APPROVAL = false;
});

test('duplicate application numbers block corrections', () => {
  sheet.getRange(rowOf('APP202600005'), col('Application Number')).setValue('APP202600004');
  const row = 6; // now duplicate APP202600004
  const t = ok(g.login(cell(row, 'Email Address'), dobIso(row)));
  assert.strictEqual(t.applications[0].duplicateRecord, true);
  fail(g.loadApplication(t.token, 'APP202600004'), 'DATA_INTEGRITY');
});

test('unexpected errors are reported with a reference, not a stack trace', () => {
  const orig = g.getSheetContext_;
  g.getSheetContext_ = () => { throw new Error('internal secret detail'); };
  const r = g.loadApplication(token, 'APP202600001');
  g.getSheetContext_ = orig;
  assert.strictEqual(r.code, 'SERVER_ERROR');
  assert.ok(!r.message.includes('secret'));
});

test('doGet renders applicant and admin pages without leaking config', () => {
  const html = g.doGet({ parameter: {} }).getContent();
  assert.ok(html.includes('Applicant Login'));
  assert.ok(!html.includes(ss.getId()), 'spreadsheet ID must not appear in page');
  assert.ok(!html.includes(g.getFolder_('UPLOAD').getId()), 'folder ID must not appear');
  const adminHtml = g.doGet({ parameter: { page: 'admin' } }).getContent();
  assert.ok(adminHtml.includes('Administrator Login'));
});

// --- Performance (mock timings indicate call counts, not real latency) --------
test('login reads only a handful of columns (no full-sheet read)', () => {
  let cellsRead = 0;
  const proto = Object.getPrototypeOf(sheet.getRange(1, 1));
  const ov = proto.getValues, od = proto.getDisplayValues;
  proto.getValues = function () { cellsRead += this.nr * this.nc; return ov.call(this); };
  proto.getDisplayValues = function () { cellsRead += this.nr * this.nc; return od.call(this); };
  ok(g.login(cell(10, 'Email Address'), dobIso(10)));
  proto.getValues = ov; proto.getDisplayValues = od;
  assert.ok(cellsRead < 2500 * 12, 'cells read: ' + cellsRead);
});

console.log(results.join('\n'));
console.log('\n' + passed + '/' + results.length + ' passed. (sample data generation: ' + tSample + ' ms in mock)');
