/**
 * Tests the portal against an exported copy of YOUR sheet.
 *   node local-test/sheet-data-test.js /path/to/sheet.json [reportDir]
 * sheet.json = {sheetName, values:[[...]], formulas:[[...]]} with dates as {"d": ISO}.
 * Real data must never be committed – keep the JSON outside the repository.
 */
'use strict';
process.env.TZ = 'Asia/Kolkata';
const fs = require('fs');
const path = require('path');
const { createContext, loadSheetJson } = require('./gas-mock');

const data = JSON.parse(fs.readFileSync(process.argv[2], 'utf8'));
const OUT = process.argv[3] || null;
const g = createContext();
const { sheet } = loadSheetJson(g, data);
const log = [];
const say = s => { log.push(s); console.log(s); };
let failures = 0;
const check = (cond, msg) => { if (!cond) { failures++; say('  ✗ ' + msg); } return cond; };

/* 1. Setup ------------------------------------------------------------- */
say('== setupPortal() report ==');
say(g.setupPortal());
const ctx = g.getSheetContext_();
const H = ctx.headers;
const cols = ctx.cols;
say('\n== Key columns ==');
['appNo', 'email', 'mobile', 'dob', 'name'].forEach(k => say('  ' + k.padEnd(6) + ' -> ' + (cols[k] === -1 ? 'NOT FOUND' : (cols[k] + 1) + '. ' + H[cols[k]])));

say('\n== Field mapping (column. header -> section / type / flags) ==');
ctx.fields.forEach(f => say('  ' + String(f.colIndex + 1).padStart(3) + '. ' + f.header.slice(0, 70) + '\n        -> ' + f.section + ' / ' + f.type +
  (f.required ? ' / required' : '') + (f.editable ? '' : ' / LOCKED') + (f.pattern ? ' / pattern' : '') +
  (f.options.length ? ' / options: ' + f.options.join(' | ') : '')));

/* 2. Data quality --------------------------------------------------------- */
const N = ctx.dataRows;
const col = (i, disp) => g.readColumn_(ctx, i, disp);
const appNos = col(cols.appNo, true), emails = col(cols.email), mobiles = col(cols.mobile), dobs = col(cols.dob);
const rowsWithDob = [], noDob = [];
for (let i = 0; i < N; i++) (g.toIsoDate_(dobs[i]) ? rowsWithDob : noDob).push(i);
const appCount = {};
appNos.forEach(a => { appCount[a] = (appCount[a] || 0) + 1; });
const badAppNos = Object.keys(appCount).filter(a => appCount[a] > 1 || !a || /^#/.test(a));
const mobCount = {};
mobiles.forEach(m => { const k = g.normalizeMobile_(m); if (k) mobCount[k] = (mobCount[k] || 0) + 1; });
const dupMob = Object.keys(mobCount).filter(k => mobCount[k] > 1);
const emCount = {};
emails.forEach(e => { const k = g.normalizeEmail_(e); if (k) emCount[k] = (emCount[k] || 0) + 1; });
const dupEm = Object.keys(emCount).filter(k => emCount[k] > 1);
say('\n== Data quality ==');
say('  rows: ' + N + ', with a valid DOB: ' + rowsWithDob.length + ', without DOB (cannot log in): ' + noDob.length);
say('  problem application numbers: ' + (badAppNos.map(a => '"' + a + '" x' + appCount[a]).join(', ') || 'none'));
say('  shared emails: ' + dupEm.length + ', shared mobile numbers: ' + dupMob.length + ' (' + dupMob.reduce((s, k) => s + mobCount[k], 0) + ' rows)');

/* 3. Login for every row ---------------------------------------------------- */
say('\n== Login test for every row ==');
const clearRl = () => { for (const k of [...g.__state.cache.keys()]) if (k.startsWith('RL_')) g.__state.cache.delete(k); };
let emailOk = 0, mobileOk = 0, emailFail = [], mobileFail = [], dupFlagged = 0, dupExpected = 0;
const tokens = {};
rowsWithDob.forEach(i => {
  const dob = g.toIsoDate_(dobs[i]);
  const app = appNos[i];
  const r1 = g.login(String(emails[i]), dob);
  if (r1.ok && r1.data.applications.some(a => a.applicationNumber === app)) { emailOk++; tokens[app] = r1.data.token; }
  else emailFail.push(app + ' (' + (r1.message || 'not listed') + ')');
  const m = g.normalizeMobile_(mobiles[i]);
  const r2 = g.login(m, dob);
  if (r2.ok && r2.data.applications.some(a => a.applicationNumber === app)) {
    mobileOk++;
    if (mobCount[m] > 1) {
      dupExpected++;
      if (r2.data.applications.find(a => a.applicationNumber === app).duplicateMobile) dupFlagged++;
    }
  } else mobileFail.push(app + ' (' + (r2.message || 'not listed') + ')');
  clearRl();
});
say('  email + DOB : ' + emailOk + '/' + rowsWithDob.length + ' succeeded' + (emailFail.length ? '; failed: ' + emailFail.slice(0, 10).join('; ') : ''));
say('  mobile + DOB: ' + mobileOk + '/' + rowsWithDob.length + ' succeeded' + (mobileFail.length ? '; failed: ' + mobileFail.slice(0, 10).join('; ') : ''));
say('  rows sharing a mobile flagged "Duplicate": ' + dupFlagged + '/' + dupExpected);
const nodobRow = noDob[0];
if (nodobRow !== undefined) {
  const r = g.login(String(emails[nodobRow]), '1990-01-01');
  say('  row without DOB (' + appNos[nodobRow] + ') login: ' + (r.ok ? 'ALLOWED (unexpected)' : 'refused – ' + r.code));
  check(!r.ok, 'incomplete row was able to log in');
  clearRl();
}
check(emailFail.filter(s => !/^#/.test(s)).length === 0, 'some valid rows could not log in by email');

/* Shared mobile examples */
say('\n== Shared mobile numbers (what the applicant sees) ==');
dupMob.slice(0, 6).forEach(m => {
  const idx = []; mobiles.forEach((v, i) => { if (g.normalizeMobile_(v) === m) idx.push(i); });
  const sameDob = new Set(idx.map(i => g.toIsoDate_(dobs[i]))).size === 1;
  const r = g.login(m, g.toIsoDate_(dobs[idx[0]]));
  clearRl();
  say('  ******' + m.slice(-4) + ': rows ' + idx.map(i => appNos[i]).join(', ') + (sameDob ? ' (same DOB)' : ' (different DOBs)') +
      ' -> login with first DOB shows: ' + (r.ok ? r.data.applications.map(a => a.applicationNumber + (a.duplicateMobile ? ' [Duplicate]' : '')).join(', ') : r.code));
});

/* 4. Load every application ------------------------------------------------ */
say('\n== Load every application ==');
let loaded = 0, loadErr = [], docsTotal = 0, docsUnavailable = 0, multiFile = 0, dateUnparsed = 0;
const t0 = Date.now();
Object.keys(tokens).forEach(app => {
  const r = g.loadApplication(tokens[app], app);
  if (!r.ok) { loadErr.push(app + ': ' + r.code); return; }
  loaded++;
  r.data.documents.forEach(d => { docsTotal += d.files.length; docsUnavailable += d.files.filter(f => !f.available).length; if (d.files.length > 1) multiFile++; });
  r.data.sections.forEach(s => s.fields.forEach(f => { if (f.unparsedValue) dateUnparsed++; }));
});
say('  loaded ' + loaded + '/' + Object.keys(tokens).length + (loadErr.length ? ', errors: ' + loadErr.slice(0, 5).join('; ') : '') + ' (' + (Date.now() - t0) + ' ms in mock)');
say('  document links resolved: ' + docsTotal + ' (unavailable: ' + docsUnavailable + ', cells with several files: ' + multiFile + ')');
say('  date cells that could not be read as dates: ' + dateUnparsed);
check(loadErr.filter(e => badAppNos.indexOf(e.split(':')[0]) === -1).length === 0, 'some applications failed to load');

/* 5. Corrections on real rows ---------------------------------------------- */
say('\n== Corrections on real rows ==');
const byHeader = ctx.byHeader;
const editableOf = type => ctx.fields.filter(f => f.editable && f.type === type);
const goodApps = Object.keys(tokens).filter(a => badAppNos.indexOf(a) === -1);
const sampleApps = goodApps.slice(0, 25);
let submitted = 0;
sampleApps.forEach((app, n) => {
  const token = tokens[app];
  const before = g.readRow_(ctx, g.findRowByAppNo_(ctx, app));
  const a = g.loadApplication(token, app).data;
  const flat = [].concat(...a.sections.map(s => s.fields));
  const changes = {};
  const nameF = flat.find(f => f.role === 'name');
  if (nameF && nameF.editable) changes[nameF.key] = 'CORRECTED NAME ' + String.fromCharCode(65 + (n % 26));
  const pin = flat.find(f => f.type === 'pincode' && f.editable);
  if (pin) changes[pin.key] = '497' + String(100 + n).slice(-3);
  const num = flat.find(f => f.type === 'number' && f.editable && f.value);
  if (num) changes[num.key] = String(Number(num.value) + 1);
  const ta = flat.find(f => f.type === 'textarea' && f.editable);
  if (ta) changes[ta.key] = 'ग्राम परीक्षण, पोस्ट सूरजपुर, जिला सूरजपुर (छ.ग.)';
  const sel = flat.find(f => f.type === 'select' && f.editable && f.value && f.options.length > 1);
  if (sel) changes[sel.key] = sel.options.find(o => o !== sel.value);
  const replaced = [];
  const doc = a.documents.find(d => d.editable && d.files.length);
  if (doc && n % 3 === 0) {
    const up = g.uploadReplacement(token, app, doc.key, { fileName: 'new.pdf', mimeType: 'application/pdf', data: Buffer.from('%PDF-1.4 corrected').toString('base64') });
    if (check(up.ok, 'upload failed for ' + app + ': ' + up.message)) replaced.push(doc.key);
  }
  const r = g.submitCorrection(token, app, { version: a.version, declaration: true, changes, replacedDocuments: replaced });
  if (!check(r.ok, 'submit failed for ' + app + ': ' + JSON.stringify(r.details || r.message))) return;
  submitted++;
  const after = g.readRow_(ctx, before.rowIndex);
  const touched = new Set(Object.keys(changes).map(k => byHeader[k].colIndex).concat(replaced.map(k => byHeader[k].colIndex))
    .concat([cols.lastModified, cols.count, cols.status, cols.txn]));
  for (let c = 0; c < H.length; c++) {
    if (touched.has(c)) continue;
    const b = before.values[c], x = after.values[c];
    const same = (b instanceof Date || Object.prototype.toString.call(b) === '[object Date]') ? +b === +x : b === x;
    check(same && before.formulas[c] === after.formulas[c], app + ': untouched column ' + (c + 1) + ' changed');
  }
  check(after.formulas[cols.appNo] === before.formulas[cols.appNo], app + ': application number formula lost');
  if (nameF && nameF.editable) check(after.values[byHeader[nameF.key].colIndex] === changes[nameF.key], app + ': name not written');
  if (num) check(typeof after.values[byHeader[num.key].colIndex] === typeof before.values[byHeader[num.key].colIndex], app + ': number type changed');
});
say('  submitted ' + submitted + '/' + sampleApps.length + ' corrections; every other cell in those rows (incl. the application-number formula) unchanged');
const log2 = g.getLogSheet_();
say('  Correction_Log rows: ' + (log2.getLastRow() - 1));

/* 6. Field rules ------------------------------------------------------------ */
say('\n== Field rules ==');
const app0 = goodApps[30];
const a0 = g.loadApplication(tokens[app0], app0).data;
const flat0 = [].concat(...a0.sections.map(s => s.fields));
const tryChange = (key, value) => g.submitCorrection(tokens[app0], app0, { version: a0.version, declaration: true, changes: { [key]: value } });
const nameF0 = flat0.find(f => f.role === 'name');
let r = tryChange(nameF0.key, 'Rohan Singh');
say('  English-capital name "Rohan Singh": ' + (r.ok ? 'accepted' : 'rejected – ' + (r.details ? Object.values(r.details.fieldErrors)[0] : r.message)));
r = tryChange(H[cols.email], 'someone.else@gmail.com');
say('  change email: ' + (r.ok ? 'accepted' : 'rejected – ' + (r.details ? Object.values(r.details.fieldErrors)[0] : r.message)));
r = tryChange(H[cols.mobile], '12345');
say('  mobile "12345": ' + (r.ok ? 'accepted' : 'rejected – ' + (r.details ? Object.values(r.details.fieldErrors)[0] : r.message)));
r = tryChange(H[cols.appNo], 'GPS/SJP/2026/9999');
say('  change application number: ' + (r.ok ? 'accepted' : 'rejected – ' + (r.details ? Object.values(r.details.fieldErrors)[0] : r.message)));

/* 7. Admin ------------------------------------------------------------------ */
g.__state.props.set('NEW_ADMIN_USERNAME', 'admin');
g.__state.props.set('NEW_ADMIN_PASSWORD', 'Admin-Pass-123');
g.setupAdminAccount();
const at = g.adminLogin('admin', 'Admin-Pass-123').data.token;
const dash = g.adminGetDashboard(at).data.stats;
say('\n== Admin dashboard ==\n  ' + JSON.stringify(dash));
const d = g.adminGetDuplicates(at).data;
say('  duplicate groups – email: ' + d.email.length + ', mobile: ' + d.mobile.length + ', duplicate application numbers: ' + d.duplicateApplicationNumbers.join(', '));

say('\n' + (failures ? failures + ' CHECK(S) FAILED' : 'ALL CHECKS PASSED'));
if (OUT) fs.writeFileSync(path.join(OUT, 'sheet-test-report.txt'), log.join('\n'));
process.exitCode = failures ? 1 : 0;
