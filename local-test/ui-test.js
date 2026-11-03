/**
 * End-to-end UI walkthrough in Chromium against the mocked backend.
 *   node local-test/ui-test.js [screenshotDir]
 * Optional: UI_DEPS=/path/with/node_modules serves Bootstrap, Font Awesome
 * and pdf.js locally (for machines without CDN access).
 */
'use strict';
process.env.TZ = 'Asia/Kolkata';
const fs = require('fs');
const path = require('path');
const { chromium } = require('playwright');
const { createContext } = require('./gas-mock');

const OUT = process.argv[2] || path.join(__dirname, 'screenshots');
const DEPS = process.env.UI_DEPS ? path.join(process.env.UI_DEPS, 'node_modules') : null;
fs.mkdirSync(OUT, { recursive: true });

const g = createContext();
g.createSampleData();
g.setupPortal();
g.__state.props.set('NEW_ADMIN_USERNAME', 'admin');
g.__state.props.set('NEW_ADMIN_PASSWORD', 'Admin-Pass-123');
g.setupAdminAccount();

const ss = Object.values(g.__state.spreadsheets)[0];
const sheet = ss.getSheetByName('Form Responses 1');
const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
const cell = (row, h) => sheet.getRange(row, headers.indexOf(h) + 1).getValue();
const row26 = g.findRowByAppNo_(g.getSheetContext_(), 'APP202600026');
const EMAIL = cell(row26, 'Email Address');
const DOB = g.toIsoDate_(cell(row26, 'Date of Birth'));

const CDN = [
  [/bootstrap@5\.3\.3\/dist\/(.*)$/, m => 'bootstrap/dist/' + m[1]],
  [/font-awesome\/6\.5\.2\/(css|webfonts)\/(.*)$/, m => '@fortawesome/fontawesome-free/' + m[1] + '/' + m[2]],
  [/pdf\.js\/3\.11\.174\/(pdf(?:\.worker)?\.min\.js)$/, m => 'pdfjs-dist/build/' + m[1]]
];
const TYPES = { '.css': 'text/css', '.js': 'application/javascript', '.woff2': 'font/woff2', '.ttf': 'font/ttf' };

(async () => {
  const errors = [];
  const browser = await chromium.launch();

  async function newPage(viewport, isMobile) {
    const context = await browser.newContext({ viewport, deviceScaleFactor: 1, isMobile: !!isMobile, hasTouch: !!isMobile });
    const page = await context.newPage();
    page.on('pageerror', e => errors.push('pageerror: ' + e.message));
    page.on('console', m => { if (m.type() === 'error' && !/favicon|fonts\.g|ERR_FAILED|net::/.test(m.text())) errors.push('console: ' + m.text()); });
    await page.route('**/*', async route => {
      const url = route.request().url();
      if (url.startsWith('https://portal.test/')) {
        const u = new URL(url);
        return route.fulfill({ contentType: 'text/html', body: g.doGet({ parameter: Object.fromEntries(u.searchParams) }).getContent() });
      }
      for (const [re, map] of CDN) {
        const m = url.match(re);
        if (m && DEPS) {
          const file = path.join(DEPS, map(m));
          if (fs.existsSync(file)) return route.fulfill({ path: file, contentType: TYPES[path.extname(file)] || 'application/octet-stream', headers: { 'Access-Control-Allow-Origin': '*' } });
        }
      }
      if (DEPS) return route.abort();
      return route.continue();
    });
    await page.exposeFunction('__gasCall', (fn, argsJson) => {
      if (typeof g[fn] !== 'function' || /_$/.test(fn)) throw new Error('Unknown function ' + fn);
      return JSON.stringify(g[fn].apply(null, JSON.parse(argsJson)));
    });
    await page.addInitScript(() => {
      function runner(success, failure) {
        return new Proxy({}, {
          get(_, name) {
            if (name === 'withSuccessHandler') return fn => runner(fn, failure);
            if (name === 'withFailureHandler') return fn => runner(success, fn);
            return (...args) => {
              setTimeout(() => {
                window.__gasCall(name, JSON.stringify(args)).then(r => success && success(JSON.parse(r)), e => failure && failure(e));
              }, 30);
            };
          }
        });
      }
      window.google = { script: { run: runner(null, null) } };
    });
    return page;
  }
  const idle = page => page.waitForSelector('#loader.d-none', { state: 'attached' });
  const shot = (page, name, full) => page.screenshot({ path: path.join(OUT, name + '.png'), fullPage: !!full });

  /* ---------------- Applicant on a phone ---------------- */
  const m = await newPage({ width: 390, height: 844 }, true);
  await m.goto('https://portal.test/');
  await m.waitForSelector('#loginForm');
  await shot(m, '01-login-mobile');

  // wrong DOB
  await m.fill('#loginIdentifier', EMAIL);
  await m.fill('#loginDob', '1970-01-01');
  await m.click('#btnLogin');
  await m.waitForSelector('#loginError:not(.d-none)');
  if (!/do not match/.test(await m.textContent('#loginError'))) errors.push('wrong DOB message missing');

  await m.fill('#loginDob', DOB);
  await m.click('#btnLogin');
  await m.waitForSelector('#view-select:not(.d-none)');
  await idle(m);
  const items = await m.$$('.js-open-app');
  if (items.length !== 2) errors.push('expected 2 duplicate applications, got ' + items.length);
  if (!(await m.isVisible('#duplicateNotice'))) errors.push('duplicate notice not shown');
  await shot(m, '02-duplicate-list-mobile', true);

  await m.click('.js-open-app[data-app="APP202600026"]');
  await m.waitForSelector('#view-edit:not(.d-none)');
  await idle(m);
  await shot(m, '03-edit-form-mobile');

  // Desktop for the rest of the applicant flow
  const p = await newPage({ width: 1280, height: 900 });
  await p.goto('https://portal.test/');
  await p.waitForSelector('#loginForm');
  await p.fill('#loginIdentifier', EMAIL);
  await p.fill('#loginDob', DOB);
  await p.click('#btnLogin');
  await p.waitForSelector('#view-select:not(.d-none)');
  await p.click('.js-open-app[data-app="APP202600026"]');
  await p.waitForSelector('#view-edit:not(.d-none)');
  await idle(p);

  const fieldId = async label => p.evaluate(l => {
    const lab = Array.from(document.querySelectorAll('#formSections label.form-label')).find(x => x.textContent.trim().replace(/\*|Modified/g, '').trim() === l);
    return lab ? lab.getAttribute('for') : null;
  }, label);
  const nameId = await fieldId('Applicant Name');
  const mobileId = await fieldId('Mobile Number');
  const genderId = await fieldId('Gender');
  const pinId = await fieldId('Pin Code');
  await p.fill('#' + nameId, 'Ravi Kumar Sahu');
  await p.fill('#' + mobileId, '12345');
  await p.locator('#' + mobileId).blur();
  if (!(await p.textContent('#err-' + mobileId.split('-')[1])).includes('10-digit')) errors.push('mobile validation message missing');
  await p.click('#btnPreview');
  await p.waitForTimeout(400);
  if (!(await p.isVisible('#view-edit'))) errors.push('preview should be blocked by invalid mobile');
  await p.fill('#' + mobileId, '9876543210');
  await p.selectOption('#' + genderId, 'Female');
  await p.fill('#' + pinId, '497001');
  await p.locator('#' + pinId).blur();

  // Replace a document
  const pdfPath = path.join(OUT, 'replacement.pdf');
  fs.writeFileSync(pdfPath, Buffer.from(g.samplePdf_('Corrected domicile certificate').map(b => b & 0xff)));
  const domicileIdx = await p.evaluate(() => Array.from(document.querySelectorAll('.doc-card .fw-semibold')).findIndex(x => /Domicile Certificate/.test(x.textContent)));
  await p.setInputFiles('#docfile-' + domicileIdx, pdfPath);
  await p.waitForSelector('.doc-new');
  await idle(p);
  await shot(p, '04-edit-form-desktop-modified', true);

  // View the new document (pdf.js)
  await p.click('.doc-new .js-view-doc');
  await p.waitForSelector('#docModal.show');
  await p.waitForSelector('#docModalBody canvas, #docModalBody iframe', { timeout: 15000 });
  await p.waitForTimeout(800);
  await shot(p, '05-document-viewer');
  await p.click('#docModal .btn-close');
  await p.waitForSelector('#docModal', { state: 'hidden' });

  await p.click('#btnPreview');
  await p.waitForSelector('#view-preview:not(.d-none)');
  const modRows = await p.$$('#previewSections tr.is-modified');
  if (modRows.length !== 5) errors.push('expected 5 modified rows in preview (4 fields + 1 doc), got ' + modRows.length);
  await shot(p, '06-preview', true);

  await p.click('#btnConfirmSubmit');
  if (!(await p.textContent('#declarationError')).includes('declaration')) errors.push('declaration check missing');
  await p.check('#declaration');
  await p.click('#btnConfirmSubmit');
  await p.waitForSelector('#confirmModal.show');
  await p.waitForTimeout(300);
  await shot(p, '07-confirm-modal');
  await p.click('#btnFinalSubmit');
  await p.waitForSelector('#view-confirm:not(.d-none)');
  await idle(p);
  await shot(p, '08-confirmation');
  const txn = (await p.textContent('#confirmTxnId')).trim();
  if (!/^CRN-/.test(txn)) errors.push('no transaction id');
  if (cell(row26, 'Applicant Name') !== 'Ravi Kumar Sahu') errors.push('sheet not updated: ' + cell(row26, 'Applicant Name'));
  if (String(cell(row26, 'Mobile Number')) !== '9876543210') errors.push('mobile not updated');
  if (cell(row26, 'Correction Status') !== 'Modified') errors.push('status not updated');

  /* ---------------- Admin ---------------- */
  const a = await newPage({ width: 1366, height: 900 });
  await a.goto('https://portal.test/?page=admin');
  await a.waitForSelector('#loginForm');
  await a.fill('#adminUsername', 'admin');
  await a.fill('#adminPassword', 'Admin-Pass-123');
  await a.click('#btnLogin');
  await a.waitForSelector('#view-dashboard:not(.d-none)');
  await idle(a);
  await shot(a, '09-admin-dashboard', true);
  await a.fill('#searchInput', 'APP202600026');
  await a.click('#searchForm button[type=submit]');
  await a.waitForSelector('#view-detail:not(.d-none)');
  await idle(a);
  await shot(a, '10-admin-application', true);
  await a.click('button[data-bs-target="#dTabHistory"]');
  await a.waitForTimeout(400);
  await shot(a, '11-admin-history', true);
  await a.click('#btnBackDashboard');
  await a.waitForSelector('#view-dashboard:not(.d-none)');
  await a.click('#tabDuplicatesBtn');
  await a.waitForSelector('#duplicatesArea table', { timeout: 10000 });
  await shot(a, '12-admin-duplicates');

  await browser.close();
  if (errors.length) {
    console.log('UI TEST FAILED:\n - ' + errors.join('\n - '));
    process.exitCode = 1;
  } else {
    console.log('UI walkthrough passed. Screenshots in ' + OUT + ' (transaction ' + txn + ')');
  }
})().catch(e => { console.error(e); process.exit(1); });
