/**
 * Load test: many simulated applicants using the portal at the same time,
 * each in its own headless Chromium browser.
 *
 *   node local-test/load-test.js --url <web app URL> --data <sheet.json> [options]
 *
 * Options
 *   --url URL          deployed /exec URL (TEST deployment on a COPY of the sheet!)
 *                      or http://localhost:8080/ for the local dev server
 *   --data FILE        sheet.json from xlsx-to-json.py (used only to pick logins)
 *   --users N          total simulated applicants            (default 20)
 *   --concurrency N    how many run at the same moment       (default 10)
 *   --ramp SECONDS     spread the start of the first wave     (default 10)
 *   --mode MODE        login | view | submit                  (default view)
 *                        login : log in only
 *                        view  : log in + open application (+ open one document)
 *                        submit: view + change one field + preview + submit
 *   --field TEXT       label (or part of it) of the field to change in submit mode
 *                      (default "स्थान"); the value gets " (LT)" toggled on/off
 *   --login email|mobile   which identifier to log in with    (default email)
 *   --headed           show the browsers (debugging)
 *   --shots DIR        save a screenshot of every failure into DIR
 *   --deps DIR         serve Bootstrap/Font Awesome/pdf.js from DIR/node_modules
 *                      (only for machines without internet access to the CDNs)
 *
 * Output: per-step timings (median / 90th percentile / max), success counts
 * and grouped error messages. Requires: npm install playwright
 *
 * ⚠ submit mode WRITES to the sheet and Drive. Only use it against a test
 *   copy of the spreadsheet with its own Apps Script project and deployment.
 */
'use strict';
const fs = require('fs');
const { chromium } = require('playwright');

/* ------------------------------ arguments ------------------------------ */
const args = {};
process.argv.slice(2).forEach((a, i, arr) => {
  if (a.startsWith('--')) args[a.slice(2)] = arr[i + 1] && !arr[i + 1].startsWith('--') ? arr[i + 1] : true;
});
if (!args.url || !args.data) {
  console.log(fs.readFileSync(__filename, 'utf8').split('*/')[0]);
  process.exit(1);
}
const USERS = Number(args.users || 20);
const CONC = Number(args.concurrency || 10);
const RAMP = Number(args.ramp || 10) * 1000;
const MODE = args.mode || 'view';
const FIELD = args.field || 'स्थान';
const LOGIN_BY = args.login || 'email';
const STEP_TIMEOUT = 90000;

/* --------------------------- test credentials --------------------------- */
function pickLogins(file) {
  const d = JSON.parse(fs.readFileSync(file, 'utf8'));
  const H = d.values[0].map(h => String(h).trim().toLowerCase());
  const find = names => H.findIndex(h => names.some(n => h === n.toLowerCase()));
  const cApp = find(['Application No', 'Application Number', 'आवेदन क्रमांक']);
  const cEmail = find(['Email Address', 'Email', 'ईमेल']);
  const cMob = find(['मोबाइल नंबर', 'Mobile Number', 'Mobile No']);
  const cDob = find(['जन्म तिथि', 'Date of Birth', 'DOB']);
  if (cDob === -1 || (cEmail === -1 && cMob === -1)) throw new Error('Could not find DOB / email / mobile columns in ' + file);
  const out = [];
  d.values.slice(1).forEach(r => {
    const dob = r[cDob] && r[cDob].d ? r[cDob].d.slice(0, 10) : '';
    const app = String(r[cApp] || '');
    if (!dob || !app || app.startsWith('#')) return;
    const mobile = String(r[cMob] || '').replace(/\D/g, '').slice(-10);
    const id = LOGIN_BY === 'mobile' ? mobile : String(r[cEmail] || '').trim();
    if (id) out.push({ app, id, dob });
  });
  return out;
}

/* ------------------------------ statistics ------------------------------ */
const timings = {};   // step -> [ms]
const errors = {};    // message -> count
let completed = 0, failed = 0;
const record = (step, ms) => { (timings[step] = timings[step] || []).push(ms); };
const pct = (arr, p) => { const s = arr.slice().sort((a, b) => a - b); return s[Math.min(s.length - 1, Math.floor(p / 100 * s.length))]; };
const fmt = ms => (ms / 1000).toFixed(1) + 's';

/* Finds the frame holding the portal (Apps Script nests it inside two iframes). */
async function appFrame(page) {
  const deadline = Date.now() + STEP_TIMEOUT;
  while (Date.now() < deadline) {
    for (const f of page.frames()) {
      try { if (await f.$('#loginForm')) return f; } catch (e) { /* frame navigating */ }
    }
    await page.waitForTimeout(250);
  }
  throw new Error('Portal page did not load');
}
async function waitIdle(f) {
  await f.waitForSelector('#loader.d-none', { state: 'attached', timeout: STEP_TIMEOUT });
}
async function visibleError(f) {
  for (const sel of ['#loginError', '#globalAlert']) {
    const el = await f.$(sel + ':not(.d-none)');
    if (el) return (await el.textContent()).trim().slice(0, 160);
  }
  return '';
}

/* ------------------------------ one applicant ------------------------------ */
async function runUser(browser, cred, n) {
  const context = await browser.newContext({ viewport: { width: 390, height: 844 }, isMobile: true, hasTouch: true });
  if (args.deps) {
    const path = require('path');
    const CDN = [[/bootstrap@5\.3\.3\/dist\/(.*)$/, m => 'bootstrap/dist/' + m[1]], [/font-awesome\/6\.5\.2\/(css|webfonts)\/(.*)$/, m => '@fortawesome/fontawesome-free/' + m[1] + '/' + m[2]], [/pdf\.js\/3\.11\.174\/(pdf(?:\.worker)?\.min\.js)$/, m => 'pdfjs-dist/build/' + m[1]]];
    await context.route(/cdn\.jsdelivr\.net|cdnjs\.cloudflare\.com|fonts\.(googleapis|gstatic)\.com/, route => {
      const url = route.request().url();
      for (const [re, map] of CDN) { const m = url.match(re); if (m) return route.fulfill({ path: path.join(args.deps, 'node_modules', map(m)), headers: { 'Access-Control-Allow-Origin': '*' } }); }
      return route.abort();
    });
  }
  const page = await context.newPage();
  let step = 'open page';
  let ok = false;
  try {
    let t = Date.now();
    await page.goto(args.url, { timeout: STEP_TIMEOUT });
    const f = await appFrame(page);
    record('1 open page', Date.now() - t);

    step = 'login';
    await f.fill('#loginIdentifier', cred.id);
    await f.fill('#loginDob', cred.dob);
    t = Date.now();
    await f.click('#btnLogin');
    await f.waitForSelector('#view-select:not(.d-none), #view-edit:not(.d-none), #loginError:not(.d-none)', { timeout: STEP_TIMEOUT });
    const loginErr = await f.$('#loginError:not(.d-none)');
    if (loginErr) throw new Error('login: ' + (await loginErr.textContent()).trim());
    await waitIdle(f);
    record('2 login', Date.now() - t);

    if (MODE !== 'login') {
      step = 'open application';
      if (await f.$('#view-select:not(.d-none)')) {
        t = Date.now();
        await f.click('.js-open-app[data-app="' + cred.app.replace(/"/g, '\\"') + '"], .js-open-app');
        await f.waitForSelector('#view-edit:not(.d-none)', { timeout: STEP_TIMEOUT });
        await waitIdle(f);
        record('3 open application', Date.now() - t);
      } else {
        record('3 open application', 0); // opened automatically as part of login
      }
      const err = await visibleError(f);
      if (err) throw new Error(err);

      step = 'view document';
      const viewBtn = await f.$('.doc-card .js-view-doc');
      if (viewBtn) {
        t = Date.now();
        await viewBtn.click();
        await f.waitForSelector('#docModal.show', { timeout: STEP_TIMEOUT });
        await waitIdle(f);
        record('4 view document', Date.now() - t);
        await f.waitForTimeout(600); // let the open animation finish (Bootstrap ignores close clicks during it)
        await f.click('#docModal .btn-close');
        await f.waitForSelector('#docModal', { state: 'hidden', timeout: STEP_TIMEOUT });
      }
    }

    if (MODE === 'submit') {
      step = 'edit';
      const id = await f.evaluate(text => {
        const lab = Array.from(document.querySelectorAll('#formSections label.form-label[for]'))
          .find(l => l.textContent.indexOf(text) !== -1 && document.getElementById(l.getAttribute('for')).classList.contains('js-field'));
        return lab ? lab.getAttribute('for') : null;
      }, FIELD);
      if (!id) throw new Error('editable field containing "' + FIELD + '" not found (use --field)');
      const current = await f.inputValue('#' + id);
      await f.fill('#' + id, / \(LT\)$/.test(current) ? current.replace(/ \(LT\)$/, '') : (current || 'Test') + ' (LT)');
      await f.locator('#' + id).blur();
      await f.click('#btnPreview');
      await f.waitForSelector('#view-preview:not(.d-none)', { timeout: STEP_TIMEOUT });
      await f.locator('#declaration').scrollIntoViewIfNeeded();
      await f.evaluate(() => document.getElementById('declaration').scrollIntoView({ block: 'center' })); // keep it clear of the sticky action bar
      await f.check('#declaration');
      await f.click('#btnConfirmSubmit');
      await f.waitForSelector('#confirmModal.show', { timeout: STEP_TIMEOUT });

      step = 'submit';
      t = Date.now();
      await f.click('#btnFinalSubmit');
      await f.waitForSelector('#view-confirm:not(.d-none), #globalAlert:not(.d-none)', { timeout: STEP_TIMEOUT });
      const subErr = await f.$('#globalAlert:not(.d-none)');
      if (subErr) throw new Error('submit: ' + (await subErr.textContent()).trim().slice(0, 160));
      await waitIdle(f);
      record('5 submit', Date.now() - t);
    }

    step = 'logout';
    const logout = await f.$('#btnLogout:not(.d-none)');
    if (logout) {
      f.page().once('dialog', d => d.accept());
      await logout.click();
    }
    completed++;
    ok = true;
  } catch (e) {
    failed++;
    let shown = '';
    try {
      for (const fr of page.frames()) {
        const t = await fr.$$eval('.toast-body, #globalAlert:not(.d-none), #loginError:not(.d-none), .invalid-feedback', els => els.map(x => x.textContent.trim()).filter(Boolean).join(' | ')).catch(() => '');
        if (t) shown += t;
      }
      if (args.shots) await page.screenshot({ path: args.shots + '/fail-' + n + '.png', fullPage: true });
    } catch (x) { /* ignore */ }
    const msg = '[' + step + '] ' + String(e.message || e).split('\n')[0].replace(/\d{6}-\d{4}/g, '…').slice(0, 160) + (shown ? ' — on screen: ' + shown.slice(0, 200) : '');
    errors[msg] = (errors[msg] || 0) + 1;
  } finally {
    await context.close().catch(() => {});
    process.stdout.write(ok ? '.' : 'x');
  }
}

/* ------------------------------ main ------------------------------ */
(async () => {
  const logins = pickLogins(args.data);
  if (logins.length < USERS) console.log('Only ' + logins.length + ' usable logins – some will be reused.');
  // Spread users over the sheet so they are different applicants.
  const step = Math.max(1, Math.floor(logins.length / USERS));
  const queue = Array.from({ length: USERS }, (_, i) => logins[(i * step) % logins.length]);

  console.log('Load test: ' + USERS + ' users, ' + CONC + ' at a time, mode=' + MODE + ', url=' + args.url);
  if (MODE === 'submit') console.log('⚠ submit mode writes to the sheet – make sure this is a TEST copy.');
  const browser = await chromium.launch({ headless: !args.headed });
  const started = Date.now();
  let next = 0;
  async function worker(w) {
    await new Promise(r => setTimeout(r, Math.round(w * RAMP / CONC)));
    while (next < queue.length) {
      const n = next++;
      await runUser(browser, queue[n], n);
    }
  }
  await Promise.all(Array.from({ length: Math.min(CONC, USERS) }, (_, w) => worker(w)));
  await browser.close();
  const total = (Date.now() - started) / 1000;

  console.log('\n\n=== Results ===');
  console.log('Completed: ' + completed + '/' + USERS + '   Failed: ' + failed + '   Wall time: ' + total.toFixed(0) + 's   Throughput: ' + (completed / total * 60).toFixed(1) + ' applicants/min');
  console.log('\nStep                  count   median    90%      max');
  Object.keys(timings).sort().forEach(k => {
    const a = timings[k];
    console.log(k.padEnd(20) + String(a.length).padStart(7) + fmt(pct(a, 50)).padStart(9) + fmt(pct(a, 90)).padStart(9) + fmt(Math.max(...a)).padStart(9));
  });
  if (Object.keys(errors).length) {
    console.log('\nErrors:');
    Object.entries(errors).sort((a, b) => b[1] - a[1]).forEach(([m, c]) => console.log('  ' + String(c).padStart(4) + ' × ' + m));
  }
})().catch(e => { console.error(e); process.exit(1); });
