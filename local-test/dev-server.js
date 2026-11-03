/**
 * Local preview server: runs the real .gs code on top of the in-memory mock
 * and serves the real HTML pages, so you can click through the portal in a
 * browser without deploying to Google.
 *
 *   node local-test/dev-server.js                      # 2,500 synthetic records
 *   node local-test/dev-server.js /path/to/sheet.json  # your exported sheet
 *   PORT=9000 node local-test/dev-server.js
 *
 * Applicant portal: http://localhost:8080/
 * Admin dashboard : http://localhost:8080/?page=admin   (admin / Admin-Pass-123)
 *
 * Everything is in memory and is lost when the server stops. Drive documents
 * are placeholders. For local use only – never expose this server publicly.
 */
'use strict';
process.env.TZ = 'Asia/Kolkata';
const http = require('http');
const fs = require('fs');
const { createContext, loadSheetJson } = require('./gas-mock');

const PORT = Number(process.env.PORT || 8080);
const dataFile = process.argv[2];

const g = createContext({ verbose: true });
let sampleLogins = [];
if (dataFile) {
  loadSheetJson(g, JSON.parse(fs.readFileSync(dataFile, 'utf8')));
  console.log('Loaded ' + dataFile);
} else {
  g.createSampleData();
  console.log('Using 2,500 synthetic sample records');
}
console.log(g.setupPortal().split('\n').filter(l => /WARNING|NOTE|Records/.test(l)).join('\n'));
g.__state.props.set('NEW_ADMIN_USERNAME', 'admin');
g.__state.props.set('NEW_ADMIN_PASSWORD', 'Admin-Pass-123');
g.setupAdminAccount();
// From here on, behave like the deployed web app: visitors are anonymous,
// so owner-only functions (setupPortal, setupAdminAccount, ...) are refused.
g.__state.activeUser = '';

// A few logins to try.
const ctx = g.getSheetContext_();
const appNos = g.readColumn_(ctx, ctx.cols.appNo, true);
const emails = ctx.cols.email !== -1 ? g.readColumn_(ctx, ctx.cols.email, false) : [];
const mobiles = ctx.cols.mobile !== -1 ? g.readColumn_(ctx, ctx.cols.mobile, false) : [];
const dobs = g.readColumn_(ctx, ctx.cols.dob, false);
for (let i = 0; i < appNos.length && sampleLogins.length < 5; i++) {
  const dob = g.toIsoDate_(dobs[i]);
  if (!dob || /^#/.test(appNos[i])) continue;
  sampleLogins.push(appNos[i] + '  |  ' + (emails[i] || '-') + '  |  ' + (g.normalizeMobile_(mobiles[i]) || '-') + '  |  DOB ' + dob);
}

// Browser-side replacement for google.script.run: forwards calls to /__gas.
const SHIM = `<script>
(function () {
  function runner(success, failure) {
    return new Proxy({}, { get: function (_, name) {
      if (name === 'withSuccessHandler') return function (fn) { return runner(fn, failure); };
      if (name === 'withFailureHandler') return function (fn) { return runner(success, fn); };
      return function () {
        var args = Array.prototype.slice.call(arguments);
        fetch('/__gas', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ fn: name, args: args }) })
          .then(function (r) { if (!r.ok) throw new Error('HTTP ' + r.status); return r.json(); })
          .then(function (res) { if (success) success(res); })
          .catch(function (e) { if (failure) failure(e); });
      };
    } });
  }
  window.google = { script: { run: runner(null, null) } };
})();
</script>`;

const server = http.createServer((req, res) => {
  const url = new URL(req.url, 'http://localhost');
  if (req.method === 'POST' && url.pathname === '/__gas') {
    let body = '';
    req.on('data', c => { body += c; if (body.length > 20 * 1024 * 1024) req.destroy(); });
    req.on('end', () => {
      try {
        const { fn, args } = JSON.parse(body);
        // Same rule as Apps Script: functions ending in "_" are private.
        if (typeof fn !== 'string' || /_$/.test(fn) || typeof g[fn] !== 'function') throw new Error('Unknown function ' + fn);
        const result = g[fn].apply(null, Array.isArray(args) ? args : []);
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify(result === undefined ? null : result));
      } catch (e) {
        console.error(e);
        res.writeHead(500, { 'Content-Type': 'text/plain' });
        res.end(String(e.message || e));
      }
    });
    return;
  }
  if (req.method === 'GET' && url.pathname === '/') {
    const html = g.doGet({ parameter: Object.fromEntries(url.searchParams) }).getContent().replace('<head>', '<head>' + SHIM);
    res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
    res.end(html);
    return;
  }
  res.writeHead(404);
  res.end('Not found');
});

server.listen(PORT, '127.0.0.1', () => {
  console.log('\nApplicant portal : http://localhost:' + PORT + '/');
  console.log('Admin dashboard  : http://localhost:' + PORT + '/?page=admin   (admin / Admin-Pass-123)');
  console.log('\nTry these logins (Application No | Email | Mobile | DOB):\n  ' + sampleLogins.join('\n  '));
  console.log('\nData is in memory only and resets when you stop the server (Ctrl+C).');
});
