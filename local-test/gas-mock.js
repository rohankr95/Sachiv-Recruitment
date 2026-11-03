/**
 * Minimal in-memory mock of the Google Apps Script services used by the
 * portal, so the real .gs files can be executed and tested with Node.js.
 * This is a test aid only – it is NOT pushed to Apps Script (.claspignore).
 */
'use strict';
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const crypto = require('crypto');

const TZ = 'Asia/Kolkata';
const SRC = path.join(__dirname, '..');

function pad(n, w = 2) { return String(n).padStart(w, '0'); }
function parts(date, tz) {
  const f = new Intl.DateTimeFormat('en-GB', { timeZone: tz || TZ, year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit', hourCycle: 'h23' });
  const o = {};
  f.formatToParts(date).forEach(p => { o[p.type] = p.value; });
  return o;
}
function formatDate(date, tz, fmt) {
  const p = parts(date, tz);
  return fmt.replace('yyyy', p.year).replace('yy', p.year.slice(2)).replace('MM', p.month).replace('dd', p.day)
    .replace('HH', p.hour).replace('mm', p.minute).replace('ss', p.second);
}
const isDate = v => Object.prototype.toString.call(v) === '[object Date]';

/* ------------------------------ Sheets ------------------------------ */
class Sheet {
  constructor(ss, name) { this.ss = ss; this.name = name; this.cells = []; this.fmts = []; this.formulas = []; this.maxRows = 1000; this.maxCols = 26; this.protections = []; }
  getName() { return this.name; }
  setName(n) { this.name = n; return this; }
  _row(r) { while (this.cells.length < r) { this.cells.push([]); this.fmts.push([]); this.formulas.push([]); } }
  getLastRow() { for (let r = this.cells.length; r > 0; r--) if (this.cells[r - 1].some(v => v !== '' && v !== undefined && v !== null)) return r; return 0; }
  getLastColumn() { let m = 0; this.cells.forEach(row => row.forEach((v, i) => { if (v !== '' && v !== undefined && v !== null) m = Math.max(m, i + 1); })); return m; }
  getMaxRows() { return this.maxRows; }
  getMaxColumns() { return this.maxCols; }
  insertColumnsAfter(_, n) { this.maxCols += n; return this; }
  insertRowsAfter(_, n) { this.maxRows += n; return this; }
  getRange(r, c, nr = 1, nc = 1) {
    if (typeof r === 'string') throw new Error('A1 notation not supported in mock');
    if (r < 1 || c < 1 || r + nr - 1 > this.maxRows || c + nc - 1 > this.maxCols) throw new Error('The coordinates of the range are outside the dimensions of the sheet.');
    return new Range(this, r, c, nr, nc);
  }
  setFrozenRows() { return this; }
  autoResizeColumns() { return this; }
  getProtections() { return this.protections; }
  protect() {
    const p = { setDescription() { return p; }, addEditor() { return p; }, removeEditors() { return p; }, getEditors() { return []; }, canDomainEdit() { return false; }, setDomainEdit() { return p; } };
    this.protections.push(p); return p;
  }
  _get(r, c) { const row = this.cells[r - 1]; const v = row ? row[c - 1] : undefined; return v === undefined || v === null ? '' : v; }
  _set(r, c, v) {
    this._row(r);
    const fmt = this.fmts[r - 1][c - 1];
    this.formulas[r - 1][c - 1] = '';
    if (typeof v === 'string') {
      if (v.startsWith("'")) v = v.slice(1);
      else if (v.startsWith('=') && fmt !== '@') { this.formulas[r - 1][c - 1] = v; v = '[formula]'; }
      else if (fmt !== '@' && /^-?\d+(\.\d+)?$/.test(v) && !/^0\d/.test(v)) v = Number(v); // emulate auto-parse
    }
    this.cells[r - 1][c - 1] = v;
  }
  _display(r, c) {
    const v = this._get(r, c);
    const fmt = (this.fmts[r - 1] || [])[c - 1];
    if (isDate(v)) return formatDate(v, this.ss.tz, fmt === 'dd/MM/yyyy HH:mm:ss' ? 'dd/MM/yyyy HH:mm:ss' : (fmt === 'dd/MM/yyyy' ? 'dd/MM/yyyy' : (formatDate(v, this.ss.tz, 'HH:mm:ss') === '00:00:00' ? 'dd/MM/yyyy' : 'dd/MM/yyyy HH:mm:ss')));
    return String(v);
  }
}
class Range {
  constructor(sheet, r, c, nr, nc) { Object.assign(this, { sheet, r, c, nr, nc }); }
  _map(fn) { const out = []; for (let i = 0; i < this.nr; i++) { const row = []; for (let j = 0; j < this.nc; j++) row.push(fn(this.r + i, this.c + j)); out.push(row); } return out; }
  getValues() { return this._map((r, c) => { const v = this.sheet._get(r, c); return isDate(v) ? new Date(v.getTime()) : v; }); }
  getDisplayValues() { return this._map((r, c) => this.sheet._display(r, c)); }
  getFormulas() { return this._map((r, c) => (this.sheet.formulas[r - 1] || [])[c - 1] || ''); }
  getValue() { return this.getValues()[0][0]; }
  getDisplayValue() { return this.getDisplayValues()[0][0]; }
  getFormula() { return this.getFormulas()[0][0]; }
  getRow() { return this.r; }
  setValues(vals) {
    if (vals.length !== this.nr || vals.some(r => r.length !== this.nc)) throw new Error('setValues dimension mismatch');
    vals.forEach((row, i) => row.forEach((v, j) => this.sheet._set(this.r + i, this.c + j, v))); return this;
  }
  setValue(v) { return this.setValues(this._map(() => v)); }
  setFormula(f) { this.sheet._set(this.r, this.c, f); return this; }
  setNumberFormat(f) { this._map((r, c) => { this.sheet._row(r); this.sheet.fmts[r - 1][c - 1] = f; }); return this; }
  setFontWeight() { return this; }
  setBackground() { return this; }
  setDataValidation() { return this; }
  setNote() { return this; }
  createTextFinder(q) {
    const self = this; let entire = false, mcase = false;
    const tf = {
      matchEntireCell(b) { entire = b; return tf; }, matchCase(b) { mcase = b; return tf; },
      findAll() {
        const hits = [];
        for (let i = 0; i < self.nr; i++) for (let j = 0; j < self.nc; j++) {
          let d = self.sheet._display(self.r + i, self.c + j), s = String(q);
          if (!mcase) { d = d.toLowerCase(); s = s.toLowerCase(); }
          if (entire ? d === s : d.includes(s)) hits.push(new Range(self.sheet, self.r + i, self.c + j, 1, 1));
        }
        return hits;
      }
    };
    return tf;
  }
}
class Spreadsheet {
  constructor(id, name) { this.id = id; this.name = name; this.sheets = []; this.tz = TZ; }
  getId() { return this.id; }
  getUrl() { return 'https://docs.google.com/spreadsheets/d/' + this.id + '/edit'; }
  getSpreadsheetTimeZone() { return this.tz; }
  getSheetByName(n) { return this.sheets.find(s => s.name === n) || null; }
  getSheets() { return this.sheets; }
  insertSheet(n) { const s = new Sheet(this, n); this.sheets.push(s); return s; }
}

/* ------------------------------ Drive ------------------------------ */
function makeDrive(state) {
  let seq = 0;
  const newId = p => p + '_' + crypto.randomBytes(12).toString('hex') + (seq++);
  class Blob {
    constructor(bytes, type, name) { this.bytes = bytes; this.type = type; this.name = name; }
    getBytes() { return this.bytes.slice(); }
    getContentType() { return this.type; }
    getName() { return this.name; }
    getDataAsString() { return Buffer.from(this.bytes.map(b => b & 0xff)).toString('utf8'); }
    getAs(t) { if (t === this.type) return this; throw new Error('Conversion not supported'); }
  }
  class File {
    constructor(blob, folder) { this.id = newId('file'); this.blob = blob; this.name = blob.getName(); this.parents = [folder]; this.trashed = false; this.created = new Date(); this.desc = ''; state.files[this.id] = this; }
    getId() { return this.id; } getName() { return this.name; } setName(n) { this.name = n; return this; }
    getMimeType() { return this.blob.getContentType(); } getSize() { return this.blob.bytes.length; }
    getBlob() { return new Blob(this.blob.bytes, this.blob.type, this.name); }
    getAs(t) { return this.getBlob().getAs(t); }
    getUrl() { return 'https://drive.google.com/file/d/' + this.id + '/view?usp=drivesdk'; }
    setDescription(d) { this.desc = d; return this; }
    moveTo(f) { this.parents = [f]; return this; }
    makeCopy(name, folder) { const c = new File(new Blob(this.blob.bytes, this.blob.type, name), folder); return c; }
    getParents() { const a = this.parents.slice(); return { hasNext: () => a.length > 0, next: () => a.shift() }; }
    setTrashed(b) { this.trashed = b; return this; }
    getDateCreated() { return this.created; }
    getSharingAccess() { return 'PRIVATE'; }
  }
  class Folder {
    constructor(name) { this.id = newId('folder'); this.name = name; state.folders[this.id] = this; }
    getId() { return this.id; } getName() { return this.name; }
    createFile(blob) { return new File(blob, this); }
    createFolder(n) { return new Folder(n); }
    getFiles() { const a = Object.values(state.files).filter(f => f.parents.includes(this) && !f.trashed); return { hasNext: () => a.length > 0, next: () => a.shift() }; }
    getSharingAccess() { return 'PRIVATE'; }
  }
  const DriveApp = {
    Access: { ANYONE: 'ANYONE', ANYONE_WITH_LINK: 'ANYONE_WITH_LINK', PRIVATE: 'PRIVATE' },
    getFileById(id) { const f = state.files[id]; if (!f || f.trashed) throw new Error('No item with the given ID could be found.'); return f; },
    getFolderById(id) { const f = state.folders[id]; if (!f) throw new Error('No folder ' + id); return f; },
    createFolder(n) { return new Folder(n); }
  };
  return { DriveApp, Blob, Folder, File };
}

/* ------------------------------ Templates ------------------------------ */
function compileTemplate(src) {
  let code = 'let __o = "";\nwith (__data) {\n';
  const re = /<\?(!=|=)?([\s\S]*?)\?>/g;
  let last = 0, m;
  while ((m = re.exec(src)) !== null) {
    code += '__o += ' + JSON.stringify(src.slice(last, m.index)) + ';\n';
    if (m[1] === '!=') code += '__o += String(' + m[2].trim().replace(/;$/, '') + ');\n';
    else if (m[1] === '=') code += '__o += __esc(' + m[2].trim().replace(/;$/, '') + ');\n';
    else code += m[2] + '\n';
    last = re.lastIndex;
  }
  code += '__o += ' + JSON.stringify(src.slice(last)) + ';\n}\nreturn __o;';
  return code;
}

/* ------------------------------ Context ------------------------------ */
function createContext(opts = {}) {
  const state = { files: {}, folders: {}, spreadsheets: {}, cache: new Map(), props: new Map(), activeUser: opts.activeUser || 'owner@example.com', triggers: [] };
  const drive = makeDrive(state);
  const toSigned = buf => Array.from(buf, b => (b > 127 ? b - 256 : b));
  const toBuf = arr => Buffer.from(arr.map(b => b & 0xff));

  const Utilities = {
    DigestAlgorithm: { SHA_256: 'sha256' }, Charset: { UTF_8: 'utf8' },
    computeDigest(alg, text) { return toSigned(crypto.createHash(alg).update(String(text), 'utf8').digest()); },
    getUuid() { return crypto.randomUUID(); },
    formatDate,
    parseDate(str, tz, fmt) {
      if (fmt !== 'yyyy-MM-dd') throw new Error('mock parseDate supports yyyy-MM-dd');
      return new Date(str + 'T00:00:00+05:30');
    },
    base64Decode(s) { if (/[^A-Za-z0-9+/=\s]/.test(s)) throw new Error('Could not decode string.'); return toSigned(Buffer.from(s, 'base64')); },
    base64Encode(arr) { return toBuf(arr).toString('base64'); },
    newBlob(data, type, name) { const bytes = typeof data === 'string' ? toSigned(Buffer.from(data, 'utf8')) : data; return new drive.Blob(bytes, type || 'text/plain', name || 'blob'); },
    sleep() { }
  };
  const cache = {
    get: k => (state.cache.has(k) ? state.cache.get(k) : null),
    put: (k, v) => { if (String(v).length > 100000) throw new Error('Argument too large'); state.cache.set(k, String(v)); },
    remove: k => state.cache.delete(k),
    getAll: keys => { const o = {}; keys.forEach(k => { if (state.cache.has(k)) o[k] = state.cache.get(k); }); return o; },
    removeAll: keys => keys.forEach(k => state.cache.delete(k))
  };
  const props = {
    getProperty: k => (state.props.has(k) ? state.props.get(k) : null),
    setProperty: (k, v) => { state.props.set(k, String(v)); return props; },
    deleteProperty: k => { state.props.delete(k); return props; }
  };
  const SpreadsheetApp = {
    openById(id) { const s = state.spreadsheets[id]; if (!s) throw new Error('Spreadsheet not found ' + id); return s; },
    getActiveSpreadsheet() { return null; },
    create(name) { const id = 'ss_' + crypto.randomBytes(8).toString('hex'); const s = new Spreadsheet(id, name); s.insertSheet('Sheet1'); state.spreadsheets[id] = s; state.files[id] = { getSharingAccess: () => 'PRIVATE', moveTo() { }, trashed: false }; return s; },
    flush() { },
    newDataValidation() { const b = { requireValueInList: () => b, setAllowInvalid: () => b, build: () => ({}) }; return b; },
    ProtectionType: { SHEET: 'SHEET' }
  };
  const fileCache = {};
  const readSrc = n => {
    for (const ext of ['.html', '']) { const p = path.join(SRC, n + ext); if (fs.existsSync(p)) return fs.readFileSync(p, 'utf8'); }
    throw new Error('No HTML file named ' + n);
  };
  const escHtml = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  const output = content => {
    const o = { getContent: () => content, setTitle: () => o, addMetaTag: () => o, setXFrameOptionsMode: () => o };
    return o;
  };
  let ctx;
  const HtmlService = {
    XFrameOptionsMode: { DEFAULT: 'DEFAULT', ALLOWALL: 'ALLOWALL' },
    createHtmlOutputFromFile: n => output(readSrc(n)),
    createTemplateFromFile(n) {
      const tpl = { evaluate() {
        const code = fileCache[n] || (fileCache[n] = compileTemplate(readSrc(n)));
        const data = new Proxy(tpl, { has: (t, k) => typeof k === 'string' && (k in t || k in ctx), get: (t, k) => (k in t ? t[k] : (k === Symbol.unscopables ? undefined : ctx[k])) });
        const fn = new Function('__data', '__esc', code);
        return output(fn(data, escHtml));
      } };
      return tpl;
    }
  };
  ctx = {
    console: { log() { }, warn() { }, error: (...a) => { if (opts.verbose) console.error(...a); } },
    Utilities, SpreadsheetApp, DriveApp: drive.DriveApp, HtmlService,
    CacheService: { getScriptCache: () => cache },
    PropertiesService: { getScriptProperties: () => props },
    LockService: { getScriptLock: () => ({ tryLock: () => true, waitLock() { }, releaseLock() { } }) },
    Session: {
      getActiveUser: () => ({ getEmail: () => state.activeUser }),
      getEffectiveUser: () => ({ getEmail: () => 'owner@example.com' }),
      getScriptTimeZone: () => TZ
    },
    ScriptApp: {
      getProjectTriggers: () => state.triggers,
      newTrigger: fn => { const b = { timeBased: () => b, everyDays: () => b, atHour: () => b, create: () => { state.triggers.push({ getHandlerFunction: () => fn }); } }; return b; }
    },
    Date, JSON, Math, Object, Array, String, Number, RegExp, Error, parseInt, parseFloat, isNaN, Intl
  };
  vm.createContext(ctx);
  const files = fs.readdirSync(SRC).filter(f => f.endsWith('.gs')).sort();
  files.forEach(f => vm.runInContext(fs.readFileSync(path.join(SRC, f), 'utf8'), ctx, { filename: f }));
  ctx.__state = state;
  ctx.__drive = drive;
  return ctx;
}

module.exports = { createContext, formatDate, Spreadsheet, TZ };
