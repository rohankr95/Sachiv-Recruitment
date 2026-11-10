/**
 * Authentication.gs
 * ---------------------------------------------------------------------------
 * Applicant login (registered email OR mobile number + date of birth),
 * session management and rate limiting.
 *
 * Sessions live in CacheService (server side). The browser only holds a
 * random 64-character token; it never receives a list of applications it is
 * not entitled to. Every public function re-validates the token and checks
 * that the requested application number belongs to the session.
 * ---------------------------------------------------------------------------
 */

var SESSION_PREFIX_ = 'SESS_';
var STAGED_PREFIX_ = 'STG_';
var RATE_PREFIX_ = 'RL_';

/* ======================================================================== */
/* Public API                                                               */
/* ======================================================================== */

/**
 * Applicant login.
 * @param {string} identifier registered email address or mobile number
 * @param {string} dob        date of birth, yyyy-mm-dd (from <input type=date>)
 * @return {{ok:boolean, data:{token:string, idleTimeoutSeconds:number, applications:Array}}}
 */
function login(identifier, dob) {
  return respond_(function () {
    if (!isPortalOpen_() && CONFIG.CLOSED_ALLOW_LOGIN !== true) {
      throw portalError_('PORTAL_CLOSED', CONFIG.PORTAL_CLOSED_MESSAGE);
    }
    var parsed = parseIdentifier_(identifier);
    var dobIso = parseDobInput_(dob);

    var limiter = checkRateLimit_('login', parsed.kind + ':' + parsed.value, CONFIG.RATE_LIMIT.LOCKOUT_MINUTES);

    var ctx = getSheetContext_();
    var idCol = parsed.kind === 'email' ? ctx.cols.email : ctx.cols.mobile;
    if (idCol === -1) {
      throw portalError_('VALIDATION', 'Login using ' + (parsed.kind === 'email' ? 'email address' : 'mobile number') +
        ' is not available. Please use your registered ' + (parsed.kind === 'email' ? 'mobile number' : 'email address') + '.');
    }

    // Batch read only the columns needed for verification (not the whole sheet).
    var idValues = readColumn_(ctx, idCol, false);
    var candidates = [];
    for (var i = 0; i < idValues.length; i++) {
      var v = parsed.kind === 'email' ? normalizeEmail_(idValues[i]) : normalizeMobile_(idValues[i]);
      if (v && v === parsed.value) candidates.push(i);
    }

    var matched = [];
    if (candidates.length) {
      var dobValues = readColumn_(ctx, ctx.cols.dob, false);
      var dobDisplay = readColumn_(ctx, ctx.cols.dob, true);
      matched = candidates.filter(function (idx) {
        var iso = toIsoDate_(dobValues[idx]) || toIsoDate_(dobDisplay[idx]);
        return iso === dobIso;
      });
    }

    if (!matched.length) {
      var remaining = registerFailure_(limiter, CONFIG.RATE_LIMIT.MAX_FAILED_ATTEMPTS, CONFIG.RATE_LIMIT.LOCKOUT_MINUTES);
      Utilities.sleep(CONFIG.RATE_LIMIT.FAILURE_DELAY_MS);
      throw portalError_('AUTH_FAILED', remaining > 0
        ? 'The details entered do not match our records. ' + remaining + ' attempt(s) remaining before a temporary lock.'
        : 'Too many failed attempts. Login for this email/mobile is locked for ' + CONFIG.RATE_LIMIT.LOCKOUT_MINUTES + ' minutes.');
    }

    var accessibleIdx = (CONFIG.DUPLICATE_IDENTIFIER_POLICY === 'SHOW_ALL' ? candidates : matched).filter(function (idx) {
      return !isExcludedRow_(dobValues[idx], dobDisplay[idx]); // never expose incomplete (excluded) rows
    });
    var appNoValues = readColumn_(ctx, ctx.cols.appNo, true);

    // Duplicate detection: how many rows share this applicant's email / mobile.
    var emailValues = ctx.cols.email !== -1 ? (parsed.kind === 'email' ? idValues : readColumn_(ctx, ctx.cols.email, false)) : [];
    var mobileValues = ctx.cols.mobile !== -1 ? (parsed.kind === 'mobile' ? idValues : readColumn_(ctx, ctx.cols.mobile, false)) : [];
    var emailCounts = countOccurrences_(emailValues, normalizeEmail_);
    var mobileCounts = countOccurrences_(mobileValues, normalizeMobile_);
    var appNoCounts = countOccurrences_(appNoValues, function (x) { return String(x).trim().toLowerCase(); });

    var apps = [];
    var dup = {};
    accessibleIdx.forEach(function (idx) {
      var appNo = String(appNoValues[idx] || '').trim();
      if (!appNo || apps.indexOf(appNo) !== -1) return;
      apps.push(appNo);
      var em = emailValues.length ? normalizeEmail_(emailValues[idx]) : '';
      var mo = mobileValues.length ? normalizeMobile_(mobileValues[idx]) : '';
      dup[appNo] = {
        email: em ? (emailCounts[em] || 0) : 0,
        mobile: mo ? (mobileCounts[mo] || 0) : 0,
        record: appNoCounts[appNo.toLowerCase()] || 0,
        dobMatched: matched.indexOf(idx) !== -1
      };
    });

    if (!apps.length) {
      throw portalError_('AUTH_FAILED', 'Your record was found but it has no application number. Please contact the helpdesk.');
    }

    clearRateLimit_(limiter);
    var token = createSession_('applicant', {
      apps: apps,
      dup: dup,
      who: parsed.kind + ' login: ' + maskIdentifier_(parsed.value),
      uploads: 0,
      staged: []
    }, CONFIG.SESSION.IDLE_TIMEOUT_MINUTES);

    return {
      token: token,
      idleTimeoutSeconds: CONFIG.SESSION.IDLE_TIMEOUT_MINUTES * 60,
      loginType: parsed.kind,
      maskedIdentifier: maskIdentifier_(parsed.value),
      applications: summarizeApplications_(ctx, apps, dup, appNoValues)
    };
  });
}

/**
 * Lists the applications the current session may access (used after a page
 * refresh and after a successful submission).
 */
function getMyApplications(token) {
  return respond_(function () {
    var session = getSession_(token, 'applicant');
    var ctx = getSheetContext_();
    return {
      idleTimeoutSeconds: CONFIG.SESSION.IDLE_TIMEOUT_MINUTES * 60,
      maskedIdentifier: String(session.who || '').replace(/^.*?:\s*/, ''),
      applications: summarizeApplications_(ctx, session.apps, session.dup, null)
    };
  });
}

/** Extends the session; returns the remaining idle time. */
function keepAlive(token) {
  return respond_(function () {
    getSession_(token, 'applicant');
    return { idleTimeoutSeconds: CONFIG.SESSION.IDLE_TIMEOUT_MINUTES * 60 };
  });
}

/** Ends the session and trashes any uploads that were never submitted. */
function logout(token) {
  return respond_(function () {
    var session = peekSession_(token);
    if (session && session.type === 'applicant' && session.staged && session.staged.length) {
      discardStagedUploads_(session.staged);
    }
    destroySession_(token);
    return { loggedOut: true };
  });
}

/* ======================================================================== */
/* Input parsing                                                            */
/* ======================================================================== */

function parseIdentifier_(identifier) {
  var raw = String(identifier == null ? '' : identifier).trim();
  if (!raw) throw portalError_('VALIDATION', 'Enter your registered email address or mobile number.');
  if (raw.length > 254) throw portalError_('VALIDATION', 'The email address or mobile number is too long.');
  if (raw.indexOf('@') !== -1) {
    var email = normalizeEmail_(raw);
    if (!isValidEmail_(email)) throw portalError_('VALIDATION', 'Enter a valid email address.');
    return { kind: 'email', value: email };
  }
  if (/[A-Za-z]/.test(raw)) throw portalError_('VALIDATION', 'Enter a valid email address or 10-digit mobile number.');
  var mobile = normalizeMobile_(raw);
  if (!isValidMobile_(mobile)) throw portalError_('VALIDATION', 'Enter a valid 10-digit mobile number.');
  return { kind: 'mobile', value: mobile };
}

function parseDobInput_(dob) {
  var s = String(dob == null ? '' : dob).trim();
  var iso = /^\d{4}-\d{2}-\d{2}$/.test(s) ? buildIsoDate_(s.slice(0, 4), s.slice(5, 7), s.slice(8, 10)) : toIsoDate_(s);
  if (!iso) throw portalError_('VALIDATION', 'Enter a valid date of birth.');
  var today = Utilities.formatDate(new Date(), getTimeZone_(), 'yyyy-MM-dd');
  if (iso > today || iso < '1900-01-01') throw portalError_('VALIDATION', 'Enter a valid date of birth.');
  return iso;
}

function countOccurrences_(values, normalizer) {
  var counts = {};
  for (var i = 0; i < values.length; i++) {
    var k = normalizer(values[i]);
    if (!k) continue;
    counts[k] = (counts[k] || 0) + 1;
  }
  return counts;
}

/**
 * Builds the list shown on the "Verify Application" step: number, name,
 * status, correction count and duplicate flags. Only reads 4 columns.
 */
function summarizeApplications_(ctx, apps, dup, appNoValues) {
  var appNos = appNoValues || readColumn_(ctx, ctx.cols.appNo, true);
  var names = ctx.cols.name !== -1 ? readColumn_(ctx, ctx.cols.name, true) : [];
  var statuses = readColumn_(ctx, ctx.cols.status, true);
  var counts = readColumn_(ctx, ctx.cols.count, false);
  var lastMod = readColumn_(ctx, ctx.cols.lastModified, true);
  var index = {};
  for (var i = 0; i < appNos.length; i++) {
    var a = String(appNos[i] || '').trim();
    if (a && index[a] === undefined) index[a] = i;
  }
  return apps.map(function (appNo) {
    var idx = index[appNo];
    var d = (dup && dup[appNo]) || { email: 0, mobile: 0, record: 0, dobMatched: true };
    return {
      applicationNumber: appNo,
      applicantName: idx !== undefined && names.length ? String(names[idx] || '') : '',
      status: idx !== undefined ? String(statuses[idx] || '') || 'Not corrected' : 'Not found',
      correctionCount: idx !== undefined ? (Number(counts[idx]) || 0) : 0,
      lastModified: idx !== undefined ? String(lastMod[idx] || '') : '',
      duplicateEmail: d.email > 1,
      emailShareCount: d.email,
      duplicateMobile: d.mobile > 1,
      mobileShareCount: d.mobile,
      duplicateRecord: d.record > 1,
      dobMatched: d.dobMatched !== false
    };
  });
}

/* ======================================================================== */
/* Sessions                                                                 */
/* ======================================================================== */

/**
 * Creates a server-side session and returns its token.
 * @param {string} type 'applicant' | 'admin'
 * @param {Object} data session payload
 * @param {number} idleMinutes idle timeout
 */
function createSession_(type, data, idleMinutes) {
  var token = randomToken_();
  var now = Date.now();
  var session = { type: type, created: now, last: now, idle: idleMinutes * 60 };
  Object.keys(data || {}).forEach(function (k) { session[k] = data[k]; });
  putSession_(token, session);
  return token;
}

function putSession_(token, session) {
  var ttl = Math.max(60, Math.min(21600, session.idle || 900));
  CacheService.getScriptCache().put(SESSION_PREFIX_ + token, JSON.stringify(session), ttl);
}

function isValidTokenFormat_(token) {
  return typeof token === 'string' && /^[a-f0-9]{64}$/.test(token);
}

/** Reads a session without extending it. Returns null if absent. */
function peekSession_(token) {
  if (!isValidTokenFormat_(token)) return null;
  var raw = CacheService.getScriptCache().get(SESSION_PREFIX_ + token);
  if (!raw) return null;
  try { return JSON.parse(raw); } catch (e) { return null; }
}

/**
 * Validates a token, enforces idle and absolute expiry, slides the idle
 * timeout and returns the session. Throws SESSION_EXPIRED otherwise.
 */
function getSession_(token, type) {
  var session = peekSession_(token);
  if (!session || session.type !== type) {
    throw portalError_('SESSION_EXPIRED', 'Your session has expired. Please log in again.');
  }
  var now = Date.now();
  var absoluteMinutes = type === 'admin' ? CONFIG.SESSION.ADMIN_ABSOLUTE_TIMEOUT_MINUTES : CONFIG.SESSION.ABSOLUTE_TIMEOUT_MINUTES;
  if (now - session.created > absoluteMinutes * 60000 || now - session.last > session.idle * 1000) {
    destroySession_(token);
    throw portalError_('SESSION_EXPIRED', 'Your session has expired. Please log in again.');
  }
  session.last = now;
  putSession_(token, session);
  return session;
}

function destroySession_(token) {
  if (isValidTokenFormat_(token)) CacheService.getScriptCache().remove(SESSION_PREFIX_ + token);
}

/** Throws unless the session is allowed to access this application. */
function requireApplicationAccess_(session, appNo) {
  if (!session || !Array.isArray(session.apps) || session.apps.indexOf(String(appNo)) === -1) {
    console.warn('Blocked access attempt to application ' + appNo + ' by ' + (session ? session.who : 'unknown'));
    throw portalError_('FORBIDDEN', 'You are not authorised to access this application.');
  }
}

/* ======================================================================== */
/* Rate limiting (per email / mobile / admin username)                      */
/* ======================================================================== */

/**
 * Throws RATE_LIMITED while the key is locked. Returns a limiter handle used
 * by registerFailure_ / clearRateLimit_.
 */
function checkRateLimit_(scope, key, lockMinutes) {
  var cacheKey = RATE_PREFIX_ + scope + '_' + sha256Hex_(scope + '|' + key).slice(0, 48);
  var raw = CacheService.getScriptCache().get(cacheKey);
  var state = raw ? JSON.parse(raw) : { n: 0, until: 0 };
  if (state.until && state.until > Date.now()) {
    var mins = Math.ceil((state.until - Date.now()) / 60000);
    throw portalError_('RATE_LIMITED', 'Too many failed attempts. Please try again after ' + mins + ' minute(s).');
  }
  return { key: cacheKey, state: state, lockMinutes: lockMinutes };
}

/** Records a failed attempt. Returns the number of attempts left. */
function registerFailure_(limiter, maxAttempts, lockMinutes) {
  var state = limiter.state;
  state.n = (state.n || 0) + 1;
  var remaining = maxAttempts - state.n;
  if (state.n >= maxAttempts) {
    state.until = Date.now() + lockMinutes * 60000;
    state.n = 0;
    remaining = 0;
    console.warn('Rate limit lock applied: ' + limiter.key);
  }
  CacheService.getScriptCache().put(limiter.key, JSON.stringify(state), Math.min(21600, Math.max(lockMinutes * 60, 900)));
  return remaining;
}

function clearRateLimit_(limiter) {
  CacheService.getScriptCache().remove(limiter.key);
}
