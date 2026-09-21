// ========== Token authentication ==========
function generateToken() {
  // CSPRNG-backed examiner token (was Math.random, which is predictable). Three
  // UUIDs of hex → 96 hex chars, well above the prior 48-char entropy.
  return (Utilities.getUuid() + Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
}

// r23: a valid verdict is cached for TOKEN_VERDICT_CACHE_SEC. Every examiner
// poll (2-5 s) used to re-read the whole 'בוחנים' sheet just to check a token
// that had been valid a few seconds earlier. Only positives are cached: a token
// that was just created must be usable at once, and one that was evicted or
// expired may linger for at most a minute.
var TOKEN_VERDICT_CACHE_SEC = 60;
function verifyToken(examinerId, token) {
  if (!examinerId || !token) return false;
  var key = QUESTION_CACHE_PREFIX + 'tok_' + normalizeId(examinerId) + '_' + String(token).slice(0, 80), cache = null;
  try { cache = CacheService.getScriptCache(); if (cache.get(key) === '1') return true; } catch (eGet) { cache = null; }
  var sheet = getSheet('בוחנים');
  var data = sheet.getDataRange().getValues();
  var valid = false;
  for (var i = 1; i < data.length; i++) {
    if (normalizeId(data[i][1]) === normalizeId(examinerId)) {
      var storedTokens = String(data[i][6] || '').split(',');
      var expiry = data[i][7];
      if (storedTokens.indexOf(token) === -1) break;
      if (!expiry) break;
      var expiryDate = expiry instanceof Date ? expiry : new Date(expiry);
      if (new Date() > expiryDate) break;
      valid = true;
      break;
    }
  }
  if (valid) { try { if (!cache) cache = CacheService.getScriptCache(); cache.put(key, '1', TOKEN_VERDICT_CACHE_SEC); } catch (ePut) {} }
  return valid;
}

function requireToken(p) {
  var valid = verifyToken(p.examinerId, p.token);
  diagMark('auth:token');   // a full 'בוחנים' read, paid by every examiner poll before its handler starts
  if (!valid) {
    return jsonResponse({ status: 'error', message: 'טוקן לא תקין — יש להתחבר מחדש', tokenExpired: true });
  }
  return null;
}

// ========== Origin allowlist (soft check) ==========
// Apps Script cannot read HTTP headers, so the client must pass an `origin` parameter.
// This is bypassable (anyone reading the HTML sees the magic string) but blocks
// casual API exploration / generic scrapers / curl scripts. Real security comes
// from token enforcement.
var ALLOWED_ORIGINS = [
  'examiner-app',      // examiner.html
  'examinee-app',      // examinee.html / exam.html (standalone practice)
  'teacher-app',       // teacher.html
  'student-app',       // student.html
  'admin-app',         // admin.html
  'bohanyzahal-site',  // bohan-site (IDF portal — server-side auth)
  'localhost-dev'      // local development
];
function checkOrigin(p) {
  // Allowed: actions called from external services (none currently) or no origin enforcement on certain reads
  // For now: reject unknown origins on all actions except 'viewResult' (HTML output, opened in browser tab).
  var action = p.action || '';
  if (action === 'viewResult' || action === '') return null;
  var origin = String(p.origin || '').trim();
  if (!origin) {
    return jsonResponse({ status: 'error', message: 'Missing origin', code: 'origin_required' });
  }
  if (ALLOWED_ORIGINS.indexOf(origin) === -1) {
    return jsonResponse({ status: 'error', message: 'Unauthorized origin', code: 'origin_denied' });
  }
  return null;
}

// ========== Rate limiting (Stage 2b) ==========
// Sliding-window rate limit backed by CacheService. The cache stores a JSON
// array of recent request timestamps per (action, identifier). On each call we
// filter to the window, count, and either allow + append, or reject. CacheService
// auto-evicts entries by TTL — no manual cleanup needed.
//
// Trade-off note: Apps Script doesn't expose client IP, so identifiers must come
// from the request payload (sessionCode, idNumber, examinerId). This means an
// attacker who varies the identifier can avoid limits — but they'd still need
// valid creds to be useful, since the data still has to go through token and
// origin checks. This rate limit primarily defends quotas + same-target floods.
function checkRateLimit(action, identifier, maxRequests, windowSeconds) {
  try {
    var cache = CacheService.getScriptCache();
    var key = 'rl_' + action + '_' + identifier;
    var raw = cache.get(key);
    var now = Date.now();
    var windowMs = windowSeconds * 1000;
    var timestamps = [];
    if (raw) {
      try { timestamps = JSON.parse(raw) || []; } catch(_e) { timestamps = []; }
    }
    // Drop timestamps older than the window
    var fresh = [];
    for (var i = 0; i < timestamps.length; i++) {
      if ((now - timestamps[i]) < windowMs) fresh.push(timestamps[i]);
    }
    if (fresh.length >= maxRequests) {
      var oldest = fresh[0];
      var waitSec = Math.max(1, Math.ceil((windowMs - (now - oldest)) / 1000));
      return { ok: false, waitSec: waitSec };
    }
    fresh.push(now);
    cache.put(key, JSON.stringify(fresh), windowSeconds + 60); // TTL slightly over window
    return { ok: true };
  } catch (e) {
    // If the cache is unavailable, fail open — don't break the system.
    return { ok: true };
  }
}

// Convenience wrapper that returns a jsonResponse error on hit, or null on pass.
function requireRateLimit(action, identifier, maxRequests, windowSeconds) {
  if (!identifier) return null; // can't enforce without an identifier
  var result = checkRateLimit(action, identifier, maxRequests, windowSeconds || 60);
  if (!result.ok) {
    return jsonResponse({
      status: 'error',
      message: 'יותר מדי בקשות. נסה שוב בעוד ' + result.waitSec + ' שניות.',
      rateLimited: true,
      waitSec: result.waitSec
    });
  }
  return null;
}

// ========== Examinee token (Stage 1c) ==========
// Each examinee receives a random token at registration time. All subsequent
// examinee-side calls (poll, submit, self-DQ) must echo that token. Prevents
// an attacker registered to the same session from acting on a victim's row
// using only sessionCode + idNumber.
function generateExamineeToken() {
  // CSPRNG-backed (was Math.random, predictable). Two UUIDs of hex → 64 hex chars.
  return (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
}

// Returns { valid: bool, legacy: bool, reason: string }
// - legacy: true when the stored row predates token support (empty cell) —
//   we accept the call but flag it so we can audit / tighten later.
// - reason values (when invalid): 'not_found', 'missing', 'mismatch'.
function verifyExamineeToken(sessionCode, idNumber, examineeToken) {
  var data = readPendingTail().rows;   // read-only: no row-index writes here
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === String(sessionCode) && normalizeId(data[i][1]) === normalizeId(idNumber)) {
      // Per-examinee audio (column J) rides along on the row we already read, so
      // callers get it without a second sheet scan — matters on getExamQuestions,
      // which is the exam-start hot path.
      var rowAudio = String(data[i][9] || '').trim() === 'on' ? 'on' : 'off';
      var storedToken = String((data[i].length > 12 ? data[i][12] : '') || '').trim();
      if (!storedToken) return { valid: true, legacy: true, audioMode: rowAudio };
      if (!examineeToken) return { valid: false, reason: 'missing' };
      if (String(examineeToken).trim() === storedToken) return { valid: true, legacy: false, audioMode: rowAudio };
      return { valid: false, reason: 'mismatch' };
    }
  }
  return { valid: false, reason: 'not_found' };
}

// Convenience wrapper for handlers. Returns null when OK, or a jsonResponse error.
function requireExamineeToken(p) {
  if (!p.sessionCode || !p.idNumber) {
    return jsonResponse({ status: 'error', message: 'חסרים פרטי נבחן' });
  }
  var result = verifyExamineeToken(p.sessionCode, p.idNumber, p.examineeToken);
  if (!result.valid) {
    return jsonResponse({ status: 'error', message: 'טוקן נבחן לא תקין', examineeTokenError: result.reason });
  }
  return null;
}

// Look up the list of sites a center commander oversees (column K in בוחנים).
// Format in the sheet: comma-separated site names matching column K in תוצאות.
// Returns array of trimmed names (empty if not found / empty cell).
function getExaminerManagedSites(examinerId) {
  if (!examinerId) return [];
  var sheet = getSheet('בוחנים');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (normalizeId(data[i][1]) === normalizeId(examinerId)) {
      var raw = (data[i].length > 10) ? String(data[i][10] || '') : '';
      if (!raw) return [];
      return raw.split(',').map(function(s) { return s.trim(); }).filter(function(s) { return s; });
    }
  }
  return [];
}

// Look up examiner's role ('בוחן' or 'מפקד'). Returns '' if not found.
function getExaminerRole(examinerId) {
  if (!examinerId) return '';
  var sheet = getSheet('בוחנים');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (normalizeId(data[i][1]) === normalizeId(examinerId)) {
      return String(data[i][5] || 'בוחן').trim();
    }
  }
  return '';
}

// Verify examiner owns the session (for sensitive actions)
function verifyExaminerForSession(sessionCode, examinerId) {
  if (!examinerId) return false;
  var sheet = getSheet('סשנים');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(sessionCode).trim()) {
      return normalizeId(data[i][1]) === normalizeId(examinerId);
    }
  }
  return false;
}

