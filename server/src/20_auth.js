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
  var key = CACHE_KEY_PREFIX + 'tok_' + normalizeId(examinerId) + '_' + String(token).slice(0, 80), cache = null;
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
  'gateway',           // the session-poll Worker (sessionSnapshot; it also holds GATEWAY_KEY)
  'localhost-dev'      // local development
];
function checkOrigin(p) {
  // Allowed: actions called from external services (none currently) or no origin enforcement on certain reads
  // An empty action is the public 'is the API running' ping and needs no origin.
  if (String(p.action || '') === '') return null;
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

// One row of 'ממתינים' → token verdict. Per-examinee audio (column J) rides
// along on the row we already read, so callers get it without a second scan.
// - legacy: true when the stored row predates token support (empty cell) —
//   we accept the call but flag it so we can audit / tighten later.
// - reason values (when invalid): 'not_found', 'missing', 'mismatch'.
function examineeTokenVerdict(row, examineeToken) {
  var rowAudio = String(row[9] || '').trim() === 'on' ? 'on' : 'off';
  var storedToken = String((row.length > 12 ? row[12] : '') || '').trim();
  if (!storedToken) return { valid: true, legacy: true, audioMode: rowAudio };
  if (!examineeToken) return { valid: false, reason: 'missing' };
  if (String(examineeToken).trim() === storedToken) return { valid: true, legacy: false, audioMode: rowAudio };
  return { valid: false, reason: 'mismatch' };
}

// The token check reads the tail of 'ממתינים'; the handler that runs right after
// it needs the very same row — its status, its sheet row number (the status
// write needs it), the audio flag and the time extension. Handing that read
// forward is what keeps startExam and submitResult at ONE read of the sheet.
// The context is SINGLE USE — the auth check hands it to the handler that runs
// immediately after it, and anything later reads the sheet again. Nothing is
// ever decided from a row this request did not just read.
//   latest = the newest row of this examinee, whatever its status (the token
//            and the "may they submit" rule are decided on it, as before)
//   active = the newest approved/in_exam row (the attempt being started)
var EXAMINEE_ROW_CONTEXT = null;
function examineeRowContext(sessionCode, idNumber, fresh) {
  var cached = EXAMINEE_ROW_CONTEXT;
  EXAMINEE_ROW_CONTEXT = null;
  if (!fresh && cached && String(cached.sessionCode) === String(sessionCode) &&
      normalizeId(cached.idNumber) === normalizeId(idNumber)) return cached;
  diagMark('sheet:pending-examinee');
  var tail = readPendingTail(), rows = tail.rows;
  var ctx = { sessionCode: sessionCode, idNumber: idNumber, tail: tail, latest: null, active: null };
  for (var i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][0]) !== String(sessionCode) || normalizeId(rows[i][1]) !== normalizeId(idNumber)) continue;
    var entry = { row: rows[i], rowNumber: i + tail.off + 1, status: String(rows[i][5] || '').trim() };
    if (!ctx.latest) ctx.latest = entry;
    if (!ctx.active && (entry.status === 'approved' || entry.status === 'in_exam')) ctx.active = entry;
    if (ctx.latest && ctx.active) break;
  }
  EXAMINEE_ROW_CONTEXT = ctx;
  return ctx;
}

// Returns { valid: bool, legacy: bool, reason: string }
function verifyExamineeToken(sessionCode, idNumber, examineeToken) {
  var ctx = examineeRowContext(sessionCode, idNumber, true);   // auth always reads fresh
  if (!ctx.latest) return { valid: false, reason: 'not_found' };
  var verdict = examineeTokenVerdict(ctx.latest.row, examineeToken);
  ctx.audioMode = verdict.audioMode || 'off';
  return verdict;
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
  // One rule, one read: examinerOwnsSession (44_sessions_misc.js) serves the
  // check from the per-execution memo of 'סשנים', so a handler that verifies
  // ownership and then reads the session row pays for the sheet once.
  return examinerOwnsSession(sessionCode, examinerId);
}

