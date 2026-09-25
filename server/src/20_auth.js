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
      // r35 (review 09 F-10 / 01 D23): the same "active" rule login and
      // verifyLogin apply. Without it an examiner disabled in the sheet kept
      // full API access until the 12 h expiry; now at most the 60 s verdict
      // cache below.
      if (!(data[i][3] === 'כן' || data[i][3] === true || data[i][3] === 'TRUE')) break;
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

// ---- Who may READ a session's board (r35, KNOWN_ISSUES #43) ----------------
// examinerDashboard returned every name, phone, ID number and score of ANY
// session to ANY valid examiner token (review 09 F-06, 01 D16) — and ran the
// board's reconciliation writes for it too. The rule is now the one the rest of
// the system already has: the session's own examiner (examinerOwnsSession — the
// creator in 'סשנים' column B; there is no co-examiner concept in the data), or
// the role 'מפקד', which is exactly who listAllSessions shows every live
// session to (examiner.html "📂 הצג כל הסשנים הפעילים" → loadForeignSession).
// A guest SITE of a multi-site session is a row in the quotas, not a second
// examiner, so nothing changes for it. The Worker does not come through here:
// it reads sessionSnapshot with GATEWAY_KEY.
// Only a positive verdict is cached — ownership of a session never changes, and
// a role taken away costs at most SESSION_VIEWER_CACHE_SEC of grace — so a
// board that re-reads itself pays for 'סשנים' once, not on every read.
var SESSION_VIEWER_CACHE_SEC = 300;
function mayViewSession(sessionCode, examinerId) {
  var code = String(sessionCode || '').trim();
  if (!code || !examinerId) return false;
  var key = CACHE_KEY_PREFIX + 'sview_' + normalizeId(examinerId) + '_' + code.slice(0, 40), cache = null;
  try { cache = CacheService.getScriptCache(); if (cache.get(key) === '1') return true; } catch (eGet) { cache = null; }
  var allowed = examinerOwnsSession(code, examinerId) || getExaminerRole(examinerId) === 'מפקד';
  if (allowed) { try { if (!cache) cache = CacheService.getScriptCache(); cache.put(key, '1', SESSION_VIEWER_CACHE_SEC); } catch (ePut) {} }
  return allowed;
}
function requireSessionViewer(p) {
  if (mayViewSession(p.sessionCode, p.examinerId)) return null;
  return jsonResponse({ status: 'error', code: 'not_session_owner', message: 'אין הרשאה — בוחן לא תואם לסשן' });
}

// Verify examiner owns the session (for sensitive actions)
function verifyExaminerForSession(sessionCode, examinerId) {
  // One rule, one read: examinerOwnsSession (44_sessions_misc.js) serves the
  // check from the per-execution memo of 'סשנים', so a handler that verifies
  // ownership and then reads the session row pays for the sheet once.
  return examinerOwnsSession(sessionCode, examinerId);
}

// ========== Signed bank grant (DESIGN §11.2) =================================
// The question TEXTS are not public any more: they live as PRIVATE Workers
// assets and the session-gateway hands a device only what it was issued — the
// 30 ids of this exam, the ids of this practice draw, or (for an examiner) any
// id and a whole language bank. What authorises that is a grant this script
// signs; the client only carries the string, and only the Worker verifies it.
//
// Shape: '<payloadB64url>.<sigB64url>', exactly the construction
// handleGetResultUploadToken uses — base64url without padding, and the HMAC is
// taken over the ENCODED payload so the Worker verifies the bytes it received
// rather than a re-serialisation of them.

// The one place that reads GATEWAY_KEY. requireGatewayKey and the grant signer
// must never disagree about what "configured" means.
function gatewayKey() {
  try { return String(PropertiesService.getScriptProperties().getProperty('GATEWAY_KEY') || ''); }
  catch (e) { return ''; }
}

// An exam grant has to outlive the exam plus every extension and every mid-exam
// refresh (a session code lives 8h); practice is one shorter sitting; an
// examiner holds one for a working day.
var BANK_GRANT_TTL_MS = { exam: 4 * 3600 * 1000, practice: 2 * 3600 * 1000, examiner: 8 * 3600 * 1000 };
var BANK_GRANT_SUB_MAX = 128;   // an identity, not a payload: never let a caller grow the token

// Escape every non-ASCII character. A practice `sub` carries a class code and a
// student id that may be Hebrew, and the Worker decodes the payload byte-wise
// (atob); \uXXXX keeps the JSON pure ASCII so both sides read the same string.
function bankGrantJson(obj) {
  return JSON.stringify(obj).replace(/[\u0080-\uFFFF]/g, function(ch) {
    return '\\u' + ('000' + ch.charCodeAt(0).toString(16)).slice(-4);
  });
}

// `key` is optional and exists only so a caller that has already read the
// property does not pay for a second Properties round trip.
function signBankGrant(payloadObj, key) {
  var secret = key || gatewayKey();
  var payloadB64 = Utilities.base64EncodeWebSafe(bankGrantJson(payloadObj)).replace(/=+$/, '');
  var sigB64 = Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(payloadB64, secret)).replace(/=+$/, '');
  return payloadB64 + '.' + sigB64;
}

// r33 (24/09/2026): this script now also READS its own grants. A phone that
// cannot reach the Worker asks bankRelay (60_exam.js) to fetch its texts server
// to server, and the relay must never become a proxy for anything but a grant
// this script signed — so it is verified here before a single byte is fetched.
// The exact mirror of signBankGrant: the HMAC is recomputed over the ENCODED
// payload and compared as the same unpadded base64url string the signer
// produced, in a compare whose time does not depend on where the strings first
// differ. The payload is decoded only once the signature holds, and a stale
// `exp` or a version other than 1 is refused like the Worker refuses it.
// Returns the payload object, or null for anything at all wrong.
var BANK_GRANT_MAX_CHARS = 4096;
function verifyBankGrant(raw, key) {
  if (typeof raw !== 'string' || !raw || raw.length > BANK_GRANT_MAX_CHARS) return null;
  var parts = raw.split('.');
  if (parts.length !== 2 || !/^[A-Za-z0-9_-]+$/.test(parts[0]) || !/^[A-Za-z0-9_-]+$/.test(parts[1])) return null;
  var secret = key || gatewayKey();
  if (!secret) return null;
  var expected = '';
  try {
    expected = Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(parts[0], secret)).replace(/=+$/, '');
  } catch (eSign) { return null; }
  if (!constantTimeEqual(expected, parts[1])) return null;
  var payload = null;
  try { payload = JSON.parse(base64UrlDecodeAscii(parts[0])); } catch (eParse) { return null; }
  if (!payload || typeof payload !== 'object' || payload.v !== 1) return null;
  if (!(Number(payload.exp) > Date.now())) return null;
  return payload;
}

function constantTimeEqual(a, b) {
  var x = String(a), y = String(b);
  if (x.length !== y.length) return false;
  var diff = 0;
  for (var i = 0; i < x.length; i++) diff |= x.charCodeAt(i) ^ y.charCodeAt(i);
  return diff === 0;
}

// base64url (padding optional) → one character per decoded byte. Enough for a
// grant payload and nothing else: bankGrantJson escapes every character above
// U+007F, so the signed JSON is pure ASCII. Throws outside the alphabet.
var BASE64URL_ALPHABET = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_';
function base64UrlDecodeAscii(text) {
  var s = String(text || '').replace(/=+$/, ''), out = '', buffer = 0, bits = 0;
  for (var i = 0; i < s.length; i++) {
    var v = BASE64URL_ALPHABET.indexOf(s.charAt(i));
    if (v < 0) throw new Error('not base64url');
    buffer = ((buffer << 6) | v) & 0x3FFF;   // bits < 8 before the shift, so 14 live bits at most
    bits += 6;
    if (bits >= 8) { bits -= 8; out += String.fromCharCode((buffer >> bits) & 0xFF); }
  }
  return out;
}

// Is the Worker that serves the texts wired up at all? startExam/startPractice
// ask BEFORE they write anything: an exam whose questions can never load must
// not consume the examinee's attempt.
function bankGrantConfigured() { return Boolean(gatewayUrl() && gatewayKey()); }

// { url, grant, exp } — or null when the Worker is not configured, which every
// caller must translate into bankNotConfiguredResponse() rather than an exam
// that starts and then cannot show a question.
function bankGrantFor(scope, ids, sub) {
  var ttl = BANK_GRANT_TTL_MS[String(scope)];
  var url = gatewayUrl(), key = gatewayKey();   // read once: this runs inside exam start
  if (!ttl || !url || !key) return null;
  var exp = Date.now() + ttl;
  var payload = { v: 1, s: String(scope) };
  // The examiner scope carries no id list — it may read any id and a full
  // language bank, so an `ids` field would only be a lie the Worker ignores.
  if (String(scope) !== 'examiner') {
    var list = [];
    for (var i = 0; ids && i < ids.length; i++) list.push(Number(ids[i]));
    payload.ids = list;
  }
  payload.sub = String(sub || '').slice(0, BANK_GRANT_SUB_MAX);
  payload.exp = exp;
  return { url: url, grant: signBankGrant(payload, key), exp: exp };
}

// One answer for every path that cannot issue a grant: this is an administrator
// problem, and the examinee/student is told so instead of "try again".
function bankNotConfiguredResponse() {
  return jsonResponse({ status: 'error', code: 'bank_not_configured',
    message: 'מאגר השאלות אינו מוגדר בשרת — פנה למנהל המערכת' });
}

// ---- Sites that moved to the new system (r35, KNOWN_ISSUES #44) -------------
// Script Property MOVED_SITES = a JSON array of site names, e.g.
//   ["בסיס 6","בח\"א 6"]
// Missing or empty = nothing moved and nothing changes. A listed site can no
// longer OPEN a session here — createSession answers site_moved (with the
// address in MOVED_SITES_URL, when that property is set) — whether it is the
// session's host site or a guest site of a multi-site session. Sessions that
// are already open are not touched and finish normally: registration, the
// exam and the result never look at this list. No redirect, no fallback.
// A value that is not a JSON array of strings refuses every new session with
// moved_sites_invalid: a typo must be seen the first time anyone opens a
// session, not silently read as "nothing moved". health reports the count.
// One property read per request (MOVED_SITES_URL only on an actual refusal),
// through the same PropertiesService every other setting here uses.
var MOVED_SITES_PROPERTY = 'MOVED_SITES';
var MOVED_SITES_URL_PROPERTY = 'MOVED_SITES_URL';
function normalizeSiteName(name) { return String(name === null || name === undefined ? '' : name).trim().replace(/\s+/g, ' '); }
// { sites: [normalized names], invalid: bool }
function movedSitesSetting() {
  var raw = String(PropertiesService.getScriptProperties().getProperty(MOVED_SITES_PROPERTY) || '').trim();
  if (!raw) return { sites: [], invalid: false };
  var parsed = null;
  try { parsed = JSON.parse(raw); } catch (eParse) { return { sites: [], invalid: true }; }
  if (!Array.isArray(parsed)) return { sites: [], invalid: true };
  var sites = [];
  for (var i = 0; i < parsed.length; i++) {
    if (typeof parsed[i] !== 'string') return { sites: [], invalid: true };
    var name = normalizeSiteName(parsed[i]);
    if (name) sites.push(name);
  }
  return { sites: sites, invalid: false };
}
// null = every name may open a session; otherwise the refusal to answer with.
function movedSiteRefusal(siteNames) {
  var moved = movedSitesSetting();
  if (moved.invalid) {
    return jsonResponse({ status: 'error', code: 'moved_sites_invalid',
      message: 'הגדרת MOVED_SITES בשרת אינה תקינה (צריך מערך JSON של שמות אתרים) — פנה למנהל המערכת' });
  }
  if (!moved.sites.length) return null;
  for (var i = 0; i < siteNames.length; i++) {
    var name = normalizeSiteName(siteNames[i]);
    if (!name || moved.sites.indexOf(name) === -1) continue;
    var url = String(PropertiesService.getScriptProperties().getProperty(MOVED_SITES_URL_PROPERTY) || '').trim();
    var body = { status: 'error', code: 'site_moved', site: name,
      message: 'האתר "' + name + '" עבר למערכת החדשה — יש להשתמש בקישור החדש' + (url ? ': ' + url : '') };
    if (url) body.url = url;
    return jsonResponse(body);
  }
  return null;
}
// For health: how many sites are listed, or 'invalid' — never the names.
function movedSitesHealth() {
  try {
    var moved = movedSitesSetting();
    return moved.invalid ? 'invalid' : moved.sites.length;
  } catch (e) { return 'error'; }
}

// Address of the polling Worker (DESIGN §3.4). Empty = examinees poll this
// script directly; setting/clearing the ScriptProperty switches the whole fleet
// within one getSessionInfo, without a Pages deploy.
function gatewayUrl() {
  try { return String(PropertiesService.getScriptProperties().getProperty('GATEWAY_URL') || '').trim(); }
  catch (e) { return ''; }
}
