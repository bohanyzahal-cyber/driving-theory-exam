// ========== Exam start, practice draw, result submission ====================
//
// One call starts an exam (startExam) and one call ends it (submitResult). The
// server draws the questions from questionIndex() (70_questions.js), stores the map it drew, and
// scores the submission against the answer key — the client is trusted for
// nothing but WHAT IT DISPLAYED (question and answer texts, for the certificate).
//
// What the old flow did instead: getExamQuestions (Drive/pool read, ~10s) →
// registerExamQuestions (full read of 'ממתינים', no idempotency, four client
// retries) → markExamStarted (another full read) → submitResult (full read of
// 'מבחנים', 8.2MB, plus three full reads of 'תוצאות' and a Drive read for the
// wrong-answer texts). That is the 15MB submit and the 10-second start.

// ---- The question map of one exam ------------------------------------------
// Stored in 'מבחנים' exactly as before — six columns, one JSON map per row —
// so nothing downstream changes; the map entries gained `topic` (the blueprint
// bucket, for the certificate) since the server can no longer look a category up
// in a bank it does not have.
var EXAM_BASE_MINUTES = 40;
var EXAM_MIN_QUESTIONS = 25;               // review E S1: never register or score a short map
var EXAM_MAP_CACHE_SEC = 10800;            // 3h — longer than any exam plus its extensions
var EXAM_MAP_MAX_AGE_MS = 8 * 3600 * 1000; // a session code lives 8h; an older row is a previous attempt
var EXAM_SUSPICIOUS_SEC = 180;             // a "finished" exam faster than this is flagged, not blocked

// The cache key carries the ATTEMPT, not just the person: an examinee who was
// reset and registered again gets a new 'ממתינים' row with a new registration
// time and must be drawn a fresh exam, not handed the cached one.
function examAttemptKey(row) {
  var stamp = (row && row[4] instanceof Date) ? row[4].toISOString() : String((row && row[4]) || '');
  return stamp.replace(/[^0-9A-Za-z]/g, '').slice(-14) || 'na';
}
function examMapCacheKey(sessionCode, idNumber, attemptKey) {
  return CACHE_KEY_PREFIX + 'qmap_' + String(sessionCode || '').trim() + '_' + normalizeId(idNumber) + '_' + attemptKey;
}
function readExamMapCache(sessionCode, idNumber, attemptKey) {
  try {
    var hit = CacheService.getScriptCache().get(examMapCacheKey(sessionCode, idNumber, attemptKey));
    if (!hit) return null;
    var record = JSON.parse(hit);
    return (record && record.map && record.map.length) ? record : null;
  } catch (e) { return null; }
}
function writeExamMapCache(sessionCode, idNumber, attemptKey, record) {
  try {
    CacheService.getScriptCache().put(examMapCacheKey(sessionCode, idNumber, attemptKey), JSON.stringify(record), EXAM_MAP_CACHE_SEC);
  } catch (e) { /* the sheet is the source of truth; the cache only saves a read */ }
}

// 'מבחנים' carries one JSON map per row and grows forever — reading it whole to
// use a single row was 8.2MB of the 15MB submit. Columns A-B (session, id) are
// scanned bottom-up and only the matching row's C-F is read.
// maxAgeMs bounds the search to the current session; 0 accepts any age (a submit
// must find its own registration however long the exam ran).
// Returns null when there is no row, or { map: null } when the row cannot be
// parsed — "a registration exists but is unreadable" is not "no registration".
function readExamRegistration(sessionCode, idNumber, maxAgeMs) {
  var sheet = getSheet('מבחנים'), lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  diagMark('sheet:exam-keys');
  var keys = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  for (var i = keys.length - 1; i >= 0; i--) {
    if (String(keys[i][0]) !== String(sessionCode)) continue;
    if (normalizeId(keys[i][1]) !== normalizeId(idNumber)) continue;
    diagMark('sheet:exam-row');
    var record = parseExamRegistrationRow(sheet.getRange(i + 2, 3, 1, 4).getValues()[0]);
    if (maxAgeMs && examRegistrationAgeMs(record) > maxAgeMs) return null;
    return record;
  }
  return null;
}
function parseExamRegistrationRow(cells) {
  var record = {
    map: null,
    at: (cells[1] instanceof Date) ? cells[1].toISOString() : String(cells[1] || ''),
    lang: String(cells[2] || ''),
    unverified: Number(cells[3] || 0)
  };
  try {
    var map = JSON.parse(cells[0]);
    // An empty array IS a readable map — and a refusable one (review E S1).
    // Only unparseable JSON leaves map null, i.e. "registered but unreadable".
    if (Array.isArray(map)) record.map = map;
  } catch (e) { /* record.map stays null → the result is stored unverified */ }
  return record;
}
function examRegistrationAgeMs(record) {
  var at = record && record.at ? new Date(record.at) : null;
  if (!at || isNaN(at.getTime())) return 0;   // unparseable stamp: do not discard the map over it
  return Date.now() - at.getTime();
}

function appendExamRegistration(sessionCode, idNumber, map, at, lang) {
  diagMark('sheet:append-exam');
  var sheet = getSheet('מבחנים');
  // Every reader here skips row 1 as a header, so a sheet that somehow has none
  // would swallow its first exam.
  if (sheet.getLastRow() === 0 && SHEET_HEADERS['מבחנים']) sheet.appendRow(SHEET_HEADERS['מבחנים']);
  sheet.appendRow([String(sessionCode), normalizeId(idNumber), JSON.stringify(map), at, lang, 0]);
}

// ---- startExam --------------------------------------------------------------
// POST, examinee token. Replaces getExamQuestions + registerExamQuestions +
// markExamStarted with one idempotent call: the same (session, attempt) always
// gets the same 30 questions back, so a retry after a lost response resumes the
// exam instead of drawing a second one.
defineAction('startExam', { methods: ['POST'], auth: 'examinee', handler: handleStartExam,
  rateLimit: { max: 10, windowSec: 60, id: function(p) { return String(p.sessionCode || '') + '_' + normalizeId(p.idNumber); } } });
function handleStartExam(data) {
  var sessionCode = String(data.sessionCode || '').trim();
  if (!sessionCode || !data.idNumber) return jsonResponse({ status: 'error', message: 'חסרים פרטי נבחן' });
  var ctx = examineeRowContext(sessionCode, data.idNumber);
  if (!ctx.active) return jsonResponse({ status: 'error', code: 'not_approved', message: 'נבחן לא מאושר למבחן' });

  var row = ctx.active.row;
  var lang = String(data.language || row[6] || 'he').toLowerCase();
  var license = String(data.license || row[8] || 'B').trim();
  if (!EXAM_STRUCTURE_SERVER[license]) {
    return jsonResponse({ status: 'error', code: 'unknown_license', message: 'דרגה לא מוכרת: ' + license });
  }
  // Before the draw, before the cache write, before the status flip: without a
  // grant the device can never fetch a question text, and a row flipped to
  // in_exam would have spent the examinee's attempt on an exam that cannot run.
  if (!bankGrantConfigured()) return bankNotConfiguredResponse();

  var attempt = examAttemptKey(row);
  var record = readExamMapCache(sessionCode, data.idNumber, attempt);
  // Only an exam already in progress may reuse a stored map. An 'approved' row
  // is a NEW attempt (a reset examinee re-registered), and it must be drawn
  // fresh even though a map of the previous attempt is still in the sheet.
  if (!record && ctx.active.status === 'in_exam') {
    record = readExamRegistration(sessionCode, data.idNumber, EXAM_MAP_MAX_AGE_MS);
    if (record && (!record.map || record.map.length < EXAM_MIN_QUESTIONS)) record = null;
  }
  if (!record) {
    try { record = drawExamRegistration(sessionCode, data.idNumber, license, lang); }
    catch (err) {
      if (!err || err.code !== 'bank_unavailable') throw err;
      return jsonResponse({ status: 'error', code: 'bank_unavailable', detail: err.detail,
        message: 'מאגר השאלות אינו זמין כעת — פנה לבוחן' });
    }
  }
  writeExamMapCache(sessionCode, data.idNumber, attempt, record);

  // approved → in_exam, through the single status writer (it flushes and drops
  // the poller snapshot). An in_exam row keeps its original start time.
  if (ctx.active.status === 'approved') {
    setPendingStatus(getSheet('ממתינים'), ctx.active.rowNumber, sessionCode, 'in_exam', { examStart: nowISO() });
    ctx.active.status = 'in_exam';
  }

  // A retry gets the same ids with a FRESHLY signed grant: the map is
  // idempotent, the clock is not, and an examinee who reloads at minute 38 must
  // not be handed a grant that expires before the extension does.
  var bank = bankGrantFor('exam', examMapIds(record.map), sessionCode + ':' + normalizeId(data.idNumber));

  return jsonResponse({
    status: 'ok', build: THEORY_API_BUILD,
    examMinutes: examMinutesFor(row), extraMinutes: sumExtraMinutes(sessionCode, data.idNumber),
    audioMode: String(row[9] || '').trim() === 'on' ? 'on' : 'off',
    language: record.lang || lang, license: license, registeredAt: record.at,
    bank: bank,
    questions: examQuestionsForClient(record.map)
  });
}

// The ids of a stored map, in map order — that is the order the client shows
// them in, and the order the grant authorises.
function examMapIds(map) {
  var ids = [];
  for (var i = 0; i < map.length; i++) ids.push(map[i].qId);
  return ids;
}

function drawExamRegistration(sessionCode, idNumber, license, lang) {
  var drawn = drawExamIds(license, lang);
  if (drawn.length < EXAM_MIN_QUESTIONS) {
    throw questionBankUnavailable('נדרשות לפחות ' + EXAM_MIN_QUESTIONS + ' שאלות', String(drawn.length));
  }
  var map = [];
  for (var i = 0; i < drawn.length; i++) {
    var order = drawShuffleOrder();
    // Non-null by construction: drawExamIds skips every id the key cannot answer.
    var correct = answerKeyIndex(drawn[i].id, lang);
    map.push({ qIdx: i, qId: drawn[i].id, shuffleOrder: order, correctShuffledIdx: order.indexOf(correct), topic: drawn[i].topic });
  }
  var at = nowISO();
  appendExamRegistration(sessionCode, idNumber, map, at, lang);
  return { map: map, at: at, lang: lang, unverified: 0 };
}

// The client fetches the texts from the Worker with the grant above and needs
// only what the server decided: which questions, in which answer order, under
// which topic.
function examQuestionsForClient(map) {
  var out = [];
  for (var i = 0; i < map.length; i++) {
    out.push({ id: map[i].qId, order: map[i].shuffleOrder, topic: map[i].topic || '' });
  }
  return out;
}

// Column K of 'ממתינים' holds the examiner's time extension — same whitelist the
// approval poll applies, so both sides compute the same deadline.
function examMinutesFor(row) {
  var ext = parseFloat(row[10]) || 1;
  if (ext !== 1.25 && ext !== 1.5) ext = 1;
  return Math.round(EXAM_BASE_MINUTES * ext);
}

// ---- bankRelay (r33, 24/09/2026, KNOWN_ISSUES #38) --------------------------
// On 23/09 many phones registered through Google and then never managed ONE
// request to the Worker — something in front of workers.dev refused them
// before our code ran. Such a phone now asks THIS script for its question
// texts, and the script fetches them from the Worker server to server with the
// phone's own grant: the answer is the Worker's /v1/bank body, unchanged, plus
// relay:true, so the page ingests it with the very code that ingests the
// Worker's own answer (shared/bank.js ingestAnswer).
// Not an open proxy: the URL is fixed, and the grant must be one this script
// signed, in the exam scope, for THIS examinee (the sub startExam writes), and
// unexpired — the examinee token alone fetches nothing. The grant is a bearer
// credential for 30 texts: it is never logged, and every error detail is
// scrubbed of it. Every UrlFetchApp call of the server lives in this module
// (exam-only): the reports project never needs the external-request scope.
var RELAY_FAILED_MESSAGE = 'לא הצלחנו לטעון את השאלות דרך השרת — נסה שוב או פנה לבוחן';
var RELAY_DETAIL_MAX = 80;
defineAction('bankRelay', { methods: ['POST'], auth: 'examinee', handler: handleBankRelay,
  rateLimit: { max: 10, windowSec: 300, id: function(p) { return String(p.sessionCode || '') + '_' + normalizeId(p.idNumber); } } });
function handleBankRelay(data) {
  if (!gatewayUrl()) return bankNotConfiguredResponse();
  if (!examGrantFor(data.grant, data.sessionCode, data.idNumber)) {
    return jsonResponse({ status: 'error', code: 'grant_invalid',
      message: 'הרשאת השאלות אינה תקפה — נסה להתחיל שוב' });
  }
  var fetched = fetchBankFromGateway(data.grant);
  if (fetched.body) {
    fetched.body.relay = true;
    return jsonResponse(fetched.body);
  }
  return jsonResponse({ status: 'error', code: 'relay_failed', http: fetched.http, retryable: true,
    message: RELAY_FAILED_MESSAGE, detail: fetched.detail });
}

// The payload of `grant` when it is a live exam grant of (sessionCode, idNumber)
// — the very sub startExam signs — or null.
function examGrantFor(grant, sessionCode, idNumber) {
  var payload = verifyBankGrant(grant);
  if (!payload || payload.s !== 'exam') return null;
  var sub = (String(sessionCode || '').trim() + ':' + normalizeId(idNumber)).slice(0, BANK_GRANT_SUB_MAX);
  if (typeof payload.sub !== 'string' || payload.sub !== sub) return null;
  if (!Array.isArray(payload.ids) || !payload.ids.length) return null;
  return payload;
}

// One GET of the Worker's /v1/bank, as bankRelay and testGatewayReachability
// both make it. { http, body } for a real bank answer (HTTP 200, status ok,
// questions an array); otherwise { http, detail } — http 0 when the fetch
// itself threw. getContentText('UTF-8') by name: the texts are Hebrew,
// Russian, Arabic and Amharic, and must not depend on a Content-Type default.
function fetchBankFromGateway(grant) {
  var url = gatewayUrl().replace(/\/+$/, '') + '/v1/bank?grant=' + encodeURIComponent(String(grant));
  var http = 0, text = '';
  diagMark('relay:bank');
  try {
    var res = UrlFetchApp.fetch(url, { method: 'get', muteHttpExceptions: true, followRedirects: true,
      headers: { 'Accept': 'application/json' } });
    http = Number(res.getResponseCode()) || 0;
    text = String(res.getContentText('UTF-8') || '');
  } catch (err) {
    diagMark('relay:bank-error');
    return { http: 0, detail: relayDetail(String(err && err.message ? err.message : err), grant) };
  }
  diagMark('relay:bank-done');
  if (http === 200) {
    var body = null;
    try { body = JSON.parse(text); } catch (eParse) { body = null; }
    if (body && typeof body === 'object' && body.status === 'ok' && Array.isArray(body.questions)) {
      return { http: http, body: body };
    }
  }
  return { http: http, detail: relayDetail(text, grant) };
}

// What the page (and 'אבחון', through the page's own report) may see of a
// failure: Cloudflare's "error code: NNNN" when the body carries one, otherwise
// the first RELAY_DETAIL_MAX printable characters — never the grant, which an
// exception message may echo inside the URL.
function relayDetail(text, grant) {
  var s = String(text || '');
  if (grant) s = s.split(String(grant)).join('[grant]');
  s = s.replace(/grant=[^&\s"'<>]*/g, 'grant=[grant]');
  var cloudflare = /error code:\s*\d{3,5}/i.exec(s);
  if (cloudflare) return cloudflare[0];
  return s.replace(/[\u0000-\u001F\u007F-\u009F]+/g, ' ').replace(/\s+/g, ' ').trim().slice(0, RELAY_DETAIL_MAX);
}

// ---- testGatewayReachability — run BY HAND from the Apps Script editor ------
// Proves (or disproves) the one thing bankRelay depends on: that GOOGLE can
// reach the Worker. Three GETs, each logged with its HTTP code and the first
// 120 characters of its body:
//   /                      the Worker's front door (never touches Google)
//   /v1/bank?grant=x.y     a fake grant — our Worker answers a JSON 403
//   /v1/bank, real grant   ONE question, exactly the call bankRelay makes
//                          (logged as a count; no text, no grant)
// A JSON answer = our Worker ran. An HTML 403 with "error code: 1010", or a
// Cloudflare challenge page = Cloudflare stops Google before the Worker, and
// bankRelay cannot help those phones. The first run in the editor is also
// where Google asks for any missing permission (external requests).
// Never prints GATEWAY_KEY, a real grant or a question text.
function testGatewayReachability() {
  var base = gatewayUrl().replace(/\/+$/, '');
  if (!base) {
    var unset = 'testGatewayReachability: GATEWAY_URL is not set in the Script properties — nothing to test';
    Logger.log(unset);
    return unset;
  }
  var probes = [{ name: 'front-door', path: '/' }, { name: 'fake-grant', path: '/v1/bank?grant=x.y' }];
  var parts = [], kinds = [];
  for (var i = 0; i < probes.length; i++) {
    var r = gatewayProbe(base + probes[i].path);
    Logger.log(probes[i].name + ' GET ' + probes[i].path + ' -> HTTP ' + r.http + ' [' + r.kind + '] ' + r.snippet);
    parts.push(probes[i].name + '=' + r.http + ' ' + r.kind);
    kinds.push(r.kind);
  }
  var real = { http: 0, ok: false, kind: 'skipped', snippet: 'no grant could be signed (GATEWAY_KEY missing?)' };
  var ids = Object.keys(questionIndex());
  var bank = ids.length ? bankGrantFor('exam', [Number(ids[0])], 'probe:reachability') : null;
  if (bank) {
    var fetched = fetchBankFromGateway(bank.grant);
    real.http = fetched.http;
    real.ok = !!fetched.body;
    real.kind = fetched.body ? 'worker' : gatewayProbeKind(fetched.http, fetched.detail);
    real.snippet = fetched.body
      ? ('status ok, ' + fetched.body.questions.length + ' question(s), missing ' +
         (Array.isArray(fetched.body.missing) ? fetched.body.missing.length : 0))
      : fetched.detail;
  }
  Logger.log('real-grant GET /v1/bank?grant=<1 question> -> HTTP ' + real.http + ' [' + real.kind + '] ' + real.snippet);
  parts.push('real-grant=' + real.http + ' ' + (real.ok ? 'ok' : real.kind));
  kinds.push(real.kind);

  var verdict;
  if (kinds.indexOf('cloudflare-block') !== -1) {
    verdict = 'CLOUDFLARE BLOCKS GOOGLE - bankRelay cannot reach the Worker';
  } else if (kinds[0] === 'worker' && kinds[1] === 'worker' && real.ok) {
    verdict = 'OK - Google reaches the Worker; bankRelay will work';
  } else if (kinds[0] === 'worker' && kinds[1] === 'worker' && !bank) {
    verdict = 'the Worker answers Google, but no grant can be signed here - set GATEWAY_KEY';
  } else if (kinds[0] === 'worker' && kinds[1] === 'worker') {
    verdict = 'the Worker answers Google but refused a real grant - compare GATEWAY_KEY here and in the Worker';
  } else {
    verdict = 'UNCLEAR - read the three lines above';
  }
  var summary = 'testGatewayReachability: ' + parts.join(' | ') + ' => ' + verdict;
  Logger.log(summary);
  return summary;
}
function gatewayProbe(url) {
  var http = 0, text = '';
  try {
    var res = UrlFetchApp.fetch(url, { method: 'get', muteHttpExceptions: true, followRedirects: true,
      headers: { 'Accept': 'application/json' } });
    http = Number(res.getResponseCode()) || 0;
    text = String(res.getContentText('UTF-8') || '');
  } catch (err) {
    text = 'EXCEPTION ' + String(err && err.message ? err.message : err);
  }
  return { http: http, kind: gatewayProbeKind(http, text),
    snippet: text.replace(/[\u0000-\u001F\u007F-\u009F]+/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 120) };
}
// 'worker' = our Worker answered (every answer of it is a JSON envelope with a
// status); 'cloudflare-block' = Cloudflare's own error or challenge page.
function gatewayProbeKind(http, text) {
  var body = String(text || ''), json = null;
  try { json = JSON.parse(body); } catch (e) { json = null; }
  if (json && typeof json === 'object' && typeof json.status === 'string') return 'worker';
  if (/error code:\s*\d{3,5}/i.test(body) ||
      /cf-chl|challenge-platform|cf_chl_opt|Just a moment|Attention Required|cf-error-details|cf-browser-verification/i.test(body)) {
    return 'cloudflare-block';
  }
  return http ? 'unexpected' : 'unreachable';
}

// ---- Retired actions --------------------------------------------------------
// One release of grace for a client that was loaded before the deploy: it asks
// for questions, gets a clear "refresh the page" instead of a broken exam.
defineAction('getExamQuestions', { methods: ['GET'], auth: 'none', handler: handleClientOutdated });
defineAction('registerExamQuestions', { methods: ['POST'], auth: 'none', handler: handleClientOutdated });
function handleClientOutdated() {
  return jsonResponse({ status: 'error', code: 'client_outdated',
    message: 'גרסה חדשה של המערכת — יש לרענן את הדף (F5)' });
}

// startExam already flipped the row to in_exam; the old client's separate ping
// has nothing left to do. Kept (as a no-op) only so that client does not treat
// an unknown action as a failure. Remove in the next release.
defineAction('markExamStarted', { methods: ['GET'], auth: 'examinee', handler: handleMarkExamStartedNoop });
function handleMarkExamStartedNoop() {
  return jsonResponse({ status: 'ok', already: true });
}

// ---- submitResult -----------------------------------------------------------
// Columns of 'תוצאות' by name — the handler used to index 30 positions by hand.
var RESULT_COL = { date: 0, id: 1, name: 2, phone: 3, license: 4, score: 5, percent: 6, verdict: 7,
  time: 8, examiner: 9, site: 10, classroom: 11, language: 12, session: 13, attempt: 14, wrongDetails: 15,
  sent: 16, dq: 17, waLink: 18, population: 19, corrected: 20, audio: 21, verified: 22, suspicious: 23,
  dqEventId: 24, correctedBy: 25, correctionReason: 26, correctionDate: 27, langPath: 28, device: 29 };
var RESULT_PASS_RATIO = 0.86;             // 26/30
var FABRICATED_FAIL_MARKERS = ['סגירת דפדפן', 'טיימאאוט', 'סיום ידני'];
var UNVERIFIED_PREFIX = '⚠️ ציון לא אומת';

defineAction('submitResult', { methods: ['POST'], auth: 'examinee', handler: handleSubmitResult });
function handleSubmitResult(data) {
  var gate = submitGate(data);
  if (gate.error) return gate.error;
  recordSubmitClientLog(data);

  var registration = readExamRegistration(data.sessionCode, data.idNumber, 0);
  var guard = submitRegistrationGuard(registration, data);
  if (guard) return guard;

  var scored = registration && registration.map && registration.map.length
    ? scoreRegisteredExam(registration.map, data.answers, registration.lang)
    : null;
  applyScore(data, scored, registration, !!(data.answers && data.answers.length));
  var wrongAnswers = scored ? wrongAnswerItems(scored, data.license) : [];

  // 'תוצאות' is read ONCE, as late as possible: supersede, duplicate, פסול and
  // idempotency all decide from the same snapshot. Three full reads of a sheet
  // that grows forever were most of what was left of the submit's cost.
  diagMark('sheet:results-submit');
  var sheet = getSheet('תוצאות');
  var tail = readTail(sheet, RESULT_COL.date);
  supersedeFabricatedFails(sheet, tail, data);

  var duplicate = findDuplicateResult(tail, data, gate);
  if (duplicate) {
    markPendingCompleted(data.sessionCode, data.idNumber, gate.pending);
    return jsonResponse({ status: 'ok', waLink: duplicate[RESULT_COL.waLink] || '', duplicate: true });
  }
  // Counted BEFORE the פסול rows are voided below: a disqualified exam was
  // still an attempt, and voiding it is only about not leaving two live rows.
  var attemptNum = countAttempts(data.idNumber, data.license, attemptRows(tail)) + 1;
  supersedeDisqualifications(sheet, tail, data);

  var waLink = buildResultWaLink(data, wrongAnswers, attemptNum);
  if (findIdenticalResult(tail, data)) {
    markPendingCompleted(data.sessionCode, data.idNumber, gate.pending);
    return jsonResponse({ status: 'ok', duplicate: true, waLink: waLink });
  }

  sheet.appendRow(buildResultRow(data, wrongAnswers, attemptNum, waLink));
  markPendingCompleted(data.sessionCode, data.idNumber, gate.pending);
  diagMark('compute:submit-done');
  return jsonResponse({ status: 'ok', waLink: waLink });
}

// Rate limit, examinee token and "is this person actually in this exam".
// Returns { error } or { pending, status } — the pending snapshot is handed on
// so the completion write does not read 'ממתינים' a second time.
function submitGate(data) {
  var rlErr = requireRateLimit('submitResult', String(data.sessionCode || '') + '_' + normalizeId(data.idNumber), 5, 60);
  if (rlErr) return { error: rlErr };
  var ctx = examineeRowContext(data.sessionCode, data.idNumber);
  var status = ctx.latest ? ctx.latest.status : '';
  // 'cancelled' is accepted: if an examiner reset an examinee who was in fact
  // still mid-exam, a genuine finished submit must be RECORDED, not lost. The
  // fabricated-fail supersede and the duplicate check keep the sheet clean.
  if (data.sessionCode && data.idNumber && ['in_exam', 'approved', 'completed', 'cancelled'].indexOf(status) === -1) {
    return { error: jsonResponse({ status: 'error', message: 'נבחן לא מאושר — לא ניתן לשלוח תוצאות' }) };
  }
  return { pending: pendingSnapshotFromTail(ctx), status: status, ctx: ctx };
}

// markPendingCompleted writes by absolute row index, so a tail read's rows are
// padded back to their sheet positions instead of paying for a second full read.
function pendingSnapshotFromTail(ctx) {
  var sheet = getSheet('ממתינים'), tail = ctx.tail;
  if (!tail || !tail.rows.length) return { sheet: sheet, rows: null };
  if (!tail.off) return { sheet: sheet, rows: tail.rows };
  var padded = [tail.rows[0]];
  for (var i = 0; i < tail.off; i++) padded.push([]);
  return { sheet: sheet, rows: padded.concat(tail.rows.slice(1)) };
}

// The client's last events (≤2KB) ride along with the submit; S2's recorder
// parks them next to the server-side diagnostics. Never fatal to a result.
function recordSubmitClientLog(data) {
  if (!data.clientLog) return;
  try {
    if (typeof diagRecordClientLog === 'function') diagRecordClientLog(data.sessionCode, data.idNumber, data.clientLog);
  } catch (e) { /* a diagnostic must never cost a result */ }
}

// A registered exam MUST come with answers (otherwise a forged score would skip
// the re-score entirely), and a map that is empty or shorter than a real exam is
// a data fault — refuse it loudly instead of scoring 3 questions out of 30 and
// possibly declaring a pass (review E S1).
function submitRegistrationGuard(registration, data) {
  if (!registration) return null;
  if (!data.answers || !Array.isArray(data.answers) || data.answers.length === 0) {
    return jsonResponse({ status: 'error', message: 'הגשה לא תקינה — חסרות תשובות למבחן רשום' });
  }
  if (registration.map && registration.map.length < EXAM_MIN_QUESTIONS) {
    return jsonResponse({ status: 'error', code: 'invalid_registration',
      message: 'רישום המבחן פגום — לא ניתן לנקד. פנה לבוחן.' });
  }
  return null;
}

// ---- scoring ---------------------------------------------------------------
// review E S2: only a real selection counts. null / '' / undefined / a
// non-number / a negative or fractional index all mean "not answered".
function submitAnswerIndex(answer) {
  if (!answer) return -1;
  var raw = answer.selected;
  if (raw === null || raw === undefined || raw === '') return -1;
  var n = Number(raw);
  if (!isFinite(n) || n < 0 || Math.floor(n) !== n) return -1;
  return n;
}

function answerLanguage(answer, registeredLang) {
  var lang = (answer && answer.langAtAnswer) ? String(answer.langAtAnswer) : String(registeredLang || 'he');
  return lang.toLowerCase() || 'he';
}

// review E S4: the correctShuffledIdx stored at registration is NOT read back.
// It is recomputed from the answer key every time, which is also what scores a
// mid-exam language switch against the key of the language the question was
// ANSWERED in (translators reorder answers — en/fr/es/ar have their own order).
// null = this entry cannot be verified at all, and then it can never be correct.
function correctIndexForEntry(entry, lang) {
  if (!entry || !entry.qId || !Array.isArray(entry.shuffleOrder)) return null;
  var orig = answerKeyIndex(entry.qId, lang);
  if (orig === null) return null;
  var pos = entry.shuffleOrder.indexOf(Number(orig));
  return pos >= 0 ? pos : null;
}

function scoreRegisteredExam(map, answers, registeredLang) {
  var scored = { correct: 0, total: map.length, verifiable: 0, items: [] };
  for (var i = 0; i < map.length; i++) {
    var entry = map[i] || null, answer = (answers && answers[i]) || null;
    var lang = answerLanguage(answer, registeredLang);
    var key = correctIndexForEntry(entry, lang);
    var selected = submitAnswerIndex(answer);
    var item = { qIdx: i, qId: entry ? entry.qId : null, topic: (entry && entry.topic) || '',
      lang: lang, key: key, selected: selected, answer: answer };
    item.right = (key !== null && selected === key);
    if (key !== null) scored.verifiable++;
    if (item.right) scored.correct++;
    scored.items.push(item);
  }
  scored.verified = scored.total > 0 && scored.verifiable === scored.total;
  return scored;
}

// The server's tally replaces whatever the client claimed. Without a readable
// registration nothing can be verified, so the row is stored with the unverified
// marker and the examiner reviews it — never silently trusted.
function applyScore(data, scored, registration, hasAnswers) {
  if (!scored) {
    data.verified = false;
    // Without answers there was nothing to verify in the first place (an
    // examiner-entered or legacy row) — only a real submission is flagged.
    if (hasAnswers) {
      data.scoreUnverified = true;
      data.unverifiedReason = registration ? 'רישום מבחן פגום' : 'רישום מבחן חסר';
    }
    // The new client sends no score at all (it never holds the key); an old one
    // sent its own tally. Either way an unverifiable result is stored as what it
    // is — a number the examiner must review — and never as a pass by default.
    data.total = Number(data.total) || (Array.isArray(data.answers) ? data.answers.length : 0) || 30;
    data.score = Math.max(0, Math.min(data.total, Number(data.score) || 0));
    data.percent = Number(data.percent) || Math.round((data.score / data.total) * 100);
    data.passed = data.passed === true || data.passed === 'true';
    return;
  }
  data.score = scored.correct;
  data.total = scored.total;
  data.percent = Math.round((scored.correct / scored.total) * 100);
  data.passed = scored.correct >= Math.ceil(scored.total * RESULT_PASS_RATIO);
  data.verified = scored.verified;
  if (!scored.verified) {
    data.scoreUnverified = true;
    data.unverifiedReason = 'שאלות ללא מפתח תשובות';
  }
  data.suspicious = registration && examRegistrationAgeMs(registration) > 0 &&
    examRegistrationAgeMs(registration) < EXAM_SUSPICIOUS_SEC * 1000;
}

// ---- feedback (certificate + WhatsApp) -------------------------------------
// Built from what the client DISPLAYED: q = the question text, a = the four
// answers in displayed order. The server owns which of them is correct. An old
// client that sends no texts still gets a scored result.
var ANSWER_LABELS_HE = ['א', 'ב', 'ג', 'ד', 'ה', 'ו'];
var ANSWER_LABELS_LATIN = ['A', 'B', 'C', 'D', 'E', 'F'];
var TEXT_UNAVAILABLE = '(טקסט לא זמין)';
function answerLabel(lang, idx) {
  var labels = (String(lang) === 'he') ? ANSWER_LABELS_HE : ANSWER_LABELS_LATIN;
  return labels[idx] || '';
}
function displayedAnswer(shown, idx, lang) {
  if (idx === null) return '(לא זמין כעת)';                 // no key for this language
  var count = (shown && shown.length) ? shown.length : QUESTION_ANSWER_COUNT;
  if (idx < 0 || idx >= count) return (String(lang) === 'he') ? 'לא נענתה' : 'Not answered';
  var text = shown ? String(shown[idx] || '') : '';
  return answerLabel(lang, idx) + ' - ' + (text || TEXT_UNAVAILABLE);
}

function wrongAnswerItems(scored, license) {
  var out = [];
  for (var i = 0; i < scored.items.length; i++) {
    var item = scored.items[i];
    if (item.right) continue;
    var shown = (item.answer && item.answer.a && item.answer.a.length) ? item.answer.a : null;
    out.push({
      questionId: item.qId || '',
      question: (item.answer && item.answer.q) ? String(item.answer.q) : TEXT_UNAVAILABLE,
      yourAnswer: displayedAnswer(shown, item.selected, item.lang),
      correctAnswer: displayedAnswer(shown, item.key, item.lang),
      category: item.topic || questionTopic(item.qId, license)
    });
  }
  return out;
}

// 'פירוט שגויות' (column P) — the commander dashboard aggregates by the מזהה
// שאלה line, and readers tolerate any line being missing.
function formatWrongDetails(items, data) {
  var text = '';
  for (var i = 0; i < items.length; i++) {
    var w = items[i];
    if (w.questionId) text += 'מזהה שאלה: ' + w.questionId + '\n';
    text += 'שאלה: ' + w.question + '\n';
    text += 'תשובת הנבחן: ' + w.yourAnswer + '\n';
    text += 'תשובה נכונה: ' + w.correctAnswer + '\n';
    if (w.category) text += 'קטגוריה: ' + w.category + '\n';
    text += '\n';
  }
  if (data.scoreUnverified) {
    text = UNVERIFIED_PREFIX + ' בשרת (' + (data.unverifiedReason || '') + ') — נדרש אימות ידני\n\n' + text;
  }
  return text;
}

function buildResultWaLink(data, items, attemptNum) {
  var message = '*🚗 אישור תוצאת מבחן תאוריה חיצוני*\n\n' +
    'שם: ' + data.fullName + '\n' +
    'ת.ז.: ' + data.idNumber + '\n' +
    'דרגה: ' + data.license + '\n' +
    (data.population ? 'אוכלוסיה: ' + data.population + '\n' : '') +
    'תאריך: ' + todayStr() + '\n' +
    'תוצאה: *' + resultVerdict(data) + '* (' + data.score + '/' + data.total + ')\n' +
    'זמן: ' + data.time + '\n';
  if (items.length > 0) {
    var wrongForWA = '';
    for (var i = 0; i < items.length; i++) {
      wrongForWA += '❌ ' + items[i].question + '\n' + 'ענית: ' + items[i].yourAnswer + '\n' + '✅ נכון: ' + items[i].correctAnswer + '\n\n';
    }
    message += '\n*שאלות שגויות (' + items.length + '):*\n\n' + wrongForWA;
  } else if (Number(data.total) - Number(data.score) === 0) {
    message += '\nכל התשובות נכונות! 🎉';
  }
  if (attemptNum > 1) message += 'ניסיון: ' + attemptNum + '\n';
  return 'https://wa.me/' + formatPhoneForWA(data.phone) + '?text=' + encodeURIComponent(message);
}

function resultVerdict(data) { return data.passed ? 'עבר' : 'נכשל'; }

// "he → ru → he" tells the examiner at a glance that the examinee switched.
function languagePath(data) {
  if (Array.isArray(data.languageHistory) && data.languageHistory.length > 0) {
    return data.languageHistory.length === 1 ? String(data.languageHistory[0]) : data.languageHistory.join(' → ');
  }
  return data.language || 'he';
}

function buildResultRow(data, wrongAnswers, attemptNum, waLink) {
  var row = [];
  row[RESULT_COL.date] = todayStr();
  row[RESULT_COL.id] = data.idNumber;
  row[RESULT_COL.name] = data.fullName;
  row[RESULT_COL.phone] = data.phone;
  row[RESULT_COL.license] = data.license;
  row[RESULT_COL.score] = data.score + '/' + data.total;
  row[RESULT_COL.percent] = data.percent + '%';
  row[RESULT_COL.verdict] = resultVerdict(data);
  row[RESULT_COL.time] = data.time;
  row[RESULT_COL.examiner] = data.examinerName || '';
  row[RESULT_COL.site] = data.site || '';
  row[RESULT_COL.classroom] = data.classroom || '';
  row[RESULT_COL.language] = data.language || 'he';
  row[RESULT_COL.session] = data.sessionCode || '';
  row[RESULT_COL.attempt] = attemptNum;
  row[RESULT_COL.wrongDetails] = formatWrongDetails(wrongAnswers, data);
  row[RESULT_COL.sent] = false;
  row[RESULT_COL.dq] = false;
  row[RESULT_COL.waLink] = waLink;
  row[RESULT_COL.population] = data.population || '';
  row[RESULT_COL.corrected] = false;
  row[RESULT_COL.audio] = data.audioMode || 'off';
  row[RESULT_COL.verified] = data.verified ? 'מאומת' : '';
  row[RESULT_COL.suspicious] = data.suspicious ? 'חשוד' : '';
  row[RESULT_COL.dqEventId] = '';
  row[RESULT_COL.correctedBy] = '';
  row[RESULT_COL.correctionReason] = '';
  row[RESULT_COL.correctionDate] = '';
  row[RESULT_COL.langPath] = languagePath(data);
  row[RESULT_COL.device] = String(data.device || '');
  return row;
}

// The attempt number counts a lifetime of results, so it may not be decided
// from a tail: an earlier attempt this month can sit above the last 1,000 rows.
// When the snapshot IS the whole sheet it is reused as it is; otherwise
// countAttempts reads the three columns it needs from live + archive itself.
function attemptRows(tail) { return tail.off ? null : tail.rows; }

// ---- the four passes over the one 'תוצאות' snapshot ------------------------
function resultRowMatchesExaminee(row, data) {
  return String(row[RESULT_COL.session]) === String(data.sessionCode) &&
    normalizeId(row[RESULT_COL.id]) === normalizeId(data.idNumber);
}

// A real finished submit must WIN over a system-written fail (the close beacon,
// the dashboard timeout row, an examiner's manual disconnect). Matched on
// session+id ONLY: the fabricated row carries the REGISTRATION language while a
// real submit may carry a different final one after a mid-exam switch.
function supersedeFabricatedFails(sheet, tail, data) {
  var rows = tail.rows, changed = false;
  for (var i = 1; i < rows.length; i++) {
    if (!resultRowMatchesExaminee(rows[i], data)) continue;
    if (String(rows[i][RESULT_COL.verdict]).trim() !== 'נכשל') continue;
    if (!isFabricatedFailNote(String(rows[i][RESULT_COL.wrongDetails] || ''))) continue;
    var sheetRow = i + tail.off + 1;
    sheet.getRange(sheetRow, RESULT_COL.verdict + 1).setValue('בוטל');
    sheet.getRange(sheetRow, RESULT_COL.correctionReason + 1).setValue('בוטל אוטומטית — הנבחן השלים והגיש מבחן');
    sheet.getRange(sheetRow, RESULT_COL.correctionDate + 1).setValue(todayStr());
    rows[i][RESULT_COL.verdict] = 'בוטל';
    changed = true;
  }
  if (changed) SpreadsheetApp.flush();
}
function isFabricatedFailNote(note) {
  for (var i = 0; i < FABRICATED_FAIL_MARKERS.length; i++) {
    if (note.indexOf(FABRICATED_FAIL_MARKERS[i]) !== -1) return true;
  }
  return false;
}

// A genuine prior result for this exact exam. Skipped while the examinee still
// has an in_exam row (a retake after a disqualification is not a duplicate).
function findDuplicateResult(tail, data, gate) {
  if (hasPendingInExam(gate)) return null;
  var rows = tail.rows;
  for (var i = 1; i < rows.length; i++) {
    var verdict = String(rows[i][RESULT_COL.verdict] || '').trim();
    if (verdict === 'פסול' || verdict === 'בוטל') continue;
    if (!resultRowMatchesExaminee(rows[i], data)) continue;
    if (String(rows[i][RESULT_COL.license]) !== String(data.license)) continue;
    if (String(rows[i][RESULT_COL.language]) !== String(data.language || 'he')) continue;
    return rows[i];
  }
  return null;
}

// Re-read this examinee's pending rows: an examiner may have decided something
// while the submit was in flight. Never overwrite a newer examiner decision.
function hasPendingInExam(gate) {
  var snapshot = gate.pending;
  snapshot.rows = refreshExamineePendingRows(snapshot.sheet, snapshot.rows, gate.ctx.sessionCode, gate.ctx.idNumber);
  var rows = snapshot.rows;
  for (var i = rows.length - 1; i >= 1; i--) {
    if (!rows[i] || String(rows[i][0]) !== String(gate.ctx.sessionCode)) continue;
    if (normalizeId(rows[i][1]) !== normalizeId(gate.ctx.idNumber)) continue;
    if (String(rows[i][5]).trim() === 'in_exam') return true;
  }
  return false;
}

// An examinee who was auto-disqualified, let back in and then finished must not
// keep both a פסול row and a real one (base 14, 2026). The audit trail stays.
function supersedeDisqualifications(sheet, tail, data) {
  var rows = tail.rows;
  for (var i = rows.length - 1; i >= 1; i--) {
    if (!resultRowMatchesExaminee(rows[i], data)) continue;
    if (String(rows[i][RESULT_COL.verdict]).trim() !== 'פסול') continue;
    var sheetRow = i + tail.off + 1;
    sheet.getRange(sheetRow, RESULT_COL.verdict + 1).setValue('בוטל');
    sheet.getRange(sheetRow, RESULT_COL.dq + 1).setValue(false);
    sheet.getRange(sheetRow, RESULT_COL.correctionReason + 1).setValue('בוטל אוטומטית — נבחן ניגש למבחן מחדש');
    sheet.getRange(sheetRow, RESULT_COL.correctionDate + 1).setValue(todayStr());
    rows[i][RESULT_COL.verdict] = 'בוטל';
    rows[i][RESULT_COL.dq] = false;
  }
}

// Idempotency: the same result sent twice (a retry whose response was lost)
// must not append a second row. A retake differs in score or time, so it does.
function findIdenticalResult(tail, data) {
  var rows = tail.rows, verdict = resultVerdict(data);
  for (var i = rows.length - 1; i >= 1; i--) {
    if (!resultRowMatchesExaminee(rows[i], data)) continue;
    if (String(rows[i][RESULT_COL.license]) !== String(data.license)) continue;
    if (String(rows[i][RESULT_COL.score]) !== (data.score + '/' + data.total)) continue;
    if (String(rows[i][RESULT_COL.verdict]).trim() !== verdict) continue;
    if (String(rows[i][RESULT_COL.time]) !== String(data.time)) continue;
    return rows[i];
  }
  return null;
}

// ---- submitFailOnClose / cancelFailOnClose ---------------------------------
// The browser-close beacon. It writes a 0/30 fail, which a genuine submit later
// supersedes; it must never write one for an examinee the examiner reset.
defineAction('submitFailOnClose', { methods: ['POST'], auth: 'examinee', handler: handleSubmitFailOnClose });
function handleSubmitFailOnClose(data) {
  var ctx = examineeRowContext(data.sessionCode, data.idNumber);
  var status = ctx.latest ? ctx.latest.status : '';
  if (status === 'cancelled' || status === 'rejected') return jsonResponse({ status: 'ok', skipped: 'cancelled' });

  var sheet = getSheet('תוצאות');
  var tail = readTail(sheet, RESULT_COL.date);
  // Any non-בוטל result for this session+id means the examinee already has a
  // real outcome. Matched on session+id only — see supersedeFabricatedFails.
  for (var i = 1; i < tail.rows.length; i++) {
    if (!resultRowMatchesExaminee(tail.rows[i], data)) continue;
    if (String(tail.rows[i][RESULT_COL.verdict] || '').trim() === 'בוטל') continue;
    markPendingCompleted(data.sessionCode, data.idNumber, pendingSnapshotFromTail(ctx));
    return jsonResponse({ status: 'ok', duplicate: true });
  }

  var attemptNum = countAttempts(data.idNumber, data.license || '', attemptRows(tail)) + 1;
  var row = buildResultRow({
    idNumber: data.idNumber, fullName: data.fullName, phone: data.phone, license: data.license || '',
    score: 0, total: data.totalQuestions || 30, percent: 0, passed: false, time: data.time || '00:00',
    examinerName: data.examinerName, site: data.site, classroom: data.classroom,
    language: data.language || 'he', sessionCode: data.sessionCode, population: data.population,
    audioMode: data.audioMode, device: data.device, languageHistory: null, verified: false
  }, [], attemptNum, '');
  row[RESULT_COL.wrongDetails] = 'סגירת דפדפן באמצע מבחן (נענו ' + (data.answeredCount || 0) + ' שאלות)';
  row[RESULT_COL.audio] = data.audioMode || 'off';
  row[RESULT_COL.verified] = '';
  row[RESULT_COL.langPath] = '';      // a close-fail has no language path to report
  sheet.appendRow(row);

  markPendingCompleted(data.sessionCode, data.idNumber, pendingSnapshotFromTail(ctx));
  return jsonResponse({ status: 'ok' });
}

// The page reloaded rather than closed: undo the premature close-fail. The row
// is marked בוטל, never deleted — deleteRow was the only path that could destroy
// a result — and the examinee goes back to in_exam so they can finish.
defineAction('cancelFailOnClose', { methods: ['POST'], auth: 'examinee', handler: handleCancelFailOnClose });
function handleCancelFailOnClose(data) {
  var sessionCode = String(data.sessionCode || ''), id = normalizeId(data.idNumber || '');
  if (!sessionCode || !id) return jsonResponse({ status: 'ok' });
  var sheet = getSheet('תוצאות');
  var tail = readTail(sheet, RESULT_COL.date);
  for (var i = tail.rows.length - 1; i >= 1; i--) {
    if (!resultRowMatchesExaminee(tail.rows[i], data)) continue;
    if (String(tail.rows[i][RESULT_COL.wrongDetails] || '').indexOf('סגירת דפדפן') !== -1) {
      var sheetRow = i + tail.off + 1;
      sheet.getRange(sheetRow, RESULT_COL.verdict + 1).setValue('בוטל');
      sheet.getRange(sheetRow, RESULT_COL.correctionReason + 1).setValue('בוטל אוטומטי - רענון/חזרה למבחן');
      sheet.getRange(sheetRow, RESULT_COL.correctionDate + 1).setValue(todayStr());
      restorePendingToInExam(sessionCode, data.idNumber);
    }
    break;   // only the most recent row of this examinee
  }
  return jsonResponse({ status: 'ok' });
}

// The mirror of markPendingCompleted for a resumed exam: the newest completed
// row of this examinee goes back to in_exam, through the single status writer.
function restorePendingToInExam(sessionCode, idNumber) {
  var ctx = examineeRowContext(sessionCode, idNumber, true);
  if (!ctx.latest || ctx.latest.status !== 'completed') return;
  setPendingStatus(getSheet('ממתינים'), ctx.latest.rowNumber, sessionCode, 'in_exam', null);
}

// ---- Result-upload token (examiner → results Worker) ------------------------
// A short-lived HMAC the browser sends as X-Auth-Token when it POSTs the result
// HTML to the Cloudflare Worker; the Worker verifies it with the same secret.
// Set ScriptProperty RESULT_UPLOAD_SECRET and the Worker secret UPLOAD_SECRET to
// the same long random string. The secret never reaches the browser.
defineAction('getResultUploadToken', { methods: ['GET'], auth: 'examiner', handler: handleGetResultUploadToken });
function handleGetResultUploadToken() {
  var secret = PropertiesService.getScriptProperties().getProperty('RESULT_UPLOAD_SECRET');
  if (!secret) {
    return jsonResponse({ status: 'error', code: 'not_configured',
      message: 'RESULT_UPLOAD_SECRET not configured in Apps Script properties' });
  }
  var payloadB64 = Utilities.base64EncodeWebSafe(JSON.stringify({ exp: Date.now() + 5 * 60 * 1000 })).replace(/=+$/, '');
  var sigB64 = Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(payloadB64, secret)).replace(/=+$/, '');
  return jsonResponse({ status: 'ok', token: payloadB64 + '.' + sigB64 });
}
