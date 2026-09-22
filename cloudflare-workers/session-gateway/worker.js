/**
 * session-gateway — the exam system's edge. Two jobs, one Worker:
 *
 *   1. It coalesces the examinee poll: ONE upstream call per session per
 *      2 seconds, however many examinees are polling.
 *   2. It serves the question texts, which are PRIVATE Worker assets, handing
 *      each device only the ids its signed grant names.
 *
 * ES module worker; deploy with `npx wrangler deploy` (that also uploads
 * assets/, built by `node tools/build_bank.js`).
 *
 * WHY (1): every approval/status poll used to be its own Apps Script execution
 * (1.5-2.5 s of cold start, one of ~30 slots). 40 waiting examinees = ~480
 * executions per minute for data that is identical for all of them. This
 * gateway asks the server once per session (`action=sessionSnapshot`) and
 * answers every examinee from that snapshot — the poll cost stops scaling with
 * the number of examinees. (DESIGN_2026-09-21 §3.4.)
 *
 * WHY (2): the bank must not be public (decision 19), and it must not cost an
 * Apps Script execution either. The texts ship as Workers Static Assets with
 * `run_worker_first`, so nothing reaches them except this code, and a device
 * gets them only against an HMAC grant that startExam/startPractice/bankGrant
 * issued. The Worker never signs and never trusts the client. (§11.)
 *
 * The poll answers are byte-compatible with the server's own checkApproval /
 * getExamStatus so the client only swaps the URL, never the logic. The
 * examinee token never leaves Apps Script: the snapshot carries a SHA-256 hash
 * and the gateway hashes what the client sent to compare.
 *
 * Endpoints:
 *   GET  /                — health: {status:'ok', service, build, bank}
 *   GET  /v1/poll?kind=approval|status&sessionCode&idNumber&examineeToken
 *                         [&wait=<1-25>&fp=<the last fingerprint>] — long poll
 *   GET  /v1/session/watch?grant=<examiner>&sessionCode=X[&wait=1-25&fp=…]
 *                         — the examiner dashboard's long poll: it answers
 *                           when the session's ROWS or RESULTS change, and
 *                           (against an r32 server) carries the data itself
 *                           in `session` (r31 §13.2, r32 §14.1)
 *   GET  /v1/bank?grant=…[&ids=1,2&langs=he,en]   — texts for the granted ids
 *   GET  /v1/bank/full?grant=…&lang=he            — a whole language (examiner)
 *   POST /v1/invalidate?grant=<examiner>&sessionCode=X[&idNumber&status&…]
 *                         — drop the session snapshot, or patch the decision
 *                           straight into it (see `invalidate`)
 *   POST /v1/invalidate?sessionCode=X&idNumber=Y&examineeToken=T
 *                         — the EXAMINEE's own nudge after a write of theirs
 *                           (registration, submit, "finished on device"); drop
 *                           only, never a patch (r31, §13.5). An id the copy
 *                           does not have yet is a budgeted drop, not a 403
 *                           (r32.1, 22/09): that is the registration nudge
 *   OPTIONS *             — CORS preflight
 *
 * WHY (3) — the HOLD: a client that sends `wait` and the `fp` it already has
 * gets its request HELD here until the answer actually changes (or `wait`
 * elapses). That is ~4-5x fewer requests — the whole account shares a hard
 * 100,000 requests/day on the free plan, and a heavy exam day was already
 * ~130k — AND it puts the examiner's decision on the screen in ~1 s, because
 * nothing waits for the next poll tick any more. Short holds (≤25 s) and a
 * sleep-based loop on purpose: the free plan allows 10 ms of CPU per request,
 * which is what rules out a WebSocket or an SSE stream held for 40 minutes.
 *
 * WHY (4) — the WATCH (r31, "everything fast, no timers"): the examiner's
 * dashboard used to re-read the whole session from Apps Script every 5 s
 * whether or not anything had happened, and still showed a registration ~20 s
 * late. It now holds ONE request here and asks "did this session change?". The
 * answer costs Google nothing extra: the snapshot being fingerprinted is the
 * same one the examinees' own polls keep fresh. Both holds are the SAME loop
 * (`holdUntilChanged`), never a copy.
 *
 * WHY (5) — the watch CARRIES THE DATA (r32, 22/09/2026, §14.1): until r32 a
 * change still cost the dashboard a second round trip to Google
 * (`examinerDashboard`) just to see WHAT changed, and every round trip to
 * Google is a lottery ticket — its delivery hop stalls 25-60 s for our
 * projects (KNOWN_ISSUES #35). An r32 server answers `sessionSnapshot` with
 * `v:2`, the full rows and the session's `results`, so the watch answer now
 * carries `session:{rows,results}` and the dashboard paints from it: ONE round
 * trip per change instead of two. Nothing else about the Worker changes — the
 * examinee's poll answers stay byte-identical (contracts.test.cjs), an r31
 * server (no `v`, no `results`) simply produces no `session`, and an r31 page
 * ignores the field it does not know.
 *
 * Failure policy: a snapshot up to 60 s old is served with `stale:true` rather
 * than an error; with nothing cached the answer is a RETRYABLE error (HTTP
 * 200), never Google's HTML. The client slows down on it but must NOT fall back
 * to direct polling for it, or the storm returns. None of those are ever held:
 * a client that must pace itself has to be told so at once.
 */

const BUILD = '2026-09-23.1';   // r32.1: the registration nudge drops instead of 403

const ALLOWED_ORIGINS = [
  'https://bohanyzahal-cyber.github.io',
  'http://localhost',
  'http://127.0.0.1'
];

const FRESH_MS = 2000;            // a FRESH CHAIN's snapshot: this young answers without asking
const STALE_MS = 60000;           // older than this and we would rather error
const REREAD_GAP_MS = 2000;       // forced upstream re-read: once per session per gap

/**
 * The passive safety cadence — how old a snapshot may get before a request
 * that is RE-ARMING a chain (it sent `fp`, so it already holds exactly this
 * state) asks Google again.
 *
 * WHY (22/09 11:15, KNOWN_ISSUES #35): Google's response-delivery hop stalls
 * 25-60 s for our projects while an idle project in the same account is
 * untouched, and the busiest thing we send Google is this read — every 2 s per
 * session, which is 30/min for a running exam and 10/min for a dashboard
 * nobody is looking at. It does not need to be a clock at all: every write a
 * client makes already announces itself (the examiner's patch/drop, the
 * examinee's nudge after submit/markFinished/registration/DQ), and a fresh
 * chain still demands a copy younger than FRESH_MS. So the Worker reads
 * Google (a) for a fresh chain, (b) when a nudge dropped the snapshot, (c)
 * when a row is missing, and (d) this cadence, for the changes nobody pushed.
 *
 * 20 s, not 45: at 12:29 the same day an examiner's approve request to Google
 * timed out on the page (30 s stall), so the page never got its `ok` and never
 * sent the nudge — and the examinee waited out the whole safety cadence. The
 * page now also nudges when a decision request fails, but the passive floor
 * has to be a number we can live with when BOTH the write and the nudge are
 * lost: 3 reads a minute per held session, still ~10x fewer than the 2 s clock.
 *
 * It MUST stay below STALE_MS, or a held chain would let its copy die before
 * refreshing it and start answering `stale` to everyone.
 */
const HELD_REREAD_MS = 20000;
if (HELD_REREAD_MS >= STALE_MS) throw new Error('HELD_REREAD_MS must stay below STALE_MS');

/** How old a snapshot may be for THIS request: a re-arm trusts what it holds. */
const freshnessFor = clientFp => (clientFp ? HELD_REREAD_MS : FRESH_MS);
const UPSTREAM_TIMEOUT_MS = 25000;
const SESSION_RE = /^[A-Z0-9]{6,8}$/;

// The hold. 25 s is short enough to stay far inside Cloudflare's own limits and
// long enough to cut the request count by 4-5x; one re-evaluation per second is
// what makes a decision another isolate patched in visible "within a second".
// HOLD_MAX_STEPS is a CPU guard, not a timing rule: whatever wakes the loop, it
// evaluates at most this many times and then answers with what it has.
const WAIT_MIN_S = 1;
const WAIT_MAX_S = 25;
const HOLD_TICK_MS = 1000;
const HOLD_MAX_STEPS = 26;

// A hold answers at its deadline plus this grace even when the look it started
// is still in flight. A look may open an upstream read, and a slow Apps Script
// (4-20 s a call on a bad morning) once stretched a 25 s hold to 44 s of wall
// time: the client aborts at 40 s and counts that as a real communication
// failure. The hold is bounded by its own deadline, never by Google's answer.
const HOLD_GRACE_MS = 1000;
// The loser of a `within()` race and nothing else: a value no view can be.
const HOLD_TIMED_OUT = Symbol('hold-timed-out');

// A request that sends no `wait` is not held - but its FIRST look can still
// join an upstream read that Google is taking 20 s to answer, and it must not
// sit there for the 25 s of UPSTREAM_TIMEOUT_MS either. Longer than any
// healthy read (0.7-2 s), shorter than the client's own deadline.
const FIRST_LOOK_MAX_MS = 20000;

// The assets binding is addressed by URL; the host is arbitrary and never
// leaves the isolate. Paths are built from validated numbers only.
const ASSET_ORIGIN = 'https://assets.local';
const ASSET_BATCH = 6;            // parallel asset reads, so 30 ids are ~5 rounds
const EXAMINER_MAX_IDS = 60;      // the wrong-question table never needs more
const LANGS = ['he', 'ru', 'en', 'ar', 'fr', 'es', 'am'];

const GRANT_INVALID = { status: 'error', code: 'grant_invalid' };
const BANK_UNAVAILABLE = { status: 'error', code: 'bank_unavailable', retryable: true };

// handleCheckApproval skips these and keeps looking for an active row; see the
// long comment there about the shared-ID incident that put 'rejected' on it.
const TERMINAL_APPROVALS = { completed: 1, disqualified: 1, cancelled: 1, rejected: 1 };
// ...and these two, when they are the NEWEST row an id has, are the examiner's
// decision about the registration itself, so they are answered as themselves
// instead of 'no registration'. A live row is never either one.
const FINAL_APPROVALS = { rejected: 1, cancelled: 1 };

// The statuses a /v1/invalidate nudge may write into a snapshot — the complete
// set ממתינים ever holds. Anything else is a bug or a probe, and the examinee
// would be answered it as gospel until the next upstream read.
const PATCH_STATUSES = {
  waiting: 1, approved: 1, rejected: 1, cancelled: 1,
  in_exam: 1, completed: 1, disqualified: 1, dq_confirmed: 1
};
// A bound on the minutes a nudge may carry. The real values are 20-120 (the
// server itself caps one time grant at 180); the point is only that an
// unauthenticated caller cannot put an absurd timer in front of an examinee,
// not even for the two seconds until the truth is read.
const MAX_PATCH_MINUTES = 600;

// --- small helpers ---------------------------------------------------------

/** Same as the server's normalizeId: digits only, left-padded to 9. */
function normalizeId(value) {
  let s = String(value == null ? '' : value).replace(/[^0-9]/g, '');
  while (s.length < 9) s = '0' + s;
  return s;
}

async function sha256Hex(text) {
  const bytes = new TextEncoder().encode(text);
  const digest = await crypto.subtle.digest('SHA-256', bytes);
  return [...new Uint8Array(digest)].map(b => b.toString(16).padStart(2, '0')).join('');
}

/** Throws on anything that is not base64url — the caller reads that as "invalid". */
function b64urlBytes(text) {
  const b64 = String(text).replace(/-/g, '+').replace(/_/g, '/');
  const binary = atob(b64 + '==='.slice((b64.length + 3) % 4));
  const out = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) out[i] = binary.charCodeAt(i);
  return out;
}

/** Only whole positive numbers survive: an id is never concatenated into a path raw. */
function toIds(list) {
  if (!Array.isArray(list)) return [];
  const out = [];
  for (const raw of list) {
    const n = Number(raw);
    if (Number.isInteger(n) && n > 0 && n < 1e7) out.push(n);
  }
  return out;
}

/** '' for an absent parameter, so "was it given?" is one truthiness test. */
function param(url, name) {
  return String(url.searchParams.get(name) || '').trim();
}

/**
 * `wait` as milliseconds to hold, clamped to [1 s, 25 s]. 0 means "answer now",
 * and that is what anything else becomes: an empty value, a fraction, a word, a
 * negative. A typo must degrade to today's immediate poll, never to a surprise
 * 25 s hold.
 */
function waitMillis(raw) {
  if (raw == null) return 0;
  const n = Number(String(raw).trim());
  if (!Number.isInteger(n) || n < WAIT_MIN_S) return 0;
  return Math.min(n, WAIT_MAX_S) * 1000;
}

/** A whole number in [min, MAX_PATCH_MINUTES], or null — a bad value is dropped. */
function patchMinutes(raw, min) {
  if (!raw) return null;
  const n = Number(raw);
  return (Number.isInteger(n) && n >= min && n <= MAX_PATCH_MINUTES) ? n : null;
}

function corsHeaders(request) {
  const origin = (request.headers && request.headers.get('Origin')) || '';
  const allowed = ALLOWED_ORIGINS.find(a => origin === a || origin.indexOf(a + ':') === 0);
  return {
    'Access-Control-Allow-Origin': allowed ? origin : ALLOWED_ORIGINS[0],
    'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type',
    'Access-Control-Max-Age': '86400',
    'Vary': 'Origin'
  };
}

function jsonHeaders(request) {
  return Object.assign({
    'Content-Type': 'application/json; charset=utf-8',
    'Cache-Control': 'no-store'
  }, corsHeaders(request));
}

/** For a body that is already JSON text — the bank never re-serialises assets. */
function rawJsonResponse(request, body, status) {
  return new Response(body, { status: status || 200, headers: jsonHeaders(request) });
}

function jsonResponse(request, body, status) {
  return rawJsonResponse(request, JSON.stringify(body), status);
}

const badRequest = (request, message) => jsonResponse(request, { status: 'error', message }, 400);

// --- answers (mirror scanApprovalRows / scanExamStatusRows) ----------------

/** Only a row that stored a token rejects a client that sent a different one. */
function tokenMismatch(row, tokenHex) {
  const stored = String(row.tokenHash || '').trim().toLowerCase();
  return Boolean(stored && tokenHex && stored !== tokenHex);
}

/**
 * null = this id has no row this session, or its newest finished row says
 * nothing to a device that is still polling (completed / disqualified).
 *
 * A LIVE row answers, exactly as the server's scanApprovalRows does. When none
 * is left, the NEWEST finished row decides, and only when it carries the
 * examiner's own decision about the registration: 'rejected' and 'cancelled'
 * are answered as themselves, so the examinee is told what happened instead of
 * being left on a 'שגיאת שרת' banner. Newest — never "the first terminal row
 * the loop likes" — is what keeps the shared-ID incident impossible; the long
 * comment in server/src/50_pending.js tells that story in full.
 */
function approvalAnswer(rows, idNumber, tokenHex) {
  let newestFinished = null;
  for (let i = rows.length - 1; i >= 0; i--) {
    const row = rows[i];
    if (normalizeId(row.id) !== idNumber) continue;
    const approval = String(row.status || 'waiting').trim();
    if (TERMINAL_APPROVALS[approval]) {
      if (!newestFinished) newestFinished = row;   // walking newest→oldest: the first one IS the newest
      continue;
    }
    if (tokenMismatch(row, tokenHex)) {
      return { status: 'error', message: 'טוקן נבחן לא תקין', examineeTokenError: 'mismatch' };
    }
    const answer = {
      status: 'ok',
      approval: approval,
      audioMode: String(row.audio || '').trim() === 'on' ? 'on' : 'off'
    };
    // The server computes the authorised duration; a snapshot without it is a
    // server bug, and omitting the field is safer than inventing minutes.
    const minutes = Number(row.examMinutes);
    if ((approval === 'approved' || approval === 'in_exam') && minutes > 0) answer.examMinutes = minutes;
    return answer;
  }
  if (!newestFinished) return null;
  const decided = String(newestFinished.status || '').trim();
  if (!FINAL_APPROVALS[decided]) return null;
  if (tokenMismatch(newestFinished, tokenHex)) {
    return { status: 'error', message: 'טוקן נבחן לא תקין', examineeTokenError: 'mismatch' };
  }
  // Byte for byte what the server answers here: no audioMode, no examMinutes —
  // there is no exam left to configure (tests/contracts.test.cjs pins it).
  return { status: 'ok', approval: decided };
}

/** null = this id has no row at all in the session. */
function statusAnswer(rows, idNumber, tokenHex) {
  for (let i = rows.length - 1; i >= 0; i--) {
    const row = rows[i];
    if (normalizeId(row.id) !== idNumber) continue;
    if (tokenMismatch(row, tokenHex)) return { status: 'error', examineeTokenError: 'mismatch' };
    return {
      status: 'ok',
      examStatus: String(row.status || '').trim(),
      extraMinutes: Number(row.extraMinutes) || 0
    };
  }
  return null;
}

const NOT_FOUND = {
  approval: { status: 'error', message: 'לא נמצא רישום' },
  status: { status: 'ok', examStatus: 'not_found' }
};

// --- the fingerprint (`fp`) ------------------------------------------------

/**
 * `fp` changes exactly when THIS examinee's answer changes — and for nothing
 * else: not for the snapshot's age, not for another examinee's row, not for
 * `stale` (a stale answer carries the fingerprint of the answer inside it, so
 * a client does not churn its `fp` while the upstream is down).
 *
 *   approval  a:<approval>:<on|off>:<examMinutes|->   a:approved:on:50
 *   status    s:<examStatus>:<extraMinutes>           s:in_exam:7
 *   decided   a:rejected:off:- / a:cancelled:off:-    (no audio, no minutes)
 *   no row    a:none / s:none        token mismatch   a:tok / s:tok
 *   errors    x:up (no snapshot at all)               x:kind / x:sess / x:id
 *
 * It is opaque to the client, which only ever echoes it back, and short because
 * it travels in the query string of every single poll.
 */
const fpWord = value => String(value == null ? '' : value).trim().replace(/:/g, ';').slice(0, 16);

function fingerprint(kind, answer, missing) {
  const tag = kind === 'approval' ? 'a' : 's';
  if (missing) return tag + ':none';
  if (answer.examineeTokenError) return tag + ':tok';
  if (kind === 'approval') {
    return tag + ':' + fpWord(answer.approval) +
      ':' + (answer.audioMode === 'on' ? 'on' : 'off') +
      ':' + (answer.examMinutes > 0 ? fpWord(answer.examMinutes) : '-');
  }
  return tag + ':' + fpWord(answer.examStatus) + ':' + fpWord(Number(answer.extraMinutes) || 0);
}

/**
 * One evaluation of one snapshot for one examinee: the body to answer with, its
 * fingerprint, whether this id had a row at all, and whether a HELD request may
 * keep waiting on it.
 */
function evaluate(kind, snapshot, idNumber, tokenHex, stale) {
  let answer = kind === 'approval'
    ? approvalAnswer(snapshot.rows, idNumber, tokenHex)
    : statusAnswer(snapshot.rows, idNumber, tokenHex);
  const missing = answer === null;
  if (missing) answer = NOT_FOUND[kind];
  if (stale) answer = Object.assign({}, answer, { stale: true });
  return {
    answer: answer,
    fp: fingerprint(kind, answer, missing),
    missing: missing,
    // Both of these END the examinee's polling: "no registration" sends the
    // device back to the code screen, and a rejected/cancelled decision puts a
    // final message in front of them. A snapshot taken before they registered
    // again would end it wrongly, so firstLook re-reads once for either — the
    // same rule the server applies to its own cached snapshot.
    provisional: missing || Boolean(FINAL_APPROVALS[answer.approval]),
    // A stale copy is never held (the client must back off while we cannot
    // refresh it) and neither is a token mismatch (that one never changes by
    // waiting — the device has to stop). "No row yet" IS held: the row appears
    // the moment the examinee registers, which is exactly what to wait for.
    holdable: !stale && !answer.examineeTokenError
  };
}

/** The one answer that means "there is no snapshot at all" — never held. */
const UNAVAILABLE_VIEW = {
  answer: {
    status: 'error', code: 'upstream_unavailable', retryable: true,
    message: 'השרת עמוס — ננסה שוב אוטומטית'
  },
  fp: 'x:up', missing: false, provisional: false, holdable: false
};

// --- the SESSION fingerprint (the examiner's watch, r31 §13.2) -------------

/**
 * What the examiner's dashboard holds on. It changes exactly when something
 * the dashboard DISPLAYS changes, and for nothing else:
 *
 *   s:<12 hex of SHA-256 over the rows, in snapshot order>
 *   s:<12 hex of SHA-256 over `<rows>#<results>`>   …once the server sends
 *             results (r32, v2). With no results the input is the rows alone,
 *             BYTE-IDENTICAL to r31: deploying this Worker ahead of the server
 *             paste (§14.5) must not move a single fingerprint, or two Worker
 *             versions would hash one unchanged session two ways and the
 *             dashboard would flip-flop for ever — the 22/09 lesson below.
 *   s:none    a session with no rows and no results at all — HELD, because the
 *             very next registration creates one and wakes the watch
 *
 * Deliberately NOT in it: `snapshot.at` (every upstream read moves it, and the
 * dashboard would re-read Apps Script every 2 s for nothing — that is the 5 s
 * timer this replaces) and `tokenHash` (the examiner never sees it, and it
 * cannot change inside an attempt anyway). A field an older server does not
 * send — warn/fin/ext/dq arrive with r31's sessionSnapshot (§13.6) — hashes as
 * '', so a Worker deployed ahead of the server simply watches fewer fields and
 * still wakes on a status change or a new registration.
 *
 * 12 hex = 48 bits. This is a change detector, not a security boundary: a
 * collision would only delay one dashboard read until its own safety-net read
 * (examiner.html, 60 s) or the next change.
 */
const SESSION_FP_FIELDS = ['id', 'status', 'audio', 'examMinutes', 'extraMinutes', 'warn', 'fin', 'ext', 'dq'];

/**
 * ...and, since r32 (§14.1), the session's RESULTS — every field of them, in
 * this fixed order. The dashboard now draws the "completed" table from the
 * watch answer, so anything the examiner can see there has to wake the watch:
 * a corrected score, a "sent" tick, a reinstated result. `wrongDetails` is
 * deliberately not in the snapshot at all (the heavy block; the page fetches
 * it on the click that needs it), and `fabricated` is the one bit of it the
 * result row itself shows.
 */
const RESULT_FP_FIELDS = [
  'date', 'idNumber', 'name', 'phone', 'license', 'score', 'percent', 'passed', 'time',
  'examiner', 'site', 'classroom', 'language', 'attempt', 'sent', 'disqualified', 'waLink',
  'population', 'corrected', 'audioMode', 'verified', 'suspicious', 'device', 'fabricated'
];

// The four r31 fields, where "absent" and "zero" are the SAME state and must
// hash the same. WHY (22/09, from `wrangler tail`): one examiner page saw its
// answer alternate between two fingerprints for the same unchanged session,
// 250 ms apart, in ~2 s bursts. The exam project moved r30 -> r31 that
// morning, so two snapshots of DIFFERENT SHAPE were alive at once - an old one
// with no warn/fin/ext/dq (hashing '') beside a new one carrying 0 (hashing
// '0') - and `lookAtNewest` answers whichever copy DIFFERS from what the
// client holds, so memory and cache flip-flopped the dashboard forever. Only
// these four collapse: `extraMinutes: 0` has always been sent, and `id` /
// `status` / `audio` / `examMinutes` have no "absent" state.
const SESSION_FP_ZEROABLE = { warn: 1, fin: 1, ext: 1, dq: 1 };

/**
 * One cell of the fingerprint's input, for a pending row or for a result.
 * Absent, null and '' are the same empty cell; everything else is its own
 * `String()` — `true`/`false` included, which is how a result's booleans
 * (`passed`, `sent`, `corrected`, `verified`, `suspicious`, `fabricated`)
 * hash. Deterministic is the whole requirement: the same snapshot must hash
 * the same in every isolate, for ever.
 */
function fpCell(row, field) {
  const value = row[field];
  if (value == null || value === '') return '';
  if (SESSION_FP_ZEROABLE[field] && (value === 0 || value === '0' || value === false)) return '';
  return String(value);
}

/** `<rows>` or `<rows>#<results>` — the exact text the session hash is taken over. */
function fingerprintInput(rows, results) {
  const rowsPart = rows.map(row => SESSION_FP_FIELDS.map(field => fpCell(row, field)).join('|')).join(';');
  if (!results.length) return rowsPart;   // r31's input, to the byte
  return rowsPart + '#' +
    results.map(res => RESULT_FP_FIELDS.map(field => fpCell(res, field)).join('|')).join(';');
}

// Per snapshot OBJECT, so a request held for 25 s hashes each copy once:
// memory hands back the same object on every tick of the hold. A WeakMap, so
// a snapshot that falls out of memory takes its entry with it; and the PROMISE
// is what is cached, so two holds that meet on one snapshot share one digest.
const sessionFpCache = new WeakMap();

/**
 * `snapshot.sfp` is the fingerprint stamped ONCE, where the snapshot was
 * stored (fetchCoalesced / patchSnapshot), and it travels into caches.default
 * inside the JSON. WHY: a held watch re-reads the cached copy every second,
 * and `caches.default` hands back a NEW object each time, so the WeakMap could
 * never hit and every tick paid for another SHA-256 - 7-12 ms of CPU per held
 * request, against a free-plan ceiling of 10 ms (`wrangler tail`, 22/09).
 * A copy written by an older Worker has no `sfp`, and is hashed here as before.
 */
function sessionFingerprint(snapshot) {
  const stamped = snapshot && snapshot.sfp;
  if (typeof stamped === 'string' && stamped) return Promise.resolve(stamped);
  const known = sessionFpCache.get(snapshot);
  if (known) return known;
  const rows = Array.isArray(snapshot.rows) ? snapshot.rows : [];
  const results = Array.isArray(snapshot.results) ? snapshot.results : [];
  const pending = (rows.length || results.length)
    ? sha256Hex(fingerprintInput(rows, results)).then(hex => 's:' + hex.slice(0, 12))
    : Promise.resolve('s:none');
  sessionFpCache.set(snapshot, pending);
  return pending;
}

/**
 * One evaluation of one snapshot for the examiner: how many rows the session
 * has and when it was read — the change detector. The DATA is attached later,
 * in `watch`, and only when it is worth sending (see `sessionPayload`): the
 * view is computed once per tick of a hold, so building the payload here would
 * copy every row 25 times for an answer that carries it at most once.
 *
 * A stale copy answers at once with `stale:true` and is NEVER held, exactly
 * like a stale poll answer: while we cannot refresh, the dashboard has to fall
 * back to its own safety-net read instead of parking here.
 */
async function sessionView(snapshot, stale) {
  const rows = Array.isArray(snapshot.rows) ? snapshot.rows : [];
  const answer = { status: 'ok', rows: rows.length, at: Number(snapshot.at) || 0 };
  if (stale) answer.stale = true;
  return {
    answer: answer, fp: await sessionFingerprint(snapshot), holdable: !stale,
    // The copy this view was computed from, and whether it was servable only
    // as stale — `watch` needs both to decide about `session`.
    snapshot: snapshot, stale: Boolean(stale)
  };
}

/** The results of the copy a view was computed from — for the log, sent or not. */
function sessionResults(view) {
  const results = view && view.snapshot && view.snapshot.results;
  return Array.isArray(results) ? results : [];
}

/** A row as the examiner may see it: everything the server sent, minus the hash. */
function publicRow(row) {
  const copy = Object.assign({}, row);
  delete copy.tokenHash;   // the token never leaves Apps Script, and its hash never leaves here
  return copy;
}

/**
 * `session:{rows,results}` — the dashboard's data, or null when it must not be
 * sent. Three gates, all from §14.1:
 *
 *   1. `v >= 2`: an r31 server sends neither the extra row fields nor the
 *      results, and half a dashboard is worse than none — the page falls back
 *      to its `examinerDashboard` read, exactly as in r31.3.
 *   2. not `stale`: a stale view is one we could not refresh, and the page
 *      must not paint a class from a copy that may be a minute behind. It
 *      falls back to its own safety net instead (§13.2).
 *   3. the fingerprint actually MOVED (or the client sent none, i.e. this is
 *      the first answer of a chain and the page has nothing yet). A hold that
 *      runs out unchanged answers the same `fp` and must stay small — that is
 *      one answer per 25 s per open dashboard, with nothing to say.
 */
function sessionPayload(view, clientFp) {
  const snapshot = view && view.snapshot;
  if (!snapshot || view.stale) return null;
  if (!(Number(snapshot.v) >= 2)) return null;
  if (clientFp && view.fp === clientFp) return null;
  return {
    rows: (Array.isArray(snapshot.rows) ? snapshot.rows : []).map(publicRow),
    results: Array.isArray(snapshot.results) ? snapshot.results : []
  };
}

// --- the gateway -----------------------------------------------------------

/**
 * Dependency-injected so tests can drive it with a fake clock, a counting
 * fetch, an in-memory ASSETS binding and a fake `sleep` (a held request must be
 * testable without waiting 25 real seconds). The state (in-flight map, memory
 * snapshots, waiters, imported HMAC key) lives in the closure, so the
 * module-level singleton below coalesces across requests of one isolate while
 * each test gets its own clean instance.
 */
export function createGateway({ fetch, caches, now, env, sleep, log }) {
  const clock = now || (() => Date.now());
  // The default timer carries its own `cancel`, and so does the tests' fake
  // one: a timer that LOST its race must go, or every held request leaves a
  // trail of live timers behind it for as long as the isolate lives.
  const nap = sleep || (ms => {
    let id;
    const timer = new Promise(resolve => { id = setTimeout(resolve, ms); });
    timer.cancel = () => clearTimeout(id);
    return timer;
  });
  // One line per answered request, so `wrangler tail` can explain a
  // fingerprint that moved without a write (22/09). Injected in tests so the
  // suite stays silent and can assert on what was written.
  const emit = log || (line => { try { console.log(line); } catch (e) { /* logging never fails a request */ } });
  const memory = new Map();   // session -> { at, snapshot }
  const inflight = new Map(); // session -> Promise<snapshot|null>
  const lastUpstream = new Map(); // session -> { ms, sfp } of the last read, for the log
  const lastForced = new Map(); // session -> ms of the last forced re-read
  const waiters = new Map();  // session -> Set<resolve> — the held requests

  const cacheKey = (kind, session) =>
    'https://session-gateway.internal/' + kind + '/' + encodeURIComponent(session);

  /**
   * `maxAgeMs` is checked against `rat` — the time THIS gateway made the copy
   * current (an upstream read, or a patch), stamped into the body before it is
   * written. The cache entry's own max-age is the outer bound (HELD_REREAD_MS
   * for 'snap' since 22/09, so another isolate can serve a re-armed chain
   * without asking Google); a fresh chain narrows it to FRESH_MS here. A copy
   * written before r31.2 has no `rat` and is taken at the entry's word.
   */
  async function cacheRead(kind, session, maxAgeMs) {
    if (!caches || !caches.default) return null;
    try {
      const hit = await caches.default.match(new Request(cacheKey(kind, session)));
      if (!hit) return null;
      const body = await hit.json();
      if (!body || !Array.isArray(body.rows)) return null;
      if (maxAgeMs != null && Number.isFinite(body.rat) && clock() - body.rat > maxAgeMs) return null;
      return body;
    } catch (e) {
      return null;
    }
  }

  /**
   * The stored Response carries the fingerprint and the read time in HEADERS
   * as well as in the body (r32, TODO 0א.6). WHY: a held request re-reads
   * `caches.default` every second — that is the channel a patch from another
   * isolate arrives through — and `JSON.parse` of a 5-40 KB snapshot every
   * second is most of what a held request costs, against a free-plan ceiling
   * of 10 ms of CPU. With these two headers the usual tick is `match` plus a
   * string compare, and the body is parsed only when the fingerprint says the
   * copy really is a different one. See `cacheFingerprint`.
   */
  async function cacheWrite(session, snapshot) {
    if (!caches || !caches.default) return;
    const body = JSON.stringify(snapshot);
    const store = (kind, maxAge) => caches.default.put(
      new Request(cacheKey(kind, session)),
      new Response(body, {
        headers: {
          'Content-Type': 'application/json',
          'Cache-Control': 'max-age=' + maxAge,
          'X-SFP': String(snapshot.sfp || ''),
          'X-RAT': String(Number(snapshot.rat) || 0)
        }
      })
    );
    try {
      await Promise.all([store('snap', HELD_REREAD_MS / 1000), store('stale', STALE_MS / 1000)]);
    } catch (e) { /* cache is best effort */ }
  }

  /**
   * The cached copy's fingerprint and read time from its HEADERS ALONE — never
   * `json()`, which is the whole point (see `cacheWrite`). `null` means "not
   * known": no entry, an entry past `maxAgeMs`, or one written before r32 and
   * therefore carrying no headers. Every caller falls back to the full read on
   * `null`, so a Worker rolled back mid-deploy behaves exactly as r31 did.
   */
  async function cacheFingerprint(kind, session, maxAgeMs) {
    if (!caches || !caches.default) return null;
    try {
      const hit = await caches.default.match(new Request(cacheKey(kind, session)));
      if (!hit || !hit.headers) return null;
      const sfp = hit.headers.get('X-SFP') || '';
      const rat = Number(hit.headers.get('X-RAT'));
      if (!sfp || !Number.isFinite(rat) || rat <= 0) return null;
      if (maxAgeMs != null && clock() - rat > maxAgeMs) return null;
      return { sfp: sfp, rat: rat };
    } catch (e) {
      return null;
    }
  }

  function memoryRead(session, maxAgeMs) {
    const hit = memory.get(session);
    if (!hit || clock() - hit.at > maxAgeMs) return null;
    return hit.snapshot;
  }

  /**
   * A copy that came out of `caches.default` while this isolate held nothing
   * is kept in memory too (r32). WHY: without it every tick of a held request
   * in a cold isolate re-read and re-parsed the same cached body; with it the
   * next tick finds the copy in memory and only compares `cacheFingerprint`'s
   * header against it. Its age is the age it already had (`rat`, OUR clock
   * when the copy was made current), never `now` — adopting a copy must not
   * make it look younger than it is, or the safety cadence would never fire.
   */
  function adopt(session, snapshot) {
    if (!snapshot) return snapshot;   // "nothing in the cache either" passes straight through
    const rat = Number(snapshot.rat);
    memory.set(session, { at: Number.isFinite(rat) && rat > 0 ? rat : clock(), snapshot: snapshot });
    return snapshot;
  }

  /** How old the copy we are about to answer from is — for the log, only. */
  const snapshotAge = snapshot => {
    const made = Number(snapshot && snapshot.rat);
    return Number.isFinite(made) ? Math.max(0, clock() - made) : -1;
  };

  /**
   * Awaits `promise` for at most `ms` on a timer THIS request owns, and drops
   * that timer as soon as the race is decided. HOLD_TIMED_OUT = the timer won.
   *
   * WHY every join of a shared promise must go through here (22/09, from
   * `wrangler tail`): 15 of 74 requests died as HTTP 500 with outcome
   * "exception" — «the Workers runtime canceled this request because it
   * detected that your Worker's code had hung and would never generate a
   * response» — after 2-3 ms of wall time and 0 CPU. That is Cloudflare's hang
   * detection, and it fires on a request that is awaiting a promise ANOTHER
   * request created (our coalesced `inflight` read, kept alive by that other
   * request's waitUntil) while having no pending I/O or timer of its own. The
   * timer here IS that pending timer, and it also bounds the wait.
   */
  async function within(ms, promise) {
    const timer = nap(Math.max(0, ms));
    try {
      return await Promise.race([promise, timer.then(() => HOLD_TIMED_OUT)]);
    } finally {
      if (timer && typeof timer.cancel === 'function') timer.cancel();
    }
  }

  /** A promise this request walked away from: let it finish for whoever is next. */
  function abandon(promise, ctx) {
    const orphan = promise.catch(() => { /* a late failure is no longer ours */ });
    if (ctx && typeof ctx.waitUntil === 'function') ctx.waitUntil(orphan);
  }

  // --- waking the held requests --------------------------------------------

  /**
   * Everything that can change a session's answer ends here: a patch, a drop,
   * and a finished upstream read. Every request held on this session wakes and
   * re-evaluates at once, so an examiner's decision reaches an examinee served
   * by THIS isolate in ~0 ms, and one served by another within one tick.
   *
   * The set is detached before it is resolved: a waiter that immediately parks
   * again registers in a new one and cannot be woken twice by the same event.
   */
  function wake(session) {
    const parked = waiters.get(session);
    if (!parked || !parked.size) return;
    waiters.delete(session);
    for (const resolve of parked) resolve();
  }

  /**
   * Parks until the session changes or `ms` passes, whichever is first, and
   * ALWAYS takes its resolver back out — a leaked resolver would pin the
   * session's Set for the life of the isolate.
   */
  function waitForChange(session, ms) {
    let parked = waiters.get(session);
    if (!parked) { parked = new Set(); waiters.set(session, parked); }
    let resolve;
    const woken = new Promise(r => { resolve = r; });
    parked.add(resolve);
    const timer = nap(ms);
    return Promise.race([woken, timer]).then(() => {
      if (timer && typeof timer.cancel === 'function') timer.cancel();
      const current = waiters.get(session);
      if (!current) return;
      current.delete(resolve);
      if (!current.size) waiters.delete(session);
    });
  }

  /** Never rejects: a failed upstream is `null`, and null means "ask the cache". */
  async function fetchUpstream(session) {
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), UPSTREAM_TIMEOUT_MS);
    try {
      const url = env.API_URL +
        (String(env.API_URL).indexOf('?') >= 0 ? '&' : '?') +
        'action=sessionSnapshot&sessionCode=' + encodeURIComponent(session) +
        '&gatewayKey=' + encodeURIComponent(env.GATEWAY_KEY || '') +
        '&origin=gateway';
      const res = await fetch(url, {
        method: 'GET',
        redirect: 'follow',
        signal: controller.signal,
        headers: { 'Accept': 'application/json' }
      });
      if (!res || res.status !== 200) return null;
      // Apps Script answers an overload with an HTML error page — JSON.parse
      // throws on it and we fall through to the stale copy.
      const data = JSON.parse(await res.text());
      if (!data || data.status !== 'ok' || !Array.isArray(data.rows)) return null;
      // `v` and `results` arrive with r32's sessionSnapshot (§14.1). An r31
      // server sends neither: v = 1 and no results, which is exactly what
      // keeps this Worker's fingerprints identical to r31's (see
      // sessionFingerprint) and its watch answers free of `session`.
      return {
        at: Number(data.at) || clock(),
        rows: data.rows,
        v: Number(data.v) || 1,
        results: Array.isArray(data.results) ? data.results : []
      };
    } catch (e) {
      return null;
    } finally {
      clearTimeout(timer);
    }
  }

  /** One live upstream request per session, whoever asks. */
  function fetchCoalesced(session) {
    const pending = inflight.get(session);
    if (pending) return pending;
    const startedAt = clock();
    // The cache write is awaited, not floating: a Worker may be torn down as
    // soon as it answers, and a dropped write means the next isolate asks
    // Apps Script again — exactly what this gateway exists to prevent.
    const promise = fetchUpstream(session).then(async snapshot => {
      inflight.delete(session);
      if (snapshot) {
        // Stamped HERE, once, before it is stored: `sfp` then travels into
        // caches.default inside the JSON, so a held watch that re-reads the
        // cached copy every second never hashes anything (see
        // sessionFingerprint — this is the 7-12 ms of CPU it costs otherwise).
        snapshot.sfp = await sessionFingerprint(snapshot);
        snapshot.rat = clock();   // OUR clock, so another isolate can age it
        lastUpstream.set(session, { ms: Math.max(0, clock() - startedAt), sfp: snapshot.sfp });
        memory.set(session, { at: clock(), snapshot });
        await cacheWrite(session, snapshot);
        wake(session); // fresh truth: whoever is held on this session re-reads it
      }
      return snapshot;
    });
    inflight.set(session, promise);
    return promise;
  }

  /**
   * { ok, snapshot, fromCache, stale } — `ok:false` only when nothing exists.
   *
   * Bounded by `until`, which is THIS request's own deadline. WHY (22/09):
   * this is the first look, it runs BEFORE the hold, and it used to await the
   * coalesced read bare — so a request that joined a read Google was taking
   * 20 s over started its 25 s hold 20 s late and answered after 45 s, past
   * the client's own 40 s abort (four such requests in one 24-minute tail).
   * The budget is the request's, not Google's: when the read does not land in
   * time we answer from the best copy we have, marked stale, which is never
   * held. The read itself is abandoned, not cancelled — it is coalesced, and
   * whoever asks next is served by it.
   */
  async function loadSnapshot(session, forceFresh, until, ctx, trace, maxAgeMs) {
    if (!forceFresh) {
      const fresh = memoryRead(session, maxAgeMs);
      if (fresh) { trace.src = 'memory'; trace.age = snapshotAge(fresh); return { ok: true, snapshot: fresh, fromCache: true, stale: false }; }
      const cached = await cacheRead('snap', session, maxAgeMs);
      if (cached) {
        adopt(session, cached);   // so the next tick of this request compares a header, not a body
        trace.src = 'cache'; trace.age = snapshotAge(cached);
        return { ok: true, snapshot: cached, fromCache: true, stale: false };
      }
    }
    const pending = fetchCoalesced(session);
    const fetched = await within(Math.max(0, until - clock()) + HOLD_GRACE_MS, pending);
    if (fetched === HOLD_TIMED_OUT) {
      abandon(pending, ctx);
      trace.late = 1;
    } else if (fetched) {
      trace.src = 'fetch';
      trace.age = snapshotAge(fetched);
      return { ok: true, snapshot: fetched, fromCache: false, stale: false };
    }
    const old = memoryRead(session, STALE_MS) || adopt(session, await cacheRead('stale', session));
    if (old) { trace.src = 'stale'; trace.age = snapshotAge(old); return { ok: true, snapshot: old, fromCache: true, stale: true }; }
    trace.src = 'none';
    trace.age = -1;
    return { ok: false };
  }

  /**
   * One extra upstream read per session per gap, whatever asked for it: the
   * server's own re-read before it says "no registration", and the examiner's
   * /v1/invalidate after a decision. They share the budget on purpose — both
   * mean "the snapshot is behind", and together they must still not cost more
   * than one additional Apps Script execution per session per REREAD_GAP_MS.
   */
  function mayForceReread(session) {
    const last = lastForced.get(session);
    if (last != null && clock() - last < REREAD_GAP_MS) return false;
    lastForced.set(session, clock());
    return true;
  }

  /**
   * Writes a decision into the session's newest row for one id and stores the
   * result as a FRESH snapshot, so the next poll answers it with no upstream
   * call at all. `false` = there was nothing to patch (no snapshot for the
   * session, or this id has no row in it) and the caller drops instead.
   *
   * The whole snapshot is re-stamped fresh, not just the row: for the next
   * FRESH_MS the other rows are served as they were read, which is at most the
   * age they already had. The copy dies on the normal clock, so the poll after
   * it reads the truth upstream either way.
   */
  async function patchSnapshot(session, idNumber, fields) {
    const current = memoryRead(session, STALE_MS)
      || await cacheRead('snap', session)
      || await cacheRead('stale', session);
    if (!current || !Array.isArray(current.rows)) return false;
    // Sheet order, oldest first — the last match is the live attempt, exactly
    // the row approvalAnswer/statusAnswer would have answered from.
    let index = -1;
    for (let i = current.rows.length - 1; i >= 0; i--) {
      if (normalizeId(current.rows[i].id) === idNumber) { index = i; break; }
    }
    if (index < 0) return false;
    const rows = current.rows.slice();
    rows[index] = Object.assign({}, rows[index], fields);  // only the named row
    // `at` stays the read time; `v` and `results` travel with the copy (r32) —
    // a patch touches ONE pending row and knows nothing about the results, so
    // dropping them would blank the dashboard's completed table on every
    // decision and move the fingerprint for a change nobody made.
    const snapshot = {
      at: current.at, rows: rows,
      v: Number(current.v) || 1,
      results: Array.isArray(current.results) ? current.results : []
    };
    // Re-stamped, never inherited: the whole point of a patch is that the
    // session's fingerprint moves, and the copy in caches.default must carry
    // the NEW one (see sessionFingerprint).
    snapshot.sfp = await sessionFingerprint(snapshot);
    snapshot.rat = clock();   // a confirmed write makes the copy current again
    memory.set(session, { at: clock(), snapshot: snapshot });
    await cacheWrite(session, snapshot);
    wake(session); // a request held on this session answers the decision NOW
    return true;
  }

  /** The next poll of this session reads upstream instead of a stale snapshot. */
  async function dropSnapshot(session) {
    memory.delete(session);
    if (caches && caches.default) {
      try {
        await Promise.all(['snap', 'stale'].map(kind =>
          caches.default.delete(new Request(cacheKey(kind, session)))));
      } catch (e) { /* cache is best effort */ }
    }
    wake(session); // held requests re-read upstream instead of waiting it out
  }

  // --- the private question bank (assets) ----------------------------------

  const hasAssets = () => Boolean(env.ASSETS && typeof env.ASSETS.fetch === 'function');

  /** null = no binding, or the read threw. A 404 comes back as a real Response. */
  async function assetFetch(pathname) {
    if (!hasAssets()) return null;
    try {
      return await env.ASSETS.fetch(ASSET_ORIGIN + pathname);
    } catch (e) {
      return null;
    }
  }

  // The build id of the deployed bank, read once per isolate. A failed read is
  // not remembered: a blip must not pin '' for the life of the isolate.
  let buildPromise = null;
  function bankBuild() {
    if (!buildPromise) {
      buildPromise = (async () => {
        const res = await assetFetch('/manifest.json');
        if (!res || res.status !== 200) return '';
        try {
          return String((JSON.parse(await res.text()) || {}).build || '');
        } catch (e) {
          return '';
        }
      })().then(build => {
        if (!build) buildPromise = null;
        return build;
      });
    }
    return buildPromise;
  }

  // The imported KEY is cached, never the promise that produces it. Once it is
  // in hand every later grant check is synchronous; until then each request
  // imports its own. WHY not share the promise (22/09): a request that awaits
  // a promise ANOTHER request created, with no I/O or timer of its own, is
  // killed by the runtime's hang detection — and this await is the very first
  // thing a watch or a nudge does. An import is cheap; being cancelled is not.
  // An empty secret is kept out of importKey, which rejects it.
  let cryptoKey = null;
  async function hmacKey() {
    if (cryptoKey) return cryptoKey;
    const imported = await crypto.subtle.importKey(
      'raw', new TextEncoder().encode(String(env.GATEWAY_KEY || '')),
      { name: 'HMAC', hash: 'SHA-256' }, false, ['verify']
    ).catch(() => null);
    if (imported) cryptoKey = imported;
    return imported;
  }

  /**
   * `<payload>.<sig>` — payload is base64url JSON, sig is base64url
   * HMAC-SHA256(GATEWAY_KEY, payload). Only Apps Script ever signs one
   * (§11.2). Returns the payload, or null for anything at all wrong: a bad
   * signature, a stale `exp`, a scope this route does not serve.
   */
  async function verifyGrant(raw, scopes) {
    if (!env.GATEWAY_KEY) return null;
    const parts = String(raw || '').split('.');
    if (parts.length !== 2 || !parts[0] || !parts[1]) return null;
    try {
      const key = await hmacKey();
      if (!key) return null;
      // crypto.subtle.verify compares the MACs in constant time.
      const ok = await crypto.subtle.verify(
        'HMAC', key, b64urlBytes(parts[1]), new TextEncoder().encode(parts[0]));
      if (!ok) return null;
      const payload = JSON.parse(new TextDecoder().decode(b64urlBytes(parts[0])));
      if (!payload || payload.v !== 1) return null;
      if (!(Number(payload.exp) > clock())) return null;
      if (scopes.indexOf(payload.s) < 0) return null;
      return payload;
    } catch (e) {
      return null; // malformed base64url or JSON is just an invalid grant
    }
  }

  /** { text } served · {} no such asset · { failed } the read itself broke. */
  async function readQuestion(id) {
    const res = await assetFetch('/q/' + id + '.json');
    if (!res) return { failed: true };
    if (res.status !== 200) return {};
    try {
      return { text: await res.text() };
    } catch (e) {
      return { failed: true };
    }
  }

  /** ASSET_BATCH connections at a time, so 30 ids are five short rounds. */
  async function readQuestions(ids) {
    const texts = [];
    const missing = [];
    let failed = 0;
    for (let i = 0; i < ids.length; i += ASSET_BATCH) {
      const batch = ids.slice(i, i + ASSET_BATCH);
      const results = await Promise.all(batch.map(readQuestion));
      results.forEach((result, n) => {
        if (result.text !== undefined) return texts.push(result.text);
        missing.push(batch[n]);
        if (result.failed) failed++;
      });
    }
    return { texts, missing, failed };
  }

  /**
   * Examiner tooling only: a handful of ids, so one parse each is affordable.
   * The exam path never does this — it concatenates the asset text untouched.
   */
  function filterLangs(text, langs) {
    try {
      const question = JSON.parse(text);
      const kept = {};
      for (const lang of langs) if (question.l && question.l[lang]) kept[lang] = question.l[lang];
      return JSON.stringify({ id: question.id, l: kept });
    } catch (e) {
      return text; // an unparsable asset of ours: serve it whole rather than 500
    }
  }

  async function bank(request, url) {
    const grant = await verifyGrant(url.searchParams.get('grant'), ['exam', 'practice', 'examiner']);
    if (!grant) return jsonResponse(request, GRANT_INVALID, 403);

    let ids;
    let langs = null;
    if (grant.s === 'examiner') {
      ids = toIds(String(url.searchParams.get('ids') || '').split(','));
      if (!ids.length) return badRequest(request, 'ids is required for an examiner grant');
      if (ids.length > EXAMINER_MAX_IDS) return badRequest(request, 'at most ' + EXAMINER_MAX_IDS + ' ids');
      const rawLangs = url.searchParams.get('langs');
      if (rawLangs != null) {
        langs = String(rawLangs).split(',').map(l => l.trim()).filter(l => LANGS.indexOf(l) >= 0);
        if (!langs.length) return badRequest(request, 'langs names no known language');
      }
    } else {
      // The grant IS the authorisation: an exam device gets its 30 ids, nothing
      // else, whatever it puts in the query string.
      ids = toIds(grant.ids);
      if (!ids.length) return jsonResponse(request, GRANT_INVALID, 403);
    }

    if (!hasAssets()) return jsonResponse(request, BANK_UNAVAILABLE, 503);
    const build = bankBuild(); // in flight beside the questions, not after them
    const { texts, missing, failed } = await readQuestions(ids);
    // Every single read broke — that is the binding or the deploy, not the ids.
    if (failed === ids.length) return jsonResponse(request, BANK_UNAVAILABLE, 503);

    const questions = langs ? texts.map(text => filterLangs(text, langs)) : texts;
    return rawJsonResponse(request,
      '{"status":"ok","build":' + JSON.stringify(await build) +
      ',"questions":[' + questions.join(',') +
      '],"missing":' + JSON.stringify(missing) + '}');
  }

  async function bankFull(request, url) {
    const grant = await verifyGrant(url.searchParams.get('grant'), ['examiner']);
    if (!grant) return jsonResponse(request, GRANT_INVALID, 403);
    const lang = String(url.searchParams.get('lang') || '').trim();
    if (LANGS.indexOf(lang) < 0) return badRequest(request, 'lang must be one of ' + LANGS.join(','));

    if (!hasAssets()) return jsonResponse(request, BANK_UNAVAILABLE, 503);
    const res = await assetFetch('/bank/' + lang + '.json');
    if (!res) return jsonResponse(request, BANK_UNAVAILABLE, 503);
    if (res.status !== 200) {
      return jsonResponse(request, { status: 'error', code: 'bank_missing', message: 'no bank for ' + lang }, 404);
    }
    // Streamed, not buffered: 0.6-1.1 MB per language, and nothing here needs
    // to look inside it. The asset response's own headers are kept as the base
    // so that whatever content encoding came with that body stays with it; only
    // ours are overwritten on top.
    const out = new Response(res.body, res);
    for (const [name, value] of Object.entries(jsonHeaders(request))) out.headers.set(name, value);
    return out;
  }

  /**
   * The EXAMINEE's own nudge, after a write of theirs (r31, §13.5): the result
   * POST that answered `ok`, and `markFinished`. Until r31 only the examiner
   * pushed, so a submit reached the dashboard only on the Worker's next read
   * of Google (≤2 s) and then on the dashboard's next tick (≤5 s).
   *
   * It is authenticated by the examinee's OWN token — the one thing this
   * device has that nobody else does. The snapshot carries `tokenHash`
   * (SHA-256, computed in Apps Script), so the comparison is hash to hash and
   * the token itself still never leaves Google. The newest row of that id is
   * the live attempt, exactly the row the poll answers from.
   *
   * It DROPS and never patches: a patch writes into what every examinee of the
   * session is then answered from, and only the examiner may do that. The drop
   * costs at most one extra upstream read per REREAD_GAP_MS — the same budget
   * as the examiner's plain nudge, deliberately shared, so a device looping on
   * it cannot buy more Apps Script executions than one examiner clicking.
   */
  async function examineeNudge(request, session, rawId, token) {
    // Memory → snap → stale, exactly like patchSnapshot: whatever copy the
    // session has is the one the row has to be found in.
    const current = memoryRead(session, STALE_MS)
      || await cacheRead('snap', session)
      || await cacheRead('stale', session);
    // Nothing cached for this session: there is nothing to drop, nobody parked
    // on it, and the re-read budget stays whole for the examiner.
    if (!current || !Array.isArray(current.rows)) {
      return jsonResponse(request, { status: 'ok', dropped: false });
    }

    const idNumber = normalizeId(rawId);
    let newest = null;
    for (let i = current.rows.length - 1; i >= 0; i--) {
      if (normalizeId(current.rows[i].id) === idNumber) { newest = current.rows[i]; break; }
    }
    // No row for this id in the copy we hold (r32.1, 22/09/2026 15:25 live):
    // the nudge that follows a REGISTRATION arrives before this Worker has read
    // the row it announces, so there is no hash to check it against. Refusing
    // it (403, as until now) left the examiner AND the phone waiting out the
    // 20 s safety read - the whole point of the nudge, lost. The only thing
    // this door can do is a drop, and a drop is budgeted (one upstream read per
    // REREAD_GAP_MS per session, the same budget as the examiner's plain nudge)
    // and needs the session code: an unknown id cannot cost Google more than
    // one examiner clicking, and it can never write anything into an answer.
    // A row that IS here with a different hash stays a refusal below: that is
    // a stolen or stale token, not a row we have not seen yet.
    if (!newest) {
      const droppedUnknown = mayForceReread(session);
      if (droppedUnknown) await dropSnapshot(session);
      return jsonResponse(request, { status: 'ok', dropped: droppedUnknown });
    }
    // A row with no stored hash cannot authenticate anyone: that is a refusal,
    // not a free pass (tokenMismatch may let it through on the READ path, where
    // the worst case is answering an old sheet's row).
    const stored = String(newest.tokenHash || '').trim().toLowerCase();
    if (!stored || stored !== await sha256Hex(token)) {
      return jsonResponse(request, GRANT_INVALID, 403);
    }

    const dropped = mayForceReread(session);
    if (dropped) await dropSnapshot(session);
    return jsonResponse(request, { status: 'ok', dropped: dropped });
  }

  /**
   * The examiner's nudge after a decision. Two shapes:
   *
   *   ?sessionCode=X                    drop the snapshot, so the examinee's
   *                                     next poll reads the server at once
   *                                     instead of waiting out FRESH_MS.
   *   …&idNumber=Y&status=S[&extraMinutes&examMinutes&audio]
   *                                     PATCH: write the decision into the
   *                                     cached snapshot, so the next poll —
   *                                     about a second after the click —
   *                                     answers it with ZERO upstream calls.
   *
   * WHY a patch is not a lie: examiner.html fires the nudge only AFTER Apps
   * Script answered `status:'ok'`, i.e. after the row was written, so the patch
   * can never be ahead of the truth by more than that one confirmed write. And
   * it is short lived — the patched copy expires on the normal FRESH_MS clock
   * and the poll behind it reads the row from the server, which overwrites it
   * either way. This is what turns the 2-4 s of DESIGN §11.8 into ≈1 s for an
   * approval without a push channel; the floor that remains is Google's own
   * ~1 s sheet write.
   *
   * Always HTTP 200 `{status:'ok'}` for a well-formed nudge: the examiner fires
   * it and forgets it. `patched` tells a caller that ASKED for a patch whether
   * it landed; a plain invalidate keeps its old body exactly.
   */
  async function invalidate(request, url) {
    const session = param(url, 'sessionCode');
    const rawId = param(url, 'idNumber');
    const status = param(url, 'status');
    const examineeToken = param(url, 'examineeToken');

    // The examinee's door, before the examiner's: no grant and NO `status` at
    // all — a device may only ever say "look again", never what to look at.
    // A request carrying `status` without a grant falls through to the
    // examiner's check below and is refused there, and so is one with no
    // token: these three parameters and no others open this door.
    if (!param(url, 'grant') && !status && session && rawId && examineeToken) {
      if (!SESSION_RE.test(session)) return badRequest(request, 'קוד סשן לא תקין');
      return examineeNudge(request, session, rawId, examineeToken);
    }

    // An examiner-only door. A nudge writes a status straight into what the
    // examinees are answered from: without this check anyone who knows a
    // session code and a classmate's id could show them "disqualified" or
    // "rejected" from a phone — and both screens stop polling for good. The
    // examiner page already holds an examiner-scope grant (bankGrant), so
    // requiring it costs nothing. Refused = nothing dropped, nothing patched.
    if (!(await verifyGrant(url.searchParams.get('grant'), ['examiner']))) {
      return jsonResponse(request, GRANT_INVALID, 403);
    }
    if (!SESSION_RE.test(session)) return badRequest(request, 'קוד סשן לא תקין');

    if (status && !PATCH_STATUSES[status]) return badRequest(request, 'סטטוס לא תקין');
    if (rawId && !rawId.replace(/[^0-9]/g, '')) return badRequest(request, 'מזהה לא תקין');

    if (rawId && status) {
      // Only what came through validation is written — a nudge that carries a
      // broken `audio` still delivers its status rather than failing whole.
      const fields = { status: status };
      const extra = patchMinutes(param(url, 'extraMinutes'), 0);
      const exam = patchMinutes(param(url, 'examMinutes'), 1);
      const audio = param(url, 'audio');
      if (extra !== null) fields.extraMinutes = extra;
      if (exam !== null) fields.examMinutes = exam;
      if (audio === 'on' || audio === 'off') fields.audio = audio;
      // A patch is a WRITE. It must not spend mayForceReread's budget, which
      // pays for upstream READS — this nudge is the one case that saves one.
      if (await patchSnapshot(session, normalizeId(rawId), fields)) {
        return jsonResponse(request, { status: 'ok', patched: true });
      }
    }

    // No patch was asked for, or there was no row to write into (the examinee
    // registered after the snapshot was taken). Drop it: the next poll re-reads
    // upstream and finds the row itself. Whether the drop actually happened is
    // the throttle's business, not the caller's.
    if (mayForceReread(session)) await dropSnapshot(session);
    return jsonResponse(request,
      (rawId || status) ? { status: 'ok', patched: false } : { status: 'ok' });
  }

  /** The answer as an ordinary poll computes it — today's path, within the budget. */
  async function firstLook(session, kind, idNumber, tokenHex, until, ctx, trace, maxAgeMs) {
    const loaded = await loadSnapshot(session, false, until, ctx, trace, maxAgeMs);
    if (!loaded.ok) return UNAVAILABLE_VIEW;
    const view = evaluate(kind, loaded.snapshot, idNumber, tokenHex, loaded.stale);
    // A row the examinee just created is missing from a snapshot taken before
    // it existed — read once more before telling them they are not registered,
    // or that the registration they are polling for was rejected or reset.
    // The second read shares the same deadline: two bounded reads, never two
    // unbounded ones.
    if (view.provisional && loaded.fromCache && mayForceReread(session)) {
      const refreshed = await loadSnapshot(session, true, until, ctx, trace, maxAgeMs);
      if (refreshed.ok) return evaluate(kind, refreshed.snapshot, idNumber, tokenHex, refreshed.stale);
    }
    return view;
  }

  /**
   * What a HELD request looks at, once per iteration. This is the read ORDER,
   * which is identical for both holds; `view(snapshot, stale)` is the only
   * thing that differs between the examinee's poll and the examiner's watch.
   *
   * Memory is free, so it goes first; `caches.default` is read EVERY
   * iteration, because the decision may have been patched in by another
   * isolate and delivering it within the second is the entire point of
   * holding; upstream only when both are past FRESH_MS, and then through
   * fetchCoalesced — forty held requests still cost one Apps Script execution
   * per 2 s, exactly like forty ordinary polls.
   *
   * The first view that actually DIFFERS from what the client holds wins, so a
   * memory copy of ours can never hide a newer decision sitting in the cache.
   * It never spends mayForceReread's budget: that one pays for the examiner's
   * nudge, and a held request re-reads on the freshness clock anyway.
   */
  async function lookAtNewest(session, clientFp, view, trace) {
    // A held request always carries an `fp`, so its window is the safety
    // cadence: a tick costs a cache read (the cross-isolate patch channel),
    // never an Apps Script execution, until the copy crosses HELD_REREAD_MS.
    const maxAgeMs = freshnessFor(clientFp);
    const mine = memoryRead(session, maxAgeMs);
    if (mine) {
      const fromMemory = await view(mine, false);
      trace.src = 'memory';
      trace.age = snapshotAge(mine);
      if (fromMemory.fp !== clientFp) return fromMemory;
      // The cheap question first (r32, TODO 0א.6): is the cached copy even a
      // DIFFERENT snapshot? Its fingerprint travels in a header, so a quiet
      // tick costs a `match` and a string compare — not `JSON.parse` of 5-40
      // KB, every second, for a copy that turns out to be the same one. A
      // `null` here means the header could not answer (a copy written before
      // r32, or one past the window), and then the body is read exactly as
      // r31 read it. Only a DIFFERING fingerprint opens the body.
      const header = await cacheFingerprint('snap', session, maxAgeMs);
      if (header && mine.sfp && header.sfp === mine.sfp) return fromMemory;
      const cached = await cacheRead('snap', session, maxAgeMs);
      if (cached) {
        const patched = await view(cached, false);
        if (patched.fp !== clientFp) { trace.src = 'cache'; trace.age = snapshotAge(cached); return patched; }
      }
      return fromMemory;
    }
    const cached = adopt(session, await cacheRead('snap', session, maxAgeMs));
    if (cached) { trace.src = 'cache'; trace.age = snapshotAge(cached); return view(cached, false); }
    // The hold's own grace timer is what bounds this join of the coalesced
    // read (see `within`), so this await is never a bare one.
    const fetched = await fetchCoalesced(session);
    if (fetched) { trace.src = 'fetch'; trace.age = snapshotAge(fetched); return view(fetched, false); }
    const old = memoryRead(session, STALE_MS) || adopt(session, await cacheRead('stale', session));
    if (old) { trace.src = 'stale'; trace.age = snapshotAge(old); return view(old, true); }
    trace.src = 'none';
    trace.age = -1;
    return UNAVAILABLE_VIEW;
  }

  /** One iteration of an examinee's held poll. */
  function holdLook(session, kind, idNumber, tokenHex, clientFp, trace) {
    return lookAtNewest(session, clientFp,
      (snapshot, stale) => evaluate(kind, snapshot, idNumber, tokenHex, stale), trace);
  }

  /** One iteration of an examiner's held watch — same order, session view. */
  function holdLookSession(session, clientFp, trace) {
    return lookAtNewest(session, clientFp, sessionView, trace);
  }

  /**
   * THE hold — one implementation for both routes (§13.2 asks for exactly
   * that, not a copy). Park until something wakes this session, look again,
   * and answer the moment the view stops being the one the client already has
   * — or when `wait` runs out, which answers that same view with its same
   * fingerprint and lets the client simply ask again.
   *
   * `look(clientFp)` is the route's own reader and must return a promise of a
   * view `{ answer, fp, holdable }`. Returns `{ view, held }`; a view that is
   * not holdable, an empty `clientFp` or a `wait` of 0 answers at once with
   * held = 0, which is what a client that must pace itself has to be told.
   *
   * `startedAt`/`deadline` belong to the REQUEST, not to the hold: the first
   * look has already spent part of that budget (22/09 — before this, a slow
   * first look and a full hold added up to 45 s of wall time). What this
   * guarantees is `first look + hold ≤ wait + HOLD_GRACE_MS`, always, and
   * `held` is measured from the request's own start so that it is the number
   * the client sees on its stopwatch.
   */
  async function holdUntilChanged({ session, view, clientFp, waitMs, startedAt, deadline, ctx, look }) {
    if (!(waitMs && clientFp && view.holdable && view.fp === clientFp)) return { view: view, held: 0 };
    for (let step = 0; step < HOLD_MAX_STEPS && clock() < deadline; step++) {
      await waitForChange(session, Math.min(HOLD_TICK_MS, deadline - clock()));
      // The look may open an upstream read, so it races what is left of the
      // hold plus HOLD_GRACE_MS. A read started a second before the deadline
      // and answered 10 s later would otherwise hold the request past the
      // client's own 40 s abort: a slow Google must never turn a held poll
      // into a client-side communication failure.
      const pending = look(clientFp);
      const next = await within(Math.max(0, deadline - clock()) + HOLD_GRACE_MS, pending);
      if (next === HOLD_TIMED_OUT) {
        // Abandoned, not cancelled. The read is coalesced, so letting it
        // finish lands the snapshot in memory and in caches.default for the
        // other holders and for this client's next request; without waitUntil
        // the runtime may kill it together with this response and the next
        // one pays for the same read again. The answer is the view we already
        // had - same fp, so the client just asks again.
        abandon(pending, ctx);
        break;
      }
      view = next;
      if (!view.holdable || view.fp !== clientFp) break;
    }
    return { view: view, held: Math.max(0, clock() - startedAt) };
  }

  /**
   * The examiner dashboard's long poll (§13.2). It never answers WHAT changed,
   * only THAT something did: the dashboard then calls `examinerDashboard`
   * once, and that single Apps Script execution is the whole cost of the
   * screen. The watch itself costs Google nothing extra — the snapshot it
   * fingerprints is the one the examinees' polls already keep fresh.
   *
   * An examiner grant is required: `rows`/`at` describe a session's state, and
   * more to the point an open door here would let anyone holding a session
   * code park a request and buy one upstream read per 2 s with it.
   */
  async function watch(request, url, ctx) {
    const startedAt = clock();
    if (!(await verifyGrant(url.searchParams.get('grant'), ['examiner']))) {
      return jsonResponse(request, GRANT_INVALID, 403);
    }
    const session = param(url, 'sessionCode');
    if (!SESSION_RE.test(session)) return badRequest(request, 'קוד סשן לא תקין');

    const clientFp = param(url, 'fp');
    const waitMs = waitMillis(url.searchParams.get('wait'));
    // ONE budget for the whole request — the first look and the hold share it.
    const deadline = startedAt + (waitMs || FIRST_LOOK_MAX_MS);
    const trace = { src: 'none', late: 0, age: -1 };

    // No `provisional` re-read here, unlike firstLook: an empty session is not
    // a mistake to correct, it is the normal state before the first examinee
    // registers — and `s:none` is held until that registration wakes it.
    const loaded = await loadSnapshot(session, false, deadline, ctx, trace, freshnessFor(clientFp));
    const first = loaded.ok ? await sessionView(loaded.snapshot, loaded.stale) : UNAVAILABLE_VIEW;
    const firstLookMs = Math.max(0, clock() - startedAt);
    const held = await holdUntilChanged({
      session: session, view: first, clientFp: clientFp, waitMs: waitMs,
      startedAt: startedAt, deadline: deadline, ctx: ctx,
      look: fp => holdLookSession(session, fp, trace)
    });
    trace.fl = firstLookMs;
    // The data itself, when there is a reason to send it (r32, §14.1). Built
    // ONCE, here, after the hold has picked the view that will be answered.
    const payload = sessionPayload(held.view, clientFp);
    // `nr` is in the line whether or not the payload went out: an `nr` stuck at
    // 0 in a session that HAS finished exams is how a tail says "the server is
    // still r31" (§14.5's verification), and `sess` is how it says "the
    // dashboard was painted from here".
    const line = { rows: held.view.answer.rows, nr: sessionResults(held.view).length };
    if (payload) line.sess = 1;
    logAnswer('watch', session, trace, held.view.fp, held.held, line);
    // `status` leads, then the two long-poll fields, then the view's own body
    // (whose `status` re-states the same value and keeps that first place),
    // and `session` last — the dashboard reads the small fields either way.
    const answer = Object.assign({ status: 'ok', fp: held.view.fp, held: held.held | 0 }, held.view.answer);
    if (payload) answer.session = payload;
    return jsonResponse(request, answer);
  }

  /**
   * One line per answered watch/poll, so the next `wrangler tail` explains
   * itself: WHERE the answered view came from, which fingerprint it carries,
   * how long the first look took, and what the last upstream read of that
   * session cost and produced. That is the whole diagnosis of a fingerprint
   * that moves without a write — `fp` different from `usfp` means a copy of
   * another shape is alive beside the one Google last gave us.
   */
  function logAnswer(route, session, trace, fp, held, extra) {
    const last = lastUpstream.get(session) || {};
    emit(JSON.stringify(Object.assign({
      r: route, s: session, src: trace.src, fp: fp, held: held | 0,
      fl: trace.fl | 0, age: trace.age | 0, late: trace.late | 0,
      up: last.ms | 0, usfp: last.sfp || ''
    }, extra)));
  }

  async function poll(request, url, ctx) {
    const startedAt = clock();
    const kind = url.searchParams.get('kind') || '';
    const session = String(url.searchParams.get('sessionCode') || '').trim();
    const rawId = url.searchParams.get('idNumber') || '';
    const token = String(url.searchParams.get('examineeToken') || '').trim();
    // A client that speaks long polling names itself by sending either field,
    // and only it is served `fp`/`held`. For everyone else the answer stays
    // byte-identical to the server's own checkApproval/getExamStatus, which is
    // the invariant tests/contracts.test.cjs pins — and the reason the client
    // can swap a URL rather than logic.
    const longPoll = url.searchParams.has('wait') || url.searchParams.has('fp');
    const clientFp = param(url, 'fp');
    const waitMs = waitMillis(url.searchParams.get('wait'));

    /** Every poll answer leaves through here: CORS, no-store, HTTP 200 as today. */
    const reply = (body, fp, held, status) => jsonResponse(request,
      longPoll ? Object.assign({}, body, { fp: fp, held: held | 0 }) : body, status);

    if (kind !== 'approval' && kind !== 'status') {
      return reply({ status: 'error', message: 'kind must be approval or status' }, 'x:kind', 0, 400);
    }
    if (!SESSION_RE.test(session)) {
      return reply({ status: 'error', message: 'קוד סשן לא תקין' }, 'x:sess', 0, 400);
    }
    if (!String(rawId).replace(/[^0-9]/g, '')) {
      return reply({ status: 'error', message: 'חסר מזהה' }, 'x:id', 0, 400);
    }

    const idNumber = normalizeId(rawId);
    const tokenHex = token ? await sha256Hex(token) : '';
    // ONE budget for the whole request — the first look and the hold share it,
    // and a request that sends no `wait` still gets a bound on its first look.
    const deadline = startedAt + (waitMs || FIRST_LOOK_MAX_MS);
    const trace = { src: 'none', late: 0, age: -1 };
    const first = await firstLook(session, kind, idNumber, tokenHex, deadline, ctx, trace, freshnessFor(clientFp));
    const firstLookMs = Math.max(0, clock() - startedAt);

    // The hold. Only while the answer is EXACTLY the one the client already
    // has: a changed answer, a stale copy, a dead upstream and a token error
    // all return at once. Each iteration re-evaluates and either answers or
    // parks again, until the deadline or the step guard.
    const held = await holdUntilChanged({
      session: session, view: first, clientFp: clientFp, waitMs: waitMs,
      startedAt: startedAt, deadline: deadline, ctx: ctx,
      look: fp => holdLook(session, kind, idNumber, tokenHex, fp, trace)
    });
    trace.fl = firstLookMs;
    logAnswer('poll', session, trace, held.view.fp, held.held, { k: kind });
    return reply(held.view.answer, held.view.fp, held.held);
  }

  // `ctx` is the Workers execution context and may be absent (a test, a direct
  // call): only the hold uses it, and only to let an abandoned look finish.
  const handle = async function handle(request, ctx) {
    if (request.method === 'OPTIONS') {
      return new Response(null, { status: 204, headers: corsHeaders(request) });
    }
    const url = new URL(request.url);
    if (request.method === 'POST') {
      if (url.pathname === '/v1/invalidate') return invalidate(request, url);
      return jsonResponse(request, { status: 'error', message: 'method not allowed' }, 405);
    }
    if (request.method !== 'GET') {
      return jsonResponse(request, { status: 'error', message: 'method not allowed' }, 405);
    }
    if (url.pathname === '/' || url.pathname === '') {
      // `bankBuild()` is a promise shared by the whole isolate, so this join
      // gets its own timer like every other one (see `within`) — the watchdog
      // must never be told the Worker is down because the runtime cancelled a
      // health check that was waiting on another request's asset read.
      const build = await within(5000, bankBuild());
      return jsonResponse(request, {
        status: 'ok', service: 'session-gateway', build: BUILD,
        bank: build === HOLD_TIMED_OUT ? '' : build
      });
    }
    if (url.pathname === '/v1/poll') return poll(request, url, ctx);
    if (url.pathname === '/v1/session/watch') return watch(request, url, ctx);
    if (url.pathname === '/v1/bank') return bank(request, url);
    if (url.pathname === '/v1/bank/full') return bankFull(request, url);
    return jsonResponse(request, { status: 'error', message: 'not found' }, 404);
  };

  // A hook, not a route — nothing from the internet can reach it. It exists so
  // the bookkeeping can be asserted directly: after every request has answered,
  // `waiting` must be 0, or a held request leaked a resolver into the isolate.
  handle._debug = () => ({
    sessions: waiters.size,
    waiting: [...waiters.values()].reduce((total, set) => total + set.size, 0),
    memory: memory.size,
    inflight: inflight.size
  });
  return handle;
}

// One gateway per isolate: the in-flight map and the memory snapshots must
// survive between requests or nothing coalesces. `env` is stable for the life
// of an isolate, so capturing it from the first request is safe.
let singleton = null;

export default {
  fetch(request, env, ctx) {
    if (!singleton) {
      singleton = createGateway({
        fetch: (url, init) => globalThis.fetch(url, init),
        caches: globalThis.caches,
        now: () => Date.now(),
        env: env
      });
    }
    return singleton(request, ctx);
  }
};
