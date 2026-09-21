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
 *   GET  /v1/bank?grant=…[&ids=1,2&langs=he,en]   — texts for the granted ids
 *   GET  /v1/bank/full?grant=…&lang=he            — a whole language (examiner)
 *   POST /v1/invalidate?grant=<examiner>&sessionCode=X[&idNumber&status&…]
 *                         — drop the session snapshot, or patch the decision
 *                           straight into it (see `invalidate`)
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
 * Failure policy: a snapshot up to 60 s old is served with `stale:true` rather
 * than an error; with nothing cached the answer is a RETRYABLE error (HTTP
 * 200), never Google's HTML. The client slows down on it but must NOT fall back
 * to direct polling for it, or the storm returns. None of those are ever held:
 * a client that must pace itself has to be told so at once.
 */

const BUILD = '2026-09-21';

const ALLOWED_ORIGINS = [
  'https://bohanyzahal-cyber.github.io',
  'http://localhost',
  'http://127.0.0.1'
];

const FRESH_MS = 2000;            // a snapshot this young answers without asking
const STALE_MS = 60000;           // older than this and we would rather error
const REREAD_GAP_MS = 2000;       // forced upstream re-read: once per session per gap
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

/** null = no active row for this id (the caller decides what that means). */
function approvalAnswer(rows, idNumber, tokenHex) {
  for (let i = rows.length - 1; i >= 0; i--) {
    const row = rows[i];
    if (normalizeId(row.id) !== idNumber) continue;
    const approval = String(row.status || 'waiting').trim();
    if (TERMINAL_APPROVALS[approval]) continue;
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
  return null;
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
  fp: 'x:up', missing: false, holdable: false
};

// --- the gateway -----------------------------------------------------------

/**
 * Dependency-injected so tests can drive it with a fake clock, a counting
 * fetch, an in-memory ASSETS binding and a fake `sleep` (a held request must be
 * testable without waiting 25 real seconds). The state (in-flight map, memory
 * snapshots, waiters, imported HMAC key) lives in the closure, so the
 * module-level singleton below coalesces across requests of one isolate while
 * each test gets its own clean instance.
 */
export function createGateway({ fetch, caches, now, env, sleep }) {
  const clock = now || (() => Date.now());
  const nap = sleep || (ms => new Promise(resolve => setTimeout(resolve, ms)));
  const memory = new Map();   // session -> { at, snapshot }
  const inflight = new Map(); // session -> Promise<snapshot|null>
  const lastForced = new Map(); // session -> ms of the last forced re-read
  const waiters = new Map();  // session -> Set<resolve> — the held requests

  const cacheKey = (kind, session) =>
    'https://session-gateway.internal/' + kind + '/' + encodeURIComponent(session);

  async function cacheRead(kind, session) {
    if (!caches || !caches.default) return null;
    try {
      const hit = await caches.default.match(new Request(cacheKey(kind, session)));
      if (!hit) return null;
      const body = await hit.json();
      return body && Array.isArray(body.rows) ? body : null;
    } catch (e) {
      return null;
    }
  }

  async function cacheWrite(session, snapshot) {
    if (!caches || !caches.default) return;
    const body = JSON.stringify(snapshot);
    const store = (kind, maxAge) => caches.default.put(
      new Request(cacheKey(kind, session)),
      new Response(body, {
        headers: { 'Content-Type': 'application/json', 'Cache-Control': 'max-age=' + maxAge }
      })
    );
    try { await Promise.all([store('snap', FRESH_MS / 1000), store('stale', STALE_MS / 1000)]); } catch (e) { /* cache is best effort */ }
  }

  function memoryRead(session, maxAgeMs) {
    const hit = memory.get(session);
    if (!hit || clock() - hit.at > maxAgeMs) return null;
    return hit.snapshot;
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
    return Promise.race([woken, nap(ms)]).then(() => {
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
      return { at: Number(data.at) || clock(), rows: data.rows };
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
    // The cache write is awaited, not floating: a Worker may be torn down as
    // soon as it answers, and a dropped write means the next isolate asks
    // Apps Script again — exactly what this gateway exists to prevent.
    const promise = fetchUpstream(session).then(async snapshot => {
      inflight.delete(session);
      if (snapshot) {
        memory.set(session, { at: clock(), snapshot });
        await cacheWrite(session, snapshot);
        wake(session); // fresh truth: whoever is held on this session re-reads it
      }
      return snapshot;
    });
    inflight.set(session, promise);
    return promise;
  }

  /** { ok, snapshot, fromCache, stale } — `ok:false` only when nothing exists. */
  async function loadSnapshot(session, forceFresh) {
    if (!forceFresh) {
      const fresh = memoryRead(session, FRESH_MS) || await cacheRead('snap', session);
      if (fresh) return { ok: true, snapshot: fresh, fromCache: true, stale: false };
    }
    const fetched = await fetchCoalesced(session);
    if (fetched) return { ok: true, snapshot: fetched, fromCache: false, stale: false };
    const old = memoryRead(session, STALE_MS) || await cacheRead('stale', session);
    if (old) return { ok: true, snapshot: old, fromCache: true, stale: true };
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
    const snapshot = { at: current.at, rows: rows };       // `at` stays the read time
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

  // Imported once per isolate: importKey on every request would be pure waste
  // on the hot path. An empty secret is kept out of importKey, which rejects it.
  let keyPromise = null;
  function hmacKey() {
    if (!keyPromise) {
      keyPromise = crypto.subtle.importKey(
        'raw', new TextEncoder().encode(String(env.GATEWAY_KEY || '')),
        { name: 'HMAC', hash: 'SHA-256' }, false, ['verify']
      ).catch(() => null);
    }
    return keyPromise;
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
    // An examiner-only door. A nudge writes a status straight into what the
    // examinees are answered from: without this check anyone who knows a
    // session code and a classmate's id could show them "disqualified" or
    // "rejected" from a phone — and both screens stop polling for good. The
    // examiner page already holds an examiner-scope grant (bankGrant), so
    // requiring it costs nothing. Refused = nothing dropped, nothing patched.
    if (!(await verifyGrant(url.searchParams.get('grant'), ['examiner']))) {
      return jsonResponse(request, GRANT_INVALID, 403);
    }
    const session = param(url, 'sessionCode');
    if (!SESSION_RE.test(session)) return badRequest(request, 'קוד סשן לא תקין');

    const rawId = param(url, 'idNumber');
    const status = param(url, 'status');
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

  /** The answer as an ordinary poll computes it — today's path, unchanged. */
  async function firstLook(session, kind, idNumber, tokenHex) {
    const loaded = await loadSnapshot(session, false);
    if (!loaded.ok) return UNAVAILABLE_VIEW;
    const view = evaluate(kind, loaded.snapshot, idNumber, tokenHex, loaded.stale);
    // A row the examinee just created is missing from a snapshot taken before
    // it existed — read once more before telling them they are not registered.
    if (view.missing && loaded.fromCache && mayForceReread(session)) {
      const refreshed = await loadSnapshot(session, true);
      if (refreshed.ok) return evaluate(kind, refreshed.snapshot, idNumber, tokenHex, refreshed.stale);
    }
    return view;
  }

  /**
   * What a HELD request looks at, once per iteration. Memory is free, so it
   * goes first; `caches.default` is read EVERY iteration, because the decision
   * may have been patched in by another isolate and delivering it within the
   * second is the entire point of holding; upstream only when both are past
   * FRESH_MS, and then through fetchCoalesced — forty held requests still cost
   * one Apps Script execution per 2 s, exactly like forty ordinary polls.
   *
   * The first view that actually DIFFERS from what the client holds wins, so a
   * memory copy of ours can never hide a newer decision sitting in the cache.
   * It never spends mayForceReread's budget: that one pays for the examiner's
   * nudge, and a held request re-reads on the freshness clock anyway.
   */
  async function holdLook(session, kind, idNumber, tokenHex, clientFp) {
    const look = (snapshot, stale) => evaluate(kind, snapshot, idNumber, tokenHex, stale);
    const mine = memoryRead(session, FRESH_MS);
    if (mine) {
      const view = look(mine, false);
      if (view.fp !== clientFp) return view;
      const cached = await cacheRead('snap', session);
      if (cached) {
        const patched = look(cached, false);
        if (patched.fp !== clientFp) return patched;
      }
      return view;
    }
    const cached = await cacheRead('snap', session);
    if (cached) return look(cached, false);
    const fetched = await fetchCoalesced(session);
    if (fetched) return look(fetched, false);
    const old = memoryRead(session, STALE_MS) || await cacheRead('stale', session);
    return old ? look(old, true) : UNAVAILABLE_VIEW;
  }

  async function poll(request, url) {
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
    let view = await firstLook(session, kind, idNumber, tokenHex);

    // The hold. Only while the answer is EXACTLY the one the client already
    // has: a changed answer, a stale copy, a dead upstream and a token error
    // all return at once. Each iteration re-evaluates and either answers or
    // parks again, until the deadline or the step guard.
    let held = 0;
    if (waitMs && clientFp && view.holdable && view.fp === clientFp) {
      const start = clock();
      const deadline = start + waitMs;
      for (let step = 0; step < HOLD_MAX_STEPS && clock() < deadline; step++) {
        await waitForChange(session, Math.min(HOLD_TICK_MS, deadline - clock()));
        view = await holdLook(session, kind, idNumber, tokenHex, clientFp);
        if (!view.holdable || view.fp !== clientFp) break;
      }
      held = Math.max(0, clock() - start);
    }
    return reply(view.answer, view.fp, held);
  }

  const handle = async function handle(request) {
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
      return jsonResponse(request, {
        status: 'ok', service: 'session-gateway', build: BUILD, bank: await bankBuild()
      });
    }
    if (url.pathname === '/v1/poll') return poll(request, url);
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
  fetch(request, env) {
    if (!singleton) {
      singleton = createGateway({
        fetch: (url, init) => globalThis.fetch(url, init),
        caches: globalThis.caches,
        now: () => Date.now(),
        env: env
      });
    }
    return singleton(request);
  }
};
