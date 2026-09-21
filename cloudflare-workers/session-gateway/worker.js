/**
 * session-gateway — one upstream call per session per 3 seconds, however many
 * examinees are polling. ES module worker; deploy with `npx wrangler deploy`.
 *
 * WHY: every approval/status poll used to be its own Apps Script execution
 * (1.5-2.5 s of cold start, one of ~30 slots). 40 waiting examinees = ~480
 * executions per minute for data that is identical for all of them. This
 * gateway asks the server once per session (`action=sessionSnapshot`) and
 * answers every examinee from that snapshot — the poll cost stops scaling with
 * the number of examinees. (DESIGN_2026-09-21 §3.4.)
 *
 * The answers are byte-compatible with the server's own checkApproval /
 * getExamStatus so the client only swaps the URL, never the logic. The
 * examinee token never leaves Apps Script: the snapshot carries a SHA-256 hash
 * and the gateway hashes what the client sent to compare.
 *
 * Endpoints:
 *   GET /                 — health: {status:'ok', service, build}
 *   GET /v1/poll?kind=approval|status&sessionCode&idNumber&examineeToken
 *   OPTIONS *             — CORS preflight
 *
 * Failure policy: a snapshot up to 60 s old is served with `stale:true` rather
 * than an error; with nothing cached the answer is a RETRYABLE error (HTTP
 * 200), never Google's HTML. The client slows down on it but must NOT fall back
 * to direct polling for it, or the storm returns.
 */

const BUILD = '2026-09-21';

const ALLOWED_ORIGINS = [
  'https://bohanyzahal-cyber.github.io',
  'http://localhost',
  'http://127.0.0.1'
];

const FRESH_MS = 3000;            // a snapshot this young answers without asking
const STALE_MS = 60000;           // older than this and we would rather error
const FORCE_GAP_MS = 10000;       // re-read-on-miss, at most once per session
const UPSTREAM_TIMEOUT_MS = 25000;
const SESSION_RE = /^[A-Z0-9]{6,8}$/;

// handleCheckApproval skips these and keeps looking for an active row; see the
// long comment there about the shared-ID incident that put 'rejected' on it.
const TERMINAL_APPROVALS = { completed: 1, disqualified: 1, cancelled: 1, rejected: 1 };

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

function corsHeaders(request) {
  const origin = (request.headers && request.headers.get('Origin')) || '';
  const allowed = ALLOWED_ORIGINS.find(a => origin === a || origin.indexOf(a + ':') === 0);
  return {
    'Access-Control-Allow-Origin': allowed ? origin : ALLOWED_ORIGINS[0],
    'Access-Control-Allow-Methods': 'GET, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type',
    'Access-Control-Max-Age': '86400',
    'Vary': 'Origin'
  };
}

function jsonResponse(request, body, status) {
  return new Response(JSON.stringify(body), {
    status: status || 200,
    headers: Object.assign({
      'Content-Type': 'application/json; charset=utf-8',
      'Cache-Control': 'no-store'
    }, corsHeaders(request))
  });
}

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

// --- the gateway -----------------------------------------------------------

/**
 * Dependency-injected so tests can drive it with a fake clock and a counting
 * fetch. The state (in-flight map, memory snapshots) lives in the closure, so
 * the module-level singleton below coalesces across requests of one isolate
 * while each test gets its own clean instance.
 */
export function createGateway({ fetch, caches, now, env }) {
  const clock = now || (() => Date.now());
  const memory = new Map();   // session -> { at, snapshot }
  const inflight = new Map(); // session -> Promise<snapshot|null>
  const lastForced = new Map(); // session -> ms of the last re-read-on-miss

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

  /** Mirrors the server's own re-read before it says "no registration". */
  function mayForceReread(session) {
    const last = lastForced.get(session);
    if (last != null && clock() - last < FORCE_GAP_MS) return false;
    lastForced.set(session, clock());
    return true;
  }

  async function poll(request, url) {
    const kind = url.searchParams.get('kind') || '';
    const session = String(url.searchParams.get('sessionCode') || '').trim();
    const rawId = url.searchParams.get('idNumber') || '';
    const token = String(url.searchParams.get('examineeToken') || '').trim();

    if (kind !== 'approval' && kind !== 'status') {
      return jsonResponse(request, { status: 'error', message: 'kind must be approval or status' }, 400);
    }
    if (!SESSION_RE.test(session)) {
      return jsonResponse(request, { status: 'error', message: 'קוד סשן לא תקין' }, 400);
    }
    if (!String(rawId).replace(/[^0-9]/g, '')) {
      return jsonResponse(request, { status: 'error', message: 'חסר מזהה' }, 400);
    }

    const idNumber = normalizeId(rawId);
    const tokenHex = token ? await sha256Hex(token) : '';
    const compute = rows => (kind === 'approval'
      ? approvalAnswer(rows, idNumber, tokenHex)
      : statusAnswer(rows, idNumber, tokenHex));

    let loaded = await loadSnapshot(session, false);
    if (!loaded.ok) {
      return jsonResponse(request, {
        status: 'error', code: 'upstream_unavailable', retryable: true,
        message: 'השרת עמוס — ננסה שוב אוטומטית'
      });
    }
    let answer = compute(loaded.snapshot.rows);
    // A row the examinee just created is missing from a snapshot taken before
    // it existed — read once more before telling them they are not registered.
    if (answer === null && loaded.fromCache && mayForceReread(session)) {
      const refreshed = await loadSnapshot(session, true);
      if (refreshed.ok) {
        loaded = refreshed;
        answer = compute(refreshed.snapshot.rows);
      }
    }
    if (answer === null) answer = NOT_FOUND[kind];
    if (loaded.stale) answer = Object.assign({}, answer, { stale: true });
    return jsonResponse(request, answer);
  }

  return async function handle(request) {
    if (request.method === 'OPTIONS') {
      return new Response(null, { status: 204, headers: corsHeaders(request) });
    }
    if (request.method !== 'GET') {
      return jsonResponse(request, { status: 'error', message: 'method not allowed' }, 405);
    }
    const url = new URL(request.url);
    if (url.pathname === '/' || url.pathname === '') {
      return jsonResponse(request, { status: 'ok', service: 'session-gateway', build: BUILD });
    }
    if (url.pathname === '/v1/poll') return poll(request, url);
    return jsonResponse(request, { status: 'error', message: 'not found' }, 404);
  };
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
