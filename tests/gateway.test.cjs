// session-gateway regression gates. No network: the upstream fetch is faked and
// counted, the clock is faked. What is proven here is the reason the gateway
// exists — 40 examinees polling one session must cost ONE Apps Script run.
const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');
const crypto = require('node:crypto');
const { pathToFileURL } = require('node:url');

const WORKER_URL = pathToFileURL(
  path.join(__dirname, '..', 'cloudflare-workers', 'session-gateway', 'worker.js')
).href;

let createGateway;
test.before(async () => { ({ createGateway } = await import(WORKER_URL)); });

const ENV = { API_URL: 'https://script.example/macros/s/AKtest/exec', GATEWAY_KEY: 'test-secret' };
const SESSION = 'ABC12345';
const PAGES_ORIGIN = 'https://bohanyzahal-cyber.github.io';
const CLOCK0 = 1000000;
const sha256Hex = text => crypto.createHash('sha256').update(text).digest('hex');
const row = over => Object.assign(
  { id: '900000001', status: 'waiting', tokenHash: '', audio: 'off', examMinutes: 40, extraMinutes: 0 }, over);

// --- the private bank (Workers Static Assets) ------------------------------

const BANK_BUILD = 'bank0001deadbeef';
const Q14 = '{"id":14,"l":{"he":{"t":"שאלה 14","a":["א","ב"],"i":""},"ru":{"t":"RU 14","a":["A","B"],"i":""},' +
  '"en":{"t":"EN 14","a":["A","B"],"i":"TQ_PIC_14.jpg"}}}';
const Q120 = '{"id":120,"l":{"he":{"t":"שאלה 120","a":["א","ב"],"i":""},"en":{"t":"EN 120","a":["A","B"],"i":""}}}';
const HE_BANK = '[{"id":14,"t":"שאלה 14","a":["א","ב"],"i":""},{"id":120,"t":"שאלה 120","a":["א","ב"],"i":""}]';
const ASSETS = {
  '/manifest.json': JSON.stringify({ build: BANK_BUILD, questions: 2, langs: {} }),
  '/q/14.json': Q14,
  '/q/120.json': Q120,
  '/bank/he.json': HE_BANK
};

/** An in-memory Workers Static Assets binding: published path -> file text. */
function assetsBinding(files) {
  return {
    fetch: async reqOrUrl => {
      const url = new URL(typeof reqOrUrl === 'string' ? reqOrUrl : reqOrUrl.url);
      const body = files[url.pathname];
      if (body === undefined) return new Response('not found', { status: 404 });
      return new Response(body, { status: 200, headers: { 'Content-Type': 'application/json' } });
    }
  };
}

/** base64url, exactly what Utilities.base64EncodeWebSafe produces server-side. */
const b64url = buf => Buffer.from(buf).toString('base64url');

/** Signs a grant the way Apps Script does — the Worker must never sign one. */
function grant(payload, key) {
  const claims = Object.assign({ v: 1, exp: CLOCK0 + 3600000 }, payload);
  const body = b64url(Buffer.from(JSON.stringify(claims), 'utf8'));
  const sig = b64url(crypto.createHmac('sha256', key || ENV.GATEWAY_KEY).update(body).digest());
  return body + '.' + sig;
}
const examGrant = (ids, over) => grant(Object.assign({ s: 'exam', ids, sub: SESSION + ':012345678' }, over));
const examinerGrant = over => grant(Object.assign({ s: 'examiner', sub: 'ex:7' }, over));

/**
 * What a cached Response answers to the gateway. Since r32 that is TWO things:
 * `headers.get('X-SFP'/'X-RAT')`, which is all a quiet tick of a held request
 * may read, and `json()`, which parses 5-40 KB and is therefore COUNTED — the
 * CPU gate of TODO 0א.6. `get` is case-insensitive like the real Headers, and
 * an entry written before r32 (a test that seeds `store` by hand) has no
 * headers at all and answers null, which is the compatibility path.
 */
const cachedResponse = (text, headers, counts) => ({
  status: 200,
  headers: {
    get(name) {
      const wanted = String(name).toLowerCase();
      for (const key of Object.keys(headers || {})) {
        if (key.toLowerCase() === wanted) return headers[key];
      }
      return null;
    }
  },
  text: async () => text,
  json: async () => { counts.json++; return JSON.parse(text); }
});

/**
 * `caches.default` in memory: max-age is honoured against the same fake clock
 * the gateway reads, so an expiring `snap` copy behaves as it does at the edge.
 * `counts` is what the CPU tests assert on — `match` is cheap, `json` is not.
 */
function cacheDouble(now) {
  const store = new Map();
  const counts = { match: 0, json: 0 };
  return { store, counts, caches: { default: {
    async match(request) {
      counts.match++;
      const hit = store.get(request.url);
      if (!hit || now() - hit.at > hit.maxAge * 1000) return undefined;
      return cachedResponse(hit.body, hit.headers, counts);
    },
    async put(request, response) {
      const maxAge = Number(/max-age=(\d+)/.exec(response.headers.get('Cache-Control') || '')[1]);
      const headers = {};
      for (const name of ['X-SFP', 'X-RAT']) {
        const value = response.headers.get(name);
        if (value != null) headers[name] = value;
      }
      store.set(request.url, { body: await response.text(), maxAge, at: now(), headers });
    },
    async delete(request) { return store.delete(request.url); }
  } } };
}

/**
 * What the gateway reads a body from — `status`, `text()`, `json()` — and
 * NOTHING else. Deliberately not a real `Response`: undici answers `text()`
 * through its stream machinery, which takes an unpredictable number of event
 * loop turns, and those turns sit inside the very chain a held request races
 * its grace timer against. A fake clock cannot be fair to a race it cannot
 * see, so the fake upstream stays on the microtask queue.
 */
const body = (text, status) => ({
  status: status,
  text: async () => text,
  json: async () => JSON.parse(text)
});

/**
 * The fake `sleep` a HELD request parks on. Virtual time moves only while the
 * gateway is asleep: `fire()` jumps the clock to the earliest deadline and
 * resolves everything due at once, so forty held requests wake together exactly
 * as they do at the edge, and a 25 s hold is tested in milliseconds.
 * `calls` counts iterations — the CPU guard a held request must respect.
 */
function timerQueue(state) {
  const parked = [];
  const queue = {
    calls: 0,
    parked,
    sleep(ms) {
      queue.calls++;
      const at = state.clock + Math.max(0, Number(ms) || 0);
      let entry;
      const timer = new Promise(resolve => { entry = { at, resolve }; parked.push(entry); });
      // The gateway drops the loser of every race (`within`, `waitForChange`).
      // A timer left parked would go on steering the fake clock — `fire()`
      // jumps to the earliest deadline — long after nobody is listening.
      timer.cancel = () => {
        const i = parked.indexOf(entry);
        if (i >= 0) parked.splice(i, 1);
      };
      return timer;
    },
    /** The earliest deadline parked, or null. */
    next() {
      return parked.length ? parked.reduce((soonest, t) => Math.min(soonest, t.at), Infinity) : null;
    },
    /** Jumps to that deadline and resolves everything due. false = nothing parked. */
    fire() {
      const at = queue.next();
      if (at == null) return false;
      if (at > state.clock) state.clock = at;
      for (let i = parked.length - 1; i >= 0; i--) {
        if (parked[i].at <= state.clock) parked.splice(i, 1)[0].resolve();
      }
      return true;
    }
  };
  return queue;
}

/**
 * WebCrypto is answered on Node's THREAD POOL, not on the microtask queue a
 * few `setImmediate` rounds drain — and the gateway is full of it: the grant
 * `verify` on every watch and every nudge, the `digest` behind every session
 * fingerprint (r31) and every examinee token. Virtual time must never move
 * while one of those is in flight, or a hold's grace timer beats the very
 * fingerprint it was about to be woken with and the test measures this harness
 * instead of the Worker. Counting them here is the only way to know; the
 * gateway itself is told nothing.
 */
let subtleInFlight = 0;
// ...and how many there were, which is the CPU gate: a held watch that hashes
// the session on every tick costs 7-12 ms of CPU against a free-plan ceiling
// of 10 (22/09). `digest` must scale with the upstream READS, never the ticks.
const subtleCalls = { digest: 0, verify: 0, importKey: 0 };
for (const name of ['digest', 'verify', 'importKey']) {
  const real = globalThis.crypto.subtle[name];
  globalThis.crypto.subtle[name] = function (...args) {
    subtleInFlight++;
    subtleCalls[name]++;
    return real.apply(globalThis.crypto.subtle, args).finally(() => { subtleInFlight--; });
  };
}

/** Lets every ready promise callback run. The fake clock does not move here. */
const flush = async rounds => {
  for (let i = 0; i < (rounds || 12); i++) await new Promise(r => setImmediate(r));
  // ...and never hands control back while WebCrypto is still working, or with
  // the next call about to be made: the caller is either about to assert on
  // what it produces or about to move the clock past it. "Quiet" therefore
  // means quiet for several turns, because one request chains verify -> read
  // -> digest with ordinary async between the links.
  for (let quiet = 0, guard = 0; quiet < 4 && guard < 200; guard++) {
    quiet = subtleInFlight ? 0 : quiet + 1;
    await new Promise(r => setImmediate(r));
  }
};

/**
 * Waits, WITHOUT moving the fake clock, until `n` requests are parked on this
 * gateway — the barrier every "now nudge it" test needs. `flush()` alone is a
 * guess: a request reaches `waitForChange` only after its grant has been
 * verified and its first look computed, and WebCrypto answers on Node's thread
 * pool, so a fixed number of turns is a race. Returns what it found, so the
 * caller asserts on it and a request that ANSWERED instead of parking still
 * fails the test.
 */
async function parked(gateway, n) {
  for (let guard = 0; guard < 300 && gateway._debug().waiting < n; guard++) await flush(2);
  return gateway._debug().waiting;
}

/**
 * Awaits work that HOLDS: drains what is ready, and when everything is parked
 * on a fake sleep, jumps the clock to the next deadline. Nothing here waits on
 * real time.
 */
async function settle(state, promise) {
  let done = false;
  const tracked = Promise.resolve(promise).then(
    value => { done = true; return value; },
    error => { done = true; throw error; });
  tracked.catch(() => { /* re-thrown to the caller by `return tracked` */ });
  for (let guard = 0; guard < 400 && !done; guard++) {
    await flush();
    if (!done) state.timers.fire();
  }
  return tracked;
}

/**
 * `settle`, plus the number of timers THIS request armed. Since 22/09 a first
 * look that has to ask Google arms one of its own — the bound that stops a
 * `wait=25` request from answering after 45 s of wall time (`wrangler tail`) —
 * so the interesting number is always a delta, never the harness total: a HOLD
 * costs two per second on top of it (a tick and a grace), and a request that
 * never holds costs at most that one guard.
 */
async function timed(state, promise) {
  const before = state.timers.calls;
  const answer = await settle(state, promise);
  answer.timers = state.timers.calls - before;
  return answer;
}

/** Runs the fake clock forward, waking every held request on the way. */
async function advance(state, ms) {
  const target = state.clock + ms;
  for (let guard = 0; guard < 400 && state.clock < target; guard++) {
    await flush();
    const next = state.timers.next();
    if (next == null || next > target) break;
    state.timers.fire();
  }
  await flush();
  if (state.clock < target) state.clock = target;
}

/**
 * Fake upstream + fake clock; `state.calls` is the Apps Script execution count.
 * `withCache` adds a shared caches.default and `spawn()`, which is a COLD
 * isolate: new memory, same cache, same clock, same upstream counter.
 */
function harness(snapshots, assets, withCache) {
  // `serverV` is which sessionSnapshot the fake Apps Script speaks: 1 = r31
  // (no `v`, no `results` — the shape the Worker must still hash exactly as
  // r31 did), 2 = r32 (§14.1), which also answers `results[session]`.
  const state = {
    calls: [], clock: CLOCK0, mode: 'ok', snapshots: snapshots || {},
    upstreamDelayMs: 0, serverV: 1, results: {}
  };
  state.timers = timerQueue(state);
  const fetchFn = async url => {
    state.calls.push(String(url));
    // A slow Google, through the same fake clock: `upstreamDelayMs` is what a
    // held request's look can still be waiting on when its deadline passes.
    if (state.upstreamDelayMs) await state.timers.sleep(state.upstreamDelayMs);
    if (state.mode === 'error500') return body('<html>Google internal error</html>', 500);
    if (state.mode === 'notfound') return body('Not Found', 404);
    if (state.mode === 'html200') return body('<!DOCTYPE html><html>Drive error</html>', 200);
    const session = new URL(url).searchParams.get('sessionCode');
    const rows = state.snapshots[session] || [];
    const answer = { status: 'ok', at: state.clock, rows };
    if (state.serverV >= 2) {
      answer.v = 2;
      answer.results = state.results[session] || [];
    }
    return body(JSON.stringify(answer), 200);
  };
  const env = assets ? Object.assign({}, ENV, { ASSETS: assets.fetch ? assets : assetsBinding(assets) }) : ENV;
  const double = withCache ? cacheDouble(() => state.clock) : null;
  state.store = double ? double.store : null;
  state.cacheCounts = double ? double.counts : null;
  // One structured line per answered watch/poll. Collected instead of printed:
  // the suite stays readable, and a test can assert what `wrangler tail` will
  // show (`state.logs.at(-1)`).
  state.logs = [];
  const spawn = () => createGateway({
    fetch: fetchFn, caches: double ? double.caches : undefined, now: () => state.clock, env,
    sleep: ms => state.timers.sleep(ms),
    log: line => state.logs.push(JSON.parse(line))
  });
  return { state, gateway: spawn(), spawn };
}

/** Any route: the raw text matters for the bank, which never re-serialises. */
async function call(gateway, pathAndQuery, init, ctx) {
  const res = await gateway(new Request('https://session-gateway.test' + pathAndQuery, init), ctx);
  const text = await res.text();
  let body = null;
  try { body = JSON.parse(text); } catch (e) { /* the test asserts on text */ }
  return { res, text, body };
}

function pollRequest(params, origin) {
  const url = 'https://session-gateway.test/v1/poll?' + new URLSearchParams(params).toString();
  return new Request(url, origin ? { headers: { Origin: origin } } : undefined);
}
async function poll(gateway, params, origin, ctx) {
  const res = await gateway(pollRequest(params, origin), ctx);
  return { res, body: await res.json() };
}
const approvalPoll = (id, over) => Object.assign({ kind: 'approval', sessionCode: SESSION, idNumber: id }, over);
const statusPoll = (id, over) => Object.assign({ kind: 'status', sessionCode: SESSION, idNumber: id }, over);

test('40 examinees polling one session cost exactly one upstream execution', async () => {
  const ids = Array.from({ length: 40 }, (_, i) => String(900000001 + i));
  const { state, gateway } = harness({ [SESSION]: ids.map(id => row({ id, status: 'approved', examMinutes: 50 })) });

  const answers = await Promise.all(ids.map(id => poll(gateway, approvalPoll(id))));

  assert.equal(state.calls.length, 1, 'one snapshot read for the whole class');
  assert.match(state.calls[0], /action=sessionSnapshot/);
  assert.match(state.calls[0], /gatewayKey=test-secret/);
  assert.equal(answers.length, 40);
  for (const { res, body } of answers) {
    assert.equal(res.status, 200);
    assert.equal(res.headers.get('Cache-Control'), 'no-store');
    assert.deepEqual(body, { status: 'ok', approval: 'approved', audioMode: 'off', examMinutes: 50 });
  }
});

test('a poll inside the 2 s window is free; past it costs one more execution', async () => {
  const { state, gateway } = harness({ [SESSION]: [row({ status: 'approved' })] });

  await poll(gateway, approvalPoll('900000001'));
  assert.equal(state.calls.length, 1);

  state.clock += 1500;
  await poll(gateway, approvalPoll('900000001'));
  assert.equal(state.calls.length, 1, 'still inside the freshness window');

  state.clock += 1000; // 2.5s after the first read
  await poll(gateway, approvalPoll('900000001'));
  assert.equal(state.calls.length, 2, 'snapshot expired, one new read');
});

test('sessions are coalesced separately', async () => {
  const { state, gateway } = harness({ ABC12345: [row()], ZZ999999: [row({ id: '900000002' })] });
  await Promise.all([
    poll(gateway, approvalPoll('900000001')),
    poll(gateway, { kind: 'approval', sessionCode: 'ZZ999999', idNumber: '900000002' })
  ]);
  assert.equal(state.calls.length, 2);
});

test('approval skips terminal rows and reports audio and authorised minutes', async () => {
  const { gateway } = harness({ [SESSION]: [
    row({ status: 'completed' }),
    row({ status: 'approved', audio: 'on', examMinutes: 60 }),
    row({ id: '900000002', status: 'completed' })
  ] });

  const active = await poll(gateway, approvalPoll('900000001'));
  assert.deepEqual(active.body, { status: 'ok', approval: 'approved', audioMode: 'on', examMinutes: 60 });

  const allTerminal = await poll(gateway, approvalPoll('900000002'));
  assert.deepEqual(allTerminal.body, { status: 'error', message: 'לא נמצא רישום' });
});

// The examiner's decision about the REGISTRATION reaches the examinee as
// itself: 'הבוחן דחה' / 'ההרשמה בוטלה' with a way back to the code screen,
// instead of the 'לא נמצא רישום' error the page used to show them.
test('the newest decision answers when nothing live is left, and audio/minutes are dropped with it', async () => {
  const { gateway } = harness({ [SESSION]: [
    row({ id: '900000001', status: 'rejected', audio: 'on', examMinutes: 60 }),
    row({ id: '900000002', status: 'cancelled', audio: 'on', examMinutes: 60 }),
    row({ id: '900000003', status: 'disqualified' }),
    // The shared-ID incident, as the sheet held it: rejected at 17:47, then a
    // second examinee on the same id cancelled at 18:05.
    row({ id: '900000004', status: 'rejected' }),
    row({ id: '900000004', status: 'cancelled' }),
    // ...and a live row still outranks any decision above it.
    row({ id: '900000005', status: 'rejected' }),
    row({ id: '900000005', status: 'waiting' })
  ] });

  assert.deepEqual((await poll(gateway, approvalPoll('900000001'))).body, { status: 'ok', approval: 'rejected' });
  assert.deepEqual((await poll(gateway, approvalPoll('900000002'))).body, { status: 'ok', approval: 'cancelled' });
  assert.deepEqual((await poll(gateway, approvalPoll('900000003'))).body, { status: 'error', message: 'לא נמצא רישום' });
  assert.deepEqual((await poll(gateway, approvalPoll('900000004'))).body, { status: 'ok', approval: 'cancelled' },
    'the NEWEST row decides — the third visitor is never shown the older rejection');
  assert.deepEqual((await poll(gateway, approvalPoll('900000005'))).body,
    { status: 'ok', approval: 'waiting', audioMode: 'off' });
});

test('approval defaults an empty status to waiting and omits minutes until approved', async () => {
  const { gateway } = harness({ [SESSION]: [row({ status: '' })] });
  const { body } = await poll(gateway, approvalPoll('900000001'));
  assert.deepEqual(body, { status: 'ok', approval: 'waiting', audioMode: 'off' });
});

test('status returns the last row whatever its state, plus extraMinutes', async () => {
  const { gateway } = harness({ [SESSION]: [
    row({ status: 'approved' }),
    row({ status: 'disqualified', extraMinutes: 12 })
  ] });

  const found = await poll(gateway, statusPoll('900000001'));
  assert.deepEqual(found.body, { status: 'ok', examStatus: 'disqualified', extraMinutes: 12 });

  const missing = await poll(gateway, statusPoll('900000009'));
  assert.deepEqual(missing.body, { status: 'ok', examStatus: 'not_found' });
});

test('ids are normalised the way the server normalises them', async () => {
  const { gateway } = harness({ [SESSION]: [row({ id: 123456789, status: 'approved' })] });
  const { body } = await poll(gateway, approvalPoll('123-456-789'));
  assert.equal(body.approval, 'approved');

  const padded = harness({ [SESSION]: [row({ id: '000000012', status: 'in_exam' })] });
  const short = await poll(padded.gateway, approvalPoll('12'));
  assert.equal(short.body.approval, 'in_exam');
});

test('a token whose hash differs from the stored one is rejected', async () => {
  const rows = [row({ status: 'approved', tokenHash: sha256Hex('real-token') })];
  const { gateway } = harness({ [SESSION]: rows });

  const wrong = await poll(gateway, approvalPoll('900000001', { examineeToken: 'stolen-token' }));
  assert.deepEqual(wrong.body, { status: 'error', message: 'טוקן נבחן לא תקין', examineeTokenError: 'mismatch' });

  const right = await poll(gateway, approvalPoll('900000001', { examineeToken: '  real-token  ' }));
  assert.equal(right.body.approval, 'approved', 'the token is trimmed before hashing, like the server');

  const noToken = await poll(gateway, approvalPoll('900000001'));
  assert.equal(noToken.body.approval, 'approved', 'a client that sent no token is not rejected');

  const statusWrong = await poll(gateway, statusPoll('900000001', { examineeToken: 'stolen-token' }));
  assert.deepEqual(statusWrong.body, { status: 'error', examineeTokenError: 'mismatch' });
});

test('a row missing from a cached snapshot forces one re-read, at most once per 2 s', async () => {
  const { state, gateway } = harness({ [SESSION]: [row({ id: '900000001', status: 'approved' })] });

  await poll(gateway, approvalPoll('900000001'));
  assert.equal(state.calls.length, 1);

  // The examinee registered a moment after the snapshot was taken.
  state.snapshots[SESSION].push(row({ id: '900000002', status: 'waiting' }));
  const late = await poll(gateway, approvalPoll('900000002'));
  assert.equal(state.calls.length, 2, 'one forced re-read before answering "not registered"');
  assert.equal(late.body.approval, 'waiting');

  state.clock += 500;
  const again = await poll(gateway, approvalPoll('900000003'));
  assert.equal(state.calls.length, 2, 'no second forced re-read inside the 2 s window');
  assert.deepEqual(again.body, { status: 'error', message: 'לא נמצא רישום' });

  state.clock += 500;
  await poll(gateway, statusPoll('900000003'));
  assert.equal(state.calls.length, 2, 'the throttle is per session, not per kind');
});

test('a fresh snapshot with no row is answered without a second read', async () => {
  const { state, gateway } = harness({ [SESSION]: [] });
  const { body } = await poll(gateway, approvalPoll('900000001'));
  assert.equal(state.calls.length, 1, 'the data was read this instant — re-reading it proves nothing');
  assert.deepEqual(body, { status: 'error', message: 'לא נמצא רישום' });
});

for (const mode of ['error500', 'notfound', 'html200']) {
  test('upstream ' + mode + ': the stale snapshot is served, marked stale', async () => {
    const { state, gateway } = harness({ [SESSION]: [row({ status: 'approved', examMinutes: 50 })] });
    await poll(gateway, approvalPoll('900000001'));

    state.mode = mode;
    state.clock += 4000; // past the freshness window, inside the 60s stale window
    const { res, body } = await poll(gateway, approvalPoll('900000001'));

    assert.equal(state.calls.length, 2);
    assert.equal(res.status, 200);
    assert.deepEqual(body, { status: 'ok', approval: 'approved', audioMode: 'off', examMinutes: 50, stale: true });
    assert.ok(!/html/i.test(JSON.stringify(body)), 'Google HTML never reaches the examinee');
  });
}

test('upstream failure with nothing cached is a retryable error, not an HTML page', async () => {
  const { state, gateway } = harness({ [SESSION]: [row()] });
  state.mode = 'error500';

  const { res, body } = await poll(gateway, approvalPoll('900000001'));
  assert.equal(res.status, 200, 'HTTP 200 so the client parses it instead of guessing');
  assert.deepEqual(body, {
    status: 'error', code: 'upstream_unavailable', retryable: true,
    message: 'השרת עמוס — ננסה שוב אוטומטית'
  });
  assert.equal(state.calls.length, 1);
});

test('a snapshot older than 60s is not served at all', async () => {
  const { state, gateway } = harness({ [SESSION]: [row({ status: 'approved' })] });
  await poll(gateway, approvalPoll('900000001'));

  state.mode = 'error500';
  state.clock += 61000;
  const { body } = await poll(gateway, approvalPoll('900000001'));
  assert.equal(body.code, 'upstream_unavailable');
  assert.equal(state.calls.length, 2);
});

test('a second isolate is served from caches.default without a new read', async () => {
  // The in-memory map dies with the isolate; caches.default is what keeps the
  // coalescing working across the isolates of one Cloudflare location.
  const { state, spawn } = harness({ [SESSION]: [row({ status: 'approved' })] }, null, true);

  const first = await poll(spawn(), approvalPoll('900000001'));
  assert.equal(first.body.approval, 'approved');
  assert.equal(state.calls.length, 1);
  assert.equal(state.store.size, 2, 'a fresh copy and a stale copy');

  const second = await poll(spawn(), approvalPoll('900000001'));
  assert.equal(state.calls.length, 1, 'the cold isolate reused the cached snapshot');
  assert.equal(second.body.approval, 'approved');

  state.clock += 4000; // fresh copy expired, stale copy still there
  const third = spawn();
  await poll(third, approvalPoll('900000001'));
  assert.equal(state.calls.length, 2);
});

test('CORS: the Pages origin and localhost are echoed, anything else is not', async () => {
  const { gateway } = harness({ [SESSION]: [row()] });

  const pages = await poll(gateway, approvalPoll('900000001'), PAGES_ORIGIN);
  assert.equal(pages.res.headers.get('Access-Control-Allow-Origin'), PAGES_ORIGIN);

  const local = await poll(gateway, approvalPoll('900000001'), 'http://localhost:8000');
  assert.equal(local.res.headers.get('Access-Control-Allow-Origin'), 'http://localhost:8000');

  const evil = await poll(gateway, approvalPoll('900000001'), 'https://evil.example');
  assert.equal(evil.res.headers.get('Access-Control-Allow-Origin'), PAGES_ORIGIN, 'not echoed');
});

test('OPTIONS preflight is answered without touching the upstream', async () => {
  const { state, gateway } = harness({ [SESSION]: [row()] });
  const res = await gateway(new Request('https://session-gateway.test/v1/poll', {
    method: 'OPTIONS', headers: { Origin: PAGES_ORIGIN }
  }));
  assert.equal(res.status, 204);
  assert.equal(res.headers.get('Access-Control-Allow-Origin'), PAGES_ORIGIN);
  assert.equal(res.headers.get('Access-Control-Allow-Methods'), 'GET, POST, OPTIONS');
  assert.equal(state.calls.length, 0);
});

test('the health route names the service, its build and the deployed bank', async () => {
  const { gateway } = harness({}, ASSETS);
  const { body } = await call(gateway, '/');
  assert.equal(body.status, 'ok');
  assert.equal(body.service, 'session-gateway');
  assert.match(body.build, /^\d{4}-\d{2}-\d{2}$/);
  assert.equal(body.bank, BANK_BUILD, 'which bank is deployed is the first thing to check after a deploy');

  const bare = await call(harness({}).gateway, '/');
  assert.equal(bare.body.bank, '', 'no assets is an empty build id, not a crash');
});

test('bad parameters are refused with 400 before any upstream read', async () => {
  const { state, gateway } = harness({ [SESSION]: [row()] });
  const bad = [
    approvalPoll('900000001', { kind: 'everything' }),
    approvalPoll('900000001', { kind: '' }),
    { kind: 'approval', sessionCode: 'abc12345', idNumber: '900000001' },   // lower case
    { kind: 'approval', sessionCode: 'ABC12', idNumber: '900000001' },      // too short
    { kind: 'approval', sessionCode: 'ABC123456', idNumber: '900000001' },  // too long
    { kind: 'approval', sessionCode: 'ABC-1234', idNumber: '900000001' },   // punctuation
    { kind: 'approval', sessionCode: SESSION, idNumber: '' },
    { kind: 'approval', sessionCode: SESSION, idNumber: 'abcdefghi' }
  ];
  for (const params of bad) {
    const { res, body } = await poll(gateway, params);
    assert.equal(res.status, 400, JSON.stringify(params));
    assert.equal(body.status, 'error');
    assert.equal(res.headers.get('Access-Control-Allow-Origin'), PAGES_ORIGIN);
  }
  assert.equal(state.calls.length, 0, 'garbage never reaches Apps Script');

  const unknownPath = await gateway(new Request('https://session-gateway.test/v1/nope'));
  assert.equal(unknownPath.status, 404);
  const posted = await gateway(new Request('https://session-gateway.test/v1/poll', { method: 'POST' }));
  assert.equal(posted.status, 405);
});

// --- the private bank ------------------------------------------------------
// The whole point of §11: the texts are not public, and what a device gets is
// decided by a signature it cannot produce.

test('bank: an exam grant is served exactly its own ids, as raw asset text, in order', async () => {
  const { gateway } = harness({}, ASSETS);
  const { res, text, body } = await call(gateway, '/v1/bank?grant=' + examGrant([120, 999, 14]));

  assert.equal(res.status, 200);
  assert.equal(res.headers.get('Cache-Control'), 'no-store');
  assert.equal(res.headers.get('Access-Control-Allow-Origin'), PAGES_ORIGIN);
  assert.equal(body.status, 'ok');
  assert.equal(body.build, BANK_BUILD);
  assert.deepEqual(body.missing, [999], 'an id with no asset file is named, not silently dropped');
  // Raw text, in grant order: the hot path must never JSON.parse the assets.
  assert.equal(text, '{"status":"ok","build":"' + BANK_BUILD + '","questions":[' + Q120 + ',' + Q14 +
    '],"missing":[999]}');
  assert.deepEqual(body.questions.map(q => q.id), [120, 14]);
  assert.deepEqual(Object.keys(body.questions[1].l), ['he', 'ru', 'en'], 'every language of the id');
});

test('bank: a practice grant is served the same way', async () => {
  const { gateway } = harness({}, ASSETS);
  const practice = grant({ s: 'practice', ids: [14], sub: 'CLASS1:student7' });
  const { res, body } = await call(gateway, '/v1/bank?grant=' + practice);
  assert.equal(res.status, 200);
  assert.deepEqual(body.questions.map(q => q.id), [14]);
  assert.deepEqual(body.missing, []);
});

test('bank: the grant decides the ids — a query string cannot widen an exam grant', async () => {
  const { gateway } = harness({}, ASSETS);
  const { body } = await call(gateway, '/v1/bank?ids=14,120&langs=he&grant=' + examGrant([14]));
  assert.deepEqual(body.questions.map(q => q.id), [14]);
  assert.deepEqual(Object.keys(body.questions[0].l), ['he', 'ru', 'en'], 'langs is examiner-only');
});

test('bank: an expired grant is refused with 403 grant_invalid', async () => {
  const { gateway } = harness({}, ASSETS);
  const expired = examGrant([14], { exp: CLOCK0 - 1 });
  const { res, body } = await call(gateway, '/v1/bank?grant=' + expired);
  assert.equal(res.status, 403);
  assert.deepEqual(body, { status: 'error', code: 'grant_invalid' });
});

test('bank: a tampered payload, a foreign key and rubbish are all grant_invalid', async () => {
  const { gateway } = harness({}, ASSETS);
  const honest = examGrant([14]);
  const [payload, sig] = honest.split('.');
  const forged = Buffer.from(JSON.stringify(
    { v: 1, s: 'exam', ids: [14, 120, 777], sub: 'x', exp: CLOCK0 + 3600000 }), 'utf8').toString('base64url');

  for (const bad of [
    forged + '.' + sig,                                        // payload swapped, old signature
    payload + '.' + b64url(Buffer.alloc(32)),                  // signature swapped
    grant({ s: 'exam', ids: [14], sub: 'x' }, 'another-secret'), // signed with the wrong key
    grant({ s: 'exam', ids: [14], v: 2 }),                     // unknown grant version
    payload,                                                   // no signature at all
    'bogus', '', '.', 'a.b'
  ]) {
    const { res, body } = await call(gateway, '/v1/bank?grant=' + encodeURIComponent(bad));
    assert.equal(res.status, 403, JSON.stringify(bad).slice(0, 40));
    assert.deepEqual(body, { status: 'error', code: 'grant_invalid' });
  }

  const still = await call(gateway, '/v1/bank?grant=' + honest);
  assert.equal(still.res.status, 200, 'the honest grant still works');
});

test('bank: an examiner grant asks for its own ids, and langs narrows the languages', async () => {
  const { gateway } = harness({}, ASSETS);
  const g = examinerGrant();

  const all = await call(gateway, '/v1/bank?grant=' + g + '&ids=14,120');
  assert.deepEqual(all.body.questions.map(q => q.id), [14, 120]);

  const narrow = await call(gateway, '/v1/bank?grant=' + g + '&ids=14,120&langs=he,en');
  assert.deepEqual(narrow.body.questions.map(q => Object.keys(q.l)), [['he', 'en'], ['he', 'en']]);
  assert.equal(narrow.body.questions[0].l.he.t, 'שאלה 14');

  const noIds = await call(gateway, '/v1/bank?grant=' + g);
  assert.equal(noIds.res.status, 400);

  const tooMany = await call(gateway,
    '/v1/bank?grant=' + g + '&ids=' + Array.from({ length: 61 }, (_, i) => i + 1).join(','));
  assert.equal(tooMany.res.status, 400, 'the cap keeps one request inside the subrequest budget');

  const unknownLang = await call(gateway, '/v1/bank?grant=' + g + '&ids=14&langs=klingon');
  assert.equal(unknownLang.res.status, 400);
});

test('bank/full: only an examiner grant streams a whole language', async () => {
  const { gateway } = harness({}, ASSETS);

  const forbidden = await call(gateway, '/v1/bank/full?lang=he&grant=' + examGrant([14]));
  assert.equal(forbidden.res.status, 403, 'an exam device must never get the whole bank');
  assert.deepEqual(forbidden.body, { status: 'error', code: 'grant_invalid' });

  const g = examinerGrant();
  const full = await call(gateway, '/v1/bank/full?lang=he&grant=' + g);
  assert.equal(full.res.status, 200);
  assert.equal(full.res.headers.get('Cache-Control'), 'no-store');
  assert.equal(full.text, HE_BANK, 'the asset body is passed through untouched');

  const undeployed = await call(gateway, '/v1/bank/full?lang=ru&grant=' + g);
  assert.equal(undeployed.res.status, 404);
  assert.equal(undeployed.body.code, 'bank_missing');

  const nonsense = await call(gateway, '/v1/bank/full?lang=klingon&grant=' + g);
  assert.equal(nonsense.res.status, 400);
});

test('bank: no assets binding, or reads that all throw, is a retryable 503', async () => {
  const unbound = await call(harness({}).gateway, '/v1/bank?grant=' + examGrant([14]));
  assert.equal(unbound.res.status, 503);
  assert.deepEqual(unbound.body, { status: 'error', code: 'bank_unavailable', retryable: true });

  const broken = harness({}, { fetch: async () => { throw new Error('asset server down'); } });
  const thrown = await call(broken.gateway, '/v1/bank?grant=' + examGrant([14, 120]));
  assert.equal(thrown.res.status, 503, 'a broken binding is not 30 "missing" questions');
  assert.equal(thrown.body.code, 'bank_unavailable');

  const full = await call(broken.gateway, '/v1/bank/full?lang=he&grant=' + examinerGrant());
  assert.equal(full.res.status, 503);
});

test('/v1/invalidate drops the snapshot, at most once per 2 s', async () => {
  const { state, gateway } = harness({ [SESSION]: [row({ status: 'waiting' })] });
  const push = () => call(gateway, '/v1/invalidate?sessionCode=' + SESSION + '&grant=' + examinerGrant(), { method: 'POST' });

  await poll(gateway, approvalPoll('900000001'));
  assert.equal(state.calls.length, 1);

  // The examiner approves and pushes. The examinee is still inside the 2 s
  // freshness window, and must see the decision on its very next poll.
  state.snapshots[SESSION] = [row({ status: 'approved', examMinutes: 40 })];
  const pushed = await push();
  assert.equal(pushed.res.status, 200);
  assert.deepEqual(pushed.body, { status: 'ok' });

  const after = await poll(gateway, approvalPoll('900000001'));
  assert.equal(state.calls.length, 2, 'the dropped snapshot forced a fresh read inside the window');
  assert.equal(after.body.approval, 'approved');

  // A burst of pushes must not buy a burst of Apps Script executions.
  state.snapshots[SESSION] = [row({ status: 'in_exam' })];
  const throttled = await push();
  assert.deepEqual(throttled.body, { status: 'ok' }, 'still ok — the caller fires and forgets');
  const again = await poll(gateway, approvalPoll('900000001'));
  assert.equal(state.calls.length, 2, 'the second push inside 2 s bought nothing');
  assert.equal(again.body.approval, 'approved');

  state.clock += 2000;
  await push();
  await poll(gateway, approvalPoll('900000001'));
  assert.equal(state.calls.length, 3, 'past the gap the next push works again');

  const junk = await call(gateway, '/v1/invalidate?sessionCode=nope&grant=' + examinerGrant(), { method: 'POST' });
  assert.equal(junk.res.status, 400);
  assert.equal(state.calls.length, 3);
});

// --- the nudge that CARRIES the decision -----------------------------------
// A drop still costs the examinee one Apps Script read before they see the
// approval. A patch costs none: the examiner already has the answer Apps Script
// confirmed, so it is written straight into the snapshot.

/** The examiner's fire-and-forget POST, optionally carrying the decision. */
const nudge = (gateway, query, grantStr) =>
  call(gateway, '/v1/invalidate?sessionCode=' + SESSION + (query || '') +
    '&grant=' + (grantStr === undefined ? examinerGrant() : grantStr), { method: 'POST' });

test('invalidate is an examiner-only door: no grant, a forged one, another scope or an expired one is refused', async () => {
  const { state, gateway } = harness({ [SESSION]: [row({ status: 'in_exam' })] });
  await poll(gateway, statusPoll('900000001'));
  assert.equal(state.calls.length, 1);

  // What a classmate with the session code and an id could send from a phone.
  const forged = grant({ s: 'examiner', sub: 'ex:7' }, 'not-the-key');
  const expired = examinerGrant({ exp: CLOCK0 - 1 });
  const wrongScope = examGrant([1, 2]);
  for (const [name, bad] of [['none', ''], ['forged', forged], ['expired', expired], ['exam scope', wrongScope]]) {
    const { res, body } = await nudge(gateway, '&idNumber=900000001&status=disqualified', bad);
    assert.equal(res.status, 403, name);
    assert.deepEqual(body, { status: 'error', code: 'grant_invalid' }, name);
    const plain = await nudge(gateway, '', bad);
    assert.equal(plain.res.status, 403, name + ' (plain drop)');
  }

  // Neither a patch nor a drop happened: the examinee is still in the exam and
  // the snapshot was not dropped (no extra upstream read).
  const still = await poll(gateway, statusPoll('900000001'));
  assert.equal(still.body.examStatus, 'in_exam', 'a refused nudge never reaches the snapshot');
  assert.equal(state.calls.length, 1, 'and never buys an upstream read either');

  // The real examiner, with the grant the server signed, still gets through.
  const ok = await nudge(gateway, '&idNumber=900000001&status=disqualified');
  assert.deepEqual(ok.body, { status: 'ok', patched: true });
  assert.equal((await poll(gateway, statusPoll('900000001'))).body.examStatus, 'disqualified');
});

test('a patched approval is answered on the next poll with no upstream read', async () => {
  const { state, gateway } = harness({ [SESSION]: [row({ status: 'waiting' })] });

  const before = await poll(gateway, approvalPoll('900000001'));
  assert.equal(before.body.approval, 'waiting');
  assert.equal(state.calls.length, 1);

  // The examiner approved. Apps Script has already written the row — the nudge
  // carries what it wrote. The fake upstream is deliberately NOT updated here:
  // what is proven is that the examinee sees the decision without reading it.
  const patched = await nudge(gateway, '&idNumber=900000001&status=approved&examMinutes=50&audio=on');
  assert.equal(patched.res.status, 200);
  assert.deepEqual(patched.body, { status: 'ok', patched: true });

  state.clock += 1000; // the examinee's very next poll, ~1 s after the click
  const after = await poll(gateway, approvalPoll('900000001'));
  assert.deepEqual(after.body, { status: 'ok', approval: 'approved', audioMode: 'on', examMinutes: 50 });
  assert.equal(state.calls.length, 1, 'the decision reached the device for zero executions');

  // And the patch is short lived: past FRESH_MS the server is the truth again.
  state.snapshots[SESSION] = [row({ status: 'in_exam', examMinutes: 40 })];
  state.clock += 1500;
  const truth = await poll(gateway, approvalPoll('900000001'));
  assert.equal(state.calls.length, 2, 'one read, exactly like an unpatched snapshot');
  assert.equal(truth.body.approval, 'in_exam');
});

test('a patched disqualification shows at once in the status kind, with extraMinutes', async () => {
  const { state, gateway } = harness({ [SESSION]: [
    row({ id: '900000001', status: 'in_exam' }),
    row({ id: '900000002', status: 'completed' })
  ] });
  await poll(gateway, statusPoll('900000001'));
  assert.equal(state.calls.length, 1);

  assert.deepEqual((await nudge(gateway,
    '&idNumber=900000001&status=disqualified&extraMinutes=15')).body, { status: 'ok', patched: true });

  const dq = await poll(gateway, statusPoll('900000001'));
  assert.deepEqual(dq.body, { status: 'ok', examStatus: 'disqualified', extraMinutes: 15 });

  const neighbour = await poll(gateway, statusPoll('900000002'));
  assert.deepEqual(neighbour.body, { status: 'ok', examStatus: 'completed', extraMinutes: 0 },
    'only the named row changes');
  assert.equal(state.calls.length, 1);
});

test('invalidate refuses an unknown status, and a nudge with no id just drops', async () => {
  const { state, gateway } = harness({ [SESSION]: [row({ status: 'waiting' })] });
  await poll(gateway, approvalPoll('900000001'));
  assert.equal(state.calls.length, 1);

  for (const bad of ['&idNumber=900000001&status=approve',
                     '&idNumber=900000001&status=ended',
                     '&idNumber=900000001&status=in exam',
                     '&idNumber=abcdefghi&status=approved']) {
    const { res, body } = await nudge(gateway, bad);
    assert.equal(res.status, 400, bad);
    assert.equal(body.status, 'error');
  }
  const stillThere = await poll(gateway, approvalPoll('900000001'));
  assert.equal(state.calls.length, 1, 'a refused nudge changed nothing and dropped nothing');
  assert.equal(stillThere.body.approval, 'waiting');

  // A status with no id names no row, so it degrades to today's plain drop.
  const dropped = await nudge(gateway, '&status=approved');
  assert.deepEqual(dropped.body, { status: 'ok', patched: false });

  state.snapshots[SESSION] = [row({ status: 'approved' })];
  const after = await poll(gateway, approvalPoll('900000001'));
  assert.equal(state.calls.length, 2, 'the snapshot was gone, so the poll read the truth');
  assert.equal(after.body.approval, 'approved');
});

test('a patch for an id the snapshot never had falls back to dropping it', async () => {
  const { state, gateway } = harness({ [SESSION]: [row({ id: '900000001', status: 'waiting' })] });
  await poll(gateway, approvalPoll('900000001'));
  assert.equal(state.calls.length, 1);

  // This examinee registered after the snapshot was taken: there is no row to
  // write into, and the nudge must not invent one.
  state.snapshots[SESSION] = [
    row({ id: '900000001', status: 'waiting' }),
    row({ id: '900000002', status: 'approved', examMinutes: 40 })
  ];
  assert.deepEqual((await nudge(gateway, '&idNumber=900000002&status=approved')).body,
    { status: 'ok', patched: false });

  const known = await poll(gateway, approvalPoll('900000001'));
  assert.equal(state.calls.length, 2, 'the snapshot was dropped — even a row it held is re-read');
  assert.equal(known.body.approval, 'waiting');

  const late = await poll(gateway, approvalPoll('900000002'));
  assert.equal(late.body.approval, 'approved', 'and the row the server has is there');
});

test('a patch does not spend the forced re-read budget', async () => {
  const { state, gateway } = harness({ [SESSION]: [row({ status: 'waiting' })] });
  await poll(gateway, approvalPoll('900000001'));
  assert.equal(state.calls.length, 1);

  assert.equal((await nudge(gateway, '&idNumber=900000001&status=approved')).body.patched, true);

  // A write is not an upstream read, so the plain invalidate right behind it
  // still gets its drop — the examiner fires several of these in a row.
  const plain = await nudge(gateway);
  assert.deepEqual(plain.body, { status: 'ok' }, 'a plain nudge keeps its old body exactly');

  state.snapshots[SESSION] = [row({ status: 'in_exam' })];
  const after = await poll(gateway, approvalPoll('900000001'));
  assert.equal(state.calls.length, 2, 'the drop worked: the patch never touched the gate');
  assert.equal(after.body.approval, 'in_exam');
});

test('a cold isolate serves the patched snapshot from caches.default', async () => {
  const { state, spawn } = harness({ [SESSION]: [row({ status: 'waiting' })] }, null, true);

  await poll(spawn(), approvalPoll('900000001'));
  assert.equal(state.calls.length, 1);

  // A different isolate takes the nudge: it has no memory of the session and
  // must patch the cached copy.
  const patched = await nudge(spawn(), '&idNumber=900000001&status=approved&examMinutes=50&audio=on');
  assert.deepEqual(patched.body, { status: 'ok', patched: true });

  state.clock += 1000;
  const after = await poll(spawn(), approvalPoll('900000001'));
  assert.deepEqual(after.body, { status: 'ok', approval: 'approved', audioMode: 'on', examMinutes: 50 });
  assert.equal(state.calls.length, 1, 'a third isolate needed no upstream read either');

  state.snapshots[SESSION] = [row({ status: 'in_exam' })];
  state.clock += 1500; // the patched copy has expired like any other
  await poll(spawn(), approvalPoll('900000001'));
  assert.equal(state.calls.length, 2);
});

test('with only the stale copy left, a nudge still patches it', async () => {
  const { state, spawn } = harness({ [SESSION]: [row({ status: 'waiting' })] }, null, true);
  await poll(spawn(), approvalPoll('900000001'));

  state.clock += 3000;  // the `snap` copy is gone; `stale` lives 60 s
  const patched = await nudge(spawn(), '&idNumber=900000001&status=approved&examMinutes=50');
  assert.deepEqual(patched.body, { status: 'ok', patched: true });

  const after = await poll(spawn(), approvalPoll('900000001'));
  assert.deepEqual(after.body, { status: 'ok', approval: 'approved', audioMode: 'off', examMinutes: 50 });
  assert.equal(state.calls.length, 1, 'still no upstream read — and no stale:true, the copy is fresh again');
});

// --- long polling: `fp` + `wait` -------------------------------------------
// The client sends the fingerprint it already has and how long it is willing to
// wait; the request is HELD here until the answer actually changes. That is
// ~4-5x fewer requests against the account's 100k/day, and an examiner's
// decision on the screen in ~1 s instead of one poll interval later.

const WAITING_FP = 'a:waiting:off:-';

test('fp names every answer kind, is stable while the answer is, and changes with it', async () => {
  const { state, gateway } = harness({ [SESSION]: [
    row({ id: '900000001', status: 'approved', examMinutes: 50 }),
    row({ id: '900000002', status: 'waiting' }),
    row({ id: '900000003', status: 'in_exam', extraMinutes: 7 }),
    row({ id: '900000004', status: 'approved', tokenHash: sha256Hex('real-token') })
  ] });
  const fp = async params => (await poll(gateway, params)).body.fp;

  assert.equal(await fp(approvalPoll('900000001', { wait: 0 })), 'a:approved:off:50');
  assert.equal(await fp(approvalPoll('900000002', { wait: 0 })), WAITING_FP, 'no minutes until approved');
  assert.equal(await fp(statusPoll('900000003', { fp: '' })), 's:in_exam:7');
  assert.equal(await fp(approvalPoll('900000009', { wait: 0 })), 'a:none', 'no row has its own fingerprint');
  assert.equal(await fp(statusPoll('900000009', { wait: 0 })), 's:none');
  assert.equal(await fp(approvalPoll('900000004', { wait: 0, examineeToken: 'stolen' })), 'a:tok');
  assert.equal(await fp(statusPoll('900000004', { wait: 0, examineeToken: 'stolen' })), 's:tok');
  assert.equal(await fp({ kind: 'nonsense', sessionCode: SESSION, idNumber: '1', wait: 0 }), 'x:kind');
  assert.equal(await fp({ kind: 'approval', sessionCode: 'nope', idNumber: '1', wait: 0 }), 'x:sess');
  assert.equal(await fp({ kind: 'approval', sessionCode: SESSION, idNumber: '', wait: 0 }), 'x:id');

  // The same answer, later, from a re-read snapshot: the fingerprint may not
  // move, or every client would poll on for ever.
  state.clock += 5000;
  assert.equal(await fp(approvalPoll('900000001', { wait: 0 })), 'a:approved:off:50');
  assert.ok(state.calls.length > 1, 'and that really was a fresh read, not the same snapshot');

  // Each of the three things an approval answer carries moves it, and so does
  // each of the two a status answer carries.
  state.snapshots[SESSION] = [
    row({ id: '900000001', status: 'approved', audio: 'on', examMinutes: 50 }),
    row({ id: '900000003', status: 'in_exam', extraMinutes: 12 })
  ];
  state.clock += 5000;
  assert.equal(await fp(approvalPoll('900000001', { wait: 0 })), 'a:approved:on:50');
  assert.equal(await fp(statusPoll('900000003', { wait: 0 })), 's:in_exam:12');
  state.snapshots[SESSION] = [
    row({ id: '900000001', status: 'in_exam', audio: 'on', examMinutes: 60 }),
    row({ id: '900000003', status: 'disqualified', extraMinutes: 12 })
  ];
  state.clock += 5000;
  assert.equal(await fp(approvalPoll('900000001', { wait: 0 })), 'a:in_exam:on:60');
  assert.equal(await fp(statusPoll('900000003', { wait: 0 })), 's:disqualified:12');
  assert.ok(state.logs.every(line => line.held === 0), 'nothing here ever held');
  assert.ok(state.logs.every(line => line.r === 'poll' && line.s === SESSION),
    'and every answer left exactly one line for `wrangler tail`');
});

test('a client that says nothing about long polling gets byte-for-byte the old answer', async () => {
  const { gateway } = harness({ [SESSION]: [row({ status: 'approved', examMinutes: 50 })] });

  // This is what tests/contracts.test.cjs pins: the gateway's answer IS the
  // server's answer. `fp`/`held` are served only to a client that named itself
  // by sending `wait` or `fp`.
  const old = await poll(gateway, approvalPoll('900000001'));
  assert.deepEqual(old.body, { status: 'ok', approval: 'approved', audioMode: 'off', examMinutes: 50 });

  const modern = await poll(gateway, approvalPoll('900000001', { wait: 0 }));
  assert.deepEqual(modern.body, {
    status: 'ok', approval: 'approved', audioMode: 'off', examMinutes: 50,
    fp: 'a:approved:off:50', held: 0
  });
  const byFpAlone = await poll(gateway, approvalPoll('900000001', { fp: 'whatever' }));
  assert.equal(byFpAlone.body.fp, 'a:approved:off:50', 'fp alone opts in too, and never holds');
  assert.equal(byFpAlone.body.held, 0);
});

test('a hold runs out its `wait` and answers the same fingerprint', async () => {
  const { state, gateway } = harness({ [SESSION]: [row({ status: 'waiting' })] });
  const first = await poll(gateway, approvalPoll('900000001', { wait: 25, fp: '' }));
  assert.equal(first.body.fp, WAITING_FP, 'no fp to send yet, so it is answered at once');
  assert.equal(first.body.held, 0);
  assert.equal(state.calls.length, 1);

  const { body, timers } = await timed(state, poll(gateway, approvalPoll('900000001', { wait: 5, fp: WAITING_FP })));

  assert.deepEqual(body, { status: 'ok', approval: 'waiting', audioMode: 'off', fp: WAITING_FP, held: 5000 });
  assert.equal(state.clock, CLOCK0 + 5000, 'it really waited the five seconds');
  assert.equal(timers, 10, 'five evaluations, one per second, not a spin');
  assert.equal(state.calls.length, 1,
    'and since 22/09 five seconds of holding cost Google NOTHING: this chain already holds this state');
  assert.deepEqual(gateway._debug().waiting, 0, 'and it took its resolver back out');
});

test('an examiner nudge ends the hold at once, for zero executions', async () => {
  const { state, gateway } = harness({ [SESSION]: [row({ status: 'waiting' })] });
  await poll(gateway, approvalPoll('900000001'));
  assert.equal(state.calls.length, 1);

  const holding = poll(gateway, approvalPoll('900000001', { wait: 25, fp: WAITING_FP }));
  assert.equal(await parked(gateway, 1), 1, 'parked');

  const pushed = await nudge(gateway, '&idNumber=900000001&status=approved&examMinutes=50&audio=on');
  assert.deepEqual(pushed.body, { status: 'ok', patched: true });

  const { body } = await settle(state, holding);
  assert.deepEqual(body, {
    status: 'ok', approval: 'approved', audioMode: 'on', examMinutes: 50,
    fp: 'a:approved:on:50', held: 0
  });
  assert.equal(state.calls.length, 1, 'the decision reached the device with no upstream read at all');
  assert.equal(state.clock, CLOCK0, 'and without waiting out a single tick');
  assert.equal(gateway._debug().waiting, 0);
});

test('a hold ends the moment the examiner rejects or resets the registration', async () => {
  for (const [status, decided] of [['rejected', 'a:rejected:off:-'], ['cancelled', 'a:cancelled:off:-']]) {
    const { state, gateway } = harness({ [SESSION]: [row({ status: 'waiting', audio: 'on' })] });
    await poll(gateway, approvalPoll('900000001'));
    assert.equal(state.calls.length, 1, status);

    const holding = poll(gateway, approvalPoll('900000001', { wait: 25, fp: 'a:waiting:on:-' }));
    assert.equal(await parked(gateway, 1), 1, status + ': parked');

    assert.deepEqual((await nudge(gateway, '&idNumber=900000001&status=' + status)).body,
      { status: 'ok', patched: true });

    const { body } = await settle(state, holding);
    assert.deepEqual(body, { status: 'ok', approval: status, fp: decided, held: 0 },
      'the decision itself, with no audio and no minutes left to carry');
    assert.equal(state.calls.length, 1, status + ': it reached the device with no upstream read at all');
    assert.equal(state.clock, CLOCK0, status + ': and without waiting out a single tick');
    assert.equal(gateway._debug().waiting, 0);
  }
});

test('the not-found fingerprint is unchanged by the decisions: completed, disqualified and no row at all', async () => {
  const { gateway } = harness({ [SESSION]: [
    row({ id: '900000001', status: 'completed' }),
    row({ id: '900000002', status: 'disqualified' })
  ] });
  const fp = async params => (await poll(gateway, params)).body.fp;
  assert.equal(await fp(approvalPoll('900000001', { wait: 0 })), 'a:none');
  assert.equal(await fp(approvalPoll('900000002', { wait: 0 })), 'a:none');
  assert.equal(await fp(approvalPoll('900000009', { wait: 0 })), 'a:none');
});

test('a status hold ends on the disqualification the nudge carries', async () => {
  const { state, gateway } = harness({ [SESSION]: [row({ status: 'in_exam' })] });
  await poll(gateway, statusPoll('900000001'));

  const holding = poll(gateway, statusPoll('900000001', { wait: 25, fp: 's:in_exam:0' }));
  await parked(gateway, 1);
  await nudge(gateway, '&idNumber=900000001&status=disqualified&extraMinutes=15');

  const { body } = await settle(state, holding);
  assert.deepEqual(body, {
    status: 'ok', examStatus: 'disqualified', extraMinutes: 15, fp: 's:disqualified:15', held: 0
  });
  assert.equal(state.calls.length, 1);
});

test('a plain drop wakes the hold, which reads the row the server now has', async () => {
  const { state, gateway } = harness({ [SESSION]: [row({ status: 'in_exam' })] });
  await poll(gateway, statusPoll('900000001'));

  const holding = poll(gateway, statusPoll('900000001', { wait: 25, fp: 's:in_exam:0' }));
  await parked(gateway, 1);

  // The examiner extended the time: the row is already written, the nudge only
  // says "the snapshot is behind".
  state.snapshots[SESSION] = [row({ status: 'in_exam', extraMinutes: 15 })];
  await nudge(gateway, '');

  const { body } = await settle(state, holding);
  assert.deepEqual(body, {
    status: 'ok', examStatus: 'in_exam', extraMinutes: 15, fp: 's:in_exam:15', held: 0
  });
  assert.equal(state.calls.length, 2, 'the woken request read the truth once');
});

test('a patch another isolate wrote reaches a held request through caches.default', async () => {
  const { state, spawn } = harness({ [SESSION]: [row({ status: 'waiting' })] }, null, true);
  const first = spawn(), second = spawn();

  await poll(first, approvalPoll('900000001'));
  const holding = poll(first, approvalPoll('900000001', { wait: 25, fp: WAITING_FP }));
  await parked(first, 1);

  // A DIFFERENT isolate takes the examiner's nudge: it cannot wake anything in
  // the first one, so the only channel left is the cache — which is why a held
  // request re-reads it on every tick.
  assert.equal((await nudge(second, '&idNumber=900000001&status=approved&examMinutes=50')).body.patched, true);
  assert.equal(first._debug().waiting, 1, 'nothing woke it — it is still parked');

  const { body } = await settle(state, holding);
  assert.deepEqual(body, {
    status: 'ok', approval: 'approved', audioMode: 'off', examMinutes: 50,
    fp: 'a:approved:off:50', held: 1000
  });
  assert.equal(state.calls.length, 1, 'seen within one tick, and still no upstream read');
});

test('forty held requests cost nothing while nothing happens, and one drop ends them all', async () => {
  const ids = Array.from({ length: 40 }, (_, i) => String(900000001 + i));
  const { state, gateway } = harness({ [SESSION]: ids.map(id => row({ id, status: 'waiting' })) });

  await Promise.all(ids.map(id => poll(gateway, approvalPoll(id))));
  assert.equal(state.calls.length, 1);

  const holds = ids.map(id => poll(gateway, approvalPoll(id, { wait: 25, fp: WAITING_FP })));
  assert.equal(await parked(gateway, 40), 40, 'forty parked requests, one session');

  await advance(state, 6000);
  assert.equal(state.calls.length, 1,
    'six seconds of holding, and not one execution: nothing happened, so nobody asked');
  assert.equal(gateway._debug().waiting, 40, 'still holding — nothing they care about changed');

  // The examiner approves ONE examinee. Only that request may end.
  await nudge(gateway, '&idNumber=900000001&status=approved&examMinutes=50');
  const one = await settle(state, holds[0]);
  assert.equal(one.body.fp, 'a:approved:off:50');
  assert.equal(one.body.held, 6000);
  assert.equal(await parked(gateway, 39), 39, 'the rest looked, saw their own row unchanged and parked again');
  assert.equal(state.calls.length, 1, 'and the patch bought no execution');

  // Now the whole class is approved and the examiner fires a plain drop.
  state.snapshots[SESSION] = ids.map(id => row({ id, status: 'approved', examMinutes: 40 }));
  await nudge(gateway, '');
  const rest = await settle(state, Promise.all(holds.slice(1)));

  assert.equal(rest.length, 39);
  for (const answer of rest) {
    assert.equal(answer.body.approval, 'approved');
    assert.equal(answer.body.fp, 'a:approved:off:40');
  }
  assert.equal(state.calls.length, 2, 'thirty-nine woken requests shared ONE execution');
  assert.equal(gateway._debug().waiting, 0, 'no resolver was left behind');
  assert.equal(gateway._debug().sessions, 0, 'and no empty Set either');
});

test('nothing that asks the client to slow down is ever held', async () => {
  // 1. A stale copy: the upstream is down, so waiting on it proves nothing.
  const stale = harness({ [SESSION]: [row({ status: 'approved', examMinutes: 50 })] });
  await poll(stale.gateway, approvalPoll('900000001'));
  stale.state.mode = 'error500';
  stale.state.clock += 50000;   // past HELD_REREAD_MS, so even a re-arm refreshes
  const onStale = await timed(stale.state,
    poll(stale.gateway, approvalPoll('900000001', { wait: 25, fp: 'a:approved:off:50' })));
  assert.deepEqual(onStale.body, {
    status: 'ok', approval: 'approved', audioMode: 'off', examMinutes: 50, stale: true,
    fp: 'a:approved:off:50', held: 0
  }, 'the fingerprint is the answer inside it, and it comes back at once');
  assert.ok(onStale.timers <= 1, 'it never held - at most the first look\'s own bound');

  // 2. No snapshot at all.
  const dead = harness({ [SESSION]: [row()] });
  dead.state.mode = 'error500';
  const onDead = await timed(dead.state, poll(dead.gateway, approvalPoll('900000001', { wait: 25, fp: 'x:up' })));
  assert.deepEqual(onDead.body, {
    status: 'error', code: 'upstream_unavailable', retryable: true,
    message: 'השרת עמוס — ננסה שוב אוטומטית', fp: 'x:up', held: 0
  });
  assert.ok(onDead.timers <= 1);

  // 3. A token mismatch never becomes anything else by waiting.
  const token = harness({ [SESSION]: [row({ status: 'approved', tokenHash: sha256Hex('real-token') })] });
  const onToken = await timed(token.state, poll(token.gateway,
    approvalPoll('900000001', { wait: 25, fp: 'a:tok', examineeToken: 'stolen-token' })));
  assert.equal(onToken.body.examineeTokenError, 'mismatch');
  assert.equal(onToken.body.held, 0);
  assert.ok(onToken.timers <= 1);

  // 4. A bad request is a bug in the caller, not something to wait out.
  const bad = harness({ [SESSION]: [row()] });
  const onBad = await timed(bad.state, poll(bad.gateway,
    { kind: 'approval', sessionCode: 'nope', idNumber: '900000001', wait: 25, fp: 'x:sess' }));
  assert.equal(onBad.res.status, 400);
  assert.equal(onBad.body.fp, 'x:sess');
  assert.equal(onBad.timers, 0, 'a 400 is refused before anything can sleep');
  assert.equal(bad.state.calls.length, 0, 'and it never reached Apps Script');
});

test('a hold is refused when the client is already behind, and `wait` is clamped', async () => {
  const { state, gateway } = harness({ [SESSION]: [row({ status: 'approved', examMinutes: 50 })] });

  // The answer already differs from what the client holds: nothing to wait for.
  const behind = await timed(state, poll(gateway, approvalPoll('900000001', { wait: 25, fp: WAITING_FP })));
  assert.equal(behind.body.fp, 'a:approved:off:50');
  assert.equal(behind.body.held, 0);
  assert.ok(behind.timers <= 1, 'nothing to wait for - at most the first look\'s own bound');

  // wait=0, a fraction, a word and a negative all mean "answer now", exactly
  // like the client that never heard of holding.
  for (const wait of [0, '0', 2.5, 'abc', -5, '']) {
    const { body, timers } = await timed(state,
      poll(gateway, approvalPoll('900000001', { wait: wait, fp: 'a:approved:off:50' })));
    assert.equal(body.held, 0, 'wait=' + JSON.stringify(wait));
    assert.equal(timers, 0, 'wait=' + JSON.stringify(wait) + ' never slept');
  }

  // 26 is clamped to 25, and that is also the CPU guard: one evaluation per
  // second and not one more, whatever wakes it.
  const { body, timers } = await timed(state, poll(gateway, approvalPoll('900000001', { wait: 26, fp: 'a:approved:off:50' })));
  assert.equal(body.held, 25000, 'clamped to the 25 s ceiling');
  assert.ok(timers <= 26 * 2, 'at most 26 evaluations for a full hold, got ' + timers / 2);
  assert.equal(timers, 50, '25 evaluations, each parking a tick and a grace timer');
  assert.equal(gateway._debug().waiting, 0);
});

// --- the hold never outlives its own deadline ------------------------------

test('a slow upstream started late in the hold does not stretch it past the deadline', async () => {
  const { state, gateway } = harness({ [SESSION]: [row({ status: 'waiting' })] });
  await poll(gateway, approvalPoll('900000001'));
  assert.equal(state.calls.length, 1, 'a fresh snapshot to hold on');

  // Parked one second short of the safety cadence, so the copy expires with a
  // second of the hold still to run: that look opens an upstream read which
  // Google answers only ten seconds later. This is the shape seen live, where
  // a read started at second 24 of a 25 s hold made it 44 s long.
  await advance(state, 19000);
  state.upstreamDelayMs = 10000;

  const { body } = await settle(state,
    poll(gateway, approvalPoll('900000001', { wait: 3, fp: WAITING_FP })));

  assert.equal(state.calls.length, 2, 'the look really did open the read');
  assert.deepEqual(body, {
    status: 'ok', approval: 'waiting', audioMode: 'off', fp: WAITING_FP, held: 4000
  }, 'the unchanged answer, so the client simply polls again');
  assert.ok(body.held <= 3000 + 1000, 'the deadline plus one second of grace, got ' + body.held);
  assert.equal(state.clock, CLOCK0 + 23000,
    'answered at its own deadline - NOT ten seconds later, when Google finally replied');
  assert.equal(gateway._debug().waiting, 0, 'and it took its resolver back out');

  // Let the abandoned read finish, so no real abort timer outlives the test.
  await advance(state, 20000);
});

test('the abandoned look still lands, and the next poll is served from it', async () => {
  const { state, gateway } = harness({ [SESSION]: [row({ status: 'waiting' })] });
  await poll(gateway, approvalPoll('900000001'));
  await advance(state, 19000);
  state.upstreamDelayMs = 10000;

  const pending = [];
  const ctx = { waitUntil: promise => pending.push(promise) };
  const held = await settle(state,
    poll(gateway, approvalPoll('900000001', { wait: 3, fp: WAITING_FP }), undefined, ctx));

  assert.equal(held.body.held, 4000);
  assert.equal(held.body.fp, WAITING_FP, 'the client keeps the answer it had');
  assert.equal(pending.length, 1, 'the look was handed to waitUntil, not dropped with the response');

  // While Google was answering, the examiner approved: the read this hold
  // walked away from is the one carrying the decision.
  state.snapshots[SESSION] = [row({ status: 'approved', examMinutes: 50 })];
  await settle(state, Promise.all(pending));
  assert.equal(state.clock, CLOCK0 + 31000, 'it finished ten seconds after it started');
  assert.equal(state.calls.length, 2, 'and it was that same read, not a new one');

  // Half a second later the examinee polls again. Memory holds what the
  // abandoned look put there, so Apps Script is not asked a second time.
  await advance(state, 500);
  const after = await poll(gateway, approvalPoll('900000001', { wait: 25, fp: WAITING_FP }));
  assert.deepEqual(after.body, {
    status: 'ok', approval: 'approved', audioMode: 'off', examMinutes: 50,
    fp: 'a:approved:off:50', held: 0
  });
  assert.equal(state.calls.length, 2, 'the abandoned read paid for this answer');
  assert.equal(gateway._debug().waiting, 0);
});

// --- the examiner's watch: GET /v1/session/watch (r31, DESIGN §13.2) -------
// The dashboard used to re-read the whole session from Apps Script every 5 s
// whether or not anything had happened. Now it holds ONE request here and asks
// "did this session change?" — and reads Apps Script only when the answer is
// yes. The watch itself costs Google nothing extra: it fingerprints the very
// snapshot the examinees' own polls keep fresh.

const watchParams = over => Object.assign({ sessionCode: SESSION, grant: examinerGrant() }, over);

async function watch(gateway, params, ctx) {
  const url = 'https://session-gateway.test/v1/session/watch?' + new URLSearchParams(params).toString();
  const res = await gateway(new Request(url), ctx);
  return { res, body: await res.json() };
}

/** The fingerprint a session of exactly these rows produces, in isolation. */
async function sessionFp(rows) {
  const { gateway } = harness({ [SESSION]: rows });
  return (await watch(gateway, watchParams())).body.fp;
}

test('watch: only a valid examiner grant opens it, and a bad session code is 400', async () => {
  const { state, gateway } = harness({ [SESSION]: [row({ status: 'waiting' })] });

  for (const [name, bad] of [
    ['none', ''],
    ['exam scope', examGrant([14])],
    ['forged', grant({ s: 'examiner', sub: 'ex:7' }, 'not-the-key')],
    ['expired', examinerGrant({ exp: CLOCK0 - 1 })],
    ['rubbish', 'bogus']
  ]) {
    const { res, body } = await watch(gateway, { sessionCode: SESSION, grant: bad });
    assert.equal(res.status, 403, name);
    assert.deepEqual(body, { status: 'error', code: 'grant_invalid' }, name);
  }
  assert.equal(state.calls.length, 0, 'a refused watch never reaches Apps Script');

  const badSession = await watch(gateway, { sessionCode: 'nope', grant: examinerGrant() });
  assert.equal(badSession.res.status, 400);
  assert.equal(badSession.body.status, 'error');
  assert.equal(state.calls.length, 0);

  const ok = await watch(gateway, watchParams());
  assert.equal(ok.res.status, 200);
  assert.equal(ok.res.headers.get('Cache-Control'), 'no-store');
  assert.equal(ok.res.headers.get('Access-Control-Allow-Origin'), PAGES_ORIGIN);
  assert.equal(ok.body.status, 'ok');
  assert.equal(ok.body.rows, 1, 'how many rows, never the rows themselves');
  assert.equal(ok.body.at, CLOCK0);
  assert.equal(ok.body.held, 0);
  assert.match(ok.body.fp, /^s:[0-9a-f]{12}$/);
});

test('watch: the session fingerprint moves on every field the dashboard shows, and on nothing else', async () => {
  const base = [
    row({ id: '900000001', status: 'in_exam', audio: 'on', examMinutes: 50, extraMinutes: 5,
          warn: 1, fin: 0, ext: 0, dq: '' }),
    row({ id: '900000002', status: 'waiting' })
  ];
  const baseline = await sessionFp(base);
  assert.match(baseline, /^s:[0-9a-f]{12}$/);

  // The examiner never sees the token hash, and it cannot change inside an
  // attempt: hashing it would wake the dashboard for nothing.
  const tokened = base.slice();
  tokened[0] = Object.assign({}, base[0], { tokenHash: sha256Hex('real-token') });
  assert.equal(await sessionFp(tokened), baseline, 'tokenHash is not watched');

  // Every field r31's sessionSnapshot carries is watched — this is the whole
  // list, and a change to any one of them must reach the screen.
  for (const [field, value] of [
    ['status', 'completed'], ['audio', 'off'], ['examMinutes', 60], ['extraMinutes', 6],
    ['warn', 2], ['fin', 1], ['ext', 1], ['dq', 1], ['id', '900000007']
  ]) {
    const changed = base.slice();
    changed[0] = Object.assign({}, base[0], { [field]: value });
    assert.notEqual(await sessionFp(changed), baseline, field + ' must move the fingerprint');
  }

  assert.notEqual(await sessionFp(base.concat([row({ id: '900000003', status: 'waiting' })])), baseline,
    'a new registration must move it');
  assert.notEqual(await sessionFp([base[1], base[0]]), baseline, 'and so must the row order');
  assert.equal(await sessionFp([]), 's:none', 'a session with no rows has its own fingerprint');
});

test('watch: a re-read of unchanged rows answers the same fingerprint', async () => {
  const { state, gateway } = harness({ [SESSION]: [row({ status: 'waiting' })] });
  const first = await watch(gateway, watchParams());

  state.clock += 5000;
  const second = await watch(gateway, watchParams());
  assert.equal(state.calls.length, 2, 'that really was a second read');
  assert.ok(second.body.at > first.body.at, 'of a newer snapshot');
  assert.equal(second.body.fp, first.body.fp,
    '`at` is not in the fingerprint — otherwise the dashboard would re-read every 2 s for ever');
});

test("watch: an older server's rows (no warn/fin/ext/dq) hash exactly like empty ones", async () => {
  // A Worker deployed ahead of the server (§13.8: the Worker goes out at
  // night, the paste is Yossi's in the morning) must be stable, not noisy.
  const legacy = { id: '900000001', status: 'waiting', tokenHash: '', audio: 'off', examMinutes: 40, extraMinutes: 0 };
  const fp = await sessionFp([legacy]);
  assert.equal(await sessionFp([Object.assign({}, legacy, { warn: '', fin: '', ext: '', dq: '' })]), fp);
  assert.equal(await sessionFp([Object.assign({}, legacy, { warn: undefined, fin: null, ext: '', dq: undefined })]), fp);
  assert.notEqual(await sessionFp([Object.assign({}, legacy, { fin: 1 })]), fp,
    'and the moment the new server sends one, it counts');
});

test('watch: a hold runs out its `wait` and answers the same fingerprint', async () => {
  const { state, gateway } = harness({ [SESSION]: [row({ status: 'waiting' })] });
  const first = await watch(gateway, watchParams());
  assert.equal(first.body.held, 0, 'no fp to send yet, so it is answered at once');
  assert.equal(state.calls.length, 1);

  const { body, timers } = await timed(state, watch(gateway, watchParams({ wait: 5, fp: first.body.fp })));

  assert.equal(body.fp, first.body.fp);
  assert.equal(body.held, 5000);
  assert.equal(body.rows, 1);
  assert.equal(state.clock, CLOCK0 + 5000, 'it really waited the five seconds');
  assert.equal(timers, 10, 'five evaluations, one per second, not a spin');
  assert.equal(state.calls.length, 1, 'and five seconds of watching cost Google nothing at all');
  assert.equal(gateway._debug().waiting, 0, 'and it took its resolver back out');
});

test('watch: an examiner nudge wakes it at once with a new fingerprint', async () => {
  const { state, gateway } = harness({ [SESSION]: [row({ status: 'waiting' })] });
  const first = await watch(gateway, watchParams());

  const holding = watch(gateway, watchParams({ wait: 25, fp: first.body.fp }));
  assert.equal(await parked(gateway, 1), 1, 'parked');

  // The examiner approved in another tab: the decision-carrying nudge patches
  // the snapshot, which is exactly what the dashboard is watching.
  assert.equal((await nudge(gateway, '&idNumber=900000001&status=approved&examMinutes=50&audio=on')).body.patched, true);

  const { body } = await settle(state, holding);
  assert.notEqual(body.fp, first.body.fp, 'the dashboard now knows to re-read examinerDashboard');
  assert.equal(body.rows, 1);
  assert.equal(body.held, 0);
  assert.equal(state.clock, CLOCK0, 'without waiting out a single tick');
  assert.equal(state.calls.length, 1, 'and without an upstream read');
  assert.equal(gateway._debug().waiting, 0);
});

test('watch: a drop wakes it, and the rows the server now has move the fingerprint', async () => {
  const { state, gateway } = harness({ [SESSION]: [row({ status: 'waiting' })] });
  const first = await watch(gateway, watchParams());

  const holding = watch(gateway, watchParams({ wait: 25, fp: first.body.fp }));
  await parked(gateway, 1);

  // A second examinee registered; the nudge says only "the snapshot is behind".
  state.snapshots[SESSION] = [row({ status: 'waiting' }), row({ id: '900000002', status: 'waiting' })];
  await nudge(gateway, '');

  const { body } = await settle(state, holding);
  assert.equal(body.rows, 2);
  assert.notEqual(body.fp, first.body.fp);
  assert.equal(body.held, 0);
  assert.equal(state.calls.length, 2, 'the woken watch read the truth once');
  assert.equal(gateway._debug().waiting, 0);
});

test('watch: an upstream read that changes nothing keeps it holding to the deadline', async () => {
  const { state, gateway } = harness({ [SESSION]: [row({ status: 'waiting' })] });
  const first = await watch(gateway, watchParams());

  // Parked one second short of the safety cadence, so the copy expires mid
  // hold and Google is read again — a NEW snapshot, with a new `at` and a
  // token hash that was not there before. Neither is in the fingerprint, so
  // the dashboard must not be woken by it.
  await advance(state, 19000);
  state.snapshots[SESSION] = [row({ status: 'waiting', tokenHash: sha256Hex('real-token') })];
  const { body } = await settle(state, watch(gateway, watchParams({ wait: 6, fp: first.body.fp })));

  assert.equal(state.calls.length, 2, 'the hold really did re-read Google — exactly once');
  assert.equal(body.fp, first.body.fp, 'same fingerprint: nothing the examiner sees changed');
  assert.equal(body.held, 6000, 'so it held all the way to its own deadline');
  assert.ok(body.at > CLOCK0, 'even though it answers from a newer snapshot');
  assert.equal(gateway._debug().waiting, 0);
});

test('watch: a full hold is clamped to 25 s and costs at most HOLD_MAX_STEPS evaluations', async () => {
  const { state, gateway } = harness({ [SESSION]: [row({ status: 'waiting' })] });
  const first = await watch(gateway, watchParams());

  const { body, timers } = await timed(state, watch(gateway, watchParams({ wait: 26, fp: first.body.fp })));
  assert.equal(body.held, 25000, 'clamped to the 25 s ceiling');
  assert.ok(timers <= 26 * 2, 'at most 26 evaluations, got ' + timers / 2);
  assert.equal(timers, 50, '25 evaluations, each parking a tick and a grace timer');
  assert.equal(gateway._debug().waiting, 0);

  // wait=0, a fraction, a word and a negative all mean "answer now".
  for (const wait of [0, '0', 2.5, 'abc', -5, '']) {
    const now = await timed(state, watch(gateway, watchParams({ wait: wait, fp: first.body.fp })));
    assert.equal(now.body.held, 0, 'wait=' + JSON.stringify(wait));
    assert.equal(now.timers, 0, 'wait=' + JSON.stringify(wait) + ' never slept');
  }
});

test('watch: a stale copy and a dead upstream are never held', async () => {
  const stale = harness({ [SESSION]: [row({ status: 'waiting' })] });
  const first = await watch(stale.gateway, watchParams());
  stale.state.mode = 'error500';
  stale.state.clock += 50000;   // past HELD_REREAD_MS, so even a re-arm refreshes

  const onStale = await timed(stale.state, watch(stale.gateway, watchParams({ wait: 25, fp: first.body.fp })));
  assert.equal(onStale.res.status, 200);
  assert.equal(onStale.body.stale, true);
  assert.equal(onStale.body.fp, first.body.fp, 'the fingerprint of the copy inside it');
  assert.equal(onStale.body.held, 0, 'the dashboard falls back to its own safety net instead');
  assert.ok(onStale.timers <= 1, 'it never held - at most the first look\'s own bound');

  const dead = harness({ [SESSION]: [row()] });
  dead.state.mode = 'error500';
  const onDead = await timed(dead.state, watch(dead.gateway, watchParams({ wait: 25, fp: 'x:up' })));
  assert.equal(onDead.res.status, 200);
  assert.deepEqual(onDead.body, {
    status: 'error', code: 'upstream_unavailable', retryable: true,
    message: 'השרת עמוס — ננסה שוב אוטומטית', fp: 'x:up', held: 0
  });
  assert.ok(onDead.timers <= 1);
});

test('watch: an empty session is held, and the first registration wakes it', async () => {
  const { state, gateway } = harness({ [SESSION]: [] });
  const empty = await watch(gateway, watchParams());
  assert.equal(empty.body.fp, 's:none');
  assert.equal(empty.body.rows, 0);
  assert.equal(state.calls.length, 1);

  const holding = watch(gateway, watchParams({ wait: 25, fp: 's:none' }));
  assert.equal(await parked(gateway, 1), 1,
    's:none is exactly the wait that matters before a class starts');

  // The first examinee registers. Since 22/09 the Worker does NOT ask Google
  // every two seconds whether that happened: the examinee's own first poll is
  // a FRESH chain (no fp), so it reads the truth — and that read is what wakes
  // the dashboard, for no execution of its own.
  state.snapshots[SESSION] = [row({ status: 'waiting' })];
  const device = await poll(gateway, approvalPoll('900000001'));
  assert.equal(device.body.approval, 'waiting');
  assert.equal(state.calls.length, 2, 'one read, shared by the device and the dashboard');

  const { body } = await settle(state, holding);
  assert.equal(body.rows, 1);
  assert.notEqual(body.fp, 's:none');
  assert.equal(body.held, 0, 'the dashboard saw it the moment that read landed');
  assert.equal(state.clock, CLOCK0, 'and not one tick was waited out');
  assert.equal(gateway._debug().waiting, 0);
});

test('forty held polls and one held watch cost nothing while nothing happens', async () => {
  const ids = Array.from({ length: 40 }, (_, i) => String(900000001 + i));
  const { state, gateway } = harness({ [SESSION]: ids.map(id => row({ id, status: 'waiting' })) });

  await Promise.all(ids.map(id => poll(gateway, approvalPoll(id))));
  const seen = await watch(gateway, watchParams());
  assert.equal(state.calls.length, 1, 'the watch rode on the snapshot the polls had already read');

  const holds = ids.map(id => poll(gateway, approvalPoll(id, { wait: 25, fp: WAITING_FP })));
  const watching = watch(gateway, watchParams({ wait: 25, fp: seen.body.fp }));
  assert.equal(await parked(gateway, 41), 41, 'forty examinees and one dashboard, one session');

  await advance(state, 6000);
  assert.equal(state.calls.length, 1,
    'six seconds of holding, forty-one open requests, and not one Apps Script execution');
  assert.equal(gateway._debug().waiting, 41, 'still holding — nothing anyone cares about changed');

  // The examiner approves the whole class: the dashboard and every examinee
  // learn it from the SAME execution.
  state.snapshots[SESSION] = ids.map(id => row({ id, status: 'approved', examMinutes: 40 }));
  await nudge(gateway, '');
  const answers = await settle(state, Promise.all(holds.concat([watching])));

  assert.equal(state.calls.length, 2, 'forty-one woken requests shared ONE execution');
  for (const answer of answers.slice(0, 40)) assert.equal(answer.body.approval, 'approved');
  const dashboard = answers[40].body;
  assert.equal(dashboard.rows, 40);
  assert.notEqual(dashboard.fp, seen.body.fp, 'and the dashboard reads examinerDashboard once, now');
  assert.equal(gateway._debug().waiting, 0, 'no resolver was left behind');
  assert.equal(gateway._debug().sessions, 0);
});

// --- the examinee's own nudge (r31, DESIGN §13.5) --------------------------
// Until r31 only the examiner pushed, so a submit reached the dashboard only
// on the Worker's next read of Google (≤2 s) and then on the dashboard's next
// tick (≤5 s). The device now says "look again" itself — authenticated by its
// own token, and never carrying a decision.

const TOKEN = 'examinee-token-7';
const devicePush = (gateway, query) =>
  call(gateway, '/v1/invalidate?sessionCode=' + SESSION + query, { method: 'POST' });

test('the examinee pushes their own submit: the token matches, the snapshot is dropped', async () => {
  const { state, gateway } = harness({ [SESSION]: [
    row({ status: 'in_exam', tokenHash: sha256Hex(TOKEN) })
  ] });
  const seen = await watch(gateway, watchParams());
  assert.equal(state.calls.length, 1);

  const holding = watch(gateway, watchParams({ wait: 25, fp: seen.body.fp }));
  const polling = poll(gateway, statusPoll('900000001', { wait: 25, fp: 's:in_exam:0', examineeToken: TOKEN }));
  assert.equal(await parked(gateway, 2), 2, 'the dashboard and the device itself');

  // The result POST answered ok — the row is already 'completed' in the sheet.
  state.snapshots[SESSION] = [row({ status: 'completed', tokenHash: sha256Hex(TOKEN) })];
  const pushed = await devicePush(gateway, '&idNumber=900000001&examineeToken=' + TOKEN);
  assert.equal(pushed.res.status, 200);
  assert.deepEqual(pushed.body, { status: 'ok', dropped: true });

  const dash = await settle(state, holding);
  assert.notEqual(dash.body.fp, seen.body.fp, 'the dashboard sees the submit at once');
  assert.equal(dash.body.held, 0);
  const device = await settle(state, polling);
  assert.equal(device.body.examStatus, 'completed');
  assert.equal(state.calls.length, 2, 'the two woken requests shared ONE re-read');
  assert.equal(state.clock, CLOCK0, 'and nothing waited out a tick');
  assert.equal(gateway._debug().waiting, 0);
});

test('an examinee nudge with the wrong token, no token or no row of its own is refused', async () => {
  const { state, gateway } = harness({ [SESSION]: [
    row({ id: '900000001', status: 'in_exam', tokenHash: sha256Hex(TOKEN) }),
    row({ id: '900000002', status: 'in_exam' })            // an older row: no token stored
  ] });
  await poll(gateway, statusPoll('900000001'));
  assert.equal(state.calls.length, 1);

  for (const [name, query] of [
    ['a stolen token', '&idNumber=900000001&examineeToken=not-the-token'],
    ["a classmate's row with no stored hash", '&idNumber=900000002&examineeToken=' + TOKEN],
    ['an id with no row at all', '&idNumber=900000009&examineeToken=' + TOKEN],
    ['no token at all', '&idNumber=900000001'],
    ['no id at all', '&examineeToken=' + TOKEN]
  ]) {
    const { res, body } = await devicePush(gateway, query);
    assert.equal(res.status, 403, name);
    assert.deepEqual(body, { status: 'error', code: 'grant_invalid' }, name);
  }

  state.snapshots[SESSION] = [row({ status: 'completed' })];
  const after = await poll(gateway, statusPoll('900000001'));
  assert.equal(state.calls.length, 1, 'a refused nudge never dropped anything, and never read upstream');
  assert.equal(after.body.examStatus, 'in_exam');
});

test('an examinee may say "look again" but never what to look at', async () => {
  const { state, gateway } = harness({ [SESSION]: [row({ status: 'in_exam', tokenHash: sha256Hex(TOKEN) })] });
  await poll(gateway, statusPoll('900000001'));

  // The one thing this door must never open: a device writing its own verdict.
  // Carrying `status` takes the request straight to the examiner's door.
  const forged = await devicePush(gateway,
    '&idNumber=900000001&examineeToken=' + TOKEN + '&status=completed&extraMinutes=600');
  assert.equal(forged.res.status, 403);
  assert.deepEqual(forged.body, { status: 'error', code: 'grant_invalid' });

  const after = await poll(gateway, statusPoll('900000001'));
  assert.deepEqual(after.body, { status: 'ok', examStatus: 'in_exam', extraMinutes: 0 }, 'nothing was patched');
  assert.equal(state.calls.length, 1, 'and nothing was dropped');

  const badSession = await call(gateway,
    '/v1/invalidate?sessionCode=nope&idNumber=900000001&examineeToken=' + TOKEN, { method: 'POST' });
  assert.equal(badSession.res.status, 400);
});

test('an examinee nudge for a session with no snapshot is ok, dropped:false, and free', async () => {
  const { state, gateway } = harness({ [SESSION]: [row({ status: 'in_exam', tokenHash: sha256Hex(TOKEN) })] });

  const { res, body } = await devicePush(gateway, '&idNumber=900000001&examineeToken=' + TOKEN);
  assert.equal(res.status, 200);
  assert.deepEqual(body, { status: 'ok', dropped: false });
  assert.equal(state.calls.length, 0, 'nothing to drop is not a reason to read Apps Script');
});

test('the examinee nudge spends the same re-read budget as the examiner', async () => {
  const { state, gateway } = harness({ [SESSION]: [row({ status: 'in_exam', tokenHash: sha256Hex(TOKEN) })] });
  await poll(gateway, statusPoll('900000001'));
  assert.equal(state.calls.length, 1);

  const query = '&idNumber=900000001&examineeToken=' + TOKEN;
  assert.deepEqual((await devicePush(gateway, query)).body, { status: 'ok', dropped: true });

  // submitResult and markFinished push within a couple of seconds of each
  // other, and a flaky device may fire either twice: the second must not buy a
  // second Apps Script execution, and neither must the examiner's own drop.
  state.clock += 500;
  assert.deepEqual((await devicePush(gateway, query)).body, { status: 'ok', dropped: false });
  assert.deepEqual((await nudge(gateway, '')).body, { status: 'ok' });

  state.snapshots[SESSION] = [row({ status: 'completed', tokenHash: sha256Hex(TOKEN) })];
  const after = await poll(gateway, statusPoll('900000001'));
  assert.equal(after.body.examStatus, 'completed');
  assert.equal(state.calls.length, 2, 'the first drop is the only one that was paid for');

  state.clock += 2000;
  assert.deepEqual((await devicePush(gateway, query)).body, { status: 'ok', dropped: true },
    'past the gap it works again');
});

// --- one budget per request (22/09, from `wrangler tail`) ------------------
// The hold was bounded, the FIRST LOOK was not: a request that joined an
// upstream read Google took 20 s over started its 25 s hold 20 s late and
// answered after 45 s — past the client's own 40 s abort, so four watches in
// one 24-minute tail were counted as communication failures. First look and
// hold now share ONE budget, measured from the moment the request arrived.

test('a first look Google never answers still answers inside the request budget', async () => {
  const { state, gateway } = harness({ [SESSION]: [row({ status: 'waiting' })] });
  await poll(gateway, approvalPoll('900000001'));     // a copy to fall back on
  await advance(state, 21000);                        // past HELD_REREAD_MS: even a re-arm must read
  state.upstreamDelayMs = 30000;                      // ...and Google has stopped answering
  const startedAt = state.clock;

  const pending = [];
  const { body } = await settle(state,
    poll(gateway, approvalPoll('900000001', { wait: 25, fp: WAITING_FP }), undefined,
      { waitUntil: p => pending.push(p) }));

  assert.equal(state.calls.length, 2, 'the first look really did open the read');
  assert.equal(state.clock - startedAt, 26000, 'the wait plus one second of grace, and not a second more');
  assert.ok(state.clock - startedAt <= 25000 + 1000,
    'a wait=25 request must never reach the client deadline of 40 s');
  assert.equal(body.stale, true, 'answered from the copy we had, marked so');
  assert.equal(body.held, 0, 'and a stale answer is never held');
  assert.equal(body.fp, WAITING_FP);
  assert.equal(pending.length, 1, 'the read was abandoned to waitUntil, not dropped');
  assert.equal(gateway._debug().waiting, 0);

  await advance(state, 40000);   // let the abandoned read land, so nothing dangles
});

test('a slow first look eats into the hold, never past it', async () => {
  const { state, gateway } = harness({ [SESSION]: [row({ status: 'waiting' })] });
  await poll(gateway, approvalPoll('900000001'));
  await advance(state, 21000);
  state.upstreamDelayMs = 10000;                      // a bad Apps Script morning
  const startedAt = state.clock;

  const { body } = await settle(state, poll(gateway, approvalPoll('900000001', { wait: 25, fp: WAITING_FP })));

  assert.equal(state.clock - startedAt, 25000, 'ten seconds of first look plus fifteen of holding');
  assert.equal(body.held, 25000, '`held` is the whole request, which is what the client times');
  assert.equal(body.fp, WAITING_FP, 'nothing changed, so the client simply asks again');
  assert.equal(gateway._debug().waiting, 0);
  await advance(state, 40000);
});

test('a watch inherits the same one-budget rule, and falls back to the copy it has', async () => {
  const { state, gateway } = harness({ [SESSION]: [row({ status: 'waiting' })] });
  const first = await watch(gateway, watchParams());
  await advance(state, 21000);
  state.upstreamDelayMs = 30000;
  const startedAt = state.clock;

  const { body } = await settle(state, watch(gateway, watchParams({ wait: 10, fp: first.body.fp })));
  assert.equal(state.clock - startedAt, 11000, 'wait + grace, whatever Google is doing');
  assert.equal(body.stale, true, 'the copy is 32 s old — still inside STALE_MS, so it is served');
  assert.equal(body.fp, first.body.fp);
  assert.equal(body.held, 0);
  await advance(state, 40000);
});

test('a request with no `wait` still bounds its own first look', async () => {
  const { state, gateway } = harness({ [SESSION]: [row({ status: 'waiting' })] });
  await poll(gateway, approvalPoll('900000001'));
  await advance(state, 3000);
  state.upstreamDelayMs = 60000;
  const startedAt = state.clock;

  // No wait, no fp: the answer must stay byte-identical to the server's, and
  // the request must not sit there for UPSTREAM_TIMEOUT_MS either.
  const { body } = await settle(state, poll(gateway, approvalPoll('900000001')));
  assert.equal(state.clock - startedAt, 21000, 'FIRST_LOOK_MAX_MS plus the grace');
  assert.deepEqual(body, { status: 'ok', approval: 'waiting', audioMode: 'off', stale: true },
    'and still no fp/held for a client that asked for neither');
  await advance(state, 70000);
});

test("a request that joins another request's upstream read answers on its own timer", async () => {
  // Cloudflare cancels a request that awaits a promise ANOTHER request created
  // while having no I/O or timer of its own: «the Workers runtime canceled
  // this request because it detected that your Worker's code had hung» — 15 of
  // 74 requests, HTTP 500, 2-3 ms of wall time, 0 CPU (22/09). The join must
  // always race a timer this request owns.
  const { state, gateway } = harness({ [SESSION]: [row({ status: 'waiting' })] });
  await poll(gateway, approvalPoll('900000001'));
  await advance(state, 21000);
  state.upstreamDelayMs = 30000;

  // A opens the read, runs out of budget and walks away, leaving it in flight.
  const pending = [];
  await settle(state, poll(gateway, approvalPoll('900000001', { wait: 3, fp: WAITING_FP }), undefined,
    { waitUntil: p => pending.push(p) }));
  assert.equal(state.calls.length, 2, 'A opened it');
  assert.equal(pending.length, 1, 'and it is still running under waitUntil');

  // B arrives while that read is in flight and joins it — a foreign promise.
  const startedAt = state.clock;
  const b = await timed(state, poll(gateway, approvalPoll('900000001', { wait: 2, fp: '' })));
  assert.equal(state.calls.length, 2, "B joined A's read instead of opening a second one");
  assert.equal(state.clock - startedAt, 3000, 'and answered on its own bound, not on Google');
  assert.ok(b.timers >= 1, 'the timer it owns is what keeps the runtime from cancelling it');
  assert.equal(b.body.stale, true);
  assert.equal(b.body.held, 0);
  await advance(state, 40000);
});

// --- the fingerprint is stamped once, where the snapshot is stored ---------

test('a held watch hashes the session once per upstream read, never once per tick', async () => {
  // 7-12 ms of CPU per held watch (22/09) against a free-plan ceiling of 10:
  // `caches.default` hands back a NEW object every tick, so the WeakMap could
  // never hit and every second paid for another SHA-256. `sfp` travels inside
  // the cached JSON, so a tick costs nothing.
  const { state, spawn } = harness({ [SESSION]: [row({ status: 'waiting' })] }, null, true);
  const gateway = spawn();
  const first = await watch(gateway, watchParams());
  assert.match(JSON.parse([...state.store.values()][0].body).sfp, /^s:[0-9a-f]{12}$/,
    'the copy in caches.default carries the fingerprint it was stamped with');

  await advance(state, 19000);   // so the safety reads fall inside the hold
  const readsBefore = state.calls.length;
  const digestsBefore = subtleCalls.digest;
  const { body } = await settle(state, watch(gateway, watchParams({ wait: 25, fp: first.body.fp })));
  const reads = state.calls.length - readsBefore;
  const digests = subtleCalls.digest - digestsBefore;

  assert.equal(body.held, 25000);
  assert.equal(reads, 2, 'two safety reads in twenty-five seconds, got ' + reads);
  assert.equal(digests, reads, 'one hash per read — NOT one per tick (that was 25)');
});

test('a cache entry written before r31 has no sfp and is hashed as before', async () => {
  const { state, spawn } = harness({ [SESSION]: [row({ status: 'waiting' })] }, null, true);
  const gateway = spawn();
  const withSfp = await watch(gateway, watchParams());

  // Exactly what the previous deploy left in caches.default: the same rows,
  // no `sfp`. It must produce the very same fingerprint.
  const key = 'https://session-gateway.internal/snap/' + SESSION;
  const stored = JSON.parse(state.store.get(key).body);
  delete stored.sfp;
  state.store.set(key, { body: JSON.stringify(stored), maxAge: 2, at: state.clock });

  const cold = spawn();   // no memory: it must read that cache entry
  const { body } = await watch(cold, watchParams());
  assert.equal(body.fp, withSfp.body.fp);
});

// --- the flip-flop of 22/09 ------------------------------------------------

test('an old-shape snapshot and a new-shape one hash alike, so nothing flip-flops', async () => {
  // Three ~2 s bursts, ~every 60-80 s: the same session, no writes, and the
  // answers alternating between two fingerprints 250 ms apart. Two snapshots
  // of DIFFERENT SHAPE were alive at once — r30's, with no warn/fin/ext/dq,
  // beside r31's carrying 0 — and `lookAtNewest` answers whichever copy
  // DIFFERS from what the client holds, so memory and cache took turns.
  const OLD = { id: '900000001', status: 'in_exam', tokenHash: '', audio: 'off', examMinutes: 40, extraMinutes: 0 };
  const NEW = Object.assign({}, OLD, { warn: 0, fin: 0, ext: 0, dq: 0 });
  assert.equal(await sessionFp([NEW]), await sessionFp([OLD]), 'absent and zero are the same state');
  assert.equal(await sessionFp([Object.assign({}, OLD, { warn: '0', fin: false, ext: '', dq: null })]),
    await sessionFp([OLD]), 'and so are the shapes a sheet can produce for "nothing"');
  assert.notEqual(await sessionFp([Object.assign({}, OLD, { warn: 1 })]), await sessionFp([OLD]),
    'a real warning still moves it');

  // The live shape: an isolate holding the OLD snapshot in memory while
  // another isolate has written the NEW one into caches.default.
  const { state, spawn } = harness({ [SESSION]: [OLD] }, null, true);
  const gateway = spawn();
  const seen = await watch(gateway, watchParams());

  const key = 'https://session-gateway.internal/snap/' + SESSION;
  state.store.set(key, {
    body: JSON.stringify({ at: state.clock, rows: [NEW] }), maxAge: 2, at: state.clock
  });

  const holding = watch(gateway, watchParams({ wait: 25, fp: seen.body.fp }));
  assert.equal(await parked(gateway, 1), 1,
    'it is HOLDING — not bouncing between two copies of the same unchanged session');

  // ...and a real change still ends it at once.
  await nudge(gateway, '&idNumber=900000001&status=completed');
  const { body } = await settle(state, holding);
  assert.notEqual(body.fp, seen.body.fp);
  assert.equal(body.held, 0);
});

// --- one line per answered request, for the next `wrangler tail` -----------

test('every answered watch and poll leaves one structured line naming where the view came from', async () => {
  const { state, gateway } = harness({ [SESSION]: [row({ status: 'waiting' })] });

  const first = await watch(gateway, watchParams());
  assert.deepEqual(state.logs.at(-1), {
    r: 'watch', s: SESSION, src: 'fetch', fp: first.body.fp, held: 0, fl: 0, age: 0, late: 0,
    up: 0, usfp: first.body.fp, rows: 1, nr: 0
  }, 'src says which copy answered, age how old it was, usfp what Google last gave us');
  assert.equal(state.logs.at(-1).sess, undefined,
    'an r31 server sends no results and no `v`, so there is no session payload to announce');

  await poll(gateway, approvalPoll('900000001', { wait: 0 }));
  const line = state.logs.at(-1);
  assert.equal(line.r, 'poll');
  assert.equal(line.k, 'approval');
  assert.equal(line.src, 'memory', 'served from the copy the watch had already read');
  assert.equal(line.fp, 'a:waiting:off:-');

  // A held request reports the hold and the first look separately.
  const held = await settle(state, watch(gateway, watchParams({ wait: 5, fp: first.body.fp })));
  const after = state.logs.at(-1);
  assert.equal(after.held, held.body.held);
  assert.equal(after.fl, 0, 'the first look was free — it was all hold');
  assert.equal(state.logs.filter(l => l.r === 'watch').length, 2, 'one line per request, no more');
});

// --- reading Google only when something happened (r31.2, 22/09 11:15) ------
// KNOWN_ISSUES #35: Google's response-delivery hop stalls 25-60 s for our
// projects while an idle project in the same account is untouched — and the
// busiest thing we send Google is this snapshot read, every 2 s per session.
// Every client write already announces itself (the examiner's patch/drop, the
// examinee's nudge after submit/markFinished/registration/DQ), so a chain that
// re-arms with an `fp` trusts the copy it already holds for HELD_REREAD_MS,
// and a FRESH chain still demands one younger than FRESH_MS.

test('a re-arm trusts a copy the fresh chain would refuse', async () => {
  const { state, gateway } = harness({ [SESSION]: [row({ status: 'waiting' })] });
  await poll(gateway, approvalPoll('900000001', { fp: '' }));
  assert.equal(state.calls.length, 1);

  state.clock += 15000;
  const rearm = await poll(gateway, approvalPoll('900000001', { fp: WAITING_FP }));
  assert.equal(state.calls.length, 1,
    'fifteen seconds old is fine for a chain that already holds exactly this answer');
  assert.equal(rearm.body.fp, WAITING_FP);
  assert.equal(rearm.body.stale, undefined, 'and it is NOT stale — nobody said anything changed');

  const fresh = await poll(gateway, approvalPoll('900000001'));
  assert.equal(state.calls.length, 2, 'a fresh chain still demands a copy younger than two seconds');
  assert.equal(fresh.body.approval, 'waiting');

  // A watch behaves the same way.
  const seen = await watch(gateway, watchParams());
  state.clock += 15000;
  await watch(gateway, watchParams({ fp: seen.body.fp }));
  assert.equal(state.calls.length, 2, 'the dashboard re-arms for free too');
  await watch(gateway, watchParams());
  assert.equal(state.calls.length, 3, 'and a reloaded dashboard reads the truth');
});

test('a held chain takes one safety read when its copy crosses the cadence, and keeps holding', async () => {
  const { state, gateway } = harness({ [SESSION]: [row({ status: 'waiting' })] });
  await poll(gateway, approvalPoll('900000001', { fp: '' }));
  assert.equal(state.calls.length, 1);

  await advance(state, 15000);   // still inside the cadence when the hold starts
  const { body } = await settle(state, poll(gateway, approvalPoll('900000001', { wait: 25, fp: WAITING_FP })));

  assert.ok(state.calls.length >= 2, 'it refreshed when the copy crossed the cadence');
  assert.equal(body.held, 25000, 'and the hold simply continued — nothing had changed');
  assert.equal(body.fp, WAITING_FP);
  assert.equal(gateway._debug().waiting, 0);
});

test('the safety cadence stays inside the window a copy is servable in', async () => {
  // HELD_REREAD_MS < STALE_MS, proven where it matters: a copy that has just
  // crossed the cadence is still good enough to answer from when Google is
  // down. If the two were the wrong way round, every held chain would go stale
  // before it ever refreshed.
  const { state, gateway } = harness({ [SESSION]: [row({ status: 'waiting' })] });
  await poll(gateway, approvalPoll('900000001', { fp: '' }));

  state.mode = 'error500';
  state.clock += 50000;
  const { body } = await settle(state, poll(gateway, approvalPoll('900000001', { wait: 25, fp: WAITING_FP })));
  assert.equal(state.calls.length, 2, 'past the cadence it really did try to refresh');
  assert.equal(body.stale, true, 'and what it has is still servable, marked stale');
  assert.equal(body.held, 0, 'which is never held');
});

test('forty examinees and a dashboard, ten minutes, nothing happening: ~30 executions, not ~300', async () => {
  const ids = Array.from({ length: 40 }, (_, i) => String(900000001 + i));
  const { state, gateway } = harness({ [SESSION]: ids.map(id => row({ id, status: 'waiting' })) });

  // Every page opens its chain: one read for the whole room.
  const opened = await Promise.all(ids.map(id => poll(gateway, approvalPoll(id, { wait: 25, fp: '' }))));
  const firstWatch = await watch(gateway, watchParams({ wait: 25, fp: '' }));
  assert.equal(state.calls.length, 1, 'one read for forty-one fresh chains');

  let pollFp = opened[0].body.fp;
  let watchFp = firstWatch.body.fp;
  for (let round = 0; round < 24; round++) {          // 24 x 25 s = ten minutes
    const holds = ids.map(id => poll(gateway, approvalPoll(id, { wait: 25, fp: pollFp })));
    const watching = watch(gateway, watchParams({ wait: 25, fp: watchFp }));
    // Every chain must be parked before virtual time moves, or a straggler
    // would find its budget already spent and read Google for nothing — on a
    // loaded machine that is the difference between a gate and a coin toss.
    assert.equal(await parked(gateway, 41), 41, 'round ' + round + ': all forty-one parked');
    const answers = await settle(state, Promise.all(holds.concat([watching])));
    assert.equal(answers[0].body.held, 25000, 'round ' + round + ': every chain held its full 25 s');
    pollFp = answers[0].body.fp;
    watchFp = answers[40].body.fp;
  }

  assert.equal(state.clock - CLOCK0, 600000, 'ten minutes of forty-one open requests');
  assert.ok(state.calls.length <= 34 && state.calls.length >= 28,
    'about one execution per 20 s — got ' + state.calls.length + ', it used to be ~300');
  assert.equal(gateway._debug().waiting, 0, 'and nothing was left parked');
});

test('a nudge is what makes the Worker read: the drop, not the clock', async () => {
  const { state, gateway } = harness({ [SESSION]: [row({ status: 'in_exam', tokenHash: sha256Hex(TOKEN) })] });
  const seen = await watch(gateway, watchParams({ fp: '' }));
  assert.equal(state.calls.length, 1);

  const holding = watch(gateway, watchParams({ wait: 25, fp: seen.body.fp }));
  const polling = poll(gateway, statusPoll('900000001', { wait: 25, fp: 's:in_exam:0', examineeToken: TOKEN }));
  assert.equal(await parked(gateway, 2), 2);

  // Twenty seconds of silence buy nothing at all...
  await advance(state, 15000);
  assert.equal(state.calls.length, 1, 'fifteen seconds of holding, zero executions');
  assert.equal(gateway._debug().waiting, 2);

  // ...and then the device submits and says so.
  state.snapshots[SESSION] = [row({ status: 'completed', tokenHash: sha256Hex(TOKEN) })];
  assert.deepEqual((await devicePush(gateway, '&idNumber=900000001&examineeToken=' + TOKEN)).body,
    { status: 'ok', dropped: true });

  const dash = await settle(state, holding);
  const device = await settle(state, polling);
  assert.equal(state.calls.length, 2, 'ONE read, because something actually happened');
  assert.notEqual(dash.body.fp, seen.body.fp);
  assert.equal(device.body.examStatus, 'completed');
  assert.equal(state.clock - CLOCK0, 15000, 'and both saw it without waiting out a tick');
});

// --- r32: the watch answer CARRIES the dashboard (DESIGN §14.1) ------------
// Until r32 a change still cost the dashboard a second round trip to Google
// just to see WHAT changed — and every round trip is a lottery ticket while
// Google's delivery hop stalls 25-60 s for our projects (KNOWN_ISSUES #35).
// An r32 server answers sessionSnapshot with `v:2`, the full rows and the
// session's results; the watch now hands them to the page in `session`.

/** An r32 snapshot row: every r31 field, plus what the dashboard paints with. */
const fullRow = over => row(Object.assign({
  name: 'ישראל ישראלי', phone: '0500000001', time: '2026-09-22T06:30:00.000Z', start: '',
  lang: 'he', pop: 'חוגר', site: 'דימונה', lic: 'B', timeExt: '', lastWarn: '', attemptsToday: 1
}, over));

/** The same rows an r32 server would send: same watched fields, more payload. */
const fullRowsOf = rows => rows.map(r => fullRow(Object.assign({}, r)));

/** One item of `results[]` — the fields §14.1 names, in the order it names them. */
const result = over => Object.assign({
  date: '22/09/2026 09:41', idNumber: '900000001', name: 'ישראל ישראלי', phone: '0500000001',
  license: 'B', score: 27, percent: 90, passed: true, time: '18:42', examiner: 'בוחן א',
  site: 'דימונה', classroom: 'כיתה 1', language: 'he', attempt: 1, sent: '', disqualified: '',
  waLink: 'https://wa.me/972500000001', population: 'חוגר', corrected: '', audioMode: 'off',
  verified: '', suspicious: '', device: 'Windows'
}, over);

/** A harness whose fake Apps Script speaks snapshot v2. */
function v2Harness(rows, results, withCache) {
  const made = harness({ [SESSION]: rows }, null, withCache);
  made.state.serverV = 2;
  made.state.results[SESSION] = results || [];
  return made;
}

const SNAP_KEY = 'https://session-gateway.internal/snap/' + SESSION;

/** The fingerprint a v2 session of exactly these rows and results produces. */
async function sessionFpV2(rows, results) {
  const { gateway } = v2Harness(rows, results);
  return (await watch(gateway, watchParams())).body.fp;
}

test('watch: `session` carries the rows and the results, and never a token hash', async () => {
  const rows = [
    fullRow({ id: '900000001', status: 'in_exam', tokenHash: sha256Hex(TOKEN), start: '09:12',
              attemptsToday: 2, todayExams: [{ license: 'B', score: 24, passed: false, language: 'he' }] }),
    fullRow({ id: '900000002', status: 'waiting', name: 'דנה כהן' })
  ];
  const results = [result(), result({ idNumber: '900000003', attempt: 2, passed: false, score: 21 })];
  const { state, gateway } = v2Harness(rows, results);

  const { body } = await watch(gateway, watchParams());
  assert.equal(state.calls.length, 1, 'one snapshot read — and the dashboard needs nothing else');
  assert.equal(body.rows, 2, '`rows` is still the COUNT: an old page reads exactly what it read in r31');
  assert.equal(body.session.rows.length, 2);
  const withoutHash = Object.assign({}, rows[1]);
  delete withoutHash.tokenHash;
  assert.deepEqual(body.session.rows[1], withoutHash, 'every other field of the row travels whole');
  assert.deepEqual(body.session.results, results, 'and the results exactly as the server sent them');
  assert.deepEqual(body.session.rows[0].todayExams, rows[0].todayExams,
    "today's other attempts come through too");

  // The one field that must never leave this Worker.
  assert.equal('tokenHash' in body.session.rows[0], false, 'the hash is stripped, not emptied');
  assert.equal(JSON.stringify(body).includes(sha256Hex(TOKEN)), false,
    'and nothing else in the answer carries it either');

  assert.equal(state.logs.at(-1).sess, 1, 'the tail says a payload went out');
  assert.equal(state.logs.at(-1).nr, 2, '...and how many results it carried');
});

test('watch: an r31 server answers no `session` at all, and the page falls back', async () => {
  // The deploy order of §14.5 is Worker first, server second: for those hours
  // the Worker talks to an r31 server and must behave exactly like r31.3c.
  const { state, gateway } = harness({ [SESSION]: [fullRow({ status: 'waiting' })] });
  const { body } = await watch(gateway, watchParams());
  assert.equal(body.session, undefined, 'no `v` in the snapshot means no payload');
  assert.equal(body.rows, 1);
  assert.equal(state.logs.at(-1).nr, 0);
  assert.equal(state.logs.at(-1).sess, undefined);
});

test('watch: `session` is attached only when the fingerprint moved, and never on a stale view', async () => {
  const rows = [fullRow({ status: 'waiting' })];
  const { state, gateway } = v2Harness(rows, [result()]);

  const first = await watch(gateway, watchParams());
  assert.ok(first.body.session, 'the first answer of a chain has no `fp` to compare, so it carries the data');

  // Re-arming with the fingerprint it already holds: nothing changed, so the
  // answer stays the small one. 25 s x every open dashboard is why.
  const rearm = await watch(gateway, watchParams({ fp: first.body.fp }));
  assert.equal(rearm.body.fp, first.body.fp);
  assert.equal(rearm.body.session, undefined, 'the page already has exactly this');

  const ranOut = await settle(state, watch(gateway, watchParams({ wait: 5, fp: first.body.fp })));
  assert.equal(ranOut.body.held, 5000, 'a hold that expires unchanged...');
  assert.equal(ranOut.body.session, undefined, '...carries nothing either');
  assert.equal(state.logs.at(-1).sess, undefined);
  assert.equal(gateway._debug().waiting, 0);

  // A change: the data comes with the news, in one round trip.
  state.snapshots[SESSION] = [fullRow({ status: 'approved', examMinutes: 50 })];
  await nudge(gateway, '');
  const changed = await watch(gateway, watchParams({ fp: first.body.fp }));
  assert.notEqual(changed.body.fp, first.body.fp);
  assert.equal(changed.body.session.rows[0].status, 'approved');

  // ...and a copy we could not refresh is never painted from: the dashboard
  // falls back to its own safety net instead of showing a minute-old class.
  state.mode = 'error500';
  await advance(state, 50000);
  const onStale = await watch(gateway, watchParams({ wait: 25, fp: 's:something-else' }));
  assert.equal(onStale.body.stale, true);
  assert.equal(onStale.body.session, undefined, 'stale never paints');
  assert.equal(onStale.body.held, 0);
});

test('watch: the fingerprint moves on every result field, and on a new result', async () => {
  const rows = [fullRow({ status: 'completed' })];
  const baseline = await sessionFpV2(rows, [result()]);
  assert.match(baseline, /^s:[0-9a-f]{12}$/);

  // Every field of `results[]`, including the `fabricated` flag that is only
  // there when it is 1. A corrected score or a "sent" tick that did not move
  // the fingerprint would simply never reach the examiner's screen.
  const changed = {
    date: '22/09/2026 10:02', idNumber: '900000009', name: 'שם אחר', phone: '0500000002',
    license: 'C1', score: 28, percent: 93, passed: false, time: '19:00', examiner: 'בוחן ב',
    site: 'באר שבע', classroom: 'כיתה 2', language: 'ru', attempt: 2, sent: 'נשלח',
    disqualified: 'פסול', waLink: 'https://wa.me/972500000002', population: 'קבע',
    corrected: 'תוקן', audioMode: 'on', verified: 'מאומת', suspicious: 'חשוד',
    device: 'Android', fabricated: 1
  };
  for (const field of Object.keys(changed)) {
    assert.notEqual(await sessionFpV2(rows, [result({ [field]: changed[field] })]), baseline,
      field + ' must move the fingerprint');
  }

  assert.notEqual(await sessionFpV2(rows, [result(), result({ idNumber: '900000002' })]), baseline,
    'a second result must move it');
  assert.notEqual(await sessionFpV2(rows, []), baseline, 'and so must losing one');
  assert.equal(await sessionFpV2(rows, [result()]), baseline, 'the same session hashes the same, always');
});

// The r31 formula, computed here from first principles: if these two ever
// disagree, deploying this Worker ahead of the server paste (§14.5) would move
// every fingerprint in the fleet and two Worker versions would hash one
// unchanged session two ways — the flip-flop of 22/09, all over again.
const R31_FP_FIELDS = ['id', 'status', 'audio', 'examMinutes', 'extraMinutes', 'warn', 'fin', 'ext', 'dq'];
const R31_ZEROABLE = { warn: 1, fin: 1, ext: 1, dq: 1 };
function r31SessionFp(rows) {
  if (!rows.length) return 's:none';
  const cell = (r, f) => {
    const value = r[f];
    if (value == null || value === '') return '';
    if (R31_ZEROABLE[f] && (value === 0 || value === '0' || value === false)) return '';
    return String(value);
  };
  const text = rows.map(r => R31_FP_FIELDS.map(f => cell(r, f)).join('|')).join(';');
  return 's:' + crypto.createHash('sha256').update(text, 'utf8').digest('hex').slice(0, 12);
}

test('watch: with no results the fingerprint is byte-for-byte the one r31 computed', async () => {
  const rows = [
    row({ id: '900000001', status: 'in_exam', audio: 'on', examMinutes: 50, extraMinutes: 5,
          warn: 1, fin: 0, ext: 0, dq: '' }),
    row({ id: '900000002', status: 'waiting' })
  ];
  const r31 = r31SessionFp(rows);
  assert.match(r31, /^s:[0-9a-f]{12}$/);

  assert.equal(await sessionFp(rows), r31, 'an r31 server: unchanged, to the byte');
  assert.equal(await sessionFpV2(rows, []), r31,
    'and an r32 server with no results yet hashes the very same text');
  assert.equal(await sessionFpV2(fullRowsOf(rows), []), r31,
    "the row fields r32 added are not watched — they cannot change without one that is");

  // ...and the empty session keeps its own name, whatever the server speaks.
  assert.equal(await sessionFp([]), 's:none');
  assert.equal(await sessionFpV2([], []), 's:none');
  assert.notEqual(await sessionFpV2([], [result()]), 's:none',
    'a session whose pending rows were archived still wakes on its results');
});

test('the examinee poll never sees a v2 row — not one extra field', async () => {
  const { gateway } = v2Harness([fullRow({ status: 'approved', examMinutes: 50 })], [result()]);

  const approval = await poll(gateway, approvalPoll('900000001'));
  assert.deepEqual(approval.body, { status: 'ok', approval: 'approved', audioMode: 'off', examMinutes: 50 },
    'byte for byte the server\'s own checkApproval, exactly as contracts.test.cjs pins it');
  const status = await poll(gateway, statusPoll('900000001'));
  assert.deepEqual(status.body, { status: 'ok', examStatus: 'approved', extraMinutes: 0 });

  const longPoll = await poll(gateway, approvalPoll('900000001', { fp: '' }));
  assert.deepEqual(longPoll.body, {
    status: 'ok', approval: 'approved', audioMode: 'off', examMinutes: 50,
    fp: 'a:approved:off:50', held: 0
  }, 'and a long-polling client still gets only fp/held on top');

  // The names, phones and results the dashboard needs are none of a device's
  // business — and an examinee's phone is the least trusted thing we serve.
  for (const answer of [approval.body, status.body, longPoll.body]) {
    const text = JSON.stringify(answer);
    assert.equal(text.includes('ישראל'), false, 'no name reaches a device');
    assert.equal(text.includes('0500000001'), false, 'no phone either');
    assert.equal(text.includes('results'), false);
    assert.equal(text.includes('session'), false);
  }
});

test('a patch keeps `v` and the results, and the held watch is answered with them', async () => {
  const { state, gateway } = v2Harness([fullRow({ status: 'waiting' })], [result()]);
  const seen = await watch(gateway, watchParams());
  assert.ok(seen.body.session, 'v2 from the start');

  const holding = watch(gateway, watchParams({ wait: 25, fp: seen.body.fp }));
  assert.equal(await parked(gateway, 1), 1);

  // The examiner approved in another tab: the decision-carrying nudge patches
  // ONE pending row and must not lose the rest of the snapshot with it.
  assert.equal((await nudge(gateway, '&idNumber=900000001&status=approved&examMinutes=50&audio=on')).body.patched, true);

  const { body } = await settle(state, holding);
  assert.notEqual(body.fp, seen.body.fp);
  assert.ok(body.session, 'the patched copy is still a v2 copy');
  assert.equal(body.session.rows[0].status, 'approved');
  assert.equal(body.session.rows[0].examMinutes, 50);
  assert.equal(body.session.rows[0].name, 'ישראל ישראלי', 'the row the server sent, with the decision written in');
  assert.deepEqual(body.session.results, seen.body.session.results, 'and the results the patch never touched');
  assert.equal(state.calls.length, 1, 'all of it for zero Apps Script executions');
  assert.equal(gateway._debug().waiting, 0);
});

// --- CPU: a quiet tick reads HEADERS, not a 5-40 KB body (TODO 0א.6) -------
// A held request re-reads caches.default every second — that is the channel a
// patch from another isolate arrives through — and `JSON.parse` of a whole
// session, every second, is most of what the request costs against a free-plan
// ceiling of 10 ms. `cacheWrite` now stamps `X-SFP`/`X-RAT` on the stored
// Response, so the usual tick is a `match` and a string compare.

test('a quiet 25 s watch looks at the cache every tick and parses it at most once', async () => {
  const { state, spawn } = v2Harness([fullRow({ status: 'waiting' })], [result()], true);
  const gateway = spawn();
  const first = await watch(gateway, watchParams());
  assert.equal(state.store.get(SNAP_KEY).headers['X-SFP'], first.body.fp,
    'the stored Response names its fingerprint in a header');
  assert.equal(state.store.get(SNAP_KEY).headers['X-RAT'], String(CLOCK0));

  await advance(state, 19000);   // so a safety read falls inside the hold
  const matchesBefore = state.cacheCounts.match;
  const parsesBefore = state.cacheCounts.json;
  const { body } = await settle(state, watch(gateway, watchParams({ wait: 25, fp: first.body.fp })));
  const matches = state.cacheCounts.match - matchesBefore;
  const parses = state.cacheCounts.json - parsesBefore;

  assert.equal(body.held, 25000, 'it really held the whole time');
  assert.equal(body.session, undefined, 'and nothing changed');
  assert.ok(matches >= 20, 'it looked at the cache on every tick — got ' + matches);
  assert.ok(parses <= 1, 'but parsed the body at most once in 25 seconds — got ' + parses +
    ' (r31 parsed it on every tick a memory copy answered: 23 of them)');
  assert.equal(gateway._debug().waiting, 0);
});

test('a cache copy written before r32 carries no headers, and its body is read as before', async () => {
  const { state, spawn } = harness({ [SESSION]: [row({ status: 'waiting' })] }, null, true);
  const gateway = spawn();
  const first = await watch(gateway, watchParams());

  // Exactly what the previous deploy left behind: a body, and no headers of
  // ours at all. "Not known" must fall back to the full read, never to a
  // wrong answer.
  const stored = state.store.get(SNAP_KEY);
  state.store.set(SNAP_KEY, { body: stored.body, maxAge: stored.maxAge, at: stored.at });

  const parsesBefore = state.cacheCounts.json;
  const held = await settle(state, watch(gateway, watchParams({ wait: 3, fp: first.body.fp })));
  assert.equal(held.body.fp, first.body.fp, 'the same answer r31 gave');
  assert.equal(held.body.held, 3000);
  assert.ok(state.cacheCounts.json - parsesBefore >= 1,
    'and it did read the body — the header check only ever skips work it can prove is redundant');
});

test('a snapshot another isolate wrote into the cache still reaches a held watch in one tick', async () => {
  const { state, spawn } = v2Harness([fullRow({ status: 'waiting' })], [result()], true);
  const first = spawn(), second = spawn();
  const seen = await watch(first, watchParams());

  const holding = watch(first, watchParams({ wait: 25, fp: seen.body.fp }));
  assert.equal(await parked(first, 1), 1);

  // A DIFFERENT isolate takes the examiner's nudge: it cannot wake anything in
  // the first one, so the only channel left is caches.default — which is the
  // very thing the header check must not be allowed to hide.
  assert.equal((await nudge(second, '&idNumber=900000001&status=approved&examMinutes=50')).body.patched, true);
  assert.equal(first._debug().waiting, 1, 'nothing woke it — it is still parked');

  const { body } = await settle(state, holding);
  assert.equal(body.held, 1000, 'seen on the very next tick');
  assert.notEqual(body.fp, seen.body.fp);
  assert.equal(body.session.rows[0].status, 'approved', 'with the data, from the cached copy');
  assert.deepEqual(body.session.results, seen.body.session.results);
  assert.equal(state.calls.length, 1, 'and still no upstream read');
});
