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
 * `caches.default` in memory: max-age is honoured against the same fake clock
 * the gateway reads, so an expiring `snap` copy behaves as it does at the edge.
 */
function cacheDouble(now) {
  const store = new Map();
  return { store, caches: { default: {
    async match(request) {
      const hit = store.get(request.url);
      if (!hit || now() - hit.at > hit.maxAge * 1000) return undefined;
      return new Response(hit.body, { status: 200 });
    },
    async put(request, response) {
      const maxAge = Number(/max-age=(\d+)/.exec(response.headers.get('Cache-Control') || '')[1]);
      store.set(request.url, { body: await response.text(), maxAge, at: now() });
    },
    async delete(request) { return store.delete(request.url); }
  } } };
}

/**
 * Fake upstream + fake clock; `state.calls` is the Apps Script execution count.
 * `withCache` adds a shared caches.default and `spawn()`, which is a COLD
 * isolate: new memory, same cache, same clock, same upstream counter.
 */
function harness(snapshots, assets, withCache) {
  const state = { calls: [], clock: CLOCK0, mode: 'ok', snapshots: snapshots || {} };
  const fetchFn = async url => {
    state.calls.push(String(url));
    if (state.mode === 'error500') return new Response('<html>Google internal error</html>', { status: 500 });
    if (state.mode === 'notfound') return new Response('Not Found', { status: 404 });
    if (state.mode === 'html200') return new Response('<!DOCTYPE html><html>Drive error</html>', { status: 200 });
    const session = new URL(url).searchParams.get('sessionCode');
    const rows = state.snapshots[session] || [];
    return new Response(JSON.stringify({ status: 'ok', at: state.clock, rows }), { status: 200 });
  };
  const env = assets ? Object.assign({}, ENV, { ASSETS: assets.fetch ? assets : assetsBinding(assets) }) : ENV;
  const double = withCache ? cacheDouble(() => state.clock) : null;
  state.store = double ? double.store : null;
  const spawn = () => createGateway({
    fetch: fetchFn, caches: double ? double.caches : undefined, now: () => state.clock, env
  });
  return { state, gateway: spawn(), spawn };
}

/** Any route: the raw text matters for the bank, which never re-serialises. */
async function call(gateway, pathAndQuery, init) {
  const res = await gateway(new Request('https://session-gateway.test' + pathAndQuery, init));
  const text = await res.text();
  let body = null;
  try { body = JSON.parse(text); } catch (e) { /* the test asserts on text */ }
  return { res, text, body };
}

function pollRequest(params, origin) {
  const url = 'https://session-gateway.test/v1/poll?' + new URLSearchParams(params).toString();
  return new Request(url, origin ? { headers: { Origin: origin } } : undefined);
}
async function poll(gateway, params, origin) {
  const res = await gateway(pollRequest(params, origin));
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
    row({ id: '900000002', status: 'cancelled' })
  ] });

  const active = await poll(gateway, approvalPoll('900000001'));
  assert.deepEqual(active.body, { status: 'ok', approval: 'approved', audioMode: 'on', examMinutes: 60 });

  const allTerminal = await poll(gateway, approvalPoll('900000002'));
  assert.deepEqual(allTerminal.body, { status: 'error', message: 'לא נמצא רישום' });
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
