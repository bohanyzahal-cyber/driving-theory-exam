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
const sha256Hex = text => crypto.createHash('sha256').update(text).digest('hex');
const row = over => Object.assign(
  { id: '900000001', status: 'waiting', tokenHash: '', audio: 'off', examMinutes: 40, extraMinutes: 0 }, over);

/** Fake upstream + fake clock; `state.calls` is the Apps Script execution count. */
function harness(snapshots) {
  const state = { calls: [], clock: 1000000, mode: 'ok', snapshots: snapshots || {} };
  const fetchFn = async url => {
    state.calls.push(String(url));
    if (state.mode === 'error500') return new Response('<html>Google internal error</html>', { status: 500 });
    if (state.mode === 'notfound') return new Response('Not Found', { status: 404 });
    if (state.mode === 'html200') return new Response('<!DOCTYPE html><html>Drive error</html>', { status: 200 });
    const session = new URL(url).searchParams.get('sessionCode');
    const rows = state.snapshots[session] || [];
    return new Response(JSON.stringify({ status: 'ok', at: state.clock, rows }), { status: 200 });
  };
  const gateway = createGateway({ fetch: fetchFn, caches: undefined, now: () => state.clock, env: ENV });
  return { state, gateway };
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

test('a poll inside the 3s window is free; past it costs one more execution', async () => {
  const { state, gateway } = harness({ [SESSION]: [row({ status: 'approved' })] });

  await poll(gateway, approvalPoll('900000001'));
  assert.equal(state.calls.length, 1);

  state.clock += 2000;
  await poll(gateway, approvalPoll('900000001'));
  assert.equal(state.calls.length, 1, 'still inside the freshness window');

  state.clock += 1500; // 3.5s after the first read
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

test('a row missing from a cached snapshot forces one re-read, at most once per 10s', async () => {
  const { state, gateway } = harness({ [SESSION]: [row({ id: '900000001', status: 'approved' })] });

  await poll(gateway, approvalPoll('900000001'));
  assert.equal(state.calls.length, 1);

  // The examinee registered a moment after the snapshot was taken.
  state.snapshots[SESSION].push(row({ id: '900000002', status: 'waiting' }));
  const late = await poll(gateway, approvalPoll('900000002'));
  assert.equal(state.calls.length, 2, 'one forced re-read before answering "not registered"');
  assert.equal(late.body.approval, 'waiting');

  state.clock += 1000;
  const again = await poll(gateway, approvalPoll('900000003'));
  assert.equal(state.calls.length, 2, 'no second forced re-read inside the 10s window');
  assert.deepEqual(again.body, { status: 'error', message: 'לא נמצא רישום' });

  state.clock += 1000;
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
  const store = new Map();
  const fakeCaches = { default: {
    async match(request) {
      const hit = store.get(request.url);
      if (!hit || clock.now - hit.at > hit.maxAge * 1000) return undefined;
      return new Response(hit.body, { status: 200 });
    },
    async put(request, response) {
      const maxAge = Number(/max-age=(\d+)/.exec(response.headers.get('Cache-Control') || '')[1]);
      store.set(request.url, { body: await response.text(), maxAge, at: clock.now });
    }
  } };
  const clock = { now: 1000000 };
  const calls = [];
  const fetchFn = async url => {
    calls.push(String(url));
    return new Response(JSON.stringify({ status: 'ok', at: clock.now, rows: [row({ status: 'approved' })] }));
  };
  const spawn = () => createGateway({ fetch: fetchFn, caches: fakeCaches, now: () => clock.now, env: ENV });

  const first = await poll(spawn(), approvalPoll('900000001'));
  assert.equal(first.body.approval, 'approved');
  assert.equal(calls.length, 1);
  assert.equal(store.size, 2, 'a fresh copy and a stale copy');

  const second = await poll(spawn(), approvalPoll('900000001'));
  assert.equal(calls.length, 1, 'the cold isolate reused the cached snapshot');
  assert.equal(second.body.approval, 'approved');

  clock.now += 4000; // fresh copy expired, stale copy still there
  const third = spawn();
  await poll(third, approvalPoll('900000001'));
  assert.equal(calls.length, 2);
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
  assert.equal(res.headers.get('Access-Control-Allow-Methods'), 'GET, OPTIONS');
  assert.equal(state.calls.length, 0);
});

test('the health route names the service and its build', async () => {
  const { gateway } = harness({});
  const res = await gateway(new Request('https://session-gateway.test/'));
  const body = await res.json();
  assert.equal(body.status, 'ok');
  assert.equal(body.service, 'session-gateway');
  assert.match(body.build, /^\d{4}-\d{2}-\d{2}$/);
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
