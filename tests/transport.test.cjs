// shared/transport.js — the one client transport layer (examinee + examiner +
// teacher + student). Everything here runs the REAL module in a VM on a fake
// clock; only fetch and the clock are synthetic. No network, no browser.
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const app = path.resolve(__dirname, '..');
const SOURCE = fs.readFileSync(path.join(app, 'shared', 'transport.js'), 'utf8');

async function drain() { for (let i = 0; i < 30; i++) await Promise.resolve(); }

class Timers {
  constructor() { this.now = 0; this.nextId = 0; this.jobs = new Map(); }
  set = (cb, ms) => { const id = ++this.nextId; this.jobs.set(id, { cb, at: this.now + Number(ms || 0), every: 0 }); return id; };
  setInterval = (cb, ms) => { const id = ++this.nextId; this.jobs.set(id, { cb, at: this.now + Number(ms || 0), every: Number(ms || 0) }); return id; };
  clear = id => this.jobs.delete(id);
  pending() { return this.jobs.size; }
  async advance(ms) {
    const until = this.now + ms;
    for (let guard = 0; guard < 10000; guard++) {
      const next = [...this.jobs].filter(([, job]) => job.at <= until).sort((a, b) => a[1].at - b[1].at)[0];
      if (!next) { this.now = until; await drain(); return; }
      this.now = next[1].at;
      if (next[1].every) next[1].at = this.now + next[1].every; else this.jobs.delete(next[0]);
      next[1].cb();
      await drain();
    }
    throw new Error('timer loop did not settle');
  }
}

/** Loads the real module into a fresh VM. `fetch` is whatever the test needs. */
function load(fetchImpl) {
  const timer = new Timers();
  const clock = class extends Date { static now() { return timer.now; } };
  const ctx = {
    console: { log() {}, warn() {}, error() {} },
    Date: clock, Math, JSON, Promise, Error, String, Number, Object, Array, encodeURIComponent,
    setTimeout: timer.set, clearTimeout: timer.clear,
    setInterval: timer.setInterval, clearInterval: timer.clear,
    AbortController, fetch: fetchImpl || (() => new Promise(() => {}))
  };
  ctx.window = ctx;
  vm.createContext(ctx);
  vm.runInContext(SOURCE, ctx, { filename: 'shared/transport.js' });
  return { T: ctx.ExamTransport, timer, ctx };
}

const reply = (body, { ok = true, status = 200 } = {}) => () =>
  Promise.resolve({ ok, status, text: () => Promise.resolve(body) });
const settle = p => p.then(value => ({ value }), error => ({ error }));

// ===== 1. failure classification =====
test('transport: a hung request fails as a timeout and aborts the fetch', async () => {
  let signal;
  const { T, timer } = load((url, opts) => { signal = opts.signal; return new Promise(() => {}); });
  const out = settle(T.fetchJsonWithTimeout('synthetic', {}, 1000));
  await drain();
  await timer.advance(1000);
  const { error } = await out;
  assert.equal(error.name, 'TimeoutError');
  assert.equal(error.transport, 'timeout');
  assert.equal(signal.aborted, true);
  assert.equal(timer.pending(), 0, 'the deadline timer is cleared');
});

test('transport: network, HTTP and non-JSON answers are each classified', async () => {
  const cases = [
    [() => Promise.reject(new TypeError('Failed to fetch')), 'network', null],
    [reply('{}', { ok: false, status: 503 }), 'http', 503],
    [reply('<html>Google error</html>'), 'nonjson', null]
  ];
  for (const [fetchImpl, transport, status] of cases) {
    const { T } = load(fetchImpl);
    const { error } = await settle(T.fetchJsonWithTimeout('synthetic', {}, 1000));
    assert.equal(error.transport, transport);
    if (status) assert.equal(error.status, status);
    if (transport === 'nonjson') assert.match(error.bodyHead, /Google error/);
  }
});

test('transport: a JSON answer resolves with the parsed body', async () => {
  const { T } = load(reply('{"status":"ok","approval":"waiting"}'));
  const { value } = await settle(T.fetchJsonWithTimeout('synthetic', {}, 1000));
  assert.deepEqual(value, { status: 'ok', approval: 'waiting' });
});

// ===== 2. backend health =====
test('transport health: an HTTP error or an HTML page degrades at once, JSON clears it', async () => {
  for (const fetchImpl of [reply('{}', { ok: false, status: 429 }), reply('<html>busy</html>')]) {
    const { T } = load(fetchImpl);
    assert.equal(T.isBackendDegraded(), false);
    await settle(T.fetchJsonWithTimeout('synthetic', {}, 1000));
    assert.equal(T.isBackendDegraded(), true, 'our server always answers 200 + JSON');
    assert.ok(T.degradedSince() >= 0);
  }
});

test('transport health: ONE timeout is a hiccup, two in a row are a degraded backend (S10)', async () => {
  const { T, timer } = load(() => new Promise(() => {}));
  settle(T.fetchJsonWithTimeout('a', {}, 1000)); await drain(); await timer.advance(1000);
  assert.equal(T.isBackendDegraded(), false, 'one ~60 s hang looks exactly like no network');
  settle(T.fetchJsonWithTimeout('b', {}, 1000)); await drain(); await timer.advance(1000);
  assert.equal(T.isBackendDegraded(), true, 'the most common stall shape now gets the strong response');
});

test('transport health: a network error never degrades, and it resets the timeout run', async () => {
  const { T, timer } = load(() => new Promise(() => {}));
  settle(T.fetchJsonWithTimeout('a', {}, 1000)); await drain(); await timer.advance(1000);
  T.noteTransport(Object.assign(new Error('offline'), { transport: 'network' }));
  assert.equal(T.isBackendDegraded(), false);
  settle(T.fetchJsonWithTimeout('b', {}, 1000)); await drain(); await timer.advance(1000);
  assert.equal(T.isBackendDegraded(), false, 'the run was broken by this device losing the network');
});

test('transport health: a good JSON answer clears a degraded state', async () => {
  const { T } = load(reply('<html>busy</html>'));
  await settle(T.fetchJsonWithTimeout('synthetic', {}, 1000));
  assert.equal(T.isBackendDegraded(), true);
  T.noteTransport(null);
  assert.equal(T.isBackendDegraded(), false);
  assert.equal(T.degradedSince(), 0);
});

// ===== 3. pacing =====
test('pacing: jitter stays within ±30% and actually varies', () => {
  const { T } = load();
  const seen = new Set();
  for (let i = 0; i < 300; i++) {
    const v = T.jitterMs(10000);
    assert.ok(Number.isInteger(v) && v >= 7000 && v <= 13000, 'within ±30%: ' + v);
    seen.add(v);
  }
  assert.ok(seen.size > 50, 'a fleet released together must not re-arrive together');
});

test('pacing: the ladder steps up ×1.5 to the cap and down one step per healthy answer', () => {
  const { T } = load();
  let ms = 5000;
  const up = [];
  for (let i = 0; i < 5; i++) { ms = T.nextPollDelay(ms, 5000, 20000, true); up.push(ms); }
  assert.deepEqual(up, [7500, 11250, 16875, 20000, 20000]);
  const down = [];
  for (let i = 0; i < 4; i++) { ms = T.nextPollDelay(ms, 5000, 20000, false); down.push(ms); }
  assert.deepEqual(down, [13333, 8889, 5926, 5000]);
});

test('pacing: a degraded backend floors the whole fleet at 30-60 s', async () => {
  const { T } = load(reply('<html>busy</html>'));
  assert.equal(T.pacePoll(5000, 5000, 20000, false), 5000);
  await settle(T.fetchJsonWithTimeout('synthetic', {}, 1000));
  assert.equal(T.pacePoll(5000, 5000, 20000, false), 30000, 'never faster than 30 s');
  assert.equal(T.pacePoll(50000, 5000, 20000, true), 60000, 'and never slower than 60 s');
  T.noteTransport(null);
  assert.equal(T.pacePoll(30000, 5000, 20000, false), 20000, 'recovery goes back through the normal ladder');
});

// ===== 4. createPollLoop =====
function loopWith(tick, opts = {}) {
  const { T, timer } = load();
  T._setJitter(ms => ms);                       // exact fake clock
  const loop = T.createPollLoop(Object.assign({ name: 'test', baseMs: 5000, maxMs: 20000, tick }, opts));
  return { T, timer, loop };
}

test('poll loop: a hanging tick never gets a second request beside it (D2)', async () => {
  let calls = 0;
  const { timer, loop } = loopWith(() => { calls++; return new Promise(() => {}); });
  loop.start(); await drain();
  assert.equal(calls, 1);
  await timer.advance(10 * 60 * 1000);
  assert.equal(calls, 1, 'answer → wait → ask again: nothing is scheduled while a request is out');
  assert.equal(loop.isInFlight(), true);
});

test('poll loop: restartIfStuck waits for the poll deadline before reviving a chain', async () => {
  let calls = 0, restarts = 0;
  const { timer, loop } = loopWith(() => { calls++; return new Promise(() => {}); }, { onRestart: () => restarts++ });
  loop.start(); await drain();
  await timer.advance(59000);
  assert.equal(loop.restartIfStuck(), false, 'a request that is merely slow is left alone');
  assert.equal(calls, 1);
  await timer.advance(2000);
  assert.equal(loop.restartIfStuck(), true);
  assert.equal(restarts, 1, 'the page gets to invalidate its generation first');
  assert.equal(calls, 2, 'and only then is the chain revived');
});

test('poll loop: four lock/wake cycles during one hang still leave ONE live request', async () => {
  let calls = 0;
  const { timer, loop } = loopWith(() => { calls++; return new Promise(() => {}); });
  loop.start(); await drain();
  for (let i = 0; i < 4; i++) { await timer.advance(8000); loop.restartIfStuck(); await drain(); }
  assert.equal(calls, 1, 'the 16/09 orphan multiplier is closed');
});

test('poll loop: stop() kills a late finaliser instead of letting it reschedule', async () => {
  let resolve, calls = 0;
  const { timer, loop } = loopWith(() => { calls++; return new Promise(r => { resolve = r; }); });
  loop.start(); await drain();
  loop.stop();
  resolve({ status: 'ok' }); await drain();
  await timer.advance(60000);
  assert.equal(calls, 1);
  assert.equal(timer.pending(), 0, 'no timer survives a stop');
});

test('poll loop: a failed tick backs off, a healthy one steps back down', async () => {
  let answer = { status: 'error', code: 'upstream_unavailable' };
  const { timer, loop } = loopWith(() => Promise.resolve(answer));
  loop.start(); await drain();
  assert.equal(loop.currentDelayMs(), 7500, 'a JSON error is a failed poll');
  await timer.advance(7500);
  assert.equal(loop.currentDelayMs(), 11250);
  answer = { status: 'ok' };
  await timer.advance(11250);
  assert.equal(loop.currentDelayMs(), 7500, 'one step down per healthy answer');
  loop.stop();
});

test('poll loop: a rejected tick is a failed poll and the loop survives it', async () => {
  let calls = 0;
  const { timer, loop } = loopWith(() => { calls++; return Promise.reject(new Error('network')); });
  loop.start(); await drain();
  assert.equal(loop.currentDelayMs(), 7500);
  await timer.advance(7500); assert.equal(calls, 2, 'the loop never dies');
  loop.stop();
});

test('poll loop: nextDelay can override the pacing (the 2 s result-sync window)', async () => {
  const { timer, loop } = loopWith(() => Promise.resolve({ status: 'ok' }), { nextDelay: () => 2000 });
  let calls = 0;
  loop.start(); await drain();
  const t0 = timer.now;
  await timer.advance(2000);
  assert.ok(timer.now - t0 === 2000);
  loop.stop();
});

// ===== 5. createFailover =====
function failoverWith(primary, fallback) {
  const { T, timer } = load();
  return { T, timer, fo: T.createFailover({ primary, fallback, failuresBeforeFallback: 3, fallbackMs: 300000 }) };
}

test('failover: three gateway TRANSPORT failures hand the next five minutes to the direct call', async () => {
  let primaryCalls = 0, fallbackCalls = 0, primaryFails = true;
  const { timer, fo } = failoverWith(
    () => { primaryCalls++; return primaryFails ? Promise.reject(Object.assign(new Error('gone'), { transport: 'http' })) : Promise.resolve({ status: 'ok', via: 'gateway' }); },
    () => { fallbackCalls++; return Promise.resolve({ status: 'ok', via: 'direct' }); });
  for (let i = 0; i < 3; i++) assert.equal((await fo.call()).via, 'direct');
  assert.equal(primaryCalls, 3); assert.equal(fallbackCalls, 3);
  assert.equal(fo.usingFallback(), true);
  primaryFails = false;
  assert.equal((await fo.call()).via, 'direct', 'still direct inside the five minutes');
  assert.equal(primaryCalls, 3, 'the gateway is not probed while it is parked');
  await timer.advance(300001);
  assert.equal((await fo.call()).via, 'gateway', 'and is tried again afterwards');
});

test('failover: upstream_unavailable is an ANSWER, not a reason to storm Apps Script', async () => {
  let primaryCalls = 0, fallbackCalls = 0;
  const { fo } = failoverWith(
    () => { primaryCalls++; return Promise.resolve({ status: 'error', code: 'upstream_unavailable', retryable: true }); },
    () => { fallbackCalls++; return Promise.resolve({ status: 'ok', via: 'direct' }); });
  for (let i = 0; i < 10; i++) {
    const answer = await fo.call();
    assert.equal(answer.code, 'upstream_unavailable');
  }
  assert.equal(fallbackCalls, 0, 'the gateway exists precisely to shield Google when it is busy');
  assert.equal(fo.usingFallback(), false);
  assert.equal(primaryCalls, 10);
});

test('failover: a good gateway answer resets the failure run', async () => {
  let fail = true, fallbackCalls = 0;
  const { fo } = failoverWith(
    () => fail ? Promise.reject(Object.assign(new Error('x'), { transport: 'timeout' })) : Promise.resolve({ via: 'gateway' }),
    () => { fallbackCalls++; return Promise.resolve({ via: 'direct' }); });
  await fo.call(); await fo.call();            // two failures
  fail = false; await fo.call();               // one success resets
  fail = true; await fo.call(); await fo.call();
  assert.equal(fo.usingFallback(), false, 'two + two is not three in a row');
  assert.equal(fallbackCalls, 4);
});

// ===== 6. createUpdateCheck =====
function updateCheckWith(answers) {
  const queue = answers.slice();
  const { T, timer, ctx } = load(() => {
    const a = queue.length ? queue.shift() : answers[answers.length - 1];
    if (!a || a.fail) return Promise.reject(new TypeError('network'));
    return Promise.resolve({ ok: a.status === undefined ? true : a.status < 400, status: a.status || 200,
      text: () => Promise.resolve(a.body !== undefined ? a.body : JSON.stringify({ build: a.build || 'b', pages: { 'examinee.html': a.hash } })) });
  });
  const seen = [];
  const check = T.createUpdateCheck({ page: 'examinee.html', intervalMs: 120000, onNewVersion: build => seen.push(build) });
  return { T, timer, check, seen, ctx };
}

test('update check: the first good answer is the base, and a change must be seen twice', async () => {
  const { timer, check, seen } = updateCheckWith([{ hash: 'aaa' }, { hash: 'bbb' }, { hash: 'bbb' }]);
  check.start(); await drain();
  await timer.advance(120000); assert.deepEqual(seen, [], 'one sighting is not a deploy');
  await timer.advance(120000); assert.equal(seen.length, 1, 'two consecutive sightings are');
  await timer.advance(600000); assert.equal(seen.length, 1, 'and it is announced exactly once');
  check.stop();
});

test('update check: 404s, HTML and network failures are never a new version', async () => {
  const noise = [{ hash: 'aaa' }];
  for (let i = 0; i < 20; i++) noise.push(i % 3 === 0 ? { status: 404 } : i % 3 === 1 ? { fail: true } : { body: '<html>portal</html>' });
  const { timer, check, seen } = updateCheckWith(noise);
  check.start(); await drain();
  await timer.advance(20 * 120000);
  assert.deepEqual(seen, []);
  check.stop();
});

test('update check: one odd hash between two good ones does not announce anything', async () => {
  const { timer, check, seen } = updateCheckWith([{ hash: 'aaa' }, { hash: 'aaa' }, { hash: 'zzz' }, { hash: 'aaa' }, { hash: 'aaa' }]);
  check.start(); await drain();
  await timer.advance(5 * 120000);
  assert.deepEqual(seen, []);
  check.stop();
});

test('update check: a version.json failure does not degrade the exam backend', async () => {
  const { T, timer, check } = updateCheckWith([{ status: 404 }, { status: 404 }, { status: 404 }]);
  check.start(); await drain();
  await timer.advance(3 * 120000);
  assert.equal(T.isBackendDegraded(), false, 'Pages is not Apps Script — the polls must not slow down');
  check.stop();
});

// ===== 7. client log ring =====
test('client log: bounded to 50 entries and ~2 KB, and draining empties it', () => {
  const { T } = load();
  for (let i = 0; i < 120; i++) T.log('event' + i, 'x'.repeat(40));
  const drained = T.drainLog();
  assert.ok(drained.length <= 50, 'ring is bounded: ' + drained.length);
  assert.ok(JSON.stringify(drained).length <= 2048, 'payload stays under 2 KB');
  assert.equal(drained[drained.length - 1].e, 'event119', 'the newest events are the ones kept');
  assert.equal(T.drainLog().length, 0, 'a drained log is not sent twice');
});

test('client log: long data is truncated and a failure never throws into the caller', () => {
  const { T } = load();
  T.log('x', 'y'.repeat(500));
  const [entry] = T.drainLog();
  assert.equal(entry.d.length, 120);
  assert.doesNotThrow(() => T.log('obj', { a: 1 }));
});

// ===== 8. createApi =====
test('api: every request carries the origin, a cache-buster and the page decorator', async () => {
  const seen = [];
  const { T } = load((url, opts) => { seen.push({ url, opts }); return Promise.resolve({ ok: true, status: 200, text: () => Promise.resolve('{"status":"ok"}') }); });
  const api = T.createApi({ apiUrl: 'https://api/exec', origin: 'examinee-app',
    decorate: p => { p.examineeToken = 'tok'; return p; } });
  await api.get({ action: 'checkApproval', sessionCode: 'ABC' });
  assert.match(seen[0].url, /^https:\/\/api\/exec\?/);
  assert.match(seen[0].url, /action=checkApproval/);
  assert.match(seen[0].url, /origin=examinee-app/);
  assert.match(seen[0].url, /examineeToken=tok/);
  assert.match(seen[0].url, /_t=\d+/);
  assert.equal(seen[0].opts.cache, 'no-store');
  await api.post({ action: 'submitResult' });
  assert.equal(seen[1].opts.method, 'POST');
  assert.deepEqual(JSON.parse(seen[1].opts.body), { action: 'submitResult', examineeToken: 'tok', origin: 'examinee-app' });
});

test('api: the caller always gets a deadline, even without asking for one', async () => {
  const { T, timer } = load(() => new Promise(() => {}));
  const api = T.createApi({ apiUrl: 'https://api/exec', origin: 'examinee-app' });
  const out = settle(api.get({ action: 'x' }));
  await drain();
  await timer.advance(T.API_TIMEOUT_MS);
  assert.equal((await out).error.name, 'TimeoutError');
});
