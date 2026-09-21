// client_fleet — what a ROOM full of these pages does to the server.
//
// The 16/09 "request storms" were not retry storms: every poll chain is
// answer → wait → ask again, so the fleet's arrival rate is INVERSELY
// proportional to server latency, while the number of ALIVE executions grows
// with it (aborting a fetch never stops an Apps Script run). This suite runs the
// REAL createPollLoop and the REAL fetchJsonWithTimeout, N clients at a time, on
// one fake clock, and measures both numbers.
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const app = path.resolve(__dirname, '..');
const TRANSPORT = fs.readFileSync(path.join(app, 'shared', 'transport.js'), 'utf8');

async function drain() { for (let i = 0; i < 12; i++) await Promise.resolve(); }

class Clock {
  constructor() { this.now = 0; this.nextId = 0; this.jobs = new Map(); }
  set = (cb, ms) => { const id = ++this.nextId; this.jobs.set(id, { cb, at: this.now + Number(ms || 0) }); return id; };
  clear = id => this.jobs.delete(id);
  async advance(ms) {
    const until = this.now + ms;
    for (let guard = 0; guard < 2000000; guard++) {
      let next = null, nextId = null;
      for (const [id, job] of this.jobs) if (job.at <= until && (!next || job.at < next.at)) { next = job; nextId = id; }
      if (!next) { this.now = until; await drain(); return; }
      this.now = next.at; this.jobs.delete(nextId);
      next.cb();
      await drain();
    }
    throw new Error('timer loop did not settle');
  }
}

function loadTransport(clock, fetchImpl) {
  const ctx = {
    console: { log() {}, warn() {}, error() {} },
    Date: class extends Date { static now() { return clock.now; } },
    Math, JSON, Promise, Error, TypeError, String, Number, Object, Array, encodeURIComponent,
    setTimeout: clock.set, clearTimeout: clock.clear, setInterval: clock.set, clearInterval: clock.clear,
    AbortController, fetch: fetchImpl
  };
  ctx.window = ctx;
  vm.createContext(ctx);
  vm.runInContext(TRANSPORT, ctx, { filename: 'shared/transport.js' });
  return ctx.ExamTransport;
}

/**
 * One fleet, one backend.
 *  clients: [{ baseMs, maxMs, count }]
 *  latency: how long the SERVER takes (an execution lives that long even after
 *           the client abandons it at the 60 s deadline)
 */
async function simulate({ clients, latency, minutes = 10, deadlineMs = 60000 }) {
  const clock = new Clock();
  const runs = [];                       // { client, start, end }
  const transport = loadTransport(clock, () => new Promise((resolve, reject) => {
    const run = { start: clock.now, end: clock.now + latency };
    runs.push(run);
    clock.set(() => resolve({ ok: true, status: 200, text: () => Promise.resolve('{"status":"ok"}') }), latency);
  }));
  const loops = [];
  let index = 0;
  for (const group of clients) {
    for (let i = 0; i < group.count; i++) {
      const id = index++;
      const loop = transport.createPollLoop({
        name: 'c' + id, baseMs: group.baseMs, maxMs: group.maxMs,
        tick: () => { const mine = runs.length; return transport.fetchJsonWithTimeout('poll', {}, deadlineMs).then(v => v, e => { throw e; }); }
      });
      loop.__client = id;
      loops.push(loop);
    }
  }
  // stagger the start over one base interval, as a real room does
  loops.forEach((loop, i) => clock.set(() => loop.start(), Math.floor((i * 1000) % 5000)));
  await clock.advance(minutes * 60 * 1000);
  loops.forEach(loop => loop.stop());

  const windowMs = minutes * 60 * 1000;
  const total = runs.length;
  const clientCount = loops.length;
  // time-average number of executions alive, which is what fills Apps Script's
  // ~30 slots — not the arrival rate everyone looks at.
  const aliveTime = runs.reduce((sum, r) => sum + (Math.min(r.end, windowMs) - r.start), 0);
  return {
    clock, transport, runs, loops,
    requestsPerMinute: total / minutes,
    aliveAverage: aliveTime / windowMs,
    alivePerClient: aliveTime / windowMs / clientCount,
    degraded: transport.isBackendDegraded(),
    periodOf: loop => loop.currentDelayMs()
  };
}

const APPROVAL_DIRECT = { baseMs: 8000, maxMs: 20000 };     // the constants examinee.html ships
const STATUS_DIRECT = { baseMs: 12000, maxMs: 20000 };
const APPROVAL_GATEWAY = { baseMs: 3000, maxMs: 20000 };    // 2 s for the first two minutes, then this
const STATUS_GATEWAY = { baseMs: 6000, maxMs: 20000 };

test('fleet: a healthy morning — 38 polling pages, one live request each at most', async () => {
  const fleet = await simulate({ latency: 2000, clients: [
    { ...APPROVAL_DIRECT, count: 18 },      // waiting for approval
    { ...STATUS_DIRECT, count: 20 }         // taking the exam
  ] });
  assert.ok(fleet.alivePerClient < 1, 'answer → wait → ask again: never two at once from one device (' + fleet.alivePerClient.toFixed(2) + ')');
  assert.ok(fleet.aliveAverage < 10, 'the ~30 execution slots are nowhere near full: ' + fleet.aliveAverage.toFixed(1));
  // the review measured 254 req/min for the same 38 clients on the OLD 5 s/10 s
  // constants; the shipped direct intervals must be below that.
  assert.ok(fleet.requestsPerMinute < 254, 'direct polling is cheaper than it was: ' + fleet.requestsPerMinute.toFixed(0) + '/min');
  assert.ok(fleet.requestsPerMinute > 150, 'and still answers within seconds: ' + fleet.requestsPerMinute.toFixed(0) + '/min');
});

test('fleet: the old 5 s/10 s constants reproduce the review\'s 254 req/min, so the model is the same one', async () => {
  const fleet = await simulate({ latency: 2000, clients: [
    { baseMs: 5000, maxMs: 20000, count: 18 },
    { baseMs: 10000, maxMs: 20000, count: 20 }
  ] });
  assert.ok(Math.abs(fleet.requestsPerMinute - 254) < 30, 'measured ' + fleet.requestsPerMinute.toFixed(0) + '/min (appendix D: 254)');
  assert.ok(fleet.alivePerClient < 1);
});

test('fleet: through the gateway the SAME room costs Apps Script one call per session per 2 s', async () => {
  const fleet = await simulate({ latency: 300, clients: [   // a Worker answers in ~300 ms
    { ...APPROVAL_GATEWAY, count: 18 },
    { ...STATUS_GATEWAY, count: 20 }
  ] });
  assert.ok(fleet.requestsPerMinute > 400, 'the examinees poll faster, because it is free: ' + fleet.requestsPerMinute.toFixed(0) + '/min');
  // What reaches Apps Script is the Worker's coalesced upstream: one snapshot
  // per session per FRESH_MS, no matter how many examinees are in the room.
  const upstreamPerMinute = 60 / 2;
  assert.equal(upstreamPerMinute, 30);
  assert.ok(upstreamPerMinute < fleet.requestsPerMinute / 10, 'a 15x reduction in executions for this room');
  assert.ok(fleet.alivePerClient < 1);
});

test('fleet: the 2 s opening window is still one live request per device', async () => {
  // The first two minutes after registration poll at 2 s through the gateway —
  // the window in which the examiner actually approves. Even at that cadence a
  // device holds at most one request, because the chain is answer → wait → ask.
  const fleet = await simulate({ latency: 300, minutes: 2, clients: [
    { baseMs: 2000, maxMs: 20000, count: 40 }
  ] });
  assert.ok(fleet.alivePerClient < 1, 'measured ' + fleet.alivePerClient.toFixed(2) + ' per device');
  assert.equal(fleet.degraded, false, 'and a Worker answering in 300 ms is never "slow"');
  assert.ok(60 / 2 < fleet.requestsPerMinute / 20, 'still one upstream per session per 2 s behind it');
});

test('fleet: a 93 s stall — the 60 s deadline plus backoff keeps orphans near one per device', async () => {
  const fleet = await simulate({ latency: 93000, minutes: 20, clients: [
    { ...APPROVAL_DIRECT, count: 18 },
    { ...STATUS_DIRECT, count: 20 }
  ] });
  assert.ok(fleet.alivePerClient <= 1.2,
    'orphan multiplication is what fills the slots; measured ' + fleet.alivePerClient.toFixed(2) + ' per device');
  assert.equal(fleet.degraded, true, 'two consecutive 60 s timeouts are a degraded backend (S10)');
  for (const loop of fleet.loops) {
    assert.ok(loop.currentDelayMs() >= 30000, 'and the whole fleet is floored at 30-60 s: ' + loop.currentDelayMs());
    assert.ok(loop.currentDelayMs() <= 60000);
  }
  assert.ok(fleet.requestsPerMinute < 60, 'arrivals COLLAPSE during a stall: ' + fleet.requestsPerMinute.toFixed(0) + '/min');
});

test('fleet: the 60 s poll deadline is what keeps the orphans down, not the intervals', async () => {
  // A1 of the 18/09 review. Same fleet, same stall, only the deadline differs:
  // abandoning at 30 s frees nothing server-side (the execution runs to 93 s)
  // and buys a second live request per device.
  const shortDeadline = await simulate({ latency: 93000, minutes: 20, deadlineMs: 30000, clients: [
    { ...APPROVAL_DIRECT, count: 18 }, { ...STATUS_DIRECT, count: 20 }
  ] });
  const shipped = await simulate({ latency: 93000, minutes: 20, clients: [
    { ...APPROVAL_DIRECT, count: 18 }, { ...STATUS_DIRECT, count: 20 }
  ] });
  assert.ok(shortDeadline.alivePerClient > shipped.alivePerClient * 1.2,
    'orphans per device: ' + shortDeadline.alivePerClient.toFixed(2) + ' at 30 s vs ' + shipped.alivePerClient.toFixed(2) + ' at 60 s');
});

test('fleet: a device that locks and wakes four times during one hang still holds ONE request (D2)', async () => {
  const clock = new Clock();
  const runs = [];
  const transport = loadTransport(clock, () => new Promise(() => { runs.push(clock.now); }));  // the server never answers
  const loop = transport.createPollLoop({ name: 'iphone', baseMs: 5000, maxMs: 20000,
    tick: () => transport.fetchJsonWithTimeout('poll', {}, 60000) });
  loop.start();
  await drain();
  for (let i = 0; i < 4; i++) {           // screen lock → wake, every 8 s
    await clock.advance(8000);
    loop.restartIfStuck();                // the visibilitychange rescue
    await drain();
  }
  assert.equal(runs.length, 1, 'the rescue refuses while a request is genuinely in flight');
  assert.equal(loop.isInFlight(), true);
  // The deadline is what ends this one: at 60 s the request is abandoned, the
  // tick fails, and the chain asks again by itself — still one at a time.
  await clock.advance(60000 - clock.now + 1);
  await drain();
  assert.equal(loop.isInFlight(), false);
  assert.equal(runs.length, 1, 'nothing new is sent until the backed-off wait elapses');
  await clock.advance(61000);
  assert.equal(runs.length, 2, 'exactly one request per cycle, however often the screen locks');
  loop.stop();
});

test('fleet: after a page freeze long enough to kill the chain, the rescue is what revives it', async () => {
  const clock = new Clock();
  const runs = [];
  // iOS can suspend the page mid-request and never settle the promise: no
  // timeout fires either, because its timer was frozen too.
  const transport = loadTransport(clock, () => new Promise(() => { runs.push(clock.now); }));
  const loop = transport.createPollLoop({ name: 'frozen', baseMs: 5000, maxMs: 20000,
    tick: () => new Promise(() => { runs.push(clock.now); }) });
  loop.start(); await drain();
  await clock.advance(10 * 60 * 1000);
  assert.equal(runs.length, 1, 'the chain is dead — this is the מחנה עמוס stall');
  assert.equal(loop.restartIfStuck(), true);
  await drain();
  assert.equal(runs.length, 2, 'and one wake brings the examinee their approval');
  loop.stop();
});

test('fleet: twenty devices released together do not re-arrive together', async () => {
  const clock = new Clock();
  const starts = [];
  const transport = loadTransport(clock, () => {
    starts.push(clock.now);
    return new Promise(resolve => clock.set(() => resolve({ ok: true, status: 200, text: () => Promise.resolve('{"status":"ok"}') }), 2000));
  });
  const loops = [];
  for (let i = 0; i < 20; i++) {
    loops.push(transport.createPollLoop({ name: 'd' + i, baseMs: 8000, maxMs: 20000,
      tick: () => transport.fetchJsonWithTimeout('poll', {}, 60000) }));
  }
  loops.forEach(loop => loop.start());          // one stall ends: all released at t=0
  await clock.advance(60000);
  loops.forEach(loop => loop.stop());
  const perSecond = new Map();
  for (const t of starts.slice(20)) {           // ignore the synchronised first wave
    const second = Math.floor(t / 1000);
    perSecond.set(second, (perSecond.get(second) || 0) + 1);
  }
  const busiest = Math.max(...perSecond.values());
  assert.ok(busiest <= 8, 'jitter spreads the fleet: busiest second had ' + busiest + ' of 20');
});

/**
 * A room of WAITING examinees against the session gateway, with and without the
 * hold. The fake Worker answers at once when this examinee's answer differs from
 * the fingerprint it was given (or when it was given none), and otherwise holds
 * the request until the answer changes or `wait` seconds elapse — the contract
 * the Worker implements.
 */
async function approvalFleet({ longPoll, ignoreHold = false, count = 40, minutes = 10, changeAtMs = 5 * 60 * 1000, latency = 300, baseMs = 3000 }) {
  const clock = new Clock();
  const requests = [];
  const live = new Map();               // client -> requests in flight right now
  let maxLive = 0;
  const answerAt = t => (t < changeAtMs ? 'waiting' : 'approved');

  const transport = loadTransport(clock, url => {
    const params = new URL(url).searchParams;
    const who = params.get('id');
    const wait = Number(params.get('wait') || 0) * 1000;
    const fp = params.get('fp') || '';
    const sent = clock.now;
    requests.push(sent);
    live.set(who, (live.get(who) || 0) + 1);
    maxLive = Math.max(maxLive, live.get(who));
    return new Promise(resolve => {
      const deliver = () => {
        live.set(who, live.get(who) - 1);
        const approval = answerAt(clock.now);
        const body = JSON.stringify({ status: 'ok', approval: approval, fp: approval, held: Math.max(0, clock.now - sent - latency) });
        resolve({ ok: true, status: 200, text: () => Promise.resolve(body) });
      };
      // the Worker holds only while its answer still matches the fingerprint it
      // was given — and an un-upgraded Worker (ignoreHold) never holds at all
      const holds = !ignoreHold && wait && fp && fp === answerAt(sent);
      const releaseAt = !holds ? sent : (changeAtMs > sent ? Math.min(sent + wait, changeAtMs) : sent + wait);
      clock.set(deliver, (releaseAt - sent) + latency);
    });
  });

  const clients = [];
  for (let i = 0; i < count; i++) {
    const me = { fp: '', sawApprovalAt: null };
    const loop = transport.createPollLoop({
      name: 'e' + i, baseMs: baseMs, maxMs: 20000,
      tick: () => {
        let url = 'https://gw.test/v1/poll?kind=approval&id=' + i;
        if (longPoll) {
          url += '&wait=' + transport.LONGPOLL_WAIT_SEC;
          if (me.fp) url += '&fp=' + me.fp;
        }
        return transport.fetchJsonWithTimeout(url, { cache: 'no-store' },
          longPoll ? transport.LONGPOLL_TIMEOUT_MS : transport.POLL_TIMEOUT_MS).then(data => {
            if (data.fp) me.fp = data.fp;
            if (data.approval === 'approved' && me.sawApprovalAt === null) me.sawApprovalAt = clock.now;
            return data;
          });
      }
    });
    me.loop = loop;
    clients.push(me);
  }
  clients.forEach((me, i) => clock.set(() => me.loop.start(), (i * 1000) % baseMs));   // a room does not arrive together
  await clock.advance(minutes * 60 * 1000);
  clients.forEach(me => me.loop.stop());
  return {
    requests: requests.length,
    perMinute: requests.length / minutes,
    maxLivePerClient: maxLive,
    approvalDelays: clients.map(me => (me.sawApprovalAt === null ? Infinity : me.sawApprovalAt - changeAtMs))
  };
}

test('fleet: long polling cuts a waiting room to a quarter of the requests and shows the decision in a second', async () => {
  // Today: 40 waiting examinees, one poll each every 3 s through the gateway.
  // The account's cap is 100,000 Worker requests/day and a heavy exam day is
  // already ~130k, so this is the number that has to come down.
  const today = await approvalFleet({ longPoll: false });
  const held = await approvalFleet({ longPoll: true });

  assert.ok(Math.abs(today.perMinute - 40 * 60 / 3.3) < 60,
    'the baseline is the shipped 3 s cadence: ' + today.perMinute.toFixed(0) + '/min');
  assert.ok(held.requests <= today.requests / 4,
    'long polling: ' + held.requests + ' requests where today sends ' + today.requests +
    ' (' + (today.requests / held.requests).toFixed(1) + 'x)');
  assert.equal(held.maxLivePerClient, 1, 'and still exactly one request in flight per device');

  const worst = Math.max(...held.approvalDelays);
  assert.ok(worst <= 1500, 'every examinee saw the approval within ' + worst + ' ms of the examiner\'s decision');
  const todayWorst = Math.max(...today.approvalDelays);
  assert.ok(worst < todayWorst, 'which is faster than a 3 s poll ever was (' + todayWorst + ' ms)');
});

test('fleet: against an un-upgraded Worker the new page costs exactly what today costs', async () => {
  // The page always offers the hold. A Worker that does not know wait/fp answers
  // immediately, every answer carries held = 0, and the cadence constants in
  // examinee.html pace the loop exactly as they do today — so the page can be
  // deployed before the Worker, or the Worker rolled back under it.
  const ignoring = await approvalFleet({ longPoll: true, ignoreHold: true, count: 20, minutes: 5 });
  const today = await approvalFleet({ longPoll: false, count: 20, minutes: 5 });
  assert.ok(Math.abs(ignoring.requests - today.requests) / today.requests < 0.1,
    'same cadence, same cost: ' + ignoring.requests + ' vs ' + today.requests);
  assert.equal(ignoring.maxLivePerClient, 1, 'and still one request in flight per device');
});

test('fleet: a slow-but-answering server (6 s) backs the fleet off instead of digging in', async () => {
  const fleet = await simulate({ latency: 7000, minutes: 15, clients: [
    { ...APPROVAL_DIRECT, count: 18 }, { ...STATUS_DIRECT, count: 20 }
  ] });
  for (const loop of fleet.loops) assert.ok(loop.currentDelayMs() > 8000, 'every device eased off: ' + loop.currentDelayMs());
  assert.ok(fleet.alivePerClient < 1);
  assert.ok(fleet.requestsPerMinute < 120, 'measured ' + fleet.requestsPerMinute.toFixed(0) + '/min');
});
