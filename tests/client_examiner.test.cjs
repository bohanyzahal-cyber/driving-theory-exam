// examiner.html + its service workers, executed for real.
//
// The pattern is the one tests/client_reliability.test.cjs established: cut a
// named section out of the page, run it in a vm context with a fake DOM and a
// fake clock, and drive it. Nothing here touches the network, the browser or
// any production data.
//
// What it gates (review ids from docs/reviews/2026-09-21/):
//   D1  an unattended poll must never log an examiner out
//   D17 the reset confirmation must warn about a result held on the device
//   D22 the combined-report probe must carry the 90 s deadline
//   D3/D4 the update check shows a banner AND reloads the page itself 60 s
//         later - but only when no dialog is open, no decision is waiting for
//         the server and nobody is typing the login form (21/09 message 20)
//   S3  the "not verified" badge keys on the stored marker, not on "0/"
//   plus: the dashboard loop never overlaps and honours the 2 s sync window,
//         the examiner bank grant and the gateway nudge after every decision
//         (which carries the decision itself, so the examinee sees it in <=2 s),
//         top-wrong rendering with and without a grant, and both SWs.
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const app = path.resolve(__dirname, '..');
const examiner = fs.readFileSync(path.join(app, 'examiner.html'), 'utf8');
const transportSrc = fs.readFileSync(path.join(app, 'shared', 'transport.js'), 'utf8');
const quiet = { log() {}, warn() {}, error() {} };

function section(src, start, end) {
  const i = src.indexOf(start), j = src.indexOf(end, i + start.length);
  assert.ok(i >= 0 && j > i, 'source section found: ' + start);
  return src.slice(i, j);
}
const deferred = () => { let resolve, reject; const promise = new Promise((a, b) => { resolve = a; reject = b; }); return { promise, resolve, reject }; };
async function drain() { for (let i = 0; i < 40; i++) await Promise.resolve(); }

// A fake clock with real setInterval semantics. It starts at a realistic epoch
// on purpose: several guards in the page compare `Date.now()` against a stored
// timestamp and treat 0 as "never", so a clock that starts at 0 hides them.
const EPOCH = 1758400000000;   // 2025-09-20T20:26:40Z — any plausible "now"
class Timers {
  now = EPOCH; nextId = 0; jobs = new Map();
  set = (cb, ms) => { const id = ++this.nextId; this.jobs.set(id, { cb, at: this.now + Number(ms) }); return id; };
  setInterval = (cb, ms) => { const id = ++this.nextId; this.jobs.set(id, { cb, at: this.now + Number(ms), every: Number(ms) }); return id; };
  clear = id => this.jobs.delete(id);
  async advance(ms) {
    const until = this.now + ms;
    for (let count = 0; count < 10000; count++) {
      const next = [...this.jobs].filter(([, job]) => job.at <= until).sort((a, b) => a[1].at - b[1].at)[0];
      if (!next) { this.now = until; await drain(); return; }
      const [id, job] = next;
      this.now = job.at;
      if (job.every) job.at = this.now + job.every; else this.jobs.delete(id);
      job.cb();
      await drain();
    }
    throw new Error('Unexpected timer loop');
  }
  // timers that are not repeating intervals — what "nothing is pending" means
  get pending() { return [...this.jobs.values()].filter(j => !j.every).length; }
}

// A DOM just rich enough for the banners, the toasts and the list renderers.
function dom() {
  const nodes = new Map();
  const created = [];
  const document = {
    activeElement: null,
    visibilityState: 'visible',
    getElementById: id => nodes.get(id) || null,
    querySelector: () => null,
    querySelectorAll: () => [],
    addEventListener() {}, removeEventListener() {}
  };
  function element(id, tag) {
    const el = {
      id, tagName: (tag || 'div').toUpperCase(), textContent: '', disabled: false, title: '',
      style: { cssText: '' }, handlers: {}, attrs: {}, children: [], parentNode: null,
      classList: {
        values: new Set(),
        add(v) { this.values.add(v); }, remove(v) { this.values.delete(v); }, contains(v) { return this.values.has(v); }
      },
      addEventListener(type, cb) { (this.handlers[type] = this.handlers[type] || []).push(cb); },
      setAttribute(k, v) { this.attrs[k] = v; },
      appendChild(child) { this.children.push(child); child.parentNode = this; if (child.id) nodes.set(child.id, child); return child; },
      removeChild(child) { this.children = this.children.filter(c => c !== child); if (child.id) nodes.delete(child.id); child.parentNode = null; },
      querySelector: () => null, querySelectorAll: () => [],
      focus() { document.activeElement = this; },
      click() { if (!this.disabled) (this.handlers.click || []).forEach(cb => cb()); }
    };
    let html = '';
    Object.defineProperty(el, 'innerHTML', {
      get: () => html,
      set(value) { html = value; el.children = []; }
    });
    if (id) nodes.set(id, el);
    return el;
  }
  document.createElement = tag => { const el = element('', tag); created.push(el); return el; };
  document.body = {
    children: [],
    appendChild(el) { this.children.push(el); el.parentNode = this; if (el.id) nodes.set(el.id, el); return el; },
    removeChild(el) { this.children = this.children.filter(c => c !== el); if (el.id) nodes.delete(el.id); el.parentNode = null; }
  };
  return { document, nodes, element, created };
}

function baseContext(extra = {}, timer = new Timers()) {
  const clockDate = class extends Date { static now() { return timer.now; } };
  const store = new Map();
  const localStorage = {
    getItem: k => (store.has(k) ? store.get(k) : null),
    setItem: (k, v) => store.set(k, String(v)),
    removeItem: k => store.delete(k)
  };
  const ctx = {
    console: quiet, Date: clockDate,
    setTimeout: timer.set, clearTimeout: timer.clear, setInterval: timer.setInterval, clearInterval: timer.clear,
    AbortController, Promise, JSON, Math, Object, Array, String, Number, RegExp, encodeURIComponent, decodeURIComponent,
    localStorage, sessionStorage: localStorage,
    ...extra
  };
  ctx.window = ctx;
  ctx.globalThis = ctx;
  vm.createContext(ctx);
  vm.runInContext(transportSrc, ctx);
  ctx.ExamTransport._resetHealth();
  return { ctx, timer, store };
}
const load = (ctx, code) => vm.runInContext(code, ctx);

// ---------------------------------------------------------------- D1
// Section: the API helper + the whole token-expiry state machine.
const apiSection = src => section(src, '  // ========== API Helper ==========', '  // ========== Login ==========');

function d1Context(answers) {
  const ui = dom();
  ui.element('screenLogin');
  const calls = [];
  const setup = baseContext({
    ...ui,
    API_URL: 'https://synthetic/exec', API_ORIGIN: 'examiner-app',
    examinerToken: 'T1', examinerData: { id: '111', name: 'synthetic' },
    stopDashboardPolling() { setup.ctx.pollingStopped = true; },
    showScreen(id) { setup.ctx.screen = id; },
    pollingStopped: false, screen: ''
  });
  setup.ctx.fetch = (url, opts) => {
    // GET carries the action in the query string, POST in the JSON body.
    const match = /action=([a-zA-Z]+)/.exec(url);
    const action = match ? match[1] : JSON.parse((opts && opts.body) || '{}').action;
    calls.push(action);
    const body = answers(action, calls.length);
    return Promise.resolve({ ok: true, text: () => Promise.resolve(JSON.stringify(body)) });
  };
  setup.ctx.localStorage.setItem('ext_examiner_remember', JSON.stringify({ id: '111', token: 'T1' }));
  load(setup.ctx, apiSection(examiner));
  return { ...setup, ...ui, calls };
}

test('D1: one tokenExpired on a dashboard poll asks for a second opinion instead of logging out', async () => {
  const { ctx, calls, nodes, store } = d1Context(action =>
    action === 'examinerDashboard' ? { status: 'error', tokenExpired: true }
      : { status: 'ok', examiner: { id: '111' } });
  await ctx.apiGet({ action: 'examinerDashboard', sessionCode: 'ABC' });
  await drain();
  assert.deepEqual(calls, ['examinerDashboard', 'verifyLogin'], 'exactly one re-verify');
  assert.equal(nodes.get('reauthBanner') || null, null, 'no banner after a single rumour');
  assert.ok(store.get('ext_examiner_remember'), 'remember-me survives');
  assert.equal(ctx.pollingStopped, false, 'polling is untouched');
  assert.equal(ctx.screen, '', 'nobody was sent to the login screen');
});

test('D1: a re-verify that confirms the expiry shows the banner and only then drops remember-me', async () => {
  const { ctx, calls, nodes, store } = d1Context(() => ({ status: 'error', tokenExpired: true }));
  await ctx.apiGet({ action: 'examinerDashboard', sessionCode: 'ABC' });
  await drain();
  assert.deepEqual(calls, ['examinerDashboard', 'verifyLogin']);
  const banner = nodes.get('reauthBanner');
  assert.ok(banner, 'in-page banner, not an alert()');
  assert.ok(banner.children.some(c => /פג תוקף ההתחברות/.test(c.textContent)),
    'banner says the login expired');
  assert.equal(store.get('ext_examiner_remember'), undefined, 'remember-me removed on the CONFIRMED path only');
  assert.equal(ctx.pollingStopped, true, 'polling stops once the token is known dead');
  assert.equal(ctx.screen, '', 'the login screen opens only when the examiner presses the button');
  nodes.get('reauthBtn').click();
  assert.equal(ctx.screen, 'screenLogin');
  assert.equal(ctx.sessionCode, undefined, 'the session was never cleared by this path');
});

test('D1: two consecutive tokenExpired polls confirm without waiting for the re-verify', async () => {
  // verifyLogin never answers, so only the second poll can confirm.
  const pending = deferred();
  const { ctx, nodes, store } = d1Context(action => {
    if (action === 'verifyLogin') return null;   // replaced below
    return { status: 'error', tokenExpired: true };
  });
  ctx.fetch = (url) => {
    const action = /action=([a-zA-Z]+)/.exec(url)[1];
    if (action === 'verifyLogin') return pending.promise;
    return Promise.resolve({ ok: true, text: () => Promise.resolve('{"status":"error","tokenExpired":true}') });
  };
  await ctx.apiGet({ action: 'examinerDashboard' });
  await drain();
  assert.equal(nodes.get('reauthBanner') || null, null, 'first strike alone proves nothing');
  await ctx.apiGet({ action: 'examinerDashboard' });
  await drain();
  assert.ok(nodes.get('reauthBanner'), 'the second consecutive answer confirms it');
  assert.equal(store.get('ext_examiner_remember'), undefined);
});

test('D1: a healthy answer in between clears the strike', async () => {
  let expired = true;
  const { ctx, nodes } = d1Context(action => {
    if (action === 'verifyLogin') return { status: 'ok' };
    return expired ? { status: 'error', tokenExpired: true } : { status: 'ok', pending: [] };
  });
  await ctx.apiGet({ action: 'examinerDashboard' });
  await drain();
  expired = false;
  await ctx.apiGet({ action: 'examinerDashboard' });
  await drain();
  expired = true;
  await ctx.apiGet({ action: 'examinerDashboard' });
  await drain();
  assert.equal(nodes.get('reauthBanner') || null, null, 'strikes are consecutive, not cumulative');
});

test('D1: login itself is never treated as an expiry', async () => {
  const { ctx, calls, nodes } = d1Context(() => ({ status: 'error', tokenExpired: true }));
  await ctx.apiPost({ action: 'login', idNumber: '1', password: 'x' });
  await drain();
  assert.deepEqual(calls, ['login'], 'no re-verify triggered by a failed login');
  assert.equal(nodes.get('reauthBanner') || null, null);
});

test('D1: timeouts and Google error pages never confirm an expiry', async () => {
  const { ctx, nodes, store } = d1Context(() => ({ status: 'ok' }));
  ctx.fetch = () => Promise.resolve({ ok: false, status: 503, text: () => Promise.resolve('<html>busy</html>') });
  await ctx.apiGet({ action: 'examinerDashboard' }).catch(() => {});
  await ctx.apiGet({ action: 'examinerDashboard' }).catch(() => {});
  await drain();
  assert.equal(nodes.get('reauthBanner') || null, null, 'a degraded backend is not an expired token');
  assert.ok(store.get('ext_examiner_remember'));
});

test('apiGet forwards its timeout (the commander-dashboard bug) and apiPost attaches credentials', async () => {
  let seenUrl = '', seenBody = null;
  const { ctx, timer } = d1Context(() => ({ status: 'ok' }));
  ctx.fetch = (url, opts) => {
    seenUrl = url;
    if (opts && opts.body) seenBody = JSON.parse(opts.body);
    return new Promise(() => {});                     // hang, so only the deadline settles it
  };
  const p = ctx.apiGet({ action: 'commanderDashboard' }, 90000).catch(e => e.name);
  await drain();
  await timer.advance(30000);
  assert.equal(timer.jobs.size, 1, 'still waiting at 30 s: the 90 s was forwarded');
  await timer.advance(60000);
  assert.equal(await p, 'TimeoutError');
  assert.match(seenUrl, /token=T1/, 'apiGet attaches the token');
  assert.match(seenUrl, /examinerId=111/);

  ctx.fetch = (url, opts) => { seenBody = JSON.parse(opts.body); return Promise.resolve({ ok: true, text: () => Promise.resolve('{"status":"ok"}') }); };
  await ctx.apiPost({ action: 'closeSession' });
  assert.equal(seenBody.token, 'T1');
  assert.equal(seenBody.examinerId, '111');
});

// ---------------------------------------------------------------- bank grant + gateway nudge
// The question texts are not in the repo any more: the gateway serves them
// against a signed grant, and every examiner decision has to tell the gateway
// to drop its snapshot of the session or the examinee waits for a stale copy.
const grantSection = src => section(src,
  '  // ========== Question-bank grant + gateway nudge ==========',
  '  // ========== Login ==========');

function grantContext(firstAnswer) {
  const ui = dom();
  const calls = [], posts = [];
  let answer = firstAnswer;
  const setup = baseContext({
    ...ui,
    sessionCode: 'ABC12345',
    apiGet(params) { calls.push(params.action); return Promise.resolve(answer(params)); },
    fetch(url, opts) { posts.push({ url, opts }); return Promise.resolve({ ok: true, text: () => Promise.resolve('{}') }); }
  });
  load(setup.ctx, grantSection(examiner));
  return { ...setup, ...ui, calls, posts, setAnswer(fn) { answer = fn; } };
}
const GRANT = { url: 'https://gateway.example', grant: 'payload.sig', exp: EPOCH + 8 * 3600 * 1000 };

test('the examiner grant is fetched once and kept for the next page load', async () => {
  const { ctx, calls, store } = grantContext(() => ({ status: 'ok', bank: GRANT }));
  const bank = await ctx.fetchBankGrant(true);
  assert.deepEqual(calls, ['bankGrant']);
  assert.equal(bank.grant, GRANT.grant);
  assert.deepEqual(JSON.parse(store.get('ext_examiner_bank')), GRANT, 'stored, so a reload costs nothing');
  await ctx.fetchBankGrant(true);
  assert.deepEqual(calls, ['bankGrant'], 'the grant in memory is reused');
});

test('a stored grant is adopted on a reload, and one about to expire is replaced', async () => {
  const stored = { url: 'https://gateway.example', grant: 'old', exp: EPOCH + 3600 * 1000 };
  const { ctx, calls, store } = grantContext(() => ({ status: 'ok', bank: GRANT }));
  store.set('ext_examiner_bank', JSON.stringify(stored));
  assert.equal((await ctx.fetchBankGrant()).grant, 'old');
  assert.deepEqual(calls, [], 'an auto-restore does not pay for a grant it already holds');

  ctx.examinerBank = null;
  store.set('ext_examiner_bank', JSON.stringify({ url: 'https://gateway.example', grant: 'old', exp: EPOCH + 5 * 60 * 1000 }));
  assert.equal((await ctx.fetchBankGrant()).grant, GRANT.grant, 'inside the 10-minute floor it is worth a request');
  assert.deepEqual(calls, ['bankGrant']);
});

test('a login on a device somebody else used does not inherit his grant', async () => {
  const { ctx, calls, store } = grantContext(() => ({ status: 'ok', bank: GRANT }));
  store.set('ext_examiner_bank', JSON.stringify({ url: 'https://gateway.example', grant: 'his', exp: EPOCH + 8 * 3600 * 1000 }));
  assert.equal((await ctx.fetchBankGrant(true)).grant, GRANT.grant, 'force=true at a fresh login');
  assert.deepEqual(calls, ['bankGrant']);
});

test('a grant that never arrives is not fatal - the next feature asks again', async () => {
  const { ctx, calls, setAnswer } = grantContext(() => ({ status: 'error', code: 'bank_not_configured' }));
  assert.equal(await ctx.fetchBankGrant(true), null, 'no throw, no banner, nothing breaks');
  setAnswer(() => ({ status: 'ok', bank: GRANT }));
  assert.equal((await ctx.ensureBankGrant()).grant, GRANT.grant);
  assert.deepEqual(calls, ['bankGrant', 'bankGrant']);
});

test('a successful decision nudges the gateway; a failed one does not', async () => {
  const { ctx, posts, setAnswer } = grantContext(() => ({ status: 'ok', bank: GRANT }));
  await ctx.fetchBankGrant(true);
  setAnswer(() => ({ status: 'ok' }));
  await ctx.examinerDecision({ action: 'approveExaminee', idNumber: '1' });
  assert.equal(posts.length, 1);
  assert.equal(posts[0].url, 'https://gateway.example/v1/invalidate?sessionCode=ABC12345&idNumber=1&status=approved',
    'the nudge carries the decision, so the next poll (<=2 s) already has it');
  assert.equal(posts[0].opts.method, 'POST');
  assert.equal(posts[0].opts.keepalive, true, 'the re-render that follows must not cancel it');
  assert.equal(posts[0].opts.cache, 'no-store');

  setAnswer(() => ({ status: 'error', message: 'busy' }));
  await ctx.examinerDecision({ action: 'approveExaminee', idNumber: '1' });
  assert.equal(posts.length, 1, 'nothing changed, so there is nothing to invalidate');
});

// The nudge may carry the decision because Apps Script has ALREADY confirmed the
// write by the time examinerDecision resolves; the gateway's own upstream read
// overwrites the patch with the same values within 2 s, so the patch can never
// be ahead of the truth - it only removes the wait.
test('the status table covers every examinee decision and leaves closeSession alone', () => {
  const { ctx } = grantContext(() => ({ status: 'ok' }));
  assert.deepEqual({ ...ctx.DECISION_STATUS }, {
    approveExaminee: 'approved',
    rejectExaminee: 'rejected',
    resetExaminee: 'cancelled',
    confirmDQ: 'dq_confirmed',
    overturnDQ: 'in_exam',
    forceComplete: 'completed',
    disqualify: 'disqualified',
    addExamTime: 'current'
  }, 'closeSession is absent on purpose - it is not about one examinee; a time grant keeps the row\'s CURRENT status');
});

test('the approval nudge carries the audio choice and the extended exam length', async () => {
  const { ctx, posts, setAnswer } = grantContext(() => ({ status: 'ok', bank: GRANT }));
  await ctx.fetchBankGrant(true);
  setAnswer(() => ({ status: 'ok' }));

  await ctx.examinerDecision({ action: 'approveExaminee', sessionCode: 'ABC12345', idNumber: '123456789', examinerId: '111', timeExtension: '1.25', audioMode: 'on' });
  assert.equal(posts[0].url,
    'https://gateway.example/v1/invalidate?sessionCode=ABC12345&idNumber=123456789&status=approved&examMinutes=50&audio=on',
    '+25% is the 50 minutes the server itself computes as round(40 * 1.25)');

  await ctx.examinerDecision({ action: 'approveExaminee', sessionCode: 'ABC12345', idNumber: '2', timeExtension: '1.5', audioMode: 'off' });
  assert.equal(posts[1].url,
    'https://gateway.example/v1/invalidate?sessionCode=ABC12345&idNumber=2&status=approved&examMinutes=60&audio=off',
    '+50% -> 60 minutes, audio explicitly off');

  await ctx.examinerDecision({ action: 'approveExaminee', sessionCode: 'ABC12345', idNumber: '3', audioMode: 'off' });
  assert.equal(posts[2].url,
    'https://gateway.example/v1/invalidate?sessionCode=ABC12345&idNumber=3&status=approved&audio=off',
    'no extension chosen -> no examMinutes, the examinee keeps the default 40');
});

test('every other decision nudges with its own status, addExamTime with the running total', async () => {
  const { ctx, posts, setAnswer } = grantContext(() => ({ status: 'ok', bank: GRANT }));
  await ctx.fetchBankGrant(true);
  setAnswer(() => ({ status: 'ok' }));
  const query = url => url.slice(url.indexOf('?') + 1);

  for (const [action, status] of [['rejectExaminee', 'rejected'], ['resetExaminee', 'cancelled'],
                                  ['confirmDQ', 'dq_confirmed'], ['overturnDQ', 'in_exam'],
                                  ['forceComplete', 'completed'], ['disqualify', 'disqualified']]) {
    posts.length = 0;
    await ctx.examinerDecision({ action: action, sessionCode: 'ABC12345', idNumber: '5', examinerId: '111' });
    assert.equal(query(posts[0].url), 'sessionCode=ABC12345&idNumber=5&status=' + status, action);
  }

  posts.length = 0;
  setAnswer(() => ({ status: 'ok', addedMinutes: 10, totalExtraMinutes: 25 }));
  await ctx.examinerDecision({ action: 'addExamTime', sessionCode: 'ABC12345', idNumber: '5', minutes: 10, reason: 'printer jam' },
                             undefined, { status: 'in_exam' });
  assert.equal(query(posts[0].url), 'sessionCode=ABC12345&idNumber=5&status=in_exam&extraMinutes=25',
    'the TOTAL the server just returned, not only the 10 minutes added now');

  posts.length = 0;
  setAnswer(() => ({ status: 'ok' }));
  await ctx.examinerDecision({ action: 'addExamTime', sessionCode: 'ABC12345', idNumber: '5', minutes: 10 },
                             undefined, { status: 'in_exam' });
  assert.equal(query(posts[0].url), 'sessionCode=ABC12345&idNumber=5&status=in_exam&extraMinutes=10',
    'an answer without a total falls back to what was just added');

  // A grant to an examinee who has NOT started must never be patched as in_exam:
  // the waiting screen treats "in_exam" as "your exam already started" and stops
  // polling for good. Without a running row the nudge is the plain drop.
  for (const hint of [{ status: 'waiting' }, { status: 'approved' }, undefined]) {
    posts.length = 0;
    setAnswer(() => ({ status: 'ok', addedMinutes: 10, totalExtraMinutes: 10 }));
    await ctx.examinerDecision({ action: 'addExamTime', sessionCode: 'ABC12345', idNumber: '5', minutes: 10 }, undefined, hint);
    assert.equal(query(posts[0].url), 'sessionCode=ABC12345', 'no status patch for ' + JSON.stringify(hint));
  }
});

test('closeSession is about the session, so its nudge names no examinee', async () => {
  const { ctx, posts, setAnswer } = grantContext(() => ({ status: 'ok', bank: GRANT }));
  await ctx.fetchBankGrant(true);
  setAnswer(() => ({ status: 'ok' }));
  await ctx.examinerDecision({ action: 'closeSession', sessionCode: 'ABC12345', examinerId: '111' });
  assert.equal(posts[0].url, 'https://gateway.example/v1/invalidate?sessionCode=ABC12345',
    'the plain drop it always was - no idNumber, no status');
});

test('a nudge whose fetch throws on the spot still resolves the decision', async () => {
  const { ctx, setAnswer } = grantContext(() => ({ status: 'ok', bank: GRANT }));
  await ctx.fetchBankGrant(true);
  setAnswer(() => ({ status: 'ok', addedMinutes: 5, totalExtraMinutes: 5 }));
  ctx.fetch = () => { throw new Error('blocked before it left the page'); };
  const result = await ctx.examinerDecision({ action: 'addExamTime', sessionCode: 'ABC12345', idNumber: '7', minutes: 5 })
    .then(d => d, e => 'rejected: ' + e.message);
  assert.equal(result.status, 'ok', 'a gateway that cannot be reached is never a failed decision');
  assert.equal(result.totalExtraMinutes, 5, 'the caller still gets the whole server answer');
});

test('a trailing slash on the gateway url does not become a double slash', async () => {
  const { ctx, posts, setAnswer } = grantContext(() => ({ status: 'ok', bank: { url: 'https://gateway.example/', grant: 'g', exp: EPOCH + 8 * 3600 * 1000 } }));
  await ctx.fetchBankGrant(true);
  setAnswer(() => ({ status: 'ok' }));
  await ctx.examinerDecision({ action: 'closeSession' });
  assert.equal(posts[0].url, 'https://gateway.example/v1/invalidate?sessionCode=ABC12345');
});

test('without a grant a decision still works - it just does not nudge', async () => {
  const { ctx, posts } = grantContext(() => ({ status: 'ok' }));
  const data = await ctx.examinerDecision({ action: 'disqualify', idNumber: '1' });
  assert.equal(data.status, 'ok');
  assert.equal(posts.length, 0, 'no gateway url, no POST into the void');
});

test('a decision is counted while it is in flight, and the count survives a rejection', async () => {
  const pending = deferred();
  const { ctx, setAnswer } = grantContext(() => pending.promise);
  assert.equal(ctx.decisionsInFlight, 0);
  const first = ctx.examinerDecision({ action: 'forceComplete' });
  assert.equal(ctx.decisionsInFlight, 1, 'the self-reload must be able to see this');
  pending.resolve({ status: 'error' });
  await first;
  assert.equal(ctx.decisionsInFlight, 0);

  setAnswer(() => Promise.reject(new Error('network down')));
  const failed = await ctx.examinerDecision({ action: 'forceComplete' }).then(() => 'resolved', e => e.message);
  assert.equal(failed, 'network down', 'the caller still sees its own failure');
  assert.equal(ctx.decisionsInFlight, 0, 'a rejection must not pin the counter at 1 forever');
});

test('every examiner decision goes through the wrapper (and so nudges the gateway)', () => {
  const wrapped = ['rejectExaminee', 'resetExaminee', 'confirmDQ', 'overturnDQ',
                   'forceComplete', 'disqualify', 'addExamTime', 'closeSession'];
  const offenders = examiner.split('\r\n').filter(l =>
    /(^|[^a-zA-Z])apiGet\(\{ action: '(approveExaminee|rejectExaminee|resetExaminee|confirmDQ|overturnDQ|forceComplete|disqualify|addExamTime|closeSession)'/.test(l));
  assert.deepEqual(offenders, [], 'a decision that skips examinerDecision leaves the examinee on a stale snapshot');
  for (const action of wrapped) {
    assert.ok(new RegExp("examinerDecision\\(\\{ action: '" + action + "'").test(examiner), action + ' is wrapped');
  }
  // approve builds its params object first (time extension, audio), then calls
  assert.match(section(examiner, "var params = { action: 'approveExaminee'", 'actions.appendChild(approveBtn);'),
    /examinerDecision\(params\)/, 'approveExaminee too');
});

test('the grant is fetched on all three ways in, and leaves with the examiner', () => {
  assert.match(section(examiner, '  // ========== Login ==========', '  window.logout = function()'),
    /fetchBankGrant\(true\);/, 'a fresh login');
  assert.match(section(examiner, '  function enterAsRemembered(creds, data) {', '  // ========== Auto-restore on page load =========='),
    /fetchBankGrant\(\);/, 'the remembered login');
  assert.match(section(examiner, '  // ========== Auto-restore on page load ==========', '  function showQR(code)'),
    /fetchBankGrant\(\);/, 'the auto-restore path');
  assert.match(examiner, /localStorage\.removeItem\('ext_examiner_bank'\)/, 'and logout takes it away');
});

// ---------------------------------------------------------------- dashboard loop
function dashboardContext(apiGet) {
  const ui = dom();
  ui.element('offlineBanner');
  const setup = baseContext({
    ...ui,
    sessionCode: 'TEST00', examinerToken: 'synthetic', failedPolls: 0,
    dashboardInterval: null, countdownInterval: null, apiGet,
    POLL_TIMEOUT_MS: 60000,
    updatePendingList() {}, updateActiveList() {}, updateCompletedList() {}
  });
  setup.ctx.isBackendDegraded = () => setup.ctx.ExamTransport.isBackendDegraded();
  load(setup.ctx, section(examiner, '  // ===== Dashboard polling =====', '  // ===== toast ====='));
  load(setup.ctx, section(examiner, '  var OFFLINE_BANNER_TEXT', '  // Text of the "'));
  setup.ctx.ExamTransport._setJitter(ms => ms);          // exact fake clock
  return { ...setup, ...ui };
}

test('dashboard: a failing poll raises the banner, a good one clears it, and the loop never overlaps', async () => {
  let response = { status: 'error', message: 'server busy' };
  let inFlight = 0, maxInFlight = 0;
  const { ctx, timer, nodes } = dashboardContext(() => {
    inFlight++; maxInFlight = Math.max(maxInFlight, inFlight);
    return Promise.resolve(response).then(r => { inFlight--; return r; });
  });
  ctx.startDashboardPolling();
  await drain();
  for (let i = 0; i < 3; i++) await timer.advance(20000);
  assert.ok(ctx.failedPolls >= 3);
  assert.equal(nodes.get('offlineBanner').classList.contains('show'), true);
  assert.equal(maxInFlight, 1, 'answer -> wait -> ask again; never two chains');
  response = { status: 'ok', pending: [], active: [], completed: [] };
  await timer.advance(60000);
  assert.equal(ctx.failedPolls, 0);
  assert.equal(nodes.get('offlineBanner').classList.contains('show'), false);
  ctx.stopDashboardPolling();
  assert.equal(timer.pending, 0, 'stopping leaves no timer behind');
});

test('dashboard: the 2 s sync cadence applies only while healthy and only for 30 s', async () => {
  const syncing = { status: 'ok', pending: [], completed: [], active: [{ idNumber: 'A', finishedOnDevice: true }] };
  const settled = { status: 'ok', pending: [], completed: [], active: [{ idNumber: 'A', finishedOnDevice: false }] };
  let response = settled;
  const { ctx, timer } = dashboardContext(() => Promise.resolve(response));
  ctx.startDashboardPolling();
  await drain();
  assert.equal(ctx.dashPollDelayMs, 5000, 'idle dashboards stay at 5 s');
  response = syncing;
  await timer.advance(5000);
  assert.equal(ctx.dashSyncSince, timer.now, 'the syncing stretch is stamped');
  const syncStarted = timer.now;
  await timer.advance(2000);
  assert.equal(ctx.dashPollDelayMs, 2000, 'a syncing result pulls the next tick in to 2 s');
  // past the 30 s window the stuck row must not keep the dashboard at 2 s
  await timer.advance(35000);
  assert.ok(timer.now - syncStarted > 30000);
  assert.equal(ctx.dashPollDelayMs, 5000, 'the fast cadence gives up after its 30 s window');
  // and a syncing row while the server is failing never wins over the backoff
  ctx.dashSyncSince = ctx.Date.now();
  response = { status: 'error', message: 'busy' };
  await timer.advance(10000);
  assert.ok(ctx.dashPollDelayMs > 5000, 'the backoff outranks the sync cadence');
  ctx.stopDashboardPolling();
});

test('dashboard: a manual refresh shares the in-flight request, and a session switch discards the answer', async () => {
  const response = deferred();
  let calls = 0, renders = 0;
  const { ctx } = dashboardContext(() => { calls++; return response.promise; });
  ctx.updatePendingList = () => renders++;
  const first = ctx.pollDashboard(), second = ctx.pollDashboard();
  assert.equal(first, second, 'one promise for both');
  assert.equal(calls, 1);
  ctx.sessionCode = 'TEST01';
  response.resolve({ status: 'ok', pending: [] });
  await first;
  assert.equal(renders, 0, 'the answer for the old session is dropped');
});

test('dashboard: a fast JSON error still backs the loop off', async () => {
  const { ctx, timer } = dashboardContext(() => Promise.resolve({ status: 'error', message: 'busy' }));
  ctx.startDashboardPolling();
  await drain();
  const first = ctx.dashboardLoop.currentDelayMs();
  await timer.advance(20000);
  assert.ok(ctx.dashboardLoop.currentDelayMs() > first,
    'a failure that answered in 1 ms must pace like a failure, not like a healthy poll');
  ctx.stopDashboardPolling();
});

// ---------------------------------------------------------------- S3
test('S3: the "not verified" badge keys on the stored marker, not on a 0/ score', () => {
  const ui = dom();
  const { ctx } = baseContext({ ...ui });
  load(ctx, section(examiner, '  var FABRICATED_FAIL_MARKERS', '  // ========== Session persistence =========='));
  // the rule, lifted verbatim out of updateCompletedList
  ctx.shows = function (r) {
    const isDQ = r.disqualified === true || r.disqualified === 'TRUE' || r.passed === 'פסול';
    const verifiedKnown = (typeof r.verified !== 'undefined');
    const vstate = String(r.verified || '');
    const isFabricatedFail = ctx.FABRICATED_FAIL_MARKERS.test(String(r.wrongDetails || ''));
    return !!(verifiedKnown && !isDQ && !isFabricatedFail && vstate !== 'מאומת' && vstate !== 'ידני');
  };
  const rows = [
    ['a real fail the server could not re-score', { score: '24/30', verified: '', wrongDetails: 'מזהה שאלה: 14' }, true],
    ['a 0/30 the server could not re-score (the 03/06 pattern)', { score: '0/30', verified: '', wrongDetails: '⚠️ ציון לא אומת בשרת' }, true],
    ['browser close', { score: '0/30', verified: '', wrongDetails: 'סגירת דפדפן' }, false],
    ['timeout', { score: '0/30', verified: '', wrongDetails: 'טיימאאוט' }, false],
    ['manual end', { score: '0/30', verified: '', wrongDetails: 'סיום ידני' }, false],
    ['a paper entry', { score: '27/30', verified: 'ידני', wrongDetails: '' }, false],
    ['a disqualification', { score: '0/30', verified: '', disqualified: true, wrongDetails: '' }, false],
    ['a verified pass', { score: '28/30', verified: 'מאומת', wrongDetails: '' }, false],
    ['an old payload without the column', { score: '24/30', wrongDetails: '' }, false]
  ];
  for (const [name, row, expected] of rows) {
    assert.equal(ctx.shows(row), expected, name);
  }
});

// ---------------------------------------------------------------- D17
test('D17: the reset confirmation warns when the device is holding a result', () => {
  const ui = dom();
  const { ctx } = baseContext({ ...ui });
  load(ctx, section(examiner, '  // Text of the "', '  // ========== Pending list =========='));
  const plain = ctx.resetConfirmText({ name: 'A', finishedOnDevice: false });
  const holding = ctx.resetConfirmText({ name: 'A', finishedOnDevice: true });
  assert.ok(!/בלתי ניתנת לשליחה/.test(plain), 'no scare text when there is nothing to lose');
  assert.match(holding, /הנבחן סימן שסיים במכשיר/);
  assert.match(holding, /בלתי ניתנת לשליחה/);
  assert.ok(holding.endsWith(plain), 'the original question is still asked');
});

// ---------------------------------------------------------------- D22
test('D22: the combined-report probe carries the 90 s deadline, throttles, and never hides on a failure', async () => {
  const ui = dom();
  const btn = ui.element('siteCombinedBtn');
  btn.style.display = 'none';
  const timeouts = [];
  const { ctx, timer } = baseContext({
    ...ui,
    sessionCode: 'S1', examinerData: { id: '9' }, examinerToken: 'T',
    HEAVY_REPORT_TIMEOUT_MS: 90000,
    apiGet(params, timeoutMs) { timeouts.push(timeoutMs); return ctx.nextAnswer(); },
    nextAnswer: () => Promise.resolve({ status: 'ok', sessions: [{}, {}] })
  });
  load(ctx, section(examiner, '  var SITE_COMBINED_PROBE_MIN_MS', '  // Build and open the combined-site report'));
  ctx.checkSiteCombinedAvailability();
  await drain();
  assert.deepEqual(timeouts, [90000], 'the 18-81 s report is not probed on the 30 s default');
  assert.equal(btn.style.display, '', 'two sessions at the site: the button appears');
  ctx.checkSiteCombinedAvailability();
  await drain();
  assert.equal(timeouts.length, 1, 'throttled to one probe per 10 minutes');
  await timer.advance(10 * 60 * 1000 + 1);
  ctx.nextAnswer = () => Promise.reject(new Error('timeout'));
  ctx.checkSiteCombinedAvailability();
  await drain();
  assert.equal(timeouts.length, 2);
  assert.equal(btn.style.display, '', 'a transient failure says nothing about the site');
});

// ---------------------------------------------------------------- D3/D4
// 21/09 message 20 reverses decision 4: the banner and its button stay, and the
// page also reloads ITSELF 60 s later - but only at a safe moment. The section
// now carries those safety rules, so the context has to supply what they read
// out of the page's closure (saveState, the in-flight decision counter).
const updateSection = src => section(src,
  '  // ===== Auto-update: a banner, a button, and a self-reload ONLY when it is safe',
  '  var deferredInstallPrompt = null;');

function updateContext(nextAnswer) {
  const ui = dom();
  const setup = baseContext({
    ...ui,
    sessionCode: '',
    reloads: 0, saved: 0, decisionsInFlight: 0,
    saveState() { setup.ctx.saved++; },
    location: { pathname: '/examiner.html', reload() { setup.ctx.reloads++; } },
    fetch: () => Promise.resolve(nextAnswer())
  });
  load(setup.ctx, updateSection(examiner));
  return { ...setup, ...ui };
}
const versionAnswer = v => ({ ok: true, text: () => Promise.resolve(JSON.stringify(v)) });

test('D3/D4: two sightings raise the banner, and 60 s later the page reloads itself', async () => {
  let version = { build: 'b1', pages: { 'examiner.html': 'hash-1' } };
  const { ctx, timer, nodes } = updateContext(() => versionAnswer(version));
  await drain();
  version = { build: 'b2', pages: { 'examiner.html': 'hash-2' } };
  await timer.advance(120000);                  // first sighting: not enough
  assert.equal(nodes.get('examinerUpdateBanner') || null, null);
  await timer.advance(120000);                  // the same new hash twice = a real deploy
  assert.ok(nodes.get('examinerUpdateBanner'), 'the banner appears');
  await timer.advance(59000);
  assert.equal(ctx.reloads, 0, 'the examiner gets his minute first');
  await timer.advance(2000);
  assert.equal(ctx.reloads, 1, 'and then the page updates itself (message 20)');
  assert.equal(ctx.saved, 1, 'saveState ran first, so he comes back to the same screen');
  await timer.advance(10 * 60 * 1000);
  assert.equal(ctx.reloads, 1, 'exactly once');
});

test('D3/D4: an open modal postpones the reload, and it happens 15 s after the modal closes', async () => {
  let version = { build: 'b1', pages: { 'examiner.html': 'h1' } };
  const { ctx, timer, nodes, element } = updateContext(() => versionAnswer(version));
  const modal = element('settingsModal');
  modal.classList.add('show');
  await drain();
  version = { build: 'b2', pages: { 'examiner.html': 'h2' } };
  await timer.advance(120000); await timer.advance(120000);
  assert.ok(nodes.get('examinerUpdateBanner'));
  await timer.advance(60000);
  assert.equal(ctx.reloads, 0, 'never wipe an open dialog');
  await timer.advance(5 * 60 * 1000);
  assert.equal(ctx.reloads, 0, 'and it keeps waiting for as long as the dialog is open');
  modal.classList.remove('show');
  await timer.advance(15000);
  assert.equal(ctx.reloads, 1, 'D4: the guard is re-evaluated every 15 s, not once before the timer');
  assert.equal(ctx.saved, 1);
});

test('D3/D4: the share dialog, the add-time overlay, a decision in flight and a half-typed login all hold the reload', async () => {
  let version = { build: 'b1', pages: { 'examiner.html': 'h1' } };
  const { ctx, timer, nodes, element, document } = updateContext(() => versionAnswer(version));
  const share = element('shareDialog');
  share.style.display = 'flex';
  await drain();
  version = { build: 'b2', pages: { 'examiner.html': 'h2' } };
  await timer.advance(120000); await timer.advance(120000);
  await timer.advance(60000);
  assert.equal(ctx.reloads, 0, 'a share dialog is open');

  // the overlay the add-time modal builds at click time has no id, only the class
  share.style.display = 'none';
  const overlay = { className: 'examiner-modal' };
  document.querySelector = sel => (sel === '.examiner-modal' ? overlay : null);
  await timer.advance(15000);
  assert.equal(ctx.reloads, 0, 'an overlay built at click time counts too');

  document.querySelector = () => null;
  ctx.decisionsInFlight = 1;
  await timer.advance(15000);
  assert.equal(ctx.reloads, 0, 'an approval the server has not answered yet');

  ctx.decisionsInFlight = 0;
  const loginScreen = element('screenLogin');
  loginScreen.classList.add('active');           // he is looking at the login form
  const loginId = element('loginId');
  loginId.value = '12345';
  await timer.advance(15000);
  assert.equal(ctx.reloads, 0, 'nine digits already typed must not be wiped');

  loginId.value = '';
  await timer.advance(15000);
  assert.equal(ctx.reloads, 1, 'once everything is clear it finally reloads');
});

test('D3/D4: values left in the login form of a logged-in examiner do not block the reload forever', async () => {
  // Nothing clears loginId/loginPass on a successful login, and the remembered
  // path pre-fills the id - so reading them while the dashboard is on screen
  // would have pinned isSafeToReload() at false for the rest of the day.
  let version = { build: 'b1', pages: { 'examiner.html': 'h1' } };
  const { ctx, timer, nodes, element } = updateContext(() => versionAnswer(version));
  const loginScreen = element('screenLogin');
  element('screenSetup').classList.add('active');   // he is inside, on the setup screen
  const loginId = element('loginId'), loginPass = element('loginPass');
  loginId.value = '123456789';
  loginPass.value = 'still here from the login';
  await drain();
  version = { build: 'b2', pages: { 'examiner.html': 'h2' } };
  await timer.advance(120000); await timer.advance(120000);
  assert.ok(nodes.get('examinerUpdateBanner'));
  await timer.advance(61000);
  assert.equal(ctx.reloads, 1, 'stale values in a hidden section are not "somebody typing"');

  assert.equal(loginScreen.classList.contains('active'), false, 'the login screen really was not the active one');
});

test('D3/D4: the button reloads at once, whatever is open', async () => {
  let version = { build: 'b1', pages: { 'examiner.html': 'h1' } };
  const { ctx, timer, nodes, element } = updateContext(() => versionAnswer(version));
  element('settingsModal').classList.add('show');
  await drain();
  version = { build: 'b2', pages: { 'examiner.html': 'h2' } };
  await timer.advance(120000); await timer.advance(120000);
  nodes.get('swUpdNow').click();
  assert.equal(ctx.reloads, 1, 'the examiner asked for it himself');
});

test('D3/D4: a rewritten header or an error page is not a new version', async () => {
  let answer = versionAnswer({ build: 'b1', pages: { 'examiner.html': 'hash-1' } });
  const { ctx, timer, nodes } = updateContext(() => answer);
  ctx.location.reload = () => { throw new Error('must not reload'); };
  await drain();
  answer = { ok: true, text: () => Promise.resolve('<html>captive portal</html>') };
  await timer.advance(120000); await timer.advance(120000);
  answer = { ok: false, status: 503, text: () => Promise.resolve('{}') };
  await timer.advance(120000); await timer.advance(120000);
  assert.equal(nodes.get('examinerUpdateBanner') || null, null,
    'only a 200 JSON with a hash for THIS page counts');
});

// ---------------------------------------------------------------- top wrong
// The table gets { questionId, count, category, text } from the server and
// nothing else. The canonical text and the image are pulled from the gateway
// for exactly the ids on screen, with the examiner grant; the stored text is
// what shows until they arrive, and for good when they never do.
function topWrongContext(arriving, grant) {
  const ui = dom();
  ui.element('cmdTopWrong');
  const bankEntries = {};
  const loadIdsCalls = [];
  const held = { grant: grant === undefined ? { url: 'https://gateway.example', grant: 'g' } : grant };
  const { ctx } = baseContext({
    ...ui,
    cmdData: null,
    escapeHtml: s => String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'),
    ensureBankGrant: () => Promise.resolve(held.grant),
    QuestionBank: {
      get: id => bankEntries[id] || null,
      imageUrl: e => (e && e.image ? 'images/' + e.image : ''),
      loadIds(bank, ids, langs) {
        loadIdsCalls.push({ bank, ids, langs });
        ids.forEach(id => { if (arriving[id]) bankEntries[id] = arriving[id]; });
        return Promise.resolve({ build: 'b', count: ids.length, missing: [] });
      }
    }
  });
  load(ctx, section(examiner, '  // The server sends { questionId, count, category, text }', '  // Day-of-week'));
  return { ctx, loadIdsCalls, held, ...ui };
}

test('top wrong: the texts are pulled from the gateway for exactly the ids on screen', async () => {
  const { ctx, nodes, loadIdsCalls } = topWrongContext({
    14: { id: 14, text: 'מה פירוש התמרור?', image: 'TQ_PIC_14.jpg' },
    21: { id: 21, text: 'שאלה 21', image: '' }
  });
  const list = [{ questionId: 14, count: 9, category: 'תמרורים', text: 'stale copy from the sheet' },
                { questionId: 21, count: 3, category: '', text: 'stale 21' }];
  ctx.cmdData = { topWrong: list };
  ctx.renderTopWrong(list);
  assert.match(nodes.get('cmdTopWrong').innerHTML, /stale copy/, 'the first paint uses what the server stored');
  await drain();
  assert.equal(loadIdsCalls.length, 1, 'one request for the whole table');
  // Array.from: these were built inside the vm, so their prototype is not ours
  assert.deepEqual(Array.from(loadIdsCalls[0].ids), [14, 21], 'exactly the ids being shown');
  assert.deepEqual(Array.from(loadIdsCalls[0].langs), ['he']);
  assert.equal(loadIdsCalls[0].bank.grant, 'g', 'with the examiner grant');
  const html = nodes.get('cmdTopWrong').innerHTML;
  assert.match(html, /images\/TQ_PIC_14\.jpg/, 'the image name comes from the bank, not from the server');
  assert.match(html, /מה פירוש התמרור\?/, 'and the canonical text wins');
  assert.ok(!/stale copy/.test(html));
  assert.match(html, /id: 14/);
  assert.match(html, /תמרורים/);
});

test('top wrong: no grant means the stored text stays, nothing is requested, and the ids are not burnt', async () => {
  const { ctx, nodes, loadIdsCalls, held } = topWrongContext({ 9999: { id: 9999, text: 'מהמאגר החי', image: '' } }, null);
  const list = [{ questionId: 9999, count: 4, category: '', text: 'שאלה שהוצאה מהמאגר' }];
  ctx.cmdData = { topWrong: list };
  ctx.renderTopWrong(list);
  await drain();
  assert.equal(loadIdsCalls.length, 0, 'an examiner whose grant never arrived still sees the table');
  const html = nodes.get('cmdTopWrong').innerHTML;
  assert.match(html, /שאלה שהוצאה מהמאגר/);
  assert.ok(!/<img/.test(html), 'no broken image box');
  assert.match(html, /find_image\.html\?q=/, 'the deep link still works');

  // the grant turns up later (the lazy retry): the next report asks for real
  held.grant = { url: 'https://gateway.example', grant: 'g' };
  ctx.renderTopWrong(list);
  await drain();
  assert.equal(loadIdsCalls.length, 1, 'ids are only marked as asked once they really were');
  assert.match(nodes.get('cmdTopWrong').innerHTML, /מהמאגר החי/);
});

test('top wrong: an id the gateway does not carry is asked for once, not on every repaint', async () => {
  const { ctx, nodes, loadIdsCalls } = topWrongContext({});   // the gateway answers, but has nothing for it
  const list = [{ questionId: 9999, count: 4, category: '', text: 'שאלה שהוצאה מהמאגר' }];
  ctx.cmdData = { topWrong: list };
  ctx.renderTopWrong(list);
  await drain();
  assert.equal(loadIdsCalls.length, 1);
  await drain();
  assert.equal(loadIdsCalls.length, 1, 'the repaint must not start the request again');
  assert.match(nodes.get('cmdTopWrong').innerHTML, /שאלה שהוצאה מהמאגר/);
});

test('top wrong: an empty list says so instead of rendering an empty table', () => {
  const { ctx, nodes, loadIdsCalls } = topWrongContext({});
  ctx.renderTopWrong([]);
  assert.match(nodes.get('cmdTopWrong').innerHTML, /אין נתוני/);
  assert.equal(loadIdsCalls.length, 0);
});

// ---------------------------------------------------------------- toasts
test('the dashboard reports through a toast, never through a blocking dialog', async () => {
  const ui = dom();
  const { ctx, timer } = baseContext({ ...ui });
  load(ctx, section(examiner, '  // ===== toast =====', '  var OFFLINE_BANNER_TEXT'));
  ctx.showToast('אופס');
  ctx.toastError('שגיאה');
  const host = ui.nodes.get('examinerToasts');
  assert.equal(host.children.length, 2);
  await timer.advance(6500);
  assert.equal(host.children.length, 1, 'the error toast stays longer than the notice');
  await timer.advance(3000);
  assert.equal(host.children.length, 0, 'toasts clean themselves up');
});

test('no blocking alert() survives in the dashboard list renderers', () => {
  const region = section(examiner, '  // ========== Pending list ==========', '  // ========== Edit examinee phone ==========');
  const offenders = region.split('\n')
    .filter(l => /(^|[^a-zA-Z.])alert\(/.test(l))
    .filter(l => !/H\.push/.test(l))
    .filter(l => !/לא הצלחתי לפתוח חלון/.test(l));  // popup blocker: immediate click feedback
  assert.deepEqual(offenders, [], 'every answer-driven dialog is a toast now');
  assert.ok(/confirm\(resetConfirmText\(item\)\)/.test(region), 'confirm() stays for the destructive click');
});

// ---------------------------------------------------------------- removals
test('the generated report pages keep their own dialogs (they have no toast host)', () => {
  // Caught during the rebuild: the alert-to-toast sweep also rewrote an alert()
  // inside an H.push() string — code that runs in a report window the examiner
  // opens, where showToast does not exist. That would be a ReferenceError the
  // moment somebody pressed "copy".
  const generated = examiner.split('\n').filter(l => l.indexOf('H.push') >= 0);
  const offenders = generated.filter(l => /showToast\(|toastError\(/.test(l));
  assert.deepEqual(offenders, [], 'a generated report cannot call the dashboard’s toast');
});

test('the retired code really is gone', () => {
  for (const dead of ['sendBulkToManager', 'searchQuestions', 'predictiveModelPreview',
                      'getExamQuestions', 'registerExamQuestions']) {
    assert.ok(examiner.indexOf(dead) < 0, dead + ' must not appear in examiner.html');
  }
  assert.ok(examiner.indexOf("<script src=\"shared/transport.js\"></script>") >= 0);
  assert.ok(examiner.indexOf("<script src=\"shared/bank.js\"></script>") >= 0);
});

// ---------------------------------------------------------------- service workers
for (const sw of ['sw-examiner.js', 'sw-teacher.js']) {
  test(sw + ': parses, is GET-only, precaches the shared modules and keeps its build-written cache name', () => {
    const src = fs.readFileSync(path.join(app, sw), 'utf8');
    const listeners = [];
    const self = { addEventListener: (t, cb) => listeners.push([t, cb]), skipWaiting() {}, clients: { claim() {} } };
    vm.runInNewContext(src, { self, caches: { open: () => Promise.resolve({ addAll: () => Promise.resolve() }), keys: () => Promise.resolve([]), match: () => Promise.resolve(null), delete: () => Promise.resolve() }, fetch: () => Promise.resolve(), console: quiet });
    assert.deepEqual(listeners.map(l => l[0]), ['install', 'activate', 'fetch']);

    assert.match(src, /^var CACHE_NAME = '[a-z]+-[a-z0-9]+';$/m,
      'the build rewrites this exact line (tools/build_version.js)');
    assert.match(src, /'\.\/shared\/transport\.js'/);
    assert.match(src, /'\.\/shared\/bank\.js'/);

    // D7: a non-GET request must leave the worker before anything touches a cache
    const fetchHandler = listeners.find(l => l[0] === 'fetch')[1];
    let responded = false;
    fetchHandler({ request: { method: 'POST', url: 'https://example/x' }, respondWith: () => { responded = true; } });
    assert.equal(responded, false, 'POST/HEAD are never intercepted (Cache.put throws on them)');
    fetchHandler({ request: { method: 'HEAD', url: 'https://example/x' }, respondWith: () => { responded = true; } });
    assert.equal(responded, false);
    fetchHandler({ request: { method: 'GET', url: 'https://script.google.com/macros/s/x/exec' }, respondWith: () => { responded = true; } });
    assert.equal(responded, false, 'the API always goes to the network');
    fetchHandler({ request: { method: 'GET', url: 'https://example/examiner.html?cb=1' }, respondWith: () => { responded = true; } });
    assert.equal(responded, true, 'same-origin GETs are served network-first');
    assert.match(src, /ignoreSearch: true/, 'a cache-busted shell must still match its cached copy offline');

    // The question bank is not served from this origin any more (it lives
    // behind the gateway), so nothing here may special-case it or precache it.
    assert.ok(!/bank\/manifest\.json/.test(src), 'no bank manifest in the shell');
    assert.ok(!/'\.\/bank\//.test(src), 'no bank file in the shell');
  });
}
