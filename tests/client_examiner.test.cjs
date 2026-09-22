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
//   r31 the dashboard has NO cadence (DESIGN §13.2): a request held open at the
//       Worker says when the session changed, a changed fingerprint reads
//       examinerDashboard exactly once, a safety net catches the rest, and a
//       gateway that is not there puts the page back on its pre-r31 5 s retry
//   r31 the cold actions (reports, commander, forecast) are routed to the second
//       Apps Script deployment by name (DESIGN §13.3)
//   r32 the watch brings the DATA (DESIGN §14.1): an answer carrying the session
//       paints the three lists with no Google read at all, by the server's own
//       dedup rules; examinerDashboard is left for the wrong-answer blocks the
//       snapshot cannot carry (one read, on a click) and as the fallback for an
//       old Worker, an old server and a stale copy. Plus: a write to תוצאות that
//       skips examinerDecision still announces itself to the gateway
//   plus: the examiner bank grant and the gateway nudge after every decision
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

// urls: { api, reports } - the two Apps Script deployments of r31. They are the
// same string in production until Yossi creates the second one; the routing
// test below drives them apart so "which deployment" is observable.
function d1Context(answers, urls = {}) {
  const ui = dom();
  ui.element('screenLogin');
  const calls = [], sent = [];
  const setup = baseContext({
    ...ui,
    API_URL: urls.api || 'https://synthetic/exec',
    REPORTS_API_URL: urls.reports || urls.api || 'https://synthetic/exec',
    API_ORIGIN: 'examiner-app',
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
    sent.push(url);
    const body = answers(action, calls.length);
    return Promise.resolve({ ok: true, text: () => Promise.resolve(JSON.stringify(body)) });
  };
  setup.ctx.localStorage.setItem('ext_examiner_remember', JSON.stringify({ id: '111', token: 'T1' }));
  load(setup.ctx, apiSection(examiner));
  return { ...setup, ...ui, calls, sent, lastUrl: () => sent[sent.length - 1] };
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

// ---------------------------------------------------------------- r31 routing
// Google loads and compiles the whole script on every request, so since r31 the
// server is two deployments and the cold half (reports, commander, forecast,
// teachers, practice) is not carried by the exam one. The client routes by
// ACTION NAME, in shared/transport.js, from one list - this page only has to
// hand it both urls.
const EXAM_URL = 'https://exam.example/exec';
const REPORTS_URL = 'https://reports.example/exec';

test('r31: the examiner page sends the cold actions to the reports deployment and everything else to the exam one', async () => {
  const { ctx, lastUrl } = d1Context(() => ({ status: 'ok' }), { api: EXAM_URL, reports: REPORTS_URL });
  for (const action of ['commanderDashboard', 'siteCombinedReport', 'centerManagerReport', 'examinerForecast']) {
    await ctx.apiGet({ action: action });
    assert.equal(lastUrl().indexOf(REPORTS_URL + '?'), 0, action + ' is served by the reports deployment');
  }
  // The hot half - the dashboard, every decision, the grant, the upload token -
  // stays with the deployment whose url every page already knows.
  for (const action of ['examinerDashboard', 'approveExaminee', 'disqualify', 'bankGrant',
                        'getResultUploadToken', 'sessionSnapshot', 'verifyLogin']) {
    await ctx.apiGet({ action: action });
    assert.equal(lastUrl().indexOf(EXAM_URL + '?'), 0, action + ' is an exam action');
  }
  await ctx.apiPost({ action: 'login', idNumber: '1', password: 'x' });
  assert.equal(lastUrl(), EXAM_URL, 'a POST is routed by the same table');

  // ...and the whole shared list really is reachable from this page's api.
  for (const action of ctx.ExamTransport.REPORTS_ACTIONS) {
    assert.equal(ctx.api.urlFor(action), REPORTS_URL, action);
  }
});

test('r31: one url for both deployments is the state before the split, and nothing moves', async () => {
  const { ctx, lastUrl } = d1Context(() => ({ status: 'ok' }), { api: 'https://synthetic/exec' });
  for (const action of ['commanderDashboard', 'examinerDashboard']) {
    await ctx.apiGet({ action: action });
    assert.equal(lastUrl().indexOf('https://synthetic/exec?'), 0, action);
  }
});


// r31, 22/09/2026: the reports deployment EXISTS. Each page must carry its url as
// a literal of its own, and it must differ from the exam url: that difference is
// the whole point of the split (DESIGN §13.3).
const EXEC_RE = /https:\/\/script\.google\.com\/macros\/s\/[A-Za-z0-9_-]+\/exec/;
function reportsUrlOf(src, line) {
  const m = new RegExp(line.replace(/[.*+?^$()|[\]\\]/g, '\\$&').replace('@@', '(' + EXEC_RE.source + ')')).exec(src);
  assert.ok(m, 'the page carries the line: ' + line);
  return m[1];
}
function examUrlOf(src, constant) {
  const m = new RegExp('var ' + constant + ' ?= ?\'(' + EXEC_RE.source + ')\';').exec(src);
  assert.ok(m, 'the page carries ' + constant);
  return m[1];
}

test('r31: the page carries the reports deployment url as a line of its own, different from the exam url', () => {
  const reports = reportsUrlOf(examiner, "\r\n  var REPORTS_API_URL = '@@';\r\n");
  assert.notEqual(reports, examUrlOf(examiner, 'API_URL'), 'reports actions must leave the exam deployment');
  assert.match(section(examiner, '  var api = ExamTransport.createApi({', '  // timeoutMs is forwarded'),
    /reportsUrl: REPORTS_API_URL,/, 'and it is handed to the transport');

  // find_image.html carries the same line for the same reason, but keeps
  // sending to the EXAM deployment: bankGrant is an exam action and it is the
  // only thing that page asks the server for.
  const findImage = fs.readFileSync(path.join(app, 'find_image.html'), 'utf8');
  assert.equal(reportsUrlOf(findImage, "\r\n  var REPORTS_API_URL = '@@';\r\n"), reports, 'the same reports url on every page');
  assert.match(findImage, /apiUrl: API_URL,/);
  assert.ok(findImage.indexOf('reportsUrl') < 0, 'nothing on that page belongs to the reports deployment');
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
  assert.equal(posts[0].url, 'https://gateway.example/v1/invalidate?sessionCode=ABC12345&idNumber=1&status=approved&grant=payload.sig',
    'the nudge carries the decision, so the next poll (<=2 s) already has it - and the examiner grant the gateway demands');
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
    'https://gateway.example/v1/invalidate?sessionCode=ABC12345&idNumber=123456789&status=approved&examMinutes=50&audio=on&grant=payload.sig',
    '+25% is the 50 minutes the server itself computes as round(40 * 1.25)');

  await ctx.examinerDecision({ action: 'approveExaminee', sessionCode: 'ABC12345', idNumber: '2', timeExtension: '1.5', audioMode: 'off' });
  assert.equal(posts[1].url,
    'https://gateway.example/v1/invalidate?sessionCode=ABC12345&idNumber=2&status=approved&examMinutes=60&audio=off&grant=payload.sig',
    '+50% -> 60 minutes, audio explicitly off');

  await ctx.examinerDecision({ action: 'approveExaminee', sessionCode: 'ABC12345', idNumber: '3', audioMode: 'off' });
  assert.equal(posts[2].url,
    'https://gateway.example/v1/invalidate?sessionCode=ABC12345&idNumber=3&status=approved&audio=off&grant=payload.sig',
    'no extension chosen -> no examMinutes, the examinee keeps the default 40');
});

test('every other decision nudges with its own status, addExamTime with the running total', async () => {
  const { ctx, posts, setAnswer } = grantContext(() => ({ status: 'ok', bank: GRANT }));
  await ctx.fetchBankGrant(true);
  setAnswer(() => ({ status: 'ok' }));
  // The grant is always the last parameter; these assertions are about the decision.
  const query = url => url.slice(url.indexOf('?') + 1).replace(/&grant=payload\.sig$/, '');

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
  assert.equal(posts[0].url, 'https://gateway.example/v1/invalidate?sessionCode=ABC12345&grant=payload.sig',
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
  assert.equal(posts[0].url, 'https://gateway.example/v1/invalidate?sessionCode=ABC12345&grant=g');
});

test('a grant about to expire is still sent, and the next decision is armed with a fresh one', async () => {
  // 5 minutes left: the gateway would still accept it, so it goes out as is —
  // and the page asks for a new grant in the background (once, deduped).
  const dying = { url: 'https://gateway.example', grant: 'dying', exp: EPOCH + 5 * 60 * 1000 };
  const { ctx, posts, calls, setAnswer, store } = grantContext(() => ({ status: 'ok', bank: GRANT }));
  ctx.examinerBank = dying;
  setAnswer(params => params.action === 'bankGrant' ? { status: 'ok', bank: GRANT } : { status: 'ok' });
  await ctx.examinerDecision({ action: 'disqualify', idNumber: '9' });
  assert.equal(posts[0].url, 'https://gateway.example/v1/invalidate?sessionCode=ABC12345&idNumber=9&status=disqualified&grant=dying');
  assert.deepEqual(calls, ['disqualify', 'bankGrant'], 'one refresh request, after the decision itself');
  await ctx.bankGrantPromise;
  assert.equal(ctx.examinerBank.grant, GRANT.grant, 'the next nudge carries the fresh grant');

  // The copy in memory is gone but the stored one is fine (a reload, another
  // tab): it is adopted on the spot and the nudge goes out with it - no request.
  posts.length = 0; calls.length = 0;
  ctx.examinerBank = null;
  await ctx.examinerDecision({ action: 'disqualify', idNumber: '9' });
  assert.equal(posts[0].url, 'https://gateway.example/v1/invalidate?sessionCode=ABC12345&idNumber=9&status=disqualified&grant=payload.sig');
  assert.deepEqual(calls, ['disqualify'], 'the stored grant was adopted without asking the server');

  // Nothing at all in hand: no POST into the void, but a refresh is still asked for.
  posts.length = 0; calls.length = 0;
  ctx.examinerBank = null; store.delete('ext_examiner_bank');
  setAnswer(params => params.action === 'bankGrant' ? { status: 'error', code: 'bank_not_configured' } : { status: 'ok' });
  await ctx.examinerDecision({ action: 'disqualify', idNumber: '9' });
  assert.equal(posts.length, 0);
  assert.deepEqual(calls, ['disqualify', 'bankGrant']);
});

// ---------------------------------------------------------------- 22/09 12:29
// Live incident: the examiner pressed approve, Google stalled, the request hit
// its 30 s deadline - and the write landed a few seconds later anyway. The nudge
// followed only status:'ok', so the gateway was never told that anything had
// changed, and since r31.3 it re-reads Google only when somebody announces a
// change. The examinee learned he was approved from the Worker's own safety
// read instead of from the decision. A lost ANSWER is not a decision that did
// not happen.
const timeoutError = () => Object.assign(new Error('Request timed out'), { name: 'TimeoutError', transport: 'timeout' });
const DROP = 'https://gateway.example/v1/invalidate?sessionCode=ABC12345&grant=payload.sig';

test('a decision that timed out announces a plain DROP at once and again at +10 s', async () => {
  const { ctx, posts, timer, setAnswer } = grantContext(() => ({ status: 'ok', bank: GRANT }));
  await ctx.fetchBankGrant(true);
  setAnswer(() => Promise.reject(timeoutError()));
  const outcome = await ctx.examinerDecision({ action: 'approveExaminee', sessionCode: 'ABC12345', idNumber: '123456789', audioMode: 'on' })
    .then(() => 'resolved', e => e.name);
  assert.equal(outcome, 'TimeoutError', 'the caller still sees its own failure');
  assert.equal(posts.length, 1, 'the gateway is told immediately');
  assert.equal(posts[0].url, DROP,
    'a PLAIN DROP - no idNumber, no status: a patch is a lie until the server has said ok');
  await timer.advance(9999);
  assert.equal(posts.length, 1, 'and not a moment earlier');
  await timer.advance(1);
  assert.equal(posts.length, 2, 'again at +10 s, because the write can land after we gave up waiting');
  assert.equal(posts[1].url, DROP);
  await timer.advance(10 * 60 * 1000);
  assert.equal(posts.length, 2, 'exactly twice - this is an announcement, not a poll');
});

test('every unknown outcome announces; every refusal the server MEANT announces nothing', async () => {
  const unknown = [
    ['a dropped connection', () => Promise.reject(new Error('Failed to fetch'))],
    ["Google's HTML error page", () => Promise.reject(Object.assign(new Error('Non-JSON response'), { name: 'SyntaxError', transport: 'nonjson' }))],
    ['a 502 from its front door', () => Promise.reject(Object.assign(new Error('HTTP 502'), { name: 'HttpError', transport: 'http', status: 502 }))],
    ['a retryable error the server itself flagged', () => ({ status: 'error', retryable: true, message: 'השרת עמוס' })],
    ['an answer we cannot read', () => ({})]
  ];
  const definitive = [
    ['a plain refusal', () => ({ status: 'error', message: 'busy' })],
    ['no permission', () => ({ status: 'error', message: 'אין הרשאה' })],
    ['an expired token', () => ({ status: 'error', tokenExpired: true })],
    ['the wrong deployment', () => ({ status: 'error', code: 'wrong_deployment', message: 'הפעולה שייכת לשרת אחר' })]
  ];
  for (const [name, answer] of unknown.concat(definitive)) {
    const expected = unknown.some(u => u[0] === name);
    const { ctx, posts, timer, setAnswer } = grantContext(() => ({ status: 'ok', bank: GRANT }));
    await ctx.fetchBankGrant(true);
    setAnswer(answer);
    await ctx.examinerDecision({ action: 'disqualify', sessionCode: 'ABC12345', idNumber: '5' }).catch(() => {});
    await timer.advance(20000);
    assert.equal(posts.length, expected ? 2 : 0, name);
    if (expected) assert.equal(posts[0].url, DROP, name + ': a drop, never a patch');
  }
});

test('a confirmed decision is unchanged: one patch nudge, and no recheck is armed', async () => {
  const { ctx, posts, timer, setAnswer } = grantContext(() => ({ status: 'ok', bank: GRANT }));
  await ctx.fetchBankGrant(true);
  setAnswer(() => ({ status: 'ok' }));
  await ctx.examinerDecision({ action: 'approveExaminee', sessionCode: 'ABC12345', idNumber: '5', audioMode: 'off' });
  assert.equal(posts.length, 1);
  assert.equal(posts[0].url,
    'https://gateway.example/v1/invalidate?sessionCode=ABC12345&idNumber=5&status=approved&audio=off&grant=payload.sig');
  await timer.advance(10 * 60 * 1000);
  assert.equal(posts.length, 1, 'the ok path never arms the 10 s recheck');
});

test('a session that was closed in the meantime is not dropped by a recheck of the old one', async () => {
  const { ctx, posts, timer, setAnswer } = grantContext(() => ({ status: 'ok', bank: GRANT }));
  await ctx.fetchBankGrant(true);
  setAnswer(() => Promise.reject(timeoutError()));
  await ctx.examinerDecision({ action: 'forceComplete', sessionCode: 'ABC12345', idNumber: '5' }).catch(() => {});
  assert.equal(posts.length, 1);
  ctx.sessionCode = 'ZZZ99999';              // he closed it and opened another
  await timer.advance(15000);
  assert.equal(posts.length, 1, 'the recheck belongs to the session the decision was about');
});

test('the examiner is told the decision may have gone through, not just "try again"', () => {
  const { ctx } = grantContext(() => ({ status: 'ok' }));
  const timedOut = ctx.decisionErrorText({ name: 'TimeoutError' });
  const network = ctx.decisionErrorText(new Error('Failed to fetch'));
  for (const text of [timedOut, network]) {
    assert.match(text, /ייתכן שההחלטה נקלטה/, 'the sentence that stops the third click');
    assert.match(text, /הלוח יתעדכן לבד/);
  }
  assert.match(timedOut, /לא התקבל אישור בזמן/);
  assert.match(network, /שגיאת תקשורת/);

  // and it really is what the decision buttons show
  assert.match(section(examiner, "var params = { action: 'approveExaminee'", 'actions.appendChild(rejectBtn);'),
    /toastError\(decisionErrorText\(error\)\)/, 'approve');
  assert.ok(examiner.indexOf("toastError('שגיאת תקשורת')") < 0,
    'not one decision handler is left saying only "communication error"');
  // the results-table decisions used to swallow a rejection entirely, which is
  // the surest way to get a second and a third click on a decision that landed
  for (const fn of ['window.disqualifyResult', 'window.overturnDQ', 'window.confirmDQ']) {
    assert.match(section(examiner, '  ' + fn + ' = function(idx) {', '\r\n  };'),
      /\}\)\.catch\(function\(error\) \{ toastError\(decisionErrorText\(error\)\); \}\);/, fn);
  }
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

// ---------------------------------------------------------------- dashboard
// r31, DESIGN §13.2: the dashboard has no cadence. ONE request is held open at
// the Worker (GET /v1/session/watch) and answers when the session fingerprint
// changes; a changed fingerprint reads examinerDashboard once - that read is the
// only Google execution on this path. A safety net sits under it for what never
// reaches the Worker's snapshot, and for a Worker that is not answering at all.
//
// These tests drive the REAL sections, the grant section included: the watch has
// to carry an examiner grant and has to cope with a gateway that refuses it.
const GATEWAY = 'https://gateway.example';
// boundedFetch reads .ok / .status / .text(), so this is what a "response" is.
const gwOk = body => ({ ok: true, status: 200, text: () => Promise.resolve(JSON.stringify(body)) });
const gwHttp = (status, body) => ({ ok: false, status, text: () => Promise.resolve(JSON.stringify(body || {})) });
// A HELD answer: the Worker keeps the request open for `ms` and then answers.
// That is what an unchanged session looks like on the wire.
const gwHeld = (clock, ms, body) => new Promise(resolve => clock.set(() => resolve(gwOk(body)), ms));

// opts.watch(url, n, clock) -> a Promise of a response (or a rejection: no gateway)
// opts.grant()               -> what apiGet({ action: 'bankGrant' }) answers
// opts.stored === false      -> start with no stored grant at all
function dashboardContext(dashboardAnswer, opts = {}) {
  const ui = dom();
  ui.element('offlineBanner');
  const listeners = {};
  ui.document.addEventListener = (type, cb) => { (listeners[type] = listeners[type] || []).push(cb); };
  const apiCalls = [], watches = [], alerts = [];
  // r32: what each renderer was handed, in order - the two painting paths
  // (examinerDashboard and the Worker snapshot) are compared through this.
  // updateCompletedList also does what the page's own does: it is what makes
  // completedResults the list the wrong-answer cache is measured against.
  const renders = { pending: [], active: [], completed: [] };
  const setup = baseContext({
    ...ui,
    sessionCode: 'TEST00', examinerToken: 'synthetic', failedPolls: 0,
    dashboardInterval: null, countdownInterval: null,
    POLL_TIMEOUT_MS: 60000,
    completedResults: [],
    alert(msg) { alerts.push(msg); },
    updatePendingList(p) { renders.pending.push(p); },
    updateActiveList(a) { renders.active.push(a); },
    updateCompletedList(c) { renders.completed.push(c); setup.ctx.completedResults = c; }
  });
  const answerGrant = opts.grant || (() => ({ status: 'ok', bank: GRANT }));
  const answerWatch = opts.watch || (() => Promise.resolve(gwOk({ status: 'ok', fp: 's:quiet', held: 0, rows: 0 })));
  setup.ctx.apiGet = params => {
    apiCalls.push(params.action);
    if (params.action === 'bankGrant') return Promise.resolve(answerGrant());
    return Promise.resolve(dashboardAnswer ? dashboardAnswer()
      : { status: 'ok', pending: [], active: [], completed: [] });
  };
  setup.ctx.fetch = (url, fetchOpts) => {
    watches.push({ url, opts: fetchOpts });
    return answerWatch(url, watches.length, setup.timer);
  };
  setup.ctx.isBackendDegraded = () => setup.ctx.ExamTransport.isBackendDegraded();
  if (opts.stored !== false) setup.store.set('ext_examiner_bank', JSON.stringify(GRANT));
  load(setup.ctx, grantSection(examiner));
  load(setup.ctx, section(examiner, '  // ===== Dashboard polling =====', '  // ===== toast ====='));
  load(setup.ctx, section(examiner, '  var OFFLINE_BANNER_TEXT', '  // Text of the "'));
  setup.ctx.ExamTransport._setJitter(ms => ms);          // exact fake clock
  return {
    ...setup, ...ui, apiCalls, watches, listeners, renders, alerts,
    last: which => renders[which][renders[which].length - 1],
    reads: () => apiCalls.filter(a => a === 'examinerDashboard').length,
    grantRequests: () => apiCalls.filter(a => a === 'bankGrant').length,
    visibility(state) {
      ui.document.visibilityState = state;
      (listeners.visibilitychange || []).forEach(cb => cb());
    }
  };
}

test('dashboard: read once at the start, then ONLY when the fingerprint changes', async () => {
  // The Worker answers the first request (it carries no fp) at once; a request
  // that carries one it holds until something changes, or for the full 25 s.
  // No answer here carries a session: this is the r31 Worker (or an r31 server),
  // and the fallback it leaves behind has to keep behaving exactly like this.
  // What an r32 answer does instead is the "paints the board" test below.
  const script = [
    { fp: 's:a' },                 // the base
    { fp: 's:a', held: 25000 },    // 25 s in which nothing happened
    { fp: 's:b' },                 // a registration
    { fp: 's:c' },                 // an approval
    { fp: 's:c', held: 25000 }     // quiet again
  ];
  const { ctx, timer, watches, reads } = dashboardContext(undefined, {
    watch: (url, n, clock) => {
      const step = script[Math.min(n, script.length) - 1];
      const body = { status: 'ok', fp: step.fp, held: step.held || 0, rows: 1 };
      return step.held ? gwHeld(clock, step.held, body) : Promise.resolve(gwOk(body));
    }
  });
  ctx.startDashboardPolling();
  await drain();
  assert.equal(reads(), 1, 'opening the screen is the only read so far');
  assert.equal(watches.length, 1, 'and one request is now waiting at the Worker');

  await timer.advance(250);                 // LONGPOLL_GAP_MS: the hold begins
  assert.equal(watches.length, 2);
  await timer.advance(25000);
  assert.equal(reads(), 1, '25 s of holding on an unchanged session costs Google nothing');

  await timer.advance(250);
  assert.equal(reads(), 2, 'a new fingerprint reads the dashboard - once');
  await timer.advance(250);
  assert.equal(reads(), 3, 'and the next change reads it again');
  await timer.advance(20000);
  assert.equal(reads(), 3, 'nothing on this page is on a clock');
  ctx.stopDashboardPolling();
});

test('dashboard: a change that lands while a read is in flight is neither doubled nor lost', async () => {
  // The 2.7 s Google run is the reason this matters: the read that is already
  // out may have LEFT before the change landed, so its answer would paint the
  // old picture while the fingerprint already says we are up to date.
  let release = null;
  const script = [{ fp: 's:a' }, { fp: 's:b' }, { fp: 's:c' }, { fp: 's:c', held: 25000 }];
  const { ctx, timer, reads } = dashboardContext(
    () => new Promise(resolve => { release = () => resolve({ status: 'ok', pending: [], active: [], completed: [] }); }),
    {
      watch: (url, n, clock) => {
        const step = script[Math.min(n, script.length) - 1];
        const body = { status: 'ok', fp: step.fp, held: step.held || 0, rows: 1 };
        return step.held ? gwHeld(clock, step.held, body) : Promise.resolve(gwOk(body));
      }
    });
  ctx.startDashboardPolling();
  await drain();
  assert.equal(reads(), 1, 'the opening read, still hanging on Google');

  await timer.advance(250);                 // s:b arrives while that read is out
  assert.equal(reads(), 1, 'no second request on top of a live one');
  await timer.advance(250);                 // and s:c on top of it
  assert.equal(reads(), 1, 'still one');

  const finish = release;
  finish();
  await drain();
  assert.equal(reads(), 2, 'the queued change is read the moment the request settles - not in 60 s');
  ctx.stopDashboardPolling();
});

test('dashboard: 250 ms between watch requests, 5 s after a failed one - and the net follows', async () => {
  let mode = 'ok';
  const { ctx, timer, watches, reads } = dashboardContext(undefined, {
    watch: () => mode === 'ok'
      ? Promise.resolve(gwOk({ status: 'ok', fp: 's:a', held: 0 }))
      : Promise.reject(new Error('no gateway'))
  });
  ctx.startDashboardPolling();
  await drain();
  await timer.advance(249);
  assert.equal(watches.length, 1, 'not before the long-poll gap');
  await timer.advance(1);
  assert.equal(watches.length, 2, 'the next hold starts 250 ms after the answer, not 5 s');

  mode = 'down';
  await timer.advance(250);                 // t=500: this one fails
  assert.equal(watches.length, 3);
  assert.equal(reads(), 1);
  await timer.advance(4500);                // t=5000
  assert.equal(reads(), 2, 'a dead gateway puts the dashboard back on its pre-r31 5 s retry');
  assert.equal(watches.length, 3, 'and a failed watch waits DASH_FALLBACK_MS (jittered in production), not 250 ms');
  await timer.advance(5000);                // t=10000
  assert.equal(reads(), 3, 'still 5 s');
  assert.equal(watches.length, 4, 'the watch retried once at 5.5 s');

  mode = 'ok';
  await timer.advance(500);                 // t=10500: the gateway is back
  assert.equal(reads(), 4, 'after an outage we do not know what changed, so we read');
  await timer.advance(30000);
  assert.equal(reads(), 4, 'and then the 5 s cadence is gone again');
  ctx.stopDashboardPolling();
});

test('dashboard: the safety net fires at 60 s of quiet, and not one tick before', async () => {
  const { ctx, timer, reads } = dashboardContext(undefined, {
    watch: (url, n, clock) => n === 1
      ? Promise.resolve(gwOk({ status: 'ok', fp: 's:quiet', held: 0 }))
      : gwHeld(clock, 25000, { status: 'ok', fp: 's:quiet', held: 25000 })
  });
  ctx.startDashboardPolling();
  await drain();
  assert.equal(reads(), 1);
  await timer.advance(59000);
  assert.equal(reads(), 1, 'a minute of quiet costs Google one read, not twelve');
  await timer.advance(2000);
  assert.equal(reads(), 2,
    'and the net does fire: a result corrected from another page is not in the fingerprint');
  ctx.stopDashboardPolling();
});

test('dashboard: a gateway that refuses the grant gets a fresh one, and the stored copy is not re-used', async () => {
  let issued = 0;
  const { ctx, timer, watches, store, grantRequests } = dashboardContext(undefined, {
    grant: () => ({ status: 'ok', bank: { url: GATEWAY, grant: 'grant-' + (++issued), exp: EPOCH + 8 * 3600 * 1000 } }),
    watch: (url) => url.indexOf('grant=payload.sig') >= 0
      ? Promise.resolve(gwHttp(403, { status: 'error', code: 'grant_invalid' }))
      : Promise.resolve(gwOk({ status: 'ok', fp: 's:a', held: 0 }))
  });
  ctx.startDashboardPolling();
  await drain();
  assert.match(watches[0].url, /grant=payload\.sig/, 'the stored grant went out first');
  assert.equal(grantRequests(), 1, 'the 403 bought exactly one new grant');
  assert.notEqual(store.get('ext_examiner_bank'), JSON.stringify(GRANT),
    'a grant the gateway refused is not left in storage for the next reload to adopt');

  await timer.advance(5000);               // a failed watch waits DASH_FALLBACK_MS
  assert.match(watches[1].url, /grant=grant-1/, 'and the watch goes on with the new one');
  await timer.advance(250);
  assert.match(watches[2].url, /grant=grant-1/);
  assert.equal(grantRequests(), 1, 'a grant that works is not asked for again');
  ctx.stopDashboardPolling();
});

test('dashboard: an examiner with no grant at all keeps working, and asks for one at most once a minute', async () => {
  const { ctx, timer, watches, reads, grantRequests } = dashboardContext(undefined, {
    stored: false,
    grant: () => ({ status: 'error', code: 'bank_not_configured' })
  });
  ctx.startDashboardPolling();
  await drain();
  assert.equal(watches.length, 0, 'nothing is sent to a gateway we cannot authenticate to');
  assert.equal(reads(), 1);

  await timer.advance(30000);
  assert.equal(watches.length, 0);
  assert.equal(reads(), 7, 'the dashboard is simply the pre-r31 one: a read every 5 s');
  assert.equal(grantRequests(), 1,
    'and the watch does NOT turn into a bankGrant call to Google every 5 s (DASH_GRANT_RETRY_MS)');
  await timer.advance(35000);
  assert.equal(grantRequests(), 2, 'it does keep trying, once a minute');
  ctx.stopDashboardPolling();
});

test('dashboard: the watch url carries the session, the 25 s wait and the grant - and never degrades Google', async () => {
  const { ctx, timer, watches } = dashboardContext(undefined, {
    watch: (url, n) => n === 1
      ? Promise.resolve(gwOk({ status: 'ok', fp: 's:a', held: 0 }))
      : Promise.resolve(gwHttp(500, { status: 'error' }))
  });
  ctx.startDashboardPolling();
  await drain();
  assert.equal(watches[0].url,
    'https://gateway.example/v1/session/watch?sessionCode=TEST00&wait=25&grant=payload.sig',
    'the first request of a chain carries no fingerprint, so the Worker answers at once');
  assert.equal(watches[0].opts.cache, 'no-store');

  await timer.advance(250);
  assert.equal(watches[1].url,
    'https://gateway.example/v1/session/watch?sessionCode=TEST00&wait=25&fp=s%3Aa&grant=payload.sig',
    'every request after it is held against what this page already has');
  await drain();
  assert.equal(ctx.ExamTransport.isBackendDegraded(), false,
    'a 500 from the Worker is not Apps Script being ill - fetchJsonQuiet, never fetchJsonWithTimeout');
  await timer.advance(5000);                // a failed watch waits DASH_FALLBACK_MS
  assert.ok(watches[2].url.indexOf('&fp=') < 0,
    'and a failed watch drops the fingerprint, so the next answer comes back immediately');
  ctx.stopDashboardPolling();
});

test('dashboard: stopping stops the watch AND the net, and start/stop twice leaves one chain', async () => {
  let inFlight = 0, maxInFlight = 0;
  const { ctx, timer, watches, reads } = dashboardContext(() => {
    inFlight++; maxInFlight = Math.max(maxInFlight, inFlight);
    return Promise.resolve({ status: 'ok', pending: [], active: [], completed: [] }).then(r => { inFlight--; return r; });
  });
  ctx.startDashboardPolling();
  await drain();
  ctx.startDashboardPolling();               // the "resume this session" path, twice
  await drain();
  await timer.advance(10000);
  ctx.stopDashboardPolling();
  await drain();
  assert.equal(timer.pending, 0, 'no timer left waiting');
  assert.equal(timer.jobs.size, 0, 'and the 2.5 s safety interval is cleared too');

  const sent = watches.length, seen = reads();
  await timer.advance(60000);
  assert.equal(watches.length, sent, 'the watch really is stopped');
  assert.equal(reads(), seen, 'and so is the net');
  assert.equal(maxInFlight, 1, 'never two dashboard requests at once');
});

test('dashboard: a tab that comes back revives the watch instead of waiting out a dead chain', async () => {
  const { ctx, timer, watches, visibility } = dashboardContext(undefined, {
    watch: () => Promise.resolve(gwOk({ status: 'ok', fp: 's:a', held: 0 }))
  });
  ctx.startDashboardPolling();
  await drain();
  const sent = watches.length;

  visibility('hidden');
  await drain();
  assert.equal(watches.length, sent, 'going away changes nothing');

  visibility('visible');
  await drain();
  assert.equal(watches.length, sent + 1, 'coming back sends a fresh request on the spot');
  assert.ok(watches[sent].url.indexOf('&fp=') < 0,
    'without a fingerprint: after a freeze we cannot know the one we hold is still current');

  ctx.stopDashboardPolling();
  const after = watches.length;
  visibility('visible');
  await drain();
  assert.equal(watches.length, after, 'and a closed dashboard is not revived by a tab switch');
  await timer.advance(10000);
  assert.equal(watches.length, after);
});

test('dashboard: a failing read raises the banner, a good one clears it, and reads never overlap', async () => {
  // No gateway at all here - which is exactly what an examiner gets when
  // Cloudflare is unreachable, and exactly how this page behaved before r31.
  let response = { status: 'error', message: 'server busy' };
  let inFlight = 0, maxInFlight = 0;
  const { ctx, timer, nodes } = dashboardContext(() => {
    inFlight++; maxInFlight = Math.max(maxInFlight, inFlight);
    return Promise.resolve(response).then(r => { inFlight--; return r; });
  }, { watch: () => Promise.reject(new Error('no gateway')) });
  ctx.startDashboardPolling();
  await drain();
  for (let i = 0; i < 3; i++) await timer.advance(20000);
  assert.ok(ctx.failedPolls >= 3);
  assert.equal(nodes.get('offlineBanner').classList.contains('show'), true);
  assert.equal(maxInFlight, 1, 'read -> wait -> read again; never two chains');
  response = { status: 'ok', pending: [], active: [], completed: [] };
  await timer.advance(60000);
  assert.equal(ctx.failedPolls, 0);
  assert.equal(nodes.get('offlineBanner').classList.contains('show'), false);
  ctx.stopDashboardPolling();
  assert.equal(timer.pending, 0, 'stopping leaves no timer behind');
});

test('dashboard: a failed read is retried in 5 s even while the watch is perfectly healthy', async () => {
  // The watch says what changed in the SESSION; it knows nothing about whether
  // Google answered US. Without this the first failed read would sit behind the
  // 60 s net - a regression on every pre-r31 behaviour.
  let response = { status: 'error', message: 'busy' };
  const { ctx, timer, reads } = dashboardContext(() => Promise.resolve(response), {
    watch: (url, n, clock) => n === 1
      ? Promise.resolve(gwOk({ status: 'ok', fp: 's:quiet', held: 0 }))
      : gwHeld(clock, 25000, { status: 'ok', fp: 's:quiet', held: 25000 })
  });
  ctx.startDashboardPolling();
  await drain();
  assert.equal(reads(), 1);
  await timer.advance(5000);
  assert.equal(reads(), 2, 'a failure is retried at DASH_FALLBACK_MS, watch or no watch');
  response = { status: 'ok', pending: [], active: [], completed: [] };
  await timer.advance(5000);
  assert.equal(reads(), 3, 'the retry that finally works');
  await timer.advance(30000);
  assert.equal(reads(), 3, 'and then the net goes back to 60 s');
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

test('dashboard: a decision still reads the dashboard itself instead of waiting for the watch', () => {
  // The examiner pressed the button; his own screen must show the result of it
  // without a round trip through Cloudflare. (The nudge that follows is for the
  // EXAMINEE's held poll - see the grant tests above.)
  assert.match(section(examiner, "var params = { action: 'approveExaminee'", 'actions.appendChild(approveBtn);'),
    /examinerDecision\(params\)[\s\S]*?pollDashboard\(\);/, 'approve');
  for (const action of ['rejectExaminee', 'resetExaminee', 'forceComplete', 'disqualify']) {
    const region = section(examiner, "examinerDecision({ action: '" + action + "'", '});');
    assert.match(region + '});', /pollDashboard\(\)/, action + ' reads the dashboard when it succeeds');
  }
});

// ------------------------------------------------------------- r32 §14.1
// The watch now brings the DATA, not only the news. KNOWN_ISSUES #35: Google's
// delivery step takes 25-60 s (or 404s) for our projects at random, so every
// round trip is a raffle ticket - one change must cost ONE, not two. A watch
// answer that carries the session paints the board; examinerDashboard is left
// for what the snapshot cannot carry (פירוט שגויות) and as the fallback for an
// old Worker, an old server and a stale copy.

// A row of ממתינים as sessionSnapshot v2 hands it over - no token, no tokenHash.
const srow = o => Object.assign({
  id: '', status: 'waiting', audio: 'off', examMinutes: 40, extraMinutes: 0,
  warn: 0, fin: 0, ext: 0, dq: 0, name: '', phone: '', time: '', start: '',
  lang: 'he', pop: '', site: '', lic: '', timeExt: '', lastWarn: '', attemptsToday: 0
}, o);

// One session with every case the server's dedup has a rule for:
//   012345678  registered twice                 -> the LATEST row wins
//   100000002  disqualified, then in_exam again -> the DQ still wins
//   900000001  plain in_exam - and an id that sorts AFTER the one above
//   200000003  finished: in neither list, but it is where registrationTime comes from
const fixtureRows = () => [
  srow({ id: '012345678', status: 'waiting', name: 'דנה כהן', phone: '0501234567',
         time: '2026-09-22T06:05:00.000Z', pop: 'חוגרים', site: 'בח"א 6', lic: 'B' }),
  srow({ id: '012345678', status: 'approved', audio: 'on', examMinutes: 50, name: 'דנה כהן',
         phone: '0501234567', time: '2026-09-22T06:40:00.000Z', pop: 'חוגרים', site: 'בח"א 6',
         lic: 'B', timeExt: '1.25', attemptsToday: 1,
         todayExams: [{ license: 'B', score: '20/30', passed: 'נכשל', language: 'he' }] }),
  srow({ id: '900000001', status: 'in_exam', extraMinutes: 10, warn: 2, fin: 1, ext: 1,
         name: 'רון לוי', phone: '0521111111', time: '2026-09-22T06:10:00.000Z',
         start: '2026-09-22T07:00:00.000Z', lang: 'ru', pop: 'קבע', site: 'בח"א 6', lic: 'C',
         lastWarn: 'יצא מהמסך',
         // the snapshot carries this on every row; the board shows it on the
         // pending list only, so an active item must NOT come out with it
         todayExams: [{ license: 'C', score: '19/30', passed: 'נכשל', language: 'ru' }] }),
  srow({ id: '100000002', status: 'disqualified', dq: 1, warn: 1, name: 'יוסי בר',
         phone: '0533333333', time: '2026-09-22T06:12:00.000Z',
         start: '2026-09-22T06:50:00.000Z', site: 'בח"א 6', lic: 'B' }),
  srow({ id: '100000002', status: 'in_exam', dq: 1, warn: 5, name: 'יוסי בר',
         phone: '0533333333', time: '2026-09-22T06:12:00.000Z',
         start: '2026-09-22T06:50:00.000Z', site: 'בח"א 6', lic: 'B' }),
  srow({ id: '200000003', status: 'completed', name: 'מאיה גל', phone: '0544444444',
         time: '2026-09-22T05:30:00.000Z', start: '2026-09-22T06:00:00.000Z', site: 'בח"א 6', lic: 'B' })
];
const FABRICATED_TEXT = 'ניתוק/טיימאאוט — הנבחן לא סיים את המבחן';
const WRONG_TEXT = 'שאלה: מה המרחק?\nתשובת הנבחן: 10\nתשובה נכונה: 20';
// The session's results, exactly as the snapshot sends them: the latest non-בוטל
// row per id, in sheet order, WITHOUT wrongDetails - and with the server's own
// fabricated verdict where the block it replaces held one of its markers.
const fixtureResults = () => [
  { date: '22/09/2026', idNumber: 200000003, name: 'מאיה גל', phone: '0544444444', license: 'B',
    score: '27/30', percent: '90%', passed: 'עבר', time: '18:12', examiner: 'בוחן א', site: 'בח"א 6',
    classroom: 'כיתה 2', language: 'he', attempt: 1, sent: false, disqualified: false, waLink: '',
    population: 'חוגרים', corrected: false, audioMode: 'off', verified: 'מאומת', suspicious: '', device: 'Windows' },
  { date: '22/09/2026', idNumber: 300000004, name: 'שיר דוד', phone: '0555555555', license: 'B',
    score: '0/30', percent: '0%', passed: 'נכשל', time: '', examiner: 'בוחן א', site: 'בח"א 6',
    classroom: 'כיתה 2', language: 'he', attempt: 2, sent: false, disqualified: false, waLink: '',
    population: '', corrected: false, audioMode: 'off', verified: '', suspicious: '', device: '',
    fabricated: 1 }
];
const fixtureSession = () => ({ rows: fixtureRows(), results: fixtureResults() });
// The page builds its lists inside the vm, so its arrays and objects belong to
// that realm and deepStrictEqual would reject them on the prototype alone. Copy
// both sides into this one before comparing (a key whose value is undefined is
// dropped on the way, so those are asserted on their own).
const host = v => JSON.parse(JSON.stringify(v));

// What handleExaminerDashboard would have returned for exactly those rows. The
// ids are the RAW cells: Sheets keeps "012345678" as the number 12345678, which
// is why both paths normalise before the renderers ever see an item.
const fixtureDashboard = () => ({
  status: 'ok',
  pending: [{
    idNumber: 12345678, name: 'דנה כהן', phone: '0501234567', time: '2026-09-22T06:40:00.000Z',
    examStartTime: '', status: 'approved', language: 'he', population: 'חוגרים', site: 'בח"א 6',
    license: 'B', audioMode: 'on', timeExtension: '1.25', dqCount: 0, warnings: 0, lastWarning: '',
    attemptsToday: 1, hasExtendedScreen: false, extraMinutes: 0, finishedOnDevice: false,
    todayExams: [{ license: 'B', score: '20/30', passed: 'נכשל', language: 'he' }]
  }],
  active: [{
    idNumber: 100000002, name: 'יוסי בר', phone: '0533333333', time: '2026-09-22T06:12:00.000Z',
    examStartTime: '2026-09-22T06:50:00.000Z', status: 'disqualified', language: 'he', population: '',
    site: 'בח"א 6', license: 'B', audioMode: 'off', timeExtension: '', dqCount: 1, warnings: 1,
    lastWarning: '', attemptsToday: 0, hasExtendedScreen: false, extraMinutes: 0,
    finishedOnDevice: false, dqPending: true
  }, {
    idNumber: 900000001, name: 'רון לוי', phone: '0521111111', time: '2026-09-22T06:10:00.000Z',
    examStartTime: '2026-09-22T07:00:00.000Z', status: 'in_exam', language: 'ru', population: 'קבע',
    site: 'בח"א 6', license: 'C', audioMode: 'off', timeExtension: '', dqCount: 0, warnings: 2,
    lastWarning: 'יצא מהמסך', attemptsToday: 0, hasExtendedScreen: true, extraMinutes: 10,
    finishedOnDevice: true
  }],
  completed: [
    Object.assign(fixtureResults()[0], { wrongDetails: WRONG_TEXT, registrationTime: '2026-09-22T05:30:00.000Z' }),
    (r => { delete r.fabricated; r.wrongDetails = FABRICATED_TEXT; return r; })(fixtureResults()[1])
  ]
});

test('r32: a watch answer that carries the session paints the board, and asks Google nothing', async () => {
  const h = dashboardContext(undefined, {
    watch: (url, n, clock) => n === 1
      ? Promise.resolve(gwOk({ status: 'ok', fp: 's:a', held: 0, rows: 6, at: EPOCH, session: fixtureSession() }))
      : gwHeld(clock, 25000, { status: 'ok', fp: 's:a', held: 25000, rows: 6 })
  });
  h.ctx.startDashboardPolling();
  await drain();
  assert.equal(h.reads(), 1,
    'the opening read - and NOT a second one on the very answer that brought the data');
  assert.equal(h.renders.pending.length, 2, 'the opening answer, then the snapshot');

  const pending = h.last('pending'), active = h.last('active'), completed = h.last('completed');
  assert.equal(pending.length, 1, 'he registered twice; the board shows him once');
  assert.equal(pending[0].idNumber, '012345678');
  assert.equal(pending[0].status, 'approved', 'the latest of the two rows wins');
  assert.equal(pending[0].name, 'דנה כהן');
  assert.equal(pending[0].time, '2026-09-22T06:40:00.000Z', 'and its registration time, not the first one');
  assert.equal(pending[0].audioMode, 'on');
  assert.equal(pending[0].timeExtension, '1.25');
  assert.equal(pending[0].attemptsToday, 1);
  assert.equal(pending[0].examStartTime, '');
  assert.deepStrictEqual(host(pending[0].todayExams), [{ license: 'B', score: '20/30', passed: 'נכשל', language: 'he' }]);

  assert.deepStrictEqual(host(active.map(a => a.idNumber)), ['100000002', '900000001'],
    'the server enumerates its own dedup object the same way - numeric keys ascending, not insertion order');
  assert.equal(active[0].status, 'disqualified');
  assert.equal(active[0].dqPending, true, 'a DQ the examiner has to decide on');
  assert.equal(active[0].warnings, 1, 'the DQ row itself, not the in_exam row that followed it');
  assert.equal(active[1].dqPending, undefined);
  assert.equal(active[1].todayExams, undefined,
    'the badge is on the pending list; the server never attaches it to an active item');
  assert.equal(active[1].hasExtendedScreen, true);
  assert.equal(active[1].finishedOnDevice, true);
  assert.equal(active[1].extraMinutes, 10);
  assert.equal(active[1].warnings, 2);
  assert.equal(active[1].lastWarning, 'יצא מהמסך');
  assert.equal(active[1].examStartTime, '2026-09-22T07:00:00.000Z');
  assert.equal(active[1].language, 'ru');
  assert.equal(active[1].license, 'C');

  assert.equal(completed.length, 2);
  assert.equal(completed[0].idNumber, '200000003');
  assert.equal(completed[0].registrationTime, '2026-09-22T05:30:00.000Z',
    'from his row in ממתינים, whatever status that row ended in');
  assert.equal(completed[0].wrongDetails, undefined,
    'unknown - and undefined is NOT "" (a result with no wrong answers)');
  assert.equal('registrationTime' in completed[1], false, 'nobody registered under that id in this session');
  assert.equal(completed[1].fabricated, 1, 'the badge rule needs it when there is no block to test');
  assert.equal(h.ctx.window.__lastPending, pending, 'the search box re-renders from the same list');

  await h.timer.advance(60000);
  assert.equal(h.reads(), 2, 'a minute of watching costs one read: the safety net, which is not a cadence');
  h.ctx.stopDashboardPolling();
});

test('r32: the snapshot and examinerDashboard build the SAME items', async () => {
  const h = dashboardContext(() => fixtureDashboard(), { watch: () => Promise.reject(new Error('no gateway')) });
  await h.ctx.pollDashboard();
  const dash = { pending: h.last('pending'), active: h.last('active'), completed: h.last('completed') };
  assert.equal(h.reads(), 1);

  h.ctx.renderFromSession(fixtureSession());
  const snap = { pending: h.last('pending'), active: h.last('active'), completed: h.last('completed') };
  assert.equal(h.reads(), 1, 'painting from the snapshot asks Google nothing at all');

  assert.deepStrictEqual(host(snap.pending), host(dash.pending));
  assert.deepStrictEqual(host(snap.active), host(dash.active));
  // The one difference a result is allowed: the snapshot carries the server's
  // own fabricated verdict instead of the text it was computed from.
  const noFlag = r => { const c = Object.assign({}, r); delete c.fabricated; return c; };
  assert.deepStrictEqual(host(snap.completed).map(noFlag), host(dash.completed),
    'the wrong-answer blocks come back out of the cache the dashboard read filled');
  assert.equal(snap.completed[1].fabricated, 1);
});

test('r32: no session in the answer still reads the dashboard, and a stale copy is never painted', async () => {
  const script = [
    { fp: 's:a' },                                           // the base - an old Worker: no session
    { fp: 's:b' },                                           // a change, still without one
    { fp: 's:c', stale: true, session: fixtureSession() }    // a copy too old to paint from
  ];
  const h = dashboardContext(undefined, {
    watch: (url, n, clock) => {
      const step = script[Math.min(n, script.length) - 1];
      const body = Object.assign({ status: 'ok', held: 0, rows: 1 }, step);
      return n > script.length ? gwHeld(clock, 25000, body) : Promise.resolve(gwOk(body));
    }
  });
  h.ctx.startDashboardPolling();
  await drain();
  assert.equal(h.reads(), 1);
  const painted = h.renders.pending.length;

  await h.timer.advance(250);
  assert.equal(h.reads(), 2, 'an old Worker: a changed fingerprint still reads the dashboard');
  await h.timer.advance(250);
  assert.equal(h.reads(), 3, 'a stale copy is read for, not painted from');
  assert.equal(h.renders.pending.length, painted + 2,
    'two dashboard answers - and nothing at all out of the stale session');
  h.ctx.stopDashboardPolling();
});

test('r32: an examinerDashboard answer fills the wrong-answer cache the snapshot cannot carry', async () => {
  const h = dashboardContext(() => ({
    status: 'ok', pending: [], active: [], completed: [
      { idNumber: 12345678, attempt: 1, wrongDetails: 'שאלה: א' },
      { idNumber: '987654321', attempt: 2, wrongDetails: '' }
    ]
  }), { watch: () => Promise.reject(new Error('no gateway')) });
  await h.ctx.pollDashboard();
  assert.deepStrictEqual(Object.assign({}, h.ctx.resultDetails),
    { '012345678:1': 'שאלה: א', '987654321:2': '' },
    "keyed by the normalised id, and '' is cached like any other value");

  // ...and it belongs to the session that filled it: opening another one (a
  // commander loading a colleague's session) starts from nothing, and the
  // opening read refills it.
  h.ctx.sessionCode = 'TEST01';
  h.ctx.startDashboardPolling();
  assert.deepStrictEqual(host(h.ctx.resultDetails), {});
  h.ctx.stopDashboardPolling();
});

test('r32: ensureResultDetails reads once when a block is missing, and not at all when it is not', async () => {
  let answer = { status: 'ok', pending: [], active: [], completed: [{ idNumber: '000000003', attempt: 2, wrongDetails: 'שאלה: ב' }] };
  const h = dashboardContext(() => answer, { watch: () => Promise.reject(new Error('no gateway')) });

  h.ctx.completedResults = [{ idNumber: '000000001', attempt: 1, wrongDetails: '' }];
  await h.ctx.ensureResultDetails();
  assert.equal(h.reads(), 0, "'' is a result with no wrong answers - a known value, not a miss");

  h.ctx.completedResults = [{ idNumber: '000000002', attempt: 1, fabricated: 1 }];
  await h.ctx.ensureResultDetails();
  assert.equal(h.reads(), 0, 'a fabricated fail never had a block worth reading');

  h.ctx.completedResults = [{ idNumber: '000000001', attempt: 1, wrongDetails: '' },
                            { idNumber: '000000003', attempt: 2 }];
  const at = await h.ctx.ensureResultDetailsAt(1);
  assert.equal(h.reads(), 1, 'one Google run, on the examiner\'s click - never on a clock');
  assert.equal(at, 0, 'and the index is re-resolved: that read REPLACED the list');
  assert.deepEqual(h.alerts, []);

  // A read that did not answer leaves the block unknown. Say so, instead of
  // handing the examinee a certificate whose "questions you got wrong" is empty.
  answer = { status: 'error', message: 'busy' };
  h.ctx.completedResults = [{ idNumber: '000000009', attempt: 1 }];
  assert.equal(await h.ctx.ensureResultDetailsAt(0), -1);
  assert.equal(h.alerts.length, 1);
  assert.equal(await h.ctx.ensureAllResultDetails(), false);
  assert.equal(h.alerts.length, 2);
});

// Writes to תוצאות that do NOT go through examinerDecision (they are not
// decisions about an examinee) still have to reach the Worker's snapshot, or the
// next watch render paints the old row back over them until the 20 s safety read.
function resultWriterContext(src) {
  const ui = dom();
  const calls = [], posts = [];
  let answer = params => (params.action === 'bankGrant' ? { status: 'ok', bank: GRANT } : { status: 'ok' });
  const setup = baseContext({
    ...ui,
    sessionCode: 'ABC12345', examinerToken: 't', examinerData: { id: '111', name: 'בוחן' },
    completedResults: [{ idNumber: '012345678', name: 'דנה', score: '25/30', attempt: 1,
                         wrongDetails: 'שאלה: א\nתשובת הנבחן: 1\nתשובה נכונה: 2' }],
    resultDetailsMissing: () => false,
    ensureResultDetailsAt: () => Promise.resolve(0),
    pollDashboard() {}, showToast() {}, alert() {}, confirm: () => true,
    apiGet(params) { calls.push(params.action); return Promise.resolve(answer(params)); },
    fetch(url, opts) { posts.push({ url, opts }); return Promise.resolve({ ok: true, text: () => Promise.resolve('{}') }); }
  });
  load(setup.ctx, grantSection(examiner));
  load(setup.ctx, src);
  return { ...setup, calls, posts, setAnswer(fn) { answer = fn; } };
}

test('r32: markSent announces itself to the gateway, and only when the write went through', async () => {
  const h = resultWriterContext(section(examiner,
    '  // Mark as sent when WA link is clicked.', '  // Parse wrongDetails text into table rows'));
  await h.ctx.fetchBankGrant(true);
  h.ctx.window.markSentWA(0);
  await drain();
  assert.deepEqual(h.calls, ['bankGrant', 'markSent']);
  assert.equal(h.posts.length, 1);
  assert.equal(h.posts[0].url, DROP, 'a PLAIN drop: there is no ממתינים row to patch a "sent" flag into');
  assert.equal(h.posts[0].opts.method, 'POST');

  h.setAnswer(() => ({ status: 'error', message: 'busy' }));
  h.ctx.window.markSentWA(0);
  await drain();
  assert.equal(h.posts.length, 1, 'nothing was written, so there is nothing to announce');
});

test('r32: correctToPass announces itself too, so the correction is not painted back off', async () => {
  const h = resultWriterContext(section(examiner,
    '  window.correctToPass = function(idx, ready) {', '  // ========== Settings Modal =========='));
  await h.ctx.fetchBankGrant(true);
  h.ctx.window.correctToPass(0);
  await drain();
  assert.deepEqual(h.calls, ['bankGrant', 'correctToPass']);
  assert.equal(h.posts.length, 1);
  assert.equal(h.posts[0].url, DROP);

  h.setAnswer(params => (params.action === 'bankGrant' ? { status: 'ok', bank: GRANT } : { status: 'error', message: 'busy' }));
  h.ctx.window.correctToPass(0);
  await drain();
  assert.equal(h.posts.length, 1, 'a refusal changed nothing in תוצאות');
});

test('r32: every consumer of the wrong-answer block asks for it first', () => {
  const perRow = ['window.openExamineePDF', 'window.shareResult', 'window.correctToPass'];
  const perList = ['window.sendToAllExaminees', 'window.generateSiteManagerReport', 'window.shareSiteManagerReport'];
  for (const fn of perRow.concat(perList)) {
    const body = section(examiner, '  ' + fn + ' = function(', '\r\n  };');
    assert.match(body, /resultDetailsMissing\(\)/, fn + ': pays for a read only when something is missing');
    assert.match(body, perRow.indexOf(fn) >= 0 ? /ensureResultDetailsAt\(idx\)/ : /ensureAllResultDetails\(\)/,
      fn + ': waits for פירוט שגויות before it runs');
  }
  // The two that open a tab must claim it INSIDE the click - Safari will not
  // honour a window.open that runs after the round trip.
  for (const fn of ['window.openExamineePDF', 'window.generateSiteManagerReport']) {
    const body = section(examiner, '  ' + fn + ' = function(', '\r\n  };');
    assert.match(body, /openReportWindow\(\);[\s\S]*ensure/, fn + ': the tab is opened before the wait, not after it');
    assert.match(body, /writeReportWindow\(pre, html\)/, fn + ': and the HTML goes into the tab that was claimed');
  }
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
    // r32: a board painted from the Worker snapshot has no block to test, so the
    // server's own verdict (fabricated) stands in for it.
    const isFabricatedFail = (typeof r.wrongDetails === 'undefined')
      ? !!r.fabricated
      : ctx.FABRICATED_FAIL_MARKERS.test(String(r.wrongDetails || ''));
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
    ['an old payload without the column', { score: '24/30', wrongDetails: '' }, false],
    // r32: the same two rows as they arrive from the Worker snapshot - no block
    // at all, only the verdict the server computed from it
    ['a snapshot row the server called fabricated', { score: '0/30', verified: '', fabricated: 1 }, false],
    ['a snapshot row it did not', { score: '24/30', verified: '' }, true]
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
