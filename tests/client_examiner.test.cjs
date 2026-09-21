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
//   D3/D4 the update check shows a banner and NEVER reloads by itself
//   S3  the "not verified" badge keys on the stored marker, not on "0/"
//   plus: the dashboard loop never overlaps and honours the 2 s sync window,
//         top-wrong rendering with and without a bank entry, and both SWs.
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
test('D3/D4: a new version raises a banner with a button and never reloads by itself', async () => {
  const ui = dom();
  let reloads = 0, version = { build: 'b1', pages: { 'examiner.html': 'hash-1' } };
  const { ctx, timer } = baseContext({
    ...ui,
    sessionCode: '',
    location: { pathname: '/examiner.html', reload() { reloads++; } },
    fetch: () => Promise.resolve({ ok: true, text: () => Promise.resolve(JSON.stringify(version)) })
  });
  load(ctx, section(examiner, '  // ===== Auto-update: BANNER ONLY', '  var deferredInstallPrompt = null;'));
  await drain();
  version = { build: 'b2', pages: { 'examiner.html': 'hash-2' } };
  await timer.advance(120000);                  // first sighting: not enough
  assert.equal(ui.nodes.get('examinerUpdateBanner') || null, null);
  await timer.advance(120000);                  // same hash twice = a real deploy
  assert.ok(ui.nodes.get('examinerUpdateBanner'), 'the banner appears');
  await timer.advance(10 * 60 * 1000);
  assert.equal(reloads, 0, 'decision 4: an examiner page NEVER reloads itself (16/09)');
  ui.nodes.get('swUpdNow').click();
  assert.equal(reloads, 1, 'only the button reloads');
});

test('D3/D4: a rewritten header or an error page is not a new version', async () => {
  const ui = dom();
  let answer = { ok: true, text: () => Promise.resolve(JSON.stringify({ build: 'b1', pages: { 'examiner.html': 'hash-1' } })) };
  const { ctx, timer } = baseContext({
    ...ui, sessionCode: '',
    location: { pathname: '/examiner.html', reload() { throw new Error('must not reload'); } },
    fetch: () => Promise.resolve(answer)
  });
  load(ctx, section(examiner, '  // ===== Auto-update: BANNER ONLY', '  var deferredInstallPrompt = null;'));
  await drain();
  answer = { ok: true, text: () => Promise.resolve('<html>captive portal</html>') };
  await timer.advance(120000); await timer.advance(120000);
  answer = { ok: false, status: 503, text: () => Promise.resolve('{}') };
  await timer.advance(120000); await timer.advance(120000);
  assert.equal(ui.nodes.get('examinerUpdateBanner') || null, null,
    'only a 200 JSON with a hash for THIS page counts');
});

// ---------------------------------------------------------------- top wrong
function topWrongContext(bankEntries) {
  const ui = dom();
  ui.element('cmdTopWrong');
  const { ctx } = baseContext({
    ...ui,
    cmdData: null,
    escapeHtml: s => String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'),
    QuestionBank: {
      has: () => true,
      load: () => Promise.resolve(),
      get: id => bankEntries[id] || null,
      imageUrl: e => (e && e.image ? 'images/' + e.image : '')
    }
  });
  load(ctx, section(examiner, '  // The server sends { questionId, count, category, text }', '  // Day-of-week'));
  return { ctx, ...ui };
}

test('top wrong: the canonical Hebrew text and the image come from the bank', () => {
  const { ctx, nodes } = topWrongContext({
    14: { id: 14, text: 'מה פירוש התמרור?', image: 'TQ_PIC_14.jpg' }
  });
  ctx.renderTopWrong([{ questionId: 14, count: 9, category: 'תמרורים', text: 'stale copy from the sheet' }]);
  const html = nodes.get('cmdTopWrong').innerHTML;
  assert.match(html, /images\/TQ_PIC_14\.jpg/, 'the image is resolved locally, not sent by the server');
  assert.match(html, /מה פירוש התמרור\?/, 'the bank text wins');
  assert.ok(!/stale copy/.test(html));
  assert.match(html, /id: 14/);
  assert.match(html, /תמרורים/);
});

test('top wrong: an id the bank does not carry falls back to the stored text and shows no image', () => {
  const { ctx, nodes } = topWrongContext({});
  ctx.renderTopWrong([{ questionId: 9999, count: 4, category: '', text: 'שאלה שהוצאה מהמאגר' }]);
  const html = nodes.get('cmdTopWrong').innerHTML;
  assert.match(html, /שאלה שהוצאה מהמאגר/);
  assert.ok(!/<img/.test(html), 'no broken image box');
  assert.match(html, /find_image\.html\?q=/, 'the deep link still works');
});

test('top wrong: an empty list says so instead of rendering an empty table', () => {
  const { ctx, nodes } = topWrongContext({});
  ctx.renderTopWrong([]);
  assert.match(nodes.get('cmdTopWrong').innerHTML, /אין נתוני/);
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
    fetchHandler({ request: { method: 'GET', url: 'https://example/bank/he.json?v=abc' }, respondWith: () => { responded = true; } });
    assert.equal(responded, true, 'same-origin GETs are served network-first');
    assert.match(src, /ignoreSearch: true/, 'bank/<lang>.json?v=<sha> must still match its cached copy offline');
  });
}
