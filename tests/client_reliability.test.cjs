// Regression tests execute the actual client functions with synthetic inputs.
// No production requests, browser state, external dependencies, or student data.
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const app = path.resolve(__dirname, '..');
const examiner = fs.readFileSync(path.join(app, 'examiner.html'), 'utf8');
const examinee = fs.readFileSync(path.join(app, 'examinee.html'), 'utf8');
const quiet = { log() {}, warn() {}, error() {} };
function section(src, start, end) {
  const i = src.indexOf(start), j = src.indexOf(end, i + start.length);
  assert.ok(i >= 0 && j > i, 'source section found: ' + start);
  return src.slice(i, j);
}
const deferred = () => { let resolve, reject; const promise = new Promise((a, b) => { resolve = a; reject = b; }); return { promise, resolve, reject }; };
async function drain() { for (let i = 0; i < 20; i++) await Promise.resolve(); }
class Timers {
  now = 0; nextId = 0; jobs = new Map();
  set = (cb, ms) => { const id = ++this.nextId; this.jobs.set(id, { cb, at: this.now + Number(ms) }); return id; };
  clear = id => this.jobs.delete(id);
  async advance(ms) {
    const until = this.now + ms;
    for (let count = 0; count < 10000; count++) {
      const next = [...this.jobs].filter(([, job]) => job.at <= until).sort((a, b) => a[1].at - b[1].at)[0];
      if (!next) { this.now = until; await drain(); return; }
      this.now = next[1].at; this.jobs.delete(next[0]); next[1].cb(); await drain();
    }
    throw new Error('Unexpected timer loop');
  }
}
function dom() {
  const nodes = new Map();
  // visibilityState + document-level listeners: the iOS rescue nets (approval
  // chain restart, markExamStarted re-fire) hang off visibilitychange.
  const docHandlers = new Map();
  const document = { activeElement: null, getElementById: id => nodes.get(id) || null,
    visibilityState: 'visible',
    addEventListener(type, cb) { if (!docHandlers.has(type)) docHandlers.set(type, []); docHandlers.get(type).push(cb); },
    removeEventListener(type, cb) { const list = docHandlers.get(type) || []; const i = list.indexOf(cb); if (i >= 0) list.splice(i, 1); } };
  function element(id) {
    const el = { id, textContent: '', disabled: false, style: {}, handlers: {}, attrs: {},
      classList: { values: new Set(), add(v) { this.values.add(v); }, remove(v) { this.values.delete(v); }, contains(v) { return this.values.has(v); } },
      addEventListener(type, cb) { this.handlers[type] = cb; },
      setAttribute(key, value) { this.attrs[key] = value; },
      appendChild(child) { this.children = this.children || []; this.children.push(child); child.parentNode = this; nodes.set(child.id, child); },
      insertBefore(child, before) { this.children = this.children || []; const at = this.children.indexOf(before); this.children.splice(at < 0 ? this.children.length : at, 0, child); child.parentNode = this; nodes.set(child.id, child); },
      removeChild(child) { this.children = (this.children || []).filter(el => el !== child); nodes.delete(child.id); child.parentNode = null; },
      focus() { document.activeElement = this; },
      click() { if (!this.disabled && this.handlers.click) this.handlers.click(); }
    };
    let html = '';
    Object.defineProperty(el, 'innerHTML', { get: () => html, set(value) {
      html = value;
      el.children = [];   // as in a real DOM: writing innerHTML drops the appended children
      for (const match of value.matchAll(/id="([^"]+)"/g)) element(match[1]);
    } });
    nodes.set(id, el); return el;
  }
  document.createElement = () => element('');
  document.body = { appendChild(el) { nodes.set(el.id, el); el.parentNode = this; }, removeChild(el) { nodes.delete(el.id); } };
  for (const id of ['examArea', 'offlineBanner', 'approvalError']) element(id);
  // Drive the page between background and foreground, as iOS does.
  const setVisibility = state => {
    document.visibilityState = state;
    for (const cb of docHandlers.get('visibilitychange') || []) cb();
  };
  return { document, nodes, element, setVisibility };
}
function context(extra = {}, timer = new Timers()) {
  const clockDate = class extends Date { static now() { return timer.now; } };
  // CRITICAL_POST_TIMEOUT_MS is declared in the apiGet-area section of the page
  // but referenced from the separately-loaded submit/register sections; mirror
  // the page constant here the way the harness already stubs cross-section deps.
  const ctx = { console: quiet, Date: clockDate, setTimeout: timer.set, clearTimeout: timer.clear,
    clearInterval: timer.clear, AbortController, CRITICAL_POST_TIMEOUT_MS: 60000, ...extra };
  vm.createContext(ctx); return { ctx, timer };
}
function load(ctx, code) { vm.runInContext(code, ctx); }
const helper = src => section(src, '  var API_TIMEOUT_MS = 30000;', '  function apiGet(');
// 2026-09-18 poll pacing block (jitter, gradual recovery, 60s poll deadline). Every
// poll section reaches it at call time; contexts pin the jitter to identity so the
// fake clock stays exact, and the real jitter is tested on its own below.
const pollHelpers = src => section(src, '  // ===== Poll pacing', '  // ===== end poll pacing');
// The pacing block asks the transport-health block (next to fetchJsonWithTimeout)
// whether the backend is degraded, so the helper section is loaded first.
function withPacing(ctx, src) { load(ctx, helper(src)); load(ctx, pollHelpers(src)); ctx.jitterMs = ms => ms; }

for (const [name, src] of [['examiner', examiner], ['examinee', examinee]]) {
  test(name + ': deadline includes a stalled response body and ignores a late body', async () => {
    const body = deferred(); let signal;
    const { ctx, timer } = context({ fetch: (url, opts) => { signal = opts.signal; return Promise.resolve({ ok: true, text: () => body.promise }); } });
    load(ctx, helper(src));
    const result = ctx.fetchJsonWithTimeout('synthetic', {}, 100).then(value => ({ value }), error => ({ error }));
    await drain(); assert.equal(timer.jobs.size, 1);
    await timer.advance(100);
    assert.equal((await result).error.name, 'TimeoutError'); assert.equal(signal.aborted, true);
    body.resolve('{"status":"ok"}'); await drain();
    assert.equal((await result).error.name, 'TimeoutError'); assert.equal(timer.jobs.size, 0);
  });
  test(name + ': healthy JSON, HTTP errors and malformed JSON all clear the deadline', async () => {
    for (const [ok, body, expected] of [[true, '{"status":"ok"}', 'ok'], [false, '{}', 'HTTP 503'], [true, '<html>busy</html>', 'SyntaxError']]) {
      const { ctx, timer } = context({ fetch: () => Promise.resolve({ ok, status: 503, text: () => Promise.resolve(body) }) });
      load(ctx, helper(src));
      const outcome = await ctx.fetchJsonWithTimeout('synthetic', {}).then(value => value.status, error => error.name === 'SyntaxError' ? error.name : error.message);
      assert.equal(outcome, expected); assert.equal(timer.jobs.size, 0);
    }
  });
  test(name + ': requests still time out without AbortController', async () => {
    const { ctx, timer } = context({ AbortController: undefined, fetch: () => new Promise(() => {}) });
    load(ctx, helper(src));
    const outcome = ctx.fetchJsonWithTimeout('synthetic', {}, 100).catch(error => error.name);
    await drain(); await timer.advance(100); assert.equal(await outcome, 'TimeoutError');
  });
}

function dashboardContext(apiGet) {
  const ui = dom();
  const setup = context({ ...ui, sessionCode: 'TEST00', examinerToken: 'synthetic', failedPolls: 0,
    dashboardInterval: null, countdownInterval: null, apiGet,
    updatePendingList() {}, updateActiveList() {}, updateCompletedList() {} });
  setup.ctx.window = setup.ctx;
  load(setup.ctx, section(examiner, '  var DASH_POLL_BASE_MS', '  // ========== Pending list'));
  withPacing(setup.ctx, examiner);
  return { ...setup, ...ui };
}
test('dashboard counts server errors and only a successful response clears the warning', async () => {
  let response = { status: 'error', message: 'server busy' };
  const { ctx, nodes } = dashboardContext(() => Promise.resolve(response));
  for (let i = 0; i < 3; i++) await ctx.pollDashboard();
  assert.equal(ctx.failedPolls, 3); assert.equal(nodes.get('offlineBanner').classList.values.has('show'), true);
  response = { status: 'ok', pending: [], active: [], completed: [] }; await ctx.pollDashboard();
  assert.equal(ctx.failedPolls, 0); assert.equal(nodes.get('offlineBanner').classList.values.has('show'), false);
});
test('dashboard closes the gap while a result is syncing, and never at the cost of a struggling server', async () => {
  const syncing = { status: 'ok', pending: [], completed: [], active: [{ idNumber: 'A', finishedOnDevice: true }] };
  const settled = { status: 'ok', pending: [], completed: [], active: [{ idNumber: 'A', finishedOnDevice: false }] };
  let response = settled;
  const { ctx, timer } = dashboardContext(() => Promise.resolve(response));
  ctx.startDashboardPolling(); await drain();
  // normal cadence: nothing pending, so the next tick is the usual 5s
  await timer.advance(4999); assert.equal(ctx.dashPollDelayMs, 5000, 'idle dashboards stay at 5s');
  response = syncing; await timer.advance(1); await drain();
  assert.equal(ctx.dashPollDelayMs, 2000, 'a syncing result pulls the next tick in to 2s');
  response = settled; await timer.advance(2000); await drain();
  assert.equal(ctx.dashPollDelayMs, 5000, 'once the result lands it returns to 5s');
  // a stuck "syncing" row must not pin the dashboard at 2s forever
  response = syncing; await timer.advance(5000); await drain();
  assert.equal(ctx.dashPollDelayMs, 2000);
  for (let i = 0; i < 20; i++) await timer.advance(2000);
  assert.equal(ctx.dashPollDelayMs, 5000, 'fast polling gives up after its 30s window');
  // A failing server always wins over the fast cadence: a syncing result was
  // seen (dashSyncSince set) and then the polls start erroring — the dashboard
  // must ease off rather than hammer at 2s.
  ctx.dashSyncSince = ctx.Date.now();
  response = { status: 'error', message: 'server busy' };
  await timer.advance(5000); await drain();
  assert.ok(ctx.dashPollDelayMs > 5000, 'backoff still applies while a result is syncing');
  ctx.stopDashboardPolling();
});
test('dashboard shares overlapping refreshes and discards a response after switching sessions', async () => {
  const response = deferred(); let calls = 0, renders = 0;
  const { ctx } = dashboardContext(() => { calls++; return response.promise; });
  ctx.updatePendingList = () => renders++;
  const first = ctx.pollDashboard(), second = ctx.pollDashboard();
  assert.equal(first, second); assert.equal(calls, 1);
  ctx.sessionCode = 'TEST01'; response.resolve({ status: 'ok', pending: [] }); await first;
  assert.equal(renders, 0);
});
test('dashboard starts without overlap and automatically recovers after a timed-out request', async () => {
  let calls = 0;
  const { ctx, timer } = dashboardContext();
  ctx.fetch = () => ++calls === 1 ? new Promise(() => {}) : Promise.resolve({ ok: true, text: () => Promise.resolve('{"status":"ok"}') });
  load(ctx, helper(examiner)); ctx.apiGet = () => ctx.fetchJsonWithTimeout('synthetic', {});
  ctx.startDashboardPolling(); await drain(); await timer.advance(8000);
  assert.equal(calls, 1);
  await timer.advance(22000); assert.equal(ctx.failedPolls, 1);
  await timer.advance(12000); assert.equal(calls, 2); assert.equal(ctx.failedPolls, 0);
  ctx.stopDashboardPolling(); assert.equal(timer.jobs.size, 0);
});
test('create-session timeout offers read-only reconciliation without claiming the session was not created', () => {
  const ui = dom(); let reads = 0;
  const { ctx } = context({ ...ui, fetchPreviousSessions: () => reads++ });
  load(ctx, section(examiner, '  function examinerMutationErrorText(', "  document.getElementById('createSessionBtn').addEventListener"));
  const button = { disabled: true }, message = ui.nodes.get('examArea');
  ctx.showCreateSessionRequestError({ name: 'TimeoutError' }, button, message);
  assert.equal(button.disabled, false); assert.ok(message.textContent.includes('ייתכן שהסבב נוצר'));
  assert.equal(reads, 1); assert.equal(message.children.length, 1);
  message.children[0].click(); assert.equal(reads, 2);
  assert.ok(ctx.examinerMutationErrorText({ name: 'TimeoutError' }).includes('ייתכן שהפעולה בוצעה'));
});

function approvalContext(apiGet) {
  const ui = dom(); const setup = context({ ...ui, apiGet, examineeData: { idNumber: 'SYNTHETIC' },
    sessionCode: 'TEST00', examineeToken: 'synthetic', approvalPollCount: 0, approvalFailCount: 0,
    approvalInterval: null, examInProgress: false, updateApprovalDebug() {}, setupWaitingProtections() {}, teardownWaitingProtections() {},
    showApprovalError(message) { ui.nodes.get('approvalError').textContent = message; },
    notifyApprovedToExaminee() {}, showInstructions() {}, sessionData: {} });
  load(setup.ctx, section(examinee, '  function doCheckApproval()', '  function showApprovalError(msg)'));
  load(setup.ctx, section(examinee, '  var APPROVAL_POLL_BASE_MS', '  // Translation files are no longer'));
  withPacing(setup.ctx, examinee);
  return { ...setup, ...ui };
}
test('approval counts repeated server errors and stops on approval without overlapping the first poll', async () => {
  let response = { status: 'error', message: 'server busy' }, calls = 0;
  const { ctx, timer, nodes } = approvalContext(() => { calls++; return Promise.resolve(response); });
  for (let i = 0; i < 3; i++) await ctx.doCheckApproval();
  assert.equal(ctx.approvalFailCount, 3); assert.ok(nodes.get('approvalError').textContent.includes('server busy'));
  const pending = deferred(); ctx.apiGet = () => { calls++; return pending.promise; };
  const before = calls; ctx.startApprovalPolling(); await timer.advance(20000); assert.equal(calls - before, 1);
  pending.resolve({ status: 'ok', approval: 'approved' }); await drain();
  assert.equal(ctx.approvalFailCount, 0); assert.equal(ctx.approvalInterval, null); assert.equal(timer.jobs.size, 0);
});

test('an approval chain killed while the page was frozen restarts when the iPhone wakes', async () => {
  // iOS can suspend the page mid-request and never settle the promise, which
  // kills the self-rescheduling chain: the examinee then waits forever for an
  // approval that already happened (two iPhones, מחנה עמוס, 15/09/2026).
  const dead = deferred(); let calls = 0, approved = 0;
  const { ctx, timer, setVisibility } = approvalContext(() => { calls++; return calls === 1 ? dead.promise : Promise.resolve({ status: 'ok', approval: 'approved' }); });
  ctx.showInstructions = () => approved++;
  ctx.startApprovalPolling(); await drain();
  assert.equal(calls, 1, 'first check fired');
  await timer.advance(60000);
  assert.equal(calls, 1, 'the chain is dead: no further checks while the request never settles');
  setVisibility('hidden'); await drain();
  await timer.advance(4000);          // past the 3s "a healthy chain just ran" guard
  setVisibility('visible'); await drain();
  assert.equal(calls, 2, 'returning to the foreground restarts the chain immediately');
  assert.equal(approved, 1, 'the pending approval is picked up at once');
  // the resurrected chain must be the only one: a late answer from the dead
  // request cannot schedule a second chain alongside it
  dead.resolve({ status: 'ok', approval: 'pending' }); await drain();
  const settled = calls; await timer.advance(20000);
  assert.ok(calls - settled <= 1, 'no duplicate chain formed');
});
test('a healthy approval chain is not restarted by ordinary tab switching', async () => {
  let calls = 0;
  const { ctx, timer, setVisibility } = approvalContext(() => { calls++; return Promise.resolve({ status: 'ok', approval: 'pending' }); });
  ctx.startApprovalPolling(); await drain();
  const afterFirst = calls;
  setVisibility('hidden'); setVisibility('visible'); await drain();
  assert.equal(calls, afterFirst, 'a poll that just ran is not piled on');
  ctx.stopApprovalPolling && ctx.stopApprovalPolling();
  ctx.approvalInterval = null; timer.jobs.clear();
});

function startContext(apiGet) {
  const ui = dom(); const setup = context({ ...ui, apiGet,
    examineeData: { idNumber: 'SYNTHETIC', license: 'B', language: 'he' }, sessionData: {},
    sessionCode: 'TEST00', examineeToken: 'synthetic', examInProgress: false, examSubmitted: false,
    examStartConfirmed: false, TOTAL_QUESTIONS: 30, examDeadline: null,
    escHtml: value => String(value).replaceAll('&', '&amp;').replaceAll('<', '&lt;'),
    showScreen(id) { setup.ctx.currentScreen = id; }, showLoadingOverlay() {}, hideLoadingOverlay() {},
    isMobileDevice: () => true, _finishStartingExam() {} });
  setup.ctx.window = setup.ctx;
  load(setup.ctx, section(examinee, '  // ========== Start Exam ==========', '  // Second half of doStartExam'));
  return { ...setup, ...ui };
}
test('rate-limit retry honors waitSec and retains candidate, approval flow and seven-language request', async () => {
  const requests = []; let loaded = 0;
  const { ctx, timer, nodes, document } = startContext(params => {
    requests.push(params);
    return Promise.resolve(requests.length === 1 ? { status: 'error', rateLimited: true, waitSec: 60, message: 'slow down' }
      : { status: 'ok', questions: Array.from({ length: 30 }, (_, i) => ({ id: i + 1 })), translations: { ru: {} } });
  });
  ctx._finishStartingExam = () => loaded++;
  ctx.doStartExam(); await drain();
  assert.equal(ctx.currentScreen, 'screenExam'); assert.ok(!nodes.get('examArea').innerHTML.includes('אין מספיק שאלות'));
  const button = nodes.get('retryExamStartBtn'); assert.equal(button.disabled, true); assert.equal(document.activeElement.id, 'examStartError');
  button.click(); assert.equal(requests.length, 1);
  await timer.advance(59000); assert.equal(button.disabled, true);
  await timer.advance(1000); assert.equal(button.disabled, false); button.click(); await drain();
  assert.equal(loaded, 1); assert.equal(requests.length, 2);
  assert.ok(requests.every(p => p.action === 'getExamQuestions' && p.includeTranslations === 'true'));
  assert.equal(ctx.examineeData.idNumber, 'SYNTHETIC'); assert.equal(ctx.examineeToken, 'synthetic');
  assert.equal(ctx.examDeadline, null); assert.equal(ctx.examInProgress, false);
});
test('busy, old-server errors and network errors provide an accessible retry without a false question-bank diagnosis', async () => {
  for (const response of [{ status: 'error', code: 'question_cache_busy', retryable: true, waitSec: 3, message: 'busy' },
    { status: 'error', message: 'legacy server error' }, null]) {
    const { ctx, timer, nodes } = startContext(() => response ? Promise.resolve(response) : Promise.reject(new Error('offline')));
    ctx.doStartExam(); await drain();
    assert.ok(nodes.get('retryExamStartBtn')); assert.ok(!nodes.get('examArea').innerHTML.includes('אין מספיק שאלות'));
    if (response && response.retryable) { assert.equal(nodes.get('retryExamStartBtn').disabled, true); await timer.advance(3000); }
    assert.equal(nodes.get('retryExamStartBtn').disabled, false); assert.equal(ctx.examineeToken, 'synthetic');
  }
});
test('a Google error page or a timeout during exam start is presented as temporary with a cooldown', async () => {
  for (const [name, expectedText] of [['SyntaxError', 'Google'], ['TimeoutError', 'לא ענה בזמן']]) {
    const err = new Error(name === 'SyntaxError' ? 'Unexpected token <' : 'Request timed out'); err.name = name;
    const { ctx, timer, nodes } = startContext(() => Promise.reject(err));
    ctx.doStartExam(); await drain();
    assert.ok(nodes.get('retryExamStartBtn'), 'retry button rendered');
    assert.ok(nodes.get('examArea').innerHTML.includes(expectedText), 'honest message: ' + expectedText);
    assert.ok(nodes.get('examArea').innerHTML.includes('השרת עמוס כרגע'), 'classified as temporary');
    assert.equal(nodes.get('retryExamStartBtn').disabled, true, 'cooldown armed');
    await timer.advance(5000);
    assert.equal(nodes.get('retryExamStartBtn').disabled, false, 'retry enabled after the 5s cooldown');
    assert.equal(ctx.examineeToken, 'synthetic', 'registration preserved');
  }
});
test('a stale question response or retry button cannot start the next candidate', async () => {
  const response = deferred(); let loaded = 0;
  const { ctx, nodes } = startContext(() => response.promise); ctx._finishStartingExam = () => loaded++;
  ctx.doStartExam(); ctx.examStartGeneration++; ctx.examineeData.idNumber = 'NEXT';
  response.resolve({ status: 'ok', questions: Array(30).fill({ id: 1 }) }); await drain();
  assert.equal(loaded, 0); assert.equal(nodes.get('retryExamStartBtn'), undefined);
});
test('registration remains confirmed-before-start and stale registration retries cannot affect a new candidate', async () => {
  const { ctx, timer } = startContext(); let began = 0, calls = 0;
  ctx.prefetchExamImages = () => {}; ctx.getQAnswers = () => ['a', 'b', 'c', 'd'];
  ctx.getQCorrectIndex = () => 1; ctx.shuffle = values => values; ctx._beginExam = () => began++;
  load(ctx, section(examinee, '  function _finishStartingExam(', '  // Second half of exam start'));
  ctx.examStartGeneration = 1;
  const attempt = { generation: 1, idNumber: 'SYNTHETIC', sessionCode: 'TEST00', examineeToken: 'synthetic' };
  const response = deferred(); ctx.apiPost = payload => { calls++; assert.equal(payload.questions.length, 30); return response.promise; };
  ctx._finishStartingExam(Array.from({ length: 30 }, (_, id) => ({ id })), 'B', 'he', attempt);
  assert.equal(began, 0); response.resolve({ status: 'ok', examStarted: true }); await drain();
  assert.equal(began, 1); assert.equal(ctx.examStartConfirmed, true);
  ctx.apiPost = () => { calls++; return Promise.reject(new Error('offline')); };
  ctx._finishStartingExam(Array.from({ length: 30 }, (_, id) => ({ id })), 'B', 'he', attempt); await drain();
  ctx.examStartGeneration++; ctx.examineeToken = 'new-token'; await timer.advance(1500);
  assert.equal(calls, 2); assert.equal(began, 1);
});

function dqContext(apiGet) {
  const setup = context({ apiGet, sessionCode: 'TEST00', examineeData: { idNumber: 'SYNTHETIC' }, examineeToken: 'synthetic',
    examRetryWaitSeconds: data => Number(data && data.waitSec) || 0, restoreExamFromSuspended() {}, showFinalDQScreen() {} });
  load(setup.ctx, section(examinee, '  var dqOverturnInterval = null;', '  function showFinalDQScreen()'));
  withPacing(setup.ctx, examinee);
  return setup;
}
test('DQ polling never overlaps, backs off on failure and stops immediately on an overturn', async () => {
  let pending = deferred(), calls = 0, restored = 0;
  const { ctx, timer } = dqContext(() => { calls++; return pending.promise; }); ctx.restoreExamFromSuspended = () => restored++;
  ctx.startDQOverturnPolling(); await timer.advance(3000); await timer.advance(20000); assert.equal(calls, 1);
  pending.resolve({ status: 'error', waitSec: 6 }); await drain(); pending = deferred();
  await timer.advance(5999); assert.equal(calls, 1); await timer.advance(1); assert.equal(calls, 2);
  pending.resolve({ status: 'ok', approval: 'in_exam' }); await drain();
  assert.equal(restored, 1); assert.equal(timer.jobs.size, 0); assert.equal(ctx.dqOverturnInterval, null);
});
test('DQ polling ignores the result of a stopped generation', async () => {
  const pending = deferred(); let restored = 0;
  const { ctx, timer } = dqContext(() => pending.promise); ctx.restoreExamFromSuspended = () => restored++;
  ctx.startDQOverturnPolling(); await timer.advance(3000); ctx.stopDQOverturnPolling();
  pending.resolve({ status: 'ok', approval: 'in_exam' }); await drain();
  assert.equal(restored, 0); assert.equal(timer.jobs.size, 0);
});

const resultPayload = token => ({ action: 'submitResult', idNumber: 'SYNTHETIC', sessionCode: 'TEST00', examineeToken: token });
function submitContext(apiPost) {
  const stored = new Map(), statuses = [], warnings = [];
  const ui = dom();
  const setup = context({ ...ui, apiPost, API_ORIGIN: 'examinee-app', examineeData: { idNumber: 'SYNTHETIC' }, sessionCode: 'TEST00', examineeToken: 'synthetic',
    localStorage: { get length() { return stored.size; }, key: i => [...stored.keys()][i],
      getItem: key => stored.get(key) || null, setItem: (key, value) => stored.set(key, value), removeItem: key => stored.delete(key) },
    hasAnyPendingResult: () => [...stored.keys()].some(key => key.startsWith('pendingResult_')),
    disarmPendingResultGuard() {}, armPendingResultGuard() {}, clearSubmitFailureBanner() {},
    showSubmitStatus: status => statuses.push(status), showSubmitFailureBanner: (...args) => warnings.push(args) });
  load(setup.ctx, section(examinee, '  var _submitInFlight = {};', '  // Resend a single pending'));
  const payload = resultPayload('synthetic'); stored.set('pendingResult_SYNTHETIC', JSON.stringify(payload));
  stored.set('pendingWrongAnswers_SYNTHETIC', '[]');
  return { ...setup, ...ui, stored, statuses, warnings, payload };
}
test('successful manual retry cancels old submit timers and clears only confirmed pending data', async () => {
  let calls = 0;
  const { ctx, timer, stored, statuses, payload } = submitContext(() => Promise.resolve(++calls === 1 ? { status: 'error' } : { status: 'ok' }));
  ctx.submitWithRetry(payload, 3, []); await drain(); assert.ok(stored.has('pendingResult_SYNTHETIC'));
  const oldTimer = [...timer.jobs.values()][0].cb;
  ctx.submitWithRetry(payload, 3, []); await drain();
  assert.equal(timer.jobs.size, 0); assert.equal(stored.size, 0); assert.deepEqual(statuses, ['received']);
  oldTimer(); await drain(); await timer.advance(60000); assert.equal(calls, 2);
});
test('token mismatch preserves pending result and wrong answers for examiner recovery', async () => {
  let calls = 0;
  const { ctx, timer, stored, warnings, nodes, payload } = submitContext(() => { calls++; return Promise.resolve({ status: 'error', examineeTokenError: 'mismatch' }); });
  ctx.submitWithRetry(payload, 3, []); await drain();
  assert.equal(stored.size, 2); assert.equal(timer.jobs.size, 0); assert.ok(warnings[0][2]);
  assert.ok(nodes.get('pendingRecoveryNotice').textContent); assert.ok(!nodes.get('pendingRecoveryNotice').textContent.includes('SYNTHETIC'));
  ctx.submitWithRetry(payload, 3, []); await timer.advance(60000); assert.equal(calls, 1);
});
test('finishing a later attempt preserves the unconfirmed earlier result and cleanup is attempt-specific', () => {
  const { ctx, stored, payload } = submitContext();
  const newer = resultPayload('new-token');
  ctx.persistPendingResult(newer, [{ questionId: 'synthetic-question' }]);
  const results = [...stored.entries()].filter(([key]) => key.startsWith('pendingResult_'));
  assert.equal(results.length, 2);
  assert.ok(results.some(([, raw]) => JSON.parse(raw).examineeToken === 'synthetic'));
  assert.ok(results.some(([, raw]) => JSON.parse(raw).examineeToken === 'new-token'));
  ctx.clearConfirmedPendingResult(payload);
  assert.equal([...stored.keys()].filter(key => key.startsWith('pendingResult_')).length, 1);
  const [newKey, newRaw] = [...stored.entries()].find(([key]) => key.startsWith('pendingResult_'));
  assert.equal(JSON.parse(newRaw).examineeToken, 'new-token');
  assert.equal(JSON.parse(newRaw).origin, 'examinee-app');
  assert.equal(JSON.parse(stored.get('pendingWrongAnswers_' + newKey.slice('pendingResult_'.length)))[0].questionId, 'synthetic-question');
});
test('a rejected previous result is visible even when another examinee is using the device', async () => {
  const { ctx, nodes, stored, payload, statuses } = submitContext(() => Promise.resolve({ status: 'error', examineeTokenError: 'mismatch' }));
  ctx.examineeData.idNumber = 'NEXT'; ctx.examineeToken = 'next-token';
  ctx.submitWithRetry(payload, 3, []); await drain();
  assert.ok(nodes.get('pendingRecoveryNotice')); assert.equal(stored.size, 2); assert.deepEqual(statuses, []);
});
test('two-tap dismiss removes only a superseded (blocked) result and clears the notice', async () => {
  const { ctx, nodes, stored, payload } = submitContext(() => Promise.resolve({ status: 'error', examineeTokenError: 'mismatch' }));
  ctx.submitWithRetry(payload, 3, []); await drain();
  assert.ok(nodes.get('pendingRecoveryNotice'));
  assert.equal([...stored.keys()].filter(k => k.startsWith('pendingResult_')).length, 1);
  const btn = nodes.get('pendingRecoveryDismiss');
  assert.ok(btn, 'dismiss button is rendered');
  btn.click();                                   // first tap only arms — nothing removed yet
  assert.match(btn.textContent, /לחץ שוב/);
  assert.equal([...stored.keys()].filter(k => k.startsWith('pendingResult_')).length, 1);
  btn.click();                                   // second tap removes the undeliverable result
  assert.equal([...stored.keys()].filter(k => k.startsWith('pendingResult_')).length, 0);
  assert.equal(nodes.get('pendingRecoveryNotice') || null, null);
});
test('old submit response cannot erase or falsely confirm a newer result sharing the same ID', async () => {
  const pending = deferred(); let calls = 0;
  const { ctx, timer, stored, statuses, payload } = submitContext(() => { calls++; return calls === 1 ? pending.promise : Promise.resolve({ status: 'ok' }); });
  ctx.submitWithRetry(payload, 3, []);
  const newer = resultPayload('new-token'); ctx.examineeToken = 'new-token'; ctx.persistPendingResult(newer, []);
  ctx.submitWithRetry(newer, 3, []); assert.equal(calls, 1);
  pending.resolve({ status: 'ok' }); await drain();
  const remaining = [...stored.entries()].filter(([key]) => key.startsWith('pendingResult_'));
  assert.equal(remaining.length, 1); assert.equal(JSON.parse(remaining[0][1]).examineeToken, 'new-token'); assert.deepEqual(statuses, []);
  await timer.advance(3000); assert.equal(calls, 2); assert.equal(stored.size, 0); assert.deepEqual(statuses, ['received']);
});

// Execute every inline script in original order. Only browser surfaces and the
// network are synthetic; API, finish, storage, retry and bootstrap functions are
// never replaced. Expose test entry points at the end of the existing closure.
// This complements real-browser checks; it does not emulate browser layout or
// native beforeunload dialogs.
function memoryStore() {
  const entries = new Map(), writes = [];
  const store = {
    entries, writes, rejectWrite: () => false,
    get length() { return entries.size; }, key: i => [...entries.keys()][i] ?? null,
    getItem: key => entries.get(String(key)) ?? null,
    setItem(key, value) {
      key = String(key); value = String(value); writes.push({ key, value });
      if (store.rejectWrite(key, value)) throw Object.assign(new Error('Synthetic storage quota'), { name: 'QuotaExceededError' });
      entries.set(key, value);
    }, removeItem: key => entries.delete(String(key))
  };
  return store;
}
function savedResults(store) {
  return [...store.entries].filter(([key]) => key.startsWith('pendingResult_')).flatMap(([key, raw]) => {
    try { return [{ key, payload: JSON.parse(raw) }]; } catch { return []; }
  });
}
function completePage({ local = memoryStore(), session = memoryStore(), reply = () => Promise.reject(new Error('Synthetic offline')) } = {}) {
  const ui = dom(), requests = [], beacons = [], windowEvents = new Map(), documentEvents = new Map();
  for (const tag of examinee.slice(0, examinee.indexOf('<script')).matchAll(/<[^>]+\bid="([^"]+)"[^>]*>/g)) {
    const el = ui.element(tag[1]), classes = /\bclass="([^"]+)"/.exec(tag[0]);
    if (classes) classes[1].split(/\s+/).forEach(value => el.classList.add(value));
    el.value = ''; el.options = [];
  }
  function eventTarget(target, events) {
    target.addEventListener = (name, cb) => { if (!events.has(name)) events.set(name, new Set()); events.get(name).add(cb); };
    target.removeEventListener = (name, cb) => events.get(name)?.delete(cb);
  }
  eventTarget(ui.document, documentEvents);
  ui.document.head = ui.document.body;
  ui.document.documentElement = { style: {}, setAttribute() {} };
  ui.document.visibilityState = 'visible';
  ui.document.querySelectorAll = selector => selector === '.screen' ? [...ui.nodes.values()].filter(node => node.classList.contains('screen')) : [];
  ui.document.querySelector = selector => selector === '.screen.active' ? [...ui.nodes.values()].find(node => node.classList.contains('screen') && node.classList.contains('active')) || null : null;
  const setup = context({ ...ui, localStorage: local, sessionStorage: session, URL, URLSearchParams, Blob,
    navigator: { userAgent: 'Synthetic test browser', platform: 'Synthetic', maxTouchPoints: 0, onLine: true,
      sendBeacon(url, blob) { beacons.push(blob); return true; } },
    history: { pushState() {} }, location: { search: '', pathname: '/synthetic/examinee.html', reload() { throw new Error('Unexpected automatic reload'); } },
    fetch(url, opts = {}) {
      if (opts.method === 'HEAD') return Promise.resolve({ headers: { get: () => null } });
      const payload = opts.body ? JSON.parse(opts.body) : Object.fromEntries(new URL(url).searchParams);
      requests.push(payload);
      return Promise.resolve(reply(payload)).then(data => ({ ok: true, text: () => Promise.resolve(JSON.stringify(data)) }));
    }
  });
  setup.ctx.window = setup.ctx;
  setup.ctx.setInterval = setup.timer.set;
  eventTarget(setup.ctx, windowEvents);
  const exposure = `
  window.__testExam = {
    seed: function(token) {
      sessionCode='TEST00'; examineeToken=token||'new-token';
      examineeData={idNumber:'SYNTHETIC',fullName:'Synthetic Test',license:'B',language:'he'};
      sessionData={license:'B',language:'he',audioMode:'off'};
      activeQuestions=Array.from({length:30},function(_,i){return {id:i+1,text:'Synthetic question',answers:['A','B','C','D'],ci:1,category:'חוק'};});
      shuffledOrders=activeQuestions.map(function(){return {order:[0,1,2,3],correctIdx:1};});
      languageHistory=['he','ar','ru','en','fr','es','am'];
      userAnswers=activeQuestions.map(function(_,i){return {chosenIndex:1,isCorrect:true,langAtAnswer:languageHistory[i%7]};});
      examInProgress=true; examSubmitted=false; examStartTime=Date.now();
      examDeadline=Date.now()+2400000; timeRemaining=2400; examTimeMinutes=40;
      saveExamineeState('screenExam'); saveActiveExam();
    },
    finish:renderExamDone,persist:persistPendingResult,flush:flushPendingResult,resend:resendAllPendingResults,
    submit:submitWithRetry,beacon:beaconAllPendingResults,hasPending:hasAnyPendingResult,
    state:function(){return {inProgress:examInProgress,submitted:examSubmitted,id:examineeData.idNumber,token:examineeToken,
      memoryOnly:Object.keys(_pendingMemoryOnly).length,guard:pendingResultGuardArmed};}
  };
`;
  const scripts = [...examinee.matchAll(/<script\b[^>]*>([\s\S]*?)<\/script>/gi)];
  for (let i = 0; i < scripts.length; i++) {
    let code = scripts[i][1];
    if (code.includes('function renderExamDone()')) {
      const end = code.lastIndexOf('})();'); assert.ok(end >= 0);
      code = code.slice(0, end) + exposure + code.slice(end);
    }
    vm.runInContext(code, setup.ctx, { filename: 'examinee.html:inline-' + (i + 1) });
  }
  assert.ok(setup.ctx.__testExam, 'entire main script reached its end');
  return { ...setup, ...ui, local, session, requests, beacons, windowEvents, exam: setup.ctx.__testExam,
    dispatch(name, event = {}) { for (const cb of [...(windowEvents.get(name) || [])]) cb(event); } };
}

test('complete page: corrupt prior result cannot block saving the finished 30-answer exam', async () => {
  const page = completePage(); page.exam.seed();
  page.local.entries.set('pendingResult_SYNTHETIC', '{corrupt');
  page.exam.finish(); await drain();
  assert.equal(page.local.getItem('pendingResult_SYNTHETIC'), '{corrupt');
  const [saved] = savedResults(page.local);
  assert.equal(saved.payload.examineeToken, 'new-token'); assert.equal(saved.payload.answers.length, 30);
  assert.equal(saved.payload.score, 30); assert.equal(saved.payload.percent, 100); assert.equal(saved.payload.passed, true);
  assert.equal(new Set(saved.payload.answers.map(answer => answer.langAtAnswer)).size, 7);
  assert.equal(page.session.getItem('ext_exam_active'), null); assert.equal(page.exam.state().submitted, true);
  let prevented = false; page.dispatch('beforeunload', { preventDefault() { prevented = true; } });
  assert.equal(prevented, true); assert.equal(page.exam.hasPending(), true);
});

test('complete page: oversized older result is not copied before the new attempt is saved', async () => {
  const page = completePage(); page.exam.seed();
  const oldRaw = JSON.stringify({ ...resultPayload('old-token'), answers: Array(30).fill({ text: 'x'.repeat(1000) }) });
  page.local.entries.set('pendingResult_SYNTHETIC', oldRaw);
  page.local.rejectWrite = (key, value) => key.startsWith('pendingResult_') && value.length > 10000;
  page.exam.finish(); await drain();
  assert.equal(page.local.getItem('pendingResult_SYNTHETIC'), oldRaw);
  assert.equal(savedResults(page.local).length, 2); assert.equal(savedResults(page.session).length, 0);
  assert.ok(!page.local.writes.some(write => write.value === oldRaw));
});

test('complete page: rejected attempt-key write preserves old local result and recovers session fallback after reload', async () => {
  const page = completePage(); page.exam.seed();
  const oldRaw = JSON.stringify(resultPayload('old-token'));
  page.local.entries.set('pendingResult_SYNTHETIC', oldRaw);
  page.local.rejectWrite = key => key.startsWith('pendingResult_') && key.includes('__');
  page.exam.finish(); await drain();
  assert.equal(page.local.getItem('pendingResult_SYNTHETIC'), oldRaw);
  assert.equal(savedResults(page.session)[0].payload.answers.length, 30);
  assert.ok(page.nodes.get('pendingTabStorageNotice').textContent.includes('סגירת הלשונית'));
  assert.equal(page.nodes.get('doneLeaveText').style.display, 'none');
  // A fresh VM with the same storage executes actual early bootstrap again.
  const reload = completePage({ local: page.local, session: page.session }); await drain();
  assert.equal(reload.exam.state().inProgress, false); assert.equal(reload.exam.state().guard, true);
  assert.ok(reload.nodes.get('pendingTabStorageNotice').textContent.includes('טרם אושרה'));
  assert.equal(reload.nodes.get('doneLeaveText').style.display, 'none');
  await reload.timer.advance(3000);
  assert.ok(reload.requests.some(payload => payload.action === 'submitResult' && payload.examineeToken === 'new-token'));
  assert.equal(savedResults(reload.session)[0].payload.answers.length, 30);
});

test('complete page: session-only warning survives transport failures and disappears only after confirmation', async () => {
  const page = completePage(); page.exam.seed(); page.local.rejectWrite = () => true;
  page.exam.finish(); await drain(); await page.timer.advance(6000);
  assert.ok(page.nodes.get('submitFailBanner'), 'actual transport-failure banner was rendered');
  assert.ok(page.nodes.get('pendingTabStorageNotice').textContent.includes('סגירת הלשונית'));
  assert.equal(page.nodes.get('doneLeaveText').style.display, 'none');
  const confirmation = deferred();
  const reload = completePage({ local: page.local, session: page.session, reply: () => confirmation.promise });
  await drain(); assert.ok(reload.nodes.get('pendingTabStorageNotice'));
  confirmation.resolve({ status: 'ok' }); await drain();
  assert.equal(savedResults(reload.session).length, 0); assert.equal(reload.exam.state().guard, false);
  assert.equal(reload.nodes.get('pendingTabStorageNotice'), undefined);
});

test('complete page: failure of both stores warns, blocks next candidate and cannot reopen a finished exam after forced reload', async () => {
  const page = completePage(); page.exam.seed();
  page.local.rejectWrite = page.session.rejectWrite = () => true;
  page.exam.finish(); await drain();
  assert.equal(page.exam.state().memoryOnly, 1); assert.equal(page.exam.state().guard, true);
  assert.ok(page.nodes.get('pendingStorageFailure').textContent.includes('רק בזיכרון הדף'));
  assert.equal(page.nodes.get('doneLeaveText').style.display, 'none');
  page.ctx.resetForNextExaminee(); assert.equal(page.exam.state().id, 'SYNTHETIC'); assert.equal(page.exam.state().submitted, true);
  page.exam.beacon(); const beaconPayloads = await Promise.all(page.beacons.map(async blob => JSON.parse(await blob.text())));
  assert.ok(beaconPayloads.some(payload => payload.action === 'submitResult' && payload.answers.length === 30));
  assert.equal(page.session.getItem('ext_exam_active'), null);
  const forcedReload = completePage({ local: page.local, session: page.session }); await drain();
  assert.equal(forcedReload.exam.state().inProgress, false); assert.equal(forcedReload.exam.state().id, '');
  // Both unavailable stores cannot provide forced-reload durability. Never
  // conceal that limitation by resuming the finished active-exam backup.
  assert.equal(forcedReload.exam.hasPending(), false);
});

test('complete page: real apiPost may attach a legacy token without changing confirmed cleanup identity', async () => {
  const page = completePage({ reply: () => ({ status: 'ok' }) }); page.exam.seed();
  const old = resultPayload(undefined); delete old.examineeToken;
  page.local.entries.set('pendingResult_SYNTHETIC', JSON.stringify(old));
  page.exam.flush('SYNTHETIC'); await drain();
  assert.equal(page.requests.find(payload => payload.action === 'submitResult').examineeToken, 'new-token');
  assert.equal(page.local.getItem('pendingResult_SYNTHETIC'), null); assert.equal(page.exam.state().guard, false);
});

test('complete page: a late older acknowledgement leaves the new attempt queued and stored', async () => {
  const olderAck = deferred(), newerAck = deferred();
  const page = completePage({ reply: payload => payload.examineeToken === 'old-token' ? olderAck.promise : newerAck.promise });
  page.exam.seed('old-token'); page.exam.finish(); await drain();
  page.exam.seed('new-token'); page.exam.finish(); await drain();
  assert.equal(page.requests.filter(payload => payload.action === 'submitResult').length, 1);
  assert.equal(savedResults(page.local).length, 2);
  olderAck.resolve({ status: 'ok' }); await drain();
  const remaining = savedResults(page.local);
  assert.equal(remaining.length, 1); assert.equal(remaining[0].payload.examineeToken, 'new-token');
  assert.equal(page.exam.hasPending(), true);
  await page.timer.advance(3000);
  assert.equal(page.requests.filter(payload => payload.action === 'submitResult').length, 2);
  assert.equal(savedResults(page.local).length, 1);
  newerAck.resolve({ status: 'ok' }); await drain();
  assert.equal(savedResults(page.local).length, 0); assert.equal(page.exam.state().guard, false);
});

for (const code of ['submission_busy', 'result_commit_busy', 'result_commit_uncertain']) {
  test('complete page: ' + code + ' keeps retrying past fast attempts and clears only after a successful acknowledgement', async () => {
    let accepted = false;
    const page = completePage({ reply: () => accepted ? { status: 'ok' } : { status: 'error', code, retryable: true, waitSec: 3 } });
    page.exam.seed(); page.exam.finish(); await drain();
    await page.timer.advance(60000);
    const submits = page.requests.filter(payload => payload.action === 'submitResult');
    assert.ok(submits.length >= 5, 'slow retries continue after the three fast attempts');
    assert.ok(submits.every(payload => payload.examineeToken === 'new-token' && payload.answers.length === 30));
    assert.equal(savedResults(page.local).length, 1); assert.equal(page.exam.state().guard, true);
    assert.ok(page.nodes.get('submitFailBanner').innerHTML.includes('ממשיכה לנסות'));
    assert.ok(!page.nodes.get('submitStatusBanner').innerHTML.includes('התקבלה ונשמרה'));
    accepted = true; page.dispatch('online'); await drain();
    assert.equal(savedResults(page.local).length, 0); assert.equal(page.exam.state().guard, false);
    assert.ok(page.nodes.get('submitStatusBanner').innerHTML.includes('התקבלה ונשמרה'));
    const count = page.requests.length;
    await page.timer.advance(60000);
    assert.equal(page.requests.length, count, 'successful online retry cancelled the scheduled retry');
  });
}

test('complete page: online flush shares an in-flight request while pagehide beacon retains the same attempt identity', async () => {
  const first = deferred();
  const page = completePage({ reply: () => first.promise }); page.exam.seed(); page.exam.finish(); await drain();
  page.dispatch('online'); page.dispatch('online'); page.exam.resend(); await drain();
  assert.equal(page.requests.filter(payload => payload.action === 'submitResult').length, 1);
  page.dispatch('pagehide');
  const beaconPayloads = await Promise.all(page.beacons.map(async blob => JSON.parse(await blob.text())));
  const resultBeacons = beaconPayloads.filter(payload => payload.action === 'submitResult');
  assert.equal(resultBeacons.length, 1); assert.equal(resultBeacons[0].examineeToken, 'new-token');
  assert.equal(resultBeacons[0].answers.length, 30);
  assert.equal(savedResults(page.local).length, 1, 'beacon enqueue is not a persistence acknowledgement');
  first.resolve({ status: 'error', code: 'result_commit_uncertain', retryable: true, waitSec: 3 }); await drain();
  assert.equal(savedResults(page.local).length, 1); assert.equal(page.exam.state().guard, true);
});

// ===== "you have no open session" must mean exactly that =====================
// 15/09/2026: an examiner with a live session opened the dashboard and saw
// nothing on his name. The request had failed (the endpoint answered in 35.5s,
// then 2.2s on retry), and this loader swallowed it — an empty `.catch` plus a
// bare `return` on a non-ok status rendered the identical empty screen as a
// genuine "no active sessions". The next move an examiner makes from that
// screen is to open a SECOND session, splitting the class between the session
// the examinees registered to and the one the examiner is watching.
function prevSessionsPage(apiGet) {
  const ui = dom();
  ui.element('prevSessionsArea'); ui.element('prevSessionsList');
  const setup = context({ ...ui, apiGet, examinerData: { id: '111', name: 'בוחן' },
    LICENSE_LABELS: { B: 'דרגה B' }, escHtml: value => String(value == null ? '' : value),
    resumeSession() {} });
  setup.ctx.window = setup.ctx;
  load(setup.ctx, section(examiner, '  // ========== Fetch Previous Sessions ==========',
    '  // ========== Commander: list all active sessions'));
  return { ...setup, ...ui };
}
const shown = ui => {
  const area = ui.nodes.get('prevSessionsArea'), list = ui.nodes.get('prevSessionsList');
  const texts = (list.children || []).flatMap(card => (card.children || []).map(part => part.textContent || ''));
  return { visible: area.style.display === 'block', cards: (list.children || []).length,
    failed: texts.some(text => text.indexOf('לא הצלחנו לטעון') >= 0), retry: (list.children || [])[0] };
};
const liveSession = () => ({ status: 'ok', sessions: [{ code: 'ABC12345', active: true, site: 'אתר', classroom: '1',
  license: 'B', validUntil: new Date(Date.now() + 3 * 3600 * 1000).toISOString() }] });

test('a failed session load says so instead of looking like "no open session"', async () => {
  for (const [label, apiGet] of [
    ['rejected request', () => Promise.reject(Object.assign(new Error('timed out'), { name: 'TimeoutError' }))],
    ['server error', () => Promise.resolve({ status: 'error', message: 'server busy' })],
    ['malformed payload', () => Promise.resolve({ status: 'ok' })],
  ]) {
    const page = prevSessionsPage(apiGet);
    page.ctx.fetchPreviousSessions(); await drain();
    const state = shown(page);
    assert.equal(state.visible, true, label + ': the failure is on screen');
    assert.equal(state.failed, true, label + ': it names the failure, not an empty list');
  }
});
test('an examiner with a live session still sees it, and a genuine empty result stays silent', async () => {
  const live = prevSessionsPage(() => Promise.resolve(liveSession()));
  live.ctx.fetchPreviousSessions(); await drain();
  assert.equal(shown(live).cards, 1); assert.equal(shown(live).failed, false);

  for (const empty of [{ status: 'ok', sessions: [] },
    { status: 'ok', sessions: [{ code: 'OLD', active: true, validUntil: new Date(Date.now() - 1000).toISOString() }] }]) {
    const page = prevSessionsPage(() => Promise.resolve(empty));
    page.ctx.fetchPreviousSessions(); await drain();
    assert.equal(shown(page).failed, false, 'nothing active is not an error');
    assert.equal(shown(page).visible, false, 'and it must not be confusable with one');
  }
});
test('retrying after a failure recovers the session list', async () => {
  let reply = () => Promise.reject(new Error('network'));
  const page = prevSessionsPage(() => reply());
  page.ctx.fetchPreviousSessions(); await drain();
  assert.equal(shown(page).failed, true);
  reply = () => Promise.resolve(liveSession());
  const box = shown(page).retry;
  box.children.find(child => child.handlers && child.handlers.click).click(); await drain();
  assert.equal(shown(page).failed, false, 'the failure box is gone');
  assert.equal(shown(page).cards, 1, 'and the real session took its place');
});

// ===== 16/09/2026: the examiner self-update check took the examiners down =====
// From ~09:30 every examiner in an exam was thrown to the login screen, and the
// pages that came back reloaded forever under "גרסה חדשה זמינה" — with NO deploy
// since 07:57. Measured that day: GitHub Pages serves ONE build as
// "6aaa21d2-76fe0" (identity) or W/"6aaa21d2-76fe0" (gzip). The check compared
// them as strings, trusted error responses, and reloaded mid-exam.
// The verbatim pre-fix block is kept here so the harness is proven to REPRODUCE
// the incident — a test that only passes on the new code proves nothing.
const UPDATE_CHECK_BEFORE_16_09 = String.raw`(function() {
    if (!window.fetch) return;
    var baseTag = null, notified = false;
    var tagOf = function(r) { return r.headers.get('ETag') || r.headers.get('Last-Modified') || null; };
    var probe = function(cb) {
      fetch(location.pathname, { method: 'HEAD', cache: 'no-store' })
        .then(function(r) { cb(tagOf(r)); }).catch(function() { cb(null); });
    };
    probe(function(t) { baseTag = t; });
    setInterval(function() {
      if (notified) return;
      probe(function(t) {
        if (!t || !baseTag || t === baseTag) return;
        notified = true;
        var b = document.createElement('div');
        b.style.cssText = 'position:fixed;left:0;right:0;bottom:0;z-index:99999;background:#1a73e8;color:#fff;padding:12px 16px;text-align:center;font-size:15px;font-weight:700;box-shadow:0 -2px 10px rgba(0,0,0,.2);';
        b.innerHTML = '🔄 גרסה חדשה זמינה — מתעדכן אוטומטית… <button id="swUpdNow" style="margin-right:10px;padding:6px 16px;font-weight:800;background:#fff;color:#1a73e8;border:none;border-radius:8px;cursor:pointer;">רענן עכשיו</button>';
        document.body.appendChild(b);
        var btn = document.getElementById('swUpdNow');
        if (btn) btn.addEventListener('click', function() { location.reload(); });
        setTimeout(function() { location.reload(); }, 60000);  // grace period to finish an action
      });
    }, 120000);  // check every 2 min
  })();`;
function currentUpdateCheck() {
  const src = examiner.replace(/\r/g, '');   // the working copy is CRLF on Windows
  const a = src.indexOf('(function() {\n    if (!window.fetch) return;');
  const t = src.indexOf('// check every 2 min', a);
  const e = src.indexOf('})();', t) + 5;
  assert.ok(a > 0 && t > a && e > t, 'update-check block located in examiner.html');
  return src.slice(a, e);
}
// Runs the real block against a scripted sequence of HEAD answers, on a fake clock.
function runUpdateCheck(code, answers, { session = '', globals = {} } = {}) {
  let now = 0, reloads = 0, banners = 0; const jobs = []; const banner = {};
  const queue = answers.slice();
  const ctx = {
    ...globals,
    window: { fetch: true }, String, Promise, sessionCode: session,
    fetch: () => {
      const a = queue.length ? queue.shift() : null;
      if (!a || a.fail) return Promise.reject(new Error('network'));
      return Promise.resolve({ ok: a.status === undefined ? true : a.status >= 200 && a.status < 300,
        headers: { get: n => (n === 'ETag' ? (a.etag || null) : n === 'Last-Modified' ? (a.lm || null) : null) } });
    },
    location: { pathname: '/driving-theory-exam/examiner.html', reload() { reloads++; } },
    document: { createElement: () => { const el = { style: {} }; Object.defineProperty(el, 'innerHTML', { set(v) { banner.html = v; }, get() { return banner.html; } }); return el; },
      body: { appendChild() { banners++; } }, getElementById: () => null, querySelector: () => null },
    setInterval: (fn, ms) => { jobs.push({ at: now + ms, fn, every: ms }); },
    setTimeout: (fn, ms) => { jobs.push({ at: now + ms, fn, every: 0 }); },
  };
  vm.createContext(ctx); vm.runInContext(code, ctx);
  const settle = async () => { for (let i = 0; i < 20; i++) await Promise.resolve(); };
  return (async () => {
    await settle();                                     // the base probe
    const until = 2 * 60 * 60 * 1000;                   // two hours of an open dashboard
    for (let guard = 0; guard < 10000; guard++) {
      if (reloads) break;                               // a real reload tears the page (and this timer) down
      jobs.sort((x, y) => x.at - y.at);
      const j = jobs[0];
      if (!j || j.at > until) break;
      jobs.shift(); now = j.at; j.fn(); await settle();
      if (j.every) jobs.push({ at: now + j.every, fn: j.fn, every: j.every });
    }
    return { reloads, banners, banner: banner.html || '' };
  })();
}
const IDENTITY = { etag: '"6aaa21d2-76fe0"' }, GZIP = { etag: 'W/"6aaa21d2-76fe0"' };
const flapping = n => Array.from({ length: n }, (_, i) => (i % 2 ? GZIP : IDENTITY));

test('update check: the harness reproduces the 16/09 incident on the pre-fix code', async () => {
  // base = identity, first poll = gzip spelling of the SAME build
  const r = await runUpdateCheck(UPDATE_CHECK_BEFORE_16_09, [IDENTITY, GZIP]);
  assert.equal(r.banners, 1, 'old code announced a "new version" that did not exist');
  assert.equal(r.reloads, 1, 'and reloaded the examiner out of the dashboard');
  const mid = await runUpdateCheck(UPDATE_CHECK_BEFORE_16_09, [IDENTITY, GZIP], { session: 'LIVE0001' });
  assert.equal(mid.reloads, 1, 'even in the middle of an open exam session');
  const err = await runUpdateCheck(UPDATE_CHECK_BEFORE_16_09, [IDENTITY, { status: 503, etag: '"edge-error-page"' }]);
  assert.equal(err.reloads, 1, 'and an error page counted as a new version too');
});
test('update check: one build in two ETag spellings is never a new version', async () => {
  const r = await runUpdateCheck(currentUpdateCheck(), flapping(80));
  assert.equal(r.banners, 0); assert.equal(r.reloads, 0);
});
test('update check: error responses and network failures are ignored', async () => {
  const noise = [IDENTITY];
  for (let i = 0; i < 30; i++) noise.push(i % 3 === 0 ? { status: 503, etag: '"edge-error"' } : i % 3 === 1 ? { fail: true } : { status: 404, etag: '"not-found"' });
  const r = await runUpdateCheck(currentUpdateCheck(), noise);
  assert.equal(r.banners, 0); assert.equal(r.reloads, 0);
});
test('update check: a single odd answer between two good ones does not reload', async () => {
  const r = await runUpdateCheck(currentUpdateCheck(), [IDENTITY, IDENTITY, { etag: '"deploy-xyz"' }, IDENTITY, IDENTITY, GZIP, IDENTITY]);
  assert.equal(r.banners, 0); assert.equal(r.reloads, 0);
});
test('update check: a real deploy is still picked up, and an idle page reloads', async () => {
  const NEW = { etag: '"6aab0000-77000"' }, NEW_GZ = { etag: 'W/"6aab0000-77000"' };
  const r = await runUpdateCheck(currentUpdateCheck(), [IDENTITY, GZIP, NEW, NEW_GZ, NEW]);
  assert.equal(r.banners, 1, 'seen on two consecutive polls (in either spelling) = a real deploy');
  assert.equal(r.reloads, 1, 'no session open, so the page updates itself');
});
test('update check: a real deploy never pulls an examiner out of an open session', async () => {
  const NEW = { etag: '"6aab0000-77000"' };
  const r = await runUpdateCheck(currentUpdateCheck(), [IDENTITY, NEW, NEW, NEW], { session: 'LIVE0001' });
  assert.equal(r.banners, 1, 'the examiner is told');
  assert.equal(r.reloads, 0, 'but the dashboard is not reloaded under them');
  assert.match(r.banner, /אחרי הבחינה/, 'and the banner no longer claims it is updating automatically');
});

// The same defect lived in all four pages; one push re-stamps every file's ETag
// on GitHub Pages, so every page must be safe before that push goes out.
for (const [name, globals] of [
  ['teacher.html', {}],
  ['student.html', {}],
  ['examinee.html', { examInProgress: false, examSubmitted: false, examineeData: null }],
]) {
  const pageSrc = fs.readFileSync(path.join(app, name), 'utf8').replace(/\r/g, '');
  const a = pageSrc.indexOf('===== Auto-update:');
  const f = pageSrc.indexOf('(function() {', a);
  const t = pageSrc.indexOf('// check every 2 min', f);
  const blk = pageSrc.slice(f, pageSrc.indexOf('})();', t) + 5);
  test(name + ': update check ignores ETag spelling flips and error pages', async () => {
    assert.ok(a > 0 && f > a && t > f, 'block located');
    const flap = await runUpdateCheck(blk, flapping(80), { globals });
    assert.equal(flap.reloads, 0, 'one build in two spellings is not a new version');
    const noise = [IDENTITY];
    for (let i = 0; i < 30; i++) noise.push(i % 2 ? { status: 503, etag: '"edge-error"' } : { fail: true });
    assert.equal((await runUpdateCheck(blk, noise, { globals })).reloads, 0, 'errors are not versions');
  });
  test(name + ': update check still reloads an idle page on a real deploy', async () => {
    const NEW = { etag: '"6aab0000-77000"' };
    const r = await runUpdateCheck(blk, [IDENTITY, NEW, NEW, NEW, NEW], { globals });
    assert.equal(r.reloads, 1, 'a confirmed deploy reaches an idle page exactly once');
  });
}
test('examinee.html: a confirmed deploy never reloads a registered examinee', async () => {
  const pageSrc = fs.readFileSync(path.join(app, 'examinee.html'), 'utf8').replace(/\r/g, '');
  const a = pageSrc.indexOf('===== Auto-update:'), f = pageSrc.indexOf('(function() {', a);
  const blk = pageSrc.slice(f, pageSrc.indexOf('})();', pageSrc.indexOf('// check every 2 min', f)) + 5);
  const NEW = { etag: '"6aab0000-77000"' };
  for (const globals of [{ examInProgress: true, examSubmitted: false, examineeData: null },
    { examInProgress: false, examSubmitted: false, examineeData: { idNumber: '123456789' } }]) {
    assert.equal((await runUpdateCheck(blk, [IDENTITY, NEW, NEW, NEW], { globals })).reloads, 0);
  }
});

// ===== 16/09/2026: the dashboard poll fired the full combined report =====
// updateCompletedList() runs on every examinerDashboard poll (5s, 2s while a
// result syncs) and called checkSiteCombinedAvailability() each time — the FULL
// siteCombinedReport, since 2026-06-01, from the first completed result to the
// end of the session. Nobody opened a combined report on 16/09, yet 'אבחון'
// recorded two at 18s and 66s. Verbatim pre-fix probe kept to prove the harness
// reproduces the storm before trusting it on the fix.
const SITE_PROBE_BEFORE_16_09 = String.raw`  function checkSiteCombinedAvailability() {
    var btn = document.getElementById('siteCombinedBtn');
    if (!btn) return;
    if (!sessionCode || !examinerData || !examinerData.id) { btn.style.display = 'none'; return; }
    apiGet({
      action: 'siteCombinedReport',
      examinerId: examinerData.id,
      token: examinerToken,
      sessionCode: sessionCode
    }).then(function(data) {
      if (!data || data.status !== 'ok' || !Array.isArray(data.sessions) || data.sessions.length <= 1) {
        btn.style.display = 'none';
        return;
      }
      // Cache the payload — generateSiteCombinedReport reuses it to avoid
      // a second round-trip when the user clicks.
      window._siteCombinedCache = data;
      var others = data.sessions.length - 1;
      btn.innerHTML = '📊 דו"ח משותף לאתר (' + data.sessions.length + ' סשנים, +' + others + ')';
      btn.style.display = '';
    }).catch(function() { btn.style.display = 'none'; });
  }
`;
function siteProbePage(code, { reply } = {}) {
  let now = 1_000_000;
  const calls = [];
  const btn = { style: { display: 'none' }, innerHTML: '' };
  const opened = [];
  const ctx = {
    console: quiet, Promise, Array, JSON,
    Date: class extends Date { static now() { return now; } },
    sessionCode: 'SESS0001', examinerData: { id: '111' }, examinerToken: 'tok',
    document: { getElementById: id => (id === 'siteCombinedBtn' ? btn : null) },
    escHtml: v => String(v), buildSiteCombinedReportHtml: d => '<report sessions=' + (d.sessions || []).length + '>',
    alert() {},
    apiGet: (params, timeoutMs) => { calls.push({ params, timeoutMs, at: now }); return reply ? reply(params) : new Promise(() => {}); },
  };
  ctx.window = ctx;
  ctx.window.open = () => { const doc = { html: '', open() { this.html = ''; }, write(h) { this.html += h; }, close() {} }; opened.push(doc); return { document: doc }; };
  vm.createContext(ctx); vm.runInContext(code, ctx);
  return { ctx, calls, btn, opened, advance: ms => { now += ms; }, setSession: s => { ctx.sessionCode = s; } };
}
const currentSiteProbeCode = () => {
  const src = examiner.replace(/\r/g, '');
  const a = src.indexOf('  var SITE_COMBINED_PROBE_MIN_MS');
  const e = src.indexOf('  // Builds the HTML for the combined report.', a);
  assert.ok(a > 0 && e > a, 'probe + click handler located');
  return src.slice(a, e);
};
const settle = async () => { for (let i = 0; i < 10; i++) await Promise.resolve(); };
const twoSessions = () => Promise.resolve({ status: 'ok', sessions: [{ code: 'SESS0001' }, { code: 'OTHER' }], results: [] });

test('site probe: the harness reproduces the 16/09 storm on the pre-fix code', async () => {
  const page = siteProbePage(SITE_PROBE_BEFORE_16_09, { reply: twoSessions });
  for (let poll = 0; poll < 60; poll++) { page.ctx.checkSiteCombinedAvailability(); await settle(); page.advance(5000); }
  assert.equal(page.calls.length, 60, 'five minutes of 5-second polls = sixty full combined reports');
});
test('site probe: five minutes of polling sends ONE report, not sixty', async () => {
  const page = siteProbePage(currentSiteProbeCode(), { reply: twoSessions });
  for (let poll = 0; poll < 60; poll++) { page.ctx.checkSiteCombinedAvailability(); await settle(); page.advance(5000); }
  assert.equal(page.calls.length, 1);
  assert.equal(page.btn.style.display, '', 'and the button is still shown');
});
test('site probe: never a second request while one is still out', async () => {
  const page = siteProbePage(currentSiteProbeCode());        // reply never resolves
  for (let poll = 0; poll < 200; poll++) { page.ctx.checkSiteCombinedAvailability(); page.advance(5000); }
  assert.equal(page.calls.length, 1, '1000 seconds of polls against a hung request still = 1');
});
test('site probe: re-asks after 10 minutes, and at once for a new session', async () => {
  const page = siteProbePage(currentSiteProbeCode(), { reply: twoSessions });
  page.ctx.checkSiteCombinedAvailability(); await settle();
  page.advance(9 * 60 * 1000); page.ctx.checkSiteCombinedAvailability(); await settle();
  assert.equal(page.calls.length, 1, 'not before 10 minutes');
  page.advance(61 * 1000); page.ctx.checkSiteCombinedAvailability(); await settle();
  assert.equal(page.calls.length, 2, 'again after 10 minutes');
  page.setSession('SESS0002'); page.ctx.checkSiteCombinedAvailability(); await settle();
  assert.equal(page.calls.length, 3, 'a different session is asked immediately');
  assert.equal(page.calls[2].params.sessionCode, 'SESS0002');
});
test('site probe: a transient failure does not hide a button that was showing', async () => {
  let fail = false;
  const page = siteProbePage(currentSiteProbeCode(), { reply: () => (fail ? Promise.reject(new Error('net')) : twoSessions()) });
  page.ctx.checkSiteCombinedAvailability(); await settle();
  assert.equal(page.btn.style.display, '');
  fail = true; page.advance(11 * 60 * 1000); page.ctx.checkSiteCombinedAvailability(); await settle();
  assert.equal(page.btn.style.display, '', 'still shown');
});
test('combined report click: opens the window first, then always fetches fresh with a 90s deadline', async () => {
  let resolve;
  const page = siteProbePage(currentSiteProbeCode(), { reply: p => new Promise(r => { resolve = r; }) });
  page.ctx.window.generateSiteCombinedReport();
  assert.equal(page.opened.length, 1, 'window opened synchronously inside the click (popup blockers)');
  assert.match(page.opened[0].html, /טוען/, 'showing a loading message');
  assert.equal(page.calls.length, 1); assert.equal(page.calls[0].timeoutMs, 90000);
  resolve({ status: 'ok', sessions: [{ code: 'A' }, { code: 'B' }, { code: 'C' }], results: [] }); await settle();
  assert.match(page.opened[0].html, /<report sessions=3>/, 'the fresh report replaced the loading message');
  assert.equal(page.ctx._siteCombinedCache, undefined, 'no stale payload is kept for next time');
});

test('all inline client scripts and both service workers parse', () => {
  for (const [name, html] of [['examiner.html', examiner], ['examinee.html', examinee]]) {
    let count = 0;
    for (const script of html.matchAll(/<script\b[^>]*>([\s\S]*?)<\/script>/gi)) {
      new vm.Script(script[1], { filename: name + ':script' + (++count) });
    }
    assert.ok(count > 0);
  }
  for (const name of ['sw-examiner.js', 'sw-examinee.js']) new vm.Script(fs.readFileSync(path.join(app, name), 'utf8'), { filename: name });
});

// ===== 2026-09-18 review, action 1: poll pacing — jitter, gradual recovery, 60s poll deadline =====
// A fleet released by one server stall used to re-arrive as a single burst
// (no randomness anywhere in the poll loops, and one fast answer snapped every
// device back to its base interval at the same moment). Polls also abandoned
// the server at 30s and asked again while the abandoned execution kept running.
for (const [name, src] of [['examiner', examiner], ['examinee', examinee]]) {
  test(name + ': poll jitter stays within ±30% and varies; backoff steps up to the cap and down to the base', () => {
    const { ctx } = context(); load(ctx, pollHelpers(src));
    const seen = new Set();
    for (let i = 0; i < 300; i++) {
      const v = ctx.jitterMs(10000);
      assert.ok(Number.isInteger(v) && v >= 7000 && v <= 13000, 'jitter within ±30%: ' + v);
      seen.add(v);
    }
    assert.ok(seen.size > 10, 'jitter is random, not a constant');
    assert.equal(ctx.POLL_TIMEOUT_MS, 60000);
    let d = 5000; const up = [], down = [];
    for (let i = 0; i < 5; i++) { d = ctx.nextPollDelay(d, 5000, 20000, true); up.push(d); }
    assert.deepEqual(up, [7500, 11250, 16875, 20000, 20000], 'backoff ×1.5 capped at the max');
    for (let i = 0; i < 5; i++) { d = ctx.nextPollDelay(d, 5000, 20000, false); down.push(d); }
    assert.deepEqual(down, [13333, 8889, 5926, 5000, 5000], 'recovery ÷1.5 floored at the base');
  });
}
test('dashboard forwards the 60s poll deadline, jitters every wait and comes down from a backoff one step at a time', async () => {
  let deadline = null, gate = deferred(); const jittered = [];
  const { ctx, timer } = dashboardContext((params, timeoutMs) => { deadline = timeoutMs; return gate.promise; });
  ctx.jitterMs = ms => { jittered.push(ms); return ms; };
  const ok = { status: 'ok', pending: [], active: [], completed: [] };
  ctx.startDashboardPolling(); await drain();
  assert.equal(deadline, 60000, 'polls get the 60s deadline, not the 30s action default');
  await timer.advance(7000); gate.resolve(ok); await drain(); assert.equal(ctx.dashPollDelayMs, 7500, 'a 7s answer is slow: ×1.5');
  gate = deferred(); await timer.advance(7500); await timer.advance(7000); gate.resolve(ok); await drain(); assert.equal(ctx.dashPollDelayMs, 11250);
  gate = deferred(); await timer.advance(11250); gate.resolve(ok); await drain(); assert.equal(ctx.dashPollDelayMs, 7500, 'a fast answer steps down; it no longer snaps to 5s');
  gate = deferred(); await timer.advance(7500); gate.resolve(ok); await drain(); assert.equal(ctx.dashPollDelayMs, 5000);
  assert.deepEqual(jittered, [7500, 11250, 7500, 5000], 'every scheduled wait went through the jitter');
  ctx.stopDashboardPolling(); assert.equal(timer.jobs.size, 0);
});
test('approval poll forwards the 60s deadline, jitters every wait and steps back down after slow answers', async () => {
  let deadline = null, gate = deferred(), calls = 0; const jittered = [];
  const { ctx, timer } = approvalContext((params, timeoutMs) => { calls++; deadline = timeoutMs; return gate.promise; });
  ctx.jitterMs = ms => { jittered.push(ms); return ms; };
  const waiting = { status: 'ok', approval: 'waiting' };
  ctx.startApprovalPolling(); await drain();
  assert.equal(calls, 1); assert.equal(deadline, 60000);
  await timer.advance(7000); gate.resolve(waiting); await drain(); assert.equal(ctx.approvalPollDelayMs, 7500);
  gate = deferred(); await timer.advance(7500); assert.equal(calls, 2); await timer.advance(7000); gate.resolve(waiting); await drain(); assert.equal(ctx.approvalPollDelayMs, 11250);
  gate = deferred(); await timer.advance(11250); assert.equal(calls, 3); gate.resolve(waiting); await drain(); assert.equal(ctx.approvalPollDelayMs, 7500);
  gate = deferred(); await timer.advance(7500); assert.equal(calls, 4); gate.resolve(waiting); await drain(); assert.equal(ctx.approvalPollDelayMs, 5000);
  assert.deepEqual(jittered, [7500, 11250, 7500, 5000]);
  timer.clear(ctx.approvalInterval); ctx.approvalInterval = null; assert.equal(timer.jobs.size, 0);
});
test('DQ polling forwards the 60s deadline and never retries before the wait the server asked for, even at the shortest jitter', async () => {
  let pending = deferred(), calls = 0, deadline = null;
  const { ctx, timer } = dqContext((params, timeoutMs) => { calls++; deadline = timeoutMs; return pending.promise; });
  ctx.jitterMs = ms => Math.round(ms * 0.7);   // the shortest draw the page can make
  ctx.startDQOverturnPolling(); await timer.advance(3000); assert.equal(calls, 1); assert.equal(deadline, 60000);
  pending.resolve({ status: 'error', waitSec: 6 }); await drain(); pending = deferred();
  await timer.advance(5999); assert.equal(calls, 1, 'server-requested wait wins over the jittered backoff');
  await timer.advance(1); assert.equal(calls, 2);
  pending.resolve({ status: 'ok', approval: 'dq_confirmed' }); await drain();
  assert.equal(timer.jobs.size, 0); assert.equal(ctx.dqOverturnInterval, null);
});
function examStatusContext(apiGet) {
  const setup = context({ apiGet, sessionCode: 'TEST00', examineeData: { idNumber: 'SYNTHETIC' }, examineeToken: 'synthetic',
    examInProgress: true, examSubmitted: false, tabSwitchDQConfirmed: false, applyExtraMinutes() {}, showDQScreen() {},
    localStorage: { setItem() {} } });
  load(setup.ctx, section(examinee, '  var EXAMSTATUS_POLL_BASE_MS', '  var resizeDebounceTimer'));
  withPacing(setup.ctx, examinee);
  return setup;
}
test('exam-status poll forwards the 60s deadline, backs off ×1.5 on failure and steps back down on a healthy answer', async () => {
  let deadline = null, gate = deferred(), calls = 0; const jittered = [];
  const { ctx, timer } = examStatusContext((params, timeoutMs) => { calls++; deadline = timeoutMs; return gate.promise; });
  ctx.jitterMs = ms => { jittered.push(ms); return ms; };
  ctx.startExamStatusPoll(); await timer.advance(10000); assert.equal(calls, 1); assert.equal(deadline, 60000);
  gate.resolve({ status: 'error' }); await drain(); assert.equal(ctx.examStatusDelayMs, 15000);
  gate = deferred(); await timer.advance(15000); assert.equal(calls, 2);
  gate.resolve({ status: 'ok', examStatus: 'in_exam', extraMinutes: 0 }); await drain(); assert.equal(ctx.examStatusDelayMs, 10000);
  assert.deepEqual(jittered, [15000, 10000]);
  ctx.stopExamStatusPoll(); assert.equal(timer.jobs.size, 0);
});

// ===== 2026-09-18 review, action 2: a non-JSON / non-200 answer is a degraded backend =====
// Google's own HTML error page, or a 404/5xx from its front door, must slow the
// whole fleet to 30-60 s and must never be read as "reload", "new version" or
// "session expired". A remembered login is discarded only on the server's own
// tokenExpired verdict — on 16/09 every transient failure at boot sent the
// examiner to the password screen.
for (const [name, src] of [['examiner', examiner], ['examinee', examinee]]) {
  test(name + ': non-JSON and non-200 answers mark the backend degraded, a JSON answer clears it, timeouts and network errors leave it alone', async () => {
    let answer;
    const { ctx, timer } = context({ fetch: () => answer() });
    load(ctx, helper(src));
    const call = () => ctx.fetchJsonWithTimeout('synthetic', {}, 100).then(v => ({ v }), e => ({ e }));
    const page = body => () => Promise.resolve({ ok: true, status: 200, text: () => Promise.resolve(body) });
    answer = page('<html>Google error</html>');
    let r = await call(); assert.equal(r.e.name, 'SyntaxError'); assert.equal(r.e.transport, 'nonjson'); assert.equal(ctx.isBackendDegraded(), true);
    answer = page('{"status":"ok"}');
    r = await call(); assert.equal(r.v.status, 'ok'); assert.equal(ctx.isBackendDegraded(), false, 'a JSON answer clears the state');
    answer = () => Promise.resolve({ ok: false, status: 503, text: () => Promise.resolve('') });
    r = await call(); assert.equal(r.e.message, 'HTTP 503'); assert.equal(r.e.transport, 'http'); assert.equal(r.e.status, 503); assert.equal(ctx.isBackendDegraded(), true);
    answer = page('{"status":"ok"}'); await call(); assert.equal(ctx.isBackendDegraded(), false);
    answer = () => Promise.reject(new TypeError('Failed to fetch'));
    r = await call(); assert.equal(r.e.name, 'TypeError'); assert.equal(ctx.isBackendDegraded(), false, 'a network error says nothing about the backend');
    answer = () => new Promise(() => {});
    const pending = call(); await drain(); await timer.advance(100); r = await pending;
    assert.equal(r.e.name, 'TimeoutError'); assert.equal(ctx.isBackendDegraded(), false, 'a timeout says nothing about the backend');
  });
}
test('dashboard slows to 30-60s on a degraded backend, tells the examiner not to reload, and steps down once JSON is back', async () => {
  let body = '<html>Google error</html>';
  const { ctx, timer, nodes } = dashboardContext();
  ctx.fetch = () => Promise.resolve({ ok: true, status: 200, text: () => Promise.resolve(body) });
  ctx.apiGet = (params, timeoutMs) => ctx.fetchJsonWithTimeout('synthetic', {}, timeoutMs);
  const banner = nodes.get('offlineBanner');
  ctx.startDashboardPolling(); await drain();
  assert.equal(ctx.failedPolls, 1); assert.equal(ctx.dashPollDelayMs, 30000, 'first degraded answer: straight to the 30s floor');
  assert.equal(banner.classList.values.has('show'), true, 'shown at once, not after three failures');
  assert.ok(banner.textContent.includes('אין צורך לרענן'), 'the banner says not to reload');
  await timer.advance(30000); assert.equal(ctx.dashPollDelayMs, 45000);
  await timer.advance(45000); assert.equal(ctx.dashPollDelayMs, 60000, 'capped at 60s');
  body = '{"status":"ok","pending":[],"active":[],"completed":[]}';
  await timer.advance(60000);
  assert.equal(ctx.failedPolls, 0); assert.equal(banner.classList.values.has('show'), false);
  assert.equal(ctx.dashPollDelayMs, 40000, 'recovery steps down from 60s; it does not snap to 5s');
  ctx.stopDashboardPolling(); assert.equal(timer.jobs.size, 0);
});
test('dashboard still needs three plain failures before the offline banner, with the offline text', async () => {
  const { ctx, nodes } = dashboardContext(() => Promise.reject(new TypeError('Failed to fetch')));
  for (let i = 0; i < 2; i++) await ctx.pollDashboard();
  assert.equal(nodes.get('offlineBanner').classList.values.has('show'), false);
  await ctx.pollDashboard();
  assert.equal(nodes.get('offlineBanner').classList.values.has('show'), true);
  assert.ok(nodes.get('offlineBanner').textContent.includes('אין חיבור לשרת'));
});
test('approval poll slows to the degraded cadence and tells the examinee not to reload', async () => {
  const { ctx, timer, nodes } = approvalContext();
  ctx.fetch = () => Promise.resolve({ ok: false, status: 503, text: () => Promise.resolve('') });
  ctx.apiGet = (params, timeoutMs) => ctx.fetchJsonWithTimeout('synthetic', {}, timeoutMs);
  ctx.startApprovalPolling(); await drain();
  assert.equal(ctx.approvalFailCount, 1); assert.equal(ctx.approvalPollDelayMs, 30000);
  await timer.advance(30000); await timer.advance(45000);
  assert.equal(ctx.approvalFailCount, 3); assert.equal(ctx.approvalPollDelayMs, 60000);
  assert.ok(nodes.get('approvalError').textContent.includes('אין צורך לרענן'));
  timer.clear(ctx.approvalInterval); ctx.approvalInterval = null; assert.equal(timer.jobs.size, 0);
});
function rememberedContext(fetchImpl) {
  const ui = dom(); const removed = []; let entered = 0;
  for (const id of ['loginError', 'rememberMe']) ui.element(id);
  const setup = context({ ...ui, fetch: fetchImpl, API_URL: 'https://synthetic.invalid/exec', API_ORIGIN: 'examiner-app',
    localStorage: { removeItem: key => removed.push(key) }, enterAsRemembered() { entered++; } });
  load(setup.ctx, helper(examiner));
  load(setup.ctx, section(examiner, '  // ===== Remembered login', '  // ===== end remembered login'));
  return { ...setup, ...ui, removed, entered: () => entered };
}
test('a remembered login survives transient failures with a retry in place, and only the server\'s tokenExpired verdict discards it', async () => {
  let ok = true, body = '<html>Google error</html>';
  const { ctx, nodes, removed, entered } = rememberedContext(() => Promise.resolve({ ok, status: ok ? 200 : 503, text: () => Promise.resolve(body) }));
  const creds = { id: '123456789', token: 'synthetic' };
  assert.equal(await ctx.verifyRememberedLogin(creds), 'retry');
  assert.deepEqual(removed, [], "Google's HTML page does not delete the saved login");
  assert.equal(nodes.get('loginError').style.display, 'block');
  assert.ok(nodes.get('rememberedLoginMsg').textContent.includes('לא הצלחנו לאמת'));
  // the retry button re-verifies; a JSON error WITHOUT tokenExpired still keeps the login and shows the message
  body = '{"status":"error","message":"החשבון אינו פעיל"}';
  nodes.get('rememberedLoginRetry').click(); await drain();
  assert.deepEqual(removed, []); assert.ok(nodes.get('rememberedLoginMsg').textContent.includes('החשבון אינו פעיל'));
  // a 503 from Google's front door: same — keep it
  ok = false; body = ''; nodes.get('rememberedLoginRetry').click(); await drain();
  assert.deepEqual(removed, []);
  // the server's verdict discards it
  ok = true; body = '{"status":"error","message":"פג תוקף ההתחברות","tokenExpired":true}';
  assert.equal(await ctx.verifyRememberedLogin(creds), 'expired');
  assert.deepEqual(removed, ['ext_examiner_remember']);
  // and a healthy answer enters without touching the saved login
  body = '{"status":"ok","examiner":{"name":"synthetic","id":"123456789","role":"בוחן"}}';
  assert.equal(await ctx.verifyRememberedLogin(creds), 'ok');
  assert.equal(entered(), 1); assert.deepEqual(removed, ['ext_examiner_remember']);
});
