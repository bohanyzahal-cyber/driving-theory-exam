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
