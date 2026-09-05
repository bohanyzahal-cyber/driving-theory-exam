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
  const document = { activeElement: null, getElementById: id => nodes.get(id) || null };
  function element(id) {
    const el = { id, textContent: '', disabled: false, style: {}, handlers: {}, attrs: {},
      classList: { values: new Set(), add(v) { this.values.add(v); }, remove(v) { this.values.delete(v); } },
      addEventListener(type, cb) { this.handlers[type] = cb; },
      setAttribute(key, value) { this.attrs[key] = value; },
      appendChild(child) { this.children = this.children || []; this.children.push(child); child.parentNode = this; },
      focus() { document.activeElement = this; },
      click() { if (!this.disabled && this.handlers.click) this.handlers.click(); }
    };
    let html = '';
    Object.defineProperty(el, 'innerHTML', { get: () => html, set(value) {
      html = value;
      for (const match of value.matchAll(/id="([^"]+)"/g)) element(match[1]);
    } });
    nodes.set(id, el); return el;
  }
  document.createElement = () => element('');
  document.body = { appendChild(el) { nodes.set(el.id, el); el.parentNode = this; }, removeChild(el) { nodes.delete(el.id); } };
  for (const id of ['examArea', 'offlineBanner', 'approvalError']) element(id);
  return { document, nodes };
}
function context(extra = {}, timer = new Timers()) {
  const clockDate = class extends Date { static now() { return timer.now; } };
  const ctx = { console: quiet, Date: clockDate, setTimeout: timer.set, clearTimeout: timer.clear,
    clearInterval: timer.clear, AbortController, ...extra };
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
    approvalInterval: null, updateApprovalDebug() {}, setupWaitingProtections() {}, teardownWaitingProtections() {},
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
test('finishing a later attempt archives the unconfirmed earlier result and confirmed cleanup is attempt-specific', () => {
  const { ctx, stored, payload } = submitContext();
  const newer = resultPayload('new-token');
  ctx.persistPendingResult(newer, [{ questionId: 'synthetic-question' }]);
  const results = [...stored.entries()].filter(([key]) => key.startsWith('pendingResult_'));
  assert.equal(results.length, 2);
  assert.ok(results.some(([, raw]) => JSON.parse(raw).examineeToken === 'synthetic'));
  assert.ok(results.some(([, raw]) => JSON.parse(raw).examineeToken === 'new-token'));
  ctx.clearConfirmedPendingResult(payload);
  assert.equal([...stored.keys()].filter(key => key.startsWith('pendingResult_')).length, 1);
  assert.equal(JSON.parse(stored.get('pendingResult_SYNTHETIC')).examineeToken, 'new-token');
  assert.equal(JSON.parse(stored.get('pendingResult_SYNTHETIC')).origin, 'examinee-app');
  assert.equal(JSON.parse(stored.get('pendingWrongAnswers_SYNTHETIC'))[0].questionId, 'synthetic-question');
});
test('a rejected previous result is visible even when another examinee is using the device', async () => {
  const { ctx, nodes, stored, payload, statuses } = submitContext(() => Promise.resolve({ status: 'error', examineeTokenError: 'mismatch' }));
  ctx.examineeData.idNumber = 'NEXT'; ctx.examineeToken = 'next-token';
  ctx.submitWithRetry(payload, 3, []); await drain();
  assert.ok(nodes.get('pendingRecoveryNotice')); assert.equal(stored.size, 2); assert.deepEqual(statuses, []);
});
test('old submit response cannot erase or falsely confirm a newer result sharing the same ID', async () => {
  const pending = deferred(); let calls = 0;
  const { ctx, timer, stored, statuses, payload } = submitContext(() => { calls++; return calls === 1 ? pending.promise : Promise.resolve({ status: 'ok' }); });
  ctx.submitWithRetry(payload, 3, []);
  const newer = resultPayload('new-token'); ctx.examineeToken = 'new-token'; stored.set('pendingResult_SYNTHETIC', JSON.stringify(newer));
  ctx.submitWithRetry(newer, 3, []); assert.equal(calls, 1);
  pending.resolve({ status: 'ok' }); await drain();
  assert.equal(JSON.parse(stored.get('pendingResult_SYNTHETIC')).examineeToken, 'new-token'); assert.deepEqual(statuses, []);
  await timer.advance(3000); assert.equal(calls, 2); assert.equal(stored.size, 0); assert.deepEqual(statuses, ['received']);
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
