// teacher.html, student.html and exam.html (standalone), executed for real.
//
// Same harness as tests/client_examiner.test.cjs: sections cut out of the page,
// run in a vm context against a fake DOM, a fake clock and a fake bank. No
// network, no browser, no production data.
//
// What it gates (review ids from docs/reviews/2026-09-21/):
//   D5  teacher: a stored login is dropped ONLY on tokenExpired === true
//   D6  teacher: every request has a deadline; no raw fetch survives
//   D4  teacher: the page reloads itself 60 s after the banner, but the modal
//       guard is evaluated AT FIRE TIME and every 15 s after it (21/09 msg 20)
//   D16 student: progress and studentId belong to class+name, not to the device
//   S8  student: one pass rule, ceil(total * 0.86), for every practice mode
//   r31 every action these three pages send is a REPORTS action (practice,
//       progress, classes, the commander dashboard), so all three talk to the
//       second Apps Script deployment - DESIGN §13.3
//   plus: startPractice + the signed grant it returns, a purely local language
//         switch, a missing grant named instead of a blank practice, spaced
//         repetition via mode=ids, the idle-only student reload, and the
//         standalone exam.html flow.
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const app = path.resolve(__dirname, '..');
const teacher = fs.readFileSync(path.join(app, 'teacher.html'), 'utf8');
const student = fs.readFileSync(path.join(app, 'student.html'), 'utf8');
const examPage = fs.readFileSync(path.join(app, 'exam.html'), 'utf8');
const transportSrc = fs.readFileSync(path.join(app, 'shared', 'transport.js'), 'utf8');
const quiet = { log() {}, warn() {}, error() {} };

function section(src, start, end) {
  const i = src.indexOf(start), j = src.indexOf(end, i + start.length);
  assert.ok(i >= 0 && j > i, 'source section found: ' + start);
  return src.slice(i, j);
}
async function drain() { for (let i = 0; i < 40; i++) await Promise.resolve(); }

const EPOCH = 1758400000000;
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
}

function dom(screens = []) {
  const nodes = new Map();
  const document = {
    activeElement: null, visibilityState: 'visible',
    getElementById: id => nodes.get(id) || null,
    querySelector: () => null, querySelectorAll: () => [],
    addEventListener() {}, removeEventListener() {},
    documentElement: { dir: 'rtl', lang: 'he' },
    head: { appendChild() {} }
  };
  function element(id, tag) {
    const el = {
      id, tagName: (tag || 'div').toUpperCase(), textContent: '', value: '', disabled: false, type: '',
      style: { cssText: '', display: '' }, handlers: {}, children: [], parentNode: null,
      classList: {
        values: new Set(),
        add(v) { this.values.add(v); }, remove(v) { this.values.delete(v); }, contains(v) { return this.values.has(v); }
      },
      addEventListener(t, cb) { (this.handlers[t] = this.handlers[t] || []).push(cb); },
      setAttribute() {}, getAttribute: () => null,
      appendChild(c) { this.children.push(c); c.parentNode = this; if (c.id) nodes.set(c.id, c); return c; },
      removeChild(c) { this.children = this.children.filter(x => x !== c); if (c.id) nodes.delete(c.id); c.parentNode = null; },
      insertBefore(c) { this.children.unshift(c); c.parentNode = this; return c; },
      querySelector: () => null, querySelectorAll: () => [],
      closest: () => null, focus() {}, click() { (this.handlers.click || []).forEach(cb => cb()); }
    };
    let html = '';
    Object.defineProperty(el, 'innerHTML', { get: () => html, set(v) { html = v; el.children = []; } });
    if (id) nodes.set(id, el);
    return el;
  }
  document.createElement = tag => element('', tag);
  document.body = {
    children: [],
    appendChild(el) { this.children.push(el); el.parentNode = this; if (el.id) nodes.set(el.id, el); return el; },
    removeChild(el) { this.children = this.children.filter(c => c !== el); if (el.id) nodes.delete(el.id); el.parentNode = null; }
  };
  screens.forEach(id => element(id));
  return { document, nodes, element };
}

function baseContext(extra = {}, timer = new Timers()) {
  const clockDate = class extends Date { static now() { return timer.now; } };
  const store = new Map();
  const localStorage = {
    getItem: k => (store.has(k) ? store.get(k) : null),
    setItem: (k, v) => store.set(k, String(v)),
    removeItem: k => store.delete(k),
    get length() { return store.size; }
  };
  const ctx = {
    console: quiet, Date: clockDate,
    setTimeout: timer.set, clearTimeout: timer.clear, setInterval: timer.setInterval, clearInterval: timer.clear,
    AbortController, Promise, JSON, Math, Object, Array, String, Number, RegExp, isNaN, parseInt, parseFloat,
    encodeURIComponent, decodeURIComponent,
    localStorage, sessionStorage: { getItem: () => null, setItem() {}, removeItem() {} },
    ...extra
  };
  ctx.window = ctx; ctx.globalThis = ctx;
  vm.createContext(ctx);
  vm.runInContext(transportSrc, ctx);
  ctx.ExamTransport._resetHealth();
  return { ctx, timer, store, localStorage };
}
const load = (ctx, code) => vm.runInContext(code, ctx);

// A stand-in for shared/bank.js with a handful of questions. `he` carries them
// all; `en` is missing 907 on purpose (the real ru bank is missing exactly that
// id) so the Hebrew fallback is exercised.
//
// The three loaders mirror the real ones: loadGrant is the practice/exam path
// (one request, the granted ids in EVERY language, which is what makes a
// language switch local), loadIds is the examiner's, loadFull is find_image's.
// `load` is gone - there is no public bank to fetch by language any more.
function fakeBank() {
  const data = {
    he: {
      1: { id: 1, t: 'שאלה 1', a: ['א', 'ב', 'ג', 'ד'], i: 'TQ_PIC_1.jpg' },
      2: { id: 2, t: 'שאלה 2', a: ['א', 'ב', 'ג', 'ד'], i: '' },
      907: { id: 907, t: 'שאלה 907', a: ['א', 'ב', 'ג', 'ד'], i: '' }
    },
    en: {
      1: { id: 1, t: 'question 1', a: ['A', 'B', 'C', 'D'], i: 'TQ_PIC_1.jpg' },
      2: { id: 2, t: 'question 2', a: ['A', 'B', 'C', 'D'], i: '' }
    }
  };
  const loaded = new Set();
  const grants = [], idCalls = [], fullCalls = [];
  const result = ids => Promise.resolve({ build: 'test', count: ids, missing: [] });
  return {
    grants, idCalls, fullCalls,
    api: {
      LANGS: ['he', 'ru', 'en', 'ar', 'fr', 'es', 'am'],
      has: l => loaded.has(l),
      loadGrant(bank) {
        grants.push(bank);
        Object.keys(data).forEach(l => loaded.add(l));   // every language of the granted ids
        return result(3);
      },
      loadIds(bank, ids, langs) {
        idCalls.push({ bank, ids, langs });
        (langs || ['he']).forEach(l => loaded.add(l));
        return result(ids.length);
      },
      loadFull(bank, lang) { fullCalls.push({ bank, lang }); loaded.add(lang); return result(3); },
      get(id, lang) {
        if (!loaded.has(lang)) return null;
        const e = (data[lang] || {})[id];
        return e ? { id: e.id, text: e.t, answers: e.a, image: e.i || '' } : null;
      },
      imageUrl: e => (e && (e.image || e.i) ? 'images/' + (e.image || e.i) : ''),
      search: () => []
    }
  };
}
// What startPractice now answers with, next to the questions.
const PRACTICE_BANK = { url: 'https://gateway.example', grant: 'payload.sig', exp: 1758400000000 + 2 * 3600 * 1000 };

// r31 (DESIGN §13.3): the two Apps Script deployments. They are one url in
// production until Yossi creates the second one; the contexts below drive them
// apart so that "which deployment answered this" is observable in a test.
const EXAM_URL = 'https://exam.example/exec';
const REPORTS_URL = 'https://reports.example/exec';

// ======================================================================
// teacher.html
// ======================================================================
function teacherContext(answer) {
  const ui = dom();
  const { ctx, timer, store } = baseContext({
    ...ui,
    API_URL: EXAM_URL, REPORTS_API_URL: REPORTS_URL, API_ORIGIN: 'teacher-app',
    teacher: null, escHtml: s => String(s == null ? '' : s)
  });
  ctx.seen = [];
  ctx.fetch = (url, opts) => {
    const match = /action=([a-zA-Z]+)/.exec(url);
    const action = match ? match[1] : JSON.parse((opts && opts.body) || '{}').action;
    ctx.seen.push({ action, url, opts });
    return answer(action);
  };
  load(ctx, section(teacher, '// D6: teacher.html had no request deadline', '// ===== Auth ====='));
  return { ctx, timer, store, ...ui };
}

test('D6: a teacher request is bounded, sent no-store, and never reaches r.json() on an HTML page', async () => {
  const { ctx, timer } = teacherContext(() => new Promise(() => {}));
  const p = ctx.apiGet({ action: 'teacherGetClasses' }).catch(e => e.name);
  await drain();
  assert.match(ctx.seen[0].url, /cache=|_t=/, 'cache-busted');
  assert.equal(ctx.seen[0].opts.cache, 'no-store');
  await timer.advance(30000);
  assert.equal(await p, 'TimeoutError', 'the default deadline is 30 s, not "forever"');

  ctx.fetch = () => Promise.resolve({ ok: true, text: () => Promise.resolve('<html>Google error</html>') });
  const err = await ctx.apiGet({ action: 'teacherGetClasses' }).catch(e => e);
  assert.equal(err.transport, 'nonjson', 'an HTML page is classified, not thrown raw at the UI');
  assert.match(ctx.transportErrorText(), /השרת עמוס/, 'and it says "busy", not "reload"');
  assert.match(ctx.transportErrorText(), /אין צורך לרענן/);
});

test('D6: the two heavy reads get 90 s, not the 30 s default', () => {
  assert.match(teacher, /apiGet\(\{ action: 'teacherClassDetails'[^)]*\}, HEAVY_TIMEOUT_MS\)/,
    'teacherClassDetails (109k rows, measured 25 s) must not run on the 30 s default');
  assert.match(teacher, /\}, HEAVY_TIMEOUT_MS\)\.then\(function\(resp\)/,
    'teacherCommanderDashboard too');
  assert.match(teacher, /var HEAVY_TIMEOUT_MS = 90000;/);
});

test('r31: every request teacher.html makes goes to the reports deployment', async () => {
  // Nothing this page asks for runs during an exam - logins, classes, the
  // commander dashboard, the at-risk list, exports - so the exam deployment
  // must not have to carry a byte of it. apiUrl IS the reports url here.
  const { ctx } = teacherContext(() =>
    Promise.resolve({ ok: true, text: () => Promise.resolve('{"status":"ok"}') }));
  for (const action of ['teacherLogin', 'teacherVerifyLogin', 'teacherGetClasses', 'teacherClassDetails',
                        'teacherCommanderDashboard', 'teacherAtRiskList', 'teacherExportData',
                        'teacherCreateClass', 'teacherDeleteClass']) {
    await ctx.apiGet({ action: action });
    assert.equal(ctx.seen[ctx.seen.length - 1].url.indexOf(REPORTS_URL + '?'), 0, action);
  }
  assert.ok(ctx.seen.every(c => c.url.indexOf(EXAM_URL) !== 0), 'not one of them touched the exam deployment');
});

test('r31: teacher.html carries REPORTS_API_URL as a line of its own', () => {
  assert.match(teacher, /\r\nvar REPORTS_API_URL = API_URL;\r\n/);
  assert.match(teacher, /apiUrl: REPORTS_API_URL,/);
});

test('D6: no raw fetch survives in teacher.html', () => {
  const code = section(teacher, '<script>\r\n(function(){', '</script>');
  const offenders = code.split('\r\n')
    .filter(l => /(^|[^a-zA-Z._])fetch\(/.test(l) || /\.then\(function\(r\) \{ return r\.json\(\)/.test(l))
    .filter(l => !/^\s*\/\//.test(l));
  assert.deepEqual(offenders, [], 'every request goes through the shared transport');
});

test('D5: the saved teacher login is deleted ONLY when the server says tokenExpired', async () => {
  for (const [name, answer, shouldDelete] of [
    ['tokenExpired', { status: 'error', tokenExpired: true }, true],
    ['a plain server error', { status: 'error', message: 'שגיאה' }, false],
    ['a rate limit', { status: 'error', code: 'rate_limited' }, false],
    ['an empty answer', {}, false]
  ]) {
    const { ctx, store } = teacherContext(() =>
      Promise.resolve({ ok: true, text: () => Promise.resolve(JSON.stringify(answer)) }));
    load(ctx, section(teacher, 'function tryAutoLogin() {', '// ===== Dashboard ====='));
    ctx.onLoggedIn = () => { ctx.enteredApp = true; };
    store.set('teacher_auth', JSON.stringify({ id: '5', token: 'T', name: 'x' }));
    ctx.sessionStorage = { getItem: () => null, setItem() {}, removeItem() {} };
    ctx.tryAutoLogin();
    await drain();
    assert.equal(store.has('teacher_auth'), !shouldDelete, name);
    assert.notEqual(ctx.enteredApp, true, name + ': and the app is not entered');
  }
});

test('D5: a transport failure keeps the stored login untouched', async () => {
  const { ctx, store } = teacherContext(() => Promise.reject(new Error('network down')));
  load(ctx, section(teacher, 'function tryAutoLogin() {', '// ===== Dashboard ====='));
  ctx.onLoggedIn = () => { ctx.enteredApp = true; };
  store.set('teacher_auth', JSON.stringify({ id: '5', token: 'T' }));
  ctx.tryAutoLogin();
  await drain();
  assert.ok(store.has('teacher_auth'), 'no network is not an expired token');
});

// 21/09 message 20: the teacher page updates itself again. What stays fixed is
// the D4 bug - the modal guard is evaluated when the timer fires and every 15 s
// after that, never once before the grace period starts.
function teacherUpdateContext() {
  const ui = dom();
  const state = { reloads: 0, version: { build: 'b1', pages: { 'teacher.html': 'h1' } }, modalOpen: false };
  const { ctx, timer } = baseContext({
    ...ui,
    location: { pathname: '/teacher.html', reload() { state.reloads++; } },
    fetch: () => Promise.resolve({ ok: true, text: () => Promise.resolve(JSON.stringify(state.version)) })
  });
  ctx.document.querySelector = sel => (sel === '.modal-overlay.active' && state.modalOpen ? { id: 'modalNewClass' } : null);
  load(ctx, section(teacher, '// ===== Auto-update: a banner, a button, and a self-reload ONLY when it is safe', '\nvar deferredInstallPrompt'));
  return { ctx, timer, state, ...ui };
}

test('D4: the teacher banner appears on the second sighting and the page then reloads itself', async () => {
  const { timer, state, nodes } = teacherUpdateContext();
  await drain();
  state.version = { build: 'b2', pages: { 'teacher.html': 'h2' } };
  await timer.advance(120000);
  assert.equal(nodes.get('teacherUpdateBanner') || null, null, 'one sighting proves nothing');
  await timer.advance(120000);
  assert.ok(nodes.get('teacherUpdateBanner'), 'the banner appears on the second');
  await timer.advance(59000);
  assert.equal(state.reloads, 0, 'not before the grace period is over');
  await timer.advance(2000);
  assert.equal(state.reloads, 1, 'and then it updates itself');
  await timer.advance(10 * 60 * 1000);
  assert.equal(state.reloads, 1, 'exactly once');
});

test('D4: an open modal postpones the teacher reload, and the guard is re-checked every 15 s', async () => {
  const { timer, state, nodes } = teacherUpdateContext();
  await drain();
  state.version = { build: 'b2', pages: { 'teacher.html': 'h2' } };
  await timer.advance(120000); await timer.advance(120000);
  assert.ok(nodes.get('teacherUpdateBanner'));
  // the D4 bug: the dialog was opened INSIDE the grace period, after the guard
  // had already been evaluated - and the reload wiped what was typed in it
  state.modalOpen = true;
  await timer.advance(60000);
  assert.equal(state.reloads, 0, 'the dialog opened after the timer was armed, and still counts');
  await timer.advance(5 * 60 * 1000);
  assert.equal(state.reloads, 0);
  state.modalOpen = false;
  await timer.advance(15000);
  assert.equal(state.reloads, 1, 'within one re-check of the dialog closing');
});

test('D4: the teacher button reloads immediately', async () => {
  const { timer, state, nodes } = teacherUpdateContext();
  await drain();
  state.version = { build: 'b2', pages: { 'teacher.html': 'h2' } };
  state.modalOpen = true;
  await timer.advance(120000); await timer.advance(120000);
  nodes.get('swUpdNow').click();
  assert.equal(state.reloads, 1, 'the teacher asked for it himself');
});

// ======================================================================
// student.html
// ======================================================================
function studentContext(extra = {}) {
  const ui = dom(['screenPractice', 'screenFlashcards', 'classStatus']);
  const bank = fakeBank();
  const { ctx, timer, store } = baseContext({
    ...ui,
    API_URL: EXAM_URL, REPORTS_API_URL: REPORTS_URL, API_ORIGIN: 'student-app',
    QuestionBank: bank.api,
    currentName: '', currentClassCode: '', currentLicense: 'B', currentLanguage: 'he',
    currentStudentId: '', currentMode: 'exam', currentCategory: '',
    examQuestions: [], flashcardQuestions: [], userAnswers: [], shuffledOrders: [],
    showLoadingOverlay() {}, hideLoadingOverlay() {},
    showScreen() {}, renderQuestion() { ctx.rendered = (ctx.rendered || 0) + 1; },
    renderFlashcard() {}, stopSpeaking() {}, alert(m) { ctx.alerted = m; },
    ...extra
  });
  load(ctx, section(student, '// ===== D16: progress belongs to a STUDENT', '// Join a class'));
  load(ctx, section(student, '// ===== one transport for the whole page =====', '// Join a class'));
  load(ctx, section(student, 'function classifyCategory(cat)', 'function shuffleAnswers(q,qIdx)'));
  load(ctx, section(student, 'function startPractice(extraParams,onOk,onFail)', '// ===== TTS Engine ====='));
  return { ctx, timer, store, bank, ...ui };
}

test('D16: studentId and every progress key follow class + name, not the device', () => {
  const { ctx, store } = studentContext();
  ctx.currentName = 'דני לוי';
  ctx.currentClassCode = 'AB12';
  const idA = ctx.getStudentId();
  const keyA = ctx.progressKey('student_streak');
  assert.match(idA, /^S[0-9a-f]{8}$/);
  assert.notEqual(keyA, 'student_streak', 'the key is namespaced');

  // same student, different capitalisation / spacing -> the same identity
  ctx.currentName = '  דני   לוי ';
  assert.equal(ctx.getStudentId(), idA, 'trim + collapse: the same person is the same id');
  assert.equal(ctx.progressKey('student_streak'), keyA);

  // a different name on the SAME device -> a different student
  ctx.currentName = 'רות כהן';
  const idB = ctx.getStudentId();
  assert.notEqual(idB, idA);
  assert.notEqual(ctx.progressKey('student_streak'), keyA);

  // and a different class -> different again
  ctx.currentClassCode = 'ZZ99';
  assert.notEqual(ctx.getStudentId(), idB);

  // the wrong-question list of one never reaches the other
  ctx.currentName = 'דני לוי'; ctx.currentClassCode = 'AB12';
  ctx.writeProgress('student_wrong_qs', JSON.stringify([{ id: 1 }]));
  ctx.currentName = 'רות כהן';
  assert.equal(ctx.readProgress('student_wrong_qs', '[]'), '[]', 'student B does not inherit A’s mistakes');
  assert.ok([...store.keys()].some(k => k.startsWith('student_wrong_qs_')));
});

test('D16: the old device-wide keys move once, to the student whose name was saved', () => {
  const { ctx, store } = studentContext();
  store.set('student_name', 'דני');
  store.set('student_streak', '{"count":4}');
  store.set('student_wrong_qs', '[{"id":7}]');
  store.set('student_history', '[{"score":20}]');
  ctx.currentName = 'דני';
  ctx.currentClassCode = 'AB12';
  ctx.migrateLegacyProgress();
  assert.equal(ctx.readProgress('student_streak', '{}'), '{"count":4}', 'the streak followed the student');
  assert.equal(store.has('student_streak'), false, 'and the device-wide copy is gone');
  assert.equal(store.get('student_progress_migrated_v1'), '1');

  // a second run must not move anything again
  store.set('student_streak', '{"count":99}');
  ctx.migrateLegacyProgress();
  assert.equal(ctx.readProgress('student_streak', '{}'), '{"count":4}', 'migration happens exactly once');
});

test('D16: a device whose saved name belongs to someone else keeps its legacy keys untouched', () => {
  const { ctx, store } = studentContext();
  store.set('student_name', 'דני');
  store.set('student_streak', '{"count":4}');
  ctx.currentName = 'רות';
  ctx.currentClassCode = 'AB12';
  ctx.migrateLegacyProgress();
  assert.equal(store.get('student_streak'), '{"count":4}', 'nothing is handed to the wrong student');
  assert.equal(ctx.readProgress('student_streak', '{}'), '{}');
});

test('startPractice hydrates from the grant it was handed, with the per-language correct index', async () => {
  const { ctx, bank } = studentContext();
  ctx.currentLanguage = 'he';
  const enc = (ci, id) => ci ^ (id % 256);
  ctx.fetch = (url) => {
    assert.match(url, /action=startPractice/);
    assert.match(url, /mode=exam/);
    assert.match(url, /language=he/);
    assert.match(url, /license=B/);
    assert.ok(url.indexOf('getExamQuestions') < 0);
    return Promise.resolve({ ok: true, text: () => Promise.resolve(JSON.stringify({
      status: 'ok', mode: 'exam', bank: PRACTICE_BANK,
      questions: [
        { id: 1, topic: 'תמרורים', ci: { he: enc(2, 1), en: enc(0, 1) } },
        { id: 2, topic: 'חוק', ci: { he: enc(1, 2), en: enc(3, 2) } }
      ]
    })) });
  };
  let got = null;
  ctx.startPractice({ mode: 'exam' }, qs => { got = qs; }, m => { throw new Error('unexpected failure ' + m); });
  await drain();
  assert.equal(got.length, 2);
  assert.equal(got[0].text, 'שאלה 1', 'text comes from the gateway, not the wire');
  assert.deepEqual(got[0].answers, ['א', 'ב', 'ג', 'ד']);
  assert.equal(got[0].imageUrl, 'images/TQ_PIC_1.jpg');
  assert.equal(got[0].category, 'תמרורים', 'the blueprint topic drives the category chip');
  assert.equal(ctx.getQCorrectIndex(got[0]), 2);
  assert.equal(ctx.getQCorrectIndex(got[1]), 1);
  assert.equal(bank.grants.length, 1, 'exactly one request for the texts');
  assert.equal(bank.grants[0].grant, PRACTICE_BANK.grant, 'the grant startPractice returned');
  assert.equal(bank.grants[0].url, PRACTICE_BANK.url);
});

test('a practice answer without a grant says the bank is not configured, and opens nothing', async () => {
  const { ctx, bank } = studentContext();
  const enc = (ci, id) => ci ^ (id % 256);
  ctx.fetch = () => Promise.resolve({ ok: true, text: () => Promise.resolve(JSON.stringify({
    status: 'ok', questions: [{ id: 1, topic: 'חוק', ci: { he: enc(0, 1) } }]
  })) });
  let failure = '', opened = false;
  ctx.startPractice({ mode: 'exam' }, () => { opened = true; }, m => { failure = m; });
  await drain();
  assert.equal(opened, false, 'never a practice with blank questions');
  assert.match(failure, /מאגר השאלות אינו מוגדר בשרת/);
  assert.equal(bank.grants.length, 0);
});

test('a language switch is local: no request at all, and the answers already given are preserved', async () => {
  const { ctx, bank } = studentContext();
  const enc = (ci, id) => ci ^ (id % 256);
  ctx.examQuestions = [
    { id: 1, category: '', ciByLang: { he: enc(2, 1), en: enc(0, 1) } },
    { id: 2, category: '', ciByLang: { he: enc(1, 2), en: enc(3, 2) } }
  ];
  await ctx.QuestionBank.loadGrant(PRACTICE_BANK);   // what startPractice did a moment earlier
  ctx.examQuestions.forEach(q => ctx.applyLanguageToQuestion(q, 'he'));
  load(ctx, section(student, 'function switchLanguage(newLang)', 'function finishExam()'));
  ctx.nodes.get('screenPractice').classList.add('active');
  let requests = 0;
  ctx.fetch = () => { requests++; return Promise.reject(new Error('must not ask the server')); };

  ctx.switchLanguage('en');
  await drain();
  assert.equal(requests, 0, 'no getQuestionsByIds, no round trip at all');
  assert.equal(ctx.currentLanguage, 'en');
  assert.equal(ctx.examQuestions[0].text, 'question 1', 'repainted from the English texts already on the device');
  assert.equal(ctx.getQCorrectIndex(ctx.examQuestions[0]), 0, 'and the correct index moved with it');
  assert.equal(ctx.document.documentElement.dir, 'ltr');
  assert.equal(bank.grants.length, 1, 'the switch itself asked the gateway for nothing');

  ctx.switchLanguage('he');
  await drain();
  assert.equal(ctx.examQuestions[0].text, 'שאלה 1');
  assert.equal(ctx.getQCorrectIndex(ctx.examQuestions[0]), 2);
  assert.equal(bank.grants.length, 1);
});

test('an id missing from a translated bank falls back to Hebrew, index included', async () => {
  const { ctx } = studentContext();
  const enc = (ci, id) => ci ^ (id % 256);
  ctx.currentLanguage = 'en';
  ctx.fetch = () => Promise.resolve({ ok: true, text: () => Promise.resolve(JSON.stringify({
    status: 'ok', bank: PRACTICE_BANK,
    questions: [{ id: 907, topic: 'חוק', ci: { he: enc(3, 907), en: enc(1, 907) } }]
  })) });
  let got = null;
  ctx.startPractice({ mode: 'exam' }, qs => { got = qs; }, m => { throw new Error(m); });
  await drain();
  assert.equal(got.length, 1, 'the question is still shown');
  assert.equal(got[0].text, 'שאלה 907', 'in Hebrew');
  assert.equal(got[0].displayLang, 'he');
  assert.equal(ctx.getQCorrectIndex(got[0]), 3, 'scored against the HEBREW key, which is what is on screen');
});

test('an old server that does not know startPractice says so in plain Hebrew', async () => {
  const { ctx } = studentContext();
  ctx.fetch = () => Promise.resolve({ ok: true, text: () => Promise.resolve(
    JSON.stringify({ status: 'error', message: 'Unknown action: startPractice' })) });
  let failure = '';
  ctx.startPractice({ mode: 'exam' }, () => { throw new Error('should not succeed'); }, m => { failure = m; });
  await drain();
  assert.match(failure, /המערכת מתעדכנת/);
});

test('a rate-limited practice request is named, and a degraded backend says "no need to reload"', async () => {
  const { ctx } = studentContext();
  ctx.fetch = () => Promise.resolve({ ok: true, text: () => Promise.resolve(
    JSON.stringify({ status: 'error', code: 'rate_limited' })) });
  let failure = '';
  ctx.startPractice({ mode: 'exam' }, () => {}, m => { failure = m; });
  await drain();
  assert.match(failure, /יותר מדי בקשות/);

  ctx.fetch = () => Promise.resolve({ ok: false, status: 500, text: () => Promise.resolve('<html>busy</html>') });
  ctx.startPractice({ mode: 'exam' }, () => {}, m => { failure = m; });
  await drain();
  assert.match(failure, /אין צורך לרענן/);
});

test('spaced repetition sends its own ids (mode=ids) and never getQuestionsByIds', async () => {
  const { ctx } = studentContext();
  ctx.currentName = 'דני'; ctx.currentClassCode = 'AB12';
  ctx.writeProgress('student_wrong_qs', JSON.stringify([
    { id: '1', license: 'B' }, { id: '2', license: 'B' }, { id: '5', license: 'C' }
  ]));
  load(ctx, section(student, '// Spaced repetition: the ids live on this device', 'function saveToHistory('));
  let seenUrl = '';
  const enc = (ci, id) => ci ^ (id % 256);
  ctx.fetch = (url) => {
    seenUrl = url;
    return Promise.resolve({ ok: true, text: () => Promise.resolve(JSON.stringify({
      status: 'ok', bank: PRACTICE_BANK, questions: [
        { id: 1, topic: 'חוק', ci: { he: enc(0, 1) } },
        { id: 2, topic: 'חוק', ci: { he: enc(1, 2) } }
      ]
    })) });
  };
  let ready = null;
  ctx.getSpacedRepetitionQuestions('B', 'he', qs => { ready = qs; });
  await drain();
  assert.match(seenUrl, /action=startPractice/);
  assert.match(seenUrl, /mode=ids/);
  assert.ok(/ids=(1%2C2|2%2C1)/.test(seenUrl), 'only the B-licence ids, comma separated: ' + seenUrl);
  assert.ok(seenUrl.indexOf('getQuestionsByIds') < 0);
  assert.equal(ready.length, 2);
  assert.equal(ready[0].text.indexOf('שאלה') === 0, true);
});

test('spaced repetition with nothing stored asks the server nothing', async () => {
  const { ctx } = studentContext();
  ctx.currentName = 'דני';
  load(ctx, section(student, '// Spaced repetition: the ids live on this device', 'function saveToHistory('));
  let requests = 0;
  ctx.fetch = () => { requests++; return Promise.reject(new Error('x')); };
  let ready = 'untouched';
  ctx.getSpacedRepetitionQuestions('B', 'he', qs => { ready = qs; });
  await drain();
  assert.ok(Array.isArray(ready) && ready.length === 0, 'an empty list, not a request');
  assert.equal(requests, 0);
});

test('S8: one pass rule, ceil(total * 0.86), for a 30-question exam and a 15-question topic quiz', () => {
  // the rule as showResults computes it
  const passed = (correct, total) => total > 0 && correct >= Math.ceil(total * 0.86);
  assert.equal(passed(26, 30), true, '26/30 is the pass mark');
  assert.equal(passed(25, 30), false);
  assert.equal(passed(30, 30), true);
  assert.equal(passed(13, 15), true, 'ceil(15*0.86) = 13');
  assert.equal(passed(12, 15), false, '12/15 = 80% used to show עבר and store נכשל');
  assert.equal(passed(0, 0), false, 'an empty set is never a pass');
  assert.equal(passed(18, 20), true);
  assert.equal(passed(17, 20), false, 'ceil(20*0.86) = 18');
  // and the page really uses it
  assert.match(student, /var passed=total>0&&correct>=Math\.ceil\(total\*0\.86\);/);
  assert.ok(student.indexOf('pct>=80') < 0, 'the 80% branch is gone');
  assert.match(student, /passed:resultData\.passed \? 1 : 0/, 'and the same verdict is what we send');
});

test('the student update check reloads only when nothing is being practised', async () => {
  const ui = dom(['screenPractice', 'screenFlashcards']);
  let reloads = 0, version = { build: 'b1', pages: { 'student.html': 'h1' } };
  const { ctx, timer } = baseContext({
    ...ui,
    location: { pathname: '/student.html', reload() { reloads++; } },
    fetch: () => Promise.resolve({ ok: true, text: () => Promise.resolve(JSON.stringify(version)) })
  });
  load(ctx, section(student, '// Update check (D3): version.json', '\nvar deferredInstallPrompt'));
  await drain();
  ui.nodes.get('screenPractice').classList.add('active');
  version = { build: 'b2', pages: { 'student.html': 'h2' } };
  await timer.advance(120000);
  await timer.advance(120000);
  await timer.advance(5 * 60 * 1000);
  assert.equal(reloads, 0, 'never in the middle of a practice');
  ui.nodes.get('screenPractice').classList.remove('active');
  await timer.advance(35000);
  assert.equal(reloads, 1, 'and it does happen once the student is idle');
});

test('joining a class still fires after the namespace is resolved', () => {
  // Caught during the rebuild: the namespace has to know the class BEFORE any
  // progress key is touched, but the "did the student type a NEW class code?"
  // test compares against the class joined BEFORE this click. Assigning
  // currentClassCode first made that comparison always false and silently
  // stopped every joinClass call.
  const init = section(student, "var classInput=document.getElementById('inputClassCode')", 'btnBackToWelcome');
  assert.match(init, /var priorClassCode=currentClassCode;/);
  assert.match(init, /if\(classInput&&classInput!==priorClassCode\)\{joinClass\(/);
  assert.ok(init.indexOf('classInput!==currentClassCode') < 0,
    'comparing against the freshly-assigned value would never be true');
  assert.ok(init.indexOf('var priorClassCode=currentClassCode;') <
            init.indexOf('currentClassCode=classInput||savedClass'),
    'the previous class is captured before it is overwritten');
  assert.ok(init.indexOf('migrateLegacyProgress();') <
            init.indexOf('currentStudentId=getStudentId();'),
    'the legacy keys move before the id is derived');
});

test('joinClass adopts a server-side studentId under the namespace of the class being joined', async () => {
  const { ctx, store } = studentContext();
  ctx.currentName = 'דני לוי';
  ctx.fetch = () => Promise.resolve({ ok: true, text: () => Promise.resolve(JSON.stringify({
    status: 'ok', className: 'כיתה א', teacherName: 'מורה',
    existingStudentId: 'S-from-server', license: 'B'
  })) });
  load(ctx, section(student, 'function joinClass(code,name,studentId,onDone){', '// S8: the verdict travels WITH'));
  ctx.nodes.get('classStatus').style = { display: '', color: '' };
  let done = false;
  ctx.joinClass('ab12', ctx.currentName, 'S-local', () => { done = true; });
  await drain();
  assert.equal(done, true);
  assert.equal(ctx.currentStudentId, 'S-from-server', 'the roster id the server already has wins');
  assert.equal(ctx.currentClassCode, 'AB12');
  assert.equal(ctx.readProgress('student_practice_id', ''), 'S-from-server');
  assert.ok([...store.keys()].some(k => k.startsWith('student_practice_id_')),
    'stored under class+name, never as one device-wide key');
});

test('r31: student.html practises against the reports deployment', async () => {
  const { ctx } = studentContext();
  let seenUrl = '';
  const enc = (ci, id) => ci ^ (id % 256);
  ctx.fetch = (url) => {
    seenUrl = url;
    return Promise.resolve({ ok: true, text: () => Promise.resolve(JSON.stringify({
      status: 'ok', bank: PRACTICE_BANK, questions: [{ id: 1, topic: 'חוק', ci: { he: enc(0, 1) } }]
    })) });
  };
  ctx.startPractice({ mode: 'exam' }, () => {}, m => { throw new Error(m); });
  await drain();
  assert.equal(seenUrl.indexOf(REPORTS_URL + '?'), 0, 'startPractice never runs during an exam');
  assert.match(student, /\r\nvar REPORTS_API_URL=API_URL;\r\n/);
  assert.match(student, /createApi\(\{ apiUrl: REPORTS_API_URL, origin: API_ORIGIN \}\)/);
});

test('the retired student code really is gone', () => {
  for (const dead of ['getExamQuestions&', 'getQuestionsByIds&', '_practiceTranslations',
                      'TRANSLATION_VARS', 'getTransDict', 'getRuQuestion', 'getFilteredQuestions',
                      'buildCategoryQuiz(_license', 'buildFlashcardSet', 'loadTranslationFile',
                      'QuestionBank.load(', 'QuestionBank.prefetch']) {
    assert.ok(student.indexOf(dead) < 0, dead + ' must not appear in student.html');
  }
  assert.match(student, /<script src="shared\/transport\.js"><\/script>/);
  assert.match(student, /<script src="shared\/bank\.js"><\/script>/);
});

// ======================================================================
// exam.html — the standalone practice exam (no examiner, Hebrew only)
// ======================================================================
test('exam.html standalone: startPractice with the id number, texts from the Hebrew bank', async () => {
  const ui = dom();
  const examArea = ui.element('examArea');
  const bank = fakeBank();
  const { ctx } = baseContext({
    ...ui,
    QUESTIONS_API_URL: EXAM_URL, REPORTS_API_URL: REPORTS_URL,
    QuestionBank: bank.api,
    TOTAL_QUESTIONS: 2,
    ExamTransportOrigin: 'examinee-app'
  });
  load(ctx, section(examPage, '  // One shared transport: bounded requests', '  // ========== Config =========='));
  let seenUrl = '';
  const enc = (ci, id) => ci ^ (id % 256);
  ctx.fetch = (url) => {
    seenUrl = url;
    return Promise.resolve({ ok: true, text: () => Promise.resolve(JSON.stringify({
      status: 'ok', bank: PRACTICE_BANK, questions: [
        { id: 1, topic: 'תמרורים', ci: { he: enc(2, 1) } },
        { id: 2, topic: 'חוק', ci: { he: enc(1, 2) } }
      ]
    })) });
  };
  const resp = await ctx.api.get({ action: 'startPractice', mode: 'exam', license: 'B', language: 'he', standaloneIdNumber: '123456789' });
  assert.equal(seenUrl.indexOf(REPORTS_URL + '?'), 0, 'r31: the standalone page is a reports client too');
  assert.match(seenUrl, /action=startPractice/);
  assert.match(seenUrl, /standaloneIdNumber=123456789/);
  assert.match(seenUrl, /origin=examinee-app/);
  await bank.api.loadGrant(resp.bank);
  assert.equal(bank.grants[0].grant, PRACTICE_BANK.grant, 'the standalone page loads the grant it was given');
  const built = resp.questions.map(r => {
    const e = bank.api.get(r.id, 'he', 'B');
    return { id: r.id, text: e.text, answers: e.answers, imageUrl: bank.api.imageUrl(e), category: r.topic, ci: r.ci.he };
  });
  assert.equal(built.length, 2);
  assert.equal(built[0].text, 'שאלה 1');
  assert.equal(built[0].imageUrl, 'images/TQ_PIC_1.jpg');
  // the scoring the page does at render time, unchanged
  assert.equal(built[0].ci ^ (built[0].id % 256), 2);
  assert.equal(built[1].ci ^ (built[1].id % 256), 1);
});

test('exam.html: the legacy client-side picker and the old action are gone', () => {
  for (const dead of ['_legacyBuildExam_DEPRECATED', '_legacyGetFilteredQuestions_DEPRECATED',
                      "action=getExamQuestions", 'window.QUESTIONS']) {
    assert.ok(examPage.indexOf(dead) < 0, dead + ' must not appear in exam.html');
  }
  assert.match(examPage, /action: 'startPractice'/);
  assert.match(examPage, /<script src="shared\/bank\.js"><\/script>/);
});

test('r31: exam.html keeps QUESTIONS_API_URL and routes through REPORTS_API_URL', () => {
  // The name QUESTIONS_API_URL stays: it is the line that says "this is the
  // OTHER Apps Script project, the one with the question DB", and the reports
  // url is derived from it so there is still exactly one url to edit per page.
  assert.match(examPage, /\r\n  var QUESTIONS_API_URL = '/);
  assert.match(examPage, /\r\n  var REPORTS_API_URL = QUESTIONS_API_URL;\r\n/);
  assert.match(examPage, /createApi\(\{ apiUrl: REPORTS_API_URL, origin: 'examinee-app' \}\)/);
});

test('exam.html: the standalone flow requires the grant and never loads a bank by language', () => {
  assert.match(examPage, /if \(!resp\.bank \|\| !resp\.bank\.url \|\| !resp\.bank\.grant\) \{/,
    'no grant is a named failure, not an exam with blank questions');
  assert.match(examPage, /message: BANK_NOT_CONFIGURED_TEXT/);
  assert.match(examPage, /QuestionBank\.loadGrant\(resp\.bank\)/);
  assert.ok(examPage.indexOf("QuestionBank.load(") < 0, 'QuestionBank.load no longer exists');
});

// ======================================================================
// service worker
// ======================================================================
test('sw-student.js: parses, is GET-only and precaches the shared modules', () => {
  const src = fs.readFileSync(path.join(app, 'sw-student.js'), 'utf8');
  const listeners = [];
  const self = { addEventListener: (t, cb) => listeners.push([t, cb]), skipWaiting() {}, clients: { claim() {} } };
  vm.runInNewContext(src, {
    self,
    caches: { open: () => Promise.resolve({ addAll: () => Promise.resolve() }), keys: () => Promise.resolve([]), match: () => Promise.resolve(null), delete: () => Promise.resolve() },
    fetch: () => Promise.resolve(), console: quiet
  });
  assert.deepEqual(listeners.map(l => l[0]), ['install', 'activate', 'fetch']);
  assert.match(src, /^var CACHE_NAME = '[a-z]+-[a-z0-9]+';$/m);
  assert.match(src, /'\.\/shared\/transport\.js'/);
  assert.match(src, /'\.\/shared\/bank\.js'/);
  const fetchHandler = listeners.find(l => l[0] === 'fetch')[1];
  let responded = false;
  fetchHandler({ request: { method: 'POST', url: 'https://example/x' }, respondWith: () => { responded = true; } });
  assert.equal(responded, false, 'D7: Cache.put throws on a non-GET request');
  fetchHandler({ request: { method: 'GET', url: 'https://example/student.html?cb=1' }, respondWith: () => { responded = true; } });
  assert.equal(responded, true);
  assert.match(src, /ignoreSearch: true/, 'a cache-busted shell still matches its cached copy offline');
  // The questions are not served from this origin any more - they come from the
  // gateway, against a grant, and are never put in a cache.
  assert.ok(!/bank\/manifest\.json/.test(src), 'no bank manifest in the shell');
  assert.ok(!/'\.\/bank\//.test(src), 'no bank file in the shell');
});
