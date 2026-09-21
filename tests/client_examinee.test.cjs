// examinee.html — the real page, executed. Every inline script runs in order in
// a VM, on top of the real shared/transport.js and shared/bank.js, with a fake
// clock, a small DOM and a synthetic network. Nothing is stubbed that the page
// itself owns: the API calls, the bank reads, the exam start, the submit ladder
// and the anti-cheat timers are the shipping code.
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const app = path.resolve(__dirname, '..');
const examinee = fs.readFileSync(path.join(app, 'examinee.html'), 'utf8');
const TRANSPORT = fs.readFileSync(path.join(app, 'shared', 'transport.js'), 'utf8');
const BANK = fs.readFileSync(path.join(app, 'shared', 'bank.js'), 'utf8');

const TOTAL = 30;
const plain = v => JSON.parse(JSON.stringify(v));   // VM arrays are not host arrays
const quiet = { log() {}, warn() {}, error() {} };
async function drain() { for (let i = 0; i < 40; i++) await Promise.resolve(); }

function section(src, start, end) {
  const i = src.indexOf(start), j = src.indexOf(end, i + start.length);
  assert.ok(i >= 0 && j > i, 'source section found: ' + start);
  return src.slice(i, j);
}

// ---------- fake clock ----------
class Timers {
  constructor() { this.now = 0; this.nextId = 0; this.jobs = new Map(); }
  set = (cb, ms) => { const id = ++this.nextId; this.jobs.set(id, { cb, at: this.now + Number(ms || 0), every: 0 }); return id; };
  setInterval = (cb, ms) => { const id = ++this.nextId; this.jobs.set(id, { cb, at: this.now + Number(ms || 0), every: Number(ms || 0) }); return id; };
  clear = id => this.jobs.delete(id);
  async advance(ms) {
    const until = this.now + ms;
    for (let guard = 0; guard < 20000; guard++) {
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

// ---------- minimal DOM (enough to render a question and click an answer) ----------
function makeDom() {
  const byId = new Map();
  const docHandlers = new Map();
  const document = {
    title: '', visibilityState: 'visible',
    addEventListener(type, cb) { if (!docHandlers.has(type)) docHandlers.set(type, []); docHandlers.get(type).push(cb); },
    removeEventListener(type, cb) { const l = docHandlers.get(type) || []; const i = l.indexOf(cb); if (i >= 0) l.splice(i, 1); },
    getElementById: id => byId.get(id) || null
  };
  // Enough CSS for this page: tag, .a, .a.b, #id and any comma list of those.
  // Descendant selectors are not supported (the page only uses them in paths
  // that already tolerate a null, e.g. '#waitingPhase h2').
  const matches = (el, sel) => String(sel).split(',').some(one => {
    one = one.trim();
    if (!one || one.includes(' ')) return false;
    const id = /#([\w-]+)/.exec(one);
    if (id && el.id !== id[1]) return false;
    const classes = (one.match(/\.[\w-]+/g) || []).map(c => c.slice(1));
    if (!classes.every(c => el.classList.contains(c))) return false;
    const tag = /^[a-zA-Z]+/.exec(one);
    if (tag && String(el.tagName || '').toLowerCase() !== tag[0].toLowerCase()) return false;
    return true;
  });
  const descendants = el => { const out = []; (function walk(n) { for (const c of n.children || []) { out.push(c); walk(c); } })(el); return out; };

  function element(tag) {
    const classes = new Set();
    const el = {
      tagName: String(tag || 'div').toUpperCase(), textContent: '', value: '', checked: false, disabled: false,
      type: '', src: '', alt: '', loading: '', referrerPolicy: '', dir: '', options: [],
      style: {}, children: [], parentNode: null, handlers: {}, attrs: {},
      classList: { add: v => classes.add(v), remove: v => classes.delete(v), contains: v => classes.has(v), values: classes },
      addEventListener(type, cb) { (this.handlers[type] = this.handlers[type] || []).push(cb); },
      removeEventListener(type, cb) { this.handlers[type] = (this.handlers[type] || []).filter(x => x !== cb); },
      setAttribute(k, v) { this.attrs[k] = v; }, getAttribute(k) { return this.attrs[k]; },
      appendChild(child) { this.children.push(child); child.parentNode = this; return child; },
      insertBefore(child, before) { const at = this.children.indexOf(before); this.children.splice(at < 0 ? this.children.length : at, 0, child); child.parentNode = this; return child; },
      // A detached node must stop answering getElementById, as in a real DOM —
      // otherwise the page finds a node whose parentNode is already null.
      removeChild(child) {
        this.children = this.children.filter(c => c !== child);
        child.parentNode = null;
        if (child.id && byId.get(child.id) === child) byId.delete(child.id);
        return child;
      },
      querySelectorAll(sel) { return descendants(this).filter(c => matches(c, sel)); },
      querySelector(sel) { return this.querySelectorAll(sel)[0] || null; },
      closest(sel) { let n = this; while (n) { if (matches(n, sel)) return n; n = n.parentNode; } return null; },
      focus() { document.activeElement = el; },
      fire(type, ev) { for (const cb of this.handlers[type] || []) cb.call(el, ev || {}); },
      click() { if (!this.disabled) this.fire('click', { stopPropagation() {}, preventDefault() {} }); }
    };
    let id = '';
    Object.defineProperty(el, 'id', { get: () => id, set(v) { id = String(v); if (id) byId.set(id, el); } });
    Object.defineProperty(el, 'className', { get: () => [...classes].join(' '), set(v) { classes.clear(); String(v).split(/\s+/).filter(Boolean).forEach(c => classes.add(c)); } });
    let html = '';
    Object.defineProperty(el, 'innerHTML', { get: () => html, set(v) {
      html = String(v);
      el.children = [];
      for (const m of html.matchAll(/id="([^"]+)"/g)) { const child = element('div'); child.id = m[1]; el.appendChild(child); }
    } });
    return el;
  }
  document.createElement = tag => element(tag);
  document.body = element('body');
  document.head = element('head');
  document.documentElement = element('html');
  document.querySelectorAll = sel => descendants(document.body).filter(c => matches(c, sel));
  document.querySelector = sel => document.querySelectorAll(sel)[0] || null;
  const setVisibility = state => { document.visibilityState = state; for (const cb of docHandlers.get('visibilitychange') || []) cb(); };
  return { document, element, byId, setVisibility, docHandlers };
}

function memoryStore() {
  const entries = new Map();
  const store = {
    entries, reject: () => false,
    get length() { return entries.size; }, key: i => [...entries.keys()][i] ?? null,
    getItem: k => entries.get(String(k)) ?? null,
    setItem(k, v) { if (store.reject(String(k))) throw Object.assign(new Error('quota'), { name: 'QuotaExceededError' }); entries.set(String(k), String(v)); },
    removeItem: k => entries.delete(String(k))
  };
  return store;
}

// ---------- the synthetic bank ----------
const LANGS = ['he', 'ru', 'en', 'ar', 'fr', 'es', 'am'];
const BANK_FILES = {};
for (const lang of LANGS) {
  BANK_FILES[lang] = [];
  for (let id = 1; id <= 60; id++) {
    // en/fr/es/ar list their answers in their own order — the 20/09 finding.
    const own = (lang === 'en' || lang === 'fr' || lang === 'es' || lang === 'ar');
    const answers = own ? [lang + '-D' + id, lang + '-C' + id, lang + '-B' + id, lang + '-A' + id]
                        : [lang + '-A' + id, lang + '-B' + id, lang + '-C' + id, lang + '-D' + id];
    BANK_FILES[lang].push({ id, t: lang + ' question ' + id, a: answers, i: id % 5 === 0 ? 'TQ_PIC_' + id + '.jpg' : '' });
  }
}
const MANIFEST = { build: 'bank-test', langs: {} };
for (const lang of LANGS) MANIFEST.langs[lang] = { sha: lang + 'sha', count: BANK_FILES[lang].length };

const SERVER_QUESTIONS = Array.from({ length: TOTAL }, (_, i) => ({
  id: i + 1, order: [[2, 0, 3, 1], [1, 3, 0, 2], [0, 1, 2, 3], [3, 2, 1, 0]][i % 4], topic: 'חוק'
}));

// ---------- the whole page ----------
function completePage({ local = memoryStore(), session = memoryStore(), reply, gateway = '', onReload, userAgent = 'Synthetic desktop', touchPoints = 0 } = {}) {
  const ui = makeDom();
  const timer = new Timers();
  const requests = [], beacons = [];
  let reloads = 0;
  // Every element the markup declares, so getElementById never returns null.
  for (const tag of examinee.slice(0, examinee.indexOf('<script')).matchAll(/<[^>]+\bid="([^"]+)"[^>]*>/g)) {
    const el = ui.element('div');
    el.id = tag[1];
    const cls = /\bclass="([^"]+)"/.exec(tag[0]);
    if (cls) el.className = cls[1];
    ui.document.body.appendChild(el);
  }
  const windowEvents = new Map();
  const clock = class extends Date { static now() { return timer.now; } };

  const answer = request => {
    requests.push(request);
    if (reply) {
      const given = reply(request);
      if (given !== undefined) return given;
    }
    return defaultReply(request, gateway);
  };

  const ctx = {
    console: quiet, Date: clock, Math, JSON, Promise, Error, TypeError, String, Number, Object, Array, RegExp,
    isFinite, parseFloat, parseInt, encodeURIComponent, decodeURIComponent, URL, URLSearchParams, Blob: class { constructor(parts) { this.parts = parts; } },
    setTimeout: timer.set, clearTimeout: timer.clear, setInterval: timer.setInterval, clearInterval: timer.clear,
    AbortController,
    document: ui.document, screen: { width: 1920, height: 1080, availWidth: 1920, availHeight: 1080, isExtended: false },
    innerWidth: 1920, innerHeight: 1080, outerWidth: 1920, outerHeight: 1080,
    localStorage: local, sessionStorage: session,
    navigator: { userAgent: userAgent, platform: 'Synthetic', maxTouchPoints: touchPoints, onLine: true,
      sendBeacon(url, blob) { beacons.push(JSON.parse(blob.parts[0])); return true; } },
    history: { pushState() {} },
    location: { search: '', pathname: '/examinee.html', reload() { reloads++; if (onReload) onReload(); } },
    fetch(url, opts = {}) {
      const isPost = opts && opts.method === 'POST';
      const request = isPost ? JSON.parse(opts.body) : Object.fromEntries(new URL(url, 'https://synthetic.test/').searchParams);
      request.__url = String(url);
      const data = answer(request);
      if (data && data.__network) return Promise.reject(new TypeError('Failed to fetch'));
      if (data && data.__hang) return new Promise(() => {});
      const body = data && data.__raw !== undefined ? data.__raw : JSON.stringify(data);
      return Promise.resolve({ ok: !(data && data.__status >= 400), status: (data && data.__status) || 200,
        text: () => Promise.resolve(body), json: () => Promise.resolve(JSON.parse(body)) });
    }
  };
  ctx.window = ctx;
  ctx.self = ctx;
  ctx.addEventListener = (name, cb) => { if (!windowEvents.has(name)) windowEvents.set(name, new Set()); windowEvents.get(name).add(cb); };
  ctx.removeEventListener = (name, cb) => { windowEvents.get(name)?.delete(cb); };
  vm.createContext(ctx);
  vm.runInContext(TRANSPORT, ctx, { filename: 'shared/transport.js' });
  vm.runInContext(BANK, ctx, { filename: 'shared/bank.js' });
  ctx.ExamTransport._setJitter(ms => ms);      // exact waits on the fake clock

  const exposure = `
  window.__t = {
    state: function() { return { screen: (function(){ var a = document.querySelector('.screen.active'); return a ? a.id : ''; })(),
      inProgress: examInProgress, submitted: examSubmitted, id: examineeData.idNumber, token: examineeToken,
      lang: getExamLang(), questions: activeQuestions, orders: shuffledOrders, answers: userAnswers,
      timeMinutes: examTimeMinutes, deadline: examDeadline, dq: tabSwitchDQConfirmed, warnings: tabSwitchWarnings,
      gateway: gatewayUrl(), audio: sessionData.audioMode }; },
    render: renderQuestion, finish: renderExamDone, switchLang: switchExamLanguage,
    answerCurrent: function(i) { var b = document.querySelectorAll('.answer-audio-btn')[i]; if (b) b.click(); return !!b; },
    goTo: function(i) { currentIndex = i; renderQuestion(); },
    images: imageSources, setDegraded: function() { ExamTransport.noteTransport({ transport: 'http' }); },
    retryDelay: submitRetryDelayMs, hasPending: hasAnyPendingResult
  };
`;
  const scripts = [...examinee.matchAll(/<script\b[^>]*>([\s\S]*?)<\/script>/gi)].map(m => m[1]).filter(code => code.trim());
  for (let i = 0; i < scripts.length; i++) {
    let code = scripts[i];
    if (code.includes('function renderExamDone()')) {
      const end = code.lastIndexOf('})();');
      assert.ok(end >= 0, 'closure end found');
      code = code.slice(0, end) + exposure + code.slice(end);
    }
    vm.runInContext(code, ctx, { filename: 'examinee.html:inline-' + (i + 1) });
  }
  assert.ok(ctx.__t, 'the whole inline script reached its end');

  const page = {
    ctx, timer, ui, requests, beacons, local, session,
    el: id => ui.byId.get(id),
    get reloads() { return reloads; },
    t: ctx.__t,
    dispatch(name, ev = {}) { for (const cb of [...(windowEvents.get(name) || [])]) cb(ev); },
    setVisibility: ui.setVisibility,
    actions: () => requests.map(r => r.action || (r.kind ? 'gateway:' + r.kind : '?')),
    sent: action => requests.filter(r => r.action === action)
  };
  return page;
}

function defaultReply(request, gateway) {
  const url = String(request.__url || '');
  if (url.indexOf('/v1/poll') !== -1) {
    return request.kind === 'approval' ? { status: 'ok', approval: 'approved', audioMode: 'off', examMinutes: 40 }
                                       : { status: 'ok', examStatus: 'in_exam', extraMinutes: 0 };
  }
  if (url.indexOf('bank/manifest.json') !== -1) return MANIFEST;
  const bank = /bank\/([a-z]{2})\.json/.exec(url);
  if (bank) return BANK_FILES[bank[1]] || { __status: 404, __raw: 'not found' };
  if (url.indexOf('version.json') !== -1) return { build: 'v1', pages: { 'examinee.html': 'hash-1' } };
  switch (request.action) {
    case 'getSessionInfo':
      return { status: 'ok', build: 'r25', session: { site: 'בדיקת נתונים', license: 'B', language: 'he', audioMode: 'off',
        examinerName: 'בוחן', classroom: '1', sites: ['בדיקת נתונים'], gateway: { url: gateway } } };
    case 'registerExaminee': return { status: 'ok', examineeToken: 'tok-1' };
    case 'checkApproval': return { status: 'ok', approval: 'approved', audioMode: 'off', examMinutes: 40 };
    case 'getExamStatus': return { status: 'ok', examStatus: 'in_exam', extraMinutes: 0 };
    case 'startExam': return { status: 'ok', build: 'r25', examMinutes: 40, extraMinutes: 0, audioMode: 'off',
      language: request.language, license: request.license, registeredAt: '', questions: SERVER_QUESTIONS };
    case 'submitResult': return { status: 'ok', waLink: '' };
    default: return { status: 'ok' };
  }
}

/** code entry → registration → approval → instructions. */
async function register(page, { language = 'he', license = 'B' } = {}) {
  page.el('sessionCodeInput').value = 'ABC12345';
  page.el('codeSubmitBtn').click();
  await drain();
  for (const [id, value] of [['idNumber', '123456789'], ['firstName', 'ישראל'], ['lastName', 'ישראלי'], ['phoneNumber', '0501234567']]) page.el(id).value = value;
  page.el('langSelect').value = language;
  page.el('licenseSelect').value = license;
  page.el('populationSelect').value = 'צבא';
  page.el('registerBtn').click();
  await drain();
  await page.timer.advance(100);
  await drain();
}

async function startExam(page) {
  page.el('airplaneCheckbox').checked = true;
  page.el('airplaneCheckbox').fire('change');
  page.el('startExamBtn').click();
  await drain();
  await page.timer.advance(50);
  await drain();
}

// ===================== 1. exam start =====================
test('start: ONE startExam call builds the 30 questions from the static bank, in the server order', async () => {
  const page = completePage();
  await register(page);
  await startExam(page);
  const state = page.t.state();
  assert.equal(state.inProgress, true);
  assert.equal(state.screen, 'screenExam');
  assert.equal(state.questions.length, TOTAL);
  const apiCalls = page.requests.filter(r => r.action).map(r => r.action);
  assert.deepEqual(apiCalls.slice(0, 4), ['getSessionInfo', 'registerExaminee', 'checkApproval', 'startExam'],
    'four calls take an examinee from the code screen into the exam');
  assert.equal(page.sent('startExam').length, 1);
  assert.equal(page.sent('getExamQuestions').length, 0);
  assert.equal(page.sent('registerExamQuestions').length, 0);
  assert.equal(page.sent('markExamStarted').length, 0);
  // the rendered answers are the bank's, permuted by the server's order
  const shown = page.el('examArea').querySelectorAll('.ans-text').map(n => n.textContent);
  const entry = BANK_FILES.he[0];
  assert.deepEqual(shown, SERVER_QUESTIONS[0].order.map(i => entry.a[i]));
  assert.equal(page.el('examArea').querySelector('.q-text').textContent, entry.t);
  // and the client is never told which one is right
  assert.deepEqual(plain(state.orders).map(o => o.correctIdx), Array(TOTAL).fill(null));
  assert.deepEqual(plain(state.orders[0].order), SERVER_QUESTIONS[0].order);
});

test('start: a retry after a lost answer reuses the identical question map', async () => {
  let attempts = 0;
  const page = completePage({ reply(request) {
    if (request.action !== 'startExam') return undefined;
    attempts++;
    if (attempts === 1) return { __raw: '<html>Google is having trouble</html>' };   // the answer never arrives as JSON
    return undefined;   // the server is idempotent: the same map comes back
  } });
  await register(page);
  await startExam(page);
  assert.equal(page.t.state().inProgress, false, 'a lost answer does not start an exam');
  assert.ok(page.el('examStartRetryStatus'), 'the examinee gets a retry with a cooldown');
  await page.timer.advance(5000);
  page.el('retryExamStartBtn').click();
  await drain(); await page.timer.advance(50); await drain();
  assert.equal(page.sent('startExam').length, 2);
  const state = page.t.state();
  assert.equal(state.inProgress, true);
  assert.deepEqual(plain(state.questions).map(q => q.id), SERVER_QUESTIONS.map(q => q.id));
  assert.deepEqual(plain(state.orders).map(o => o.order), SERVER_QUESTIONS.map(q => q.order));
});

test('start: an old server that does not know startExam says the system is updating', async () => {
  const page = completePage({ reply: r => r.action === 'startExam' ? { status: 'error', message: 'Unknown action: startExam' } : undefined });
  await register(page);
  await startExam(page);
  const html = page.el('examArea').innerHTML;
  assert.match(html, /המערכת מתעדכנת/);
  assert.equal(page.el('retryExamStartBtn').disabled, true, 'with a cooldown instead of a re-click storm');
  await page.timer.advance(60000);
  assert.equal(page.el('retryExamStartBtn').disabled, false);
});

test('start: a bad answer permutation falls back to the natural order instead of scrambling', async () => {
  const page = completePage({ reply(request) {
    if (request.action !== 'startExam') return undefined;
    const questions = SERVER_QUESTIONS.map((q, i) => i === 0 ? { id: q.id, order: [0, 0, 1, 2], topic: q.topic } : q);
    return { status: 'ok', examMinutes: 40, extraMinutes: 0, audioMode: 'off', questions };
  } });
  await register(page);
  await startExam(page);
  assert.deepEqual(plain(page.t.state().orders[0].order), [0, 1, 2, 3]);
});

test('start: examMinutes and a pre-granted extension both reach the clock', async () => {
  const page = completePage({ reply: r => r.action === 'startExam'
    ? { status: 'ok', examMinutes: 50, extraMinutes: 10, audioMode: 'off', questions: SERVER_QUESTIONS } : undefined });
  await register(page);
  await startExam(page);
  const state = page.t.state();
  assert.equal(state.timeMinutes, 50);
  assert.ok(Math.abs(state.deadline - page.timer.now - 60 * 60 * 1000) < 1000, '50 authorised + 10 granted');
});

// ===================== 2. language =====================
test('language: a switch is local — no request — and the answers follow the new language', async () => {
  const page = completePage();
  await register(page);
  await startExam(page);
  const before = page.requests.length;
  page.t.switchLang('ru');
  await drain();
  assert.equal(page.requests.length, before, 'every bank was prefetched while waiting for approval');
  assert.equal(page.t.state().lang, 'ru');
  const shown = page.el('examArea').querySelectorAll('.ans-text').map(n => n.textContent);
  assert.deepEqual(shown, SERVER_QUESTIONS[0].order.map(i => BANK_FILES.ru[0].a[i]));
});

test('language: an answer given in another ORDER group is shown frozen, and D11 keeps it until a new pick', async () => {
  const page = completePage();
  await register(page);
  await startExam(page);
  page.t.answerCurrent(1);
  const first = plain(page.t.state().answers[0]);
  assert.equal(first.langAtAnswer, 'he');
  assert.equal(first.q, BANK_FILES.he[0].t);
  assert.deepEqual(first.a, SERVER_QUESTIONS[0].order.map(i => BANK_FILES.he[0].a[i]));
  const chosenText = first.a[1];

  page.t.switchLang('en'); await drain();
  let buttons = page.el('examArea').querySelectorAll('.answer-audio-btn');
  assert.ok(buttons.every(b => b.disabled), 'the frozen list cannot be clicked');
  assert.deepEqual(page.el('examArea').querySelectorAll('.ans-text').map(n => n.textContent), first.a,
    'the answers are shown exactly as they were seen');
  assert.ok(page.el('examArea').querySelector('.frozen-answer-notice'), 'and the notice explains why');

  // "answer again in the current language" — D11: nothing is lost yet
  page.el('examArea').querySelector('.frozen-answer-reanswer').click();
  await drain();
  const kept = plain(page.t.state().answers[0]);
  assert.ok(kept, 'the previous answer is still recorded');
  assert.equal(kept.q, first.q);
  assert.equal(kept.a[kept.chosenIndex], chosenText);
  buttons = page.el('examArea').querySelectorAll('.answer-audio-btn');
  assert.ok(buttons.every(b => !b.disabled), 'the current-language list is live');
  assert.ok(buttons.every(b => !b.classList.contains('selected')), 'and nothing is highlighted: the position would lie');
  assert.ok(page.el('examArea').querySelector('.keep-previous-notice'));

  // leaving the question and coming back must not lose it either
  page.t.goTo(1); page.t.goTo(0); await drain();
  assert.ok(page.t.state().answers[0], 'still there');

  // only an actual click replaces it
  page.t.answerCurrent(2);
  const replaced = plain(page.t.state().answers[0]);
  assert.equal(replaced.langAtAnswer, 'en');
  assert.equal(replaced.chosenIndex, 2);
  assert.deepEqual(replaced.a, SERVER_QUESTIONS[0].order.map(i => BANK_FILES.en[0].a[i]));
});

test('language: he/ru/am share an order, so an answer given in Hebrew is NOT frozen in Russian', async () => {
  const page = completePage();
  await register(page);
  await startExam(page);
  page.t.answerCurrent(0);
  page.t.switchLang('ru'); await drain();
  const buttons = page.el('examArea').querySelectorAll('.answer-audio-btn');
  assert.ok(buttons.every(b => !b.disabled));
  assert.equal(buttons[0].classList.contains('selected'), true, 'the same position is the same answer');
  assert.equal(page.el('examArea').querySelector('.frozen-answer-notice'), null);
});

test('language: a language whose bank cannot be fetched keeps the exam running', async () => {
  const page = completePage({ reply: r => String(r.__url).includes('bank/am.json') ? { __network: true } : undefined });
  await register(page);
  await startExam(page);
  page.t.switchLang('am');
  await drain(); await page.timer.advance(100); await drain();
  assert.equal(page.t.state().lang, 'he', 'the exam stays in the language it was in');
  assert.match(page.el('examNotice').textContent, /לא ניתן לטעון/);
  assert.equal(page.t.state().inProgress, true);
});

// ===================== 3. submit =====================
test('submit: every question carries the texts as displayed, unanswered ones are -1, and the client log rides along', async () => {
  const page = completePage();
  await register(page);
  await startExam(page);
  page.t.answerCurrent(2);                       // Q1 in Hebrew
  page.t.goTo(1); page.t.switchLang('en'); await drain();
  page.t.answerCurrent(0);                       // Q2 in English
  page.t.finish();
  await drain();
  const [submit] = page.sent('submitResult');
  assert.ok(submit, 'the result was sent');
  assert.equal(submit.answers.length, TOTAL);
  assert.equal(submit.total, TOTAL);
  assert.equal('score' in submit, false, 'the client cannot know the score and does not pretend to');
  assert.equal('passed' in submit, false);
  assert.equal('wrongAnswers' in submit, false, 'the server builds the feedback');
  assert.ok(Array.isArray(submit.clientLog));
  assert.ok(JSON.stringify(submit.clientLog).length <= 2048);

  const q1 = submit.answers[0], q2 = submit.answers[1], q3 = submit.answers[2];
  assert.deepEqual([q1.qIdx, q1.selected, q1.langAtAnswer], [0, 2, 'he']);
  assert.equal(q1.q, BANK_FILES.he[0].t);
  assert.deepEqual(q1.a, SERVER_QUESTIONS[0].order.map(i => BANK_FILES.he[0].a[i]));
  assert.deepEqual([q2.qIdx, q2.selected, q2.langAtAnswer], [1, 0, 'en']);
  assert.equal(q2.q, BANK_FILES.en[1].t);
  assert.deepEqual(q2.a, SERVER_QUESTIONS[1].order.map(i => BANK_FILES.en[1].a[i]));
  assert.equal(q3.selected, -1, 'an unanswered question is still reported');
  assert.equal(q3.langAtAnswer, 'en', 'in the language the examinee was last looking at');
  assert.equal(q3.q, BANK_FILES.en[2].t);
  assert.deepEqual(q3.a, SERVER_QUESTIONS[2].order.map(i => BANK_FILES.en[2].a[i]));
  assert.equal(submit.languageHistory.join('>'), 'he>en');
});

test('submit: the examinee never sees a score, only that the exam was submitted', async () => {
  const page = completePage();
  await register(page);
  await startExam(page);
  page.t.answerCurrent(0);
  page.t.finish();
  await drain();
  assert.equal(page.t.state().screen, 'screenDone');
  const done = page.el('screenDone').innerHTML + page.el('doneTitle').textContent + page.el('doneResultsText').textContent;
  assert.match(page.el('submitStatusBanner').innerHTML, /התקבלה|שולח/);
  assert.ok(!/\b(26|30)\s*\/\s*30\b/.test(done), 'no score anywhere on the done screen');
  assert.equal(page.el('doneNameDisplay').textContent, 'ישראל ישראלי');
});

test('submit: D18 — three fast tries, then a growing wait that never exceeds 120 s, never faster than 30 s while degraded', async () => {
  const page = completePage({ reply: r => r.action === 'submitResult' ? { __network: true } : undefined });
  await register(page);
  await startExam(page);
  page.t.finish();
  await drain();
  assert.equal(page.sent('submitResult').length, 1);
  await page.timer.advance(3000); assert.equal(page.sent('submitResult').length, 2);
  await page.timer.advance(3000); assert.equal(page.sent('submitResult').length, 3);
  // a fresh attempt key, so the ladder is read from its first rung
  const key = JSON.stringify(['LADDER', 'LADDER', 'LADDER']);
  assert.equal(page.t.retryDelay(key, 3), 3000, 'the first three tries are fast');
  const ladder = [];
  for (let i = 0; i < 12; i++) ladder.push(page.t.retryDelay(key, 1));
  assert.equal(ladder[0], 15000);
  assert.ok(ladder.every((ms, i) => i === 0 || ms >= ladder[i - 1]), 'monotonic: ' + ladder.join(','));
  assert.ok(ladder.every(ms => ms <= 120000), 'never past the 120 s ceiling: ' + ladder.join(','));
  assert.equal(ladder[ladder.length - 1], 120000);
  page.t.setDegraded();
  assert.ok(page.t.retryDelay(JSON.stringify(['x', 'y', 'z']), 1) >= 30000, 'a degraded backend is never asked every 15 s');
  assert.ok(page.el('submitFailBanner'), 'and the examinee is told loudly that the result is not on the server yet');
  assert.equal(page.t.hasPending(), true, 'while the result itself stays on the device');
});

test('submit: a confirmed result clears the local copy and the close guard', async () => {
  const page = completePage();
  await register(page);
  await startExam(page);
  page.t.finish();
  await drain();
  assert.equal(page.sent('submitResult').length, 1);
  assert.equal([...page.local.entries.keys()].filter(k => k.startsWith('pendingResult_')).length, 0);
  assert.equal(page.t.hasPending(), false);
  assert.match(page.el('submitStatusBanner').innerHTML, /התקבלה/);
});

// ===================== 4. anti-cheat =====================
test('D9: the first hidden event on a desktop starts a 2 s grace instead of disqualifying', async () => {
  const page = completePage();
  await register(page);
  await startExam(page);
  page.setVisibility('hidden');
  await page.timer.advance(1500);
  assert.equal(page.t.state().dq, false, 'a screen lock or a forced minimise is not yet a disqualification');
  page.setVisibility('visible');
  await drain();
  assert.equal(page.t.state().dq, false, 'returning within the grace keeps the exam');
  assert.equal(page.t.state().screen, 'screenExam');
  assert.equal(page.beacons.filter(b => b.action === 'disqualify').length, 0);
});

test('D9: staying away past the grace still disqualifies, and the second event is immediate', async () => {
  const page = completePage();
  await register(page);
  await startExam(page);
  page.setVisibility('hidden');
  await page.timer.advance(2100);
  assert.equal(page.t.state().dq, true, 'the grace is a grace, not an amnesty');
  assert.equal(page.beacons.filter(b => b.action === 'disqualify').length, 1);

  const second = completePage();
  await register(second); await startExam(second);
  second.setVisibility('hidden'); await second.timer.advance(1000); second.setVisibility('visible'); await drain();
  assert.equal(second.t.state().warnings, 1, 'the first return is counted even though no banner is shown');
  second.setVisibility('hidden'); await drain();
  assert.equal(second.t.state().dq, true, 'no second chance on a desktop');
});

test('D9: a phone keeps exactly its three warned chances', async () => {
  const page = completePage({ userAgent: 'Mozilla/5.0 (iPhone; CPU iPhone OS 17_0 like Mac OS X) AppleWebKit Mobile/15E148', touchPoints: 5 });
  await register(page);
  await startExam(page);
  for (let i = 0; i < 3; i++) {
    page.setVisibility('hidden');
    await page.timer.advance(3000);                       // away 3 s: past the 2 s ignore, inside the 10 s grace
    page.setVisibility('visible'); await drain();
    assert.equal(page.t.state().dq, false, 'chance ' + (i + 1) + ' of 3');
    assert.equal(page.t.state().warnings, i + 1);
  }
  page.setVisibility('hidden'); await drain();
  assert.equal(page.t.state().dq, true, 'the fourth switch disqualifies, exactly as before');
});

test('D10: a disqualification costs ONE execution — the beacon, with the fetch only as a fallback', async () => {
  const page = completePage();
  await register(page);
  await startExam(page);
  const before = page.requests.length;
  page.setVisibility('hidden');
  await page.timer.advance(2100);
  assert.equal(page.beacons.filter(b => b.action === 'disqualify').length, 1);
  assert.equal(page.requests.slice(before).filter(r => r.action === 'disqualify').length, 0, 'no duplicate fetch');

  const noBeacon = completePage();
  noBeacon.ctx.navigator.sendBeacon = null;
  await register(noBeacon); await startExam(noBeacon);
  noBeacon.setVisibility('hidden');
  await noBeacon.timer.advance(2100);
  assert.equal(noBeacon.sent('disqualify').length, 1, 'a browser without sendBeacon still reports it');
});

test('D14: an examinee whose row is already in_exam is told, not left polling forever', async () => {
  const page = completePage({ reply: r => (r.action === 'checkApproval' || r.kind === 'approval')
    ? { status: 'ok', approval: 'in_exam', audioMode: 'off' } : undefined });
  await register(page);
  assert.match(page.el('rejectedMsg').textContent, /המבחן שלך כבר התחיל/);
  const polls = page.sent('checkApproval').length;
  await page.timer.advance(120000);
  assert.equal(page.sent('checkApproval').length, polls, 'the chain stopped');
});

test('D13: the timer-expiry extension check waits the full poll deadline', async () => {
  let asked = 0;
  const page = completePage({ reply(request) {
    if (request.action !== 'getExamStatus') return undefined;
    asked++;
    return asked === 1 ? undefined : { __hang: true };
  } });
  await register(page);
  await startExam(page);
  const src = examinee.replace(/\r/g, '');
  assert.match(section(src, '  function onTimerExpired()', '  function updateTimerDisplay()'), /statusPollCall\(\)/,
    'the 30 s default would auto-submit an exam whose extension was already granted');
});

test('extension: minutes granted mid-exam extend the deadline exactly once', async () => {
  let extra = 0;
  const page = completePage({ reply: r => r.action === 'getExamStatus' || r.kind === 'status'
    ? { status: 'ok', examStatus: 'in_exam', extraMinutes: extra } : undefined });
  await register(page);
  await startExam(page);
  const deadline = page.t.state().deadline;
  extra = 10;
  await page.timer.advance(30000);
  assert.equal(page.t.state().deadline, deadline + 10 * 60 * 1000);
  await page.timer.advance(60000);
  assert.equal(page.t.state().deadline, deadline + 10 * 60 * 1000, 'the same grant is not applied twice');
});

test('DQ: an examiner-initiated disqualification reaches the examinee through the status poll', async () => {
  let status = 'in_exam';
  const page = completePage({ reply: r => (r.action === 'getExamStatus' || r.kind === 'status')
    ? { status: 'ok', examStatus: status, extraMinutes: 0 } : undefined });
  await register(page);
  await startExam(page);
  status = 'disqualified';
  await page.timer.advance(30000);
  assert.equal(page.t.state().dq, true, 'the exam runs locally — without this poll the examinee kept answering');
  assert.match(page.el('examArea').innerHTML, /נפסל/);
  assert.equal(page.beacons.filter(b => b.action === 'disqualify').length, 0, 'no self-DQ: the server already knows');
});

test('DQ: an overturned disqualification resumes the exam with its questions and its clock', async () => {
  // 'approved' is also what the row says while the examiner is still deciding:
  // the overturn poll waits for 'in_exam' and does nothing until then.
  let approval = 'approved';
  const page = completePage({ reply(r) {
    if (r.action === 'getExamStatus' || r.kind === 'status') return { status: 'ok', examStatus: 'disqualified', extraMinutes: 0 };
    if (r.action === 'checkApproval' || r.kind === 'approval') return { status: 'ok', approval: approval, audioMode: 'off' };
    return undefined;
  } });
  await register(page);
  await startExam(page);
  page.t.answerCurrent(1);
  const answered = plain(page.t.state().answers[0]);
  await page.timer.advance(30000);              // the examiner disqualifies
  assert.equal(page.t.state().dq, true);
  assert.equal(page.t.state().inProgress, false, 'the clock is stopped while the examiner decides');

  approval = 'in_exam';                          // and then overturns it
  await page.timer.advance(10000);
  await drain();
  const state = page.t.state();
  assert.equal(state.inProgress, true, 'the exam resumes');
  assert.equal(state.questions.length, TOTAL);
  assert.deepEqual(plain(state.answers[0]), answered, 'with the answers it had');
  assert.equal(page.el('examArea').querySelector('.q-text').textContent, BANK_FILES.he[0].t, 'and its texts, from the bank');
});

// ===================== 5. images =====================
test('D15: the local copy is the only source tried first; the proxy and gov.il are fallbacks', async () => {
  const page = completePage();
  await register(page);
  await startExam(page);
  const sources = page.t.images('images/TQ_PIC_5.jpg');
  assert.equal(sources[0], 'images/TQ_PIC_5.jpg');
  assert.ok(sources.length > 1, 'there is still a fallback chain');
  assert.ok(sources.slice(1).every(u => /^https:\/\//.test(u)));
  assert.ok(sources.some(u => u.includes('image-proxy')));
  assert.ok(sources.some(u => u.includes('gov.il')));
  assert.equal(sources.filter(u => u.includes('gov.il')).length, 2, 'both gov.il folders, since the bank stores only the file name');
});

// ===================== 6. gateway =====================
test('gateway: both polls go to the Worker when the session names one', async () => {
  const page = completePage({ gateway: 'https://gw.example/' });
  await register(page);
  await startExam(page);
  await page.timer.advance(20000);
  const polls = page.requests.filter(r => String(r.__url).includes('/v1/poll'));
  assert.ok(polls.length >= 2, 'the approval and the status poll both went through the gateway');
  assert.ok(polls.some(p => p.kind === 'approval'));
  assert.ok(polls.some(p => p.kind === 'status'));
  assert.ok(polls.every(p => p.sessionCode === 'ABC12345' && p.idNumber === '123456789' && p.examineeToken === 'tok-1'));
  assert.equal(page.sent('checkApproval').length, 0, 'nothing reached Apps Script');
});

test('gateway: upstream_unavailable slows the poll down but never sends the fleet at Apps Script', async () => {
  const page = completePage({ gateway: 'https://gw.example/',
    reply: r => String(r.__url).includes('/v1/poll') ? { status: 'error', code: 'upstream_unavailable', retryable: true } : undefined });
  page.el('sessionCodeInput').value = 'ABC12345';
  page.el('codeSubmitBtn').click(); await drain();
  for (const [id, value] of [['idNumber', '123456789'], ['firstName', 'א'], ['lastName', 'ב'], ['phoneNumber', '0501234567']]) page.el(id).value = value;
  page.el('registerBtn').click(); await drain();
  await page.timer.advance(120000);
  assert.equal(page.sent('checkApproval').length, 0, 'the gateway exists to shield Google exactly when it is busy');
  const polls = page.requests.filter(r => String(r.__url).includes('/v1/poll')).length;
  assert.ok(polls < 20, 'and the client backs off instead of hammering: ' + polls);
});

test('gateway: three transport failures of the Worker itself fall back to the direct call', async () => {
  let gatewayDown = true;
  const page = completePage({ gateway: 'https://gw.example/',
    reply: r => (String(r.__url).includes('/v1/poll') && gatewayDown) ? { __status: 502, __raw: 'bad gateway' } : undefined });
  page.el('sessionCodeInput').value = 'ABC12345';
  page.el('codeSubmitBtn').click(); await drain();
  for (const [id, value] of [['idNumber', '123456789'], ['firstName', 'א'], ['lastName', 'ב'], ['phoneNumber', '0501234567']]) page.el(id).value = value;
  page.el('registerBtn').click(); await drain();
  await page.timer.advance(60000);
  assert.ok(page.sent('checkApproval').length >= 1, 'the examinee is not stranded by a broken Worker');
});

// ===================== 7. restore =====================
test('D20: a state left by another examinee is never adopted silently', async () => {
  const local = memoryStore();
  local.setItem('ext_examinee_state_123456789', JSON.stringify({
    sessionCode: 'ABC12345', sessionData: { site: 'בדיקת נתונים', license: 'B', language: 'he' },
    examineeData: { idNumber: '123456789', fullName: 'ישראל ישראלי', license: 'B', language: 'he' },
    examineeToken: 'tok-old', screen: 'screenInstructions', savedAt: 1
  }));
  const page = completePage({ local });
  await drain();
  const card = page.el('restoreConfirm');
  assert.ok(card, 'a fresh tab asks before it becomes somebody else');
  assert.match(card.innerHTML, /ישראל ישראלי/);
  assert.match(card.innerHTML, /\*\*\*\*6789/, 'the id is masked');
  assert.equal(page.t.state().id, '', 'nothing is restored while the question is open');
  page.el('restoreNo').click();
  await drain();
  assert.equal(page.t.state().id, '', 'saying no starts a new examinee');
  assert.equal(page.t.state().screen, 'screenIdForm', 'on the same session');
  assert.equal(local.getItem('ext_examinee_state_123456789'), null, 'and the old identity is gone from the device');
});

test('D20: saying "yes, it is me" restores the waiting screen as before', async () => {
  const local = memoryStore();
  local.setItem('ext_examinee_state_123456789', JSON.stringify({
    sessionCode: 'ABC12345', sessionData: { site: 'בדיקת נתונים', license: 'B', language: 'he' },
    examineeData: { idNumber: '123456789', fullName: 'ישראל ישראלי', license: 'B', language: 'he' },
    examineeToken: 'tok-old', screen: 'screenInstructions', savedAt: 1
  }));
  const page = completePage({ local, reply: r => (r.action === 'checkApproval') ? { status: 'ok', approval: 'waiting', audioMode: 'off' } : undefined });
  await drain();
  page.el('restoreYes').click();
  await drain(); await page.timer.advance(50); await drain();
  assert.equal(page.t.state().id, '123456789');
  assert.equal(page.t.state().token, 'tok-old');
  assert.equal(page.t.state().screen, 'screenInstructions');
});

test('D20: this tab\'s own state is restored without a question', async () => {
  const session = memoryStore();
  session.setItem('ext_examinee_state', JSON.stringify({
    sessionCode: 'ABC12345', sessionData: { site: 'בדיקת נתונים', license: 'B', language: 'he' },
    examineeData: { idNumber: '123456789', fullName: 'ישראל ישראלי', license: 'B', language: 'he' },
    examineeToken: 'tok-old', screen: 'screenInstructions', savedAt: 1
  }));
  const page = completePage({ session, reply: r => (r.action === 'checkApproval') ? { status: 'ok', approval: 'waiting', audioMode: 'off' } : undefined });
  await drain(); await page.timer.advance(50); await drain();
  assert.ok(!page.el('restoreConfirm'), 'the tab that wrote the state owns it');
  assert.equal(page.t.state().id, '123456789');
});

test('resume: a reload during the exam rebuilds the questions from the bank and keeps the answers', async () => {
  const local = memoryStore(), session = memoryStore();
  const first = completePage({ local, session });
  await register(first);
  await startExam(first);
  first.t.answerCurrent(3);
  const savedAnswer = plain(first.t.state().answers[0]);
  await first.timer.advance(5000);

  const second = completePage({ local, session });
  await drain(); await second.timer.advance(100); await drain();
  const state = second.t.state();
  assert.equal(state.inProgress, true, 'the exam resumes');
  assert.equal(state.questions.length, TOTAL);
  assert.deepEqual(plain(state.answers[0]), savedAnswer, 'with the answer exactly as it was displayed');
  assert.equal(second.el('examArea').querySelector('.q-text').textContent, BANK_FILES.he[0].t);
});

// ===================== 8. update check =====================
test('update check: a real deploy never reloads a registered examinee, and reaches an idle page later', async () => {
  let hash = 'hash-1';
  const page = completePage({ reply: r => String(r.__url).includes('version.json') ? { build: 'v', pages: { 'examinee.html': hash } } : undefined });
  await register(page);
  hash = 'hash-2';
  await page.timer.advance(10 * 60 * 1000);
  assert.equal(page.reloads, 0, 'the 16/09 shape: a waiting examinee must never be bounced to the login screen');
  assert.equal(page.t.state().screen, 'screenInstructions');
  // the examinee leaves; the page is genuinely idle
  page.ctx.resetForNextExaminee();
  await page.timer.advance(120000);
  assert.equal(page.reloads, 1, 'and only then does the page update itself');
});

test('update check: a page mid-exam is never reloaded', async () => {
  let hash = 'hash-1';
  const page = completePage({ reply: r => String(r.__url).includes('version.json') ? { build: 'v', pages: { 'examinee.html': hash } } : undefined });
  await register(page);
  await startExam(page);
  hash = 'hash-9';
  await page.timer.advance(30 * 60 * 1000);
  assert.equal(page.reloads, 0);
  assert.equal(page.t.state().inProgress, true);
});

// ===================== 9. source-level guarantees =====================
test('source: every inline script parses, and the dead legacy paths are gone', () => {
  const scripts = [...examinee.matchAll(/<script\b[^>]*>([\s\S]*?)<\/script>/gi)].map(m => m[1]).filter(c => c.trim());
  assert.equal(scripts.length, 2, 'the page is still one inline closure plus the service-worker registration');
  for (const code of scripts) new Function(code);          // parses standalone
  for (const gone of ['_examTranslations', 'getQuestionsByIds', 'TRANSLATION_VARS', 'getTransDict', 'hasTranslation',
    'getRuQuestion', '_legacyGetFilteredQuestions_DEPRECATED', '_legacyBuildExam_DEPRECATED', 'getQCorrectIndex',
    'decodeCI', 'registerExamQuestions', 'getExamQuestions', 'markExamStarted', 'submitWrongAnswers',
    'loadTranslationFile', 'dqServerConfirmed', 'PASS_SCORE']) {
    assert.ok(!new RegExp('\\b' + gone + '\\b').test(examinee.replace(/markExamStarted is gone[^\n]*/g, '')), gone + ' is gone');
  }
  assert.ok(examinee.includes('<script src="shared/transport.js"></script>'));
  assert.ok(examinee.includes('<script src="shared/bank.js"></script>'));
});

test('source: the page owns no transport of its own any more', () => {
  const src = examinee.replace(/\r/g, '');
  assert.ok(!/function fetchJsonWithTimeout/.test(src), 'the bounded fetch is shared');
  assert.ok(!/function noteTransport/.test(src));
  assert.ok(!/function pacePoll/.test(src));
  assert.ok(!/function jitterMs/.test(src));
  assert.ok(!/setInterval\(approvalPollTick/.test(src));
  assert.equal((src.match(/ExamTransport\.createPollLoop/g) || []).length, 3, 'approval, exam status and DQ overturn');
  assert.ok(/ExamTransport\.createUpdateCheck/.test(src));
  assert.ok(/ExamTransport\.createFailover/.test(src));
  assert.ok(/ExamTransport\.drainLog\(\)/.test(src));
});

test('source: the service worker precaches the shared layers and matches the bank without its query', () => {
  const sw = fs.readFileSync(path.join(app, 'sw-examinee.js'), 'utf8');
  new vm.Script(sw, { filename: 'sw-examinee.js' });
  for (const asset of ['./shared/transport.js', './shared/bank.js', './bank/manifest.json']) assert.ok(sw.includes(asset), asset);
  assert.match(sw, /ignoreSearch: true/);
  assert.match(sw, /req\.method !== 'GET'/, 'GET only — a HEAD can never be cached');
  const cacheLine = sw.split('\n').filter(l => /^var CACHE = '/.test(l));
  assert.equal(cacheLine.length, 1, 'the build tool rewrites exactly this line');
  assert.match(cacheLine[0], /^var CACHE = '[^']*';$/);
});
