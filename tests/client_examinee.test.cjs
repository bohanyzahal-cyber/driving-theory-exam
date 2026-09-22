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
    entries, reject: () => false, rejectBig: 0,
    get length() { return entries.size; }, key: i => [...entries.keys()][i] ?? null,
    getItem: k => entries.get(String(k)) ?? null,
    setItem(k, v) {
      // rejectBig models the real quota: the SAME key is accepted once it is small
      // enough, which is what the drop-the-texts fallback depends on.
      if (store.reject(String(k)) || (store.rejectBig && String(v).length > store.rejectBig)) {
        throw Object.assign(new Error('quota'), { name: 'QuotaExceededError' });
      }
      entries.set(String(k), String(v));
    },
    removeItem: k => entries.delete(String(k))
  };
  return store;
}

// ---------- the synthetic bank ----------
// The texts are not on this site any more: the Worker serves the 30 ids the
// grant names, in every language at once. BANK_FILES is still the source of the
// texts, so every rendering assertion below means exactly what it meant before.
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
const EXAM_BANK = { url: 'https://gw.test', grant: 'g1' };

/** One /v1/bank record: the question in every language `langs` lists. */
function bankRecord(id, langs) {
  const l = {};
  for (const lang of langs) {
    const entry = BANK_FILES[lang][id - 1];
    if (entry) l[lang] = { t: entry.t, a: entry.a, i: entry.i };
  }
  return { id, l };
}
function bankAnswer(ids, langs = LANGS) {
  return { status: 'ok', build: 'bank-test', missing: [], questions: ids.map(id => bankRecord(id, langs)) };
}

const SERVER_QUESTIONS = Array.from({ length: TOTAL }, (_, i) => ({
  id: i + 1, order: [[2, 0, 3, 1], [1, 3, 0, 2], [0, 1, 2, 3], [3, 2, 1, 0]][i % 4], topic: 'חוק'
}));
const ISSUED_IDS = SERVER_QUESTIONS.map(q => q.id);

// ---------- the whole page ----------
// gateway: the Worker every session names. It is the page's ONLY poll route, so
// a session without one is a misconfigured deployment — pass gateway: '' to test
// exactly that, and nothing else.
function completePage({ local = memoryStore(), session = memoryStore(), reply, gateway = 'https://gw.example/', onReload, userAgent = 'Synthetic desktop', touchPoints = 0, search = '' } = {}) {
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
    location: { search: search, pathname: '/examinee.html', reload() { reloads++; if (onReload) onReload(); } },
    fetch(url, opts = {}) {
      const isPost = opts && opts.method === 'POST';
      // A POST with NO body is a real shape on this page since r31: the nudge to
      // the Worker (§13.5) says everything it has to say in the query string.
      const request = (isPost && opts.body !== undefined)
        ? JSON.parse(opts.body)
        : Object.fromEntries(new URL(url, 'https://synthetic.test/').searchParams);
      request.__url = String(url);
      request.__method = isPost ? 'POST' : 'GET';
      request.__keepalive = !!(opts && opts.keepalive);
      request.__mode = (opts && opts.mode) || '';
      request.__at = timer.now;      // when it left the device, for cadence assertions
      const data = answer(request);
      if (data && data.__network) return Promise.reject(new TypeError('Failed to fetch'));
      if (data && data.__hang) return new Promise(() => {});
      // the knobs are the harness's, never part of the body the page parses
      const body = data && data.__raw !== undefined ? data.__raw : JSON.stringify(data, (k, v) => k.slice(0, 2) === '__' ? undefined : v);
      const response = { ok: !(data && data.__status >= 400), status: (data && data.__status) || 200,
        text: () => Promise.resolve(body), json: () => Promise.resolve(JSON.parse(body)) };
      // __delay models a request the server HOLDS (a long poll) before answering
      if (data && data.__delay) return new Promise(resolve => timer.set(() => resolve(response), data.__delay));
      return Promise.resolve(response);
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
    retryDelay: submitRetryDelayMs, hasPending: hasAnyPendingResult,
    nudge: nudgeGatewayAfterWrite, degraded: function() { return ExamTransport.isBackendDegraded(); },
    // sendCancelDQToServer has no live caller in the page today (the "returned
    // within the grace" path clears the timer locally); it is still the one
    // place that sends cancelDisqualify, so its push is tested from here.
    cancelDQ: sendCancelDQToServer
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
  if (url.indexOf('/v1/bank') !== -1) return bankAnswer(ISSUED_IDS);
  if (url.indexOf('version.json') !== -1) return { build: 'v1', pages: { 'examinee.html': 'hash-1' } };
  switch (request.action) {
    case 'getSessionInfo':
      return { status: 'ok', build: 'r25', session: { site: 'בדיקת נתונים', license: 'B', language: 'he', audioMode: 'off',
        examinerName: 'בוחן', classroom: '1', sites: ['בדיקת נתונים'], gateway: { url: gateway } } };
    case 'registerExaminee': return { status: 'ok', examineeToken: 'tok-1' };
    // No checkApproval / getExamStatus: the page cannot reach Apps Script with
    // either one any more. A request carrying them would be a regression, and
    // the assertions below name it.
    case 'startExam': return { status: 'ok', build: 'r25', examMinutes: 40, extraMinutes: 0, audioMode: 'off',
      language: request.language, license: request.license, registeredAt: '', questions: SERVER_QUESTIONS,
      bank: EXAM_BANK };
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
test('start: ONE startExam call builds the 30 questions, in the server order, with texts from the grant', async () => {
  const page = completePage();
  await register(page);
  await startExam(page);
  const state = page.t.state();
  assert.equal(state.inProgress, true);
  assert.equal(state.screen, 'screenExam');
  assert.equal(state.questions.length, TOTAL);
  const apiCalls = page.requests.filter(r => r.action).map(r => r.action);
  assert.deepEqual(apiCalls.slice(0, 3), ['getSessionInfo', 'registerExaminee', 'startExam'],
    'three Apps Script calls take an examinee from the code screen into the exam; the approval wait is the Worker\'s');
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
    return { status: 'ok', examMinutes: 40, extraMinutes: 0, audioMode: 'off', questions, bank: EXAM_BANK };
  } });
  await register(page);
  await startExam(page);
  assert.deepEqual(plain(page.t.state().orders[0].order), [0, 1, 2, 3]);
});

test('start: examMinutes and a pre-granted extension both reach the clock', async () => {
  const page = completePage({ reply: r => r.action === 'startExam'
    ? { status: 'ok', examMinutes: 50, extraMinutes: 10, audioMode: 'off', questions: SERVER_QUESTIONS, bank: EXAM_BANK } : undefined });
  await register(page);
  await startExam(page);
  const state = page.t.state();
  assert.equal(state.timeMinutes, 50);
  assert.ok(Math.abs(state.deadline - page.timer.now - 60 * 60 * 1000) < 1000, '50 authorised + 10 granted');
});

// ===================== 1b. the question texts =====================
const bankCalls = page => page.requests.filter(r => String(r.__url).includes('/v1/bank'));

test('bank: nothing is fetched while the examinee waits — the grant does not exist yet', async () => {
  const page = completePage();
  await register(page);
  await page.timer.advance(120000);
  assert.equal(bankCalls(page).length, 0, 'registration and the whole approval wait cost the bank Worker nothing');
  assert.equal(page.t.state().screen, 'screenInstructions');
  await startExam(page);
  assert.equal(bankCalls(page).length, 1, 'ONE request, right after startExam, carrying the grant it just issued');
  assert.equal(bankCalls(page)[0].grant, 'g1');
});

test('bank: the texts the examinee reads are the ones the grant bought', async () => {
  const page = completePage();
  await register(page);
  await startExam(page);
  assert.equal(page.el('examArea').querySelector('.q-text').textContent, BANK_FILES.he[0].t);
  page.t.goTo(7);
  assert.equal(page.el('examArea').querySelector('.q-text').textContent, BANK_FILES.he[7].t);
  assert.deepEqual(page.el('examArea').querySelectorAll('.ans-text').map(n => n.textContent),
    SERVER_QUESTIONS[7].order.map(i => BANK_FILES.he[7].a[i]));
});

test('bank: a server with no gateway configured says so instead of offering a pointless cooldown', async () => {
  const page = completePage({ reply: r => r.action === 'startExam'
    ? { status: 'ok', examMinutes: 40, extraMinutes: 0, audioMode: 'off', questions: SERVER_QUESTIONS } : undefined });
  await register(page);
  await startExam(page);
  assert.equal(page.t.state().inProgress, false, 'an exam without texts is not an exam');
  assert.equal(bankCalls(page).length, 0, 'there is nothing to ask and no grant to ask with');
  const html = page.el('examArea').innerHTML;
  assert.match(html, /מאגר השאלות אינו מוגדר בשרת/);
  assert.ok(!/השרת עמוס/.test(html), 'this is a missing server setting, not a busy server');
  assert.equal(page.el('retryExamStartBtn').disabled, false, 'no cooldown: re-clicking cannot fix it, but it is not blocked either');
});

test('bank: a Worker that will not answer is retried on the ladder, then the examinee gets a retry', async () => {
  let down = true;
  const page = completePage({ reply: r => (String(r.__url).includes('/v1/bank') && down)
    ? { __status: 503, __raw: '{"status":"error","code":"bank_unavailable"}' } : undefined });
  await register(page);
  page.el('airplaneCheckbox').checked = true;
  page.el('airplaneCheckbox').fire('change');
  page.el('startExamBtn').click();
  await drain();
  await page.timer.advance(6000);                       // the 1.5 s + 3 s ladder inside bank.js
  assert.equal(bankCalls(page).length, 3, 'three attempts and then it stops — a ladder, not a storm');
  assert.equal(page.t.state().inProgress, false);
  assert.match(page.el('examArea').innerHTML, /לא הצלחנו לטעון את השאלות/);
  assert.equal(page.sent('startExam').length, 1, 'the row was already written; the retry does not write a second one');

  down = false;
  await page.timer.advance(5000);
  assert.equal(page.el('retryExamStartBtn').disabled, false);
  page.el('retryExamStartBtn').click();
  await drain(); await page.timer.advance(100); await drain();
  assert.equal(page.t.state().inProgress, true, 'and the retry starts the exam');
  assert.deepEqual(plain(page.t.state().questions).map(q => q.id), ISSUED_IDS, 'with the same 30 ids');
});

test('bank: an id no language could serve is a failed load, never a blank question', async () => {
  const page = completePage({ reply(r) {
    if (!String(r.__url).includes('/v1/bank')) return undefined;
    const body = bankAnswer(ISSUED_IDS);
    body.questions = body.questions.slice(1);
    body.missing = [ISSUED_IDS[0]];
    return body;
  } });
  await register(page);
  page.el('airplaneCheckbox').checked = true;
  page.el('airplaneCheckbox').fire('change');
  page.el('startExamBtn').click();
  await drain(); await page.timer.advance(6000);
  assert.equal(page.t.state().inProgress, false);
  assert.match(page.el('examArea').innerHTML, /לא הצלחנו לטעון את השאלות/);
});

// ===================== 2. language =====================
test('language: a switch is local — no request at all — and the answers follow the new language', async () => {
  const page = completePage();
  await register(page);
  await startExam(page);
  const before = page.requests.length;
  for (const lang of ['ru', 'en', 'ar', 'fr', 'es', 'am', 'he']) { page.t.switchLang(lang); await drain(); }
  await page.timer.advance(50);
  assert.equal(page.requests.filter(r => String(r.__url).includes('/v1/bank')).length, 1,
    'ONE request bought all seven languages; a switch never touches the network again');
  page.t.switchLang('ru'); await drain();
  assert.equal(page.requests.length, before, 'and nothing else was sent either');
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

test('language: a language the Worker had no text for keeps the exam running', async () => {
  const served = LANGS.filter(l => l !== 'am');
  const page = completePage({ reply: r => String(r.__url).includes('/v1/bank') ? bankAnswer(ISSUED_IDS, served) : undefined });
  await register(page);
  await startExam(page);
  page.t.switchLang('am');
  await drain(); await page.timer.advance(100); await drain();
  assert.equal(page.t.state().lang, 'he', 'the exam stays in the language it was in');
  assert.match(page.el('examNotice').textContent, /אינן זמינות בשפה/);
  assert.equal(page.t.state().inProgress, true);
  assert.equal(page.requests.filter(r => String(r.__url).includes('/v1/bank')).length, 1, 'and it did not go looking for one');
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
  const page = completePage({ reply: r => r.kind === 'approval'
    ? { status: 'ok', approval: 'in_exam', audioMode: 'off' } : undefined });
  await register(page);
  assert.match(page.el('rejectedMsg').textContent, /המבחן שלך כבר התחיל/);
  const polls = pollsOf(page, 'approval').length;
  await page.timer.advance(120000);
  assert.equal(pollsOf(page, 'approval').length, polls, 'the chain stopped');
});

// The examiner's decision about the registration itself. Until r30 the server
// answered these as 'לא נמצא רישום' and the examinee watched a 'שגיאת שרת'
// banner with nothing to click.
test('rejected: the examinee is told at once, polling stops, and one button takes them back to registration', async () => {
  let approval = 'waiting';
  const page = completePage({ reply: r => r.kind === 'approval'
    ? { status: 'ok', approval: approval, audioMode: 'off' } : undefined });
  await register(page);
  assert.equal(page.t.state().screen, 'screenInstructions');

  approval = 'rejected';
  await page.timer.advance(5000);

  assert.equal(page.el('rejectedMsg').textContent, 'הבוחן דחה את הכניסה שלך. פנה לבוחן.');
  assert.equal(page.el('rejectedMsg').style.display, 'block');
  assert.equal(page.el('waitingPhase').style.display, 'none');
  const back = page.el('backToRegisterBtn');
  assert.equal(back.textContent, 'חזרה להרשמה');
  assert.equal(back.style.display, 'block');
  assert.ok(!page.el('approvalError') || page.el('approvalError').style.display !== 'block',
    'a decision is not a server error');

  const stopped = pollsOf(page, 'approval').length;
  await page.timer.advance(120000);
  assert.equal(pollsOf(page, 'approval').length, stopped, 'the chain stopped');

  back.click();
  await drain();
  assert.equal(page.t.state().screen, 'screenCode');
  assert.equal(page.local.getItem('ext_examinee_state_123456789'), null, 'the dead registration is off the device');
  assert.equal(page.session.getItem('ext_examinee_state'), null);
  assert.equal(back.style.display, 'none', 'and the decision screen is cleared behind them');
  assert.equal(page.el('rejectedMsg').style.display, 'none');
  assert.equal(page.el('waitingPhase').style.display, 'block');
  assert.equal(pollsOf(page, 'approval').length, stopped, 'leaving started no new chain');
});

test('cancelled: the reset is named in the examinee\'s own language, with the same way back', async () => {
  let approval = 'waiting';
  const page = completePage({ reply: r => r.kind === 'approval'
    ? { status: 'ok', approval: approval, audioMode: 'off' } : undefined });
  await register(page, { language: 'ru' });

  approval = 'cancelled';
  await page.timer.advance(5000);

  assert.equal(page.el('rejectedMsg').textContent, 'Регистрация отменена экзаменатором. Зарегистрируйтесь заново.');
  assert.equal(page.el('rejectedMsg').style.display, 'block');
  assert.equal(page.el('backToRegisterBtn').textContent, 'Вернуться к регистрации');
  const stopped = pollsOf(page, 'approval').length;
  await page.timer.advance(120000);
  assert.equal(pollsOf(page, 'approval').length, stopped, 'the chain stopped');

  page.el('backToRegisterBtn').click();
  await drain();
  assert.equal(page.t.state().screen, 'screenCode');
  assert.equal(page.local.getItem('ext_examinee_state_123456789'), null);
});

test('D13: the timer-expiry extension check waits the full poll deadline', async () => {
  const page = completePage();
  await register(page);
  await startExam(page);
  const src = examinee.replace(/\r/g, '');
  // statusPollCall keeps the 60 s poll deadline (the 30 s API default would
  // auto-submit an exam whose extension was already granted), and noWait keeps
  // the Worker from HOLDING this one answer: the examinee is watching a timer
  // that just hit zero, so it must come back at once, not in up to 25 s.
  assert.match(section(src, '  function onTimerExpired()', '  function updateTimerDisplay()'),
    /statusPollCall\(\{ noWait: true \}\)/,
    'the 30 s default would auto-submit an exam whose extension was already granted');
});

test('extension: minutes granted mid-exam extend the deadline exactly once', async () => {
  let extra = 0;
  const page = completePage({ reply: r => r.kind === 'status' ? { status: 'ok', examStatus: 'in_exam', extraMinutes: extra } : undefined });
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
  const page = completePage({ reply: r => r.kind === 'status' ? { status: 'ok', examStatus: status, extraMinutes: 0 } : undefined });
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
    if (r.kind === 'status') return { status: 'ok', examStatus: 'disqualified', extraMinutes: 0 };
    if (r.kind === 'approval') return { status: 'ok', approval: approval, audioMode: 'off' };
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

test('DQ: an overturn after a RELOAD resumes from the texts the suspended state kept, with no network', async () => {
  let approval = 'approved';
  const dqReply = r => {
    if (r.kind === 'status') return { status: 'ok', examStatus: 'disqualified', extraMinutes: 0 };
    if (r.kind === 'approval') return { status: 'ok', approval: approval, audioMode: 'off' };
    return undefined;
  };
  const local = memoryStore(), session = memoryStore();
  const first = completePage({ local, session, reply: dqReply });
  await register(first);
  await startExam(first);
  first.t.answerCurrent(1);
  const answered = plain(first.t.state().answers[0]);
  await first.timer.advance(30000);                  // the examiner disqualifies
  assert.equal(first.t.state().dq, true);
  const suspended = JSON.parse(local.getItem('examSuspended_ABC12345_123456789'));
  assert.equal(suspended.records.length, TOTAL, 'the texts were suspended with the exam');

  // the examinee reloads while disqualified, and only then is it overturned
  approval = 'in_exam';
  const second = completePage({ local, session, reply: r => (String(r.__url).includes('/v1/bank') || r.action === 'startExam')
    ? { __network: true } : dqReply(r) });
  await drain(); await second.timer.advance(30000); await drain();
  const state = second.t.state();
  assert.equal(state.inProgress, true, 'the exam resumes');
  assert.deepEqual(plain(state.answers[0]), answered, 'with the answers it had');
  assert.equal(second.el('examArea').querySelector('.q-text').textContent, BANK_FILES.he[0].t, 'and its texts, from the device');
  assert.equal(bankCalls(second).length, 0);
  assert.equal(second.sent('startExam').length, 0);
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
test('gateway: a session with no Worker is a misconfiguration, said at the code screen', async () => {
  // The server refuses to start an exam without GATEWAY_URL, and this page has
  // no second route for its polls. Nobody should get as far as registering.
  const page = completePage({ gateway: '' });
  page.el('sessionCodeInput').value = 'ABC12345';
  page.el('codeSubmitBtn').click();
  await drain();
  await page.timer.advance(120000);
  assert.match(page.el('codeApiError').textContent, /המערכת אינה מוגדרת \(Worker\)/);
  assert.equal(page.el('codeApiError').style.display, 'block');
  assert.equal(page.t.state().screen, 'screenCode', 'the examinee stays where they are');
  assert.equal(page.requests.filter(r => String(r.__url).includes('/v1/poll')).length, 0, 'nothing was polled');
  assert.equal(page.sent('checkApproval').length, 0, 'and Apps Script was not asked to stand in');
  assert.equal(page.sent('registerExaminee').length, 0);
  assert.equal(page.el('codeSubmitBtn').disabled, false, 'the button is usable again once it is fixed');
});

test('gateway: both polls go to the Worker — it is the only route there is', async () => {
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

test('gateway: against an OLD Worker the approval poll falls back to 2 s, then to 3 s', async () => {
  // The FALLBACK cadence, and since r31 (§13.4) that is all it is. This Worker
  // answers without a fingerprint — it has never heard of wait/fp — so there is
  // nothing the next request could be held against and the page paces itself:
  // 2 s through the window in which the examiner is actually walking the room,
  // then 3 s. Against an upgraded Worker the same chain re-arms in 250 ms and
  // spends its time inside a hold instead (the re-arm tests below).
  const page = completePage({ gateway: 'https://gw.example/',
    reply: r => String(r.__url).includes('/v1/poll') ? { status: 'ok', approval: 'waiting', audioMode: 'off' } : undefined });
  await register(page);
  const polls = () => page.requests.filter(r => String(r.__url).includes('/v1/poll') && r.kind === 'approval').length;
  const opening = polls();
  await page.timer.advance(60000);
  const fast = polls() - opening;
  assert.ok(Math.abs(fast - 30) <= 2, 'about one every 2 s in the opening minute: ' + fast);
  await page.timer.advance(70000);                 // now past the 120 s window
  const settled = polls();
  await page.timer.advance(60000);
  const slow = polls() - settled;
  assert.ok(Math.abs(slow - 20) <= 2, 'and one every 3 s afterwards: ' + slow);
  assert.equal(page.sent('checkApproval').length, 0, 'none of it reached Apps Script');
});

test('gateway: against an OLD Worker the in-exam status poll falls back to 6 s', async () => {
  // Same as above: no fingerprint in either answer, so nothing can be held and
  // nothing can be re-armed, and the constant in the page is what paces it.
  const page = completePage({ gateway: 'https://gw.example/', reply(r) {
    if (r.kind === 'approval') return { status: 'ok', approval: 'approved', audioMode: 'off', examMinutes: 40 };
    if (r.kind === 'status') return { status: 'ok', examStatus: 'in_exam', extraMinutes: 0 };
    return undefined;
  } });
  await register(page);
  await startExam(page);
  const polls = () => page.requests.filter(r => String(r.__url).includes('/v1/poll') && r.kind === 'status').length;
  const before = polls();
  await page.timer.advance(60000);
  const n = polls() - before;
  assert.ok(Math.abs(n - 10) <= 2, 'one every 6 s: ' + n);
  assert.equal(page.sent('getExamStatus').length, 0);
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

test('gateway: a Worker outage is retried on the ladder — never answered by sending the room at Apps Script', async () => {
  // Five consecutive real failures (a 502 the page can see, while it is on
  // screen). There is nowhere else to go: the loop keeps asking the Worker, more
  // and more slowly, and it never dies.
  let gatewayDown = true;
  const page = completePage({
    reply: r => (String(r.__url).includes('/v1/poll') && gatewayDown) ? { __status: 502, __raw: 'bad gateway' } : undefined });
  page.el('sessionCodeInput').value = 'ABC12345';
  page.el('codeSubmitBtn').click(); await drain();
  for (const [id, value] of [['idNumber', '123456789'], ['firstName', 'א'], ['lastName', 'ב'], ['phoneNumber', '0501234567']]) page.el(id).value = value;
  page.el('registerBtn').click(); await drain();

  await page.timer.advance(300000);
  const failed = pollsOf(page, 'approval');
  assert.ok(failed.length >= 5, 'five failures and more: ' + failed.length);
  assert.equal(page.sent('checkApproval').length, 0, 'not one of them went to Apps Script');
  assert.equal(page.sent('getExamStatus').length, 0);
  // growing delays, and a ceiling: the loop's own ladder ends at maxMs (20 s)
  // and stays there — a Worker 502 never marks the BACKEND degraded (that would
  // slow the submit path too). Five minutes of outage is a handful of requests
  // per device, not a storm.
  const gap = i => failed[i + 1].__at - failed[i].__at;
  assert.ok(gap(failed.length - 2) > gap(0), 'the retries spread out: ' + gap(0) + ' ms then ' + gap(failed.length - 2) + ' ms');
  assert.ok(gap(failed.length - 2) <= 20000, 'and stop growing at maxMs: ' + gap(failed.length - 2));
  assert.ok(failed.length <= 20, 'a whole outage cost this device ' + failed.length + ' requests');
  assert.match(page.el('approvalError').textContent, /\S/, 'the examinee is told that the line is down');

  gatewayDown = false;                           // the Worker comes back
  await page.timer.advance(120000);
  assert.ok(pollsOf(page, 'approval').length > failed.length, 'the chain was still alive and picked it straight back up');
  assert.equal(page.el('instructionsPhase').style.display, 'block', 'and the approval it had been waiting for lands');
});

// ===================== 6b. long polling =====================
// The Worker holds a poll until THIS examinee's answer changes, up to 25 s. The
// page carries the hold in the URL: `wait` (how long it may hold) and `fp` (the
// fingerprint of the answer this device already has).
const pollsOf = (page, kind) => page.requests.filter(r => String(r.__url).includes('/v1/poll') && r.kind === kind);
const waitingAnswer = fp => ({ status: 'ok', approval: 'waiting', audioMode: 'off', fp: fp, held: 0 });

test('long poll: the approval poll offers a 25 s hold, without a fingerprint the first time', async () => {
  let fp = 'fp-1';
  const page = completePage({ gateway: 'https://gw.example/',
    reply: r => r.kind === 'approval' ? waitingAnswer(fp) : undefined });
  await register(page);
  const first = pollsOf(page, 'approval')[0];
  assert.equal(first.wait, '25', 'every gateway poll offers the Worker the hold');
  assert.equal(first.fp, undefined, 'the first one has nothing to hold against, so it is answered at once');
  await page.timer.advance(2000);
  assert.equal(pollsOf(page, 'approval')[1].fp, 'fp-1', 'from the second on it holds against the answer it already has');
  fp = 'fp-2';                                   // the examiner did something: the answer, and its fingerprint, changed
  await page.timer.advance(8000);
  const last = pollsOf(page, 'approval').pop();
  assert.equal(last.fp, 'fp-2', 'and the next hold is against the NEW answer');
  assert.equal(last.wait, '25');
  assert.equal(page.sent('checkApproval').length, 0, 'none of it reached Apps Script');
});

test('long poll: the in-exam status poll holds the same way', async () => {
  const page = completePage({ gateway: 'https://gw.example/',
    reply: r => r.kind === 'status' ? { status: 'ok', examStatus: 'in_exam', extraMinutes: 0, fp: 'st-1', held: 0 } : undefined });
  await register(page);
  await startExam(page);
  assert.equal(pollsOf(page, 'status')[0].wait, '25');
  assert.equal(pollsOf(page, 'status')[0].fp, undefined);
  await page.timer.advance(12000);
  assert.equal(pollsOf(page, 'status')[1].fp, 'st-1');
  assert.equal(page.sent('getExamStatus').length, 0);
});

test('long poll: an approval that lands inside a held answer is applied exactly as before', async () => {
  let approved = false;
  const current = () => approved
    ? { status: 'ok', approval: 'approved', audioMode: 'off', examMinutes: 40, fp: 'f-approved' }
    : { status: 'ok', approval: 'waiting', audioMode: 'off', fp: 'f-waiting' };
  const page = completePage({ gateway: 'https://gw.example/', reply(r) {
    if (r.kind !== 'approval') return undefined;
    const now = current();
    // the Worker holds only while its answer still matches the fingerprint it was given
    return r.fp === now.fp ? Object.assign({ held: 25000, __delay: 25000 }, now) : Object.assign({ held: 0 }, now);
  } });
  await register(page);
  assert.notEqual(page.el('instructionsPhase').style.display, 'block', 'still waiting');
  await page.timer.advance(2200);                // the un-held first answer is paced as before; then the hold starts
  assert.equal(pollsOf(page, 'approval').filter(r => r.fp).length, 1,
    'one request is holding; nothing else is being sent while it does');
  await page.timer.advance(25500);               // the hold expires unchanged, and the next one starts after the gap
  assert.equal(pollsOf(page, 'approval').filter(r => r.fp).length, 2);
  approved = true;                               // the examiner approves
  await page.timer.advance(26000);               // the pending hold expires, the next request sees the change
  assert.equal(page.el('instructionsPhase').style.display, 'block', 'the examinee is on the instructions screen');
  assert.equal(page.t.state().timeMinutes, 40);
  assert.ok(pollsOf(page, 'approval').length < 10, 'a whole approval wait cost a handful of requests: ' + pollsOf(page, 'approval').length);
});

test('long poll: a dropped hold takes its fingerprint with it, so the next poll is answered at once', async () => {
  let gatewayDown = false;
  const page = completePage({ gateway: 'https://gw.example/', reply(r) {
    if (!String(r.__url).includes('/v1/poll')) return undefined;
    return gatewayDown ? { __network: true } : waitingAnswer('gw-1');
  } });
  await register(page);
  await page.timer.advance(6000);
  assert.ok(pollsOf(page, 'approval').pop().fp === 'gw-1', 'a fingerprint is established');

  page.setVisibility('hidden');                  // the phone locks and the held request dies
  gatewayDown = true;
  await page.timer.advance(30000);
  page.setVisibility('visible');
  gatewayDown = false;
  await page.timer.advance(5000);
  const back = pollsOf(page, 'approval').filter(r => r.__at > 36000);
  assert.ok(back.length >= 1, 'the chain is still alive');
  assert.equal(back[0].fp, undefined, 'and asks without a fingerprint, so the Worker answers immediately');
  assert.equal(back[0].wait, '25', 'while still offering the next hold');
  assert.equal(page.sent('checkApproval').length, 0, 'none of this involved Apps Script');
});

test('long poll: a held request killed by a screen lock is not a broken Worker', async () => {
  let mode = 'ok';
  const page = completePage({ gateway: 'https://gw.example/', reply(r) {
    if (!String(r.__url).includes('/v1/poll')) return undefined;
    if (mode === 'network') return { __network: true };
    if (mode === 'http') return { __status: 502, __raw: 'bad gateway' };
    return waitingAnswer('gw-1');
  } });
  await register(page);
  page.setVisibility('hidden');
  mode = 'network';                              // the phone locked: every held request dies
  await page.timer.advance(180000);
  assert.equal(page.sent('checkApproval').length, 0,
    'a room of locked phones must not be answered by sending all of them at Apps Script');
  const whileHidden = pollsOf(page, 'approval').length;
  assert.ok(whileHidden >= 3, 'the chain kept trying, backing off: ' + whileHidden);
  assert.ok(whileHidden < 40, 'and it did back off: ' + whileHidden);

  // The difference a dropped hold makes is on the SCREEN, not in the routing:
  // both go on asking the Worker, but only a failure the examinee could see is
  // allowed to accuse the server.
  assert.ok(!page.el('approvalError') || page.el('approvalError').style.display !== 'block',
    'nothing was said while the phone was asleep');
  page.setVisibility('visible');
  await page.timer.advance(3000);                // past the 2 s wake grace
  mode = 'http';                                 // now the Worker really is answering 502, on screen
  await page.timer.advance(180000);
  assert.equal(page.sent('checkApproval').length, 0, 'and a real failure is still no reason to call Apps Script');
  assert.equal(page.el('approvalError').style.display, 'block', 'it is a reason to tell the examinee');
});

test('long poll: three screen locks never raise a server error on the waiting screen', async () => {
  let mode = 'ok';
  const page = completePage({ gateway: 'https://gw.example/', reply(r) {
    if (!String(r.__url).includes('/v1/poll')) return undefined;
    return mode === 'network' ? { __network: true } : waitingAnswer('gw-1');
  } });
  await register(page);
  page.setVisibility('hidden');
  mode = 'network';
  await page.timer.advance(120000);
  const banner = page.el('approvalError');
  assert.ok(!banner || banner.style.display !== 'block', 'the examinee sees nothing: their phone was asleep');
  // poll_interrupted, not a failure: the debug line says the wait was cut short
  // and the chain simply asks again.
  assert.match(page.el('approvalDebugResponse').textContent, /ההמתנה נקטעה/);
  assert.ok(pollsOf(page, 'approval').length >= 3, 'while the chain kept asking the Worker');
  assert.equal(page.sent('checkApproval').length, 0);
});

test('long poll: the DQ-overturn wait is held too, starting from the decision as it is now', async () => {
  let approval = 'approved';
  const page = completePage({ gateway: 'https://gw.example/',
    reply: r => r.kind === 'approval'
      ? { status: 'ok', approval: approval, audioMode: 'off', examMinutes: 40, fp: 'a-' + approval, held: 0 } : undefined });
  await register(page);
  await startExam(page);
  const before = pollsOf(page, 'approval').length;
  approval = 'disqualified';
  page.setVisibility('hidden');                  // tab switch during the exam
  await page.timer.advance(2100);                // past the grace: disqualified
  page.setVisibility('visible'); await drain();  // the DQ screen, and the wait for the examiner's decision
  assert.equal(page.t.state().dq, true);
  const dqPolls = pollsOf(page, 'approval').slice(before);
  assert.ok(dqPolls.length >= 1, 'the overturn wait is running');
  assert.equal(dqPolls[0].wait, '25', 'and it is a held poll, so the decision lands in about a second');
  assert.equal(dqPolls[0].fp, undefined, 'starting from the decision as it is right now');
  approval = 'dq_confirmed';
  await page.timer.advance(8000);
  assert.match(page.el('dqWaitingMsg').innerHTML, /הבוחן אישר את הפסילה/);
});

test('the DQ-overturn wait ignores a rejected/cancelled answer instead of acting on it', async () => {
  // Both chains speak to the same route. A decision about a REGISTRATION
  // cannot happen while that registration is mid-exam, so the only requirement
  // here is that one would do nothing at all — no resume, no final DQ screen.
  let approval = 'approved';
  const page = completePage({ reply: r => r.kind === 'approval'
    ? { status: 'ok', approval: approval, audioMode: 'off', examMinutes: 40 } : undefined });
  await register(page);
  await startExam(page);
  approval = 'disqualified';
  page.setVisibility('hidden');
  await page.timer.advance(2100);
  page.setVisibility('visible'); await drain();
  assert.equal(page.t.state().dq, true);

  const before = pollsOf(page, 'approval').length;
  approval = 'cancelled';
  await page.timer.advance(10000);
  assert.ok(pollsOf(page, 'approval').length > before, 'it kept waiting for the examiner');
  assert.equal(page.t.state().screen, 'screenExam', 'nothing moved');
  assert.ok(!/הבוחן אישר את הפסילה/.test(page.el('dqWaitingMsg').innerHTML));
  approval = 'rejected';
  await page.timer.advance(10000);
  assert.equal(page.t.state().screen, 'screenExam');
  assert.equal(page.t.state().inProgress, false, 'and the suspended exam is still suspended');
});

// ===================== 6c. re-arm: "why not 0" (r31, DESIGN §13.4) =====================
// The first answer of a chain, and every answer that CHANGED, comes back
// unheld — and used to be followed by 2-3 s (6 in the exam) before the next
// request went out, which is the request the Worker would have held. So the
// hold began seconds after it could have. Now any answer that carries a
// fingerprint is followed by the 250 ms gap: the chain spends its life inside a
// hold, and the examiner's decision is never more than a gap from the screen.
const gapsOf = polls => polls.slice(1).map((p, i) => p.__at - polls[i].__at);

test('re-arm: the approval chain re-arms 250 ms after a holdable answer, carrying its fingerprint', async () => {
  const page = completePage({ gateway: 'https://gw.example/',
    reply: r => r.kind === 'approval' ? waitingAnswer('fp-1') : undefined });   // held 0, but holdable
  await register(page);
  await page.timer.advance(1000);
  const polls = pollsOf(page, 'approval');
  assert.deepEqual(gapsOf(polls).slice(0, 4), [250, 250, 250, 250],
    'not 2 s, not 3 s: the next request IS the hold, so it goes out now');
  assert.equal(polls[0].fp, undefined, 'the first of a chain still has nothing to hold against');
  assert.equal(polls[1].fp, 'fp-1', 'and every one after it carries the fingerprint it was given');
  assert.equal(polls[1].wait, '25', 'while still offering the hold');
  assert.equal(page.sent('checkApproval').length, 0);
});

test('re-arm: the in-exam status chain re-arms the same way', async () => {
  const page = completePage({ gateway: 'https://gw.example/',
    reply: r => r.kind === 'status' ? { status: 'ok', examStatus: 'in_exam', extraMinutes: 0, fp: 'st-1', held: 0 } : undefined });
  await register(page);
  await startExam(page);
  const from = pollsOf(page, 'status').length;
  await page.timer.advance(1000);
  const polls = pollsOf(page, 'status').slice(from - 1);
  assert.deepEqual(gapsOf(polls).slice(0, 3), [250, 250, 250], 'a mid-exam DQ or extension no longer waits out a 6 s tick');
  assert.equal(polls[1].fp, 'st-1');
  assert.equal(page.sent('getExamStatus').length, 0);
});

test('re-arm: the DQ-overturn wait re-arms too — the examiner\'s decision lands in a gap', async () => {
  let approval = 'approved';
  const page = completePage({ gateway: 'https://gw.example/',
    reply: r => r.kind === 'approval'
      ? { status: 'ok', approval: approval, audioMode: 'off', examMinutes: 40, fp: 'a-' + approval, held: 0 } : undefined });
  await register(page);
  await startExam(page);
  const before = pollsOf(page, 'approval').length;
  approval = 'disqualified';
  page.setVisibility('hidden');
  await page.timer.advance(2100);                // past the grace: disqualified
  page.setVisibility('visible'); await drain();
  assert.equal(page.t.state().dq, true);
  const from = pollsOf(page, 'approval').length;
  await page.timer.advance(1000);
  const dqPolls = pollsOf(page, 'approval').slice(Math.max(before, from - 1));
  assert.deepEqual(gapsOf(dqPolls).slice(0, 3), [250, 250, 250]);
  assert.equal(dqPolls[1].fp, 'a-disqualified', 'held against the decision as it stands');
});

test('re-arm: a STALE answer does not re-arm — the Worker never holds a stale copy', async () => {
  // Re-arming against one would hammer the Worker every 250 ms for as long as
  // Google is unreachable behind it, which is the one moment it must not be.
  const page = completePage({ gateway: 'https://gw.example/',
    reply: r => r.kind === 'approval'
      ? { status: 'ok', approval: 'waiting', audioMode: 'off', fp: 'fp-1', held: 0, stale: true } : undefined });
  await register(page);
  const t0 = pollsOf(page, 'approval')[0].__at;
  await page.timer.advance(6000);
  const polls = pollsOf(page, 'approval');
  assert.deepEqual(polls.slice(0, 4).map(p => p.__at - t0), [0, 2000, 4000, 6000],
    'the fallback cadence, exactly as before r31');
  assert.equal(polls[1].fp, 'fp-1', 'the fingerprint is still offered — the Worker will hold once its copy is fresh');
});

// ===================== 6d. pushing our own writes (r31, DESIGN §13.5) =====================
// Every write the EXAMINEE makes announces itself to the Worker. Until r31 the
// examiner's dashboard learned of them only when the Worker next re-read Google
// and then only on its own next tick; since the Worker stopped reading on a 2 s
// clock (it reads when something announces a change, plus a 45 s safety read —
// KNOWN_ISSUES #35) an unannounced write can sit unseen for the whole 45 s.
// The push carries no decision at all, only proof of who is speaking: the
// examinee's own token, which the Worker matches against the row it holds.
//
// Confirmed writes (registerExaminee, startExam, submitResult) push at once.
// Fire-and-forget writes (the markFinished beacon, disqualify, cancelDisqualify,
// reportWarning) push 2 s later, because a Worker that re-read the sheet before
// the write landed would cache the row exactly as it was.
const nudges = page => page.requests.filter(r => String(r.__url).includes('/v1/invalidate'));
const NUDGE_DELAY = 2000;

/** Every push has the same shape, whatever produced it. */
function assertNudgeShape(push, { token = 'tok-1', id = '123456789', session = 'ABC12345' } = {}) {
  assert.ok(push, 'a push was sent');
  assert.equal(String(push.__url).split('?')[0], 'https://gw.example/v1/invalidate');
  assert.equal(push.__method, 'POST');
  assert.equal(push.__keepalive, true, 'it has to survive the examinee closing the tab behind it');
  assert.equal(push.sessionCode, session);
  assert.equal(push.idNumber, id);
  assert.equal(push.examineeToken, token);
  assert.equal(push.grant, undefined, 'an examinee holds no examiner grant and must never need one');
  assert.equal(push.status, undefined, 'and pushes no decision: only the examiner writes into what examinees read');
  assert.equal(push.extraMinutes, undefined);
}

test('nudge: a registration the server confirmed is pushed at once', async () => {
  // The row exists in ממתינים the moment the answer came back, so the examiner's
  // watch can wake on it immediately — the first poll's own forced read of
  // Google coalesces with the one this drop causes.
  const page = completePage({ gateway: 'https://gw.example/' });
  await register(page);
  assert.equal(page.sent('registerExaminee').length, 1);
  assert.equal(nudges(page).length, 1, 'exactly one push for one registration');
  assertNudgeShape(nudges(page)[0]);
  assert.equal(page.t.state().screen, 'screenInstructions');
});

test('nudge: a registration RESUMED on a retry is pushed with the token the row already had', async () => {
  let tries = 0;
  const page = completePage({ gateway: 'https://gw.example/', reply(r) {
    if (r.action !== 'registerExaminee') return undefined;
    return ++tries === 1 ? { __network: true } : { status: 'ok', examineeToken: 'tok-resumed', resumed: true };
  } });
  await register(page);
  assert.equal(nudges(page).length, 0, 'a registration that never reached the server announces nothing');
  page.el('registerBtn').click();
  await drain(); await page.timer.advance(100); await drain();
  assert.equal(nudges(page).length, 1);
  assertNudgeShape(nudges(page)[0], { token: 'tok-resumed' });
});

test('nudge: a registration the server refused pushes nothing', async () => {
  const page = completePage({ gateway: 'https://gw.example/',
    reply: r => r.action === 'registerExaminee' ? { status: 'error', message: 'הסשן נסגר' } : undefined });
  await register(page);
  await page.timer.advance(10000);
  assert.equal(nudges(page).length, 0, 'nothing was written, so there is nothing to announce');
  assert.equal(page.el('registerError').style.display, 'block');
});

test('nudge: a started exam is pushed at once — the row is in_exam the moment the server answered', async () => {
  const page = completePage({ gateway: 'https://gw.example/' });
  await register(page);
  const afterRegistration = nudges(page).length;
  await startExam(page);
  assert.equal(page.t.state().inProgress, true);
  const pushed = nudges(page).slice(afterRegistration);
  assert.equal(pushed.length, 1, 'one startExam, one push');
  assertNudgeShape(pushed[0]);
  const bank = page.requests.filter(r => String(r.__url).includes('/v1/bank'))[0];
  assert.ok(pushed[0].__at <= bank.__at, 'and it goes out before the texts are even fetched');
});

test('nudge: an exam start the server did not complete pushes nothing', async () => {
  for (const failure of [{ __raw: '<html>Google is having trouble</html>' }, { status: 'error', message: 'טוקן נבחן לא תקין' }]) {
    const page = completePage({ gateway: 'https://gw.example/',
      reply: r => r.action === 'startExam' ? failure : undefined });
    await register(page);
    const afterRegistration = nudges(page).length;
    await startExam(page);
    await page.timer.advance(10000);
    assert.equal(page.t.state().inProgress, false);
    assert.equal(nudges(page).length, afterRegistration, 'no exam, no announcement');
  }
});

test('nudge: a disqualification is pushed 2 s after it is sent — the examiner must see "needs decision"', async () => {
  const page = completePage({ gateway: 'https://gw.example/' });
  await register(page);
  await startExam(page);
  const before = nudges(page).length;
  page.setVisibility('hidden');
  await page.timer.advance(2000);                // the desktop grace expires: the DQ beacon goes out
  assert.equal(page.beacons.filter(b => b.action === 'disqualify').length, 1);
  assert.equal(nudges(page).length, before, 'nothing is pushed while the beacon is still in the air');
  await page.timer.advance(NUDGE_DELAY - 1);
  assert.equal(nudges(page).length, before, 'still not');
  await page.timer.advance(1);
  const pushed = nudges(page).slice(before);
  assert.equal(pushed.length, 1);
  assertNudgeShape(pushed[0]);
});

test('nudge: a disqualification sent while the page is UNLOADING pushes nothing', async () => {
  // The beacon is all that can still leave; a fetch would be killed with the
  // document and a timer 2 s out would never fire. That DQ reaches the examiner
  // on the Worker's safety read.
  const page = completePage({ gateway: 'https://gw.example/' });
  await register(page);
  await startExam(page);
  page.setVisibility('hidden');
  await page.timer.advance(2000);                // a DQ is now confirmed on this device
  await page.timer.advance(NUDGE_DELAY);
  const before = nudges(page).length;
  const beaconsBefore = page.beacons.filter(b => b.action === 'disqualify').length;
  page.dispatch('beforeunload', { preventDefault() {}, returnValue: '' });
  await page.timer.advance(30000);
  assert.equal(page.beacons.filter(b => b.action === 'disqualify').length, beaconsBefore + 1,
    'the same beacon as always still goes out');
  assert.equal(nudges(page).length, before, 'and nothing is scheduled behind a page that is gone');
});

test('nudge: a cancelled disqualification is pushed too, so "needs decision" stops being shown', async () => {
  // sendCancelDQToServer is a beacon like the disqualification it undoes and
  // nothing reads an answer from it, so the push waits the same 2 s: a Worker
  // that re-read before the write landed would cache the row still פסול, and
  // that copy would stand until the safety read — the opposite of the point.
  const page = completePage({ gateway: 'https://gw.example/' });
  await register(page);
  await startExam(page);
  const before = nudges(page).length;
  page.t.cancelDQ();
  await drain();
  assert.equal(page.beacons.filter(b => b.action === 'cancelDisqualify').length, 1);
  assert.equal(nudges(page).length, before);
  await page.timer.advance(NUDGE_DELAY);
  const pushed = nudges(page).slice(before);
  assert.equal(pushed.length, 1);
  assertNudgeShape(pushed[0]);
});

test('nudge: a reported warning is pushed 2 s later — the counter is a column on the dashboard', async () => {
  const page = completePage({ gateway: 'https://gw.example/' });
  await register(page);
  await startExam(page);
  const before = nudges(page).length;
  page.setVisibility('hidden');
  await page.timer.advance(1000);                // back inside the desktop grace: a warning, not a DQ
  page.setVisibility('visible');
  await drain();
  assert.equal(page.t.state().dq, false);
  assert.equal(page.t.state().warnings, 1);
  assert.equal(page.sent('reportWarning').length, 1);
  assert.equal(nudges(page).length, before, 'the report is fire-and-forget: the push waits for it to land');
  await page.timer.advance(NUDGE_DELAY);
  const pushed = nudges(page).slice(before);
  assert.equal(pushed.length, 1);
  assertNudgeShape(pushed[0]);
});

test('nudge: a confirmed result is pushed at once — with the token, and with nothing else', async () => {
  const page = completePage({ gateway: 'https://gw.example/' });
  await register(page);
  await startExam(page);
  const before = nudges(page).length;
  page.t.finish();
  await drain();
  assert.equal(page.sent('submitResult').length, 1);
  const pushed = nudges(page).slice(before);
  assert.equal(pushed.length, 1, 'exactly one push, the moment the server confirmed the result');
  assertNudgeShape(pushed[0]);
  assert.equal(page.sent('markFinished').length, 0, 'the finished ping is a beacon, not a POST');
  assert.ok(page.beacons.some(b => b.action === 'markFinished'));
});

test('nudge: "finished on device" is pushed 2 s later, so the beacon lands in Google first', async () => {
  // A Worker that re-read the sheet BEFORE the beacon landed would cache the row
  // exactly as it was and show the examiner nothing new. Nobody is watching this
  // timer: the green banner is already on the examinee's screen.
  const page = completePage({ gateway: 'https://gw.example/',
    reply: r => r.action === 'submitResult' ? { __hang: true } : undefined });
  await register(page);
  await startExam(page);
  const before = nudges(page).length;
  page.t.finish();
  await drain();
  assert.ok(page.beacons.some(b => b.action === 'markFinished'));
  assert.equal(nudges(page).length, before, 'nothing is pushed while the beacon is still in the air');
  await page.timer.advance(NUDGE_DELAY - 1);
  assert.equal(nudges(page).length, before);
  await page.timer.advance(1);
  const pushed = nudges(page).slice(before);
  assert.equal(pushed.length, 1, 'and then the Worker is told to drop its copy of the session');
  assertNudgeShape(pushed[0]);
});

test('nudge: a result the server did NOT confirm pushes nothing', async () => {
  for (const failure of [{ __network: true }, { status: 'error', examineeTokenError: 'mismatch' }]) {
    const page = completePage({ gateway: 'https://gw.example/',
      reply: r => r.action === 'submitResult' ? failure : undefined });
    await register(page);
    await startExam(page);
    const before = nudges(page).length;
    page.t.finish();
    await drain();
    await page.timer.advance(NUDGE_DELAY - 1);   // before the markFinished push, which is a different write
    assert.equal(page.sent('submitResult').length, 1);
    assert.equal(nudges(page).length, before, 'the Worker is told about a write only once the server owns it');
  }
});

test('nudge: a whole happy path announces itself exactly four times, in order', async () => {
  // register → start → submit → "finished on device". Four writes, four pushes:
  // no write goes unannounced, and none is announced twice.
  const page = completePage({ gateway: 'https://gw.example/' });
  await register(page);
  const registeredAt = page.timer.now;
  assert.equal(nudges(page).length, 1);
  await startExam(page);
  assert.equal(nudges(page).length, 2);
  const startedAt = page.timer.now;
  page.t.finish();
  await drain();
  assert.equal(nudges(page).length, 3, 'the result is confirmed synchronously here, so its push is immediate');
  const finishedAt = page.timer.now;
  await page.timer.advance(NUDGE_DELAY - 1);
  assert.equal(nudges(page).length, 3, 'the markFinished push is still waiting for its beacon to land');
  await page.timer.advance(1);
  const all = nudges(page);
  assert.equal(all.length, 4, 'exactly four');
  for (const push of all) assertNudgeShape(push);
  assert.ok(all[0].__at <= registeredAt, 'registration: at once');
  assert.ok(all[1].__at <= startedAt && all[1].__at >= registeredAt, 'startExam: at once');
  assert.equal(all[2].__at, finishedAt, 'submitResult: at once');
  assert.equal(all[3].__at, finishedAt + NUDGE_DELAY, 'markFinished: exactly 2 s behind its beacon');
  await page.timer.advance(10 * 60 * 1000);
  assert.equal(nudges(page).length, 4, 'and nothing keeps pushing afterwards');
});

test('nudge: its own failure changes nothing on screen and is counted nowhere', async () => {
  const page = completePage({ gateway: 'https://gw.example/',
    reply: r => String(r.__url).includes('/v1/invalidate') ? { __network: true } : undefined });
  await register(page);
  await startExam(page);
  page.t.finish();
  await drain();
  await page.timer.advance(5000);
  assert.ok(nudges(page).length >= 1, 'it was attempted (and it failed)');
  assert.ok(!page.el('submitFailBanner'), 'the result IS on the server: the examinee is shown no failure');
  assert.match(page.el('submitStatusBanner').innerHTML, /התקבלה/);
  assert.equal(page.t.hasPending(), false, 'nothing was re-armed for retry');
  assert.equal(page.sent('submitResult').length, 1, 'and the result was not sent again');
  assert.equal(page.t.degraded(), false, 'a failed push must never mark Apps Script degraded');
});

test('nudge: it is never sent without a Worker url', async () => {
  // A session that names no Worker never gets past the code screen (§11.5), so
  // there is no url to push to — and the guard says so even when the caller
  // hands over a complete identity.
  const noGateway = completePage({ gateway: '' });
  noGateway.el('sessionCodeInput').value = 'ABC12345';
  noGateway.el('codeSubmitBtn').click();
  await drain();
  noGateway.t.nudge('ABC12345', '123456789', 'tok-1');
  await drain();
  assert.equal(nudges(noGateway).length, 0);

  // ...and a page that knows nothing yet (no code, no id, no token) pushes
  // nothing either, however it is called.
  const fresh = completePage({ gateway: 'https://gw.example/' });
  fresh.t.nudge();
  await drain();
  assert.equal(nudges(fresh).length, 0);

  // With all three, it goes — this is the path a RESEND uses, announcing itself
  // with the identity the stored attempt carries rather than the page's current one.
  const live = completePage({ gateway: 'https://gw.example/' });
  await register(live);
  const before = nudges(live).length;
  live.t.nudge('OLD12345', '987654321', 'tok-old');
  await drain();
  const pushed = nudges(live).slice(before);
  assert.equal(pushed.length, 1);
  assert.equal(pushed[0].sessionCode, 'OLD12345');
  assert.equal(pushed[0].idNumber, '987654321');
  assert.equal(pushed[0].examineeToken, 'tok-old');
});

// ===================== 7. restore =====================
test('D20: a state left by another examinee is never adopted silently', async () => {
  const local = memoryStore();
  local.setItem('ext_examinee_state_123456789', JSON.stringify({
    sessionCode: 'ABC12345', sessionData: { site: 'בדיקת נתונים', license: 'B', language: 'he', gateway: { url: 'https://gw.example/' } },
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
    sessionCode: 'ABC12345', sessionData: { site: 'בדיקת נתונים', license: 'B', language: 'he', gateway: { url: 'https://gw.example/' } },
    examineeData: { idNumber: '123456789', fullName: 'ישראל ישראלי', license: 'B', language: 'he' },
    examineeToken: 'tok-old', screen: 'screenInstructions', savedAt: 1
  }));
  const page = completePage({ local, reply: r => r.kind === 'approval' ? { status: 'ok', approval: 'waiting', audioMode: 'off' } : undefined });
  await drain();
  page.el('restoreYes').click();
  await drain(); await page.timer.advance(50); await drain();
  assert.equal(page.t.state().id, '123456789');
  assert.equal(page.t.state().token, 'tok-old');
  assert.equal(page.t.state().screen, 'screenInstructions');
});

test('D20: a saved wait whose registration was cancelled starts fresh, not on the decision screen', async () => {
  const local = memoryStore();
  local.setItem('ext_examinee_state_123456789', JSON.stringify({
    sessionCode: 'ABC12345', sessionData: { site: 'בדיקת נתונים', license: 'B', language: 'he', gateway: { url: 'https://gw.example/' } },
    examineeData: { idNumber: '123456789', fullName: 'ישראל ישראלי', license: 'B', language: 'he' },
    examineeToken: 'tok-old', screen: 'screenInstructions', savedAt: 1
  }));
  const page = completePage({ local, reply: r => r.kind === 'approval' ? { status: 'ok', approval: 'cancelled' } : undefined });
  await drain();
  page.el('restoreYes').click();
  await drain(); await page.timer.advance(50); await drain();

  // A decision found on RELOAD still means "that registration is over": the
  // examinee is put back on the code screen, not in front of a message about a
  // wait they were not watching.
  assert.equal(page.t.state().screen, 'screenCode');
  assert.equal(local.getItem('ext_examinee_state_123456789'), null, 'the dead row is not carried forward');
  assert.equal(page.el('rejectedMsg').style.display, 'none');
  assert.equal(page.el('backToRegisterBtn').style.display, 'none');
  assert.equal(pollsOf(page, 'approval').length, 1, 'one verification poll, and no chain behind it');
  await page.timer.advance(60000);
  assert.equal(pollsOf(page, 'approval').length, 1);
});

test('D20: this tab\'s own state is restored without a question', async () => {
  const session = memoryStore();
  session.setItem('ext_examinee_state', JSON.stringify({
    sessionCode: 'ABC12345', sessionData: { site: 'בדיקת נתונים', license: 'B', language: 'he', gateway: { url: 'https://gw.example/' } },
    examineeData: { idNumber: '123456789', fullName: 'ישראל ישראלי', license: 'B', language: 'he' },
    examineeToken: 'tok-old', screen: 'screenInstructions', savedAt: 1
  }));
  const page = completePage({ session, reply: r => r.kind === 'approval' ? { status: 'ok', approval: 'waiting', audioMode: 'off' } : undefined });
  await drain(); await page.timer.advance(50); await drain();
  assert.ok(!page.el('restoreConfirm'), 'the tab that wrote the state owns it');
  assert.equal(page.t.state().id, '123456789');
});

/** Runs one exam up to an answer and hands back the stores it wrote. */
async function examInProgressStores() {
  const local = memoryStore(), session = memoryStore();
  const first = completePage({ local, session });
  await register(first);
  await startExam(first);
  first.t.answerCurrent(3);
  const savedAnswer = plain(first.t.state().answers[0]);
  await first.timer.advance(5000);
  return { local, session, savedAnswer, first };
}
const activeBlob = session => JSON.parse(session.getItem('ext_exam_active'));

test('resume: an exam that has started is LOCAL — a reload with no network at all still resumes it', async () => {
  const { local, session, savedAnswer } = await examInProgressStores();
  assert.equal(activeBlob(session).records.length, TOTAL, 'the texts were saved next to the exam');

  // Nothing answers: no Worker, no Apps Script, no Pages. This is the invariant.
  const second = completePage({ local, session, reply: () => ({ __network: true }) });
  await drain(); await second.timer.advance(10000); await drain();
  const state = second.t.state();
  assert.equal(state.inProgress, true, 'the exam resumes');
  assert.equal(state.questions.length, TOTAL);
  assert.deepEqual(plain(state.answers[0]), savedAnswer, 'with the answer exactly as it was displayed');
  assert.equal(second.el('examArea').querySelector('.q-text').textContent, BANK_FILES.he[0].t);
  assert.equal(bankCalls(second).length, 0, 'it never asked the Worker');
  assert.equal(second.sent('startExam').length, 0, 'and never asked the server for a new grant');
  // and every language is still there, so a switch mid-reconnection is local too
  second.t.switchLang('ru'); await drain();
  assert.equal(second.el('examArea').querySelector('.q-text').textContent, BANK_FILES.ru[0].t);
});

test('resume: the resumed exam saves the texts and the grant again, so a second reload is local too', async () => {
  const { local, session } = await examInProgressStores();
  const second = completePage({ local, session, reply: () => ({ __network: true }) });
  await drain(); await second.timer.advance(10000); await drain();
  const blob = activeBlob(session);
  assert.equal(blob.records.length, TOTAL);
  assert.deepEqual(blob.bank, EXAM_BANK, 'the grant rode along even though it was never used');

  const third = completePage({ local, session, reply: () => ({ __network: true }) });
  await drain(); await third.timer.advance(10000); await drain();
  assert.equal(third.t.state().inProgress, true);
  assert.equal(bankCalls(third).length, 0);
});

test('resume: a blob with no texts (an older one, or one that did not fit) uses the STORED grant', async () => {
  const { local, session, savedAnswer } = await examInProgressStores();
  const blob = activeBlob(session);
  delete blob.records;                       // exactly what a pre-snapshot blob looks like
  session.setItem('ext_exam_active', JSON.stringify(blob));

  const second = completePage({ local, session });
  await drain(); await second.timer.advance(100); await drain();
  const state = second.t.state();
  assert.equal(state.inProgress, true);
  assert.deepEqual(plain(state.answers[0]), savedAnswer);
  assert.equal(second.el('examArea').querySelector('.q-text').textContent, BANK_FILES.he[0].t);
  assert.equal(bankCalls(second).length, 1);
  assert.equal(bankCalls(second)[0].grant, 'g1', 'the grant that was saved with the exam');
  assert.equal(second.sent('startExam').length, 0, 'a reload still costs the server nothing');
});

test('resume: a truncated snapshot is not trusted — it falls back to the grant', async () => {
  const { local, session } = await examInProgressStores();
  const blob = activeBlob(session);
  blob.records = blob.records.slice(0, 5);   // 5 of 30: some questions would be blank
  session.setItem('ext_exam_active', JSON.stringify(blob));

  const second = completePage({ local, session });
  await drain(); await second.timer.advance(100); await drain();
  assert.equal(second.t.state().inProgress, true);
  assert.equal(bankCalls(second).length, 1, 'a partial copy is no better than none');
  assert.equal(second.el('examArea').querySelector('.q-text').textContent, BANK_FILES.he[0].t);
});

test('resume: a stored grant the Worker refuses costs ONE startExam, not the exam', async () => {
  const { local, session, savedAnswer } = await examInProgressStores();
  const blob = activeBlob(session);
  delete blob.records;                       // force the network path
  session.setItem('ext_exam_active', JSON.stringify(blob));

  const second = completePage({ local, session, reply(r) {
    const url = String(r.__url);
    if (url.includes('/v1/bank') && r.grant === 'g1') return { __status: 403, __raw: '{"status":"error","code":"grant_invalid"}' };
    if (r.action === 'startExam') return { status: 'ok', examMinutes: 40, extraMinutes: 0, audioMode: 'off',
      questions: SERVER_QUESTIONS, bank: { url: 'https://gw.test', grant: 'g2' } };
    return undefined;
  } });
  await drain(); await second.timer.advance(10000); await drain();
  assert.equal(second.sent('startExam').length, 1, 'startExam is idempotent: the same ids come back with a fresh grant');
  assert.deepEqual(bankCalls(second).map(r => r.grant), ['g1', 'g1', 'g1', 'g2'],
    'the stored grant on its own ladder first, then the new one');
  const state = second.t.state();
  assert.equal(state.inProgress, true);
  assert.deepEqual(plain(state.questions).map(q => q.id), ISSUED_IDS);
  assert.deepEqual(plain(state.answers[0]), savedAnswer);
  assert.equal(second.el('examArea').querySelector('.q-text').textContent, BANK_FILES.he[0].t);
});

test('resume: a snapshot too big for storage is skipped, and the exam is unharmed', async () => {
  // 30 questions x 7 languages of ~8 KB each is past the 1.5 MB cap. The exam
  // must run exactly as before; only the offline resume is given up.
  const huge = 'x'.repeat(8000);
  const local = memoryStore(), session = memoryStore();
  const first = completePage({ local, session, reply(r) {
    if (!String(r.__url).includes('/v1/bank')) return undefined;
    const body = bankAnswer(ISSUED_IDS);
    for (const q of body.questions) for (const lang of Object.keys(q.l)) q.l[lang].t = huge + q.l[lang].t;
    return body;
  } });
  await register(first);
  await startExam(first);
  first.t.answerCurrent(1);
  assert.equal(first.t.state().inProgress, true, 'the exam itself is untouched');
  const blob = activeBlob(session);
  assert.equal(blob.records, null, 'the snapshot was skipped');
  assert.equal(blob.activeQuestions.length, TOTAL, 'but the exam state was stored — that is the part with no fallback');
  assert.deepEqual(blob.bank, EXAM_BANK, 'and the grant, which is now the only way back');

  const second = completePage({ local, session, reply(r) {
    if (!String(r.__url).includes('/v1/bank')) return undefined;
    const body = bankAnswer(ISSUED_IDS);
    for (const q of body.questions) for (const lang of Object.keys(q.l)) q.l[lang].t = huge + q.l[lang].t;
    return body;
  } });
  await drain(); await second.timer.advance(100); await drain();
  assert.equal(second.t.state().inProgress, true, 'and the reload resumes over the network');
  assert.equal(bankCalls(second).length, 1);
});

test('resume: a sessionStorage quota error drops the texts, never the exam state', async () => {
  const local = memoryStore(), session = memoryStore();
  // The blob with the texts is ~19 KB here, the one without ~2.5 KB: a 12 KB
  // ceiling refuses the first and accepts the second, as a real quota would.
  session.rejectBig = 12000;
  const page = completePage({ local, session });
  await register(page);
  await startExam(page);
  page.t.answerCurrent(1);
  const blob = activeBlob(session);
  assert.equal(blob.records, null);
  assert.equal(blob.activeQuestions.length, TOTAL);
  assert.equal(blob.userAnswers.filter(Boolean).length, 1, 'the answer is on the device, which is the whole point');
  assert.equal(page.t.state().inProgress, true);
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

test('source: the page has no way left to read a public bank', () => {
  const src = examinee.replace(/\r/g, '');
  assert.ok(!/loadExamBanks/.test(src), 'the per-language loader is gone with the public bank');
  assert.ok(!/QuestionBank\.prefetch/.test(src), 'and so is warming languages during the wait');
  assert.ok(!/QuestionBank\.load\s*\(/.test(src), 'only loadGrant remains');
  assert.ok(!/QuestionBank\.(loadManifest|configure|build)\b/.test(src));
  assert.ok(!/bank\/[a-z]{2}\.json|bank\/manifest\.json/.test(src), 'no same-origin bank path is referenced');
  assert.equal((src.match(/QuestionBank\.loadGrant/g) || []).length, 1, 'ONE place knows how to fetch the texts');
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
  assert.ok(/ExamTransport\.drainLog\(\)/.test(src));
});

test('source: the direct poll route is gone — the Worker is the only one', () => {
  const src = examinee.replace(/\r/g, '');
  assert.ok(!/createFailover/.test(src), 'nothing to fail over to');
  for (const gone of ['directApprovalCall', 'directStatusCall', 'APPROVAL_DIRECT_BASE_MS', 'EXAMSTATUS_DIRECT_BASE_MS']) {
    assert.ok(!new RegExp('\\b' + gone + '\\b').test(src), gone + ' is gone');
  }
  assert.ok(!/action: 'checkApproval'|action: 'getExamStatus'/.test(src),
    'the page has no way left to poll Apps Script directly');
  assert.ok(!/gatewayUrl\(\)\s*\?/.test(src), 'and no cadence branches on whether a gateway exists');
  assert.ok(src.includes('המערכת אינה מוגדרת (Worker) — פנה למנהל המערכת'), 'a missing Worker is named, not polled');
  assert.ok(!/falls? back to (the )?direct|five minutes/.test(src), 'and the comments do not promise a fallback');
});

test('source: the re-arm and the device push are wired exactly where §13.4/§13.5 put them', () => {
  const src = examinee.replace(/\r/g, '');
  // §13.4: the ONE loop that overrides the pacing must let a re-armed answer
  // through, or the fast window would pull it back to 2 s. The other two chains
  // pass no nextDelay at all, so transport's gap is already the last word.
  assert.match(src, /nextDelay: function\(info, paced\) \{\s*\n\s*if \(info\.held > 0 \|\| info\.rearm\) return paced;/,
    'the approval loop follows info.rearm');
  assert.equal((src.match(/nextDelay:/g) || []).length, 1, 'and it is still the only loop that overrides the pacing');
  // §13.5: a RESEND after a re-registration must announce itself with the token
  // the stored attempt was stamped with, never with whatever the page holds now.
  assert.match(src, /nudgeGatewayAfterWrite\(payload\.sessionCode, payload\.idNumber, payload\.examineeToken\);/,
    'the submit push carries the payload\'s own identity');
  assert.match(src, /nudgeGatewayAfterWrite\(attempt\.sessionCode, attempt\.idNumber, attempt\.examineeToken\);/,
    'and the startExam push carries the attempt\'s');
  // Every write this page makes is announced: three confirmed ones push at once,
  // four fire-and-forget ones push NUDGE_AFTER_BEACON_MS later. One declaration,
  // seven call sites, and no bare millisecond anywhere.
  assert.equal((src.match(/\bnudgeGatewayAfterWrite\b/g) || []).length, 8, 'declared once, reached from seven writes');
  assert.equal((src.match(/setTimeout\(nudgeGatewayAfterWrite, NUDGE_AFTER_BEACON_MS\)/g) || []).length, 4,
    'markFinished, disqualify, cancelDisqualify, reportWarning');
  assert.ok(!/nudgeGatewayAfterWrite,\s*\d/.test(src), 'the delay is the named constant, never a number');
  // ONE helper sends every disqualification, so a path cannot be added without
  // its push — and the unload path opts out explicitly, because nothing it
  // schedules would ever run.
  assert.equal((src.match(/action: 'disqualify'/g) || []).length, 1, 'one payload builder, one literal');
  assert.match(src, /function sendDQToServer\(unloading\)/);
  assert.match(src, /if \(!unloading\) setTimeout\(nudgeGatewayAfterWrite, NUDGE_AFTER_BEACON_MS\);/);
  assert.match(src, /sendDQToServer\(true\);/, 'and onBeforeUnload is the only caller that opts out');
  // never a decision, and never an examiner's grant, from this page
  const fn = section(src, 'function nudgeGatewayAfterWrite(', '\n  }\n');
  assert.ok(!/grant|&status=|examinerId/.test(fn), 'an examinee pushes no decision and holds no grant');
  assert.match(fn, /keepalive: true/);
  assert.match(fn, /catch/, 'and it can never throw into the caller');
  // submitFailOnClose is a sendBeacon on the way out (onBeforeUnload): nothing
  // reads its answer, so there is no "the server owns it" moment to announce.
  assert.match(src, /action: 'submitFailOnClose'/);
  assert.ok(!/submitFailOnClose[\s\S]{0,1200}nudgeGatewayAfterWrite/.test(src));
});

test('source: the service worker precaches the shared layers and never the question texts', () => {
  const sw = fs.readFileSync(path.join(app, 'sw-examinee.js'), 'utf8');
  new vm.Script(sw, { filename: 'sw-examinee.js' });
  for (const asset of ['./examinee.html', './shared/transport.js', './shared/bank.js']) assert.ok(sw.includes(asset), asset);
  assert.ok(!/bank\//.test(sw), 'no same-origin bank files exist any more, so no special case for them');
  assert.ok(!/v1\/bank/.test(sw), 'and a grant-bearing Worker answer is never put in a shared cache');
  assert.match(sw, /ignoreSearch: true/, 'still there for the cross-origin image cache');
  assert.match(sw, /req\.method !== 'GET'/, 'GET only — a HEAD can never be cached');
  const cacheLine = sw.split('\n').filter(l => /^var CACHE = '/.test(l));
  assert.equal(cacheLine.length, 1, 'the build tool rewrites exactly this line');
  assert.match(cacheLine[0], /^var CACHE = '[^']*';$/);
});

// ---- r31.2 (22/09/2026): the registration key -------------------------------
// Seen live at 08:35: a registration that Google answered after the phone's
// deadline, a second press, 'כבר רשום' treated as success, and a waiting
// screen with NO token — approved by the examiner, refused by every startExam.
// The page now mints one key per registration attempt series, sends it on
// every try, and adopts the token a resumed answer hands back.
test('r31.2: every registration carries the same regKey across a retry, and a resumed answer restores the token', async () => {
  let registrations = 0;
  const page = completePage({ reply(request) {
    if (request.action !== 'registerExaminee') return undefined;
    registrations++;
    if (registrations === 1) return { __raw: '<html>Google is having trouble</html>' };   // the phone gave up; the row was written
    return { status: 'ok', examineeToken: 'tok-resumed', resumed: true };                 // the same key gets the row's token back
  } });
  await register(page);
  const first = page.sent('registerExaminee');
  assert.equal(first.length, 1);
  assert.match(String(first[0].regKey || ''), /^[A-Za-z0-9_-]{16,64}$/, 'a random key rides along');
  assert.equal(page.el('registerError').style.display, 'block', 'the lost answer is shown as a communication error');
  page.el('registerBtn').click();
  await drain(); await page.timer.advance(100); await drain();
  const sent = page.sent('registerExaminee');
  assert.equal(sent.length, 2);
  assert.equal(sent[1].regKey, sent[0].regKey, 'the retry names the SAME attempt');
  await startExam(page);
  const start = page.sent('startExam');
  assert.equal(start.length, 1);
  assert.equal(start[0].examineeToken, 'tok-resumed', 'the token the resumed answer returned is the one the exam uses');
  assert.equal(page.t.state().inProgress, true);
});

test('r31.2: the registration request waits the unattended deadline, not the 30 s default', () => {
  const src = examinee.replace(/\r/g, '');
  const call = section(src, "      action: 'registerExaminee',", '.then(function(data) {');
  assert.match(call, /regKey: registrationKeyFor\(examineeData\.idNumber\)/);
  assert.match(call, /\}, POLL_TIMEOUT_MS\)\s*$/, 'a registration is a write Google may answer in 30-75 s');
});

// ---- r31.3 (22/09/2026): a link naming another session beats a saved state --
// Seen live: the examiner's link carried today's session, the phone still held
// this morning's (closed) session in its saved state, and the restore won —
// the examinee kept registering into the old code whatever link they opened.
const savedWait = (code, over) => JSON.stringify(Object.assign({
  sessionCode: code, sessionData: { site: 'בדיקת נתונים', license: 'B', language: 'he', gateway: { url: 'https://gw.example/' } },
  examineeData: { idNumber: '123456789', fullName: 'ישראל ישראלי', license: 'B', language: 'he' },
  examineeToken: 'tok-old', screen: 'screenInstructions', savedAt: 1
}, over || {}));

test('r31.3: a link with ANOTHER session code drops a saved pre-exam state and enters the linked session', async () => {
  const session = memoryStore(), local = memoryStore();
  session.setItem('ext_examinee_state', savedWait('OLDCODE1'));
  local.setItem('ext_examinee_state_123456789', savedWait('OLDCODE1'));
  const page = completePage({ session, local, search: '?code=NEWCODE2' });
  await drain(); await page.timer.advance(400); await drain();
  assert.ok(!page.el('restoreConfirm'), 'no restore prompt: the link decided');
  const info = page.sent('getSessionInfo');
  assert.equal(info.length, 1, 'the linked code was submitted');
  assert.equal(info[0].sessionCode, 'NEWCODE2');
  assert.equal(JSON.parse(session.getItem('ext_examinee_state')).sessionCode, 'NEWCODE2', 'the tab now saves the LINKED session, not the stale one');
  assert.equal(local.getItem('ext_examinee_state_123456789'), null, '...and from the per-examinee copy');
});

test('r31.3: a link with the SAME code restores as before, and a link never drops an exam in progress', async () => {
  const same = memoryStore();
  same.setItem('ext_examinee_state', savedWait('ABC12345'));
  const pageSame = completePage({ session: same, search: '?code=ABC12345',
    reply: r => r.kind === 'approval' ? { status: 'ok', approval: 'waiting', audioMode: 'off' } : undefined });
  await drain(); await pageSame.timer.advance(400); await drain();
  assert.equal(pageSame.t.state().id, '123456789', 'same code: the saved wait is restored');
  assert.equal(pageSame.sent('getSessionInfo').length, 0, 'and the code was not re-submitted');

  const { local, session } = await examInProgressStores();
  const pageExam = completePage({ local, session, search: '?code=OTHERCD9' });
  await drain(); await pageExam.timer.advance(400); await drain();
  assert.equal(pageExam.t.state().inProgress, true, 'the running exam is restored');
  assert.equal(pageExam.sent('getSessionInfo').length, 0, 'the stale link in the address bar is ignored');
});
