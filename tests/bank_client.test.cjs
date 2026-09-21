// shared/bank.js — the client-side reader of the question bank, which is no
// longer a public file but a private asset of the session-gateway Worker served
// against a signed grant. The module runs for real in a VM, on a fake clock (the
// retry ladder is a real wait), with the real shared/transport.js underneath it
// so the "a slow Worker must not degrade Apps Script" rule is exercised, not
// assumed. Only fetch and the clock are synthetic.
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const app = path.resolve(__dirname, '..');
const SOURCE = fs.readFileSync(path.join(app, 'shared', 'bank.js'), 'utf8');
const TRANSPORT = fs.readFileSync(path.join(app, 'shared', 'transport.js'), 'utf8');

const plain = v => JSON.parse(JSON.stringify(v));   // VM objects are not host objects
async function drain() { for (let i = 0; i < 30; i++) await Promise.resolve(); }

class Timers {
  constructor() { this.now = 0; this.nextId = 0; this.jobs = new Map(); }
  set = (cb, ms) => { const id = ++this.nextId; this.jobs.set(id, { cb, at: this.now + Number(ms || 0) }); return id; };
  clear = id => this.jobs.delete(id);
  async advance(ms) {
    await drain();                       // let whatever is already in flight settle and schedule
    const until = this.now + ms;
    for (let guard = 0; guard < 10000; guard++) {
      const next = [...this.jobs].filter(([, job]) => job.at <= until).sort((a, b) => a[1].at - b[1].at)[0];
      if (!next) { this.now = until; await drain(); return; }
      this.now = next[1].at;
      this.jobs.delete(next[0]);
      next[1].cb();
      await drain();
    }
    throw new Error('timer loop did not settle');
  }
}

const BANK = { url: 'https://gw.test/', grant: 'PAY.SIG' };

// One /v1/bank record: the question in every language that has it.
const RECORDS = {
  1: { id: 1, l: {
    he: { t: 'שאלה עברית', a: ['א', 'ב', 'ג', 'ד'], i: 'TQ_PIC_1.jpg' },
    ru: { t: 'Вопрос', a: ['а', 'б', 'в', 'г'], i: 'TQ_PIC_1.jpg' },
    en: { t: 'English question', a: ['a', 'b', 'c', 'd'], i: 'TQ_PIC_1.jpg' }
  } },
  124: { id: 124, l: {
    he: { t: 'ניסוח רגיל', a: ['1', '2', '3', '4'], i: '', v: { D: { t: 'ניסוח לדרגה D', a: ['1D', '2D', '3D', '4D'] } } },
    en: { t: 'Plain wording', a: ['w', 'x', 'y', 'z'], i: '', v: { D: { t: 'Bus wording' } } }
  } }
};
const okBody = (ids, missing) => ({ status: 'ok', build: 'bank-2026-09-21',
  questions: ids.map(id => RECORDS[id]).filter(Boolean), missing: missing || [] });

/**
 * Loads the real module. `reply(url, attemptNumber)` returns either a JSON body,
 * `{ __network: true }`, or `{ __status, __raw }`.
 */
function load(reply, { withTransport = true } = {}) {
  const timer = new Timers();
  const requested = [];
  const fetchImpl = (url, opts) => {
    requested.push(String(url));
    const data = reply(String(url), requested.length, opts);
    if (data && data.__network) return Promise.reject(new TypeError('Failed to fetch'));
    const body = data && data.__raw !== undefined ? data.__raw : JSON.stringify(data);
    return Promise.resolve({
      ok: !(data && data.__status >= 400), status: (data && data.__status) || 200,
      text: () => Promise.resolve(body), json: () => Promise.resolve(JSON.parse(body))
    });
  };
  const ctx = {
    console: { log() {}, warn() {} },
    Date: class extends Date { static now() { return timer.now; } },
    Math, JSON, Promise, Error, TypeError, String, Number, Array, Object, encodeURIComponent,
    setTimeout: timer.set, clearTimeout: timer.clear, setInterval: timer.set, clearInterval: timer.clear,
    AbortController, fetch: fetchImpl
  };
  ctx.window = ctx;
  vm.createContext(ctx);
  if (withTransport) vm.runInContext(TRANSPORT, ctx, { filename: 'shared/transport.js' });
  vm.runInContext(SOURCE, ctx, { filename: 'shared/bank.js' });
  return { B: ctx.QuestionBank, T: ctx.ExamTransport, requested, timer };
}

const settle = p => p.then(value => ({ value }), error => ({ error }));

// ===== 1. loadGrant =====
test('bank: loadGrant buys the issued questions with the grant and ingests every language', async () => {
  const { B, requested } = load(() => okBody([1, 124]));
  const res = await B.loadGrant(BANK);
  assert.deepEqual(requested, ['https://gw.test/v1/bank?grant=PAY.SIG']);
  assert.deepEqual(plain(res), { build: 'bank-2026-09-21', count: 2, missing: [] });
  // one request, seven-language-ready: this is what makes a switch local
  assert.equal(B.has('he'), true);
  assert.equal(B.has('ru'), true);
  assert.equal(B.has('en'), true);
  assert.equal(B.has('ar'), false, 'a language the question has no text in stays absent');
  assert.equal(B.get(1, 'ru', 'B').text, 'Вопрос');
});

test('bank: the grant is URL-encoded, and a trailing slash on the url never doubles', async () => {
  const { B, requested } = load(() => okBody([1]));
  await B.loadGrant({ url: 'https://gw.test///', grant: 'a+b/c=.s i g' });
  assert.equal(requested[0], 'https://gw.test/v1/bank?grant=a%2Bb%2Fc%3D.s%20i%20g');
});

test('bank: no url or no grant is rejected without a request', async () => {
  const { B, requested } = load(() => okBody([1]));
  for (const bad of [null, {}, { url: 'https://gw.test' }, { grant: 'g' }]) {
    const { error } = await settle(B.loadGrant(bad));
    assert.match(String(error), /grant missing/);
  }
  assert.equal(requested.length, 0);
});

// ===== 2. the retry ladder =====
test('bank: three attempts, 1.5 s and 3 s apart, and only then a rejection', async () => {
  const { B, requested, timer } = load(() => ({ __network: true }));
  const out = settle(B.loadGrant(BANK));
  await drain();
  assert.equal(requested.length, 1);
  await timer.advance(1499); assert.equal(requested.length, 1, 'the first wait is 1.5 s');
  await timer.advance(1);    assert.equal(requested.length, 2);
  await timer.advance(2999); assert.equal(requested.length, 2, 'the second is 3 s');
  await timer.advance(1);    assert.equal(requested.length, 3);
  const { error } = await out;
  assert.ok(error, 'all three attempts failed, so the caller is told');
  assert.equal(B.has('he'), false, 'and nothing half-loaded is left behind');
  await timer.advance(60000);
  assert.equal(requested.length, 3, 'the ladder ends — it never becomes a storm');
});

test('bank: a blip on the first attempt costs 1.5 s, not an exam', async () => {
  const { B, requested, timer } = load((url, attempt) => attempt === 1 ? { __network: true } : okBody([1, 124]));
  const out = settle(B.loadGrant(BANK));
  await timer.advance(1500);
  const { value } = await out;
  assert.equal(value.count, 2);
  assert.equal(requested.length, 2);
  assert.equal(B.get(1, 'he', 'B').text, 'שאלה עברית');
});

test('bank: a refused grant is retried on the same ladder and no further', async () => {
  const { B, requested, timer } = load(() => ({ __status: 403, __raw: '{"status":"error","code":"grant_invalid"}' }));
  const out = settle(B.loadGrant(BANK));
  await timer.advance(10000);
  const { error } = await out;
  assert.ok(error);
  assert.equal(requested.length, 3, 'three attempts, exactly like a network failure: the page has to ask for a new grant either way');
});

test('bank: an error envelope or a body of the wrong shape is a failure, never an empty exam', async () => {
  for (const body of [{ status: 'error', code: 'bank_unavailable', retryable: true }, { status: 'ok' }, { status: 'ok', questions: 'nope' }, null]) {
    const { B, timer } = load(() => body);
    const out = settle(B.loadGrant(BANK));
    await timer.advance(10000);
    const { error } = await out;
    assert.ok(error, 'body ' + JSON.stringify(body));
    assert.equal(B.has('he'), false);
  }
});

// ===== 3. missing ids =====
test('bank: ids the Worker has no asset for come back in missing', async () => {
  const { B } = load(() => okBody([1], [999, 1000]));
  const res = await B.loadGrant(BANK);
  assert.deepEqual(plain(res.missing), [999, 1000]);
  assert.equal(res.count, 1);
});

test('bank: a record with no usable language joins missing instead of drawing a blank question', async () => {
  const { B } = load(() => ({ status: 'ok', build: 'b', missing: [],
    questions: [RECORDS[1], { id: 77, l: { he: { t: '', a: [] } } }, { id: 78, l: {} }] }));
  const res = await B.loadGrant(BANK);
  assert.deepEqual(plain(res.missing), [77, 78]);
  assert.equal(res.count, 1);
  assert.equal(B.get(77, 'he', 'B'), null);
});

// ===== 4. reading what was ingested =====
test('bank: get() returns the displayed shape and applies the licence variant', async () => {
  const { B } = load(() => okBody([1, 124]));
  await B.loadGrant(BANK);
  const base = B.get(124, 'he', 'B');
  assert.equal(base.text, 'ניסוח רגיל');
  assert.deepEqual(plain(base.answers), ['1', '2', '3', '4']);
  const bus = B.get(124, 'he', 'D');
  assert.equal(bus.text, 'ניסוח לדרגה D', 'the D wording of the Hebrew ids that differ by licence');
  assert.deepEqual(plain(bus.answers), ['1D', '2D', '3D', '4D']);
  const busEn = B.get(124, 'en', 'D');
  assert.equal(busEn.text, 'Bus wording');
  assert.deepEqual(plain(busEn.answers), ['w', 'x', 'y', 'z'], 'a variant with no answers keeps the base answers');
});

test('bank: get() on an unloaded language, or an unknown id, returns null', async () => {
  const { B } = load(() => okBody([1]));
  assert.equal(B.get(1, 'he', 'B'), null, 'nothing is invented before the texts are here');
  await B.loadGrant(BANK);
  assert.equal(B.get(999999, 'he', 'B'), null);
  assert.equal(B.get(1, 'ar', 'B'), null);
});

test('bank: imageUrl is same-origin, and empty when the question has no image', async () => {
  const { B } = load(() => okBody([1, 124]));
  await B.loadGrant(BANK);
  assert.equal(B.imageUrl(B.get(1, 'he', 'B')), 'images/TQ_PIC_1.jpg');
  assert.equal(B.imageUrl(B.get(124, 'he', 'B')), '');
  assert.equal(B.imageUrl(null), '');
  assert.equal(B.imageUrl({ i: 'TQ_PIC_9.jpg' }), 'images/TQ_PIC_9.jpg', 'a raw entry works too');
});

test('bank: a second load replaces an id in place instead of duplicating it', async () => {
  const { B } = load((url, attempt) => attempt === 1 ? okBody([1]) : ({ status: 'ok', build: 'b', missing: [],
    questions: [{ id: 1, l: { he: { t: 'ניסוח מתוקן', a: ['א', 'ב', 'ג', 'ד'], i: '' } } }] }));
  await B.loadGrant(BANK);
  await B.loadGrant(BANK);
  assert.equal(B.get(1, 'he', 'B').text, 'ניסוח מתוקן');
  assert.equal(B.search('he', 'ניסוח').length, 1, 'one entry, not two');
});

test('bank: _reset drops every language', async () => {
  const { B } = load(() => okBody([1]));
  await B.loadGrant(BANK);
  B._reset();
  assert.equal(B.has('he'), false);
  assert.equal(B.get(1, 'he', 'B'), null);
});

// ===== 4b. saving an exam's texts with the exam =====
test('bank: exportIds hands the questions back in the Worker\'s own shape', async () => {
  const { B } = load(() => okBody([1, 124]));
  await B.loadGrant(BANK);
  const records = plain(B.exportIds([1, 124]));
  assert.deepEqual(Object.keys(records[0].l).sort(), ['en', 'he', 'ru']);
  assert.equal(records[0].l.he.t, 'שאלה עברית');
  assert.deepEqual(records[0].l.ru.a, ['а', 'б', 'в', 'г']);
  assert.equal(records[0].l.he.i, 'TQ_PIC_1.jpg');
  assert.deepEqual(records[1].l.he.v.D.a, ['1D', '2D', '3D', '4D'], 'the licence variants ride along');
  assert.deepEqual(plain(B.exportIds([1, 999])).map(r => r.id), [1], 'an id with no text is simply not exported');
  assert.deepEqual(plain(B.exportIds([])), []);
});

test('bank: a snapshot goes back in with no network, and reads exactly as it did before', async () => {
  const { B, requested } = load(() => okBody([1, 124]));
  await B.loadGrant(BANK);
  const saved = JSON.parse(JSON.stringify(plain(B.exportIds([1, 124]))));   // through storage and back
  const before = { he: B.get(1, 'he', 'B'), ru: B.get(1, 'ru', 'B'), busD: B.get(124, 'he', 'D') };
  B._reset();
  assert.equal(B.has('he'), false);

  assert.equal(B.importRecords(saved), 2, 'both ids came back');
  assert.equal(requested.length, 1, 'and not one request was made to do it');
  assert.deepEqual(plain(B.get(1, 'he', 'B')), plain(before.he));
  assert.deepEqual(plain(B.get(1, 'ru', 'B')), plain(before.ru));
  assert.deepEqual(plain(B.get(124, 'he', 'D')), plain(before.busD), 'including the licence variant');
  assert.equal(B.has('en'), true, 'every language is there, so a switch is still local');
});

test('bank: importRecords tolerates rubbish and reports what actually arrived', () => {
  const { B } = load(() => okBody([1]));
  for (const bad of [null, undefined, 'x', 42, {}]) assert.equal(B.importRecords(bad), 0);
  assert.equal(B.importRecords([null, { id: 5 }, { id: 6, l: {} }, { id: 7, l: { he: { t: '', a: [] } } }]), 0,
    'a truncated or corrupted copy counts as nothing, so the caller falls back to the grant');
  assert.equal(B.has('he'), false);
});

// ===== 5. examiner scope =====
test('bank: loadIds asks for exactly those ids and languages', async () => {
  const { B, requested } = load(() => okBody([1, 124]));
  const res = await B.loadIds({ url: 'https://gw.test', grant: 'EX.SIG' }, [1, 124], ['he']);
  assert.equal(requested[0], 'https://gw.test/v1/bank?grant=EX.SIG&ids=1%2C124&langs=he');
  assert.equal(res.count, 2);
  assert.equal(B.has('he'), true);
});

test('bank: loadIds without a language list leaves the choice to the Worker', async () => {
  const { B, requested } = load(() => okBody([1]));
  await B.loadIds({ url: 'https://gw.test', grant: 'EX.SIG' }, [1]);
  assert.equal(requested[0], 'https://gw.test/v1/bank?grant=EX.SIG&ids=1');
});

test('bank: loadFull replaces the whole language so search sees more than the issued ids', async () => {
  const full = [];
  for (let id = 1; id <= 40; id++) full.push({ id, t: 'שאלה מספר ' + id, a: ['א', 'ב', 'ג', 'ד'], i: id === 5 ? 'TQ_PIC_5.jpg' : '' });
  const { B, requested } = load(url => url.indexOf('/v1/bank/full') !== -1 ? full : okBody([1]));
  await B.loadGrant(BANK);
  assert.equal(B.search('he', 'שאלה מספר').length, 0, 'the issued questions are not the bank');
  const res = await B.loadFull({ url: 'https://gw.test', grant: 'EX.SIG' }, 'he');
  assert.equal(requested[1], 'https://gw.test/v1/bank/full?grant=EX.SIG&lang=he');
  assert.deepEqual(plain(res), { lang: 'he', count: 40 });
  assert.equal(B.search('he', 'שאלה מספר', 100).length, 40);
  assert.equal(B.imageUrl(B.get(5, 'he', 'B')), 'images/TQ_PIC_5.jpg');
});

test('bank: an empty full bank is a failure, never an empty search box', async () => {
  const { B, timer } = load(() => []);
  const out = settle(B.loadFull({ url: 'https://gw.test', grant: 'EX.SIG' }, 'he'));
  await timer.advance(10000);
  const { error } = await out;
  assert.match(String(error), /empty bank he/);
});

test('bank: search matches question text and answers, and is bounded', async () => {
  const { B } = load(() => okBody([1, 124]));
  await B.loadGrant(BANK);
  assert.equal(B.search('en', '').length, 0);
  assert.equal(B.search('fr', 'x').length, 0, 'a language with nothing in it returns nothing');
  const hits = B.search('en', 'english');
  assert.equal(hits.length, 1);
  assert.equal(hits[0].id, 1);
  assert.equal(hits[0].image, 'TQ_PIC_1.jpg');
  assert.equal(B.search('en', 'w').length, 1, 'answers are searched too');
  assert.equal(B.search('en', 'o', 1).length, 1, 'the limit is honoured');
});

// ===== 6. it is not the exam backend =====
test('bank: a failing Worker never marks the Apps Script backend degraded', async () => {
  const { B, T, timer } = load(() => ({ __status: 503, __raw: '{"status":"error"}' }));
  const out = settle(B.loadGrant(BANK));
  await timer.advance(10000);
  assert.ok((await out).error);
  assert.equal(T.isBackendDegraded(), false,
    'an HTTP error from the bank Worker would otherwise floor every poll in the page at 30-60 s');
});

test('bank: without ExamTransport the reader still works on a plain fetch', async () => {
  const { B } = load(() => okBody([1]), { withTransport: false });
  const res = await B.loadGrant(BANK);
  assert.equal(res.count, 1);
  assert.equal(B.get(1, 'he', 'B').text, 'שאלה עברית');
});

test('bank: LANGS is still the seven the system ships', () => {
  const { B } = load(() => okBody([1]));
  assert.deepEqual(plain(B.LANGS), ['he', 'ru', 'en', 'ar', 'fr', 'es', 'am']);
});

test('bank: the reader offers no way to fetch a public bank any more', () => {
  const { B } = load(() => okBody([1]));
  for (const gone of ['load', 'loadManifest', 'prefetch', 'configure', 'build']) {
    assert.equal(typeof B[gone], 'undefined', gone + ' is gone — there is no public bank to read');
  }
  assert.ok(!/bank\/manifest\.json/.test(SOURCE), 'and no same-origin bank path is left in the source');
});
