// shared/bank.js — the client-side reader of the static question bank.
// The module runs for real in a VM; fetch is synthetic, and one test reads the
// REAL bank/he.json so the shape the builder produces and the shape the reader
// expects can never drift apart.
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const app = path.resolve(__dirname, '..');
const SOURCE = fs.readFileSync(path.join(app, 'shared', 'bank.js'), 'utf8');

const MANIFEST = {
  build: 'test-build',
  langs: { he: { sha: 'he1', count: 2 }, en: { sha: 'en1', count: 2 }, ru: { sha: 'ru1', count: 2 },
           ar: { sha: 'ar1', count: 1 }, fr: { sha: 'fr1', count: 1 }, es: { sha: 'es1', count: 1 }, am: { sha: 'am1', count: 1 } }
};
const BANKS = {
  he: [
    { id: 1, t: 'שאלה עברית', a: ['א', 'ב', 'ג', 'ד'], i: 'TQ_PIC_1.jpg' },
    { id: 124, t: 'ניסוח רגיל', a: ['1', '2', '3', '4'], i: '', v: { D: { t: 'ניסוח לדרגה D', a: ['1D', '2D', '3D', '4D'] } } }
  ],
  en: [
    { id: 1, t: 'English question', a: ['a', 'b', 'c', 'd'], i: 'TQ_PIC_1.jpg' },
    { id: 124, t: 'Plain wording', a: ['w', 'x', 'y', 'z'], i: '', v: { D: { t: 'Bus wording' } } }
  ],
  ru: [{ id: 1, t: 'Вопрос', a: ['а', 'б', 'в', 'г'], i: '' }]
};

function load({ manifest = MANIFEST, banks = BANKS, fail = () => false } = {}) {
  const requested = [];
  const fetchImpl = url => {
    requested.push(url);
    const failure = fail(url);
    if (failure) return Promise.reject(new TypeError('offline'));
    const body = url.indexOf('manifest.json') !== -1 ? manifest : banks[/bank\/([a-z]{2})\.json/.exec(url)[1]];
    if (body === undefined) return Promise.resolve({ ok: false, status: 404, json: () => Promise.reject(new Error('404')) });
    return Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve(JSON.parse(JSON.stringify(body))) });
  };
  const ctx = { console: { log() {}, warn() {} }, Promise, Error, TypeError, String, Number, Array, Object, JSON, fetch: fetchImpl };
  ctx.window = ctx;
  vm.createContext(ctx);
  vm.runInContext(SOURCE, ctx, { filename: 'shared/bank.js' });
  return { B: ctx.QuestionBank, requested };
}

test('bank: a language is fetched once, versioned by the manifest sha, and memoised', async () => {
  const { B, requested } = load();
  const bank = await B.load('he');
  assert.equal(bank.list.length, 2);
  assert.deepEqual(requested, ['bank/manifest.json', 'bank/he.json?v=he1']);
  await B.load('he'); await B.load('he');
  assert.equal(requested.length, 2, 'the bank is read from memory afterwards');
  assert.equal(B.has('he'), true);
  assert.equal(B.build(), 'test-build');
});

test('bank: two languages share one manifest read', async () => {
  const { B, requested } = load();
  await Promise.all([B.load('he'), B.load('en')]);
  assert.equal(requested.filter(u => u.includes('manifest')).length, 1);
  assert.ok(requested.includes('bank/he.json?v=he1'));
  assert.ok(requested.includes('bank/en.json?v=en1'));
});

test('bank: get() returns the displayed shape and applies the licence variant', async () => {
  const { B } = load();
  await B.load('he');
  const plain = B.get(124, 'he', 'B');
  assert.equal(plain.text, 'ניסוח רגיל');
  assert.deepEqual(plain.answers, ['1', '2', '3', '4']);
  const bus = B.get(124, 'he', 'D');
  assert.equal(bus.text, 'ניסוח לדרגה D', 'the D wording of the 5 Hebrew ids that differ by licence');
  assert.deepEqual(bus.answers, ['1D', '2D', '3D', '4D']);
  await B.load('en');
  const busEn = B.get(124, 'en', 'D');
  assert.equal(busEn.text, 'Bus wording');
  assert.deepEqual(busEn.answers, ['w', 'x', 'y', 'z'], 'a variant with no answers keeps the base answers');
});

test('bank: get() on a language that is not loaded, or an unknown id, returns null', async () => {
  const { B } = load();
  assert.equal(B.get(1, 'he', 'B'), null, 'nothing is invented before the bank is here');
  await B.load('he');
  assert.equal(B.get(999999, 'he', 'B'), null);
  assert.equal(B.get(1, 'ar', 'B'), null);
});

test('bank: imageUrl is same-origin, and empty when the question has no image', async () => {
  const { B } = load();
  await B.load('he');
  assert.equal(B.imageUrl(B.get(1, 'he', 'B')), 'images/TQ_PIC_1.jpg');
  assert.equal(B.imageUrl(B.get(124, 'he', 'B')), '');
  assert.equal(B.imageUrl(null), '');
  assert.equal(B.imageUrl({ i: 'TQ_PIC_9.jpg' }), 'images/TQ_PIC_9.jpg', 'a raw entry works too');
});

test('bank: a missing manifest falls back to the unversioned url instead of failing', async () => {
  const { B, requested } = load({ fail: url => url.includes('manifest.json') });
  const bank = await B.load('he');
  assert.equal(bank.list.length, 2);
  assert.deepEqual(requested, ['bank/manifest.json', 'bank/he.json'], 'the service-worker copy is still reachable');
});

test('bank: a failed load rejects and can be retried (the promise is not poisoned)', async () => {
  let offline = true;
  const { B, requested } = load({ fail: url => offline && url.includes('he.json') });
  await assert.rejects(() => B.load('he'));
  assert.equal(B.has('he'), false);
  offline = false;
  const bank = await B.load('he');
  assert.equal(bank.list.length, 2);
  assert.ok(requested.filter(u => u.includes('he.json')).length >= 2);
});

test('bank: an unknown language is rejected without a request', async () => {
  const { B, requested } = load();
  await assert.rejects(() => B.load('zz'), /unknown language/);
  assert.equal(requested.length, 0);
});

test('bank: an empty file is treated as a failure, never as an empty exam', async () => {
  const { B } = load({ banks: { he: [] } });
  await assert.rejects(() => B.load('he'), /empty bank/);
});

test('bank: prefetch warms every language and swallows the failures', async () => {
  const { B, requested } = load({ banks: { he: BANKS.he, en: BANKS.en } });   // the rest 404
  B.prefetch();
  await new Promise(r => setTimeout(r, 10));
  assert.ok(requested.some(u => u.includes('he.json')));
  assert.ok(requested.some(u => u.includes('en.json')));
  assert.equal(B.has('he'), true);
  assert.equal(B.has('ar'), false, 'a language with no file simply stays unloaded');
});

test('bank: prefetch of a chosen pair loads exactly those', async () => {
  const { B, requested } = load();
  B.prefetch(['he', 'ru']);
  await new Promise(r => setTimeout(r, 10));
  assert.ok(requested.some(u => u.includes('he.json')));
  assert.ok(requested.some(u => u.includes('ru.json')));
  assert.ok(!requested.some(u => u.includes('en.json')));
});

test('bank: search matches question text and answers, and is bounded', async () => {
  const { B } = load();
  await B.load('en');
  assert.equal(B.search('en', '').length, 0);
  assert.equal(B.search('fr', 'x').length, 0, 'a language that is not loaded returns nothing');
  const hits = B.search('en', 'english');
  assert.equal(hits.length, 1);
  assert.equal(hits[0].id, 1);
  assert.equal(hits[0].image, 'TQ_PIC_1.jpg');
  assert.equal(B.search('en', 'w').length, 1, 'answers are searched too');
  assert.equal(B.search('en', 'o', 1).length, 1, 'the limit is honoured');
});

test('bank: configure() moves the base path for pages served from a sub-folder', async () => {
  const { B, requested } = load();
  B.configure({ base: '../bank' });
  await B.load('he');
  assert.ok(requested[0].startsWith('../bank/'), requested[0]);
});

// ===== the real files =====
test('bank: the real bank/he.json + manifest load through the real reader', async () => {
  const manifest = JSON.parse(fs.readFileSync(path.join(app, 'bank', 'manifest.json'), 'utf8'));
  const he = JSON.parse(fs.readFileSync(path.join(app, 'bank', 'he.json'), 'utf8'));
  const { B, requested } = load({ manifest, banks: { he } });
  const bank = await B.load('he');
  assert.equal(bank.list.length, manifest.langs.he.count, 'the manifest count is the file');
  assert.equal(requested[1], 'bank/he.json?v=' + manifest.langs.he.sha);
  const first = B.get(he[0].id, 'he', 'B');
  assert.equal(typeof first.text, 'string');
  assert.equal(first.answers.length, 4);
  assert.equal('ci' in first, false, 'the public bank carries no correct answer');
  const withVariant = he.find(e => e.v);
  if (withVariant) {
    const licence = Object.keys(withVariant.v)[0];
    const applied = B.get(withVariant.id, 'he', licence);
    const base = B.get(withVariant.id, 'he', 'B');
    assert.notDeepEqual([applied.text, applied.answers], [base.text, base.answers]);
  }
  const withImage = he.find(e => e.i);
  assert.match(B.imageUrl(B.get(withImage.id, 'he', 'B')), /^images\/TQ_PIC_[A-Za-z0-9_]+\.jpg$/);
});
