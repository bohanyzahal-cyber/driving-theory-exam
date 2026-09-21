// Run: node tests/database_reliability.test.cjs
//
// What the exam WRITES, replayed against explicit expected tables. Until the
// 21/09/2026 rewrite this suite compared every scenario against a frozen copy of
// the audited server (git show 232ed3c) — a differential gate that could only
// ever say "same as before", which is exactly what a rewrite cannot promise.
// The scenarios are the same ones; each now states the rows it expects.
//
// Everything here is synthetic — no network, no Google, no real ID numbers.
const test = require('node:test');
const assert = require('node:assert/strict');
const { createEnv } = require('./helpers/server_env.cjs');

const SESSION = 'SYNTHETIC';
const ID = '900000001';
const TOKEN = 'synthetic-examinee-token';
const REGISTERED_AT = '2026-09-22T06:00:00Z';
const QUESTION_COUNT = 30;
const PENDING_HEADER = Array(19).fill('header');
const RESULTS_HEADER = Array(30).fill('header');
const EXAMS_HEADER = Array(6).fill('header');
// Synthetic answer key: question id → correct original index, per language.
// he/ru/ar deliberately disagree, as the real banks do.
const KEY = { he: {}, ru: {}, ar: {} };
const IDS = Array.from({ length: QUESTION_COUNT }, (_, i) => 101 + i);
for (const id of IDS) {
  KEY.he[id] = id % 4;
  KEY.ru[id] = (id + 1) % 4;
  KEY.ar[id] = (id + 2) % 4;
}
const ORDER = [2, 0, 3, 1];
function questionMap(overrides) {
  return IDS.map((id, i) => Object.assign({
    qIdx: i, qId: id, shuffleOrder: ORDER.slice(), correctShuffledIdx: ORDER.indexOf(KEY.he[id]), topic: 'חוק'
  }, overrides || {}));
}
function answers(count, lang) {
  return questionMap().map((entry, i) => {
    const key = ORDER.indexOf(KEY[lang || 'he'][entry.qId]);
    return {
      qIdx: i, selected: i < count ? key : (key + 1) % 4, langAtAnswer: lang || 'he',
      q: 'שאלה ' + entry.qId, a: ORDER.map(o => 'תשובה ' + entry.qId + '-' + o)
    };
  });
}
function pendingRow(status, overrides) {
  const row = Array(19).fill('');
  row[0] = SESSION; row[1] = ID; row[2] = 'נבחן סינתטי'; row[4] = REGISTERED_AT;
  row[5] = status || 'in_exam'; row[6] = 'he'; row[8] = 'B'; row[9] = 'off'; row[12] = TOKEN;
  return Object.assign(row, overrides || {});
}
function examRow(overrides) {
  return Object.assign([SESSION, ID, JSON.stringify(questionMap()), REGISTERED_AT, 'he', 0], overrides || {});
}
function resultRow(overrides) {
  const row = Array(30).fill('');
  row[0] = '22/09/2026 07:10'; row[1] = ID; row[2] = 'נבחן סינתטי'; row[4] = 'B';
  row[5] = '0/30'; row[6] = '0%'; row[7] = 'נכשל'; row[8] = '00:30'; row[12] = 'he';
  row[13] = SESSION; row[14] = 1; row[18] = 'synthetic-existing-link';
  return Object.assign(row, overrides || {});
}
function env(spec) {
  const s = spec || {};
  const e = createEnv({
    sheets: {
      'ממתינים': [PENDING_HEADER, ...(s.pending || [pendingRow()])],
      'מבחנים': [EXAMS_HEADER, ...(s.exams || [examRow()])],
      'תוצאות': [RESULTS_HEADER, ...(s.results || [])]
    }
  });
  e.ctx.ANSWER_KEY_BY_LANG = KEY;
  // The 30 fixture questions have a hand-written key per language; every other
  // id (the real index has 1,700, and a draw must be able to fill a blueprint)
  // gets a deterministic synthetic one that still differs between languages.
  e.ctx.lookupCorrectIndex = (id, lang) => {
    const byLang = KEY[lang];
    if (byLang && Object.prototype.hasOwnProperty.call(byLang, id)) return byLang[id];
    return (Number(id) + Math.max(0, Object.keys(KEY).indexOf(lang))) % 4;
  };
  return e;
}
function submit(e, overrides) {
  return e.json(e.ctx.doPost({ postData: { contents: JSON.stringify(Object.assign({
    action: 'submitResult', origin: 'examinee-app', sessionCode: SESSION, idNumber: ID, examineeToken: TOKEN,
    fullName: 'נבחן סינתטי', phone: '000', license: 'B', language: 'he',
    score: 0, total: 30, percent: 0, passed: false, time: '10:00',
    examinerName: 'בוחן סינתטי', site: 'אתר סינתטי', classroom: 'כיתה', population: 'אוכלוסיה',
    audioMode: 'on', device: 'tablet', languageHistory: ['he'], answers: answers(30)
  }, overrides || {})) } }));
}
const results = e => e.rows('תוצאות');
const pendingStatuses = e => e.rows('ממתינים').slice(1).map(r => r[5]);

test('the sheet headers the exam writes to are the ones it assumes', () => {
  const e = env();
  assert.equal(e.ctx.SHEET_HEADERS['תוצאות'].length, 30, 'the result row is built column by column');
  assert.equal(e.ctx.SHEET_HEADERS['ממתינים'].length, 19, 'the status writer addresses up to column 19');
  assert.equal(e.ctx.SHEET_HEADERS['מבחנים'].length, 6, 'the registration row is six columns');
});

test('a verified result carries its metadata, and a lost-response retry adds nothing', () => {
  const e = env();
  const first = submit(e);
  assert.equal(first.status, 'ok');
  assert.equal(results(e).length, 2);
  const row = results(e)[1];
  assert.equal(row[5], '30/30');
  assert.equal(row[6], '100%');
  assert.equal(row[7], 'עבר');
  assert.equal(row[12], 'he');
  assert.equal(row[13], SESSION);
  assert.equal(row[14], 1);
  assert.equal(row[15], '', 'a perfect score has no wrong-answer detail');
  assert.equal(row[19], 'אוכלוסיה');
  assert.equal(row[21], 'on');
  assert.equal(row[22], 'מאומת');
  assert.equal(row[28], 'he');
  assert.equal(row[29], 'tablet');
  assert.equal(pendingStatuses(e)[0], 'completed');

  const retry = submit(e);
  assert.equal(retry.duplicate, true);
  assert.equal(retry.waLink, first.waLink);
  assert.equal(results(e).length, 2, 'the retry did not append');
});

test('mixed-language answers use the index of the language they were answered in', () => {
  const e = env();
  const mixed = answers(30, 'he').map((a, i) => (i < 10 ? a : answers(30, 'ar')[i]));
  submit(e, { answers: mixed, language: 'ar', languageHistory: ['he', 'ru', 'ar'] });
  const row = results(e)[1];
  assert.equal(row[5], '30/30', 'every answer scored against its own language');
  assert.equal(row[28], 'he → ru → ar');

  // The same selections all claimed as Hebrew are wrong for the Arabic half.
  const e2 = env();
  submit(e2, { answers: mixed.map(a => Object.assign({}, a, { langAtAnswer: 'he' })) });
  assert.equal(results(e2)[1][5], '10/30');
});

test('an unanswered question is reported with the text the examinee saw', () => {
  const e = env();
  const list = answers(30);
  list[0].selected = null;
  submit(e, { answers: list });
  const row = results(e)[1];
  assert.equal(row[5], '29/30');
  assert.match(row[15], /שאלה: שאלה 101/);
  assert.match(row[15], /תשובת הנבחן: לא נענתה/);
  assert.match(row[15], /תשובה נכונה: [אבגד] - תשובה 101-/);
  assert.match(row[15], /קטגוריה: חוק/);
});

test('a registered exam without answers is refused', () => {
  for (const list of [undefined, []]) {
    const e = env();
    const reply = submit(e, { answers: list, score: 30, total: 30, passed: true });
    assert.equal(reply.status, 'error');
    assert.match(reply.message, /חסרות תשובות/);
    assert.equal(results(e).length, 1);
    assert.equal(pendingStatuses(e)[0], 'in_exam', 'the examinee is left in the exam');
  }
});

test('a map entry the key cannot verify keeps the whole result unverified', () => {
  const e = env({ exams: [examRow({ 2: JSON.stringify(questionMap().map((entry, i) => (i === 0 ? { qIdx: 0, correctShuffledIdx: 1 } : entry))) })] });
  submit(e);
  const row = results(e)[1];
  assert.equal(row[22], '', 'not מאומת');
  assert.match(row[15], /^⚠️ ציון לא אומת/);
  assert.equal(row[5], '29/30', 'the unverifiable entry cannot be correct');
});

test('a missing or malformed registration is stored for manual review', () => {
  const missing = env({ exams: [] });
  assert.equal(submit(missing, { score: 27, total: 30, percent: 90, passed: true }).status, 'ok');
  assert.equal(results(missing)[1][5], '27/30', 'the client figure is kept — and flagged');
  assert.equal(results(missing)[1][22], '');
  assert.match(results(missing)[1][15], /^⚠️ ציון לא אומת/);

  const malformed = env({ exams: [examRow({ 2: '{bad-json' })] });
  assert.equal(submit(malformed, { score: 27, total: 30, percent: 90, passed: true }).status, 'ok');
  assert.equal(results(malformed)[1][22], '');
  assert.match(results(malformed)[1][15], /^⚠️ ציון לא אומת/);
  assert.equal(submit(env({ exams: [examRow({ 2: '{bad-json' })] }), { answers: [] }).status, 'error',
    'a registration that exists still makes answers mandatory');
});

test('an empty registered map is refused and can never pass', () => {
  const e = env({ exams: [examRow({ 2: '[]' })] });
  const reply = submit(e, { score: 30, passed: true });
  assert.equal(reply.code, 'invalid_registration');
  assert.equal(results(e).length, 1);
});

test('the latest registration of the session wins', () => {
  const stale = questionMap().map(entry => Object.assign({}, entry, { correctShuffledIdx: 3, shuffleOrder: [3, 2, 1, 0] }));
  const e = env({ exams: [examRow({ 2: JSON.stringify(stale), 3: '2026-09-22T05:00:00Z' }), examRow()] });
  submit(e);
  assert.equal(results(e)[1][5], '30/30', 'scored against the newest map');
});

test('a recent registration makes a fast finish suspicious', () => {
  const e = env({ exams: [examRow({ 3: '2026-09-22T06:29:00Z' })] });
  submit(e);
  assert.equal(results(e)[1][23], 'חשוד');
  const slow = env();
  submit(slow);
  assert.equal(results(slow)[1][23], '');
});

for (const marker of ['סגירת דפדפן', 'טיימאאוט', 'סיום ידני בעקבות ניתוק']) {
  test('a fabricated failure is superseded across languages: ' + marker, () => {
    const e = env({ pending: [pendingRow('completed')], results: [resultRow({ 15: marker, 12: 'ru' })] });
    submit(e);
    assert.equal(results(e)[1][7], 'בוטל');
    assert.equal(results(e)[1][26], 'בוטל אוטומטית — הנבחן השלים והגיש מבחן');
    assert.equal(results(e)[1][27], e.ctx.todayStr());
    assert.equal(results(e).length, 3);
    assert.equal(results(e)[2][14], 1, 'a cancelled fabricated failure is not an attempt');
  });
}

test('a genuine previous result is returned as a duplicate, not superseded', () => {
  const e = env({ pending: [pendingRow('completed')], results: [resultRow()] });
  const reply = submit(e);
  assert.equal(reply.duplicate, true);
  assert.equal(reply.waLink, 'synthetic-existing-link');
  assert.equal(results(e).length, 2);
  assert.equal(results(e)[1][7], 'נכשל', 'the genuine row is untouched');
});

test('a retake after a disqualification: history, voiding and every pending row', () => {
  const e = env({
    pending: [pendingRow('approved'), pendingRow('cancelled'), pendingRow()],
    results: [
      resultRow({ 13: 'OLDER', 7: 'עבר' }),
      resultRow({ 4: 'C1', 7: 'עבר' }),
      resultRow({ 7: 'בוטל' }),
      resultRow({ 7: 'פסול', 17: true })
    ]
  });
  submit(e);
  const rows = results(e);
  assert.equal(rows.at(-1)[14], 3, 'B attempts: the older session and the voided-DQ row, plus this one');
  assert.equal(rows[4][7], 'בוטל', 'the פסול row is voided');
  assert.equal(rows[4][17], false);
  assert.equal(rows[4][26], 'בוטל אוטומטית — נבחן ניגש למבחן מחדש');
  assert.deepEqual(pendingStatuses(e), ['completed', 'cancelled', 'completed'], 'every active row is closed');
});

test('a reset examinee can still deliver a genuine result, a disqualified one cannot', () => {
  const cancelled = env({ pending: [pendingRow('cancelled')] });
  assert.equal(submit(cancelled).status, 'ok');
  assert.equal(results(cancelled).length, 2);

  for (const status of ['disqualified', 'dq_confirmed', 'rejected', 'waiting']) {
    const e = env({ pending: [pendingRow(status)] });
    const reply = submit(e);
    assert.equal(reply.status, 'error', status);
    assert.match(reply.message, /לא מאושר/);
    assert.equal(results(e).length, 1, status + ' writes nothing');
  }
});

test('startExam registers the exam language and a key-derived shuffle', () => {
  const e = env({ exams: [], pending: [pendingRow('approved', { 6: 'ru' })] });
  const reply = e.json(e.ctx.doPost({ postData: { contents: JSON.stringify({
    action: 'startExam', origin: 'examinee-app', sessionCode: SESSION, idNumber: ID,
    examineeToken: TOKEN, language: 'ru', license: 'B' }) } }));
  assert.equal(reply.status, 'ok');
  const row = e.rows('מבחנים')[1];
  assert.equal(row[0], SESSION);
  assert.equal(row[1], ID);
  assert.equal(row[4], 'ru', 'the registration language is stored');
  assert.equal(row[5], 0, 'nothing unverified was registered');
  const map = JSON.parse(row[2]);
  assert.equal(map.length, 30);
  for (const entry of map) {
    assert.equal(entry.shuffleOrder.indexOf(e.ctx.lookupCorrectIndex(entry.qId, 'ru')), entry.correctShuffledIdx,
      'the stored index matches the Russian key');
  }
  // …and it is really a different key from the Hebrew one.
  const differing = map.filter(entry =>
    entry.shuffleOrder.indexOf(e.ctx.lookupCorrectIndex(entry.qId, 'he')) !== entry.correctShuffledIdx);
  assert.ok(differing.length > 0, 'the Russian key is not the Hebrew key');
});

test('a pending row that appears while the result is being written is still closed', () => {
  const e = env();
  const pending = e.sheet('ממתינים');
  const appendRow = e.sheet('תוצאות').appendRow.bind(e.sheet('תוצאות'));
  e.sheet('תוצאות').appendRow = row => { pending.rows.push(pendingRow('approved')); appendRow(row); };
  submit(e);
  assert.deepEqual(pendingStatuses(e), ['completed', 'completed']);
});

test('rows that move while the submit runs are re-read, not overwritten blindly', () => {
  const e = env({ pending: [pendingRow('in_exam', { 0: 'OTHER', 1: '900000002' }), pendingRow()] });
  const pending = e.sheet('ממתינים');
  const appendRow = e.sheet('תוצאות').appendRow.bind(e.sheet('תוצאות'));
  e.sheet('תוצאות').appendRow = row => {
    const tmp = pending.rows[1]; pending.rows[1] = pending.rows[2]; pending.rows[2] = tmp;   // same size, different order
    appendRow(row);
  };
  submit(e);
  assert.deepEqual(pendingStatuses(e), ['completed', 'in_exam'], 'the other examinee is untouched');
  assert.equal(pending.rows[2][1], '900000002');
});

test('an examiner decision taken during the submit is not overwritten', () => {
  const e = env({ results: [resultRow({ 15: 'סגירת דפדפן' })] });
  const pending = e.sheet('ממתינים');
  const appendRow = e.sheet('תוצאות').appendRow.bind(e.sheet('תוצאות'));
  e.sheet('תוצאות').appendRow = row => { pending.rows[1][5] = 'completed'; appendRow(row); };
  submit(e);
  assert.equal(results(e).length, 3);
  assert.equal(pendingStatuses(e)[0], 'completed');
});

test('many duplicate pending rows are all completed, without per-row reads', () => {
  const e = env({ pending: Array.from({ length: 8 }, (_, i) => pendingRow(i === 0 ? 'cancelled' : 'in_exam')) });
  e.resetCounters();
  submit(e);
  assert.deepEqual(pendingStatuses(e), ['cancelled', ...Array(7).fill('completed')]);
  // Past four rows of one examinee, refreshing them one by one costs more round
  // trips than re-reading the sheet — the refresh switches to a single read.
  assert.equal(e.counters().perSheet['ממתינים'].rangeReads, 0);
  assert.ok(e.counters().perSheet['ממתינים'].fullReads <= 3, e.counters().perSheet['ממתינים'].fullReads + ' reads');
});

test('voided rows never count as attempts', () => {
  const e = env({ results: Array.from({ length: 8 }, (_, i) => resultRow({ 13: 'OLD-' + i, 7: i === 0 ? 'בוטל' : 'נכשל', 8: '0' + i + ':00' })) });
  submit(e);
  assert.equal(results(e).at(-1)[14], 8, 'seven genuine failures plus this attempt');
});

test('a big history is read from the tail, and never from מבחנים', () => {
  const oldPending = Array.from({ length: 3999 }, (_, i) =>
    pendingRow('completed', { 0: 'HISTORY', 1: '8' + String(i).padStart(8, '0'), 4: '2026-08-01T08:00:00Z' }));
  const oldExams = Array.from({ length: 3999 }, (_, i) => ['HISTORY', '8' + String(i).padStart(8, '0'), '[]', '2026-08-01T08:00:00Z', 'he', 0]);
  const oldResults = Array.from({ length: 5000 }, (_, i) =>
    resultRow({ 0: '01/08/2026 08:00', 13: 'HISTORY', 1: '8' + String(i).padStart(8, '0') }));
  // An earlier attempt of THIS examinee, far above any 1,000-row tail.
  oldResults[0] = resultRow({ 0: '01/08/2026 08:00', 13: 'OLD-ATTEMPT', 7: 'נכשל' });
  const e = env({ pending: [...oldPending, pendingRow()], exams: [...oldExams, examRow()], results: oldResults });
  e.resetCounters();
  submit(e);
  const counters = e.counters().perSheet;
  assert.equal(results(e).length, 5002, 'the result was appended');
  assert.equal(results(e).at(-1)[5], '30/30');
  assert.equal(counters['מבחנים'].fullReads, 0, 'the question maps of 4,000 exams are never pulled');
  assert.ok(counters['מבחנים'].cellsRead <= 2 * 4000 + 10, 'only the id columns and one row: ' + counters['מבחנים'].cellsRead);
  assert.equal(results(e).at(-1)[14], 2, 'an attempt above the tail is still counted');
  assert.equal(counters['תוצאות'].fullReads, 0, 'a 5,000-row sheet is read as a tail');
  // One 1,000-row tail (30 columns) plus the three attempt columns of the whole
  // sheet — a full read would be 150,000 cells.
  assert.ok(counters['תוצאות'].cellsRead <= 1001 * 30 + 3 * 5001 + 60, 'one tail + attempt columns: ' + counters['תוצאות'].cellsRead);
  assert.equal(counters['ממתינים'].fullReads, 0);
  assert.equal(pendingStatuses(e).at(-1), 'completed', 'the live row was found inside the tail');
});
