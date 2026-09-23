// Run: node tests/scoring_real.test.cjs
//
// The exam-critical path against the REAL answer key (deployment/answer_key.gs)
// and the REAL question banks (deployment/generated/questions_*.json): draw 30
// questions, register them, score what the examinee sent back, write one row.
//
// Every check here exists because something went wrong once:
//   * 03/06/2026 — questions with no answer key scored 0/30 silently, so a draw
//     now skips any id the key cannot answer and a short map is refused (E S1).
//   * 20/09/2026 — en/fr/es/ar order their answers differently, so each answer
//     is scored against the key of the language it was ANSWERED in.
//   * 21/09/2026 — Amharic 1442/1443 were scored against a Hebrew index.
//   * the client's claimed score, its claimed correct index and its claimed
//     wrong-answer list are all ignored: the server re-scores from the key (S4).
//   * cost is a correctness property: the submit used to move ~15MB.
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const { createEnv, ROOT } = require('./helpers/server_env.cjs');

// ---- real data -------------------------------------------------------------
const keySandbox = {};
vm.runInNewContext(fs.readFileSync(path.join(ROOT, 'deployment', 'answer_key.gs'), 'utf8'), keySandbox);
const ANSWER_KEY = keySandbox.ANSWER_KEY_BY_LANG;
const INDEX = JSON.parse(fs.readFileSync(path.join(ROOT, 'deployment', 'question_index.json'), 'utf8'));
const ANSWER_KEY_SOURCE = path.join('deployment', 'answer_key.gs');
const LANGS = ['he', 'ru', 'en', 'ar', 'fr', 'es', 'am'];
const LICENSES = ['B', '1', 'C1', 'C', 'D'];
const BLUEPRINT = {
  B: { 'בטיחות': 7, 'הכרת הרכב': 7, 'חוק': 7, 'תמרורים': 9 },
  1: { 'בטיחות': 5, 'הכרת הרכב': 5, 'חוק': 6, 'תמרורים': 6, 'ספציפי': 8 },
  C1: { 'בטיחות': 5, 'הכרת הרכב': 5, 'חוק': 5, 'תמרורים': 5, 'ספציפי': 10 },
  C: { 'בטיחות': 5, 'הכרת הרכב': 4, 'חוק': 3, 'תמרורים': 4, 'ספציפי': 14 },
  D: { 'בטיחות': 4, 'הכרת הרכב': 2, 'חוק': 5, 'תמרורים': 4, 'ספציפי': 15 }
};
const banks = {};
function bank(lang) {
  if (!banks[lang]) {
    const rows = JSON.parse(fs.readFileSync(path.join(ROOT, 'deployment', 'generated', `questions_${lang}.json`), 'utf8'));
    const byId = {};
    for (const row of rows) if (!byId[row.id]) byId[row.id] = row;
    banks[lang] = byId;
  }
  return banks[lang];
}

// ---- fixtures --------------------------------------------------------------
const SESSION = 'TEST1234';
const ID = '900000001';
const TOKEN = 'examinee-token-1';
const REGISTERED_AT = '2026-09-22T06:00:00Z';
const PENDING_HEADER = ['קוד סשן', 'ת.ז.', 'שם', 'טלפון', 'זמן הרשמה', 'סטטוס', 'שפה', 'אוכלוסיה', 'דרגה', 'שמע',
  'הארכת זמן', 'התחלת מבחן', 'טוקן נבחן', 'ספירת DQ', 'מסך נוסף', 'ספירת אזהרות', 'אזהרה אחרונה', 'אתר', 'סיים במכשיר'];
const RESULTS_HEADER = ['תאריך', 'ת.ז.', 'שם', 'טלפון', 'דרגה', 'ציון', 'אחוז', 'עבר/נכשל', 'זמן', 'בוחן', 'אתר',
  'כיתה', 'שפה', 'קוד סשן', 'ניסיון', 'פירוט שגויות', 'נשלח?', 'פסול?', 'קישור וואטסאפ', 'אוכלוסיה', 'תוקן?', 'שמע',
  'מאומת', 'חשוד', 'dqEventId', 'תוקן ע"י', 'סיבת תיקון', 'תאריך תיקון', 'מסלול שפות', 'מכשיר'];
const EXAMS_HEADER = ['קוד סשן', 'ת.ז.', 'שאלות JSON', 'זמן רישום', 'שפה', 'שגויות לא מאומתות'];

function pendingRow(overrides) {
  const row = Array(19).fill('');
  row[0] = SESSION; row[1] = ID; row[2] = 'נבחן בדיקה'; row[3] = '0501234567';
  row[4] = REGISTERED_AT; row[5] = 'approved'; row[6] = 'he'; row[8] = 'B'; row[9] = 'off'; row[12] = TOKEN;
  return Object.assign(row, overrides || {});
}
function resultRow(overrides) {
  const row = Array(30).fill('');
  row[0] = '22/09/2026 07:10'; row[1] = ID; row[2] = 'נבחן בדיקה'; row[4] = 'B';
  row[5] = '0/30'; row[6] = '0%'; row[7] = 'נכשל'; row[8] = '00:30'; row[12] = 'he'; row[13] = SESSION; row[14] = 1;
  return Object.assign(row, overrides || {});
}
function env(options) {
  const opts = options || {};
  return createEnv({
    sheets: Object.assign({
      'ממתינים': [PENDING_HEADER, ...(opts.pending || [pendingRow()])],
      'מבחנים': [EXAMS_HEADER, ...(opts.exams || [])],
      'תוצאות': [RESULTS_HEADER, ...(opts.results || [])]
    }, opts.sheets || {}),
    // startExam refuses to write anything when the Worker that serves the
    // question texts is unset (DESIGN §11.2), so the fixture is configured.
    properties: Object.assign({ GATEWAY_KEY: 'gateway-secret', GATEWAY_URL: 'https://gw.example.workers.dev' },
      opts.properties || {}),
    sources: [ANSWER_KEY_SOURCE]
  });
}
const post = (e, body) => e.json(e.ctx.doPost({ postData: { contents: JSON.stringify(Object.assign({ origin: 'examinee-app' }, body)) } }));
const get = (e, params) => e.json(e.ctx.doGet({ parameter: Object.assign({ origin: 'examinee-app' }, params) }));
const startExam = (e, overrides) => post(e, Object.assign({
  action: 'startExam', sessionCode: SESSION, idNumber: ID, examineeToken: TOKEN, language: 'he', license: 'B'
}, overrides));

// A registration written straight into 'מבחנים', so a scoring test can pin the
// exact ids, languages and shuffles it needs.
function registration(ids, lang, options) {
  const opts = options || {};
  const map = ids.map((id, i) => {
    const order = opts.orders ? opts.orders[i] : [2, 0, 3, 1];
    const key = ANSWER_KEY[lang][id];
    return { qIdx: i, qId: id, shuffleOrder: order, correctShuffledIdx: order.indexOf(key), topic: opts.topic || 'חוק' };
  });
  return [SESSION, ID, JSON.stringify(map), opts.at || REGISTERED_AT, lang, 0];
}
function mapOf(examRow) { return JSON.parse(examRow[2]); }

// What the client sends back: the texts as DISPLAYED plus the chosen position.
function answerFor(entry, lang, options) {
  const opts = options || {};
  const question = bank(lang)[entry.qId];
  const key = entry.shuffleOrder.indexOf(ANSWER_KEY[lang][entry.qId]);
  const answer = {
    qIdx: entry.qIdx,
    selected: opts.selected !== undefined ? opts.selected : (opts.wrong ? (key + 1) % 4 : key),
    langAtAnswer: lang
  };
  if (!opts.noText && question) {
    answer.q = question.text;
    answer.a = entry.shuffleOrder.map(orig => question.answers[orig]);
  }
  return answer;
}
function answersFor(map, lang, options) {
  const opts = options || {};
  return map.map((entry, i) => answerFor(entry, lang, Object.assign({}, opts, { wrong: i >= (opts.correctCount === undefined ? map.length : opts.correctCount) })));
}
const submit = (e, answers, overrides) => post(e, Object.assign({
  action: 'submitResult', sessionCode: SESSION, idNumber: ID, examineeToken: TOKEN,
  fullName: 'נבחן בדיקה', phone: '0501234567', license: 'B', language: 'he',
  score: 0, total: 30, percent: 0, passed: false, time: '12:34',
  examinerName: 'בוחן בדיקה', site: 'אתר בדיקה', classroom: '1', population: 'חובה',
  audioMode: 'off', device: 'desktop', languageHistory: ['he'], answers: answers
}, overrides));
const lastResult = e => e.rows('תוצאות').at(-1);

// ---- the draw --------------------------------------------------------------
test('the draw follows the blueprint for every license and language', () => {
  const e = env();
  for (const license of LICENSES) {
    for (const lang of LANGS) {
      const drawn = e.ctx.drawExamIds(license, lang);
      assert.equal(drawn.length, 30, `${license}/${lang} draws 30`);
      const byTopic = {}, seen = new Set();
      for (const q of drawn) {
        byTopic[q.topic] = (byTopic[q.topic] || 0) + 1;
        assert.equal(seen.has(q.id), false, `${license}/${lang} id ${q.id} drawn twice`);
        seen.add(q.id);
        assert.equal(INDEX[String(q.id)].c[license], q.topic, 'topic comes from the index');
        assert.ok(INDEX[String(q.id)].l & (1 << LANGS.indexOf(lang)), `${q.id} exists in ${lang}`);
        assert.notEqual(e.ctx.answerKeyIndex(q.id, lang), null, `${q.id} has a ${lang} answer key`);
      }
      assert.deepEqual(byTopic, BLUEPRINT[license], `${license}/${lang} blueprint`);
    }
  }
});

test('an unknown license is refused rather than drawn from', () => {
  const e = env();
  assert.throws(() => e.ctx.drawExamIds('X', 'he'), err => err.code === 'bank_unavailable');
  assert.equal(startExam(e, { license: 'X' }).code, 'unknown_license');
});

// ---- startExam -------------------------------------------------------------
test('startExam registers one map and hands the same one back on a retry', () => {
  const e = env();
  const first = startExam(e);
  assert.equal(first.status, 'ok');
  assert.equal(first.questions.length, 30);
  assert.equal(first.build, e.ctx.THEORY_API_BUILD);
  assert.equal(first.registeredAt, e.rows('מבחנים')[1][3]);
  assert.equal(e.rows('מבחנים').length, 2, 'one registration row');
  for (const q of first.questions) {
    assert.deepEqual([...q.order].sort(), [0, 1, 2, 3], 'the answer order is a permutation');
    assert.ok(q.topic, 'every question carries its topic');
  }

  const cached = startExam(e);
  assert.deepEqual(cached.questions, first.questions, 'a retry is served from the cache');
  assert.equal(e.rows('מבחנים').length, 2, 'no second registration');

  e.entries.clear();                                   // cache evicted mid-exam
  const fromSheet = startExam(e);
  assert.deepEqual(fromSheet.questions, first.questions, 'a retry falls back to the registered row');
  assert.equal(e.rows('מבחנים').length, 2, 'still one registration');
  assert.equal(e.sheet('מבחנים').fullReads, 0, 'the map is never found by reading the whole sheet');
});

test('startExam flips approved to in_exam and drops the poller snapshot', () => {
  const e = env();
  e.cache.put(e.ctx.pendingSnapshotKey(SESSION), '[["stale"]]', 60);
  startExam(e);
  assert.equal(e.rows('ממתינים')[1][5], 'in_exam');
  assert.ok(e.rows('ממתינים')[1][11], 'the exam start time is stamped');
  assert.equal(e.cache.get(e.ctx.pendingSnapshotKey(SESSION)), null, 'the snapshot the pollers read is invalidated');

  const startedAt = e.rows('ממתינים')[1][11];
  startExam(e);
  assert.equal(e.rows('ממתינים')[1][11], startedAt, 'a retry does not restart the clock');
});

test('startExam refuses an examinee the examiner has not approved', () => {
  for (const status of ['waiting', 'completed', 'cancelled', 'rejected', 'disqualified']) {
    const e = env({ pending: [pendingRow({ 5: status })] });
    const reply = startExam(e);
    assert.equal(reply.code, 'not_approved', status);
    assert.equal(e.rows('מבחנים').length, 1, status + ' registers nothing');
  }
});

test('startExam returns the approved duration, the extra minutes and the audio flag', () => {
  const e = env({
    pending: [pendingRow({ 9: 'on', 10: '1.5' })],
    sheets: { 'הארכות זמן': [['תאריך', 'קוד סשן', 'ת.ז.', 'שם', 'דקות', 'סיבה', 'בוחן'], ['', SESSION, ID, '', 7, 'פינוי', '']] }
  });
  const reply = startExam(e);
  assert.equal(reply.examMinutes, 60, '40 × 1.5');
  assert.equal(reply.extraMinutes, 7);
  assert.equal(reply.audioMode, 'on');
  assert.equal(reply.license, 'B');
  assert.equal(reply.language, 'he');
});

test('a reset examinee who registered again is drawn a NEW exam', () => {
  const e = env();
  const first = startExam(e);
  // The examiner reset them; the client registered again → a new row, a new
  // registration time. The old map must not come back.
  e.rows('ממתינים')[1][5] = 'cancelled';
  e.rows('ממתינים').push(pendingRow({ 4: '2026-09-22T07:30:00Z', 5: 'approved', 12: 'examinee-token-2' }));
  const retake = startExam(e, { examineeToken: 'examinee-token-2' });
  assert.equal(retake.status, 'ok');
  assert.equal(e.rows('מבחנים').length, 3, 'the retake is registered separately');
  assert.notDeepEqual(retake.questions.map(q => q.id), first.questions.map(q => q.id));
});

test('startExam costs one pending read, one append and no read of מבחנים', () => {
  const e = env();
  e.resetCounters();
  startExam(e);
  const counters = e.counters().perSheet;
  assert.equal(counters['ממתינים'].fullReads, 1, 'the token check and the handler share one read');
  assert.equal(counters['ממתינים'].setValues, 2, 'status + exam start');
  assert.equal(counters['מבחנים'].fullReads, 0);
  assert.equal(counters['מבחנים'].rangeReads, 0, 'a fresh draw does not look for an old map');
  assert.equal(counters['מבחנים'].appends, 1);

  e.entries.clear();
  e.resetCounters();
  startExam(e);
  const retry = e.counters().perSheet;
  assert.equal(retry['מבחנים'].fullReads, 0);
  assert.ok(retry['מבחנים'].rangeReads <= 2, 'a cache miss reads the id columns and one row: ' + retry['מבחנים'].rangeReads);
  assert.equal(retry['מבחנים'].appends, 0);
});

// ---- scoring ---------------------------------------------------------------
test('a Hebrew exam is scored from the answer key, not from the client', () => {
  const ids = Object.keys(INDEX).slice(0, 30).map(Number);
  for (const [correct, verdict] of [[30, 'עבר'], [26, 'עבר'], [25, 'נכשל'], [0, 'נכשל']]) {
    const e = env({ exams: [registration(ids, 'he')], pending: [pendingRow({ 5: 'in_exam' })] });
    const map = mapOf(e.rows('מבחנים')[1]);
    // The client claims a perfect pass every time; the server must ignore it.
    const reply = submit(e, answersFor(map, 'he', { correctCount: correct }), { score: 30, total: 30, percent: 100, passed: true });
    assert.equal(reply.status, 'ok');
    const row = lastResult(e);
    assert.equal(row[5], correct + '/30', 'score');
    assert.equal(row[7], verdict, correct + ' correct → ' + verdict);
    assert.equal(row[6], Math.round((correct / 30) * 100) + '%');
    assert.equal(row[22], 'מאומת', 'scored against the real key');
  }
});

test('a selection that is not a real answer never counts', () => {
  const ids = Object.keys(INDEX).slice(0, 30).map(Number);
  for (const selected of [null, '', undefined, -1, 'x', 99, 1.5, NaN]) {
    const e = env({ exams: [registration(ids, 'he')], pending: [pendingRow({ 5: 'in_exam' })] });
    const map = mapOf(e.rows('מבחנים')[1]);
    const answers = answersFor(map, 'he', { correctCount: 30 });
    answers[0].selected = selected;
    submit(e, answers);
    assert.equal(lastResult(e)[5], '29/30', 'selected=' + String(selected) + ' is unanswered');
  }
  // An answer object that is missing entirely is unanswered too.
  const e = env({ exams: [registration(ids, 'he')], pending: [pendingRow({ 5: 'in_exam' })] });
  const map = mapOf(e.rows('מבחנים')[1]);
  const answers = answersFor(map, 'he', { correctCount: 30 });
  answers[3] = null;
  submit(e, answers);
  assert.equal(lastResult(e)[5], '29/30');
});

test('a mid-exam language switch is scored against the language answered in', () => {
  // ids whose English answer order differs from the Hebrew one — with a single
  // key per question, answering in English would score these wrong.
  const differing = Object.keys(INDEX).map(Number)
    .filter(id => (INDEX[String(id)].l & 0b101) === 0b101 && ANSWER_KEY.he[id] !== undefined &&
      ANSWER_KEY.en[id] !== undefined && ANSWER_KEY.he[id] !== ANSWER_KEY.en[id]);
  assert.ok(differing.length >= 30, 'the banks still disagree on answer order');
  const ids = differing.slice(0, 30);
  const e = env({ exams: [registration(ids, 'he')], pending: [pendingRow({ 5: 'in_exam' })] });
  const map = mapOf(e.rows('מבחנים')[1]);
  // First half answered in Hebrew, second half after switching to English.
  const answers = map.map((entry, i) => answerFor(entry, i < 15 ? 'he' : 'en'));
  submit(e, answers, { language: 'en', languageHistory: ['he', 'en'] });
  const row = lastResult(e);
  assert.equal(row[5], '30/30', 'every answer scored against its own language');
  assert.equal(row[28], 'he → en', 'the language path is recorded');

  // The same answers claimed as Hebrew would be wrong — proving the key differs.
  const e2 = env({ exams: [registration(ids, 'he')], pending: [pendingRow({ 5: 'in_exam' })] });
  const map2 = mapOf(e2.rows('מבחנים')[1]);
  const mislabelled = map2.map((entry, i) => Object.assign(answerFor(entry, i < 15 ? 'he' : 'en'), { langAtAnswer: 'he' }));
  submit(e2, mislabelled);
  assert.equal(lastResult(e2)[5], '15/30');
});

test('Amharic 1442 and 1443 are scored against the Amharic bank (21/09 fix)', () => {
  // 28/03/2026 a rebuild reordered the answers of exactly these two while the
  // key stayed at the Hebrew index; the bank was restored on 21/09/2026.
  assert.equal(ANSWER_KEY.am[1442], ANSWER_KEY.he[1442]);
  assert.equal(ANSWER_KEY.am[1443], ANSWER_KEY.he[1443]);
  const filler = Object.keys(INDEX).map(Number).filter(id => id !== 1442 && id !== 1443 && (INDEX[String(id)].l & 0b1000000)).slice(0, 28);
  const ids = [1442, 1443, ...filler];
  const e = env({ exams: [registration(ids, 'am')], pending: [pendingRow({ 5: 'in_exam', 6: 'am', 8: 'C' })] });
  const map = mapOf(e.rows('מבחנים')[1]);
  const answers = map.map((entry, i) => answerFor(entry, 'am', { wrong: i >= 2 }));
  submit(e, answers, { language: 'am', license: 'C' });
  const row = lastResult(e);
  assert.equal(row[5], '2/30', 'the two Amharic questions scored correct');
  assert.equal(row[15].indexOf('מזהה שאלה: 1442'), -1, '1442 is not listed as wrong');
  assert.equal(row[15].indexOf('מזהה שאלה: 1443'), -1, '1443 is not listed as wrong');
});

test('a short or empty registration is refused and can never pass', () => {
  const ids = Object.keys(INDEX).slice(0, 24).map(Number);
  for (const exams of [[registration(ids, 'he')], [registration([], 'he')]]) {
    const e = env({ exams, pending: [pendingRow({ 5: 'in_exam' })] });
    const stored = mapOf(e.rows('מבחנים')[1]);
    const reply = submit(e, stored.length ? answersFor(stored, 'he', { correctCount: stored.length })
      : [{ qIdx: 0, selected: 0, langAtAnswer: 'he' }], { score: 30, total: 30, passed: true });
    assert.equal(reply.status, 'error');
    if (stored.length) assert.equal(reply.code, 'invalid_registration');
    assert.equal(e.rows('תוצאות').length, 1, 'nothing is written');
  }
});

test('a registered exam without answers is refused', () => {
  const ids = Object.keys(INDEX).slice(0, 30).map(Number);
  const e = env({ exams: [registration(ids, 'he')], pending: [pendingRow({ 5: 'in_exam' })] });
  assert.match(submit(e, [], { score: 30, passed: true }).message, /חסרות תשובות/);
  assert.match(submit(e, undefined, { score: 30, passed: true }).message, /חסרות תשובות/);
  assert.equal(e.rows('תוצאות').length, 1);
});

test('an entry the key cannot verify leaves the whole result unverified', () => {
  const ids = Object.keys(INDEX).slice(0, 30).map(Number);
  const e = env({ exams: [registration(ids, 'he')], pending: [pendingRow({ 5: 'in_exam' })] });
  const examRow = e.rows('מבחנים')[1];
  const map = mapOf(examRow);
  delete map[0].shuffleOrder;                     // a legacy entry: correctShuffledIdx only
  map[0].correctShuffledIdx = 0;
  examRow[2] = JSON.stringify(map);
  const answers = map.map((entry, i) => (i === 0 ? { qIdx: 0, selected: 0, langAtAnswer: 'he' } : answerFor(entry, 'he')));
  submit(e, answers);
  const row = lastResult(e);
  assert.equal(row[5], '29/30', 'the unverifiable answer is not counted as correct');
  assert.equal(row[22], '', 'the row is not marked מאומת');
  assert.match(row[15], /^⚠️ ציון לא אומת/, 'the certificate says so');
});

test('a submit with no registration at all is stored for manual review', () => {
  const e = env({ pending: [pendingRow({ 5: 'in_exam' })] });
  const reply = submit(e, [{ qIdx: 0, selected: 0, langAtAnswer: 'he' }], { score: 28, total: 30, percent: 93, passed: true });
  assert.equal(reply.status, 'ok');
  const row = lastResult(e);
  assert.equal(row[22], '', 'never מאומת');
  assert.match(row[15], /^⚠️ ציון לא אומת/);
});

// ---- the certificate -------------------------------------------------------
test('the texts the examinee saw reach the certificate and the WhatsApp link', () => {
  const ids = Object.keys(INDEX).slice(0, 30).map(Number);
  const e = env({ exams: [registration(ids, 'he')], pending: [pendingRow({ 5: 'in_exam' })] });
  const map = mapOf(e.rows('מבחנים')[1]);
  const answers = answersFor(map, 'he', { correctCount: 29 });
  const reply = submit(e, answers);
  const wrongEntry = map[29], question = bank('he')[wrongEntry.qId];
  const chosen = question.answers[wrongEntry.shuffleOrder[answers[29].selected]];
  const correct = question.answers[ANSWER_KEY.he[wrongEntry.qId]];
  const details = lastResult(e)[15];
  assert.match(details, new RegExp('מזהה שאלה: ' + wrongEntry.qId));
  assert.ok(details.includes(question.text), 'the question text as displayed');
  assert.ok(details.includes(chosen), 'what the examinee chose');
  assert.ok(details.includes(correct), 'what was correct');
  assert.ok(details.includes('קטגוריה: ' + wrongEntry.topic), 'the blueprint topic');
  const waText = decodeURIComponent(reply.waLink.split('?text=')[1]);
  assert.ok(waText.includes(question.text) && waText.includes(correct), 'the WhatsApp message carries them too');
  assert.equal(details.split('מזהה שאלה:').length - 1, 1, 'only the wrong answer is listed');
});

test('an old client that sends no texts is still scored', () => {
  const ids = Object.keys(INDEX).slice(0, 30).map(Number);
  const e = env({ exams: [registration(ids, 'he')], pending: [pendingRow({ 5: 'in_exam' })] });
  const map = mapOf(e.rows('מבחנים')[1]);
  const answers = map.map((entry, i) => answerFor(entry, 'he', { noText: true, wrong: i >= 28 }));
  submit(e, answers);
  const row = lastResult(e);
  assert.equal(row[5], '28/30');
  assert.ok(row[15].includes('(טקסט לא זמין)'), 'the certificate says the text is unavailable');
  assert.equal(row[15].indexOf('undefined'), -1, 'and never the word undefined');
});

test('the appended row matches the sheet header, column for column', () => {
  const ids = Object.keys(INDEX).slice(0, 30).map(Number);
  const e = env({ exams: [registration(ids, 'he')], pending: [pendingRow({ 5: 'in_exam' })] });
  const map = mapOf(e.rows('מבחנים')[1]);
  submit(e, answersFor(map, 'he', { correctCount: 27 }), { time: '11:22', device: 'tablet', audioMode: 'on' });
  const row = lastResult(e);
  assert.deepEqual(Array.from(e.ctx.SHEET_HEADERS['תוצאות']), RESULTS_HEADER, 'the header this test was written against');
  assert.equal(row.length, 30);
  assert.equal(row[1], ID);
  assert.equal(row[4], 'B');
  assert.equal(row[5], '27/30');
  assert.equal(row[7], 'עבר');
  assert.equal(row[8], '11:22');
  assert.equal(row[13], SESSION);
  assert.equal(row[14], 1, 'first attempt');
  assert.equal(row[16], false);
  assert.equal(row[17], false);
  assert.match(row[18], /^https:\/\/wa\.me\/972501234567\?text=/);
  assert.equal(row[20], false);
  assert.equal(row[21], 'on');
  assert.equal(row[22], 'מאומת');
  assert.equal(row[23], '', 'a 12-minute exam is not suspicious');
  assert.equal(row[29], 'tablet');
});

test('an exam finished in under three minutes is flagged', () => {
  const ids = Object.keys(INDEX).slice(0, 30).map(Number);
  const e = env({
    exams: [registration(ids, 'he', { at: '2026-09-22T06:29:00Z' })],   // 60s before the fixed clock
    pending: [pendingRow({ 5: 'in_exam' })]
  });
  const map = mapOf(e.rows('מבחנים')[1]);
  submit(e, answersFor(map, 'he', { correctCount: 30 }), { time: '00:59' });
  assert.equal(lastResult(e)[23], 'חשוד');
});

// ---- the passes over 'תוצאות' ----------------------------------------------
test('a fabricated close-fail is superseded by the real result', () => {
  const ids = Object.keys(INDEX).slice(0, 30).map(Number);
  for (const marker of ['סגירת דפדפן באמצע מבחן', 'טיימאאוט', 'סיום ידני בעקבות ניתוק']) {
    const e = env({
      exams: [registration(ids, 'he')],
      pending: [pendingRow({ 5: 'completed' })],
      results: [resultRow({ 15: marker, 12: 'ru' })]      // written in the REGISTRATION language
    });
    const map = mapOf(e.rows('מבחנים')[1]);
    submit(e, answersFor(map, 'he', { correctCount: 30 }));
    assert.equal(e.rows('תוצאות')[1][7], 'בוטל', marker + ' is voided');
    assert.equal(e.rows('תוצאות')[1][26], 'בוטל אוטומטית — הנבחן השלים והגיש מבחן');
    assert.equal(e.rows('תוצאות').length, 3, 'the real result is appended');
    assert.equal(lastResult(e)[5], '30/30');
    assert.equal(lastResult(e)[14], 1, 'a voided row is not an attempt');
  }
});

test('a previous disqualification is voided when the examinee finishes', () => {
  const ids = Object.keys(INDEX).slice(0, 30).map(Number);
  const e = env({
    exams: [registration(ids, 'he')],
    pending: [pendingRow({ 5: 'in_exam' })],
    results: [resultRow({ 7: 'פסול', 17: true })]
  });
  const map = mapOf(e.rows('מבחנים')[1]);
  submit(e, answersFor(map, 'he', { correctCount: 26 }));
  assert.equal(e.rows('תוצאות')[1][7], 'בוטל');
  assert.equal(e.rows('תוצאות')[1][17], false);
  assert.equal(e.rows('תוצאות')[1][26], 'בוטל אוטומטית — נבחן ניגש למבחן מחדש');
  assert.equal(lastResult(e)[7], 'עבר');
});

test('a genuine earlier result is returned instead of a second row', () => {
  const ids = Object.keys(INDEX).slice(0, 30).map(Number);
  const e = env({
    exams: [registration(ids, 'he')],
    pending: [pendingRow({ 5: 'completed' })],
    results: [resultRow({ 5: '27/30', 7: 'עבר', 18: 'https://wa.me/existing' })]
  });
  const map = mapOf(e.rows('מבחנים')[1]);
  const reply = submit(e, answersFor(map, 'he', { correctCount: 30 }));
  assert.equal(reply.duplicate, true);
  assert.equal(reply.waLink, 'https://wa.me/existing');
  assert.equal(e.rows('תוצאות').length, 2, 'no second row');
  assert.equal(e.rows('ממתינים')[1][5], 'completed');
});

test('the same result sent twice is stored once', () => {
  const ids = Object.keys(INDEX).slice(0, 30).map(Number);
  const e = env({ exams: [registration(ids, 'he')], pending: [pendingRow({ 5: 'in_exam' })] });
  const map = mapOf(e.rows('מבחנים')[1]);
  const answers = answersFor(map, 'he', { correctCount: 26 });
  const first = submit(e, answers);
  const retry = submit(e, answers);
  assert.equal(first.status, 'ok');
  assert.equal(retry.duplicate, true);
  assert.equal(retry.waLink, first.waLink);
  assert.equal(e.rows('תוצאות').length, 2, 'one result row');
  assert.equal(e.rows('ממתינים')[1][5], 'completed');
});

test('a retake after a disqualification is not treated as a duplicate', () => {
  const ids = Object.keys(INDEX).slice(0, 30).map(Number);
  const e = env({
    exams: [registration(ids, 'he')],
    pending: [pendingRow({ 5: 'in_exam' })],
    results: [resultRow({ 5: '20/30', 7: 'נכשל', 8: '20:00' })]
  });
  const map = mapOf(e.rows('מבחנים')[1]);
  submit(e, answersFor(map, 'he', { correctCount: 26 }), { time: '30:00' });
  assert.equal(e.rows('תוצאות').length, 3);
  assert.equal(lastResult(e)[14], 2, 'second attempt');
});

test('submitResult reads תוצאות once and never reads מבחנים whole', () => {
  const ids = Object.keys(INDEX).slice(0, 30).map(Number);
  const e = env({ exams: [registration(ids, 'he')], pending: [pendingRow({ 5: 'in_exam' })] });
  const map = mapOf(e.rows('מבחנים')[1]);
  e.resetCounters();
  submit(e, answersFor(map, 'he', { correctCount: 26 }));
  const counters = e.counters().perSheet;
  assert.equal(counters['תוצאות'].fullReads, 1, 'one snapshot serves supersede, duplicate and idempotency');
  assert.equal(counters['תוצאות'].appends, 1);
  assert.equal(counters['מבחנים'].fullReads, 0, 'the question maps of every exam ever are never read');
  assert.ok(counters['מבחנים'].rangeReads <= 2, 'id columns + one row');
  assert.equal(counters['ממתינים'].fullReads, 1, 'the token check and the completion share one read');
});

test('a submit from an examinee who is not in this exam is refused', () => {
  const ids = Object.keys(INDEX).slice(0, 30).map(Number);
  const e = env({ exams: [registration(ids, 'he')], pending: [pendingRow({ 5: 'disqualified' })] });
  const map = mapOf(e.rows('מבחנים')[1]);
  assert.match(submit(e, answersFor(map, 'he', { correctCount: 30 })).message, /לא מאושר/);
  assert.equal(e.rows('תוצאות').length, 1);
});

test('a wrong examinee token never reaches the handler', () => {
  const ids = Object.keys(INDEX).slice(0, 30).map(Number);
  const e = env({ exams: [registration(ids, 'he')], pending: [pendingRow({ 5: 'in_exam' })] });
  const map = mapOf(e.rows('מבחנים')[1]);
  const reply = submit(e, answersFor(map, 'he', { correctCount: 30 }), { examineeToken: 'forged' });
  assert.equal(reply.examineeTokenError, 'mismatch');
  assert.equal(e.rows('תוצאות').length, 1);
  assert.equal(startExam(e, { examineeToken: 'forged' }).examineeTokenError, 'mismatch');
});

// ---- the close beacon ------------------------------------------------------
test('the close beacon writes one fail, and a reset examinee gets none', () => {
  const e = env({ pending: [pendingRow({ 5: 'in_exam' })] });
  const body = { action: 'submitFailOnClose', sessionCode: SESSION, idNumber: ID, examineeToken: TOKEN,
    fullName: 'נבחן בדיקה', phone: '0501234567', license: 'B', language: 'he', answeredCount: 12,
    totalQuestions: 30, time: '05:00', device: 'phone' };
  assert.equal(post(e, body).status, 'ok');
  const row = lastResult(e);
  assert.equal(row.length, 30);
  assert.equal(row[5], '0/30');
  assert.equal(row[7], 'נכשל');
  assert.match(row[15], /סגירת דפדפן באמצע מבחן \(נענו 12 שאלות\)/);
  assert.equal(row[29], 'phone');
  assert.equal(e.rows('ממתינים')[1][5], 'completed');
  assert.equal(post(e, body).duplicate, true, 'a second beacon adds nothing');
  assert.equal(e.rows('תוצאות').length, 2);

  const reset = env({ pending: [pendingRow({ 5: 'cancelled' })] });
  assert.equal(post(reset, body).skipped, 'cancelled');
  assert.equal(reset.rows('תוצאות').length, 1);
});

test('a reload cancels the close-fail and puts the examinee back in the exam', () => {
  const e = env({
    pending: [pendingRow({ 5: 'completed' })],
    results: [resultRow({ 15: 'סגירת דפדפן באמצע מבחן (נענו 3 שאלות)' })]
  });
  assert.equal(post(e, { action: 'cancelFailOnClose', sessionCode: SESSION, idNumber: ID, examineeToken: TOKEN }).status, 'ok');
  assert.equal(e.rows('תוצאות')[1][7], 'בוטל');
  assert.equal(e.rows('תוצאות')[1][26], 'בוטל אוטומטי - רענון/חזרה למבחן');
  assert.equal(e.rows('ממתינים')[1][5], 'in_exam', 'the examinee can finish');
});

// ---- practice --------------------------------------------------------------
test('startPractice returns ids, topics and a correct index per language', () => {
  const e = env();
  const reply = get(e, { action: 'startPractice', language: 'he', license: 'B', classCode: 'C1', studentId: 'S1' });
  assert.equal(reply.status, 'ok');
  assert.equal(reply.questions.length, 30);
  const byTopic = {};
  for (const q of reply.questions) {
    byTopic[q.topic] = (byTopic[q.topic] || 0) + 1;
    const entry = INDEX[String(q.id)];
    for (let bit = 0; bit < LANGS.length; bit++) {
      const lang = LANGS[bit];
      if (!(entry.l & (1 << bit))) { assert.equal(q.ci[lang], undefined, `${q.id} is not translated to ${lang}`); continue; }
      assert.equal(q.ci[lang], ANSWER_KEY[lang][q.id] ^ (q.id % 256), `${q.id} ${lang} ci is XOR-encoded`);
    }
  }
  assert.deepEqual(byTopic, BLUEPRINT.B, 'practice in exam mode follows the blueprint');
});

test('startPractice serves a single topic and an explicit id list', () => {
  const e = env();
  const category = get(e, { action: 'startPractice', mode: 'category', categoryFilter: 'תמרורים',
    maxCount: '8', language: 'ru', license: 'C', classCode: 'C1', studentId: 'S1' });
  assert.equal(category.questions.length, 8);
  for (const q of category.questions) {
    assert.equal(q.topic, 'תמרורים');
    assert.equal(INDEX[String(q.id)].c['C'], 'תמרורים');
    assert.ok(INDEX[String(q.id)].l & 0b10, 'exists in Russian');
  }
  const known = Object.keys(INDEX).slice(0, 3).map(Number);
  const byIds = get(e, { action: 'startPractice', mode: 'ids', ids: known.join(',') + ',999999',
    language: 'he', license: 'B', classCode: 'C1', studentId: 'S1' });
  assert.deepEqual(byIds.questions.map(q => q.id), known, 'unknown ids are dropped, not fatal');
  // 30 is the ceiling in EVERY mode: one draw is one /v1/bank request and the
  // Worker reads at most 30 assets per request (DESIGN §11.3).
  const capped = get(e, { action: 'startPractice', mode: 'category', categoryFilter: 'חוק', maxCount: '500',
    language: 'he', license: 'B', classCode: 'C1', studentId: 'S1' });
  assert.equal(capped.questions.length, 30, 'never more than 30 questions in one request');
  const manyIds = get(e, { action: 'startPractice', mode: 'ids', language: 'he', license: 'B',
    classCode: 'C1', studentId: 'S1', ids: Object.keys(INDEX).slice(0, 40).join(',') });
  assert.equal(manyIds.questions.length, 30, 'a spaced-repetition list is capped too');
  // Every draw carries the grant that lets the device fetch those texts.
  for (const reply of [category, byIds, capped, manyIds]) {
    const payload = JSON.parse(Buffer.from(reply.bank.grant.split('.')[0], 'base64url').toString('utf8'));
    assert.equal(payload.s, 'practice');
    assert.equal(payload.sub, 'C1:S1');
    assert.deepEqual(payload.ids, reply.questions.map(q => q.id));
  }
});

test('practice allowances are per identity, and guests are capped globally', () => {
  const e = env();
  const guest = () => get(e, { action: 'startPractice', language: 'he', license: 'B' });
  for (let i = 0; i < 5; i++) assert.equal(guest().status, 'ok');
  assert.equal(guest().rateLimited, true, 'a guest gets five draws a minute');

  const student = n => get(e, { action: 'startPractice', language: 'he', license: 'B', classCode: 'C1', studentId: 'S' + n });
  for (let i = 0; i < 20; i++) assert.equal(student(1).status, 'ok');
  assert.equal(student(1).rateLimited, true);
  assert.equal(student(2).status, 'ok', 'one student does not spend the class allowance');

  const standalone = get(e, { action: 'startPractice', language: 'he', license: 'B', standaloneIdNumber: '900000002' });
  assert.equal(standalone.status, 'ok');
});

// ---- the router ------------------------------------------------------------
test('retired actions tell an old client to refresh, unknown ones are refused', () => {
  const e = env();
  for (const action of ['getExamQuestions']) {
    assert.equal(get(e, { action }).code, 'client_outdated', action);
  }
  assert.equal(post(e, { action: 'registerExamQuestions' }).code, 'client_outdated');
  for (const action of ['getQuestionsByIds', 'searchQuestions', 'uploadResultHtml', 'getUploadResult', 'viewResult',
    'predictiveModelPreview', 'submitWrongAnswers']) {
    assert.match(get(e, { action }).message, /Unknown action/, action);
    assert.match(post(e, { action }).message, /Unknown action/, action + ' (POST)');
  }
  assert.equal(get(e, { action: 'markExamStarted', sessionCode: SESSION, idNumber: ID, examineeToken: TOKEN }).already, true);
  assert.match(get(e, { action: 'submitResult' }).message, /דורשת POST/);
  assert.match(get(e, { action: 'login' }).message, /דורשת POST/);
  assert.match(post(e, { action: 'startPractice' }).message, /דורשת GET/);
});

test('health reports the build and the size of the deployed index', () => {
  const e = env();
  const health = get(e, { action: 'health' });
  assert.equal(health.status, 'ok');
  assert.equal(health.build, '2026-09-24-r33');
  assert.equal(health.indexIds, 1700);
  assert.equal(Object.keys(INDEX).length, 1700, 'the generated index still holds every question');
  assert.equal(e.counters().fullReads, 0, 'health reads nothing');
});
