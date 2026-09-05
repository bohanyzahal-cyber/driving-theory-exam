// Offline regression replay against the last audited production-source commit.
// All records and service responses are synthetic; no network or Google access.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const { execFileSync } = require('node:child_process');

const appDir = path.resolve(__dirname, '..');
const baseSource = execFileSync('git', ['show', '232ed3c:external_exam_apps_script.js'], { cwd: appDir, encoding: 'utf8', maxBuffer: 5e6 });
const currentSource = fs.readFileSync(path.join(appDir, 'external_exam_apps_script.js'), 'utf8');
const clone = value => JSON.parse(JSON.stringify(value));
const ID = '900000001';
const TOKEN = 'synthetic-examinee-token';
const NOW = Date.parse('2026-09-06T08:30:00Z');
const names = { pending: 'ממתינים', exams: 'מבחנים', results: 'תוצאות' };
const questionMap = [
  { qIdx: 0, qId: 1, shuffleOrder: [0, 1, 2, 3], correctShuffledIdx: 0 },
  { qIdx: 1, qId: 2, shuffleOrder: [0, 1, 2, 3], correctShuffledIdx: 1 }
];
const correctIndexes = { he: { 1: 0, 2: 1 }, ru: { 1: 2, 2: 0 }, ar: { 1: 3, 2: 2 } };

function pendingRow(status = 'in_exam', overrides = {}) {
  return Object.assign(Array(19).fill(''), {
    0: 'SYNTHETIC', 1: ID, 4: '2026-09-06T08:00:00Z', 5: status,
    6: 'he', 8: 'B', 9: 'off', 12: TOKEN
  }, overrides);
}
function examRow(overrides = {}) {
  return Object.assign(['SYNTHETIC', ID, JSON.stringify(questionMap), '2026-09-06T08:00:00Z', 'he', 0], overrides);
}
function resultRow(overrides = {}) {
  return Object.assign(Array(30).fill(''), {
    0: '06/09/2026 08:10', 1: ID, 2: 'Synthetic examinee', 4: 'B',
    5: '0/2', 6: '0%', 7: 'נכשל', 8: '00:30', 12: 'he',
    13: 'SYNTHETIC', 14: 1, 18: 'synthetic-existing-link'
  }, overrides);
}
function input(overrides = {}) {
  return Object.assign({
    sessionCode: 'SYNTHETIC', idNumber: ID, examineeToken: TOKEN,
    fullName: 'Synthetic examinee', phone: '000', license: 'B', language: 'he',
    answers: [{ selected: 0, langAtAnswer: 'he' }, { selected: 1, langAtAnswer: 'he' }],
    questions: questionMap.map(q => ({ qIdx: q.qIdx, qId: q.qId, shuffleOrder: q.shuffleOrder })),
    score: 0, total: 2, percent: 0, passed: false, time: '10:00',
    examinerName: 'Synthetic examiner', site: 'Synthetic site', classroom: 'Synthetic classroom',
    population: 'Synthetic population', audioMode: 'on', device: 'tablet',
    languageHistory: ['he'], wrongAnswers: [{ question: 'client placeholder', yourAnswer: 'undefined', correctAnswer: 'undefined' }]
  }, overrides);
}
function fixture(overrides = {}) {
  return Object.assign({ pending: [pendingRow()], exams: [examRow()], results: [], input: input() }, overrides);
}

function runtime(source, spec) {
  const tables = {
    [names.pending]: [Array(19).fill('header'), ...clone(spec.pending)],
    [names.exams]: [Array(6).fill('header'), ...clone(spec.exams)],
    [names.results]: [Array(30).fill('header'), ...clone(spec.results)]
  };
  const reads = {}, writes = [], bankLoads = [];
  let run;
  const cache = { get() { return null; }, put() {}, getAll() { return {}; }, putAll() {} };
  function read(name, kind, r, c, h, w) {
    reads[name] ||= { full: 0, range: 0, cells: 0 };
    reads[name][kind]++;
    reads[name].cells += h * w;
    const snapshot = tables[name].slice(r - 1, r - 1 + h).map(row => row.slice(c - 1, c - 1 + w));
    if (spec.afterRead) spec.afterRead({ name, kind, count: reads[name][kind], tables, reads, writes });
    return snapshot;
  }
  class FixedDate extends Date {
    constructor(...args) { super(...(args.length ? args : [NOW])); }
    static now() { return NOW; }
  }
  run = {
    Date: FixedDate, Logger: { log() {} }, CacheService: { getScriptCache: () => cache },
    SpreadsheetApp: { flush() {} }, Session: { getScriptTimeZone: () => 'UTC' },
    Utilities: { formatDate: () => '06/09/2026 08:30' }
  };
  vm.createContext(run);
  vm.runInContext(source, run);
  run.requireRateLimit = () => null;
  run.jsonResponse = result => result;
  run.nowISO = () => '2026-09-06T08:30:00.000Z';
  run.todayStr = () => '06/09/2026 08:30';
  run.ANSWER_KEY_BY_LANG = correctIndexes;
  run.lookupCorrectIndex = (id, lang) => correctIndexes[lang]?.[id] ?? null;
  run.loadQuestionsForLanguageServer = lang => {
    bankLoads.push(lang);
    if (spec.onBankLoad) spec.onBankLoad({ lang, tables, reads, writes });
    if (spec.failLanguages?.includes(lang)) throw new Error('Synthetic unavailable bank');
    return [1, 2].map(id => ({ id, text: `synthetic ${lang} question ${id}`, answers: [0, 1, 2, 3].map(a => `${lang}-${id}-${a}`), category: 'חוק' }));
  };
  run.getSheet = name => {
    if (!tables[name]) throw new Error('Unexpected synthetic table ' + name);
    return {
      getLastRow: () => tables[name].length,
      getLastColumn: () => tables[name][0].length,
      getDataRange: () => ({ getValues: () => read(name, 'full', 1, 1, tables[name].length, tables[name][0].length) }),
      getRange: (r, c, h = 1, w = 1) => ({
        getValues: () => read(name, 'range', r, c, h, w),
        setValue(value) {
          assert(tables[name][r - 1], `write stays within ${name}`);
          tables[name][r - 1][c - 1] = value;
          writes.push({ name, r, c, value });
        }
      }),
      appendRow(row) {
        tables[name].push(clone(row));
        writes.push({ name, append: clone(row) });
        if (spec.onAppend) spec.onAppend({ name, tables, reads, writes });
      }
    };
  };
  return {
    run, tables, reads, writes, bankLoads,
    execute(action, data) { return clone(run[action](clone(data))); }
  };
}

let passed = 0;
const summaries = [];
function compare(label, spec, action = 'handleSubmitResult', options = {}) {
  const oldRun = runtime(baseSource, spec), newRun = runtime(currentSource, spec);
  const before = oldRun.execute(action, spec.input);
  const after = newRun.execute(action, spec.input);
  assert.deepEqual(after, before, label + ': API response matches audited code');
  assert.deepEqual(newRun.tables, oldRun.tables, label + ': complete tables (scoring, metadata, audit, attempts) match');
  if (options.retry) {
    const oldRetry = oldRun.execute(action, spec.input), newRetry = newRun.execute(action, spec.input);
    assert.deepEqual(newRetry, oldRetry, label + ': lost-response retry matches');
    assert.deepEqual(newRun.tables, oldRun.tables, label + ': retry rows match');
  }
  if (options.check) options.check({ oldRun, newRun, result: after });
  summaries.push({ label, oldFullReads: Object.values(oldRun.reads).reduce((n, r) => n + r.full, 0), newFullReads: Object.values(newRun.reads).reduce((n, r) => n + r.full, 0), oldBankLoads: oldRun.bankLoads.length, newBankLoads: newRun.bankLoads.length });
  passed++;
  return { oldRun, newRun, result: after };
}

compare('perfect verified result, metadata and lost-response retry', fixture(), 'handleSubmitResult', {
  retry: true,
  check: ({ newRun }) => {
    assert.equal(newRun.tables[names.results].length, 2);
    assert.equal(newRun.tables[names.results][1][5], '2/2');
    assert.equal(newRun.tables[names.results][1][22], 'מאומת');
    assert.equal(newRun.bankLoads.length, 0, 'perfect score never loads language banks');
  }
});
compare('mixed-language answers use language-at-answer indexes and feedback', fixture({ input: input({
  language: 'ru', languageHistory: ['he', 'ru', 'ar'],
  answers: [{ selected: 2, langAtAnswer: 'ru' }, { selected: 0, langAtAnswer: 'ar' }]
}) }), 'handleSubmitResult', { check: ({ newRun }) => {
  const row = newRun.tables[names.results][1];
  assert.equal(row[5], '1/2');
  assert.match(row[15], /synthetic ar question 2/);
  assert.equal(row[28], 'he → ru → ar');
} });
compare('unanswered question retains wrong-answer feedback', fixture({ input: input({ answers: [null, { selected: 1, langAtAnswer: 'he' }] }) }));
compare('missing translated bank preserves default-language fallback', fixture({ failLanguages: ['ru'], input: input({ answers: [{ selected: 1, langAtAnswer: 'ru' }, { selected: 1, langAtAnswer: 'he' }] }) }));
compare('missing default bank preserves existing client feedback', fixture({ failLanguages: ['he'], input: input({ answers: [{ selected: 1, langAtAnswer: 'he' }, { selected: 1, langAtAnswer: 'he' }] }) }));
compare('registered exam rejects absent answers', fixture({ input: input({ answers: undefined }) }));
compare('registered exam rejects empty answers', fixture({ input: input({ answers: [] }) }));
compare('unverified map keeps verification flag', fixture({ exams: [examRow({ 5: 1 })] }));
compare('missing registration retains manual-review behavior', fixture({ exams: [] }));
compare('malformed latest registration retains verification failure behavior', fixture({ exams: [examRow({ 2: '{bad-json' })] }));
compare('latest registration and timing win over older same-session map', fixture({ exams: [examRow(), examRow({ 2: JSON.stringify(questionMap.map(q => ({ ...q, correctShuffledIdx: 3 }))), 3: '2026-09-06T08:29:00Z' })] }));
for (const marker of ['סגירת דפדפן', 'טיימאאוט', 'סיום ידני בעקבות ניתוק']) {
  compare('fabricated failure superseded across languages: ' + marker, fixture({
    pending: [pendingRow('completed')],
    results: [resultRow({ 15: marker, 12: 'ru' })],
    input: input({ language: 'he' })
  }), 'handleSubmitResult', { check: ({ newRun }) => {
    assert.equal(newRun.tables[names.results][1][7], 'בוטל');
    assert.equal(newRun.tables[names.results][2][14], 1, 'cancelled fabricated failure is excluded from attempts');
  } });
}
compare('genuine previous failure returns duplicate without supersession', fixture({ pending: [pendingRow('completed')], results: [resultRow()] }));
compare('DQ retake, old license history and multiple pending rows', fixture({
  pending: [pendingRow('approved'), pendingRow('cancelled'), pendingRow('in_exam')],
  results: [resultRow({ 13: 'OLDER', 7: 'עבר' }), resultRow({ 4: 'C1', 7: 'עבר' }), resultRow({ 7: 'בוטל' }), resultRow({ 7: 'פסול', 17: true })]
}), 'handleSubmitResult', { check: ({ newRun }) => {
  const rows = newRun.tables[names.results];
  assert.equal(rows[rows.length - 1][14], 3, 'attempt count retains historical ordering before DQ cancellation');
  assert.equal(rows[4][7], 'בוטל');
  assert.equal(rows[4][17], false);
  assert.deepEqual(newRun.tables[names.pending].slice(1).map(r => r[5]), ['completed', 'cancelled', 'completed']);
} });
compare('cancelled examinee can recover genuine result', fixture({ pending: [pendingRow('cancelled')] }));
compare('disqualified examinee remains rejected', fixture({ pending: [pendingRow('disqualified')] }));
compare('approved registration and repeat registration retain append behavior', fixture({ pending: [pendingRow('approved')], exams: [] }), 'handleRegisterExamQuestions', { retry: true });
compare('registration language and shuffled key preserved', fixture({ input: input({ language: 'ru', questions: [{ qIdx: 0, qId: 1, shuffleOrder: [3, 2, 1, 0] }] }) }), 'handleRegisterExamQuestions');
compare('registration rechecks examiner cancellation after storing map', fixture({
  pending: [pendingRow('approved')],
  onAppend({ name, tables }) { if (name === names.exams) tables[names.pending][1][5] = 'cancelled'; }
}), 'handleRegisterExamQuestions', { check: ({ result }) => assert.equal(result.examStarted, false) });

compare('slow bank work sees concurrent fabricated fail and status change', fixture({
  input: input({ answers: [{ selected: 1, langAtAnswer: 'he' }, { selected: 1, langAtAnswer: 'he' }] }),
  onBankLoad({ tables }) {
    tables[names.pending][1][5] = 'completed';
    if (tables[names.results].length === 1) tables[names.results].push(resultRow({ 15: 'סגירת דפדפן' }));
  }
}));
compare('late successful retry is detected after fabricated-fail snapshot', fixture({
  afterRead({ name, kind, count, tables }) {
    if (name === names.results && kind === 'full' && count === 1) {
      tables[names.results].push(resultRow({ 5: '2/2', 6: '100%', 7: 'עבר', 8: '10:00' }));
      tables[names.pending][1][5] = 'completed';
    }
  }
}));
compare('new pending row during scoring is included in completion', fixture({
  pending: [pendingRow('completed')],
  input: input({ answers: [{ selected: 1, langAtAnswer: 'he' }, { selected: 1, langAtAnswer: 'he' }] }),
  onBankLoad({ tables }) { if (tables[names.pending].length === 2) tables[names.pending].push(pendingRow('in_exam')); }
}));
compare('same-size row movement during scoring recovers row indexes', fixture({
  pending: [pendingRow('in_exam', { 0: 'OTHER', 1: '900000002' }), pendingRow()],
  input: input({ answers: [{ selected: 1, langAtAnswer: 'he' }, { selected: 1, langAtAnswer: 'he' }] }),
  onBankLoad({ tables }) { if (tables[names.pending][2][1] === ID) [tables[names.pending][1], tables[names.pending][2]] = [tables[names.pending][2], tables[names.pending][1]]; }
}));
compare('new pending duplicate just after result append is completed', fixture({
  onAppend({ name, tables }) { if (name === names.results) tables[names.pending].push(pendingRow('approved')); }
}));
compare('concurrent overturn without append updates historical attempt count', fixture({
  results: [resultRow({ 13: 'OLD-ATTEMPT', 7: 'פסול', 17: true })],
  afterRead({ name, kind, count, tables }) {
    if (name === names.results && kind === 'full' && count === 1) {
      tables[names.results][1][7] = 'בוטל';
      tables[names.results][1][17] = false;
    }
  }
}), 'handleSubmitResult', { check: ({ newRun }) => assert.equal(newRun.tables[names.results].at(-1)[14], 1) });
compare('registration appended after mandatory-answers guard remains visible to scoring', fixture({
  afterRead({ name, kind, count, tables }) {
    if (name === names.exams && kind === 'full' && count === 1) {
      tables[names.exams].push(examRow({ 2: JSON.stringify(questionMap.map(q => ({ ...q, correctShuffledIdx: 3 }))) }));
    }
  }
}));
compare('registration appended during bank work remains visible to timing check', fixture({
  input: input({ answers: [{ selected: 1, langAtAnswer: 'he' }, { selected: 1, langAtAnswer: 'he' }] }),
  onBankLoad({ tables }) {
    if (tables[names.exams].length === 2) tables[names.exams].push(examRow({ 3: '2026-09-06T08:29:00Z' }));
  }
}), 'handleSubmitResult', { check: ({ newRun }) => assert.equal(newRun.tables[names.results].at(-1)[23], 'חשוד') });

const oldPending = Array.from({ length: 3999 }, (_, i) => pendingRow('completed', { 0: 'HISTORY', 1: '8' + String(i).padStart(8, '0'), 4: '2026-08-01T08:00:00Z' }));
const oldExams = Array.from({ length: 3999 }, (_, i) => examRow({ 0: 'HISTORY', 1: '8' + String(i).padStart(8, '0') }));
const oldResults = Array.from({ length: 5000 }, (_, i) => resultRow({ 0: '01/08/2026 08:00', 13: 'HISTORY', 1: '8' + String(i).padStart(8, '0') }));
const large = compare('4000 pending/4000 exams/5000 results preserves full historical semantics', fixture({ pending: [...oldPending, pendingRow()], exams: [...oldExams, examRow()], results: oldResults }), 'handleSubmitResult', {
  check: ({ oldRun, newRun }) => {
    assert.equal(oldRun.reads[names.exams].full, 3);
    assert.equal(newRun.reads[names.exams].full, 1);
    assert.equal(newRun.reads[names.pending].full, 1);
    assert.equal(newRun.reads[names.results].full, 2);
    assert.equal(newRun.bankLoads.length, 0);
    assert(Object.values(newRun.reads).reduce((n, r) => n + r.cells, 0) < Object.values(oldRun.reads).reduce((n, r) => n + r.cells, 0) * 0.6);
  }
});
// Historical attempts at the top of a large sheet still affect attempt numbers.
compare('attempt count includes oldest history outside any 1000-row tail', fixture({ results: [resultRow({ 13: 'OLD-ATTEMPT', 7: 'עבר' }), ...oldResults] }), 'handleSubmitResult', { check: ({ newRun }) => assert.equal(newRun.tables[names.results].at(-1)[14], 2) });
compare('many historical retakes use one fresh history read', fixture({ results: Array.from({ length: 8 }, (_, i) => resultRow({ 13: 'OLD-ATTEMPT-' + i, 7: i === 0 ? 'בוטל' : 'נכשל' })) }), 'handleSubmitResult', { check: ({ newRun }) => {
  assert.equal(newRun.tables[names.results].at(-1)[14], 8);
  assert.equal(newRun.reads[names.results].range, 0);
} });
compare('many pending duplicates remain bounded and all active rows complete', fixture({ pending: Array.from({ length: 8 }, (_, i) => pendingRow(i === 0 ? 'cancelled' : 'in_exam')) }), 'handleSubmitResult', { check: ({ newRun }) => {
  assert.deepEqual(newRun.tables[names.pending].slice(1).map(r => r[5]), ['cancelled', ...Array(7).fill('completed')]);
  assert.equal(newRun.reads[names.pending].range, 0);
} });

// Intentional optimization: a verified perfect result has no wrong answers even
// if the text bank is unavailable. Its score comes from the registered key.
const perfectUnavailable = runtime(currentSource, fixture({ failLanguages: ['he'] }));
assert.equal(perfectUnavailable.execute('handleSubmitResult', input()).status, 'ok');
assert.equal(perfectUnavailable.bankLoads.length, 0);
assert.equal(perfectUnavailable.tables[names.results][1][15], '');
passed++;

// Existing callers without snapshots retain their full-history behavior.
const helpers = runtime(currentSource, fixture({ pending: [pendingRow('approved'), pendingRow('in_exam'), pendingRow('disqualified')], results: [resultRow(), resultRow({ 7: 'בוטל' }), resultRow({ 4: 'C1' })] }));
assert.equal(helpers.run.countAttempts(ID, 'B'), 1);
helpers.run.markPendingCompleted('SYNTHETIC', ID);
assert.deepEqual(helpers.tables[names.pending].slice(1).map(r => r[5]), ['completed', 'completed', 'disqualified']);
passed++;

console.log(JSON.stringify({ passed, comparisons: summaries, largeReplay: {
  oldReads: large.oldRun.reads, newReads: large.newRun.reads,
  oldCells: Object.values(large.oldRun.reads).reduce((n, r) => n + r.cells, 0),
  newCells: Object.values(large.newRun.reads).reduce((n, r) => n + r.cells, 0)
} }, null, 2));
