// Run: node tests/server_cost.test.cjs
//
// COST is the regression this suite guards. Every exam-day incident of
// September 2026 was a read, not a wrong answer: the examiner dashboard's stale
// -row reconciliation reached 247 Sheets round trips and 6.86 M cells in ONE
// poll (review C R1), submitResult pulled ~15 MB (R2), teacherClassDetails read
// 109k practice rows including two JSON blob columns, and the reports fell back
// to reading every result ever recorded. The fixture is the size the audit
// measured against, and the assertions are budgets: a change that re-introduces
// a full read of a forever-growing sheet fails here.
//
// "reads" = calls that cross the Sheets service to fetch data. A tail read is
// two of them (the header row and the tail block).
'use strict';
const assert = require('node:assert/strict');
const { createEnv } = require('./helpers/server_env.cjs');

let checks = 0;
const check = (label, fn) => { fn(); checks++; console.log('ok  ' + label); };

const DAY = 86400000;
const NOW = Date.parse('2026-09-22T06:30:00Z');
const SESSION = 'SESS0001';

// 'תוצאות' col A is written by todayStr(): DD/MM/YYYY HH:MM.
function sheetDate(ms) {
  const d = new Date(ms);
  const p = n => String(n).padStart(2, '0');
  return `${p(d.getDate())}/${p(d.getMonth() + 1)}/${d.getFullYear()} ${p(d.getHours())}:${p(d.getMinutes())}`;
}
function iso(ms) { return new Date(ms).toISOString(); }

const PEND_HEADER = ['קוד סשן', 'ת.ז.', 'שם', 'טלפון', 'זמן הרשמה', 'סטטוס', 'שפה', 'אוכלוסיה', 'דרגה', 'שמע',
  'הארכת זמן', 'התחלת מבחן', 'טוקן נבחן', 'ספירת DQ', 'מסך נוסף', 'ספירת אזהרות', 'אזהרה אחרונה', 'אתר', 'סיים במכשיר'];
const RES_HEADER = ['תאריך', 'ת.ז.', 'שם', 'טלפון', 'דרגה', 'ציון', 'אחוז', 'עבר/נכשל', 'זמן', 'בוחן', 'אתר', 'כיתה',
  'שפה', 'קוד סשן', 'ניסיון', 'פירוט שגויות', 'נשלח?', 'פסול?', 'קישור וואטסאפ', 'אוכלוסיה', 'תוקן?', 'שמע', 'מאומת',
  'חשוד', 'dqEventId', 'תוקן ע"י', 'סיבת תיקון', 'תאריך תיקון', 'מסלול שפות', 'מכשיר'];
const EXAMS_HEADER = ['קוד סשן', 'ת.ז.', 'שאלות JSON', 'זמן רישום', 'שפה', 'שגויות לא מאומתות'];
const SESS_HEADER = ['קוד', 'בוחן ת.ז.', 'שם בוחן', 'אתר', 'כיתה', 'דרגה', 'שפה', 'מצב שמע', 'זמן יצירה', 'תקף עד',
  'פעיל', 'כמויות JSON', 'מאושרים JSON', 'בוחן אחראי', 'אוכלוסיה'];

function pendRow(opts) {
  const row = new Array(19).fill('');
  row[0] = opts.session; row[1] = opts.id; row[2] = opts.name || ('נבחן ' + opts.id); row[3] = opts.phone || '0500000000';
  row[4] = iso(opts.registered); row[5] = opts.status; row[6] = 'he'; row[7] = 'חיילים'; row[8] = opts.license || 'B';
  row[9] = opts.audio || 'off'; row[10] = opts.ext || ''; row[11] = opts.started ? iso(opts.started) : '';
  row[12] = opts.token || ('tok-' + opts.id); row[13] = 0; row[17] = 'בסיס 6';
  return row;
}
function resRow(opts) {
  const row = new Array(30).fill('');
  row[0] = sheetDate(opts.at); row[1] = opts.id; row[2] = 'נבחן ' + opts.id; row[3] = '0500000000';
  row[4] = opts.license || 'B'; row[5] = opts.score || '27/30'; row[6] = '90%'; row[7] = opts.passed || 'עבר';
  row[8] = '30 דק\' 00 שנ\''; row[9] = 'בוחן א'; row[10] = opts.site || 'בסיס 6'; row[11] = 'כיתה 1'; row[12] = 'he';
  row[13] = opts.session || SESSION; row[14] = opts.attempt || 1; row[15] = opts.wrong || ''; row[19] = 'חיילים';
  row[21] = 'off'; row[22] = 'מאומת';
  return row;
}

// ---- Fixture: the sizes appendix C measured (ממתינים 1,441×19, תוצאות 4,501×30,
// מבחנים 4,001, סשנים 302, בוחנים 42), with the ages the nightly archive job
// leaves behind: ממתינים 14 days, תוצאות 28 days, מבחנים 2 days. Rows ascend in
// time, like an append-only sheet, so readTail's 48-hour guard takes its cheap
// path and readHistorySince can reason about what is below the live sheet.
function buildFixture(extra) {
  const opts = extra || {};
  const spread = (i, n, days) => NOW - Math.round((days - (i * days) / n) * DAY);
  const pending = [PEND_HEADER];
  for (let i = 0; i < 1400; i++) {
    pending.push(pendRow({ session: 'OLD' + (i % 90), id: String(300000000 + i), status: 'completed',
      registered: spread(i, 1400, 14) }));
  }
  const results = [RES_HEADER];
  for (let i = 0; i < 4500; i++) {
    results.push(resRow({ at: spread(i, 4500, 28), id: String(300000000 + i), session: 'OLD' + (i % 90) }));
  }
  const exams = [EXAMS_HEADER];
  for (let i = 0; i < 4000; i++) {
    exams.push(['OLD' + (i % 90), String(300000000 + i), '{"q":1}', iso(spread(i, 4000, 2)), 'he', 0]);
  }
  const sessions = [SESS_HEADER];
  for (let i = 0; i < 300; i++) {
    sessions.push(['OLD' + i, '111111111', 'בוחן א', 'בסיס 6', 'כיתה 1', 'B', 'he', 'off',
      iso(spread(i, 300, 90)), iso(spread(i, 300, 90) + 8 * 3600000), false, '[]', '', '', '']);
  }
  sessions.push([SESSION, '111111111', 'בוחן א', 'בסיס 6', 'כיתה 1', 'B', 'he', 'off',
    iso(NOW - 3600000), iso(NOW + 7 * 3600000), true, '[]', '', 'בוחן א', '']);
  const examiners = [['שם', 'ת.ז.', 'סיסמה', 'פעיל', 'מס בוחן', 'תפקיד', 'טוקן', 'תוקף טוקן', 'ניסיונות כושלים', 'נעילה עד', 'אתרים מנוהלים']];
  for (let i = 0; i < 41; i++) {
    examiners.push(['בוחן ' + i, String(111111111 + i), 'pw', 'כן', String(i), 'בוחן', 'tokX', iso(NOW + DAY), 0, '', '']);
  }
  // Live examinees of the session under test.
  for (let i = 0; i < (opts.waiting || 3); i++) {
    pending.push(pendRow({ session: SESSION, id: String(400000000 + i), status: 'waiting', registered: NOW - 600000 }));
  }
  for (let i = 0; i < (opts.stale || 0); i++) {
    pending.push(pendRow({ session: SESSION, id: String(500000000 + i), status: 'in_exam',
      registered: NOW - 3 * 3600000, started: NOW - 3 * 3600000 }));
  }
  for (let i = 0; i < (opts.inExam || 2); i++) {
    pending.push(pendRow({ session: SESSION, id: String(600000000 + i), status: 'in_exam',
      registered: NOW - 600000, started: NOW - 300000 }));
  }
  for (let i = 0; i < (opts.done || 2); i++) {
    const id = String(700000000 + i);
    pending.push(pendRow({ session: SESSION, id, status: 'completed', registered: NOW - 3600000, started: NOW - 3000000 }));
    results.push(resRow({ at: NOW - 600000, id, session: SESSION }));
  }
  return { 'ממתינים': pending, 'תוצאות': results, 'מבחנים': exams, 'סשנים': sessions, 'בוחנים': examiners,
    'הארכות זמן': [['תאריך', 'קוד סשן', 'ת.ז.', 'שם', 'דקות', 'סיבה', 'בוחן']] };
}

const GATEWAY_URL = 'https://gw.example.workers.dev';
// startExam/startPractice refuse to write without a Worker to fetch the texts
// from (DESIGN §11.2), so the fixture is a CONFIGURED deployment.
const GATEWAY_PROPS = { GATEWAY_KEY: 'gateway-secret', GATEWAY_URL };
function envWith(extra, properties) {
  // The real answer key: startExam refuses to draw a question it cannot score,
  // so without it the draw finds nothing (and the snapshot check below would
  // pass vacuously).
  return createEnv({ sheets: buildFixture(extra), now: NOW,
    properties: Object.assign({}, GATEWAY_PROPS, properties || {}),
    sources: ['deployment/answer_key.gs'] });
}

// ---- 1. examinerDashboard, steady state ------------------------------------
{
  const env = envWith({});
  env.json(env.ctx.handleExaminerDashboard({ sessionCode: SESSION }));   // first poll warms the 30 s extensions cache
  env.resetCounters();
  const out = env.json(env.ctx.handleExaminerDashboard({ sessionCode: SESSION }));
  const c = env.counters();
  check('examinerDashboard steady: THREE sheet reads — ממתינים tail, תוצאות tail, nothing else', () => {
    assert.equal(out.status, 'ok');
    assert.equal(c.perSheet['ממתינים'].rangeReads, 2, 'ממתינים: one tail read = header + tail');
    assert.equal(c.perSheet['תוצאות'].rangeReads, 2, 'תוצאות: ONE tail read (the second one is gone)');
    assert.equal(c.perSheet['ממתינים'].fullReads + c.perSheet['תוצאות'].fullReads, 0, 'no full read of a growing sheet');
    assert.equal(c.perSheet['סשנים'].fullReads + c.perSheet['סשנים'].rangeReads, 0, 'the session row is not read when nothing is reconciled');
    assert.equal(c.perSheet['הארכות זמן'].fullReads, 0, 'the extensions map is cached for 30 s');
    assert.ok(c.reads <= 4, 'reads=' + c.reads);
  });
  check('examinerDashboard steady: under 100k cells', () => assert.ok(c.cellsRead <= 100000, 'cells=' + c.cellsRead));
  check('examinerDashboard steady: the board still lists everyone', () => {
    assert.equal(out.pending.length, 3);
    assert.equal(out.active.length, 2);
    assert.equal(out.completed.length, 2);
    assert.equal(out.completed[0].registrationTime !== undefined, true);
  });
}

// ---- 2. examinerDashboard after an outage: 40 stale rows at once ------------
{
  const env = envWith({ stale: 40 });
  env.resetCounters();
  const out = env.json(env.ctx.handleExaminerDashboard({ sessionCode: SESSION }));
  const c = env.counters();
  check('40 stale rows: at most 3 fabricated fail rows in one poll (was 40)', () => {
    assert.equal(out.status, 'ok');
    assert.equal(c.perSheet['תוצאות'].appends, 3);
  });
  check('40 stale rows: at most 3 status flips in one poll, the rest next poll', () => {
    assert.equal(c.perSheet['ממתינים'].setValues, 3);
  });
  check('40 stale rows: 12 Sheets reads or fewer (measured 247 round trips before)', () => {
    assert.ok(c.reads <= 12, 'reads=' + c.reads);
    assert.ok(c.cellsRead <= 150000, 'cells=' + c.cellsRead);
  });
  check('40 stale rows: סשנים read at most once, attempt history at most once', () => {
    assert.ok((c.perSheet['סשנים'].fullReads + c.perSheet['סשנים'].rangeReads) <= 1);
  });
  check('the fabricated rows are timeout fails with a real attempt number', () => {
    const rows = env.rows('תוצאות').slice(-3);
    for (const r of rows) {
      assert.equal(r[7], 'נכשל');
      assert.match(String(r[15]), /ניתוק\/טיימאאוט/);
      assert.equal(r[14], 1);
      assert.equal(r[9], 'בוחן א');   // examiner name came from the session row
    }
  });
  // The next poll continues where this one stopped.
  env.resetCounters();
  env.json(env.ctx.handleExaminerDashboard({ sessionCode: SESSION }));
  check('the next poll reconciles the next three', () => assert.equal(env.counters().perSheet['תוצאות'].appends, 3));
}

// ---- 3. the examinee pollers ------------------------------------------------
{
  const env = envWith({});
  const id = '600000000';
  env.json(env.ctx.handleCheckApproval({ sessionCode: SESSION, idNumber: id, examineeToken: 'tok-' + id }));
  env.resetCounters();
  const warm = env.json(env.ctx.handleCheckApproval({ sessionCode: SESSION, idNumber: id, examineeToken: 'tok-' + id }));
  check('checkApproval warm: zero Sheets reads (r23 snapshot)', () => {
    assert.equal(warm.status, 'ok');
    assert.equal(env.counters().reads, 0);
  });

  env.json(env.ctx.handleGetExamStatus({ sessionCode: SESSION, idNumber: id, examineeToken: 'tok-' + id }));
  env.resetCounters();
  const st = env.json(env.ctx.handleGetExamStatus({ sessionCode: SESSION, idNumber: id, examineeToken: 'tok-' + id }));
  check('getExamStatus warm: at most one read (the 30 s extensions read)', () => {
    assert.equal(st.examStatus, 'in_exam');
    assert.ok(env.counters().reads <= 1, 'reads=' + env.counters().reads);
  });
}

// ---- 3b. startExam pays for the grant in ScriptProperties, not in Sheets ----
// The bank grant (DESIGN §11.2) is signed inside the exam-start request. It
// must not add a single Sheets round trip to the hottest write path of the
// morning — the whole point of the r30 rewrite was 0 Drive and 1 pending read.
{
  const env = envWith({});
  const id = '400000000';   // a waiting row the examiner approves first
  env.ctx.handleApproveExaminee({ sessionCode: SESSION, idNumber: id, examinerId: '111111111', token: 'tokX' });
  let gatewayPropertyReads = 0;
  const realProps = env.ctx.PropertiesService;
  env.ctx.PropertiesService = { getScriptProperties: function() {
    const store = realProps.getScriptProperties();
    return Object.assign({}, store, {
      getProperty: k => { if (String(k).indexOf('GATEWAY') === 0) gatewayPropertyReads++; return store.getProperty(k); }
    });
  } };
  env.resetCounters();
  const started = env.json(env.ctx.doPost({ postData: { contents: JSON.stringify({
    action: 'startExam', origin: 'examinee-app', sessionCode: SESSION, idNumber: id,
    examineeToken: 'tok-' + id, language: 'he', license: 'B' }) } }));
  const c = env.counters();
  env.ctx.PropertiesService = realProps;
  check('startExam still costs one ממתינים tail read, one append and two cell writes', () => {
    assert.equal(started.status, 'ok');
    assert.equal(c.perSheet['ממתינים'].fullReads, 0, 'never a full read of a sheet that grows all day');
    assert.equal(c.perSheet['ממתינים'].setValues, 2, 'status + exam start');
    assert.equal(c.perSheet['מבחנים'].fullReads + c.perSheet['מבחנים'].rangeReads, 0, 'a fresh draw never reads מבחנים');
    assert.equal(c.perSheet['מבחנים'].appends, 1);
    // One tail read of ממתינים (header + tail) and the 30 s extensions map.
    assert.equal(c.perSheet['ממתינים'].rangeReads, 2);
    assert.equal(c.perSheet['הארכות זמן'].fullReads, 1);
    assert.equal(c.reads, 3, 'reads=' + c.reads + ' ' + JSON.stringify(c.perSheet));
  });
  check('the grant is paid for in ScriptProperties, never in Sheets', () => {
    assert.ok(started.bank && started.bank.grant, 'a grant was issued');
    // GATEWAY_URL + GATEWAY_KEY, once for the pre-write guard and once when the
    // grant is signed — four tiny property reads, zero Sheets round trips.
    assert.equal(gatewayPropertyReads, 4, 'gateway property reads=' + gatewayPropertyReads);
  });
}

// ---- 3c. r33: the Google fallback costs the auth read and one cell ----------
// A phone that cannot reach the Worker (KNOWN_ISSUES #38) reports it
// (reportGateway) and gets its texts through this script (bankRelay). On the
// audited fixture neither may cost more than the examinee-token check's ONE
// tail read of 'ממתינים': the report writes a single cell (Q), the relay
// writes nothing at all, and neither ever reads a growing sheet whole.
{
  const env = envWith({});
  const id = '600000000';   // in_exam in SESSION
  const post = body => env.json(env.ctx.doPost({ postData: { contents: JSON.stringify(Object.assign(
    { origin: 'examinee-app', sessionCode: SESSION, idNumber: id, examineeToken: 'tok-' + id }, body)) } }));
  env.resetCounters();
  const reported = post({ action: 'reportGateway', mode: 'google', diag: 'v1|why=probe|os=Android14' });
  const c = env.counters();
  check('reportGateway: one ממתינים tail read (the auth check), ONE cell written, nothing else in Sheets', () => {
    assert.deepEqual(reported, { status: 'ok' });
    assert.equal(c.perSheet['ממתינים'].fullReads, 0, 'never a full read of a sheet that grows all day');
    assert.equal(c.perSheet['ממתינים'].rangeReads, 2, 'one tail read = header + tail, handed from the auth check');
    assert.equal(c.perSheet['ממתינים'].setValues, 1, 'column Q only');
    assert.equal(c.perSheet['ממתינים'].appends, 0);
    const row = env.rows('ממתינים').filter(r => r[0] === SESSION && r[1] === id)[0];
    assert.equal(row[16], '📡 גיבוי גוגל (probe)');
    assert.equal(row[15], '', 'the warning counter cell is exactly as it was');
  });

  env.ctx.UrlFetchApp = { fetch: () => ({ getResponseCode: () => 200,
    getContentText: () => '{"status":"ok","build":"b","questions":[{"id":1}],"missing":[]}' }) };
  const grant = env.ctx.bankGrantFor('exam', [1, 2, 3], SESSION + ':' + id).grant;
  env.resetCounters();
  const relayed = post({ action: 'bankRelay', grant });
  const r = env.counters();
  check('bankRelay: the auth check\'s tail read and nothing else — no write, no second read', () => {
    assert.equal(relayed.status, 'ok');
    assert.equal(relayed.relay, true);
    assert.equal(r.reads, 2, 'reads=' + r.reads + ' ' + JSON.stringify(r.perSheet));
    assert.equal(r.appends + r.setValues, 0, 'the relay writes nothing');
  });
}

// ---- 4. sessionSnapshot: one upstream call for the whole session ------------
// r32 (DESIGN §14.1): the board is drawn from THIS answer, so it costs one more
// tail read — 'תוצאות', the same one the board itself pays. That is the whole
// price of the change, and it replaces a SECOND round trip to Google
// (examinerDashboard) per change, whose delivery hop is what stalls 25-60 s
// (KNOWN_ISSUES #35). The budget below is what must never grow again.
{
  const env = envWith({ waiting: 2, inExam: 2 });
  env.resetCounters();
  const cold = env.json(env.ctx.handleSessionSnapshot({ sessionCode: SESSION }));
  const coldC = env.counters();
  check('sessionSnapshot cold: two tail reads (ממתינים, תוצאות) + the extensions read, and no write', () => {
    assert.equal(cold.status, 'ok');
    assert.equal(cold.v, 2);
    assert.equal(coldC.perSheet['ממתינים'].rangeReads, 2, 'ממתינים: one tail read = header + tail');
    assert.equal(coldC.perSheet['תוצאות'].rangeReads, 2, 'תוצאות: one tail read, no cache');
    assert.equal(coldC.perSheet['ממתינים'].fullReads + coldC.perSheet['תוצאות'].fullReads, 0,
      'never a full read of a sheet that grows all day');
    assert.equal(coldC.perSheet['הארכות זמן'].fullReads, 1, 'the 30 s extensions map');
    assert.equal(coldC.reads, 5, 'reads=' + coldC.reads + ' ' + JSON.stringify(coldC.perSheet));
    assert.equal(coldC.appends + coldC.setValues, 0, 'the snapshot is read-only');
  });
  env.resetCounters();
  const snap = env.json(env.ctx.handleSessionSnapshot({ sessionCode: SESSION }));
  const warmC = env.counters();
  check('sessionSnapshot warm: ממתינים comes from the 4 s snapshot, only תוצאות is re-read', () => {
    assert.equal(warmC.perSheet['ממתינים'].fullReads + warmC.perSheet['ממתינים'].rangeReads, 0,
      'the r23 per-session snapshot still serves the rows');
    assert.equal(warmC.perSheet['הארכות זמן'].fullReads, 0, 'the extensions map is cached for 30 s');
    assert.equal(warmC.reads, 2, 'reads=' + warmC.reads + ' — the results tail, nothing else');
  });
  check('sessionSnapshot returns every row of the session, oldest first', () => {
    assert.equal(snap.status, 'ok');
    assert.equal(snap.rows.length, 6);   // 2 waiting + 2 in_exam + 2 completed
    assert.equal(snap.rows[0].id, '400000000');
    assert.equal(snap.rows[0].status, 'waiting');
    assert.equal(typeof snap.at, 'number');
  });
  check('sessionSnapshot never leaks the examinee token, and its row shape is pinned', () => {
    const text = JSON.stringify(snap);
    assert.equal(text.indexOf('tok-'), -1, 'the token itself is still never sent — only its SHA-256');
    for (const row of snap.rows) {
      assert.match(row.tokenHash, /^[0-9a-f]{64}$/);
      // r31 (DESIGN §13.6) added warn/fin/ext/dq — counters and flags the
      // examiner board shows, so the Worker's fingerprint can wake it on them.
      // r32 (§14.1) added the board's own columns: the answer now DOES carry
      // names and phones, on purpose, because the board is drawn from it and
      // the only route that forwards them (/v1/session/watch) answers an
      // examiner grant. The list stays exhaustive: it is the guard against a
      // field being added some day without anybody deciding to.
      assert.deepEqual(Object.keys(row).sort(),
        ['attemptsToday', 'audio', 'dq', 'examMinutes', 'ext', 'extraMinutes', 'fin', 'id', 'lang', 'lastWarn',
          'lic', 'name', 'phone', 'pop', 'site', 'start', 'status', 'time', 'timeExt', 'tokenHash', 'warn']
          .concat(row.todayExams ? ['todayExams'] : []).sort());
    }
  });
  check('sessionSnapshot hashes the token the way the Worker does (SHA-256 hex)', () => {
    const expected = require('node:crypto').createHash('sha256').update('tok-400000000', 'utf8').digest('hex');
    assert.equal(snap.rows[0].tokenHash, expected);
  });
  check('sessionSnapshot reports the authorised exam length', () => {
    assert.equal(snap.rows[0].examMinutes, 40);
    assert.equal(snap.rows[0].extraMinutes, 0);
    assert.equal(snap.rows[0].audio, 'off');
  });
  check('sessionSnapshot carries the session results, without the wrong-answers blob', () => {
    assert.equal(snap.results.length, 2, 'the two examinees who finished');
    assert.deepEqual(snap.results.map(r => r.idNumber).sort(), ['700000000', '700000001']);
    assert.equal(snap.results[0].score, '27/30');
    assert.equal(JSON.stringify(snap).indexOf('wrongDetails'), -1);
  });
}

// ---- 5. getSessionInfo carries the build and the gateway switch -------------
{
  const env = envWith({});
  const info = env.json(env.ctx.handleGetSessionInfo({ sessionCode: SESSION }));
  check('getSessionInfo returns build + gateway url', () => {
    assert.equal(info.status, 'ok');
    assert.equal(typeof info.session.build, 'string');
    assert.equal(info.session.gateway.url, GATEWAY_URL);
  });
  const env2 = envWith({}, { GATEWAY_URL: '' });
  const info2 = env2.json(env2.ctx.handleGetSessionInfo({ sessionCode: SESSION }));
  check('an unset GATEWAY_URL means "poll me directly"', () => assert.equal(info2.session.gateway.url, ''));
  // ...and since r30 clearing it also takes the question texts down: there is
  // deliberately no partial switch (Yossi, 21/09) — the Worker is required.
  check('an unset GATEWAY_URL also means no bank grant', () => assert.equal(env2.ctx.bankGrantFor('exam', [1], 'x'), null));
}

// ---- 6. siteCombinedReport: today is cheap, an old day reads the archive ----
{
  const env = envWith({});
  env.resetCounters();
  const rep = env.json(env.ctx.handleSiteCombinedReport({ sessionCode: SESSION, examinerId: '111111111', token: 'tokX' }));
  const c = env.counters();
  check('siteCombinedReport for today: three reads or fewer, no full results read', () => {
    assert.equal(rep.status, 'ok');
    assert.equal(c.perSheet['תוצאות'].fullReads, 0);
    assert.ok(c.perSheet['תוצאות'].rangeReads <= 3, 'result reads=' + c.perSheet['תוצאות'].rangeReads);
  });

  // 60 days ago: beyond the 30-day retention window, so the rows live in the archive.
  const oldDay = NOW - 60 * DAY;
  const fixture = buildFixture({});
  fixture['סשנים'].push(['OLDSESS1', '111111111', 'בוחן א', 'אתר ישן', 'כיתה 9', 'B', 'he', 'off',
    iso(oldDay), iso(oldDay + 8 * 3600000), false, '[]', '', 'בוחן א', '']);
  fixture['תוצאות_ארכיון'] = [RES_HEADER,
    resRow({ at: oldDay, id: '900000001', session: 'OLDSESS1', site: 'אתר ישן' }),
    resRow({ at: oldDay, id: '900000002', session: 'OLDSESS1', site: 'אתר ישן', passed: 'נכשל' })];
  const envOld = createEnv({ sheets: fixture, now: NOW });
  const old = envOld.json(envOld.ctx.handleSiteCombinedReport({ sessionCode: 'OLDSESS1', examinerId: '111111111', token: 'tokX' }));
  check('siteCombinedReport for a 60-day-old day reads the archive', () => {
    assert.equal(old.status, 'ok');
    assert.equal(old.results.length, 2);
    assert.deepEqual(old.results.map(r => r.idNumber).sort(), ['900000001', '900000002']);
  });
}

// ---- 7. teacherClassDetails: rows since the class, never column N ----------
{
  const classCreated = NOW - 30 * DAY;
  const practice = [['תאריך', 'מזהה תלמיד', 'שם תלמיד', 'קוד כיתה', 'מצב', 'דרגה', 'ציון', 'סה"כ', 'אחוז', 'עבר/נכשל',
    'זמן', 'נושא', 'שפה', 'פירוט שגויות', 'פירוט לפי נושא', 'טלפון']];
  const blob = JSON.stringify(Array.from({ length: 20 }, (_, i) => ({ qNum: i, qText: 'x'.repeat(80) })));
  for (let i = 0; i < 20000; i++) {
    const at = NOW - Math.round((200 - (i * 200) / 20000) * DAY);
    practice.push([sheetDate(at), i % 3 === 0 ? 'stu-1' : 'stu-' + i, 'תלמיד', i % 3 === 0 ? 'CLS1' : 'CLS9',
      'exam', 'B', 26, 30, 87, 'עבר', '12:30', '', 'he', blob,
      JSON.stringify({ 'חוק': { correct: 3, total: 4 } }), '0500000000']);
  }
  const env = createEnv({ now: NOW, sheets: {
    'כיתות': [['קוד כיתה', 'שם כיתה', 'מורה ת.ז.', 'שם מורה', 'דרגה', 'תאריך יצירה', 'פעיל', 'אתר'],
      ['CLS1', 'כיתה א', '222222222', 'מורה', 'B', iso(classCreated), 'כן', 'בסיס 6']],
    'תלמידי כיתות': [['קוד כיתה', 'שם תלמיד', 'מזהה תלמיד', 'תאריך הצטרפות'], ['CLS1', 'תלמיד', 'stu-1', iso(classCreated)]],
    'תוצאות תרגול': practice
  } });
  env.resetCounters();
  const details = env.json(env.ctx.handleTeacherClassDetails({ classCode: 'CLS1', teacherId: '222222222' }));
  const sheetCounters = env.counters().perSheet['תוצאות תרגול'];
  check('teacherClassDetails answers from the rows since the class was created', () => {
    assert.equal(details.status, 'ok');
    assert.equal(details.students.length, 1);
    assert.ok(details.students[0].totalPractices > 0);
    assert.equal(sheetCounters.fullReads, 0, 'never a full read of the practice sheet');
    // A full read is 20,000 × 16 = 320,000 cells. What is left: one pass over
    // the date column to find the boundary exactly (r19), then the 3,000 rows
    // since the class was created, in 14 of the 16 columns.
    assert.ok(sheetCounters.cellsRead < 80000, 'cells=' + sheetCounters.cellsRead);
  });
  check('teacherClassDetails never reads the 2 KB wrong-answers blob (column N)', () => {
    const text = JSON.stringify(details);
    assert.equal(text.indexOf('x'.repeat(80)), -1);
    assert.equal(text.indexOf('wrongDetails'), -1);
  });
  check('the per-topic breakdown (column O) is still there — the client draws it', () => {
    assert.deepEqual(details.students[0].categoryErrors['חוק'] !== undefined, true);
  });
}

// ---- 7b. submitPracticeResult: the class check costs two column reads -------
// The action is auth:'none', so the class code it is handed is verified before
// it can reach the teacher/commander statistics (TODO 1.2). That is two reads
// the handler did not pay before, and both are bounded by the number of classes
// and of enrolled students — never by the practice history, which is still only
// appended to. An unknown class pays ONE: the roster is not read at all.
{
  const classes = [['קוד כיתה', 'שם כיתה', 'מורה ת.ז.', 'שם מורה', 'דרגה', 'תאריך יצירה', 'פעיל', 'אתר']];
  for (let i = 0; i < 200; i++) classes.push(['CLS' + i, 'כיתה ' + i, '222222222', 'מורה', 'B', iso(NOW - 60 * DAY), 'כן', 'בסיס 6']);
  const roster = [['קוד כיתה', 'שם תלמיד', 'מזהה תלמיד', 'תאריך הצטרפות']];
  for (let i = 0; i < 5000; i++) roster.push(['CLS' + (i % 200), 'תלמיד ' + i, 'stu-' + i, iso(NOW - 40 * DAY)]);
  const history = [['תאריך', 'מזהה תלמיד', 'שם תלמיד', 'קוד כיתה', 'מצב', 'דרגה', 'ציון', 'סה"כ', 'אחוז', 'עבר/נכשל',
    'זמן', 'נושא', 'שפה', 'פירוט שגויות', 'פירוט לפי נושא', 'טלפון']];
  for (let i = 0; i < 20000; i++) {
    history.push([sheetDate(NOW - DAY), 'stu-' + (i % 5000), 'תלמיד', 'CLS' + (i % 200), 'exam', 'B', 26, 30, 87,
      'עבר', '12:30', '', 'he', '', '', '0500000000']);
  }
  const env = createEnv({ now: NOW, sheets: { 'כיתות': classes, 'תלמידי כיתות': roster, 'תוצאות תרגול': history } });
  const submit = params => env.json(env.ctx.handleSubmitPracticeResult(Object.assign(
    { studentName: 'תלמיד 7', mode: 'exam', license: 'B', score: 26, total: 30, percent: 87, time: '12:30' }, params)));
  env.resetCounters();
  const stored = submit({ studentId: 'stu-7', classCode: 'CLS7' });
  const c = env.counters();
  check('submitPracticeResult verifies the class in two column reads and one append', () => {
    assert.deepEqual(stored, { status: 'ok' });
    assert.equal(c.reads, 2, 'reads=' + c.reads);
    assert.equal(c.perSheet['כיתות'].fullReads + c.perSheet['תלמידי כיתות'].fullReads, 0, 'columns only, never getDataRange');
    assert.equal(c.perSheet['תוצאות תרגול'].fullReads + c.perSheet['תוצאות תרגול'].rangeReads, 0,
      'the 20,000-row practice history is only appended to');
    assert.equal(c.appends, 1);
    // 201 class codes + 5,001 roster rows × 3 columns (A the code, C the id, B
    // along for the ride). Two getDataRange() reads would have been 201 × 8 +
    // 5,001 × 4 and would grow with every column ever added to those sheets.
    assert.ok(c.cellsRead <= 201 + 3 * 5001, 'cells=' + c.cellsRead);
    assert.equal(env.rows('תוצאות תרגול').at(-1)[3], 'CLS7', 'an enrolled student keeps the class');
  });
  env.resetCounters();
  const unknown = submit({ studentId: 'stu-7', classCode: 'NOSUCH' });
  check('a class that does not exist costs ONE read, and the row is stored without it', () => {
    assert.deepEqual(unknown, { status: 'ok', classUnknown: true });
    assert.equal(env.counters().reads, 1, 'reads=' + env.counters().reads);
    assert.equal(env.rows('תוצאות תרגול').at(-1)[3], '');
  });
}

// ---- 8. countAttempts: live + archive ---------------------------------------
{
  const fixture = buildFixture({});
  const id = '800000001';
  fixture['תוצאות'].push(resRow({ at: NOW - 2 * DAY, id, passed: 'נכשל' }));
  fixture['תוצאות'].push(resRow({ at: NOW - DAY, id, passed: 'בוטל' }));           // overturned: not an attempt
  fixture['תוצאות_ארכיון'] = [RES_HEADER,
    resRow({ at: NOW - 100 * DAY, id, passed: 'נכשל' }),
    resRow({ at: NOW - 90 * DAY, id, passed: 'פסול' }),
    resRow({ at: NOW - 80 * DAY, id: '899999999', passed: 'נכשל' })];              // someone else
  const env = createEnv({ sheets: fixture, now: NOW });
  check('countAttempts counts live + archive and skips בוטל', () => {
    assert.equal(env.ctx.countAttempts(id, 'B'), 3);
    assert.equal(env.ctx.countAttempts(id, 'C'), 0);
  });
  const envNoArchive = createEnv({ sheets: buildFixture({}), now: NOW });
  check('countAttempts works when no archive sheet exists yet', () => {
    assert.equal(envNoArchive.ctx.countAttempts('300000005', 'B'), 1);
  });
  env.resetCounters();
  env.ctx.countAttempts(id, 'B');
  check('countAttempts reads three narrow columns, not the whole sheet', () => {
    const c = env.counters();
    assert.equal(c.perSheet['תוצאות'].fullReads, 0);
    assert.equal(c.perSheet['תוצאות'].rangeReads, 3);
    assert.ok(c.cellsRead < 30000, 'cells=' + c.cellsRead);
  });
}

// ---- 9. readResultsSince: one header, archive first -------------------------
{
  const fixture = buildFixture({});
  fixture['תוצאות_ארכיון'] = [RES_HEADER,
    resRow({ at: NOW - 200 * DAY, id: '910000001' }),
    resRow({ at: NOW - 150 * DAY, id: '910000002' })];
  const env = createEnv({ sheets: fixture, now: NOW });
  const merged = env.ctx.readResultsSince(new Date(NOW - 300 * DAY));
  check('readResultsSince merges live + archive under ONE header', () => {
    assert.equal(merged.rows[0][0], 'תאריך');
    assert.equal(merged.rows.filter(r => r[0] === 'תאריך').length, 1);
    assert.equal(merged.rows.length, 1 + 2 + (env.rows('תוצאות').length - 1));
    assert.equal(String(merged.rows[1][1]), '910000001', 'archive rows come first (oldest first)');
  });
  const recent = env.ctx.readResultsSince(new Date(NOW - 2 * DAY));
  check('a recent cutoff never touches the archive', () => {
    assert.equal(recent.rows.some(r => String(r[1]) === '910000001'), false);
    assert.match(recent.mode, /arch:0$/);
  });
}

// ---- 10. every status writer drops the pollers' snapshot --------------------
{
  const removedKeys = [];
  function spyEnv(extra) {
    const env = envWith(extra || {});
    const realRemove = env.cache.remove.bind(env.cache);
    env.cache.remove = key => { removedKeys.push(key); return realRemove(key); };
    return env;
  }
  const snapKey = 'qv2_pendsnap_' + SESSION;
  function writesSnapshot(label, run) {
    removedKeys.length = 0;
    const env = spyEnv();
    env.ctx.pendingRowsForSession(SESSION);             // a poller cached the state
    assert.ok(env.cache.get(snapKey), 'precondition: snapshot cached');
    run(env);
    check(label + ' invalidates the pending snapshot', () => {
      assert.ok(removedKeys.indexOf(snapKey) !== -1, 'snapshot key never removed by ' + label);
      assert.equal(env.cache.get(snapKey), null);
    });
  }
  const waitingId = '400000000', inExamId = '600000000';
  const examiner = { examinerId: '111111111', token: 'tokX' };
  writesSnapshot('approve', env => env.ctx.handleApproveExaminee(Object.assign({ sessionCode: SESSION, idNumber: waitingId }, examiner)));
  writesSnapshot('reject', env => env.ctx.handleRejectExaminee(Object.assign({ sessionCode: SESSION, idNumber: waitingId }, examiner)));
  writesSnapshot('cancelRegistration', env => env.ctx.handleCancelRegistration({ sessionCode: SESSION, idNumber: waitingId, phone: '0500000000', examineeToken: 'tok-' + waitingId }));
  writesSnapshot('startExam (approved → in_exam)', env => {
    env.ctx.handleApproveExaminee(Object.assign({ sessionCode: SESSION, idNumber: waitingId }, examiner));
    removedKeys.length = 0;
    env.ctx.pendingRowsForSession(SESSION);
    env.ctx.handleStartExam({ sessionCode: SESSION, idNumber: waitingId, examineeToken: 'tok-' + waitingId });
  });
  writesSnapshot('disqualify', env => env.ctx.handleDisqualify(Object.assign({ sessionCode: SESSION, idNumber: inExamId, dqEventId: 'e1' }, examiner)));
  writesSnapshot('confirmDQ', env => {
    env.ctx.handleDisqualify(Object.assign({ sessionCode: SESSION, idNumber: inExamId, dqEventId: 'e2' }, examiner));
    removedKeys.length = 0;
    env.ctx.pendingRowsForSession(SESSION);
    env.ctx.handleConfirmDQ(Object.assign({ sessionCode: SESSION, idNumber: inExamId }, examiner));
  });
  writesSnapshot('overturnDQ', env => {
    env.ctx.handleDisqualify(Object.assign({ sessionCode: SESSION, idNumber: inExamId, dqEventId: 'e3' }, examiner));
    removedKeys.length = 0;
    env.ctx.pendingRowsForSession(SESSION);
    env.ctx.handleOverturnDQ(Object.assign({ sessionCode: SESSION, idNumber: inExamId }, examiner));
  });
  writesSnapshot('resetExaminee', env => env.ctx.handleResetExaminee(Object.assign({ sessionCode: SESSION, idNumber: inExamId }, examiner)));
  writesSnapshot('forceComplete', env => env.ctx.handleForceComplete(Object.assign({ sessionCode: SESSION, idNumber: inExamId }, examiner)));
  writesSnapshot('markFinished', env => env.ctx.handleMarkFinished({ sessionCode: SESSION, idNumber: inExamId, examineeToken: 'tok-' + inExamId }));
  writesSnapshot('reportWarning', env => env.ctx.handleReportWarning({ sessionCode: SESSION, idNumber: inExamId, examineeToken: 'tok-' + inExamId, reason: 'tab-switch' }));
  writesSnapshot('reportGateway (r33)', env => env.ctx.handleReportGateway({ sessionCode: SESSION, idNumber: inExamId, examineeToken: 'tok-' + inExamId, mode: 'google', diag: 'v1|why=probe' }));
  writesSnapshot('self-disqualify with a reason (r33)', env => env.ctx.handleDisqualify({ sessionCode: SESSION, idNumber: inExamId, examineeToken: 'tok-' + inExamId, dqEventId: 'r33', reason: 'split-area' }));
  writesSnapshot('the dashboard reconciliation', env => {
    const sheet = env.sheet('ממתינים');
    // make one in_exam row stale: registered long before the exam window ends
    for (const row of sheet.rows) {
      if (row[0] === SESSION && row[5] === 'in_exam') { row[4] = iso(NOW - 5 * 3600000); row[11] = iso(NOW - 5 * 3600000); break; }
    }
    env.ctx.handleExaminerDashboard({ sessionCode: SESSION });
  });
}

// ---- 11. every converted reader still answers -------------------------------
// The history readers moved to readResultsSince / readRowsSince and the
// commander stopped loading question banks. These are the reports nobody runs
// in a test until an exam evening, so each one is called once here against the
// real fixture: a missing helper or a colSpec that forgot a column shows up as
// an error or an empty report instead of at 20:00 on a Sunday.
{
  const fixture = buildFixture({});
  const wrongBlock = [
    'מזהה שאלה: 1442', 'שאלה: מה המרחק המינימלי?', 'תשובת הנבחן: א - 10 מטר',
    'תשובה נכונה: ב - 20 מטר', 'קטגוריה: חוקי התנועה'
  ].join('\n');
  for (let i = 0; i < 6; i++) {
    fixture['תוצאות'].push(resRow({ at: NOW - 3600000, id: String(950000000 + i), passed: i % 2 ? 'נכשל' : 'עבר',
      wrong: wrongBlock, session: SESSION }));
  }
  // A commander, a center manager, a teacher-commander and an admin.
  fixture['בוחנים'].push(['מפקד', '999999999', 'pw', 'כן', '99', 'מפקד', 'tokC', iso(NOW + DAY), 0, '', '']);
  fixture['בוחנים'].push(['מפקד מרכז', '999999998', 'pw', 'כן', '98', 'מפקד מרכז', 'tokM', iso(NOW + DAY), 0, '', 'בסיס 6']);
  fixture['מורים'] = [['שם', 'ת.ז.', 'סיסמה', 'פעיל', 'טוקן', 'תוקף טוקן', 'ניסיונות כושלים', 'נעילה עד', 'תפקיד', 'אתר'],
    ['מורה מפקד', '222222222', 'pw', 'כן', 'tokT', iso(NOW + DAY), 0, '', 'מפקד ראשי', ''],
    ['אדמין', '222222223', 'pw', 'כן', 'tokA', iso(NOW + DAY), 0, '', 'אדמין', '']];
  fixture['כיתות'] = [['קוד כיתה', 'שם כיתה', 'מורה ת.ז.', 'שם מורה', 'דרגה', 'תאריך יצירה', 'פעיל', 'אתר'],
    ['CLS1', 'כיתה א', '222222222', 'מורה מפקד', 'B', iso(NOW - 40 * DAY), 'כן', 'בסיס 6']];
  fixture['תלמידי כיתות'] = [['קוד כיתה', 'שם תלמיד', 'מזהה תלמיד', 'תאריך הצטרפות'], ['CLS1', 'נבחן 950000000', 'stu-1', iso(NOW - 40 * DAY)]];
  fixture['תוצאות תרגול'] = [['תאריך', 'מזהה תלמיד', 'שם תלמיד', 'קוד כיתה', 'מצב', 'דרגה', 'ציון', 'סה"כ', 'אחוז',
    'עבר/נכשל', 'זמן', 'נושא', 'שפה', 'פירוט שגויות', 'פירוט לפי נושא', 'טלפון']];
  for (let i = 0; i < 40; i++) {
    fixture['תוצאות תרגול'].push([sheetDate(NOW - (20 - i / 2) * DAY), 'stu-1', 'נבחן 950000000', 'CLS1', 'exam', 'B',
      26, 30, 87, 'עבר', '12:30', '', 'he', JSON.stringify([{ qNum: 1, category: 'חוקי התנועה', qText: 'שאלה' }]),
      JSON.stringify({ 'חוק': { correct: 3, total: 4 } }), '0500000000']);
  }
  // A soldier who practices badly and has never taken the real exam: the at-risk
  // list needs one, and the forecast joins him to the live registrants by phone.
  for (let i = 0; i < 6; i++) {
    fixture['תוצאות תרגול'].push([sheetDate(NOW - (10 - i) * DAY), 'stu-2', 'דני לוי', 'CLS1', 'exam', 'B',
      12, 30, 40, 'נכשל', '09:00', '', 'he', '[]', '{}', '0521234567']);
  }
  fixture['ממתינים'].push(pendRow({ session: SESSION, id: '410000000', status: 'waiting',
    registered: NOW - 300000, name: 'דני לוי', phone: '0521234567' }));
  const env = createEnv({ sheets: fixture, now: NOW });
  const range = { dateFrom: '01/09/2026', dateTo: '22/09/2026' };

  // This fixture IS an exam morning (SESSION: 3 + 1 waiting, 2 in the exam), so
  // since 24/09/2026 the two heavy reports refuse it (examHoursRefusal,
  // 86_commander.js; the rule is tested in api_reliability, its cost in 11b).
  // This section is about the readers, so it asks with the emergency override.
  const refused = env.json(env.ctx.handleCommanderDashboard(Object.assign({ examinerId: '999999999', token: 'tokC' }, range)));
  check('on this exam-morning fixture the commander dashboard is refused (exam_hours) without the override', () => {
    assert.equal(refused.code, 'exam_hours');
    assert.deepEqual(refused.live, { sessions: 1, examinees: 6 });
  });
  const force = { force: '1' };
  const cmd = env.json(env.ctx.handleCommanderDashboard(Object.assign({ examinerId: '999999999', token: 'tokC' }, range, force)));
  check('commanderDashboard runs with no question bank at all', () => {
    assert.equal(cmd.status, 'ok');
    assert.ok(cmd.data.overall.total > 0);
    assert.ok(cmd.data.waitTimes, 'the wait-time block survived the date bound');
  });
  check('topWrong is {questionId, count, category, text} — the client resolves the rest', () => {
    assert.equal(cmd.data.topWrong.length, 1);
    const entry = cmd.data.topWrong[0];
    assert.deepEqual(Object.keys(entry).sort(), ['category', 'count', 'langCounts', 'questionId', 'text']);
    assert.equal(entry.questionId, '1442');
    assert.equal(entry.count, 6);
    assert.equal(entry.text, 'מה המרחק המינימלי?');
    assert.equal(entry.category, 'חוק', 'the row\'s "קטגוריה:" classified into a blueprint topic');
    assert.equal(entry.imageUrl, undefined, 'no image resolution on the server any more');
  });
  check('weak topics come from the row itself (קטגוריה:), not from a bank', () => {
    const overall = cmd.data.weakTopics.overall;
    assert.ok(overall.length > 0);
    assert.ok(overall.some(t => t.topic === 'חוק' && t.wrong === 6), JSON.stringify(overall));
  });

  check('centerManagerReport runs', () => {
    const rep = env.json(env.ctx.handleCenterManagerReport(Object.assign({ examinerId: '999999998', token: 'tokM' }, range, force)));
    assert.equal(rep.status, 'ok');
    assert.ok(rep.overall.total > 0);
    assert.ok(rep.results.length > 0);
  });
  check('teacherCommanderDashboard runs (practice + repeat exam failures)', () => {
    const t = env.json(env.ctx.handleTeacherCommanderDashboard(Object.assign({ teacherId: '222222222', token: 'tokT' }, range)));
    assert.equal(t.status, 'ok');
    assert.ok(t.data.overall.total > 0);
    assert.ok(Array.isArray(t.data.repeatFailures));
  });
  check('adminDashboard runs', () => {
    const a = env.json(env.ctx.handleAdminDashboard(Object.assign({ teacherId: '222222223', token: 'tokA' }, range)));
    assert.equal(a.status, 'ok');
    assert.ok(a.data.overall.total > 0);
  });
  check('teacherClassDetails + export run on the same class', () => {
    const d = env.json(env.ctx.handleTeacherClassDetails({ classCode: 'CLS1', teacherId: '222222222' }));
    assert.equal(d.status, 'ok');
    assert.equal(d.students.length, 1);
    const ex = env.json(env.ctx.handleTeacherExportData({ classCode: 'CLS1', teacherId: '222222222' }));
    assert.equal(ex.status, 'ok');
    assert.equal(ex.rows.length, 46, 'every practice row of the class, both students');
    assert.ok(String(ex.rows[0]['פירוט שגויות']).length > 0, 'the export keeps the JSON columns');
  });
  check('the predictive model and the at-risk list build from live + archive', () => {
    const model = env.ctx.buildPassProbabilityModel({ lookbackDays: 30 });
    assert.ok(model.base.n > 0);
    const atRisk = env.ctx.computeAtRiskAll({ lookbackDays: 30 });
    assert.ok(Array.isArray(atRisk.students));
    env.ctx.rebuildAtRiskCache();
    const forecast = env.json(env.ctx.handleExaminerForecast({ examinerId: '999999999', token: 'tokC' }));
    assert.equal(forecast.status, 'ok');
    assert.ok(forecast.data.examDayForecast, 'the live-registrant forecast read ממתינים');
  });
  check('verifyLogin caches its positive verdict (no sheet read the second time)', () => {
    const first = env.json(env.ctx.handleVerifyLogin({ examinerId: '999999999', token: 'tokC' }));
    assert.equal(first.status, 'ok');
    env.resetCounters();
    const second = env.json(env.ctx.handleVerifyLogin({ examinerId: '999999999', token: 'tokC' }));
    assert.equal(second.status, 'ok');
    assert.equal(second.examiner.role, 'מפקד');
    assert.equal(env.counters().perSheet['בוחנים'].fullReads, 0);
  });
  check('a client log lands in אבחון, truncated to 2 KB', () => {
    const res = env.ctx.diagRecordClientLog(SESSION, '400000000', Array.from({ length: 50 }, (_, i) => ({ t: i, e: 'poll-timeout' })));
    assert.equal(res, 'appended');
    const row = env.rows('אבחון')[env.rows('אבחון').length - 1];
    assert.equal(row[1], 'CLIENT');
    assert.equal(row[2], SESSION);
    assert.ok(String(row[6]).length <= 2048);
    assert.equal(env.ctx.diagRecordClientLog(SESSION, '400000000', null), 'empty');
  });
}

// ---- 11b. a refused report costs about one examiner poll, never the history ---
// 24/09/2026: three commander-dashboard opens (60 s, 55 s, 42 s) on an exam
// morning, then a 243 s hang of the exam project's own spreadsheet open. The two
// heavy reports now refuse while exams are running — and the refusal is only
// worth having if it is cheap: the role, 'סשנים' and ONE tail of 'ממתינים', in
// the REPORTS build that serves them, and not one cell of what the report itself
// would have read ('תוצאות', its archive, the practice sheet).
{
  const fixture = buildFixture({});
  fixture['בוחנים'].push(['מפקד', '999999999', 'pw', 'כן', '99', 'מפקד', 'tokC', iso(NOW + DAY), 0, '', '']);
  fixture['בוחנים'].push(['מפקד מרכז', '999999998', 'pw', 'כן', '98', 'מפקד מרכז', 'tokM', iso(NOW + DAY), 0, '', 'בסיס 6']);
  fixture['תוצאות_ארכיון'] = [RES_HEADER, resRow({ at: NOW - 60 * DAY, id: '910000001' })];
  fixture['ממתינים_ארכיון'] = [PEND_HEADER];
  fixture['תוצאות תרגול'] = [['תאריך', 'מזהה תלמיד', 'שם תלמיד', 'קוד כיתה', 'מצב', 'דרגה', 'ציון', 'סה"כ', 'אחוז',
    'עבר/נכשל', 'זמן', 'נושא', 'שפה', 'פירוט שגויות', 'פירוט לפי נושא', 'טלפון']];
  const env = createEnv({ sheets: fixture, now: NOW, serverFile: 'external_exam_apps_script.reports.js' });
  const reports = [
    ['commanderDashboard', 'handleCommanderDashboard', { examinerId: '999999999', token: 'tokC', dateFrom: '01/09/2026', dateTo: '22/09/2026' }],
    ['centerManagerReport', 'handleCenterManagerReport', { examinerId: '999999998', token: 'tokM' }]
  ];
  for (const [action, handler, params] of reports) {
    env.ctx._sessionRowsMemo = null;   // each request is its own execution in production; the vm context lives on
    env.resetCounters();
    const out = env.json(env.ctx[handler](params));
    const c = env.counters();
    check(action + ' refused during exams: the role, סשנים and one ממתינים tail — nothing of the history', () => {
      assert.equal(out.code, 'exam_hours');
      assert.deepEqual(out.live, { sessions: 1, examinees: 5 });
      for (const name of ['תוצאות', 'תוצאות_ארכיון', 'ממתינים_ארכיון', 'תוצאות תרגול', 'כיתות', 'מבחנים']) {
        const s = c.perSheet[name];
        assert.ok(!s || s.fullReads + s.rangeReads === 0, name + ' was read: ' + JSON.stringify(s));
      }
      assert.equal(c.perSheet['ממתינים'].fullReads, 0, 'never a full read of ממתינים');
      assert.equal(c.perSheet['ממתינים'].rangeReads, 2, 'one tail read = header + tail');
      assert.equal(c.perSheet['סשנים'].fullReads, 1, 'the per-execution סשנים memo');
      assert.ok(c.cellsRead <= 30000, 'cells=' + c.cellsRead);
      assert.equal(c.appends + c.setValues, 0, 'a refusal writes nothing');
    });
  }
  // With no session open the guard does not even read 'ממתינים'.
  for (const row of env.rows('סשנים')) if (row[0] === SESSION) row[10] = false;
  env.ctx._sessionRowsMemo = null;
  env.resetCounters();
  const idle = env.ctx.liveExamActivity();
  check('no open session: the guard costs the סשנים read alone', () => {
    assert.deepEqual({ sessions: idle.sessions, examinees: idle.examinees, error: idle.error }, { sessions: 0, examinees: 0, error: undefined });
    assert.equal(env.counters().reads, 1, 'reads=' + env.counters().reads);
  });
}

// ---- 12. SHEET_HEADERS matches what the code actually uses -----------------
// Review E §5: a sheet that has to be RE-CREATED (new spreadsheet, restored
// backup, the practice migration) is born from SHEET_HEADERS. 'ממתינים'
// declared 15 columns while the code read 15-18 and wrote 19, 'מורים' declared
// 8 with the role at 8 and the site at 9, 'כיתות' declared 7 and createClass
// appended 8 — so a re-created sheet was born up to four columns short and the
// writes landed outside the header.
{
  const env = envWith({});
  const expected = { 'ממתינים': 19, 'ממתינים_ארכיון': 19, 'תוצאות': 30, 'תוצאות_ארכיון': 30, 'מבחנים': 6,
    'מבחנים_ארכיון': 6, 'סשנים': 15, 'מורים': 10, 'כיתות': 8, 'תוצאות תרגול': 16, 'בוחנים': 11, 'הארכות זמן': 7 };
  check('SHEET_HEADERS is as wide as the highest column the code touches', () => {
    for (const name of Object.keys(expected)) {
      const header = env.ctx.SHEET_HEADERS[name];
      assert.ok(header, 'SHEET_HEADERS is missing ' + name);
      assert.equal(header.length, expected[name], name + ' header width');
      for (const cell of header) assert.ok(String(cell || '').trim(), name + ' has an empty header cell');
    }
  });
  check('a sheet created from scratch is born with every column', () => {
    const fresh = createEnv({ sheets: {}, now: NOW });
    const sheet = fresh.ctx.getSheet('ממתינים');
    assert.equal(sheet.getLastColumn(), 19);
    assert.equal(fresh.rows('ממתינים')[0][18], 'סיים במכשיר');
  });
}

console.log('\n' + checks + ' cost checks passed');
