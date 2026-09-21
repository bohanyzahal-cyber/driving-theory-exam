// Cross-layer contracts: the REAL built server (external_exam_apps_script.js in
// the vm mock), the REAL session-gateway Worker module and the REAL shape of
// the examinee page's submit payload, driven end to end without a network.
//
// Each layer has its own suite with fakes of its neighbours; what can still go
// wrong is the seam between them — a field renamed on one side only. These
// tests take the output of one real layer and feed it to the next:
//   1. sessionSnapshot (server) → Worker → the Worker's answer must equal the
//      server's own checkApproval / getExamStatus answer for the same state,
//      including token mismatch, terminal rows and "not registered".
//   2. startExam (server) → the shape the examinee page consumes.
//   3. the examinee page's submit payload (q/a texts, no score) → server scoring.
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const crypto = require('node:crypto');
const { pathToFileURL } = require('node:url');
const { createEnv, ROOT } = require('./helpers/server_env.cjs');

const WORKER_URL = pathToFileURL(path.join(ROOT, 'cloudflare-workers', 'session-gateway', 'worker.js')).href;
let createGateway;
test.before(async () => { ({ createGateway } = await import(WORKER_URL)); });

const keySandbox = {};
vm.runInNewContext(fs.readFileSync(path.join(ROOT, 'deployment', 'answer_key.gs'), 'utf8'), keySandbox);
const ANSWER_KEY = keySandbox.ANSWER_KEY_BY_LANG;
const HE_BANK = JSON.parse(fs.readFileSync(path.join(ROOT, 'bank', 'he.json'), 'utf8'));
const HE_BY_ID = Object.fromEntries(HE_BANK.map(e => [e.id, e]));

const SESSION = 'ABC12345';
const GATEWAY_KEY = 'shared-secret-for-tests';
const PENDING_HEADER = ['קוד סשן', 'ת.ז.', 'שם', 'טלפון', 'זמן הרשמה', 'סטטוס', 'שפה', 'אוכלוסיה', 'דרגה', 'שמע',
  'הארכת זמן', 'התחלת מבחן', 'טוקן נבחן', 'ספירת DQ', 'מסך נוסף', 'ספירת אזהרות', 'אזהרה אחרונה', 'אתר', 'סיים במכשיר'];
const RESULTS_HEADER = ['תאריך', 'ת.ז.', 'שם', 'טלפון', 'דרגה', 'ציון', 'אחוז', 'עבר/נכשל', 'זמן', 'בוחן', 'אתר',
  'כיתה', 'שפה', 'קוד סשן', 'ניסיון', 'פירוט שגויות', 'נשלח?', 'פסול?', 'קישור וואטסאפ', 'אוכלוסיה', 'תוקן?', 'שמע',
  'מאומת', 'חשוד', 'dqEventId', 'תוקן ע"י', 'סיבת תיקון', 'תאריך תיקון', 'מסלול שפות', 'מכשיר'];
const EXAMS_HEADER = ['קוד סשן', 'ת.ז.', 'שאלות JSON', 'זמן רישום', 'שפה', 'שגויות לא מאומתות'];

function pendingRow(id, status, over) {
  const row = Array(19).fill('');
  row[0] = SESSION; row[1] = id; row[2] = 'נבחן ' + id; row[3] = '0501234567';
  row[4] = '2026-09-22T06:00:00Z'; row[5] = status; row[6] = 'he'; row[8] = 'B'; row[9] = 'off'; row[12] = 'tok-' + id;
  return Object.assign(row, over || {});
}
function serverEnv(pending, extensions) {
  return createEnv({
    sheets: {
      'ממתינים': [PENDING_HEADER, ...pending],
      'מבחנים': [EXAMS_HEADER],
      'תוצאות': [RESULTS_HEADER],
      'הארכות זמן': [['זמן', 'קוד סשן', 'ת.ז.', 'שם', 'דקות', 'סיבה', 'בוחן'], ...(extensions || [])]
    },
    properties: { GATEWAY_KEY },
    sources: [path.join('deployment', 'answer_key.gs')]
  });
}
const get = (e, params) => e.json(e.ctx.doGet({ parameter: Object.assign({ origin: 'examinee-app' }, params) }));
const post = (e, body) => e.json(e.ctx.doPost({ postData: { contents: JSON.stringify(Object.assign({ origin: 'examinee-app' }, body)) } }));

// The Worker's upstream IS the server: every fetch it makes runs doGet in the vm.
function gatewayOver(env) {
  const calls = [];
  const fetchFn = async url => {
    calls.push(String(url));
    const u = new URL(url);
    const params = Object.fromEntries(u.searchParams.entries());
    const body = env.json(env.ctx.doGet({ parameter: params }));
    return new Response(JSON.stringify(body), { status: 200 });
  };
  const gateway = createGateway({ fetch: fetchFn, caches: undefined, now: () => Date.now(), env: { API_URL: 'https://api.test/exec', GATEWAY_KEY } });
  const poll = async params => {
    const res = await gateway(new Request('https://gw.test/v1/poll?' + new URLSearchParams(params).toString()));
    return res.json();
  };
  return { poll, calls };
}

test('the gateway answers exactly what the server answers, for every row state', async () => {
  const pending = [
    pendingRow('900000001', 'waiting'),
    pendingRow('900000002', 'approved', { 10: '1.25', 9: 'on' }),
    pendingRow('900000003', 'in_exam', { 11: '2026-09-22T06:10:00Z' }),
    pendingRow('900000004', 'completed'),
    pendingRow('900000005', 'rejected'),
    pendingRow('900000006', 'dq_confirmed'),
    pendingRow('900000007', 'disqualified'),
    pendingRow('900000008', 'waiting', { 12: '' })   // legacy row without a token
  ];
  const env = serverEnv(pending, [['2026-09-22T06:20:00Z', SESSION, '900000003', 'x', 7, 'evacuation', 'examiner']]);
  const gw = gatewayOver(env);
  for (const row of pending) {
    const id = row[1], token = row[12];
    const direct = get(env, { action: 'checkApproval', sessionCode: SESSION, idNumber: id, examineeToken: token });
    const viaGateway = await gw.poll({ kind: 'approval', sessionCode: SESSION, idNumber: id, examineeToken: token });
    assert.deepEqual(viaGateway, direct, 'approval answer for ' + row[5]);
    const directStatus = get(env, { action: 'getExamStatus', sessionCode: SESSION, idNumber: id, examineeToken: token });
    const gatewayStatus = await gw.poll({ kind: 'status', sessionCode: SESSION, idNumber: id, examineeToken: token });
    assert.deepEqual(gatewayStatus, directStatus, 'status answer for ' + row[5]);
  }
  // in_exam carries the examiner's extra minutes on both routes
  const s = await gw.poll({ kind: 'status', sessionCode: SESSION, idNumber: '900000003', examineeToken: 'tok-900000003' });
  assert.equal(s.extraMinutes, 7);
  // the approved row carries the extended duration on both routes
  const a = await gw.poll({ kind: 'approval', sessionCode: SESSION, idNumber: '900000002', examineeToken: 'tok-900000002' });
  assert.equal(a.examMinutes, 50);
  assert.equal(a.audioMode, 'on');
  // Eight examinees polled twice each → the server ran the snapshot ONCE for
  // the fresh window, plus ONE forced re-read when a terminal row produced
  // 'not registered' (the Worker mirrors the server's re-read-on-miss rule, at
  // most once per session per 10 s). Sixteen polls, two executions.
  assert.equal(gw.calls.filter(u => u.includes('action=sessionSnapshot')).length, 2);
});

test('a wrong token and an unknown examinee get the same answer from both routes', async () => {
  const env = serverEnv([pendingRow('900000001', 'approved')]);
  const gw = gatewayOver(env);
  const wrongDirect = get(env, { action: 'checkApproval', sessionCode: SESSION, idNumber: '900000001', examineeToken: 'wrong' });
  const wrongGateway = await gw.poll({ kind: 'approval', sessionCode: SESSION, idNumber: '900000001', examineeToken: 'wrong' });
  assert.deepEqual(wrongGateway, wrongDirect);
  assert.equal(wrongGateway.examineeTokenError, 'mismatch');
  const missingDirect = get(env, { action: 'checkApproval', sessionCode: SESSION, idNumber: '900000099', examineeToken: 'x' });
  const missingGateway = await gw.poll({ kind: 'approval', sessionCode: SESSION, idNumber: '900000099', examineeToken: 'x' });
  assert.deepEqual(missingGateway, missingDirect);
  assert.equal(missingGateway.status, 'error');
});

test('the snapshot never leaks names, phones or tokens, and refuses a wrong key', () => {
  const env = serverEnv([pendingRow('900000001', 'approved')]);
  const snap = get(env, { action: 'sessionSnapshot', sessionCode: SESSION, gatewayKey: GATEWAY_KEY, origin: 'gateway' });
  assert.equal(snap.status, 'ok');
  const text = JSON.stringify(snap);
  assert.ok(!text.includes('נבחן 900000001') && !text.includes('0501234567') && !text.includes('tok-900000001'));
  assert.equal(snap.rows[0].tokenHash, crypto.createHash('sha256').update('tok-900000001').digest('hex'));
  const denied = get(env, { action: 'sessionSnapshot', sessionCode: SESSION, gatewayKey: 'nope', origin: 'gateway' });
  assert.equal(denied.code, 'gateway_denied');
});

test('startExam hands the page ids the static bank can render, in the server order, and getSessionInfo names the gateway', () => {
  const env = createEnv({
    sheets: {
      'ממתינים': [PENDING_HEADER, pendingRow('900000001', 'approved')],
      'מבחנים': [EXAMS_HEADER], 'תוצאות': [RESULTS_HEADER],
      'סשנים': [['קוד', 'בוחן', 'שם', 'אתר', 'כיתה', 'דרגה', 'שפה', 'שמע', 'נוצר', 'תקף עד', 'פעיל', 'מכסות', 'מאושרים', 'אחראי', 'אוכלוסיה'],
        [SESSION, '123456789', 'בוחן', 'בדיקת נתונים', '1', 'B', 'he', 'off', '2026-09-22T05:00:00Z', '2026-09-23T05:00:00Z', true, '', '', '', '']]
    },
    properties: { GATEWAY_URL: 'https://session-gateway.example.workers.dev' },
    sources: [path.join('deployment', 'answer_key.gs')]
  });
  const info = get(env, { action: 'getSessionInfo', sessionCode: SESSION });
  assert.equal(info.session.gateway.url, 'https://session-gateway.example.workers.dev');
  assert.match(String(info.session.build), /^2026-/);
  const started = post(env, { action: 'startExam', sessionCode: SESSION, idNumber: '900000001', examineeToken: 'tok-900000001', language: 'he', license: 'B' });
  assert.equal(started.status, 'ok');
  assert.equal(started.questions.length, 30);
  for (const q of started.questions) {
    const entry = HE_BY_ID[q.id];
    assert.ok(entry, 'bank/he.json has id ' + q.id);
    assert.equal(entry.a.length, 4);
    assert.deepEqual(q.order.slice().sort(), [0, 1, 2, 3], 'order is a permutation');
    assert.ok(['בטיחות', 'הכרת הרכב', 'חוק', 'תמרורים', 'ספציפי'].includes(q.topic));
  }
  assert.equal(env.sheet('ממתינים').rows[1][5], 'in_exam');
});

test('the examinee page payload (texts, no score) is scored by the server and lands as a 30-column verified row', () => {
  const env = createEnv({
    sheets: { 'ממתינים': [PENDING_HEADER, pendingRow('900000001', 'approved')], 'מבחנים': [EXAMS_HEADER], 'תוצאות': [RESULTS_HEADER] },
    sources: [path.join('deployment', 'answer_key.gs')]
  });
  const started = post(env, { action: 'startExam', sessionCode: SESSION, idNumber: '900000001', examineeToken: 'tok-900000001', language: 'he', license: 'B' });
  // Build the payload exactly as examinee.html does: displayed texts per the
  // server order; 26 correct picks (via the key — the page itself never has it),
  // 3 wrong, 1 unanswered.
  const answers = started.questions.map((q, i) => {
    const entry = HE_BY_ID[q.id];
    const displayed = q.order.map(k => entry.a[k]);
    const correctDisplayed = q.order.indexOf(ANSWER_KEY.he[q.id]);
    let selected = correctDisplayed;
    if (i < 3) selected = (correctDisplayed + 1) % 4;
    if (i === 3) selected = -1;
    return { qIdx: i, selected, langAtAnswer: 'he', q: entry.t, a: displayed };
  });
  const res = post(env, {
    action: 'submitResult', examineeToken: 'tok-900000001', sessionCode: SESSION, idNumber: '900000001',
    fullName: 'נבחן בדיקה', phone: '0501234567', license: 'B', language: 'he', languageHistory: ['he'],
    total: 30, time: "12 דק' 03 שנ'", examinerName: 'בוחן', site: 'בדיקת נתונים', classroom: '1', population: 'צבא',
    audioMode: 'off', device: 'desktop', answers, clientLog: [{ t: 1, e: 'start_exam', d: 'B/he' }]
  });
  assert.equal(res.status, 'ok');
  const row = env.sheet('תוצאות').rows[1];
  assert.equal(row.length, 30);
  assert.equal(row[5], '26/30');
  assert.equal(row[7], 'עבר');
  assert.equal(row[22], 'מאומת');
  const details = String(row[15]);
  assert.equal((details.match(/מזהה שאלה:/g) || []).length, 4, 'three wrong + one unanswered');
  assert.ok(details.includes('לא נענתה'));
  assert.ok(details.includes(HE_BY_ID[started.questions[0].id].t), 'the displayed question text reached the certificate');
  assert.ok(!details.includes('undefined'));
  assert.equal(env.sheet('ממתינים').rows[1][5], 'completed');
});
