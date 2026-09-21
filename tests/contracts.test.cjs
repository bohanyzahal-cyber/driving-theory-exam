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
//   3. the signed bank grant (server, HMAC) → the real Worker's /v1/bank: the
//      texts are private assets now, so "the exam can show its questions" is a
//      property of TWO layers agreeing on one signature (DESIGN §11).
//   4. the examinee page's submit payload (q/a texts, no score) → server scoring.
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

// The texts are the Worker's PRIVATE assets — not the repo, not Pages — so the
// tests read them where `node tools/build_bank.js` writes them (DESIGN §11.1).
const ASSETS_DIR = path.join(ROOT, 'cloudflare-workers', 'session-gateway', 'assets');
const HE_BANK_FILE = path.join(ASSETS_DIR, 'bank', 'he.json');
if (!fs.existsSync(HE_BANK_FILE)) {
  throw new Error('run node tools/build_bank.js first (missing ' + HE_BANK_FILE + ')');
}
const HE_BANK_TEXT = fs.readFileSync(HE_BANK_FILE, 'utf8');
const HE_BY_ID = Object.fromEntries(JSON.parse(HE_BANK_TEXT).map(e => [e.id, e]));

// The Workers Static Assets binding: addressed by URL, answering with the file
// or a 404. `run_worker_first` means nothing else can reach these bytes.
function assetsBinding() {
  return {
    async fetch(reqOrUrl) {
      const href = typeof reqOrUrl === 'string' ? reqOrUrl : reqOrUrl.url;
      const rel = decodeURIComponent(new URL(href).pathname).replace(/^\/+/, '');
      const file = path.join(ASSETS_DIR, rel);
      if (!file.startsWith(ASSETS_DIR) || !fs.existsSync(file)) return new Response('not found', { status: 404 });
      return new Response(fs.readFileSync(file), { status: 200, headers: { 'Content-Type': 'application/json' } });
    }
  };
}

// The vm mock runs on a fixed clock; the Worker must be given the same one or
// an expiry assertion would depend on the day the suite is run.
const ENV_NOW = Date.parse('2026-09-22T06:30:00Z');
const SESSION = 'ABC12345';
const GATEWAY_KEY = 'shared-secret-for-tests';
const GATEWAY_URL = 'https://session-gateway.example.workers.dev';
// Every server environment here is a CONFIGURED one: startExam/startPractice
// refuse to write anything without a gateway to fetch the texts from (§11.2).
const GATEWAY_PROPS = { GATEWAY_KEY, GATEWAY_URL };
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
    properties: Object.assign({}, GATEWAY_PROPS),
    sources: [path.join('deployment', 'answer_key.gs')]
  });
}
const get = (e, params) => e.json(e.ctx.doGet({ parameter: Object.assign({ origin: 'examinee-app' }, params) }));
// The two direct poll ACTIONS are retired from the API (an old page is told to
// reload); their handlers stay as the reference every gateway answer must match.
const direct = (e, handler, params) => e.json(e.ctx[handler](Object.assign({ origin: 'examinee-app' }, params)));
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
  const gateway = createGateway({ fetch: fetchFn, caches: undefined, now: () => ENV_NOW,
    env: { API_URL: 'https://api.test/exec', GATEWAY_KEY, ASSETS: assetsBinding() } });
  const poll = async params => {
    const res = await gateway(new Request('https://gw.test/v1/poll?' + new URLSearchParams(params).toString()));
    return res.json();
  };
  return { poll, calls, gateway };
}

// A Worker that only serves the bank: its upstream must never be reached, so a
// fetch here is itself the failure.
function bankWorker(key) {
  return createGateway({
    fetch: async () => { throw new Error('the bank must not call Apps Script'); },
    caches: undefined, now: () => ENV_NOW,
    env: { API_URL: 'https://api.test/exec', GATEWAY_KEY: key || GATEWAY_KEY, ASSETS: assetsBinding() }
  });
}
// `bank` is the {url, grant, exp} the server just issued — the client appends
// nothing but the query the page needs.
const askBank = (worker, bank, query) =>
  worker(new Request(bank.url + '/v1/bank?grant=' + encodeURIComponent(bank.grant) + (query || '')));
const askBankFull = (worker, bank, lang) =>
  worker(new Request(bank.url + '/v1/bank/full?grant=' + encodeURIComponent(bank.grant) + '&lang=' + lang));

test('the gateway answers exactly what the server answers, for every row state', async () => {
  const pending = [
    pendingRow('900000001', 'waiting'),
    pendingRow('900000002', 'approved', { 10: '1.25', 9: 'on' }),
    pendingRow('900000003', 'in_exam', { 11: '2026-09-22T06:10:00Z' }),
    pendingRow('900000004', 'completed'),
    pendingRow('900000005', 'rejected'),
    pendingRow('900000006', 'dq_confirmed'),
    pendingRow('900000007', 'disqualified'),
    pendingRow('900000008', 'waiting', { 12: '' }),  // legacy row without a token
    pendingRow('900000009', 'cancelled'),
    // The shared-ID incident (see scanApprovalRows): rejected at 17:47, the
    // second examinee on the same id cancelled at 18:05 — the NEWEST decides.
    pendingRow('900000010', 'rejected'),
    pendingRow('900000010', 'cancelled'),
    // ...and a live row still outranks a decision written above it.
    pendingRow('900000011', 'rejected'),
    pendingRow('900000011', 'waiting')
  ];
  const env = serverEnv(pending, [['2026-09-22T06:20:00Z', SESSION, '900000003', 'x', 7, 'evacuation', 'examiner']]);
  const gw = gatewayOver(env);
  for (const row of pending) {
    const id = row[1], token = row[12];
    const directApproval = direct(env, 'handleCheckApproval', { sessionCode: SESSION, idNumber: id, examineeToken: token });
    const viaGateway = await gw.poll({ kind: 'approval', sessionCode: SESSION, idNumber: id, examineeToken: token });
    assert.deepEqual(viaGateway, directApproval, 'approval answer for ' + row[5]);
    const directStatus = direct(env, 'handleGetExamStatus', { sessionCode: SESSION, idNumber: id, examineeToken: token });
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
  // The examiner's decision on a registration nobody is testing under any more:
  // the status itself, and NOTHING else — no audioMode, no examMinutes.
  const decided = id => gw.poll({ kind: 'approval', sessionCode: SESSION, idNumber: id, examineeToken: 'tok-' + id });
  assert.deepEqual(await decided('900000005'), { status: 'ok', approval: 'rejected' });
  assert.deepEqual(await decided('900000009'), { status: 'ok', approval: 'cancelled' });
  assert.deepEqual(await decided('900000010'), { status: 'ok', approval: 'cancelled' },
    'the newest row decides: the third visitor is told "cancelled", never "rejected"');
  assert.equal((await decided('900000011')).approval, 'waiting', 'a live row still wins');
  // A finished exam still answers "not registered" — there is nothing to say to
  // a device that is still polling one.
  assert.equal((await decided('900000004')).message, 'לא נמצא רישום');
  assert.equal((await decided('900000007')).message, 'לא נמצא רישום');
  // Eight examinees polled twice each → the server ran the snapshot ONCE for
  // the fresh window, plus ONE forced re-read when a terminal row produced
  // 'not registered' (the Worker mirrors the server's re-read-on-miss rule, at
  // most once per session per 10 s). Sixteen polls, two executions.
  assert.equal(gw.calls.filter(u => u.includes('action=sessionSnapshot')).length, 2);
});

test('a wrong token and an unknown examinee get the same answer from both routes', async () => {
  const env = serverEnv([pendingRow('900000001', 'approved')]);
  const gw = gatewayOver(env);
  const wrongDirect = direct(env, 'handleCheckApproval', { sessionCode: SESSION, idNumber: '900000001', examineeToken: 'wrong' });
  const wrongGateway = await gw.poll({ kind: 'approval', sessionCode: SESSION, idNumber: '900000001', examineeToken: 'wrong' });
  assert.deepEqual(wrongGateway, wrongDirect);
  assert.equal(wrongGateway.examineeTokenError, 'mismatch');
  const missingDirect = direct(env, 'handleCheckApproval', { sessionCode: SESSION, idNumber: '900000099', examineeToken: 'x' });
  const missingGateway = await gw.poll({ kind: 'approval', sessionCode: SESSION, idNumber: '900000099', examineeToken: 'x' });
  assert.deepEqual(missingGateway, missingDirect);
  assert.equal(missingGateway.status, 'error');

  // The token check applies to the decision below a dead registration exactly
  // as it applies to a live row: a device holding another examinee's stale
  // token is refused, not told "you were rejected".
  const decidedEnv = serverEnv([pendingRow('900000001', 'rejected')]);
  const decidedGw = gatewayOver(decidedEnv);
  const params = { sessionCode: SESSION, idNumber: '900000001', examineeToken: 'stale-token' };
  const stolenDirect = direct(decidedEnv, 'handleCheckApproval', params);
  const stolenGateway = await decidedGw.poll(Object.assign({ kind: 'approval' }, params));
  assert.deepEqual(stolenGateway, stolenDirect);
  assert.equal(stolenGateway.examineeTokenError, 'mismatch');
});

test('the snapshot never leaks names, phones or tokens, and refuses a wrong key', () => {
  // The row carries every counter and flag the examiner board displays (r31,
  // DESIGN §13.6: the board waits on the Worker's fingerprint of these), so the
  // "nothing identifying" rule is asserted on a row where they are all set.
  const loaded = pendingRow('900000001', 'in_exam');
  loaded[13] = 2;        // N: ספירת DQ
  loaded[14] = 'כן';     // O: מסך נוסף
  loaded[15] = 3;        // P: ספירת אזהרות
  loaded[16] = 'החלפת חלון';
  loaded[18] = '2026-09-22T07:10:00Z';   // S: סיים במכשיר
  const env = serverEnv([loaded]);
  const snap = get(env, { action: 'sessionSnapshot', sessionCode: SESSION, gatewayKey: GATEWAY_KEY, origin: 'gateway' });
  assert.equal(snap.status, 'ok');
  const text = JSON.stringify(snap);
  assert.ok(!text.includes('נבחן 900000001') && !text.includes('0501234567') && !text.includes('tok-900000001'));
  assert.ok(!text.includes('החלפת חלון'), 'the warning REASON is free text an examinee typed into — it stays on the sheet');
  assert.equal(snap.rows[0].tokenHash, crypto.createHash('sha256').update('tok-900000001').digest('hex'));
  assert.deepEqual(Object.keys(snap.rows[0]).sort(),
    ['audio', 'dq', 'examMinutes', 'ext', 'extraMinutes', 'fin', 'id', 'status', 'tokenHash', 'warn']);
  assert.deepEqual([snap.rows[0].warn, snap.rows[0].fin, snap.rows[0].ext, snap.rows[0].dq], [3, 1, 1, 2]);
  const denied = get(env, { action: 'sessionSnapshot', sessionCode: SESSION, gatewayKey: 'nope', origin: 'gateway' });
  assert.equal(denied.code, 'gateway_denied');
});

function sessionEnv(properties) {
  return createEnv({
    sheets: {
      'ממתינים': [PENDING_HEADER, pendingRow('900000001', 'approved')],
      'מבחנים': [EXAMS_HEADER], 'תוצאות': [RESULTS_HEADER],
      'סשנים': [['קוד', 'בוחן', 'שם', 'אתר', 'כיתה', 'דרגה', 'שפה', 'שמע', 'נוצר', 'תקף עד', 'פעיל', 'מכסות', 'מאושרים', 'אחראי', 'אוכלוסיה'],
        [SESSION, '123456789', 'בוחן', 'בדיקת נתונים', '1', 'B', 'he', 'off', '2026-09-22T05:00:00Z', '2026-09-23T05:00:00Z', true, '', '', '', '']]
    },
    properties: Object.assign({}, GATEWAY_PROPS, properties || {}),
    sources: [path.join('deployment', 'answer_key.gs')]
  });
}
const startOne = env => post(env, { action: 'startExam', sessionCode: SESSION, idNumber: '900000001',
  examineeToken: 'tok-900000001', language: 'he', license: 'B' });

test('startExam hands the page ids the bank can render, in the server order, and getSessionInfo names the gateway', () => {
  const env = sessionEnv();
  const info = get(env, { action: 'getSessionInfo', sessionCode: SESSION });
  assert.equal(info.session.gateway.url, GATEWAY_URL);
  assert.match(String(info.session.build), /^2026-/);
  const started = startOne(env);
  assert.equal(started.status, 'ok');
  assert.equal(started.questions.length, 30);
  for (const q of started.questions) {
    const entry = HE_BY_ID[q.id];
    assert.ok(entry, 'assets/bank/he.json has id ' + q.id);
    assert.equal(entry.a.length, 4);
    assert.deepEqual(q.order.slice().sort(), [0, 1, 2, 3], 'order is a permutation');
    assert.ok(['בטיחות', 'הכרת הרכב', 'חוק', 'תמרורים', 'ספציפי'].includes(q.topic));
  }
  assert.equal(env.sheet('ממתינים').rows[1][5], 'in_exam');
});

// One Worker, always on: health says whether it is wired up, and nothing else
// decides where the fleet polls (Yossi, 21/09: no partial switches).
test('health reports the Worker wiring as booleans only', () => {
  const on = get(sessionEnv(), { action: 'health' });
  assert.deepEqual(on.gateway, { url: true, key: true });
  const info = get(sessionEnv(), { action: 'getSessionInfo', sessionCode: SESSION });
  assert.equal(info.session.gateway.url, GATEWAY_URL, 'the fleet polls the Worker');
});

test('the grant startExam issues opens exactly those 30 questions in the real Worker', async () => {
  const env = sessionEnv();
  const started = startOne(env);
  const ids = started.questions.map(q => q.id);
  assert.equal(typeof started.bank.grant, 'string');
  assert.equal(started.bank.url, GATEWAY_URL);
  assert.equal(started.bank.exp, ENV_NOW + 4 * 3600 * 1000);

  const res = await askBank(bankWorker(), started.bank);
  assert.equal(res.status, 200);
  const body = await res.json();
  assert.equal(body.status, 'ok');
  assert.deepEqual(body.missing, []);
  assert.deepEqual(body.questions.map(q => q.id), ids, 'exactly the granted ids, in the granted order');
  for (const q of body.questions) {
    const he = q.l.he, fromBank = HE_BY_ID[q.id];
    assert.equal(he.t, fromBank.t, 'question ' + q.id + ' text');
    assert.deepEqual(he.a, fromBank.a, 'question ' + q.id + ' answers');
    assert.equal(he.i || '', fromBank.i || '', 'question ' + q.id + ' image');
  }
  // An exam grant is a WHITELIST: the ids it names, and nothing else.
  const extra = await askBank(bankWorker(), started.bank, '&ids=' + (ids[0] + 1));
  const extraBody = await extra.json();
  assert.deepEqual((extraBody.questions || []).map(q => q.id), ids, 'an exam grant ignores an id list from the client');
});

test('a grant signed with another key is refused by the Worker', async () => {
  const impostor = createEnv({
    sheets: { 'ממתינים': [PENDING_HEADER, pendingRow('900000001', 'approved')], 'מבחנים': [EXAMS_HEADER], 'תוצאות': [RESULTS_HEADER] },
    properties: { GATEWAY_KEY: 'not-the-shared-secret', GATEWAY_URL },
    sources: [path.join('deployment', 'answer_key.gs')]
  });
  const forged = startOne(impostor).bank;
  const res = await askBank(bankWorker(), forged);
  assert.equal(res.status, 403);
  assert.equal((await res.json()).code, 'grant_invalid');
  // ...and a grant whose payload was edited after signing fails the same way.
  const real = startOne(sessionEnv()).bank;
  const tampered = Object.assign({}, real, { grant: real.grant.replace(/^[^.]+/, m => m.slice(0, -1) + (m.slice(-1) === 'A' ? 'B' : 'A')) });
  assert.equal((await askBank(bankWorker(), tampered)).status, 403);
});

test('startPractice returns a practice grant the Worker accepts', async () => {
  const env = sessionEnv();
  const practice = get(env, { action: 'startPractice', license: 'B', language: 'he', mode: 'category',
    categoryFilter: 'תמרורים', maxCount: '50', classCode: 'CLS1', studentId: 'stu-1' });
  assert.equal(practice.status, 'ok');
  assert.ok(practice.count <= 30, 'a draw never exceeds one Worker request: ' + practice.count);
  assert.equal(practice.bank.exp, ENV_NOW + 2 * 3600 * 1000);
  const body = await (await askBank(bankWorker(), practice.bank)).json();
  assert.equal(body.status, 'ok');
  assert.deepEqual(body.questions.map(q => q.id), practice.questions.map(q => q.id));
  // A practice grant is not an examiner grant: it cannot open a whole language.
  assert.equal((await askBankFull(bankWorker(), practice.bank, 'he')).status, 403);
});

test('bankGrant needs an examiner token, and its grant opens the whole Hebrew bank', async () => {
  const env = createEnv({
    sheets: { 'ממתינים': [PENDING_HEADER], 'מבחנים': [EXAMS_HEADER], 'תוצאות': [RESULTS_HEADER],
      'בוחנים': [['שם', 'ת.ז.', 'סיסמה', 'פעיל', 'מספר', 'תפקיד', 'טוקן', 'תוקף', 'כשלונות', 'נעילה', 'אתרים'],
        ['בוחן', '123456789', 'pw', 'כן', '7', 'בוחן', 'tokE', '2026-09-23T06:00:00Z', 0, '', '']] },
    properties: Object.assign({}, GATEWAY_PROPS),
    sources: [path.join('deployment', 'answer_key.gs')]
  });
  assert.equal(get(env, { action: 'bankGrant', origin: 'examiner-app', examinerId: '123456789' }).tokenExpired, true);
  const granted = get(env, { action: 'bankGrant', origin: 'examiner-app', examinerId: '123456789', token: 'tokE' });
  assert.equal(granted.status, 'ok');
  assert.equal(granted.bank.url, GATEWAY_URL);
  assert.equal(granted.bank.exp, ENV_NOW + 8 * 3600 * 1000);

  const worker = bankWorker();
  const full = await askBankFull(worker, granted.bank, 'he');
  assert.equal(full.status, 200);
  assert.equal(await full.text(), HE_BANK_TEXT, 'the Worker streams the asset unchanged');
  // The same grant also answers a named id list — the commander's wrong-question table.
  const some = await (await askBank(worker, granted.bank, '&ids=1,2&langs=he')).json();
  assert.deepEqual(some.questions.map(q => q.id), [1, 2]);
  assert.equal(some.questions[0].l.he.t, HE_BY_ID[1].t);
});

test('the examinee page payload (texts, no score) is scored by the server and lands as a 30-column verified row', () => {
  const env = createEnv({
    sheets: { 'ממתינים': [PENDING_HEADER, pendingRow('900000001', 'approved')], 'מבחנים': [EXAMS_HEADER], 'תוצאות': [RESULTS_HEADER] },
    properties: Object.assign({}, GATEWAY_PROPS),
    sources: [path.join('deployment', 'answer_key.gs')]
  });
  const started = startOne(env);
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
