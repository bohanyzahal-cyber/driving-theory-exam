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
  assert.equal(ok.build, '2026-09-23-r32');
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
  assert.equal(result.build, '2026-09-23-r32');
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
// The handler is called directly: the API action itself is retired (below).
const checkApproval = (e, id, token) => e.json(e.ctx.handleCheckApproval({ origin: 'examinee-app', sessionCode: SESSION, idNumber: id,
  examineeToken: token === undefined ? 'token-' + id : token }));

test('checkApproval / getExamStatus are retired API actions: an old page is told to reload, never answered', () => {
  const e = approvals([approvalRow(ID, 'waiting')]);
  for (const action of ['checkApproval', 'getExamStatus']) {
    const body = get(e, { action: action, sessionCode: SESSION, idNumber: ID, examineeToken: 'token-' + ID });
    assert.equal(body.code, 'client_outdated', action);
    assert.equal(body.approval, undefined, action + ' leaks nothing');
  }
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
