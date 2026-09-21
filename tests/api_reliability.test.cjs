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

test('guest practice allowance is unchanged', () => {
  const e = runtime();
  for (let i = 0; i < 5; i++) assert.equal(get(e, { action: 'startPractice', license: 'B' }).status, 'ok');
  assert.equal(get(e, { action: 'startPractice', license: 'B' }).rateLimited, true);
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
  identical.ctx.defineAction('checkApproval', { methods: ['GET'], auth: 'none', handler: identical.ctx.handleCheckApproval });
  assert.doesNotThrow(() => identical.ctx.ensureLegacyActions());
});

test('health&deep=1 times one cell of our own document and reports a failure instead of throwing', () => {
  const e = runtime();
  const ok = get(e, { action: 'health', deep: '1' });
  assert.equal(ok.status, 'ok');
  assert.equal(ok.build, '2026-09-22-r30');
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
  assert.equal(result.build, '2026-09-22-r30');
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
