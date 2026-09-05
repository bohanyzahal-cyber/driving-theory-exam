const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const source = fs.readFileSync(path.join(__dirname, '../external_exam_apps_script.js'), 'utf8');

function runtime() {
  const memory = new Map(), logs = [];
  const ctx = { Logger: { log: value => logs.push(String(value)) },
    CacheService: { getScriptCache: () => ({ get: key => memory.get(key) ?? null,
      put: (key, value) => memory.set(key, value) }) } };
  vm.createContext(ctx);
  vm.runInContext(source, ctx);
  ctx.jsonResponse = value => value;
  ctx.verifyExamineeToken = () => ({ valid: true, audioMode: 'off' });
  ctx.EXAM_STRUCTURE_SERVER = { B: { law: 1 } };
  ctx.classifyCategoryServer = category => category;
  ctx.loadLicensePoolServer = () => [{ id: 1, category: 'law', text: 'Synthetic question', answers: ['A', 'B'] }];
  ctx.loadQuestionsForLanguageServer = () => [{ id: 1, category: 'law', text: 'Synthetic question', answers: ['A', 'B'] }];
  return { ctx, logs };
}
const candidate = id => ({ sessionCode: 'SYNTHETIC_CLASS', idNumber: String(id),
  examineeToken: 'synthetic-token', license: 'B', language: 'he' });

test('a whole class can draw questions without sharing one candidate allowance', () => {
  const { ctx } = runtime();
  for (let id = 1; id <= 40; id++) {
    const reply = ctx.handleGetExamQuestions(candidate(id));
    assert.equal(reply.status, 'ok', 'candidate ' + id);
    assert.equal(reply.count, 1);
    assert.equal(reply.questions[0].ci, undefined);
  }
});

test('repeated draws by the same candidate remain rate limited', () => {
  const { ctx } = runtime();
  for (let i = 0; i < 20; i++) assert.equal(ctx.handleGetExamQuestions(candidate(1)).status, 'ok');
  const limited = ctx.handleGetExamQuestions(candidate(1));
  assert.equal(limited.rateLimited, true);
  assert.ok(limited.waitSec > 0 && limited.waitSec <= 60);
  assert.equal(ctx.handleGetExamQuestions(candidate(2)).status, 'ok');
});

test('guest draw allowance is unchanged', () => {
  const { ctx } = runtime();
  for (let i = 0; i < 5; i++) assert.equal(ctx.handleGetExamQuestions({ license: 'B' }).status, 'ok');
  assert.equal(ctx.handleGetExamQuestions({ license: 'B' }).rateLimited, true);
});

test('question lookup rate identity separates candidates but normalizes the same ID', () => {
  const { ctx } = runtime();
  assert.notEqual(ctx.questionRequestRateId(candidate(1), 'examinee'), ctx.questionRequestRateId(candidate(2), 'examinee'));
  assert.equal(ctx.questionRequestRateId(candidate(1), 'examinee'), ctx.questionRequestRateId(candidate('000000001'), 'examinee'));
});

test('cache construction contention is retryable through the real question handler', () => {
  const { ctx } = runtime();
  ctx.loadLicensePoolServer = () => { throw Object.assign(new Error('synthetic contention'), { retryable: true, waitSec: 3 }); };
  const reply = ctx.handleGetExamQuestions(candidate(1));
  assert.equal(reply.code, 'question_cache_busy');
  assert.equal(reply.retryable, true);
  assert.equal(reply.waitSec, 3);
});

test('health identifies build without Sheets, Drive or private parameters', () => {
  const { ctx, logs } = runtime();
  ctx.getSheet = () => { throw new Error('health must not access Sheets'); };
  const result = ctx.doGet({ parameter: { action: 'health', origin: 'examinee-app', token: 'DO_NOT_LOG_ME' } });
  assert.equal(result.status, 'ok');
  assert.equal(result.build, '2026-09-05-r5');
  assert.equal(logs.length, 2);
  assert.ok(logs[0].includes('"phase":"start"'));
  assert.ok(logs[1].includes('"phase":"end"'));
  assert.ok(!logs.join('').includes('DO_NOT_LOG_ME'));
  assert.equal(ctx.doGet({ parameter: { action: 'health' } }).status, 'error');
});

test('timing log does not echo arbitrary action names', () => {
  const { ctx, logs } = runtime();
  ctx.doGet({ parameter: { action: 'PRIVATE_VALUE_DO_NOT_LOG', origin: 'examinee-app' } });
  assert.ok(logs.every(line => !line.includes('PRIVATE_VALUE_DO_NOT_LOG')));
  assert.ok(logs.every(line => line.includes('"action":"unknown"')));
});

test('router logs completion and preserves structured retryable failures', () => {
  const { ctx, logs } = runtime();
  ctx.handleGetSessionInfo = () => { throw Object.assign(new Error('busy'), { retryable: true, waitSec: 3 }); };
  const result = ctx.doGet({ parameter: { action: 'getSessionInfo', origin: 'examinee-app' } });
  assert.equal(result.retryable, true);
  assert.equal(result.code, 'question_cache_busy');
  assert.ok(logs.at(-1).includes('"phase":"end"'));
  const malformed = ctx.doPost({ postData: { contents: '{' } });
  assert.equal(malformed.status, 'error');
  assert.ok(logs.at(-1).includes('"method":"POST"'));
});
