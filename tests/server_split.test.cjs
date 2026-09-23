// The r31 build and the two-project split (DESIGN_2026-09-21 §13.1, §13.3, §13.6).
//
// Everything here runs the REAL generated files. The split is the one change of
// this release that cannot be seen by looking at one layer: an action lives in a
// table in 30_api.js, its handler lives in a module that a build target may or
// may not include, and a PAGE decides which of two URLs to send it to. Three
// lists that must agree, and nothing at runtime would say they do not — a page
// would simply talk to a server that answers "wrong_deployment" forever.
// So this suite pins all three against each other:
//   1. the three committed files are what the sources build (`--check`)
//   2. the exam file alone serves exam+both and refuses reports, and vice versa
//   3. the packed question index unpacks to deployment/question_index.json
//   4. a standalone project opens the exam spreadsheet by EXAM_SPREADSHEET_ID
//   5. ACTION_TARGETS === ExamTransport.REPORTS_ACTIONS, and covers every action
//   6. every action a page sends is served by the URL that page routes to
//   7. sessionSnapshot carries what the Worker's fingerprint needs, and the one
//      writer that used to leave a stale snapshot behind now drops it
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const { spawnSync } = require('node:child_process');
const { createEnv, ROOT } = require('./helpers/server_env.cjs');

const MONOLITH = 'external_exam_apps_script.js';
const EXAM_FILE = 'external_exam_apps_script.exam.js';
const REPORTS_FILE = 'external_exam_apps_script.reports.js';

// The client's half of the contract, loaded the way tests/transport.test.cjs
// loads it: the real module in a vm, no browser.
function transportModule() {
  const ctx = { console: { log() {}, warn() {}, error() {} }, setTimeout, clearTimeout, setInterval,
    clearInterval, fetch: async () => { throw new Error('no network in this suite'); }, Date, Math, JSON };
  ctx.window = ctx;
  ctx.global = ctx;
  vm.createContext(ctx);
  vm.runInContext(fs.readFileSync(path.join(ROOT, 'shared', 'transport.js'), 'utf8'), ctx, { filename: 'transport.js' });
  return ctx.ExamTransport;
}

// A server environment with enough of the document for health and a dispatch to
// answer. Nothing here is exercised past the target check.
function env(serverFile, extra) {
  return createEnv(Object.assign({
    serverFile: serverFile,
    sheets: { 'אתרים': [['שם אתר', 'מזהה', 'טלפון מנהל', 'כיתות'], ['בדיקת נתונים', '1', '', '']] },
    properties: { GATEWAY_KEY: 'k', GATEWAY_URL: 'https://gw.test' }
  }, extra || {}));
}
const dispatch = (e, action, method) => e.json(e.ctx.dispatchApiAction(method || 'GET', action, { origin: 'examiner-app' }));
// Arrays and objects built inside a vm carry that vm's prototypes, and
// deepStrictEqual compares them; plain() brings a value back to this realm.
const plain = v => JSON.parse(JSON.stringify(v));

// ---- 1. the committed files are the ones the sources build -----------------

test('the three committed outputs equal a fresh build', () => {
  const run = spawnSync(process.execPath, [path.join(ROOT, 'tools', 'build_server.js'), '--check'],
    { cwd: ROOT, encoding: 'utf8' });
  assert.equal(run.status, 0, 'node tools/build_server.js --check failed:\n' + run.stdout + run.stderr);
  for (const name of [MONOLITH, EXAM_FILE, REPORTS_FILE]) {
    assert.match(run.stdout, new RegExp(name.replace(/\./g, '\\.') + ' is up to date'));
  }
});

test('stripping left no comment but the file header, and no marker behind', () => {
  const acorn = require(path.join(ROOT, 'tools', 'vendor', 'acorn.js'));
  for (const name of [MONOLITH, EXAM_FILE, REPORTS_FILE]) {
    const text = fs.readFileSync(path.join(ROOT, name), 'utf8');
    assert.ok(!text.includes('@@QUESTION_INDEX@@'), name + ' still carries the index marker');
    assert.ok(!text.includes('@@API_DEPLOYMENT@@'), name + ' still carries the deployment marker');
    assert.ok(text.includes('© 2026'), name + ' lost the copyright notice');
    // Found with the same tokenizer the build strips with, so a `//` inside a
    // string literal (the classic regex-stripper bug, in the other direction)
    // cannot be mistaken for a comment.
    const comments = [];
    for (const token of acorn.tokenizer(text, { ecmaVersion: 2022,
      onComment: (block, body, start, end) => comments.push({ body, end }) })) { void token; }
    // Everything that survives is the FILE HEADER: the generated block and the
    // copyright block of 00_config.js, all of it above the first statement.
    // 1,800+ source comments are gone; a comment anywhere below is a bug.
    const firstStatement = acorn.parse(text, { ecmaVersion: 2022, sourceType: 'script' }).body[0].start;
    const strays = comments.filter(c => c.end > firstStatement).map(c => c.body.slice(0, 60));
    assert.deepEqual(strays, [], name + ' still carries comments below the header');
    assert.ok(comments.some(c => c.body.includes('© 2026')), name + ' lost the copyright comment');
    assert.ok(comments.some(c => c.body.includes('GENERATED FILE')), name + ' lost the generated header');
  }
});

// The check that the split itself makes necessary: a `both` or `reports` module
// calling a function that only the exam modules carry compiles, passes --check,
// and then throws ReferenceError in production at the worst possible moment.
// Free identifiers of each built file, scope-unaware on purpose (it collects
// EVERY name the file declares anywhere, so it can only under-report — never a
// false alarm).
test('no built file references a function it does not contain', () => {
  const acorn = require(path.join(ROOT, 'tools', 'vendor', 'acorn.js'));
  // Apps Script's own globals, plus the ECMAScript ones.
  const PROVIDED = new Set(['SpreadsheetApp', 'CacheService', 'PropertiesService', 'Utilities', 'LockService',
    'ScriptApp', 'Session', 'Logger', 'ContentService', 'HtmlService', 'MimeType', 'UrlFetchApp', 'DriveApp',
    'MailApp', 'GmailApp', 'Date', 'Math', 'JSON', 'Object', 'Array', 'String', 'Number', 'Boolean', 'RegExp',
    'Error', 'isNaN', 'isFinite', 'parseInt', 'parseFloat', 'encodeURIComponent', 'decodeURIComponent',
    'globalThis', 'console', 'undefined', 'NaN', 'Infinity', 'Map', 'Set', 'Promise', 'arguments',
    // deployment/answer_key.gs is pasted as its own file and is not in the repo;
    // answerKeyIndex() already refuses to score when it is absent.
    'lookupCorrectIndex']);
  const addPattern = (pattern, into) => {
    if (!pattern) return;
    if (pattern.type === 'Identifier') { into.add(pattern.name); return; }
    if (pattern.type === 'ObjectPattern') for (const p of pattern.properties) addPattern(p.value || p.argument, into);
    if (pattern.type === 'ArrayPattern') for (const e of pattern.elements) addPattern(e, into);
    if (pattern.type === 'AssignmentPattern') addPattern(pattern.left, into);
    if (pattern.type === 'RestElement') addPattern(pattern.argument, into);
  };
  for (const name of [MONOLITH, EXAM_FILE, REPORTS_FILE]) {
    const ast = acorn.parse(fs.readFileSync(path.join(ROOT, name), 'utf8'),
      { ecmaVersion: 2022, sourceType: 'script', locations: true });
    const declared = new Set(), used = new Map();
    (function walk(node, parent) {
      if (Array.isArray(node)) { for (const item of node) walk(item, parent); return; }
      if (!node || typeof node !== 'object' || !node.type) return;
      if (node.type === 'FunctionDeclaration' || node.type === 'FunctionExpression') {
        if (node.id) declared.add(node.id.name);
        for (const param of node.params) addPattern(param, declared);
      }
      if (node.type === 'VariableDeclarator') addPattern(node.id, declared);
      if (node.type === 'CatchClause') addPattern(node.param, declared);
      if (node.type === 'Identifier') {
        const isKey = parent && ((parent.type === 'MemberExpression' && parent.property === node && !parent.computed) ||
          (parent.type === 'Property' && parent.key === node && !parent.computed) ||
          ((parent.type === 'FunctionDeclaration' || parent.type === 'FunctionExpression') && parent.id === node));
        if (!isKey && !used.has(node.name)) used.set(node.name, node.loc.start.line);
        return;
      }
      for (const key in node) { if (key !== 'loc') walk(node[key], node); }
    })(ast.body, null);
    const free = [...used].filter(([id]) => !declared.has(id) && !PROVIDED.has(id))
      .map(([id, line]) => id + ' (line ' + line + ')');
    assert.deepEqual(free, [], name + ' calls something it does not carry — check server/BUILD_TARGETS.json');
  }
});

// ---- 2. each half serves its own actions and refuses the other's -----------

for (const [file, deployment, mine, theirs] of [
  [EXAM_FILE, 'exam', 'examinerDashboard', 'commanderDashboard'],
  [REPORTS_FILE, 'reports', 'startPractice', 'examinerDashboard']
]) {
  test(file + ' answers only for its own half', () => {
    const e = env(file);
    assert.equal(e.ctx.API_DEPLOYMENT, deployment);
    const health = dispatch(e, 'health');
    assert.equal(health.status, 'ok');
    assert.equal(health.deployment, deployment);
    assert.equal(health.indexIds, 1700, 'the packed index unpacks in this deployment too');
    assert.equal(dispatch(e, 'getOfficeNumber').status, 'ok', 'getOfficeNumber is the other `both` action');

    // A foreign action is refused WITHOUT a credential: the target check runs
    // before auth, so an examiner whose page has not reloaded yet is told to
    // reload rather than sent to re-login.
    const foreign = dispatch(e, theirs);
    assert.equal(foreign.status, 'error');
    assert.equal(foreign.code, 'wrong_deployment');
    assert.match(foreign.message, /שרת אחר/);

    // And an action of this half really is registered here — not "Unknown
    // action" and not the namedHandler fallback "Action not available".
    const own = dispatch(e, mine);
    assert.notEqual(own.code, 'wrong_deployment');
    assert.ok(!/Unknown action|Action not available/.test(String(own.message || '')), mine + ': ' + own.message);
  });

  test(file + ' registers every action it is supposed to serve', () => {
    const e = env(file);
    e.ctx.ensureLegacyActions();
    const targets = e.ctx.ACTION_TARGETS;
    const registry = e.ctx.apiRegistry();
    // The legacy rows name their handler as a string and resolve it at call
    // time, so "registered" is not enough — the function has to be in the file.
    const legacyHandler = Object.fromEntries(e.ctx.legacyActionTable().map(row => [row[0], row[3]]));
    const missing = [], unresolvable = [], intruders = [];
    for (const action of Object.keys(targets)) {
      const target = targets[action];
      const belongs = target === 'both' || target === deployment;
      if (!belongs) {
        // A reports action in the exam file is allowed to be registered (the
        // legacy table is in 30_api.js, which is `both`) — what must NOT happen
        // is that it runs, and test above proves it does not.
        if (registry[action] && legacyHandler[action] && typeof e.ctx[legacyHandler[action]] === 'function') {
          intruders.push(action + ' -> ' + legacyHandler[action]);
        }
        continue;
      }
      if (!registry[action]) { missing.push(action); continue; }
      const fnName = legacyHandler[action];
      if (fnName && typeof e.ctx[fnName] !== 'function') unresolvable.push(action + ' -> ' + fnName);
    }
    assert.deepEqual(missing, [], 'not registered in ' + file);
    assert.deepEqual(unresolvable, [], 'registered but the handler is not in ' + file);
    assert.deepEqual(intruders, [], 'the handler of a foreign action is still in ' + file);
  });
}

test('the reports deployment keeps the nightly jobs and the exam one removes them', () => {
  const reports = env(REPORTS_FILE);
  assert.equal(typeof reports.ctx.installNightlyJobs, 'function');
  assert.equal(typeof reports.ctx.archiveSheets, 'function');
  assert.equal(typeof reports.ctx.uninstallNightlyJobs, 'function', 'uninstall is `both`');

  const exam = env(EXAM_FILE);
  assert.equal(typeof exam.ctx.installNightlyJobs, 'undefined', 'installNightlyJobs must NOT be in the exam file');
  for (const fn of ['archiveSheets', 'rebuildAtRiskCache', 'warmupQuestionCaches']) {
    exam.ctx.ScriptApp.newTrigger(fn).timeBased().everyHours(1).create();
  }
  exam.ctx.ScriptApp.newTrigger('doSomethingElse').timeBased().everyHours(6).create();
  const msg = exam.ctx.uninstallNightlyJobs();
  assert.deepEqual(exam.triggers.map(t => t.getHandlerFunction()), ['doSomethingElse']);
  assert.match(msg, /removed 3 trigger\(s\)/);
  assert.equal(exam.ctx.uninstallNightlyJobs(), 'uninstallNightlyJobs: removed 0 trigger(s) []; 1 other trigger(s) left untouched');
});

// ---- 3. the packed index is the committed index -----------------------------

test('QUESTION_INDEX_PACKED unpacks to exactly deployment/question_index.json', () => {
  const source = JSON.parse(fs.readFileSync(path.join(ROOT, 'deployment', 'question_index.json'), 'utf8'));
  for (const file of [MONOLITH, EXAM_FILE, REPORTS_FILE]) {
    const e = env(file);
    // Back into this realm: the object is built inside the vm and carries its
    // prototypes, which deepStrictEqual compares.
    const unpacked = JSON.parse(JSON.stringify(e.ctx.questionIndex()));
    assert.deepEqual(Object.keys(unpacked).sort(), Object.keys(source).sort(), file);
    assert.deepEqual(unpacked, source, file + ': the unpacked index is not the committed one');
    assert.equal(e.ctx.questionIndexCount(), Object.keys(source).length, file);
    // Memoised, like apiRegistry: a second call is the same object.
    assert.equal(e.ctx.questionIndex(), e.ctx.questionIndex(), file + ': questionIndex() is not memoised');
  }
});

test('the packed string is a fraction of the object literal it replaced', () => {
  const text = fs.readFileSync(path.join(ROOT, MONOLITH), 'utf8');
  const packed = text.match(/var QUESTION_INDEX_PACKED = \{[\s\S]*?\n\};/);
  assert.ok(packed, 'the packed index is not in the built file');
  const rawJson = fs.statSync(path.join(ROOT, 'deployment', 'question_index.json')).size;
  assert.ok(packed[0].length < rawJson / 4,
    'the packed index is ' + packed[0].length + ' chars against ' + rawJson + ' of JSON — the point was the compile time');
});

test('the index still drives a real draw in the exam deployment', () => {
  const e = env(EXAM_FILE, { sources: [path.join('deployment', 'answer_key.gs')] });
  const drawn = JSON.parse(JSON.stringify(e.ctx.drawExamIds('B', 'he')));
  assert.equal(drawn.length, 30);
  const byTopic = {};
  for (const q of drawn) byTopic[q.topic] = (byTopic[q.topic] || 0) + 1;
  assert.deepEqual(byTopic, { 'בטיחות': 7, 'הכרת הרכב': 7, 'חוק': 7, 'תמרורים': 9 });
});

// ---- 4. a standalone project has no active spreadsheet ---------------------

test('getSpreadsheet falls back to EXAM_SPREADSHEET_ID, and says so when it is not set', () => {
  const bound = env(REPORTS_FILE);
  assert.equal(bound.ctx.getSpreadsheet(), bound.active, 'a bound project still uses the active spreadsheet');

  const standalone = createEnv({ serverFile: REPORTS_FILE, standalone: true,
    properties: { EXAM_SPREADSHEET_ID: 'active' },
    sheets: { 'אתרים': [['שם אתר'], ['בדיקת נתונים']] } });
  assert.equal(standalone.ctx.getSpreadsheet(), standalone.active);
  assert.equal(standalone.json(standalone.ctx.dispatchApiAction('GET', 'health', { deep: '1' })).sheetError, '',
    'health&deep=1 reads a cell of the opened document');

  const unconfigured = createEnv({ serverFile: REPORTS_FILE, standalone: true });
  assert.throws(() => unconfigured.ctx.getSpreadsheet(),
    /EXAM_SPREADSHEET_ID is not set — standalone deployment needs the exam spreadsheet id/);
});

// ---- 5. the server's table and the client's list are one list --------------

test("ACTION_TARGETS's reports set is exactly ExamTransport.REPORTS_ACTIONS", () => {
  const e = env(MONOLITH);
  const targets = plain(e.ctx.ACTION_TARGETS);
  const serverReports = Object.keys(targets).filter(a => targets[a] === 'reports').sort();
  const clientReports = plain(transportModule().REPORTS_ACTIONS).sort();
  assert.equal(clientReports.length, 22, 'DESIGN §13.3 names 22 cold actions');
  assert.deepEqual(serverReports, clientReports,
    'the pages route by the client list and the server refuses by the table — they must be the same list');
});

test('every action has a target and every target names a real action', () => {
  const e = env(MONOLITH);
  // ensureLegacyActions throws when an action has no target, so reaching here
  // is already half the assertion; the reverse direction is checked explicitly.
  e.ctx.ensureLegacyActions();
  const targets = plain(e.ctx.ACTION_TARGETS);
  const registered = plain(e.ctx.apiActionNames()).sort();
  const targeted = Object.keys(targets).sort();
  assert.deepEqual(targeted, registered,
    'the monolith registers every action, so the two lists are the same list');
  assert.equal(registered.length, 67, 'r33 added bankRelay and reportGateway');
  for (const action of registered) {
    assert.ok(['both', 'exam', 'reports'].includes(targets[action]), action + ' has target ' + targets[action]);
  }
  assert.deepEqual(targeted.filter(a => targets[a] === 'both'), ['getOfficeNumber', 'health']);
});

// r33 (24/09/2026): the Google fallback of a phone that cannot reach the Worker
// runs on the EXAM project — the examinee page only knows that URL — and the
// server's first UrlFetchApp calls stay out of the reports file: the exam
// project's manifest (appsscript.json, tracked) already declares
// script.external_request, while the reports project (created 22/09) has a
// manifest this repo does not track — a UrlFetchApp there could make Google
// demand a new authorisation in the middle of a paste.
test('r33: bankRelay and reportGateway are exam actions, and UrlFetchApp stays in the exam half', () => {
  const e = env(MONOLITH);
  const targets = plain(e.ctx.ACTION_TARGETS);
  for (const action of ['bankRelay', 'reportGateway', 'checkApproval', 'getExamStatus']) {
    assert.equal(targets[action], 'exam', action);
  }
  const clientReports = plain(transportModule().REPORTS_ACTIONS);
  assert.ok(!clientReports.includes('bankRelay') && !clientReports.includes('reportGateway'),
    'the page sends them to the exam URL');
  const exam = env(EXAM_FILE), reports = env(REPORTS_FILE);
  for (const action of ['bankRelay', 'reportGateway']) {
    const served = dispatch(exam, action, 'POST');
    assert.notEqual(served.code, 'wrong_deployment', action);
    assert.equal(served.message, 'חסרים פרטי נבחן', action + ' reached its examinee-token check');
    assert.equal(dispatch(reports, action, 'POST').code, 'wrong_deployment', action);
  }
  // The two polls are real answers again in the exam file, not client_outdated.
  for (const action of ['checkApproval', 'getExamStatus']) {
    assert.notEqual(dispatch(exam, action, 'GET').code, 'client_outdated', action);
  }
  assert.equal(typeof exam.ctx.testGatewayReachability, 'function', 'the operator runs it in the exam project');
  assert.equal(typeof reports.ctx.testGatewayReachability, 'undefined');
  assert.ok(!fs.readFileSync(path.join(ROOT, REPORTS_FILE), 'utf8').includes('UrlFetchApp'),
    'no UrlFetchApp in the reports file');
  assert.ok(fs.readFileSync(path.join(ROOT, EXAM_FILE), 'utf8').includes('UrlFetchApp'));
});

test('an action without a target is refused at registration, not at 06:30', () => {
  const e = env(MONOLITH);
  e.ctx.defineAction('somethingNobodyRouted', { methods: ['GET'], auth: 'none', handler: () => null });
  e.ctx.ensureLegacyActions._done = false;   // force the registration pass again
  assert.throws(() => e.ctx.ensureLegacyActions(),
    /ACTION_TARGETS has no entry for: somethingNobodyRouted/);
});

test('the monolith serves everything — API_DEPLOYMENT "all" never refuses', () => {
  const e = env(MONOLITH);
  assert.equal(e.ctx.API_DEPLOYMENT, 'all');
  for (const action of Object.keys(e.ctx.ACTION_TARGETS)) {
    assert.notEqual(dispatch(e, action, 'GET').code, 'wrong_deployment', action);
    assert.notEqual(dispatch(e, action, 'POST').code, 'wrong_deployment', action);
  }
  assert.equal(dispatch(e, 'health').deployment, 'all');
});

// ---- 6. every page talks to a server that serves it ------------------------

// The URL each page routes an action to (DESIGN §13.3). examiner.html is the
// only page that routes per action — it holds both URLs — so its actions may be
// anything. The others send everything to one URL, so every action they send
// must be served there. (REPORTS_API_URL itself is package E's; this asserts the
// actions, which is what breaks if the routing is wrong.)
const PAGE_ROUTING = {
  'examiner.html': 'either',
  'examinee.html': 'exam',
  'find_image.html': 'exam',
  'teacher.html': 'reports',
  'student.html': 'reports',
  'exam.html': 'reports'
};

function actionsOf(page) {
  const text = fs.readFileSync(path.join(ROOT, page), 'utf8');
  const found = new Set();
  const re = /\baction\s*:\s*'([A-Za-z][A-Za-z0-9_]*)'/g;
  for (let m = re.exec(text); m; m = re.exec(text)) found.add(m[1]);
  return [...found].sort();
}

test('every action a page sends is known to the server', () => {
  const e = env(MONOLITH);
  e.ctx.ensureLegacyActions();
  const known = new Set(e.ctx.apiActionNames());
  for (const page of Object.keys(PAGE_ROUTING)) {
    const actions = actionsOf(page);
    assert.ok(actions.length, page + ' sends no action at all — the scan is broken');
    for (const action of actions) assert.ok(known.has(action), page + ' sends unknown action ' + action);
  }
});

test('every action a page sends is served by the URL that page routes to', () => {
  const e = env(MONOLITH);
  const targets = e.ctx.ACTION_TARGETS;
  const wrong = [];
  for (const [page, route] of Object.entries(PAGE_ROUTING)) {
    if (route === 'either') continue;
    for (const action of actionsOf(page)) {
      if (targets[action] !== route && targets[action] !== 'both') {
        wrong.push(page + ' sends ' + action + ' (' + targets[action] + ') to the ' + route + ' url');
      }
    }
  }
  assert.deepEqual(wrong, []);
});

test('examiner.html is the page that needs both urls', () => {
  const e = env(MONOLITH);
  const targets = e.ctx.ACTION_TARGETS;
  const sent = actionsOf('examiner.html');
  assert.ok(sent.some(a => targets[a] === 'exam'), 'examiner.html must keep hot actions on the exam url');
  assert.ok(sent.some(a => targets[a] === 'reports'), 'examiner.html must route its reports to the cold url');
  // These four are why: they are the examiner's own reports, and they are the
  // slowest handlers in the system.
  for (const action of ['commanderDashboard', 'centerManagerReport', 'siteCombinedReport', 'examinerForecast']) {
    assert.ok(sent.includes(action), 'examiner.html no longer sends ' + action);
    assert.equal(targets[action], 'reports');
  }
});

// ---- 7. the snapshot and the writer that used to leave it stale ------------

const PENDING_HEADER = ['קוד סשן', 'ת.ז.', 'שם', 'טלפון', 'זמן הרשמה', 'סטטוס', 'שפה', 'אוכלוסיה', 'דרגה', 'שמע',
  'הארכת זמן', 'התחלת מבחן', 'טוקן נבחן', 'ספירת DQ', 'מסך נוסף', 'ספירת אזהרות', 'אזהרה אחרונה', 'אתר', 'סיים במכשיר'];
const SESSION = 'ABC12345';
function pendingRow(id, status, over) {
  const row = new Array(19).fill('');
  row[0] = SESSION; row[1] = id; row[2] = 'נבחן ' + id; row[3] = '0501234567';
  row[4] = '2026-09-22T06:00:00Z'; row[5] = status; row[6] = 'he'; row[8] = 'B'; row[9] = 'off'; row[12] = 'tok-' + id;
  return Object.assign(row, over || {});
}
function pendingEnv(rows) {
  return createEnv({
    serverFile: EXAM_FILE,
    sheets: { 'ממתינים': [PENDING_HEADER, ...rows], 'הארכות זמן': [['תאריך', 'קוד סשן', 'ת.ז.', 'שם', 'דקות', 'סיבה', 'בוחן']] },
    properties: { GATEWAY_KEY: 'k', GATEWAY_URL: 'https://gw.test' }
  });
}
const snapshotOf = e => e.json(e.ctx.dispatchApiAction('GET', 'sessionSnapshot', { sessionCode: SESSION, gatewayKey: 'k' }));

test('sessionSnapshot carries warn, fin, ext and dq for the Worker fingerprint', () => {
  const e = pendingEnv([
    pendingRow('900000001', 'in_exam', { 13: 2, 14: 'כן', 15: 3, 18: '2026-09-22T07:10:00Z' }),
    pendingRow('900000002', 'waiting')
  ]);
  const snap = snapshotOf(e);
  assert.equal(snap.status, 'ok');
  assert.deepEqual(
    { warn: snap.rows[0].warn, fin: snap.rows[0].fin, ext: snap.rows[0].ext, dq: snap.rows[0].dq },
    { warn: 3, fin: 1, ext: 1, dq: 2 });
  // The quiet row must read 0, not '' — the Worker joins these into a string.
  assert.deepEqual(
    { warn: snap.rows[1].warn, fin: snap.rows[1].fin, ext: snap.rows[1].ext, dq: snap.rows[1].dq },
    { warn: 0, fin: 0, ext: 0, dq: 0 });
  for (const row of snap.rows) {
    for (const field of ['warn', 'fin', 'ext', 'dq']) assert.equal(typeof row[field], 'number', field);
  }
});

const snapshotKey = e => e.ctx.pendingSnapshotKey(SESSION);

test('markPendingCompleted flushes and drops the session snapshot', () => {
  const e = pendingEnv([pendingRow('900000001', 'in_exam')]);
  // Warm the snapshot the pollers (and therefore the Worker) read.
  assert.equal(snapshotOf(e).rows[0].status, 'in_exam');
  assert.ok(e.cache.get(snapshotKey(e)), 'the per-session snapshot was not cached — this proves nothing without it');

  const flushesBefore = e.flushes.count;
  e.ctx.markPendingCompleted(SESSION, '900000001', null);
  assert.ok(e.flushes.count > flushesBefore, 'the write must be flushed BEFORE the snapshot is dropped');
  assert.equal(e.cache.get(snapshotKey(e)), null,
    'the stale snapshot survived the write — that is the r30 behaviour DESIGN §13.6 fixes');
  // What the Worker would have served for up to PENDING_SNAPSHOT_SEC.
  assert.equal(snapshotOf(e).rows[0].status, 'completed');
});

test('markPendingCompleted does not drop a snapshot when it wrote nothing', () => {
  const e = pendingEnv([pendingRow('900000001', 'completed')]);
  snapshotOf(e);
  assert.ok(e.cache.get(snapshotKey(e)));
  const flushesBefore = e.flushes.count;
  e.ctx.markPendingCompleted(SESSION, '900000001', null);
  assert.equal(e.flushes.count, flushesBefore, 'nothing was written, so nothing needed flushing');
  assert.ok(e.cache.get(snapshotKey(e)) !== null, 'a no-op must not cost the next poller an upstream read');
});
