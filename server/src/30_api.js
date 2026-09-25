// Which of the three generated files this is — 'all' (the monolith), 'exam' (the
// hot project every page already points at) or 'reports' (the standalone cold
// project). tools/build_server.js replaces the marker line below per output;
// nothing else in the sources may assign it. health reports it, so a paste can
// be verified from outside without reading a single private cell.
// @@API_DEPLOYMENT@@

// Public build marker: identifies the deployed API without reading private data.
var THEORY_API_BUILD = '2026-09-27-r35.2';
// When the current request entered the script — health&deep=1 reports the whole
// request against it, so a watchdog can separate our time from Google's.
var API_STARTED_AT = 0;

// Every action name known to the deployment, for the timing log (which must
// never echo an arbitrary string a caller sent) and for feature detection.
function apiActionList() {
  ensureLegacyActions();
  return apiActionNames();
}

function logTheoryApiTiming(phase, method, action, startedAt) {
  // Never log request parameters, IDs, credentials, answers or arbitrary action text.
  // A start without an end can identify a runtime timeout in the execution log.
  try {
    Logger.log('[API] ' + JSON.stringify({ build: THEORY_API_BUILD, phase: phase,
      method: method, action: apiActionList().indexOf(action) >= 0 ? action : 'unknown',
      elapsedMs: Math.max(0, Date.now() - startedAt) }));
  } catch (logErr) { /* diagnostics must never break an exam */ }
}

function theoryRetryableErrorResponse(err) {
  if (!err || err.retryable !== true) return null;
  return jsonResponse({ status: 'error', code: err.code || 'retry_later', retryable: true,
    waitSec: Math.max(1, Math.min(30, Number(err.waitSec) || 3)),
    message: err.userMessage || 'המערכת עמוסה כעת. אפשר לנסות שוב בעוד מספר שניות.' });
}

// ========== Dispatch ==========
// doGet/doPost were a 300-line switch plus three hand-maintained lists of which
// action needs which token. Now every action is a registry row (name, methods,
// auth, handler) and the two entry points do the same four things: parse, check
// the method, check the auth rule, call the handler.

function dispatchApiAction(method, action, p) {
  ensureLegacyActions();
  // The target check runs BEFORE the registry lookup, and before any auth, on
  // purpose: in a split deployment the handler of a foreign action is not in
  // this file at all, so the honest answer is "wrong server", not "unknown
  // action" (which a page would report as a bug) and not "token invalid" (which
  // would send an examiner to re-login for nothing). No credential is read to
  // produce it, so it leaks nothing an anonymous caller could not already guess
  // from health.
  var target = ACTION_TARGETS[action];
  if (target && target !== 'both' && API_DEPLOYMENT !== 'all' && target !== API_DEPLOYMENT) {
    return jsonResponse({ status: 'error', code: 'wrong_deployment',
      message: 'הפעולה שייכת לשרת אחר — יש לרענן את הדף' });
  }
  // r35.1: practice moved to the new system (20_auth.js PRACTICE_MOVED). Only
  // the practice-flow actions read the property, and before any auth. r35.2:
  // exam.html's startPractice (the audio exam) is exempt.
  var movedErr = practiceMovedRefusal(action, p);
  if (movedErr) return movedErr;
  var spec = apiRegistry()[action];
  if (!spec) return jsonResponse({ status: 'error', message: 'Unknown action: ' + action });
  if (spec.methods.indexOf(method) === -1) {
    return jsonResponse({ status: 'error',
      message: method === 'GET' ? 'פעולה זו דורשת POST' : 'פעולה זו דורשת GET' });
  }
  var authErr = requireActionAuth(spec.auth, p);
  if (authErr) return authErr;
  if (spec.rateLimit) {
    var rlErr = requireRateLimit(action, spec.rateLimit.id(p), spec.rateLimit.max, spec.rateLimit.windowSec || 60);
    if (rlErr) return rlErr;
  }
  return spec.handler(p);
}

function requireActionAuth(auth, p) {
  if (auth === 'examiner') return requireToken(p);
  // r35 (KNOWN_ISSUES #43): a valid examiner token AND the right to read THIS
  // session's board — its own examiner or a 'מפקד' (20_auth.js mayViewSession).
  if (auth === 'examinerSession') return requireToken(p) || requireSessionViewer(p);
  // requireTeacherToken lives in 90_teacher.js, which the exam deployment does
  // not carry. Every teacher action is a `reports` action, so the target check
  // above already refused it and this line is unreachable there — but resolving
  // the name instead of referencing it means the day somebody gives a
  // teacher-authenticated action an `exam` target, the exam deployment REFUSES
  // it rather than throwing ReferenceError out of the middle of the router.
  if (auth === 'teacher') {
    var teacherCheck = globalFunction('requireTeacherToken');
    return teacherCheck ? teacherCheck(p) : jsonResponse({ status: 'error', code: 'wrong_deployment',
      message: 'הפעולה שייכת לשרת אחר — יש לרענן את הדף' });
  }
  if (auth === 'examinee') return requireExamineeToken(p);
  if (auth === 'gateway') return requireGatewayKey(p);
  return null;   // 'none' — either public, or the handler enforces its own rule
}

// The session-poll Worker is the only caller that reads a whole session's rows
// in one request; it authenticates with a shared secret kept in ScriptProperties
// (never in the client), so an examinee token is not involved.
function requireGatewayKey(p) {
  var expected = gatewayKey();   // the same property the bank grants are signed with
  if (!expected || String(p.gatewayKey || '') !== expected) {
    return jsonResponse({ status: 'error', code: 'gateway_denied', message: 'gateway key invalid' });
  }
  return null;
}

// ---- Which deployment serves which action (DESIGN §13.3) --------------------
// ONE table for every action the system has, whether it is declared here, in the
// legacy table below or with defineAction() inside a module. It is the contract
// the client half of the split is checked against: the 22 'reports' names must
// be exactly ExamTransport.REPORTS_ACTIONS (tests/server_split.test.cjs), because
// a page routes by that list and the server refuses by this one — if they ever
// drift, a page sends to a server that answers wrong_deployment forever.
// 'both' is only what EVERY page may need whatever it is doing: health (the
// deploy check and the transport's own probe) and getOfficeNumber (a public
// display string). Everything else belongs to exactly one project, and the
// default for a new action is 'exam' — the hot file — only if it really runs
// during an exam; a report, a teacher screen or practice is 'reports'.
var ACTION_TARGETS = {
  // -- served by both deployments --
  health: 'both', getOfficeNumber: 'both',
  // -- the cold project: reports, commanders, teachers, students, practice --
  startPractice: 'reports', submitPracticeResult: 'reports', loadStudentProgress: 'reports',
  saveStudentProgress: 'reports', studentJoinClass: 'reports',
  teacherLogin: 'reports', teacherVerifyLogin: 'reports', teacherDashboard: 'reports',
  teacherCreateClass: 'reports', teacherCloseClass: 'reports', teacherDeleteClass: 'reports',
  teacherRemoveStudent: 'reports', teacherGetClasses: 'reports', teacherClassDetails: 'reports',
  teacherExportData: 'reports', teacherCommanderDashboard: 'reports', teacherAtRiskList: 'reports',
  adminDashboard: 'reports', commanderDashboard: 'reports', centerManagerReport: 'reports',
  siteCombinedReport: 'reports', examinerForecast: 'reports',
  // -- the hot project: everything an exam morning touches --
  login: 'exam', verifyLogin: 'exam', getSites: 'exam', listSessions: 'exam',
  listAllSessions: 'exam', listActiveExaminers: 'exam', createSession: 'exam',
  updateSession: 'exam', closeSession: 'exam', getSessionInfo: 'exam',
  registerExaminee: 'exam', cancelRegistration: 'exam', approveExaminee: 'exam',
  rejectExaminee: 'exam', examinerDashboard: 'exam', resetExaminee: 'exam',
  forceComplete: 'exam', markSent: 'exam', correctToPass: 'exam',
  commanderCorrectResult: 'exam', submitManualResult: 'exam', correctExamineeMeta: 'exam',
  checkApproval: 'exam', getExamStatus: 'exam', addExamTime: 'exam', markFinished: 'exam',
  reportWarning: 'exam', disqualify: 'exam', cancelDisqualify: 'exam',
  overturnDQ: 'exam', confirmDQ: 'exam',
  startExam: 'exam', markExamStarted: 'exam', getExamQuestions: 'exam',
  registerExamQuestions: 'exam', submitResult: 'exam', submitFailOnClose: 'exam',
  cancelFailOnClose: 'exam', getResultUploadToken: 'exam',
  sessionSnapshot: 'exam', bankGrant: 'exam',
  // r33 (24/09/2026): the Google fallback of a phone that cannot reach the Worker
  bankRelay: 'exam', reportGateway: 'exam'
};

// ---- Where the action rows live --------------------------------------------
// A defineAction row is evaluated when the file LOADS, so `handler: handleX`
// must be a function this deployment actually contains. The rows for the exam
// handlers therefore sit in 60_exam.js next to them (22/09/2026, DESIGN §13.3)
// — this module is `both` and would otherwise crash the reports project on
// load with "handleStartExam is not defined". The legacy table below is safe
// the same way for a different reason: it names its handlers as STRINGS and
// resolves them at call time.

// ---- Every other action, with today's method and auth rule ------------------
// S2 will move these rows next to their handlers; until then this table is the
// single declaration of them, and it reproduces exactly what the old doGet/doPost
// enforced (its examinerActions / teacherActions / postOnlyActions lists), so the
// dispatcher is a refactor and not a policy change. Handlers are named, not
// referenced, so a module that replaces one is picked up at call time.
//   [name, methods, auth, handler function name]
function legacyActionTable() {
  return [
    // -- examiner token (the old examinerActions list) --
    ['getSites', 'GET', 'examiner', 'handleGetSites'],
    ['listSessions', 'GET', 'examiner', 'handleListSessions'],
    ['listAllSessions', 'GET', 'examiner', 'handleListAllSessions'],
    ['createSession', 'GET', 'examiner', 'handleCreateSession'],
    ['updateSession', 'GET', 'examiner', 'handleUpdateSession'],
    ['closeSession', 'GET', 'examiner', 'handleCloseSession'],
    ['approveExaminee', 'GET', 'examiner', 'handleApproveExaminee'],
    ['rejectExaminee', 'GET', 'examiner', 'handleRejectExaminee'],
    // r35: the board of ONE session — token + owner or מפקד (review 09 F-06)
    ['examinerDashboard', 'GET', 'examinerSession', 'handleExaminerDashboard'],
    ['resetExaminee', 'GET', 'examiner', 'handleResetExaminee'],
    ['correctToPass', 'GET', 'examiner', 'handleCorrectToPass'],
    ['overturnDQ', 'GET', 'examiner', 'handleOverturnDQ'],
    ['confirmDQ', 'GET', 'examiner', 'handleConfirmDQ'],
    ['forceComplete', 'GET', 'examiner', 'handleForceComplete'],
    ['markSent', 'GET', 'examiner', 'handleMarkSent'],
    ['commanderDashboard', 'GET', 'examiner', 'handleCommanderDashboard'],
    ['centerManagerReport', 'GET', 'examiner', 'handleCenterManagerReport'],
    ['examinerForecast', 'GET', 'examiner', 'handleExaminerForecast'],
    // POST-only in the old router; the handlers verify the examiner themselves too
    ['commanderCorrectResult', 'POST', 'examiner', 'handleCommanderCorrectResult'],
    ['submitManualResult', 'POST', 'examiner', 'handleSubmitManualResult'],
    ['correctExamineeMeta', 'GET,POST', 'examiner', 'handleCorrectExamineeMeta'],
    // -- teacher token (the old teacherActions list) --
    ['teacherDashboard', 'GET', 'teacher', 'handleTeacherDashboard'],
    ['teacherCreateClass', 'GET', 'teacher', 'handleTeacherCreateClass'],
    ['teacherCloseClass', 'GET', 'teacher', 'handleTeacherCloseClass'],
    ['teacherDeleteClass', 'GET', 'teacher', 'handleTeacherDeleteClass'],
    ['teacherRemoveStudent', 'GET', 'teacher', 'handleTeacherRemoveStudent'],
    ['teacherGetClasses', 'GET', 'teacher', 'handleTeacherGetClasses'],
    ['teacherClassDetails', 'GET', 'teacher', 'handleTeacherClassDetails'],
    ['teacherExportData', 'GET', 'teacher', 'handleTeacherExportData'],
    ['teacherCommanderDashboard', 'GET', 'teacher', 'handleTeacherCommanderDashboard'],
    ['teacherAtRiskList', 'GET', 'teacher', 'handleTeacherAtRiskList'],
    ['adminDashboard', 'GET', 'teacher', 'handleAdminDashboard'],
    // -- public / handler-enforced auth --
    ['login', 'POST', 'none', 'handleLogin'],
    ['verifyLogin', 'GET', 'none', 'handleVerifyLogin'],
    ['teacherLogin', 'POST', 'none', 'handleTeacherLogin'],
    ['teacherVerifyLogin', 'GET', 'none', 'handleTeacherVerifyLogin'],
    ['getOfficeNumber', 'GET', 'none', 'handleGetOfficeNumber'],
    ['listActiveExaminers', 'GET', 'none', 'handleListActiveExaminers'],
    ['siteCombinedReport', 'GET', 'none', 'handleSiteCombinedReport'],
    ['getSessionInfo', 'GET', 'none', 'handleGetSessionInfo'],
    ['registerExaminee', 'GET', 'none', 'handleRegisterExaminee'],
    // r35: the handler requires the examinee token of the row it cancels
    ['cancelRegistration', 'GET', 'none', 'handleCancelRegistration'],
    // Retired 21/09/2026 evening (the page polled ONLY through the session
    // gateway) and SERVED AGAIN in r33 (24/09/2026, KNOWN_ISSUES #38): on 23/09
    // many phones reached Google but never the Worker, so a phone that fails
    // the Worker now polls these two directly — paced by the page (12 s / 20 s)
    // and only on the phones that failed; every other phone stays on the
    // Worker. Same handlers, same answers: they are also the reference the
    // Worker is tested against (tests/contracts.test.cjs). Auth is theirs: a
    // per-examinee rate limit and the token-mismatch rule inside the handler.
    ['checkApproval', 'GET', 'none', 'handleCheckApproval'],
    ['getExamStatus', 'GET', 'none', 'handleGetExamStatus'],
    ['addExamTime', 'GET', 'none', 'handleAddExamTime'],
    // GET,POST since r33.1 (24/09/2026): the page has sent this ping as a BEACON
    // (sendBeacon = POST, like disqualify / reportWarning below) since r30, and a
    // GET-only row refused every one of them — the examiner never saw
    // 'סיים — מסנכרן תוצאה' (found 23/09 evening).
    ['markFinished', 'GET,POST', 'none', 'handleMarkFinished'],
    // 'disqualify' is deliberately not examiner-gated: the examinee client sends
    // it too, and handleDisqualify accepts either an examiner token or an active
    // pending row of that examinee.
    ['disqualify', 'GET,POST', 'none', 'handleDisqualify'],
    ['reportWarning', 'GET,POST', 'none', 'handleReportWarning'],
    ['studentJoinClass', 'GET', 'none', 'handleStudentJoinClass'],
    ['submitPracticeResult', 'GET,POST', 'none', 'handleSubmitPracticeResult'],
    ['loadStudentProgress', 'GET', 'none', 'handleLoadStudentProgress'],
    ['saveStudentProgress', 'POST', 'none', 'handleSaveStudentProgress']
  ];
}

// Registered on first dispatch rather than at load time, so a module that
// declared the same action with defineAction() wins and a DIFFERENT declaration
// of the same name — a merge accident between two packages — throws instead of
// silently taking effect.
function ensureLegacyActions() {
  if (ensureLegacyActions._done) return;
  ensureLegacyActions._done = true;
  var table = legacyActionTable();
  for (var i = 0; i < table.length; i++) {
    var name = table[i][0], methods = table[i][1].split(','), auth = table[i][2], fnName = table[i][3];
    var existing = apiRegistry()[name];
    if (!existing) {
      defineAction(name, { methods: methods, auth: auth, handler: namedHandler(fnName) });
      continue;
    }
    if (existing.auth !== auth || existing.methods.join(',') !== methods.join(',')) {
      throw new Error('conflicting defineAction for ' + name + ': ' + existing.methods.join('/') + '/' + existing.auth +
        ' vs ' + methods.join('/') + '/' + auth);
    }
  }
  assertActionTargets();
}

// An action with no row in ACTION_TARGETS is not a small omission: in the split
// it would be served by whichever of the two projects happens to hold its
// handler and answer "Unknown action" in the other, with nothing in either log
// saying why. Loud here, at registration, exactly like a conflicting
// defineAction — and the test gate (node tools/test.js) runs before every paste,
// so this can only fire on a tree that was never built.
function assertActionTargets() {
  var names = apiActionNames(), missing = [];
  for (var i = 0; i < names.length; i++) {
    if (!ACTION_TARGETS[names[i]]) missing.push(names[i]);
  }
  if (missing.length) {
    throw new Error('ACTION_TARGETS has no entry for: ' + missing.join(', ') +
      ' — every action must name the deployment that serves it (DESIGN §13.3)');
  }
}

// Resolved at call time: the handler may live in any module, and a test or a
// later module may replace it.
function namedHandler(fnName) {
  return function(p) {
    var fn = globalFunction(fnName);
    if (!fn) return jsonResponse({ status: 'error', message: 'Action not available: ' + fnName });
    return fn(p);
  };
}
function globalFunction(fnName) {
  var fn = null;
  try { fn = globalThis[fnName]; } catch (e) { fn = null; }
  return (typeof fn === 'function') ? fn : null;
}

function handleGetOfficeNumber() {
  // Public read of the office WA number — used by clients for display.
  return jsonResponse({ status: 'ok', officeWhatsApp: getOfficeWhatsAppNumber() });
}

// health&deep=1 (2026-09-19, review action 7): the plain health does no work at
// all, so it can only say "Google is slow". This one also reads a single cell of
// OUR document and reports that time separately, so a watchdog can tell "our
// document stalls" from "Google's front door stalls" every minute of an exam
// morning (tools/exam_watchdog.gs). indexIds is the deployed question index —
// the client compares it against the bank build the Worker served it.
function handleHealth(p) {
  // gateway: booleans only. The question texts are served by the Worker against
  // a signed grant, so "is it wired up" is the first thing a deploy check needs
  // — and neither the URL nor the key is ever printed by a public probe.
  // deployment: which of the three pasted files this is (DESIGN §13.3). It is
  // the whole verification of a split paste — "did the right file land in the
  // right project" — and it costs nothing to read.
  // movedSites (r35, KNOWN_ISSUES #44): how many sites MOVED_SITES lists, or
  // 'invalid' / 'error' — the check after editing that property. Never the names.
  // practiceMoved (r35.1, #45): true / false / 'invalid' / 'error' — the same
  // check for PRACTICE_MOVED (set in the reports project).
  var body = { status: 'ok', build: THEORY_API_BUILD, deployment: API_DEPLOYMENT,
    indexIds: questionIndexCount(),
    gateway: { url: Boolean(gatewayUrl()), key: Boolean(gatewayKey()) },
    movedSites: movedSitesHealth(), practiceMoved: practiceMovedHealth() };
  if (String(p.deep || '') !== '1') return jsonResponse(body);
  var deepT0 = Date.now(), sheetMs = -1, sheetError = '';
  try { getSheet('אתרים').getRange(1, 1).getValue(); sheetMs = Date.now() - deepT0; }
  catch (eDeep) { sheetError = String(eDeep && eDeep.message ? eDeep.message : eDeep).slice(0, 120); }
  body.deep = true;
  body.sheetMs = sheetMs;
  body.sheetError = sheetError;
  body.totalMs = Date.now() - API_STARTED_AT;
  return jsonResponse(body);
}
defineAction('health', { methods: ['GET'], auth: 'none', handler: handleHealth });

// ========== doGet — קריאות קריאה + פעולות קלות ==========

function doGet(e) {
  var apiStartedAt = API_STARTED_AT = Date.now();
  var action = '';
  diagBegin('GET');
  try {
    var p = (e && e.parameter) || {};
    action = p.action || '';
    if (DIAG_EXEC) { DIAG_EXEC.action = action; DIAG_EXEC.t0 = apiStartedAt; }
    logTheoryApiTiming('start', 'GET', action, apiStartedAt);

    // Soft origin check — log unauthorized origins (deterrent, bypassable but raises bar)
    var originErr = checkOrigin(p);
    if (originErr) return originErr;

    if (action === '') return jsonResponse({ status: 'ok', message: 'External Exam API is running' });
    return dispatchApiAction('GET', action, p);
  } catch (err) {
    return theoryRetryableErrorResponse(err) || jsonResponse({ status: 'error', message: err.toString() });
  } finally {
    diagFinish(action, apiStartedAt);
    logTheoryApiTiming('end', 'GET', action, apiStartedAt);
  }
}

// ========== doPost — שמירת תוצאות (נתונים גדולים) ==========

function doPost(e) {
  var apiStartedAt = API_STARTED_AT = Date.now();
  var action = '';
  diagBegin('POST');
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse({ status: 'error', message: 'No POST data received' });
    }
    var data = JSON.parse(e.postData.contents);
    action = data.action || '';
    if (DIAG_EXEC) { DIAG_EXEC.action = action; DIAG_EXEC.t0 = apiStartedAt; }
    logTheoryApiTiming('start', 'POST', action, apiStartedAt);

    // Soft origin check (deters casual scripts; bypassable by reading client source)
    var originErr = checkOrigin(data);
    if (originErr) return originErr;

    return dispatchApiAction('POST', action, data);
  } catch (err) {
    return theoryRetryableErrorResponse(err) || jsonResponse({ status: 'error', message: 'doPost error: ' + err.toString() });
  } finally {
    diagFinish(action, apiStartedAt);
    logTheoryApiTiming('end', 'POST', action, apiStartedAt);
  }
}
