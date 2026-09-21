// Public build marker: identifies the deployed API without reading private data.
var THEORY_API_BUILD = '2026-09-22-r30';
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
  if (auth === 'teacher') return requireTeacherToken(p);
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

// ---- Actions owned by this package -----------------------------------------
defineAction('startExam', { methods: ['POST'], auth: 'examinee', handler: handleStartExam,
  rateLimit: { max: 10, windowSec: 60, id: function(p) { return String(p.sessionCode || '') + '_' + normalizeId(p.idNumber); } } });
defineAction('startPractice', { methods: ['GET'], auth: 'none', handler: handleStartPractice });
defineAction('markExamStarted', { methods: ['GET'], auth: 'examinee', handler: handleMarkExamStartedNoop });
defineAction('getExamQuestions', { methods: ['GET'], auth: 'none', handler: handleClientOutdated });
defineAction('registerExamQuestions', { methods: ['POST'], auth: 'none', handler: handleClientOutdated });
defineAction('submitResult', { methods: ['POST'], auth: 'examinee', handler: handleSubmitResult });
defineAction('submitFailOnClose', { methods: ['POST'], auth: 'examinee', handler: handleSubmitFailOnClose });
defineAction('cancelFailOnClose', { methods: ['POST'], auth: 'examinee', handler: handleCancelFailOnClose });
defineAction('getResultUploadToken', { methods: ['GET'], auth: 'examiner', handler: handleGetResultUploadToken });

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
    ['examinerDashboard', 'GET', 'examiner', 'handleExaminerDashboard'],
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
    ['cancelRegistration', 'GET', 'none', 'handleCancelRegistration'],
    ['checkApproval', 'GET', 'none', 'handleCheckApproval'],
    ['getExamStatus', 'GET', 'none', 'handleGetExamStatus'],
    ['addExamTime', 'GET', 'none', 'handleAddExamTime'],
    ['markFinished', 'GET', 'none', 'handleMarkFinished'],
    // 'disqualify' is deliberately not examiner-gated: the examinee client sends
    // it too, and handleDisqualify accepts either an examiner token or an active
    // pending row of that examinee.
    ['disqualify', 'GET,POST', 'none', 'handleDisqualify'],
    ['reportWarning', 'GET,POST', 'none', 'handleReportWarning'],
    ['cancelDisqualify', 'GET,POST', 'none', 'handleCancelDisqualify'],
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
  var body = { status: 'ok', build: THEORY_API_BUILD, indexIds: questionIndexCount(),
    gateway: { url: Boolean(gatewayUrl()), key: Boolean(gatewayKey()) } };
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
