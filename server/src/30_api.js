// Public build marker: identifies the deployed API without reading private data.
var THEORY_API_BUILD = '2026-09-19-r24';
var THEORY_API_ACTIONS = ('health addExamTime adminDashboard approveExaminee cancelDisqualify cancelFailOnClose cancelRegistration centerManagerReport checkApproval closeSession commanderCorrectResult commanderDashboard confirmDQ correctExamineeMeta correctToPass createSession disqualify examinerDashboard examinerForecast forceComplete getExamQuestions getExamStatus getOfficeNumber getQuestionsByIds getResultUploadToken getSessionInfo getSites getUploadResult listActiveExaminers listAllSessions listSessions loadStudentProgress login markExamStarted markFinished markSent overturnDQ predictiveModelPreview registerExamQuestions registerExaminee rejectExaminee reportWarning resetExaminee saveStudentProgress searchQuestions siteCombinedReport studentJoinClass submitFailOnClose submitManualResult submitPracticeResult submitResult submitWrongAnswers teacherAtRiskList teacherClassDetails teacherCloseClass teacherCommanderDashboard teacherCreateClass teacherDashboard teacherDeleteClass teacherExportData teacherGetClasses teacherLogin teacherRemoveStudent teacherVerifyLogin updateSession uploadResultHtml verifyLogin viewResult').split(' ');

function logTheoryApiTiming(phase, method, action, startedAt) {
  // Never log request parameters, IDs, credentials, answers or arbitrary action text.
  // A start without an end can identify a runtime timeout in the execution log.
  try {
    Logger.log('[API] ' + JSON.stringify({ build: THEORY_API_BUILD, phase: phase,
      method: method, action: THEORY_API_ACTIONS.indexOf(action) >= 0 ? action : 'unknown',
      elapsedMs: Math.max(0, Date.now() - startedAt) }));
  } catch (logErr) { /* diagnostics must never break an exam */ }
}

function theoryRetryableErrorResponse(err) {
  if (!err || err.retryable !== true) return null;
  return jsonResponse({ status: 'error', code: 'question_cache_busy', retryable: true,
    waitSec: Math.max(1, Math.min(30, Number(err.waitSec) || 3)),
    message: 'מאגר השאלות מתעדכן כעת. אפשר לנסות שוב בעוד מספר שניות.' });
}

function questionRequestRateId(p, auth) {
  // A class starting together must not share one candidate's allowance.
  if (auth === 'examinee') return String(p.sessionCode || '') + '_' + normalizeId(p.idNumber);
  return p.idNumber || p.examinerId || p.studentId || p.standaloneIdNumber || p.sessionCode || 'anon';
}

// ========== doGet — קריאות קריאה + פעולות קלות ==========

function doGet(e) {
  var apiStartedAt = Date.now();
  var action = '';
  diagBegin('GET');
  try {
    var p = (e && e.parameter) || {};
    action = p.action || '';
    if (DIAG_EXEC) { DIAG_EXEC.action = action; DIAG_EXEC.t0 = apiStartedAt; }
    logTheoryApiTiming('start', 'GET', action, apiStartedAt);

    // Block sensitive state-mutating actions from GET — must come via POST.
    // Prevents URL-based forging (URLs leak to logs/history; trivially craftable).
    // Clients already use POST for these (apiPost / sendBeacon with JSON body).
    var postOnlyActions = ['submitResult','submitFailOnClose','submitWrongAnswers','uploadResultHtml','registerExamQuestions','saveStudentProgress','commanderCorrectResult','submitManualResult'];
    if (postOnlyActions.indexOf(action) !== -1) {
      return jsonResponse({ status: 'error', message: 'פעולה זו דורשת POST' });
    }

    // Soft origin check — log unauthorized origins (deterrent, bypassable but raises bar)
    var originErr = checkOrigin(p);
    if (originErr) return originErr;

    if (action === 'health') {
      if (String(p.deep || '') === '1') {
        // health&deep=1 (2026-09-19, review action 7): the plain health does no
        // work at all, so it can only say "Google is slow". This one also reads a
        // single cell of OUR document and reports that time separately, so a
        // watchdog can tell "our document stalls" from "Google's front door
        // stalls" every minute of an exam morning (tools/exam_watchdog.gs).
        var deepT0 = Date.now(), sheetMs = -1, sheetError = '';
        try { getSheet('אתרים').getRange(1, 1).getValue(); sheetMs = Date.now() - deepT0; }
        catch (eDeep) { sheetError = String(eDeep && eDeep.message ? eDeep.message : eDeep).slice(0, 120); }
        return jsonResponse({ status: 'ok', build: THEORY_API_BUILD, deep: true, sheetMs: sheetMs, sheetError: sheetError,
          totalMs: Date.now() - apiStartedAt });
      }
      return jsonResponse({ status: 'ok', build: THEORY_API_BUILD });
    }

    // Actions that require examiner token authentication
    var examinerActions = ['getSites','listSessions','listAllSessions','createSession','updateSession','closeSession',
      'approveExaminee','rejectExaminee','examinerDashboard','resetExaminee',
      'correctToPass','overturnDQ','confirmDQ','forceComplete','markSent','commanderDashboard',
      'commanderCorrectResult','correctExamineeMeta','getResultUploadToken','centerManagerReport',
      'predictiveModelPreview','examinerForecast'];
    // Note: 'disqualify' is NOT in this list because it can be sent by the examinee client (no token)
    // — auth is enforced inside handleDisqualify itself (examiner token OR active pending row).
    if (examinerActions.indexOf(action) !== -1) {
      var tokenErr = requireToken(p);
      if (tokenErr) return tokenErr;
    }

    // Actions that require teacher token authentication
    var teacherActions = ['teacherDashboard','teacherCreateClass','teacherCloseClass','teacherDeleteClass',
      'teacherRemoveStudent','teacherGetClasses','teacherClassDetails','teacherExportData',
      'teacherCommanderDashboard','teacherAtRiskList','adminDashboard'];
    if (teacherActions.indexOf(action) !== -1) {
      var tErr = requireTeacherToken(p);
      if (tErr) return tErr;
    }

    switch (action) {

      case 'login':
        // Login only via POST — block GET to prevent password in URL
        return jsonResponse({ status: 'error', message: 'יש להתחבר דרך POST בלבד' });

      case 'verifyLogin':
        return handleVerifyLogin(p);

      case 'getSites':
        return handleGetSites();

      case 'listSessions':
        return handleListSessions(p);

      case 'listAllSessions':
        return handleListAllSessions(p);

      case 'centerManagerReport':
        return handleCenterManagerReport(p);

      case 'getOfficeNumber':
        // Public read of the office WA number — used by clients for display.
        return jsonResponse({ status: 'ok', officeWhatsApp: getOfficeWhatsAppNumber() });

      case 'createSession':
        return handleCreateSession(p);

      case 'listActiveExaminers':
        return handleListActiveExaminers(p);

      case 'siteCombinedReport':
        return handleSiteCombinedReport(p);

      case 'updateSession':
        return handleUpdateSession(p);

      case 'closeSession':
        return handleCloseSession(p);

      case 'getSessionInfo':
        return handleGetSessionInfo(p);

      case 'registerExaminee':
        return handleRegisterExaminee(p);

      case 'cancelRegistration':
        return handleCancelRegistration(p);

      case 'checkApproval':
        return handleCheckApproval(p);

      case 'approveExaminee':
        return handleApproveExaminee(p);

      case 'rejectExaminee':
        return handleRejectExaminee(p);

      case 'markExamStarted':
        return handleMarkExamStarted(p);

      case 'examinerDashboard':
        return handleExaminerDashboard(p);

      case 'disqualify':
        return handleDisqualify(p);

      case 'reportWarning':
        return handleReportWarning(p);

      case 'getExamStatus':
        return handleGetExamStatus(p);

      case 'addExamTime':
        return handleAddExamTime(p);

      case 'markFinished':
        return handleMarkFinished(p);

      case 'cancelDisqualify':
        return handleCancelDisqualify(p);

      case 'resetExaminee':
        return handleResetExaminee(p);

      case 'overturnDQ':
        return handleOverturnDQ(p);

      case 'confirmDQ':
        return handleConfirmDQ(p);

      case 'correctToPass':
        return handleCorrectToPass(p);

      case 'correctExamineeMeta':
        return handleCorrectExamineeMeta(p);

      case 'forceComplete':
        return handleForceComplete(p);

      case 'markSent':
        return handleMarkSent(p);

      case 'commanderDashboard':
        return handleCommanderDashboard(p);

      case 'predictiveModelPreview':
        return handlePredictiveModelPreview(p);

      case 'examinerForecast':
        return handleExaminerForecast(p);

      case 'submitResult':
        // Decode wrongAnswers from JSON string parameter
        var resultData = {
          action: 'submitResult',
          sessionCode: p.sessionCode || '',
          idNumber: p.idNumber || '',
          fullName: p.fullName || '',
          phone: p.phone || '',
          license: p.license || 'B',
          language: p.language || 'he',
          score: Number(p.score) || 0,
          total: Number(p.total) || 30,
          percent: Number(p.percent) || 0,
          passed: p.passed === 'true' || p.passed === true,
          time: p.time || '',
          examinerName: p.examinerName || '',
          site: p.site || '',
          classroom: p.classroom || '',
          population: p.population || '',
          audioMode: p.audioMode || 'off',
          device: p.device || '',
          wrongAnswers: []
        };
        try { if (p.wrongAnswers) resultData.wrongAnswers = JSON.parse(p.wrongAnswers); } catch(ex) {}
        return handleSubmitResult(resultData);

      case 'submitWrongAnswers':
        return handleSubmitWrongAnswers(p);

      case 'submitFailOnClose':
        var failData = {
          action: 'submitFailOnClose',
          sessionCode: p.sessionCode || '',
          idNumber: p.idNumber || '',
          fullName: p.fullName || '',
          phone: p.phone || '',
          license: p.license || 'B',
          language: p.language || 'he',
          examinerName: p.examinerName || '',
          site: p.site || '',
          classroom: p.classroom || '',
          answeredCount: Number(p.answeredCount) || 0,
          totalQuestions: Number(p.totalQuestions) || 30,
          time: p.time || '',
          population: p.population || '',
          audioMode: p.audioMode || 'off',
          device: p.device || ''
        };
        return handleSubmitFailOnClose(failData);

      case 'getUploadResult':
        return handleGetUploadResult(p);

      case 'getResultUploadToken':
        return handleGetResultUploadToken(p);

      case 'getExamQuestions':
        return handleGetExamQuestions(p);

      case 'searchQuestions':
        return handleSearchQuestions(p);

      case 'getQuestionsByIds':
        return handleGetQuestionsByIds(p);

      case 'viewResult':
        // DISABLED: see handleUploadResultHtml. Result viewing moved to the
        // authenticated Cloudflare Worker; this no longer serves cached HTML (it
        // used ALLOWALL framing on the trusted Google origin \u2014 an XSS/phishing vector).
        return HtmlService.createHtmlOutput('<h1 style="text-align:center;padding:40px;font-family:Arial;direction:rtl;">\u05DC\u05D0 \u05D6\u05DE\u05D9\u05DF</h1>');

      // ===== Teacher actions =====
      case 'teacherVerifyLogin':
        return handleTeacherVerifyLogin(p);

      case 'teacherGetClasses':
        return handleTeacherGetClasses(p);

      case 'teacherCreateClass':
        return handleTeacherCreateClass(p);

      case 'teacherCloseClass':
        return handleTeacherCloseClass(p);

      case 'teacherDeleteClass':
        return handleTeacherDeleteClass(p);

      case 'teacherRemoveStudent':
        return handleTeacherRemoveStudent(p);

      case 'teacherDashboard':
        return handleTeacherDashboard(p);

      case 'teacherClassDetails':
        return handleTeacherClassDetails(p);

      case 'teacherExportData':
        return handleTeacherExportData(p);

      case 'teacherCommanderDashboard':
        return handleTeacherCommanderDashboard(p);

      case 'teacherAtRiskList':
        return handleTeacherAtRiskList(p);

      case 'adminDashboard':
        return handleAdminDashboard(p);

      // ===== Student join class (no auth) =====
      case 'studentJoinClass':
        return handleStudentJoinClass(p);

      case 'submitPracticeResult':
        return handleSubmitPracticeResult(p);

      case 'loadStudentProgress':
        return handleLoadStudentProgress(p);
      default:
        return jsonResponse({ status: 'ok', message: 'External Exam API is running' });
    }

  } catch (err) {
    return theoryRetryableErrorResponse(err) || jsonResponse({ status: 'error', message: err.toString() });
  } finally {
    diagFinish(action, apiStartedAt);
    logTheoryApiTiming('end', 'GET', action, apiStartedAt);
  }
}

// ========== doPost — שמירת תוצאות (נתונים גדולים) ==========

function doPost(e) {
  var apiStartedAt = Date.now();
  var action = '';
  diagBegin('POST');
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse({ status: 'error', message: 'No POST data received' });
    }
    var raw = e.postData.contents;
    var data = JSON.parse(raw);
    action = data.action || '';
    if (DIAG_EXEC) { DIAG_EXEC.action = action; DIAG_EXEC.t0 = apiStartedAt; }
    logTheoryApiTiming('start', 'POST', action, apiStartedAt);

    // Soft origin check (deters casual scripts; bypassable by reading client source)
    var originErr = checkOrigin(data);
    if (originErr) return originErr;

    if (action === 'login') {
      return handleLogin(data);
    } else if (action === 'teacherLogin') {
      return handleTeacherLogin(data);
    } else if (action === 'submitPracticeResult') {
      return handleSubmitPracticeResult(data);
    } else if (action === 'registerExamQuestions') {
      return handleRegisterExamQuestions(data);
    } else if (action === 'submitResult') {
      return handleSubmitResult(data);
    } else if (action === 'submitFailOnClose') {
      return handleSubmitFailOnClose(data);
    } else if (action === 'submitWrongAnswers') {
      return handleSubmitWrongAnswersBulk(data);
    } else if (action === 'cancelFailOnClose') {
      return handleCancelFailOnClose(data);
    } else if (action === 'uploadResultHtml') {
      return handleUploadResultHtml(data);
    } else if (action === 'disqualify') {
      return handleDisqualify(data);
    } else if (action === 'reportWarning') {
      return handleReportWarning(data);
    } else if (action === 'cancelDisqualify') {
      return handleCancelDisqualify(data);
    } else if (action === 'saveStudentProgress') {
      return handleSaveStudentProgress(data);
    } else if (action === 'commanderCorrectResult') {
      return handleCommanderCorrectResult(data);
    } else if (action === 'correctExamineeMeta') {
      return handleCorrectExamineeMeta(data);
    } else if (action === 'submitManualResult') {
      return handleSubmitManualResult(data);
    } else {
      return jsonResponse({ status: 'error', message: 'Unknown POST action: ' + action });
    }

  } catch (err) {
    return theoryRetryableErrorResponse(err) || jsonResponse({ status: 'error', message: 'doPost error: ' + err.toString() });
  } finally {
    diagFinish(action, apiStartedAt);
    logTheoryApiTiming('end', 'POST', action, apiStartedAt);
  }
}

