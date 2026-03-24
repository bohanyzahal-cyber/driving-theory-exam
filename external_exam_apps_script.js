// © 2026 Vitaly Gitelman. All Rights Reserved.
// Unauthorized copying, modification or distribution is prohibited.
// ===== Google Apps Script — מערכת בחינות חיצונית =====
// הדבק את הקוד הזה ב-Apps Script של גיליון Google Sheets חדש
// Deploy → New deployment → Web app
// Execute as: Me | Who has access: Anyone
// העתק את ה-URL שמקבלים והדבק ב-examiner.html וב-examinee.html

// ========== פונקציות עזר ==========

var SHEET_HEADERS = {
  'בוחנים': ['שם', 'ת.ז.', 'סיסמה', 'פעיל', 'מס בוחן', 'תפקיד', 'טוקן', 'תוקף טוקן', 'ניסיונות כושלים', 'נעילה עד'],
  'אתרים': ['שם אתר', 'מזהה', 'טלפון מנהל', 'כיתות'],
  'סשנים': ['קוד', 'בוחן ת.ז.', 'שם בוחן', 'אתר', 'כיתה', 'דרגה', 'שפה', 'מצב שמע', 'זמן יצירה', 'תקף עד', 'פעיל'],
  'ממתינים': ['קוד סשן', 'ת.ז.', 'שם', 'טלפון', 'זמן הרשמה', 'סטטוס', 'שפה', 'אוכלוסיה', 'דרגה', 'שמע', 'הארכת זמן'],
  'תוצאות': ['תאריך', 'ת.ז.', 'שם', 'טלפון', 'דרגה', 'ציון', 'אחוז', 'עבר/נכשל', 'זמן', 'בוחן', 'אתר', 'כיתה', 'שפה', 'קוד סשן', 'ניסיון', 'פירוט שגויות', 'נשלח?', 'פסול?', 'קישור וואטסאפ', 'אוכלוסיה', 'תוקן?', 'שמע', 'מאומת', 'חשוד', 'dqEventId'],
  'מורים': ['שם', 'ת.ז.', 'סיסמה', 'פעיל', 'טוקן', 'תוקף טוקן', 'ניסיונות כושלים', 'נעילה עד'],
  'כיתות': ['קוד כיתה', 'שם כיתה', 'מורה ת.ז.', 'שם מורה', 'דרגה', 'תאריך יצירה', 'פעיל'],
  'תלמידי כיתות': ['קוד כיתה', 'שם תלמיד', 'מזהה תלמיד', 'תאריך הצטרפות'],
  'תוצאות תרגול': ['תאריך', 'מזהה תלמיד', 'שם תלמיד', 'קוד כיתה', 'מצב', 'דרגה', 'ציון', 'סה"כ', 'אחוז', 'עבר/נכשל', 'זמן', 'נושא', 'שפה', 'פירוט שגויות', 'פירוט לפי נושא']
};

function getSheet(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    var headers = SHEET_HEADERS[name];
    if (headers) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    }
  }
  return sheet;
}

function findRow(sheet, colIndex, value) {
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][colIndex]) === String(value)) return i + 1;
  }
  return -1;
}

function findAllRows(sheet, colIndex, value) {
  var data = sheet.getDataRange().getValues();
  var results = [];
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][colIndex]) === String(value)) results.push({ row: i + 1, data: data[i] });
  }
  return results;
}

function generateSessionCode() {
  var sessSheet = getSheet('סשנים');
  var data = sessSheet.getDataRange().getValues();
  var existingCodes = {};
  for (var i = 1; i < data.length; i++) {
    // Check ALL session codes (not just active) to prevent data mixing with closed sessions
    existingCodes[String(data[i][0]).trim()] = true;
  }
  // 8-character alphanumeric code (unambiguous chars: no O/0/I/1/L)
  var chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  var code;
  do {
    code = '';
    for (var c = 0; c < 8; c++) {
      code += chars.charAt(Math.floor(Math.random() * chars.length));
    }
  } while (existingCodes[code]);
  return code;
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ========== Token authentication ==========
function generateToken() {
  var chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  var token = '';
  for (var i = 0; i < 48; i++) {
    token += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return token;
}

function verifyToken(examinerId, token) {
  if (!examinerId || !token) return false;
  var sheet = getSheet('בוחנים');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (normalizeId(data[i][1]) === normalizeId(examinerId)) {
      var storedTokens = String(data[i][6] || '').split(',');
      var expiry = data[i][7];
      if (storedTokens.indexOf(token) === -1) return false;
      if (!expiry) return false;
      var expiryDate = expiry instanceof Date ? expiry : new Date(expiry);
      if (new Date() > expiryDate) return false;
      return true;
    }
  }
  return false;
}

function requireToken(p) {
  if (!verifyToken(p.examinerId, p.token)) {
    return jsonResponse({ status: 'error', message: 'טוקן לא תקין — יש להתחבר מחדש', tokenExpired: true });
  }
  return null;
}

// Verify examiner owns the session (for sensitive actions)
function verifyExaminerForSession(sessionCode, examinerId) {
  if (!examinerId) return false;
  var sheet = getSheet('סשנים');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(sessionCode).trim()) {
      return normalizeId(data[i][1]) === normalizeId(examinerId);
    }
  }
  return false;
}

function countAttempts(idNumber, license) {
  var sheet = getSheet('תוצאות');
  var data = sheet.getDataRange().getValues();
  var count = 0;
  for (var i = 1; i < data.length; i++) {
    if (normalizeId(data[i][1]) === normalizeId(idNumber) && String(data[i][4]) === String(license)) {
      var status = String(data[i][7] || '').trim();
      if (status === 'בוטל') continue; // overturned DQ is not a real attempt
      count++;
    }
  }
  return count;
}

function formatPhoneForWA(phone) {
  phone = String(phone || '').replace(/[^0-9]/g, '');
  if (phone.charAt(0) === '0') phone = '972' + phone.substring(1);
  else if (phone.length === 9 && phone.charAt(0) === '5') phone = '972' + phone;
  return phone;
}

function normalizeId(val) {
  var s = String(val || '').replace(/[^0-9]/g, '');
  while (s.length < 9) s = '0' + s;
  return s;
}

function nowISO() {
  return new Date().toISOString();
}

function todayStr() {
  var d = new Date();
  var dd = ('0' + d.getDate()).slice(-2);
  var mm = ('0' + (d.getMonth() + 1)).slice(-2);
  var yyyy = d.getFullYear();
  var hh = ('0' + d.getHours()).slice(-2);
  var mi = ('0' + d.getMinutes()).slice(-2);
  return dd + '/' + mm + '/' + yyyy + ' ' + hh + ':' + mi;
}

// ========== doGet — קריאות קריאה + פעולות קלות ==========

function doGet(e) {
  try {
    var p = e.parameter || {};
    var action = p.action || '';

    // Actions that require examiner token authentication
    var examinerActions = ['getSites','listSessions','createSession','updateSession','closeSession',
      'approveExaminee','rejectExaminee','examinerDashboard','resetExaminee',
      'correctToPass','overturnDQ','forceComplete','markSent','commanderDashboard'];
    // Note: 'disqualify' is NOT in this list because it can be sent by the examinee client (no token)
    if (examinerActions.indexOf(action) !== -1) {
      var tokenErr = requireToken(p);
      if (tokenErr) return tokenErr;
    }

    // Actions that require teacher token authentication
    var teacherActions = ['teacherDashboard','teacherCreateClass','teacherCloseClass',
      'teacherRemoveStudent','teacherGetClasses','teacherClassDetails','teacherExportData',
      'teacherCommanderDashboard','adminDashboard'];
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

      case 'createSession':
        return handleCreateSession(p);

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

      case 'forceComplete':
        return handleForceComplete(p);

      case 'markSent':
        return handleMarkSent(p);

      case 'commanderDashboard':
        return handleCommanderDashboard(p);

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
          audioMode: p.audioMode || 'off'
        };
        return handleSubmitFailOnClose(failData);

      case 'getUploadResult':
        return handleGetUploadResult(p);

      case 'viewResult':
        var resultId = p.id;
        if (!resultId) return HtmlService.createHtmlOutput('<h1 style="text-align:center;padding:40px;font-family:Arial;">Missing ID</h1>');
        var vc = CacheService.getScriptCache();
        var metaStr = vc.get('result_' + resultId + '_meta');
        if (!metaStr) return HtmlService.createHtmlOutput('<h1 style="text-align:center;padding:40px;font-family:Arial;direction:rtl;">\u05D4\u05E7\u05D9\u05E9\u05D5\u05E8 \u05E4\u05D2 \u05EA\u05D5\u05E7\u05E3</h1>');
        var numChunks = parseInt(metaStr, 10);
        var keys = [];
        for (var ci = 0; ci < numChunks; ci++) keys.push('result_' + resultId + '_' + ci);
        var chunkMap = vc.getAll(keys);
        var fullHtml = '';
        for (var ci2 = 0; ci2 < numChunks; ci2++) {
          var chunk = chunkMap['result_' + resultId + '_' + ci2];
          if (!chunk) return HtmlService.createHtmlOutput('<h1 style="text-align:center;padding:40px;font-family:Arial;direction:rtl;">\u05D4\u05E7\u05D9\u05E9\u05D5\u05E8 \u05E4\u05D2 \u05EA\u05D5\u05E7\u05E3</h1>');
          fullHtml += chunk;
        }
        return HtmlService.createHtmlOutput(fullHtml).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

      // ===== Teacher actions =====
      case 'teacherVerifyLogin':
        return handleTeacherVerifyLogin(p);

      case 'teacherGetClasses':
        return handleTeacherGetClasses(p);

      case 'teacherCreateClass':
        return handleTeacherCreateClass(p);

      case 'teacherCloseClass':
        return handleTeacherCloseClass(p);

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

      case 'adminDashboard':
        return handleAdminDashboard(p);

      // ===== Student join class (no auth) =====
      case 'studentJoinClass':
        return handleStudentJoinClass(p);

      case 'submitPracticeResult':
        return handleSubmitPracticeResult(p);

      default:
        return jsonResponse({ status: 'ok', message: 'External Exam API is running' });
    }

  } catch (err) {
    return jsonResponse({ status: 'error', message: err.toString() });
  }
}

// ========== doPost — שמירת תוצאות (נתונים גדולים) ==========

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse({ status: 'error', message: 'No POST data received' });
    }
    var raw = e.postData.contents;
    var data = JSON.parse(raw);
    var action = data.action || '';

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
    } else if (action === 'cancelDisqualify') {
      return handleCancelDisqualify(data);
    } else {
      return jsonResponse({ status: 'error', message: 'Unknown POST action: ' + action });
    }

  } catch (err) {
    return jsonResponse({ status: 'error', message: 'doPost error: ' + err.toString() });
  }
}

// ========== handlers ==========

function handleLogin(p) {
  var sheet = getSheet('בוחנים');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (normalizeId(data[i][1]) === normalizeId(p.idNumber)) {
      var row = i + 1; // sheet rows are 1-indexed
      // Rate limiting: column I (index 8) = failed attempts, column J (index 9) = lockout until
      var failedAttempts = Number(data[i][8]) || 0;
      var lockoutUntil = data[i][9];
      if (lockoutUntil) {
        var lockoutDate = lockoutUntil instanceof Date ? lockoutUntil : new Date(lockoutUntil);
        if (new Date() < lockoutDate) {
          var minsLeft = Math.ceil((lockoutDate - new Date()) / 60000);
          return jsonResponse({ status: 'error', message: 'החשבון נעול עקב ניסיונות כושלים. נסה שוב בעוד ' + minsLeft + ' דקות' });
        }
        // Lockout expired — reset counter
        failedAttempts = 0;
        sheet.getRange(row, 9).setValue(0);    // column I = failed attempts reset
        sheet.getRange(row, 10).setValue('');   // column J = lockout cleared
      }
      if (String(data[i][2]) === String(p.password)) {
        if (data[i][3] === 'כן' || data[i][3] === true || data[i][3] === 'TRUE') {
          // Successful login — reset failed attempts
          if (failedAttempts > 0) {
            sheet.getRange(row, 9).setValue(0);    // column I = failed attempts reset
            sheet.getRange(row, 10).setValue('');   // column J = lockout cleared
          }
          // Generate token and store in sheet (columns G=7, H=8 → indices 6,7)
          // Support multiple tokens (multi-device) separated by comma, max 5
          var token = generateToken();
          var expiry = new Date();
          expiry.setHours(expiry.getHours() + 12);
          var existingTokens = String(data[i][6] || '').trim();
          var tokenList = existingTokens ? existingTokens.split(',') : [];
          tokenList.push(token);
          if (tokenList.length > 5) tokenList = tokenList.slice(-5); // keep last 5
          sheet.getRange(row, 7).setValue(tokenList.join(','));   // column G = tokens
          sheet.getRange(row, 8).setValue(expiry);   // column H = expiry
          return jsonResponse({ status: 'ok', examiner: { name: data[i][0], id: normalizeId(data[i][1]), examinerNumber: String(data[i][4] || ''), role: String(data[i][5] || 'בוחן'), token: token } });
        } else {
          return jsonResponse({ status: 'error', message: 'החשבון אינו פעיל' });
        }
      } else {
        // Wrong password — increment failed attempts
        failedAttempts++;
        sheet.getRange(row, 9).setValue(failedAttempts);   // column I = failed attempts
        if (failedAttempts >= 5) {
          var lockout = new Date();
          lockout.setMinutes(lockout.getMinutes() + 15);
          sheet.getRange(row, 10).setValue(lockout);        // column J = lockout until
          return jsonResponse({ status: 'error', message: 'יותר מדי ניסיונות כושלים. החשבון ננעל ל-15 דקות' });
        }
        return jsonResponse({ status: 'error', message: 'סיסמה שגויה' });
      }
    }
  }
  return jsonResponse({ status: 'error', message: 'בוחן לא נמצא' });
}

function handleVerifyLogin(p) {
  if (!p.examinerId || !p.token) {
    return jsonResponse({ status: 'error', message: 'חסרים פרטי אימות', tokenExpired: true });
  }
  var sheet = getSheet('בוחנים');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (normalizeId(data[i][1]) === normalizeId(p.examinerId)) {
      var storedTokens = String(data[i][6] || '').split(',');
      var expiry = data[i][7];
      if (storedTokens.indexOf(p.token) === -1) {
        return jsonResponse({ status: 'error', message: 'טוקן לא תקין', tokenExpired: true });
      }
      if (!expiry) {
        return jsonResponse({ status: 'error', message: 'טוקן לא תקין', tokenExpired: true });
      }
      var expiryDate = expiry instanceof Date ? expiry : new Date(expiry);
      if (new Date() > expiryDate) {
        return jsonResponse({ status: 'error', message: 'פג תוקף ההתחברות', tokenExpired: true });
      }
      if (!(data[i][3] === 'כן' || data[i][3] === true || data[i][3] === 'TRUE')) {
        return jsonResponse({ status: 'error', message: 'החשבון אינו פעיל' });
      }
      return jsonResponse({ status: 'ok', examiner: { name: data[i][0], id: normalizeId(data[i][1]), examinerNumber: String(data[i][4] || ''), role: String(data[i][5] || 'בוחן'), token: p.token } });
    }
  }
  return jsonResponse({ status: 'error', message: 'בוחן לא נמצא', tokenExpired: true });
}

function handleGetSites() {
  var sheet = getSheet('אתרים');
  var data = sheet.getDataRange().getValues();
  var sites = [];
  for (var i = 1; i < data.length; i++) {
    var classrooms = String(data[i][3] || '').split(',').map(function(s) { return s.trim(); }).filter(function(s) { return s; });
    sites.push({
      name: data[i][0],
      id: data[i][1],
      managerPhone: data[i][2],
      classrooms: classrooms
    });
  }
  return jsonResponse({ status: 'ok', sites: sites });
}

function handleListSessions(p) {
  var sheet = getSheet('סשנים');
  var data = sheet.getDataRange().getValues();
  var examinerId = normalizeId(p.examinerId);
  var sitesSheet = getSheet('אתרים');
  var sitesData = sitesSheet.getDataRange().getValues();
  // Build sites lookup for manager phone
  var sitesMap = {};
  for (var s = 1; s < sitesData.length; s++) {
    sitesMap[String(sitesData[s][0]).trim()] = { managerPhone: sitesData[s][2] || '' };
  }
  var sessions = [];
  for (var i = data.length - 1; i >= 1; i--) {
    if (normalizeId(data[i][1]) === examinerId) {
      var siteName = String(data[i][3] || '').trim();
      sessions.push({
        code: String(data[i][0]),
        site: data[i][3] || '',
        classroom: data[i][4] || '',
        license: data[i][5] || '',
        language: data[i][6] || 'he',
        audioMode: data[i][7] || 'off',
        created: data[i][8] || '',
        validUntil: data[i][9] || '',
        active: data[i][10] === true || String(data[i][10]).toUpperCase() === 'TRUE',
        managerPhone: sitesMap[siteName] ? sitesMap[siteName].managerPhone : ''
      });
    }
  }
  // Return up to 20 most recent sessions
  return jsonResponse({ status: 'ok', sessions: sessions.slice(0, 20) });
}

function handleCreateSession(p) {
  var sheet = getSheet('סשנים');
  var code = generateSessionCode();
  var now = new Date();
  var validUntil = new Date(now.getTime() + 8 * 60 * 60 * 1000);

  // Lookup examiner name (normalize ID to handle leading zeros)
  var exSheet = getSheet('בוחנים');
  var exData = exSheet.getDataRange().getValues();
  var exRow = -1;
  var examinerName = '';
  for (var ei = 1; ei < exData.length; ei++) {
    if (normalizeId(exData[ei][1]) === normalizeId(p.examinerId)) { exRow = ei + 1; examinerName = exData[ei][0]; break; }
  }

  sheet.appendRow([
    code,
    p.examinerId,
    examinerName,
    p.site || '',
    p.classroom || '',
    p.license || 'B',
    p.language || 'he',
    p.audioMode || 'off',
    now.toISOString(),
    validUntil.toISOString(),
    true
  ]);

  return jsonResponse({ status: 'ok', sessionCode: code, validUntil: validUntil.toISOString(), examinerName: examinerName });
}

function handleUpdateSession(p) {
  var sheet = getSheet('סשנים');
  var data = sheet.getDataRange().getValues();
  var searchCode = String(p.sessionCode).trim();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === searchCode && (data[i][10] === true || String(data[i][10]).toUpperCase() === 'TRUE')) {
      if (normalizeId(data[i][1]) !== normalizeId(p.examinerId)) {
        return jsonResponse({ status: 'error', message: 'אין הרשאה לעדכן סשן זה' });
      }
      var row = i + 1;
      if (p.license) sheet.getRange(row, 6).setValue(p.license);
      if (p.language) sheet.getRange(row, 7).setValue(p.language);
      if (p.audioMode) sheet.getRange(row, 8).setValue(p.audioMode);
      return jsonResponse({ status: 'ok' });
    }
  }
  return jsonResponse({ status: 'error', message: 'סשן לא נמצא' });
}

function handleCloseSession(p) {
  var sheet = getSheet('סשנים');
  var data = sheet.getDataRange().getValues();
  var searchCode = String(p.sessionCode).trim();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === searchCode && normalizeId(data[i][1]) === normalizeId(p.examinerId)) {
      sheet.getRange(i + 1, 11).setValue(false);
      return jsonResponse({ status: 'ok' });
    }
  }
  return jsonResponse({ status: 'error', message: 'סשן לא נמצא' });
}

function handleGetSessionInfo(p) {
  var sheet = getSheet('סשנים');
  var data = sheet.getDataRange().getValues();
  var searchCode = String(p.sessionCode).trim();
  for (var i = 1; i < data.length; i++) {
    var rowCode = String(data[i][0]).trim();
    if (rowCode === searchCode) {
      var active = data[i][10];
      if (active !== true && active !== 'TRUE' && String(active).toUpperCase() !== 'TRUE') {
        return jsonResponse({ status: 'error', message: 'הסשן הסתיים' });
      }
      var validUntil = new Date(data[i][9]);
      if (new Date() > validUntil) {
        sheet.getRange(i + 1, 11).setValue(false);
        return jsonResponse({ status: 'error', message: 'תוקף הסשן פג' });
      }
      return jsonResponse({
        status: 'ok',
        session: {
          site: data[i][3],
          classroom: data[i][4],
          license: data[i][5],
          language: data[i][6],
          audioMode: data[i][7],
          examinerName: data[i][2],
          validUntil: data[i][9]
        }
      });
    }
  }
  return jsonResponse({ status: 'error', message: 'קוד סשן לא תקין' });
}

function handleRegisterExaminee(p) {
  var MAX_PENDING_PER_SESSION = 50;
  var pendSheet = getSheet('ממתינים');
  var data = pendSheet.getDataRange().getValues();
  var activeCount = 0;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(p.sessionCode)) {
      var status = String(data[i][5] || '').trim();
      if (normalizeId(data[i][1]) === normalizeId(p.idNumber)) {
        if (status === 'waiting' || status === 'approved' || status === 'in_exam') {
          return jsonResponse({ status: 'error', message: 'כבר רשום בסשן זה' });
        }
      }
      if (status === 'waiting' || status === 'approved' || status === 'in_exam') {
        activeCount++;
      }
    }
  }
  if (activeCount >= MAX_PENDING_PER_SESSION) {
    return jsonResponse({ status: 'error', message: 'הסשן מלא — לא ניתן לרשום נבחנים נוספים' });
  }
  pendSheet.appendRow([
    p.sessionCode,
    p.idNumber,
    p.fullName || '',
    p.phone || '',
    nowISO(),
    'waiting',
    p.language || '',
    p.population || '',
    p.license || '',
    p.audioMode || 'off',
    ''  // הארכת זמן — נקבע ע"י הבוחן בעת אישור
  ]);
  return jsonResponse({ status: 'ok' });
}

function handleCancelRegistration(p) {
  var sheet = getSheet('ממתינים');
  var data = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === String(p.sessionCode) && normalizeId(data[i][1]) === normalizeId(p.idNumber)) {
      var s = String(data[i][5]).trim();
      if (s === 'waiting' || s === 'approved') {
        // Verify phone matches to prevent unauthorized cancellation
        var storedPhone = String(data[i][3] || '').replace(/[^0-9]/g, '');
        var givenPhone = String(p.phone || '').replace(/[^0-9]/g, '');
        if (storedPhone && givenPhone && storedPhone.slice(-7) !== givenPhone.slice(-7)) {
          return jsonResponse({ status: 'error', message: 'פרטים לא תואמים' });
        }
        sheet.getRange(i + 1, 6).setValue('cancelled');
        SpreadsheetApp.flush();
        return jsonResponse({ status: 'ok' });
      }
    }
  }
  return jsonResponse({ status: 'error', message: 'לא נמצא רישום פעיל לביטול' });
}

function handleCheckApproval(p) {
  var BASE_EXAM_MINUTES = 40;
  var sheet = getSheet('ממתינים');
  var data = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]).trim() === String(p.sessionCode).trim() && normalizeId(data[i][1]) === normalizeId(p.idNumber)) {
      var approval = String(data[i][5] || 'waiting').trim();
      // Skip terminal statuses from previous exams — keep looking for active row
      // Note: dq_confirmed is NOT skipped — examinee needs to receive this status
      if (approval === 'completed' || approval === 'disqualified' || approval === 'cancelled') continue;
      var response = { status: 'ok', approval: approval };
      // When approved, compute and return authorized exam duration
      if (approval === 'approved' || approval === 'in_exam') {
        var ext = parseFloat(data[i][10]) || 1;
        if (ext !== 1.25 && ext !== 1.5) ext = 1;
        response.examMinutes = Math.round(BASE_EXAM_MINUTES * ext);
      }
      return jsonResponse(response);
    }
  }
  return jsonResponse({ status: 'error', message: 'לא נמצא רישום' });
}

function handleApproveExaminee(p) {
  if (p.examinerId && !verifyExaminerForSession(p.sessionCode, p.examinerId)) {
    return jsonResponse({ status: 'error', message: 'אין הרשאה — בוחן לא תואם לסשן' });
  }
  // Validate time extension (whitelist)
  var validExt = { '': true, '1.25': true, '1.5': true };
  var timeExt = String(p.timeExtension || '');
  if (!validExt[timeExt]) timeExt = '';

  var sheet = getSheet('ממתינים');
  var data = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === String(p.sessionCode) && normalizeId(data[i][1]) === normalizeId(p.idNumber) && String(data[i][5]).trim() === 'waiting') {
      sheet.getRange(i + 1, 6).setValue('approved');
      if (timeExt) sheet.getRange(i + 1, 11).setValue(timeExt);  // column K = הארכת זמן
      SpreadsheetApp.flush();
      return jsonResponse({ status: 'ok' });
    }
  }
  return jsonResponse({ status: 'error', message: 'נבחן ממתין לא נמצא (סטטוס נוכחי: ' + (data.length > 1 ? findStatus(data, p) : 'אין נתונים') + ')' });
}

function findStatus(data, p) {
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === String(p.sessionCode) && normalizeId(data[i][1]) === normalizeId(p.idNumber)) {
      return String(data[i][5]);
    }
  }
  return 'לא נמצא';
}

function handleRejectExaminee(p) {
  if (p.examinerId && !verifyExaminerForSession(p.sessionCode, p.examinerId)) {
    return jsonResponse({ status: 'error', message: 'אין הרשאה — בוחן לא תואם לסשן' });
  }
  var sheet = getSheet('ממתינים');
  var data = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === String(p.sessionCode) && normalizeId(data[i][1]) === normalizeId(p.idNumber) && String(data[i][5]).trim() === 'waiting') {
      sheet.getRange(i + 1, 6).setValue('rejected');
      SpreadsheetApp.flush();
      return jsonResponse({ status: 'ok' });
    }
  }
  return jsonResponse({ status: 'error', message: 'נבחן ממתין לא נמצא' });
}

function handleMarkExamStarted(p) {
  var sheet = getSheet('ממתינים');
  var data = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === String(p.sessionCode) && normalizeId(data[i][1]) === normalizeId(p.idNumber) && String(data[i][5]).trim() === 'approved') {
      sheet.getRange(i + 1, 6).setValue('in_exam');
      sheet.getRange(i + 1, 12).setValue(nowISO()); // column L = exam actual start time
      SpreadsheetApp.flush();
      return jsonResponse({ status: 'ok' });
    }
  }
  return jsonResponse({ status: 'error', message: 'נבחן מאושר לא נמצא' });
}

function handleExaminerDashboard(p) {
  var code = String(p.sessionCode);
  var pendSheet = getSheet('ממתינים');
  var resSheet = getSheet('תוצאות');

  var pendData = pendSheet.getDataRange().getValues();
  var resData = resSheet.getDataRange().getValues();
  var pending = [];
  var active = [];

  // Auto-cleanup: detect stale in_exam entries that already have a result or are way past exam time
  var now = new Date();
  var BASE_EXAM_MS = 40 * 60 * 1000;
  var STALE_BUFFER_MS = 20 * 60 * 1000; // 20 minutes buffer (approval wait + instructions)
  for (var ci = 1; ci < pendData.length; ci++) {
    if (String(pendData[ci][0]) !== code) continue;
    if (String(pendData[ci][5]).trim() !== 'in_exam') continue;
    var ciId = pendData[ci][1];
    var examStart = pendData[ci][11] ? new Date(pendData[ci][11]) : null;
    var regTime = examStart || (pendData[ci][4] ? new Date(pendData[ci][4]) : null);
    // Dynamic stale threshold: exam time (based on extension) + buffer
    var ciExt = parseFloat(pendData[ci][10]) || 1;
    if (ciExt !== 1.25 && ciExt !== 1.5) ciExt = 1;
    var maxMs = Math.round(BASE_EXAM_MS * ciExt) + STALE_BUFFER_MS;
    var isStale = regTime && (now.getTime() - regTime.getTime() > maxMs);

    // Count results by type for this examinee in this session
    var dqResults = 0, otherResults = 0;
    for (var ri = 1; ri < resData.length; ri++) {
      if (String(resData[ri][13]) === code && normalizeId(resData[ri][1]) === normalizeId(ciId)) {
        if (String(resData[ri][7] || '') === 'בוטל') continue;
        if (String(resData[ri][7] || '').trim() === 'פסול') dqResults++;
        else otherResults++;
      }
    }
    // Count terminal entries by type in pending sheet for this examinee
    var dqTerminals = 0, otherTerminals = 0;
    for (var cc = 1; cc < pendData.length; cc++) {
      if (String(pendData[cc][0]) !== code || normalizeId(pendData[cc][1]) !== normalizeId(ciId)) continue;
      var ccStatus = String(pendData[cc][5]).trim();
      if (ccStatus === 'disqualified' || ccStatus === 'dq_confirmed') dqTerminals++;
      else if (ccStatus === 'completed') otherTerminals++;
    }
    // Cap DQ results to DQ terminals — handles duplicate פסול rows from page refreshes
    var effectiveResults = Math.min(dqResults, dqTerminals) + otherResults;
    var totalTerminals = dqTerminals + otherTerminals;
    var hasUnmatchedResult = effectiveResults > totalTerminals;

    if (hasUnmatchedResult || isStale) {
      // Fix dangling status — mark as completed
      pendSheet.getRange(ci + 1, 6).setValue('completed');
      pendData[ci][5] = 'completed'; // update local copy
      if (isStale && !hasUnmatchedResult) {
        // Create a timeout fail result
        var sesData2 = getSheet('סשנים').getDataRange().getValues();
        var license2 = pendData[ci][8] || '', site2 = '', classroom2 = '', examinerName2 = '', language2 = pendData[ci][6] || 'he';
        for (var si = 1; si < sesData2.length; si++) {
          if (String(sesData2[si][0]).trim() === code) {
            examinerName2 = sesData2[si][2] || '';
            site2 = sesData2[si][3] || '';
            classroom2 = sesData2[si][4] || '';
            if (!license2) license2 = sesData2[si][5] || '';
            break;
          }
        }
        var attemptNum2 = countAttempts(String(ciId), license2) + 1;
        resSheet.appendRow([
          todayStr(), ciId, pendData[ci][2] || '', pendData[ci][3] || '', license2,
          '0/30', '0%', 'נכשל', '', examinerName2,
          site2, classroom2, language2, code,
          attemptNum2, 'ניתוק/טיימאאוט — הנבחן לא סיים את המבחן', false, false, '',
          pendData[ci][7] || '', false, pendData[ci][9] || 'off'
        ]);
        // Refresh resData after append
        resData = resSheet.getDataRange().getValues();
      }
    }
  }

  for (var i = 1; i < pendData.length; i++) {
    if (String(pendData[i][0]) !== code) continue;
    var s = String(pendData[i][5] || '').trim();
    var item = { idNumber: pendData[i][1], name: pendData[i][2], phone: pendData[i][3], time: pendData[i][4], examStartTime: pendData[i][11] || '', status: s, language: pendData[i][6] || '', population: pendData[i][7] || '', license: pendData[i][8] || '', audioMode: pendData[i][9] || 'off', timeExtension: String(pendData[i][10] || '') };
    if (s === 'waiting') pending.push(item);
    else if (s === 'approved') pending.push(item);
    else if (s === 'in_exam') active.push(item);
  }

  // Re-read resData in case cleanup added new results
  resData = resSheet.getDataRange().getValues();
  var completed = [];
  for (var j = 1; j < resData.length; j++) {
    if (String(resData[j][13]) !== code) continue;
    if (String(resData[j][7] || '') === 'בוטל') continue;
    completed.push({
      date: resData[j][0],
      idNumber: resData[j][1],
      name: resData[j][2],
      phone: resData[j][3],
      license: resData[j][4],
      score: resData[j][5],
      percent: resData[j][6],
      passed: resData[j][7],
      time: resData[j][8],
      examiner: resData[j][9],
      site: resData[j][10],
      classroom: resData[j][11],
      language: resData[j][12],
      attempt: resData[j][14],
      wrongDetails: resData[j][15],
      sent: resData[j][16],
      disqualified: resData[j][17],
      waLink: resData[j][18],
      population: resData[j][19] || '',
      corrected: resData[j][20] || false,
      audioMode: resData[j][21] || 'off'
    });
  }

  // Flag repeat examinees: check if any pending examinee already tested today (any session)
  var now = new Date();
  var todayDD = ('0' + now.getDate()).slice(-2);
  var todayMM = ('0' + (now.getMonth() + 1)).slice(-2);
  var todayYYYY = now.getFullYear();
  var todayDate = todayDD + '/' + todayMM + '/' + todayYYYY; // "DD/MM/YYYY"
  for (var pi = 0; pi < pending.length; pi++) {
    var todayExams = [];
    for (var ri = 1; ri < resData.length; ri++) {
      if (normalizeId(resData[ri][1]) !== normalizeId(pending[pi].idNumber)) continue;
      if (String(resData[ri][7] || '') === 'בוטל') continue;
      // Handle both Date objects and string dates from Sheets
      var cellDate = resData[ri][0];
      var dateStr = '';
      if (cellDate instanceof Date) {
        dateStr = ('0' + cellDate.getDate()).slice(-2) + '/' + ('0' + (cellDate.getMonth() + 1)).slice(-2) + '/' + cellDate.getFullYear();
      } else {
        dateStr = String(cellDate);
      }
      if (dateStr.indexOf(todayDate) === 0) {
        todayExams.push({ license: String(resData[ri][4]), score: String(resData[ri][5]), passed: String(resData[ri][7]), language: String(resData[ri][12] || '') });
      }
    }
    if (todayExams.length > 0) {
      pending[pi].todayExams = todayExams;
    }
  }

  // Cross-reference registration times from ממתינים for completed results
  for (var c = 0; c < completed.length; c++) {
    for (var p2 = pendData.length - 1; p2 >= 1; p2--) {
      if (String(pendData[p2][0]) === code && normalizeId(pendData[p2][1]) === normalizeId(completed[c].idNumber)) {
        completed[c].registrationTime = pendData[p2][4];
        break;
      }
    }
  }

  return jsonResponse({ status: 'ok', pending: pending, active: active, completed: completed });
}

function handleDisqualify(p) {
  // Verify examiner if provided (examinee-side DQ doesn't send examinerId)
  if (p.examinerId && !verifyExaminerForSession(p.sessionCode, p.examinerId)) {
    return jsonResponse({ status: 'error', message: 'אין הרשאה — בוחן לא תואם לסשן' });
  }
  // ALWAYS remove from ממתינים first (so examinee leaves "active" list)
  var pendSheet = getSheet('ממתינים');
  var pendData = pendSheet.getDataRange().getValues();
  var name = '', phone = '', population = '', examineeLicense = '', examineeAudio = 'off';
  for (var j = pendData.length - 1; j >= 1; j--) {
    if (String(pendData[j][0]) === String(p.sessionCode) && normalizeId(pendData[j][1]) === normalizeId(p.idNumber)) {
      name = pendData[j][2] || '';
      phone = pendData[j][3] || '';
      population = pendData[j][7] || '';
      examineeLicense = pendData[j][8] || '';
      examineeAudio = pendData[j][9] || 'off';
      pendSheet.getRange(j + 1, 6).setValue('disqualified');
      break;
    }
  }

  // Idempotency: if the latest result is already "פסול" with the SAME dqEventId, this is a retry — skip.
  // Different dqEventId or no dqEventId = new DQ event → create new row.
  var dqEventId = String(p.dqEventId || '');
  var sheet = getSheet('תוצאות');
  var data = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][13]) === String(p.sessionCode) && normalizeId(data[i][1]) === normalizeId(p.idNumber)) {
      var rowStatus = String(data[i][7]).trim();
      if ((rowStatus === 'פסול' || rowStatus === 'בוטל') && dqEventId && String(data[i][24] || '') === dqEventId) {
        // Same DQ event (active or cancelled) — retry from sendDQToServer, skip
        return jsonResponse({ status: 'ok' });
      }
      // Either latest is not "פסול", or different/missing dqEventId → new attempt, create new row
      break;
    }
  }

  // Create new disqualified result row
  var sesSheet = getSheet('סשנים');
  var sesData = sesSheet.getDataRange().getValues();
  var license = '', language = 'he', site = '', classroom = '', examinerName = '';
  for (var s = 1; s < sesData.length; s++) {
    if (String(sesData[s][0]).trim() === String(p.sessionCode).trim()) {
      examinerName = sesData[s][2] || '';
      site = sesData[s][3] || '';
      classroom = sesData[s][4] || '';
      license = examineeLicense || sesData[s][5] || '';
      language = sesData[s][6] || 'he';
      break;
    }
  }
  if (!license) license = examineeLicense;
  var attemptNum = countAttempts(String(p.idNumber), license) + 1;
  sheet.appendRow([
    todayStr(), p.idNumber, name, phone, license,
    '0/30', '0%', 'פסול', '', examinerName,
    site, classroom, language, String(p.sessionCode),
    attemptNum, '', false, true, '',
    population, false, examineeAudio, '', '', dqEventId
  ]);
  SpreadsheetApp.flush();
  return jsonResponse({ status: 'ok' });
}

// Cancel a provisional disqualification — called when examinee returns within grace period
function handleCancelDisqualify(p) {
  var sc = String(p.sessionCode || '');
  var id = normalizeId(p.idNumber || '');
  if (!sc || !id) return jsonResponse({ status: 'ok' });

  // 1. Revert pending status from 'disqualified' back to 'in_exam'
  var pendSheet = getSheet('ממתינים');
  var pendData = pendSheet.getDataRange().getValues();
  for (var j = pendData.length - 1; j >= 1; j--) {
    if (String(pendData[j][0]) === sc && normalizeId(pendData[j][1]) === id) {
      if (String(pendData[j][5]).trim() === 'disqualified') {
        pendSheet.getRange(j + 1, 6).setValue('in_exam');
        break;
      }
    }
  }

  // 2. Delete the DQ result row matching this dqEventId (or latest פסול if no eventId)
  var dqEventId = String(p.dqEventId || '');
  var resSheet = getSheet('תוצאות');
  var resData = resSheet.getDataRange().getValues();
  for (var i = resData.length - 1; i >= 1; i--) {
    if (String(resData[i][13]) === sc && normalizeId(resData[i][1]) === id) {
      if (String(resData[i][7]).trim() === 'פסול') {
        // Only delete if dqEventId matches (or if no eventId provided for backwards compat)
        if (!dqEventId || String(resData[i][24] || '') === dqEventId) {
          resSheet.getRange(i + 1, 8).setValue('בוטל');
          break;
        }
      }
      // If latest result is NOT פסול or eventId doesn't match — stop
      break;
    }
  }

  SpreadsheetApp.flush();
  return jsonResponse({ status: 'ok' });
}

function handleResetExaminee(p) {
  if (p.examinerId && !verifyExaminerForSession(p.sessionCode, p.examinerId)) {
    return jsonResponse({ status: 'error', message: 'אין הרשאה — בוחן לא תואם לסשן' });
  }
  var sheet = getSheet('ממתינים');
  var data = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === String(p.sessionCode) && normalizeId(data[i][1]) === normalizeId(p.idNumber)) {
      var s = String(data[i][5]).trim();
      if (s === 'in_exam' || s === 'approved' || s === 'waiting') {
        sheet.getRange(i + 1, 6).setValue('cancelled');
        SpreadsheetApp.flush();
        return jsonResponse({ status: 'ok' });
      }
    }
  }
  return jsonResponse({ status: 'error', message: 'לא נמצא נבחן פעיל לאיפוס' });
}

// Force-complete a stuck in_exam examinee (examiner manual action)
function handleForceComplete(p) {
  if (p.examinerId && !verifyExaminerForSession(p.sessionCode, p.examinerId)) {
    return jsonResponse({ status: 'error', message: 'אין הרשאה — בוחן לא תואם לסשן' });
  }
  var pendSheet = getSheet('ממתינים');
  var pendData = pendSheet.getDataRange().getValues();
  var found = false;
  var name = '', phone = '', population = '', examineeLicense = '', examineeAudio = 'off', language = 'he';
  for (var j = pendData.length - 1; j >= 1; j--) {
    if (String(pendData[j][0]) === String(p.sessionCode) && normalizeId(pendData[j][1]) === normalizeId(p.idNumber) && String(pendData[j][5]).trim() === 'in_exam') {
      name = pendData[j][2] || '';
      phone = pendData[j][3] || '';
      language = pendData[j][6] || 'he';
      population = pendData[j][7] || '';
      examineeLicense = pendData[j][8] || '';
      examineeAudio = pendData[j][9] || 'off';
      pendSheet.getRange(j + 1, 6).setValue('completed');
      found = true;
      break;
    }
  }
  if (!found) {
    return jsonResponse({ status: 'error', message: 'לא נמצא נבחן עם סטטוס in_exam' });
  }

  // Check if result already exists — if so, just mark pending as completed (done above)
  var resSheet = getSheet('תוצאות');
  var resData = resSheet.getDataRange().getValues();
  for (var i = resData.length - 1; i >= 1; i--) {
    if (String(resData[i][13]) === String(p.sessionCode) && normalizeId(resData[i][1]) === normalizeId(p.idNumber)) {
      SpreadsheetApp.flush();
      return jsonResponse({ status: 'ok', message: 'נמצאה תוצאה קיימת — הסטטוס עודכן' });
    }
  }

  // No result exists — create a fail result
  var sesSheet = getSheet('סשנים');
  var sesData = sesSheet.getDataRange().getValues();
  var license = examineeLicense, site = '', classroom = '', examinerName = '';
  for (var s = 1; s < sesData.length; s++) {
    if (String(sesData[s][0]).trim() === String(p.sessionCode).trim()) {
      examinerName = sesData[s][2] || '';
      site = sesData[s][3] || '';
      classroom = sesData[s][4] || '';
      if (!license) license = sesData[s][5] || '';
      break;
    }
  }
  var attemptNum = countAttempts(String(p.idNumber), license) + 1;
  resSheet.appendRow([
    todayStr(), p.idNumber, name, phone, license,
    '0/30', '0%', 'נכשל', '', examinerName,
    site, classroom, language, String(p.sessionCode),
    attemptNum, 'סיום ידני ע"י בוחן — ניתוק/תקלה', false, false, '',
    population, false, examineeAudio
  ]);
  SpreadsheetApp.flush();
  return jsonResponse({ status: 'ok', message: 'נבחן סומן כנכשל (ניתוק)' });
}

function handleOverturnDQ(p) {
  if (p.examinerId && !verifyExaminerForSession(p.sessionCode, p.examinerId)) {
    return jsonResponse({ status: 'error', message: 'אין הרשאה — בוחן לא תואם לסשן' });
  }
  var sheet = getSheet('תוצאות');
  var data = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][13]) === String(p.sessionCode) && normalizeId(data[i][1]) === normalizeId(p.idNumber)) {
      var status = String(data[i][7]).trim();
      if (status !== 'פסול') {
        return jsonResponse({ status: 'error', message: 'תוצאה זו אינה פסולה' });
      }
      // Change status from פסול to בוטל (hidden from results, allow retake)
      sheet.getRange(i + 1, 8).setValue('בוטל');
      // Clear the DQ flag
      sheet.getRange(i + 1, 18).setValue(false);
      // Revert pending status from 'disqualified' back to 'in_exam' so examinee can resume
      var pendSheet = getSheet('ממתינים');
      var pendData = pendSheet.getDataRange().getValues();
      for (var j = pendData.length - 1; j >= 1; j--) {
        if (String(pendData[j][0]) === String(p.sessionCode) && normalizeId(pendData[j][1]) === normalizeId(p.idNumber)) {
          if (String(pendData[j][5]).trim() === 'disqualified') {
            pendSheet.getRange(j + 1, 6).setValue('in_exam');
          }
          break;
        }
      }
      SpreadsheetApp.flush();
      return jsonResponse({ status: 'ok' });
    }
  }
  return jsonResponse({ status: 'error', message: 'תוצאה לא נמצאה' });
}

function handleConfirmDQ(p) {
  if (p.examinerId && !verifyExaminerForSession(p.sessionCode, p.examinerId)) {
    return jsonResponse({ status: 'error', message: 'אין הרשאה — בוחן לא תואם לסשן' });
  }
  // Mark pending status as dq_confirmed so examinee polling gets a final answer
  var pendSheet = getSheet('ממתינים');
  var pendData = pendSheet.getDataRange().getValues();
  for (var j = pendData.length - 1; j >= 1; j--) {
    if (String(pendData[j][0]) === String(p.sessionCode) && normalizeId(pendData[j][1]) === normalizeId(p.idNumber)) {
      if (String(pendData[j][5]).trim() === 'disqualified') {
        pendSheet.getRange(j + 1, 6).setValue('dq_confirmed');
        SpreadsheetApp.flush();
        return jsonResponse({ status: 'ok' });
      }
      break;
    }
  }
  return jsonResponse({ status: 'error', message: 'לא נמצא רישום פסול לאישור' });
}

function handleCorrectToPass(p) {
  if (p.examinerId && !verifyExaminerForSession(p.sessionCode, p.examinerId)) {
    return jsonResponse({ status: 'error', message: 'אין הרשאה — בוחן לא תואם לסשן' });
  }
  var sheet = getSheet('תוצאות');
  var data = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][13]) === String(p.sessionCode) && normalizeId(data[i][1]) === normalizeId(p.idNumber)) {
      // Verify score is eligible (>= 24/30)
      var scoreParts = String(data[i][5]).split('/');
      var scoreNum = parseInt(scoreParts[0]) || 0;
      if (scoreNum < 24) {
        return jsonResponse({ status: 'error', message: 'ציון נמוך מדי לתיקון (מתחת ל-24)' });
      }
      // Update pass/fail to עבר
      sheet.getRange(i + 1, 8).setValue('עבר');     // column H = עבר/נכשל
      // Clear disqualified flag (in case correcting a DQ result directly)
      sheet.getRange(i + 1, 18).setValue(false);     // column R = disqualified
      // Mark as corrected
      sheet.getRange(i + 1, 21).setValue(true);      // column U = תוקן?
      // Regenerate WhatsApp link — corrected result shows only "עבר" (no score/errors)
      var phone = formatPhoneForWA(data[i][3]);
      var waMsg = '*🚗 אישור תוצאת מבחן תאוריה חיצוני*\n\n' +
        'שם: ' + data[i][2] + '\n' +
        'ת.ז.: ' + data[i][1] + '\n' +
        'דרגה: ' + data[i][4] + '\n' +
        (data[i][19] ? 'אוכלוסיה: ' + data[i][19] + '\n' : '') +
        'תאריך: ' + data[i][0] + '\n' +
        'תוצאה: *עבר*\n';
      var waLink = 'https://wa.me/' + phone + '?text=' + encodeURIComponent(waMsg);
      sheet.getRange(i + 1, 19).setValue(waLink);    // column S = קישור וואטסאפ
      SpreadsheetApp.flush();
      return jsonResponse({ status: 'ok' });
    }
  }
  return jsonResponse({ status: 'error', message: 'תוצאה לא נמצאה' });
}

function handleMarkSent(p) {
  var sheet = getSheet('תוצאות');
  var data = sheet.getDataRange().getValues();
  var ids = p.idNumbers ? p.idNumbers.split(',') : [p.idNumber];
  var count = 0;
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][13]) === String(p.sessionCode)) {
      for (var k = 0; k < ids.length; k++) {
        if (normalizeId(data[i][1]) === normalizeId(ids[k])) {
          sheet.getRange(i + 1, 17).setValue(true);  // נשלח? — column Q (17)
          count++;
        }
      }
    }
  }
  return jsonResponse({ status: 'ok', updated: count });
}

function handleRegisterExamQuestions(data) {
  // Store question map for server-side score verification
  // data: { sessionCode, idNumber, questions: [{qIdx, correctShuffledIdx}] }
  if (!data.sessionCode || !data.idNumber || !data.questions) {
    return jsonResponse({ status: 'error', message: 'חסרים נתונים לרישום מבחן' });
  }
  // Verify examinee is in_exam status
  var pendSheet = getSheet('ממתינים');
  var pendData = pendSheet.getDataRange().getValues();
  var found = false;
  for (var i = pendData.length - 1; i >= 1; i--) {
    if (String(pendData[i][0]) === String(data.sessionCode) && normalizeId(pendData[i][1]) === normalizeId(data.idNumber)) {
      var status = String(pendData[i][5]).trim();
      if (status === 'in_exam' || status === 'approved') {
        found = true;
        break;
      }
    }
  }
  if (!found) {
    return jsonResponse({ status: 'error', message: 'נבחן לא מאושר למבחן' });
  }
  // Store in מבחנים sheet (create if needed)
  var examSheet;
  try { examSheet = getSheet('מבחנים'); } catch(e) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    examSheet = ss.insertSheet('מבחנים');
    examSheet.appendRow(['קוד סשן', 'ת.ז.', 'שאלות JSON', 'זמן רישום']);
  }
  examSheet.appendRow([
    String(data.sessionCode),
    normalizeId(data.idNumber),
    JSON.stringify(data.questions),
    nowISO()
  ]);
  return jsonResponse({ status: 'ok' });
}

function handleSubmitResult(data) {
  var sheet = getSheet('תוצאות');

  // Verify examinee is approved (in_exam status) before accepting results
  if (data.sessionCode && data.idNumber) {
    var pendSheet = getSheet('ממתינים');
    var pendData = pendSheet.getDataRange().getValues();
    var isApproved = false;
    for (var pi = pendData.length - 1; pi >= 1; pi--) {
      if (String(pendData[pi][0]) === String(data.sessionCode) && normalizeId(pendData[pi][1]) === normalizeId(data.idNumber)) {
        var pStatus = String(pendData[pi][5]).trim();
        if (pStatus === 'in_exam' || pStatus === 'approved' || pStatus === 'completed') {
          isApproved = true;
        }
        break;
      }
    }
    if (!isApproved) {
      return jsonResponse({ status: 'error', message: 'נבחן לא מאושר — לא ניתן לשלוח תוצאות' });
    }
  }

  // Server-side score verification: if answers array is present, recalculate score
  if (data.answers && Array.isArray(data.answers)) {
    try {
      var examSheet = getSheet('מבחנים');
      var examData = examSheet.getDataRange().getValues();
      var questionMap = null;
      // Find the registered exam for this session+ID (latest)
      for (var ei = examData.length - 1; ei >= 1; ei--) {
        if (String(examData[ei][0]) === String(data.sessionCode) && normalizeId(examData[ei][1]) === normalizeId(data.idNumber)) {
          questionMap = JSON.parse(examData[ei][2]);
          break;
        }
      }
      if (questionMap) {
        var correctCount = 0;
        var totalQ = questionMap.length;
        for (var ai = 0; ai < data.answers.length && ai < totalQ; ai++) {
          if (data.answers[ai] !== null && data.answers[ai] !== undefined) {
            var selected = Number(data.answers[ai].selected);
            var correctIdx = Number(questionMap[ai].correctShuffledIdx);
            if (selected === correctIdx) correctCount++;
          }
        }
        var pct = Math.round((correctCount / totalQ) * 100);
        var passThreshold = Math.ceil(totalQ * 0.86); // ~26/30
        data.score = correctCount;
        data.total = totalQ;
        data.percent = pct;
        data.passed = correctCount >= passThreshold;
        data.verified = true;
      }
    } catch(ve) {
      // If verification fails, fall through to client-provided score with flag
      data.verified = false;
    }
  }

  // Server-side timing check: if exam took less than 3 minutes, flag as suspicious
  if (data.sessionCode && data.idNumber) {
    try {
      var examSheet2 = getSheet('מבחנים');
      var examData2 = examSheet2.getDataRange().getValues();
      for (var ti = examData2.length - 1; ti >= 1; ti--) {
        if (String(examData2[ti][0]) === String(data.sessionCode) && normalizeId(examData2[ti][1]) === normalizeId(data.idNumber)) {
          var regTime = new Date(examData2[ti][3]);
          var elapsed = (new Date() - regTime) / 1000; // seconds
          if (elapsed < 180 && elapsed > 0) { // less than 3 minutes
            data.suspicious = true;
          }
          break;
        }
      }
    } catch(te) {}
  }

  // Duplicate protection: check if result already exists for this session+ID+license+language
  // Skip disqualified (פסול) and cancelled (בוטל) rows — those are not real results and should not block retakes
  // Also skip duplicate check entirely if examinee has an active in_exam pending row (retake after DQ)
  var hasPendingInExam = false;
  var pendCheck = getSheet('ממתינים').getDataRange().getValues();
  for (var pc = pendCheck.length - 1; pc >= 1; pc--) {
    if (String(pendCheck[pc][0]) === String(data.sessionCode) && normalizeId(pendCheck[pc][1]) === normalizeId(data.idNumber) && String(pendCheck[pc][5]).trim() === 'in_exam') {
      hasPendingInExam = true;
      break;
    }
  }
  if (!hasPendingInExam) {
    var existingData = sheet.getDataRange().getValues();
    for (var d = 1; d < existingData.length; d++) {
      var existingStatus = String(existingData[d][7] || '').trim();
      if (existingStatus === 'פסול' || existingStatus === 'בוטל') continue;
      if (String(existingData[d][13]) === String(data.sessionCode) && normalizeId(existingData[d][1]) === normalizeId(data.idNumber) && String(existingData[d][4]) === String(data.license) && String(existingData[d][12]) === String(data.language || 'he')) {
        // Already submitted — still update pending status so examinee doesn't stay stuck in "in_exam"
        markPendingCompleted(data.sessionCode, data.idNumber);
        return jsonResponse({ status: 'ok', waLink: existingData[d][18] || '', duplicate: true });
      }
    }
  }

  var wrongDetails = '';
  var wrongForWA = '';
  if (data.wrongAnswers && data.wrongAnswers.length > 0) {
    for (var i = 0; i < data.wrongAnswers.length; i++) {
      var w = data.wrongAnswers[i];
      wrongDetails += 'שאלה: ' + w.question + '\n';
      wrongDetails += 'תשובת הנבחן: ' + w.yourAnswer + '\n';
      wrongDetails += 'תשובה נכונה: ' + w.correctAnswer + '\n';
      if (w.category) wrongDetails += 'קטגוריה: ' + w.category + '\n';
      wrongDetails += '\n';

      wrongForWA += '❌ ' + w.question + '\n';
      wrongForWA += 'ענית: ' + w.yourAnswer + '\n';
      wrongForWA += '✅ נכון: ' + w.correctAnswer + '\n\n';
    }
  }

  var passText = data.passed ? 'עבר' : 'נכשל';
  var waMessage = '*🚗 אישור תוצאת מבחן תאוריה חיצוני*\n\n' +
    'שם: ' + data.fullName + '\n' +
    'ת.ז.: ' + data.idNumber + '\n' +
    'דרגה: ' + data.license + '\n' +
    (data.population ? 'אוכלוסיה: ' + data.population + '\n' : '') +
    'תאריך: ' + todayStr() + '\n' +
    'תוצאה: *' + passText + '* (' + data.score + '/' + data.total + ')\n' +
    'זמן: ' + data.time + '\n';

  var wrongCount = Number(data.total) - Number(data.score);
  if (data.wrongAnswers && data.wrongAnswers.length > 0) {
    waMessage += '\n*שאלות שגויות (' + data.wrongAnswers.length + '):*\n\n' + wrongForWA;
  } else if (wrongCount === 0) {
    waMessage += '\nכל התשובות נכונות! 🎉';
  }

  var phone = formatPhoneForWA(data.phone);

  // Count attempt number for this examinee + license combination
  var attemptNum = countAttempts(data.idNumber, data.license) + 1;

  var waMessage2 = waMessage; // preserve for link
  if (attemptNum > 1) {
    waMessage2 = waMessage + 'ניסיון: ' + attemptNum + '\n';
  }
  var waLink = 'https://wa.me/' + phone + '?text=' + encodeURIComponent(waMessage2);

  sheet.appendRow([
    todayStr(),
    data.idNumber,
    data.fullName,
    data.phone,
    data.license,
    data.score + '/' + data.total,
    data.percent + '%',
    passText,
    data.time,
    data.examinerName || '',
    data.site || '',
    data.classroom || '',
    data.language || 'he',
    data.sessionCode || '',
    attemptNum,
    wrongDetails,
    false,
    false,
    waLink,
    data.population || '',
    false,
    data.audioMode || 'off',
    data.verified ? 'מאומת' : '',
    data.suspicious ? 'חשוד' : ''
  ]);

  // Update pending status to completed
  markPendingCompleted(data.sessionCode, data.idNumber);

  return jsonResponse({ status: 'ok', waLink: waLink });
}

// Helper: mark the latest pending row for this session+ID as completed
function markPendingCompleted(sessionCode, idNumber) {
  var pendSheet = getSheet('ממתינים');
  var pendData = pendSheet.getDataRange().getValues();
  for (var j = pendData.length - 1; j >= 1; j--) {
    if (String(pendData[j][0]) === String(sessionCode) && normalizeId(pendData[j][1]) === normalizeId(idNumber) && (String(pendData[j][5]).trim() === 'in_exam' || String(pendData[j][5]).trim() === 'approved')) {
      pendSheet.getRange(j + 1, 6).setValue('completed');
      break;
    }
  }
}

function handleSubmitWrongAnswers(p) {
  // Append a single wrong answer item to existing result row
  var sheet = getSheet('תוצאות');
  var data = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][13]) === String(p.sessionCode) && normalizeId(data[i][1]) === normalizeId(p.idNumber)) {
      var existing = String(data[i][15] || '');

      // New format: individual item with question/yourAnswer/correctAnswer params
      if (p.question) {
        var line = 'שאלה: ' + p.question + '\n' +
                   'תשובת הנבחן: ' + p.yourAnswer + '\n' +
                   'תשובה נכונה: ' + p.correctAnswer + '\n\n';
        sheet.getRange(i + 1, 16).setValue(existing + line);
        SpreadsheetApp.flush();
        return jsonResponse({ status: 'ok' });
      }

      // Legacy format: chunk with JSON array
      var chunk = p.chunk || '';
      var totalChunks = Number(p.totalChunks) || 1;
      if (totalChunks === 1) {
        try {
          var wrongArr = JSON.parse(chunk);
          var formatted = '';
          for (var w = 0; w < wrongArr.length; w++) {
            formatted += 'שאלה: ' + wrongArr[w].question + '\n';
            formatted += 'תשובת הנבחן: ' + wrongArr[w].yourAnswer + '\n';
            formatted += 'תשובה נכונה: ' + wrongArr[w].correctAnswer + '\n\n';
          }
          sheet.getRange(i + 1, 16).setValue(formatted);
        } catch(ex) {
          sheet.getRange(i + 1, 16).setValue(existing + chunk);
        }
      } else {
        sheet.getRange(i + 1, 16).setValue(existing + chunk);
      }
      return jsonResponse({ status: 'ok' });
    }
  }
  return jsonResponse({ status: 'error', message: 'Result row not found for wrong answers' });
}

function handleSubmitWrongAnswersBulk(data) {
  // Receive ALL wrong answers in a single POST and write to result row
  var sheet = getSheet('תוצאות');
  var rows = sheet.getDataRange().getValues();
  for (var i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][13]) === String(data.sessionCode) && normalizeId(rows[i][1]) === normalizeId(data.idNumber)) {
      var wrongDetails = '';
      var wrongForWA = '';
      if (data.wrongAnswers && data.wrongAnswers.length > 0) {
        for (var w = 0; w < data.wrongAnswers.length; w++) {
          var item = data.wrongAnswers[w];
          wrongDetails += 'שאלה: ' + item.question + '\n';
          wrongDetails += 'תשובת הנבחן: ' + item.yourAnswer + '\n';
          wrongDetails += 'תשובה נכונה: ' + item.correctAnswer + '\n';
          if (item.category) wrongDetails += 'קטגוריה: ' + item.category + '\n';
          wrongDetails += '\n';
          wrongForWA += '❌ ' + item.question + '\n';
          wrongForWA += 'ענית: ' + item.yourAnswer + '\n';
          wrongForWA += '✅ נכון: ' + item.correctAnswer + '\n\n';
        }
      }
      // Update wrong details column
      sheet.getRange(i + 1, 16).setValue(wrongDetails);

      // Regenerate WA link with wrong answers included
      var isCorrected = rows[i][20] === true || String(rows[i][20]) === 'TRUE';
      if (!isCorrected && data.wrongAnswers && data.wrongAnswers.length > 0) {
        var passText = String(rows[i][7] || 'נכשל');
        var phone = formatPhoneForWA(rows[i][3]);
        var waMsg = '*🚗 אישור תוצאת מבחן תאוריה חיצוני*\n\n' +
          'שם: ' + rows[i][2] + '\n' +
          'ת.ז.: ' + rows[i][1] + '\n' +
          'דרגה: ' + rows[i][4] + '\n' +
          (rows[i][19] ? 'אוכלוסיה: ' + rows[i][19] + '\n' : '') +
          'תאריך: ' + rows[i][0] + '\n' +
          'תוצאה: *' + passText + '* (' + rows[i][5] + ')\n' +
          'זמן: ' + rows[i][8] + '\n' +
          '\n*שאלות שגויות (' + data.wrongAnswers.length + '):*\n\n' + wrongForWA;
        var attemptNum = rows[i][14] || 1;
        if (attemptNum > 1) waMsg += 'ניסיון: ' + attemptNum + '\n';
        var waLink = 'https://wa.me/' + phone + '?text=' + encodeURIComponent(waMsg);
        sheet.getRange(i + 1, 19).setValue(waLink);
      }

      SpreadsheetApp.flush();
      return jsonResponse({ status: 'ok', count: data.wrongAnswers ? data.wrongAnswers.length : 0 });
    }
  }
  return jsonResponse({ status: 'error', message: 'Result row not found for wrong answers' });
}

function handleSubmitFailOnClose(data) {
  var sheet = getSheet('תוצאות');

  // Duplicate protection: check if result already exists for this session+ID+license+language
  var existingData = sheet.getDataRange().getValues();
  for (var d = 1; d < existingData.length; d++) {
    if (String(existingData[d][13]) === String(data.sessionCode) && normalizeId(existingData[d][1]) === normalizeId(data.idNumber) && String(existingData[d][4]) === String(data.license) && String(existingData[d][12]) === String(data.language || 'he')) {
      markPendingCompleted(data.sessionCode, data.idNumber);
      return jsonResponse({ status: 'ok', duplicate: true });
    }
  }

  var attemptNum = countAttempts(data.idNumber, data.license || '') + 1;

  sheet.appendRow([
    todayStr(),
    data.idNumber,
    data.fullName,
    data.phone,
    data.license || '',
    '0/' + (data.totalQuestions || 30),
    '0%',
    'נכשל',
    data.time || '00:00',
    data.examinerName || '',
    data.site || '',
    data.classroom || '',
    data.language || 'he',
    data.sessionCode || '',
    attemptNum,
    'סגירת דפדפן באמצע מבחן (נענו ' + (data.answeredCount || 0) + ' שאלות)',
    false,
    false,
    '',
    data.population || '',
    false,
    data.audioMode || 'off'
  ]);

  markPendingCompleted(data.sessionCode, data.idNumber);

  return jsonResponse({ status: 'ok' });
}

function handleCancelFailOnClose(data) {
  // Called when page reloads (refresh, not actual close) — undo the fail
  var sc = String(data.sessionCode || '');
  var id = normalizeId(data.idNumber || '');
  if (!sc || !id) return jsonResponse({ status: 'ok' });

  var sheet = getSheet('תוצאות');
  var rows = sheet.getDataRange().getValues();
  // Find the most recent row for this session+ID that is a "close" fail
  for (var r = rows.length - 1; r >= 1; r--) {
    if (String(rows[r][13]) === sc && normalizeId(rows[r][1]) === id) {
      var notes = String(rows[r][15] || '');
      if (notes.indexOf('\u05E1\u05D2\u05D9\u05E8\u05EA \u05D3\u05E4\u05D3\u05E4\u05DF') !== -1) {
        // Delete this row — examinee resumed, so the fail was premature
        sheet.deleteRow(r + 1);
        // Also un-mark pending as completed so exam can continue
        unmarkPendingCompleted(sc, id);
      }
      break; // only check the most recent match
    }
  }
  return jsonResponse({ status: 'ok' });
}

function unmarkPendingCompleted(sessionCode, idNumber) {
  // Restore examinee status from 'done' back to 'in_exam' in pending sheet
  var pendingSheet = getSheet('ממתינים');
  if (!pendingSheet) return;
  var data = pendingSheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === sessionCode && normalizeId(data[i][1]) === normalizeId(idNumber)) {
      var status = String(data[i][6] || '');
      if (status === 'done') {
        pendingSheet.getRange(i + 1, 7).setValue('in_exam');
      }
      break;
    }
  }
}

// ========== שיתוף תוצאה דרך CacheService ==========

function handleUploadResultHtml(data) {
  if (!data.html || !data.requestId) {
    return jsonResponse({ status: 'error', message: 'Missing html or requestId' });
  }

  try {
    var cache = CacheService.getScriptCache();
    var html = data.html;
    var CHUNK_SIZE = 90000; // 90KB per chunk (limit is 100KB)
    var numChunks = Math.ceil(html.length / CHUNK_SIZE);

    // שומר HTML בחלקים ב-CacheService (עד 6 שעות)
    var chunks = {};
    for (var i = 0; i < numChunks; i++) {
      chunks['result_' + data.requestId + '_' + i] = html.substring(i * CHUNK_SIZE, (i + 1) * CHUNK_SIZE);
    }
    chunks['result_' + data.requestId + '_meta'] = String(numChunks);
    cache.putAll(chunks, 21600);

    // בונה קישור לצפייה דרך doGet
    var viewLink = ScriptApp.getService().getUrl() + '?action=viewResult&id=' + data.requestId;

    // שומר תוצאה ב-ScriptProperties כדי שה-client יוכל לקרוא דרך GET polling
    var props = PropertiesService.getScriptProperties();
    props.setProperty('upload_' + data.requestId, JSON.stringify({ link: viewLink }));

    return jsonResponse({ status: 'ok', link: viewLink });
  } catch (err) {
    if (data.requestId) {
      var props2 = PropertiesService.getScriptProperties();
      props2.setProperty('upload_' + data.requestId, JSON.stringify({ error: err.toString() }));
    }
    return jsonResponse({ status: 'error', message: err.toString() });
  }
}

function handleGetUploadResult(p) {
  var requestId = p.requestId;
  if (!requestId) return jsonResponse({ status: 'error', message: 'Missing requestId' });

  var props = PropertiesService.getScriptProperties();
  var stored = props.getProperty('upload_' + requestId);
  if (!stored) return jsonResponse({ status: 'pending' });

  // ניקוי
  props.deleteProperty('upload_' + requestId);

  var result = JSON.parse(stored);
  if (result.error) {
    return jsonResponse({ status: 'error', message: result.error });
  }
  return jsonResponse({ status: 'ok', link: result.link });
}

// ========== Commander Dashboard ==========

function parseDateParam(str) {
  if (!str) return null;
  var parts = String(str).split('/');
  if (parts.length !== 3) return null;
  var d = parseInt(parts[0], 10);
  var m = parseInt(parts[1], 10) - 1;
  var y = parseInt(parts[2], 10);
  if (isNaN(d) || isNaN(m) || isNaN(y)) return null;
  return new Date(y, m, d);
}

function parseSheetDate(val) {
  if (val instanceof Date) return val;
  var s = String(val || '');
  var match = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (match) return new Date(parseInt(match[3]), parseInt(match[2]) - 1, parseInt(match[1]));
  var d = new Date(s);
  if (!isNaN(d.getTime())) return d;
  return null;
}

function handleCommanderDashboard(p) {
  // Verify role
  var exSheet = getSheet('בוחנים');
  var exData = exSheet.getDataRange().getValues();
  var role = '';
  for (var i = 1; i < exData.length; i++) {
    if (normalizeId(exData[i][1]) === normalizeId(p.examinerId)) {
      role = String(exData[i][5] || 'בוחן');
      break;
    }
  }
  if (role !== 'מפקד') {
    return jsonResponse({ status: 'error', message: 'אין הרשאת מפקד' });
  }

  // Parse date range
  var dateFrom = parseDateParam(p.dateFrom);
  var dateTo = parseDateParam(p.dateTo);
  if (!dateFrom || !dateTo) {
    return jsonResponse({ status: 'error', message: 'תאריכים לא תקינים' });
  }
  dateTo.setHours(23, 59, 59, 999);

  // Read results
  var resSheet = getSheet('תוצאות');
  var resData = resSheet.getDataRange().getValues();

  // Aggregate
  var overall = { total: 0, passed: 0, failed: 0, disqualified: 0, scores: [] };
  var byExaminer = {};
  var bySite = {};
  var byLicense = {};
  var byPopulation = {};

  for (var r = 1; r < resData.length; r++) {
    var rowDate = parseSheetDate(resData[r][0]);
    if (!rowDate || rowDate < dateFrom || rowDate > dateTo) continue;

    var examinerName = String(resData[r][9] || '');
    var siteName = String(resData[r][10] || '');
    var license = String(resData[r][4] || '');
    var population = String(resData[r][19] || '');
    var passedStr = String(resData[r][7] || '');
    if (passedStr === 'בוטל') continue;
    var isDQ = resData[r][17] === true || String(resData[r][17]).toUpperCase() === 'TRUE' || passedStr === 'פסול';
    var isPassed = !isDQ && (passedStr === 'עבר');

    var pctVal = 0;
    var pctStr = String(resData[r][6] || '');
    if (pctStr.indexOf('%') !== -1) {
      pctVal = parseFloat(pctStr.replace('%', '')) || 0;
    } else {
      var pctNum = Number(resData[r][6]);
      if (!isNaN(pctNum)) {
        pctVal = pctNum <= 1 ? pctNum * 100 : pctNum;
      }
    }

    overall.total++;
    if (isDQ) overall.disqualified++;
    else if (isPassed) overall.passed++;
    else overall.failed++;
    overall.scores.push(pctVal);

    var eName = examinerName || 'לא צוין';
    var sName = siteName || 'לא צוין';
    var lName = license || 'לא צוין';
    var pName = population || 'לא צוין';

    addToGroup(byExaminer, eName, isPassed, isDQ, pctVal);
    addToGroup(bySite, sName, isPassed, isDQ, pctVal);
    addToGroup(byLicense, lName, isPassed, isDQ, pctVal);
    addToGroup(byPopulation, pName, isPassed, isDQ, pctVal);

    // Cross-tabulation sub-groups
    addToSubGroup(byExaminer, eName, 'byLicense', lName, isPassed, isDQ, pctVal);
    addToSubGroup(byExaminer, eName, 'bySite', sName, isPassed, isDQ, pctVal);
    addToSubGroup(bySite, sName, 'byLicense', lName, isPassed, isDQ, pctVal);
    addToSubGroup(bySite, sName, 'byExaminer', eName, isPassed, isDQ, pctVal);
    addToSubGroup(byLicense, lName, 'bySite', sName, isPassed, isDQ, pctVal);
    addToSubGroup(byLicense, lName, 'byExaminer', eName, isPassed, isDQ, pctVal);
    addToSubGroup(byPopulation, pName, 'byLicense', lName, isPassed, isDQ, pctVal);
    addToSubGroup(byPopulation, pName, 'bySite', sName, isPassed, isDQ, pctVal);
  }

  function addToGroup(map, key, isPassed, isDQ, pctVal) {
    if (!map[key]) map[key] = { total: 0, passed: 0, failed: 0, disqualified: 0, scores: [] };
    map[key].total++;
    if (isDQ) map[key].disqualified++;
    else if (isPassed) map[key].passed++;
    else map[key].failed++;
    map[key].scores.push(pctVal);
  }

  function addToSubGroup(map, primaryKey, subDim, subKey, isPassed, isDQ, pctVal) {
    if (!map[primaryKey]) return;
    if (!map[primaryKey]._sub) map[primaryKey]._sub = {};
    if (!map[primaryKey]._sub[subDim]) map[primaryKey]._sub[subDim] = {};
    addToGroup(map[primaryKey]._sub[subDim], subKey, isPassed, isDQ, pctVal);
  }

  function computeStats(obj) {
    var avg = 0, median = 0;
    if (obj.scores.length > 0) {
      var sum = 0;
      for (var s = 0; s < obj.scores.length; s++) sum += obj.scores[s];
      avg = Math.round(sum / obj.scores.length);
      var sorted = obj.scores.slice().sort(function(a, b) { return a - b; });
      var mid = Math.floor(sorted.length / 2);
      median = sorted.length % 2 !== 0 ? sorted[mid] : Math.round((sorted[mid - 1] + sorted[mid]) / 2);
    }
    var passRate = obj.total > 0 ? Math.round((obj.passed / obj.total) * 100) : 0;
    return { total: obj.total, passed: obj.passed, failed: obj.failed, disqualified: obj.disqualified, passRate: passRate, avgScore: avg, medianScore: median };
  }

  function computeGroupWithSub(map) {
    var out = {};
    for (var key in map) {
      out[key] = computeStats(map[key]);
      if (map[key]._sub) {
        out[key].sub = {};
        for (var subDim in map[key]._sub) {
          out[key].sub[subDim] = {};
          for (var subKey in map[key]._sub[subDim]) {
            out[key].sub[subDim][subKey] = computeStats(map[key]._sub[subDim][subKey]);
          }
        }
      }
    }
    return out;
  }

  var result = {
    overall: computeStats(overall),
    byExaminer: computeGroupWithSub(byExaminer),
    bySite: computeGroupWithSub(bySite),
    byLicense: computeGroupWithSub(byLicense),
    byPopulation: computeGroupWithSub(byPopulation)
  };

  return jsonResponse({ status: 'ok', data: result });
}

// ========== מערכת מורים — Teacher System ==========

function verifyTeacherToken(teacherId, token) {
  if (!teacherId || !token) return false;
  var sheet = getSheet('מורים');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (normalizeId(data[i][1]) === normalizeId(teacherId)) {
      var storedToken = String(data[i][4] || '');
      var expiry = data[i][5];
      if (storedToken === token && expiry) {
        var expiryDate = expiry instanceof Date ? expiry : new Date(expiry);
        if (new Date() <= expiryDate) return true;
      }
      // Continue searching other rows with same ID
    }
  }
  return false;
}

function requireTeacherToken(p) {
  if (!verifyTeacherToken(p.teacherId, p.token)) {
    return jsonResponse({ status: 'error', message: 'טוקן לא תקין — יש להתחבר מחדש', tokenExpired: true });
  }
  return null;
}

function generateClassCode() {
  var chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  var sheet = getSheet('כיתות');
  var data = sheet.getDataRange().getValues();
  var existing = {};
  for (var i = 1; i < data.length; i++) existing[String(data[i][0]).trim()] = true;
  var code;
  do {
    code = '';
    for (var c = 0; c < 6; c++) code += chars.charAt(Math.floor(Math.random() * chars.length));
  } while (existing[code]);
  return code;
}

function handleTeacherLogin(p) {
  var sheet = getSheet('מורים');
  var data = sheet.getDataRange().getValues();
  var matchedRows = [];
  for (var i = 1; i < data.length; i++) {
    if (normalizeId(data[i][1]) === normalizeId(p.idNumber)) {
      matchedRows.push(i);
    }
  }
  if (matchedRows.length === 0) {
    return jsonResponse({ status: 'error', message: 'מורה לא נמצא' });
  }
  var lastError = '';
  for (var m = 0; m < matchedRows.length; m++) {
    var i = matchedRows[m];
    var row = i + 1;
    var failedAttempts = Number(data[i][6]) || 0;
    var lockoutUntil = data[i][7];
    if (lockoutUntil) {
      var lockoutDate = lockoutUntil instanceof Date ? lockoutUntil : new Date(lockoutUntil);
      if (new Date() < lockoutDate) {
        var minsLeft = Math.ceil((lockoutDate - new Date()) / 60000);
        lastError = 'החשבון נעול. נסה שוב בעוד ' + minsLeft + ' דקות';
        continue;
      }
      failedAttempts = 0;
      sheet.getRange(row, 7).setValue(0);
      sheet.getRange(row, 8).setValue('');
    }
    if (String(data[i][2]) === String(p.password)) {
      if (data[i][3] === 'כן' || data[i][3] === true || data[i][3] === 'TRUE') {
        if (failedAttempts > 0) {
          sheet.getRange(row, 7).setValue(0);
          sheet.getRange(row, 8).setValue('');
        }
        var token = generateToken();
        var expiry = new Date();
        expiry.setHours(expiry.getHours() + 12);
        sheet.getRange(row, 5).setValue(token);
        sheet.getRange(row, 6).setValue(expiry);
        return jsonResponse({ status: 'ok', teacher: { name: data[i][0], id: normalizeId(data[i][1]), token: token, role: String(data[i][8] || 'מורה'), site: String(data[i][9] || '') } });
      } else {
        lastError = 'החשבון אינו פעיל';
        continue;
      }
    } else {
      lastError = 'סיסמה שגויה';
    }
  }
  // If no row matched successfully, increment failed attempts on first active row
  if (lastError === 'סיסמה שגויה' && matchedRows.length > 0) {
    var fi = matchedRows[0];
    var fRow = fi + 1;
    var fa = (Number(data[fi][6]) || 0) + 1;
    sheet.getRange(fRow, 7).setValue(fa);
    if (fa >= 5) {
      var lockout = new Date();
      lockout.setMinutes(lockout.getMinutes() + 15);
      sheet.getRange(fRow, 8).setValue(lockout);
      return jsonResponse({ status: 'error', message: 'יותר מדי ניסיונות. החשבון ננעל ל-15 דקות' });
    }
  }
  return jsonResponse({ status: 'error', message: lastError || 'שגיאה בהתחברות' });
}

function handleTeacherVerifyLogin(p) {
  if (!verifyTeacherToken(p.teacherId, p.token)) {
    return jsonResponse({ status: 'error', message: 'טוקן לא תקין', tokenExpired: true });
  }
  var sheet = getSheet('מורים');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (normalizeId(data[i][1]) === normalizeId(p.teacherId)) {
      return jsonResponse({ status: 'ok', teacher: { name: data[i][0], id: normalizeId(data[i][1]), role: String(data[i][8] || 'מורה'), site: String(data[i][9] || '') } });
    }
  }
  return jsonResponse({ status: 'error', message: 'מורה לא נמצא' });
}

function handleTeacherCommanderDashboard(p) {
  // Verify commander role
  var tSheet = getSheet('מורים');
  var tData = tSheet.getDataRange().getValues();
  var role = '';
  var userSite = '';
  for (var i = 1; i < tData.length; i++) {
    if (normalizeId(tData[i][1]) === normalizeId(p.teacherId)) {
      role = String(tData[i][8] || 'מורה');
      userSite = String(tData[i][9] || '');
      break;
    }
  }
  if (role !== 'מפקד' && role !== 'מפקד מקומי' && role !== 'מפקד ראשי') {
    return jsonResponse({ status: 'error', message: 'אין הרשאת מפקד' });
  }

  // Determine if local or global commander
  // Legacy: role === 'מפקד' treated as 'מפקד ראשי'
  var isGlobal = (role === 'מפקד ראשי' || role === 'מפקד');
  var isLocal = (role === 'מפקד מקומי');

  // Parse date range
  var dateFrom = parseDateParam(p.dateFrom);
  var dateTo = parseDateParam(p.dateTo);
  if (!dateFrom || !dateTo) {
    return jsonResponse({ status: 'error', message: 'תאריכים לא תקינים' });
  }
  dateTo.setHours(23, 59, 59, 999);

  // Build class→teacher map from כיתות sheet
  var classSheet = getSheet('כיתות');
  var classData = classSheet.getDataRange().getValues();
  var classMap = {}; // classCode → { teacherName, className, license, site }
  for (var c = 1; c < classData.length; c++) {
    var cc = String(classData[c][0]).trim();
    classMap[cc] = {
      teacherName: String(classData[c][3] || ''),
      teacherId: normalizeId(classData[c][2]),
      className: String(classData[c][1] || ''),
      license: String(classData[c][4] || ''),
      site: String(classData[c][7] || '')
    };
  }

  // Read practice results
  var resSheet = getSheet('תוצאות תרגול');
  var resData = resSheet.getDataRange().getValues();

  var overall = { total: 0, passed: 0, failed: 0, scores: [], students: {}, teachers: {}, classes: {}, sites: {} };
  var byTeacher = {};
  var byClass = {};
  var byLicense = {};
  var byMode = {};
  var bySite = {};

  for (var r = 1; r < resData.length; r++) {
    var rowDate = parseSheetDate(resData[r][0]);
    if (!rowDate || rowDate < dateFrom || rowDate > dateTo) continue;

    var classCode = String(resData[r][3] || '').trim();
    if (!classCode) continue;

    var cInfo = classMap[classCode] || { teacherName: 'לא ידוע', className: classCode, license: '', site: '' };
    var classSite = cInfo.site || '';

    // Site filtering for local commander
    if (isLocal && userSite && classSite !== userSite) continue;

    var teacherName = cInfo.teacherName || 'לא ידוע';
    var className = cInfo.className || classCode;
    var license = String(resData[r][5] || cInfo.license || 'לא צוין');
    var mode = String(resData[r][4] || 'לא צוין');
    var studentId = String(resData[r][1] || '');
    var passedStr = String(resData[r][9] || '');
    var isPassed = (passedStr === 'עבר' || passedStr === 'true' || passedStr === true);
    var isFailed = (passedStr === 'נכשל' || passedStr === 'false' || passedStr === false);

    var pctVal = 0;
    var pctRaw = resData[r][8];
    if (typeof pctRaw === 'string' && pctRaw.indexOf('%') !== -1) {
      pctVal = parseFloat(pctRaw.replace('%', '')) || 0;
    } else {
      var pctNum = Number(pctRaw);
      if (!isNaN(pctNum)) {
        pctVal = pctNum <= 1 ? pctNum * 100 : pctNum;
      }
    }

    overall.total++;
    if (isPassed) overall.passed++;
    else if (isFailed) overall.failed++;
    overall.scores.push(pctVal);
    overall.students[studentId] = true;
    overall.teachers[teacherName] = true;
    overall.classes[classCode] = true;
    if (classSite) overall.sites[classSite] = true;

    addToGroup(byTeacher, teacherName, isPassed, isFailed, pctVal, studentId);
    addToGroup(byClass, className + ' (' + classCode + ')', isPassed, isFailed, pctVal, studentId);
    addToGroup(byLicense, license, isPassed, isFailed, pctVal, studentId);
    addToGroup(byMode, mode, isPassed, isFailed, pctVal, studentId);

    // bySite aggregation (for global commander)
    if (isGlobal && classSite) {
      addToGroup(bySite, classSite, isPassed, isFailed, pctVal, studentId);
      addToSubGroup(bySite, classSite, 'byTeacher', teacherName, isPassed, isFailed, pctVal, studentId);
      addToSubGroup(bySite, classSite, 'byClass', className + ' (' + classCode + ')', isPassed, isFailed, pctVal, studentId);
    }

    // Cross-tabulation sub-groups
    addToSubGroup(byTeacher, teacherName, 'byClass', className + ' (' + classCode + ')', isPassed, isFailed, pctVal, studentId);
    addToSubGroup(byTeacher, teacherName, 'byLicense', license, isPassed, isFailed, pctVal, studentId);
    addToSubGroup(byClass, className + ' (' + classCode + ')', 'byLicense', license, isPassed, isFailed, pctVal, studentId);
    addToSubGroup(byClass, className + ' (' + classCode + ')', 'byMode', mode, isPassed, isFailed, pctVal, studentId);
    addToSubGroup(byLicense, license, 'byTeacher', teacherName, isPassed, isFailed, pctVal, studentId);
    addToSubGroup(byLicense, license, 'byClass', className + ' (' + classCode + ')', isPassed, isFailed, pctVal, studentId);
    addToSubGroup(byMode, mode, 'byLicense', license, isPassed, isFailed, pctVal, studentId);
    addToSubGroup(byMode, mode, 'byTeacher', teacherName, isPassed, isFailed, pctVal, studentId);
  }

  function addToGroup(map, key, isPassed, isFailed, pctVal, studentId) {
    if (!map[key]) map[key] = { total: 0, passed: 0, failed: 0, scores: [], students: {} };
    map[key].total++;
    if (isPassed) map[key].passed++;
    else if (isFailed) map[key].failed++;
    map[key].scores.push(pctVal);
    map[key].students[studentId] = true;
  }

  function addToSubGroup(map, primaryKey, subDim, subKey, isPassed, isFailed, pctVal, studentId) {
    if (!map[primaryKey]) return;
    if (!map[primaryKey]._sub) map[primaryKey]._sub = {};
    if (!map[primaryKey]._sub[subDim]) map[primaryKey]._sub[subDim] = {};
    addToGroup(map[primaryKey]._sub[subDim], subKey, isPassed, isFailed, pctVal, studentId);
  }

  function computeStats(obj) {
    var avg = 0, median = 0;
    if (obj.scores.length > 0) {
      var sum = 0;
      for (var s = 0; s < obj.scores.length; s++) sum += obj.scores[s];
      avg = Math.round(sum / obj.scores.length);
      var sorted = obj.scores.slice().sort(function(a, b) { return a - b; });
      var mid = Math.floor(sorted.length / 2);
      median = sorted.length % 2 !== 0 ? sorted[mid] : Math.round((sorted[mid - 1] + sorted[mid]) / 2);
    }
    var passRate = obj.total > 0 ? Math.round((obj.passed / obj.total) * 100) : 0;
    var studentCount = Object.keys(obj.students || {}).length;
    return { total: obj.total, passed: obj.passed, failed: obj.failed, passRate: passRate, avgScore: avg, medianScore: median, students: studentCount };
  }

  function computeGroupWithSub(map) {
    var out = {};
    for (var key in map) {
      out[key] = computeStats(map[key]);
      if (map[key]._sub) {
        out[key].sub = {};
        for (var subDim in map[key]._sub) {
          out[key].sub[subDim] = {};
          for (var subKey in map[key]._sub[subDim]) {
            out[key].sub[subDim][subKey] = computeStats(map[key]._sub[subDim][subKey]);
          }
        }
      }
    }
    return out;
  }

  var overallStats = computeStats(overall);
  overallStats.activeTeachers = Object.keys(overall.teachers).length;
  overallStats.activeClasses = Object.keys(overall.classes).length;
  overallStats.activeSites = Object.keys(overall.sites).length;

  var result = {
    overall: overallStats,
    byTeacher: computeGroupWithSub(byTeacher),
    byClass: computeGroupWithSub(byClass),
    byLicense: computeGroupWithSub(byLicense),
    byMode: computeGroupWithSub(byMode),
    commanderRole: role,
    commanderSite: userSite
  };

  // Add bySite only for global commander
  if (isGlobal) {
    result.bySite = computeGroupWithSub(bySite);
  }

  // Active classes list
  var studSheet = getSheet('תלמידי כיתות');
  var studData = studSheet.getDataRange().getValues();
  var studCountMap = {};
  for (var sc = 1; sc < studData.length; sc++) {
    var scc = String(studData[sc][0]).trim();
    studCountMap[scc] = (studCountMap[scc] || 0) + 1;
  }
  var activeClasses = [];
  for (var ac = 1; ac < classData.length; ac++) {
    if (String(classData[ac][6]) !== 'כן') continue; // only active
    var acCode = String(classData[ac][0]).trim();
    var acSite = String(classData[ac][7] || '');
    if (isLocal && userSite && acSite !== userSite) continue;
    activeClasses.push({
      code: acCode,
      name: String(classData[ac][1] || ''),
      teacherName: String(classData[ac][3] || ''),
      license: String(classData[ac][4] || ''),
      site: acSite,
      students: studCountMap[acCode] || 0
    });
  }
  result.activeClasses = activeClasses;

  return jsonResponse({ status: 'ok', data: result });
}

function handleAdminDashboard(p) {
  // Verify admin role - check ALL rows for this ID
  var tSheet = getSheet('מורים');
  var tData = tSheet.getDataRange().getValues();
  var isAdmin = false;
  for (var i = 1; i < tData.length; i++) {
    if (normalizeId(tData[i][1]) === normalizeId(p.teacherId)) {
      if (String(tData[i][8] || '') === 'אדמין') { isAdmin = true; break; }
    }
  }
  if (!isAdmin) {
    return jsonResponse({ status: 'error', message: 'אין הרשאת אדמין' });
  }

  var dateFrom = parseDateParam(p.dateFrom);
  var dateTo = parseDateParam(p.dateTo);
  if (!dateFrom || !dateTo) {
    return jsonResponse({ status: 'error', message: 'תאריכים לא תקינים' });
  }
  dateTo.setHours(23, 59, 59, 999);

  // Build class map
  var classSheet = getSheet('כיתות');
  var classData = classSheet.getDataRange().getValues();
  var classMap = {};
  for (var c = 1; c < classData.length; c++) {
    var cc = String(classData[c][0]).trim();
    classMap[cc] = {
      teacherName: String(classData[c][3] || ''),
      className: String(classData[c][1] || ''),
      license: String(classData[c][4] || ''),
      site: String(classData[c][7] || '')
    };
  }

  // Read practice results - INCLUDING rows without classCode
  var resSheet = getSheet('תוצאות תרגול');
  var resData = resSheet.getDataRange().getValues();

  var overall = { total: 0, passed: 0, failed: 0, scores: [], students: {}, classes: {}, independentStudents: {} };
  var byLicense = {};
  var byMode = {};
  var byEnrollment = {};
  var byClass = {};
  var byDay = {};

  function addToGroup(map, key, isPassed, isFailed, pctVal, studentId) {
    if (!map[key]) map[key] = { total: 0, passed: 0, failed: 0, scores: [], students: {} };
    map[key].total++;
    if (isPassed) map[key].passed++;
    else if (isFailed) map[key].failed++;
    map[key].scores.push(pctVal);
    map[key].students[studentId] = true;
  }

  function addToSubGroup(map, primaryKey, subDim, subKey, isPassed, isFailed, pctVal, studentId) {
    if (!map[primaryKey]) return;
    if (!map[primaryKey]._sub) map[primaryKey]._sub = {};
    if (!map[primaryKey]._sub[subDim]) map[primaryKey]._sub[subDim] = {};
    addToGroup(map[primaryKey]._sub[subDim], subKey, isPassed, isFailed, pctVal, studentId);
  }

  function computeStats(obj) {
    var avg = 0, median = 0;
    if (obj.scores.length > 0) {
      var sum = 0;
      for (var s = 0; s < obj.scores.length; s++) sum += obj.scores[s];
      avg = Math.round(sum / obj.scores.length);
      var sorted = obj.scores.slice().sort(function(a, b) { return a - b; });
      var mid = Math.floor(sorted.length / 2);
      median = sorted.length % 2 !== 0 ? sorted[mid] : Math.round((sorted[mid - 1] + sorted[mid]) / 2);
    }
    var passRate = obj.total > 0 ? Math.round((obj.passed / obj.total) * 100) : 0;
    var studentCount = Object.keys(obj.students || {}).length;
    return { total: obj.total, passed: obj.passed, failed: obj.failed, passRate: passRate, avgScore: avg, medianScore: median, students: studentCount };
  }

  function computeGroupWithSub(map) {
    var out = {};
    for (var key in map) {
      out[key] = computeStats(map[key]);
      if (map[key]._sub) {
        out[key].sub = {};
        for (var subDim in map[key]._sub) {
          out[key].sub[subDim] = {};
          for (var subKey in map[key]._sub[subDim]) {
            out[key].sub[subDim][subKey] = computeStats(map[key]._sub[subDim][subKey]);
          }
        }
      }
    }
    return out;
  }

  function fmtDay(d) {
    var dd = d.getDate(), mm = d.getMonth() + 1, yyyy = d.getFullYear();
    return (dd < 10 ? '0' : '') + dd + '/' + (mm < 10 ? '0' : '') + mm + '/' + yyyy;
  }

  for (var r = 1; r < resData.length; r++) {
    var rowDate = parseSheetDate(resData[r][0]);
    if (!rowDate || rowDate < dateFrom || rowDate > dateTo) continue;

    var classCode = String(resData[r][3] || '').trim();
    var isIndependent = !classCode;
    var enrollmentStatus = isIndependent ? 'עצמאי' : 'כיתה';
    var cInfo = classCode ? (classMap[classCode] || { teacherName: 'לא ידוע', className: classCode, license: '', site: '' }) : null;

    var license = String(resData[r][5] || (cInfo ? cInfo.license : '') || 'לא צוין');
    var mode = String(resData[r][4] || 'לא צוין');
    var studentId = String(resData[r][1] || '');
    var studentName = String(resData[r][2] || '');
    var passedStr = String(resData[r][9] || '');
    var isPassed = (passedStr === 'עבר' || passedStr === 'true' || passedStr === true);
    var isFailed = (passedStr === 'נכשל' || passedStr === 'false' || passedStr === false);

    var pctVal = 0;
    var pctRaw = resData[r][8];
    if (typeof pctRaw === 'string' && pctRaw.indexOf('%') !== -1) {
      pctVal = parseFloat(pctRaw.replace('%', '')) || 0;
    } else {
      var pctNum = Number(pctRaw);
      if (!isNaN(pctNum)) {
        pctVal = pctNum <= 1 ? pctNum * 100 : pctNum;
      }
    }

    overall.total++;
    if (isPassed) overall.passed++;
    else if (isFailed) overall.failed++;
    overall.scores.push(pctVal);
    overall.students[studentId] = true;
    if (classCode) overall.classes[classCode] = true;
    if (isIndependent) overall.independentStudents[studentId] = true;

    // By license
    addToGroup(byLicense, license, isPassed, isFailed, pctVal, studentId);
    addToSubGroup(byLicense, license, 'byMode', mode, isPassed, isFailed, pctVal, studentId);
    addToSubGroup(byLicense, license, 'byEnrollment', enrollmentStatus, isPassed, isFailed, pctVal, studentId);

    // By mode
    addToGroup(byMode, mode, isPassed, isFailed, pctVal, studentId);
    addToSubGroup(byMode, mode, 'byLicense', license, isPassed, isFailed, pctVal, studentId);

    // By enrollment
    addToGroup(byEnrollment, enrollmentStatus, isPassed, isFailed, pctVal, studentId);
    addToSubGroup(byEnrollment, enrollmentStatus, 'byLicense', license, isPassed, isFailed, pctVal, studentId);

    // By class (only for enrolled students)
    if (classCode) {
      var className = cInfo ? cInfo.className : classCode;
      addToGroup(byClass, className + ' (' + classCode + ')', isPassed, isFailed, pctVal, studentId);
      addToSubGroup(byClass, className + ' (' + classCode + ')', 'byLicense', license, isPassed, isFailed, pctVal, studentId);
      addToSubGroup(byClass, className + ' (' + classCode + ')', 'byMode', mode, isPassed, isFailed, pctVal, studentId);
    }

    // By day
    var dayKey = fmtDay(rowDate);
    addToGroup(byDay, dayKey, isPassed, isFailed, pctVal, studentId);
  }

  var overallStats = computeStats(overall);
  overallStats.activeClasses = Object.keys(overall.classes).length;
  overallStats.independentStudents = Object.keys(overall.independentStudents).length;

  return jsonResponse({
    status: 'ok',
    data: {
      overall: overallStats,
      byLicense: computeGroupWithSub(byLicense),
      byMode: computeGroupWithSub(byMode),
      byEnrollment: computeGroupWithSub(byEnrollment),
      byClass: computeGroupWithSub(byClass),
      byDay: computeGroupWithSub(byDay)
    }
  });
}

function handleTeacherCreateClass(p) {
  var code = generateClassCode();
  var className = p.className || 'כיתה חדשה';
  var license = p.license || 'B';
  var sheet = getSheet('כיתות');
  // Get teacher name
  var tSheet = getSheet('מורים');
  var tData = tSheet.getDataRange().getValues();
  var teacherName = '';
  for (var i = 1; i < tData.length; i++) {
    if (normalizeId(tData[i][1]) === normalizeId(p.teacherId)) { teacherName = tData[i][0]; break; }
  }
  sheet.appendRow([code, className, normalizeId(p.teacherId), teacherName, license, nowISO(), 'כן']);
  return jsonResponse({ status: 'ok', classCode: code, className: className });
}

function handleTeacherCloseClass(p) {
  var sheet = getSheet('כיתות');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(p.classCode).trim() &&
        normalizeId(data[i][2]) === normalizeId(p.teacherId)) {
      sheet.getRange(i + 1, 7).setValue('לא');
      return jsonResponse({ status: 'ok' });
    }
  }
  return jsonResponse({ status: 'error', message: 'כיתה לא נמצאה' });
}

function handleTeacherRemoveStudent(p) {
  var sheet = getSheet('תלמידי כיתות');
  var data = sheet.getDataRange().getValues();
  // Verify teacher owns this class
  var classSheet = getSheet('כיתות');
  var classData = classSheet.getDataRange().getValues();
  var ownsClass = false;
  for (var c = 1; c < classData.length; c++) {
    if (String(classData[c][0]).trim() === String(p.classCode).trim() &&
        normalizeId(classData[c][2]) === normalizeId(p.teacherId)) {
      ownsClass = true; break;
    }
  }
  if (!ownsClass) return jsonResponse({ status: 'error', message: 'אין הרשאה' });

  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]).trim() === String(p.classCode).trim() &&
        String(data[i][2]).trim() === String(p.studentId).trim()) {
      sheet.deleteRow(i + 1);
      return jsonResponse({ status: 'ok' });
    }
  }
  return jsonResponse({ status: 'error', message: 'תלמיד לא נמצא' });
}

function handleTeacherGetClasses(p) {
  var sheet = getSheet('כיתות');
  var data = sheet.getDataRange().getValues();
  var studSheet = getSheet('תלמידי כיתות');
  var studData = studSheet.getDataRange().getValues();

  // Count students per class
  var studentCounts = {};
  for (var s = 1; s < studData.length; s++) {
    var cc = String(studData[s][0]).trim();
    studentCounts[cc] = (studentCounts[cc] || 0) + 1;
  }

  var classes = [];
  for (var i = 1; i < data.length; i++) {
    if (normalizeId(data[i][2]) === normalizeId(p.teacherId)) {
      var classCode = String(data[i][0]).trim();
      classes.push({
        code: classCode,
        name: data[i][1],
        license: data[i][4] || 'B',
        created: data[i][5],
        active: data[i][6] === 'כן',
        studentCount: studentCounts[classCode] || 0
      });
    }
  }
  return jsonResponse({ status: 'ok', classes: classes });
}

function handleTeacherClassDetails(p) {
  var classCode = String(p.classCode || '').trim();
  if (!classCode) return jsonResponse({ status: 'error', message: 'חסר קוד כיתה' });

  // Verify teacher owns this class
  var classSheet = getSheet('כיתות');
  var classData = classSheet.getDataRange().getValues();
  var classInfo = null;
  for (var c = 1; c < classData.length; c++) {
    if (String(classData[c][0]).trim() === classCode && normalizeId(classData[c][2]) === normalizeId(p.teacherId)) {
      classInfo = { code: classCode, name: classData[c][1], license: classData[c][4], active: classData[c][6] === 'כן' };
      break;
    }
  }
  if (!classInfo) return jsonResponse({ status: 'error', message: 'כיתה לא נמצאה' });

  // Get students in class
  var studSheet = getSheet('תלמידי כיתות');
  var studData = studSheet.getDataRange().getValues();
  var studentIds = [];
  var studentMap = {};
  for (var s = 1; s < studData.length; s++) {
    if (String(studData[s][0]).trim() === classCode) {
      var sid = String(studData[s][2]).trim();
      studentIds.push(sid);
      studentMap[sid] = { name: studData[s][1], id: sid, joined: studData[s][3] };
    }
  }

  // Get practice results for these students
  var resSheet = getSheet('תוצאות תרגול');
  var resData = resSheet.getDataRange().getValues();
  var studentResults = {};
  for (var r = 1; r < resData.length; r++) {
    var rSid = String(resData[r][1]).trim();
    var rClass = String(resData[r][3]).trim();
    if (rClass === classCode && studentIds.indexOf(rSid) !== -1) {
      if (!studentResults[rSid]) studentResults[rSid] = [];
      studentResults[rSid].push({
        date: resData[r][0],
        mode: resData[r][4],
        license: resData[r][5],
        score: resData[r][6],
        total: resData[r][7],
        percent: resData[r][8],
        passed: resData[r][9],
        time: resData[r][10],
        category: resData[r][11] || '',
        language: resData[r][12] || 'he',
        wrongDetails: resData[r][13] || '',
        categoryBreakdown: resData[r][14] || ''
      });
    }
  }

  // Build student summaries
  var students = [];
  for (var si = 0; si < studentIds.length; si++) {
    var id = studentIds[si];
    var info = studentMap[id];
    var results = studentResults[id] || [];
    var totalExams = 0, totalPassed = 0, scores = [], lastActive = '';
    var categoryErrors = {};
    for (var ri = 0; ri < results.length; ri++) {
      var res = results[ri];
      var pctVal = Number(res.percent) || 0;
      if (pctVal > 0) scores.push(pctVal);
      if (res.mode === 'exam') {
        totalExams++;
        if (res.passed === 'עבר' || res.passed === true) totalPassed++;
      }
      if (res.date && (!lastActive || String(res.date) > String(lastActive))) lastActive = res.date;
      // Aggregate category errors
      if (res.categoryBreakdown) {
        try {
          var cb = typeof res.categoryBreakdown === 'string' ? JSON.parse(res.categoryBreakdown) : res.categoryBreakdown;
          for (var cat in cb) {
            if (!categoryErrors[cat]) categoryErrors[cat] = { correct: 0, total: 0 };
            categoryErrors[cat].correct += (cb[cat].correct || 0);
            categoryErrors[cat].total += (cb[cat].total || 0);
          }
        } catch(e) {}
      }
    }
    var avgScore = 0;
    if (scores.length > 0) {
      var sum = 0;
      for (var sc = 0; sc < scores.length; sc++) sum += scores[sc];
      avgScore = Math.round(sum / scores.length);
    }
    students.push({
      name: info.name,
      id: id,
      joined: info.joined,
      totalPractices: results.length,
      totalExams: totalExams,
      totalPassed: totalPassed,
      avgScore: avgScore,
      lastActive: lastActive,
      categoryErrors: categoryErrors,
      recentResults: results.slice(-10) // Last 10 results
    });
  }

  return jsonResponse({ status: 'ok', classInfo: classInfo, students: students });
}

function handleTeacherDashboard(p) {
  // Overview of all classes for a teacher
  return handleTeacherGetClasses(p);
}

function handleTeacherExportData(p) {
  var classCode = String(p.classCode || '').trim();
  if (!classCode) return jsonResponse({ status: 'error', message: 'חסר קוד כיתה' });

  // Verify ownership
  var classSheet = getSheet('כיתות');
  var classData = classSheet.getDataRange().getValues();
  var owns = false;
  for (var c = 1; c < classData.length; c++) {
    if (String(classData[c][0]).trim() === classCode && normalizeId(classData[c][2]) === normalizeId(p.teacherId)) {
      owns = true; break;
    }
  }
  if (!owns) return jsonResponse({ status: 'error', message: 'אין הרשאה' });

  // Get all results for this class
  var resSheet = getSheet('תוצאות תרגול');
  var resData = resSheet.getDataRange().getValues();
  var headers = resData[0];
  var rows = [];
  for (var r = 1; r < resData.length; r++) {
    if (String(resData[r][3]).trim() === classCode) {
      var row = {};
      for (var h = 0; h < headers.length; h++) row[headers[h]] = resData[r][h];
      rows.push(row);
    }
  }
  return jsonResponse({ status: 'ok', headers: headers, rows: rows });
}

function handleStudentJoinClass(p) {
  var classCode = String(p.classCode || '').trim().toUpperCase();
  var studentName = String(p.studentName || '').trim();
  var studentId = String(p.studentId || '').trim();
  if (!classCode || !studentName || !studentId) {
    return jsonResponse({ status: 'error', message: 'חסרים פרטים (קוד כיתה, שם, מזהה)' });
  }

  // Verify class exists and active
  var classSheet = getSheet('כיתות');
  var classData = classSheet.getDataRange().getValues();
  var classInfo = null;
  for (var c = 1; c < classData.length; c++) {
    if (String(classData[c][0]).trim() === classCode) {
      if (classData[c][6] !== 'כן') return jsonResponse({ status: 'error', message: 'הכיתה אינה פעילה' });
      classInfo = { name: classData[c][1], teacherName: classData[c][3], license: classData[c][4] };
      break;
    }
  }
  if (!classInfo) return jsonResponse({ status: 'error', message: 'כיתה לא נמצאה' });

  // Check if already enrolled
  var studSheet = getSheet('תלמידי כיתות');
  var studData = studSheet.getDataRange().getValues();
  for (var s = 1; s < studData.length; s++) {
    if (String(studData[s][0]).trim() === classCode && String(studData[s][2]).trim() === studentId) {
      return jsonResponse({ status: 'ok', message: 'כבר רשום בכיתה', className: classInfo.name, teacherName: classInfo.teacherName, license: classInfo.license });
    }
  }

  studSheet.appendRow([classCode, studentName, studentId, nowISO()]);
  return jsonResponse({ status: 'ok', message: 'הצטרפת לכיתה בהצלחה!', className: classInfo.name, teacherName: classInfo.teacherName, license: classInfo.license });
}

function handleSubmitPracticeResult(p) {
  var sheet = getSheet('תוצאות תרגול');
  var studentId = String(p.studentId || '').trim();
  var classCode = String(p.classCode || '').trim();
  var mode = String(p.mode || 'exam');
  var license = String(p.license || 'B');
  var score = Number(p.score) || 0;
  var total = Number(p.total) || 0;
  var percent = Number(p.percent) || 0;
  var passed = percent >= 86 ? 'עבר' : 'נכשל';
  var time = String(p.time || '');
  var category = String(p.category || '');
  var language = String(p.language || 'he');
  var wrongDetails = '';
  try { wrongDetails = typeof p.wrongDetails === 'string' ? p.wrongDetails : JSON.stringify(p.wrongDetails || ''); } catch(e) {}
  var categoryBreakdown = '';
  try { categoryBreakdown = typeof p.categoryBreakdown === 'string' ? p.categoryBreakdown : JSON.stringify(p.categoryBreakdown || ''); } catch(e) {}

  sheet.appendRow([todayStr(), studentId, String(p.studentName || ''), classCode, mode, license, score, total, percent, passed, time, category, language, wrongDetails, categoryBreakdown]);
  return jsonResponse({ status: 'ok' });
}
