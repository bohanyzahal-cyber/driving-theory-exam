// ===== Google Apps Script — מערכת בחינות חיצונית =====
// הדבק את הקוד הזה ב-Apps Script של גיליון Google Sheets חדש
// Deploy → New deployment → Web app
// Execute as: Me | Who has access: Anyone
// העתק את ה-URL שמקבלים והדבק ב-examiner.html וב-examinee.html

// ========== פונקציות עזר ==========

var SHEET_HEADERS = {
  'בוחנים': ['שם', 'ת.ז.', 'סיסמה', 'פעיל', 'מס בוחן'],
  'אתרים': ['שם אתר', 'מזהה', 'טלפון מנהל', 'כיתות'],
  'סשנים': ['קוד', 'בוחן ת.ז.', 'שם בוחן', 'אתר', 'כיתה', 'דרגה', 'שפה', 'מצב שמע', 'זמן יצירה', 'תקף עד', 'פעיל'],
  'ממתינים': ['קוד סשן', 'ת.ז.', 'שם', 'טלפון', 'זמן הרשמה', 'סטטוס', 'שפה', 'אוכלוסיה', 'דרגה', 'שמע'],
  'תוצאות': ['תאריך', 'ת.ז.', 'שם', 'טלפון', 'דרגה', 'ציון', 'אחוז', 'עבר/נכשל', 'זמן', 'בוחן', 'אתר', 'כיתה', 'שפה', 'קוד סשן', 'ניסיון', 'פירוט שגויות', 'נשלח?', 'פסול?', 'קישור וואטסאפ', 'אוכלוסיה', 'תוקן?', 'שמע']
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
  var code;
  do {
    code = String(Math.floor(100000 + Math.random() * 900000));
  } while (existingCodes[code]);
  return code;
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
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

    switch (action) {

      case 'login':
        return handleLogin(p);

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

      case 'resetExaminee':
        return handleResetExaminee(p);

      case 'correctToPass':
        return handleCorrectToPass(p);

      case 'forceComplete':
        return handleForceComplete(p);

      case 'markSent':
        return handleMarkSent(p);

      case 'debugSession':
        return handleDebugSession(p);

      case 'debugResults':
        return handleDebugResults();

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

    if (action === 'submitResult') {
      return handleSubmitResult(data);
    } else if (action === 'submitFailOnClose') {
      return handleSubmitFailOnClose(data);
    } else if (action === 'submitWrongAnswers') {
      return handleSubmitWrongAnswersBulk(data);
    } else if (action === 'uploadResultHtml') {
      return handleUploadResultHtml(data);
    } else if (action === 'disqualify') {
      return handleDisqualify(data);
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
      if (String(data[i][2]) === String(p.password)) {
        if (data[i][3] === 'כן' || data[i][3] === true || data[i][3] === 'TRUE') {
          return jsonResponse({ status: 'ok', examiner: { name: data[i][0], id: normalizeId(data[i][1]), examinerNumber: String(data[i][4] || '') } });
        } else {
          return jsonResponse({ status: 'error', message: 'החשבון אינו פעיל' });
        }
      } else {
        return jsonResponse({ status: 'error', message: 'סיסמה שגויה' });
      }
    }
  }
  return jsonResponse({ status: 'error', message: 'בוחן לא נמצא' });
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
  // Debug info: return how many rows and first code found for troubleshooting
  var debugInfo = 'rows=' + data.length;
  if (data.length > 1) {
    debugInfo += ', firstCode=[' + String(data[1][0]) + '], type=' + typeof data[1][0] + ', cols=' + data[1].length;
  }
  return jsonResponse({ status: 'error', message: 'קוד סשן לא תקין (' + debugInfo + ')' });
}

function handleRegisterExaminee(p) {
  var pendSheet = getSheet('ממתינים');
  var data = pendSheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(p.sessionCode) && normalizeId(data[i][1]) === normalizeId(p.idNumber)) {
      var status = data[i][5];
      if (status === 'waiting' || status === 'approved' || status === 'in_exam') {
        return jsonResponse({ status: 'error', message: 'כבר רשום בסשן זה' });
      }
    }
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
    p.audioMode || 'off'
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
        sheet.getRange(i + 1, 6).setValue('cancelled');
        SpreadsheetApp.flush();
        return jsonResponse({ status: 'ok' });
      }
    }
  }
  return jsonResponse({ status: 'error', message: 'לא נמצא רישום פעיל לביטול' });
}

function handleCheckApproval(p) {
  var sheet = getSheet('ממתינים');
  var data = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]).trim() === String(p.sessionCode).trim() && normalizeId(data[i][1]) === normalizeId(p.idNumber)) {
      var approval = String(data[i][5] || 'waiting').trim();
      // Skip terminal statuses from previous exams — keep looking for active row
      if (approval === 'completed' || approval === 'disqualified' || approval === 'cancelled') continue;
      return jsonResponse({ status: 'ok', approval: approval });
    }
  }
  return jsonResponse({ status: 'error', message: 'לא נמצא רישום' });
}

function handleApproveExaminee(p) {
  if (p.examinerId && !verifyExaminerForSession(p.sessionCode, p.examinerId)) {
    return jsonResponse({ status: 'error', message: 'אין הרשאה — בוחן לא תואם לסשן' });
  }
  var sheet = getSheet('ממתינים');
  var data = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === String(p.sessionCode) && normalizeId(data[i][1]) === normalizeId(p.idNumber) && String(data[i][5]).trim() === 'waiting') {
      sheet.getRange(i + 1, 6).setValue('approved');
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
  var MAX_EXAM_MS = 90 * 60 * 1000; // 90 minutes — very conservative threshold
  for (var ci = 1; ci < pendData.length; ci++) {
    if (String(pendData[ci][0]) !== code) continue;
    if (String(pendData[ci][5]).trim() !== 'in_exam') continue;
    var ciId = pendData[ci][1];
    var regTime = pendData[ci][4] ? new Date(pendData[ci][4]) : null;
    var isStale = regTime && (now.getTime() - regTime.getTime() > MAX_EXAM_MS);

    // Check if result already exists for this examinee in this session
    var hasResult = false;
    for (var ri = 1; ri < resData.length; ri++) {
      if (String(resData[ri][13]) === code && normalizeId(resData[ri][1]) === normalizeId(ciId)) {
        hasResult = true;
        break;
      }
    }

    if (hasResult || isStale) {
      // Fix dangling status — mark as completed
      pendSheet.getRange(ci + 1, 6).setValue('completed');
      pendData[ci][5] = 'completed'; // update local copy
      if (isStale && !hasResult) {
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
    var s = pendData[i][5];
    var item = { idNumber: pendData[i][1], name: pendData[i][2], phone: pendData[i][3], time: pendData[i][4], status: s, language: pendData[i][6] || '', population: pendData[i][7] || '', license: pendData[i][8] || '', audioMode: pendData[i][9] || 'off' };
    if (s === 'waiting') pending.push(item);
    else if (s === 'approved') pending.push(item);
    else if (s === 'in_exam') active.push(item);
  }

  // Re-read resData in case cleanup added new results
  resData = resSheet.getDataRange().getValues();
  var completed = [];
  for (var j = 1; j < resData.length; j++) {
    if (String(resData[j][13]) !== code) continue;
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

  // Check if result already exists in תוצאות
  var sheet = getSheet('תוצאות');
  var data = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][13]) === String(p.sessionCode) && normalizeId(data[i][1]) === normalizeId(p.idNumber)) {
      sheet.getRange(i + 1, 18).setValue(true);  // פסול? — column R (18)
      sheet.getRange(i + 1, 8).setValue('פסול');
      SpreadsheetApp.flush();
      return jsonResponse({ status: 'ok' });
    }
  }

  // No result yet — create disqualified result
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
    population, false, examineeAudio
  ]);
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

function handleDebugSession(p) {
  var sheet = getSheet('סשנים');
  var data = sheet.getDataRange().getValues();
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    rows.push({
      code: String(data[i][0]),
      codeType: typeof data[i][0],
      examiner: String(data[i][2]),
      site: String(data[i][3]),
      license: String(data[i][5]),
      active: String(data[i][10]),
      created: String(data[i][8])
    });
  }
  return jsonResponse({
    status: 'ok',
    totalSessions: data.length - 1,
    sessions: rows
  });
}

function handleDebugResults() {
  var sheet = getSheet('תוצאות');
  var data = sheet.getDataRange().getValues();
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    rows.push({
      date: String(data[i][0]),
      idNumber: String(data[i][1]),
      name: String(data[i][2]),
      score: String(data[i][5]),
      passed: String(data[i][7]),
      sessionCode: String(data[i][13]),
      sessionCodeType: typeof data[i][13]
    });
  }
  return jsonResponse({ status: 'ok', totalResults: data.length - 1, results: rows });
}

function handleSubmitResult(data) {
  var sheet = getSheet('תוצאות');

  // Duplicate protection: check if result already exists for this session+ID+license+language
  var existingData = sheet.getDataRange().getValues();
  for (var d = 1; d < existingData.length; d++) {
    if (String(existingData[d][13]) === String(data.sessionCode) && normalizeId(existingData[d][1]) === normalizeId(data.idNumber) && String(existingData[d][4]) === String(data.license) && String(existingData[d][12]) === String(data.language || 'he')) {
      // Already submitted — still update pending status so examinee doesn't stay stuck in "in_exam"
      markPendingCompleted(data.sessionCode, data.idNumber);
      return jsonResponse({ status: 'ok', waLink: existingData[d][18] || '', duplicate: true });
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
    data.audioMode || 'off'
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
