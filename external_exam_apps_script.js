// © 2026 Vitaly Gitelman. All Rights Reserved.
// Unauthorized copying, modification or distribution is prohibited.
// ===== Google Apps Script — מערכת בחינות חיצונית =====
// הדבק את הקוד הזה ב-Apps Script של גיליון Google Sheets חדש
// Deploy → New deployment → Web app
// Execute as: Me | Who has access: Anyone
// העתק את ה-URL שמקבלים והדבק ב-examiner.html וב-examinee.html

// ========== פונקציות עזר ==========

var SHEET_HEADERS = {
  'בוחנים': ['שם', 'ת.ז.', 'סיסמה', 'פעיל', 'מס בוחן', 'תפקיד', 'טוקן', 'תוקף טוקן', 'ניסיונות כושלים', 'נעילה עד', 'אתרים מנוהלים'],
  'אתרים': ['שם אתר', 'מזהה', 'טלפון מנהל', 'כיתות'],
  'סשנים': ['קוד', 'בוחן ת.ז.', 'שם בוחן', 'אתר', 'כיתה', 'דרגה', 'שפה', 'מצב שמע', 'זמן יצירה', 'תקף עד', 'פעיל', 'כמויות JSON', 'מאושרים JSON', 'בוחן אחראי'],
  'ממתינים': ['קוד סשן', 'ת.ז.', 'שם', 'טלפון', 'זמן הרשמה', 'סטטוס', 'שפה', 'אוכלוסיה', 'דרגה', 'שמע', 'הארכת זמן', 'התחלת מבחן', 'טוקן נבחן', 'ספירת DQ', 'מסך נוסף'],
  'תוצאות': ['תאריך', 'ת.ז.', 'שם', 'טלפון', 'דרגה', 'ציון', 'אחוז', 'עבר/נכשל', 'זמן', 'בוחן', 'אתר', 'כיתה', 'שפה', 'קוד סשן', 'ניסיון', 'פירוט שגויות', 'נשלח?', 'פסול?', 'קישור וואטסאפ', 'אוכלוסיה', 'תוקן?', 'שמע', 'מאומת', 'חשוד', 'dqEventId', 'תוקן ע"י', 'סיבת תיקון', 'תאריך תיקון', 'מסלול שפות', 'מכשיר'],
  'הארכות זמן': ['תאריך', 'קוד סשן', 'ת.ז.', 'שם', 'דקות', 'סיבה', 'בוחן'],
  'מורים': ['שם', 'ת.ז.', 'סיסמה', 'פעיל', 'טוקן', 'תוקף טוקן', 'ניסיונות כושלים', 'נעילה עד'],
  'כיתות': ['קוד כיתה', 'שם כיתה', 'מורה ת.ז.', 'שם מורה', 'דרגה', 'תאריך יצירה', 'פעיל'],
  'תלמידי כיתות': ['קוד כיתה', 'שם תלמיד', 'מזהה תלמיד', 'תאריך הצטרפות'],
  'תוצאות תרגול': ['תאריך', 'מזהה תלמיד', 'שם תלמיד', 'קוד כיתה', 'מצב', 'דרגה', 'ציון', 'סה"כ', 'אחוז', 'עבר/נכשל', 'זמן', 'נושא', 'שפה', 'פירוט שגויות', 'פירוט לפי נושא', 'טלפון'],
  'התקדמות תלמידים': ['שם תלמיד', 'קוד כיתה', 'מפתח', 'streak', 'wrong_qs', 'history', 'עדכון אחרון']
};

// Sites used ONLY for system testing by examiners (not real exams). Their rows are
// EXCLUDED from the commander dashboard statistics so test data doesn't pollute the
// real numbers. They are NOT filtered from the live examiner dashboard — a tester
// still needs to see their own test session. Add more names here if needed.
var TEST_SITES = ['בדיקת נתונים', 'דימונה דוגית 35'];
function isTestSite(site) {
  return TEST_SITES.indexOf(String(site || '').trim()) !== -1;
}

// Person-name normalizer for fuzzy matching: strip punctuation, collapse spaces,
// lowercase, token-sort (so "ויטלי גיטלמן" and "גיטלמן ויטלי" hash the same).
function normalizeNameKey(s) {
  if (!s) return '';
  var t = String(s).replace(/[׳״'".\-]/g, ' ').replace(/\s+/g, ' ').trim().toLowerCase();
  if (!t) return '';
  var tokens = t.split(' ').filter(function(x) { return x; });
  tokens.sort();
  return tokens.join(' ');
}

// Examiner identity sets (names + IDs) from the בוחנים sheet, for excluding examiners
// who registered as EXAMINEES to test the system. ID match is exact (no false positives,
// since a ת.ז. is unique to the examiner); name match is fuzzy (normalizeNameKey) and can
// rarely catch a real same-named candidate. Read once per report.
function getExaminerExclusion() {
  var names = {}, ids = {};
  try {
    var d = getSheet('בוחנים').getDataRange().getValues();
    for (var i = 1; i < d.length; i++) {
      var nk = normalizeNameKey(d[i][0]);   // col 0 = שם
      if (nk) names[nk] = true;
      var ik = normalizeId(d[i][1]);         // col 1 = ת.ז.
      if (ik) ids[ik] = true;
    }
  } catch (e) {}
  return { names: names, ids: ids };
}
function isExaminerSelfTest(name, id, excl) {
  if (!excl) return false;
  var ik = normalizeId(id);
  if (ik && excl.ids[ik]) return true;       // ת.ז. match — precise
  var nk = normalizeNameKey(name);
  return !!(nk && excl.names[nk]);           // name match — fuzzy fallback
}

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

// ========== Office WhatsApp number (sender) ==========
// Dedicated number for outbound WhatsApp messages to examinees, registered
// 2026-05-15. Currently inactive — used only as a future-ready config for
// when the Meta Cloud API is wired up. Read via getOfficeWhatsAppNumber()
// so it can be overridden at runtime through Script Properties without redeploy.
var OFFICE_WHATSAPP_NUMBER_DEFAULT = '0529151157';
function getOfficeWhatsAppNumber() {
  try {
    var prop = PropertiesService.getScriptProperties().getProperty('OFFICE_WHATSAPP_NUMBER');
    if (prop && String(prop).trim()) return String(prop).trim();
  } catch (e) {}
  return OFFICE_WHATSAPP_NUMBER_DEFAULT;
}

// ========== Token authentication ==========
function generateToken() {
  // CSPRNG-backed examiner token (was Math.random, which is predictable). Three
  // UUIDs of hex → 96 hex chars, well above the prior 48-char entropy.
  return (Utilities.getUuid() + Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
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

// ========== Origin allowlist (soft check) ==========
// Apps Script cannot read HTTP headers, so the client must pass an `origin` parameter.
// This is bypassable (anyone reading the HTML sees the magic string) but blocks
// casual API exploration / generic scrapers / curl scripts. Real security comes
// from token enforcement.
var ALLOWED_ORIGINS = [
  'examiner-app',      // examiner.html
  'examinee-app',      // examinee.html / exam.html (standalone practice)
  'teacher-app',       // teacher.html
  'student-app',       // student.html
  'admin-app',         // admin.html
  'bohanyzahal-site',  // bohan-site (IDF portal — server-side auth)
  'localhost-dev'      // local development
];
function checkOrigin(p) {
  // Allowed: actions called from external services (none currently) or no origin enforcement on certain reads
  // For now: reject unknown origins on all actions except 'viewResult' (HTML output, opened in browser tab).
  var action = p.action || '';
  if (action === 'viewResult' || action === '') return null;
  var origin = String(p.origin || '').trim();
  if (!origin) {
    return jsonResponse({ status: 'error', message: 'Missing origin', code: 'origin_required' });
  }
  if (ALLOWED_ORIGINS.indexOf(origin) === -1) {
    return jsonResponse({ status: 'error', message: 'Unauthorized origin', code: 'origin_denied' });
  }
  return null;
}

// ========== Rate limiting (Stage 2b) ==========
// Sliding-window rate limit backed by CacheService. The cache stores a JSON
// array of recent request timestamps per (action, identifier). On each call we
// filter to the window, count, and either allow + append, or reject. CacheService
// auto-evicts entries by TTL — no manual cleanup needed.
//
// Trade-off note: Apps Script doesn't expose client IP, so identifiers must come
// from the request payload (sessionCode, idNumber, examinerId). This means an
// attacker who varies the identifier can avoid limits — but they'd still need
// valid creds to be useful, since the data still has to go through token and
// origin checks. This rate limit primarily defends quotas + same-target floods.
function checkRateLimit(action, identifier, maxRequests, windowSeconds) {
  try {
    var cache = CacheService.getScriptCache();
    var key = 'rl_' + action + '_' + identifier;
    var raw = cache.get(key);
    var now = Date.now();
    var windowMs = windowSeconds * 1000;
    var timestamps = [];
    if (raw) {
      try { timestamps = JSON.parse(raw) || []; } catch(_e) { timestamps = []; }
    }
    // Drop timestamps older than the window
    var fresh = [];
    for (var i = 0; i < timestamps.length; i++) {
      if ((now - timestamps[i]) < windowMs) fresh.push(timestamps[i]);
    }
    if (fresh.length >= maxRequests) {
      var oldest = fresh[0];
      var waitSec = Math.max(1, Math.ceil((windowMs - (now - oldest)) / 1000));
      return { ok: false, waitSec: waitSec };
    }
    fresh.push(now);
    cache.put(key, JSON.stringify(fresh), windowSeconds + 60); // TTL slightly over window
    return { ok: true };
  } catch (e) {
    // If the cache is unavailable, fail open — don't break the system.
    return { ok: true };
  }
}

// Convenience wrapper that returns a jsonResponse error on hit, or null on pass.
function requireRateLimit(action, identifier, maxRequests, windowSeconds) {
  if (!identifier) return null; // can't enforce without an identifier
  var result = checkRateLimit(action, identifier, maxRequests, windowSeconds || 60);
  if (!result.ok) {
    return jsonResponse({
      status: 'error',
      message: 'יותר מדי בקשות. נסה שוב בעוד ' + result.waitSec + ' שניות.',
      rateLimited: true,
      waitSec: result.waitSec
    });
  }
  return null;
}

// ========== Examinee token (Stage 1c) ==========
// Each examinee receives a random token at registration time. All subsequent
// examinee-side calls (poll, submit, self-DQ) must echo that token. Prevents
// an attacker registered to the same session from acting on a victim's row
// using only sessionCode + idNumber.
function generateExamineeToken() {
  // CSPRNG-backed (was Math.random, predictable). Two UUIDs of hex → 64 hex chars.
  return (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
}

// Returns { valid: bool, legacy: bool, reason: string }
// - legacy: true when the stored row predates token support (empty cell) —
//   we accept the call but flag it so we can audit / tighten later.
// - reason values (when invalid): 'not_found', 'missing', 'mismatch'.
function verifyExamineeToken(sessionCode, idNumber, examineeToken) {
  var sheet = getSheet('ממתינים');
  var data = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === String(sessionCode) && normalizeId(data[i][1]) === normalizeId(idNumber)) {
      var storedToken = String((data[i].length > 12 ? data[i][12] : '') || '').trim();
      if (!storedToken) return { valid: true, legacy: true };
      if (!examineeToken) return { valid: false, reason: 'missing' };
      if (String(examineeToken).trim() === storedToken) return { valid: true, legacy: false };
      return { valid: false, reason: 'mismatch' };
    }
  }
  return { valid: false, reason: 'not_found' };
}

// Convenience wrapper for handlers. Returns null when OK, or a jsonResponse error.
function requireExamineeToken(p) {
  if (!p.sessionCode || !p.idNumber) {
    return jsonResponse({ status: 'error', message: 'חסרים פרטי נבחן' });
  }
  var result = verifyExamineeToken(p.sessionCode, p.idNumber, p.examineeToken);
  if (!result.valid) {
    return jsonResponse({ status: 'error', message: 'טוקן נבחן לא תקין', examineeTokenError: result.reason });
  }
  return null;
}

// Look up the list of sites a center commander oversees (column K in בוחנים).
// Format in the sheet: comma-separated site names matching column K in תוצאות.
// Returns array of trimmed names (empty if not found / empty cell).
function getExaminerManagedSites(examinerId) {
  if (!examinerId) return [];
  var sheet = getSheet('בוחנים');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (normalizeId(data[i][1]) === normalizeId(examinerId)) {
      var raw = (data[i].length > 10) ? String(data[i][10] || '') : '';
      if (!raw) return [];
      return raw.split(',').map(function(s) { return s.trim(); }).filter(function(s) { return s; });
    }
  }
  return [];
}

// Look up examiner's role ('בוחן' or 'מפקד'). Returns '' if not found.
function getExaminerRole(examinerId) {
  if (!examinerId) return '';
  var sheet = getSheet('בוחנים');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (normalizeId(data[i][1]) === normalizeId(examinerId)) {
      return String(data[i][5] || 'בוחן').trim();
    }
  }
  return '';
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

// "מפקד קד״ץ" gets typed in many forms: with Hebrew gershayim ״, ASCII " or ',
// no separator at all ("מפקד קדץ"), with extra spaces. Match all of them so a
// sheet entry typed casually still resolves to the role.
function isKdtzRole(role) {
  return /^\s*מפקד\s+קד[\s׳״'"]*ץ\s*$/.test(String(role || ''));
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

    // Actions that require examiner token authentication
    var examinerActions = ['getSites','listSessions','listAllSessions','createSession','updateSession','closeSession',
      'approveExaminee','rejectExaminee','examinerDashboard','resetExaminee',
      'correctToPass','overturnDQ','confirmDQ','forceComplete','markSent','commanderDashboard',
      'commanderCorrectResult','correctExamineeMeta','getResultUploadToken','centerManagerReport'];
    // Note: 'disqualify' is NOT in this list because it can be sent by the examinee client (no token)
    // — auth is enforced inside handleDisqualify itself (examiner token OR active pending row).
    if (examinerActions.indexOf(action) !== -1) {
      var tokenErr = requireToken(p);
      if (tokenErr) return tokenErr;
    }

    // Actions that require teacher token authentication
    var teacherActions = ['teacherDashboard','teacherCreateClass','teacherCloseClass','teacherDeleteClass',
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

      case 'bohanSiteAuth':
        return handleBohanSiteAuth(p);

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
        quotas: decodeSessionQuotas(data[i][11], data[i][12], data[i][5]),
        // Defensive read — see handleGetSessionInfo comment.
        responsibleExaminer: String((data[i].length > 13 ? data[i][13] : '') || ''),
        defaultPopulation: String((data[i].length > 14 ? data[i][14] : '') || ''),
        managerPhone: sitesMap[siteName] ? sitesMap[siteName].managerPhone : ''
      });
    }
  }
  // Return up to 20 most recent sessions
  return jsonResponse({ status: 'ok', sessions: sessions.slice(0, 20) });
}

// Commander-only: return every active, non-expired session across all examiners.
// Used by the commander UI to load and inspect/correct results in another
// examiner's session (audit / appeal-committee workflow).
// Center-commander aggregate report across multiple sites.
// Role 'מפקד מרכז' has read-only access — cannot enter sessions, cannot correct
// results. Just sees aggregated stats across their assigned sites (column K of
// בוחנים). Date range optional (defaults to today). Three categories: overall,
// per-site, per-license.
function handleCenterManagerReport(p) {
  if (!verifyToken(p.examinerId, p.token)) {
    return jsonResponse({ status: 'error', message: 'טוקן לא תקין', tokenExpired: true });
  }
  var role = getExaminerRole(p.examinerId);
  // 'מפקד קד״ץ' shares the same dashboard as 'מפקד מרכז' — both are
  // multi-site read-only commander roles, gated only on the site list in
  // column K. Add new commander roles here to give them the same view.
  if (role !== 'מפקד מרכז' && !isKdtzRole(role)) {
    return jsonResponse({ status: 'error', message: 'פעולה זו זמינה רק למפקד' });
  }
  var managedSites = getExaminerManagedSites(p.examinerId);
  if (!managedSites.length) {
    return jsonResponse({ status: 'error', message: 'לא הוקצו אתרים מנוהלים — פנה למנהל המערכת' });
  }
  // Normalise site names: strip all whitespace + lowercase for forgiving match
  // (handles "ב.ה. 6910" vs "ב.ה.6910" vs " ב.ה. 6910 " — common manual-entry drift).
  function normalizeSiteName(s) {
    return String(s || '').replace(/\s+/g, '').toLowerCase();
  }
  var sitesNormalized = {};
  for (var s = 0; s < managedSites.length; s++) {
    var ns = normalizeSiteName(managedSites[s]);
    if (ns) sitesNormalized[ns] = managedSites[s]; // map normalized → display name
  }
  // Diagnostic counters so the UI can show why a report is empty.
  var dbg = { totalRows: 0, inDateRange: 0, statusCancelled: 0, siteMatched: 0, siteMismatched: 0, distinctSitesSeenInRange: {} };

  // Parse date range. Defaults to today (00:00 today → now).
  var dateFrom, dateTo;
  if (p.dateFrom) {
    dateFrom = new Date(p.dateFrom);
    if (isNaN(dateFrom.getTime())) dateFrom = null;
  }
  if (p.dateTo) {
    dateTo = new Date(p.dateTo);
    if (isNaN(dateTo.getTime())) dateTo = null;
    else dateTo.setHours(23, 59, 59, 999);
  }
  if (!dateFrom) {
    dateFrom = new Date();
    dateFrom.setHours(0, 0, 0, 0);
  }
  if (!dateTo) {
    dateTo = new Date();
    dateTo.setHours(23, 59, 59, 999);
  }

  // Walk תוצאות, filter by site IN managed + date range. Skip 'בוטל' (cancelled DQ).
  var sheet = getSheet('תוצאות');
  var rows = sheet.getDataRange().getValues();
  var overall = { total: 0, passed: 0, failed: 0, dq: 0 };
  var bySite = {};
  var byLicense = {};
  // Per-examinee details — needed to render the same rich report style as
  // the site-manager report (KPIs + pie + weak topics + per-examinee table).
  var results = [];
  var examinerExcl = getExaminerExclusion();   // exclude examiners who self-tested as examinees (name or ת.ז.)
  for (var ri = 1; ri < rows.length; ri++) {
    var r = rows[ri];
    dbg.totalRows++;
    var status = String(r[7] || '').trim();
    // Parse row date (column A is "DD/MM/YYYY HH:mm" — see todayStr())
    var rawDate = r[0];
    var rowDate = null;
    if (rawDate instanceof Date) rowDate = rawDate;
    else if (rawDate) {
      var m = String(rawDate).match(/(\d{1,2})\/(\d{1,2})\/(\d{4})\s+(\d{1,2}):(\d{2})/);
      if (m) rowDate = new Date(+m[3], (+m[2]) - 1, +m[1], +m[4], +m[5]);
    }
    if (!rowDate) continue;
    if (rowDate < dateFrom || rowDate > dateTo) continue;
    dbg.inDateRange++;
    if (status === 'בוטל') { dbg.statusCancelled++; continue; }
    var rowSite = String(r[10] || '').trim();
    if (isTestSite(rowSite)) continue;   // system-test site — exclude from the manager report stats
    if (isExaminerSelfTest(r[2], r[1], examinerExcl)) continue;   // examiner self-testing (name or ת.ז.) — exclude
    // Track every distinct site we see in range so the commander can see
    // exactly what site names appear in the sheet vs what they configured.
    if (rowSite) dbg.distinctSitesSeenInRange[rowSite] = (dbg.distinctSitesSeenInRange[rowSite] || 0) + 1;
    var rowSiteNorm = normalizeSiteName(rowSite);
    var matchedDisplay = sitesNormalized[rowSiteNorm];
    if (!matchedDisplay) { dbg.siteMismatched++; continue; }
    dbg.siteMatched++;
    // Use the configured display name so aggregation is consistent
    var siteKey = matchedDisplay;

    var rowLic = String(r[4] || '').trim() || '-';
    var isDQ = (status === 'פסול');
    var isPassed = (status === 'עבר');
    overall.total++;
    if (isDQ) overall.dq++;
    else if (isPassed) overall.passed++;
    else overall.failed++;

    if (!bySite[siteKey]) bySite[siteKey] = { site: siteKey, total: 0, passed: 0, failed: 0, dq: 0 };
    bySite[siteKey].total++;
    if (isDQ) bySite[siteKey].dq++;
    else if (isPassed) bySite[siteKey].passed++;
    else bySite[siteKey].failed++;

    // Capture per-examinee row for the rich report
    results.push({
      date: r[0],
      idNumber: r[1],
      name: r[2],
      phone: r[3],
      license: r[4],
      score: r[5],
      percent: r[6],
      passed: r[7],
      time: r[8],
      examiner: r[9],
      site: siteKey,
      classroom: r[11],
      language: r[12],
      attempt: r[14],
      wrongDetails: r[15],
      disqualified: r[17],
      population: r[19] || '',
      corrected: r[20] || false,
      audioMode: r[21] || 'off'
    });

    if (!byLicense[rowLic]) byLicense[rowLic] = { license: rowLic, total: 0, passed: 0, failed: 0, dq: 0 };
    byLicense[rowLic].total++;
    if (isDQ) byLicense[rowLic].dq++;
    else if (isPassed) byLicense[rowLic].passed++;
    else byLicense[rowLic].failed++;
  }

  // Ensure every managed site appears in bySite (even with zero rows) so the
  // commander can spot missing data instead of being confused by absence.
  for (var ms = 0; ms < managedSites.length; ms++) {
    var name = managedSites[ms];
    if (isTestSite(name)) continue;   // never surface the system-test site, even as a zero row
    if (!bySite[name]) bySite[name] = { site: name, total: 0, passed: 0, failed: 0, dq: 0 };
  }

  function pct(part, whole) { return whole > 0 ? Math.round((part / whole) * 100) : 0; }
  overall.passRate = pct(overall.passed, overall.total);

  var bySiteArr = [];
  for (var sk in bySite) {
    var bsr = bySite[sk];
    bsr.passRate = pct(bsr.passed, bsr.total);
    bySiteArr.push(bsr);
  }
  bySiteArr.sort(function(a, b) { return a.site.localeCompare(b.site, 'he'); });

  var byLicArr = [];
  for (var lk in byLicense) {
    var blr = byLicense[lk];
    blr.passRate = pct(blr.passed, blr.total);
    byLicArr.push(blr);
  }
  // Sort by typical license order: B, 1, C1, C, D, other
  var licOrder = { 'B': 1, '1': 2, 'C1': 3, 'C': 4, 'D': 5 };
  byLicArr.sort(function(a, b) {
    var oa = licOrder[a.license] || 99, ob = licOrder[b.license] || 99;
    return oa - ob || a.license.localeCompare(b.license);
  });

  // Convert distinct-sites-seen map → sorted array for display
  var seenArr = [];
  for (var ds in dbg.distinctSitesSeenInRange) {
    seenArr.push({ site: ds, count: dbg.distinctSitesSeenInRange[ds] });
  }
  seenArr.sort(function(a, b) { return b.count - a.count; });

  return jsonResponse({
    status: 'ok',
    managedSites: managedSites,
    dateFrom: dateFrom.toISOString(),
    dateTo: dateTo.toISOString(),
    overall: overall,
    bySite: bySiteArr,
    byLicense: byLicArr,
    results: results,
    // Diagnostic info shown when the report is empty — helps identify the
    // cause (wrong site name in column K, no exams in date range, etc.)
    diagnostics: {
      totalRowsInSheet: dbg.totalRows,
      rowsInDateRange: dbg.inDateRange,
      rowsCancelled: dbg.statusCancelled,
      rowsMatchedSite: dbg.siteMatched,
      rowsMismatchedSite: dbg.siteMismatched,
      sitesSeenInRange: seenArr,
      configuredSites: managedSites
    }
  });
}

function handleListAllSessions(p) {
  // Token already verified upstream (in examinerActions allowlist). Add a role
  // check here since the action isn't restricted by ownership.
  var role = getExaminerRole(p.examinerId);
  if (role !== 'מפקד') {
    return jsonResponse({ status: 'error', message: 'פעולה זו זמינה רק למפקדים' });
  }
  var sheet = getSheet('סשנים');
  var data = sheet.getDataRange().getValues();
  var sitesSheet = getSheet('אתרים');
  var sitesData = sitesSheet.getDataRange().getValues();
  var sitesMap = {};
  for (var s = 1; s < sitesData.length; s++) {
    sitesMap[String(sitesData[s][0]).trim()] = { managerPhone: sitesData[s][2] || '' };
  }
  var now = new Date();
  var sessions = [];
  for (var i = data.length - 1; i >= 1; i--) {
    var active = data[i][10] === true || String(data[i][10]).toUpperCase() === 'TRUE';
    if (!active) continue;
    var validUntil = data[i][9] ? new Date(data[i][9]) : null;
    if (validUntil && now > validUntil) continue;
    var siteName = String(data[i][3] || '').trim();
    sessions.push({
      code: String(data[i][0]),
      examinerId: normalizeId(data[i][1]),
      examinerName: data[i][2] || '',
      site: data[i][3] || '',
      classroom: data[i][4] || '',
      license: data[i][5] || '',
      language: data[i][6] || 'he',
      audioMode: data[i][7] || 'off',
      created: data[i][8] || '',
      validUntil: data[i][9] || '',
      active: true,
      quotas: decodeSessionQuotas(data[i][11], data[i][12], data[i][5]),
      managerPhone: sitesMap[siteName] ? sitesMap[siteName].managerPhone : ''
    });
  }
  // Cap response size — newest first (we already iterate in reverse)
  return jsonResponse({ status: 'ok', sessions: sessions.slice(0, 100) });
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

  // Column L: per-license quotas, stored as JSON. Array of rows like:
  //   [{license:'B', requested:20, approved:18}, {license:'C1', requested:5, approved:5}]
  // Mirrors the plan table in the examiner report — one quota row per license.
  // Column M is reserved (was approvedCount in the previous single-pair design;
  // kept blank now to leave room for future extension without renumbering).
  var quotas = parseAndValidateQuotas(p.quotas);
  if (quotas.error) {
    return jsonResponse({ status: 'error', message: quotas.error });
  }

  // Column N (13): בוחן אחראי — name of the senior/responsible examiner when
  // multiple examiners work the same site/day per the פקודת עבודה. When the
  // session is opened by a solo examiner this can equal the examiner himself,
  // or be left blank if he's the responsible. The Rav-Bochen / commander
  // reports surface this field so the chain of responsibility matches the
  // physical staffing on the ground.
  var responsibleExaminer = String(p.responsibleExaminer || '').trim();

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
    true,
    JSON.stringify(quotas.rows),
    '',
    responsibleExaminer,
    String(p.defaultPopulation || '').trim()   // O (idx 14) = default population for the session
  ]);

  return jsonResponse({
    status: 'ok',
    sessionCode: code,
    validUntil: validUntil.toISOString(),
    examinerName: examinerName,
    responsibleExaminer: responsibleExaminer
  });
}

// Returns every session that took place at the same site on the same day as
// the caller's sessionCode, plus every result row belonging to those sessions.
// Powers the "דו"ח משותף לאתר" (combined site report) button on the examiner
// dashboard — when two examiners share a site, the Rav-Bochen wants one
// report with both their sessions side by side.
//
// Authorization: the caller must be the responsible examiner of at least one
// of the sessions, OR hold a commander role. Regular examiners who happen to
// have a session at the same site can still see results for their own session
// via the existing examinerDashboard — the combined view is a chain-of-
// command audit lens.
function handleSiteCombinedReport(p) {
  if (!verifyToken(p.examinerId, p.token)) {
    return jsonResponse({ status: 'error', message: 'טוקן לא תקין', tokenExpired: true });
  }

  var sessSheet = getSheet('סשנים');
  var sessData = sessSheet.getDataRange().getValues();

  // Locate the calling session to discover its site + date.
  var anchorSite = '';
  var anchorDate = null;
  var callerSessionCode = String(p.sessionCode || '').trim();
  for (var i = 1; i < sessData.length; i++) {
    if (String(sessData[i][0]).trim() === callerSessionCode) {
      anchorSite = String(sessData[i][3] || '').trim();
      anchorDate = sessData[i][8] instanceof Date ? sessData[i][8] : new Date(sessData[i][8]);
      break;
    }
  }
  if (!anchorSite || !anchorDate || isNaN(anchorDate.getTime())) {
    return jsonResponse({ status: 'error', message: 'סשן לא נמצא או חסר תאריך' });
  }

  var dayStart = new Date(anchorDate);
  dayStart.setHours(0, 0, 0, 0);
  var dayEnd = new Date(dayStart);
  dayEnd.setDate(dayEnd.getDate() + 1);

  // Pull every session at the same (site, day). We accept sessions in any
  // status (active / closed / expired) — the combined report is a historical
  // record, not a live view.
  var sessions = [];
  var callerExaminerName = '';
  var callerIsResponsibleOnAny = false;
  for (var j = 1; j < sessData.length; j++) {
    var site = String(sessData[j][3] || '').trim();
    if (site !== anchorSite) continue;
    var created = sessData[j][8] instanceof Date ? sessData[j][8] : new Date(sessData[j][8]);
    if (isNaN(created.getTime()) || created < dayStart || created >= dayEnd) continue;
    var responsible = String((sessData[j].length > 13 ? sessData[j][13] : '') || '').trim();
    var examinerName = String(sessData[j][2] || '').trim();
    if (normalizeId(sessData[j][1]) === normalizeId(p.examinerId)) {
      callerExaminerName = examinerName;
      if (responsible && responsible === examinerName) callerIsResponsibleOnAny = true;
    }
    sessions.push({
      code: String(sessData[j][0]).trim(),
      examinerId: String(sessData[j][1] || '').trim(),
      examinerName: examinerName,
      site: site,
      classroom: String(sessData[j][4] || '').trim(),
      license: String(sessData[j][5] || '').trim(),
      language: String(sessData[j][6] || '').trim(),
      audioMode: String(sessData[j][7] || '').trim(),
      created: created.toISOString(),
      responsibleExaminer: responsible,
      quotas: decodeSessionQuotas(sessData[j][11], sessData[j][12], sessData[j][5])
    });
  }
  if (sessions.length === 0) {
    return jsonResponse({ status: 'error', message: 'לא נמצאו סשנים תואמים' });
  }

  // Authorization: pass if caller is responsible-on-some-session OR commander.
  var role = getExaminerRole(p.examinerId);
  var isCommander = (
    role === 'מפקד' || role === 'מפקד מקומי' || role === 'מפקד ראשי' ||
    role === 'מפקד מרכז' || isKdtzRole(role) || role === 'רב בוחן'
  );
  // Also: caller is the responsible examiner of any session in the set,
  // even if their own session doesn't list themselves as responsible.
  var callerNamedAsResponsibleAnywhere = false;
  if (callerExaminerName) {
    for (var s = 0; s < sessions.length; s++) {
      if (sessions[s].responsibleExaminer === callerExaminerName) {
        callerNamedAsResponsibleAnywhere = true;
        break;
      }
    }
  }
  if (!isCommander && !callerIsResponsibleOnAny && !callerNamedAsResponsibleAnywhere) {
    return jsonResponse({
      status: 'error',
      message: 'הדו"ח המשותף זמין רק לבוחן האחראי או למפקד'
    });
  }

  // Now pull all results that belong to any of these session codes.
  var sessionCodesSet = {};
  for (var sc = 0; sc < sessions.length; sc++) sessionCodesSet[sessions[sc].code] = true;

  var resSheet = getSheet('תוצאות');
  var resData = resSheet.getDataRange().getValues();
  var results = [];
  for (var r = 1; r < resData.length; r++) {
    var sCode = String(resData[r][13] || '').trim();
    if (!sessionCodesSet[sCode]) continue;
    if (String(resData[r][7] || '').trim() === 'בוטל') continue; // skip overturned/superseded rows (consistent with the other report handlers)
    var rDate = resData[r][0] instanceof Date ? resData[r][0] : new Date(resData[r][0]);
    results.push({
      date: rDate.toISOString(),
      idNumber: String(resData[r][1] || ''),
      name: String(resData[r][2] || ''),
      phone: String(resData[r][3] || ''),
      license: String(resData[r][4] || ''),
      score: resData[r][5],
      percent: resData[r][6],
      passed: String(resData[r][7] || ''),
      time: String(resData[r][8] || ''),
      examiner: String(resData[r][9] || ''),
      site: String(resData[r][10] || ''),
      classroom: String(resData[r][11] || ''),
      language: String(resData[r][12] || ''),
      sessionCode: sCode,
      attemptNum: resData[r][14],
      population: String(resData[r][19] || ''),
      disqualified: resData[r][17] === true || String(resData[r][17]).toUpperCase() === 'TRUE',
      audioMode: String(resData[r][21] || ''),
      verified: (resData[r].length > 22) ? String(resData[r][22] || '') : '',
      device: (resData[r].length > 29) ? String(resData[r][29] || '') : ''
    });
  }

  return jsonResponse({
    status: 'ok',
    site: anchorSite,
    dayStart: dayStart.toISOString(),
    sessions: sessions,
    results: results
  });
}

// Returns a sorted list of active examiners' names, used by the session-create
// dropdown to pick the בוחן אחראי. Only sends `name` — IDs/roles/tokens
// don't belong on the client. Active = column D in 'בוחנים' is exactly 'כן'
// (the same truthiness check used elsewhere).
//
// RESPONSIBLE_EXAMINER_HIDE_LIST: names to omit from this dropdown even when
// they're marked active in the sheet. Use when an examiner is still active in
// the system (can log in, see their dashboard) but shouldn't be selectable as
// a "responsible examiner" on the work order. Comparison is normalized
// (trim + lowercase + collapsed whitespace) so casing/spacing variants match.
var RESPONSIBLE_EXAMINER_HIDE_LIST = [
  'תומר לוי',
  'אביאור שמעוני',
  'דוד בטיטו'
];
function _normalizeNameForHideList(s) {
  return String(s || '').trim().replace(/\s+/g, ' ').toLowerCase();
}

function handleListActiveExaminers(p) {
  var sheet = getSheet('בוחנים');
  var data = sheet.getDataRange().getValues();
  var hideSet = {};
  for (var hi = 0; hi < RESPONSIBLE_EXAMINER_HIDE_LIST.length; hi++) {
    hideSet[_normalizeNameForHideList(RESPONSIBLE_EXAMINER_HIDE_LIST[hi])] = true;
  }
  var names = [];
  for (var i = 1; i < data.length; i++) {
    var active = data[i][3];
    var isActive = (active === true) || (active === 'כן') || (String(active).toUpperCase() === 'TRUE');
    if (!isActive) continue;
    var name = String(data[i][0] || '').trim();
    if (!name) continue;
    if (hideSet[_normalizeNameForHideList(name)]) continue;
    names.push(name);
  }
  names.sort(function(a, b) { return a.localeCompare(b, 'he'); });
  return jsonResponse({ status: 'ok', examiners: names });
}

// Validates the JSON quotas payload sent from the examiner UI. Returns either
// { rows: [...] } on success or { error: 'msg' }. The same checks are mirrored
// in examiner.html createSessionBtn — kept in sync so a forged client still
// fails server-side.
var QUOTA_VALID_LICENSES = { B:1, '1':1, C1:1, C:1, D:1 };
function parseAndValidateQuotas(raw) {
  if (!raw) return { error: 'יש להזין כמויות נבחנים לפי דרגה' };
  var parsed;
  try { parsed = JSON.parse(raw); } catch(e) { return { error: 'מבנה כמויות לא תקין' }; }
  if (!Array.isArray(parsed) || parsed.length === 0) {
    return { error: 'יש להזין לפחות שורת כמויות אחת' };
  }
  var seen = {};
  var clean = [];
  for (var i = 0; i < parsed.length; i++) {
    var r = parsed[i] || {};
    var lic = String(r.license || '').trim();
    var site = String(r.site || '').trim();  // '' = host site (backward compatible)
    var req = parseInt(r.requested, 10);
    var appr = parseInt(r.approved, 10);
    if (!QUOTA_VALID_LICENSES[lic]) {
      return { error: 'דרגה לא חוקית בשורה ' + (i + 1) };
    }
    // Uniqueness is per (site, license): the same license may appear once per
    // site (host + guest) but not twice for the same site.
    var _qkey = site + '|' + lic;
    if (seen[_qkey]) {
      return { error: 'דרגה "' + lic + '" מופיעה יותר מפעם אחת' + (site ? ' לאתר "' + site + '"' : '') };
    }
    seen[_qkey] = true;
    // Quantity is OPTIONAL (mirrors the client redesign e0708f2, which removed the
    // requested/approved fields from session opening). A missing/blank/0 quantity
    // defaults to 0 instead of blocking session creation — the examiner no longer
    // has to type a number to open an exam.
    if (!isFinite(req) || req < 0) req = 0;
    if (!isFinite(appr) || appr < 0) appr = 0;
    if (appr > req) appr = req;
    clean.push({ site: site, license: lic, requested: req, approved: appr });
  }
  return { rows: clean };
}

// Decode column L into an array of quota rows. Handles three historical shapes:
//   1. Empty cell           → []
//   2. Plain number         → [{license: <session.license>, requested: <num>, approved: <colM>}]
//      (early prototype that stored requested/approved as separate columns L,M)
//   3. JSON array string    → parsed array
// Used by every session reader so backward-compat is centralised.
function decodeSessionQuotas(colL, colM, sessionLicense) {
  if (colL === '' || colL == null) return [];
  var s = String(colL).trim();
  if (s.charAt(0) === '[') {
    try {
      var arr = JSON.parse(s);
      if (Array.isArray(arr)) {
        var out = [];
        for (var i = 0; i < arr.length; i++) {
          var r = arr[i] || {};
          out.push({
            site: String(r.site || ''),
            license: String(r.license || ''),
            requested: Number(r.requested) || 0,
            approved: Number(r.approved) || 0
          });
        }
        return out;
      }
    } catch(e) {}
    return [];
  }
  // Legacy single-pair format: column L = requested, column M = approved
  var legacyReq = parseInt(s, 10);
  var legacyAppr = parseInt(colM, 10);
  if (isFinite(legacyReq) && legacyReq > 0) {
    return [{
      site: '',
      license: String(sessionLicense || 'B'),
      requested: legacyReq,
      approved: isFinite(legacyAppr) ? legacyAppr : 0
    }];
  }
  return [];
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
      var cleanup = cleanupStuckDisqualified(searchCode);
      return jsonResponse({ status: 'ok', cleanup: cleanup });
    }
  }
  return jsonResponse({ status: 'error', message: 'סשן לא נמצא' });
}

// Triggered from handleCloseSession. Sweeps pending rows in this session whose
// status got stuck on 'disqualified' without a real result, and moves them to a
// terminal status so the session closes clean.
//   no result row    → 'cancelled' (DQ fired but nothing was ever recorded)
//   latest is 'בוטל' → 'completed' (a result existed but was overturned)
// Rows whose latest result is 'פסול'/'עבר'/'נכשל' are left as 'disqualified' —
// those are real outcomes awaiting examiner confirm/overturn.
function cleanupStuckDisqualified(sessionCode) {
  var pendSheet = getSheet('ממתינים');
  var pendData = pendSheet.getDataRange().getValues();
  var stuck = [];
  for (var i = 1; i < pendData.length; i++) {
    if (String(pendData[i][0]) === String(sessionCode) &&
        String(pendData[i][5] || '').trim() === 'disqualified') {
      stuck.push({ rowIdx: i, idKey: normalizeId(pendData[i][1]) });
    }
  }
  if (stuck.length === 0) return { cancelled: 0, completed: 0, skipped: 0 };

  var resSheet = getSheet('תוצאות');
  var resData = resSheet.getDataRange().getValues();
  var latestByExaminee = {};
  for (var r = 1; r < resData.length; r++) {
    if (String(resData[r][13]) !== String(sessionCode)) continue;
    // resData is in append order; later row wins as "latest"
    latestByExaminee[normalizeId(resData[r][1])] = String(resData[r][7] || '').trim();
  }

  var cancelled = 0, completed = 0, skipped = 0;
  for (var k = 0; k < stuck.length; k++) {
    var latest = latestByExaminee[stuck[k].idKey];
    if (!latest) {
      pendSheet.getRange(stuck[k].rowIdx + 1, 6).setValue('cancelled');
      cancelled++;
    } else if (latest === 'בוטל') {
      pendSheet.getRange(stuck[k].rowIdx + 1, 6).setValue('completed');
      completed++;
    } else {
      skipped++;
    }
  }
  if (cancelled || completed) SpreadsheetApp.flush();
  return { cancelled: cancelled, completed: completed, skipped: skipped };
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
      var _siQuotas = decodeSessionQuotas(data[i][11], data[i][12], data[i][5]);
      // Build the distinct site list (host first, then guest sites declared in
      // the quotas). Quota rows with an empty site belong to the host (column D).
      // The examinee picks from this list when more than one site exists.
      var _hostSite = String(data[i][3] || '').trim();
      var _siteSeen = {};
      var _sites = [];
      if (_hostSite) { _sites.push(_hostSite); _siteSeen[_hostSite] = true; }
      for (var _sq = 0; _sq < _siQuotas.length; _sq++) {
        var _sName = String(_siQuotas[_sq].site || '').trim() || _hostSite;
        if (_sName && !_siteSeen[_sName]) { _siteSeen[_sName] = true; _sites.push(_sName); }
      }
      return jsonResponse({
        status: 'ok',
        session: {
          site: data[i][3],
          sites: _sites,
          classroom: data[i][4],
          license: data[i][5],
          language: data[i][6],
          audioMode: data[i][7],
          examinerName: data[i][2],
          validUntil: data[i][9],
          quotas: _siQuotas,
          // Column N (13) may be missing on rows created before this feature
          // shipped — defensive read returns '' for those, treating them as
          // sessions without a designated responsible examiner.
          responsibleExaminer: String((data[i].length > 13 ? data[i][13] : '') || ''),
          // Default population set by the examiner at session open (col O, idx 14).
          // The examinee's form pre-selects it but can change it. '' on old rows.
          defaultPopulation: String((data[i].length > 14 ? data[i][14] : '') || '')
        }
      });
    }
  }
  return jsonResponse({ status: 'error', message: 'קוד סשן לא תקין' });
}

function handleRegisterExaminee(p) {
  // Rate limit: max 30 registrations per minute per session. Prevents an
  // attacker with the session code from spamming hundreds of fake registrations.
  var rlErr = requireRateLimit('registerExaminee', String(p.sessionCode || ''), 30, 60);
  if (rlErr) return rlErr;
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
        // A PENDING disqualification (anti-cheat fired, examiner hasn't decided)
        // must NOT allow a fresh registration. Re-registering while the DQ was
        // still on the examiner's screen created a SECOND ממתינים row, so the
        // soldier appeared twice in "במבחן כרגע" (incident @ בוחן יניר 2026-06-08).
        // The examiner must first decide — "בטל פסילה" (resume) or "אשר פסילה"
        // (finalize). 'dq_confirmed'/'completed' are intentionally NOT blocked:
        // they're final (a legitimate retake may re-register) and don't surface
        // in "במבחן כרגע".
        if (status === 'disqualified') {
          return jsonResponse({ status: 'error', message: 'יש פסילה הממתינה להחלטת הבוחן — פנה לבוחן לפני רישום מחדש' });
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
  var examineeToken = generateExamineeToken();
  // External-monitor indicator from client (screen.isExtended). Cheating risk
  // signal — examinee may be sharing window to a second screen with accomplice.
  var hasExtendedScreen = (p.hasExtendedScreen === '1' || p.hasExtendedScreen === 1 || p.hasExtendedScreen === true);
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
    '',                       // K (10): הארכת זמן — נקבע ע"י הבוחן בעת אישור
    '',                       // L (11): התחלת מבחן — נקבע ע"י markExamStarted
    examineeToken,            // M (12): טוקן נבחן — מוחזר ללקוח, נדרש בקריאות עוקבות
    0,                        // N (13): ספירת DQ — מתעלה עם כל disqualify
    hasExtendedScreen ? 'כן' : '', // O (14): מסך נוסף — סימן אזהרה
    0,                        // P (15): ספירת אזהרות — מאותחל ל-0 (נכתב ע"י warning)
    '',                       // Q (16): אזהרה אחרונה — נכתב ע"י warning
    p.site || ''              // R (17): אתר — האתר שהנבחן בחר (מארח/אורח), לתצוגה חיה לבוחן
  ]);
  return jsonResponse({ status: 'ok', examineeToken: examineeToken });
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
  // Rate limit: max 60 polls per minute per (sessionCode, idNumber). Normal
  // polling is ~20-30/min, so this gives 2× headroom while blocking floods.
  var rlErr = requireRateLimit('checkApproval', String(p.sessionCode || '') + '_' + normalizeId(p.idNumber), 60, 60);
  if (rlErr) return rlErr;
  var BASE_EXAM_MINUTES = 40;
  var sheet = getSheet('ממתינים');
  var data = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]).trim() === String(p.sessionCode).trim() && normalizeId(data[i][1]) === normalizeId(p.idNumber)) {
      var approval = String(data[i][5] || 'waiting').trim();
      // Skip terminal statuses from previous exams — keep looking for active row
      // Note: dq_confirmed is NOT skipped — examinee needs to receive this status
      //
      // 'rejected' is intentionally on the skip list. Real exam-day incident:
      // two examinees shared an ID number (family), first was rejected at
      // 17:47, second cancelled at 18:05. A third visitor with stale
      // localStorage polled later — the loop skipped the newest cancelled row
      // and returned the older 'rejected' status, showing "הבוחן דחה" on a
      // screen that nobody actually rejected. Skipping rejected here forces
      // the response to "no registration found" when all rows are terminal,
      // which the client interprets as "your saved state is stale, start over".
      //
      // Trade-off: when an examiner rejects a CURRENT registration, the
      // examinee no longer sees an in-app rejection notice — they see "no
      // registration" and reset to the code screen. Acceptable because the
      // examiner is physically next to them and can explain verbally.
      if (approval === 'completed' || approval === 'disqualified' || approval === 'cancelled' || approval === 'rejected') continue;
      // Token check: when a token is stored for this row, reject mismatches.
      // Legacy rows (no stored token) and the very first poll (client may not
      // have echoed the token yet) are accepted so we don't break in-flight
      // registrations during the deploy window.
      var storedToken = String((data[i].length > 12 ? data[i][12] : '') || '').trim();
      if (storedToken && p.examineeToken && String(p.examineeToken).trim() !== storedToken) {
        return jsonResponse({ status: 'error', message: 'טוקן נבחן לא תקין', examineeTokenError: 'mismatch' });
      }
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
  var tokenErr = requireExamineeToken(p);
  if (tokenErr) return tokenErr;
  var sheet = getSheet('ממתינים');
  var data = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) !== String(p.sessionCode) || normalizeId(data[i][1]) !== normalizeId(p.idNumber)) continue;
    var st = String(data[i][5]).trim();
    if (st === 'approved') {
      sheet.getRange(i + 1, 6).setValue('in_exam');
      sheet.getRange(i + 1, 12).setValue(nowISO()); // column L = exam actual start time
      SpreadsheetApp.flush();
      return jsonResponse({ status: 'ok' });
    }
    if (st === 'in_exam') {
      // Idempotent: a retry whose earlier response was lost (common on iOS when
      // the page backgrounds) should still report success so the client stops
      // retrying — and not overwrite the original start time.
      return jsonResponse({ status: 'ok', already: true });
    }
    // Other statuses (completed/cancelled/disqualified): keep scanning for an
    // approved/in_exam row belonging to this examinee.
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

  // Sum of mid-exam time grants per examinee (minutes) — extends the stale/timeout
  // threshold below and is shown as a badge in the active list. One read, by id.
  var extraMinById = {};
  try {
    var extData = getSheet('הארכות זמן').getDataRange().getValues();
    for (var exr = 1; exr < extData.length; exr++) {
      if (String(extData[exr][1]).trim() !== code) continue;
      var exk = normalizeId(extData[exr][2]);
      extraMinById[exk] = (extraMinById[exk] || 0) + (Number(extData[exr][4]) || 0);
    }
  } catch (e) {}

  // Auto-cleanup: detect stale in_exam entries that already have a result or are way past exam time
  var now = new Date();
  var BASE_EXAM_MS = 40 * 60 * 1000;
  var STALE_BUFFER_MS = 20 * 60 * 1000; // 20 minutes buffer (approval wait + instructions)
  for (var ci = 1; ci < pendData.length; ci++) {
    if (String(pendData[ci][0]) !== code) continue;
    // Reconcile stuck 'in_exam' AND 'approved' entries. 'approved' that never
    // advanced to 'in_exam' happens when markExamStarted failed on the device
    // (common on iOS) — leaving the examinee stuck in "ממתינים" forever, even
    // after finishing. We clear those once a result exists for them (below).
    var _ciStatus = String(pendData[ci][5]).trim();
    if (_ciStatus !== 'in_exam' && _ciStatus !== 'approved') continue;
    var _startedExam = (_ciStatus === 'in_exam');
    var ciId = pendData[ci][1];
    var examStart = pendData[ci][11] ? new Date(pendData[ci][11]) : null;
    var regTime = examStart || (pendData[ci][4] ? new Date(pendData[ci][4]) : null);
    // Dynamic stale threshold: exam time (based on extension) + buffer
    var ciExt = parseFloat(pendData[ci][10]) || 1;
    if (ciExt !== 1.25 && ciExt !== 1.5) ciExt = 1;
    var maxMs = Math.round(BASE_EXAM_MS * ciExt) + STALE_BUFFER_MS + ((extraMinById[normalizeId(ciId)] || 0) * 60 * 1000);
    var isStale = regTime && (now.getTime() - regTime.getTime() > maxMs);
    // Only someone who actually STARTED (in_exam) can time out into a fail. A
    // stale 'approved' never started → never fabricate a 0/30 fail for it; it is
    // only reconciled when a real result already exists.
    var effectiveStale = isStale && _startedExam;

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

    if (hasUnmatchedResult || effectiveStale) {
      // Fix dangling status — mark as completed
      pendSheet.getRange(ci + 1, 6).setValue('completed');
      pendData[ci][5] = 'completed'; // update local copy
      if (effectiveStale && !hasUnmatchedResult) {
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

  // Pre-compute attempts-today by examinee id (for "second attempt today" warning).
  // Counts non-disqualified terminal entries (real attempts) made today regardless
  // of which session — so an examinee who tried earlier today in another session
  // also triggers the warning.
  var attemptsTodayById = {};
  var todayDateStr = (function() {
    var d = new Date();
    return d.getFullYear() + '-' + (d.getMonth() + 1) + '-' + d.getDate();
  })();
  function isToday(val) {
    if (!val) return false;
    try {
      var d = (val instanceof Date) ? val : new Date(val);
      if (isNaN(d.getTime())) return false;
      return (d.getFullYear() + '-' + (d.getMonth() + 1) + '-' + d.getDate()) === todayDateStr;
    } catch(_) { return false; }
  }
  for (var ai2 = 1; ai2 < resData.length; ai2++) {
    if (!isToday(resData[ai2][0])) continue;
    var aiPassed = String(resData[ai2][7] || '').trim();
    if (aiPassed === 'בוטל') continue; // overturned, not a real attempt
    var aiId = normalizeId(resData[ai2][1]);
    attemptsTodayById[aiId] = (attemptsTodayById[aiId] || 0) + 1;
  }

  // Build pending (waiting/approved) and active (in_exam/disqualified) lists,
  // DEDUPED per examinee. A soldier must appear ONCE in each list even when the
  // ממתינים sheet holds duplicate rows for them (re-registration after a stuck
  // row, mid-incident states). Without this the examiner saw the same person
  // two/three times in "במבחן כרגע" (reported incident @ בוחן יניר). Dedup keys
  // on normalized id within this session. Collapse priority:
  //   • pending: latest row wins.
  //   • active: a 'disqualified' (needs-decision) row beats an 'in_exam' row;
  //     within the same status, the latest row wins.
  var pendingById = {};
  var activeById = {};
  for (var i = 1; i < pendData.length; i++) {
    if (String(pendData[i][0]) !== code) continue;
    var s = String(pendData[i][5] || '').trim();
    if (s !== 'waiting' && s !== 'approved' && s !== 'in_exam' && s !== 'disqualified') continue;
    var dqCount = (pendData[i].length > 13) ? (Number(pendData[i][13]) || 0) : 0;
    var hasExtScreen = (pendData[i].length > 14) ? (String(pendData[i][14] || '').trim() === 'כן') : false;
    var warnCount = (pendData[i].length > 15) ? (Number(pendData[i][15]) || 0) : 0;
    var lastWarn = (pendData[i].length > 16) ? String(pendData[i][16] || '') : '';
    var idNorm = normalizeId(pendData[i][1]);
    var item = { idNumber: pendData[i][1], name: pendData[i][2], phone: pendData[i][3], time: pendData[i][4], examStartTime: pendData[i][11] || '', status: s, language: pendData[i][6] || '', population: pendData[i][7] || '', site: (pendData[i].length > 17) ? (pendData[i][17] || '') : '', license: pendData[i][8] || '', audioMode: pendData[i][9] || 'off', timeExtension: String(pendData[i][10] || ''), dqCount: dqCount, warnings: warnCount, lastWarning: lastWarn, attemptsToday: attemptsTodayById[idNorm] || 0, hasExtendedScreen: hasExtScreen, extraMinutes: extraMinById[idNorm] || 0, finishedOnDevice: (pendData[i].length > 18 ? !!pendData[i][18] : false) };
    if (s === 'waiting' || s === 'approved') {
      pendingById[idNorm] = item; // ascending loop → latest row wins
    } else {
      // Surface DQ events the examiner must decide on (in_exam + disqualified).
      if (s === 'disqualified') item.dqPending = true;
      var prevA = activeById[idNorm];
      if (!prevA) {
        activeById[idNorm] = item;
      } else {
        // 'disqualified' (needs decision) beats 'in_exam'; same status → latest wins.
        var curDQ = (s === 'disqualified');
        var prevDQ = (prevA.status === 'disqualified');
        if (curDQ || !prevDQ) activeById[idNorm] = item;
      }
    }
  }
  for (var pkA in pendingById) pending.push(pendingById[pkA]);
  for (var akA in activeById) active.push(activeById[akA]);

  // Re-read resData in case cleanup added new results
  resData = resSheet.getDataRange().getValues();
  // DEDUP results per examinee: the תוצאות sheet can end up with several
  // non-בוטל rows for one (session, id) when recovery paths (timeout-fail,
  // manual force-complete, disqualify) appended rows that weren't superseded.
  // The examiner must see each soldier ONCE — keep only the LATEST row, which
  // matches the system's own canonical rule (every supersede appends the newest
  // and marks older ones בוטל; latest-wins is the safety net when that didn't run).
  var latestResRowById = {};
  for (var jd = 1; jd < resData.length; jd++) {
    if (String(resData[jd][13]) !== code) continue;
    if (String(resData[jd][7] || '') === 'בוטל') continue;
    latestResRowById[normalizeId(resData[jd][1])] = jd; // ascending → ends as latest
  }
  var completed = [];
  for (var j = 1; j < resData.length; j++) {
    if (String(resData[j][13]) !== code) continue;
    if (String(resData[j][7] || '') === 'בוטל') continue;
    if (latestResRowById[normalizeId(resData[j][1])] !== j) continue; // keep latest only
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
      audioMode: resData[j][21] || 'off',
      // Integrity flags the server already computes & stores but the dashboard
      // never showed: verified='מאומת' when the score was re-computed against the
      // trusted answer key; suspicious='חשוד' when the exam took <3 min. Surfacing
      // these lets the examiner spot any result that was NOT server-verified
      // (missing answer key, missing exam-registration, or a tampered/forged
      // submit) instead of it looking identical to a clean pass.
      verified: (resData[j].length > 22) ? (resData[j][22] || '') : '',
      suspicious: (resData[j].length > 23) ? (resData[j][23] || '') : '',
      device: (resData[j].length > 29) ? (resData[j][29] || '') : ''
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

// Lightweight WARNING counter (suspicious-but-not-DQ events: tab/app-switch
// warning, split-screen detected). The examinee client reports each warning so the
// examiner dashboard can surface repeated suspicious behavior even when it never
// reached a full disqualification. Examinee-token gated + rate-limited; best-effort
// — a failed report never affects the exam. Stored in ממתינים col 16 (idx15) =
// count, col 17 (idx16) = last reason.
function handleReportWarning(p) {
  if (!p.sessionCode || !p.idNumber) return jsonResponse({ status: 'error', message: 'חסר מזהה' });
  var rlErr = requireRateLimit('reportWarning', String(p.sessionCode || '') + '_' + normalizeId(p.idNumber), 30, 60);
  if (rlErr) return rlErr;
  var tokenCheck = verifyExamineeToken(p.sessionCode, p.idNumber, p.examineeToken);
  if (!tokenCheck.valid) return jsonResponse({ status: 'error', message: 'טוקן נבחן לא תקין', examineeTokenError: tokenCheck.reason });
  try {
    var sheet = getSheet('ממתינים');
    var data = sheet.getDataRange().getValues();
    for (var i = data.length - 1; i >= 1; i--) {
      if (String(data[i][0]) === String(p.sessionCode) && normalizeId(data[i][1]) === normalizeId(p.idNumber)) {
        var st = String(data[i][5] || '').trim();
        if (st === 'in_exam' || st === 'approved') {
          var prev = (data[i].length > 15) ? (Number(data[i][15]) || 0) : 0;
          sheet.getRange(i + 1, 16).setValue(prev + 1);                                   // col 16 (idx15) = warnings count
          if (p.reason) sheet.getRange(i + 1, 17).setValue(String(p.reason).slice(0, 40)); // col 17 (idx16) = last reason
        }
        break;
      }
    }
  } catch(e) {}
  return jsonResponse({ status: 'ok' });
}

// Raw status probe for the examinee DURING the exam. handleCheckApproval can't be
// reused — it deliberately SKIPS 'disqualified'. This returns the live ממתינים
// status so an examiner-initiated disqualification is reflected on the examinee's
// device; until now the exam ran locally and the examinee never knew they were DQ'd.
function handleGetExamStatus(p) {
  if (!p.sessionCode || !p.idNumber) return jsonResponse({ status: 'error', message: 'חסר מזהה' });
  var rlErr = requireRateLimit('getExamStatus', String(p.sessionCode || '') + '_' + normalizeId(p.idNumber), 60, 60);
  if (rlErr) return rlErr;
  var sheet = getSheet('ממתינים');
  var data = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]).trim() === String(p.sessionCode).trim() && normalizeId(data[i][1]) === normalizeId(p.idNumber)) {
      var storedToken = String((data[i].length > 12 ? data[i][12] : '') || '').trim();
      if (storedToken && p.examineeToken && String(p.examineeToken).trim() !== storedToken) {
        return jsonResponse({ status: 'error', examineeTokenError: 'mismatch' });
      }
      return jsonResponse({ status: 'ok', examStatus: String(data[i][5] || '').trim(), extraMinutes: sumExtraMinutes(p.sessionCode, p.idNumber) });
    }
  }
  return jsonResponse({ status: 'ok', examStatus: 'not_found' });
}

// ===== Mid-exam time addition (security evacuation / technical / medical) =====
// The examiner grants extra minutes to a RUNNING exam. Every grant is appended to
// the 'הארכות זמן' audit sheet with a mandatory reason, so the record is preserved.
// getExamStatus + the dashboard read the SUM of grants per examinee:
//   - the examinee extends examDeadline (idempotently: start + base + sum)
//   - the dashboard pushes back the stale/timeout-fail threshold by the same sum
// Examiner-authenticated only (mirrors handleDisqualify path A).
function sumExtraMinutes(sessionCode, idNumber) {
  try {
    var d = getSheet('הארכות זמן').getDataRange().getValues();
    var total = 0;
    for (var i = 1; i < d.length; i++) {
      if (String(d[i][1]).trim() === String(sessionCode).trim() && normalizeId(d[i][2]) === normalizeId(idNumber)) {
        total += Number(d[i][4]) || 0;
      }
    }
    return total;
  } catch (e) { return 0; }
}

function handleAddExamTime(p) {
  if (!p.sessionCode || !p.idNumber) return jsonResponse({ status: 'error', message: 'חסר מזהה' });
  // Examiner auth — must hold a valid token AND own the session (same as DQ).
  if (!verifyToken(p.examinerId, p.token)) {
    return jsonResponse({ status: 'error', message: 'טוקן בוחן לא תקין', tokenExpired: true });
  }
  if (!verifyExaminerForSession(p.sessionCode, p.examinerId)) {
    return jsonResponse({ status: 'error', message: 'אין הרשאה — בוחן לא תואם לסשן' });
  }
  var minutes = Math.round(Number(p.minutes) || 0);
  if (!(minutes > 0) || minutes > 180) {
    return jsonResponse({ status: 'error', message: 'מספר דקות לא תקין' });
  }
  var reason = String(p.reason || '').trim();
  if (!reason) return jsonResponse({ status: 'error', message: 'חובה לציין סיבה' });

  // Confirm the examinee exists in this session and grab their name for the audit row.
  var pendData = getSheet('ממתינים').getDataRange().getValues();
  var name = '', found = false;
  for (var j = pendData.length - 1; j >= 1; j--) {
    if (String(pendData[j][0]).trim() === String(p.sessionCode).trim() && normalizeId(pendData[j][1]) === normalizeId(p.idNumber)) {
      name = pendData[j][2] || '';
      found = true;
      break;
    }
  }
  if (!found) return jsonResponse({ status: 'error', message: 'נבחן לא נמצא בסשן' });

  // Examiner display name for the audit row.
  var examinerName = '';
  try {
    var sData = getSheet('סשנים').getDataRange().getValues();
    for (var s = 1; s < sData.length; s++) {
      if (String(sData[s][0]).trim() === String(p.sessionCode).trim()) { examinerName = sData[s][2] || ''; break; }
    }
  } catch (e) {}

  getSheet('הארכות זמן').appendRow([new Date(), p.sessionCode, p.idNumber, name, minutes, reason, examinerName]);

  return jsonResponse({ status: 'ok', addedMinutes: minutes, totalExtraMinutes: sumExtraMinutes(p.sessionCode, p.idNumber) });
}

// The examinee's device reports it FINISHED the exam — a tiny keepalive ping fired at
// submit time, separate from the heavier (retried) result POST. On a weak connection the
// ping often lands even when the full result is still syncing, so the examiner sees
// "finished — syncing result" instead of mistaking a finished examinee for one who is
// still testing and forcing a needless redo. Stamps the in_exam row (col 19 = סיים במכשיר);
// the row IS the attempt, so the flag is naturally scoped to this attempt (a retake is a
// new row) and becomes irrelevant once the result lands (the row flips to completed).
function handleMarkFinished(p) {
  if (!p.sessionCode || !p.idNumber) return jsonResponse({ status: 'error', message: 'חסר מזהה' });
  var pendSheet = getSheet('ממתינים');
  var data = pendSheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]).trim() === String(p.sessionCode).trim() && normalizeId(data[i][1]) === normalizeId(p.idNumber)) {
      var storedToken = String((data[i].length > 12 ? data[i][12] : '') || '').trim();
      if (storedToken && p.examineeToken && String(p.examineeToken).trim() !== storedToken) {
        return jsonResponse({ status: 'error', examineeTokenError: 'mismatch' });
      }
      if (String(data[i][5]).trim() === 'in_exam') {
        if (pendSheet.getMaxColumns() < 19) pendSheet.insertColumnsAfter(pendSheet.getMaxColumns(), 19 - pendSheet.getMaxColumns());
        if (!String(pendSheet.getRange(1, 19).getValue() || '').trim()) pendSheet.getRange(1, 19).setValue('סיים במכשיר');
        pendSheet.getRange(i + 1, 19).setValue(nowISO());
      }
      return jsonResponse({ status: 'ok' });
    }
  }
  return jsonResponse({ status: 'ok' });  // no matching row — harmless no-op
}

function handleDisqualify(p) {
  // Auth: two valid paths
  //   A) Examiner-initiated DQ — must include valid token AND own the session
  //   B) Self-DQ from examinee client (cheat detection) — pending row must exist with active status
  // Without one of these, reject. Prevents an attacker with just sessionCode+victim's idNumber
  // from disqualifying other examinees.
  var pendSheet = getSheet('ממתינים');
  var pendData = pendSheet.getDataRange().getValues();
  var name = '', phone = '', population = '', examineeLicense = '', examineeAudio = 'off';
  var pendRowIdx = -1, pendStatus = '';
  for (var j = pendData.length - 1; j >= 1; j--) {
    if (String(pendData[j][0]) === String(p.sessionCode) && normalizeId(pendData[j][1]) === normalizeId(p.idNumber)) {
      name = pendData[j][2] || '';
      phone = pendData[j][3] || '';
      population = pendData[j][7] || '';
      examineeLicense = pendData[j][8] || '';
      examineeAudio = pendData[j][9] || 'off';
      pendRowIdx = j;
      pendStatus = String(pendData[j][5] || '').trim();
      break;
    }
  }

  if (p.examinerId) {
    // Path A: examiner-initiated — require valid token + ownership
    if (!verifyToken(p.examinerId, p.token)) {
      return jsonResponse({ status: 'error', message: 'טוקן בוחן לא תקין', tokenExpired: true });
    }
    if (!verifyExaminerForSession(p.sessionCode, p.examinerId)) {
      return jsonResponse({ status: 'error', message: 'אין הרשאה — בוחן לא תואם לסשן' });
    }
  } else {
    // Path B: self-DQ — pending row must exist in active state AND the caller
    // must hold the examinee token issued at registration time. Legacy rows
    // (no stored token) are accepted as a transitional measure.
    if (pendRowIdx === -1) {
      return jsonResponse({ status: 'error', message: 'אין נבחן רשום בסשן זה' });
    }
    if (pendStatus !== 'in_exam' && pendStatus !== 'approved' && pendStatus !== 'disqualified') {
      return jsonResponse({ status: 'error', message: 'מצב לא תקף לפסילה: ' + pendStatus });
    }
    // Rate limit: max 10 self-DQ events per minute per (sessionCode, idNumber).
    // Anti-cheat can legitimately fire multiple beacons (retries, visibility +
    // blur racing); 10/min is well above any normal pattern.
    var dqRlErr = requireRateLimit('disqualify', String(p.sessionCode || '') + '_' + normalizeId(p.idNumber), 10, 60);
    if (dqRlErr) return dqRlErr;
    var tokenCheck = verifyExamineeToken(p.sessionCode, p.idNumber, p.examineeToken);
    if (!tokenCheck.valid) {
      return jsonResponse({ status: 'error', message: 'טוקן נבחן לא תקין לפסילה עצמית', examineeTokenError: tokenCheck.reason });
    }
  }

  // Update pending status to 'disqualified' (only if a row exists) AND increment
  // the DQ-event counter in column N so the examiner can see how many times this
  // examinee triggered an anti-cheat event — even if some were auto-reverted in
  // grace period via cancelDisqualify.
  if (pendRowIdx !== -1) {
    pendSheet.getRange(pendRowIdx + 1, 6).setValue('disqualified');
    var prevCount = 0;
    if (pendData[pendRowIdx].length > 13) prevCount = Number(pendData[pendRowIdx][13]) || 0;
    pendSheet.getRange(pendRowIdx + 1, 14).setValue(prevCount + 1);
    // Clear any OTHER active (in_exam/approved) rows for this examinee so a
    // duplicate row doesn't linger on the board beside the disqualified one.
    for (var dqd = 1; dqd < pendData.length; dqd++) {
      if (dqd === pendRowIdx) continue;
      if (String(pendData[dqd][0]) !== String(p.sessionCode) || normalizeId(pendData[dqd][1]) !== normalizeId(p.idNumber)) continue;
      var dqdStatus = String(pendData[dqd][5]).trim();
      if (dqdStatus === 'in_exam' || dqdStatus === 'approved') {
        pendSheet.getRange(dqd + 1, 6).setValue('cancelled');
      }
    }
  }

  // Idempotency: prevent duplicate פסול rows when examinee anti-cheat AND examiner
  // manual DQ fire on the same examinee close in time (different dqEventIds).
  // Rules:
  //   1. Same dqEventId on a פסול/בוטל row -> retry, skip silently.
  //   2. Recent (≤2 min) פסול row WITHOUT 'בוטל' status -> same logical DQ event from
  //      another path (e.g. examiner clicked after auto-DQ already fired) -> skip.
  //   3. Otherwise (latest is not פסול, or it's old/cancelled) -> create new row.
  var dqEventId = String(p.dqEventId || '');
  var sheet = getSheet('תוצאות');
  var data = sheet.getDataRange().getValues();
  var nowMs = Date.now();
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][13]) === String(p.sessionCode) && normalizeId(data[i][1]) === normalizeId(p.idNumber)) {
      var rowStatus = String(data[i][7]).trim();
      // Rule 1: same dqEventId (active or cancelled) — retry from sendDQToServer, skip
      if ((rowStatus === 'פסול' || rowStatus === 'בוטל') && dqEventId && String(data[i][24] || '') === dqEventId) {
        return jsonResponse({ status: 'ok' });
      }
      // Rule 2: latest is an active 'פסול' (not cancelled) within last 2 minutes
      // → treat as the same DQ episode even if dqEventId differs/missing.
      if (rowStatus === 'פסול') {
        var rowDateRaw = data[i][0];
        var rowDate = null;
        try {
          if (rowDateRaw instanceof Date) rowDate = rowDateRaw;
          else if (rowDateRaw) {
            // Sheet date column F may be "DD/MM/YYYY HH:mm" — parse manually
            var m = String(rowDateRaw).match(/(\d{1,2})\/(\d{1,2})\/(\d{4})\s+(\d{1,2}):(\d{2})/);
            if (m) rowDate = new Date(+m[3], (+m[2]) - 1, +m[1], +m[4], +m[5]);
          }
        } catch (e) { rowDate = null; }
        if (rowDate && (nowMs - rowDate.getTime()) < 120000) {
          // Within 2 minutes of an active פסול → duplicate from race between
          // examiner button and examinee anti-cheat. Skip.
          return jsonResponse({ status: 'ok', deduped: true });
        }
      }
      // Rule 3: not a duplicate — fall through to create new row
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
  // Only the examinee whose token matches the row may cancel their provisional DQ.
  var cdTokenErr = requireExamineeToken(p);
  if (cdTokenErr) return cdTokenErr;
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
  // Reset EVERY non-final row for this examinee (not just the latest) and accept
  // ALL stuck states — including 'disqualified'/'dq_confirmed'. Previously reset
  // refused those, so a soldier stuck on a pending DQ could not be cleared at all.
  // "אפס" should fully remove a stuck soldier from the board so they can re-register.
  var resetCount = 0;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) !== String(p.sessionCode) || normalizeId(data[i][1]) !== normalizeId(p.idNumber)) continue;
    var s = String(data[i][5]).trim();
    if (s === 'waiting' || s === 'approved' || s === 'in_exam' || s === 'disqualified' || s === 'dq_confirmed') {
      sheet.getRange(i + 1, 6).setValue('cancelled');
      resetCount++;
    }
  }
  if (resetCount === 0) {
    return jsonResponse({ status: 'error', message: 'לא נמצא נבחן פעיל לאיפוס' });
  }
  SpreadsheetApp.flush();
  return jsonResponse({ status: 'ok', resetCount: resetCount });
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
  // Close EVERY in_exam/approved row for this examinee (not just the latest, and
  // 'approved' too — markExamStarted can fail on iOS, leaving a stuck 'approved'
  // even after the soldier finished). One "סיים ידנית" must clear them all.
  for (var j = pendData.length - 1; j >= 1; j--) {
    if (String(pendData[j][0]) !== String(p.sessionCode) || normalizeId(pendData[j][1]) !== normalizeId(p.idNumber)) continue;
    var fcStatus = String(pendData[j][5]).trim();
    if (fcStatus !== 'in_exam' && fcStatus !== 'approved') continue;
    if (!found) { // capture details from the latest matching row
      name = pendData[j][2] || '';
      phone = pendData[j][3] || '';
      language = pendData[j][6] || 'he';
      population = pendData[j][7] || '';
      examineeLicense = pendData[j][8] || '';
      examineeAudio = pendData[j][9] || 'off';
    }
    pendSheet.getRange(j + 1, 6).setValue('completed');
    found = true;
  }
  if (!found) {
    return jsonResponse({ status: 'error', message: 'לא נמצא נבחן עם סטטוס in_exam/approved' });
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

  // Find the latest result row + the pending row for this examinee in one pass each.
  var sheet = getSheet('תוצאות');
  var data = sheet.getDataRange().getValues();
  var resultRowIdx = -1;
  var resultStatus = '';
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][13]) === String(p.sessionCode) && normalizeId(data[i][1]) === normalizeId(p.idNumber)) {
      resultRowIdx = i;
      resultStatus = String(data[i][7]).trim();
      break;
    }
  }

  var pendSheet = getSheet('ממתינים');
  var pendData = pendSheet.getDataRange().getValues();
  var pendRowIdx = -1;
  var pendStatusNow = '';
  for (var j = pendData.length - 1; j >= 1; j--) {
    if (String(pendData[j][0]) === String(p.sessionCode) && normalizeId(pendData[j][1]) === normalizeId(p.idNumber)) {
      pendRowIdx = j;
      pendStatusNow = String(pendData[j][5] || '').trim();
      break;
    }
  }

  // Case 1: latest result is פסול → normal overturn flow.
  // Pending revert covers BOTH 'disqualified' (auto-DQ, not yet confirmed) and
  // 'dq_confirmed' (examiner already clicked ✔ אשר). Without the dq_confirmed
  // branch, the examiner who pressed "אשר" by accident could overturn the
  // result row but the examinee stays locked out — they'd need a fresh
  // registration, which is what created duplicate rows at base 14.
  if (resultStatus === 'פסול') {
    sheet.getRange(resultRowIdx + 1, 8).setValue('בוטל');
    sheet.getRange(resultRowIdx + 1, 18).setValue(false);
    if (pendRowIdx !== -1 && (pendStatusNow === 'disqualified' || pendStatusNow === 'dq_confirmed')) {
      pendSheet.getRange(pendRowIdx + 1, 6).setValue('in_exam');
    }
    SpreadsheetApp.flush();
    return jsonResponse({ status: 'ok' });
  }

  // Case 2: stuck pending in 'disqualified' but latest result is already
  // a final outcome (עבר/נכשל/בוטל). Happens when DQ fired transiently
  // during a deploy window — examinee continued and finished the exam, but
  // the pending row stayed stuck. Just clean up the pending row.
  if (pendRowIdx !== -1 && pendStatusNow === 'disqualified' &&
      (resultStatus === 'עבר' || resultStatus === 'נכשל' || resultStatus === 'בוטל')) {
    pendSheet.getRange(pendRowIdx + 1, 6).setValue('completed');
    SpreadsheetApp.flush();
    return jsonResponse({ status: 'ok', resolved: 'stale_dq_cleared' });
  }

  // Case 3: pending is disqualified but no result row yet → revert so the
  // examinee can resume the exam (in_exam state, just like case 1).
  if (pendRowIdx !== -1 && pendStatusNow === 'disqualified' && resultRowIdx === -1) {
    pendSheet.getRange(pendRowIdx + 1, 6).setValue('in_exam');
    SpreadsheetApp.flush();
    return jsonResponse({ status: 'ok', resolved: 'no_result_reverted' });
  }

  // Fall-through: nothing to do
  return jsonResponse({ status: 'error', message: 'תוצאה לא נמצאה' });
}

function handleConfirmDQ(p) {
  // Unconditional ownership check. confirmDQ is now in examinerActions → requireToken
  // already forced a valid examinerId. The old `if (p.examinerId && ...)` form could
  // be bypassed by simply OMITTING examinerId, letting anyone who knows session+id
  // finalize a victim's provisional DQ (robbing their grace-period recovery).
  if (!verifyExaminerForSession(p.sessionCode, p.examinerId)) {
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

// Commander-only result correction. Allows changing score and pass/fail/DQ
// Manual result entry for transition period — examiner enters a paper-based
// exam outcome that bypassed the digital system. Appends a row to תוצאות with
// the same shape submitResult uses; marks column W as 'ידני' so reports can
// distinguish it from system-scored results. Requires examiner-token +
// session ownership (same auth as overturnDQ/correctToPass).
//
// Required: sessionCode, idNumber, examinerId, token, fullName, score, total.
// Optional: phone, license, population, audioMode, time.
function handleSubmitManualResult(p) {
  if (!verifyToken(p.examinerId, p.token)) {
    return jsonResponse({ status: 'error', message: 'טוקן בוחן לא תקין', tokenExpired: true });
  }
  if (!verifyExaminerForSession(p.sessionCode, p.examinerId)) {
    return jsonResponse({ status: 'error', message: 'אין הרשאה — בוחן לא תואם לסשן' });
  }
  // Required field validation. ID/name keep manual entries debuggable; score
  // pair lets the spreadsheet compute the same "כך/סה״כ" string the digital
  // flow writes, so existing reports parse it without special-casing.
  var fullName = String(p.fullName || '').trim();
  var idNumber = String(p.idNumber || '').trim();
  var scoreNum = parseInt(p.score, 10);
  var totalNum = parseInt(p.total, 10) || 30;
  if (!fullName) return jsonResponse({ status: 'error', message: 'חובה למלא שם מלא' });
  if (!idNumber) return jsonResponse({ status: 'error', message: 'חובה למלא ת.ז.' });
  if (isNaN(scoreNum) || scoreNum < 0 || scoreNum > totalNum) {
    return jsonResponse({ status: 'error', message: 'ציון לא תקין (חייב להיות בין 0 ל-' + totalNum + ')' });
  }

  // Pull session context so manual rows match the rest of the session's rows
  // (same site/classroom/language) without the examiner re-typing them.
  var site = '', classroom = '', sessLicense = '', sessLanguage = 'he', examinerName = '';
  var sesSheet = getSheet('סשנים');
  var sesData = sesSheet.getDataRange().getValues();
  for (var s = 1; s < sesData.length; s++) {
    if (String(sesData[s][0]).trim() === String(p.sessionCode).trim()) {
      examinerName = sesData[s][2] || '';
      site = sesData[s][3] || '';
      classroom = sesData[s][4] || '';
      sessLicense = sesData[s][5] || '';
      sessLanguage = sesData[s][6] || 'he';
      break;
    }
  }
  var license = String(p.license || sessLicense || 'B');
  var language = String(p.language || sessLanguage || 'he');

  // Percent + pass/fail mirror submitResult's behavior: 86% threshold (26/30).
  var percent = Math.round((scoreNum / totalNum) * 100);
  var passThreshold = Math.ceil(totalNum * 0.86);
  var passText = scoreNum >= passThreshold ? 'עבר' : 'נכשל';

  // WhatsApp link is convenient even for manual rows — examiner often wants to
  // send the same confirmation message they'd send for a digital exam.
  var waLink = '';
  if (p.phone) {
    var phoneFmt = formatPhoneForWA(p.phone);
    var waMsg = '*🚗 אישור תוצאת מבחן תאוריה חיצוני*\n\n' +
      'שם: ' + fullName + '\n' +
      'ת.ז.: ' + idNumber + '\n' +
      'דרגה: ' + license + '\n' +
      (p.population ? 'אוכלוסיה: ' + p.population + '\n' : '') +
      'תאריך: ' + todayStr() + '\n' +
      'ציון: ' + scoreNum + '/' + totalNum + ' (' + percent + '%)\n' +
      'תוצאה: *' + passText + '*\n';
    if (phoneFmt) waLink = 'https://wa.me/' + phoneFmt + '?text=' + encodeURIComponent(waMsg);
  }

  var attemptNum = countAttempts(idNumber, license) + 1;
  var sheet = getSheet('תוצאות');
  // Idempotency: a lost-response retry (request landed, reply dropped, examiner
  // re-saves) must not create a second identical manual row. Skip if a non-בוטל
  // row already exists for this session+id+license+score.
  var manExisting = sheet.getDataRange().getValues();
  for (var mx = manExisting.length - 1; mx >= 1; mx--) {
    if (String(manExisting[mx][13]) === String(p.sessionCode) &&
        normalizeId(manExisting[mx][1]) === normalizeId(idNumber) &&
        String(manExisting[mx][4]) === String(license) &&
        String(manExisting[mx][5]) === (scoreNum + '/' + totalNum) &&
        String(manExisting[mx][7] || '').trim() !== 'בוטל') {
      return jsonResponse({ status: 'ok', duplicate: true, waLink: manExisting[mx][18] || '' });
    }
  }
  sheet.appendRow([
    todayStr(),
    idNumber,
    fullName,
    p.phone || '',
    license,
    scoreNum + '/' + totalNum,
    percent + '%',
    passText,
    p.time || '',
    examinerName,
    site,
    classroom,
    language,
    String(p.sessionCode),
    attemptNum,
    '',                                 // P (15) wrongDetails — N/A for manual
    false,                              // Q (16) corrected
    false,                              // R (17) disqualified
    waLink,
    p.population || '',
    false,                              // U (20) suspicious
    p.audioMode || 'off',
    'ידני',                             // W (22) verified flag — marks paper-based entry
    '',                                 // X (23) suspicious text
    '',                                 // Y (24) dqEventId
    '',                                 // Z (25) תוקן ע"י
    '',                                 // AA (26) סיבת תיקון
    '',                                 // AB (27) תאריך תיקון
    ''                                  // AC (28) מסלול שפות
  ]);
  SpreadsheetApp.flush();
  return jsonResponse({ status: 'ok', waLink: waLink, attempt: attemptNum });
}

// status on any result row, with a mandatory reason recorded for audit.
// Caller must have a valid examiner token AND role 'מפקד' in the בוחנים sheet.
// Required params: sessionCode, idNumber, newScore (e.g. "28"), newTotal (e.g. "30"),
//                  newStatus ('עבר' | 'נכשל' | 'פסול'), reason (non-empty).
// Examiner-level correction of an examinee's site + population on their result
// row (the examinee picked the wrong site/population at registration). Reached
// via doPost (apiPost auto-attaches examinerId+token), which does NOT run the
// examinerActions allowlist — so the token is verified HERE, then session
// ownership. Updates תוצאות col 11 (אתר, idx 10) / col 20 (אוכלוסיה, idx 19).
function handleCorrectExamineeMeta(p) {
  if (!verifyToken(p.examinerId, p.token)) {
    return jsonResponse({ status: 'error', message: 'טוקן בוחן לא תקין', tokenExpired: true });
  }
  if (!verifyExaminerForSession(p.sessionCode, p.examinerId)) {
    return jsonResponse({ status: 'error', message: 'אין הרשאה — בוחן לא תואם לסשן' });
  }
  var newSite = (typeof p.site !== 'undefined' && p.site !== null) ? String(p.site).trim() : '';
  var newPop = (typeof p.population !== 'undefined' && p.population !== null) ? String(p.population).trim() : '';
  var newPhone = (typeof p.phone !== 'undefined' && p.phone !== null) ? String(p.phone).trim() : null;  // null = "not sent" → don't touch
  var newId = (typeof p.newIdNumber !== 'undefined' && p.newIdNumber !== null) ? String(p.newIdNumber).trim() : '';
  // Only apply an id change when it's a valid digit string AND actually different.
  var applyId = (newId && /^\d{5,10}$/.test(newId) && normalizeId(newId) !== normalizeId(p.idNumber));
  if (!newSite && !newPop && newPhone === null && !applyId) {
    return jsonResponse({ status: 'error', message: 'לא הוזנו שדות לעדכון' });
  }
  var sheet = getSheet('תוצאות');
  var rows = sheet.getDataRange().getValues();
  for (var i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][13]) === String(p.sessionCode) && normalizeId(rows[i][1]) === normalizeId(p.idNumber)) {
      var rowIdx = i + 1;
      if (applyId) {
        var idCell = sheet.getRange(rowIdx, 2);   // B (idx 1) = ת.ז.
        idCell.setNumberFormat('@');              // store as text — preserve leading zeros / avoid number formatting
        idCell.setValue(newId);
      }
      if (newPhone !== null) {
        var phoneCell = sheet.getRange(rowIdx, 4); // D (idx 3) = טלפון
        phoneCell.setNumberFormat('@');
        phoneCell.setValue(newPhone);
      }
      if (newSite) sheet.getRange(rowIdx, 11).setValue(newSite);   // K (idx 10) = אתר
      if (newPop) sheet.getRange(rowIdx, 20).setValue(newPop);     // T (idx 19) = אוכלוסיה
      SpreadsheetApp.flush();
      return jsonResponse({ status: 'ok' });
    }
  }
  return jsonResponse({ status: 'error', message: 'תוצאה לא נמצאה' });
}

function handleCommanderCorrectResult(data) {
  // Token + role check (token already verified by examinerActions allowlist,
  // but we re-check role here since the role doesn't appear in that allowlist).
  if (!verifyToken(data.examinerId, data.token)) {
    return jsonResponse({ status: 'error', message: 'טוקן בוחן לא תקין', tokenExpired: true });
  }
  var role = getExaminerRole(data.examinerId);
  if (role !== 'מפקד') {
    return jsonResponse({ status: 'error', message: 'פעולה זו זמינה רק למפקדים' });
  }

  // Validate inputs
  var reason = String(data.reason || '').trim();
  if (!reason) {
    return jsonResponse({ status: 'error', message: 'יש להזין סיבת תיקון' });
  }
  var newScore = parseInt(data.newScore, 10);
  var newTotal = parseInt(data.newTotal, 10);
  if (isNaN(newScore) || isNaN(newTotal) || newTotal <= 0 || newScore < 0 || newScore > newTotal) {
    return jsonResponse({ status: 'error', message: 'ציון חדש לא תקין' });
  }
  var newStatus = String(data.newStatus || '').trim();
  if (newStatus !== 'עבר' && newStatus !== 'נכשל' && newStatus !== 'פסול') {
    return jsonResponse({ status: 'error', message: 'סטטוס חדש לא תקין' });
  }

  var sheet = getSheet('תוצאות');
  var rows = sheet.getDataRange().getValues();
  for (var i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][13]) === String(data.sessionCode) && normalizeId(rows[i][1]) === normalizeId(data.idNumber)) {
      var rowIdx = i + 1;
      var pct = Math.round((newScore / newTotal) * 100);
      // Apply updates
      sheet.getRange(rowIdx, 6).setValue(newScore + '/' + newTotal);  // F: ציון
      sheet.getRange(rowIdx, 7).setValue(pct + '%');                   // G: אחוז
      sheet.getRange(rowIdx, 8).setValue(newStatus);                   // H: עבר/נכשל
      sheet.getRange(rowIdx, 18).setValue(newStatus === 'פסול');       // R: פסול?
      sheet.getRange(rowIdx, 21).setValue(true);                       // U: תוקן?
      // Audit trail (columns Z=26, AA=27, AB=28)
      // Look up commander's display name from בוחנים sheet
      var commanderName = '';
      try {
        var examSheet = getSheet('בוחנים');
        var examData = examSheet.getDataRange().getValues();
        for (var x = 1; x < examData.length; x++) {
          if (normalizeId(examData[x][1]) === normalizeId(data.examinerId)) {
            commanderName = String(examData[x][0] || '');
            break;
          }
        }
      } catch(e) {}
      sheet.getRange(rowIdx, 26).setValue(commanderName + ' (' + normalizeId(data.examinerId) + ')');
      sheet.getRange(rowIdx, 27).setValue(reason);
      sheet.getRange(rowIdx, 28).setValue(todayStr());
      SpreadsheetApp.flush();
      return jsonResponse({ status: 'ok' });
    }
  }
  return jsonResponse({ status: 'error', message: 'תוצאה לא נמצאה' });
}

function handleMarkSent(p) {
  // Ownership check — consistent with the other examiner mutations; prevents an
  // authenticated examiner from flipping the "נשלח?" flag on another session's rows.
  if (!verifyExaminerForSession(p.sessionCode, p.examinerId)) {
    return jsonResponse({ status: 'error', message: 'אין הרשאה — בוחן לא תואם לסשן' });
  }
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
  // Reject the call if the examinee token doesn't match the registered row
  // (legacy rows without a stored token still pass).
  var reqTokenErr = requireExamineeToken(data);
  if (reqTokenErr) return reqTokenErr;
  // Server-side score verification setup. The client tells us which questions
  // came up and how each was shuffled — but NOT which answer is correct. The
  // server looks up the canonical correct index in ANSWER_KEY_BY_LANG and
  // computes the shuffled-correct index itself, so a tampered client cannot
  // claim "answer 0 is always correct" and pass without taking the exam.
  //
  // Expected payload:
  //   { sessionCode, idNumber, language?, questions: [{qIdx, qId, shuffleOrder}] }
  // Where shuffleOrder is an array like [2,0,1,3] meaning:
  //   "displayed answer A = original answer 2, B = original 0, C = original 1, D = original 3".
  //
  // Backward-compat: old clients send {qIdx, correctShuffledIdx} (legacy, trusted).
  // If we detect the legacy shape, we accept it but mark the registration
  // as unverified so submitResult flags the result row accordingly.
  if (!data.sessionCode || !data.idNumber || !data.questions) {
    return jsonResponse({ status: 'error', message: 'חסרים נתונים לרישום מבחן' });
  }

  // If server-side question delivery (getExamQuestions) was used, an entry
  // for `issued_qs_<session>_<id>` will exist in cache. Verify the IDs the
  // client is registering all came from that set — otherwise reject.
  // (No cache entry means legacy flow where the client picked questions
  // locally from questions.js; that path stays open for now.)
  try {
    var issuedKey = 'issued_qs_' + String(data.sessionCode) + '_' + normalizeId(data.idNumber);
    var issuedJson = CacheService.getScriptCache().get(issuedKey);
    if (issuedJson) {
      var issuedSet = {};
      var issuedArr = JSON.parse(issuedJson) || [];
      for (var iz = 0; iz < issuedArr.length; iz++) issuedSet[String(issuedArr[iz])] = true;
      for (var iq = 0; iq < data.questions.length; iq++) {
        var qq = data.questions[iq];
        if (!qq || !qq.qId) continue;
        if (!issuedSet[String(qq.qId)]) {
          return jsonResponse({
            status: 'error',
            message: 'Question ID not in issued set — rejecting registration',
            unexpectedId: qq.qId
          });
        }
      }
    }
  } catch (e) { /* cache failure — fall through to existing flow */ }

  // Verify examinee is in_exam / approved status
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

  // Build the canonical question map. Each entry stores {qIdx, correctShuffledIdx}
  // — same shape submitResult already consumes — but the correctShuffledIdx is
  // computed server-side whenever possible.
  var lang = String(data.language || 'he').toLowerCase();
  var canonicalMap = [];
  var unverifiedCount = 0;
  var hasAnswerKey = (typeof ANSWER_KEY_BY_LANG !== 'undefined') && (typeof lookupCorrectIndex === 'function');

  for (var qi = 0; qi < data.questions.length; qi++) {
    var q = data.questions[qi];
    if (!q) { canonicalMap.push(null); unverifiedCount++; continue; }

    // Modern shape: client sent qId + shuffleOrder → server computes
    if (hasAnswerKey && q.qId && Array.isArray(q.shuffleOrder)) {
      var origCorrect = lookupCorrectIndex(Number(q.qId), lang);
      if (origCorrect === null || origCorrect === undefined) {
        // Question id missing from answer key → fall back to client's claim if present
        canonicalMap.push({ qIdx: q.qIdx, qId: Number(q.qId), shuffleOrder: q.shuffleOrder, correctShuffledIdx: Number(q.correctShuffledIdx || 0) });
        unverifiedCount++;
        continue;
      }
      var idxInShuffle = q.shuffleOrder.indexOf(Number(origCorrect));
      if (idxInShuffle < 0) {
        // shuffleOrder doesn't contain the correct original index → malformed
        canonicalMap.push({ qIdx: q.qIdx, qId: Number(q.qId), shuffleOrder: q.shuffleOrder, correctShuffledIdx: Number(q.correctShuffledIdx || 0) });
        unverifiedCount++;
        continue;
      }
      // Store qId + shuffleOrder so submitResult can build wrongDetails server-side
      // (needed because we no longer send `ci` to examinees — see handleGetExamQuestions).
      canonicalMap.push({ qIdx: q.qIdx, qId: Number(q.qId), shuffleOrder: q.shuffleOrder, correctShuffledIdx: idxInShuffle });
      continue;
    }

    // Legacy / fallback: client sent correctShuffledIdx directly → trust but flag
    canonicalMap.push({ qIdx: q.qIdx, qId: q.qId ? Number(q.qId) : null, shuffleOrder: Array.isArray(q.shuffleOrder) ? q.shuffleOrder : null, correctShuffledIdx: Number(q.correctShuffledIdx || 0) });
    unverifiedCount++;
  }

  // Guard: if the server answer key is entirely unavailable (answer_key.gs not
  // deployed, or every selected qId missing from it), do NOT register a map full
  // of unverifiable garbage — that is what produced silent 0/30 fails (פינטו/דיין
  // 03/06). Return an error so the confirmed-register client blocks the exam and
  // retries, surfacing the problem instead of mis-scoring a real examinee.
  if (!hasAnswerKey || (data.questions.length > 0 && unverifiedCount >= data.questions.length)) {
    return jsonResponse({ status: 'error', message: 'מפתח התשובות אינו זמין בשרת — פנה למנהל המערכת', keyUnavailable: true });
  }

  // Store in מבחנים sheet (create if needed). Add a fifth column for unverified
  // count so submitResult can flag results scored from unverified data.
  var examSheet;
  try { examSheet = getSheet('מבחנים'); } catch(e) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    examSheet = ss.insertSheet('מבחנים');
    examSheet.appendRow(['קוד סשן', 'ת.ז.', 'שאלות JSON', 'זמן רישום', 'שפה', 'שגויות לא מאומתות']);
  }
  examSheet.appendRow([
    String(data.sessionCode),
    normalizeId(data.idNumber),
    JSON.stringify(canonicalMap),
    nowISO(),
    lang,
    unverifiedCount
  ]);

  // Layer-1 consolidation: also mark the examinee in_exam here — the same write
  // markExamStarted did — so the start no longer needs a separate markExamStarted
  // round-trip. Only flips 'approved' → 'in_exam' (same guard). Because this is
  // the CONFIRMED/blocking call, it also strengthens the iPhone "stuck in
  // ממתינים" fix. `marked` is returned so the client knows whether to keep the
  // visibilitychange fallback armed.
  var marked = false;
  try {
    var penSheet = getSheet('ממתינים');
    var penData = penSheet.getDataRange().getValues();
    for (var mi = penData.length - 1; mi >= 1; mi--) {
      if (String(penData[mi][0]) !== String(data.sessionCode) || normalizeId(penData[mi][1]) !== normalizeId(data.idNumber)) continue;
      var mst = String(penData[mi][5]).trim();
      if (mst === 'approved') { penSheet.getRange(mi + 1, 6).setValue('in_exam'); penSheet.getRange(mi + 1, 12).setValue(nowISO()); marked = true; break; }
      if (mst === 'in_exam') { marked = true; break; }
      // other status (cancelled/completed): keep scanning for an approved/in_exam row
    }
  } catch (msErr) { /* non-fatal — client fallback + dashboard cleanup cover it */ }

  return jsonResponse({ status: 'ok', verified: unverifiedCount === 0, unverifiedCount: unverifiedCount, examStarted: marked });
}

function handleSubmitResult(data) {
  // Rate limit: max 5 submissions per minute per (sessionCode, idNumber).
  // One legitimate submission + retries on flaky network; floods are blocked.
  var srRlErr = requireRateLimit('submitResult', String(data.sessionCode || '') + '_' + normalizeId(data.idNumber), 5, 60);
  if (srRlErr) return srRlErr;
  // Require the examinee token before accepting any result. Legacy rows
  // (no stored token) pass through requireExamineeToken with legacy=true.
  var srTokenErr = requireExamineeToken(data);
  if (srTokenErr) return srTokenErr;
  var sheet = getSheet('תוצאות');

  // Verify examinee is approved (in_exam status) before accepting results
  if (data.sessionCode && data.idNumber) {
    var pendSheet = getSheet('ממתינים');
    var pendData = pendSheet.getDataRange().getValues();
    var isApproved = false;
    for (var pi = pendData.length - 1; pi >= 1; pi--) {
      if (String(pendData[pi][0]) === String(data.sessionCode) && normalizeId(pendData[pi][1]) === normalizeId(data.idNumber)) {
        var pStatus = String(pendData[pi][5]).trim();
        // 'cancelled' accepted too: if an examiner reset an examinee who was
        // actually still mid-exam, a genuine finished submit must be RECORDED,
        // not rejected and lost. The fabricated-fail supersede + dup-check below
        // prevent a double-row; close-fails for 'cancelled' are suppressed.
        if (pStatus === 'in_exam' || pStatus === 'approved' || pStatus === 'completed' || pStatus === 'cancelled') {
          isApproved = true;
        }
        break;
      }
    }
    if (!isApproved) {
      return jsonResponse({ status: 'error', message: 'נבחן לא מאושר — לא ניתן לשלוח תוצאות' });
    }
  }

  // SECURITY (anti score-forge): if a registered exam (מבחנים row) exists for this
  // session+id, the answers array is MANDATORY so the server re-scores from the
  // answer key. Without this, an examinee could POST a forged score with NO answers
  // and skip BOTH the re-score and the unverified-guard below (both answers-gated).
  var hasRegisteredExam = false;
  try {
    var regChk = getSheet('מבחנים').getDataRange().getValues();
    for (var rc = regChk.length - 1; rc >= 1; rc--) {
      if (String(regChk[rc][0]) === String(data.sessionCode) && normalizeId(regChk[rc][1]) === normalizeId(data.idNumber)) { hasRegisteredExam = true; break; }
    }
  } catch (regChkErr) {}
  if (hasRegisteredExam && (!data.answers || !Array.isArray(data.answers) || data.answers.length === 0)) {
    return jsonResponse({ status: 'error', message: 'הגשה לא תקינה — חסרות תשובות למבחן רשום' });
  }

  // Server-side score verification: if answers array is present, recalculate
  // score using the question map registered at exam start. The map's
  // correctShuffledIdx values are server-computed (from ANSWER_KEY_BY_LANG)
  // when possible — only fall back to client-claimed values for questions
  // missing from the answer key, in which case we mark the result unverified.
  if (data.answers && Array.isArray(data.answers)) {
    try {
      var examSheet = getSheet('מבחנים');
      var examData = examSheet.getDataRange().getValues();
      var questionMap = null;
      var unverifiedCount = 0;
      var registeredLang = '';
      // Find the latest registered exam for this session+ID
      for (var ei = examData.length - 1; ei >= 1; ei--) {
        if (String(examData[ei][0]) === String(data.sessionCode) && normalizeId(examData[ei][1]) === normalizeId(data.idNumber)) {
          questionMap = JSON.parse(examData[ei][2]);
          // Column F (index 5) = unverified-count (added when registerExamQuestions stored this row).
          // Older rows may not have this column → treat as fully unverified to be safe.
          unverifiedCount = (examData[ei].length > 5) ? Number(examData[ei][5] || 0) : questionMap.length;
          // Column E (index 4) = language (added in registerExamQuestions)
          registeredLang = (examData[ei].length > 4) ? String(examData[ei][4] || '') : '';
          break;
        }
      }
      if (questionMap) {
        // Shuffle indexes refer to original answer positions in whatever
        // language the questions were registered in. Translators reorder
        // answers, so the original "correct index" can differ between
        // languages (e.g. he Q128 → idx 1, ar Q128 → idx 2). When the
        // examinee switched language mid-exam, score each answer against
        // the correct index for THE LANGUAGE THEY SAW IT IN, not the one
        // captured at registration.
        function correctIdxForLang(mapEntry, lang) {
          if (!mapEntry || !lang) return null;
          if (typeof lookupCorrectIndex !== 'function') return null;
          if (!mapEntry.qId || !Array.isArray(mapEntry.shuffleOrder)) return null;
          var orig = lookupCorrectIndex(Number(mapEntry.qId), String(lang).toLowerCase());
          if (orig === null || orig === undefined) return null;
          var pos = mapEntry.shuffleOrder.indexOf(Number(orig));
          return pos >= 0 ? pos : null;
        }
        function effectiveCorrectIdx(mapEntry, langAtAnswer) {
          var lang = langAtAnswer ? String(langAtAnswer).toLowerCase() : '';
          if (lang && lang !== registeredLang) {
            var alt = correctIdxForLang(mapEntry, lang);
            if (alt !== null) return alt;
          }
          return Number(mapEntry.correctShuffledIdx);
        }
        var correctCount = 0;
        var totalQ = questionMap.length;
        for (var ai = 0; ai < data.answers.length && ai < totalQ; ai++) {
          if (data.answers[ai] !== null && data.answers[ai] !== undefined && questionMap[ai]) {
            var selected = Number(data.answers[ai].selected);
            var correctIdx = effectiveCorrectIdx(questionMap[ai], data.answers[ai].langAtAnswer);
            if (selected === correctIdx) correctCount++;
          }
        }
        var pct = Math.round((correctCount / totalQ) * 100);
        var passThreshold = Math.ceil(totalQ * 0.86); // ~26/30
        data.score = correctCount;
        data.total = totalQ;
        data.percent = pct;
        data.passed = correctCount >= passThreshold;
        // verified=true ONLY when every question in the map was scored against
        // a server-trusted answer key. Any fallback entry → unverified.
        data.verified = (unverifiedCount === 0);

        // ===== Server-side wrong-answers reconstruction =====
        // Each wrong answer is rendered in the language the examinee was viewing
        // WHEN they answered that specific question (data.answers[i].langAtAnswer).
        // Without this, an examinee who switched mid-exam sees mixed-language
        // feedback that doesn't match what they actually saw.
        try {
          // Lazy per-language cache: questions DB + byId map per language code.
          // Avoids loading every language up-front when most exams use one.
          var langDbCache = {};
          function getLangDb(lang) {
            var safeLang = String(lang || 'he').toLowerCase();
            if (langDbCache[safeLang]) return langDbCache[safeLang];
            try {
              var qs = loadQuestionsForLanguageServer(safeLang);
              if (!qs || !qs.length) return null;
              var idx = {};
              for (var q = 0; q < qs.length; q++) {
                if (qs[q] && qs[q].id !== undefined) idx[String(qs[q].id)] = qs[q];
              }
              langDbCache[safeLang] = { byId: idx, labels: (safeLang === 'he') ? ['א','ב','ג','ד','ה','ו'] : ['A','B','C','D','E','F'] };
              return langDbCache[safeLang];
            } catch (loadErr) {
              langDbCache[safeLang] = null;
              return null;
            }
          }
          var defaultLang = String(data.language || registeredLang || 'he').toLowerCase();
          // Pre-warm the default so we have a fallback for missing per-question langs.
          var defaultDb = getLangDb(defaultLang);

          var serverWrong = [];
          for (var wi = 0; wi < data.answers.length && wi < questionMap.length; wi++) {
            var mapEntry = questionMap[wi];
            var ans2 = data.answers[wi];
            if (!mapEntry) continue;
            var selected2 = ans2 ? Number(ans2.selected) : -1;
            // Use the per-answer language's correctIdx — same logic as the
            // scoring loop above, so wrong-answer reconstruction matches the
            // pass/fail tally instead of contradicting it after a mid-exam
            // language switch.
            var correctIdx2 = effectiveCorrectIdx(mapEntry, ans2 && ans2.langAtAnswer);
            if (selected2 === correctIdx2) continue; // got it right
            // Pick the language the examinee was viewing when they answered this Q.
            // Falls back to the exam's primary language when missing (old clients).
            var perAnsLang = (ans2 && ans2.langAtAnswer) ? String(ans2.langAtAnswer).toLowerCase() : defaultLang;
            var db = getLangDb(perAnsLang) || defaultDb;
            if (!db || !db.byId) continue;
            var qInfo = (mapEntry.qId !== undefined && mapEntry.qId !== null) ? db.byId[String(mapEntry.qId)] : null;
            if (!qInfo || !Array.isArray(mapEntry.shuffleOrder) || !Array.isArray(qInfo.answers)) continue;
            var shuffled = mapEntry.shuffleOrder.map(function(origIdx) { return qInfo.answers[origIdx]; });
            var yourLabel = '', yourText = '';
            if (selected2 === -1 || selected2 < 0 || selected2 >= shuffled.length) {
              yourText = (perAnsLang === 'he') ? 'לא נענתה' : 'Not answered';
            } else {
              yourLabel = db.labels[selected2] || '';
              yourText = shuffled[selected2] || '';
            }
            var correctLabel = (correctIdx2 >= 0 && correctIdx2 < shuffled.length) ? (db.labels[correctIdx2] || '') : '';
            var correctText = (correctIdx2 >= 0 && correctIdx2 < shuffled.length) ? (shuffled[correctIdx2] || '') : '';
            // Classify the raw question category to the bucket name used by
            // EXAM_STRUCTURE (בטיחות / הכרת הרכב / חוק / תמרורים / ספציפי).
            // Without classification the certificate shows all-100%.
            var classifiedCat = (typeof classifyCategoryServer === 'function')
              ? classifyCategoryServer(qInfo.category)
              : '';
            serverWrong.push({
              question: qInfo.text || ('שאלה ' + (wi + 1)),
              yourAnswer: yourLabel ? (yourLabel + ' - ' + yourText) : yourText,
              correctAnswer: correctLabel ? (correctLabel + ' - ' + correctText) : correctText,
              category: classifiedCat || qInfo.category || ''
            });
          }
          // Always replace client-provided wrongAnswers — server is authoritative.
          if (defaultDb) data.wrongAnswers = serverWrong;
        } catch (rwe) {
          // Reconstruction failed (Drive load, etc.) — keep whatever client sent
          // rather than wiping it. Log for diagnosis.
          try { Logger.log('wrong-answer rebuild failed: ' + (rwe && rwe.message)); } catch(_) {}
        }
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

  // Guard: answers were present but the server never re-scored (no מבחנים row /
  // questionMap missing → data.verified left undefined above). Do NOT silently
  // trust the client's score — it is computed from a deliberately-stripped `ci`
  // and can be garbage (the historical false-0/30). Flag the row unverified so
  // the examiner reviews it instead of recording a bogus pass/fail.
  if (data.answers && Array.isArray(data.answers) && typeof data.verified === 'undefined') {
    data.verified = false;
    data.scoreUnverified = true;
  }

  // ===== Supersede any SYSTEM-FABRICATED fail for this session+id =====
  // A real finished submit must WIN over a system-written fail — the browser-
  // close beacon (handleSubmitFailOnClose) or the dashboard timeout/disconnect
  // row — created while the examinee was offline/backgrounded. Those are 'נכשל'
  // rows whose note carries a machine marker. Match on session+id ONLY (NOT
  // language/license): the fabricated row is stamped with the REGISTRATION
  // language, but the real submit may carry a DIFFERENT final language after a
  // mid-exam switch (Russian/Arabic/Amharic examinees), so a language-scoped
  // match would miss it and the dup-check below would swallow the real result →
  // a false 0/30 "vanished" exam, especially on iOS. Mark them בוטל (audit kept).
  // Mirrors the פסול-supersede pass below; genuine real נכשל rows lack the marker.
  var fabRows = sheet.getDataRange().getValues();
  var fabSuperseded = false;
  for (var fb = 1; fb < fabRows.length; fb++) {
    if (String(fabRows[fb][13]) !== String(data.sessionCode)) continue;
    if (normalizeId(fabRows[fb][1]) !== normalizeId(data.idNumber)) continue;
    if (String(fabRows[fb][7]).trim() !== 'נכשל') continue;
    var fbNote = String(fabRows[fb][15] || '');
    // markers: close-beacon ('סגירת דפדפן'), dashboard timeout ('טיימאאוט'), and
    // examiner manual-disconnect ('סיום ידני ... ניתוק/תקלה'). All three mean "did
    // not finish properly" — a real finished submit must override them.
    if (fbNote.indexOf('סגירת דפדפן') === -1 && fbNote.indexOf('טיימאאוט') === -1 && fbNote.indexOf('סיום ידני') === -1) continue;
    sheet.getRange(fb + 1, 8).setValue('בוטל');                                    // H = pass/fail
    sheet.getRange(fb + 1, 27).setValue('בוטל אוטומטית — הנבחן השלים והגיש מבחן');  // AA = reason
    sheet.getRange(fb + 1, 28).setValue(todayStr());                               // AB = correction date
    fabSuperseded = true;
  }
  if (fabSuperseded) SpreadsheetApp.flush(); // make the בוטל visible to the dup-check read below

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
        // Genuine prior real result for this exact exam — a true duplicate.
        // (Fabricated close/timeout fails were already superseded to בוטל above
        // and are skipped by the status filter, so they can't masquerade here.)
        markPendingCompleted(data.sessionCode, data.idNumber);
        return jsonResponse({ status: 'ok', waLink: existingData[d][18] || '', duplicate: true });
      }
    }
  }

  // Belt-and-suspenders: never let the literal "undefined" reach the certificate.
  // The client never receives `ci`, so its locally-built wrongAnswers carry
  // "undefined - undefined" as the correct answer; the server normally rebuilds
  // them, but if that failed (Drive/cache down) the client text is kept. Replace
  // any "undefined" with a neutral placeholder so feedback is never garbled.
  if (Array.isArray(data.wrongAnswers)) {
    for (var sw = 0; sw < data.wrongAnswers.length; sw++) {
      var swItem = data.wrongAnswers[sw];
      if (swItem && typeof swItem.correctAnswer === 'string' && swItem.correctAnswer.indexOf('undefined') !== -1) swItem.correctAnswer = '(לא זמין כעת)';
      if (swItem && typeof swItem.yourAnswer === 'string' && swItem.yourAnswer.indexOf('undefined') !== -1) swItem.yourAnswer = '(לא זמין)';
    }
  }

  var wrongDetails = '';
  var wrongForWA = '';
  if (data.wrongAnswers && data.wrongAnswers.length > 0) {
    for (var i = 0; i < data.wrongAnswers.length; i++) {
      var w = data.wrongAnswers[i];
      // Question ID prefix lets the commander dashboard aggregate by the exact
      // question (not just generic text "מה פירוש התמרור?" that collapses 50+
      // distinct sign questions into one row). Backward-compatible — readers
      // tolerate the line being missing for legacy rows.
      if (w.questionId) wrongDetails += 'מזהה שאלה: ' + w.questionId + '\n';
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

  // Surface the unverified-score guard (set above) loudly in the stored detail.
  if (data.scoreUnverified) {
    wrongDetails = '⚠️ ציון לא אומת בשרת (רישום מבחן חסר) — נדרש אימות ידני\n\n' + wrongDetails;
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

  // Format language history into a readable path. Single language = just the
  // code (e.g. "he"). Multiple = arrow-joined (e.g. "he → ru → he") so the
  // examiner can see at a glance that the examinee switched languages.
  var langPath = '';
  if (Array.isArray(data.languageHistory) && data.languageHistory.length > 0) {
    langPath = data.languageHistory.length === 1
      ? String(data.languageHistory[0])
      : data.languageHistory.join(' → ');
  } else {
    langPath = data.language || 'he';
  }

  // Supersede any prior פסול row for THIS session+id. Scenario: examinee was
  // auto-DQ'd, the overturn flow didn't finish (examiner clicked אשר, or hit
  // a stale "תוצאה לא נמצאה" path), then the examinee was re-allowed in and
  // finished the exam. Without this cleanup the sheet ends up with both a
  // פסול row AND a עבר/נכשל row — which is what happened at base 14 today.
  // We mark the old row as בוטל (audit trail preserved) and log the reason.
  var existingRows = sheet.getDataRange().getValues();
  for (var ex = existingRows.length - 1; ex >= 1; ex--) {
    if (String(existingRows[ex][13]) === String(data.sessionCode) &&
        normalizeId(existingRows[ex][1]) === normalizeId(data.idNumber) &&
        String(existingRows[ex][7]).trim() === 'פסול') {
      sheet.getRange(ex + 1, 8).setValue('בוטל');           // H = pass/fail
      sheet.getRange(ex + 1, 18).setValue(false);            // R = disqualified flag
      sheet.getRange(ex + 1, 27).setValue('בוטל אוטומטית — נבחן ניגש למבחן מחדש'); // AA = reason
      sheet.getRange(ex + 1, 28).setValue(todayStr());       // AB = correction date
    }
  }

  // Idempotency: skip if an identical result row already exists. A retry/resend
  // whose original response was lost (flaky network) would otherwise create a
  // duplicate. Matches session+id+license+score+time+result, so a re-take or a
  // post-overturn submit (different score/time/result) is still appended.
  for (var dc = existingRows.length - 1; dc >= 1; dc--) {
    if (String(existingRows[dc][13]) === String(data.sessionCode) &&
        normalizeId(existingRows[dc][1]) === normalizeId(data.idNumber) &&
        String(existingRows[dc][4]) === String(data.license) &&
        String(existingRows[dc][5]) === (data.score + '/' + data.total) &&
        String(existingRows[dc][7]).trim() === String(passText).trim() &&
        String(existingRows[dc][8]) === String(data.time)) {
      markPendingCompleted(data.sessionCode, data.idNumber);
      return jsonResponse({ status: 'ok', duplicate: true, waLink: waLink });
    }
  }

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
    data.suspicious ? 'חשוד' : '',
    '',                                 // Y (24) dqEventId — not a DQ row
    '',                                 // Z (25) תוקן ע"י — empty (no correction yet)
    '',                                 // AA (26) סיבת תיקון — empty
    '',                                 // AB (27) תאריך תיקון — empty
    langPath,                           // AC (28) מסלול שפות — full path he → ru → he
    String(data.device || '')           // AD (29) מכשיר — phone / tablet / desktop
  ]);

  // Update pending status to completed
  markPendingCompleted(data.sessionCode, data.idNumber);

  return jsonResponse({ status: 'ok', waLink: waLink });
}

// Helper: mark ALL active pending rows for this session+ID as completed.
// Closes EVERY in_exam/approved row (not just the latest) — a duplicate pending
// row otherwise leaves the soldier stuck on the board even though they finished
// and submitted (reported: "stuck in ממתינים/במבחן despite finishing").
function markPendingCompleted(sessionCode, idNumber) {
  var pendSheet = getSheet('ממתינים');
  var pendData = pendSheet.getDataRange().getValues();
  for (var j = pendData.length - 1; j >= 1; j--) {
    if (String(pendData[j][0]) === String(sessionCode) && normalizeId(pendData[j][1]) === normalizeId(idNumber) && (String(pendData[j][5]).trim() === 'in_exam' || String(pendData[j][5]).trim() === 'approved')) {
      pendSheet.getRange(j + 1, 6).setValue('completed');
    }
  }
}

function handleSubmitWrongAnswers(p) {
  var swaTokenErr = requireExamineeToken(p);
  if (swaTokenErr) return swaTokenErr;
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
  var swabTokenErr = requireExamineeToken(data);
  if (swabTokenErr) return swabTokenErr;
  // Receive ALL wrong answers in a single POST and write to result row
  var sheet = getSheet('תוצאות');
  var rows = sheet.getDataRange().getValues();
  for (var i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][13]) === String(data.sessionCode) && normalizeId(rows[i][1]) === normalizeId(data.idNumber)) {
      // Detect the bug pattern where client sent "undefined" because it doesn't
      // know correct answers (handleGetExamQuestions strips `ci` from examinee
      // responses). If client data is bogus AND the sheet already has good data
      // (written by handleSubmitResult's server-side rebuild), keep the sheet's version.
      var clientHasBogus = false;
      if (data.wrongAnswers && data.wrongAnswers.length > 0) {
        for (var bi = 0; bi < data.wrongAnswers.length; bi++) {
          var ca = String((data.wrongAnswers[bi] && data.wrongAnswers[bi].correctAnswer) || '');
          if (ca.indexOf('undefined') !== -1) { clientHasBogus = true; break; }
        }
      }
      var existingWrong = String(rows[i][15] || '');
      if (clientHasBogus && existingWrong && existingWrong.indexOf('undefined') === -1) {
        // Sheet already has authoritative data → keep it, skip overwrite.
        SpreadsheetApp.flush();
        return jsonResponse({ status: 'ok', skipped: true, reason: 'client_bogus_server_good' });
      }

      var wrongDetails = '';
      var wrongForWA = '';
      if (data.wrongAnswers && data.wrongAnswers.length > 0) {
        for (var w = 0; w < data.wrongAnswers.length; w++) {
          var item = data.wrongAnswers[w];
          // Question ID prefix — see comment in handleSubmitWrongAnswers above.
          if (item.questionId) wrongDetails += 'מזהה שאלה: ' + item.questionId + '\n';
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
  var focTokenErr = requireExamineeToken(data);
  if (focTokenErr) return focTokenErr;
  var sheet = getSheet('תוצאות');

  // Do NOT record a close-fail for an examinee an examiner reset/removed
  // (status 'cancelled') or 'rejected' — reset semantics are "won't count as a
  // fail". A clean finish of such an examinee IS still recorded (submitResult
  // accepts 'cancelled'); only the auto-0/30-on-close is suppressed here.
  try {
    var focPend = getSheet('ממתינים').getDataRange().getValues();
    for (var fp = focPend.length - 1; fp >= 1; fp--) {
      if (String(focPend[fp][0]) === String(data.sessionCode) && normalizeId(focPend[fp][1]) === normalizeId(data.idNumber)) {
        var fpStatus = String(focPend[fp][5]).trim();
        if (fpStatus === 'cancelled' || fpStatus === 'rejected') return jsonResponse({ status: 'ok', skipped: 'cancelled' });
        break;
      }
    }
  } catch (focErr) {}

  // Duplicate protection: if ANY non-בוטל result already exists for this session+id,
  // do NOT add a close-fail. Match on session+id ONLY (not language/license): the
  // close-beacon carries the REGISTRATION language, but a real submit may carry a
  // different FINAL language after a mid-exam switch — a language-scoped check would
  // miss it and append a spurious 0/30 next to the real result.
  var existingData = sheet.getDataRange().getValues();
  for (var d = 1; d < existingData.length; d++) {
    if (String(existingData[d][13]) === String(data.sessionCode) && normalizeId(existingData[d][1]) === normalizeId(data.idNumber)) {
      if (String(existingData[d][7] || '').trim() === 'בוטל') continue; // a voided row is not a real result
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
    data.audioMode || 'off',
    '', '', '', '', '', '', '',         // idx 22-28 (מאומת..מסלול שפות) — N/A for a close-fail row
    String(data.device || '')           // AD (29) מכשיר — phone / tablet / desktop
  ]);

  markPendingCompleted(data.sessionCode, data.idNumber);

  return jsonResponse({ status: 'ok' });
}

function handleCancelFailOnClose(data) {
  // Called when page reloads (refresh, not actual close) — undo the fail
  var cfocTokenErr = requireExamineeToken(data);
  if (cfocTokenErr) return cfocTokenErr;
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
        // Examinee resumed — the fail-on-close was premature. Mark בוטל instead
        // of DELETING: sheet.deleteRow was the ONLY path that could ever destroy a
        // result row (latent "vanished result" vector). At most one בוטל row
        // results, since submitFailOnClose's dup-guard blocks further close-fails.
        sheet.getRange(r + 1, 8).setValue('בוטל');
        sheet.getRange(r + 1, 27).setValue('בוטל אוטומטי - רענון/חזרה למבחן');
        sheet.getRange(r + 1, 28).setValue(todayStr());
        // Also un-mark pending as completed so exam can continue
        unmarkPendingCompleted(sc, id);
      }
      break; // only check the most recent match
    }
  }
  return jsonResponse({ status: 'ok' });
}

// ========== Question-cache warmup (for scheduled trigger) ==========
// Apps Script's CacheService keeps entries for up to 6 hours. Loading the
// 7 language files from Drive on a cold cache takes ~10-15 sec, which is
// the main reason exam-start feels slow for the first user after a long
// idle period. This function pre-loads every language file into cache so
// real user requests always hit warm cache.
//
// Setup (one-time): in Apps Script editor →
//   Triggers (clock icon, left sidebar) → Add Trigger
//   Function: warmupQuestionCaches
//   Event source: Time-driven
//   Type: Hour timer
//   Every: 4 hours
//   Save (you'll be asked to authorize)
//
// After that, the cache is continuously warm; users always see ~3-5 sec
// exam-start instead of 15-20 sec.
function warmupQuestionCaches() {
  var LANGS = ['he', 'ru', 'en', 'ar', 'fr', 'es', 'am'];
  var summary = [];
  for (var i = 0; i < LANGS.length; i++) {
    var t0 = Date.now();
    try {
      var data = loadQuestionsForLanguageServer(LANGS[i]);
      var n = data ? data.length : 0;
      // Flag suspicious results loudly so they're easy to spot in the log.
      // loadQuestionsForLanguageServer already refuses to cache empty results,
      // but the warmup is the right place to notice that the Drive file
      // itself went bad (was edited mid-day, partial upload, etc.).
      var flag = (n === 0) ? '  ⚠️ EMPTY — check Drive source file' : '';
      summary.push(LANGS[i] + ': ' + n + ' questions in ' + (Date.now() - t0) + 'ms' + flag);
    } catch (e) {
      summary.push(LANGS[i] + ': ERROR - ' + (e && e.message ? e.message : e));
    }
  }
  Logger.log('warmupQuestionCaches complete:\n' + summary.join('\n'));
  return summary;
}

// ========== Emergency cache reset ==========
// Run this manually from the Apps Script editor when you see "0 questions"
// (or similar nonsense from a poisoned cache). Clears every cached language
// chunk, then re-warms straight from Drive. Returns a human-readable report.
//
// Real-world trigger: during one exam day, the Hebrew chunks were cached as
// empty after a race condition between two parallel loads. Every subsequent
// examinee got "0 questions" in Hebrew until the cache TTL expired. This
// function fixes that in 5 seconds without waiting for TTL.
//
// Apps Script editor → select function: emergencyClearAndRefreshCache → Run.
// Then check Logger output (View → Logs or "Execution log" panel).
function emergencyClearAndRefreshCache() {
  var cache = CacheService.getScriptCache();
  var langs = ['he', 'ru', 'en', 'ar', 'fr', 'es', 'am'];
  var report = ['=== Step 1: clearing cached chunks ==='];
  langs.forEach(function(lang) {
    var metaKey = 'qdata_' + lang + '_meta';
    var meta = cache.get(metaKey);
    if (meta) {
      var n = parseInt(meta, 10) || 0;
      var keys = [metaKey];
      for (var i = 0; i < n; i++) keys.push('qdata_' + lang + '_part_' + i);
      cache.removeAll(keys);
      report.push(lang + ': removed ' + keys.length + ' keys');
    } else {
      report.push(lang + ': no cache to clear');
    }
  });
  report.push('');
  report.push('=== Step 2: refreshing from Drive ===');
  langs.forEach(function(lang) {
    var t0 = Date.now();
    try {
      var data = loadQuestionsForLanguageServer(lang);
      var n = data ? data.length : 0;
      var flag = (n === 0) ? '  ⚠️ STILL EMPTY — check Drive file questions_' + lang + '.json' : '';
      report.push(lang + ': ' + n + ' questions (' + (Date.now() - t0) + 'ms)' + flag);
    } catch (e) {
      report.push(lang + ': ERROR - ' + (e && e.message ? e.message : e));
    }
  });
  var out = report.join('\n');
  Logger.log(out);
  return out;
}

// ========== Server-side question delivery ==========
// Loads question data from a private Google Drive folder (one JSON file per
// language) and returns a curated 30-question exam to authenticated clients.
// The full question bank never reaches the browser — only the questions for
// the current exam, without the correct-answer index.
//
// Setup:
//   1) Run deployment/generate_questions_data.js locally to produce
//      deployment/generated/questions_<lang>.json files.
//   2) Upload all 7 files to a private Drive folder (only this account
//      should have access; do NOT share publicly).
//   3) Copy the folder ID (the long string in the Drive URL) and set it
//      as ScriptProperty: QUESTIONS_DRIVE_FOLDER_ID = <folder-id>
//   4) Deploy this Apps Script.
//   5) Push updated HTMLs (examinee, exam, student, find_image) so they call
//      getExamQuestions instead of loading questions.js.

var EXAM_STRUCTURE_SERVER = {
  'B':  { 'בטיחות': 7, 'הכרת הרכב': 7, 'חוק': 7, 'תמרורים': 9 },
  '1':  { 'בטיחות': 5, 'הכרת הרכב': 5, 'חוק': 6, 'תמרורים': 6, 'ספציפי': 8 },
  'C1': { 'בטיחות': 5, 'הכרת הרכב': 5, 'חוק': 5, 'תמרורים': 5, 'ספציפי': 10 },
  'C':  { 'בטיחות': 5, 'הכרת הרכב': 4, 'חוק': 3, 'תמרורים': 4, 'ספציפי': 14 },
  'D':  { 'בטיחות': 4, 'הכרת הרכב': 2, 'חוק': 5, 'תמרורים': 4, 'ספציפי': 15 }
};

function classifyCategoryServer(cat) {
  var c = String(cat || '').trim();
  if (/ספציפי/.test(c)) return 'ספציפי'; // ספציפי
  if (/בטיחות/.test(c)) return 'בטיחות'; // בטיחות
  if (/הכרת הרכב/.test(c)) return 'הכרת הרכב'; // הכרת הרכב
  if (/חוק/.test(c)) return 'חוק'; // חוק
  if (/תמרורים/.test(c)) return 'תמרורים'; // תמרורים
  if (/זכות קדימה/.test(c)) return 'חוק'; // זכות קדימה → חוק
  return '';
}

function filterByLicenseServer(pool, license) {
  return pool.filter(function(q) {
    var cat = String(q.category || '');
    // "מתן זכות קדימה" applies to all license types
    if (/זכות קדימה/.test(cat)) return true;
    if (license === '1') {
      var lt = String(q.licenseType || '').trim();
      if (lt !== '' && lt !== 'N/A') return false;
      if (cat.indexOf('1') === -1) return false;
      return true;
    }
    if (license === 'C') {
      var lic = String(q.licenseType || '').trim();
      return lic === 'C' || lic === 'C/E' || lic === 'C+E' || lic === 'CE';
    }
    var lic2 = String(q.licenseType || '').trim();
    return lic2 === license;
  });
}

function shuffleArrayServer(arr) {
  var a = arr.slice();
  for (var i = a.length - 1; i > 0; i--) {
    var j = Math.floor(Math.random() * (i + 1));
    var t = a[i]; a[i] = a[j]; a[j] = t;
  }
  return a;
}

// Read questions for a given language. Uses chunked CacheService caching so
// subsequent calls within 6h don't hit Drive again. Cold start: ~2-4 sec
// (Drive read + JSON.parse). Warm: ~300 ms (cache reassembly).
function loadQuestionsForLanguageServer(lang) {
  var safeLang = String(lang || 'he').toLowerCase();
  if (!/^[a-z]{2}$/.test(safeLang)) throw new Error('Invalid language code');

  var cache = CacheService.getScriptCache();
  var metaKey = 'qdata_' + safeLang + '_meta';

  // Reassemble the cached chunks into the questions array, or return null on a
  // miss / partial / empty-poisoned cache. Factored out so both the initial
  // read AND the stampede-waiters below can re-check without duplicating logic.
  function readFromCache() {
    var meta = cache.get(metaKey);
    if (!meta) return null;
    var numChunks = parseInt(meta, 10);
    var keys = [];
    for (var i = 0; i < numChunks; i++) keys.push('qdata_' + safeLang + '_part_' + i);
    var chunks = cache.getAll(keys);
    var json = '';
    for (var k = 0; k < numChunks; k++) {
      var c = chunks['qdata_' + safeLang + '_part_' + k];
      if (c === null || c === undefined) return null;  // partial → treat as miss
      json += c;
    }
    try {
      var cached = JSON.parse(json);
      // Cache-poisoning guard: empty array means a previous write captured a
      // partial/empty file (real exam-day incident: Hebrew was cached as []
      // after a race). Treat empty as a miss and re-read from Drive.
      if (Array.isArray(cached) && cached.length > 0) return cached;
      Logger.log('[CACHE] empty cached result for ' + safeLang + ' — ignoring and re-reading from Drive');
    } catch (e) { /* fall through to Drive */ }
    return null;
  }

  var hit = readFromCache();
  if (hit) return hit;

  // ===== Stampede protection (fixes the ~13:00 daily outage) =====
  // The question cache has a 6h TTL. When it expires, a whole exam-start wave
  // hits a COLD cache at the same instant. Without this guard, EVERY request
  // reads the big Drive file simultaneously — each taking tens of seconds —
  // which saturates Apps Script's ~30 execution slots. That is exactly what
  // froze exam-start + result-submit + examiner sync for ~15 min until one
  // read finally re-warmed the cache. Fix: only ONE execution reads Drive; the
  // rest wait briefly and reuse the freshly-warmed cache. A cold-cache moment
  // becomes a ~3s blip for one call instead of a 15-min outage for everyone.
  var lockKey = 'qload_lock_' + safeLang;
  var haveLock = false;
  if (!cache.get(lockKey)) {
    cache.put(lockKey, '1', 60);   // hold the loader lock up to 60s
    haveLock = true;
  }
  if (!haveLock) {
    // Another execution is already loading — wait for it to warm the cache
    // instead of stampeding Drive ourselves.
    for (var w = 0; w < 10; w++) {
      Utilities.sleep(1000);
      var warmed = readFromCache();
      if (warmed) return warmed;
      if (!cache.get(lockKey)) break;  // loader finished/failed — load it ourselves
    }
    // Timed out or the loader vanished: fall through and read Drive as a fallback.
  }

  try {
    // Cache miss → read from Drive
    var folderId = PropertiesService.getScriptProperties().getProperty('QUESTIONS_DRIVE_FOLDER_ID');
    if (!folderId) {
      throw new Error('QUESTIONS_DRIVE_FOLDER_ID not configured in ScriptProperties');
    }
    var folder;
    try { folder = DriveApp.getFolderById(folderId); }
    catch (e) { throw new Error('Cannot access Drive folder: ' + e.message); }

    var fileName = 'questions_' + safeLang + '.json';
    var files = folder.getFilesByName(fileName);
    if (!files.hasNext()) throw new Error(fileName + ' not found in Drive folder');
    var file = files.next();

    var jsonStr = file.getBlob().getDataAsString('UTF-8');

    // Parse FIRST, then decide whether to cache. We never cache an empty result
    // — that's how Hebrew got stuck at 0 questions for hours after a race.
    var parsed;
    try { parsed = JSON.parse(jsonStr); }
    catch (e) { throw new Error('Failed to parse ' + fileName + ' from Drive: ' + e.message); }

    if (Array.isArray(parsed) && parsed.length > 0) {
      // Write back to cache in chunks (CacheService cap: 100 KB per key)
      var CHUNK_SIZE = 90000;
      var totalChunks = Math.ceil(jsonStr.length / CHUNK_SIZE);
      var putMap = {};
      for (var pi = 0; pi < totalChunks; pi++) {
        putMap['qdata_' + safeLang + '_part_' + pi] = jsonStr.substr(pi * CHUNK_SIZE, CHUNK_SIZE);
      }
      putMap[metaKey] = String(totalChunks);
      try { cache.putAll(putMap, 21600); } catch (e) { /* cache full or unavailable — proceed without */ }
    } else {
      // Don't poison the cache. Log so we can spot a corrupted source file.
      Logger.log('[DRIVE] ' + fileName + ' parsed to empty/non-array — NOT caching. Check the source file.');
    }

    return parsed;
  } finally {
    if (haveLock) cache.remove(lockKey);
  }
}

// Pick 30 questions per the license blueprint, return them WITHOUT the
// correct-answer index. Authenticated clients only — falls back to a
// rate-limited guest path for the standalone exam.html flow.
function handleGetExamQuestions(p) {
  // Determine auth context
  var auth = 'guest';
  if (p.sessionCode && p.idNumber && p.examineeToken) {
    var ev = verifyExamineeToken(p.sessionCode, p.idNumber, p.examineeToken);
    if (!ev.valid) {
      return jsonResponse({ status: 'error', message: 'Examinee token invalid', reason: ev.reason });
    }
    auth = 'examinee';
  } else if (p.token && p.examinerId) {
    if (!verifyToken(p.examinerId, p.token)) {
      return jsonResponse({ status: 'error', message: 'Examiner token invalid', tokenExpired: true });
    }
    auth = 'examiner';
  } else if (p.classCode && p.studentId) {
    // Student practice mode — looser auth, just rate-limit
    auth = 'student';
  } else if (p.standaloneIdNumber) {
    // Standalone exam.html — examinee enters their ID, no token; rate-limit hard
    auth = 'standalone';
  }

  // Rate limit (per auth + identifier)
  var rlId = p.sessionCode || p.idNumber || p.examinerId || p.studentId || p.standaloneIdNumber || 'anon';
  var rlMax = (auth === 'guest' || auth === 'standalone') ? 5 : 20;
  var rlErr = requireRateLimit('getExamQuestions_' + auth, rlId, rlMax, 60);
  if (rlErr) return rlErr;

  var lang = String(p.language || 'he').toLowerCase();
  var license = String(p.license || p.licenseType || 'B');
  if (!EXAM_STRUCTURE_SERVER[license]) {
    return jsonResponse({ status: 'error', message: 'Unknown license: ' + license });
  }

  var allQuestions;
  try { allQuestions = loadQuestionsForLanguageServer(lang); }
  catch (e) {
    Logger.log('loadQuestionsForLanguageServer(' + lang + ') failed: ' + (e && e.message));
    return jsonResponse({ status: 'error', message: 'שגיאה בטעינת שאלות. נסה שוב.' });
  }

  // Order matters: filter by license BEFORE dedupe. The source data has the
  // same question id repeated for multiple license types (e.g. id 1276 appears
  // once with licenseType=B and once with C1). If we dedupe first, we might
  // keep the C1 row and then the license filter rejects it. Filter first so
  // we only see rows that already match the license, then dedupe within that.
  var filtered = filterByLicenseServer(allQuestions, license);
  var seen = {};
  var pool = [];
  for (var i = 0; i < filtered.length; i++) {
    var q = filtered[i];
    if (!q || !q.id || seen[q.id]) continue;
    if (!Array.isArray(q.answers) || q.answers.length < 2) continue;
    seen[q.id] = true;
    pool.push(q);
  }

  // Category-quiz mode (student.html practice by topic): return up to N
  // questions matching one category, skipping the 30-question blueprint.
  var mode = String(p.mode || 'exam');
  var selected;
  if (mode === 'category' && p.categoryFilter) {
    var wantTopic = String(p.categoryFilter);
    var catPool = pool.filter(function(q) { return classifyCategoryServer(q.category) === wantTopic; });
    catPool = shuffleArrayServer(catPool);
    var max = Number(p.maxCount) || 15;
    selected = catPool.slice(0, Math.min(catPool.length, max));
    if (selected.length === 0) {
      return jsonResponse({ status: 'error', message: 'No questions in category', topic: wantTopic });
    }
  } else {
    // Default exam mode: pick per EXAM_STRUCTURE blueprint.
    var byTopic = {};
    for (var j = 0; j < pool.length; j++) {
      var t = classifyCategoryServer(pool[j].category);
      if (!t) continue;
      if (!byTopic[t]) byTopic[t] = [];
      byTopic[t].push(pool[j]);
    }
    var blueprint = EXAM_STRUCTURE_SERVER[license];
    selected = [];
    var usedIds = {};
    for (var topic in blueprint) {
      var needed = blueprint[topic];
      var avail = shuffleArrayServer(byTopic[topic] || []);
      var count = 0;
      for (var ai = 0; ai < avail.length && count < needed; ai++) {
        if (usedIds[avail[ai].id]) continue;
        usedIds[avail[ai].id] = true;
        selected.push(avail[ai]);
        count++;
      }
      if (count < needed) {
        return jsonResponse({
          status: 'error',
          message: 'Not enough questions for topic',
          topic: topic,
          have: (byTopic[topic] || []).length,
          need: needed
        });
      }
    }
    selected = shuffleArrayServer(selected);
  }

  // Remember which IDs we issued so handleRegisterExamQuestions can refuse
  // submissions that name questions we didn't actually give the examinee.
  // Only meaningful for the examinee-token path (we have a stable identifier).
  if (auth === 'examinee' && p.sessionCode && p.idNumber) {
    var issuedKey = 'issued_qs_' + p.sessionCode + '_' + normalizeId(p.idNumber);
    var ids = selected.map(function(q) { return q.id; });
    try { CacheService.getScriptCache().put(issuedKey, JSON.stringify(ids), 21600); } catch (e) { /* skip */ }
  }

  // For real exams (auth === 'examinee') we deliberately omit the correct
  // index — scoring happens server-side via handleRegisterExamQuestions /
  // handleSubmitResult using ANSWER_KEY_BY_LANG. For practice/standalone
  // flows (exam.html, student.html) the client needs to score locally, so
  // we include the encoded `ci` field that matches the legacy questions.js
  // format: ci = correctIndex XOR (id mod 256).
  if (auth !== 'examinee' && typeof lookupCorrectIndex === 'function') {
    for (var ci_i = 0; ci_i < selected.length; ci_i++) {
      var origCorrect = lookupCorrectIndex(Number(selected[ci_i].id), lang);
      if (origCorrect !== null && origCorrect !== undefined) {
        selected[ci_i].ci = origCorrect ^ (selected[ci_i].id % 256);
      }
    }
  }

  // Pre-fetch all 7 languages so mid-exam language switches are instant
  // (no extra round trip). Adds ~120-150 KB to the response. Cold-start
  // server cost is real (7 Drive reads) but cached for 6h after that.
  //
  // For any non-examinee caller we include each translation's `ci` (correct
  // index, XOR-encoded) so the client can score correctly per language when
  // the translator put answers in a different order. This includes `guest`
  // (student practicing without a classCode) — the previous student-only gate
  // left guest practice silently wrong after a mid-practice language switch.
  // For examinee auth we keep `ci` stripped — server is sole source of truth.
  var includeCiInTranslations = (auth !== 'examinee');
  var translations = null;
  if (p.includeTranslations === 'true' || p.includeTranslations === '1') {
    translations = {};
    var SUPPORTED_LANGS = ['he', 'ru', 'en', 'ar', 'fr', 'es', 'am'];
    var idSet = {};
    for (var ix = 0; ix < selected.length; ix++) idSet[selected[ix].id] = true;
    for (var li = 0; li < SUPPORTED_LANGS.length; li++) {
      var altLang = SUPPORTED_LANGS[li];
      try {
        var altData = loadQuestionsForLanguageServer(altLang);
        var altMap = {};
        for (var ai = 0; ai < altData.length; ai++) {
          var aq = altData[ai];
          if (aq && idSet[aq.id]) {
            var entry = { t: aq.text, a: aq.answers };
            // Source JSONs don't carry `ci` — look it up per-language from the
            // answer key so a mid-exam language switch can rewire the correct
            // index to whatever order the translator put answers in.
            if (includeCiInTranslations && typeof lookupCorrectIndex === 'function') {
              var altCorrect = lookupCorrectIndex(Number(aq.id), altLang);
              if (altCorrect !== null && altCorrect !== undefined) {
                entry.ci = altCorrect ^ (aq.id % 256);
              }
            }
            altMap[aq.id] = entry;
          }
        }
        translations[altLang] = altMap;
      } catch (e) { /* language file missing — skip */ }
    }
  }

  var responseBody = { status: 'ok', auth: auth, count: selected.length, questions: selected };
  if (translations) responseBody.translations = translations;
  return jsonResponse(responseBody);
}

// ========== Re-fetch questions in a different language ==========
// When an examinee/student/practice user changes language mid-exam, the
// client calls this with the set of question IDs already shown and the new
// language. Server returns those same IDs with text/answers in the new
// language so the exam can continue without losing progress.
function handleGetQuestionsByIds(p) {
  // Match auth model of handleGetExamQuestions
  var auth = 'guest';
  if (p.sessionCode && p.idNumber && p.examineeToken) {
    var ev = verifyExamineeToken(p.sessionCode, p.idNumber, p.examineeToken);
    if (!ev.valid) {
      return jsonResponse({ status: 'error', message: 'Examinee token invalid', reason: ev.reason });
    }
    auth = 'examinee';
  } else if (p.token && p.examinerId) {
    if (!verifyToken(p.examinerId, p.token)) {
      return jsonResponse({ status: 'error', message: 'Examiner token invalid', tokenExpired: true });
    }
    auth = 'examiner';
  } else if (p.classCode && p.studentId) {
    auth = 'student';
  } else if (p.standaloneIdNumber) {
    auth = 'standalone';
  }

  var rlErr = requireRateLimit('getQuestionsByIds_' + auth,
    p.sessionCode || p.idNumber || p.examinerId || p.studentId || p.standaloneIdNumber || 'anon',
    30, 60);
  if (rlErr) return rlErr;

  var lang = String(p.language || 'he').toLowerCase();
  var idsRaw = String(p.ids || '');
  var ids = idsRaw.split(',').map(function(s) {
    var n = parseInt(String(s).trim(), 10);
    return isNaN(n) ? null : n;
  }).filter(function(n) { return n !== null; });

  if (ids.length === 0) return jsonResponse({ status: 'error', message: 'No IDs provided' });
  if (ids.length > 50) return jsonResponse({ status: 'error', message: 'Too many IDs (max 50)' });

  var allQuestions;
  try { allQuestions = loadQuestionsForLanguageServer(lang); }
  catch (e) {
    Logger.log('loadQuestionsForLanguageServer(' + lang + ') failed: ' + (e && e.message));
    return jsonResponse({ status: 'error', message: 'שגיאה בטעינת שאלות. נסה שוב.' });
  }

  // Build id → question lookup
  var byId = {};
  for (var i = 0; i < allQuestions.length; i++) {
    var q = allQuestions[i];
    if (q && q.id) byId[q.id] = q;
  }

  var results = [];
  for (var j = 0; j < ids.length; j++) {
    var found = byId[ids[j]];
    if (!found) {
      results.push(null);
      continue;
    }
    var entry = {
      id: found.id,
      text: found.text,
      answers: found.answers,
      category: found.category,
      licenseType: found.licenseType,
      imageUrl: found.imageUrl,
      language: found.language || lang
    };
    // Include ci for non-examinee callers (practice/standalone) so they can score locally.
    if (auth !== 'examinee' && typeof lookupCorrectIndex === 'function') {
      var origCorrect = lookupCorrectIndex(Number(found.id), lang);
      if (origCorrect !== null && origCorrect !== undefined) {
        entry.ci = origCorrect ^ (found.id % 256);
      }
    }
    results.push(entry);
  }

  return jsonResponse({ status: 'ok', count: results.length, questions: results });
}

// ========== Question search (for find_image.html examiner utility) ==========
// Returns up to 20 questions whose text/answers/category match the query
// substring. Requires examiner token — this is an internal staff utility,
// not for examinees. Cross-language search: caller passes the language and
// we search that language's pre-translated dataset.
function handleSearchQuestions(p) {
  var authErr = requireToken(p);
  if (authErr) return authErr;
  var rlErr = requireRateLimit('searchQuestions', String(p.examinerId || ''), 30, 60);
  if (rlErr) return rlErr;

  var query = String(p.q || '').trim().toLowerCase();
  if (query.length < 2) return jsonResponse({ status: 'ok', matches: [], note: 'Query too short' });

  var lang = String(p.language || 'he').toLowerCase();
  var allQuestions;
  try { allQuestions = loadQuestionsForLanguageServer(lang); }
  catch (e) {
    Logger.log('loadQuestionsForLanguageServer(' + lang + ') failed: ' + (e && e.message));
    return jsonResponse({ status: 'error', message: 'שגיאה בטעינת שאלות. נסה שוב.' });
  }

  var matches = [];
  var MAX_MATCHES = 20;
  for (var i = 0; i < allQuestions.length && matches.length < MAX_MATCHES; i++) {
    var q = allQuestions[i];
    if (!q || !q.text) continue;
    var hay = (q.text + ' ' + (q.answers || []).join(' ') + ' ' + (q.category || '')).toLowerCase();
    if (hay.indexOf(query) === -1) continue;
    var match = {
      id: q.id,
      text: q.text,
      answers: q.answers,
      category: q.category,
      licenseType: q.licenseType,
      imageUrl: q.imageUrl
    };
    // Include encoded ci for the search utility (examiner trusted view)
    if (typeof lookupCorrectIndex === 'function') {
      var orig = lookupCorrectIndex(Number(q.id), lang);
      if (orig !== null && orig !== undefined) {
        match.ci = orig ^ (q.id % 256);
      }
    }
    matches.push(match);
  }

  return jsonResponse({ status: 'ok', matches: matches, language: lang, query: query });
}

// ========== Bohan-site (IDF examiners portal) — server-side data + auth ==========
// The names, Waze links, and external URLs used by bohan-site.pages.dev live
// here on the server. The HTML on Pages is a thin shell that asks the user
// for the shared password, then fetches this data via bohanSiteAuth and
// renders it. Without the password, the browser receives no sensitive data
// — protects both the live Pages URL AND the public GitHub source.
//
// Setup:
//   1) Set ScriptProperty BOHAN_SITE_PASSWORD = <shared examiner password>
//   2) Set ScriptProperty BOHAN_SITE_SECRET = <random 32-char HMAC secret>
//   3) Deploy. Push the updated bohan-site/index.html to GitHub.

var BOHAN_SITE_EXAMINERS = [
  'אללוף יצחק',
  'גיטלמן ויטלי',
  'אלבלה אברהם',
  'נגאוקר שמשון',
  'מלל דרור',
  'קורנבליט אלכס',
  'שרון צרימי',
  'יניר נאוגאוקר',
  'פלס דוד',
  'לוי בנימין',
  'רון יהוד'
];

var BOHAN_SITE_LOCATIONS = [
  { name: 'אופקים', wazeUrl: 'https://ul.waze.com/ul?ll=31.32266385%2C34.62266207&navigate=yes&zoom=17&utm_campaign=default&utm_source=waze_website&utm_medium=lm_share_location' },
  { name: 'אשדוד', wazeUrl: 'https://waze.com/ul/hsv8sudz4t' },
  { name: 'אשקלון', wazeUrl: 'https://ul.waze.com/ul?ll=31.66465059%2C34.55991983&navigate=yes&zoom=17&utm_campaign=default&utm_source=waze_website&utm_medium=lm_share_location' },
  { name: 'אילת', wazeUrl: 'https://waze.com/ul/hsv2b5mzn7' },
  { name: 'באר שבע', wazeUrl: 'https://ul.waze.com/ul?ll=31.24712197%2C34.76847768&navigate=yes&zoom=17&utm_campaign=default&utm_source=waze_website&utm_medium=lm_share_location' },
  { name: 'באר שבע בית החייל', wazeUrl: 'https://www.ufis.org.il/?categoryId=123860' },
  { name: 'באר שבע ל"ג', wazeUrl: 'https://waze.com/ul/hsv89zc1eu' },
  { name: 'בח"א 6 חצרים', wazeUrl: 'https://ul.waze.com/ul?ll=31.24938755%2C34.66291666&navigate=yes&zoom=17&utm_campaign=default&utm_source=waze_website&utm_medium=lm_share_location' },
  { name: 'בח"א 8 תל נוף', wazeUrl: 'https://ul.waze.com/ul?ll=31.83986809%2C34.81412888&navigate=yes&zoom=14&utm_campaign=default&utm_source=waze_website&utm_medium=lm_share_location' },
  { name: 'פלמחים', wazeUrl: 'https://ul.waze.com/ul?ll=31.92554940%2C34.71307397&navigate=yes&zoom=17&utm_campaign=default&utm_source=waze_website&utm_medium=lm_share_location' },
  { name: 'ביס"ט 21 טכני', wazeUrl: 'https://ul.waze.com/ul?ll=32.86343939%2C35.04089355&navigate=yes&zoom=9&utm_campaign=default&utm_source=waze_website&utm_medium=lm_share_location' },
  { name: 'בית נבאללה', wazeUrl: 'https://waze.com/ul/hsv8vgj8r9' },
  { name: 'בית שמש', wazeUrl: 'https://waze.com/ul/hsv8us7hey' },
  { name: 'בהל"צ', wazeUrl: 'https://waze.com/ul/hsv2fff46z' },
  { name: 'דימונה', wazeUrl: 'https://waze.com/ul/hsv8btpv88' },
  { name: 'הדר יוסף', wazeUrl: 'https://waze.com/ul/hsv8y8hwnu' },
  { name: 'חדרה', wazeUrl: 'https://waze.com/ul/hsvbb6zm8x' },
  { name: 'חיפה', wazeUrl: 'https://waze.com/ul/hsvbfe9sg0' },
  { name: 'חיפה ל"ג', wazeUrl: 'https://waze.com/ul/hsvbft3gwz' },
  { name: 'טבריה', wazeUrl: 'https://ul.waze.com/ul?ll=32.78839400%2C35.53756700&navigate=yes&utm_campaign=share_drive&utm_source=waze_app&utm_medium=lm_share_location' },
  { name: 'טבריה ל"ג', wazeUrl: 'https://waze.com/ul/hsvc62ppwf' },
  { name: 'יפו', wazeUrl: 'https://waze.com/ul/hsv8wr1gvv' },
  { name: 'ירושלים', wazeUrl: 'https://waze.com/ul/hsv9h8u42r' },
  { name: 'ירושלים ל"ג', wazeUrl: 'https://waze.com/ul/hsv9h9ryht' },
  { name: 'כנף 1 רמת דוד', wazeUrl: 'https://waze.com/ul/hsvc1b53fy' },
  { name: 'כנף 25 רמון', wazeUrl: 'https://waze.com/ul/hsv2xhgew9' },
  { name: 'כנף 28 נבטים', wazeUrl: 'https://waze.com/ul/hsv8ct9tb0' },
  { name: 'כפר סבא', wazeUrl: 'https://waze.com/ul/hsv8yfxfp0' },
  { name: 'לוד', wazeUrl: 'https://waze.com/ul/hsv8v9vemx' },
  { name: 'מחנה עמוס', wazeUrl: 'https://waze.com/ul/hsvc16rd7e' },
  { name: 'משמר הנגב', wazeUrl: 'https://waze.com/ul/hsv8djxbxr' },
  { name: 'נתיבות', wazeUrl: 'https://waze.com/ul/hsv8ddwrsd' },
  { name: 'נתניה', wazeUrl: 'https://waze.com/ul/hsv8zcd4ss' },
  { name: 'עיר הבהדים', wazeUrl: 'https://waze.com/ul/hsv8b8y09k' },
  { name: 'עכו', wazeUrl: 'https://waze.com/ul/hsvbgq3ccg' },
  { name: 'עפולה', wazeUrl: 'https://waze.com/ul/hsvc17pke7' },
  { name: 'עפולה משא כבד', wazeUrl: 'https://waze.com/ul/hsvc1ed03e' },
  { name: 'פתח תקווה', wazeUrl: 'https://waze.com/ul/hsv8y9kdg1' },
  { name: 'קרית גת', wazeUrl: 'https://waze.com/ul/hsv8ez31n2' },
  { name: 'קרית חיים', wazeUrl: 'https://waze.com/ul/hsvbftyx0z' },
  { name: 'קרית שמונה', wazeUrl: 'https://waze.com/ul/hsvckc7vk8' },
  { name: 'ראשון לציון', wazeUrl: 'https://waze.com/ul/hsv8tzcr54' },
  { name: 'רחובות', wazeUrl: 'https://waze.com/ul/hsv8trzeke' },
  { name: 'תה"ש ל"ג', wazeUrl: 'https://waze.com/ul/hsv8wrkzz3' }
];

var BOHAN_SITE_SHEETS_URL = 'https://docs.google.com/spreadsheets/d/1KAX96KcGNQU7aOS7lf6oMZY-FiDuJ_j5npcDSVfiW5E/edit?usp=sharing';
var BOHAN_SITE_SITES_URL = 'https://sites.google.com/view/bohanyzahal/%D7%91%D7%99%D7%AA';

function _bohanSiteData() {
  return {
    examiners: BOHAN_SITE_EXAMINERS,
    locations: BOHAN_SITE_LOCATIONS,
    sheetsUrl: BOHAN_SITE_SHEETS_URL,
    sitesUrl: BOHAN_SITE_SITES_URL
  };
}

function _bohanSiteIssueToken(secret) {
  var payload = JSON.stringify({ exp: Date.now() + 24 * 60 * 60 * 1000 });
  var pB64 = Utilities.base64EncodeWebSafe(payload).replace(/=+$/, '');
  var sigBytes = Utilities.computeHmacSha256Signature(pB64, secret);
  var sigB64 = Utilities.base64EncodeWebSafe(sigBytes).replace(/=+$/, '');
  return pB64 + '.' + sigB64;
}

function _bohanSiteVerifyToken(token, secret) {
  if (!token || token.indexOf('.') < 0) return false;
  var parts = String(token).split('.');
  if (parts.length !== 2) return false;
  var pB64 = parts[0];
  var sigB64 = parts[1];
  var sigBytes = Utilities.computeHmacSha256Signature(pB64, secret);
  var expectedB64 = Utilities.base64EncodeWebSafe(sigBytes).replace(/=+$/, '');
  if (sigB64 !== expectedB64) return false;
  try {
    var pad = (4 - (pB64.length % 4)) % 4;
    var decoded = Utilities.base64DecodeWebSafe(pB64 + Array(pad + 1).join('='));
    var payload = JSON.parse(Utilities.newBlob(decoded).getDataAsString());
    if (typeof payload.exp !== 'number' || Date.now() > payload.exp) return false;
    return true;
  } catch (e) { return false; }
}

function handleBohanSiteAuth(p) {
  // Light rate limit
  var rlErr = requireRateLimit('bohanSiteAuth', p.token ? 'tok' : 'pwd', 20, 60);
  if (rlErr) return rlErr;

  var props = PropertiesService.getScriptProperties();
  var secret = props.getProperty('BOHAN_SITE_SECRET');
  var password = props.getProperty('BOHAN_SITE_PASSWORD');
  if (!secret || !password) {
    return jsonResponse({ status: 'error', message: 'BOHAN_SITE_PASSWORD / BOHAN_SITE_SECRET not configured' });
  }

  // Path A: token-based access (subsequent visits)
  if (p.token) {
    if (!_bohanSiteVerifyToken(p.token, secret)) {
      return jsonResponse({ status: 'error', message: 'טוקן לא חוקי או פג תוקף', tokenExpired: true });
    }
    return jsonResponse({ status: 'ok', data: _bohanSiteData() });
  }

  // Path B: password login (first visit / after token expiry)
  if (p.password) {
    if (String(p.password) !== String(password)) {
      return jsonResponse({ status: 'error', message: 'סיסמה שגויה' });
    }
    return jsonResponse({
      status: 'ok',
      token: _bohanSiteIssueToken(secret),
      data: _bohanSiteData()
    });
  }

  return jsonResponse({ status: 'error', message: 'Missing password or token' });
}

// ========== Result-upload HMAC token (for Cloudflare Worker auth) ==========
// Issues a short-lived signed token that the browser sends in X-Auth-Token
// when POSTing exam-result HTML to the exam-results Cloudflare Worker.
// The Worker verifies the same HMAC with its own copy of the secret.
//
// Setup: in Apps Script, set ScriptProperty 'RESULT_UPLOAD_SECRET' to a long
// random string. Set the IDENTICAL value as a Worker secret binding named
// UPLOAD_SECRET. The secret never reaches the browser.
function handleGetResultUploadToken(p) {
  // Examiner auth (token+id) is already enforced by the doGet dispatcher
  // before this handler runs — see examinerActions list.
  var props = PropertiesService.getScriptProperties();
  var secret = props.getProperty('RESULT_UPLOAD_SECRET');
  if (!secret) {
    return jsonResponse({
      status: 'error',
      message: 'RESULT_UPLOAD_SECRET not configured in Apps Script properties',
      code: 'not_configured'
    });
  }
  var payload = JSON.stringify({ exp: Date.now() + 5 * 60 * 1000 });
  var payloadB64 = Utilities.base64EncodeWebSafe(payload).replace(/=+$/, '');
  var sigBytes = Utilities.computeHmacSha256Signature(payloadB64, secret);
  var sigB64 = Utilities.base64EncodeWebSafe(sigBytes).replace(/=+$/, '');
  return jsonResponse({ status: 'ok', token: payloadB64 + '.' + sigB64 });
}

function unmarkPendingCompleted(sessionCode, idNumber) {
  // Restore the latest pending row from 'completed' back to 'in_exam' so a resumed
  // exam (after a premature close-fail was reverted) can be re-submitted. BUG FIX:
  // this previously read col index 6 (=language) and tested 'done' (a status that
  // is NEVER written — markPendingCompleted writes 'completed'), then wrote to
  // column 7 (=language) and scanned oldest-first — so it silently no-op'd and
  // could corrupt the language cell. Now mirrors markPendingCompleted exactly:
  // status = col index 5 / column 6, newest row first.
  var pendingSheet = getSheet('ממתינים');
  if (!pendingSheet) return;
  var data = pendingSheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === String(sessionCode) && normalizeId(data[i][1]) === normalizeId(idNumber)) {
      if (String(data[i][5] || '').trim() === 'completed') {
        pendingSheet.getRange(i + 1, 6).setValue('in_exam');
      }
      break;
    }
  }
}

// ========== שיתוף תוצאה דרך CacheService ==========

function handleUploadResultHtml(data) {
  // DISABLED: result-HTML hosting moved to the authenticated Cloudflare Worker
  // (RESULTS_URL). This Apps Script endpoint had NO auth and served attacker HTML
  // from the trusted Google origin (stored-XSS/phishing). The live client never
  // calls it (it posts to the Worker), so returning early is safe and closes the hole.
  return jsonResponse({ status: 'error', message: 'disabled — use the results worker' });
  // eslint-disable-next-line
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

  // Read practice results too — we'll join real-exam outcomes against the
  // practice history of the same name+license to surface a "did practice
  // before exam predict success?" metric. The student app stores its own
  // "מזהה תלמיד" (not the national ID), so we match only on full name +
  // license. Note that this is best-effort: identical names will collapse.
  var practiceSheet = getSheet('תוצאות תרגול');
  var practiceData = practiceSheet.getDataRange().getValues();

  // Class → site map (from כיתות) — practice rows store the class code, not the
  // site, so this lets the name+site fallback match scope by base.
  var pClassSiteMap = {};
  try {
    var pClassData = getSheet('כיתות').getDataRange().getValues();
    for (var pcs = 1; pcs < pClassData.length; pcs++) {
      pClassSiteMap[String(pClassData[pcs][0]).trim()] = String(pClassData[pcs][7] || '').trim();
    }
  } catch (ePCS) { /* no כיתות sheet → name+site index stays empty, name fallback still works */ }
  // Phone normaliser: digits only, last 9 (so "050-1234567", "0501234567" and
  // "972501234567" all collapse to the same key on both practice and exam sides).
  function normPhoneCmd(v) {
    var d = String(v || '').replace(/\D/g, '');
    return d.length >= 9 ? d.slice(-9) : '';
  }

  // Parse a stay-time string from the 'זמן' column into seconds. The examinee
  // client writes Hebrew-formatted strings like "32 דק' 14 שנ'" (see
  // getElapsedTimeStr in examinee.html). Older rows or other clients may use
  // "MM:SS"; we handle both shapes so the dashboard works across the whole
  // historical dataset.
  //
  // Anything we can't parse returns 0 and the row is dropped from the stay-
  // time aggregate — better to under-count than to poison the median with
  // garbage values (negative durations, hh:mm:ss strings that look like
  // minutes when truncated, etc.).
  function parseStayTimeToSeconds(str) {
    if (!str) return 0;
    var s = String(str).trim();
    if (!s) return 0;
    // Format A — Hebrew: "32 דק' 14 שנ'", optionally with geresh ׳ or U+2019.
    // The regex tolerates extra whitespace and either quote-mark variant.
    var hebMatch = s.match(/(\d+)\s*דק['׳’]?\s*(\d+)\s*שנ['׳’]?/);
    if (hebMatch) {
      var hmins = parseInt(hebMatch[1], 10);
      var hsecs = parseInt(hebMatch[2], 10);
      if (!isNaN(hmins) && !isNaN(hsecs) && hsecs < 60) {
        var ht = hmins * 60 + hsecs;
        if (ht > 0 && ht <= 5400) return ht;
      }
      return 0;
    }
    // Format B — "MM:SS" colon-separated (legacy and other clients).
    var colonMatch = s.match(/^(\d{1,3}):(\d{2})$/);
    if (colonMatch) {
      var cmins = parseInt(colonMatch[1], 10);
      var csecs = parseInt(colonMatch[2], 10);
      if (isNaN(cmins) || isNaN(csecs) || csecs >= 60) return 0;
      var ctotal = cmins * 60 + csecs;
      if (ctotal > 5400) return 0;
      return ctotal;
    }
    return 0;
  }

  // Bucket thresholds for stay-time histograms (seconds).
  //   fast    < 20 min — efficient
  //   normal  20–35 min — typical exam pace
  //   slow    > 35 min — close to the 40-min ceiling
  var STAY_FAST_MAX = 20 * 60;
  var STAY_NORMAL_MAX = 35 * 60;

  // Aggregate. stayTimes tracks duration-in-seconds per result so we can
  // compute avg / median / p10 / p90 / 3-bucket histogram per group.
  var overall = { total: 0, passed: 0, failed: 0, disqualified: 0, stayTimes: [], reattempts: 0 };
  // Previous period of identical length, ending right before dateFrom —
  // powers the ▲▼ trend badges on the KPI cards (this period vs the last one).
  var prevOverall = { total: 0, passed: 0, failed: 0, disqualified: 0, reattempts: 0 };
  var prevWindowMs = dateTo.getTime() - dateFrom.getTime();
  var prevFrom = new Date(dateFrom.getTime() - prevWindowMs - 1);
  // Integrity flags (current window only). Definitions mirror the per-row
  // badges in the examiner results table, so commander totals always match
  // what the examiner sees row-by-row.
  var integrityOverall = { unverified: 0, suspicious: 0, corrected: 0 };
  var integrityByExaminer = {};
  var integrityBySite = {};
  var byExaminer = {};
  var bySite = {};
  var byLicense = {};
  var byPopulation = {};
  var byLanguage = {};       // 'he' / 'ru' / ... → stats. Catches translation issues
                              // (one language failing more than others is a content/HEB-RTL signal).
  var byAttempt = {};        // 'ניסיון 1' / 'ניסיון 2' / 'ניסיון 3+' → stats.
                              // Shows whether re-attempts have higher/lower
                              // pass rate (does the second try go better?).
  var byDevice = {};         // 'טלפון'/'טאבלט'/'מחשב'/'מבחן בכתב'/'לא צוין (ישן)' → stats.
  var byAudio = {};          // '🔊 שמע' / 'רגיל' → stats.
  // Weak-topic aggregation: per graded exam, the license blueprint tells how
  // many questions of each topic were asked; each parsed wrong-block is later
  // resolved (id→category via the question DB) and counted against that.
  var topicAsked = {};
  var topicAskedByLic = {};
  var weakTopicPending = [];
  var byDay = {};            // 'YYYY-MM-DD' → count. Drives the throughput line chart.
  var byHour = {};           // 'dow-hour' (0–6 dow, 0–23 hour) → count. Heatmap data.
  var wrongQuestionCounts = {}; // question text → fail count. Aggregated from
                                 // column 15 (פירוט שגויות) to surface the
                                 // top-N most-missed questions for content review.

  // Language label normalizer — short codes get human names for the UI.
  var LANG_LABELS_SERVER = {
    'he': 'עברית', 'ru': 'רוסית', 'en': 'אנגלית',
    'ar': 'ערבית', 'fr': 'צרפתית', 'es': 'ספרדית', 'am': 'אמהרית'
  };
  function isoDateStr(d) {
    var y = d.getFullYear();
    var m = d.getMonth() + 1;
    var day = d.getDate();
    return y + '-' + (m < 10 ? '0' + m : m) + '-' + (day < 10 ? '0' + day : day);
  }

  // ===== Practice impact index =====
  // Builds a lookup: normalized-name + '|' + license → list of practice rows
  // (sorted by date desc). For each real exam row, we'll find the latest
  // practice within 30 days BEFORE the exam date and bucket the impact.
  function normalizeFullName(s) {
    if (!s) return '';
    var t = String(s).trim();
    if (!t) return '';
    // Strip common punctuation and collapse internal whitespace
    t = t.replace(/[׳״'".\-]/g, ' ').replace(/\s+/g, ' ').trim().toLowerCase();
    // Token-sort so "ויטלי גיטלמן" and "גיטלמן ויטלי" hash to the same key
    var tokens = t.split(' ').filter(function(x) { return x; });
    tokens.sort();
    return tokens.join(' ');
  }
  function parsePercentValue(v) {
    if (typeof v === 'number') return v <= 1 ? v * 100 : v;
    var s = String(v || '').replace('%', '').trim();
    if (!s) return -1;
    var n = parseFloat(s);
    if (isNaN(n)) return -1;
    return n <= 1 ? n * 100 : n;
  }
  var practiceIndex = {};         // name|license → [{date, percent}]
  var practicePhoneIndex = {};    // last-9-digits phone → [...]  (exact join, fwd-only)
  var practiceNameSiteIndex = {}; // name|license|site → [...]   (collision-safe fallback)
  for (var pi = 1; pi < practiceData.length; pi++) {
    var pName = normalizeFullName(practiceData[pi][2]);
    if (!pName) continue;
    var pLic = String(practiceData[pi][5] || '').trim();
    var pDate = practiceData[pi][0];
    if (pDate && !(pDate instanceof Date)) pDate = new Date(pDate);
    if (!pDate || isNaN(pDate.getTime())) continue;
    var pPct = parsePercentValue(practiceData[pi][8]);
    var pRec = { date: pDate, percent: pPct };
    var pKey = pName + '|' + pLic;
    if (!practiceIndex[pKey]) practiceIndex[pKey] = [];
    practiceIndex[pKey].push(pRec);
    var pPhone = (practiceData[pi].length > 15) ? normPhoneCmd(practiceData[pi][15]) : '';
    if (pPhone) {
      if (!practicePhoneIndex[pPhone]) practicePhoneIndex[pPhone] = [];
      practicePhoneIndex[pPhone].push(pRec);
    }
    var pSite = pClassSiteMap[String(practiceData[pi][3] || '').trim()] || '';
    if (pSite) {
      var pnsKey = pName + '|' + pLic + '|' + pSite;
      if (!practiceNameSiteIndex[pnsKey]) practiceNameSiteIndex[pnsKey] = [];
      practiceNameSiteIndex[pnsKey].push(pRec);
    }
  }
  // Sort each list newest-first for the "latest before X" linear scan.
  function sortPracticeLists(idx) { for (var k in idx) idx[k].sort(function(a, b) { return b.date - a.date; }); }
  sortPracticeLists(practiceIndex);
  sortPracticeLists(practicePhoneIndex);
  sortPracticeLists(practiceNameSiteIndex);

  // Practice-impact accumulators. 4 buckets — 'none' (no practice in window),
  // 'low' (<70%), 'mid' (70-85.99%), 'high' (≥86%, the pass threshold).
  var practiceImpact = {
    none: { total: 0, passed: 0 },
    low:  { total: 0, passed: 0 },
    mid:  { total: 0, passed: 0 },
    high: { total: 0, passed: 0 },
    withAny: { total: 0, passed: 0 }, // sum of low+mid+high — pre-computed for the simple card
    unparseable: 0 // practice found but % couldn't be parsed; counted under withAny but not bucketed
  };
  // Match-coverage: of the eligible (non-DQ) exam-takers, how many were matched
  // to a practice record and by which key. Lets the UI show honest coverage
  // instead of silently treating unmatched as "didn't practice".
  var piCoverage = { eligible: 0, matched: 0, byPhone: 0, byNameSite: 0, byName: 0 };

  // Examiners who registered as examinees to test the system — exclude from stats (by name OR ת.ז.).
  var examinerExcl = getExaminerExclusion();

  for (var r = 1; r < resData.length; r++) {
    var rowDate = parseSheetDate(resData[r][0]);
    if (!rowDate) continue;
    var inPrevWindow = rowDate >= prevFrom && rowDate < dateFrom;
    if ((rowDate < dateFrom || rowDate > dateTo) && !inPrevWindow) continue;

    var examinerName = String(resData[r][9] || '');
    var siteName = String(resData[r][10] || '');
    if (isTestSite(siteName)) continue;   // system-test site — exclude from ALL commander stats (current + previous window)
    if (isExaminerSelfTest(resData[r][2], resData[r][1], examinerExcl)) continue;   // examiner self-testing (name or ת.ז.) — exclude
    var license = String(resData[r][4] || '');
    var population = String(resData[r][19] || '');
    var passedStr = String(resData[r][7] || '');
    if (passedStr === 'בוטל') continue;
    var isDQ = resData[r][17] === true || String(resData[r][17]).toUpperCase() === 'TRUE' || passedStr === 'פסול';
    var isPassed = !isDQ && (passedStr === 'עבר');

    // Previous-window rows feed ONLY the trend comparison — none of the
    // breakdowns, charts or integrity tallies below.
    if (inPrevWindow) {
      prevOverall.total++;
      if (isDQ) prevOverall.disqualified++;
      else if (isPassed) prevOverall.passed++;
      else prevOverall.failed++;
      if ((Number(resData[r][14]) || 1) > 1) prevOverall.reattempts++;
      continue;
    }

    // Integrity flags — same definitions as the examiner results-table badges:
    // unverified = score not re-verified against the trusted answer key
    // (anything except 'מאומת'/'ידני'), excluding DQ rows and 0/X system-fails;
    // suspicious = exam finished in under 3 minutes; corrected = manually
    // amended result (col תוקן?).
    var integVState = (resData[r].length > 22) ? String(resData[r][22] || '') : '';
    var integSuspicious = (resData[r].length > 23) && String(resData[r][23] || '') === 'חשוד';
    var integCorrected = resData[r][20] === true || String(resData[r][20]).toUpperCase() === 'TRUE';
    var integZeroScore = /^0\//.test(String(resData[r][5] || ''));
    var integUnverified = !isDQ && !integZeroScore && integVState !== 'מאומת' && integVState !== 'ידני';
    if (integUnverified || integSuspicious || integCorrected) {
      var integEx = examinerName || 'לא צוין';
      var integSite = siteName || 'לא צוין';
      if (!integrityByExaminer[integEx]) integrityByExaminer[integEx] = { unverified: 0, suspicious: 0, corrected: 0 };
      if (!integrityBySite[integSite]) integrityBySite[integSite] = { unverified: 0, suspicious: 0, corrected: 0 };
      if (integUnverified) { integrityOverall.unverified++; integrityByExaminer[integEx].unverified++; integrityBySite[integSite].unverified++; }
      if (integSuspicious) { integrityOverall.suspicious++; integrityByExaminer[integEx].suspicious++; integrityBySite[integSite].suspicious++; }
      if (integCorrected) { integrityOverall.corrected++; integrityByExaminer[integEx].corrected++; integrityBySite[integSite].corrected++; }
    }

    // Stay-time: approval → submit. We use the exam's elapsed-time column 8
    // ('זמן', MM:SS) as the proxy — examinee clicks "Start Exam" within a
    // few seconds of approval, and submit happens at the time we record.
    // Adding a separate approval-timestamp column would tighten this but
    // requires a schema change; current proxy is within ~30 seconds.
    var timeSec = parseStayTimeToSeconds(resData[r][8]);

    // Re-attempt detection: column 14 (ניסיון) holds the attempt number for
    // this exam (1, 2, 3...). Anything > 1 is the same examinee taking it
    // again after a previous fail/DQ — useful signal for tracking how many
    // failures actually come back vs walk away.
    var attemptNum = Number(resData[r][14]) || 1;
    var isReattempt = attemptNum > 1;

    // Practice impact — look up most recent practice for this examinee
    // (matched by name+license) within 30 days before the exam date.
    // DQ rows are excluded because they don't reflect knowledge level.
    if (!isDQ) {
      var examineeName = normalizeFullName(resData[r][2]);
      var realLic = String(resData[r][4] || '').trim();
      var realDate = resData[r][0];
      if (realDate && !(realDate instanceof Date)) realDate = new Date(realDate);
      if (examineeName && realDate && !isNaN(realDate.getTime())) {
        var thirtyBefore = new Date(realDate);
        thirtyBefore.setDate(thirtyBefore.getDate() - 30);
        piCoverage.eligible++;
        // Match priority: exact phone → name+license+site → name+license.
        var examPhone = normPhoneCmd(resData[r][3]);
        var examSite = String(resData[r][10] || '').trim();
        var lookupList = null, matchType = '';
        if (examPhone && practicePhoneIndex[examPhone]) {
          lookupList = practicePhoneIndex[examPhone]; matchType = 'byPhone';
        }
        if (!lookupList && examSite && practiceNameSiteIndex[examineeName + '|' + realLic + '|' + examSite]) {
          lookupList = practiceNameSiteIndex[examineeName + '|' + realLic + '|' + examSite]; matchType = 'byNameSite';
        }
        if (!lookupList) {
          lookupList = practiceIndex[examineeName + '|' + realLic] || null;
          if (lookupList) matchType = 'byName';
        }
        lookupList = lookupList || [];
        var matchedPractice = null;
        for (var ml = 0; ml < lookupList.length; ml++) {
          var item = lookupList[ml];
          if (item.date <= realDate && item.date >= thirtyBefore) {
            matchedPractice = item;
            break;
          }
        }
        if (!matchedPractice) {
          practiceImpact.none.total++;
          if (isPassed) practiceImpact.none.passed++;
        } else {
          piCoverage.matched++;
          if (matchType && typeof piCoverage[matchType] === 'number') piCoverage[matchType]++;
          practiceImpact.withAny.total++;
          if (isPassed) practiceImpact.withAny.passed++;
          var pct = matchedPractice.percent;
          if (pct < 0) {
            practiceImpact.unparseable++;
          } else if (pct < 70) {
            practiceImpact.low.total++;
            if (isPassed) practiceImpact.low.passed++;
          } else if (pct < 86) {
            practiceImpact.mid.total++;
            if (isPassed) practiceImpact.mid.passed++;
          } else {
            practiceImpact.high.total++;
            if (isPassed) practiceImpact.high.passed++;
          }
        }
      }
    }

    // Language (col 12) — drives the byLanguage breakdown. Default to Hebrew
    // since that's the source language and missing values pre-date the column.
    var langCode = String(resData[r][12] || 'he').toLowerCase().trim();
    var langName = LANG_LABELS_SERVER[langCode] || langCode;

    overall.total++;
    if (isDQ) overall.disqualified++;
    else if (isPassed) overall.passed++;
    else overall.failed++;
    if (timeSec > 0) overall.stayTimes.push(timeSec);
    if (isReattempt) overall.reattempts++;

    // Time-series — one increment per row, no sub-groups (keeps payload small)
    var dayKey = isoDateStr(rowDate);
    if (!byDay[dayKey]) byDay[dayKey] = { total: 0, passed: 0, failed: 0, dq: 0 };
    byDay[dayKey].total++;
    if (isDQ) byDay[dayKey].dq++;
    else if (isPassed) byDay[dayKey].passed++;
    else byDay[dayKey].failed++;
    var hourKey = rowDate.getDay() + '-' + rowDate.getHours();
    byHour[hourKey] = (byHour[hourKey] || 0) + 1;

    // Wrong-question aggregation. Column 15 (פירוט שגויות) is a multi-line
    // string with one block per missed question:
    //   "שאלה: <text>\nתשובת הנבחן: <ans>\nתשובה נכונה: <correct>\n\n"
    // We split on blank lines and extract the "שאלה:" line as the natural
    // key. Question text is a stable identifier across rows because the same
    // text is rendered for every examinee who got that question wrong.
    var wrongDetails = String(resData[r][15] || '');
    var topicBlocksParsed = 0;
    if (wrongDetails) {
      var blocks = wrongDetails.split(/\n\s*\n/);
      for (var wb = 0; wb < blocks.length; wb++) {
        var lines = blocks[wb].split('\n');
        var qText = '', qCorrect = '', qId = '';
        for (var wl = 0; wl < lines.length; wl++) {
          var line = lines[wl];
          if (line.indexOf('מזהה שאלה:') === 0) {
            qId = line.replace(/^מזהה שאלה:\s*/, '').trim();
          } else if (line.indexOf('שאלה:') === 0) {
            qText = line.replace(/^שאלה:\s*/, '').trim();
            if (qText.length > 200) qText = qText.substring(0, 200);
          } else if (line.indexOf('תשובה נכונה:') === 0) {
            qCorrect = line.replace(/^תשובה נכונה:\s*/, '').trim();
            // Historical data has many "undefined - undefined" entries from the
            // legacy per-language ci bug (memory: project_per_language_ci_bug).
            // Treat those as if no correct answer was captured.
            if (qCorrect.indexOf('undefined') !== -1 || qCorrect === '-' || qCorrect === '') {
              qCorrect = '';
            } else {
              var labelStripMatch = qCorrect.match(/^[A-Za-dא-לА-Г]\s*[-–]\s*(.+)$/);
              if (labelStripMatch) qCorrect = labelStripMatch[1].trim();
            }
          }
        }
        // Preferred aggregation key: question ID (added 2026-06-02). Uniquely
        // identifies the question across all license/language variants — no
        // false collisions, no "מה פירוש התמרור?" lumping.
        // Fallback for legacy rows without ID: (text + correctAnswer), or
        // text alone if correctAnswer is also missing/garbage.
        var key;
        if (qId) {
          key = 'id:' + qId;
        } else if (qText) {
          key = 't:' + qText + '|||' + qCorrect;
        } else {
          continue;
        }
        if (!wrongQuestionCounts[key]) {
          wrongQuestionCounts[key] = { count: 0, text: qText, correctAnswer: qCorrect, questionId: qId, langCounts: {} };
        }
        wrongQuestionCounts[key].count++;
        // Per-language split — a question failing mostly in one non-Hebrew
        // language is a translation-bug signal for the content team.
        wrongQuestionCounts[key].langCounts[langName] = (wrongQuestionCounts[key].langCounts[langName] || 0) + 1;
        // Weak-topic pending item: resolved to a topic after the loop via the
        // question DB (id preferred; text fallback for legacy rows).
        topicBlocksParsed++;
        weakTopicPending.push({ id: qId || '', text: qText || '', license: license, topic: '' });
      }
    }

    // Blueprint-based "asked" totals per topic: count a row's blueprint once
    // per graded digital exam (server-verified, or legacy rows that at least
    // carry parsed wrong-blocks). DQ rows, 0/X system-fails and manual paper
    // entries (no per-question data) stay out of both numerator & denominator.
    if (!isDQ && !integZeroScore && (integVState === 'מאומת' || topicBlocksParsed > 0)) {
      var topicBp = EXAM_STRUCTURE_SERVER[license];
      if (topicBp) {
        if (!topicAskedByLic[license]) topicAskedByLic[license] = {};
        for (var tbk in topicBp) {
          topicAsked[tbk] = (topicAsked[tbk] || 0) + topicBp[tbk];
          topicAskedByLic[license][tbk] = (topicAskedByLic[license][tbk] || 0) + topicBp[tbk];
        }
      }
    }

    var eName = examinerName || 'לא צוין';
    var sName = siteName || 'לא צוין';
    var lName = license || 'לא צוין';
    var pName = population || 'לא צוין';

    // Attempt-bucket label: 1, 2, 3+ (anything ≥ 3 collapses to a single
    // bucket — the long tail is too small to be useful on its own).
    var attemptLabel = attemptNum <= 1 ? 'ניסיון 1'
                       : attemptNum === 2 ? 'ניסיון 2'
                       : 'ניסיון 3+';

    addToGroup(byExaminer, eName, isPassed, isDQ, timeSec);
    addToGroup(bySite, sName, isPassed, isDQ, timeSec);
    addToGroup(byLicense, lName, isPassed, isDQ, timeSec);
    addToGroup(byPopulation, pName, isPassed, isDQ, timeSec);
    addToGroup(byLanguage, langName, isPassed, isDQ, timeSec);
    addToGroup(byAttempt, attemptLabel, isPassed, isDQ, timeSec);

    // Device + audio dimensions. Device (col 30) exists only on new rows —
    // older rows group under 'לא צוין (ישן)'; paper entries show 'מבחן בכתב'.
    var deviceRaw = (resData[r].length > 29) ? String(resData[r][29] || '').trim() : '';
    var deviceLabel = deviceRaw === 'phone' ? 'טלפון'
                      : deviceRaw === 'tablet' ? 'טאבלט'
                      : deviceRaw === 'desktop' ? 'מחשב'
                      : (integVState === 'ידני' ? 'מבחן בכתב' : 'לא צוין (ישן)');
    var audioLabel = String(resData[r][21] || 'off') === 'on' ? '🔊 שמע' : 'רגיל';
    addToGroup(byDevice, deviceLabel, isPassed, isDQ, timeSec);
    addToGroup(byAudio, audioLabel, isPassed, isDQ, timeSec);

    // Cross-tabulation sub-groups
    addToSubGroup(byExaminer, eName, 'byLicense', lName, isPassed, isDQ, timeSec);
    addToSubGroup(byExaminer, eName, 'bySite', sName, isPassed, isDQ, timeSec);
    addToSubGroup(bySite, sName, 'byLicense', lName, isPassed, isDQ, timeSec);
    addToSubGroup(bySite, sName, 'byExaminer', eName, isPassed, isDQ, timeSec);
    addToSubGroup(byLicense, lName, 'bySite', sName, isPassed, isDQ, timeSec);
    addToSubGroup(byLicense, lName, 'byExaminer', eName, isPassed, isDQ, timeSec);
    addToSubGroup(byPopulation, pName, 'byLicense', lName, isPassed, isDQ, timeSec);
    addToSubGroup(byPopulation, pName, 'bySite', sName, isPassed, isDQ, timeSec);
  }

  function addToGroup(map, key, isPassed, isDQ, timeSec) {
    if (!map[key]) map[key] = { total: 0, passed: 0, failed: 0, disqualified: 0, stayTimes: [] };
    map[key].total++;
    if (isDQ) map[key].disqualified++;
    else if (isPassed) map[key].passed++;
    else map[key].failed++;
    if (timeSec > 0) map[key].stayTimes.push(timeSec);
  }

  function addToSubGroup(map, primaryKey, subDim, subKey, isPassed, isDQ, timeSec) {
    if (!map[primaryKey]) return;
    if (!map[primaryKey]._sub) map[primaryKey]._sub = {};
    if (!map[primaryKey]._sub[subDim]) map[primaryKey]._sub[subDim] = {};
    addToGroup(map[primaryKey]._sub[subDim], subKey, isPassed, isDQ, timeSec);
  }

  // Percentile helper. arr is assumed already sorted ascending.
  function percentileSorted(sortedArr, p) {
    if (!sortedArr || sortedArr.length === 0) return 0;
    if (sortedArr.length === 1) return sortedArr[0];
    var rank = (p / 100) * (sortedArr.length - 1);
    var lo = Math.floor(rank);
    var hi = Math.ceil(rank);
    if (lo === hi) return sortedArr[lo];
    var w = rank - lo;
    return Math.round(sortedArr[lo] * (1 - w) + sortedArr[hi] * w);
  }

  function computeStats(obj) {
    var stayAvg = 0, stayMedian = 0, stayP10 = 0, stayP90 = 0;
    var stayFast = 0, stayNormal = 0, staySlow = 0;
    var stayTimes = obj.stayTimes || [];
    if (stayTimes.length > 0) {
      var sum = 0;
      for (var s = 0; s < stayTimes.length; s++) {
        sum += stayTimes[s];
        if (stayTimes[s] < STAY_FAST_MAX) stayFast++;
        else if (stayTimes[s] <= STAY_NORMAL_MAX) stayNormal++;
        else staySlow++;
      }
      stayAvg = Math.round(sum / stayTimes.length);
      var sorted = stayTimes.slice().sort(function(a, b) { return a - b; });
      stayMedian = percentileSorted(sorted, 50);
      stayP10 = percentileSorted(sorted, 10);
      stayP90 = percentileSorted(sorted, 90);
    }
    var passRate = obj.total > 0 ? Math.round((obj.passed / obj.total) * 100) : 0;
    var dqRate = obj.total > 0 ? Math.round((obj.disqualified / obj.total) * 100) : 0;
    return {
      total: obj.total,
      passed: obj.passed,
      failed: obj.failed,
      disqualified: obj.disqualified,
      passRate: passRate,
      dqRate: dqRate,
      // Stay-time metrics (all in seconds; client formats as MM:SS)
      stayAvg: stayAvg,
      stayMedian: stayMedian,
      stayP10: stayP10,
      stayP90: stayP90,
      stayFast: stayFast,
      stayNormal: stayNormal,
      staySlow: staySlow,
      stayCount: stayTimes.length
    };
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

  // Throughput timeline: fill in zero-count days between dateFrom and dateTo
  // so the client gets a contiguous series instead of a sparse one (cleaner
  // chart, no false-impression gaps).
  var timeline = [];
  var cursor = new Date(dateFrom);
  cursor.setHours(0, 0, 0, 0);
  var endDay = new Date(dateTo);
  endDay.setHours(0, 0, 0, 0);
  var safetyLimit = 0;
  while (cursor <= endDay && safetyLimit < 400) {
    var k = isoDateStr(cursor);
    var dayAgg = byDay[k] || { total: 0, passed: 0, failed: 0, dq: 0 };
    // `count` kept so an older client (plain-count polyline) keeps working.
    timeline.push({ date: k, count: dayAgg.total, passed: dayAgg.passed, failed: dayAgg.failed, dq: dayAgg.dq });
    cursor.setDate(cursor.getDate() + 1);
    safetyLimit++;
  }

  // Heatmap: flatten to a 7×24 array of counts (0 = Sunday in JS date.getDay())
  var heatmap = [];
  for (var dow = 0; dow < 7; dow++) {
    var hourRow = [];
    for (var hr = 0; hr < 24; hr++) hourRow.push(byHour[dow + '-' + hr] || 0);
    heatmap.push(hourRow);
  }

  // Re-attempt summary — overall.reattempts already counted in the loop;
  // turn it into a rate so the client can show both raw count and %.
  var overallStats = computeStats(overall);
  overallStats.reattempts = overall.reattempts;
  overallStats.reattemptRate = overall.total > 0
    ? Math.round((overall.reattempts / overall.total) * 100)
    : 0;

  // Previous-period rates for the KPI trend badges.
  prevOverall.passRate = prevOverall.total > 0 ? Math.round((prevOverall.passed / prevOverall.total) * 100) : 0;
  prevOverall.dqRate = prevOverall.total > 0 ? Math.round((prevOverall.disqualified / prevOverall.total) * 100) : 0;
  prevOverall.reattemptRate = prevOverall.total > 0 ? Math.round((prevOverall.reattempts / prevOverall.total) * 100) : 0;

  // ===== Weak topics: resolve pending wrong-blocks to topics =====
  // id is language-independent (all language files carry the Hebrew category),
  // so most items resolve on the first (Hebrew) pass; text-fallback items from
  // legacy rows resolve when their exam language comes up. Cache makes the
  // repeated loads cheap (~300ms warm per language).
  var topicWrong = {};
  var topicWrongByLic = {};
  try {
    var WT_LANGS = ['he', 'ru', 'en', 'ar', 'fr', 'es', 'am'];
    for (var wtl = 0; wtl < WT_LANGS.length; wtl++) {
      var wtUnresolved = false;
      for (var wtc = 0; wtc < weakTopicPending.length; wtc++) {
        if (!weakTopicPending[wtc].topic) { wtUnresolved = true; break; }
      }
      if (!wtUnresolved) break;
      var wtQs;
      try { wtQs = loadQuestionsForLanguageServer(WT_LANGS[wtl]); } catch (eWtLoad) { continue; }
      if (!Array.isArray(wtQs) || wtQs.length === 0) continue;
      var wtById = {}, wtByText = {};
      for (var wtq = 0; wtq < wtQs.length; wtq++) {
        var wtRec = wtQs[wtq];
        if (!wtRec) continue;
        if (wtRec.id) wtById[String(wtRec.id)] = wtRec;
        if (wtRec.text) wtByText[String(wtRec.text).substring(0, 200)] = wtRec;
      }
      for (var wtp = 0; wtp < weakTopicPending.length; wtp++) {
        var wtItem = weakTopicPending[wtp];
        if (wtItem.topic) continue;
        var wtFound = wtItem.id ? wtById[wtItem.id] : null;
        if (!wtFound && wtItem.text) wtFound = wtByText[wtItem.text];
        if (wtFound) {
          var wtTopic = classifyCategoryServer(wtFound.category);
          if (wtTopic) wtItem.topic = wtTopic;
        }
      }
    }
  } catch (eWtAll) { /* weak-topic section degrades to empty, dashboard still works */ }
  for (var wtf = 0; wtf < weakTopicPending.length; wtf++) {
    var wtFin = weakTopicPending[wtf];
    if (!wtFin.topic) continue;
    topicWrong[wtFin.topic] = (topicWrong[wtFin.topic] || 0) + 1;
    if (!topicWrongByLic[wtFin.license]) topicWrongByLic[wtFin.license] = {};
    topicWrongByLic[wtFin.license][wtFin.topic] = (topicWrongByLic[wtFin.license][wtFin.topic] || 0) + 1;
  }
  function buildTopicArr(wrongMap, askedMap) {
    var tArr = [];
    for (var tk in askedMap) {
      var tAsked = askedMap[tk] || 0;
      if (!tAsked) continue;
      var tWrong = wrongMap[tk] || 0;
      tArr.push({ topic: tk, wrong: tWrong, asked: tAsked, pct: Math.round((tWrong / tAsked) * 100) });
    }
    tArr.sort(function(a, b) { return b.pct - a.pct; });
    return tArr;
  }
  var weakTopicsOut = { overall: buildTopicArr(topicWrong, topicAsked), byLicense: {} };
  for (var wtLic in topicAskedByLic) {
    weakTopicsOut.byLicense[wtLic] = buildTopicArr(topicWrongByLic[wtLic] || {}, topicAskedByLic[wtLic]);
  }

  // ===== Wait time: registration → exam start (ממתינים col E + col L) =====
  // Approval time isn't stored, so this measures the full soldier experience:
  // registered → waited for approval → pressed Start. Per-site via the row's
  // own site (new rows) or the session's host site (fallback).
  var waitTimesOut = { overall: { avg: 0, median: 0, p90: 0, count: 0 }, bySite: {} };
  try {
    var pendSheetW = getSheet('ממתינים');
    var pendDataW = pendSheetW.getDataRange().getValues();
    var sessSheetW = getSheet('סשנים');
    var sessDataW = sessSheetW.getDataRange().getValues();
    var sessSiteMapW = {};
    for (var swi = 1; swi < sessDataW.length; swi++) {
      sessSiteMapW[String(sessDataW[swi][0]).trim()] = String(sessDataW[swi][3] || '');
    }
    var waitAll = [];
    var waitBySiteArr = {};
    for (var pwi = 1; pwi < pendDataW.length; pwi++) {
      var regRaw = pendDataW[pwi][4];
      var startRaw = pendDataW[pwi][11];
      if (!regRaw || !startRaw) continue;
      var regD = regRaw instanceof Date ? regRaw : new Date(regRaw);
      var startD = startRaw instanceof Date ? startRaw : new Date(startRaw);
      if (isNaN(regD.getTime()) || isNaN(startD.getTime())) continue;
      if (regD < dateFrom || regD > dateTo) continue;
      var waitSec = Math.round((startD.getTime() - regD.getTime()) / 1000);
      if (waitSec <= 0 || waitSec > 4 * 3600) continue; // clock skew / stuck rows
      waitAll.push(waitSec);
      var waitSite = ((pendDataW[pwi].length > 17 ? String(pendDataW[pwi][17] || '') : '').trim())
        || sessSiteMapW[String(pendDataW[pwi][0]).trim()] || 'לא צוין';
      if (!waitBySiteArr[waitSite]) waitBySiteArr[waitSite] = [];
      waitBySiteArr[waitSite].push(waitSec);
    }
    function waitStatsOf(arr) {
      if (!arr.length) return { avg: 0, median: 0, p90: 0, count: 0 };
      var wSum = 0;
      for (var wsi = 0; wsi < arr.length; wsi++) wSum += arr[wsi];
      var wSorted = arr.slice().sort(function(a, b) { return a - b; });
      return {
        avg: Math.round(wSum / arr.length),
        median: percentileSorted(wSorted, 50),
        p90: percentileSorted(wSorted, 90),
        count: arr.length
      };
    }
    waitTimesOut.overall = waitStatsOf(waitAll);
    for (var wbs in waitBySiteArr) waitTimesOut.bySite[wbs] = waitStatsOf(waitBySiteArr[wbs]);
  } catch (eWait) { /* wait-time card simply stays hidden */ }

  // Top-N most-missed questions, sorted by count descending. Capped at 10 —
  // beyond that the list gets noisy and stops driving decisions.
  // Values are objects {count, text, correctAnswer, questionId}; questionId
  // takes precedence (post-2026-06-02 data), text+correctAnswer is the
  // fallback for legacy rows.
  var topWrong = [];
  var wrongKeys = Object.keys(wrongQuestionCounts);
  wrongKeys.sort(function(a, b) {
    return wrongQuestionCounts[b].count - wrongQuestionCounts[a].count;
  });
  for (var wk = 0; wk < Math.min(wrongKeys.length, 10); wk++) {
    var entry = wrongQuestionCounts[wrongKeys[wk]];
    topWrong.push({
      question: entry.text || '',
      correctAnswer: entry.correctAnswer || '',
      questionId: entry.questionId || '',
      count: entry.count,
      langCounts: entry.langCounts || {}
    });
  }

  // Image lookup: many top-N questions are traffic-sign prompts ("מה פירוש
  // התמרור?") that don't make sense without seeing the sign. Try to resolve
  // each top-wrong entry to its real question record so we can include the
  // imageUrl + id. Hebrew first (most exams); other languages as fallback
  // for entries that didn't resolve.
  //
  // The match key is (text, correctAnswer-in-answers, correctIndex points to
  // that answer). This is strict enough that even if two real questions share
  // the same text, we only attach the image when the correct-answer text also
  // matches the answer-key index — so we either get the right sign or no
  // image. Better silent miss than a wrong picture.
  try {
    if (typeof loadQuestionsForLanguageServer === 'function') {
      var SUPPORTED_LANGS_FOR_IMG = ['he', 'ru', 'en', 'ar', 'fr', 'es', 'am'];
      for (var lgi = 0; lgi < SUPPORTED_LANGS_FOR_IMG.length; lgi++) {
        var stillMissing = false;
        for (var tw0 = 0; tw0 < topWrong.length; tw0++) {
          if (!topWrong[tw0].imageUrl) { stillMissing = true; break; }
        }
        if (!stillMissing) break;
        var langQs;
        try { langQs = loadQuestionsForLanguageServer(SUPPORTED_LANGS_FOR_IMG[lgi]); }
        catch (eLoad) { continue; }
        if (!Array.isArray(langQs) || langQs.length === 0) continue;
        // Index by ID (fast path) and by text (fallback path)
        var qById = {}, qByText = {};
        for (var qIdx = 0; qIdx < langQs.length; qIdx++) {
          var qRec = langQs[qIdx];
          if (!qRec) continue;
          if (qRec.id) qById[String(qRec.id)] = qRec;
          if (qRec.text) {
            if (!qByText[qRec.text]) qByText[qRec.text] = [];
            qByText[qRec.text].push(qRec);
          }
        }
        for (var twi = 0; twi < topWrong.length; twi++) {
          if (topWrong[twi].imageUrl) continue;
          // Path 0: question ID known (new data) — direct lookup, no
          // ambiguity. Best of all paths.
          if (topWrong[twi].questionId) {
            var idMatch = qById[String(topWrong[twi].questionId)];
            if (idMatch && idMatch.imageUrl) {
              topWrong[twi].imageUrl = idMatch.imageUrl;
              continue;
            }
          }
          var candidates = qByText[topWrong[twi].question] || [];
          if (candidates.length === 0) continue;
          // Path A: correct answer known — match it precisely against the
          // candidate's answers array.
          if (topWrong[twi].correctAnswer) {
            for (var ci = 0; ci < candidates.length; ci++) {
              var cand = candidates[ci];
              if (!Array.isArray(cand.answers)) continue;
              if (cand.answers.indexOf(topWrong[twi].correctAnswer) !== -1) {
                if (cand.imageUrl) topWrong[twi].imageUrl = cand.imageUrl;
                if (cand.id && !topWrong[twi].questionId) topWrong[twi].questionId = cand.id;
                break;
              }
            }
          }
          // Path B: only one candidate for this exact text — unambiguous.
          if (!topWrong[twi].imageUrl && candidates.length === 1) {
            if (candidates[0].imageUrl) topWrong[twi].imageUrl = candidates[0].imageUrl;
            if (candidates[0].id && !topWrong[twi].questionId) topWrong[twi].questionId = candidates[0].id;
          }
        }
      }
    }
  } catch (eImg) { /* image resolution best-effort; ignore failures */ }

  // Practice impact — finalize pass rates for each bucket. Pass rate is
  // computed only on the non-DQ sample (DQs were excluded above).
  function finalizePI(bucket) {
    return {
      total: bucket.total,
      passed: bucket.passed,
      passRate: bucket.total > 0 ? Math.round((bucket.passed / bucket.total) * 100) : 0
    };
  }
  var practiceImpactOut = {
    none:        finalizePI(practiceImpact.none),
    withAny:     finalizePI(practiceImpact.withAny),
    low:         finalizePI(practiceImpact.low),
    mid:         finalizePI(practiceImpact.mid),
    high:        finalizePI(practiceImpact.high),
    unparseable: practiceImpact.unparseable,
    lookbackDays: 30,
    coverage: piCoverage
  };

  var result = {
    overall: overallStats,
    byExaminer: computeGroupWithSub(byExaminer),
    bySite: computeGroupWithSub(bySite),
    byLicense: computeGroupWithSub(byLicense),
    byPopulation: computeGroupWithSub(byPopulation),
    byLanguage: computeGroupWithSub(byLanguage),
    byAttempt: computeGroupWithSub(byAttempt),
    timeline: timeline,
    heatmap: heatmap,
    topWrong: topWrong,
    practiceImpact: practiceImpactOut,
    prevOverall: prevOverall,
    integrity: { overall: integrityOverall, byExaminer: integrityByExaminer, bySite: integrityBySite },
    byDevice: computeGroupWithSub(byDevice),
    byAudio: computeGroupWithSub(byAudio),
    weakTopics: weakTopicsOut,
    waitTimes: waitTimesOut
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
  if (role !== 'מפקד' && role !== 'מפקד מקומי' && role !== 'מפקד ראשי' && !isKdtzRole(role)) {
    return jsonResponse({ status: 'error', message: 'אין הרשאת מפקד' });
  }

  // Determine commander scope:
  //   isGlobal    — sees every site (no site filter)
  //   isLocal     — single site (column 9 holds the one site name)
  //   isMultiSite — fixed list of sites (column 9 holds a comma-separated list)
  // Legacy: role === 'מפקד' treated as 'מפקד ראשי'
  var isGlobal = (role === 'מפקד ראשי' || role === 'מפקד');
  var isLocal = (role === 'מפקד מקומי');
  var isMultiSite = isKdtzRole(role);
  var managedSites = [];
  if (isMultiSite) {
    managedSites = String(userSite || '').split(',').map(function(s) { return s.trim(); }).filter(function(s) { return s; });
    if (!managedSites.length) {
      return jsonResponse({ status: 'error', message: 'לא הוקצו אתרים — מלא רשימה מופרדת בפסיקים בעמודת האתר במורים' });
    }
  }

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

  var overall = { total: 0, passed: 0, failed: 0, scores: [], stayTimes: [], students: {}, teachers: {}, classes: {}, sites: {} };
  var byTeacher = {};
  var byClass = {};
  var byLicense = {};
  var byMode = {};
  var bySite = {};

  // Activity-by-hour (0–23) + most-failed-questions, mirroring the examiner
  // commander view. Practice rows store the submit time inside the date cell
  // ("DD/MM/YYYY HH:mm" — see todayStr) and the missed questions as a JSON
  // array in col 13 (פירוט שגויות), one {qNum, category, qText} per question.
  var hourBuckets = [];
  for (var hb = 0; hb < 24; hb++) hourBuckets.push(0);
  var wrongCounts = {}; // qText → { count, category }

  for (var r = 1; r < resData.length; r++) {
    var rowDate = parseSheetDate(resData[r][0]);
    if (!rowDate || rowDate < dateFrom || rowDate > dateTo) continue;

    var classCode = String(resData[r][3] || '').trim();
    if (!classCode) continue;

    var cInfo = classMap[classCode] || { teacherName: 'לא ידוע', className: classCode, license: '', site: '' };
    var classSite = cInfo.site || '';

    // Site filtering for local + multi-site commanders
    if (isLocal && userSite && classSite !== userSite) continue;
    if (isMultiSite && managedSites.indexOf(classSite) === -1) continue;

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

    // Stay-time (col 10, "M:SS"/"MM:SS") → seconds, for the avg/median time KPIs.
    var tSec = parsePracticeTimeSec(resData[r][10]);
    if (tSec > 0) overall.stayTimes.push(tSec);

    // Activity-by-hour from the submit timestamp embedded in the date cell.
    var hh = practiceRowHour(resData[r][0], rowDate);
    if (hh >= 0 && hh < 24) hourBuckets[hh]++;

    // Most-failed questions — the practice client sends a JSON array of
    // {qNum, category, qText}. Aggregate by question text (no ID/correct-answer
    // is captured in practice mode, unlike the real-exam path).
    var wdRaw = resData[r][13];
    if (wdRaw) {
      var wdArr = null;
      try { wdArr = (typeof wdRaw === 'string') ? JSON.parse(wdRaw) : wdRaw; } catch (eWD) { wdArr = null; }
      if (Array.isArray(wdArr)) {
        for (var wdi = 0; wdi < wdArr.length; wdi++) {
          var wit = wdArr[wdi];
          if (!wit) continue;
          var qt = String(wit.qText || wit.question || '').trim();
          if (!qt) continue;
          if (qt.length > 200) qt = qt.substring(0, 200);
          if (!wrongCounts[qt]) wrongCounts[qt] = { count: 0, category: String(wit.category || '') };
          wrongCounts[qt].count++;
        }
      }
    }

    addToGroup(byTeacher, teacherName, isPassed, isFailed, pctVal, studentId);
    addToGroup(byClass, className + ' (' + classCode + ')', isPassed, isFailed, pctVal, studentId);
    addToGroup(byLicense, license, isPassed, isFailed, pctVal, studentId);
    addToGroup(byMode, mode, isPassed, isFailed, pctVal, studentId);

    // bySite aggregation — global and multi-site commanders both want
    // a per-site breakdown. Local commander has only one site, so the
    // tab is hidden client-side; no aggregation needed.
    if ((isGlobal || isMultiSite) && classSite) {
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

  // Parse a practice duration into seconds. The student app sends "MM:SS"
  // (e.g. "5:03" = 5 min 3 sec), but Google Sheets AUTO-CONVERTS the string on
  // write, MISREADING "MM:SS" as "HH:MM". So getValues() never returns the
  // original string — it returns one of:
  //   • Date   — short sessions (<24 min). e.g. "5:03" → 05:03 time → Date.
  //   • number — long sessions (≥24 min). "24:29" can't be a time-of-day, so
  //              Sheets stores it as a DURATION serial (fraction of a day, e.g.
  //              ~1.02 for 24h29m). Verified against real data: 28,637 Date
  //              cells + 1,871 duration-number cells.
  //   • string — only if a value somehow wasn't auto-converted.
  // In every case the stored clock is H:M:S where the original minutes landed in
  // H and the original seconds in M. We recover by mapping H→minutes, M→seconds.
  // Cap at 2h to drop garbage (abandoned tabs produce multi-day serials).
  function parsePracticeTimeSec(v) {
    if (v === null || v === undefined || v === '') return 0;
    var mm, ss;
    if (v instanceof Date) {
      mm = v.getHours();      // original minutes (Sheets read them as hours)
      ss = v.getMinutes();    // original seconds (Sheets read them as minutes)
    } else if (typeof v === 'number') {
      // Day-fraction serial (works for both <1 time serials and ≥1 durations).
      if (v <= 0) return 0;
      var totalClockSec = Math.round(v * 86400); // the misread H:M:S, in seconds
      mm = Math.floor(totalClockSec / 3600);     // clock-hours → original minutes
      ss = Math.floor((totalClockSec % 3600) / 60); // clock-minutes → original seconds
    } else {
      var m = String(v).trim().match(/^(\d{1,3}):(\d{2})$/);
      if (!m) return 0;
      mm = parseInt(m[1], 10);
      ss = parseInt(m[2], 10);
    }
    if (isNaN(mm) || isNaN(ss) || ss >= 60) return 0;
    var t = mm * 60 + ss;
    return (t > 0 && t <= 7200) ? t : 0;
  }
  // Hour-of-day from the date cell. Sheets usually auto-parses "DD/MM/YYYY
  // HH:mm" into a real Date (hour preserved); for string cells we regex the
  // HH out, since parseSheetDate drops the time component.
  function practiceRowHour(cell, parsed) {
    if (cell instanceof Date) return cell.getHours();
    var s = String(cell || '');
    var m = s.match(/\d{1,2}\/\d{1,2}\/\d{4}\s+(\d{1,2}):(\d{2})/);
    if (m) return parseInt(m[1], 10);
    if (parsed && parsed instanceof Date) return parsed.getHours();
    return -1;
  }
  function avgMedianSec(arr) {
    if (!arr || !arr.length) return { avg: 0, median: 0 };
    var sum = 0;
    for (var i = 0; i < arr.length; i++) sum += arr[i];
    var avg = Math.round(sum / arr.length);
    var sorted = arr.slice().sort(function(a, b) { return a - b; });
    var mid = Math.floor(sorted.length / 2);
    var median = sorted.length % 2 !== 0 ? sorted[mid] : Math.round((sorted[mid - 1] + sorted[mid]) / 2);
    return { avg: avg, median: median };
  }

  var overallStats = computeStats(overall);
  overallStats.activeTeachers = Object.keys(overall.teachers).length;
  overallStats.activeClasses = Object.keys(overall.classes).length;
  overallStats.activeSites = Object.keys(overall.sites).length;
  var ovTime = avgMedianSec(overall.stayTimes);
  overallStats.stayAvg = ovTime.avg;
  overallStats.stayMedian = ovTime.median;
  overallStats.stayCount = overall.stayTimes.length;

  var result = {
    overall: overallStats,
    byTeacher: computeGroupWithSub(byTeacher),
    byClass: computeGroupWithSub(byClass),
    byLicense: computeGroupWithSub(byLicense),
    byMode: computeGroupWithSub(byMode),
    commanderRole: role,
    commanderSite: userSite
  };

  // bySite breakdown for any cross-site commander (global or multi-site).
  if (isGlobal || isMultiSite) {
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
    if (isMultiSite && managedSites.indexOf(acSite) === -1) continue;
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

  // Top-10 most-failed questions (sorted by fail count) + activity-by-hour.
  var topWrong = [];
  var wKeys = Object.keys(wrongCounts);
  wKeys.sort(function(a, b) { return wrongCounts[b].count - wrongCounts[a].count; });
  for (var twk = 0; twk < Math.min(wKeys.length, 10); twk++) {
    topWrong.push({ question: wKeys[twk], category: wrongCounts[wKeys[twk]].category, count: wrongCounts[wKeys[twk]].count });
  }
  result.topWrong = topWrong;
  result.hourly = hourBuckets;

  // Repeat REAL-exam failures — soldiers who failed the external theory exam
  // (תוצאות, not practice) 2+ times in the window. The training commander owns
  // the intervention: targeted practice before they burn another exam slot.
  // Site scoping mirrors the practice rows (local/multi-site commanders see
  // only their sites; exam rows carry the site in col 11).
  try {
    var examResSheet = getSheet('תוצאות');
    var examResData = examResSheet.getDataRange().getValues();
    var failsById = {};
    for (var er = 1; er < examResData.length; er++) {
      var erDate = parseSheetDate(examResData[er][0]);
      if (!erDate || erDate < dateFrom || erDate > dateTo) continue;
      // Only genuine knowledge fails: skip עבר, פסול (anti-cheat, not knowledge)
      // and בוטל (superseded/overturned rows).
      if (String(examResData[er][7] || '').trim() !== 'נכשל') continue;
      var erSite = String(examResData[er][10] || '');
      if (isLocal && userSite && erSite !== userSite) continue;
      if (isMultiSite && managedSites.indexOf(erSite) === -1) continue;
      var erId = normalizeId(examResData[er][1]);
      if (!erId) continue;
      if (!failsById[erId]) {
        failsById[erId] = { name: '', idLast4: String(examResData[er][1] || '').slice(-4), license: '', site: '', fails: 0, lastDate: '', lastScore: '' };
      }
      failsById[erId].fails++;
      // Rows are appended chronologically — the last in-range row wins the
      // "latest" fields.
      if (examResData[er][2]) failsById[erId].name = String(examResData[er][2]);
      if (examResData[er][4]) failsById[erId].license = String(examResData[er][4]);
      if (erSite) failsById[erId].site = erSite;
      failsById[erId].lastDate = erDate.getDate() + '/' + (erDate.getMonth() + 1) + '/' + erDate.getFullYear();
      failsById[erId].lastScore = String(examResData[er][5] || '');
    }
    var repeatFailures = [];
    for (var rfk in failsById) {
      if (failsById[rfk].fails >= 2) repeatFailures.push(failsById[rfk]);
    }
    repeatFailures.sort(function(a, b) { return b.fails - a.fails; });
    result.repeatFailures = repeatFailures.slice(0, 50);
  } catch (eRF) { result.repeatFailures = []; }

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
  var teacherSite = '';
  for (var i = 1; i < tData.length; i++) {
    if (normalizeId(tData[i][1]) === normalizeId(p.teacherId)) {
      teacherName = tData[i][0];
      teacherSite = String(tData[i][9] || '');
      break;
    }
  }
  sheet.appendRow([code, className, normalizeId(p.teacherId), teacherName, license, nowISO(), 'כן', teacherSite]);
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

function handleTeacherDeleteClass(p) {
  var classCode = String(p.classCode || '').trim();
  if (!classCode) return jsonResponse({ status: 'error', message: 'חסר קוד כיתה' });

  var classSheet = getSheet('כיתות');
  var classData = classSheet.getDataRange().getValues();
  var classRowIdx = -1;
  for (var i = 1; i < classData.length; i++) {
    if (String(classData[i][0]).trim() === classCode &&
        normalizeId(classData[i][2]) === normalizeId(p.teacherId)) {
      classRowIdx = i;
      break;
    }
  }
  if (classRowIdx === -1) return jsonResponse({ status: 'error', message: 'כיתה לא נמצאה או שאין הרשאה' });

  // Safety: only allow deletion of CLOSED classes (active = 'לא')
  if (String(classData[classRowIdx][6]).trim() === 'כן') {
    return jsonResponse({ status: 'error', message: 'יש לסגור את הכיתה לפני מחיקה' });
  }

  // Delete the class row
  classSheet.deleteRow(classRowIdx + 1);

  // Delete all students enrolled in this class (cleanup roster)
  var studentsRemoved = 0;
  var studSheet = getSheet('תלמידי כיתות');
  var studData = studSheet.getDataRange().getValues();
  for (var s = studData.length - 1; s >= 1; s--) {
    if (String(studData[s][0]).trim() === classCode) {
      studSheet.deleteRow(s + 1);
      studentsRemoved++;
    }
  }

  // NOTE: practice results in 'תוצאות תרגול' are intentionally preserved for historical reporting.

  return jsonResponse({ status: 'ok', studentsRemoved: studentsRemoved });
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

  // Check if already enrolled by studentId (same device/browser)
  var studSheet = getSheet('תלמידי כיתות');
  var studData = studSheet.getDataRange().getValues();
  for (var s = 1; s < studData.length; s++) {
    if (String(studData[s][0]).trim() === classCode && String(studData[s][2]).trim() === studentId) {
      return jsonResponse({ status: 'ok', message: 'כבר רשום בכיתה', className: classInfo.name, teacherName: classInfo.teacherName, license: classInfo.license });
    }
  }

  // Check if same name is already in this class with a DIFFERENT studentId (joined from another device/browser).
  // If so, return the existing studentId so the new device adopts it — prevents duplicate roster entries.
  var normName = studentName.toLowerCase().replace(/\s+/g, ' ');
  for (var s2 = 1; s2 < studData.length; s2++) {
    if (String(studData[s2][0]).trim() === classCode) {
      var existingName = String(studData[s2][1]).trim().toLowerCase().replace(/\s+/g, ' ');
      if (existingName === normName) {
        return jsonResponse({
          status: 'ok',
          message: 'מצאנו שאתה כבר רשום בכיתה הזו ממכשיר אחר. הנתונים שלך אוחדו.',
          existingStudentId: String(studData[s2][2]).trim(),
          className: classInfo.name,
          teacherName: classInfo.teacherName,
          license: classInfo.license
        });
      }
    }
  }

  studSheet.appendRow([classCode, studentName, studentId, nowISO()]);
  return jsonResponse({ status: 'ok', message: 'הצטרפת לכיתה בהצלחה!', className: classInfo.name, teacherName: classInfo.teacherName, license: classInfo.license });
}

function handleSubmitPracticeResult(p) {
  var studentId = String(p.studentId || '').trim();
  var classCode = String(p.classCode || '').trim();
  // Rate limit: cap public practice-result writes so the תוצאות תרגול sheet (which
  // feeds the teacher/commander stats) can't be flooded with fabricated rows.
  var prRlErr = requireRateLimit('submitPracticeResult', (studentId || classCode || 'anon'), 30, 60);
  if (prRlErr) return prRlErr;
  var sheet = getSheet('תוצאות תרגול');
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

  sheet.appendRow([todayStr(), studentId, String(p.studentName || ''), classCode, mode, license, score, total, percent, passed, time, category, language, wrongDetails, categoryBreakdown, String(p.phone || '')]);
  return jsonResponse({ status: 'ok' });
}

function handleLoadStudentProgress(p) {
  var name = String(p.studentName || '').trim();
  var classCode = String(p.classCode || '').trim().toUpperCase();
  if (!name || !classCode) {
    return jsonResponse({ status: 'error', message: 'חסרים פרטים' });
  }
  var key = name.toLowerCase() + '|' + classCode;
  var sheet = getSheet('התקדמות תלמידים');
  var row = findRow(sheet, 2, key);
  if (row === -1) {
    return jsonResponse({ status: 'ok', found: false });
  }
  var data = sheet.getRange(row, 1, 1, 7).getValues()[0];
  return jsonResponse({
    status: 'ok',
    found: true,
    streak: data[3] || '{}',
    wrongQs: data[4] || '[]',
    history: data[5] || '[]',
    lastUpdated: data[6] || ''
  });
}

function handleSaveStudentProgress(p) {
  var name = String(p.studentName || '').trim();
  var classCode = String(p.classCode || '').trim().toUpperCase();
  if (!name || !classCode) {
    return jsonResponse({ status: 'error', message: 'חסרים פרטים' });
  }
  var key = name.toLowerCase() + '|' + classCode;
  var streak = String(p.streak || '{}');
  var wrongQs = String(p.wrongQs || '[]');
  var history = String(p.history || '[]');
  var sheet = getSheet('התקדמות תלמידים');
  var row = findRow(sheet, 2, key);
  if (row === -1) {
    sheet.appendRow([name, classCode, key, streak, wrongQs, history, nowISO()]);
  } else {
    sheet.getRange(row, 1, 1, 7).setValues([[name, classCode, key, streak, wrongQs, history, nowISO()]]);
  }
  return jsonResponse({ status: 'ok' });
}
