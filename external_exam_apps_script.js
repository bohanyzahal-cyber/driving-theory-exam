// ============================================================================
// GENERATED FILE — do not edit here. Source: server/src/*.js (order: server/BUILD_ORDER.json).
// Rebuild with:  node tools/build_server.js     (tools/build.js runs it too)
// Deploy: paste this whole file into the Apps Script editor → Deploy → Manage deployments → New version.
// Modules: 00_config, 05_registry, 10_spreadsheet, 08_pending_writes, 98_migration_practice, 12_reads, 14_pending_archive, 16_lookup_util, 18_whatsapp, 20_auth, 22_util, 70_questions, 30_api, 40_sessions_login, 82_report_center, 42_sessions_manage, 84_report_site, 44_sessions_misc, 50_pending, 55_dashboard, 52_pending_status, 65_dq, 60_exam, 75_diag, 86_commander, 88_predictive, 90_teacher, 92_at_risk, 94_forecast, 96_admin, 91_teacher_classes
// ============================================================================
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
  // 15 columns: createSession appends the default population (idx 14) and
  // getSessionInfo / listSessions read it (same drift as ממתינים, review E §5).
  'סשנים': ['קוד', 'בוחן ת.ז.', 'שם בוחן', 'אתר', 'כיתה', 'דרגה', 'שפה', 'מצב שמע', 'זמן יצירה', 'תקף עד', 'פעיל', 'כמויות JSON', 'מאושרים JSON', 'בוחן אחראי', 'אוכלוסיית ברירת מחדל'],
  // 19 columns: the code reads 15-18 and writes 19 (review E §5 — a sheet
  // re-created from a 15-column header would be born four columns short and the
  // warning counter / site / "finished on device" writes would land outside it).
  'ממתינים': ['קוד סשן', 'ת.ז.', 'שם', 'טלפון', 'זמן הרשמה', 'סטטוס', 'שפה', 'אוכלוסיה', 'דרגה', 'שמע', 'הארכת זמן', 'התחלת מבחן', 'טוקן נבחן', 'ספירת DQ', 'מסך נוסף', 'ספירת אזהרות', 'אזהרה אחרונה', 'אתר', 'סיים במכשיר'],
  // Rows moved out of ממתינים by archiveSheets (full 19-col width, nothing deleted).
  'ממתינים_ארכיון': ['קוד סשן', 'ת.ז.', 'שם', 'טלפון', 'זמן הרשמה', 'סטטוס', 'שפה', 'אוכלוסיה', 'דרגה', 'שמע', 'הארכת זמן', 'התחלת מבחן', 'טוקן נבחן', 'ספירת DQ', 'מסך נוסף', 'ספירת אזהרות', 'אזהרה אחרונה', 'אתר', 'סיים במכשיר'],
  'תוצאות': ['תאריך', 'ת.ז.', 'שם', 'טלפון', 'דרגה', 'ציון', 'אחוז', 'עבר/נכשל', 'זמן', 'בוחן', 'אתר', 'כיתה', 'שפה', 'קוד סשן', 'ניסיון', 'פירוט שגויות', 'נשלח?', 'פסול?', 'קישור וואטסאפ', 'אוכלוסיה', 'תוקן?', 'שמע', 'מאומת', 'חשוד', 'dqEventId', 'תוקן ע"י', 'סיבת תיקון', 'תאריך תיקון', 'מסלול שפות', 'מכשיר'],
  // Rows moved out of תוצאות by archiveSheets (30 days) — same 30 columns.
  'תוצאות_ארכיון': ['תאריך', 'ת.ז.', 'שם', 'טלפון', 'דרגה', 'ציון', 'אחוז', 'עבר/נכשל', 'זמן', 'בוחן', 'אתר', 'כיתה', 'שפה', 'קוד סשן', 'ניסיון', 'פירוט שגויות', 'נשלח?', 'פסול?', 'קישור וואטסאפ', 'אוכלוסיה', 'תוקן?', 'שמע', 'מאומת', 'חשוד', 'dqEventId', 'תוקן ע"י', 'סיבת תיקון', 'תאריך תיקון', 'מסלול שפות', 'מכשיר'],
  // The question map of one attempt. Written only by the exam-start path and
  // read only while that attempt is scored, so it is archived after 2 days.
  'מבחנים': ['קוד סשן', 'ת.ז.', 'שאלות JSON', 'זמן רישום', 'שפה', 'שגויות לא מאומתות'],
  'מבחנים_ארכיון': ['קוד סשן', 'ת.ז.', 'שאלות JSON', 'זמן רישום', 'שפה', 'שגויות לא מאומתות'],
  'הארכות זמן': ['תאריך', 'קוד סשן', 'ת.ז.', 'שם', 'דקות', 'סיבה', 'בוחן'],
  // 10 columns: role (idx 8) and site (idx 9) are read by teacherLogin,
  // teacherVerifyLogin, teacherCommanderDashboard and adminDashboard.
  'מורים': ['שם', 'ת.ז.', 'סיסמה', 'פעיל', 'טוקן', 'תוקף טוקן', 'ניסיונות כושלים', 'נעילה עד', 'תפקיד', 'אתר'],
  // 8 columns: createClass appends the site (idx 7) and every commander report reads it.
  'כיתות': ['קוד כיתה', 'שם כיתה', 'מורה ת.ז.', 'שם מורה', 'דרגה', 'תאריך יצירה', 'פעיל', 'אתר'],
  'כיתות שנמחקו': ['קוד כיתה', 'שם כיתה', 'מורה ת.ז.', 'שם מורה', 'דרגה', 'אתר', 'תאריך מחיקה'],
  'תלמידי כיתות': ['קוד כיתה', 'שם תלמיד', 'מזהה תלמיד', 'תאריך הצטרפות'],
  'תוצאות תרגול': ['תאריך', 'מזהה תלמיד', 'שם תלמיד', 'קוד כיתה', 'מצב', 'דרגה', 'ציון', 'סה"כ', 'אחוז', 'עבר/נכשל', 'זמן', 'נושא', 'שפה', 'פירוט שגויות', 'פירוט לפי נושא', 'טלפון'],
  'התקדמות תלמידים': ['שם תלמיד', 'קוד כיתה', 'מפתח', 'streak', 'wrong_qs', 'history', 'עדכון אחרון'],
  'חיזוי סיכון': ['חושב בתאריך', 'שם', 'דרגה', 'קוד כיתה', 'מורה ת.ז.', 'שם מורה', 'שם כיתה', 'אתר', 'ציון תרגול', 'תרגולים', 'מגמה', 'ניסיון צפוי', 'ניגש בעבר', 'סיכוי מעבר', 'רמת סיכון', 'ביטחון', 'זוהה בטלפון', 'טלפון']
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
// lowercase, token-sort (so "ישראל ישראלי" and "ישראלי ישראל" hash the same).
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

// ========== API action registry ==========
// Every handler module declares its own actions with defineAction(); doGet/doPost
// dispatch through the registry instead of a 300-line switch, so adding an
// action is one line next to its handler and the auth rule lives with it.
//   defineAction('startExam', { methods: ['POST'], auth: 'examinee', handler: handleStartExam });
//   auth: 'none' | 'examiner' | 'teacher' | 'examinee' | 'gateway'
//   methods: any of 'GET', 'POST' (a GET to a POST-only action is refused).
//   rateLimit: optional { max, windowSec, id: function(p) -> identifier }
// The registry is kept on the function object so module load order does not
// matter (a module may register before this file's vars would have run).
function apiRegistry() {
  if (!apiRegistry._actions) apiRegistry._actions = {};
  return apiRegistry._actions;
}
function defineAction(name, spec) {
  if (!name || !spec || typeof spec.handler !== 'function') throw new Error('defineAction: bad spec for ' + name);
  apiRegistry()[name] = {
    name: name,
    methods: spec.methods || ['GET'],
    auth: spec.auth || 'none',
    handler: spec.handler,
    rateLimit: spec.rateLimit || null
  };
}
function apiActionNames() { return Object.keys(apiRegistry()).sort(); }
// One spreadsheet handle per execution, and the first open is marked: on 17/09
// an examinerDashboard spent 84.8 s before its first sheet mark, and the trail
// could not say whether opening the document or reading 'בוחנים' took it.
var _spreadsheetHandle = null;
function getSpreadsheet() {
  if (!_spreadsheetHandle) {
    _spreadsheetHandle = SpreadsheetApp.getActiveSpreadsheet();
    diagMark('ss:open');
  }
  return _spreadsheetHandle;
}

// ---- Practice/teacher data lives in its OWN spreadsheet (r24, review B1) ----
// Teachers' class details were full reads of a 107,900-row sheet on the very
// document the exam pollers read and write, and a practice evening alone
// reached the platform's rejection threshold (16/09 19:46, 98 executions alive).
// The document is the unit of contention, so the seven practice/teacher sheets
// move to a second spreadsheet whose id is kept in the Script Property
// PRACTICE_SPREADSHEET_ID. Until that property is set (migratePracticeSpreadsheet
// sets it after a verified copy) everything still comes from the active
// spreadsheet — this code is safe to deploy before migrating, and
// rollbackPracticeSpreadsheet() points everything back.
var PRACTICE_SPREADSHEET_PROPERTY = 'PRACTICE_SPREADSHEET_ID';
var PRACTICE_MIGRATING_PROPERTY = 'PRACTICE_MIGRATING';
var PRACTICE_SHEET_NAMES = ['תוצאות תרגול', 'כיתות', 'תלמידי כיתות', 'מורים', 'התקדמות תלמידים', 'כיתות שנמחקו', 'חיזוי סיכון'];
var PRACTICE_RETIRED_PREFIX = '_migrated_';
function isPracticeSheetName(name) { return PRACTICE_SHEET_NAMES.indexOf(String(name || '')) >= 0; }
var _practiceHandle = null, _practiceIdChecked = false, _practiceId = '';
function getPracticeSpreadsheetId() {
  if (!_practiceIdChecked) {
    _practiceIdChecked = true;
    try { _practiceId = String(PropertiesService.getScriptProperties().getProperty(PRACTICE_SPREADSHEET_PROPERTY) || '').trim(); }
    catch (e) { _practiceId = ''; }
  }
  return _practiceId;
}
function getPracticeSpreadsheet() {
  var id = getPracticeSpreadsheetId();
  if (!id) return getSpreadsheet();
  if (!_practiceHandle) {
    _practiceHandle = SpreadsheetApp.openById(id);
    diagMark('ss:open-practice');
  }
  return _practiceHandle;
}
function spreadsheetFor(name) {
  return isPracticeSheetName(name) ? getPracticeSpreadsheet() : getSpreadsheet();
}
// While migratePracticeSpreadsheet copies the sheets, a practice write would land
// in the old document and be lost; the practice/teacher write handlers answer a
// retryable maintenance error for the minutes the copy takes.
function practiceWriteGuard() {
  try {
    if (String(PropertiesService.getScriptProperties().getProperty(PRACTICE_MIGRATING_PROPERTY) || '') !== '1') return null;
  } catch (e) { return null; }
  return jsonResponse({ status: 'error', code: 'practice_maintenance', retryable: true, waitSec: 60,
    message: 'מערכת התרגול בתחזוקה של כמה דקות — נסה שוב בעוד דקה.' });
}

function getSheet(name) {
  var ss = spreadsheetFor(name);
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

function getSheetIfExists(name) {
  return spreadsheetFor(name).getSheetByName(name);
}

// ========== Writes to ממתינים shared by several modules ==========
// Every status change of an examinee row goes through setPendingStatus so that
// (a) the per-session snapshot the pollers read is dropped at once — a status
// written past the snapshot was the r23 gap (review C R8), and (b) the flush
// happens exactly once. Column numbers are 1-based getRange columns.
var PENDING_STATUS_COL = 6;
var PENDING_COLS = { status: 6, language: 7, population: 8, license: 9, audio: 10, timeExtension: 11, examStart: 12, token: 13, dqCount: 14, extScreen: 15, warnCount: 16, lastWarning: 17, site: 18, finishedOnDevice: 19 };
// extras: optional { <PENDING_COLS name>: value } written in the same flush.
function setPendingStatus(sheet, rowNumber, sessionCode, status, extras) {
  sheet.getRange(rowNumber, PENDING_STATUS_COL).setValue(status);
  if (extras) {
    for (var name in extras) {
      if (!Object.prototype.hasOwnProperty.call(extras, name) || !PENDING_COLS[name]) continue;
      sheet.getRange(rowNumber, PENDING_COLS[name]).setValue(extras[name]);
    }
  }
  SpreadsheetApp.flush();
  invalidatePendingSnapshot(sessionCode);
}

// Refresh current rows without repeatedly copying the whole growing sheet.
// Full snapshots retain old recovery rows; a changed row count/identity falls
// back to a full read so registration, retakes and maintenance remain visible.
function refreshExamineePendingRows(sheet, rows, sessionCode, idNumber) {
  if (!rows || sheet.getLastRow() !== rows.length) return sheet.getDataRange().getValues();
  var matchingRows = 0;
  for (var m = 1; m < rows.length; m++) {
    if (String(rows[m][0]) === String(sessionCode) && normalizeId(rows[m][1]) === normalizeId(idNumber)) matchingRows++;
  }
  if (matchingRows > 4) return sheet.getDataRange().getValues();
  for (var i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][0]) !== String(sessionCode) || normalizeId(rows[i][1]) !== normalizeId(idNumber)) continue;
    var live = sheet.getRange(i + 1, 1, 1, rows[i].length).getValues()[0];
    if (!live || String(live[0]) !== String(sessionCode) || normalizeId(live[1]) !== normalizeId(idNumber)) {
      return sheet.getDataRange().getValues();
    }
    rows[i] = live;
  }
  return rows;
}

// Helper: mark ALL active pending rows for this session+ID as completed.
// Closes EVERY in_exam/approved row (not just the latest) — a duplicate pending
// row otherwise leaves the soldier stuck on the board even though they finished
// and submitted (reported: "stuck in ממתינים/במבחן despite finishing").
function markPendingCompleted(sessionCode, idNumber, pendingSnapshot) {
  var pendSheet = pendingSnapshot ? pendingSnapshot.sheet : getSheet('ממתינים');
  var pendData = pendingSnapshot
    ? refreshExamineePendingRows(pendSheet, pendingSnapshot.rows, sessionCode, idNumber)
    : pendSheet.getDataRange().getValues();
  for (var j = pendData.length - 1; j >= 1; j--) {
    if (String(pendData[j][0]) === String(sessionCode) && normalizeId(pendData[j][1]) === normalizeId(idNumber) && (String(pendData[j][5]).trim() === 'in_exam' || String(pendData[j][5]).trim() === 'approved')) {
      pendSheet.getRange(j + 1, 6).setValue('completed');
      pendData[j][5] = 'completed';
    }
  }
}

// ---- Migration: run from the editor, in this order, see docs/OPERATIONS.md ----
// practiceMigrationPreflight  (read-only; also triggers the one-time permission prompt)
// migratePracticeSpreadsheet  (at a quiet hour; copies, verifies, cuts over; resumable)
// verifyPracticeMigration     (right after; every sheet must say ok)
// retirePracticeSheetsFromExamSpreadsheet (a day later: renames the old copies away)
// rollbackPracticeSpreadsheet (any time: points everything back to the exam spreadsheet)
function practiceMigrationPreflight() {
  var src = getSpreadsheet(), lines = [], id = getPracticeSpreadsheetId(), migrating = '';
  try { migrating = String(PropertiesService.getScriptProperties().getProperty(PRACTICE_MIGRATING_PROPERTY) || ''); } catch (e) {}
  lines.push('practice spreadsheet id: ' + (id || '(not set — the practice sheets still live in the exam spreadsheet)'));
  lines.push('migrating flag: ' + (migrating === '1' ? 'ON (practice writes are being refused)' : 'off'));
  for (var i = 0; i < PRACTICE_SHEET_NAMES.length; i++) {
    var s = src.getSheetByName(PRACTICE_SHEET_NAMES[i]);
    lines.push(PRACTICE_SHEET_NAMES[i] + ': ' + (s ? s.getLastRow() + ' rows × ' + s.getLastColumn() + ' cols' : 'missing in the exam spreadsheet'));
  }
  var scan = practiceCrossReferences(src);
  if (scan.skipped.length) lines.push('not scanned for formulas (script-written rows, too large): ' + scan.skipped.join(', '));
  lines.push(scan.refs.length ? 'CROSS-SHEET FORMULAS — handle these before migrating: ' + scan.refs.join('; ')
    : 'no formulas reference sheets across the practice/exam boundary');
  var out = lines.join('\n'); Logger.log(out); return out;
}
// Formulas in exam sheets that name a practice sheet (or the reverse) would break
// once the sheets live in different files. Sheets above 20,000 rows are not
// scanned (the practice results are script-written rows, never formulas).
function practiceCrossReferences(ss) {
  var refs = [], skipped = [], sheets = ss.getSheets(), examNames = [];
  for (var key in SHEET_HEADERS) { if (Object.prototype.hasOwnProperty.call(SHEET_HEADERS, key) && !isPracticeSheetName(key)) examNames.push(key); }
  examNames.push(DIAG_SHEET);
  for (var i = 0; i < sheets.length; i++) {
    var sh = sheets[i], name = sh.getName(), lookFor = isPracticeSheetName(name) ? examNames : PRACTICE_SHEET_NAMES;
    if (sh.getLastRow() < 1 || sh.getLastColumn() < 1) continue;
    if (sh.getLastRow() > 20000) { skipped.push(name + ' (' + sh.getLastRow() + ' rows)'); continue; }
    var formulas = sh.getDataRange().getFormulas();
    for (var r = 0; r < formulas.length && refs.length < 30; r++) {
      for (var c = 0; c < formulas[r].length; c++) {
        var f = formulas[r][c]; if (!f) continue;
        for (var k = 0; k < lookFor.length; k++) {
          if (f.indexOf(lookFor[k]) >= 0) { refs.push(name + '!R' + (r + 1) + 'C' + (c + 1) + ' → ' + lookFor[k]); break; }
        }
      }
    }
  }
  return { refs: refs, skipped: skipped };
}
function migratePracticeSpreadsheet() {
  var props = PropertiesService.getScriptProperties(), lines = [];
  if (getPracticeSpreadsheetId()) return 'already migrated to ' + getPracticeSpreadsheetId() + ' — run verifyPracticeMigration; rollbackPracticeSpreadsheet undoes it';
  var src = getSpreadsheet(), pendingKey = PRACTICE_SPREADSHEET_PROPERTY + '_PENDING';
  var pendingId = String(props.getProperty(pendingKey) || '').trim(), target = null;
  if (pendingId) {
    try { target = SpreadsheetApp.openById(pendingId); lines.push('resuming into ' + pendingId); } catch (eOpen) { target = null; }
  }
  if (!target) {
    target = SpreadsheetApp.create('תרגול ומורים — נתונים (מ-' + Utilities.formatDate(new Date(), 'Asia/Jerusalem', 'yyyy-MM-dd') + ')');
    props.setProperty(pendingKey, target.getId());
    lines.push('created ' + target.getUrl());
  }
  // Practice writes stay refused from the first run until the cut-over (the flag
  // is cleared below only on success, or by rollbackPracticeSpreadsheet).
  props.setProperty(PRACTICE_MIGRATING_PROPERTY, '1');
  var ok = true, complete = true, deadline = Date.now() + MIGRATION_BUDGET_MS;
  try {
    for (var i = 0; i < PRACTICE_SHEET_NAMES.length; i++) {
      var name = PRACTICE_SHEET_NAMES[i], s = src.getSheetByName(name);
      if (!s) { lines.push(name + ': not in the exam spreadsheet, skipped'); continue; }
      var existing = target.getSheetByName(name);
      if (existing && s.getLastRow() > 0 && existing.getLastRow() === s.getLastRow()) { lines.push(name + ': already copied (' + s.getLastRow() + ' rows)'); continue; }
      if (!copySheetInChunks(s, target, name, deadline, lines)) { complete = false; break; }
    }
    if (complete) {
      var sheets = target.getSheets();
      if (sheets.length > 1) {
        for (var j = 0; j < sheets.length; j++) {
          if (!isPracticeSheetName(sheets[j].getName()) && sheets[j].getLastRow() <= 1) { target.deleteSheet(sheets[j]); lines.push('removed the empty default sheet'); break; }
        }
      }
      for (var v = 0; v < PRACTICE_SHEET_NAMES.length; v++) {
        var sv = src.getSheetByName(PRACTICE_SHEET_NAMES[v]), dv = target.getSheetByName(PRACTICE_SHEET_NAMES[v]);
        if (sv && (!dv || dv.getLastRow() !== sv.getLastRow())) { ok = false; lines.push(PRACTICE_SHEET_NAMES[v] + ': row count differs after the copy'); }
      }
    }
  } catch (e) { ok = false; lines.push('ERROR: ' + (e && e.message ? e.message : e)); }
  if (ok && complete) {
    props.setProperty(PRACTICE_SPREADSHEET_PROPERTY, target.getId());
    props.deleteProperty(pendingKey);
    props.deleteProperty(PRACTICE_MIGRATING_PROPERTY);
    lines.push('CUT OVER: practice/teacher sheets are now served from ' + target.getUrl());
  } else if (!complete) {
    lines.push('NOT cut over yet — run migratePracticeSpreadsheet again to continue (practice writes stay refused until the cut-over)');
  } else {
    lines.push('NOT cut over — fix the errors above and run again (it resumes where it stopped; rollbackPracticeSpreadsheet abandons it)');
  }
  _practiceIdChecked = false; _practiceHandle = null;
  var out = lines.join('\n'); Logger.log(out); return out;
}
// Sheet.copyTo() failed on the 109,491-row practice sheet ("Unable to load
// document" on the exam spreadsheet, 19/09/2026 20:25) — the same document that
// stalls single reads on exam mornings cannot be materialised whole for a copy.
// So the rows are copied MIGRATION_CHUNK_ROWS at a time, values only (these are
// script-written data sheets: no formulas, formatting does not matter). The
// target's row count is the progress marker, so a run that stops on the time
// budget continues from where it left off; append-only sources that grew in
// between simply get their new rows copied too.
var MIGRATION_CHUNK_ROWS = 4000;
var MIGRATION_BUDGET_MS = 270000;   // 4.5 of the 6 minutes; the run stops cleanly and is re-run
function copySheetInChunks(src, target, name, deadline, lines) {
  var rowsSrc = src.getLastRow(), cols = src.getLastColumn();
  var dst = target.getSheetByName(name) || target.insertSheet(name);
  var done = dst.getLastRow();
  if (done > rowsSrc) { target.deleteSheet(dst); dst = target.insertSheet(name); done = 0; }
  if (rowsSrc === 0) { lines.push(name + ': empty, nothing to copy'); return true; }
  if (dst.getMaxRows() < rowsSrc) dst.insertRowsAfter(dst.getMaxRows(), rowsSrc - dst.getMaxRows());
  if (dst.getMaxColumns() < cols) dst.insertColumnsAfter(dst.getMaxColumns(), cols - dst.getMaxColumns());
  var t0 = Date.now(), startedAt = done;
  while (done < rowsSrc) {
    if (Date.now() > deadline) {
      SpreadsheetApp.flush();
      lines.push(name + ': ' + done + '/' + rowsSrc + ' rows so far (' + (done - startedAt) + ' this run, ' + (Date.now() - t0) + ' ms) — out of time');
      return false;
    }
    var n = Math.min(MIGRATION_CHUNK_ROWS, rowsSrc - done);
    var values = src.getRange(done + 1, 1, n, cols).getValues();
    dst.getRange(done + 1, 1, n, cols).setValues(values);
    done += n;
  }
  SpreadsheetApp.flush();
  lines.push(name + ': copied ' + done + '/' + rowsSrc + ' rows in chunks (' + (done - startedAt) + ' this run, ' + (Date.now() - t0) + ' ms)');
  return true;
}
function verifyPracticeMigration() {
  var id = getPracticeSpreadsheetId();
  if (!id) return 'not migrated: PRACTICE_SPREADSHEET_ID is not set';
  var src = getSpreadsheet(), dst = SpreadsheetApp.openById(id), lines = ['practice spreadsheet: ' + dst.getUrl()], bad = 0;
  for (var i = 0; i < PRACTICE_SHEET_NAMES.length; i++) {
    var name = PRACTICE_SHEET_NAMES[i], s = src.getSheetByName(name), d = dst.getSheetByName(name);
    if (!s) { lines.push(name + ': ' + (d ? 'only in the new spreadsheet (' + d.getLastRow() + ' rows)' : 'in neither')); continue; }
    if (!d) { bad++; lines.push(name + ': MISSING in the new spreadsheet'); continue; }
    var rs = s.getLastRow(), rd = d.getLastRow(), same = rs === rd;
    if (same && rs > 0) {
      var lastS = JSON.stringify(s.getRange(rs, 1, 1, s.getLastColumn()).getValues()[0]);
      var lastD = JSON.stringify(d.getRange(rd, 1, 1, d.getLastColumn()).getValues()[0]);
      same = lastS === lastD;
    }
    if (!same) bad++;
    lines.push(name + ': ' + (same ? 'ok' : 'MISMATCH') + ' (exam copy ' + rs + ' rows, new ' + rd + ' rows)');
  }
  lines.push(bad ? 'MISMATCHES: ' + bad + ' — expected only if practice traffic already wrote to the new spreadsheet' : 'ALL OK');
  var out = lines.join('\n'); Logger.log(out); return out;
}
function retirePracticeSheetsFromExamSpreadsheet() {
  var id = getPracticeSpreadsheetId();
  if (!id) return 'not migrated — nothing to retire';
  var src = getSpreadsheet(), dst = SpreadsheetApp.openById(id), lines = [], renamed = 0;
  for (var i = 0; i < PRACTICE_SHEET_NAMES.length; i++) {
    var name = PRACTICE_SHEET_NAMES[i], s = src.getSheetByName(name);
    if (!s) continue;
    if (!dst.getSheetByName(name)) { lines.push(name + ': kept — the new spreadsheet has no such sheet'); continue; }
    s.setName(PRACTICE_RETIRED_PREFIX + name); renamed++;
    lines.push(name + ': renamed to ' + PRACTICE_RETIRED_PREFIX + name);
  }
  lines.push(renamed + ' sheet(s) renamed; delete them by hand once a week of practice has run cleanly — that is when the exam spreadsheet actually shrinks');
  var out = lines.join('\n'); Logger.log(out); return out;
}
function rollbackPracticeSpreadsheet() {
  var props = PropertiesService.getScriptProperties(), src = getSpreadsheet(), lines = [];
  var id = getPracticeSpreadsheetId();
  props.deleteProperty(PRACTICE_SPREADSHEET_PROPERTY);
  props.deleteProperty(PRACTICE_MIGRATING_PROPERTY);
  props.deleteProperty(PRACTICE_SPREADSHEET_PROPERTY + '_PENDING');   // a new migration starts a fresh spreadsheet
  for (var i = 0; i < PRACTICE_SHEET_NAMES.length; i++) {
    var retired = src.getSheetByName(PRACTICE_RETIRED_PREFIX + PRACTICE_SHEET_NAMES[i]);
    if (retired && !src.getSheetByName(PRACTICE_SHEET_NAMES[i])) { retired.setName(PRACTICE_SHEET_NAMES[i]); lines.push(PRACTICE_SHEET_NAMES[i] + ': restored from ' + PRACTICE_RETIRED_PREFIX); }
  }
  _practiceIdChecked = false; _practiceHandle = null;
  lines.push('rolled back: practice sheets are served from the exam spreadsheet again' + (id ? ' (rows written to ' + id + ' since the cut-over are NOT copied back)' : ''));
  var out = lines.join('\n'); Logger.log(out); return out;
}

// ========== Tail reads for append-only sheets (perf) ==========
// ממתינים / תוצאות are append-only and every live-path handler filters by the
// CURRENT sessionCode, whose rows always sit at the bottom. Reading the whole
// sheet (4,000+ rows × 19-30 cols) on every poll was the steady-state slowness:
// checkApproval (20-30/min per examinee) + examinerDashboard (every 5s per
// examiner, 3 full reads each) alone kept most of the ~30 execution slots busy
// with no cold cache and no stampede involved.
// readTail reads only the last TAIL_ROWS data rows. Safety guard: if the OLDEST
// row in the tail is younger than TAIL_MAX_AGE_HOURS, rows of a live session
// could still exist above the tail → fall back to a full read (correct, slower).
var TAIL_ROWS = 1000;
var TAIL_MAX_AGE_HOURS = 48;

// NOTE: distinct from the existing parseSheetDate() (~line 4900), which drops the
// HH:MM part of todayStr() dates. The tail guard and the archive cutoff need the
// time of day, so this one keeps it. Do not merge the two.
function parseSheetDateTime(v) {
  if (!v) return null;
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
  var s = String(v).trim();
  // todayStr() format DD/MM/YYYY HH:MM — V8's Date parser would read it as MM/DD.
  var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2}))?/);
  if (m) return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]), Number(m[4] || 0), Number(m[5] || 0));
  var d = new Date(s);   // ISO (nowISO) and anything else Date understands
  return isNaN(d.getTime()) ? null : d;
}

// Returns { rows, off }. rows[0] = header; rows[i] (i >= 1) is sheet row (i + 1 + off).
// off === 0 on a full read, so the existing `getRange(i + 1, ...)` write pattern
// stays valid as long as callers add `off`.
function readTail(sheet, tsColIdx) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  // The mode is marked so the 'אבחון' trail says which read the caller paid
  // for: tail:small/<rows> (sheet fits in one tail), tail:<TAIL_ROWS>/<rows>, or
  // tail:full/<rows> when the 48-hour guard fired — a burst period can push a
  // live session above the last 1000 rows and turn every poll into a full read.
  if (lastRow - 1 <= TAIL_ROWS || lastCol < 1) {
    var small = sheet.getDataRange().getValues();
    diagMark('tail:small/' + lastRow);
    return { rows: small, off: 0 };
  }
  var startRow = lastRow - TAIL_ROWS + 1;
  var tail = sheet.getRange(startRow, 1, TAIL_ROWS, lastCol).getValues();
  var oldest = parseSheetDateTime(tail[0][tsColIdx]);
  if (!oldest || (Date.now() - oldest.getTime()) < TAIL_MAX_AGE_HOURS * 3600 * 1000) {
    // Tail might not cover a live session (burst day / unparseable timestamp).
    var full = sheet.getDataRange().getValues();
    diagMark('tail:full/' + lastRow);
    return { rows: full, off: 0 };
  }
  var header = sheet.getRange(1, 1, 1, lastCol).getValues();
  diagMark('tail:' + TAIL_ROWS + '/' + lastRow);
  return { rows: header.concat(tail), off: startRow - 2 };
}

function readPendingTail() { return readTail(getSheet('ממתינים'), 4); }   // col E = זמן הרשמה
function readResultsTail() { return readTail(getSheet('תוצאות'), 0); }    // col A = תאריך

// ---- Per-session snapshot of 'ממתינים' for the two read-only pollers (r23) ----
// checkApproval and getExamStatus are the most frequent requests of an exam
// morning (every waiting phone every 5-20 s, every phone in the exam every
// 10-20 s), and each paid a tail read of 'ממתינים' — three Sheets round trips.
// On 17/09 the stalls were single Sheets calls hanging ~93 s, about one call in
// a hundred; every round trip removed is a lottery ticket not bought. A
// session's rows are cached for PENDING_SNAPSHOT_SEC (a status can reach the
// phone that much later — under one poll interval). The writes an examinee
// waits for (register, approve, reject, reset) drop the snapshot, and a row
// missing from a cached snapshot is re-read from the sheet before "not found"
// is answered: the examinee page treats "not found" on reload as a dead
// registration and sends the soldier back to the code screen.
// The examiner dashboard keeps reading the sheet itself: it writes by row index.
var PENDING_SNAPSHOT_SEC = 4;
var PENDING_SNAPSHOT_PREFIX = 'pendsnap_';
function pendingSnapshotKey(sessionCode) {
  return CACHE_KEY_PREFIX + PENDING_SNAPSHOT_PREFIX + String(sessionCode || '').trim();
}
// Returns { rows, cached }. rows[0] is a header placeholder so the callers'
// `i >= 1` loops stay exactly as they were.
function pendingRowsForSession(sessionCode, fresh) {
  var key = pendingSnapshotKey(sessionCode), cache = null;
  if (!fresh) {
    try {
      cache = CacheService.getScriptCache();
      var hit = cache.get(key);
      if (hit) {
        var parsed = JSON.parse(hit);
        if (parsed && parsed.length) return { rows: [[]].concat(parsed), cached: true };
      }
    } catch (eGet) { cache = null; }
  }
  var all = readPendingTail().rows, mine = [], want = String(sessionCode || '').trim();
  for (var i = 1; i < all.length; i++) {
    if (String(all[i][0]).trim() === want) mine.push(all[i]);
  }
  try {
    if (!cache) cache = CacheService.getScriptCache();
    cache.put(key, JSON.stringify(mine), PENDING_SNAPSHOT_SEC);
  } catch (ePut) {}
  return { rows: [[]].concat(mine), cached: false };
}
function invalidatePendingSnapshot(sessionCode) {
  try { CacheService.getScriptCache().remove(pendingSnapshotKey(sessionCode)); } catch (e) {}
}

// r16: read only the rows a date-bounded caller can use.
// Every heavy aggregation skips rows older than some lower bound - the commander
// dashboard ignores results before prevFrom and practice rows more than 30 days
// before any such result - yet each read the whole sheet: 'אבחון' 15/09 showed
// practice-commander at 20-27s on every one of 12 dashboard opens. Take the tail
// first. The guard is the oldest row in hand: sheets here are append-only in
// time order, so only when that row is strictly older than `cutoff` can nothing
// relevant sit above it. Otherwise grow once (x4), then read everything. The
// result is correct whatever the sheet's size or a day's row count - the size
// only decides how often the cheap path wins. `mode` goes into the diag trail.
//
// r17 — the second half of the same problem: WIDTH. Bounding the rows bought
// almost nothing on 'תוצאות תרגול' (28.5s, mode `full`) while 'תוצאות' — the
// longer sheet — read in 1.3s, because a practice row carries two JSON blobs:
// col N 'פירוט שגויות' holds every wrong question's full text (~2KB/row) and
// col O 'פירוט לפי נושא' another. The commander handler reads NEITHER. So a
// caller may name the column ranges it needs and the skipped ones come back as
// '' — absolute indices are preserved, so `row[15]` is still the phone.
//
// `colSpec` is [[firstCol, numCols], ...], 1-based and ascending, and applies
// to the full read as well — that is where the 28 seconds actually were.
// ⚠ A caller passing colSpec MUST list every column it touches; a column left
// out reads as empty, it does not fail. Keep each call site's list next to the
// code that indexes the row.
function readSheetSlice(sheet, startRow, numRows, lastCol, colSpec) {
  if (numRows <= 0) return [];
  if (!colSpec || !colSpec.length) return sheet.getRange(startRow, 1, numRows, lastCol).getValues();
  var parts = [];
  for (var i = 0; i < colSpec.length; i++) {
    var first = colSpec[i][0], count = Math.min(colSpec[i][1], lastCol - first + 1);
    parts.push(count > 0 ? sheet.getRange(startRow, first, numRows, count).getValues() : null);
  }
  var out = [];
  for (var r = 0; r < numRows; r++) {
    var row = [];
    for (var j = 0; j < colSpec.length; j++) {
      while (row.length < colSpec[j][0] - 1) row.push('');   // the columns we deliberately skipped
      var piece = parts[j] ? parts[j][r] : null;
      if (piece) for (var k = 0; k < piece.length; k++) row.push(piece[k]);
    }
    out.push(row);
  }
  return out;
}

// r19: find the boundary EXACTLY, stop estimating it.
//
// Two guesses in a row missed. r16's fixed ladder (1000, 4000) was built without
// knowing the sheet's size — it turned out to be 107,614 rows. r18 replaced it
// with a row-rate estimate from the last 1000 rows, and that missed too: the
// 15/09 13:28 trail still reported `full/107626`, so whatever the real arrival
// pattern is, extrapolating a burst of 1000 rows does not describe it.
//
// So ask for the answer instead of modelling it. ONE read of the timestamp
// column alone (1 cell per row — for this sheet 107k cells against the 1.7M the
// full read moves) gives every row's date; the first row at or after the cutoff
// is then a fact. No assumption about row-rate, and none about ordering either:
// `first` is the TOPMOST qualifying row, so everything above it was examined and
// ruled out, whatever order the sheet is in.
//
// A row whose date will not parse counts as QUALIFYING. It may be junk the
// caller's own loop skips anyway, but the two loops here use two different
// parsers (parseSheetDate / parseSheetDateTime) — dropping a row this one
// cannot read risks losing a row the caller could. Never trade correctness for
// the read; if that forces a near-full read, `mode` will say so.
function firstRowSince(sheet, tsColIdx, cutoff, lastRow) {
  var col = sheet.getRange(2, tsColIdx + 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < col.length; i++) {
    var d = parseSheetDateTime(col[i][0]);
    if (!d || d.getTime() >= cutoff.getTime()) return i + 2;   // sheet row number
  }
  return 0;                                                    // nothing in range
}

function readRowsSince(sheet, tsColIdx, cutoff, colSpec) {
  var lastRow = sheet.getLastRow(), lastCol = sheet.getLastColumn(), dataRows = lastRow - 1;
  var bounded = cutoff instanceof Date && !isNaN(cutoff.getTime());
  if (bounded && lastCol >= 1 && dataRows > TAIL_ROWS) {
    var startRow = 0;
    try { startRow = firstRowSince(sheet, tsColIdx, cutoff, lastRow); } catch (eScan) { startRow = -1; }
    if (startRow === 0) {   // no row is recent enough: the caller wants the header and nothing else
      return { rows: readSheetSlice(sheet, 1, 1, lastCol, colSpec), off: 0, mode: 'none/' + lastRow };
    }
    if (startRow > 2) {
      var n = lastRow - startRow + 1;
      var tail = readSheetSlice(sheet, startRow, n, lastCol, colSpec);
      var header = readSheetSlice(sheet, 1, 1, lastCol, colSpec);
      return { rows: header.concat(tail), off: startRow - 2, mode: 'rows' + n + '/' + lastRow };
    }
  }
  // The row count rides along in `mode`: when a read still falls back to `full`
  // the trail must say how big the sheet actually is, instead of us guessing.
  if (colSpec && lastCol >= 1) return { rows: readSheetSlice(sheet, 1, lastRow, lastCol, colSpec), off: 0, mode: 'full/' + lastRow };
  return { rows: sheet.getDataRange().getValues(), off: 0, mode: 'full/' + lastRow };
}

// ---- History readers: live sheet + its archive, as one table ---------------
// The nightly job (B5) moves rows older than the retention window out of
// 'תוצאות' / 'ממתינים'. Every date-bounded aggregation must therefore read BOTH
// or it would silently report a shorter history than it did yesterday. Rows
// come back oldest-first (archive rows, then live rows) under ONE header, in
// the requested columns only.
//
// The archive is read only when it can still hold a row in range:
//   • readRowsSince answering 'rows…'/'none…' means it FOUND the boundary
//     inside the live sheet — everything above it, archive included, is older;
//   • otherwise the oldest live row decides; an unbounded cutoff always reads.
// No `off` is returned: a merged table has no single sheet-row mapping, so a
// caller that writes by row index must read the live sheet itself.
function readHistorySince(sheetName, archiveName, tsColIdx, cutoff, colSpec) {
  var live = readRowsSince(getSheet(sheetName), tsColIdx, cutoff, colSpec);
  var bounded = cutoff instanceof Date && !isNaN(cutoff.getTime());
  var needArchive = true;
  if (bounded && /^(rows|none)/.test(String(live.mode))) needArchive = false;
  if (bounded && needArchive && live.rows.length > 1) {
    var oldestLive = parseSheetDateTime(live.rows[1][tsColIdx]);
    if (oldestLive && oldestLive.getTime() <= cutoff.getTime()) needArchive = false;
  }
  var arch = needArchive ? getSheetIfExists(archiveName) : null;
  if (!arch || arch.getLastRow() < 2) return { rows: live.rows, mode: live.mode + '+arch:0' };
  var archRead = readRowsSince(arch, tsColIdx, cutoff, colSpec);
  var header = live.rows.length ? live.rows.slice(0, 1) : archRead.rows.slice(0, 1);
  return { rows: header.concat(archRead.rows.slice(1), live.rows.slice(1)),
    mode: live.mode + '+arch:' + Math.max(0, archRead.rows.length - 1) };
}
function readResultsSince(cutoff, colSpec) { return readHistorySince('תוצאות', RESULTS_ARCHIVE_SHEET, 0, cutoff, colSpec); }
function readPendingSince(cutoff, colSpec) { return readHistorySince('ממתינים', PENDING_ARCHIVE_SHEET, 4, cutoff, colSpec); }

// ========== Nightly archive of the sheets that grow forever (B5) ==========
// Every remaining heavy handler reads a sheet that grows with every exam ever
// taken: submitResult pulled ~15 MB per call (8.2 MB of it 'מבחנים', read whole
// to use ONE row), and the dashboards read 'תוצאות' several times per poll.
// 'ממתינים' has been archived since r5; this job does the same for the other
// two, under one lock, one guard and one time budget:
//   ממתינים 14 days → ממתינים_ארכיון   (live readers filter by session code)
//   מבחנים   2 days → מבחנים_ארכיון    (a question map is read while the attempt is scored)
//   תוצאות  30 days → תוצאות_ארכיון    (history readers use readResultsSince)
// Nothing is deleted: every row is copied first and the archive keeps the full
// width. Run archiveSheets() once by hand, then installNightlyJobs().
var PENDING_ARCHIVE_SHEET = 'ממתינים_ארכיון';
var RESULTS_ARCHIVE_SHEET = 'תוצאות_ארכיון';
var EXAMS_ARCHIVE_SHEET = 'מבחנים_ארכיון';
var PENDING_ARCHIVE_RETAIN_DAYS = 14;
var PENDING_TERMINAL = { completed: 1, disqualified: 1, dq_confirmed: 1, cancelled: 1, rejected: 1 };

// tsCol is the 0-based index of the timestamp column each sheet is ordered by.
var ARCHIVE_PLAN = [
  { name: 'ממתינים', archive: PENDING_ARCHIVE_SHEET, tsCol: 4, retainDays: PENDING_ARCHIVE_RETAIN_DAYS },
  { name: 'מבחנים', archive: EXAMS_ARCHIVE_SHEET, tsCol: 3, retainDays: 2 },
  { name: 'תוצאות', archive: RESULTS_ARCHIVE_SHEET, tsCol: 0, retainDays: 30 }
];
var ARCHIVE_BUDGET_MS = 4.5 * 60 * 1000;   // stay under the 6-min ceiling; the rest moves next run
var ARCHIVE_CHUNK = 300;
var ARCHIVE_QUIET_HOURS = 3;               // no non-terminal registration this recent

function archiveSheets() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) { Logger.log('archive: another run holds the lock'); return { skipped: 'locked' }; }
  var t0 = Date.now(), deadline = t0 + ARCHIVE_BUDGET_MS;
  try {
    var pendingRows = getSheet('ממתינים').getDataRange().getValues();
    var blocker = archiveBlockedReason(pendingRows);
    if (blocker) { Logger.log('archive: ' + blocker + ' — skipping this run'); return { skipped: blocker }; }
    var moved = {}, stopped = '';
    for (var i = 0; i < ARCHIVE_PLAN.length; i++) {
      var plan = ARCHIVE_PLAN[i];
      if (Date.now() > deadline) { stopped = plan.name; break; }
      var res = archiveOneSheet(plan, deadline, plan.name === 'ממתינים' ? pendingRows : null);
      moved[plan.name] = res.moved;
      if (res.stopped) { stopped = plan.name; break; }
    }
    // The only scheduled run of the day is also the only sweeper of the
    // diagnostics that a killed execution left behind (the hourly warmup used
    // to do it, and the warmup is gone). flushDiagnostics() does it on demand.
    var diag = diagSweep(null);
    Logger.log('archive: ' + JSON.stringify(moved) + (stopped ? ' — budget hit at ' + stopped + ', rest next run' : '') +
      ' in ' + (Date.now() - t0) + 'ms; diagnostics: ' + diag.swept + ' killed / ' + diag.flushed + ' parked');
    return { moved: moved, stopped: stopped, diagnostics: diag, ms: Date.now() - t0 };
  } finally {
    lock.releaseLock();
  }
}

// Safety: never move a row while an exam may be in progress — a live handler
// holds row indexes across its own reads and writes. Two independent signals,
// either one blocks the WHOLE run (all three sheets share the same risk).
function archiveBlockedReason(pendingRows) {
  var quietSince = Date.now() - ARCHIVE_QUIET_HOURS * 3600000;
  for (var i = 1; i < pendingRows.length; i++) {
    var st = String(pendingRows[i][5] || '').trim();
    if (PENDING_TERMINAL[st]) continue;
    var reg = parseSheetDateTime(pendingRows[i][4]);
    if (reg && reg.getTime() > quietSince) return 'a non-terminal registration in the last ' + ARCHIVE_QUIET_HOURS + 'h';
  }
  var sessions = getSheet('סשנים').getDataRange().getValues();
  var now = Date.now();
  for (var s = 1; s < sessions.length; s++) {
    var active = sessions[s][10] === true || String(sessions[s][10]).toUpperCase() === 'TRUE';
    if (!active) continue;
    var validUntil = parseSheetDateTime(sessions[s][9]);
    if (!validUntil || validUntil.getTime() > now) return 'session ' + String(sessions[s][0]) + ' is still open';
  }
  return '';
}

// Copy → delete, in chunks, OLDEST FIRST. Two invariants depend on the order:
//   • within a chunk the rows are deleted bottom-up, so the row numbers still
//     to be deleted stay valid;
//   • across chunks the oldest rows leave first, so an interrupted run (time
//     budget) still leaves "live = the newest rows, archive = everything older"
//     — which is exactly what readHistorySince relies on when it decides it can
//     skip the archive. Moving the newest candidates first would shuffle the
//     two sheets and make that decision wrong.
// Each chunk deletes rows ABOVE every later candidate, so the row numbers of
// the later chunks shift up by the number already deleted; `deleted` tracks it.
// Rows appended by a concurrent execution sit BELOW the whole set.
function archiveOneSheet(plan, deadline, preloadedRows) {
  var src = getSheetIfExists(plan.name);
  if (!src || src.getLastRow() <= 1) return { moved: 0, stopped: false };
  var rows = preloadedRows || src.getDataRange().getValues();
  if (rows.length <= 1) return { moved: 0, stopped: false };
  var cutoff = Date.now() - plan.retainDays * 86400000;
  var rowNums = [], vals = [], width = rows[0].length;
  for (var i = 1; i < rows.length; i++) {
    var ts = parseSheetDateTime(rows[i][plan.tsCol]);
    if (!ts || ts.getTime() >= cutoff) continue;   // unparseable timestamp → keep the row
    rowNums.push(i + 1);
    vals.push(rows[i]);
    if (rows[i].length > width) width = rows[i].length;
  }
  if (!rowNums.length) return { moved: 0, stopped: false };

  var arch = getSheet(plan.archive);
  if (arch.getLastRow() === 0) arch.getRange(1, 1, 1, width).setValues([padArchiveRow(rows[0], width)]);
  var moved = 0, deleted = 0;
  for (var from = 0; from < rowNums.length; from += ARCHIVE_CHUNK) {
    var to = Math.min(rowNums.length, from + ARCHIVE_CHUNK);
    var chunkRows = [], chunkVals = [];
    for (var v = from; v < to; v++) {
      chunkRows.push(rowNums[v] - deleted);
      chunkVals.push(padArchiveRow(vals[v], width));
    }
    arch.getRange(arch.getLastRow() + 1, 1, chunkVals.length, width).setValues(chunkVals);
    SpreadsheetApp.flush();
    deleteRowRuns(src, chunkRows);
    deleted += chunkRows.length;
    moved += chunkVals.length;
    if (Date.now() > deadline) return { moved: moved, stopped: true };
  }
  return { moved: moved, stopped: false };
}

function padArchiveRow(row, width) {
  var out = row.slice(0, width);
  while (out.length < width) out.push('');
  return out;
}

// chunkRows is ascending; delete contiguous runs from the bottom up.
function deleteRowRuns(sheet, chunkRows) {
  var k = chunkRows.length - 1;
  while (k >= 0) {
    var end = chunkRows[k], start = end;
    while (k - 1 >= 0 && chunkRows[k - 1] === start - 1) { k--; start = chunkRows[k]; }
    sheet.deleteRows(start, end - start + 1);
    k--;
  }
}

// Kept for one release: the 01:00 trigger of the previous version calls this
// name, and an operator may still have it in a runbook. Delete after 10/2026.
function archiveOldPendingRows() { return archiveSheets(); }

// Run ONCE from the editor after a deploy. Leaves EXACTLY two time triggers:
// archiveSheets 01:00 and rebuildAtRiskCache 03:00 (Asia/Jerusalem). Every
// trigger of a retired job is removed by NAME — the functions themselves may no
// longer exist in the script, but Apps Script keeps running their triggers and
// each run burns from the 90-minutes-a-day trigger budget.
var NIGHTLY_OBSOLETE_HANDLERS = ['archiveOldPendingRows', 'archiveSheets', 'warmupQuestionCaches',
  'ensureQuestionCachesWarm', 'rebuildMissingQuestionCaches', 'rebuildAtRiskCache'];
function installNightlyJobs() {
  var trigs = ScriptApp.getProjectTriggers(), removed = [];
  for (var i = 0; i < trigs.length; i++) {
    var fn = trigs[i].getHandlerFunction();
    if (NIGHTLY_OBSOLETE_HANDLERS.indexOf(fn) === -1) continue;
    ScriptApp.deleteTrigger(trigs[i]);
    removed.push(fn);
  }
  // 01:00 archive before the 03:00 forecast rebuild, so the forecast reads a settled sheet.
  ScriptApp.newTrigger('archiveSheets').timeBased().atHour(1).everyDays(1).inTimezone('Asia/Jerusalem').create();
  ScriptApp.newTrigger('rebuildAtRiskCache').timeBased().atHour(3).everyDays(1).inTimezone('Asia/Jerusalem').create();
  var msg = 'installNightlyJobs: removed ' + removed.length + ' old trigger(s) [' + removed.join(', ') +
    ']; created archiveSheets 01:00 + rebuildAtRiskCache 03:00 (Asia/Jerusalem)';
  Logger.log(msg);
  return msg;
}
// ---- One scan for "this examinee's current row" ----------------------------
// Ten handlers wrote this reverse loop by hand, which is how their status and
// 'בוטל' filters drifted apart (review E S7, C R12). rows[0] is a header.
//
// findLatestPendingRow: newest ממתינים row of (session, id). `statuses`, when
// given, keeps scanning past rows in other states instead of stopping at the
// first match — that is what "reset every stuck row" and "approve the waiting
// one" need. Returns { idx, row, status }; idx is an index into `rows` (the
// sheet row is idx + 1 + off) and is -1 when nothing matched.
function findLatestPendingRow(rows, sessionCode, idNumber, statuses) {
  var code = String(sessionCode || '').trim(), id = normalizeId(idNumber);
  for (var i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][0]).trim() !== code || normalizeId(rows[i][1]) !== id) continue;
    var st = String(rows[i][5] || '').trim();
    if (statuses && statuses.indexOf(st) === -1) continue;
    return { idx: i, row: rows[i], status: st };
  }
  return { idx: -1, row: null, status: '' };
}

// Newest 'תוצאות' row of (session, id). skipCancelled leaves 'בוטל' rows out:
// a correction must never land on a row that was already overturned (E S7).
function findLatestResultRow(rows, sessionCode, idNumber, skipCancelled) {
  var code = String(sessionCode || '').trim(), id = normalizeId(idNumber);
  for (var i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][13]).trim() !== code || normalizeId(rows[i][1]) !== id) continue;
    var st = String(rows[i][7] || '').trim();
    if (skipCancelled && st === 'בוטל') continue;
    return { idx: i, row: rows[i], status: st };
  }
  return { idx: -1, row: null, status: '' };
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

// r23: a valid verdict is cached for TOKEN_VERDICT_CACHE_SEC. Every examiner
// poll (2-5 s) used to re-read the whole 'בוחנים' sheet just to check a token
// that had been valid a few seconds earlier. Only positives are cached: a token
// that was just created must be usable at once, and one that was evicted or
// expired may linger for at most a minute.
var TOKEN_VERDICT_CACHE_SEC = 60;
function verifyToken(examinerId, token) {
  if (!examinerId || !token) return false;
  var key = CACHE_KEY_PREFIX + 'tok_' + normalizeId(examinerId) + '_' + String(token).slice(0, 80), cache = null;
  try { cache = CacheService.getScriptCache(); if (cache.get(key) === '1') return true; } catch (eGet) { cache = null; }
  var sheet = getSheet('בוחנים');
  var data = sheet.getDataRange().getValues();
  var valid = false;
  for (var i = 1; i < data.length; i++) {
    if (normalizeId(data[i][1]) === normalizeId(examinerId)) {
      var storedTokens = String(data[i][6] || '').split(',');
      var expiry = data[i][7];
      if (storedTokens.indexOf(token) === -1) break;
      if (!expiry) break;
      var expiryDate = expiry instanceof Date ? expiry : new Date(expiry);
      if (new Date() > expiryDate) break;
      valid = true;
      break;
    }
  }
  if (valid) { try { if (!cache) cache = CacheService.getScriptCache(); cache.put(key, '1', TOKEN_VERDICT_CACHE_SEC); } catch (ePut) {} }
  return valid;
}

function requireToken(p) {
  var valid = verifyToken(p.examinerId, p.token);
  diagMark('auth:token');   // a full 'בוחנים' read, paid by every examiner poll before its handler starts
  if (!valid) {
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
  'gateway',           // the session-poll Worker (sessionSnapshot; it also holds GATEWAY_KEY)
  'localhost-dev'      // local development
];
function checkOrigin(p) {
  // Allowed: actions called from external services (none currently) or no origin enforcement on certain reads
  // An empty action is the public 'is the API running' ping and needs no origin.
  if (String(p.action || '') === '') return null;
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

// One row of 'ממתינים' → token verdict. Per-examinee audio (column J) rides
// along on the row we already read, so callers get it without a second scan.
// - legacy: true when the stored row predates token support (empty cell) —
//   we accept the call but flag it so we can audit / tighten later.
// - reason values (when invalid): 'not_found', 'missing', 'mismatch'.
function examineeTokenVerdict(row, examineeToken) {
  var rowAudio = String(row[9] || '').trim() === 'on' ? 'on' : 'off';
  var storedToken = String((row.length > 12 ? row[12] : '') || '').trim();
  if (!storedToken) return { valid: true, legacy: true, audioMode: rowAudio };
  if (!examineeToken) return { valid: false, reason: 'missing' };
  if (String(examineeToken).trim() === storedToken) return { valid: true, legacy: false, audioMode: rowAudio };
  return { valid: false, reason: 'mismatch' };
}

// The token check reads the tail of 'ממתינים'; the handler that runs right after
// it needs the very same row — its status, its sheet row number (the status
// write needs it), the audio flag and the time extension. Handing that read
// forward is what keeps startExam and submitResult at ONE read of the sheet.
// The context is SINGLE USE — the auth check hands it to the handler that runs
// immediately after it, and anything later reads the sheet again. Nothing is
// ever decided from a row this request did not just read.
//   latest = the newest row of this examinee, whatever its status (the token
//            and the "may they submit" rule are decided on it, as before)
//   active = the newest approved/in_exam row (the attempt being started)
var EXAMINEE_ROW_CONTEXT = null;
function examineeRowContext(sessionCode, idNumber, fresh) {
  var cached = EXAMINEE_ROW_CONTEXT;
  EXAMINEE_ROW_CONTEXT = null;
  if (!fresh && cached && String(cached.sessionCode) === String(sessionCode) &&
      normalizeId(cached.idNumber) === normalizeId(idNumber)) return cached;
  diagMark('sheet:pending-examinee');
  var tail = readPendingTail(), rows = tail.rows;
  var ctx = { sessionCode: sessionCode, idNumber: idNumber, tail: tail, latest: null, active: null };
  for (var i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][0]) !== String(sessionCode) || normalizeId(rows[i][1]) !== normalizeId(idNumber)) continue;
    var entry = { row: rows[i], rowNumber: i + tail.off + 1, status: String(rows[i][5] || '').trim() };
    if (!ctx.latest) ctx.latest = entry;
    if (!ctx.active && (entry.status === 'approved' || entry.status === 'in_exam')) ctx.active = entry;
    if (ctx.latest && ctx.active) break;
  }
  EXAMINEE_ROW_CONTEXT = ctx;
  return ctx;
}

// Returns { valid: bool, legacy: bool, reason: string }
function verifyExamineeToken(sessionCode, idNumber, examineeToken) {
  var ctx = examineeRowContext(sessionCode, idNumber, true);   // auth always reads fresh
  if (!ctx.latest) return { valid: false, reason: 'not_found' };
  var verdict = examineeTokenVerdict(ctx.latest.row, examineeToken);
  ctx.audioMode = verdict.audioMode || 'off';
  return verdict;
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
  // One rule, one read: examinerOwnsSession (44_sessions_misc.js) serves the
  // check from the per-execution memo of 'סשנים', so a handler that verifies
  // ownership and then reads the session row pays for the sheet once.
  return examinerOwnsSession(sessionCode, examinerId);
}

// Key prefix shared by every CacheService / ScriptProperties entry this script
// owns (pending snapshots, extra-minutes maps, token verdicts, diagnostics).
// Declared here — a module that nothing in the roadmap deletes — so the live
// keys keep their names no matter which subsystem is retired; the question
// cache declares the same literal value for as long as it exists.
var CACHE_KEY_PREFIX = 'qv2_';

// Attempt number = how many non-'בוטל' result rows this examinee already has
// for this licence — in the LIVE sheet and in the archive (B5: 'תוצאות' is
// archived after 30 days, and without the archive a retake three months later
// would be recorded as attempt 1).
// `liveRows` is an optional rows array the caller already holds (rows[0] =
// header); without it we read the live sheet ourselves in the three columns
// this needs: B (ת.ז.), E (דרגה), H (עבר/נכשל). The old version re-read the
// whole sheet up to three times per call to re-validate its own input.
var RESULTS_ATTEMPT_COLSPEC = [[2, 1], [5, 1], [8, 1]];
function countAttempts(idNumber, license, liveRows) {
  var wantId = normalizeId(idNumber), wantLic = String(license);
  var count = countAttemptRows(liveRows || readAttemptColumns(getSheet('תוצאות')), wantId, wantLic);
  var arch = getSheetIfExists(RESULTS_ARCHIVE_SHEET);
  if (arch) count += countAttemptRows(readAttemptColumns(arch), wantId, wantLic);
  return count;
}
// Live + archive attempt columns as ONE table (oldest first), for a caller that
// counts attempts for SEVERAL examinees in one request — the dashboard's
// reconciliation used to re-read the whole 'תוצאות' sheet once per stale row
// (review C R1: 40 stale rows = 6.8 M cells in one poll).
function readAttemptHistory() {
  var live = readAttemptColumns(getSheet('תוצאות'));
  var arch = getSheetIfExists(RESULTS_ARCHIVE_SHEET);
  if (!arch) return live;
  var archRows = readAttemptColumns(arch);
  var header = live.length ? live.slice(0, 1) : archRows.slice(0, 1);
  return header.concat(archRows.slice(1), live.slice(1));
}

function readAttemptColumns(sheet) {
  var lastRow = sheet.getLastRow(), lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];
  return readSheetSlice(sheet, 1, lastRow, lastCol, RESULTS_ATTEMPT_COLSPEC);
}
function countAttemptRows(rows, wantId, wantLic) {
  var count = 0;
  for (var i = 1; i < rows.length; i++) {
    if (normalizeId(rows[i][1]) !== wantId || String(rows[i][4]) !== wantLic) continue;
    if (String(rows[i][7] || '').trim() === 'בוטל') continue;   // overturned DQ is not a real attempt
    count++;
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

// ========== The question index — all the server knows about questions ========
//
// Until 21/09/2026 the server held the question TEXTS: seven language banks in a
// private Drive folder, copied into CacheService as per-license pools, kept warm
// by a trigger, guarded by leases and rebuilt out of band. That subsystem was
// ~1,500 lines and it is what a killed execution died inside (the 360s kills of
// 09-09/09-11), what Drive reads added to every commander dashboard (r12/r13),
// and what made exam-start a 10-second request. It bought nothing: the texts AND
// the correct answers were already served to anyone who asked (getQuestionsByIds
// with any studentId, verified live 21/09).
//
// So the texts are now static files the client loads (bank/<lang>.json, built by
// tools/build_bank.js), and the server keeps only:
//   * QUESTION_INDEX — id → { c: {license: topic}, l: language bitmask, img }
//     (generated into this file at build time from deployment/question_index.json)
//   * the answer key (deployment/answer_key.gs, pasted separately, never public)
// From those two it can draw an exam and score it, with zero Drive and zero
// question data in CacheService.

// Generated from deployment/question_index.json by tools/build_server.js — 1700 ids. Do not edit.
var QUESTION_INDEX = {
"1":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"2":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"3":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"4":{"c":{"B":"חוק","C1":"חוק"},"l":127,"img":0},
"5":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"6":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"7":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"8":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"9":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"10":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"11":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"12":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":1},
"13":{"c":{"1":"חוק","B":"חוק","C1":"חוק"},"l":127,"img":0},
"14":{"c":{"1":"חוק","B":"חוק","C1":"חוק"},"l":127,"img":0},
"15":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"16":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"17":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"18":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":1},
"19":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"20":{"c":{"C1":"ספציפי"},"l":127,"img":0},
"21":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"22":{"c":{"B":"חוק","C1":"חוק"},"l":127,"img":0},
"23":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"24":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"25":{"c":{"C1":"ספציפי"},"l":127,"img":0},
"26":{"c":{"1":"חוק","C1":"חוק","C":"חוק"},"l":127,"img":0},
"28":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"29":{"c":{"C1":"ספציפי","D":"ספציפי"},"l":127,"img":0},
"30":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"31":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"32":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"33":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"34":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"35":{"c":{"C1":"ספציפי"},"l":127,"img":0},
"36":{"c":{"1":"ספציפי"},"l":127,"img":0},
"37":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"39":{"c":{"B":"חוק","C1":"חוק"},"l":127,"img":0},
"40":{"c":{"B":"חוק","C1":"חוק"},"l":127,"img":0},
"41":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"42":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"43":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"44":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"45":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"46":{"c":{"B":"חוק","C1":"חוק"},"l":127,"img":1},
"47":{"c":{"B":"חוק","C1":"חוק","C":"חוק"},"l":127,"img":0},
"48":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"50":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"52":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"54":{"c":{"B":"חוק","C1":"חוק"},"l":127,"img":0},
"55":{"c":{"B":"חוק","C1":"חוק"},"l":127,"img":0},
"56":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"57":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"58":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"59":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"60":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"61":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"62":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"63":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"64":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"65":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"66":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"67":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"68":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"69":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"70":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"71":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"72":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"73":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"74":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"75":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"76":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"77":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"78":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"79":{"c":{"B":"חוק"},"l":127,"img":1},
"80":{"c":{"B":"חוק","C1":"חוק","D":"חוק"},"l":127,"img":0},
"81":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":1},
"82":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":1},
"83":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"84":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":1},
"85":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":1},
"86":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"87":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"88":{"c":{"1":"ספציפי"},"l":127,"img":1},
"89":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"90":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"91":{"c":{"B":"חוק","C1":"חוק"},"l":127,"img":0},
"92":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"93":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"94":{"c":{"1":"חוק","B":"חוק","C1":"חוק"},"l":127,"img":1},
"95":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"96":{"c":{"1":"חוק","B":"חוק","C1":"חוק"},"l":127,"img":0},
"97":{"c":{"1":"חוק","B":"חוק","C1":"חוק"},"l":127,"img":0},
"98":{"c":{"1":"חוק","B":"חוק","C1":"חוק"},"l":127,"img":0},
"99":{"c":{"1":"חוק","B":"חוק","C1":"חוק"},"l":127,"img":0},
"100":{"c":{"1":"חוק","B":"חוק","C1":"חוק"},"l":127,"img":0},
"101":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":1},
"102":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":1},
"103":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":1},
"104":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"105":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":1},
"106":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":1},
"107":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"108":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":1},
"109":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"110":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":1},
"111":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":1},
"112":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":1},
"113":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":1},
"114":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":1},
"115":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":1},
"116":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":1},
"117":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"118":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"119":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"120":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"121":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"122":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"123":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"124":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"125":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"126":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"127":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"128":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"129":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"130":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"131":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"132":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":1},
"133":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"134":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"135":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"136":{"c":{"1":"חוק","C1":"חוק","C":"חוק"},"l":127,"img":0},
"137":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"138":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":1},
"139":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":1},
"140":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"141":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"142":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"143":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"144":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"145":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"146":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"147":{"c":{"B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":1},
"148":{"c":{"1":"חוק","B":"חוק","C1":"חוק"},"l":127,"img":1},
"149":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"150":{"c":{"B":"חוק","C1":"חוק","C":"חוק","D":"ספציפי"},"l":127,"img":1},
"151":{"c":{"B":"חוק","C1":"חוק"},"l":127,"img":1},
"152":{"c":{"B":"חוק","C1":"חוק"},"l":127,"img":1},
"153":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"154":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":1},
"155":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":1},
"156":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":1},
"157":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"158":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":1},
"159":{"c":{"B":"חוק","C1":"חוק","C":"חוק"},"l":127,"img":0},
"160":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":1},
"161":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"162":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":1},
"163":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"164":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"165":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"166":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"167":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"168":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"169":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"170":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"171":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"172":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":1},
"173":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"174":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"175":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"176":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":1},
"177":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"178":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"179":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"180":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"181":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"182":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"183":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"184":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"185":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"186":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"187":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":1},
"188":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"189":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"190":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":1},
"191":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"192":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"193":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":1},
"194":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"195":{"c":{"B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"196":{"c":{"B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":1},
"197":{"c":{"B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":1},
"198":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"199":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"200":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"201":{"c":{"B":"חוק","C1":"חוק"},"l":127,"img":0},
"202":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":4,"img":0},
"203":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"205":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":1},
"206":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"207":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":1},
"208":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":1},
"209":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"210":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":1},
"211":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":1},
"212":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"213":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"214":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":1},
"215":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"216":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":4,"img":1},
"217":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"218":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"219":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"220":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"221":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"222":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"223":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"224":{"c":{"B":"חוק","C1":"חוק","C":"חוק","D":"ספציפי"},"l":127,"img":0},
"225":{"c":{"B":"חוק","C1":"חוק","C":"חוק","D":"ספציפי"},"l":127,"img":1},
"226":{"c":{"B":"חוק","C1":"חוק","C":"חוק","D":"ספציפי"},"l":127,"img":0},
"227":{"c":{"B":"חוק","C1":"חוק","C":"חוק"},"l":127,"img":1},
"228":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"229":{"c":{"B":"חוק","C1":"חוק","C":"חוק"},"l":127,"img":0},
"230":{"c":{"B":"חוק","C1":"חוק","C":"חוק"},"l":127,"img":0},
"231":{"c":{"B":"חוק","C1":"חוק","C":"חוק"},"l":127,"img":1},
"232":{"c":{"B":"חוק","C1":"חוק","C":"חוק"},"l":127,"img":0},
"233":{"c":{"B":"חוק","C1":"חוק","C":"חוק"},"l":127,"img":0},
"234":{"c":{"B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"235":{"c":{"B":"חוק","C1":"חוק","C":"חוק","D":"ספציפי"},"l":127,"img":0},
"236":{"c":{"B":"חוק","C1":"חוק","C":"חוק","D":"ספציפי"},"l":127,"img":0},
"237":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"238":{"c":{"B":"חוק","C1":"חוק"},"l":127,"img":0},
"240":{"c":{"1":"חוק","B":"חוק","C1":"חוק"},"l":127,"img":0},
"241":{"c":{"B":"חוק","C1":"חוק","C":"חוק"},"l":127,"img":0},
"242":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":1},
"243":{"c":{"1":"חוק","C1":"חוק","C":"חוק"},"l":127,"img":0},
"244":{"c":{"B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"245":{"c":{"1":"ספציפי"},"l":127,"img":0},
"246":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"247":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"248":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":4,"img":0},
"249":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"250":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"252":{"c":{"1":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"253":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"254":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"255":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"256":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"257":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"258":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"259":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"260":{"c":{"B":"חוק","C1":"חוק","C":"חוק","D":"ספציפי"},"l":127,"img":0},
"261":{"c":{"B":"חוק","C1":"חוק","C":"חוק","D":"ספציפי"},"l":127,"img":0},
"262":{"c":{"C1":"ספציפי","D":"ספציפי"},"l":127,"img":0},
"263":{"c":{"1":"חוק","C1":"חוק","C":"חוק","D":"ספציפי"},"l":127,"img":0},
"264":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"265":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"266":{"c":{"B":"חוק","C1":"חוק"},"l":4,"img":0},
"267":{"c":{"B":"חוק","C1":"חוק"},"l":127,"img":0},
"268":{"c":{"B":"חוק","C1":"חוק"},"l":127,"img":0},
"269":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"270":{"c":{"1":"חוק","B":"חוק","C1":"חוק"},"l":127,"img":0},
"271":{"c":{"1":"חוק","B":"חוק","C1":"חוק"},"l":127,"img":0},
"272":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"273":{"c":{"B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"274":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"275":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"276":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"277":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"278":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"279":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"280":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":1},
"282":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"283":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"284":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"285":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"286":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"287":{"c":{"C1":"ספציפי","D":"ספציפי"},"l":127,"img":0},
"288":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"289":{"c":{"D":"ספציפי"},"l":127,"img":0},
"290":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"291":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"292":{"c":{"C1":"ספציפי"},"l":127,"img":0},
"293":{"c":{"C":"ספציפי"},"l":127,"img":0},
"294":{"c":{"B":"חוק","C1":"חוק","C":"חוק","D":"ספציפי"},"l":127,"img":0},
"295":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"296":{"c":{"B":"חוק","C1":"חוק"},"l":127,"img":0},
"297":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"298":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"299":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"300":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"301":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"302":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"303":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"304":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"305":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"306":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"307":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"308":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"309":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"310":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"311":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"312":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"313":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"314":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"315":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"316":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"317":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"318":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"319":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"320":{"c":{"C1":"ספציפי"},"l":127,"img":0},
"321":{"c":{"C":"ספציפי"},"l":127,"img":0},
"322":{"c":{"C1":"ספציפי"},"l":127,"img":0},
"323":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"324":{"c":{"C1":"ספציפי","D":"ספציפי"},"l":127,"img":0},
"325":{"c":{"C1":"ספציפי"},"l":127,"img":0},
"326":{"c":{"C1":"ספציפי"},"l":127,"img":0},
"327":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"328":{"c":{"1":"חוק","B":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"329":{"c":{"1":"חוק","B":"חוק","C1":"חוק"},"l":127,"img":0},
"330":{"c":{"1":"חוק","B":"חוק","C1":"חוק"},"l":127,"img":0},
"331":{"c":{"B":"חוק","C1":"חוק","C":"חוק"},"l":127,"img":0},
"332":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"333":{"c":{"B":"חוק"},"l":127,"img":0},
"334":{"c":{"1":"ספציפי"},"l":127,"img":0},
"335":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"336":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"337":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"338":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"339":{"c":{"B":"חוק","C1":"חוק"},"l":127,"img":0},
"340":{"c":{"1":"ספציפי"},"l":127,"img":0},
"341":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"342":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"343":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"344":{"c":{"B":"חוק","C1":"חוק"},"l":127,"img":0},
"345":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"346":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"347":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"348":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":1},
"349":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"350":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"351":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"352":{"c":{"B":"חוק","C1":"חוק","C":"חוק","D":"ספציפי"},"l":127,"img":0},
"353":{"c":{"B":"חוק","C1":"חוק"},"l":127,"img":0},
"354":{"c":{"B":"חוק","C1":"חוק","C":"חוק","D":"ספציפי"},"l":127,"img":0},
"355":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"356":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"357":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"358":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"359":{"c":{"C1":"ספציפי","C":"ספציפי"},"l":127,"img":1},
"360":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"361":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"362":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"363":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"364":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"365":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"366":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"367":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"368":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"369":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"370":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"371":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"372":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"373":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"374":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"375":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"376":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"377":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"378":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"379":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"380":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"381":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"382":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"383":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"384":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"385":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"386":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"387":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"388":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"389":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":1},
"390":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"391":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"392":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"393":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"394":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"395":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"396":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"397":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"398":{"c":{"B":"חוק","C1":"חוק"},"l":127,"img":0},
"399":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"400":{"c":{"B":"תמרורים","C1":"תמרורים"},"l":127,"img":1},
"401":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":0},
"402":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"403":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"404":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"405":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"406":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"407":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"408":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"409":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"410":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"411":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"412":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"413":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"414":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"415":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"416":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"417":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"418":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"419":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"420":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"421":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"422":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"423":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"424":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"425":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"426":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"427":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"428":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"429":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"430":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"431":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"432":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"433":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"434":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"435":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"436":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"437":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"438":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"439":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"440":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"441":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"442":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"443":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"444":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"445":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"446":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"447":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"448":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"449":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"450":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"451":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"452":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"453":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"454":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"455":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"456":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"457":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"458":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"459":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"460":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"461":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"462":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"463":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"464":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"465":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"466":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"467":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"468":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"469":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"470":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"471":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"472":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"473":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"474":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"475":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"476":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"477":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":0},
"478":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":1},
"479":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"480":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"481":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"482":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"484":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"485":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"486":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"487":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"488":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"489":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים"},"l":127,"img":1},
"490":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"491":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"492":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"493":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"494":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"495":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"496":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"497":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"498":{"c":{"B":"תמרורים"},"l":127,"img":1},
"499":{"c":{"C1":"ספציפי","C":"ספציפי"},"l":127,"img":1},
"500":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"501":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"502":{"c":{"B":"תמרורים"},"l":127,"img":1},
"503":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"504":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"505":{"c":{"B":"תמרורים"},"l":127,"img":1},
"506":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"507":{"c":{"C":"ספציפי","D":"ספציפי"},"l":127,"img":1},
"508":{"c":{"C":"ספציפי","D":"ספציפי"},"l":127,"img":1},
"509":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים"},"l":127,"img":1},
"510":{"c":{"C1":"ספציפי","C":"ספציפי"},"l":127,"img":1},
"511":{"c":{"C":"ספציפי","D":"ספציפי"},"l":127,"img":1},
"512":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"513":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"514":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"515":{"c":{"C":"ספציפי","D":"ספציפי"},"l":127,"img":1},
"516":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"517":{"c":{"D":"ספציפי"},"l":127,"img":1},
"518":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"519":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"520":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"521":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"522":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"523":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"524":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"525":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"526":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"527":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"528":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"529":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"530":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"531":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"532":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"533":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"534":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"535":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"536":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"537":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"538":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"539":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"540":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"541":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"542":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"543":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"544":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"545":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"546":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"547":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"548":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"549":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"550":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"551":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"552":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"553":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"554":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"555":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"556":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"557":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"558":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"559":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"560":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"561":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"562":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"563":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"564":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"565":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"566":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"567":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"568":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"569":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"570":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"571":{"c":{"1":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"572":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"573":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"574":{"c":{"C1":"תמרורים","C":"תמרורים"},"l":4,"img":1},
"575":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"576":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"577":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"578":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"579":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"580":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"581":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"582":{"c":{"C1":"ספציפי","C":"ספציפי"},"l":127,"img":1},
"583":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"584":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"585":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"586":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"587":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"588":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"589":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"590":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"591":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"592":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"593":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"594":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"595":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"596":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"597":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"598":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"599":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"600":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"601":{"c":{"D":"ספציפי"},"l":127,"img":1},
"602":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"603":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"604":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"605":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"606":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":0},
"607":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"608":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"609":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"610":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"611":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"612":{"c":{"D":"ספציפי"},"l":127,"img":0},
"613":{"c":{"D":"ספציפי"},"l":127,"img":0},
"614":{"c":{"D":"ספציפי"},"l":127,"img":0},
"615":{"c":{"D":"ספציפי"},"l":127,"img":0},
"616":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"617":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"618":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"619":{"c":{"B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"620":{"c":{"D":"ספציפי"},"l":127,"img":0},
"621":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"622":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"623":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"624":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"625":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"626":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"627":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"628":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"629":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"630":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"631":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"632":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"633":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"634":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"635":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"636":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"637":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"638":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"639":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"640":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"641":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"642":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"643":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"644":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"645":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"646":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"647":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"648":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"649":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"650":{"c":{"D":"ספציפי"},"l":127,"img":1},
"651":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"652":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"653":{"c":{"B":"חוק","C1":"חוק"},"l":127,"img":1},
"654":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"655":{"c":{"B":"חוק","C1":"חוק","C":"חוק","D":"ספציפי"},"l":127,"img":0},
"656":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"657":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"658":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"659":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"660":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"661":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"662":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"663":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"664":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"665":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"666":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"667":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"668":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"669":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"670":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"671":{"c":{"B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"672":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"673":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"674":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"675":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"676":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"677":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"678":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"679":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"680":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"681":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"682":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"683":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"684":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"685":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"686":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"687":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"688":{"c":{"1":"הכרת הרכב","B":"הכרת הרכב","C1":"הכרת הרכב","C":"הכרת הרכב","D":"הכרת הרכב"},"l":127,"img":0},
"689":{"c":{"1":"הכרת הרכב","B":"הכרת הרכב","C1":"הכרת הרכב","C":"הכרת הרכב","D":"הכרת הרכב"},"l":127,"img":0},
"690":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"691":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":0},
"692":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"693":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"694":{"c":{"C1":"ספציפי","C":"ספציפי"},"l":127,"img":1},
"695":{"c":{"C1":"ספציפי","C":"ספציפי"},"l":127,"img":1},
"696":{"c":{"B":"הכרת הרכב","C1":"הכרת הרכב","C":"הכרת הרכב","D":"הכרת הרכב"},"l":127,"img":0},
"697":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":1},
"698":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":1},
"699":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":1},
"700":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":1},
"701":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":1},
"702":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":1},
"703":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":1},
"704":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":1},
"705":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":1},
"706":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":1},
"707":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":1},
"708":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":1},
"709":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":1},
"710":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":1},
"711":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":1},
"712":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":1},
"714":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":1},
"715":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":1},
"716":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":1},
"717":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":1},
"718":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":1},
"719":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":1},
"720":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":1},
"721":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":1},
"722":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"723":{"c":{"D":"ספציפי"},"l":127,"img":0},
"724":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":1},
"725":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":1},
"726":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":1},
"727":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":1},
"728":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":1},
"729":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":1},
"730":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"731":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"732":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"733":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"734":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"735":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"736":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"737":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות"},"l":127,"img":0},
"738":{"c":{"B":"בטיחות","C1":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"739":{"c":{"B":"בטיחות","C1":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"740":{"c":{"B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"741":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"742":{"c":{"B":"הכרת הרכב","C1":"הכרת הרכב","C":"הכרת הרכב","D":"הכרת הרכב"},"l":127,"img":0},
"744":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"745":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"746":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":4,"img":1},
"747":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"748":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"749":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"750":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"752":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"753":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"754":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"755":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"756":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"757":{"c":{"C1":"ספציפי"},"l":127,"img":0},
"758":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"759":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"760":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"761":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"762":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"763":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"764":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"765":{"c":{"1":"הכרת הרכב","B":"הכרת הרכב","C1":"הכרת הרכב","C":"הכרת הרכב","D":"הכרת הרכב"},"l":127,"img":0},
"766":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"767":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"768":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"769":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"770":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"771":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"772":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"774":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"775":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"776":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"777":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"778":{"c":{"B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"779":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"780":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"781":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"782":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"783":{"c":{"B":"חוק"},"l":127,"img":0},
"784":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":1},
"785":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"786":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"787":{"c":{"D":"ספציפי"},"l":127,"img":1},
"788":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"789":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"790":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"791":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"792":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"793":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"794":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"795":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"796":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":1},
"797":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"798":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"799":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"800":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"801":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"802":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"803":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"804":{"c":{"B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"805":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"806":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"807":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"808":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"809":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"810":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"811":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"812":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"813":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"814":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"815":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"816":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"817":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"818":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"819":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"820":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"821":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"822":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"823":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"824":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"825":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"826":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"827":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"828":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"829":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"830":{"c":{"1":"הכרת הרכב","B":"הכרת הרכב","C1":"הכרת הרכב","C":"הכרת הרכב","D":"הכרת הרכב"},"l":127,"img":0},
"831":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"832":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"833":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"834":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"835":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"836":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"837":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"838":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"839":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"840":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"841":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"842":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"843":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"844":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"845":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"846":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"847":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"848":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"849":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"850":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"851":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"852":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"853":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"854":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"855":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"856":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"857":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"858":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"859":{"c":{"B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"860":{"c":{"B":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"861":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"862":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"863":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"864":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"865":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"866":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"867":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"868":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"869":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"870":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"871":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"872":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"873":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"874":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"875":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"876":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"877":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"878":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"879":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"880":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"881":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"882":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"883":{"c":{"C1":"ספציפי","D":"ספציפי"},"l":127,"img":1},
"884":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"885":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"886":{"c":{"B":"הכרת הרכב","C1":"הכרת הרכב","C":"הכרת הרכב","D":"הכרת הרכב"},"l":127,"img":0},
"887":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"888":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"889":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"890":{"c":{"B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"891":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"892":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"893":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"895":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"896":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"897":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"898":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"899":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"900":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"901":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"902":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"903":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"904":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"905":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"906":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"907":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות"},"l":125,"img":1},
"908":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"909":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"910":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"911":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"912":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"913":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"914":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"915":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"916":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"917":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"918":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"919":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"920":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"921":{"c":{"B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"922":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"923":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"924":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"925":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"926":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"927":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"928":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"929":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"930":{"c":{"B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"931":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"932":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"933":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"934":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"935":{"c":{"1":"הכרת הרכב","B":"הכרת הרכב","C1":"הכרת הרכב","C":"הכרת הרכב","D":"הכרת הרכב"},"l":127,"img":0},
"936":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"937":{"c":{"B":"בטיחות","C1":"בטיחות"},"l":127,"img":1},
"938":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"939":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"940":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"941":{"c":{"B":"חוק","C1":"חוק"},"l":127,"img":1},
"942":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"943":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"944":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"945":{"c":{"B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"946":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"947":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"948":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"949":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"950":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"951":{"c":{"B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"952":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"953":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"954":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"955":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"956":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"957":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"958":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"960":{"c":{"B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"ספציפי"},"l":127,"img":1},
"961":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"962":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":1},
"963":{"c":{"B":"חוק"},"l":127,"img":1},
"964":{"c":{"B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"ספציפי"},"l":127,"img":1},
"965":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"966":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"967":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":1},
"968":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"969":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"970":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"971":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"972":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"973":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"974":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"975":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"976":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"977":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"978":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"979":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"980":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"981":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"982":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"983":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"984":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"985":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"986":{"c":{"1":"הכרת הרכב","B":"הכרת הרכב","C1":"הכרת הרכב","C":"הכרת הרכב","D":"הכרת הרכב"},"l":127,"img":1},
"987":{"c":{"1":"הכרת הרכב","B":"הכרת הרכב","C1":"הכרת הרכב","C":"הכרת הרכב","D":"הכרת הרכב"},"l":127,"img":0},
"988":{"c":{"1":"הכרת הרכב","B":"הכרת הרכב","C1":"הכרת הרכב","C":"הכרת הרכב","D":"הכרת הרכב"},"l":127,"img":0},
"989":{"c":{"1":"הכרת הרכב","B":"הכרת הרכב","C1":"הכרת הרכב"},"l":127,"img":0},
"990":{"c":{"1":"הכרת הרכב","B":"הכרת הרכב","C1":"הכרת הרכב"},"l":127,"img":0},
"991":{"c":{"1":"הכרת הרכב","B":"הכרת הרכב","C1":"הכרת הרכב"},"l":127,"img":0},
"992":{"c":{"1":"הכרת הרכב","B":"הכרת הרכב","C1":"הכרת הרכב"},"l":127,"img":0},
"993":{"c":{"1":"הכרת הרכב","B":"הכרת הרכב","C1":"הכרת הרכב"},"l":127,"img":0},
"994":{"c":{"B":"הכרת הרכב","C1":"הכרת הרכב"},"l":127,"img":0},
"995":{"c":{"B":"הכרת הרכב","C1":"הכרת הרכב"},"l":127,"img":0},
"996":{"c":{"1":"הכרת הרכב","B":"הכרת הרכב"},"l":127,"img":0},
"997":{"c":{"B":"הכרת הרכב","C1":"הכרת הרכב"},"l":127,"img":0},
"998":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"999":{"c":{"1":"הכרת הרכב","B":"הכרת הרכב","C1":"הכרת הרכב","C":"הכרת הרכב","D":"הכרת הרכב"},"l":127,"img":0},
"1000":{"c":{"B":"הכרת הרכב","C1":"הכרת הרכב"},"l":127,"img":0},
"1001":{"c":{"B":"הכרת הרכב","C1":"הכרת הרכב"},"l":127,"img":0},
"1002":{"c":{"B":"הכרת הרכב","C1":"הכרת הרכב"},"l":127,"img":0},
"1003":{"c":{"B":"הכרת הרכב","C1":"הכרת הרכב"},"l":127,"img":0},
"1004":{"c":{"B":"הכרת הרכב","C1":"הכרת הרכב"},"l":127,"img":0},
"1005":{"c":{"B":"הכרת הרכב"},"l":127,"img":0},
"1006":{"c":{"B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"1007":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"1008":{"c":{"B":"הכרת הרכב","C1":"הכרת הרכב"},"l":127,"img":0},
"1009":{"c":{"B":"הכרת הרכב","C1":"הכרת הרכב"},"l":127,"img":0},
"1010":{"c":{"B":"הכרת הרכב","C1":"הכרת הרכב"},"l":127,"img":0},
"1011":{"c":{"B":"הכרת הרכב","C1":"הכרת הרכב"},"l":127,"img":0},
"1012":{"c":{"B":"הכרת הרכב","C1":"הכרת הרכב"},"l":127,"img":0},
"1013":{"c":{"B":"הכרת הרכב"},"l":127,"img":0},
"1014":{"c":{"B":"הכרת הרכב"},"l":127,"img":0},
"1015":{"c":{"B":"הכרת הרכב","C1":"הכרת הרכב"},"l":127,"img":0},
"1016":{"c":{"B":"הכרת הרכב","C1":"הכרת הרכב"},"l":127,"img":0},
"1017":{"c":{"B":"הכרת הרכב","C1":"הכרת הרכב"},"l":127,"img":0},
"1018":{"c":{"B":"הכרת הרכב","C1":"הכרת הרכב"},"l":127,"img":0},
"1019":{"c":{"B":"הכרת הרכב","C1":"הכרת הרכב"},"l":127,"img":0},
"1020":{"c":{"B":"הכרת הרכב","C1":"הכרת הרכב"},"l":127,"img":0},
"1021":{"c":{"B":"הכרת הרכב","C1":"הכרת הרכב"},"l":127,"img":0},
"1022":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"1023":{"c":{"1":"הכרת הרכב","B":"הכרת הרכב","C1":"הכרת הרכב"},"l":127,"img":0},
"1024":{"c":{"B":"הכרת הרכב","C1":"הכרת הרכב"},"l":127,"img":0},
"1025":{"c":{"B":"הכרת הרכב","C1":"הכרת הרכב"},"l":127,"img":0},
"1026":{"c":{"B":"הכרת הרכב","C1":"הכרת הרכב"},"l":127,"img":0},
"1027":{"c":{"1":"הכרת הרכב","B":"הכרת הרכב","C1":"הכרת הרכב","C":"הכרת הרכב","D":"הכרת הרכב"},"l":127,"img":0},
"1028":{"c":{"B":"הכרת הרכב","C1":"הכרת הרכב"},"l":127,"img":0},
"1029":{"c":{"B":"הכרת הרכב","C1":"הכרת הרכב"},"l":127,"img":0},
"1030":{"c":{"B":"הכרת הרכב","C1":"הכרת הרכב"},"l":127,"img":0},
"1031":{"c":{"1":"הכרת הרכב","B":"הכרת הרכב","C1":"הכרת הרכב"},"l":127,"img":0},
"1032":{"c":{"B":"הכרת הרכב","C1":"הכרת הרכב"},"l":127,"img":0},
"1033":{"c":{"B":"הכרת הרכב","C1":"הכרת הרכב"},"l":127,"img":0},
"1034":{"c":{"1":"הכרת הרכב","B":"הכרת הרכב","C1":"הכרת הרכב"},"l":127,"img":0},
"1035":{"c":{"B":"הכרת הרכב","C1":"הכרת הרכב"},"l":127,"img":0},
"1036":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"1037":{"c":{"B":"הכרת הרכב","C1":"הכרת הרכב"},"l":127,"img":0},
"1038":{"c":{"B":"הכרת הרכב","C1":"הכרת הרכב"},"l":127,"img":0},
"1039":{"c":{"B":"הכרת הרכב","C1":"הכרת הרכב"},"l":127,"img":0},
"1040":{"c":{"B":"הכרת הרכב"},"l":127,"img":0},
"1041":{"c":{"B":"הכרת הרכב","C1":"הכרת הרכב"},"l":127,"img":0},
"1042":{"c":{"1":"הכרת הרכב","B":"הכרת הרכב","C1":"הכרת הרכב"},"l":127,"img":0},
"1043":{"c":{"B":"בטיחות"},"l":127,"img":0},
"1044":{"c":{"B":"בטיחות","C1":"בטיחות"},"l":127,"img":0},
"1045":{"c":{"B":"הכרת הרכב","C1":"הכרת הרכב"},"l":127,"img":0},
"1046":{"c":{"1":"הכרת הרכב","B":"הכרת הרכב","C1":"הכרת הרכב"},"l":127,"img":0},
"1047":{"c":{"B":"הכרת הרכב","C1":"הכרת הרכב"},"l":127,"img":0},
"1048":{"c":{"B":"הכרת הרכב","C1":"הכרת הרכב"},"l":127,"img":0},
"1049":{"c":{"B":"הכרת הרכב","C1":"הכרת הרכב"},"l":127,"img":0},
"1050":{"c":{"B":"הכרת הרכב","C1":"הכרת הרכב"},"l":127,"img":0},
"1051":{"c":{"B":"הכרת הרכב","C1":"הכרת הרכב"},"l":127,"img":0},
"1052":{"c":{"1":"הכרת הרכב","B":"הכרת הרכב","C1":"הכרת הרכב"},"l":127,"img":0},
"1053":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"1054":{"c":{"B":"בטיחות","C1":"בטיחות"},"l":127,"img":0},
"1055":{"c":{"B":"הכרת הרכב","C1":"הכרת הרכב"},"l":127,"img":0},
"1056":{"c":{"1":"הכרת הרכב","B":"הכרת הרכב","C1":"הכרת הרכב"},"l":127,"img":0},
"1057":{"c":{"1":"הכרת הרכב","B":"הכרת הרכב","C1":"הכרת הרכב"},"l":127,"img":0},
"1058":{"c":{"1":"הכרת הרכב","B":"הכרת הרכב","C1":"הכרת הרכב"},"l":127,"img":0},
"1059":{"c":{"1":"הכרת הרכב","B":"הכרת הרכב","C1":"הכרת הרכב"},"l":127,"img":0},
"1060":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"1061":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"1062":{"c":{"1":"הכרת הרכב","B":"הכרת הרכב","C1":"הכרת הרכב"},"l":127,"img":0},
"1063":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות"},"l":127,"img":0},
"1064":{"c":{"1":"הכרת הרכב","B":"הכרת הרכב","C1":"הכרת הרכב"},"l":127,"img":0},
"1065":{"c":{"1":"הכרת הרכב","B":"הכרת הרכב","C1":"הכרת הרכב"},"l":127,"img":0},
"1066":{"c":{"B":"הכרת הרכב","C1":"הכרת הרכב"},"l":127,"img":0},
"1067":{"c":{"1":"הכרת הרכב","B":"הכרת הרכב","C1":"הכרת הרכב"},"l":127,"img":0},
"1068":{"c":{"B":"הכרת הרכב","C1":"הכרת הרכב"},"l":127,"img":0},
"1069":{"c":{"B":"הכרת הרכב","C1":"הכרת הרכב"},"l":127,"img":0},
"1070":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1071":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"1072":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"1073":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"1074":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"1075":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"1076":{"c":{"B":"בטיחות"},"l":127,"img":0},
"1077":{"c":{"B":"הכרת הרכב","C1":"הכרת הרכב"},"l":127,"img":0},
"1078":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"1079":{"c":{"B":"הכרת הרכב","C1":"הכרת הרכב"},"l":127,"img":0},
"1080":{"c":{"B":"הכרת הרכב","C1":"הכרת הרכב"},"l":127,"img":0},
"1081":{"c":{"B":"הכרת הרכב","C1":"הכרת הרכב"},"l":127,"img":0},
"1082":{"c":{"B":"בטיחות","C1":"בטיחות"},"l":127,"img":0},
"1083":{"c":{"1":"ספציפי"},"l":127,"img":0},
"1084":{"c":{"1":"ספציפי"},"l":127,"img":0},
"1085":{"c":{"1":"בטיחות"},"l":127,"img":0},
"1086":{"c":{"1":"בטיחות"},"l":127,"img":0},
"1087":{"c":{"1":"בטיחות"},"l":127,"img":0},
"1088":{"c":{"1":"בטיחות"},"l":127,"img":0},
"1089":{"c":{"1":"ספציפי"},"l":127,"img":0},
"1090":{"c":{"1":"ספציפי"},"l":127,"img":0},
"1091":{"c":{"1":"בטיחות"},"l":127,"img":0},
"1092":{"c":{"1":"בטיחות"},"l":127,"img":0},
"1093":{"c":{"1":"בטיחות"},"l":127,"img":0},
"1094":{"c":{"1":"ספציפי"},"l":127,"img":0},
"1095":{"c":{"1":"ספציפי"},"l":127,"img":0},
"1096":{"c":{"1":"ספציפי"},"l":127,"img":0},
"1097":{"c":{"1":"ספציפי"},"l":127,"img":0},
"1098":{"c":{"1":"ספציפי"},"l":127,"img":0},
"1099":{"c":{"1":"ספציפי"},"l":127,"img":0},
"1100":{"c":{"1":"ספציפי"},"l":127,"img":0},
"1101":{"c":{"1":"ספציפי"},"l":127,"img":0},
"1102":{"c":{"1":"ספציפי"},"l":127,"img":0},
"1103":{"c":{"1":"ספציפי"},"l":127,"img":0},
"1104":{"c":{"1":"בטיחות"},"l":127,"img":0},
"1105":{"c":{"1":"ספציפי"},"l":127,"img":0},
"1106":{"c":{"1":"ספציפי"},"l":127,"img":0},
"1107":{"c":{"1":"בטיחות"},"l":127,"img":0},
"1108":{"c":{"1":"בטיחות"},"l":127,"img":0},
"1109":{"c":{"1":"ספציפי"},"l":127,"img":0},
"1110":{"c":{"1":"בטיחות"},"l":127,"img":0},
"1111":{"c":{"1":"בטיחות"},"l":127,"img":0},
"1112":{"c":{"1":"בטיחות"},"l":127,"img":0},
"1113":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1114":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1115":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1116":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1117":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1118":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1119":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1120":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1121":{"c":{"C1":"ספציפי","D":"ספציפי"},"l":127,"img":0},
"1122":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1123":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"1124":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1125":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1126":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1127":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1128":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1129":{"c":{"B":"חוק","D":"חוק"},"l":127,"img":0},
"1130":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1131":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1132":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1133":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1134":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1135":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1136":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1137":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1138":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1139":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1140":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1141":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1142":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1143":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1144":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1145":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1146":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1147":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1148":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1149":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1150":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1151":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1152":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1153":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1154":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1155":{"c":{"1":"חוק","D":"ספציפי"},"l":127,"img":0},
"1156":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1157":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1158":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1159":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1160":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1161":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1162":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1163":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1164":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1165":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1166":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1167":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1168":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1169":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1170":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1171":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1172":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1173":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1174":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1175":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1176":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1177":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1178":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1179":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1180":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1181":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1182":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1183":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"1184":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1185":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1186":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1187":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1188":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1189":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1190":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1191":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1192":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1193":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1194":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1195":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1196":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1197":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1198":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1199":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1200":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1201":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1202":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1203":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1204":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"1205":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1206":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1207":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1208":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1209":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1210":{"c":{"B":"חוק","C1":"חוק","D":"חוק"},"l":127,"img":0},
"1211":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1212":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1213":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1214":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1215":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1216":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1217":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1218":{"c":{"B":"חוק","C1":"חוק","C":"חוק","D":"ספציפי"},"l":127,"img":0},
"1219":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"1220":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1221":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1222":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1223":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1224":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1225":{"c":{"B":"חוק","C1":"חוק"},"l":127,"img":0},
"1226":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1227":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1228":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1229":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1230":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1231":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1232":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1233":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1234":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1235":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1236":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1237":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"1238":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1239":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1240":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1241":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1242":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1243":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1244":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1245":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1246":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1247":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1248":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1249":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1250":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1251":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1252":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1253":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1254":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1255":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1256":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1257":{"c":{"D":"ספציפי"},"l":127,"img":1},
"1258":{"c":{"D":"ספציפי"},"l":127,"img":1},
"1259":{"c":{"D":"ספציפי"},"l":127,"img":1},
"1260":{"c":{"D":"ספציפי"},"l":127,"img":1},
"1261":{"c":{"D":"ספציפי"},"l":127,"img":1},
"1262":{"c":{"D":"ספציפי"},"l":127,"img":1},
"1263":{"c":{"D":"ספציפי"},"l":127,"img":1},
"1264":{"c":{"D":"ספציפי"},"l":127,"img":1},
"1265":{"c":{"D":"ספציפי"},"l":127,"img":1},
"1266":{"c":{"D":"ספציפי"},"l":127,"img":1},
"1267":{"c":{"B":"תמרורים","D":"ספציפי"},"l":127,"img":1},
"1268":{"c":{"D":"ספציפי"},"l":127,"img":1},
"1269":{"c":{"D":"ספציפי"},"l":127,"img":1},
"1270":{"c":{"D":"ספציפי"},"l":127,"img":1},
"1271":{"c":{"D":"ספציפי"},"l":127,"img":1},
"1272":{"c":{"D":"ספציפי"},"l":127,"img":1},
"1273":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1274":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1275":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1276":{"c":{"C1":"ספציפי","C":"ספציפי"},"l":127,"img":1},
"1277":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1278":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"1279":{"c":{"C":"ספציפי"},"l":127,"img":0},
"1280":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1281":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1282":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1283":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1284":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1285":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1286":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1287":{"c":{"B":"הכרת הרכב","C1":"הכרת הרכב"},"l":127,"img":0},
"1288":{"c":{"B":"הכרת הרכב","C1":"הכרת הרכב"},"l":127,"img":0},
"1289":{"c":{"1":"הכרת הרכב","B":"הכרת הרכב","C1":"הכרת הרכב"},"l":127,"img":0},
"1290":{"c":{"1":"הכרת הרכב","B":"הכרת הרכב","C1":"הכרת הרכב"},"l":127,"img":0},
"1291":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"1292":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"1293":{"c":{"B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"1294":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1295":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1296":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"1297":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1298":{"c":{"B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"1299":{"c":{"C":"ספציפי"},"l":127,"img":0},
"1300":{"c":{"C":"ספציפי"},"l":127,"img":0},
"1301":{"c":{"C":"ספציפי"},"l":127,"img":0},
"1302":{"c":{"1":"חוק","C":"חוק"},"l":127,"img":0},
"1303":{"c":{"C":"ספציפי"},"l":127,"img":0},
"1304":{"c":{"C":"ספציפי"},"l":127,"img":0},
"1305":{"c":{"C":"ספציפי"},"l":127,"img":0},
"1306":{"c":{"C1":"ספציפי"},"l":127,"img":0},
"1307":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"1308":{"c":{"C1":"ספציפי"},"l":127,"img":0},
"1309":{"c":{"C":"ספציפי"},"l":127,"img":0},
"1310":{"c":{"C":"ספציפי"},"l":127,"img":0},
"1311":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1312":{"c":{"C":"ספציפי"},"l":127,"img":0},
"1313":{"c":{"C":"ספציפי"},"l":127,"img":0},
"1314":{"c":{"C":"ספציפי"},"l":127,"img":0},
"1315":{"c":{"C1":"ספציפי"},"l":127,"img":0},
"1316":{"c":{"C":"ספציפי"},"l":127,"img":0},
"1317":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1318":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1319":{"c":{"C1":"ספציפי"},"l":127,"img":0},
"1320":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק"},"l":127,"img":0},
"1321":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":1},
"1322":{"c":{"C":"ספציפי"},"l":127,"img":0},
"1323":{"c":{"C":"ספציפי"},"l":127,"img":0},
"1324":{"c":{"C":"ספציפי"},"l":127,"img":0},
"1325":{"c":{"B":"חוק","C1":"חוק","C":"חוק","D":"ספציפי"},"l":127,"img":0},
"1326":{"c":{"C":"ספציפי"},"l":127,"img":0},
"1327":{"c":{"C":"ספציפי"},"l":127,"img":0},
"1328":{"c":{"C":"ספציפי"},"l":127,"img":0},
"1329":{"c":{"C":"ספציפי"},"l":127,"img":0},
"1330":{"c":{"C":"ספציפי"},"l":127,"img":0},
"1331":{"c":{"C":"ספציפי"},"l":127,"img":0},
"1332":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"1333":{"c":{"C1":"ספציפי"},"l":127,"img":0},
"1334":{"c":{"B":"חוק","C1":"חוק"},"l":127,"img":0},
"1335":{"c":{"C1":"ספציפי"},"l":127,"img":0},
"1336":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1337":{"c":{"C":"ספציפי"},"l":127,"img":0},
"1338":{"c":{"1":"חוק","C1":"חוק","C":"חוק","D":"ספציפי"},"l":127,"img":0},
"1339":{"c":{"C1":"ספציפי"},"l":127,"img":0},
"1340":{"c":{"1":"חוק","C1":"חוק","C":"חוק"},"l":127,"img":0},
"1341":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1342":{"c":{"C1":"ספציפי"},"l":127,"img":0},
"1343":{"c":{"C":"ספציפי"},"l":127,"img":0},
"1344":{"c":{"C":"ספציפי"},"l":127,"img":0},
"1345":{"c":{"B":"חוק","C1":"חוק","C":"חוק","D":"ספציפי"},"l":127,"img":0},
"1346":{"c":{"B":"חוק","C1":"חוק","C":"חוק","D":"ספציפי"},"l":127,"img":0},
"1347":{"c":{"B":"חוק","C1":"חוק","C":"חוק","D":"ספציפי"},"l":127,"img":0},
"1348":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"1349":{"c":{"C1":"ספציפי"},"l":127,"img":0},
"1350":{"c":{"C1":"ספציפי"},"l":127,"img":0},
"1351":{"c":{"C1":"ספציפי"},"l":127,"img":0},
"1352":{"c":{"C":"ספציפי"},"l":127,"img":0},
"1353":{"c":{"C1":"ספציפי"},"l":127,"img":0},
"1354":{"c":{"C":"ספציפי"},"l":127,"img":0},
"1355":{"c":{"C":"ספציפי"},"l":127,"img":0},
"1356":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1357":{"c":{"B":"חוק","C1":"חוק","C":"חוק"},"l":127,"img":0},
"1358":{"c":{"C":"ספציפי"},"l":127,"img":0},
"1359":{"c":{"C1":"ספציפי"},"l":127,"img":0},
"1360":{"c":{"C1":"ספציפי","D":"ספציפי"},"l":127,"img":0},
"1361":{"c":{"C":"ספציפי"},"l":127,"img":0},
"1362":{"c":{"C1":"ספציפי"},"l":127,"img":0},
"1363":{"c":{"C1":"ספציפי","D":"ספציפי"},"l":127,"img":0},
"1364":{"c":{"C1":"ספציפי","D":"ספציפי"},"l":127,"img":0},
"1365":{"c":{"C":"ספציפי"},"l":127,"img":0},
"1366":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1367":{"c":{"C1":"ספציפי"},"l":127,"img":0},
"1368":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1369":{"c":{"C1":"ספציפי"},"l":127,"img":0},
"1370":{"c":{"1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1371":{"c":{"C1":"ספציפי","D":"ספציפי"},"l":127,"img":0},
"1372":{"c":{"C1":"ספציפי","D":"ספציפי"},"l":127,"img":0},
"1373":{"c":{"C1":"ספציפי"},"l":127,"img":0},
"1374":{"c":{"C":"ספציפי"},"l":127,"img":0},
"1375":{"c":{"C1":"ספציפי"},"l":127,"img":0},
"1376":{"c":{"1":"חוק","B":"חוק","C1":"חוק"},"l":127,"img":0},
"1377":{"c":{"C1":"ספציפי","D":"ספציפי"},"l":127,"img":0},
"1378":{"c":{"C1":"ספציפי","D":"ספציפי"},"l":127,"img":0},
"1379":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1380":{"c":{"1":"חוק","C":"חוק"},"l":127,"img":0},
"1381":{"c":{"B":"חוק","C1":"חוק","C":"חוק","D":"ספציפי"},"l":127,"img":0},
"1382":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"1383":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1384":{"c":{"C1":"חוק","C":"חוק"},"l":127,"img":0},
"1385":{"c":{"C":"ספציפי"},"l":127,"img":0},
"1386":{"c":{"C1":"ספציפי"},"l":127,"img":0},
"1387":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"1388":{"c":{"C1":"ספציפי","D":"ספציפי"},"l":127,"img":0},
"1389":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"1390":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"1391":{"c":{"C1":"ספציפי"},"l":127,"img":1},
"1392":{"c":{"C1":"ספציפי"},"l":127,"img":0},
"1393":{"c":{"C1":"ספציפי"},"l":127,"img":0},
"1394":{"c":{"C1":"ספציפי","D":"ספציפי"},"l":127,"img":0},
"1395":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1396":{"c":{"C1":"ספציפי","D":"ספציפי"},"l":127,"img":0},
"1397":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"1398":{"c":{"B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1399":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1400":{"c":{"1":"חוק","C1":"חוק","C":"חוק","D":"ספציפי"},"l":127,"img":0},
"1401":{"c":{"C":"ספציפי"},"l":127,"img":0},
"1402":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק"},"l":127,"img":0},
"1403":{"c":{"B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"ספציפי"},"l":127,"img":0},
"1404":{"c":{"1":"חוק","B":"חוק","C1":"חוק"},"l":127,"img":0},
"1405":{"c":{"C1":"ספציפי"},"l":127,"img":0},
"1406":{"c":{"C1":"ספציפי"},"l":127,"img":0},
"1407":{"c":{"C1":"ספציפי"},"l":127,"img":0},
"1408":{"c":{"C1":"ספציפי"},"l":127,"img":0},
"1409":{"c":{"C1":"ספציפי"},"l":127,"img":0},
"1410":{"c":{"C1":"ספציפי"},"l":127,"img":0},
"1411":{"c":{"C1":"ספציפי"},"l":127,"img":0},
"1412":{"c":{"C":"ספציפי"},"l":127,"img":0},
"1413":{"c":{"C":"ספציפי"},"l":127,"img":0},
"1414":{"c":{"C1":"ספציפי"},"l":127,"img":0},
"1415":{"c":{"C1":"ספציפי"},"l":127,"img":0},
"1416":{"c":{"C1":"ספציפי"},"l":127,"img":0},
"1417":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"1418":{"c":{"C":"ספציפי"},"l":127,"img":0},
"1419":{"c":{"C":"ספציפי"},"l":127,"img":0},
"1420":{"c":{"C":"ספציפי"},"l":127,"img":0},
"1421":{"c":{"C":"ספציפי"},"l":127,"img":0},
"1422":{"c":{"C":"ספציפי"},"l":127,"img":0},
"1423":{"c":{"C":"ספציפי"},"l":127,"img":0},
"1424":{"c":{"C":"ספציפי"},"l":127,"img":0},
"1425":{"c":{"C":"ספציפי"},"l":127,"img":0},
"1426":{"c":{"1":"ספציפי"},"l":127,"img":0},
"1427":{"c":{"C1":"ספציפי"},"l":127,"img":0},
"1428":{"c":{"C1":"ספציפי"},"l":127,"img":0},
"1429":{"c":{"C1":"ספציפי"},"l":127,"img":0},
"1430":{"c":{"C1":"ספציפי"},"l":127,"img":0},
"1431":{"c":{"C1":"ספציפי"},"l":127,"img":0},
"1432":{"c":{"C1":"ספציפי"},"l":127,"img":0},
"1433":{"c":{"C1":"ספציפי"},"l":127,"img":0},
"1434":{"c":{"C1":"ספציפי"},"l":127,"img":0},
"1435":{"c":{"C1":"ספציפי","D":"ספציפי"},"l":127,"img":0},
"1436":{"c":{"C1":"ספציפי"},"l":127,"img":0},
"1437":{"c":{"C":"ספציפי"},"l":127,"img":0},
"1438":{"c":{"C":"ספציפי"},"l":127,"img":0},
"1439":{"c":{"C":"ספציפי"},"l":127,"img":0},
"1440":{"c":{"C1":"ספציפי"},"l":127,"img":0},
"1441":{"c":{"C1":"ספציפי"},"l":127,"img":0},
"1442":{"c":{"C":"ספציפי"},"l":127,"img":0},
"1443":{"c":{"C":"ספציפי"},"l":127,"img":0},
"1444":{"c":{"B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1445":{"c":{"1":"ספציפי"},"l":127,"img":0},
"1446":{"c":{"C1":"ספציפי","D":"ספציפי"},"l":127,"img":0},
"1447":{"c":{"C1":"ספציפי"},"l":127,"img":0},
"1448":{"c":{"C1":"ספציפי"},"l":127,"img":0},
"1449":{"c":{"C1":"ספציפי"},"l":127,"img":0},
"1450":{"c":{"C1":"ספציפי"},"l":127,"img":0},
"1451":{"c":{"C1":"ספציפי"},"l":127,"img":0},
"1452":{"c":{"C1":"ספציפי"},"l":127,"img":0},
"1453":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1454":{"c":{"C1":"ספציפי"},"l":127,"img":0},
"1455":{"c":{"C1":"ספציפי"},"l":127,"img":0},
"1456":{"c":{"C1":"ספציפי"},"l":127,"img":0},
"1457":{"c":{"C1":"ספציפי"},"l":127,"img":0},
"1458":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1459":{"c":{"C1":"ספציפי"},"l":127,"img":0},
"1460":{"c":{"C1":"ספציפי"},"l":127,"img":0},
"1461":{"c":{"C":"ספציפי"},"l":127,"img":0},
"1462":{"c":{"C1":"ספציפי"},"l":127,"img":0},
"1463":{"c":{"C1":"ספציפי","D":"ספציפי"},"l":127,"img":0},
"1464":{"c":{"C1":"ספציפי"},"l":127,"img":0},
"1465":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1466":{"c":{"C1":"ספציפי","D":"ספציפי"},"l":127,"img":0},
"1467":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"1468":{"c":{"1":"הכרת הרכב","B":"הכרת הרכב","C1":"הכרת הרכב","C":"הכרת הרכב","D":"הכרת הרכב"},"l":127,"img":0},
"1469":{"c":{"C":"ספציפי"},"l":127,"img":0},
"1470":{"c":{"C":"ספציפי"},"l":127,"img":0},
"1471":{"c":{"B":"חוק","C1":"חוק"},"l":127,"img":0},
"1472":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1473":{"c":{"C":"ספציפי"},"l":127,"img":0},
"1474":{"c":{"C":"ספציפי"},"l":127,"img":0},
"1475":{"c":{"C":"ספציפי"},"l":127,"img":0},
"1476":{"c":{"B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"1477":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"1478":{"c":{"C":"ספציפי"},"l":127,"img":0},
"1479":{"c":{"C":"ספציפי"},"l":127,"img":0},
"1480":{"c":{"B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"1481":{"c":{"B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"1482":{"c":{"C1":"ספציפי"},"l":127,"img":0},
"1483":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"1484":{"c":{"B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1485":{"c":{"B":"הכרת הרכב","C1":"הכרת הרכב"},"l":127,"img":0},
"1486":{"c":{"B":"חוק"},"l":127,"img":0},
"1487":{"c":{"B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1488":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1489":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"1490":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"1491":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"1492":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"1493":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"1494":{"c":{"C1":"ספציפי"},"l":127,"img":0},
"1495":{"c":{"C1":"ספציפי"},"l":127,"img":0},
"1496":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1497":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"1498":{"c":{"1":"הכרת הרכב","B":"הכרת הרכב","C1":"הכרת הרכב"},"l":127,"img":0},
"1499":{"c":{"B":"הכרת הרכב","C1":"הכרת הרכב"},"l":127,"img":0},
"1500":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1501":{"c":{"B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":1},
"1502":{"c":{"C1":"ספציפי"},"l":127,"img":0},
"1503":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"1504":{"c":{"C":"ספציפי"},"l":127,"img":0},
"1505":{"c":{"C":"ספציפי"},"l":127,"img":0},
"1506":{"c":{"C1":"ספציפי"},"l":127,"img":0},
"1507":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"1508":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"1509":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"1510":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"1511":{"c":{"C1":"ספציפי"},"l":127,"img":0},
"1513":{"c":{"C1":"ספציפי"},"l":127,"img":0},
"1514":{"c":{"C":"ספציפי"},"l":127,"img":0},
"1515":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"1516":{"c":{"C1":"ספציפי"},"l":127,"img":0},
"1517":{"c":{"B":"חוק","C1":"חוק"},"l":127,"img":1},
"1518":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"1519":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"1520":{"c":{"C1":"ספציפי"},"l":127,"img":0},
"1521":{"c":{"C":"ספציפי"},"l":127,"img":0},
"1522":{"c":{"C1":"ספציפי"},"l":127,"img":0},
"1523":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"1524":{"c":{"C1":"ספציפי"},"l":127,"img":0},
"1525":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1526":{"c":{"C1":"ספציפי"},"l":127,"img":0},
"1527":{"c":{"C1":"ספציפי"},"l":127,"img":0},
"1528":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1529":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"1530":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1531":{"c":{"B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"1532":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"1533":{"c":{"B":"חוק","C1":"חוק","C":"חוק"},"l":127,"img":0},
"1534":{"c":{"B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"1535":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"1536":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1537":{"c":{"C":"ספציפי"},"l":127,"img":0},
"1538":{"c":{"B":"חוק","C1":"חוק","C":"חוק"},"l":127,"img":0},
"1539":{"c":{"B":"חוק","C1":"חוק","C":"חוק"},"l":127,"img":0},
"1540":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1541":{"c":{"C1":"ספציפי"},"l":127,"img":0},
"1542":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"1543":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1544":{"c":{"C1":"ספציפי"},"l":127,"img":0},
"1545":{"c":{"C1":"ספציפי"},"l":127,"img":0},
"1546":{"c":{"B":"חוק","C1":"חוק","C":"חוק"},"l":127,"img":0},
"1547":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1548":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1549":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1550":{"c":{"B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"1551":{"c":{"C1":"ספציפי"},"l":127,"img":0},
"1552":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1553":{"c":{"B":"חוק","C1":"חוק","C":"חוק","D":"ספציפי"},"l":127,"img":0},
"1554":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"1555":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"1556":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"1557":{"c":{"C1":"ספציפי"},"l":127,"img":0},
"1558":{"c":{"B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"1559":{"c":{"B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"1560":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1561":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":0},
"1562":{"c":{"B":"הכרת הרכב","C1":"הכרת הרכב"},"l":127,"img":0},
"1563":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"1564":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"1565":{"c":{"C1":"ספציפי","D":"ספציפי"},"l":127,"img":0},
"1566":{"c":{"C1":"ספציפי","D":"ספציפי"},"l":127,"img":0},
"1567":{"c":{"C1":"ספציפי","D":"ספציפי"},"l":127,"img":0},
"1568":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1569":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"1570":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"1572":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"1573":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1574":{"c":{"1":"ספציפי"},"l":127,"img":0},
"1578":{"c":{"1":"הכרת הרכב","B":"הכרת הרכב","C1":"הכרת הרכב"},"l":127,"img":0},
"1580":{"c":{"B":"בטיחות","C1":"בטיחות"},"l":127,"img":0},
"1582":{"c":{"B":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1590":{"c":{"1":"ספציפי"},"l":127,"img":0},
"1591":{"c":{"B":"חוק"},"l":127,"img":0},
"1593":{"c":{"B":"חוק"},"l":127,"img":0},
"1598":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"1607":{"c":{"D":"ספציפי"},"l":127,"img":1},
"1608":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"1622":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1625":{"c":{"1":"ספציפי"},"l":127,"img":0},
"1628":{"c":{"1":"ספציפי"},"l":127,"img":0},
"1636":{"c":{"C1":"ספציפי"},"l":127,"img":0},
"1637":{"c":{"C1":"ספציפי"},"l":127,"img":1},
"1638":{"c":{"D":"ספציפי"},"l":127,"img":1},
"1641":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1646":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"1650":{"c":{"1":"חוק","B":"חוק","C1":"חוק"},"l":127,"img":0},
"1652":{"c":{"C1":"ספציפי"},"l":127,"img":1},
"1664":{"c":{"1":"ספציפי"},"l":127,"img":0},
"1665":{"c":{"1":"ספציפי"},"l":127,"img":0},
"1671":{"c":{"C":"ספציפי"},"l":127,"img":1},
"1672":{"c":{"C":"ספציפי"},"l":127,"img":1},
"1673":{"c":{"1":"הכרת הרכב","B":"הכרת הרכב","C1":"הכרת הרכב","C":"הכרת הרכב","D":"הכרת הרכב"},"l":127,"img":0},
"1674":{"c":{"B":"חוק","C1":"חוק"},"l":127,"img":0},
"1675":{"c":{"1":"הכרת הרכב","B":"הכרת הרכב","C1":"הכרת הרכב","C":"הכרת הרכב","D":"הכרת הרכב"},"l":127,"img":0},
"1676":{"c":{"C1":"ספציפי","D":"ספציפי"},"l":127,"img":0},
"1677":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1678":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1679":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"1681":{"c":{"B":"חוק","C1":"חוק"},"l":127,"img":0},
"1682":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1683":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1684":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1685":{"c":{"B":"חוק","C1":"חוק"},"l":127,"img":0},
"1686":{"c":{"1":"חוק","B":"חוק","C1":"חוק"},"l":127,"img":0},
"1687":{"c":{"B":"חוק","C1":"חוק","C":"חוק","D":"ספציפי"},"l":127,"img":0},
"1688":{"c":{"1":"חוק","B":"חוק","C1":"חוק"},"l":127,"img":1},
"1690":{"c":{"B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"1691":{"c":{"B":"הכרת הרכב","C1":"הכרת הרכב","C":"הכרת הרכב","D":"הכרת הרכב"},"l":127,"img":0},
"1692":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1693":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"1694":{"c":{"1":"הכרת הרכב","B":"הכרת הרכב","C1":"הכרת הרכב"},"l":127,"img":0},
"1695":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"1696":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1697":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"1698":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1699":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"1700":{"c":{"B":"חוק","C1":"חוק","C":"חוק","D":"ספציפי"},"l":127,"img":0},
"1701":{"c":{"B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"1702":{"c":{"1":"הכרת הרכב","B":"הכרת הרכב","C1":"הכרת הרכב","C":"הכרת הרכב","D":"הכרת הרכב"},"l":127,"img":0},
"1703":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"1704":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"1705":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"1706":{"c":{"1":"הכרת הרכב","B":"הכרת הרכב","C1":"הכרת הרכב","C":"הכרת הרכב","D":"הכרת הרכב"},"l":127,"img":0},
"1707":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1708":{"c":{"1":"הכרת הרכב","B":"הכרת הרכב","C1":"הכרת הרכב"},"l":127,"img":1},
"1709":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"1710":{"c":{"1":"ספציפי"},"l":127,"img":0},
"1711":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"1712":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1713":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1714":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1715":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"1716":{"c":{"1":"חוק","B":"חוק","C1":"חוק"},"l":127,"img":0},
"1717":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1718":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1719":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"1720":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1721":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"1722":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1723":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":1},
"1724":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1725":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1726":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1727":{"c":{"B":"חוק","C1":"חוק"},"l":127,"img":1},
"1729":{"c":{"D":"ספציפי"},"l":127,"img":1},
"1731":{"c":{"D":"ספציפי"},"l":127,"img":1},
"1736":{"c":{"1":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1737":{"c":{"1":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1738":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"1739":{"c":{"1":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1740":{"c":{"B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1741":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1742":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1743":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1744":{"c":{"B":"הכרת הרכב","C1":"הכרת הרכב"},"l":127,"img":0},
"1745":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1746":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1747":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1748":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1749":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1750":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":0},
"1751":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1752":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1753":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1754":{"c":{"B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"1755":{"c":{"B":"חוק","C1":"חוק"},"l":127,"img":0},
"1756":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1757":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1758":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1759":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1760":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"1761":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":0},
"1762":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1763":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":0},
"1764":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1765":{"c":{"B":"בטיחות","C1":"בטיחות"},"l":127,"img":0},
"1766":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1767":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1768":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1769":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"1770":{"c":{"C":"ספציפי"},"l":127,"img":0},
"1771":{"c":{"C":"ספציפי"},"l":127,"img":0},
"1772":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":0},
"1773":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"1774":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1775":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1776":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":0},
"1777":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1778":{"c":{"1":"הכרת הרכב","B":"הכרת הרכב","C1":"הכרת הרכב","C":"הכרת הרכב","D":"הכרת הרכב"},"l":127,"img":0},
"1779":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1780":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1781":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1782":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1783":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1784":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1785":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1786":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1787":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":0},
"1788":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1789":{"c":{"B":"הכרת הרכב","C1":"הכרת הרכב"},"l":127,"img":0},
"1790":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"1791":{"c":{"1":"בטיחות","B":"בטיחות","C1":"בטיחות","C":"בטיחות","D":"בטיחות"},"l":127,"img":0},
"1792":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1793":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1794":{"c":{"D":"ספציפי"},"l":127,"img":0},
"1795":{"c":{"1":"ספציפי"},"l":127,"img":0},
"1797":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1798":{"c":{"1":"חוק","B":"חוק","C1":"חוק","C":"חוק","D":"חוק"},"l":127,"img":0},
"1799":{"c":{"C":"ספציפי"},"l":127,"img":1},
"1800":{"c":{"C":"ספציפי"},"l":127,"img":0},
"1801":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1},
"1803":{"c":{"1":"תמרורים","B":"תמרורים","C1":"תמרורים","C":"תמרורים","D":"תמרורים"},"l":127,"img":1}
};

// Bit 0 = he … bit 6 = am in QUESTION_INDEX[id].l — the order tools/build_bank.js
// writes. A bit says the id exists in that language's bank, which is also what
// makes its answer key trustworthy for that language (en/fr/es/ar order their
// answers differently from Hebrew — see TRANSLATION_LINEAGE_2026-09-20.md).
var QUESTION_LANGS = ['he', 'ru', 'en', 'ar', 'fr', 'es', 'am'];
// Every question in every one of the seven generated banks has exactly four
// answers (checked over all 6,245-6,270 rows per bank, 21/09/2026), so the
// displayed order is a permutation of [0..3].
var QUESTION_ANSWER_COUNT = 4;

var EXAM_STRUCTURE_SERVER = {
  'B':  { 'בטיחות': 7, 'הכרת הרכב': 7, 'חוק': 7, 'תמרורים': 9 },
  '1':  { 'בטיחות': 5, 'הכרת הרכב': 5, 'חוק': 6, 'תמרורים': 6, 'ספציפי': 8 },
  'C1': { 'בטיחות': 5, 'הכרת הרכב': 5, 'חוק': 5, 'תמרורים': 5, 'ספציפי': 10 },
  'C':  { 'בטיחות': 5, 'הכרת הרכב': 4, 'חוק': 3, 'תמרורים': 4, 'ספציפי': 14 },
  'D':  { 'בטיחות': 4, 'הכרת הרכב': 2, 'חוק': 5, 'תמרורים': 4, 'ספציפי': 15 }
};

// The bank's raw category → one of the five blueprint topics. Kept on the server
// because the index stores the classified topic and the reports classify the
// categories they read out of result rows.
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

function shuffleArrayServer(arr) {
  var a = arr.slice();
  for (var i = a.length - 1; i > 0; i--) {
    var j = Math.floor(Math.random() * (i + 1));
    var t = a[i]; a[i] = a[j]; a[j] = t;
  }
  return a;
}

function questionIndexEntry(id) {
  var entry = QUESTION_INDEX[String(id)];
  return entry || null;
}

function questionIndexCount() { return Object.keys(QUESTION_INDEX).length; }

function questionLangBit(lang) {
  var i = QUESTION_LANGS.indexOf(String(lang || 'he').toLowerCase());
  return i < 0 ? 0 : (1 << i);
}

// The blueprint topic of a question for one license ('' when the question does
// not belong to that license at all).
function questionTopic(id, license) {
  var entry = questionIndexEntry(id);
  return (entry && entry.c[String(license)]) || '';
}

// The one place that reads the answer key. null means "the key cannot answer for
// this id in this language" — callers must treat that as NOT VERIFIABLE and
// never as index 0, which is how the false 0/30 of 03/06/2026 happened.
function answerKeyIndex(id, lang) {
  if (typeof lookupCorrectIndex !== 'function') return null;
  var idx = lookupCorrectIndex(Number(id), String(lang || 'he').toLowerCase());
  if (idx === null || idx === undefined) return null;
  var n = Number(idx);
  return (isFinite(n) && n >= 0 && n < QUESTION_ANSWER_COUNT) ? n : null;
}

// ids of one license+language grouped by blueprint topic. One pass over 1,700
// index entries — measured in microseconds, so no cache (and no cache bug).
function indexIdsByTopic(license, lang) {
  var bit = questionLangBit(lang), lic = String(license), byTopic = {};
  for (var id in QUESTION_INDEX) {
    if (!Object.prototype.hasOwnProperty.call(QUESTION_INDEX, id)) continue;
    var entry = QUESTION_INDEX[id];
    if (!(entry.l & bit)) continue;
    var topic = entry.c[lic];
    if (!topic) continue;
    if (!byTopic[topic]) byTopic[topic] = [];
    byTopic[topic].push(Number(id));
  }
  return byTopic;
}

// Every id available for a license+language, whatever its topic.
function indexIdsFor(license, lang) {
  var byTopic = indexIdsByTopic(license, lang), out = [];
  for (var topic in byTopic) {
    if (!Object.prototype.hasOwnProperty.call(byTopic, topic)) continue;
    out = out.concat(byTopic[topic]);
  }
  return out;
}

function questionBankUnavailable(message, detail) {
  var err = new Error(message);
  err.code = 'bank_unavailable';
  err.detail = detail || '';
  return err;
}

// 30 ids per the license blueprint, as [{id, topic}] in random order.
// A question whose answer key is missing for THIS language is skipped here:
// registering it would mean scoring it later against a key that does not exist.
function drawExamIds(license, lang) {
  var blueprint = EXAM_STRUCTURE_SERVER[String(license)];
  if (!blueprint) throw questionBankUnavailable('דרגה לא מוכרת: ' + license, String(license));
  var byTopic = indexIdsByTopic(license, lang), picked = [], used = {};
  for (var topic in blueprint) {
    if (!Object.prototype.hasOwnProperty.call(blueprint, topic)) continue;
    var need = blueprint[topic], pool = shuffleArrayServer(byTopic[topic] || []), got = 0;
    for (var i = 0; i < pool.length && got < need; i++) {
      var id = pool[i];
      if (used[id] || answerKeyIndex(id, lang) === null) continue;
      used[id] = true;
      picked.push({ id: id, topic: topic });
      got++;
    }
    if (got < need) {
      throw questionBankUnavailable('אין מספיק שאלות בנושא ' + topic, topic + ' ' + got + '/' + need);
    }
  }
  return shuffleArrayServer(picked);
}

// A fresh display order for one question: a permutation of the answer positions.
function drawShuffleOrder() {
  var order = [];
  for (var i = 0; i < QUESTION_ANSWER_COUNT; i++) order.push(i);
  return shuffleArrayServer(order);
}

// Practice scores locally, so it needs the correct index for every language the
// question exists in, XOR-encoded exactly as the legacy questions.js did
// (ci = correctIndex ^ (id % 256)) so the client keeps a single decoder.
// Languages the question is not translated into are omitted — the answer key
// falls back to Hebrew, and for en/fr/es/ar that fallback is a different order.
function practiceCiByLang(id) {
  var entry = questionIndexEntry(id);
  if (!entry) return null;
  var out = {};
  for (var i = 0; i < QUESTION_LANGS.length; i++) {
    if (!(entry.l & (1 << i))) continue;
    var idx = answerKeyIndex(id, QUESTION_LANGS[i]);
    if (idx === null) continue;
    out[QUESTION_LANGS[i]] = idx ^ (Number(id) % 256);
  }
  return out;
}
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
  var expected = '';
  try { expected = String(PropertiesService.getScriptProperties().getProperty('GATEWAY_KEY') || ''); } catch (e) { expected = ''; }
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
// the client compares it against the static bank it loaded.
function handleHealth(p) {
  var body = { status: 'ok', build: THEORY_API_BUILD, indexIds: questionIndexCount() };
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

// A remembered login is verified on every page load and every reload — and on
// 16/09 the reload storm made that a full read of 'בוחנים' per reload per
// device. The POSITIVE verdict is cached exactly as verifyToken's is (review C
// R13): same 60 s, same key shape, so a disabled account or a rotated token
// costs at most one minute of grace.
var LOGIN_VERDICT_CACHE_SEC = 60;
function handleVerifyLogin(p) {
  if (!p.examinerId || !p.token) {
    return jsonResponse({ status: 'error', message: 'חסרים פרטי אימות', tokenExpired: true });
  }
  var vKey = CACHE_KEY_PREFIX + 'vlog_' + normalizeId(p.examinerId) + '_' + String(p.token).slice(0, 80), vCache = null;
  try {
    vCache = CacheService.getScriptCache();
    var vHit = vCache.get(vKey);
    if (vHit) return jsonResponse({ status: 'ok', examiner: JSON.parse(vHit) });
  } catch (eGet) { vCache = null; }
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
      var examiner = { name: data[i][0], id: normalizeId(data[i][1]), examinerNumber: String(data[i][4] || ''), role: String(data[i][5] || 'בוחן'), token: p.token };
      try { if (!vCache) vCache = CacheService.getScriptCache(); vCache.put(vKey, JSON.stringify(examiner), LOGIN_VERDICT_CACHE_SEC); } catch (ePut) {}
      return jsonResponse({ status: 'ok', examiner: examiner });
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
  // Date-bounded and archive-aware: the range defaults to today, and a range
  // older than the 30-day retention window must still see the archive (B5).
  diagMark('sheet:results-center-report');
  var centerRead = readResultsSince(dateFrom);
  var rows = centerRead.rows;
  diagMark('sheet:results-center-report-done:' + centerRead.mode);
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
  diagMark('sheet:sessions-list');
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

  diagMark('sheet:sessions-report');
  var sessData = sessionRows();

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
  diagMark('sheet:role-report');
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

  // r11 read the tail and fell back to the WHOLE sheet whenever the reported
  // day was older than the tail — the very first line the 'אבחון' sheet ever
  // recorded was this report at 81.5 s (2026-09-15), and the 360 s doGet kills
  // cluster at end-of-exam report time. r25: one date-bounded read that spans
  // the live sheet and the archive, so an old day costs the rows of that day
  // instead of every result ever recorded.
  diagMark('sheet:results-report');
  var resRead = readResultsSince(dayStart);
  var resData = resRead.rows;
  diagMark('sheet:results-report-done:' + resRead.mode);
  var results = [];
  for (var r = 1; r < resData.length; r++) {
    var sCode = String(resData[r][13] || '').trim();
    if (!sessionCodesSet[sCode]) continue;
    if (String(resData[r][7] || '').trim() === 'בוטל') continue; // skip overturned/superseded rows (consistent with the other report handlers)
    // parseSheetDateTime, not new Date(): column A is written as the STRING
    // "DD/MM/YYYY HH:MM" and only becomes a real date when the spreadsheet's
    // locale parses it. Where it does not (a text-formatted column, an imported
    // row), new Date() returned Invalid Date and toISOString() threw — taking
    // the whole report down with a 500 instead of one odd row.
    var rDate = parseSheetDateTime(resData[r][0]);
    results.push({
      date: rDate ? rDate.toISOString() : String(resData[r][0] || ''),
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
  diagMark('sheet:examiners-list');   // 17/09: 196s for this one small read
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

// Address of the polling Worker (DESIGN §3.4). Empty = examinees poll this
// script directly; setting/clearing the ScriptProperty switches the whole fleet
// within one getSessionInfo, without a Pages deploy.
function gatewayUrl() {
  try { return String(PropertiesService.getScriptProperties().getProperty('GATEWAY_URL') || '').trim(); }
  catch (e) { return ''; }
}

// ---- One 'סשנים' read per execution ----------------------------------------
// addExamTime and disqualify each read the whole sheet twice — once for the
// ownership check, once for the session's examiner name (review C R12). The
// memo lives for one request, which is far shorter than any state it caches.
var _sessionRowsMemo = null;
function sessionRows() {
  if (!_sessionRowsMemo) _sessionRowsMemo = getSheet('סשנים').getDataRange().getValues();
  return _sessionRowsMemo;
}
function sessionRowByCode(sessionCode) {
  var rows = sessionRows(), want = String(sessionCode || '').trim();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === want) return rows[i];
  }
  return null;
}
// Same rule as verifyExaminerForSession (20_auth.js): the session's own
// examiner, and nothing when the session does not exist. Served from the memo
// so the caller's later lookups are free. ⚠ The two must stay in step until the
// auth module is rewritten to take a row.
function examinerOwnsSession(sessionCode, examinerId) {
  if (!examinerId) return false;
  var row = sessionRowByCode(sessionCode);
  return !!row && normalizeId(row[1]) === normalizeId(examinerId);
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

  // A session closes at the end of its day; results older than the retention
  // window cannot belong to it, so the tail is enough.
  var resData = readResultsTail().rows;
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
      setPendingStatus(pendSheet, stuck[k].rowIdx + 1, sessionCode, 'cancelled');
      cancelled++;
    } else if (latest === 'בוטל') {
      setPendingStatus(pendSheet, stuck[k].rowIdx + 1, sessionCode, 'completed');
      completed++;
    } else {
      skipped++;
    }
  }
  return { cancelled: cancelled, completed: completed, skipped: skipped };
}

function handleGetSessionInfo(p) {
  var sheet = getSheet('סשנים');
  var data = sessionRows();
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
          // The client checks `build` to notice an old server behind a new page,
          // and reads `gateway.url` to decide where the examinee polls. An empty
          // url (ScriptProperty GATEWAY_URL unset) means "poll me directly" —
          // that is the kill switch for the Worker, with no Pages push.
          build: THEORY_API_BUILD,
          gateway: { url: gatewayUrl() },
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
  invalidatePendingSnapshot(p.sessionCode);   // r23: the first poll must find the new row
  return jsonResponse({ status: 'ok', examineeToken: examineeToken });
}

// Write columns of a ממתינים row WITHOUT changing its status (audio, extension,
// DQ counter, warnings, "finished on device"). Goes through the same flush and
// the same snapshot invalidation as setPendingStatus: a poller that reads a
// snapshot written before the change would otherwise show the old value for up
// to PENDING_SNAPSHOT_SEC (review C R8).
function writePendingCells(sheet, rowNumber, sessionCode, extras) {
  var wrote = false;
  for (var name in extras) {
    if (!Object.prototype.hasOwnProperty.call(extras, name) || !PENDING_COLS[name]) continue;
    sheet.getRange(rowNumber, PENDING_COLS[name]).setValue(extras[name]);
    wrote = true;
  }
  if (!wrote) return;
  SpreadsheetApp.flush();
  invalidatePendingSnapshot(sessionCode);
}

function handleCancelRegistration(p) {
  var sheet = getSheet('ממתינים');
  var data = sheet.getDataRange().getValues();
  var hit = findLatestPendingRow(data, p.sessionCode, p.idNumber, ['waiting', 'approved']);
  if (hit.idx === -1) return jsonResponse({ status: 'error', message: 'לא נמצא רישום פעיל לביטול' });
  // Verify phone matches to prevent unauthorized cancellation
  var storedPhone = String(hit.row[3] || '').replace(/[^0-9]/g, '');
  var givenPhone = String(p.phone || '').replace(/[^0-9]/g, '');
  if (storedPhone && givenPhone && storedPhone.slice(-7) !== givenPhone.slice(-7)) {
    return jsonResponse({ status: 'error', message: 'פרטים לא תואמים' });
  }
  setPendingStatus(sheet, hit.idx + 1, p.sessionCode, 'cancelled');
  return jsonResponse({ status: 'ok' });
}

function handleCheckApproval(p) {
  // Rate limit: max 60 polls per minute per (sessionCode, idNumber). Normal
  // polling is ~20-30/min, so this gives 2× headroom while blocking floods.
  var rlErr = requireRateLimit('checkApproval', String(p.sessionCode || '') + '_' + normalizeId(p.idNumber), 60, 60);
  if (rlErr) return rlErr;
  var BASE_EXAM_MINUTES = 40;
  // r23: served from the per-session snapshot (pendingRowsForSession); a row
  // missing from a cached snapshot is re-read from the sheet before "not found".
  var snap = pendingRowsForSession(p.sessionCode);
  var found = scanApprovalRows(snap.rows, p, BASE_EXAM_MINUTES);
  if (!found && snap.cached) found = scanApprovalRows(pendingRowsForSession(p.sessionCode, true).rows, p, BASE_EXAM_MINUTES);
  return found || jsonResponse({ status: 'error', message: 'לא נמצא רישום' });
}

// The scan handleCheckApproval used to run inline — unchanged; null = no active row.
function scanApprovalRows(data, p, BASE_EXAM_MINUTES) {
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
      // Per-examinee audio (column J). Returned on EVERY poll so the examinee's
      // client stays in sync with what the examiner set on their row — the
      // client used to freeze the session-level flag at code-entry time and had
      // no refresh path at all.
      response.audioMode = String(data[i][9] || '').trim() === 'on' ? 'on' : 'off';
      // When approved, compute and return authorized exam duration
      if (approval === 'approved' || approval === 'in_exam') {
        var ext = parseFloat(data[i][10]) || 1;
        if (ext !== 1.25 && ext !== 1.5) ext = 1;
        response.examMinutes = Math.round(BASE_EXAM_MINUTES * ext);
      }
      return jsonResponse(response);
    }
  }
  return null;
}

function handleApproveExaminee(p) {
  if (p.examinerId && !verifyExaminerForSession(p.sessionCode, p.examinerId)) {
    return jsonResponse({ status: 'error', message: 'אין הרשאה — בוחן לא תואם לסשן' });
  }
  // Validate time extension (whitelist)
  var validExt = { '': true, '1.25': true, '1.5': true };
  var timeExt = String(p.timeExtension || '');
  if (!validExt[timeExt]) timeExt = '';

  // Per-examinee audio. The examiner decides this on the specific examinee's
  // row, so it no longer depends on the session-level flag being on at the
  // moment the examinee typed the session code (that snapshot was the bug:
  // audio turned on after the examinee registered never reached them).
  // Omitted param → leave column J as the examinee registered with it.
  var audioMode = String(p.audioMode || '');
  if (audioMode !== 'on' && audioMode !== 'off') audioMode = '';

  var sheet = getSheet('ממתינים');
  var data = sheet.getDataRange().getValues();
  var hit = findLatestPendingRow(data, p.sessionCode, p.idNumber, ['waiting']);
  if (hit.idx === -1) {
    var current = data.length > 1 ? (findLatestPendingRow(data, p.sessionCode, p.idNumber).status || 'לא נמצא') : 'אין נתונים';
    return jsonResponse({ status: 'error', message: 'נבחן ממתין לא נמצא (סטטוס נוכחי: ' + current + ')' });
  }
  var extras = {};
  if (timeExt) extras.timeExtension = timeExt;   // column K
  if (audioMode) extras.audio = audioMode;       // column J
  setPendingStatus(sheet, hit.idx + 1, p.sessionCode, 'approved', extras);
  return jsonResponse({ status: 'ok' });
}

function handleRejectExaminee(p) {
  if (p.examinerId && !verifyExaminerForSession(p.sessionCode, p.examinerId)) {
    return jsonResponse({ status: 'error', message: 'אין הרשאה — בוחן לא תואם לסשן' });
  }
  var sheet = getSheet('ממתינים');
  var data = sheet.getDataRange().getValues();
  var hit = findLatestPendingRow(data, p.sessionCode, p.idNumber, ['waiting']);
  if (hit.idx === -1) return jsonResponse({ status: 'error', message: 'נבחן ממתין לא נמצא' });
  setPendingStatus(sheet, hit.idx + 1, p.sessionCode, 'rejected');
  return jsonResponse({ status: 'ok' });
}

// handleMarkExamStarted is gone (review C R14): startExam performs the
// approved → in_exam flip itself, through the single status writer, and the
// action now answers handleMarkExamStartedNoop (60_exam.js) for the one release
// in which an old client may still call it.

function handleExaminerDashboard(p) {
  var code = String(p.sessionCode);
  var pendSheet = getSheet('ממתינים');
  var resSheet = getSheet('תוצאות');

  // Tail reads (see readTail). pendOff shifts the one row-index write below;
  // resSheet is only ever appended to in this handler, so it needs no offset.
  // r15: instrumented. This is THE exam hot path — every examiner polls it every
  // 2-5s for the whole exam — and on 2026-09-15 it recorded 27.9s with no marks
  // at all, so the trail could not say where the time went. It re-reads the
  // results tail three times per request (here, after a state change, and before
  // the completed list); the marks will finally price that.
  diagMark('sheet:pending-dash');
  var _pendT = readTail(pendSheet, 4);
  var pendData = _pendT.rows, pendOff = _pendT.off;
  diagMark('sheet:results-dash');
  var _resT = readTail(resSheet, 0);
  var resData = _resT.rows;
  diagMark('sheet:extensions-dash');
  var pending = [];
  var active = [];

  // Sum of mid-exam time grants per examinee (minutes) — extends the stale/timeout
  // threshold below and is shown as a badge in the active list. One read, by id.
  var extraMinById = {};
  try { extraMinById = extraMinutesBySession(code); } catch (e) {}   // r23: cached 30 s, dropped by addExamTime

  // ---- Pre-built indexes (perf) ----------------------------------------------
  // handleExaminerDashboard runs every 5s per examiner. The old code re-scanned
  // the WHOLE תוצאות/ממתינים sheets inside per-examinee loops, making it
  // O(examinees × תוצאות) — which silently degraded as תוצאות grew each day and
  // spiked during the morning registration rush. These indexes reproduce EXACTLY
  // what those inner scans computed, but build once → O(1) lookups.
  //   resBySessId[id]  = { dqResults, otherResults }  (this session, non-בוטל)
  //   pendTermBySessId[id] = { dqTerminals, otherTerminals } (this session)
  function buildResBySessId(rows) {
    var idx = {};
    for (var r = 1; r < rows.length; r++) {
      if (String(rows[r][13]) !== code) continue;
      if (String(rows[r][7] || '') === 'בוטל') continue;
      var k = normalizeId(rows[r][1]);
      if (!idx[k]) idx[k] = { dqResults: 0, otherResults: 0 };
      if (String(rows[r][7] || '').trim() === 'פסול') idx[k].dqResults++;
      else idx[k].otherResults++;
    }
    return idx;
  }
  function buildPendTermBySessId(rows) {
    var idx = {};
    for (var r = 1; r < rows.length; r++) {
      if (String(rows[r][0]) !== code) continue;
      var k = normalizeId(rows[r][1]);
      if (!idx[k]) idx[k] = { dqTerminals: 0, otherTerminals: 0 };
      var st = String(rows[r][5]).trim();
      if (st === 'disqualified' || st === 'dq_confirmed') idx[k].dqTerminals++;
      else if (st === 'completed') idx[k].otherTerminals++;
    }
    return idx;
  }
  var resBySessId = buildResBySessId(resData);
  var pendTermBySessId = buildPendTermBySessId(pendData);

  // ---- Auto-cleanup of stale in_exam/approved rows ---------------------------
  // Bounded on purpose (review C R1 / fix A1). This loop used to do, INSIDE
  // itself and per stale row: a full 'סשנים' read, a full 'תוצאות' read for the
  // attempt number, an append, another results tail read and an index rebuild —
  // 6 round trips and ~169,000 cells each. It fires hardest right after an
  // outage, when every examinee who could not submit is stale at once: 40 of
  // them measured 247 round trips / 6.86 M cells in ONE 2-5 s poll, which is the
  // best match at HEAD for the end-of-exam 360 s doGet kills.
  // Now: at most DASH_MAX_RECONCILE_PER_POLL rows per request (the rest are
  // reconciled by the next poll, 2-5 s later), the session row is read at most
  // once, the attempt history at most once, and nothing is re-read after an
  // append — the appended row is added to the in-memory table instead.
  var DASH_MAX_RECONCILE_PER_POLL = 3;
  var reconciled = 0;
  var sessionRowForDash = null, sessionRowRead = false;
  function dashSessionRow() {
    if (!sessionRowRead) { sessionRowRead = true; diagMark('sheet:sessions-dash'); sessionRowForDash = sessionRowByCode(code); }
    return sessionRowForDash;
  }
  var attemptHistory = null;
  function dashAttemptCount(idNumber, license) {
    if (!attemptHistory) { diagMark('sheet:attempts-dash'); attemptHistory = readAttemptHistory(); }
    return countAttemptRows(attemptHistory, normalizeId(idNumber), String(license));
  }
  var now = new Date();
  var BASE_EXAM_MS = 40 * 60 * 1000;
  var STALE_BUFFER_MS = 20 * 60 * 1000; // 20 minutes buffer (approval wait + instructions)
  for (var ci = 1; ci < pendData.length; ci++) {
    if (reconciled >= DASH_MAX_RECONCILE_PER_POLL) break;
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

    // Count results by type for this examinee in this session (indexed lookup —
    // was a full scan of תוצאות per examinee).
    var _rc = resBySessId[normalizeId(ciId)] || { dqResults: 0, otherResults: 0 };
    var dqResults = _rc.dqResults, otherResults = _rc.otherResults;
    // Count terminal entries by type in pending sheet for this examinee (indexed —
    // was a full scan of ממתינים per examinee).
    var _pt = pendTermBySessId[normalizeId(ciId)] || { dqTerminals: 0, otherTerminals: 0 };
    var dqTerminals = _pt.dqTerminals, otherTerminals = _pt.otherTerminals;
    // Cap DQ results to DQ terminals — handles duplicate פסול rows from page refreshes
    var effectiveResults = Math.min(dqResults, dqTerminals) + otherResults;
    var totalTerminals = dqTerminals + otherTerminals;
    var hasUnmatchedResult = effectiveResults > totalTerminals;

    if (hasUnmatchedResult || effectiveStale) {
      reconciled++;
      // Fix dangling status — mark as completed (single status writer: the
      // snapshot the examinee's poller reads is dropped in the same call).
      setPendingStatus(pendSheet, ci + 1 + pendOff, code, 'completed');
      pendData[ci][5] = 'completed'; // update local copy
      // Keep pendTermBySessId in sync: this row was in_exam/approved (loop guard
      // above) → now a 'completed' terminal, so a fresh rescan would count it here.
      var _mk = normalizeId(ciId);
      if (!pendTermBySessId[_mk]) pendTermBySessId[_mk] = { dqTerminals: 0, otherTerminals: 0 };
      pendTermBySessId[_mk].otherTerminals++;
      if (effectiveStale && !hasUnmatchedResult) {
        // Create a timeout fail result
        var ses2 = dashSessionRow();
        var license2 = pendData[ci][8] || '', site2 = '', classroom2 = '', examinerName2 = '', language2 = pendData[ci][6] || 'he';
        if (ses2) {
          examinerName2 = ses2[2] || '';
          site2 = ses2[3] || '';
          classroom2 = ses2[4] || '';
          if (!license2) license2 = ses2[5] || '';
        }
        var failRow = [
          todayStr(), ciId, pendData[ci][2] || '', pendData[ci][3] || '', license2,
          '0/30', '0%', 'נכשל', '', examinerName2,
          site2, classroom2, language2, code,
          dashAttemptCount(ciId, license2) + 1, 'ניתוק/טיימאאוט — הנבחן לא סיים את המבחן', false, false, '',
          pendData[ci][7] || '', false, pendData[ci][9] || 'off'
        ];
        resSheet.appendRow(failRow);
        // The row we just wrote is the only thing a re-read would have added, so
        // add it in memory: later iterations, the completed list and the
        // attempts-today tally all see it without another read of 'תוצאות'.
        resData.push(failRow);
        if (attemptHistory) attemptHistory.push([failRow[0], failRow[1], '', '', failRow[4], '', '', failRow[7]]);
        if (!resBySessId[_mk]) resBySessId[_mk] = { dqResults: 0, otherResults: 0 };
        resBySessId[_mk].otherResults++;
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

  // r25: the second tail read of 'תוצאות' is GONE. On a normal poll it
  // re-fetched identical data (1,000 × 30 cells, 12 times a minute per
  // examiner); its only other effect was picking up a result another execution
  // appended during this request, and that arrives one poll later anyway — the
  // dashboard polls every 2 s while a result is syncing.
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
  var todayDD = ('0' + now.getDate()).slice(-2);
  var todayMM = ('0' + (now.getMonth() + 1)).slice(-2);
  var todayYYYY = now.getFullYear();
  var todayDate = todayDD + '/' + todayMM + '/' + todayYYYY; // "DD/MM/YYYY"
  // Index today's non-בוטל results by examinee id (any session), then attach —
  // was a full scan of תוצאות per pending examinee.
  var todayExamsById = {};
  for (var ti = 1; ti < resData.length; ti++) {
    if (String(resData[ti][7] || '') === 'בוטל') continue;
    // Handle both Date objects and string dates from Sheets
    var _cd = resData[ti][0];
    var _ds = '';
    if (_cd instanceof Date) {
      _ds = ('0' + _cd.getDate()).slice(-2) + '/' + ('0' + (_cd.getMonth() + 1)).slice(-2) + '/' + _cd.getFullYear();
    } else {
      _ds = String(_cd);
    }
    if (_ds.indexOf(todayDate) !== 0) continue;
    var _tk = normalizeId(resData[ti][1]);
    (todayExamsById[_tk] = todayExamsById[_tk] || []).push({ license: String(resData[ti][4]), score: String(resData[ti][5]), passed: String(resData[ti][7]), language: String(resData[ti][12] || '') });
  }
  for (var pi = 0; pi < pending.length; pi++) {
    var _te = todayExamsById[normalizeId(pending[pi].idNumber)];
    if (_te && _te.length > 0) pending[pi].todayExams = _te;
  }

  // Cross-reference registration times from ממתינים for completed results (indexed
  // — was a reverse scan of ממתינים per completed examinee). Last matching row for
  // (code, id) wins, exactly as the reverse-from-end + break did.
  var pendRegTimeById = {};
  for (var pr = 1; pr < pendData.length; pr++) {
    if (String(pendData[pr][0]) !== code) continue;
    pendRegTimeById[normalizeId(pendData[pr][1])] = pendData[pr][4];
  }
  for (var c = 0; c < completed.length; c++) {
    var _rk = normalizeId(completed[c].idNumber);
    if (Object.prototype.hasOwnProperty.call(pendRegTimeById, _rk)) {
      completed[c].registrationTime = pendRegTimeById[_rk];
    }
  }

  diagMark('compute:dash-done');
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
    var hit = findLatestPendingRow(data, p.sessionCode, p.idNumber);
    if (hit.idx !== -1 && (hit.status === 'in_exam' || hit.status === 'approved')) {
      var prev = (hit.row.length > 15) ? (Number(hit.row[15]) || 0) : 0;
      var extras = { warnCount: prev + 1 };
      if (p.reason) extras.lastWarning = String(p.reason).slice(0, 40);
      writePendingCells(sheet, hit.idx + 1, p.sessionCode, extras);
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
  // 17/09/2026: one of these polls ran 354s and was killed, others 58-93s, with
  // nothing here but a 1000-row tail read and a small extensions read. Marks
  // (free under 8s) so the next stall says WHERE — spreadsheet, cache or before.
  diagMark('sheet:pending-status');
  // r23: per-session snapshot, see pendingRowsForSession; a missing row is re-read
  var snap = pendingRowsForSession(p.sessionCode);
  var found = scanExamStatusRows(snap.rows, p);
  if (!found && snap.cached) found = scanExamStatusRows(pendingRowsForSession(p.sessionCode, true).rows, p);
  return found || jsonResponse({ status: 'ok', examStatus: 'not_found' });
}

function scanExamStatusRows(data, p) {
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]).trim() === String(p.sessionCode).trim() && normalizeId(data[i][1]) === normalizeId(p.idNumber)) {
      var storedToken = String((data[i].length > 12 ? data[i][12] : '') || '').trim();
      if (storedToken && p.examineeToken && String(p.examineeToken).trim() !== storedToken) {
        return jsonResponse({ status: 'error', examineeTokenError: 'mismatch' });
      }
      return jsonResponse({ status: 'ok', examStatus: String(data[i][5] || '').trim(), extraMinutes: sumExtraMinutes(p.sessionCode, p.idNumber) });
    }
  }
  return null;
}

// ===== Mid-exam time addition (security evacuation / technical / medical) =====
// The examiner grants extra minutes to a RUNNING exam. Every grant is appended to
// the 'הארכות זמן' audit sheet with a mandatory reason, so the record is preserved.
// getExamStatus + the dashboard read the SUM of grants per examinee:
//   - the examinee extends examDeadline (idempotently: start + base + sum)
//   - the dashboard pushes back the stale/timeout-fail threshold by the same sum
// Examiner-authenticated only (mirrors handleDisqualify path A).
// r23: the grants of a session are read once per EXTRA_MINUTES_CACHE_SEC and
// served to every getExamStatus poll and every dashboard poll from the cache;
// handleAddExamTime drops the entry, so a new grant is visible at once.
var EXTRA_MINUTES_CACHE_SEC = 30;
function extraMinutesKey(sessionCode) { return CACHE_KEY_PREFIX + 'extmin_' + String(sessionCode || '').trim(); }
// { normalizedId: minutes } for one session
function extraMinutesBySession(sessionCode) {
  var key = extraMinutesKey(sessionCode), cache = null;
  try { cache = CacheService.getScriptCache(); var hit = cache.get(key); if (hit) return JSON.parse(hit); } catch (eGet) { cache = null; }
  diagMark('sheet:extensions');
  var d = getSheet('הארכות זמן').getDataRange().getValues(), map = {}, want = String(sessionCode || '').trim();
  for (var i = 1; i < d.length; i++) {
    if (String(d[i][1]).trim() !== want) continue;
    var k = normalizeId(d[i][2]);
    map[k] = (map[k] || 0) + (Number(d[i][4]) || 0);
  }
  try { if (!cache) cache = CacheService.getScriptCache(); cache.put(key, JSON.stringify(map), EXTRA_MINUTES_CACHE_SEC); } catch (ePut) {}
  return map;
}
function invalidateExtraMinutes(sessionCode) {
  try { CacheService.getScriptCache().remove(extraMinutesKey(sessionCode)); } catch (e) {}
}
function sumExtraMinutes(sessionCode, idNumber) {
  try { return extraMinutesBySession(sessionCode)[normalizeId(idNumber)] || 0; } catch (e) { return 0; }
}

function handleAddExamTime(p) {
  if (!p.sessionCode || !p.idNumber) return jsonResponse({ status: 'error', message: 'חסר מזהה' });
  // Examiner auth — must hold a valid token AND own the session (same as DQ).
  // examinerOwnsSession serves the check from the per-execution 'סשנים' memo
  // the examiner-name lookup below reuses: this handler read that sheet twice
  // (review C R12).
  if (!verifyToken(p.examinerId, p.token)) {
    return jsonResponse({ status: 'error', message: 'טוקן בוחן לא תקין', tokenExpired: true });
  }
  if (!examinerOwnsSession(p.sessionCode, p.examinerId)) {
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
  var hit = findLatestPendingRow(pendData, p.sessionCode, p.idNumber);
  if (hit.idx === -1) return jsonResponse({ status: 'error', message: 'נבחן לא נמצא בסשן' });
  var name = hit.row[2] || '';

  // Examiner display name for the audit row — from the same memo as the auth check.
  var sessionRow = sessionRowByCode(p.sessionCode);
  var examinerName = sessionRow ? (sessionRow[2] || '') : '';

  getSheet('הארכות זמן').appendRow([new Date(), p.sessionCode, p.idNumber, name, minutes, reason, examinerName]);
  invalidateExtraMinutes(p.sessionCode);   // r23: the next status poll must see the grant

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
  var hit = findLatestPendingRow(data, p.sessionCode, p.idNumber);
  if (hit.idx === -1) return jsonResponse({ status: 'ok' });   // no matching row — harmless no-op
  var storedToken = String((hit.row.length > 12 ? hit.row[12] : '') || '').trim();
  if (storedToken && p.examineeToken && String(p.examineeToken).trim() !== storedToken) {
    return jsonResponse({ status: 'error', examineeTokenError: 'mismatch' });
  }
  if (hit.status === 'in_exam') {
    // Older sheets stop at 18 columns (SHEET_HEADERS now declares 19).
    if (pendSheet.getMaxColumns() < 19) pendSheet.insertColumnsAfter(pendSheet.getMaxColumns(), 19 - pendSheet.getMaxColumns());
    if (!String(pendSheet.getRange(1, 19).getValue() || '').trim()) pendSheet.getRange(1, 19).setValue('סיים במכשיר');
    writePendingCells(pendSheet, hit.idx + 1, p.sessionCode, { finishedOnDevice: nowISO() });
  }
  return jsonResponse({ status: 'ok' });
}

// ---- One upstream read for the whole session (gateway, DESIGN §3.4) --------
// The examinee pollers are 87% of an exam morning's requests: 40 phones × 12
// polls/min = 480 Apps Script executions a minute, each one a container start
// against ~30 slots. The Worker collapses them into ONE upstream call per
// session every 3 s and answers the phones itself, so this is the only shape in
// which examinee state leaves the script.
// It carries NO names and NO phones, and never the examinee token itself: the
// Worker compares SHA-256 hashes, so a leak of this response cannot be replayed
// as an examinee. Rows come back in sheet order (oldest first).
defineAction('sessionSnapshot', { methods: ['GET'], auth: 'gateway', handler: handleSessionSnapshot,
  rateLimit: { max: 60, windowSec: 60, id: function(p) { return String(p.sessionCode || ''); } } });
function handleSessionSnapshot(p) {
  var code = String(p.sessionCode || '').trim();
  if (!code) return jsonResponse({ status: 'error', message: 'חסר קוד סשן' });
  var snap = pendingRowsForSession(code);         // the same 4-second snapshot the pollers use
  var extraMin = {};
  try { extraMin = extraMinutesBySession(code); } catch (eExt) { extraMin = {}; }
  var rows = [];
  for (var i = 1; i < snap.rows.length; i++) {
    var r = snap.rows[i], id = normalizeId(r[1]);
    rows.push({
      id: id,
      status: String(r[5] || '').trim(),
      tokenHash: hashExamineeToken(r.length > 12 ? r[12] : ''),
      audio: String(r[9] || '').trim() === 'on' ? 'on' : 'off',
      examMinutes: examMinutesFor(r),   // one rule for the exam length (60_exam.js)
      extraMinutes: extraMin[id] || 0
    });
  }
  return jsonResponse({ status: 'ok', at: Date.now(), rows: rows });
}

function hashExamineeToken(token) {
  var t = String(token || '').trim();
  if (!t) return '';
  try {
    var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, t, Utilities.Charset.UTF_8), hex = '';
    for (var i = 0; i < bytes.length; i++) {
      var b = (bytes[i] + 256) % 256;
      hex += (b < 16 ? '0' : '') + b.toString(16);
    }
    return hex;
  } catch (e) { return ''; }
}

function handleDisqualify(p) {
  // Auth: two valid paths
  //   A) Examiner-initiated DQ — must include valid token AND own the session
  //   B) Self-DQ from examinee client (cheat detection) — pending row must exist with active status
  // Without one of these, reject. Prevents an attacker with just sessionCode+victim's idNumber
  // from disqualifying other examinees.
  var pendSheet = getSheet('ממתינים');
  var pendData = pendSheet.getDataRange().getValues();
  var hit = findLatestPendingRow(pendData, p.sessionCode, p.idNumber);
  var name = '', phone = '', population = '', examineeLicense = '', examineeAudio = 'off';
  var pendRowIdx = hit.idx, pendStatus = hit.status;
  if (hit.idx !== -1) {
    name = hit.row[2] || '';
    phone = hit.row[3] || '';
    population = hit.row[7] || '';
    examineeLicense = hit.row[8] || '';
    examineeAudio = hit.row[9] || 'off';
  }

  if (p.examinerId) {
    // Path A: examiner-initiated — require valid token + ownership.
    // examinerOwnsSession reads 'סשנים' through the per-execution memo the
    // result row below reuses; this handler read the sheet twice (C R12).
    if (!verifyToken(p.examinerId, p.token)) {
      return jsonResponse({ status: 'error', message: 'טוקן בוחן לא תקין', tokenExpired: true });
    }
    if (!examinerOwnsSession(p.sessionCode, p.examinerId)) {
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
    var prevCount = (pendData[pendRowIdx].length > 13) ? (Number(pendData[pendRowIdx][13]) || 0) : 0;
    setPendingStatus(pendSheet, pendRowIdx + 1, p.sessionCode, 'disqualified', { dqCount: prevCount + 1 });
    // Clear any OTHER active (in_exam/approved) rows for this examinee so a
    // duplicate row doesn't linger on the board beside the disqualified one.
    for (var dqd = 1; dqd < pendData.length; dqd++) {
      if (dqd === pendRowIdx) continue;
      if (String(pendData[dqd][0]) !== String(p.sessionCode) || normalizeId(pendData[dqd][1]) !== normalizeId(p.idNumber)) continue;
      var dqdStatus = String(pendData[dqd][5]).trim();
      if (dqdStatus === 'in_exam' || dqdStatus === 'approved') {
        setPendingStatus(pendSheet, dqd + 1, p.sessionCode, 'cancelled');
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
  // The dedupe window is 2 minutes and a retry lands within seconds, so the tail
  // is always enough — this used to read every result ever recorded.
  var data = readResultsTail().rows;
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
  var sesRow = sessionRowByCode(p.sessionCode);
  var license = '', language = 'he', site = '', classroom = '', examinerName = '';
  if (sesRow) {
    examinerName = sesRow[2] || '';
    site = sesRow[3] || '';
    classroom = sesRow[4] || '';
    license = examineeLicense || sesRow[5] || '';
    language = sesRow[6] || 'he';
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
  var pendHit = findLatestPendingRow(pendData, sc, id);
  if (pendHit.idx !== -1 && pendHit.status === 'disqualified') {
    setPendingStatus(pendSheet, pendHit.idx + 1, sc, 'in_exam');
  }

  // 2. Cancel the DQ result row matching this dqEventId (or the latest פסול).
  // The row was written seconds ago, inside the grace period — tail is enough.
  var dqEventId = String(p.dqEventId || '');
  var resSheet = getSheet('תוצאות');
  var resRead = readResultsTail();
  var resHit = findLatestResultRow(resRead.rows, sc, id, false);
  if (resHit.idx !== -1 && resHit.status === 'פסול' &&
      (!dqEventId || String(resHit.row[24] || '') === dqEventId)) {
    resSheet.getRange(resHit.idx + 1 + resRead.off, 8).setValue('בוטל');
    SpreadsheetApp.flush();
  }
  return jsonResponse({ status: 'ok' });
}

function handleResetExaminee(p) {
  if (p.examinerId && !examinerOwnsSession(p.sessionCode, p.examinerId)) {
    return jsonResponse({ status: 'error', message: 'אין הרשאה — בוחן לא תואם לסשן' });
  }
  var sheet = getSheet('ממתינים');
  var data = sheet.getDataRange().getValues();
  // Reset EVERY non-final row for this examinee (not just the latest) and accept
  // ALL stuck states — including 'disqualified'/'dq_confirmed'. Previously reset
  // refused those, so a soldier stuck on a pending DQ could not be cleared at all.
  // "אפס" should fully remove a stuck soldier from the board so they can re-register.
  var RESETTABLE = { waiting: 1, approved: 1, in_exam: 1, disqualified: 1, dq_confirmed: 1 };
  var resetCount = 0;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) !== String(p.sessionCode) || normalizeId(data[i][1]) !== normalizeId(p.idNumber)) continue;
    if (!RESETTABLE[String(data[i][5]).trim()]) continue;
    setPendingStatus(sheet, i + 1, p.sessionCode, 'cancelled');
    resetCount++;
  }
  if (resetCount === 0) {
    return jsonResponse({ status: 'error', message: 'לא נמצא נבחן פעיל לאיפוס' });
  }
  return jsonResponse({ status: 'ok', resetCount: resetCount });
}

// Force-complete a stuck in_exam examinee (examiner manual action)
function handleForceComplete(p) {
  if (p.examinerId && !examinerOwnsSession(p.sessionCode, p.examinerId)) {
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
    setPendingStatus(pendSheet, j + 1, p.sessionCode, 'completed');
    found = true;
  }
  if (!found) {
    return jsonResponse({ status: 'error', message: 'לא נמצא נבחן עם סטטוס in_exam/approved' });
  }

  // Check if result already exists — if so, just mark pending as completed (done
  // above). The examinee is in THIS session, so their row is in the tail.
  var resSheet = getSheet('תוצאות');
  var resData = readResultsTail().rows;
  if (findLatestResultRow(resData, p.sessionCode, p.idNumber, false).idx !== -1) {
    return jsonResponse({ status: 'ok', message: 'נמצאה תוצאה קיימת — הסטטוס עודכן' });
  }

  // No result exists — create a fail result
  var sesRow = sessionRowByCode(p.sessionCode);
  var license = examineeLicense, site = '', classroom = '', examinerName = '';
  if (sesRow) {
    examinerName = sesRow[2] || '';
    site = sesRow[3] || '';
    classroom = sesRow[4] || '';
    if (!license) license = sesRow[5] || '';
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
  if (p.examinerId && !examinerOwnsSession(p.sessionCode, p.examinerId)) {
    return jsonResponse({ status: 'error', message: 'אין הרשאה — בוחן לא תואם לסשן' });
  }

  // Find the latest result row + the pending row for this examinee in one pass
  // each. An overturn always targets a result of the running session.
  var sheet = getSheet('תוצאות');
  var resRead = readResultsTail();
  var resHit = findLatestResultRow(resRead.rows, p.sessionCode, p.idNumber, false);
  var resultRowIdx = resHit.idx, resultStatus = resHit.status;

  var pendSheet = getSheet('ממתינים');
  var pendData = pendSheet.getDataRange().getValues();
  var pendHit = findLatestPendingRow(pendData, p.sessionCode, p.idNumber);
  var pendRowIdx = pendHit.idx, pendStatusNow = pendHit.status;

  // Case 1: latest result is פסול → normal overturn flow.
  // Pending revert covers BOTH 'disqualified' (auto-DQ, not yet confirmed) and
  // 'dq_confirmed' (examiner already clicked ✔ אשר). Without the dq_confirmed
  // branch, the examiner who pressed "אשר" by accident could overturn the
  // result row but the examinee stays locked out — they'd need a fresh
  // registration, which is what created duplicate rows at base 14.
  if (resultStatus === 'פסול') {
    sheet.getRange(resultRowIdx + 1 + resRead.off, 8).setValue('בוטל');
    sheet.getRange(resultRowIdx + 1 + resRead.off, 18).setValue(false);
    SpreadsheetApp.flush();
    if (pendRowIdx !== -1 && (pendStatusNow === 'disqualified' || pendStatusNow === 'dq_confirmed')) {
      setPendingStatus(pendSheet, pendRowIdx + 1, p.sessionCode, 'in_exam');
    }
    return jsonResponse({ status: 'ok' });
  }

  // Case 2: stuck pending in 'disqualified' but latest result is already
  // a final outcome (עבר/נכשל/בוטל). Happens when DQ fired transiently
  // during a deploy window — examinee continued and finished the exam, but
  // the pending row stayed stuck. Just clean up the pending row.
  if (pendRowIdx !== -1 && pendStatusNow === 'disqualified' &&
      (resultStatus === 'עבר' || resultStatus === 'נכשל' || resultStatus === 'בוטל')) {
    setPendingStatus(pendSheet, pendRowIdx + 1, p.sessionCode, 'completed');
    return jsonResponse({ status: 'ok', resolved: 'stale_dq_cleared' });
  }

  // Case 3: pending is disqualified but no result row yet → revert so the
  // examinee can resume the exam (in_exam state, just like case 1).
  if (pendRowIdx !== -1 && pendStatusNow === 'disqualified' && resultRowIdx === -1) {
    setPendingStatus(pendSheet, pendRowIdx + 1, p.sessionCode, 'in_exam');
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
  if (!examinerOwnsSession(p.sessionCode, p.examinerId)) {
    return jsonResponse({ status: 'error', message: 'אין הרשאה — בוחן לא תואם לסשן' });
  }
  // Mark pending status as dq_confirmed so examinee polling gets a final answer
  var pendSheet = getSheet('ממתינים');
  var pendData = pendSheet.getDataRange().getValues();
  var hit = findLatestPendingRow(pendData, p.sessionCode, p.idNumber);
  if (hit.idx === -1 || hit.status !== 'disqualified') {
    return jsonResponse({ status: 'error', message: 'לא נמצא רישום פסול לאישור' });
  }
  setPendingStatus(pendSheet, hit.idx + 1, p.sessionCode, 'dq_confirmed');
  return jsonResponse({ status: 'ok' });
}

function handleCorrectToPass(p) {
  if (p.examinerId && !examinerOwnsSession(p.sessionCode, p.examinerId)) {
    return jsonResponse({ status: 'error', message: 'אין הרשאה — בוחן לא תואם לסשן' });
  }
  var sheet = getSheet('תוצאות');
  var read = readResultsTail();
  // skipCancelled: a 'בוטל' row was already overturned/superseded — correcting
  // it would resurrect a dead row and leave two live results (review E S7).
  var hit = findLatestResultRow(read.rows, p.sessionCode, p.idNumber, true);
  if (hit.idx === -1) return jsonResponse({ status: 'error', message: 'תוצאה לא נמצאה' });
  var row = hit.row, rowNumber = hit.idx + 1 + read.off;
  // Verify score is eligible (>= 24/30)
  var scoreNum = parseInt(String(row[5]).split('/')[0]) || 0;
  if (scoreNum < 24) {
    return jsonResponse({ status: 'error', message: 'ציון נמוך מדי לתיקון (מתחת ל-24)' });
  }
  sheet.getRange(rowNumber, 8).setValue('עבר');      // column H = עבר/נכשל
  sheet.getRange(rowNumber, 18).setValue(false);     // column R = disqualified (may be a DQ row)
  sheet.getRange(rowNumber, 21).setValue(true);      // column U = תוקן?
  // Regenerate WhatsApp link — corrected result shows only "עבר" (no score/errors)
  var waMsg = '*🚗 אישור תוצאת מבחן תאוריה חיצוני*\n\n' +
    'שם: ' + row[2] + '\n' +
    'ת.ז.: ' + row[1] + '\n' +
    'דרגה: ' + row[4] + '\n' +
    (row[19] ? 'אוכלוסיה: ' + row[19] + '\n' : '') +
    'תאריך: ' + row[0] + '\n' +
    'תוצאה: *עבר*\n';
  sheet.getRange(rowNumber, 19).setValue('https://wa.me/' + formatPhoneForWA(row[3]) + '?text=' + encodeURIComponent(waMsg));
  SpreadsheetApp.flush();
  return jsonResponse({ status: 'ok' });
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
  var sesRow = sessionRowByCode(p.sessionCode);
  if (sesRow) {
    examinerName = sesRow[2] || '';
    site = sesRow[3] || '';
    classroom = sesRow[4] || '';
    sessLicense = sesRow[5] || '';
    sessLanguage = sesRow[6] || 'he';
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
  // row already exists for this session+id+license+score. The retry follows
  // within seconds, so the tail covers it.
  var manExisting = readResultsTail().rows;
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
  if (!examinerOwnsSession(p.sessionCode, p.examinerId)) {
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
  var metaRead = readResultsTail();   // the examiner fixes a row of the session in front of them
  var rows = metaRead.rows;
  for (var i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][13]) === String(p.sessionCode) && normalizeId(rows[i][1]) === normalizeId(p.idNumber)) {
      var rowIdx = i + 1 + metaRead.off;
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

  // A commander corrects results of ANY session, including one from weeks ago,
  // so this is the one correction handler that keeps the full live read (it runs
  // a handful of times a month and must not miss a row a tail would cut off).
  var sheet = getSheet('תוצאות');
  var hit = findLatestResultRow(sheet.getDataRange().getValues(), data.sessionCode, data.idNumber, true);   // skip 'בוטל' (E S7)
  if (hit.idx === -1) return jsonResponse({ status: 'error', message: 'תוצאה לא נמצאה' });
  var rowIdx = hit.idx + 1;
  var pct = Math.round((newScore / newTotal) * 100);
  sheet.getRange(rowIdx, 6).setValue(newScore + '/' + newTotal);  // F: ציון
  sheet.getRange(rowIdx, 7).setValue(pct + '%');                   // G: אחוז
  sheet.getRange(rowIdx, 8).setValue(newStatus);                   // H: עבר/נכשל
  sheet.getRange(rowIdx, 18).setValue(newStatus === 'פסול');       // R: פסול?
  sheet.getRange(rowIdx, 21).setValue(true);                       // U: תוקן?
  // Audit trail (columns Z=26, AA=27, AB=28) — commander's display name from בוחנים
  var commanderName = '';
  try {
    var examData = getSheet('בוחנים').getDataRange().getValues();
    for (var x = 1; x < examData.length; x++) {
      if (normalizeId(examData[x][1]) === normalizeId(data.examinerId)) { commanderName = String(examData[x][0] || ''); break; }
    }
  } catch(e) {}
  sheet.getRange(rowIdx, 26).setValue(commanderName + ' (' + normalizeId(data.examinerId) + ')');
  sheet.getRange(rowIdx, 27).setValue(reason);
  sheet.getRange(rowIdx, 28).setValue(todayStr());
  SpreadsheetApp.flush();
  return jsonResponse({ status: 'ok' });
}

function handleMarkSent(p) {
  // Ownership check — consistent with the other examiner mutations; prevents an
  // authenticated examiner from flipping the "נשלח?" flag on another session's rows.
  if (!examinerOwnsSession(p.sessionCode, p.examinerId)) {
    return jsonResponse({ status: 'error', message: 'אין הרשאה — בוחן לא תואם לסשן' });
  }
  var sheet = getSheet('תוצאות');
  var read = readResultsTail();   // rows of the session the examiner is sending from
  var data = read.rows;
  var wanted = {};
  var ids = p.idNumbers ? p.idNumbers.split(',') : [p.idNumber];
  for (var k = 0; k < ids.length; k++) wanted[normalizeId(ids[k])] = true;
  var count = 0;
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][13]) !== String(p.sessionCode)) continue;
    if (!wanted[normalizeId(data[i][1])]) continue;
    sheet.getRange(i + 1 + read.off, 17).setValue(true);  // נשלח? — column Q (17)
    count++;
  }
  return jsonResponse({ status: 'ok', updated: count });
}

// ========== Exam start, practice draw, result submission ====================
//
// One call starts an exam (startExam) and one call ends it (submitResult). The
// server draws the questions from QUESTION_INDEX, stores the map it drew, and
// scores the submission against the answer key — the client is trusted for
// nothing but WHAT IT DISPLAYED (question and answer texts, for the certificate).
//
// What the old flow did instead: getExamQuestions (Drive/pool read, ~10s) →
// registerExamQuestions (full read of 'ממתינים', no idempotency, four client
// retries) → markExamStarted (another full read) → submitResult (full read of
// 'מבחנים', 8.2MB, plus three full reads of 'תוצאות' and a Drive read for the
// wrong-answer texts). That is the 15MB submit and the 10-second start.

// ---- The question map of one exam ------------------------------------------
// Stored in 'מבחנים' exactly as before — six columns, one JSON map per row —
// so nothing downstream changes; the map entries gained `topic` (the blueprint
// bucket, for the certificate) since the server can no longer look a category up
// in a bank it does not have.
var EXAM_BASE_MINUTES = 40;
var EXAM_MIN_QUESTIONS = 25;               // review E S1: never register or score a short map
var EXAM_MAP_CACHE_SEC = 10800;            // 3h — longer than any exam plus its extensions
var EXAM_MAP_MAX_AGE_MS = 8 * 3600 * 1000; // a session code lives 8h; an older row is a previous attempt
var EXAM_SUSPICIOUS_SEC = 180;             // a "finished" exam faster than this is flagged, not blocked

// The cache key carries the ATTEMPT, not just the person: an examinee who was
// reset and registered again gets a new 'ממתינים' row with a new registration
// time and must be drawn a fresh exam, not handed the cached one.
function examAttemptKey(row) {
  var stamp = (row && row[4] instanceof Date) ? row[4].toISOString() : String((row && row[4]) || '');
  return stamp.replace(/[^0-9A-Za-z]/g, '').slice(-14) || 'na';
}
function examMapCacheKey(sessionCode, idNumber, attemptKey) {
  return CACHE_KEY_PREFIX + 'qmap_' + String(sessionCode || '').trim() + '_' + normalizeId(idNumber) + '_' + attemptKey;
}
function readExamMapCache(sessionCode, idNumber, attemptKey) {
  try {
    var hit = CacheService.getScriptCache().get(examMapCacheKey(sessionCode, idNumber, attemptKey));
    if (!hit) return null;
    var record = JSON.parse(hit);
    return (record && record.map && record.map.length) ? record : null;
  } catch (e) { return null; }
}
function writeExamMapCache(sessionCode, idNumber, attemptKey, record) {
  try {
    CacheService.getScriptCache().put(examMapCacheKey(sessionCode, idNumber, attemptKey), JSON.stringify(record), EXAM_MAP_CACHE_SEC);
  } catch (e) { /* the sheet is the source of truth; the cache only saves a read */ }
}

// 'מבחנים' carries one JSON map per row and grows forever — reading it whole to
// use a single row was 8.2MB of the 15MB submit. Columns A-B (session, id) are
// scanned bottom-up and only the matching row's C-F is read.
// maxAgeMs bounds the search to the current session; 0 accepts any age (a submit
// must find its own registration however long the exam ran).
// Returns null when there is no row, or { map: null } when the row cannot be
// parsed — "a registration exists but is unreadable" is not "no registration".
function readExamRegistration(sessionCode, idNumber, maxAgeMs) {
  var sheet = getSheet('מבחנים'), lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  diagMark('sheet:exam-keys');
  var keys = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  for (var i = keys.length - 1; i >= 0; i--) {
    if (String(keys[i][0]) !== String(sessionCode)) continue;
    if (normalizeId(keys[i][1]) !== normalizeId(idNumber)) continue;
    diagMark('sheet:exam-row');
    var record = parseExamRegistrationRow(sheet.getRange(i + 2, 3, 1, 4).getValues()[0]);
    if (maxAgeMs && examRegistrationAgeMs(record) > maxAgeMs) return null;
    return record;
  }
  return null;
}
function parseExamRegistrationRow(cells) {
  var record = {
    map: null,
    at: (cells[1] instanceof Date) ? cells[1].toISOString() : String(cells[1] || ''),
    lang: String(cells[2] || ''),
    unverified: Number(cells[3] || 0)
  };
  try {
    var map = JSON.parse(cells[0]);
    // An empty array IS a readable map — and a refusable one (review E S1).
    // Only unparseable JSON leaves map null, i.e. "registered but unreadable".
    if (Array.isArray(map)) record.map = map;
  } catch (e) { /* record.map stays null → the result is stored unverified */ }
  return record;
}
function examRegistrationAgeMs(record) {
  var at = record && record.at ? new Date(record.at) : null;
  if (!at || isNaN(at.getTime())) return 0;   // unparseable stamp: do not discard the map over it
  return Date.now() - at.getTime();
}

function appendExamRegistration(sessionCode, idNumber, map, at, lang) {
  diagMark('sheet:append-exam');
  var sheet = getSheet('מבחנים');
  // Every reader here skips row 1 as a header, so a sheet that somehow has none
  // would swallow its first exam.
  if (sheet.getLastRow() === 0 && SHEET_HEADERS['מבחנים']) sheet.appendRow(SHEET_HEADERS['מבחנים']);
  sheet.appendRow([String(sessionCode), normalizeId(idNumber), JSON.stringify(map), at, lang, 0]);
}

// ---- startExam --------------------------------------------------------------
// POST, examinee token. Replaces getExamQuestions + registerExamQuestions +
// markExamStarted with one idempotent call: the same (session, attempt) always
// gets the same 30 questions back, so a retry after a lost response resumes the
// exam instead of drawing a second one.
function handleStartExam(data) {
  var sessionCode = String(data.sessionCode || '').trim();
  if (!sessionCode || !data.idNumber) return jsonResponse({ status: 'error', message: 'חסרים פרטי נבחן' });
  var ctx = examineeRowContext(sessionCode, data.idNumber);
  if (!ctx.active) return jsonResponse({ status: 'error', code: 'not_approved', message: 'נבחן לא מאושר למבחן' });

  var row = ctx.active.row;
  var lang = String(data.language || row[6] || 'he').toLowerCase();
  var license = String(data.license || row[8] || 'B').trim();
  if (!EXAM_STRUCTURE_SERVER[license]) {
    return jsonResponse({ status: 'error', code: 'unknown_license', message: 'דרגה לא מוכרת: ' + license });
  }

  var attempt = examAttemptKey(row);
  var record = readExamMapCache(sessionCode, data.idNumber, attempt);
  // Only an exam already in progress may reuse a stored map. An 'approved' row
  // is a NEW attempt (a reset examinee re-registered), and it must be drawn
  // fresh even though a map of the previous attempt is still in the sheet.
  if (!record && ctx.active.status === 'in_exam') {
    record = readExamRegistration(sessionCode, data.idNumber, EXAM_MAP_MAX_AGE_MS);
    if (record && (!record.map || record.map.length < EXAM_MIN_QUESTIONS)) record = null;
  }
  if (!record) {
    try { record = drawExamRegistration(sessionCode, data.idNumber, license, lang); }
    catch (err) {
      if (!err || err.code !== 'bank_unavailable') throw err;
      return jsonResponse({ status: 'error', code: 'bank_unavailable', detail: err.detail,
        message: 'מאגר השאלות אינו זמין כעת — פנה לבוחן' });
    }
  }
  writeExamMapCache(sessionCode, data.idNumber, attempt, record);

  // approved → in_exam, through the single status writer (it flushes and drops
  // the poller snapshot). An in_exam row keeps its original start time.
  if (ctx.active.status === 'approved') {
    setPendingStatus(getSheet('ממתינים'), ctx.active.rowNumber, sessionCode, 'in_exam', { examStart: nowISO() });
    ctx.active.status = 'in_exam';
  }

  return jsonResponse({
    status: 'ok', build: THEORY_API_BUILD,
    examMinutes: examMinutesFor(row), extraMinutes: sumExtraMinutes(sessionCode, data.idNumber),
    audioMode: String(row[9] || '').trim() === 'on' ? 'on' : 'off',
    language: record.lang || lang, license: license, registeredAt: record.at,
    questions: examQuestionsForClient(record.map)
  });
}

function drawExamRegistration(sessionCode, idNumber, license, lang) {
  var drawn = drawExamIds(license, lang);
  if (drawn.length < EXAM_MIN_QUESTIONS) {
    throw questionBankUnavailable('נדרשות לפחות ' + EXAM_MIN_QUESTIONS + ' שאלות', String(drawn.length));
  }
  var map = [];
  for (var i = 0; i < drawn.length; i++) {
    var order = drawShuffleOrder();
    // Non-null by construction: drawExamIds skips every id the key cannot answer.
    var correct = answerKeyIndex(drawn[i].id, lang);
    map.push({ qIdx: i, qId: drawn[i].id, shuffleOrder: order, correctShuffledIdx: order.indexOf(correct), topic: drawn[i].topic });
  }
  var at = nowISO();
  appendExamRegistration(sessionCode, idNumber, map, at, lang);
  return { map: map, at: at, lang: lang, unverified: 0 };
}

// The client holds the texts (bank/<lang>.json) and needs only what the server
// decided: which questions, in which answer order, under which topic.
function examQuestionsForClient(map) {
  var out = [];
  for (var i = 0; i < map.length; i++) {
    out.push({ id: map[i].qId, order: map[i].shuffleOrder, topic: map[i].topic || '' });
  }
  return out;
}

// Column K of 'ממתינים' holds the examiner's time extension — same whitelist the
// approval poll applies, so both sides compute the same deadline.
function examMinutesFor(row) {
  var ext = parseFloat(row[10]) || 1;
  if (ext !== 1.25 && ext !== 1.5) ext = 1;
  return Math.round(EXAM_BASE_MINUTES * ext);
}

// ---- startPractice ----------------------------------------------------------
// GET, no token. Practice scores on the client, so it gets the correct index of
// every language the question exists in (XOR-encoded) and never needs another
// round trip — a language switch mid-practice is local.
var PRACTICE_MAX_COUNT = 50;
var PRACTICE_DEFAULT_COUNT = 15;
function handleStartPractice(p) {
  var rlErr = practiceRateLimit(p);
  if (rlErr) return rlErr;
  var lang = String(p.language || 'he').toLowerCase();
  var license = String(p.license || p.licenseType || 'B').trim();
  if (!EXAM_STRUCTURE_SERVER[license]) {
    return jsonResponse({ status: 'error', code: 'unknown_license', message: 'דרגה לא מוכרת: ' + license });
  }
  var mode = String(p.mode || 'exam');
  var picked;
  try { picked = practiceSelection(mode, license, lang, p); }
  catch (err) {
    if (!err || err.code !== 'bank_unavailable') throw err;
    return jsonResponse({ status: 'error', code: 'bank_unavailable', detail: err.detail, message: 'אין מספיק שאלות לתרגול' });
  }
  if (!picked.length) return jsonResponse({ status: 'error', code: 'no_questions', message: 'לא נמצאו שאלות לתרגול' });
  var questions = [];
  for (var i = 0; i < picked.length; i++) {
    questions.push({ id: picked[i].id, topic: picked[i].topic, ci: practiceCiByLang(picked[i].id) || {} });
  }
  return jsonResponse({ status: 'ok', mode: mode, count: questions.length, questions: questions });
}

function practiceSelection(mode, license, lang, p) {
  if (mode === 'ids') return practiceByIds(p.ids, license, lang);
  if (mode === 'category' && p.categoryFilter) return practiceByCategory(String(p.categoryFilter), license, lang, practiceCount(p));
  return drawExamIds(license, lang);   // full 30-question blueprint
}

function practiceCount(p) {
  var n = Number(p.maxCount) || PRACTICE_DEFAULT_COUNT;
  return Math.max(1, Math.min(PRACTICE_MAX_COUNT, n));
}

function practiceByCategory(topic, license, lang, maxCount) {
  var byTopic = indexIdsByTopic(license, lang), pool = shuffleArrayServer(byTopic[topic] || []), out = [];
  for (var i = 0; i < pool.length && out.length < maxCount; i++) out.push({ id: pool[i], topic: topic });
  return out;
}

// Spaced repetition: the client names the ids it wants back. Unknown ids and ids
// missing from this language are dropped rather than failing the whole request.
function practiceByIds(raw, license, lang) {
  var bit = questionLangBit(lang), out = [], seen = {};
  var parts = String(raw || '').split(',');
  for (var i = 0; i < parts.length && out.length < PRACTICE_MAX_COUNT; i++) {
    var id = parseInt(String(parts[i]).trim(), 10);
    if (isNaN(id) || seen[id]) continue;
    seen[id] = true;
    var entry = questionIndexEntry(id);
    if (!entry || !(entry.l & bit)) continue;
    out.push({ id: id, topic: entry.c[license] || '' });
  }
  return out;
}

// Class practice is identified by class+student, standalone by the ID typed into
// exam.html, and everything else is a guest. The guest allowance is also capped
// globally: `ci` makes the draw an answer oracle, so a scraper must not be able
// to walk the bank quickly by inventing identifiers.
var PRACTICE_GUEST_GLOBAL_MAX = 120;
function practiceRateLimit(p) {
  if (p.classCode && p.studentId) {
    return requireRateLimit('startPractice_student', String(p.classCode) + '_' + String(p.studentId), 20, 60);
  }
  if (p.standaloneIdNumber) {
    return requireRateLimit('startPractice_standalone', normalizeId(p.standaloneIdNumber), 5, 60);
  }
  return requireRateLimit('startPractice_guest', 'anon', 5, 60)
    || requireRateLimit('startPractice_guest', 'guest_global', PRACTICE_GUEST_GLOBAL_MAX, 60);
}

// ---- Retired actions --------------------------------------------------------
// One release of grace for a client that was loaded before the deploy: it asks
// for questions, gets a clear "refresh the page" instead of a broken exam.
function handleClientOutdated() {
  return jsonResponse({ status: 'error', code: 'client_outdated',
    message: 'גרסה חדשה של המערכת — יש לרענן את הדף (F5)' });
}

// startExam already flipped the row to in_exam; the old client's separate ping
// has nothing left to do. Kept (as a no-op) only so that client does not treat
// an unknown action as a failure. Remove in the next release.
function handleMarkExamStartedNoop() {
  return jsonResponse({ status: 'ok', already: true });
}

// ---- submitResult -----------------------------------------------------------
// Columns of 'תוצאות' by name — the handler used to index 30 positions by hand.
var RESULT_COL = { date: 0, id: 1, name: 2, phone: 3, license: 4, score: 5, percent: 6, verdict: 7,
  time: 8, examiner: 9, site: 10, classroom: 11, language: 12, session: 13, attempt: 14, wrongDetails: 15,
  sent: 16, dq: 17, waLink: 18, population: 19, corrected: 20, audio: 21, verified: 22, suspicious: 23,
  dqEventId: 24, correctedBy: 25, correctionReason: 26, correctionDate: 27, langPath: 28, device: 29 };
var RESULT_PASS_RATIO = 0.86;             // 26/30
var FABRICATED_FAIL_MARKERS = ['סגירת דפדפן', 'טיימאאוט', 'סיום ידני'];
var UNVERIFIED_PREFIX = '⚠️ ציון לא אומת';

function handleSubmitResult(data) {
  var gate = submitGate(data);
  if (gate.error) return gate.error;
  recordSubmitClientLog(data);

  var registration = readExamRegistration(data.sessionCode, data.idNumber, 0);
  var guard = submitRegistrationGuard(registration, data);
  if (guard) return guard;

  var scored = registration && registration.map && registration.map.length
    ? scoreRegisteredExam(registration.map, data.answers, registration.lang)
    : null;
  applyScore(data, scored, registration, !!(data.answers && data.answers.length));
  var wrongAnswers = scored ? wrongAnswerItems(scored, data.license) : [];

  // 'תוצאות' is read ONCE, as late as possible: supersede, duplicate, פסול and
  // idempotency all decide from the same snapshot. Three full reads of a sheet
  // that grows forever were most of what was left of the submit's cost.
  diagMark('sheet:results-submit');
  var sheet = getSheet('תוצאות');
  var tail = readTail(sheet, RESULT_COL.date);
  supersedeFabricatedFails(sheet, tail, data);

  var duplicate = findDuplicateResult(tail, data, gate);
  if (duplicate) {
    markPendingCompleted(data.sessionCode, data.idNumber, gate.pending);
    return jsonResponse({ status: 'ok', waLink: duplicate[RESULT_COL.waLink] || '', duplicate: true });
  }
  // Counted BEFORE the פסול rows are voided below: a disqualified exam was
  // still an attempt, and voiding it is only about not leaving two live rows.
  var attemptNum = countAttempts(data.idNumber, data.license, attemptRows(tail)) + 1;
  supersedeDisqualifications(sheet, tail, data);

  var waLink = buildResultWaLink(data, wrongAnswers, attemptNum);
  if (findIdenticalResult(tail, data)) {
    markPendingCompleted(data.sessionCode, data.idNumber, gate.pending);
    return jsonResponse({ status: 'ok', duplicate: true, waLink: waLink });
  }

  sheet.appendRow(buildResultRow(data, wrongAnswers, attemptNum, waLink));
  markPendingCompleted(data.sessionCode, data.idNumber, gate.pending);
  diagMark('compute:submit-done');
  return jsonResponse({ status: 'ok', waLink: waLink });
}

// Rate limit, examinee token and "is this person actually in this exam".
// Returns { error } or { pending, status } — the pending snapshot is handed on
// so the completion write does not read 'ממתינים' a second time.
function submitGate(data) {
  var rlErr = requireRateLimit('submitResult', String(data.sessionCode || '') + '_' + normalizeId(data.idNumber), 5, 60);
  if (rlErr) return { error: rlErr };
  var ctx = examineeRowContext(data.sessionCode, data.idNumber);
  var status = ctx.latest ? ctx.latest.status : '';
  // 'cancelled' is accepted: if an examiner reset an examinee who was in fact
  // still mid-exam, a genuine finished submit must be RECORDED, not lost. The
  // fabricated-fail supersede and the duplicate check keep the sheet clean.
  if (data.sessionCode && data.idNumber && ['in_exam', 'approved', 'completed', 'cancelled'].indexOf(status) === -1) {
    return { error: jsonResponse({ status: 'error', message: 'נבחן לא מאושר — לא ניתן לשלוח תוצאות' }) };
  }
  return { pending: pendingSnapshotFromTail(ctx), status: status, ctx: ctx };
}

// markPendingCompleted writes by absolute row index, so a tail read's rows are
// padded back to their sheet positions instead of paying for a second full read.
function pendingSnapshotFromTail(ctx) {
  var sheet = getSheet('ממתינים'), tail = ctx.tail;
  if (!tail || !tail.rows.length) return { sheet: sheet, rows: null };
  if (!tail.off) return { sheet: sheet, rows: tail.rows };
  var padded = [tail.rows[0]];
  for (var i = 0; i < tail.off; i++) padded.push([]);
  return { sheet: sheet, rows: padded.concat(tail.rows.slice(1)) };
}

// The client's last events (≤2KB) ride along with the submit; S2's recorder
// parks them next to the server-side diagnostics. Never fatal to a result.
function recordSubmitClientLog(data) {
  if (!data.clientLog) return;
  try {
    if (typeof diagRecordClientLog === 'function') diagRecordClientLog(data.sessionCode, data.idNumber, data.clientLog);
  } catch (e) { /* a diagnostic must never cost a result */ }
}

// A registered exam MUST come with answers (otherwise a forged score would skip
// the re-score entirely), and a map that is empty or shorter than a real exam is
// a data fault — refuse it loudly instead of scoring 3 questions out of 30 and
// possibly declaring a pass (review E S1).
function submitRegistrationGuard(registration, data) {
  if (!registration) return null;
  if (!data.answers || !Array.isArray(data.answers) || data.answers.length === 0) {
    return jsonResponse({ status: 'error', message: 'הגשה לא תקינה — חסרות תשובות למבחן רשום' });
  }
  if (registration.map && registration.map.length < EXAM_MIN_QUESTIONS) {
    return jsonResponse({ status: 'error', code: 'invalid_registration',
      message: 'רישום המבחן פגום — לא ניתן לנקד. פנה לבוחן.' });
  }
  return null;
}

// ---- scoring ---------------------------------------------------------------
// review E S2: only a real selection counts. null / '' / undefined / a
// non-number / a negative or fractional index all mean "not answered".
function submitAnswerIndex(answer) {
  if (!answer) return -1;
  var raw = answer.selected;
  if (raw === null || raw === undefined || raw === '') return -1;
  var n = Number(raw);
  if (!isFinite(n) || n < 0 || Math.floor(n) !== n) return -1;
  return n;
}

function answerLanguage(answer, registeredLang) {
  var lang = (answer && answer.langAtAnswer) ? String(answer.langAtAnswer) : String(registeredLang || 'he');
  return lang.toLowerCase() || 'he';
}

// review E S4: the correctShuffledIdx stored at registration is NOT read back.
// It is recomputed from the answer key every time, which is also what scores a
// mid-exam language switch against the key of the language the question was
// ANSWERED in (translators reorder answers — en/fr/es/ar have their own order).
// null = this entry cannot be verified at all, and then it can never be correct.
function correctIndexForEntry(entry, lang) {
  if (!entry || !entry.qId || !Array.isArray(entry.shuffleOrder)) return null;
  var orig = answerKeyIndex(entry.qId, lang);
  if (orig === null) return null;
  var pos = entry.shuffleOrder.indexOf(Number(orig));
  return pos >= 0 ? pos : null;
}

function scoreRegisteredExam(map, answers, registeredLang) {
  var scored = { correct: 0, total: map.length, verifiable: 0, items: [] };
  for (var i = 0; i < map.length; i++) {
    var entry = map[i] || null, answer = (answers && answers[i]) || null;
    var lang = answerLanguage(answer, registeredLang);
    var key = correctIndexForEntry(entry, lang);
    var selected = submitAnswerIndex(answer);
    var item = { qIdx: i, qId: entry ? entry.qId : null, topic: (entry && entry.topic) || '',
      lang: lang, key: key, selected: selected, answer: answer };
    item.right = (key !== null && selected === key);
    if (key !== null) scored.verifiable++;
    if (item.right) scored.correct++;
    scored.items.push(item);
  }
  scored.verified = scored.total > 0 && scored.verifiable === scored.total;
  return scored;
}

// The server's tally replaces whatever the client claimed. Without a readable
// registration nothing can be verified, so the row is stored with the unverified
// marker and the examiner reviews it — never silently trusted.
function applyScore(data, scored, registration, hasAnswers) {
  if (!scored) {
    data.verified = false;
    // Without answers there was nothing to verify in the first place (an
    // examiner-entered or legacy row) — only a real submission is flagged.
    if (hasAnswers) {
      data.scoreUnverified = true;
      data.unverifiedReason = registration ? 'רישום מבחן פגום' : 'רישום מבחן חסר';
    }
    // The new client sends no score at all (it never holds the key); an old one
    // sent its own tally. Either way an unverifiable result is stored as what it
    // is — a number the examiner must review — and never as a pass by default.
    data.total = Number(data.total) || (Array.isArray(data.answers) ? data.answers.length : 0) || 30;
    data.score = Math.max(0, Math.min(data.total, Number(data.score) || 0));
    data.percent = Number(data.percent) || Math.round((data.score / data.total) * 100);
    data.passed = data.passed === true || data.passed === 'true';
    return;
  }
  data.score = scored.correct;
  data.total = scored.total;
  data.percent = Math.round((scored.correct / scored.total) * 100);
  data.passed = scored.correct >= Math.ceil(scored.total * RESULT_PASS_RATIO);
  data.verified = scored.verified;
  if (!scored.verified) {
    data.scoreUnverified = true;
    data.unverifiedReason = 'שאלות ללא מפתח תשובות';
  }
  data.suspicious = registration && examRegistrationAgeMs(registration) > 0 &&
    examRegistrationAgeMs(registration) < EXAM_SUSPICIOUS_SEC * 1000;
}

// ---- feedback (certificate + WhatsApp) -------------------------------------
// Built from what the client DISPLAYED: q = the question text, a = the four
// answers in displayed order. The server owns which of them is correct. An old
// client that sends no texts still gets a scored result.
var ANSWER_LABELS_HE = ['א', 'ב', 'ג', 'ד', 'ה', 'ו'];
var ANSWER_LABELS_LATIN = ['A', 'B', 'C', 'D', 'E', 'F'];
var TEXT_UNAVAILABLE = '(טקסט לא זמין)';
function answerLabel(lang, idx) {
  var labels = (String(lang) === 'he') ? ANSWER_LABELS_HE : ANSWER_LABELS_LATIN;
  return labels[idx] || '';
}
function displayedAnswer(shown, idx, lang) {
  if (idx === null) return '(לא זמין כעת)';                 // no key for this language
  var count = (shown && shown.length) ? shown.length : QUESTION_ANSWER_COUNT;
  if (idx < 0 || idx >= count) return (String(lang) === 'he') ? 'לא נענתה' : 'Not answered';
  var text = shown ? String(shown[idx] || '') : '';
  return answerLabel(lang, idx) + ' - ' + (text || TEXT_UNAVAILABLE);
}

function wrongAnswerItems(scored, license) {
  var out = [];
  for (var i = 0; i < scored.items.length; i++) {
    var item = scored.items[i];
    if (item.right) continue;
    var shown = (item.answer && item.answer.a && item.answer.a.length) ? item.answer.a : null;
    out.push({
      questionId: item.qId || '',
      question: (item.answer && item.answer.q) ? String(item.answer.q) : TEXT_UNAVAILABLE,
      yourAnswer: displayedAnswer(shown, item.selected, item.lang),
      correctAnswer: displayedAnswer(shown, item.key, item.lang),
      category: item.topic || questionTopic(item.qId, license)
    });
  }
  return out;
}

// 'פירוט שגויות' (column P) — the commander dashboard aggregates by the מזהה
// שאלה line, and readers tolerate any line being missing.
function formatWrongDetails(items, data) {
  var text = '';
  for (var i = 0; i < items.length; i++) {
    var w = items[i];
    if (w.questionId) text += 'מזהה שאלה: ' + w.questionId + '\n';
    text += 'שאלה: ' + w.question + '\n';
    text += 'תשובת הנבחן: ' + w.yourAnswer + '\n';
    text += 'תשובה נכונה: ' + w.correctAnswer + '\n';
    if (w.category) text += 'קטגוריה: ' + w.category + '\n';
    text += '\n';
  }
  if (data.scoreUnverified) {
    text = UNVERIFIED_PREFIX + ' בשרת (' + (data.unverifiedReason || '') + ') — נדרש אימות ידני\n\n' + text;
  }
  return text;
}

function buildResultWaLink(data, items, attemptNum) {
  var message = '*🚗 אישור תוצאת מבחן תאוריה חיצוני*\n\n' +
    'שם: ' + data.fullName + '\n' +
    'ת.ז.: ' + data.idNumber + '\n' +
    'דרגה: ' + data.license + '\n' +
    (data.population ? 'אוכלוסיה: ' + data.population + '\n' : '') +
    'תאריך: ' + todayStr() + '\n' +
    'תוצאה: *' + resultVerdict(data) + '* (' + data.score + '/' + data.total + ')\n' +
    'זמן: ' + data.time + '\n';
  if (items.length > 0) {
    var wrongForWA = '';
    for (var i = 0; i < items.length; i++) {
      wrongForWA += '❌ ' + items[i].question + '\n' + 'ענית: ' + items[i].yourAnswer + '\n' + '✅ נכון: ' + items[i].correctAnswer + '\n\n';
    }
    message += '\n*שאלות שגויות (' + items.length + '):*\n\n' + wrongForWA;
  } else if (Number(data.total) - Number(data.score) === 0) {
    message += '\nכל התשובות נכונות! 🎉';
  }
  if (attemptNum > 1) message += 'ניסיון: ' + attemptNum + '\n';
  return 'https://wa.me/' + formatPhoneForWA(data.phone) + '?text=' + encodeURIComponent(message);
}

function resultVerdict(data) { return data.passed ? 'עבר' : 'נכשל'; }

// "he → ru → he" tells the examiner at a glance that the examinee switched.
function languagePath(data) {
  if (Array.isArray(data.languageHistory) && data.languageHistory.length > 0) {
    return data.languageHistory.length === 1 ? String(data.languageHistory[0]) : data.languageHistory.join(' → ');
  }
  return data.language || 'he';
}

function buildResultRow(data, wrongAnswers, attemptNum, waLink) {
  var row = [];
  row[RESULT_COL.date] = todayStr();
  row[RESULT_COL.id] = data.idNumber;
  row[RESULT_COL.name] = data.fullName;
  row[RESULT_COL.phone] = data.phone;
  row[RESULT_COL.license] = data.license;
  row[RESULT_COL.score] = data.score + '/' + data.total;
  row[RESULT_COL.percent] = data.percent + '%';
  row[RESULT_COL.verdict] = resultVerdict(data);
  row[RESULT_COL.time] = data.time;
  row[RESULT_COL.examiner] = data.examinerName || '';
  row[RESULT_COL.site] = data.site || '';
  row[RESULT_COL.classroom] = data.classroom || '';
  row[RESULT_COL.language] = data.language || 'he';
  row[RESULT_COL.session] = data.sessionCode || '';
  row[RESULT_COL.attempt] = attemptNum;
  row[RESULT_COL.wrongDetails] = formatWrongDetails(wrongAnswers, data);
  row[RESULT_COL.sent] = false;
  row[RESULT_COL.dq] = false;
  row[RESULT_COL.waLink] = waLink;
  row[RESULT_COL.population] = data.population || '';
  row[RESULT_COL.corrected] = false;
  row[RESULT_COL.audio] = data.audioMode || 'off';
  row[RESULT_COL.verified] = data.verified ? 'מאומת' : '';
  row[RESULT_COL.suspicious] = data.suspicious ? 'חשוד' : '';
  row[RESULT_COL.dqEventId] = '';
  row[RESULT_COL.correctedBy] = '';
  row[RESULT_COL.correctionReason] = '';
  row[RESULT_COL.correctionDate] = '';
  row[RESULT_COL.langPath] = languagePath(data);
  row[RESULT_COL.device] = String(data.device || '');
  return row;
}

// The attempt number counts a lifetime of results, so it may not be decided
// from a tail: an earlier attempt this month can sit above the last 1,000 rows.
// When the snapshot IS the whole sheet it is reused as it is; otherwise
// countAttempts reads the three columns it needs from live + archive itself.
function attemptRows(tail) { return tail.off ? null : tail.rows; }

// ---- the four passes over the one 'תוצאות' snapshot ------------------------
function resultRowMatchesExaminee(row, data) {
  return String(row[RESULT_COL.session]) === String(data.sessionCode) &&
    normalizeId(row[RESULT_COL.id]) === normalizeId(data.idNumber);
}

// A real finished submit must WIN over a system-written fail (the close beacon,
// the dashboard timeout row, an examiner's manual disconnect). Matched on
// session+id ONLY: the fabricated row carries the REGISTRATION language while a
// real submit may carry a different final one after a mid-exam switch.
function supersedeFabricatedFails(sheet, tail, data) {
  var rows = tail.rows, changed = false;
  for (var i = 1; i < rows.length; i++) {
    if (!resultRowMatchesExaminee(rows[i], data)) continue;
    if (String(rows[i][RESULT_COL.verdict]).trim() !== 'נכשל') continue;
    if (!isFabricatedFailNote(String(rows[i][RESULT_COL.wrongDetails] || ''))) continue;
    var sheetRow = i + tail.off + 1;
    sheet.getRange(sheetRow, RESULT_COL.verdict + 1).setValue('בוטל');
    sheet.getRange(sheetRow, RESULT_COL.correctionReason + 1).setValue('בוטל אוטומטית — הנבחן השלים והגיש מבחן');
    sheet.getRange(sheetRow, RESULT_COL.correctionDate + 1).setValue(todayStr());
    rows[i][RESULT_COL.verdict] = 'בוטל';
    changed = true;
  }
  if (changed) SpreadsheetApp.flush();
}
function isFabricatedFailNote(note) {
  for (var i = 0; i < FABRICATED_FAIL_MARKERS.length; i++) {
    if (note.indexOf(FABRICATED_FAIL_MARKERS[i]) !== -1) return true;
  }
  return false;
}

// A genuine prior result for this exact exam. Skipped while the examinee still
// has an in_exam row (a retake after a disqualification is not a duplicate).
function findDuplicateResult(tail, data, gate) {
  if (hasPendingInExam(gate)) return null;
  var rows = tail.rows;
  for (var i = 1; i < rows.length; i++) {
    var verdict = String(rows[i][RESULT_COL.verdict] || '').trim();
    if (verdict === 'פסול' || verdict === 'בוטל') continue;
    if (!resultRowMatchesExaminee(rows[i], data)) continue;
    if (String(rows[i][RESULT_COL.license]) !== String(data.license)) continue;
    if (String(rows[i][RESULT_COL.language]) !== String(data.language || 'he')) continue;
    return rows[i];
  }
  return null;
}

// Re-read this examinee's pending rows: an examiner may have decided something
// while the submit was in flight. Never overwrite a newer examiner decision.
function hasPendingInExam(gate) {
  var snapshot = gate.pending;
  snapshot.rows = refreshExamineePendingRows(snapshot.sheet, snapshot.rows, gate.ctx.sessionCode, gate.ctx.idNumber);
  var rows = snapshot.rows;
  for (var i = rows.length - 1; i >= 1; i--) {
    if (!rows[i] || String(rows[i][0]) !== String(gate.ctx.sessionCode)) continue;
    if (normalizeId(rows[i][1]) !== normalizeId(gate.ctx.idNumber)) continue;
    if (String(rows[i][5]).trim() === 'in_exam') return true;
  }
  return false;
}

// An examinee who was auto-disqualified, let back in and then finished must not
// keep both a פסול row and a real one (base 14, 2026). The audit trail stays.
function supersedeDisqualifications(sheet, tail, data) {
  var rows = tail.rows;
  for (var i = rows.length - 1; i >= 1; i--) {
    if (!resultRowMatchesExaminee(rows[i], data)) continue;
    if (String(rows[i][RESULT_COL.verdict]).trim() !== 'פסול') continue;
    var sheetRow = i + tail.off + 1;
    sheet.getRange(sheetRow, RESULT_COL.verdict + 1).setValue('בוטל');
    sheet.getRange(sheetRow, RESULT_COL.dq + 1).setValue(false);
    sheet.getRange(sheetRow, RESULT_COL.correctionReason + 1).setValue('בוטל אוטומטית — נבחן ניגש למבחן מחדש');
    sheet.getRange(sheetRow, RESULT_COL.correctionDate + 1).setValue(todayStr());
    rows[i][RESULT_COL.verdict] = 'בוטל';
    rows[i][RESULT_COL.dq] = false;
  }
}

// Idempotency: the same result sent twice (a retry whose response was lost)
// must not append a second row. A retake differs in score or time, so it does.
function findIdenticalResult(tail, data) {
  var rows = tail.rows, verdict = resultVerdict(data);
  for (var i = rows.length - 1; i >= 1; i--) {
    if (!resultRowMatchesExaminee(rows[i], data)) continue;
    if (String(rows[i][RESULT_COL.license]) !== String(data.license)) continue;
    if (String(rows[i][RESULT_COL.score]) !== (data.score + '/' + data.total)) continue;
    if (String(rows[i][RESULT_COL.verdict]).trim() !== verdict) continue;
    if (String(rows[i][RESULT_COL.time]) !== String(data.time)) continue;
    return rows[i];
  }
  return null;
}

// ---- submitFailOnClose / cancelFailOnClose ---------------------------------
// The browser-close beacon. It writes a 0/30 fail, which a genuine submit later
// supersedes; it must never write one for an examinee the examiner reset.
function handleSubmitFailOnClose(data) {
  var ctx = examineeRowContext(data.sessionCode, data.idNumber);
  var status = ctx.latest ? ctx.latest.status : '';
  if (status === 'cancelled' || status === 'rejected') return jsonResponse({ status: 'ok', skipped: 'cancelled' });

  var sheet = getSheet('תוצאות');
  var tail = readTail(sheet, RESULT_COL.date);
  // Any non-בוטל result for this session+id means the examinee already has a
  // real outcome. Matched on session+id only — see supersedeFabricatedFails.
  for (var i = 1; i < tail.rows.length; i++) {
    if (!resultRowMatchesExaminee(tail.rows[i], data)) continue;
    if (String(tail.rows[i][RESULT_COL.verdict] || '').trim() === 'בוטל') continue;
    markPendingCompleted(data.sessionCode, data.idNumber, pendingSnapshotFromTail(ctx));
    return jsonResponse({ status: 'ok', duplicate: true });
  }

  var attemptNum = countAttempts(data.idNumber, data.license || '', attemptRows(tail)) + 1;
  var row = buildResultRow({
    idNumber: data.idNumber, fullName: data.fullName, phone: data.phone, license: data.license || '',
    score: 0, total: data.totalQuestions || 30, percent: 0, passed: false, time: data.time || '00:00',
    examinerName: data.examinerName, site: data.site, classroom: data.classroom,
    language: data.language || 'he', sessionCode: data.sessionCode, population: data.population,
    audioMode: data.audioMode, device: data.device, languageHistory: null, verified: false
  }, [], attemptNum, '');
  row[RESULT_COL.wrongDetails] = 'סגירת דפדפן באמצע מבחן (נענו ' + (data.answeredCount || 0) + ' שאלות)';
  row[RESULT_COL.audio] = data.audioMode || 'off';
  row[RESULT_COL.verified] = '';
  row[RESULT_COL.langPath] = '';      // a close-fail has no language path to report
  sheet.appendRow(row);

  markPendingCompleted(data.sessionCode, data.idNumber, pendingSnapshotFromTail(ctx));
  return jsonResponse({ status: 'ok' });
}

// The page reloaded rather than closed: undo the premature close-fail. The row
// is marked בוטל, never deleted — deleteRow was the only path that could destroy
// a result — and the examinee goes back to in_exam so they can finish.
function handleCancelFailOnClose(data) {
  var sessionCode = String(data.sessionCode || ''), id = normalizeId(data.idNumber || '');
  if (!sessionCode || !id) return jsonResponse({ status: 'ok' });
  var sheet = getSheet('תוצאות');
  var tail = readTail(sheet, RESULT_COL.date);
  for (var i = tail.rows.length - 1; i >= 1; i--) {
    if (!resultRowMatchesExaminee(tail.rows[i], data)) continue;
    if (String(tail.rows[i][RESULT_COL.wrongDetails] || '').indexOf('סגירת דפדפן') !== -1) {
      var sheetRow = i + tail.off + 1;
      sheet.getRange(sheetRow, RESULT_COL.verdict + 1).setValue('בוטל');
      sheet.getRange(sheetRow, RESULT_COL.correctionReason + 1).setValue('בוטל אוטומטי - רענון/חזרה למבחן');
      sheet.getRange(sheetRow, RESULT_COL.correctionDate + 1).setValue(todayStr());
      restorePendingToInExam(sessionCode, data.idNumber);
    }
    break;   // only the most recent row of this examinee
  }
  return jsonResponse({ status: 'ok' });
}

// The mirror of markPendingCompleted for a resumed exam: the newest completed
// row of this examinee goes back to in_exam, through the single status writer.
function restorePendingToInExam(sessionCode, idNumber) {
  var ctx = examineeRowContext(sessionCode, idNumber, true);
  if (!ctx.latest || ctx.latest.status !== 'completed') return;
  setPendingStatus(getSheet('ממתינים'), ctx.latest.rowNumber, sessionCode, 'in_exam', null);
}

// ---- Result-upload token (examiner → results Worker) ------------------------
// A short-lived HMAC the browser sends as X-Auth-Token when it POSTs the result
// HTML to the Cloudflare Worker; the Worker verifies it with the same secret.
// Set ScriptProperty RESULT_UPLOAD_SECRET and the Worker secret UPLOAD_SECRET to
// the same long random string. The secret never reaches the browser.
function handleGetResultUploadToken() {
  var secret = PropertiesService.getScriptProperties().getProperty('RESULT_UPLOAD_SECRET');
  if (!secret) {
    return jsonResponse({ status: 'error', code: 'not_configured',
      message: 'RESULT_UPLOAD_SECRET not configured in Apps Script properties' });
  }
  var payloadB64 = Utilities.base64EncodeWebSafe(JSON.stringify({ exp: Date.now() + 5 * 60 * 1000 })).replace(/=+$/, '');
  var sigB64 = Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(payloadB64, secret)).replace(/=+$/, '');
  return jsonResponse({ status: 'ok', token: payloadB64 + '.' + sigB64 });
}
// ---- Diagnostics that survive a killed execution ---------------------------
// Per-execution logs are unreachable in this project's Executions page (no
// Cloud project is linked), so a 360-second row says nothing about WHERE it
// hung. A phase marker is written to ScriptProperties once a request is already
// in trouble and deleted when it finishes; a killed execution never reaches the
// delete, so its marker is still there for the nightly sweep to record in the
// 'אבחון' sheet. Requests that finish but take longer than DIAG_SLOW_MS write
// their own row. Marks may be placed anywhere, polling handlers included: a
// healthy request writes nothing at all (see diagMark).
var DIAG_SHEET = 'אבחון';
var DIAG_SLOW_MS = 15000;
var DIAG_STALE_MS = 420000;
var DIAG_EXEC = null;

function diagBegin(method) {
  // Runs before the request's try/catch: it must be incapable of throwing.
  try { DIAG_EXEC = { id: Utilities.getUuid(), method: method, action: '', phase: '', marked: false, notes: [] }; }
  catch (e) { DIAG_EXEC = null; }
}

// r15: a mark is FREE until the request is already in trouble; r25: and then it
// costs exactly ONE service call, not one per mark.
//
// Every mark used to cost a ScriptProperties round-trip, which was affordable
// only because marks were kept off the polling handlers — leaving
// examinerDashboard, the handler we most needed to understand, with no trail at
// all (27.9 s on 2026-09-15, not a single phase). So the phase trail is kept in
// memory (free; diagFinish's SLOW row reads it from there) and the property —
// whose only job is to survive a 360 s kill — is written when the request first
// crosses DIAG_MARK_MIN_MS and NOT again (review C R9: a stalled dashboard paid
// six extra Properties round trips, a stalled submit about ten, on exactly the
// executions that were already failing).
// Trade-off, deliberate: a KILLED row names the action and the phase the
// request was in when it crossed 8 s, not the phase it died in. A request that
// finishes still reports its full phase list in its SLOW row.
var DIAG_MARK_MIN_MS = 8000;

function diagMark(phase) {
  try {
    if (!DIAG_EXEC) return;
    var elapsed = Date.now() - (DIAG_EXEC.t0 || Date.now());
    DIAG_EXEC.phase = phase;
    DIAG_EXEC.notes.push(phase + '@' + elapsed);
    if (elapsed < DIAG_MARK_MIN_MS || DIAG_EXEC.marked) return;   // healthy, or already marked: no service call
    DIAG_EXEC.marked = true;
    PropertiesService.getScriptProperties().setProperty(CACHE_KEY_PREFIX + 'diag_' + DIAG_EXEC.id,
      JSON.stringify({ a: DIAG_EXEC.action, m: DIAG_EXEC.method, ph: phase, t: Date.now() }));
  } catch (e) { /* diagnostics must never break a request */ }
}

function diagFinish(action, startedAt) {
  try {
    if (!DIAG_EXEC) return;
    var elapsed = Date.now() - startedAt;
    if (DIAG_EXEC.marked) {
      try { PropertiesService.getScriptProperties().deleteProperty(CACHE_KEY_PREFIX + 'diag_' + DIAG_EXEC.id); } catch (eDel) {}
    }
    if (elapsed >= DIAG_SLOW_MS) {
      diagRecordRow(DIAG_EXEC.id, [nowISO(), 'SLOW', DIAG_EXEC.method, action || DIAG_EXEC.action || '', elapsed,
        DIAG_EXEC.phase || '', DIAG_EXEC.notes.join(' ')]);
    }
  } catch (e) { /* never throw into the response path */ }
  finally { DIAG_EXEC = null; }
}

// A SLOW row used to be appended to the document straight from the response
// path of the slow request itself — one more write to the very document that
// had just stalled. 17/09 10:43: the append itself hung ~93 s (187 s execution,
// 92.9 s row); 16/09 morning: at least seven rows of 45-279 s executions never
// arrived, so the sheet read "healthy" at the worst moment. Now the row is
// parked in ScriptProperties first (a different service) and appended in place
// only while appends are healthy: one append that fails or takes longer than
// DIAG_APPEND_SLOW_MS opens a breaker for DIAG_APPEND_BREAKER_SEC, and the rows
// wait for the nightly sweep (or flushDiagnostics() from the editor).
// Parked rows are capped (review C R10): 200 rows were ~60 KB of the 500 KB
// ScriptProperties quota, cleared only by a sweep that reads the whole store.
var DIAG_ROW_PREFIX = 'diagrow_';
var DIAG_ROW_INDEX_KEY = 'diagrows';
var DIAG_MAX_PARKED_ROWS = 50;
var DIAG_APPEND_BREAKER_KEY = 'diagbreaker';
var DIAG_APPEND_BREAKER_SEC = 300;
var DIAG_APPEND_SLOW_MS = 2000;
function diagRecordRow(id, row) {
  var key = CACHE_KEY_PREFIX + DIAG_ROW_PREFIX + id;
  var parked = parkDiagRow(key, row);
  var cache = null, breakerOpen = false;
  try { cache = CacheService.getScriptCache(); breakerOpen = !!cache.get(CACHE_KEY_PREFIX + DIAG_APPEND_BREAKER_KEY); } catch (eGet) {}
  if (breakerOpen) return 'parked';
  var t0 = Date.now(), appended = false;
  try { getDiagnosticsSheet().appendRow(row); appended = true; } catch (eAppend) {}
  if (!appended || Date.now() - t0 > DIAG_APPEND_SLOW_MS) {
    try { if (cache) cache.put(CACHE_KEY_PREFIX + DIAG_APPEND_BREAKER_KEY, String(Date.now()), DIAG_APPEND_BREAKER_SEC); } catch (ePut) {}
  }
  if (appended && parked) unparkDiagRow(key);
  return appended ? 'appended' : 'parked';
}

// The index is a bounded FIFO of parked keys, so capping costs two Properties
// calls instead of reading the entire store on every park.
function parkDiagRow(key, row) {
  try {
    var props = PropertiesService.getScriptProperties();
    var index = [];
    try { index = JSON.parse(props.getProperty(CACHE_KEY_PREFIX + DIAG_ROW_INDEX_KEY) || '[]') || []; } catch (eIdx) { index = []; }
    props.setProperty(key, JSON.stringify(row));
    index.push(key);
    while (index.length > DIAG_MAX_PARKED_ROWS) {
      var oldest = index.shift();
      try { props.deleteProperty(oldest); } catch (eOld) {}
    }
    props.setProperty(CACHE_KEY_PREFIX + DIAG_ROW_INDEX_KEY, JSON.stringify(index));
    return true;
  } catch (eProp) { return false; }
}

function unparkDiagRow(key) {
  try {
    var props = PropertiesService.getScriptProperties();
    props.deleteProperty(key);
    var index = JSON.parse(props.getProperty(CACHE_KEY_PREFIX + DIAG_ROW_INDEX_KEY) || '[]') || [];
    var out = [];
    for (var i = 0; i < index.length; i++) if (index[i] !== key) out.push(index[i]);
    props.setProperty(CACHE_KEY_PREFIX + DIAG_ROW_INDEX_KEY, JSON.stringify(out));
  } catch (e) {}
}

// The client's own last-50-events ring, attached to a submit or to a failure
// report (DESIGN §3.6). Without it the only evidence of a client-side stall is
// the soldier's word: the 16/09 reload storm and the 17/09 "שגיאת תקשורת" both
// had to be reconstructed from screenshots. Goes through the same breaker as
// SLOW rows, is capped at 2 KB, and never throws into the response path.
var DIAG_CLIENT_LOG_MAX_CHARS = 2048;
function diagRecordClientLog(sessionCode, idNumber, entries) {
  try {
    if (!entries) return 'empty';
    var text = (typeof entries === 'string') ? entries : JSON.stringify(entries);
    if (!text || text === '[]' || text === '{}') return 'empty';
    if (text.length > DIAG_CLIENT_LOG_MAX_CHARS) text = text.slice(0, DIAG_CLIENT_LOG_MAX_CHARS - 1) + '…';
    var id = '';
    try { id = Utilities.getUuid(); } catch (eId) { id = 'client_' + Date.now(); }
    return diagRecordRow(id, [nowISO(), 'CLIENT', String(sessionCode || ''), normalizeId(idNumber), '', '', text]);
  } catch (e) { return 'error'; }
}

// Run from the editor during an exam morning if 'אבחון' looks empty while the
// dashboards are slow: writes the parked SLOW rows and records killed executions.
function flushDiagnostics() {
  var r = diagSweep(null);
  var msg = 'flushDiagnostics: ' + r.flushed + ' parked row(s) written, ' + r.swept + ' killed-execution marker(s) recorded';
  Logger.log(msg);
  return msg;
}

// Called by the nightly archive job (and by hand): markers older than
// DIAG_STALE_MS belong to executions that never finished (killed at 360 s, or
// crashed) — record where they were.
function diagSweep(summary) {
  var swept = 0, flushed = 0;
  try {
    var props = PropertiesService.getScriptProperties(), all = props.getProperties(), prefix = CACHE_KEY_PREFIX + 'diag_';
    var rowPrefix = CACHE_KEY_PREFIX + DIAG_ROW_PREFIX;
    var sheet = null;
    for (var key in all) {
      if (!Object.prototype.hasOwnProperty.call(all, key)) continue;
      if (key.indexOf(rowPrefix) === 0) {
        // a row parked while the append breaker was open (diagRecordRow)
        var row = null;
        try { row = JSON.parse(all[key]); } catch (eRow) {}
        if (row && row.length) {
          if (!sheet) sheet = getDiagnosticsSheet();
          sheet.appendRow(row);
          flushed++;
        }
        props.deleteProperty(key);
        continue;
      }
      if (key.indexOf(prefix) !== 0) continue;
      var entry = null;
      try { entry = JSON.parse(all[key]); } catch (e) {}
      if (entry && entry.t && Date.now() - entry.t < DIAG_STALE_MS) continue; // still running, leave it
      if (entry) {
        if (!sheet) sheet = getDiagnosticsSheet();
        sheet.appendRow([nowISO(), 'KILLED', entry.m || '', entry.a || '', '', entry.ph || '',
          'started ' + new Date(entry.t).toISOString()]);
      }
      props.deleteProperty(key);
      swept++;
    }
    try { props.deleteProperty(CACHE_KEY_PREFIX + DIAG_ROW_INDEX_KEY); } catch (eIdx) {}
  } catch (e) { if (summary) summary.push('diagnostics sweep: skipped (' + (e && e.message ? e.message : e) + ')'); }
  if (summary) summary.push('diagnostics sweep: ' + swept + ' stale marker(s) recorded, ' + flushed + ' parked row(s) flushed');
  return { swept: swept, flushed: flushed };
}

// Through getSpreadsheet(), not SpreadsheetApp.getActiveSpreadsheet(): the memo
// is opened once per execution and its open is what the 'ss:open' mark times
// (review C R9 — the diagnostics were the one caller still bypassing it).
function getDiagnosticsSheet() {
  var ss = getSpreadsheet(), sheet = ss.getSheetByName(DIAG_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(DIAG_SHEET);
    sheet.getRange(1, 1, 1, 7).setValues([['זמן', 'סוג', 'שיטה', 'פעולה', 'משך (ms)', 'שלב אחרון', 'הערות']]);
    sheet.getRange(1, 1, 1, 7).setFontWeight('bold');
  }
  return sheet;
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
  // r16: both big reads below have a provable lower date bound. This handler
  // looks at a result only when its date is >= prevFrom (the trend window that
  // precedes the requested one), and at a practice row only within 30 days
  // before such a result. So rows older than prevFrom (results) or prevFrom-30d
  // (practice) can never reach the output; readRowsSince stops reading there.
  // One extra day of margin on each covers the two date parsers disagreeing on
  // a boundary row. prevFrom is computed here, before the reads, and reused by
  // the trend logic further down - keep the two in step.
  var DAY_MS = 86400000;
  var prevFrom = (dateFrom && dateTo && dateFrom.getTime && dateTo.getTime)
    ? new Date(dateFrom.getTime() - (dateTo.getTime() - dateFrom.getTime()) - 1) : null;
  diagMark('sheet:results-commander');
  // readResultsSince, not readRowsSince: a range that reaches past the 30-day
  // retention window must include 'תוצאות_ארכיון' or the dashboard would report
  // a shorter history every night (B5).
  var resRead = readResultsSince(prevFrom ? new Date(prevFrom.getTime() - DAY_MS) : null);
  var resData = resRead.rows;
  diagMark('sheet:results-commander-done:' + resRead.mode);

  // Read practice results too — we'll join real-exam outcomes against the
  // practice history of the same name+license to surface a "did practice
  // before exam predict success?" metric. The student app stores its own
  // "מזהה תלמיד" (not the national ID), so we match only on full name +
  // license. Note that this is best-effort: identical names will collapse.
  // r17/r18: 'תוצאות תרגול' is 107,614 rows — 24x 'תוצאות' — so this read is
  // bounded on BOTH axes. Rows: only back to 31 days before prevFrom (see
  // rowsNeededSince). Columns: exactly the six this loop indexes —
  // date(0)=A, name(2)=C, class(3)=D, license(5)=F, percent(8)=I, phone(15)=P.
  // Everything else, above all the per-row JSON blobs in N ('פירוט שגויות', the
  // full text of every wrong question) and O, is never fetched.
  // ⚠ Index another column here and you MUST add it below — a column outside
  // the list reads as '' rather than failing. A test enforces the pairing.
  diagMark('sheet:practice-commander');
  var practiceSheet = getSheet('תוצאות תרגול');
  var practiceRead = readRowsSince(practiceSheet, 0, prevFrom ? new Date(prevFrom.getTime() - 31 * DAY_MS) : null,
    [[1, 1], [3, 2], [6, 1], [9, 1], [16, 1]]);
  var practiceData = practiceRead.rows;
  diagMark('sheet:practice-commander-done:' + practiceRead.mode);

  // Class → site map (from כיתות) — practice rows store the class code, not the
  // site, so this lets the name+site fallback match scope by base.
  var pClassSiteMap = {};
  try {
    diagMark('sheet:classes-commander');
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
  // prevFrom itself is computed above, before the sheet reads, because it
  // bounds how much of תוצאות / תוצאות תרגול readRowsSince has to fetch.
  if (!prevFrom) prevFrom = new Date(dateFrom.getTime() - prevWindowMs - 1);
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
    // Token-sort so "ישראל ישראלי" and "ישראלי ישראל" hash to the same key
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
        var qText = '', qCorrect = '', qId = '', qCategory = '';
        for (var wl = 0; wl < lines.length; wl++) {
          var line = lines[wl];
          if (line.indexOf('מזהה שאלה:') === 0) {
            qId = line.replace(/^מזהה שאלה:\s*/, '').trim();
          } else if (line.indexOf('קטגוריה:') === 0) {
            // submitResult writes the question's own category into the wrong
            // block, so the topic needs no question bank at all (§3.6).
            qCategory = line.replace(/^קטגוריה:\s*/, '').trim();
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
          wrongQuestionCounts[key] = { count: 0, text: qText, category: classifyCategoryServer(qCategory) || '', questionId: qId, langCounts: {} };
        }
        wrongQuestionCounts[key].count++;
        if (!wrongQuestionCounts[key].category && qCategory) wrongQuestionCounts[key].category = classifyCategoryServer(qCategory) || '';
        // Per-language split — a question failing mostly in one non-Hebrew
        // language is a translation-bug signal for the content team.
        wrongQuestionCounts[key].langCounts[langName] = (wrongQuestionCounts[key].langCounts[langName] || 0) + 1;
        // Weak topic, straight from the row: no language bank, no Drive, no
        // resolver loop (§3.6). A row written before the 'קטגוריה:' line
        // existed simply carries no topic and is not counted.
        topicBlocksParsed++;
        weakTopicPending.push({ license: license, topic: classifyCategoryServer(qCategory) || '' });
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

  // ===== Weak topics =====
  // Each pending item already carries its topic, parsed from the row's own
  // 'קטגוריה:' line. Until r25 this section loaded the question bank of up to
  // seven languages to map id → category: 14 Drive reads per request before
  // r12/r13, and even from the cache it was the reason the server had to keep
  // whole language banks in memory. What the commander loses: rows written
  // before 02/06/2026, which have no category line, no longer resolve — the
  // "asked" denominators are unchanged, so those rows lower the wrong-rate of
  // an old date range instead of raising it.
  var topicWrong = {};
  var topicWrongByLic = {};
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
    // Live + archive, both bounded by the requested range (review C R11: the
    // archive is never pruned and this block read all of it, whole, on every
    // commander dashboard). Columns: reg time (E), exam start (L), session code
    // (A), site (R) — see the loop below.
    diagMark('sheet:pending-commander');
    var pendReadW = readPendingSince(dateFrom, [[1, 1], [5, 1], [12, 1], [18, 1]]);
    var pendDataW = pendReadW.rows;
    diagMark('sheet:pending-commander-done:' + pendReadW.mode);
    var sessDataW = sessionRows();
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
  // §3.6: the server sends the ID, the count, the topic and the text AS IT WAS
  // SHOWN (the first occurrence in 'פירוט שגויות'); the client resolves the
  // canonical text and the image from the static bank by questionId. The old
  // shape carried `correctAnswer`/`imageUrl`, which cost up to seven language
  // banks per request to fill in.
  var topWrong = [];
  var wrongKeys = Object.keys(wrongQuestionCounts);
  wrongKeys.sort(function(a, b) {
    return wrongQuestionCounts[b].count - wrongQuestionCounts[a].count;
  });
  for (var wk = 0; wk < Math.min(wrongKeys.length, 10); wk++) {
    var entry = wrongQuestionCounts[wrongKeys[wk]];
    topWrong.push({
      questionId: entry.questionId || '',
      count: entry.count,
      category: entry.category || '',
      text: entry.text || '',
      langCounts: entry.langCounts || {}
    });
  }

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

  // 15/09 measured 12.1s between compute:commander-resolvers@32775 and the
  // request's 44932ms total — a quarter of the request after the last mark,
  // unattributed. This one closes the trail: anything left between it and the
  // total is serialisation of the payload plus diagFinish's own sheet append.
  diagMark('compute:commander-payload');
  return jsonResponse({ status: 'ok', data: result });
}

// ============================================================================
// ========== Pass-probability calibration engine (predictive BI) ==========
// ============================================================================
// STEP 1 of the predictive layer. Learns, from HISTORICAL real-exam outcomes
// joined to each examinee's practice history, how a practice profile maps to
// the probability of PASSING the real theory exam.
//
// This is the SINGLE SOURCE OF TRUTH that both commander dashboards will use:
//   - teacher-commander  → per-soldier PREVENTIVE risk score (before the exam)
//   - examiner-commander → cohort pass-rate FORECAST (exam-day planning)
//
// Method: hierarchical shrinkage (empirical-Bayes style). We deliberately do
// NOT train a heavy ML model — with a few thousand rows that would over-fit and
// be impossible to explain to a commander. Instead we compute the observed pass
// rate per (license, last-practice-score bin) cell, and shrink each cell toward
// its license base rate, then toward the global base rate, by sample size.
// Sparse cells fall back gracefully; rich cells keep their specific signal.
// Every estimate ships with its support (n) so the UI can be honest about
// confidence. Same join logic as handleCommanderDashboard (phone → name+site →
// name, latest practice within a lookback window before the exam).

// Ordered score bins for the last practice attempt before the exam. Finer than
// the 3 practice-impact buckets so the probability curve has resolution near
// the pass threshold (~86% in practice terms).
var PP_BINS = ['0-49', '50-59', '60-69', '70-79', '80-85', '86-92', '93-100'];
function ppBin(pct) {
  if (pct == null || pct < 0) return null;   // unparseable / no score
  if (pct < 50) return '0-49';
  if (pct < 60) return '50-59';
  if (pct < 70) return '60-69';
  if (pct < 80) return '70-79';
  if (pct < 86) return '80-85';
  if (pct < 93) return '86-92';
  return '93-100';
}

// Empirical-Bayes shrinkage: blend the cell's observed rate with a prior
// (parent) rate, weighted by K "virtual" observations of the prior. Returns a
// probability in [0,1]. K controls how much support a cell needs before it
// out-weighs its parent — K=12 means a cell needs ~12 samples to carry half the
// weight. priorRate is already a probability in [0,1].
function ppShrink(passed, n, priorRate, k) {
  return (passed + k * priorRate) / (n + k);
}

// Local, self-contained copies of the join/normalization helpers (the versions
// inside handleCommanderDashboard are private to that function). Kept identical
// on purpose so the model's join matches the dashboard's practice-impact join.
function ppNormName(s) {
  if (!s) return '';
  var t = String(s).trim();
  if (!t) return '';
  t = t.replace(/[׳״'".\-]/g, ' ').replace(/\s+/g, ' ').trim().toLowerCase();
  var tokens = t.split(' ').filter(function(x) { return x; });
  tokens.sort();
  return tokens.join(' ');
}
function ppNormPhone(v) {
  var d = String(v || '').replace(/\D/g, '');
  return d.length >= 9 ? d.slice(-9) : '';
}
function ppParsePct(v) {
  if (typeof v === 'number') return v <= 1 ? v * 100 : v;
  var s = String(v || '').replace('%', '').trim();
  if (!s) return -1;
  var n = parseFloat(s);
  if (isNaN(n)) return -1;
  return n <= 1 ? n * 100 : n;
}

// Given the list of practice records (each {date, percent}) that fall in the
// lookback window before an exam, derive the feature vector the model keys on.
// list must be sorted newest-first. Returns null when there's no usable
// practice (the "no practice" branch is modelled separately).
function ppExtractFeatures(list, examDate, lookbackDays) {
  if (!list || !list.length) return null;
  var windowStart = new Date(examDate.getTime() - lookbackDays * 86400000);
  var inWin = [];
  for (var i = 0; i < list.length; i++) {
    var rec = list[i];
    if (rec.date <= examDate && rec.date >= windowStart && rec.percent >= 0) inWin.push(rec);
  }
  if (!inWin.length) return null;
  // inWin is newest-first. Latest = the attempt closest to the exam.
  var latest = inWin[0];
  var oldest = inWin[inWin.length - 1];
  var best = -1;
  for (var j = 0; j < inWin.length; j++) if (inWin[j].percent > best) best = inWin[j].percent;
  var daysSince = Math.round((examDate.getTime() - latest.date.getTime()) / 86400000);
  // Trend across the window: newest minus oldest. >3pt = improving, <-3 = declining.
  var trendPts = inWin.length > 1 ? (latest.percent - oldest.percent) : 0;
  return {
    lastPct: latest.percent,
    bestPct: best,
    sessions: inWin.length,
    daysSince: daysSince,
    trendPts: trendPts,
    trend: inWin.length < 2 ? 'single' : (trendPts > 3 ? 'up' : (trendPts < -3 ? 'down' : 'flat'))
  };
}

// Build the calibration model from the full history. Reads תוצאות + תוצאות תרגול
// directly so it's independent of handleCommanderDashboard. opts:
//   lookbackDays (default 30) — practice window before each exam
//   sinceDate    (optional Date) — ignore exams before this (bound the history)
function buildPassProbabilityModel(opts) {
  opts = opts || {};
  var lookbackDays = opts.lookbackDays || 30;
  var sinceDate = opts.sinceDate || null;

  // Both reads are column-pruned, and the results read spans live + archive
  // (B5). Exam columns used below: A date, B id, C name, D phone, E licence,
  // H pass, K site, O attempt, R פסול. Practice columns: A date, C name,
  // D class, F licence, I percent, P phone — never N/O, the two JSON blobs
  // that made the practice read 28.5 s (r17).
  // ⚠ Index another column here and it MUST be added to the colSpec; a column
  // outside the list reads as '' instead of failing.
  var resData = readResultsSince(sinceDate, [[1, 5], [8, 1], [11, 1], [15, 1], [18, 1]]).rows;
  var practiceData = readRowsSince(getSheet('תוצאות תרגול'), 0,
    sinceDate ? new Date(sinceDate.getTime() - lookbackDays * 86400000) : null,
    [[1, 1], [3, 2], [6, 1], [9, 1], [16, 1]]).rows;

  // Class → site map (practice rows store the class code, not the site).
  var classSiteMap = {};
  try {
    var classData = getSheet('כיתות').getDataRange().getValues();
    for (var c = 1; c < classData.length; c++) {
      classSiteMap[String(classData[c][0]).trim()] = String(classData[c][7] || '').trim();
    }
  } catch (eC) { /* no כיתות → name+site fallback stays empty */ }

  // Practice indexes (same three keys as the dashboard).
  var byName = {}, byPhone = {}, byNameSite = {};
  for (var pi = 1; pi < practiceData.length; pi++) {
    var pName = ppNormName(practiceData[pi][2]);
    if (!pName) continue;
    var pLic = String(practiceData[pi][5] || '').trim();
    var pDate = practiceData[pi][0];
    if (pDate && !(pDate instanceof Date)) pDate = new Date(pDate);
    if (!pDate || isNaN(pDate.getTime())) continue;
    var rec = { date: pDate, percent: ppParsePct(practiceData[pi][8]) };
    var nk = pName + '|' + pLic;
    (byName[nk] = byName[nk] || []).push(rec);
    var ph = (practiceData[pi].length > 15) ? ppNormPhone(practiceData[pi][15]) : '';
    if (ph) (byPhone[ph] = byPhone[ph] || []).push(rec);
    var site = classSiteMap[String(practiceData[pi][3] || '').trim()] || '';
    if (site) { var nsk = pName + '|' + pLic + '|' + site; (byNameSite[nsk] = byNameSite[nsk] || []).push(rec); }
  }
  function sortDesc(idx) { for (var k in idx) idx[k].sort(function(a, b) { return b.date - a.date; }); }
  sortDesc(byName); sortDesc(byPhone); sortDesc(byNameSite);

  var examinerExcl = getExaminerExclusion();

  // Accumulators. The model keys cells on license × attempt × score-bin — real
  // data shows attempt number is as strong a predictor as license (attempt 1
  // ~45% pass, 2 ~36%, 3+ ~19%) and it's known before the exam, so it's a
  // first-class axis, not just a marginal. Hierarchy for shrinkage:
  // cell(lic|att|bin) → lic|att → lic → base.
  var base = { n: 0, passed: 0 };
  var byLic = {};                 // license → {n, passed}
  var byLicAtt = {};              // 'license|attempt' → {n, passed}  (shrinkage prior)
  var cells = {};                 // 'license|attempt|bin' → {n, passed}
  var noPractice = { _all: { n: 0, passed: 0 } };   // 'license|attempt' → {n,passed}, plus _all
  var byAttempt = {};             // '1' / '2' / '3+' → {n, passed} (marginal diagnostic)
  var bySessions = {};            // '1' / '2' / '3+' → {n, passed} (marginal diagnostic)
  var byTrend = {};               // up/flat/down/single → {n, passed} (marginal diagnostic)
  var coverage = { eligible: 0, matched: 0, byPhone: 0, byNameSite: 0, byName: 0 };

  // Attempt number → bucket. Col 14 holds the attempt index (1,2,3...).
  function attBucket(v) { var n = Number(v) || 1; return n >= 3 ? '3+' : String(n); }

  function bump(obj, key, passed) {
    if (!obj[key]) obj[key] = { n: 0, passed: 0 };
    obj[key].n++; if (passed) obj[key].passed++;
  }

  for (var r = 1; r < resData.length; r++) {
    var rowDate = parseSheetDate(resData[r][0]);
    if (!rowDate) continue;
    if (sinceDate && rowDate < sinceDate) continue;

    var siteName = String(resData[r][10] || '');
    if (isTestSite(siteName)) continue;
    if (isExaminerSelfTest(resData[r][2], resData[r][1], examinerExcl)) continue;

    var passedStr = String(resData[r][7] || '');
    if (passedStr === 'בוטל') continue;
    var isDQ = resData[r][17] === true || String(resData[r][17]).toUpperCase() === 'TRUE' || passedStr === 'פסול';
    if (isDQ) continue;   // DQ ≠ knowledge; excluded from the pass model (matches practice-impact)
    var isPassed = (passedStr === 'עבר') ? 1 : 0;

    var license = String(resData[r][4] || '').trim() || 'לא צוין';
    var att = attBucket(resData[r][14]);
    var licAtt = license + '|' + att;
    base.n++; if (isPassed) base.passed++;
    bump(byLic, license, isPassed);
    bump(byLicAtt, licAtt, isPassed);
    bump(byAttempt, att, isPassed);

    // Join to practice history (phone → name+site → name), same priority as the
    // dashboard. examDate must be a real Date for the window math.
    var examName = ppNormName(resData[r][2]);
    var examDate = resData[r][0];
    if (examDate && !(examDate instanceof Date)) examDate = new Date(examDate);
    if (!examName || !examDate || isNaN(examDate.getTime())) { bump(noPractice, licAtt, isPassed); noPractice._all.n++; if (isPassed) noPractice._all.passed++; continue; }

    coverage.eligible++;
    var examPhone = ppNormPhone(resData[r][3]);
    var examSite = siteName.trim();
    var list = null, matchType = '';
    if (examPhone && byPhone[examPhone]) { list = byPhone[examPhone]; matchType = 'byPhone'; }
    if (!list && examSite && byNameSite[examName + '|' + license + '|' + examSite]) { list = byNameSite[examName + '|' + license + '|' + examSite]; matchType = 'byNameSite'; }
    if (!list && byName[examName + '|' + license]) { list = byName[examName + '|' + license]; matchType = 'byName'; }

    var feat = list ? ppExtractFeatures(list, examDate, lookbackDays) : null;
    if (!feat) {
      bump(noPractice, licAtt, isPassed);
      noPractice._all.n++; if (isPassed) noPractice._all.passed++;
      continue;
    }
    coverage.matched++;
    if (coverage[matchType] != null) coverage[matchType]++;

    var bin = ppBin(feat.lastPct);
    if (bin) bump(cells, licAtt + '|' + bin, isPassed);
    var sessKey = feat.sessions >= 3 ? '3+' : String(feat.sessions);
    bump(bySessions, sessKey, isPassed);
    bump(byTrend, feat.trend, isPassed);
  }

  // Finalize rates (probability in [0,1]) for every accumulator.
  function rate(o) { return o && o.n > 0 ? o.passed / o.n : 0; }
  base.rate = rate(base);
  for (var lk in byLic) byLic[lk].rate = rate(byLic[lk]);
  for (var lak in byLicAtt) byLicAtt[lak].rate = rate(byLicAtt[lak]);
  for (var ck in cells) cells[ck].rate = rate(cells[ck]);
  for (var nk2 in noPractice) noPractice[nk2].rate = rate(noPractice[nk2]);
  for (var ak in byAttempt) byAttempt[ak].rate = rate(byAttempt[ak]);
  for (var sk in bySessions) bySessions[sk].rate = rate(bySessions[sk]);
  for (var tk in byTrend) byTrend[tk].rate = rate(byTrend[tk]);

  // ---- Monotonic (isotonic) enforcement across the score bins ----
  // A higher practice score must never predict a LOWER pass probability. Raw
  // shrunk cell rates violate this in sparse low bins (e.g. C1 60-69 observed
  // 6% while 0-49 observed 11% — pure small-sample noise), which would float
  // mid-scorers above low-scorers in the at-risk ranking. We pool-adjacent-
  // violators (PAV, weighted by cell support) the shrunk probabilities per
  // (license, attempt) so predictions are non-decreasing in the score bin.
  var K = 12;
  function pav(values, weights) {
    var blocks = [];
    for (var j = 0; j < values.length; j++) {
      blocks.push({ v: values[j], w: weights[j], len: 1 });
      while (blocks.length > 1 && blocks[blocks.length - 2].v > blocks[blocks.length - 1].v) {
        var b2 = blocks.pop(), b1 = blocks.pop();
        var w = b1.w + b2.w;
        blocks.push({ v: (b1.v * b1.w + b2.v * b2.w) / (w || 1), w: w, len: b1.len + b2.len });
      }
    }
    var out = [];
    for (var q = 0; q < blocks.length; q++) for (var t = 0; t < blocks[q].len; t++) out.push(blocks[q].v);
    return out;
  }
  var monoCells = {};
  var attList = ['1', '2', '3+'];
  for (var lic2 in byLic) {
    var licRate2 = ppShrink(byLic[lic2].passed, byLic[lic2].n, base.rate, K);
    for (var ai = 0; ai < attList.length; ai++) {
      var att2 = attList[ai];
      var la = byLicAtt[lic2 + '|' + att2];
      var laRate2 = ppShrink(la ? la.passed : 0, la ? la.n : 0, licRate2, K);
      var vals = [], wts = [];
      for (var bi2 = 0; bi2 < PP_BINS.length; bi2++) {
        var cc = cells[lic2 + '|' + att2 + '|' + PP_BINS[bi2]];
        vals.push(ppShrink(cc ? cc.passed : 0, cc ? cc.n : 0, laRate2, K));
        wts.push(cc ? Math.max(cc.n, 1) : 1);
      }
      var mono = pav(vals, wts);
      for (var bi3 = 0; bi3 < PP_BINS.length; bi3++) monoCells[lic2 + '|' + att2 + '|' + PP_BINS[bi3]] = mono[bi3];
    }
  }

  return {
    version: 3,
    k: 12,
    lookbackDays: lookbackDays,
    bins: PP_BINS,
    base: base,
    byLicense: byLic,
    byLicAtt: byLicAtt,
    cells: cells,
    monoCells: monoCells,
    noPractice: noPractice,
    byAttempt: byAttempt,
    bySessions: bySessions,
    byTrend: byTrend,
    coverage: coverage
  };
}

// Predict pass probability for one examinee profile from a built model.
// features: { license, attempt (1/2/3+ — the UPCOMING attempt number; default 1),
//             lastPct (number, or null/undefined if no practice) }
// Shrinkage chain: cell(lic|att|bin) → lic|att → lic → base. Returns
//   { prob (0-100 int), n (support of the most specific cell used),
//     basis ('cell'|'licenseAttempt'|'license'|'base'|'noPractice'), confidence }.
function predictPassProbability(model, features) {
  if (!model || !model.base) return null;
  var k = model.k || 12;
  var license = (features && features.license) ? String(features.license).trim() : '';
  var attNum = features && features.attempt != null ? (Number(features.attempt) || 1) : 1;
  var att = attNum >= 3 ? '3+' : String(attNum);
  var licAtt = license + '|' + att;

  var licNode = license && model.byLicense[license] ? model.byLicense[license] : null;
  // License rate shrunk toward the global base.
  var licRate = ppShrink(licNode ? licNode.passed : 0, licNode ? licNode.n : 0, model.base.rate, k);
  // License+attempt rate shrunk toward the license rate — the working prior.
  var laNode = model.byLicAtt && model.byLicAtt[licAtt] ? model.byLicAtt[licAtt] : null;
  var laRate = ppShrink(laNode ? laNode.passed : 0, laNode ? laNode.n : 0, licRate, k);

  var hasPractice = features && features.lastPct != null && features.lastPct >= 0;
  var prob, n, basis;
  if (!hasPractice) {
    var npLA = model.noPractice[licAtt] || null;
    // no-practice(lic|att) shrunk toward the lic|att overall rate.
    prob = ppShrink(npLA ? npLA.passed : 0, npLA ? npLA.n : 0, laRate, k);
    n = npLA ? npLA.n : 0;
    basis = 'noPractice';
  } else {
    var bin = ppBin(features.lastPct);
    var cell = bin ? model.cells[licAtt + '|' + bin] : null;
    // Prefer the monotonic (isotonic) probability so a higher practice score
    // never predicts a lower pass chance; fall back to the raw shrunk rate on
    // older models that predate monoCells.
    var monoKey = licAtt + '|' + bin;
    if (bin && model.monoCells && model.monoCells[monoKey] != null) {
      prob = model.monoCells[monoKey];
    } else {
      prob = ppShrink(cell ? cell.passed : 0, cell ? cell.n : 0, laRate, k);
    }
    n = cell ? cell.n : 0;
    basis = (cell && cell.n >= k) ? 'cell' : (laNode && laNode.n >= k ? 'licenseAttempt' : (licNode ? 'license' : 'base'));
  }
  // Confidence from the support behind the estimate.
  var confidence = n >= 40 ? 'high' : (n >= 12 ? 'medium' : 'low');
  return { prob: Math.round(prob * 100), n: n, basis: basis, confidence: confidence, licenseBaseRate: Math.round(licRate * 100), licenseAttemptRate: Math.round(laRate * 100) };
}
// handlePredictiveModelPreview removed (review C R14): a diagnostic endpoint no
// client ever called, whose only job was to eyeball the model before it was
// wired into the dashboards — which it now is (at-risk list, examinerForecast).
// It also built the whole model synchronously on a doGet.
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
  var valid = verifyTeacherToken(p.teacherId, p.token);
  diagMark('auth:teacher');
  if (!valid) {
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

// Read-only lookup of classes that were deleted (archived by handleTeacherDeleteClass).
// Uses getSheetByName (NOT getSheet) so a report read never auto-creates the sheet.
// Returns {} if the archive doesn't exist yet (no class has ever been deleted).
function getDeletedClassMap() {
  var map = {};
  try {
    var sheet = getSheetIfExists('כיתות שנמחקו');
    if (!sheet) return map;
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var cc = String(data[i][0] || '').trim();
      if (!cc) continue;
      map[cc] = {
        teacherName: String(data[i][3] || ''),
        className: String(data[i][1] || ''),
        license: String(data[i][4] || ''),
        site: String(data[i][5] || ''),
        deleted: true
      };
    }
  } catch (e) { /* archive missing/unreadable → treat as empty */ }
  return map;
}

// Resolve a practice-result class code to display info for the commander reports.
// Precedence: active class → deleted class (real teacher + "(כיתה שנמחקה)" tag) →
// truly unrecognized code ("קוד לא מזוהה"). This is what lets the dashboard tell a
// legitimate deleted class apart from a forged/never-existed code, instead of
// lumping both under "לא ידוע".
function resolveClassInfo(classCode, classMap, deletedMap) {
  if (classMap && classMap[classCode]) return classMap[classCode];
  if (deletedMap && deletedMap[classCode]) {
    var d = deletedMap[classCode];
    return {
      teacherName: d.teacherName || 'לא ידוע',
      teacherId: '',
      className: (d.className || classCode) + ' (כיתה שנמחקה)',
      license: d.license || '',
      site: d.site || '',
      deleted: true
    };
  }
  return { teacherName: 'קוד לא מזוהה', teacherId: '', className: classCode, license: '', site: '', unresolved: true };
}

function handleTeacherLogin(p) {
  var maintenance = practiceWriteGuard(); if (maintenance) return maintenance;   // r24
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
  var deletedClassMap = getDeletedClassMap();

  // Read practice results — bounded by the requested range (the loop drops
  // anything outside it anyway). All columns: unlike the commander view this
  // one aggregates the wrong-question JSON in column N.
  diagMark('sheet:practice-teacher-commander');
  var resRead = readRowsSince(getSheet('תוצאות תרגול'), 0, dateFrom);
  var resData = resRead.rows;
  diagMark('sheet:practice-teacher-commander-done:' + resRead.mode);

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

    var cInfo = resolveClassInfo(classCode, classMap, deletedClassMap);
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
    // Live + archive (B5), bounded by the range, columns A-H + K (site).
    diagMark('sheet:results-teacher-commander');
    var examResData = readResultsSince(dateFrom, [[1, 8], [11, 1]]).rows;
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

// ========== At-risk examinee list (PREVENTIVE — step 2 of predictive BI) ==========
// The teacher-commander's preventive view: for every student who practiced in the
// window and has NOT yet passed the real theory exam, predict their pass
// probability from the calibration model (license × upcoming-attempt × last
// practice score) and rank worst-first. Lets the training commander route
// weak soldiers to more practice BEFORE they burn a real exam slot. Highest
// value for C1 first-timers (huge volume, ~22% base pass, and practice score
// cleanly separates ~10% from ~73%). Complements repeatFailures (which is
// retrospective — already failed 2+); this catches them before the first fail.
// Heavy computation for the WHOLE fleet — builds the model once and scores every
// recently-practiced student (no scope filter). Called ONLY by the nightly cache
// rebuild, never on a dashboard request. Returns the full ranked list + build
// summary + computedAt so the cache can be scope-filtered cheaply on read.
function computeAtRiskAll(opts) {
  opts = opts || {};
  var lookbackDays = opts.lookbackDays || 30;
  var model = buildPassProbabilityModel({ lookbackDays: lookbackDays });

  // Class map: code → {teacherId, teacherName, className, license, site}.
  var classData = getSheet('כיתות').getDataRange().getValues();
  var classMap = {};
  for (var c = 1; c < classData.length; c++) {
    classMap[String(classData[c][0]).trim()] = {
      teacherId: normalizeId(classData[c][2]),
      teacherName: String(classData[c][3] || ''),
      className: String(classData[c][1] || ''),
      license: String(classData[c][4] || ''),
      site: String(classData[c][7] || '')
    };
  }
  var deletedClassMap = getDeletedClassMap();

  // Exam-history index → upcoming attempt number + everPassed. Keyed by phone AND
  // name|license, mirroring the model join.
  // Live + archive (B5), columns C name, D phone, E licence, H pass, O attempt.
  // ⚠ Index another column and add it to the colSpec — a missing column reads ''.
  var examData = readResultsSince(null, [[3, 3], [8, 1], [15, 1]]).rows;
  var histByPhone = {}, histByName = {};
  function histBump(idx, key, attempt, passed) {
    if (!key) return;
    if (!idx[key]) idx[key] = { attempts: 0, everPassed: false };
    if (attempt > idx[key].attempts) idx[key].attempts = attempt;
    if (passed) idx[key].everPassed = true;
  }
  for (var e = 1; e < examData.length; e++) {
    var ePassedStr = String(examData[e][7] || '');
    if (ePassedStr === 'בוטל') continue;
    var eName = ppNormName(examData[e][2]);
    var eLic = String(examData[e][4] || '').trim();
    var ePhone = ppNormPhone(examData[e][3]);
    var eAtt = Number(examData[e][14]) || 1;
    var ePassed = (ePassedStr === 'עבר');
    histBump(histByPhone, ePhone, eAtt, ePassed);
    histBump(histByName, eName + '|' + eLic, eAtt, ePassed);
  }

  // Group practice rows by student within the window — no scope filter here (the
  // cache holds everyone; the read handler filters by caller scope).
  var windowStart = new Date();
  windowStart.setDate(windowStart.getDate() - lookbackDays);
  // Rows: only the window this loop keeps (it drops anything older itself).
  // Columns: A date, B studentId, C name, D class, F licence, I percent,
  // P phone — never the two JSON blobs in N/O.
  var practiceData = readRowsSince(getSheet('תוצאות תרגול'), 0, windowStart,
    [[1, 6], [9, 1], [16, 1]]).rows;
  var students = {};
  for (var r = 1; r < practiceData.length; r++) {
    var pDate = parseSheetDate(practiceData[r][0]);
    if (!pDate || pDate < windowStart) continue;
    var classCode = String(practiceData[r][3] || '').trim();
    var cInfo = resolveClassInfo(classCode, classMap, deletedClassMap);
    var pct = ppParsePct(practiceData[r][8]);
    if (pct < 0) continue;
    var name = String(practiceData[r][2] || '');
    var lic = String(practiceData[r][5] || cInfo.license || '').trim();
    var phone = (practiceData[r].length > 15) ? ppNormPhone(practiceData[r][15]) : '';
    var studentId = String(practiceData[r][1] || '');
    var rowResolved = !cInfo.unresolved;   // class code mapped to a real (active/deleted) class
    var key = phone ? ('p:' + phone) : (studentId ? ('s:' + studentId) : ('n:' + ppNormName(name) + '|' + lic));
    if (!students[key]) {
      students[key] = { name: name, license: lic, phone: phone, studentId: studentId, classCode: classCode, teacherId: cInfo.teacherId || '', className: cInfo.className, teacherName: cInfo.teacherName, site: cInfo.site || '', resolved: rowResolved, recs: [] };
    } else if (!students[key].resolved && rowResolved) {
      // Upgrade to a resolved class if an earlier row had a blank/unknown code.
      students[key].classCode = classCode; students[key].teacherId = cInfo.teacherId || '';
      students[key].className = cInfo.className; students[key].teacherName = cInfo.teacherName;
      students[key].site = cInfo.site || ''; students[key].resolved = true;
    }
    students[key].recs.push({ date: pDate, pct: pct });
    if (name) students[key].name = name;
    if (lic) students[key].license = lic;
    if (studentId && !students[key].studentId) students[key].studentId = studentId;
  }

  // Roster fallback — many students practice WITHOUT a class code (it's optional)
  // so their rows show "קוד לא מזוהה" and can't be attributed to a teacher. Recover
  // the attribution from the class rosters (תלמידי כיתות: code, name, studentId):
  // match by student id first, then by normalized name. Best-effort — a name that
  // appears in two classes maps to whichever was seen first.
  var rosterById = {}, rosterByName = {};
  try {
    var rosterData = getSheet('תלמידי כיתות').getDataRange().getValues();
    for (var rs = 1; rs < rosterData.length; rs++) {
      var rCode = String(rosterData[rs][0] || '').trim();
      if (!rCode) continue;
      var rSid = String(rosterData[rs][2] || '').trim();
      var rNk = ppNormName(rosterData[rs][1]);
      if (rSid && !rosterById[rSid]) rosterById[rSid] = rCode;
      if (rNk && !rosterByName[rNk]) rosterByName[rNk] = rCode;
    }
  } catch (eRoster) { /* no roster sheet → fallback simply does nothing */ }
  for (var uk in students) {
    var us = students[uk];
    if (us.resolved) continue;
    var rc = (us.studentId && rosterById[us.studentId]) || rosterByName[ppNormName(us.name)] || '';
    if (!rc) continue;
    var rci = resolveClassInfo(rc, classMap, deletedClassMap);
    if (rci.unresolved) continue;
    us.classCode = rc; us.teacherId = rci.teacherId || '';
    us.className = rci.className; us.teacherName = rci.teacherName;
    us.site = rci.site || ''; us.resolved = true; us.attributedViaRoster = true;
  }

  // Junk-name filter — drops shared/demo entries (a "." with 100+ sessions, a
  // cycle name typed into the name field, punctuation/number-only names).
  function looksLikeName(s) {
    var t = String(s || '').trim();
    if (t.length < 2) return false;
    var letters = t.replace(/[^A-Za-z֐-׿]/g, '');
    return letters.length >= 2;
  }
  var out = [];
  var summary = { high: 0, medium: 0, low: 0, alreadyPassed: 0, junk: 0, total: 0 };
  for (var k in students) {
    var st = students[k];
    if (!looksLikeName(st.name)) { summary.junk++; continue; }
    st.recs.sort(function(a, b) { return b.date - a.date; });
    var latest = st.recs[0];
    var oldest = st.recs[st.recs.length - 1];
    var sessions = st.recs.length;
    var trendPts = sessions > 1 ? (latest.pct - oldest.pct) : 0;
    var trend = sessions < 2 ? 'single' : (trendPts > 3 ? 'up' : (trendPts < -3 ? 'down' : 'flat'));
    var hist = (st.phone && histByPhone[st.phone]) || histByName[ppNormName(st.name) + '|' + st.license] || null;
    if (hist && hist.everPassed) { summary.alreadyPassed++; continue; }
    var upcomingAttempt = (hist ? hist.attempts : 0) + 1;
    var pred = predictPassProbability(model, { license: st.license, attempt: upcomingAttempt, lastPct: latest.pct });
    var prob = pred ? pred.prob : null;
    var tier = prob == null ? 'low' : (prob < 40 ? 'high' : (prob < 65 ? 'medium' : 'low'));
    summary[tier]++; summary.total++;
    out.push({
      name: st.name, license: st.license, classCode: st.classCode, teacherId: st.teacherId,
      teacherName: st.teacherName, className: st.className, site: st.site,
      lastPct: Math.round(latest.pct), sessions: sessions, trend: trend,
      attempt: upcomingAttempt, everTested: !!hist,
      prob: prob, tier: tier, confidence: pred ? pred.confidence : 'low', matchedByPhone: !!st.phone, phone: st.phone || ''
    });
  }
  out.sort(function(a, b) {
    var pa = a.prob == null ? 999 : a.prob, pb = b.prob == null ? 999 : b.prob;
    if (pa !== pb) return pa - pb;
    if (a.lastPct !== b.lastPct) return a.lastPct - b.lastPct;
    return b.sessions - a.sessions;
  });
  return { computedAtMs: new Date().getTime(), students: out, summary: summary, modelBaseRate: Math.round(model.base.rate * 100) };
}

// Nightly job — recompute the whole at-risk list and persist it to the
// 'חיזוי סיכון' sheet + script properties (timestamp, summary, base rate). Run
// by a time-based trigger (see installAtRiskTrigger) at ~03:00 so the heavy
// model build never lands during exam/practice hours. Dashboards then READ this
// cache instead of rebuilding — no per-request model build, no midday load.
// The job-running FLAG is gone with the warmup that read it (r25). What the two
// nightly jobs need from each other is stronger than a flag anyway: this one
// reads 'תוצאות' live + archive, and the 01:00 archive MOVES rows between those
// two sheets, so a run that overlapped it could count a row twice or not at all.
// The script lock the archive already holds is the real interlock; 03:00 is two
// hours later, so waiting for it is a formality that costs nothing.
function rebuildAtRiskCache() {
  var lock = LockService.getScriptLock();
  var held = false;
  try { held = lock.tryLock(60000); } catch (e) { held = false; }
  if (!held) Logger.log('rebuildAtRiskCache: the archive still holds the lock — computing anyway');
  try { return rebuildAtRiskCacheInner(); }
  finally { if (held) lock.releaseLock(); }
}
function rebuildAtRiskCacheInner() {
  var res = computeAtRiskAll({ lookbackDays: 30 });
  var sheet = getSheet('חיזוי סיכון');
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
  var computedAtStr = Utilities.formatDate(new Date(res.computedAtMs), 'Asia/Jerusalem', 'yyyy-MM-dd HH:mm');
  var rows = res.students.map(function(s) {
    return [computedAtStr, s.name, s.license, s.classCode, s.teacherId, s.teacherName, s.className, s.site,
      s.lastPct, s.sessions, s.trend, s.attempt, s.everTested, (s.prob == null ? '' : s.prob), s.tier, s.confidence, s.matchedByPhone, (s.phone || '')];
  });
  if (rows.length) sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  var props = PropertiesService.getScriptProperties();
  props.setProperty('atRisk_computedAt', computedAtStr);
  props.setProperty('atRisk_summary', JSON.stringify(res.summary));
  props.setProperty('atRisk_modelBaseRate', String(res.modelBaseRate));
  return { computed: res.students.length, computedAt: computedAtStr };
}

// One-time setup — installs the nightly trigger. Run once from the Apps Script
// editor (Run → installAtRiskTrigger). Idempotent: removes any prior copy first.
function installAtRiskTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'rebuildAtRiskCache') ScriptApp.deleteTrigger(triggers[i]);
  }
  ScriptApp.newTrigger('rebuildAtRiskCache').timeBased().atHour(3).everyDays(1).inTimezone('Asia/Jerusalem').create();
  return 'Nightly at-risk trigger installed (~03:00 Asia/Jerusalem).';
}

// Dashboard read — fast. Reads the pre-computed cache and filters by the caller's
// scope: commanders see by level (global / their site / their managed sites),
// a regular teacher sees ONLY their own classes (by teacher ת.ז.). No model
// build, no big-sheet scan of תוצאות/תוצאות תרגול on the request path.
function handleTeacherAtRiskList(p) {
  var tData = getSheet('מורים').getDataRange().getValues();
  var role = '', userSite = '', myId = normalizeId(p.teacherId), found = false;
  for (var i = 1; i < tData.length; i++) {
    if (normalizeId(tData[i][1]) === myId) { role = String(tData[i][8] || 'מורה'); userSite = String(tData[i][9] || ''); found = true; break; }
  }
  if (!found) return jsonResponse({ status: 'error', message: 'מורה לא נמצא' });

  var isGlobal = (role === 'מפקד' || role === 'מפקד ראשי' || role === 'אדמין');
  var isLocal = (role === 'מפקד מקומי');
  var isMultiSite = isKdtzRole(role);
  var isCommander = isGlobal || isLocal || isMultiSite;
  var managedSites = [];
  if (isMultiSite) {
    managedSites = String(userSite || '').split(',').map(function(s) { return s.trim(); }).filter(function(s) { return s; });
  }

  var props = PropertiesService.getScriptProperties();
  var computedAt = props.getProperty('atRisk_computedAt') || null;
  var modelBaseRate = Number(props.getProperty('atRisk_modelBaseRate') || 0);
  var buildSummary = null;
  try { buildSummary = JSON.parse(props.getProperty('atRisk_summary') || 'null'); } catch (eBS) { buildSummary = null; }

  var sheet = getSheetIfExists('חיזוי סיכון');
  if (!sheet || sheet.getLastRow() < 2) {
    return jsonResponse({ status: 'ok', data: {
      computedAt: computedAt, notComputed: true,
      summary: { high: 0, medium: 0, low: 0, total: 0 }, students: []
    } });
  }
  var rows = sheet.getDataRange().getValues();
  var out = [];
  var summary = { high: 0, medium: 0, low: 0, total: 0 };
  for (var r2 = 1; r2 < rows.length; r2++) {
    var row = rows[r2];
    var site = String(row[7] || '');
    var teacherId = normalizeId(row[4]);
    if (isCommander) {
      if (isLocal && userSite && site !== userSite) continue;
      if (isMultiSite && managedSites.indexOf(site) === -1) continue;
      // isGlobal → no filter
    } else {
      if (!teacherId || teacherId !== myId) continue;   // regular teacher → own classes only
    }
    var tier = String(row[14] || 'low');
    if (tier === 'high') summary.high++; else if (tier === 'medium') summary.medium++; else summary.low++;
    summary.total++;
    out.push({
      name: String(row[1] || ''), license: String(row[2] || ''), className: String(row[6] || ''),
      teacherName: String(row[5] || ''), site: site,
      lastPct: Number(row[8]) || 0, sessions: Number(row[9]) || 0, trend: String(row[10] || ''),
      attempt: Number(row[11]) || 1, everTested: (row[12] === true || String(row[12]).toUpperCase() === 'TRUE'),
      prob: (row[13] === '' || row[13] == null) ? null : Number(row[13]),
      tier: tier, confidence: String(row[15] || 'low'),
      matchedByPhone: (row[16] === true || String(row[16]).toUpperCase() === 'TRUE')
    });
  }
  out.sort(function(a, b) {
    var pa = a.prob == null ? 999 : a.prob, pb = b.prob == null ? 999 : b.prob;
    if (pa !== pb) return pa - pb;
    if (a.lastPct !== b.lastPct) return a.lastPct - b.lastPct;
    return b.sessions - a.sessions;
  });

  return jsonResponse({ status: 'ok', data: {
    computedAt: computedAt,
    modelBaseRate: modelBaseRate,
    scope: isCommander ? (isGlobal ? 'כל האתרים' : (isLocal ? userSite : managedSites.join(', '))) : 'הכיתות שלי',
    buildSummary: buildSummary,
    coverageNote: 'תלמיד מזוהה לפי טלפון אם הוזן, אחרת לפי שם+דרגה. ניסיון = מספר הניסיון הצפוי במבחן האמיתי.',
    summary: summary,
    truncated: out.length > 200,
    students: out.slice(0, 200)
  } });
}

// ========== Examiner-commander forecast (predictive PLANNING view) ==========
// Reads the nightly at-risk cache (NO model build on request) and produces two
// forecasts for the examiner-commander:
//   cohortForecast (B') — aggregate expected pass rate of the current practicing
//     pool, by license + overall ("a typical upcoming C1 cohort → ~X% pass").
//   examDayForecast (A') — joins the LIVE registrants (ממתינים in active sessions)
//     to their nightly prediction → "of N registered now, ~X expected to pass,
//     Y at risk → prepare re-exam slots". No-practice-record registrants are
//     counted separately, never silently assumed pass/fail.
function handleExaminerForecast(p) {
  var exData = getSheet('בוחנים').getDataRange().getValues();
  var role = '';
  for (var i = 1; i < exData.length; i++) {
    if (normalizeId(exData[i][1]) === normalizeId(p.examinerId)) { role = String(exData[i][5] || 'בוחן'); break; }
  }
  if (role !== 'מפקד') return jsonResponse({ status: 'error', message: 'אין הרשאת מפקד' });

  var props = PropertiesService.getScriptProperties();
  var computedAt = props.getProperty('atRisk_computedAt') || null;
  var sheet = getSheetIfExists('חיזוי סיכון');
  if (!sheet || sheet.getLastRow() < 2) {
    return jsonResponse({ status: 'ok', data: { computedAt: computedAt, notComputed: true, cohortForecast: null, examDayForecast: null } });
  }
  var rows = sheet.getDataRange().getValues();
  // Cache columns: 1 name, 2 license, 7 site, 8 lastPct, 11 attempt, 13 prob, 14 tier, 17 phone.
  var byPhone = {}, byName = {};
  var cohortByLic = {};
  var cohortBySite = {};
  var cohortBySiteLic = {};
  var cohortAll = { n: 0, sumProb: 0, high: 0, medium: 0, low: 0 };
  function cohortBump(o, prob, tier) {
    o.n++; if (prob != null) o.sumProb += prob;
    if (tier === 'high') o.high++; else if (tier === 'medium') o.medium++; else o.low++;
  }
  for (var r = 1; r < rows.length; r++) {
    var row = rows[r];
    var lic = String(row[2] || '');
    var prob = (row[13] === '' || row[13] == null) ? null : Number(row[13]);
    var tier = String(row[14] || 'low');
    var phone = String(row[17] || '');
    var rec = { name: String(row[1] || ''), license: lic, lastPct: Number(row[8]) || 0, attempt: Number(row[11]) || 1, prob: prob, tier: tier, site: String(row[7] || '') };
    if (phone) byPhone[phone] = rec;
    var nk = ppNormName(row[1]) + '|' + lic;
    if (!byName[nk]) byName[nk] = rec;
    if (!cohortByLic[lic]) cohortByLic[lic] = { n: 0, sumProb: 0, high: 0, medium: 0, low: 0 };
    cohortBump(cohortByLic[lic], prob, tier);
    cohortBump(cohortAll, prob, tier);
    var cSite = String(row[7] || '') || 'לא צוין';
    if (!cohortBySite[cSite]) cohortBySite[cSite] = { n: 0, sumProb: 0, high: 0, medium: 0, low: 0 };
    cohortBump(cohortBySite[cSite], prob, tier);
    var clKey = cSite + '|' + (lic || 'לא צוין');
    if (!cohortBySiteLic[clKey]) cohortBySiteLic[clKey] = { site: cSite, license: (lic || 'לא צוין'), n: 0, sumProb: 0, high: 0, medium: 0, low: 0 };
    cohortBump(cohortBySiteLic[clKey], prob, tier);
  }
  function finalizeCohort(o) {
    return { n: o.n, expectedPassRate: o.n ? Math.round(o.sumProb / o.n) : 0, expectedPasses: Math.round(o.sumProb / 100), high: o.high, medium: o.medium, low: o.low };
  }
  var cohortForecast = { overall: finalizeCohort(cohortAll), byLicense: {}, bySite: {}, bySiteLicense: [] };
  for (var lk in cohortByLic) cohortForecast.byLicense[lk] = finalizeCohort(cohortByLic[lk]);
  for (var csk in cohortBySite) cohortForecast.bySite[csk] = finalizeCohort(cohortBySite[csk]);
  for (var clk in cohortBySiteLic) {
    var cl = cohortBySiteLic[clk], cf2 = finalizeCohort(cl);
    cohortForecast.bySiteLicense.push({ site: cl.site, license: cl.license, n: cf2.n, expectedPassRate: cf2.expectedPassRate, expectedPasses: cf2.expectedPasses, high: cf2.high, medium: cf2.medium, low: cf2.low });
  }
  cohortForecast.bySiteLicense.sort(function(a, b) { return a.site === b.site ? (b.n - a.n) : (a.site < b.site ? -1 : 1); });

  // A' — live registrants in ACTIVE sessions only. Track each session's site so
  // the forecast can be split per site (examiner-allocation planning).
  var activeSessions = {}, sessionSite = {};
  try {
    diagMark('sheet:sessions-forecast');
    var sess = sessionRows();
    var nowT = new Date().getTime();
    for (var s = 1; s < sess.length; s++) {
      var active = sess[s][10] === true || String(sess[s][10]).toUpperCase() === 'TRUE';
      var validUntil = sess[s][9] ? new Date(sess[s][9]).getTime() : 0;
      if (active && (!validUntil || validUntil > nowT)) {
        var scode = String(sess[s][0] || '').trim();
        activeSessions[scode] = true;
        sessionSite[scode] = String(sess[s][3] || '');   // אתר
      }
    }
  } catch (eS) { /* no sessions → examDay stays empty */ }

  var examDay = { registered: 0, matched: 0, noRecord: 0, expectedPasses: 0, expectedFails: 0, high: 0, medium: 0, low: 0, atRisk: [], bySite: {}, bySiteLicense: [] };
  var edSiteLicAcc = {};
  function edSite(site) {
    if (!examDay.bySite[site]) examDay.bySite[site] = { registered: 0, matched: 0, noRecord: 0, expectedPasses: 0, expectedFails: 0, high: 0, medium: 0, low: 0 };
    return examDay.bySite[site];
  }
  function edSiteLic(site, lic) {
    var kk = site + '|' + lic;
    if (!edSiteLicAcc[kk]) edSiteLicAcc[kk] = { site: site, license: lic, registered: 0, matched: 0, noRecord: 0, expectedPasses: 0, high: 0, medium: 0, low: 0 };
    return edSiteLicAcc[kk];
  }
  try {
    // Only registrations of sessions that are still open matter, and a session
    // lives 8 hours — so two days of rows cover every one of them. Columns:
    // A code, C name, D phone, F status, I licence.
    diagMark('sheet:pending-forecast');
    var waitCutoff = new Date(Date.now() - 2 * 86400000);
    var wait = readRowsSince(getSheet('ממתינים'), 4, waitCutoff, [[1, 1], [3, 2], [6, 1], [9, 1]]).rows;
    for (var w = 1; w < wait.length; w++) {
      var code = String(wait[w][0] || '').trim();
      if (!activeSessions[code]) continue;
      var status = String(wait[w][5] || '');
      if (status === 'הושלם' || status === 'פסול' || status === 'בוטל' || status === 'נדחה') continue;   // already resolved
      var siteA = sessionSite[code] || 'לא צוין';
      var bs = edSite(siteA);
      examDay.registered++; bs.registered++;
      var wPhone = ppNormPhone(wait[w][3]);
      var wLic = String(wait[w][8] || '');
      var wName = String(wait[w][2] || '');
      var sl = edSiteLic(siteA, wLic || 'לא צוין'); sl.registered++;
      var hit = (wPhone && byPhone[wPhone]) || byName[ppNormName(wName) + '|' + wLic] || null;
      if (!hit || hit.prob == null) { examDay.noRecord++; bs.noRecord++; sl.noRecord++; continue; }
      examDay.matched++; bs.matched++; sl.matched++;
      examDay.expectedPasses += hit.prob / 100; bs.expectedPasses += hit.prob / 100; sl.expectedPasses += hit.prob / 100;
      if (hit.tier === 'high') { examDay.high++; bs.high++; sl.high++; } else if (hit.tier === 'medium') { examDay.medium++; bs.medium++; sl.medium++; } else { examDay.low++; bs.low++; sl.low++; }
      if (hit.tier === 'high' || hit.tier === 'medium') examDay.atRisk.push({ name: wName, license: wLic, site: siteA, prob: hit.prob, tier: hit.tier, lastPct: hit.lastPct, attempt: hit.attempt });
    }
  } catch (eW) { /* no waiting sheet */ }
  examDay.expectedPasses = Math.round(examDay.expectedPasses);
  examDay.expectedFails = Math.max(0, examDay.matched - examDay.expectedPasses);
  for (var bsk in examDay.bySite) {
    var b2 = examDay.bySite[bsk];
    b2.expectedPasses = Math.round(b2.expectedPasses);
    b2.expectedFails = Math.max(0, b2.matched - b2.expectedPasses);
  }
  for (var slk in edSiteLicAcc) {
    var sl2 = edSiteLicAcc[slk];
    sl2.expectedPasses = Math.round(sl2.expectedPasses);
    sl2.expectedFails = Math.max(0, sl2.matched - sl2.expectedPasses);
    examDay.bySiteLicense.push(sl2);
  }
  examDay.bySiteLicense.sort(function(a, b) { return a.site === b.site ? (b.registered - a.registered) : (a.site < b.site ? -1 : 1); });
  examDay.atRisk.sort(function(a, b) { return a.prob - b.prob; });
  examDay.atRisk = examDay.atRisk.slice(0, 50);

  return jsonResponse({ status: 'ok', data: {
    computedAt: computedAt,
    cohortForecast: cohortForecast,
    examDayForecast: examDay
  } });
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
  var deletedClassMap = getDeletedClassMap();

  // Read practice results - INCLUDING rows without classCode.
  // Bounded by the requested range; columns A-F (date, student, name, class,
  // mode, licence) + I (percent) + J (pass) — never the two JSON blobs.
  // ⚠ Index another column here and add it to the colSpec.
  diagMark('sheet:practice-admin');
  var adminRead = readRowsSince(getSheet('תוצאות תרגול'), 0, dateFrom, [[1, 6], [9, 2]]);
  var resData = adminRead.rows;
  diagMark('sheet:practice-admin-done:' + adminRead.mode);

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
    var cInfo = classCode ? resolveClassInfo(classCode, classMap, deletedClassMap) : null;

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
  var maintenance = practiceWriteGuard(); if (maintenance) return maintenance;   // r24
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
  var maintenance = practiceWriteGuard(); if (maintenance) return maintenance;   // r24
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
  var maintenance = practiceWriteGuard(); if (maintenance) return maintenance;   // r24
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

  // Archive the class metadata BEFORE deleting the row. Practice results in
  // 'תוצאות תרגול' are preserved for history (see NOTE below), and once the class
  // row is gone the reports can no longer resolve its teacher/name → they showed
  // "לא ידוע". This archive lets the commander dashboard still attribute those
  // orphaned rows to the real teacher and tag them "(כיתה שנמחקה)", so a genuine
  // deletion is distinguishable from a truly unrecognized/forged class code.
  try {
    var cRow = classData[classRowIdx];
    getSheet('כיתות שנמחקו').appendRow([
      String(cRow[0] || '').trim(), // קוד כיתה
      String(cRow[1] || ''),        // שם כיתה
      normalizeId(cRow[2]),         // מורה ת.ז.
      String(cRow[3] || ''),        // שם מורה
      String(cRow[4] || ''),        // דרגה
      String(cRow[7] || ''),        // אתר
      nowISO()                      // תאריך מחיקה
    ]);
  } catch (archiveErr) { /* non-fatal: deletion proceeds even if archiving fails */ }

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
  var maintenance = practiceWriteGuard(); if (maintenance) return maintenance;   // r24
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
  var classInfo = null, classCreated = null;
  for (var c = 1; c < classData.length; c++) {
    if (String(classData[c][0]).trim() === classCode && normalizeId(classData[c][2]) === normalizeId(p.teacherId)) {
      classInfo = { code: classCode, name: classData[c][1], license: classData[c][4], active: classData[c][6] === 'כן' };
      classCreated = parseSheetDateTime(classData[c][5]);   // col F = תאריך יצירה
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

  // Get practice results for these students.
  // 17/09/2026: this ran 13 times in one exam window at 19-32 s each, every call
  // a FULL read of the 107k-row practice sheet including its two JSON blob
  // columns. Two bounds, both exact:
  //   rows    — nothing before the class existed can belong to the class;
  //   columns — everything except N ('פירוט שגויות', ~2 KB of wrong-question
  //             text per row) which nothing on this screen shows. Column O
  //             ('פירוט לפי נושא') STAYS: it is the per-topic breakdown the
  //             student card draws, and it is two orders of magnitude smaller.
  // ⚠ Index another column here and add it to the colSpec.
  diagMark('sheet:practice-class');
  var classRead = readRowsSince(getSheet('תוצאות תרגול'), 0, classCreated, [[1, 13], [15, 1]]);
  var resData = classRead.rows;
  diagMark('sheet:practice-class-done:' + classRead.mode);
  var inClass = {};
  for (var sj = 0; sj < studentIds.length; sj++) inClass[studentIds[sj]] = true;
  var studentResults = {};
  for (var r = 1; r < resData.length; r++) {
    var rSid = String(resData[r][1]).trim();
    if (String(resData[r][3]).trim() !== classCode || !inClass[rSid]) continue;
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
      categoryBreakdown: resData[r][14] || ''
    });
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
  var owns = false, exportCreated = null;
  for (var c = 1; c < classData.length; c++) {
    if (String(classData[c][0]).trim() === classCode && normalizeId(classData[c][2]) === normalizeId(p.teacherId)) {
      owns = true;
      exportCreated = parseSheetDateTime(classData[c][5]);
      break;
    }
  }
  if (!owns) return jsonResponse({ status: 'error', message: 'אין הרשאה' });

  // Get all results for this class. Rows are bounded by the class's creation
  // date (nothing older can belong to it); every COLUMN stays — this is the
  // export, and the JSON blobs are the point of it.
  var resData = readRowsSince(getSheet('תוצאות תרגול'), 0, exportCreated).rows;
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
  var maintenance = practiceWriteGuard(); if (maintenance) return maintenance;   // r24
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
  var maintenance = practiceWriteGuard(); if (maintenance) return maintenance;   // r24
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
  var maintenance = practiceWriteGuard(); if (maintenance) return maintenance;   // r24
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
