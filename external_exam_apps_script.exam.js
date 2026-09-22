// ============================================================================
// GENERATED FILE — do not edit here. Source: server/src/*.js (order: server/BUILD_ORDER.json,
// split: server/BUILD_TARGETS.json). Comments are stripped and the question index is packed
// by the build; read the SOURCE, not this.
// Rebuild with:  node tools/build_server.js     (tools/build.js runs it too)
// Deployment: exam   (API_DEPLOYMENT — health reports it)
// Deploy (split, step ג): paste this whole file into the EXISTING Apps Script project — the
// URL every page already knows → New version. health must answer "deployment":"exam".
// Run uninstallNightlyJobs() once here: the nightly jobs move to the reports project.
// A reports action that reaches this deployment is answered wrong_deployment, not a broken page.
// Modules: 00_config, 05_registry, 10_spreadsheet, 08_pending_writes, 12_reads, 16_lookup_util, 18_whatsapp, 20_auth, 22_util, 24_triggers, 70_questions, 30_api, 40_sessions_login, 42_sessions_manage, 44_sessions_misc, 50_pending, 55_dashboard, 52_pending_status, 65_dq, 60_exam, 75_diag
// ============================================================================
// © 2026 Vitaly Gitelman. All Rights Reserved.
// Unauthorized copying, modification or distribution is prohibited.
// ===== Google Apps Script — מערכת בחינות חיצונית =====
// הדבק את הקוד הזה ב-Apps Script של גיליון Google Sheets חדש
// Deploy → New deployment → Web app
// Execute as: Me | Who has access: Anyone
// העתק את ה-URL שמקבלים והדבק ב-examiner.html וב-examinee.html


var PENDING_ARCHIVE_SHEET = 'ממתינים_ארכיון';
var RESULTS_ARCHIVE_SHEET = 'תוצאות_ארכיון';
var EXAMS_ARCHIVE_SHEET = 'מבחנים_ארכיון';

var SHEET_HEADERS = {
  'בוחנים': ['שם', 'ת.ז.', 'סיסמה', 'פעיל', 'מס בוחן', 'תפקיד', 'טוקן', 'תוקף טוקן', 'ניסיונות כושלים', 'נעילה עד', 'אתרים מנוהלים'],
  'אתרים': ['שם אתר', 'מזהה', 'טלפון מנהל', 'כיתות'],
  'סשנים': ['קוד', 'בוחן ת.ז.', 'שם בוחן', 'אתר', 'כיתה', 'דרגה', 'שפה', 'מצב שמע', 'זמן יצירה', 'תקף עד', 'פעיל', 'כמויות JSON', 'מאושרים JSON', 'בוחן אחראי', 'אוכלוסיית ברירת מחדל'],
  'ממתינים': ['קוד סשן', 'ת.ז.', 'שם', 'טלפון', 'זמן הרשמה', 'סטטוס', 'שפה', 'אוכלוסיה', 'דרגה', 'שמע', 'הארכת זמן', 'התחלת מבחן', 'טוקן נבחן', 'ספירת DQ', 'מסך נוסף', 'ספירת אזהרות', 'אזהרה אחרונה', 'אתר', 'סיים במכשיר'],
  'ממתינים_ארכיון': ['קוד סשן', 'ת.ז.', 'שם', 'טלפון', 'זמן הרשמה', 'סטטוס', 'שפה', 'אוכלוסיה', 'דרגה', 'שמע', 'הארכת זמן', 'התחלת מבחן', 'טוקן נבחן', 'ספירת DQ', 'מסך נוסף', 'ספירת אזהרות', 'אזהרה אחרונה', 'אתר', 'סיים במכשיר'],
  'תוצאות': ['תאריך', 'ת.ז.', 'שם', 'טלפון', 'דרגה', 'ציון', 'אחוז', 'עבר/נכשל', 'זמן', 'בוחן', 'אתר', 'כיתה', 'שפה', 'קוד סשן', 'ניסיון', 'פירוט שגויות', 'נשלח?', 'פסול?', 'קישור וואטסאפ', 'אוכלוסיה', 'תוקן?', 'שמע', 'מאומת', 'חשוד', 'dqEventId', 'תוקן ע"י', 'סיבת תיקון', 'תאריך תיקון', 'מסלול שפות', 'מכשיר'],
  'תוצאות_ארכיון': ['תאריך', 'ת.ז.', 'שם', 'טלפון', 'דרגה', 'ציון', 'אחוז', 'עבר/נכשל', 'זמן', 'בוחן', 'אתר', 'כיתה', 'שפה', 'קוד סשן', 'ניסיון', 'פירוט שגויות', 'נשלח?', 'פסול?', 'קישור וואטסאפ', 'אוכלוסיה', 'תוקן?', 'שמע', 'מאומת', 'חשוד', 'dqEventId', 'תוקן ע"י', 'סיבת תיקון', 'תאריך תיקון', 'מסלול שפות', 'מכשיר'],
  'מבחנים': ['קוד סשן', 'ת.ז.', 'שאלות JSON', 'זמן רישום', 'שפה', 'שגויות לא מאומתות'],
  'מבחנים_ארכיון': ['קוד סשן', 'ת.ז.', 'שאלות JSON', 'זמן רישום', 'שפה', 'שגויות לא מאומתות'],
  'הארכות זמן': ['תאריך', 'קוד סשן', 'ת.ז.', 'שם', 'דקות', 'סיבה', 'בוחן'],
  'מורים': ['שם', 'ת.ז.', 'סיסמה', 'פעיל', 'טוקן', 'תוקף טוקן', 'ניסיונות כושלים', 'נעילה עד', 'תפקיד', 'אתר'],
  'כיתות': ['קוד כיתה', 'שם כיתה', 'מורה ת.ז.', 'שם מורה', 'דרגה', 'תאריך יצירה', 'פעיל', 'אתר'],
  'כיתות שנמחקו': ['קוד כיתה', 'שם כיתה', 'מורה ת.ז.', 'שם מורה', 'דרגה', 'אתר', 'תאריך מחיקה'],
  'תלמידי כיתות': ['קוד כיתה', 'שם תלמיד', 'מזהה תלמיד', 'תאריך הצטרפות'],
  'תוצאות תרגול': ['תאריך', 'מזהה תלמיד', 'שם תלמיד', 'קוד כיתה', 'מצב', 'דרגה', 'ציון', 'סה"כ', 'אחוז', 'עבר/נכשל', 'זמן', 'נושא', 'שפה', 'פירוט שגויות', 'פירוט לפי נושא', 'טלפון'],
  'התקדמות תלמידים': ['שם תלמיד', 'קוד כיתה', 'מפתח', 'streak', 'wrong_qs', 'history', 'עדכון אחרון'],
  'חיזוי סיכון': ['חושב בתאריך', 'שם', 'דרגה', 'קוד כיתה', 'מורה ת.ז.', 'שם מורה', 'שם כיתה', 'אתר', 'ציון תרגול', 'תרגולים', 'מגמה', 'ניסיון צפוי', 'ניגש בעבר', 'סיכוי מעבר', 'רמת סיכון', 'ביטחון', 'זוהה בטלפון', 'טלפון']
};

var TEST_SITES = ['בדיקת נתונים', 'דימונה דוגית 35'];
function isTestSite(site) {
  return TEST_SITES.indexOf(String(site || '').trim()) !== -1;
}

function normalizeNameKey(s) {
  if (!s) return '';
  var t = String(s).replace(/[׳״'".\-]/g, ' ').replace(/\s+/g, ' ').trim().toLowerCase();
  if (!t) return '';
  var tokens = t.split(' ').filter(function(x) { return x; });
  tokens.sort();
  return tokens.join(' ');
}

function getExaminerExclusion() {
  var names = {}, ids = {};
  try {
    var d = getSheet('בוחנים').getDataRange().getValues();
    for (var i = 1; i < d.length; i++) {
      var nk = normalizeNameKey(d[i][0]);
      if (nk) names[nk] = true;
      var ik = normalizeId(d[i][1]);
      if (ik) ids[ik] = true;
    }
  } catch (e) {}
  return { names: names, ids: ids };
}
function isExaminerSelfTest(name, id, excl) {
  if (!excl) return false;
  var ik = normalizeId(id);
  if (ik && excl.ids[ik]) return true;
  var nk = normalizeNameKey(name);
  return !!(nk && excl.names[nk]);
}

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
var EXAM_SPREADSHEET_PROPERTY = 'EXAM_SPREADSHEET_ID';
var _spreadsheetHandle = null;
function getSpreadsheet() {
  if (!_spreadsheetHandle) {
    _spreadsheetHandle = SpreadsheetApp.getActiveSpreadsheet() || openExamSpreadsheetById();
    diagMark('ss:open');
  }
  return _spreadsheetHandle;
}
function openExamSpreadsheetById() {
  var id = '';
  try { id = String(PropertiesService.getScriptProperties().getProperty(EXAM_SPREADSHEET_PROPERTY) || '').trim(); }
  catch (e) { id = ''; }
  if (!id) {
    throw new Error('EXAM_SPREADSHEET_ID is not set — standalone deployment needs the exam spreadsheet id');
  }
  return SpreadsheetApp.openById(id);
}

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

var PENDING_STATUS_COL = 6;
var PENDING_COLS = { status: 6, language: 7, population: 8, license: 9, audio: 10, timeExtension: 11, examStart: 12, token: 13, dqCount: 14, extScreen: 15, warnCount: 16, lastWarning: 17, site: 18, finishedOnDevice: 19 };
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

function markPendingCompleted(sessionCode, idNumber, pendingSnapshot) {
  var pendSheet = pendingSnapshot ? pendingSnapshot.sheet : getSheet('ממתינים');
  var pendData = pendingSnapshot
    ? refreshExamineePendingRows(pendSheet, pendingSnapshot.rows, sessionCode, idNumber)
    : pendSheet.getDataRange().getValues();
  var wrote = false;
  for (var j = pendData.length - 1; j >= 1; j--) {
    if (String(pendData[j][0]) === String(sessionCode) && normalizeId(pendData[j][1]) === normalizeId(idNumber) && (String(pendData[j][5]).trim() === 'in_exam' || String(pendData[j][5]).trim() === 'approved')) {
      pendSheet.getRange(j + 1, 6).setValue('completed');
      pendData[j][5] = 'completed';
      wrote = true;
    }
  }
  if (wrote) {
    SpreadsheetApp.flush();
    invalidatePendingSnapshot(sessionCode);
  }
}

var TAIL_ROWS = 1000;
var TAIL_MAX_AGE_HOURS = 48;

function parseSheetDateTime(v) {
  if (!v) return null;
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
  var s = String(v).trim();
  var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2}))?/);
  if (m) return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]), Number(m[4] || 0), Number(m[5] || 0));
  var d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

function readTail(sheet, tsColIdx) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow - 1 <= TAIL_ROWS || lastCol < 1) {
    var small = sheet.getDataRange().getValues();
    diagMark('tail:small/' + lastRow);
    return { rows: small, off: 0 };
  }
  var startRow = lastRow - TAIL_ROWS + 1;
  var tail = sheet.getRange(startRow, 1, TAIL_ROWS, lastCol).getValues();
  var oldest = parseSheetDateTime(tail[0][tsColIdx]);
  if (!oldest || (Date.now() - oldest.getTime()) < TAIL_MAX_AGE_HOURS * 3600 * 1000) {
    var full = sheet.getDataRange().getValues();
    diagMark('tail:full/' + lastRow);
    return { rows: full, off: 0 };
  }
  var header = sheet.getRange(1, 1, 1, lastCol).getValues();
  diagMark('tail:' + TAIL_ROWS + '/' + lastRow);
  return { rows: header.concat(tail), off: startRow - 2 };
}

function readPendingTail() { return readTail(getSheet('ממתינים'), 4); }
function readResultsTail() { return readTail(getSheet('תוצאות'), 0); }

var PENDING_SNAPSHOT_SEC = 4;
var PENDING_SNAPSHOT_PREFIX = 'pendsnap_';
function pendingSnapshotKey(sessionCode) {
  return CACHE_KEY_PREFIX + PENDING_SNAPSHOT_PREFIX + String(sessionCode || '').trim();
}
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
      while (row.length < colSpec[j][0] - 1) row.push('');
      var piece = parts[j] ? parts[j][r] : null;
      if (piece) for (var k = 0; k < piece.length; k++) row.push(piece[k]);
    }
    out.push(row);
  }
  return out;
}

function firstRowSince(sheet, tsColIdx, cutoff, lastRow) {
  var col = sheet.getRange(2, tsColIdx + 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < col.length; i++) {
    var d = parseSheetDateTime(col[i][0]);
    if (!d || d.getTime() >= cutoff.getTime()) return i + 2;
  }
  return 0;
}

function readRowsSince(sheet, tsColIdx, cutoff, colSpec) {
  var lastRow = sheet.getLastRow(), lastCol = sheet.getLastColumn(), dataRows = lastRow - 1;
  var bounded = cutoff instanceof Date && !isNaN(cutoff.getTime());
  if (bounded && lastCol >= 1 && dataRows > TAIL_ROWS) {
    var startRow = 0;
    try { startRow = firstRowSince(sheet, tsColIdx, cutoff, lastRow); } catch (eScan) { startRow = -1; }
    if (startRow === 0) {
      return { rows: readSheetSlice(sheet, 1, 1, lastCol, colSpec), off: 0, mode: 'none/' + lastRow };
    }
    if (startRow > 2) {
      var n = lastRow - startRow + 1;
      var tail = readSheetSlice(sheet, startRow, n, lastCol, colSpec);
      var header = readSheetSlice(sheet, 1, 1, lastCol, colSpec);
      return { rows: header.concat(tail), off: startRow - 2, mode: 'rows' + n + '/' + lastRow };
    }
  }
  if (colSpec && lastCol >= 1) return { rows: readSheetSlice(sheet, 1, lastRow, lastCol, colSpec), off: 0, mode: 'full/' + lastRow };
  return { rows: sheet.getDataRange().getValues(), off: 0, mode: 'full/' + lastRow };
}

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
function examinerOwnsSession(sessionCode, examinerId) {
  if (!examinerId) return false;
  var row = sessionRowByCode(sessionCode);
  return !!row && normalizeId(row[1]) === normalizeId(examinerId);
}
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
    existingCodes[String(data[i][0]).trim()] = true;
  }
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

var OFFICE_WHATSAPP_NUMBER_DEFAULT = '0529151157';
function getOfficeWhatsAppNumber() {
  try {
    var prop = PropertiesService.getScriptProperties().getProperty('OFFICE_WHATSAPP_NUMBER');
    if (prop && String(prop).trim()) return String(prop).trim();
  } catch (e) {}
  return OFFICE_WHATSAPP_NUMBER_DEFAULT;
}

function generateToken() {
  return (Utilities.getUuid() + Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
}

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
  diagMark('auth:token');
  if (!valid) {
    return jsonResponse({ status: 'error', message: 'טוקן לא תקין — יש להתחבר מחדש', tokenExpired: true });
  }
  return null;
}

var ALLOWED_ORIGINS = [
  'examiner-app',
  'examinee-app',
  'teacher-app',
  'student-app',
  'admin-app',
  'bohanyzahal-site',
  'gateway',
  'localhost-dev'
];
function checkOrigin(p) {
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
    cache.put(key, JSON.stringify(fresh), windowSeconds + 60);
    return { ok: true };
  } catch (e) {
    return { ok: true };
  }
}

function requireRateLimit(action, identifier, maxRequests, windowSeconds) {
  if (!identifier) return null;
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

function generateExamineeToken() {
  return (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
}

function examineeTokenVerdict(row, examineeToken) {
  var rowAudio = String(row[9] || '').trim() === 'on' ? 'on' : 'off';
  var storedToken = String((row.length > 12 ? row[12] : '') || '').trim();
  if (!storedToken) return { valid: true, legacy: true, audioMode: rowAudio };
  if (!examineeToken) return { valid: false, reason: 'missing' };
  if (String(examineeToken).trim() === storedToken) return { valid: true, legacy: false, audioMode: rowAudio };
  return { valid: false, reason: 'mismatch' };
}

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

function verifyExamineeToken(sessionCode, idNumber, examineeToken) {
  var ctx = examineeRowContext(sessionCode, idNumber, true);
  if (!ctx.latest) return { valid: false, reason: 'not_found' };
  var verdict = examineeTokenVerdict(ctx.latest.row, examineeToken);
  ctx.audioMode = verdict.audioMode || 'off';
  return verdict;
}

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

function verifyExaminerForSession(sessionCode, examinerId) {
  return examinerOwnsSession(sessionCode, examinerId);
}


function gatewayKey() {
  try { return String(PropertiesService.getScriptProperties().getProperty('GATEWAY_KEY') || ''); }
  catch (e) { return ''; }
}

var BANK_GRANT_TTL_MS = { exam: 4 * 3600 * 1000, practice: 2 * 3600 * 1000, examiner: 8 * 3600 * 1000 };
var BANK_GRANT_SUB_MAX = 128;

function bankGrantJson(obj) {
  return JSON.stringify(obj).replace(/[\u0080-\uFFFF]/g, function(ch) {
    return '\\u' + ('000' + ch.charCodeAt(0).toString(16)).slice(-4);
  });
}

function signBankGrant(payloadObj, key) {
  var secret = key || gatewayKey();
  var payloadB64 = Utilities.base64EncodeWebSafe(bankGrantJson(payloadObj)).replace(/=+$/, '');
  var sigB64 = Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(payloadB64, secret)).replace(/=+$/, '');
  return payloadB64 + '.' + sigB64;
}

function bankGrantConfigured() { return Boolean(gatewayUrl() && gatewayKey()); }

function bankGrantFor(scope, ids, sub) {
  var ttl = BANK_GRANT_TTL_MS[String(scope)];
  var url = gatewayUrl(), key = gatewayKey();
  if (!ttl || !url || !key) return null;
  var exp = Date.now() + ttl;
  var payload = { v: 1, s: String(scope) };
  if (String(scope) !== 'examiner') {
    var list = [];
    for (var i = 0; ids && i < ids.length; i++) list.push(Number(ids[i]));
    payload.ids = list;
  }
  payload.sub = String(sub || '').slice(0, BANK_GRANT_SUB_MAX);
  payload.exp = exp;
  return { url: url, grant: signBankGrant(payload, key), exp: exp };
}

function bankNotConfiguredResponse() {
  return jsonResponse({ status: 'error', code: 'bank_not_configured',
    message: 'מאגר השאלות אינו מוגדר בשרת — פנה למנהל המערכת' });
}

function gatewayUrl() {
  try { return String(PropertiesService.getScriptProperties().getProperty('GATEWAY_URL') || '').trim(); }
  catch (e) { return ''; }
}
var CACHE_KEY_PREFIX = 'qv2_';

var RESULTS_ATTEMPT_COLSPEC = [[2, 1], [5, 1], [8, 1]];
function countAttempts(idNumber, license, liveRows) {
  var wantId = normalizeId(idNumber), wantLic = String(license);
  var count = countAttemptRows(liveRows || readAttemptColumns(getSheet('תוצאות')), wantId, wantLic);
  var arch = getSheetIfExists(RESULTS_ARCHIVE_SHEET);
  if (arch) count += countAttemptRows(readAttemptColumns(arch), wantId, wantLic);
  return count;
}
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
    if (String(rows[i][7] || '').trim() === 'בוטל') continue;
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

var NIGHTLY_OBSOLETE_HANDLERS = ['archiveOldPendingRows', 'archiveSheets', 'warmupQuestionCaches',
  'ensureQuestionCachesWarm', 'rebuildMissingQuestionCaches', 'rebuildAtRiskCache'];

function uninstallNightlyJobs() {
  var wanted = NIGHTLY_OBSOLETE_HANDLERS.slice();
  var live = ['archiveSheets', 'rebuildAtRiskCache'];
  for (var l = 0; l < live.length; l++) if (wanted.indexOf(live[l]) === -1) wanted.push(live[l]);
  var trigs = ScriptApp.getProjectTriggers(), removed = [];
  for (var i = 0; i < trigs.length; i++) {
    var fn = trigs[i].getHandlerFunction();
    if (wanted.indexOf(fn) === -1) continue;
    ScriptApp.deleteTrigger(trigs[i]);
    removed.push(fn);
  }
  var msg = 'uninstallNightlyJobs: removed ' + removed.length + ' trigger(s) [' + removed.join(', ') +
    ']; ' + (trigs.length - removed.length) + ' other trigger(s) left untouched';
  Logger.log(msg);
  return msg;
}

var QUESTION_INDEX_PACKED = {
max: 1803,
lic: ["B","1","C1","C","D"],
topics: ["בטיחות","הכרת הרכב","חוק","תמרורים","ספציפי"],
rec: "333330333330333330303000333330333330333330333330333330333330333330333331333000333000444441333330333330333331333330005000333330303000333330333330005000033300000000333330005050333330333330333330333330333330005000050000333330000000303000303000333330333330111110333330333330303001303300333330000000333330000000333330000000303000303000333330333330333330333330333330333330333330333330333330333330333330333330333330333330333330333330333330333330333330333330333330333330333330300001303030333331333331333330333331333331333330333330050001333330333330303000333330444441333001333330333000333000333000333000333000333331333331333331333330333331333331333330333331333330333331333331333331333331333331333331333331333330333330333330333330333330333330333330333330333330333330333330333330333330333330333330333331333330333330333330033300333330333331333331111110333330111111333330111111111110333330303331333001333330303351303001303001444441333331333331333331333330333331303300333331333330333331333330333330333330333330333330333330333330333330333330333331333330333330333330333331333330333330111111444441444441444441444441444441111111444441333331111111333330333331333330333330333331111110101110303331303331333330333330333330303000333330333330000000333331333330333331333331333330333331333331333330333330333331444441333331333330333330333330333330333330333330333330303350303351303350303301333330303300303300303301303300303300303330303350303350333330303000000000333000303300333331033300303330050000444441333330333330333330333330000000033330333330333330333330333330111110333330333330303350303350005050033350333330333330303000303000303000333330333000333000333330101110333330333330333330333330333330333330333331000000333330333330333330333330333330005050333330000050333330333330005000000500303350333330303000333330333330333330333330111110111110333330333330333330333330333330333330111110444441111110111110111110111110111110111110111110111110111110005000000500005000333330005050005000005000333330330330333000333000303300333330300000050000333330333330333330333330303000050000333330333330333330303000333330333330333330333331444441333330333330303350303000303350444441444441333330444441005501444441444441444441444441333330444441333330444441444441444441444441444441444441444441444441444441444441444441444441333330444441444441444441444441444441444441444441444441333330333331444441444441333330444441444441444441444441444441303000333330404001444440444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441111110444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441444440333331333330333330333330444441000000444441444441444441444441444441444401444441444441444441444441444441444441444441444441400001005501444441444441400001444441444441400001444441000551000551444401005501000551444441444441444441000551333330000051444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441333330444441444441444441444441444441444441444441444441444441444441444441044441444441444441004401444441444441444441444441444441444441444441005501444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441000051444441444441444441444441444440444441444441444441444441111110000050000050000050000050444441444441444441303330000050444441444441444441444441444441111110444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441000051444441444441303001111110303350444441444441444441444441111110444441444441444441444441444441444441444441444441444441444441101110444441444441444441444441444441444441444441444441444441444441444441444441444441444441444441111110222220222220444441444440111110444441005501005501202220333331333331333331333331333331333331333331333331333331333331333331333331333331333331333331333331000000333331333331333331333331333331333331333331333331333330000050333331333331333331333331333331333331111111111110111110111110333330111110333330111000101010101010101110111010202220000000333330111111111111333330111110111110111110000000333330333330333330333330333330005000333330444441111110111110111110111110333330222220111110333330333330111111111111111111111110000000111111111110111111333330101110111110333330111110111110300000333331444441444441000051444441111110333330333330333330333330111111333330333331111110333330111110111110111110333330333330101110111110111110111110111110111110111110111110111110111110111110111111111110111110111110111110111110111110111110111110111110111110111111111110111110333330222220111110111111111110111111111110111110111110111110111111111110111110111111111110111110111110111110111110111110111110111110111110111110111110111110111110111110111110111110101110100010111110111110111110111110111110111110111110111110111110111110111110111110111110111110111110111110111110111111111111111111111111111111005051111111111111202220111111111111111111101111333330111111111110000000111111111111111110111111111111111111111111111111111110111111111111111111111101111111111111111111444441444441111111111111111111111111111111111111111111111111101111111111111111111111111111111111333330111111111111101110111111111111111111111111222220111111101001111111111111111111303001111111111111111111101111111111111111111111111111333330101110111111111111111111111111111111444441111111000000404451444441333331300001404451444441444441333331111111111111111111111111111111111111111111111111444441111111111111111111111111111111111111111111111111111111222221222220222220222000222000222000222000222000202000202000220000202000333330222220202000202000202000202000202000200000101110444441202000202000202000202000202000200000200000202000202000202000202000202000202000202000444441222000202000202000202000222220202000202000202000222000202000202000222000202000444441202000202000202000200000202000222000100000101000202000222000202000202000202000202000202000222000444441101000202000222000222000222000222000111110111110222000111000222000222000202000222000202000202000000050111110111110444441111110111110100000202000444441202000202000202000101000050000050000010000010000010000010000050000050000010000010000010000050000050000050000050000050000050000050000050000050000050000010000050000050000010000010000050000010000010000010000000050000050000050000050000050000050000050000050005050000050444441000050000050000050000050000050300030000050000050000050333330000050000050000050000050000050000050000050000050000050000050000050000050000050000050000050000050000050000050000050000050000050030050000050000050000050000050000050000050000050000050000050000050000050000050000050000050000050000050000050000050000050000050000050000050000050000050000050000050000050111110000050000050000050000050000050000050000050000050000050000050000050000050000050000050000050000050000050000050000050000050444441000050000050000050000050000050303030333330000050000050000050000050000050000050303350444441000050000050000050000050000050303000000050000050000050000050000050000050000050000050000050000050000050444441000050000050000050000050000050000050000050000050000050000050000050000050000050000050000050000050000050000050000050000051000051000051000051000051000051000051000051000051000051400051000051000051000051000051000051000050000050000050005501000050444441000500000050000050000050333330333330333330333330202000202000222000222000111110111110101110333330333330111110333330101110000500000500000500030300000500000500000500005000444441005000000500000500333330000500000500000500005000000500000050333330005000333300333331000500000500000500303350000500000500000500000500000500000500111110005000303000005000333330000500033350005000033300000050005000000500000500303350303350303350444441005000005000005000000500005000000500000500333330303300000500005000005050000500005000005050005050000500333330005000000050005000030330005050005050005000000500005000333000005050005050333330030300303350444441333330003300000500005000444441005050111110444441005001005000005000005050333330005050111110303330333330033350000500333300404450333000005000005000005000005000005000005000005000000500000500005000005000005000111110000500000500000500000500000500000500000500000500050000005000005000005000005000005000005000005000005000005050005000000500000500000500005000005000000500000500303330050000005050005000005000005000005000005000005000000050005000005000005000005000000050005000005000000500005000005050005000000050005050444441222220000500000500303000000050000500000500000500101110444441000500000500101110101110005000111110303330202000300000303330333330111110111110111110111110111110005000005000333330111110222000202000333330303331005000111110000500000500005000444441111110111110111110005000000000005000000500111110005000303001111110444441005000000500005000111110005000333330005000005000333330111110333330101110111110303300101110444441333330000500303300303300333330005000111110333330005000005000303300333330333330000050101110005000333330303350111110111110444441005000101110101110333330444440202000111110111110005050005050005050333330444441444441000000444441333330050000000000000000000000222000000000101000000000300330000000000000000000000000000000000000000000050000300000000000300000000000000000000000000000111110000000000000000000000000000000000000000000000000000051111110000000000000000000000000000000000000000000000000000000000000000000000000000000333330000000000000050000000000000000050000000000000000000000000000000000000000000000005000005001000051000000000000333330000000000000000000000000111110000000000000000000333000000000005001000000000000000000000000000000000000000000000000000000000000000000050000050000000000000000000000000000000000000501000501222220303000222220005050333330333330111110000000303000333330333330333330303000333000303350333001000000101110202220000050111110222000111110333330111110333330111110303350101110222220111110111110111110222220333330222001111110050000444441000050333330333330111110333000333330333330111110333330111110333330111111000050000050000050303001000000000051000000000051000000000000000000000000033330033330111110033330303330333330333330333330202000333330333330333330333330000050444440333330333330333330101110303000333330333330333330000050444441444440333330444440333330101000333330333330333330444441000500000500444440111110333330333330444440333330222220333330333330333330333330333330333330333330000050444440333330202000111110111110333330000050000050050000000000333330333330000501000500444441000000444441",
lang: {"202":4,"216":4,"248":4,"266":4,"574":4,"746":4,"907":125},
langDefault: 127
};

function questionIndex() {
  if (questionIndex._index) return questionIndex._index;
  var packed = QUESTION_INDEX_PACKED, rec = packed.rec, lic = packed.lic, topics = packed.topics;
  var lang = packed.lang || {}, stride = lic.length + 1, out = {};
  for (var id = 1; id <= packed.max; id++) {
    var at = (id - 1) * stride, classified = null;
    for (var k = 0; k < lic.length; k++) {
      var digit = rec.charCodeAt(at + k) - 48;
      if (!digit) continue;
      if (!classified) classified = {};
      classified[lic[k]] = topics[digit - 1];
    }
    if (!classified) continue;
    var key = String(id);
    out[key] = { c: classified,
      l: Object.prototype.hasOwnProperty.call(lang, key) ? lang[key] : packed.langDefault,
      img: rec.charCodeAt(at + lic.length) === 49 ? 1 : 0 };
  }
  questionIndex._index = out;
  return out;
}

var QUESTION_LANGS = ['he', 'ru', 'en', 'ar', 'fr', 'es', 'am'];
var QUESTION_ANSWER_COUNT = 4;

var EXAM_STRUCTURE_SERVER = {
  'B':  { 'בטיחות': 7, 'הכרת הרכב': 7, 'חוק': 7, 'תמרורים': 9 },
  '1':  { 'בטיחות': 5, 'הכרת הרכב': 5, 'חוק': 6, 'תמרורים': 6, 'ספציפי': 8 },
  'C1': { 'בטיחות': 5, 'הכרת הרכב': 5, 'חוק': 5, 'תמרורים': 5, 'ספציפי': 10 },
  'C':  { 'בטיחות': 5, 'הכרת הרכב': 4, 'חוק': 3, 'תמרורים': 4, 'ספציפי': 14 },
  'D':  { 'בטיחות': 4, 'הכרת הרכב': 2, 'חוק': 5, 'תמרורים': 4, 'ספציפי': 15 }
};

function classifyCategoryServer(cat) {
  var c = String(cat || '').trim();
  if (/ספציפי/.test(c)) return 'ספציפי';
  if (/בטיחות/.test(c)) return 'בטיחות';
  if (/הכרת הרכב/.test(c)) return 'הכרת הרכב';
  if (/חוק/.test(c)) return 'חוק';
  if (/תמרורים/.test(c)) return 'תמרורים';
  if (/זכות קדימה/.test(c)) return 'חוק';
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
  var entry = questionIndex()[String(id)];
  return entry || null;
}

function questionIndexCount() { return Object.keys(questionIndex()).length; }

function questionLangBit(lang) {
  var i = QUESTION_LANGS.indexOf(String(lang || 'he').toLowerCase());
  return i < 0 ? 0 : (1 << i);
}

function questionTopic(id, license) {
  var entry = questionIndexEntry(id);
  return (entry && entry.c[String(license)]) || '';
}

function answerKeyIndex(id, lang) {
  if (typeof lookupCorrectIndex !== 'function') return null;
  var idx = lookupCorrectIndex(Number(id), String(lang || 'he').toLowerCase());
  if (idx === null || idx === undefined) return null;
  var n = Number(idx);
  return (isFinite(n) && n >= 0 && n < QUESTION_ANSWER_COUNT) ? n : null;
}

function indexIdsByTopic(license, lang) {
  var bit = questionLangBit(lang), lic = String(license), byTopic = {};
  var index = questionIndex();
  for (var id in index) {
    if (!Object.prototype.hasOwnProperty.call(index, id)) continue;
    var entry = index[id];
    if (!(entry.l & bit)) continue;
    var topic = entry.c[lic];
    if (!topic) continue;
    if (!byTopic[topic]) byTopic[topic] = [];
    byTopic[topic].push(Number(id));
  }
  return byTopic;
}

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

function drawShuffleOrder() {
  var order = [];
  for (var i = 0; i < QUESTION_ANSWER_COUNT; i++) order.push(i);
  return shuffleArrayServer(order);
}

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

defineAction('bankGrant', { methods: ['GET'], auth: 'examiner', handler: handleBankGrant,
  rateLimit: { max: 30, windowSec: 60, id: function(p) { return normalizeId(p.examinerId); } } });
function handleBankGrant(p) {
  var bank = bankGrantFor('examiner', null, 'ex:' + normalizeId(p.examinerId));
  if (!bank) return bankNotConfiguredResponse();
  return jsonResponse({ status: 'ok', bank: bank });
}
var API_DEPLOYMENT = "exam";

var THEORY_API_BUILD = '2026-09-22-r31';
var API_STARTED_AT = 0;

function apiActionList() {
  ensureLegacyActions();
  return apiActionNames();
}

function logTheoryApiTiming(phase, method, action, startedAt) {
  try {
    Logger.log('[API] ' + JSON.stringify({ build: THEORY_API_BUILD, phase: phase,
      method: method, action: apiActionList().indexOf(action) >= 0 ? action : 'unknown',
      elapsedMs: Math.max(0, Date.now() - startedAt) }));
  } catch (logErr) {  }
}

function theoryRetryableErrorResponse(err) {
  if (!err || err.retryable !== true) return null;
  return jsonResponse({ status: 'error', code: err.code || 'retry_later', retryable: true,
    waitSec: Math.max(1, Math.min(30, Number(err.waitSec) || 3)),
    message: err.userMessage || 'המערכת עמוסה כעת. אפשר לנסות שוב בעוד מספר שניות.' });
}


function dispatchApiAction(method, action, p) {
  ensureLegacyActions();
  var target = ACTION_TARGETS[action];
  if (target && target !== 'both' && API_DEPLOYMENT !== 'all' && target !== API_DEPLOYMENT) {
    return jsonResponse({ status: 'error', code: 'wrong_deployment',
      message: 'הפעולה שייכת לשרת אחר — יש לרענן את הדף' });
  }
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
  if (auth === 'teacher') {
    var teacherCheck = globalFunction('requireTeacherToken');
    return teacherCheck ? teacherCheck(p) : jsonResponse({ status: 'error', code: 'wrong_deployment',
      message: 'הפעולה שייכת לשרת אחר — יש לרענן את הדף' });
  }
  if (auth === 'examinee') return requireExamineeToken(p);
  if (auth === 'gateway') return requireGatewayKey(p);
  return null;
}

function requireGatewayKey(p) {
  var expected = gatewayKey();
  if (!expected || String(p.gatewayKey || '') !== expected) {
    return jsonResponse({ status: 'error', code: 'gateway_denied', message: 'gateway key invalid' });
  }
  return null;
}

var ACTION_TARGETS = {
  health: 'both', getOfficeNumber: 'both',
  startPractice: 'reports', submitPracticeResult: 'reports', loadStudentProgress: 'reports',
  saveStudentProgress: 'reports', studentJoinClass: 'reports',
  teacherLogin: 'reports', teacherVerifyLogin: 'reports', teacherDashboard: 'reports',
  teacherCreateClass: 'reports', teacherCloseClass: 'reports', teacherDeleteClass: 'reports',
  teacherRemoveStudent: 'reports', teacherGetClasses: 'reports', teacherClassDetails: 'reports',
  teacherExportData: 'reports', teacherCommanderDashboard: 'reports', teacherAtRiskList: 'reports',
  adminDashboard: 'reports', commanderDashboard: 'reports', centerManagerReport: 'reports',
  siteCombinedReport: 'reports', examinerForecast: 'reports',
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
  sessionSnapshot: 'exam', bankGrant: 'exam'
};


function legacyActionTable() {
  return [
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
    ['commanderCorrectResult', 'POST', 'examiner', 'handleCommanderCorrectResult'],
    ['submitManualResult', 'POST', 'examiner', 'handleSubmitManualResult'],
    ['correctExamineeMeta', 'GET,POST', 'examiner', 'handleCorrectExamineeMeta'],
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
    ['checkApproval', 'GET', 'none', 'handleClientOutdated'],
    ['getExamStatus', 'GET', 'none', 'handleClientOutdated'],
    ['addExamTime', 'GET', 'none', 'handleAddExamTime'],
    ['markFinished', 'GET', 'none', 'handleMarkFinished'],
    ['disqualify', 'GET,POST', 'none', 'handleDisqualify'],
    ['reportWarning', 'GET,POST', 'none', 'handleReportWarning'],
    ['cancelDisqualify', 'GET,POST', 'none', 'handleCancelDisqualify'],
    ['studentJoinClass', 'GET', 'none', 'handleStudentJoinClass'],
    ['submitPracticeResult', 'GET,POST', 'none', 'handleSubmitPracticeResult'],
    ['loadStudentProgress', 'GET', 'none', 'handleLoadStudentProgress'],
    ['saveStudentProgress', 'POST', 'none', 'handleSaveStudentProgress']
  ];
}

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
  return jsonResponse({ status: 'ok', officeWhatsApp: getOfficeWhatsAppNumber() });
}

function handleHealth(p) {
  var body = { status: 'ok', build: THEORY_API_BUILD, deployment: API_DEPLOYMENT,
    indexIds: questionIndexCount(),
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


function doGet(e) {
  var apiStartedAt = API_STARTED_AT = Date.now();
  var action = '';
  diagBegin('GET');
  try {
    var p = (e && e.parameter) || {};
    action = p.action || '';
    if (DIAG_EXEC) { DIAG_EXEC.action = action; DIAG_EXEC.t0 = apiStartedAt; }
    logTheoryApiTiming('start', 'GET', action, apiStartedAt);

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

function handleLogin(p) {
  var sheet = getSheet('בוחנים');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (normalizeId(data[i][1]) === normalizeId(p.idNumber)) {
      var row = i + 1;
      var failedAttempts = Number(data[i][8]) || 0;
      var lockoutUntil = data[i][9];
      if (lockoutUntil) {
        var lockoutDate = lockoutUntil instanceof Date ? lockoutUntil : new Date(lockoutUntil);
        if (new Date() < lockoutDate) {
          var minsLeft = Math.ceil((lockoutDate - new Date()) / 60000);
          return jsonResponse({ status: 'error', message: 'החשבון נעול עקב ניסיונות כושלים. נסה שוב בעוד ' + minsLeft + ' דקות' });
        }
        failedAttempts = 0;
        sheet.getRange(row, 9).setValue(0);
        sheet.getRange(row, 10).setValue('');
      }
      if (String(data[i][2]) === String(p.password)) {
        if (data[i][3] === 'כן' || data[i][3] === true || data[i][3] === 'TRUE') {
          if (failedAttempts > 0) {
            sheet.getRange(row, 9).setValue(0);
            sheet.getRange(row, 10).setValue('');
          }
          var token = generateToken();
          var expiry = new Date();
          expiry.setHours(expiry.getHours() + 12);
          var existingTokens = String(data[i][6] || '').trim();
          var tokenList = existingTokens ? existingTokens.split(',') : [];
          tokenList.push(token);
          if (tokenList.length > 5) tokenList = tokenList.slice(-5);
          sheet.getRange(row, 7).setValue(tokenList.join(','));
          sheet.getRange(row, 8).setValue(expiry);
          return jsonResponse({ status: 'ok', examiner: { name: data[i][0], id: normalizeId(data[i][1]), examinerNumber: String(data[i][4] || ''), role: String(data[i][5] || 'בוחן'), token: token } });
        } else {
          return jsonResponse({ status: 'error', message: 'החשבון אינו פעיל' });
        }
      } else {
        failedAttempts++;
        sheet.getRange(row, 9).setValue(failedAttempts);
        if (failedAttempts >= 5) {
          var lockout = new Date();
          lockout.setMinutes(lockout.getMinutes() + 15);
          sheet.getRange(row, 10).setValue(lockout);
          return jsonResponse({ status: 'error', message: 'יותר מדי ניסיונות כושלים. החשבון ננעל ל-15 דקות' });
        }
        return jsonResponse({ status: 'error', message: 'סיסמה שגויה' });
      }
    }
  }
  return jsonResponse({ status: 'error', message: 'בוחן לא נמצא' });
}

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
        responsibleExaminer: String((data[i].length > 13 ? data[i][13] : '') || ''),
        defaultPopulation: String((data[i].length > 14 ? data[i][14] : '') || ''),
        managerPhone: sitesMap[siteName] ? sitesMap[siteName].managerPhone : ''
      });
    }
  }
  return jsonResponse({ status: 'ok', sessions: sessions.slice(0, 20) });
}

function handleListAllSessions(p) {
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
  return jsonResponse({ status: 'ok', sessions: sessions.slice(0, 100) });
}

function handleCreateSession(p) {
  var sheet = getSheet('סשנים');
  var code = generateSessionCode();
  var now = new Date();
  var validUntil = new Date(now.getTime() + 8 * 60 * 60 * 1000);

  var exSheet = getSheet('בוחנים');
  var exData = exSheet.getDataRange().getValues();
  var exRow = -1;
  var examinerName = '';
  for (var ei = 1; ei < exData.length; ei++) {
    if (normalizeId(exData[ei][1]) === normalizeId(p.examinerId)) { exRow = ei + 1; examinerName = exData[ei][0]; break; }
  }

  var quotas = parseAndValidateQuotas(p.quotas);
  if (quotas.error) {
    return jsonResponse({ status: 'error', message: quotas.error });
  }

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
    String(p.defaultPopulation || '').trim()
  ]);

  return jsonResponse({
    status: 'ok',
    sessionCode: code,
    validUntil: validUntil.toISOString(),
    examinerName: examinerName,
    responsibleExaminer: responsibleExaminer
  });
}

var RESPONSIBLE_EXAMINER_HIDE_LIST = [
  'תומר לוי',
  'אביאור שמעוני',
  'דוד בטיטו'
];
function _normalizeNameForHideList(s) {
  return String(s || '').trim().replace(/\s+/g, ' ').toLowerCase();
}

function handleListActiveExaminers(p) {
  diagMark('sheet:examiners-list');
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
    var site = String(r.site || '').trim();
    var req = parseInt(r.requested, 10);
    var appr = parseInt(r.approved, 10);
    if (!QUOTA_VALID_LICENSES[lic]) {
      return { error: 'דרגה לא חוקית בשורה ' + (i + 1) };
    }
    var _qkey = site + '|' + lic;
    if (seen[_qkey]) {
      return { error: 'דרגה "' + lic + '" מופיעה יותר מפעם אחת' + (site ? ' לאתר "' + site + '"' : '') };
    }
    seen[_qkey] = true;
    if (!isFinite(req) || req < 0) req = 0;
    if (!isFinite(appr) || appr < 0) appr = 0;
    if (appr > req) appr = req;
    clean.push({ site: site, license: lic, requested: req, approved: appr });
  }
  return { rows: clean };
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

  var resData = readResultsTail().rows;
  var latestByExaminee = {};
  for (var r = 1; r < resData.length; r++) {
    if (String(resData[r][13]) !== String(sessionCode)) continue;
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
          responsibleExaminer: String((data[i].length > 13 ? data[i][13] : '') || ''),
          defaultPopulation: String((data[i].length > 14 ? data[i][14] : '') || '')
        }
      });
    }
  }
  return jsonResponse({ status: 'error', message: 'קוד סשן לא תקין' });
}

var REG_KEY_MEMO_SEC = 1800;
function regKeyMemoKey(sessionCode, idNumber, regKey) {
  return CACHE_KEY_PREFIX + 'reg_' + String(sessionCode || '').trim() + '_' + normalizeId(idNumber) + '_' + String(regKey || '').trim();
}
function rememberRegistrationToken(sessionCode, idNumber, regKey, token) {
  if (!regKey || !token) return;
  try { CacheService.getScriptCache().put(regKeyMemoKey(sessionCode, idNumber, regKey), String(token), REG_KEY_MEMO_SEC); } catch (e) {}
}
function recallRegistrationToken(sessionCode, idNumber, regKey) {
  if (!regKey) return '';
  try { return String(CacheService.getScriptCache().get(regKeyMemoKey(sessionCode, idNumber, regKey)) || ''); } catch (e) { return ''; }
}
function validRegKey(raw) {
  var key = String(raw || '').trim();
  return /^[A-Za-z0-9_-]{8,64}$/.test(key) ? key : '';
}

function handleRegisterExaminee(p) {
  var rlErr = requireRateLimit('registerExaminee', String(p.sessionCode || ''), 30, 60);
  if (rlErr) return rlErr;
  var regKey = validRegKey(p.regKey);
  var lock = null, held = false;
  try { lock = LockService.getScriptLock(); held = lock.tryLock(5000); } catch (eLock) { held = false; }
  try {
    return registerExamineeLocked(p, regKey);
  } finally {
    if (held) { try { lock.releaseLock(); } catch (eRel) {} }
  }
}

function registerExamineeLocked(p, regKey) {
  var MAX_PENDING_PER_SESSION = 50;
  var pendSheet = getSheet('ממתינים');
  var data = pendSheet.getDataRange().getValues();
  var activeCount = 0;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(p.sessionCode)) {
      var status = String(data[i][5] || '').trim();
      if (normalizeId(data[i][1]) === normalizeId(p.idNumber)) {
        if (status === 'waiting' || status === 'approved' || status === 'in_exam') {
          var remembered = recallRegistrationToken(p.sessionCode, p.idNumber, regKey);
          var rowToken = String((data[i].length > 12 ? data[i][12] : '') || '').trim();
          if (regKey && remembered && rowToken && remembered === rowToken) {
            return jsonResponse({ status: 'ok', examineeToken: rowToken, resumed: true });
          }
          return jsonResponse({ status: 'error', message: 'כבר רשום בסשן זה' });
        }
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
    '',
    '',
    examineeToken,
    0,
    hasExtendedScreen ? 'כן' : '',
    0,
    '',
    p.site || ''
  ]);
  invalidatePendingSnapshot(p.sessionCode);
  rememberRegistrationToken(p.sessionCode, p.idNumber, regKey, examineeToken);
  return jsonResponse({ status: 'ok', examineeToken: examineeToken });
}

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
  var storedPhone = String(hit.row[3] || '').replace(/[^0-9]/g, '');
  var givenPhone = String(p.phone || '').replace(/[^0-9]/g, '');
  if (storedPhone && givenPhone && storedPhone.slice(-7) !== givenPhone.slice(-7)) {
    return jsonResponse({ status: 'error', message: 'פרטים לא תואמים' });
  }
  setPendingStatus(sheet, hit.idx + 1, p.sessionCode, 'cancelled');
  return jsonResponse({ status: 'ok' });
}

function handleCheckApproval(p) {
  var rlErr = requireRateLimit('checkApproval', String(p.sessionCode || '') + '_' + normalizeId(p.idNumber), 60, 60);
  if (rlErr) return rlErr;
  var BASE_EXAM_MINUTES = 40;
  var snap = pendingRowsForSession(p.sessionCode);
  var found = scanApprovalRows(snap.rows, p, BASE_EXAM_MINUTES);
  if (snap.cached && (!found || found.terminal)) {
    found = scanApprovalRows(pendingRowsForSession(p.sessionCode, true).rows, p, BASE_EXAM_MINUTES);
  }
  return found ? jsonResponse(found.body) : jsonResponse({ status: 'error', message: 'לא נמצא רישום' });
}

function approvalTokenMismatch(row, p) {
  var storedToken = String((row.length > 12 ? row[12] : '') || '').trim();
  return Boolean(storedToken && p.examineeToken && String(p.examineeToken).trim() !== storedToken);
}

function scanApprovalRows(data, p, BASE_EXAM_MINUTES) {
  var newestFinished = null;
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]).trim() !== String(p.sessionCode).trim() || normalizeId(data[i][1]) !== normalizeId(p.idNumber)) continue;
    var approval = String(data[i][5] || 'waiting').trim();
    if (approval === 'completed' || approval === 'disqualified' || approval === 'cancelled' || approval === 'rejected') {
      if (!newestFinished) newestFinished = data[i];
      continue;
    }
    if (approvalTokenMismatch(data[i], p)) {
      return { body: { status: 'error', message: 'טוקן נבחן לא תקין', examineeTokenError: 'mismatch' }, terminal: false };
    }
    var response = { status: 'ok', approval: approval };
    response.audioMode = String(data[i][9] || '').trim() === 'on' ? 'on' : 'off';
    if (approval === 'approved' || approval === 'in_exam') {
      var ext = parseFloat(data[i][10]) || 1;
      if (ext !== 1.25 && ext !== 1.5) ext = 1;
      response.examMinutes = Math.round(BASE_EXAM_MINUTES * ext);
    }
    return { body: response, terminal: false };
  }
  if (!newestFinished) return null;
  var decided = String(newestFinished[5] || '').trim();
  if (decided !== 'rejected' && decided !== 'cancelled') return null;
  if (approvalTokenMismatch(newestFinished, p)) {
    return { body: { status: 'error', message: 'טוקן נבחן לא תקין', examineeTokenError: 'mismatch' }, terminal: true };
  }
  return { body: { status: 'ok', approval: decided }, terminal: true };
}

function handleApproveExaminee(p) {
  if (p.examinerId && !verifyExaminerForSession(p.sessionCode, p.examinerId)) {
    return jsonResponse({ status: 'error', message: 'אין הרשאה — בוחן לא תואם לסשן' });
  }
  var validExt = { '': true, '1.25': true, '1.5': true };
  var timeExt = String(p.timeExtension || '');
  if (!validExt[timeExt]) timeExt = '';

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
  if (timeExt) extras.timeExtension = timeExt;
  if (audioMode) extras.audio = audioMode;
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


function handleExaminerDashboard(p) {
  var code = String(p.sessionCode);
  var pendSheet = getSheet('ממתינים');
  var resSheet = getSheet('תוצאות');

  diagMark('sheet:pending-dash');
  var _pendT = readTail(pendSheet, 4);
  var pendData = _pendT.rows, pendOff = _pendT.off;
  diagMark('sheet:results-dash');
  var _resT = readTail(resSheet, 0);
  var resData = _resT.rows;
  diagMark('sheet:extensions-dash');
  var pending = [];
  var active = [];

  var extraMinById = {};
  try { extraMinById = extraMinutesBySession(code); } catch (e) {}

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
  var STALE_BUFFER_MS = 20 * 60 * 1000;
  for (var ci = 1; ci < pendData.length; ci++) {
    if (reconciled >= DASH_MAX_RECONCILE_PER_POLL) break;
    if (String(pendData[ci][0]) !== code) continue;
    var _ciStatus = String(pendData[ci][5]).trim();
    if (_ciStatus !== 'in_exam' && _ciStatus !== 'approved') continue;
    var _startedExam = (_ciStatus === 'in_exam');
    var ciId = pendData[ci][1];
    var examStart = pendData[ci][11] ? new Date(pendData[ci][11]) : null;
    var regTime = examStart || (pendData[ci][4] ? new Date(pendData[ci][4]) : null);
    var ciExt = parseFloat(pendData[ci][10]) || 1;
    if (ciExt !== 1.25 && ciExt !== 1.5) ciExt = 1;
    var maxMs = Math.round(BASE_EXAM_MS * ciExt) + STALE_BUFFER_MS + ((extraMinById[normalizeId(ciId)] || 0) * 60 * 1000);
    var isStale = regTime && (now.getTime() - regTime.getTime() > maxMs);
    var effectiveStale = isStale && _startedExam;

    var _rc = resBySessId[normalizeId(ciId)] || { dqResults: 0, otherResults: 0 };
    var dqResults = _rc.dqResults, otherResults = _rc.otherResults;
    var _pt = pendTermBySessId[normalizeId(ciId)] || { dqTerminals: 0, otherTerminals: 0 };
    var dqTerminals = _pt.dqTerminals, otherTerminals = _pt.otherTerminals;
    var effectiveResults = Math.min(dqResults, dqTerminals) + otherResults;
    var totalTerminals = dqTerminals + otherTerminals;
    var hasUnmatchedResult = effectiveResults > totalTerminals;

    if (hasUnmatchedResult || effectiveStale) {
      reconciled++;
      setPendingStatus(pendSheet, ci + 1 + pendOff, code, 'completed');
      pendData[ci][5] = 'completed';
      var _mk = normalizeId(ciId);
      if (!pendTermBySessId[_mk]) pendTermBySessId[_mk] = { dqTerminals: 0, otherTerminals: 0 };
      pendTermBySessId[_mk].otherTerminals++;
      if (effectiveStale && !hasUnmatchedResult) {
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
        resData.push(failRow);
        if (attemptHistory) attemptHistory.push([failRow[0], failRow[1], '', '', failRow[4], '', '', failRow[7]]);
        if (!resBySessId[_mk]) resBySessId[_mk] = { dqResults: 0, otherResults: 0 };
        resBySessId[_mk].otherResults++;
      }
    }
  }

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
    if (aiPassed === 'בוטל') continue;
    var aiId = normalizeId(resData[ai2][1]);
    attemptsTodayById[aiId] = (attemptsTodayById[aiId] || 0) + 1;
  }

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
      pendingById[idNorm] = item;
    } else {
      if (s === 'disqualified') item.dqPending = true;
      var prevA = activeById[idNorm];
      if (!prevA) {
        activeById[idNorm] = item;
      } else {
        var curDQ = (s === 'disqualified');
        var prevDQ = (prevA.status === 'disqualified');
        if (curDQ || !prevDQ) activeById[idNorm] = item;
      }
    }
  }
  for (var pkA in pendingById) pending.push(pendingById[pkA]);
  for (var akA in activeById) active.push(activeById[akA]);

  var latestResRowById = {};
  for (var jd = 1; jd < resData.length; jd++) {
    if (String(resData[jd][13]) !== code) continue;
    if (String(resData[jd][7] || '') === 'בוטל') continue;
    latestResRowById[normalizeId(resData[jd][1])] = jd;
  }
  var completed = [];
  for (var j = 1; j < resData.length; j++) {
    if (String(resData[j][13]) !== code) continue;
    if (String(resData[j][7] || '') === 'בוטל') continue;
    if (latestResRowById[normalizeId(resData[j][1])] !== j) continue;
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
      verified: (resData[j].length > 22) ? (resData[j][22] || '') : '',
      suspicious: (resData[j].length > 23) ? (resData[j][23] || '') : '',
      device: (resData[j].length > 29) ? (resData[j][29] || '') : ''
    });
  }

  var todayDD = ('0' + now.getDate()).slice(-2);
  var todayMM = ('0' + (now.getMonth() + 1)).slice(-2);
  var todayYYYY = now.getFullYear();
  var todayDate = todayDD + '/' + todayMM + '/' + todayYYYY;
  var todayExamsById = {};
  for (var ti = 1; ti < resData.length; ti++) {
    if (String(resData[ti][7] || '') === 'בוטל') continue;
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

function handleGetExamStatus(p) {
  if (!p.sessionCode || !p.idNumber) return jsonResponse({ status: 'error', message: 'חסר מזהה' });
  var rlErr = requireRateLimit('getExamStatus', String(p.sessionCode || '') + '_' + normalizeId(p.idNumber), 60, 60);
  if (rlErr) return rlErr;
  diagMark('sheet:pending-status');
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

var EXTRA_MINUTES_CACHE_SEC = 30;
function extraMinutesKey(sessionCode) { return CACHE_KEY_PREFIX + 'extmin_' + String(sessionCode || '').trim(); }
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

  var pendData = getSheet('ממתינים').getDataRange().getValues();
  var hit = findLatestPendingRow(pendData, p.sessionCode, p.idNumber);
  if (hit.idx === -1) return jsonResponse({ status: 'error', message: 'נבחן לא נמצא בסשן' });
  var name = hit.row[2] || '';

  var sessionRow = sessionRowByCode(p.sessionCode);
  var examinerName = sessionRow ? (sessionRow[2] || '') : '';

  getSheet('הארכות זמן').appendRow([new Date(), p.sessionCode, p.idNumber, name, minutes, reason, examinerName]);
  invalidateExtraMinutes(p.sessionCode);

  return jsonResponse({ status: 'ok', addedMinutes: minutes, totalExtraMinutes: sumExtraMinutes(p.sessionCode, p.idNumber) });
}

function handleMarkFinished(p) {
  if (!p.sessionCode || !p.idNumber) return jsonResponse({ status: 'error', message: 'חסר מזהה' });
  var pendSheet = getSheet('ממתינים');
  var data = pendSheet.getDataRange().getValues();
  var hit = findLatestPendingRow(data, p.sessionCode, p.idNumber);
  if (hit.idx === -1) return jsonResponse({ status: 'ok' });
  var storedToken = String((hit.row.length > 12 ? hit.row[12] : '') || '').trim();
  if (storedToken && p.examineeToken && String(p.examineeToken).trim() !== storedToken) {
    return jsonResponse({ status: 'error', examineeTokenError: 'mismatch' });
  }
  if (hit.status === 'in_exam') {
    if (pendSheet.getMaxColumns() < 19) pendSheet.insertColumnsAfter(pendSheet.getMaxColumns(), 19 - pendSheet.getMaxColumns());
    if (!String(pendSheet.getRange(1, 19).getValue() || '').trim()) pendSheet.getRange(1, 19).setValue('סיים במכשיר');
    writePendingCells(pendSheet, hit.idx + 1, p.sessionCode, { finishedOnDevice: nowISO() });
  }
  return jsonResponse({ status: 'ok' });
}

defineAction('sessionSnapshot', { methods: ['GET'], auth: 'gateway', handler: handleSessionSnapshot,
  rateLimit: { max: 60, windowSec: 60, id: function(p) { return String(p.sessionCode || ''); } } });
function handleSessionSnapshot(p) {
  var code = String(p.sessionCode || '').trim();
  if (!code) return jsonResponse({ status: 'error', message: 'חסר קוד סשן' });
  var snap = pendingRowsForSession(code);
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
      examMinutes: examMinutesFor(r),
      extraMinutes: extraMin[id] || 0,
      warn: Number(r[15]) || 0,
      fin: r.length > 18 && r[18] ? 1 : 0,
      ext: String(r[14] || '').trim() === 'כן' ? 1 : 0,
      dq: Number(r[13]) || 0
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
    if (!verifyToken(p.examinerId, p.token)) {
      return jsonResponse({ status: 'error', message: 'טוקן בוחן לא תקין', tokenExpired: true });
    }
    if (!examinerOwnsSession(p.sessionCode, p.examinerId)) {
      return jsonResponse({ status: 'error', message: 'אין הרשאה — בוחן לא תואם לסשן' });
    }
  } else {
    if (pendRowIdx === -1) {
      return jsonResponse({ status: 'error', message: 'אין נבחן רשום בסשן זה' });
    }
    if (pendStatus !== 'in_exam' && pendStatus !== 'approved' && pendStatus !== 'disqualified') {
      return jsonResponse({ status: 'error', message: 'מצב לא תקף לפסילה: ' + pendStatus });
    }
    var dqRlErr = requireRateLimit('disqualify', String(p.sessionCode || '') + '_' + normalizeId(p.idNumber), 10, 60);
    if (dqRlErr) return dqRlErr;
    var tokenCheck = verifyExamineeToken(p.sessionCode, p.idNumber, p.examineeToken);
    if (!tokenCheck.valid) {
      return jsonResponse({ status: 'error', message: 'טוקן נבחן לא תקין לפסילה עצמית', examineeTokenError: tokenCheck.reason });
    }
  }

  if (pendRowIdx !== -1) {
    var prevCount = (pendData[pendRowIdx].length > 13) ? (Number(pendData[pendRowIdx][13]) || 0) : 0;
    setPendingStatus(pendSheet, pendRowIdx + 1, p.sessionCode, 'disqualified', { dqCount: prevCount + 1 });
    for (var dqd = 1; dqd < pendData.length; dqd++) {
      if (dqd === pendRowIdx) continue;
      if (String(pendData[dqd][0]) !== String(p.sessionCode) || normalizeId(pendData[dqd][1]) !== normalizeId(p.idNumber)) continue;
      var dqdStatus = String(pendData[dqd][5]).trim();
      if (dqdStatus === 'in_exam' || dqdStatus === 'approved') {
        setPendingStatus(pendSheet, dqd + 1, p.sessionCode, 'cancelled');
      }
    }
  }

  var dqEventId = String(p.dqEventId || '');
  var sheet = getSheet('תוצאות');
  var data = readResultsTail().rows;
  var nowMs = Date.now();
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][13]) === String(p.sessionCode) && normalizeId(data[i][1]) === normalizeId(p.idNumber)) {
      var rowStatus = String(data[i][7]).trim();
      if ((rowStatus === 'פסול' || rowStatus === 'בוטל') && dqEventId && String(data[i][24] || '') === dqEventId) {
        return jsonResponse({ status: 'ok' });
      }
      if (rowStatus === 'פסול') {
        var rowDateRaw = data[i][0];
        var rowDate = null;
        try {
          if (rowDateRaw instanceof Date) rowDate = rowDateRaw;
          else if (rowDateRaw) {
            var m = String(rowDateRaw).match(/(\d{1,2})\/(\d{1,2})\/(\d{4})\s+(\d{1,2}):(\d{2})/);
            if (m) rowDate = new Date(+m[3], (+m[2]) - 1, +m[1], +m[4], +m[5]);
          }
        } catch (e) { rowDate = null; }
        if (rowDate && (nowMs - rowDate.getTime()) < 120000) {
          return jsonResponse({ status: 'ok', deduped: true });
        }
      }
      break;
    }
  }

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

function handleCancelDisqualify(p) {
  var cdTokenErr = requireExamineeToken(p);
  if (cdTokenErr) return cdTokenErr;
  var sc = String(p.sessionCode || '');
  var id = normalizeId(p.idNumber || '');
  if (!sc || !id) return jsonResponse({ status: 'ok' });

  var pendSheet = getSheet('ממתינים');
  var pendData = pendSheet.getDataRange().getValues();
  var pendHit = findLatestPendingRow(pendData, sc, id);
  if (pendHit.idx !== -1 && pendHit.status === 'disqualified') {
    setPendingStatus(pendSheet, pendHit.idx + 1, sc, 'in_exam');
  }

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

function handleForceComplete(p) {
  if (p.examinerId && !examinerOwnsSession(p.sessionCode, p.examinerId)) {
    return jsonResponse({ status: 'error', message: 'אין הרשאה — בוחן לא תואם לסשן' });
  }
  var pendSheet = getSheet('ממתינים');
  var pendData = pendSheet.getDataRange().getValues();
  var found = false;
  var name = '', phone = '', population = '', examineeLicense = '', examineeAudio = 'off', language = 'he';
  for (var j = pendData.length - 1; j >= 1; j--) {
    if (String(pendData[j][0]) !== String(p.sessionCode) || normalizeId(pendData[j][1]) !== normalizeId(p.idNumber)) continue;
    var fcStatus = String(pendData[j][5]).trim();
    if (fcStatus !== 'in_exam' && fcStatus !== 'approved') continue;
    if (!found) {
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

  var resSheet = getSheet('תוצאות');
  var resData = readResultsTail().rows;
  if (findLatestResultRow(resData, p.sessionCode, p.idNumber, false).idx !== -1) {
    return jsonResponse({ status: 'ok', message: 'נמצאה תוצאה קיימת — הסטטוס עודכן' });
  }

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

  var sheet = getSheet('תוצאות');
  var resRead = readResultsTail();
  var resHit = findLatestResultRow(resRead.rows, p.sessionCode, p.idNumber, false);
  var resultRowIdx = resHit.idx, resultStatus = resHit.status;

  var pendSheet = getSheet('ממתינים');
  var pendData = pendSheet.getDataRange().getValues();
  var pendHit = findLatestPendingRow(pendData, p.sessionCode, p.idNumber);
  var pendRowIdx = pendHit.idx, pendStatusNow = pendHit.status;

  if (resultStatus === 'פסול') {
    sheet.getRange(resultRowIdx + 1 + resRead.off, 8).setValue('בוטל');
    sheet.getRange(resultRowIdx + 1 + resRead.off, 18).setValue(false);
    SpreadsheetApp.flush();
    if (pendRowIdx !== -1 && (pendStatusNow === 'disqualified' || pendStatusNow === 'dq_confirmed')) {
      setPendingStatus(pendSheet, pendRowIdx + 1, p.sessionCode, 'in_exam');
    }
    return jsonResponse({ status: 'ok' });
  }

  if (pendRowIdx !== -1 && pendStatusNow === 'disqualified' &&
      (resultStatus === 'עבר' || resultStatus === 'נכשל' || resultStatus === 'בוטל')) {
    setPendingStatus(pendSheet, pendRowIdx + 1, p.sessionCode, 'completed');
    return jsonResponse({ status: 'ok', resolved: 'stale_dq_cleared' });
  }

  if (pendRowIdx !== -1 && pendStatusNow === 'disqualified' && resultRowIdx === -1) {
    setPendingStatus(pendSheet, pendRowIdx + 1, p.sessionCode, 'in_exam');
    return jsonResponse({ status: 'ok', resolved: 'no_result_reverted' });
  }

  return jsonResponse({ status: 'error', message: 'תוצאה לא נמצאה' });
}

function handleConfirmDQ(p) {
  if (!examinerOwnsSession(p.sessionCode, p.examinerId)) {
    return jsonResponse({ status: 'error', message: 'אין הרשאה — בוחן לא תואם לסשן' });
  }
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
  var hit = findLatestResultRow(read.rows, p.sessionCode, p.idNumber, true);
  if (hit.idx === -1) return jsonResponse({ status: 'error', message: 'תוצאה לא נמצאה' });
  var row = hit.row, rowNumber = hit.idx + 1 + read.off;
  var scoreNum = parseInt(String(row[5]).split('/')[0]) || 0;
  if (scoreNum < 24) {
    return jsonResponse({ status: 'error', message: 'ציון נמוך מדי לתיקון (מתחת ל-24)' });
  }
  sheet.getRange(rowNumber, 8).setValue('עבר');
  sheet.getRange(rowNumber, 18).setValue(false);
  sheet.getRange(rowNumber, 21).setValue(true);
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

function handleSubmitManualResult(p) {
  if (!verifyToken(p.examinerId, p.token)) {
    return jsonResponse({ status: 'error', message: 'טוקן בוחן לא תקין', tokenExpired: true });
  }
  if (!verifyExaminerForSession(p.sessionCode, p.examinerId)) {
    return jsonResponse({ status: 'error', message: 'אין הרשאה — בוחן לא תואם לסשן' });
  }
  var fullName = String(p.fullName || '').trim();
  var idNumber = String(p.idNumber || '').trim();
  var scoreNum = parseInt(p.score, 10);
  var totalNum = parseInt(p.total, 10) || 30;
  if (!fullName) return jsonResponse({ status: 'error', message: 'חובה למלא שם מלא' });
  if (!idNumber) return jsonResponse({ status: 'error', message: 'חובה למלא ת.ז.' });
  if (isNaN(scoreNum) || scoreNum < 0 || scoreNum > totalNum) {
    return jsonResponse({ status: 'error', message: 'ציון לא תקין (חייב להיות בין 0 ל-' + totalNum + ')' });
  }

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

  var percent = Math.round((scoreNum / totalNum) * 100);
  var passThreshold = Math.ceil(totalNum * 0.86);
  var passText = scoreNum >= passThreshold ? 'עבר' : 'נכשל';

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
    '',
    false,
    false,
    waLink,
    p.population || '',
    false,
    p.audioMode || 'off',
    'ידני',
    '',
    '',
    '',
    '',
    '',
    ''
  ]);
  SpreadsheetApp.flush();
  return jsonResponse({ status: 'ok', waLink: waLink, attempt: attemptNum });
}

function handleCorrectExamineeMeta(p) {
  if (!verifyToken(p.examinerId, p.token)) {
    return jsonResponse({ status: 'error', message: 'טוקן בוחן לא תקין', tokenExpired: true });
  }
  if (!examinerOwnsSession(p.sessionCode, p.examinerId)) {
    return jsonResponse({ status: 'error', message: 'אין הרשאה — בוחן לא תואם לסשן' });
  }
  var newSite = (typeof p.site !== 'undefined' && p.site !== null) ? String(p.site).trim() : '';
  var newPop = (typeof p.population !== 'undefined' && p.population !== null) ? String(p.population).trim() : '';
  var newPhone = (typeof p.phone !== 'undefined' && p.phone !== null) ? String(p.phone).trim() : null;
  var newId = (typeof p.newIdNumber !== 'undefined' && p.newIdNumber !== null) ? String(p.newIdNumber).trim() : '';
  var applyId = (newId && /^\d{5,10}$/.test(newId) && normalizeId(newId) !== normalizeId(p.idNumber));
  if (!newSite && !newPop && newPhone === null && !applyId) {
    return jsonResponse({ status: 'error', message: 'לא הוזנו שדות לעדכון' });
  }
  var sheet = getSheet('תוצאות');
  var metaRead = readResultsTail();
  var rows = metaRead.rows;
  for (var i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][13]) === String(p.sessionCode) && normalizeId(rows[i][1]) === normalizeId(p.idNumber)) {
      var rowIdx = i + 1 + metaRead.off;
      if (applyId) {
        var idCell = sheet.getRange(rowIdx, 2);
        idCell.setNumberFormat('@');
        idCell.setValue(newId);
      }
      if (newPhone !== null) {
        var phoneCell = sheet.getRange(rowIdx, 4);
        phoneCell.setNumberFormat('@');
        phoneCell.setValue(newPhone);
      }
      if (newSite) sheet.getRange(rowIdx, 11).setValue(newSite);
      if (newPop) sheet.getRange(rowIdx, 20).setValue(newPop);
      SpreadsheetApp.flush();
      return jsonResponse({ status: 'ok' });
    }
  }
  return jsonResponse({ status: 'error', message: 'תוצאה לא נמצאה' });
}

function handleCommanderCorrectResult(data) {
  if (!verifyToken(data.examinerId, data.token)) {
    return jsonResponse({ status: 'error', message: 'טוקן בוחן לא תקין', tokenExpired: true });
  }
  var role = getExaminerRole(data.examinerId);
  if (role !== 'מפקד') {
    return jsonResponse({ status: 'error', message: 'פעולה זו זמינה רק למפקדים' });
  }

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
  var hit = findLatestResultRow(sheet.getDataRange().getValues(), data.sessionCode, data.idNumber, true);
  if (hit.idx === -1) return jsonResponse({ status: 'error', message: 'תוצאה לא נמצאה' });
  var rowIdx = hit.idx + 1;
  var pct = Math.round((newScore / newTotal) * 100);
  sheet.getRange(rowIdx, 6).setValue(newScore + '/' + newTotal);
  sheet.getRange(rowIdx, 7).setValue(pct + '%');
  sheet.getRange(rowIdx, 8).setValue(newStatus);
  sheet.getRange(rowIdx, 18).setValue(newStatus === 'פסול');
  sheet.getRange(rowIdx, 21).setValue(true);
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
  if (!examinerOwnsSession(p.sessionCode, p.examinerId)) {
    return jsonResponse({ status: 'error', message: 'אין הרשאה — בוחן לא תואם לסשן' });
  }
  var sheet = getSheet('תוצאות');
  var read = readResultsTail();
  var data = read.rows;
  var wanted = {};
  var ids = p.idNumbers ? p.idNumbers.split(',') : [p.idNumber];
  for (var k = 0; k < ids.length; k++) wanted[normalizeId(ids[k])] = true;
  var count = 0;
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][13]) !== String(p.sessionCode)) continue;
    if (!wanted[normalizeId(data[i][1])]) continue;
    sheet.getRange(i + 1 + read.off, 17).setValue(true);
    count++;
  }
  return jsonResponse({ status: 'ok', updated: count });
}


var EXAM_BASE_MINUTES = 40;
var EXAM_MIN_QUESTIONS = 25;
var EXAM_MAP_CACHE_SEC = 10800;
var EXAM_MAP_MAX_AGE_MS = 8 * 3600 * 1000;
var EXAM_SUSPICIOUS_SEC = 180;

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
  } catch (e) {  }
}

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
    if (Array.isArray(map)) record.map = map;
  } catch (e) {  }
  return record;
}
function examRegistrationAgeMs(record) {
  var at = record && record.at ? new Date(record.at) : null;
  if (!at || isNaN(at.getTime())) return 0;
  return Date.now() - at.getTime();
}

function appendExamRegistration(sessionCode, idNumber, map, at, lang) {
  diagMark('sheet:append-exam');
  var sheet = getSheet('מבחנים');
  if (sheet.getLastRow() === 0 && SHEET_HEADERS['מבחנים']) sheet.appendRow(SHEET_HEADERS['מבחנים']);
  sheet.appendRow([String(sessionCode), normalizeId(idNumber), JSON.stringify(map), at, lang, 0]);
}

defineAction('startExam', { methods: ['POST'], auth: 'examinee', handler: handleStartExam,
  rateLimit: { max: 10, windowSec: 60, id: function(p) { return String(p.sessionCode || '') + '_' + normalizeId(p.idNumber); } } });
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
  if (!bankGrantConfigured()) return bankNotConfiguredResponse();

  var attempt = examAttemptKey(row);
  var record = readExamMapCache(sessionCode, data.idNumber, attempt);
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

  if (ctx.active.status === 'approved') {
    setPendingStatus(getSheet('ממתינים'), ctx.active.rowNumber, sessionCode, 'in_exam', { examStart: nowISO() });
    ctx.active.status = 'in_exam';
  }

  var bank = bankGrantFor('exam', examMapIds(record.map), sessionCode + ':' + normalizeId(data.idNumber));

  return jsonResponse({
    status: 'ok', build: THEORY_API_BUILD,
    examMinutes: examMinutesFor(row), extraMinutes: sumExtraMinutes(sessionCode, data.idNumber),
    audioMode: String(row[9] || '').trim() === 'on' ? 'on' : 'off',
    language: record.lang || lang, license: license, registeredAt: record.at,
    bank: bank,
    questions: examQuestionsForClient(record.map)
  });
}

function examMapIds(map) {
  var ids = [];
  for (var i = 0; i < map.length; i++) ids.push(map[i].qId);
  return ids;
}

function drawExamRegistration(sessionCode, idNumber, license, lang) {
  var drawn = drawExamIds(license, lang);
  if (drawn.length < EXAM_MIN_QUESTIONS) {
    throw questionBankUnavailable('נדרשות לפחות ' + EXAM_MIN_QUESTIONS + ' שאלות', String(drawn.length));
  }
  var map = [];
  for (var i = 0; i < drawn.length; i++) {
    var order = drawShuffleOrder();
    var correct = answerKeyIndex(drawn[i].id, lang);
    map.push({ qIdx: i, qId: drawn[i].id, shuffleOrder: order, correctShuffledIdx: order.indexOf(correct), topic: drawn[i].topic });
  }
  var at = nowISO();
  appendExamRegistration(sessionCode, idNumber, map, at, lang);
  return { map: map, at: at, lang: lang, unverified: 0 };
}

function examQuestionsForClient(map) {
  var out = [];
  for (var i = 0; i < map.length; i++) {
    out.push({ id: map[i].qId, order: map[i].shuffleOrder, topic: map[i].topic || '' });
  }
  return out;
}

function examMinutesFor(row) {
  var ext = parseFloat(row[10]) || 1;
  if (ext !== 1.25 && ext !== 1.5) ext = 1;
  return Math.round(EXAM_BASE_MINUTES * ext);
}

defineAction('getExamQuestions', { methods: ['GET'], auth: 'none', handler: handleClientOutdated });
defineAction('registerExamQuestions', { methods: ['POST'], auth: 'none', handler: handleClientOutdated });
function handleClientOutdated() {
  return jsonResponse({ status: 'error', code: 'client_outdated',
    message: 'גרסה חדשה של המערכת — יש לרענן את הדף (F5)' });
}

defineAction('markExamStarted', { methods: ['GET'], auth: 'examinee', handler: handleMarkExamStartedNoop });
function handleMarkExamStartedNoop() {
  return jsonResponse({ status: 'ok', already: true });
}

var RESULT_COL = { date: 0, id: 1, name: 2, phone: 3, license: 4, score: 5, percent: 6, verdict: 7,
  time: 8, examiner: 9, site: 10, classroom: 11, language: 12, session: 13, attempt: 14, wrongDetails: 15,
  sent: 16, dq: 17, waLink: 18, population: 19, corrected: 20, audio: 21, verified: 22, suspicious: 23,
  dqEventId: 24, correctedBy: 25, correctionReason: 26, correctionDate: 27, langPath: 28, device: 29 };
var RESULT_PASS_RATIO = 0.86;
var FABRICATED_FAIL_MARKERS = ['סגירת דפדפן', 'טיימאאוט', 'סיום ידני'];
var UNVERIFIED_PREFIX = '⚠️ ציון לא אומת';

defineAction('submitResult', { methods: ['POST'], auth: 'examinee', handler: handleSubmitResult });
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

  diagMark('sheet:results-submit');
  var sheet = getSheet('תוצאות');
  var tail = readTail(sheet, RESULT_COL.date);
  supersedeFabricatedFails(sheet, tail, data);

  var duplicate = findDuplicateResult(tail, data, gate);
  if (duplicate) {
    markPendingCompleted(data.sessionCode, data.idNumber, gate.pending);
    return jsonResponse({ status: 'ok', waLink: duplicate[RESULT_COL.waLink] || '', duplicate: true });
  }
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

function submitGate(data) {
  var rlErr = requireRateLimit('submitResult', String(data.sessionCode || '') + '_' + normalizeId(data.idNumber), 5, 60);
  if (rlErr) return { error: rlErr };
  var ctx = examineeRowContext(data.sessionCode, data.idNumber);
  var status = ctx.latest ? ctx.latest.status : '';
  if (data.sessionCode && data.idNumber && ['in_exam', 'approved', 'completed', 'cancelled'].indexOf(status) === -1) {
    return { error: jsonResponse({ status: 'error', message: 'נבחן לא מאושר — לא ניתן לשלוח תוצאות' }) };
  }
  return { pending: pendingSnapshotFromTail(ctx), status: status, ctx: ctx };
}

function pendingSnapshotFromTail(ctx) {
  var sheet = getSheet('ממתינים'), tail = ctx.tail;
  if (!tail || !tail.rows.length) return { sheet: sheet, rows: null };
  if (!tail.off) return { sheet: sheet, rows: tail.rows };
  var padded = [tail.rows[0]];
  for (var i = 0; i < tail.off; i++) padded.push([]);
  return { sheet: sheet, rows: padded.concat(tail.rows.slice(1)) };
}

function recordSubmitClientLog(data) {
  if (!data.clientLog) return;
  try {
    if (typeof diagRecordClientLog === 'function') diagRecordClientLog(data.sessionCode, data.idNumber, data.clientLog);
  } catch (e) {  }
}

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

function applyScore(data, scored, registration, hasAnswers) {
  if (!scored) {
    data.verified = false;
    if (hasAnswers) {
      data.scoreUnverified = true;
      data.unverifiedReason = registration ? 'רישום מבחן פגום' : 'רישום מבחן חסר';
    }
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

var ANSWER_LABELS_HE = ['א', 'ב', 'ג', 'ד', 'ה', 'ו'];
var ANSWER_LABELS_LATIN = ['A', 'B', 'C', 'D', 'E', 'F'];
var TEXT_UNAVAILABLE = '(טקסט לא זמין)';
function answerLabel(lang, idx) {
  var labels = (String(lang) === 'he') ? ANSWER_LABELS_HE : ANSWER_LABELS_LATIN;
  return labels[idx] || '';
}
function displayedAnswer(shown, idx, lang) {
  if (idx === null) return '(לא זמין כעת)';
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

function attemptRows(tail) { return tail.off ? null : tail.rows; }

function resultRowMatchesExaminee(row, data) {
  return String(row[RESULT_COL.session]) === String(data.sessionCode) &&
    normalizeId(row[RESULT_COL.id]) === normalizeId(data.idNumber);
}

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

defineAction('submitFailOnClose', { methods: ['POST'], auth: 'examinee', handler: handleSubmitFailOnClose });
function handleSubmitFailOnClose(data) {
  var ctx = examineeRowContext(data.sessionCode, data.idNumber);
  var status = ctx.latest ? ctx.latest.status : '';
  if (status === 'cancelled' || status === 'rejected') return jsonResponse({ status: 'ok', skipped: 'cancelled' });

  var sheet = getSheet('תוצאות');
  var tail = readTail(sheet, RESULT_COL.date);
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
  row[RESULT_COL.langPath] = '';
  sheet.appendRow(row);

  markPendingCompleted(data.sessionCode, data.idNumber, pendingSnapshotFromTail(ctx));
  return jsonResponse({ status: 'ok' });
}

defineAction('cancelFailOnClose', { methods: ['POST'], auth: 'examinee', handler: handleCancelFailOnClose });
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
    break;
  }
  return jsonResponse({ status: 'ok' });
}

function restorePendingToInExam(sessionCode, idNumber) {
  var ctx = examineeRowContext(sessionCode, idNumber, true);
  if (!ctx.latest || ctx.latest.status !== 'completed') return;
  setPendingStatus(getSheet('ממתינים'), ctx.latest.rowNumber, sessionCode, 'in_exam', null);
}

defineAction('getResultUploadToken', { methods: ['GET'], auth: 'examiner', handler: handleGetResultUploadToken });
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
var DIAG_SHEET = 'אבחון';
var DIAG_SLOW_MS = 15000;
var DIAG_STALE_MS = 420000;
var DIAG_EXEC = null;

function diagBegin(method) {
  try { DIAG_EXEC = { id: Utilities.getUuid(), method: method, action: '', phase: '', marked: false, notes: [] }; }
  catch (e) { DIAG_EXEC = null; }
}

var DIAG_MARK_MIN_MS = 8000;

function diagMark(phase) {
  try {
    if (!DIAG_EXEC) return;
    var elapsed = Date.now() - (DIAG_EXEC.t0 || Date.now());
    DIAG_EXEC.phase = phase;
    DIAG_EXEC.notes.push(phase + '@' + elapsed);
    if (elapsed < DIAG_MARK_MIN_MS || DIAG_EXEC.marked) return;
    DIAG_EXEC.marked = true;
    PropertiesService.getScriptProperties().setProperty(CACHE_KEY_PREFIX + 'diag_' + DIAG_EXEC.id,
      JSON.stringify({ a: DIAG_EXEC.action, m: DIAG_EXEC.method, ph: phase, t: Date.now() }));
  } catch (e) {  }
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
  } catch (e) {  }
  finally { DIAG_EXEC = null; }
}

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

function flushDiagnostics() {
  var r = diagSweep(null);
  var msg = 'flushDiagnostics: ' + r.flushed + ' parked row(s) written, ' + r.swept + ' killed-execution marker(s) recorded';
  Logger.log(msg);
  return msg;
}

function diagSweep(summary) {
  var swept = 0, flushed = 0;
  try {
    var props = PropertiesService.getScriptProperties(), all = props.getProperties(), prefix = CACHE_KEY_PREFIX + 'diag_';
    var rowPrefix = CACHE_KEY_PREFIX + DIAG_ROW_PREFIX;
    var sheet = null;
    for (var key in all) {
      if (!Object.prototype.hasOwnProperty.call(all, key)) continue;
      if (key.indexOf(rowPrefix) === 0) {
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
      if (entry && entry.t && Date.now() - entry.t < DIAG_STALE_MS) continue;
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

function getDiagnosticsSheet() {
  var ss = getSpreadsheet(), sheet = ss.getSheetByName(DIAG_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(DIAG_SHEET);
    sheet.getRange(1, 1, 1, 7).setValues([['זמן', 'סוג', 'שיטה', 'פעולה', 'משך (ms)', 'שלב אחרון', 'הערות']]);
    sheet.getRange(1, 1, 1, 7).setFontWeight('bold');
  }
  return sheet;
}
