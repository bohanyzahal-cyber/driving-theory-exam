// One spreadsheet handle per execution, and the first open is marked: on 17/09
// an examinerDashboard spent 84.8 s before its first sheet mark, and the trail
// could not say whether opening the document or reading 'בוחנים' took it.
//
// Since the split (22/09/2026, DESIGN §13.3) this may run in a STANDALONE Apps
// Script project — the reports deployment is not bound to any spreadsheet, so
// getActiveSpreadsheet() answers null there and the script must be told which
// document is the exam one. One Script Property, checked once per execution;
// a missing property is a configuration error and says so, rather than
// surfacing later as "cannot read property getSheetByName of null".
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

