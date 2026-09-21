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

