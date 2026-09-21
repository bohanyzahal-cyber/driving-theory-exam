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
// The three archive sheet NAMES live in 00_config.js since 22/09/2026: 12_reads
// and 22_util read them and they are `both`, while this module is `reports`.
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

// Run ONCE from the editor after a deploy — in the REPORTS project once the
// split is live (DESIGN §13.3), because both handlers ship there. Leaves EXACTLY
// two time triggers: archiveSheets 01:00 and rebuildAtRiskCache 03:00
// (Asia/Jerusalem). Every trigger of a retired job is removed by NAME — the
// functions themselves may no longer exist in the script, but Apps Script keeps
// running their triggers and each run burns from the 90-minutes-a-day trigger
// budget. NIGHTLY_OBSOLETE_HANDLERS is in 24_triggers.js (`both`), next to
// uninstallNightlyJobs, which is what the exam project runs.
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
