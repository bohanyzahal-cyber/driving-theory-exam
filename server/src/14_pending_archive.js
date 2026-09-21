// ========== Nightly archive of ממתינים (perf) ==========
// Every live-path reader of ממתינים filters by the current sessionCode; the only
// history reader is the commander wait-time stat, which reads the archive too.
// Rows older than PENDING_ARCHIVE_RETAIN_DAYS are therefore dead weight on the
// hot path → moved (not deleted) to 'ממתינים_ארכיון'.
// Run archiveOldPendingRows() once by hand, then installPendingArchiveTrigger()
// for a daily 03:00 run (script time zone).
var PENDING_ARCHIVE_SHEET = 'ממתינים_ארכיון';
var PENDING_ARCHIVE_RETAIN_DAYS = 14;
var PENDING_TERMINAL = { completed: 1, disqualified: 1, dq_confirmed: 1, cancelled: 1, rejected: 1 };

function archiveOldPendingRows() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) { Logger.log('archive: another run holds the lock'); return; }
  markJobRunning('archiveOldPendingRows', true);   // the lock already keeps the cache check away; the flag names it
  var t0 = Date.now();
  var BUDGET_MS = 4.5 * 60 * 1000;   // stay under the 6-min ceiling; the rest moves next run
  try {
    var src = getSheet('ממתינים');
    var rows = src.getDataRange().getValues();
    if (rows.length <= 1) return;
    var now = Date.now();
    var cutoff = now - PENDING_ARCHIVE_RETAIN_DAYS * 86400000;
    var quietSince = now - 3 * 3600000;
    var rowNums = [];   // 1-based sheet rows to move, ascending
    var vals = [];
    for (var i = 1; i < rows.length; i++) {
      var st = String(rows[i][5] || '').trim();
      var reg = parseSheetDateTime(rows[i][4]);
      // Safety: never run while an exam may be in progress (a non-terminal row
      // registered in the last 3h) — concurrent handlers hold row indexes.
      if (!PENDING_TERMINAL[st] && reg && reg.getTime() > quietSince) {
        Logger.log('archive: exam activity detected — skipping this run');
        return;
      }
      if (reg && reg.getTime() < cutoff) { rowNums.push(i + 1); vals.push(rows[i]); }
    }
    if (!rowNums.length) { Logger.log('archive: nothing to move'); return; }

    var arch = getSheet(PENDING_ARCHIVE_SHEET);
    var width = rows[0].length;
    var CHUNK = 300;
    var moved = 0;
    // Chunks from the BOTTOM: copy → delete → next. Deleting bottom-up keeps the
    // remaining (smaller) row numbers valid, and concurrent registrations append
    // BELOW the snapshot so they are never touched.
    for (var c = rowNums.length; c > 0; c -= CHUNK) {
      var from = Math.max(0, c - CHUNK);
      var chunkRows = rowNums.slice(from, c);
      var chunkVals = vals.slice(from, c);
      arch.getRange(arch.getLastRow() + 1, 1, chunkVals.length, width).setValues(chunkVals);
      SpreadsheetApp.flush();
      var k = chunkRows.length - 1;
      while (k >= 0) {
        var end = chunkRows[k], start = end;
        while (k - 1 >= 0 && chunkRows[k - 1] === start - 1) { k--; start = chunkRows[k]; }
        src.deleteRows(start, end - start + 1);
        k--;
      }
      moved += chunkVals.length;
      if (Date.now() - t0 > BUDGET_MS) {
        Logger.log('archive: time budget hit — ' + moved + '/' + rowNums.length + ' moved, rest next run');
        return;
      }
    }
    Logger.log('archive: moved ' + moved + ' rows in ' + (Date.now() - t0) + 'ms');
  } finally {
    markJobRunning('archiveOldPendingRows', false);
    lock.releaseLock();
  }
}

function installPendingArchiveTrigger() {
  var trigs = ScriptApp.getProjectTriggers();
  for (var i = 0; i < trigs.length; i++) {
    if (trigs[i].getHandlerFunction() === 'archiveOldPendingRows') ScriptApp.deleteTrigger(trigs[i]);
  }
  // 01:00 — deliberately BEFORE the nightly rebuildAtRiskCache (03:00 Asia/Jerusalem)
  // so the two jobs never overlap and the forecast sees a settled sheet.
  ScriptApp.newTrigger('archiveOldPendingRows').timeBased().atHour(1).everyDays(1).inTimezone('Asia/Jerusalem').create();
  Logger.log('installed daily 01:00 Asia/Jerusalem trigger for archiveOldPendingRows');
}

