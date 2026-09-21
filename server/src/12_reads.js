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
  return QUESTION_CACHE_PREFIX + PENDING_SNAPSHOT_PREFIX + String(sessionCode || '').trim();
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

