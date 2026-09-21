// ---- Diagnostics that survive a killed execution ---------------------------
// Per-execution logs are unreachable in this project's Executions page (no
// Cloud project is linked), so a 360-second row says nothing about WHERE it
// hung. A phase marker is written to ScriptProperties before each risky
// Google call and deleted when the request finishes; a killed execution never
// reaches the delete, so its last phase is still there for the warmup to
// sweep into the 'אבחון' sheet. Requests that finish but take longer than
// DIAG_SLOW_MS write their own row. Marks may now be placed anywhere, polling
// handlers included: the property is written only after DIAG_MARK_MIN_MS, so a
// healthy request costs nothing at all (see diagMark).
var DIAG_SHEET = 'אבחון';
var DIAG_SLOW_MS = 15000;
var DIAG_STALE_MS = 420000;
var DIAG_EXEC = null;

function diagBegin(method) {
  // Runs before the request's try/catch: it must be incapable of throwing.
  try { DIAG_EXEC = { id: Utilities.getUuid(), method: method, action: '', phase: '', marked: false, notes: [] }; }
  catch (e) { DIAG_EXEC = null; }
}

// r15: a mark is now FREE until the request is already in trouble.
//
// Every mark used to cost a ScriptProperties round-trip, which was affordable
// only because marks were kept off the polling handlers. That restriction made
// the one handler we most needed to understand - examinerDashboard, polled every
// 2s by every examiner for the whole exam - the one handler with no trail at
// all: it recorded 27.9s on 2026-09-15 with not a single phase to show for it.
// So the phase trail is kept in memory (free, and diagFinish's SLOW row reads it
// from there), and the property - whose only job is to survive a 360s kill so
// the sweep can report where the execution died - is written only once the
// request has already passed DIAG_MARK_MIN_MS. A healthy 2s poll now pays
// nothing; anything slow enough to be killed crossed the threshold long before.
var DIAG_MARK_MIN_MS = 8000;

function diagMark(phase) {
  try {
    if (!DIAG_EXEC) return;
    var elapsed = Date.now() - (DIAG_EXEC.t0 || Date.now());
    DIAG_EXEC.phase = phase;
    DIAG_EXEC.notes.push(phase + '@' + elapsed);
    if (elapsed < DIAG_MARK_MIN_MS) return;   // still healthy: no service call
    PropertiesService.getScriptProperties().setProperty(QUESTION_CACHE_PREFIX + 'diag_' + DIAG_EXEC.id,
      JSON.stringify({ a: DIAG_EXEC.action, m: DIAG_EXEC.method, ph: phase, t: Date.now() }));
    DIAG_EXEC.marked = true;
  } catch (e) { /* diagnostics must never break a request */ }
}

function diagFinish(action, startedAt) {
  try {
    if (!DIAG_EXEC) return;
    var elapsed = Date.now() - startedAt;
    if (DIAG_EXEC.marked) {
      try { PropertiesService.getScriptProperties().deleteProperty(QUESTION_CACHE_PREFIX + 'diag_' + DIAG_EXEC.id); } catch (eDel) {}
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
// parked in ScriptProperties first (a different service), and appended in
// place only while appends are healthy: one append that fails or takes longer
// than DIAG_APPEND_SLOW_MS opens a breaker for DIAG_APPEND_BREAKER_SEC, and the
// rows wait for the hourly sweep (or flushDiagnostics() from the editor).
var DIAG_ROW_PREFIX = 'diagrow_';
var DIAG_APPEND_BREAKER_KEY = 'diagbreaker';
var DIAG_APPEND_BREAKER_SEC = 300;
var DIAG_APPEND_SLOW_MS = 2000;
function diagRecordRow(id, row) {
  var key = QUESTION_CACHE_PREFIX + DIAG_ROW_PREFIX + id, parked = false;
  try { PropertiesService.getScriptProperties().setProperty(key, JSON.stringify(row)); parked = true; } catch (eProp) {}
  var cache = null, breakerOpen = false;
  try { cache = CacheService.getScriptCache(); breakerOpen = !!cache.get(QUESTION_CACHE_PREFIX + DIAG_APPEND_BREAKER_KEY); } catch (eGet) {}
  if (breakerOpen) return 'parked';
  var t0 = Date.now(), appended = false;
  try { getDiagnosticsSheet().appendRow(row); appended = true; } catch (eAppend) {}
  if (!appended || Date.now() - t0 > DIAG_APPEND_SLOW_MS) {
    try { if (cache) cache.put(QUESTION_CACHE_PREFIX + DIAG_APPEND_BREAKER_KEY, String(Date.now()), DIAG_APPEND_BREAKER_SEC); } catch (ePut) {}
  }
  if (appended && parked) { try { PropertiesService.getScriptProperties().deleteProperty(key); } catch (eDel) {} }
  return appended ? 'appended' : 'parked';
}

// Run from the editor during an exam morning if 'אבחון' looks empty while the
// dashboards are slow: writes the parked SLOW rows and records killed executions.
function flushDiagnostics() {
  var r = diagSweep(null);
  var msg = 'flushDiagnostics: ' + r.flushed + ' parked SLOW row(s) written, ' + r.swept + ' killed-execution marker(s) recorded';
  Logger.log(msg);
  return msg;
}

// Called by the warmup: markers older than DIAG_STALE_MS belong to executions
// that never finished (killed at 360s, or crashed) - record where they were.
function diagSweep(summary) {
  var swept = 0, flushed = 0;
  try {
    var props = PropertiesService.getScriptProperties(), all = props.getProperties(), prefix = QUESTION_CACHE_PREFIX + 'diag_';
    var rowPrefix = QUESTION_CACHE_PREFIX + DIAG_ROW_PREFIX;
    var sheet = null;
    for (var key in all) {
      if (!Object.prototype.hasOwnProperty.call(all, key)) continue;
      if (key.indexOf(rowPrefix) === 0) {
        // a SLOW row parked while the append breaker was open (diagRecordRow)
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
  } catch (e) { if (summary) summary.push('diagnostics sweep: skipped (' + (e && e.message ? e.message : e) + ')'); }
  if (summary) summary.push('diagnostics sweep: ' + swept + ' stale marker(s) recorded, ' + flushed + ' parked row(s) flushed');
  return { swept: swept, flushed: flushed };
}

function getDiagnosticsSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), sheet = ss.getSheetByName(DIAG_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(DIAG_SHEET);
    sheet.getRange(1, 1, 1, 7).setValues([['זמן', 'סוג', 'שיטה', 'פעולה', 'משך (ms)', 'שלב אחרון', 'הערות']]);
    sheet.getRange(1, 1, 1, 7).setFontWeight('bold');
  }
  return sheet;
}

