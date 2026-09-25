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
      diagRecordRow(DIAG_EXEC.id, [nowISO(), 'SLOW', DIAG_EXEC.method, cellSafe(String(action || DIAG_EXEC.action || '')), elapsed,
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
    // r35: the text and the session code are the device's — cellSafe (22_util.js).
    return diagRecordRow(id, [nowISO(), 'CLIENT', cellSafe(String(sessionCode || '')), normalizeId(idNumber), '', '', cellSafe(text)]);
  } catch (e) { return 'error'; }
}

// r33 (24/09/2026, KNOWN_ISSUES #38): a phone that cannot reach the Worker says
// how the Worker failed it (examinee.html gwDiagString: 'v1|why=…|e=…|os=…|br=…'),
// with its registration (gwDiag) or on its own (reportGateway). Until now the
// only evidence of such a phone was its ABSENCE from the Worker's log. The text
// reaches the sheet, so only printable characters survive, and never more than
// DIAG_GATEWAY_MAX_CHARS. A missing value serialised by a careless caller
// ('undefined' / 'null') is no diagnosis at all.
var DIAG_GATEWAY_MAX_CHARS = 300;
function sanitizeGatewayDiag(raw) {
  if (raw === null || raw === undefined) return '';
  var text = String(raw);
  if (text === 'undefined' || text === 'null') return '';
  text = text.replace(/[\t\r\n]+/g, ' ').replace(/[^\x20-\x7E\u0590-\u05FF]/g, '').trim();
  return text.slice(0, DIAG_GATEWAY_MAX_CHARS);
}
// Lands exactly where the client logs land ('אבחון', type CLIENT), as ONE
// entry. JSON, so the cell starts with '[' and can never be read as a formula.
function recordGatewayDiag(sessionCode, idNumber, mode, diag) {
  if (!diag) return 'empty';
  try { return diagRecordClientLog(sessionCode, idNumber, [{ t: Date.now(), e: 'gw', m: String(mode || ''), d: diag }]); }
  catch (e) { return 'error'; }
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
        sheet.appendRow([nowISO(), 'KILLED', entry.m || '', cellSafe(String(entry.a || '')), '', entry.ph || '',
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
