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
// r31 (22/09/2026, DESIGN §13.6) added warn/fin/ext/dq: the examiner board no
// longer polls on a timer — it waits on the Worker's fingerprint of these rows
// (/v1/session/watch), so every field the board DISPLAYS has to be in the
// fingerprint or a change to it would never wake anybody.
//
// r32 (22/09/2026 evening, DESIGN §14.1) — version 2. Until now a change cost
// the board TWO Google round trips: the Worker read this snapshot (which moved
// the fingerprint) and the page then called examinerDashboard to see WHAT
// changed. Google's delivery hop stalls 25-60 s at random for our projects
// (KNOWN_ISSUES #35), so every round trip is a lottery ticket; the second one
// buys nothing this one could not carry. So the snapshot now carries the whole
// board: the ממתינים columns the lists display, and the session's results.
// That deliberately ENDS the "no names, no phones" property of this response —
// the board displays names, and the only consumer of the extra fields is the
// Worker's /v1/session/watch, which answers an examiner grant. The examinee
// answers the Worker builds (approval/status) are unchanged and still carry
// none of it (tests/contracts.test.cjs pins them byte for byte). The examinee
// TOKEN is still never sent — only its SHA-256.
defineAction('sessionSnapshot', { methods: ['GET'], auth: 'gateway', handler: handleSessionSnapshot,
  rateLimit: { max: 60, windowSec: 60, id: function(p) { return String(p.sessionCode || ''); } } });
function handleSessionSnapshot(p) {
  var code = String(p.sessionCode || '').trim();
  if (!code) return jsonResponse({ status: 'error', message: 'חסר קוד סשן' });
  var snap = pendingRowsForSession(code);         // the same 4-second snapshot the pollers use
  var extraMin = {};
  try { extraMin = extraMinutesBySession(code); } catch (eExt) { extraMin = {}; }
  // The ONE read r32 adds: the same 'תוצאות' tail the board itself reads. No
  // cache — this handler only runs after a write announced itself or on the
  // Worker's 20 s safety re-read of a held session (r31.3), so it is ~3-4 a
  // minute per live session, and a result the examiner cannot see is worse
  // than a read.
  diagMark('sheet:results-snapshot');
  var resData = readResultsTail().rows;
  var attemptsToday = attemptsTodayFromResults(resData);   // 55_dashboard.js — the board's own rules
  var todayExams = todayExamsFromResults(resData);
  var rows = [];
  for (var i = 1; i < snap.rows.length; i++) {
    var r = snap.rows[i], id = normalizeId(r[1]);
    var row = {
      id: id,
      status: String(r[5] || '').trim(),
      tokenHash: hashExamineeToken(r.length > 12 ? r[12] : ''),
      audio: String(r[9] || '').trim() === 'on' ? 'on' : 'off',
      examMinutes: examMinutesFor(r),   // one rule for the exam length (60_exam.js)
      extraMinutes: extraMin[id] || 0,
      warn: Number(r[15]) || 0,                                  // P (16): ספירת אזהרות
      fin: r.length > 18 && r[18] ? 1 : 0,                       // S (19): סיים במכשיר
      ext: String(r[14] || '').trim() === 'כן' ? 1 : 0,             // O (15): מסך נוסף
      dq: Number(r[13]) || 0,                                    // N (14): ספירת DQ
      // r32: what the board's pending/active rows display. Same columns, same
      // defaults as the items handleExaminerDashboard builds, so the page can
      // render either source without a second rule.
      name: r[2] === undefined ? '' : r[2],                      // C (3)
      phone: r[3] === undefined ? '' : r[3],                     // D (4)
      time: r[4] === undefined ? '' : r[4],                      // E (5): זמן הרשמה — a Date serialises to ISO
      start: r[11] || '',                                        // L (12): התחלת מבחן
      lang: r[6] || '',                                          // G (7)
      pop: r[7] || '',                                           // H (8)
      site: (r.length > 17) ? (r[17] || '') : '',                // R (18)
      lic: r[8] || '',                                           // I (9)
      timeExt: String(r[10] || ''),                              // K (11)
      lastWarn: (r.length > 16) ? String(r[16] || '') : '',      // Q (17)
      attemptsToday: attemptsToday[id] || 0
    };
    // Sent only when there is something to say — an empty array on every row
    // would be pure weight in the Worker's fingerprint and in the cache.
    var te = todayExams[id];
    if (te && te.length > 0) row.todayExams = te;
    rows.push(row);
  }
  return jsonResponse({ status: 'ok', v: 2, at: Date.now(), rows: rows, results: snapshotResultsForSession(resData, code) });
}

// The board's completed list WITHOUT column P (פירוט שגויות). That blob is
// ~2 KB per result and the board's own row only ever asked one question of it —
// "is this a fabricated fail?" (a browser-close / timeout / manual-finish row,
// which is shown differently). So the question is answered here, as
// `fabricated: 1`, and the blob stays on the sheet; the page fetches it from
// examinerDashboard on the click that actually needs it (a report, the wrong-
// answers table). Every other field keeps the board's name, so one renderer
// serves both routes.
function snapshotResultsForSession(resData, code) {
  var items = completedResultsForSession(resData, code), out = [];
  for (var i = 0; i < items.length; i++) {
    var item = items[i], slim = {};
    for (var key in item) {
      if (!Object.prototype.hasOwnProperty.call(item, key) || key === 'wrongDetails') continue;
      slim[key] = item[key];
    }
    if (isFabricatedFailNote(String(item.wrongDetails || ''))) slim.fabricated = 1;
    out.push(slim);
  }
  return out;
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

