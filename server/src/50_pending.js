// ---- Idempotent registration (r31.2, 22/09/2026) --------------------------
// A registration is a Google write that took 30-75 s on a stalling morning.
// The phone gave up at its deadline, the examinee pressed "הירשם" again, and the
// server answered the retry with 'כבר רשום בסשן זה' — which the page treats as
// success and goes on to the waiting screen WITHOUT A TOKEN. The examiner then
// approves (the Worker does not require a token to answer a poll), and every
// startExam fails with 'טוקן נבחן לא תקין' (reason: missing) until someone
// resets the examinee. Seen live at 08:35 on the test site; on an exam day it
// would have looked like "the exam never starts".
// Now the page sends a regKey — a random id it generates once per registration
// attempt and keeps across its retries — and the server remembers
// regKey → token for REG_KEY_MEMO_SEC. A retry with the SAME regKey gets the
// existing row's token back ({resumed:true}); a different device (another
// regKey, or none: an older page) still gets 'כבר רשום'. No new sheet column:
// the memo lives in CacheService, which is exactly as long-lived as a retry.
// A best-effort script lock closes the remaining window where two executions
// of the same registration read the sheet before either appended.
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
  // Rate limit: max 30 registrations per minute per session. Prevents an
  // attacker with the session code from spamming hundreds of fake registrations.
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
          // The same device asking again: hand back the row it already has.
          var remembered = recallRegistrationToken(p.sessionCode, p.idNumber, regKey);
          var rowToken = String((data[i].length > 12 ? data[i][12] : '') || '').trim();
          if (regKey && remembered && rowToken && remembered === rowToken) {
            return jsonResponse({ status: 'ok', examineeToken: rowToken, resumed: true });
          }
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
  rememberRegistrationToken(p.sessionCode, p.idNumber, regKey, examineeToken);
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
  // r23: served from the per-session snapshot (pendingRowsForSession). Every
  // answer that is NOT a live row ends the examinee's polling — 'לא נמצא רישום'
  // sends them back to the code screen, and the examiner's rejected/cancelled
  // decision puts a final message in front of them — and a cached snapshot can
  // simply predate a re-registration. So those answers are re-read from the
  // sheet first; only a live row is answered straight out of the snapshot.
  var snap = pendingRowsForSession(p.sessionCode);
  var found = scanApprovalRows(snap.rows, p, BASE_EXAM_MINUTES);
  if (snap.cached && (!found || found.terminal)) {
    found = scanApprovalRows(pendingRowsForSession(p.sessionCode, true).rows, p, BASE_EXAM_MINUTES);
  }
  return found ? jsonResponse(found.body) : jsonResponse({ status: 'error', message: 'לא נמצא רישום' });
}

// Token check: when a token is stored for this row, reject mismatches. Legacy
// rows (no stored token) and the very first poll (the client may not have
// echoed the token yet) are accepted so we don't break in-flight registrations
// during the deploy window.
function approvalTokenMismatch(row, p) {
  var storedToken = String((row.length > 12 ? row[12] : '') || '').trim();
  return Boolean(storedToken && p.examineeToken && String(p.examineeToken).trim() !== storedToken);
}

// The scan handleCheckApproval runs: { body, terminal } — `terminal` marks an
// answer taken from a FINISHED row, which handleCheckApproval refuses to serve
// from a cached snapshot — or null when this session has no row for this id.
function scanApprovalRows(data, p, BASE_EXAM_MINUTES) {
  var newestFinished = null;
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]).trim() !== String(p.sessionCode).trim() || normalizeId(data[i][1]) !== normalizeId(p.idNumber)) continue;
    var approval = String(data[i][5] || 'waiting').trim();
    // Terminal statuses are skipped so that the LIVE attempt is what answers:
    // a finished row from earlier in the day must never outrank the row the
    // examinee is waiting on right now. dq_confirmed is NOT terminal — the
    // examinee has to receive it.
    //
    // When nothing live is left, the NEWEST finished row decides, and only
    // when it is the examiner's own decision about the registration:
    //   rejected  → {approval:'rejected'}  — 'הבוחן דחה', + 'חזרה להרשמה'
    //   cancelled → {approval:'cancelled'} — 'ההרשמה בוטלה', + 'חזרה להרשמה'
    //   completed / disqualified / no row → 'לא נמצא רישום' (start over)
    //
    // Until r30 'rejected' was skipped outright and every all-terminal id was
    // answered 'not found', because of a real exam-day incident: two examinees
    // shared an ID number (family), the first was rejected at 17:47, the second
    // cancelled at 18:05, and a third visitor with stale localStorage polled
    // later — the scan skipped the newest (cancelled) row and returned the
    // older 'rejected' one, showing 'הבוחן דחה' on a screen nobody had
    // rejected. Taking the NEWEST finished row instead of the first one the
    // loop happens to like keeps that impossible: for that visitor the newest
    // row is the 18:05 'cancelled', so they are told the registration was
    // cancelled and sent to register again — never that they were rejected.
    // And the examinee whose CURRENT registration the examiner just rejected or
    // reset now learns it in-app instead of waiting out a 'שגיאת שרת' banner.
    if (approval === 'completed' || approval === 'disqualified' || approval === 'cancelled' || approval === 'rejected') {
      if (!newestFinished) newestFinished = data[i];   // walking newest→oldest: the first one IS the newest
      continue;
    }
    if (approvalTokenMismatch(data[i], p)) {
      return { body: { status: 'error', message: 'טוקן נבחן לא תקין', examineeTokenError: 'mismatch' }, terminal: false };
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
    return { body: response, terminal: false };
  }
  if (!newestFinished) return null;
  var decided = String(newestFinished[5] || '').trim();
  // completed / disqualified say nothing to a device that is still polling: the
  // exam is over, and 'לא נמצא רישום' is what sends it back to the code screen.
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

