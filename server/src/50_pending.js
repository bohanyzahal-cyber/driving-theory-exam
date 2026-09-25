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
//
// r32 (22/09/2026 evening, DESIGN §14.2) — CLAIM BEFORE WRITE. The lock alone
// was not enough once the phone stopped waiting: Google's delivery hop stalls
// 25-60 s for our projects (KNOWN_ISSUES #35), so the examinee page now gives
// up at 30 s and retries the SAME regKey while the first execution may still be
// queued — the two executions genuinely OVERLAP, and the script lock is only
// held for 5 s of tryLock. So the memo is written BEFORE the sheet is read,
// with the value 'pending' = "an execution is appending this registration right
// now". A second execution that finds 'pending' waits for the token instead of
// reading the sheet and appending a second row. Two executions of one regKey =
// one row, one token. A claim never survives its execution: it expires in
// REG_CLAIM_SEC and is dropped when the registration is refused.
var REG_KEY_MEMO_SEC = 1800;
var REG_CLAIM_PENDING = 'pending';   // never a token — see recallRegistrationToken
var REG_CLAIM_SEC = 60;              // an execution that has not appended by then is dead
var REG_CLAIM_WAIT_MS = 25000;       // under the page's 30 s deadline, so the waiter still answers
var REG_CLAIM_POLL_MS = 500;
function regKeyMemoKey(sessionCode, idNumber, regKey) {
  return CACHE_KEY_PREFIX + 'reg_' + String(sessionCode || '').trim() + '_' + normalizeId(idNumber) + '_' + String(regKey || '').trim();
}
function rememberRegistrationToken(sessionCode, idNumber, regKey, token) {
  if (!regKey || !token) return;
  try { CacheService.getScriptCache().put(regKeyMemoKey(sessionCode, idNumber, regKey), String(token), REG_KEY_MEMO_SEC); } catch (e) {}
}
// The raw memo: a token, REG_CLAIM_PENDING, or '' when nothing is remembered.
function readRegistrationMemo(sessionCode, idNumber, regKey) {
  if (!regKey) return '';
  try { return String(CacheService.getScriptCache().get(regKeyMemoKey(sessionCode, idNumber, regKey)) || ''); } catch (e) { return ''; }
}
// The memo as a TOKEN. A claim is not a token: answering 'pending' as an
// examineeToken would write it into the sheet and every later call would be
// refused with 'טוקן נבחן לא תקין' — the exact shape of #34.
function recallRegistrationToken(sessionCode, idNumber, regKey) {
  var memo = readRegistrationMemo(sessionCode, idNumber, regKey);
  return memo === REG_CLAIM_PENDING ? '' : memo;
}
function validRegKey(raw) {
  var key = String(raw || '').trim();
  return /^[A-Za-z0-9_-]{8,64}$/.test(key) ? key : '';
}
function claimRegistration(sessionCode, idNumber, regKey) {
  if (!regKey) return;
  try { CacheService.getScriptCache().put(regKeyMemoKey(sessionCode, idNumber, regKey), REG_CLAIM_PENDING, REG_CLAIM_SEC); } catch (e) {}
}
// Drop our own claim when we are answering without having appended. The value
// is checked first so this can never eat a token another execution just wrote.
function releaseRegistrationClaim(sessionCode, idNumber, regKey) {
  if (!regKey) return;
  try {
    var cache = CacheService.getScriptCache(), key = regKeyMemoKey(sessionCode, idNumber, regKey);
    if (String(cache.get(key) || '') === REG_CLAIM_PENDING) cache.remove(key);
  } catch (e) {}
}
// Wait for the execution that claimed this regKey to publish its token.
// Returns the token, or '' when the claim expired (that execution died) or the
// wait ran out — in both cases the caller registers normally.
function awaitRegistrationToken(sessionCode, idNumber, regKey) {
  for (var waited = 0; waited < REG_CLAIM_WAIT_MS; waited += REG_CLAIM_POLL_MS) {
    try { Utilities.sleep(REG_CLAIM_POLL_MS); } catch (eSleep) { return ''; }
    var memo = readRegistrationMemo(sessionCode, idNumber, regKey);
    if (!memo) return '';                                  // claim gone: the first execution died
    if (memo !== REG_CLAIM_PENDING) return memo;           // the token landed
  }
  return '';
}

function handleRegisterExaminee(p) {
  // Rate limit: max 120 registrations per minute per session. Prevents an
  // attacker with the session code from spamming hundreds of fake registrations
  // (the session itself is capped at MAX_PENDING_PER_SESSION live rows below).
  // 30 until r32; raised because the examinee page now retries a stalled
  // registration by itself (30 s attempts, sequential, DESIGN §14.2), so a
  // class of 40-50 phones on a stalling Google morning can legitimately send
  // ~2 per phone per minute - and a rate-limited answer there would only add
  // another round of retries.
  var rlErr = requireRateLimit('registerExaminee', String(p.sessionCode || ''), 120, 60);
  if (rlErr) return rlErr;
  // TODO 1.5 (r32): the session is validated BEFORE anything is written. Until
  // now registerExaminee was the only live action that never looked at 'סשנים',
  // so a stale page (or a code typed by hand) could append a row to a session
  // that had ended or expired — the examinee then waited forever for an
  // examiner who was not there. Same three messages as getSessionInfo, which
  // the page shows as they are. READ-ONLY on purpose: getSessionInfo also
  // deactivates an expired session, and a registration must not write to
  // 'סשנים' behind the examiner's back.
  var sessionErr = registrationSessionError(p.sessionCode);
  if (sessionErr) return sessionErr;
  var regKey = validRegKey(p.regKey);
  // Claim before write. A retry that overlaps the first execution waits for its
  // token instead of racing it to the append.
  if (regKey) {
    var memo = readRegistrationMemo(p.sessionCode, p.idNumber, regKey);
    if (memo === REG_CLAIM_PENDING) memo = awaitRegistrationToken(p.sessionCode, p.idNumber, regKey);
    // No token to resume from → this execution is the one that appends, and it
    // says so before it reads the sheet. (A token memo is left alone: the
    // locked path below verifies it against the live row, exactly as in r31.2.)
    if (!memo) claimRegistration(p.sessionCode, p.idNumber, regKey);
  }
  var lock = null, held = false;
  try { lock = LockService.getScriptLock(); held = lock.tryLock(5000); } catch (eLock) { held = false; }
  var outcome = {};
  try {
    return registerExamineeLocked(p, regKey, outcome);
  } finally {
    if (held) { try { lock.releaseLock(); } catch (eRel) {} }
    // r33: the fallback diagnosis of a row this execution appended, recorded
    // AFTER the lock is released — an append to 'אבחון' must never hold up the
    // next examinee's registration.
    if (outcome.gwDiag) { try { recordGatewayDiag(p.sessionCode, p.idNumber, 'register', outcome.gwDiag); } catch (eDiag) {} }
  }
}

// TODO 1.5: the three refusals of handleGetSessionInfo (44_sessions_misc.js),
// read-only. Served from the per-execution 'סשנים' memo (12_reads.js), so this
// costs one sheet read at most and nothing at all when the request already read
// the session for another reason.
function registrationSessionError(sessionCode) {
  var row = sessionRowByCode(sessionCode);
  if (!row) return jsonResponse({ status: 'error', message: 'קוד סשן לא תקין' });
  var active = row[10];                                    // K (11): פעיל
  if (active !== true && active !== 'TRUE' && String(active).toUpperCase() !== 'TRUE') {
    return jsonResponse({ status: 'error', message: 'הסשן הסתיים' });
  }
  if (new Date() > new Date(row[9])) {                     // J (10): תקף עד
    return jsonResponse({ status: 'error', message: 'תוקף הסשן פג' });
  }
  return null;
}

// outcome (optional, r33): told whether this execution appended a row that
// carries a fallback diagnosis, so the caller can record it outside the lock.
function registerExamineeLocked(p, regKey, outcome) {
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
          // Nothing was appended, so nothing is "pending" any more: a retry of
          // this same key must be answered at once, not made to wait 25 s.
          releaseRegistrationClaim(p.sessionCode, p.idNumber, regKey);
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
          releaseRegistrationClaim(p.sessionCode, p.idNumber, regKey);
          return jsonResponse({ status: 'error', message: 'יש פסילה הממתינה להחלטת הבוחן — פנה לבוחן לפני רישום מחדש' });
        }
      }
      if (status === 'waiting' || status === 'approved' || status === 'in_exam') {
        activeCount++;
      }
    }
  }
  if (activeCount >= MAX_PENDING_PER_SESSION) {
    releaseRegistrationClaim(p.sessionCode, p.idNumber, regKey);
    return jsonResponse({ status: 'error', message: 'הסשן מלא — לא ניתן לרשום נבחנים נוספים' });
  }
  var examineeToken = generateExamineeToken();
  // External-monitor indicator from client (screen.isExtended). Cheating risk
  // signal — examinee may be sharing window to a second screen with accomplice.
  var hasExtendedScreen = (p.hasExtendedScreen === '1' || p.hasExtendedScreen === 1 || p.hasExtendedScreen === true);
  // r33: a phone that already knows it cannot reach the Worker registers with
  // gwDiag, and its row carries the '📡' badge from the very first render
  // (column Q, see reportGateway in 52_pending_status.js).
  var gwDiag = sanitizeGatewayDiag(p.gwDiag);
  // r35: every field the examinee typed goes through cellSafe (22_util.js) — a
  // name that starts with '=' must land as text, never as a formula.
  pendSheet.appendRow([
    p.sessionCode,
    cellSafe(p.idNumber),
    cellSafe(p.fullName || ''),
    cellSafe(p.phone || ''),
    nowISO(),
    'waiting',
    cellSafe(p.language || ''),
    cellSafe(p.population || ''),
    cellSafe(p.license || ''),
    cellSafe(p.audioMode || 'off'),
    '',                       // K (10): הארכת זמן — נקבע ע"י הבוחן בעת אישור
    '',                       // L (11): התחלת מבחן — נקבע ע"י markExamStarted
    examineeToken,            // M (12): טוקן נבחן — מוחזר ללקוח, נדרש בקריאות עוקבות
    0,                        // N (13): ספירת DQ — מתעלה עם כל disqualify
    hasExtendedScreen ? 'כן' : '', // O (14): מסך נוסף — סימן אזהרה
    0,                        // P (15): ספירת אזהרות — מאותחל ל-0 (נכתב ע"י warning)
    gwDiag ? GATEWAY_LABEL_GOOGLE : '', // Q (16): אזהרה אחרונה — נכתב ע"י warning / r33: '📡' של גיבוי גוגל
    cellSafe(p.site || '')    // R (17): אתר — האתר שהנבחן בחר (מארח/אורח), לתצוגה חיה לבוחן
  ]);
  if (outcome && gwDiag) outcome.gwDiag = gwDiag;
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

// r35 (review 09 F-16 / 01 D7, KNOWN_ISSUES #43): the examinee token of the
// registration being cancelled is REQUIRED. Until now anyone with the session
// code and a classmate's ID number could cancel that classmate's waiting or
// approved registration — the phone check below was skipped whenever `phone`
// was simply left out. examinee.html sends the token on this call (its api
// decorator attaches it to every action but registerExaminee). A row written
// before tokens existed (empty column M) is still accepted, as everywhere.
//
// r35.1 (review_r35 F1): the regKey is accepted as the same proof when the
// token is MISSING. On a slow morning the row lands and the token answer does
// not reach the phone for 25-60 s or longer (KNOWN_ISSUES #34/#35) — the page
// sits on the waiting screen without a token, and "טעיתי? לחזרה ולתיקון
// הפרטים" in that window was refused. The page then re-registered with the
// same regKey, got the OLD row back ({resumed:true}), and the correction was
// silently lost (a corrected licence was even enforced as the old one at
// startExam). The regKey is random, generated and kept on the device, never
// shown to anyone; the server remembers regKey → token for REG_KEY_MEMO_SEC, so
// "this regKey's memo holds exactly the token of the row being cancelled"
// proves it is the device that created that row. It never rescues a token that
// is present and wrong ('mismatch'), and a claim ('pending') is not a token.
// examinee.html starts sending regKey on this call in a Pages push (KNOWN_ISSUES
// #45) — until then the window stays as it is.
function handleCancelRegistration(p) {
  var sheet = getSheet('ממתינים');
  var data = sheet.getDataRange().getValues();
  var hit = findLatestPendingRow(data, p.sessionCode, p.idNumber, ['waiting', 'approved']);
  if (hit.idx === -1) return jsonResponse({ status: 'error', message: 'לא נמצא רישום פעיל לביטול' });
  var tokenCheck = examineeTokenVerdict(hit.row, p.examineeToken);
  if (!tokenCheck.valid && !(tokenCheck.reason === 'missing' && regKeyProvesRow(hit.row, p))) {
    return jsonResponse({ status: 'error', message: 'טוקן נבחן לא תקין', examineeTokenError: tokenCheck.reason });
  }
  // Verify phone matches to prevent unauthorized cancellation
  var storedPhone = String(hit.row[3] || '').replace(/[^0-9]/g, '');
  var givenPhone = String(p.phone || '').replace(/[^0-9]/g, '');
  if (storedPhone && givenPhone && storedPhone.slice(-7) !== givenPhone.slice(-7)) {
    return jsonResponse({ status: 'error', message: 'פרטים לא תואמים' });
  }
  setPendingStatus(sheet, hit.idx + 1, p.sessionCode, 'cancelled');
  return jsonResponse({ status: 'ok' });
}

// True when p.regKey is the key this device registered this very row with: the
// regKey memo (registerExaminee) holds exactly the row's token. One cache read.
function regKeyProvesRow(row, p) {
  var key = validRegKey(p.regKey);
  var rowToken = String((row.length > 12 ? row[12] : '') || '').trim();
  if (!key || !rowToken) return false;
  return recallRegistrationToken(p.sessionCode, p.idNumber, key) === rowToken;
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

