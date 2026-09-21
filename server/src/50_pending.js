function handleRegisterExaminee(p) {
  // Rate limit: max 30 registrations per minute per session. Prevents an
  // attacker with the session code from spamming hundreds of fake registrations.
  var rlErr = requireRateLimit('registerExaminee', String(p.sessionCode || ''), 30, 60);
  if (rlErr) return rlErr;
  var MAX_PENDING_PER_SESSION = 50;
  var pendSheet = getSheet('ממתינים');
  var data = pendSheet.getDataRange().getValues();
  var activeCount = 0;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(p.sessionCode)) {
      var status = String(data[i][5] || '').trim();
      if (normalizeId(data[i][1]) === normalizeId(p.idNumber)) {
        if (status === 'waiting' || status === 'approved' || status === 'in_exam') {
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
  // r23: served from the per-session snapshot (pendingRowsForSession); a row
  // missing from a cached snapshot is re-read from the sheet before "not found".
  var snap = pendingRowsForSession(p.sessionCode);
  var found = scanApprovalRows(snap.rows, p, BASE_EXAM_MINUTES);
  if (!found && snap.cached) found = scanApprovalRows(pendingRowsForSession(p.sessionCode, true).rows, p, BASE_EXAM_MINUTES);
  return found || jsonResponse({ status: 'error', message: 'לא נמצא רישום' });
}

// The scan handleCheckApproval used to run inline — unchanged; null = no active row.
function scanApprovalRows(data, p, BASE_EXAM_MINUTES) {
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]).trim() === String(p.sessionCode).trim() && normalizeId(data[i][1]) === normalizeId(p.idNumber)) {
      var approval = String(data[i][5] || 'waiting').trim();
      // Skip terminal statuses from previous exams — keep looking for active row
      // Note: dq_confirmed is NOT skipped — examinee needs to receive this status
      //
      // 'rejected' is intentionally on the skip list. Real exam-day incident:
      // two examinees shared an ID number (family), first was rejected at
      // 17:47, second cancelled at 18:05. A third visitor with stale
      // localStorage polled later — the loop skipped the newest cancelled row
      // and returned the older 'rejected' status, showing "הבוחן דחה" on a
      // screen that nobody actually rejected. Skipping rejected here forces
      // the response to "no registration found" when all rows are terminal,
      // which the client interprets as "your saved state is stale, start over".
      //
      // Trade-off: when an examiner rejects a CURRENT registration, the
      // examinee no longer sees an in-app rejection notice — they see "no
      // registration" and reset to the code screen. Acceptable because the
      // examiner is physically next to them and can explain verbally.
      if (approval === 'completed' || approval === 'disqualified' || approval === 'cancelled' || approval === 'rejected') continue;
      // Token check: when a token is stored for this row, reject mismatches.
      // Legacy rows (no stored token) and the very first poll (client may not
      // have echoed the token yet) are accepted so we don't break in-flight
      // registrations during the deploy window.
      var storedToken = String((data[i].length > 12 ? data[i][12] : '') || '').trim();
      if (storedToken && p.examineeToken && String(p.examineeToken).trim() !== storedToken) {
        return jsonResponse({ status: 'error', message: 'טוקן נבחן לא תקין', examineeTokenError: 'mismatch' });
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
      return jsonResponse(response);
    }
  }
  return null;
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

