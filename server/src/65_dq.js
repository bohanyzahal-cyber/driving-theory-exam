function handleDisqualify(p) {
  // Auth: two valid paths
  //   A) Examiner-initiated DQ — must include valid token AND own the session
  //   B) Self-DQ from examinee client (cheat detection) — pending row must exist with active status
  // Without one of these, reject. Prevents an attacker with just sessionCode+victim's idNumber
  // from disqualifying other examinees.
  var pendSheet = getSheet('ממתינים');
  var pendData = pendSheet.getDataRange().getValues();
  var hit = findLatestPendingRow(pendData, p.sessionCode, p.idNumber);
  var name = '', phone = '', population = '', examineeLicense = '', examineeAudio = 'off';
  var pendRowIdx = hit.idx, pendStatus = hit.status;
  if (hit.idx !== -1) {
    name = hit.row[2] || '';
    phone = hit.row[3] || '';
    population = hit.row[7] || '';
    examineeLicense = hit.row[8] || '';
    examineeAudio = hit.row[9] || 'off';
  }

  if (p.examinerId) {
    // Path A: examiner-initiated — require valid token + ownership.
    // examinerOwnsSession reads 'סשנים' through the per-execution memo the
    // result row below reuses; this handler read the sheet twice (C R12).
    if (!verifyToken(p.examinerId, p.token)) {
      return jsonResponse({ status: 'error', message: 'טוקן בוחן לא תקין', tokenExpired: true });
    }
    if (!examinerOwnsSession(p.sessionCode, p.examinerId)) {
      return jsonResponse({ status: 'error', message: 'אין הרשאה — בוחן לא תואם לסשן' });
    }
  } else {
    // Path B: self-DQ — pending row must exist in active state AND the caller
    // must hold the examinee token issued at registration time. Legacy rows
    // (no stored token) are accepted as a transitional measure.
    if (pendRowIdx === -1) {
      return jsonResponse({ status: 'error', message: 'אין נבחן רשום בסשן זה' });
    }
    if (pendStatus !== 'in_exam' && pendStatus !== 'approved' && pendStatus !== 'disqualified') {
      return jsonResponse({ status: 'error', message: 'מצב לא תקף לפסילה: ' + pendStatus });
    }
    // Rate limit: max 10 self-DQ events per minute per (sessionCode, idNumber).
    // Anti-cheat can legitimately fire multiple beacons (retries, visibility +
    // blur racing); 10/min is well above any normal pattern.
    var dqRlErr = requireRateLimit('disqualify', String(p.sessionCode || '') + '_' + normalizeId(p.idNumber), 10, 60);
    if (dqRlErr) return dqRlErr;
    var tokenCheck = verifyExamineeToken(p.sessionCode, p.idNumber, p.examineeToken);
    if (!tokenCheck.valid) {
      return jsonResponse({ status: 'error', message: 'טוקן נבחן לא תקין לפסילה עצמית', examineeTokenError: tokenCheck.reason });
    }
  }

  // r33 (24/09/2026): which detector fired on the examinee's device
  // (examinee.html currentDQReason: split-area, zoom-out, hidden-10s, ...), for
  // the examiner's row as 'פסילה: <reason>' in column Q. Self-DQ only, and only
  // a short lowercase token — anything else is dropped, never written.
  var dqReason = p.examinerId ? '' : selfDqReason(p.reason);

  // r34 (24/09/2026): a disqualification this server ALREADY has — a result row
  // of this session and id with the same dqEventId, still 'פסול' or already
  // 'בוטל' by the examiner — changes nothing at all, 'ממתינים' included. The
  // examinee page now retries an unconfirmed DQ until it gets an answer
  // (examinee.html postDQUntilConfirmed; KNOWN_ISSUES #40). A retry whose first
  // attempt DID land but whose answer Google lost (#35) must not count the DQ a
  // second time in column N — and one that lands AFTER the examiner overturned
  // it must not disqualify the examinee again. The same check used to run only
  // after the pending row had already been rewritten below.
  var dqEventId = String(p.dqEventId || '');
  var sheet = getSheet('תוצאות');
  // The dedupe window is 2 minutes and a retry lands within minutes, so the tail
  // is always enough — this used to read every result ever recorded.
  var data = readResultsTail().rows;
  if (dqEventId && dqEventAlreadyRecorded(data, p.sessionCode, p.idNumber, dqEventId)) {
    return jsonResponse({ status: 'ok', duplicate: true });
  }

  // Update pending status to 'disqualified' (only if a row exists) AND increment
  // the DQ-event counter in column N so the examiner can see how many times this
  // examinee triggered an anti-cheat event — even if the examiner later
  // overturned some of them (overturnDQ).
  if (pendRowIdx !== -1) {
    var prevCount = (pendData[pendRowIdx].length > 13) ? (Number(pendData[pendRowIdx][13]) || 0) : 0;
    var dqExtras = { dqCount: prevCount + 1 };
    if (dqReason) dqExtras.lastWarning = 'פסילה: ' + dqReason;
    setPendingStatus(pendSheet, pendRowIdx + 1, p.sessionCode, 'disqualified', dqExtras);
    // Clear any OTHER active (in_exam/approved) rows for this examinee so a
    // duplicate row doesn't linger on the board beside the disqualified one.
    for (var dqd = 1; dqd < pendData.length; dqd++) {
      if (dqd === pendRowIdx) continue;
      if (String(pendData[dqd][0]) !== String(p.sessionCode) || normalizeId(pendData[dqd][1]) !== normalizeId(p.idNumber)) continue;
      var dqdStatus = String(pendData[dqd][5]).trim();
      if (dqdStatus === 'in_exam' || dqdStatus === 'approved') {
        setPendingStatus(pendSheet, dqd + 1, p.sessionCode, 'cancelled');
      }
    }
  }

  // Idempotency: prevent duplicate פסול rows when examinee anti-cheat AND examiner
  // manual DQ fire on the same examinee close in time (different dqEventIds).
  // Rules:
  //   1. Same dqEventId on a פסול/בוטל row -> retry, skip silently (checked above,
  //      before 'ממתינים' is touched — r34).
  //   2. Recent (≤2 min) פסול row WITHOUT 'בוטל' status -> same logical DQ event from
  //      another path (e.g. examiner clicked after auto-DQ already fired) -> skip.
  //   3. Otherwise (latest is not פסול, or it's old/cancelled) -> create new row.
  var nowMs = Date.now();
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][13]) === String(p.sessionCode) && normalizeId(data[i][1]) === normalizeId(p.idNumber)) {
      var rowStatus = String(data[i][7]).trim();
      // Rule 2: latest is an active 'פסול' (not cancelled) within last 2 minutes
      // → treat as the same DQ episode even if dqEventId differs/missing.
      if (rowStatus === 'פסול') {
        var rowDateRaw = data[i][0];
        var rowDate = null;
        try {
          if (rowDateRaw instanceof Date) rowDate = rowDateRaw;
          else if (rowDateRaw) {
            // Sheet date column F may be "DD/MM/YYYY HH:mm" — parse manually
            var m = String(rowDateRaw).match(/(\d{1,2})\/(\d{1,2})\/(\d{4})\s+(\d{1,2}):(\d{2})/);
            if (m) rowDate = new Date(+m[3], (+m[2]) - 1, +m[1], +m[4], +m[5]);
          }
        } catch (e) { rowDate = null; }
        if (rowDate && (nowMs - rowDate.getTime()) < 120000) {
          // Within 2 minutes of an active פסול → duplicate from race between
          // examiner button and examinee anti-cheat. Skip.
          return jsonResponse({ status: 'ok', deduped: true });
        }
      }
      // Rule 3: not a duplicate — fall through to create new row
      break;
    }
  }

  // Create new disqualified result row
  var sesRow = sessionRowByCode(p.sessionCode);
  var license = '', language = 'he', site = '', classroom = '', examinerName = '';
  if (sesRow) {
    examinerName = sesRow[2] || '';
    site = sesRow[3] || '';
    classroom = sesRow[4] || '';
    license = examineeLicense || sesRow[5] || '';
    language = sesRow[6] || 'he';
  }
  if (!license) license = examineeLicense;
  var attemptNum = countAttempts(String(p.idNumber), license) + 1;
  // r35.1 (review_r35 L1, F-15): the request's idNumber and dqEventId are
  // escaped — a self-DQ holds only the examinee token, and normalizeId matching
  // keeps only the digits, so '=…("<own id>")' passed auth and landed as a
  // formula in 'תוצאות'. r35.2 (review_r35_1_server m1): the WHOLE row goes
  // through cellSafeRow — name, phone, licence, audio and population are read
  // back from 'ממתינים' and site/classroom/examiner from 'סשנים', and Sheets
  // returns them WITHOUT the protecting apostrophe.
  sheet.appendRow(cellSafeRow([
    todayStr(), String(p.idNumber), name, phone, license,
    '0/30', '0%', 'פסול', '', examinerName,
    site, classroom, language, String(p.sessionCode),
    attemptNum, '', false, true, '',
    population, false, examineeAudio, '', '', dqEventId
  ]));
  SpreadsheetApp.flush();
  return jsonResponse({ status: 'ok' });
}

// Rule 1 of handleDisqualify: a result row of this session and id that carries
// this very dqEventId (column Y) and is still 'פסול' or was overturned to 'בוטל'.
function dqEventAlreadyRecorded(rows, sessionCode, idNumber, dqEventId) {
  var want = String(sessionCode), id = normalizeId(idNumber);
  for (var i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][13]) !== want || normalizeId(rows[i][1]) !== id) continue;
    if (String(rows[i][24] || '') !== dqEventId) continue;
    var status = String(rows[i][7]).trim();
    if (status === 'פסול' || status === 'בוטל') return true;
  }
  return false;
}

function selfDqReason(raw) {
  var reason = (raw === null || raw === undefined) ? '' : String(raw).trim();
  return /^[a-z0-9-]{1,24}$/.test(reason) ? reason : '';
}

// ---- cancelDisqualify: REMOVED in r35 (review 09 F-03, 01 D6, KNOWN_ISSUES #43)
// It let the EXAMINEE token undo a disqualification — any of them, the
// examiner's own included, with no time limit, even after confirmDQ (it voided
// the result row too). No page calls it: examinee.html clears the grace timer
// locally when the examinee comes back in time, BEFORE any DQ is sent, and its
// sendCancelDQToServer has no caller; the Worker never sends it. Undoing a DQ is
// the examiner's decision — overturnDQ (examiner token + session ownership).
// The name stays registered only to say so clearly instead of "Unknown action".
defineAction('cancelDisqualify', { methods: ['GET', 'POST'], auth: 'none', handler: handleCancelDisqualifyRemoved });
function handleCancelDisqualifyRemoved() {
  return jsonResponse({ status: 'error', code: 'action_removed',
    message: 'ביטול פסילה נעשה רק על ידי הבוחן' });
}

function handleResetExaminee(p) {
  if (p.examinerId && !examinerOwnsSession(p.sessionCode, p.examinerId)) {
    return jsonResponse({ status: 'error', message: 'אין הרשאה — בוחן לא תואם לסשן' });
  }
  var sheet = getSheet('ממתינים');
  var data = sheet.getDataRange().getValues();
  // Reset EVERY non-final row for this examinee (not just the latest) and accept
  // ALL stuck states — including 'disqualified'/'dq_confirmed'. Previously reset
  // refused those, so a soldier stuck on a pending DQ could not be cleared at all.
  // "אפס" should fully remove a stuck soldier from the board so they can re-register.
  var RESETTABLE = { waiting: 1, approved: 1, in_exam: 1, disqualified: 1, dq_confirmed: 1 };
  var resetCount = 0;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) !== String(p.sessionCode) || normalizeId(data[i][1]) !== normalizeId(p.idNumber)) continue;
    if (!RESETTABLE[String(data[i][5]).trim()]) continue;
    setPendingStatus(sheet, i + 1, p.sessionCode, 'cancelled');
    resetCount++;
  }
  if (resetCount === 0) {
    return jsonResponse({ status: 'error', message: 'לא נמצא נבחן פעיל לאיפוס' });
  }
  return jsonResponse({ status: 'ok', resetCount: resetCount });
}

// Force-complete a stuck in_exam examinee (examiner manual action)
function handleForceComplete(p) {
  if (p.examinerId && !examinerOwnsSession(p.sessionCode, p.examinerId)) {
    return jsonResponse({ status: 'error', message: 'אין הרשאה — בוחן לא תואם לסשן' });
  }
  var pendSheet = getSheet('ממתינים');
  var pendData = pendSheet.getDataRange().getValues();
  var found = false;
  var name = '', phone = '', population = '', examineeLicense = '', examineeAudio = 'off', language = 'he';
  // Close EVERY in_exam/approved row for this examinee (not just the latest, and
  // 'approved' too — markExamStarted can fail on iOS, leaving a stuck 'approved'
  // even after the soldier finished). One "סיים ידנית" must clear them all.
  for (var j = pendData.length - 1; j >= 1; j--) {
    if (String(pendData[j][0]) !== String(p.sessionCode) || normalizeId(pendData[j][1]) !== normalizeId(p.idNumber)) continue;
    var fcStatus = String(pendData[j][5]).trim();
    if (fcStatus !== 'in_exam' && fcStatus !== 'approved') continue;
    if (!found) { // capture details from the latest matching row
      name = pendData[j][2] || '';
      phone = pendData[j][3] || '';
      language = pendData[j][6] || 'he';
      population = pendData[j][7] || '';
      examineeLicense = pendData[j][8] || '';
      examineeAudio = pendData[j][9] || 'off';
    }
    setPendingStatus(pendSheet, j + 1, p.sessionCode, 'completed');
    found = true;
  }
  if (!found) {
    return jsonResponse({ status: 'error', message: 'לא נמצא נבחן עם סטטוס in_exam/approved' });
  }

  // Check if result already exists — if so, just mark pending as completed (done
  // above). The examinee is in THIS session, so their row is in the tail.
  var resSheet = getSheet('תוצאות');
  var resData = readResultsTail().rows;
  if (findLatestResultRow(resData, p.sessionCode, p.idNumber, false).idx !== -1) {
    return jsonResponse({ status: 'ok', message: 'נמצאה תוצאה קיימת — הסטטוס עודכן' });
  }

  // No result exists — create a fail result
  var sesRow = sessionRowByCode(p.sessionCode);
  var license = examineeLicense, site = '', classroom = '', examinerName = '';
  if (sesRow) {
    examinerName = sesRow[2] || '';
    site = sesRow[3] || '';
    classroom = sesRow[4] || '';
    if (!license) license = sesRow[5] || '';
  }
  var attemptNum = countAttempts(String(p.idNumber), license) + 1;
  // r35.2 (review_r35_1_server m1): read back from 'ממתינים' / 'סשנים' → cellSafeRow.
  resSheet.appendRow(cellSafeRow([
    todayStr(), String(p.idNumber), name, phone, license,
    '0/30', '0%', 'נכשל', '', examinerName,
    site, classroom, language, String(p.sessionCode),
    attemptNum, 'סיום ידני ע"י בוחן — ניתוק/תקלה', false, false, '',
    population, false, examineeAudio
  ]));
  SpreadsheetApp.flush();
  return jsonResponse({ status: 'ok', message: 'נבחן סומן כנכשל (ניתוק)' });
}

function handleOverturnDQ(p) {
  if (p.examinerId && !examinerOwnsSession(p.sessionCode, p.examinerId)) {
    return jsonResponse({ status: 'error', message: 'אין הרשאה — בוחן לא תואם לסשן' });
  }

  // Find the latest result row + the pending row for this examinee in one pass
  // each. An overturn always targets a result of the running session.
  var sheet = getSheet('תוצאות');
  var resRead = readResultsTail();
  var resHit = findLatestResultRow(resRead.rows, p.sessionCode, p.idNumber, false);
  var resultRowIdx = resHit.idx, resultStatus = resHit.status;

  var pendSheet = getSheet('ממתינים');
  var pendData = pendSheet.getDataRange().getValues();
  var pendHit = findLatestPendingRow(pendData, p.sessionCode, p.idNumber);
  var pendRowIdx = pendHit.idx, pendStatusNow = pendHit.status;

  // Case 1: latest result is פסול → normal overturn flow.
  // Pending revert covers BOTH 'disqualified' (auto-DQ, not yet confirmed) and
  // 'dq_confirmed' (examiner already clicked ✔ אשר). Without the dq_confirmed
  // branch, the examiner who pressed "אשר" by accident could overturn the
  // result row but the examinee stays locked out — they'd need a fresh
  // registration, which is what created duplicate rows at base 14.
  if (resultStatus === 'פסול') {
    sheet.getRange(resultRowIdx + 1 + resRead.off, 8).setValue('בוטל');
    sheet.getRange(resultRowIdx + 1 + resRead.off, 18).setValue(false);
    SpreadsheetApp.flush();
    if (pendRowIdx !== -1 && (pendStatusNow === 'disqualified' || pendStatusNow === 'dq_confirmed')) {
      setPendingStatus(pendSheet, pendRowIdx + 1, p.sessionCode, 'in_exam');
    }
    return jsonResponse({ status: 'ok' });
  }

  // Case 2: stuck pending in 'disqualified' but latest result is already
  // a final outcome (עבר/נכשל/בוטל). Happens when DQ fired transiently
  // during a deploy window — examinee continued and finished the exam, but
  // the pending row stayed stuck. Just clean up the pending row.
  if (pendRowIdx !== -1 && pendStatusNow === 'disqualified' &&
      (resultStatus === 'עבר' || resultStatus === 'נכשל' || resultStatus === 'בוטל')) {
    setPendingStatus(pendSheet, pendRowIdx + 1, p.sessionCode, 'completed');
    return jsonResponse({ status: 'ok', resolved: 'stale_dq_cleared' });
  }

  // Case 3: pending is disqualified but no result row yet → revert so the
  // examinee can resume the exam (in_exam state, just like case 1).
  if (pendRowIdx !== -1 && pendStatusNow === 'disqualified' && resultRowIdx === -1) {
    setPendingStatus(pendSheet, pendRowIdx + 1, p.sessionCode, 'in_exam');
    return jsonResponse({ status: 'ok', resolved: 'no_result_reverted' });
  }

  // Fall-through: nothing to do
  return jsonResponse({ status: 'error', message: 'תוצאה לא נמצאה' });
}

function handleConfirmDQ(p) {
  // Unconditional ownership check. confirmDQ is now in examinerActions → requireToken
  // already forced a valid examinerId. The old `if (p.examinerId && ...)` form could
  // be bypassed by simply OMITTING examinerId, letting anyone who knows session+id
  // finalize a victim's provisional DQ (robbing their grace-period recovery).
  if (!examinerOwnsSession(p.sessionCode, p.examinerId)) {
    return jsonResponse({ status: 'error', message: 'אין הרשאה — בוחן לא תואם לסשן' });
  }
  // Mark pending status as dq_confirmed so examinee polling gets a final answer
  var pendSheet = getSheet('ממתינים');
  var pendData = pendSheet.getDataRange().getValues();
  var hit = findLatestPendingRow(pendData, p.sessionCode, p.idNumber);
  if (hit.idx === -1 || hit.status !== 'disqualified') {
    return jsonResponse({ status: 'error', message: 'לא נמצא רישום פסול לאישור' });
  }
  setPendingStatus(pendSheet, hit.idx + 1, p.sessionCode, 'dq_confirmed');
  return jsonResponse({ status: 'ok' });
}

function handleCorrectToPass(p) {
  if (p.examinerId && !examinerOwnsSession(p.sessionCode, p.examinerId)) {
    return jsonResponse({ status: 'error', message: 'אין הרשאה — בוחן לא תואם לסשן' });
  }
  var sheet = getSheet('תוצאות');
  var read = readResultsTail();
  // skipCancelled: a 'בוטל' row was already overturned/superseded — correcting
  // it would resurrect a dead row and leave two live results (review E S7).
  var hit = findLatestResultRow(read.rows, p.sessionCode, p.idNumber, true);
  if (hit.idx === -1) return jsonResponse({ status: 'error', message: 'תוצאה לא נמצאה' });
  var row = hit.row, rowNumber = hit.idx + 1 + read.off;
  // Verify score is eligible (>= 24/30)
  var scoreNum = parseInt(String(row[5]).split('/')[0]) || 0;
  if (scoreNum < 24) {
    return jsonResponse({ status: 'error', message: 'ציון נמוך מדי לתיקון (מתחת ל-24)' });
  }
  sheet.getRange(rowNumber, 8).setValue('עבר');      // column H = עבר/נכשל
  sheet.getRange(rowNumber, 18).setValue(false);     // column R = disqualified (may be a DQ row)
  sheet.getRange(rowNumber, 21).setValue(true);      // column U = תוקן?
  // Regenerate WhatsApp link — corrected result shows only "עבר" (no score/errors)
  var waMsg = '*🚗 אישור תוצאת מבחן תאוריה חיצוני*\n\n' +
    'שם: ' + row[2] + '\n' +
    'ת.ז.: ' + row[1] + '\n' +
    'דרגה: ' + row[4] + '\n' +
    (row[19] ? 'אוכלוסיה: ' + row[19] + '\n' : '') +
    'תאריך: ' + row[0] + '\n' +
    'תוצאה: *עבר*\n';
  sheet.getRange(rowNumber, 19).setValue('https://wa.me/' + formatPhoneForWA(row[3]) + '?text=' + encodeURIComponent(waMsg));
  SpreadsheetApp.flush();
  return jsonResponse({ status: 'ok' });
}

// Commander-only result correction. Allows changing score and pass/fail/DQ
// Manual result entry for transition period — examiner enters a paper-based
// exam outcome that bypassed the digital system. Appends a row to תוצאות with
// the same shape submitResult uses; marks column W as 'ידני' so reports can
// distinguish it from system-scored results. Requires examiner-token +
// session ownership (same auth as overturnDQ/correctToPass).
//
// Required: sessionCode, idNumber, examinerId, token, fullName, score, total.
// Optional: phone, license, population, audioMode, time.
function handleSubmitManualResult(p) {
  if (!verifyToken(p.examinerId, p.token)) {
    return jsonResponse({ status: 'error', message: 'טוקן בוחן לא תקין', tokenExpired: true });
  }
  if (!verifyExaminerForSession(p.sessionCode, p.examinerId)) {
    return jsonResponse({ status: 'error', message: 'אין הרשאה — בוחן לא תואם לסשן' });
  }
  // Required field validation. ID/name keep manual entries debuggable; score
  // pair lets the spreadsheet compute the same "כך/סה״כ" string the digital
  // flow writes, so existing reports parse it without special-casing.
  var fullName = String(p.fullName || '').trim();
  var idNumber = String(p.idNumber || '').trim();
  var scoreNum = parseInt(p.score, 10);
  var totalNum = parseInt(p.total, 10) || 30;
  if (!fullName) return jsonResponse({ status: 'error', message: 'חובה למלא שם מלא' });
  if (!idNumber) return jsonResponse({ status: 'error', message: 'חובה למלא ת.ז.' });
  if (isNaN(scoreNum) || scoreNum < 0 || scoreNum > totalNum) {
    return jsonResponse({ status: 'error', message: 'ציון לא תקין (חייב להיות בין 0 ל-' + totalNum + ')' });
  }

  // Pull session context so manual rows match the rest of the session's rows
  // (same site/classroom/language) without the examiner re-typing them.
  var site = '', classroom = '', sessLicense = '', sessLanguage = 'he', examinerName = '';
  var sesRow = sessionRowByCode(p.sessionCode);
  if (sesRow) {
    examinerName = sesRow[2] || '';
    site = sesRow[3] || '';
    classroom = sesRow[4] || '';
    sessLicense = sesRow[5] || '';
    sessLanguage = sesRow[6] || 'he';
  }
  var license = String(p.license || sessLicense || 'B');
  var language = String(p.language || sessLanguage || 'he');

  // Percent + pass/fail mirror submitResult's behavior: 86% threshold (26/30).
  var percent = Math.round((scoreNum / totalNum) * 100);
  var passThreshold = Math.ceil(totalNum * 0.86);
  var passText = scoreNum >= passThreshold ? 'עבר' : 'נכשל';

  // WhatsApp link is convenient even for manual rows — examiner often wants to
  // send the same confirmation message they'd send for a digital exam.
  var waLink = '';
  if (p.phone) {
    var phoneFmt = formatPhoneForWA(p.phone);
    var waMsg = '*🚗 אישור תוצאת מבחן תאוריה חיצוני*\n\n' +
      'שם: ' + fullName + '\n' +
      'ת.ז.: ' + idNumber + '\n' +
      'דרגה: ' + license + '\n' +
      (p.population ? 'אוכלוסיה: ' + p.population + '\n' : '') +
      'תאריך: ' + todayStr() + '\n' +
      'ציון: ' + scoreNum + '/' + totalNum + ' (' + percent + '%)\n' +
      'תוצאה: *' + passText + '*\n';
    if (phoneFmt) waLink = 'https://wa.me/' + phoneFmt + '?text=' + encodeURIComponent(waMsg);
  }

  var attemptNum = countAttempts(idNumber, license) + 1;
  var sheet = getSheet('תוצאות');
  // Idempotency: a lost-response retry (request landed, reply dropped, examiner
  // re-saves) must not create a second identical manual row. Skip if a non-בוטל
  // row already exists for this session+id+license+score. The retry follows
  // within seconds, so the tail covers it.
  var manExisting = readResultsTail().rows;
  for (var mx = manExisting.length - 1; mx >= 1; mx--) {
    if (String(manExisting[mx][13]) === String(p.sessionCode) &&
        normalizeId(manExisting[mx][1]) === normalizeId(idNumber) &&
        String(manExisting[mx][4]) === String(license) &&
        String(manExisting[mx][5]) === (scoreNum + '/' + totalNum) &&
        String(manExisting[mx][7] || '').trim() !== 'בוטל') {
      return jsonResponse({ status: 'ok', duplicate: true, waLink: manExisting[mx][18] || '' });
    }
  }
  // r35.2: the examiner's typed fields and the session values read back from
  // 'סשנים' are written as text (cellSafeRow).
  sheet.appendRow(cellSafeRow([
    todayStr(),
    idNumber,
    fullName,
    p.phone || '',
    license,
    scoreNum + '/' + totalNum,
    percent + '%',
    passText,
    p.time || '',
    examinerName,
    site,
    classroom,
    language,
    String(p.sessionCode),
    attemptNum,
    '',                                 // P (15) wrongDetails — N/A for manual
    false,                              // Q (16) corrected
    false,                              // R (17) disqualified
    waLink,
    p.population || '',
    false,                              // U (20) suspicious
    p.audioMode || 'off',
    'ידני',                             // W (22) verified flag — marks paper-based entry
    '',                                 // X (23) suspicious text
    '',                                 // Y (24) dqEventId
    '',                                 // Z (25) תוקן ע"י
    '',                                 // AA (26) סיבת תיקון
    '',                                 // AB (27) תאריך תיקון
    ''                                  // AC (28) מסלול שפות
  ]));
  SpreadsheetApp.flush();
  return jsonResponse({ status: 'ok', waLink: waLink, attempt: attemptNum });
}

// status on any result row, with a mandatory reason recorded for audit.
// Caller must have a valid examiner token AND role 'מפקד' in the בוחנים sheet.
// Required params: sessionCode, idNumber, newScore (e.g. "28"), newTotal (e.g. "30"),
//                  newStatus ('עבר' | 'נכשל' | 'פסול'), reason (non-empty).
// Examiner-level correction of an examinee's site + population on their result
// row (the examinee picked the wrong site/population at registration). Reached
// via doPost (apiPost auto-attaches examinerId+token), which does NOT run the
// examinerActions allowlist — so the token is verified HERE, then session
// ownership. Updates תוצאות col 11 (אתר, idx 10) / col 20 (אוכלוסיה, idx 19).
function handleCorrectExamineeMeta(p) {
  if (!verifyToken(p.examinerId, p.token)) {
    return jsonResponse({ status: 'error', message: 'טוקן בוחן לא תקין', tokenExpired: true });
  }
  if (!examinerOwnsSession(p.sessionCode, p.examinerId)) {
    return jsonResponse({ status: 'error', message: 'אין הרשאה — בוחן לא תואם לסשן' });
  }
  var newSite = (typeof p.site !== 'undefined' && p.site !== null) ? String(p.site).trim() : '';
  var newPop = (typeof p.population !== 'undefined' && p.population !== null) ? String(p.population).trim() : '';
  var newPhone = (typeof p.phone !== 'undefined' && p.phone !== null) ? String(p.phone).trim() : null;  // null = "not sent" → don't touch
  // r35.2 (review r35.2 verification): the corrected phone is held to the rule
  // registration holds it to (examinee.html validateIdForm: 9-10 digits once
  // everything else is stripped), and refused BEFORE anything is written — it
  // used to land as whatever the request said. Empty = leave the phone as it is:
  // examiner.html pre-fills the field with the row's phone, which a manual row
  // may not have, and an empty field must not block a site or ID correction.
  if (newPhone === '') newPhone = null;
  if (newPhone !== null) {
    var phoneDigits = newPhone.replace(/[^0-9]/g, '');
    if (phoneDigits.length < 9 || phoneDigits.length > 10) {
      return jsonResponse({ status: 'error', code: 'invalid_phone', message: 'מספר טלפון לא תקין — נדרשות 9–10 ספרות' });
    }
  }
  var newId = (typeof p.newIdNumber !== 'undefined' && p.newIdNumber !== null) ? String(p.newIdNumber).trim() : '';
  // Only apply an id change when it's a valid digit string AND actually different.
  var applyId = (newId && /^\d{5,10}$/.test(newId) && normalizeId(newId) !== normalizeId(p.idNumber));
  if (!newSite && !newPop && newPhone === null && !applyId) {
    return jsonResponse({ status: 'error', message: 'לא הוזנו שדות לעדכון' });
  }
  var sheet = getSheet('תוצאות');
  var metaRead = readResultsTail();   // the examiner fixes a row of the session in front of them
  var rows = metaRead.rows;
  for (var i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][13]) === String(p.sessionCode) && normalizeId(rows[i][1]) === normalizeId(p.idNumber)) {
      var rowIdx = i + 1 + metaRead.off;
      if (applyId) {
        var idCell = sheet.getRange(rowIdx, 2);   // B (idx 1) = ת.ז.
        idCell.setNumberFormat('@');              // store as text — preserve leading zeros / avoid number formatting
        idCell.setValue(newId);
      }
      if (newPhone !== null) {
        var phoneCell = sheet.getRange(rowIdx, 4); // D (idx 3) = טלפון
        phoneCell.setNumberFormat('@');
        phoneCell.setValue(cellSafe(newPhone));   // r35.2: '+972…' stays text, whatever the cell format
      }
      if (newSite) sheet.getRange(rowIdx, 11).setValue(cellSafe(newSite));   // K (idx 10) = אתר (r35.2: as text)
      if (newPop) sheet.getRange(rowIdx, 20).setValue(cellSafe(newPop));     // T (idx 19) = אוכלוסיה
      SpreadsheetApp.flush();
      return jsonResponse({ status: 'ok' });
    }
  }
  return jsonResponse({ status: 'error', message: 'תוצאה לא נמצאה' });
}

function handleCommanderCorrectResult(data) {
  // Token + role check (token already verified by examinerActions allowlist,
  // but we re-check role here since the role doesn't appear in that allowlist).
  if (!verifyToken(data.examinerId, data.token)) {
    return jsonResponse({ status: 'error', message: 'טוקן בוחן לא תקין', tokenExpired: true });
  }
  var role = getExaminerRole(data.examinerId);
  if (role !== 'מפקד') {
    return jsonResponse({ status: 'error', message: 'פעולה זו זמינה רק למפקדים' });
  }

  // Validate inputs
  var reason = String(data.reason || '').trim();
  if (!reason) {
    return jsonResponse({ status: 'error', message: 'יש להזין סיבת תיקון' });
  }
  var newScore = parseInt(data.newScore, 10);
  var newTotal = parseInt(data.newTotal, 10);
  if (isNaN(newScore) || isNaN(newTotal) || newTotal <= 0 || newScore < 0 || newScore > newTotal) {
    return jsonResponse({ status: 'error', message: 'ציון חדש לא תקין' });
  }
  var newStatus = String(data.newStatus || '').trim();
  if (newStatus !== 'עבר' && newStatus !== 'נכשל' && newStatus !== 'פסול') {
    return jsonResponse({ status: 'error', message: 'סטטוס חדש לא תקין' });
  }

  // A commander corrects results of ANY session, including one from weeks ago,
  // so this is the one correction handler that keeps the full live read (it runs
  // a handful of times a month and must not miss a row a tail would cut off).
  var sheet = getSheet('תוצאות');
  var hit = findLatestResultRow(sheet.getDataRange().getValues(), data.sessionCode, data.idNumber, true);   // skip 'בוטל' (E S7)
  if (hit.idx === -1) return jsonResponse({ status: 'error', message: 'תוצאה לא נמצאה' });
  var rowIdx = hit.idx + 1;
  var pct = Math.round((newScore / newTotal) * 100);
  sheet.getRange(rowIdx, 6).setValue(newScore + '/' + newTotal);  // F: ציון
  sheet.getRange(rowIdx, 7).setValue(pct + '%');                   // G: אחוז
  sheet.getRange(rowIdx, 8).setValue(newStatus);                   // H: עבר/נכשל
  sheet.getRange(rowIdx, 18).setValue(newStatus === 'פסול');       // R: פסול?
  sheet.getRange(rowIdx, 21).setValue(true);                       // U: תוקן?
  // Audit trail (columns Z=26, AA=27, AB=28) — commander's display name from בוחנים
  var commanderName = '';
  try {
    var examData = getSheet('בוחנים').getDataRange().getValues();
    for (var x = 1; x < examData.length; x++) {
      if (normalizeId(examData[x][1]) === normalizeId(data.examinerId)) { commanderName = String(examData[x][0] || ''); break; }
    }
  } catch(e) {}
  // r35.2: the name read back from 'בוחנים' and the typed reason, as text.
  sheet.getRange(rowIdx, 26).setValue(cellSafe(commanderName + ' (' + normalizeId(data.examinerId) + ')'));
  sheet.getRange(rowIdx, 27).setValue(cellSafe(reason));
  sheet.getRange(rowIdx, 28).setValue(todayStr());
  SpreadsheetApp.flush();
  return jsonResponse({ status: 'ok' });
}

function handleMarkSent(p) {
  // Ownership check — consistent with the other examiner mutations; prevents an
  // authenticated examiner from flipping the "נשלח?" flag on another session's rows.
  if (!examinerOwnsSession(p.sessionCode, p.examinerId)) {
    return jsonResponse({ status: 'error', message: 'אין הרשאה — בוחן לא תואם לסשן' });
  }
  var sheet = getSheet('תוצאות');
  var read = readResultsTail();   // rows of the session the examiner is sending from
  var data = read.rows;
  var wanted = {};
  var ids = p.idNumbers ? p.idNumbers.split(',') : [p.idNumber];
  for (var k = 0; k < ids.length; k++) wanted[normalizeId(ids[k])] = true;
  var count = 0;
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][13]) !== String(p.sessionCode)) continue;
    if (!wanted[normalizeId(data[i][1])]) continue;
    sheet.getRange(i + 1 + read.off, 17).setValue(true);  // נשלח? — column Q (17)
    count++;
  }
  return jsonResponse({ status: 'ok', updated: count });
}

