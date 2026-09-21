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
    for (var i = data.length - 1; i >= 1; i--) {
      if (String(data[i][0]) === String(p.sessionCode) && normalizeId(data[i][1]) === normalizeId(p.idNumber)) {
        var st = String(data[i][5] || '').trim();
        if (st === 'in_exam' || st === 'approved') {
          var prev = (data[i].length > 15) ? (Number(data[i][15]) || 0) : 0;
          sheet.getRange(i + 1, 16).setValue(prev + 1);                                   // col 16 (idx15) = warnings count
          if (p.reason) sheet.getRange(i + 1, 17).setValue(String(p.reason).slice(0, 40)); // col 17 (idx16) = last reason
        }
        break;
      }
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
function extraMinutesKey(sessionCode) { return QUESTION_CACHE_PREFIX + 'extmin_' + String(sessionCode || '').trim(); }
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
  if (!verifyToken(p.examinerId, p.token)) {
    return jsonResponse({ status: 'error', message: 'טוקן בוחן לא תקין', tokenExpired: true });
  }
  if (!verifyExaminerForSession(p.sessionCode, p.examinerId)) {
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
  var name = '', found = false;
  for (var j = pendData.length - 1; j >= 1; j--) {
    if (String(pendData[j][0]).trim() === String(p.sessionCode).trim() && normalizeId(pendData[j][1]) === normalizeId(p.idNumber)) {
      name = pendData[j][2] || '';
      found = true;
      break;
    }
  }
  if (!found) return jsonResponse({ status: 'error', message: 'נבחן לא נמצא בסשן' });

  // Examiner display name for the audit row.
  var examinerName = '';
  try {
    var sData = getSheet('סשנים').getDataRange().getValues();
    for (var s = 1; s < sData.length; s++) {
      if (String(sData[s][0]).trim() === String(p.sessionCode).trim()) { examinerName = sData[s][2] || ''; break; }
    }
  } catch (e) {}

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
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]).trim() === String(p.sessionCode).trim() && normalizeId(data[i][1]) === normalizeId(p.idNumber)) {
      var storedToken = String((data[i].length > 12 ? data[i][12] : '') || '').trim();
      if (storedToken && p.examineeToken && String(p.examineeToken).trim() !== storedToken) {
        return jsonResponse({ status: 'error', examineeTokenError: 'mismatch' });
      }
      if (String(data[i][5]).trim() === 'in_exam') {
        if (pendSheet.getMaxColumns() < 19) pendSheet.insertColumnsAfter(pendSheet.getMaxColumns(), 19 - pendSheet.getMaxColumns());
        if (!String(pendSheet.getRange(1, 19).getValue() || '').trim()) pendSheet.getRange(1, 19).setValue('סיים במכשיר');
        pendSheet.getRange(i + 1, 19).setValue(nowISO());
      }
      return jsonResponse({ status: 'ok' });
    }
  }
  return jsonResponse({ status: 'ok' });  // no matching row — harmless no-op
}

