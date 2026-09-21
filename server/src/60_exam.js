function handleRegisterExamQuestions(data) {
  // Reject the call if the examinee token doesn't match the registered row
  // (legacy rows without a stored token still pass).
  var reqTokenErr = requireExamineeToken(data);
  if (reqTokenErr) return reqTokenErr;
  // Server-side score verification setup. The client tells us which questions
  // came up and how each was shuffled — but NOT which answer is correct. The
  // server looks up the canonical correct index in ANSWER_KEY_BY_LANG and
  // computes the shuffled-correct index itself, so a tampered client cannot
  // claim "answer 0 is always correct" and pass without taking the exam.
  //
  // Expected payload:
  //   { sessionCode, idNumber, language?, questions: [{qIdx, qId, shuffleOrder}] }
  // Where shuffleOrder is an array like [2,0,1,3] meaning:
  //   "displayed answer A = original answer 2, B = original 0, C = original 1, D = original 3".
  //
  // Backward-compat: old clients send {qIdx, correctShuffledIdx} (legacy, trusted).
  // If we detect the legacy shape, we accept it but mark the registration
  // as unverified so submitResult flags the result row accordingly.
  if (!data.sessionCode || !data.idNumber || !data.questions) {
    return jsonResponse({ status: 'error', message: 'חסרים נתונים לרישום מבחן' });
  }

  // If server-side question delivery (getExamQuestions) was used, an entry
  // for `issued_qs_<session>_<id>` will exist in cache. Verify the IDs the
  // client is registering all came from that set — otherwise reject.
  // (No cache entry means legacy flow where the client picked questions
  // locally from questions.js; that path stays open for now.)
  try {
    var issuedKey = 'issued_qs_' + String(data.sessionCode) + '_' + normalizeId(data.idNumber);
    var issuedJson = CacheService.getScriptCache().get(issuedKey);
    if (issuedJson) {
      var issuedSet = {};
      var issuedArr = JSON.parse(issuedJson) || [];
      for (var iz = 0; iz < issuedArr.length; iz++) issuedSet[String(issuedArr[iz])] = true;
      for (var iq = 0; iq < data.questions.length; iq++) {
        var qq = data.questions[iq];
        if (!qq || !qq.qId) continue;
        if (!issuedSet[String(qq.qId)]) {
          return jsonResponse({
            status: 'error',
            message: 'Question ID not in issued set — rejecting registration',
            unexpectedId: qq.qId
          });
        }
      }
    }
  } catch (e) { /* cache failure — fall through to existing flow */ }

  // Verify examinee is in_exam / approved status
  diagMark('sheet:pending-register');
  var pendSheet = getSheet('ממתינים');
  var pendData = pendSheet.getDataRange().getValues();
  var found = false;
  for (var i = pendData.length - 1; i >= 1; i--) {
    if (String(pendData[i][0]) === String(data.sessionCode) && normalizeId(pendData[i][1]) === normalizeId(data.idNumber)) {
      var status = String(pendData[i][5]).trim();
      if (status === 'in_exam' || status === 'approved') {
        found = true;
        break;
      }
    }
  }
  if (!found) {
    return jsonResponse({ status: 'error', message: 'נבחן לא מאושר למבחן' });
  }

  // Build the canonical question map. Each entry stores {qIdx, correctShuffledIdx}
  // — same shape submitResult already consumes — but the correctShuffledIdx is
  // computed server-side whenever possible.
  var lang = String(data.language || 'he').toLowerCase();
  var canonicalMap = [];
  var unverifiedCount = 0;
  var hasAnswerKey = (typeof ANSWER_KEY_BY_LANG !== 'undefined') && (typeof lookupCorrectIndex === 'function');

  for (var qi = 0; qi < data.questions.length; qi++) {
    var q = data.questions[qi];
    if (!q) { canonicalMap.push(null); unverifiedCount++; continue; }

    // Modern shape: client sent qId + shuffleOrder → server computes
    if (hasAnswerKey && q.qId && Array.isArray(q.shuffleOrder)) {
      var origCorrect = lookupCorrectIndex(Number(q.qId), lang);
      if (origCorrect === null || origCorrect === undefined) {
        // Question id missing from answer key → fall back to client's claim if present
        canonicalMap.push({ qIdx: q.qIdx, qId: Number(q.qId), shuffleOrder: q.shuffleOrder, correctShuffledIdx: Number(q.correctShuffledIdx || 0) });
        unverifiedCount++;
        continue;
      }
      var idxInShuffle = q.shuffleOrder.indexOf(Number(origCorrect));
      if (idxInShuffle < 0) {
        // shuffleOrder doesn't contain the correct original index → malformed
        canonicalMap.push({ qIdx: q.qIdx, qId: Number(q.qId), shuffleOrder: q.shuffleOrder, correctShuffledIdx: Number(q.correctShuffledIdx || 0) });
        unverifiedCount++;
        continue;
      }
      // Store qId + shuffleOrder so submitResult can build wrongDetails server-side
      // (needed because we no longer send `ci` to examinees — see handleGetExamQuestions).
      canonicalMap.push({ qIdx: q.qIdx, qId: Number(q.qId), shuffleOrder: q.shuffleOrder, correctShuffledIdx: idxInShuffle });
      continue;
    }

    // Legacy / fallback: client sent correctShuffledIdx directly → trust but flag
    canonicalMap.push({ qIdx: q.qIdx, qId: q.qId ? Number(q.qId) : null, shuffleOrder: Array.isArray(q.shuffleOrder) ? q.shuffleOrder : null, correctShuffledIdx: Number(q.correctShuffledIdx || 0) });
    unverifiedCount++;
  }

  // Guard: if the server answer key is entirely unavailable (answer_key.gs not
  // deployed, or every selected qId missing from it), do NOT register a map full
  // of unverifiable garbage — that is what produced silent 0/30 fails (פינטו/דיין
  // 03/06). Return an error so the confirmed-register client blocks the exam and
  // retries, surfacing the problem instead of mis-scoring a real examinee.
  if (!hasAnswerKey || (data.questions.length > 0 && unverifiedCount >= data.questions.length)) {
    return jsonResponse({ status: 'error', message: 'מפתח התשובות אינו זמין בשרת — פנה למנהל המערכת', keyUnavailable: true });
  }

  // Store in מבחנים sheet (create if needed). Add a fifth column for unverified
  // count so submitResult can flag results scored from unverified data.
  var examSheet;
  try { examSheet = getSheet('מבחנים'); } catch(e) {
    var ss = getSpreadsheet();
    examSheet = ss.insertSheet('מבחנים');
    examSheet.appendRow(['קוד סשן', 'ת.ז.', 'שאלות JSON', 'זמן רישום', 'שפה', 'שגויות לא מאומתות']);
  }
  diagMark('sheet:append-register');
  examSheet.appendRow([
    String(data.sessionCode),
    normalizeId(data.idNumber),
    JSON.stringify(canonicalMap),
    nowISO(),
    lang,
    unverifiedCount
  ]);

  // Layer-1 consolidation: also mark the examinee in_exam here — the same write
  // markExamStarted did — so the start no longer needs a separate markExamStarted
  // round-trip. Only flips 'approved' → 'in_exam' (same guard). Because this is
  // the CONFIRMED/blocking call, it also strengthens the iPhone "stuck in
  // ממתינים" fix. `marked` is returned so the client knows whether to keep the
  // visibilitychange fallback armed.
  var marked = false;
  try {
    var penSheet = pendSheet;
    var penData = refreshExamineePendingRows(penSheet, pendData, data.sessionCode, data.idNumber);
    for (var mi = penData.length - 1; mi >= 1; mi--) {
      if (String(penData[mi][0]) !== String(data.sessionCode) || normalizeId(penData[mi][1]) !== normalizeId(data.idNumber)) continue;
      var mst = String(penData[mi][5]).trim();
      if (mst === 'approved') { penSheet.getRange(mi + 1, 6).setValue('in_exam'); penSheet.getRange(mi + 1, 12).setValue(nowISO()); marked = true; break; }
      if (mst === 'in_exam') { marked = true; break; }
      // other status (cancelled/completed): keep scanning for an approved/in_exam row
    }
  } catch (msErr) { /* non-fatal — client fallback + dashboard cleanup cover it */ }

  return jsonResponse({ status: 'ok', verified: unverifiedCount === 0, unverifiedCount: unverifiedCount, examStarted: marked });
}

function handleSubmitResult(data) {
  // Rate limit: max 5 submissions per minute per (sessionCode, idNumber).
  // One legitimate submission + retries on flaky network; floods are blocked.
  var srRlErr = requireRateLimit('submitResult', String(data.sessionCode || '') + '_' + normalizeId(data.idNumber), 5, 60);
  if (srRlErr) return srRlErr;
  // Require the examinee token before accepting any result. Legacy rows
  // (no stored token) pass through requireExamineeToken with legacy=true.
  diagMark('sheet:token-submit');
  var srTokenErr = requireExamineeToken(data);
  if (srTokenErr) return srTokenErr;
  var sheet = getSheet('תוצאות');

  // Verify examinee is approved (in_exam status) before accepting results
  if (data.sessionCode && data.idNumber) {
    diagMark('sheet:pending-submit');
    var pendSheet = getSheet('ממתינים');
    var pendData = pendSheet.getDataRange().getValues();
    var isApproved = false;
    for (var pi = pendData.length - 1; pi >= 1; pi--) {
      if (String(pendData[pi][0]) === String(data.sessionCode) && normalizeId(pendData[pi][1]) === normalizeId(data.idNumber)) {
        var pStatus = String(pendData[pi][5]).trim();
        // 'cancelled' accepted too: if an examiner reset an examinee who was
        // actually still mid-exam, a genuine finished submit must be RECORDED,
        // not rejected and lost. The fabricated-fail supersede + dup-check below
        // prevent a double-row; close-fails for 'cancelled' are suppressed.
        if (pStatus === 'in_exam' || pStatus === 'approved' || pStatus === 'completed' || pStatus === 'cancelled') {
          isApproved = true;
        }
        break;
      }
    }
    if (!isApproved) {
      return jsonResponse({ status: 'error', message: 'נבחן לא מאושר — לא ניתן לשלוח תוצאות' });
    }
  }

  // SECURITY (anti score-forge): if a registered exam (מבחנים row) exists for this
  // session+id, the answers array is MANDATORY so the server re-scores from the
  // answer key. Without this, an examinee could POST a forged score with NO answers
  // and skip BOTH the re-score and the unverified-guard below (both answers-gated).
  // Keep the full history for old-result recovery and retakes; do not tail-read
  // it. Reuse only within this call, and retry a failed read in the later guards.
  var registeredExamRows = null;
  function readRegisteredExams() {
    var registeredSheet = getSheet('מבחנים');
    if (!registeredExamRows || registeredSheet.getLastRow() !== registeredExamRows.length) {
      // r20 (marks only): prime suspect for the 32s this handler spends after
      // meta:wrong-answers — 'מבחנים' carries the question-map JSON of EVERY
      // exam ever registered, one blob per row, and this pulls all of it to use
      // a single row. Do NOT narrow it blindly: column C IS the answer key the
      // re-score depends on (line ~3461). Measure first.
      diagMark('sheet:registered-submit');
      registeredExamRows = registeredSheet.getDataRange().getValues();
    }
    return registeredExamRows;
  }
  var hasRegisteredExam = false;
  try {
    var regChk = readRegisteredExams();
    for (var rc = regChk.length - 1; rc >= 1; rc--) {
      if (String(regChk[rc][0]) === String(data.sessionCode) && normalizeId(regChk[rc][1]) === normalizeId(data.idNumber)) { hasRegisteredExam = true; break; }
    }
  } catch (regChkErr) {}
  if (hasRegisteredExam && (!data.answers || !Array.isArray(data.answers) || data.answers.length === 0)) {
    return jsonResponse({ status: 'error', message: 'הגשה לא תקינה — חסרות תשובות למבחן רשום' });
  }

  // Server-side score verification: if answers array is present, recalculate
  // score using the question map registered at exam start. The map's
  // correctShuffledIdx values are server-computed (from ANSWER_KEY_BY_LANG)
  // when possible — only fall back to client-claimed values for questions
  // missing from the answer key, in which case we mark the result unverified.
  if (data.answers && Array.isArray(data.answers)) {
    try {
      var examData = readRegisteredExams();
      var questionMap = null;
      var unverifiedCount = 0;
      var registeredLang = '';
      // Find the latest registered exam for this session+ID
      for (var ei = examData.length - 1; ei >= 1; ei--) {
        if (String(examData[ei][0]) === String(data.sessionCode) && normalizeId(examData[ei][1]) === normalizeId(data.idNumber)) {
          questionMap = JSON.parse(examData[ei][2]);
          // Column F (index 5) = unverified-count (added when registerExamQuestions stored this row).
          // Older rows may not have this column → treat as fully unverified to be safe.
          unverifiedCount = (examData[ei].length > 5) ? Number(examData[ei][5] || 0) : questionMap.length;
          // Column E (index 4) = language (added in registerExamQuestions)
          registeredLang = (examData[ei].length > 4) ? String(examData[ei][4] || '') : '';
          break;
        }
      }
      if (questionMap) {
        // Shuffle indexes refer to original answer positions in whatever
        // language the questions were registered in. Translators reorder
        // answers, so the original "correct index" can differ between
        // languages (e.g. he Q128 → idx 1, ar Q128 → idx 2). When the
        // examinee switched language mid-exam, score each answer against
        // the correct index for THE LANGUAGE THEY SAW IT IN, not the one
        // captured at registration.
        function correctIdxForLang(mapEntry, lang) {
          if (!mapEntry || !lang) return null;
          if (typeof lookupCorrectIndex !== 'function') return null;
          if (!mapEntry.qId || !Array.isArray(mapEntry.shuffleOrder)) return null;
          var orig = lookupCorrectIndex(Number(mapEntry.qId), String(lang).toLowerCase());
          if (orig === null || orig === undefined) return null;
          var pos = mapEntry.shuffleOrder.indexOf(Number(orig));
          return pos >= 0 ? pos : null;
        }
        function effectiveCorrectIdx(mapEntry, langAtAnswer) {
          var lang = langAtAnswer ? String(langAtAnswer).toLowerCase() : '';
          if (lang && lang !== registeredLang) {
            var alt = correctIdxForLang(mapEntry, lang);
            if (alt !== null) return alt;
          }
          return Number(mapEntry.correctShuffledIdx);
        }
        var correctCount = 0;
        var totalQ = questionMap.length;
        for (var ai = 0; ai < data.answers.length && ai < totalQ; ai++) {
          if (data.answers[ai] !== null && data.answers[ai] !== undefined && questionMap[ai]) {
            var selected = Number(data.answers[ai].selected);
            var correctIdx = effectiveCorrectIdx(questionMap[ai], data.answers[ai].langAtAnswer);
            if (selected === correctIdx) correctCount++;
          }
        }
        var pct = Math.round((correctCount / totalQ) * 100);
        var passThreshold = Math.ceil(totalQ * 0.86); // ~26/30
        data.score = correctCount;
        data.total = totalQ;
        data.percent = pct;
        data.passed = correctCount >= passThreshold;
        // verified=true ONLY when every question in the map was scored against
        // a server-trusted answer key. Any fallback entry → unverified.
        data.verified = (unverifiedCount === 0);

        // ===== Server-side wrong-answers reconstruction =====
        // Each wrong answer is rendered in the language the examinee was viewing
        // WHEN they answered that specific question (data.answers[i].langAtAnswer).
        // Without this, an examinee who switched mid-exam sees mixed-language
        // feedback that doesn't match what they actually saw.
        try {
          // Lazy per-language cache: questions DB + byId map per language code.
          // Avoids loading every language up-front when most exams use one.
          //
          // r15: this used to call loadQuestionsForLanguageServer — a DRIVE read,
          // on the result-submission hot path. The 'אבחון' sheet caught it live on
          // 2026-09-15: `SLOW POST submitResult 20066 ... drive:he@4000`, i.e. 16
          // of those 20 seconds were Drive, while the examinee's device sat on a
          // 60s deadline and the examiner waited for a result that never arrived.
          // The cached per-license pools hold the same objects, and they are the
          // very pools this exam was served from, so every question the examinee
          // saw is provably in the union — questionMetaForLanguage still falls
          // back to Drive if a pool is missing, so nothing is reconstructed from
          // a partial bank.
          var langDbCache = {}, qMetaMemo = {};
          function getLangDb(lang) {
            var safeLang = String(lang || 'he').toLowerCase();
            if (langDbCache[safeLang]) return langDbCache[safeLang];
            try {
              diagMark('meta:wrong-answers');
              var qs = questionMetaForLanguage(safeLang, qMetaMemo);
              if (!qs || !qs.length) return null;
              var idx = {};
              for (var q = 0; q < qs.length; q++) {
                if (qs[q] && qs[q].id !== undefined) idx[String(qs[q].id)] = qs[q];
              }
              langDbCache[safeLang] = { byId: idx, labels: (safeLang === 'he') ? ['א','ב','ג','ד','ה','ו'] : ['A','B','C','D','E','F'] };
              return langDbCache[safeLang];
            } catch (loadErr) {
              langDbCache[safeLang] = null;
              return null;
            }
          }
          var defaultLang = String(data.language || registeredLang || 'he').toLowerCase();
          // No feedback needs a bank when the authoritative score is perfect.
          // For wrong answers keep the existing default-language fallback intact.
          var allCorrect = totalQ > 0 && correctCount === totalQ;
          var defaultDb = allCorrect ? null : getLangDb(defaultLang);

          var serverWrong = [];
          for (var wi = 0; wi < data.answers.length && wi < questionMap.length; wi++) {
            var mapEntry = questionMap[wi];
            var ans2 = data.answers[wi];
            if (!mapEntry) continue;
            var selected2 = ans2 ? Number(ans2.selected) : -1;
            // Use the per-answer language's correctIdx — same logic as the
            // scoring loop above, so wrong-answer reconstruction matches the
            // pass/fail tally instead of contradicting it after a mid-exam
            // language switch.
            var correctIdx2 = effectiveCorrectIdx(mapEntry, ans2 && ans2.langAtAnswer);
            if (selected2 === correctIdx2) continue; // got it right
            // Pick the language the examinee was viewing when they answered this Q.
            // Falls back to the exam's primary language when missing (old clients).
            var perAnsLang = (ans2 && ans2.langAtAnswer) ? String(ans2.langAtAnswer).toLowerCase() : defaultLang;
            var db = getLangDb(perAnsLang) || defaultDb;
            if (!db || !db.byId) continue;
            var qInfo = (mapEntry.qId !== undefined && mapEntry.qId !== null) ? db.byId[String(mapEntry.qId)] : null;
            if (!qInfo || !Array.isArray(mapEntry.shuffleOrder) || !Array.isArray(qInfo.answers)) continue;
            var shuffled = mapEntry.shuffleOrder.map(function(origIdx) { return qInfo.answers[origIdx]; });
            var yourLabel = '', yourText = '';
            if (selected2 === -1 || selected2 < 0 || selected2 >= shuffled.length) {
              yourText = (perAnsLang === 'he') ? 'לא נענתה' : 'Not answered';
            } else {
              yourLabel = db.labels[selected2] || '';
              yourText = shuffled[selected2] || '';
            }
            var correctLabel = (correctIdx2 >= 0 && correctIdx2 < shuffled.length) ? (db.labels[correctIdx2] || '') : '';
            var correctText = (correctIdx2 >= 0 && correctIdx2 < shuffled.length) ? (shuffled[correctIdx2] || '') : '';
            // Classify the raw question category to the bucket name used by
            // EXAM_STRUCTURE (בטיחות / הכרת הרכב / חוק / תמרורים / ספציפי).
            // Without classification the certificate shows all-100%.
            var classifiedCat = (typeof classifyCategoryServer === 'function')
              ? classifyCategoryServer(qInfo.category)
              : '';
            serverWrong.push({
              question: qInfo.text || ('שאלה ' + (wi + 1)),
              yourAnswer: yourLabel ? (yourLabel + ' - ' + yourText) : yourText,
              correctAnswer: correctLabel ? (correctLabel + ' - ' + correctText) : correctText,
              category: classifiedCat || qInfo.category || ''
            });
          }
          // Always replace client-provided wrongAnswers — server is authoritative.
          if (allCorrect || defaultDb) data.wrongAnswers = serverWrong;
        } catch (rwe) {
          // Reconstruction failed (Drive load, etc.) — keep whatever client sent
          // rather than wiping it. Log for diagnosis.
          try { Logger.log('wrong-answer rebuild failed: ' + (rwe && rwe.message)); } catch(_) {}
        }
      }
    } catch(ve) {
      // If verification fails, fall through to client-provided score with flag
      data.verified = false;
    }
  }

  // Server-side timing check: if exam took less than 3 minutes, flag as suspicious
  if (data.sessionCode && data.idNumber) {
    try {
      var examData2 = readRegisteredExams();
      for (var ti = examData2.length - 1; ti >= 1; ti--) {
        if (String(examData2[ti][0]) === String(data.sessionCode) && normalizeId(examData2[ti][1]) === normalizeId(data.idNumber)) {
          var regTime = new Date(examData2[ti][3]);
          var elapsed = (new Date() - regTime) / 1000; // seconds
          if (elapsed < 180 && elapsed > 0) { // less than 3 minutes
            data.suspicious = true;
          }
          break;
        }
      }
    } catch(te) {}
  }

  // Guard: answers were present but the server never re-scored (no מבחנים row /
  // questionMap missing → data.verified left undefined above). Do NOT silently
  // trust the client's score — it is computed from a deliberately-stripped `ci`
  // and can be garbage (the historical false-0/30). Flag the row unverified so
  // the examiner reviews it instead of recording a bogus pass/fail.
  if (data.answers && Array.isArray(data.answers) && typeof data.verified === 'undefined') {
    data.verified = false;
    data.scoreUnverified = true;
  }

  // ===== Supersede any SYSTEM-FABRICATED fail for this session+id =====
  // A real finished submit must WIN over a system-written fail — the browser-
  // close beacon (handleSubmitFailOnClose) or the dashboard timeout/disconnect
  // row — created while the examinee was offline/backgrounded. Those are 'נכשל'
  // rows whose note carries a machine marker. Match on session+id ONLY (NOT
  // language/license): the fabricated row is stamped with the REGISTRATION
  // language, but the real submit may carry a DIFFERENT final language after a
  // mid-exam switch (Russian/Arabic/Amharic examinees), so a language-scoped
  // match would miss it and the dup-check below would swallow the real result →
  // a false 0/30 "vanished" exam, especially on iOS. Mark them בוטל (audit kept).
  // Mirrors the פסול-supersede pass below; genuine real נכשל rows lack the marker.
  diagMark('sheet:results-submit');
  var fabRows = sheet.getDataRange().getValues();
  var fabSuperseded = false;
  for (var fb = 1; fb < fabRows.length; fb++) {
    if (String(fabRows[fb][13]) !== String(data.sessionCode)) continue;
    if (normalizeId(fabRows[fb][1]) !== normalizeId(data.idNumber)) continue;
    if (String(fabRows[fb][7]).trim() !== 'נכשל') continue;
    var fbNote = String(fabRows[fb][15] || '');
    // markers: close-beacon ('סגירת דפדפן'), dashboard timeout ('טיימאאוט'), and
    // examiner manual-disconnect ('סיום ידני ... ניתוק/תקלה'). All three mean "did
    // not finish properly" — a real finished submit must override them.
    if (fbNote.indexOf('סגירת דפדפן') === -1 && fbNote.indexOf('טיימאאוט') === -1 && fbNote.indexOf('סיום ידני') === -1) continue;
    sheet.getRange(fb + 1, 8).setValue('בוטל');                                    // H = pass/fail
    sheet.getRange(fb + 1, 27).setValue('בוטל אוטומטית — הנבחן השלים והגיש מבחן');  // AA = reason
    sheet.getRange(fb + 1, 28).setValue(todayStr());                               // AB = correction date
    // The duplicate and attempt checks below reuse this complete snapshot.
    fabRows[fb][7] = 'בוטל';
    fabSuperseded = true;
  }
  if (fabSuperseded) SpreadsheetApp.flush();

  // Duplicate protection: check if result already exists for this session+ID+license+language
  // Skip disqualified (פסול) and cancelled (בוטל) rows — those are not real results and should not block retakes
  // Also skip duplicate check entirely if examinee has an active in_exam pending row (retake after DQ)
  var hasPendingInExam = false;
  // Re-check current status after potentially slow scoring/bank work. New rows
  // or changed row positions require a full read; otherwise only this person's
  // rows need refreshing. Never overwrite a newer examiner decision.
  pendData = refreshExamineePendingRows(pendSheet, pendData, data.sessionCode, data.idNumber);
  var pendCheck = pendData;
  for (var pc = pendCheck.length - 1; pc >= 1; pc--) {
    if (String(pendCheck[pc][0]) === String(data.sessionCode) && normalizeId(pendCheck[pc][1]) === normalizeId(data.idNumber) && String(pendCheck[pc][5]).trim() === 'in_exam') {
      hasPendingInExam = true;
      break;
    }
  }
  if (!hasPendingInExam) {
    diagMark('sheet:results-submit-2');
    var existingData = sheet.getDataRange().getValues();
    for (var d = 1; d < existingData.length; d++) {
      var existingStatus = String(existingData[d][7] || '').trim();
      if (existingStatus === 'פסול' || existingStatus === 'בוטל') continue;
      if (String(existingData[d][13]) === String(data.sessionCode) && normalizeId(existingData[d][1]) === normalizeId(data.idNumber) && String(existingData[d][4]) === String(data.license) && String(existingData[d][12]) === String(data.language || 'he')) {
        // Genuine prior real result for this exact exam — a true duplicate.
        // (Fabricated close/timeout fails were already superseded to בוטל above
        // and are skipped by the status filter, so they can't masquerade here.)
        markPendingCompleted(data.sessionCode, data.idNumber, { sheet: pendSheet, rows: pendData });
        return jsonResponse({ status: 'ok', waLink: existingData[d][18] || '', duplicate: true });
      }
    }
  }

  // Belt-and-suspenders: never let the literal "undefined" reach the certificate.
  // The client never receives `ci`, so its locally-built wrongAnswers carry
  // "undefined - undefined" as the correct answer; the server normally rebuilds
  // them, but if that failed (Drive/cache down) the client text is kept. Replace
  // any "undefined" with a neutral placeholder so feedback is never garbled.
  if (Array.isArray(data.wrongAnswers)) {
    for (var sw = 0; sw < data.wrongAnswers.length; sw++) {
      var swItem = data.wrongAnswers[sw];
      if (swItem && typeof swItem.correctAnswer === 'string' && swItem.correctAnswer.indexOf('undefined') !== -1) swItem.correctAnswer = '(לא זמין כעת)';
      if (swItem && typeof swItem.yourAnswer === 'string' && swItem.yourAnswer.indexOf('undefined') !== -1) swItem.yourAnswer = '(לא זמין)';
    }
  }

  var wrongDetails = '';
  var wrongForWA = '';
  if (data.wrongAnswers && data.wrongAnswers.length > 0) {
    for (var i = 0; i < data.wrongAnswers.length; i++) {
      var w = data.wrongAnswers[i];
      // Question ID prefix lets the commander dashboard aggregate by the exact
      // question (not just generic text "מה פירוש התמרור?" that collapses 50+
      // distinct sign questions into one row). Backward-compatible — readers
      // tolerate the line being missing for legacy rows.
      if (w.questionId) wrongDetails += 'מזהה שאלה: ' + w.questionId + '\n';
      wrongDetails += 'שאלה: ' + w.question + '\n';
      wrongDetails += 'תשובת הנבחן: ' + w.yourAnswer + '\n';
      wrongDetails += 'תשובה נכונה: ' + w.correctAnswer + '\n';
      if (w.category) wrongDetails += 'קטגוריה: ' + w.category + '\n';
      wrongDetails += '\n';

      wrongForWA += '❌ ' + w.question + '\n';
      wrongForWA += 'ענית: ' + w.yourAnswer + '\n';
      wrongForWA += '✅ נכון: ' + w.correctAnswer + '\n\n';
    }
  }

  // Surface the unverified-score guard (set above) loudly in the stored detail.
  if (data.scoreUnverified) {
    wrongDetails = '⚠️ ציון לא אומת בשרת (רישום מבחן חסר) — נדרש אימות ידני\n\n' + wrongDetails;
  }

  var passText = data.passed ? 'עבר' : 'נכשל';
  var waMessage = '*🚗 אישור תוצאת מבחן תאוריה חיצוני*\n\n' +
    'שם: ' + data.fullName + '\n' +
    'ת.ז.: ' + data.idNumber + '\n' +
    'דרגה: ' + data.license + '\n' +
    (data.population ? 'אוכלוסיה: ' + data.population + '\n' : '') +
    'תאריך: ' + todayStr() + '\n' +
    'תוצאה: *' + passText + '* (' + data.score + '/' + data.total + ')\n' +
    'זמן: ' + data.time + '\n';

  var wrongCount = Number(data.total) - Number(data.score);
  if (data.wrongAnswers && data.wrongAnswers.length > 0) {
    waMessage += '\n*שאלות שגויות (' + data.wrongAnswers.length + '):*\n\n' + wrongForWA;
  } else if (wrongCount === 0) {
    waMessage += '\nכל התשובות נכונות! 🎉';
  }

  var phone = formatPhoneForWA(data.phone);

  // Count attempt number for this examinee + license combination
  var attemptNum = countAttempts(data.idNumber, data.license, fabRows, sheet) + 1;

  var waMessage2 = waMessage; // preserve for link
  if (attemptNum > 1) {
    waMessage2 = waMessage + 'ניסיון: ' + attemptNum + '\n';
  }
  var waLink = 'https://wa.me/' + phone + '?text=' + encodeURIComponent(waMessage2);

  // Format language history into a readable path. Single language = just the
  // code (e.g. "he"). Multiple = arrow-joined (e.g. "he → ru → he") so the
  // examiner can see at a glance that the examinee switched languages.
  var langPath = '';
  if (Array.isArray(data.languageHistory) && data.languageHistory.length > 0) {
    langPath = data.languageHistory.length === 1
      ? String(data.languageHistory[0])
      : data.languageHistory.join(' → ');
  } else {
    langPath = data.language || 'he';
  }

  // Supersede any prior פסול row for THIS session+id. Scenario: examinee was
  // auto-DQ'd, the overturn flow didn't finish (examiner clicked אשר, or hit
  // a stale "תוצאה לא נמצאה" path), then the examinee was re-allowed in and
  // finished the exam. Without this cleanup the sheet ends up with both a
  // פסול row AND a עבר/נכשל row — which is what happened at base 14 today.
  // We mark the old row as בוטל (audit trail preserved) and log the reason.
  // Preserve the late complete read: another submission may have completed
  // since scoring. In particular, do not move final retry detection earlier.
  diagMark('sheet:results-submit-3');
  var existingRows = sheet.getDataRange().getValues();
  for (var ex = existingRows.length - 1; ex >= 1; ex--) {
    if (String(existingRows[ex][13]) === String(data.sessionCode) &&
        normalizeId(existingRows[ex][1]) === normalizeId(data.idNumber) &&
        String(existingRows[ex][7]).trim() === 'פסול') {
      sheet.getRange(ex + 1, 8).setValue('בוטל');           // H = pass/fail
      sheet.getRange(ex + 1, 18).setValue(false);            // R = disqualified flag
      sheet.getRange(ex + 1, 27).setValue('בוטל אוטומטית — נבחן ניגש למבחן מחדש'); // AA = reason
      sheet.getRange(ex + 1, 28).setValue(todayStr());       // AB = correction date
      existingRows[ex][7] = 'בוטל';
      existingRows[ex][17] = false;
    }
  }

  // Idempotency: skip if an identical result row already exists. A retry/resend
  // whose original response was lost (flaky network) would otherwise create a
  // duplicate. Matches session+id+license+score+time+result, so a re-take or a
  // post-overturn submit (different score/time/result) is still appended.
  for (var dc = existingRows.length - 1; dc >= 1; dc--) {
    if (String(existingRows[dc][13]) === String(data.sessionCode) &&
        normalizeId(existingRows[dc][1]) === normalizeId(data.idNumber) &&
        String(existingRows[dc][4]) === String(data.license) &&
        String(existingRows[dc][5]) === (data.score + '/' + data.total) &&
        String(existingRows[dc][7]).trim() === String(passText).trim() &&
        String(existingRows[dc][8]) === String(data.time)) {
      markPendingCompleted(data.sessionCode, data.idNumber, { sheet: pendSheet, rows: pendData });
      return jsonResponse({ status: 'ok', duplicate: true, waLink: waLink });
    }
  }

  sheet.appendRow([
    todayStr(),
    data.idNumber,
    data.fullName,
    data.phone,
    data.license,
    data.score + '/' + data.total,
    data.percent + '%',
    passText,
    data.time,
    data.examinerName || '',
    data.site || '',
    data.classroom || '',
    data.language || 'he',
    data.sessionCode || '',
    attemptNum,
    wrongDetails,
    false,
    false,
    waLink,
    data.population || '',
    false,
    data.audioMode || 'off',
    data.verified ? 'מאומת' : '',
    data.suspicious ? 'חשוד' : '',
    '',                                 // Y (24) dqEventId — not a DQ row
    '',                                 // Z (25) תוקן ע"י — empty (no correction yet)
    '',                                 // AA (26) סיבת תיקון — empty
    '',                                 // AB (27) תאריך תיקון — empty
    langPath,                           // AC (28) מסלול שפות — full path he → ru → he
    String(data.device || '')           // AD (29) מכשיר — phone / tablet / desktop
  ]);

  // Update pending status to completed
  markPendingCompleted(data.sessionCode, data.idNumber, { sheet: pendSheet, rows: pendData });

  diagMark('compute:submit-done');
  return jsonResponse({ status: 'ok', waLink: waLink });
}

function handleSubmitWrongAnswers(p) {
  var swaTokenErr = requireExamineeToken(p);
  if (swaTokenErr) return swaTokenErr;
  // Append a single wrong answer item to existing result row
  var sheet = getSheet('תוצאות');
  var data = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][13]) === String(p.sessionCode) && normalizeId(data[i][1]) === normalizeId(p.idNumber)) {
      var existing = String(data[i][15] || '');

      // New format: individual item with question/yourAnswer/correctAnswer params
      if (p.question) {
        var line = 'שאלה: ' + p.question + '\n' +
                   'תשובת הנבחן: ' + p.yourAnswer + '\n' +
                   'תשובה נכונה: ' + p.correctAnswer + '\n\n';
        sheet.getRange(i + 1, 16).setValue(existing + line);
        SpreadsheetApp.flush();
        return jsonResponse({ status: 'ok' });
      }

      // Legacy format: chunk with JSON array
      var chunk = p.chunk || '';
      var totalChunks = Number(p.totalChunks) || 1;
      if (totalChunks === 1) {
        try {
          var wrongArr = JSON.parse(chunk);
          var formatted = '';
          for (var w = 0; w < wrongArr.length; w++) {
            formatted += 'שאלה: ' + wrongArr[w].question + '\n';
            formatted += 'תשובת הנבחן: ' + wrongArr[w].yourAnswer + '\n';
            formatted += 'תשובה נכונה: ' + wrongArr[w].correctAnswer + '\n\n';
          }
          sheet.getRange(i + 1, 16).setValue(formatted);
        } catch(ex) {
          sheet.getRange(i + 1, 16).setValue(existing + chunk);
        }
      } else {
        sheet.getRange(i + 1, 16).setValue(existing + chunk);
      }
      return jsonResponse({ status: 'ok' });
    }
  }
  return jsonResponse({ status: 'error', message: 'Result row not found for wrong answers' });
}

function handleSubmitWrongAnswersBulk(data) {
  var swabTokenErr = requireExamineeToken(data);
  if (swabTokenErr) return swabTokenErr;
  // Receive ALL wrong answers in a single POST and write to result row
  var sheet = getSheet('תוצאות');
  var rows = sheet.getDataRange().getValues();
  for (var i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][13]) === String(data.sessionCode) && normalizeId(rows[i][1]) === normalizeId(data.idNumber)) {
      // Detect the bug pattern where client sent "undefined" because it doesn't
      // know correct answers (handleGetExamQuestions strips `ci` from examinee
      // responses). If client data is bogus AND the sheet already has good data
      // (written by handleSubmitResult's server-side rebuild), keep the sheet's version.
      var clientHasBogus = false;
      if (data.wrongAnswers && data.wrongAnswers.length > 0) {
        for (var bi = 0; bi < data.wrongAnswers.length; bi++) {
          var ca = String((data.wrongAnswers[bi] && data.wrongAnswers[bi].correctAnswer) || '');
          if (ca.indexOf('undefined') !== -1) { clientHasBogus = true; break; }
        }
      }
      var existingWrong = String(rows[i][15] || '');
      if (clientHasBogus && existingWrong && existingWrong.indexOf('undefined') === -1) {
        // Sheet already has authoritative data → keep it, skip overwrite.
        SpreadsheetApp.flush();
        return jsonResponse({ status: 'ok', skipped: true, reason: 'client_bogus_server_good' });
      }

      var wrongDetails = '';
      var wrongForWA = '';
      if (data.wrongAnswers && data.wrongAnswers.length > 0) {
        for (var w = 0; w < data.wrongAnswers.length; w++) {
          var item = data.wrongAnswers[w];
          // Question ID prefix — see comment in handleSubmitWrongAnswers above.
          if (item.questionId) wrongDetails += 'מזהה שאלה: ' + item.questionId + '\n';
          wrongDetails += 'שאלה: ' + item.question + '\n';
          wrongDetails += 'תשובת הנבחן: ' + item.yourAnswer + '\n';
          wrongDetails += 'תשובה נכונה: ' + item.correctAnswer + '\n';
          if (item.category) wrongDetails += 'קטגוריה: ' + item.category + '\n';
          wrongDetails += '\n';
          wrongForWA += '❌ ' + item.question + '\n';
          wrongForWA += 'ענית: ' + item.yourAnswer + '\n';
          wrongForWA += '✅ נכון: ' + item.correctAnswer + '\n\n';
        }
      }
      // Update wrong details column
      sheet.getRange(i + 1, 16).setValue(wrongDetails);

      // Regenerate WA link with wrong answers included
      var isCorrected = rows[i][20] === true || String(rows[i][20]) === 'TRUE';
      if (!isCorrected && data.wrongAnswers && data.wrongAnswers.length > 0) {
        var passText = String(rows[i][7] || 'נכשל');
        var phone = formatPhoneForWA(rows[i][3]);
        var waMsg = '*🚗 אישור תוצאת מבחן תאוריה חיצוני*\n\n' +
          'שם: ' + rows[i][2] + '\n' +
          'ת.ז.: ' + rows[i][1] + '\n' +
          'דרגה: ' + rows[i][4] + '\n' +
          (rows[i][19] ? 'אוכלוסיה: ' + rows[i][19] + '\n' : '') +
          'תאריך: ' + rows[i][0] + '\n' +
          'תוצאה: *' + passText + '* (' + rows[i][5] + ')\n' +
          'זמן: ' + rows[i][8] + '\n' +
          '\n*שאלות שגויות (' + data.wrongAnswers.length + '):*\n\n' + wrongForWA;
        var attemptNum = rows[i][14] || 1;
        if (attemptNum > 1) waMsg += 'ניסיון: ' + attemptNum + '\n';
        var waLink = 'https://wa.me/' + phone + '?text=' + encodeURIComponent(waMsg);
        sheet.getRange(i + 1, 19).setValue(waLink);
      }

      SpreadsheetApp.flush();
      return jsonResponse({ status: 'ok', count: data.wrongAnswers ? data.wrongAnswers.length : 0 });
    }
  }
  return jsonResponse({ status: 'error', message: 'Result row not found for wrong answers' });
}

function handleSubmitFailOnClose(data) {
  var focTokenErr = requireExamineeToken(data);
  if (focTokenErr) return focTokenErr;
  var sheet = getSheet('תוצאות');

  // Do NOT record a close-fail for an examinee an examiner reset/removed
  // (status 'cancelled') or 'rejected' — reset semantics are "won't count as a
  // fail". A clean finish of such an examinee IS still recorded (submitResult
  // accepts 'cancelled'); only the auto-0/30-on-close is suppressed here.
  try {
    var focPend = getSheet('ממתינים').getDataRange().getValues();
    for (var fp = focPend.length - 1; fp >= 1; fp--) {
      if (String(focPend[fp][0]) === String(data.sessionCode) && normalizeId(focPend[fp][1]) === normalizeId(data.idNumber)) {
        var fpStatus = String(focPend[fp][5]).trim();
        if (fpStatus === 'cancelled' || fpStatus === 'rejected') return jsonResponse({ status: 'ok', skipped: 'cancelled' });
        break;
      }
    }
  } catch (focErr) {}

  // Duplicate protection: if ANY non-בוטל result already exists for this session+id,
  // do NOT add a close-fail. Match on session+id ONLY (not language/license): the
  // close-beacon carries the REGISTRATION language, but a real submit may carry a
  // different FINAL language after a mid-exam switch — a language-scoped check would
  // miss it and append a spurious 0/30 next to the real result.
  var existingData = sheet.getDataRange().getValues();
  for (var d = 1; d < existingData.length; d++) {
    if (String(existingData[d][13]) === String(data.sessionCode) && normalizeId(existingData[d][1]) === normalizeId(data.idNumber)) {
      if (String(existingData[d][7] || '').trim() === 'בוטל') continue; // a voided row is not a real result
      markPendingCompleted(data.sessionCode, data.idNumber);
      return jsonResponse({ status: 'ok', duplicate: true });
    }
  }

  var attemptNum = countAttempts(data.idNumber, data.license || '') + 1;

  sheet.appendRow([
    todayStr(),
    data.idNumber,
    data.fullName,
    data.phone,
    data.license || '',
    '0/' + (data.totalQuestions || 30),
    '0%',
    'נכשל',
    data.time || '00:00',
    data.examinerName || '',
    data.site || '',
    data.classroom || '',
    data.language || 'he',
    data.sessionCode || '',
    attemptNum,
    'סגירת דפדפן באמצע מבחן (נענו ' + (data.answeredCount || 0) + ' שאלות)',
    false,
    false,
    '',
    data.population || '',
    false,
    data.audioMode || 'off',
    '', '', '', '', '', '', '',         // idx 22-28 (מאומת..מסלול שפות) — N/A for a close-fail row
    String(data.device || '')           // AD (29) מכשיר — phone / tablet / desktop
  ]);

  markPendingCompleted(data.sessionCode, data.idNumber);

  return jsonResponse({ status: 'ok' });
}

function handleCancelFailOnClose(data) {
  // Called when page reloads (refresh, not actual close) — undo the fail
  var cfocTokenErr = requireExamineeToken(data);
  if (cfocTokenErr) return cfocTokenErr;
  var sc = String(data.sessionCode || '');
  var id = normalizeId(data.idNumber || '');
  if (!sc || !id) return jsonResponse({ status: 'ok' });

  var sheet = getSheet('תוצאות');
  var rows = sheet.getDataRange().getValues();
  // Find the most recent row for this session+ID that is a "close" fail
  for (var r = rows.length - 1; r >= 1; r--) {
    if (String(rows[r][13]) === sc && normalizeId(rows[r][1]) === id) {
      var notes = String(rows[r][15] || '');
      if (notes.indexOf('\u05E1\u05D2\u05D9\u05E8\u05EA \u05D3\u05E4\u05D3\u05E4\u05DF') !== -1) {
        // Examinee resumed — the fail-on-close was premature. Mark בוטל instead
        // of DELETING: sheet.deleteRow was the ONLY path that could ever destroy a
        // result row (latent "vanished result" vector). At most one בוטל row
        // results, since submitFailOnClose's dup-guard blocks further close-fails.
        sheet.getRange(r + 1, 8).setValue('בוטל');
        sheet.getRange(r + 1, 27).setValue('בוטל אוטומטי - רענון/חזרה למבחן');
        sheet.getRange(r + 1, 28).setValue(todayStr());
        // Also un-mark pending as completed so exam can continue
        unmarkPendingCompleted(sc, id);
      }
      break; // only check the most recent match
    }
  }
  return jsonResponse({ status: 'ok' });
}

