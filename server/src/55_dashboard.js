function handleExaminerDashboard(p) {
  var code = String(p.sessionCode);
  var pendSheet = getSheet('ממתינים');
  var resSheet = getSheet('תוצאות');

  // Tail reads (see readTail). pendOff shifts the one row-index write below;
  // resSheet is only ever appended to in this handler, so it needs no offset.
  // r15: instrumented. This is THE exam hot path — every examiner polls it every
  // 2-5s for the whole exam — and on 2026-09-15 it recorded 27.9s with no marks
  // at all, so the trail could not say where the time went. It re-reads the
  // results tail three times per request (here, after a state change, and before
  // the completed list); the marks will finally price that.
  diagMark('sheet:pending-dash');
  var _pendT = readTail(pendSheet, 4);
  var pendData = _pendT.rows, pendOff = _pendT.off;
  diagMark('sheet:results-dash');
  var resData = readTail(resSheet, 0).rows;
  diagMark('sheet:extensions-dash');
  var pending = [];
  var active = [];

  // Sum of mid-exam time grants per examinee (minutes) — extends the stale/timeout
  // threshold below and is shown as a badge in the active list. One read, by id.
  var extraMinById = {};
  try { extraMinById = extraMinutesBySession(code); } catch (e) {}   // r23: cached 30 s, dropped by addExamTime

  // ---- Pre-built indexes (perf) ----------------------------------------------
  // handleExaminerDashboard runs every 5s per examiner. The old code re-scanned
  // the WHOLE תוצאות/ממתינים sheets inside per-examinee loops, making it
  // O(examinees × תוצאות) — which silently degraded as תוצאות grew each day and
  // spiked during the morning registration rush. These indexes reproduce EXACTLY
  // what those inner scans computed, but build once → O(1) lookups.
  //   resBySessId[id]  = { dqResults, otherResults }  (this session, non-בוטל)
  //   pendTermBySessId[id] = { dqTerminals, otherTerminals } (this session)
  function buildResBySessId(rows) {
    var idx = {};
    for (var r = 1; r < rows.length; r++) {
      if (String(rows[r][13]) !== code) continue;
      if (String(rows[r][7] || '') === 'בוטל') continue;
      var k = normalizeId(rows[r][1]);
      if (!idx[k]) idx[k] = { dqResults: 0, otherResults: 0 };
      if (String(rows[r][7] || '').trim() === 'פסול') idx[k].dqResults++;
      else idx[k].otherResults++;
    }
    return idx;
  }
  function buildPendTermBySessId(rows) {
    var idx = {};
    for (var r = 1; r < rows.length; r++) {
      if (String(rows[r][0]) !== code) continue;
      var k = normalizeId(rows[r][1]);
      if (!idx[k]) idx[k] = { dqTerminals: 0, otherTerminals: 0 };
      var st = String(rows[r][5]).trim();
      if (st === 'disqualified' || st === 'dq_confirmed') idx[k].dqTerminals++;
      else if (st === 'completed') idx[k].otherTerminals++;
    }
    return idx;
  }
  var resBySessId = buildResBySessId(resData);
  var pendTermBySessId = buildPendTermBySessId(pendData);

  // Auto-cleanup: detect stale in_exam entries that already have a result or are way past exam time
  var now = new Date();
  var BASE_EXAM_MS = 40 * 60 * 1000;
  var STALE_BUFFER_MS = 20 * 60 * 1000; // 20 minutes buffer (approval wait + instructions)
  for (var ci = 1; ci < pendData.length; ci++) {
    if (String(pendData[ci][0]) !== code) continue;
    // Reconcile stuck 'in_exam' AND 'approved' entries. 'approved' that never
    // advanced to 'in_exam' happens when markExamStarted failed on the device
    // (common on iOS) — leaving the examinee stuck in "ממתינים" forever, even
    // after finishing. We clear those once a result exists for them (below).
    var _ciStatus = String(pendData[ci][5]).trim();
    if (_ciStatus !== 'in_exam' && _ciStatus !== 'approved') continue;
    var _startedExam = (_ciStatus === 'in_exam');
    var ciId = pendData[ci][1];
    var examStart = pendData[ci][11] ? new Date(pendData[ci][11]) : null;
    var regTime = examStart || (pendData[ci][4] ? new Date(pendData[ci][4]) : null);
    // Dynamic stale threshold: exam time (based on extension) + buffer
    var ciExt = parseFloat(pendData[ci][10]) || 1;
    if (ciExt !== 1.25 && ciExt !== 1.5) ciExt = 1;
    var maxMs = Math.round(BASE_EXAM_MS * ciExt) + STALE_BUFFER_MS + ((extraMinById[normalizeId(ciId)] || 0) * 60 * 1000);
    var isStale = regTime && (now.getTime() - regTime.getTime() > maxMs);
    // Only someone who actually STARTED (in_exam) can time out into a fail. A
    // stale 'approved' never started → never fabricate a 0/30 fail for it; it is
    // only reconciled when a real result already exists.
    var effectiveStale = isStale && _startedExam;

    // Count results by type for this examinee in this session (indexed lookup —
    // was a full scan of תוצאות per examinee).
    var _rc = resBySessId[normalizeId(ciId)] || { dqResults: 0, otherResults: 0 };
    var dqResults = _rc.dqResults, otherResults = _rc.otherResults;
    // Count terminal entries by type in pending sheet for this examinee (indexed —
    // was a full scan of ממתינים per examinee).
    var _pt = pendTermBySessId[normalizeId(ciId)] || { dqTerminals: 0, otherTerminals: 0 };
    var dqTerminals = _pt.dqTerminals, otherTerminals = _pt.otherTerminals;
    // Cap DQ results to DQ terminals — handles duplicate פסול rows from page refreshes
    var effectiveResults = Math.min(dqResults, dqTerminals) + otherResults;
    var totalTerminals = dqTerminals + otherTerminals;
    var hasUnmatchedResult = effectiveResults > totalTerminals;

    if (hasUnmatchedResult || effectiveStale) {
      // Fix dangling status — mark as completed
      pendSheet.getRange(ci + 1 + pendOff, 6).setValue('completed');
      pendData[ci][5] = 'completed'; // update local copy
      // Keep pendTermBySessId in sync: this row was in_exam/approved (loop guard
      // above) → now a 'completed' terminal, so a fresh rescan would count it here.
      var _mk = normalizeId(ciId);
      if (!pendTermBySessId[_mk]) pendTermBySessId[_mk] = { dqTerminals: 0, otherTerminals: 0 };
      pendTermBySessId[_mk].otherTerminals++;
      if (effectiveStale && !hasUnmatchedResult) {
        // Create a timeout fail result
        var sesData2 = getSheet('סשנים').getDataRange().getValues();
        var license2 = pendData[ci][8] || '', site2 = '', classroom2 = '', examinerName2 = '', language2 = pendData[ci][6] || 'he';
        for (var si = 1; si < sesData2.length; si++) {
          if (String(sesData2[si][0]).trim() === code) {
            examinerName2 = sesData2[si][2] || '';
            site2 = sesData2[si][3] || '';
            classroom2 = sesData2[si][4] || '';
            if (!license2) license2 = sesData2[si][5] || '';
            break;
          }
        }
        var attemptNum2 = countAttempts(String(ciId), license2) + 1;
        resSheet.appendRow([
          todayStr(), ciId, pendData[ci][2] || '', pendData[ci][3] || '', license2,
          '0/30', '0%', 'נכשל', '', examinerName2,
          site2, classroom2, language2, code,
          attemptNum2, 'ניתוק/טיימאאוט — הנבחן לא סיים את המבחן', false, false, '',
          pendData[ci][7] || '', false, pendData[ci][9] || 'off'
        ]);
        // Refresh resData after append, and rebuild the results index so later
        // iterations' counts include the row just appended (behavior-identical to
        // the old per-iteration rescan of the freshly re-read sheet).
        resData = readTail(resSheet, 0).rows;
        resBySessId = buildResBySessId(resData);
      }
    }
  }

  // Pre-compute attempts-today by examinee id (for "second attempt today" warning).
  // Counts non-disqualified terminal entries (real attempts) made today regardless
  // of which session — so an examinee who tried earlier today in another session
  // also triggers the warning.
  var attemptsTodayById = {};
  var todayDateStr = (function() {
    var d = new Date();
    return d.getFullYear() + '-' + (d.getMonth() + 1) + '-' + d.getDate();
  })();
  function isToday(val) {
    if (!val) return false;
    try {
      var d = (val instanceof Date) ? val : new Date(val);
      if (isNaN(d.getTime())) return false;
      return (d.getFullYear() + '-' + (d.getMonth() + 1) + '-' + d.getDate()) === todayDateStr;
    } catch(_) { return false; }
  }
  for (var ai2 = 1; ai2 < resData.length; ai2++) {
    if (!isToday(resData[ai2][0])) continue;
    var aiPassed = String(resData[ai2][7] || '').trim();
    if (aiPassed === 'בוטל') continue; // overturned, not a real attempt
    var aiId = normalizeId(resData[ai2][1]);
    attemptsTodayById[aiId] = (attemptsTodayById[aiId] || 0) + 1;
  }

  // Build pending (waiting/approved) and active (in_exam/disqualified) lists,
  // DEDUPED per examinee. A soldier must appear ONCE in each list even when the
  // ממתינים sheet holds duplicate rows for them (re-registration after a stuck
  // row, mid-incident states). Without this the examiner saw the same person
  // two/three times in "במבחן כרגע" (reported incident @ בוחן יניר). Dedup keys
  // on normalized id within this session. Collapse priority:
  //   • pending: latest row wins.
  //   • active: a 'disqualified' (needs-decision) row beats an 'in_exam' row;
  //     within the same status, the latest row wins.
  var pendingById = {};
  var activeById = {};
  for (var i = 1; i < pendData.length; i++) {
    if (String(pendData[i][0]) !== code) continue;
    var s = String(pendData[i][5] || '').trim();
    if (s !== 'waiting' && s !== 'approved' && s !== 'in_exam' && s !== 'disqualified') continue;
    var dqCount = (pendData[i].length > 13) ? (Number(pendData[i][13]) || 0) : 0;
    var hasExtScreen = (pendData[i].length > 14) ? (String(pendData[i][14] || '').trim() === 'כן') : false;
    var warnCount = (pendData[i].length > 15) ? (Number(pendData[i][15]) || 0) : 0;
    var lastWarn = (pendData[i].length > 16) ? String(pendData[i][16] || '') : '';
    var idNorm = normalizeId(pendData[i][1]);
    var item = { idNumber: pendData[i][1], name: pendData[i][2], phone: pendData[i][3], time: pendData[i][4], examStartTime: pendData[i][11] || '', status: s, language: pendData[i][6] || '', population: pendData[i][7] || '', site: (pendData[i].length > 17) ? (pendData[i][17] || '') : '', license: pendData[i][8] || '', audioMode: pendData[i][9] || 'off', timeExtension: String(pendData[i][10] || ''), dqCount: dqCount, warnings: warnCount, lastWarning: lastWarn, attemptsToday: attemptsTodayById[idNorm] || 0, hasExtendedScreen: hasExtScreen, extraMinutes: extraMinById[idNorm] || 0, finishedOnDevice: (pendData[i].length > 18 ? !!pendData[i][18] : false) };
    if (s === 'waiting' || s === 'approved') {
      pendingById[idNorm] = item; // ascending loop → latest row wins
    } else {
      // Surface DQ events the examiner must decide on (in_exam + disqualified).
      if (s === 'disqualified') item.dqPending = true;
      var prevA = activeById[idNorm];
      if (!prevA) {
        activeById[idNorm] = item;
      } else {
        // 'disqualified' (needs decision) beats 'in_exam'; same status → latest wins.
        var curDQ = (s === 'disqualified');
        var prevDQ = (prevA.status === 'disqualified');
        if (curDQ || !prevDQ) activeById[idNorm] = item;
      }
    }
  }
  for (var pkA in pendingById) pending.push(pendingById[pkA]);
  for (var akA in activeById) active.push(activeById[akA]);

  // Re-read resData in case cleanup added new results.
  // NOTE (r15): the only write to resSheet above is inside the rare timeout-fail
  // branch, which re-reads by itself — so on a normal poll this second tail read
  // re-fetches identical data. It is kept for now because it ALSO picks up a
  // result another execution appended mid-request, which is exactly the latency
  // the 2s fast-sync was built to remove. Measure it before trading that away.
  diagMark('sheet:results-dash-2');
  resData = readTail(resSheet, 0).rows;
  // DEDUP results per examinee: the תוצאות sheet can end up with several
  // non-בוטל rows for one (session, id) when recovery paths (timeout-fail,
  // manual force-complete, disqualify) appended rows that weren't superseded.
  // The examiner must see each soldier ONCE — keep only the LATEST row, which
  // matches the system's own canonical rule (every supersede appends the newest
  // and marks older ones בוטל; latest-wins is the safety net when that didn't run).
  var latestResRowById = {};
  for (var jd = 1; jd < resData.length; jd++) {
    if (String(resData[jd][13]) !== code) continue;
    if (String(resData[jd][7] || '') === 'בוטל') continue;
    latestResRowById[normalizeId(resData[jd][1])] = jd; // ascending → ends as latest
  }
  var completed = [];
  for (var j = 1; j < resData.length; j++) {
    if (String(resData[j][13]) !== code) continue;
    if (String(resData[j][7] || '') === 'בוטל') continue;
    if (latestResRowById[normalizeId(resData[j][1])] !== j) continue; // keep latest only
    completed.push({
      date: resData[j][0],
      idNumber: resData[j][1],
      name: resData[j][2],
      phone: resData[j][3],
      license: resData[j][4],
      score: resData[j][5],
      percent: resData[j][6],
      passed: resData[j][7],
      time: resData[j][8],
      examiner: resData[j][9],
      site: resData[j][10],
      classroom: resData[j][11],
      language: resData[j][12],
      attempt: resData[j][14],
      wrongDetails: resData[j][15],
      sent: resData[j][16],
      disqualified: resData[j][17],
      waLink: resData[j][18],
      population: resData[j][19] || '',
      corrected: resData[j][20] || false,
      audioMode: resData[j][21] || 'off',
      // Integrity flags the server already computes & stores but the dashboard
      // never showed: verified='מאומת' when the score was re-computed against the
      // trusted answer key; suspicious='חשוד' when the exam took <3 min. Surfacing
      // these lets the examiner spot any result that was NOT server-verified
      // (missing answer key, missing exam-registration, or a tampered/forged
      // submit) instead of it looking identical to a clean pass.
      verified: (resData[j].length > 22) ? (resData[j][22] || '') : '',
      suspicious: (resData[j].length > 23) ? (resData[j][23] || '') : '',
      device: (resData[j].length > 29) ? (resData[j][29] || '') : ''
    });
  }

  // Flag repeat examinees: check if any pending examinee already tested today (any session)
  var now = new Date();
  var todayDD = ('0' + now.getDate()).slice(-2);
  var todayMM = ('0' + (now.getMonth() + 1)).slice(-2);
  var todayYYYY = now.getFullYear();
  var todayDate = todayDD + '/' + todayMM + '/' + todayYYYY; // "DD/MM/YYYY"
  // Index today's non-בוטל results by examinee id (any session), then attach —
  // was a full scan of תוצאות per pending examinee.
  var todayExamsById = {};
  for (var ti = 1; ti < resData.length; ti++) {
    if (String(resData[ti][7] || '') === 'בוטל') continue;
    // Handle both Date objects and string dates from Sheets
    var _cd = resData[ti][0];
    var _ds = '';
    if (_cd instanceof Date) {
      _ds = ('0' + _cd.getDate()).slice(-2) + '/' + ('0' + (_cd.getMonth() + 1)).slice(-2) + '/' + _cd.getFullYear();
    } else {
      _ds = String(_cd);
    }
    if (_ds.indexOf(todayDate) !== 0) continue;
    var _tk = normalizeId(resData[ti][1]);
    (todayExamsById[_tk] = todayExamsById[_tk] || []).push({ license: String(resData[ti][4]), score: String(resData[ti][5]), passed: String(resData[ti][7]), language: String(resData[ti][12] || '') });
  }
  for (var pi = 0; pi < pending.length; pi++) {
    var _te = todayExamsById[normalizeId(pending[pi].idNumber)];
    if (_te && _te.length > 0) pending[pi].todayExams = _te;
  }

  // Cross-reference registration times from ממתינים for completed results (indexed
  // — was a reverse scan of ממתינים per completed examinee). Last matching row for
  // (code, id) wins, exactly as the reverse-from-end + break did.
  var pendRegTimeById = {};
  for (var pr = 1; pr < pendData.length; pr++) {
    if (String(pendData[pr][0]) !== code) continue;
    pendRegTimeById[normalizeId(pendData[pr][1])] = pendData[pr][4];
  }
  for (var c = 0; c < completed.length; c++) {
    var _rk = normalizeId(completed[c].idNumber);
    if (Object.prototype.hasOwnProperty.call(pendRegTimeById, _rk)) {
      completed[c].registrationTime = pendRegTimeById[_rk];
    }
  }

  diagMark('compute:dash-done');
  return jsonResponse({ status: 'ok', pending: pending, active: active, completed: completed });
}

