// ---- The board's own rules over 'תוצאות', as functions (r32, 22/09/2026) ----
// Until r32 the examiner board was the only reader of these three rules, so
// they lived inside handleExaminerDashboard as loops. DESIGN §14.1 makes the
// board draw itself from sessionSnapshot v2 (one Google round trip per change
// instead of two — KNOWN_ISSUES #35), which means handleSessionSnapshot has to
// produce the SAME lists from the same sheet. Two copies of a dedup rule is how
// the two screens drift apart, so the rules live here once and both handlers
// call them. Behaviour is unchanged, line for line — including the quirks: the
// attempts tally parses the date cell with the Date constructor while the
// today-exams list compares the formatted DD/MM/YYYY prefix.
// Both callers are `exam` modules (server/BUILD_TARGETS.json), so the reports
// deployment carries neither.

// "Already tested today" tally: non-בוטל result rows written today, by
// examinee id, in ANY session — someone who tried earlier today elsewhere must
// still raise the flag.
function attemptsTodayFromResults(resData) {
  var now = new Date();
  var todayDateStr = now.getFullYear() + '-' + (now.getMonth() + 1) + '-' + now.getDate();
  function isToday(val) {
    if (!val) return false;
    try {
      var d = (val instanceof Date) ? val : new Date(val);
      if (isNaN(d.getTime())) return false;
      return (d.getFullYear() + '-' + (d.getMonth() + 1) + '-' + d.getDate()) === todayDateStr;
    } catch(_) { return false; }
  }
  var byId = {};
  for (var i = 1; i < resData.length; i++) {
    if (!isToday(resData[i][0])) continue;
    if (String(resData[i][7] || '').trim() === 'בוטל') continue;   // overturned, not a real attempt
    var k = normalizeId(resData[i][1]);
    byId[k] = (byId[k] || 0) + 1;
  }
  return byId;
}

// Today's non-בוטל results by examinee id, any session — what the board shows
// under a repeat examinee. The date cell is either a Date or the DD/MM/YYYY
// string todayStr() writes, so the match is on the formatted prefix.
function todayExamsFromResults(resData) {
  var now = new Date();
  var todayDate = ('0' + now.getDate()).slice(-2) + '/' + ('0' + (now.getMonth() + 1)).slice(-2) + '/' + now.getFullYear();
  var byId = {};
  for (var ti = 1; ti < resData.length; ti++) {
    if (String(resData[ti][7] || '') === 'בוטל') continue;
    var _cd = resData[ti][0], _ds = '';
    if (_cd instanceof Date) {
      _ds = ('0' + _cd.getDate()).slice(-2) + '/' + ('0' + (_cd.getMonth() + 1)).slice(-2) + '/' + _cd.getFullYear();
    } else {
      _ds = String(_cd);
    }
    if (_ds.indexOf(todayDate) !== 0) continue;
    var _tk = normalizeId(resData[ti][1]);
    (byId[_tk] = byId[_tk] || []).push({ license: String(resData[ti][4]), score: String(resData[ti][5]), passed: String(resData[ti][7]), language: String(resData[ti][12] || '') });
  }
  return byId;
}

// The 'completed' list of one session, in sheet order. DEDUP per examinee: the
// תוצאות sheet can end up with several non-בוטל rows for one (session, id) when
// recovery paths (timeout-fail, manual force-complete, disqualify) appended
// rows that weren't superseded. The examiner must see each soldier ONCE — keep
// only the LATEST row, which matches the system's own canonical rule (every
// supersede appends the newest and marks older ones בוטל; latest-wins is the
// safety net when that didn't run).
// `registrationTime` is NOT added here: it comes from ממתינים and only the
// board has that table (the snapshot's reader computes it from its own rows).
function completedResultsForSession(resData, code) {
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
  return completed;
}

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
  var _resT = readTail(resSheet, 0);
  var resData = _resT.rows;
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

  // ---- Auto-cleanup of stale in_exam/approved rows ---------------------------
  // Bounded on purpose (review C R1 / fix A1). This loop used to do, INSIDE
  // itself and per stale row: a full 'סשנים' read, a full 'תוצאות' read for the
  // attempt number, an append, another results tail read and an index rebuild —
  // 6 round trips and ~169,000 cells each. It fires hardest right after an
  // outage, when every examinee who could not submit is stale at once: 40 of
  // them measured 247 round trips / 6.86 M cells in ONE 2-5 s poll, which is the
  // best match at HEAD for the end-of-exam 360 s doGet kills.
  // Now: at most DASH_MAX_RECONCILE_PER_POLL rows per request (the rest are
  // reconciled by the next poll, 2-5 s later), the session row is read at most
  // once, the attempt history at most once, and nothing is re-read after an
  // append — the appended row is added to the in-memory table instead.
  var DASH_MAX_RECONCILE_PER_POLL = 3;
  var reconciled = 0;
  var sessionRowForDash = null, sessionRowRead = false;
  function dashSessionRow() {
    if (!sessionRowRead) { sessionRowRead = true; diagMark('sheet:sessions-dash'); sessionRowForDash = sessionRowByCode(code); }
    return sessionRowForDash;
  }
  var attemptHistory = null;
  function dashAttemptCount(idNumber, license) {
    if (!attemptHistory) { diagMark('sheet:attempts-dash'); attemptHistory = readAttemptHistory(); }
    return countAttemptRows(attemptHistory, normalizeId(idNumber), String(license));
  }
  var now = new Date();
  var BASE_EXAM_MS = 40 * 60 * 1000;
  var STALE_BUFFER_MS = 20 * 60 * 1000; // 20 minutes buffer (approval wait + instructions)
  for (var ci = 1; ci < pendData.length; ci++) {
    if (reconciled >= DASH_MAX_RECONCILE_PER_POLL) break;
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
      reconciled++;
      // Fix dangling status — mark as completed (single status writer: the
      // snapshot the examinee's poller reads is dropped in the same call).
      setPendingStatus(pendSheet, ci + 1 + pendOff, code, 'completed');
      pendData[ci][5] = 'completed'; // update local copy
      // Keep pendTermBySessId in sync: this row was in_exam/approved (loop guard
      // above) → now a 'completed' terminal, so a fresh rescan would count it here.
      var _mk = normalizeId(ciId);
      if (!pendTermBySessId[_mk]) pendTermBySessId[_mk] = { dqTerminals: 0, otherTerminals: 0 };
      pendTermBySessId[_mk].otherTerminals++;
      if (effectiveStale && !hasUnmatchedResult) {
        // Create a timeout fail result
        var ses2 = dashSessionRow();
        var license2 = pendData[ci][8] || '', site2 = '', classroom2 = '', examinerName2 = '', language2 = pendData[ci][6] || 'he';
        if (ses2) {
          examinerName2 = ses2[2] || '';
          site2 = ses2[3] || '';
          classroom2 = ses2[4] || '';
          if (!license2) license2 = ses2[5] || '';
        }
        var failRow = [
          todayStr(), ciId, pendData[ci][2] || '', pendData[ci][3] || '', license2,
          '0/30', '0%', 'נכשל', '', examinerName2,
          site2, classroom2, language2, code,
          dashAttemptCount(ciId, license2) + 1, 'ניתוק/טיימאאוט — הנבחן לא סיים את המבחן', false, false, '',
          pendData[ci][7] || '', false, pendData[ci][9] || 'off'
        ];
        resSheet.appendRow(failRow);
        // The row we just wrote is the only thing a re-read would have added, so
        // add it in memory: later iterations, the completed list and the
        // attempts-today tally all see it without another read of 'תוצאות'.
        resData.push(failRow);
        if (attemptHistory) attemptHistory.push([failRow[0], failRow[1], '', '', failRow[4], '', '', failRow[7]]);
        if (!resBySessId[_mk]) resBySessId[_mk] = { dqResults: 0, otherResults: 0 };
        resBySessId[_mk].otherResults++;
      }
    }
  }

  // Pre-compute attempts-today by examinee id (for "second attempt today"
  // warning) — the rule itself is above, shared with sessionSnapshot v2.
  var attemptsTodayById = attemptsTodayFromResults(resData);

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

  // r25: the second tail read of 'תוצאות' is GONE. On a normal poll it
  // re-fetched identical data (1,000 × 30 cells, 12 times a minute per
  // examiner); its only other effect was picking up a result another execution
  // appended during this request, and that arrives one poll later anyway — the
  // dashboard polls every 2 s while a result is syncing.
  // The dedup-per-examinee rule is above, shared with sessionSnapshot v2.
  var completed = completedResultsForSession(resData, code);

  // Flag repeat examinees: check if any pending examinee already tested today
  // (any session) — the rule itself is above, shared with sessionSnapshot v2.
  var todayExamsById = todayExamsFromResults(resData);
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

