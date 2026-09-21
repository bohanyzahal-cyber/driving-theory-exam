// Returns every session that took place at the same site on the same day as
// the caller's sessionCode, plus every result row belonging to those sessions.
// Powers the "דו"ח משותף לאתר" (combined site report) button on the examiner
// dashboard — when two examiners share a site, the Rav-Bochen wants one
// report with both their sessions side by side.
//
// Authorization: the caller must be the responsible examiner of at least one
// of the sessions, OR hold a commander role. Regular examiners who happen to
// have a session at the same site can still see results for their own session
// via the existing examinerDashboard — the combined view is a chain-of-
// command audit lens.
function handleSiteCombinedReport(p) {
  if (!verifyToken(p.examinerId, p.token)) {
    return jsonResponse({ status: 'error', message: 'טוקן לא תקין', tokenExpired: true });
  }

  diagMark('sheet:sessions-report');
  var sessData = sessionRows();

  // Locate the calling session to discover its site + date.
  var anchorSite = '';
  var anchorDate = null;
  var callerSessionCode = String(p.sessionCode || '').trim();
  for (var i = 1; i < sessData.length; i++) {
    if (String(sessData[i][0]).trim() === callerSessionCode) {
      anchorSite = String(sessData[i][3] || '').trim();
      anchorDate = sessData[i][8] instanceof Date ? sessData[i][8] : new Date(sessData[i][8]);
      break;
    }
  }
  if (!anchorSite || !anchorDate || isNaN(anchorDate.getTime())) {
    return jsonResponse({ status: 'error', message: 'סשן לא נמצא או חסר תאריך' });
  }

  var dayStart = new Date(anchorDate);
  dayStart.setHours(0, 0, 0, 0);
  var dayEnd = new Date(dayStart);
  dayEnd.setDate(dayEnd.getDate() + 1);

  // Pull every session at the same (site, day). We accept sessions in any
  // status (active / closed / expired) — the combined report is a historical
  // record, not a live view.
  var sessions = [];
  var callerExaminerName = '';
  var callerIsResponsibleOnAny = false;
  for (var j = 1; j < sessData.length; j++) {
    var site = String(sessData[j][3] || '').trim();
    if (site !== anchorSite) continue;
    var created = sessData[j][8] instanceof Date ? sessData[j][8] : new Date(sessData[j][8]);
    if (isNaN(created.getTime()) || created < dayStart || created >= dayEnd) continue;
    var responsible = String((sessData[j].length > 13 ? sessData[j][13] : '') || '').trim();
    var examinerName = String(sessData[j][2] || '').trim();
    if (normalizeId(sessData[j][1]) === normalizeId(p.examinerId)) {
      callerExaminerName = examinerName;
      if (responsible && responsible === examinerName) callerIsResponsibleOnAny = true;
    }
    sessions.push({
      code: String(sessData[j][0]).trim(),
      examinerId: String(sessData[j][1] || '').trim(),
      examinerName: examinerName,
      site: site,
      classroom: String(sessData[j][4] || '').trim(),
      license: String(sessData[j][5] || '').trim(),
      language: String(sessData[j][6] || '').trim(),
      audioMode: String(sessData[j][7] || '').trim(),
      created: created.toISOString(),
      responsibleExaminer: responsible,
      quotas: decodeSessionQuotas(sessData[j][11], sessData[j][12], sessData[j][5])
    });
  }
  if (sessions.length === 0) {
    return jsonResponse({ status: 'error', message: 'לא נמצאו סשנים תואמים' });
  }

  // Authorization: pass if caller is responsible-on-some-session OR commander.
  diagMark('sheet:role-report');
  var role = getExaminerRole(p.examinerId);
  var isCommander = (
    role === 'מפקד' || role === 'מפקד מקומי' || role === 'מפקד ראשי' ||
    role === 'מפקד מרכז' || isKdtzRole(role) || role === 'רב בוחן'
  );
  // Also: caller is the responsible examiner of any session in the set,
  // even if their own session doesn't list themselves as responsible.
  var callerNamedAsResponsibleAnywhere = false;
  if (callerExaminerName) {
    for (var s = 0; s < sessions.length; s++) {
      if (sessions[s].responsibleExaminer === callerExaminerName) {
        callerNamedAsResponsibleAnywhere = true;
        break;
      }
    }
  }
  if (!isCommander && !callerIsResponsibleOnAny && !callerNamedAsResponsibleAnywhere) {
    return jsonResponse({
      status: 'error',
      message: 'הדו"ח המשותף זמין רק לבוחן האחראי או למפקד'
    });
  }

  // Now pull all results that belong to any of these session codes.
  var sessionCodesSet = {};
  for (var sc = 0; sc < sessions.length; sc++) sessionCodesSet[sessions[sc].code] = true;

  // r11 read the tail and fell back to the WHOLE sheet whenever the reported
  // day was older than the tail — the very first line the 'אבחון' sheet ever
  // recorded was this report at 81.5 s (2026-09-15), and the 360 s doGet kills
  // cluster at end-of-exam report time. r25: one date-bounded read that spans
  // the live sheet and the archive, so an old day costs the rows of that day
  // instead of every result ever recorded.
  diagMark('sheet:results-report');
  var resRead = readResultsSince(dayStart);
  var resData = resRead.rows;
  diagMark('sheet:results-report-done:' + resRead.mode);
  var results = [];
  for (var r = 1; r < resData.length; r++) {
    var sCode = String(resData[r][13] || '').trim();
    if (!sessionCodesSet[sCode]) continue;
    if (String(resData[r][7] || '').trim() === 'בוטל') continue; // skip overturned/superseded rows (consistent with the other report handlers)
    // parseSheetDateTime, not new Date(): column A is written as the STRING
    // "DD/MM/YYYY HH:MM" and only becomes a real date when the spreadsheet's
    // locale parses it. Where it does not (a text-formatted column, an imported
    // row), new Date() returned Invalid Date and toISOString() threw — taking
    // the whole report down with a 500 instead of one odd row.
    var rDate = parseSheetDateTime(resData[r][0]);
    results.push({
      date: rDate ? rDate.toISOString() : String(resData[r][0] || ''),
      idNumber: String(resData[r][1] || ''),
      name: String(resData[r][2] || ''),
      phone: String(resData[r][3] || ''),
      license: String(resData[r][4] || ''),
      score: resData[r][5],
      percent: resData[r][6],
      passed: String(resData[r][7] || ''),
      time: String(resData[r][8] || ''),
      examiner: String(resData[r][9] || ''),
      site: String(resData[r][10] || ''),
      classroom: String(resData[r][11] || ''),
      language: String(resData[r][12] || ''),
      sessionCode: sCode,
      attemptNum: resData[r][14],
      population: String(resData[r][19] || ''),
      disqualified: resData[r][17] === true || String(resData[r][17]).toUpperCase() === 'TRUE',
      audioMode: String(resData[r][21] || ''),
      verified: (resData[r].length > 22) ? String(resData[r][22] || '') : '',
      device: (resData[r].length > 29) ? String(resData[r][29] || '') : ''
    });
  }

  return jsonResponse({
    status: 'ok',
    site: anchorSite,
    dayStart: dayStart.toISOString(),
    sessions: sessions,
    results: results
  });
}

