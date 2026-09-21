// ========== מערכת מורים — Teacher System ==========

function verifyTeacherToken(teacherId, token) {
  if (!teacherId || !token) return false;
  var sheet = getSheet('מורים');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (normalizeId(data[i][1]) === normalizeId(teacherId)) {
      var storedToken = String(data[i][4] || '');
      var expiry = data[i][5];
      if (storedToken === token && expiry) {
        var expiryDate = expiry instanceof Date ? expiry : new Date(expiry);
        if (new Date() <= expiryDate) return true;
      }
      // Continue searching other rows with same ID
    }
  }
  return false;
}

function requireTeacherToken(p) {
  var valid = verifyTeacherToken(p.teacherId, p.token);
  diagMark('auth:teacher');
  if (!valid) {
    return jsonResponse({ status: 'error', message: 'טוקן לא תקין — יש להתחבר מחדש', tokenExpired: true });
  }
  return null;
}

function generateClassCode() {
  var chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  var sheet = getSheet('כיתות');
  var data = sheet.getDataRange().getValues();
  var existing = {};
  for (var i = 1; i < data.length; i++) existing[String(data[i][0]).trim()] = true;
  var code;
  do {
    code = '';
    for (var c = 0; c < 6; c++) code += chars.charAt(Math.floor(Math.random() * chars.length));
  } while (existing[code]);
  return code;
}

// Read-only lookup of classes that were deleted (archived by handleTeacherDeleteClass).
// Uses getSheetByName (NOT getSheet) so a report read never auto-creates the sheet.
// Returns {} if the archive doesn't exist yet (no class has ever been deleted).
function getDeletedClassMap() {
  var map = {};
  try {
    var sheet = getSheetIfExists('כיתות שנמחקו');
    if (!sheet) return map;
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var cc = String(data[i][0] || '').trim();
      if (!cc) continue;
      map[cc] = {
        teacherName: String(data[i][3] || ''),
        className: String(data[i][1] || ''),
        license: String(data[i][4] || ''),
        site: String(data[i][5] || ''),
        deleted: true
      };
    }
  } catch (e) { /* archive missing/unreadable → treat as empty */ }
  return map;
}

// Resolve a practice-result class code to display info for the commander reports.
// Precedence: active class → deleted class (real teacher + "(כיתה שנמחקה)" tag) →
// truly unrecognized code ("קוד לא מזוהה"). This is what lets the dashboard tell a
// legitimate deleted class apart from a forged/never-existed code, instead of
// lumping both under "לא ידוע".
function resolveClassInfo(classCode, classMap, deletedMap) {
  if (classMap && classMap[classCode]) return classMap[classCode];
  if (deletedMap && deletedMap[classCode]) {
    var d = deletedMap[classCode];
    return {
      teacherName: d.teacherName || 'לא ידוע',
      teacherId: '',
      className: (d.className || classCode) + ' (כיתה שנמחקה)',
      license: d.license || '',
      site: d.site || '',
      deleted: true
    };
  }
  return { teacherName: 'קוד לא מזוהה', teacherId: '', className: classCode, license: '', site: '', unresolved: true };
}

function handleTeacherLogin(p) {
  var maintenance = practiceWriteGuard(); if (maintenance) return maintenance;   // r24
  var sheet = getSheet('מורים');
  var data = sheet.getDataRange().getValues();
  var matchedRows = [];
  for (var i = 1; i < data.length; i++) {
    if (normalizeId(data[i][1]) === normalizeId(p.idNumber)) {
      matchedRows.push(i);
    }
  }
  if (matchedRows.length === 0) {
    return jsonResponse({ status: 'error', message: 'מורה לא נמצא' });
  }
  var lastError = '';
  for (var m = 0; m < matchedRows.length; m++) {
    var i = matchedRows[m];
    var row = i + 1;
    var failedAttempts = Number(data[i][6]) || 0;
    var lockoutUntil = data[i][7];
    if (lockoutUntil) {
      var lockoutDate = lockoutUntil instanceof Date ? lockoutUntil : new Date(lockoutUntil);
      if (new Date() < lockoutDate) {
        var minsLeft = Math.ceil((lockoutDate - new Date()) / 60000);
        lastError = 'החשבון נעול. נסה שוב בעוד ' + minsLeft + ' דקות';
        continue;
      }
      failedAttempts = 0;
      sheet.getRange(row, 7).setValue(0);
      sheet.getRange(row, 8).setValue('');
    }
    if (String(data[i][2]) === String(p.password)) {
      if (data[i][3] === 'כן' || data[i][3] === true || data[i][3] === 'TRUE') {
        if (failedAttempts > 0) {
          sheet.getRange(row, 7).setValue(0);
          sheet.getRange(row, 8).setValue('');
        }
        var token = generateToken();
        var expiry = new Date();
        expiry.setHours(expiry.getHours() + 12);
        sheet.getRange(row, 5).setValue(token);
        sheet.getRange(row, 6).setValue(expiry);
        return jsonResponse({ status: 'ok', teacher: { name: data[i][0], id: normalizeId(data[i][1]), token: token, role: String(data[i][8] || 'מורה'), site: String(data[i][9] || '') } });
      } else {
        lastError = 'החשבון אינו פעיל';
        continue;
      }
    } else {
      lastError = 'סיסמה שגויה';
    }
  }
  // If no row matched successfully, increment failed attempts on first active row
  if (lastError === 'סיסמה שגויה' && matchedRows.length > 0) {
    var fi = matchedRows[0];
    var fRow = fi + 1;
    var fa = (Number(data[fi][6]) || 0) + 1;
    sheet.getRange(fRow, 7).setValue(fa);
    if (fa >= 5) {
      var lockout = new Date();
      lockout.setMinutes(lockout.getMinutes() + 15);
      sheet.getRange(fRow, 8).setValue(lockout);
      return jsonResponse({ status: 'error', message: 'יותר מדי ניסיונות. החשבון ננעל ל-15 דקות' });
    }
  }
  return jsonResponse({ status: 'error', message: lastError || 'שגיאה בהתחברות' });
}

function handleTeacherVerifyLogin(p) {
  if (!verifyTeacherToken(p.teacherId, p.token)) {
    return jsonResponse({ status: 'error', message: 'טוקן לא תקין', tokenExpired: true });
  }
  var sheet = getSheet('מורים');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (normalizeId(data[i][1]) === normalizeId(p.teacherId)) {
      return jsonResponse({ status: 'ok', teacher: { name: data[i][0], id: normalizeId(data[i][1]), role: String(data[i][8] || 'מורה'), site: String(data[i][9] || '') } });
    }
  }
  return jsonResponse({ status: 'error', message: 'מורה לא נמצא' });
}

function handleTeacherCommanderDashboard(p) {
  // Verify commander role
  var tSheet = getSheet('מורים');
  var tData = tSheet.getDataRange().getValues();
  var role = '';
  var userSite = '';
  for (var i = 1; i < tData.length; i++) {
    if (normalizeId(tData[i][1]) === normalizeId(p.teacherId)) {
      role = String(tData[i][8] || 'מורה');
      userSite = String(tData[i][9] || '');
      break;
    }
  }
  if (role !== 'מפקד' && role !== 'מפקד מקומי' && role !== 'מפקד ראשי' && !isKdtzRole(role)) {
    return jsonResponse({ status: 'error', message: 'אין הרשאת מפקד' });
  }

  // Determine commander scope:
  //   isGlobal    — sees every site (no site filter)
  //   isLocal     — single site (column 9 holds the one site name)
  //   isMultiSite — fixed list of sites (column 9 holds a comma-separated list)
  // Legacy: role === 'מפקד' treated as 'מפקד ראשי'
  var isGlobal = (role === 'מפקד ראשי' || role === 'מפקד');
  var isLocal = (role === 'מפקד מקומי');
  var isMultiSite = isKdtzRole(role);
  var managedSites = [];
  if (isMultiSite) {
    managedSites = String(userSite || '').split(',').map(function(s) { return s.trim(); }).filter(function(s) { return s; });
    if (!managedSites.length) {
      return jsonResponse({ status: 'error', message: 'לא הוקצו אתרים — מלא רשימה מופרדת בפסיקים בעמודת האתר במורים' });
    }
  }

  // Parse date range
  var dateFrom = parseDateParam(p.dateFrom);
  var dateTo = parseDateParam(p.dateTo);
  if (!dateFrom || !dateTo) {
    return jsonResponse({ status: 'error', message: 'תאריכים לא תקינים' });
  }
  dateTo.setHours(23, 59, 59, 999);

  // Build class→teacher map from כיתות sheet
  var classSheet = getSheet('כיתות');
  var classData = classSheet.getDataRange().getValues();
  var classMap = {}; // classCode → { teacherName, className, license, site }
  for (var c = 1; c < classData.length; c++) {
    var cc = String(classData[c][0]).trim();
    classMap[cc] = {
      teacherName: String(classData[c][3] || ''),
      teacherId: normalizeId(classData[c][2]),
      className: String(classData[c][1] || ''),
      license: String(classData[c][4] || ''),
      site: String(classData[c][7] || '')
    };
  }
  var deletedClassMap = getDeletedClassMap();

  // Read practice results
  var resSheet = getSheet('תוצאות תרגול');
  var resData = resSheet.getDataRange().getValues();

  var overall = { total: 0, passed: 0, failed: 0, scores: [], stayTimes: [], students: {}, teachers: {}, classes: {}, sites: {} };
  var byTeacher = {};
  var byClass = {};
  var byLicense = {};
  var byMode = {};
  var bySite = {};

  // Activity-by-hour (0–23) + most-failed-questions, mirroring the examiner
  // commander view. Practice rows store the submit time inside the date cell
  // ("DD/MM/YYYY HH:mm" — see todayStr) and the missed questions as a JSON
  // array in col 13 (פירוט שגויות), one {qNum, category, qText} per question.
  var hourBuckets = [];
  for (var hb = 0; hb < 24; hb++) hourBuckets.push(0);
  var wrongCounts = {}; // qText → { count, category }

  for (var r = 1; r < resData.length; r++) {
    var rowDate = parseSheetDate(resData[r][0]);
    if (!rowDate || rowDate < dateFrom || rowDate > dateTo) continue;

    var classCode = String(resData[r][3] || '').trim();
    if (!classCode) continue;

    var cInfo = resolveClassInfo(classCode, classMap, deletedClassMap);
    var classSite = cInfo.site || '';

    // Site filtering for local + multi-site commanders
    if (isLocal && userSite && classSite !== userSite) continue;
    if (isMultiSite && managedSites.indexOf(classSite) === -1) continue;

    var teacherName = cInfo.teacherName || 'לא ידוע';
    var className = cInfo.className || classCode;
    var license = String(resData[r][5] || cInfo.license || 'לא צוין');
    var mode = String(resData[r][4] || 'לא צוין');
    var studentId = String(resData[r][1] || '');
    var passedStr = String(resData[r][9] || '');
    var isPassed = (passedStr === 'עבר' || passedStr === 'true' || passedStr === true);
    var isFailed = (passedStr === 'נכשל' || passedStr === 'false' || passedStr === false);

    var pctVal = 0;
    var pctRaw = resData[r][8];
    if (typeof pctRaw === 'string' && pctRaw.indexOf('%') !== -1) {
      pctVal = parseFloat(pctRaw.replace('%', '')) || 0;
    } else {
      var pctNum = Number(pctRaw);
      if (!isNaN(pctNum)) {
        pctVal = pctNum <= 1 ? pctNum * 100 : pctNum;
      }
    }

    overall.total++;
    if (isPassed) overall.passed++;
    else if (isFailed) overall.failed++;
    overall.scores.push(pctVal);
    overall.students[studentId] = true;
    overall.teachers[teacherName] = true;
    overall.classes[classCode] = true;
    if (classSite) overall.sites[classSite] = true;

    // Stay-time (col 10, "M:SS"/"MM:SS") → seconds, for the avg/median time KPIs.
    var tSec = parsePracticeTimeSec(resData[r][10]);
    if (tSec > 0) overall.stayTimes.push(tSec);

    // Activity-by-hour from the submit timestamp embedded in the date cell.
    var hh = practiceRowHour(resData[r][0], rowDate);
    if (hh >= 0 && hh < 24) hourBuckets[hh]++;

    // Most-failed questions — the practice client sends a JSON array of
    // {qNum, category, qText}. Aggregate by question text (no ID/correct-answer
    // is captured in practice mode, unlike the real-exam path).
    var wdRaw = resData[r][13];
    if (wdRaw) {
      var wdArr = null;
      try { wdArr = (typeof wdRaw === 'string') ? JSON.parse(wdRaw) : wdRaw; } catch (eWD) { wdArr = null; }
      if (Array.isArray(wdArr)) {
        for (var wdi = 0; wdi < wdArr.length; wdi++) {
          var wit = wdArr[wdi];
          if (!wit) continue;
          var qt = String(wit.qText || wit.question || '').trim();
          if (!qt) continue;
          if (qt.length > 200) qt = qt.substring(0, 200);
          if (!wrongCounts[qt]) wrongCounts[qt] = { count: 0, category: String(wit.category || '') };
          wrongCounts[qt].count++;
        }
      }
    }

    addToGroup(byTeacher, teacherName, isPassed, isFailed, pctVal, studentId);
    addToGroup(byClass, className + ' (' + classCode + ')', isPassed, isFailed, pctVal, studentId);
    addToGroup(byLicense, license, isPassed, isFailed, pctVal, studentId);
    addToGroup(byMode, mode, isPassed, isFailed, pctVal, studentId);

    // bySite aggregation — global and multi-site commanders both want
    // a per-site breakdown. Local commander has only one site, so the
    // tab is hidden client-side; no aggregation needed.
    if ((isGlobal || isMultiSite) && classSite) {
      addToGroup(bySite, classSite, isPassed, isFailed, pctVal, studentId);
      addToSubGroup(bySite, classSite, 'byTeacher', teacherName, isPassed, isFailed, pctVal, studentId);
      addToSubGroup(bySite, classSite, 'byClass', className + ' (' + classCode + ')', isPassed, isFailed, pctVal, studentId);
    }

    // Cross-tabulation sub-groups
    addToSubGroup(byTeacher, teacherName, 'byClass', className + ' (' + classCode + ')', isPassed, isFailed, pctVal, studentId);
    addToSubGroup(byTeacher, teacherName, 'byLicense', license, isPassed, isFailed, pctVal, studentId);
    addToSubGroup(byClass, className + ' (' + classCode + ')', 'byLicense', license, isPassed, isFailed, pctVal, studentId);
    addToSubGroup(byClass, className + ' (' + classCode + ')', 'byMode', mode, isPassed, isFailed, pctVal, studentId);
    addToSubGroup(byLicense, license, 'byTeacher', teacherName, isPassed, isFailed, pctVal, studentId);
    addToSubGroup(byLicense, license, 'byClass', className + ' (' + classCode + ')', isPassed, isFailed, pctVal, studentId);
    addToSubGroup(byMode, mode, 'byLicense', license, isPassed, isFailed, pctVal, studentId);
    addToSubGroup(byMode, mode, 'byTeacher', teacherName, isPassed, isFailed, pctVal, studentId);
  }

  function addToGroup(map, key, isPassed, isFailed, pctVal, studentId) {
    if (!map[key]) map[key] = { total: 0, passed: 0, failed: 0, scores: [], students: {} };
    map[key].total++;
    if (isPassed) map[key].passed++;
    else if (isFailed) map[key].failed++;
    map[key].scores.push(pctVal);
    map[key].students[studentId] = true;
  }

  function addToSubGroup(map, primaryKey, subDim, subKey, isPassed, isFailed, pctVal, studentId) {
    if (!map[primaryKey]) return;
    if (!map[primaryKey]._sub) map[primaryKey]._sub = {};
    if (!map[primaryKey]._sub[subDim]) map[primaryKey]._sub[subDim] = {};
    addToGroup(map[primaryKey]._sub[subDim], subKey, isPassed, isFailed, pctVal, studentId);
  }

  function computeStats(obj) {
    var avg = 0, median = 0;
    if (obj.scores.length > 0) {
      var sum = 0;
      for (var s = 0; s < obj.scores.length; s++) sum += obj.scores[s];
      avg = Math.round(sum / obj.scores.length);
      var sorted = obj.scores.slice().sort(function(a, b) { return a - b; });
      var mid = Math.floor(sorted.length / 2);
      median = sorted.length % 2 !== 0 ? sorted[mid] : Math.round((sorted[mid - 1] + sorted[mid]) / 2);
    }
    var passRate = obj.total > 0 ? Math.round((obj.passed / obj.total) * 100) : 0;
    var studentCount = Object.keys(obj.students || {}).length;
    return { total: obj.total, passed: obj.passed, failed: obj.failed, passRate: passRate, avgScore: avg, medianScore: median, students: studentCount };
  }

  function computeGroupWithSub(map) {
    var out = {};
    for (var key in map) {
      out[key] = computeStats(map[key]);
      if (map[key]._sub) {
        out[key].sub = {};
        for (var subDim in map[key]._sub) {
          out[key].sub[subDim] = {};
          for (var subKey in map[key]._sub[subDim]) {
            out[key].sub[subDim][subKey] = computeStats(map[key]._sub[subDim][subKey]);
          }
        }
      }
    }
    return out;
  }

  // Parse a practice duration into seconds. The student app sends "MM:SS"
  // (e.g. "5:03" = 5 min 3 sec), but Google Sheets AUTO-CONVERTS the string on
  // write, MISREADING "MM:SS" as "HH:MM". So getValues() never returns the
  // original string — it returns one of:
  //   • Date   — short sessions (<24 min). e.g. "5:03" → 05:03 time → Date.
  //   • number — long sessions (≥24 min). "24:29" can't be a time-of-day, so
  //              Sheets stores it as a DURATION serial (fraction of a day, e.g.
  //              ~1.02 for 24h29m). Verified against real data: 28,637 Date
  //              cells + 1,871 duration-number cells.
  //   • string — only if a value somehow wasn't auto-converted.
  // In every case the stored clock is H:M:S where the original minutes landed in
  // H and the original seconds in M. We recover by mapping H→minutes, M→seconds.
  // Cap at 2h to drop garbage (abandoned tabs produce multi-day serials).
  function parsePracticeTimeSec(v) {
    if (v === null || v === undefined || v === '') return 0;
    var mm, ss;
    if (v instanceof Date) {
      mm = v.getHours();      // original minutes (Sheets read them as hours)
      ss = v.getMinutes();    // original seconds (Sheets read them as minutes)
    } else if (typeof v === 'number') {
      // Day-fraction serial (works for both <1 time serials and ≥1 durations).
      if (v <= 0) return 0;
      var totalClockSec = Math.round(v * 86400); // the misread H:M:S, in seconds
      mm = Math.floor(totalClockSec / 3600);     // clock-hours → original minutes
      ss = Math.floor((totalClockSec % 3600) / 60); // clock-minutes → original seconds
    } else {
      var m = String(v).trim().match(/^(\d{1,3}):(\d{2})$/);
      if (!m) return 0;
      mm = parseInt(m[1], 10);
      ss = parseInt(m[2], 10);
    }
    if (isNaN(mm) || isNaN(ss) || ss >= 60) return 0;
    var t = mm * 60 + ss;
    return (t > 0 && t <= 7200) ? t : 0;
  }
  // Hour-of-day from the date cell. Sheets usually auto-parses "DD/MM/YYYY
  // HH:mm" into a real Date (hour preserved); for string cells we regex the
  // HH out, since parseSheetDate drops the time component.
  function practiceRowHour(cell, parsed) {
    if (cell instanceof Date) return cell.getHours();
    var s = String(cell || '');
    var m = s.match(/\d{1,2}\/\d{1,2}\/\d{4}\s+(\d{1,2}):(\d{2})/);
    if (m) return parseInt(m[1], 10);
    if (parsed && parsed instanceof Date) return parsed.getHours();
    return -1;
  }
  function avgMedianSec(arr) {
    if (!arr || !arr.length) return { avg: 0, median: 0 };
    var sum = 0;
    for (var i = 0; i < arr.length; i++) sum += arr[i];
    var avg = Math.round(sum / arr.length);
    var sorted = arr.slice().sort(function(a, b) { return a - b; });
    var mid = Math.floor(sorted.length / 2);
    var median = sorted.length % 2 !== 0 ? sorted[mid] : Math.round((sorted[mid - 1] + sorted[mid]) / 2);
    return { avg: avg, median: median };
  }

  var overallStats = computeStats(overall);
  overallStats.activeTeachers = Object.keys(overall.teachers).length;
  overallStats.activeClasses = Object.keys(overall.classes).length;
  overallStats.activeSites = Object.keys(overall.sites).length;
  var ovTime = avgMedianSec(overall.stayTimes);
  overallStats.stayAvg = ovTime.avg;
  overallStats.stayMedian = ovTime.median;
  overallStats.stayCount = overall.stayTimes.length;

  var result = {
    overall: overallStats,
    byTeacher: computeGroupWithSub(byTeacher),
    byClass: computeGroupWithSub(byClass),
    byLicense: computeGroupWithSub(byLicense),
    byMode: computeGroupWithSub(byMode),
    commanderRole: role,
    commanderSite: userSite
  };

  // bySite breakdown for any cross-site commander (global or multi-site).
  if (isGlobal || isMultiSite) {
    result.bySite = computeGroupWithSub(bySite);
  }

  // Active classes list
  var studSheet = getSheet('תלמידי כיתות');
  var studData = studSheet.getDataRange().getValues();
  var studCountMap = {};
  for (var sc = 1; sc < studData.length; sc++) {
    var scc = String(studData[sc][0]).trim();
    studCountMap[scc] = (studCountMap[scc] || 0) + 1;
  }
  var activeClasses = [];
  for (var ac = 1; ac < classData.length; ac++) {
    if (String(classData[ac][6]) !== 'כן') continue; // only active
    var acCode = String(classData[ac][0]).trim();
    var acSite = String(classData[ac][7] || '');
    if (isLocal && userSite && acSite !== userSite) continue;
    if (isMultiSite && managedSites.indexOf(acSite) === -1) continue;
    activeClasses.push({
      code: acCode,
      name: String(classData[ac][1] || ''),
      teacherName: String(classData[ac][3] || ''),
      license: String(classData[ac][4] || ''),
      site: acSite,
      students: studCountMap[acCode] || 0
    });
  }
  result.activeClasses = activeClasses;

  // Top-10 most-failed questions (sorted by fail count) + activity-by-hour.
  var topWrong = [];
  var wKeys = Object.keys(wrongCounts);
  wKeys.sort(function(a, b) { return wrongCounts[b].count - wrongCounts[a].count; });
  for (var twk = 0; twk < Math.min(wKeys.length, 10); twk++) {
    topWrong.push({ question: wKeys[twk], category: wrongCounts[wKeys[twk]].category, count: wrongCounts[wKeys[twk]].count });
  }
  result.topWrong = topWrong;
  result.hourly = hourBuckets;

  // Repeat REAL-exam failures — soldiers who failed the external theory exam
  // (תוצאות, not practice) 2+ times in the window. The training commander owns
  // the intervention: targeted practice before they burn another exam slot.
  // Site scoping mirrors the practice rows (local/multi-site commanders see
  // only their sites; exam rows carry the site in col 11).
  try {
    var examResSheet = getSheet('תוצאות');
    var examResData = examResSheet.getDataRange().getValues();
    var failsById = {};
    for (var er = 1; er < examResData.length; er++) {
      var erDate = parseSheetDate(examResData[er][0]);
      if (!erDate || erDate < dateFrom || erDate > dateTo) continue;
      // Only genuine knowledge fails: skip עבר, פסול (anti-cheat, not knowledge)
      // and בוטל (superseded/overturned rows).
      if (String(examResData[er][7] || '').trim() !== 'נכשל') continue;
      var erSite = String(examResData[er][10] || '');
      if (isLocal && userSite && erSite !== userSite) continue;
      if (isMultiSite && managedSites.indexOf(erSite) === -1) continue;
      var erId = normalizeId(examResData[er][1]);
      if (!erId) continue;
      if (!failsById[erId]) {
        failsById[erId] = { name: '', idLast4: String(examResData[er][1] || '').slice(-4), license: '', site: '', fails: 0, lastDate: '', lastScore: '' };
      }
      failsById[erId].fails++;
      // Rows are appended chronologically — the last in-range row wins the
      // "latest" fields.
      if (examResData[er][2]) failsById[erId].name = String(examResData[er][2]);
      if (examResData[er][4]) failsById[erId].license = String(examResData[er][4]);
      if (erSite) failsById[erId].site = erSite;
      failsById[erId].lastDate = erDate.getDate() + '/' + (erDate.getMonth() + 1) + '/' + erDate.getFullYear();
      failsById[erId].lastScore = String(examResData[er][5] || '');
    }
    var repeatFailures = [];
    for (var rfk in failsById) {
      if (failsById[rfk].fails >= 2) repeatFailures.push(failsById[rfk]);
    }
    repeatFailures.sort(function(a, b) { return b.fails - a.fails; });
    result.repeatFailures = repeatFailures.slice(0, 50);
  } catch (eRF) { result.repeatFailures = []; }

  return jsonResponse({ status: 'ok', data: result });
}

