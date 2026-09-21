function handleAdminDashboard(p) {
  // Verify admin role - check ALL rows for this ID
  var tSheet = getSheet('מורים');
  var tData = tSheet.getDataRange().getValues();
  var isAdmin = false;
  for (var i = 1; i < tData.length; i++) {
    if (normalizeId(tData[i][1]) === normalizeId(p.teacherId)) {
      if (String(tData[i][8] || '') === 'אדמין') { isAdmin = true; break; }
    }
  }
  if (!isAdmin) {
    return jsonResponse({ status: 'error', message: 'אין הרשאת אדמין' });
  }

  var dateFrom = parseDateParam(p.dateFrom);
  var dateTo = parseDateParam(p.dateTo);
  if (!dateFrom || !dateTo) {
    return jsonResponse({ status: 'error', message: 'תאריכים לא תקינים' });
  }
  dateTo.setHours(23, 59, 59, 999);

  // Build class map
  var classSheet = getSheet('כיתות');
  var classData = classSheet.getDataRange().getValues();
  var classMap = {};
  for (var c = 1; c < classData.length; c++) {
    var cc = String(classData[c][0]).trim();
    classMap[cc] = {
      teacherName: String(classData[c][3] || ''),
      className: String(classData[c][1] || ''),
      license: String(classData[c][4] || ''),
      site: String(classData[c][7] || '')
    };
  }
  var deletedClassMap = getDeletedClassMap();

  // Read practice results - INCLUDING rows without classCode
  var resSheet = getSheet('תוצאות תרגול');
  var resData = resSheet.getDataRange().getValues();

  var overall = { total: 0, passed: 0, failed: 0, scores: [], students: {}, classes: {}, independentStudents: {} };
  var byLicense = {};
  var byMode = {};
  var byEnrollment = {};
  var byClass = {};
  var byDay = {};

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

  function fmtDay(d) {
    var dd = d.getDate(), mm = d.getMonth() + 1, yyyy = d.getFullYear();
    return (dd < 10 ? '0' : '') + dd + '/' + (mm < 10 ? '0' : '') + mm + '/' + yyyy;
  }

  for (var r = 1; r < resData.length; r++) {
    var rowDate = parseSheetDate(resData[r][0]);
    if (!rowDate || rowDate < dateFrom || rowDate > dateTo) continue;

    var classCode = String(resData[r][3] || '').trim();
    var isIndependent = !classCode;
    var enrollmentStatus = isIndependent ? 'עצמאי' : 'כיתה';
    var cInfo = classCode ? resolveClassInfo(classCode, classMap, deletedClassMap) : null;

    var license = String(resData[r][5] || (cInfo ? cInfo.license : '') || 'לא צוין');
    var mode = String(resData[r][4] || 'לא צוין');
    var studentId = String(resData[r][1] || '');
    var studentName = String(resData[r][2] || '');
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
    if (classCode) overall.classes[classCode] = true;
    if (isIndependent) overall.independentStudents[studentId] = true;

    // By license
    addToGroup(byLicense, license, isPassed, isFailed, pctVal, studentId);
    addToSubGroup(byLicense, license, 'byMode', mode, isPassed, isFailed, pctVal, studentId);
    addToSubGroup(byLicense, license, 'byEnrollment', enrollmentStatus, isPassed, isFailed, pctVal, studentId);

    // By mode
    addToGroup(byMode, mode, isPassed, isFailed, pctVal, studentId);
    addToSubGroup(byMode, mode, 'byLicense', license, isPassed, isFailed, pctVal, studentId);

    // By enrollment
    addToGroup(byEnrollment, enrollmentStatus, isPassed, isFailed, pctVal, studentId);
    addToSubGroup(byEnrollment, enrollmentStatus, 'byLicense', license, isPassed, isFailed, pctVal, studentId);

    // By class (only for enrolled students)
    if (classCode) {
      var className = cInfo ? cInfo.className : classCode;
      addToGroup(byClass, className + ' (' + classCode + ')', isPassed, isFailed, pctVal, studentId);
      addToSubGroup(byClass, className + ' (' + classCode + ')', 'byLicense', license, isPassed, isFailed, pctVal, studentId);
      addToSubGroup(byClass, className + ' (' + classCode + ')', 'byMode', mode, isPassed, isFailed, pctVal, studentId);
    }

    // By day
    var dayKey = fmtDay(rowDate);
    addToGroup(byDay, dayKey, isPassed, isFailed, pctVal, studentId);
  }

  var overallStats = computeStats(overall);
  overallStats.activeClasses = Object.keys(overall.classes).length;
  overallStats.independentStudents = Object.keys(overall.independentStudents).length;

  return jsonResponse({
    status: 'ok',
    data: {
      overall: overallStats,
      byLicense: computeGroupWithSub(byLicense),
      byMode: computeGroupWithSub(byMode),
      byEnrollment: computeGroupWithSub(byEnrollment),
      byClass: computeGroupWithSub(byClass),
      byDay: computeGroupWithSub(byDay)
    }
  });
}

