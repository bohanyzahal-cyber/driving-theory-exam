function handleTeacherCreateClass(p) {
  var maintenance = practiceWriteGuard(); if (maintenance) return maintenance;   // r24
  var code = generateClassCode();
  var className = p.className || 'כיתה חדשה';
  var license = p.license || 'B';
  var sheet = getSheet('כיתות');
  // Get teacher name
  var tSheet = getSheet('מורים');
  var tData = tSheet.getDataRange().getValues();
  var teacherName = '';
  var teacherSite = '';
  for (var i = 1; i < tData.length; i++) {
    if (normalizeId(tData[i][1]) === normalizeId(p.teacherId)) {
      teacherName = tData[i][0];
      teacherSite = String(tData[i][9] || '');
      break;
    }
  }
  sheet.appendRow([code, className, normalizeId(p.teacherId), teacherName, license, nowISO(), 'כן', teacherSite]);
  return jsonResponse({ status: 'ok', classCode: code, className: className });
}

function handleTeacherCloseClass(p) {
  var maintenance = practiceWriteGuard(); if (maintenance) return maintenance;   // r24
  var sheet = getSheet('כיתות');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(p.classCode).trim() &&
        normalizeId(data[i][2]) === normalizeId(p.teacherId)) {
      sheet.getRange(i + 1, 7).setValue('לא');
      return jsonResponse({ status: 'ok' });
    }
  }
  return jsonResponse({ status: 'error', message: 'כיתה לא נמצאה' });
}

function handleTeacherDeleteClass(p) {
  var maintenance = practiceWriteGuard(); if (maintenance) return maintenance;   // r24
  var classCode = String(p.classCode || '').trim();
  if (!classCode) return jsonResponse({ status: 'error', message: 'חסר קוד כיתה' });

  var classSheet = getSheet('כיתות');
  var classData = classSheet.getDataRange().getValues();
  var classRowIdx = -1;
  for (var i = 1; i < classData.length; i++) {
    if (String(classData[i][0]).trim() === classCode &&
        normalizeId(classData[i][2]) === normalizeId(p.teacherId)) {
      classRowIdx = i;
      break;
    }
  }
  if (classRowIdx === -1) return jsonResponse({ status: 'error', message: 'כיתה לא נמצאה או שאין הרשאה' });

  // Safety: only allow deletion of CLOSED classes (active = 'לא')
  if (String(classData[classRowIdx][6]).trim() === 'כן') {
    return jsonResponse({ status: 'error', message: 'יש לסגור את הכיתה לפני מחיקה' });
  }

  // Archive the class metadata BEFORE deleting the row. Practice results in
  // 'תוצאות תרגול' are preserved for history (see NOTE below), and once the class
  // row is gone the reports can no longer resolve its teacher/name → they showed
  // "לא ידוע". This archive lets the commander dashboard still attribute those
  // orphaned rows to the real teacher and tag them "(כיתה שנמחקה)", so a genuine
  // deletion is distinguishable from a truly unrecognized/forged class code.
  try {
    var cRow = classData[classRowIdx];
    getSheet('כיתות שנמחקו').appendRow([
      String(cRow[0] || '').trim(), // קוד כיתה
      String(cRow[1] || ''),        // שם כיתה
      normalizeId(cRow[2]),         // מורה ת.ז.
      String(cRow[3] || ''),        // שם מורה
      String(cRow[4] || ''),        // דרגה
      String(cRow[7] || ''),        // אתר
      nowISO()                      // תאריך מחיקה
    ]);
  } catch (archiveErr) { /* non-fatal: deletion proceeds even if archiving fails */ }

  // Delete the class row
  classSheet.deleteRow(classRowIdx + 1);

  // Delete all students enrolled in this class (cleanup roster)
  var studentsRemoved = 0;
  var studSheet = getSheet('תלמידי כיתות');
  var studData = studSheet.getDataRange().getValues();
  for (var s = studData.length - 1; s >= 1; s--) {
    if (String(studData[s][0]).trim() === classCode) {
      studSheet.deleteRow(s + 1);
      studentsRemoved++;
    }
  }

  // NOTE: practice results in 'תוצאות תרגול' are intentionally preserved for historical reporting.

  return jsonResponse({ status: 'ok', studentsRemoved: studentsRemoved });
}

function handleTeacherRemoveStudent(p) {
  var maintenance = practiceWriteGuard(); if (maintenance) return maintenance;   // r24
  var sheet = getSheet('תלמידי כיתות');
  var data = sheet.getDataRange().getValues();
  // Verify teacher owns this class
  var classSheet = getSheet('כיתות');
  var classData = classSheet.getDataRange().getValues();
  var ownsClass = false;
  for (var c = 1; c < classData.length; c++) {
    if (String(classData[c][0]).trim() === String(p.classCode).trim() &&
        normalizeId(classData[c][2]) === normalizeId(p.teacherId)) {
      ownsClass = true; break;
    }
  }
  if (!ownsClass) return jsonResponse({ status: 'error', message: 'אין הרשאה' });

  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]).trim() === String(p.classCode).trim() &&
        String(data[i][2]).trim() === String(p.studentId).trim()) {
      sheet.deleteRow(i + 1);
      return jsonResponse({ status: 'ok' });
    }
  }
  return jsonResponse({ status: 'error', message: 'תלמיד לא נמצא' });
}

function handleTeacherGetClasses(p) {
  var sheet = getSheet('כיתות');
  var data = sheet.getDataRange().getValues();
  var studSheet = getSheet('תלמידי כיתות');
  var studData = studSheet.getDataRange().getValues();

  // Count students per class
  var studentCounts = {};
  for (var s = 1; s < studData.length; s++) {
    var cc = String(studData[s][0]).trim();
    studentCounts[cc] = (studentCounts[cc] || 0) + 1;
  }

  var classes = [];
  for (var i = 1; i < data.length; i++) {
    if (normalizeId(data[i][2]) === normalizeId(p.teacherId)) {
      var classCode = String(data[i][0]).trim();
      classes.push({
        code: classCode,
        name: data[i][1],
        license: data[i][4] || 'B',
        created: data[i][5],
        active: data[i][6] === 'כן',
        studentCount: studentCounts[classCode] || 0
      });
    }
  }
  return jsonResponse({ status: 'ok', classes: classes });
}

function handleTeacherClassDetails(p) {
  var classCode = String(p.classCode || '').trim();
  if (!classCode) return jsonResponse({ status: 'error', message: 'חסר קוד כיתה' });

  // Verify teacher owns this class
  var classSheet = getSheet('כיתות');
  var classData = classSheet.getDataRange().getValues();
  var classInfo = null, classCreated = null;
  for (var c = 1; c < classData.length; c++) {
    if (String(classData[c][0]).trim() === classCode && normalizeId(classData[c][2]) === normalizeId(p.teacherId)) {
      classInfo = { code: classCode, name: classData[c][1], license: classData[c][4], active: classData[c][6] === 'כן' };
      classCreated = parseSheetDateTime(classData[c][5]);   // col F = תאריך יצירה
      break;
    }
  }
  if (!classInfo) return jsonResponse({ status: 'error', message: 'כיתה לא נמצאה' });

  // Get students in class
  var studSheet = getSheet('תלמידי כיתות');
  var studData = studSheet.getDataRange().getValues();
  var studentIds = [];
  var studentMap = {};
  for (var s = 1; s < studData.length; s++) {
    if (String(studData[s][0]).trim() === classCode) {
      var sid = String(studData[s][2]).trim();
      studentIds.push(sid);
      studentMap[sid] = { name: studData[s][1], id: sid, joined: studData[s][3] };
    }
  }

  // Get practice results for these students.
  // 17/09/2026: this ran 13 times in one exam window at 19-32 s each, every call
  // a FULL read of the 107k-row practice sheet including its two JSON blob
  // columns. Two bounds, both exact:
  //   rows    — nothing before the class existed can belong to the class;
  //   columns — everything except N ('פירוט שגויות', ~2 KB of wrong-question
  //             text per row) which nothing on this screen shows. Column O
  //             ('פירוט לפי נושא') STAYS: it is the per-topic breakdown the
  //             student card draws, and it is two orders of magnitude smaller.
  // ⚠ Index another column here and add it to the colSpec.
  diagMark('sheet:practice-class');
  var classRead = readRowsSince(getSheet('תוצאות תרגול'), 0, classCreated, [[1, 13], [15, 1]]);
  var resData = classRead.rows;
  diagMark('sheet:practice-class-done:' + classRead.mode);
  var inClass = {};
  for (var sj = 0; sj < studentIds.length; sj++) inClass[studentIds[sj]] = true;
  var studentResults = {};
  for (var r = 1; r < resData.length; r++) {
    var rSid = String(resData[r][1]).trim();
    if (String(resData[r][3]).trim() !== classCode || !inClass[rSid]) continue;
    if (!studentResults[rSid]) studentResults[rSid] = [];
    studentResults[rSid].push({
      date: resData[r][0],
      mode: resData[r][4],
      license: resData[r][5],
      score: resData[r][6],
      total: resData[r][7],
      percent: resData[r][8],
      passed: resData[r][9],
      time: resData[r][10],
      category: resData[r][11] || '',
      language: resData[r][12] || 'he',
      categoryBreakdown: resData[r][14] || ''
    });
  }

  // Build student summaries
  var students = [];
  for (var si = 0; si < studentIds.length; si++) {
    var id = studentIds[si];
    var info = studentMap[id];
    var results = studentResults[id] || [];
    var totalExams = 0, totalPassed = 0, scores = [], lastActive = '';
    var categoryErrors = {};
    for (var ri = 0; ri < results.length; ri++) {
      var res = results[ri];
      var pctVal = Number(res.percent) || 0;
      if (pctVal > 0) scores.push(pctVal);
      if (res.mode === 'exam') {
        totalExams++;
        if (res.passed === 'עבר' || res.passed === true) totalPassed++;
      }
      if (res.date && (!lastActive || String(res.date) > String(lastActive))) lastActive = res.date;
      // Aggregate category errors
      if (res.categoryBreakdown) {
        try {
          var cb = typeof res.categoryBreakdown === 'string' ? JSON.parse(res.categoryBreakdown) : res.categoryBreakdown;
          for (var cat in cb) {
            if (!categoryErrors[cat]) categoryErrors[cat] = { correct: 0, total: 0 };
            categoryErrors[cat].correct += (cb[cat].correct || 0);
            categoryErrors[cat].total += (cb[cat].total || 0);
          }
        } catch(e) {}
      }
    }
    var avgScore = 0;
    if (scores.length > 0) {
      var sum = 0;
      for (var sc = 0; sc < scores.length; sc++) sum += scores[sc];
      avgScore = Math.round(sum / scores.length);
    }
    students.push({
      name: info.name,
      id: id,
      joined: info.joined,
      totalPractices: results.length,
      totalExams: totalExams,
      totalPassed: totalPassed,
      avgScore: avgScore,
      lastActive: lastActive,
      categoryErrors: categoryErrors,
      recentResults: results.slice(-10) // Last 10 results
    });
  }

  return jsonResponse({ status: 'ok', classInfo: classInfo, students: students });
}

function handleTeacherDashboard(p) {
  // Overview of all classes for a teacher
  return handleTeacherGetClasses(p);
}

function handleTeacherExportData(p) {
  var classCode = String(p.classCode || '').trim();
  if (!classCode) return jsonResponse({ status: 'error', message: 'חסר קוד כיתה' });

  // Verify ownership
  var classSheet = getSheet('כיתות');
  var classData = classSheet.getDataRange().getValues();
  var owns = false, exportCreated = null;
  for (var c = 1; c < classData.length; c++) {
    if (String(classData[c][0]).trim() === classCode && normalizeId(classData[c][2]) === normalizeId(p.teacherId)) {
      owns = true;
      exportCreated = parseSheetDateTime(classData[c][5]);
      break;
    }
  }
  if (!owns) return jsonResponse({ status: 'error', message: 'אין הרשאה' });

  // Get all results for this class. Rows are bounded by the class's creation
  // date (nothing older can belong to it); every COLUMN stays — this is the
  // export, and the JSON blobs are the point of it.
  var resData = readRowsSince(getSheet('תוצאות תרגול'), 0, exportCreated).rows;
  var headers = resData[0];
  var rows = [];
  for (var r = 1; r < resData.length; r++) {
    if (String(resData[r][3]).trim() === classCode) {
      var row = {};
      for (var h = 0; h < headers.length; h++) row[headers[h]] = resData[r][h];
      rows.push(row);
    }
  }
  return jsonResponse({ status: 'ok', headers: headers, rows: rows });
}

function handleStudentJoinClass(p) {
  var maintenance = practiceWriteGuard(); if (maintenance) return maintenance;   // r24
  var classCode = String(p.classCode || '').trim().toUpperCase();
  var studentName = String(p.studentName || '').trim();
  var studentId = String(p.studentId || '').trim();
  if (!classCode || !studentName || !studentId) {
    return jsonResponse({ status: 'error', message: 'חסרים פרטים (קוד כיתה, שם, מזהה)' });
  }

  // Verify class exists and active
  var classSheet = getSheet('כיתות');
  var classData = classSheet.getDataRange().getValues();
  var classInfo = null;
  for (var c = 1; c < classData.length; c++) {
    if (String(classData[c][0]).trim() === classCode) {
      if (classData[c][6] !== 'כן') return jsonResponse({ status: 'error', message: 'הכיתה אינה פעילה' });
      classInfo = { name: classData[c][1], teacherName: classData[c][3], license: classData[c][4] };
      break;
    }
  }
  if (!classInfo) return jsonResponse({ status: 'error', message: 'כיתה לא נמצאה' });

  // Check if already enrolled by studentId (same device/browser)
  var studSheet = getSheet('תלמידי כיתות');
  var studData = studSheet.getDataRange().getValues();
  for (var s = 1; s < studData.length; s++) {
    if (String(studData[s][0]).trim() === classCode && String(studData[s][2]).trim() === studentId) {
      return jsonResponse({ status: 'ok', message: 'כבר רשום בכיתה', className: classInfo.name, teacherName: classInfo.teacherName, license: classInfo.license });
    }
  }

  // Check if same name is already in this class with a DIFFERENT studentId (joined from another device/browser).
  // If so, return the existing studentId so the new device adopts it — prevents duplicate roster entries.
  var normName = studentName.toLowerCase().replace(/\s+/g, ' ');
  for (var s2 = 1; s2 < studData.length; s2++) {
    if (String(studData[s2][0]).trim() === classCode) {
      var existingName = String(studData[s2][1]).trim().toLowerCase().replace(/\s+/g, ' ');
      if (existingName === normName) {
        return jsonResponse({
          status: 'ok',
          message: 'מצאנו שאתה כבר רשום בכיתה הזו ממכשיר אחר. הנתונים שלך אוחדו.',
          existingStudentId: String(studData[s2][2]).trim(),
          className: classInfo.name,
          teacherName: classInfo.teacherName,
          license: classInfo.license
        });
      }
    }
  }

  studSheet.appendRow([classCode, studentName, studentId, nowISO()]);
  return jsonResponse({ status: 'ok', message: 'הצטרפת לכיתה בהצלחה!', className: classInfo.name, teacherName: classInfo.teacherName, license: classInfo.license });
}

function handleSubmitPracticeResult(p) {
  var maintenance = practiceWriteGuard(); if (maintenance) return maintenance;   // r24
  var studentId = String(p.studentId || '').trim();
  var classCode = String(p.classCode || '').trim();
  // Rate limit: cap public practice-result writes so the תוצאות תרגול sheet (which
  // feeds the teacher/commander stats) can't be flooded with fabricated rows.
  var prRlErr = requireRateLimit('submitPracticeResult', (studentId || classCode || 'anon'), 30, 60);
  if (prRlErr) return prRlErr;
  var sheet = getSheet('תוצאות תרגול');
  var mode = String(p.mode || 'exam');
  var license = String(p.license || 'B');
  var score = Number(p.score) || 0;
  var total = Number(p.total) || 0;
  var percent = Number(p.percent) || 0;
  var passed = percent >= 86 ? 'עבר' : 'נכשל';
  var time = String(p.time || '');
  var category = String(p.category || '');
  var language = String(p.language || 'he');
  var wrongDetails = '';
  try { wrongDetails = typeof p.wrongDetails === 'string' ? p.wrongDetails : JSON.stringify(p.wrongDetails || ''); } catch(e) {}
  var categoryBreakdown = '';
  try { categoryBreakdown = typeof p.categoryBreakdown === 'string' ? p.categoryBreakdown : JSON.stringify(p.categoryBreakdown || ''); } catch(e) {}

  sheet.appendRow([todayStr(), studentId, String(p.studentName || ''), classCode, mode, license, score, total, percent, passed, time, category, language, wrongDetails, categoryBreakdown, String(p.phone || '')]);
  return jsonResponse({ status: 'ok' });
}

function handleLoadStudentProgress(p) {
  var name = String(p.studentName || '').trim();
  var classCode = String(p.classCode || '').trim().toUpperCase();
  if (!name || !classCode) {
    return jsonResponse({ status: 'error', message: 'חסרים פרטים' });
  }
  var key = name.toLowerCase() + '|' + classCode;
  var sheet = getSheet('התקדמות תלמידים');
  var row = findRow(sheet, 2, key);
  if (row === -1) {
    return jsonResponse({ status: 'ok', found: false });
  }
  var data = sheet.getRange(row, 1, 1, 7).getValues()[0];
  return jsonResponse({
    status: 'ok',
    found: true,
    streak: data[3] || '{}',
    wrongQs: data[4] || '[]',
    history: data[5] || '[]',
    lastUpdated: data[6] || ''
  });
}

function handleSaveStudentProgress(p) {
  var maintenance = practiceWriteGuard(); if (maintenance) return maintenance;   // r24
  var name = String(p.studentName || '').trim();
  var classCode = String(p.classCode || '').trim().toUpperCase();
  if (!name || !classCode) {
    return jsonResponse({ status: 'error', message: 'חסרים פרטים' });
  }
  var key = name.toLowerCase() + '|' + classCode;
  var streak = String(p.streak || '{}');
  var wrongQs = String(p.wrongQs || '[]');
  var history = String(p.history || '[]');
  var sheet = getSheet('התקדמות תלמידים');
  var row = findRow(sheet, 2, key);
  if (row === -1) {
    sheet.appendRow([name, classCode, key, streak, wrongQs, history, nowISO()]);
  } else {
    sheet.getRange(row, 1, 1, 7).setValues([[name, classCode, key, streak, wrongQs, history, nowISO()]]);
  }
  return jsonResponse({ status: 'ok' });
}
