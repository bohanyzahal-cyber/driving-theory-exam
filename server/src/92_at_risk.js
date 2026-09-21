// ========== At-risk examinee list (PREVENTIVE — step 2 of predictive BI) ==========
// The teacher-commander's preventive view: for every student who practiced in the
// window and has NOT yet passed the real theory exam, predict their pass
// probability from the calibration model (license × upcoming-attempt × last
// practice score) and rank worst-first. Lets the training commander route
// weak soldiers to more practice BEFORE they burn a real exam slot. Highest
// value for C1 first-timers (huge volume, ~22% base pass, and practice score
// cleanly separates ~10% from ~73%). Complements repeatFailures (which is
// retrospective — already failed 2+); this catches them before the first fail.
// Heavy computation for the WHOLE fleet — builds the model once and scores every
// recently-practiced student (no scope filter). Called ONLY by the nightly cache
// rebuild, never on a dashboard request. Returns the full ranked list + build
// summary + computedAt so the cache can be scope-filtered cheaply on read.
function computeAtRiskAll(opts) {
  opts = opts || {};
  var lookbackDays = opts.lookbackDays || 30;
  var model = buildPassProbabilityModel({ lookbackDays: lookbackDays });

  // Class map: code → {teacherId, teacherName, className, license, site}.
  var classData = getSheet('כיתות').getDataRange().getValues();
  var classMap = {};
  for (var c = 1; c < classData.length; c++) {
    classMap[String(classData[c][0]).trim()] = {
      teacherId: normalizeId(classData[c][2]),
      teacherName: String(classData[c][3] || ''),
      className: String(classData[c][1] || ''),
      license: String(classData[c][4] || ''),
      site: String(classData[c][7] || '')
    };
  }
  var deletedClassMap = getDeletedClassMap();

  // Exam-history index → upcoming attempt number + everPassed. Keyed by phone AND
  // name|license, mirroring the model join.
  // Live + archive (B5), columns C name, D phone, E licence, H pass, O attempt.
  // ⚠ Index another column and add it to the colSpec — a missing column reads ''.
  var examData = readResultsSince(null, [[3, 3], [8, 1], [15, 1]]).rows;
  var histByPhone = {}, histByName = {};
  function histBump(idx, key, attempt, passed) {
    if (!key) return;
    if (!idx[key]) idx[key] = { attempts: 0, everPassed: false };
    if (attempt > idx[key].attempts) idx[key].attempts = attempt;
    if (passed) idx[key].everPassed = true;
  }
  for (var e = 1; e < examData.length; e++) {
    var ePassedStr = String(examData[e][7] || '');
    if (ePassedStr === 'בוטל') continue;
    var eName = ppNormName(examData[e][2]);
    var eLic = String(examData[e][4] || '').trim();
    var ePhone = ppNormPhone(examData[e][3]);
    var eAtt = Number(examData[e][14]) || 1;
    var ePassed = (ePassedStr === 'עבר');
    histBump(histByPhone, ePhone, eAtt, ePassed);
    histBump(histByName, eName + '|' + eLic, eAtt, ePassed);
  }

  // Group practice rows by student within the window — no scope filter here (the
  // cache holds everyone; the read handler filters by caller scope).
  var windowStart = new Date();
  windowStart.setDate(windowStart.getDate() - lookbackDays);
  // Rows: only the window this loop keeps (it drops anything older itself).
  // Columns: A date, B studentId, C name, D class, F licence, I percent,
  // P phone — never the two JSON blobs in N/O.
  var practiceData = readRowsSince(getSheet('תוצאות תרגול'), 0, windowStart,
    [[1, 6], [9, 1], [16, 1]]).rows;
  var students = {};
  for (var r = 1; r < practiceData.length; r++) {
    var pDate = parseSheetDate(practiceData[r][0]);
    if (!pDate || pDate < windowStart) continue;
    var classCode = String(practiceData[r][3] || '').trim();
    var cInfo = resolveClassInfo(classCode, classMap, deletedClassMap);
    var pct = ppParsePct(practiceData[r][8]);
    if (pct < 0) continue;
    var name = String(practiceData[r][2] || '');
    var lic = String(practiceData[r][5] || cInfo.license || '').trim();
    var phone = (practiceData[r].length > 15) ? ppNormPhone(practiceData[r][15]) : '';
    var studentId = String(practiceData[r][1] || '');
    var rowResolved = !cInfo.unresolved;   // class code mapped to a real (active/deleted) class
    var key = phone ? ('p:' + phone) : (studentId ? ('s:' + studentId) : ('n:' + ppNormName(name) + '|' + lic));
    if (!students[key]) {
      students[key] = { name: name, license: lic, phone: phone, studentId: studentId, classCode: classCode, teacherId: cInfo.teacherId || '', className: cInfo.className, teacherName: cInfo.teacherName, site: cInfo.site || '', resolved: rowResolved, recs: [] };
    } else if (!students[key].resolved && rowResolved) {
      // Upgrade to a resolved class if an earlier row had a blank/unknown code.
      students[key].classCode = classCode; students[key].teacherId = cInfo.teacherId || '';
      students[key].className = cInfo.className; students[key].teacherName = cInfo.teacherName;
      students[key].site = cInfo.site || ''; students[key].resolved = true;
    }
    students[key].recs.push({ date: pDate, pct: pct });
    if (name) students[key].name = name;
    if (lic) students[key].license = lic;
    if (studentId && !students[key].studentId) students[key].studentId = studentId;
  }

  // Roster fallback — many students practice WITHOUT a class code (it's optional)
  // so their rows show "קוד לא מזוהה" and can't be attributed to a teacher. Recover
  // the attribution from the class rosters (תלמידי כיתות: code, name, studentId):
  // match by student id first, then by normalized name. Best-effort — a name that
  // appears in two classes maps to whichever was seen first.
  var rosterById = {}, rosterByName = {};
  try {
    var rosterData = getSheet('תלמידי כיתות').getDataRange().getValues();
    for (var rs = 1; rs < rosterData.length; rs++) {
      var rCode = String(rosterData[rs][0] || '').trim();
      if (!rCode) continue;
      var rSid = String(rosterData[rs][2] || '').trim();
      var rNk = ppNormName(rosterData[rs][1]);
      if (rSid && !rosterById[rSid]) rosterById[rSid] = rCode;
      if (rNk && !rosterByName[rNk]) rosterByName[rNk] = rCode;
    }
  } catch (eRoster) { /* no roster sheet → fallback simply does nothing */ }
  for (var uk in students) {
    var us = students[uk];
    if (us.resolved) continue;
    var rc = (us.studentId && rosterById[us.studentId]) || rosterByName[ppNormName(us.name)] || '';
    if (!rc) continue;
    var rci = resolveClassInfo(rc, classMap, deletedClassMap);
    if (rci.unresolved) continue;
    us.classCode = rc; us.teacherId = rci.teacherId || '';
    us.className = rci.className; us.teacherName = rci.teacherName;
    us.site = rci.site || ''; us.resolved = true; us.attributedViaRoster = true;
  }

  // Junk-name filter — drops shared/demo entries (a "." with 100+ sessions, a
  // cycle name typed into the name field, punctuation/number-only names).
  function looksLikeName(s) {
    var t = String(s || '').trim();
    if (t.length < 2) return false;
    var letters = t.replace(/[^A-Za-z֐-׿]/g, '');
    return letters.length >= 2;
  }
  var out = [];
  var summary = { high: 0, medium: 0, low: 0, alreadyPassed: 0, junk: 0, total: 0 };
  for (var k in students) {
    var st = students[k];
    if (!looksLikeName(st.name)) { summary.junk++; continue; }
    st.recs.sort(function(a, b) { return b.date - a.date; });
    var latest = st.recs[0];
    var oldest = st.recs[st.recs.length - 1];
    var sessions = st.recs.length;
    var trendPts = sessions > 1 ? (latest.pct - oldest.pct) : 0;
    var trend = sessions < 2 ? 'single' : (trendPts > 3 ? 'up' : (trendPts < -3 ? 'down' : 'flat'));
    var hist = (st.phone && histByPhone[st.phone]) || histByName[ppNormName(st.name) + '|' + st.license] || null;
    if (hist && hist.everPassed) { summary.alreadyPassed++; continue; }
    var upcomingAttempt = (hist ? hist.attempts : 0) + 1;
    var pred = predictPassProbability(model, { license: st.license, attempt: upcomingAttempt, lastPct: latest.pct });
    var prob = pred ? pred.prob : null;
    var tier = prob == null ? 'low' : (prob < 40 ? 'high' : (prob < 65 ? 'medium' : 'low'));
    summary[tier]++; summary.total++;
    out.push({
      name: st.name, license: st.license, classCode: st.classCode, teacherId: st.teacherId,
      teacherName: st.teacherName, className: st.className, site: st.site,
      lastPct: Math.round(latest.pct), sessions: sessions, trend: trend,
      attempt: upcomingAttempt, everTested: !!hist,
      prob: prob, tier: tier, confidence: pred ? pred.confidence : 'low', matchedByPhone: !!st.phone, phone: st.phone || ''
    });
  }
  out.sort(function(a, b) {
    var pa = a.prob == null ? 999 : a.prob, pb = b.prob == null ? 999 : b.prob;
    if (pa !== pb) return pa - pb;
    if (a.lastPct !== b.lastPct) return a.lastPct - b.lastPct;
    return b.sessions - a.sessions;
  });
  return { computedAtMs: new Date().getTime(), students: out, summary: summary, modelBaseRate: Math.round(model.base.rate * 100) };
}

// Nightly job — recompute the whole at-risk list and persist it to the
// 'חיזוי סיכון' sheet + script properties (timestamp, summary, base rate). Run
// by a time-based trigger (see installAtRiskTrigger) at ~03:00 so the heavy
// model build never lands during exam/practice hours. Dashboards then READ this
// cache instead of rebuilding — no per-request model build, no midday load.
// The job-running FLAG is gone with the warmup that read it (r25). What the two
// nightly jobs need from each other is stronger than a flag anyway: this one
// reads 'תוצאות' live + archive, and the 01:00 archive MOVES rows between those
// two sheets, so a run that overlapped it could count a row twice or not at all.
// The script lock the archive already holds is the real interlock; 03:00 is two
// hours later, so waiting for it is a formality that costs nothing.
function rebuildAtRiskCache() {
  var lock = LockService.getScriptLock();
  var held = false;
  try { held = lock.tryLock(60000); } catch (e) { held = false; }
  if (!held) Logger.log('rebuildAtRiskCache: the archive still holds the lock — computing anyway');
  try { return rebuildAtRiskCacheInner(); }
  finally { if (held) lock.releaseLock(); }
}
function rebuildAtRiskCacheInner() {
  var res = computeAtRiskAll({ lookbackDays: 30 });
  var sheet = getSheet('חיזוי סיכון');
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
  var computedAtStr = Utilities.formatDate(new Date(res.computedAtMs), 'Asia/Jerusalem', 'yyyy-MM-dd HH:mm');
  var rows = res.students.map(function(s) {
    return [computedAtStr, s.name, s.license, s.classCode, s.teacherId, s.teacherName, s.className, s.site,
      s.lastPct, s.sessions, s.trend, s.attempt, s.everTested, (s.prob == null ? '' : s.prob), s.tier, s.confidence, s.matchedByPhone, (s.phone || '')];
  });
  if (rows.length) sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  var props = PropertiesService.getScriptProperties();
  props.setProperty('atRisk_computedAt', computedAtStr);
  props.setProperty('atRisk_summary', JSON.stringify(res.summary));
  props.setProperty('atRisk_modelBaseRate', String(res.modelBaseRate));
  return { computed: res.students.length, computedAt: computedAtStr };
}

// One-time setup — installs the nightly trigger. Run once from the Apps Script
// editor (Run → installAtRiskTrigger). Idempotent: removes any prior copy first.
function installAtRiskTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'rebuildAtRiskCache') ScriptApp.deleteTrigger(triggers[i]);
  }
  ScriptApp.newTrigger('rebuildAtRiskCache').timeBased().atHour(3).everyDays(1).inTimezone('Asia/Jerusalem').create();
  return 'Nightly at-risk trigger installed (~03:00 Asia/Jerusalem).';
}

// Dashboard read — fast. Reads the pre-computed cache and filters by the caller's
// scope: commanders see by level (global / their site / their managed sites),
// a regular teacher sees ONLY their own classes (by teacher ת.ז.). No model
// build, no big-sheet scan of תוצאות/תוצאות תרגול on the request path.
function handleTeacherAtRiskList(p) {
  var tData = getSheet('מורים').getDataRange().getValues();
  var role = '', userSite = '', myId = normalizeId(p.teacherId), found = false;
  for (var i = 1; i < tData.length; i++) {
    if (normalizeId(tData[i][1]) === myId) { role = String(tData[i][8] || 'מורה'); userSite = String(tData[i][9] || ''); found = true; break; }
  }
  if (!found) return jsonResponse({ status: 'error', message: 'מורה לא נמצא' });

  var isGlobal = (role === 'מפקד' || role === 'מפקד ראשי' || role === 'אדמין');
  var isLocal = (role === 'מפקד מקומי');
  var isMultiSite = isKdtzRole(role);
  var isCommander = isGlobal || isLocal || isMultiSite;
  var managedSites = [];
  if (isMultiSite) {
    managedSites = String(userSite || '').split(',').map(function(s) { return s.trim(); }).filter(function(s) { return s; });
  }

  var props = PropertiesService.getScriptProperties();
  var computedAt = props.getProperty('atRisk_computedAt') || null;
  var modelBaseRate = Number(props.getProperty('atRisk_modelBaseRate') || 0);
  var buildSummary = null;
  try { buildSummary = JSON.parse(props.getProperty('atRisk_summary') || 'null'); } catch (eBS) { buildSummary = null; }

  var sheet = getSheetIfExists('חיזוי סיכון');
  if (!sheet || sheet.getLastRow() < 2) {
    return jsonResponse({ status: 'ok', data: {
      computedAt: computedAt, notComputed: true,
      summary: { high: 0, medium: 0, low: 0, total: 0 }, students: []
    } });
  }
  var rows = sheet.getDataRange().getValues();
  var out = [];
  var summary = { high: 0, medium: 0, low: 0, total: 0 };
  for (var r2 = 1; r2 < rows.length; r2++) {
    var row = rows[r2];
    var site = String(row[7] || '');
    var teacherId = normalizeId(row[4]);
    if (isCommander) {
      if (isLocal && userSite && site !== userSite) continue;
      if (isMultiSite && managedSites.indexOf(site) === -1) continue;
      // isGlobal → no filter
    } else {
      if (!teacherId || teacherId !== myId) continue;   // regular teacher → own classes only
    }
    var tier = String(row[14] || 'low');
    if (tier === 'high') summary.high++; else if (tier === 'medium') summary.medium++; else summary.low++;
    summary.total++;
    out.push({
      name: String(row[1] || ''), license: String(row[2] || ''), className: String(row[6] || ''),
      teacherName: String(row[5] || ''), site: site,
      lastPct: Number(row[8]) || 0, sessions: Number(row[9]) || 0, trend: String(row[10] || ''),
      attempt: Number(row[11]) || 1, everTested: (row[12] === true || String(row[12]).toUpperCase() === 'TRUE'),
      prob: (row[13] === '' || row[13] == null) ? null : Number(row[13]),
      tier: tier, confidence: String(row[15] || 'low'),
      matchedByPhone: (row[16] === true || String(row[16]).toUpperCase() === 'TRUE')
    });
  }
  out.sort(function(a, b) {
    var pa = a.prob == null ? 999 : a.prob, pb = b.prob == null ? 999 : b.prob;
    if (pa !== pb) return pa - pb;
    if (a.lastPct !== b.lastPct) return a.lastPct - b.lastPct;
    return b.sessions - a.sessions;
  });

  return jsonResponse({ status: 'ok', data: {
    computedAt: computedAt,
    modelBaseRate: modelBaseRate,
    scope: isCommander ? (isGlobal ? 'כל האתרים' : (isLocal ? userSite : managedSites.join(', '))) : 'הכיתות שלי',
    buildSummary: buildSummary,
    coverageNote: 'תלמיד מזוהה לפי טלפון אם הוזן, אחרת לפי שם+דרגה. ניסיון = מספר הניסיון הצפוי במבחן האמיתי.',
    summary: summary,
    truncated: out.length > 200,
    students: out.slice(0, 200)
  } });
}

