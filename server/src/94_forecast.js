// ========== Examiner-commander forecast (predictive PLANNING view) ==========
// Reads the nightly at-risk cache (NO model build on request) and produces two
// forecasts for the examiner-commander:
//   cohortForecast (B') — aggregate expected pass rate of the current practicing
//     pool, by license + overall ("a typical upcoming C1 cohort → ~X% pass").
//   examDayForecast (A') — joins the LIVE registrants (ממתינים in active sessions)
//     to their nightly prediction → "of N registered now, ~X expected to pass,
//     Y at risk → prepare re-exam slots". No-practice-record registrants are
//     counted separately, never silently assumed pass/fail.
function handleExaminerForecast(p) {
  var exData = getSheet('בוחנים').getDataRange().getValues();
  var role = '';
  for (var i = 1; i < exData.length; i++) {
    if (normalizeId(exData[i][1]) === normalizeId(p.examinerId)) { role = String(exData[i][5] || 'בוחן'); break; }
  }
  if (role !== 'מפקד') return jsonResponse({ status: 'error', message: 'אין הרשאת מפקד' });

  var props = PropertiesService.getScriptProperties();
  var computedAt = props.getProperty('atRisk_computedAt') || null;
  var sheet = getSheetIfExists('חיזוי סיכון');
  if (!sheet || sheet.getLastRow() < 2) {
    return jsonResponse({ status: 'ok', data: { computedAt: computedAt, notComputed: true, cohortForecast: null, examDayForecast: null } });
  }
  var rows = sheet.getDataRange().getValues();
  // Cache columns: 1 name, 2 license, 7 site, 8 lastPct, 11 attempt, 13 prob, 14 tier, 17 phone.
  var byPhone = {}, byName = {};
  var cohortByLic = {};
  var cohortBySite = {};
  var cohortBySiteLic = {};
  var cohortAll = { n: 0, sumProb: 0, high: 0, medium: 0, low: 0 };
  function cohortBump(o, prob, tier) {
    o.n++; if (prob != null) o.sumProb += prob;
    if (tier === 'high') o.high++; else if (tier === 'medium') o.medium++; else o.low++;
  }
  for (var r = 1; r < rows.length; r++) {
    var row = rows[r];
    var lic = String(row[2] || '');
    var prob = (row[13] === '' || row[13] == null) ? null : Number(row[13]);
    var tier = String(row[14] || 'low');
    var phone = String(row[17] || '');
    var rec = { name: String(row[1] || ''), license: lic, lastPct: Number(row[8]) || 0, attempt: Number(row[11]) || 1, prob: prob, tier: tier, site: String(row[7] || '') };
    if (phone) byPhone[phone] = rec;
    var nk = ppNormName(row[1]) + '|' + lic;
    if (!byName[nk]) byName[nk] = rec;
    if (!cohortByLic[lic]) cohortByLic[lic] = { n: 0, sumProb: 0, high: 0, medium: 0, low: 0 };
    cohortBump(cohortByLic[lic], prob, tier);
    cohortBump(cohortAll, prob, tier);
    var cSite = String(row[7] || '') || 'לא צוין';
    if (!cohortBySite[cSite]) cohortBySite[cSite] = { n: 0, sumProb: 0, high: 0, medium: 0, low: 0 };
    cohortBump(cohortBySite[cSite], prob, tier);
    var clKey = cSite + '|' + (lic || 'לא צוין');
    if (!cohortBySiteLic[clKey]) cohortBySiteLic[clKey] = { site: cSite, license: (lic || 'לא צוין'), n: 0, sumProb: 0, high: 0, medium: 0, low: 0 };
    cohortBump(cohortBySiteLic[clKey], prob, tier);
  }
  function finalizeCohort(o) {
    return { n: o.n, expectedPassRate: o.n ? Math.round(o.sumProb / o.n) : 0, expectedPasses: Math.round(o.sumProb / 100), high: o.high, medium: o.medium, low: o.low };
  }
  var cohortForecast = { overall: finalizeCohort(cohortAll), byLicense: {}, bySite: {}, bySiteLicense: [] };
  for (var lk in cohortByLic) cohortForecast.byLicense[lk] = finalizeCohort(cohortByLic[lk]);
  for (var csk in cohortBySite) cohortForecast.bySite[csk] = finalizeCohort(cohortBySite[csk]);
  for (var clk in cohortBySiteLic) {
    var cl = cohortBySiteLic[clk], cf2 = finalizeCohort(cl);
    cohortForecast.bySiteLicense.push({ site: cl.site, license: cl.license, n: cf2.n, expectedPassRate: cf2.expectedPassRate, expectedPasses: cf2.expectedPasses, high: cf2.high, medium: cf2.medium, low: cf2.low });
  }
  cohortForecast.bySiteLicense.sort(function(a, b) { return a.site === b.site ? (b.n - a.n) : (a.site < b.site ? -1 : 1); });

  // A' — live registrants in ACTIVE sessions only. Track each session's site so
  // the forecast can be split per site (examiner-allocation planning).
  var activeSessions = {}, sessionSite = {};
  try {
    diagMark('sheet:sessions-forecast');
    var sess = sessionRows();
    var nowT = new Date().getTime();
    for (var s = 1; s < sess.length; s++) {
      var active = sess[s][10] === true || String(sess[s][10]).toUpperCase() === 'TRUE';
      var validUntil = sess[s][9] ? new Date(sess[s][9]).getTime() : 0;
      if (active && (!validUntil || validUntil > nowT)) {
        var scode = String(sess[s][0] || '').trim();
        activeSessions[scode] = true;
        sessionSite[scode] = String(sess[s][3] || '');   // אתר
      }
    }
  } catch (eS) { /* no sessions → examDay stays empty */ }

  var examDay = { registered: 0, matched: 0, noRecord: 0, expectedPasses: 0, expectedFails: 0, high: 0, medium: 0, low: 0, atRisk: [], bySite: {}, bySiteLicense: [] };
  var edSiteLicAcc = {};
  function edSite(site) {
    if (!examDay.bySite[site]) examDay.bySite[site] = { registered: 0, matched: 0, noRecord: 0, expectedPasses: 0, expectedFails: 0, high: 0, medium: 0, low: 0 };
    return examDay.bySite[site];
  }
  function edSiteLic(site, lic) {
    var kk = site + '|' + lic;
    if (!edSiteLicAcc[kk]) edSiteLicAcc[kk] = { site: site, license: lic, registered: 0, matched: 0, noRecord: 0, expectedPasses: 0, high: 0, medium: 0, low: 0 };
    return edSiteLicAcc[kk];
  }
  try {
    // Only registrations of sessions that are still open matter, and a session
    // lives 8 hours — so two days of rows cover every one of them. Columns:
    // A code, C name, D phone, F status, I licence.
    diagMark('sheet:pending-forecast');
    var waitCutoff = new Date(Date.now() - 2 * 86400000);
    var wait = readRowsSince(getSheet('ממתינים'), 4, waitCutoff, [[1, 1], [3, 2], [6, 1], [9, 1]]).rows;
    for (var w = 1; w < wait.length; w++) {
      var code = String(wait[w][0] || '').trim();
      if (!activeSessions[code]) continue;
      var status = String(wait[w][5] || '');
      if (status === 'הושלם' || status === 'פסול' || status === 'בוטל' || status === 'נדחה') continue;   // already resolved
      var siteA = sessionSite[code] || 'לא צוין';
      var bs = edSite(siteA);
      examDay.registered++; bs.registered++;
      var wPhone = ppNormPhone(wait[w][3]);
      var wLic = String(wait[w][8] || '');
      var wName = String(wait[w][2] || '');
      var sl = edSiteLic(siteA, wLic || 'לא צוין'); sl.registered++;
      var hit = (wPhone && byPhone[wPhone]) || byName[ppNormName(wName) + '|' + wLic] || null;
      if (!hit || hit.prob == null) { examDay.noRecord++; bs.noRecord++; sl.noRecord++; continue; }
      examDay.matched++; bs.matched++; sl.matched++;
      examDay.expectedPasses += hit.prob / 100; bs.expectedPasses += hit.prob / 100; sl.expectedPasses += hit.prob / 100;
      if (hit.tier === 'high') { examDay.high++; bs.high++; sl.high++; } else if (hit.tier === 'medium') { examDay.medium++; bs.medium++; sl.medium++; } else { examDay.low++; bs.low++; sl.low++; }
      if (hit.tier === 'high' || hit.tier === 'medium') examDay.atRisk.push({ name: wName, license: wLic, site: siteA, prob: hit.prob, tier: hit.tier, lastPct: hit.lastPct, attempt: hit.attempt });
    }
  } catch (eW) { /* no waiting sheet */ }
  examDay.expectedPasses = Math.round(examDay.expectedPasses);
  examDay.expectedFails = Math.max(0, examDay.matched - examDay.expectedPasses);
  for (var bsk in examDay.bySite) {
    var b2 = examDay.bySite[bsk];
    b2.expectedPasses = Math.round(b2.expectedPasses);
    b2.expectedFails = Math.max(0, b2.matched - b2.expectedPasses);
  }
  for (var slk in edSiteLicAcc) {
    var sl2 = edSiteLicAcc[slk];
    sl2.expectedPasses = Math.round(sl2.expectedPasses);
    sl2.expectedFails = Math.max(0, sl2.matched - sl2.expectedPasses);
    examDay.bySiteLicense.push(sl2);
  }
  examDay.bySiteLicense.sort(function(a, b) { return a.site === b.site ? (b.registered - a.registered) : (a.site < b.site ? -1 : 1); });
  examDay.atRisk.sort(function(a, b) { return a.prob - b.prob; });
  examDay.atRisk = examDay.atRisk.slice(0, 50);

  return jsonResponse({ status: 'ok', data: {
    computedAt: computedAt,
    cohortForecast: cohortForecast,
    examDayForecast: examDay
  } });
}

