// Commander-only: return every active, non-expired session across all examiners.
// Used by the commander UI to load and inspect/correct results in another
// examiner's session (audit / appeal-committee workflow).
// Center-commander aggregate report across multiple sites.
// Role 'מפקד מרכז' has read-only access — cannot enter sessions, cannot correct
// results. Just sees aggregated stats across their assigned sites (column K of
// בוחנים). Date range optional (defaults to today). Three categories: overall,
// per-site, per-license.
function handleCenterManagerReport(p) {
  if (!verifyToken(p.examinerId, p.token)) {
    return jsonResponse({ status: 'error', message: 'טוקן לא תקין', tokenExpired: true });
  }
  var role = getExaminerRole(p.examinerId);
  // 'מפקד קד״ץ' shares the same dashboard as 'מפקד מרכז' — both are
  // multi-site read-only commander roles, gated only on the site list in
  // column K. Add new commander roles here to give them the same view.
  if (role !== 'מפקד מרכז' && !isKdtzRole(role)) {
    return jsonResponse({ status: 'error', message: 'פעולה זו זמינה רק למפקד' });
  }
  var managedSites = getExaminerManagedSites(p.examinerId);
  if (!managedSites.length) {
    return jsonResponse({ status: 'error', message: 'לא הוקצו אתרים מנוהלים — פנה למנהל המערכת' });
  }
  // Normalise site names: strip all whitespace + lowercase for forgiving match
  // (handles "ב.ה. 6910" vs "ב.ה.6910" vs " ב.ה. 6910 " — common manual-entry drift).
  function normalizeSiteName(s) {
    return String(s || '').replace(/\s+/g, '').toLowerCase();
  }
  var sitesNormalized = {};
  for (var s = 0; s < managedSites.length; s++) {
    var ns = normalizeSiteName(managedSites[s]);
    if (ns) sitesNormalized[ns] = managedSites[s]; // map normalized → display name
  }
  // Diagnostic counters so the UI can show why a report is empty.
  var dbg = { totalRows: 0, inDateRange: 0, statusCancelled: 0, siteMatched: 0, siteMismatched: 0, distinctSitesSeenInRange: {} };

  // Parse date range. Defaults to today (00:00 today → now).
  var dateFrom, dateTo;
  if (p.dateFrom) {
    dateFrom = new Date(p.dateFrom);
    if (isNaN(dateFrom.getTime())) dateFrom = null;
  }
  if (p.dateTo) {
    dateTo = new Date(p.dateTo);
    if (isNaN(dateTo.getTime())) dateTo = null;
    else dateTo.setHours(23, 59, 59, 999);
  }
  if (!dateFrom) {
    dateFrom = new Date();
    dateFrom.setHours(0, 0, 0, 0);
  }
  if (!dateTo) {
    dateTo = new Date();
    dateTo.setHours(23, 59, 59, 999);
  }

  // Walk תוצאות, filter by site IN managed + date range. Skip 'בוטל' (cancelled DQ).
  // Date-bounded and archive-aware: the range defaults to today, and a range
  // older than the 30-day retention window must still see the archive (B5).
  diagMark('sheet:results-center-report');
  var centerRead = readResultsSince(dateFrom);
  var rows = centerRead.rows;
  diagMark('sheet:results-center-report-done:' + centerRead.mode);
  var overall = { total: 0, passed: 0, failed: 0, dq: 0 };
  var bySite = {};
  var byLicense = {};
  // Per-examinee details — needed to render the same rich report style as
  // the site-manager report (KPIs + pie + weak topics + per-examinee table).
  var results = [];
  var examinerExcl = getExaminerExclusion();   // exclude examiners who self-tested as examinees (name or ת.ז.)
  for (var ri = 1; ri < rows.length; ri++) {
    var r = rows[ri];
    dbg.totalRows++;
    var status = String(r[7] || '').trim();
    // Parse row date (column A is "DD/MM/YYYY HH:mm" — see todayStr())
    var rawDate = r[0];
    var rowDate = null;
    if (rawDate instanceof Date) rowDate = rawDate;
    else if (rawDate) {
      var m = String(rawDate).match(/(\d{1,2})\/(\d{1,2})\/(\d{4})\s+(\d{1,2}):(\d{2})/);
      if (m) rowDate = new Date(+m[3], (+m[2]) - 1, +m[1], +m[4], +m[5]);
    }
    if (!rowDate) continue;
    if (rowDate < dateFrom || rowDate > dateTo) continue;
    dbg.inDateRange++;
    if (status === 'בוטל') { dbg.statusCancelled++; continue; }
    var rowSite = String(r[10] || '').trim();
    if (isTestSite(rowSite)) continue;   // system-test site — exclude from the manager report stats
    if (isExaminerSelfTest(r[2], r[1], examinerExcl)) continue;   // examiner self-testing (name or ת.ז.) — exclude
    // Track every distinct site we see in range so the commander can see
    // exactly what site names appear in the sheet vs what they configured.
    if (rowSite) dbg.distinctSitesSeenInRange[rowSite] = (dbg.distinctSitesSeenInRange[rowSite] || 0) + 1;
    var rowSiteNorm = normalizeSiteName(rowSite);
    var matchedDisplay = sitesNormalized[rowSiteNorm];
    if (!matchedDisplay) { dbg.siteMismatched++; continue; }
    dbg.siteMatched++;
    // Use the configured display name so aggregation is consistent
    var siteKey = matchedDisplay;

    var rowLic = String(r[4] || '').trim() || '-';
    var isDQ = (status === 'פסול');
    var isPassed = (status === 'עבר');
    overall.total++;
    if (isDQ) overall.dq++;
    else if (isPassed) overall.passed++;
    else overall.failed++;

    if (!bySite[siteKey]) bySite[siteKey] = { site: siteKey, total: 0, passed: 0, failed: 0, dq: 0 };
    bySite[siteKey].total++;
    if (isDQ) bySite[siteKey].dq++;
    else if (isPassed) bySite[siteKey].passed++;
    else bySite[siteKey].failed++;

    // Capture per-examinee row for the rich report
    results.push({
      date: r[0],
      idNumber: r[1],
      name: r[2],
      phone: r[3],
      license: r[4],
      score: r[5],
      percent: r[6],
      passed: r[7],
      time: r[8],
      examiner: r[9],
      site: siteKey,
      classroom: r[11],
      language: r[12],
      attempt: r[14],
      wrongDetails: r[15],
      disqualified: r[17],
      population: r[19] || '',
      corrected: r[20] || false,
      audioMode: r[21] || 'off'
    });

    if (!byLicense[rowLic]) byLicense[rowLic] = { license: rowLic, total: 0, passed: 0, failed: 0, dq: 0 };
    byLicense[rowLic].total++;
    if (isDQ) byLicense[rowLic].dq++;
    else if (isPassed) byLicense[rowLic].passed++;
    else byLicense[rowLic].failed++;
  }

  // Ensure every managed site appears in bySite (even with zero rows) so the
  // commander can spot missing data instead of being confused by absence.
  for (var ms = 0; ms < managedSites.length; ms++) {
    var name = managedSites[ms];
    if (isTestSite(name)) continue;   // never surface the system-test site, even as a zero row
    if (!bySite[name]) bySite[name] = { site: name, total: 0, passed: 0, failed: 0, dq: 0 };
  }

  function pct(part, whole) { return whole > 0 ? Math.round((part / whole) * 100) : 0; }
  overall.passRate = pct(overall.passed, overall.total);

  var bySiteArr = [];
  for (var sk in bySite) {
    var bsr = bySite[sk];
    bsr.passRate = pct(bsr.passed, bsr.total);
    bySiteArr.push(bsr);
  }
  bySiteArr.sort(function(a, b) { return a.site.localeCompare(b.site, 'he'); });

  var byLicArr = [];
  for (var lk in byLicense) {
    var blr = byLicense[lk];
    blr.passRate = pct(blr.passed, blr.total);
    byLicArr.push(blr);
  }
  // Sort by typical license order: B, 1, C1, C, D, other
  var licOrder = { 'B': 1, '1': 2, 'C1': 3, 'C': 4, 'D': 5 };
  byLicArr.sort(function(a, b) {
    var oa = licOrder[a.license] || 99, ob = licOrder[b.license] || 99;
    return oa - ob || a.license.localeCompare(b.license);
  });

  // Convert distinct-sites-seen map → sorted array for display
  var seenArr = [];
  for (var ds in dbg.distinctSitesSeenInRange) {
    seenArr.push({ site: ds, count: dbg.distinctSitesSeenInRange[ds] });
  }
  seenArr.sort(function(a, b) { return b.count - a.count; });

  return jsonResponse({
    status: 'ok',
    managedSites: managedSites,
    dateFrom: dateFrom.toISOString(),
    dateTo: dateTo.toISOString(),
    overall: overall,
    bySite: bySiteArr,
    byLicense: byLicArr,
    results: results,
    // Diagnostic info shown when the report is empty — helps identify the
    // cause (wrong site name in column K, no exams in date range, etc.)
    diagnostics: {
      totalRowsInSheet: dbg.totalRows,
      rowsInDateRange: dbg.inDateRange,
      rowsCancelled: dbg.statusCancelled,
      rowsMatchedSite: dbg.siteMatched,
      rowsMismatchedSite: dbg.siteMismatched,
      sitesSeenInRange: seenArr,
      configuredSites: managedSites
    }
  });
}

