// ========== Commander Dashboard ==========

function parseDateParam(str) {
  if (!str) return null;
  var parts = String(str).split('/');
  if (parts.length !== 3) return null;
  var d = parseInt(parts[0], 10);
  var m = parseInt(parts[1], 10) - 1;
  var y = parseInt(parts[2], 10);
  if (isNaN(d) || isNaN(m) || isNaN(y)) return null;
  return new Date(y, m, d);
}

function parseSheetDate(val) {
  if (val instanceof Date) return val;
  var s = String(val || '');
  var match = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (match) return new Date(parseInt(match[3]), parseInt(match[2]) - 1, parseInt(match[1]));
  var d = new Date(s);
  if (!isNaN(d.getTime())) return d;
  return null;
}

function handleCommanderDashboard(p) {
  // Verify role
  var exSheet = getSheet('בוחנים');
  var exData = exSheet.getDataRange().getValues();
  var role = '';
  for (var i = 1; i < exData.length; i++) {
    if (normalizeId(exData[i][1]) === normalizeId(p.examinerId)) {
      role = String(exData[i][5] || 'בוחן');
      break;
    }
  }
  if (role !== 'מפקד') {
    return jsonResponse({ status: 'error', message: 'אין הרשאת מפקד' });
  }

  // Parse date range
  var dateFrom = parseDateParam(p.dateFrom);
  var dateTo = parseDateParam(p.dateTo);
  if (!dateFrom || !dateTo) {
    return jsonResponse({ status: 'error', message: 'תאריכים לא תקינים' });
  }
  dateTo.setHours(23, 59, 59, 999);

  // Read results
  // r16: both big reads below have a provable lower date bound. This handler
  // looks at a result only when its date is >= prevFrom (the trend window that
  // precedes the requested one), and at a practice row only within 30 days
  // before such a result. So rows older than prevFrom (results) or prevFrom-30d
  // (practice) can never reach the output; readRowsSince stops reading there.
  // One extra day of margin on each covers the two date parsers disagreeing on
  // a boundary row. prevFrom is computed here, before the reads, and reused by
  // the trend logic further down - keep the two in step.
  var DAY_MS = 86400000;
  var prevFrom = (dateFrom && dateTo && dateFrom.getTime && dateTo.getTime)
    ? new Date(dateFrom.getTime() - (dateTo.getTime() - dateFrom.getTime()) - 1) : null;
  diagMark('sheet:results-commander');
  var resSheet = getSheet('תוצאות');
  var resRead = readRowsSince(resSheet, 0, prevFrom ? new Date(prevFrom.getTime() - DAY_MS) : null);
  var resData = resRead.rows;
  diagMark('sheet:results-commander-done:' + resRead.mode);

  // Read practice results too — we'll join real-exam outcomes against the
  // practice history of the same name+license to surface a "did practice
  // before exam predict success?" metric. The student app stores its own
  // "מזהה תלמיד" (not the national ID), so we match only on full name +
  // license. Note that this is best-effort: identical names will collapse.
  // r17/r18: 'תוצאות תרגול' is 107,614 rows — 24x 'תוצאות' — so this read is
  // bounded on BOTH axes. Rows: only back to 31 days before prevFrom (see
  // rowsNeededSince). Columns: exactly the six this loop indexes —
  // date(0)=A, name(2)=C, class(3)=D, license(5)=F, percent(8)=I, phone(15)=P.
  // Everything else, above all the per-row JSON blobs in N ('פירוט שגויות', the
  // full text of every wrong question) and O, is never fetched.
  // ⚠ Index another column here and you MUST add it below — a column outside
  // the list reads as '' rather than failing. A test enforces the pairing.
  diagMark('sheet:practice-commander');
  var practiceSheet = getSheet('תוצאות תרגול');
  var practiceRead = readRowsSince(practiceSheet, 0, prevFrom ? new Date(prevFrom.getTime() - 31 * DAY_MS) : null,
    [[1, 1], [3, 2], [6, 1], [9, 1], [16, 1]]);
  var practiceData = practiceRead.rows;
  diagMark('sheet:practice-commander-done:' + practiceRead.mode);

  // Class → site map (from כיתות) — practice rows store the class code, not the
  // site, so this lets the name+site fallback match scope by base.
  var pClassSiteMap = {};
  try {
    diagMark('sheet:classes-commander');
    var pClassData = getSheet('כיתות').getDataRange().getValues();
    for (var pcs = 1; pcs < pClassData.length; pcs++) {
      pClassSiteMap[String(pClassData[pcs][0]).trim()] = String(pClassData[pcs][7] || '').trim();
    }
  } catch (ePCS) { /* no כיתות sheet → name+site index stays empty, name fallback still works */ }
  // Phone normaliser: digits only, last 9 (so "050-1234567", "0501234567" and
  // "972501234567" all collapse to the same key on both practice and exam sides).
  function normPhoneCmd(v) {
    var d = String(v || '').replace(/\D/g, '');
    return d.length >= 9 ? d.slice(-9) : '';
  }

  // Parse a stay-time string from the 'זמן' column into seconds. The examinee
  // client writes Hebrew-formatted strings like "32 דק' 14 שנ'" (see
  // getElapsedTimeStr in examinee.html). Older rows or other clients may use
  // "MM:SS"; we handle both shapes so the dashboard works across the whole
  // historical dataset.
  //
  // Anything we can't parse returns 0 and the row is dropped from the stay-
  // time aggregate — better to under-count than to poison the median with
  // garbage values (negative durations, hh:mm:ss strings that look like
  // minutes when truncated, etc.).
  function parseStayTimeToSeconds(str) {
    if (!str) return 0;
    var s = String(str).trim();
    if (!s) return 0;
    // Format A — Hebrew: "32 דק' 14 שנ'", optionally with geresh ׳ or U+2019.
    // The regex tolerates extra whitespace and either quote-mark variant.
    var hebMatch = s.match(/(\d+)\s*דק['׳’]?\s*(\d+)\s*שנ['׳’]?/);
    if (hebMatch) {
      var hmins = parseInt(hebMatch[1], 10);
      var hsecs = parseInt(hebMatch[2], 10);
      if (!isNaN(hmins) && !isNaN(hsecs) && hsecs < 60) {
        var ht = hmins * 60 + hsecs;
        if (ht > 0 && ht <= 5400) return ht;
      }
      return 0;
    }
    // Format B — "MM:SS" colon-separated (legacy and other clients).
    var colonMatch = s.match(/^(\d{1,3}):(\d{2})$/);
    if (colonMatch) {
      var cmins = parseInt(colonMatch[1], 10);
      var csecs = parseInt(colonMatch[2], 10);
      if (isNaN(cmins) || isNaN(csecs) || csecs >= 60) return 0;
      var ctotal = cmins * 60 + csecs;
      if (ctotal > 5400) return 0;
      return ctotal;
    }
    return 0;
  }

  // Bucket thresholds for stay-time histograms (seconds).
  //   fast    < 20 min — efficient
  //   normal  20–35 min — typical exam pace
  //   slow    > 35 min — close to the 40-min ceiling
  var STAY_FAST_MAX = 20 * 60;
  var STAY_NORMAL_MAX = 35 * 60;

  // Aggregate. stayTimes tracks duration-in-seconds per result so we can
  // compute avg / median / p10 / p90 / 3-bucket histogram per group.
  var overall = { total: 0, passed: 0, failed: 0, disqualified: 0, stayTimes: [], reattempts: 0 };
  // Previous period of identical length, ending right before dateFrom —
  // powers the ▲▼ trend badges on the KPI cards (this period vs the last one).
  var prevOverall = { total: 0, passed: 0, failed: 0, disqualified: 0, reattempts: 0 };
  var prevWindowMs = dateTo.getTime() - dateFrom.getTime();
  // prevFrom itself is computed above, before the sheet reads, because it
  // bounds how much of תוצאות / תוצאות תרגול readRowsSince has to fetch.
  if (!prevFrom) prevFrom = new Date(dateFrom.getTime() - prevWindowMs - 1);
  // Integrity flags (current window only). Definitions mirror the per-row
  // badges in the examiner results table, so commander totals always match
  // what the examiner sees row-by-row.
  var integrityOverall = { unverified: 0, suspicious: 0, corrected: 0 };
  var integrityByExaminer = {};
  var integrityBySite = {};
  var byExaminer = {};
  var bySite = {};
  var byLicense = {};
  var byPopulation = {};
  var byLanguage = {};       // 'he' / 'ru' / ... → stats. Catches translation issues
                              // (one language failing more than others is a content/HEB-RTL signal).
  var byAttempt = {};        // 'ניסיון 1' / 'ניסיון 2' / 'ניסיון 3+' → stats.
                              // Shows whether re-attempts have higher/lower
                              // pass rate (does the second try go better?).
  var byDevice = {};         // 'טלפון'/'טאבלט'/'מחשב'/'מבחן בכתב'/'לא צוין (ישן)' → stats.
  var byAudio = {};          // '🔊 שמע' / 'רגיל' → stats.
  // Weak-topic aggregation: per graded exam, the license blueprint tells how
  // many questions of each topic were asked; each parsed wrong-block is later
  // resolved (id→category via the question DB) and counted against that.
  var topicAsked = {};
  var topicAskedByLic = {};
  var weakTopicPending = [];
  var byDay = {};            // 'YYYY-MM-DD' → count. Drives the throughput line chart.
  var byHour = {};           // 'dow-hour' (0–6 dow, 0–23 hour) → count. Heatmap data.
  var wrongQuestionCounts = {}; // question text → fail count. Aggregated from
                                 // column 15 (פירוט שגויות) to surface the
                                 // top-N most-missed questions for content review.

  // Language label normalizer — short codes get human names for the UI.
  var LANG_LABELS_SERVER = {
    'he': 'עברית', 'ru': 'רוסית', 'en': 'אנגלית',
    'ar': 'ערבית', 'fr': 'צרפתית', 'es': 'ספרדית', 'am': 'אמהרית'
  };
  function isoDateStr(d) {
    var y = d.getFullYear();
    var m = d.getMonth() + 1;
    var day = d.getDate();
    return y + '-' + (m < 10 ? '0' + m : m) + '-' + (day < 10 ? '0' + day : day);
  }

  // ===== Practice impact index =====
  // Builds a lookup: normalized-name + '|' + license → list of practice rows
  // (sorted by date desc). For each real exam row, we'll find the latest
  // practice within 30 days BEFORE the exam date and bucket the impact.
  function normalizeFullName(s) {
    if (!s) return '';
    var t = String(s).trim();
    if (!t) return '';
    // Strip common punctuation and collapse internal whitespace
    t = t.replace(/[׳״'".\-]/g, ' ').replace(/\s+/g, ' ').trim().toLowerCase();
    // Token-sort so "ישראל ישראלי" and "ישראלי ישראל" hash to the same key
    var tokens = t.split(' ').filter(function(x) { return x; });
    tokens.sort();
    return tokens.join(' ');
  }
  function parsePercentValue(v) {
    if (typeof v === 'number') return v <= 1 ? v * 100 : v;
    var s = String(v || '').replace('%', '').trim();
    if (!s) return -1;
    var n = parseFloat(s);
    if (isNaN(n)) return -1;
    return n <= 1 ? n * 100 : n;
  }
  var practiceIndex = {};         // name|license → [{date, percent}]
  var practicePhoneIndex = {};    // last-9-digits phone → [...]  (exact join, fwd-only)
  var practiceNameSiteIndex = {}; // name|license|site → [...]   (collision-safe fallback)
  for (var pi = 1; pi < practiceData.length; pi++) {
    var pName = normalizeFullName(practiceData[pi][2]);
    if (!pName) continue;
    var pLic = String(practiceData[pi][5] || '').trim();
    var pDate = practiceData[pi][0];
    if (pDate && !(pDate instanceof Date)) pDate = new Date(pDate);
    if (!pDate || isNaN(pDate.getTime())) continue;
    var pPct = parsePercentValue(practiceData[pi][8]);
    var pRec = { date: pDate, percent: pPct };
    var pKey = pName + '|' + pLic;
    if (!practiceIndex[pKey]) practiceIndex[pKey] = [];
    practiceIndex[pKey].push(pRec);
    var pPhone = (practiceData[pi].length > 15) ? normPhoneCmd(practiceData[pi][15]) : '';
    if (pPhone) {
      if (!practicePhoneIndex[pPhone]) practicePhoneIndex[pPhone] = [];
      practicePhoneIndex[pPhone].push(pRec);
    }
    var pSite = pClassSiteMap[String(practiceData[pi][3] || '').trim()] || '';
    if (pSite) {
      var pnsKey = pName + '|' + pLic + '|' + pSite;
      if (!practiceNameSiteIndex[pnsKey]) practiceNameSiteIndex[pnsKey] = [];
      practiceNameSiteIndex[pnsKey].push(pRec);
    }
  }
  // Sort each list newest-first for the "latest before X" linear scan.
  function sortPracticeLists(idx) { for (var k in idx) idx[k].sort(function(a, b) { return b.date - a.date; }); }
  sortPracticeLists(practiceIndex);
  sortPracticeLists(practicePhoneIndex);
  sortPracticeLists(practiceNameSiteIndex);

  // Practice-impact accumulators. 4 buckets — 'none' (no practice in window),
  // 'low' (<70%), 'mid' (70-85.99%), 'high' (≥86%, the pass threshold).
  var practiceImpact = {
    none: { total: 0, passed: 0 },
    low:  { total: 0, passed: 0 },
    mid:  { total: 0, passed: 0 },
    high: { total: 0, passed: 0 },
    withAny: { total: 0, passed: 0 }, // sum of low+mid+high — pre-computed for the simple card
    unparseable: 0 // practice found but % couldn't be parsed; counted under withAny but not bucketed
  };
  // Match-coverage: of the eligible (non-DQ) exam-takers, how many were matched
  // to a practice record and by which key. Lets the UI show honest coverage
  // instead of silently treating unmatched as "didn't practice".
  var piCoverage = { eligible: 0, matched: 0, byPhone: 0, byNameSite: 0, byName: 0 };

  // Examiners who registered as examinees to test the system — exclude from stats (by name OR ת.ז.).
  var examinerExcl = getExaminerExclusion();

  for (var r = 1; r < resData.length; r++) {
    var rowDate = parseSheetDate(resData[r][0]);
    if (!rowDate) continue;
    var inPrevWindow = rowDate >= prevFrom && rowDate < dateFrom;
    if ((rowDate < dateFrom || rowDate > dateTo) && !inPrevWindow) continue;

    var examinerName = String(resData[r][9] || '');
    var siteName = String(resData[r][10] || '');
    if (isTestSite(siteName)) continue;   // system-test site — exclude from ALL commander stats (current + previous window)
    if (isExaminerSelfTest(resData[r][2], resData[r][1], examinerExcl)) continue;   // examiner self-testing (name or ת.ז.) — exclude
    var license = String(resData[r][4] || '');
    var population = String(resData[r][19] || '');
    var passedStr = String(resData[r][7] || '');
    if (passedStr === 'בוטל') continue;
    var isDQ = resData[r][17] === true || String(resData[r][17]).toUpperCase() === 'TRUE' || passedStr === 'פסול';
    var isPassed = !isDQ && (passedStr === 'עבר');

    // Previous-window rows feed ONLY the trend comparison — none of the
    // breakdowns, charts or integrity tallies below.
    if (inPrevWindow) {
      prevOverall.total++;
      if (isDQ) prevOverall.disqualified++;
      else if (isPassed) prevOverall.passed++;
      else prevOverall.failed++;
      if ((Number(resData[r][14]) || 1) > 1) prevOverall.reattempts++;
      continue;
    }

    // Integrity flags — same definitions as the examiner results-table badges:
    // unverified = score not re-verified against the trusted answer key
    // (anything except 'מאומת'/'ידני'), excluding DQ rows and 0/X system-fails;
    // suspicious = exam finished in under 3 minutes; corrected = manually
    // amended result (col תוקן?).
    var integVState = (resData[r].length > 22) ? String(resData[r][22] || '') : '';
    var integSuspicious = (resData[r].length > 23) && String(resData[r][23] || '') === 'חשוד';
    var integCorrected = resData[r][20] === true || String(resData[r][20]).toUpperCase() === 'TRUE';
    var integZeroScore = /^0\//.test(String(resData[r][5] || ''));
    var integUnverified = !isDQ && !integZeroScore && integVState !== 'מאומת' && integVState !== 'ידני';
    if (integUnverified || integSuspicious || integCorrected) {
      var integEx = examinerName || 'לא צוין';
      var integSite = siteName || 'לא צוין';
      if (!integrityByExaminer[integEx]) integrityByExaminer[integEx] = { unverified: 0, suspicious: 0, corrected: 0 };
      if (!integrityBySite[integSite]) integrityBySite[integSite] = { unverified: 0, suspicious: 0, corrected: 0 };
      if (integUnverified) { integrityOverall.unverified++; integrityByExaminer[integEx].unverified++; integrityBySite[integSite].unverified++; }
      if (integSuspicious) { integrityOverall.suspicious++; integrityByExaminer[integEx].suspicious++; integrityBySite[integSite].suspicious++; }
      if (integCorrected) { integrityOverall.corrected++; integrityByExaminer[integEx].corrected++; integrityBySite[integSite].corrected++; }
    }

    // Stay-time: approval → submit. We use the exam's elapsed-time column 8
    // ('זמן', MM:SS) as the proxy — examinee clicks "Start Exam" within a
    // few seconds of approval, and submit happens at the time we record.
    // Adding a separate approval-timestamp column would tighten this but
    // requires a schema change; current proxy is within ~30 seconds.
    var timeSec = parseStayTimeToSeconds(resData[r][8]);

    // Re-attempt detection: column 14 (ניסיון) holds the attempt number for
    // this exam (1, 2, 3...). Anything > 1 is the same examinee taking it
    // again after a previous fail/DQ — useful signal for tracking how many
    // failures actually come back vs walk away.
    var attemptNum = Number(resData[r][14]) || 1;
    var isReattempt = attemptNum > 1;

    // Practice impact — look up most recent practice for this examinee
    // (matched by name+license) within 30 days before the exam date.
    // DQ rows are excluded because they don't reflect knowledge level.
    if (!isDQ) {
      var examineeName = normalizeFullName(resData[r][2]);
      var realLic = String(resData[r][4] || '').trim();
      var realDate = resData[r][0];
      if (realDate && !(realDate instanceof Date)) realDate = new Date(realDate);
      if (examineeName && realDate && !isNaN(realDate.getTime())) {
        var thirtyBefore = new Date(realDate);
        thirtyBefore.setDate(thirtyBefore.getDate() - 30);
        piCoverage.eligible++;
        // Match priority: exact phone → name+license+site → name+license.
        var examPhone = normPhoneCmd(resData[r][3]);
        var examSite = String(resData[r][10] || '').trim();
        var lookupList = null, matchType = '';
        if (examPhone && practicePhoneIndex[examPhone]) {
          lookupList = practicePhoneIndex[examPhone]; matchType = 'byPhone';
        }
        if (!lookupList && examSite && practiceNameSiteIndex[examineeName + '|' + realLic + '|' + examSite]) {
          lookupList = practiceNameSiteIndex[examineeName + '|' + realLic + '|' + examSite]; matchType = 'byNameSite';
        }
        if (!lookupList) {
          lookupList = practiceIndex[examineeName + '|' + realLic] || null;
          if (lookupList) matchType = 'byName';
        }
        lookupList = lookupList || [];
        var matchedPractice = null;
        for (var ml = 0; ml < lookupList.length; ml++) {
          var item = lookupList[ml];
          if (item.date <= realDate && item.date >= thirtyBefore) {
            matchedPractice = item;
            break;
          }
        }
        if (!matchedPractice) {
          practiceImpact.none.total++;
          if (isPassed) practiceImpact.none.passed++;
        } else {
          piCoverage.matched++;
          if (matchType && typeof piCoverage[matchType] === 'number') piCoverage[matchType]++;
          practiceImpact.withAny.total++;
          if (isPassed) practiceImpact.withAny.passed++;
          var pct = matchedPractice.percent;
          if (pct < 0) {
            practiceImpact.unparseable++;
          } else if (pct < 70) {
            practiceImpact.low.total++;
            if (isPassed) practiceImpact.low.passed++;
          } else if (pct < 86) {
            practiceImpact.mid.total++;
            if (isPassed) practiceImpact.mid.passed++;
          } else {
            practiceImpact.high.total++;
            if (isPassed) practiceImpact.high.passed++;
          }
        }
      }
    }

    // Language (col 12) — drives the byLanguage breakdown. Default to Hebrew
    // since that's the source language and missing values pre-date the column.
    var langCode = String(resData[r][12] || 'he').toLowerCase().trim();
    var langName = LANG_LABELS_SERVER[langCode] || langCode;

    overall.total++;
    if (isDQ) overall.disqualified++;
    else if (isPassed) overall.passed++;
    else overall.failed++;
    if (timeSec > 0) overall.stayTimes.push(timeSec);
    if (isReattempt) overall.reattempts++;

    // Time-series — one increment per row, no sub-groups (keeps payload small)
    var dayKey = isoDateStr(rowDate);
    if (!byDay[dayKey]) byDay[dayKey] = { total: 0, passed: 0, failed: 0, dq: 0 };
    byDay[dayKey].total++;
    if (isDQ) byDay[dayKey].dq++;
    else if (isPassed) byDay[dayKey].passed++;
    else byDay[dayKey].failed++;
    var hourKey = rowDate.getDay() + '-' + rowDate.getHours();
    byHour[hourKey] = (byHour[hourKey] || 0) + 1;

    // Wrong-question aggregation. Column 15 (פירוט שגויות) is a multi-line
    // string with one block per missed question:
    //   "שאלה: <text>\nתשובת הנבחן: <ans>\nתשובה נכונה: <correct>\n\n"
    // We split on blank lines and extract the "שאלה:" line as the natural
    // key. Question text is a stable identifier across rows because the same
    // text is rendered for every examinee who got that question wrong.
    var wrongDetails = String(resData[r][15] || '');
    var topicBlocksParsed = 0;
    if (wrongDetails) {
      var blocks = wrongDetails.split(/\n\s*\n/);
      for (var wb = 0; wb < blocks.length; wb++) {
        var lines = blocks[wb].split('\n');
        var qText = '', qCorrect = '', qId = '';
        for (var wl = 0; wl < lines.length; wl++) {
          var line = lines[wl];
          if (line.indexOf('מזהה שאלה:') === 0) {
            qId = line.replace(/^מזהה שאלה:\s*/, '').trim();
          } else if (line.indexOf('שאלה:') === 0) {
            qText = line.replace(/^שאלה:\s*/, '').trim();
            if (qText.length > 200) qText = qText.substring(0, 200);
          } else if (line.indexOf('תשובה נכונה:') === 0) {
            qCorrect = line.replace(/^תשובה נכונה:\s*/, '').trim();
            // Historical data has many "undefined - undefined" entries from the
            // legacy per-language ci bug (memory: project_per_language_ci_bug).
            // Treat those as if no correct answer was captured.
            if (qCorrect.indexOf('undefined') !== -1 || qCorrect === '-' || qCorrect === '') {
              qCorrect = '';
            } else {
              var labelStripMatch = qCorrect.match(/^[A-Za-dא-לА-Г]\s*[-–]\s*(.+)$/);
              if (labelStripMatch) qCorrect = labelStripMatch[1].trim();
            }
          }
        }
        // Preferred aggregation key: question ID (added 2026-06-02). Uniquely
        // identifies the question across all license/language variants — no
        // false collisions, no "מה פירוש התמרור?" lumping.
        // Fallback for legacy rows without ID: (text + correctAnswer), or
        // text alone if correctAnswer is also missing/garbage.
        var key;
        if (qId) {
          key = 'id:' + qId;
        } else if (qText) {
          key = 't:' + qText + '|||' + qCorrect;
        } else {
          continue;
        }
        if (!wrongQuestionCounts[key]) {
          wrongQuestionCounts[key] = { count: 0, text: qText, correctAnswer: qCorrect, questionId: qId, langCounts: {} };
        }
        wrongQuestionCounts[key].count++;
        // Per-language split — a question failing mostly in one non-Hebrew
        // language is a translation-bug signal for the content team.
        wrongQuestionCounts[key].langCounts[langName] = (wrongQuestionCounts[key].langCounts[langName] || 0) + 1;
        // Weak-topic pending item: resolved to a topic after the loop via the
        // question DB (id preferred; text fallback for legacy rows).
        topicBlocksParsed++;
        weakTopicPending.push({ id: qId || '', text: qText || '', license: license, topic: '' });
      }
    }

    // Blueprint-based "asked" totals per topic: count a row's blueprint once
    // per graded digital exam (server-verified, or legacy rows that at least
    // carry parsed wrong-blocks). DQ rows, 0/X system-fails and manual paper
    // entries (no per-question data) stay out of both numerator & denominator.
    if (!isDQ && !integZeroScore && (integVState === 'מאומת' || topicBlocksParsed > 0)) {
      var topicBp = EXAM_STRUCTURE_SERVER[license];
      if (topicBp) {
        if (!topicAskedByLic[license]) topicAskedByLic[license] = {};
        for (var tbk in topicBp) {
          topicAsked[tbk] = (topicAsked[tbk] || 0) + topicBp[tbk];
          topicAskedByLic[license][tbk] = (topicAskedByLic[license][tbk] || 0) + topicBp[tbk];
        }
      }
    }

    var eName = examinerName || 'לא צוין';
    var sName = siteName || 'לא צוין';
    var lName = license || 'לא צוין';
    var pName = population || 'לא צוין';

    // Attempt-bucket label: 1, 2, 3+ (anything ≥ 3 collapses to a single
    // bucket — the long tail is too small to be useful on its own).
    var attemptLabel = attemptNum <= 1 ? 'ניסיון 1'
                       : attemptNum === 2 ? 'ניסיון 2'
                       : 'ניסיון 3+';

    addToGroup(byExaminer, eName, isPassed, isDQ, timeSec);
    addToGroup(bySite, sName, isPassed, isDQ, timeSec);
    addToGroup(byLicense, lName, isPassed, isDQ, timeSec);
    addToGroup(byPopulation, pName, isPassed, isDQ, timeSec);
    addToGroup(byLanguage, langName, isPassed, isDQ, timeSec);
    addToGroup(byAttempt, attemptLabel, isPassed, isDQ, timeSec);

    // Device + audio dimensions. Device (col 30) exists only on new rows —
    // older rows group under 'לא צוין (ישן)'; paper entries show 'מבחן בכתב'.
    var deviceRaw = (resData[r].length > 29) ? String(resData[r][29] || '').trim() : '';
    var deviceLabel = deviceRaw === 'phone' ? 'טלפון'
                      : deviceRaw === 'tablet' ? 'טאבלט'
                      : deviceRaw === 'desktop' ? 'מחשב'
                      : (integVState === 'ידני' ? 'מבחן בכתב' : 'לא צוין (ישן)');
    var audioLabel = String(resData[r][21] || 'off') === 'on' ? '🔊 שמע' : 'רגיל';
    addToGroup(byDevice, deviceLabel, isPassed, isDQ, timeSec);
    addToGroup(byAudio, audioLabel, isPassed, isDQ, timeSec);

    // Cross-tabulation sub-groups
    addToSubGroup(byExaminer, eName, 'byLicense', lName, isPassed, isDQ, timeSec);
    addToSubGroup(byExaminer, eName, 'bySite', sName, isPassed, isDQ, timeSec);
    addToSubGroup(bySite, sName, 'byLicense', lName, isPassed, isDQ, timeSec);
    addToSubGroup(bySite, sName, 'byExaminer', eName, isPassed, isDQ, timeSec);
    addToSubGroup(byLicense, lName, 'bySite', sName, isPassed, isDQ, timeSec);
    addToSubGroup(byLicense, lName, 'byExaminer', eName, isPassed, isDQ, timeSec);
    addToSubGroup(byPopulation, pName, 'byLicense', lName, isPassed, isDQ, timeSec);
    addToSubGroup(byPopulation, pName, 'bySite', sName, isPassed, isDQ, timeSec);
  }

  function addToGroup(map, key, isPassed, isDQ, timeSec) {
    if (!map[key]) map[key] = { total: 0, passed: 0, failed: 0, disqualified: 0, stayTimes: [] };
    map[key].total++;
    if (isDQ) map[key].disqualified++;
    else if (isPassed) map[key].passed++;
    else map[key].failed++;
    if (timeSec > 0) map[key].stayTimes.push(timeSec);
  }

  function addToSubGroup(map, primaryKey, subDim, subKey, isPassed, isDQ, timeSec) {
    if (!map[primaryKey]) return;
    if (!map[primaryKey]._sub) map[primaryKey]._sub = {};
    if (!map[primaryKey]._sub[subDim]) map[primaryKey]._sub[subDim] = {};
    addToGroup(map[primaryKey]._sub[subDim], subKey, isPassed, isDQ, timeSec);
  }

  // Percentile helper. arr is assumed already sorted ascending.
  function percentileSorted(sortedArr, p) {
    if (!sortedArr || sortedArr.length === 0) return 0;
    if (sortedArr.length === 1) return sortedArr[0];
    var rank = (p / 100) * (sortedArr.length - 1);
    var lo = Math.floor(rank);
    var hi = Math.ceil(rank);
    if (lo === hi) return sortedArr[lo];
    var w = rank - lo;
    return Math.round(sortedArr[lo] * (1 - w) + sortedArr[hi] * w);
  }

  function computeStats(obj) {
    var stayAvg = 0, stayMedian = 0, stayP10 = 0, stayP90 = 0;
    var stayFast = 0, stayNormal = 0, staySlow = 0;
    var stayTimes = obj.stayTimes || [];
    if (stayTimes.length > 0) {
      var sum = 0;
      for (var s = 0; s < stayTimes.length; s++) {
        sum += stayTimes[s];
        if (stayTimes[s] < STAY_FAST_MAX) stayFast++;
        else if (stayTimes[s] <= STAY_NORMAL_MAX) stayNormal++;
        else staySlow++;
      }
      stayAvg = Math.round(sum / stayTimes.length);
      var sorted = stayTimes.slice().sort(function(a, b) { return a - b; });
      stayMedian = percentileSorted(sorted, 50);
      stayP10 = percentileSorted(sorted, 10);
      stayP90 = percentileSorted(sorted, 90);
    }
    var passRate = obj.total > 0 ? Math.round((obj.passed / obj.total) * 100) : 0;
    var dqRate = obj.total > 0 ? Math.round((obj.disqualified / obj.total) * 100) : 0;
    return {
      total: obj.total,
      passed: obj.passed,
      failed: obj.failed,
      disqualified: obj.disqualified,
      passRate: passRate,
      dqRate: dqRate,
      // Stay-time metrics (all in seconds; client formats as MM:SS)
      stayAvg: stayAvg,
      stayMedian: stayMedian,
      stayP10: stayP10,
      stayP90: stayP90,
      stayFast: stayFast,
      stayNormal: stayNormal,
      staySlow: staySlow,
      stayCount: stayTimes.length
    };
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

  // Throughput timeline: fill in zero-count days between dateFrom and dateTo
  // so the client gets a contiguous series instead of a sparse one (cleaner
  // chart, no false-impression gaps).
  var timeline = [];
  var cursor = new Date(dateFrom);
  cursor.setHours(0, 0, 0, 0);
  var endDay = new Date(dateTo);
  endDay.setHours(0, 0, 0, 0);
  var safetyLimit = 0;
  while (cursor <= endDay && safetyLimit < 400) {
    var k = isoDateStr(cursor);
    var dayAgg = byDay[k] || { total: 0, passed: 0, failed: 0, dq: 0 };
    // `count` kept so an older client (plain-count polyline) keeps working.
    timeline.push({ date: k, count: dayAgg.total, passed: dayAgg.passed, failed: dayAgg.failed, dq: dayAgg.dq });
    cursor.setDate(cursor.getDate() + 1);
    safetyLimit++;
  }

  // Heatmap: flatten to a 7×24 array of counts (0 = Sunday in JS date.getDay())
  var heatmap = [];
  for (var dow = 0; dow < 7; dow++) {
    var hourRow = [];
    for (var hr = 0; hr < 24; hr++) hourRow.push(byHour[dow + '-' + hr] || 0);
    heatmap.push(hourRow);
  }

  // Re-attempt summary — overall.reattempts already counted in the loop;
  // turn it into a rate so the client can show both raw count and %.
  var overallStats = computeStats(overall);
  overallStats.reattempts = overall.reattempts;
  overallStats.reattemptRate = overall.total > 0
    ? Math.round((overall.reattempts / overall.total) * 100)
    : 0;

  // Previous-period rates for the KPI trend badges.
  prevOverall.passRate = prevOverall.total > 0 ? Math.round((prevOverall.passed / prevOverall.total) * 100) : 0;
  prevOverall.dqRate = prevOverall.total > 0 ? Math.round((prevOverall.disqualified / prevOverall.total) * 100) : 0;
  prevOverall.reattemptRate = prevOverall.total > 0 ? Math.round((prevOverall.reattempts / prevOverall.total) * 100) : 0;

  // ===== Weak topics: resolve pending wrong-blocks to topics =====
  // id is language-independent (all language files carry the Hebrew category),
  // so most items resolve on the first (Hebrew) pass; text-fallback items from
  // legacy rows resolve when their exam language comes up. Cache makes the
  // repeated loads cheap (~300ms warm per language).
  var topicWrong = {};
  diagMark('compute:commander-resolvers');
  var topicWrongByLic = {};
  // Shared across BOTH resolver loops below: without it each language was
  // resolved twice per request (see questionMetaForLanguage).
  var qMetaMemo = {};
  try {
    var WT_LANGS = ['he', 'ru', 'en', 'ar', 'fr', 'es', 'am'];
    for (var wtl = 0; wtl < WT_LANGS.length; wtl++) {
      var wtUnresolved = false;
      for (var wtc = 0; wtc < weakTopicPending.length; wtc++) {
        if (!weakTopicPending[wtc].topic) { wtUnresolved = true; break; }
      }
      if (!wtUnresolved) break;
      var wtQs;
      try { wtQs = questionMetaForLanguage(WT_LANGS[wtl], qMetaMemo); } catch (eWtLoad) { continue; }
      if (!Array.isArray(wtQs) || wtQs.length === 0) continue;
      var wtById = {}, wtByText = {};
      for (var wtq = 0; wtq < wtQs.length; wtq++) {
        var wtRec = wtQs[wtq];
        if (!wtRec) continue;
        if (wtRec.id) wtById[String(wtRec.id)] = wtRec;
        if (wtRec.text) wtByText[String(wtRec.text).substring(0, 200)] = wtRec;
      }
      for (var wtp = 0; wtp < weakTopicPending.length; wtp++) {
        var wtItem = weakTopicPending[wtp];
        if (wtItem.topic) continue;
        var wtFound = wtItem.id ? wtById[wtItem.id] : null;
        if (!wtFound && wtItem.text) wtFound = wtByText[wtItem.text];
        if (wtFound) {
          var wtTopic = classifyCategoryServer(wtFound.category);
          if (wtTopic) wtItem.topic = wtTopic;
        }
      }
    }
  } catch (eWtAll) { /* weak-topic section degrades to empty, dashboard still works */ }
  for (var wtf = 0; wtf < weakTopicPending.length; wtf++) {
    var wtFin = weakTopicPending[wtf];
    if (!wtFin.topic) continue;
    topicWrong[wtFin.topic] = (topicWrong[wtFin.topic] || 0) + 1;
    if (!topicWrongByLic[wtFin.license]) topicWrongByLic[wtFin.license] = {};
    topicWrongByLic[wtFin.license][wtFin.topic] = (topicWrongByLic[wtFin.license][wtFin.topic] || 0) + 1;
  }
  function buildTopicArr(wrongMap, askedMap) {
    var tArr = [];
    for (var tk in askedMap) {
      var tAsked = askedMap[tk] || 0;
      if (!tAsked) continue;
      var tWrong = wrongMap[tk] || 0;
      tArr.push({ topic: tk, wrong: tWrong, asked: tAsked, pct: Math.round((tWrong / tAsked) * 100) });
    }
    tArr.sort(function(a, b) { return b.pct - a.pct; });
    return tArr;
  }
  var weakTopicsOut = { overall: buildTopicArr(topicWrong, topicAsked), byLicense: {} };
  for (var wtLic in topicAskedByLic) {
    weakTopicsOut.byLicense[wtLic] = buildTopicArr(topicWrongByLic[wtLic] || {}, topicAskedByLic[wtLic]);
  }

  // ===== Wait time: registration → exam start (ממתינים col E + col L) =====
  // Approval time isn't stored, so this measures the full soldier experience:
  // registered → waited for approval → pressed Start. Per-site via the row's
  // own site (new rows) or the session's host site (fallback).
  var waitTimesOut = { overall: { avg: 0, median: 0, p90: 0, count: 0 }, bySite: {} };
  try {
    var pendDataW = getSheet('ממתינים').getDataRange().getValues();
    // Include rows archiveOldPendingRows moved out of the live sheet, so a
    // date-range wait-time report stays complete beyond the retention window.
    var archW = getSheetIfExists(PENDING_ARCHIVE_SHEET);
    if (archW && archW.getLastRow() > 1) pendDataW = pendDataW.concat(archW.getDataRange().getValues().slice(1));
    var sessSheetW = getSheet('סשנים');
    var sessDataW = sessSheetW.getDataRange().getValues();
    var sessSiteMapW = {};
    for (var swi = 1; swi < sessDataW.length; swi++) {
      sessSiteMapW[String(sessDataW[swi][0]).trim()] = String(sessDataW[swi][3] || '');
    }
    var waitAll = [];
    var waitBySiteArr = {};
    for (var pwi = 1; pwi < pendDataW.length; pwi++) {
      var regRaw = pendDataW[pwi][4];
      var startRaw = pendDataW[pwi][11];
      if (!regRaw || !startRaw) continue;
      var regD = regRaw instanceof Date ? regRaw : new Date(regRaw);
      var startD = startRaw instanceof Date ? startRaw : new Date(startRaw);
      if (isNaN(regD.getTime()) || isNaN(startD.getTime())) continue;
      if (regD < dateFrom || regD > dateTo) continue;
      var waitSec = Math.round((startD.getTime() - regD.getTime()) / 1000);
      if (waitSec <= 0 || waitSec > 4 * 3600) continue; // clock skew / stuck rows
      waitAll.push(waitSec);
      var waitSite = ((pendDataW[pwi].length > 17 ? String(pendDataW[pwi][17] || '') : '').trim())
        || sessSiteMapW[String(pendDataW[pwi][0]).trim()] || 'לא צוין';
      if (!waitBySiteArr[waitSite]) waitBySiteArr[waitSite] = [];
      waitBySiteArr[waitSite].push(waitSec);
    }
    function waitStatsOf(arr) {
      if (!arr.length) return { avg: 0, median: 0, p90: 0, count: 0 };
      var wSum = 0;
      for (var wsi = 0; wsi < arr.length; wsi++) wSum += arr[wsi];
      var wSorted = arr.slice().sort(function(a, b) { return a - b; });
      return {
        avg: Math.round(wSum / arr.length),
        median: percentileSorted(wSorted, 50),
        p90: percentileSorted(wSorted, 90),
        count: arr.length
      };
    }
    waitTimesOut.overall = waitStatsOf(waitAll);
    for (var wbs in waitBySiteArr) waitTimesOut.bySite[wbs] = waitStatsOf(waitBySiteArr[wbs]);
  } catch (eWait) { /* wait-time card simply stays hidden */ }

  // Top-N most-missed questions, sorted by count descending. Capped at 10 —
  // beyond that the list gets noisy and stops driving decisions.
  // Values are objects {count, text, correctAnswer, questionId}; questionId
  // takes precedence (post-2026-06-02 data), text+correctAnswer is the
  // fallback for legacy rows.
  var topWrong = [];
  var wrongKeys = Object.keys(wrongQuestionCounts);
  wrongKeys.sort(function(a, b) {
    return wrongQuestionCounts[b].count - wrongQuestionCounts[a].count;
  });
  for (var wk = 0; wk < Math.min(wrongKeys.length, 10); wk++) {
    var entry = wrongQuestionCounts[wrongKeys[wk]];
    topWrong.push({
      question: entry.text || '',
      correctAnswer: entry.correctAnswer || '',
      questionId: entry.questionId || '',
      count: entry.count,
      langCounts: entry.langCounts || {}
    });
  }

  // Image lookup: many top-N questions are traffic-sign prompts ("מה פירוש
  // התמרור?") that don't make sense without seeing the sign. Try to resolve
  // each top-wrong entry to its real question record so we can include the
  // imageUrl + id. Hebrew first (most exams); other languages as fallback
  // for entries that didn't resolve.
  //
  // The match key is (text, correctAnswer-in-answers, correctIndex points to
  // that answer). This is strict enough that even if two real questions share
  // the same text, we only attach the image when the correct-answer text also
  // matches the answer-key index — so we either get the right sign or no
  // image. Better silent miss than a wrong picture.
  try {
    if (typeof loadQuestionsForLanguageServer === 'function') {
      var SUPPORTED_LANGS_FOR_IMG = ['he', 'ru', 'en', 'ar', 'fr', 'es', 'am'];
      for (var lgi = 0; lgi < SUPPORTED_LANGS_FOR_IMG.length; lgi++) {
        var stillMissing = false;
        for (var tw0 = 0; tw0 < topWrong.length; tw0++) {
          if (!topWrong[tw0].imageUrl) { stillMissing = true; break; }
        }
        if (!stillMissing) break;
        var langQs;
        try { langQs = questionMetaForLanguage(SUPPORTED_LANGS_FOR_IMG[lgi], qMetaMemo); }
        catch (eLoad) { continue; }
        if (!Array.isArray(langQs) || langQs.length === 0) continue;
        // Index by ID (fast path) and by text (fallback path)
        var qById = {}, qByText = {};
        for (var qIdx = 0; qIdx < langQs.length; qIdx++) {
          var qRec = langQs[qIdx];
          if (!qRec) continue;
          if (qRec.id) qById[String(qRec.id)] = qRec;
          if (qRec.text) {
            if (!qByText[qRec.text]) qByText[qRec.text] = [];
            qByText[qRec.text].push(qRec);
          }
        }
        for (var twi = 0; twi < topWrong.length; twi++) {
          if (topWrong[twi].imageUrl) continue;
          // Path 0: question ID known (new data) — direct lookup, no
          // ambiguity. Best of all paths.
          if (topWrong[twi].questionId) {
            var idMatch = qById[String(topWrong[twi].questionId)];
            if (idMatch && idMatch.imageUrl) {
              topWrong[twi].imageUrl = idMatch.imageUrl;
              continue;
            }
          }
          var candidates = qByText[topWrong[twi].question] || [];
          if (candidates.length === 0) continue;
          // Path A: correct answer known — match it precisely against the
          // candidate's answers array.
          if (topWrong[twi].correctAnswer) {
            for (var ci = 0; ci < candidates.length; ci++) {
              var cand = candidates[ci];
              if (!Array.isArray(cand.answers)) continue;
              if (cand.answers.indexOf(topWrong[twi].correctAnswer) !== -1) {
                if (cand.imageUrl) topWrong[twi].imageUrl = cand.imageUrl;
                if (cand.id && !topWrong[twi].questionId) topWrong[twi].questionId = cand.id;
                break;
              }
            }
          }
          // Path B: only one candidate for this exact text — unambiguous.
          if (!topWrong[twi].imageUrl && candidates.length === 1) {
            if (candidates[0].imageUrl) topWrong[twi].imageUrl = candidates[0].imageUrl;
            if (candidates[0].id && !topWrong[twi].questionId) topWrong[twi].questionId = candidates[0].id;
          }
        }
      }
    }
  } catch (eImg) { /* image resolution best-effort; ignore failures */ }

  // Practice impact — finalize pass rates for each bucket. Pass rate is
  // computed only on the non-DQ sample (DQs were excluded above).
  function finalizePI(bucket) {
    return {
      total: bucket.total,
      passed: bucket.passed,
      passRate: bucket.total > 0 ? Math.round((bucket.passed / bucket.total) * 100) : 0
    };
  }
  var practiceImpactOut = {
    none:        finalizePI(practiceImpact.none),
    withAny:     finalizePI(practiceImpact.withAny),
    low:         finalizePI(practiceImpact.low),
    mid:         finalizePI(practiceImpact.mid),
    high:        finalizePI(practiceImpact.high),
    unparseable: practiceImpact.unparseable,
    lookbackDays: 30,
    coverage: piCoverage
  };

  var result = {
    overall: overallStats,
    byExaminer: computeGroupWithSub(byExaminer),
    bySite: computeGroupWithSub(bySite),
    byLicense: computeGroupWithSub(byLicense),
    byPopulation: computeGroupWithSub(byPopulation),
    byLanguage: computeGroupWithSub(byLanguage),
    byAttempt: computeGroupWithSub(byAttempt),
    timeline: timeline,
    heatmap: heatmap,
    topWrong: topWrong,
    practiceImpact: practiceImpactOut,
    prevOverall: prevOverall,
    integrity: { overall: integrityOverall, byExaminer: integrityByExaminer, bySite: integrityBySite },
    byDevice: computeGroupWithSub(byDevice),
    byAudio: computeGroupWithSub(byAudio),
    weakTopics: weakTopicsOut,
    waitTimes: waitTimesOut
  };

  // 15/09 measured 12.1s between compute:commander-resolvers@32775 and the
  // request's 44932ms total — a quarter of the request after the last mark,
  // unattributed. This one closes the trail: anything left between it and the
  // total is serialisation of the payload plus diagFinish's own sheet append.
  diagMark('compute:commander-payload');
  return jsonResponse({ status: 'ok', data: result });
}

// ============================================================================
// ========== Pass-probability calibration engine (predictive BI) ==========
// ============================================================================
// STEP 1 of the predictive layer. Learns, from HISTORICAL real-exam outcomes
// joined to each examinee's practice history, how a practice profile maps to
// the probability of PASSING the real theory exam.
//
// This is the SINGLE SOURCE OF TRUTH that both commander dashboards will use:
//   - teacher-commander  → per-soldier PREVENTIVE risk score (before the exam)
//   - examiner-commander → cohort pass-rate FORECAST (exam-day planning)
//
// Method: hierarchical shrinkage (empirical-Bayes style). We deliberately do
// NOT train a heavy ML model — with a few thousand rows that would over-fit and
// be impossible to explain to a commander. Instead we compute the observed pass
// rate per (license, last-practice-score bin) cell, and shrink each cell toward
// its license base rate, then toward the global base rate, by sample size.
// Sparse cells fall back gracefully; rich cells keep their specific signal.
// Every estimate ships with its support (n) so the UI can be honest about
// confidence. Same join logic as handleCommanderDashboard (phone → name+site →
// name, latest practice within a lookback window before the exam).

