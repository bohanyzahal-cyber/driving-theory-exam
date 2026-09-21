// Ordered score bins for the last practice attempt before the exam. Finer than
// the 3 practice-impact buckets so the probability curve has resolution near
// the pass threshold (~86% in practice terms).
var PP_BINS = ['0-49', '50-59', '60-69', '70-79', '80-85', '86-92', '93-100'];
function ppBin(pct) {
  if (pct == null || pct < 0) return null;   // unparseable / no score
  if (pct < 50) return '0-49';
  if (pct < 60) return '50-59';
  if (pct < 70) return '60-69';
  if (pct < 80) return '70-79';
  if (pct < 86) return '80-85';
  if (pct < 93) return '86-92';
  return '93-100';
}

// Empirical-Bayes shrinkage: blend the cell's observed rate with a prior
// (parent) rate, weighted by K "virtual" observations of the prior. Returns a
// probability in [0,1]. K controls how much support a cell needs before it
// out-weighs its parent — K=12 means a cell needs ~12 samples to carry half the
// weight. priorRate is already a probability in [0,1].
function ppShrink(passed, n, priorRate, k) {
  return (passed + k * priorRate) / (n + k);
}

// Local, self-contained copies of the join/normalization helpers (the versions
// inside handleCommanderDashboard are private to that function). Kept identical
// on purpose so the model's join matches the dashboard's practice-impact join.
function ppNormName(s) {
  if (!s) return '';
  var t = String(s).trim();
  if (!t) return '';
  t = t.replace(/[׳״'".\-]/g, ' ').replace(/\s+/g, ' ').trim().toLowerCase();
  var tokens = t.split(' ').filter(function(x) { return x; });
  tokens.sort();
  return tokens.join(' ');
}
function ppNormPhone(v) {
  var d = String(v || '').replace(/\D/g, '');
  return d.length >= 9 ? d.slice(-9) : '';
}
function ppParsePct(v) {
  if (typeof v === 'number') return v <= 1 ? v * 100 : v;
  var s = String(v || '').replace('%', '').trim();
  if (!s) return -1;
  var n = parseFloat(s);
  if (isNaN(n)) return -1;
  return n <= 1 ? n * 100 : n;
}

// Given the list of practice records (each {date, percent}) that fall in the
// lookback window before an exam, derive the feature vector the model keys on.
// list must be sorted newest-first. Returns null when there's no usable
// practice (the "no practice" branch is modelled separately).
function ppExtractFeatures(list, examDate, lookbackDays) {
  if (!list || !list.length) return null;
  var windowStart = new Date(examDate.getTime() - lookbackDays * 86400000);
  var inWin = [];
  for (var i = 0; i < list.length; i++) {
    var rec = list[i];
    if (rec.date <= examDate && rec.date >= windowStart && rec.percent >= 0) inWin.push(rec);
  }
  if (!inWin.length) return null;
  // inWin is newest-first. Latest = the attempt closest to the exam.
  var latest = inWin[0];
  var oldest = inWin[inWin.length - 1];
  var best = -1;
  for (var j = 0; j < inWin.length; j++) if (inWin[j].percent > best) best = inWin[j].percent;
  var daysSince = Math.round((examDate.getTime() - latest.date.getTime()) / 86400000);
  // Trend across the window: newest minus oldest. >3pt = improving, <-3 = declining.
  var trendPts = inWin.length > 1 ? (latest.percent - oldest.percent) : 0;
  return {
    lastPct: latest.percent,
    bestPct: best,
    sessions: inWin.length,
    daysSince: daysSince,
    trendPts: trendPts,
    trend: inWin.length < 2 ? 'single' : (trendPts > 3 ? 'up' : (trendPts < -3 ? 'down' : 'flat'))
  };
}

// Build the calibration model from the full history. Reads תוצאות + תוצאות תרגול
// directly so it's independent of handleCommanderDashboard. opts:
//   lookbackDays (default 30) — practice window before each exam
//   sinceDate    (optional Date) — ignore exams before this (bound the history)
function buildPassProbabilityModel(opts) {
  opts = opts || {};
  var lookbackDays = opts.lookbackDays || 30;
  var sinceDate = opts.sinceDate || null;

  // Both reads are column-pruned, and the results read spans live + archive
  // (B5). Exam columns used below: A date, B id, C name, D phone, E licence,
  // H pass, K site, O attempt, R פסול. Practice columns: A date, C name,
  // D class, F licence, I percent, P phone — never N/O, the two JSON blobs
  // that made the practice read 28.5 s (r17).
  // ⚠ Index another column here and it MUST be added to the colSpec; a column
  // outside the list reads as '' instead of failing.
  var resData = readResultsSince(sinceDate, [[1, 5], [8, 1], [11, 1], [15, 1], [18, 1]]).rows;
  var practiceData = readRowsSince(getSheet('תוצאות תרגול'), 0,
    sinceDate ? new Date(sinceDate.getTime() - lookbackDays * 86400000) : null,
    [[1, 1], [3, 2], [6, 1], [9, 1], [16, 1]]).rows;

  // Class → site map (practice rows store the class code, not the site).
  var classSiteMap = {};
  try {
    var classData = getSheet('כיתות').getDataRange().getValues();
    for (var c = 1; c < classData.length; c++) {
      classSiteMap[String(classData[c][0]).trim()] = String(classData[c][7] || '').trim();
    }
  } catch (eC) { /* no כיתות → name+site fallback stays empty */ }

  // Practice indexes (same three keys as the dashboard).
  var byName = {}, byPhone = {}, byNameSite = {};
  for (var pi = 1; pi < practiceData.length; pi++) {
    var pName = ppNormName(practiceData[pi][2]);
    if (!pName) continue;
    var pLic = String(practiceData[pi][5] || '').trim();
    var pDate = practiceData[pi][0];
    if (pDate && !(pDate instanceof Date)) pDate = new Date(pDate);
    if (!pDate || isNaN(pDate.getTime())) continue;
    var rec = { date: pDate, percent: ppParsePct(practiceData[pi][8]) };
    var nk = pName + '|' + pLic;
    (byName[nk] = byName[nk] || []).push(rec);
    var ph = (practiceData[pi].length > 15) ? ppNormPhone(practiceData[pi][15]) : '';
    if (ph) (byPhone[ph] = byPhone[ph] || []).push(rec);
    var site = classSiteMap[String(practiceData[pi][3] || '').trim()] || '';
    if (site) { var nsk = pName + '|' + pLic + '|' + site; (byNameSite[nsk] = byNameSite[nsk] || []).push(rec); }
  }
  function sortDesc(idx) { for (var k in idx) idx[k].sort(function(a, b) { return b.date - a.date; }); }
  sortDesc(byName); sortDesc(byPhone); sortDesc(byNameSite);

  var examinerExcl = getExaminerExclusion();

  // Accumulators. The model keys cells on license × attempt × score-bin — real
  // data shows attempt number is as strong a predictor as license (attempt 1
  // ~45% pass, 2 ~36%, 3+ ~19%) and it's known before the exam, so it's a
  // first-class axis, not just a marginal. Hierarchy for shrinkage:
  // cell(lic|att|bin) → lic|att → lic → base.
  var base = { n: 0, passed: 0 };
  var byLic = {};                 // license → {n, passed}
  var byLicAtt = {};              // 'license|attempt' → {n, passed}  (shrinkage prior)
  var cells = {};                 // 'license|attempt|bin' → {n, passed}
  var noPractice = { _all: { n: 0, passed: 0 } };   // 'license|attempt' → {n,passed}, plus _all
  var byAttempt = {};             // '1' / '2' / '3+' → {n, passed} (marginal diagnostic)
  var bySessions = {};            // '1' / '2' / '3+' → {n, passed} (marginal diagnostic)
  var byTrend = {};               // up/flat/down/single → {n, passed} (marginal diagnostic)
  var coverage = { eligible: 0, matched: 0, byPhone: 0, byNameSite: 0, byName: 0 };

  // Attempt number → bucket. Col 14 holds the attempt index (1,2,3...).
  function attBucket(v) { var n = Number(v) || 1; return n >= 3 ? '3+' : String(n); }

  function bump(obj, key, passed) {
    if (!obj[key]) obj[key] = { n: 0, passed: 0 };
    obj[key].n++; if (passed) obj[key].passed++;
  }

  for (var r = 1; r < resData.length; r++) {
    var rowDate = parseSheetDate(resData[r][0]);
    if (!rowDate) continue;
    if (sinceDate && rowDate < sinceDate) continue;

    var siteName = String(resData[r][10] || '');
    if (isTestSite(siteName)) continue;
    if (isExaminerSelfTest(resData[r][2], resData[r][1], examinerExcl)) continue;

    var passedStr = String(resData[r][7] || '');
    if (passedStr === 'בוטל') continue;
    var isDQ = resData[r][17] === true || String(resData[r][17]).toUpperCase() === 'TRUE' || passedStr === 'פסול';
    if (isDQ) continue;   // DQ ≠ knowledge; excluded from the pass model (matches practice-impact)
    var isPassed = (passedStr === 'עבר') ? 1 : 0;

    var license = String(resData[r][4] || '').trim() || 'לא צוין';
    var att = attBucket(resData[r][14]);
    var licAtt = license + '|' + att;
    base.n++; if (isPassed) base.passed++;
    bump(byLic, license, isPassed);
    bump(byLicAtt, licAtt, isPassed);
    bump(byAttempt, att, isPassed);

    // Join to practice history (phone → name+site → name), same priority as the
    // dashboard. examDate must be a real Date for the window math.
    var examName = ppNormName(resData[r][2]);
    var examDate = resData[r][0];
    if (examDate && !(examDate instanceof Date)) examDate = new Date(examDate);
    if (!examName || !examDate || isNaN(examDate.getTime())) { bump(noPractice, licAtt, isPassed); noPractice._all.n++; if (isPassed) noPractice._all.passed++; continue; }

    coverage.eligible++;
    var examPhone = ppNormPhone(resData[r][3]);
    var examSite = siteName.trim();
    var list = null, matchType = '';
    if (examPhone && byPhone[examPhone]) { list = byPhone[examPhone]; matchType = 'byPhone'; }
    if (!list && examSite && byNameSite[examName + '|' + license + '|' + examSite]) { list = byNameSite[examName + '|' + license + '|' + examSite]; matchType = 'byNameSite'; }
    if (!list && byName[examName + '|' + license]) { list = byName[examName + '|' + license]; matchType = 'byName'; }

    var feat = list ? ppExtractFeatures(list, examDate, lookbackDays) : null;
    if (!feat) {
      bump(noPractice, licAtt, isPassed);
      noPractice._all.n++; if (isPassed) noPractice._all.passed++;
      continue;
    }
    coverage.matched++;
    if (coverage[matchType] != null) coverage[matchType]++;

    var bin = ppBin(feat.lastPct);
    if (bin) bump(cells, licAtt + '|' + bin, isPassed);
    var sessKey = feat.sessions >= 3 ? '3+' : String(feat.sessions);
    bump(bySessions, sessKey, isPassed);
    bump(byTrend, feat.trend, isPassed);
  }

  // Finalize rates (probability in [0,1]) for every accumulator.
  function rate(o) { return o && o.n > 0 ? o.passed / o.n : 0; }
  base.rate = rate(base);
  for (var lk in byLic) byLic[lk].rate = rate(byLic[lk]);
  for (var lak in byLicAtt) byLicAtt[lak].rate = rate(byLicAtt[lak]);
  for (var ck in cells) cells[ck].rate = rate(cells[ck]);
  for (var nk2 in noPractice) noPractice[nk2].rate = rate(noPractice[nk2]);
  for (var ak in byAttempt) byAttempt[ak].rate = rate(byAttempt[ak]);
  for (var sk in bySessions) bySessions[sk].rate = rate(bySessions[sk]);
  for (var tk in byTrend) byTrend[tk].rate = rate(byTrend[tk]);

  // ---- Monotonic (isotonic) enforcement across the score bins ----
  // A higher practice score must never predict a LOWER pass probability. Raw
  // shrunk cell rates violate this in sparse low bins (e.g. C1 60-69 observed
  // 6% while 0-49 observed 11% — pure small-sample noise), which would float
  // mid-scorers above low-scorers in the at-risk ranking. We pool-adjacent-
  // violators (PAV, weighted by cell support) the shrunk probabilities per
  // (license, attempt) so predictions are non-decreasing in the score bin.
  var K = 12;
  function pav(values, weights) {
    var blocks = [];
    for (var j = 0; j < values.length; j++) {
      blocks.push({ v: values[j], w: weights[j], len: 1 });
      while (blocks.length > 1 && blocks[blocks.length - 2].v > blocks[blocks.length - 1].v) {
        var b2 = blocks.pop(), b1 = blocks.pop();
        var w = b1.w + b2.w;
        blocks.push({ v: (b1.v * b1.w + b2.v * b2.w) / (w || 1), w: w, len: b1.len + b2.len });
      }
    }
    var out = [];
    for (var q = 0; q < blocks.length; q++) for (var t = 0; t < blocks[q].len; t++) out.push(blocks[q].v);
    return out;
  }
  var monoCells = {};
  var attList = ['1', '2', '3+'];
  for (var lic2 in byLic) {
    var licRate2 = ppShrink(byLic[lic2].passed, byLic[lic2].n, base.rate, K);
    for (var ai = 0; ai < attList.length; ai++) {
      var att2 = attList[ai];
      var la = byLicAtt[lic2 + '|' + att2];
      var laRate2 = ppShrink(la ? la.passed : 0, la ? la.n : 0, licRate2, K);
      var vals = [], wts = [];
      for (var bi2 = 0; bi2 < PP_BINS.length; bi2++) {
        var cc = cells[lic2 + '|' + att2 + '|' + PP_BINS[bi2]];
        vals.push(ppShrink(cc ? cc.passed : 0, cc ? cc.n : 0, laRate2, K));
        wts.push(cc ? Math.max(cc.n, 1) : 1);
      }
      var mono = pav(vals, wts);
      for (var bi3 = 0; bi3 < PP_BINS.length; bi3++) monoCells[lic2 + '|' + att2 + '|' + PP_BINS[bi3]] = mono[bi3];
    }
  }

  return {
    version: 3,
    k: 12,
    lookbackDays: lookbackDays,
    bins: PP_BINS,
    base: base,
    byLicense: byLic,
    byLicAtt: byLicAtt,
    cells: cells,
    monoCells: monoCells,
    noPractice: noPractice,
    byAttempt: byAttempt,
    bySessions: bySessions,
    byTrend: byTrend,
    coverage: coverage
  };
}

// Predict pass probability for one examinee profile from a built model.
// features: { license, attempt (1/2/3+ — the UPCOMING attempt number; default 1),
//             lastPct (number, or null/undefined if no practice) }
// Shrinkage chain: cell(lic|att|bin) → lic|att → lic → base. Returns
//   { prob (0-100 int), n (support of the most specific cell used),
//     basis ('cell'|'licenseAttempt'|'license'|'base'|'noPractice'), confidence }.
function predictPassProbability(model, features) {
  if (!model || !model.base) return null;
  var k = model.k || 12;
  var license = (features && features.license) ? String(features.license).trim() : '';
  var attNum = features && features.attempt != null ? (Number(features.attempt) || 1) : 1;
  var att = attNum >= 3 ? '3+' : String(attNum);
  var licAtt = license + '|' + att;

  var licNode = license && model.byLicense[license] ? model.byLicense[license] : null;
  // License rate shrunk toward the global base.
  var licRate = ppShrink(licNode ? licNode.passed : 0, licNode ? licNode.n : 0, model.base.rate, k);
  // License+attempt rate shrunk toward the license rate — the working prior.
  var laNode = model.byLicAtt && model.byLicAtt[licAtt] ? model.byLicAtt[licAtt] : null;
  var laRate = ppShrink(laNode ? laNode.passed : 0, laNode ? laNode.n : 0, licRate, k);

  var hasPractice = features && features.lastPct != null && features.lastPct >= 0;
  var prob, n, basis;
  if (!hasPractice) {
    var npLA = model.noPractice[licAtt] || null;
    // no-practice(lic|att) shrunk toward the lic|att overall rate.
    prob = ppShrink(npLA ? npLA.passed : 0, npLA ? npLA.n : 0, laRate, k);
    n = npLA ? npLA.n : 0;
    basis = 'noPractice';
  } else {
    var bin = ppBin(features.lastPct);
    var cell = bin ? model.cells[licAtt + '|' + bin] : null;
    // Prefer the monotonic (isotonic) probability so a higher practice score
    // never predicts a lower pass chance; fall back to the raw shrunk rate on
    // older models that predate monoCells.
    var monoKey = licAtt + '|' + bin;
    if (bin && model.monoCells && model.monoCells[monoKey] != null) {
      prob = model.monoCells[monoKey];
    } else {
      prob = ppShrink(cell ? cell.passed : 0, cell ? cell.n : 0, laRate, k);
    }
    n = cell ? cell.n : 0;
    basis = (cell && cell.n >= k) ? 'cell' : (laNode && laNode.n >= k ? 'licenseAttempt' : (licNode ? 'license' : 'base'));
  }
  // Confidence from the support behind the estimate.
  var confidence = n >= 40 ? 'high' : (n >= 12 ? 'medium' : 'low');
  return { prob: Math.round(prob * 100), n: n, basis: basis, confidence: confidence, licenseBaseRate: Math.round(licRate * 100), licenseAttemptRate: Math.round(laRate * 100) };
}
// handlePredictiveModelPreview removed (review C R14): a diagnostic endpoint no
// client ever called, whose only job was to eyeball the model before it was
// wired into the dashboards — which it now is (at-risk list, examinerForecast).
// It also built the whole model synchronously on a doGet.
