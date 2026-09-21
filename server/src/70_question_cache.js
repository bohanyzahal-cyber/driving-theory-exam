// ========== Question-cache warmup (for scheduled trigger) ==========
// Cache entries can expire early. Warmup prepares bounded, compressed banks,
// pools and translation shards; its final verification reports whether they
// all survived in the shared cache. Timings depend on Drive/service health.
//
// The run is time-boxed well under Google's 360-second kill and resumes:
// it rebuilds the translation index only when the published one is missing
// or older than four hours, refreshes pool languages from a stored cursor,
// and stops early rather than being killed mid-build. A summary line ending
// in PARTIAL is normal and means the next scheduled run continues from the
// cursor it reports; only ERROR lines need attention.
//
// Setup (one-time): do NOT add a trigger for this function by hand any more.
// Run installWarmupTriggers() once from the editor instead — it replaces any
// hand-made warmupQuestionCaches trigger with an hourly ensureQuestionCachesWarm,
// which verifies the cache and calls this function only when it is needed.
//
// Check the returned cache verification, not only the trigger's completion.
function warmupQuestionCaches(options) {
  var opts = options || {};
  var started = Date.now();
  var budget = opts.budgetMs > 0 ? Math.min(opts.budgetMs, WARMUP_MAX_BUDGET_MS) : WARMUP_BUDGET_MS;
  var deadline = started + budget;
  var memo = { banks: {}, cacheStatus: {} }, summary = [], transientFailures = 0;
  try { diagSweep(summary); } catch (eSweep) { summary.push('diagnostics sweep: skipped (' + (eSweep && eSweep.message ? eSweep.message : eSweep) + ')'); }
  var cache = CacheService.getScriptCache();
  var state = opts.resetCursor === true ? { langIdx: 0 } : readWarmupState();
  var partial = false;
  // Reserves grow from this run's own measurements: a unit is started only
  // while the time left still covers the slowest unit of its kind so far.
  var bankMaxMs = WARMUP_BANK_RESERVE_MS, poolMaxMs = WARMUP_POOL_RESERVE_MS;

  // ---- Language banks + translation index ----------------------------------
  // The index is rebuilt on every run, as it always was. Measured in production
  // 2026-09-12: the build itself is 5.8s, and the run has to read all seven
  // banks for the pools anyway, so refreshing it costs almost nothing and its
  // six-hour TTL is renewed every time. Skipping it while it is "fresh enough"
  // was a false economy that also made the refresh depend on the trigger
  // interval: an index exactly at the threshold was never renewed and expired.
  var txMissing = translationIndexAgeMs(cache) === null;
  var txDue = true;
  // No published index at all is the worst state to be in: every exam start
  // falls back to parsing whole banks. Such a run gets the larger budget, and
  // if even that cannot load all seven banks it still spends what it loaded on
  // pools rather than wasting the reads.
  if (txMissing && !(opts.budgetMs > 0)) {
    deadline = started + WARMUP_MAX_BUDGET_MS;
    summary.push('translation-index: MISSING - this run takes the larger ' + WARMUP_MAX_BUDGET_MS + 'ms budget');
  }
  // A rebuild needs every bank inside one execution, so it is attempted only
  // when the whole phase can be expected to fit. Under a Drive slow enough to
  // make that impossible, a published index that keeps serving beats a run that
  // spends its budget on banks it cannot use.
  var txEstimateMs = TX_LANGS.length * bankMaxMs + WARMUP_TX_RESERVE_MS;
  if (!txMissing && deadline - Date.now() < txEstimateMs) {
    txDue = false;
    partial = true;
    summary.push('translation-index: SKIPPED - a rebuild needs about ' + txEstimateMs +
      'ms and this run has ' + (deadline - Date.now()) + 'ms; the published index keeps serving');
  }
  if (txDue) {
    var loadedAll = true;
    for (var i = 0; i < TX_LANGS.length; i++) {
      if (deadline - Date.now() < bankMaxMs + WARMUP_TX_RESERVE_MS) {
        loadedAll = false; partial = true;
        summary.push('bank loads: PARTIAL - budget too short to reach ' + TX_LANGS[i] +
          '; the published index is left untouched');
        break;
      }
      var lang = TX_LANGS[i], t0 = Date.now();
      try {
        var data = loadQuestionsForLanguageServer(lang, memo);
        var bankMs = Date.now() - t0;
        if (bankMs > bankMaxMs) bankMaxMs = bankMs;
        summary.push(lang + ': loaded ' + data.length + ' questions in ' + bankMs + 'ms');
      } catch (e) {
        if (!(e && e.code === 'question_language_unavailable')) transientFailures++;
        summary.push(lang + ': ERROR - ' + (e && e.message ? e.message : e));
      }
    }
    // A language whose JSON is genuinely absent from Drive is left out of the
    // index (clients fetch it on demand). A transient failure (Drive error,
    // malformed file) must not replace a fuller index that is already
    // published - and neither may a run that ran out of budget mid-load.
    var loadedLangs = Object.keys(memo.banks).length;
    if (loadedAll && loadedLangs > 0 &&
        (transientFailures === 0 || loadedLangs >= publishedTranslationLanguageCount(cache))) {
      // A shard that vanishes immediately after a successful write means the
      // shared cache is full of records nothing reads any more: r1-r3 keys left
      // behind by an upgrade or by a rollback. Reclaiming them costs ~72
      // CacheService round-trips (~300s in production), so it runs only when
      // the index actually failed, and only once - never on a healthy run.
      var tx = null, txError = null;
      try { tx = buildTranslationIndexCache(memo, true); }
      catch (eTx) { txError = eTx; }
      if (txError && deadline - Date.now() > WARMUP_TAIL_RESERVE_MS &&
          sweepLegacyQuestionCachesOnce(cache, memo.banks, summary)) {
        try { tx = buildTranslationIndexCache(memo, true); txError = null; }
        catch (eRetry) { txError = eRetry; }
      }
      if (tx) {
        summary.push('translation-index: ' + tx.count + ' questions; languages=' + tx.langs.join(',') + '; cached=' + tx.cached);
      } else {
        summary.push('translation-index: ERROR - ' + (txError && txError.message ? txError.message : txError));
      }
    } else if (loadedAll) {
      summary.push('translation-index: ERROR - skipped; a language failed transiently and the published index is fuller');
    }
  }

  // ---- Per-license pools, resumed from the stored cursor -------------------
  // Each language costs one Drive read (banks are not cached) plus five pool
  // builds. Whatever does not fit in this run is picked up by the next run from
  // the same cursor, so every language is refreshed well inside the six-hour
  // pool TTL while no single run approaches the kill limit.
  var licenses = Object.keys(EXAM_STRUCTURE_SERVER), langsBuilt = 0;
  // The rotation is computed from where this run started: reading the cursor
  // inside the loop, while the loop itself advances it, walks the languages in
  // a stride that repeats some and never reaches others.
  var startIdx = state.langIdx;
  for (var step = 0; step < TX_LANGS.length; step++) {
    var idx = (startIdx + step) % TX_LANGS.length, code = TX_LANGS[idx];
    var need = (memo.banks[code] ? 0 : bankMaxMs) + poolMaxMs + WARMUP_TAIL_RESERVE_MS;
    if (deadline - Date.now() < need) { partial = true; break; }
    if (!memo.banks[code]) {
      var tBank = Date.now();
      try {
        loadQuestionsForLanguageServer(code, memo);
        var lazyMs = Date.now() - tBank;
        if (lazyMs > bankMaxMs) bankMaxMs = lazyMs;
      } catch (eBank) {
        // An optional language with no JSON in Drive is expected, not a fault:
        // keep ERROR lines meaningful for the post-deploy check.
        var absent = eBank && eBank.code === 'question_language_unavailable';
        summary.push('pools ' + code + (absent ? ': SKIPPED - no bank in Drive' :
          ': ERROR - bank unavailable (' + (eBank && eBank.message ? eBank.message : eBank) + ')'));
        state.langIdx = (idx + 1) % TX_LANGS.length;
        continue;
      }
    }
    var parts = [], langPartial = false;
    for (var c = 0; c < licenses.length; c++) {
      if (deadline - Date.now() < poolMaxMs + WARMUP_TAIL_RESERVE_MS) { langPartial = true; partial = true; break; }
      var tPool = Date.now();
      try {
        var pool = loadLicensePoolServer(code, licenses[c], true, memo);
        var poolMs = Date.now() - tPool;
        if (poolMs > poolMaxMs) poolMaxMs = poolMs;
        parts.push(licenses[c] + '=' + pool.length + '; cached=' + memo.cacheStatus[code + '/' + licenses[c]]);
      } catch (ePool) { parts.push(licenses[c] + '=ERROR(' + (ePool && ePool.message) + ')'); }
    }
    summary.push('pools ' + code + (langPartial ? ' PARTIAL' : '') + ': ' + parts.join(' '));
    // A language cut in half is rebuilt from its first license next run: the
    // cursor advances only past a language whose five pools were all written.
    if (langPartial) break;
    state.langIdx = (idx + 1) % TX_LANGS.length;
    langsBuilt++;
  }
  // Only a run that rebuilt every language counts as a refresh: the hourly
  // ensureQuestionCachesWarm measures the cache's age from this stamp, and a
  // PARTIAL run left some pools at their old age.
  if (!partial && langsBuilt === TX_LANGS.length) state.lastCompleteAt = Date.now();
  writeWarmupState(state);

  summary.push('persistent cache budget: <=' + QUESTION_CACHE_RESERVED_KEYS + ' keys (pools + translation shards; banks are not cached); individual values <81KB');
  if (deadline - Date.now() > WARMUP_VERIFY_RESERVE_MS) {
    try { summary.push('cache verification: ' + JSON.stringify(questionCacheStatus())); }
    catch (eCheck) { summary.push('cache verification: ERROR - ' + (eCheck && eCheck.message)); }
  } else {
    summary.push('cache verification: SKIPPED - out of budget; run questionCacheStatus() from the editor');
  }
  summary.push('warmup ' + (partial ? 'PARTIAL' : 'COMPLETE') + ': ' + langsBuilt + '/' + TX_LANGS.length +
    ' pool languages in ' + (Date.now() - started) + 'ms of a ' + (deadline - started) + 'ms budget; next cursor=' + TX_LANGS[state.langIdx] +
    (partial ? ' (the next scheduled run continues from there)' : ''));
  Logger.log('warmupQuestionCaches complete:\n' + summary.join('\n'));
  return summary;
}

// ========== Keep the question cache warm with nobody touching it ==========
// 2026-09-17, from the operator: "I don't want to run the warmup by hand every
// day, and not several times a day." He had been, because the automatic path
// could not be trusted:
//  - the only warmup trigger was added by hand in the editor ("every 4 hours"),
//    so nothing in the code knew whether it existed or when it ran;
//  - it rebuilt blindly and never LOOKED at the cache, so pools evicted early
//    (CacheService does not promise the full 6h) stayed missing until the next
//    blind run — and on 15/09 a test examinee met "נסה שוב" at exam start.
// This runs every hour. It asks questionCacheStatus() (every pool and shard
// present, one real pool and shard decoded — about a second) and rebuilds only
// when the cache is not ready, or when a refresh is due. Early eviction is
// repaired within the hour, and the one-shot rebuild that a live cache miss
// already requests still covers the minutes in between.
//
// It must not collide with the other scheduled work (operator, same day: "check
// it doesn't overlap in run time with other functions"). Measured in the code:
//  - archiveOldPendingRows, daily in the 01:00 hour, holds the SCRIPT LOCK for
//    its whole run — up to 4.5 minutes. Every pool build claims a lease under
//    that lock with tryLock(200), so a warmup inside that window fails all 35.
//  - rebuildAtRiskCache, daily in the 03:00 hour, is a heavy practice-sheet job.
//  - rebuildMissingQuestionCaches, a one-shot a live cache miss schedules, would
//    fight a warmup for the same pool leases.
// Overlap is avoided by looking at what is RUNNING, not at the clock: skip the
// tick while another job holds the script lock (the archive does, for its whole
// run), while a nightly job's running flag is fresh (rebuildAtRiskCache takes no
// lock, so it raises one), and while a one-shot rebuild is pending. A first
// version instead declared the 01:00 and 03:00 HOURS off-limits. Sweeping the
// timer's minute in the simulation showed why that is wrong: a timer at :58 with
// ±15 min drift lands two consecutive ticks inside one hour, so an hour-wide
// window swallowed two ticks and the cache went cold (margin −182 min at :58).
// The jobs themselves are minutes long; only the minutes are skipped now.
//
// WHEN to refresh is decided by looking ahead, not by a fixed list of hours.
// (Fixed hours with a 5h safety net were tried first; that simulation went cold
// at 05:00 after one dropped run next to a protected hour.) Every tick asks: what
// is the LONGEST I may have to wait before a tick can rebuild again — skipping
// hours that may not start a refresh, assuming Apps Script drops
// WARMUP_ASSUME_MISSED_TICKS of the ticks, and assuming the tick that finally
// does the work lands as late in its hour as the drift allows? If the cache
// could expire within that wait, rebuild now.
//
// The ONLY hour that may not START a scheduled refresh is 08 (operator: "and if
// exams start at 8?"). A rebuild rewrites pools one by one and an examinee whose
// exam starts on the pool being written gets a short wait-and-retry; a 07:50 tick
// drifting +15 min runs at 08:05, where a 07:00-only pre-exam rule no longer
// applies. So the pre-exam refresh may run in the 06:00 or 07:00 hour, the
// look-ahead treats 08 as unable to refresh, and a cache that is actually broken
// is still repaired in 08 (that is worse than a rebuild).
// The nightly jobs get NO hour of their own here: a :58 timer landed three ticks
// in a row inside "their" hours (01:58, 03:12, 03:55) and the cache went cold at
// 05:09 — the third hour-wide rule to fail this simulation. Their running flag
// and the script lock keep the check away for the minutes they actually run.
// 2026-09-19 (review action 5): 09 and 10 joined 08. With 08 alone the age rule
// started a full rebuild at ~10:50 on exam days (07:50 pre-exam refresh + 3 h of
// look-ahead) — inside the results wave, leases and all. 11 stays free on
// purpose: the timer-minute sweep showed that protecting up to 12 goes cold at
// 12:43–13:13 when a :58 timer's 07 tick drifts into 08 and the 06:43 refresh
// has to last until the 12 tick; with 11 free the day's refresh lands at
// ~11:45–12:15 instead, after the morning's exam starts are over.
var WARMUP_NO_SCHEDULED_REFRESH_HOURS = [8, 9, 10];
var WARMUP_EXAM_START_HOURS = [8, 9, 10];
var WARMUP_PREEXAM_HOURS = [6, 7];
var WARMUP_PREEXAM_MIN_AGE_MS = 90 * 60 * 1000;
var WARMUP_JOB_FLAG = 'job_running';
var WARMUP_JOB_FLAG_MS = 6 * 60 * 1000;          // no job outlives the 6-minute kill

// Nightly jobs raise this while they run, so the hourly check stays out of their
// way for exactly as long as they take. A stale flag (a killed job) expires.
function markJobRunning(name, running) {
  try {
    var props = PropertiesService.getScriptProperties(), key = QUESTION_CACHE_PREFIX + WARMUP_JOB_FLAG;
    if (running) props.setProperty(key, JSON.stringify({ name: name, at: Date.now() }));
    else props.deleteProperty(key);
  } catch (e) {}
}
function runningJobName() {
  try {
    var flag = JSON.parse(PropertiesService.getScriptProperties().getProperty(QUESTION_CACHE_PREFIX + WARMUP_JOB_FLAG) || 'null');
    return (flag && flag.at && Date.now() - flag.at < WARMUP_JOB_FLAG_MS) ? String(flag.name || 'job') : '';
  } catch (e) { return ''; }
}
var WARMUP_CACHE_TTL_MS = 6 * 60 * 60 * 1000;    // pools and index are written with 21600s
var WARMUP_EXPIRY_BUFFER_MS = 15 * 60 * 1000;
// An hour timer lands within ~15 min of its minute, but two ticks can drift in
// opposite directions — one 15 min early, the next 15 min late — so the spacing
// can exceed whole hours by 30 min. (15 here left a 9-minute margin in the
// three-day simulation.)
var WARMUP_TICK_DRIFT_MS = 30 * 60 * 1000;
var WARMUP_ASSUME_MISSED_TICKS = 1;              // survive Apps Script skipping a run
var WARMUP_REBUILD_RUNNING_MS = 6 * 60 * 1000;   // a one-shot rebuild cannot outlive the 6-min kill
var WARMUP_ENSURE_FUNCTION = 'ensureQuestionCachesWarm';

// Can a tick in this hour START a scheduled or age-based rebuild?
function warmupCanRefreshInHour(h) {
  return WARMUP_NO_SCHEDULED_REFRESH_HOURS.indexOf(h) < 0;
}
// A tick that lands in the first minutes of the first protected hour is the
// previous hour's tick, drifted late (a 07:50 timer runs at 08:05). It may still
// refresh: the rebuild ends well before the 08:30 wave, and refusing it left the
// cache 9 minutes from expiry at 12:xx in the timer-minute sweep (06 tick
// dropped, 07 tick drifted into 08, nothing allowed until 11). The look-ahead
// keeps counting the whole hour as unable to refresh — that stays conservative.
var WARMUP_DRIFTED_TICK_MINUTES = 15;
function warmupCanRefreshNow(h, m) {
  if (warmupCanRefreshInHour(h)) return true;
  return h === WARMUP_NO_SCHEDULED_REFRESH_HOURS[0] && m < WARMUP_DRIFTED_TICK_MINUTES;
}

// Worst-case wait from a tick in `hour` until a later tick can rebuild: walk
// forward hour by hour, counting only hours that may refresh, until 1 + the
// assumed missed ticks of them have passed. The tick that does the work may land
// anywhere in its hour, so count that whole hour too, plus the timer's drift.
function warmupWorstWaitMs(hour) {
  var needed = 1 + WARMUP_ASSUME_MISSED_TICKS, h = hour, steps = 0;
  while (needed > 0 && steps < 48) {
    h = (h + 1) % 24; steps++;
    if (warmupCanRefreshInHour(h)) needed--;
  }
  return (steps + 1) * 60 * 60 * 1000 + WARMUP_TICK_DRIFT_MS;
}

function ensureQuestionCachesWarm() {
  var t0 = Date.now();
  var hour = Number(Utilities.formatDate(new Date(), 'Asia/Jerusalem', 'H'));
  var minute = new Date().getUTCMinutes();   // Israel is a whole number of hours from UTC
  var job = runningJobName();
  if (job) {
    Logger.log('[ENSURE] skipped: ' + job + ' is running');
    return 'skipped: ' + job + ' is running';
  }
  var probe = LockService.getScriptLock();
  if (!probe.tryLock(100)) {
    Logger.log('[ENSURE] skipped: another job holds the script lock');
    return 'skipped: another job holds the script lock';
  }
  probe.releaseLock();
  try {
    var pending = JSON.parse(PropertiesService.getScriptProperties().getProperty(QUESTION_CACHE_PREFIX + QUESTION_REBUILD_FLAG) || 'null');
    if (pending && pending.at && Date.now() - pending.at < WARMUP_REBUILD_RUNNING_MS) {
      Logger.log('[ENSURE] skipped: a one-shot rebuild is already repairing (' + pending.resource + ')');
      return 'skipped: a one-shot rebuild is already repairing';
    }
  } catch (ePending) {}

  var status = null, statusError = '';
  try { status = questionCacheStatus(); } catch (eStatus) { statusError = (eStatus && eStatus.message) || String(eStatus); }
  var state = readWarmupState();
  var ageMs = typeof state.lastCompleteAt === 'number' ? Date.now() - state.lastCompleteAt : null;
  var reason = '';
  if (!status) reason = 'status check failed' + (statusError ? ' (' + statusError + ')' : '');
  else if (!status.ready) reason = 'cache not ready: ' + status.missingOrMixedKeys + ' missing/mixed keys, ' + status.decodeFailures + ' decode failures';
  else if (ageMs === null) reason = 'no complete warmup on record';   // e.g. the first run after deploy: establish the stamp
  else if (WARMUP_PREEXAM_HOURS.indexOf(hour) >= 0 && ageMs >= WARMUP_PREEXAM_MIN_AGE_MS) reason = 'pre-exam refresh (' + hour + ':00 hour)';
  else if (warmupCanRefreshNow(hour, minute) && ageMs + warmupWorstWaitMs(hour) >= WARMUP_CACHE_TTL_MS - WARMUP_EXPIRY_BUFFER_MS) {
    reason = 'age ' + Math.round(ageMs / 60000) + ' min could outlive the cache before the next safe tick (worst wait ' +
      Math.round(warmupWorstWaitMs(hour) / 60000) + ' min)';
  }
  if (!reason) {
    // Nothing to rebuild — still record killed executions promptly instead of
    // waiting hours for the next full warmup to sweep them.
    try { diagSweep(null); } catch (eSweep) {}
    var quiet = 'warm: ' + status.presentKeys + ' keys ready, last complete warmup ' + Math.round(ageMs / 60000) +
      ' min ago' + (WARMUP_EXAM_START_HOURS.indexOf(hour) >= 0 ? ', exam start hour — no scheduled rebuild' : '') +
      ' (' + (Date.now() - t0) + 'ms)';
    Logger.log('[ENSURE] ' + quiet);
    return quiet;
  }
  Logger.log('[ENSURE] rebuilding — ' + reason);
  var summary = warmupQuestionCaches();
  return 'rebuilt (' + reason + '): ' + summary[summary.length - 1];
}

// Run ONCE from the Apps Script editor. Safe to run again: it always leaves
// exactly one hourly ensureQuestionCachesWarm trigger. It removes hand-made
// warmupQuestionCaches triggers (they would double the work), leaves every
// other trigger alone — including the one-shot rebuildMissingQuestionCaches
// triggers a live cache miss creates — and warms the cache right away.
function installWarmupTriggers() {
  var removed = [], triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    var fn = triggers[i].getHandlerFunction();
    if (fn === 'warmupQuestionCaches' || fn === WARMUP_ENSURE_FUNCTION) {
      ScriptApp.deleteTrigger(triggers[i]);
      removed.push(fn);
    }
  }
  ScriptApp.newTrigger(WARMUP_ENSURE_FUNCTION).timeBased().everyHours(1).create();
  var first = ensureQuestionCachesWarm();
  var msg = 'installWarmupTriggers: removed ' + removed.length + ' old trigger(s) [' + removed.join(', ') +
    ']; created one hourly ' + WARMUP_ENSURE_FUNCTION + '; first check → ' + first;
  Logger.log(msg);
  return msg;
}

// ---- Warmup budget, resume cursor and one-time legacy sweep ---------------
// Apps Script kills an execution at 360 seconds, and a killed run never reaches
// a finally block: every cache lease it held stays locked for its full
// lease TTL, on exactly the resource it failed to publish. That resource
// is then both missing from the cache and unbuildable, so live requests for it
// answer question_cache_busy until the lease expires. Two production runs died
// that way (2026-09-09 and 2026-09-11, both starting 08:29, killed at 361s),
// inside the exam-morning start wave. Hence the budget: the run stops on its
// own terms, releases its leases, records how far it got, and the next
// scheduled run continues from there.
var WARMUP_BUDGET_MS = 240000;       // four minutes of the six-minute ceiling
var WARMUP_MAX_BUDGET_MS = 300000;   // cap on a caller-supplied budget
var WARMUP_BANK_RESERVE_MS = 8000;   // measured Drive read + parse: 5-6s/bank
var WARMUP_POOL_RESERVE_MS = 8000;   // filter + dedupe + gzip + write per pool
var WARMUP_TX_RESERVE_MS = 30000;    // index build: 5.8s measured, 5x margin
var WARMUP_TAIL_RESERVE_MS = 5000;
var WARMUP_VERIFY_RESERVE_MS = 25000;
var WARMUP_STATE_KEY = 'warmup_state';
var WARMUP_LEGACY_KEY = 'warmup_legacy_swept';

function readWarmupState() {
  try {
    var raw = PropertiesService.getScriptProperties().getProperty(QUESTION_CACHE_PREFIX + WARMUP_STATE_KEY);
    var state = raw ? JSON.parse(raw) : null;
    if (state && typeof state.langIdx === 'number' && state.langIdx >= 0 && state.langIdx < TX_LANGS.length) return state;
  } catch (e) {}
  return { langIdx: 0 };
}

function writeWarmupState(state) {
  try {
    PropertiesService.getScriptProperties().setProperty(QUESTION_CACHE_PREFIX + WARMUP_STATE_KEY, JSON.stringify(state));
  } catch (e) { Logger.log('[WARMUP] cursor write failed: ' + (e && e.message)); }
}

// Age of the published translation index, or null when none is published.
function translationIndexAgeMs(cache) {
  try {
    var meta = JSON.parse(cache.get(QUESTION_CACHE_PREFIX + 'tx_meta') || 'null');
    if (!meta || !meta.builtAt) return null;
    var age = Date.now() - meta.builtAt;
    return age > 0 ? age : 0;
  } catch (e) { return null; }
}

// The r1-r3 key sweep is no longer part of a normal run. It deleted ~7,160
// keys in ~72 CacheService round-trips every single time, and measurement in
// production on 2026-09-12 showed it was ~300 of the 361 seconds Google killed
// on 09-09 and 09-11 (banks + all 35 pools + verification are only ~58s).
// Nothing has written those key names since r4 and a CacheService entry lives
// at most six hours, so on a running system there is nothing left to find.
// It is kept as recovery for the one case where it still matters: the shared
// cache is so full of them that the translation index cannot be written. Then
// it runs once, and the persisted flag stops it from ever running again.
// Delete the ScriptProperty qv2_warmup_legacy_swept to re-arm it by hand.
function sweepLegacyQuestionCachesOnce(cache, banks, summary) {
  var props = PropertiesService.getScriptProperties();
  var flag = QUESTION_CACHE_PREFIX + WARMUP_LEGACY_KEY;
  if (props.getProperty(flag) === '1') return false;
  // The per-question legacy key list is derived from the banks, so a partial
  // load would leave most of them behind and waste the single attempt.
  if (Object.keys(banks).length < TX_LANGS.length) return false;
  clearLegacyQuestionCaches(cache, banks);
  // Only a completed sweep sets the flag: a run killed mid-sweep holds no
  // lease and the next one simply tries again.
  props.setProperty(flag, '1');
  summary.push('legacy cleanup: one-time sweep of r1-r3 keys, to make room for the index');
  return true;
}


// Editor helper: rebuild ONLY the translation index and report how long its
// phases really take. Run it from the editor in a quiet window to measure the
// heaviest unit in the system, and to leave a fresh index before an exam
// morning. It never touches the per-license pools.
function warmupTranslationIndexOnly() {
  var t0 = Date.now(), memo = { banks: {}, cacheStatus: {} }, summary = [];
  for (var i = 0; i < TX_LANGS.length; i++) {
    var lang = TX_LANGS[i], tb = Date.now();
    try {
      var rows = loadQuestionsForLanguageServer(lang, memo);
      summary.push(lang + ': loaded ' + rows.length + ' questions in ' + (Date.now() - tb) + 'ms');
    } catch (e) {
      summary.push(lang + ': ERROR - ' + (e && e.message ? e.message : e));
    }
  }
  var loaded = Object.keys(memo.banks).length, banksMs = Date.now() - t0, indexMs = 0;
  if (loaded < TX_LANGS.length) {
    summary.push('translation-index: ABORTED - only ' + loaded + '/' + TX_LANGS.length +
      ' banks loaded; the published index is left untouched');
  } else {
    var tIndex = Date.now();
    try {
      var tx = buildTranslationIndexCache(memo, true);
      indexMs = Date.now() - tIndex;
      summary.push('translation-index: ' + tx.count + ' questions; languages=' + tx.langs.join(',') +
        '; cached=' + tx.cached);
    } catch (eTx) {
      indexMs = Date.now() - tIndex;
      summary.push('translation-index: ERROR - ' + (eTx && eTx.message));
    }
  }
  summary.push('phase timings: banks ' + banksMs + 'ms, index build ' + indexMs + 'ms, total ' +
    (Date.now() - t0) + 'ms (Google kills an execution at 360000ms)');
  Logger.log('warmupTranslationIndexOnly:' + '\n' + summary.join('\n'));
  return summary;
}

// ========== r10: hot-path hardening (2026-09-15) ==========
// Evidence from the Executions log of 08/09, 10/09 and 14/09: on every exam
// day a handful of doGet executions ran the full 360 seconds until Google
// killed them, each pinned to a Google service call (Drive/Sheets/Cache have
// no per-call timeout) while holding an execution slot; a killed pool builder
// also left its lease locked for 370s, freezing that license for six more
// minutes. Google's slowness is the trigger; the exposure below is ours.

// A live builder that genuinely needs more than this is slower than the
// slowest Drive read seen under degradation (100s); after it a second builder
// may start, which costs one duplicate read instead of a six-minute freeze.
var QUESTION_CACHE_LEASE_MS = 150000;

// Never read Drive inside an examinee request. On a live pool miss every
// caller answers question_cache_busy (the client retries in 3s) and ONE
// one-shot trigger rebuilds whatever is missing out of band. The flag dedupes
// a start wave; Apps Script allows 20 triggers per script, so creation is
// gated and the rebuild deletes its own triggers when it runs.
var QUESTION_REBUILD_FLAG = 'rebuild_pending';
var QUESTION_REBUILD_FLAG_MS = 180000;
var QUESTION_REBUILD_FUNCTION = 'rebuildMissingQuestionCaches';

function requestQuestionCacheRebuild(resource) {
  try {
    var props = PropertiesService.getScriptProperties();
    var flagKey = QUESTION_CACHE_PREFIX + QUESTION_REBUILD_FLAG;
    var raw = props.getProperty(flagKey), pending = null;
    try { pending = raw ? JSON.parse(raw) : null; } catch (e) {}
    if (pending && pending.at && Date.now() - pending.at < QUESTION_REBUILD_FLAG_MS) return true;
    var lock = LockService.getScriptLock();
    if (!lock.tryLock(200)) return true; // someone else is scheduling it right now
    try {
      raw = props.getProperty(flagKey);
      try { pending = raw ? JSON.parse(raw) : null; } catch (e2) { pending = null; }
      if (pending && pending.at && Date.now() - pending.at < QUESTION_REBUILD_FLAG_MS) return true;
      ScriptApp.newTrigger(QUESTION_REBUILD_FUNCTION).timeBased().after(1000).create();
      props.setProperty(flagKey, JSON.stringify({ at: Date.now(), resource: String(resource || '') }));
      Logger.log('[POOL] live MISS on ' + resource + ': rebuild scheduled out of band');
      return true;
    } finally {
      lock.releaseLock();
    }
  } catch (e) {
    // No trigger (quota, permission) means the caller must build inline as before.
    Logger.log('[POOL] rebuild scheduling failed, falling back to inline build: ' + (e && e.message ? e.message : e));
    return false;
  }
}

// Trigger target: rebuild only what is missing, inside the warmup's budget,
// then remove every one-shot trigger pointing here and clear the flag.
function rebuildMissingQuestionCaches() {
  var started = Date.now(), summary = [], memo = { banks: {}, cacheStatus: {} };
  try {
    var triggers = ScriptApp.getProjectTriggers();
    for (var t = 0; t < triggers.length; t++) {
      if (triggers[t].getHandlerFunction() === QUESTION_REBUILD_FUNCTION) ScriptApp.deleteTrigger(triggers[t]);
    }
  } catch (eTrig) { summary.push('trigger cleanup: ERROR - ' + (eTrig && eTrig.message)); }
  var cache = CacheService.getScriptCache(), licenses = Object.keys(EXAM_STRUCTURE_SERVER), rebuilt = 0;
  try {
    if (translationIndexAgeMs(cache) === null) {
      for (var i = 0; i < TX_LANGS.length; i++) {
        if (Date.now() - started > WARMUP_BUDGET_MS - WARMUP_TX_RESERVE_MS) break;
        try { loadQuestionsForLanguageServer(TX_LANGS[i], memo); } catch (eLang) {}
      }
      if (Object.keys(memo.banks).length === TX_LANGS.length) {
        try { var tx = buildTranslationIndexCache(memo, true); summary.push('translation-index rebuilt: ' + tx.count); rebuilt++; }
        catch (eTx) { summary.push('translation-index: ERROR - ' + (eTx && eTx.message)); }
      } else {
        summary.push('translation-index: left missing, banks incomplete within budget');
      }
    }
    for (var l = 0; l < TX_LANGS.length; l++) {
      for (var c = 0; c < licenses.length; c++) {
        if (Date.now() - started > WARMUP_BUDGET_MS - WARMUP_POOL_RESERVE_MS) { summary.push('budget reached'); l = TX_LANGS.length; break; }
        var key = QUESTION_CACHE_PREFIX + 'pool_' + TX_LANGS[l] + '_' + licenses[c];
        var hit = readQuestionCacheRecord(cache, key, QUESTION_POOL_MAX_PARTS);
        if (Array.isArray(hit) && hit.length) continue;
        try { loadLicensePoolServer(TX_LANGS[l], licenses[c], true, memo); rebuilt++; }
        catch (ePool) { summary.push('pool ' + TX_LANGS[l] + '/' + licenses[c] + ': ERROR - ' + (ePool && ePool.message)); }
      }
    }
  } finally {
    try { PropertiesService.getScriptProperties().deleteProperty(QUESTION_CACHE_PREFIX + QUESTION_REBUILD_FLAG); } catch (eFlag) {}
  }
  summary.push('rebuilt ' + rebuilt + ' missing cache records in ' + (Date.now() - started) + 'ms');
  Logger.log('rebuildMissingQuestionCaches:\n' + summary.join('\n'));
  return summary;
}

// Mid-exam language switches (getQuestionsByIds) used to read the whole bank
// from Drive on every call. The per-license pools already hold every question
// of that language in full; answer from them when they cover the request.
function questionsFromCachedPools(lang, ids) {
  var cache = CacheService.getScriptCache(), licenses = Object.keys(EXAM_STRUCTURE_SERVER), byId = {}, missing = ids.length;
  for (var c = 0; c < licenses.length && missing > 0; c++) {
    var pool = readQuestionCacheRecord(cache, QUESTION_CACHE_PREFIX + 'pool_' + lang + '_' + licenses[c], QUESTION_POOL_MAX_PARTS);
    if (!Array.isArray(pool)) continue;
    for (var i = 0; i < pool.length; i++) {
      var q = pool[i];
      if (q && q.id && !byId[q.id] && ids.indexOf(Number(q.id)) !== -1) { byId[q.id] = q; missing--; }
    }
  }
  return missing === 0 ? byId : null;
}

