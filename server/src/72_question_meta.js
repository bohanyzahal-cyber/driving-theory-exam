// Question metadata (topic / image) for the reports, WITHOUT touching Drive.
//
// Measured in the 'אבחון' sheet on 2026-09-15, the day the diagnostics shipped:
// commanderDashboard ran 33-70s, and the trail showed the Sheets read finishing
// in 0.6s while the remaining 30-60s were Drive reads of the question banks —
// every language, twice, because the two resolver loops each called
// loadQuestionsForLanguageServer with no shared memo (up to 14 reads in one
// request). That is what times the client out at 30s, and on a slow Google
// afternoon it is exactly the shape of the 360s kills.
//
// The per-license pools already hold the same question objects (id, text,
// category, imageUrl) and live in CacheService, so the union of the five pools
// answers both loops. Two guards keep it honest:
//   - all five pools must be present and non-empty; a missing or unreadable one
//     falls back to Drive, so a partial union is never returned.
// MEASURED 2026-09-15 against the real banks (local copies verified identical
// to production by comparing pool sizes with the warmup log): EVERY question of
// a language belongs to at least one license, so five present pools ARE the
// whole bank - he/ar/fr/es/am union 1694 of 1694, ru 1693 of 1693, en 1700 of
// 1700. An earlier version of this guard compared the union against the
// translation index count (1700 = the union ACROSS languages) and so rejected
// every language except English, sending them all back to Drive. Do not
// reintroduce a cross-language total as a per-language expectation.
// The memo makes a language cost at most one resolution per request.
function questionMetaForLanguage(lang, memo) {
  var safeLang;
  try { safeLang = normalizeQuestionCacheLanguage(lang); } catch (eLang) { return []; }
  if (memo && memo[safeLang]) return memo[safeLang];
  var rows = null;
  try {
    var cache = CacheService.getScriptCache();
    var licenses = Object.keys(EXAM_STRUCTURE_SERVER), seen = {}, out = [];
    for (var c = 0; c < licenses.length; c++) {
      var pool = readQuestionCacheRecord(cache, QUESTION_CACHE_PREFIX + 'pool_' + safeLang + '_' + licenses[c], QUESTION_POOL_MAX_PARTS);
      if (!Array.isArray(pool) || !pool.length) { out = null; break; }
      for (var i = 0; i < pool.length; i++) {
        var q = pool[i];
        if (q && q.id && !seen[q.id]) { seen[q.id] = true; out.push(q); }
      }
    }
    if (out && out.length) rows = out;
  } catch (eCache) { rows = null; }
  if (!rows) rows = loadQuestionsForLanguageServer(safeLang);   // marks drive:<lang> itself
  if (memo) memo[safeLang] = rows;
  return rows;
}

// ========== Emergency cache reset ==========
// Run this manually from the Apps Script editor when you see "0 questions"
// (or similar nonsense from a poisoned cache). Clears every cached language
// chunk, then re-warms straight from Drive. Returns a human-readable report.
//
// Real-world trigger: during one exam day, the Hebrew chunks were cached as
// empty after a race condition between two parallel loads. Every subsequent
// examinee got "0 questions" in Hebrew until the cache TTL expired. This
// reset invalidates all derived data; completion time depends on Drive.
//
// Apps Script editor → select function: emergencyClearAndRefreshCache → Run.
// Then check Logger output (View → Logs or "Execution log" panel).
function emergencyClearAndRefreshCache() {
  var cache = CacheService.getScriptCache(), report = [];
  clearQuestionBankCacheKeys(cache);
  clearAllLicensePools(cache, TX_LANGS);
  var keys = [QUESTION_CACHE_PREFIX + 'tx_meta'];
  for (var s = 0; s < QUESTION_TX_SHARDS; s++) keys.push(QUESTION_CACHE_PREFIX + 'tx_' + s);
  cache.removeAll(keys);
  report.push('Cleared pools AND translations (and any legacy bank records); rebuilding from Drive.');
  report = report.concat(warmupQuestionCaches({ resetCursor: true }));
  var out = report.join('\n');
  Logger.log(out);
  return out;
}

// ========== Server-side question delivery ==========
// Loads question data from a private Google Drive folder (one JSON file per
// language) and returns a curated 30-question exam to authenticated clients.
// The full question bank never reaches the browser — only the questions for
// the current exam, without the correct-answer index.
//
// Setup:
//   1) Run deployment/generate_questions_data.js locally to produce
//      deployment/generated/questions_<lang>.json files.
//   2) Upload all 7 files to a private Drive folder (only this account
//      should have access; do NOT share publicly).
//   3) Copy the folder ID (the long string in the Drive URL) and set it
//      as ScriptProperty: QUESTIONS_DRIVE_FOLDER_ID = <folder-id>
//   4) Deploy this Apps Script.
//   5) Push updated HTMLs (examinee, exam, student, find_image) so they call
//      getExamQuestions instead of loading questions.js.

