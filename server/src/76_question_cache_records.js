// Read a compressed cached bank, or let one execution rebuild it from Drive.
// A request-local memo can reuse the loaded bank within a warmup operation.
// CacheService limits: 100 KB/value, 1,000 items shared by the whole script.
// gzip + base64 makes chunk length equal to its byte count (ASCII). Fixed slots
// bound persistent storage: 7*(16+1) banks + 35*(4+1) pools + 128 shards + 1
// manifest = 423 keys, leaving >500 for rate limits and active examinees.
// Cache is evictable; generation tags prevent mixing old/new chunks on refresh.
var QUESTION_CACHE_PREFIX = 'qv2_';
var QUESTION_CACHE_PART_BYTES = 80000;
var QUESTION_BANK_MAX_PARTS = 16;
var QUESTION_POOL_MAX_PARTS = 4;
var QUESTION_TX_SHARDS = 128;
// 35 pools x (manifest + 4 parts) + 128 translation shards + 1 manifest.
var QUESTION_CACHE_RESERVED_KEYS = 304;

function questionCacheBusy() {
  var err = new Error('מאגר השאלות מתעדכן כעת. יש לנסות שוב בעוד מספר שניות.');
  err.code = 'question_cache_busy';
  err.retryable = true;
  err.waitSec = 3;
  return err;
}

function encodeQuestionCache(value) {
  return Utilities.base64Encode(Utilities.gzip(Utilities.newBlob(JSON.stringify(value), 'application/json')).getBytes());
}

function decodeQuestionCache(encoded) {
  // Utilities.ungzip refuses a Blob without a content type ("Blob object must
  // have non-null content type for this operation"). Without the explicit type
  // every cache read failed in production (r1-r4) and every request rebuilt
  // its pool from Drive while the shared cache looked healthy.
  var gz = Utilities.newBlob(Utilities.base64Decode(encoded), 'application/x-gzip', 'cache.gz');
  return JSON.parse(Utilities.ungzip(gz).getDataAsString('UTF-8'));
}

function readQuestionCacheRecord(cache, key, maxParts) {
  try {
    var raw = cache.get(key + '_meta');
    if (!raw) return null;
    var meta = JSON.parse(raw);
    if (!meta || !meta.g || !Number.isInteger(meta.n) || meta.n < 1 || meta.n > maxParts) return null;
    var keys = [];
    for (var i = 0; i < meta.n; i++) keys.push(key + '_' + i);
    var values = cache.getAll(keys), joined = '', prefix = meta.g + ':';
    for (var j = 0; j < keys.length; j++) {
      var part = values[keys[j]];
      if (typeof part !== 'string' || part.indexOf(prefix) !== 0) return null;
      joined += part.substring(prefix.length);
    }
    return decodeQuestionCache(joined);
  } catch (e) {
    Logger.log('[CACHE] invalid/read failed ' + key + ': ' + (e && e.message ? e.message : e));
    return null;
  }
}

function writeQuestionCacheRecord(cache, key, value, maxParts) {
  try {
    var encoded = encodeQuestionCache(value);
    var n = Math.ceil(encoded.length / QUESTION_CACHE_PART_BYTES);
    if (n < 1 || n > maxParts) throw new Error('compressed record exceeds reserved cache budget (' + n + '/' + maxParts + ' parts)');
    var generation = Utilities.getUuid(), values = {}, keys = [];
    for (var i = 0; i < n; i++) {
      var partKey = key + '_' + i;
      keys.push(partKey);
      values[partKey] = generation + ':' + encoded.substring(i * QUESTION_CACHE_PART_BYTES, (i + 1) * QUESTION_CACHE_PART_BYTES);
    }
    cache.putAll(values, 21600);
    var check = cache.getAll(keys), joined = '';
    for (var j = 0; j < keys.length; j++) {
      if (check[keys[j]] !== values[keys[j]]) throw new Error('chunk missing immediately after write');
      joined += check[keys[j]].substring(generation.length + 1);
    }
    // Round-trip the read-back bytes through the real decoder. A string-only
    // comparison passed for four releases while every decode was failing.
    var decoded = decodeQuestionCache(joined);
    var expectedLength = Array.isArray(value) ? value.length : Object.keys(value).length;
    var decodedLength = Array.isArray(decoded) ? decoded.length : Object.keys(decoded).length;
    if (decodedLength !== expectedLength) throw new Error('read-back decode mismatch (' + decodedLength + '/' + expectedLength + ')');
    // Publish only after all chunks were verified. Readers reject mixed generations.
    var meta = JSON.stringify({ g: generation, n: n });
    cache.put(key + '_meta', meta, 21600);
    if (cache.get(key + '_meta') !== meta) throw new Error('manifest missing immediately after write');
    var obsolete = [];
    for (var k = n; k < maxParts; k++) obsolete.push(key + '_' + k);
    if (obsolete.length) cache.removeAll(obsolete);
    return true;
  } catch (e) {
    Logger.log('[CACHE] WRITE FAILED ' + key + ': ' + (e && e.message ? e.message : e));
    return false;
  }
}

function clearQuestionCacheRecord(cache, key, maxParts) {
  var keys = [key + '_meta'];
  for (var i = 0; i < maxParts; i++) keys.push(key + '_' + i);
  cache.removeAll(keys);
  return keys.length;
}

// Durable lease ownership is claimed under a short true mutex. The mutex is
// released BEFORE Drive access, gzip or filtering. Waiters return a retryable
// response, instead of occupying execution slots with 10-20 second sleeps.
// The lease outlives the six-minute Apps Script execution limit after a crash.
function claimQuestionCacheLease(resource) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(200)) throw questionCacheBusy();
  try {
    var props = PropertiesService.getScriptProperties();
    var key = QUESTION_CACHE_PREFIX + 'lease_' + resource;
    var prior = props.getProperty(key), lease = null;
    try { lease = prior ? JSON.parse(prior) : null; } catch (e) {}
    if (lease && lease.until > Date.now()) throw questionCacheBusy();
    var owner = Utilities.getUuid();
    props.setProperty(key, JSON.stringify({ owner: owner, until: Date.now() + QUESTION_CACHE_LEASE_MS }));
    return { key: key, owner: owner };
  } finally {
    lock.releaseLock();
  }
}

function releaseQuestionCacheLease(lease) {
  if (!lease) return;
  var lock = LockService.getScriptLock();
  var locked = lock.tryLock(200);
  // A busy mutex must NOT leave the lease in place: a start wave keeps the
  // mutex hot exactly when the builder finishes, and an unreleased lease
  // makes every later miss on this resource "busy" for the whole lease TTL.
  // Owner comparison is sufficient without the mutex: while this lease is
  // unexpired nobody else can claim the key, and after expiry a replacement
  // carries a different owner, so the delete below cannot remove it.
  try {
    var props = PropertiesService.getScriptProperties();
    var raw = props.getProperty(lease.key);
    var current = raw ? JSON.parse(raw) : null;
    if (current && current.owner === lease.owner) props.deleteProperty(lease.key);
    if (!locked) Logger.log('[CACHE] lease released without mutex: ' + lease.key);
  } catch (e) {
    Logger.log('[CACHE] lease release failed: ' + (e && e.message ? e.message : e));
  } finally {
    if (locked) lock.releaseLock();
  }
}

function normalizeQuestionCacheLanguage(lang) {
  var safeLang = String(lang || 'he').toLowerCase();
  if (TX_LANGS.indexOf(safeLang) === -1) throw new Error('Invalid language code');
  return safeLang;
}

// Read-only post-warmup verification. It returns counts only, never question
// text, IDs, tokens, folder IDs or student data. All required chunks must still
// exist AFTER banks, pools and translations have shared the cache capacity.
function questionCacheStatus() {
  var cache = CacheService.getScriptCache(), records = [], licenses = Object.keys(EXAM_STRUCTURE_SERVER);
  for (var l = 0; l < TX_LANGS.length; l++) {
    for (var c = 0; c < licenses.length; c++) records.push({ key: QUESTION_CACHE_PREFIX + 'pool_' + TX_LANGS[l] + '_' + licenses[c], max: QUESTION_POOL_MAX_PARTS });
  }
  var metaKeys = records.map(function(r) { return r.key + '_meta'; });
  var txKey = QUESTION_CACHE_PREFIX + 'tx_meta';
  metaKeys.push(txKey);
  var metas = cache.getAll(metaKeys), required = {}, missing = 0, present = 0;
  for (var i = 0; i < records.length; i++) {
    var record = records[i], meta = null;
    try { meta = JSON.parse(metas[record.key + '_meta'] || 'null'); } catch (e) {}
    if (!meta || !meta.g || !Number.isInteger(meta.n) || meta.n < 1 || meta.n > record.max) { missing++; continue; }
    present++;
    for (var p = 0; p < meta.n; p++) required[record.key + '_' + p] = meta.g + ':';
  }
  var tx = null;
  try { tx = JSON.parse(metas[txKey] || 'null'); } catch (eTx) {}
  if (!tx || !tx.g || !Array.isArray(tx.langs) || !tx.langs.length) missing++;
  else {
    present++;
    for (var s = 0; s < QUESTION_TX_SHARDS; s++) required[QUESTION_CACHE_PREFIX + 'tx_' + s] = tx.g + ':';
  }
  var keys = Object.keys(required), maxBytes = 0, decodeChecks = 0, decodeFailures = 0;
  // Decode a real pool record and a real shard, not only their prefixes.
  try {
    var probePool = readQuestionCacheRecord(cache, QUESTION_CACHE_PREFIX + 'pool_he_B', QUESTION_POOL_MAX_PARTS);
    decodeChecks++; if (!Array.isArray(probePool) || !probePool.length) decodeFailures++;
    var probeTx = tryTranslationsFromIndex(probePool && probePool.length ? [probePool[0].id] : [1], false);
    decodeChecks++; if (!probeTx) decodeFailures++;
  } catch (eProbe) { decodeFailures++; }
  for (var start = 0; start < keys.length; start += 50) {
    var batch = keys.slice(start, start + 50), values = cache.getAll(batch);
    for (var k = 0; k < batch.length; k++) {
      var value = values[batch[k]];
      if (typeof value !== 'string' || value.indexOf(required[batch[k]]) !== 0) missing++;
      else { present++; maxBytes = Math.max(maxBytes, value.length); }
    }
  }
  return { ready: missing === 0 && decodeFailures === 0, presentKeys: present, missingOrMixedKeys: missing, maxValueBytes: maxBytes,
    decodeChecks: decodeChecks, decodeFailures: decodeFailures,
    translationLanguages: tx && Array.isArray(tx.langs) ? tx.langs.length : 0, reservedKeyLimit: QUESTION_CACHE_RESERVED_KEYS };
}

// Removal does not depend on old metadata surviving eviction. Old code used up
// to ~35 chunks/bank and ~12/pool; removing 128 fixed legacy slots also covers
// larger past banks. Legacy per-question keys are removed from all loaded IDs.
function clearLegacyQuestionCaches(cache, banks) {
  var keys = ['tx_meta'];
  for (var l = 0; l < TX_LANGS.length; l++) {
    var lang = TX_LANGS[l];
    keys.push('qdata_' + lang + '_meta', 'qload_lock_' + lang);
    for (var p = 0; p < 128; p++) keys.push('qdata_' + lang + '_part_' + p);
    var licenses = Object.keys(EXAM_STRUCTURE_SERVER);
    for (var c = 0; c < licenses.length; c++) {
      var base = 'qpool_' + lang + '_' + licenses[c];
      keys.push(base + '_meta', 'qpool_lock_' + lang + '_' + licenses[c]);
      for (var k = 0; k < 128; k++) keys.push(base + '_part_' + k);
    }
  }
  var seen = {};
  for (var code in banks) {
    var rows = banks[code] || [];
    for (var q = 0; q < rows.length; q++) if (rows[q] && rows[q].id !== undefined) seen[String(rows[q].id)] = true;
  }
  Object.keys(seen).forEach(function(id) { keys.push('tx_' + id); });
  for (var i = 0; i < keys.length; i += 100) cache.removeAll(keys.slice(i, i + 100));
  var bankKeys = clearQuestionBankCacheKeys(cache);
  Logger.log('[CACHE] removed legacy bank/pool keys, ' + bankKeys + ' r1-r3 bank record keys and ' + Object.keys(seen).length + ' legacy translation keys');
}
// Full language banks are NOT cached in CacheService any more (r4). Measured in
// production: reading a bank back from the cache (16 x 80KB base64 chunks,
// base64-decode, gunzip, JSON.parse of ~2MB) took 5-6s, the same as reading the
// JSON from Drive, while the seven banks occupied ~40% of the shared cache and
// pushed the pools/translation shards that live requests actually need out of
// it. Banks are read from Drive and memoised per request; only the derived
// per-license pools and the packed translation index are cached.
function loadQuestionsForLanguageServer(lang, memo) {
  var safeLang = normalizeQuestionCacheLanguage(lang);
  if (memo && memo.banks && memo.banks[safeLang]) return memo.banks[safeLang];
  var t0 = Date.now();
  diagMark('drive:' + safeLang);
  var folderId = PropertiesService.getScriptProperties().getProperty('QUESTIONS_DRIVE_FOLDER_ID');
  if (!folderId) throw new Error('QUESTIONS_DRIVE_FOLDER_ID not configured in ScriptProperties');
  var folder = DriveApp.getFolderById(folderId);
  var fileName = 'questions_' + safeLang + '.json';
  var files = folder.getFilesByName(fileName);
  if (!files.hasNext()) {
    var missing = new Error(fileName + ' not found in Drive folder');
    missing.code = 'question_language_unavailable';
    throw missing;
  }
  var parsed = JSON.parse(files.next().getBlob().getDataAsString('UTF-8'));
  if (!Array.isArray(parsed) || !parsed.length) throw new Error(fileName + ' is empty or invalid');
  Logger.log('[BANK] Drive read ' + safeLang + ': ' + parsed.length + ' rows; ' + (Date.now() - t0) + 'ms');
  if (memo) {
    if (!memo.banks) memo.banks = {};
    memo.banks[safeLang] = parsed;
  }
  return parsed;
}

// Legacy (r1-r3) bank records: remove them so the space goes to pools/shards.
function clearQuestionBankCacheKeys(cache) {
  var removed = 0;
  for (var l = 0; l < TX_LANGS.length; l++) {
    removed += clearQuestionCacheRecord(cache, QUESTION_CACHE_PREFIX + 'bank_' + TX_LANGS[l], QUESTION_BANK_MAX_PARTS);
  }
  return removed;
}

// ========== Per-license question pools ==========
// A pool contains license-filtered, deduplicated valid questions. Warmup shares
// its local language-bank memo across pools. A live miss elects one builder;
// contending requests receive an explicit retryable response.
// Filter BEFORE dedupe because source rows repeat IDs across license types.
function loadLicensePoolServer(lang, license, forceRebuild, memo) {
  var safeLang = normalizeQuestionCacheLanguage(lang);
  var lic = String(license || '').trim();
  if (!Object.prototype.hasOwnProperty.call(EXAM_STRUCTURE_SERVER, lic)) throw new Error('Invalid license for pool: ' + lic);
  var cache = CacheService.getScriptCache(), key = QUESTION_CACHE_PREFIX + 'pool_' + safeLang + '_' + lic;
  var hit;
  if (!forceRebuild) {
    hit = readQuestionCacheRecord(cache, key, QUESTION_POOL_MAX_PARTS);
    if (Array.isArray(hit) && hit.length) return hit;
    // r10: a live miss never reads Drive inside the request. Every caller gets
    // question_cache_busy (client retries in 3s) and one out-of-band rebuild
    // fills the gap; the inline build below remains only as the fallback when
    // no trigger could be scheduled.
    if (requestQuestionCacheRebuild('pool_' + safeLang + '_' + lic)) throw questionCacheBusy();
    diagMark('pool-build-inline:' + safeLang + '/' + lic);
  }
  var lease = claimQuestionCacheLease('pool_' + safeLang + '_' + lic);
  try {
    if (!forceRebuild) {
      hit = readQuestionCacheRecord(cache, key, QUESTION_POOL_MAX_PARTS);
      if (Array.isArray(hit) && hit.length) return hit;
    }
    var t0 = Date.now();
    var filtered = filterByLicenseServer(loadQuestionsForLanguageServer(safeLang, memo), lic);
    var seen = {}, pool = [];
    // Filter BEFORE dedupe: repeated IDs have distinct license rows.
    for (var f = 0; f < filtered.length; f++) {
      var q = filtered[f];
      if (!q || !q.id || seen[q.id] || !Array.isArray(q.answers) || q.answers.length < 2) continue;
      seen[q.id] = true;
      pool.push(q);
    }
    if (!pool.length) throw new Error('Empty pool ' + safeLang + '/' + lic + '; check question bank');
    var cached = writeQuestionCacheRecord(cache, key, pool, QUESTION_POOL_MAX_PARTS);
    if (memo) {
      if (!memo.cacheStatus) memo.cacheStatus = {};
      memo.cacheStatus[safeLang + '/' + lic] = cached;
    }
    Logger.log('[POOL] built ' + safeLang + '/' + lic + ': ' + pool.length + ' questions; cached=' + cached + '; ' + (Date.now() - t0) + 'ms; ' + (forceRebuild ? 'warmup' : 'MISS'));
    return pool;
  } finally {
    releaseQuestionCacheLease(lease);
  }
}

// Remove every cached pool chunk (all langs × all licenses). Used by the
// emergency reset so a bank refresh can never serve stale pools.
function clearAllLicensePools(cache, langs) {
  var removed = 0, licenses = Object.keys(EXAM_STRUCTURE_SERVER);
  for (var l = 0; l < langs.length; l++) {
    for (var c = 0; c < licenses.length; c++) {
      removed += clearQuestionCacheRecord(cache, QUESTION_CACHE_PREFIX + 'pool_' + langs[l] + '_' + licenses[c], QUESTION_POOL_MAX_PARTS);
    }
  }
  return removed;
}

// ========== Packed translation index ==========
// 128 fixed gzip/base64 shards replace the oversized one-key-per-question
// index. Exam starts batch-read only shards containing their selected IDs.
// The response retains all available language texts/answers and the original
// ci encoding for non-examinee callers. Correct-answer indices are never cached
// in translation shards; they are attached from the server answer key on read.
var TX_LANGS = ['he', 'ru', 'en', 'ar', 'fr', 'es', 'am'];

function questionTranslationShard(id) {
  var text = String(id), hash = 0;
  for (var i = 0; i < text.length; i++) hash = ((hash * 31) + text.charCodeAt(i)) >>> 0;
  return hash % QUESTION_TX_SHARDS;
}

// skipFailedLanguages: the caller (warmup) has already decided that a partial
// index is acceptable; any language that fails to load is left out. Without it
// only a language whose JSON is genuinely absent from Drive is skipped.
function buildTranslationIndexCache(memo, skipFailedLanguages) {
  var lease = claimQuestionCacheLease('translations');
  try {
    var cache = CacheService.getScriptCache(), banks = {}, byId = {};
    memo = memo || { banks: {}, cacheStatus: {} };
    var langs = [];
    for (var l = 0; l < TX_LANGS.length; l++) {
      var lang = TX_LANGS[l];
      var rows;
      try { rows = loadQuestionsForLanguageServer(lang, memo); }
      catch (eLang) {
        // Only a genuinely absent optional language is skipped; anything else
        // aborts before touching the published index.
        if (skipFailedLanguages || (eLang && eLang.code === 'question_language_unavailable')) {
          Logger.log('[TX] language omitted from index: ' + lang + ' (' + (eLang && eLang.message ? eLang.message : eLang) + ')');
          continue;
        }
        throw eLang;
      }
      langs.push(lang);
      banks[lang] = rows;
      for (var q = 0; q < rows.length; q++) {
        var row = rows[q];
        if (!row || row.id === undefined || row.id === null) continue;
        var per = byId[row.id] || (byId[row.id] = {});
        per[lang] = { t: row.text, a: row.answers };
      }
    }
    var shards = [], ids = Object.keys(byId);
    for (var s = 0; s < QUESTION_TX_SHARDS; s++) shards.push({});
    for (var i = 0; i < ids.length; i++) shards[questionTranslationShard(ids[i])][ids[i]] = byId[ids[i]];
    var generation = Utilities.getUuid(), values = {}, keys = [];
    for (var b = 0; b < shards.length; b++) {
      var encoded = encodeQuestionCache(shards[b]);
      if (encoded.length > QUESTION_CACHE_PART_BYTES) throw new Error('Translation shard ' + b + ' exceeds reserved cache budget');
      var key = QUESTION_CACHE_PREFIX + 'tx_' + b;
      keys.push(key);
      values[key] = generation + ':' + encoded;
    }
    cache.putAll(values, 21600);
    var verified = cache.getAll(keys);
    for (var k = 0; k < keys.length; k++) {
      if (verified[keys[k]] !== values[keys[k]]) throw new Error('Translation shard missing immediately after write');
    }
    // Decode one read-back shard with the real decoder (see writeQuestionCacheRecord).
    var probe = decodeQuestionCache(verified[keys[0]].substring(generation.length + 1));
    if (Object.keys(probe).length !== Object.keys(shards[0]).length) throw new Error('Translation shard read-back decode mismatch');
    if (!langs.length) throw new Error('No language bank could be loaded');
    var meta = JSON.stringify({ g: generation, langs: langs, count: ids.length, builtAt: Date.now() });
    cache.put(QUESTION_CACHE_PREFIX + 'tx_meta', meta, 21600);
    if (cache.get(QUESTION_CACHE_PREFIX + 'tx_meta') !== meta) throw new Error('Translation manifest missing after write');
    Logger.log('[TX] index cached: ' + ids.length + ' questions, ' + langs.length + ' languages, in ' + QUESTION_TX_SHARDS + ' shards');
    return { count: ids.length, langs: langs, cached: true };
  } catch (e) {
    Logger.log('[TX] WRITE FAILED: ' + (e && e.message ? e.message : e));
    throw e;
  } finally {
    releaseQuestionCacheLease(lease);
  }
}

function publishedTranslationLanguageCount(cache) {
  try {
    var meta = JSON.parse(cache.get(QUESTION_CACHE_PREFIX + 'tx_meta') || 'null');
    return meta && Array.isArray(meta.langs) ? meta.langs.length : 0;
  } catch (e) { return 0; }
}

// Assemble translations for the selected questions from the per-id index. Returns
// the SAME { lang: { id: {t,a[,ci]} } } shape as the bank path, or null to signal
// "index not ready / incomplete → caller should fall back to the banks".
function tryTranslationsFromIndex(idList, includeCi) {
  try {
    var cache = CacheService.getScriptCache();
    var raw = cache.get(QUESTION_CACHE_PREFIX + 'tx_meta');
    if (!raw) return null;
    var meta = JSON.parse(raw);
    if (!meta || !meta.g || !Array.isArray(meta.langs) || !meta.langs.length) return null;
    var keys = [], needed = {};
    for (var i = 0; i < idList.length; i++) {
      var shard = questionTranslationShard(idList[i]);
      if (!needed[shard]) { needed[shard] = true; keys.push(QUESTION_CACHE_PREFIX + 'tx_' + shard); }
    }
    // Two cache reads total: manifest + only the needed packed shards.
    var got = cache.getAll(keys), byId = {}, prefix = meta.g + ':';
    for (var k = 0; k < keys.length; k++) {
      var packed = got[keys[k]];
      if (typeof packed !== 'string' || packed.indexOf(prefix) !== 0) return null;
      var entries = decodeQuestionCache(packed.substring(prefix.length));
      Object.keys(entries).forEach(function(id) { byId[id] = entries[id]; });
    }
    for (var j = 0; j < idList.length; j++) if (!byId[idList[j]]) return null;
    var translations = {};
    for (var l = 0; l < meta.langs.length; l++) {
      var lang = meta.langs[l], altMap = {};
      for (var q = 0; q < idList.length; q++) {
        var id = idList[q], per = byId[id];
        if (!per || !per[lang]) continue;
        var entry = { t: per[lang].t, a: per[lang].a };
        if (includeCi && typeof lookupCorrectIndex === 'function') {
          var ci = lookupCorrectIndex(Number(id), lang);
          if (ci !== null && ci !== undefined) entry.ci = ci ^ (id % 256);
        }
        altMap[id] = entry;
      }
      translations[lang] = altMap;
    }
    return translations;
  } catch (e) {
    Logger.log('[TX] invalid/read failed: ' + (e && e.message ? e.message : e));
    return null;
  }
}

// Live exam start never loads full language banks (r4). When the packed index
// is not ready or a shard is missing, the exam starts WITHOUT prefetched
// translations; the client already falls back to getQuestionsByIds on a
// mid-exam language switch. The previous fallback loaded all seven banks in the
// request (~5s each, ~35s total) and was the measured cause of 30-50s exam
// starts whenever a single shard had been evicted.
function buildExamTranslations(selected, includeCi) {
  var idList = [];
  for (var i = 0; i < selected.length; i++) idList.push(selected[i].id);
  var fast = tryTranslationsFromIndex(idList, includeCi);
  if (fast !== null) return fast;
  Logger.log('[TX] shard MISS: exam starts without prefetched translations for ' + idList.length + ' questions (client fetches on demand)');
  return null;
}

// Pick 30 questions per the license blueprint, return them WITHOUT the
// correct-answer index. Authenticated clients only — falls back to a
// rate-limited guest path for the standalone exam.html flow.
// ========== Practice block while an exam is running ==========
// Class practice (student.html) and exams share ONE Apps Script account — the
// same ~30 concurrent execution slots and the same question cache/pools. A
// practice wave (the daily peak) can therefore saturate the server and freeze
// the exam side: the documented ~13:00 outage was a practice peak, and single-
// site exam mornings still stalled because a class was practicing in parallel.
// This guard turns class practice OFF whenever at least one exam session is
// open. Exam draws (auth 'examinee') and standalone exam.html are NEVER
// affected. The answer is cached ~60s so practice calls don't re-scan the
// sessions sheet (which would add the very load we're removing). It self-clears
// within ~60s after the last session closes or expires.
// r6: OFF by default. The block was introduced while every practice request
// rebuilt a full question pool (the cache was unreadable, see r5). A practice
// request is now a ~2s cache hit, so class practice no longer competes with
// exams. Re-enable without a deployment by setting the Script Property
// PRACTICE_BLOCK_DURING_EXAMS to "on" (takes effect within ~60s); remove it or
// set anything else to disable again.
var PRACTICE_BLOCK_PROPERTY = 'PRACTICE_BLOCK_DURING_EXAMS';
function isExamSessionActiveForPracticeBlock() {
  try {
    var cache = CacheService.getScriptCache();
    var flag = cache.get('exam_active_block');
    if (flag === 'Y') return true;
    if (flag === 'N') return false;
    var enabled = String(PropertiesService.getScriptProperties().getProperty(PRACTICE_BLOCK_PROPERTY) || '').trim().toLowerCase() === 'on';
    if (!enabled) {
      try { cache.put('exam_active_block', 'N', 60); } catch (eOff) {}
      return false;
    }
    // Cache miss (at most once per 60s) → scan the sessions sheet once. Same
    // active-session rule the examiner dashboard uses: פעיל(10) true AND not
    // past תקף עד(9).
    var active = false;
    var sess = getSheet('סשנים').getDataRange().getValues();
    var nowT = new Date().getTime();
    for (var s = 1; s < sess.length; s++) {
      var isActive = sess[s][10] === true || String(sess[s][10]).toUpperCase() === 'TRUE';
      if (!isActive) continue;
      var validUntil = sess[s][9] ? new Date(sess[s][9]).getTime() : 0;
      if (!validUntil || validUntil > nowT) { active = true; break; }
    }
    try { cache.put('exam_active_block', active ? 'Y' : 'N', 60); } catch (ePut) {}
    return active;
  } catch (e) {
    // Fail OPEN: a transient sheet/cache error must never block ALL practice.
    Logger.log('[PRACTICE-BLOCK] check failed, allowing practice: ' + (e && e.message ? e.message : e));
    return false;
  }
}

