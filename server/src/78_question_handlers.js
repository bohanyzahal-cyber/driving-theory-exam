function handleGetExamQuestions(p) {
  // Determine auth context
  var auth = 'guest';
  var examineeAudioMode = null;
  if (p.sessionCode && p.idNumber && p.examineeToken) {
    diagMark('sheet:token-examstart');
    var ev = verifyExamineeToken(p.sessionCode, p.idNumber, p.examineeToken);
    if (!ev.valid) {
      return jsonResponse({ status: 'error', message: 'Examinee token invalid', reason: ev.reason });
    }
    auth = 'examinee';
    // Second capture point for per-examinee audio: covers an examiner who set
    // it AFTER approving, when the client's approval polling has already
    // stopped but the examinee hasn't pressed "start" yet. Free — the token
    // check above already read the row.
    examineeAudioMode = ev.audioMode || 'off';
  } else if (p.token && p.examinerId) {
    if (!verifyToken(p.examinerId, p.token)) {
      return jsonResponse({ status: 'error', message: 'Examiner token invalid', tokenExpired: true });
    }
    auth = 'examiner';
  } else if (p.classCode && p.studentId) {
    // Student practice mode — looser auth, just rate-limit
    auth = 'student';
  } else if (p.standaloneIdNumber) {
    // Standalone exam.html — examinee enters their ID, no token; rate-limit hard
    auth = 'standalone';
  }

  // Physical practice block: class practice (student.html) is turned OFF while
  // any exam session is open, so it can't share the ~30 slots with the exam
  // wave. Only auth 'student' (classCode+studentId) is affected — exam draws
  // and standalone are not. Cheap: a cached flag, no per-call sheet scan.
  if (auth === 'student' && isExamSessionActiveForPracticeBlock()) {
    return jsonResponse({ status: 'error', code: 'practice_blocked_exam_active',
      message: 'התרגול סגור כעת מפני שמתקיים מבחן במערכת. נסו שוב מאוחר יותר.' });
  }

  // Rate limit (per auth + identifier)
  var rlId = questionRequestRateId(p, auth);
  var rlMax = (auth === 'guest' || auth === 'standalone') ? 5 : 20;
  var rlErr = requireRateLimit('getExamQuestions_' + auth, rlId, rlMax, 60);
  if (rlErr) return rlErr;

  var lang = String(p.language || 'he').toLowerCase();
  var license = String(p.license || p.licenseType || 'B');
  if (!EXAM_STRUCTURE_SERVER[license]) {
    return jsonResponse({ status: 'error', message: 'Unknown license: ' + license });
  }

  // Per-license pool: license-filtered + deduped, prebuilt by the warmup
  // trigger and self-healing on miss (see loadLicensePoolServer — that is the
  // only copy of the filter/dedupe logic now). Loads ~1/3 of the bytes the
  // full bank did, which is what took exam-start from ~10s to ~2-3s.
  var pool;
  try { pool = loadLicensePoolServer(lang, license); }
  catch (e) {
    var busyResponse = theoryRetryableErrorResponse(e);
    if (busyResponse) return busyResponse;
    Logger.log('loadLicensePoolServer(' + lang + ',' + license + ') failed: ' + (e && e.message));
    return jsonResponse({ status: 'error', message: 'שגיאה בטעינת שאלות. נסה שוב.' });
  }

  // Category-quiz mode (student.html practice by topic): return up to N
  // questions matching one category, skipping the 30-question blueprint.
  var mode = String(p.mode || 'exam');
  var selected;
  if (mode === 'category' && p.categoryFilter) {
    var wantTopic = String(p.categoryFilter);
    var catPool = pool.filter(function(q) { return classifyCategoryServer(q.category) === wantTopic; });
    catPool = shuffleArrayServer(catPool);
    var max = Number(p.maxCount) || 15;
    selected = catPool.slice(0, Math.min(catPool.length, max));
    if (selected.length === 0) {
      return jsonResponse({ status: 'error', message: 'No questions in category', topic: wantTopic });
    }
  } else {
    // Default exam mode: pick per EXAM_STRUCTURE blueprint.
    var byTopic = {};
    for (var j = 0; j < pool.length; j++) {
      var t = classifyCategoryServer(pool[j].category);
      if (!t) continue;
      if (!byTopic[t]) byTopic[t] = [];
      byTopic[t].push(pool[j]);
    }
    var blueprint = EXAM_STRUCTURE_SERVER[license];
    selected = [];
    var usedIds = {};
    for (var topic in blueprint) {
      var needed = blueprint[topic];
      var avail = shuffleArrayServer(byTopic[topic] || []);
      var count = 0;
      for (var ai = 0; ai < avail.length && count < needed; ai++) {
        if (usedIds[avail[ai].id]) continue;
        usedIds[avail[ai].id] = true;
        selected.push(avail[ai]);
        count++;
      }
      if (count < needed) {
        return jsonResponse({
          status: 'error',
          message: 'Not enough questions for topic',
          topic: topic,
          have: (byTopic[topic] || []).length,
          need: needed
        });
      }
    }
    selected = shuffleArrayServer(selected);
  }

  // Remember which IDs we issued so handleRegisterExamQuestions can refuse
  // submissions that name questions we didn't actually give the examinee.
  // Only meaningful for the examinee-token path (we have a stable identifier).
  if (auth === 'examinee' && p.sessionCode && p.idNumber) {
    var issuedKey = 'issued_qs_' + p.sessionCode + '_' + normalizeId(p.idNumber);
    var ids = selected.map(function(q) { return q.id; });
    try { CacheService.getScriptCache().put(issuedKey, JSON.stringify(ids), 21600); } catch (e) { /* skip */ }
  }

  // For real exams (auth === 'examinee') we deliberately omit the correct
  // index — scoring happens server-side via handleRegisterExamQuestions /
  // handleSubmitResult using ANSWER_KEY_BY_LANG. For practice/standalone
  // flows (exam.html, student.html) the client needs to score locally, so
  // we include the encoded `ci` field that matches the legacy questions.js
  // format: ci = correctIndex XOR (id mod 256).
  if (auth !== 'examinee' && typeof lookupCorrectIndex === 'function') {
    for (var ci_i = 0; ci_i < selected.length; ci_i++) {
      var origCorrect = lookupCorrectIndex(Number(selected[ci_i].id), lang);
      if (origCorrect !== null && origCorrect !== undefined) {
        selected[ci_i].ci = origCorrect ^ (selected[ci_i].id % 256);
      }
    }
  }

  // Pre-fetch all 7 languages so mid-exam language switches are instant
  // (no extra round trip). Adds ~120-150 KB to the response. Cold-start
  // server cost is real (7 Drive reads) but cached for 6h after that.
  //
  // For any non-examinee caller we include each translation's `ci` (correct
  // index, XOR-encoded) so the client can score correctly per language when
  // the translator put answers in a different order. This includes `guest`
  // (student practicing without a classCode) — the previous student-only gate
  // left guest practice silently wrong after a mid-practice language switch.
  // For examinee auth we keep `ci` stripped — server is sole source of truth.
  var includeCiInTranslations = (auth !== 'examinee');
  var translations = null;
  if (p.includeTranslations === 'true' || p.includeTranslations === '1') {
    // Fast path: assemble from the pre-built per-id translation index (built by
    // the warmup trigger) so we DON'T parse 6 extra full language banks on every
    // exam-start — the main per-request cost that saturated execution slots in
    // the morning start-wave. Falls back to loading the banks when the index
    // isn't ready (e.g. first request after a cache clear), so behavior is never
    // worse than before. Response shape is identical → offline mid-exam language
    // switching unchanged. See buildExamTranslations / buildTranslationIndexCache.
    // `ci` (non-examinee only) is still layered per-language from the answer key.
    translations = buildExamTranslations(selected, includeCiInTranslations);
  }

  var responseBody = { status: 'ok', auth: auth, count: selected.length, questions: selected };
  if (examineeAudioMode !== null) responseBody.audioMode = examineeAudioMode;
  if (translations) responseBody.translations = translations;
  return jsonResponse(responseBody);
}

// ========== Re-fetch questions in a different language ==========
// When an examinee/student/practice user changes language mid-exam, the
// client calls this with the set of question IDs already shown and the new
// language. Server returns those same IDs with text/answers in the new
// language so the exam can continue without losing progress.
function handleGetQuestionsByIds(p) {
  // Match auth model of handleGetExamQuestions
  var auth = 'guest';
  if (p.sessionCode && p.idNumber && p.examineeToken) {
    var ev = verifyExamineeToken(p.sessionCode, p.idNumber, p.examineeToken);
    if (!ev.valid) {
      return jsonResponse({ status: 'error', message: 'Examinee token invalid', reason: ev.reason });
    }
    auth = 'examinee';
  } else if (p.token && p.examinerId) {
    if (!verifyToken(p.examinerId, p.token)) {
      return jsonResponse({ status: 'error', message: 'Examiner token invalid', tokenExpired: true });
    }
    auth = 'examiner';
  } else if (p.classCode && p.studentId) {
    auth = 'student';
  } else if (p.standaloneIdNumber) {
    auth = 'standalone';
  }

  // Practice block (see isExamSessionActiveForPracticeBlock): flashcards /
  // language-switch in class practice are off while an exam session is open.
  if (auth === 'student' && isExamSessionActiveForPracticeBlock()) {
    return jsonResponse({ status: 'error', code: 'practice_blocked_exam_active',
      message: 'התרגול סגור כעת מפני שמתקיים מבחן במערכת. נסו שוב מאוחר יותר.' });
  }

  var rlErr = requireRateLimit('getQuestionsByIds_' + auth,
    questionRequestRateId(p, auth),
    30, 60);
  if (rlErr) return rlErr;

  var lang = String(p.language || 'he').toLowerCase();
  var idsRaw = String(p.ids || '');
  var ids = idsRaw.split(',').map(function(s) {
    var n = parseInt(String(s).trim(), 10);
    return isNaN(n) ? null : n;
  }).filter(function(n) { return n !== null; });

  if (ids.length === 0) return jsonResponse({ status: 'error', message: 'No IDs provided' });
  if (ids.length > 50) return jsonResponse({ status: 'error', message: 'Too many IDs (max 50)' });

  // r10: the per-license pools already hold every question of this language;
  // a full Drive read happens only when they do not cover the request.
  var allQuestions, cachedById = null;
  try { cachedById = questionsFromCachedPools(lang, ids); } catch (eCached) { cachedById = null; }
  if (cachedById) {
    allQuestions = [];
    for (var ck in cachedById) if (Object.prototype.hasOwnProperty.call(cachedById, ck)) allQuestions.push(cachedById[ck]);
  } else {
    try { allQuestions = loadQuestionsForLanguageServer(lang); }
    catch (e) {
      var busyResponse = theoryRetryableErrorResponse(e);
      if (busyResponse) return busyResponse;
      Logger.log('loadQuestionsForLanguageServer(' + lang + ') failed: ' + (e && e.message));
      return jsonResponse({ status: 'error', message: 'שגיאה בטעינת שאלות. נסה שוב.' });
    }
  }

  // Build id → question lookup
  var byId = {};
  for (var i = 0; i < allQuestions.length; i++) {
    var q = allQuestions[i];
    if (q && q.id) byId[q.id] = q;
  }

  var results = [];
  for (var j = 0; j < ids.length; j++) {
    var found = byId[ids[j]];
    if (!found) {
      results.push(null);
      continue;
    }
    var entry = {
      id: found.id,
      text: found.text,
      answers: found.answers,
      category: found.category,
      licenseType: found.licenseType,
      imageUrl: found.imageUrl,
      language: found.language || lang
    };
    // Include ci for non-examinee callers (practice/standalone) so they can score locally.
    if (auth !== 'examinee' && typeof lookupCorrectIndex === 'function') {
      var origCorrect = lookupCorrectIndex(Number(found.id), lang);
      if (origCorrect !== null && origCorrect !== undefined) {
        entry.ci = origCorrect ^ (found.id % 256);
      }
    }
    results.push(entry);
  }

  return jsonResponse({ status: 'ok', count: results.length, questions: results });
}

// ========== Question search (for find_image.html examiner utility) ==========
// Returns up to 20 questions whose text/answers/category match the query
// substring. Requires examiner token — this is an internal staff utility,
// not for examinees. Cross-language search: caller passes the language and
// we search that language's pre-translated dataset.
function handleSearchQuestions(p) {
  var authErr = requireToken(p);
  if (authErr) return authErr;
  var rlErr = requireRateLimit('searchQuestions', String(p.examinerId || ''), 30, 60);
  if (rlErr) return rlErr;

  var query = String(p.q || '').trim().toLowerCase();
  if (query.length < 2) return jsonResponse({ status: 'ok', matches: [], note: 'Query too short' });

  var lang = String(p.language || 'he').toLowerCase();
  var allQuestions;
  try { allQuestions = loadQuestionsForLanguageServer(lang); }
  catch (e) {
    Logger.log('loadQuestionsForLanguageServer(' + lang + ') failed: ' + (e && e.message));
    return jsonResponse({ status: 'error', message: 'שגיאה בטעינת שאלות. נסה שוב.' });
  }

  var matches = [];
  var MAX_MATCHES = 20;
  for (var i = 0; i < allQuestions.length && matches.length < MAX_MATCHES; i++) {
    var q = allQuestions[i];
    if (!q || !q.text) continue;
    var hay = (q.text + ' ' + (q.answers || []).join(' ') + ' ' + (q.category || '')).toLowerCase();
    if (hay.indexOf(query) === -1) continue;
    var match = {
      id: q.id,
      text: q.text,
      answers: q.answers,
      category: q.category,
      licenseType: q.licenseType,
      imageUrl: q.imageUrl
    };
    // Include encoded ci for the search utility (examiner trusted view)
    if (typeof lookupCorrectIndex === 'function') {
      var orig = lookupCorrectIndex(Number(q.id), lang);
      if (orig !== null && orig !== undefined) {
        match.ci = orig ^ (q.id % 256);
      }
    }
    matches.push(match);
  }

  return jsonResponse({ status: 'ok', matches: matches, language: lang, query: query });
}

// ========== Bohan-site portal — MOVED OUT (2026-07-31) ==========
// The examiners-portal auth + data (action=bohanSiteAuth: shared password,
// HMAC tokens, examiner + Waze-location lists) moved to its OWN standalone
// Apps Script project ("bohan-site-server"), fully separate from this exam
// system. This deployment no longer serves the portal.

// ========== Result-upload HMAC token (for Cloudflare Worker auth) ==========
// Issues a short-lived signed token that the browser sends in X-Auth-Token
// when POSTing exam-result HTML to the exam-results Cloudflare Worker.
// The Worker verifies the same HMAC with its own copy of the secret.
//
// Setup: in Apps Script, set ScriptProperty 'RESULT_UPLOAD_SECRET' to a long
// random string. Set the IDENTICAL value as a Worker secret binding named
// UPLOAD_SECRET. The secret never reaches the browser.
function handleGetResultUploadToken(p) {
  // Examiner auth (token+id) is already enforced by the doGet dispatcher
  // before this handler runs — see examinerActions list.
  var props = PropertiesService.getScriptProperties();
  var secret = props.getProperty('RESULT_UPLOAD_SECRET');
  if (!secret) {
    return jsonResponse({
      status: 'error',
      message: 'RESULT_UPLOAD_SECRET not configured in Apps Script properties',
      code: 'not_configured'
    });
  }
  var payload = JSON.stringify({ exp: Date.now() + 5 * 60 * 1000 });
  var payloadB64 = Utilities.base64EncodeWebSafe(payload).replace(/=+$/, '');
  var sigBytes = Utilities.computeHmacSha256Signature(payloadB64, secret);
  var sigB64 = Utilities.base64EncodeWebSafe(sigBytes).replace(/=+$/, '');
  return jsonResponse({ status: 'ok', token: payloadB64 + '.' + sigB64 });
}

function unmarkPendingCompleted(sessionCode, idNumber) {
  // Restore the latest pending row from 'completed' back to 'in_exam' so a resumed
  // exam (after a premature close-fail was reverted) can be re-submitted. BUG FIX:
  // this previously read col index 6 (=language) and tested 'done' (a status that
  // is NEVER written — markPendingCompleted writes 'completed'), then wrote to
  // column 7 (=language) and scanned oldest-first — so it silently no-op'd and
  // could corrupt the language cell. Now mirrors markPendingCompleted exactly:
  // status = col index 5 / column 6, newest row first.
  var pendingSheet = getSheet('ממתינים');
  if (!pendingSheet) return;
  var data = pendingSheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === String(sessionCode) && normalizeId(data[i][1]) === normalizeId(idNumber)) {
      if (String(data[i][5] || '').trim() === 'completed') {
        pendingSheet.getRange(i + 1, 6).setValue('in_exam');
      }
      break;
    }
  }
}

// ========== שיתוף תוצאה דרך CacheService ==========

function handleUploadResultHtml(data) {
  // DISABLED: result-HTML hosting moved to the authenticated Cloudflare Worker
  // (RESULTS_URL). This Apps Script endpoint had NO auth and served attacker HTML
  // from the trusted Google origin (stored-XSS/phishing). The live client never
  // calls it (it posts to the Worker), so returning early is safe and closes the hole.
  return jsonResponse({ status: 'error', message: 'disabled — use the results worker' });
  // eslint-disable-next-line
  if (!data.html || !data.requestId) {
    return jsonResponse({ status: 'error', message: 'Missing html or requestId' });
  }

  try {
    var cache = CacheService.getScriptCache();
    var html = data.html;
    var CHUNK_SIZE = 90000; // 90KB per chunk (limit is 100KB)
    var numChunks = Math.ceil(html.length / CHUNK_SIZE);

    // שומר HTML בחלקים ב-CacheService (עד 6 שעות)
    var chunks = {};
    for (var i = 0; i < numChunks; i++) {
      chunks['result_' + data.requestId + '_' + i] = html.substring(i * CHUNK_SIZE, (i + 1) * CHUNK_SIZE);
    }
    chunks['result_' + data.requestId + '_meta'] = String(numChunks);
    cache.putAll(chunks, 21600);

    // בונה קישור לצפייה דרך doGet
    var viewLink = ScriptApp.getService().getUrl() + '?action=viewResult&id=' + data.requestId;

    // שומר תוצאה ב-ScriptProperties כדי שה-client יוכל לקרוא דרך GET polling
    var props = PropertiesService.getScriptProperties();
    props.setProperty('upload_' + data.requestId, JSON.stringify({ link: viewLink }));

    return jsonResponse({ status: 'ok', link: viewLink });
  } catch (err) {
    if (data.requestId) {
      var props2 = PropertiesService.getScriptProperties();
      props2.setProperty('upload_' + data.requestId, JSON.stringify({ error: err.toString() }));
    }
    return jsonResponse({ status: 'error', message: err.toString() });
  }
}

function handleGetUploadResult(p) {
  var requestId = p.requestId;
  if (!requestId) return jsonResponse({ status: 'error', message: 'Missing requestId' });

  var props = PropertiesService.getScriptProperties();
  var stored = props.getProperty('upload_' + requestId);
  if (!stored) return jsonResponse({ status: 'pending' });

  // ניקוי
  props.deleteProperty('upload_' + requestId);

  var result = JSON.parse(stored);
  if (result.error) {
    return jsonResponse({ status: 'error', message: result.error });
  }
  return jsonResponse({ status: 'ok', link: result.link });
}

