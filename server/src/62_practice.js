// ========== Practice draws (cold path, its OWN deployment) =================
//
// Moved out of 60_exam.js on 22/09/2026 (DESIGN §13.3) unchanged: practice is
// a student/teacher feature that never runs during an exam, and every byte in
// the exam deployment is compiled by Google on every single exam request.
// What stays in the hot file is practiceCiByLang (70_questions.js, `both`) —
// it is the answer-key reader and the exam path is built on the same index.

// ---- startPractice ----------------------------------------------------------
// GET, no token. Practice scores on the client, so it gets the correct index of
// every language the question exists in (XOR-encoded) and never needs another
// round trip — a language switch mid-practice is local.
// One draw is one Worker request, and that request may ask for at most 30 assets
// (DESIGN §11.3) — so 30 is the cap in EVERY mode, not just the exam blueprint.
var PRACTICE_MAX_COUNT = 30;
var PRACTICE_DEFAULT_COUNT = 15;
defineAction('startPractice', { methods: ['GET'], auth: 'none', handler: handleStartPractice });
function handleStartPractice(p) {
  var rlErr = practiceRateLimit(p);
  if (rlErr) return rlErr;
  var lang = String(p.language || 'he').toLowerCase();
  var license = String(p.license || p.licenseType || 'B').trim();
  if (!EXAM_STRUCTURE_SERVER[license]) {
    return jsonResponse({ status: 'error', code: 'unknown_license', message: 'דרגה לא מוכרת: ' + license });
  }
  // Same rule as startExam: no grant, no texts, so say so before drawing.
  if (!bankGrantConfigured()) return bankNotConfiguredResponse();
  var mode = String(p.mode || 'exam');
  var picked;
  try { picked = practiceSelection(mode, license, lang, p); }
  catch (err) {
    if (!err || err.code !== 'bank_unavailable') throw err;
    return jsonResponse({ status: 'error', code: 'bank_unavailable', detail: err.detail, message: 'אין מספיק שאלות לתרגול' });
  }
  if (!picked.length) return jsonResponse({ status: 'error', code: 'no_questions', message: 'לא נמצאו שאלות לתרגול' });
  var questions = [], ids = [];
  for (var i = 0; i < picked.length; i++) {
    questions.push({ id: picked[i].id, topic: picked[i].topic, ci: practiceCiByLang(picked[i].id) || {} });
    ids.push(picked[i].id);
  }
  return jsonResponse({ status: 'ok', mode: mode, count: questions.length,
    bank: bankGrantFor('practice', ids, practiceSubject(p)), questions: questions });
}

// Who the practice grant was issued to — the same four identities the rate
// limit separates, so a grant can be traced back to the draw that produced it.
function practiceSubject(p) {
  if (p.classCode && p.studentId) return String(p.classCode) + ':' + String(p.studentId);
  if (p.studentId) return 'home:' + String(p.studentId);   // student.html, no class code typed
  if (p.standaloneIdNumber) return normalizeId(p.standaloneIdNumber);
  return 'guest';
}

function practiceSelection(mode, license, lang, p) {
  if (mode === 'ids') return practiceByIds(p.ids, license, lang);
  if (mode === 'category' && p.categoryFilter) return practiceByCategory(String(p.categoryFilter), license, lang, practiceCount(p));
  return drawExamIds(license, lang);   // full 30-question blueprint
}

function practiceCount(p) {
  var n = Number(p.maxCount) || PRACTICE_DEFAULT_COUNT;
  return Math.max(1, Math.min(PRACTICE_MAX_COUNT, n));
}

function practiceByCategory(topic, license, lang, maxCount) {
  var byTopic = indexIdsByTopic(license, lang), pool = shuffleArrayServer(byTopic[topic] || []), out = [];
  for (var i = 0; i < pool.length && out.length < maxCount; i++) out.push({ id: pool[i], topic: topic });
  return out;
}

// Spaced repetition: the client names the ids it wants back. Unknown ids and ids
// missing from this language are dropped rather than failing the whole request.
function practiceByIds(raw, license, lang) {
  var bit = questionLangBit(lang), out = [], seen = {};
  var parts = String(raw || '').split(',');
  for (var i = 0; i < parts.length && out.length < PRACTICE_MAX_COUNT; i++) {
    var id = parseInt(String(parts[i]).trim(), 10);
    if (isNaN(id) || seen[id]) continue;
    seen[id] = true;
    var entry = questionIndexEntry(id);
    if (!entry || !(entry.l & bit)) continue;
    out.push({ id: id, topic: entry.c[license] || '' });
  }
  return out;
}

// ---- Who is asking, and what that caller is allowed -------------------------
// FOUR identities, and no two real users share a bucket:
//
//   class student   classCode + studentId    20/min per (class, student)
//   home practice   studentId, no class      20/min per studentId
//   standalone      standaloneIdNumber       5/min per ID (exam.html)
//   guest           nothing at all           5/min, ONE shared bucket
//
// The home identity is the bug this fixes. student.html asks for the class code
// as OPTIONAL — a soldier practising at home has no class — so those callers
// sent a studentId and no classCode and fell through to the guest bucket, whose
// identifier is the constant 'anon'. That is five practice draws a minute for
// the whole country TOGETHER: on a busy evening most of them would have got
// "יותר מדי בקשות" and nothing else. What is left in the guest bucket is a
// caller that names nothing at all — our own pages always identify themselves,
// so that is scripts, and five a minute is the right allowance for a script.
//
// One ceiling then covers everyone WITHOUT a class code (home + standalone +
// guest) together, because an identifier the caller invents is not an identity:
// vary the studentId and each one is worth another 20/min. The ceiling does NOT
// protect the bank — practice returns `ci`, so a draw is an answer oracle by
// design and no rate a real student would accept can stop a patient scraper
// from walking 1,700 questions. What it protects is Apps Script: it bounds how
// many of the shared executions one flood can hold, which is the resource an
// exam morning on the same script actually competes for. Class students are
// outside the ceiling — they are enrolled, traceable, and the last people a
// flood may lock out.
var PRACTICE_NOCLASS_GLOBAL_MAX = 300;
function practiceRateLimit(p) {
  if (p.classCode && p.studentId) {
    return requireRateLimit('startPractice_student', String(p.classCode) + '_' + String(p.studentId), 20, 60);
  }
  var own;
  if (p.studentId) own = requireRateLimit('startPractice_home', String(p.studentId), 20, 60);
  else if (p.standaloneIdNumber) own = requireRateLimit('startPractice_standalone', normalizeId(p.standaloneIdNumber), 5, 60);
  else own = requireRateLimit('startPractice_guest', 'anon', 5, 60);
  // The ceiling is charged only for a draw the caller's own identity allowed, so
  // one blocked device cannot spend everybody else's allowance while it retries.
  return own || requireRateLimit('startPractice_noclass', 'global', PRACTICE_NOCLASS_GLOBAL_MAX, 60);
}
