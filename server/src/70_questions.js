// ========== The question index — all the server knows about questions ========
//
// Until 21/09/2026 the server held the question TEXTS: seven language banks in a
// private Drive folder, copied into CacheService as per-license pools, kept warm
// by a trigger, guarded by leases and rebuilt out of band. That subsystem was
// ~1,500 lines and it is what a killed execution died inside (the 360s kills of
// 09-09/09-11), what Drive reads added to every commander dashboard (r12/r13),
// and what made exam-start a 10-second request. It bought nothing: the texts AND
// the correct answers were already served to anyone who asked (getQuestionsByIds
// with any studentId, verified live 21/09).
//
// So the texts left the script. They are NOT public either (DESIGN §11): they
// are private Workers assets of the session-gateway (built by tools/build_bank.js
// into cloudflare-workers/session-gateway/assets/, never in the repo and never on
// Pages), and the Worker serves each device only the ids it was issued, against a
// grant this script signs (bankGrantFor, 20_auth.js). The server keeps only:
//   * QUESTION_INDEX — id → { c: {license: topic}, l: language bitmask, img }
//     (generated into this file at build time from deployment/question_index.json)
//   * the answer key (deployment/answer_key.gs, pasted separately, never public)
// From those two it can draw an exam and score it, with zero Drive and zero
// question data in CacheService.

// @@QUESTION_INDEX@@

// Bit 0 = he … bit 6 = am in QUESTION_INDEX[id].l — the order tools/build_bank.js
// writes. A bit says the id exists in that language's bank, which is also what
// makes its answer key trustworthy for that language (en/fr/es/ar order their
// answers differently from Hebrew — see TRANSLATION_LINEAGE_2026-09-20.md).
var QUESTION_LANGS = ['he', 'ru', 'en', 'ar', 'fr', 'es', 'am'];
// Every question in every one of the seven generated banks has exactly four
// answers (checked over all 6,245-6,270 rows per bank, 21/09/2026), so the
// displayed order is a permutation of [0..3].
var QUESTION_ANSWER_COUNT = 4;

var EXAM_STRUCTURE_SERVER = {
  'B':  { 'בטיחות': 7, 'הכרת הרכב': 7, 'חוק': 7, 'תמרורים': 9 },
  '1':  { 'בטיחות': 5, 'הכרת הרכב': 5, 'חוק': 6, 'תמרורים': 6, 'ספציפי': 8 },
  'C1': { 'בטיחות': 5, 'הכרת הרכב': 5, 'חוק': 5, 'תמרורים': 5, 'ספציפי': 10 },
  'C':  { 'בטיחות': 5, 'הכרת הרכב': 4, 'חוק': 3, 'תמרורים': 4, 'ספציפי': 14 },
  'D':  { 'בטיחות': 4, 'הכרת הרכב': 2, 'חוק': 5, 'תמרורים': 4, 'ספציפי': 15 }
};

// The bank's raw category → one of the five blueprint topics. Kept on the server
// because the index stores the classified topic and the reports classify the
// categories they read out of result rows.
function classifyCategoryServer(cat) {
  var c = String(cat || '').trim();
  if (/ספציפי/.test(c)) return 'ספציפי'; // ספציפי
  if (/בטיחות/.test(c)) return 'בטיחות'; // בטיחות
  if (/הכרת הרכב/.test(c)) return 'הכרת הרכב'; // הכרת הרכב
  if (/חוק/.test(c)) return 'חוק'; // חוק
  if (/תמרורים/.test(c)) return 'תמרורים'; // תמרורים
  if (/זכות קדימה/.test(c)) return 'חוק'; // זכות קדימה → חוק
  return '';
}

function shuffleArrayServer(arr) {
  var a = arr.slice();
  for (var i = a.length - 1; i > 0; i--) {
    var j = Math.floor(Math.random() * (i + 1));
    var t = a[i]; a[i] = a[j]; a[j] = t;
  }
  return a;
}

function questionIndexEntry(id) {
  var entry = QUESTION_INDEX[String(id)];
  return entry || null;
}

function questionIndexCount() { return Object.keys(QUESTION_INDEX).length; }

function questionLangBit(lang) {
  var i = QUESTION_LANGS.indexOf(String(lang || 'he').toLowerCase());
  return i < 0 ? 0 : (1 << i);
}

// The blueprint topic of a question for one license ('' when the question does
// not belong to that license at all).
function questionTopic(id, license) {
  var entry = questionIndexEntry(id);
  return (entry && entry.c[String(license)]) || '';
}

// The one place that reads the answer key. null means "the key cannot answer for
// this id in this language" — callers must treat that as NOT VERIFIABLE and
// never as index 0, which is how the false 0/30 of 03/06/2026 happened.
function answerKeyIndex(id, lang) {
  if (typeof lookupCorrectIndex !== 'function') return null;
  var idx = lookupCorrectIndex(Number(id), String(lang || 'he').toLowerCase());
  if (idx === null || idx === undefined) return null;
  var n = Number(idx);
  return (isFinite(n) && n >= 0 && n < QUESTION_ANSWER_COUNT) ? n : null;
}

// ids of one license+language grouped by blueprint topic. One pass over 1,700
// index entries — measured in microseconds, so no cache (and no cache bug).
function indexIdsByTopic(license, lang) {
  var bit = questionLangBit(lang), lic = String(license), byTopic = {};
  for (var id in QUESTION_INDEX) {
    if (!Object.prototype.hasOwnProperty.call(QUESTION_INDEX, id)) continue;
    var entry = QUESTION_INDEX[id];
    if (!(entry.l & bit)) continue;
    var topic = entry.c[lic];
    if (!topic) continue;
    if (!byTopic[topic]) byTopic[topic] = [];
    byTopic[topic].push(Number(id));
  }
  return byTopic;
}

// Every id available for a license+language, whatever its topic.
function indexIdsFor(license, lang) {
  var byTopic = indexIdsByTopic(license, lang), out = [];
  for (var topic in byTopic) {
    if (!Object.prototype.hasOwnProperty.call(byTopic, topic)) continue;
    out = out.concat(byTopic[topic]);
  }
  return out;
}

function questionBankUnavailable(message, detail) {
  var err = new Error(message);
  err.code = 'bank_unavailable';
  err.detail = detail || '';
  return err;
}

// 30 ids per the license blueprint, as [{id, topic}] in random order.
// A question whose answer key is missing for THIS language is skipped here:
// registering it would mean scoring it later against a key that does not exist.
function drawExamIds(license, lang) {
  var blueprint = EXAM_STRUCTURE_SERVER[String(license)];
  if (!blueprint) throw questionBankUnavailable('דרגה לא מוכרת: ' + license, String(license));
  var byTopic = indexIdsByTopic(license, lang), picked = [], used = {};
  for (var topic in blueprint) {
    if (!Object.prototype.hasOwnProperty.call(blueprint, topic)) continue;
    var need = blueprint[topic], pool = shuffleArrayServer(byTopic[topic] || []), got = 0;
    for (var i = 0; i < pool.length && got < need; i++) {
      var id = pool[i];
      if (used[id] || answerKeyIndex(id, lang) === null) continue;
      used[id] = true;
      picked.push({ id: id, topic: topic });
      got++;
    }
    if (got < need) {
      throw questionBankUnavailable('אין מספיק שאלות בנושא ' + topic, topic + ' ' + got + '/' + need);
    }
  }
  return shuffleArrayServer(picked);
}

// A fresh display order for one question: a permutation of the answer positions.
function drawShuffleOrder() {
  var order = [];
  for (var i = 0; i < QUESTION_ANSWER_COUNT; i++) order.push(i);
  return shuffleArrayServer(order);
}

// Practice scores locally, so it needs the correct index for every language the
// question exists in, XOR-encoded exactly as the legacy questions.js did
// (ci = correctIndex ^ (id % 256)) so the client keeps a single decoder.
// Languages the question is not translated into are omitted — the answer key
// falls back to Hebrew, and for en/fr/es/ar that fallback is a different order.
function practiceCiByLang(id) {
  var entry = questionIndexEntry(id);
  if (!entry) return null;
  var out = {};
  for (var i = 0; i < QUESTION_LANGS.length; i++) {
    if (!(entry.l & (1 << i))) continue;
    var idx = answerKeyIndex(id, QUESTION_LANGS[i]);
    if (idx === null) continue;
    out[QUESTION_LANGS[i]] = idx ^ (Number(id) % 256);
  }
  return out;
}

// ---- bankGrant --------------------------------------------------------------
// The examiner tools still need question TEXTS: the commander's wrong-answer
// table (ids only, one language) and find_image.html (a full language bank, so
// that a search can see all of it). Since the texts stopped being public, an
// examiner gets the same kind of signed grant an examinee does — scoped to no
// id list, valid for a working day, and read straight from the Worker. The
// examiner token is checked by the router before this runs.
defineAction('bankGrant', { methods: ['GET'], auth: 'examiner', handler: handleBankGrant,
  rateLimit: { max: 30, windowSec: 60, id: function(p) { return normalizeId(p.examinerId); } } });
function handleBankGrant(p) {
  // normalizeId, not the raw field: the same examiner must produce the same
  // subject whether they typed leading zeros or not.
  var bank = bankGrantFor('examiner', null, 'ex:' + normalizeId(p.examinerId));
  if (!bank) return bankNotConfiguredResponse();
  return jsonResponse({ status: 'ok', bank: bank });
}
