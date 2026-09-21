// shared/bank.js — the ONE client-side reader of the question bank.
// Loaded by examinee.html, student.html, exam.html, examiner.html (commander
// "top wrong" questions) and find_image.html. ES5 only.
//
// The texts are NOT public. They live as private static assets of the
// session-gateway Worker, which serves a device only the questions that device
// was actually issued, and only against a signed grant. The grant is an opaque
// `<payload>.<sig>` string: the server signs it (startExam / startPractice /
// bankGrant), the Worker verifies it, and nothing in this file — or anywhere
// else on the client — reads or produces one.
//
//   GET <url>/v1/bank?grant=…[&ids=1,2&langs=he,en]
//     -> { status:'ok', build, questions:[{id, l:{he:{t,a,i,v?}, ru:{…}, …}}], missing:[…] }
//   GET <url>/v1/bank/full?grant=…&lang=he           (examiner scope only)
//     -> [{id, t, a, i, v?}, …]  — the whole bank of one language, so search works
//
// Every answer is ingested into banks[lang].byId[id] = {id, t, a, i, v}, which is
// the shape get(id, lang, license) has always returned. `v` carries the few
// license-specific wordings (Hebrew ids 1276, 829, 124, 126, 621 …). A question
// arrives in ALL its languages at once, which is what keeps a mid-exam language
// switch local: after loadGrant nothing more is fetched.
(function (global) {
  'use strict';

  var LANGS = ['he', 'ru', 'en', 'ar', 'fr', 'es', 'am'];
  var LOAD_TIMEOUT_MS = 20000;   // a cold Worker on a classroom network
  var LOAD_ATTEMPTS = 3;
  var banks = {};                // lang -> { byId: {id: entry}, list: [entries] }

  // 1.5 / 3 / 6 s between attempts.
  function retryWaitMs(attempt) { return 1500 * Math.pow(2, attempt - 1); }
  function wait(ms) { return new Promise(function (resolve) { setTimeout(resolve, ms); }); }

  // ExamTransport owns every bounded fetch in the client, but the bank Worker is
  // NOT the Apps Script backend: a slow Worker must never mark that backend
  // degraded, or every poll in the page would drop to the 30-60 s floor for
  // nothing. Hence the quiet variant. Pages that load bank.js on its own (tests,
  // find_image) fall back to a plain fetch.
  function fetchJson(url) {
    var T = global.ExamTransport;
    if (T && T.fetchJsonQuiet) return T.fetchJsonQuiet(url, { cache: 'no-store' }, LOAD_TIMEOUT_MS);
    return fetch(url, { cache: 'no-store' }).then(function (r) {
      if (!r.ok) throw new Error('HTTP ' + r.status);
      return r.json();
    });
  }

  // Three attempts on the ladder above. A refused grant (403) fails on the same
  // ladder as a network blip: the page has to ask the server for a fresh grant
  // either way, and retrying longer only keeps the examinee waiting.
  function fetchWithRetry(url) {
    var attempt = 0;
    function once() {
      attempt++;
      return fetchJson(url).catch(function (err) {
        if (attempt >= LOAD_ATTEMPTS) throw err;
        return wait(retryWaitMs(attempt)).then(once);
      });
    }
    return once();
  }

  // '' when the caller has no usable grant — the callers turn that into a
  // rejection rather than a request the Worker would refuse anyway.
  function endpoint(bank, path, params) {
    if (!bank || !bank.url || !bank.grant) return '';
    var qs = 'grant=' + encodeURIComponent(String(bank.grant));
    for (var k in params) {
      if (!Object.prototype.hasOwnProperty.call(params, k)) continue;
      if (params[k] === undefined || params[k] === null || params[k] === '') continue;
      qs += '&' + k + '=' + encodeURIComponent(String(params[k]));
    }
    return String(bank.url).replace(/\/+$/, '') + path + '?' + qs;
  }

  function bankFor(lang) {
    if (!banks[lang]) banks[lang] = { byId: {}, list: [] };
    return banks[lang];
  }

  function put(lang, entry) {
    var bank = bankFor(lang);
    if (bank.byId[entry.id]) {
      for (var i = 0; i < bank.list.length; i++) {
        if (bank.list[i].id === entry.id) { bank.list[i] = entry; break; }
      }
    } else {
      bank.list.push(entry);
    }
    bank.byId[entry.id] = entry;
  }

  // One record -> one entry per language it carries. A language with no text for
  // this id is simply absent, and bankEntry() falls back to Hebrew per question.
  // Returns how many languages were ingested.
  function ingest(record) {
    if (!record || record.id == null || !record.l) return 0;
    var langs = 0;
    for (var lang in record.l) {
      if (!Object.prototype.hasOwnProperty.call(record.l, lang)) continue;
      var src = record.l[lang];
      if (!src || !src.t || !Array.isArray(src.a)) continue;
      put(String(lang).toLowerCase(), { id: record.id, t: src.t, a: src.a, i: src.i || '', v: src.v || null });
      langs++;
    }
    return langs;
  }

  // Shared by loadGrant and loadIds. Resolves { build, count, missing }; rejects
  // only when all three attempts failed or the body is not the documented shape
  // (an error envelope from the Worker lands here too, with its code).
  function loadQuestions(url) {
    if (!url) return Promise.reject(new Error('bank grant missing'));
    return fetchWithRetry(url).then(function (body) {
      if (!body || body.status !== 'ok' || !Array.isArray(body.questions)) {
        throw new Error('bad bank answer' + (body && body.code ? ' (' + body.code + ')' : ''));
      }
      var missing = Array.isArray(body.missing) ? body.missing.slice() : [];
      var count = 0;
      for (var i = 0; i < body.questions.length; i++) {
        // An id that yielded no language at all is as absent as one the Worker
        // never found: the caller must not draw it as an empty question.
        if (ingest(body.questions[i])) count++;
        else if (body.questions[i] && body.questions[i].id != null) missing.push(body.questions[i].id);
      }
      return { build: String(body.build || ''), count: count, missing: missing };
    });
  }

  // Exam / practice: the grant itself names the ids, so the device gets exactly
  // what it was issued, in every language.
  function loadGrant(bank) { return loadQuestions(endpoint(bank, '/v1/bank', {})); }

  // Examiner scope: an explicit id list, optionally narrowed to some languages.
  function loadIds(bank, ids, langs) {
    return loadQuestions(endpoint(bank, '/v1/bank', {
      ids: (ids || []).join(','),
      langs: (langs || []).join(',')
    }));
  }

  // Examiner scope: replace one language's bank with the whole thing, which is
  // what makes search() see more than the questions this device was issued.
  function loadFull(bank, lang) {
    lang = String(lang || 'he').toLowerCase();
    var url = endpoint(bank, '/v1/bank/full', { lang: lang });
    if (!url) return Promise.reject(new Error('bank grant missing'));
    return fetchWithRetry(url).then(function (list) {
      if (!Array.isArray(list) || !list.length) throw new Error('empty bank ' + lang);
      var byId = {};
      for (var i = 0; i < list.length; i++) byId[list[i].id] = list[i];
      banks[lang] = { byId: byId, list: list };
      return { lang: lang, count: list.length };
    });
  }

  // ---- saving an exam's texts with the exam ----
  // Once an exam has started it runs entirely on the device — a promise that
  // cannot depend on a Worker still being reachable after a reload. These two
  // hand the ingested questions out in the Worker's own shape and take them
  // back through the same ingest, so a page can keep them next to the exam
  // state and resume with no network at all.
  function exportIds(ids) {
    var out = [];
    for (var i = 0; i < (ids || []).length; i++) {
      var id = ids[i], l = null;
      for (var lang in banks) {
        if (!Object.prototype.hasOwnProperty.call(banks, lang)) continue;
        var e = banks[lang].byId[id];
        if (!e) continue;
        if (!l) l = {};
        l[lang] = { t: e.t, a: e.a, i: e.i || '' };
        if (e.v) l[lang].v = e.v;
      }
      if (l) out.push({ id: id, l: l });
    }
    return out;
  }

  // Returns how many ids arrived with a usable text, so the caller can tell a
  // complete snapshot from a truncated or corrupted one.
  function importRecords(records) {
    if (!Array.isArray(records)) return 0;
    var count = 0;
    for (var i = 0; i < records.length; i++) if (ingest(records[i])) count++;
    return count;
  }

  function has(lang) {
    var bank = banks[String(lang || '').toLowerCase()];
    return !!(bank && bank.list.length);
  }

  // Synchronous read for a loaded language. Returns null when the language holds
  // nothing for this id (the caller falls back to Hebrew, then to a placeholder).
  function get(id, lang, license) {
    var bank = banks[String(lang || '').toLowerCase()];
    if (!bank) return null;
    var e = bank.byId[id];
    if (!e) return null;
    var out = { id: e.id, text: e.t, answers: e.a, image: e.i || '' };
    var v = e.v && license ? e.v[String(license)] : null;
    if (v) {
      if (v.t) out.text = v.t;
      if (v.a) out.answers = v.a;
    }
    return out;
  }

  // Same-origin image path for an entry (images/ is self-hosted in the repo) or ''.
  function imageUrl(entry) {
    var name = entry && (entry.image || entry.i);
    return name ? 'images/' + name : '';
  }

  // Substring search for the examiner's find-image utility: [{id, text, image}].
  function search(lang, query, limit) {
    var bank = banks[String(lang || '').toLowerCase()];
    if (!bank) return [];
    var q = String(query || '').trim().toLowerCase();
    if (!q) return [];
    var out = [];
    for (var i = 0; i < bank.list.length && out.length < (limit || 20); i++) {
      var e = bank.list[i];
      if (String(e.t).toLowerCase().indexOf(q) !== -1 || e.a.some(function (a) { return String(a).toLowerCase().indexOf(q) !== -1; })) {
        out.push({ id: e.id, text: e.t, answers: e.a, image: e.i || '' });
      }
    }
    return out;
  }

  global.QuestionBank = {
    LANGS: LANGS,
    loadGrant: loadGrant,
    loadIds: loadIds,
    loadFull: loadFull,
    exportIds: exportIds,
    importRecords: importRecords,
    has: has,
    get: get,
    imageUrl: imageUrl,
    search: search,
    // test hook
    _reset: function () { banks = {}; }
  };
})(typeof window !== 'undefined' ? window : this);
