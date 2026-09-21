// shared/bank.js — the ONE client-side reader of the static question bank.
// Loaded by examinee.html, student.html, exam.html, examiner.html (commander
// "top wrong" questions) and find_image.html. ES5 only.
//
// The bank lives on GitHub Pages as bank/<lang>.json (one entry per question id,
// no correct-answer index) plus bank/manifest.json (content hashes). Built by
// tools/build_bank.js from deployment/generated/questions_<lang>.json. Entry:
//   { id, t: text, a: [answers], i: 'TQ_PIC_xxx.jpg' | '', v?: { <license>: { t?, a? } } }
// `v` carries the few license-specific wordings (Hebrew ids 1276, 829, 124, 126,
// 621 …); get(id, lang, license) applies it.
//
// Loading: fetch('bank/<lang>.json?v=<sha>') — the sha comes from the manifest, so
// a new bank build is a new URL and an old one stays valid in every cache. Offline
// the service worker answers from its cache (it ignores the query string for
// bank/ files). Everything is memoised in memory once loaded.
(function (global) {
  'use strict';

  var BASE = 'bank/';
  var LANGS = ['he', 'ru', 'en', 'ar', 'fr', 'es', 'am'];
  var manifest = null, manifestPromise = null;
  var banks = {};          // lang -> { byId: {id: entry}, list: [entries] }
  var loading = {};        // lang -> Promise

  function configure(opts) {
    if (opts && opts.base) BASE = String(opts.base).replace(/\/?$/, '/');
  }

  function fetchJson(url) {
    return fetch(url, { cache: 'no-store' }).then(function (r) {
      if (!r.ok) throw new Error('HTTP ' + r.status + ' for ' + url);
      return r.json();
    });
  }

  function loadManifest() {
    if (manifest) return Promise.resolve(manifest);
    if (!manifestPromise) {
      manifestPromise = fetchJson(BASE + 'manifest.json').then(function (m) {
        if (!m || !m.langs) throw new Error('bad manifest');
        manifest = m;
        return m;
      }).catch(function () { manifestPromise = null; return null; });   // offline: fall back to unversioned URLs
    }
    return manifestPromise;
  }

  function bankUrl(lang, m) {
    var sha = m && m.langs && m.langs[lang] && m.langs[lang].sha;
    return BASE + lang + '.json' + (sha ? '?v=' + sha : '');
  }

  function index(list) {
    var byId = {};
    for (var i = 0; i < list.length; i++) byId[list[i].id] = list[i];
    return { byId: byId, list: list };
  }

  // Resolves to the bank of `lang` ({ byId, list }); rejects only when neither the
  // network nor the service-worker cache has it.
  function load(lang) {
    lang = String(lang || 'he').toLowerCase();
    if (LANGS.indexOf(lang) === -1) return Promise.reject(new Error('unknown language ' + lang));
    if (banks[lang]) return Promise.resolve(banks[lang]);
    if (!loading[lang]) {
      loading[lang] = loadManifest().then(function (m) {
        return fetch(bankUrl(lang, m), { cache: 'default' }).then(function (r) {
          if (!r.ok) throw new Error('HTTP ' + r.status);
          return r.json();
        });
      }).then(function (list) {
        if (!Array.isArray(list) || !list.length) throw new Error('empty bank ' + lang);
        banks[lang] = index(list);
        return banks[lang];
      }).catch(function (err) { delete loading[lang]; throw err; });
    }
    return loading[lang];
  }

  function has(lang) { return !!banks[String(lang || '').toLowerCase()]; }

  // Synchronous read for a loaded language. Returns null when the language is not
  // loaded or the id is absent (the caller keeps the previous language's text).
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

  // Fire-and-forget warm-up of several languages (the examinee waits minutes for
  // approval — that is when the banks should arrive, not at "start exam").
  function prefetch(langs) {
    (langs || LANGS).forEach(function (l) { load(l).catch(function () {}); });
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
    configure: configure,
    loadManifest: loadManifest,
    load: load,
    has: has,
    get: get,
    imageUrl: imageUrl,
    prefetch: prefetch,
    search: search,
    build: function () { return manifest ? manifest.build : ''; },
    // test hook
    _reset: function () { manifest = null; manifestPromise = null; banks = {}; loading = {}; }
  };
})(typeof window !== 'undefined' ? window : this);
