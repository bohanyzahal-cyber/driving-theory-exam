// Gates on the static question bank and the server's id index.
//
// The point of this file: the exam draw moved from "the server reads the whole
// bank" to "the server reads an index of ids". That is only safe if the index
// selects exactly what filterByLicenseServer + dedupe used to select, and if
// every id the client can be handed really has a text, an image and an answer
// key in the language it will be answered in. Both server functions are copied
// here VERBATIM on purpose — the test must not share code with the builder.
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const crypto = require('node:crypto');
const { execFileSync } = require('node:child_process');

const ROOT = path.join(__dirname, '..');
const GENERATED = path.join(ROOT, 'deployment', 'generated');
const BANK_DIR = path.join(ROOT, 'bank');
const LANGS = ['he', 'ru', 'en', 'ar', 'fr', 'es', 'am'];
const LICENSES = ['B', '1', 'C1', 'C', 'D'];

// One language may make a question eligible for a license the others do not:
// only the English dump has a C1 row for id 14. The index is the UNION over
// languages (DESIGN §3.1), so this is expected — but ONLY this.
const UNION_ONLY = { C1: [14] };

// --- verbatim copies from external_exam_apps_script.js ---------------------
function classifyCategoryServer(cat) {
  var c = String(cat || '').trim();
  if (/ספציפי/.test(c)) return 'ספציפי';
  if (/בטיחות/.test(c)) return 'בטיחות';
  if (/הכרת הרכב/.test(c)) return 'הכרת הרכב';
  if (/חוק/.test(c)) return 'חוק';
  if (/תמרורים/.test(c)) return 'תמרורים';
  if (/זכות קדימה/.test(c)) return 'חוק';
  return '';
}
function filterByLicenseServer(pool, license) {
  return pool.filter(function (q) {
    var cat = String(q.category || '');
    if (/זכות קדימה/.test(cat)) return true;
    if (license === '1') {
      var lt = String(q.licenseType || '').trim();
      if (lt !== '' && lt !== 'N/A') return false;
      if (cat.indexOf('1') === -1) return false;
      return true;
    }
    if (license === 'C') {
      var lic = String(q.licenseType || '').trim();
      return lic === 'C' || lic === 'C/E' || lic === 'C+E' || lic === 'CE';
    }
    var lic2 = String(q.licenseType || '').trim();
    return lic2 === license;
  });
}
/** What a legacy exam draw really saw: filter, drop unusable, first row wins. */
function legacyPool(rows, license) {
  const seen = new Set(), out = [];
  for (const q of filterByLicenseServer(rows, license)) {
    if (!q.id || !Array.isArray(q.answers) || q.answers.length < 2) continue;
    if (seen.has(q.id)) continue;
    seen.add(q.id);
    out.push(q);
  }
  return out;
}

// --- fixtures --------------------------------------------------------------
if (!fs.existsSync(path.join(BANK_DIR, 'manifest.json'))) {
  execFileSync('node', [path.join(ROOT, 'tools', 'build_bank.js')], { cwd: ROOT, stdio: 'inherit' });
}

const read = file => JSON.parse(fs.readFileSync(file, 'utf8'));
const source = {};
const bank = {};
for (const lang of LANGS) {
  source[lang] = read(path.join(GENERATED, 'questions_' + lang + '.json'));
  bank[lang] = read(path.join(BANK_DIR, lang + '.json'));
}
const manifest = read(path.join(BANK_DIR, 'manifest.json'));
const index = read(path.join(ROOT, 'deployment', 'question_index.json'));
const imageFiles = new Set(fs.readdirSync(path.join(ROOT, 'images')));

const answerKey = (() => {
  const ctx = vm.createContext({});
  vm.runInContext(fs.readFileSync(path.join(ROOT, 'deployment', 'answer_key.gs'), 'utf8'), ctx);
  return ctx.ANSWER_KEY_BY_LANG;
})();

const entryMap = lang => new Map(bank[lang].map(e => [e.id, e]));
const rowsById = lang => {
  const map = new Map();
  for (const row of source[lang]) {
    if (!map.has(row.id)) map.set(row.id, []);
    map.get(row.id).push(row);
  }
  return map;
};
const licenseKeyOf = row => {
  const lt = String(row.licenseType || '').trim();
  return (lt === '' || lt === 'N/A') ? '1' : lt;
};
const sigOf = (text, answers) => JSON.stringify([String(text), answers]);

// --- the gates -------------------------------------------------------------

test('every bank id keeps its own-language answer key, pointing inside the answers', () => {
  for (const lang of LANGS) {
    const key = answerKey[lang];
    assert.ok(key, 'no answer key for ' + lang);
    for (const entry of bank[lang]) {
      assert.ok(Object.prototype.hasOwnProperty.call(key, entry.id),
        lang + ' has no own-language key for id ' + entry.id);
      const idx = key[entry.id];
      assert.ok(Number.isInteger(idx) && idx >= 0 && idx < entry.a.length,
        lang + ' id ' + entry.id + ': key index ' + idx + ' outside 0..' + (entry.a.length - 1));
      for (const [license, variant] of Object.entries(entry.v || {})) {
        if (!variant.a) continue;
        assert.ok(idx < variant.a.length,
          lang + ' id ' + entry.id + ' variant ' + license + ': key index ' + idx + ' outside the variant');
      }
    }
  }
});

test('every question carries exactly four answers', () => {
  // Measured, not assumed: all 7 dumps hold 4 answers in every single row, so
  // "at least 2" (what the server tolerates) understates the real invariant.
  for (const lang of LANGS) {
    for (const entry of bank[lang]) {
      assert.equal(entry.a.length, 4, lang + ' id ' + entry.id + ' has ' + entry.a.length + ' answers');
      for (const variant of Object.values(entry.v || {})) {
        if (variant.a) assert.equal(variant.a.length, 4, lang + ' id ' + entry.id + ' variant');
      }
    }
  }
});

test('every image reference resolves to a self-hosted file', () => {
  for (const lang of LANGS) {
    for (const entry of bank[lang]) {
      if (!entry.i) continue;
      assert.ok(!/[\\/]/.test(entry.i), lang + ' id ' + entry.id + ': ' + entry.i + ' is not a basename');
      assert.ok(imageFiles.has(entry.i), lang + ' id ' + entry.id + ': images/' + entry.i + ' is missing');
    }
  }
});

test('no source row carries a non-http image, and id 120 carries none at all', () => {
  for (const lang of LANGS) {
    for (const row of source[lang]) {
      const url = String(row.imageUrl || '');
      assert.ok(url === '' || /^https?:\/\//.test(url),
        lang + ' id ' + row.id + ' has imageUrl "' + url + '" — run tools/fix_bank_images.js');
    }
    assert.equal(entryMap(lang).get(120).i, '', lang + ' id 120 must have no image');
  }
  assert.equal(index[120].img, 0);
});

test('the manifest describes the files that are actually on disk', () => {
  const shas = [];
  for (const lang of LANGS) {
    const bytes = fs.readFileSync(path.join(BANK_DIR, lang + '.json'));
    const sha = crypto.createHash('sha1').update(bytes).digest('hex');
    shas.push(sha);
    assert.equal(manifest.langs[lang].sha, sha, lang + ': manifest sha is stale');
    assert.equal(manifest.langs[lang].bytes, bytes.length, lang + ': manifest byte count is stale');
    assert.equal(manifest.langs[lang].count, bank[lang].length, lang + ': manifest id count is stale');
  }
  assert.equal(manifest.build, crypto.createHash('sha1').update(shas.join('')).digest('hex'));
  assert.deepEqual(Object.keys(manifest.langs), LANGS, 'language order is load-bearing (the index bitmask)');
});

test('each bank holds one sorted entry per unique source id', () => {
  for (const lang of LANGS) {
    const ids = bank[lang].map(e => e.id);
    assert.deepEqual(ids, [...ids].sort((a, b) => a - b), lang + ' is not sorted by id');
    assert.equal(new Set(ids).size, ids.length, lang + ' repeats an id');
    assert.deepEqual(new Set(ids), new Set(source[lang].map(r => r.id)), lang + ' lost or invented ids');
  }
  const indexIds = Object.keys(index).map(Number);
  assert.deepEqual(indexIds, [...indexIds].sort((a, b) => a - b), 'index keys are not sorted numerically');
  const union = new Set(LANGS.flatMap(l => bank[l].map(e => e.id)));
  assert.deepEqual(new Set(indexIds), union, 'the index must cover every id of every language');
});

test('applying the variant map reproduces every source row exactly', () => {
  for (const lang of LANGS) {
    const entries = entryMap(lang);
    for (const [id, rows] of rowsById(lang)) {
      const entry = entries.get(id);
      const canonSig = sigOf(entry.t, entry.a);
      const counts = new Map();
      for (const row of rows) {
        const sig = sigOf(row.text, row.answers);
        counts.set(sig, (counts.get(sig) || 0) + 1);
        const variant = (entry.v || {})[licenseKeyOf(row)] || {};
        assert.equal(variant.t === undefined ? entry.t : variant.t, String(row.text),
          lang + ' id ' + id + ' license ' + licenseKeyOf(row) + ': text not reproducible');
        assert.deepEqual(variant.a === undefined ? entry.a : variant.a, row.answers,
          lang + ' id ' + id + ' license ' + licenseKeyOf(row) + ': answers not reproducible');
      }
      assert.ok(counts.has(canonSig), lang + ' id ' + id + ': canonical row is not one of the source rows');
      assert.equal(counts.get(canonSig), Math.max(...counts.values()),
        lang + ' id ' + id + ': canonical row is not the majority row');
      if (counts.size === 1) {
        assert.equal(entry.v, undefined, lang + ' id ' + id + ' has a variant map but all rows agree');
      } else {
        assert.ok(entry.v, lang + ' id ' + id + ' rows differ but no variant map');
        for (const [license, variant] of Object.entries(entry.v)) {
          assert.ok(LICENSES.includes(license), lang + ' id ' + id + ': unknown license key ' + license);
          assert.ok(Object.keys(variant).length, lang + ' id ' + id + ' ' + license + ': empty variant');
          assert.notEqual(sigOf(variant.t === undefined ? entry.t : variant.t,
            variant.a === undefined ? entry.a : variant.a), canonSig,
            lang + ' id ' + id + ' ' + license + ': variant equals the canonical row');
        }
      }
    }
  }
});

test('only the Hebrew dump has per-license variants, on the nine known rows', () => {
  const variantRows = lang => bank[lang].reduce((n, e) => n + Object.keys(e.v || {}).length, 0);
  assert.equal(variantRows('he'), 9);
  assert.deepEqual(bank.he.filter(e => e.v).map(e => e.id), [124, 125, 126, 128, 621, 829, 1276]);
  for (const lang of LANGS.filter(l => l !== 'he')) {
    assert.equal(variantRows(lang), 0, lang + ' grew per-license variants — translation drift?');
  }
});

test('the language bitmask names exactly the languages that hold the id', () => {
  for (const [id, record] of Object.entries(index)) {
    let expected = 0;
    LANGS.forEach((lang, bit) => { if (entryMap(lang).has(Number(id))) expected |= (1 << bit); });
    assert.equal(record.l, expected, 'id ' + id + ': mask ' + record.l + ' should be ' + expected);
    assert.ok(record.l > 0, 'id ' + id + ' exists in no language');
  }
  assert.equal(index[907].l & 1 << LANGS.indexOf('ru'), 0, 'id 907 is still missing from the Russian dump');
});

test('the index selects exactly what the legacy server selected, per language and license', () => {
  const sizes = {};
  for (const license of LICENSES) {
    for (const lang of LANGS) {
      const bit = 1 << LANGS.indexOf(lang);
      const pool = legacyPool(source[lang], license);
      const legacy = new Map(pool.map(q => [q.id, classifyCategoryServer(q.category)]));
      const derived = Object.keys(index)
        .filter(id => index[id].c[license] && (index[id].l & bit))
        .map(Number);

      sizes[lang] = sizes[lang] || {};
      sizes[lang][license] = pool.length;

      const missing = [...legacy.keys()].filter(id => !derived.includes(id));
      assert.deepEqual(missing, [], lang + '/' + license + ': the index lost ids the server would draw');

      const extra = derived.filter(id => !legacy.has(id)).sort((a, b) => a - b);
      const allowed = (UNION_ONLY[license] || []).filter(id => entryMap(lang).has(id) && !legacy.has(id));
      assert.deepEqual(extra, allowed,
        lang + '/' + license + ': unexpected union-only ids ' + JSON.stringify(extra));

      for (const [id, topic] of legacy) {
        assert.ok(topic, lang + '/' + license + ' id ' + id + ' classifies to nothing');
        assert.equal(index[id].c[license], topic,
          lang + '/' + license + ' id ' + id + ': topic ' + index[id].c[license] + ' should be ' + topic);
      }
    }
  }
  // The five numbers the whole exam blueprint rests on.
  assert.deepEqual(sizes.he, { B: 1253, '1': 1118, C1: 1361, C: 1187, D: 1303 });
});

test('every license the blueprint draws from has enough questions per topic', () => {
  const BLUEPRINT = {
    B: { 'בטיחות': 7, 'הכרת הרכב': 7, 'חוק': 7, 'תמרורים': 9 },
    '1': { 'בטיחות': 5, 'הכרת הרכב': 5, 'חוק': 6, 'תמרורים': 6, 'ספציפי': 8 },
    C1: { 'בטיחות': 5, 'הכרת הרכב': 5, 'חוק': 5, 'תמרורים': 5, 'ספציפי': 10 },
    C: { 'בטיחות': 5, 'הכרת הרכב': 4, 'חוק': 3, 'תמרורים': 4, 'ספציפי': 14 },
    D: { 'בטיחות': 4, 'הכרת הרכב': 2, 'חוק': 5, 'תמרורים': 4, 'ספציפי': 15 }
  };
  for (const [license, topics] of Object.entries(BLUEPRINT)) {
    assert.equal(Object.values(topics).reduce((a, b) => a + b, 0), 30, license + ' is not a 30-question exam');
    for (const lang of LANGS) {
      const bit = 1 << LANGS.indexOf(lang);
      const available = {};
      for (const [id, record] of Object.entries(index)) {
        if (!record.c[license] || !(record.l & bit)) continue;
        available[record.c[license]] = (available[record.c[license]] || 0) + 1;
      }
      for (const [topic, needed] of Object.entries(topics)) {
        assert.ok((available[topic] || 0) >= needed,
          lang + '/' + license + '/' + topic + ': only ' + (available[topic] || 0) + ' of ' + needed);
      }
    }
  }
});
