#!/usr/bin/env node
/**
 * build_bank.js — builds the STATIC public question bank and the server index.
 *
 * Inputs  (source of truth, gitignored, == the Drive copies):
 *   deployment/generated/questions_<lang>.json   — one row per question×license
 * Outputs:
 *   bank/<lang>.json            — one entry per UNIQUE id: {id,t,a,i} (+v)
 *   bank/manifest.json          — sha per language, used for ?v= cache-busting
 *   deployment/question_index.json — {id: {c:{license:topic}, l:mask, img}}
 *
 * WHY this shape:
 *  - The texts are already public (gov.il publishes the Hebrew bank), so the
 *    client can hold them; what stays on the server is the ANSWER KEY and this
 *    id index. That removes Drive, the caches, the warmups and the leases from
 *    the hot path (DESIGN_2026-09-21 §3.1).
 *  - The dumps repeat a question once per license. Only 9 Hebrew rows really
 *    differ between licenses, so the bank keeps ONE canonical entry per id and
 *    a tiny `v` map for the exceptions — the client applies it by license.
 *  - The index reproduces the legacy selection exactly: eligibility per license
 *    is `filterByLicenseServer` + dedupe-by-first-row, the topic is
 *    `classifyCategoryServer` of that same row. tests/bank_invariants.test.cjs
 *    re-derives both from the raw JSON with a verbatim copy of the server
 *    functions and demands set equality.
 *
 * Deterministic: same input bytes -> identical bank/<lang>.json bytes.
 * Usage: node tools/build_bank.js
 */
'use strict';
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const zlib = require('zlib');

const ROOT = path.join(__dirname, '..');
const GENERATED = path.join(ROOT, 'deployment', 'generated');
const IMAGES = path.join(ROOT, 'images');
const BANK_DIR = path.join(ROOT, 'bank');
const INDEX_FILE = path.join(ROOT, 'deployment', 'question_index.json');

// Fixed order — bit 0 = he … bit 6 = am in the index language mask.
const LANGS = ['he', 'ru', 'en', 'ar', 'fr', 'es', 'am'];
const LICENSES = ['B', '1', 'C1', 'C', 'D'];

// Legacy pool sizes measured on the Hebrew dump (21/09/2026). A change here is
// either a data edit or a builder bug — both must be looked at, never ignored.
const EXPECTED_HE_POOLS = { B: 1253, '1': 1118, C1: 1361, C: 1187, D: 1303 };

// ---------------------------------------------------------------------------
// Legacy selection semantics — must stay byte-for-byte equivalent to
// classifyCategoryServer / filterByLicenseServer in external_exam_apps_script.js.
// ---------------------------------------------------------------------------
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

/** The rows an exam draw would really see: license filter, then first row wins. */
function legacyPool(rows, license) {
  const seen = new Set();
  const out = [];
  for (const q of filterByLicenseServer(rows, license)) {
    if (!q.id || !Array.isArray(q.answers) || q.answers.length < 2) continue;
    if (seen.has(q.id)) continue;
    seen.add(q.id);
    out.push(q);
  }
  return out;
}

// ---------------------------------------------------------------------------
// Bank entries
// ---------------------------------------------------------------------------
const fail = msg => { throw new Error('build_bank: ' + msg); };

/** '' and 'N/A' mean "the motorcycle/1 rows" in the dumps. */
function licenseKey(row) {
  const lt = String(row.licenseType || '').trim();
  return (lt === '' || lt === 'N/A') ? '1' : lt;
}

function imageBasename(url) {
  const u = String(url || '');
  if (!u) return '';
  if (!/^https?:\/\//.test(u)) fail('non-http imageUrl "' + u + '" — run tools/fix_bank_images.js');
  return u.split('/').pop().split('?')[0];
}

/** Canonical = the (text, answers) pair most license rows agree on; ties: first. */
function pickCanonical(rows) {
  const groups = new Map();
  rows.forEach((row, i) => {
    const sig = JSON.stringify([row.text, row.answers]);
    const g = groups.get(sig);
    if (g) g.count++;
    else groups.set(sig, { count: 1, first: i, row });
  });
  let best = null;
  for (const g of groups.values()) {
    if (!best || g.count > best.count || (g.count === best.count && g.first < best.first)) best = g;
  }
  return best.row;
}

function buildEntry(id, rows, stats) {
  const canon = pickCanonical(rows);
  const entry = { id, t: String(canon.text), a: canon.answers.slice(), i: imageBasename(canon.imageUrl) };
  if (entry.a.length < 2) fail('id ' + id + ' has fewer than 2 answers');
  if (entry.i && !stats.imageFiles.has(entry.i)) fail('id ' + id + ' points at missing image images/' + entry.i);
  stats.minAnswers = Math.min(stats.minAnswers, entry.a.length);
  stats.maxAnswers = Math.max(stats.maxAnswers, entry.a.length);

  // 23 ids carry a second licenseType-'' row whose only difference is the
  // category ('מתן זכות קדימה', which the license filter lets through for every
  // license). First row wins, exactly like the legacy dedupe — but if two rows
  // of one license ever disagreed on the TEXT there would be no right answer,
  // so that fails the build.
  const variants = {};
  const seenKeys = new Map();
  for (const row of rows) {
    const key = licenseKey(row);
    const sig = JSON.stringify([row.text, row.answers]);
    if (seenKeys.has(key)) {
      if (seenKeys.get(key) !== sig) fail('id ' + id + ' has two different rows for license ' + key);
      continue;
    }
    seenKeys.set(key, sig);
    const diff = {};
    if (String(row.text) !== entry.t) diff.t = String(row.text);
    if (JSON.stringify(row.answers) !== JSON.stringify(entry.a)) diff.a = row.answers.slice();
    if (Object.keys(diff).length) variants[key] = diff;
  }
  if (Object.keys(variants).length) {
    entry.v = variants;
    stats.variantRows += Object.keys(variants).length;
    stats.variantIds.add(id);
  }
  return entry;
}

function buildLangBank(rows, stats) {
  const byId = new Map();
  for (const row of rows) {
    if (!byId.has(row.id)) byId.set(row.id, []);
    byId.get(row.id).push(row);
  }
  return [...byId.keys()].sort((a, b) => a - b).map(id => buildEntry(id, byId.get(id), stats));
}

// ---------------------------------------------------------------------------
// Server index
// ---------------------------------------------------------------------------
/**
 * Eligibility is the UNION over languages: a question the English dump marks as
 * C1 is offered to every language, because the text exists everywhere and the
 * answer key is per language anyway (id 14 is the only such case today). The
 * language mask `l` is what keeps a draw honest — a language that lacks the id
 * (ru lacks 907) is never handed it.
 */
function buildIndex(banks, entriesByLang) {
  const index = new Map();
  const poolSizes = {};
  const entryOf = id => {
    for (const lang of LANGS) {
      const e = entriesByLang[lang].get(id);
      if (e) return e;
    }
    return null;
  };

  for (const lang of LANGS) {
    poolSizes[lang] = {};
    const bit = 1 << LANGS.indexOf(lang);
    for (const row of banks[lang]) {
      if (!index.has(row.id)) index.set(row.id, { c: {}, l: 0, img: 0 });
      index.get(row.id).l |= bit;
    }
    for (const license of LICENSES) {
      const pool = legacyPool(banks[lang], license);
      poolSizes[lang][license] = pool.length;
      for (const q of pool) {
        const topic = classifyCategoryServer(q.category);
        if (!topic) continue; // unusable for the blueprint — never drawn
        const rec = index.get(q.id);
        if (rec.c[license] && rec.c[license] !== topic) {
          fail('id ' + q.id + ' license ' + license + ' classifies as both "' +
               rec.c[license] + '" and "' + topic + '" (' + lang + ')');
        }
        rec.c[license] = topic;
      }
    }
  }

  const out = {};
  for (const id of [...index.keys()].sort((a, b) => a - b)) {
    const rec = index.get(id);
    if (!Object.keys(rec.c).length) fail('id ' + id + ' belongs to no license');
    const entry = entryOf(id);
    const c = {};
    for (const license of LICENSES) if (rec.c[license]) c[license] = rec.c[license]; // stable key order
    out[id] = { c, l: rec.l, img: entry && entry.i ? 1 : 0 };
  }
  return { index: out, poolSizes };
}

// ---------------------------------------------------------------------------
const sha1 = buf => crypto.createHash('sha1').update(buf).digest('hex');

function main() {
  const stats = {
    imageFiles: new Set(fs.readdirSync(IMAGES)),
    minAnswers: Infinity, maxAnswers: 0, variantRows: 0, variantIds: new Set()
  };
  const banks = {};
  const entriesByLang = {};
  const manifest = { build: '', generatedAt: new Date().toISOString(), langs: {} };
  const shas = [];

  fs.mkdirSync(BANK_DIR, { recursive: true });

  for (const lang of LANGS) {
    banks[lang] = JSON.parse(fs.readFileSync(path.join(GENERATED, 'questions_' + lang + '.json'), 'utf8'));
    const entries = buildLangBank(banks[lang], stats);
    entriesByLang[lang] = new Map(entries.map(e => [e.id, e]));
    const bytes = Buffer.from(JSON.stringify(entries), 'utf8');
    fs.writeFileSync(path.join(BANK_DIR, lang + '.json'), bytes);
    const sha = sha1(bytes);
    shas.push(sha);
    manifest.langs[lang] = { sha, count: entries.length, bytes: bytes.length };
    const gz = zlib.gzipSync(bytes, { level: 9 }).length;
    console.log(lang.padEnd(3), 'rows ' + String(banks[lang].length).padStart(5),
      '-> ids ' + String(entries.length).padStart(4),
      '| ' + String(bytes.length).padStart(7) + ' B',
      '| gzip ' + String(gz).padStart(6) + ' B',
      '| sha ' + sha.slice(0, 8));
  }
  manifest.build = sha1(shas.join(''));
  fs.writeFileSync(path.join(BANK_DIR, 'manifest.json'), JSON.stringify(manifest, null, 2) + '\n');

  const { index, poolSizes } = buildIndex(banks, entriesByLang);
  fs.writeFileSync(INDEX_FILE, JSON.stringify(index));

  // --- summary -------------------------------------------------------------
  console.log('');
  console.log('bank build ' + manifest.build.slice(0, 12) + ' | index ids ' + Object.keys(index).length +
    ' | answers per question ' + stats.minAnswers + '-' + stats.maxAnswers +
    ' | variant rows ' + stats.variantRows + ' on ids ' + [...stats.variantIds].sort((a, b) => a - b).join(','));
  console.log('eligible per license (legacy pool sizes):');
  console.log('     ' + LICENSES.map(l => l.padStart(6)).join(''));
  for (const lang of LANGS) {
    console.log('  ' + lang.padEnd(3) + LICENSES.map(l => String(poolSizes[lang][l]).padStart(6)).join(''));
  }
  const withImage = Object.values(index).filter(r => r.img).length;
  console.log('  index: ' + withImage + ' ids with an image, ' +
    LICENSES.map(l => l + ' ' + Object.values(index).filter(r => r.c[l]).length).join(' / ') + ' (union over languages)');

  for (const license of LICENSES) {
    if (poolSizes.he[license] !== EXPECTED_HE_POOLS[license]) {
      fail('Hebrew pool ' + license + ' is ' + poolSizes.he[license] + ', expected ' +
           EXPECTED_HE_POOLS[license] + ' — data changed? update EXPECTED_HE_POOLS deliberately');
    }
  }
  console.log('wrote bank/<lang>.json, bank/manifest.json, deployment/question_index.json');
}

main();
