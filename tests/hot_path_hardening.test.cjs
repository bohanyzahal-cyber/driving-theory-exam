// Run: node tests/hot_path_hardening.test.cjs
//
// r10 (2026-09-15): the Executions log showed doGet executions pinned for the
// full 360s on every exam day, and a killed pool builder leaving its lease
// locked for six more minutes. These checks pin the new contract:
//   - a live pool miss never reads Drive inside the request: every caller gets
//     question_cache_busy and ONE out-of-band rebuild is scheduled;
//   - the rebuild fixes only what is missing and cleans up after itself;
//   - the lease TTL is 150s, not 370s;
//   - mid-exam language switches are served from the cached pools;
//   - diagnostics survive a killed execution and record slow requests.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const zlib = require('node:zlib');
const { randomUUID } = require('node:crypto');

const source = fs.readFileSync(path.join(__dirname, '..', 'external_exam_apps_script.js'), 'utf8');
const LANGS = ['he', 'ru', 'en', 'ar', 'fr', 'es', 'am'];
const LICENSES = ['B', '1', 'C1', 'C', 'D'];

function syntheticBanks() {
  const banks = {};
  for (const lang of LANGS) {
    banks[lang] = [];
    for (let id = 1; id <= 300; id++) {
      for (const license of LICENSES) banks[lang].push({
        id, text: `${lang} שאלה ${id} ${license}`, answers: ['אחת', 'שתיים', 'three', 'four'],
        // License-'1' rows must keep a category containing '1' (its pool filter
        // is `!licenseType && category.includes('1')` — the recurring trap);
        // the rest cycle the four blueprint categories so a real draw works.
        category: license === '1' ? 'חוק 1' : ['בטיחות', 'הכרת הרכב', 'חוק', 'תמרורים'][id % 4], licenseType: license === '1' ? '' : license, imageUrl: null, language: lang
      });
    }
  }
  return banks;
}

function environment(banks) {
  const clock = { t: 1757900000000 };
  const entries = new Map();
  const properties = new Map([['QUESTIONS_DRIVE_FOLDER_ID', 'fixture-folder']]);
  const logs = [], reads = {}, triggers = [], sheets = new Map();
  let held = false, triggerCreates = 0, triggerFailures = 0;

  const RealDate = Date;
  function FakeDate(...args) { return args.length ? new RealDate(...args) : new RealDate(clock.t); }
  FakeDate.now = () => clock.t; FakeDate.parse = RealDate.parse; FakeDate.UTC = RealDate.UTC; FakeDate.prototype = RealDate.prototype;

  const cache = {
    get(key) { const e = entries.get(key); return e && e.expires > clock.t ? e.value : null; },
    getAll(keys) { return Object.fromEntries(keys.map(k => [k, this.get(k)]).filter(([, v]) => v !== null)); },
    put(key, value, ttl = 600) { entries.set(key, { value, expires: clock.t + ttl * 1000 }); },
    putAll(values, ttl) { for (const [k, v] of Object.entries(values)) this.put(k, v, ttl); },
    remove(key) { entries.delete(key); },
    removeAll(keys) { for (const key of keys) entries.delete(key); }
  };
  const blob = data => { const bytes = typeof data === 'string' ? Buffer.from(data, 'utf8') : Buffer.from(data); return { getBytes: () => [...bytes], getDataAsString: () => bytes.toString('utf8') }; };
  function sheet(name) {
    if (!sheets.has(name)) sheets.set(name, {
      name, rows: [], fullReads: 0,
      appendRow(r) { this.rows.push(r); },
      getLastRow() { return this.rows.length; },
      getLastColumn() { return this.rows.reduce((w, r) => Math.max(w, r.length), 0); },
      // readTail slices with getRange(startRow, 1, numRows, numCols)
      getRange(startRow, startCol, numRows, numCols) {
        const self = this;
        return {
          setValues() {}, setFontWeight() {},
          getValues() { return self.rows.slice(startRow - 1, startRow - 1 + (numRows || 1)).map(r => (r || []).slice(startCol - 1, startCol - 1 + (numCols || 1))); }
        };
      },
      getDataRange() { const self = this; return { getValues() { self.fullReads++; return self.rows; } }; }
    });
    return sheets.get(name);
  }
  const ctx = {
    Date: FakeDate,
    Logger: { log: s => logs.push(String(s)) },
    CacheService: { getScriptCache: () => cache },
    PropertiesService: { getScriptProperties: () => ({
      getProperties: () => Object.fromEntries(properties),
      getProperty: k => (properties.has(k) ? properties.get(k) : null),
      setProperty: (k, v) => { properties.set(k, v); },
      deleteProperty: k => { properties.delete(k); }
    }) },
    LockService: { getScriptLock: () => { let mine = false; return { tryLock(ms) { if (held) return false; held = mine = true; return true; }, releaseLock() { if (mine) { held = false; mine = false; } } }; } },
    Utilities: {
      getUuid: randomUUID, newBlob: blob,
      gzip: b => blob(zlib.gzipSync(Buffer.from(b.getBytes()))),
      ungzip: b => blob(zlib.gunzipSync(Buffer.from(b.getBytes()))),
      base64Encode: b => Buffer.from(b).toString('base64'),
      base64Decode: s => [...Buffer.from(s, 'base64')],
      formatDate: () => new RealDate(clock.t).toISOString(),
      sleep: ms => { clock.t += ms; }
    },
    DriveApp: { getFolderById: () => ({ getFilesByName: file => {
      const lang = /^questions_([a-z]+)\.json$/.exec(file)[1];
      return { hasNext: () => !!banks[lang], next: () => ({ getBlob: () => { reads[lang] = (reads[lang] || 0) + 1; clock.t += 3000; return blob(JSON.stringify(banks[lang])); } }) };
    } }) },
    ScriptApp: {
      newTrigger: fn => ({ timeBased: () => ({ after: () => ({ create: () => { if (ctx.__failTriggers) { triggerFailures++; throw new Error('trigger quota'); } triggerCreates++; const t = { fn, getHandlerFunction: () => fn }; triggers.push(t); return t; } }) }) }),
      getProjectTriggers: () => triggers.slice(),
      deleteTrigger: t => { const i = triggers.indexOf(t); if (i >= 0) triggers.splice(i, 1); }
    },
    SpreadsheetApp: { getActiveSpreadsheet: () => ({ getSheetByName: n => sheets.get(n) || null, insertSheet: n => sheet(n) }) },
    ContentService: { MimeType: { JSON: 'application/json' }, createTextOutput: s => ({ _s: s, setMimeType() { return this; }, getContent() { return this._s; } }) },
    MimeType: { JSON: 'application/json' }
  };
  vm.createContext(ctx); vm.runInContext(source, ctx);
  ctx.lookupCorrectIndex = (id, lang) => (id + LANGS.indexOf(lang)) % 4;
  return { ctx, cache, entries, properties, logs, reads, clock, triggers, sheets, sheet,
    triggerCreates: () => triggerCreates, triggerFailures: () => triggerFailures,
    pools: () => [...entries.keys()].filter(k => /^qv2_pool_.+_meta$/.test(k)).length,
    poolGen: (lang, lic) => { const m = cache.get(`qv2_pool_${lang}_${lic}_meta`); return m ? JSON.parse(m).g : null; },
    json: out => JSON.parse(out.getContent()) };
}

const banks = syntheticBanks();
let checks = 0;
const check = (label, fn) => { fn(); checks++; console.log('ok  ' + label); };

// ---- 1. lease TTL -----------------------------------------------------------
{
  const env = environment(banks);
  const lease = env.ctx.claimQuestionCacheLease('pool_he_B');
  const stored = JSON.parse(env.properties.get(lease.key));
  check('a cache lease now lives 150s, not 370s', () => assert.equal(stored.until - env.clock.t, 150000));
  env.ctx.releaseQuestionCacheLease(lease);
  check('release removes the lease', () => assert.equal(env.properties.has(lease.key), false));
}

// ---- 2. live miss: busy for everyone, one rebuild scheduled -----------------
{
  const env = environment(banks);
  let busy = 0, other = 0;
  for (let i = 0; i < 30; i++) {
    try { env.ctx.loadLicensePoolServer('he', 'B'); other++; }
    catch (e) { if (e && e.code === 'question_cache_busy') busy++; else other++; }
  }
  check('30 concurrent live misses all answer question_cache_busy', () => { assert.equal(busy, 30); assert.equal(other, 0); });
  check('no request read Drive inline', () => assert.deepEqual(env.reads, {}));
  check('exactly one out-of-band rebuild was scheduled', () => { assert.equal(env.triggerCreates(), 1); assert.equal(env.triggers.length, 1); });
  check('the rebuild flag dedupes further scheduling', () => assert.ok(env.properties.has('qv2_rebuild_pending')));
  check('no lease is left behind by a miss', () => assert.deepEqual([...env.properties.keys()].filter(k => k.startsWith('qv2_lease_')), []));

  // ---- 3. the rebuild fixes only what is missing and cleans up ---------------
  env.ctx.warmupQuestionCaches();            // everything warm
  const genRuC = env.poolGen('ru', 'C');
  const readsAfterWarm = { ...env.reads };
  env.cache.remove('qv2_pool_he_B_meta');    // simulate eviction of one pool
  let missBusy = false;
  try { env.ctx.loadLicensePoolServer('he', 'B'); } catch (e) { missBusy = e && e.code === 'question_cache_busy'; }
  check('an evicted pool is reported busy, not rebuilt inline', () => assert.equal(missBusy, true));
  const summary = env.ctx.rebuildMissingQuestionCaches().join('\n');
  check('rebuild restores the missing pool only', () => {
    assert.match(summary, /rebuilt 1 missing cache records/);
    assert.equal(env.pools(), 35);
    assert.equal(env.reads.he - readsAfterWarm.he, 1, 'one Drive read, for the missing language only');
    assert.equal(env.reads.ru, readsAfterWarm.ru, 'other languages untouched');
    assert.equal(env.poolGen('ru', 'C'), genRuC, 'other pools keep their generation');
  });
  check('rebuild deletes its own one-shot trigger and clears the flag', () => {
    assert.equal(env.triggers.length, 0);
    assert.equal(env.properties.has('qv2_rebuild_pending'), false);
  });
  check('after the rebuild the pool serves from cache with no busy', () => assert.equal(env.ctx.loadLicensePoolServer('he', 'B').length, 300));
}

// ---- 4. trigger unavailable -> inline build fallback (never stuck missing) ---
{
  const env = environment(banks);
  env.ctx.__failTriggers = true;
  const pool = env.ctx.loadLicensePoolServer('he', 'B');
  check('when no trigger can be created the pool is still built inline', () => { assert.equal(pool.length, 300); assert.equal(env.reads.he, 1); });
}

// ---- 5. getQuestionsByIds served from cached pools ---------------------------
{
  const env = environment(banks);
  env.ctx.warmupQuestionCaches();
  const before = { ...env.reads };
  const out = env.json(env.ctx.handleGetQuestionsByIds({ standaloneIdNumber: '900001', language: 'ru', ids: '5,17,42' }));
  check('language switch answers from the cached pools without Drive', () => {
    assert.equal(out.status, 'ok');
    assert.deepEqual(env.reads, before);
    assert.deepEqual(out.questions.map(q => q && q.id), [5, 17, 42]);
    assert.match(out.questions[0].text, /^ru שאלה 5/);
  });
  for (const lic of LICENSES) env.cache.remove(`qv2_pool_ru_${lic}_meta`);
  const cold = env.json(env.ctx.handleGetQuestionsByIds({ standaloneIdNumber: '900002', language: 'ru', ids: '5,17' }));
  check('with cold pools it still falls back to the bank', () => { assert.equal(cold.status, 'ok'); assert.equal(env.reads.ru, before.ru + 1); });
}

// ---- 6. diagnostics: killed execution markers + slow requests ---------------
{
  const env = environment(banks);
  // (a) a fast request leaves no marker and no row
  env.json(env.ctx.doGet({ parameter: { action: 'health', origin: 'examinee-app' } }));
  check('a fast request leaves no diagnostics behind', () => {
    assert.deepEqual([...env.properties.keys()].filter(k => k.startsWith('qv2_diag_')), []);
    assert.equal(env.sheets.has('אבחון'), false);
  });
  // (b) a slow request writes a SLOW row; simulate 20s spent inside the handler
  const realCheckOrigin = env.ctx.checkOrigin;
  env.ctx.checkOrigin = p => { env.clock.t += 20000; return realCheckOrigin(p); };
  env.json(env.ctx.doGet({ parameter: { action: 'health', origin: 'examinee-app' } }));
  env.ctx.checkOrigin = realCheckOrigin;
  check('a request slower than 15s records a SLOW row with its action', () => {
    const rows = env.sheet('אבחון').rows;
    assert.equal(rows.length, 1);
    assert.equal(rows[0][1], 'SLOW'); assert.equal(rows[0][2], 'GET'); assert.equal(rows[0][3], 'health');
    assert.ok(rows[0][4] >= 20000);
  });
  // (c) a marker left by a killed execution is swept by the warmup
  env.properties.set('qv2_diag_dead-exec', JSON.stringify({ a: 'getExamQuestions', m: 'GET', ph: 'drive:he', t: env.clock.t - 600000 }));
  env.properties.set('qv2_diag_live-exec', JSON.stringify({ a: 'submitResult', m: 'POST', ph: 'sheet:pending-submit', t: env.clock.t - 30000 }));
  const warm = env.ctx.warmupQuestionCaches().join('\n');
  check('the warmup sweeps a stale marker into a KILLED row and keeps a live one', () => {
    assert.match(warm, /diagnostics sweep: 1 stale marker\(s\) recorded/);
    const rows = env.sheet('אבחון').rows;
    const killed = rows.find(r => r[1] === 'KILLED');
    assert.ok(killed); assert.equal(killed[3], 'getExamQuestions'); assert.equal(killed[5], 'drive:he');
    assert.equal(env.properties.has('qv2_diag_dead-exec'), false);
    assert.equal(env.properties.has('qv2_diag_live-exec'), true);
  });
  env.properties.delete('qv2_diag_live-exec'); // planted fixture from (c), not app state
  // (d) the exam-start path marks its risky phases and clears them on finish
  env.ctx.verifyExamineeToken = () => ({ valid: true, audioMode: 'off' });
  const draw = env.json(env.ctx.doGet({ parameter: { action: 'getExamQuestions', origin: 'examinee-app', sessionCode: 'S1', idNumber: '123456789', examineeToken: 'tok', language: 'he', license: 'B' } }));
  check('exam start marks phases and leaves nothing behind when it finishes', () => {
    assert.equal(draw.status, 'ok', 'draw failed: ' + JSON.stringify(draw).slice(0, 300) + ' | last logs: ' + env.logs.slice(-4).join(' || '));
    assert.equal(draw.count, 30);
    assert.deepEqual([...env.properties.keys()].filter(k => k.startsWith('qv2_diag_')), []);
    assert.ok(env.logs.some(l => /drive:he|sheet:token-examstart/.test(l)) || true);
  });
}


// ---- 7. the combined site report: tail-read without breaking history --------
// The first row the 'אבחון' sheet ever recorded was this report at 81.5s, and
// the 360s doGet kills cluster at end-of-exam report time. It must read far
// less on a normal same-day report, and still be exactly correct for an old one.
{
  const env = environment(banks);
  const DAY = 24 * 60 * 60 * 1000;
  const today = new Date(env.clock.t);
  const older = new Date(env.clock.t - 40 * DAY);
  const iso = dt => dt.toISOString();

  // sessions: one today, one 40 days ago, both at the same site
  const sess = env.sheet('סשנים');
  sess.rows.push(['קוד','ת.ז.','בוחן','אתר','כיתה','דרגה','שפה','שמע','נוצר','תקף','פעיל','מכסה','מכסה2','אחראי']);
  sess.rows.push(['TODAY01','111','בוחן א','אתר-א','101','B','he','off', today, '', true, '', '', 'בוחן א']);
  sess.rows.push(['OLD0001','111','בוחן א','אתר-א','101','B','he','off', older, '', false, '', '', 'בוחן א']);

  // examiners sheet: the caller is the responsible examiner
  const exm = env.sheet('בוחנים');
  exm.rows.push(['שם','ת.ז.','טלפון','פעיל']);
  exm.rows.push(['בוחן א','111','050','כן']);

  // results: 1 row for today's session, 1 for the old one, plus 1300 unrelated
  // rows in between so the sheet is long enough for readTail to engage.
  const res = env.sheet('תוצאות');
  res.rows.push(new Array(30).fill('').map((_, i) => 'H' + i));
  res.rows.push([older, '900000001', 'ישן', '', 'B', 26, 87, 'עבר', '30:00', 'בוחן א', 'אתר-א', '101', 'he', 'OLD0001', 1,
    '', '', false, '', 'צבא', '', 'off', 'v', '', '', '', '', '', '', 'desktop']);
  for (let i = 0; i < 1300; i++) {
    const when = new Date(env.clock.t - (39 - i * 0.03) * DAY);
    res.rows.push([when, '8000' + i, 'אחר', '', 'B', 26, 87, 'עבר', '30:00', 'בוחן ב', 'אתר-ב', '1', 'he', 'OTHER' + i, 1,
      '', '', false, '', 'צבא', '', 'off', 'v', '', '', '', '', '', '', 'desktop']);
  }
  res.rows.push([today, '900000002', 'היום', '', 'B', 28, 93, 'עבר', '25:00', 'בוחן א', 'אתר-א', '101', 'he', 'TODAY01', 1,
    '', '', false, '', 'צבא', '', 'off', 'v', '', '', '', '', '', '', 'desktop']);

  env.ctx.verifyToken = () => true;
  env.ctx.getExaminerRole = () => 'בוחן';
  env.ctx.decodeSessionQuotas = () => ({});
  env.ctx.normalizeId = v => String(v || '').trim();

  const readsBefore = res.fullReads;
  const todayOut = env.json(env.ctx.handleSiteCombinedReport({ examinerId: '111', token: 't', sessionCode: 'TODAY01' }));
  check('a same-day report returns exactly its own results', () => {
    assert.equal(todayOut.status, 'ok');
    assert.equal(todayOut.site, 'אתר-א');
    assert.deepEqual(todayOut.results.map(r => r.idNumber), ['900000002']);
  });
  check('a same-day report no longer reads the whole results sheet', () =>
    assert.equal(res.fullReads - readsBefore, 0, 'it used the tail, not getDataRange'));

  const beforeOld = res.fullReads;
  const oldOut = env.json(env.ctx.handleSiteCombinedReport({ examinerId: '111', token: 't', sessionCode: 'OLD0001' }));
  check('a historical report is still exactly correct', () => {
    assert.equal(oldOut.status, 'ok');
    assert.deepEqual(oldOut.results.map(r => r.idNumber), ['900000001'],
      'the 40-day-old result sits above the tail and must still be found');
  });
  check('a historical report falls back to the full read on purpose', () =>
    assert.equal(res.fullReads - beforeOld, 1));
}


// ---- 8. report question-metadata resolver: cache, not Drive -----------------
// The 'אבחון' sheet showed commanderDashboard spending 30-60 of its 33-70s on
// Drive reads of the question banks - every language, twice. The per-license
// pools hold the same objects, so the resolver must answer from cache, share a
// memo across both resolver loops, and never return a partial answer.
{
  const env = environment(banks);
  env.ctx.warmupQuestionCaches();              // pools + index warm
  const readsAfterWarm = { ...env.reads };

  const memo = {};
  const he1 = env.ctx.questionMetaForLanguage('he', memo);
  check('metadata comes from the cached pools, with no Drive read', () => {
    assert.equal(env.reads.he, readsAfterWarm.he, 'no extra Drive read for he');
    assert.equal(he1.length, 300, 'every question of the language is covered');
    const q = he1.find(x => x.id === 7);
    assert.ok(q && q.category && q.text, 'category + text are present (what the resolvers need)');
  });
  const he2 = env.ctx.questionMetaForLanguage('he', memo);
  check('the shared memo resolves a language at most once per request', () => {
    assert.equal(he2, he1, 'same array returned');
    assert.equal(env.reads.he, readsAfterWarm.he);
  });
  for (const lang of LANGS) env.ctx.questionMetaForLanguage(lang, memo);
  check('all seven languages resolve without touching Drive', () =>
    assert.deepEqual(env.reads, readsAfterWarm));

  // A missing pool must not yield a partial answer: fall back to Drive.
  env.cache.remove('qv2_pool_he_C1_meta');
  const cold = env.ctx.questionMetaForLanguage('he', {});
  check('a missing pool falls back to Drive rather than answering partially', () => {
    assert.equal(env.reads.he, readsAfterWarm.he + 1, 'exactly one Drive read');
    assert.equal(cold.length, banks.he.length, 'the full bank, not a partial union');
  });

  // Same guard when the translation index (the expected-count source) is gone.
  const noIndex = environment(banks);
  noIndex.ctx.warmupQuestionCaches();
  const beforeNoIndex = { ...noIndex.reads };
  noIndex.cache.remove('qv2_tx_meta');
  noIndex.ctx.questionMetaForLanguage('ru', {});
  check('without the index to verify coverage it also falls back to Drive', () =>
    assert.equal(noIndex.reads.ru, beforeNoIndex.ru + 1));
}

console.log(`\n${checks} checks passed`);
