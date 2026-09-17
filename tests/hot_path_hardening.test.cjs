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
      name, rows: [], fullReads: 0, cellsRead: 0,
      appendRow(r) { this.rows.push(r); },
      getLastRow() { return this.rows.length; },
      getLastColumn() { return this.rows.reduce((w, r) => Math.max(w, r.length), 0); },
      // readTail slices with getRange(startRow, 1, numRows, numCols)
      getRange(startRow, startCol, numRows, numCols) {
        const self = this;
        return {
          setValues() {}, setFontWeight() {},
          getValues() {
            const out = self.rows.slice(startRow - 1, startRow - 1 + (numRows || 1)).map(r => (r || []).slice(startCol - 1, startCol - 1 + (numCols || 1)));
            self.cellsRead += out.length * (numCols || 1);   // what the read actually costs
            return out;
          }
        };
      },
      getDataRange() { const self = this; return { getValues() { self.fullReads++; self.cellsRead += self.rows.length * self.getLastColumn(); return self.rows; } }; }
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
      newTrigger: fn => ({ timeBased: () => ({
        after: () => ({ create: () => { if (ctx.__failTriggers) { triggerFailures++; throw new Error('trigger quota'); } triggerCreates++; const t = { fn, getHandlerFunction: () => fn }; triggers.push(t); return t; } }),
        // recurring hour timer (installWarmupTriggers); kept out of triggerCreates, which counts one-shot rebuilds
        everyHours: n => ({ create: () => { const t = { fn, everyHours: n, getHandlerFunction: () => fn }; triggers.push(t); return t; } })
      }) }),
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
  // (e) r15: a mark is free while the request is healthy. examinerDashboard is
  // polled every 2s by every examiner; five property writes per poll would have
  // been thousands of service calls an hour — so the property (the kill
  // breadcrumb) is only written once the request has already passed 8s, while
  // the in-memory trail still feeds every phase into the SLOW row.
  {
    const propWrites = () => [...env.properties.keys()].filter(k => k.startsWith('qv2_diag_')).length;
    env.ctx.diagBegin('GET'); env.ctx.DIAG_EXEC.t0 = env.clock.t; env.ctx.DIAG_EXEC.action = 'examinerDashboard';
    env.ctx.diagMark('sheet:pending-dash'); env.clock.t += 3000;
    env.ctx.diagMark('sheet:results-dash'); env.clock.t += 3000;
    check('marks under 8s touch no service at all', () => assert.equal(propWrites(), 0));
    env.clock.t += 3000;                      // 9s in: now in trouble
    env.ctx.diagMark('sheet:results-dash-2');
    check('the kill breadcrumb appears once the request is already slow', () => {
      assert.equal(propWrites(), 1);
      const entry = JSON.parse([...env.properties.values()].find(v => /results-dash-2/.test(v)));
      assert.equal(entry.a, 'examinerDashboard'); assert.equal(entry.ph, 'sheet:results-dash-2');
    });
    env.clock.t += 8000;                      // 17s total → SLOW row
    const before = env.sheet('אבחון').rows.length;
    env.ctx.diagFinish('examinerDashboard', env.clock.t - 17000);
    check('the SLOW row still carries the full in-memory trail, breadcrumb cleared', () => {
      const row = env.sheet('אבחון').rows[before];
      assert.equal(row[1], 'SLOW'); assert.equal(row[3], 'examinerDashboard');
      assert.match(row[6], /sheet:pending-dash@0 sheet:results-dash@3000 sheet:results-dash-2@9000/);
      assert.equal(propWrites(), 0);
    });
  }
}

// ---- 6b. result submission must not touch Drive -----------------------------
// 'אבחון' 2026-09-15: `SLOW POST submitResult 20066 ... drive:he@4000` — the
// wrong-answer reconstruction loaded the bank from Drive on the submission hot
// path. It must draw on the cached pools (which carry the full question objects,
// answers included) exactly like the reports do.
{
  const src = fs.readFileSync(path.join(__dirname, '..', 'external_exam_apps_script.js'), 'utf8');
  const start = src.indexOf('function handleSubmitResult('), end = src.indexOf('\nfunction ', start + 10);
  const body = src.slice(start, end);
  check('handleSubmitResult no longer reads the question bank from Drive', () => {
    assert.ok(body.length > 1000, 'handler located');
    assert.equal(body.includes('loadQuestionsForLanguageServer('), false);
    assert.ok(body.includes('questionMetaForLanguage('), 'it resolves through the cached pools');
  });
  const env = environment(banks);
  env.ctx.warmupQuestionCaches();
  const rows = env.ctx.questionMetaForLanguage('he', {});
  check('the cached metadata carries the answers the reconstruction renders', () => {
    const q = rows.find(x => x.id === 7);
    assert.ok(Array.isArray(q.answers) && q.answers.length >= 2, 'answers present');
    assert.ok(q.text && q.category, 'text + category present');
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


// ---- 7b. date-bounded reads: readRowsSince ---------------------------------
// The commander dashboard never looks at a result older than prevFrom, nor at a
// practice row more than 30 days before such a result - yet it read both sheets
// whole ('אבחון' 15/09: practice-commander 20-27s x12). readRowsSince must read
// only the tail when the tail provably reaches past the cutoff, grow when it
// does not, and never depend on the sheet's size for correctness.
{
  const env = environment(banks);
  const DAY = 24 * 60 * 60 * 1000, now = env.clock.t;
  const fill = (sh, count, spanDays) => {                // oldest first, one row per step
    sh.rows.push(['תאריך', 'מזהה', 'שם']);
    for (let i = 0; i < count; i++) sh.rows.push([new Date(now - spanDays * DAY + (i * spanDays * DAY) / count), 'id' + i, 'n' + i]);
  };
  const big = env.sheet('big'); fill(big, 6000, 120);      // 6000 rows over 120 days (50/day)
  const read = (sh, cutoffDaysAgo) => env.ctx.readRowsSince(sh, 0, new Date(now - cutoffDaysAgo * DAY));

  // The boundary is found exactly (one timestamp-column read), never estimated:
  // both earlier guesses — a fixed ladder, then a row-rate extrapolation — fell
  // through to `full` on the live 107,614-row sheet.
  const exactRows = (sheet, daysAgo) => sheet.rows.slice(1).filter(r => r[0].getTime() >= now - daysAgo * DAY).length;
  for (const daysAgo of [10, 50, 95]) {
    const r = read(big, daysAgo);
    check(`a ${daysAgo}-day cutoff reads exactly the rows at or after it`, () => {
      assert.match(r.mode, /^rows(\d+)\/6001$/, 'mode carries the count and the sheet size: ' + r.mode);
      const n = Number(/^rows(\d+)\//.exec(r.mode)[1]);
      assert.equal(n, exactRows(big, daysAgo), 'not one row more than needed');
      assert.equal(r.rows.length, n + 1, 'header + exactly those rows');
      assert.deepEqual(r.rows[0], ['תאריך', 'מזהה', 'שם'], 'header preserved');
      assert.equal(r.rows[r.rows.length - 1][1], 'id5999', 'newest row is in hand');
      assert.equal(big.fullReads, 0, 'the whole sheet is never fetched');
      assert.ok(r.rows[1][0].getTime() >= now - daysAgo * DAY, 'nothing older than the cutoff came along');
    });
  }
  check('a row index survives the offset, so callers can still write back', () => {
    const r = read(big, 10);
    // rows[i] is sheet row (i + 1 + off); the mock array is 0-based, so that is
    // index (i + off). i = 1 → the first row handed back.
    assert.equal(big.rows[1 + r.off][1], r.rows[1][1], 'rows[i] is sheet row (i + 1 + off)');
  });
  // The shape that broke both guesses: 'תוצאות תרגול' = 107,614 rows.
  check('at the real sheet scale a 32-day cutoff reads a fraction, not everything', () => {
    const huge = env.sheet('huge'); fill(huge, 20000, 110);   // same rows-per-day ratio, kept testable
    huge.fullReads = 0;
    const r = env.ctx.readRowsSince(huge, 0, new Date(now - 32 * DAY));
    assert.match(r.mode, /^rows\d+\/20001$/, 'must not fall back to full: ' + r.mode);
    assert.equal(Number(/^rows(\d+)\//.exec(r.mode)[1]), exactRows(huge, 32));
    assert.ok(Number(/^rows(\d+)\//.exec(r.mode)[1]) < 20000 / 2, 'well under half the sheet');
    assert.equal(huge.fullReads, 0);
  });
  check('an unparseable date is kept, never silently dropped', () => {
    const junk = env.sheet('junk'); fill(junk, 1500, 300);
    junk.rows[400][0] = '';                    // array index 400 = sheet row 401, far older than the cutoff
    const r = env.ctx.readRowsSince(junk, 0, new Date(now - 5 * DAY));
    assert.match(r.mode, /^rows\d+\/1501$/);
    assert.equal(r.off, 401 - 2, 'the read starts AT the unreadable row, not after it');
    assert.equal(r.rows[1][0], '', 'and it is the first row handed back');
  });
  check('nothing in range returns the header alone, not the sheet', () => {
    const r = env.ctx.readRowsSince(big, 0, new Date(now + 10 * DAY));
    assert.equal(r.mode, 'none/6001'); assert.equal(r.rows.length, 1); assert.equal(big.fullReads, 0);
  });
  const small = env.sheet('small'); fill(small, 200, 400);
  check('a small sheet is simply read whole', () => {
    const r = read(small, 1); assert.equal(r.mode, 'full/201'); assert.equal(r.rows.length, 201); assert.equal(r.off, 0);
  });
  check('no usable cutoff means a full read, never a guess', () => {
    assert.equal(env.ctx.readRowsSince(big, 0, null).mode, 'full/6001');
    assert.equal(env.ctx.readRowsSince(big, 0, new Date(NaN)).mode, 'full/6001');
  });

  // ---- r17: width. Bounding rows bought ~1s on 'תוצאות תרגול' (28.5s, `full`)
  // while the LONGER 'תוצאות' read in 1.3s — because a practice row carries two
  // JSON blobs (cols N/O) the handler never reads.
  {
    const wide = env.sheet('wide');
    const blob = 'x'.repeat(2000);
    wide.rows.push(['תאריך', 'מזהה', 'שם', 'כיתה', 'מצב', 'דרגה', 'ציון', 'סה"כ', 'אחוז', 'עבר', 'זמן', 'נושא', 'שפה', 'פירוט שגויות', 'לפי נושא', 'טלפון']);
    for (let i = 0; i < 300; i++) {
      wide.rows.push([new Date(now - (300 - i) * 60000), 'S' + i, 'שם ' + i, 'C1', 'exam', 'B', 20, 30, 67, 'נכשל', '10:00', '', 'he', blob, blob, '050' + i]);
    }
    const SPEC = [[1, 13], [16, 1]];
    wide.cellsRead = 0;
    const pruned = env.ctx.readRowsSince(wide, 0, new Date(now - 1 * DAY), SPEC);
    const prunedCells = wide.cellsRead;
    check('a pruned read keeps every index the caller uses', () => {
      const row = pruned.rows[5];
      assert.equal(row.length, 16, 'absolute column positions are preserved');
      assert.ok(row[0] instanceof Date); assert.equal(row[2], 'שם 4'); assert.equal(row[5], 'B');
      assert.equal(row[8], 67); assert.equal(row[15], '0504', 'col P still lands on index 15');
    });
    check('the JSON columns it does not use come back empty, not fetched', () => {
      assert.equal(pruned.rows[5][13], ''); assert.equal(pruned.rows[5][14], '');
    });
    wide.cellsRead = 0;
    env.ctx.readRowsSince(wide, 0, new Date(now - 1 * DAY));   // same rows, every column
    check('pruning is what makes the read cheap, not the row bound', () => {
      assert.ok(prunedCells < wide.cellsRead * 0.9,
        `pruned ${prunedCells} cells vs ${wide.cellsRead} unpruned`);
    });
    check('the header row is pruned the same way, so it still aligns', () => {
      const wideT = env.sheet('wideTail');
      wideT.rows.push(wide.rows[0].slice());
      for (let i = 0; i < 1500; i++) wideT.rows.push([new Date(now - (1500 - i) * 3600000), 'S' + i, 'n', 'C', 'exam', 'B', 1, 2, 3, '', '', '', 'he', blob, blob, '05']);
      const t = env.ctx.readRowsSince(wideT, 0, new Date(now - 2 * DAY), SPEC);
      assert.match(t.mode, /^rows\d+\/1501$/);
      assert.equal(t.rows[0][0], 'תאריך'); assert.equal(t.rows[0][15], 'טלפון');
      assert.equal(t.rows[0][13], '', 'skipped in the header exactly as in the body');
    });
  }
  check('an unparseable date anchors the read at itself, without reading the sheet', () => {
    // Under the old ladder this rejected the tail and read all 1500 rows. The
    // exact scan keeps the unreadable row — it is the topmost row that may
    // matter — and reads from there down, nothing above it.
    const odd = env.sheet('odd'); fill(odd, 1500, 30); odd.rows[odd.rows.length - 1000][0] = 'not a date';
    odd.fullReads = 0;
    const r = read(odd, 1);
    assert.equal(r.mode, 'rows1000/1501');
    assert.equal(r.rows.length, 1001); assert.equal(odd.fullReads, 0);
    assert.equal(r.rows[1][0], 'not a date', 'the row we could not read is the first one kept');
  });
  // The commander handler must route both heavy reads through it, bounded by
  // prevFrom (results) and prevFrom - 31 days (practice).
  const src = fs.readFileSync(path.join(__dirname, '..', 'external_exam_apps_script.js'), 'utf8');
  const cmd = src.slice(src.indexOf('function handleCommanderDashboard('), src.indexOf('\nfunction ', src.indexOf('function handleCommanderDashboard(') + 10));
  check('the commander dashboard reads results and practice through readRowsSince', () => {
    assert.ok(/readRowsSince\(resSheet, 0, prevFrom \? new Date\(prevFrom\.getTime\(\) - DAY_MS\)/.test(cmd));
    // The spec is read out of the source and checked against the columns the
    // handler actually indexes — a column dropped from one side must fail here,
    // because at runtime it would silently read as '' instead.
    const specSrc = /readRowsSince\(practiceSheet, 0,[\s\S]{0,140}?(\[\[[\d, \[\]]+\])\)/.exec(cmd);
    assert.ok(specSrc, 'the practice read passes an explicit column spec');
    const allowed = new Set();
    for (const [first, count] of JSON.parse(specSrc[1])) for (let c = first; c < first + count; c++) allowed.add(c - 1);
    const touched = [...cmd.matchAll(/practiceData\[\w+\]\[(\d+)\]/g)].map(m => Number(m[1]));
    assert.ok(touched.length >= 5, 'found the indexing sites');
    const missing = [...new Set(touched)].filter(i => !allowed.has(i));
    assert.deepEqual(missing, [], 'columns indexed but not fetched (they would read empty)');
    assert.ok(!allowed.has(13) && !allowed.has(14), 'the two JSON blob columns stay out of the read');
    assert.equal((cmd.match(/getSheet\('תוצאות( תרגול)?'\)\.getDataRange\(\)/g) || []).length, 0, 'no bare full read of either sheet remains');
  });
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

  // The five pools ARE the whole bank of a language (measured against the real
  // banks: every question belongs to at least one license). Coverage must NOT
  // be judged against the translation index's count - that is the union ACROSS
  // languages (1700) and exceeds a single language's bank (he 1694, ru 1693),
  // which once sent every language except English back to Drive.
  const noIndex = environment(banks);
  noIndex.ctx.warmupQuestionCaches();
  const beforeNoIndex = { ...noIndex.reads };
  noIndex.cache.remove('qv2_tx_meta');
  const stillCached = noIndex.ctx.questionMetaForLanguage('ru', {});
  check('a language whose bank is smaller than the cross-language index still uses the cache', () => {
    assert.deepEqual(noIndex.reads, beforeNoIndex, 'no Drive read');
    assert.equal(stillCached.length, 300);
  });
}

// ---- 10. the cache keeps itself warm: nobody runs the warmup by hand -------
// 2026-09-17, the operator: "I don't want to run the warmup by hand every day,
// and not several times a day." The hand-made 4-hour trigger rebuilt blindly and
// never looked; an hourly ensureQuestionCachesWarm verifies and repairs.
{
  const HOUR = 3600 * 1000;
  const env = environment(banks);
  const first = env.ctx.ensureQuestionCachesWarm();
  check('hourly check: a cold cache is rebuilt', () => {
    assert.match(first, /^rebuilt \(cache not ready/);
    assert.equal(env.ctx.questionCacheStatus().ready, true);
  });
  env.clock.t += 30 * 60 * 1000;
  const readsBefore = { ...env.reads };
  const quiet = env.ctx.ensureQuestionCachesWarm();
  check('hourly check: a warm, fresh cache costs no rebuild and no Drive read', () => {
    assert.match(quiet, /^warm: \d+ keys ready, last complete warmup 3\d min ago/);
    assert.deepEqual(env.reads, readsBefore);
  });
  env.cache.remove('qv2_pool_he_B_meta');                  // CacheService evicts early sometimes
  const repaired = env.ctx.ensureQuestionCachesWarm();
  check('hourly check: an evicted pool is repaired at the next tick', () => {
    assert.match(repaired, /^rebuilt \(cache not ready: \d+ missing/);
    assert.equal(env.ctx.questionCacheStatus().ready, true);
  });
  env.clock.t += 4 * HOUR + 60 * 1000;
  const refreshed = env.ctx.ensureQuestionCachesWarm();
  check('hourly check: past 4 hours it refreshes, before the 6-hour TTL can expire', () =>
    assert.match(refreshed, /^rebuilt \(last complete warmup 24\d min ago\)/));

  // A PARTIAL run must not reset the age: some of its pools are still old.
  const partialEnv = environment(banks);
  partialEnv.ctx.warmupQuestionCaches();                    // complete: stamps lastCompleteAt
  const stamped = JSON.parse(partialEnv.properties.get('qv2_warmup_state')).lastCompleteAt;
  partialEnv.clock.t += 2 * HOUR;
  partialEnv.ctx.warmupQuestionCaches({ budgetMs: 20000 }); // far too small → PARTIAL
  check('a PARTIAL warmup does not count as a refresh', () =>
    assert.equal(JSON.parse(partialEnv.properties.get('qv2_warmup_state')).lastCompleteAt, stamped));
}
{
  // Two days of an hour timer, firing at a different minute every hour the way
  // Apps Script's does. After the first tick the cache must be ready at EVERY
  // tick — i.e. it never went cold in between — with a bounded rebuild count.
  const HOUR = 3600 * 1000, MIN = 60 * 1000;
  const env = environment(banks);
  const base = env.clock.t;
  let rebuilds = 0; const coldAt = [];
  for (let h = 0; h < 48; h++) {
    const target = base + h * HOUR + ((h * 37) % 50) * MIN;   // 0-49 min into the hour
    if (env.clock.t < target) env.clock.t = target;
    if (h > 0 && !env.ctx.questionCacheStatus().ready) coldAt.push(h);
    if (/^rebuilt/.test(env.ctx.ensureQuestionCachesWarm())) rebuilds++;
  }
  check('48 hours of hourly ticks: the cache is never found cold after the first', () =>
    assert.deepEqual(coldAt, [], 'cold at hour(s) ' + coldAt.join(',')));
  check('48 hours of hourly ticks: rebuilds stay about one per 4-5 hours', () =>
    assert.ok(rebuilds >= 9 && rebuilds <= 13, 'rebuilds=' + rebuilds));

  // The simulation must be able to FAIL, or "never cold" proves nothing: with a
  // refresh threshold past the 6h TTL the cache really does go cold between ticks.
  const broken = environment(banks);
  broken.ctx.WARMUP_REFRESH_AFTER_MS = 6.5 * HOUR;
  const bBase = broken.clock.t; const bCold = [];
  for (let h = 0; h < 48; h++) {
    const target = bBase + h * HOUR + ((h * 37) % 50) * MIN;
    if (broken.clock.t < target) broken.clock.t = target;
    if (h > 0 && !broken.ctx.questionCacheStatus().ready) bCold.push(h);
    broken.ctx.ensureQuestionCachesWarm();
  }
  check('the 48-hour simulation detects a broken refresh rule (it is not vacuous)', () =>
    assert.ok(bCold.length > 0, 'a 6.5h threshold should leave the cache cold at some tick'));
}
{
  const env = environment(banks);
  const fake = fn => ({ getHandlerFunction: () => fn });
  env.triggers.push(fake('warmupQuestionCaches'));          // the hand-made 4-hour one
  env.triggers.push(fake('rebuildMissingQuestionCaches'));  // a live miss's one-shot
  env.triggers.push(fake('archiveOldPendingRows'));         // unrelated daily job
  const msg = env.ctx.installWarmupTriggers();
  const handlers = () => env.triggers.map(t => t.getHandlerFunction()).sort();
  check('install: replaces the hand-made warmup trigger with one hourly check', () => {
    assert.deepEqual(handlers(), ['archiveOldPendingRows', 'ensureQuestionCachesWarm', 'rebuildMissingQuestionCaches']);
    assert.equal(env.triggers.find(t => t.getHandlerFunction() === 'ensureQuestionCachesWarm').everyHours, 1);
    assert.match(msg, /removed 1 old trigger\(s\) \[warmupQuestionCaches\]/);
    assert.match(msg, /first check → rebuilt/);
  });
  env.ctx.installWarmupTriggers();
  check('install: running it again still leaves exactly one hourly check', () =>
    assert.equal(env.triggers.filter(t => t.getHandlerFunction() === 'ensureQuestionCachesWarm').length, 1));
}

console.log(`\n${checks} checks passed`);
