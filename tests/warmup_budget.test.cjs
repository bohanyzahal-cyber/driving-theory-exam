// Run: node tests/warmup_budget.test.cjs
//
// Proves the properties that the 09/09 and 11/09/2026 production kills violated:
// a warmup run must end on its own terms inside its budget, must never leave a
// cache lease behind, must resume from a stored cursor, and must never publish
// a thinner translation index than the one already serving traffic.
//
// Google's execution ceiling is 360s. The clock here is simulated, so slow
// Drive reads and slow compression are modelled without waiting for them.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const zlib = require('node:zlib');
const { randomUUID } = require('node:crypto');

const source = fs.readFileSync(path.join(__dirname, '..', 'external_exam_apps_script.js'), 'utf8');
const LANGS = ['he', 'ru', 'en', 'ar', 'fr', 'es', 'am'];
const LICENSES = ['B', '1', 'C1', 'C', 'D'];
const KILL_MS = 360000;
const ORIGIN = 1757000000000;

function syntheticBanks() {
  const banks = {};
  for (const lang of LANGS) {
    banks[lang] = [];
    for (let id = 1; id <= 1700; id++) {
      // Production format: one row per (question, license), IDs repeat.
      for (const license of LICENSES) banks[lang].push({
        id, text: `${lang} שאלה ${id} ${license}`,
        answers: ['אחת', 'שתיים', 'three', 'four'],
        category: 'חוק 1', licenseType: license === '1' ? '' : license,
        imageUrl: null, language: lang
      });
    }
  }
  return banks;
}

// costs is read on every call, so a scenario can make Google turn slow midway.
function environment(banks, costs) {
  const clock = { t: ORIGIN };
  const entries = new Map();
  const properties = new Map([['QUESTIONS_DRIVE_FOLDER_ID', 'fixture-folder']]);
  const logs = [], reads = {};
  let held = false;

  const RealDate = Date;
  function FakeDate(...args) { return args.length ? new RealDate(...args) : new RealDate(clock.t); }
  FakeDate.now = () => clock.t;
  FakeDate.parse = RealDate.parse;
  FakeDate.UTC = RealDate.UTC;
  FakeDate.prototype = RealDate.prototype;

  const cache = {
    get(key) { const e = entries.get(key); return e && e.expires > clock.t ? e.value : null; },
    getAll(keys) { return Object.fromEntries(keys.map(k => [k, this.get(k)]).filter(([, v]) => v !== null)); },
    put(key, value, ttl = 600) {
      assert.equal(typeof value, 'string');
      assert.ok(Buffer.byteLength(value, 'utf8') <= 100 * 1024, 'cache value within 100KB');
      entries.set(key, { value, expires: clock.t + ttl * 1000 });
    },
    putAll(values, ttl) { for (const [k, v] of Object.entries(values)) this.put(k, v, ttl); },
    remove(key) { entries.delete(key); },
    removeAll(keys) { for (const key of keys) entries.delete(key); }
  };
  const blob = data => {
    const bytes = typeof data === 'string' ? Buffer.from(data, 'utf8') : Buffer.from(data);
    return { getBytes: () => [...bytes], getDataAsString: () => bytes.toString('utf8') };
  };
  const ctx = {
    Date: FakeDate,
    Logger: { log: s => logs.push(s) },
    CacheService: { getScriptCache: () => cache },
    PropertiesService: { getScriptProperties: () => ({
      getProperties: () => Object.fromEntries(properties),
      getProperty: k => (properties.has(k) ? properties.get(k) : null),
      setProperty: (k, v) => { properties.set(k, v); },
      deleteProperty: k => { properties.delete(k); }
    }) },
    LockService: { getScriptLock: () => {
      let mine = false;
      return {
        tryLock(ms) { assert.ok(ms <= 200, 'mutex waits must stay short'); if (held) return false; held = mine = true; return true; },
        releaseLock() { if (mine) { held = false; mine = false; } }
      };
    } },
    Utilities: {
      getUuid: randomUUID,
      newBlob: blob,
      gzip: b => { const bytes = Buffer.from(b.getBytes()); clock.t += costs.gzipMs(bytes.length); return blob(zlib.gzipSync(bytes)); },
      ungzip: b => blob(zlib.gunzipSync(Buffer.from(b.getBytes()))),
      base64Encode: b => Buffer.from(b).toString('base64'),
      base64Decode: s => [...Buffer.from(s, 'base64')],
      sleep: ms => { clock.t += ms; }
    },
    DriveApp: { getFolderById: id => {
      assert.equal(id, 'fixture-folder');
      return { getFilesByName: file => {
        const lang = /^questions_([a-z]+)\.json$/.exec(file)[1];
        return { hasNext: () => !!banks[lang], next: () => ({ getBlob: () => {
          reads[lang] = (reads[lang] || 0) + 1;
          clock.t += costs.driveMs;
          return blob(JSON.stringify(banks[lang]));
        } }) };
      } };
    } }
  };
  vm.createContext(ctx);
  vm.runInContext(source, ctx);
  ctx.lookupCorrectIndex = (id, lang) => (id + LANGS.indexOf(lang)) % 4;
  return { ctx, cache, entries, properties, logs, reads, clock,
    run(options) {
      const at = clock.t;
      const report = ctx.warmupQuestionCaches(options);
      return { report, elapsed: clock.t - at, text: report.join('\n') };
    },
    leases() { return [...properties.keys()].filter(k => k.startsWith('qv2_lease_')); },
    pools() { return [...entries.keys()].filter(k => /^qv2_pool_.+_meta$/.test(k)).length; },
    txLangs() { const m = cache.get('qv2_tx_meta'); return m ? JSON.parse(m).langs.length : 0; },
    txGeneration() { const m = cache.get('qv2_tx_meta'); return m ? JSON.parse(m).g : null; }
  };
}

const HEALTHY = () => ({ driveMs: 6000, gzipMs: bytes => Math.max(40, Math.round(bytes / 300)) });
const SLOW = () => ({ driveMs: 26000, gzipMs: bytes => Math.max(200, Math.round(bytes / 90)) });

const banks = syntheticBanks();
let checks = 0;
const check = (label, fn) => { fn(); checks++; console.log('ok  ' + label); };

// ---- 1. A healthy run completes everything, inside the budget -------------
const env = environment(banks, HEALTHY());
const first = env.run();
check('healthy run reports COMPLETE with no errors', () => {
  assert.match(first.text, /warmup COMPLETE: 7\/7 pool languages/);
  assert.ok(!/ERROR/.test(first.text), first.text);
  assert.ok(!/cached=false/.test(first.text), first.text);
});
check('healthy run stays inside the budget, far below the 360s kill', () => {
  assert.ok(first.elapsed < KILL_MS, `elapsed ${first.elapsed}ms`);
  assert.ok(first.elapsed <= 300000, `elapsed ${first.elapsed}ms must respect the budget`);
});
check('healthy run publishes all 35 pools and all 7 index languages', () => {
  assert.equal(env.pools(), 35);
  assert.equal(env.txLangs(), 7);
  assert.equal(env.ctx.questionCacheStatus().ready, true);
});
check('no lease survives the run', () => assert.deepEqual(env.leases(), []));
check('each bank is read from Drive exactly once per run', () =>
  assert.deepEqual(env.reads, Object.fromEntries(LANGS.map(l => [l, 1]))));

// ---- 2. Every run refreshes the index and the pools ----------------------
// The index build is ~6s on a run that must read all seven banks for the pools
// anyway, so it is never skipped for being "fresh enough": a threshold sitting
// on the trigger interval left an index unrenewed until it expired.
const generation = env.txGeneration();
env.clock.t += 60 * 60 * 1000;
const second = env.run();
check('the next run rebuilds the index rather than trusting its age', () => {
  assert.match(second.text, /translation-index: 1700 questions/);
  assert.notEqual(env.txGeneration(), generation, 'a new generation is published');
  assert.equal(env.txLangs(), 7);
});
check('the next run also refreshes every pool and renews their TTL', () => {
  assert.equal(env.pools(), 35);
  assert.equal(env.ctx.questionCacheStatus().ready, true);
});
check('an index older than its TTL can never be served', () => {
  // The failure this replaces: a 6h TTL with a 4h refresh threshold meant a
  // run at exactly 4h skipped the rebuild, and the index lapsed at 6h.
  const age = env.clock.t - (JSON.parse(env.cache.get('qv2_tx_meta')).builtAt);
  assert.ok(age < 60 * 60 * 1000, `index was rebuilt this run (age ${age}ms)`);
});
check('the r1-r3 key sweep is gone from every run', () => {
  // It cost ~72 CacheService round-trips per run and cannot match anything any
  // more; nothing has written those key names since r4.
  assert.ok(!/legacy cleanup/.test(first.text), first.text);
  assert.ok(!/legacy cleanup/.test(second.text), second.text);
});

// ---- 3. A run after the TTL has lapsed republishes from scratch -----------
env.clock.t += 7 * 60 * 60 * 1000;   // past the six-hour cache TTL
const third = env.run();
check('a run whose index has expired republishes it with the larger budget', () => {
  assert.match(third.text, /translation-index: MISSING - this run takes the larger 300000ms budget/);
  assert.match(third.text, /translation-index: 1700 questions/);
  assert.equal(env.ctx.questionCacheStatus().ready, true);
});

// ---- 4. A sick Google must yield PARTIAL, never a kill -------------------
const sick = environment(banks, SLOW());
const runs = [];
for (let i = 0; i < 6; i++) { sick.clock.t += 60 * 60 * 1000; runs.push(sick.run()); }
check('every run under slow Drive and slow gzip stays inside the ceiling', () => {
  for (const r of runs) assert.ok(r.elapsed < KILL_MS, `a run took ${r.elapsed}ms; the kill is ${KILL_MS}ms`);
});
check('slow runs leave no lease behind', () => assert.deepEqual(sick.leases(), []));
check('consecutive runs resume the cursor until every pool exists', () => {
  assert.ok(runs.some(r => /warmup PARTIAL/.test(r.text)), 'a slow run is expected to stop early');
  assert.equal(sick.pools(), 35, 'the cursor eventually covers 7 languages x 5 licenses');
  assert.equal(sick.ctx.questionCacheStatus().ready, true);
});
check('a partial run names the language the next run continues from', () => {
  const partial = runs.find(r => /warmup PARTIAL/.test(r.text));
  assert.match(partial.text, /next cursor=(he|ru|en|ar|fr|es|am) \(the next scheduled run continues from there\)/);
});

// ---- 5. Drive turning pathological must not thin a published index -------
const costs = HEALTHY();
const flip = environment(banks, costs);
flip.run();
const published = flip.txGeneration();
assert.equal(flip.txLangs(), 7, 'precondition: a complete index is in service');
flip.clock.t += 5 * 60 * 60 * 1000;      // the index is now due for a rebuild
costs.driveMs = 95000;                   // one bank read now eats a third of the ceiling
const starved = flip.run();
check('a run that cannot load every bank leaves the published index untouched', () => {
  assert.match(starved.text, /bank loads: PARTIAL/);
  assert.equal(flip.txLangs(), 7, 'never a thinned index');
  assert.equal(flip.txGeneration(), published, 'the serving index is not replaced');
});
check('the starved run still ends inside the ceiling and releases its leases', () => {
  assert.ok(starved.elapsed < KILL_MS, `starved run consumed ${starved.elapsed}ms`);
  assert.deepEqual(flip.leases(), []);
});

// ---- 6. Emergency reset forces a full rebuild from the first language ----
const emergency = environment(banks, HEALTHY());
emergency.run();
emergency.clock.t += 10 * 60 * 1000;
const before = emergency.txGeneration();
const reset = emergency.ctx.emergencyClearAndRefreshCache();
check('emergency reset rebuilds the index even when it is fresh', () => {
  assert.match(reset, /translation-index: 1700 questions/);
  assert.notEqual(emergency.txGeneration(), before);
});
check('emergency reset leaves no lease behind', () => assert.deepEqual(emergency.leases(), []));

// ---- 7. The index-only editor helper ------------------------------------
const indexOnly = environment(banks, HEALTHY());
const measured = indexOnly.ctx.warmupTranslationIndexOnly().join('\n');
check('index-only helper publishes the index and reports both phases', () => {
  assert.match(measured, /translation-index: 1700 questions; languages=he,ru,en,ar,fr,es,am; cached=true/);
  assert.match(measured, /phase timings: banks \d+ms, index build \d+ms, total \d+ms/);
  assert.equal(indexOnly.txLangs(), 7);
  assert.equal(indexOnly.pools(), 0, 'it must not touch the pools');
  assert.deepEqual(indexOnly.leases(), []);
});
// A missing bank must abort the helper instead of publishing a thinner index.
const missingRu = JSON.parse(JSON.stringify(banks));
delete missingRu.ru;
const thin = environment(missingRu, HEALTHY());
const thinReport = thin.ctx.warmupTranslationIndexOnly().join('\n');
check('index-only helper aborts rather than publish a thinner index', () => {
  assert.match(thinReport, /translation-index: ABORTED - only 6\/7 banks loaded/);
  assert.equal(thin.txLangs(), 0, 'nothing published at all');
  assert.deepEqual(thin.leases(), []);
});

console.log(`\n${checks} checks passed`);
