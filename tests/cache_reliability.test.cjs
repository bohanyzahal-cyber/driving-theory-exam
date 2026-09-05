// Run: node tests/cache_reliability.test.cjs [optional-private-bank-directory]
// Synthetic fixtures only are committed. Optional real data produces counts only.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const zlib = require('node:zlib');
const { randomUUID, randomBytes } = require('node:crypto');
const source = fs.readFileSync(path.join(__dirname, '..', 'external_exam_apps_script.js'), 'utf8');
const langs = ['he', 'ru', 'en', 'ar', 'fr', 'es', 'am'];
const licenses = ['B', '1', 'C1', 'C', 'D'];
const json = value => JSON.parse(JSON.stringify(value));

function syntheticBanks() {
  const banks = {};
  for (const lang of langs) {
    banks[lang] = [];
    for (let id = 1; id <= 1750; id++) {
      // Repeated IDs have distinct license rows, as in the production format.
      for (const license of licenses) banks[lang].push({
        id, text: `${lang} שאלה عربية русский አማርኛ ${id} ${license}`,
        answers: ['אחת', 'שתיים', 'three', 'four'],
        category: 'חוק 1', licenseType: license === '1' ? '' : license,
        imageUrl: null, language: lang
      });
    }
    // A question absent in one language must stay absent, not be invented.
    if (lang === 'am') banks[lang] = banks[lang].filter(q => q.id !== 1729);
  }
  return banks;
}

function environment(banks) {
  const entries = new Map(), properties = new Map([['QUESTIONS_DRIVE_FOLDER_ID', 'fixture-folder']]);
  const logs = [], reads = {}, stats = { evictions: 0, maxKeys: 0, maxBytes: 0, sleeps: 0, lockedWork: 0 };
  let now = 1000000000, held = false, beforeDrive = null, rejectWrites = false;
  const cache = {
    get(key) { const e = entries.get(key); return e && e.expires > now ? e.value : null; },
    getAll(keys) { return Object.fromEntries(keys.map(k => [k, this.get(k)]).filter(([,v]) => v !== null)); },
    put(key, value, ttl = 600) {
      if (rejectWrites) throw new Error('injected cache failure');
      assert.equal(typeof value, 'string');
      assert.ok(key.length <= 250, 'cache key size');
      const bytes = Buffer.byteLength(value, 'utf8');
      assert.ok(bytes <= 100 * 1024, `oversized cache value: ${bytes}`);
      stats.maxBytes = Math.max(stats.maxBytes, bytes);
      entries.set(key, { value, expires: now + ttl * 1000 });
      if (entries.size > 1000) {
        const retained = [...entries].sort((a,b) => b[1].expires - a[1].expires).slice(0,900);
        stats.evictions += entries.size - retained.length;
        entries.clear(); for (const [k,v] of retained) entries.set(k,v);
      }
      stats.maxKeys = Math.max(stats.maxKeys, entries.size);
    },
    putAll(values, ttl) { for (const [k,v] of Object.entries(values)) this.put(k,v,ttl); },
    remove(key) { entries.delete(key); },
    removeAll(keys) { for (const key of keys) entries.delete(key); }
  };
  function blob(data) {
    const bytes = typeof data === 'string' ? Buffer.from(data,'utf8') : Buffer.from(data);
    return { getBytes: () => [...bytes], getDataAsString: () => bytes.toString('utf8') };
  }
  function unlocked() { if (held) stats.lockedWork++; assert.equal(held, false, 'global mutex must not cover expensive work'); }
  const ctx = {
    Logger: { log: s => logs.push(s) },
    CacheService: { getScriptCache: () => cache },
    PropertiesService: { getScriptProperties: () => ({
      getProperty: k => properties.get(k) || null,
      setProperty: (k,v) => { properties.set(k,v); },
      deleteProperty: k => properties.delete(k)
    }) },
    LockService: { getScriptLock: () => {
      let mine = false;
      return { tryLock(ms) { assert.ok(ms <= 200); if (held) return false; held = mine = true; return true; }, releaseLock() { if (mine) held = mine = false; } };
    } },
    Utilities: {
      getUuid: randomUUID, newBlob: blob,
      gzip: b => { unlocked(); return blob(zlib.gzipSync(Buffer.from(b.getBytes()))); },
      ungzip: b => { unlocked(); return blob(zlib.gunzipSync(Buffer.from(b.getBytes()))); },
      base64Encode: b => Buffer.from(b).toString('base64'),
      base64Decode: s => [...Buffer.from(s,'base64')],
      sleep() { stats.sleeps++; throw new Error('sleep is forbidden on cache contention'); }
    },
    DriveApp: { getFolderById: id => {
      unlocked(); assert.equal(id,'fixture-folder');
      return { getFilesByName: file => {
        const lang = /^questions_([a-z]+)\.json$/.exec(file)[1];
        return { hasNext: () => !!banks[lang], next: () => ({ getBlob: () => {
          unlocked(); reads[lang] = (reads[lang] || 0) + 1;
          if (beforeDrive) beforeDrive(lang);
          return blob(JSON.stringify(banks[lang]));
        } }) };
      } };
    } }
  };
  vm.createContext(ctx); vm.runInContext(source, ctx);
  ctx.lookupCorrectIndex = (id, lang) => (id + langs.indexOf(lang)) % 4;
  return { ctx, cache, entries, properties, stats, reads, logs,
    beforeDrive(fn) { beforeDrive = fn; }, rejectWrites(on) { rejectWrites = on; },
    advance(ms) { now += ms; }, setHeld(on) { held = on; }
  };
}

function expectedTranslations(banks, ids, includeCi) {
  const out = {};
  for (const [lang, rows] of Object.entries(banks)) {
    out[lang] = {};
    for (const q of rows) if (ids.includes(q.id)) {
      out[lang][q.id] = { t:q.text, a:q.answers };
      if (includeCi) out[lang][q.id].ci = ((q.id + langs.indexOf(lang)) % 4) ^ (q.id % 256);
    }
  }
  return out;
}

const banks = syntheticBanks();
const env = environment(banks);
// Seed the old oversized index (including no surviving tx_meta) to verify
// first-deploy cleanup does not depend on old manifests still being cached.
for (let id=1;id<=1000;id++) env.cache.put('tx_'+id, '{}', 21600);
assert.equal(env.entries.size,1000,'migration begins at the cache capacity limit');
const report = env.ctx.warmupQuestionCaches();
assert.ok(report.every(line => !line.includes('ERROR') && !line.includes('cached=false')), report.join('\n'));
assert.deepEqual(env.reads, Object.fromEntries(langs.map(l => [l,1])), 'warmup reads each Drive bank once');
assert.ok(![...env.entries.keys()].some(k => /^tx_/.test(k)), 'legacy translation keys removed even without metadata');
assert.equal(env.ctx.questionCacheStatus().ready,true);
assert.ok(env.ctx.questionCacheStatus().presentKeys <=423);
assert.ok(env.stats.maxBytes <81000);
const evictionsAfterWarmup = env.stats.evictions;
// Reserve realistic ephemeral traffic alongside all long-lived cache records.
for (let i=0;i<400;i++) env.cache.put('active-exam-fixture-'+i,'opaque',3600);
assert.equal(env.stats.evictions,evictionsAfterWarmup,'persistent cache leaves room for 400 active/ephemeral keys');

for (const includeCi of [false,true]) {
  for (let start=1;start<=1750;start+=30) {
    const ids=Array.from({length:Math.min(30,1751-start)},(_,i)=>start+i);
    assert.deepEqual(json(env.ctx.tryTranslationsFromIndex(ids,includeCi)),expectedTranslations(banks,ids,includeCi));
  }
}
for (const lang of langs) for (const license of licenses) {
  const got=json(env.ctx.loadLicensePoolServer(lang,license));
  const seen=new Set();
  const expected=banks[lang].filter(q => license==='1' ? !q.licenseType && q.category.includes('1') : q.licenseType===license)
    .filter(q=>!seen.has(q.id)&&(seen.add(q.id),true));
  assert.deepEqual(got,expected,'license filter must precede deduplication');
}
assert.deepEqual(env.reads,Object.fromEntries(langs.map(l=>[l,1])),'warm reads never touch Drive');
// A four-hour warmup must renew cached banks as well as derived pools/index.
const initialExpiry=env.entries.get('qv2_bank_he_meta').expires;
env.advance(4*60*60*1000);
env.ctx.warmupQuestionCaches();
assert.ok(env.entries.get('qv2_bank_he_meta').expires>initialExpiry,'cache hits receive a renewed TTL');
env.advance(3*60*60*1000);
assert.equal(env.ctx.loadQuestionsForLanguageServer('he').length,banks.he.length);
assert.deepEqual(env.reads,Object.fromEntries(langs.map(l=>[l,1])),'renewed bank survives beyond original six-hour expiry');

// Missing/corrupt/mixed shards must never produce incomplete translations.
const ids=[1,2,3,1729], shardKey='qv2_tx_'+env.ctx.questionTranslationShard(1);
const oldShard=env.cache.get(shardKey);
env.cache.remove(shardKey);
assert.equal(env.ctx.tryTranslationsFromIndex(ids,false),null);
assert.deepEqual(json(env.ctx.buildExamTranslations(ids.map(id=>({id})),false)),expectedTranslations(banks,ids,false));
env.cache.put(shardKey, 'different-generation:'+oldShard.split(':').slice(1).join(':'),21600);
assert.equal(env.ctx.tryTranslationsFromIndex(ids,false),null);
env.cache.put(shardKey, oldShard,21600);

// Byte encoding protects multibyte JSON; truncated gzip and mixed generation
// chunks are detected before returning a value.
const unicode={text:'שלום مرحبا Привет አማርኛ 😀'.repeat(40000)};
assert.equal(env.ctx.writeQuestionCacheRecord(env.cache,'unicode-fixture',unicode,16),true);
assert.deepEqual(json(env.ctx.readQuestionCacheRecord(env.cache,'unicode-fixture',16)),unicode);
const unicodeManifest=JSON.parse(env.cache.get('unicode-fixture_meta'));
env.cache.put('unicode-fixture_0',unicodeManifest.g+':'+Buffer.from('truncated gzip').toString('base64'),21600);
assert.equal(env.ctx.readQuestionCacheRecord(env.cache,'unicode-fixture',16),null);
assert.equal(env.ctx.writeQuestionCacheRecord(env.cache,'over-budget-fixture',{noise:randomBytes(200000).toString('base64')},1),false);
assert.equal(env.cache.get('over-budget-fixture_meta'),null,'oversized record never publishes a manifest');

// True lease contention returns immediately: no extra Drive read and no sleeps.
const cold=environment(banks);
cold.beforeDrive(lang=>assert.throws(()=>cold.ctx.loadQuestionsForLanguageServer(lang),e=>e.code==='question_cache_busy'&&e.retryable&&e.waitSec===3));
cold.ctx.loadQuestionsForLanguageServer('he');
assert.equal(cold.reads.he,1);
assert.equal(cold.stats.sleeps,0);
assert.equal(cold.stats.lockedWork,0);
assert.equal([...cold.properties.keys()].filter(k=>k.includes('lease_')).length,0);
cold.setHeld(true);
assert.throws(()=>cold.ctx.claimQuestionCacheLease('test'),e=>e.code==='question_cache_busy');
cold.setHeld(false);

// Builder failure releases its lease and the next request can recover.
const failure=environment(banks);
failure.beforeDrive(()=>{throw new Error('injected Drive failure');});
assert.throws(()=>failure.ctx.loadQuestionsForLanguageServer('he'),/injected Drive failure/);
assert.equal([...failure.properties.keys()].filter(k=>k.includes('lease_')).length,0);
failure.beforeDrive(null);
assert.equal(failure.ctx.loadQuestionsForLanguageServer('he').length,banks.he.length);
const owned=failure.ctx.claimQuestionCacheLease('owner-test');
failure.properties.set(owned.key,JSON.stringify({owner:'new-owner',until:Date.now()+100000}));
failure.ctx.releaseQuestionCacheLease(owned);
assert.equal(JSON.parse(failure.properties.get(owned.key)).owner,'new-owner','old owner cannot remove replacement lease');
failure.properties.set('qv2_lease_expired',JSON.stringify({owner:'expired',until:Date.now()-1}));
const expiredReplacement=failure.ctx.claimQuestionCacheLease('expired');
assert.notEqual(expiredReplacement.owner,'expired');
failure.ctx.releaseQuestionCacheLease(expiredReplacement);

// Cache-service failures remain visible while the requesting candidate can
// still receive the correctly loaded bank. Memo prevents repeated Drive reads
// within that run even when every cache write fails.
const outage=environment(banks), memo={banks:{},cacheStatus:{}};
outage.rejectWrites(true);
assert.equal(outage.ctx.loadQuestionsForLanguageServer('he',memo).length,banks.he.length);
assert.equal(outage.ctx.loadQuestionsForLanguageServer('he',memo).length,banks.he.length);
assert.equal(outage.reads.he,1);
assert.equal(memo.cacheStatus.he,false);
assert.ok(outage.logs.some(s=>s.includes('WRITE FAILED')));

// A missing index plus busy bank must bubble up retryability, not omit a language.
const busy=environment(banks);
busy.ctx.claimQuestionCacheLease('bank_he');
assert.throws(()=>busy.ctx.buildExamTranslations([{id:1}],false),e=>e.code==='question_cache_busy');
const optionalBanks={...banks}; delete optionalBanks.am;
const optional=environment(optionalBanks);
assert.deepEqual(json(optional.ctx.buildExamTranslations([{id:1}],false)),expectedTranslations(optionalBanks,[1],false),'genuinely absent optional language retains legacy behavior');
for (const optionalFailure of ['empty','malformed','unreadable']) {
  const altered={...banks,am:optionalFailure==='empty'?[]:'not-a-question-array'};
  const degraded=environment(altered);
  if (optionalFailure==='unreadable') degraded.beforeDrive(lang=>{if(lang==='am')throw new Error('optional Drive unavailable');});
  assert.deepEqual(json(degraded.ctx.buildExamTranslations([{id:1}],false)),expectedTranslations(optionalBanks,[1],false),'nonretryable '+optionalFailure+' optional bank retains baseline behavior');
  assert.ok(degraded.logs.some(line=>line.includes('optional language omitted am')));
}

// Emergency reset invalidates translations too and reloads all seven banks.
env.ctx.emergencyClearAndRefreshCache();
assert.deepEqual(env.reads,Object.fromEntries(langs.map(l=>[l,2])));
assert.equal(env.ctx.questionCacheStatus().ready,true);
console.log('PASS synthetic: 1,750 IDs; exact seven-language/ci parity; bounded UTF-8 cache; legacy cleanup; leases; failure recovery; reset');

if (process.argv[2]) {
  const directory=path.resolve(process.argv[2]), realBanks={};
  for (const lang of langs) realBanks[lang]=JSON.parse(fs.readFileSync(path.join(directory,'questions_'+lang+'.json'),'utf8'));
  const real=environment(realBanks);
  for (const id of [...new Set(Object.values(realBanks).flat().map(q=>q.id))].slice(0,1000)) real.cache.put('tx_'+id,'{}',21600);
  assert.equal(real.entries.size,1000);
  const summary=real.ctx.warmupQuestionCaches();
  assert.ok(summary.every(line=>!line.includes('ERROR')&&!line.includes('cached=false')),summary.join('\n'));
  assert.deepEqual(real.reads,Object.fromEntries(langs.map(l=>[l,1])));
  const realIds=[...new Set(Object.values(realBanks).flat().map(q=>q.id))];
  for (let i=0;i<realIds.length;i+=30) for (const ci of [false,true]) {
    const chunk=realIds.slice(i,i+30);
    assert.deepEqual(json(real.ctx.tryTranslationsFromIndex(chunk,ci)),expectedTranslations(realBanks,chunk,ci));
  }
  const status=real.ctx.questionCacheStatus();
  assert.equal(status.ready,true);
  assert.ok(status.presentKeys<=423);
  console.log('PASS optional real banks: '+JSON.stringify({uniqueIds:realIds.length,...json(status)}));
}
