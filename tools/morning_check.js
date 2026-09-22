#!/usr/bin/env node
/**
 * morning_check.js — the exam-morning 07:00 check (OPERATIONS §10) in one
 * command, plus the "us or Google?" two-hop probe of KNOWN_ISSUES #35.
 *
 *   node tools/morning_check.js            # Worker + both Apps Script projects
 *   node tools/morning_check.js --deep     # health&deep=1 (one sheet read each)
 *   node tools/morning_check.js --hops 3   # repeat the two-hop probe N times
 *
 * What it prints, per target:
 *   Worker   /                 -> build + bank id (bank empty = assets missing)
 *   exam     hop1 /exec        -> the execution itself (302 in ~1-1.5 s is healthy)
 *            hop2 echo         -> Google's delivery hop; 25-60 s or a 404/HTML
 *                                 page here with a fast hop1 = KNOWN_ISSUES #35,
 *                                 nothing on our side to fix
 *   reports  same two hops
 *
 * Verdicts: OK / SLOW-DELIVERY (hop1 fast, hop2 slow or not JSON) / SLOW-FRONT
 * (hop1 itself slow: Google's front door + queueing + our execution — compare
 * with the Executions page: a short run there = Google again) / DOWN. Nothing here writes anything anywhere: every call
 * is `action=health`, which is the cheapest read the server has.
 *
 * Node 22, no dependencies (global fetch; redirects followed by hand so the
 * two hops are timed separately).
 */
'use strict';

const WORKER = 'https://session-gateway.bohanyzahal.workers.dev/';
const TARGETS = {
  exam: 'https://script.google.com/macros/s/AKfycbzOI0zrDEngP-GvlRblhOk8tQsYBvWZ2gGliIQHTpS67WrDZl4la8NPpwtJr_Vjsh3Gzg/exec',
  reports: 'https://script.google.com/macros/s/AKfycbw7FwTioHoEMvl6Plk-IlHii1rb3FSs9CXan-8lCVP5K7FTuz594rsEOc2y6LDVS-DXcQ/exec'
};
const EXPECT_BUILD_PREFIX = '2026-09-2';   // any r31/r32 build of this September
const HOP_TIMEOUT_MS = 90000;

const args = process.argv.slice(2);
const deep = args.includes('--deep');
const hopsArg = args.indexOf('--hops');
const rounds = hopsArg >= 0 ? Math.max(1, Number(args[hopsArg + 1]) || 1) : 1;

const ms = t => (t / 1000).toFixed(2) + 's';
const pad = (s, n) => String(s).padEnd(n);

async function timedFetch(url, init) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), HOP_TIMEOUT_MS);
  const t0 = Date.now();
  try {
    const res = await fetch(url, Object.assign({ signal: controller.signal, cache: 'no-store' }, init));
    const text = await res.text();
    return { status: res.status, headers: res.headers, text, ms: Date.now() - t0 };
  } catch (e) {
    return { status: 0, headers: new Headers(), text: '', ms: Date.now() - t0, error: e && e.name === 'AbortError' ? 'timeout' : String(e && e.message || e) };
  } finally {
    clearTimeout(timer);
  }
}

function parseJson(text) {
  try { const v = JSON.parse(text); return v && typeof v === 'object' ? v : null; } catch (e) { return null; }
}

async function checkWorker() {
  const r = await timedFetch(WORKER);
  const body = parseJson(r.text);
  const ok = Boolean(r.status === 200 && body && body.status === 'ok' && body.bank);
  console.log(pad('Worker /', 14) + pad(ok ? 'OK' : 'DOWN', 15) + ms(r.ms) +
    (body ? '  build=' + body.build + '  bank=' + String(body.bank || '').slice(0, 10) + (body.bank ? '' : '  <-- bank EMPTY: assets not deployed') : '  ' + (r.error || 'HTTP ' + r.status)));
  return ok;
}

/** hop 1 = /exec (302 with Location), hop 2 = the echo that carries the body. */
async function twoHop(name, base) {
  const url = base + '?action=health' + (deep ? '&deep=1' : '') + '&origin=examinee-app&_t=' + Date.now();
  const hop1 = await timedFetch(url, { redirect: 'manual' });
  const location = hop1.headers.get('location') || '';
  let line = pad(name + ' hop1', 14) + pad(hop1.status + (location ? ' ->echo' : ''), 15) + ms(hop1.ms);
  if (!location) {
    // Google answered the body straight from /exec (it does that sometimes),
    // or refused. Judge what we have.
    const body = parseJson(hop1.text);
    const verdict = hop1.status === 0 ? 'DOWN' : body && body.status === 'ok' ? 'OK' : 'DOWN';
    console.log(line + '  ' + verdict + (body ? '  build=' + body.build + ' deployment=' + body.deployment : '  ' + (hop1.error || 'no JSON, HTTP ' + hop1.status)));
    return verdict;
  }
  console.log(line);
  const hop2 = await timedFetch(location);
  const body = parseJson(hop2.text);
  const isJson = Boolean(body && body.status === 'ok');
  let verdict;
  if (hop1.ms > 8000) verdict = 'SLOW-FRONT';
  else if (!isJson || hop2.ms > 8000) verdict = isJson ? 'SLOW-DELIVERY' : 'SLOW-DELIVERY (no JSON: HTTP ' + hop2.status + ')';
  else verdict = 'OK';
  const detail = body
    ? '  build=' + body.build + ' deployment=' + body.deployment +
      (deep ? ' sheetMs=' + body.sheetMs + (body.sheetError ? ' sheetError=' + body.sheetError : '') : '') +
      ' gateway=' + (body.gateway ? body.gateway.url + '/' + body.gateway.key : '?') +
      (String(body.build || '').indexOf(EXPECT_BUILD_PREFIX) === 0 ? '' : '  <-- unexpected build')
    : '  ' + (hop2.error || ('HTTP ' + hop2.status + ' ' + (hop2.text || '').replace(/\s+/g, ' ').slice(0, 60)));
  console.log(pad(name + ' hop2', 14) + pad(verdict, 15) + ms(hop2.ms) + detail);
  return verdict;
}

(async () => {
  console.log('morning check ' + new Date().toISOString() + (deep ? ' (deep)' : '') + ' — nothing here writes anything');
  const worker = await checkWorker();
  const verdicts = { worker };
  for (let round = 1; round <= rounds; round++) {
    if (rounds > 1) console.log('-- round ' + round + '/' + rounds);
    for (const name of Object.keys(TARGETS)) verdicts[name + round] = await twoHop(name, TARGETS[name]);
  }
  const bad = Object.entries(verdicts).filter(([, v]) => v !== true && v !== 'OK');
  console.log(bad.length ? '\nATTENTION: ' + bad.map(([k, v]) => k + '=' + v).join(', ') +
    '\n  SLOW-DELIVERY with a fast hop1 = Google (#35): do not deploy, do not refresh, do not re-click.' +
    '\n  SLOW-FRONT = hop1 itself slow: open Executions - a run of 0.6-4 s there means Google queued/delivered it, not us.' +
    '\n  DOWN on the Worker = cloudflarestatus.com, then DEPLOY_2026-09-22.md §8 if it stays down.'
    : '\nAll good: run the exams. Do not deploy anything today (OPERATIONS §10).');
})();
