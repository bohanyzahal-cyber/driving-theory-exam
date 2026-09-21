#!/usr/bin/env node
/**
 * build.js — produce everything that gets deployed, in dependency order.
 *
 *   1. build_server.js  — assembles server/src/*.js into external_exam_apps_script.js
 *                         (it injects deployment/question_index.json, so the bank
 *                         build must be able to run before it on a clean tree —
 *                         but the index only changes when the dumps change)
 *   2. build_bank.js    — deployment/question_index.json (committed) and the
 *                         PRIVATE gateway assets:
 *                         cloudflare-workers/session-gateway/assets/{q,bank,manifest}
 *   3. build_version.js — version.json + the four service-worker cache names
 *
 * Stops at the first failure: a half-built release is worse than none.
 *
 * Two deploy targets come out of this, and they are separate: Pages gets the
 * pages (git push), the Worker gets the assets (`npx wrangler deploy` from
 * cloudflare-workers/session-gateway). A tree without deployment/generated/
 * still builds — but build_bank.js warns, because it cannot produce the assets
 * and the Worker would be deployed with an empty bank.
 * Usage: node tools/build.js
 */
'use strict';
const fs = require('fs');
const path = require('path');
const { spawnSync } = require('child_process');

const ROOT = path.join(__dirname, '..');
const STEPS = [
  { script: 'tools/build_server.js', optional: true },
  { script: 'tools/build_bank.js' },
  { script: 'tools/build_version.js' }
];

let failed = null;
for (const step of STEPS) {
  const file = path.join(ROOT, step.script);
  if (!fs.existsSync(file)) {
    if (step.optional) { console.warn('SKIP  ' + step.script + ' — not present yet'); continue; }
    failed = step.script + ' is missing';
    break;
  }
  console.log('\n=== ' + step.script + ' ===');
  const run = spawnSync(process.execPath, [file], { cwd: ROOT, stdio: 'inherit' });
  if (run.status !== 0) { failed = step.script + ' exited with ' + run.status; break; }
}

console.log('');
if (failed) {
  console.error('BUILD FAILED: ' + failed);
  process.exit(1);
}
console.log('BUILD OK — run `node tools/test.js` before deploying anything.');
