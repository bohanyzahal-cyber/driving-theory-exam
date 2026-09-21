#!/usr/bin/env node
/**
 * test.js — run every tests/*.test.cjs sequentially and summarise.
 *
 * Sequential on purpose: the suites read the same 25MB of question dumps and
 * a few of them shell out to git, so running them in parallel turns a 30s
 * gate into a memory spike with confusing interleaved output.
 *
 * Usage: node tools/test.js [name-fragment ...]
 */
'use strict';
const fs = require('fs');
const path = require('path');
const { spawnSync } = require('child_process');

const ROOT = path.join(__dirname, '..');
const TESTS = path.join(ROOT, 'tests');
const filters = process.argv.slice(2);

const files = fs.readdirSync(TESTS)
  .filter(f => f.endsWith('.test.cjs'))
  .filter(f => !filters.length || filters.some(x => f.includes(x)))
  .sort();

if (!files.length) {
  console.error('no tests matched ' + JSON.stringify(filters));
  process.exit(1);
}

const results = [];
for (const file of files) {
  console.log('\n=== ' + file + ' ===');
  const started = Date.now();
  const run = spawnSync(process.execPath, [path.join(TESTS, file)], { cwd: ROOT, stdio: 'inherit' });
  results.push({ file, ok: run.status === 0, status: run.status, ms: Date.now() - started });
}

const width = Math.max(...results.map(r => r.file.length));
console.log('\n' + '-'.repeat(width + 20));
for (const r of results) {
  console.log((r.ok ? 'PASS  ' : 'FAIL  ') + r.file.padEnd(width) + '  ' + (r.ms / 1000).toFixed(1) + 's' +
    (r.ok ? '' : '  (exit ' + r.status + ')'));
}
const failed = results.filter(r => !r.ok);
console.log('-'.repeat(width + 20));
console.log(results.length - failed.length + '/' + results.length + ' suites passed');
process.exit(failed.length ? 1 : 0);
