#!/usr/bin/env node
/**
 * build_version.js — one hash per page, used for BOTH update checks and the
 * service-worker cache names.
 *
 * WHY: the 16/09 outage was an ETag comparison ("x" vs W/"x") that declared a
 * new version on every GitHub Pages push, whether or not the page had changed,
 * and reloaded examiners mid-session. A content hash cannot lie: a push that
 * did not change examinee.html leaves its hash alone, so nothing reloads.
 * The same hash names the SW cache, so a page that really did change also
 * invalidates its offline copy — one source of truth instead of two.
 *
 * Writes version.json and patches the cache-name constant in the four service
 * workers, changing nothing else in them (line endings included).
 * Idempotent: a second run reports "unchanged" and rewrites nothing.
 *
 * Usage: node tools/build_version.js
 */
'use strict';
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const { execFileSync } = require('child_process');

const ROOT = path.join(__dirname, '..');
const VERSION_FILE = path.join(ROOT, 'version.json');

// page -> service worker and the constant that names its cache
const PAGES = [
  { page: 'examinee.html', sw: 'sw-examinee.js', name: 'examinee', constant: 'CACHE' },
  { page: 'examiner.html', sw: 'sw-examiner.js', name: 'examiner', constant: 'CACHE_NAME' },
  { page: 'teacher.html', sw: 'sw-teacher.js', name: 'teacher', constant: 'CACHE_NAME' },
  { page: 'student.html', sw: 'sw-student.js', name: 'student', constant: 'CACHE_NAME' }
];

const sha1 = buf => crypto.createHash('sha1').update(buf).digest('hex');

function gitShortSha() {
  try {
    return execFileSync('git', ['rev-parse', '--short', 'HEAD'], { cwd: ROOT, encoding: 'utf8' }).trim() || 'local';
  } catch (e) {
    return 'local'; // a copy without git history still builds
  }
}

/** Replaces only the quoted value of `var <constant> = '...'`. */
function patchServiceWorker(swFile, constant, cacheName) {
  const file = path.join(ROOT, swFile);
  const before = fs.readFileSync(file, 'utf8');
  const pattern = new RegExp('(var\\s+' + constant + '\\b\\s*=\\s*\')([^\']*)(\')', 'g');
  const hits = before.match(pattern) || [];
  if (hits.length !== 1) {
    throw new Error('build_version: ' + swFile + ' has ' + hits.length + ' definitions of ' + constant + ', expected 1');
  }
  let previous = '';
  const after = before.replace(pattern, (m, head, old, tail) => { previous = old; return head + cacheName + tail; });
  if (after === before) return { file: swFile, changed: false, name: cacheName };
  fs.writeFileSync(file, after, 'utf8');
  return { file: swFile, changed: true, from: previous, name: cacheName };
}

function main() {
  const pages = {};
  const changes = [];
  for (const spec of PAGES) {
    const hash = sha1(fs.readFileSync(path.join(ROOT, spec.page)));
    pages[spec.page] = hash;
    const patched = patchServiceWorker(spec.sw, spec.constant, spec.name + '-' + hash.slice(0, 8));
    if (patched.changed) changes.push(spec.sw + ': ' + patched.from + ' -> ' + patched.name);
  }

  const previous = fs.existsSync(VERSION_FILE) ? JSON.parse(fs.readFileSync(VERSION_FILE, 'utf8')) : null;
  const build = gitShortSha();
  const same = previous && previous.build === build &&
    JSON.stringify(previous.pages) === JSON.stringify(pages);
  // generatedAt is kept when nothing else moved, so a second run is a no-op.
  const version = {
    build,
    generatedAt: same ? previous.generatedAt : new Date().toISOString(),
    pages
  };
  if (!same) fs.writeFileSync(VERSION_FILE, JSON.stringify(version, null, 2) + '\n', 'utf8');

  for (const spec of PAGES) console.log(spec.page.padEnd(15), pages[spec.page].slice(0, 8), '->', spec.sw);
  console.log('version.json ' + (same ? 'unchanged' : (previous ? 'updated' : 'created')) + ' (build ' + build + ')');
  console.log(changes.length ? 'service workers repointed:\n  ' + changes.join('\n  ') : 'service workers already current');
}

main();
