#!/usr/bin/env node
// Assembles external_exam_apps_script.js (the file that is pasted into the Apps
// Script editor) from the modules in server/src/, in the order listed by
// server/BUILD_ORDER.json. The pasted file stays ONE file because the deploy is a
// single full-file paste; the source stays many files because a 9,000-line file
// is where duplicated logic hides.
//
//   node tools/build_server.js            → writes external_exam_apps_script.js
//   node tools/build_server.js --check    → exits 1 if the committed output is stale
//
// Generated data: a module may contain the line `// @@QUESTION_INDEX@@`; it is
// replaced by `var QUESTION_INDEX = {...};` built from deployment/question_index.json
// (tools/build_bank.js writes that file). The answer key stays a separate pasted
// file (deployment/answer_key.gs) because it is secret and changes rarely.
'use strict';
const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');

const ROOT = path.resolve(__dirname, '..');
const SRC_DIR = path.join(ROOT, 'server', 'src');
const ORDER_FILE = path.join(ROOT, 'server', 'BUILD_ORDER.json');
const OUT = path.join(ROOT, 'external_exam_apps_script.js');
const INDEX_JSON = path.join(ROOT, 'deployment', 'question_index.json');
const INDEX_MARKER = '// @@QUESTION_INDEX@@';

function readOrder() {
  const order = JSON.parse(fs.readFileSync(ORDER_FILE, 'utf8'));
  const onDisk = fs.readdirSync(SRC_DIR).filter(f => f.endsWith('.js')).map(f => f.slice(0, -3)).sort();
  const listed = order.slice().sort();
  const missing = listed.filter(n => !onDisk.includes(n));
  const unlisted = onDisk.filter(n => !listed.includes(n));
  if (missing.length || unlisted.length) {
    throw new Error('server/BUILD_ORDER.json and server/src disagree — missing: [' + missing.join(', ') + '] unlisted: [' + unlisted.join(', ') + ']');
  }
  return order;
}

function questionIndexSource() {
  if (!fs.existsSync(INDEX_JSON)) return null;
  const index = JSON.parse(fs.readFileSync(INDEX_JSON, 'utf8'));
  const ids = Object.keys(index);
  // One line per id keeps the pasted file diff-able and the editor responsive.
  const body = ids.map(id => JSON.stringify(id) + ':' + JSON.stringify(index[id])).join(',\n');
  return '// Generated from deployment/question_index.json by tools/build_server.js — ' + ids.length + ' ids. Do not edit.\n' +
    'var QUESTION_INDEX = {\n' + body + '\n};';
}

function header(order) {
  return [
    '// ============================================================================',
    '// GENERATED FILE — do not edit here. Source: server/src/*.js (order: server/BUILD_ORDER.json).',
    '// Rebuild with:  node tools/build_server.js     (tools/build.js runs it too)',
    '// Deploy: paste this whole file into the Apps Script editor → Deploy → Manage deployments → New version.',
    '// Modules: ' + order.join(', '),
    '// ============================================================================',
    ''
  ].join('\n');
}

function build() {
  const order = readOrder();
  let markerSeen = 0;
  const parts = order.map(name => {
    // One line-ending convention for the pasted file whatever a module was
    // saved with (a CRLF module used to leave the generated file mixed).
    let text = fs.readFileSync(path.join(SRC_DIR, name + '.js'), 'utf8').replace(/\r\n?/g, '\n');
    if (text.includes(INDEX_MARKER)) {
      markerSeen++;
      const generated = questionIndexSource();
      if (!generated) throw new Error(name + '.js expects deployment/question_index.json — run node tools/build_bank.js first');
      text = text.replace(INDEX_MARKER, generated);
    }
    if (!text.endsWith('\n')) text += '\n';
    return text;
  });
  if (markerSeen > 1) throw new Error('the QUESTION_INDEX marker appears in more than one module');
  return header(order) + parts.join('');
}

function syntaxCheck(file) {
  execFileSync(process.execPath, ['--check', file], { stdio: 'inherit' });
}

function main() {
  const output = build();
  const check = process.argv.includes('--check');
  if (check) {
    const current = fs.existsSync(OUT) ? fs.readFileSync(OUT, 'utf8') : '';
    if (current !== output) { console.error('external_exam_apps_script.js is stale — run node tools/build_server.js'); process.exit(1); }
    console.log('external_exam_apps_script.js is up to date (' + output.length + ' bytes)');
    return;
  }
  fs.writeFileSync(OUT, output);
  syntaxCheck(OUT);
  console.log('wrote external_exam_apps_script.js: ' + output.length + ' bytes, ' + output.split('\n').length + ' lines');
}

main();
