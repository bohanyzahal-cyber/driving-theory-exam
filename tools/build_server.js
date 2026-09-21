#!/usr/bin/env node
// Assembles the files that are pasted into the Apps Script editor from the
// modules in server/src/, in the order listed by server/BUILD_ORDER.json. The
// pasted file stays ONE file because the deploy is a single full-file paste; the
// source stays many files because a 9,000-line file is where duplicated logic
// hides.
//
//   node tools/build_server.js            → writes the three outputs below
//   node tools/build_server.js --check    → exits 1 if any committed output is stale
//
// THREE outputs (DESIGN_2026-09-21 §13.3), all committed, all built from the
// same sources, each one carrying its own `var API_DEPLOYMENT`:
//   external_exam_apps_script.js          every module          'all'
//   external_exam_apps_script.exam.js     both + exam           'exam'
//   external_exam_apps_script.reports.js  both + reports        'reports'
// server/BUILD_TARGETS.json assigns every module to exactly one of both/exam/
// reports; the split exists because Google loads and compiles the WHOLE file on
// every request, and the hot (exam) path never runs reports, teachers, practice,
// prediction or archiving (measured 21/09/2026: an empty action = 2.7 s, 1.2 s
// of it ours).
//
// Two things happen to a module's text on the way in, and both are measured in
// milliseconds of Google's per-request compile (§13.1):
//   1. Generated data: a module may contain the line `// @@QUESTION_INDEX@@`; it
//      is replaced by `var QUESTION_INDEX_PACKED = {...};` built from
//      deployment/question_index.json (tools/build_bank.js writes that file) —
//      ~11 KB of packed string instead of 132 KB of object literal, unpacked
//      once per execution by questionIndex() in 70_questions.js. The line
//      `// @@API_DEPLOYMENT@@` is replaced the same way, per output.
//      The answer key stays a separate pasted file (deployment/answer_key.gs)
//      because it is secret and changes rarely.
//   2. Comments are removed with a REAL tokenizer (tools/vendor/acorn.js, MIT,
//      vendored so a clean tree needs no npm install) — never a regex, which
//      cannot tell a comment from `'// not a comment'` inside a string. The
//      output is then parsed again and its AST compared to the AST of the input
//      with the position fields removed: if stripping changed the program in any
//      way (an ASI edge case, a cut that ate code) the build FAILS. Only the
//      copyright block survives, and the GENERATED header is added afterwards.
'use strict';
const fs = require('fs');
const path = require('path');
const assert = require('assert');
const { execFileSync } = require('child_process');
const acorn = require('./vendor/acorn.js');

const ROOT = path.resolve(__dirname, '..');
const SRC_DIR = path.join(ROOT, 'server', 'src');
const ORDER_FILE = path.join(ROOT, 'server', 'BUILD_ORDER.json');
const TARGETS_FILE = path.join(ROOT, 'server', 'BUILD_TARGETS.json');
const INDEX_JSON = path.join(ROOT, 'deployment', 'question_index.json');
const INDEX_MARKER = '// @@QUESTION_INDEX@@';
const DEPLOYMENT_MARKER = '// @@API_DEPLOYMENT@@';

// The five licenses and the five blueprint topics, in the order the packed
// record encodes them. Fixed here (not derived from the data) so that a build
// on a tree whose index happens to be missing a license cannot silently shift
// every digit of every record.
const PACK_LICENSES = ['B', '1', 'C1', 'C', 'D'];
const PACK_TOPICS = ['בטיחות', 'הכרת הרכב', 'חוק', 'תמרורים', 'ספציפי'];
const PACK_LANG_DEFAULT = 127;   // all seven languages — 1,693 of the 1,700 ids

const OUTPUTS = [
  { deployment: 'all', file: 'external_exam_apps_script.js', groups: ['both', 'exam', 'reports'],
    paste: ['Deploy (monolith option, DESIGN §13.3 step ג\'): paste this whole file into the EXISTING',
      'Apps Script project → Deploy → Manage deployments → New version. One project, every action.'] },
  { deployment: 'exam', file: 'external_exam_apps_script.exam.js', groups: ['both', 'exam'],
    paste: ['Deploy (split, step ג): paste this whole file into the EXISTING Apps Script project — the',
      'URL every page already knows → New version. health must answer "deployment":"exam".',
      'Run uninstallNightlyJobs() once here: the nightly jobs move to the reports project.',
      'A reports action that reaches this deployment is answered wrong_deployment, not a broken page.'] },
  { deployment: 'reports', file: 'external_exam_apps_script.reports.js', groups: ['both', 'reports'],
    paste: ['Deploy (split, step א): paste this whole file into a NEW STANDALONE Apps Script project,',
      'together with deployment/answer_key.gs (practice scores from it). Script properties:',
      'EXAM_SPREADSHEET_ID (required — a standalone script has no active spreadsheet),',
      'PRACTICE_SPREADSHEET_ID, GATEWAY_KEY, GATEWAY_URL, OFFICE_WHATSAPP_NUMBER (optional).',
      'Deploy as Web app (Execute as: Me, Who has access: Anyone) → health must answer',
      '"deployment":"reports" → run installNightlyJobs() once.'] }
];

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

// Every module belongs to exactly one target and every target names only modules
// that exist: a module that fell out of BUILD_TARGETS.json would silently vanish
// from BOTH split files while the monolith kept working, which is the one kind
// of split bug a paste cannot reveal.
function readTargets(order) {
  const targets = JSON.parse(fs.readFileSync(TARGETS_FILE, 'utf8'));
  const seen = {}, problems = [];
  for (const group of ['both', 'exam', 'reports']) {
    if (!Array.isArray(targets[group])) throw new Error('server/BUILD_TARGETS.json is missing the "' + group + '" list');
    for (const name of targets[group]) {
      if (seen[name]) problems.push(name + ' is in both "' + seen[name] + '" and "' + group + '"');
      seen[name] = group;
      if (order.indexOf(name) === -1) problems.push(name + ' is in "' + group + '" but not in BUILD_ORDER.json');
    }
  }
  for (const name of order) if (!seen[name]) problems.push(name + ' is in BUILD_ORDER.json but in no target');
  if (problems.length) throw new Error('server/BUILD_TARGETS.json: ' + problems.join('; '));
  return seen;
}

// ---- the question index, packed --------------------------------------------
// 1,700 objects (132 K characters that Google re-parses on every request) become
// one string of 6 characters per id: five topic digits (0 = not in that license,
// 1-5 = index in PACK_TOPICS + 1), then the image digit. The language mask is
// carried separately because all but 7 ids share the same one.
// 70_questions.js questionIndex() is the only reader, and it rebuilds EXACTLY
// { '<id>': { c: {license: topic}, l: mask, img: 0|1 } } — tests/server_split
// proves the round trip against deployment/question_index.json.
let packedIndexMemo = null;
function packedIndexSource() {
  if (packedIndexMemo !== null) return packedIndexMemo;
  if (!fs.existsSync(INDEX_JSON)) return (packedIndexMemo = false);
  const index = JSON.parse(fs.readFileSync(INDEX_JSON, 'utf8'));
  const ids = Object.keys(index);
  let max = 0;
  for (const id of ids) {
    const n = Number(id);
    if (!Number.isInteger(n) || n < 1) throw new Error('question_index.json: id ' + JSON.stringify(id) + ' is not a positive integer');
    if (n > max) max = n;
  }
  const stride = PACK_LICENSES.length + 1;
  const records = [], lang = {};
  let present = 0;
  for (let id = 1; id <= max; id++) {
    const entry = index[String(id)];
    if (!entry) { records.push('0'.repeat(stride)); continue; }
    for (const license in entry.c) {
      if (!Object.prototype.hasOwnProperty.call(entry.c, license)) continue;
      if (PACK_LICENSES.indexOf(license) === -1) throw new Error('question ' + id + ': unknown license ' + JSON.stringify(license));
    }
    let digits = '', licensed = 0;
    for (const license of PACK_LICENSES) {
      const topic = entry.c[license] || '';
      if (!topic) { digits += '0'; continue; }
      const at = PACK_TOPICS.indexOf(topic);
      if (at === -1) throw new Error('question ' + id + ': unknown topic ' + JSON.stringify(topic));
      digits += String(at + 1);
      licensed++;
    }
    // An id in the index with no license at all would encode as '000000' —
    // indistinguishable from "no such question", so the whole entry would
    // disappear from the deployed server and the exam would be drawn short.
    if (!licensed) throw new Error('question ' + id + ' is in the index with no license — packing it would make it vanish');
    if (entry.img !== 0 && entry.img !== 1) throw new Error('question ' + id + ': img must be 0 or 1, got ' + JSON.stringify(entry.img));
    records.push(digits + String(entry.img));
    if (entry.l !== PACK_LANG_DEFAULT) lang[String(id)] = entry.l;
    present++;
  }
  if (present !== ids.length) throw new Error('packed index holds ' + present + ' of ' + ids.length + ' ids');
  const rec = records.join('');
  if (rec.length !== max * stride) throw new Error('packed record is ' + rec.length + ' chars, expected ' + (max * stride));
  packedIndexMemo = 'var QUESTION_INDEX_PACKED = {\n' +
    'max: ' + max + ',\n' +
    'lic: ' + JSON.stringify(PACK_LICENSES) + ',\n' +
    'topics: ' + JSON.stringify(PACK_TOPICS) + ',\n' +
    'rec: ' + JSON.stringify(rec) + ',\n' +
    'lang: ' + JSON.stringify(lang) + ',\n' +
    'langDefault: ' + PACK_LANG_DEFAULT + '\n};';
  return packedIndexMemo;
}

// ---- comment stripping ------------------------------------------------------
// The tokenizer reports every comment's exact range; draining the token stream is
// what makes onComment fire for the whole file. A regex would eat '//' inside a
// string or a regex literal, which is why this is a real parse.
function commentsOf(src, label) {
  const found = [];
  try {
    const tokenizer = acorn.tokenizer(src, { ecmaVersion: 2022,
      onComment: (block, text, start, end) => { found.push({ block: block, text: text, start: start, end: end }); } });
    for (const token of tokenizer) { void token; }
  } catch (err) {
    throw new Error('acorn could not tokenize ' + label + ': ' + err.message);
  }
  return found;
}

// A copyright notice is a BLOCK, not a line: a comment containing '©' keeps the
// unbroken run of line comments under it, so "All Rights Reserved" never ships
// without "Unauthorized copying ... is prohibited". In practice that is the
// whole opening header of 00_config.js (notice + the original paste
// instructions), ~350 bytes, and keeping it whole is worth more than the bytes.
function keepFlagsFor(src, comments) {
  const keep = comments.map(c => c.text.indexOf('©') >= 0);
  for (let i = 1; i < comments.length; i++) {
    if (keep[i] || !keep[i - 1] || comments[i - 1].block || comments[i].block) continue;
    if (/^\n[ \t]*$/.test(src.slice(comments[i - 1].end, comments[i].start))) keep[i] = true;
  }
  return keep;
}

function stripComments(src, label) {
  const comments = commentsOf(src, label);
  const keep = keepFlagsFor(src, comments);
  // NUL marks where a comment was, so a line that held nothing else can be told
  // from a line the source itself left blank (those are kept: they are the
  // paragraphing of the code).
  let cut = src;
  for (let i = comments.length - 1; i >= 0; i--) {
    if (keep[i]) continue;
    cut = cut.slice(0, comments[i].start) + '\u0000' + cut.slice(comments[i].end);
  }
  const out = [];
  for (const line of cut.split('\n')) {
    if (line.indexOf('\u0000') === -1) { out.push(line); continue; }
    const bare = line.split('\u0000').join('').replace(/\s+$/, '');
    if (bare === '') continue;
    out.push(bare);
  }
  const stripped = out.join('\n');
  assertSameProgram(src, stripped, label);
  return stripped;
}

function parseProgram(src, label) {
  try { return acorn.parse(src, { ecmaVersion: 2022, sourceType: 'script' }); }
  catch (err) { throw new Error('acorn could not parse ' + label + ': ' + err.message); }
}

// Positions are the ONLY thing stripping is allowed to change.
function withoutPositions(node) {
  if (Array.isArray(node)) { for (const item of node) withoutPositions(item); return node; }
  if (!node || typeof node !== 'object') return node;
  delete node.start; delete node.end; delete node.loc; delete node.range;
  for (const key in node) if (Object.prototype.hasOwnProperty.call(node, key)) withoutPositions(node[key]);
  return node;
}

function assertSameProgram(before, after, label) {
  const a = withoutPositions(parseProgram(before, label + ' (before stripping)'));
  const b = withoutPositions(parseProgram(after, label + ' (after stripping)'));
  if (JSON.stringify(a) === JSON.stringify(b)) return;
  try { assert.deepStrictEqual(b, a); }
  catch (err) { throw new Error('stripping comments CHANGED the program in ' + label + ' — ' + err.message); }
  throw new Error('stripping comments changed ' + label + ' in a way deepStrictEqual cannot show');
}

// ---- assembly ---------------------------------------------------------------
// One line-ending convention for the pasted file whatever a module was saved
// with (a CRLF module used to leave the generated file mixed).
const rawMemo = new Map();
function rawModule(name) {
  if (!rawMemo.has(name)) {
    rawMemo.set(name, fs.readFileSync(path.join(SRC_DIR, name + '.js'), 'utf8').replace(/\r\n?/g, '\n'));
  }
  return rawMemo.get(name);
}

const moduleMemo = new Map();
function moduleText(name, deployment) {
  const key = name + '\u0000' + deployment;
  if (moduleMemo.has(key)) return moduleMemo.get(key);
  let text = rawModule(name);
  // Markers are comments, so every marker is replaced BEFORE stripping runs.
  if (text.includes(INDEX_MARKER)) {
    const generated = packedIndexSource();
    if (generated === false) throw new Error(name + '.js expects deployment/question_index.json — run node tools/build_bank.js first');
    text = text.replace(INDEX_MARKER, generated);
  }
  if (text.includes(DEPLOYMENT_MARKER)) {
    text = text.replace(DEPLOYMENT_MARKER, 'var API_DEPLOYMENT = ' + JSON.stringify(deployment) + ';');
  }
  text = stripComments(text, name + '.js');
  if (!text.endsWith('\n')) text += '\n';
  moduleMemo.set(key, text);
  return text;
}

function header(output, modules) {
  return [
    '// ============================================================================',
    '// GENERATED FILE — do not edit here. Source: server/src/*.js (order: server/BUILD_ORDER.json,',
    '// split: server/BUILD_TARGETS.json). Comments are stripped and the question index is packed',
    '// by the build; read the SOURCE, not this.',
    '// Rebuild with:  node tools/build_server.js     (tools/build.js runs it too)',
    '// Deployment: ' + output.deployment + '   (API_DEPLOYMENT — health reports it)'
  ].concat(output.paste.map(line => '// ' + line)).concat([
    '// Modules: ' + modules.join(', '),
    '// ============================================================================',
    ''
  ]).join('\n');
}

function build(output, order, targetOf) {
  const modules = order.filter(name => output.groups.indexOf(targetOf[name]) !== -1);
  if (output.deployment !== 'all' && modules.length === order.length) {
    throw new Error(output.file + ' would carry every module — BUILD_TARGETS.json is not splitting anything');
  }
  // Both markers must appear EXACTLY once across the modules of every output: a
  // missing @@API_DEPLOYMENT@@ leaves API_DEPLOYMENT undefined, which makes the
  // dispatcher answer wrong_deployment to every action it has a target for.
  for (const marker of [INDEX_MARKER, DEPLOYMENT_MARKER]) {
    const seen = modules.filter(name => rawModule(name).includes(marker));
    if (seen.length !== 1) {
      throw new Error(marker + ' appears in ' + seen.length + ' of the modules of ' + output.file +
        ' [' + seen.join(', ') + '] — it must appear in exactly one');
    }
  }
  return header(output, modules) + modules.map(name => moduleText(name, output.deployment)).join('');
}

function syntaxCheck(file) {
  execFileSync(process.execPath, ['--check', file], { stdio: 'inherit' });
}

function main() {
  const order = readOrder();
  const targetOf = readTargets(order);
  const check = process.argv.includes('--check');
  let stale = [];
  for (const output of OUTPUTS) {
    const text = build(output, order, targetOf);
    const file = path.join(ROOT, output.file);
    if (check) {
      const current = fs.existsSync(file) ? fs.readFileSync(file, 'utf8') : '';
      if (current !== text) { stale.push(output.file); continue; }
      console.log(output.file + ' is up to date (' + Buffer.byteLength(text) + ' bytes)');
      continue;
    }
    fs.writeFileSync(file, text);
    syntaxCheck(file);
    console.log('wrote ' + output.file + ': ' + Buffer.byteLength(text) + ' bytes, ' +
      text.split('\n').length + ' lines, ' + output.groups.join('+'));
  }
  if (check && stale.length) {
    console.error(stale.join(', ') + ' — stale, run node tools/build_server.js');
    process.exit(1);
  }
}

main();
