// The 26/09/2026 cut-over: the exams and the practice run on the new system,
// and every old entry point of this site sends people there. Only exam.html
// (the standalone audio exam) stays.
//
//   examiner.html, examiner/index.html (the installed app), find_image.html,
//   report.html   -> https://teoria-digital-vitaly.com/examiner/
//                    (examiners and commanders both sign in there)
//   examinee.html (every old link and QR, ?code=...), examinee/index.html
//                 -> https://teoria-digital-vitaly.com/exam/, without the old
//                    session code (the new system does not know it)
//   teacher.html, teacher/index.html -> https://teoria-digital-vitaly.com/teacher/
//   student.html, student/index.html -> https://teoria-digital-vitaly.com/student/
//   admin.html (the practice admin board) -> https://teoria-digital-vitaly.com/admin/
//   Nothing from an old URL is carried over.
//
// Each redirect page runs for real in a vm, in the three ways it can load: on
// its own, inside the old installed app's iframe (so it must move the TOP
// window), and inside a frame that may not move its top (so it moves itself).
// It must register no service worker and touch no storage: what a device kept
// from the old pages stays exactly as it was.
//
// version.json must describe the pages as they are: that is how a page that is
// already open learns there is a new version (ExamTransport.createUpdateCheck)
// and reloads - into the redirect.
'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const crypto = require('node:crypto');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const app = path.resolve(__dirname, '..');
const read = rel => fs.readFileSync(path.join(app, rel), 'utf8');

const EXAMINER = 'https://teoria-digital-vitaly.com/examiner/';
const EXAM = 'https://teoria-digital-vitaly.com/exam/';
const TEACHER = 'https://teoria-digital-vitaly.com/teacher/';
const STUDENT = 'https://teoria-digital-vitaly.com/student/';
const ADMIN = 'https://teoria-digital-vitaly.com/admin/';
const REDIRECTS = {
  'examiner.html': EXAMINER,
  'examiner/index.html': EXAMINER,
  'find_image.html': EXAMINER,
  'report.html': EXAMINER,
  'examinee.html': EXAM,
  'examinee/index.html': EXAM,
  'teacher.html': TEACHER,
  'teacher/index.html': TEACHER,
  'student.html': STUDENT,
  'student/index.html': STUDENT,
  'admin.html': ADMIN
};
const MOVED_LINE = 'המערכת עברה לכתובת חדשה';   // המערכת עברה לכתובת חדשה

function inlineScripts(rel, html) {
  const scripts = [];
  const re = /<script\b([^>]*)>([\s\S]*?)<\/script>/gi;
  for (let m = re.exec(html); m; m = re.exec(html)) {
    assert.ok(!/\bsrc\s*=/i.test(m[1]), rel + ' loads no script file');
    scripts.push(m[2]);
  }
  assert.ok(scripts.length, rel + ' has its redirect script');
  return scripts;
}

// Runs the page's scripts against a window that records every navigation and
// every reach for a service worker or storage (even one the page would catch).
function load(rel, { framed = false, topRefuses = false, search = '' } = {}) {
  const navigations = [], touched = [];
  const locationOf = who => ({
    href: 'https://bohanyzahal-cyber.github.io/driving-theory-exam/' + rel + search,
    search,
    replace(url) {
      if (who === 'top' && topRefuses) throw new Error('SecurityError: this frame may not navigate its top');
      navigations.push({ who, url: String(url) });
    },
    assign(url) { navigations.push({ who, url: String(url), assign: true }); }
  });
  const navigator = {};
  Object.defineProperty(navigator, 'serviceWorker', {
    get() { touched.push('navigator.serviceWorker'); throw new Error(rel + ' touched navigator.serviceWorker'); }
  });
  const ctx = vm.createContext({ location: locationOf('self'), navigator, __touch: name => touched.push(name),
                                 __top: framed ? { location: locationOf('top') } : null });
  // window, top and the storage traps live on the context's own global: node's
  // vm does not run getters that are set on the sandbox object.
  vm.runInContext('var window = globalThis, self = globalThis, top = __top || globalThis, parent = top;' +
    "['localStorage', 'sessionStorage', 'indexedDB', 'caches'].forEach(function (name) {" +
    "  Object.defineProperty(globalThis, name, { get: function () { __touch(name); throw new Error('touched ' + name); } });" +
    '});', ctx);
  for (const src of inlineScripts(rel, read(rel))) vm.runInContext(src, ctx, { filename: rel });
  assert.deepEqual(touched, [], rel + ' reached for a service worker or storage');
  return navigations;
}

for (const [rel, target] of Object.entries(REDIRECTS)) {
  test(rel + ' sends every visitor to ' + target, () => {
    const html = read(rel);
    assert.ok(html.includes('CUTOVER_REDIRECT'), 'the ASCII marker a deploy check greps for');
    // Opened directly: a bookmark, an old link, the old QR.
    assert.deepEqual(load(rel), [{ who: 'self', url: target }]);
    // Inside the old installed app's iframe: the whole window moves, not the frame.
    assert.deepEqual(load(rel, { framed: true }), [{ who: 'top', url: target }]);
    // A frame that may not move its top still leaves this page.
    assert.deepEqual(load(rel, { framed: true, topRefuses: true }), [{ who: 'self', url: target }]);
    // The line people see if the automatic move does not happen, and its link.
    assert.ok(html.includes(MOVED_LINE));
    assert.ok(html.includes('<a href="' + target + '" target="_top"'), 'the link opens the new address in the whole window');
    // Nothing that reaches the shared worker scope or the practice pages' storage.
    assert.ok(!/serviceWorker|localStorage|sessionStorage|rel="manifest"/.test(html));
  });
}

test('an old examinee link or QR lands on the new exam page without its old session code', () => {
  for (const rel of ['examinee.html', 'examinee/index.html']) {
    assert.deepEqual(load(rel, { search: '?code=AB12CD34' }), [{ who: 'self', url: EXAM }]);
    assert.deepEqual(load(rel, { framed: true, search: '?code=AB12CD34&gw=google' }), [{ who: 'top', url: EXAM }]);
  }
});

test('no redirect carries anything from the old URL', () => {
  for (const [rel, target] of Object.entries(REDIRECTS)) {
    assert.deepEqual(load(rel, { search: '?class=K7&name=x' }), [{ who: 'self', url: target }], rel);
    assert.deepEqual(load(rel, { framed: true, search: '?q=abc&cb=1' }), [{ who: 'top', url: target }], rel);
  }
});

test('exam.html, the standalone audio exam, stays on this site', () => {
  assert.ok(!read('exam.html').includes('CUTOVER_REDIRECT'), 'exam.html must keep working here');
});

// The build contract (tools/build_version.js): one hash per page, and the same
// hash names that page's service-worker cache. An open page polls version.json
// every 2 minutes and reloads only on a new hash for ITS page - so each of the
// four pages reloads into its redirect once its own hash has changed.
test('version.json describes the pages as they are, and names the service-worker caches', () => {
  const version = JSON.parse(read('version.json'));
  const WORKERS = { 'examinee.html': 'sw-examinee.js', 'examiner.html': 'sw-examiner.js',
                    'teacher.html': 'sw-teacher.js', 'student.html': 'sw-student.js' };
  assert.deepEqual(Object.keys(version.pages).sort(), Object.keys(WORKERS).sort());
  const sha1 = s => crypto.createHash('sha1').update(Buffer.from(s, 'utf8')).digest('hex');
  for (const [page, worker] of Object.entries(WORKERS)) {
    // The build hashes the bytes on disk, and a checkout may hold either line ending.
    const raw = read(page), lf = raw.replace(/\r\n/g, '\n');
    assert.ok([sha1(raw), sha1(lf), sha1(lf.replace(/\n/g, '\r\n'))].includes(version.pages[page]),
      page + ' changed after the last `node tools/build_version.js`');
    const cache = page.replace('.html', '') + '-' + version.pages[page].slice(0, 8);
    assert.match(read(worker), new RegExp("^var CACHE(?:_NAME)? = '" + cache + "';\\r?$", 'm'), worker);
  }
});
