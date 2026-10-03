// Service Worker for Examinee PWA — NETWORK-FIRST.
// Always serves the freshest deployed examinee.html (cache is an offline-only
// fallback). Was cache-first, which left devices running a STALE examinee.html
// after every deploy until a manual cache bump — a real source of "old code on
// some devices". Mirrors sw-examiner.js (already network-first).
var CACHE = 'examinee-ad920083';
var IMG_CACHE = 'exam-images-v1';   // question images, warmed by the page at exam start, served offline
// 26/09/2026 cut-over: where examinee.html itself now sends every visitor.
var CUTOVER_REDIRECT = 'https://teoria-digital-vitaly.com/exam/';

self.addEventListener('install', function(e) {
  // The shared layers are part of the shell: without transport.js or bank.js the
  // page cannot run at all. The question TEXTS are not here and never will be —
  // they come from the session-gateway Worker against a per-exam signed grant,
  // and a grant is not something a cache may hand to the next device.
  e.waitUntil(caches.open(CACHE).then(function(c) {
    return c.addAll(['./examinee.html', './shared/transport.js', './shared/bank.js',
                     './icon-examinee-192.png', './icon-examinee-512.png']);
  }));
  self.skipWaiting();
});

self.addEventListener('activate', function(e) {
  e.waitUntil(caches.keys().then(function(ks) {
    return Promise.all(ks.filter(function(k) { return k !== CACHE && k !== IMG_CACHE; }).map(function(k) { return caches.delete(k); }));
  }));
  // 03/10/2026: the worker of 15/03-04/06/2026 was cache-first, so a device that
  // last opened examinee.html in those weeks is shown its OLD copy once more,
  // while this worker replaces that one behind it — and nothing moves that open
  // page (it has no update check, and the old system no longer answers it).
  // Once this worker controls the page, send it where examinee.html itself now
  // sends everyone. Only that page: every other page of this scope always came
  // from the network.
  e.waitUntil(Promise.resolve(self.clients.claim()).then(function() {
    return self.clients.matchAll({ type: 'window' });
  }).then(function(list) {
    return Promise.all(list.map(function(c) {
      var stale = false;
      try { stale = /\/examinee\.html$/.test(new URL(c.url).pathname); } catch (errUrl) { stale = false; }
      if (!stale || typeof c.navigate !== 'function') return null;
      return c.navigate(CUTOVER_REDIRECT).catch(function() { return null; });
    }));
  }).catch(function() {}));
});

self.addEventListener('fetch', function(e) {
  var req = e.request;
  if (req.method !== 'GET') return;                  // POST etc. (API writes) — untouched
  var url;
  try { url = new URL(req.url); } catch (_) { return; }
  if (url.origin === self.location.origin) {
    // NETWORK-FIRST same-origin: try the network so the latest HTML is always
    // served; refresh the cache on success; fall back to cache only when offline.
    e.respondWith(
      fetch(req).then(function(resp) {
        if (resp && resp.ok) { var clone = resp.clone(); caches.open(CACHE).then(function(c) { c.put(req, clone); }); }
        return resp;
      }).catch(function() {
        // Offline: serve the cached copy, and the page itself for anything else —
        // a reload mid-exam must land on examinee.html even with no network.
        return caches.match(req).then(function(r) { return r || caches.match('./examinee.html'); });
      })
    );
    return;
  }

  // CROSS-ORIGIN question images (gov.il / image proxy): network-first so online
  // stays fresh, but on failure serve the copy the exam page cached at start — so
  // an image question still renders if the connection drops mid-exam (risk #5).
  // ignoreSearch so buildResilientImage's cache-busting (?t=...) still matches.
  // All other cross-origin (API / TTS) is left untouched.
  if (req.destination === 'image') {
    e.respondWith(
      fetch(req).then(function(resp) {
        if (resp && (resp.ok || resp.type === 'opaque')) {
          var clone = resp.clone();
          caches.open(IMG_CACHE).then(function(c) { c.put(req, clone); });
        }
        return resp;
      }).catch(function() {
        return caches.open(IMG_CACHE).then(function(c) {
          return c.match(req, { ignoreSearch: true }).then(function(r) { return r || Response.error(); });
        });
      })
    );
  }
});
