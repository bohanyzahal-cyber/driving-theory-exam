// Service Worker for Examinee PWA — NETWORK-FIRST.
// Always serves the freshest deployed examinee.html (cache is an offline-only
// fallback). Was cache-first, which left devices running a STALE examinee.html
// after every deploy until a manual cache bump — a real source of "old code on
// some devices". Mirrors sw-examiner.js (already network-first).
var CACHE = 'examinee-b438919a';
var IMG_CACHE = 'exam-images-v1';   // question images, warmed by the page at exam start, served offline

self.addEventListener('install', function(e) {
  // The shared layers and the bank manifest are part of the shell now: without
  // transport.js or bank.js the page cannot run at all, and without the manifest
  // it cannot even name the bank files it needs.
  e.waitUntil(caches.open(CACHE).then(function(c) {
    return c.addAll(['./examinee.html', './shared/transport.js', './shared/bank.js', './bank/manifest.json',
                     './icon-examinee-192.png', './icon-examinee-512.png']);
  }));
  self.skipWaiting();
});

self.addEventListener('activate', function(e) {
  e.waitUntil(caches.keys().then(function(ks) {
    return Promise.all(ks.filter(function(k) { return k !== CACHE && k !== IMG_CACHE; }).map(function(k) { return caches.delete(k); }));
  }));
  self.clients.claim();
});

self.addEventListener('fetch', function(e) {
  var req = e.request;
  if (req.method !== 'GET') return;                  // POST etc. (API writes) — untouched
  var url;
  try { url = new URL(req.url); } catch (_) { return; }
  if (url.origin === self.location.origin) {
    // NETWORK-FIRST same-origin: try the network so the latest HTML is always
    // served; refresh the cache on success; fall back to cache only when offline.
    var isBank = url.pathname.indexOf('/bank/') !== -1 && /\.json$/.test(url.pathname);
    e.respondWith(
      fetch(req).then(function(resp) {
        if (resp && resp.ok) { var clone = resp.clone(); caches.open(CACHE).then(function(c) { c.put(req, clone); }); }
        return resp;
      }).catch(function() {
        // bank/<lang>.json is requested with ?v=<sha>; offline, the cached copy
        // of the previous build is the only text the examinee has, and a query
        // string must not be what stands between them and their exam.
        return caches.match(req, isBank ? { ignoreSearch: true } : undefined)
          .then(function(r) { return r || (isBank ? Response.error() : caches.match('./examinee.html')); });
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
