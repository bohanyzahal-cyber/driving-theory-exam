// Service Worker for Examinee PWA — NETWORK-FIRST.
// Always serves the freshest deployed examinee.html (cache is an offline-only
// fallback). Was cache-first, which left devices running a STALE examinee.html
// after every deploy until a manual cache bump — a real source of "old code on
// some devices". Mirrors sw-examiner.js (already network-first).
var CACHE = 'examinee-v3';

self.addEventListener('install', function(e) {
  e.waitUntil(caches.open(CACHE).then(function(c) {
    return c.addAll(['./examinee.html', './icon-examinee-192.png', './icon-examinee-512.png']);
  }));
  self.skipWaiting();
});

self.addEventListener('activate', function(e) {
  e.waitUntil(caches.keys().then(function(ks) {
    return Promise.all(ks.filter(function(k) { return k !== CACHE; }).map(function(k) { return caches.delete(k); }));
  }));
  self.clients.claim();
});

self.addEventListener('fetch', function(e) {
  var req = e.request;
  if (req.method !== 'GET') return;                  // POST etc. (API writes) — untouched
  var url;
  try { url = new URL(req.url); } catch (_) { return; }
  if (url.origin !== self.location.origin) return;   // cross-origin (API / TTS / gov images) — untouched
  // NETWORK-FIRST: try the network so the latest HTML is always served; refresh
  // the cache on success; fall back to cache only when offline.
  e.respondWith(
    fetch(req).then(function(resp) {
      if (resp && resp.ok) { var clone = resp.clone(); caches.open(CACHE).then(function(c) { c.put(req, clone); }); }
      return resp;
    }).catch(function() {
      return caches.match(req).then(function(r) { return r || caches.match('./examinee.html'); });
    })
  );
});
