// Service Worker for Teacher PWA
//
// CACHE_NAME below is REWRITTEN BY THE BUILD (tools/build_version.js) from the
// sha1 of teacher.html — never edit it by hand and never move it off its own
// line: one hash names both the update check (version.json) and this cache, so
// a page that really changed also invalidates its offline copy, and a push that
// did not change it invalidates nothing. (D8: the old hand-bumped vNN drifted
// eight deploys behind.)
var CACHE_NAME = 'teacher-ad17a04b';

// Install — cache the page shell plus the shared client modules. transport.js
// and bank.js are separate files now, so an offline shell without them is a
// blank page.
self.addEventListener('install', function(e) {
  e.waitUntil(
    caches.open(CACHE_NAME).then(function(cache) {
      return cache.addAll([
        './teacher.html',
        './shared/transport.js',
        './shared/bank.js',
        './icon-teacher-192.png',
        './icon-teacher-512.png',
        './manifest-teacher.json'
      ]);
    })
  );
  self.skipWaiting();
});

// Activate — clean old caches
self.addEventListener('activate', function(e) {
  e.waitUntil(
    caches.keys().then(function(names) {
      return Promise.all(
        names.filter(function(n) { return n !== CACHE_NAME; })
             .map(function(n) { return caches.delete(n); })
      );
    })
  );
  self.clients.claim();
});

// Fetch — network first, cache only as an offline fallback.
//
// D7: GET ONLY. The previous version filtered by URL alone, so every POST to
// the API and every HEAD from the old update check reached cache.put() and threw
// "Request method 'HEAD' is unsupported" in the live console — and caches.match()
// on a non-GET request can never hit anyway. Same guard as sw-examinee.js.
self.addEventListener('fetch', function(e) {
  var req = e.request;
  if (req.method !== 'GET') return;
  var url = req.url;
  if (url.indexOf('script.google.com') !== -1) return;   // API — always network

  e.respondWith(
    fetch(req).then(function(response) {
      if (response && response.ok) {
        var clone = response.clone();
        caches.open(CACHE_NAME).then(function(cache) { cache.put(req, clone); });
      }
      return response;
    }).catch(function() {
      // bank/<lang>.json is fetched with a ?v=<sha> cache-buster; ignoreSearch so
      // the copy taken at load time still answers when the network is gone.
      return caches.match(req, { ignoreSearch: true });
    })
  );
});
