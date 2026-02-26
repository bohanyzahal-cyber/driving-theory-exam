// Service Worker for Examiner PWA
var CACHE_NAME = 'examiner-v2';

// Install — cache the examiner page shell
self.addEventListener('install', function(e) {
  e.waitUntil(
    caches.open(CACHE_NAME).then(function(cache) {
      return cache.addAll([
        './examiner.html',
        './logo.png'
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

// Fetch — network first, fallback to cache (API calls always go to network)
self.addEventListener('fetch', function(e) {
  var url = e.request.url;
  // Always go to network for API calls
  if (url.indexOf('script.google.com') !== -1) return;
  // Always go to network for QR fallback API
  if (url.indexOf('qrserver.com') !== -1) return;

  e.respondWith(
    fetch(e.request).then(function(response) {
      // Update cache with fresh version
      if (response.ok) {
        var clone = response.clone();
        caches.open(CACHE_NAME).then(function(cache) { cache.put(e.request, clone); });
      }
      return response;
    }).catch(function() {
      return caches.match(e.request);
    })
  );
});
