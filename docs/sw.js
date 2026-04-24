// Service Worker — Consegne Reparto
// Strategia: Network-first per le API GAS, Cache-first per gli asset statici

var CACHE_NAME = 'consegne-v1';

// Asset statici da pre-cachare all'installazione
var PRECACHE_ASSETS = [
  './',
  './index.html',
  './print.html',
  './favicon.png',
  './manifest.json',
  './css/styles.css',
  './js/api.js',
  './js/app.js',
  './js/app2.js'
];

// ── Installazione: pre-carica asset statici ──────────────────────
self.addEventListener('install', function(e) {
  e.waitUntil(
    caches.open(CACHE_NAME).then(function(cache) {
      return cache.addAll(PRECACHE_ASSETS);
    }).then(function() {
      return self.skipWaiting();
    })
  );
});

// ── Attivazione: rimuove cache vecchie ───────────────────────────
self.addEventListener('activate', function(e) {
  e.waitUntil(
    caches.keys().then(function(keys) {
      return Promise.all(
        keys.filter(function(k) { return k !== CACHE_NAME; })
            .map(function(k) { return caches.delete(k); })
      );
    }).then(function() {
      return self.clients.claim();
    })
  );
});

// ── Fetch: strategia ibrida ──────────────────────────────────────
self.addEventListener('fetch', function(e) {
  var url = e.request.url;

  // Chiamate al backend GAS → sempre Network (mai cache)
  if (url.indexOf('script.google.com') !== -1 ||
      url.indexOf('script.googleusercontent.com') !== -1) {
    e.respondWith(fetch(e.request));
    return;
  }

  // CDN esterni (Bootstrap, Bootstrap Icons, SweetAlert2) → Cache-first
  if (url.indexOf('cdn.jsdelivr.net') !== -1) {
    e.respondWith(
      caches.match(e.request).then(function(cached) {
        if (cached) return cached;
        return fetch(e.request).then(function(res) {
          var clone = res.clone();
          caches.open(CACHE_NAME).then(function(c) { c.put(e.request, clone); });
          return res;
        });
      })
    );
    return;
  }

  // Asset propri → Network-first con fallback alla cache
  e.respondWith(
    fetch(e.request).then(function(res) {
      // Aggiorna la cache con la versione fresca
      var clone = res.clone();
      caches.open(CACHE_NAME).then(function(c) { c.put(e.request, clone); });
      return res;
    }).catch(function() {
      // Rete non disponibile → usa la cache
      return caches.match(e.request);
    })
  );
});
