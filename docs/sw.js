// Service Worker — Consegne Reparto
// Strategia: Network-first per le API GAS, Cache-first per gli asset statici

// Bump CACHE_NAME quando si aggiungono/cambiano asset nella precache,
// così i client già connessi invalidano la vecchia cache e ricaricano.
// v83: fix privacy — la cache non deve MAI contenere risposte Supabase/Google
// (dati pazienti a riposo su disco). Il bump cancella anche le cache
// precedenti che li contenevano (handler 'activate').
var CACHE_NAME = 'consegne-v87';

// Asset statici da pre-cachare all'installazione
var PRECACHE_ASSETS = [
  './',
  './index.html',
  './print.html',
  './favicon.png',
  './favicon-16.png',
  './favicon-32.png',
  './favicon-48.png',
  './apple-touch-icon.png',
  './icon-192.png',
  './icon-512.png',
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

// ── Fetch: strategia ibrida — ALLOWLIST DI CACHE (privacy) ──────────
// PRINCIPIO: su disco (Cache Storage) finiscono SOLO gli asset statici
// dell'app (same-origin) e i CDN. MAI le risposte con dati dei pazienti.
//
// Bug privacy corretto 19/07: il vecchio ramo "network-first" catturava
// QUALSIASI richiesta non-jsdelivr/non-GAS — incluse le GET a
// supabase.co (elenco pazienti, ~190KB) e a googleapis.com (backup Drive,
// userinfo). Quelle risposte venivano scritte in Cache Storage: dati
// sanitari a riposo sul disco di OGNI PC (anche condivisi), leggibili da
// DevTools → Application → Cache anche dopo logout/scadenza sessione.
// Ora Supabase/Google/OAuth passano dritti alla rete, senza toccare il
// disco. (Il bump di CACHE_NAME cancella pure le cache vecchie che li
// contenevano, via l'handler 'activate'.)
self.addEventListener('fetch', function(e) {
  var url = e.request.url;
  var sameOrigin = url.indexOf(self.location.origin) === 0;

  // CDN statici (Bootstrap, Bootstrap Icons, SweetAlert2, supabase-js) →
  // Cache-first: sono codice pubblico, nessun dato paziente.
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

  // Asset PROPRI dell'app (same-origin: index.html, js, css, icone) →
  // Network-first con fallback alla cache. Nessun dato paziente qui:
  // github.io serve solo file statici.
  if (sameOrigin) {
    e.respondWith(
      fetch(e.request).then(function(res) {
        var clone = res.clone();
        caches.open(CACHE_NAME).then(function(c) { c.put(e.request, clone); });
        return res;
      }).catch(function() {
        return caches.match(e.request);
      })
    );
    return;
  }

  // TUTTO IL RESTO (supabase.co = dati pazienti, googleapis.com = Drive,
  // accounts.google.com = login, ecc.) → SOLO rete, MAI in cache.
  e.respondWith(fetch(e.request));
});
