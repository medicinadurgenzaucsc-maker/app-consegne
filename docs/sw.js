// Service Worker — Consegne Reparto
// Strategia: per i file dell'app prima la rete e, se la rete non dà una
// risposta buona, la copia in cache; per le librerie esterne prima la cache;
// tutto il resto (i dati dei pazienti) solo rete, mai su disco.

// Bump CACHE_NAME quando si aggiungono/cambiano asset nella precache,
// così i client già connessi invalidano la vecchia cache e ricaricano.
// v83: fix privacy — la cache non deve MAI contenere risposte Supabase/Google
// (dati pazienti a riposo su disco). Il bump cancella anche le cache
// precedenti che li contenevano (handler 'activate').
var CACHE_NAME = 'consegne-v188';

// Asset statici da pre-cachare all'installazione.
// La rubrica («Numeri Telefono») è una pagina dell'app: il riquadro la chiede
// come cartella («rubrica/»), quindi in cache servono sia la cartella sia il
// file. La rubrica non ha, e non deve avere, un service worker suo.
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
  './js/librerie/purify-3.4.16.min.js',
  './js/sanifica.js',
  './js/api.js',
  './js/app.js',
  './js/app2.js',
  './rubrica/',
  './rubrica/index.html',
  './rubrica/app.css',
  './rubrica/app.js'
];

// Una risposta arrivata dopo un rinvio (su Cloudflare «print.html» viene
// servita come «/print») non può rispondere a una navigazione: il browser la
// rifiuta. In cache se ne tiene una copia identica, senza il segno del rinvio.
function senzaRinvio(res) {
  if (!res.redirected) return Promise.resolve(res);
  return res.blob().then(function(corpo) {
    return new Response(corpo, { status: res.status, statusText: res.statusText, headers: res.headers });
  });
}

// I file dell'app non cambiano con i parametri dell'indirizzo: in cache hanno
// una sola voce ciascuno («/?_r=123» e «/» sono la stessa pagina).
function senzaParametri(url) {
  return url.split('#')[0].split('?')[0];
}

function dallaCache(request) {
  return caches.match(request, { ignoreSearch: true, ignoreVary: true });
}

// ── Installazione: pre-carica asset statici ─────────────────────
// Ogni file è chiesto di nuovo al sito (mai dalla memoria del browser) e deve
// arrivare con una risposta buona, altrimenti l'installazione fallisce e resta
// al lavoro il service worker di prima.
self.addEventListener('install', function(e) {
  e.waitUntil(
    caches.open(CACHE_NAME).then(function(cache) {
      return Promise.all(PRECACHE_ASSETS.map(function(percorso) {
        return fetch(new Request(percorso, { cache: 'no-cache' })).then(function(res) {
          if (!res.ok) throw new Error('file non disponibile: ' + percorso + ' (HTTP ' + res.status + ')');
          return senzaRinvio(res).then(function(pulita) { return cache.put(percorso, pulita); });
        });
      }));
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
  var sameOrigin = false;
  try { sameOrigin = new URL(url).origin === self.location.origin; } catch (err) {}

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
  // prima la rete, e in cache finisce SOLO una risposta buona. Nessun dato
  // paziente qui: il sito serve solo file statici.
  // Se la rete non risponde, o risponde con un errore (il «404» dei minuti in
  // cui un sito viene ripubblicato), o al posto del file arriva la pagina di
  // accesso di un cancello messo davanti al sito (risposta «opaca»), si usa la
  // copia buona già in cache: fino alla v184 l'errore veniva mostrato e
  // SALVATO al posto della copia buona.
  if (sameOrigin && e.request.method === 'GET') {
    e.respondWith(
      fetch(e.request).then(function(res) {
        // rinvio di una navigazione: lo segue il browser, non si salva
        if (res.type === 'opaqueredirect') return res;
        if (res.ok) {
          var clone = res.clone();
          e.waitUntil(
            senzaRinvio(clone).then(function(pulita) {
              return caches.open(CACHE_NAME).then(function(c) { return c.put(senzaParametri(url), pulita); });
            }).catch(function() {})
          );
          return res;
        }
        return dallaCache(e.request).then(function(buona) { return buona || res; });
      }).catch(function() {
        return dallaCache(e.request).then(function(buona) { return buona || Response.error(); });
      })
    );
    return;
  }

  // TUTTO IL RESTO (supabase.co = dati pazienti, googleapis.com,
  // accounts.google.com = login, ecc.) → SOLO rete, MAI in cache.
  e.respondWith(fetch(e.request));
});
