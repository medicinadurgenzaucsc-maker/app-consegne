// Questo indirizzo ora rimanda soltanto al nuovo sito dell'applicazione. Il
// service worker rimasto nei browser dalla versione precedente viene sostituito
// da questo, che svuota la cache dell'applicazione e si toglie da solo: da quel
// momento ogni richiesta va dritta alla rete, cioè alla pagina di rinvio.
// Non ricarica le pagine aperte: chi sta lavorando continua finché non ricarica.
self.addEventListener('install', function () {
  self.skipWaiting();
});

self.addEventListener('activate', function (e) {
  e.waitUntil(
    caches.keys().then(function (nomi) {
      // solo le cache di questa applicazione: sullo stesso dominio possono vivere altri siti
      return Promise.all(nomi.filter(function (n) { return n.indexOf('consegne-') === 0; }).map(function (n) { return caches.delete(n); }));
    }).then(function () {
      return self.registration.unregister();
    })
  );
});
