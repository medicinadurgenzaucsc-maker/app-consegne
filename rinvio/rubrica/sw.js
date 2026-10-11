// Questo indirizzo ora rimanda soltanto all'applicazione delle consegne, dove
// vive la rubrica. Il service worker lasciato nei browser dalla rubrica di prima
// viene sostituito da questo, che toglie la SUA cache (ci stavano anche nomi e
// numeri) e si toglie da solo: da quel momento ogni richiesta va dritta alla
// rete, cioè alla pagina di rinvio. Le altre cache dello stesso dominio non si
// toccano: lì vive anche l'applicazione delle consegne.
self.addEventListener('install', function () {
  self.skipWaiting();
});

self.addEventListener('activate', function (e) {
  e.waitUntil(
    caches.keys().then(function (nomi) {
      return Promise.all(nomi.filter(function (n) { return n.indexOf('rubrica-') === 0; }).map(function (n) { return caches.delete(n); }));
    }).then(function () {
      return self.registration.unregister();
    })
  );
});
