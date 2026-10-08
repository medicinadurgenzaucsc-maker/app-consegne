// Prove del service worker. Girano nella pagina /prove/ del banco del service
// worker (collaudo/strumenti/banco-sw.js), che sta FUORI dall'ambito del service
// worker: il sito è caricato dentro cornici, e questa pagina comanda al banco i
// guasti da fingere (sito che risponde con un errore, rete assente, cancello di
// accesso davanti al sito) e il cambio di versione.
//
//   await window.__proveSw()                             tutto, nei due stili
//   await window.__proveSw({ stili: ['cloudflare'] })    un solo stile
//   await window.__proveSw({ versione: 'compressa' })    le stesse prove sulla copia compressa del
//                                                        sito, la forma pubblicata su Cloudflare
//   await window.__proveSw({ versione: 'produzione' })   CONTROPROVA: le stesse prove sul service
//                                                        worker oggi in produzione, che ai guasti
//                                                        non regge (le prove devono saperlo dire)
//   window.__avanzamentoSw()                             a che punto è
//
// Esito (anche in window.__esitoSw): { prove, superate, fallite, elenco }.
(function () {
  var BASE = location.origin + '/sito/';
  var esiti = [], gruppo = '', inCorso = '';
  var attendi = function (ms) { return new Promise(function (r) { setTimeout(r, ms); }); };
  var unico = function () { return Date.now() + '' + Math.floor(Math.random() * 1000); };

  async function finche(cond, ms, passo) {
    var t0 = Date.now();
    while (Date.now() - t0 < ms) {
      var v = false;
      try { v = await cond(); } catch (e) {}
      if (v) return v;
      await attendi(passo || 200);
    }
    return false;
  }
  // una prova è superata se la funzione restituisce true (o nulla); una stringa è il motivo del fallimento
  async function prova(nome, fn) {
    var t0 = Date.now(), esito;
    inCorso = gruppo + ' › ' + nome;
    try { esito = await fn(); } catch (e) { esito = 'eccezione: ' + ((e && e.message) || e); }
    var ok = esito === true || esito === undefined;
    esiti.push({ gruppo: gruppo, nome: nome, ok: ok, dettaglio: ok ? '' : String(esito).slice(0, 500), ms: Date.now() - t0 });
    return ok;
  }
  async function imposta(o) {
    var q = Object.keys(o).map(function (k) { return k + '=' + encodeURIComponent(o[k]); }).join('&');
    var r = await fetch('/controllo/imposta?' + q, { cache: 'no-store' });
    if (!r.ok) throw new Error('banco: ' + (await r.text()));
    return r.json();
  }
  async function richieste(azzera) {
    return (await fetch('/controllo/richieste' + (azzera ? '?azzera=1' : ''), { cache: 'no-store' })).json();
  }

  // Carica un indirizzo in una cornice. «doc» resta vuoto se la pagina arrivata
  // non è del sito (pagina d'errore del browser, oppure un altro sito).
  function cornice(url, senzaPretese) {
    return new Promise(function (ok, ko) {
      var f = document.createElement('iframe');
      f.style.cssText = 'width:420px;height:240px;border:1px solid #bbb;margin:4px';
      var fatto = false;
      var t = setTimeout(function () {
        if (fatto) return;
        fatto = true;
        if (senzaPretese) ok({ f: f, win: null, doc: null }); else ko(new Error('la pagina ' + url + ' non si è caricata in 40 s'));
      }, senzaPretese ? 5000 : 40000);
      f.addEventListener('load', function () {
        if (fatto) return;
        fatto = true; clearTimeout(t);
        var win = null, doc = null;
        try { win = f.contentWindow; doc = win.document; if (!doc.body) doc = null; } catch (e) { win = null; doc = null; }
        ok({ f: f, win: win, doc: doc });
      });
      f.src = url;
      document.getElementById('cornici').appendChild(f);
    });
  }
  function togliCornici() { document.getElementById('cornici').innerHTML = ''; }

  // Via service worker e cache: ogni batteria parte da un browser «nuovo».
  async function azzera() {
    togliCornici();
    await attendi(400);
    var regs = await navigator.serviceWorker.getRegistrations();
    for (var i = 0; i < regs.length; i++) await regs[i].unregister();
    var nomi = await caches.keys();
    for (var j = 0; j < nomi.length; j++) await caches.delete(nomi[j]);
    await attendi(400);
  }
  async function attivo(c) {
    if (!c.win) return 'la pagina del sito non si è caricata';
    var reg = await Promise.race([c.win.navigator.serviceWorker.ready, attendi(60000).then(function () { return null; })]);
    if (!reg) return 'service worker non pronto dopo 60 s';
    var ok = await finche(function () { return !!c.win.navigator.serviceWorker.controller; }, 20000);
    return ok ? true : 'la pagina non è passata sotto il controllo del service worker';
  }
  // Nome della cache ed elenco dei file che il service worker pubblicato tiene in cache.
  async function elencoPrecache() {
    var t = await (await fetch(BASE + 'sw.js?' + unico(), { cache: 'no-store' })).text();
    // nella copia compressa la stessa riga è scritta senza spazi e con le virgolette doppie
    var nome = (/CACHE_NAME\s*=\s*['"]([^'"]+)['"]/.exec(t) || [])[1];
    var corpo = (/PRECACHE_ASSETS\s*=\s*\[([^\]]*)\]/.exec(t) || [])[1] || '';
    return { nome: nome, percorsi: (corpo.match(/['"][^'"]+['"]/g) || []).map(function (s) { return s.slice(1, -1); }) };
  }

  // La cache del service worker nuovo: una sola, con tutti i file dell'elenco, e
  // dentro niente che non sia un file buono del sito (o una libreria).
  async function cacheSana(atteso) {
    var nomi = await caches.keys();
    if (nomi.length !== 1 || nomi[0] !== atteso.nome) return 'cache presenti: ' + (nomi.join(', ') || 'nessuna') + ' (attesa solo ' + atteso.nome + ')';
    var cache = await caches.open(nomi[0]);
    for (var i = 0; i < atteso.percorsi.length; i++) {
      var p = atteso.percorsi[i];
      var r = await cache.match(new URL(p, BASE).href);
      if (!r) return 'in cache manca ' + p;
      if (!r.ok) return p + ': in cache con stato ' + r.status;
      if (r.redirected) return p + ': in cache con il segno del rinvio';
    }
    var chiavi = await cache.keys();
    for (var j = 0; j < chiavi.length; j++) {
      var u = chiavi[j].url;
      if (u.indexOf(location.origin + '/') !== 0) {
        if (u.indexOf('https://cdn.jsdelivr.net/') !== 0) return 'in cache una voce di un altro sito: ' + u;
        continue;
      }
      if (u.indexOf('?') >= 0) return 'in cache una voce con parametri: ' + u;
      var voce = await cache.match(chiavi[j]);
      if (!voce.ok) return 'in cache una risposta non buona: ' + u + ' (' + voce.status + ')';
      if (voce.redirected) return 'in cache una risposta col segno del rinvio: ' + u;
      if (/\.(html|js|css|json)$/.test(u) || u.slice(-1) === '/') {
        var testo = await voce.text();
        if (testo.indexOf('ERRORE-FINTO') >= 0 || testo.indexOf('ACCESSO-FINTO') >= 0) return 'in cache è finita una pagina di errore o di accesso: ' + u;
      }
    }
    return true;
  }
  // Un file del sito chiesto dalla pagina controllata: deve arrivare buono. Il
  // parametro unico obbliga il service worker a provare la rete ogni volta.
  async function fileSano(c, percorso, segno) {
    var r = await c.win.fetch(percorso + (percorso.indexOf('?') < 0 ? '?' : '&') + 'prova=' + unico());
    var t = await r.text();
    var ok = r.status === 200 && t.indexOf(segno) >= 0 && t.indexOf('ERRORE-FINTO') < 0 && t.indexOf('ACCESSO-FINTO') < 0;
    return ok || (percorso + ': stato ' + r.status + ', inizio «' + t.slice(0, 70) + '»');
  }
  // Una navigazione verso la pagina principale: deve arrivare la pagina dell'app.
  async function paginaSana(url) {
    var c = await cornice(url + (url.indexOf('?') < 0 ? '?' : '&') + '_r=' + unico());
    var esito = true;
    if (!c.doc) esito = url + ': al posto della pagina un errore del browser';
    else if (!c.doc.getElementById('navVersioneApp')) esito = url + ': titolo «' + c.doc.title + '», testo «' + c.doc.body.innerText.slice(0, 80) + '»';
    c.f.remove();
    return esito;
  }
  function immagine(c, percorso) {
    return new Promise(function (ok) {
      var img = new c.win.Image();
      img.onload = function () { ok(true); };
      img.onerror = function () { ok(false); };
      img.src = percorso + '?prova=' + unico();
    });
  }

  // ── il service worker della versione di lavoro davanti ai guasti ─────────
  async function batteria(stile, versione) {
    gruppo = stile + (versione === 'produzione' ? ' · versione in produzione (controprova)' : versione === 'compressa' ? ' · copia compressa' : ' · versione di lavoro');
    await azzera();
    await imposta({ stile: stile, modo: 'ok', versione: versione });
    var atteso = await elencoPrecache();
    var c = await cornice(BASE);
    if (!(await prova('il service worker si installa e prende il controllo della pagina', function () { return attivo(c); }))) return;
    await prova('in cache tutti i file dell\'elenco, con risposte buone e senza segno di rinvio', function () { return cacheSana(atteso); });

    var guasti = [['404', 'il sito risponde 404'], ['500', 'il sito risponde 500'], ['giu', 'la rete non risponde']];
    for (var i = 0; i < guasti.length; i++) {
      await imposta({ modo: guasti[i][0] });
      var G = 'quando ' + guasti[i][1] + ': ';
      await prova(G + 'i file dell\'app arrivano dalla cache', async function () {
        var a = await fileSano(c, 'js/api.js', '_AMBIENTI'); if (a !== true) return a;
        a = await fileSano(c, 'css/styles.css', '{'); if (a !== true) return a;
        return fileSano(c, 'print.html?layout=alt&saltaVuoti=1', 'Stampa Consegne');
      });
      await prova(G + 'la pagina principale si apre dalla cache', async function () {
        var a = await paginaSana(BASE); if (a !== true) return a;
        return paginaSana(BASE + 'index.html');
      });
      await prova(G + 'in cache non finisce nulla di sbagliato', function () { return cacheSana(atteso); });
    }
    await prova('dopo i guasti il service worker è ancora registrato', async function () {
      var regs = await navigator.serviceWorker.getRegistrations();
      return (regs.length === 1 && !!regs[0].active) || ('registrazioni: ' + regs.length);
    });

    await imposta({ modo: 'accesso' });
    var A = 'con un cancello di accesso che rinvia a un altro sito: ';
    await prova(A + 'i file dell\'app arrivano dalla cache', async function () {
      var a = await fileSano(c, 'js/app.js', 'function'); if (a !== true) return a;
      return (await immagine(c, 'favicon.png')) || 'un\'immagine del sito non si carica';
    });
    await prova(A + 'una navigazione viene lasciata andare alla pagina di accesso', async function () {
      await richieste(true);
      var x = await cornice(BASE + '?cancello=' + unico(), true);
      await attendi(1500);
      var e = await richieste();
      x.f.remove();
      return e.some(function (r) { return r.percorso === '/accesso/login'; }) || ('la pagina di accesso non è stata chiesta: ' + JSON.stringify(e.slice(0, 4)));
    });
    await prova(A + 'la pagina di accesso non finisce in cache', function () { return cacheSana(atteso); });

    await imposta({ modo: 'ok' });
    await prova('tornato il sito, i file si chiedono di nuovo alla rete', async function () {
      await richieste(true);
      var a = await fileSano(c, 'js/app2.js', 'function'); if (a !== true) return a;
      var e = await richieste();
      return e.some(function (r) { return r.percorso.indexOf('js/app2.js') === 0 && r.esito === 200; }) || 'al sito non è arrivata nessuna richiesta';
    });
    await prova('una pagina che non esiste riceve un 404 vero, non la pagina principale', async function () {
      var r = await c.win.fetch('non-esiste-' + unico() + '.js');
      var t = await r.text();
      return (r.status === 404 && t.indexOf('navVersioneApp') < 0) || ('stato ' + r.status + ', inizio «' + t.slice(0, 60) + '»');
    });
    await prova('alla fine la cache è ancora sana', function () { return cacheSana(atteso); });
    togliCornici();
  }

  // ── il passaggio dalla versione oggi in produzione a quella di lavoro ────
  async function passaggio(stile, arrivo) {
    gruppo = stile + ' · passaggio dalla versione in produzione';
    await azzera();
    await imposta({ stile: stile, modo: 'ok', versione: 'produzione' });
    var vecchio = await elencoPrecache();
    var c = await cornice(BASE);
    if (!(await prova('la versione in produzione (' + vecchio.nome + ') si installa', function () { return attivo(c); }))) return;
    await prova('la sua cache è quella attesa', async function () {
      var n = await caches.keys();
      return (n.length === 1 && n[0] === vecchio.nome) || ('cache: ' + n.join(', '));
    });
    await imposta({ versione: arrivo || 'lavoro' });
    var nuovo = await elencoPrecache();
    await prova('esce la versione nuova (' + nuovo.nome + '): il service worker si aggiorna, resta una sola cache ed è completa', async function () {
      if (nuovo.nome === vecchio.nome) return 'le due versioni hanno lo stesso nome di cache: ' + nuovo.nome;
      var reg = await c.win.navigator.serviceWorker.getRegistration();
      await reg.update();
      var ok = await finche(async function () { var n = await caches.keys(); return n.length === 1 && n[0] === nuovo.nome; }, 60000, 300);
      if (!ok) return 'cache dopo 60 s: ' + (await caches.keys()).join(', ');
      await attendi(800);
      return cacheSana(nuovo);
    });
    await prova('la pagina ricaricata è quella nuova', async function () {
      var x = await cornice(BASE + '?_r=' + unico());
      var el = x.doc && x.doc.getElementById('navVersioneApp');
      var v = el ? el.textContent.trim() : '(assente)';
      x.f.remove();
      return ('consegne-v' + v === nuovo.nome) || ('versione mostrata: ' + v + ', attesa ' + nuovo.nome);
    });
    await prova('subito dopo il passaggio regge un sito che risponde 404', async function () {
      await imposta({ modo: '404' });
      var a = await paginaSana(BASE);
      await imposta({ modo: 'ok' });
      return a;
    });
    togliCornici();
  }

  window.__avanzamentoSw = function () {
    var fallite = esiti.filter(function (e) { return !e.ok; }).length;
    return { fatte: esiti.length, fallite: fallite, inCorso: inCorso, finito: !!window.__esitoSw };
  };
  window.__proveSw = async function (opz) {
    opz = opz || {};
    esiti = []; window.__esitoSw = null;
    var stili = opz.stili || ['cloudflare', 'github'];
    try {
      for (var i = 0; i < stili.length; i++) {
        var daProvare = opz.versione === 'produzione' ? 'produzione' : opz.versione === 'compressa' ? 'compressa' : 'lavoro';
        await batteria(stili[i], daProvare);
        if (!opz.senzaPassaggio && opz.versione !== 'produzione') await passaggio(stili[i], daProvare);
      }
    } catch (e) {
      esiti.push({ gruppo: gruppo, nome: 'le prove si sono interrotte', ok: false, dettaglio: String((e && e.message) || e).slice(0, 500), ms: 0 });
    }
    try { await azzera(); await imposta({ stile: 'cloudflare', modo: 'ok', versione: 'lavoro' }); } catch (e) {}
    var fallite = esiti.filter(function (e) { return !e.ok; });
    window.__esitoSw = { prove: esiti.length, superate: esiti.length - fallite.length, fallite: fallite.length, elenco: esiti };
    return window.__esitoSw;
  };
})();
