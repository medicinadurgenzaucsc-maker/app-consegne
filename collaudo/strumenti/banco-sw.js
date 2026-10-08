// Banco del service worker: serve il sito come lo servirebbero Cloudflare Pages
// o GitHub Pages, con il service worker ACCESO (nel banco di prova normale è
// spento), e sa fingere i guasti che il service worker deve reggere.
//
//   node collaudo/strumenti/banco-sw.js [porta]        (predefinita 8766)
//
//   http://localhost:8766/prove/     pagina delle prove: await window.__proveSw()
//   http://localhost:8766/sito/      il sito (nessun utente collegato: basta la pagina di accesso)
//   /controllo/imposta?stile=cloudflare|github&modo=ok|404|500|giu|accesso&versione=lavoro|produzione
//   /controllo/richieste[?azzera=1]  le richieste arrivate al sito
//
// stile      «cloudflare»: x.html rinvia a x, senza 404.html ogni indirizzo
//            sconosciuto riceve la pagina principale; «github»: i file così come sono.
// modo       «404» e «500»: il sito risponde con un errore (i minuti in cui viene
//            ripubblicato); «giu»: la connessione cade; «accesso»: ogni richiesta
//            è rinviata alla pagina di accesso di un cancello, su un altro sito.
// versione   «lavoro»: i file della cartella di lavoro; «produzione»: quelli di
//            origin/master, per provare il passaggio dalla versione in reparto;
//            «compressa»: la copia compressa della cartella di lavoro, cioè la
//            forma pubblicata su Cloudflare (preparata al primo uso in _sito/banco-sw).
const http = require('http');
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const { execFileSync } = require('child_process');
const { elencoSito, prepara, RADICE } = require('./prepara-sito.js');

const PORTA = Number(process.argv[2] || 8766);
const RIF_PRODUZIONE = 'origin/master';
const MIME = {
  '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.css': 'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8', '.png': 'image/png', '.svg': 'image/svg+xml', '.ico': 'image/x-icon',
};
const stato = { stile: 'cloudflare', modo: 'ok', versione: 'lavoro' };
const AMMESSI = { stile: ['cloudflare', 'github'], modo: ['ok', '404', '500', 'giu', 'accesso'], versione: ['lavoro', 'produzione', 'compressa'] };
let richieste = [];

// I file delle due «pubblicazioni». Quelli di produzione si leggono da git una volta sola.
const memoria = {};
const COMPRESSA = path.join(RADICE, '_sito', 'banco-sw');
function elenco(versione) {
  if (versione === 'lavoro' || versione === 'compressa') return elencoSito();
  if (!memoria.elenco) memoria.elenco = elencoSito(RIF_PRODUZIONE);
  return memoria.elenco;
}
function leggi(versione, p) {
  if (elenco(versione).indexOf(p) < 0 || p.charAt(0) === '_') return null;
  if (versione === 'lavoro') return fs.readFileSync(path.join(RADICE, 'docs', p));
  if (versione === 'compressa') {
    if (!memoria.compressa) { prepara(COMPRESSA, { compressa: true }); memoria.compressa = true; }
    return fs.readFileSync(path.join(COMPRESSA, p));
  }
  const k = 'file:' + p;
  if (!memoria[k]) memoria[k] = execFileSync('git', ['-C', RADICE, 'show', RIF_PRODUZIONE + ':docs/' + p], { maxBuffer: 64 * 1024 * 1024 });
  return memoria[k];
}

function rispondi(res, statoHttp, intestazioni, corpo) {
  res.writeHead(statoHttp, intestazioni);
  res.end(corpo);
}
function testo(res, statoHttp, corpo, tipo) {
  rispondi(res, statoHttp, { 'Content-Type': tipo || 'text/plain; charset=utf-8', 'Cache-Control': 'no-store' }, corpo);
}
function file(req, res, statoHttp, p, corpo) {
  const etag = '"' + crypto.createHash('sha1').update(corpo).digest('hex') + '"';
  const h = {
    'Content-Type': MIME[path.extname(p).toLowerCase()] || 'application/octet-stream',
    'Cache-Control': stato.stile === 'cloudflare' ? 'public, max-age=0, must-revalidate' : 'max-age=600',
    ETag: etag,
  };
  if (statoHttp === 200 && req.headers['if-none-match'] === etag) return rispondi(res, 304, h);
  rispondi(res, statoHttp, h, corpo);
}

function sito(req, res, url) {
  const p = decodeURIComponent(url.pathname).slice('/sito/'.length);
  const nota = (esito) => richieste.push({ metodo: req.method, percorso: p + url.search, modo: stato.modo, esito: esito });
  if (stato.modo === 'giu') { nota('caduta'); return req.socket.destroy(); }
  if (stato.modo === '404') { nota(404); return testo(res, 404, '<h1>404</h1><p>ERRORE-FINTO: sito non trovato</p>', MIME['.html']); }
  if (stato.modo === '500') { nota(500); return testo(res, 500, '<h1>500</h1><p>ERRORE-FINTO: errore del server</p>', MIME['.html']); }
  if (stato.modo === 'accesso') {
    nota('rinvio al cancello');
    return rispondi(res, 302, { Location: 'http://127.0.0.1:' + PORTA + '/accesso/login?da=' + encodeURIComponent(url.pathname), 'Cache-Control': 'no-store' });
  }
  const v = stato.versione;
  if (stato.stile === 'cloudflare' && p.slice(-5) === '.html') {
    // come Cloudflare Pages: «x.html» rinvia a «x», «index.html» alla cartella
    let senza = p.slice(0, -5);
    if (senza === 'index') senza = '';
    else if (senza.slice(-6) === '/index') senza = senza.slice(0, -5);
    nota('rinvio 308');
    return rispondi(res, 308, { Location: '/sito/' + senza + url.search, 'Cache-Control': 'no-store' });
  }
  const candidati = (p === '' || p.slice(-1) === '/') ? [p + 'index.html'] : [p, p + '.html'];
  for (let i = 0; i < candidati.length; i++) {
    const corpo = leggi(v, candidati[i]);
    if (corpo) { nota(200); return file(req, res, 200, candidati[i], corpo); }
  }
  const nonTrovata = leggi(v, '404.html');
  if (nonTrovata) { nota(404); return file(req, res, 404, '404.html', nonTrovata); }
  // Cloudflare Pages senza 404.html: ogni indirizzo sconosciuto riceve la pagina principale
  if (stato.stile === 'cloudflare') { nota('200 (pagina principale al posto del file)'); return file(req, res, 200, 'index.html', leggi(v, 'index.html')); }
  nota(404);
  testo(res, 404, 'non trovato');
}

const PAGINA_PROVE = '<!DOCTYPE html><meta charset="utf-8"><title>Prove del service worker</title>'
  + '<body style="font:15px system-ui;margin:24px"><h1>Prove del service worker</h1>'
  + '<p>Dalla console: <code>await window.__proveSw()</code></p><div id="cornici"></div>'
  + '<script src="/prove/prove-sw.js"></script>';

http.createServer((req, res) => {
  try {
    const url = new URL(req.url || '/', 'http://localhost');
    if (url.pathname === '/controllo/imposta') {
      for (const k of Object.keys(AMMESSI)) {
        const val = url.searchParams.get(k);
        if (val == null) continue;
        if (AMMESSI[k].indexOf(val) < 0) return testo(res, 400, k + ': valori ammessi ' + AMMESSI[k].join(', '));
        stato[k] = val;
      }
      return testo(res, 200, JSON.stringify(stato), MIME['.json']);
    }
    if (url.pathname === '/controllo/richieste') {
      const copia = richieste;
      if (url.searchParams.has('azzera')) richieste = [];
      return testo(res, 200, JSON.stringify(copia), MIME['.json']);
    }
    if (url.pathname === '/accesso/login') {
      richieste.push({ metodo: req.method, percorso: '/accesso/login', modo: stato.modo, esito: 200 });
      return testo(res, 200, '<title>Accesso</title><h1>ACCESSO-FINTO</h1>', MIME['.html']);
    }
    if (url.pathname === '/prove/' || url.pathname === '/prove') return testo(res, 200, PAGINA_PROVE, MIME['.html']);
    if (url.pathname === '/prove/prove-sw.js') return testo(res, 200, fs.readFileSync(path.join(RADICE, 'collaudo', 'pagina', 'prove-sw.js')), MIME['.js']);
    if (url.pathname === '/sito') return rispondi(res, 301, { Location: '/sito/' });
    if (url.pathname.indexOf('/sito/') === 0) {
      if (req.method !== 'GET' && req.method !== 'HEAD') return testo(res, 405, 'metodo non ammesso');
      return sito(req, res, url);
    }
    testo(res, 404, 'non trovato');
  } catch (e) {
    testo(res, 500, 'errore del banco: ' + e.message);
  }
}).listen(PORTA, '127.0.0.1', () => console.log('banco del service worker su http://localhost:' + PORTA + '/prove/  (sito: /sito/)'));
