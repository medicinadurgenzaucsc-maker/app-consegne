// Banco di prova locale: serve l'app così com'è nella cartella di lavoro,
// collegata al database di COLLAUDO (su localhost l'app sceglie il collaudo),
// con la sessione dell'utente fittizio delle prove già pronta. Nessun file
// viene copiato o modificato: l'avvio del banco è aggiunto al volo alle pagine.
//
//   node collaudo/strumenti/banco.js [porta]
//   node collaudo/strumenti/banco.js [porta] compressa   serve la copia COMPRESSA del sito,
//        quella che viene pubblicata su Cloudflare (preparata all'avvio in _sito/banco:
//        dopo una modifica a docs/ il banco va riavviato)
//
//   http://localhost:8765/                  l'app (cartella docs/)
//   http://localhost:8765/?sloggato         la stessa, senza sessione (prova di fumo)
//   http://localhost:8765/?senzaFiltro      la stessa, senza la libreria del filtro HTML
//   http://localhost:8765/rubrica/          la rubrica aperta DA SOLA: deve restare sul messaggio fisso
//   http://localhost:8765/trak-finto/       il finto TrakCare
//   http://localhost:8765/pagina/…          generatore e prove che girano dentro l'app
//   /banco/cassaforte, /banco/posta         servono alle prove della mail (vedi sotto)
const http = require('http');
const fs = require('fs');
const path = require('path');
const { genera } = require('./gettone-prova.js');

const RADICE = path.resolve(__dirname, '../..');
const PORTA = Number(process.argv[2] || 8765);
const COMPRESSA = process.argv[3] === 'compressa';
if (process.argv[3] && !COMPRESSA) { console.log('secondo argomento non riconosciuto: ' + process.argv[3]); process.exit(1); }
// La copia compressa si prepara una volta, all'avvio: se non riesce il banco non parte.
const SITO = COMPRESSA ? path.join(RADICE, '_sito', 'banco') : path.join(RADICE, 'docs');
if (COMPRESSA) {
  const r = require('./prepara-sito.js').prepara(SITO, { compressa: true });
  console.log('copia compressa pronta: ' + r.elenco.length + ' file, ' + Math.round(r.byte / 1024) + ' KB (' + r.compressione.strumento + ')');
}
const MONTAGGI = [
  ['/trak-finto/', path.join(RADICE, 'collaudo', 'trak-finto')],
  ['/pagina/', path.join(RADICE, 'collaudo', 'pagina')],
  ['/', SITO],
];
const MIME = {
  '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.css': 'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8', '.png': 'image/png', '.svg': 'image/svg+xml', '.ico': 'image/x-icon',
  '.md': 'text/plain; charset=utf-8',
};
const SDK = /<script src="https:\/\/cdn\.jsdelivr\.net\/npm\/@supabase\/supabase-js@2\/dist\/umd\/supabase\.min\.js"><\/script>/;
// Nella pagina della rubrica il raccoglitore degli errori va in testa, prima del foglio di stile.
const TESTA_RUBRICA = /<link rel="stylesheet" href="app\.css">/;

// Sessione di prova: firmata al bisogno e rinnovata quando sta per scadere.
let sessione = null;
async function sessioneValida() {
  if (!sessione || sessione.scade * 1000 - Date.now() < 30 * 60 * 1000) sessione = await genera(10);
  return sessione;
}

function avvio(chiave, valore) {
  return '// BANCO LOCALE: generato al volo, non esiste su disco.\n(function () {\n'
    + (valore
      ? '  try { localStorage.setItem(' + JSON.stringify(chiave) + ', ' + JSON.stringify(valore) + '); } catch (e) {}\n'
      : '  try { localStorage.removeItem(' + JSON.stringify(chiave) + '); } catch (e) {}\n')
    + '  try { navigator.serviceWorker.register = function () { return Promise.reject(new Error("service worker spento nel banco")); }; } catch (e) {}\n'
    // le prove aprono la pagina di stampa dentro una cornice: lì la finestra di stampa non deve aprirsi
    + '  if (/[?&]senzaStampa=/.test(location.search)) window.print = function () { window.__stampaChiesta = (window.__stampaChiesta || 0) + 1; };\n'
    + '  window.__erroriBanco = [];\n'
    + '  window.addEventListener("error", function (e) { window.__erroriBanco.push("UNCAUGHT: " + e.message); });\n'
    + '  window.addEventListener("unhandledrejection", function (e) { window.__erroriBanco.push("PROMISE: " + String((e.reason && e.reason.message) || e.reason)); });\n'
    + '})();\n';
}

// La pagina della rubrica (docs/rubrica/) non ha una sessione sua: la chiede alla
// pagina che la contiene. Dal banco riceve SOLO la raccolta degli errori, mai la
// sessione (aperta da sola deve restare senza), e come file a parte: la sua
// regola CSP non ammette script scritti in pagina. Fra gli errori finisce anche
// ciò che quella regola ha bloccato (uno stile o un gestore in linea).
function raccoglitore() {
  return '// BANCO LOCALE: generato al volo, non esiste su disco. Solo raccolta degli errori.\n(function () {\n'
    + '  window.__erroriBanco = [];\n'
    + '  window.addEventListener("error", function (e) { window.__erroriBanco.push("UNCAUGHT: " + e.message); });\n'
    + '  window.addEventListener("unhandledrejection", function (e) { window.__erroriBanco.push("PROMISE: " + String((e.reason && e.reason.message) || e.reason)); });\n'
    + '  document.addEventListener("securitypolicyviolation", function (e) { window.__erroriBanco.push("CSP: " + e.violatedDirective + " " + (e.blockedURI || "") + " " + String(e.sample || "").slice(0, 60)); });\n'
    + '})();\n';
}

function rispondi(res, stato, tipo, corpo) {
  res.writeHead(stato, { 'Content-Type': tipo, 'Cache-Control': 'no-store' });
  res.end(corpo);
}

http.createServer(async (req, res) => {
  try {
    const url = new URL(req.url || '/', 'http://localhost');
    let percorso = decodeURIComponent(url.pathname);

    if (percorso === '/banco-avvio.js') {
      // Il token di prova esce solo verso le pagine del banco stesso.
      if (req.headers['sec-fetch-site'] !== 'same-origin') return rispondi(res, 403, 'text/plain; charset=utf-8', 'riservato alle pagine del banco');
      const { collaudo } = require('./sb.js');
      const chiave = 'sb-' + collaudo().ref + '-auth-token';
      if (url.searchParams.has('sloggato')) return rispondi(res, 200, MIME['.js'], avvio(chiave, null));
      try {
        const s = await sessioneValida();
        return rispondi(res, 200, MIME['.js'], avvio(chiave, JSON.stringify(s.sessione)));
      } catch (e) {
        console.log('sessione di prova non disponibile: ' + e.message);
        return rispondi(res, 200, MIME['.js'], avvio(chiave, null) + 'console.error("banco: sessione di prova non disponibile");\n');
      }
    }

    if (percorso === '/banco-errori.js') return rispondi(res, 200, MIME['.js'], raccoglitore());

    // Prove della mail: la cassaforte del collaudo passa alle credenziali finte
    // (e torna com'era alla fine), e il finto Google dice che cosa ha «spedito».
    // Solo dalle pagine del banco: di qui passa il segreto del finto Google.
    if (percorso === '/banco/cassaforte' || percorso === '/banco/posta') {
      if (req.headers['sec-fetch-site'] !== 'same-origin') return rispondi(res, 403, 'text/plain; charset=utf-8', 'riservato alle pagine del banco');
      const cassaforte = require('./cassaforte-collaudo.js');
      if (percorso === '/banco/posta') {
        const oggetto = String(url.searchParams.get('oggetto') || '').replace(/\$/g, '');
        const righe = await require('./sb.js').collaudo().query("select mittente, destinatari, oggetto, reale from public.posta_simulata where oggetto = $o$" + oggetto + "$o$ order by id desc limit 1");
        return rispondi(res, 200, MIME['.json'], JSON.stringify(righe[0] || null));
      }
      const modo = url.searchParams.get('modo');
      if (req.method !== 'POST' || ['finta', 'vera', 'senza-segreto', 'segreto-sbagliato'].indexOf(modo) < 0) return rispondi(res, 400, 'text/plain; charset=utf-8', 'uso: POST /banco/cassaforte?modo=finta|vera|senza-segreto|segreto-sbagliato');
      if (modo === 'finta') return rispondi(res, 200, MIME['.json'], JSON.stringify({ segreto: await cassaforte.finta(), clientFinto: cassaforte.CLIENT_FINTO }));
      // come al primo avvio: né client secret né consenso (solo su una cassaforte finta)
      if (modo === 'senza-segreto') return rispondi(res, 200, MIME['.json'], JSON.stringify(await cassaforte.senzaSegreto()));
      // il secret custodito non è più riconosciuto da «Google» (solo su una cassaforte finta)
      if (modo === 'segreto-sbagliato') return rispondi(res, 200, MIME['.json'], JSON.stringify(await cassaforte.segretoSbagliato()));
      return rispondi(res, 200, MIME['.json'], JSON.stringify(await cassaforte.vera()));
    }

    // Le prove salvano qui la «fotografia» degli esiti attesi: un solo file,
    // dal solo banco, e solo se è JSON valido.
    if (req.method === 'PUT' && percorso === '/pagina/attesi-terapia.json') {
      if (req.headers['sec-fetch-site'] !== 'same-origin') return rispondi(res, 403, 'text/plain; charset=utf-8', 'riservato alle pagine del banco');
      const pezzi = [];
      for await (const p of req) { pezzi.push(p); if (pezzi.reduce((s, x) => s + x.length, 0) > 2 * 1024 * 1024) return rispondi(res, 413, 'text/plain; charset=utf-8', 'troppo grande'); }
      const testo = Buffer.concat(pezzi).toString('utf8');
      JSON.parse(testo);
      fs.writeFileSync(path.join(RADICE, 'collaudo', 'pagina', 'attesi-terapia.json'), testo);
      return rispondi(res, 200, 'text/plain; charset=utf-8', 'salvato');
    }
    if (req.method !== 'GET' && req.method !== 'HEAD') return rispondi(res, 405, 'text/plain; charset=utf-8', 'metodo non ammesso');

    const montaggio = MONTAGGI.find((m) => percorso.startsWith(m[0]));
    if (percorso.endsWith('/')) percorso += 'index.html';
    const file = path.normalize(path.join(montaggio[1], percorso.slice(montaggio[0].length)));
    // Una cartella chiesta senza la barra finale («/rubrica»): rinvio alla cartella, come fanno GitHub e Cloudflare.
    if (file.startsWith(montaggio[1] + path.sep) && fs.existsSync(file) && fs.statSync(file).isDirectory() && fs.existsSync(path.join(file, 'index.html'))) {
      res.writeHead(301, { Location: url.pathname + '/' + url.search, 'Cache-Control': 'no-store' });
      return res.end();
    }
    if (!file.startsWith(montaggio[1] + path.sep) || !fs.existsSync(file) || !fs.statSync(file).isFile()) {
      return rispondi(res, 404, 'text/plain; charset=utf-8', 'non trovato');
    }
    const est = path.extname(file).toLowerCase();
    // Solo le pagine dell'app ricevono l'avvio del banco, subito prima dell'SDK.
    if (est === '.html' && montaggio[0] === '/') {
      const html = fs.readFileSync(file, 'utf8');
      const relativo = path.relative(montaggio[1], file).split(path.sep).join('/');
      if (relativo.indexOf('rubrica/') === 0) {
        if (!TESTA_RUBRICA.test(html)) console.log('ATTENZIONE: in ' + relativo + ' non trovo il foglio di stile app.css: pagina servita senza raccolta degli errori');
        return rispondi(res, 200, MIME[est], html.replace(TESTA_RUBRICA, (tag) => '<script src="/banco-errori.js"></script>\n  ' + tag));
      }
      if (!SDK.test(html)) console.log('ATTENZIONE: in ' + relativo + ' non trovo il tag dell\'SDK Supabase: pagina servita senza sessione di prova');
      const sloggato = url.searchParams.has('sloggato') ? '?sloggato' : '';
      let pagina = html.replace(SDK, (tag) => '<script src="/banco-avvio.js' + sloggato + '"></script>\n  ' + tag);
      // ?senzaFiltro: la pagina arriva senza la libreria del filtro dell'HTML,
      // come se il suo caricamento fosse fallito (l'app deve fermarsi e dirlo).
      if (url.searchParams.has('senzaFiltro')) pagina = pagina.replace(/<script\b[^>]*\bpurify[^>]*><\/script>/i, '<!-- banco: filtro tolto -->');
      return rispondi(res, 200, MIME[est], pagina);
    }
    rispondi(res, 200, MIME[est] || 'application/octet-stream', fs.readFileSync(file));
  } catch (e) {
    rispondi(res, 500, 'text/plain; charset=utf-8', 'errore del banco: ' + e.message);
  }
}).listen(PORTA, '127.0.0.1', () => console.log('banco di prova su http://localhost:' + PORTA + '/  (app: ' + (COMPRESSA ? 'copia COMPRESSA di docs/' : 'docs/') + ' · finto TrakCare: /trak-finto/)'));
