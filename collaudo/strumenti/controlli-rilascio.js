// Controlli da fare prima di ogni pubblicazione, su ciò che sta in docs/.
// Sono i controlli che un test nel browser non vede o vede tardi:
//
//  1. sintassi dei file .js e degli script scritti dentro index.html e print.html
//  2. i due segnalibri TrakCare compilano anche collassati su una riga; se il
//     loro codice è cambiato, deve essere cambiato anche il numero di versione
//  3. nessun tag <script src> perso (un taglio sbagliato li porta via senza che
//     la sintassi se ne accorga) e ordine giusto: libreria del filtro →
//     sanifica.js → api.js, in tutte e due le pagine
//  4. la libreria del filtro nel repository è quella dichiarata: la sua impronta
//     è quella scritta in «integrity» nelle due pagine
//  5. service worker: versione della cache alzata, e ogni file dell'elenco esiste
//  6. ogni signOut dichiara scope 'local' (quello predefinito scollegherebbe
//     tutti i PC del reparto, che usano lo stesso account)
//  7. la versione scritta nel menu della rotellina è la stessa di CACHE_NAME
//  8. indirizzi: ogni ambiente elenca i suoi, nessuno è di tutti e due, e un
//     indirizzo sconosciuto non è di nessuno (l'app lì non deve partire)
//  9. fra i file pubblicati su Cloudflare ci sono tutti quelli che le pagine e
//     il service worker caricano
// 10. flusso di pubblicazione: azione di GitHub fissata all'identificativo
//     completo, strumento con versione fissa e impronta di ogni pacchetto
// 11. pagina di rinvio (ciò che resta al vecchio indirizzo): porta al nuovo con
//     gli stessi parametri, e il suo service worker si toglie da solo
// 12. copia compressa (ciò che viene pubblicato su Cloudflare): strumento con
//     versione fissa e impronte; la copia si costruisce e supera i suoi
//     controlli: script che compilano, struttura delle pagine e nomi globali
//     identici, nessun indirizzo di rete nuovo, segnalibri interi
// 13. rubrica («Numeri Telefono», docs/rubrica/): solo i suoi tre file, nessun
//     service worker né manifesto, nessun indirizzo o chiave nel codice, regola
//     CSP stretta e niente scritto in linea, richieste e scritture in memoria
//     da un punto solo, e il riquadro dell'app apre proprio quella pagina
//
//   node collaudo/strumenti/controlli-rilascio.js [riferimento]
//
// «riferimento» è la versione con cui confrontare (predefinito: origin/master,
// cioè ciò che è in produzione).
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const { execFileSync } = require('child_process');

const RADICE = path.resolve(__dirname, '../..');
const RIF = process.argv[2] || 'origin/master';
let errori = 0;
const ko = (t) => { errori++; console.log('KO  ' + t); };
const ok = (t) => console.log('ok  ' + t);
const leggi = (f) => fs.readFileSync(path.join(RADICE, f), 'utf8');
function diRif(f) {
  try { return execFileSync('git', ['-C', RADICE, 'show', RIF + ':' + f], { encoding: 'utf8', maxBuffer: 64 * 1024 * 1024, stdio: ['ignore', 'pipe', 'ignore'] }); }
  catch (e) { return null; }
}

// ── 1. sintassi ──────────────────────────────────────────────────────────
function compila(file) {
  try { execFileSync(process.execPath, ['--check', path.join(RADICE, file)], { stdio: 'pipe' }); ok('sintassi ' + file); }
  catch (e) { ko('sintassi ' + file + ': ' + String(e.stderr || e.message).split('\n').slice(0, 4).join(' | ')); }
}
fs.readdirSync(path.join(RADICE, 'docs/js')).filter((f) => f.endsWith('.js')).forEach((f) => compila('docs/js/' + f));
compila('docs/sw.js');
compila('docs/rubrica/app.js');
['docs/index.html', 'docs/print.html', 'docs/rubrica/index.html'].forEach((f) => {
  const html = leggi(f);
  const re = /<script(?![^>]*\bsrc=)[^>]*>([\s\S]*?)<\/script>/gi;
  let m, n = 0, male = 0;
  while ((m = re.exec(html))) {
    n++;
    try { new Function(m[1]); } catch (e) { male++; ko(f + ', script scritto in pagina n. ' + n + ': ' + e.message); }
  }
  if (!male) ok(f + ': ' + n + ' script scritti in pagina compilano');
});

// ── 2. segnalibri ────────────────────────────────────────────────────────
function segnalibro(html, nome, fine) {
  if (!html) return null;
  const i0 = html.indexOf('function ' + nome + '(){');
  const i1 = html.indexOf(fine);
  if (i0 < 0 || i1 < 0) return null;
  return html.slice(i0, i1).trim().replace(/\}\s*$/, '}').replace(/\r\n/g, '\n');
}
const ora = leggi('docs/index.html'), prima = diRif('docs/index.html');
if (prima == null) console.log('--  riferimento «' + RIF + '» non leggibile: salto i confronti con la versione precedente');
[['_bmTerapiaSorgente', "var _BM_TERAPIA = 'javascript:(", '_BM_VERSIONE'], ['_bmLabSorgente', "var _BM_LAB = 'javascript:(", '_LAB_BM_VERSIONE']].forEach((c) => {
  const a = segnalibro(ora, c[0], c[1]), b = segnalibro(prima, c[0], c[1]);
  if (!a) return ko(c[0] + ': non trovato in index.html');
  try { new Function('return (' + a.replace(/\n\s*/g, ' ') + ')'); ok(c[0] + ' collassato su una riga: compila'); }
  catch (e) { ko(c[0] + ' collassato su una riga: ' + e.message + ' (un commento // al suo interno?)'); }
  const versione = (s) => s ? (s.match(new RegExp('var ' + c[2] + '\\s*=\\s*([0-9]+)')) || [])[1] : null;
  const va = versione(ora), vb = versione(prima);
  if (b == null) return;
  if (a === b) ok(c[0] + ': identico a ' + RIF + ' (' + c[2] + ' = ' + va + ')');
  else if (va && vb && Number(va) > Number(vb)) ok(c[0] + ': cambiato, e ' + c[2] + ' è salito da ' + vb + ' a ' + va);
  else ko(c[0] + ': cambiato rispetto a ' + RIF + ' ma ' + c[2] + ' è rimasto ' + va + ' → va alzato (i colleghi devono ritrascinare il segnalibro)');
});

// ── 3 e 4. tag script e libreria del filtro ──────────────────────────────
const tagSrc = (html) => (html.match(/<script\b[^>]*\bsrc=["'][^"']+["'][^>]*>/gi) || []).map((t) => ({ tag: t, src: (t.match(/\bsrc=["']([^"']+)/) || [])[1] }));
[['docs/index.html', ora, prima], ['docs/print.html', leggi('docs/print.html'), diRif('docs/print.html')], ['docs/rubrica/index.html', leggi('docs/rubrica/index.html'), diRif('docs/rubrica/index.html'), 'senza filtro']].forEach((x) => {
  const adesso = tagSrc(x[1]);
  const nomi = adesso.map((t) => t.src);
  if (x[2] != null) {
    const persi = tagSrc(x[2]).map((t) => t.src).filter((s) => nomi.indexOf(s) < 0);
    persi.length ? ko(x[0] + ': tag script presenti in ' + RIF + ' e ora ASSENTI: ' + persi.join(', ')) : ok(x[0] + ': ' + adesso.length + ' tag script con src, nessuno perso rispetto a ' + RIF);
  }
  // la pagina della rubrica mostra solo testo, codificato dal suo codice: non carica il filtro,
  // e per lei basta il controllo sui tag persi
  if (x[3]) return;
  const lib = adesso.filter((t) => /(^|\/)purify[^/]*\.js$/.test(t.src))[0];
  if (!lib) return ko(x[0] + ': manca il tag della libreria del filtro (DOMPurify)');
  const iL = nomi.indexOf(lib.src), iS = nomi.indexOf('js/sanifica.js'), iA = nomi.indexOf('js/api.js');
  (iL >= 0 && iL < iS && iS < iA) ? ok(x[0] + ': ordine libreria del filtro → sanifica.js → api.js') : ko(x[0] + ': ordine sbagliato (libreria ' + iL + ', sanifica.js ' + iS + ', api.js ' + iA + ')');
  if (/^https?:/i.test(lib.src)) return ko(x[0] + ': la libreria del filtro è caricata da fuori (' + lib.src + ') invece che dal repository');
  const file = path.join(RADICE, 'docs', lib.src);
  if (!fs.existsSync(file)) return ko(x[0] + ': il file ' + lib.src + ' non esiste');
  const impronta = 'sha256-' + crypto.createHash('sha256').update(fs.readFileSync(file)).digest('base64');
  const dichiarata = (lib.tag.match(/\bintegrity=["']([^"']+)/) || [])[1];
  if (!dichiarata) ko(x[0] + ': il tag della libreria non ha «integrity»');
  else if (dichiarata !== impronta) ko(x[0] + ': l\'impronta di ' + lib.src + ' è ' + impronta + ' ma la pagina dichiara ' + dichiarata + ' → il browser rifiuterebbe la libreria e l\'app si fermerebbe');
  else ok(x[0] + ': ' + lib.src + ' corrisponde all\'impronta dichiarata (' + fs.statSync(file).size + ' byte)');
});
try {
  const attr = execFileSync('git', ['-C', RADICE, 'check-attr', 'text', '--', 'docs/js/librerie/purify-3.4.16.min.js'], { encoding: 'utf8' });
  /text: unset/.test(attr) ? ok('git non converte i fine riga dei file in docs/js/librerie/') : ko('.gitattributes: manca «docs/js/librerie/* -text» (' + attr.trim() + ')');
} catch (e) { ko('git check-attr: ' + e.message); }

// ── 5. service worker ────────────────────────────────────────────────────
const sw = leggi('docs/sw.js'), swPrima = diRif('docs/sw.js');
const versione = (s) => Number(((s || '').match(/consegne-v(\d+)/) || [])[1]);
if (swPrima != null) {
  let cambiato = true;
  try { cambiato = execFileSync('git', ['-C', RADICE, 'diff', '--name-only', RIF, '--', 'docs'], { encoding: 'utf8', stdio: ['ignore', 'pipe', 'ignore'] }).trim() !== ''; } catch (e) {}
  if (!cambiato) ok('docs/ identica a ' + RIF + ': nessun cambio di versione necessario');
  else (versione(sw) > versione(swPrima)) ? ok('CACHE_NAME: v' + versione(swPrima) + ' → v' + versione(sw)) : ko('docs/ è cambiata ma CACHE_NAME è rimasto v' + versione(sw) + ': i PC resterebbero sulla versione vecchia');
}
const elenco = ((sw.match(/PRECACHE_ASSETS\s*=\s*\[([\s\S]*?)\]/) || [])[1] || '').match(/'[^']+'/g) || [];
const mancanti = elenco.map((s) => s.slice(1, -1)).filter((p) => p !== './' && !fs.existsSync(path.join(RADICE, 'docs', p)));
mancanti.length ? ko('sw.js elenca file che non esistono (l\'installazione del service worker fallirebbe): ' + mancanti.join(', ')) : ok('sw.js: i ' + elenco.length + ' file tenuti in cache esistono tutti');
['./js/sanifica.js', './js/api.js', './js/app.js', './js/app2.js', './rubrica/', './rubrica/index.html', './rubrica/app.css', './rubrica/app.js'].concat(tagSrc(ora).map((t) => t.src).filter((s) => /purify/.test(s)).map((s) => './' + s)).forEach((p) => {
  if (elenco.indexOf("'" + p + "'") < 0) ko('sw.js: ' + p + ' non è fra i file tenuti in cache');
});

// ── 6. uscita dall'applicazione ──────────────────────────────────────────
// signOut() senza argomenti è «global»: revoca la sessione dell'account su
// tutti i dispositivi, e i PC del reparto usano lo stesso account.
{
  const sorgenti = fs.readdirSync(path.join(RADICE, 'docs/js')).filter((f) => f.endsWith('.js')).map((f) => 'docs/js/' + f).concat(['docs/index.html', 'docs/print.html', 'docs/rubrica/app.js', 'docs/rubrica/index.html']);
  let chiamate = 0, sbagliate = 0;
  sorgenti.forEach((f) => {
    const re = /\.signOut\s*\(([^)]*)\)/g;
    const testo = leggi(f);
    let m;
    while ((m = re.exec(testo))) {
      chiamate++;
      if (!/scope\s*:\s*['"]local['"]/.test(m[1])) { sbagliate++; ko(f + ': signOut(' + m[1].trim() + ') senza scope \'local\' → scollegherebbe TUTTI i PC del reparto'); }
    }
  });
  if (!sbagliate) ok('uscita: ' + chiamate + ' chiamate a signOut, tutte con scope \'local\'');
}

// ── 7. versione mostrata nel menu ────────────────────────────────────────
// Il menu della rotellina mostra la versione dell'applicazione: è scritta in
// index.html e deve essere la stessa cifra di CACHE_NAME in sw.js.
{
  const m = /id="navVersioneApp"[^>]*>(\d+)</.exec(ora);
  if (!m) ko('index.html: manca la versione dell\'applicazione nel menu della rotellina (id="navVersioneApp")');
  else if (Number(m[1]) !== versione(sw)) ko('index.html mostra la versione ' + m[1] + ' ma CACHE_NAME in sw.js è v' + versione(sw) + ': vanno alzate insieme');
  else ok('versione nel menu della rotellina: ' + m[1] + ', la stessa di CACHE_NAME');
}

// ── 8. indirizzi e ambienti ──────────────────────────────────────────────
// Ogni ambiente elenca i propri indirizzi (api.js): nessuno deve stare in tutti
// e due, fra quelli di produzione non deve essercene uno di prova, e un
// indirizzo sconosciuto non deve valere né da collaudo né da produzione.
{
  const api = leggi('docs/js/api.js').replace(/\r\n/g, '\n');
  const i0 = api.indexOf('var _AMBIENTI = {'), i1 = api.indexOf('var AMBIENTE = ');
  let f = null;
  if (i0 < 0 || i1 < i0) ko('api.js: blocco degli ambienti non trovato');
  else {
    try { f = new Function(api.slice(i0, i1) + '\nreturn { ambienti: _AMBIENTI, da: _ambienteDaHost };')(); }
    catch (e) { ko('api.js: il blocco degli ambienti non compila: ' + e.message); }
  }
  if (f) {
    const P = f.ambienti.produzione.host, C = f.ambienti.collaudo.host;
    const doppi = P.filter((h) => C.indexOf(h) >= 0);
    doppi.length ? ko('indirizzi presenti in tutti e due gli ambienti: ' + doppi.join(', ')) : ok('ambienti: ' + P.length + ' indirizzi di produzione, ' + C.length + ' di collaudo, nessuno in comune');
    const diProva = P.filter((h) => /collaudo|gistech|localhost|127[.]0[.]0[.]1/.test(h));
    if (diProva.length) ko('fra gli indirizzi di PRODUZIONE ce n\'è uno di prova: ' + diProva.join(', '));
    if (f.ambienti.produzione.supabaseUrl === f.ambienti.collaudo.supabaseUrl) ko('i due ambienti puntano allo stesso database');
    const attesi = [['medicinadurgenzaucsc-maker.github.io', 'produzione'], ['gistech2026.github.io', 'collaudo'], ['localhost', 'collaudo'], ['LOCALHOST', 'collaudo'],
      ['', null], ['example.org', null], ['github.io', null], ['pages.dev', null], ['xgistech2026.github.io', null],
      ['gistech2026.github.io.example.org', null], ['medicinadurgenzaucsc-maker.github.io.example.org', null]];
    // una voce col punto vale per i sottodomini, non per un nome incollato davanti né in coda
    [[P, 'produzione'], [C, 'collaudo']].forEach((x) => x[0].filter((h) => h.charAt(0) === '.').forEach((h) => {
      attesi.push(['anteprima' + h, x[1]], ['falso' + h.slice(1), null], [h.slice(1) + '.example.org', null]);
    }));
    // ogni indirizzo esatto è del suo ambiente; un sottodominio davanti vale solo se c'è anche la voce col punto
    [[P, 'produzione'], [C, 'collaudo']].forEach((x) => x[0].filter((h) => h.charAt(0) !== '.').forEach((h) => {
      attesi.push([h, x[1]], ['anteprima.' + h, x[0].indexOf('.' + h) >= 0 ? x[1] : null], [h + '.example.org', null]);
    }));
    // in produzione le anteprime di Cloudflare non devono essere riconosciute: sui pazienti veri si lavora solo dal sito
    if (P.some((h) => h.charAt(0) === '.')) ko('fra gli indirizzi di PRODUZIONE c\'è una voce col punto: le anteprime lavorerebbero sui pazienti veri');
    const sbagliati = attesi.filter((x) => f.da(x[0]) !== x[1]).map((x) => '«' + x[0] + '» → ' + f.da(x[0]) + ' (atteso ' + x[1] + ')');
    sbagliati.length ? ko('riconoscimento degli indirizzi: ' + sbagliati.join('; ')) : ok('riconoscimento degli indirizzi: ' + attesi.length + ' casi giusti, sconosciuti compresi');
  }
}

// ── 9. i file pubblicati su Cloudflare ───────────────────────────────────
// Su Cloudflare arriva una copia dei soli file del sito (prepara-sito.js): deve
// contenere tutto ciò che le pagine e il service worker caricano.
{
  const { elencoSito } = require('./prepara-sito.js');
  const sito = elencoSito();
  const locale = (u) => u.split('#')[0].split('?')[0].replace(/^[.][/]/, '');
  const servono = new Set(['index.html', 'print.html', 'sw.js', '404.html', '_headers']);
  // una voce che finisce con la barra è una cartella: il file che la serve è il suo index.html
  elenco.map((s) => s.slice(1, -1)).filter((p) => p !== './').forEach((p) => servono.add(locale(p.slice(-1) === '/' ? p + 'index.html' : p)));
  // [pagina, cartella rispetto a cui valgono i suoi indirizzi]
  [[ora, ''], [leggi('docs/print.html'), ''], [leggi('docs/rubrica/index.html'), 'rubrica/']].forEach((x) => {
    (x[0].match(/<(?:script|link)\b[^>]*\b(?:src|href)=["'][^"']+["']/gi) || []).forEach((t) => {
      const u = (t.match(/\b(?:src|href)=["']([^"']+)/) || [])[1];
      if (u && !/^(https?:|data:|#)/i.test(u)) servono.add(path.posix.normalize(x[1] + locale(u)));
    });
  });
  try { JSON.parse(leggi('docs/manifest.json')).icons.forEach((i) => servono.add(locale(i.src))); } catch (e) { ko('manifest.json non leggibile: ' + e.message); }
  const assenti = [...servono].filter((p) => p && sito.indexOf(p) < 0);
  assenti.length
    ? ko('file caricati dalle pagine ma NON fra quelli pubblicati su Cloudflare (non registrati in git, o in una cartella dal nome che comincia con «_»): ' + assenti.join(', '))
    : ok('sito: ' + sito.length + ' file da pubblicare, compresi i ' + servono.size + ' che pagine e service worker caricano');
  const fuoriCache = sito.filter((p) => ['_headers', '_redirects', '404.html', 'sw.js'].indexOf(p) < 0 && elenco.indexOf("'./" + p + "'") < 0);
  if (fuoriCache.length) ko('sw.js: file del sito che il service worker non tiene in cache: ' + fuoriCache.join(', '));
}

// ── 10. flusso di pubblicazione su Cloudflare ────────────────────────────
// Ciò che pubblica il sito può cambiare il codice che arriva ai PC del reparto:
// l'azione di GitHub va fissata al suo identificativo completo, lo strumento di
// pubblicazione ha versione fissa e l'impronta di ogni pacchetto.
{
  compila('.github/pubblicazione/passi.js');
  compila('collaudo/strumenti/prepara-sito.js');
  const flusso = leggi('.github/workflows/pubblica-cloudflare.yml');
  const usi = (flusso.match(/^\s*(?:-\s*)?uses:\s*\S+/gm) || []).map((u) => u.trim());
  const liberi = usi.filter((u) => !/@[0-9a-f]{40}$/.test(u));
  liberi.length ? ko('pubblica-cloudflare.yml: azioni non fissate all\'identificativo completo: ' + liberi.join(', ')) : ok('pubblica-cloudflare.yml: ' + usi.length + ' azione di GitHub, fissata all\'identificativo completo');
  if (!/npm ci --ignore-scripts/.test(flusso)) ko('pubblica-cloudflare.yml: lo strumento va installato con «npm ci --ignore-scripts»');
  try {
    const pacchetto = JSON.parse(leggi('.github/pubblicazione/package.json')), impronte = JSON.parse(leggi('.github/pubblicazione/package-lock.json'));
    const voluta = pacchetto.dependencies.wrangler, fissata = (impronte.packages['node_modules/wrangler'] || {}).version;
    const nomi = Object.keys(impronte.packages).filter(Boolean);
    const senza = nomi.filter((k) => !impronte.packages[k].integrity);
    if (!/^\d+[.]\d+[.]\d+$/.test(voluta)) ko('strumento di pubblicazione: la versione di wrangler non è fissa (' + voluta + ')');
    else if (voluta !== fissata) ko('strumento di pubblicazione: package.json chiede wrangler ' + voluta + ' ma le impronte sono della ' + fissata + ' (rigenerare package-lock.json)');
    else if (senza.length) ko('strumento di pubblicazione: pacchetti senza impronta: ' + senza.slice(0, 5).join(', '));
    else ok('strumento di pubblicazione: wrangler ' + voluta + ', ' + nomi.length + ' pacchetti, tutti con impronta');
  } catch (e) { ko('strumento di pubblicazione: ' + e.message); }
}

// ── 11. pagina di rinvio ─────────────────────────────────────────────────
// È ciò che resta al vecchio indirizzo quando il sito si sposta: deve portare
// al nuovo con gli stessi parametri, anche per la stampa e per le pagine che
// non esistono, e il suo service worker deve solo togliersi di mezzo.
{
  compila('collaudo/strumenti/prepara-rinvio.js');
  compila('collaudo/strumenti/scambio-repository.js');
  try {
    const { fileRinvio } = require('./prepara-rinvio.js');
    const NUOVO = 'https://esempio-nuovo.pages.dev/';
    const file = Object.fromEntries(fileRinvio(NUOVO).map((x) => [x[0], x[1].toString('utf8')]));
    const vm = require('vm');
    const arriva = (nome, search, hash) => {
      let dove = null;
      vm.runInNewContext((/<script>([\s\S]*?)<\/script>/.exec(file[nome]) || [])[1] || '', { location: { search: search, hash: hash, replace: (u) => { dove = u; } } });
      return dove;
    };
    const casi = [['index.html', '', '', NUOVO], ['index.html', '?a=1', '#b', NUOVO + '?a=1#b'], ['print.html', '?layout=alt', '', NUOVO + 'print.html?layout=alt'], ['404.html', '?a=1', '', NUOVO + '?a=1']];
    const sbagliati = casi.filter((c) => arriva(c[0], c[1], c[2]) !== c[3]).map((c) => c[0] + c[1] + c[2] + ' → ' + arriva(c[0], c[1], c[2]));
    sbagliati.length ? ko('pagina di rinvio: ' + sbagliati.join('; ')) : ok('pagina di rinvio: ' + casi.length + ' casi portano al nuovo indirizzo con parametri e ancora');
    if (!/<noscript><meta http-equiv="refresh"/.test(file['index.html'])) ko('pagina di rinvio: manca il rinvio per chi ha JavaScript spento');
    if (/addEventListener\('fetch'/.test(file['sw.js']) || !/registration\.unregister\(\)/.test(file['sw.js'])) ko('rinvio/sw.js deve solo svuotare la cache e togliersi: niente gestore delle richieste');
    else ok('rinvio/sw.js: si toglie da solo e non intercetta richieste');
    let rifiutati = 0;
    ['http://esempio.org/', 'https://esempio.org/?a=1', "https://x'y.org/", 'javascript:alert(1)'].forEach((u) => { try { fileRinvio(u); } catch (e) { rifiutati++; } });
    rifiutati === 4 ? ok('pagina di rinvio: gli indirizzi non validi sono rifiutati') : ko('pagina di rinvio: un indirizzo non valido è stato accettato');
  } catch (e) { ko('pagina di rinvio: ' + e.message); }
}

// ── 12. copia compressa ──────────────────────────────────────────────
// Lo strumento che comprime trasforma il codice che arriva ai PC: versione
// fissa, impronta di ogni pacchetto, e la copia prodotta deve superare i
// controlli di comprimi.js. Va installato in locale una volta:
//   npm ci --ignore-scripts   (dentro .github/compressione)
{
  compila('.github/compressione/comprimi.js');
  try {
    const pacchetto = JSON.parse(leggi('.github/compressione/package.json')), impronte = JSON.parse(leggi('.github/compressione/package-lock.json'));
    const nomi = Object.keys(impronte.packages).filter(Boolean);
    const senza = nomi.filter((k) => !impronte.packages[k].integrity);
    const dichiarati = Object.keys(pacchetto.dependencies || {});
    const liberi = dichiarati.filter((d) => !/^\d+[.]\d+[.]\d+$/.test(pacchetto.dependencies[d]) || (impronte.packages['node_modules/' + d] || {}).version !== pacchetto.dependencies[d]);
    if (dichiarati.indexOf('terser') < 0) ko('strumento di compressione: terser non è fra le dipendenze');
    else if (liberi.length) ko('strumento di compressione: versione non fissa o diversa da quella delle impronte: ' + liberi.join(', ') + ' (rigenerare package-lock.json)');
    else if (senza.length) ko('strumento di compressione: pacchetti senza impronta: ' + senza.slice(0, 5).join(', '));
    else ok('strumento di compressione: terser ' + pacchetto.dependencies.terser + ', ' + nomi.length + ' pacchetti, tutti con impronta');
    const flusso = leggi('.github/workflows/pubblica-cloudflare.yml');
    if ((flusso.match(/npm ci --ignore-scripts/g) || []).length < 2) ko('pubblica-cloudflare.yml: anche lo strumento di compressione va installato con «npm ci --ignore-scripts»');
    if (!/[.]github\/compressione\/\*\*/.test(flusso)) ko('pubblica-cloudflare.yml: una modifica a .github/compressione deve far ripartire la pubblicazione');
  } catch (e) { ko('strumento di compressione: ' + e.message); }
  try {
    const r = require('./prepara-sito.js').prepara(path.join(RADICE, '_sito', 'controlli'), { compressa: true });
    const c = r.compressione;
    ok('copia compressa: ' + r.elenco.length + ' file, da ' + Math.round(c.prima / 1024) + ' a ' + Math.round(c.dopo / 1024) + ' KB; script che compilano, struttura delle pagine e ' + c.nomiGlobali + ' nomi globali identici, nessun indirizzo nuovo, segnalibri interi');
  } catch (e) { ko('copia compressa: ' + String((e && e.message) || e).split('\n').join(' | ').slice(0, 900)); }
}

// ── 13. rubrica («Numeri Telefono») ──────────────────────────────────
// La rubrica è una pagina dell'app, della stessa origine: da lì si arriva alla
// sessione e alle schede. Deve restare com'è nata: tre file, nessun service
// worker né manifesto suoi, nessun indirizzo e nessuna chiave nel codice (li
// chiede alla pagina che la contiene), niente scritto in linea, e una regola
// CSP che spegne ciò che dovesse sfuggire. Nomi e numeri non finiscono nella
// memoria del browser: le scritture passano da una sola funzione, le richieste
// da un'altra.
{
  const ATTESI = ['app.css', 'app.js', 'index.html'];
  const suDisco = fs.readdirSync(path.join(RADICE, 'docs/rubrica')).sort();
  const inGit = require('./prepara-sito.js').elencoSito().filter((p) => p.indexOf('rubrica/') === 0).map((p) => p.slice('rubrica/'.length)).sort();
  (suDisco.join() === ATTESI.join() && inGit.join() === ATTESI.join())
    ? ok('rubrica: in docs/rubrica/ solo i tre file attesi, tutti registrati in git')
    : ko('rubrica: in docs/rubrica/ devono esserci solo ' + ATTESI.join(', ') + ' (niente sw.js, manifesto o icone), tutti registrati in git. Su disco: ' + suDisco.join(', ') + '; in git: ' + (inGit.join(', ') || 'nessuno'));

  const codice = leggi('docs/rubrica/app.js'), pagina = leggi('docs/rubrica/index.html'), stile = leggi('docs/rubrica/app.css');
  const guai = [];
  [
    [/serviceWorker/, 'un service worker'],
    [/https?:\/\/[A-Za-z0-9]/, 'un indirizzo di rete (indirizzo e chiave si chiedono alla pagina madre)'],
    [/eyJ[A-Za-z0-9_-]{10,}/, 'una chiave'],
    [/createClient|\.signOut\b/, 'un client del database proprio o un\'uscita dalla sessione'],
    [/sessionStorage|indexedDB|\bcaches\b/, 'un\'altra memoria del browser'],
    [/\bstyle\s*=|\son[a-z]+\s*=/, 'uno stile o un gestore scritto in linea nell\'HTML che costruisce (la regola CSP lo spegnerebbe)'],
    [/\beval\s*\(|new Function|document\.write|insertAdjacentHTML|outerHTML/, 'una costruzione che esegue o inserisce testo senza controllo'],
  ].forEach((v) => { if (v[0].test(codice)) guai.push('app.js contiene ' + v[1]); });
  const quanti = (re) => (codice.match(re) || []).length;
  if (quanti(/\.setItem\s*\(/g) !== 1 || !/function lsScrivi\(/.test(codice)) guai.push('app.js: le scritture nella memoria del browser devono passare tutte da lsScrivi (trovati ' + quanti(/\.setItem\s*\(/g) + ' setItem)');
  if (quanti(/\bfetch\s*\(/g) !== 1 || !/function supaRisposta\(/.test(codice)) guai.push('app.js: le richieste devono partire tutte da supaRisposta (trovati ' + quanti(/\bfetch\s*\(/g) + ' fetch)');
  if (!/window\.parent/.test(codice)) guai.push('app.js non chiede più nulla alla pagina madre');

  const csp = (pagina.match(/<meta\s+http-equiv="Content-Security-Policy"\s+content="([^"]*)"/i) || [])[1];
  if (!csp) guai.push('index.html: manca la regola CSP');
  else {
    const regole = {};
    csp.split(';').map((r) => r.trim()).filter(Boolean).forEach((r) => { const p = r.split(/\s+/); regole[p[0]] = p.slice(1).join(' '); });
    const attese = { 'default-src': "'none'", 'script-src': "'self'", 'style-src': "'self'", 'base-uri': "'none'", 'form-action': "'none'" };
    Object.keys(attese).forEach((k) => { if (regole[k] !== attese[k]) guai.push('index.html, regola CSP: «' + k + '» vale «' + (regole[k] || 'assente') + '», atteso ' + attese[k]); });
    if (regole['connect-src'] !== 'https://*.supabase.co') guai.push('index.html, regola CSP: «connect-src» deve ammettere solo il database (https://*.supabase.co), trovato «' + (regole['connect-src'] || 'assente') + '»');
    const altre = Object.keys(regole).filter((k) => k !== 'connect-src').map((k) => regole[k]).join(' ');
    if (/unsafe|[*]|data:|blob:|https?:/.test(altre)) guai.push('index.html, regola CSP: una voce larga (unsafe, *, data:, blob: o un sito esterno)');
    if (pagina.indexOf('http-equiv="Content-Security-Policy"') > pagina.search(/<(?:script|link)\b/i)) guai.push('index.html: la regola CSP deve venire prima di ogni script e di ogni foglio di stile');
  }
  if (/<script(?![^>]*\bsrc=)/i.test(pagina)) guai.push('index.html: uno script scritto in pagina');
  if (/<style\b|\sstyle\s*=/i.test(pagina)) guai.push('index.html: uno stile scritto in pagina');
  if (/<[^>]*\son[a-z]+\s*=/i.test(pagina)) guai.push('index.html: un gestore scritto in linea');
  if (/manifest|<iframe|<object|<embed|<base\b/i.test(pagina)) guai.push('index.html: manifesto, cornice o elemento non ammesso');
  const script = tagSrc(pagina).map((t) => t.src);
  if (script.join() !== 'app.js') guai.push('index.html: atteso un solo script, app.js (trovati: ' + (script.join(', ') || 'nessuno') + ')');
  if (/url\s*\(|@import|https?:/i.test(stile)) guai.push('app.css: un indirizzo o un\'importazione');

  // lato app: il riquadro apre la pagina dell'app, senza «sandbox», e l'uscita la svuota
  if (!/var RUBRICA_URL = 'rubrica\/';/.test(ora) || /rubrica-gemelli/.test(ora)) guai.push('docs/index.html: il riquadro deve aprire la pagina dell\'app («rubrica/»), non un sito esterno');
  if (/id="rubricaIframe"[^>]*\bsandbox\b/.test(ora.replace(/\s+/g, ' '))) guai.push('docs/index.html: la cornice della rubrica non deve avere «sandbox» (la rubrica chiede la sessione alla pagina)');
  if (!/function _eseguiUscita\(\)[\s\S]{0,1600}?_rubricaSvuotaCopie\(\)/.test(ora)) guai.push('docs/index.html: l\'uscita (_eseguiUscita) non svuota più la rubrica');
  guai.length ? guai.forEach((g) => ko('rubrica: ' + g)) : ok('rubrica: nessun service worker, indirizzo o chiave nel codice; regola CSP stretta e niente in linea; richieste e memoria da un punto solo; il riquadro apre «rubrica/»');
}

console.log(errori ? ('\n' + errori + ' CONTROLLI FALLITI: non pubblicare') : '\ntutti i controlli superati');
process.exit(errori ? 1 : 0);
