// Controlli da fare prima di ogni pubblicazione, su ciò che sta in docs/.
// Sono i controlli che un test nel browser non vede o vede tardi:
//
//  1. sintassi dei file .js e degli script scritti dentro index.html e print.html
//  2. i due segnalibri TrakCare compilano anche collassati su una riga; se il
//     loro codice è cambiato, deve essere cambiato anche il numero di versione
//  3. nessun tag <script src> perso (un taglio sbagliato li porta via senza che
//     la sintassi se ne accorga) e ordine giusto: ambiente.js → libreria del
//     filtro → sanifica.js → api.js, in tutte e due le pagine
//  4. la libreria del filtro nel repository è quella dichiarata: la sua impronta
//     è quella scritta in «integrity» nelle due pagine
//  5. service worker: versione della cache alzata, e ogni file dell'elenco esiste
//  6. ogni signOut dichiara scope 'local' (quello predefinito scollegherebbe
//     tutti i PC del reparto, che usano lo stesso account)
//  7. la versione scritta nel menu della rotellina è la stessa di CACHE_NAME
//  8. indirizzi: ogni reparto elenca i suoi (reparti/<codice>.json), nessuno è
//     di due reparti, e con le funzioni vere di api.js un sito parte solo dai
//     propri indirizzi: non da quelli di un altro, non da uno sconosciuto, non
//     con un file dell'ambiente sbagliato; la tabella di riserva rimasta in
//     api.js è identica a quella della v189 e concorda con i file dei reparti
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
//     (tutto guardato nel codice vivo: ciò che sta in un commento non conta)
// 14. reparti: i dati dei due ambienti scritti negli strumenti concordano con i
//     file dei reparti; il codice di un reparto vale da nome d'ambiente senza
//     togliere le barriere della produzione; il flusso riparte se cambia un
//     reparto; il sito di OGNI reparto si costruisce, non porta dati di un
//     altro e differisce dagli altri solo in js/ambiente.js (punto 12)
// 16. modo rubrica (pagina aperta con ?rubrica=1): il segno è lo stesso
//     dovunque; l'app parte da un punto solo, e in modo rubrica nessuna delle
//     strade che leggono da sole i dati dei pazienti resta aperta; il segno
//     sopravvive a «Update»; il secondo manifesto riapre lo stesso indirizzo
// 15. riservatezza: il piano «un codice, tanti reparti», che è un documento
//     riservato, non sta fra i file né nella storia che un rilascio porterebbe
//     nel repository del reparto, che è pubblico
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
if (process.argv.length > 3) { console.log('uso: controlli-rilascio.js [riferimento git]'); process.exit(2); }
// Il nome di un reparto non è un riferimento git. «collaudo» però è anche un ramo: i confronti con la
// versione precedente diventerebbero confronti del ramo con se stesso, e direbbero «nulla è cambiato».
if (process.argv[2] && require('./reparti.js').nomeStorico(process.argv[2])) {
  console.log('controlli-rilascio.js vuole un riferimento git, non il nome di un reparto: senza argomenti confronta con origin/master; per il ramo si scrive refs/heads/collaudo');
  process.exit(2);
}
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
['docs/index.html', 'docs/print.html', 'docs/rubrica/index.html', 'docs/rete.html'].forEach((f) => {
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
// Un riferimento scritto a mano e non leggibile è un errore (il nome di un reparto al posto di un ramo, per
// esempio): saltare i confronti in silenzio farebbe passare un segnalibro cambiato senza la sua versione.
if (prima == null && process.argv[2]) ko('riferimento «' + RIF + '» non leggibile: va scritto un ramo o un commit che esiste (i confronti con la versione precedente non si possono fare)');
else if (prima == null) console.log('--  riferimento «' + RIF + '» non leggibile: salto i confronti con la versione precedente');
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
  const iL = nomi.indexOf(lib.src), iS = nomi.indexOf('js/sanifica.js'), iA = nomi.indexOf('js/api.js'), iE = nomi.indexOf('js/ambiente.js');
  (iE >= 0 && iE < iL && iL < iS && iS < iA) ? ok(x[0] + ': ordine ambiente.js → libreria del filtro → sanifica.js → api.js') : ko(x[0] + ': ordine sbagliato (ambiente.js ' + iE + ', libreria ' + iL + ', sanifica.js ' + iS + ', api.js ' + iA + ')');
  // Il file dell'ambiente deve essere ESEGUITO prima di api.js: un solo tag, scritto semplice.
  // Con async, defer o type=module il browser lo eseguirebbe dopo, e api.js partirebbe dalla tabella di riserva.
  const tagE = adesso.filter((t) => (t.src || '').split('?')[0] === 'js/ambiente.js');
  if (tagE.length !== 1 || tagE[0].tag !== '<script src="js/ambiente.js">') ko(x[0] + ': il tag di js/ambiente.js deve essere uno solo e scritto così: <script src="js/ambiente.js"> (trovati: ' + (tagE.map((t) => t.tag).join(' ') || 'nessuno') + ')');
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
['./js/ambiente.js', './js/sanifica.js', './js/api.js', './js/app.js', './js/app2.js', './rubrica/', './rubrica/index.html', './rubrica/app.css', './rubrica/app.js'].concat(tagSrc(ora).map((t) => t.src).filter((s) => /purify/.test(s)).map((s) => './' + s)).forEach((p) => {
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
// Chi è un sito lo dice js/ambiente.js, generato dal file del suo reparto
// (reparti/<codice>.json); in api.js decide _ambienteEsame. Qui, con le
// funzioni VERE di api.js e i file VERI dei reparti, si prova che ogni sito
// parta dai propri indirizzi e da nessun altro, e che un file sbagliato lo
// fermi. È la sola prova possibile del file della 11N prima del rilascio: su
// questo PC quel file non si serve mai.
// TRANSITORIO: la tabella _AMBIENTI della v189 resta in api.js come riserva
// per i minuti di un aggiornamento. Deve restare identica a com'era e
// concordare con i file dei reparti; quando verrà tolta questo blocco andrà
// riscritto senza di lei (e con lui prove.js, prove-sw.js, prepara-sito.js).
{
  const TABELLA_V189 = 'faa288d7b25f9f58db86572cd608ac8aa10e06c543d9c717871547030d5ce690';
  const reparti = require('./reparti.js');
  const vm = require('vm');
  const api = leggi('docs/js/api.js').replace(/\r\n/g, '\n');
  const i0 = api.indexOf('var _AMBIENTI = {'), iT = api.indexOf('// Un nome host è coperto da una voce'), i1 = api.indexOf('var _AMB_ESAME = ');
  let tutti = null, g = null;
  try { tutti = reparti.tutti(); ok('reparti: ' + tutti.length + ' file validi (' + tutti.map((r) => r.codice).join(', ') + '); nessun nome, indirizzo, database o progetto in comune; nei reparti di produzione nessuna anteprima, né localhost né 127.0.0.1'); }
  catch (e) { ko('reparti: ' + String(e.message).split('\n').join(' | ')); }
  if (tutti) { const c = reparti.controllaDocs(); c.ok ? ok('docs/js/ambiente.js è quello generato per il reparto «' + c.reparto + '», che pubblica da docs/') : ko(c.motivo); }
  if (i0 < 0 || iT < i0 || i1 < iT) ko('api.js: blocco degli ambienti non trovato');
  else {
    try { g = new Function(api.slice(i0, i1) + '\nreturn { ambienti: _AMBIENTI, da: _ambienteDaHost, esame: _ambienteEsame, di: _ambienteDi };')(); }
    catch (e) { ko('api.js: il blocco degli ambienti non compila: ' + e.message); }
  }
  if (g && tutti) {
    // la tabella di riserva: quella della v189, e d'accordo con i file dei reparti
    (crypto.createHash('sha256').update(api.slice(i0, iT).trim()).digest('hex') === TABELLA_V189)
      ? ok('tabella di riserva in api.js: identica a quella della v189')
      : ko('tabella di riserva in api.js: è cambiata rispetto alla v189 (sta lì solo per i minuti di un aggiornamento: non si allunga e non si corregge)');
    Object.keys(g.ambienti).forEach((nome) => {
      const t = g.ambienti[nome];
      const r = tutti.filter((x) => x.ambiente === nome && x.supabaseUrl === t.supabaseUrl)[0];
      if (!r) return ko('tabella di riserva: nessun reparto ha l\'ambiente «' + nome + '» e il suo database');
      const diversi = [];
      if (r.supabaseAnonKey !== t.supabaseAnonKey) diversi.push('chiave pubblica');
      if (r.googleClientId !== t.googleClientId) diversi.push('client Google');
      const fuori = t.host.filter((h) => r.host.indexOf(h) < 0);
      if (fuori.length) diversi.push('indirizzi (la tabella conosce ' + fuori.join(', ') + ', il file del reparto no: lì il sito partirebbe o si fermerebbe a seconda del file arrivato)');
      diversi.length ? ko('reparti/' + r.codice + '.json e la tabella di riserva non concordano su: ' + diversi.join('; ')) : ok('reparti/' + r.codice + '.json concorda con la tabella di riserva («' + nome + '»: database, chiave pubblica, client Google, indirizzi)');
    });

    // il file di ogni reparto, così come lo riceve il browser
    const file = {};
    tutti.forEach((r) => { const w = {}; vm.runInNewContext(reparti.ambienteJs(r), { window: w }); file[r.codice] = w._AMBIENTE; });
    const male = [];
    let casi = 0;
    const atteso = (cosa, ottenuto, voluto) => { casi++; if (ottenuto !== voluto) male.push(cosa + ' → ' + ottenuto + ' (atteso ' + voluto + ')'); };
    const detto = (e) => e.esito ? 'PARTE' : 'fermo:' + e.motivo;
    const SCONOSCIUTI = ['', 'example.org', 'github.io', 'pages.dev', 'supabase.co', 'xgistech2026.github.io', 'localhost.example.org'];
    const CAMPI = ['ambiente', 'supabaseUrl', 'supabaseAnonKey', 'googleClientId'];
    tutti.forEach((r) => {
      const F = file[r.codice];
      r.host.forEach((h) => {
        const colPunto = h.charAt(0) === '.', nudo = h.replace(/^[.]/, '');
        // dai propri indirizzi il sito parte, dal file, con i propri dati
        (colPunto ? ['anteprima' + h, 'a.b' + h] : [h, h.toUpperCase()]).forEach((host) => {
          const e = g.esame(F, host);
          const suoi = e.esito && e.esito.ambiente === r.ambiente && e.esito.supabaseUrl === r.supabaseUrl && e.esito.supabaseAnonKey === r.supabaseAnonKey && e.esito.googleClientId === r.googleClientId;
          atteso('sito ' + r.codice + ' da «' + host + '»', e.esito ? (e.esito.origine + (suoi ? ', dati suoi' : ', DATI DI ALTRI')) : 'fermo:' + e.motivo, 'file, dati suoi');
        });
        // un nome incollato davanti o in coda non vale; un sottodominio vale solo con la voce col punto
        const finti = ['falso' + nudo, nudo + '.example.org'];
        if (!colPunto && r.host.indexOf('.' + h) < 0) finti.push('anteprima.' + h);
        finti.forEach((host) => atteso('sito ' + r.codice + ' da «' + host + '»', detto(g.esame(F, host)), 'fermo:indirizzo'));
      });
      SCONOSCIUTI.forEach((host) => atteso('sito ' + r.codice + ' da «' + host + '»', detto(g.esame(F, host)), 'fermo:indirizzo'));
      // dagli indirizzi di un ALTRO reparto non parte
      tutti.filter((x) => x.codice !== r.codice).forEach((x) => x.host.forEach((h) => {
        const host = h.charAt(0) === '.' ? 'anteprima' + h : h;
        atteso('sito ' + r.codice + ' dall\'indirizzo di ' + x.codice + ' «' + host + '»', detto(g.esame(F, host)), 'fermo:indirizzo');
      }));
      // un file che non è scritto come deve ferma l'app anche dal suo indirizzo
      const buono = r.host.filter((h) => h.charAt(0) !== '.')[0];
      [['vuoto', {}], ['una stringa', 'x'], ['zero', 0], ['falso', false], ['un elenco', []],
        ['senza host', Object.assign({}, F, { host: undefined })], ['host non in elenco', Object.assign({}, F, { host: buono })],
        ['ambiente sconosciuto', Object.assign({}, F, { ambiente: 'prova' })], ['database non https', Object.assign({}, F, { supabaseUrl: F.supabaseUrl.replace('https://', 'http://') })],
        ['senza chiave', Object.assign({}, F, { supabaseAnonKey: '' })], ['senza client', Object.assign({}, F, { googleClientId: null })],
      ].forEach((x) => atteso('sito ' + r.codice + ', file ' + x[0], detto(g.esame(x[1], buono)), 'fermo:file'));
      // il file di un altro sito messo su questo indirizzo (stessi indirizzi, database dell'altro)
      tutti.filter((x) => x.codice !== r.codice).forEach((x) => {
        if (!g.da(buono)) return;
        atteso('sito ' + r.codice + ' col file di ' + x.codice, detto(g.esame(Object.assign({}, file[x.codice], { host: r.host.slice() }), buono)), 'fermo:discorde');
      });
      // le due metà della regola, una per volta, col SUO file: solo il database cambiato, solo l'ambiente cambiato
      if (g.da(buono)) {
        atteso('sito ' + r.codice + ', suo file col database cambiato', detto(g.esame(Object.assign({}, F, { supabaseUrl: 'https://aaaaaaaaaaaaaaaaaaaa.supabase.co' }), buono)), 'fermo:discorde');
        atteso('sito ' + r.codice + ', suo file con l\'ambiente cambiato', detto(g.esame(Object.assign({}, F, { ambiente: r.ambiente === 'produzione' ? 'collaudo' : 'produzione' }), buono)), 'fermo:discorde');
      }
      // senza il file (pagina della versione di prima) vale la tabella, e dà gli stessi dati del file
      r.host.filter((h) => h.charAt(0) !== '.' && g.da(h)).forEach((h) => {
        const a = g.esame(undefined, h), b = g.esame(null, h), c = g.esame(F, h);
        const pari = a.esito && b.esito && c.esito && a.esito.origine === 'tabella' && b.esito.origine === 'tabella' && CAMPI.every((k) => a.esito[k] === c.esito[k] && b.esito[k] === c.esito[k]);
        atteso('sito ' + r.codice + ' senza file da «' + h + '»', pari ? 'tabella, stessi dati del file' : 'DIVERSO', 'tabella, stessi dati del file');
      });
    });
    // senza file, un indirizzo sconosciuto resta fermo
    SCONOSCIUTI.forEach((host) => atteso('senza file da «' + host + '»', detto(g.esame(undefined, host)), 'fermo:indirizzo'));
    atteso('_ambienteDi è l\'esito di _ambienteEsame', String(g.di({}, 'localhost') === null && g.di(undefined, 'example.org') === null && !!g.di(undefined, 'localhost')), 'true');
    male.length
      ? ko('riconoscimento degli indirizzi: ' + male.slice(0, 6).join('; ') + (male.length > 6 ? ' … e altri ' + (male.length - 6) : ''))
      : ok('riconoscimento degli indirizzi: ' + casi + ' casi giusti con le funzioni di api.js e i file dei reparti (indirizzi propri, di altri reparti, sconosciuti; file sbagliato, di un altro sito, assente)');
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
  // i due manifesti: quello delle consegne e quello dei numeri di telefono (modo rubrica)
  ['docs/manifest.json', 'docs/manifest-rubrica.json'].forEach((m) => {
    try { JSON.parse(leggi(m)).icons.forEach((i) => servono.add(locale(i.src))); } catch (e) { ko(m + ' non leggibile: ' + e.message); }
  });
  const assenti = [...servono].filter((p) => p && sito.indexOf(p) < 0);
  assenti.length
    ? ko('file caricati dalle pagine ma NON fra quelli pubblicati su Cloudflare (non registrati in git, o in una cartella dal nome che comincia con «_»): ' + assenti.join(', '))
    : ok('sito: ' + sito.length + ' file da pubblicare, compresi i ' + servono.size + ' che pagine e service worker caricano');
  // rete.html è la pagina di prova della rete: deve dire ciò che passa ADESSO, quindi non sta fra i file tenuti in cache
  const fuoriCache = sito.filter((p) => ['_headers', '_redirects', '404.html', 'sw.js', 'rete.html'].indexOf(p) < 0 && elenco.indexOf("'./" + p + "'") < 0);
  if (elenco.indexOf("'./rete.html'") >= 0) ko('sw.js: rete.html non deve stare fra i file tenuti in cache (è la pagina che prova la rete)');
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

// ── 11 bis. rinvio dal vecchio indirizzo della rubrica ───────────────
// La rubrica di prima aveva un indirizzo suo, con la sua icona sui telefoni.
// Lì resta un repository pubblico con la SOLA pagina che porta all'app delle
// consegne aperta per la rubrica, e un service worker che toglie quello di
// prima con la sua cache (ci stavano nomi e numeri) senza toccare le cache
// dell'app, che vive sullo stesso dominio.
{
  compila('collaudo/strumenti/rinvio-rubrica.js');
  try {
    const { fileRubrica, destinazioneVera, SEGNO_RUBRICA } = require('./rinvio-rubrica.js');
    const vm = require('vm');
    const D = destinazioneVera(require('./gh-api.js').AMBIENTI.produzione.utente);
    const file = Object.fromEntries(fileRubrica(D.intera).map((x) => [x[0], x[1].toString('utf8')]));
    const guai = [];
    if (Object.keys(file).sort().join() !== ['.nojekyll', '404.html', 'README.md', 'index.html', 'robots.txt', 'sw.js'].join()) guai.push('file inattesi: ' + Object.keys(file).sort().join(', '));
    if (D.intera !== D.app + SEGNO_RUBRICA || !/^https:\/\/[a-z0-9-]+[.]github[.]io\/app-consegne\/[?]rubrica=1$/.test(D.intera)) guai.push('destinazione inattesa');
    // le chiavi che tenevano nomi e numeri: lo stesso elenco che la rubrica dentro l'app toglie all'avvio
    const residui = ((/const LS_RESIDUI\s*=\s*\[([^\]]*)\]/.exec(leggi('docs/rubrica/app.js')) || [])[1] || '').match(/'[^']+'/g) || [];
    ['index.html', '404.html'].forEach((n) => {
      const codice = (/<script>([\s\S]*?)<\/script>/.exec(file[n]) || [])[1] || '';
      const giro = (inCornice) => {
        let dove = null; const tolte = [];
        const finestra = {}; finestra.self = finestra; finestra.top = inCornice ? {} : finestra;
        vm.runInNewContext(codice, { window: finestra, localStorage: { removeItem: (k) => { tolte.push(k); } }, location: { search: '?x=1', hash: '#y', replace: (u) => { dove = u; } } });
        return { dove, tolte };
      };
      const a = giro(false), b = giro(true);
      if (a.dove !== D.intera) guai.push(n + ' porta a «' + a.dove + '» invece che all\'app aperta per la rubrica');
      if (b.dove !== null) guai.push(n + ' rimanda anche dentro una cornice: l\'app intera finirebbe nel riquadro di una vecchia pagina');
      if (!residui.length || a.tolte.slice().sort().join() !== residui.map((x) => x.slice(1, -1)).sort().join()) guai.push(n + ' deve togliere dal dispositivo le stesse chiavi residue della rubrica di prima che toglie docs/rubrica/app.js (LS_RESIDUI), e solo quelle');
      if (!/<noscript><meta http-equiv="refresh"/.test(file[n])) guai.push(n + ': manca il rinvio per chi ha JavaScript spento');
      if (/eyJ[A-Za-z0-9_-]{10,}|supabase/i.test(file[n])) guai.push(n + ' contiene una chiave o un indirizzo di database');
    });
    const sw = file['sw.js'];
    if (/addEventListener\('fetch'/.test(sw)) guai.push('sw.js non deve avere un gestore delle richieste');
    if (!/registration\.unregister\(\)/.test(sw)) guai.push('sw.js deve togliersi da solo');
    if ((sw.match(/caches\.delete\(/g) || []).length !== 1 || !/nomi\.filter\(function \(n\) \{ return n\.indexOf\('rubrica-'\) === 0; \}\)\.map\(function \(n\) \{ return caches\.delete\(n\); \}\)/.test(sw)) guai.push('sw.js deve togliere SOLO le cache «rubrica-…»: sullo stesso dominio vive l\'app delle consegne');
    guai.length ? ko('rinvio della rubrica: ' + guai.join('; ')) : ok('rinvio della rubrica: 6 file, la pagina porta a ' + D.intera + ' (non dentro una cornice) e toglie la vecchia copia dei numeri; il service worker toglie solo le cache «rubrica-…» e se stesso');
  } catch (e) { ko('rinvio della rubrica: ' + String((e && e.message) || e).slice(0, 300)); }
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
  // La copia si costruisce per OGNI reparto: ognuna supera i controlli di
  // comprimi.js e quello di prepara-sito.js (nessun dato di un altro reparto),
  // e fra loro differiscono in un file solo, js/ambiente.js. È il senso di
  // «un codice, tante configurazioni».
  try {
    const { prepara } = require('./prepara-sito.js');
    const impronte = {};
    require('./reparti.js').tutti().forEach((rep, i) => {
      const r = prepara(path.join(RADICE, '_sito', 'controlli-' + rep.codice), { compressa: true, reparto: rep.codice });
      const c = r.compressione;
      impronte[rep.codice] = Object.fromEntries(c.righe.map((x) => [x.file, x.impronta]));
      if (i === 0) ok('copia compressa: ' + r.elenco.length + ' file, da ' + Math.round(c.prima / 1024) + ' a ' + Math.round(c.dopo / 1024) + ' KB; script che compilano, struttura delle pagine e ' + c.nomiGlobali + ' nomi globali identici, nessun indirizzo nuovo, segnalibri interi, contenuto di js/ambiente.js intatto');
    });
    const codici = Object.keys(impronte), diversi = new Set();
    codici.slice(1).forEach((c) => Object.keys(impronte[codici[0]]).concat(Object.keys(impronte[c])).forEach((p) => { if (impronte[codici[0]][p] !== impronte[c][p]) diversi.add(p); }));
    (codici.length < 2 || (diversi.size === 1 && diversi.has('js/ambiente.js')))
      ? ok('sito per reparto: ' + codici.length + ' siti costruiti (' + codici.join(', ') + '), nessuno porta dati di un altro reparto, e differiscono SOLO in js/ambiente.js')
      : ko('sito per reparto: i siti dei reparti devono differire solo in js/ambiente.js; differiscono in: ' + ([...diversi].join(', ') || 'nessun file'));
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

  const guai = [];
  // I controlli guardano il codice VIVO: una regola o una chiamata rimasta solo
  // dentro un commento non c'è (spenta «per una prova» e dimenticata, passava
  // da «ok»). I commenti del codice li trova l'analizzatore usato per la copia
  // compressa, non una ricerca nel testo: «//» sta anche dentro le stringhe.
  function senzaCommenti(js, nome) {
    let acorn = null;
    try { acorn = require(path.join(RADICE, '.github/compressione/node_modules/acorn')); }
    catch (e) { guai.push('per leggere ' + nome + ' senza i commenti serve lo strumento della copia compressa: «npm ci --ignore-scripts» dentro .github/compressione'); return js; }
    const tratti = [];
    try { acorn.parse(js, { ecmaVersion: 'latest', sourceType: 'script', onComment: (blocco, testo, da, a) => { tratti.push([da, a]); } }); }
    catch (e) { guai.push(nome + ' non si analizza: ' + e.message); return js; }
    let vivo = '', i = 0;
    tratti.forEach((t) => { vivo += js.slice(i, t[0]) + ' '; i = t[1]; });
    return vivo + js.slice(i);
  }
  const stile = leggi('docs/rubrica/app.css');
  const pagina = leggi('docs/rubrica/index.html').replace(/<!--[\s\S]*?-->/g, '');
  const codice = senzaCommenti(leggi('docs/rubrica/app.js'), 'docs/rubrica/app.js');
  [
    [/serviceWorker/, 'un service worker'],
    [/https?:\/\/[A-Za-z0-9]/, 'un indirizzo di rete (indirizzo e chiave si chiedono alla pagina madre)'],
    [/eyJ[A-Za-z0-9_-]{10,}/, 'una chiave'],
    [/createClient|\.signOut\b/, 'un client del database proprio o un\'uscita dalla sessione'],
    [/sessionStorage|indexedDB|\bcaches\b/, 'un\'altra memoria del browser'],
    [/XMLHttpRequest|sendBeacon|WebSocket|EventSource|importScripts|\bimport\s*\(/, 'una strada di rete diversa da quella di supaRisposta'],
    [/\bstyle\s*=|\son[a-z]+\s*=/, 'uno stile o un gestore scritto in linea nell\'HTML che costruisce (la regola CSP lo spegnerebbe)'],
    [/\beval\s*\(|new Function|document\.write|insertAdjacentHTML|outerHTML/, 'una costruzione che esegue o inserisce testo senza controllo'],
  ].forEach((v) => { if (v[0].test(codice)) guai.push('app.js contiene ' + v[1]); });
  const quanti = (re) => (codice.match(re) || []).length;
  // Richieste: la parola «fetch» compare una volta sola, nella chiamata di
  // supaRisposta (darle un secondo nome sarebbe una seconda occorrenza).
  if (quanti(/\bfetch\b/g) !== 1 || !/\bawait fetch\s*\(/.test(codice) || !/function supaRisposta\(/.test(codice)) guai.push('app.js: le richieste devono partire tutte da supaRisposta (la parola «fetch» compare ' + quanti(/\bfetch\b/g) + ' volte, attesa 1)');
  // Memoria del browser: ci si arriva solo dalle tre funzioni lsLeggi, lsScrivi
  // e lsTogli, una chiamata ciascuna…
  const dalleTre = ['getItem', 'setItem', 'removeItem'].every((m) => quanti(new RegExp('window\\.localStorage\\.' + m + '\\s*\\(', 'g')) === 1);
  if (quanti(/\blocalStorage\b/g) !== 3 || quanti(/\.setItem\s*\(/g) !== 1 || !dalleTre || !/function lsScrivi\(/.test(codice)) guai.push('app.js: alla memoria del browser si arriva solo da lsLeggi, lsScrivi e lsTogli (localStorage compare ' + quanti(/\blocalStorage\b/g) + ' volte, attese 3; setItem ' + quanti(/\.setItem\s*\(/g) + ', atteso 1)');
  // … e ci si scrivono solo le tre preferenze senza dati (tema, preferiti,
  // chiamati di recente), sotto chiavi «rubrica-…» scritte per esteso: sono
  // quelle che «Esci» conosce.
  const CHIAVI = ['LS_TEMA', 'LS_PREFERITI', 'LS_CHIAMATI'];
  const scritte = (codice.match(/\blsScrivi\s*\(\s*[^,)]*/g) || []).map((c) => c.replace(/^lsScrivi\s*\(\s*/, '').trim()).filter((a) => a !== 'chiave');
  const altre = scritte.filter((a) => CHIAVI.indexOf(a) < 0);
  if (!scritte.length || altre.length) guai.push('app.js: lsScrivi deve ricevere solo ' + CHIAVI.join(', ') + ' (trovato: ' + (altre.join(', ') || 'nessuna scrittura') + ')');
  CHIAVI.forEach((k) => { if (!new RegExp('const ' + k + '\\s*=\\s*\'rubrica-[a-z-]+\';').test(codice)) guai.push('app.js: ' + k + ' deve essere una chiave «rubrica-…» scritta per esteso'); });
  if (!/window\.parent/.test(codice)) guai.push('app.js non chiede più nulla alla pagina madre');

  const csp = (pagina.match(/<meta\s+http-equiv="Content-Security-Policy"\s+content="([^"]*)"/i) || [])[1];
  if (!csp) guai.push('index.html: manca la regola CSP (una regola messa fra segni di commento non vale)');
  else {
    const regole = {};
    csp.split(';').map((r) => r.trim()).filter(Boolean).forEach((r) => { const p = r.split(/\s+/); regole[p[0]] = p.slice(1).join(' '); });
    const attese = { 'default-src': "'none'", 'script-src': "'self'", 'style-src': "'self'", 'base-uri': "'none'", 'form-action': "'none'" };
    Object.keys(attese).forEach((k) => { if (regole[k] !== attese[k]) guai.push('index.html, regola CSP: «' + k + '» vale «' + (regole[k] || 'assente') + '», atteso ' + attese[k]); });
    if (regole['connect-src'] !== 'https://*.supabase.co') guai.push('index.html, regola CSP: «connect-src» deve ammettere solo il database (https://*.supabase.co), trovato «' + (regole['connect-src'] || 'assente') + '»');
    const larghe = Object.keys(regole).filter((k) => k !== 'connect-src').map((k) => regole[k]).join(' ');
    if (/unsafe|[*]|data:|blob:|https?:/.test(larghe)) guai.push('index.html, regola CSP: una voce larga (unsafe, *, data:, blob: o un sito esterno)');
    if (pagina.indexOf('http-equiv="Content-Security-Policy"') > pagina.search(/<(?:script|link)\b/i)) guai.push('index.html: la regola CSP deve venire prima di ogni script e di ogni foglio di stile');
  }
  if (/<script(?![^>]*\bsrc=)/i.test(pagina)) guai.push('index.html: uno script scritto in pagina');
  if (/<style\b|\sstyle\s*=/i.test(pagina)) guai.push('index.html: uno stile scritto in pagina');
  if (/<[^>]*\son[a-z]+\s*=/i.test(pagina)) guai.push('index.html: un gestore scritto in linea');
  if (/manifest|<iframe|<object|<embed|<base\b/i.test(pagina)) guai.push('index.html: manifesto, cornice o elemento non ammesso');
  const script = tagSrc(pagina).map((t) => t.src);
  if (script.join() !== 'app.js') guai.push('index.html: atteso un solo script, app.js (trovati: ' + (script.join(', ') || 'nessuno') + ')');
  if (/url\s*\(|@import|https?:/i.test(stile)) guai.push('app.css: un indirizzo o un\'importazione');

  // lato app: il riquadro apre la pagina dell'app, senza «sandbox», e l'uscita la svuota (anche qui conta il codice vivo)
  const p0 = ora.indexOf("var panel = document.getElementById('rubricaPanel');");
  const riquadro = p0 < 0 ? '' : senzaCommenti(ora.slice(ora.lastIndexOf('<script>', p0) + '<script>'.length, ora.indexOf('</script>', p0)), 'lo script del riquadro della rubrica in docs/index.html');
  const indirizzi = riquadro.match(/\bRUBRICA_URL\s*=[^=][^;]*;/g) || [];
  if (indirizzi.length !== 1 || !/^RUBRICA_URL\s*=\s*'rubrica\/';$/.test(indirizzi[0]) || /rubrica-gemelli/.test(ora)) guai.push('docs/index.html: il riquadro deve aprire la pagina dell\'app («rubrica/»), non un sito esterno');
  const alCornice = (riquadro.match(/\biframe\.src\s*=[^=][^;]*;/g) || []).filter((a) => !/^iframe\.src\s*=\s*(RUBRICA_URL|'about:blank');$/.test(a));
  if (alCornice.length || /setAttribute\(\s*['"]src/.test(riquadro)) guai.push('docs/index.html: alla cornice della rubrica si dà un indirizzo che non è RUBRICA_URL');
  if (/id="rubricaIframe"[^>]*\bsandbox\b/.test(ora.replace(/\s+/g, ' '))) guai.push('docs/index.html: la cornice della rubrica non deve avere «sandbox» (la rubrica chiede la sessione alla pagina)');
  const u0 = ora.indexOf('function _eseguiUscita() {'), u1 = ora.indexOf('window._esciSessione = function()', u0);
  const uscita = (u0 >= 0 && u1 > u0) ? senzaCommenti(ora.slice(u0, u1), 'la funzione _eseguiUscita di docs/index.html') : '';
  if (!/window\._rubricaSvuotaCopie\(\)/.test(uscita)) guai.push('docs/index.html: l\'uscita (_eseguiUscita) non svuota più la rubrica');
  // Prima di uscire l'app chiede alla rubrica se ha un lavoro a metà: la
  // chiamata deve stare nel codice vivo, e _lavoroInSospeso deve usarla
  // davvero (le due funzioni stanno subito prima di _eseguiUscita).
  const s0 = ora.indexOf('function _rubricaInSospeso() {');
  const sospeso = (s0 >= 0 && u0 > s0) ? senzaCommenti(ora.slice(s0, u0), 'le funzioni del lavoro in sospeso di docs/index.html') : '';
  if (!/\.rubricaInSospeso\(\)/.test(sospeso) || !/_schedeInSospeso\(\)\s*\|\|\s*_rubricaInSospeso\(\)/.test(sospeso) || !/function rubricaInSospeso\(/.test(codice)) guai.push('l\'uscita non chiede più alla rubrica se ha un lavoro in sospeso (rubricaInSospeso)');
  guai.length ? guai.forEach((g) => ko('rubrica: ' + g)) : ok('rubrica: nessun service worker, indirizzo o chiave nel codice vivo; regola CSP stretta e niente in linea; richieste e memoria da un punto solo, solo le tre preferenze; il riquadro apre «rubrica/» e l\'uscita la svuota');
}

// ── 14. reparti e strumenti ──────────────────────────────────────────
// I file dei reparti sono la fonte di indirizzi, database e chiavi pubbliche,
// ma gli strumenti nati con due ambienti portano ancora le loro copie: devono
// dire le stesse cose. Il codice di un reparto vale da nome d'ambiente
// («gemelli-11n» è la produzione) e passa dalle stesse barriere.
{
  compila('collaudo/strumenti/reparti.js');
  compila('collaudo/strumenti/banco.js');
  compila('collaudo/strumenti/banco-sw.js');
  try {
    const reparti = require('./reparti.js'), sb = require('./sb.js'), gh = require('./gh-api.js');
    const c = reparti.risolvi('collaudo'), p = reparti.risolvi('produzione');
    const guai = [];
    if (c.tipo !== 'collaudo' || p.tipo !== 'produzione') guai.push('i nomi «collaudo» e «produzione» non indicano reparti di quel tipo');
    if (c.supabaseUrl !== 'https://' + sb.REF_COLLAUDO + '.supabase.co') guai.push('sb.js, REF_COLLAUDO');
    if (p.supabaseUrl !== 'https://' + sb.REF_PRODUZIONE + '.supabase.co') guai.push('sb.js, REF_PRODUZIONE');
    const aj = JSON.parse(leggi('collaudo/strumenti/ambiente.json'));
    if (aj.url !== c.supabaseUrl || aj.anon !== c.supabaseAnonKey || aj.ref !== sb.REF_COLLAUDO) guai.push('collaudo/strumenti/ambiente.json');
    if (gh.AMBIENTI.collaudo.sitoNuovo !== 'https://' + c.pubblicazione.cloudflare.progetto + '.pages.dev/') guai.push('gh-api.js, sito nuovo del collaudo');
    if (gh.AMBIENTI.produzione.sitoNuovo !== 'https://' + p.pubblicazione.cloudflare.progetto + '.pages.dev/') guai.push('gh-api.js, sito nuovo della produzione');
    if (p.pubblicazione.githubPages.repository.split('/')[0] !== gh.AMBIENTI.produzione.utente) guai.push('gh-api.js, account della produzione');
    if (leggi('collaudo/pagina/prove.js').indexOf("var REF_PRODUZIONE = '" + sb.REF_PRODUZIONE + "';") < 0) guai.push('collaudo/pagina/prove.js, REF_PRODUZIONE');
    guai.length ? ko('i dati dei due ambienti scritti negli strumenti non concordano con i file dei reparti: ' + guai.join('; ')) : ok('strumenti: i dati dei due ambienti in sb.js, ambiente.json, gh-api.js e prove.js concordano con i file dei reparti');

    // nomi: il codice di un reparto è tradotto nel nome storico, e le barriere restano
    const nomi = [['gemelli-11n', 'produzione'], ['GEMELLI-11N', 'produzione'], ['produzione', 'produzione'], ['collaudo', 'collaudo'], ['inesistente', null], ['constructor', null], ['', null], ['--confermo-produzione', null]];
    const storti = nomi.filter((x) => reparti.nomeStorico(x[0]) !== x[1]).map((x) => '«' + x[0] + '» → ' + reparti.nomeStorico(x[0]));
    const rifiuta = (f) => { try { f(); return false; } catch (e) { return true; } };
    const a = gh.argomenti(['gemelli-11n', 'stato']);
    if (a.ambiente !== 'produzione' || !a.esplicito || a.resto.join() !== 'stato') storti.push('gh-api.js non legge «gemelli-11n» come produzione');
    if (!rifiuta(() => gh.argomenti(['stato', 'gemelli-11n']))) storti.push('gh-api.js accetta il nome del reparto fuori posto');
    if (!rifiuta(() => gh.argomenti(['collaudo', gh.CONFERMA]))) storti.push('gh-api.js accetta la conferma di produzione sul collaudo');
    if (gh.argomenti(['gemelli-11n']).confermato || !gh.per('gemelli-11n').produzione || gh.per('collaudo').produzione) storti.push('gh-api.js: «gemelli-11n» non passa dalle barriere della produzione');
    if (!rifiuta(() => gh.per('inesistente'))) storti.push('gh-api.js accetta un reparto sconosciuto');
    storti.length ? ko('nomi dei reparti negli strumenti: ' + storti.join('; ')) : ok('nomi dei reparti: «gemelli-11n» vale «produzione» negli strumenti, con la stessa conferma e la stessa sola lettura; un nome sconosciuto o fuori posto è rifiutato');

    // la pubblicazione riparte quando cambia un reparto, e chi la segue lo sa
    const flusso = leggi('.github/workflows/pubblica-cloudflare.yml'), passi = leggi('.github/pubblicazione/passi.js');
    const manca = [];
    if (!/- 'reparti\/\*\*'/.test(flusso) || !/- 'collaudo\/strumenti\/reparti[.]js'/.test(flusso)) manca.push('pubblica-cloudflare.yml non riparte se cambia reparti/ o reparti.js');
    ['collaudo/strumenti/pubblicazione.js', 'collaudo/strumenti/scambio-repository.js'].forEach((s) => { if (!/DEL_SITO = \/\^\(docs\\\/\|reparti\\\/\|/.test(leggi(s))) manca.push(s + ': DEL_SITO non conta reparti/'); });
    if (!/repartoDelFlusso\(\)/.test(passi) || !/reparto: reparto \? reparto[.]codice : undefined/.test(passi) || !/hostAmmesso\(reparto, p[.]subdomain\)/.test(passi)) manca.push('passi.js non prepara il sito per il reparto del progetto, o non controlla il suo indirizzo');
    if (!/reparto[.]supabaseUrl\) errore\(/.test(passi)) manca.push('passi.js non confronta il database dell\'avviso con quello del reparto');
    manca.length ? ko('pubblicazione per reparto: ' + manca.join('; ')) : ok('pubblicazione: il flusso riparte se cambia un reparto, prepara il sito del reparto del suo progetto, ne controlla l\'indirizzo e avvisa solo i suoi PC');

    // la cartella di destinazione viene svuotata: prepara-sito.js accetta solo _sito
    const fuori = path.join(RADICE, '_prova-fuori-da-sito-' + process.pid);
    const rifiutata = rifiuta(() => require('./prepara-sito.js').prepara(fuori, {}));
    fs.rmSync(fuori, { recursive: true, force: true });
    rifiutata ? ok('prepara-sito.js rifiuta una cartella di destinazione fuori da _sito (la svuoterebbe)') : ko('prepara-sito.js ha accettato una cartella di destinazione fuori da _sito');
  } catch (e) { ko('reparti e strumenti: ' + String((e && e.message) || e).split('\n').join(' | ').slice(0, 600)); }
}

// ── 16. modo rubrica ─────────────────────────────────────────────────
// La pagina aperta con «?rubrica=1» mostra, dopo l'accesso, la sola rubrica:
// i dati dei pazienti non devono essere né scaricati né disegnati. Qui si
// sorveglia nel codice ciò che lo garantisce; il comportamento lo provano
// prove.js (sezione rubrica) e prove-sw.js.
{
  const guai = [];
  try {
    const api = leggi('docs/js/api.js'), app = leggi('docs/js/app.js'), stile = leggi('docs/css/styles.css'), rub = leggi('docs/rubrica/app.js');
    const { SEGNO_RUBRICA } = require('./rinvio-rubrica.js');
    // il segno: uno solo, letto in un punto solo e con un confronto esatto
    if (SEGNO_RUBRICA !== '?rubrica=1') guai.push('il segno del rinvio non è «?rubrica=1»');
    const iTesta = ora.indexOf("modo = new URLSearchParams(window.location.search).get('rubrica') === '1';");
    if (iTesta < 0 || ora.indexOf('window._MODO_RUBRICA = modo;') < iTesta) guai.push('index.html: lo script in testa non legge il segno con un confronto esatto');
    else if (iTesta > ora.indexOf('<link href="https://cdn.jsdelivr.net') || iTesta > ora.indexOf('js/api.js')) guai.push('index.html: il modo va deciso in testa, prima dei fogli di stile e di ogni altro script');
    if ((ora.match(/_MODO_RUBRICA\s*=[^=]/g) || []).length !== 1 || /_MODO_RUBRICA\s*=[^=]/.test(api + app + rub)) guai.push('il modo si assegna in un punto solo, lo script in testa a index.html');
    // il secondo manifesto riapre lo stesso indirizzo, senza identità propria né indirizzi assoluti
    const m = JSON.parse(leggi('docs/manifest-rubrica.json')), c = JSON.parse(leggi('docs/manifest.json'));
    if (m.start_url !== './' + SEGNO_RUBRICA || m.scope !== './' || m.id !== undefined || /https?:|\/\//.test(JSON.stringify(m))) guai.push('manifest-rubrica.json: attesi start_url «./?rubrica=1», scope «./», nessun id e nessun indirizzo assoluto');
    if (c.start_url !== './' || c.id !== undefined || c.scope !== undefined) guai.push('manifest.json delle consegne è cambiato: le icone già installate perderebbero la loro app');
    if (ora.indexOf("imposta('link[rel=\"manifest\"]', 'href', 'manifest-rubrica.json');") < 0) guai.push('index.html: in modo rubrica il manifesto non viene scambiato');
    // l'app parte da un punto solo
    if ((ora.match(/window\._avviaFlussoApp\(\)/g) || []).length !== 1 || !/function _avviaDopoAccesso\(\) \{\s*if \(window\._MODO_RUBRICA === true\) \{ _avviaModoRubrica\(\); return; \}/.test(ora)) guai.push('index.html: l\'app deve partire da un punto solo, _avviaDopoAccesso, che in modo rubrica apre la rubrica');
    if ((ora.match(/_avviaDopoAccesso\(\);/g) || []).length !== 2) guai.push('index.html: attese due chiamate a _avviaDopoAccesso (sessione ritrovata, accesso appena fatto)');
    // le strade che leggono da sole i dati dei pazienti: tutte chiuse in modo rubrica
    const COMMENTI = '(?:\\s*\\/\\/[^\\n]*\\n)*\\s*';
    [[app, 'window\\._avviaFlussoApp = function\\(\\) \\{' + COMMENTI + 'if \\(window\\._MODO_RUBRICA === true\\) return;', 'app.js, _avviaFlussoApp'],
      [api, 'function _inizializzaRealtime\\(\\) \\{' + COMMENTI + 'if \\(window\\._MODO_RUBRICA === true\\) return;', 'api.js, _inizializzaRealtime'],
      [api, 'function _rtMiniResync\\(\\) \\{' + COMMENTI + 'if \\(window\\._MODO_RUBRICA === true\\) return;', 'api.js, _rtMiniResync'],
      [api, 'function _emergenzaAvviaPolling\\(\\) \\{' + COMMENTI + 'if \\(window\\._MODO_RUBRICA === true\\) return;', 'api.js, _emergenzaAvviaPolling'],
      [api, 'function _emergenzaCheckSeMutato\\(\\) \\{' + COMMENTI + 'if \\(window\\._MODO_RUBRICA === true\\) return Promise\\.resolve\\(false\\);', 'api.js, _emergenzaCheckSeMutato'],
      [api, "addEventListener\\('visibilitychange', function\\(\\) \\{\\s*if \\(document\\.visibilityState !== 'visible'\\) return;" + COMMENTI + 'if \\(window\\._MODO_RUBRICA === true\\) return;', 'api.js, ritorno in primo piano'],
      [ora, 'function _maybeDiagAuto\\(\\) \\{' + COMMENTI + 'if \\(window\\._MODO_RUBRICA === true\\) return;', 'index.html, diagnostica automatica'],
    ].forEach((x) => { if (!new RegExp(x[1]).test(x[0])) guai.push(x[2] + ': manca in testa la guardia del modo rubrica (lì partirebbe la lettura dei dati dei pazienti)'); });
    // il segno sopravvive al ricaricamento di «Update»
    const f = /function _indirizzoDiRicarica\([^)]*\) \{[\s\S]*?\n\}/.exec(api.replace(/\r\n/g, '\n'));
    let ricarica = null;
    try { ricarica = f && new Function(f[0] + '\nreturn _indirizzoDiRicarica;')(); } catch (e) { ricarica = null; }
    if (!ricarica || ricarica('/x/', '#a', 5, true) !== '/x/?_r=5&rubrica=1#a' || ricarica('/x/', '', 5, false) !== '/x/?_r=5' || api.indexOf("_indirizzoDiRicarica(window.location.pathname, window.location.hash, Date.now(), window._MODO_RUBRICA === true)") < 0) guai.push('api.js: il ricaricamento di «Update» non conserva il segno del modo rubrica');
    // ciò che si vede: solo la rubrica, a tutto schermo
    ['html.modo-rubrica #appBodyContent,', 'html.modo-rubrica #backupCheckOverlay,', 'html.modo-rubrica #versionUpdateOverlay { display: none !important; }', 'html.modo-rubrica #rubricaPanelHeader { display: none !important; }'].forEach((r) => { if (stile.indexOf(r) < 0) guai.push('styles.css: manca la regola «' + r + '»'); });
    // l'icona di uscita: solo in modo rubrica, e l'uscita la fa la pagina madre
    if (!/madreDelModo\._MODO_RUBRICA === true && typeof madreDelModo\._esciSessione === 'function'/.test(rub) || !/id="btnEsci"[^>]*\bhidden\b/.test(leggi('docs/rubrica/index.html'))) guai.push('rubrica: l\'icona di uscita deve nascere nascosta e comparire solo in modo rubrica');
    // il collegamento all'app intera nasce nascosto, e in modo rubrica resta tale
    if (!/class="avviso-link"[^>]*\bhidden\b/.test(leggi('docs/rubrica/index.html')) || rub.indexOf('collegamento.hidden = inModoRubrica;') < 0) guai.push('rubrica: «Apri le Consegne» deve nascere nascosto e comparire solo fuori dal modo rubrica');
    // il riquadro lì non si richiude mai: l'elenco si ricontrolla al ritorno in primo piano
    if (ora.indexOf("document.addEventListener('visibilitychange', _modoRubricaAlRitorno);") < 0 || !/function _modoRubricaAlRitorno\(\) \{\s*if \(window\._MODO_RUBRICA !== true \|\| !_modoRubricaAvviato \|\| _uscitaInCorso /.test(ora)) guai.push('index.html: in modo rubrica l\'elenco deve essere ricontrollato quando la pagina torna in primo piano (_modoRubricaAlRitorno)');
  } catch (e) { guai.push(String((e && e.message) || e).slice(0, 300)); }
  guai.length ? guai.forEach((g) => ko('modo rubrica: ' + g)) : ok('modo rubrica: segno unico «?rubrica=1» deciso in testa alla pagina; l\'app parte da un punto solo e in modo rubrica sono chiuse le 7 strade verso i dati dei pazienti; il segno sopravvive a «Update»; secondo manifesto e icona di uscita al loro posto');
}

// ── 15. ciò che non deve diventare pubblico ──────────────────────────
// Il repository del reparto è pubblico. Il piano, riservato, sta su un ramo a
// parte («piano») del repository privato di collaudo: non deve comparire né
// fra i file registrati né in un commit che un rilascio porterebbe su origin.
// Una volta pubblicato non si richiama.
{
  try {
    const git = (a) => execFileSync('git', ['-C', RADICE].concat(a), { encoding: 'utf8', maxBuffer: 64 * 1024 * 1024, stdio: ['ignore', 'pipe', 'ignore'] });
    const fraIFile = git(['ls-files']).split('\n').filter((f) => /PIANO-REPARTI/i.test(f));
    let nellaStoria = [];
    try { nellaStoria = git(['rev-list', '--objects', 'HEAD', '--not', RIF]).split('\n').filter((r) => /PIANO-REPARTI/i.test(r)); }
    catch (e) { nellaStoria = ['storia non leggibile rispetto a ' + RIF]; }
    (fraIFile.length || nellaStoria.length)
      ? ko('riservatezza: il piano comparirebbe nel repository pubblico (' + (fraIFile.length ? 'fra i file: ' + fraIFile.join(', ') : 'nella storia da rilasciare: ' + nellaStoria.length + ' oggetti') + '). Sta nel ramo «piano» del repository privato di collaudo e non va rilasciato')
      : ok('riservatezza: il piano non sta né fra i file né nella storia che un rilascio porterebbe nel repository pubblico');
  } catch (e) { ko('riservatezza: ' + String((e && e.message) || e).slice(0, 300)); }
}

console.log(errori ? ('\n' + errori + ' CONTROLLI FALLITI: non pubblicare') : '\ntutti i controlli superati');
process.exit(errori ? 1 : 0);
