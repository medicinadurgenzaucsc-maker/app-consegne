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
['docs/index.html', 'docs/print.html'].forEach((f) => {
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
[['docs/index.html', ora, prima], ['docs/print.html', leggi('docs/print.html'), diRif('docs/print.html')]].forEach((x) => {
  const adesso = tagSrc(x[1]);
  const nomi = adesso.map((t) => t.src);
  if (x[2] != null) {
    const persi = tagSrc(x[2]).map((t) => t.src).filter((s) => nomi.indexOf(s) < 0);
    persi.length ? ko(x[0] + ': tag script presenti in ' + RIF + ' e ora ASSENTI: ' + persi.join(', ')) : ok(x[0] + ': ' + adesso.length + ' tag script con src, nessuno perso rispetto a ' + RIF);
  }
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
['./js/sanifica.js', './js/api.js', './js/app.js', './js/app2.js'].concat(tagSrc(ora).map((t) => t.src).filter((s) => /purify/.test(s)).map((s) => './' + s)).forEach((p) => {
  if (elenco.indexOf("'" + p + "'") < 0) ko('sw.js: ' + p + ' non è fra i file tenuti in cache');
});

console.log(errori ? ('\n' + errori + ' CONTROLLI FALLITI: non pubblicare') : '\ntutti i controlli superati');
process.exit(errori ? 1 : 0);
