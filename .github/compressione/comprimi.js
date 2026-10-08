// Comprime la copia del sito già preparata in una cartella (prepara-sito.js).
// È la copia che Cloudflare pubblica: lo stesso programma, più scomodo da
// leggere e da copiare. docs/ resta com'è, leggibile, ed è ciò su cui si lavora.
//
//  - file .js del sito (tranne js/librerie/, già compressi e protetti da
//    «integrity»): terser toglie commenti e spazi e accorcia i nomi LOCALI;
//  - pagine .html: lo stesso per gli script scritti in pagina; dei commenti
//    HTML viene tolto il testo e resta il segno vuoto;
//  - .css e blocchi <style>: viene tolto il testo dei commenti.
//
// Che cosa NON cambia mai: i nomi visibili da fuori (funzioni e variabili
// globali, proprietà, attributi), perché pagine, gestori scritti negli
// attributi e file diversi si chiamano fra loro per nome; la struttura delle
// pagine; i file delle librerie. Nessuna trasformazione «unsafe».
//
// Alla fine la copia viene controllata (vedi «verifica»): se un controllo
// fallisce la compressione è un errore e non si pubblica nulla.
//
// Lo strumento ha versione fissa (package.json e impronte in package-lock.json)
// e si installa con «npm ci --ignore-scripts» in questa cartella.
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const crypto = require('crypto');

function strumenti() {
  try { return { terser: require('terser'), acorn: require('acorn') }; } catch (e) {
    throw new Error('lo strumento di compressione non è installato: lanciare «npm ci --ignore-scripts» in .github/compressione');
  }
}

const OPZIONI = {
  compress: { passes: 1 },
  mangle: { safari10: true },
  format: { comments: false, inline_script: true },
  toplevel: false,
  module: false,
};

const SEGNALIBRI = ['_bmTerapiaSorgente', '_bmLabSorgente'];
const SCRIPT_IN_PAGINA = /<script(?![^>]*\bsrc=)([^>]*)>([\s\S]*?)<\/script>/gi;
const eJs = (attributi) => { const m = /\btype\s*=\s*["']?([^"'\s>]+)/i.exec(attributi); return !m || /^(text|application)\/javascript$/i.test(m[1]); };
const impronta = (b) => crypto.createHash('sha256').update(b).digest('hex');

function comprimiJs(codice, nome) {
  const r = strumenti().terser.minify_sync(codice, OPZIONI);
  if (r.error) throw new Error(nome + ': ' + r.error.message);
  if (typeof r.code !== 'string' || !r.code.length) throw new Error(nome + ': la compressione non ha prodotto codice');
  return r.code;
}

// Toglie il testo dei commenti CSS lasciando il segno vuoto: i pezzi ai due
// lati restano separati esattamente come prima.
function cssSenzaCommenti(css) {
  let fuori = '', i = 0, tolti = 0;
  while (i < css.length) {
    const c = css[i];
    if (c === '"' || c === "'") {
      let j = i + 1;
      while (j < css.length && css[j] !== c) { if (css[j] === '\\') j++; j++; }
      fuori += css.slice(i, j + 1); i = j + 1; continue;
    }
    if (c === '/' && css[i + 1] === '*') {
      const fine = css.indexOf('*/', i + 2);
      if (fine < 0) { fuori += css.slice(i); break; }
      if (fine > i + 2) tolti++;
      fuori += '/**/'; i = fine + 2; continue;
    }
    fuori += c; i++;
  }
  return { testo: fuori, tolti };
}

// La pagina a pezzi: script, stili e zone di solo testo non si toccano nel
// cercare i commenti HTML.
const BLOCCHI = /<(script|style|textarea|title)\b[^>]*>[\s\S]*?<\/\1>/gi;
function fuoriDaiBlocchi(html, f) {
  let fuori = '', ultimo = 0, m;
  BLOCCHI.lastIndex = 0;
  while ((m = BLOCCHI.exec(html))) { fuori += f(html.slice(ultimo, m.index)) + m[0]; ultimo = m.index + m[0].length; }
  return fuori + f(html.slice(ultimo));
}

function scriptInPagina(html) {
  const fuori = []; let m;
  SCRIPT_IN_PAGINA.lastIndex = 0;
  while ((m = SCRIPT_IN_PAGINA.exec(html))) fuori.push({ attributi: m[1], codice: m[2] });
  return fuori;
}

function comprimiHtml(html, nome) {
  // Un commento che racchiude un blocco (script, stile…) confonderebbe il taglio a pezzi.
  const annidati = (html.match(/<!--[\s\S]*?-->/g) || []).filter((c) => /<(script|style|textarea|title)\b/i.test(c));
  if (annidati.length) throw new Error(nome + ': un commento HTML contiene un blocco <script>, <style>, <textarea> o <title>: va tolto dal sorgente');
  const esito = { script: 0, commenti: 0, stili: 0 };
  let n = 0;
  let pagina = html.replace(SCRIPT_IN_PAGINA, (tutto, attributi, codice) => {
    n++;
    if (!eJs(attributi) || !codice.trim()) return tutto;
    const compresso = comprimiJs(codice, nome + ', script n. ' + n);
    if (/<\/script/i.test(compresso)) throw new Error(nome + ', script n. ' + n + ': il codice compresso contiene «</script»');
    esito.script++;
    return '<script' + attributi + '>' + compresso + '</script>';
  });
  pagina = pagina.replace(/(<style\b[^>]*>)([\s\S]*?)(<\/style>)/gi, (tutto, apre, css, chiude) => {
    const r = cssSenzaCommenti(css); esito.stili += r.tolti; return apre + r.testo + chiude;
  });
  pagina = fuoriDaiBlocchi(pagina, (pezzo) => pezzo.replace(/<!--[\s\S]*?-->/g, (c) => { if (c.length > 7) esito.commenti++; return '<!---->'; }));
  return { pagina, esito };
}

// ── Verifica della copia ───────────────────────────────────────────────
// Lo scheletro di una pagina: tutto tranne il contenuto di script, stili e commenti.
function scheletro(html) {
  return html
    .replace(SCRIPT_IN_PAGINA, (t, a) => '<script' + a + '></script>')
    .replace(/(<style\b[^>]*>)[\s\S]*?(<\/style>)/gi, '$1$2')
    .replace(/<!--[\s\S]*?-->/g, '<!---->');
}

// Analisi del codice con lo stesso analizzatore usato da terser.
function albero(codice) { return strumenti().acorn.parse(codice, { ecmaVersion: 'latest', sourceType: 'script' }); }
function visita(nodo, f) {
  if (!nodo || typeof nodo.type !== 'string') return;
  f(nodo);
  for (const k of Object.keys(nodo)) {
    const v = nodo[k];
    if (Array.isArray(v)) v.forEach((x) => visita(x, f)); else if (v && typeof v.type === 'string') visita(v, f);
  }
}
// Il testo di una funzione dichiarata per nome, così come lo restituirebbe toString().
function funzione(codice, nome) {
  let trovata = null;
  visita(albero(codice), (n) => { if (!trovata && n.type === 'FunctionDeclaration' && n.id && n.id.name === nome) trovata = codice.slice(n.start, n.end); });
  return trovata;
}
// I nomi dichiarati al livello più esterno: sono quelli visibili agli altri file e alle pagine.
function nomiGlobali(codice) {
  const nomi = new Set();
  const dalModello = (m) => {
    if (!m) return;
    if (m.type === 'Identifier') nomi.add(m.name);
    else if (m.type === 'ObjectPattern') m.properties.forEach((q) => dalModello(q.value || q.argument));
    else if (m.type === 'ArrayPattern') m.elements.forEach(dalModello);
    else if (m.type === 'AssignmentPattern') dalModello(m.left);
    else if (m.type === 'RestElement') dalModello(m.argument);
  };
  albero(codice).body.forEach((n) => {
    if ((n.type === 'FunctionDeclaration' || n.type === 'ClassDeclaration') && n.id) nomi.add(n.id.name);
    else if (n.type === 'VariableDeclaration') n.declarations.forEach((d) => dalModello(d.id));
  });
  return nomi;
}
function confrontaNomi(prima, dopo, dove, errori) {
  const a = nomiGlobali(prima), b = nomiGlobali(dopo);
  const persi = Array.from(a).filter((x) => !b.has(x)), nuovi = Array.from(b).filter((x) => !a.has(x));
  if (persi.length) errori.push(dove + ': nomi globali spariti nella copia compressa: ' + persi.slice(0, 8).join(', '));
  if (nuovi.length) errori.push(dove + ': nomi globali comparsi nella copia compressa: ' + nuovi.slice(0, 8).join(', '));
  return a.size;
}
const indirizzi = (testo) => Array.from(new Set((testo.match(/\b(?:https?|wss?):\/\/[A-Za-z0-9.-]+/g) || [])));

function verifica(cartella, originali, elenco) {
  const errori = [];
  const leggi = (p) => fs.readFileSync(path.join(cartella, p), 'utf8');
  const tuttoIlSorgente = Object.keys(originali).map((p) => originali[p]).join('\n');
  let globali = 0;

  elenco.filter((p) => /\.js$/.test(p)).forEach((p) => {
    try { new vm.Script(leggi(p), { filename: p }); } catch (e) { return errori.push(p + ' non compila: ' + e.message); }
    if (originali[p] != null) globali += confrontaNomi(originali[p], leggi(p), p, errori);
  });
  elenco.filter((p) => /\.html$/.test(p)).forEach((p) => {
    const dopo = leggi(p);
    const a = scriptInPagina(originali[p]), b = scriptInPagina(dopo);
    if (a.length !== b.length) return errori.push(p + ': numero di script in pagina cambiato (' + a.length + ' → ' + b.length + ')');
    a.forEach((x, i) => {
      if (!eJs(x.attributi) || !x.codice.trim()) return;
      try { new vm.Script(b[i].codice, { filename: p + '#' + (i + 1) }); } catch (e) { return errori.push(p + ', script n. ' + (i + 1) + ' non compila: ' + e.message); }
      globali += confrontaNomi(x.codice, b[i].codice, p + ', script n. ' + (i + 1), errori);
    });
    if (scheletro(dopo) !== scheletro(originali[p])) errori.push(p + ': la struttura della pagina è cambiata (deve cambiare solo il contenuto di script, stili e commenti)');
  });

  // Nessun indirizzo di rete che non sia già nel sorgente.
  elenco.filter((p) => originali[p] != null).forEach((p) => {
    indirizzi(leggi(p)).forEach((u) => { if (tuttoIlSorgente.indexOf(u) < 0) errori.push(p + ': nella copia compressa compare un indirizzo assente dal sorgente: ' + u); });
  });

  // I segnalibri di TrakCare nascono dal testo di due funzioni: devono restare
  // funzioni intere, compilabili, senza a capo e senza sequenze «%XX» (in un
  // indirizzo javascript: verrebbero decodificate e romperebbero il codice).
  if (originali['index.html'] != null) {
    const cerca = (html, nome) => {
      for (const s of scriptInPagina(html)) {
        if (eJs(s.attributi) && s.codice.indexOf('function ' + nome + '(') >= 0) { const f = funzione(s.codice, nome); if (f) return f; }
      }
      return null;
    };
    const pagina = leggi('index.html');
    SEGNALIBRI.forEach((nome) => {
      if (!cerca(originali['index.html'], nome)) return;
      const f = cerca(pagina, nome);
      if (!f) return errori.push('index.html: nella copia compressa non trovo più la funzione ' + nome);
      try { new vm.Script('(' + f + ')'); } catch (e) { errori.push(nome + ' compresso non compila da solo: ' + e.message); }
      if (/[\r\n]/.test(f)) errori.push(nome + ' compresso contiene a capo');
      const sospette = f.match(/%[0-9A-Fa-f]{2}/g);
      if (sospette) errori.push(nome + ' compresso contiene sequenze che un indirizzo javascript: decodificherebbe: ' + Array.from(new Set(sospette)).slice(0, 5).join(' '));
    });
  }

  // Service worker: stessa versione e stessi file tenuti in cache.
  if (originali['sw.js'] != null) {
    const prima = originali['sw.js'], dopo = leggi('sw.js');
    const v = (/consegne-v\d+/.exec(prima) || [])[0];
    if (!v || dopo.indexOf(v) < 0) errori.push('sw.js: nella copia compressa non trovo la versione ' + v);
    const lista = ((prima.match(/PRECACHE_ASSETS\s*=\s*\[([\s\S]*?)\]/) || [])[1] || '').match(/'[^']+'/g) || [];
    lista.map((x) => x.slice(1, -1)).forEach((x) => { if (dopo.indexOf(x) < 0) errori.push('sw.js: nella copia compressa manca il file in cache ' + x); });
  }
  return { errori, globali };
}

// Comprime sul posto i file dell'elenco dentro «cartella».
function comprimi(cartella, elenco) {
  const versione = 'terser ' + require('terser/package.json').version;
  strumenti();
  const originali = {};
  const righe = [];
  let prima = 0, dopo = 0;
  elenco.forEach((p) => {
    const file = path.join(cartella, p);
    const grezzo = fs.readFileSync(file);
    prima += grezzo.length;
    let nuovo = null, nota = 'copiato così com\'è';
    if (/^js\/librerie\//.test(p)) nota = 'libreria: non si tocca';
    else if (/\.js$/.test(p)) {
      originali[p] = grezzo.toString('utf8');
      nuovo = comprimiJs(originali[p], p); nota = 'compresso';
    } else if (/\.html$/.test(p)) {
      originali[p] = grezzo.toString('utf8');
      const r = comprimiHtml(originali[p], p);
      nuovo = r.pagina; nota = r.esito.script + ' script compressi, ' + r.esito.commenti + ' commenti svuotati';
    } else if (/\.css$/.test(p)) {
      originali[p] = grezzo.toString('utf8');
      const r = cssSenzaCommenti(originali[p]);
      nuovo = r.testo; nota = r.tolti + ' commenti svuotati';
    }
    if (nuovo != null) fs.writeFileSync(file, nuovo);
    const finale = fs.readFileSync(file);
    dopo += finale.length;
    righe.push({ file: p, prima: grezzo.length, dopo: finale.length, nota, impronta: impronta(finale) });
  });
  const v = verifica(cartella, originali, elenco);
  if (v.errori.length) throw new Error('la copia compressa non supera i controlli:\n  - ' + v.errori.join('\n  - '));
  return { righe, prima, dopo, strumento: versione, nomiGlobali: v.globali };
}

module.exports = { comprimi, comprimiJs, comprimiHtml, cssSenzaCommenti, verifica, OPZIONI };
