// Prepara la cartella da pubblicare: una copia dei SOLI file del sito, per UN
// reparto. Sono i file di docs/ registrati in git, tolto tutto ciò che sta in
// una cartella dal nome che comincia con «_» o con «.» (come docs/_legacy): gli
// stessi che pubblica GitHub Pages. In più passano i due file di configurazione
// di Cloudflare, _headers e _redirects, se ci sono.
//
// Il codice è uno solo per tutti i reparti: di un reparto cambia soltanto
// js/ambiente.js, che qui viene GENERATO dal suo file reparti/<codice>.json
// (collaudo/strumenti/reparti.js). Senza «--reparto» il sito è quello del
// reparto che pubblica da docs/ così com'è (oggi la 11N). A copia fatta si
// controlla che nel sito non ci siano indirizzi, chiavi o database di un ALTRO
// reparto: se ce ne sono la preparazione fallisce.
//
// Senza altro i file sono copiati così come sono, byte per byte. Con
// «compressa» la copia passa poi da .github/compressione/comprimi.js: è la
// forma in cui il sito viene pubblicato su Cloudflare, e va provata in quella
// forma (banco.js e banco-sw.js sanno servirla).
//
// La cartella è sempre _sito o una sua sottocartella (viene svuotata prima).
//   node collaudo/strumenti/prepara-sito.js <cartella>                       prepara la cartella
//   node collaudo/strumenti/prepara-sito.js <cartella> compressa             la prepara e la comprime
//   node collaudo/strumenti/prepara-sito.js <cartella> … --reparto <codice>  per quel reparto
//   node collaudo/strumenti/prepara-sito.js                                  stampa solo l'elenco
//
// Usato dal flusso .github/workflows/pubblica-cloudflare.yml, dai controlli di
// rilascio e dai banchi di prova (banco.js, banco-sw.js).
const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');
const reparti = require('./reparti.js');

const RADICE = path.resolve(__dirname, '../..');
const CONFIGURAZIONE = ['_headers', '_redirects'];
const AMBIENTE = 'js/ambiente.js';
// Indirizzi di questo PC: li può elencare più di un reparto di prova, non dicono di chi è un sito.
const LOCALI = ['localhost', '127.0.0.1'];

// Percorsi relativi a docs/, con la barra «/».
function elencoSito(riferimento) {
  const args = riferimento ? ['ls-tree', '-r', '-z', '--name-only', riferimento, 'docs'] : ['ls-files', '-z', 'docs'];
  const tutti = execFileSync('git', ['-C', RADICE].concat(args), { encoding: 'utf8', maxBuffer: 16 * 1024 * 1024 })
    .split('\0').filter(Boolean).map((p) => p.replace(/^docs\//, ''));
  return tutti.filter((p) => {
    const pezzi = p.split('/');
    if (pezzi.length === 1 && CONFIGURAZIONE.indexOf(p) >= 0) return true;
    return !pezzi.some((x) => x.charAt(0) === '_' || x.charAt(0) === '.');
  }).sort();
}

// Ciò che in un sito dice «questo è il reparto X»: il suo database, la sua
// chiave pubblica, il suo client Google, i suoi indirizzi, dove pubblica.
function segniDi(reparto) {
  const segni = [
    // il solo riferimento del progetto (venti lettere): lo si trova anche dentro un altro indirizzo o in una chiave di memoria
    ['il database', reparto.supabaseUrl.replace(/^https:\/\//, '').replace(/[.]supabase[.]co$/, '')],
    ['la chiave pubblica', reparto.supabaseAnonKey],
    ['il client Google', reparto.googleClientId],
  ];
  reparto.host.forEach((h) => { if (LOCALI.indexOf(h) < 0) segni.push(["l'indirizzo", h.replace(/^\./, '')]); });
  const p = reparto.pubblicazione;
  if (p.cloudflare) segni.push(['il progetto Cloudflare', p.cloudflare.progetto + '.pages.dev']);
  if (p.githubPages) segni.push(['il repository', p.githubPages.repository]);
  return segni;
}

// TRANSITORIO (v190): js/api.js porta ancora la tabella _AMBIENTI della v189,
// con i dati dei due ambienti di allora, come riserva per i minuti di un
// aggiornamento (vedi il commento in testa ad api.js). È l'unico punto in cui
// un sito può nominare un altro reparto, e qui lo si salta: SOLO la tabella,
// il resto del file si controlla. Quando la tabella sparisce da api.js questa
// funzione non trova più nulla da togliere e il file si controlla per intero.
function senzaTabellaTransitoria(testo) {
  const i = testo.indexOf('var _AMBIENTI = {');
  if (i < 0) return testo;
  const j = testo.indexOf('\n};', i);
  if (j < 0) throw new Error('js/api.js: la tabella _AMBIENTI comincia ma non finisce');
  return testo.slice(0, i) + testo.slice(j + 3);
}

// Nel sito preparato per un reparto non deve comparire nulla che appartenga a
// un altro. Restituisce l'elenco dei difetti (vuoto = a posto): nomi di file e
// di cose, mai i valori.
function datiDiAltri(cartella, elenco, reparto) {
  const miei = segniDi(reparto).map((s) => s[1]);
  const altri = reparti.tutti().filter((r) => r.codice !== reparto.codice);
  const difetti = [];
  elenco.forEach((p) => {
    let testo = fs.readFileSync(path.join(cartella, p), 'latin1');
    if (p === 'js/api.js') testo = senzaTabellaTransitoria(testo.replace(/\r\n/g, '\n'));
    altri.forEach((r) => {
      segniDi(r).forEach((s) => {
        // un valore che i due reparti hanno in comune non dice di chi è il sito
        if (miei.indexOf(s[1]) >= 0) return;
        if (testo.indexOf(s[1]) >= 0) difetti.push(p + ' contiene ' + s[0] + ' del reparto «' + r.codice + '»');
      });
    });
  });
  return difetti;
}

function prepara(destinazione, opzioni) {
  opzioni = opzioni || {};
  const dest = path.resolve(destinazione);
  // La cartella di destinazione viene SVUOTATA prima della copia: per questo
  // può essere solo _sito o una sua sottocartella (fuori da git). Fino alla
  // v189 bastava che stesse nel repository: «prepara-sito.js collaudo», o il
  // codice di un reparto scritto al posto della cartella, avrebbe cancellato
  // quella cartella.
  const SITI = path.join(RADICE, '_sito');
  if (dest !== SITI && !dest.startsWith(SITI + path.sep)) {
    throw new Error('la cartella di destinazione viene svuotata: può essere solo _sito o una sua sottocartella, non ' + dest);
  }
  // Il reparto: quello chiesto, altrimenti quello che pubblica da docs/ così com'è. Mai un reparto «indovinato».
  const diDocs = reparti.delSitoInDocs();
  const reparto = (opzioni.reparto === undefined || opzioni.reparto === null) ? diDocs : reparti.risolvi(opzioni.reparto);
  const elenco = elencoSito();
  if (elenco.indexOf(AMBIENTE) < 0) throw new Error('fra i file del sito registrati in git manca ' + AMBIENTE);
  // docs/js/ambiente.js è un file generato: se qualcuno l'ha cambiato a mano ci si ferma prima di copiare
  const c = reparti.controllaDocs();
  if (!c.ok) throw new Error(c.motivo);

  fs.rmSync(dest, { recursive: true, force: true });
  elenco.forEach((p) => {
    const da = path.join(RADICE, 'docs', p), a = path.join(dest, p);
    if (!fs.existsSync(da)) throw new Error('file registrato in git ma assente: docs/' + p);
    fs.mkdirSync(path.dirname(a), { recursive: true });
    fs.copyFileSync(da, a);
  });
  // L'unico file che cambia da un reparto all'altro. Per il reparto di docs/ resta la copia identica appena fatta.
  if (reparto.codice !== diDocs.codice) fs.writeFileSync(path.join(dest, AMBIENTE), reparti.ambienteJs(reparto));
  const scritto = fs.readFileSync(path.join(dest, AMBIENTE), 'utf8').replace(/\r\n/g, '\n');
  if (scritto !== reparti.ambienteJs(reparto)) throw new Error(AMBIENTE + ' nella copia non è quello del reparto «' + reparto.codice + '»');

  const estranei = datiDiAltri(dest, elenco, reparto);
  if (estranei.length) throw new Error('il sito preparato per «' + reparto.codice + '» porta dati di un altro reparto:\n  ' + estranei.join('\n  '));

  let byte = 0;
  elenco.forEach((p) => { byte += fs.statSync(path.join(dest, p)).size; });
  if (!opzioni.compressa) return { elenco, byte, reparto: reparto.codice };
  const esito = require(path.join(RADICE, '.github', 'compressione', 'comprimi.js')).comprimi(dest, elenco);
  return { elenco, byte: esito.dopo, compressione: esito, reparto: reparto.codice };
}

module.exports = { elencoSito, prepara, datiDiAltri, senzaTabellaTransitoria, RADICE, AMBIENTE };

if (require.main === module) {
  try {
    const args = process.argv.slice(2);
    let reparto;
    const i = args.indexOf('--reparto');
    if (i >= 0) {
      reparto = args[i + 1];
      if (!reparto || reparto.indexOf('--') === 0) throw new Error('dopo --reparto va scritto il codice del reparto');
      args.splice(i, 2);
    }
    const dest = args[0];
    if (!dest) {
      if (reparto !== undefined || args.length) throw new Error('per preparare un sito va scritta la cartella di destinazione');
      elencoSito().forEach((p) => console.log(p));
    } else {
      const compressa = args[1] === 'compressa';
      if (args.length > 2 || (args[1] && !compressa)) throw new Error('argomento non riconosciuto: ' + (args[1] && !compressa ? args[1] : args[2]));
      const r = prepara(dest, { compressa, reparto });
      const dove = path.relative(RADICE, path.resolve(dest)).split(path.sep).join('/') + '/';
      if (!compressa) {
        r.elenco.forEach((p) => console.log('  ' + p));
        console.log(r.elenco.length + ' file, ' + r.byte + ' byte in ' + dove + ' per il reparto «' + r.reparto + '»');
      } else {
        const c = r.compressione, kb = (n) => (n / 1024).toFixed(0).padStart(5) + ' KB';
        c.righe.forEach((x) => console.log('  ' + kb(x.prima) + ' → ' + kb(x.dopo) + '  ' + x.file.padEnd(34) + x.nota));
        console.log(r.elenco.length + ' file in ' + dove + ' per il reparto «' + r.reparto + '»: da ' + kb(c.prima).trim() + ' a ' + kb(c.dopo).trim() + ' (' + c.strumento + '); ' + c.nomiGlobali + ' nomi globali rimasti identici');
      }
    }
  } catch (e) { console.log('ERRORE: ' + e.message); process.exitCode = 1; }
}
