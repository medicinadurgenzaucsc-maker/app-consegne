// Chiamate all'API di GitHub per gli account dei due ambienti.
//
//   collaudo    account gistech2026: la chiave si legge dall'archivio credenziali
//               di Windows tramite GCM (vedi github-accesso.js);
//   produzione  account medicinadurgenzaucsc-maker: la chiave si cerca nello
//               stesso archivio e, se non c'è, nell'indirizzo del remoto «origin»
//               di questa cartella. In produzione ogni chiamata che CAMBIA
//               qualcosa è rifiutata, a meno che lo strumento che la fa non sia
//               stato lanciato con «--confermo-produzione».
//
// Nessuna chiave viene mai stampata. Uso da riga di comando (solo collaudo):
//   node gh-api.js <METODO> <percorso> ['<corpo json>']
const { spawnSync, execFileSync } = require('child_process');
const path = require('path');

const RADICE = path.resolve(__dirname, '../..');
// «sitoNuovo» è l'indirizzo del sito su Cloudflare: è lì che la pagina di rinvio
// di quell'ambiente deve portare, e non altrove.
const AMBIENTI = {
  collaudo: { utente: 'gistech2026', remoto: 'collaudo', sitoNuovo: 'https://consegne-collaudo.pages.dev/' },
  produzione: { utente: 'medicinadurgenzaucsc-maker', remoto: 'origin', sitoNuovo: 'https://consegne-reparto.pages.dev/' },
};
// Solo i due nomi scritti qui sopra: «constructor» e simili non sono ambienti.
const eAmbiente = (x) => Object.prototype.hasOwnProperty.call(AMBIENTI, String(x));
const CONFERMA = '--confermo-produzione';
const maschera = (s) => String(s || '').replace(/(gh[pousr]_|github_pat_)[A-Za-z0-9_]+/g, 'gh*_***');

function daGcm(utente) {
  const r = spawnSync('git', ['credential-manager', 'get'], {
    input: 'protocol=https\nhost=github.com\nusername=' + utente + '\n\n',
    encoding: 'utf8',
    env: { ...process.env, GCM_INTERACTIVE: 'never', GIT_TERMINAL_PROMPT: '0' },
  });
  const m = /^password=(.+)$/m.exec(r.stdout || '');
  const u = /^username=(.+)$/m.exec(r.stdout || '');
  if (!m) return null;
  if (u && u[1].trim().toLowerCase() !== utente.toLowerCase()) return null;   // GCM ha risposto con un altro account
  return m[1].trim();
}
// La chiave scritta nell'indirizzo di un remoto (https://chiave@github.com/proprietario/…), solo se il remoto è di quel proprietario.
function dalRemoto(remoto, utente) {
  let u = '';
  try { u = execFileSync('git', ['-C', RADICE, 'remote', 'get-url', remoto], { encoding: 'utf8' }).trim(); } catch (e) { return null; }
  const m = /^https:\/\/([^@/]+)@github\.com\/([^/]+)\//.exec(u);
  if (!m || m[2].toLowerCase() !== utente.toLowerCase()) return null;
  const chiave = m[1].indexOf(':') >= 0 ? m[1].split(':').pop() : m[1];
  return /^(gh[pousr]_|github_pat_)/.test(chiave) ? chiave : null;
}

// Tutto ciò che serve per lavorare su un ambiente.
function per(ambiente) {
  if (!eAmbiente(ambiente)) throw new Error('ambiente sconosciuto: ' + ambiente + ' (validi: ' + Object.keys(AMBIENTI).join(', ') + ')');
  const a = AMBIENTI[ambiente];
  const UTENTE = a.utente;
  const produzione = ambiente === 'produzione';
  let memoria = null;
  function token() {
    if (memoria) return memoria;
    memoria = daGcm(UTENTE) || (produzione ? dalRemoto(a.remoto, UTENTE) : null);
    if (!memoria) throw new Error('chiave di GitHub per ' + UTENTE + ' non trovata' + (produzione ? ' né in GCM né nel remoto «' + a.remoto + '»' : ' in GCM'));
    return memoria;
  }
  async function gh(metodo, percorso, corpo) {
    // Solo l'account di questo ambiente. Il percorso si controlla nella forma in cui parte
    // davvero: «..» viene risolto prima di spedire, e un percorso senza la barra iniziale
    // si attaccherebbe al nome del sito, mandando la chiave altrove.
    if (typeof percorso !== 'string' || percorso.charAt(0) !== '/') throw new Error('percorso non valido: deve cominciare con «/»');
    const u = new URL('https://api.github.com' + percorso);
    if (u.origin !== 'https://api.github.com') throw new Error('percorso non valido');
    const p = u.pathname.toLowerCase();
    if (p.indexOf('/repos/' + UTENTE.toLowerCase() + '/') !== 0 && p !== '/user' && p !== '/user/repos') {
      throw new Error('percorso fuori dall\'account di ' + ambiente + ': ' + u.pathname);
    }
    if (produzione && String(metodo).toUpperCase() !== 'GET' && process.argv.indexOf(CONFERMA) < 0) {
      throw new Error('in PRODUZIONE questa chiamata cambierebbe qualcosa (' + metodo + ' ' + percorso + '): rilanciare lo strumento con ' + CONFERMA + ', e solo con l\'ok di chi gestisce l\'app');
    }
    const r = await fetch(u.href, {
      method: metodo,
      headers: {
        Authorization: 'Bearer ' + token(),
        Accept: 'application/vnd.github+json',
        'X-GitHub-Api-Version': '2022-11-28',
        'User-Agent': 'consegne-' + ambiente,
        ...(corpo ? { 'Content-Type': 'application/json' } : {}),
      },
      body: corpo ? JSON.stringify(corpo) : undefined,
    });
    const testo = await r.text();
    let dati = null;
    try { dati = testo ? JSON.parse(testo) : null; } catch (e) { dati = testo; }
    return { stato: r.status, dati, permessi: r.headers.get('x-oauth-scopes') || '' };
  }
  return { ambiente, UTENTE, remoto: a.remoto, sitoNuovo: a.sitoNuovo, gh, token, produzione };
}

// Da un elenco di argomenti: l'ambiente (se è il primo) e il resto, senza la conferma.
// Un ambiente scritto fuori posto non viene ignorato: lo strumento si ferma,
// altrimenti «annulla produzione» lavorerebbe sul collaudo senza dirlo.
// «esplicito» dice se l'ambiente è stato scritto: i comandi che cambiano qualcosa
// lo pretendono (vedi pretendiAmbiente), così un comando copiato da un documento
// senza l'ambiente non lavora sul collaudo al posto della produzione.
function argomenti(argv) {
  const confermato = argv.indexOf(CONFERMA) >= 0;
  const resto = argv.filter((x) => x !== CONFERMA);
  const sconosciuta = resto.filter((x) => /^--/.test(String(x)))[0];
  if (sconosciuta) throw new Error('opzione sconosciuta: ' + sconosciuta + ' (l\'unica ammessa è ' + CONFERMA + ')');
  let ambiente = 'collaudo', esplicito = false;
  if (resto.length && eAmbiente(String(resto[0]).toLowerCase())) { ambiente = String(resto.shift()).toLowerCase(); esplicito = true; }
  const fuoriPosto = resto.filter((x) => eAmbiente(String(x).toLowerCase()));
  if (fuoriPosto.length) throw new Error('il nome dell\'ambiente («' + fuoriPosto[0] + '») va scritto per primo, subito dopo il nome dello strumento, e una volta sola');
  if (confermato && ambiente !== 'produzione') throw new Error(CONFERMA + ' vale solo per la produzione: scrivere «produzione» come primo argomento');
  return { ambiente, resto, confermato, esplicito };
}
// Per i comandi che cambiano qualcosa: l'ambiente va scritto, non sottinteso.
function pretendiAmbiente(arg, comando) {
  if (!arg.esplicito) throw new Error('«' + comando + '» cambia qualcosa: l\'ambiente va scritto per esteso come primo argomento (collaudo oppure produzione)');
}

const collaudo = per('collaudo');
module.exports = { gh: collaudo.gh, token: collaudo.token, UTENTE: collaudo.UTENTE, per, argomenti, pretendiAmbiente, maschera, AMBIENTI, CONFERMA };

if (require.main === module) {
  const [metodo, percorso, corpo] = process.argv.slice(2);
  collaudo.gh(metodo || 'GET', percorso || '/user', corpo ? JSON.parse(corpo) : undefined)
    .then((x) => { console.log('HTTP ' + x.stato); console.log(maschera(JSON.stringify(x.dati, null, 1)).slice(0, 4000)); })
    .catch((e) => { console.log('ERRORE: ' + maschera(e.message)); process.exit(1); });
}
