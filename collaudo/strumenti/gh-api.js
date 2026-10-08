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
const AMBIENTI = {
  collaudo: { utente: 'gistech2026', remoto: 'collaudo' },
  produzione: { utente: 'medicinadurgenzaucsc-maker', remoto: 'origin' },
};
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
  const a = AMBIENTI[ambiente];
  if (!a) throw new Error('ambiente sconosciuto: ' + ambiente + ' (validi: ' + Object.keys(AMBIENTI).join(', ') + ')');
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
    // Solo l'account di questo ambiente: nessuna chiamata può riguardare altri proprietari.
    if (/^\/repos\//.test(percorso) && percorso.toLowerCase().indexOf('/repos/' + UTENTE.toLowerCase() + '/') !== 0) {
      throw new Error('percorso fuori dall\'account di ' + ambiente + ': ' + percorso);
    }
    if (produzione && String(metodo).toUpperCase() !== 'GET' && process.argv.indexOf(CONFERMA) < 0) {
      throw new Error('in PRODUZIONE questa chiamata cambierebbe qualcosa (' + metodo + ' ' + percorso + '): rilanciare lo strumento con ' + CONFERMA + ', e solo con l\'ok di chi gestisce l\'app');
    }
    const r = await fetch('https://api.github.com' + percorso, {
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
  return { ambiente, UTENTE, remoto: a.remoto, gh, token, produzione };
}

// Da un elenco di argomenti: l'ambiente (se è il primo) e il resto, senza la conferma.
function argomenti(argv) {
  const resto = argv.filter((x) => x !== CONFERMA);
  const ambiente = AMBIENTI[resto[0]] ? resto.shift() : 'collaudo';
  return { ambiente, resto, confermato: argv.indexOf(CONFERMA) >= 0 };
}

const collaudo = per('collaudo');
module.exports = { gh: collaudo.gh, token: collaudo.token, UTENTE: collaudo.UTENTE, per, argomenti, maschera, AMBIENTI, CONFERMA };

if (require.main === module) {
  const [metodo, percorso, corpo] = process.argv.slice(2);
  collaudo.gh(metodo || 'GET', percorso || '/user', corpo ? JSON.parse(corpo) : undefined)
    .then((x) => { console.log('HTTP ' + x.stato); console.log(maschera(JSON.stringify(x.dati, null, 1)).slice(0, 4000)); })
    .catch((e) => { console.log('ERRORE: ' + maschera(e.message)); process.exit(1); });
}
