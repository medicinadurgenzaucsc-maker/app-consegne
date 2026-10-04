// Chiamate all'API di GitHub per l'account di COLLAUDO (gistech2026).
// Il token si legge dall'archivio credenziali di Windows tramite GCM e non
// viene mai stampato. Uso:
//   node gh-api.js <METODO> <percorso> ['<corpo json>']
// Stampa lo stato HTTP e la risposta (JSON) con i token mascherati.
const { spawnSync } = require('child_process');

const UTENTE = 'gistech2026';
const maschera = (s) => String(s || '').replace(/gh[pousr]_[A-Za-z0-9_]+/g, 'gh*_***');

function token() {
  const r = spawnSync('git', ['credential-manager', 'get'], {
    input: 'protocol=https\nhost=github.com\nusername=' + UTENTE + '\n\n',
    encoding: 'utf8',
    env: { ...process.env, GCM_INTERACTIVE: 'never', GIT_TERMINAL_PROMPT: '0' },
  });
  const m = /^password=(.+)$/m.exec(r.stdout || '');
  const u = /^username=(.+)$/m.exec(r.stdout || '');
  if (!m) throw new Error('credenziale di ' + UTENTE + ' non trovata in GCM');
  if (u && u[1].trim().toLowerCase() !== UTENTE) throw new Error('GCM ha restituito un altro account: ' + u[1].trim());
  return m[1].trim();
}

async function gh(metodo, percorso, corpo) {
  // Solo l'account di collaudo: nessuna chiamata può riguardare altri proprietari.
  if (/^\/repos\//.test(percorso) && !percorso.startsWith('/repos/' + UTENTE + '/')) {
    throw new Error('percorso fuori dall\'account di collaudo: ' + percorso);
  }
  const r = await fetch('https://api.github.com' + percorso, {
    method: metodo,
    headers: {
      Authorization: 'Bearer ' + token(),
      Accept: 'application/vnd.github+json',
      'X-GitHub-Api-Version': '2022-11-28',
      'User-Agent': 'collaudo-consegne',
      ...(corpo ? { 'Content-Type': 'application/json' } : {}),
    },
    body: corpo ? JSON.stringify(corpo) : undefined,
  });
  const testo = await r.text();
  let dati = null;
  try { dati = testo ? JSON.parse(testo) : null; } catch (e) { dati = testo; }
  return { stato: r.status, dati, permessi: r.headers.get('x-oauth-scopes') || '' };
}

module.exports = { gh, token, UTENTE };

if (require.main === module) {
  const [metodo, percorso, corpo] = process.argv.slice(2);
  gh(metodo || 'GET', percorso || '/user', corpo ? JSON.parse(corpo) : undefined)
    .then((x) => { console.log('HTTP ' + x.stato); console.log(maschera(JSON.stringify(x.dati, null, 1)).slice(0, 4000)); })
    .catch((e) => { console.log('ERRORE: ' + maschera(e.message)); process.exit(1); });
}
