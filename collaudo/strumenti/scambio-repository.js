// Prova generale dello «scambio dei repository», per ora SOLO nel collaudo
// (account gistech2026: gh-api.js non lascia toccare altri proprietari).
//
// Prima dello scambio:  app-consegne            pubblico, sorgente + sito su GitHub Pages
// Dopo lo scambio:      app-consegne-sorgente   PRIVATO, il sorgente (storia, segreti, flussi)
//                       app-consegne            pubblico, SOLO la pagina di rinvio al sito nuovo
// Il repository della pagina di rinvio si prepara prima, con un altro nome
// (app-consegne-rinvio) e il sito già pubblicato: allo scambio bastano due
// cambi di nome, e si misura per quanti secondi il vecchio indirizzo non risponde.
//
//   node collaudo/strumenti/scambio-repository.js stato
//   node collaudo/strumenti/scambio-repository.js prepara <nuovo indirizzo>   es. https://consegne-collaudo.pages.dev/
//   node collaudo/strumenti/scambio-repository.js scambia
//   node collaudo/strumenti/scambio-repository.js annulla                     rimette tutto com'era (per riprovare)
//
// Dopo lo scambio il remoto «collaudo» di questa cartella punta al repository
// privato. ATTENZIONE alle altre copie della cartella: il loro remoto punterebbe
// ancora al vecchio nome, che ora è il repository pubblico del rinvio. Per questo
// quel repository viene chiuso alle scritture (regola «sola-lettura»).
const { execFileSync } = require('child_process');
const { gh, UTENTE } = require('./gh-api.js');
const { fileRinvio, destinazione, RADICE } = require('./prepara-rinvio.js');

const NOMI = { app: 'app-consegne', sorgente: 'app-consegne-sorgente', rinvio: 'app-consegne-rinvio' };
const SITO = 'https://' + UTENTE + '.github.io/';
const REMOTO = 'collaudo';
const REGOLA = 'sola-lettura';
const attendi = (ms) => new Promise((r) => setTimeout(r, ms));
const secondi = (t0) => ((Date.now() - t0) / 1000).toFixed(1) + ' s';

async function api(metodo, percorso, corpo, ammessi) {
  const r = await gh(metodo, percorso, corpo);
  if (r.stato >= 300 && !(ammessi || []).includes(r.stato)) throw new Error(metodo + ' ' + percorso + ' → HTTP ' + r.stato + ' ' + JSON.stringify(r.dati).slice(0, 300));
  return r;
}
const base = (nome) => '/repos/' + UTENTE + '/' + nome;
async function repo(nome) {
  const r = await api('GET', base(nome), null, [404]);
  // un nome appena cambiato viene ancora «rinviato» da GitHub al repository nuovo: non è lui
  if (r.stato === 404 || !r.dati || r.dati.name !== nome) return null;
  return r.dati;
}
// Che cosa risponde un indirizzo: l'app, la pagina di rinvio, un 404, altro.
async function cheCosa(url) {
  try {
    const r = await fetch(url + (url.indexOf('?') < 0 ? '?' : '&') + 'cb=' + Date.now(), { redirect: 'manual', cache: 'no-store' });
    if (r.status === 404) return '404';
    if (r.status >= 300) return 'HTTP ' + r.status;
    const t = await r.text();
    if (t.indexOf('navVersioneApp') >= 0) return 'app';
    if (t.indexOf('ha cambiato indirizzo') >= 0) return 'rinvio';
    return 'altro';
  } catch (e) { return 'errore di rete'; }
}
async function regola(nome) {
  // 403: in un repository privato di un account gratuito le regole non esistono
  const r = await api('GET', base(nome) + '/rulesets', null, [404, 403]);
  return (Array.isArray(r.dati) ? r.dati : []).filter((x) => x.name === REGOLA)[0] || null;
}
async function sblocca(nome) {
  const x = await regola(nome);
  if (x) await api('DELETE', base(nome) + '/rulesets/' + x.id);
}
// Nessuno può più scrivere nei rami del repository del rinvio, nemmeno il
// proprietario: un push fatto per sbaglio col sorgente viene rifiutato.
async function blocca(nome) {
  if (await regola(nome)) return;
  await api('POST', base(nome) + '/rulesets', {
    name: REGOLA, target: 'branch', enforcement: 'active', bypass_actors: [],
    conditions: { ref_name: { include: ['~ALL'], exclude: [] } },
    rules: [{ type: 'creation' }, { type: 'update' }, { type: 'deletion' }, { type: 'non_fast_forward' }],
  });
}
// Il remoto locale, senza mai stampare la chiave che contiene.
function remoto() { return execFileSync('git', ['-C', RADICE, 'remote', 'get-url', REMOTO], { encoding: 'utf8' }).trim(); }
function nomeDelRemoto() { return (remoto().split('/').pop() || '').replace(/[.]git$/, ''); }
function puntaIlRemoto(nome) {
  const pezzi = remoto().split('/');
  pezzi[pezzi.length - 1] = nome + '.git';
  execFileSync('git', ['-C', RADICE, 'remote', 'set-url', REMOTO, pezzi.join('/')]);
}

async function stato() {
  for (const k of Object.keys(NOMI)) {
    const r = await repo(NOMI[k]);
    if (!r) { console.log(NOMI[k].padEnd(24) + 'non esiste'); continue; }
    const p = r.has_pages ? (await api('GET', base(NOMI[k]) + '/pages', null, [404])).dati : null;
    const chiuso = await regola(NOMI[k]);
    console.log(NOMI[k].padEnd(24) + r.visibility + ' | sito: ' + (p && p.source ? p.source.branch + ':' + p.source.path + ' (' + p.status + ')' : 'no') + (chiuso ? ' | chiuso alle scritture' : ''));
  }
  console.log('indirizzo ' + SITO + NOMI.app + '/         → ' + await cheCosa(SITO + NOMI.app + '/'));
  console.log('indirizzo ' + SITO + NOMI.rinvio + '/  → ' + await cheCosa(SITO + NOMI.rinvio + '/'));
  console.log('remoto «' + REMOTO + '» di questa cartella → ' + UTENTE + '/' + nomeDelRemoto());
}

// Crea (o aggiorna) il repository della pagina di rinvio, col suo sito già in linea.
async function prepara(nuovo) {
  const dest = destinazione(nuovo);
  if (!(await repo(NOMI.app))) throw new Error('il repository ' + NOMI.app + ' non esiste');
  if (await repo(NOMI.sorgente)) throw new Error('esiste già ' + NOMI.sorgente + ': lo scambio è stato fatto (per riprovare: «annulla»)');
  // il sito nuovo deve rispondere: 200, oppure il rinvio alla pagina di accesso del cancello
  const prova = await fetch(dest + 'sw.js', { redirect: 'manual', cache: 'no-store' }).then((r) => r.status, () => 0);
  if (prova !== 200 && prova !== 302) throw new Error('il sito nuovo ' + dest + ' non risponde come atteso (HTTP ' + prova + ')');
  let r = await repo(NOMI.rinvio);
  if (!r) {
    await api('POST', '/user/repos', { name: NOMI.rinvio, description: "Rinvio al nuovo indirizzo dell'applicazione", private: false, auto_init: true, has_issues: false, has_wiki: false, has_projects: false });
    console.log('repository ' + UTENTE + '/' + NOMI.rinvio + ' creato');
    await attendi(3000);
    r = await repo(NOMI.rinvio);
  }
  await sblocca(NOMI.rinvio);
  const b = base(NOMI.rinvio), ramo = r.default_branch;
  const albero = [];
  for (const [dove, corpo] of fileRinvio(dest)) {
    const blob = await api('POST', b + '/git/blobs', { content: corpo.toString('base64'), encoding: 'base64' });
    albero.push({ path: dove, mode: '100644', type: 'blob', sha: blob.dati.sha });
  }
  const tree = await api('POST', b + '/git/trees', { tree: albero });
  const testa = (await api('GET', b + '/git/ref/heads/' + ramo)).dati.object.sha;
  const attuale = (await api('GET', b + '/git/commits/' + testa)).dati;
  if (attuale.tree.sha === tree.dati.sha && !attuale.parents.length) console.log('la pagina di rinvio è già quella giusta');
  else {
    // un solo commit, senza storia: nel repository pubblico non resta altro
    const commit = await api('POST', b + '/git/commits', { message: 'Pagina di rinvio', tree: tree.dati.sha, parents: [] });
    await api('PATCH', b + '/git/refs/heads/' + ramo, { sha: commit.dati.sha, force: true });
    console.log('pagina di rinvio pubblicata (' + albero.length + ' file) verso ' + dest);
  }
  const pagine = await api('POST', b + '/pages', { source: { branch: ramo, path: '/' } }, [409, 422]);
  console.log(pagine.stato < 300 ? 'GitHub Pages attivato' : 'GitHub Pages già attivo');
  await blocca(NOMI.rinvio);
  const t0 = Date.now();
  let visto = '';
  for (let i = 0; i < 60; i++) { visto = await cheCosa(SITO + NOMI.rinvio + '/'); if (visto === 'rinvio') break; await attendi(5000); }
  if (visto !== 'rinvio') throw new Error('dopo ' + secondi(t0) + ' il sito del rinvio risponde ancora «' + visto + '»');
  console.log('pronto: ' + SITO + NOMI.rinvio + '/ rimanda a ' + dest + ' (in linea dopo ' + secondi(t0) + '); repository chiuso alle scritture');
}

// Aspetta che un indirizzo risponda nel modo voluto per tre volte di fila, annotando i cambi.
async function finche(url, voluto, nota, t0, repoDelSito) {
  let ultimo = '', buone = 0, primo = null, nonTrovato = 0, chiesta = false;
  for (let i = 0; i < 240 && buone < 3; i++) {
    const c = await cheCosa(url);
    if (c !== ultimo) { nota('il vecchio indirizzo risponde: ' + c); ultimo = c; }
    if (c === voluto) { buone++; if (primo == null) primo = Date.now() - t0; } else { buone = 0; primo = null; }
    if (c === '404') nonTrovato++;
    if (!chiesta && c !== voluto && Date.now() - t0 > 40000) {
      chiesta = true;
      const x = await api('POST', base(repoDelSito) + '/pages/builds', null, [403, 404, 409, 422]);
      nota('dopo 40 s chiesta a GitHub Pages una nuova pubblicazione (HTTP ' + x.stato + ')');
    }
    await attendi(1500);
  }
  return { riuscito: buone >= 3, primo: primo, nonTrovato: nonTrovato };
}

async function scambia() {
  if (!(await repo(NOMI.app)) || !(await repo(NOMI.rinvio))) throw new Error('servono sia ' + NOMI.app + ' sia ' + NOMI.rinvio + ' (prima «prepara»)');
  if (await repo(NOMI.sorgente)) throw new Error('esiste già ' + NOMI.sorgente + ': lo scambio è stato fatto');
  if ((await api('GET', base(NOMI.app) + '/contents/docs/sw.js', null, [404])).stato !== 200) throw new Error(NOMI.app + ' non sembra il repository del sorgente (manca docs/sw.js)');
  if (nomeDelRemoto() !== NOMI.app) throw new Error('il remoto «' + REMOTO + '» punta a ' + nomeDelRemoto() + ', atteso ' + NOMI.app);
  if (await cheCosa(SITO + NOMI.rinvio + '/') !== 'rinvio') throw new Error('il sito del rinvio non è in linea');
  if (await cheCosa(SITO + NOMI.app + '/') !== 'app') throw new Error("il vecchio indirizzo non sta servendo l'applicazione");
  const t0 = Date.now();
  const nota = (t) => console.log(secondi(t0).padStart(8) + '  ' + t);
  await api('PATCH', base(NOMI.app), { name: NOMI.sorgente });
  nota(NOMI.app + ' rinominato in ' + NOMI.sorgente);
  await api('PATCH', base(NOMI.rinvio), { name: NOMI.app });
  nota(NOMI.rinvio + ' rinominato in ' + NOMI.app);
  puntaIlRemoto(NOMI.sorgente);
  nota('remoto «' + REMOTO + '» puntato su ' + NOMI.sorgente);
  const esito = await finche(SITO + NOMI.app + '/', 'rinvio', nota, t0, NOMI.app);
  await api('PATCH', base(NOMI.sorgente), { private: true });
  nota(NOMI.sorgente + ' reso privato');
  console.log(esito.riuscito
    ? 'FATTO: il vecchio indirizzo rimanda al nuovo dopo ' + (esito.primo / 1000).toFixed(1) + ' s dalla prima rinomina; risposte «404» viste nel frattempo: ' + esito.nonTrovato
    : 'ATTENZIONE: dopo 6 minuti il vecchio indirizzo non serve ancora la pagina di rinvio');
}

async function annulla() {
  const s = await repo(NOMI.sorgente);
  if (!s || !(await repo(NOMI.app))) throw new Error('nulla da annullare: servono ' + NOMI.sorgente + ' e ' + NOMI.app);
  if ((await api('GET', base(NOMI.app) + '/contents/docs/sw.js', null, [404])).stato === 200) throw new Error(NOMI.app + ' contiene già il sorgente: stato inatteso, non tocco nulla');
  const t0 = Date.now();
  const nota = (t) => console.log(secondi(t0).padStart(8) + '  ' + t);
  if (s.private) { await api('PATCH', base(NOMI.sorgente), { private: false }); nota(NOMI.sorgente + ' di nuovo pubblico'); }
  await api('PATCH', base(NOMI.app), { name: NOMI.rinvio });
  nota(NOMI.app + ' rinominato in ' + NOMI.rinvio);
  await api('PATCH', base(NOMI.sorgente), { name: NOMI.app });
  nota(NOMI.sorgente + ' rinominato in ' + NOMI.app);
  puntaIlRemoto(NOMI.app);
  nota('remoto «' + REMOTO + '» puntato su ' + NOMI.app);
  const p = await api('POST', base(NOMI.app) + '/pages', { source: { branch: 'master', path: '/docs' } }, [409, 422]);
  nota('GitHub Pages sul sorgente: ' + (p.stato < 300 ? 'riattivato' : 'già attivo'));
  const esito = await finche(SITO + NOMI.app + '/', 'app', nota, t0, NOMI.app);
  console.log(esito.riuscito
    ? "TORNATO COM'ERA: il vecchio indirizzo serve di nuovo l'applicazione dopo " + (esito.primo / 1000).toFixed(1) + ' s'
    : "ATTENZIONE: dopo 6 minuti il vecchio indirizzo non serve ancora l'applicazione");
}

const AZIONI = { stato: stato, prepara: prepara, scambia: scambia, annulla: annulla };
const azione = AZIONI[process.argv[2] || 'stato'];
if (!azione) { console.log('uso: stato | prepara <nuovo indirizzo> | scambia | annulla'); process.exit(1); }
Promise.resolve(process.argv[3]).then(azione).catch((e) => { console.log('ERRORE: ' + String(e.message).replace(/gh[pousr]_[A-Za-z0-9_]+/g, 'gh*_***')); process.exit(1); });
