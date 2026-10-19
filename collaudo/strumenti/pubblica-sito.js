// Pubblica il sito di servizio del collaudo (oggi https://gistech2026.github.io/collaudo/):
//   index.html          pagina d'ingresso
//   informativa.html    informativa sull'accesso (è il link «privacy» del client Google di collaudo)
//   trak-finto/         il finto TrakCare
// Sta in un repository a parte perché NON deve finire in docs/: ciò che sta in
// docs/ arriva anche in produzione. Quale repository lo dice il file del reparto
// di collaudo, voce «github.sitoServizio» (dalla v199): può essere anche di
// un'organizzazione.
// Ogni pubblicazione è un solo commit che rispecchia esattamente i file locali:
// sostituisce tutto il contenuto del ramo, quindi si scrive solo in un repository
// appena nato o che contiene già soltanto i file di questo sito (un nome sbagliato
// nel file del reparto non deve svuotare un altro repository), e un repository
// che manca si crea solo con «--crea».
//
//   node collaudo/strumenti/pubblica-sito.js           pubblica e aspetta che il sito la serva
//   node collaudo/strumenti/pubblica-sito.js --crea    come sopra, e se il repository non c'è lo crea
const crypto = require('crypto');
const fs = require('fs');
const path = require('path');
const AMB = require('./gh-api.js').per('collaudo');
const gh = AMB.gh;

if (!AMB.sitoServizio) { console.log('ERRORE: il file del reparto di collaudo non dice dove sta il sito di servizio (voce github.sitoServizio)'); process.exit(1); }
const [PROPRIETARIO, REPO] = AMB.sitoServizio.split('/');
const SITO = 'https://' + PROPRIETARIO.toLowerCase() + '.github.io/' + REPO + '/';
const COLLAUDO = path.resolve(__dirname, '..');
const FILE = [
  ['index.html', path.join(COLLAUDO, 'sito', 'index.html')],
  ['informativa.html', path.join(COLLAUDO, 'sito', 'informativa.html')],
  ['trak-finto/index.html', path.join(COLLAUDO, 'trak-finto', 'index.html')],
  ['trak-finto/catalogo.js', path.join(COLLAUDO, 'trak-finto', 'catalogo.js')],
];
// i soli file che il repository del sito di servizio può contenere (README.md e .nojekyll: un repository appena nato)
const DEL_SITO = FILE.map((f) => f[0]).concat(['versione.txt', '.nojekyll', 'README.md']);
const argomenti = process.argv.slice(2);
if (argomenti.length > 1 || (argomenti.length && argomenti[0] !== '--crea')) { console.log('uso: pubblica-sito.js [--crea]'); process.exit(2); }
const CREA = argomenti[0] === '--crea';
const attendi = (ms) => new Promise((r) => setTimeout(r, ms));
const api = async (metodo, percorso, corpo, ammessi) => {
  const r = await gh(metodo, percorso, corpo);
  if (r.stato >= 300 && !(ammessi || []).includes(r.stato)) throw new Error(metodo + ' ' + percorso + ' → HTTP ' + r.stato + ' ' + JSON.stringify(r.dati).slice(0, 300));
  return r;
};

(async () => {
  const base = '/repos/' + PROPRIETARIO + '/' + REPO;
  let scheda = await api('GET', base, null, [404]);
  if (scheda.stato === 404) {
    if (!CREA) throw new Error('il repository ' + PROPRIETARIO + '/' + REPO + ' non esiste. Per crearlo, pubblico: «pubblica-sito.js --crea»');
    // nell'account dell'utente con /user/repos, in un'organizzazione con /orgs/<nome>/repos
    const c = await AMB.crea(PROPRIETARIO, { name: REPO, description: 'Sito di servizio del collaudo delle Consegne: finto TrakCare e informativa (dati inventati)', private: false, auto_init: true, has_issues: false, has_wiki: false });
    if (c.stato >= 300) throw new Error('creazione del repository ' + PROPRIETARIO + '/' + REPO + ' → HTTP ' + c.stato + ' ' + JSON.stringify(c.dati).slice(0, 300));
    console.log('repository ' + PROPRIETARIO + '/' + REPO + ' creato');
    await attendi(3000);
    scheda = await api('GET', base);
  } else if (!scheda.dati || String(scheda.dati.name) !== REPO || String((scheda.dati.owner || {}).login).toLowerCase() !== PROPRIETARIO.toLowerCase()) {
    // GitHub rimanda un repository rinominato o trasferito: scrivere lì vorrebbe dire scrivere altrove
    throw new Error('a ' + PROPRIETARIO + '/' + REPO + ' risponde un altro repository (' + String((scheda.dati && scheda.dati.full_name) || '?') + '): il file del reparto va corretto');
  } else if (scheda.dati.private) throw new Error('il repository ' + PROPRIETARIO + '/' + REPO + ' è privato: con un account gratuito il suo sito non verrebbe servito');
  const contenuti = FILE.map(([dove, da]) => [dove, fs.readFileSync(da)]);
  const impronta = crypto.createHash('sha256').update(Buffer.concat(contenuti.map(([d, c]) => Buffer.concat([Buffer.from(d + '\0'), c])))).digest('hex').slice(0, 16);
  contenuti.push(['versione.txt', Buffer.from(impronta + '\n')], ['.nojekyll', Buffer.from('')]);

  const ramo = scheda.dati.default_branch;
  const testa = (await api('GET', base + '/git/ref/heads/' + ramo)).dati.object.sha;
  // PRIMA di scrivere: il ramo contiene solo i file di questo sito (o quelli di un repository appena nato)?
  const vecchio = (await api('GET', base + '/git/trees/' + (await api('GET', base + '/git/commits/' + testa)).dati.tree.sha + '?recursive=1')).dati || {};
  const estranei = (vecchio.tree || []).filter((x) => x.type === 'blob' && DEL_SITO.indexOf(x.path) < 0).map((x) => x.path);
  if (vecchio.truncated || estranei.length) throw new Error('il repository ' + PROPRIETARIO + '/' + REPO + ' contiene altro (' + (vecchio.truncated ? 'troppi file' : estranei.slice(0, 4).join(', ')) + '): la pubblicazione ne sostituirebbe tutto il contenuto. Nulla è stato scritto');
  const albero = [];
  for (const [dove, corpo] of contenuti) {
    const blob = await api('POST', base + '/git/blobs', { content: corpo.toString('base64'), encoding: 'base64' });
    albero.push({ path: dove, mode: '100644', type: 'blob', sha: blob.dati.sha });
  }
  const tree = await api('POST', base + '/git/trees', { tree: albero });
  const attuale = (await api('GET', base + '/git/commits/' + testa)).dati.tree.sha;
  if (attuale === tree.dati.sha) console.log('nessuna differenza: il repository contiene già questi file');
  else {
    const commit = await api('POST', base + '/git/commits', { message: 'Sito del collaudo ' + impronta, tree: tree.dati.sha, parents: [testa] });
    await api('PATCH', base + '/git/refs/heads/' + ramo, { sha: commit.dati.sha });
    console.log('pubblicato il commit ' + commit.dati.sha.slice(0, 7) + ' (' + contenuti.length + ' file)');
  }
  // in un'organizzazione che non lascia pubblicare siti ai membri, GitHub Pages lo accende un proprietario
  const pagine = await api('POST', base + '/pages', { source: { branch: ramo, path: '/' } }, [403, 404, 409, 422]);
  if (pagine.stato < 300) console.log('GitHub Pages attivato');
  else if ((await api('GET', base + '/pages', null, [403, 404])).stato !== 200) throw new Error('i file sono nel repository, ma GitHub Pages non è acceso e questa chiave non può accenderlo (HTTP ' + pagine.stato + '): lo accende un proprietario dell\'organizzazione (Settings del repository, Pages, ramo ' + ramo + ', cartella /), poi si rilancia lo strumento');
  else console.log('GitHub Pages già attivo');

  let servita = '';
  for (let i = 0; i < 45; i++) {
    try { const r = await fetch(SITO + 'versione.txt?cb=' + Date.now()); servita = r.ok ? (await r.text()).trim() : ''; } catch (e) { servita = ''; }
    if (servita === impronta) break;
    await attendi(8000);
  }
  if (servita !== impronta) { console.log('ERRORE: il sito non serve ancora la versione ' + impronta + ' (serve «' + servita + '»)'); process.exit(1); }
  console.log('in linea: ' + SITO + ' (versione ' + impronta + ')');
})().catch((e) => { console.log('ERRORE: ' + e.message); process.exit(1); });
