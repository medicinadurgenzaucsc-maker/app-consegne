// Pubblica il sito di servizio del collaudo: https://gistech2026.github.io/collaudo/
//   index.html          pagina d'ingresso
//   informativa.html    informativa sull'accesso (è il link «privacy» del client Google di collaudo)
//   trak-finto/         il finto TrakCare
// Sta in un repository a parte (gistech2026/collaudo) perché NON deve finire in
// docs/: ciò che sta in docs/ arriva anche in produzione.
// Ogni pubblicazione è un solo commit che rispecchia esattamente i file locali.
//
//   node collaudo/strumenti/pubblica-sito.js
const crypto = require('crypto');
const fs = require('fs');
const path = require('path');
const { gh, UTENTE } = require('./gh-api.js');

const REPO = 'collaudo';
const SITO = 'https://' + UTENTE + '.github.io/' + REPO + '/';
const COLLAUDO = path.resolve(__dirname, '..');
const FILE = [
  ['index.html', path.join(COLLAUDO, 'sito', 'index.html')],
  ['informativa.html', path.join(COLLAUDO, 'sito', 'informativa.html')],
  ['trak-finto/index.html', path.join(COLLAUDO, 'trak-finto', 'index.html')],
  ['trak-finto/catalogo.js', path.join(COLLAUDO, 'trak-finto', 'catalogo.js')],
];
const attendi = (ms) => new Promise((r) => setTimeout(r, ms));
const api = async (metodo, percorso, corpo, ammessi) => {
  const r = await gh(metodo, percorso, corpo);
  if (r.stato >= 300 && !(ammessi || []).includes(r.stato)) throw new Error(metodo + ' ' + percorso + ' → HTTP ' + r.stato + ' ' + JSON.stringify(r.dati).slice(0, 300));
  return r;
};

(async () => {
  const base = '/repos/' + UTENTE + '/' + REPO;
  if ((await api('GET', base, null, [404])).stato === 404) {
    await api('POST', '/user/repos', { name: REPO, description: 'Sito di servizio del collaudo delle Consegne: finto TrakCare e informativa (dati inventati)', private: false, auto_init: true, has_issues: false, has_wiki: false });
    console.log('repository ' + UTENTE + '/' + REPO + ' creato');
    await attendi(3000);
  }
  const contenuti = FILE.map(([dove, da]) => [dove, fs.readFileSync(da)]);
  const impronta = crypto.createHash('sha256').update(Buffer.concat(contenuti.map(([d, c]) => Buffer.concat([Buffer.from(d + '\0'), c])))).digest('hex').slice(0, 16);
  contenuti.push(['versione.txt', Buffer.from(impronta + '\n')], ['.nojekyll', Buffer.from('')]);

  const testa = (await api('GET', base + '/git/ref/heads/main')).dati.object.sha;
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
    await api('PATCH', base + '/git/refs/heads/main', { sha: commit.dati.sha });
    console.log('pubblicato il commit ' + commit.dati.sha.slice(0, 7) + ' (' + contenuti.length + ' file)');
  }
  const pagine = await api('POST', base + '/pages', { source: { branch: 'main', path: '/' } }, [409, 422]);
  console.log(pagine.stato < 300 ? 'GitHub Pages attivato' : 'GitHub Pages già attivo');

  let servita = '';
  for (let i = 0; i < 45; i++) {
    try { const r = await fetch(SITO + 'versione.txt?cb=' + Date.now()); servita = r.ok ? (await r.text()).trim() : ''; } catch (e) { servita = ''; }
    if (servita === impronta) break;
    await attendi(8000);
  }
  if (servita !== impronta) { console.log('ERRORE: il sito non serve ancora la versione ' + impronta + ' (serve «' + servita + '»)'); process.exit(1); }
  console.log('in linea: ' + SITO + ' (versione ' + impronta + ')');
})().catch((e) => { console.log('ERRORE: ' + e.message); process.exit(1); });
