// LIVELLO 0: il sito di un reparto pubblicato come SOLA copia compressa in un
// repository pubblico, servito da GitHub Pages. Il sorgente (docs/ leggibile,
// strumenti, documenti, storia) resta nel repository privato: in quello
// pubblico ci sono soltanto i file che il sito serve, già compressi, cioè ciò
// che chiunque apra il sito riceve comunque.
//
//   node collaudo/strumenti/sito-pubblico.js <reparto> stato             sola lettura
//   node collaudo/strumenti/sito-pubblico.js <reparto> pubblica          prepara la copia compressa del reparto, la
//                                                                        pubblica e aspetta che il sito la serva
//   node collaudo/strumenti/sito-pubblico.js <reparto> pubblica --crea   come sopra, e se il repository non esiste lo crea
//
// Il repository lo dice il file del reparto (pubblicazione.sitoPubblico). A ogni
// pubblicazione il ramo riceve UN commit in più il cui contenuto è l'intero
// sito: nessuna spinta forzata, e quindi ciò che è stato pubblicato una volta
// resta nella storia pubblica. Per questo:
//  - si pubblica solo ciò che è REGISTRATO in git (con modifiche in sospeso ai
//    file del sito lo strumento si ferma);
//  - si scrive solo in un repository che è già il sito pubblico di quel reparto,
//    o che è appena nato: sostituendo l'intero contenuto, in un repository
//    sbagliato cancellerebbe dal ramo ciò che c'è;
//  - il repository non si crea da solo: serve «--crea», e la chiave deve essere
//    proprio quella dell'account che lo possiede.
// Da riga di comando lo strumento lavora solo per un reparto di COLLAUDO (per
// un reparto di produzione il passaggio al livello 0 è un rilascio a sé, da
// fare col suo piano). Le funzioni sono scritte perché le possa usare anche il
// flusso di pubblicazione, che oggi NON ha ancora questo passo.
// Nessuna chiave viene mai stampata.
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const { execFileSync } = require('child_process');
const reparti = require('./reparti.js');

const RADICE = path.resolve(__dirname, '../..');
// File di configurazione di Cloudflare: su GitHub Pages non servono e non si pubblicano.
const SOLO_CLOUDFLARE = ['_headers', '_redirects'];
// Ciò da cui nasce il sito: con modifiche non registrate qui dentro non si pubblica.
const ORIGINE_DEL_SITO = ['docs', 'reparti', '.github/compressione', 'collaudo/strumenti/prepara-sito.js', 'collaudo/strumenti/reparti.js'];
const attendi = (ms) => new Promise((r) => setTimeout(r, ms));
const impronta = (b) => crypto.createHash('sha256').update(b).digest('hex');

// Dove sta il sito pubblico di un reparto, oppure null se il reparto non ne ha uno.
function destinazione(reparto) {
  const s = reparto.pubblicazione && reparto.pubblicazione.sitoPubblico;
  if (!s) return null;
  const [proprietario, nome] = s.repository.split('/');
  const host = proprietario.toLowerCase() + '.github.io';
  if (!reparti.hostAmmesso(reparto, host)) throw new Error('l\'indirizzo ' + host + ' non è fra quelli del reparto «' + reparto.codice + '»: il sito pubblicato lì si fermerebbe su «Indirizzo non riconosciuto»');
  return { proprietario, nome, host, sito: 'https://' + host + '/' + nome + '/' };
}

// I file da pubblicare: quelli della cartella preparata da prepara-sito.js, tolti quelli che servono
// solo a Cloudflare, più il segno che dice a GitHub Pages di servire i file così come sono.
function fileDelSito(cartella, elenco) {
  const file = elenco.filter((p) => SOLO_CLOUDFLARE.indexOf(p) < 0).map((p) => [p, fs.readFileSync(path.join(cartella, p))]);
  if (!file.some((f) => f[0] === 'index.html') || !file.some((f) => f[0] === 'sw.js') || !file.some((f) => f[0] === 'js/ambiente.js')) throw new Error('la cartella preparata non contiene il sito (mancano index.html, sw.js o js/ambiente.js)');
  file.push(['.nojekyll', Buffer.alloc(0)]);
  return file;
}

// I database nominati in un file dell'ambiente (uno solo, in un file giusto).
function databaseDi(testo) {
  return (String(testo || '').match(/https:\/\/[a-z]{20}\.supabase\.co/g) || []).filter((v, i, a) => a.indexOf(v) === i);
}
// La pubblicazione sostituisce l'INTERO contenuto del ramo. Il repository va bene solo se è appena nato
// (c'è il solo file iniziale) oppure è già il sito pubblico di QUESTO reparto: ha i file di un sito e il
// suo file dell'ambiente nomina lo stesso database di quello che si sta per pubblicare. Restituisce null
// se va bene, altrimenti il motivo. «percorsi» sono i file che il ramo contiene oggi.
function repositoryInadatto(percorsi, ambienteCheCe, ambienteNuovo) {
  if (percorsi.every((p) => /^(README(\.md)?|\.nojekyll)$/i.test(p))) return null;
  if (percorsi.indexOf('index.html') < 0 || percorsi.indexOf('sw.js') < 0 || percorsi.indexOf('js/ambiente.js') < 0) return 'contiene altro (' + percorsi.length + ' file, e non sono quelli di un sito pubblicato da questo strumento)';
  const nuovo = databaseDi(ambienteNuovo), cheCe = databaseDi(ambienteCheCe);
  if (nuovo.length !== 1) return 'il file dell\'ambiente da pubblicare non nomina un solo database';
  if (cheCe.length !== 1 || cheCe[0] !== nuovo[0]) return 'contiene il sito di un altro reparto (il suo file dell\'ambiente nomina un altro database)';
  return null;
}

// «gh» è una funzione (metodo, percorso, corpo) → { stato, dati }: quella di gh-api.js da questo PC,
// oppure una fatta da un flusso con la sua chiave. Il repository deve esistere già ed essere pubblico.
// Restituisce { cambiato, ramo, commit }: «commit» è quello che il ramo porta alla fine.
async function pubblica(gh, dest, file, messaggio, nota) {
  const dice = nota || (() => {});
  const api = async (metodo, percorso, corpo, ammessi) => {
    const r = await gh(metodo, percorso, corpo);
    if (r.stato >= 300 && !(ammessi || []).includes(r.stato)) throw new Error(metodo + ' ' + percorso + ' → HTTP ' + r.stato + ' ' + JSON.stringify(r.dati).slice(0, 300));
    return r;
  };
  const base = '/repos/' + dest.proprietario + '/' + dest.nome;
  const dove = dest.proprietario + '/' + dest.nome;
  const r = await api('GET', base, null, [404]);
  if (r.stato === 404 || !r.dati || String(r.dati.name).toLowerCase() !== dest.nome.toLowerCase()) throw new Error('il repository ' + dove + ' non esiste: va creato prima, con un primo file (un repository del tutto vuoto non basta)');
  if (r.dati.private) throw new Error('il repository ' + dove + ' è privato: con un account gratuito il suo sito non verrebbe servito');
  const ramo = r.dati.default_branch;
  const rif = await api('GET', base + '/git/ref/heads/' + ramo, null, [404, 409]);
  if (rif.stato !== 200) throw new Error('il repository ' + dove + ' è vuoto (nessun commit sul ramo ' + ramo + '): va creato con un primo file');
  const testa = rif.dati.object.sha;
  const attuale = (await api('GET', base + '/git/commits/' + testa)).dati;
  // PRIMA di scrivere qualunque cosa: è davvero il repository di questo sito?
  const vecchio = (await api('GET', base + '/git/trees/' + attuale.tree.sha + '?recursive=1')).dati || {};
  const voci = (vecchio.tree || []).filter((x) => x.type === 'blob');
  if (vecchio.truncated) throw new Error('il repository ' + dove + ' contiene troppi file per essere un sito pubblicato da questo strumento. Nulla è stato scritto');
  const suoAmbiente = voci.filter((x) => x.path === 'js/ambiente.js')[0];
  const cheCe = suoAmbiente ? Buffer.from(String((await api('GET', base + '/git/blobs/' + suoAmbiente.sha)).dati.content || ''), 'base64').toString('utf8') : '';
  const inadatto = repositoryInadatto(voci.map((x) => x.path), cheCe, file.filter((f) => f[0] === 'js/ambiente.js')[0][1].toString('utf8'));
  if (inadatto) throw new Error('il repository ' + dove + ' ' + inadatto + ': la pubblicazione ne sostituirebbe tutto il contenuto. Nulla è stato scritto');

  const albero = [];
  for (const [percorso, corpo] of file) {
    const blob = await api('POST', base + '/git/blobs', { content: corpo.toString('base64'), encoding: 'base64' });
    albero.push({ path: percorso, mode: '100644', type: 'blob', sha: blob.dati.sha });
  }
  // senza «base_tree»: l'albero è l'intero sito, ciò che c'era prima e non c'è più sparisce
  const tree = await api('POST', base + '/git/trees', { tree: albero });
  let cambiato = false, commit = testa;
  if (attuale.tree.sha === tree.dati.sha) dice('il repository contiene già questo sito');
  else {
    // autore neutro: l'indirizzo di posta dell'account non deve finire nella storia pubblica
    const firma = { name: 'pubblicazione del sito', email: dest.proprietario.toLowerCase() + '@users.noreply.github.com' };
    const nuovo = await api('POST', base + '/git/commits', { message: messaggio, tree: tree.dati.sha, parents: [testa], author: firma, committer: firma });
    await api('PATCH', base + '/git/refs/heads/' + ramo, { sha: nuovo.dati.sha, force: false });
    cambiato = true;
    commit = nuovo.dati.sha;
    dice('sito pubblicato nel repository (' + file.length + ' file, commit ' + commit.slice(0, 7) + ')');
  }
  const pagine = await api('POST', base + '/pages', { source: { branch: ramo, path: '/' } }, [409, 403, 404, 422]);
  if (pagine.stato < 300) dice('GitHub Pages attivato');
  else if (pagine.stato !== 409) dice('GitHub Pages: la richiesta di accensione ha risposto HTTP ' + pagine.stato);
  const acceso = await api('GET', base + '/pages', null, [404, 403]);
  if (acceso.stato !== 200) throw new Error('nel repository ' + dove + ' GitHub Pages non risulta acceso (HTTP ' + acceso.stato + '): i file ci sono, ma il sito non viene servito');
  return { cambiato, ramo, commit };
}

// Aspetta che il sito serva ESATTAMENTE i file pubblicati, confrontati per impronta. Con «opzioni.commit»
// prima aspetta che GitHub Pages dica di aver pubblicato quel commit, e solo dopo chiede i file: un file
// nuovo chiesto prima del tempo riceverebbe un 404 che la rete di GitHub tiene per dieci minuti. Se Pages
// tarda gli si chiede di ripubblicare. Restituisce { servito, secondi, diversi }.
async function attendiServito(gh, dest, file, nota, opzioni) {
  const dice = nota || (() => {});
  // la rete di GitHub può tenere una copia vecchia fino a dieci minuti: l'attesa deve superarli
  const massimo = (opzioni && opzioni.massimo) || 660000;
  const commit = opzioni && opzioni.commit;
  const base = '/repos/' + dest.proprietario + '/' + dest.nome;
  const attesi = file.filter((f) => f[0] !== '.nojekyll');
  const sw = attesi.filter((f) => f[0] === 'sw.js')[0][1];
  const servito = async (p) => {
    try { const r = await fetch(dest.sito + p + '?cb=' + Date.now() + Math.floor(Math.random() * 1000), { cache: 'no-store', redirect: 'manual' }); return r.status === 200 ? Buffer.from(await r.arrayBuffer()) : null; } catch (e) { return null; }
  };
  const confronta = async () => { const diversi = []; for (const [p, corpo] of attesi) { const x = await servito(p); if (!x || impronta(x) !== impronta(corpo)) diversi.push(p); } return diversi; };
  const t0 = Date.now(), QUANDO = [60000, 150000, 260000, 420000];
  let spinte = 0, costruito = !(gh && commit), detto = false;
  while (Date.now() - t0 < massimo) {
    if (!costruito) {
      let ultima = null;
      try { ultima = await gh('GET', base + '/pages/builds/latest'); } catch (e) { ultima = null; }
      // se la domanda a GitHub non riesce si torna al vecchio modo: prima il service worker, poi il resto
      if (!ultima || ultima.stato !== 200 || !ultima.dati) costruito = true;
      else if (ultima.dati.commit === commit && ultima.dati.status === 'built') costruito = true;
      else if (!detto) { detto = true; dice('aspetto che GitHub Pages pubblichi il commit ' + commit.slice(0, 7) + ' (ora: ' + String(ultima.dati.status) + ')'); }
    }
    if (costruito) {
      const s = await servito('sw.js');
      if (s && impronta(s) === impronta(sw)) {
        const diversi = await confronta();
        if (!diversi.length) return { servito: true, secondi: Math.round((Date.now() - t0) / 1000), diversi: [] };
        dice('il sito serve già il service worker atteso, ma ' + diversi.length + ' file non sono ancora quelli pubblicati');
      }
    }
    if (gh && spinte < QUANDO.length && Date.now() - t0 >= QUANDO[spinte]) {
      spinte++;
      try { const x = await gh('POST', base + '/pages/builds'); dice('chiesta a GitHub Pages una pubblicazione (HTTP ' + x.stato + ')'); } catch (e) { dice('richiesta di pubblicazione a GitHub Pages non riuscita: continuo ad aspettare'); }
    }
    await attendi(6000);
  }
  const diversi = costruito ? await confronta() : attesi.map((f) => f[0]);
  return { servito: !diversi.length, secondi: Math.round((Date.now() - t0) / 1000), diversi };
}

// Modifiche non registrate a ciò da cui nasce il sito (anche file nuovi, anche solo aggiunti all'indice).
function inSospeso() {
  const righe = execFileSync('git', ['-C', RADICE, 'status', '--porcelain', '--'].concat(ORIGINE_DEL_SITO), { encoding: 'utf8' });
  return righe.split('\n').filter(Boolean);
}

module.exports = { destinazione, fileDelSito, repositoryInadatto, pubblica, attendiServito, inSospeso, SOLO_CLOUDFLARE, ORIGINE_DEL_SITO };

if (require.main === module) {
  const ghApi = require('./gh-api.js');
  (async () => {
    const [nomeReparto, comando, ...resto] = process.argv.slice(2);
    const crea = resto.length === 1 && resto[0] === '--crea';
    if (!nomeReparto || ['stato', 'pubblica'].indexOf(comando) < 0 || (resto.length && !(comando === 'pubblica' && crea))) { console.log('uso: sito-pubblico.js <reparto> stato | pubblica [--crea]'); process.exitCode = 2; return; }
    const reparto = reparti.risolvi(nomeReparto);
    const dest = destinazione(reparto);
    if (!dest) throw new Error('il reparto «' + reparto.codice + '» non ha un sito pubblico (campo pubblicazione.sitoPubblico del suo file)');
    if (reparto.tipo !== 'collaudo') throw new Error('da riga di comando questo strumento lavora solo per un reparto di collaudo: per un reparto di produzione il livello 0 è un rilascio a sé');
    const AMB = ghApi.per('collaudo');
    if (AMB.UTENTE.toLowerCase() !== dest.proprietario.toLowerCase()) throw new Error('il repository ' + dest.proprietario + '/' + dest.nome + ' non è dell\'account di collaudo');
    const gh = AMB.gh, base = '/repos/' + dest.proprietario + '/' + dest.nome;

    if (comando === 'stato') {
      const r = await gh('GET', base);
      if (r.stato === 404) { console.log('repository ' + dest.proprietario + '/' + dest.nome + ': non esiste'); return; }
      if (r.stato !== 200 || !r.dati || typeof r.dati.private !== 'boolean') { console.log('ERRORE: lo stato del repository non si legge (HTTP ' + r.stato + ')'); process.exitCode = 1; return; }
      const p = r.dati.has_pages ? await gh('GET', base + '/pages') : null;
      console.log('repository ' + dest.proprietario + '/' + dest.nome + ': ' + (r.dati.private ? 'PRIVATO' : 'pubblico') + ', sito ' + (r.dati.has_pages ? 'acceso (' + String((p && p.dati && p.dati.status) || 'stato non detto') + ')' : 'spento'));
      const sw = await fetch(dest.sito + 'sw.js?cb=' + Date.now(), { cache: 'no-store' }).then(async (x) => ({ stato: x.status, testo: await x.text() }), () => ({ stato: 0, testo: '' }));
      const locale = (/consegne-v(\d+)/.exec(fs.readFileSync(path.join(RADICE, 'docs', 'sw.js'), 'utf8')) || [])[1];
      const servita = sw.stato === 200 ? ((/consegne-v(\d+)/.exec(sw.testo) || [])[1] || 'non letta') : null;
      console.log(dest.sito + ' risponde: HTTP ' + sw.stato + (servita ? ', versione ' + servita + (servita === locale ? ' (la stessa di questa cartella)' : ' (in questa cartella c\'è la ' + locale + ': per aggiornarlo, «pubblica»)') : ''));
      return;
    }

    // ── pubblica ──
    // la chiave è proprio quella dell'account che possiede il repository?
    const io = await gh('GET', '/user');
    if (io.stato !== 200 || String(io.dati && io.dati.login).toLowerCase() !== dest.proprietario.toLowerCase()) throw new Error('la chiave in uso non è quella dell\'account ' + dest.proprietario + ' (HTTP ' + io.stato + '): nulla è stato scritto');
    // nella storia pubblica finisce solo ciò che nel sorgente è registrato
    const sospese = inSospeso();
    if (sospese.length) throw new Error('ci sono modifiche non registrate ai file da cui nasce il sito (' + sospese.slice(0, 4).map((x) => x.trim()).join('; ') + (sospese.length > 4 ? '; …' : '') + '): si pubblica solo ciò che è registrato in git, perché ciò che arriva nel repository pubblico non si ritira');
    const sorgente = execFileSync('git', ['-C', RADICE, 'rev-parse', '--short', 'HEAD'], { encoding: 'utf8' }).trim();
    const { prepara } = require('./prepara-sito.js');
    const cartella = path.join(RADICE, '_sito', 'pubblico-' + reparto.codice);
    const sito = prepara(cartella, { compressa: true, reparto: reparto.codice });
    if (sito.reparto !== reparto.codice) throw new Error('il sito è stato preparato per «' + sito.reparto + '»');
    const file = fileDelSito(cartella, sito.elenco);
    const versione = (/consegne-v(\d+)/.exec(fs.readFileSync(path.join(RADICE, 'docs', 'sw.js'), 'utf8')) || [])[1];
    console.log('sito del reparto «' + reparto.codice + '» preparato: ' + file.length + ' file, versione ' + versione + ' (sorgente ' + sorgente + ')');
    const esiste = await gh('GET', base);
    if (esiste.stato === 404) {
      if (!crea) throw new Error('il repository ' + dest.proprietario + '/' + dest.nome + ' non esiste. Per crearlo, pubblico: «pubblica --crea»');
      const c = await gh('POST', '/user/repos', { name: dest.nome, description: 'Sito pubblicato: solo i file serviti, in copia compressa', private: false, auto_init: true, has_issues: false, has_wiki: false, has_projects: false });
      if (c.stato >= 300) throw new Error('creazione del repository non riuscita: HTTP ' + c.stato);
      console.log('repository ' + dest.proprietario + '/' + dest.nome + ' creato');
      await attendi(3000);
    } else if (esiste.stato !== 200) throw new Error('lo stato del repository non si legge (HTTP ' + esiste.stato + '): nulla è stato scritto');
    const esitoPub = await pubblica(gh, dest, file, 'Sito del reparto ' + reparto.codice + ', versione ' + versione + ' (sorgente ' + sorgente + ')', (t) => console.log(t));
    const esito = await attendiServito(gh, dest, file, (t) => console.log(t), { commit: esitoPub.commit });
    console.log(esito.servito ? 'pronto dopo ' + esito.secondi + ' s: ' + dest.sito + ' serve i ' + (file.length - 1) + ' file pubblicati, identici alla copia compressa' : 'NON ancora servito dopo ' + esito.secondi + ' s; file diversi: ' + esito.diversi.slice(0, 6).join(', '));
    if (!esito.servito) process.exitCode = 1;
  })().catch((e) => { console.log('ERRORE: ' + ghApi.maschera(String((e && e.message) || e)).replace(/\/\/[^/@\s]+@/g, '//***@')); process.exitCode = 1; });
}
