// Lo «scambio dei repository», provato nel collaudo (account gistech2026) e
// pronto per la produzione (account medicinadurgenzaucsc-maker).
//
// Prima dello scambio:  app-consegne            pubblico, sorgente + sito su GitHub Pages
// Dopo lo scambio:      app-consegne-sorgente   PRIVATO, il sorgente (storia, segreti, flussi)
//                       app-consegne            pubblico, SOLO la pagina di rinvio al sito nuovo
// Il repository della pagina di rinvio si prepara prima, con un altro nome
// (app-consegne-rinvio) e il sito già pubblicato: allo scambio bastano due
// cambi di nome, e si misura per quanti secondi il vecchio indirizzo non risponde.
//
//   node collaudo/strumenti/scambio-repository.js [ambiente] stato
//   node collaudo/strumenti/scambio-repository.js <ambiente> prepara <nuovo indirizzo>   es. https://consegne-collaudo.pages.dev/
//   node collaudo/strumenti/scambio-repository.js <ambiente> scambia
//   node collaudo/strumenti/scambio-repository.js <ambiente> annulla             rimette tutto com'era
//
// «ambiente» è collaudo oppure produzione, e va scritto per primo. Solo «stato»
// lo può sottintendere (vale collaudo): i comandi che cambiano qualcosa lo
// vogliono scritto, così un comando copiato senza ambiente non lavora sul
// collaudo al posto della produzione. In produzione «stato» legge soltanto; gli
// altri comandi vogliono anche «--confermo-produzione», e si lanciano solo con
// l'ok di chi gestisce l'app. Il remoto di questa cartella è «collaudo» per il
// collaudo e «origin» per la produzione.
//
// «scambia» e «annulla» si possono RILANCIARE: guardano a che punto sono i
// repository e fanno solo i passi che mancano. Se uno dei due si interrompe a
// metà, basta rilanciarlo (oppure lanciare l'altro per tornare indietro). Se
// qualcosa non è andato fino in fondo l'ultima riga comincia con «ATTENZIONE»
// e lo strumento esce con errore.
//
// «scambia» parte solo se: nessun fork, nessun collaboratore oltre al
// proprietario, nessuna chiave di pubblicazione, nessun aggancio esterno (un
// elenco che GitHub non lascia leggere vale come controllo fallito); la pagina
// di rinvio porta all'indirizzo del sito nuovo di quell'ambiente; l'ultima
// pubblicazione su Cloudflare è riuscita e contiene il sito com'è oggi nel ramo
// master (dietro il cancello di accesso il sito nuovo risponde allo stesso modo
// anche se è vuoto: lo si sa solo dall'esito del flusso).
//
// PRIMA dello scambio ogni dispositivo che usa l'app va portato sul nuovo
// indirizzo e fatto entrare lì: dopo lo scambio il vecchio indirizzo rimanda al
// nuovo, dove la sessione non c'è, e chi ricarica (anche da solo, alle 04:00) si
// ritrova sulla schermata di accesso. Lo strumento questo non lo può verificare.
//
// Dopo lo scambio il remoto di questa cartella punta al repository privato.
// ATTENZIONE alle altre copie della cartella: il loro remoto punterebbe ancora
// al vecchio nome, che ora è il repository pubblico del rinvio. Per questo quel
// repository viene chiuso alle scritture (regola «sola-lettura»).
//
// GitHub Pages pubblica al massimo una decina di volte l'ora per sito: dopo
// troppi scambi e annullamenti di fila (succede solo provando) le pubblicazioni
// vanno in errore e il vecchio indirizzo resta a 404 finché l'ora non passa.
const { execFileSync } = require('child_process');
const ghApi = require('./gh-api.js');
const { fileRinvio, destinazione, RADICE } = require('./prepara-rinvio.js');

// Nessun messaggio deve contenere una chiave: né quelle di GitHub né ciò che
// in un indirizzo sta fra «//» e «@».
const pulito = (t) => ghApi.maschera(String(t)).replace(/\/\/[^/@\s]+@/g, '//***@');

let ARG, AMB;
try { ARG = ghApi.argomenti(process.argv.slice(2)); AMB = ghApi.per(ARG.ambiente); }
catch (e) { console.log('ERRORE: ' + pulito(e.message)); process.exit(1); }
const gh = AMB.gh, UTENTE = AMB.UTENTE;

const NOMI = { app: 'app-consegne', sorgente: 'app-consegne-sorgente', rinvio: 'app-consegne-rinvio' };
const SITO = 'https://' + UTENTE.toLowerCase() + '.github.io/';
const REMOTO = AMB.remoto;
const NUOVO = AMB.sitoNuovo;
// Due regole: una per i rami e una per le etichette. Con la sola regola dei rami
// un «git push --tags» porterebbe nel repository pubblico tutta la storia che le
// etichette raggiungono.
const REGOLE = [
  { name: 'sola-lettura', target: 'branch', rules: [{ type: 'creation' }, { type: 'update' }, { type: 'deletion' }, { type: 'non_fast_forward' }] },
  { name: 'sola-lettura-etichette', target: 'tag', rules: [{ type: 'creation' }, { type: 'update' }, { type: 'deletion' }] },
];
const FLUSSO = 'pubblica-cloudflare.yml';
// Se cambia uno di questi, il sito pubblicato su Cloudflare non è più quello del ramo.
const DEL_SITO = /^(docs\/|[.]github\/compressione\/|[.]github\/pubblicazione\/|[.]github\/workflows\/pubblica-cloudflare[.]yml$|collaudo\/strumenti\/prepara-sito[.]js$)/;
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
// Che cosa risponde un indirizzo (l'app, la pagina di rinvio, un 404, altro) e,
// se è la pagina di rinvio, verso quale indirizzo porta.
async function doveRimanda(url) {
  try {
    const r = await fetch(url + (url.indexOf('?') < 0 ? '?' : '&') + 'cb=' + Date.now(), { redirect: 'manual', cache: 'no-store' });
    if (r.status === 404) return { cosa: '404', verso: null };
    if (r.status >= 300) return { cosa: 'HTTP ' + r.status, verso: null };
    const t = await r.text();
    if (t.indexOf('navVersioneApp') >= 0) return { cosa: 'app', verso: null };
    if (t.indexOf('ha cambiato indirizzo') < 0) return { cosa: 'altro', verso: null };
    const m = /var NUOVO = '([^']*)'/.exec(t);
    return { cosa: 'rinvio', verso: m ? m[1] : null };
  } catch (e) { return { cosa: 'errore di rete', verso: null }; }
}
const cheCosa = async (url) => (await doveRimanda(url)).cosa;
// Le regole «sola-lettura» presenti nel repository (al più due).
async function regolePresenti(nome) {
  // 403: in un repository privato di un account gratuito le regole non esistono
  const r = await api('GET', base(nome) + '/rulesets', null, [404, 403]);
  const nomi = REGOLE.map((x) => x.name);
  return (Array.isArray(r.dati) ? r.dati : []).filter((x) => nomi.indexOf(x.name) >= 0);
}
// «Chiuso alle scritture» vuol dire tutte e due le regole.
async function regola(nome) {
  const presenti = (await regolePresenti(nome)).map((x) => x.name);
  return REGOLE.every((x) => presenti.indexOf(x.name) >= 0);
}
async function sblocca(nome) {
  for (const x of await regolePresenti(nome)) await api('DELETE', base(nome) + '/rulesets/' + x.id);
}
// Nessuno può più scrivere nei rami e nelle etichette del repository del
// rinvio, nemmeno il proprietario: un push fatto per sbaglio col sorgente
// viene rifiutato. Si può ripetere: crea solo le regole che mancano.
async function blocca(nome) {
  const presenti = (await regolePresenti(nome)).map((x) => x.name);
  let create = 0;
  for (const x of REGOLE) {
    if (presenti.indexOf(x.name) >= 0) continue;
    await api('POST', base(nome) + '/rulesets', { name: x.name, target: x.target, enforcement: 'active', bypass_actors: [], conditions: { ref_name: { include: ['~ALL'], exclude: [] } }, rules: x.rules });
    create++;
  }
  return create > 0;
}

// Il remoto locale. Il suo indirizzo può contenere una chiave: non finisce mai
// in un messaggio, nemmeno quando git fallisce (l'errore di git riporterebbe
// l'intera riga di comando). Deve essere un repository di questo account,
// nella forma github.com/<proprietario>/<nome>: altrimenti non si tocca nulla.
const FORMA_REMOTO = /(github\.com[/:])([^/:@\s]+)\/([A-Za-z0-9._-]+?)(?:\.git)?\/?$/;
function remoto() {
  let u = '';
  try { u = execFileSync('git', ['-C', RADICE, 'remote', 'get-url', REMOTO], { encoding: 'utf8', stdio: ['ignore', 'pipe', 'pipe'] }).trim(); }
  catch (e) { throw new Error('non riesco a leggere il remoto «' + REMOTO + '» di questa cartella'); }
  const m = FORMA_REMOTO.exec(u);
  if (!m || m[2].toLowerCase() !== UTENTE.toLowerCase()) throw new Error('il remoto «' + REMOTO + '» di questa cartella non punta a un repository di ' + UTENTE + ': va corretto a mano');
  return { indirizzo: u, nome: m[3] };
}
function nomeDelRemoto() { return remoto().nome; }
function puntaIlRemoto(nome) {
  const r = remoto();
  if (r.nome === nome) return false;
  const nuovo = r.indirizzo.replace(FORMA_REMOTO, (tutto, sito, proprietario) => sito + proprietario + '/' + nome + '.git');
  try { execFileSync('git', ['-C', RADICE, 'remote', 'set-url', REMOTO, nuovo], { stdio: ['ignore', 'pipe', 'pipe'] }); }
  catch (e) { throw new Error('non riesco a spostare il remoto «' + REMOTO + '» su ' + nome + ': va corretto a mano prima di ogni push'); }
  return true;
}

// A che punto sono i tre nomi.
//   prima      il sorgente sta ancora al vecchio nome, il rinvio è pronto col nome provvisorio
//   a-meta-1   fatta solo la prima rinomina: il vecchio nome non è di nessuno
//   a-meta-2   fatte le due rinomine, ma il sorgente è ancora pubblico
//   fatto      scambio concluso
async function situazione() {
  const app = await repo(NOMI.app), sorgente = await repo(NOMI.sorgente), rinvio = await repo(NOMI.rinvio);
  const haIlSorgente = async (nome) => (await api('GET', base(nome) + '/contents/docs/sw.js', null, [404])).stato === 200;
  const appE = !app ? null : (await haIlSorgente(NOMI.app)) ? 'sorgente' : 'rinvio';
  let fase = 'inattesa';
  if (appE === 'sorgente' && !sorgente && rinvio) fase = 'prima';
  else if (appE === 'sorgente' && !sorgente && !rinvio) fase = 'senza-rinvio';
  else if (!app && sorgente && rinvio) fase = 'a-meta-1';
  else if (appE === 'rinvio' && sorgente && !rinvio) fase = sorgente.private ? 'fatto' : 'a-meta-2';
  return { app, sorgente, rinvio, appE, fase };
}
const FASI = {
  prima: 'scambio non fatto: il sorgente sta al vecchio nome, la pagina di rinvio è pronta',
  'senza-rinvio': 'scambio non fatto, e la pagina di rinvio non è ancora stata preparata («prepara»)',
  'a-meta-1': 'scambio A METÀ: fatta solo la prima rinomina, il vecchio nome non è di nessuno. Rilanciare «scambia» per finire, o «annulla» per tornare indietro',
  'a-meta-2': 'scambio A METÀ: fatte le due rinomine ma il sorgente è ancora PUBBLICO. Rilanciare «scambia» per finire, o «annulla» per tornare indietro',
  fatto: 'scambio fatto',
  inattesa: 'situazione INATTESA: i tre repository non sono in nessuno degli stati previsti, non tocco nulla',
};

// Ciò che renderebbe inutile lo scambio: copie pubbliche o altri accessi al sorgente.
// Un elenco che GitHub non lascia leggere NON vale «nessuno»: è un controllo fallito.
async function controlliDiPartenza(nome) {
  const guai = [];
  const r = (await api('GET', base(nome))).dati;
  if (r.forks_count > 0) guai.push(r.forks_count + ' fork (un fork pubblico resta pubblico anche dopo lo scambio)');
  const elenco = async (percorso, cosa) => {
    const x = await gh('GET', base(nome) + percorso);
    if (x.stato !== 200 || !Array.isArray(x.dati)) { guai.push('non riesco a leggere ' + cosa + ' (HTTP ' + x.stato + '): con questa chiave il controllo non si può fare'); return []; }
    return x.dati;
  };
  const altri = (await elenco('/collaborators?affiliation=all&per_page=100', 'i collaboratori')).filter((x) => String(x.login).toLowerCase() !== UTENTE.toLowerCase());
  if (altri.length) guai.push('collaboratori oltre al proprietario: ' + altri.map((x) => x.login).join(', '));
  const chiavi = await elenco('/keys', 'le chiavi di pubblicazione');
  if (chiavi.length) guai.push(chiavi.length + ' chiavi di pubblicazione');
  const agganci = await elenco('/hooks', 'gli agganci esterni');
  if (agganci.length) guai.push(agganci.length + ' agganci esterni');
  const inviti = await elenco('/invitations', 'gli inviti');
  if (inviti.length) guai.push(inviti.length + ' inviti in sospeso');
  return guai;
}

// Il sito su Cloudflare contiene davvero ciò che sta nel ramo master? Dietro il
// cancello di accesso il sito risponde 302 anche se è vuoto: lo dice solo il flusso.
async function pubblicazione(nome) {
  const b = await api('GET', base(nome) + '/branches/master', null, [404]);
  const cima = b.dati && b.dati.commit && b.dati.commit.sha;
  if (!cima) return { ok: false, motivo: 'ramo master non trovato' };
  const r = await api('GET', base(nome) + '/actions/workflows/' + FLUSSO + '/runs?per_page=10', null, [404]);
  if (r.stato === 404) return { ok: false, motivo: 'nel repository non c\'è ancora il flusso ' + FLUSSO };
  const riuscite = ((r.dati && r.dati.workflow_runs) || []).filter((x) => x.status === 'completed' && x.conclusion === 'success');
  for (const run of riuscite) {
    const j = await api('GET', base(nome) + '/actions/runs/' + run.id + '/jobs');
    const passi = [].concat(...((j.dati && j.dati.jobs) || []).map((x) => x.steps || []));
    if (!passi.some((p) => p.name === 'Pubblica' && p.conclusion === 'success')) continue;   // un giro che ha solo controllato i file
    if (run.head_sha === cima) return { ok: true, commit: run.head_sha, numero: run.run_number };
    const c = await api('GET', base(nome) + '/compare/' + run.head_sha + '...' + cima, null, [404]);
    if (c.stato === 404) return { ok: false, motivo: 'non riesco a confrontare l\'ultima pubblicazione con il ramo master' };
    // il ramo deve DISCENDERE dal commit pubblicato: se è rimasto indietro o ha preso un'altra strada, su Cloudflare c'è altro
    const rapporto = c.dati && c.dati.status;
    if (rapporto !== 'ahead' && rapporto !== 'identical') return { ok: false, motivo: 'il ramo master non discende dal commit pubblicato (n. ' + run.run_number + ', confronto: ' + rapporto + '): ripubblicare' };
    const file = (c.dati && c.dati.files) || [];
    if (file.length >= 300) return { ok: false, motivo: 'troppe differenze dall\'ultima pubblicazione per verificarle: ripubblicare' };
    const toccati = file.map((f) => f.filename).filter((f) => DEL_SITO.test(f));
    if (toccati.length) return { ok: false, motivo: 'dopo l\'ultima pubblicazione riuscita (n. ' + run.run_number + ') il sito è cambiato nel ramo master: ' + toccati.slice(0, 4).join(', ') };
    return { ok: true, commit: run.head_sha, numero: run.run_number };
  }
  return { ok: false, motivo: 'nessuna pubblicazione su Cloudflare riuscita fra le ultime dieci esecuzioni del flusso' };
}

async function stato() {
  for (const k of Object.keys(NOMI)) {
    const r = await repo(NOMI[k]);
    if (!r) { console.log(NOMI[k].padEnd(24) + 'non esiste'); continue; }
    const p = r.has_pages ? (await api('GET', base(NOMI[k]) + '/pages', null, [404])).dati : null;
    const quante = (await regolePresenti(NOMI[k])).length;
    console.log(NOMI[k].padEnd(24) + r.visibility + ' | sito: ' + (p && p.source ? p.source.branch + ':' + p.source.path + ' (' + p.status + ')' : 'no') + (quante === REGOLE.length ? ' | chiuso alle scritture' : quante ? ' | chiuso solo in parte' : ''));
  }
  const rinvii = [];
  for (const nome of [NOMI.app, NOMI.rinvio]) {
    const d = await doveRimanda(SITO + nome + '/');
    if (d.cosa === 'rinvio') rinvii.push(d.verso);
    console.log(('indirizzo ' + SITO + nome + '/').padEnd(78) + ' → ' + d.cosa + (d.cosa === 'rinvio' ? ' verso ' + d.verso : ''));
  }
  try { console.log('remoto «' + REMOTO + '» di questa cartella → ' + UTENTE + '/' + pulito(nomeDelRemoto())); }
  catch (e) { console.log('remoto «' + REMOTO + '» di questa cartella → DA SISTEMARE: ' + pulito(e.message)); }
  const s = await situazione();
  console.log('situazione: ' + FASI[s.fase]);
  const conIlRinvio = s.appE === 'rinvio' ? NOMI.app : s.rinvio ? NOMI.rinvio : null;
  if (conIlRinvio) {
    const giusta = rinvii.length > 0 && rinvii.every((x) => x === NUOVO);
    const quante = (await regolePresenti(conIlRinvio)).length;
    console.log('pagina di rinvio: ' + (giusta ? 'porta al sito nuovo di ' + ARG.ambiente + ' (' + NUOVO + ')' : 'DA SISTEMARE: deve portare a ' + NUOVO + (rinvii.length ? ', porta a ' + rinvii.join(', ') : ', e ora non è in linea'))
      + (quante === REGOLE.length ? '; repository chiuso alle scritture' : quante ? '; repository chiuso solo in parte: manca una delle due regole, rami ed etichette (la mette «scambia», oppure «prepara» rilanciato)' : '; DA SISTEMARE: il repository NON è chiuso alle scritture'));
  }
  const conIlSorgente = s.appE === 'sorgente' ? NOMI.app : s.sorgente ? NOMI.sorgente : null;
  if (conIlSorgente) {
    const guai = await controlliDiPartenza(conIlSorgente);
    console.log('accessi al sorgente: ' + (guai.length ? 'DA SISTEMARE: ' + guai.join('; ') : 'solo il proprietario, nessun fork, nessuna chiave di pubblicazione, nessun aggancio esterno'));
    const pub = await pubblicazione(conIlSorgente);
    console.log('sito su Cloudflare: ' + (pub.ok ? 'pubblicato dal commit ' + pub.commit.slice(0, 7) + ' (esecuzione n. ' + pub.numero + '), allineato al ramo master' : 'NON PRONTO: ' + pub.motivo));
  }
}

// Crea (o aggiorna) il repository della pagina di rinvio, col suo sito già in linea.
async function prepara(nuovo) {
  const dest = destinazione(nuovo);
  if (dest !== NUOVO) throw new Error('per ' + ARG.ambiente + ' la pagina di rinvio deve portare a ' + NUOVO + ', non a ' + dest);
  const s = await situazione();
  if (s.fase !== 'prima' && s.fase !== 'senza-rinvio') throw new Error('non si può preparare ora. ' + FASI[s.fase]);
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
  // Se da qui in poi qualcosa si interrompe il repository resta aperto alle
  // scritture: «stato» lo dice, e sia «prepara» sia «scambia» lo richiudono.
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
  const pagine = await api('POST', b + '/pages', { source: { branch: ramo, path: '/' } }, [409]);
  console.log(pagine.stato < 300 ? 'GitHub Pages attivato' : 'GitHub Pages già attivo');
  await blocca(NOMI.rinvio);
  const t0 = Date.now();
  let visto = { cosa: '', verso: null };
  for (let i = 0; i < 60; i++) { visto = await doveRimanda(SITO + NOMI.rinvio + '/'); if (visto.cosa === 'rinvio' && visto.verso === dest) break; await attendi(5000); }
  if (visto.cosa !== 'rinvio' || visto.verso !== dest) throw new Error('dopo ' + secondi(t0) + ' il sito del rinvio risponde ancora «' + visto.cosa + (visto.verso ? ' verso ' + visto.verso : '') + '»');
  console.log('pronto: ' + SITO + NOMI.rinvio + '/ rimanda a ' + dest + ' (in linea dopo ' + secondi(t0) + '); repository chiuso alle scritture');
}

// Aspetta che un indirizzo risponda nel modo voluto per tre volte di fila, annotando i cambi.
// Se si aspetta la pagina di rinvio, conta buona solo quella che porta al sito nuovo di
// questo ambiente: la destinazione si legge nella stessa risposta, senza richieste a parte.
// «pronto», se c'è, è una condizione in più da verificare quando la risposta è quella voluta.
async function finche(url, voluto, nota, t0, repoDelSito, pronto) {
  let ultimo = '', buone = 0, primo = null, nonTrovato = 0, spinte = 0, verso = null, prontoVisto = !pronto;
  const QUANDO = [40000, 130000, 220000];
  for (let i = 0; i < 240 && buone < 3; i++) {
    const d = await doveRimanda(url);
    const c = d.cosa === 'rinvio' ? 'rinvio verso ' + d.verso : d.cosa;
    if (d.cosa === 'rinvio') verso = d.verso;
    if (c !== ultimo) { nota('il vecchio indirizzo risponde: ' + c); ultimo = c; }
    let buona = d.cosa === voluto && (voluto !== 'rinvio' || d.verso === NUOVO);
    if (buona && !prontoVisto) {
      // una copia vecchia servita per qualche secondo non basta: serve anche la conferma
      try { prontoVisto = await pronto(); } catch (e) { prontoVisto = false; }
      if (!prontoVisto) buona = false;
    }
    if (buona) { buone++; if (primo == null) primo = Date.now() - t0; } else { buone = 0; primo = null; }
    if (d.cosa === '404') nonTrovato++;
    if (!buona && spinte < QUANDO.length && Date.now() - t0 > QUANDO[spinte]) {
      spinte++;
      // è solo una spinta: se non riesce si continua ad aspettare
      try {
        const x = await gh('POST', base(repoDelSito) + '/pages/builds');
        nota('dopo ' + Math.round(QUANDO[spinte - 1] / 1000) + ' s chiesta a GitHub Pages una nuova pubblicazione (HTTP ' + x.stato + ')');
      } catch (e) { nota('richiesta di nuova pubblicazione a GitHub Pages non riuscita: continuo ad aspettare'); }
    }
    await attendi(1500);
  }
  return { riuscito: buone >= 3, primo: primo, nonTrovato: nonTrovato, verso: verso, ultimo: ultimo };
}

// Ciò che non è andato fino in fondo: lo si ripete in fondo, e lo strumento esce con errore.
const avvisi = [];
function spostaIlRemoto(nome, nota) {
  try { if (puntaIlRemoto(nome)) nota('remoto «' + REMOTO + '» puntato su ' + nome); }
  catch (e) { nota('ATTENZIONE: ' + e.message); if (avvisi.indexOf(e.message) < 0) avvisi.push(e.message); }
}
function chiudi(riuscito, riga) {
  console.log(riga);
  avvisi.forEach((a) => console.log('ATTENZIONE: ' + a));
  if (!riuscito || avvisi.length) process.exitCode = 1;
}

async function scambia() {
  let fase = (await situazione()).fase;
  const t0 = Date.now();
  const nota = (t) => console.log(secondi(t0).padStart(8) + '  ' + t);
  if (fase === 'fatto') {
    if (await blocca(NOMI.app)) nota('repository della pagina di rinvio chiuso alle scritture: non lo era');
    spostaIlRemoto(NOMI.sorgente, nota);
    // i repository sono a posto: resta da vedere se il vecchio indirizzo serve già la pagina di rinvio
    const gia = await doveRimanda(SITO + NOMI.app + '/');
    if (gia.cosa === 'rinvio' && gia.verso === NUOVO) return chiudi(true, 'Lo scambio è già fatto e il vecchio indirizzo rimanda a ' + NUOVO + ': nulla da fare.');
    nota('lo scambio è fatto, ma il vecchio indirizzo non serve ancora la pagina di rinvio giusta: la aspetto');
    return codaDelloScambio(nota, t0);
  }
  if (fase === 'prima') {
    if (nomeDelRemoto() !== NOMI.app) throw new Error('il remoto «' + REMOTO + '» punta a ' + nomeDelRemoto() + ', atteso ' + NOMI.app);
    const rinvio = await doveRimanda(SITO + NOMI.rinvio + '/');
    if (rinvio.cosa !== 'rinvio') throw new Error('il sito del rinvio non è in linea');
    if (rinvio.verso !== NUOVO) throw new Error('la pagina di rinvio porta a ' + rinvio.verso + ' invece che a ' + NUOVO + ': rilanciare «prepara ' + NUOVO + '»');
    if (await cheCosa(SITO + NOMI.app + '/') !== 'app') throw new Error("il vecchio indirizzo non sta servendo l'applicazione");
    const guai = await controlliDiPartenza(NOMI.app);
    if (guai.length) throw new Error('prima dello scambio vanno sistemati: ' + guai.join('; '));
    const pub = await pubblicazione(NOMI.app);
    if (!pub.ok) throw new Error('il sito nuovo non è pronto: ' + pub.motivo);
    if (await blocca(NOMI.rinvio)) nota('repository della pagina di rinvio chiuso alle scritture: non lo era');
    nota('controlli di partenza superati; sito su Cloudflare pubblicato dal commit ' + pub.commit.slice(0, 7));
    await api('PATCH', base(NOMI.app), { name: NOMI.sorgente });
    nota(NOMI.app + ' rinominato in ' + NOMI.sorgente);
    fase = 'a-meta-1';
  } else if (fase === 'a-meta-1' || fase === 'a-meta-2') {
    nota('riprendo uno scambio rimasto a metà');
  } else throw new Error(FASI[fase]);
  // Il sorgente ha già il nome nuovo: il remoto si sposta SUBITO, prima delle
  // chiamate che possono ancora fallire, così un push non finisce mai nel
  // repository pubblico del rinvio.
  spostaIlRemoto(NOMI.sorgente, nota);
  if (fase === 'a-meta-1') {
    await api('PATCH', base(NOMI.rinvio), { name: NOMI.app });
    nota(NOMI.rinvio + ' rinominato in ' + NOMI.app);
  }
  // subito privato: il vecchio indirizzo ormai è del repository del rinvio
  await api('PATCH', base(NOMI.sorgente), { private: true });
  nota(NOMI.sorgente + ' reso privato');
  if (await blocca(NOMI.app)) nota('repository della pagina di rinvio chiuso alle scritture: non lo era');
  return codaDelloScambio(nota, t0);
}
// L'ultimo passo dello scambio: aspettare che il vecchio indirizzo serva la pagina di
// rinvio che porta al sito nuovo. Vale anche per uno «scambia» rilanciato a scambio fatto.
async function codaDelloScambio(nota, t0) {
  const esito = await finche(SITO + NOMI.app + '/', 'rinvio', nota, t0, NOMI.app);
  if (esito.riuscito) return chiudi(true, 'FATTO: il vecchio indirizzo rimanda a ' + NUOVO + ' dopo ' + (esito.primo / 1000).toFixed(1) + ' s dall\'inizio; risposte «404» viste nel frattempo: ' + esito.nonTrovato);
  if (/^rinvio verso /.test(esito.ultimo) && esito.verso !== NUOVO) return chiudi(false, 'ATTENZIONE: lo scambio è fatto, ma il vecchio indirizzo rimanda a ' + esito.verso + ' invece che a ' + NUOVO + ': la pagina di rinvio va rifatta');
  chiudi(false, 'ATTENZIONE: dopo 6 minuti il vecchio indirizzo non serve ancora la pagina di rinvio (ultima risposta: ' + esito.ultimo + '). I repository sono a posto: manca solo la pubblicazione di GitHub Pages, e «stato» ne mostra l\'esito. Rilanciare «scambia» la richiede di nuovo');
}

// Il sorgente è al vecchio nome: perché il vecchio indirizzo serva l'applicazione
// deve essere pubblico e avere il sito acceso. Si può ripetere.
async function riaccendiIlSito(nota) {
  const r = await repo(NOMI.app);
  if (!r) throw new Error('il repository ' + NOMI.app + ' non si trova');
  if (r.private) { await api('PATCH', base(NOMI.app), { private: false }); nota(NOMI.app + ' di nuovo pubblico'); }
  const p = await api('GET', base(NOMI.app) + '/pages', null, [404]);
  if (p.stato === 404) {
    const x = await api('POST', base(NOMI.app) + '/pages', { source: { branch: 'master', path: '/docs' } }, [409]);
    nota('GitHub Pages sul sorgente: ' + (x.stato < 300 ? 'riattivato' : 'già attivo'));
    return { riacceso: true, tornatoPubblico: !!r.private };
  }
  nota('GitHub Pages sul sorgente: già attivo (' + (p.dati && p.dati.status) + ')');
  return { riacceso: false, tornatoPubblico: !!r.private };
}
// La pubblicazione del sito del sorgente è conclusa? Subito dopo una rinomina il vecchio
// indirizzo può servire per qualche decina di secondi una copia di prima: non fa testo.
async function sitoPubblicato() {
  const p = await api('GET', base(NOMI.app) + '/pages', null, [404]);
  return p.stato === 200 && !!p.dati && p.dati.status === 'built';
}

async function annulla() {
  const s = await situazione();
  const t0 = Date.now();
  const nota = (t) => console.log(secondi(t0).padStart(8) + '  ' + t);
  if (s.fase === 'prima' || s.fase === 'senza-rinvio') {
    spostaIlRemoto(NOMI.app, nota);
    // Un annullamento interrotto dopo l'ultima rinomina lascia i nomi a posto e il sito
    // spento. Non ci si fida di ciò che il vecchio indirizzo serve in questo momento (può
    // essere una copia di prima): si guarda se il sorgente è pubblico e se il sito è acceso.
    const acceso = await riaccendiIlSito(nota);
    if (!acceso.riacceso && !acceso.tornatoPubblico && await sitoPubblicato() && await cheCosa(SITO + NOMI.app + '/') === 'app') {
      return chiudi(true, "Lo scambio non è fatto e il vecchio indirizzo serve l'applicazione: nulla da annullare.");
    }
    nota("lo scambio non è fatto, ma il sito del vecchio indirizzo non è a posto: aspetto che torni a servire l'applicazione");
    const ripresa = await finche(SITO + NOMI.app + '/', 'app', nota, t0, NOMI.app, sitoPubblicato);
    return chiudi(ripresa.riuscito, ripresa.riuscito
      ? "TORNATO COM'ERA: il vecchio indirizzo serve di nuovo l'applicazione dopo " + (ripresa.primo / 1000).toFixed(1) + ' s'
      : "ATTENZIONE: dopo 6 minuti il vecchio indirizzo non serve ancora l'applicazione (ultima risposta: " + ripresa.ultimo + "). I repository sono com'erano: manca la pubblicazione di GitHub Pages («stato» ne mostra l'esito). Rilanciare «annulla» la richiede di nuovo");
  } else if (s.fase === 'a-meta-1') {
    nota('era stata fatta solo la prima rinomina: la disfo');
    await api('PATCH', base(NOMI.sorgente), { name: NOMI.app });
    nota(NOMI.sorgente + ' rinominato in ' + NOMI.app);
  } else if (s.fase === 'a-meta-2' || s.fase === 'fatto') {
    if (s.sorgente.private) { await api('PATCH', base(NOMI.sorgente), { private: false }); nota(NOMI.sorgente + ' di nuovo pubblico'); }
    await api('PATCH', base(NOMI.app), { name: NOMI.rinvio });
    nota(NOMI.app + ' rinominato in ' + NOMI.rinvio);
    await api('PATCH', base(NOMI.sorgente), { name: NOMI.app });
    nota(NOMI.sorgente + ' rinominato in ' + NOMI.app);
  } else throw new Error(FASI[s.fase]);
  spostaIlRemoto(NOMI.app, nota);
  await riaccendiIlSito(nota);
  // «app» conta solo a pubblicazione conclusa: subito dopo le rinomine il vecchio indirizzo può servire una copia di prima
  const esito = await finche(SITO + NOMI.app + '/', 'app', nota, t0, NOMI.app, sitoPubblicato);
  chiudi(esito.riuscito, esito.riuscito
    ? "TORNATO COM'ERA: il vecchio indirizzo serve di nuovo l'applicazione dopo " + (esito.primo / 1000).toFixed(1) + ' s'
    : "ATTENZIONE: dopo 6 minuti il vecchio indirizzo non serve ancora l'applicazione (ultima risposta: " + esito.ultimo + "). I repository sono tornati com'erano: manca la pubblicazione di GitHub Pages («stato» ne mostra l'esito). Rilanciare «annulla» la richiede di nuovo");
}

const AZIONI = { stato: [stato, 0], prepara: [prepara, 1], scambia: [scambia, 0], annulla: [annulla, 0] };
const comando = ARG.resto[0] || 'stato';
const voce = Object.prototype.hasOwnProperty.call(AZIONI, comando) ? AZIONI[comando] : null;
if (!voce || ARG.resto.length - 1 > voce[1] || (comando === 'prepara' && ARG.resto.length !== 2)) {
  console.log('uso: scambio-repository.js [collaudo|produzione] stato | <collaudo|produzione> prepara <nuovo indirizzo> | scambia | annulla');
  process.exit(1);
}
if (comando !== 'stato') {
  try { ghApi.pretendiAmbiente(ARG, comando); }
  catch (e) { console.log('ERRORE: ' + pulito(e.message)); process.exit(1); }
}
console.log('ambiente: ' + ARG.ambiente + ' | proprietario su GitHub: ' + UTENTE + ' | remoto di questa cartella: ' + REMOTO);
if (AMB.produzione && comando !== 'stato' && !ARG.confermato) {
  console.log('In PRODUZIONE «' + comando + '» cambia i repository del reparto: va lanciato con ' + ghApi.CONFERMA + ', e solo con l\'ok di chi gestisce l\'app.');
  process.exit(1);
}
Promise.resolve(ARG.resto[1]).then(voce[0]).catch((e) => {
  console.log('ERRORE: ' + pulito(e && e.message || e));
  avvisi.forEach((a) => console.log('ATTENZIONE: ' + a));
  console.log('Per vedere a che punto si è rimasti: «stato». «scambia» e «annulla» si possono rilanciare.');
  process.exit(1);
});
