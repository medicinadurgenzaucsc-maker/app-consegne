// Dove il sito sta su Cloudflare lo pubblica il flusso «Pubblica su Cloudflare»
// del repository, a ogni push su master che tocca il sito; con la variabile
// CLOUDFLARE_AVVISA = si è lo stesso flusso ad aggiornare app_version, cioè a far
// comparire «Update» sulle pagine aperte. Finché i PC usano GitHub Pages quella
// variabile è spenta e l'avviso lo dà il flusso «notify-deploy», quando GitHub
// Pages ha finito di pubblicare. Questo strumento segue il flusso di Cloudflare
// fino alla fine e poi aspetta che l'avviso arrivi nel database, da qualunque
// dei due venga. Non scrive nulla nel database.
//
//   node collaudo/strumenti/pubblicazione.js [ambiente]          segue la pubblicazione del commit appena mandato
//                                                                (se quel commit non tocca il sito, mostra l'ultima fatta)
//   node collaudo/strumenti/pubblicazione.js [ambiente] avvia    la fa ripartire a mano (stesso commit) e la segue
//
// «ambiente» è collaudo oppure produzione. Senza «avvia» lo strumento legge
// soltanto e l'ambiente si può sottintendere (vale collaudo); «avvia» lo vuole
// scritto e, in produzione, anche «--confermo-produzione». Il repository è quello a cui punta il
// remoto dell'ambiente in questa cartella («collaudo» oppure «origin»): così lo
// strumento continua a funzionare se il repository cambia nome.
const { execFileSync } = require('child_process');
const path = require('path');
const ghApi = require('./gh-api.js');
const sb = require('./sb.js');

let ARG, AMB;
try { ARG = ghApi.argomenti(process.argv.slice(2)); AMB = ghApi.per(ARG.ambiente); }
catch (e) { console.log('ERRORE: ' + ghApi.maschera(e.message)); process.exit(1); }
const gh = AMB.gh, UTENTE = AMB.UTENTE;
const RADICE = path.resolve(__dirname, '../..');
const FLUSSO = 'pubblica-cloudflare.yml';
// Se cambia uno di questi, il push fa partire il flusso (sono i percorsi scritti nel flusso stesso).
const DEL_SITO = /^(docs\/|[.]github\/compressione\/|[.]github\/pubblicazione\/|[.]github\/workflows\/pubblica-cloudflare[.]yml$|collaudo\/strumenti\/prepara-sito[.]js$)/;
const attendi = (ms) => new Promise((r) => setTimeout(r, ms));
const git = (...a) => execFileSync('git', ['-C', RADICE].concat(a), { encoding: 'utf8' }).trim();

// Il nome del repository dal remoto, senza mai stampare la chiave che l'indirizzo può contenere.
function repository() {
  const pezzi = git('remote', 'get-url', AMB.remoto).replace(/[.]git$/, '').split('/');
  const nome = pezzi.pop(), proprietario = (pezzi.pop() || '').split('@').pop().split(':').pop();
  if (proprietario.toLowerCase() !== UTENTE.toLowerCase() || !/^[A-Za-z0-9._-]+$/.test(nome)) throw new Error('il remoto «' + AMB.remoto + '» non punta a un repository di ' + UTENTE);
  return nome;
}
// La riga che annuncia la versione ai PC: in produzione si legge soltanto.
async function versioneAnnunciata() {
  const sql = 'select sha, message from public.app_version where id = 1';
  const r = AMB.produzione ? await sb.produzione().leggi(sql) : await sb.collaudo().query(sql);
  return r[0] || {};
}

// «null» = GitHub non conosce ancora il flusso (succede nei primi secondi dopo il
// push che lo porta nel repository): chi chiama decide quanto aspettare.
async function esecuzioni(R) {
  const r = await gh('GET', R + '/actions/workflows/' + FLUSSO + '/runs?per_page=5');
  if (r.stato === 404) return null;
  if (r.stato >= 300) throw new Error('elenco delle esecuzioni: HTTP ' + r.stato);
  return (r.dati && r.dati.workflow_runs) || [];
}
// Una risposta sbagliata di GitHub o del database ogni tanto capita: nei cicli di attesa
// si riprova al giro dopo, e ci si ferma solo alla terza di fila.
function paziente(leggi, cosa) {
  let falliti = 0;
  return async (...a) => {
    try { const valore = await leggi(...a); falliti = 0; return { ok: true, valore }; }
    catch (e) { falliti++; if (falliti >= 3) throw e; console.log('… ' + cosa + ': una lettura non è riuscita, riprovo'); return { ok: false }; }
  };
}
// Il commit in cima al ramo fa partire il flusso? Lo dicono i file cambiati dopo
// l'ultima esecuzione, non il tempo che passa. true, false, oppure null se non si può stabilire.
async function toccaIlSito(R, commit, ultima) {
  if (!ultima) return true;                       // mai pubblicato prima: la prima esecuzione è attesa
  if (ultima.head_sha === commit) return true;
  const c = await gh('GET', R + '/compare/' + ultima.head_sha + '...' + commit);
  if (c.stato !== 200 || !c.dati) return null;
  const file = c.dati.files || [];
  if (file.length >= 300) return null;
  return file.some((f) => DEL_SITO.test(f.filename));
}

(async () => {
  const azione = ARG.resto[0] || 'stato';
  if ((azione !== 'stato' && azione !== 'avvia') || ARG.resto.length > 1) { console.log('uso: pubblicazione.js [collaudo|produzione]  |  pubblicazione.js <collaudo|produzione> avvia'); process.exit(1); }
  if (azione === 'avvia') ghApi.pretendiAmbiente(ARG, 'avvia');
  if (AMB.produzione && azione === 'avvia' && !ARG.confermato) {
    console.log('In PRODUZIONE «avvia» ripubblica il sito del reparto: va lanciato con ' + ghApi.CONFERMA + ', e solo con l\'ok di chi gestisce l\'app.');
    process.exit(1);
  }
  const nome = repository();
  const R = '/repos/' + UTENTE + '/' + nome;
  console.log('ambiente: ' + ARG.ambiente + ' | repository ' + UTENTE + '/' + nome);
  // «notify-deploy» scatta solo dove GitHub Pages pubblica questo repository: in produzione, finché il
  // repository ha il suo sito. Altrove (collaudo; produzione dopo lo scambio) l'avviso deve darlo il flusso.
  const scheda = await gh('GET', R);
  if (scheda.stato !== 200 || !scheda.dati) throw new Error('scheda del repository: HTTP ' + scheda.stato);
  const daNotify = AMB.produzione && scheda.dati.has_pages === true;

  // Com'era l'annuncio prima: i PC mostrano «Update» solo se il commit annunciato CAMBIA.
  const annuncioPrima = azione === 'avvia' ? String((await versioneAnnunciata()).sha || '') : '';

  // Quale esecuzione aspettare: quella nuova dopo «avvia», altrimenti quella del commit in cima al remoto.
  let dopo = 0, commit = null, attesa = true;
  if (azione === 'avvia') {
    const prima = await esecuzioni(R);
    if (prima === null) throw new Error('nel repository non c\'è il flusso ' + FLUSSO + ': non si può avviare');
    dopo = (prima[0] || {}).run_number || 0;
    const r = await gh('POST', R + '/actions/workflows/' + FLUSSO + '/dispatches', { ref: 'master' });
    if (r.stato !== 204) throw new Error('avvio a mano del flusso: HTTP ' + r.stato + ' ' + JSON.stringify(r.dati).slice(0, 200));
    console.log('pubblicazione avviata a mano');
  } else {
    const b = await gh('GET', R + '/branches/master');
    commit = b.dati && b.dati.commit && b.dati.commit.sha;
    if (!commit) throw new Error('ramo master del repository: HTTP ' + b.stato);
    console.log('commit in cima al ramo: ' + commit.slice(0, 7));
    // Lanciato prima del push (o dopo un push respinto) lo strumento seguirebbe il rilascio precedente: lo si dice.
    const ramoLocale = AMB.produzione ? 'master' : 'collaudo';
    let locale = '';
    try { locale = git('rev-parse', 'refs/heads/' + ramoLocale); } catch (e) { locale = ''; }
    if (/^[0-9a-f]{40}$/.test(locale) && locale !== commit) console.log('ATTENZIONE: in questa cartella il ramo ' + ramoLocale + ' è a ' + locale.slice(0, 7) + ', sul remoto in cima c\'è ' + commit.slice(0, 7) + ': il push di ' + locale.slice(0, 7) + ' non risulta ancora arrivato. Seguo ciò che c\'è sul remoto e, se la cima cambia, passo a quella');
    // Se il commit non tocca il sito il flusso non parte: lo si stabilisce dai file, non dal tempo.
    const gia = await esecuzioni(R);
    if (gia !== null && !gia.some((x) => x.head_sha === commit)) {
      const tocca = await toccaIlSito(R, commit, gia[0]);
      if (tocca === false) attesa = false;
      else if (tocca === null) console.log('non riesco a stabilire dai file se questo commit fa partire la pubblicazione: la aspetto');
    }
  }
  const voluta = (x) => (azione === 'avvia' ? x.run_number > dopo : x.head_sha === commit);

  let run = null, vista = false, mostroLaPrecedente = false, cimaCambiata = false;
  const elenco = paziente(esecuzioni, 'elenco delle esecuzioni');
  const cimaDelRemoto = paziente(async () => { const b = await gh('GET', R + '/branches/master'); return (b.dati && b.dati.commit && b.dati.commit.sha) || null; }, 'cima del ramo');
  for (let i = 0; i < 96; i++) {
    // finché l'esecuzione attesa non compare, si rilegge la cima: un push arrivato dopo viene seguito
    if (azione === 'stato' && !vista && i > 0) {
      const c = await cimaDelRemoto();
      if (c.ok && c.valore && c.valore !== commit) {
        console.log('sul remoto in cima ora c\'è ' + c.valore.slice(0, 7) + ' (prima ' + commit.slice(0, 7) + '): seguo questo');
        commit = c.valore; attesa = true; cimaCambiata = true;
      }
    }
    const letto = await elenco(R);
    if (!letto.ok) { await attendi(5000); continue; }
    const tutte = letto.valore;
    if (cimaCambiata && tutte !== null) {
      cimaCambiata = false;
      if (!tutte.some((x) => x.head_sha === commit) && (await toccaIlSito(R, commit, tutte[0])) === false) attesa = false;
    }
    if (tutte === null) {
      // subito dopo il primo push che porta il flusso nel repository GitHub può non conoscerlo ancora
      if (i >= 24) throw new Error('dopo due minuti nel repository non risulta il flusso ' + FLUSSO + ' (in cima al remoto c\'è ' + String(commit || '?').slice(0, 7) + '). Se il push che porta il flusso non è ancora arrivato, rilanciare questo strumento dopo il push');
      if (i % 4 === 0) console.log('… il flusso non risulta ancora nel repository: aspetto');
      await attendi(5000);
      continue;
    }
    const mia = tutte.filter(voluta)[0];
    if (mia) { vista = true; if (mia.status === 'completed') { run = mia; break; } }
    if (!vista && !attesa) {
      const ultima = tutte[0] || null;
      if (ultima && ultima.status !== 'completed') {
        if (i % 4 === 0) console.log('il commit ' + commit.slice(0, 7) + ' non cambia nessun file del sito, ma è ancora in corso la pubblicazione n. ' + ultima.run_number + ' (commit ' + String(ultima.head_sha).slice(0, 7) + '): la seguo');
        await attendi(5000);
        continue;
      }
      mostroLaPrecedente = true;
      if (!ultima) { console.log('il commit ' + commit.slice(0, 7) + ' non cambia nessun file del sito: nessuna pubblicazione è attesa, e finora non ne è stata fatta nessuna'); return; }
      run = ultima;
      console.log('il commit ' + commit.slice(0, 7) + ' non cambia nessun file del sito: nessuna pubblicazione è attesa. Mostro l\'ultima fatta');
      break;
    }
    if (i % 4 === 0) console.log('… ' + (mia ? mia.status + ' (n. ' + mia.run_number + ', ' + String(mia.head_sha).slice(0, 7) + ')' : 'in attesa che parta'));
    await attendi(5000);
  }
  if (!run) { console.log(vista || !attesa ? 'ERRORE: dopo otto minuti la pubblicazione non è ancora conclusa' : 'ERRORE: la pubblicazione attesa' + (commit ? ' per il commit ' + commit.slice(0, 7) : '') + ' non è partita'); process.exit(1); }

  console.log('esecuzione n. ' + run.run_number + ' | ' + run.event + ' | commit ' + String(run.head_sha).slice(0, 7) + ' | esito: ' + run.conclusion);
  const j = await gh('GET', R + '/actions/runs/' + run.id + '/jobs');
  const lavori = (j.dati && j.dati.jobs) || [];
  // senza l'elenco dei passi non si può dire né «pubblicato» né «solo controllato»
  const senzaPassi = !lavori.some((x) => (x.steps || []).length);
  if (j.stato >= 300 || (senzaPassi && run.conclusion === 'success')) { console.log('ERRORE: i passi dell\'esecuzione non si leggono (HTTP ' + j.stato + '): rilanciare lo strumento'); process.exit(1); }
  let avvisoFatto = false, pubblicato = false;
  for (const lavoro of lavori) {
    for (const p of (lavoro.steps || [])) {
      console.log('    ' + String(p.conclusion).padEnd(8) + ' ' + p.name);
      if (/Avvisa/.test(p.name) && p.conclusion === 'success') avvisoFatto = true;
      if (p.name === 'Pubblica' && p.conclusion === 'success') pubblicato = true;
    }
    const a = await gh('GET', R + '/check-runs/' + lavoro.id + '/annotations');
    (Array.isArray(a.dati) ? a.dati : []).forEach((x) => console.log('  [' + x.annotation_level + '] ' + (x.title ? x.title + ': ' : '') + String(x.message).slice(0, 400)));
  }

  const letta = async () => { const v = await versioneAnnunciata(); return { sha: String(v.sha), testo: String(v.sha).slice(0, 7) + ' («' + v.message + '»)' }; };
  const mio = String(run.head_sha);
  // Aspetta che «notify-deploy» annunci il commit: al più cinque minuti.
  const aspettaNotify = async () => {
    const rilegge = paziente(letta, 'versione annunciata');
    let v = await letta();
    for (let i = 0; i < 30 && v.sha !== mio; i++) {
      if (i % 3 === 0) console.log('… aspetto l\'avviso di «notify-deploy»: nel database ora c\'è ' + v.testo);
      await attendi(10000);
      const r = await rilegge();
      if (r.ok) v = r.valore;
    }
    return v;
  };
  const TARDI = 'Controllare che GitHub Pages abbia finito e che il flusso «notify-deploy» sia riuscito, poi rilanciare questo strumento';
  if (run.conclusion !== 'success' || !pubblicato) {
    console.log(run.conclusion !== 'success'
      ? 'ERRORE: il flusso di Cloudflare non è riuscito (esito: ' + run.conclusion + '): ' + (pubblicato ? 'il passo «Pubblica» è riuscito, quindi su Cloudflare la versione nuova c\'è, ma un passo successivo (verifica o avviso) è fallito' : 'il sito su Cloudflare NON è stato aggiornato')
      : 'ERRORE: il flusso ha solo controllato i file: Cloudflare non è configurato in questo repository, il sito su Cloudflare NON è stato aggiornato');
    if (daNotify && !mostroLaPrecedente && azione === 'stato') {
      console.log('Il sito che i PC usano oggi è quello di GitHub Pages, che NON passa da questo flusso: lo pubblica GitHub e l\'avviso lo dà «notify-deploy»');
      const v = await aspettaNotify();
      console.log(v.sha === mio
        ? 'su GitHub Pages il commit ' + mio.slice(0, 7) + ' è pubblicato e l\'avviso ai PC è arrivato da «notify-deploy»: nel database c\'è ' + v.testo + '. Resta da sistemare SOLO Cloudflare'
        : 'ATTENZIONE: dopo cinque minuti nemmeno l\'avviso di «notify-deploy» è nel database (c\'è ' + v.testo + '). ' + TARDI);
    }
    process.exit(1);
  }
  let v = await letta();
  if (mostroLaPrecedente) {
    console.log('versione annunciata ai PC nel database di ' + ARG.ambiente + ', in questo momento: ' + v.testo);
    console.log('ultima pubblicazione su Cloudflare: commit ' + mio.slice(0, 7) + (v.sha === mio ? ', che è anche quello annunciato ai PC' : ''));
    if (avvisoFatto && v.sha !== mio) { console.log('ERRORE: quella pubblicazione dice di aver avvisato ma app_version non porta il suo commit'); process.exit(1); }
    return;
  }
  if (avvisoFatto) {
    console.log('versione annunciata ai PC nel database di ' + ARG.ambiente + ': ' + v.testo);
    if (v.sha !== mio) { console.log('ERRORE: il flusso dice di aver avvisato ma app_version non porta il commit pubblicato'); process.exit(1); }
    if (azione === 'avvia' && annuncioPrima === mio) console.log('FATTO: ripubblicato lo stesso commit. L\'annuncio portava già questo commit, quindi sui PC aperti NON compare «Update»: non serve, la versione è la stessa');
    else console.log('FATTO: pubblicato su Cloudflare e avviso inviato ai PC aperti');
    return;
  }
  if (!daNotify) {
    console.log('versione annunciata ai PC nel database di ' + ARG.ambiente + ', in questo momento: ' + v.testo);
    console.log('ERRORE: pubblicato su Cloudflare, ma il flusso NON ha avvisato i PC: il passo «Avvisa i PC aperti» è stato saltato, cioè in questo repository la variabile CLOUDFLARE_AVVISA non vale «si». Qui GitHub Pages non pubblica e «notify-deploy» non scatta: senza quella variabile nessun PC riceve «Update»');
    process.exit(1);
  }
  console.log('pubblicato su Cloudflare. In questo repository l\'avviso ai PC non lo dà questo flusso (CLOUDFLARE_AVVISA è spenta, com\'è giusto finché i PC usano GitHub Pages): lo dà «notify-deploy» quando GitHub Pages ha finito');
  if (azione === 'avvia') { console.log('versione annunciata ai PC, in questo momento: ' + v.testo + '. Una ripubblicazione a mano non fa partire «notify-deploy»: nessun avviso è atteso'); return; }
  v = await aspettaNotify();
  if (v.sha === mio) console.log('FATTO: pubblicato su Cloudflare, e l\'avviso ai PC è arrivato da «notify-deploy»: nel database c\'è ' + v.testo);
  else { console.log('ATTENZIONE: dopo cinque minuti l\'avviso di «notify-deploy» non è ancora nel database (c\'è ' + v.testo + '). ' + TARDI); process.exit(1); }
})().catch((e) => { console.log('ERRORE: ' + ghApi.maschera(e && e.message || e).slice(0, 400)); process.exit(1); });
