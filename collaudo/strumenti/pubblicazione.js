// Dopo lo scambio dei repository il sito di COLLAUDO lo pubblica il flusso
// «Pubblica su Cloudflare» del repository privato, a ogni push su master che
// tocca il sito; con la variabile CLOUDFLARE_AVVISA = si è lo stesso flusso ad
// aggiornare app_version, cioè a far comparire «Update» sulle pagine aperte.
// Questo strumento segue il flusso fino alla fine e controlla che l'avviso sia
// arrivato nel database di collaudo. Non scrive nulla nel database.
//
//   node collaudo/strumenti/pubblicazione.js          segue l'ultima pubblicazione e ne mostra l'esito
//   node collaudo/strumenti/pubblicazione.js avvia    la fa ripartire a mano (stesso commit) e la segue
//
// Il repository è quello a cui punta il remoto «collaudo» di questa cartella:
// così lo strumento continua a funzionare se il repository cambia nome.
const { execFileSync } = require('child_process');
const path = require('path');
const { gh, UTENTE } = require('./gh-api.js');
const { collaudo } = require('./sb.js');

const RADICE = path.resolve(__dirname, '../..');
const FLUSSO = 'pubblica-cloudflare.yml';
const attendi = (ms) => new Promise((r) => setTimeout(r, ms));
const pulito = (t) => String(t).replace(/gh[pousr]_[A-Za-z0-9_]+/g, 'gh*_***');

// Il nome del repository dal remoto, senza mai stampare la chiave che l'indirizzo può contenere.
function repository() {
  const u = execFileSync('git', ['-C', RADICE, 'remote', 'get-url', 'collaudo'], { encoding: 'utf8' }).trim();
  const pezzi = u.replace(/[.]git$/, '').split('/');
  const nome = pezzi.pop(), proprietario = pezzi.pop();
  if (proprietario !== UTENTE || !/^[A-Za-z0-9._-]+$/.test(nome)) throw new Error('il remoto «collaudo» non punta a un repository di ' + UTENTE);
  return nome;
}

async function ultima(R) {
  const r = await gh('GET', R + '/actions/workflows/' + FLUSSO + '/runs?per_page=1');
  if (r.stato >= 300) throw new Error('elenco delle esecuzioni: HTTP ' + r.stato);
  return (r.dati && r.dati.workflow_runs && r.dati.workflow_runs[0]) || null;
}

(async () => {
  const azione = process.argv[2] || 'stato';
  if (azione !== 'stato' && azione !== 'avvia') { console.log('uso: pubblicazione.js [avvia]'); process.exit(1); }
  const nome = repository();
  const R = '/repos/' + UTENTE + '/' + nome;
  console.log('repository ' + UTENTE + '/' + nome);

  let dopo = 0;
  if (azione === 'avvia') {
    const prima = await ultima(R);
    dopo = prima ? prima.run_number : 0;
    const r = await gh('POST', R + '/actions/workflows/' + FLUSSO + '/dispatches', { ref: 'master' });
    if (r.stato !== 204) throw new Error('avvio a mano del flusso: HTTP ' + r.stato + ' ' + JSON.stringify(r.dati).slice(0, 200));
    console.log('pubblicazione avviata a mano');
  }

  // Aspetta la fine dell'esecuzione (dopo «avvia»: di quella nuova).
  let run = null;
  for (let i = 0; i < 80; i++) {
    run = await ultima(R);
    if (run && run.run_number > dopo && run.status === 'completed') break;
    if (i % 4 === 0) console.log('… ' + (run && run.run_number > dopo ? run.status + ' (n. ' + run.run_number + ', ' + String(run.head_sha).slice(0, 7) + ')' : 'in attesa che parta'));
    await attendi(5000);
    run = null;
  }
  if (!run) { console.log('ERRORE: il flusso non è finito in 7 minuti'); process.exit(1); }

  console.log('esecuzione n. ' + run.run_number + ' | ' + run.event + ' | commit ' + String(run.head_sha).slice(0, 7) + ' | esito: ' + run.conclusion);
  const j = await gh('GET', R + '/actions/runs/' + run.id + '/jobs');
  const lavori = (j.dati && j.dati.jobs) || [];
  let avvisoFatto = false;
  for (const lavoro of lavori) {
    for (const p of (lavoro.steps || [])) {
      console.log('    ' + String(p.conclusion).padEnd(8) + ' ' + p.name);
      if (/Avvisa/.test(p.name) && p.conclusion === 'success') avvisoFatto = true;
    }
    const a = await gh('GET', R + '/check-runs/' + lavoro.id + '/annotations');
    (Array.isArray(a.dati) ? a.dati : []).forEach((x) => console.log('  [' + x.annotation_level + '] ' + (x.title ? x.title + ': ' : '') + String(x.message).slice(0, 400)));
  }

  // L'avviso: app_version deve portare il commit appena pubblicato.
  const v = (await collaudo().query('select sha, message from public.app_version where id = 1'))[0] || {};
  const allineata = String(v.sha) === String(run.head_sha);
  console.log('app_version nel collaudo: ' + String(v.sha).slice(0, 7) + ' («' + v.message + '»)');
  if (run.conclusion !== 'success') { console.log('ERRORE: la pubblicazione non è riuscita'); process.exit(1); }
  if (avvisoFatto && allineata) console.log('FATTO: pubblicato e avviso inviato ai PC aperti');
  else if (!avvisoFatto) console.log('pubblicato; avviso ai PC NON inviato (la variabile CLOUDFLARE_AVVISA non vale «si»)');
  else { console.log('ERRORE: il flusso dice di aver avvisato ma app_version non porta il commit pubblicato'); process.exit(1); }
})().catch((e) => { console.log('ERRORE: ' + pulito(e && e.message || e).slice(0, 400)); process.exit(1); });
