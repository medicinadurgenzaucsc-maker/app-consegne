// Il flusso Keepalive di un reparto (.github/workflows/supabase-keepalive.yml), dalla v195 «un codice, tante
// configurazioni»: il reparto si riconosce dalla variabile CLOUDFLARE_PROGETTO del repository (la stessa del
// flusso di pubblicazione), oppure, senza la variabile, dal repository che pubblica la cartella docs/ con
// GitHub Pages. I suoi dati (progetto, indirizzo e chiave pubblica del database, centrale) vengono dal suo
// file in reparti/. Nessun segreto passa di qui.
//
//   node collaudo/strumenti/keepalive.js reparto <progetto Cloudflare o ""> <proprietario/repository>
//        righe «nome=valore» per le uscite del passo: codice, ambiente, ref, url, chiave, centrale.
//        Gli errori vanno sul canale degli errori, così non finiscono fra le uscite del passo.
//   node collaudo/strumenti/keepalive.js segnala
//        manda l'esito del giro alla funzione «keepalive» della centrale, se il file del reparto la nomina.
//        Legge dall'ambiente del flusso CENTRALE, REPARTO, AMBIENTE, FLUSSO e ESITO_<PASSO> per ogni passo.
//        «ok» solo se tutti i passi sono riusciti. Se la centrale non risponde il giro non diventa rosso.
const reparti = require('./reparti.js');

// i passi del flusso, nell'ordine, con ciò che si scrive alla centrale se è quello che ha fallito
const PASSI = [
  ['segreto', 'SUPABASE_MGMT_TOKEN mancante'],
  ['stato', 'Stato progetto + auto-restore se in pausa'],
  ['heartbeat', 'Heartbeat SQL (scrittura vera sul DB)'],
  ['ping', 'Verifica end-to-end (RPC ping via REST anon)'],
  ['rinnovo', 'Rinnova il repo se fermo da 45+ giorni'],
];

// il reparto di questo repository: dal progetto Cloudflare, oppure da githubPages; se il reparto pubblica
// docs/ con GitHub Pages da un altro repository, questo non è il suo (una copia, un fork): errore
function repartoDi(progetto, repository) {
  const repo = String(repository || '').toLowerCase();
  let r = null;
  if (progetto) {
    r = reparti.perProgettoCloudflare(progetto);
    if (!r) throw new Error('nessun reparto dichiara il progetto Cloudflare «' + progetto + '» (variabile CLOUDFLARE_PROGETTO del repository)');
  } else {
    const trovati = reparti.tutti().filter((x) => x.pubblicazione.githubPages && String(x.pubblicazione.githubPages.repository).toLowerCase() === repo);
    if (trovati.length !== 1) throw new Error('senza la variabile CLOUDFLARE_PROGETTO il reparto si riconosce solo dal repository che pubblica docs/ con GitHub Pages, e «' + repository + '» non lo è');
    r = trovati[0];
  }
  if (r.pubblicazione.githubPages && String(r.pubblicazione.githubPages.repository).toLowerCase() !== repo) {
    throw new Error('il reparto «' + r.codice + '» pubblica docs/ da ' + r.pubblicazione.githubPages.repository + ', non da ' + repository);
  }
  return r;
}

function datiDelReparto(progetto, repository) {
  const r = repartoDi(progetto, repository);
  const ref = (/^https:\/\/([a-z]{20})\.supabase\.co$/.exec(r.supabaseUrl) || [])[1];
  if (!ref) throw new Error('indirizzo del database non valido nel file del reparto ' + r.codice);
  return { codice: r.codice, ambiente: r.ambiente, ref, url: r.supabaseUrl, chiave: r.supabaseAnonKey, centrale: r.centrale ? r.centrale.url : '' };
}

// '' se tutti i passi sono riusciti; altrimenti il primo passo fallito o annullato, oppure «esito incompleto»
function primoFallito(env) {
  const p = PASSI.find(([id]) => ['failure', 'cancelled'].indexOf(env['ESITO_' + id.toUpperCase()]) >= 0);
  if (p) return p[1];
  return PASSI.every(([id]) => env['ESITO_' + id.toUpperCase()] === 'success') ? '' : 'esito incompleto';
}

const COMANDI = {
  reparto(progetto, repository) {
    if (!repository) throw new Error('uso: keepalive.js reparto <progetto Cloudflare o ""> <proprietario/repository>');
    const d = datiDelReparto(progetto, repository);
    Object.keys(d).forEach((k) => console.log(k + '=' + d[k]));
  },
  async segnala() {
    const e = process.env;
    if (!e.CENTRALE) { console.log('il file del reparto non nomina una centrale: nessun segnale'); return; }
    if (!/^https:\/\/[a-z]{20}\.supabase\.co$/.test(e.CENTRALE)) throw new Error('indirizzo della centrale non valido');
    const passo = primoFallito(e);
    const corpo = { reparto: String(e.REPARTO || ''), ambiente: String(e.AMBIENTE || ''), esito: passo ? 'ko' : 'ok', passo, flusso: String(e.FLUSSO || '') };
    try {
      const r = await fetch(e.CENTRALE + '/functions/v1/keepalive', { method: 'POST', body: JSON.stringify(corpo), signal: AbortSignal.timeout(15000) });
      const t = await r.text();
      // la risposta si stampa come stringa JSON: nessun a capo può diventare un comando del flusso
      console.log('centrale: HTTP ' + r.status + ' ' + JSON.stringify(t.slice(0, 80)) + ' · esito ' + corpo.esito + (passo ? ', passo «' + passo + '»' : ''));
      if (r.status >= 300) console.log('::warning::la centrale ha risposto ' + r.status + ': il segnale del keepalive non è arrivato');
    } catch (err) {
      console.log('::warning::centrale non raggiunta: il segnale del keepalive non è arrivato (' + JSON.stringify(String((err && err.message) || err).slice(0, 120)) + ')');
    }
  },
};

module.exports = { PASSI, repartoDi, datiDelReparto, primoFallito };

if (require.main === module) {
  const [comando, ...resto] = process.argv.slice(2);
  const f = COMANDI[comando];
  if (!f) { console.error('comandi: ' + Object.keys(COMANDI).join(', ')); process.exitCode = 2; }
  else Promise.resolve().then(() => f(...resto)).catch((err) => { console.error('::error::' + String((err && err.message) || err).slice(0, 300)); process.exitCode = 1; });
}
