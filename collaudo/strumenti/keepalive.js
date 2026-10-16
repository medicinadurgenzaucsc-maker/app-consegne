// Il flusso Keepalive di un reparto (.github/workflows/supabase-keepalive.yml), dalla v195 «un codice, tante
// configurazioni»: il reparto si riconosce dal proprietario del repository in cui il flusso gira, e i suoi
// dati (progetto, indirizzo e chiave pubblica del database, centrale) vengono dal suo file in reparti/.
// Nessun segreto passa di qui.
//
//   node collaudo/strumenti/keepalive.js reparto <proprietario>
//        righe «nome=valore» per le uscite del passo: codice, ambiente, ref, url, chiave, centrale
//   node collaudo/strumenti/keepalive.js segnala
//        manda l'esito del giro alla funzione «keepalive» della centrale, se il file del reparto la nomina.
//        Legge dall'ambiente del flusso CENTRALE, REPARTO, AMBIENTE, FLUSSO e ESITO_<PASSO> per ogni passo.
//        Se la centrale non risponde il giro non diventa rosso: lo dice nel registro del flusso e basta.
const reparti = require('./reparti.js');

// i passi del flusso, nell'ordine, con ciò che si scrive alla centrale se è quello che ha fallito
const PASSI = [
  ['segreto', 'SUPABASE_MGMT_TOKEN mancante'],
  ['stato', 'Stato progetto + auto-restore se in pausa'],
  ['heartbeat', 'Heartbeat SQL (scrittura vera sul DB)'],
  ['ping', 'Verifica end-to-end (RPC ping via REST anon)'],
  ['rinnovo', 'Rinnova il repo se fermo da 45+ giorni'],
];

function datiDelReparto(proprietario) {
  const r = reparti.perProprietario(proprietario);
  const ref = (/^https:\/\/([a-z]{20})\.supabase\.co$/.exec(r.supabaseUrl) || [])[1];
  if (!ref) throw new Error('indirizzo del database non valido nel file del reparto ' + r.codice);
  return { codice: r.codice, ambiente: r.ambiente, ref, url: r.supabaseUrl, chiave: r.supabaseAnonKey, centrale: r.centrale ? r.centrale.url : '' };
}

// il primo passo fallito (o annullato) nell'ordine del flusso; un passo saltato dopo un errore non conta
function primoFallito(env) {
  const p = PASSI.find(([id]) => ['failure', 'cancelled'].indexOf(env['ESITO_' + id.toUpperCase()]) >= 0);
  return p ? p[1] : '';
}

const COMANDI = {
  reparto(proprietario) {
    if (!proprietario) throw new Error('uso: keepalive.js reparto <proprietario del repository>');
    const d = datiDelReparto(proprietario);
    Object.keys(d).forEach((k) => console.log(k + '=' + d[k]));
  },
  async segnala() {
    const e = process.env;
    if (!e.CENTRALE) { console.log('il file del reparto non nomina una centrale: nessun segnale'); return; }
    if (!/^https:\/\/[a-z]{20}\.supabase\.co$/.test(e.CENTRALE)) throw new Error('indirizzo della centrale non valido');
    const passo = primoFallito(e);
    const corpo = { reparto: String(e.REPARTO || ''), ambiente: String(e.AMBIENTE || ''), esito: passo ? 'ko' : 'ok', passo, flusso: String(e.FLUSSO || '') };
    try {
      const r = await fetch(e.CENTRALE + '/functions/v1/keepalive', { method: 'POST', body: JSON.stringify(corpo) });
      const t = await r.text();
      console.log('centrale: HTTP ' + r.status + ' ' + t.slice(0, 80) + ' · esito ' + corpo.esito + (passo ? ', passo «' + passo + '»' : ''));
      if (r.status >= 300) console.log('::warning::la centrale ha risposto ' + r.status + ': il segnale del keepalive non è arrivato');
    } catch (err) {
      console.log('::warning::centrale non raggiunta: il segnale del keepalive non è arrivato (' + String((err && err.message) || err).slice(0, 120) + ')');
    }
  },
};

module.exports = { PASSI, datiDelReparto, primoFallito };

if (require.main === module) {
  const [comando, ...resto] = process.argv.slice(2);
  const f = COMANDI[comando];
  if (!f) { console.log('comandi: ' + Object.keys(COMANDI).join(', ')); process.exitCode = 2; }
  else Promise.resolve().then(() => f(...resto)).catch((err) => { console.log('::error::' + String((err && err.message) || err).slice(0, 300)); process.exitCode = 1; });
}
