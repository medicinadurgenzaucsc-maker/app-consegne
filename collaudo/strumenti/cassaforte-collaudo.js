// La cassaforte della mail nel COLLAUDO ha due vite.
//
//  VERA   il consenso dato davvero, dal sito di collaudo, con l'account Google
//         di prova scelto come mittente: le mail partono davvero, verso i
//         destinatari di prova. È lo stato di riposo.
//  FINTA  le credenziali delle prove automatiche: «Google» è la funzione
//         google-finto e nessuna mail parte. Le batterie di prove chiamano
//         finta() all'inizio e vera() alla fine.
//
// Lo scambio avviene DENTRO il database, da una riga all'altra della tabella
// privata google_oauth: i token veri non passano mai da questo script.
//   riga «reparto»       quella che la funzione google-token usa
//   riga «prova»         le credenziali finte (le crea funzioni-collaudo.js configura)
//   riga «reparto_vero»  dove riposa il consenso vero mentre girano le prove
//
// Sono «finte» tutte le righe il cui client secret comincia come quello finto
// (google-finto ne conosce due validi e considera sbagliati tutti gli altri:
// servono a provare il cambio del secret), o che hanno solo un consenso finto.
//
//   node collaudo/strumenti/cassaforte-collaudo.js stato | finta | vera
const { collaudo } = require('./sb.js');

const CLIENT_FINTO = 'segreto-client-finto-del-collaudo';   // lo stesso di google-finto
const COLONNE = 'client_secret, refresh_token, email, scopes, access_token, access_scad';

async function stato() {
  const righe = await collaudo().query(
    "select id, (left(coalesce(client_secret, ''), " + CLIENT_FINTO.length + ") = '" + CLIENT_FINTO + "') as client_finto, (client_secret is not null) as ha_client, "
    + "(refresh_token is not null) as ha_consenso, (left(coalesce(refresh_token, ''), 6) = 'finto_') as consenso_finto, email "
    + "from public.google_oauth where id in ('reparto', 'prova', 'reparto_vero')");
  const r = righe.find((x) => x.id === 'reparto');
  const finta = !!r && (r.client_finto || (!r.ha_client && r.consenso_finto));
  return {
    modo: !r ? 'assente' : finta ? 'finta' : (r.ha_client || r.ha_consenso) ? 'vera' : 'vuota',
    email: r ? r.email : null,
    consenso: !!(r && r.ha_consenso),
    credenzialiFinte: righe.some((x) => x.id === 'prova'),
    veraInAttesa: righe.some((x) => x.id === 'reparto_vero'),
  };
}

// Mette nella riga «reparto» le credenziali finte e restituisce il segreto del
// finto Google. Se lì c'era un consenso vero, prima lo mette da parte.
async function finta() {
  const c = collaudo();
  const s = await stato();
  if (!s.credenzialiFinte) throw new Error('mancano le credenziali finte: lanciare «node collaudo/strumenti/funzioni-collaudo.js configura»');
  if (s.modo === 'vera' && s.veraInAttesa) {
    throw new Error('c\'è già un consenso vero messo da parte (riga «reparto_vero») e la riga «reparto» ne contiene un altro: stato da guardare a mano');
  }
  // Riga vuota con qualcosa già da parte = una prova interrotta a metà: si
  // riparte senza mettere da parte nulla (ciò che è da parte resta lì).
  if (s.modo !== 'finta' && !s.veraInAttesa) {
    await c.query(
      "insert into public.google_oauth (id, " + COLONNE + ", richiedi_utente, updated_at) "
      + "select 'reparto_vero', " + COLONNE + ", richiedi_utente, updated_at from public.google_oauth where id = 'reparto'");
  }
  const r = await c.query(
    "insert into public.google_oauth (id, client_secret, refresh_token, email, scopes, richiedi_utente) "
    + "select 'reparto', p.client_secret, p.refresh_token, (select lower(trim(valore)) from public.impostazioni where chiave = 'ACCOUNT_LOGIN'), p.scopes, true "
    + "from public.google_oauth p where p.id = 'prova' "
    + "on conflict (id) do update set client_secret = excluded.client_secret, refresh_token = excluded.refresh_token, email = excluded.email, "
    + "scopes = excluded.scopes, access_token = null, access_scad = null, richiedi_utente = true, updated_at = now() "
    + "returning refresh_token");
  return r[0].refresh_token;
}

// Rimette nella riga «reparto» il consenso vero messo da parte. Se non ce n'è
// uno, la lascia vuota: al primo uso l'app chiederà di configurarla.
async function vera() {
  const c = collaudo();
  const s = await stato();
  if (s.veraInAttesa) {
    await c.query(
      "update public.google_oauth r set client_secret = v.client_secret, refresh_token = v.refresh_token, email = v.email, scopes = v.scopes, "
      + "access_token = v.access_token, access_scad = v.access_scad, richiedi_utente = v.richiedi_utente, updated_at = now() "
      + "from public.google_oauth v where r.id = 'reparto' and v.id = 'reparto_vero'; "
      + "delete from public.google_oauth where id = 'reparto_vero'");
  } else if (s.modo === 'finta') {
    await c.query(
      "update public.google_oauth set client_secret = null, refresh_token = null, email = null, scopes = null, access_token = null, access_scad = null, "
      + "richiedi_utente = true, updated_at = now() where id = 'reparto'");
  }
  return stato();
}

// Solo per le prove, e solo su una cassaforte finta: toglie client secret e
// consenso, come al primo avvio (l'app deve chiedere il secret).
async function senzaSegreto() {
  const s = await stato();
  if (s.modo !== 'finta') throw new Error('la cassaforte non è in modalità finta: non tolgo nulla');
  await collaudo().query(
    "update public.google_oauth set client_secret = null, refresh_token = null, email = null, scopes = null, access_token = null, access_scad = null, "
    + "updated_at = now() where id = 'reparto'");
  return stato();
}

// Solo per le prove, e solo su una cassaforte finta: mette un client secret
// che il finto Google non riconosce (come se quello custodito fosse stato
// cancellato nella console Google). Il consenso resta.
async function segretoSbagliato() {
  const s = await stato();
  if (s.modo !== 'finta') throw new Error('la cassaforte non è in modalità finta: non cambio nulla');
  await collaudo().query(
    "update public.google_oauth set client_secret = '" + CLIENT_FINTO + "-sbagliato', access_token = null, access_scad = null, updated_at = now() where id = 'reparto'");
  return stato();
}

module.exports = { stato, finta, vera, senzaSegreto, segretoSbagliato, CLIENT_FINTO };

if (require.main === module) {
  const cmd = process.argv[2] || 'stato';
  const descrivi = (s) => 'cassaforte di collaudo: ' + s.modo.toUpperCase()
    + (s.modo === 'vera' ? (s.consenso ? ', consenso dato da ' + s.email : ', in attesa del consenso') : '')
    + (s.modo === 'vuota' ? ' (da configurare dal sito di collaudo)' : '')
    + (s.veraInAttesa ? ' — un consenso vero è messo da parte' : '')
    + (s.credenzialiFinte ? '' : ' — MANCANO le credenziali finte');
  (cmd === 'finta' ? finta().then(stato) : cmd === 'vera' ? vera() : cmd === 'stato' ? stato() : Promise.reject(new Error('uso: stato | finta | vera')))
    .then((s) => console.log(descrivi(s)))
    .catch((e) => { console.log('ERRORE: ' + String(e.message).replace(/finto_[0-9a-f]+/g, 'finto_***')); process.exitCode = 1; });
}
