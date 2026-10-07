// Batteria di prove della funzione google-token nel progetto di COLLAUDO.
// Gira con le credenziali finte (google-finto fa la parte di Google): nessuna
// mail parte davvero. Se nella cassaforte c'è un consenso vero viene messo da
// parte all'inizio e rimesso al suo posto alla fine, come il mittente scelto.
const { collaudo } = require('./sb.js');
const { genera } = require('./gettone-prova.js');
const { finta, vera, CLIENT_FINTO } = require('./cassaforte-collaudo.js');
const amb = require('./ambiente.json');

const esiti = [];
// Le azioni le cui risposte contenevano il client secret custodito: deve restare vuoto.
const secretUsciti = [];
function prova(nome, ok, dettaglio) {
  esiti.push(ok);
  console.log((ok ? 'OK  ' : 'KO  ') + nome + (dettaglio ? ' — ' + dettaglio : ''));
}

(async () => {
  const c = collaudo();
  const { token: utente } = await genera(1);
  const segreto = await finta();
  const mittenteDiRiposo = (await c.query("select valore from public.impostazioni where chiave = 'MAIL_DIMISSIONI_MITTENTE'"))[0];
  await c.query("delete from public.impostazioni where chiave = 'MAIL_DIMISSIONI_MITTENTE'");
  try {
  const accountAtteso = (await c.query("select lower(trim(valore)) as v from public.impostazioni where chiave = 'ACCOUNT_LOGIN'"))[0].v;
  const chiama = async (bearer, corpo) => {
    const r = await fetch(amb.url + '/functions/v1/google-token', {
      method: 'POST', headers: { 'Content-Type': 'application/json', apikey: amb.anon, Authorization: 'Bearer ' + bearer },
      body: JSON.stringify(corpo),
    });
    const testo = await r.text();
    if (testo.indexOf(CLIENT_FINTO) >= 0) secretUsciti.push(String(corpo.azione));
    let j = null; try { j = JSON.parse(testo); } catch (e) {}
    return { http: r.status, j: j || {} };
  };
  const secretCustodito = async () => (await c.query("select client_secret as s from public.google_oauth where id = 'reparto'"))[0].s;
  const consensoCustodito = async () => (await c.query("select (refresh_token is not null) as c from public.google_oauth where id = 'reparto'"))[0].c;
  const posta = async (oggetto) => c.query("select mittente, destinatari, oggetto, corpo_html from public.posta_simulata where oggetto = " + "$o$" + oggetto + "$o$" + " order by id desc limit 1");

  await c.query("delete from public.posta_simulata where reale = false; update public.google_oauth set richiedi_utente = true where id = 'reparto'");

  // ── stato ──
  let r = await chiama(amb.anon, { azione: 'stato' });
  prova('stato da estraneo: risponde ma senza email', r.http === 200 && r.j.configurato === true && r.j.autorizzato === true && r.j.email === null, JSON.stringify(r.j));
  r = await chiama(utente, { azione: 'stato' });
  prova('stato da utente autorizzato: con email', r.http === 200 && r.j.email === accountAtteso);

  // ── invia ──
  const messaggio = { azione: 'invia', destinatari: ['destinatario1@example.com', 'destinatario2@example.com'], oggetto: 'Dimissioni di prova — già èàù €', html: '<p>Elenco <b>finto</b> — àèìòù ✓</p>' + '<div>riga</div>'.repeat(40) };
  r = await chiama(amb.anon, messaggio);
  prova('invio da estraneo: rifiutato', r.http === 403 && r.j.non_autorizzato === true, 'HTTP ' + r.http);
  prova('invio da estraneo: nulla registrato', (await posta(messaggio.oggetto)).length === 0);
  r = await chiama(utente, messaggio);
  prova('invio da utente autorizzato: accettato', r.http === 200 && r.j.ok === true && r.j.destinatari === 2, JSON.stringify(r.j));
  let p = await posta(messaggio.oggetto);
  prova('messaggio ricostruito identico (mittente, destinatari, oggetto con accenti, corpo)',
    p.length === 1 && p[0].mittente === accountAtteso && p[0].destinatari === messaggio.destinatari.join(', ') && p[0].oggetto === messaggio.oggetto && p[0].corpo_html === messaggio.html,
    p.length ? ('corpo ' + p[0].corpo_html.length + ' caratteri') : 'non trovato');

  await new Promise((res) => setTimeout(res, 2600));   // il token in cassaforte deve risultare «vecchio»
  const conRifiuto = { azione: 'invia', destinatari: ['rifiuta@example.com'], oggetto: 'Prova nuovo tentativo ' + Date.now(), html: '<p>x</p>' };
  r = await chiama(utente, conRifiuto);
  prova('401 dal servizio di posta: rinnova il token e riprova', r.http === 200 && r.j.ok === true && (await posta(conRifiuto.oggetto)).length === 1, 'HTTP ' + r.http);
  r = await chiama(utente, { azione: 'invia', destinatari: ['guasto@example.com'], oggetto: 'Prova guasto', html: '<p>x</p>' });
  prova('guasto del servizio di posta: errore leggibile', r.http === 502 && /Guasto simulato/.test(r.j.errore || ''), r.j.errore);
  r = await chiama(utente, { azione: 'invia', destinatari: ['non-un-indirizzo'], oggetto: 'x', html: '<p>x</p>' });
  prova('destinatario non valido: rifiutato prima di spedire', r.http === 400 && /non validi/.test(r.j.errore || ''));
  r = await chiama(utente, { azione: 'invia', destinatari: [], oggetto: 'x', html: '<p>x</p>' });
  prova('nessun destinatario: rifiutato', r.http === 400);
  r = await chiama(utente, { azione: 'invia', destinatari: Array.from({ length: 31 }, (_, i) => 'd' + i + '@example.com'), oggetto: 'x', html: '<p>x</p>' });
  prova('più di 30 destinatari: rifiutato', r.http === 400);

  // ── scambia: il controllo dell'account ──
  r = await chiama(amb.anon, { azione: 'scambia', code: segreto });
  prova('autorizzazione da estraneo: rifiutata', r.http === 403 && r.j.non_autorizzato === true);
  r = await chiama(utente, { azione: 'scambia', code: 'codice-sbagliato' });
  prova('codice Google non valido: rifiutato', r.http === 400 && /scambio rifiutato/.test(r.j.errore || ''));
  r = await chiama(utente, { azione: 'scambia', code: segreto + '.intruso' });
  prova('consenso dato da un ALTRO account Google: rifiutato', r.http === 403 && /account non ammesso/.test(r.j.errore || ''), r.j.errore);
  let riga = (await c.query("select email, (refresh_token = " + "$s$" + segreto + "$s$" + ") as intatto from public.google_oauth where id = 'reparto'"))[0];
  prova('…e la cassaforte resta quella del reparto', riga.email === accountAtteso && riga.intatto === true);

  await c.query("update public.impostazioni set chiave = 'ACCOUNT_LOGIN_TEMP' where chiave = 'ACCOUNT_LOGIN'");
  try {
    r = await chiama(utente, { azione: 'scambia', code: segreto });
    prova('impostazione ACCOUNT_LOGIN assente: rifiuta invece di accettare chiunque', r.http === 500 && /ACCOUNT_LOGIN/.test(r.j.errore || ''), r.j.errore);
  } finally {
    await c.query("update public.impostazioni set chiave = 'ACCOUNT_LOGIN' where chiave = 'ACCOUNT_LOGIN_TEMP'");
  }
  r = await chiama(utente, { azione: 'scambia', code: segreto });
  prova('consenso dell\'account del reparto: accettato', r.http === 200 && r.j.ok === true && r.j.email === accountAtteso, JSON.stringify(r.j));

  // ── config ──
  r = await chiama(amb.anon, { azione: 'config', client_secret: 'x'.repeat(30) });
  prova('configurazione da estraneo: rifiutata', r.http === 403 && r.j.non_autorizzato === true);
  r = await chiama(utente, { azione: 'config', client_secret: 'y'.repeat(30) });
  prova('un client secret che Google (quello vero) non riconosce: rifiutato, quello custodito resta', r.http === 400 && r.j.segreto_errato === true && (await secretCustodito()) === CLIENT_FINTO, r.j.errore);
  r = await chiama(utente, { azione: 'config', client_secret: CLIENT_FINTO + '-sbagliato' });
  prova('client secret sbagliato: errore, nulla salvato', r.http === 400 && r.j.segreto_errato === true && /non riconosce/.test(r.j.errore || '') && (await secretCustodito()) === CLIENT_FINTO, r.j.errore);
  r = await chiama(utente, { azione: 'config', client_secret: '223005241786-abcdefghijklmnopqrstuvwxyz012345.apps.googleusercontent.com' });
  prova('ID client incollato al posto del secret: riconosciuto e rifiutato', r.http === 400 && r.j.id_client === true && (await secretCustodito()) === CLIENT_FINTO, r.j.errore);
  r = await chiama(utente, { azione: 'config', client_secret: 'troppo corto' });
  prova('client secret con una forma impossibile: rifiutato', r.http === 400 && r.j.formato === true);
  r = await chiama(utente, { azione: 'config', client_secret: CLIENT_FINTO + '-bis' });
  prova('client secret valido: sostituisce quello custodito senza doverlo conoscere, il consenso resta buono', r.http === 200 && r.j.ok === true && r.j.autorizzato === true && (await secretCustodito()) === CLIENT_FINTO + '-bis', JSON.stringify(r.j));
  r = await chiama(utente, { azione: 'invia', destinatari: ['destinatario1@example.com'], oggetto: 'Prova dopo il cambio del secret ' + Date.now(), html: '<p>x</p>' });
  prova('…e con il secret nuovo la mail parte', r.http === 200 && r.j.ok === true, JSON.stringify(r.j));
  await c.query("update public.google_oauth set refresh_token = " + "$s$" + segreto + ".guasto$s$" + " where id = 'reparto'");
  r = await chiama(utente, { azione: 'config', client_secret: CLIENT_FINTO });
  prova('Google non risponde durante la verifica: il secret NON viene salvato', r.http === 502 && r.j.non_verificato === true && (await secretCustodito()) === CLIENT_FINTO + '-bis', r.j.errore);
  await c.query("update public.google_oauth set refresh_token = 'finto_revocato' where id = 'reparto'");
  r = await chiama(utente, { azione: 'config', client_secret: CLIENT_FINTO });
  prova('secret valido ma consenso revocato: salvato, e i permessi risultano da concedere', r.http === 200 && r.j.ok === true && r.j.autorizzato === false && (await secretCustodito()) === CLIENT_FINTO, JSON.stringify(r.j));
  await c.query("update public.google_oauth set refresh_token = null, access_token = null, access_scad = null where id = 'reparto'");
  r = await chiama(utente, { azione: 'config', client_secret: CLIENT_FINTO + '-sbagliato' });
  prova('primo avvio con un secret sbagliato: rifiutato prima ancora del consenso', r.http === 400 && r.j.segreto_errato === true && (await secretCustodito()) === CLIENT_FINTO);
  r = await chiama(utente, { azione: 'config', client_secret: CLIENT_FINTO + '-bis' });
  prova('primo avvio con un secret valido: salvato, in attesa del consenso', r.http === 200 && r.j.ok === true && r.j.autorizzato === false && (await secretCustodito()) === CLIENT_FINTO + '-bis', JSON.stringify(r.j));
  await c.query("update public.google_oauth set client_secret = '" + CLIENT_FINTO + "', refresh_token = " + "$s$" + segreto + "$s$" + ", access_token = null, access_scad = null where id = 'reparto'");

  // ── il secret custodito smette di valere (cancellato o sostituito nella console Google) ──
  await c.query("update public.google_oauth set client_secret = '" + CLIENT_FINTO + "-sbagliato', access_token = null, access_scad = null where id = 'reparto'");
  try {
    r = await chiama(utente, { azione: 'stato', verifica: true });
    prova('secret custodito non più riconosciuto: la verifica lo dice, il consenso non viene toccato', r.j.autorizzato === false && r.j.segreto_errato === true && r.j.verificato === true && (await consensoCustodito()) === true, JSON.stringify(r.j).replace(CLIENT_FINTO, '***'));
    r = await chiama(utente, { azione: 'scambia', code: segreto });
    prova('…lo scambio del consenso lo segnala (l\'app richiede il secret)', r.http === 400 && r.j.segreto_errato === true, r.j.errore);
    r = await chiama(utente, { azione: 'invia', destinatari: ['destinatario1@example.com'], oggetto: 'Prova secret non riconosciuto', html: '<p>x</p>' });
    prova('…e l\'invio pure, senza spedire', r.http === 401 && r.j.segreto_errato === true && (await posta('Prova secret non riconosciuto')).length === 0, r.j.errore);
    r = await chiama(utente, { azione: 'config', client_secret: CLIENT_FINTO });
    prova('…inserito un secret valido tutto torna a funzionare, senza rifare il consenso', r.http === 200 && r.j.autorizzato === true, JSON.stringify(r.j));
  } finally {
    await c.query("update public.google_oauth set client_secret = '" + CLIENT_FINTO + "', access_token = null, access_scad = null where id = 'reparto'");
  }
  r = await chiama(utente, { azione: 'invia', destinatari: ['apispenta@example.com'], oggetto: 'Prova Gmail API spente', html: '<p>x</p>' });
  prova('Gmail API non abilitate nel progetto Google: messaggio che dice dove mettere le mani', r.http === 502 && r.j.api_spenta === true && /Gmail API non sono attive/.test(r.j.errore || ''), r.j.errore);

  // ── consenso revocato ──
  await c.query("update public.google_oauth set refresh_token = 'finto_revocato', access_token = null, access_scad = null where id = 'reparto'");
  try {
    r = await chiama(utente, { azione: 'invia', destinatari: ['destinatario1@example.com'], oggetto: 'Prova revoca', html: '<p>x</p>' });
    prova('consenso revocato: chiede di riautorizzare', r.http === 401 && r.j.riautorizzare === true, r.j.errore);
    r = await chiama(utente, { azione: 'stato' });
    prova('…e lo stato torna «non autorizzato»', r.http === 200 && r.j.autorizzato === false);
  } finally {
    await c.query("update public.google_oauth set refresh_token = " + "$s$" + segreto + "$s$" + ", access_token = null, access_scad = null where id = 'reparto'");
  }
  r = await chiama(utente, { azione: 'stato' });
  prova('cassaforte di collaudo ripristinata', r.j.autorizzato === true);

  // ── verifica dal vivo dei permessi (stato con «verifica») ──
  // Del token si confronta solo l'impronta: il valore non esce dal database.
  const inCassaforte = async () => (await c.query("select md5(coalesce(access_token, '')) as t, (refresh_token is not null) as consenso from public.google_oauth where id = 'reparto'"))[0];
  let primaT = await inCassaforte();
  r = await chiama(utente, { azione: 'stato', verifica: true });
  let dopoT = await inCassaforte();
  prova('verifica dal vivo: Google rilascia un token nuovo, permessi confermati', r.http === 200 && r.j.autorizzato === true && r.j.verificato === true && primaT.t !== dopoT.t, JSON.stringify(r.j));
  r = await chiama(utente, { azione: 'stato' });
  prova('senza richiesta di verifica lo stato non interroga Google', r.j.verificato === undefined && (await inCassaforte()).t === dopoT.t);
  r = await chiama(amb.anon, { azione: 'stato', verifica: true });
  prova('un estraneo non può far interrogare Google', r.http === 200 && r.j.verificato === undefined && (await inCassaforte()).t === dopoT.t, JSON.stringify(r.j));
  await c.query("update public.google_oauth set refresh_token = " + "$s$" + segreto + ".guasto$s$" + " where id = 'reparto'");
  try {
    r = await chiama(utente, { azione: 'stato', verifica: true });
    prova('Google non risponde: permessi «non verificati», il consenso resta dov\'è', r.http === 200 && r.j.autorizzato === true && r.j.verificato === false && /Guasto simulato/.test(r.j.problema || '') && (await inCassaforte()).consenso === true, JSON.stringify(r.j));
    await c.query("update public.google_oauth set refresh_token = 'finto_revocato' where id = 'reparto'");
    r = await chiama(utente, { azione: 'stato', verifica: true });
    prova('consenso revocato: la verifica lo scopre prima dell\'invio', r.http === 200 && r.j.autorizzato === false && r.j.verificato === true && (await inCassaforte()).consenso === false, JSON.stringify(r.j));
  } finally {
    await c.query("update public.google_oauth set refresh_token = " + "$s$" + segreto + "$s$" + ", access_token = null, access_scad = null where id = 'reparto'");
  }
  r = await chiama(utente, { azione: 'stato', verifica: true });
  prova('rimesso il consenso, la verifica torna a confermare', r.j.autorizzato === true && r.j.verificato === true);

  // ── il mittente configurabile («Impostazioni email») ──
  const impostaMittente = (v) => c.query("insert into public.impostazioni (chiave, valore) values ('MAIL_DIMISSIONI_MITTENTE', $m$" + v + "$m$) on conflict (chiave) do update set valore = excluded.valore");
  const togliMittente = () => c.query("delete from public.impostazioni where chiave = 'MAIL_DIMISSIONI_MITTENTE'");
  const altro = 'mittente-prova@example.com';
  const come = (email) => segreto + '.come.' + Buffer.from(email).toString('base64url');
  r = await chiama(utente, { azione: 'stato' });
  prova('mittente mai impostato: vale l\'account del reparto', r.j.mittente === accountAtteso && r.j.autorizzato === true, JSON.stringify(r.j));
  r = await chiama(amb.anon, { azione: 'stato' });
  prova('a un estraneo il mittente non viene detto', r.j.mittente === null && r.j.email === null);
  await impostaMittente(altro);
  r = await chiama(utente, { azione: 'stato' });
  prova('mittente cambiato: i permessi risultano mancanti', r.j.autorizzato === false && r.j.mittente === altro && r.j.email === accountAtteso, JSON.stringify(r.j));
  const senzaPermessi = { azione: 'invia', destinatari: ['destinatario1@example.com'], oggetto: 'Prova senza permessi ' + Date.now(), html: '<p>x</p>' };
  r = await chiama(utente, senzaPermessi);
  prova('invio senza i permessi del mittente: rifiutato, chiede di concederli', r.http === 401 && r.j.riautorizzare === true && /mancano i permessi/.test(r.j.errore || ''), r.j.errore);
  prova('…e nulla è stato spedito', (await posta(senzaPermessi.oggetto)).length === 0);
  r = await chiama(utente, { azione: 'scambia', code: segreto });
  prova('consenso dato con un account diverso dal mittente: rifiutato', r.http === 403 && /account non ammesso/.test(r.j.errore || '') && r.j.mittente === altro, r.j.errore);
  riga = (await c.query("select email from public.google_oauth where id = 'reparto'"))[0];
  prova('…e la cassaforte non è cambiata', riga.email === accountAtteso);
  r = await chiama(utente, { azione: 'scambia', code: come(altro) });
  prova('consenso dato con l\'account del mittente: accettato', r.http === 200 && r.j.ok === true && r.j.email === altro, JSON.stringify(r.j));
  r = await chiama(utente, { azione: 'stato' });
  prova('…e i permessi risultano concessi', r.j.autorizzato === true && r.j.mittente === altro && r.j.email === altro);
  const dalNuovo = { azione: 'invia', destinatari: ['destinatario1@example.com'], oggetto: 'Prova dal nuovo mittente ' + Date.now(), html: '<p>x</p>' };
  r = await chiama(utente, dalNuovo);
  p = await posta(dalNuovo.oggetto);
  prova('la mail parte a nome del nuovo mittente', r.http === 200 && r.j.mittente === altro && p.length === 1 && p[0].mittente === altro, JSON.stringify(r.j));
  const mittenteVecchio = { azione: 'invia', mittente: accountAtteso, destinatari: ['destinatario1@example.com'], oggetto: 'Prova mittente cambiato ' + Date.now(), html: '<p>x</p>' };
  r = await chiama(utente, mittenteVecchio);
  prova('la pagina mostrava un mittente che non è più quello: la mail non parte', r.http === 409 && r.j.mittente_cambiato === true && r.j.mittente === altro && (await posta(mittenteVecchio.oggetto)).length === 0, r.j.errore);
  r = await chiama(utente, Object.assign({}, mittenteVecchio, { mittente: ' Mittente-Prova@EXAMPLE.com ' }));
  prova('…se il mittente mostrato è quello giusto, parte', r.http === 200 && r.j.ok === true && (await posta(mittenteVecchio.oggetto)).length === 1, JSON.stringify(r.j));
  await impostaMittente('  Mittente-Prova@Example.COM ');
  r = await chiama(utente, { azione: 'stato' });
  prova('maiuscole e spazi nell\'indirizzo non contano', r.j.autorizzato === true && r.j.mittente === altro);
  await impostaMittente('non-un-indirizzo');
  r = await chiama(utente, { azione: 'stato' });
  prova('mittente che non è un indirizzo: nessun permesso e nessun ripiego silenzioso', r.j.autorizzato === false && r.j.mittente === null, JSON.stringify(r.j));
  r = await chiama(utente, { azione: 'invia', destinatari: ['destinatario1@example.com'], oggetto: 'Prova mittente non valido', html: '<p>x</p>' });
  prova('…l\'invio viene rifiutato', r.http === 409 && /non determinabile/.test(r.j.errore || ''), r.j.errore);
  r = await chiama(utente, { azione: 'scambia', code: come(altro) });
  prova('…e anche il consenso', r.http === 500 && /non determinabile/.test(r.j.errore || ''));
  await togliMittente();
  r = await chiama(utente, { azione: 'stato' });
  prova('tolto il mittente si torna all\'account del reparto, che deve ridare il consenso', r.j.mittente === accountAtteso && r.j.autorizzato === false);
  r = await chiama(utente, { azione: 'scambia', code: segreto });
  prova('consenso dell\'account del reparto: di nuovo accettato', r.http === 200 && r.j.email === accountAtteso);

  // ── credenziali vere: la richiesta deve arrivare a Google ──
  // (qui con un codice inventato e un client secret inventato: Google rifiuta,
  // ed è proprio la sua risposta che si vuole vedere, non quella del finto)
  await c.query("update public.google_oauth set client_secret = 'segreto-inventato-per-la-prova' where id = 'reparto'");
  try {
    r = await chiama(utente, { azione: 'scambia', code: 'codice-inventato' });
    prova('con un client secret non finto la richiesta viene girata a Google (che non lo riconosce)', r.http === 400 && r.j.segreto_errato === true && /client secret is invalid/.test(r.j.errore || ''), r.j.errore);
  } finally {
    await c.query("update public.google_oauth set client_secret = '" + CLIENT_FINTO + "' where id = 'reparto'");
  }

  // ── google-finto non è un passaggio aperto verso Google ──
  // (gira richieste a Google: deve rispondere solo a google-token, che ne
  // conosce l'indirizzo completo; chi lo chiama da fuori non trova nulla)
  const finto = async (percorso, init) => (await fetch(amb.url + '/functions/v1/google-finto' + percorso, init)).status;
  const modulo = { method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, body: 'grant_type=authorization_code&code=x&client_id=x&client_secret=x&redirect_uri=postmessage' };
  prova('finto Google senza la chiave nell\'indirizzo: non trovato (scambio)', (await finto('/token', modulo)) === 404);
  prova('finto Google con una chiave sbagliata: non trovato', (await finto('/k/' + '0'.repeat(40) + '/token', modulo)) === 404);
  prova('finto Google senza chiave: non trovato (utente e invio)',
    (await finto('/userinfo', { headers: { Authorization: 'Bearer x' } })) === 404
    && (await finto('/gmail-invio', { method: 'POST', headers: { Authorization: 'Bearer x', 'Content-Type': 'application/json' }, body: '{"raw":""}' })) === 404);

  // ── varie ──
  r = await chiama(utente, { azione: 'token' });
  prova('la vecchia azione «token» non esiste', r.http === 400 && /sconosciuta/.test(r.j.errore || ''));

  // ── interruttore spento: comportamento della produzione di oggi ──
  await c.query("update public.google_oauth set richiedi_utente = false where id = 'reparto'");
  try {
    r = await chiama(amb.anon, { azione: 'invia', destinatari: ['destinatario1@example.com'], oggetto: 'Prova interruttore spento', html: '<p>x</p>' });
    prova('interruttore spento: l\'invio da estraneo passa (è il varco che l\'interruttore chiude)', r.http === 200 && r.j.ok === true);
    r = await chiama(amb.anon, { azione: 'scambia', code: segreto + '.intruso' });
    prova('interruttore spento: l\'account intruso è comunque rifiutato (correzione v4)', r.http === 403 && /account non ammesso/.test(r.j.errore || ''));
    r = await chiama(amb.anon, { azione: 'stato' });
    prova('interruttore spento: lo stato dice il mittente anche senza sessione (l\'app deve poter lavorare)', r.j.mittente === accountAtteso && r.j.email === accountAtteso, JSON.stringify(r.j));
  } finally {
    await c.query("update public.google_oauth set richiedi_utente = true where id = 'reparto'");
  }

  } finally {
    await c.query("delete from public.impostazioni where chiave = 'MAIL_DIMISSIONI_MITTENTE'");
    if (mittenteDiRiposo) await c.query("insert into public.impostazioni (chiave, valore) values ('MAIL_DIMISSIONI_MITTENTE', $m$" + mittenteDiRiposo.valore + "$m$)");
    const dopo = await vera();
    console.log('\ncassaforte di collaudo rimessa a riposo: ' + dopo.modo + (dopo.consenso ? ' (consenso di ' + dopo.email + ')' : ''));
  }

  prova('il client secret custodito non è comparso in nessuna risposta della funzione', secretUsciti.length === 0, secretUsciti.join(', '));

  const n = esiti.filter(Boolean).length;
  console.log('\nESITO: ' + n + '/' + esiti.length + (n === esiti.length ? ' — tutte superate' : ' — CI SONO PROVE FALLITE'));
  if (n !== esiti.length) process.exitCode = 1;
})().catch((e) => { console.log('ERRORE: ' + String(e.message).replace(/finto_[0-9a-f]+/g, 'finto_***')); process.exit(1); });
