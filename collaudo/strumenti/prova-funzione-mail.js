// Batteria di prove della funzione google-token nel progetto di COLLAUDO.
// Gira contro il finto Google (google-finto): nessuna mail parte davvero.
// Le prove che toccano la configurazione la rimettono com'era alla fine.
const { collaudo } = require('./sb.js');
const { genera } = require('./gettone-prova.js');
const amb = require('./ambiente.json');

const esiti = [];
function prova(nome, ok, dettaglio) {
  esiti.push(ok);
  console.log((ok ? 'OK  ' : 'KO  ') + nome + (dettaglio ? ' — ' + dettaglio : ''));
}

(async () => {
  const c = collaudo();
  const { token: utente } = await genera(1);
  const segreto = (await c.query("select refresh_token from public.google_oauth where id = 'reparto'"))[0].refresh_token;
  const accountAtteso = (await c.query("select lower(trim(valore)) as v from public.impostazioni where chiave = 'ACCOUNT_LOGIN'"))[0].v;
  const chiama = async (bearer, corpo) => {
    const r = await fetch(amb.url + '/functions/v1/google-token', {
      method: 'POST', headers: { 'Content-Type': 'application/json', apikey: amb.anon, Authorization: 'Bearer ' + bearer },
      body: JSON.stringify(corpo),
    });
    let j = null; try { j = await r.json(); } catch (e) {}
    return { http: r.status, j: j || {} };
  };
  const posta = async (oggetto) => c.query("select mittente, destinatari, oggetto, corpo_html from public.posta_simulata where oggetto = " + "$o$" + oggetto + "$o$" + " order by id desc limit 1");

  await c.query("delete from public.posta_simulata; update public.google_oauth set richiedi_utente = true where id = 'reparto'");

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
  prova('sostituire il segreto senza conoscerlo: rifiutato', r.http === 403 && /esiste già/.test(r.j.errore || ''));

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
  } finally {
    await c.query("update public.google_oauth set richiedi_utente = true where id = 'reparto'");
  }

  const n = esiti.filter(Boolean).length;
  console.log('\nESITO: ' + n + '/' + esiti.length + (n === esiti.length ? ' — tutte superate' : ' — CI SONO PROVE FALLITE'));
  if (n !== esiti.length) process.exitCode = 1;
})().catch((e) => { console.log('ERRORE: ' + String(e.message).replace(/finto_[0-9a-f]+/g, 'finto_***')); process.exit(1); });
