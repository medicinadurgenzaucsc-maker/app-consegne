// L'uscita da un dispositivo non deve scollegare gli altri.
//
// I PC del reparto usano tutti lo stesso account: la voce «Esci» chiama
// signOut({ scope: 'local' }), perché quello predefinito ('global') revoca
// la sessione dell'account su TUTTI i dispositivi. Questa prova lo dimostra
// sul servizio di autenticazione del COLLAUDO con sessioni vere:
//
//   1. crea un utente provvisorio (indirizzo example.com, nessuna password)
//   2. gli apre tre sessioni, come tre dispositivi diversi
//   3. esce dalla prima con scope 'local' → la prima non si rinnova più,
//      la seconda e la terza sì
//   4. esce dalla seconda con scope 'global' → cade anche la terza: è ciò
//      che succederebbe al reparto con un signOut() senza argomenti
//   5. elimina l'utente provvisorio
//
//   node collaudo/strumenti/prova-uscita.js
//
// Le chiavi si leggono dall'API di gestione e non vengono mai stampate.
const crypto = require('crypto');
const { collaudo, maschera } = require('./sb.js');

const c = collaudo();
const AUTH = 'https://' + c.ref + '.supabase.co/auth/v1';
let superate = 0, fallite = 0;
function esito(ok, testo) { ok ? superate++ : fallite++; console.log((ok ? 'ok  ' : 'KO  ') + testo); }

async function chiama(metodo, percorso, chiave, bearer, corpo) {
  const r = await fetch(AUTH + percorso, {
    method: metodo,
    headers: { apikey: chiave, Authorization: 'Bearer ' + (bearer || chiave), ...(corpo ? { 'Content-Type': 'application/json' } : {}) },
    body: corpo ? JSON.stringify(corpo) : undefined,
  });
  const testo = await r.text();
  let dati = null;
  try { dati = testo ? JSON.parse(testo) : null; } catch (e) { dati = testo; }
  return { stato: r.status, dati };
}
const motivo = (r) => maschera(JSON.stringify(r.dati || '')).slice(0, 160);

(async () => {
  const chiavi = await c.api('GET', '/api-keys?reveal=true');
  const trova = (nome) => ((chiavi.dati || []).find((k) => k.name === nome) || {}).api_key;
  const SERVIZIO = trova('service_role'), PUBBLICA = trova('anon');
  if (!SERVIZIO || !PUBBLICA) throw new Error('chiavi del collaudo non disponibili (HTTP ' + chiavi.stato + ')');

  const email = 'prova-uscita-' + crypto.randomBytes(4).toString('hex') + '@example.com';
  const creato = await chiama('POST', '/admin/users', SERVIZIO, null, { email, email_confirm: true });
  if (creato.stato >= 300 || !creato.dati || !creato.dati.id) throw new Error('utente provvisorio non creato (HTTP ' + creato.stato + '): ' + motivo(creato));
  const idUtente = creato.dati.id;
  console.log('utente provvisorio creato: ' + email);

  try {
    // Una sessione vera per «dispositivo»: collegamento generato con la chiave di servizio, poi verificato.
    async function apriSessione(nome) {
      const link = await chiama('POST', '/admin/generate_link', SERVIZIO, null, { type: 'magiclink', email });
      const impronta = link.dati && (link.dati.hashed_token || (link.dati.properties && link.dati.properties.hashed_token));
      if (link.stato >= 300 || !impronta) throw new Error(nome + ': collegamento non generato (HTTP ' + link.stato + '): ' + motivo(link));
      let v = await chiama('POST', '/verify', PUBBLICA, null, { type: 'magiclink', token_hash: impronta });
      if (v.stato >= 300) v = await chiama('POST', '/verify', PUBBLICA, null, { type: 'email', token_hash: impronta });
      if (v.stato >= 300 || !v.dati || !v.dati.refresh_token) throw new Error(nome + ': sessione non aperta (HTTP ' + v.stato + '): ' + motivo(v));
      return { nome, accesso: v.dati.access_token, rinnovo: v.dati.refresh_token };
    }
    // Il rinnovo è ciò che tiene collegato un dispositivo: se fallisce, quel dispositivo è fuori.
    async function siRinnova(s) {
      const r = await chiama('POST', '/token?grant_type=refresh_token', PUBBLICA, null, { refresh_token: s.rinnovo });
      if (r.stato === 200 && r.dati && r.dati.refresh_token) { s.accesso = r.dati.access_token; s.rinnovo = r.dati.refresh_token; return true; }
      return false;
    }

    const a = await apriSessione('dispositivo A'), b = await apriSessione('dispositivo B'), d = await apriSessione('dispositivo C');
    esito(await siRinnova(a) && await siRinnova(b) && await siRinnova(d), 'tre sessioni aperte per lo stesso account, tutte si rinnovano');

    const locale = await chiama('POST', '/logout?scope=local', PUBBLICA, a.accesso);
    esito(locale.stato === 204, 'uscita con scope «local» dal dispositivo A: HTTP ' + locale.stato);
    esito(!(await siRinnova(a)), 'dopo l\'uscita il dispositivo A non si rinnova più');
    esito(await siRinnova(b), 'il dispositivo B resta collegato');
    esito(await siRinnova(d), 'il dispositivo C resta collegato');

    const globale = await chiama('POST', '/logout?scope=global', PUBBLICA, b.accesso);
    esito(globale.stato === 204, 'controprova: uscita con scope «global» dal dispositivo B: HTTP ' + globale.stato);
    esito(!(await siRinnova(d)), 'controprova: con «global» cade anche il dispositivo C (è ciò che «Esci» NON deve fare)');
  } finally {
    const via = await chiama('DELETE', '/admin/users/' + idUtente, SERVIZIO);
    esito(via.stato === 200 || via.stato === 204, 'utente provvisorio eliminato: HTTP ' + via.stato);
    const resti = await c.query("select count(*)::int as n from auth.users where email like 'prova-uscita-%@example.com'");
    esito(resti[0].n === 0, 'nessun utente provvisorio rimasto nel collaudo');
  }
  console.log('\n' + superate + ' superate, ' + fallite + ' fallite');
  process.exit(fallite ? 1 : 0);
})().catch((e) => { console.log('ERRORE: ' + maschera(e.message)); process.exit(1); });
