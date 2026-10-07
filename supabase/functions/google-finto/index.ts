// google-finto — esiste SOLO nel progetto di collaudo.
// Sta fra google-token e Google, e ha due comportamenti:
//
//  1. CREDENZIALI FINTE (quelle delle prove automatiche) -> fa la parte di
//     Google: rilascia token finti, dice chi e' l'utente e «spedisce» la mail
//     scrivendola in posta_simulata. Gira cosi' tutto il codice vero di
//     google-token (rinnovo del token, costruzione del messaggio, nuovo
//     tentativo sul 401) senza che parta mai una mail.
//
//  2. CREDENZIALI VERE (il consenso dato davvero, dal sito di collaudo, con un
//     account Google vero) -> gira la richiesta a Google cosi' com'e' e ne
//     restituisce la risposta: la mail parte davvero, dal mittente di prova
//     ai destinatari di prova. Di ogni invio riuscito resta una copia in
//     posta_simulata, segnata come «reale».
//
// Come si distinguono: allo scambio e al rinnovo dal «client secret» (quello
// finto e' una frase fissa, CLIENT_FINTO); per userinfo e invio dal token
// (quelli finti cominciano col segreto del collaudo, FINTO_SEGRETO, che
// conoscono solo google-token e gli strumenti di prova).
//
// CHI PUO' CHIAMARLA. Questa funzione gira richieste a Google e non verifica
// il JWT: senza un controllo sarebbe un passaggio aperto per chiunque. Risponde
// quindi solo sotto un indirizzo che contiene una chiave casuale
// (…/google-finto/k/<chiave>/token, /userinfo, /gmail-invio): la chiave sta
// nella variabile FINTO_CHIAVE e nelle variabili GOOGLE_URL_* che google-token
// usa come indirizzi di Google. Nessun altro la conosce; ogni altro indirizzo
// risponde «non trovato».
//
// Codici di autorizzazione finti:
//   <segreto>                    consenso dato dall'account del reparto
//   <segreto>.come.<b64url>      consenso dato dall'indirizzo codificato
//   <segreto>.intruso            consenso dato da intruso@example.com
// Token di rinnovo «<segreto>.guasto»: Google non risponde (503), per provare
// la verifica dei permessi quando Google non e' raggiungibile.
// Client secret finti: quello delle prove, un secondo altrettanto «valido»
// (<quello>-bis, per provare la sostituzione) e, con lo stesso inizio,
// qualunque altro: per il finto Google e' un secret SBAGLIATO, e risponde
// invalid_client come farebbe Google.
// Destinatario «apispenta@example.com»: risponde come Google quando nel
// progetto le Gmail API non sono abilitate.
// Destinatari che simulano un guasto: «rifiuta@example.com» fa rispondere 401
// se il token d'accesso ha piu' di due secondi (google-token deve rinnovarlo
// e riprovare), «guasto@example.com» fa rispondere 500.

import { createClient } from 'npm:@supabase/supabase-js@2';

const SEGRETO = Deno.env.get('FINTO_SEGRETO') || '';
const CHIAVE = Deno.env.get('FINTO_CHIAVE') || '';
const CLIENT_FINTO = 'segreto-client-finto-del-collaudo';
const segretoFinto = (s: string) => s.startsWith(CLIENT_FINTO);
const segretoFintoValido = (s: string) => s === CLIENT_FINTO || s === CLIENT_FINTO + '-bis';
const GOOGLE = {
  token: 'https://oauth2.googleapis.com/token',
  userinfo: 'https://www.googleapis.com/oauth2/v3/userinfo',
  invio: 'https://gmail.googleapis.com/gmail/v1/users/me/messages/send',
};
const sb = createClient(Deno.env.get('SUPABASE_URL')!, Deno.env.get('SUPABASE_SERVICE_ROLE_KEY')!);
const AMBITI = 'openid https://www.googleapis.com/auth/userinfo.email https://www.googleapis.com/auth/userinfo.profile https://www.googleapis.com/auth/gmail.send';

const json = (body: unknown, status = 200) =>
  new Response(JSON.stringify(body), { status, headers: { 'Content-Type': 'application/json' } });

function daBase64(s: string): Uint8Array {
  const pulito = s.replace(/-/g, '+').replace(/_/g, '/').replace(/\s+/g, '');
  const bin = atob(pulito + '='.repeat((4 - (pulito.length % 4)) % 4));
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return out;
}
const utf8 = new TextDecoder();

// Intestazione RFC 2047: «=?UTF-8?B?...?=» -> testo.
function decodificaIntestazione(v: string): string {
  return v.replace(/=\?UTF-8\?B\?([^?]*)\?=/gi, (_m, b64) => utf8.decode(daBase64(b64)));
}

function leggiMessaggio(raw: string) {
  const testo = utf8.decode(daBase64(raw));
  const taglio = testo.indexOf('\r\n\r\n');
  const intestazioni: Record<string, string> = {};
  for (const riga of testo.slice(0, taglio < 0 ? testo.length : taglio).split('\r\n')) {
    const i = riga.indexOf(':');
    if (i > 0) intestazioni[riga.slice(0, i).trim().toLowerCase()] = riga.slice(i + 1).trim();
  }
  let corpo = taglio < 0 ? '' : testo.slice(taglio + 4);
  if ((intestazioni['content-transfer-encoding'] || '').toLowerCase() === 'base64') {
    try { corpo = utf8.decode(daBase64(corpo)); } catch { /* corpo lasciato com'e' */ }
  }
  return {
    mittente: intestazioni['from'] || '',
    destinatari: intestazioni['to'] || '',
    oggetto: decodificaIntestazione(intestazioni['subject'] || ''),
    corpo_html: corpo,
  };
}

// Ogni token d'accesso finto porta l'istante del rilascio (serve a simulare
// la scadenza senza tenere stato in memoria: le chiamate non lo condividono)
// e, se il consenso e' stato dato «come» un certo indirizzo, quell'indirizzo.
const nuovoToken = (come = '') => SEGRETO + '.t' + Date.now() + (come ? '.e.' + come : '');
const etaToken = (t: string) => Date.now() - (parseInt(t.slice((SEGRETO + '.t').length), 10) || 0);
function indirizzoDelToken(t: string): string {
  const i = t.indexOf('.e.');
  if (i < 0) return '';
  try { return utf8.decode(daBase64(t.slice(i + 3))).toLowerCase().trim(); } catch { return ''; }
}

// Gira la richiesta a Google e ne restituisce la risposta, stato compreso.
async function daGoogle(url: string, init: RequestInit): Promise<Response> {
  const r = await fetch(url, init);
  return new Response(await r.text(), { status: r.status, headers: { 'Content-Type': r.headers.get('Content-Type') || 'application/json' } });
}

// Confronto che impiega lo stesso tempo qualunque sia il punto in cui le due
// stringhe differiscono.
function uguali(a: string, b: string): boolean {
  if (a.length !== b.length) return false;
  let diversi = 0;
  for (let i = 0; i < a.length; i++) diversi |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return diversi === 0;
}

Deno.serve(async (req) => {
  if (!SEGRETO || CHIAVE.length < 24) return json({ error: 'collaudo non configurato' }, 500);
  // …/google-finto/k/<chiave>/<percorso>: senza la chiave giusta non c'e' nulla.
  const pezzi = /^\/k\/([^/]+)(\/.*)$/.exec(new URL(req.url).pathname.replace(/^.*\/google-finto/, ''));
  if (!pezzi || !uguali(pezzi[1], CHIAVE)) return json({ error: 'non trovato' }, 404);
  const percorso = pezzi[2];

  // ── scambio del codice e rinnovo del token ──────────────────────────
  if (percorso === '/token' && req.method === 'POST') {
    const corpo = await req.text();
    const f = new URLSearchParams(corpo);
    const segretoClient = f.get('client_secret') || '';
    if (!segretoFinto(segretoClient)) {
      return daGoogle(GOOGLE.token, { method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, body: corpo });
    }
    if (!segretoFintoValido(segretoClient)) {
      return json({ error: 'invalid_client', error_description: 'The provided client secret is invalid.' }, 401);
    }
    const rilascia = (accesso: string, rinnovo?: string) =>
      json({ access_token: accesso, ...(rinnovo ? { refresh_token: rinnovo } : {}), expires_in: 3599, scope: AMBITI, token_type: 'Bearer' });
    if (f.get('grant_type') === 'refresh_token' && f.get('refresh_token') === SEGRETO) return rilascia(nuovoToken());
    if (f.get('grant_type') === 'refresh_token' && f.get('refresh_token') === SEGRETO + '.guasto') {
      return json({ error: 'temporarily_unavailable', error_description: 'Guasto simulato di Google' }, 503);
    }
    if (f.get('grant_type') === 'authorization_code') {
      const code = f.get('code') || '';
      if (code === SEGRETO) return rilascia(nuovoToken(), SEGRETO);
      if (code === SEGRETO + '.intruso') return rilascia(SEGRETO + '.intruso', SEGRETO + '.intruso');
      if (code.startsWith(SEGRETO + '.come.')) return rilascia(nuovoToken(code.slice((SEGRETO + '.come.').length)), SEGRETO);
    }
    return json({ error: 'invalid_grant', error_description: 'Token has been expired or revoked.' }, 400);
  }

  const autorizzazione = req.headers.get('Authorization') || '';
  const bearer = autorizzazione.replace(/^Bearer\s+/i, '');
  const finto = bearer.startsWith(SEGRETO);

  // ── chi e' l'utente ──────────────────────────────────────────────────
  if (percorso === '/userinfo') {
    if (!finto) return daGoogle(GOOGLE.userinfo, { headers: { Authorization: autorizzazione } });
    if (bearer === SEGRETO + '.intruso') return json({ email: 'intruso@example.com', email_verified: true });
    if (!bearer.startsWith(SEGRETO + '.t')) return json({ error: { code: 401, message: 'Invalid Credentials' } }, 401);
    const come = indirizzoDelToken(bearer);
    if (come) return json({ email: come, email_verified: true });
    const { data } = await sb.from('impostazioni').select('valore').eq('chiave', 'ACCOUNT_LOGIN').maybeSingle();
    return json({ email: String(data?.valore || '').trim(), email_verified: true });
  }

  // ── invio della mail ─────────────────────────────────────────────────
  if (percorso === '/gmail-invio' && req.method === 'POST') {
    const corpo = await req.text();
    let raw = '';
    try { raw = String(JSON.parse(corpo).raw || ''); } catch { return json({ error: { message: 'corpo non valido' } }, 400); }

    if (!finto) {
      // Consenso vero: la mail parte davvero. Se Google la accetta ne resta copia.
      const g = await fetch(GOOGLE.invio, { method: 'POST', headers: { Authorization: autorizzazione, 'Content-Type': 'application/json' }, body: corpo });
      const risposta = await g.text();
      if (g.ok) {
        let id = '';
        try { id = String(JSON.parse(risposta).id || ''); } catch { /* risposta non JSON */ }
        try { await sb.from('posta_simulata').insert({ ...leggiMessaggio(raw), grezzo: raw, reale: true, esito: 'inviata da Google, id ' + id }); } catch { /* la copia non deve far fallire l'invio */ }
      }
      return new Response(risposta, { status: g.status, headers: { 'Content-Type': g.headers.get('Content-Type') || 'application/json' } });
    }

    if (!bearer.startsWith(SEGRETO + '.t')) return json({ error: { code: 401, message: 'Invalid Credentials' } }, 401);
    const m = leggiMessaggio(raw);
    if (/guasto@example\.com/i.test(m.destinatari)) return json({ error: { code: 500, message: 'Guasto simulato del servizio di posta' } }, 500);
    if (/apispenta@example\.com/i.test(m.destinatari)) {
      return json({ error: {
        code: 403, status: 'PERMISSION_DENIED',
        message: 'Gmail API has not been used in project 0 before or it is disabled. Enable it by visiting https://console.developers.google.com/apis/api/gmail.googleapis.com/overview?project=0 then retry.',
        details: [{ '@type': 'type.googleapis.com/google.rpc.ErrorInfo', reason: 'SERVICE_DISABLED', domain: 'googleapis.com' }],
      } }, 403);
    }
    if (/rifiuta@example\.com/i.test(m.destinatari) && etaToken(bearer) > 2000) {
      return json({ error: { code: 401, message: 'Invalid Credentials' } }, 401);
    }
    const { data, error } = await sb.from('posta_simulata').insert({ ...m, grezzo: raw, reale: false, esito: 'simulata' }).select('id').single();
    if (error) return json({ error: { code: 500, message: error.message } }, 500);
    return json({ id: 'finta-' + data.id, threadId: 'finta-' + data.id, labelIds: ['SENT'] });
  }

  return json({ error: 'percorso sconosciuto' }, 404);
});
