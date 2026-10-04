// google-finto — esiste SOLO nel progetto di collaudo.
// Fa la parte di Google per la funzione google-token: rilascia token finti,
// dice chi e' l'utente e «spedisce» la mail scrivendola nella tabella
// posta_simulata invece di inviarla. Cosi' nel collaudo gira tutto il codice
// vero di google-token (rinnovo del token, costruzione del messaggio, nuovo
// tentativo sul 401) senza che parta mai una mail.
//
// Accetta solo chi presenta il segreto del collaudo (variabile FINTO_SEGRETO):
// e' lo stesso valore salvato come refresh token nella cassaforte di collaudo,
// quindi lo conosce soltanto google-token.
//
// Per provare i percorsi d'errore: un destinatario «rifiuta@example.com» fa
// rispondere 401 se il token d'accesso ha piu' di due secondi (google-token
// deve rinnovarlo e riprovare: col token appena rilasciato l'invio passa),
// «guasto@example.com» fa rispondere 500; il codice di
// autorizzazione «<segreto>.intruso» simula il consenso dato da un account
// Google diverso da quello del reparto.

import { createClient } from 'npm:@supabase/supabase-js@2';

const SEGRETO = Deno.env.get('FINTO_SEGRETO') || '';
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

// Ogni token d'accesso porta l'istante del rilascio: serve a simulare la
// scadenza senza tenere stato in memoria (le chiamate non lo condividono).
const nuovoToken = () => SEGRETO + '.t' + Date.now();
const etaToken = (t: string) => Date.now() - Number(t.split('.t')[1] || 0);

Deno.serve(async (req) => {
  if (!SEGRETO) return json({ error: 'segreto del collaudo non configurato' }, 500);
  const percorso = new URL(req.url).pathname.replace(/^.*\/google-finto/, '') || '/';

  if (percorso === '/token' && req.method === 'POST') {
    const f = new URLSearchParams(await req.text());
    if (f.get('grant_type') === 'refresh_token' && f.get('refresh_token') === SEGRETO) {
      return json({ access_token: nuovoToken(), expires_in: 3599, scope: AMBITI, token_type: 'Bearer' });
    }
    if (f.get('grant_type') === 'authorization_code' && f.get('code') === SEGRETO) {
      return json({ access_token: nuovoToken(), refresh_token: SEGRETO, expires_in: 3599, scope: AMBITI, token_type: 'Bearer' });
    }
    if (f.get('grant_type') === 'authorization_code' && f.get('code') === SEGRETO + '.intruso') {
      return json({ access_token: SEGRETO + '.intruso', refresh_token: SEGRETO + '.intruso', expires_in: 3599, scope: AMBITI, token_type: 'Bearer' });
    }
    return json({ error: 'invalid_grant', error_description: 'Token has been expired or revoked.' }, 400);
  }

  const bearer = (req.headers.get('Authorization') || '').replace(/^Bearer\s+/i, '');
  if (percorso === '/userinfo' && bearer === SEGRETO + '.intruso') {
    return json({ email: 'intruso@example.com', email_verified: true });
  }
  if (!bearer.startsWith(SEGRETO + '.t')) return json({ error: { code: 401, message: 'Invalid Credentials' } }, 401);

  if (percorso === '/userinfo') {
    const { data } = await sb.from('impostazioni').select('valore').eq('chiave', 'ACCOUNT_LOGIN').maybeSingle();
    return json({ email: String(data?.valore || '').trim(), email_verified: true });
  }

  if (percorso === '/gmail-invio' && req.method === 'POST') {
    let raw = '';
    try { raw = String((await req.json()).raw || ''); } catch { return json({ error: { message: 'corpo non valido' } }, 400); }
    const m = leggiMessaggio(raw);
    if (/guasto@example\.com/i.test(m.destinatari)) return json({ error: { code: 500, message: 'Guasto simulato del servizio di posta' } }, 500);
    if (/rifiuta@example\.com/i.test(m.destinatari) && etaToken(bearer) > 2000) {
      return json({ error: { code: 401, message: 'Invalid Credentials' } }, 401);
    }
    const { data, error } = await sb.from('posta_simulata').insert({ ...m, grezzo: raw }).select('id').single();
    if (error) return json({ error: { code: 500, message: error.message } }, 500);
    return json({ id: 'finta-' + data.id, threadId: 'finta-' + data.id, labelIds: ['SENT'] });
  }

  return json({ error: 'percorso sconosciuto' }, 404);
});
