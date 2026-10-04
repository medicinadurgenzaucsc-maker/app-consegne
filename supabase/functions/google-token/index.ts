// google-token v4 — cassaforte OAuth del reparto, SOLO mail.
// Il token Google non lascia mai il server: l'unica operazione possibile
// e' l'invio della mail dimissioni ('invia'), eseguito interamente qui.
//
// Azioni (POST JSON {azione:...}):
//  stato   -> {configurato, autorizzato[, email]}  (aperta: nulla di riservato)
//  config  -> {client_secret} (sostituzione solo conoscendo quello attuale)
//  scambia -> {code} scambio authorization code (redirect 'postmessage');
//             accetta SOLO l'account del reparto (impostazioni.ACCOUNT_LOGIN);
//             se l'impostazione manca rifiuta: non si registra un account
//             che non si puo' verificare
//  invia   -> {destinatari:[...], oggetto, html} spedisce via Gmail API
//
// CHI PUO' CHIAMARE. La chiave pubblica del sito basta a superare il
// verify_jwt del gateway, quindi da sola non prova nulla: config, scambia e
// invia pretendono che il Bearer sia la sessione di un utente AUTORIZZATO
// (stessa regola della RLS: is_autorizzato()). L'interruttore e' la colonna
// google_oauth.richiedi_utente: a false la verifica gira lo stesso ma finisce
// solo nel registro, cosi' si accende e si spegne con un UPDATE.
//
// GOOGLE. Client e indirizzi sono quelli veri; le variabili d'ambiente
// GOOGLE_* li sostituiscono SOLO nel progetto di collaudo, dove al posto di
// Google risponde la funzione «google-finto» e nessuna mail parte davvero.

import { createClient } from 'npm:@supabase/supabase-js@2';

const CLIENT_ID = Deno.env.get('GOOGLE_CLIENT_ID') || '170256871056-gchf386c3oic77ek2j5m3b1e5pbv6cre.apps.googleusercontent.com';
const URL_TOKEN = Deno.env.get('GOOGLE_URL_TOKEN') || 'https://oauth2.googleapis.com/token';
const URL_USERINFO = Deno.env.get('GOOGLE_URL_USERINFO') || 'https://www.googleapis.com/oauth2/v3/userinfo';
const URL_GMAIL_INVIO = Deno.env.get('GOOGLE_URL_GMAIL_INVIO') || 'https://gmail.googleapis.com/gmail/v1/users/me/messages/send';
const SCOPES_RICHIESTI = ['https://www.googleapis.com/auth/gmail.send'];
const AZIONI_RISERVATE = ['config', 'scambia', 'invia'];

const CORS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'authorization, x-client-info, apikey, content-type',
  'Access-Control-Allow-Methods': 'POST, OPTIONS',
};
const json = (body: unknown, status = 200) =>
  new Response(JSON.stringify(body), { status, headers: { ...CORS, 'Content-Type': 'application/json' } });

const SB_URL = Deno.env.get('SUPABASE_URL')!;
const sb = createClient(SB_URL, Deno.env.get('SUPABASE_SERVICE_ROLE_KEY')!);

async function riga() {
  const { data, error } = await sb.from('google_oauth').select('*').eq('id', 'reparto').maybeSingle();
  if (error) throw new Error('lettura cassaforte: ' + error.message);
  return data;
}
function scopesMancanti(scopes: string | null | undefined): string[] {
  const have = String(scopes || '').split(/\s+/);
  return SCOPES_RICHIESTI.filter((s) => !have.includes(s));
}

// Il chiamante e' un utente loggato e in lista? Si chiede al database con il
// SUO Bearer: la chiave pubblica (ruolo anon) e gli account fuori lista danno
// false, come pure qualunque errore.
async function chiamanteAutorizzato(req: Request): Promise<boolean> {
  const auth = req.headers.get('Authorization') || '';
  if (!/^Bearer\s+\S+/i.test(auth)) return false;
  const apikey = Deno.env.get('SUPABASE_ANON_KEY') || req.headers.get('apikey') || '';
  if (!apikey) return false;
  try {
    const c = createClient(SB_URL, apikey, {
      global: { headers: { Authorization: auth } },
      auth: { persistSession: false, autoRefreshToken: false },
    });
    const { data, error } = await c.rpc('is_autorizzato');
    return !error && data === true;
  } catch {
    return false;
  }
}

// Access token SOLO per uso interno: mai restituito al chiamante.
async function tokenInterno(r: Record<string, string>, forza = false): Promise<{ ok: true; tok: string } | { ok: false; resp: Response }> {
  if (!r || !r.client_secret || !r.refresh_token) {
    return { ok: false, resp: json({ errore: 'cassaforte non autorizzata', riautorizzare: true }, 401) };
  }
  if (!forza && r.access_token && r.access_scad && new Date(r.access_scad).getTime() > Date.now()) {
    return { ok: true, tok: r.access_token };
  }
  const tr = await fetch(URL_TOKEN, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      client_id: CLIENT_ID, client_secret: r.client_secret,
      refresh_token: r.refresh_token, grant_type: 'refresh_token',
    }),
  });
  const tok = await tr.json();
  if (!tr.ok || !tok.access_token) {
    const revocato = tok.error === 'invalid_grant';
    if (revocato) await sb.from('google_oauth').update({ refresh_token: null, access_token: null, access_scad: null, updated_at: new Date().toISOString() }).eq('id', 'reparto');
    return { ok: false, resp: json({ errore: revocato ? 'autorizzazione revocata: va concessa di nuovo' : ('refresh rifiutato: ' + (tok.error_description || tok.error || tr.status)), riautorizzare: revocato }, 401) };
  }
  await sb.from('google_oauth').update({
    access_token: tok.access_token,
    access_scad: new Date(Date.now() + Math.max(60, (tok.expires_in || 3600) - 300) * 1000).toISOString(),
    updated_at: new Date().toISOString(),
  }).eq('id', 'reparto');
  return { ok: true, tok: tok.access_token };
}

// ── MIME (replica del vecchio client, RFC 5322 + 2047) ──────────────
const enc = new TextEncoder();
function b64(bytes: Uint8Array): string {
  let bin = '';
  for (let i = 0; i < bytes.length; i += 0x8000) bin += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  return btoa(bin);
}
const b64utf8 = (s: string) => b64(enc.encode(s));
function encodeHeader(s: string): string {
  return /[^\x00-\x7F]/.test(s || '') ? '=?UTF-8?B?' + b64utf8(String(s)) + '?=' : (s || '');
}
function costruisciRaw(from: string, dest: string[], oggetto: string, html: string): string {
  const headers = [
    'From: ' + from,
    'To: ' + dest.join(', '),
    'Subject: ' + encodeHeader(oggetto || ''),
    'MIME-Version: 1.0',
    'Content-Type: text/html; charset=UTF-8',
    'Content-Transfer-Encoding: base64',
  ];
  const bodyB64 = b64utf8(html || '').replace(/(.{76})/g, '$1\r\n');
  const rawMessage = headers.join('\r\n') + '\r\n\r\n' + bodyB64;
  return b64utf8(rawMessage).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
}
const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/;

Deno.serve(async (req) => {
  if (req.method === 'OPTIONS') return new Response('ok', { headers: CORS });
  if (req.method !== 'POST') return json({ errore: 'solo POST' }, 405);
  let body: Record<string, unknown>;
  try { body = await req.json(); } catch { return json({ errore: 'JSON non valido' }, 400); }

  try {
    const r = await riga();
    const azione = String(body.azione || '');
    const utente = await chiamanteAutorizzato(req);
    const richiedi = !!(r && r.richiedi_utente);
    console.log(JSON.stringify({ ev: 'chiamata', azione, utente, richiedi }));
    if (richiedi && !utente && AZIONI_RISERVATE.includes(azione)) {
      return json({
        errore: 'Operazione riservata agli utenti del reparto: aggiorna la pagina (F5), accedi con l\'account autorizzato e riprova.',
        non_autorizzato: true,
      }, 403);
    }


    if (azione === 'stato') {
      return json({
        configurato: !!(r && r.client_secret),
        autorizzato: !!(r && r.refresh_token) && scopesMancanti(r?.scopes).length === 0,
        email: utente ? (r?.email || null) : null,
      });
    }

    if (azione === 'config') {
      const nuovo = String(body.client_secret || '').trim();
      if (!/^[\w~.-]{20,100}$/.test(nuovo)) return json({ errore: 'client secret non valido' }, 400);
      if (r && r.client_secret && r.client_secret !== String(body.secret_attuale || '').trim()) {
        return json({ errore: 'esiste già un client secret: per sostituirlo serve quello attuale' }, 403);
      }
      const { error } = await sb.from('google_oauth').upsert({ id: 'reparto', client_secret: nuovo, updated_at: new Date().toISOString() });
      if (error) return json({ errore: error.message }, 500);
      return json({ ok: true });
    }

    if (azione === 'scambia') {
      if (!r || !r.client_secret) return json({ errore: 'client secret non ancora configurato' }, 400);
      const code = String(body.code || '');
      if (!code) return json({ errore: 'code mancante' }, 400);
      const tr = await fetch(URL_TOKEN, {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: new URLSearchParams({
          code, client_id: CLIENT_ID, client_secret: r.client_secret,
          redirect_uri: 'postmessage', grant_type: 'authorization_code',
        }),
      });
      const tok = await tr.json();
      if (!tr.ok || !tok.access_token) {
        return json({ errore: 'scambio rifiutato da Google: ' + (tok.error_description || tok.error || tr.status) }, 400);
      }
      const ui = await fetch(URL_USERINFO, {
        headers: { Authorization: 'Bearer ' + tok.access_token },
      }).then((x) => x.json()).catch(() => ({}));
      const email = String(ui.email || '').toLowerCase();
      const { data: imp, error: errImp } = await sb.from('impostazioni').select('valore').eq('chiave', 'ACCOUNT_LOGIN').maybeSingle();
      const atteso = String(imp?.valore || '').toLowerCase().trim();
      if (errImp || !atteso) {
        return json({ errore: "impostazione ACCOUNT_LOGIN non leggibile: impossibile verificare l'account" }, 500);
      }
      if (!email || email !== atteso) {
        return json({ errore: 'account non ammesso: autorizza con ' + atteso }, 403);
      }
      if (scopesMancanti(tok.scope).length > 0) {
        return json({ errore: 'manca il permesso di invio mail: rifai l\'autorizzazione spuntando tutte le caselle' }, 400);
      }
      if (!tok.refresh_token && !r.refresh_token) {
        return json({ errore: 'Google non ha fornito il refresh token: riprova' }, 400);
      }
      const { error } = await sb.from('google_oauth').upsert({
        id: 'reparto',
        refresh_token: tok.refresh_token || r.refresh_token,
        email,
        scopes: tok.scope || '',
        access_token: tok.access_token,
        access_scad: new Date(Date.now() + Math.max(60, (tok.expires_in || 3600) - 300) * 1000).toISOString(),
        updated_at: new Date().toISOString(),
      });
      if (error) return json({ errore: error.message }, 500);
      return json({ ok: true, email });
    }

    if (azione === 'invia') {
      const dest = Array.isArray(body.destinatari) ? (body.destinatari as string[]).map((d) => String(d).trim()).filter(Boolean) : [];
      const oggetto = String(body.oggetto || '').trim();
      const html = String(body.html || '');
      if (!dest.length || dest.length > 30) return json({ errore: 'destinatari: da 1 a 30 indirizzi' }, 400);
      const nonValidi = dest.filter((d) => !EMAIL_RE.test(d));
      if (nonValidi.length) return json({ errore: 'indirizzi non validi: ' + nonValidi.join(', ') }, 400);
      if (!oggetto || oggetto.length > 500) return json({ errore: 'oggetto mancante o troppo lungo' }, 400);
      if (!html || html.length > 500000) return json({ errore: 'corpo mancante o troppo grande' }, 400);

      let t = await tokenInterno(r);
      if (!t.ok) return t.resp;
      const raw = costruisciRaw(r.email || '', dest, oggetto, html);
      let g = await fetch(URL_GMAIL_INVIO, {
        method: 'POST',
        headers: { Authorization: 'Bearer ' + t.tok, 'Content-Type': 'application/json' },
        body: JSON.stringify({ raw }),
      });
      if (g.status === 401) {
        t = await tokenInterno(r, true);
        if (!t.ok) return t.resp;
        g = await fetch(URL_GMAIL_INVIO, {
          method: 'POST',
          headers: { Authorization: 'Bearer ' + t.tok, 'Content-Type': 'application/json' },
          body: JSON.stringify({ raw }),
        });
      }
      if (!g.ok) {
        const b = await g.json().catch(() => ({} as Record<string, { message?: string }>));
        return json({ errore: (b.error && b.error.message) || ('Gmail HTTP ' + g.status) }, 502);
      }
      return json({ ok: true, destinatari: dest.length });
    }

    return json({ errore: 'azione sconosciuta' }, 400);
  } catch (e) {
    return json({ errore: String((e as Error)?.message || e) }, 500);
  }
});
