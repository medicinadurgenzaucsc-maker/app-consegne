// google-token v6 — cassaforte OAuth del reparto, SOLO mail.
// Il token Google non lascia mai il server: l'unica operazione possibile
// e' l'invio della mail dimissioni ('invia'), eseguito interamente qui.
//
// Azioni (POST JSON {azione:...}):
//  stato   -> {configurato, autorizzato[, email, mittente]}  (aperta: nulla di
//             riservato). «autorizzato» = ci sono i permessi di invio PER IL
//             MITTENTE configurato; email e mittente solo a chi e' in lista.
//             Con {verifica:true}, e solo per chi e' in lista, i permessi si
//             controllano DAL VIVO: si chiede a Google un token nuovo. La
//             risposta porta allora «verificato»: true = «autorizzato» e' la
//             verita' di adesso (un consenso revocato da' false); false =
//             Google non ha risposto, «problema» dice perche'
//  config  -> {client_secret} salva il client secret, SOLO se Google lo
//             riconosce per questo client: lo si prova subito (rinnovo del
//             token se c'e' un consenso, altrimenti uno scambio con un codice
//             inventato). Uno sbagliato non prende mai il posto di quello
//             buono. Il secret custodito NON viene mai restituito, ne' intero
//             ne' in parte: per cambiarlo se ne scrive uno nuovo. Un ID client
//             incollato al posto del secret viene riconosciuto e rifiutato.
//             Risponde {ok, autorizzato}
//  scambia -> {code} scambio authorization code (redirect 'postmessage');
//             accetta SOLO l'account configurato come mittente; se il
//             mittente non e' determinabile rifiuta: non si registra un
//             account che non si puo' verificare
//  invia   -> {destinatari:[...], oggetto, html[, mittente]} spedisce via Gmail
//             API, solo se la cassaforte e' autorizzata per il mittente
//             configurato. «mittente» e' quello che la pagina ha mostrato a
//             chi scrive: se nel frattempo e' cambiato la mail non parte
//
// IL MITTENTE. E' impostazioni.MAIL_DIMISSIONI_MITTENTE (menu «Impostazioni
// email» dell'app); finche' nessuno lo imposta vale l'account del reparto
// (impostazioni.ACCOUNT_LOGIN). La cassaforte custodisce UN solo consenso
// Google: se il mittente cambia, il consenso va dato di nuovo da quell'account.
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

// L'indirizzo da cui deve partire la mail: quello scelto in «Impostazioni
// email», altrimenti l'account del reparto. Stringa vuota = non determinabile
// (impostazioni illeggibili, oppure un valore che non e' un indirizzo): in
// quel caso non si autorizza e non si spedisce.
async function mittenteConfigurato(): Promise<string> {
  const { data, error } = await sb.from('impostazioni').select('chiave,valore').in('chiave', ['MAIL_DIMISSIONI_MITTENTE', 'ACCOUNT_LOGIN']);
  if (error || !data) return '';
  const valore = (chiave: string) => String((data.find((x) => x.chiave === chiave) || {}).valore || '').toLowerCase().trim();
  const scelto = valore('MAIL_DIMISSIONI_MITTENTE') || valore('ACCOUNT_LOGIN');
  return MITTENTE_RE.test(scelto) ? scelto : '';
}
// Solo un indirizzo «semplice»: il valore arriva da una tabella che ogni utente
// dell'app puo' scrivere e torna indietro nelle risposte.
const MITTENTE_RE = /^[a-z0-9._%+-]{1,64}@[a-z0-9.-]{1,190}\.[a-z]{2,24}$/;
const stessoIndirizzo = (a: unknown, b: string) => String(a || '').toLowerCase().trim() === b;

// Chiede a Google se «secret» e' il client secret di questo client, senza
// salvare nulla.
//  - con un consenso custodito: rinnovo del token (prova anche il consenso)
//  - senza: uno scambio con un codice inventato. Google controlla prima il
//    client (invalid_client = secret sbagliato) e solo dopo il codice
//    (invalid_grant = secret giusto, codice ovviamente no)
// valido: true / false / null (Google non ha saputo rispondere).
async function provaSegreto(secret: string, r: Record<string, string> | null): Promise<{
  valido: boolean | null; consenso?: 'vivo' | 'revocato'; accesso?: { token: string; scad: string }; dettaglio?: string;
}> {
  const chiedi = async (campi: Record<string, string>) => {
    const tr = await fetch(URL_TOKEN, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({ client_id: CLIENT_ID, client_secret: secret, ...campi }),
    });
    const tok = await tr.json().catch(() => ({} as Record<string, any>));
    return { ok: tr.ok, stato: tr.status, tok };
  };
  const dettaglio = (x: { stato: number; tok: Record<string, any> }) => String(x.tok.error_description || x.tok.error || ('HTTP ' + x.stato));
  try {
    if (r && r.refresh_token) {
      const a = await chiedi({ grant_type: 'refresh_token', refresh_token: r.refresh_token });
      if (a.ok && a.tok.access_token) {
        return {
          valido: true, consenso: 'vivo',
          accesso: { token: a.tok.access_token, scad: new Date(Date.now() + Math.max(60, (a.tok.expires_in || 3600) - 300) * 1000).toISOString() },
        };
      }
      if (a.tok.error === 'invalid_client') return { valido: false, dettaglio: dettaglio(a) };
      if (a.tok.error === 'invalid_grant') return { valido: true, consenso: 'revocato' };
      return { valido: null, dettaglio: dettaglio(a) };
    }
    const b = await chiedi({ grant_type: 'authorization_code', code: 'verifica-del-client-secret', redirect_uri: 'postmessage' });
    if (b.tok.error === 'invalid_grant') return { valido: true };
    if (b.tok.error === 'invalid_client') return { valido: false, dettaglio: dettaglio(b) };
    return { valido: null, dettaglio: dettaglio(b) };
  } catch (e) {
    return { valido: null, dettaglio: String((e as Error)?.message || e) };
  }
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
    // Google non riconosce (piu') il client secret custodito: il consenso non
    // si tocca, va solo inserito un secret valido.
    if (tok.error === 'invalid_client') {
      return { ok: false, resp: json({ errore: 'Google non riconosce il client secret custodito (' + (tok.error_description || 'invalid_client') + '): va inserito di nuovo da «Impostazioni email»', segreto_errato: true, riautorizzare: false }, 401) };
    }
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
      // A interruttore spento le azioni riservate sono aperte: anche qui si
      // risponde per intero, altrimenti l'app non potrebbe lavorare.
      const fidato = utente || !richiedi;
      const mittente = await mittenteConfigurato();
      const risposta: Record<string, unknown> = {
        configurato: !!(r && r.client_secret),
        autorizzato: !!(r && r.refresh_token) && scopesMancanti(r?.scopes).length === 0
          && !!mittente && stessoIndirizzo(r?.email, mittente),
        email: fidato ? (r?.email || null) : null,
        mittente: fidato ? (mittente || null) : null,
      };
      // Verifica dal vivo. Il consenso custodito puo' essere stato revocato
      // (da Google, o dal titolare dell'account) senza che qui se ne sappia
      // nulla: lo si scopre chiedendo a Google un token nuovo. L'app lo fa
      // PRIMA di mostrare la procedura di invio.
      if (body.verifica === true && fidato) {
        risposta.verificato = true;
        if (risposta.autorizzato) {
          let esito: { ok: boolean; riautorizzare?: boolean; segreto_errato?: boolean; errore?: string };
          try {
            const t = await tokenInterno(r, true);
            esito = t.ok ? { ok: true } : { ok: false, ...(await t.resp.json()) };
          } catch (e) {
            esito = { ok: false, errore: String((e as Error)?.message || e) };
          }
          if (!esito.ok && esito.riautorizzare) {
            risposta.autorizzato = false;            // revocato: va concesso di nuovo
          } else if (!esito.ok && esito.segreto_errato) {
            risposta.autorizzato = false;            // il secret custodito non vale piu': va inserito di nuovo
            risposta.segreto_errato = true;
            risposta.problema = esito.errore;
          } else if (!esito.ok) {
            risposta.verificato = false;             // Google non ha risposto: non si sa
            risposta.problema = esito.errore || 'Google non risponde';
          }
        }
      }
      return json(risposta);
    }

    if (azione === 'config') {
      const nuovo = String(body.client_secret || '').trim();
      if (/\.apps\.googleusercontent\.com$/i.test(nuovo) || nuovo === CLIENT_ID) {
        return json({ errore: 'questo è l\'ID client, non il client secret: il secret è un\'altra stringa, di solito comincia con «GOCSPX-»', id_client: true }, 400);
      }
      if (!/^[\w~.-]{20,100}$/.test(nuovo)) {
        return json({ errore: 'il client secret non ha la forma attesa (da 20 a 100 caratteri, senza spazi)', formato: true }, 400);
      }
      // Si salva solo cio' che Google riconosce: chi scrive un secret sbagliato
      // lo sa subito, e quello custodito resta al suo posto.
      const p = await provaSegreto(nuovo, r);
      if (p.valido === false) {
        return json({ errore: 'Google non riconosce questo client secret' + (p.dettaglio ? ' («' + p.dettaglio + '»)' : '') + ': non è stato salvato. Controlla di aver copiato il secret del client giusto e riprova.', segreto_errato: true }, 400);
      }
      if (p.valido === null) {
        return json({ errore: 'Google non ha saputo verificare il client secret' + (p.dettaglio ? ' (' + p.dettaglio + ')' : '') + ': non è stato salvato, riprova fra qualche minuto.', non_verificato: true }, 502);
      }
      const campi: Record<string, unknown> = { id: 'reparto', client_secret: nuovo, updated_at: new Date().toISOString() };
      if (p.accesso) { campi.access_token = p.accesso.token; campi.access_scad = p.accesso.scad; }
      const { error } = await sb.from('google_oauth').upsert(campi);
      if (error) return json({ errore: error.message }, 500);
      const mittente = await mittenteConfigurato();
      return json({
        ok: true,
        // con questo secret il consenso custodito funziona, ed e' quello del mittente?
        autorizzato: p.consenso === 'vivo' && scopesMancanti(r?.scopes).length === 0
          && !!mittente && stessoIndirizzo(r?.email, mittente),
      });
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
        if (tok.error === 'invalid_client') {
          return json({ errore: 'Google non riconosce il client secret custodito (' + (tok.error_description || 'invalid_client') + '): va inserito di nuovo', segreto_errato: true }, 400);
        }
        return json({ errore: 'scambio rifiutato da Google: ' + (tok.error_description || tok.error || tr.status) }, 400);
      }
      const ui = await fetch(URL_USERINFO, {
        headers: { Authorization: 'Bearer ' + tok.access_token },
      }).then((x) => x.json()).catch(() => ({}));
      const email = String(ui.email || '').toLowerCase();
      const atteso = await mittenteConfigurato();
      if (!atteso) {
        return json({ errore: "mittente della mail non determinabile (impostazioni MAIL_DIMISSIONI_MITTENTE e ACCOUNT_LOGIN): impossibile verificare l'account" }, 500);
      }
      if (!email || email !== atteso) {
        return json({
          errore: 'account non ammesso: il consenso è stato dato con «' + (email || 'un account senza indirizzo') + '», ma il mittente impostato è «' + atteso + '»',
          account: email || null, mittente: atteso,
        }, 403);
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

      // Si spedisce solo a nome del mittente configurato: se la cassaforte
      // custodisce il consenso di un altro account non si prova nemmeno.
      const mittente = await mittenteConfigurato();
      if (!mittente) return json({ errore: 'mittente della mail non determinabile: aprire «Impostazioni email»' }, 409);
      // Chi scrive ha visto un mittente: se non e' piu' quello, non si spedisce
      // a nome di un altro senza che lo sappia.
      const mostrato = String(body.mittente || '').toLowerCase().trim();
      if (mostrato && mostrato !== mittente) {
        return json({
          errore: 'il mittente della mail è cambiato (ora è ' + mittente + '): la mail non è partita. Chiudi e riapri «Invia mail dimissioni».',
          mittente_cambiato: true, mittente,
        }, 409);
      }
      if (r && r.refresh_token && !stessoIndirizzo(r.email, mittente)) {
        return json({ errore: 'mancano i permessi di invio per ' + mittente, riautorizzare: true, mittente }, 401);
      }

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
        const b = await g.json().catch(() => ({} as Record<string, any>));
        const messaggio = String((b.error && b.error.message) || '');
        // Le Gmail API vanno accese una volta nel progetto Google dell'app:
        // l'errore di Google e' in inglese e non dice dove mettere le mani.
        if (g.status === 403 && (/SERVICE_DISABLED/.test(JSON.stringify((b.error && b.error.details) || '')) || /has not been used in project|it is disabled/i.test(messaggio))) {
          return json({
            errore: 'Nel progetto Google di questa app le Gmail API non sono attive: la mail non è partita. Vanno abilitate nella console Google Cloud (API e servizi, Libreria, Gmail API, Abilita); poi si riprova dopo un paio di minuti.',
            api_spenta: true,
          }, 502);
        }
        return json({ errore: messaggio || ('Gmail HTTP ' + g.status) }, 502);
      }
      return json({ ok: true, destinatari: dest.length, mittente });
    }

    return json({ errore: 'azione sconosciuta' }, 400);
  } catch (e) {
    return json({ errore: String((e as Error)?.message || e) }, 500);
  }
});
