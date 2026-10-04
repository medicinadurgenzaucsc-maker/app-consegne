// Pubblica e configura le Edge Function nel SOLO progetto di collaudo.
//
//   node funzioni-collaudo.js pubblica   -> google-token (lo stesso sorgente della
//                                          produzione) + google-finto (finto Google)
//   node funzioni-collaudo.js configura  -> tabella posta_simulata, segreto del
//                                          finto Google, variabili GOOGLE_*,
//                                          cassaforte di collaudo già «autorizzata»
//
// Prima di pubblicare verifica la sintassi (il deploy non fa type-check e un
// errore di sintassi spegnerebbe la funzione). I segreti non vengono stampati.
const crypto = require('crypto');
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { stripTypeScriptTypes } = require('node:module');
const { collaudo } = require('./sb.js');

const RADICE = path.resolve(__dirname, process.env.RADICE_REPO || '../..');
const FUNZIONI = [
  { slug: 'google-token', verify_jwt: true },
  { slug: 'google-finto', verify_jwt: false },   // la chiama google-token con credenziali «alla Google», non con un JWT Supabase
];

function sorgente(slug) {
  const f = path.join(RADICE, 'supabase', 'functions', slug, 'index.ts');
  const src = fs.readFileSync(f, 'utf8').split('\r\n').join('\n');
  new vm.SourceTextModule(stripTypeScriptTypes(src), { identifier: slug });   // lancia se la sintassi è sbagliata
  return src;
}

async function pubblica() {
  const c = collaudo();
  for (const f of FUNZIONI) {
    const src = sorgente(f.slug);
    const fd = new FormData();
    fd.append('metadata', new Blob([JSON.stringify({ name: f.slug, entrypoint_path: 'index.ts', verify_jwt: f.verify_jwt })], { type: 'application/json' }));
    fd.append('file', new Blob([src], { type: 'application/typescript' }), 'index.ts');
    const r = await c.invia('/functions/deploy?slug=' + f.slug, { method: 'POST', body: fd });
    const t = await r.text();
    let j = null; try { j = JSON.parse(t); } catch (e) {}
    console.log(f.slug + ': HTTP ' + r.status + (j ? ' | versione ' + j.version + ' | verify_jwt ' + j.verify_jwt + ' | stato ' + j.status : ' | ' + t.slice(0, 200)));
    if (r.status >= 300) process.exitCode = 1;
  }
}

async function configura() {
  const c = collaudo();
  await c.query(`
    create table if not exists public.posta_simulata (
      id bigint generated always as identity primary key,
      ricevuta_il timestamptz not null default now(),
      mittente text, destinatari text, oggetto text, corpo_html text, grezzo text
    );
    comment on table public.posta_simulata is 'SOLO COLLAUDO: le mail che google-token avrebbe spedito, intercettate da google-finto.';
    alter table public.posta_simulata enable row level security;
    revoke all on table public.posta_simulata from PUBLIC, anon, authenticated;
    grant all on table public.posta_simulata to service_role;`);
  console.log('tabella posta_simulata: pronta (privata)');

  // Il segreto del finto Google: se la cassaforte ne ha già uno lo si riusa, così
  // ripetere «configura» non invalida nulla.
  const attuale = await c.query("select refresh_token from public.google_oauth where id = 'reparto'");
  let segreto = attuale[0] && attuale[0].refresh_token;
  if (!segreto || !/^finto_[0-9a-f]{48}$/.test(segreto)) segreto = 'finto_' + crypto.randomBytes(24).toString('hex');
  const base = 'https://' + c.ref + '.supabase.co/functions/v1/google-finto';
  const s = await c.api('POST', '/secrets', [
    { name: 'FINTO_SEGRETO', value: segreto },
    { name: 'GOOGLE_URL_TOKEN', value: base + '/token' },
    { name: 'GOOGLE_URL_USERINFO', value: base + '/userinfo' },
    { name: 'GOOGLE_URL_GMAIL_INVIO', value: base + '/gmail-invio' },
  ]);
  console.log('variabili delle funzioni: HTTP ' + s.stato);
  if (s.stato >= 300) throw new Error('variabili non impostate: ' + JSON.stringify(s.dati).slice(0, 200));

  const tag = '$s_' + crypto.randomBytes(4).toString('hex') + '$';
  const r = await c.query(`
    insert into public.google_oauth (id, client_secret, refresh_token, email, scopes, richiedi_utente)
    values ('reparto', 'segreto-client-finto-del-collaudo', ${tag}${segreto}${tag},
            (select lower(trim(valore)) from public.impostazioni where chiave = 'ACCOUNT_LOGIN'),
            'https://www.googleapis.com/auth/gmail.send', true)
    on conflict (id) do update set client_secret = excluded.client_secret, refresh_token = excluded.refresh_token,
      email = excluded.email, scopes = excluded.scopes, richiedi_utente = true,
      access_token = null, access_scad = null, updated_at = now();
    select (refresh_token is not null) as autorizzata, richiedi_utente, email from public.google_oauth where id = 'reparto';`);
  console.log('cassaforte di collaudo: ' + JSON.stringify(r[0]));
}

const cmd = process.argv[2];
(cmd === 'pubblica' ? pubblica() : cmd === 'configura' ? configura() : Promise.reject(new Error('uso: pubblica | configura')))
  .catch((e) => { console.log('ERRORE: ' + String(e.message).replace(/finto_[0-9a-f]+/g, 'finto_***')); process.exit(1); });
