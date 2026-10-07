// Pubblica e configura le Edge Function nel SOLO progetto di collaudo.
//
//   node funzioni-collaudo.js pubblica   -> google-token (lo stesso sorgente della
//                                          produzione) + google-finto (finto Google)
//   node funzioni-collaudo.js configura  -> tabella posta_simulata, segreto del
//                                          finto Google, variabili GOOGLE_* (con
//                                          una chiave nuova nell'indirizzo del
//                                          finto Google, che solo google-token
//                                          conosce),
//                                          credenziali finte per le prove (riga
//                                          «prova» della cassaforte: la riga
//                                          «reparto» non viene toccata, vedi
//                                          cassaforte-collaudo.js)
//
// Prima di pubblicare verifica la sintassi (il deploy non fa type-check e un
// errore di sintassi spegnerebbe la funzione). I segreti non vengono stampati.
const crypto = require('crypto');
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { stripTypeScriptTypes } = require('node:module');
const { collaudo } = require('./sb.js');
const { stato: statoCassaforte, CLIENT_FINTO } = require('./cassaforte-collaudo.js');

const RADICE = path.resolve(__dirname, process.env.RADICE_REPO || '../..');
const FUNZIONI = [
  { slug: 'google-token', verify_jwt: true },
  { slug: 'google-finto', verify_jwt: false },   // la chiama google-token con credenziali «alla Google», non con un JWT Supabase
];

// La sintassi si controlla come modulo. vm.SourceTextModule esiste solo se node
// parte con --experimental-vm-modules: senza, si fa controllare un file
// provvisorio a «node --check» (che non risolve gli import, li legge soltanto).
function verificaSintassi(slug, js) {
  if (typeof vm.SourceTextModule === 'function') { new vm.SourceTextModule(js, { identifier: slug }); return; }
  const provvisorio = path.join(require('os').tmpdir(), 'sintassi-' + slug + '-' + process.pid + '.mjs');
  fs.writeFileSync(provvisorio, js);
  try {
    require('child_process').execFileSync(process.execPath, ['--check', provvisorio], { stdio: 'pipe' });
  } catch (e) {
    throw new Error('sintassi di ' + slug + ': ' + String(e.stderr || e.message).split('\n').filter(Boolean).slice(0, 6).join(' | '));
  } finally {
    try { fs.unlinkSync(provvisorio); } catch (e2) { /* niente da togliere */ }
  }
}

function sorgente(slug) {
  const f = path.join(RADICE, 'supabase', 'functions', slug, 'index.ts');
  const src = fs.readFileSync(f, 'utf8').split('\r\n').join('\n');
  verificaSintassi(slug, stripTypeScriptTypes(src));   // lancia se la sintassi è sbagliata
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
    alter table public.posta_simulata add column if not exists reale boolean not null default false;
    alter table public.posta_simulata add column if not exists esito text;
    comment on table public.posta_simulata is 'SOLO COLLAUDO: le mail passate da google-finto. reale = false: simulate, mai partite; reale = true: copia di una mail spedita davvero con il consenso vero.';
    alter table public.posta_simulata enable row level security;
    revoke all on table public.posta_simulata from PUBLIC, anon, authenticated;
    grant all on table public.posta_simulata to service_role;`);
  console.log('tabella posta_simulata: pronta (privata)');

  // Il segreto del finto Google: se la cassaforte ne ha già uno lo si riusa, così
  // ripetere «configura» non invalida nulla.
  const attuali = await c.query("select id, refresh_token from public.google_oauth where id in ('prova', 'reparto')");
  const buono = (id) => { const x = attuali.find((y) => y.id === id); return x && /^finto_[0-9a-f]{48}$/.test(x.refresh_token || '') ? x.refresh_token : null; };
  const segreto = buono('prova') || buono('reparto') || 'finto_' + crypto.randomBytes(24).toString('hex');
  // Il client Google del collaudo (pubblico: sta in api.js). Serve quando il
  // consenso è vero: lo scambio del codice deve nominare lo stesso client che
  // l'ha chiesto dal browser.
  const client = (/collaudo:\s*\{[\s\S]*?googleClientId:\s*'([^']+)'/.exec(fs.readFileSync(path.join(RADICE, 'docs', 'js', 'api.js'), 'utf8')) || [])[1];
  if (!client) throw new Error('client Google del collaudo non trovato in docs/js/api.js');
  // google-finto gira richieste a Google e non verifica il JWT: risponde solo
  // sotto un indirizzo che contiene questa chiave, così non è un passaggio
  // aperto. La chiave cambia a ogni «configura» (non si può rileggere) e vive
  // solo nelle variabili delle due funzioni.
  const chiave = crypto.randomBytes(20).toString('hex');
  const base = 'https://' + c.ref + '.supabase.co/functions/v1/google-finto/k/' + chiave;
  const s = await c.api('POST', '/secrets', [
    { name: 'FINTO_SEGRETO', value: segreto },
    { name: 'FINTO_CHIAVE', value: chiave },
    { name: 'GOOGLE_URL_TOKEN', value: base + '/token' },
    { name: 'GOOGLE_URL_USERINFO', value: base + '/userinfo' },
    { name: 'GOOGLE_URL_GMAIL_INVIO', value: base + '/gmail-invio' },
    { name: 'GOOGLE_CLIENT_ID', value: client },
  ]);
  console.log('variabili delle funzioni: HTTP ' + s.stato);
  if (s.stato >= 300) throw new Error('variabili non impostate: ' + JSON.stringify(s.dati).slice(0, 200));

  const tag = '$s_' + crypto.randomBytes(4).toString('hex') + '$';
  await c.query(`
    insert into public.google_oauth (id, client_secret, refresh_token, scopes, richiedi_utente)
    values ('prova', '${CLIENT_FINTO}', ${tag}${segreto}${tag}, 'https://www.googleapis.com/auth/gmail.send', true)
    on conflict (id) do update set client_secret = excluded.client_secret, refresh_token = excluded.refresh_token,
      scopes = excluded.scopes, access_token = null, access_scad = null, updated_at = now();`);
  const s2 = await statoCassaforte();
  console.log('credenziali finte per le prove: pronte | cassaforte del collaudo: ' + s2.modo + (s2.consenso ? ' (consenso di ' + s2.email + ')' : '') + (s2.veraInAttesa ? ' — consenso vero messo da parte' : ''));
}

const cmd = process.argv[2];
(cmd === 'pubblica' ? pubblica() : cmd === 'configura' ? configura() : Promise.reject(new Error('uso: pubblica | configura')))
  .catch((e) => { console.log('ERRORE: ' + String(e.message).replace(/finto_[0-9a-f]+/g, 'finto_***').replace(/\/k\/[0-9a-f]+/g, '/k/***')); process.exit(1); });
