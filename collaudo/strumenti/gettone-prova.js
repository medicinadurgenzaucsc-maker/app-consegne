// Identità delle PROVE AUTOMATICHE sul solo database di collaudo.
// Genera un token di sessione firmato con la chiave del progetto di collaudo
// (letta dall'API di gestione, mai stampata) per l'utente fittizio
// «collaudo-automatico@example.com», che sta nella lista autorizzati del SOLO
// collaudo. Non esiste un account dietro: è un token a scadenza, valido solo lì.
//
//   node gettone-prova.js [ore]   -> scrive i file per il banco di prova locale
//
// Uscite (tutte fuori dal repository, mai da pubblicare):
//   <banco>/banco-avvio.js   sessione nel localStorage + service worker spento
//   .gettone                 il solo token, per le prove da riga di comando
const crypto = require('crypto');
const fs = require('fs');
const path = require('path');
const { collaudo } = require('./sb.js');

const EMAIL = 'collaudo-automatico@example.com';
const UTENTE_ID = '00000000-0000-4000-8000-00000c011a0d';
const b64 = (o) => Buffer.from(typeof o === 'string' ? o : JSON.stringify(o)).toString('base64url');

async function genera(ore) {
  const c = collaudo();
  const pg = await c.api('GET', '/postgrest');
  const segreto = pg.dati && pg.dati.jwt_secret;
  if (!segreto) throw new Error('chiave di firma del collaudo non disponibile');
  const ora = Math.floor(Date.now() / 1000);
  const scade = ora + Math.round(ore * 3600);
  const claims = {
    iss: 'https://' + c.ref + '.supabase.co/auth/v1',
    sub: UTENTE_ID, aud: 'authenticated', exp: scade, iat: ora,
    email: EMAIL, phone: '',
    app_metadata: { provider: 'google', providers: ['google'] },
    user_metadata: { email: EMAIL, full_name: 'Collaudo Automatico' },
    role: 'authenticated', aal: 'aal1',
    amr: [{ method: 'oauth', timestamp: ora }],
    session_id: crypto.randomUUID(), is_anonymous: false,
  };
  const corpo = b64({ alg: 'HS256', typ: 'JWT' }) + '.' + b64(claims);
  const token = corpo + '.' + crypto.createHmac('sha256', segreto).update(corpo).digest('base64url');
  const adesso = new Date().toISOString();
  const sessione = {
    access_token: token, token_type: 'bearer', expires_in: scade - ora, expires_at: scade,
    refresh_token: 'gettone-di-prova-non-rinnovabile',
    user: {
      id: UTENTE_ID, aud: 'authenticated', role: 'authenticated', email: EMAIL, email_confirmed_at: adesso, phone: '',
      app_metadata: claims.app_metadata, user_metadata: claims.user_metadata, identities: [],
      created_at: adesso, updated_at: adesso, is_anonymous: false,
    },
  };
  return { token, sessione, scade, ref: c.ref };
}

module.exports = { genera, EMAIL };

if (require.main === module) {
  const ore = Number(process.argv[2] || 10);
  const banco = process.argv[3];
  genera(ore).then(({ token, sessione, scade, ref }) => {
    fs.writeFileSync(path.join(__dirname, '.gettone'), token);
    if (banco) {
      const js = '// BANCO LOCALE — generato da gettone-prova.js: non pubblicare.\n'
        + '(function () {\n'
        + '  try { localStorage.setItem(' + JSON.stringify('sb-' + ref + '-auth-token') + ', ' + JSON.stringify(JSON.stringify(sessione)) + '); } catch (e) {}\n'
        + '  try { navigator.serviceWorker.register = function () { return Promise.reject(new Error("service worker spento nel banco")); }; } catch (e) {}\n'
        + '  window.__erroriBanco = [];\n'
        + '  window.addEventListener("error", function (e) { window.__erroriBanco.push("UNCAUGHT: " + e.message); });\n'
        + '  window.addEventListener("unhandledrejection", function (e) { window.__erroriBanco.push("PROMISE: " + String((e.reason && e.reason.message) || e.reason)); });\n'
        + '})();\n';
      fs.writeFileSync(path.join(banco, 'banco-avvio.js'), js);
    }
    console.log('gettone di prova per ' + EMAIL + ' valido fino alle ' + new Date(scade * 1000).toLocaleTimeString('it-IT') + (banco ? ' — banco-avvio.js scritto' : ''));
  }).catch((e) => { console.log('ERRORE: ' + e.message); process.exit(1); });
}
