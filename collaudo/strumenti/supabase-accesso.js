// Accesso al progetto Supabase di COLLAUDO senza far passare chiavi dalla chat.
// Stesso protocollo del comando ufficiale «supabase login» (sorgente pubblico:
// supabase/cli, apps/cli/src/command-internal/ensure-login.ts): qui si genera
// una coppia di chiavi, l'utente apre il link da loggato e legge un codice di
// verifica; con quel codice si ritira il token, cifrato per la nostra chiave.
//
//   node supabase-accesso.js avvia              -> stampa il link da aprire
//   node supabase-accesso.js completa <codice>  -> ritira il token e lo registra
//
// Il token non viene MAI stampato: finisce solo nella configurazione locale
// (voce MCP «supabase-collaudo»), come quelli degli altri progetti.
const crypto = require('crypto');
const fs = require('fs');
const path = require('path');
const { spawnSync } = require('child_process');

const REF = 'rqvohwpthhumydpbwktq';
const API = 'https://api.supabase.com';
const DASH = 'https://supabase.com/dashboard';
const NOME_TOKEN = 'claude_collaudo_consegne';
const NOME_MCP = 'supabase-collaudo';
const STATO = path.join(__dirname, '.accesso-supabase.json');
const maschera = (s) => String(s || '').replace(/sbp_[A-Za-z0-9_]+/g, 'sbp_***');

async function avvia() {
  const ecdh = crypto.createECDH('prime256v1');
  ecdh.generateKeys();
  const sessione = crypto.randomUUID();
  fs.writeFileSync(STATO, JSON.stringify({ sessione, priv: ecdh.getPrivateKey('hex'), creato: Date.now() }));
  console.log(`${DASH}/cli/login?session_id=${sessione}&token_name=${NOME_TOKEN}&public_key=${ecdh.getPublicKey('hex', 'uncompressed')}`);
}

async function completa(codice) {
  if (!fs.existsSync(STATO)) { console.log('ESITO: nessun accesso avviato (manca lo stato locale)'); process.exit(2); }
  const st = JSON.parse(fs.readFileSync(STATO, 'utf8'));
  const cod = String(codice || '').trim();
  if (!cod) { console.log('ESITO: codice mancante'); process.exit(2); }

  const r = await fetch(`${API}/platform/cli/login/${st.sessione}?device_code=${encodeURIComponent(cod)}`, {
    headers: { 'User-Agent': 'SupabaseCLI/collaudo-consegne' },
  });
  if (r.status !== 200) {
    console.log('ESITO: codice rifiutato (HTTP ' + r.status + '): ' + maschera((await r.text()).slice(0, 200)));
    process.exit(3);
  }
  const j = await r.json();
  let token;
  try {
    const ecdh = crypto.createECDH('prime256v1');
    ecdh.setPrivateKey(Buffer.from(st.priv, 'hex'));
    const segreto = ecdh.computeSecret(Buffer.from(j.public_key, 'hex'));
    const cifrato = String(j.access_token);
    const d = crypto.createDecipheriv('aes-256-gcm', segreto, Buffer.from(j.nonce, 'hex'));
    d.setAuthTag(Buffer.from(cifrato.slice(-32), 'hex'));
    token = Buffer.concat([d.update(Buffer.from(cifrato.slice(0, -32), 'hex')), d.final()]).toString('utf8');
  } catch (e) {
    console.log('ESITO: impossibile decifrare il token (' + e.message + ')');
    process.exit(4);
  }

  // Il token vede davvero il progetto di collaudo? (difesa dall'account sbagliato)
  const p = await fetch(`${API}/v1/projects/${REF}`, { headers: { Authorization: 'Bearer ' + token } });
  if (p.status !== 200) {
    console.log('ESITO: il token ottenuto NON vede il progetto ' + REF + ' (HTTP ' + p.status + '): probabilmente il link è stato aperto da un altro account Supabase.');
    process.exit(5);
  }
  const prj = await p.json();
  console.log('progetto: ' + prj.name + ' | regione: ' + prj.region + ' | stato: ' + prj.status);

  // Registrazione: stessa forma delle voci «supabase» e «supabase-conti».
  const esegui = (args) => spawnSync('claude', args, { shell: true, encoding: 'utf8' });
  esegui(['mcp', 'remove', '-s', 'user', NOME_MCP]);
  const add = esegui(['mcp', 'add', '-s', 'user', NOME_MCP, '--', 'npx', '-y', '@supabase/mcp-server-supabase',
    '--project-ref=' + REF, '--access-token=' + token]);
  if (add.status !== 0) {
    console.log('ESITO: token ottenuto ma registrazione fallita: ' + maschera((add.stderr || add.stdout || '').slice(0, 300)));
    process.exit(6);
  }
  fs.unlinkSync(STATO);
  console.log('ESITO: OK — accesso registrato come «' + NOME_MCP + '» (' + maschera(String(add.stdout || '').trim().split('\n')[0]).slice(0, 120) + ')');
}

const [cmd, arg] = process.argv.slice(2);
(cmd === 'avvia' ? avvia() : cmd === 'completa' ? completa(arg) : Promise.reject(new Error('uso: avvia | completa <codice>')))
  .catch((e) => { console.log('ESITO: errore — ' + maschera(e.message)); process.exit(1); });
