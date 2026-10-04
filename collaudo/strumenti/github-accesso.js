// Accesso all'account GitHub di COLLAUDO con il flusso «a codice» (device flow):
// lo stesso di Git Credential Manager, il componente di Git già installato sul
// PC, e con il suo stesso identificativo pubblico (sorgente pubblico:
// git-ecosystem/git-credential-manager, src/GitHub/GitHubConstants.cs). Si pilota
// da qui perché, senza un terminale interattivo, GCM non riesce a mostrare il codice.
//
//   node github-accesso.js avvia    -> stampa il codice da digitare su github.com/login/device
//   node github-accesso.js attendi  -> aspetta l'autorizzazione e registra l'accesso
//
// Il token non viene MAI stampato: va nell'archivio credenziali di Windows
// tramite GCM, accanto agli altri account già presenti.
const fs = require('fs');
const path = require('path');
const { spawnSync } = require('child_process');

const UTENTE = 'gistech2026';
const CLIENT_ID = '0120e057bd645470c1ed';
const STATO = path.join(__dirname, '.accesso-github.json');
const maschera = (s) => String(s || '').replace(/gh[pousr]_[A-Za-z0-9_]+/g, 'gh*_***');
const dormi = (ms) => new Promise((r) => setTimeout(r, ms));
const form = (o) => ({
  method: 'POST',
  headers: { Accept: 'application/json', 'Content-Type': 'application/x-www-form-urlencoded' },
  body: new URLSearchParams(o),
});

async function avvia() {
  const r = await fetch('https://github.com/login/device/code', form({ client_id: CLIENT_ID, scope: 'repo gist workflow' }));
  const j = await r.json();
  if (!j.device_code || !j.user_code) { console.log('ESITO: GitHub non ha rilasciato il codice: ' + JSON.stringify(j).slice(0, 200)); process.exit(2); }
  fs.writeFileSync(STATO, JSON.stringify({ device_code: j.device_code, interval: j.interval || 5, scade: Date.now() + (j.expires_in || 900) * 1000 }));
  console.log('CODICE ' + j.user_code + ' | ' + j.verification_uri + ' | valido ' + Math.round((j.expires_in || 900) / 60) + ' minuti');
}

async function attendi() {
  if (!fs.existsSync(STATO)) { console.log('ESITO: nessun accesso avviato'); process.exit(2); }
  const st = JSON.parse(fs.readFileSync(STATO, 'utf8'));
  let pausa = (st.interval || 5) * 1000;
  let token = null;
  while (Date.now() < st.scade) {
    await dormi(pausa);
    let j;
    try {
      const r = await fetch('https://github.com/login/oauth/access_token', form({
        client_id: CLIENT_ID, device_code: st.device_code,
        grant_type: 'urn:ietf:params:oauth:grant-type:device_code',
      }));
      j = await r.json();
    } catch (e) { continue; }
    if (j.access_token) { token = j.access_token; break; }
    if (j.error === 'authorization_pending') continue;
    if (j.error === 'slow_down') { pausa += 5000; continue; }
    console.log('ESITO: autorizzazione non concessa (' + (j.error || 'errore sconosciuto') + ')');
    try { fs.unlinkSync(STATO); } catch (e) {}
    process.exit(3);
  }
  if (!token) { console.log('ESITO: codice scaduto senza autorizzazione'); try { fs.unlinkSync(STATO); } catch (e) {} process.exit(4); }

  const u = await fetch('https://api.github.com/user', { headers: { Authorization: 'Bearer ' + token, Accept: 'application/vnd.github+json', 'User-Agent': 'collaudo-consegne' } });
  const utente = await u.json();
  const permessi = u.headers.get('x-oauth-scopes') || '';
  if (String(utente.login || '').toLowerCase() !== UTENTE.toLowerCase()) {
    console.log('ESITO: autorizzato l\'account sbagliato («' + (utente.login || '?') + '» invece di «' + UTENTE + '»): accesso NON registrato');
    try { fs.unlinkSync(STATO); } catch (e) {}
    process.exit(5);
  }
  const s = spawnSync('git', ['credential-manager', 'store'], {
    input: 'protocol=https\nhost=github.com\nusername=' + UTENTE + '\npassword=' + token + '\n\n', encoding: 'utf8',
  });
  if (s.status !== 0) { console.log('ESITO: token ottenuto ma non registrato: ' + maschera((s.stderr || s.stdout || '').slice(0, 200))); process.exit(6); }
  try { fs.unlinkSync(STATO); } catch (e) {}
  console.log('ESITO: OK — account ' + utente.login + ' | permessi: ' + permessi);
}

const cmd = process.argv[2];
(cmd === 'avvia' ? avvia() : cmd === 'attendi' ? attendi() : Promise.reject(new Error('uso: avvia | attendi')))
  .catch((e) => { console.log('ESITO: errore — ' + maschera(e.message)); process.exit(1); });
