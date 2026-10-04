// Accesso ai due database dalle procedure di collaudo.
//  - collaudo():   lettura e scrittura, SOLO sul progetto di collaudo
//  - produzione(): SOLA LETTURA sul progetto vero (transazione read-only
//                  imposta dal server + controllo sul testo della query)
// I token si leggono dalla configurazione locale e non vengono mai stampati.
const fs = require('fs');
const os = require('os');
const path = require('path');

const REF_COLLAUDO = 'rqvohwpthhumydpbwktq';
const REF_PRODUZIONE = 'ifmmcvxzhwdkmzhsxcvb';
const maschera = (s) => String(s || '').replace(/sbp_[A-Za-z0-9_]+/g, 'sbp_***').replace(/eyJ[A-Za-z0-9_-]{10,}\.[A-Za-z0-9_-]{10,}\.[A-Za-z0-9_-]{10,}/g, 'eyJ***');

function token(voce, refAtteso) {
  const j = JSON.parse(fs.readFileSync(path.join(os.homedir(), '.claude.json'), 'utf8'));
  const s = (j.mcpServers || {})[voce];
  if (!s) throw new Error('voce «' + voce + '» assente dalla configurazione');
  const a = s.args || [];
  if (!a.includes('--project-ref=' + refAtteso)) throw new Error('la voce «' + voce + '» non punta al progetto ' + refAtteso);
  const t = a.find((x) => x.startsWith('--access-token='));
  if (!t) throw new Error('token assente nella voce «' + voce + '»');
  return t.slice('--access-token='.length);
}

async function chiama(ref, tok, metodo, percorso, corpo) {
  const r = await fetch('https://api.supabase.com/v1/projects/' + ref + percorso, {
    method: metodo,
    headers: { Authorization: 'Bearer ' + tok, ...(corpo ? { 'Content-Type': 'application/json' } : {}) },
    body: corpo ? JSON.stringify(corpo) : undefined,
  });
  const testo = await r.text();
  let dati = null;
  try { dati = testo ? JSON.parse(testo) : null; } catch (e) { dati = testo; }
  return { stato: r.status, dati };
}

function collaudo() {
  const tok = token('supabase-collaudo', REF_COLLAUDO);
  const api = (metodo, percorso, corpo) => chiama(REF_COLLAUDO, tok, metodo, percorso, corpo);
  const query = async (sql) => {
    const r = await api('POST', '/database/query', { query: sql });
    if (r.stato >= 300) throw new Error('SQL collaudo (HTTP ' + r.stato + '): ' + maschera(JSON.stringify(r.dati)).slice(0, 900));
    return r.dati;
  };
  // Richieste che non sono JSON (per esempio il caricamento di una funzione).
  const invia = (percorso, init) => fetch('https://api.supabase.com/v1/projects/' + REF_COLLAUDO + percorso,
    { ...init, headers: { ...((init && init.headers) || {}), Authorization: 'Bearer ' + tok } });
  return { ref: REF_COLLAUDO, api, query, invia };
}

function produzione() {
  const tok = token('supabase', REF_PRODUZIONE);
  const leggi = async (sql) => {
    // Il controllo si fa sul testo SENZA i letterali, dove ';' e parole chiave sono innocui.
    const nudo = sql.replace(/'(?:[^']|'')*'/g, "''").replace(/--[^\n]*/g, ' ');
    if (!/^\s*(select|with)\b/i.test(nudo) || /;\s*\S/.test(nudo)
      || /\b(insert|update|delete|drop|alter|create|grant|revoke|truncate|commit|rollback|begin|call|copy|vacuum|refresh|comment|do)\b/i.test(nudo)) {
      throw new Error('in produzione sono ammesse solo SELECT singole');
    }
    const r = await chiama(REF_PRODUZIONE, tok, 'POST', '/database/query', { query: sql, read_only: true });
    if (r.stato >= 300) throw new Error('SQL produzione (HTTP ' + r.stato + '): ' + maschera(JSON.stringify(r.dati)).slice(0, 900));
    return r.dati;
  };
  // Letture di configurazione (mai scritture): solo GET.
  const leggiApi = (percorso) => chiama(REF_PRODUZIONE, tok, 'GET', percorso);
  return { ref: REF_PRODUZIONE, leggi, leggiApi };
}

module.exports = { collaudo, produzione, maschera, REF_COLLAUDO, REF_PRODUZIONE };
