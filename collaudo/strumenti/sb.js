// Accesso ai due database dalle procedure di collaudo.
//  - collaudo():   lettura e scrittura, SOLO sul progetto di collaudo
//  - produzione(): SOLA LETTURA sul progetto vero. La barriera è del server:
//                  con read_only la query gira come utente di sola lettura
//                  (supabase_read_only_user) in una transazione di sola
//                  lettura. Provato sul collaudo l'08/10/2026: sono respinte
//                  scritture dirette, seconde istruzioni, «commit» seguito da
//                  una scrittura, SELECT INTO, funzioni che scrivono. Il
//                  controllo sul testo della query viene prima e ferma gli
//                  errori di distrazione.
// I token si leggono dalla configurazione locale e non vengono mai stampati.
const fs = require('fs');
const os = require('os');
const path = require('path');

const REF_COLLAUDO = 'rqvohwpthhumydpbwktq';
const REF_PRODUZIONE = 'ifmmcvxzhwdkmzhsxcvb';
const maschera = (s) => String(s || '').replace(/sbp_[A-Za-z0-9_]+/g, 'sbp_***').replace(/eyJ[A-Za-z0-9_-]{10,}\.[A-Za-z0-9_-]{10,}\.[A-Za-z0-9_-]{10,}/g, 'eyJ***');

function token(voce, refAtteso) {
  // Il messaggio di un errore di lettura riporterebbe un pezzo del file, che contiene le chiavi.
  let j = null;
  try { j = JSON.parse(fs.readFileSync(path.join(os.homedir(), '.claude.json'), 'utf8')); }
  catch (e) { throw new Error('la configurazione locale (~/.claude.json) non si legge: è assente o non è scritta bene'); }
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

// Il testo di una query senza stringhe, nomi fra virgolette e commenti (al loro
// posto resta '' oppure ""). Legge da sinistra a destra come fa PostgreSQL:
//  - stringhe fra apici, con l'apice doppiato; in quelle precedute da E la
//    barra rovesciata protegge il carattere seguente;
//  - nomi fra virgolette doppie, che possono contenere apici, punti e virgola,
//    dollari e segni di commento;
//  - stringhe fra dollari ($$…$$, $nome$…$nome$), ma non quando il dollaro è
//    la coda di un nome (a$x$ è un identificatore);
//  - commenti di riga, chiusi anche da un ritorno a capo senza \n, e commenti
//    di blocco, che si possono annidare;
//  - una stringa continuata a capo ('a' e sotto 'b') viene rifiutata.
// Ciò che non si chiude fa fallire la lettura: meglio rifiutare che indovinare.
function senzaLetterali(sql) {
  const diNome = (c) => c !== undefined && (/[A-Za-z0-9_$]/.test(c) || c.charCodeAt(0) >= 128);
  let fuori = '', i = 0;
  while (i < sql.length) {
    const c = sql[i], d = sql[i + 1], prima = sql[i - 1];
    if (c === "'") {
      const conBarre = (prima === 'E' || prima === 'e') && !diNome(sql[i - 2]);
      let j = i + 1, chiusa = false;
      while (j < sql.length) {
        if (conBarre && sql[j] === '\\') { j += 2; continue; }
        if (sql[j] === "'") { if (sql[j + 1] === "'") { j += 2; continue; } chiusa = true; break; }
        j++;
      }
      if (!chiusa) throw new Error('query con una stringa non chiusa');
      // PostgreSQL unisce due stringhe separate solo da spazi con un a capo, e la seconda
      // resta dello stesso tipo della prima: nessuna lettura degli strumenti lo usa, si rifiuta.
      // (Espressione senza ripetizioni annidate: su un rientro lungo il tempo deve restare lineare.)
      if (/^[ \t\f]*(?:--[^\n\r]*)?[\n\r](?:[ \t\n\r\f\v]|--[^\n\r]*[\n\r])*'/.test(sql.slice(j + 1))) throw new Error('query con una stringa continuata a capo');
      fuori += "''"; i = j + 1; continue;
    }
    if (c === '"') {
      let j = i + 1, chiusa = false;
      while (j < sql.length) {
        if (sql[j] === '"') { if (sql[j + 1] === '"') { j += 2; continue; } chiusa = true; break; }
        j++;
      }
      if (!chiusa) throw new Error('query con un nome fra virgolette non chiuso');
      fuori += '""'; i = j + 1; continue;
    }
    if (c === '$' && !diNome(prima)) {
      const m = /^\$((?:[A-Za-z_]|[^\x00-\x7F])(?:[A-Za-z0-9_]|[^\x00-\x7F])*)?\$/.exec(sql.slice(i));
      if (m) {
        const fine = sql.indexOf(m[0], i + m[0].length);
        if (fine < 0) throw new Error('query con una stringa fra dollari non chiusa');
        fuori += "''"; i = fine + m[0].length; continue;
      }
    }
    if (c === '-' && d === '-') {
      let j = i + 2;
      while (j < sql.length && sql[j] !== '\n' && sql[j] !== '\r') j++;
      fuori += ' '; i = j; continue;
    }
    if (c === '/' && d === '*') {
      let j = i + 2, livello = 1;
      while (j < sql.length && livello > 0) {
        if (sql[j] === '/' && sql[j + 1] === '*') { livello++; j += 2; continue; }
        if (sql[j] === '*' && sql[j + 1] === '/') { livello--; j += 2; continue; }
        j++;
      }
      if (livello > 0) throw new Error('query con un commento non chiuso');
      fuori += ' '; i = j; continue;
    }
    fuori += c; i++;
  }
  return fuori;
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
    // Il controllo si fa sul testo SENZA i letterali e i commenti, dove ';' e parole chiave sono innocui.
    // Li toglie un lettore che va da sinistra a destra: stringhe fra apici, stringhe fra dollari
    // ($$…$$ e $nome$…$nome$), commenti di riga e di blocco. La barriera vera resta read_only sul server.
    const nudo = senzaLetterali(sql);
    if (!/^\s*(select|with)\b/i.test(nudo) || /;\s*\S/.test(nudo)
      || /\b(insert|update|delete|drop|alter|create|grant|revoke|truncate|commit|rollback|begin|call|copy|vacuum|refresh|comment|do|into|share)\b/i.test(nudo)) {
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

module.exports = { collaudo, produzione, maschera, senzaLetterali, REF_COLLAUDO, REF_PRODUZIONE };
