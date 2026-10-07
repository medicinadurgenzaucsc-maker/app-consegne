// I passi del flusso pubblica-cloudflare.yml che non sono un semplice comando.
//
//   node .github/pubblicazione/passi.js sito            prepara _sito/ e legge la versione
//   node .github/pubblicazione/passi.js configurazione  Cloudflare è configurato qui?
//   node .github/pubblicazione/passi.js progetto        il progetto Pages esiste (o lo crea)
//   node .github/pubblicazione/passi.js verifica        la pubblicazione è quella di questo commit
//   node .github/pubblicazione/passi.js avvisa          aggiorna app_version: sui PC compare «Update»
//
// Tutto arriva dall'ambiente del flusso. Nessuna chiave viene mai stampata: di
// Cloudflare e di Supabase si mostrano solo lo stato della risposta e i motivi
// di un errore.
const fs = require('fs');
const path = require('path');

const RADICE = path.resolve(__dirname, '../..');
const env = process.env;
const errore = (t) => { console.log('::error::' + t); process.exit(1); };
const esporta = (file, riga) => { if (file) fs.appendFileSync(file, riga + '\n'); };

async function cloudflare(metodo, percorso, corpo) {
  const r = await fetch('https://api.cloudflare.com/client/v4/accounts/' + env.CLOUDFLARE_ACCOUNT_ID.trim() + '/pages/projects' + percorso, {
    method: metodo,
    headers: Object.assign({ Authorization: 'Bearer ' + env.CLOUDFLARE_API_TOKEN.trim() }, corpo ? { 'Content-Type': 'application/json' } : {}),
    body: corpo ? JSON.stringify(corpo) : undefined,
  });
  let dati = null;
  try { dati = await r.json(); } catch (e) { dati = null; }
  return { stato: r.status, dati };
}
function motivi(r) {
  const elenco = ((r.dati && r.dati.errors) || []).map((e) => e.code + ' ' + e.message).join(' | ') || 'nessun dettaglio';
  const aiuto = (r.stato === 401 || r.stato === 403)
    ? ". La chiave non è valida, oppure non ha il permesso «Cloudflare Pages: Edit», oppure l'identificativo dell'account non è quello a cui la chiave appartiene"
    : '';
  return 'Cloudflare ha risposto HTTP ' + r.stato + ': ' + elenco + aiuto;
}

const PASSI = {
  // I file del sito in _sito/ e il numero di versione, che serve ai passi dopo.
  sito() {
    if (Number(process.versions.node.split('.')[0]) < 22) errore('serve Node 22 o successivo, trovato ' + process.version);
    const { prepara } = require(path.join(RADICE, 'collaudo/strumenti/prepara-sito.js'));
    const r = prepara(path.join(RADICE, '_sito'));
    r.elenco.forEach((p) => console.log('  ' + p));
    console.log(r.elenco.length + ' file, ' + r.byte + ' byte');
    ['index.html', 'print.html', '404.html', 'sw.js', 'manifest.json', 'js/api.js'].forEach((p) => {
      if (r.elenco.indexOf(p) < 0) errore('fra i file del sito manca ' + p);
    });
    const m = /CACHE_NAME = 'consegne-v(\d+)'/.exec(fs.readFileSync(path.join(RADICE, '_sito/sw.js'), 'utf8'));
    if (!m) errore('numero di versione non trovato in sw.js');
    esporta(env.GITHUB_ENV, 'VERSIONE=' + m[1]);
    console.log('versione del sito: ' + m[1]);
  },

  // Senza progetto o senza chiavi non si pubblica: il flusso si ferma qui, senza errore.
  configurazione() {
    const manca = [];
    if (!env.PROGETTO) manca.push('la variabile CLOUDFLARE_PROGETTO');
    if (!env.CLOUDFLARE_API_TOKEN) manca.push('il segreto CLOUDFLARE_API_TOKEN');
    if (!env.CLOUDFLARE_ACCOUNT_ID) manca.push('il segreto CLOUDFLARE_ACCOUNT_ID');
    // errori di copia: si dicono subito, senza mostrare il valore
    const chiave = String(env.CLOUDFLARE_API_TOKEN || '').trim(), account = String(env.CLOUDFLARE_ACCOUNT_ID || '').trim();
    if (env.PROGETTO && !/^[a-z0-9]([a-z0-9-]{0,56}[a-z0-9])?$/.test(env.PROGETTO)) errore('CLOUDFLARE_PROGETTO non è un nome di progetto valido (lettere minuscole, cifre e trattini)');
    if (account && !/^[0-9a-f]{32}$/.test(account)) errore("il segreto CLOUDFLARE_ACCOUNT_ID non ha la forma di un identificativo di account (32 caratteri fra 0-9 e a-f): è stato incollato qualcos'altro");
    if (chiave && (/\s/.test(chiave) || chiave.length < 30)) errore('il segreto CLOUDFLARE_API_TOKEN non ha la forma di una chiave (è troppo corto o contiene spazi): va incollata la chiave intera, senza altro testo');
    if (chiave && chiave === account) errore('i due segreti hanno lo stesso valore: uno dei due è stato incollato nel posto sbagliato');
    esporta(env.GITHUB_OUTPUT, 'pronto=' + (manca.length ? 'no' : 'si'));
    if (manca.length) console.log('::notice title=Nulla da pubblicare::In questo repository manca ' + manca.join(', ') + '. Controllati i file del sito e lo strumento di pubblicazione; nessuna pubblicazione.');
    else console.log('Cloudflare è configurato: progetto «' + env.PROGETTO + '»');
  },

  // Il progetto Pages: se non c'è lo crea. Dice a che indirizzo risponde, e si
  // ferma se l'app non riconosce quell'indirizzo (si aprirebbe solo l'avviso
  // «Indirizzo non riconosciuto»).
  async progetto() {
    let r = await cloudflare('GET', '/' + env.PROGETTO);
    if (r.stato === 404) {
      console.log('Il progetto «' + env.PROGETTO + '» non esiste ancora: lo creo.');
      r = await cloudflare('POST', '', { name: env.PROGETTO, production_branch: 'master' });
    }
    if (!r.dati || !r.dati.success) errore(motivi(r));
    const p = r.dati.result;
    console.log('progetto: ' + p.name + ' | ramo di produzione: ' + p.production_branch + ' | indirizzo: https://' + p.subdomain);
    if (p.production_branch !== 'master') errore('il ramo di produzione del progetto è «' + p.production_branch + '», atteso «master»');
    const api = fs.readFileSync(path.join(RADICE, 'docs/js/api.js'), 'utf8');
    if (api.indexOf("'" + p.subdomain + "'") < 0) errore("l'indirizzo " + p.subdomain + " non è fra quelli che l'app riconosce (elenchi «host» in docs/js/api.js): va aggiunto all'ambiente giusto prima di pubblicare");
    esporta(env.GITHUB_ENV, 'SITO=https://' + p.subdomain);
  },

  // L'ultima pubblicazione di produzione del progetto è proprio questo commit, ed è riuscita.
  async verifica() {
    const r = await cloudflare('GET', '/' + env.PROGETTO + '/deployments?env=production');
    if (!r.dati || !r.dati.success) errore(motivi(r));
    const d = (r.dati.result || [])[0];
    if (!d) errore('nessuna pubblicazione di produzione trovata nel progetto');
    const meta = (d.deployment_trigger && d.deployment_trigger.metadata) || {};
    const fase = d.latest_stage || {};
    console.log('ultima pubblicazione: ' + d.url + ' | ambiente: ' + d.environment + ' | commit: ' + meta.commit_hash + ' | fase: ' + fase.name + ' ' + fase.status);
    if (meta.commit_hash !== env.GITHUB_SHA) errore('la pubblicazione di produzione non è quella di questo commit');
    if (fase.status !== 'success') errore('la pubblicazione non risulta conclusa: ' + fase.name + ' ' + fase.status);
    const riga = (t) => esporta(env.GITHUB_STEP_SUMMARY, t);
    riga('### Pubblicato su Cloudflare');
    riga('- progetto: ' + env.PROGETTO);
    riga('- sito: ' + env.SITO);
    riga('- versione: ' + env.VERSIONE + ' (' + String(env.GITHUB_SHA).slice(0, 7) + ')');
    riga('- questa pubblicazione: ' + d.url);
  },

  // Come faceva notify-deploy.yml dopo la pubblicazione su GitHub Pages.
  async avvisa() {
    if (!env.SUPABASE_URL || !env.SUPABASE_SERVICE_ROLE_KEY) errore('CLOUDFLARE_AVVISA vale «si» ma mancano i segreti SUPABASE_URL e SUPABASE_SERVICE_ROLE_KEY');
    const sha = String(env.GITHUB_SHA);
    const base = env.SUPABASE_URL.trim();
    const r = await fetch((base.charAt(base.length - 1) === '/' ? base.slice(0, -1) : base) + '/rest/v1/app_version?id=eq.1', {
      method: 'PATCH',
      headers: { apikey: env.SUPABASE_SERVICE_ROLE_KEY.trim(), Authorization: 'Bearer ' + env.SUPABASE_SERVICE_ROLE_KEY.trim(), 'Content-Type': 'application/json', Prefer: 'return=representation' },
      body: JSON.stringify({ sha: sha, deployed_at: Date.now(), message: 'deploy ' + sha.slice(0, 7) }),
    });
    let righe = null;
    try { righe = await r.json(); } catch (e) { righe = null; }
    console.log('app_version: HTTP ' + r.status);
    if (!r.ok) errore('aggiornamento di app_version non riuscito (HTTP ' + r.status + ')');
    if (!Array.isArray(righe) || righe.length !== 1) errore('app_version: attesa una riga aggiornata, trovate ' + (Array.isArray(righe) ? righe.length : 'nessuna'));
    console.log('avviso inviato: sui PC aperti compare «Update» (' + sha.slice(0, 7) + ')');
  },
};

const passo = PASSI[process.argv[2]];
if (!passo) errore('passo sconosciuto: ' + process.argv[2] + ' (validi: ' + Object.keys(PASSI).join(', ') + ')');
Promise.resolve().then(passo).catch((e) => errore('errore imprevisto nel passo «' + process.argv[2] + '»: ' + String((e && e.message) || e).slice(0, 300)));
