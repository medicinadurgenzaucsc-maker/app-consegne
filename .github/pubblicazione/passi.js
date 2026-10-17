// I passi del flusso pubblica-cloudflare.yml che non sono un semplice comando.
//
//   node .github/pubblicazione/passi.js sito            prepara _sito/ (copia compressa) e legge la versione
//   node .github/pubblicazione/passi.js configurazione  Cloudflare è configurato qui?
//   node .github/pubblicazione/passi.js progetto        il progetto Pages esiste (o lo crea)
//   node .github/pubblicazione/passi.js verifica        la pubblicazione è quella di questo commit
//   node .github/pubblicazione/passi.js avvisa          aggiorna app_version: sui PC compare «Update»
//
// Per quale reparto si pubblica lo dice il progetto Pages scritto nelle
// impostazioni del repository (CLOUDFLARE_PROGETTO): è il reparto che in
// reparti/<codice>.json dichiara quel progetto. Il sito si prepara per lui
// (js/ambiente.js generato dal suo file), e un progetto che nessun reparto
// dichiara ferma il flusso. Dove Cloudflare non è configurato si controlla
// soltanto, sul sito di docs/ così com'è.
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
const reparti = require(path.join(RADICE, 'collaudo/strumenti/reparti.js'));

// Il reparto di questa pubblicazione, oppure null se nel repository non è scritto nessun progetto.
function repartoDelFlusso() {
  if (!env.PROGETTO) return null;
  let r = null;
  try { r = reparti.perProgettoCloudflare(env.PROGETTO); } catch (e) { errore('i file in reparti/ non sono validi: ' + String((e && e.message) || e).split('\n').join(' | ').slice(0, 900)); }
  if (!r) errore('nessun reparto in reparti/ pubblica sul progetto «' + env.PROGETTO + '» (campo pubblicazione.cloudflare.progetto): non si pubblica un sito senza sapere di chi è');
  return r;
}

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
  // I file del sito in _sito/, nella forma compressa che viene pubblicata, e il
  // numero di versione, che serve ai passi dopo. La copia si controlla da sé:
  // se un controllo fallisce il flusso si ferma qui e non pubblica nulla.
  sito() {
    if (Number(process.versions.node.split('.')[0]) < 22) errore('serve Node 22 o successivo, trovato ' + process.version);
    const { prepara } = require(path.join(RADICE, 'collaudo/strumenti/prepara-sito.js'));
    const reparto = repartoDelFlusso();
    let r = null;
    try { r = prepara(path.join(RADICE, '_sito'), { compressa: true, reparto: reparto ? reparto.codice : undefined }); } catch (e) { errore(String((e && e.message) || e).split('\n').join(' | ').slice(0, 900)); }
    if (reparto && r.reparto !== reparto.codice) errore('il sito è stato preparato per «' + r.reparto + '» invece che per «' + reparto.codice + '»');
    console.log('reparto: ' + r.reparto + (reparto ? '' : ' (nessun progetto Cloudflare in questo repository: è il sito di docs/, solo per controllo)'));
    esporta(env.GITHUB_ENV, 'REPARTO=' + r.reparto);
    const kb = (n) => (n / 1024).toFixed(0).padStart(5) + ' KB';
    r.compressione.righe.forEach((x) => console.log('  ' + kb(x.prima) + ' → ' + kb(x.dopo) + '  ' + x.file.padEnd(34) + x.nota));
    console.log(r.elenco.length + ' file: da ' + kb(r.compressione.prima).trim() + ' a ' + kb(r.compressione.dopo).trim() + ' (' + r.compressione.strumento + '); ' + r.compressione.nomiGlobali + ' nomi globali rimasti identici');
    console.log('impronte dei file pubblicati (sha256):');
    r.compressione.righe.forEach((x) => console.log('  ' + x.impronta + '  ' + x.file));
    ['index.html', 'print.html', '404.html', 'sw.js', 'manifest.json', 'manifest-rubrica.json', 'js/ambiente.js', 'js/api.js', 'rubrica/index.html', 'rubrica/app.js', 'rubrica/app.css'].forEach((p) => {
      if (r.elenco.indexOf(p) < 0) errore('fra i file del sito manca ' + p);
    });
    // dal sorgente: nella copia compressa la riga ha un'altra forma
    const m = /CACHE_NAME = 'consegne-v(\d+)'/.exec(fs.readFileSync(path.join(RADICE, 'docs/sw.js'), 'utf8'));
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

  // Il progetto Pages: se non c'è lo crea. Dice a che indirizzi risponde (il suo
  // .pages.dev e i domini collegati), e si ferma se l'app non ne riconosce uno
  // (lì si aprirebbe solo l'avviso «Indirizzo non riconosciuto»). Il dominio
  // dichiarato nel file del reparto e non ancora collegato al progetto è solo
  // un avviso: il nome si collega da Cloudflare, dopo.
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
    const reparto = repartoDelFlusso();
    const nomi = [p.subdomain].concat(Array.isArray(p.domains) ? p.domains : []).map((x) => String(x || '').toLowerCase()).filter((x, i, a) => x && a.indexOf(x) === i);
    nomi.forEach((h) => { if (!reparti.hostAmmesso(reparto, h)) errore("l'indirizzo " + h + ' non è fra quelli del reparto «' + reparto.codice + '» (campo «host» di reparti/' + reparto.codice + '.json): il sito pubblicato lì si fermerebbe su «Indirizzo non riconosciuto»'); });
    if (nomi.length > 1) console.log('domini collegati al progetto: ' + nomi.slice(1).join(', '));
    const dominio = reparto.pubblicazione.cloudflare && reparto.pubblicazione.cloudflare.dominio;
    if (dominio && nomi.indexOf(dominio) < 0) console.log('::notice title=Dominio non ancora collegato::Il file del reparto dichiara il dominio ' + dominio + ', ma il progetto Pages non lo ha ancora: il sito risponde solo da ' + p.subdomain + '. Il nome si collega da Cloudflare insieme alla sua applicazione Access, senza minuti in mezzo (finché l\'applicazione non c\'è, da quel nome il sito è aperto a tutti), e subito dopo si prova da estraneo con verifica-chiusura.js.');
    esporta(env.GITHUB_ENV, 'SITO=https://' + p.subdomain);
    if (dominio && nomi.indexOf(dominio) >= 0) esporta(env.GITHUB_ENV, 'DOMINIO=https://' + dominio);
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
    riga('- reparto: ' + repartoDelFlusso().codice);
    riga('- progetto: ' + env.PROGETTO);
    riga('- sito: ' + env.SITO);
    if (env.DOMINIO) riga('- dominio: ' + env.DOMINIO);
    riga('- versione: ' + env.VERSIONE + ' (' + String(env.GITHUB_SHA).slice(0, 7) + ')');
    riga('- questa pubblicazione: ' + d.url);
  },

  // Come faceva notify-deploy.yml dopo la pubblicazione su GitHub Pages.
  async avvisa() {
    if (!env.SUPABASE_URL || !env.SUPABASE_SERVICE_ROLE_KEY) errore('CLOUDFLARE_AVVISA vale «si» ma mancano i segreti SUPABASE_URL e SUPABASE_SERVICE_ROLE_KEY');
    const sha = String(env.GITHUB_SHA);
    const base = env.SUPABASE_URL.trim();
    // «Update» va acceso sui PC del reparto che si è appena pubblicato, non su quelli di un altro:
    // il database dei segreti deve essere quello scritto nel file del reparto. Il valore non si stampa.
    const reparto = repartoDelFlusso();
    if (!reparto) errore('avviso ai PC: in questo repository non è scritto il progetto Cloudflare, quindi non si sa di quale reparto siano');
    if ((base.charAt(base.length - 1) === '/' ? base.slice(0, -1) : base) !== reparto.supabaseUrl) errore('il segreto SUPABASE_URL non è il database del reparto «' + reparto.codice + '»: l\'avviso arriverebbe ai PC di un altro');
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
