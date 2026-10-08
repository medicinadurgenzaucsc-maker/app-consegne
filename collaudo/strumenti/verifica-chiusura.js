// Prova DA ESTRANEO, senza alcuna credenziale, che dopo lo scambio dei
// repository il sorgente non si scarichi più da nessuna strada, che al vecchio
// indirizzo resti solo la pagina di rinvio e che il sito nuovo chieda l'accesso.
// Non scrive nulla da nessuna parte: fa solo richieste pubbliche.
//
//   node collaudo/strumenti/verifica-chiusura.js collaudo
//   node collaudo/strumenti/verifica-chiusura.js produzione [https://<progetto>.pages.dev/]
//   …  --archivi     interroga anche Software Heritage e la Wayback Machine (più lento)
//
// Esce con 0 solo se tutte le prove passano. Su un ambiente dove lo scambio non
// è stato fatto DEVE fallire: è la controprova che le prove vedono davvero un
// repository aperto. Un «limite di richieste» non vale mai come prova superata.
// Nei primi minuti dopo lo scambio i file «raw» e lo zip possono rispondere
// ancora: GitHub impiega qualche secondo a propagare il cambio e una risposta
// presa in quel momento resta 5 minuti nella cache del distributore. Per questo
// la verifica si lancia un paio di minuti dopo lo scambio e si ripete dopo dieci.
const { spawnSync } = require('child_process');

const AMBIENTI = {
  collaudo: { proprietario: 'gistech2026', nuovo: 'https://consegne-collaudo.pages.dev/' },
  produzione: { proprietario: 'medicinadurgenzaucsc-maker', nuovo: 'https://consegne-reparto.pages.dev/' },
};
const APP = 'app-consegne', SORGENTE = 'app-consegne-sorgente';
const FILE_RINVIO = ['.nojekyll', '404.html', 'README.md', 'index.html', 'print.html', 'robots.txt', 'sw.js'];
const UA = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/141.0 Safari/537.36';
const attendi = (ms) => new Promise((r) => setTimeout(r, ms));

const argomenti = process.argv.slice(2);
const nomeAmbiente = argomenti[0];
const amb = AMBIENTI[nomeAmbiente];
if (!amb) { console.log('uso: verifica-chiusura.js collaudo|produzione [indirizzo del sito nuovo] [--archivi]'); process.exit(2); }
const O = amb.proprietario;
// l'indirizzo del sito nuovo, sempre con la barra finale
const NUOVO = ((x) => (x && x.charAt(x.length - 1) !== '/' ? x + '/' : x))(argomenti.filter((a) => /^https:\/\//.test(a))[0] || amb.nuovo);
const ARCHIVI = argomenti.includes('--archivi');

const esiti = [];
function nota(passa, cosa, visto) {
  esiti.push(passa);
  console.log((passa === true ? 'ok      ' : passa === null ? 'NON SO  ' : 'APERTO  ') + cosa + '  [' + visto + ']');
}

async function chiedi(url, opzioni) {
  try {
    const r = await fetch(url, { redirect: 'manual', cache: 'no-store', headers: { 'User-Agent': UA, Accept: '*/*' }, ...(opzioni || {}) });
    const limite = (r.status === 403 || r.status === 429) && (r.headers.get('x-ratelimit-remaining') === '0' || r.status === 429);
    return { stato: r.status, limite, verso: r.headers.get('location') || '', risposta: r };
  } catch (e) { return { stato: 0, limite: false, verso: '', errore: String(e.message || e).slice(0, 80) }; }
}
// Deve rispondere 404: è il modo in cui GitHub dice «non esiste» a chi non può vederlo.
async function deveMancare(cosa, url) {
  let r = await chiedi(url);
  for (let i = 0; i < 4 && r.stato >= 300 && r.stato < 400 && r.verso; i++) r = await chiedi(new URL(r.verso, url).href);
  if (r.limite || r.stato === 0) return nota(null, cosa, r.limite ? 'limite di richieste' : 'errore di rete');
  const h = r.risposta ? r.risposta.headers : null;
  const cache = h && r.stato === 200 && /HIT/i.test(h.get('x-cache') || '') ? ', copia in cache da ' + (h.get('source-age') || h.get('age') || '?') + ' s: sparisce entro 5 minuti' : '';
  nota(r.stato === 404, cosa, 'HTTP ' + r.stato + cache);
}
// Nomi dei file dentro uno zip, letti dall'indice in fondo all'archivio.
function nomiDelloZip(b) {
  const i = b.lastIndexOf(Buffer.from([0x50, 0x4b, 0x05, 0x06]));
  if (i < 0) return null;
  const n = b.readUInt16LE(i + 10); let p = b.readUInt32LE(i + 16); const nomi = [];
  for (let k = 0; k < n; k++) { const ln = b.readUInt16LE(p + 28), le = b.readUInt16LE(p + 30), lc = b.readUInt16LE(p + 32); nomi.push(b.slice(p + 46, p + 46 + ln).toString()); p += 46 + ln + le + lc; }
  return nomi.filter((x) => !x.endsWith('/')).map((x) => x.split('/').slice(1).join('/'));
}
const soloRinvio = (nomi) => nomi.length > 0 && nomi.every((x) => FILE_RINVIO.includes(x));
const codiceApp = (t) => /navVersioneApp|_AMBIENTI|supabaseAnonKey/.test(t);
// una richiesta che non ha avuto risposta: rete assente, nome non risolto, limite di richieste
const senzaRisposta = (r) => !!(r.limite || r.stato === 0);

(async () => {
  console.log('Ambiente: ' + nomeAmbiente + ' | proprietario su GitHub: ' + O + ' | sito nuovo: ' + (NUOVO || 'non indicato'));

  console.log('\n— Il repository del sorgente (' + O + '/' + SORGENTE + ') non deve esistere per un estraneo');
  await deveMancare('scheda del repository (API)', 'https://api.github.com/repos/' + O + '/' + SORGENTE);
  await deveMancare('pagina su github.com', 'https://github.com/' + O + '/' + SORGENTE);
  await deveMancare('un file (raw)', 'https://raw.githubusercontent.com/' + O + '/' + SORGENTE + '/master/docs/js/api.js');
  await deveMancare('un file (raw), saltando la cache', 'https://raw.githubusercontent.com/' + O + '/' + SORGENTE + '/master/docs/js/app.js?cb=' + Date.now());
  await deveMancare('zip del ramo', 'https://github.com/' + O + '/' + SORGENTE + '/archive/refs/heads/master.zip');
  await deveMancare('contenuto di un file (API)', 'https://api.github.com/repos/' + O + '/' + SORGENTE + '/contents/docs/js/api.js');
  await deveMancare('storia dei commit (API)', 'https://api.github.com/repos/' + O + '/' + SORGENTE + '/commits');
  {
    const g = spawnSync('git', ['-c', 'credential.helper=', '-c', 'core.askPass=', 'ls-remote', 'https://github.com/' + O + '/' + SORGENTE + '.git'],
      { encoding: 'utf8', env: { ...process.env, GIT_TERMINAL_PROMPT: '0', GCM_INTERACTIVE: 'never', GIT_ASKPASS: '', SSH_ASKPASS: '' }, timeout: 30000 });
    // «rifiutato» vale solo se GitHub ha risposto chiedendo le credenziali o dicendo che non esiste:
    // una scadenza o la rete assente non provano nulla
    const detto = String(g.stderr || '');
    const rifiuto = g.status !== 0 && g.status !== null && /could not read Username|Authentication failed|Repository not found|terminal prompts disabled|not found/i.test(detto);
    nota(g.status === 0 ? false : rifiuto ? true : null, 'git senza credenziali', g.status === 0 ? 'ha elencato i rami' : rifiuto ? 'rifiutato' : 'git non ha raggiunto GitHub: prova non eseguita');
  }

  console.log('\n— Al vecchio nome (' + O + '/' + APP + ') deve restare solo la pagina di rinvio');
  {
    const r = await chiedi('https://api.github.com/repos/' + O + '/' + APP + '/contents/docs/sw.js');
    if (r.limite || r.stato === 0) nota(null, 'il sorgente non sta più al vecchio nome', 'limite di richieste o rete');
    else nota(r.stato === 404, 'il sorgente non sta più al vecchio nome', 'docs/sw.js: HTTP ' + r.stato);
    const c = await chiedi('https://api.github.com/repos/' + O + '/' + APP + '/commits?per_page=5');
    if (c.limite || c.stato !== 200) nota(null, 'un solo commit, senza storia', c.limite ? 'limite di richieste' : 'HTTP ' + c.stato);
    else { const j = await c.risposta.json(); nota(Array.isArray(j) && j.length === 1, 'un solo commit, senza storia', (Array.isArray(j) ? j.length : '?') + ' commit'); }
  }
  await deveMancare('file del sorgente chiesto per ramo (raw)', 'https://raw.githubusercontent.com/' + O + '/' + APP + '/master/docs/js/api.js');
  for (const ramo of ['master', 'main']) {
    let r = await chiedi('https://github.com/' + O + '/' + APP + '/archive/refs/heads/' + ramo + '.zip');
    for (let i = 0; i < 4 && r.stato >= 300 && r.stato < 400 && r.verso; i++) r = await chiedi(r.verso);
    if (r.limite || r.stato === 0) { nota(null, 'zip chiesto come «' + ramo + '»', 'limite di richieste o rete'); continue; }
    if (r.stato === 404) { nota(true, 'zip chiesto come «' + ramo + '»', 'HTTP 404'); continue; }
    const nomi = r.stato === 200 ? nomiDelloZip(Buffer.from(await r.risposta.arrayBuffer())) : null;
    nota(!!nomi && soloRinvio(nomi), 'zip chiesto come «' + ramo + '»: solo i file del rinvio', nomi ? nomi.length + ' file' + (soloRinvio(nomi) ? '' : ', fra cui ' + nomi.filter((x) => !FILE_RINVIO.includes(x)).slice(0, 3).join(', ')) : 'HTTP ' + r.stato);
  }

  console.log('\n— Il vecchio sito su GitHub non deve più servire l\'applicazione');
  const sito = 'https://' + O.toLowerCase() + '.github.io/';
  {
    const r = await chiedi(sito + APP + '/?cb=' + Date.now());
    const t = r.risposta ? await r.risposta.text() : '';
    // senza risposta non si può dire né «chiuso» né «aperto»
    const muto = senzaRisposta(r);
    nota(muto ? null : r.stato === 200 && /ha cambiato indirizzo/.test(t) && !codiceApp(t), 'il vecchio indirizzo mostra la pagina di rinvio', muto ? 'limite di richieste o rete' : 'HTTP ' + r.stato + (codiceApp(t) ? ', serve l\'applicazione' : ''));
    if (NUOVO) nota(muto ? null : t.indexOf(NUOVO) >= 0, 'la pagina di rinvio porta al sito nuovo', muto ? 'limite di richieste o rete' : t.indexOf(NUOVO) >= 0 ? NUOVO : 'indirizzo non trovato nella pagina');
  }
  for (const f of ['js/api.js', 'js/app.js', 'css/styles.css']) {
    for (const coda of ['?cb=' + Date.now(), '']) {
      const r = await chiedi(sito + APP + '/' + f + coda);
      const t = r.risposta ? await r.risposta.text() : '';
      if (senzaRisposta(r)) { nota(null, 'vecchio sito, ' + f + (coda ? '' : ' dalla cache'), 'limite di richieste o rete'); continue; }
      nota(r.stato === 404 && !codiceApp(t), 'vecchio sito, ' + f + (coda ? '' : ' dalla cache'), 'HTTP ' + r.stato + ', ' + t.length + ' byte');
    }
  }
  await deveMancare('sito del repository del sorgente', sito + SORGENTE + '/');
  await deveMancare('sito del sorgente, un file', sito + SORGENTE + '/js/api.js');

  if (!NUOVO) {
    console.log('\n— Il sito nuovo');
    nota(null, 'sito nuovo dietro l\'accesso', 'indirizzo del sito nuovo non indicato: prove non eseguite');
  }
  if (NUOVO) {
    console.log('\n— Il sito nuovo deve chiedere l\'accesso prima di dare qualunque file');
    {
      // anche gli indirizzi di anteprima che Cloudflare crea a ogni pubblicazione
      const u = new URL(NUOVO);
      const r = await chiedi(u.protocol + '//verifica-anteprima.' + u.host + '/js/api.js');
      const cancello = r.stato === 302 && /cloudflareaccess\.com\/cdn-cgi\/access\/login/.test(r.verso);
      nota(senzaRisposta(r) ? null : cancello, 'un indirizzo di anteprima del sito nuovo', senzaRisposta(r) ? 'limite di richieste o rete' : 'HTTP ' + r.stato + (cancello ? ' verso l\'accesso' : ''));
    }
    for (const f of ['', 'js/api.js', 'sw.js', 'print', 'manifest.json']) {
      const r = await chiedi(NUOVO + f);
      const cancello = r.stato === 302 && /cloudflareaccess\.com\/cdn-cgi\/access\/login/.test(r.verso);
      nota(senzaRisposta(r) ? null : cancello, 'sito nuovo, /' + f, senzaRisposta(r) ? 'limite di richieste o rete' : 'HTTP ' + r.stato + (cancello ? ' verso l\'accesso' : ''));
    }
  }

  if (ARCHIVI) {
    console.log('\n— Archivi pubblici di terzi (una copia lì non si può richiamare)');
    for (const nome of [APP, SORGENTE]) {
      const r = await fetch('https://archive.softwareheritage.org/api/1/origin/https://github.com/' + O + '/' + nome + '/get/').then((x) => x.status, () => 0);
      nota(r === 404 ? true : r === 200 ? false : null, 'Software Heritage, ' + O + '/' + nome, r === 404 ? 'non archiviato' : r === 200 ? 'ARCHIVIATO' : 'HTTP ' + r);
    }
    const cdx = async (prefisso) => {
      for (let i = 0; i < 3; i++) {
        try {
          const r = await fetch('https://web.archive.org/cdx/search/cdx?url=' + encodeURIComponent(prefisso) + '&output=json&limit=5&fl=timestamp,original');
          if (r.status !== 200) { await attendi(15000); continue; }
          const t = await r.text(); const j = JSON.parse(t || '[]');
          return Math.max(0, j.length - 1);
        } catch (e) { await attendi(5000); }
      }
      return null;
    };
    const controprova = await cdx('github.com/torvalds/linux*');
    await attendi(4000);
    for (const prefisso of ['github.com/' + O + '/' + APP + '*', O.toLowerCase() + '.github.io/' + APP + '*']) {
      const n = controprova ? await cdx(prefisso) : null;
      nota(n === 0 ? true : n > 0 ? false : null, 'Wayback Machine, ' + prefisso, n == null ? 'archivio non interrogabile ora' : n === 0 ? 'nessuna copia' : n + ' copie');
      await attendi(4000);
    }
  }

  const aperti = esiti.filter((x) => x === false).length, dubbi = esiti.filter((x) => x === null).length;
  console.log('\n' + (aperti === 0 && dubbi === 0 ? 'TUTTO CHIUSO: ' + esiti.length + ' prove superate'
    : (aperti ? 'ATTENZIONE: ' + aperti + ' prove dicono che qualcosa è ancora aperto' : '') + (aperti && dubbi ? '; ' : '') + (dubbi ? dubbi + ' prove non eseguite fino in fondo' : '')));
  process.exit(aperti === 0 && dubbi === 0 ? 0 : 1);
})().catch((e) => { console.log('ERRORE: ' + String(e && e.message || e).slice(0, 300)); process.exit(1); });
