// Chiamate all'API di GitHub per gli account dei due ambienti.
//
//   collaudo    la chiave si legge dall'archivio credenziali di Windows tramite
//               GCM (vedi github-accesso.js);
//   produzione  la chiave si cerca nello stesso archivio e, se non c'è,
//               nell'indirizzo del remoto «origin» di questa cartella. In
//               produzione ogni chiamata che CAMBIA qualcosa è rifiutata, a meno
//               che lo strumento che la fa non sia stato lanciato con
//               «--confermo-produzione».
//
// Account e repository vengono dal file del reparto (voce «github», vedi
// github() in reparti.js; dalla v199): l'utente di cui si usa la chiave, il
// repository del sorgente, quello del rinvio, il sito di servizio, il sito
// pubblico. Le chiamate possono toccare SOLO quei repository, uno per uno (non
// gli interi account), e crearne uno solo se è uno di quelli: un proprietario
// può essere anche un'organizzazione (un sorgente trasferito alla società), e lì
// un repository nuovo si crea con /orgs/<organizzazione>/repos. Da qualunque
// ambiente, una chiamata che cambia un repository della produzione vuole
// l'ambiente «produzione» e la conferma. Un rinvio di GitHub (repository
// rinominato o trasferito) non si segue: la chiave non va dove il controllo non
// ha guardato, e lo strumento vede la risposta 3xx.
//
// Nessuna chiave viene mai stampata. Uso da riga di comando (solo collaudo):
//   node gh-api.js <METODO> <percorso> ['<corpo json>']
const { spawnSync, execFileSync } = require('child_process');
const path = require('path');

const RADICE = path.resolve(__dirname, '../..');
// Il remoto di questa cartella per ciascun ambiente: è configurazione del PC, non del reparto.
const REMOTI = { collaudo: 'collaudo', produzione: 'origin' };
// La RISERVA: ciò che gli strumenti sanno anche se i file dei reparti non si
// leggono. Uno strumento d'emergenza («scambio-repository.js produzione annulla»)
// non deve fermarsi per un file scritto male. Deve dire le stesse cose dei
// file: lo pretende il controllo 14 di controlli-rilascio.js, quindi quando un
// file cambia account o repository si cambia anche qui.
// «sitoNuovo» è l'indirizzo del sito su Cloudflare: è lì che la pagina di rinvio
// di quell'ambiente deve portare, e non altrove.
// «squadra» e «dominio» sono le voci di pubblicazione.cloudflare (verifica-chiusura.js).
const RISERVA = {
  collaudo: { utente: 'gistech2026', remoto: 'collaudo', sitoNuovo: 'https://consegne-collaudo.pages.dev/',
    sorgente: 'gistech2026/app-consegne-sorgente', rinvio: 'gistech2026/app-consegne', sitoServizio: 'gistech2026/collaudo', sitoPubblico: 'gistech2026/collaudo-sito',
    squadra: 'consegne-collaudo', dominio: 'consegne-collaudo.torsadestech.com', proprietari: ['gistech2026'] },
  produzione: { utente: 'medicinadurgenzaucsc-maker', remoto: 'origin', sitoNuovo: 'https://consegne-reparto.pages.dev/',
    sorgente: 'medicinadurgenzaucsc-maker/app-consegne-sorgente', rinvio: 'medicinadurgenzaucsc-maker/app-consegne', sitoServizio: null, sitoPubblico: null,
    squadra: null, dominio: null, proprietari: ['medicinadurgenzaucsc-maker'] },
};
// Un ambiente come lo dice il file del suo reparto, oppure null se il file non si legge.
function dalFile(ambiente) {
  try {
    const reparti = require('./reparti.js');
    const r = reparti.risolvi(ambiente);
    const g = reparti.github(r), cf = r.pubblicazione.cloudflare, sp = r.pubblicazione.sitoPubblico;
    return { utente: g.utente, remoto: REMOTI[ambiente], sitoNuovo: cf ? 'https://' + cf.progetto + '.pages.dev/' : null,
      sorgente: g.sorgente.repository, rinvio: g.rinvio.repository, sitoServizio: g.sitoServizio ? g.sitoServizio.repository : null, sitoPubblico: sp ? sp.repository : null,
      squadra: (cf && cf.squadra) || null, dominio: (cf && cf.dominio) || null, proprietari: g.proprietari, reparto: r.codice, daFile: true };
  } catch (e) { return null; }
}
const AMBIENTI = {};
Object.keys(RISERVA).forEach((k) => { AMBIENTI[k] = dalFile(k) || Object.assign({}, RISERVA[k], { reparto: null, daFile: false }); });
// I repository in cui un ambiente può lavorare, in minuscolo: quelli del suo reparto, più il nome
// provvisorio della pagina di rinvio, che lo scambio prepara accanto al vecchio indirizzo.
function repositoryDi(a) {
  return [a.sorgente, a.rinvio, a.rinvio && a.rinvio + '-rinvio', a.sitoServizio, a.sitoPubblico].filter(Boolean)
    .map((x) => String(x).toLowerCase()).filter((x, i, l) => l.indexOf(x) === i);
}
// Una chiamata è ammessa? «p» è il percorso già risolto e minuscolo, «corpo» quello che parte.
// Ammessi soltanto: la scheda dell'utente; i repository dell'elenco (e i loro sottopercorsi); la
// scheda di un'organizzazione che possiede uno di quei repository, in lettura; la creazione di un
// repository, con /user/repos o /orgs/<organizzazione>/repos, solo col nome di uno dell'elenco.
// Un «%» nel percorso non passa: un separatore codificato non deve cambiare che cosa si tocca.
function percorsoAmmesso(regole, metodo, p, corpo) {
  const m = String(metodo || 'GET').toUpperCase();
  const io = String(regole.utente).toLowerCase();
  const ammessi = regole.repository.map((x) => String(x).toLowerCase());
  const proprietari = ammessi.map((x) => x.split('/')[0]);
  if (typeof p !== 'string' || p.indexOf('%') >= 0) return false;
  const nome = corpo && typeof corpo.name === 'string' ? corpo.name.toLowerCase() : null;
  if (p === '/user') return m === 'GET';
  if (p === '/user/repos') return m === 'POST' && !!nome && ammessi.indexOf(io + '/' + nome) >= 0;
  const o = /^\/orgs\/([^/]+)(\/repos)?$/.exec(p);
  if (o) {
    if (o[1] === io || proprietari.indexOf(o[1]) < 0) return false;
    return o[2] ? (m === 'POST' && !!nome && ammessi.indexOf(o[1] + '/' + nome) >= 0) : m === 'GET';
  }
  const r = /^\/repos\/([^/]+)\/([^/]+)(\/.*)?$/.exec(p);
  return !!r && ammessi.indexOf(r[1] + '/' + r[2]) >= 0;
}
// La chiamata cambia un repository della produzione? Vale da qualunque ambiente: i repository della
// produzione sono quelli del suo file e quelli della riserva, così un file scritto male non li toglie.
function toccaLaProduzione(metodo, p, corpo) {
  if (String(metodo || 'GET').toUpperCase() === 'GET') return false;
  const loro = repositoryDi(AMBIENTI.produzione).concat(repositoryDi(RISERVA.produzione));
  const r = /^\/repos\/([^/]+)\/([^/]+)/.exec(p);
  if (r && loro.indexOf(r[1] + '/' + r[2]) >= 0) return true;
  const nome = corpo && typeof corpo.name === 'string' ? corpo.name.toLowerCase() : null;
  const o = /^\/orgs\/([^/]+)\/repos$/.exec(p);
  return !!(nome && o && loro.indexOf(o[1] + '/' + nome) >= 0);
}
// Solo i due nomi scritti qui sopra: «constructor» e simili non sono ambienti.
const eAmbiente = (x) => Object.prototype.hasOwnProperty.call(AMBIENTI, String(x));
// Anche il codice di un reparto (reparti/<codice>.json) vale da nome d'ambiente:
// «gemelli-11n» è la produzione. Lo traduce reparti.js; le barriere restano
// legate al nome «produzione», che è quello che esce di qui.
function nomeAmbiente(x) {
  const n = String(x).toLowerCase();
  if (eAmbiente(n)) return n;
  const storico = require('./reparti.js').nomeStorico(n);
  return storico && eAmbiente(storico) ? storico : null;
}
const CONFERMA = '--confermo-produzione';
const maschera = (s) => String(s || '').replace(/(gh[pousr]_|github_pat_)[A-Za-z0-9_]+/g, 'gh*_***');

function daGcm(utente) {
  const r = spawnSync('git', ['credential-manager', 'get'], {
    input: 'protocol=https\nhost=github.com\nusername=' + utente + '\n\n',
    encoding: 'utf8',
    env: { ...process.env, GCM_INTERACTIVE: 'never', GIT_TERMINAL_PROMPT: '0' },
  });
  const m = /^password=(.+)$/m.exec(r.stdout || '');
  const u = /^username=(.+)$/m.exec(r.stdout || '');
  if (!m) return null;
  if (u && u[1].trim().toLowerCase() !== utente.toLowerCase()) return null;   // GCM ha risposto con un altro account
  return m[1].trim();
}
// La chiave scritta nell'indirizzo di un remoto (https://chiave@github.com/proprietario/…), solo se il remoto è di quel proprietario.
function dalRemoto(remoto, utente) {
  let u = '';
  try { u = execFileSync('git', ['-C', RADICE, 'remote', 'get-url', remoto], { encoding: 'utf8' }).trim(); } catch (e) { return null; }
  const m = /^https:\/\/([^@/]+)@github\.com\/([^/]+)\//.exec(u);
  if (!m || m[2].toLowerCase() !== utente.toLowerCase()) return null;
  const chiave = m[1].indexOf(':') >= 0 ? m[1].split(':').pop() : m[1];
  return /^(gh[pousr]_|github_pat_)/.test(chiave) ? chiave : null;
}

// Tutto ciò che serve per lavorare su un ambiente. «opzioni.ancheRepository» aggiunge all'elenco dei
// repository ammessi quelli che uno strumento dichiara per sé (es. rinvio-rubrica.js), nella forma
// proprietario/nome.
function per(ambiente, opzioni) {
  const scritto = ambiente;
  ambiente = nomeAmbiente(ambiente);
  if (!ambiente) throw new Error('ambiente sconosciuto: ' + scritto + ' (validi: ' + Object.keys(AMBIENTI).join(', ') + ', oppure il codice di un reparto)');
  const a = AMBIENTI[ambiente];
  const UTENTE = a.utente;
  const PROPRIETARI = a.proprietari.indexOf(UTENTE.toLowerCase()) >= 0 ? a.proprietari.slice() : [UTENTE.toLowerCase()].concat(a.proprietari);
  const REPOSITORY = repositoryDi(a).concat(((opzioni && opzioni.ancheRepository) || []).map((x) => String(x).toLowerCase()));
  const produzione = ambiente === 'produzione';
  // la chiave della produzione non si usa mai da un altro ambiente
  const loro = [AMBIENTI.produzione.utente, RISERVA.produzione.utente].map((x) => String(x).toLowerCase());
  if (!produzione && loro.indexOf(UTENTE.toLowerCase()) >= 0) throw new Error('l\'account di ' + ambiente + ' (' + UTENTE + ') è quello della produzione: il file del reparto va corretto');
  let memoria = null;
  function token() {
    if (memoria) return memoria;
    memoria = daGcm(UTENTE) || (produzione ? dalRemoto(a.remoto, UTENTE) : null);
    if (!memoria) throw new Error('chiave di GitHub per ' + UTENTE + ' non trovata' + (produzione ? ' né in GCM né nel remoto «' + a.remoto + '»' : ' in GCM'));
    return memoria;
  }
  async function gh(metodo, percorso, corpo) {
    // Solo i repository di questo ambiente. Il percorso si controlla nella forma in cui parte
    // davvero: «..» viene risolto prima di spedire, e un percorso senza la barra iniziale
    // si attaccherebbe al nome del sito, mandando la chiave altrove.
    if (typeof percorso !== 'string' || percorso.charAt(0) !== '/') throw new Error('percorso non valido: deve cominciare con «/»');
    const u = new URL('https://api.github.com' + percorso);
    if (u.origin !== 'https://api.github.com') throw new Error('percorso non valido');
    const p = u.pathname.toLowerCase();
    if (!percorsoAmmesso({ utente: UTENTE, repository: REPOSITORY }, metodo, p, corpo)) {
      throw new Error('chiamata fuori dai repository di ' + ambiente + ' (' + REPOSITORY.join(', ') + '): ' + String(metodo).toUpperCase() + ' ' + u.pathname);
    }
    if (!produzione && toccaLaProduzione(metodo, p, corpo)) {
      throw new Error('questa chiamata cambierebbe un repository della PRODUZIONE (' + metodo + ' ' + percorso + ') da ' + ambiente + ': non si fa');
    }
    if (produzione && String(metodo).toUpperCase() !== 'GET' && process.argv.indexOf(CONFERMA) < 0) {
      throw new Error('in PRODUZIONE questa chiamata cambierebbe qualcosa (' + metodo + ' ' + percorso + '): rilanciare lo strumento con ' + CONFERMA + ', e solo con l\'ok di chi gestisce l\'app');
    }
    const r = await fetch(u.href, {
      // un rinvio porterebbe la chiave a un indirizzo che il controllo qui sopra non ha visto
      redirect: 'manual',
      method: metodo,
      headers: {
        Authorization: 'Bearer ' + token(),
        Accept: 'application/vnd.github+json',
        'X-GitHub-Api-Version': '2022-11-28',
        'User-Agent': 'consegne-' + ambiente,
        ...(corpo ? { 'Content-Type': 'application/json' } : {}),
      },
      body: corpo ? JSON.stringify(corpo) : undefined,
    });
    const testo = await r.text();
    let dati = null;
    try { dati = testo ? JSON.parse(testo) : null; } catch (e) { dati = testo; }
    // una risposta 3xx resta tale: per uno strumento vuol dire «repository spostato»
    return { stato: r.status, dati, permessi: r.headers.get('x-oauth-scopes') || '', spostato: r.status >= 300 && r.status < 400 };
  }
  // Un repository nuovo: nell'account dell'utente con /user/repos, in un'organizzazione con /orgs/<nome>/repos.
  function crea(proprietario, corpo) {
    const mio = String(proprietario).toLowerCase() === UTENTE.toLowerCase();
    return gh('POST', mio ? '/user/repos' : '/orgs/' + proprietario + '/repos', corpo);
  }
  // true se il proprietario è un'organizzazione, false se è una persona; null se GitHub non lo dice.
  async function eOrganizzazione(proprietario) {
    if (String(proprietario).toLowerCase() === UTENTE.toLowerCase()) return false;
    const r = await gh('GET', '/orgs/' + proprietario);
    return r.stato === 200 ? true : r.stato === 404 ? false : null;
  }
  return { ambiente, UTENTE, proprietari: PROPRIETARI, repository: REPOSITORY, remoto: a.remoto, sitoNuovo: a.sitoNuovo, sorgente: a.sorgente, rinvio: a.rinvio,
    sitoServizio: a.sitoServizio, sitoPubblico: a.sitoPubblico, squadra: a.squadra, dominio: a.dominio, daFile: a.daFile, reparto: a.reparto, gh, crea, eOrganizzazione, token, produzione };
}

// Da un elenco di argomenti: l'ambiente (se è il primo) e il resto, senza la conferma.
// Un ambiente scritto fuori posto non viene ignorato: lo strumento si ferma,
// altrimenti «annulla produzione» lavorerebbe sul collaudo senza dirlo.
// «esplicito» dice se l'ambiente è stato scritto: i comandi che cambiano qualcosa
// lo pretendono (vedi pretendiAmbiente), così un comando copiato da un documento
// senza l'ambiente non lavora sul collaudo al posto della produzione.
function argomenti(argv) {
  const confermato = argv.indexOf(CONFERMA) >= 0;
  const resto = argv.filter((x) => x !== CONFERMA);
  const sconosciuta = resto.filter((x) => /^--/.test(String(x)))[0];
  if (sconosciuta) throw new Error('opzione sconosciuta: ' + sconosciuta + ' (l\'unica ammessa è ' + CONFERMA + ')');
  let ambiente = 'collaudo', esplicito = false;
  if (resto.length && nomeAmbiente(resto[0])) { ambiente = nomeAmbiente(resto.shift()); esplicito = true; }
  const fuoriPosto = resto.filter((x) => nomeAmbiente(x));
  if (fuoriPosto.length) throw new Error('il nome dell\'ambiente («' + fuoriPosto[0] + '») va scritto per primo, subito dopo il nome dello strumento, e una volta sola');
  if (confermato && ambiente !== 'produzione') throw new Error(CONFERMA + ' vale solo per la produzione: scrivere «produzione» come primo argomento');
  return { ambiente, resto, confermato, esplicito };
}
// Per i comandi che cambiano qualcosa: l'ambiente va scritto, non sottinteso.
function pretendiAmbiente(arg, comando) {
  if (!arg.esplicito) throw new Error('«' + comando + '» cambia qualcosa: l\'ambiente va scritto per esteso come primo argomento (collaudo oppure produzione)');
}

const collaudo = per('collaudo');
module.exports = { gh: collaudo.gh, token: collaudo.token, UTENTE: collaudo.UTENTE, per, argomenti, pretendiAmbiente, nomeAmbiente, maschera, percorsoAmmesso, toccaLaProduzione, repositoryDi, AMBIENTI, RISERVA, CONFERMA };

if (require.main === module) {
  const [metodo, percorso, corpo] = process.argv.slice(2);
  collaudo.gh(metodo || 'GET', percorso || '/user', corpo ? JSON.parse(corpo) : undefined)
    .then((x) => { console.log('HTTP ' + x.stato); console.log(maschera(JSON.stringify(x.dati, null, 1)).slice(0, 4000)); })
    .catch((e) => { console.log('ERRORE: ' + maschera(e.message)); process.exit(1); });
}
