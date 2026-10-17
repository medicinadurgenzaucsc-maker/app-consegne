// I reparti: un file per ciascuno in reparti/<codice>.json, e da quel file il
// piccolo file servito docs/js/ambiente.js che dice all'app chi è.
//
//   node collaudo/strumenti/reparti.js                    elenca i reparti e li convalida
//   node collaudo/strumenti/reparti.js mostra <codice>    stampa il file js/ambiente.js di quel reparto
//   node collaudo/strumenti/reparti.js controlla          docs/js/ambiente.js è quello della 11N?
//   node collaudo/strumenti/reparti.js scrivi             riscrive docs/js/ambiente.js dal file della 11N
//
// docs/ è il sito della 11N servito così com'è da GitHub Pages: per questo nel
// repository docs/js/ambiente.js è SEMPRE quello del reparto che pubblica da
// docs/ (campo «pubblicazione.githubPages»). Per ogni altro reparto il file si
// genera al momento di preparare il sito (prepara-sito.js --reparto <codice>),
// e il banco di prova lo serve generato in memoria per il collaudo.
//
// Nei file dei reparti non entra MAI un segreto: indirizzi, chiave pubblica del
// database e identificativo del client Google sono pubblici per natura (stanno
// nel sito che chiunque sia ammesso riceve). La convalida rifiuta una chiave che
// non sia quella pubblica e ogni campo che non conosce.
const fs = require('fs');
const path = require('path');

const RADICE = path.resolve(__dirname, '../..');
const CARTELLA = path.join(RADICE, 'reparti');
const AMBIENTE_JS = path.join(RADICE, 'docs', 'js', 'ambiente.js');

const CAMPI = ['codice', 'sinonimi', 'nome', 'ospedale', 'tipo', 'ambiente', 'hosting', 'accesso', 'lingua', 'canale',
  'versioneLocale', 'livello', 'host', 'pubblicazione', 'github', 'supabaseUrl', 'supabaseAnonKey', 'googleClientId', 'centrale'];
const TIPI = ['produzione', 'collaudo'];
const HOSTING = ['reparto', 'fornitore', 'condiviso', 'locale'];
const ACCESSI = ['google'];
const CANALI = ['pilota', 'stabile'];
const LINGUE = ['it'];
// Un nome host scritto per esteso, oppure (solo nel collaudo) col punto davanti: vale per i sottodomini.
const HOST = /^\.?[a-z0-9]([a-z0-9-]*[a-z0-9])?(\.[a-z0-9]([a-z0-9-]*[a-z0-9])?)*$/;
const CODICE = /^[a-z0-9]([a-z0-9-]{0,40}[a-z0-9])?$/;
const LOCALI = ['localhost', '127.0.0.1'];
// Un account (persona od organizzazione) e un repository di GitHub.
const UTENTE_GITHUB = /^[A-Za-z0-9](?:[A-Za-z0-9]|-(?=[A-Za-z0-9])){0,38}$/;
const REPOSITORY = /^[A-Za-z0-9-]+\/[A-Za-z0-9._-]+$/;

// Il ruolo scritto dentro una chiave di Supabase (la parte centrale è JSON in base64).
function ruoloDellaChiave(chiave) {
  const pezzi = String(chiave || '').split('.');
  if (pezzi.length !== 3) return null;
  try { return String(JSON.parse(Buffer.from(pezzi[1].replace(/-/g, '+').replace(/_/g, '/'), 'base64').toString('utf8')).role || ''); }
  catch (e) { return null; }
}

// Errori di un reparto preso da solo. «nomeFile» è il nome senza estensione: deve essere il codice.
function difetti(r, nomeFile) {
  const e = [];
  if (!r || typeof r !== 'object' || Array.isArray(r)) return ['il file non contiene un oggetto'];
  Object.keys(r).forEach((k) => { if (CAMPI.indexOf(k) < 0) e.push('campo sconosciuto «' + k + '» (un file di reparto non porta mai segreti né campi liberi)'); });
  const testo = (k) => typeof r[k] === 'string' && r[k].trim() === r[k] && r[k] !== '';
  if (!testo('codice') || !CODICE.test(r.codice)) e.push('codice mancante o non valido (minuscole, cifre e trattini)');
  else if (nomeFile !== undefined && r.codice !== nomeFile) e.push('il codice «' + r.codice + '» non coincide con il nome del file «' + nomeFile + '.json»');
  if (r.sinonimi !== undefined && !(Array.isArray(r.sinonimi) && r.sinonimi.every((s) => typeof s === 'string' && CODICE.test(s)))) e.push('sinonimi: atteso un elenco di nomi');
  // «collaudo» e «produzione» sono i nomi che gli strumenti usano da sempre, e a «produzione» legano le loro barriere:
  // un reparto li può portare solo se è di quel tipo
  [r.codice].concat(Array.isArray(r.sinonimi) ? r.sinonimi : []).forEach((n) => { if (TIPI.indexOf(n) >= 0 && n !== r.tipo) e.push('il nome «' + n + '» è riservato a un reparto di tipo ' + n); });
  ['nome', 'ospedale'].forEach((k) => { if (!testo(k) || r[k].length > 80) e.push(k + ' mancante o troppo lungo'); });
  if (TIPI.indexOf(r.tipo) < 0) e.push('tipo: atteso uno fra ' + TIPI.join(', '));
  if (TIPI.indexOf(r.ambiente) < 0) e.push('ambiente: atteso uno fra ' + TIPI.join(', ') + ' (è il valore della riga AMBIENTE del suo database)');
  if (TIPI.indexOf(r.tipo) >= 0 && r.ambiente !== r.tipo) e.push('ambiente «' + r.ambiente + '» diverso dal tipo «' + r.tipo + '»');
  if (HOSTING.indexOf(r.hosting) < 0) e.push('hosting: atteso uno fra ' + HOSTING.join(', '));
  if (ACCESSI.indexOf(r.accesso) < 0) e.push('accesso: atteso uno fra ' + ACCESSI.join(', '));
  if (LINGUE.indexOf(r.lingua) < 0) e.push('lingua: atteso uno fra ' + LINGUE.join(', '));
  if (CANALI.indexOf(r.canale) < 0) e.push('canale: atteso uno fra ' + CANALI.join(', '));
  if (!Number.isInteger(r.versioneLocale) || r.versioneLocale < 0) e.push('versioneLocale: atteso un intero da 0 in su');
  else if (r.versioneLocale !== 0) e.push('versioneLocale diversa da 0: la versione locale arriva con la fase 5 del piano, per ora deve valere 0');
  if ([0, 1, 2].indexOf(r.livello) < 0) e.push('livello: atteso 0, 1 o 2');
  if (!Array.isArray(r.host) || !r.host.length) e.push('host: atteso un elenco non vuoto di nomi');
  else {
    r.host.forEach((h) => {
      if (typeof h !== 'string' || h !== h.toLowerCase() || !HOST.test(h) || h.length > 120) e.push('host non valido: «' + h + '» (solo il nome, minuscolo, senza https:// e senza percorso)');
      else if (r.tipo === 'produzione' && (h.charAt(0) === '.' || LOCALI.indexOf(h) >= 0)) e.push('host «' + h + '» non ammesso in un reparto di produzione: sui pazienti veri si lavora solo da indirizzi scritti per esteso, mai da anteprime o da questo PC');
    });
    if (new Set(r.host).size !== r.host.length) e.push('host ripetuti');
  }
  const p = r.pubblicazione;
  if (!p || typeof p !== 'object' || Array.isArray(p)) e.push('pubblicazione: atteso un oggetto');
  else {
    Object.keys(p).forEach((k) => { if (['githubPages', 'cloudflare', 'sitoPubblico'].indexOf(k) < 0) e.push('pubblicazione: voce sconosciuta «' + k + '»'); });
    if (!p.githubPages && !p.cloudflare && !p.sitoPubblico) e.push('pubblicazione: serve almeno una voce fra githubPages, cloudflare e sitoPubblico');
    // Cloudflare Pages: il progetto (risponde da <progetto>.pages.dev), e due voci facoltative:
    //  - «dominio», il nome sul dominio della società collegato allo stesso progetto (dalla v199): accanto al
    //    .pages.dev, mai al suo posto, perché il progetto risponde da tutti e due;
    //  - «squadra», la squadra di Cloudflare Access che chiude il sito: il cancello sta su
    //    https://<squadra>.cloudflareaccess.com, ed è lì che un estraneo deve essere rimandato.
    const cf = p.cloudflare;
    if (cf !== undefined) {
      if (!(cf && typeof cf === 'object' && !Array.isArray(cf) && typeof cf.progetto === 'string' && /^[a-z0-9]([a-z0-9-]{0,56}[a-z0-9])?$/.test(cf.progetto))) e.push('pubblicazione.cloudflare: atteso il campo «progetto», con un nome di progetto valido');
      else {
        Object.keys(cf).forEach((k) => { if (['progetto', 'dominio', 'squadra'].indexOf(k) < 0) e.push('pubblicazione.cloudflare: voce sconosciuta «' + k + '» (ammesse: progetto, dominio, squadra)'); });
        if (Array.isArray(r.host) && r.host.indexOf(cf.progetto + '.pages.dev') < 0) e.push('pubblicazione.cloudflare: «' + cf.progetto + '.pages.dev» non è fra gli host del reparto');
        if (cf.dominio !== undefined) {
          if (typeof cf.dominio !== 'string' || cf.dominio !== cf.dominio.toLowerCase() || cf.dominio.charAt(0) === '.' || !HOST.test(cf.dominio) || cf.dominio.indexOf('.') < 0 || cf.dominio.length > 120) e.push('pubblicazione.cloudflare.dominio: atteso un nome scritto per esteso, minuscolo, senza https:// e senza percorso');
          else if (/[.](pages[.]dev|github[.]io|supabase[.]co|cloudflareaccess[.]com)$/.test(cf.dominio) || LOCALI.indexOf(cf.dominio) >= 0) e.push('pubblicazione.cloudflare.dominio: «' + cf.dominio + '» non è un dominio proprio (il .pages.dev del progetto è già detto da «progetto»)');
          else if (Array.isArray(r.host) && r.host.indexOf(cf.dominio) < 0) e.push('pubblicazione.cloudflare.dominio: «' + cf.dominio + '» non è fra gli host del reparto: il sito pubblicato lì si fermerebbe su «Indirizzo non riconosciuto»');
        }
        if (cf.squadra !== undefined && !(typeof cf.squadra === 'string' && /^[a-z0-9]([a-z0-9-]{0,61}[a-z0-9])?$/.test(cf.squadra))) e.push('pubblicazione.cloudflare.squadra: atteso il nome della squadra di Cloudflare Access (minuscole, cifre e trattini: il cancello è https://<squadra>.cloudflareaccess.com)');
      }
    }
    if (p.githubPages && !(typeof p.githubPages.repository === 'string' && /^[A-Za-z0-9-]+\/[A-Za-z0-9._-]+$/.test(p.githubPages.repository) && p.githubPages.cartella === 'docs' && Object.keys(p.githubPages).length === 2)) e.push('pubblicazione.githubPages: attesi «repository» (proprietario/nome) e «cartella» = docs');
    if (p.githubPages && Array.isArray(r.host) && typeof p.githubPages.repository === 'string' && r.host.indexOf(p.githubPages.repository.split('/')[0].toLowerCase() + '.github.io') < 0) e.push('pubblicazione.githubPages: il sito del proprietario del repository non è fra gli host del reparto');
    // il «livello 0»: un repository PUBBLICO che contiene solo il sito già compresso, servito da GitHub Pages
    const sp = p.sitoPubblico;
    if (sp !== undefined) {
      if (!(sp && typeof sp === 'object' && !Array.isArray(sp) && typeof sp.repository === 'string' && /^[A-Za-z0-9-]+\/[A-Za-z0-9._-]+$/.test(sp.repository) && Object.keys(sp).length === 1)) e.push('pubblicazione.sitoPubblico: atteso il solo campo «repository» (proprietario/nome)');
      else {
        if (Array.isArray(r.host) && r.host.indexOf(sp.repository.split('/')[0].toLowerCase() + '.github.io') < 0) e.push('pubblicazione.sitoPubblico: il sito del proprietario del repository non è fra gli host del reparto');
        if (p.githubPages && typeof p.githubPages.repository === 'string' && p.githubPages.repository.toLowerCase() === sp.repository.toLowerCase()) e.push('pubblicazione.sitoPubblico: è lo stesso repository di githubPages, che contiene il sorgente; il sito pubblico vuole un repository suo');
      }
    }
  }
  // facoltativo (dalla v199): gli account e i repository di GitHub che gli strumenti di questo PC usano per il
  // reparto. Senza, valgono quelli del repository da cui il reparto pubblica docs/ (vedi github() qui sotto)
  if (r.github !== undefined) {
    const g = r.github;
    if (!(g && typeof g === 'object' && !Array.isArray(g))) e.push('github: atteso un oggetto');
    else {
      Object.keys(g).forEach((k) => { if (['utente', 'sorgente', 'rinvio', 'sitoServizio'].indexOf(k) < 0) e.push('github: voce sconosciuta «' + k + '» (ammesse: utente, sorgente, rinvio, sitoServizio)'); });
      if (!(typeof g.utente === 'string' && UTENTE_GITHUB.test(g.utente))) e.push('github.utente: atteso il nome di un account di GitHub');
      ['sorgente', 'rinvio'].forEach((k) => { if (!(typeof g[k] === 'string' && REPOSITORY.test(g[k]))) e.push('github.' + k + ': atteso un repository, nella forma proprietario/nome'); });
      if (g.sitoServizio !== undefined && !(typeof g.sitoServizio === 'string' && REPOSITORY.test(g.sitoServizio))) e.push('github.sitoServizio: atteso un repository, nella forma proprietario/nome');
      const minuscolo = (x) => String(x || '').toLowerCase();
      // il sorgente è privato: nessun repository pubblico del reparto può essere lui
      const pubblici = [['rinvio', g.rinvio], ['sitoServizio', g.sitoServizio], ['pubblicazione.sitoPubblico', p && p.sitoPubblico && p.sitoPubblico.repository]].filter((x) => x[1]);
      pubblici.forEach((x) => { if (typeof g.sorgente === 'string' && minuscolo(x[1]) === minuscolo(g.sorgente)) e.push('github.sorgente: è lo stesso repository di ' + x[0] + ', che è pubblico'); });
      if (g.sitoServizio && p && p.sitoPubblico && minuscolo(g.sitoServizio) === minuscolo(p.sitoPubblico.repository)) e.push('github.sitoServizio: è lo stesso repository del sito pubblico');
      if (g.sitoServizio && minuscolo(g.sitoServizio) === minuscolo(g.rinvio)) e.push('github.sitoServizio: è lo stesso repository del rinvio');
      // dove il reparto pubblica docs/ con GitHub Pages, il vecchio indirizzo è quel repository
      if (p && p.githubPages && typeof p.githubPages.repository === 'string' && typeof g.rinvio === 'string' && minuscolo(g.rinvio) !== minuscolo(p.githubPages.repository)) e.push('github.rinvio: deve essere il repository di pubblicazione.githubPages');
    }
  } else if (p && typeof p === 'object' && !p.githubPages) e.push('github: un reparto che non pubblica docs/ con GitHub Pages deve dire account e repository di GitHub (voce «github»)');
  if (!testo('supabaseUrl') || !/^https:\/\/[a-z]{20}\.supabase\.co$/.test(r.supabaseUrl)) e.push('supabaseUrl: atteso https://<progetto>.supabase.co, senza barra finale');
  if (!testo('supabaseAnonKey')) e.push('supabaseAnonKey mancante');
  else if (/^sb_secret_/.test(r.supabaseAnonKey)) e.push('supabaseAnonKey è una chiave SEGRETA: in un file di reparto va solo la chiave pubblica');
  else if (!/^sb_publishable_[A-Za-z0-9_-]+$/.test(r.supabaseAnonKey)) {
    const ruolo = ruoloDellaChiave(r.supabaseAnonKey);
    if (ruolo === null) e.push('supabaseAnonKey non ha la forma di una chiave di Supabase');
    else if (ruolo !== 'anon') e.push('supabaseAnonKey ha il ruolo «' + ruolo + '»: in un file di reparto va SOLO la chiave pubblica (ruolo anon)');
    else if (testo('supabaseUrl')) {
      // la chiave deve essere di QUEL progetto: una chiave pubblica giusta sul progetto sbagliato collegherebbe il sito al database di un altro
      let ref = null;
      try { ref = JSON.parse(Buffer.from(r.supabaseAnonKey.split('.')[1].replace(/-/g, '+').replace(/_/g, '/'), 'base64').toString('utf8')).ref; } catch (err) { ref = null; }
      if (ref && r.supabaseUrl.indexOf('://' + ref + '.') < 0) e.push('supabaseAnonKey è la chiave di un altro progetto, non di quello di supabaseUrl');
    }
  }
  if (!testo('googleClientId') || !/^\d+-[a-z0-9]+\.apps\.googleusercontent\.com$/.test(r.googleClientId)) e.push('googleClientId: atteso un identificativo di client Google (…apps.googleusercontent.com), non un segreto');
  // facoltativo: la centrale che riceve il battito dei dispositivi. Il solo indirizzo, mai una chiave
  if (r.centrale !== undefined) {
    const c = r.centrale;
    if (!(c && typeof c === 'object' && !Array.isArray(c) && Object.keys(c).length === 1 && typeof c.url === 'string' && /^https:\/\/[a-z]{20}\.supabase\.co$/.test(c.url))) e.push('centrale: atteso il solo campo «url», https://<progetto>.supabase.co senza barra finale');
    else if (c.url === r.supabaseUrl) e.push('centrale: è lo stesso database del reparto');
  }
  return e;
}

// Un nome host è coperto da una voce dell'elenco? Una voce col punto davanti vale per i sottodomini.
function hostCoperto(voce, host) {
  host = String(host || '').toLowerCase();
  return voce.charAt(0) === '.' ? (host.length > voce.length && host.slice(-voce.length) === voce) : host === voce;
}
function hostAmmesso(reparto, host) { return reparto.host.some((v) => hostCoperto(v, host)); }

function codici() {
  if (!fs.existsSync(CARTELLA)) return [];
  return fs.readdirSync(CARTELLA).filter((f) => /\.json$/.test(f)).map((f) => f.replace(/\.json$/, '')).sort();
}
function leggi(codice) {
  if (!CODICE.test(String(codice))) throw new Error('codice di reparto non valido: ' + codice);
  const f = path.join(CARTELLA, codice + '.json');
  if (!fs.existsSync(f)) throw new Error('reparto sconosciuto: «' + codice + '» (reparti noti: ' + (codici().join(', ') || 'nessuno') + ')');
  let r = null;
  try { r = JSON.parse(fs.readFileSync(f, 'utf8')); } catch (e) { throw new Error('reparti/' + codice + '.json non è JSON valido: ' + e.message); }
  return r;
}

// Tutti i reparti, convalidati uno per uno e fra loro. Lancia un errore con l'elenco completo dei difetti.
function tutti() { return convalida(codici().map((c) => leggi(c)), codici()); }
// La convalida di un elenco di reparti già letto («nomiFile» sono i nomi dei file senza estensione): la usa
// tutti(), e i controlli di rilascio la provano su copie dei file con un difetto ciascuna.
function convalida(elenco, nomiFile) {
  const e = [];
  elenco.forEach((r, i) => difetti(r, nomiFile[i]).forEach((d) => e.push('reparti/' + nomiFile[i] + '.json: ' + d)));
  if (e.length) throw new Error(e.join('\n'));
  // fra reparti: nessun nome in comune, nessun indirizzo in comune, nessun database in comune
  const nomi = {}, basi = {}, progetti = {};
  elenco.forEach((r) => {
    [r.codice].concat(r.sinonimi || []).forEach((n) => { if (nomi[n]) e.push('il nome «' + n + '» indica due reparti: ' + nomi[n] + ' e ' + r.codice); nomi[n] = r.codice; });
    if (basi[r.supabaseUrl]) e.push('i reparti ' + basi[r.supabaseUrl] + ' e ' + r.codice + ' hanno lo stesso database');
    basi[r.supabaseUrl] = r.codice;
    const pr = r.pubblicazione.cloudflare && r.pubblicazione.cloudflare.progetto;
    if (pr) { if (progetti[pr]) e.push('i reparti ' + progetti[pr] + ' e ' + r.codice + ' pubblicano sullo stesso progetto Cloudflare'); progetti[pr] = r.codice; }
  });
  // la centrale è una per tutti, ma non è mai il database di un reparto
  elenco.forEach((r) => { if (r.centrale && basi[r.centrale.url]) e.push('la centrale di ' + r.codice + ' è il database del reparto ' + basi[r.centrale.url]); });
  // un repository serve un sito solo. Fra quelli DICHIARATI nei file dei reparti: né due siti pubblici nello stesso,
  // né un sito pubblico nel repository da cui un reparto pubblica docs/. Un repository col sorgente che nessun file
  // dichiara (quello del collaudo) qui non si conosce: lo ferma sito-pubblico.js, che prima di scrivere guarda che
  // cosa il repository contiene
  // Dalla v199 contano anche i repository che gli strumenti usano (sorgente, rinvio col suo nome provvisorio,
  // sito di servizio), così come li ricava github(): anche quelli di un reparto che la voce «github» non ce l'ha,
  // come la 11N. Il rinvio di un reparto che pubblica docs/ con GitHub Pages è quello stesso repository, e non
  // vale come doppione. È la barriera che tiene ogni strumento dentro i repository del suo reparto: gh-api.js
  // lascia lavorare solo lì.
  const depositi = {};
  elenco.forEach((r) => {
    const g = github(r);
    [['githubPages', r.pubblicazione.githubPages && r.pubblicazione.githubPages.repository], ['sitoPubblico', r.pubblicazione.sitoPubblico && r.pubblicazione.sitoPubblico.repository],
      ['github.sorgente', g.sorgente.repository], ['github.rinvio', g.rinvio.repository], ['github.rinvio provvisorio', g.rinvio.repository + '-rinvio'],
      ['github.sitoServizio', g.sitoServizio && g.sitoServizio.repository]].forEach(([voce, valore]) => {
      const dep = valore && String(valore).toLowerCase();
      if (!dep) return;
      const gia = depositi[dep];
      const stessoSito = gia && gia.codice === r.codice && [gia.voce, voce].sort().join() === 'github.rinvio,githubPages';
      if (gia && !stessoSito && !(gia.codice === r.codice && gia.voce === voce)) e.push('il repository ' + dep + ' è dichiarato due volte: ' + gia.codice + ' ' + gia.voce + ' e ' + r.codice + ' ' + voce);
      if (!gia) depositi[dep] = { codice: r.codice, voce };
    });
  });
  // l'account con cui lavorano gli strumenti di un reparto di collaudo non è mai quello di un reparto di
  // produzione: la sua chiave aprirebbe la produzione senza la conferma
  elenco.forEach((a) => elenco.forEach((b) => {
    if (a.codice >= b.codice || a.tipo === b.tipo) return;
    if (github(a).utente.toLowerCase() === github(b).utente.toLowerCase()) e.push('i reparti ' + a.codice + ' (' + a.tipo + ') e ' + b.codice + ' (' + b.tipo + ') usano lo stesso account di GitHub: un collaudo non lavora mai con l\'account della produzione');
  }));
  // e nessun repository di un collaudo sta nell'account con cui lavora un reparto di produzione: lì ogni
  // repository è della produzione, anche uno che il suo file non nomina
  elenco.filter((r) => r.tipo !== 'produzione').forEach((r) => elenco.filter((x) => x.tipo === 'produzione').forEach((x) => {
    const suo = github(x).utente.toLowerCase(), g = github(r);
    [g.sorgente, g.rinvio, g.sitoServizio, r.pubblicazione.sitoPubblico && { repository: r.pubblicazione.sitoPubblico.repository }].filter(Boolean).forEach((d) => {
      if (String(d.repository).split('/')[0].toLowerCase() === suo) e.push('il reparto ' + r.codice + ' (' + r.tipo + ') nomina ' + d.repository + ', che sta nell\'account della produzione ' + x.codice + ': un collaudo non lavora mai lì');
    });
  }));
  elenco.forEach((a) => elenco.forEach((b) => {
    if (a.codice >= b.codice) return;
    a.host.forEach((ha) => b.host.forEach((hb) => {
      const nudoA = ha.replace(/^\./, ''), nudoB = hb.replace(/^\./, '');
      if (ha === hb || hostCoperto(ha, nudoB) || hostCoperto(hb, nudoA)) e.push('l\'indirizzo «' + ha + '» del reparto ' + a.codice + ' e «' + hb + '» del reparto ' + b.codice + ' si sovrappongono');
      // «cardio-consegne.pages.dev» accanto a «consegne.pages.dev»: prepara-sito.js cerca nel sito di un reparto
      // gli indirizzi degli altri come testo, e troverebbe il secondo dentro il primo
      else if (nudoA.indexOf(nudoB) >= 0 || nudoB.indexOf(nudoA) >= 0) e.push('l\'indirizzo «' + ha + '» del reparto ' + a.codice + ' e «' + hb + '» del reparto ' + b.codice + ' sono scritti uno dentro l\'altro: il controllo «nessun dato di altri reparti nel sito» li confonderebbe, va scelto un altro nome');
    }));
  }));
  if (e.length) throw new Error(e.join('\n'));
  return elenco;
}

// Da un codice o da un sinonimo («produzione») al reparto. Mai un valore predefinito: un nome sconosciuto è un errore.
function risolvi(nome) {
  const n = String(nome || '').toLowerCase();
  const trovati = tutti().filter((r) => r.codice === n || (r.sinonimi || []).indexOf(n) >= 0);
  if (trovati.length !== 1) throw new Error('reparto sconosciuto: «' + nome + '» (noti: ' + tutti().map((r) => [r.codice].concat(r.sinonimi || []).join(' = ')).join('; ') + ')');
  return trovati[0];
}
// Gli strumenti nati quando gli ambienti erano due li chiamano «collaudo» e
// «produzione», e alla parola «produzione» legano le loro barriere (sola
// lettura, conferma esplicita). Da un nome scritto a mano (il codice di un
// reparto o un suo sinonimo) al nome storico, oppure null: così «gemelli-11n»
// vale «produzione» dovunque e le barriere restano dove sono. I due nomi
// storici valgono anche se i file dei reparti non si leggono: uno strumento
// d'emergenza non deve fermarsi per un file scritto male.
function nomeStorico(nome) {
  const n = String(nome === undefined || nome === null ? '' : nome).toLowerCase();
  if (TIPI.indexOf(n) >= 0) return n;
  if (!CODICE.test(n)) return null;
  let trovati = [];
  try { trovati = tutti().filter((r) => r.codice === n || (r.sinonimi || []).indexOf(n) >= 0); } catch (e) { return null; }
  if (trovati.length !== 1) return null;
  const storici = [trovati[0].codice].concat(trovati[0].sinonimi || []).filter((x) => TIPI.indexOf(x) >= 0);
  return (storici.length === 1 && storici[0] === trovati[0].tipo) ? storici[0] : null;
}
function perProgettoCloudflare(progetto) {
  const trovati = tutti().filter((r) => r.pubblicazione.cloudflare && r.pubblicazione.cloudflare.progetto === progetto);
  return trovati.length === 1 ? trovati[0] : null;
}
// Il reparto il cui sito è la cartella docs/ del repository così com'è (oggi la 11N).
function delSitoInDocs() {
  const trovati = tutti().filter((r) => r.pubblicazione.githubPages);
  if (trovati.length !== 1) throw new Error('atteso un solo reparto che pubblica da docs/ con GitHub Pages, trovati ' + trovati.length);
  return trovati[0];
}

// Gli account e i repository di GitHub di un reparto, come li usano gli strumenti di questo PC (dalla v199,
// prima erano costanti scritte in ciascuno strumento):
//   utente        l'account con cui lavorano: la sua chiave sta nell'archivio credenziali (GCM)
//   sorgente      il repository col sorgente, privato dopo lo scambio dei repository
//   rinvio        il repository pubblico del vecchio indirizzo: prima dello scambio è il sorgente stesso,
//                 dopo porta solo la pagina di rinvio
//   sitoServizio  (facoltativo) il sito di servizio del collaudo: finto TrakCare e informativa
//   proprietari   tutti gli account (persone od organizzazioni) che possiedono un repository del reparto:
//                 sono i soli in cui gh-api.js lascia lavorare
// Il proprietario di un repository può essere un'organizzazione diversa dall'utente: è il caso di un sorgente
// trasferito all'organizzazione della società, mentre il vecchio indirizzo resta all'account di prima (i siti
// di GitHub Pages non seguono un trasferimento). Senza la voce «github» nel file valgono i nomi di sempre,
// presi dal repository da cui il reparto pubblica docs/: è il caso della 11N, il cui file non cambia.
function github(reparto) {
  const p = reparto.pubblicazione || {}, g = reparto.github;
  const pezzi = (x) => { const [proprietario, nome] = String(x).split('/'); return { proprietario, nome, repository: proprietario + '/' + nome }; };
  let utente, sorgente, rinvio, sitoServizio = null;
  if (g) {
    utente = g.utente; sorgente = pezzi(g.sorgente); rinvio = pezzi(g.rinvio);
    if (g.sitoServizio) sitoServizio = pezzi(g.sitoServizio);
  } else if (p.githubPages) {
    rinvio = pezzi(p.githubPages.repository);
    utente = rinvio.proprietario;
    sorgente = pezzi(rinvio.proprietario + '/app-consegne-sorgente');
  } else throw new Error('il reparto «' + reparto.codice + '» non dice account e repository di GitHub (voce «github» del suo file)');
  const proprietari = [];
  [utente, sorgente.proprietario, rinvio.proprietario, sitoServizio && sitoServizio.proprietario,
    p.sitoPubblico && p.sitoPubblico.repository.split('/')[0], p.githubPages && p.githubPages.repository.split('/')[0]]
    .filter(Boolean).forEach((x) => { if (proprietari.indexOf(x.toLowerCase()) < 0) proprietari.push(x.toLowerCase()); });
  return { utente, sorgente, rinvio, sitoServizio, proprietari };
}
// L'indirizzo esatto della pagina d'accesso a cui la squadra Access del reparto rimanda chi non è entrato
// in un suo sito, senza la parte variabile («?kid=…»); null se il file del reparto non nomina la squadra.
function paginaAccesso(squadra, host) { return 'https://' + squadra + '.cloudflareaccess.com/cdn-cgi/access/login/' + String(host || '').toLowerCase(); }
function cancelloAccess(reparto, host) {
  const s = reparto.pubblicazione && reparto.pubblicazione.cloudflare && reparto.pubblicazione.cloudflare.squadra;
  return s ? paginaAccesso(s, host) : null;
}

// Il testo di js/ambiente.js per un reparto. Solo ciò che serve all'app, nient'altro del file del reparto.
function ambienteJs(reparto) {
  const d = difetti(reparto);
  if (d.length) throw new Error('reparto non valido: ' + d.join('; '));
  const dati = {
    codice: reparto.codice,
    nome: reparto.nome,
    ambiente: reparto.ambiente,
    lingua: reparto.lingua,
    versioneLocale: reparto.versioneLocale,
    host: reparto.host.slice(),
    supabaseUrl: reparto.supabaseUrl,
    supabaseAnonKey: reparto.supabaseAnonKey,
    googleClientId: reparto.googleClientId,
  };
  if (reparto.centrale) dati.centrale = { url: reparto.centrale.url };
  // dalla v199: la squadra di Cloudflare Access del sito, per la riga di rete.html che prova se il cancello
  // si raggiunge. Solo il nome: l'app non la usa, e un reparto senza cancello non la porta
  const squadra = reparto.pubblicazione.cloudflare && reparto.pubblicazione.cloudflare.squadra;
  if (squadra) dati.squadraAccess = squadra;
  // JSON è anche JavaScript, tolti due separatori di riga (codici 2028 e 2029) che in un file .js vanno scritti per esteso
  const json = JSON.stringify(dati, null, 2)
    .split(String.fromCharCode(0x2028)).join(String.fromCharCode(92) + 'u2028')
    .split(String.fromCharCode(0x2029)).join(String.fromCharCode(92) + 'u2029');
  return [
    '// Ambiente di questo sito: reparto «' + reparto.codice + '».',
    '// File GENERATO da collaudo/strumenti/reparti.js a partire da',
    '// reparti/' + reparto.codice + '.json: non si modifica a mano.',
    '// Contiene solo dati pubblici per natura (indirizzi del sito, indirizzo e',
    '// chiave pubblica del database, identificativo del client Google): mai un',
    '// segreto. Va caricato PRIMA di js/api.js, che da qui sa a quale database',
    '// collegarsi e da quali indirizzi può essere aperto.',
    'window._AMBIENTE = ' + json + ';',
    '',
  ].join('\n');
}

const senzaCR = (s) => String(s).replace(/\r\n/g, '\n');
// docs/js/ambiente.js deve essere, a parte i fine riga, quello generato per il reparto che pubblica da docs/.
function controllaDocs() {
  const atteso = ambienteJs(delSitoInDocs());
  if (!fs.existsSync(AMBIENTE_JS)) return { ok: false, motivo: 'docs/js/ambiente.js non esiste' };
  const trovato = senzaCR(fs.readFileSync(AMBIENTE_JS, 'utf8'));
  return trovato === atteso ? { ok: true, reparto: delSitoInDocs().codice } : { ok: false, motivo: 'docs/js/ambiente.js non è quello generato da reparti/' + delSitoInDocs().codice + '.json: rifarlo con «node collaudo/strumenti/reparti.js scrivi»' };
}

module.exports = { RADICE, CARTELLA, AMBIENTE_JS, CAMPI, difetti, codici, leggi, tutti, convalida, risolvi, nomeStorico, perProgettoCloudflare, delSitoInDocs, github, cancelloAccess, paginaAccesso, hostCoperto, hostAmmesso, ambienteJs, controllaDocs, ruoloDellaChiave };

if (require.main === module) {
  try {
    const [azione, arg] = process.argv.slice(2);
    if (process.argv.length > 4 || (azione && ['mostra', 'controlla', 'scrivi'].indexOf(azione) < 0)) throw new Error('uso: reparti.js [mostra <codice> | controlla | scrivi]');
    if (!azione) {
      tutti().forEach((r) => console.log(r.codice.padEnd(16) + r.tipo.padEnd(12) + 'livello ' + r.livello + '  ' + r.hosting.padEnd(10) + r.canale.padEnd(8) + r.host.join(', ')));
      console.log(tutti().length + ' reparti, tutti validi; da docs/ pubblica: ' + delSitoInDocs().codice);
    } else if (azione === 'mostra') {
      // la chiave pubblica non è un segreto, ma sullo schermo se ne mostra solo l'inizio
      process.stdout.write(ambienteJs(risolvi(arg)).replace(/("supabaseAnonKey": ")([^"]{12})[^"]*/, '$1$2…'));
    } else if (azione === 'controlla') {
      const c = controllaDocs();
      console.log(c.ok ? 'docs/js/ambiente.js è quello del reparto ' + c.reparto : 'ERRORE: ' + c.motivo);
      if (!c.ok) process.exitCode = 1;
    } else {
      fs.mkdirSync(path.dirname(AMBIENTE_JS), { recursive: true });
      fs.writeFileSync(AMBIENTE_JS, ambienteJs(delSitoInDocs()));
      console.log('scritto docs/js/ambiente.js per il reparto ' + delSitoInDocs().codice);
    }
  } catch (e) { console.log('ERRORE: ' + e.message); process.exitCode = 1; }
}
