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
  'versioneLocale', 'livello', 'host', 'pubblicazione', 'supabaseUrl', 'supabaseAnonKey', 'googleClientId'];
const TIPI = ['produzione', 'collaudo'];
const HOSTING = ['reparto', 'fornitore', 'condiviso', 'locale'];
const ACCESSI = ['google'];
const CANALI = ['pilota', 'stabile'];
const LINGUE = ['it'];
// Un nome host scritto per esteso, oppure (solo nel collaudo) col punto davanti: vale per i sottodomini.
const HOST = /^\.?[a-z0-9]([a-z0-9-]*[a-z0-9])?(\.[a-z0-9]([a-z0-9-]*[a-z0-9])?)*$/;
const CODICE = /^[a-z0-9]([a-z0-9-]{0,40}[a-z0-9])?$/;
const LOCALI = ['localhost', '127.0.0.1'];

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
    if (p.cloudflare && !(typeof p.cloudflare.progetto === 'string' && /^[a-z0-9]([a-z0-9-]{0,56}[a-z0-9])?$/.test(p.cloudflare.progetto) && Object.keys(p.cloudflare).length === 1)) e.push('pubblicazione.cloudflare: atteso il solo campo «progetto», con un nome di progetto valido');
    if (p.cloudflare && Array.isArray(r.host) && r.host.indexOf(p.cloudflare.progetto + '.pages.dev') < 0) e.push('pubblicazione.cloudflare: «' + p.cloudflare.progetto + '.pages.dev» non è fra gli host del reparto');
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
function tutti() {
  const elenco = codici().map((c) => leggi(c));
  const e = [];
  elenco.forEach((r, i) => difetti(r, codici()[i]).forEach((d) => e.push('reparti/' + codici()[i] + '.json: ' + d)));
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
  // un repository serve un sito solo. Fra quelli DICHIARATI nei file dei reparti: né due siti pubblici nello stesso,
  // né un sito pubblico nel repository da cui un reparto pubblica docs/. Un repository col sorgente che nessun file
  // dichiara (quello del collaudo) qui non si conosce: lo ferma sito-pubblico.js, che prima di scrivere guarda che
  // cosa il repository contiene
  const depositi = {};
  elenco.forEach((r) => ['githubPages', 'sitoPubblico'].forEach((voce) => {
    const dep = r.pubblicazione[voce] && String(r.pubblicazione[voce].repository).toLowerCase();
    if (!dep) return;
    if (depositi[dep] && depositi[dep] !== r.codice + ' ' + voce) e.push('il repository ' + dep + ' è dichiarato due volte: ' + depositi[dep] + ' e ' + r.codice + ' ' + voce);
    depositi[dep] = r.codice + ' ' + voce;
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

module.exports = { RADICE, CARTELLA, AMBIENTE_JS, CAMPI, difetti, codici, leggi, tutti, risolvi, nomeStorico, perProgettoCloudflare, delSitoInDocs, hostCoperto, hostAmmesso, ambienteJs, controllaDocs, ruoloDellaChiave };

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
