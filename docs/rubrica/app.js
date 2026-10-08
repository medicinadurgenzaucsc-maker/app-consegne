// ═══════════════════════════════════════════════════════════════════════════
//  Rubrica — i «Numeri Telefono» dell'app Consegne
//
//  La pagina vive DENTRO l'app: un riquadro della stessa origine, aperto dal
//  menu «Numeri Telefono». Non ha un database né una chiave suoi: indirizzo,
//  chiave pubblica e sessione li chiede ogni volta alla pagina che la contiene
//  (window.parent). Aperta da sola, o con l'app senza accesso, resta sul
//  messaggio fisso di index.html e non chiede nulla al database.
//
//  Regole (CLAUDE.md, «Rubrica»):
//   - nomi e numeri stanno SOLO in memoria: mai nella memoria del browser, mai
//     in una cache; sul dispositivo restano solo tema, preferiti e chiamati di
//     recente (identificativi, nessun dato);
//   - ogni testo che arriva dal database o dalla memoria del browser passa da
//     esc() prima di entrare in innerHTML; un numero diventa un collegamento
//     solo se è fatto di cifre; un identificativo è sempre un intero;
//   - niente gestori e niente stili scritti in linea, nemmeno nell'HTML
//     costruito qui: la regola CSP della pagina li spegne (un aspetto nuovo va
//     in app.css, come classe);
//   - nessun service worker proprio, nessun client del database proprio, e la
//     sessione non si chiude mai da qui: è quella dell'app;
//   - una modifica si scrive solo se la riga è ancora quella su cui la scheda
//     è stata aperta, e dopo una scrittura rimasta senza risposta non si
//     invita a riprovare alla cieca: si rilegge e si guarda se c'è già;
//   - il fuoco si prende e si rende solo se questa pagina lo ha ancora: chi
//     nel frattempo è tornato a scrivere in una scheda dell'app non va
//     interrotto.
// ═══════════════════════════════════════════════════════════════════════════

// ── Costanti ─────────────────────────────────────────────────────────────────
// Memoria del browser: solo preferenze senza dati personali. I nomi sono nuovi
// rispetto alla rubrica di prima, così ciò che quella ha lasciato nel browser
// non viene mai creduto buono. Ogni lettura è trattata da ostile (lsLeggi…).
const LS_TEMA       = 'rubrica-tema';
const LS_PREFERITI  = 'rubrica-preferiti';
const LS_CHIAMATI   = 'rubrica-chiamati';
// Chiavi della rubrica di prima che tenevano nomi, numeri o richieste in
// sospeso: all'avvio si tolgono dal dispositivo.
const LS_RESIDUI    = ['rubrica-data', 'rubrica-cats', 'rubrica-ts', 'rubrica-pending-writes', 'rubrica-recent-searches'];
// Preferenze lasciate sul dispositivo dalla rubrica di prima, quando stava
// sullo stesso sito: alla prima apertura preferiti e chiamati di recente si
// riprendono (riprendiPreferenzeDiPrima), poi le tre chiavi si tolgono.
const LS_PREFERITI_DI_PRIMA = 'rubrica-favs';
const LS_CHIAMATI_DI_PRIMA  = 'rubrica-recent-calls';
const LS_TEMA_DI_PRIMA      = 'rubrica-theme';
const MAX_RECENT    = 5;
const MAX_CALLS     = 8;
const MAX_PREFERITI = 500;
const STAR_SVG   = `<svg width="11" height="11" viewBox="0 0 24 24" fill="currentColor" stroke="currentColor" stroke-width="1" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><polygon points="12 2 15.09 8.26 22 9.27 17 14.14 18.18 21.02 12 17.77 5.82 21.02 7 14.14 2 9.27 8.91 8.26 12 2"/></svg>`;
const ALPHA_KEYS = ['0-9', ...'ABCDEFGHIJKLMNOPQRSTUVWXYZ'];
const TRASH_SVG = `<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><polyline points="3 6 5 6 21 6"/><path d="M19 6l-1 14a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2L5 6"/><path d="M10 11v6"/><path d="M14 11v6"/><path d="M9 6V4a1 1 0 0 1 1-1h4a1 1 0 0 1 1 1v2"/></svg>`;
const DRAG_SVG  = `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><line x1="4" y1="6" x2="20" y2="6"/><line x1="4" y1="12" x2="20" y2="12"/><line x1="4" y1="18" x2="20" y2="18"/></svg>`;
const EDIT_SVG  = `<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>`;

// ── DOM cache (evita decine di getElementById) ───────────────────────────────
const $ = id => document.getElementById(id);
const D = {};
function cacheDOM() {
  [
    'contactList','countLabel','searchBar','searchWrap','btnSearchClear',
    'searchSuggestions',
    'btnTheme','btnSettings','btnAll','btnAlpha','btnNew',
    'chips','installToast',
    'modalOverlay','modalTitle','modalBox','contactForm',
    'fId','fNome','fCategoria','numeriContainer','btnAddNum',
    'btnCancel','btnDelete','btnSave','btnSaveVcf',
    'catModalOverlay','catModalBox','catList','fNewCat','btnAddCat',
    'btnCatCancel','btnCatSave',
    'alphaModalOverlay','alphaGrid','btnAlphaClose',
    'numMenuOverlay','numMenuTitle','numMenuSub',
  ].forEach(id => { D[id] = $(id); });
}

// ── Stato globale ────────────────────────────────────────────────────────────
// Contatti e categorie vivono solo qui, in memoria: spariscono con la pagina.
let allContacts = [];
let categories  = [];
let versioneNota = null;   // rubrica_versione.ts dei dati che sono in memoria
let datiPronti   = false;
let activeCategory = null;
let activeSearch   = '';
let activeAlpha    = null;

// ── Pagina madre e sessione ──────────────────────────────────────────────────
// La pagina che contiene il riquadro, se è l'app: stessa origine, col client
// del database (_sb), l'indirizzo e la chiave pubblica. Da un'altra origine
// anche solo leggerla lancia un errore: vale come «non c'è».
function madre() {
  try {
    const m = window.parent;
    if (!m || m === window) return null;
    if (!m._sb || !m._sb.auth || typeof m._sb.auth.getSession !== 'function') return null;
    if (typeof m.SUPABASE_URL !== 'string' || typeof m.SUPABASE_ANON_KEY !== 'string') return null;
    if (m.SUPABASE_URL.slice(0, 8) !== 'https://' || !m.SUPABASE_ANON_KEY) return null;
    return m;
  } catch (_) {
    return null;
  }
}

// La sessione manca, è scaduta o non basta. «motivo»: fuori-app | assente |
// scaduta | non-autorizzato.
class ErroreSessione extends Error {
  constructor(motivo) {
    super('sessione: ' + motivo);
    this.name = 'ErroreSessione';
    this.motivo = motivo;
  }
}
// Un errore con un messaggio già adatto a chi usa la rubrica.
function erroreChiaro(testo) {
  const e = new Error(testo);
  e.chiaro = true;
  return e;
}

// Un errore «di rete»: la risposta non è arrivata (rete assente, tempo massimo
// scaduto, corpo interrotto), oppure non si sa se la sessione c'è ancora. Per
// una SCRITTURA vuol dire esito incerto: la richiesta può essere arrivata.
function erroreRete() {
  const e = new Error('rete');
  e.rete = true;
  return e;
}

// Il token della sessione si chiede alla pagina madre A OGNI richiesta e non si
// conserva mai qui: dura circa un'ora, e la pagina resta aperta tutto il giorno
// (lo rinnova la libreria dell'app). Con «forza» si chiede un rinnovo subito.
// Mai la chiave pubblica al posto del token: otterrebbe solo un rifiuto.
//
// Risponde { tok, certo }. «Sessione assente» si dice solo quando lo dice la
// libreria: nessuna sessione e nessun errore, oppure un «no» del servizio di
// accesso. Se invece la libreria non risponde in tempo, o risponde con un
// errore che non è un rifiuto (un rinnovo fallito per la rete), NON SI SA:
// vale l'ultimo token noto alla pagina madre, con «certo» falso, e chi chiama
// non può concluderne che la sessione sia scaduta.
const ATTESA_SESSIONE_MS = 6000;
function rifiutoDelServizio(errore) {
  const stato = Number(errore && errore.status);
  return stato >= 400 && stato < 500 && stato !== 408 && stato !== 429;
}
async function gettone(m, forza) {
  const chiesta = Promise.resolve()
    .then(() => (forza ? m._sb.auth.refreshSession() : m._sb.auth.getSession()))
    .then(r => ({ sessione: (r && r.data && r.data.session) || null, errore: (r && r.error) || null }), () => null);
  const esito = await Promise.race([chiesta, new Promise(r => setTimeout(() => r(null), ATTESA_SESSIONE_MS))]);
  const buono = t => typeof t === 'string' && !!t && t !== m.SUPABASE_ANON_KEY;
  if (esito && esito.sessione) {
    if (!buono(esito.sessione.access_token)) throw new ErroreSessione('assente');
    return { tok: esito.sessione.access_token, certo: true };
  }
  if (esito && (!esito.errore || rifiutoDelServizio(esito.errore))) throw new ErroreSessione('assente');
  // non si sa: resta l'ultimo token noto alla pagina madre
  if (!buono(m._supaAccessToken)) throw erroreRete();
  return { tok: m._supaAccessToken, certo: false };
}

// ── Accesso al database ──────────────────────────────────────────────────────
// UNICO punto da cui partono le richieste. Si possono chiedere solo le tabelle
// e le funzioni della rubrica; chiave e token vengono dopo le intestazioni di
// chi chiama, che quindi non le può sostituire. Risponde { stato, testo }.
//  - 401: un solo rinnovo forzato della sessione (chiesto solo per la
//    richiesta subito dopo, non per i tentativi seguenti), poi si rinuncia;
//    «sessione scaduta» si dice solo se il rinnovo ha risposto, altrimenti
//    è un errore di rete: basta riprovare;
//  - nuovi tentativi automatici solo sulle LETTURE (una scrittura ripetuta dopo
//    una risposta persa creerebbe un doppione);
//  - ogni richiesta ha un tempo massimo, lettura del corpo compresa: oltre,
//    viene interrotta e vale come rete assente. Senza, una risposta che non
//    arriva terrebbe ferma la fila dei caricamenti e i pulsanti di salvataggio;
//  - 403: «account non autorizzato» si dice solo per una LETTURA respinta. Lo
//    stesso rifiuto su una scrittura (un permesso che manca su una tabella,
//    una sequenza o una funzione) è un errore di quell'operazione.
const PERCORSO_AMMESSO = /^(?:rpc\/)?rubrica_[a-z_]+(?:\?[A-Za-z0-9_.,=&-]*)?$/;
const MAX_TENTATIVI = 2;
const ATTESA_RICHIESTA_MS = 12000;
async function supaRisposta(path, options = {}) {
  const m = madre();
  if (!m) throw new ErroreSessione('fuori-app');
  if (typeof path !== 'string' || !PERCORSO_AMMESSO.test(path) || path.includes('..')) throw erroreChiaro('Richiesta non ammessa');
  const metodo  = options.method || 'GET';
  const lettura = metodo === 'GET';
  let rinnovato = false;   // l'unico rinnovo forzato della sessione è già stato chiesto
  let forza     = false;   // vale per UNA richiesta: quella subito dopo un 401
  let tentativo = 0;       // nuovi tentativi già fatti (solo letture)
  for (;;) {
    const g = await gettone(m, forza);
    forza = false;
    const interruttore = typeof AbortController === 'function' ? new AbortController() : null;
    const sveglia = interruttore ? setTimeout(() => interruttore.abort(), ATTESA_RICHIESTA_MS) : null;
    let stato = 0, testo = '';
    try {
      const res = await fetch(`${m.SUPABASE_URL}/rest/v1/${path}`, {
        method: metodo,
        body: options.body,
        cache: 'no-store',          // le risposte non restano nemmeno nella memoria del browser
        signal: interruttore ? interruttore.signal : undefined,
        headers: {
          'Content-Type': 'application/json',
          ...(options.headers || {}),
          'apikey': m.SUPABASE_ANON_KEY,
          'Authorization': `Bearer ${g.tok}`,
        },
      });
      testo = await res.text();
      stato = res.status;
    } catch (_) {
      stato = 0;                    // rete assente, tempo scaduto o corpo interrotto
    }
    if (sveglia) clearTimeout(sveglia);
    if (stato === 0) {
      if (lettura && tentativo < MAX_TENTATIVI) { await sleep(200 * Math.pow(2, tentativo++)); continue; }
      throw erroreRete();
    }
    if (stato === 401) {
      if (!rinnovato) { rinnovato = true; forza = true; continue; }
      if (!g.certo) throw erroreRete();   // il rinnovo non ha risposto: non si sa se la sessione c'è
      throw new ErroreSessione('scaduta');
    }
    if (stato >= 500 && lettura && tentativo < MAX_TENTATIVI) { await sleep(200 * Math.pow(2, tentativo++)); continue; }
    if (stato < 200 || stato >= 300) {
      let err = null;
      try { err = JSON.parse(testo); } catch (_) { err = null; }
      const codice = String((err && err.code) || '');
      if (stato === 403 && lettura && (!codice || codice === '42501')) throw new ErroreSessione('non-autorizzato');
      const e = new Error(String((err && (err.message || err.hint)) || ('HTTP ' + stato)));
      e.stato  = stato;
      e.codice = codice;
      // ciò che era stato inviato: serve a spiegare un rifiuto (messaggioErrore)
      try { e.dati = options.body ? JSON.parse(options.body) : null; } catch (_) { e.dati = null; }
      throw e;
    }
    return { stato, testo };
  }
}
async function supaFetch(path, options = {}) {
  const { testo } = await supaRisposta(path, options);
  if (!testo) return [];
  try { return JSON.parse(testo); } catch (_) {
    // la risposta è «riuscita» ma non si legge: per una scrittura l'esito è incerto
    const e = erroreChiaro('Risposta del database non leggibile');
    e.incerto = true;
    throw e;
  }
}
function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

// Legge TUTTE le righe, a pagine: il servizio non ne restituisce più di un
// certo numero per richiesta (di norma 1000) e taglierebbe il resto senza
// dirlo. Si continua finché arriva una pagina vuota; «percorso» ha già il suo
// ordinamento stabile (order=id).
const PAGINA     = 1000;
const MAX_PAGINE = 50;
async function leggiTutte(percorso) {
  const tutte = [];
  for (let giro = 0; giro < MAX_PAGINE; giro++) {
    const righe = await supaFetch(`${percorso}&limit=${PAGINA}&offset=${tutte.length}`);
    if (!Array.isArray(righe)) throw erroreChiaro('Risposta del database non leggibile');
    if (!righe.length) return tutte;
    for (const r of righe) tutte.push(r);
  }
  throw erroreChiaro('La rubrica è troppo grande per essere letta');
}

// ── Messaggi d'errore ────────────────────────────────────────────────────────
// Chi usa la rubrica non deve mai leggere il testo grezzo del database.
const MESSAGGI_SESSIONE = {
  'fuori-app':       'La rubrica si apre dal menu “Numeri Telefono” dell\'app',
  'assente':         'Sessione non attiva: rientra nell\'applicazione',
  'scaduta':         'Sessione non attiva: rientra nell\'applicazione',
  'non-autorizzato': 'Account non autorizzato a usare la rubrica',
  'non-visibile':    'Il database non mostra la rubrica a questo account',
};
// Le funzioni della rubrica nel database rispondono con frasi già in italiano
// («Esiste già una categoria…»): quelle si mostrano. Le frasi del database
// vero e proprio cominciano così, e non si mostrano mai:
const GREZZO_DEL_DATABASE = /^(new row|duplicate key|null value|value too long|update or delete|insert or update|permission denied|invalid input|could not|relation |column |function |syntax error|JWT|operator )/i;
const CODICI_DELLE_FUNZIONI = ['P0001', 'P0002', '22023', '23503', '23505'];
// Tabelle o funzioni della rubrica che il database non conosce: il codice è
// arrivato su un sito il cui database non è ancora stato preparato. Non è un
// guasto di rete, e controllare la connessione non serve.
const CODICI_RUBRICA_ASSENTE = ['PGRST205', 'PGRST202', '42P01', '42883'];
function rubricaAssente(e) {
  if (!e || e instanceof ErroreSessione || e.stato == null) return false;
  return e.stato === 404 || CODICI_RUBRICA_ASSENTE.includes(String(e.codice || ''));
}

// Perché il database rifiuterebbe questi dati (gli stessi vincoli delle
// tabelle): frase da mostrare, oppure '' se sono in regola.
function motivoRifiuto(d) {
  if (!d || typeof d !== 'object') return '';
  const s = k => (typeof d[k] === 'string' ? d[k] : '');
  if (/["<>]/.test(s('numeri'))) return 'Il numero contiene caratteri non ammessi';
  if (/[<>]/.test(s('nome') + s('categoria') + s('note'))) return 'Il testo contiene caratteri non ammessi (< e >)';
  if (s('nome').length > 200 || s('categoria').length > 80 || s('numeri').length > 300 || s('note').length > 600) return 'Testo troppo lungo';
  if ('nome' in d && !s('nome').trim()) return 'Il nome è obbligatorio';
  return '';
}
function messaggioErrore(e) {
  if (e instanceof ErroreSessione) return MESSAGGI_SESSIONE[e.motivo] || MESSAGGI_SESSIONE.assente;
  if (e && e.chiaro) return String(e.message);
  if (e && e.rete) return 'Connessione assente o instabile: riprova';
  if (!e || e.stato == null) return 'Operazione non riuscita';
  if (rubricaAssente(e)) return 'La rubrica non è ancora disponibile su questo sito';
  if (e.stato === 403) return 'Il database non permette questa operazione: avvisa chi gestisce l\'applicazione';
  const cod    = String(e.codice || '');
  const grezzo = String(e.message || '');
  // 23514 = un vincolo delle tabelle ha rifiutato i dati
  if (cod === '23514') {
    return motivoRifiuto(e.dati)
      || (/numer/i.test(grezzo) ? 'Il numero contiene caratteri non ammessi'
                                : 'Dati non accettati: caratteri non ammessi oppure testo troppo lungo');
  }
  if (cod === '22001') return 'Testo troppo lungo';
  if (cod === '23502') return 'Manca un dato obbligatorio';
  if (grezzo && !GREZZO_DEL_DATABASE.test(grezzo) && CODICI_DELLE_FUNZIONI.includes(cod)) return grezzo;
  if (cod === '23505') return 'Esiste già una voce con lo stesso nome';
  if (cod === '23503') return 'La voce è ancora in uso: non può essere eliminata';
  return 'Operazione non riuscita (errore ' + (cod || e.stato) + ')';
}
// Che cosa fare, sotto il messaggio che prende il posto dell'elenco: di
// connessione si parla solo quando l'errore è di rete.
function consiglioErrore(e) {
  if (e instanceof ErroreSessione) {
    if (e.motivo === 'non-autorizzato') return 'Rivolgiti a chi gestisce l\'applicazione';
    if (e.motivo === 'non-visibile') return 'Può non essere ancora pronta, oppure l\'account non è fra quelli autorizzati: avvisa chi gestisce l\'applicazione';
    return 'Ricarica l\'applicazione ed entra di nuovo';
  }
  if (e && e.rete) return 'Controlla la connessione e riprova';
  if (rubricaAssente(e)) return 'Il database di questo sito non è ancora stato preparato: avvisa chi gestisce l\'applicazione';
  return 'Riprova; se continua, avvisa chi gestisce l\'applicazione';
}

// ── Helpers ──────────────────────────────────────────────────────────────────
// Testo da mettere in pagina dentro innerHTML o in un attributo: codifica
// anche l'apice, e un valore assente diventa testo vuoto.
const ENTITA = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' };
function esc(s) {
  return String(s == null ? '' : s).replace(/[&<>"']/g, c => ENTITA[c]);
}
// Un identificativo è un intero positivo: tutto il resto non entra né in
// pagina né nell'indirizzo di una richiesta.
function idIntero(v) {
  const s = typeof v === 'number' ? String(v) : (typeof v === 'string' ? v.trim() : '');
  if (!/^[0-9]{1,15}$/.test(s)) return null;
  const n = Number(s);
  return n > 0 ? n : null;
}
// Un numero diventa un collegamento «tel:» (in pagina, nel menu, nella scheda
// del telefono) solo se è fatto di sole cifre, con un eventuale + davanti.
const NUMERO_TEL = /^\+?[0-9]{1,20}$/;
function debounce(fn, ms) {
  let t;
  return (...args) => { clearTimeout(t); t = setTimeout(() => fn(...args), ms); };
}
function formatCatName(s) {
  return String(s == null ? '' : s).replace(/[^a-zA-ZÀ-ÿ0-9 \-_.]/g, '').toUpperCase().trim().slice(0, 80);
}

// ── Memoria del browser: letture e scritture protette ───────────────────────
// La memoria è in comune con l'app (stessa origine) e può essere piena,
// spenta o scritta da altri: nessuna di queste funzioni lancia un errore.
function lsLeggi(chiave) {
  try { return window.localStorage.getItem(chiave); } catch (_) { return null; }
}
function lsScrivi(chiave, valore) {
  try { window.localStorage.setItem(chiave, valore); } catch (_) {}
}
function lsTogli(chiave) {
  try { window.localStorage.removeItem(chiave); } catch (_) {}
}
// Un elenco di identificativi: si accetta solo un elenco, e di quello solo gli
// interi positivi, senza doppioni, fino a «massimo».
function lsIdentificativi(chiave, massimo) {
  let dati = null;
  try { dati = JSON.parse(lsLeggi(chiave) || 'null'); } catch (_) { dati = null; }
  if (!Array.isArray(dati)) return [];
  const fuori = [];
  for (const x of dati) {
    const n = idIntero(x);
    if (n === null || fuori.includes(n)) continue;
    fuori.push(n);
    if (fuori.length >= massimo) break;
  }
  return fuori;
}
function pulisciResidui() {
  LS_RESIDUI.forEach(lsTogli);
}
// Preferiti e chiamati di recente lasciati dalla rubrica di prima: si
// riprendono UNA volta, solo se qui non ce ne sono ancora, e solo come
// identificativi interi (lsIdentificativi scarta tutto il resto: ciò che sta
// in quelle chiavi non è creduto buono). Poi le chiavi di prima si tolgono.
// Il tema NON si riprende: quella rubrica lo salvava anche senza una scelta
// (scuro), e qui il tema predefinito è chiaro.
function riprendiPreferenzeDiPrima() {
  if (lsLeggi(LS_PREFERITI_DI_PRIMA) !== null && lsLeggi(LS_PREFERITI) === null) {
    const preferiti = lsIdentificativi(LS_PREFERITI_DI_PRIMA, MAX_PREFERITI);
    if (preferiti.length) lsScrivi(LS_PREFERITI, JSON.stringify(preferiti));
  }
  if (lsLeggi(LS_CHIAMATI_DI_PRIMA) !== null && lsLeggi(LS_CHIAMATI) === null) {
    const chiamati = lsIdentificativi(LS_CHIAMATI_DI_PRIMA, MAX_CALLS);
    if (chiamati.length) lsScrivi(LS_CHIAMATI, JSON.stringify(chiamati));
  }
  [LS_PREFERITI_DI_PRIMA, LS_CHIAMATI_DI_PRIMA, LS_TEMA_DI_PRIMA].forEach(lsTogli);
}

// ── Avatar: tinta stabile dalla categoria ────────────────────────────────────
function hashHue(str) {
  let h = 0;
  for (let i = 0; i < str.length; i++) h = (h * 31 + str.charCodeAt(i)) & 0xffffff;
  return Math.abs(h) % 360;
}
// Una delle 24 classi di colore di app.css (av-0 … av-23): solo un numero
// calcolato qui, mai il testo della categoria.
function avatarClasse(categoria) {
  return 'av-' + (Math.round(hashHue(String(categoria || '')) / 15) % 24);
}
function avatarLetter(nome) {
  const s = String(nome || '').trim();
  return s ? s[0].toUpperCase() : '?';
}

// ── Highlight match nei risultati ────────────────────────────────────────────
function escRegex(s) { return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }

// `raw` è il testo così com'è (NON ancora codificato); tokens sono in
// minuscolo. Prima si taglia il testo sulle corrispondenze, poi ogni pezzo
// viene codificato: così un'entità (&amp;, &#39;) non può essere spezzata e
// nel risultato l'unico elemento è <mark>.
function highlight(raw, tokens) {
  const s = String(raw == null ? '' : raw);
  if (!s || !tokens || !tokens.length) return esc(s);
  const re = new RegExp('(' + tokens.map(escRegex).join('|') + ')', 'gi');
  return s.split(re).map((p, i) => (i % 2 ? '<mark>' + esc(p) + '</mark>' : esc(p))).join('');
}

// ── Haptic feedback ──────────────────────────────────────────────────────────
function haptic(pattern = 8) {
  if ('vibrate' in navigator) navigator.vibrate(pattern);
}

// ── Ricerche recenti ─────────────────────────────────────────────────────────
// Solo in memoria: sono nomi e numeri digitati, sul dispositivo non restano.
let ricercheRecenti = [];
function getRecentSearches() {
  return ricercheRecenti.slice();
}
function pushRecentSearch(q) {
  q = String(q || '').trim().slice(0, 80);
  if (!q || q.length < 2) return;
  ricercheRecenti = [q, ...ricercheRecenti.filter(x => x.toLowerCase() !== q.toLowerCase())].slice(0, MAX_RECENT);
}
function removeRecentSearch(q) {
  ricercheRecenti = ricercheRecenti.filter(x => x !== q);
}
function renderRecentSuggestions() {
  const arr = getRecentSearches();
  if (!arr.length || activeSearch || document.activeElement !== D.searchBar) {
    D.searchSuggestions.hidden = true;
    return;
  }
  D.searchSuggestions.innerHTML = arr.map(q =>
    `<button class="search-suggestion" data-q="${esc(q)}" type="button">
       <span>${esc(q)}</span><span class="sug-x" data-x="${esc(q)}">×</span>
     </button>`
  ).join('');
  D.searchSuggestions.hidden = false;
}

// ── Salvataggio scroll position fra modal e ritorno ──────────────────────────
let savedScrollY = 0;

// ── Preferiti (★) ────────────────────────────────────────────────────────────
// In memoria del browser solo gli identificativi dei contatti.
function getFavs() {
  return new Set(lsIdentificativi(LS_PREFERITI, MAX_PREFERITI));
}
function setFavs(set) {
  lsScrivi(LS_PREFERITI, JSON.stringify([...set].slice(0, MAX_PREFERITI)));
}
function isFav(id) { return favs.has(idIntero(id)); }
function toggleFav(id) {
  id = idIntero(id);
  if (id === null) return;
  if (favs.has(id)) favs.delete(id); else favs.add(id);
  setFavs(favs);
}
let favs = new Set();   // riempito da init()

// ── Recenti chiamati ─────────────────────────────────────────────────────────
function getRecentCalls() {
  return lsIdentificativi(LS_CHIAMATI, MAX_CALLS);
}
function pushRecentCall(id) {
  id = idIntero(id);
  if (id === null) return;
  const arr = [id, ...getRecentCalls().filter(x => x !== id)].slice(0, MAX_CALLS);
  lsScrivi(LS_CHIAMATI, JSON.stringify(arr));
}

// ── Filtri speciali (Preferiti / Recenti chiamati) ───────────────────────────
let activeSpecial = null; // 'fav' | 'recent' | null

// ── Tema con 3 stati: light / dark / auto ────────────────────────────────────
// Il riquadro sta dentro un'app chiara: senza una scelta il tema è chiaro. In
// memoria finisce solo la scelta fatta col pulsante, mai il valore predefinito.
const TEMI = ['light', 'dark', 'auto'];
let temaAttuale = 'light';
let mediaQ = window.matchMedia('(prefers-color-scheme: dark)');
function resolveTheme(mode) {
  if (mode === 'auto') return mediaQ.matches ? 'dark' : 'light';
  return mode === 'dark' ? 'dark' : 'light';
}
function temaSalvato() {
  const t = lsLeggi(LS_TEMA);
  return TEMI.includes(t) ? t : 'light';
}
function applyTheme(mode, ricorda) {
  if (!TEMI.includes(mode)) mode = 'light';
  temaAttuale = mode;
  document.documentElement.setAttribute('data-theme', resolveTheme(mode));
  if (D.btnTheme) D.btnTheme.dataset.themeMode = mode;
  if (ricorda) lsScrivi(LS_TEMA, mode);
}
function cycleTheme() {
  const next = temaAttuale === 'light' ? 'dark' : (temaAttuale === 'dark' ? 'auto' : 'light');
  applyTheme(next, true);
  showToast('Tema: ' + ({ light: 'chiaro', dark: 'scuro', auto: 'automatico' })[next], 1500);
}
mediaQ.addEventListener('change', () => {
  if (temaAttuale === 'auto') applyTheme('auto', false);
});

// ── Focus trap nei modal ─────────────────────────────────────────────────────
let activeModal = null;
let lastFocused = null;
const FOCUSABLE = 'button:not([hidden]), [href], input:not([type="hidden"]), select, textarea, [tabindex]:not([tabindex="-1"])';
function trapFocus(modalEl) {
  activeModal = modalEl;
  lastFocused = document.activeElement;
  const handler = e => {
    if (e.key === 'Escape') {
      e.preventDefault();
      // con una domanda a schermo Esc vale «Annulla» per la domanda, e la finestra sotto resta
      if (confermaInSospeso) annullaConferma(); else closeActiveModal();
      return;
    }
    if (e.key !== 'Tab') return;
    const els = [...modalEl.querySelectorAll(FOCUSABLE)].filter(el => !el.disabled && el.offsetParent !== null);
    if (!els.length) return;
    const first = els[0], last = els[els.length - 1];
    if (e.shiftKey && document.activeElement === first) { e.preventDefault(); last.focus(); }
    else if (!e.shiftKey && document.activeElement === last) { e.preventDefault(); first.focus(); }
  };
  modalEl._trapHandler = handler;
  document.addEventListener('keydown', handler);
}
function releaseFocus(modalEl) {
  if (modalEl?._trapHandler) {
    document.removeEventListener('keydown', modalEl._trapHandler);
    modalEl._trapHandler = null;
  }
  activeModal = null;
  // Il fuoco si rende solo se questa pagina lo ha ancora: una finestra che si
  // chiude DOPO un'attesa di rete non deve riprenderselo a chi nel frattempo
  // è tornato a scrivere in una scheda dell'app.
  if (document.hasFocus() && lastFocused?.focus) lastFocused.focus();
  lastFocused = null;
}
function closeAlphaModal() {
  releaseFocus(D.alphaModalOverlay);
  D.alphaModalOverlay.hidden = true;
}
function closeActiveModal() {
  if (!activeModal) return;
  if (activeModal === D.modalOverlay) closeModal();
  else if (activeModal === D.catModalOverlay) closeCatModal();
  else if (activeModal === D.alphaModalOverlay) closeAlphaModal();
  else if (activeModal === D.numMenuOverlay)   closeNumMenu();
}

// ── Esportazione CSV / vCard ─────────────────────────────────────────────────
// Risponde vero se il file è stato consegnato al browser (che poi lo salvi
// davvero, da qui non si può sapere).
function downloadFile(filename, content, mime) {
  try {
    const blob = new Blob([content], { type: mime });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = filename;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
    return true;
  } catch (_) {
    return false;
  }
}
// Un foglio di calcolo tratta da FORMULA una cella che comincia con = + - @
// (o con una tabulazione o un ritorno a capo): l'apice davanti la rende testo.
// L'apice va davanti a OGNI possibile inizio di cella, non solo al primo
// carattere del campo: un programma che divide il file su un separatore
// diverso da quello usato qui vedrebbe come cella a sé anche ciò che segue un
// «;», una «,», una tabulazione o un ritorno a capo dentro il testo.
function csvEscape(s) {
  s = String(s == null ? '' : s);
  s = s.replace(/(^|[;,\t\r\n])([=+\-@])/g, "$1'$2");
  if (/^[\t\r\n]/.test(s)) s = "'" + s;
  if (/[",;\t\n\r]/.test(s)) return '"' + s.replace(/"/g, '""') + '"';
  return s;
}
// Il file: separatore «;» e segno BOM in testa, la forma che Excel in italiano
// apre già in colonne (e in cui rispetta le virgolette dei campi).
const CSV_SEPARATORE = ';';
function buildCSV(contatti) {
  const rows = [['nome', 'categoria', 'numeri', 'note'].join(CSV_SEPARATORE)];
  for (const c of contatti) {
    rows.push([c.nome, c.categoria, c.numeri, c.note].map(csvEscape).join(CSV_SEPARATORE));
  }
  return '\uFEFF' + rows.join('\r\n') + '\r\n';
}
function exportCSV() {
  if (!datiPronti || !allContacts.length) { showToast('Nessun contatto da esportare', 3000, 'warning'); return; }
  const today = new Date().toISOString().slice(0, 10);
  if (downloadFile(`rubrica-${today}.csv`, buildCSV(allContacts), 'text/csv;charset=utf-8')) showToast('CSV esportato', 2000, 'success');
  else showToast('Esportazione non riuscita', 3500, 'error');
}
// Testo dentro una scheda vCard: un ritorno a capo aggiungerebbe righe (altri
// numeri, altri indirizzi) alla scheda che finisce nel telefono.
function vcTesto(s) {
  return String(s == null ? '' : s)
    .replace(/\\/g, '\\\\')
    .replace(/\r\n|\r|\n/g, '\\n')
    .replace(/[,;]/g, x => '\\' + x)
    .replace(/[\u0000-\u001f\u007f]/g, '');
}
function buildVCard(c) {
  const nums  = String(c.numeri || '').split('|').map(s => s.trim()).filter(Boolean);
  const notes = String(c.note   || '').split('|');
  const lines = [
    'BEGIN:VCARD',
    'VERSION:3.0',
    `FN:${vcTesto(c.nome)}`,
    `N:${vcTesto(c.nome)};;;;`,
    `ORG:${vcTesto(c.categoria)}`,
  ];
  nums.forEach((n, i) => {
    if (!NUMERO_TEL.test(n)) return;   // nella scheda entrano solo numeri veri
    // La nota va nel TYPE, così nel telefono compare come etichetta: solo lettere, cifre e spazi
    const nota = (notes[i] || '').replace(/[^A-Za-zÀ-ÿ0-9 ]/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 40);
    lines.push(`TEL;TYPE=${nota ? 'WORK,' + nota : 'WORK'}:${n}`);
  });
  lines.push('END:VCARD');
  return lines.join('\r\n');
}

function exportSingleVCF(contact) {
  if (!contact) return;
  const slug = String(contact.nome).toLowerCase()
    .replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '').slice(0, 40) || 'contatto';
  if (!downloadFile(`${slug}.vcf`, buildVCard(contact), 'text/vcard;charset=utf-8')) showToast('Esportazione non riuscita', 3500, 'error');
  haptic(12);
}

// ── Drag-to-close modal (handle bar) ─────────────────────────────────────────
function setupDragHandles() {
  document.querySelectorAll('.modal-handle').forEach(handle => {
    let startY = 0, currentY = 0, dragging = false, pointerId = null;
    const box = handle.parentElement;
    const overlay = box.parentElement;

    handle.addEventListener('pointerdown', e => {
      dragging  = true;
      pointerId = e.pointerId;
      startY    = e.clientY;
      currentY  = e.clientY;
      box.style.transition = 'none';
      try { handle.setPointerCapture(e.pointerId); } catch (_) {}
    });
    handle.addEventListener('pointermove', e => {
      if (!dragging || e.pointerId !== pointerId) return;
      currentY = e.clientY;
      const dy = Math.max(0, currentY - startY);
      box.style.transform = `translateY(${dy}px)`;
      e.preventDefault();
    });
    const finish = e => {
      if (!dragging) return;
      dragging = false;
      try { handle.releasePointerCapture(pointerId); } catch (_) {}
      pointerId = null;
      const dy = currentY - startY;
      box.style.transition = 'transform .22s ease-out';
      if (dy > box.offsetHeight * 0.28) {
        box.style.transform = `translateY(${box.offsetHeight}px)`;
        setTimeout(() => {
          box.style.transform = '';
          if (overlay === D.modalOverlay)         closeModal();
          else if (overlay === D.catModalOverlay) closeCatModal();
          else                                     overlay.hidden = true;
        }, 220);
      } else {
        box.style.transform = '';
      }
    };
    handle.addEventListener('pointerup',     finish);
    handle.addEventListener('pointercancel', finish);
    handle.addEventListener('pointerleave',  finish);
  });
}

// ── Righe dal database: forma verificata, ordinamento, testo per la ricerca ──
// Di ogni riga si tiene solo ciò che ha la forma attesa: identificativo intero
// e campi di testo. Il campo `_search` è calcolato una volta sola.
function soloTesto(v, massimo) {
  const s = typeof v === 'string' ? v : (typeof v === 'number' ? String(v) : '');
  return s.slice(0, massimo);
}
function prepareContacts(arr) {
  const puliti = [];
  for (const r of (Array.isArray(arr) ? arr : [])) {
    if (!r || typeof r !== 'object') continue;
    const id = idIntero(r.id);
    if (id === null) continue;
    puliti.push({
      id,
      nome:      soloTesto(r.nome, 200),
      categoria: soloTesto(r.categoria, 80),
      numeri:    soloTesto(r.numeri, 300),
      note:      soloTesto(r.note, 600),
    });
  }
  puliti.sort((a, b) => {
    const aNum = /^\d/.test(a.nome);
    const bNum = /^\d/.test(b.nome);
    if (aNum && !bNum) return -1;
    if (!aNum && bNum) return 1;
    return a.nome.localeCompare(b.nome, 'it');
  });
  for (const c of puliti) {
    c._search = (c.nome + ' ' + c.numeri + ' ' + c.note).toLowerCase();
  }
  return puliti;
}
function prepareCategories(arr) {
  const pulite = [];
  for (const r of (Array.isArray(arr) ? arr : [])) {
    if (!r || typeof r !== 'object') continue;
    const id   = idIntero(r.id);
    const nome = soloTesto(r.nome, 80);
    if (id === null || !nome) continue;
    const ordine = Number(r.ordine);
    pulite.push({ id, nome, ordine: Number.isFinite(ordine) ? ordine : 0 });
  }
  pulite.sort((a, b) => (a.ordine - b.ordine) || (a.id - b.id));
  return pulite;
}

// ── Avvio ────────────────────────────────────────────────────────────────────
// La pagina nasce chiusa, col solo messaggio fisso. Si apre solo se è dentro
// l'app E l'app ha una sessione: prima di allora non parte nessuna richiesta
// al database e non si legge nulla dalla memoria del browser.
let avviata = false;
let avvioInCorso = false;
async function avvia() {
  if (avviata || avvioInCorso) return;
  const m = madre();
  if (!m) return;                    // aperta da sola, o dentro una pagina che non è l'app
  avvioInCorso = true;
  let conSessione = false;
  try { await gettone(m, false); conSessione = true; } catch (_) {}
  avvioInCorso = false;
  if (!conSessione) return;          // app senza accesso: resta il messaggio fisso
  avviata = true;
  $('avvisoFuori').hidden = true;
  $('testata').hidden = false;
  $('corpo').hidden = false;
  init();
}
function init() {
  cacheDOM();
  pulisciResidui();
  riprendiPreferenzeDiPrima();
  favs = getFavs();
  applyTheme(temaSalvato(), false);
  setupEvents();
  setupDragHandles();
  loadData(false);
}
// La pagina madre la chiama a ogni riapertura del riquadro: una lettura minima
// dice se qualcuno ha modificato la rubrica, e solo allora la si rilegge.
function rubricaAllApertura() {
  if (!avviata) { avvia(); return; }
  loadData(false);
}
// La pagina madre lo chiede prima di chiudere la sessione («Esci»): un
// salvataggio in corso, un contatto scritto e non salvato, o modifiche alle
// categorie non salvate, andrebbero persi senza avviso.
function rubricaInSospeso() {
  if (!avviata) return false;
  if (scrittureInCorso > 0) return true;
  if (!D.modalOverlay.hidden && fotoModulo() !== moduloAllApertura) return true;
  if (!D.catModalOverlay.hidden && hasCatChanges()) return true;
  return false;
}
// Torna al messaggio fisso (la pagina madre non c'è più).
function mostraFuori() {
  $('testata').hidden = true;
  $('corpo').hidden = true;
  $('avvisoFuori').hidden = false;
}

// ── Caricamento dei dati ─────────────────────────────────────────────────────
// Un caricamento alla volta, in fila. Senza «forza» si legge solo
// rubrica_versione (una riga): se il valore è quello dei dati in memoria non
// si scarica altro. Dopo una propria modifica si rilegge sempre tutto (forza).
let codaCaricamenti = Promise.resolve();
function loadData(forza = false) {
  const p = codaCaricamenti.then(() => caricaDati(forza));
  codaCaricamenti = p.catch(() => {});
  return p;
}
async function caricaDati(forza) {
  try {
    if (!datiPronti) { showLoading(); setLoadingStatus('Verifica ultimo aggiornamento...'); }
    const righe = await supaFetch('rubrica_versione?select=ts&id=eq.1');
    // Un account collegato ma non autorizzato non riceve un errore: riceve un
    // elenco vuoto (lo decidono le regole del database). Lo stesso elenco
    // vuoto arriva se la riga della versione non è stata ancora creata: da qui
    // i due casi non si distinguono, e il messaggio li nomina tutti e due.
    if (!Array.isArray(righe) || !righe.length) throw new ErroreSessione('non-visibile');
    const ts = Number(righe[0] && righe[0].ts);
    if (!Number.isFinite(ts)) throw erroreChiaro('Risposta del database non leggibile');
    if (!forza && datiPronti && ts === versioneNota) return;

    setLoadingStatus('Caricamento contatti...');
    const [contatti, cats] = await Promise.all([
      leggiTutte('rubrica_contatti?select=id,nome,categoria,numeri,note&order=id'),
      leggiTutte('rubrica_categorie?select=id,nome,ordine&order=id'),
    ]);
    allContacts = prepareContacts(contatti);
    categories  = prepareCategories(cats);
    // Il valore letto PRIMA dello scarico: una modifica arrivata nel frattempo
    // farà risultare i dati vecchi al prossimo controllo, e verranno riletti.
    versioneNota = ts;
    datiPronti   = true;

    // un filtro su una categoria che non esiste più non deve restare acceso
    if (activeCategory && !categories.some(c => c.nome === activeCategory)) {
      activeCategory = null;
      D.btnAll.classList.toggle('active', !activeSpecial && !activeAlpha);
    }
    renderChips();
    applyFilters();
  } catch (e) {
    // Un controllo di passaggio fallito per la rete non butta via ciò che si
    // sta guardando; tutto il resto (sessione, primo carico, rilettura dopo una
    // modifica) si dice.
    if (!(e instanceof ErroreSessione) && !forza && datiPronti) return;
    mostraErrore(e);
  }
}
// Messaggio al posto dell'elenco, col pulsante «Riprova». I dati in memoria si
// buttano: senza sessione non devono restare, e dopo un errore non sono più
// affidabili.
function mostraErrore(e) {
  allContacts = [];
  categories  = [];
  versioneNota = null;
  datiPronti   = false;
  if (e instanceof ErroreSessione && e.motivo === 'fuori-app') { mostraFuori(); return; }
  const sotto = consiglioErrore(e);
  D.chips.innerHTML = '';
  D.contactList.innerHTML =
    `<div class="state-msg"><div class="ico">×</div><div>${esc(messaggioErrore(e))}</div><div class="sub">${esc(sotto)}</div><button type="button" class="btn-riprova" id="btnRiprova">Riprova</button></div>`;
  updateCount(null);
}

// ── Render Chips (categorie + speciali) ─────────────────────────────────────
function renderChips() {
  const sorted = [...categories].sort((a, b) => a.nome.localeCompare(b.nome, 'it'));
  const favCount = favs.size;
  const recentCount = getRecentCalls().length;

  let html = '';
  if (favCount) {
    html += `<button class="chip chip-special${activeSpecial === 'fav' ? ' active' : ''}" data-special="fav" type="button">★ Preferiti</button>`;
  }
  if (recentCount) {
    html += `<button class="chip chip-special${activeSpecial === 'recent' ? ' active' : ''}" data-special="recent" type="button">🕐 Recenti</button>`;
  }
  html += sorted.map(c =>
    `<button class="chip${activeCategory === c.nome ? ' active' : ''}" data-cat="${esc(c.nome)}" type="button">${esc(c.nome)}</button>`
  ).join('');

  D.chips.innerHTML = html;
}

// ── Filtri ───────────────────────────────────────────────────────────────────
function applyFilters() {
  let result = allContacts;
  if (activeSpecial === 'fav') {
    result = result.filter(c => favs.has(c.id));
  } else if (activeSpecial === 'recent') {
    const recents = getRecentCalls();
    const order = new Map(recents.map((id, i) => [id, i]));
    result = result
      .filter(c => order.has(c.id))
      .sort((a, b) => order.get(a.id) - order.get(b.id));
  }
  if (activeCategory) result = result.filter(c => c.categoria === activeCategory);
  if (activeAlpha) {
    if (activeAlpha === '0-9') {
      result = result.filter(c => /^\d/.test(c.nome));
    } else {
      result = result.filter(c => c.nome.toUpperCase().startsWith(activeAlpha));
    }
  }
  if (activeSearch) result = smartSearch(result, activeSearch);
  renderContacts(result);
}

// ── Ricerca smart 3-tier ottimizzata ─────────────────────────────────────────
function smartSearch(contacts, rawQuery) {
  const q = rawQuery.toLowerCase().trim();
  if (!q) return contacts;

  const tokens = q.split(/\s+/).filter(Boolean);
  const escRe  = s => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

  // Pre-compile regex tier-2 una sola volta per query (non per contatto)
  const tier2Re = tokens.map(t => new RegExp('\\b' + escRe(t) + '\\b'));
  const doTier3 = q.length >= 4;

  const tier1 = [], tier2 = [], tier3 = [];

  for (const c of contacts) {
    const h = c._search || ''; // memoizzato in prepareContacts
    if (h.includes(q)) {
      tier1.push(c);
    } else if (tier2Re.every(re => re.test(h))) {
      tier2.push(c);
    } else if (doTier3 && tokens.every(t => h.includes(t))) {
      tier3.push(c);
    }
  }
  return [...tier1, ...tier2, ...tier3];
}

// ── Render contatti con avatar + highlight ──────────────────────────────────
// Tutto ciò che viene dai contatti entra in pagina codificato (esc, highlight).
// Gli attributi sono sempre fra virgolette doppie.
function renderContacts(contacts) {
  updateCount(contacts.length);
  if (!contacts.length) {
    D.contactList.innerHTML = '<div class="state-msg"><div class="ico">∅</div><div>Nessun risultato</div></div>';
    return;
  }

  // Tokens per highlighting (solo se c'è ricerca attiva)
  const searchTokens = activeSearch
    ? activeSearch.split(/\s+/).filter(t => t.length >= 2).slice(0, 8)
    : null;

  const html = contacts.map(c => {
    const id    = esc(String(c.id));
    const nomeH = highlight(c.nome, searchTokens);
    const nums  = c.numeri.split('|').map(n => n.trim()).filter(Boolean);
    const notes = c.note.split('|');

    const groups = nums.map((n, i) => {
      const nota = (notes[i] || '').trim();
      const nH   = highlight(n, searchTokens);
      const dati = `data-num="${esc(n)}" data-nome="${esc(c.nome)}"`;
      // collegamento «tel:» solo per un numero fatto di cifre; il resto è testo
      const pill = NUMERO_TEL.test(n)
        ? `<a class="num-pill" href="tel:${esc(n)}" ${dati}>${nH}</a>`
        : `<span class="num-pill num-pill-solo" ${dati}>${nH}</span>`;
      return `<span class="num-group">${pill}${nota ? `<span class="num-nota-inline">${highlight(nota, searchTokens)}</span>` : ''}</span>`;
    }).join('');

    const star = isFav(c.id) ? `<span class="contact-fav-star" aria-label="Preferito">${STAR_SVG}</span>` : '';
    return `<div class="contact-card" data-id="${id}">
      <div class="contact-avatar ${avatarClasse(c.categoria)}" aria-hidden="true">${esc(avatarLetter(c.nome))}</div>
      <div class="contact-info" data-id="${id}">
        <div class="contact-name">${star}<span class="contact-name-text">${nomeH}</span></div>
        <div class="contact-numbers">${groups}</div>
        <div class="contact-cat">${esc(c.categoria)}</div>
      </div>
      <button class="btn-edit" data-id="${id}" aria-label="Modifica contatto" type="button">${EDIT_SVG}</button>
    </div>`;
  }).join('');
  D.contactList.innerHTML = html;
}

function showLoading() {
  D.contactList.innerHTML =
    '<div class="state-msg"><div class="spinner"></div><div id="loadStatus" class="load-status" aria-live="polite"></div></div>';
  updateCount(null);
}
function setLoadingStatus(msg) {
  const el = $('loadStatus');
  if (el) el.textContent = msg;
}
function updateCount(n) {
  D.countLabel.textContent = n !== null ? `${n} contatti` : '';
}

// ── Events setup ─────────────────────────────────────────────────────────────
function setupEvents() {
  // Tema (3 stati: light → dark → auto)
  D.btnTheme.addEventListener('click', cycleTheme);

  // Ciò che si fa qui dentro conta come attività anche per la pagina madre:
  // la sua ricarica notturna aspetta chi sta lavorando.
  const segnalaAttivita = () => {
    try { const m = window.parent; if (m && m !== window) m._lastUserActivityTs = Date.now(); } catch (_) {}
  };
  ['pointerdown', 'keydown'].forEach(ev => document.addEventListener(ev, segnalaAttivita, { passive: true, capture: true }));

  // Search debounced
  const debouncedFilter = debounce(() => {
    applyFilters();
    // Aggiungi alle recenti dopo che l'utente ha smesso di scrivere (se ha trovato qualcosa)
    if (activeSearch && activeSearch.length >= 2) pushRecentSearch(activeSearch);
  }, 400);
  D.searchBar.addEventListener('input', e => {
    activeSearch = e.target.value.toLowerCase().trim();
    D.searchWrap.classList.toggle('has-text', !!activeSearch);
    if (!activeSearch) renderRecentSuggestions();
    else D.searchSuggestions.hidden = true;
    debouncedFilter();
  });
  D.btnSearchClear.addEventListener('click', () => {
    D.searchBar.value = '';
    activeSearch = '';
    D.searchWrap.classList.remove('has-text');
    applyFilters();
    D.searchBar.focus();
  });

  // Mostra Tutti
  D.btnAll.addEventListener('click', () => {
    activeCategory = null;
    activeSearch   = '';
    activeAlpha    = null;
    activeSpecial  = null;
    D.searchBar.value = '';
    D.searchWrap.classList.remove('has-text');
    renderChips();
    D.btnAll.classList.add('active');
    D.btnAlpha.classList.remove('active');
    renderContacts(allContacts);
  });

  // Alpha
  D.btnAlpha.addEventListener('click', openAlphaModal);
  D.btnAlphaClose.addEventListener('click', closeAlphaModal);
  D.alphaModalOverlay.addEventListener('click', e => {
    if (e.target === D.alphaModalOverlay) closeAlphaModal();
  });

  // Nuovo
  D.btnNew.addEventListener('click', nuovoContatto);
  D.btnAddNum.addEventListener('click', () => { const f = addNumRow(); f.focus(); });
  D.btnCancel.addEventListener('click', closeModal);
  D.modalOverlay.addEventListener('click', e => {
    if (e.target === D.modalOverlay) closeModal();
  });

  // Categorie
  D.btnSettings.addEventListener('click', openCatModal);
  D.btnCatCancel.addEventListener('click', closeCatModal);
  D.catModalOverlay.addEventListener('click', e => {
    if (e.target === D.catModalOverlay) closeCatModal();
  });
  D.btnAddCat.addEventListener('click', addNewCategory);
  D.fNewCat.addEventListener('keydown', e => {
    if (e.key === 'Enter') { e.preventDefault(); addNewCategory(); }
  });
  D.btnCatSave.addEventListener('click', saveCatChanges);

  // Form submit
  D.contactForm.addEventListener('submit', async e => {
    e.preventDefault();
    await saveContact();
  });
  D.btnDelete.addEventListener('click', async () => {
    // L'identificativo si prende ADESSO: la domanda vale per questo contatto,
    // non per quello che sarà aperto quando arriva la risposta.
    const id = idIntero(D.fId.value);
    if (id === null) { showToast('Contatto non riconosciuto: chiudi la scheda e riaprila', 4000, 'error'); return; }
    if (!(await showConfirm('Eliminare questo contatto?', 'error'))) return;
    if (D.modalOverlay.hidden || idIntero(D.fId.value) !== id) return;
    await deleteContact(id);
  });

  // ── Event delegation: lista contatti ──────────────────────────────────────
  D.contactList.addEventListener('click', e => {
    // 0) «Riprova» sotto un messaggio d'errore
    if (e.target.closest('#btnRiprova')) {
      loadData(true);
      return;
    }
    // 1) Pulsante edit (matita)
    const editBtn = e.target.closest('.btn-edit');
    if (editBtn) {
      e.stopPropagation();
      haptic(8);
      openEdit(editBtn.dataset.id);
      return;
    }
    // 2) Numero — PRIMA di .contact-info perché ne è figlio
    const numLink = e.target.closest('.num-pill');
    if (numLink && !longPressTriggered) {
      haptic(12);
      const card = numLink.closest('.contact-card');
      if (card) {
        const wasEmpty = getRecentCalls().length === 0;
        pushRecentCall(card.dataset.id);
        // Se è la prima chiamata, ridisegna le chips per mostrare "🕐 Recenti"
        if (wasEmpty) renderChips();
      }
      return; // lascia procedere il link tel: nativo
    }
    // 3) Area info (nome) → apre modal modifica
    const info = e.target.closest('.contact-info');
    if (info && !longPressTriggered) {
      haptic(8);
      openEdit(info.dataset.id);
    }
  });

  // ── Long-press su numero → menu copia/condividi ──────────────────────────
  setupLongPress();

  // ── Event delegation: chips categorie + speciali ─────────────────────────
  D.chips.addEventListener('click', e => {
    const chip = e.target.closest('.chip');
    if (!chip) return;
    haptic(8);
    if (chip.dataset.special) {
      const sp = chip.dataset.special;
      activeSpecial = activeSpecial === sp ? null : sp;
      activeCategory = null;
      renderChips();
      D.btnAll.classList.toggle('active', !activeSpecial);
      applyFilters();
      return;
    }
    const cat = chip.dataset.cat;
    activeCategory = activeCategory === cat ? null : cat;
    activeSpecial = null;
    renderChips();
    D.btnAll.classList.toggle('active', !activeCategory);
    applyFilters();
  });

  // ── Esportazione ──────────────────────────────────────────────────────────
  $('btnExportCSV')?.addEventListener('click', exportCSV);

  // Salva singolo contatto come vCard (solo modal modifica)
  D.btnSaveVcf.addEventListener('click', () => {
    const c = D.btnSaveVcf._contact;
    if (c) exportSingleVCF(c);
  });

  // ── Shortcut tastiera: "\" porta focus sulla ricerca (desktop) ────────────
  document.addEventListener('keydown', e => {
    if (e.key !== '\\') return;
    if (e.ctrlKey || e.metaKey || e.altKey) return;
    const tag = (document.activeElement?.tagName || '').toUpperCase();
    if (tag === 'INPUT' || tag === 'TEXTAREA' || tag === 'SELECT') return;
    // Modal aperto → non interferire
    if (activeModal) return;
    e.preventDefault();
    D.searchBar.focus();
    D.searchBar.select?.();
  });

  // ── Suggerimenti ricerca recenti ──────────────────────────────────────────
  D.searchBar.addEventListener('focus', renderRecentSuggestions);
  D.searchBar.addEventListener('blur',  () => {
    // Delay per permettere il click sul suggerimento
    setTimeout(() => { D.searchSuggestions.hidden = true; }, 150);
  });
  D.searchBar.addEventListener('change', () => {
    if (activeSearch) pushRecentSearch(activeSearch);
  });
  D.searchSuggestions.addEventListener('click', e => {
    const x = e.target.closest('[data-x]');
    if (x) {
      e.stopPropagation();
      removeRecentSearch(x.dataset.x);
      renderRecentSuggestions();
      return;
    }
    const sug = e.target.closest('.search-suggestion');
    if (sug) {
      const q = sug.dataset.q;
      D.searchBar.value = q;
      activeSearch = q.toLowerCase();
      D.searchWrap.classList.add('has-text');
      D.searchSuggestions.hidden = true;
      pushRecentSearch(q);
      applyFilters();
    }
  });

  // ── Menu numero: azioni ───────────────────────────────────────────────────
  D.numMenuOverlay.addEventListener('click', e => {
    if (e.target === D.numMenuOverlay) closeNumMenu();
    const action = e.target.closest('[data-action]');
    if (!action) return;
    handleNumMenuAction(action.dataset.action);
  });
}

// ── Long press detection (numero | card) ─────────────────────────────────────
let longPressTimer = null;
let longPressTriggered = false;
let longPressNum  = '';
let longPressNome = '';
let longPressMode = '';   // 'num' | 'card'
let longPressId   = '';
function setupLongPress() {
  D.contactList.addEventListener('pointerdown', e => {
    const a    = e.target.closest('.num-pill');
    const card = e.target.closest('.contact-info');
    if (!a && !card) return;
    longPressTriggered = false;
    if (a) {
      longPressMode = 'num';
      longPressNum  = a.dataset.num  || a.textContent.trim();
      longPressNome = a.dataset.nome || '';
    } else {
      longPressMode = 'card';
      longPressId   = card.dataset.id;
    }
    clearTimeout(longPressTimer);
    longPressTimer = setTimeout(() => {
      longPressTriggered = true;
      haptic([15, 30, 15]);
      if (longPressMode === 'num') openNumMenu(longPressNum, longPressNome);
      else if (longPressMode === 'card') {
        toggleFav(longPressId);
        showToast(isFav(longPressId) ? '★ Aggiunto ai preferiti' : 'Rimosso dai preferiti', 1800, 'success');
        renderChips();
        applyFilters();
      }
    }, 500);
  });
  const cancel = () => { clearTimeout(longPressTimer); };
  D.contactList.addEventListener('pointerup',     cancel);
  D.contactList.addEventListener('pointerleave',  cancel);
  D.contactList.addEventListener('pointercancel', cancel);
  D.contactList.addEventListener('pointermove', e => {
    if (Math.abs(e.movementX) + Math.abs(e.movementY) > 5) cancel();
  });
  // Quando si attiva il long-press, blocca il click successivo
  D.contactList.addEventListener('click', e => {
    if (longPressTriggered) {
      e.preventDefault();
      e.stopPropagation();
      setTimeout(() => { longPressTriggered = false; }, 50);
    }
  }, true);
}

function openNumMenu(num, nome) {
  D.numMenuTitle.textContent = num;
  D.numMenuSub.textContent   = nome || '';
  D.numMenuOverlay.hidden = false;
  trapFocus(D.numMenuOverlay);
}
function closeNumMenu() {
  releaseFocus(D.numMenuOverlay);
  D.numMenuOverlay.hidden = true;
}
const COPIA_NON_RIUSCITA = 'Copia non riuscita: il browser non l\'ha permessa';
async function handleNumMenuAction(action) {
  const num  = D.numMenuTitle.textContent;
  const nome = D.numMenuSub.textContent;
  closeNumMenu();
  haptic(8);
  switch (action) {
    case 'call':
      // la stessa regola del collegamento in elenco: si chiama solo un numero fatto di cifre
      if (NUMERO_TEL.test(num)) window.location.href = `tel:${num}`;
      else showToast('Questo numero non si può chiamare da qui', 3000, 'warning');
      break;
    case 'copy':
      if (await copyText(num)) showToast('Numero copiato', 2000, 'success');
      else showToast(COPIA_NON_RIUSCITA, 3500, 'warning');
      break;
    case 'copy-full':
      if (await copyText(`${nome}: ${num}`)) showToast('Copiato negli appunti', 2000, 'success');
      else showToast(COPIA_NON_RIUSCITA, 3500, 'warning');
      break;
    case 'share':
      if (navigator.share) {
        try {
          await navigator.share({ title: nome, text: `${nome}: ${num}` });
        } catch (_) {}
      } else if (await copyText(`${nome}: ${num}`)) {
        showToast('Condivisione non supportata — copiato', 3000);
      } else {
        showToast('Condivisione non supportata', 3000, 'warning');
      }
      break;
    case 'close':
      break;
  }
}
// Risponde vero solo se il testo è davvero finito negli appunti.
async function copyText(text) {
  if (navigator.clipboard) {
    try { await navigator.clipboard.writeText(text); return true; } catch (_) {}
  }
  // Ripiego
  const ta = document.createElement('textarea');
  ta.value = text;
  ta.style.position = 'fixed';
  ta.style.opacity = '0';
  document.body.appendChild(ta);
  ta.select();
  let riuscita = false;
  try { riuscita = document.execCommand('copy') === true; } catch (_) {}
  ta.remove();
  return riuscita;
}

// ── Modal contatto ───────────────────────────────────────────────────────────
function addNumRow(num = '', nota = '') {
  const row = document.createElement('div');
  row.className = 'num-row';
  row.innerHTML = `
    <div class="num-row-inputs">
      <input type="tel" class="f-num" placeholder="Numero" value="${esc(num)}" inputmode="numeric" autocomplete="off" maxlength="20">
      <input type="text" class="f-nota" placeholder="Nota (opzionale)" value="${esc(nota)}" autocomplete="off" maxlength="120">
    </div>
    <button type="button" class="btn-rm-num" aria-label="Rimuovi riga">×</button>`;

  const fNum = row.querySelector('.f-num');
  fNum.addEventListener('input', () => {
    fNum.value = fNum.value.replace(/[^\d]/g, '');
    fNum.classList.toggle('invalid', fNum.value.length > 0 && !/^\d+$/.test(fNum.value));
  });
  fNum.addEventListener('blur', () => {
    fNum.classList.toggle('invalid', fNum.value.length > 0 && !/^\d+$/.test(fNum.value));
  });

  const fNota = row.querySelector('.f-nota');
  fNota.addEventListener('input', () => {
    if (fNota.value.includes('|')) fNota.value = fNota.value.replace(/\|/g, '');
  });

  row.querySelector('.btn-rm-num').addEventListener('click', () => {
    if (D.numeriContainer.querySelectorAll('.num-row').length > 1) row.remove();
  });

  D.numeriContainer.appendChild(row);
  return fNum;
}

function collectPairs() {
  const rows = D.numeriContainer.querySelectorAll('.num-row');
  const numeri = [], note = [];
  let valid = true;
  rows.forEach(row => {
    const n = row.querySelector('.f-num').value.trim();
    const t = row.querySelector('.f-nota').value.replace(/\|/g, '').trim();
    if (!n) return;
    if (!/^\d+$/.test(n)) {
      row.querySelector('.f-num').classList.add('invalid');
      valid = false;
    } else {
      numeri.push(n);
      note.push(t);
    }
  });
  if (!valid) return null;
  return { numeri: numeri.join('|'), note: note.join('|') };
}

// La copia del contatto su cui la scheda è stata aperta (null = contatto
// nuovo): al salvataggio la riga viene riletta e confrontata con questa.
let contattoAperto    = null;
// I valori del modulo appena aperto: dicono se c'è qualcosa di scritto e non
// salvato (rubricaInSospeso).
let moduloAllApertura = '';
function fotoModulo() {
  const righe = [...D.numeriContainer.querySelectorAll('.num-row')]
    .map(r => [r.querySelector('.f-num').value, r.querySelector('.f-nota').value]);
  return JSON.stringify([D.fNome.value, D.fCategoria.value, righe]);
}
function stessiDati(a, b) {
  return a.nome === b.nome && a.categoria === b.categoria && a.numeri === b.numeri && a.note === b.note;
}

function openModal(contact) {
  // Salva scroll position per ripristinarla alla chiusura
  savedScrollY = window.scrollY;

  const isNew = !contact;
  contattoAperto = contact
    ? { id: contact.id, nome: contact.nome, categoria: contact.categoria, numeri: contact.numeri, note: contact.note }
    : null;
  invioSenzaRisposta = null;
  D.modalTitle.textContent = isNew ? 'Nuovo Contatto' : 'Modifica Contatto';
  D.fId.value   = contact ? String(contact.id) : '';
  D.fNome.value = contact ? contact.nome : '';
  D.btnDelete.hidden = isNew;
  D.btnSaveVcf.hidden = isNew;
  D.btnSaveVcf._contact = contact || null;
  // La categoria del contatto può non essere più in elenco: resta fra le
  // scelte, così salvando non cambia da sola.
  const nomiCat = categories.map(c => c.nome);
  if (contact && contact.categoria && !nomiCat.includes(contact.categoria)) nomiCat.push(contact.categoria);
  // Un contatto SENZA categoria: la scelta vuota è la sua, così salvando non
  // gliene viene assegnata una di nascosto (saveContact chiede di sceglierla).
  const senzaCategoria = !!contact && !contact.categoria;
  D.fCategoria.innerHTML =
    (senzaCategoria ? '<option value="" selected>— scegli una categoria —</option>' : '') +
    nomiCat.map(n =>
      `<option value="${esc(n)}"${contact && n === contact.categoria ? ' selected' : ''}>${esc(n)}</option>`
    ).join('');

  D.numeriContainer.innerHTML = '';
  const nums  = String(contact?.numeri || '').split('|').map(s => s.trim()).filter(Boolean);
  const notes = String(contact?.note   || '').split('|');
  if (nums.length) {
    nums.forEach((n, i) => addNumRow(n, notes[i] || ''));
  } else {
    addNumRow();
  }

  moduloAllApertura = fotoModulo();
  D.modalOverlay.hidden = false;
  trapFocus(D.modalOverlay);
  setTimeout(() => { if (document.hasFocus()) D.fNome.focus(); }, 50);
}
// «Nuovo Contatto» vuole l'elenco caricato: senza, non ci sono le categorie
// fra cui scegliere e il contatto nascerebbe senza categoria.
function nuovoContatto() {
  if (!datiPronti) { showToast('La rubrica non è ancora caricata: attendi, oppure premi «Riprova»', 3500, 'warning'); return; }
  openModal(null);
}
// Prima di aprire la scheda si ricontrolla, con la lettura minima della
// versione, se un altro PC ha modificato la rubrica: così la scheda si apre
// sui dati di adesso e non su una copia rimasta in memoria da ore. Se il
// controllo tarda (rete lenta) la scheda si apre lo stesso: al salvataggio la
// riga viene comunque riletta e confrontata (saveContact).
const ATTESA_APERTURA_MS = 2500;
let aperturaInCorso = false;
async function openEdit(id) {
  id = idIntero(id);
  if (id === null || aperturaInCorso || !D.modalOverlay.hidden) return;
  aperturaInCorso = true;
  try {
    await Promise.race([loadData(false), sleep(ATTESA_APERTURA_MS)]);
  } finally {
    aperturaInCorso = false;
  }
  if (!datiPronti) return;            // l'errore è già al posto dell'elenco
  if (activeModal) return;            // durante l'attesa si è aperta un'altra finestra: non la si copre
  const c = allContacts.find(x => x.id === id);
  if (!c) { showToast('Il contatto non esiste più: forse è stato eliminato da un altro PC', 4000, 'warning'); return; }
  openModal(c);
}
function closeModal() {
  // una domanda rimasta a schermo («Eliminare questo contatto?») non sopravvive alla scheda
  annullaConferma();
  invioSenzaRisposta = null;
  contattoAperto = null;
  releaseFocus(D.modalOverlay);
  D.modalOverlay.hidden = true;
  // Ripristina scroll position
  if (savedScrollY) {
    requestAnimationFrame(() => window.scrollTo(0, savedScrollY));
  }
}

// ── Salvataggio ed eliminazione di un contatto ───────────────────────────────
// Dopo una modifica riuscita non si ricostruisce nulla a mano: si rilegge la
// rubrica dal database (così si vedono anche le modifiche fatte da altri PC).
// Modifica ed eliminazione chiedono indietro la riga toccata: se non ne torna
// nessuna il contatto non c'è più, e lo si dice invece di fingere un successo.
//
// Due regole in più:
//  - AGGIORNAMENTO PERSO. Una modifica riscrive tutti i campi del contatto:
//    prima di inviarla si rilegge la riga, e si scrive solo se è ancora quella
//    su cui la scheda è stata aperta. Se un altro PC l'ha cambiata nel
//    frattempo non si scrive nulla e lo si dice.
//  - ESITO INCERTO. Se la risposta di una scrittura non arriva (rete caduta,
//    tempo massimo) la richiesta può essere arrivata lo stesso: non si invita
//    a riprovare alla cieca. Si rilegge la rubrica e si guarda se la scrittura
//    c'è già; se nemmeno la rilettura riesce lo si ricorda
//    (invioSenzaRisposta) e il controllo si rifà al prossimo «Salva», PRIMA di
//    inviare di nuovo.
let scrittureInCorso   = 0;      // salvataggi partiti e non ancora conclusi (anche delle categorie)
let invioSenzaRisposta = null;   // { nuovo, id, payload, notiPrima } dell'invio rimasto senza risposta
const AVVISO_ESITO_IGNOTO = 'Connessione assente: non è certo che il contatto sia stato salvato. Quando la rete torna premi di nuovo «Salva»: prima di inviare si controlla se c\'è già';

// Risponde 'salvato', 'non-salvato' oppure 'ignoto' (nemmeno la rilettura è
// riuscita). Un contatto nuovo si riconosce dagli stessi dati su un
// identificativo che prima dell'invio non c'era.
async function verificaInvio(v) {
  await loadData(true);
  if (!datiPronti) return 'ignoto';
  const trovato = v.nuovo
    ? allContacts.some(c => !v.notiPrima.has(c.id) && stessiDati(c, v.payload))
    : allContacts.some(c => c.id === v.id && stessiDati(c, v.payload));
  return trovato ? 'salvato' : 'non-salvato';
}
function contattoSparito() {
  const e = erroreChiaro('Il contatto non esiste più: forse è stato eliminato da un altro PC');
  e.sparito = true;
  return e;
}

async function saveContact() {
  if (scrittureInCorso) return;        // un salvataggio è già partito
  const nuovo = D.fId.value === '';
  const id    = nuovo ? null : idIntero(D.fId.value);
  if (!nuovo && id === null) { showToast('Contatto non riconosciuto: chiudi la scheda e riaprila', 4000, 'error'); return; }
  const nome = D.fNome.value.trim();
  if (!nome) { showToast('Il nome è obbligatorio', 3000, 'warning'); return; }

  const pairs = collectPairs();
  if (!pairs) { showToast('Uno o più numeri contengono caratteri non validi (solo cifre)', 3500, 'warning'); return; }
  if (!pairs.numeri) { showToast('Inserisci almeno un numero', 3000, 'warning'); return; }

  const payload = {
    nome,
    categoria: D.fCategoria.value,
    numeri:    pairs.numeri,
    note:      pairs.note,
  };
  // la categoria è obbligatoria, finché ne esiste almeno una fra cui scegliere
  if (!payload.categoria && categories.length) { showToast('Scegli una categoria', 3000, 'warning'); return; }
  // gli stessi vincoli del database, detti prima di inviare
  const rifiuto = motivoRifiuto(payload);
  if (rifiuto) { showToast(rifiuto, 4000, 'warning'); return; }

  const btn = D.btnSave;
  btn.innerHTML = '<span class="btn-spinner"></span>'; btn.disabled = true;
  scrittureInCorso++;
  // gli identificativi noti prima dell'invio: dopo una risposta persa dicono
  // qual è il contatto appena creato
  const invio = { nuovo, id, payload, notiPrima: new Set(allContacts.map(c => c.id)) };
  const riuscito = nuovo ? 'Contatto aggiunto con successo' : 'Contatto aggiornato con successo';
  try {
    if (invioSenzaRisposta) {
      // l'invio di prima è rimasto senza risposta e non si era potuto
      // controllare: lo si controlla adesso, prima di inviare di nuovo
      const prima = await verificaInvio(invioSenzaRisposta);
      if (prima === 'ignoto') { showToast(AVVISO_ESITO_IGNOTO, 7000, 'error'); return; }
      invioSenzaRisposta = null;
      if (prima === 'salvato') {
        closeModal();
        showToast('Il contatto risulta già salvato: controllalo nell\'elenco', 4500, 'success');
        return;
      }
      invio.notiPrima = new Set(allContacts.map(c => c.id));
    }
    if (!nuovo) {
      const adesso = prepareContacts(await supaFetch(`rubrica_contatti?select=id,nome,categoria,numeri,note&id=eq.${id}`))[0];
      if (!adesso) throw contattoSparito();
      if (contattoAperto && !stessiDati(adesso, contattoAperto) && !stessiDati(adesso, payload)) {
        closeModal();
        showToast('Questo contatto è stato modificato da un altro PC mentre la scheda era aperta: riaprilo e ripeti la modifica', 6500, 'warning');
        loadData(true);
        return;
      }
    }
    const rows = await supaFetch(nuovo ? 'rubrica_contatti' : `rubrica_contatti?id=eq.${id}`, {
      method: nuovo ? 'POST' : 'PATCH',
      headers: { 'Prefer': 'return=representation' },
      body: JSON.stringify(payload),
    });
    if (!Array.isArray(rows) || !rows.length) {
      if (!nuovo) throw contattoSparito();
      const e = erroreChiaro('Il contatto non risulta salvato: riprova');
      e.incerto = true;
      throw e;
    }
    closeModal();
    showToast(riuscito, 3000, 'success');
    loadData(true);
  } catch (e) {
    if (e && e.sparito) {
      // il contatto non c'è più: la scheda si chiude e l'elenco si rilegge
      closeModal();
      showToast(messaggioErrore(e), 4500, 'warning');
      loadData(true);
    } else if (e && (e.rete || e.incerto)) {
      const esito = await verificaInvio(invio);
      if (esito === 'salvato') {
        closeModal();
        showToast(riuscito, 3000, 'success');
      } else if (esito === 'ignoto') {
        invioSenzaRisposta = invio;
        showToast(AVVISO_ESITO_IGNOTO, 7000, 'error');
      } else {
        showToast('Il contatto non risulta salvato: riprova', 4500, 'error');
      }
    } else {
      showToast(messaggioErrore(e), 4500, 'error');
      if (e instanceof ErroreSessione) loadData(true);
    }
  } finally {
    scrittureInCorso--;
    btn.textContent = 'Salva'; btn.disabled = false;
  }
}

// «id» è quello preso al clic su «Elimina» (vedi il gestore del pulsante).
async function deleteContact(id) {
  id = idIntero(id);
  if (id === null) { showToast('Contatto non riconosciuto: chiudi la scheda e riaprila', 4000, 'error'); return; }
  if (scrittureInCorso) return;
  const btn = D.btnDelete;
  btn.innerHTML = '<span class="btn-spinner"></span>'; btn.disabled = true;
  scrittureInCorso++;
  try {
    const rows = await supaFetch(`rubrica_contatti?id=eq.${id}`, {
      method: 'DELETE',
      headers: { 'Prefer': 'return=representation' },
    });
    closeModal();
    showToast(Array.isArray(rows) && rows.length ? 'Contatto eliminato con successo' : 'Il contatto era già stato eliminato', 3000, 'success');
    loadData(true);
  } catch (e) {
    if (e && (e.rete || e.incerto)) {
      // risposta persa: si rilegge, e se il contatto non c'è più l'eliminazione era arrivata
      await loadData(true);
      if (datiPronti && !allContacts.some(c => c.id === id)) {
        closeModal();
        showToast('Contatto eliminato con successo', 3000, 'success');
      } else {
        showToast(datiPronti ? 'Il contatto non risulta eliminato: riprova' : messaggioErrore(e), 4500, 'error');
      }
    } else {
      showToast(messaggioErrore(e), 4500, 'error');
      if (e instanceof ErroreSessione) loadData(true);
    }
  } finally {
    scrittureInCorso--;
    btn.textContent = 'Elimina'; btn.disabled = false;
  }
}

// ── Modal Alpha ──────────────────────────────────────────────────────────────
function openAlphaModal() {
  D.alphaGrid.innerHTML = ALPHA_KEYS.map(k =>
    `<button class="alpha-btn${k === '0-9' ? ' num-btn' : ''}${activeAlpha === k ? ' active' : ''}"
             data-key="${k}" type="button">${k}</button>`
  ).join('');
  D.alphaGrid.querySelectorAll('.alpha-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const key = btn.dataset.key;
      if (activeAlpha === key) {
        activeAlpha = null;
        D.btnAlpha.classList.remove('active');
        D.btnAll.classList.add('active');
      } else {
        activeAlpha = key;
        D.btnAlpha.classList.add('active');
        D.btnAll.classList.remove('active');
      }
      closeAlphaModal();
      applyFilters();
    });
  });
  D.alphaModalOverlay.hidden = false;
  trapFocus(D.alphaModalOverlay);
}

// ── Modal Categorie ──────────────────────────────────────────────────────────
let catOriginal   = [];
let catPending    = [];
let catPendingNew = [];

// Prima di mostrare le categorie si ricontrolla se un altro PC ha modificato
// la rubrica (una lettura minima); i conteggi si fanno sui contatti in memoria.
// «pos» è il posto della categoria nella finestra (0, 1, 2…) e «pos0» quello
// che aveva all'apertura: il trascinamento scambia i POSTI, non i valori di
// «ordine», così anche due categorie rimaste con lo stesso «ordine» si
// possono separare (vedi nuoviOrdini).
async function openCatModal() {
  if (!D.catModalOverlay.hidden) return;
  D.catModalOverlay.hidden = false;
  trapFocus(D.catModalOverlay);
  D.catList.innerHTML = '<div class="cat-loading"><div class="spinner"></div></div>';
  catOriginal = []; catPending = []; catPendingNew = [];
  await loadData(false);
  if (D.catModalOverlay.hidden) return;      // chiusa durante l'attesa
  if (!datiPronti) {
    chiudiCategorie();
    return;
  }
  catOriginal = categories.map(c => ({
    ...c,
    count: allContacts.filter(x => x.categoria === c.nome).length,
  }));
  catPending    = catOriginal.map((c, i) => ({ ...c, editNome: c.nome, deleted: false, pos: i, pos0: i }));
  catPendingNew = [];
  renderCatList();
}

function chiudiCategorie() {
  annullaConferma();     // una domanda rimasta a schermo non sopravvive alla finestra
  releaseFocus(D.catModalOverlay);
  D.catModalOverlay.hidden = true;
}
async function closeCatModal() {
  if (hasCatChanges() && !(await showConfirm('Ci sono modifiche non salvate.\nChiudere ugualmente?', 'warning'))) return;
  chiudiCategorie();
}

function hasCatChanges() {
  // Cambio nome o eliminazione
  if (catPending.some(c => c.deleted || c.editNome !== c.nome)) return true;
  // Aggiunte
  if (catPendingNew.length > 0) return true;
  // Cambio di posto
  return catPending.some(c => c.pos !== c.pos0);
}
// Scambia il posto di due categorie nella finestra e ridisegna l'elenco.
function scambiaPosti(pidxA, pidxB) {
  const a = catPending[pidxA];
  const b = catPending[pidxB];
  if (!a || !b || a === b || a.deleted || b.deleted) return false;
  const tmp = a.pos;
  a.pos = b.pos;
  b.pos = tmp;
  renderCatList();
  return true;
}
// I valori di «ordine» da scrivere dopo un riordino: { id, ordine } delle sole
// categorie il cui valore cambia. Si riusano i valori di prima, messi in fila
// e resi strettamente crescenti: così due categorie scambiate sono due
// scritture (le altre non si toccano), e due valori uguali rimasti da un
// salvataggio interrotto a metà vengono separati.
function nuoviOrdini() {
  const inFila = catPending.filter(c => !c.deleted).sort((a, b) => a.pos - b.pos);
  if (!inFila.some(c => c.pos !== c.pos0)) return [];     // nessuno spostamento: non si scrive nulla
  const valori = inFila.map(c => Number(c.ordine) || 0).sort((a, b) => a - b);
  for (let i = 1; i < valori.length; i++) {
    if (valori[i] <= valori[i - 1]) valori[i] = valori[i - 1] + 1;
  }
  return inFila
    .map((c, i) => ({ id: c.id, ordine: valori[i], prima: Number(c.ordine) || 0 }))
    .filter(x => x.ordine !== x.prima)
    .map(x => ({ id: x.id, ordine: x.ordine }));
}

// ── Drag & drop riordino categorie ──────────────────────────────────────────
function setupCatDrag() {
  let dragRow = null;
  let dragPidx = null;
  let pointerId = null;

  D.catList.querySelectorAll('.cat-drag-handle[data-drag]').forEach(handle => {
    handle.addEventListener('pointerdown', e => {
      const row = handle.closest('.cat-item');
      if (!row) return;
      dragRow  = row;
      dragPidx = Number(row.dataset.pidx);
      pointerId = e.pointerId;
      row.classList.add('dragging');
      try { handle.setPointerCapture(e.pointerId); } catch (_) {}
      haptic(15);
    });

    handle.addEventListener('pointermove', e => {
      if (!dragRow || e.pointerId !== pointerId) return;
      e.preventDefault();
      // Trova la row sotto il puntatore
      const target = document.elementFromPoint(e.clientX, e.clientY)?.closest('.cat-item');
      if (!target || target === dragRow || !target.dataset.pidx) return;
      const targetPidx = Number(target.dataset.pidx);
      // Scambia il posto delle due categorie e ridisegna (la classe «dragging»
      // va poi rimessa sulla riga nuova)
      if (!scambiaPosti(dragPidx, targetPidx)) return;
      // Ritrova la row e marcala dragging
      const newRow = D.catList.querySelector(`.cat-item[data-pidx="${dragPidx}"]`);
      if (newRow) {
        newRow.classList.add('dragging');
        dragRow = newRow;
      }
    });

    const finish = e => {
      if (!dragRow) return;
      dragRow.classList.remove('dragging');
      dragRow = null;
      dragPidx = null;
      pointerId = null;
    };
    handle.addEventListener('pointerup',     finish);
    handle.addEventListener('pointercancel', finish);
    handle.addEventListener('pointerleave',  finish);
  });
}

function renderCatList() {
  // Le categorie nell'ordine dei posti (vedi openCatModal)
  const visiblePending = catPending
    .filter(c => !c.deleted)
    .sort((a, b) => a.pos - b.pos);

  const existingHTML = visiblePending.map(c => {
    const pidx = catPending.indexOf(c);
    const canDel = c.count === 0;
    return `
      <div class="cat-item" data-pidx="${pidx}">
        <span class="cat-drag-handle" data-drag="1" aria-label="Trascina per riordinare">${DRAG_SVG}</span>
        <input class="cat-name-input" type="text" value="${esc(c.editNome)}" maxlength="80"
               data-pidx="${pidx}" autocomplete="off" spellcheck="false">
        <span class="cat-count">${c.count > 0 ? c.count + ' cont.' : '—'}</span>
        <button class="cat-del-btn ${canDel ? 'can-del' : 'no-del'}"
                data-pidx="${pidx}" data-type="existing" type="button"
                title="${canDel ? 'Elimina' : 'Ha contatti — elimina prima i contatti'}">${TRASH_SVG}</button>
      </div>`;
  }).join('');

  const newHTML = catPendingNew.map((nc, ni) => `
    <div class="cat-item">
      <span class="cat-drag-handle spento" aria-hidden="true">${DRAG_SVG}</span>
      <input class="cat-name-input" type="text" value="${esc(nc.nome)}" maxlength="80"
             data-nidx="${ni}" autocomplete="off" spellcheck="false">
      <span class="cat-count nuovo">nuovo</span>
      <button class="cat-del-btn can-del" data-nidx="${ni}" data-type="new"
              type="button" title="Rimuovi">${TRASH_SVG}</button>
    </div>`).join('');

  D.catList.innerHTML = existingHTML + newHTML;
  setupCatDrag();

  D.catList.querySelectorAll('.cat-name-input[data-pidx]').forEach(input => {
    const idx = Number(input.dataset.pidx);
    input.addEventListener('input', () => { catPending[idx].editNome = input.value; });
    input.addEventListener('blur', () => {
      const f = formatCatName(input.value) || catPending[idx].nome;
      input.value = f;
      catPending[idx].editNome = f;
    });
  });

  D.catList.querySelectorAll('.cat-name-input[data-nidx]').forEach(input => {
    const ni = Number(input.dataset.nidx);
    input.addEventListener('input', () => { catPendingNew[ni].nome = input.value; });
    input.addEventListener('blur', () => {
      const f = formatCatName(input.value) || catPendingNew[ni].nome;
      input.value = f;
      catPendingNew[ni].nome = f;
    });
  });

  D.catList.querySelectorAll('.cat-del-btn').forEach(btn => {
    if (btn.classList.contains('no-del')) {
      btn.addEventListener('click', () => {
        const c = catPending[Number(btn.dataset.pidx)];
        showToast(`"${c.nome}" ha ${c.count} contatti — non può essere eliminata`, 4000, 'warning');
      });
      return;
    }
    btn.addEventListener('click', () => {
      if (btn.dataset.type === 'existing') {
        catPending[Number(btn.dataset.pidx)].deleted = true;
      } else {
        catPendingNew.splice(Number(btn.dataset.nidx), 1);
      }
      renderCatList();
    });
  });
}

function addNewCategory() {
  const nome = formatCatName(D.fNewCat.value);
  if (!nome) { showToast('Nome categoria non valido', 3000, 'warning'); return; }

  const allNames = [
    ...catPending.filter(c => !c.deleted).map(c => c.editNome.toLowerCase()),
    ...catPendingNew.map(c => c.nome.toLowerCase())
  ];
  if (allNames.includes(nome.toLowerCase())) {
    showToast('Categoria già esistente: ' + nome, 3000, 'warning'); return;
  }
  catPendingNew.push({ nome });
  D.fNewCat.value = '';
  renderCatList();
  D.fNewCat.focus();
}

// Che cosa c'è da salvare, calcolato dalla finestra com'è IN QUESTO MOMENTO:
// { errore } oppure { adds, renames, deletes, ordini, conContatti }.
function cambiCategorie() {
  const finalNames = [
    ...catPending.filter(c => !c.deleted).map(c => formatCatName(c.editNome)),
    ...catPendingNew.map(c => formatCatName(c.nome))
  ];
  if (finalNames.some(n => !n)) return { errore: 'Il nome di una categoria è vuoto o non valido' };
  const lower = finalNames.map(n => n.toLowerCase());
  if (lower.some((n, i) => lower.indexOf(n) !== i)) return { errore: 'Ci sono categorie con lo stesso nome' };

  // una categoria non toccata non si rinomina, nemmeno se il suo nome non ha la forma di formatCatName
  const rinominate = catPending.filter(c => !c.deleted && c.editNome !== c.nome && formatCatName(c.editNome) !== c.nome);
  return {
    adds:    catPendingNew.map(nc => formatCatName(nc.nome)),
    renames: rinominate.map(c => ({ old: c.nome, new: formatCatName(c.editNome) })),
    deletes: catPending.filter(c => c.deleted).map(c => c.nome),
    ordini:  nuoviOrdini(),
    conContatti: rinominate.filter(c => c.count > 0).map(c =>
      `• "${c.nome}" → "${formatCatName(c.editNome)}" (${c.count} contatti verranno aggiornati)`),
  };
}

async function saveCatChanges() {
  if (scrittureInCorso) return;        // un salvataggio è già partito
  let cambi = cambiCategorie();
  if (cambi.errore) { showToast(cambi.errore, 3000, 'warning'); return; }

  if (cambi.conContatti.length) {
    const mostrato = JSON.stringify(cambi);
    if (!(await showConfirm(`Attenzione — verranno aggiornati i contatti:\n\n${cambi.conContatti.join('\n')}\n\nProcedere?`, 'warning'))) return;
    // La risposta vale per ciò che la domanda mostrava: se nel frattempo la
    // finestra è cambiata, o è stata chiusa, non si salva nulla.
    if (D.catModalOverlay.hidden) return;
    cambi = cambiCategorie();
    if (cambi.errore || JSON.stringify(cambi) !== mostrato) {
      showToast('Le categorie sono cambiate mentre la domanda era a schermo: controlla e salva di nuovo', 5000, 'warning');
      return;
    }
  }
  if (!cambi.adds.length && !cambi.renames.length && !cambi.deletes.length && !cambi.ordini.length) {
    chiudiCategorie();
    showToast('Nessuna modifica da salvare', 2500);
    return;
  }

  const btn = D.btnCatSave;
  btn.innerHTML = '<span class="btn-spinner"></span>'; btn.disabled = true;
  scrittureInCorso++;
  let fatto = false;   // una parte delle modifiche è già nel database
  try {
    // UNA chiamata, una sola transazione nel database. Se risponde con un
    // errore (nome già esistente, categoria che ha ancora contatti, sessione…)
    // nulla è stato cambiato: lo si dice e basta. Mai rifare lo stesso lavoro
    // «a pezzi» con chiamate separate.
    if (cambi.adds.length || cambi.renames.length || cambi.deletes.length) {
      await supaFetch('rpc/rubrica_categorie_batch', {
        method: 'POST',
        body: JSON.stringify({ adds: cambi.adds, renames: cambi.renames, deletes: cambi.deletes }),
      });
      fatto = true;
    }

    // ── Riordino: le sole categorie il cui «ordine» cambia (nuoviOrdini) ──────
    // Una scrittura per categoria: la funzione del database non ha un
    // parametro per l'ordine. Se si fermano a metà possono restare due valori
    // uguali: lo si dice, e il riordino successivo li separa.
    for (const o of cambi.ordini) {
      const id = idIntero(o.id);
      if (id === null) continue;
      await supaFetch(`rubrica_categorie?id=eq.${id}`, {
        method: 'PATCH',
        body: JSON.stringify({ ordine: o.ordine }),
      });
      fatto = true;
    }

    chiudiCategorie();
    showToast('Categorie aggiornate con successo', 3000, 'success');
    loadData(true);
  } catch (e) {
    const incerto = !!(e && (e.rete || e.incerto));
    if (fatto || incerto) {
      // Una parte è già nel database, oppure non si sa (risposta persa): la
      // finestra non è più lo specchio di ciò che c'è. Si chiude e si rilegge:
      // riaprendola si vede che cosa è stato salvato davvero.
      chiudiCategorie();
      showToast((fatto ? 'Categorie salvate solo in parte' : 'Non è certo che le categorie siano state salvate') + ': riapri la finestra e controlla. ' + messaggioErrore(e), 7000, 'error');
      loadData(true);
    } else {
      showToast(messaggioErrore(e), 5000, 'error');
      if (e instanceof ErroreSessione) loadData(true);
    }
  } finally {
    scrittureInCorso--;
    btn.textContent = 'Salva modifiche'; btn.disabled = false;
  }
}

// ── Toast / Confirm ──────────────────────────────────────────────────────────
// (il contenitore si chiama ancora installToast: è quello di tutti gli avvisi)
function showToast(msg, duration = 3500, type = '') {
  annullaConferma();     // un avviso prende il posto della domanda a schermo: vale «no»
  const t = D.installToast;
  t.textContent = msg;
  t.className = 'show' + (type ? ' toast-' + type : '');
  clearTimeout(t._hideTimer);
  t._hideTimer = setTimeout(() => { t.classList.remove('show'); }, duration);
}

// Una domanda alla volta. Finché è a schermo ciò che sta sotto non si tocca:
// una scheda chiusa e riaperta su un altro contatto, o una categoria rinominata
// nel frattempo, farebbero rispondere «sì» a una domanda diversa da quella
// letta. Chi chiude la finestra sotto (closeModal, chiudiCategorie) o mostra
// un avviso fa valere «no» per la domanda rimasta.
let confermaInSospeso = null;
function annullaConferma() {
  if (confermaInSospeso) confermaInSospeso(false);
}
function showConfirm(msg, type = 'warning') {
  annullaConferma();
  return new Promise(resolve => {
    const t = D.installToast;
    clearTimeout(t._hideTimer);
    t.innerHTML = `
      <div class="tc-msg">${esc(msg).replace(/\n/g,'<br>')}</div>
      <div class="tc-btns">
        <button type="button" class="tc-no">Annulla</button>
        <button type="button" class="tc-yes">Conferma</button>
      </div>`;
    t.className = 'show toast-confirm' + (type ? ' toast-' + type : '');
    const sotto = [...document.body.children].filter(el => el !== t && el.tagName !== 'SCRIPT');
    sotto.forEach(el => { el.inert = true; });
    const colFuoco = document.activeElement;
    const done = ok => {
      if (confermaInSospeso !== done) return;
      confermaInSospeso = null;
      sotto.forEach(el => { el.inert = false; });
      t.classList.remove('show');
      if (document.hasFocus() && colFuoco && colFuoco.isConnected && colFuoco.focus) colFuoco.focus();
      resolve(ok);
    };
    confermaInSospeso = done;
    t.querySelector('.tc-yes').addEventListener('click', () => done(true),  { once: true });
    t.querySelector('.tc-no') .addEventListener('click', () => done(false), { once: true });
    if (document.hasFocus()) t.querySelector('.tc-no').focus();
  });
}

// ── Boot ─────────────────────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', avvia);
