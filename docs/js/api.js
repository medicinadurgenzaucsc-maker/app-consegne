// ============================================================
// api.js — Backend Supabase v2
// Sostituisce GAS + Google Sheets con Supabase (Postgres + Realtime)
// ============================================================


// ── CONFIGURAZIONE ────────────────────────────────────────────
var SUPABASE_URL     = 'https://ifmmcvxzhwdkmzhsxcvb.supabase.co';
var SUPABASE_ANON_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImlmbW1jdnh6aHdka216aHN4Y3ZiIiwicm9sZSI6ImFub24iLCJpYXQiOjE3Nzc2Nzk1ODAsImV4cCI6MjA5MzI1NTU4MH0.LH8h4Fivtl3-TuiA050oF8iS4b80xrd2Dn6z8JjCoeA';
var APP_URL          = 'https://medicinadurgenzaucsc-maker.github.io/app-consegne/';
var PRINT_URL        = APP_URL + 'print.html';
var LOCK_TTL_MS      = 600000; // 10 minuti — safety net DB-side, deve essere > timer client (5 min) per evitare race

// ── GOOGLE OAUTH2 ─────────────────────────────────────────────
// Client ID OAuth2 da Google Cloud Console (Web application).
// Istruzioni: console.cloud.google.com → API e servizi → Credenziali →
// Crea credenziali → ID client OAuth2 → Web application
// Origini JS autorizzate: https://medicinadurgenzaucsc-maker.github.io
var GOOGLE_CLIENT_ID = '170256871056-gchf386c3oic77ek2j5m3b1e5pbv6cre.apps.googleusercontent.com';

// Token di accesso Google — impostato dopo il login, usato per Drive API
window._googleAccessToken = null;

// ── CLIENT SUPABASE ───────────────────────────────────────────
var _sb = supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY);

// ── STATO LOCK IN MEMORIA ─────────────────────────────────────
// Aggiornato da Realtime: nessuna query aggiuntiva al DB per leggere i lock.
// Struttura: { [letto]: { token } }
var _lockState = {};


// ── MAPPING CAMPI JS ↔ COLONNE SQL ───────────────────────────
var _campi = {
  'Letto':               'letto',
  'Nome':                'nome',
  'Eta':                 'eta',
  'DataNascita':         'data_nascita',
  'DataRicovero':        'data_ricovero',
  'Diagnosi':            'diagnosi',
  'NoteTerapia':         'note_terapia',
  'Diaria':              'diaria',
  'DaFare':              'da_fare',
  'TipologiaLetto':      'tipologia_letto',
  'PianoTerapeutico':    'piano_terapeutico',
  'EsamiColturali':      'esami_colturali',
  'Allergie':            'allergie',
  'CodiceSanitario':     'codice_sanitario',
  'Ossigeno':            'ossigeno',
  'Vitto':               'vitto',
  'Sesso':               'sesso',
  'Dimissibile':         'dimissibile',
  'UltimoAggiornamento': 'ultimo_aggiornamento'
};
var _colonne = {};
Object.keys(_campi).forEach(function(k) { _colonne[_campi[k]] = k; });

function _fromDb(row) {
  var p = {};
  Object.keys(row).forEach(function(col) {
    var field = _colonne[col] || col;
    p[field] = (row[col] !== null && row[col] !== undefined) ? String(row[col]) : '';
  });
  return p;
}

function _toDb(datiPaziente) {
  var row = {};
  Object.keys(datiPaziente).forEach(function(field) {
    var col = _campi[field];
    if (col && col !== 'letto') row[col] = datiPaziente[field] != null ? datiPaziente[field] : '';
  });
  return row;
}

function _oggiStr() {
  var d = new Date();
  return d.getFullYear() + '-' +
    String(d.getMonth() + 1).padStart(2, '0') + '-' +
    String(d.getDate()).padStart(2, '0');
}

function _oraStr() {
  var d = new Date();
  return String(d.getHours()).padStart(2, '0') + ':' +
    String(d.getMinutes()).padStart(2, '0') + ':' +
    String(d.getSeconds()).padStart(2, '0');
}

// Promise wrapper per Supabase
function _q(queryPromise) {
  return new Promise(function(resolve, reject) {
    queryPromise.then(function(res) {
      if (res.error) { reject(res.error); } else { resolve(res.data); }
    }).catch(reject);
  });
}


// ── HELPER ESCAPE ATTRIBUTI HTML ─────────────────────────────
function _aEsc(s) {
  return String(s || '')
    .replace(/&/g, '&amp;')
    .replace(/"/g, '&quot;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

// ══════════════════════════════════════════════════════════════
// SISTEMA LOG — registra eventi diagnostici nella tabella `logs`
// Visibili al medico/IT dal menu rotellina → "Log".
// Identifica il PC tramite:
//   - device_id: UUID persistente in localStorage (univoco per PC+browser)
//   - device_name: descrizione AUTOMATICA generata da browser+OS+schermo+
//                  città timezone + lingua (non configurabile dall'utente
//                  per non confondere — il browser NON può ottenere il
//                  vero hostname Windows per motivi di privacy)
// ══════════════════════════════════════════════════════════════
function _getDeviceId() {
  try {
    // Cleanup retro-compatibilità: rimuovi il vecchio _deviceName
    // (ora il nome è generato automaticamente, non più configurabile)
    if (localStorage.getItem('_deviceName') !== null) {
      localStorage.removeItem('_deviceName');
    }
    var id = localStorage.getItem('_deviceId');
    if (!id) {
      id = 'dev_' + Date.now().toString(36) + '_' + Math.random().toString(36).slice(2, 10);
      localStorage.setItem('_deviceId', id);
    }
    return id;
  } catch(e) { return 'dev_unknown'; }
}
function _parseUA(ua) {
  if (!ua) return 'Sconosciuto';
  var browser = '?', os = '?';
  var m;
  if ((m = ua.match(/Edg\/(\d+)/)))           browser = 'Edge ' + m[1];
  else if ((m = ua.match(/Chrome\/(\d+)/)))   browser = 'Chrome ' + m[1];
  else if ((m = ua.match(/Firefox\/(\d+)/)))  browser = 'Firefox ' + m[1];
  else if (/Safari\//.test(ua))               browser = 'Safari';
  if (/Windows NT/.test(ua))                  os = 'Windows';
  else if (/Mac OS X/.test(ua))               os = 'macOS';
  else if (/iPhone|iPad/.test(ua))            os = 'iOS';
  else if (/Android/.test(ua))                os = 'Android';
  else if (/Linux/.test(ua))                  os = 'Linux';
  return browser + ' / ' + os;
}
// Genera una descrizione AUTOMATICA del PC con tutte le info disponibili.
// Es: "Edge 148 / Windows · 1920×1080 · Rome · it-IT"
function _getDeviceName() {
  try {
    var ua = _parseUA(navigator.userAgent || '');
    var risol = (window.screen && window.screen.width)
      ? (window.screen.width + '×' + window.screen.height)
      : '?';
    var citta = '';
    try {
      var tz = Intl.DateTimeFormat().resolvedOptions().timeZone || '';
      citta = (tz.indexOf('/') !== -1) ? tz.split('/').pop().replace(/_/g, ' ') : tz;
    } catch(_) {}
    var lang = navigator.language || '';
    var parts = [ua, risol];
    if (citta) parts.push(citta);
    if (lang) parts.push(lang);
    return parts.join(' · ');
  } catch(e) {
    return 'Sconosciuto';
  }
}

// Inserisci un log nella tabella `logs` (async, fire-and-forget).
// livello: 'info' | 'warning' | 'error'
// tipo:    stringa libera per raggruppare (es. 'proxy', 'save-failed')
// messaggio:   breve, leggibile (es. "3 fetch fallite in 30s")
// descrizione: dettagli verbose opzionali
// url:         URL coinvolto opzionale
function _log(livello, tipo, messaggio, descrizione, url) {
  try {
    var row = {
      ts:          Date.now(),
      livello:     livello || 'info',
      tipo:        tipo || 'generic',
      messaggio:   String(messaggio || '').slice(0, 500),
      descrizione: descrizione ? String(descrizione).slice(0, 2000) : null,
      device_id:   _getDeviceId(),
      device_name: _getDeviceName(),  // auto: 'Edge 148 / Windows · 1920×1080 · Rome · it-IT'
      user_agent:  _parseUA(navigator.userAgent),
      url:         url ? String(url).slice(0, 500) : null
    };
    if (typeof _sb === 'undefined' || !_sb) return; // Supabase non ancora caricato
    _q(_sb.from('logs').insert(row)).catch(function(e) {
      // Non re-loggare l'errore per evitare loop infiniti
      try { console.warn('[Log] insert fallito:', e && e.message); } catch(_) {}
    });
  } catch(e) {
    try { console.warn('[Log] errore interno:', e && e.message); } catch(_) {}
  }
}

// Recupera gli ultimi N log ordinati dal più recente al più vecchio.
// Limite default: 200. Filtri opzionali su livello.
function _sbGetLogs(limit, livelloFiltro) {
  var q = _sb.from('logs').select('*').order('ts', { ascending: false }).limit(limit || 200);
  if (livelloFiltro) q = q.eq('livello', livelloFiltro);
  return _q(q);
}

// Cleanup: elimina log più vecchi di N giorni (default 20).
// Chiamato:
//   1. All'apertura del modal log (per dare il dato fresco all'utente)
//   2. All'avvio dell'app (per non lasciare i log in eccesso anche
//      se nessuno apre mai il modal)
// Il DELETE è idempotente e usa l'indice idx_logs_ts (DESC), quindi
// è veloce anche su tabelle grandi.
function _sbPuliciLogVecchi(giorni) {
  var g = (typeof giorni === 'number' && giorni > 0) ? giorni : 20;
  var soglia = Date.now() - (g * 86400000);
  return _q(_sb.from('logs').delete().lt('ts', soglia))
    .then(function() {
      console.log('[Log] Cleanup OK (soglia: ' + g + ' giorni = ts < ' + soglia + ')');
      return { ok: true, soglia: soglia, giorni: g };
    });
}

// Cancella TUTTI i log (azione admin)
function _sbCancellaTuttiLog() {
  return _q(_sb.from('logs').delete().gte('id', 0));
}

// Espone le funzioni a window per uso globale
window._log              = _log;
window._getDeviceId      = _getDeviceId;
window._getDeviceName    = _getDeviceName;  // auto, non configurabile
window._parseUA          = _parseUA;
window._sbGetLogs        = _sbGetLogs;
window._sbPuliciLogVecchi = _sbPuliciLogVecchi;
window._sbCancellaTuttiLog = _sbCancellaTuttiLog;


// ── COLORE DA STRINGA ─────────────────────────────────────────
function stringToColor(str) {
  if (!str) return '#546e7a';
  var hash = 0;
  for (var i = 0; i < str.length; i++) {
    hash = str.charCodeAt(i) + ((hash << 5) - hash);
  }
  var h = Math.abs(hash) % 360;
  return 'hsl(' + h + ',45%,40%)';
}


// ── CALCOLI DATE ──────────────────────────────────────────────
function _parseDataRicovero(raw) {
  if (!raw) return { vis: '', giorni: '-' };
  var s = String(raw).trim();
  var d;
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) {
    var p = s.substring(0, 10).split('-');
    d = new Date(parseInt(p[0], 10), parseInt(p[1], 10) - 1, parseInt(p[2], 10));
  } else if (/^\d{2}\/\d{2}\/\d{4}/.test(s)) {
    var p = s.split('/');
    d = new Date(parseInt(p[2], 10), parseInt(p[1], 10) - 1, parseInt(p[0], 10));
  } else {
    d = new Date(s);
  }
  if (!d || isNaN(d.getTime())) return { vis: s, giorni: '-' };
  d.setHours(0, 0, 0, 0);
  var oggi = new Date(); oggi.setHours(0, 0, 0, 0);
  var giorni = Math.floor((oggi - d) / 86400000);
  var vis = String(d.getDate()).padStart(2, '0') + '/' +
            String(d.getMonth() + 1).padStart(2, '0') + '/' +
            d.getFullYear();
  return { vis: vis, giorni: giorni >= 0 ? giorni : '-' };
}

function _parseDataNascita(raw) {
  if (!raw) return { vis: '', eta: '', hasData: false };
  var s = String(raw).trim();
  var d = null;
  if (/^\d{2}\/\d{2}\/\d{4}$/.test(s)) {
    var p = s.split('/');
    d = new Date(parseInt(p[2], 10), parseInt(p[1], 10) - 1, parseInt(p[0], 10));
  } else if (/^\d{4}-\d{2}-\d{2}/.test(s)) {
    var p = s.substring(0, 10).split('-');
    d = new Date(parseInt(p[0], 10), parseInt(p[1], 10) - 1, parseInt(p[2], 10));
  }
  if (!d || isNaN(d.getTime())) return { vis: s, eta: '', hasData: false };
  d.setHours(0, 0, 0, 0);
  var vis = String(d.getDate()).padStart(2, '0') + '/' +
            String(d.getMonth() + 1).padStart(2, '0') + '/' +
            d.getFullYear();
  var oggi = new Date(); oggi.setHours(0, 0, 0, 0);
  var anni = oggi.getFullYear() - d.getFullYear();
  if (oggi.getMonth() < d.getMonth() ||
      (oggi.getMonth() === d.getMonth() && oggi.getDate() < d.getDate())) anni--;
  return { vis: vis, eta: String(anni), hasData: true };
}


// NB: _renderMainCard (vista standard) rimosso il 2026-05-11 — codice
// preservato in docs/_legacy/standard-view.md per eventuale re-integrazione.

// ── Icona "Dimissione" — due frecce orizzontali in direzioni opposte:
// rosso scuro (#b71c1c) sopra che punta a destra (paziente che esce),
// grigio scuro (#424242) sotto che punta a sinistra (letto che si libera).
// Esposta come window.* per essere riutilizzata da app.js e dall'IIFE
// snapshot in index.html senza duplicare il markup.
function _dimIconSvg(size) {
  var s = size || 13;
  return '<svg width="' + s + '" height="' + s + '" viewBox="0 0 24 24" aria-hidden="true" style="vertical-align:middle;">' +
    '<path d="M3 8h13 M12 4l4 4-4 4" stroke="#b71c1c" stroke-width="2.6" stroke-linecap="round" stroke-linejoin="round" fill="none"/>' +
    '<path d="M21 16H8 M12 12l-4 4 4 4" stroke="#424242" stroke-width="2.6" stroke-linecap="round" stroke-linejoin="round" fill="none"/>' +
  '</svg>';
}
window._dimIconSvg = _dimIconSvg;

// Icona "virus" verdognola per l'ISOLAMENTO paziente.
// SVG inline (come _dimIconSvg): look identico ovunque, indipendente
// dal font Bootstrap Icons. Usata in: toolbar Diaria, banner isolamento,
// icona accanto al numero letto, stampa.
function _virusIconSvg(size) {
  var s = size || 16;
  return '<svg width="' + s + '" height="' + s + '" viewBox="0 0 24 24" aria-hidden="true" style="vertical-align:middle;">' +
    '<g stroke="#2e7d32" stroke-width="1.8" stroke-linecap="round">' +
      '<path d="M12 1.5v3.5 M12 19v3.5 M1.5 12H5 M19 12h3.5 M4.6 4.6l2.5 2.5 M16.9 16.9l2.5 2.5 M19.4 4.6l-2.5 2.5 M7.1 16.9l-2.5 2.5" fill="none"/>' +
    '</g>' +
    '<circle cx="12" cy="12" r="6.8" fill="#66bb6a" stroke="#2e7d32" stroke-width="1.2"/>' +
    '<circle cx="9.8" cy="10.4" r="1.3" fill="#e8f5e9"/>' +
    '<circle cx="14.2" cy="12.8" r="1.6" fill="#e8f5e9"/>' +
    '<circle cx="10.8" cy="14.8" r="1.0" fill="#e8f5e9"/>' +
  '</svg>';
}
window._virusIconSvg = _virusIconSvg;

// Icona "cuore + battito ECG" per la RIANIMAZIONE paziente.
// Cuore rosso pieno con linea di battito bianca. Come le altre, SVG
// inline (identico ovunque). Usata in: toolbar Diaria, banner
// rianimazione, icona accanto al numero letto, stampa.
function _rianimIconSvg(size) {
  var s = size || 16;
  return '<svg width="' + s + '" height="' + s + '" viewBox="0 0 24 24" aria-hidden="true" style="vertical-align:middle;">' +
    '<path d="M12 20.5S3.5 15 3.5 8.9C3.5 6 5.6 4 8.1 4c1.7 0 3.1 1 3.9 2.3C12.8 5 14.2 4 15.9 4c2.5 0 4.6 2 4.6 4.9 0 6.1-8.5 11.6-8.5 11.6z" fill="#e53935" stroke="#b71c1c" stroke-width="1.1"/>' +
    '<path d="M5.5 12.4h3l1.4-3.1 2.1 5.2 1.3-2.6h4.7" stroke="#fff" stroke-width="1.5" fill="none" stroke-linecap="round" stroke-linejoin="round"/>' +
  '</svg>';
}
window._rianimIconSvg = _rianimIconSvg;

// ── RENDERER CARD (vista compatta a riga, unica vista attiva) ─
function _renderAltCard(p) {
  var letto = String(p.Letto || '');
  var tipo  = String(p.TipologiaLetto || '').trim().toUpperCase();
  var ric   = _parseDataRicovero(p.DataRicovero);
  var nasc  = _parseDataNascita(p.DataNascita);
  var eta   = nasc.hasData ? nasc.eta : (p.Eta || '');
  var colore = tipo ? stringToColor(tipo) : 'transparent';

  var badgeHtml = tipo
    ? '<span class="badge text-white text-uppercase shadow-sm"' +
      ' style="background-color:' + colore + ';cursor:pointer;"' +
      ' id="badge-tipo-alt-' + _aEsc(letto) + '"' +
      ' onclick="_apriModalTipologia(\'' + _aEsc(letto) + '\')"' +
      ' title="Click per cambiare tipologia (solo in modifica)">' + _aEsc(tipo) + '</span>'
    : '<span class="badge bg-light text-secondary border text-uppercase shadow-sm"' +
      ' style="cursor:pointer;"' +
      ' onclick="_apriModalTipologia(\'' + _aEsc(letto) + '\')"' +
      ' title="Click per cambiare tipologia (solo in modifica)">STANDARD</span>';

  // Icona "in dimissione" accanto al badge tipologia
  var dimVal = (p.Dimissibile || '').trim();
  var dimIconHtml = dimVal
    ? '<span class="dim-icon-badge ms-1" title="Paziente in dimissione' +
      (dimVal ? ' — ' + _aEsc(dimVal) : '') + '">' +
      _dimIconSvg(14) + '</span>'
    : '';

  // ── Icone di stato accanto al letto (isolamento + rianimazione).
  // Fonte di verità = banner dentro la Diaria (viaggiano con l'HTML del
  // campo: realtime, stampa, backup, import). Qui leggiamo i marker per il
  // primo render; poi _bedStatusSyncIcone (index.html) le tiene allineate.
  // Le due icone stanno AFFIANCATE in un contenitore condiviso.
  var isoMatch    = String(p.Diaria || '').match(/data-isolamento="([^"]*)"/);
  var rianimMatch = String(p.Diaria || '').match(/data-rianimazione="([^"]*)"/);
  var statusIcons = '';
  if (isoMatch) {
    statusIcons += '<span class="bed-status-ico iso" title="Paziente in isolamento' + (isoMatch[1] ? ': ' + _aEsc(isoMatch[1]) : '') + '">' + _virusIconSvg(26) + '</span>';
  }
  if (rianimMatch) {
    statusIcons += '<span class="bed-status-ico rianim" title="Rianimazione' + (rianimMatch[1] ? ': ' + _aEsc(rianimMatch[1]) : '') + '">' + _rianimIconSvg(26) + '</span>';
  }
  var statusIconsHtml = statusIcons ? '<div class="bed-status-icons">' + statusIcons + '</div>' : '';

  return '<div class="alt-row patient-card" data-bed="' + _aEsc(letto) + '" data-tipologia="' + _aEsc(tipo) + '">' +
    '<div class="alt-col alt-col-info">' +
    '<button class="focus-pencil-btn" title="Modifica scheda"' +
    ' onclick="_attivaFocusMode(this.closest(\'.patient-card\'))"><i class="bi bi-pencil-fill"></i></button>' +
    '<span id="status-alt-' + _aEsc(letto) + '" class="badge status-badge position-absolute top-0 start-0 m-1 bg-secondary"></span>' +
    statusIconsHtml +
    '<div class="alt-info-box">' +
    '<div class="bed-number-flex">' +
    '<div class="alt-bed-number">' + _aEsc(letto) + '</div>' +
    '<span class="sesso-symbol" data-field="Sesso" data-sesso="' + _aEsc(p.Sesso || '') + '"></span>' +
    '</div>' +
    '<div class="text-center tipo-badge-wrap" style="font-size:0.75rem;">' + badgeHtml + dimIconHtml + '</div>' +
    '<div class="editable-area plain-text alt-nome" contenteditable="true" data-field="Nome" data-placeholder="Cognome Nome...">' + (p.Nome || '') + '</div>' +
    '<div class="alt-allergie-label"><i class="bi bi-exclamation-triangle-fill"></i> Allergie</div>' +
    '<div class="editable-area rich-text alt-allergie-val" contenteditable="true" data-field="Allergie" data-placeholder="Allergia ad antibiotici...">' + (p.Allergie || '') + '</div>' +
    '<div class="alt-info-row"><span class="alt-info-label">Data di Nascita</span>' +
    '<input type="text" class="data-nascita-text alt-info-val" placeholder="gg/mm/aaaa"' +
    ' oninput="formattaDataNascita(this)" onblur="formattaDataNascitaBlur(this)" readonly' +
    ' value="' + _aEsc(nasc.vis) + '"' +
    ' style="border:none;border-bottom:1px solid #ccc;outline:none;background:transparent;padding:0 2px;"></div>' +
    '<div class="alt-info-row"><span class="alt-info-label">Et&agrave; </span>' +
    '<div class="editable-area plain-text alt-info-val"' +
    ' contenteditable="' + (nasc.hasData ? 'false' : 'true') + '" data-field="Eta"' +
    ' style="' + (nasc.hasData ? 'opacity:0.6;cursor:not-allowed;' : '') + '">' + eta + '</div></div>' +
    '<div class="alt-info-row"><span class="alt-info-label">Ricovero</span>' +
    '<input type="text" class="data-ricovero-text alt-info-val" placeholder="gg/mm/aaaa"' +
    ' oninput="formattaDataRicovero(this)" onblur="formattaDataRicoveroBlur(this)" readonly' +
    ' value="' + _aEsc(ric.vis) + '"' +
    ' style="border:none;border-bottom:1px solid #ccc;outline:none;background:transparent;padding:0 2px;"></div>' +
    '<div class="alt-info-row"><span class="alt-info-label">Giorni di Ricovero</span>' +
    '<span class="text-danger valore-giorni alt-info-val">' + ric.giorni + '</span></div>' +
    '<div class="alt-info-row"><span class="alt-info-label">C.S.</span>' +
    '<div class="editable-area plain-text alt-info-val" contenteditable="true" data-field="CodiceSanitario">' + (p.CodiceSanitario || '') + '</div></div>' +
    '<div class="alt-info-row"><span class="alt-info-label">Ossigeno</span>' +
    '<div class="editable-area plain-text alt-info-val" contenteditable="true" data-field="Ossigeno" data-placeholder="es. CN 4lt/min">' + (p.Ossigeno || '') + '</div></div>' +
    '<div class="alt-info-row"><span class="alt-info-label">Vitto</span>' +
    '<div class="editable-area plain-text alt-info-val" contenteditable="true" data-field="Vitto" data-placeholder="Vitto completo, diabetico, senza scorie...">' + (p.Vitto || '') + '</div></div>' +
    '<div class="alt-info-row dim-row" data-field="Dimissibile" data-value="' + _aEsc(dimVal) + '">' +
    '<label class="dim-checkbox-wrap" title="' + (dimVal ? 'Click per modificare/rimuovere' : 'Click per impostare data dimissione') + '">' +
    '<input type="checkbox" class="dim-checkbox"' + (dimVal ? ' checked' : '') + ' aria-label="Dimissibile">' +
    '<span class="dim-checkbox-label">Dimissibile</span>' +
    '</label>' +
    (dimVal
      ? '<span class="dim-data-info">' + _dimIconSvg(12) + ' Dimissione il ' + _aEsc(dimVal) + '</span>'
      : '<span class="dim-data-info-empty text-muted"></span>') +
    '</div>' +
    '</div></div>' +
    '<div class="alt-col alt-col-diag">' +
    '<div class="alt-diag-top">' +
    '<div class="alt-col-header">Diagnosi / Motivo Ricovero</div>' +
    '<div class="editable-area plain-text alt-editable" contenteditable="true" data-field="Diagnosi" data-placeholder="Diagnosi e motivo del ricovero...">' + (p.Diagnosi || '') + '</div>' +
    '</div>' +
    '<div class="alt-diag-bottom">' +
    '<div class="alt-col-header-split">Piano di Cura</div>' +
    '<div class="editable-area rich-text alt-editable" contenteditable="true" data-field="PianoTerapeutico" data-placeholder="Piano di cura...">' + (p.PianoTerapeutico || '') + '</div>' +
    '</div>' +
    '<div class="alt-diag-fourth">' +
    '<div class="alt-col-header-split">Esami Colturali e Scale Valutazione</div>' +
    '<div class="editable-area rich-text alt-editable" contenteditable="true" data-field="EsamiColturali" data-placeholder="Esami colturali... PESI SCORE, HASBLED, GENEVA...">' + (p.EsamiColturali || '') + '</div>' +
    '</div>' +
    '<div class="alt-diag-third">' +
    '<div class="alt-col-header-split">Da Fare / Richieste</div>' +
    '<div class="editable-area rich-text alt-editable" contenteditable="true" data-field="DaFare" data-placeholder="Da Fare / Richieste...">' + (p.DaFare || '') + '</div>' +
    '</div>' +
    '</div>' +
    '<div class="alt-col alt-col-diaria">' +
    '<div class="alt-col-header">Diaria ed Epicrisi</div>' +
    '<div class="editable-area rich-text alt-editable" contenteditable="true" data-field="Diaria" data-placeholder="Diaria ed epicrisi: APR, APP, storia prima del ricovero...">' + (p.Diaria || '') + '</div>' +
    '</div>' +
    '<div class="alt-col alt-col-terapia">' +
    '<div class="alt-col-header">Note e Terapia</div>' +
    '<div class="editable-area rich-text alt-editable" contenteditable="true" data-field="NoteTerapia" data-placeholder="Note e Terapia: Terapia EV, Orale, Aerosol...">' + (p.NoteTerapia || '') + '</div>' +
    '</div>' +
    '<div class="alt-row-spacer"></div>' +
    '</div>';
}


// ── RENDERER SCHEDA NOTE (layout principale) ──────────────────
// NB: _renderNoteCardMain (vista standard) rimosso il 2026-05-11 —
// codice in docs/_legacy/standard-view.md.

// ── RENDERER SCHEDA NOTE ──────────────────────────────────────
function _renderNoteCardAlt(p) {
  return '<div class="alt-row patient-card note-card" data-bed="NOTE">' +
    '<div class="alt-col alt-col-info" style="justify-content:center;">' +
    '<button class="focus-pencil-btn" title="Modifica note"' +
    ' onclick="_attivaFocusMode(this.closest(\'.patient-card\'))"><i class="bi bi-pencil-fill"></i></button>' +
    '<span id="status-alt-NOTE" class="badge status-badge position-absolute top-0 start-0 m-1 bg-secondary"></span>' +
    '<div class="alt-info-box">' +
    '<div class="alt-bed-number" style="font-size:1rem;letter-spacing:1px;">NOTE</div>' +
    '</div>' +
    '</div>' +
    '<div class="alt-col" style="flex:1;border-right:none;">' +
    '<div class="editable-area rich-text alt-editable" contenteditable="true" data-field="Diaria"' +
    ' data-placeholder="Note del reparto (visibili a tutti)..."' +
    ' style="min-height:300px;padding:10px;font-size:0.9rem;">' + (p.Diaria || '') + '</div>' +
    '</div>' +
    '<div class="alt-row-spacer"></div>' +
    '</div>';
}

// ── RENDERER COMPLETO ─────────────────────────────────────────
// Solo vista alt (compatta a riga). La vista standard è stata rimossa
// per alleggerire il bundle — codice preservato in docs/_legacy/.
// L'id "cardsContainer" è mantenuto (riusato per la vista alt) per
// retrocompatibilità con i selettori esistenti nel resto del codice.
function _renderCardsHtml(pazienti) {
  var html  = '<div class="container-fluid" id="cardsContainer">';
  var noteP = null;
  (pazienti || []).forEach(function(p) {
    if (p.Letto === 'NOTE') { noteP = p; return; }
    html += _renderAltCard(p);
  });
  // NOTE sempre in fondo
  if (noteP) html += _renderNoteCardAlt(noteP);
  html += '<div id="noLettiMsg" class="container text-center mt-5"' +
          ' style="display:none;"><i class="bi bi-inboxes text-muted" style="font-size:4rem;"></i>' +
          '<h4 class="mt-3 text-secondary">Nessun Letto configurato.</h4></div></div>';
  return html;
}


// ── CARICAMENTO INIZIALE CARD ─────────────────────────────────
function _caricaCardIniziali(container, onDone) {
  _sbGetPazienti().then(function(pazienti) {
    if (container) container.innerHTML = _renderCardsHtml(pazienti);
    if (typeof onDone === 'function') onDone(pazienti);
  }).catch(function(e) {
    console.error('Errore caricamento card:', e);
    if (typeof onDone === 'function') onDone([]);
  });
}


// ── LETTURA TOAST DA URL ──────────────────────────────────────
(function() {
  try {
    var params = new URLSearchParams(window.location.search);
    var toastType = params.get('toast');
    var toastMsg  = params.get('msg');
    if (toastType && toastMsg) {
      history.replaceState(null, '', window.location.pathname);
      window.addEventListener('load', function() {
        setTimeout(function() {
          if (typeof showToast === 'function') {
            var titolo = toastType === 'success' ? 'Successo' : 'Informazione';
            showToast(titolo, toastMsg, toastType);
          }
        }, 800);
      });
    }
  } catch(e) {}
})();


// ══════════════════════════════════════════════════════════════
// SUPABASE: FUNZIONI CRUD
// ══════════════════════════════════════════════════════════════

function _sbGetPazienti() {
  return _q(_sb.from('consegne').select('*')).then(function(rows) {
    var pazienti = (rows || []).map(_fromDb);
    pazienti.sort(function(a, b) {
      // NOTE sempre in fondo
      if (a.Letto === 'NOTE') return 1;
      if (b.Letto === 'NOTE') return -1;
      var nA = parseInt(a.Letto, 10), nB = parseInt(b.Letto, 10);
      return (!isNaN(nA) && !isNaN(nB)) ? nA - nB : String(a.Letto).localeCompare(String(b.Letto));
    });
    return pazienti;
  });
}

function _sbGetLettiFull() {
  return _q(_sb.from('consegne').select('letto,nome,tipologia_letto')).then(function(rows) {
    var letti = (rows || []).filter(function(r) { return r.letto !== 'NOTE'; }).map(function(r) {
      return { letto: r.letto, nome: r.nome || '', tipologia: (r.tipologia_letto || 'STANDARD').toUpperCase() };
    });
    letti.sort(function(a, b) {
      var nA = parseInt(a.letto, 10), nB = parseInt(b.letto, 10);
      return (!isNaN(nA) && !isNaN(nB)) ? nA - nB : a.letto.localeCompare(b.letto);
    });
    return letti;
  });
}

function _sbSalvaPaziente(letto, datiPaziente) {
  var row = _toDb(datiPaziente);
  row.ultimo_aggiornamento = _oraStr();
  row.updated_at = new Date().toISOString();
  // SAFETY: aggiunto .select() per ricevere indietro la/le riga/he aggiornata/e.
  // Senza .select() Supabase risponde con array vuoto se non ha matchato nulla,
  // ma _q lo considera comunque "success". Con .select() verifichiamo che almeno
  // una riga sia stata effettivamente aggiornata. Questo intercetta casi come:
  //  - lettoo eliminato da un altro PC tra il momento del lettura e quello del save
  //  - cache CDN/proxy che restituisce 200 OK senza che la UPDATE sia avvenuta
  //  - rete instabile in modalità emergenza con risposta corrotta
  return _q(_sb.from('consegne').update(row).eq('letto', String(letto)).select())
    .then(function(data) {
      var rowsUpdated = Array.isArray(data) ? data.length : 0;
      if (rowsUpdated === 0) {
        // CRITICO: l'UPDATE è andato ma 0 righe modificate.
        // Significa che il letto non esiste più nel DB o che la rete ha
        // restituito una risposta corrotta. NON segnare success.
        console.error('[Save] UPDATE letto '+letto+' restituito 0 righe — il letto non esiste nel DB o la rete è corrotta');
        return { success: false, message: 'Salvataggio non confermato dal server (0 righe aggiornate). Dati locali NON sovrascritti.' };
      }
      return { success: true, ora: row.ultimo_aggiornamento, rowsUpdated: rowsUpdated };
    });
}

// Legge UNA singola riga consegne per il letto specificato.
// Usato in apertura focus mode per garantire dati freschi prima
// dell'editing — sia in Realtime normale (sicurezza extra) sia in
// modalità emergenza (dove il polling può essere stale fino a 30s).
// Payload: ~12 KB (1 riga rich-text) vs ~298 KB di SELECT *.
function _sbGetPazienteSingolo(letto) {
  return _q(_sb.from('consegne').select('*').eq('letto', String(letto)).maybeSingle())
    .then(function(row) {
      if (!row) return null;
      return _fromDb(row);
    });
}

// Importa una scheda letto dal parsing Google Doc.
// Aggiorna il letto se esiste, restituisce false se il letto non è presente.
function _sbImportaLetto(letto, dati) {
  return _q(_sb.from('consegne').select('letto').eq('letto', String(letto)).maybeSingle())
    .then(function(existing) {
      if (!existing) return false; // letto non esiste nel DB
      var row = _toDb(dati);
      row.ultimo_aggiornamento = _oraStr();
      row.updated_at = new Date().toISOString();
      return _q(_sb.from('consegne').update(row).eq('letto', String(letto)))
        .then(function() { return true; });
    });
}

// Template HTML iniziale per il campo Da Fare / Richieste. Viene
// inserito automaticamente quando un letto viene CREATO (Aggiungi Posto
// Letto) o SVUOTATO (Dimetti / Svuota Letto). Serve come scaletta di
// lavoro pre-impostata che l'utente può poi compilare.
//
// Struttura: tag <b> si CHIUDE prima del <br> così quando l'utente
// preme Enter dopo una label, il bold NON si propaga al testo nuovo.
// Tra le sezioni: due <br> per dare riga vuota dove scrivere.
var _TEMPLATE_DAFARE =
  '<b>DA FARE:</b><br><br>' +
  '<b>RICHIESTI/IN ATTESA:</b><br><br>' +
  '<b>ESEGUITI:</b><br><br>' +
  '<b>NOTE:</b><br>';
window._sbGetPazienteSingolo = _sbGetPazienteSingolo;
window._aggiornaCardDaPaziente = _aggiornaCardDaPaziente;
window._TEMPLATE_DAFARE = _TEMPLATE_DAFARE;

// Template HTML iniziale per il campo Note e Terapia. Stessa logica
// del template DaFare: vie di somministrazione in grassetto + riga
// vuota sotto ciascuna dove scrivere i farmaci (testo plain).
var _TEMPLATE_NOTETERAPIA =
  '<b>TERAPIA EV</b><br><br>' +
  '<b>TERAPIA OS</b><br><br>' +
  '<b>TERAPIA SC</b><br><br>' +
  '<b>AEROSOL</b><br>';
window._TEMPLATE_NOTETERAPIA = _TEMPLATE_NOTETERAPIA;

function _sbAggiungiLetto(numeroLetto) {
  return _q(_sb.from('consegne').select('letto').eq('letto', String(numeroLetto)).maybeSingle())
    .then(function(existing) {
      if (existing) return { success: false, message: 'Il letto esiste già!' };
      return _q(_sb.from('consegne').insert({
        letto:           String(numeroLetto),
        tipologia_letto: 'STANDARD',
        da_fare:         _TEMPLATE_DAFARE,
        note_terapia:    _TEMPLATE_NOTETERAPIA
      }))
        .then(function() { return { success: true, message: 'Letto aggiunto.' }; });
    });
}

function _sbEliminaLetto(numeroLetto) {
  return _q(_sb.from('consegne').select('nome,diagnosi').eq('letto', String(numeroLetto)).maybeSingle())
    .then(function(row) {
      if (!row) return { success: false, message: 'Letto non trovato.' };
      if ((row.nome || '').trim() || (row.diagnosi || '').trim()) {
        return { success: false, message: 'Il letto non è vuoto.' };
      }
      return _q(_sb.from('consegne').delete().eq('letto', String(numeroLetto)))
        .then(function() { return { success: true, message: 'Letto eliminato.' }; });
    });
}

function _sbDimettiLetto(numeroLetto) {
  // Reset di tutti i campi a stringa vuota, TRANNE da_fare che viene
  // preimpostato al template di scaletta (DA FARE / RICHIESTI / ESEGUITI / NOTE)
  var campiVuoti = {
    nome:'', eta:'', data_nascita:'', data_ricovero:'', diagnosi:'',
    note_terapia: _TEMPLATE_NOTETERAPIA, diaria:'', da_fare: _TEMPLATE_DAFARE,
    piano_terapeutico:'', esami_colturali:'',
    allergie:'', codice_sanitario:'', ossigeno:'', vitto:'',
    dimissibile:'', sesso:'',
    ultimo_aggiornamento: _oraStr(), updated_at: new Date().toISOString()
  };
  return _q(_sb.from('consegne').update(campiVuoti).eq('letto', String(numeroLetto)))
    .then(function() { return { success: true, message: 'Letto svuotato.' }; });
}

function _sbSpostaPaziente(lettoOrigine, lettoDestinazione) {
  return _q(_sb.from('consegne').select('*').in('letto', [String(lettoOrigine), String(lettoDestinazione)]))
    .then(function(rows) {
      var orig = rows.find(function(r) { return r.letto === String(lettoOrigine); });
      var dest = rows.find(function(r) { return r.letto === String(lettoDestinazione); });
      if (!orig || !dest) return { success: false, message: 'Letto non trovato.' };

      var campiPaziente = ['nome','eta','data_nascita','data_ricovero','diagnosi',
        'note_terapia','diaria','da_fare','piano_terapeutico','esami_colturali',
        'allergie','codice_sanitario','ossigeno','vitto','dimissibile','sesso'];

      var nuovoOrig = {}, nuovoDest = {};
      campiPaziente.forEach(function(c) {
        nuovoOrig[c] = dest[c] || '';
        nuovoDest[c] = orig[c] || '';
      });
      nuovoOrig.ultimo_aggiornamento = nuovoDest.ultimo_aggiornamento = _oraStr();
      nuovoOrig.updated_at = nuovoDest.updated_at = new Date().toISOString();

      return Promise.all([
        _q(_sb.from('consegne').update(nuovoOrig).eq('letto', String(lettoOrigine))),
        _q(_sb.from('consegne').update(nuovoDest).eq('letto', String(lettoDestinazione)))
      ]).then(function() { return { success: true, message: 'Spostati con successo!' }; });
    });
}

function _sbGetRiepilogo() {
  // LIGHT: SELECT solo i 5 campi necessari per il conteggio (no rich-text).
  // Riduce egress ~30x: da ~150KB a ~5KB per query.
  var pConteggi = _q(_sb.from('consegne').select('letto,nome,tipologia_letto,sesso,dimissibile'));
  // Conteggio ISOLATI senza scaricare la diaria (rich-text pesante):
  // filtro server-side sul marker del banner + head:true → PostgREST
  // restituisce SOLO l'header con il count, zero byte di body.
  var pIsolati = _sb.from('consegne')
    .select('letto', { count: 'exact', head: true })
    .like('diaria', '%isolamento-banner%')
    .neq('letto', 'NOTE')
    .then(function(res) { return res.error ? 0 : (res.count || 0); });
  var pRianim = _sb.from('consegne')
    .select('letto', { count: 'exact', head: true })
    .like('diaria', '%rianimazione-banner%')
    .neq('letto', 'NOTE')
    .then(function(res) { return res.error ? 0 : (res.count || 0); });

  return Promise.all([pConteggi, pIsolati, pRianim]).then(function(results) {
    var rows = results[0], isolati = results[1], rianimazione = results[2];
    var uomini = 0, donne = 0, indefinito = 0, vuoti = 0, inDimissione = 0;
    var tipologie = {};
    (rows || []).forEach(function(r) {
      if (r.letto === 'NOTE') return; // non contare NOTE come letto
      var hasPaziente = (r.nome || '').trim() !== '';
      if (!hasPaziente) { vuoti++; return; }
      var sesso = (r.sesso || '').toUpperCase();
      if (sesso === 'M') uomini++;
      else if (sesso === 'F') donne++;
      else indefinito++;
      var tip = (r.tipologia_letto || 'STANDARD').toUpperCase();
      tipologie[tip] = (tipologie[tip] || 0) + 1;
      if ((r.dimissibile || '').trim() !== '') inDimissione++;
    });
    return { uomini: uomini, donne: donne, indefinito: indefinito, tipologie: tipologie, vuoti: vuoti, inDimissione: inDimissione, isolati: isolati, rianimazione: rianimazione };
  });
}


// ══════════════════════════════════════════════════════════════
// LOCK MANAGEMENT
// ══════════════════════════════════════════════════════════════

function _sbGetLocks() {
  var now = Date.now();
  return _q(_sb.from('locks').select('*')).then(function(rows) {
    var locks = {};
    _lockState = {}; // aggiorna anche lo stato in-memory
    (rows || []).forEach(function(r) {
      if (now - Number(r.ts) <= LOCK_TTL_MS) {
        locks[r.letto] = { token: r.token, ts: Number(r.ts) };
        _lockState[r.letto] = { token: r.token };
      }
    });
    return locks;
  });
}

function _sbAcquistaLock(letto, token) {
  var k = String(letto);
  var now = Date.now();

  // Fast-fail in-memory: se c'è già un lock altrui noto → blocca subito senza query
  var existing = _lockState[k];
  if (existing && existing.token !== token) {
    return Promise.resolve({ success: false, blocked: true, message: 'Scheda in aggiornamento da altro utente.' });
  }

  // INSERT che non sovrascrive un lock altrui già presente nel DB
  // (ON CONFLICT DO NOTHING): il primo dei due utenti concorrenti inserisce,
  // il secondo viene ignorato silenziosamente.
  return _q(_sb.from('locks').upsert({ letto: k, token: token, ts: now }, { onConflict: 'letto', ignoreDuplicates: true }))
    .then(function() {
      // Verifica chi possiede effettivamente il lock nel DB
      // (risolve la race condition: entrambi gli utenti passavano il check
      //  in-memory, ma solo uno ha inserito per primo)
      return _q(_sb.from('locks').select('token, ts').eq('letto', k).maybeSingle());
    })
    .then(function(row) {
      if (!row) {
        // Riga sparita nel frattempo (rarissimo): upsert normale come fallback
        return _q(_sb.from('locks').upsert({ letto: k, token: token, ts: now }, { onConflict: 'letto' }))
          .then(function() { return { success: true, blocked: false }; });
      }
      if (row.token === token) {
        return { success: true, blocked: false };
      }
      // Lock altrui: è scaduto? → subentro
      if (now - Number(row.ts) > LOCK_TTL_MS) {
        return _q(_sb.from('locks').update({ token: token, ts: now }).eq('letto', k))
          .then(function() { return { success: true, blocked: false }; });
      }
      // Lock altrui attivo → bloccato
      return { success: false, blocked: true, message: 'Scheda in aggiornamento da altro utente.' };
    })
    .catch(function() { return { success: false, blocked: false, message: 'Errore server, riprova.' }; });
}

function _sbRilasciaLock(letto, token) {
  return _q(_sb.from('locks').delete().eq('letto', String(letto)).eq('token', token))
    .then(function() { return { success: true }; })
    .catch(function() { return { success: false }; });
}

function _sbAcquistaLockMultiplo(letti, token) {
  // Fast-fail in-memory: blocca se c'è qualsiasi lock attivo su uno dei letti
  var bloccatiMem = letti.filter(function(l) {
    return (typeof _lockState !== 'undefined') && !!_lockState[String(l)];
  });
  if (bloccatiMem.length > 0) {
    return Promise.resolve({ success: false, bloccati: bloccatiMem, message: 'Letti in uso.' });
  }
  var now = Date.now();
  var ks = letti.map(String);
  var upserts = ks.map(function(l) { return { letto: l, token: token, ts: now }; });

  // INSERT ignoreDuplicates: non sovrascrive lock altrui già presenti
  return _q(_sb.from('locks').upsert(upserts, { onConflict: 'letto', ignoreDuplicates: true }))
    .then(function() {
      // Verifica chi possiede i lock nel DB
      return _q(_sb.from('locks').select('letto, token').in('letto', ks));
    })
    .then(function(rows) {
      var rowMap = {};
      (rows || []).forEach(function(r) { rowMap[r.letto] = r.token; });
      var vinti  = ks.filter(function(l) { return rowMap[l] === token; });
      var persi  = ks.filter(function(l) { return rowMap[l] !== token; });
      if (persi.length > 0) {
        // Rilascia i lock acquisiti parzialmente, poi segnala fallimento
        if (vinti.length > 0) _sbRilasciaLockMultiplo(vinti, token);
        return { success: false, bloccati: persi, message: 'Letti in uso.' };
      }
      return { success: true };
    })
    .catch(function() { return { success: false, bloccati: [], message: 'Errore server.' }; });
}

function _sbRilasciaLockMultiplo(letti, token) {
  return _q(_sb.from('locks').delete().in('letto', letti.map(String)).eq('token', token))
    .then(function() { return { success: true }; })
    .catch(function() { return { success: false }; });
}


// ══════════════════════════════════════════════════════════════
// ARCHIVIO GIORNALIERO
// ══════════════════════════════════════════════════════════════

function _sbGetGiorniConservazione() {
  return _q(_sb.from('impostazioni').select('valore').eq('chiave', 'GIORNI_ARCHIVIO').maybeSingle())
    .then(function(row) { return row ? (parseInt(row.valore, 10) || 90) : 90; });
}

function _sbSalvaGiorniConservazione(giorni) {
  return _q(_sb.from('impostazioni').upsert(
    { chiave: 'GIORNI_ARCHIVIO', valore: String(giorni) },
    { onConflict: 'chiave' }
  )).then(function() { return { success: true, giorni: giorni }; });
}

var BACKUP_INTERVALLO_MS = 1 * 60 * 60 * 1000; // 1 ora (era 6h)

function _sbArchiviaGiornoCorrente() {
  var dataStr = _oggiStr();
  var now = Date.now();

  // ── Step 1: leggi ULTIMO_BACKUP ────────────────────────────────────────────
  return _q(_sb.from('impostazioni').select('valore').eq('chiave', 'ULTIMO_BACKUP').maybeSingle())
    .then(function(row) {
      var ultimoBackup = row ? Number(row.valore) : 0;
      if (now - ultimoBackup < BACKUP_INTERVALLO_MS) {
        return { inCorso: false }; // Meno di 6h → skip, nessuna query extra
      }

      // ── Step 2: CAS — prenota il backup in modo atomico ──────────────────
      // Aggiorna ULTIMO_BACKUP a `now` SOLO SE ha ancora il valore letto prima.
      // Se due utenti si collegano insieme, solo uno avrà la riga con
      // valore = ultimoBackup: il secondo troverà 0 righe aggiornate e si ferma.
      var oldValore = row ? String(row.valore) : '0';
      return _q(
        _sb.from('impostazioni')
          .update({ valore: String(now) })
          .eq('chiave', 'ULTIMO_BACKUP')
          .eq('valore', oldValore)    // CAS: aggiorna solo se è ancora il vecchio valore
          .select()                   // NECESSARIO: senza .select() Supabase ritorna null
                                      // anche se la riga è stata aggiornata → falso negativo
      ).then(function(updated) {
        // updated = array con la riga aggiornata se CAS ha avuto successo, [] altrimenti
        if (!updated || updated.length === 0) {
          // Un altro client ha già riservato il backup (valore era già cambiato)
          return { inCorso: false };
        }

        // ── Step 3: esegui il backup (siamo gli unici autorizzati) ──────────
        var _pazientiBackup;
        return _sbGetPazienti().then(function(pazienti) {
          _pazientiBackup = pazienti;
          return _q(_sb.from('archivio').insert({
            data_str: dataStr,
            ts: now,
            dati: pazienti
          }));
        }).then(function() {
          // Log del backup DB orario riuscito: include info sul token Drive
          // così, incrociando con i log 'drive-backup', si capisce se il PC
          // che ha eseguito il backup aveva o no la sessione Drive attiva.
          if (typeof window._log === 'function') {
            try {
              window._log('info', 'backup-orario-db',
                'Backup orario su DB completato (' + _pazientiBackup.length + ' letti)',
                'token_drive_disponibile=' + (!!window._googleDriveToken));
            } catch(e) {}
          }
          // ── Step 4: backup Drive (fire-and-forget) + pulizia archivio ──────
          _driveBackupConsegne(_pazientiBackup, now);
          return _sbGetGiorniConservazione().then(function(giorni) {
            var limit = now - (giorni * 86400000);
            _q(_sb.from('archivio').delete().lt('ts', limit)).catch(function() {});
            return { inCorso: false };
          });
        }).catch(function() {
          // Se il backup fallisce, ripristina ULTIMO_BACKUP al vecchio valore
          // così il prossimo caricamento ci riprova
          _q(_sb.from('impostazioni').update({ valore: oldValore })
            .eq('chiave', 'ULTIMO_BACKUP')).catch(function() {});
          return { inCorso: false };
        });
      });
    })
    .catch(function() { return { inCorso: false }; });
}

// ══════════════════════════════════════════════════════════════
// MAIL DIMISSIONI — helper per persistenza template + check
// ══════════════════════════════════════════════════════════════

// Chiavi nella tabella `impostazioni` per i template della mail dimissioni
var _MAIL_DIM_KEYS = {
  destinatari: 'MAIL_DIMISSIONI_DESTINATARI',
  oggetto:     'MAIL_DIMISSIONI_OGGETTO',
  corpo:       'MAIL_DIMISSIONI_CORPO',
  chiusura:    'MAIL_DIMISSIONI_CHIUSURA'
};

// Sostituisce ogni occorrenza di data DD/MM/YYYY o D/M/YYYY in `str`
// con la data di domani. Utile per mantenere l'oggetto sempre aggiornato
// se l'utente ha salvato un template tipo "Dimissioni del 25/05/2026":
// alla prossima apertura la data diventa quella di domani automaticamente.
function _aggiornaDataInStringa(str) {
  if (!str) return str;
  var dom = _domaniMezzanotte();
  function pad(n) { return String(n).padStart(2, '0'); }
  var dataNuova = pad(dom.getDate()) + '/' + pad(dom.getMonth()+1) + '/' + dom.getFullYear();
  return String(str).replace(/\b\d{1,2}\/\d{1,2}\/\d{4}\b/g, dataNuova);
}

// Genera oggetto di default dinamico: "Dimissioni previste per il DD/MM/YYYY"
function _oggettoMailDefault() {
  var dom = _domaniMezzanotte();
  function pad(n) { return String(n).padStart(2, '0'); }
  var data = pad(dom.getDate()) + '/' + pad(dom.getMonth()+1) + '/' + dom.getFullYear();
  return 'Dimissioni previste per il ' + data;
}

// Carica i template salvati. Restituisce { destinatari, oggetto, corpo, chiusura }.
// Default dinamici se non presenti nel DB. Se l'oggetto salvato contiene una
// data, viene aggiornata a domani per riusabilità giorno dopo giorno.
function _sbCaricaTemplateMail() {
  return _q(_sb.from('impostazioni').select('chiave,valore')
    .in('chiave', [_MAIL_DIM_KEYS.destinatari, _MAIL_DIM_KEYS.oggetto,
                   _MAIL_DIM_KEYS.corpo, _MAIL_DIM_KEYS.chiusura]))
    .then(function(rows) {
      var map = {};
      (rows || []).forEach(function(r) { map[r.chiave] = r.valore || ''; });
      var oggettoSalvato = map[_MAIL_DIM_KEYS.oggetto] || '';
      return {
        destinatari: map[_MAIL_DIM_KEYS.destinatari] || '',
        oggetto:     oggettoSalvato
                       ? _aggiornaDataInStringa(oggettoSalvato)
                       : _oggettoMailDefault(),
        corpo:       map[_MAIL_DIM_KEYS.corpo]       ||
          'Buongiorno,\n\nsi comunica che per la giornata di domani sono previste le seguenti dimissioni dal reparto MEU 11N:',
        chiusura:    map[_MAIL_DIM_KEYS.chiusura]    ||
          'Cordiali saluti.'
      };
    });
}

// Salva i template (upsert su impostazioni)
function _sbSalvaTemplateMail(destinatari, oggetto, corpo, chiusura) {
  var rows = [
    { chiave: _MAIL_DIM_KEYS.destinatari, valore: destinatari || '' },
    { chiave: _MAIL_DIM_KEYS.oggetto,     valore: oggetto || '' },
    { chiave: _MAIL_DIM_KEYS.corpo,       valore: corpo || '' },
    { chiave: _MAIL_DIM_KEYS.chiusura,    valore: chiusura || '' }
  ];
  return _q(_sb.from('impostazioni').upsert(rows, { onConflict: 'chiave' }));
}

// Verifica se OGGI è già stata inviata una mail dimissioni.
// Cerca in `logs` tipo='mail-dimissioni' con ts >= mezzanotte di oggi.
// Restituisce { gia: bool, ultimo: log|null }.
function _sbMailDimissioniInviataOggi() {
  var oggiMezzanotte = new Date();
  oggiMezzanotte.setHours(0, 0, 0, 0);
  var soglia = oggiMezzanotte.getTime();
  return _q(_sb.from('logs')
    .select('ts,device_name,messaggio,descrizione')
    .eq('tipo', 'mail-dimissioni')
    .gte('ts', soglia)
    .order('ts', { ascending: false })
    .limit(1))
    .then(function(rows) {
      if (rows && rows.length) return { gia: true, ultimo: rows[0] };
      return { gia: false, ultimo: null };
    });
}

// Helper: data di DOMANI in formato Date (00:00:00)
function _domaniMezzanotte() {
  var d = new Date();
  d.setHours(0, 0, 0, 0);
  d.setDate(d.getDate() + 1);
  return d;
}

// Verifica se la stringa Dimissibile rappresenta esattamente DOMANI
// (data parsata col parser flex deve essere uguale a domani).
function _isDimissibileDomani(strDim) {
  if (!strDim) return false;
  var s = String(strDim).trim().toLowerCase();
  // Match testuale "domani"
  if (s === 'domani') return true;
  // Match data DD/MM/YYYY (usa lo stesso parser flex del dim modal)
  // Se _parseDimData è disponibile, usalo; altrimenti regex inline.
  var parsed = null;
  if (typeof window._parseDimData === 'function') {
    parsed = window._parseDimData(strDim);
  }
  if (!parsed) parsed = strDim;
  var m = String(parsed).match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (!m) return false;
  var d = new Date(parseInt(m[3],10), parseInt(m[2],10) - 1, parseInt(m[1],10));
  d.setHours(0, 0, 0, 0);
  var dom = _domaniMezzanotte();
  return d.getTime() === dom.getTime();
}

// Costruisce e invia un'email via Gmail API for-conto-utente.
// Richiede token con scope https://www.googleapis.com/auth/gmail.send
function _gmailInviaMail(token, fromEmail, destinatariArr, oggetto, htmlBody) {
  if (!token) return Promise.reject(new Error('Token Gmail mancante'));
  if (!destinatariArr || !destinatariArr.length) return Promise.reject(new Error('Nessun destinatario'));

  // Codifica oggetto in RFC 2047 (per accenti, lettere non ASCII)
  function encodeHeader(s) {
    var hasNonAscii = /[^\x00-\x7F]/.test(s || '');
    if (!hasNonAscii) return s || '';
    return '=?UTF-8?B?' + btoa(unescape(encodeURIComponent(String(s)))) + '?=';
  }

  var headers = [
    'From: ' + (fromEmail || ''),
    'To: ' + destinatariArr.join(', '),
    'Subject: ' + encodeHeader(oggetto || ''),
    'MIME-Version: 1.0',
    'Content-Type: text/html; charset=UTF-8',
    'Content-Transfer-Encoding: base64'
  ];
  // Body in Base64 (max 76 char per riga è standard ma Gmail accetta unbroken)
  var bodyB64 = btoa(unescape(encodeURIComponent(htmlBody || '')));
  // Chunked a 76 char (più compatibile)
  bodyB64 = bodyB64.replace(/(.{76})/g, '$1\r\n');
  // RFC 5322: headers e body separati da UNA RIGA VUOTA → \r\n\r\n
  // (un \r\n chiude l'ultimo header, l'altro è la linea vuota richiesta).
  // Senza questo separatore, Gmail interpreta il body come header malformato
  // e la mail arriva con oggetto/destinatari OK ma corpo vuoto.
  var rawMessage = headers.join('\r\n') + '\r\n\r\n' + bodyB64;

  // URL-safe Base64 (Gmail API richiede questo)
  var raw = btoa(unescape(encodeURIComponent(rawMessage)))
    .replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');

  return fetch('https://gmail.googleapis.com/gmail/v1/users/me/messages/send', {
    method: 'POST',
    headers: {
      'Authorization': 'Bearer ' + token,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({ raw: raw })
  }).then(function(r) {
    if (!r.ok) {
      return r.json().catch(function(){return {};}).then(function(b) {
        var em = (b.error && b.error.message) || 'HTTP ' + r.status;
        throw new Error(em);
      });
    }
    return r.json();
  });
}

window._sbCaricaTemplateMail        = _sbCaricaTemplateMail;
window._sbSalvaTemplateMail         = _sbSalvaTemplateMail;
window._sbMailDimissioniInviataOggi = _sbMailDimissioniInviataOggi;
window._domaniMezzanotte            = _domaniMezzanotte;
window._isDimissibileDomani         = _isDimissibileDomani;
window._gmailInviaMail              = _gmailInviaMail;


// Backup forzato: inserisce subito un record in archivio e aggiorna ULTIMO_BACKUP.
// Usato dal backup manuale; bypassa il controllo delle 6h.
function _sbBackupForzato() {
  var dataStr = _oggiStr();
  var now = Date.now();
  return _sbGetPazienti().then(function(pazienti) {
    return _q(_sb.from('archivio').insert({ data_str: dataStr, ts: now, dati: pazienti }))
      .then(function() {
        return _q(_sb.from('impostazioni')
          .upsert({ chiave: 'ULTIMO_BACKUP', valore: String(now) }, { onConflict: 'chiave' }));
      }).then(function() {
        _driveBackupConsegne(pazienti, now); // fire-and-forget
        return { success: true, ts: now, count: pazienti.length };
      });
  });
}

// Ripristina i dati di uno o più letti dal backup.
// Aggiorna solo letti che esistono nel DB corrente; salta quelli rimossi.
function _sbRipristinaLetti(pazientiDaRipristinare) {
  return _q(_sb.from('consegne').select('letto')).then(function(rows) {
    var lettiEsistenti = {};
    (rows || []).forEach(function(r) { lettiEsistenti[r.letto] = true; });
    var validi = (pazientiDaRipristinare || []).filter(function(p) {
      return lettiEsistenti[String(p.Letto)];
    });
    if (validi.length === 0) return { success: true, count: 0 };
    var promises = validi.map(function(p) {
      var row = _toDb(p);
      row.ultimo_aggiornamento = _oraStr();
      row.updated_at = new Date().toISOString();
      return _q(_sb.from('consegne').update(row).eq('letto', String(p.Letto)));
    });
    return Promise.all(promises).then(function() { return { success: true, count: validi.length }; });
  });
}

function _sbGetGiorniArchivio() {
  // Ritorna array di date uniche (YYYY-MM-DD) con almeno un backup
  return _q(_sb.from('archivio').select('data_str').order('data_str', { ascending: false }))
    .then(function(rows) {
      var seen = {};
      return (rows || []).filter(function(r) {
        if (seen[r.data_str]) return false;
        seen[r.data_str] = true;
        return true;
      }).map(function(r) { return r.data_str; });
    });
}

function _sbGetTimestampsGiorno(dataStr) {
  // Ritorna array di epoch-ms (come string) ordinati dal più recente
  return _q(_sb.from('archivio').select('ts').eq('data_str', dataStr).order('ts', { ascending: false }))
    .then(function(rows) {
      return (rows || []).map(function(r) { return String(r.ts); });
    });
}

function _sbGetDatiArchivioGiorno(key) {
  var sKey = String(key || '');
  // Supporta sia epoch ms numerico (nuovo) sia data stringa YYYY-MM-DD (legacy)
  if (/^\d{10,13}$/.test(sKey)) {
    return _q(_sb.from('archivio').select('dati,ts').eq('ts', Number(sKey)).maybeSingle())
      .then(function(row) {
        return row ? { pazienti: row.dati || [], timestamp: row.ts } : null;
      });
  }
  // Legacy: prende il backup più recente per quella data
  return _q(_sb.from('archivio').select('dati,ts').eq('data_str', sKey.substring(0, 10)).order('ts', { ascending: false }))
    .then(function(rows) {
      var row = rows && rows[0];
      return row ? { pazienti: row.dati || [], timestamp: row.ts } : null;
    });
}


// ══════════════════════════════════════════════════════════════
// TIPOLOGIE
// ══════════════════════════════════════════════════════════════

function _sbGetColoriTipologie() {
  return _q(_sb.from('tipologie').select('nome,colore')).then(function(rows) {
    var mappa = {};
    (rows || []).forEach(function(r) { mappa[r.nome] = r.colore || ''; });
    return mappa;
  });
}

function _sbSalvaColoriTipologie(mappa) {
  var rows = Object.keys(mappa || {}).filter(Boolean).map(function(nome) {
    return { nome: nome, colore: mappa[nome] || '' };
  });
  if (!rows.length) return Promise.resolve({ ok: true });
  return _q(_sb.from('tipologie').upsert(rows, { onConflict: 'nome' }))
    .then(function() { return { ok: true }; });
}

function _sbSalvaTipologieBatch(modifiche) {
  if (!modifiche || !modifiche.length) return Promise.resolve({ ok: true });
  var promesse = modifiche.map(function(m) {
    var nomeOld = (m.nomeOld || '').trim();
    var nomeNew = (m.nomeNew || '').trim().toUpperCase();
    var colore  = m.colore || '';
    if (!nomeNew) return Promise.resolve();
    if (nomeOld && nomeOld !== nomeNew) {
      // Rinomina: inserisce con nuovo nome, cancella vecchio, aggiorna consegne
      return _q(_sb.from('tipologie').upsert({ nome: nomeNew, colore: colore }, { onConflict: 'nome' }))
        .then(function() {
          return Promise.all([
            _q(_sb.from('tipologie').delete().eq('nome', nomeOld)),
            _q(_sb.from('consegne').update({ tipologia_letto: nomeNew }).eq('tipologia_letto', nomeOld))
          ]);
        });
    }
    // Nuova tipologia o aggiornamento colore
    return _q(_sb.from('tipologie').upsert({ nome: nomeNew, colore: colore }, { onConflict: 'nome' }));
  });
  return Promise.all(promesse).then(function() { return { ok: true }; });
}

function _sbEliminaTipologia(nome, force) {
  return _q(_sb.from('consegne').select('letto').eq('tipologia_letto', nome)).then(function(rows) {
    if (rows && rows.length > 0 && !force) {
      return { success: false, inUso: true, count: rows.length };
    }
    var p = _q(_sb.from('tipologie').delete().eq('nome', nome));
    if (rows && rows.length > 0) {
      p = p.then(function() {
        return _q(_sb.from('consegne').update({ tipologia_letto: 'STANDARD' }).eq('tipologia_letto', nome));
      });
    }
    return p.then(function() { return { success: true }; });
  });
}

function _sbCambiaTipologiaALetto(letto, nuovaTipologia) {
  // updated_at incluso esplicitamente (belt & braces: il trigger DB
  // trg_consegne_touch_updated_at lo garantisce comunque server-side).
  // Senza bump, il cambio tipologia era INVISIBILE ai PC in modalità
  // emergenza: il check leggero confronta MAX(updated_at).
  return _q(_sb.from('consegne').update({
    tipologia_letto: nuovaTipologia,
    updated_at: new Date().toISOString()
  }).eq('letto', String(letto)))
    .then(function() { return { success: true }; });
}

function _sbGetTipologieConfigurate() {
  return _sbGetColoriTipologie().then(function(mappa) {
    var result = {};
    Object.keys(mappa).forEach(function(k) { if (k) result[k] = mappa[k] || null; });
    return result;
  });
}

function _sbGetDatiLettiConTipologia() {
  return _sbGetPazienti().then(function(pazienti) {
    var mappa = {};
    pazienti.forEach(function(p) {
      var tip = (p.TipologiaLetto || 'STANDARD').toUpperCase();
      if (!mappa[tip]) mappa[tip] = [];
      mappa[tip].push({ letto: p.Letto, nome: p.Nome || '' });
    });
    return mappa;
  });
}

function _sbGetTipologieLettiBed() {
  // Ritorna un array ordinato di tipologie UNICHE presenti in almeno 1 letto (esclusa NOTE)
  return _q(_sb.from('consegne').select('letto,tipologia_letto')).then(function(rows) {
    var set = {};
    (rows || []).forEach(function(r) {
      if (r.letto === 'NOTE') return;
      var t = (r.tipologia_letto || 'STANDARD').trim().toUpperCase();
      if (t) set[t] = true;
    });
    return Object.keys(set).sort();
  });
}


// ══════════════════════════════════════════════════════════════
// LINK UTILI
// ══════════════════════════════════════════════════════════════

function _sbGetLinkUtili() {
  return _q(_sb.from('link_utili').select('*').order('id')).then(function(rows) {
    return (rows || []).map(function(r) { return { id: r.id, nome: r.nome, url: r.url }; });
  });
}

function _sbAggiungiLink(nome, url) {
  return _q(_sb.from('link_utili').insert({ nome: nome, url: url }).select().single())
    .then(function(row) { return { success: true, link: row }; });
}

function _sbModificaLink(indice, nome, url) {
  // indice = posizione 0-based nell'array restituito da getLinkUtili
  return _sbGetLinkUtili().then(function(links) {
    var link = links[indice];
    if (!link) return { success: false };
    return _q(_sb.from('link_utili').update({ nome: nome, url: url }).eq('id', link.id))
      .then(function() { return { success: true }; });
  });
}

function _sbEliminaLink(indice) {
  return _sbGetLinkUtili().then(function(links) {
    var link = links[indice];
    if (!link) return { success: false };
    return _q(_sb.from('link_utili').delete().eq('id', link.id))
      .then(function() { return { success: true }; });
  });
}


// ══════════════════════════════════════════════════════════════
// IMPOSTAZIONI
// ══════════════════════════════════════════════════════════════

function _sbOttieniNomeReparto() {
  return _q(_sb.from('impostazioni').select('valore').eq('chiave', 'NOME_REPARTO').maybeSingle())
    .then(function(row) { return { nome: row ? row.valore : 'Consegne Reparto' }; });
}

function _sbSalvaNomeReparto(nuovoNome) {
  return _q(_sb.from('impostazioni').upsert({ chiave: 'NOME_REPARTO', valore: nuovoNome }, { onConflict: 'chiave' }))
    .then(function() { return { nome: nuovoNome }; });
}

function _sbOttieniAccountLogin() {
  return _q(_sb.from('impostazioni').select('valore').eq('chiave', 'ACCOUNT_LOGIN').maybeSingle())
    .then(function(row) { return row ? row.valore : ''; });
}


// ══════════════════════════════════════════════════════════════
// GOOGLE DRIVE BACKUP
// ══════════════════════════════════════════════════════════════

// Helper Base64 UTF-8 safe (gestisce caratteri accentati e Unicode).
// Usato per embedare il JSON dei dati nel Google Doc di backup senza
// problemi di escaping HTML.
function _b64encodeUtf8(s) {
  try {
    return btoa(unescape(encodeURIComponent(String(s || ''))));
  } catch(e) { return ''; }
}
function _b64decodeUtf8(b64) {
  try {
    // Tollera whitespace/newline introdotti dal Doc
    var clean = String(b64 || '').replace(/[\s\n\r\t]+/g, '');
    return decodeURIComponent(escape(atob(clean)));
  } catch(e) { return ''; }
}
window._b64encodeUtf8 = _b64encodeUtf8;
window._b64decodeUtf8 = _b64decodeUtf8;

// Cerca o crea la cartella "BACKUP CONSEGNE EMERGENZA" nella root dell'utente.
// Con scope drive.file, files.list restituisce solo le cartelle create da quest'app.
function _driveGetOrCreateFolder(nome) {
  var token = window._googleDriveToken;
  if (!token) return Promise.reject(new Error('No Drive token'));
  var q = "name='" + nome.replace(/'/g, "\\'") + "'" +
          " and mimeType='application/vnd.google-apps.folder'" +
          " and trashed=false";
  return fetch('https://www.googleapis.com/drive/v3/files?q=' + encodeURIComponent(q) + '&fields=files(id)', {
    headers: { 'Authorization': 'Bearer ' + token }
  }).then(function(r) { return r.json(); })
  .then(function(data) {
    if (data.files && data.files.length > 0) return data.files[0].id;
    // Non trovata → crea nella root
    return fetch('https://www.googleapis.com/drive/v3/files', {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
      body: JSON.stringify({ name: nome, mimeType: 'application/vnd.google-apps.folder' })
    }).then(function(r) { return r.json(); }).then(function(f) {
      if (f.error) throw new Error(f.error.message);
      return f.id;
    });
  });
}

// Lista tutti i Google Doc nella cartella backup ordinati per createdTime DESC
// (dal più recente al più vecchio). Usato dal modal "Importa da backup Drive"
// come safety net se Supabase è temporaneamente down.
// Ritorna array di { id, name, createdTime, modifiedTime }.
function _driveListaBackupFiles(folderId, limit) {
  var token = window._googleDriveToken;
  if (!token) return Promise.reject(new Error('No Drive token'));
  if (!folderId) return Promise.reject(new Error('No folder ID'));
  var q = "'" + folderId + "' in parents" +
          " and mimeType='application/vnd.google-apps.document'" +
          " and trashed=false";
  var url = 'https://www.googleapis.com/drive/v3/files?' +
            'q=' + encodeURIComponent(q) +
            '&orderBy=createdTime%20desc' +
            '&pageSize=' + (limit || 100) +
            '&fields=files(id,name,createdTime,modifiedTime)';
  return fetch(url, {
    headers: { 'Authorization': 'Bearer ' + token }
  }).then(function(r) {
    if (!r.ok) {
      return r.json().catch(function(){return {};}).then(function(body) {
        var msg = (body.error && body.error.message) || 'HTTP ' + r.status;
        throw new Error(msg);
      });
    }
    return r.json();
  }).then(function(data) {
    return (data && data.files) ? data.files : [];
  });
}
window._driveListaBackupFiles = _driveListaBackupFiles;
window._driveGetOrCreateFolder = _driveGetOrCreateFolder;

// Elimina i file Drive nella cartella più vecchi di cutoffMs (basato su createdTime di Drive).
function _driveEliminaVecchi(folderId, cutoffMs) {
  var token = window._googleDriveToken;
  if (!token) return Promise.resolve();
  var cutoffISO = new Date(cutoffMs).toISOString();
  var q = "'" + folderId + "' in parents" +
          " and createdTime < '" + cutoffISO + "'" +
          " and trashed=false";
  return fetch('https://www.googleapis.com/drive/v3/files?q=' + encodeURIComponent(q) + '&fields=files(id,name)', {
    headers: { 'Authorization': 'Bearer ' + token }
  }).then(function(r) { return r.json(); })
  .then(function(data) {
    var files = data.files || [];
    if (files.length > 0) console.log('[Drive cleanup] Eliminazione ' + files.length + ' file vecchi');
    return Promise.all(files.map(function(f) {
      return fetch('https://www.googleapis.com/drive/v3/files/' + f.id, {
        method: 'DELETE',
        headers: { 'Authorization': 'Bearer ' + token }
      }).catch(function() {});
    }));
  }).catch(function() {});
}

// Crea un file di testo plain in una cartella Drive (multipart upload)
// Crea un Google Doc nativo su Drive, imposta A4 landscape e adatta le colonne
// alla larghezza reale del foglio via Docs API batchUpdate.
//
// Flusso:
// 1. Upload HTML → Google Doc (il converter usa portrait per le larghezze)
// 2. GET struttura doc (solo posizioni tabelle)
// 3. batchUpdate: imposta A4 landscape + larghezze colonne corrette
//
// A4 landscape con margini 20mm: larghezza utile = 841.89 - 2×56.69 = 728.51pt
// Colonne card (3 col): C1=17% (124pt) | C2=63% (459pt) | C3=20% (146pt)
function _driveCreaGoogleDoc(nome, htmlContent, folderId) {
  var token = window._googleDriveToken;
  if (!token) return Promise.reject(new Error('No Drive token'));

  // Step 1: carica HTML come Google Doc
  var boundary = 'app_consegne_gdoc_boundary';
  var meta = JSON.stringify({ name: nome, mimeType: 'application/vnd.google-apps.document', parents: [folderId] });
  var uploadBody = '--' + boundary + '\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n' + meta + '\r\n' +
                   '--' + boundary + '\r\nContent-Type: text/html; charset=UTF-8\r\n\r\n' + htmlContent + '\r\n' +
                   '--' + boundary + '--';

  return fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart', {
    method: 'POST',
    headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'multipart/related; boundary="' + boundary + '"' },
    body: uploadBody
  })
  .then(function(r) { return r.json(); })
  .then(function(file) {
    if (!file || !file.id) return file;
    var docId = file.id;

    // Step 2: legge la struttura del doc (solo startIndex delle tabelle)
    return fetch('https://docs.googleapis.com/v1/documents/' + docId +
                 '?fields=body.content(startIndex%2Ctable%2Fcolumns)', {
      headers: { 'Authorization': 'Bearer ' + token }
    })
    .then(function(r) { return r.json(); })
    .then(function(doc) {
      // Step 3: costruisce un unico batchUpdate con:
      //   a) orientamento A4 landscape + margini
      //   b) larghezze colonne per ogni tabella
      var TOTAL = 728; // pt utili in landscape (841.89 - 2×56.69)
      var COL3  = [Math.round(TOTAL * 0.17), Math.round(TOTAL * 0.63),
                   TOTAL - Math.round(TOTAL * 0.17) - Math.round(TOTAL * 0.63)];  // 124 + 459 + 145 = 728
      var COL2  = [Math.round(TOTAL * 0.40), TOTAL - Math.round(TOTAL * 0.40)];   // 291 + 437 = 728

      var requests = [{
        updateDocumentStyle: {
          documentStyle: {
            pageSize: { width: { magnitude: 841.89, unit: 'PT' }, height: { magnitude: 595.28, unit: 'PT' } },
            marginTop:    { magnitude: 42.52, unit: 'PT' },
            marginBottom: { magnitude: 42.52, unit: 'PT' },
            marginLeft:   { magnitude: 56.69, unit: 'PT' },
            marginRight:  { magnitude: 56.69, unit: 'PT' }
          },
          fields: 'pageSize,marginTop,marginBottom,marginLeft,marginRight'
        }
      }];

      // Aggiunge updateTableColumnProperties per ogni tabella trovata
      ((doc.body && doc.body.content) || []).forEach(function(elem) {
        if (!elem.table || elem.startIndex === undefined) return;
        var numCols = elem.table.columns || 3;
        var widths  = numCols === 3 ? COL3 : COL2;
        for (var ci = 0; ci < numCols; ci++) {
          requests.push({
            updateTableColumnProperties: {
              tableStartLocation: { index: elem.startIndex },
              columnIndices: [ci],
              tableColumnProperties: {
                widthType: 'FIXED_WIDTH',
                width: { magnitude: widths[ci] || Math.round(TOTAL / numCols), unit: 'PT' }
              },
              fields: 'widthType,width'
            }
          });
        }
      });

      return fetch('https://docs.googleapis.com/v1/documents/' + docId + ':batchUpdate', {
        method: 'POST',
        headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
        body: JSON.stringify({ requests: requests })
      });
    })
    .then(function() { return file; })
    .catch(function(e) {
      console.warn('[Drive] Post-processing fallito (non bloccante):', e.message || e);
      return file;
    });
  });
}

// ══════════════════════════════════════════════════════════════════════
// SCHEMA DINAMICO PER BACKUP/IMPORT DRIVE — derivato da `_campi`
// ──────────────────────────────────────────────────────────────────────
// Ogni paziente nel backup è una TABELLA 4-COLONNE che replica il
// layout della card paziente nell'app (info | diag | diaria | terapia).
// Il parser dell'import cerca le coppie label/valore dentro le celle:
// non importa in che colonna stanno, basta che la label sia presente.
//
// SISTEMA DATA-DRIVEN da `_campi`:
// - Per un nuovo campo, basta aggiungerlo a `_campi`: appare nel backup
//   nel gruppo "info" con label = nome JS (default).
// - Per personalizzare label + colonna + tipo HTML, aggiungere una
//   entry a `_DRIVE_LAYOUT` con { campo, gruppo, label, isHtml? }.
// - Per escludere un campo dal backup, aggiungere a `_DRIVE_EXCLUDE_FIELDS`.
//
// Quindi: NESSUNA modifica al codice di rendering/parsing per nuovi campi.
// ══════════════════════════════════════════════════════════════════════

// Definizione layout campo per campo: gruppo (colonna), label visibile
// nel Doc, isHtml (true = preserva HTML rich-text). L'ORDINE in questo
// array determina anche l'ordine visivo dei campi dentro ogni gruppo.
// Gruppi disponibili: 'info' (col sinistra), 'diag' (col centro-sx, una
// riga per campo), 'diaria' (col centro-grande), 'terapia' (col destra).
var _DRIVE_LAYOUT = [
  // ── Colonna INFO (sinistra) — replica .alt-col-info dell'app ──
  { campo: 'Letto',             gruppo: 'info',    label: 'LETTO' },
  { campo: 'TipologiaLetto',    gruppo: 'info',    label: 'TIPOLOGIA' },
  { campo: 'Nome',              gruppo: 'info',    label: 'NOME' },
  { campo: 'Sesso',             gruppo: 'info',    label: 'SESSO' },
  { campo: 'Allergie',          gruppo: 'info',    label: '⚠ ALLERGIE',       isHtml: true },
  { campo: 'DataNascita',       gruppo: 'info',    label: 'Data di Nascita' },
  { campo: 'Eta',               gruppo: 'info',    label: 'Età' },
  { campo: 'DataRicovero',      gruppo: 'info',    label: 'Ricovero' },
  { campo: 'CodiceSanitario',   gruppo: 'info',    label: 'C.S.' },
  { campo: 'Ossigeno',          gruppo: 'info',    label: 'Ossigeno' },
  { campo: 'Vitto',             gruppo: 'info',    label: 'Vitto' },
  { campo: 'Dimissibile',       gruppo: 'info',    label: 'Dimissibile' },
  // ── Colonna DIAG (centro-sinistra) — replica .alt-col-diag ──
  { campo: 'Diagnosi',          gruppo: 'diag',    label: 'DIAGNOSI / MOTIVO RICOVERO', isHtml: true },
  { campo: 'PianoTerapeutico',  gruppo: 'diag',    label: 'PIANO DI CURA',              isHtml: true },
  { campo: 'EsamiColturali',    gruppo: 'diag',    label: 'ESAMI COLTURALI',            isHtml: true },
  { campo: 'DaFare',            gruppo: 'diag',    label: 'DA FARE / RICHIESTE',        isHtml: true },
  // ── Colonna DIARIA (centro, grande) — replica .alt-col-diaria ──
  { campo: 'Diaria',            gruppo: 'diaria',  label: 'DIARIA ED EPICRISI',         isHtml: true },
  // ── Colonna TERAPIA (destra) — replica .alt-col-terapia ──
  { campo: 'NoteTerapia',       gruppo: 'terapia', label: 'NOTE E TERAPIA',             isHtml: true }
];

// Campi da NON includere nel backup (campi tecnici di sistema)
var _DRIVE_EXCLUDE_FIELDS = {
  'UltimoAggiornamento': true
};

// Costruisce dinamicamente lo schema gruppi {info:[...], diag:[...], diaria:[...], terapia:[...]}
// dai campi attualmente registrati in `_campi`, usando `_DRIVE_LAYOUT` per
// gli override. I campi nuovi non listati vanno in 'info' con label = nome JS.
function _getDriveSchema() {
  var groups = { info: [], diag: [], diaria: [], terapia: [] };
  if (typeof _campi !== 'object' || !_campi) return groups;

  // Mappa rapida campo → spec layout
  var bySpec = {};
  _DRIVE_LAYOUT.forEach(function(s) { bySpec[s.campo] = s; });

  // Prima: campi nell'ordine del _DRIVE_LAYOUT (mantengono l'ordine voluto)
  _DRIVE_LAYOUT.forEach(function(spec) {
    if (_DRIVE_EXCLUDE_FIELDS[spec.campo]) return;
    if (!_campi[spec.campo]) return; // il campo non è più in _campi
    groups[spec.gruppo || 'info'].push({
      campo:  spec.campo,
      label:  spec.label || spec.campo,
      isHtml: !!spec.isHtml
    });
  });

  // Poi: campi presenti in _campi ma NON in _DRIVE_LAYOUT (aggiunti dopo)
  // → vanno automaticamente in 'info' con label = nome JS.
  Object.keys(_campi).forEach(function(f) {
    if (_DRIVE_EXCLUDE_FIELDS[f]) return;
    if (bySpec[f]) return; // già processato sopra
    groups.info.push({ campo: f, label: f, isHtml: false });
  });

  return groups;
}

// Espone tutto a window per uso da app.js (parser import)
window._DRIVE_LAYOUT          = _DRIVE_LAYOUT;
window._DRIVE_EXCLUDE_FIELDS  = _DRIVE_EXCLUDE_FIELDS;
window._getDriveSchema        = _getDriveSchema;

// ──────────────────────────────────────────────────────────────────────
// HELPER per rendering Doc — replicano stili della card app (.alt-row)
// ──────────────────────────────────────────────────────────────────────

// Stile celle: bordo, padding
var _D_B  = '1.5pt solid #333';
var _D_Bi = '1pt solid #ccc';

// Escape minimo per i campi plain text
function _driveEscape(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

// Header colonna (es. "DIAGNOSI / MOTIVO RICOVERO")
// Replica .alt-col-header / .alt-col-header-split dell'app:
//   background #e8e6e1, bold, uppercase, center, font 8pt
function _driveColHeader(label, isSplit) {
  return '<p style="background-color:#e8e6e1;padding:4pt 6pt;font-weight:bold;' +
         'font-size:8pt;text-align:center;text-transform:uppercase;letter-spacing:0.3pt;' +
         'margin:0;color:#263238;' +
         (isSplit ? 'border-top:1pt solid #333;' : '') +
         '">' + _driveEscape(label) + '</p>';
}

// Wrap valore campo: padding standard
function _driveValueHtml(val, isHtml) {
  if (val == null || val === '') return '<p style="padding:5pt 8pt;font-size:9pt;margin:0;color:#999;font-style:italic;">&nbsp;</p>';
  var content = isHtml ? String(val) : _driveEscape(val);
  return '<p style="padding:5pt 8pt;font-size:9pt;margin:0;line-height:1.4;color:#263238;">' + content + '</p>';
}

// Rendering della cella INFO (sinistra). Contiene Letto/Tipologia/Nome
// con stili speciali (numero gigante, badge colorato, nome maiuscolo),
// poi Allergie con label rossa, e tutti gli altri campi info come
// piccola label uppercase + valore.
function _driveRenderColInfo(p, schemaInfo) {
  // Indice rapido per accesso ai campi standard nello schema info
  var bySpec = {};
  schemaInfo.forEach(function(s) { bySpec[s.campo] = s; });

  var tipo  = (p.TipologiaLetto || '').trim().toUpperCase() || 'STANDARD';
  var nome  = (p.Nome || '').trim();
  var letto = String(p.Letto || '?');
  var sesso = (p.Sesso || '').toUpperCase();
  var sessoSym = sesso === 'M' ? '♂' : sesso === 'F' ? '♀' : '';
  var sessoColor = sesso === 'M' ? '#1976d2' : sesso === 'F' ? '#c2185b' : '#666';
  var tipoColor = stringToColor(tipo);

  // Header speciale: letto + sesso + tipologia + nome
  // Le LABEL sono presenti per il parser, ma stylate piccole/discrete.
  // Il VALORE invece ha lo stile grande della card app.
  var lblStyle = 'font-size:7pt;color:#888;font-weight:bold;text-transform:uppercase;letter-spacing:0.4pt;margin:6pt 0 1pt;text-align:center;';

  var html = '';

  // LETTO grande (replica .alt-bed-number: 2rem bold)
  if (bySpec.Letto) {
    html += '<p style="' + lblStyle + '">' + _driveEscape(bySpec.Letto.label) + '</p>';
    html += '<p style="font-size:24pt;font-weight:bold;text-align:center;margin:0;line-height:1;color:#263238;">' + _driveEscape(letto) + '</p>';
  }
  // TIPOLOGIA (replica badge colorato)
  if (bySpec.TipologiaLetto) {
    html += '<p style="' + lblStyle + '">' + _driveEscape(bySpec.TipologiaLetto.label) + '</p>';
    html += '<p style="background-color:' + tipoColor + ';color:#fff;text-align:center;font-size:8.5pt;' +
            'font-weight:bold;text-transform:uppercase;padding:3pt 6pt;margin:0;letter-spacing:0.8pt;">' +
            _driveEscape(tipo) + '</p>';
  }
  // SESSO (piccola riga)
  if (bySpec.Sesso) {
    html += '<p style="' + lblStyle + '">' + _driveEscape(bySpec.Sesso.label) + '</p>';
    html += '<p style="text-align:center;font-size:11pt;margin:0;color:' + sessoColor + ';font-weight:bold;">' +
            (sessoSym || _driveEscape(sesso) || '—') + '</p>';
  }
  // NOME (replica .alt-nome: 1.1rem bold uppercase)
  if (bySpec.Nome) {
    html += '<p style="' + lblStyle + '">' + _driveEscape(bySpec.Nome.label) + '</p>';
    html += '<p style="font-size:11pt;font-weight:bold;text-align:center;text-transform:uppercase;' +
            'margin:0;padding:2pt;border-bottom:1pt solid #ccc;color:#263238;">' +
            (_driveEscape(nome) || '<span style="color:#aaa;font-style:italic;font-weight:normal;">— vuoto —</span>') + '</p>';
  }
  // ALLERGIE (replica .alt-allergie-label: rosso, uppercase, bold)
  if (bySpec.Allergie) {
    html += '<p style="margin:6pt 0 1pt;font-size:9pt;font-weight:bold;color:#c62828;text-transform:uppercase;text-align:center;">' +
            _driveEscape(bySpec.Allergie.label) + '</p>';
    var allergieVal = String(p.Allergie || '');
    html += '<p style="font-size:9.5pt;padding:2pt 4pt;margin:0;text-align:center;color:#263238;">' +
            (allergieVal || '<span style="color:#aaa;font-style:italic;">—</span>') + '</p>';
  }

  // Altri campi info (Data Nascita, Età, Ricovero, CS, Ossigeno, Vitto,
  // Dimissibile, eventuali campi NUOVI aggiunti) → label piccola + valore
  var processed = { Letto:1, TipologiaLetto:1, Nome:1, Sesso:1, Allergie:1 };
  schemaInfo.forEach(function(s) {
    if (processed[s.campo]) return;
    var val = p[s.campo];
    var labelStyle = 'font-size:8pt;font-weight:bold;color:#37474f;text-transform:none;margin:5pt 0 0;';
    var valStyle = 'font-size:9pt;margin:0;color:#263238;padding:0 0 0 2pt;';
    if (s.isHtml) {
      html += '<p style="' + labelStyle + '">' + _driveEscape(s.label) + '</p>';
      html += '<p style="' + valStyle + '">' + (val ? String(val) : '<span style="color:#aaa;font-style:italic;">—</span>') + '</p>';
    } else {
      html += '<p style="' + labelStyle + '">' + _driveEscape(s.label) + '</p>';
      html += '<p style="' + valStyle + '">' + _driveEscape(val) + '</p>';
    }
  });

  return html;
}

// Rendering della cella DIAG (centro-sinistra). Diagnosi sopra,
// poi PIANO DI CURA, ESAMI COLTURALI, DA FARE / RICHIESTE — divisi
// da split header. Ogni campo è una "sezione" con header + valore.
// MA siamo dentro una cella, non in una sottotabella, per parsing
// semplice.
function _driveRenderDiagSection(spec, val, isFirst) {
  return _driveColHeader(spec.label, !isFirst) + _driveValueHtml(val, spec.isHtml);
}

function _driveRenderCard(p) {
  var schema = _getDriveSchema();
  // Schema groups: info, diag, diaria, terapia

  // ── COLONNA 1: INFO (rowspan = numero righe = numero campi diag) ──
  var infoHtml = _driveRenderColInfo(p, schema.info);

  // ── COLONNA 2: DIAG (una riga per campo, headers stacked) ──
  // Costruita come singola cella con multiple sezioni stacked verticalmente.
  // Ogni sezione: header colonna + valore. Visivamente identico a
  // .alt-col-diag con .alt-diag-top/.alt-diag-bottom/.alt-diag-fourth/.alt-diag-third.
  var diagHtml = schema.diag.map(function(s, i) {
    return _driveRenderDiagSection(s, p[s.campo], i === 0);
  }).join('');
  if (!diagHtml) diagHtml = '<p style="padding:6pt;font-style:italic;color:#aaa;">—</p>';

  // ── COLONNA 3: DIARIA (più grande) ──
  var diariaHtml = schema.diaria.map(function(s, i) {
    return _driveRenderDiagSection(s, p[s.campo], i === 0);
  }).join('');
  if (!diariaHtml) diariaHtml = '<p style="padding:6pt;font-style:italic;color:#aaa;">—</p>';

  // ── COLONNA 4: TERAPIA (destra) ──
  var terapiaHtml = schema.terapia.map(function(s, i) {
    return _driveRenderDiagSection(s, p[s.campo], i === 0);
  }).join('');
  if (!terapiaHtml) terapiaHtml = '<p style="padding:6pt;font-style:italic;color:#aaa;">—</p>';

  // Layout tabella 4-col come la card app (col widths come .alt-col-*)
  return (
    '<table width="100%" style="border-collapse:collapse;margin-top:14pt;margin-bottom:14pt;' +
    'border:' + _D_B + ';font-family:Arial,sans-serif;page-break-inside:avoid;">' +
      '<tr>' +
        '<td width="12%" style="background-color:#f2efe9;vertical-align:top;padding:4pt;border:' + _D_Bi + ';">' +
          infoHtml +
        '</td>' +
        '<td width="18%" style="vertical-align:top;padding:0;border:' + _D_Bi + ';">' +
          diagHtml +
        '</td>' +
        '<td width="58%" style="vertical-align:top;padding:0;border:' + _D_Bi + ';">' +
          diariaHtml +
        '</td>' +
        '<td width="12%" style="vertical-align:top;padding:0;border:' + _D_Bi + ';">' +
          terapiaHtml +
        '</td>' +
      '</tr>' +
    '</table>'
  );
}

// Card NOTE: una tabella 4-col come le altre, ma quasi vuota:
// solo "LETTO" → NOTE in col-info e "DIARIA ED EPICRISI" in col-diaria
function _driveRenderNoteCard(p) {
  var diariaSpec = null;
  // Trova lo spec Diaria dallo schema
  _DRIVE_LAYOUT.forEach(function(s) {
    if (s.campo === 'Diaria') diariaSpec = s;
  });
  if (!diariaSpec) diariaSpec = { campo: 'Diaria', label: 'DIARIA ED EPICRISI', isHtml: true };

  // Cella info semplificata: solo "LETTO" + "NOTE"
  var lblStyle = 'font-size:7pt;color:#888;font-weight:bold;text-transform:uppercase;letter-spacing:0.4pt;margin:6pt 0 1pt;text-align:center;';
  var infoHtml =
    '<p style="' + lblStyle + '">LETTO</p>' +
    '<p style="font-size:13pt;font-weight:bold;letter-spacing:2pt;color:#546e7a;text-align:center;margin:0;padding:8pt 0;">NOTE</p>';

  // Cella diaria: header + valore
  var diariaHtml = _driveColHeader(diariaSpec.label, false) +
                   _driveValueHtml(p.Diaria || '', true);

  return (
    '<table width="100%" style="border-collapse:collapse;margin-top:14pt;margin-bottom:14pt;' +
    'border:' + _D_B + ';font-family:Arial,sans-serif;page-break-inside:avoid;">' +
      '<tr>' +
        '<td width="12%" style="background-color:#eceff1;vertical-align:top;padding:4pt;border:' + _D_Bi + ';">' +
          infoHtml +
        '</td>' +
        '<td colspan="3" width="88%" style="vertical-align:top;padding:0;border:' + _D_Bi + ';">' +
          diariaHtml +
        '</td>' +
      '</tr>' +
    '</table>'
  );
}

// Esegue backup su Google Drive come Google Doc con layout standard.
// Le cartelle vengono create la prima volta e l'ID salvato in Supabase.
// Fire-and-forget: non blocca il flusso principale.
function _driveBackupConsegne(pazienti, ts) {
  // Helper logging coerente (logga solo se _log è disponibile)
  function _bkLog(livello, esito, msg, dettagli) {
    console.log('[Drive backup]', esito, '—', msg);
    if (typeof window._log === 'function') {
      try { window._log(livello, 'drive-backup', esito + ' — ' + msg, dettagli || ''); }
      catch(e) {}
    }
  }
  if (!window._googleDriveToken) {
    // SKIP: questo PC ha vinto la CAS atomica ma non ha il token Drive.
    // Causa tipica: il login Google a Drive è scaduto (Google rinnova il
    // token ogni ~1h e in app aperta da molte ore senza interazione il
    // token può essere null). Conseguenza: il backup va SOLO su Supabase,
    // niente file nella cartella Drive per quest'ora.
    _bkLog('warning', 'skip-no-token',
      'Backup orario su DB OK ma Google Drive saltato',
      'Questo PC ha vinto il CAS atomico ma _googleDriveToken=null ' +
      '(login Drive scaduto o mai concesso). Per avere il file nella ' +
      'cartella Drive servirebbe un altro PC col token attivo.');
    return Promise.resolve();
  }
  var data = new Date(Number(ts));
  var pad = function(n) { return String(n).padStart(2, '0'); };
  var dataLabel = pad(data.getDate()) + '/' + pad(data.getMonth() + 1) + '/' + data.getFullYear() +
                  ' ' + pad(data.getHours()) + ':' + pad(data.getMinutes());
  var nomeSafe  = pad(data.getDate()) + '-' + pad(data.getMonth() + 1) + '-' + data.getFullYear() +
                  '_' + pad(data.getHours()) + '-' + pad(data.getMinutes());
  var nomeFile  = 'Backup_' + nomeSafe; // senza estensione → sarà un Google Doc

  // Ordina per numero letto (NOTE sempre ultima)
  var ordinati = (pazienti || []).slice().sort(function(a, b) {
    if (a.Letto === 'NOTE') return 1;
    if (b.Letto === 'NOTE') return -1;
    var nA = parseInt(a.Letto, 10), nB = parseInt(b.Letto, 10);
    return (!isNaN(nA) && !isNaN(nB)) ? nA - nB : String(a.Letto).localeCompare(String(b.Letto));
  });

  // Costruisce HTML: tabelle label-valore (autoritative per parsing/import).
  // Ogni paziente è una tabella 2-colonne (Label | Valore) con TUTTI i 18
  // campi. Il parser cerca per LABEL (no positional) quindi:
  //   - il medico può aggiornare manualmente i valori dal Doc
  //   - aggiungere righe vuote o riordinare non rompe il parser
  //   - round-trip backup → modifica → import è LOSSLESS
  var cardsHtml = ordinati.map(function(p) {
    return p.Letto === 'NOTE' ? _driveRenderNoteCard(p) : _driveRenderCard(p);
  }).join('');

  var html = '<!DOCTYPE html><html><head><meta charset="UTF-8">' +
    '<style>' +
    '@page{size:A4 landscape;margin:14mm 16mm 14mm 16mm;}' +
    'body{font-family:Arial,sans-serif;font-size:9pt;margin:0;padding:0;color:#263238;}' +
    'p{margin:0;padding:0;}' +
    '</style>' +
    '</head><body>' +
    // Header documento
    '<table width="100%" style="margin-bottom:6pt;border-bottom:3pt solid #37474f;">' +
    '<tr>' +
    '<td style="font-size:18pt;font-weight:bold;color:#37474f;padding-bottom:6pt;letter-spacing:0.5pt;">' +
      '🛏 CONSEGNE REPARTO' +
    '</td>' +
    '<td style="text-align:right;vertical-align:bottom;font-size:10pt;color:#546e7a;padding-bottom:6pt;">' +
      '<b>Backup automatico</b><br>' +
      '<span style="font-size:11pt;color:#37474f;font-weight:bold;">' + dataLabel + '</span>' +
    '</td>' +
    '</tr></table>' +
    // Banner istruzioni emergenza
    '<table width="100%" style="background-color:#fff8e1;border-left:3pt solid #f4a300;margin-bottom:14pt;">' +
    '<tr><td style="padding:6pt 10pt;font-size:8.5pt;color:#7a5800;line-height:1.4;">' +
      '<b>⚠ Safety net per emergenze</b> · ' +
      'In caso di down del sistema, modificare i valori nelle celle di destra delle tabelle ' +
      '(NON modificare le etichette di sinistra). Ri-importare poi nelle consegne via menu ⚙ → "Importa da backup Drive".' +
    '</td></tr></table>' +
    cardsHtml +
    '</body></html>';

  // Cerca o crea la cartella nella root, crea il file, poi pulisce i vecchi
  return _driveGetOrCreateFolder('BACKUP CONSEGNE EMERGENZA')
    .then(function(folderId) {
      return _driveCreaGoogleDoc(nomeFile, html, folderId)
        .then(function(file) {
          _bkLog('info', 'success',
            'File Drive creato: ' + (file && file.name),
            'fileId=' + (file && file.id) + ' | folderId=' + folderId);
          // Pulizia file vecchi con la stessa retention di Supabase archivio
          return _sbGetGiorniConservazione().then(function(giorni) {
            var cutoff = Number(ts) - (giorni * 86400000);
            _driveEliminaVecchi(folderId, cutoff); // fire-and-forget
          });
        });
    })
    .catch(function(e) {
      var msg = (e && (e.message || e.error || e.toString())) || 'errore sconosciuto';
      _bkLog('error', 'error',
        'Backup Drive FALLITO: ' + msg,
        'nomeFile=' + nomeFile + ' | stack=' + (e && e.stack ? String(e.stack).substring(0,300) : 'n/a'));
    });
}


// ══════════════════════════════════════════════════════════════
// REALTIME — aggiornamenti istantanei senza polling
// ══════════════════════════════════════════════════════════════

var _realtimeTimer = null;
var _realtimeChannel = null;

function _inizializzaRealtime() {
  if (_realtimeChannel) { _sb.removeChannel(_realtimeChannel); }

  _realtimeChannel = _sb.channel('consegne-live')
    .on('postgres_changes', { event: '*', schema: 'public', table: 'consegne' },
      function(payload) {
        // FAST-PATH delta-update: aggiorna SOLO la card interessata invece di
        // re-renderizzare l'intera lista. Su PC lento riduce il freeze da 2-5s
        // a <100ms per evento Realtime. Se il delta non gestisce il caso
        // (NOTE, INSERT, errori) → fallback automatico al full sync.
        // Compat: in alcune versioni Supabase v2 il campo è "eventType",
        // in altre "event". Accettiamo entrambi.
        var evType = payload && (payload.eventType || payload.event || '');
        var ok = false;
        try {
          if (evType === 'UPDATE' && payload.new) {
            ok = _applicaDeltaUpdate(payload.new);
          } else if (evType === 'DELETE' && payload.old) {
            ok = _applicaDeltaDelete(payload.old.letto);
          }
          // INSERT: lasciamo fare full sync (serve ordinamento corretto)
        } catch (e) {
          console.warn('[Realtime delta error, fallback full sync]', e);
          ok = false;
        }
        // Se il delta fallisce: passa il letto specifico così il fallback
        // fa SELECT mirato (~17KB) invece di SELECT * (~500KB).
        // Letto noto da payload.new (UPDATE/INSERT) o payload.old (DELETE).
        if (!ok) {
          var lettoFallback = (payload.new && payload.new.letto) ||
                              (payload.old && payload.old.letto) || null;
          // Per NOTE e INSERT serve full sync (rendering completo)
          if (lettoFallback === 'NOTE' || evType === 'INSERT') lettoFallback = null;
          _scheduleRealtimeSync(lettoFallback);
        }
      })
    // Lock: gestiti direttamente in-memory, nessuna query aggiuntiva al DB
    .on('postgres_changes', { event: 'INSERT', schema: 'public', table: 'locks' },
      function(payload) { _onLockChange('INSERT', payload); })
    .on('postgres_changes', { event: 'UPDATE', schema: 'public', table: 'locks' },
      function(payload) { _onLockChange('UPDATE', payload); })
    .on('postgres_changes', { event: 'DELETE', schema: 'public', table: 'locks' },
      function(payload) { _onLockChange('DELETE', payload); })
    // App version: nuovo deploy notificato dalla Action page_build →
    // PATCH a Supabase → Realtime → tutti i client mostrano il badge.
    .on('postgres_changes',
      { event: 'UPDATE', schema: 'public', table: 'app_version', filter: 'id=eq.1' },
      function(payload) {
        var newSha = (payload && payload.new && payload.new.sha) || '';
        if (typeof _onAppVersionChange === 'function') _onAppVersionChange(newSha);
      })
    .subscribe(function(status) {
      console.log('[Realtime]', status);
    });
}

// ── DELTA UPDATE — aggiorna SOLO la card interessata ──────────────────
// Ritorna true se gestito, false se serve fallback al full sync.
function _applicaDeltaUpdate(row) {
  if (!row || !row.letto) return false;
  var p = _fromDb(row);
  var letto = String(p.Letto);

  // Skip card lockata da questo client (focus mode su questo letto):
  // _applicaAggiornamentoDaHtml ha la stessa guard, replicata qui per coerenza.
  if (typeof _lettoLockAtt !== 'undefined' && _lettoLockAtt === letto) return true;

  // NOTE ha renderer dedicato (_renderNoteCard*): meglio fallback full sync
  if (letto === 'NOTE') return false;

  var cards = document.querySelectorAll('.patient-card[data-bed="' + letto + '"]');
  if (!cards.length) return false; // card non esistente in DOM → full sync

  cards.forEach(function(card) { _aggiornaCardDaPaziente(card, p); });

  // Aggiorna sync indicator (come fa _scheduleRealtimeSync alla fine)
  var ind = document.getElementById('syncIndicator');
  var st  = document.getElementById('syncStatus');
  if (ind) ind.className = 'badge bg-success ms-2';
  if (st)  st.innerText  = new Date().toLocaleTimeString('it-IT');

  // Aggiorna eventuali badge/riepiloghi se la funzione esiste
  if (typeof window._aggiornaBadgePrincipali === 'function') {
    try { window._aggiornaBadgePrincipali(); } catch(e) {}
  }
  return true;
}

// Rimuove la card dal DOM (delta DELETE).
function _applicaDeltaDelete(letto) {
  if (!letto) return false;
  var rimossi = 0;
  document.querySelectorAll('.patient-card[data-bed="' + String(letto) + '"]').forEach(function(c) {
    c.remove();
    rimossi++;
  });
  if (rimossi === 0) return true; // già non c'era → no-op, comunque success

  // Aggiorna messaggio "Nessun letto" se la lista è vuota
  var nessun = document.querySelectorAll('.patient-card[data-bed]').length === 0;
  var noMsg  = document.getElementById('noLettiMsg'); if (noMsg) noMsg.style.display = nessun ? '' : 'none';

  // Aggiorna riepiloghi se disponibili
  if (typeof window._aggiornaBadgePrincipali === 'function') {
    try { window._aggiornaBadgePrincipali(); } catch(e) {}
  }
  return true;
}

// Applica i dati di un paziente alla card (entrambe le viste main e alt).
// Replica la logica di _applicaAggiornamentoDaHtml ma per UN paziente,
// lavorando sull'oggetto JS senza DOMParser. Preserva caret position.
function _aggiornaCardDaPaziente(card, p) {
  var activeEl = document.activeElement;

  // Campi testuali (11 campi — include EsamiColturali aggiunto 2026-05-11)
  var campi = ['Nome','Diagnosi','Eta','NoteTerapia','Diaria','DaFare',
               'PianoTerapeutico','EsamiColturali','Allergie','CodiceSanitario','Ossigeno','Vitto'];
  campi.forEach(function(campo) {
    var dst = card.querySelector('[data-field="' + campo + '"]');
    if (!dst) return;
    var val = p[campo] != null ? String(p[campo]) : '';
    var hasFocus = (dst === activeEl || dst.contains(activeEl));
    var caretPos = (hasFocus && typeof _salvaCaretPos === 'function') ? _salvaCaretPos(dst) : null;
    if (dst.classList.contains('plain-text')) {
      if (dst.innerText !== val) dst.innerText = val;
    } else {
      if (dst.innerHTML !== val) dst.innerHTML = val;
    }
    if (hasFocus && caretPos !== null && typeof _ripristinaCaretPos === 'function') {
      dst.focus();
      _ripristinaCaretPos(dst, caretPos);
    }
    // Aggiorna stato placeholder (mostra/nasconde la classe .is-empty)
    if (typeof _placeholderCheck === 'function') _placeholderCheck(dst);
  });

  // Isolamento: la Diaria appena aggiornata può aver aggiunto/rimosso il
  // banner → sincronizza l'icona virus accanto al numero letto.
  if (typeof window._isolamentoSyncIcona === 'function') {
    try { window._isolamentoSyncIcona(card); } catch(e) {}
  }

  // Date (data ricovero + giorni calcolati, data nascita + età ricalcolata in UI)
  var ric = _parseDataRicovero(p.DataRicovero);
  var dr = card.querySelector('.data-ricovero-text');
  if (dr && dr.value !== ric.vis) dr.value = ric.vis;
  var gg = card.querySelector('.valore-giorni');
  if (gg && String(gg.innerText) !== String(ric.giorni)) gg.innerText = ric.giorni;
  var nasc = _parseDataNascita(p.DataNascita);
  var dn = card.querySelector('.data-nascita-text');
  if (dn && dn.value !== nasc.vis) dn.value = nasc.vis;

  // Sesso (simbolo + colore via _aggiornaSimboloSesso definita in index.html)
  var sesso = card.querySelector('.sesso-symbol[data-field="Sesso"]');
  if (sesso) {
    var nuovoSesso = p.Sesso || '';
    if (sesso.getAttribute('data-sesso') !== nuovoSesso) {
      sesso.setAttribute('data-sesso', nuovoSesso);
      if (typeof _aggiornaSimboloSesso === 'function') _aggiornaSimboloSesso(sesso);
    }
  }

  // Tipologia letto + badge
  var nuovaTipo = (p.TipologiaLetto || '').toUpperCase();
  if (card.getAttribute('data-tipologia') !== nuovaTipo) {
    card.setAttribute('data-tipologia', nuovaTipo);
  }
  var badge = card.querySelector('[id^="badge-tipo-"]');
  if (badge) {
    var testoBadge = nuovaTipo || 'STANDARD';
    if ((badge.innerText || '').trim() !== testoBadge) badge.innerText = testoBadge;
    if (nuovaTipo && nuovaTipo !== 'STANDARD') {
      var col = (typeof window._getColoreTipo === 'function')
        ? window._getColoreTipo(nuovaTipo)
        : stringToColor(nuovaTipo);
      if (badge.style.backgroundColor !== col) badge.style.backgroundColor = col;
    } else {
      if (badge.style.backgroundColor) badge.style.backgroundColor = '';
    }
  }

  // Dimissibile: checkbox + scritta + icona accanto al badge tipologia
  var dimNew = (p.Dimissibile || '').trim();
  var dimRow = card.querySelector('.dim-row');
  if (dimRow) {
    var attualeDim = dimRow.getAttribute('data-value') || '';
    if (attualeDim !== dimNew) {
      dimRow.setAttribute('data-value', dimNew);
      var cb = dimRow.querySelector('.dim-checkbox');
      if (cb) cb.checked = !!dimNew;
      var info = dimRow.querySelector('.dim-data-info, .dim-data-info-empty');
      if (info) {
        if (dimNew) {
          info.className = 'dim-data-info';
          info.innerHTML = _dimIconSvg(12) + ' Dimissione il ' +
            String(dimNew).replace(/[&<>"']/g, function(c) {
              return { '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;' }[c];
            });
        } else {
          info.className = 'dim-data-info-empty text-muted';
          info.innerHTML = '';
        }
      }
    }
  }
  // Icona "in dimissione" accanto al badge tipologia
  var tipoWrap = card.querySelector('.tipo-badge-wrap');
  if (tipoWrap) {
    var iconaEsistente = tipoWrap.querySelector('.dim-icon-badge');
    if (dimNew && !iconaEsistente) {
      var span = document.createElement('span');
      span.className = 'dim-icon-badge ms-1';
      span.title = 'Paziente in dimissione — ' + dimNew;
      span.innerHTML = _dimIconSvg(14);
      tipoWrap.appendChild(span);
    } else if (!dimNew && iconaEsistente) {
      iconaEsistente.remove();
    } else if (dimNew && iconaEsistente) {
      iconaEsistente.title = 'Paziente in dimissione — ' + dimNew;
    }
  }
}

// Aggiorna _lockState e UI senza toccare il DB
function _onLockChange(event, payload) {
  if (event === 'DELETE') {
    var letto = payload.old && payload.old.letto;
    if (letto) delete _lockState[letto];
  } else {
    var row = payload.new;
    if (row && row.letto) {
      _lockState[row.letto] = { token: row.token };
    }
  }
  if (typeof _applicaLocks === 'function') { _applicaLocks(_lockState); }
  var ind = document.getElementById('syncIndicator');
  var st  = document.getElementById('syncStatus');
  if (ind) ind.className = 'badge bg-success ms-2';
  if (st)  st.innerText  = new Date().toLocaleTimeString('it-IT');
}

var _pendingSync = false;
var _pendingSyncLetto = null;  // se non-null, sync mirato su un solo letto

// Sync intelligente: se lettoSpecifico è passato, fa SELECT mirato (~17KB)
// invece di SELECT * (~500KB con 30 letti). Riduce egress drasticamente
// quando il fallback è scattato solo per un singolo letto cambiato.
// Senza lettoSpecifico, mantiene il comportamento legacy (full sync).
function _scheduleRealtimeSync(lettoSpecifico) {
  // Se un sync mirato è in coda ma arriva una richiesta full → promuove a full
  if (_pendingSync) {
    if (!lettoSpecifico) _pendingSyncLetto = null;
    return;
  }
  _pendingSync = true;
  _pendingSyncLetto = lettoSpecifico || null;
  clearTimeout(_realtimeTimer);
  _realtimeTimer = setTimeout(function() {
    var lettoTarget = _pendingSyncLetto;
    _pendingSync = false;
    _pendingSyncLetto = null;

    var ind = document.getElementById('syncIndicator');
    var st  = document.getElementById('syncStatus');
    if (ind) ind.className = 'badge bg-warning text-dark ms-2';
    if (st)  st.innerText  = 'Sync...';

    var fetchPromise;
    if (lettoTarget) {
      // SELECT mirato: 1 sola riga (~17KB invece di ~500KB)
      fetchPromise = _q(_sb.from('consegne').select('*').eq('letto', lettoTarget))
        .then(function(rows) {
          if (!rows || !rows.length) return null;
          var p = _fromDb(rows[0]);
          // Applica direttamente alla card senza full re-render
          var cards = document.querySelectorAll('.patient-card[data-bed="' + lettoTarget + '"]');
          if (cards.length && typeof _aggiornaCardDaPaziente === 'function') {
            cards.forEach(function(card) { _aggiornaCardDaPaziente(card, p); });
            if (typeof window._aggiornaBadgePrincipali === 'function') {
              try { window._aggiornaBadgePrincipali(); } catch(e) {}
            }
            return 'mirato';
          }
          // Card non trovata → promuovi a full sync (rara)
          return _sbGetPazienti().then(function(pazienti) {
            if (typeof _applicaAggiornamentoCompleto === 'function') {
              _applicaAggiornamentoCompleto(_renderCardsHtml(pazienti));
            }
            return 'full-fallback';
          });
        });
    } else {
      // Full sync legacy: SELECT * (costoso, evitare quando possibile)
      fetchPromise = _sbGetPazienti().then(function(pazienti) {
        if (typeof _applicaAggiornamentoCompleto === 'function') {
          _applicaAggiornamentoCompleto(_renderCardsHtml(pazienti));
        }
        return 'full';
      });
    }

    fetchPromise.then(function() {
      if (ind) ind.className = 'badge bg-success ms-2';
      if (st)  st.innerText  = new Date().toLocaleTimeString('it-IT');
    }).catch(function(e) {
      console.warn('[Realtime sync error]', e);
      if (ind) ind.className = 'badge bg-danger ms-2';
      if (st)  st.innerText  = 'Errore sync';
    });
  }, 300); // debounce 300ms
}


// ══════════════════════════════════════════════════════════════
// MODALITÀ EMERGENZA PC LENTO — funzioni standalone, on-demand
// Non interferiscono con il resto del codice. Attivate solo se
// l'utente clicca esplicitamente il pulsante in navbar.
// ══════════════════════════════════════════════════════════════
var _emergenzaPollingId   = null;  // Polling letture (5s)
var _emergenzaForceSaveId = null;  // Force-save scan letti dirty (10s)

// Avvia polling ogni 5s: aggiorna card pazienti + stato lock
// Avvia force-save scan ogni 10s: rilancia salvataggi pending/falliti
function _emergenzaAvviaPolling() {
  if (_emergenzaPollingId) return;
  console.log('[Emergenza] Polling attivato (check 30s + force-save 30s)');
  if (typeof window._log === 'function') {
    window._log('warning', 'emergenza-on',
      'Modalità emergenza attivata (check 30s + force-save 30s)',
      'Il sistema è passato dal Realtime standard al polling periodico. ' +
      'Versione LIGHT: ogni 30s controlla solo MAX(updated_at), full sync ' +
      'solo se rilevate modifiche. Causa: diagnostica ha rilevato problemi ' +
      'di connessione/WebSocket o attivazione manuale.');
  }

  // POLLING LETTURE — versione LIGHT (riduce egress drasticamente)
  // PROBLEMA precedente: ogni 5s SELECT * FROM consegne (~150KB/query
  // con 30 letti rich-text). In 24h = ~2.4 GB di egress solo da qui.
  // Su 3 giorni di emergenza = >5 GB → quota free Supabase esaurita.
  //
  // SOLUZIONE: check leggero ogni 30s con SELECT updated_at LIMIT 1
  // (~50 byte). Full sync (_sbGetPazienti) eseguito SOLO se MAX(updated_at)
  // è cambiato dalla volta precedente. Risparmio: ~99% nelle finestre
  // di inattività, quasi-zero costo durante editing normale.
  // Cadenza scelta dopo analisi empirica: 30s permette di stare sotto
  // 2GB free tier anche nel worst case (5 PC × 20h emergenza × 30g).
  _emergenzaPollingId = setInterval(function() {
    _emergenzaCheckSeMutato().catch(function(){});
    _sbGetLocks().then(function(locks) {
      if (typeof _applicaLocks === 'function') _applicaLocks(locks || {});
    }).catch(function(){});
  }, 30000);

  // FORCE-SAVE SCAN — safety net per la modalità emergenza:
  // ogni 30s, se ci sono letti in _dirtyLetti che NON hanno un retry esponenziale
  // attivo (gestito nativamente da app.js) e NON sono in focus mode, forza il save.
  // Utile quando: (a) i 3 retry esponenziali sono esauriti, (b) un save è stato
  // bloccato dal _syncPaused di un altro ciclo, (c) un dirty rimane stagnante.
  _emergenzaForceSaveId = setInterval(function() {
    // Niente da fare se non ci sono letti dirty
    if (typeof _dirtyLetti === 'undefined' || !_dirtyLetti.size) return;
    // Salta se un save è già in corso a livello globale (riproveremo al prossimo ciclo)
    if (typeof _syncPaused !== 'undefined' && _syncPaused) {
      console.log('[Emergenza] Force-save: save in corso, salto ciclo');
      return;
    }
    console.log('[Emergenza] Force-save scan: '+_dirtyLetti.size+' letto/i dirty');
    Array.from(_dirtyLetti).forEach(function(letto) {
      // Salta se l'utente è in focus mode su questo letto (lo gestisce il flusso normale)
      if (typeof _focusCard !== 'undefined' && _focusCard &&
          _focusCard.getAttribute('data-bed') === letto) return;
      // Salta se un retry esponenziale è già pianificato per questo letto (evita race
      // tra force-save e retry nativo). Il retry nativo lo gestirà al delay previsto.
      if (typeof _saveRetryTimer !== 'undefined' && _saveRetryTimer[letto]) return;
      var card = document.querySelector('.patient-card[data-bed="' + letto + '"]');
      if (!card) return;
      // Annulla il debounce in corso (se presente) e forza il save adesso
      if (typeof timerSalvataggioLetto !== 'undefined' && timerSalvataggioLetto[letto]) {
        clearTimeout(timerSalvataggioLetto[letto]);
        delete timerSalvataggioLetto[letto];
      }
      if (typeof eseguiSalvataggioLettoCompleto === 'function') {
        eseguiSalvataggioLettoCompleto(letto, card);
      }
    });
  }, 30000);
}

// ── CHECK LEGGERO EMERGENZA — versione robusta ─────────────────────────
// Firma dello stato tabella = "MAX(updated_at)|COUNT(righe)".
// Perché entrambi:
//   - MAX(updated_at): rileva ogni UPDATE (dal trigger DB
//     trg_consegne_touch_updated_at è garantito server-side, immune agli
//     orologi sbagliati dei PC client e ai percorsi che non bumpavano)
//   - COUNT: rileva INSERT e soprattutto DELETE (una DELETE non altera il
//     MAX se la riga eliminata non era la più recente → prima era
//     invisibile ai PC in polling: "Elimina letto" non si sincronizzava)
// nullsFirst:false → un'eventuale riga con updated_at NULL non maschera
// il vero MAX (PostgREST di default mette i NULL per primi in DESC, il
// che avrebbe reso il check cieco per sempre).
// Payload: ~60 byte per tick, come prima.
var _emergFirmaSync   = null;   // firma dell'ultimo sync RIUSCITO
var _emergFirmaInit   = false;  // baseline registrata? (flag esplicito, non null-sentinel)
var _emergSyncInCorso = false;  // full sync in volo: non accodarne altri

function _emergenzaCheckSeMutato() {
  if (_emergSyncInCorso) return Promise.resolve(false);
  return _sb.from('consegne')
    .select('updated_at', { count: 'exact' })
    .order('updated_at', { ascending: false, nullsFirst: false })
    .limit(1)
    .then(function(res) {
      if (res.error) throw res.error;
      var maxTs = (res.data && res.data[0]) ? res.data[0].updated_at : null;
      var firma = String(maxTs) + '|' + String(res.count);

      // Prima esecuzione: registra baseline, no sync
      if (!_emergFirmaInit) {
        _emergFirmaInit = true;
        _emergFirmaSync = firma;
        return false;
      }
      if (firma === _emergFirmaSync) return false;

      console.log('[Emergenza] Modifica rilevata (' + firma + '), lancio full sync');
      _emergSyncInCorso = true;
      var ind = document.getElementById('syncIndicator');
      var st  = document.getElementById('syncStatus');
      if (ind) ind.className = 'badge bg-warning text-dark ms-2';
      if (st)  st.innerText  = 'Sync...';

      return _sbGetPazienti().then(function(pazienti) {
        if (typeof _applicaAggiornamentoCompleto === 'function') {
          _applicaAggiornamentoCompleto(_renderCardsHtml(pazienti));
        }
        // ⚠ Baseline avanzata SOLO a sync RIUSCITO. Prima veniva avanzata
        // subito dopo la detection: se il full sync falliva (rete instabile
        // — esattamente il contesto in cui la modalità emergenza è attiva),
        // la modifica era persa per sempre finché qualcos'altro non
        // cambiava. Ora al tick successivo (30s) si riprova.
        _emergFirmaSync = firma;
        _emergSyncInCorso = false;
        if (ind) ind.className = 'badge bg-success ms-2';
        if (st)  st.innerText  = new Date().toLocaleTimeString('it-IT');
        return true;
      }).catch(function(e) {
        _emergSyncInCorso = false;
        if (ind) ind.className = 'badge bg-danger ms-2';
        if (st)  st.innerText  = 'Errore sync';
        console.warn('[Emergenza] Full sync fallito, riprovo al prossimo tick:', e && e.message);
        return false;
      });
    });
}

function _emergenzaFermaPolling() {
  if (_emergenzaPollingId) {
    clearInterval(_emergenzaPollingId);
    _emergenzaPollingId = null;
  }
  if (_emergenzaForceSaveId) {
    clearInterval(_emergenzaForceSaveId);
    _emergenzaForceSaveId = null;
  }
  console.log('[Emergenza] Polling + force-save disattivati');
  if (typeof window._log === 'function') {
    window._log('info', 'emergenza-off',
      'Modalità emergenza disattivata (ritorno a Realtime standard)',
      'Disattivazione manuale o automatica (diagnostica non rileva più problemi).');
  }
}

function _emergenzaPollingAttivo() {
  return !!_emergenzaPollingId;
}


// ══════════════════════════════════════════════════════════════
// VERSION CHECK — badge "Aggiornamento disponibile"
// [E2E test 2026-05-10 — verifica deploy → Realtime → badge]
// [E2E test 2026-05-11 — verifica login flow ripulito]
// [E2E test 2026-05-11 — verifica comparsa badge arancione]
// [E2E test 2026-05-11 — verifica replica identity FULL]
//
// Flow:
//   push master → GitHub Pages deploy → Action page_build →
//   PATCH app_version → Realtime UPDATE → _onAppVersionChange() →
//   badge visibile → click → cleanup cache+SW → hard reload.
//
// "_localSha" rappresenta lo SHA del deploy che il client ha
// caricato. Salvato in localStorage così sopravvive ai reload:
//   - Al click sul badge lo CANCELLIAMO prima del reload, così
//     il prossimo load registra lo SHA nuovo.
//   - Senza cancellazione, dopo reload _localSha sarebbe ancora
//     il vecchio e il badge ricomparirebbe subito.
//
// Edge case: tab aperta DURANTE il deploy (CDN lag): il client
// fetcha bundle vecchio + ottiene SHA nuovo da Supabase. Mostra
// SHA nuovo in menu ma il codice è vecchio. Senza un build step
// che inietta lo SHA nel bundle è inevitabile. Frequenza bassa
// (~30s di finestra per deploy).
// ══════════════════════════════════════════════════════════════
var _localSha = (typeof localStorage !== 'undefined') ? localStorage.getItem('_appDeploySha') : null;
var _badgeAggiornamentoMostrato = false;
// Flag interno: true dopo il primo check al boot. Le chiamate successive
// (es. da _sincronizzaEPoiFai dopo un'operazione) NON fanno auto-apply.
var _versionCheckInitDone = false;

function _versionCheckInit() {
  // Memorizziamo SUBITO se questa è la prima chiamata (= boot della pagina).
  // Lo facciamo prima del fetch async per evitare race condition: se per
  // qualche motivo _versionCheckInit venisse richiamato in rapida
  // successione, solo la prima esecuzione potrà auto-applicare.
  var primaVolta = !_versionCheckInitDone;
  _versionCheckInitDone = true;
  // Carica SHA corrente di Supabase all'avvio
  _q(_sb.from('app_version').select('sha').eq('id', 1).maybeSingle())
    .then(function(row) {
      var remoteSha = (row && row.sha) || '';
      if (!_localSha && remoteSha) {
        // Prima volta assoluta (no SHA in localStorage): baseline silenziosa
        _localSha = remoteSha;
        try { localStorage.setItem('_appDeploySha', _localSha); } catch(e){}
      } else if (_localSha && remoteSha && _localSha !== remoteSha) {
        // C'è stato un deploy mentre la tab era chiusa / offline.
        if (primaVolta) {
          // AUTO-APPLY al boot: nessuna interazione utente richiesta.
          // L'utente vede l'overlay 'Aggiornamento in corso...' e la
          // pagina si ricarica con la versione nuova.
          console.log('[VersionCheck] Boot: deploy nuovo rilevato, auto-apply');
          if (typeof _applicaAggiornamentoVersione === 'function') {
            _applicaAggiornamentoVersione();
            return; // pagina sta per ricaricarsi, no menu update
          }
        } else {
          // Successive chiamate (durante sessione): badge come prima.
          _mostraBadgeAggiornamento();
        }
      }
      _aggiornaVoceMenuVersione();
    })
    .catch(function(e) {
      console.warn('[VersionCheck] init error:', e);
      _aggiornaVoceMenuVersione(); // mostra "(sconosciuta)" se fallisce
    });
}

// Chiamato dal listener Realtime postgres_changes su app_version
function _onAppVersionChange(newSha) {
  if (!newSha) return;
  if (!_localSha) {
    // primissimo evento ricevuto senza init completato: usalo come baseline
    _localSha = newSha;
    try { localStorage.setItem('_appDeploySha', _localSha); } catch(e){}
    _aggiornaVoceMenuVersione();
    return;
  }
  if (newSha !== _localSha) {
    _mostraBadgeAggiornamento();
  }
}

// POLLING FALLBACK per version check: ogni 5 minuti controlla app_version
// sul DB e mostra il badge se lo SHA è diverso da _localSha. Serve come
// safety net nel caso il Realtime listener perda eventi (WebSocket
// disconnesso, race condition al boot, ecc.). Costo: ~50 byte per query
// × 12 query/ora × 24h × 30g × 5 PC = ~450 KB/mese, trascurabile.
function _versionCheckPoll() {
  _q(_sb.from('app_version').select('sha').eq('id', 1).maybeSingle())
    .then(function(row) {
      var remoteSha = (row && row.sha) || '';
      if (!remoteSha) return;
      if (!_localSha) {
        // Baseline mancante: setta senza mostrare badge
        _localSha = remoteSha;
        try { localStorage.setItem('_appDeploySha', _localSha); } catch(e){}
        _aggiornaVoceMenuVersione();
        return;
      }
      if (remoteSha !== _localSha) {
        // SHA diverso: forza la ricomparsa del badge anche se
        // _badgeAggiornamentoMostrato era già true da un evento precedente
        // (es. utente non ha applicato e poi è arrivato un nuovo deploy).
        if (!_badgeAggiornamentoMostrato) {
          console.log('[VersionCheck poll] Nuovo deploy rilevato via polling:', remoteSha.substring(0,8));
          _mostraBadgeAggiornamento();
        }
      }
    }).catch(function(){});
}

// Avvia il polling. Chiamato da app.js dopo il primo _versionCheckInit.
function _versionCheckPollStart() {
  if (typeof window._versionCheckPollId !== 'undefined' && window._versionCheckPollId) return;
  window._versionCheckPollId = setInterval(_versionCheckPoll, 5 * 60 * 1000);
}
window._versionCheckPollStart = _versionCheckPollStart;
window._versionCheckPoll      = _versionCheckPoll;

function _aggiornaVoceMenuVersione() {
  var el = document.getElementById('navVersionShaText');
  if (!el) return;
  if (_localSha) {
    el.textContent = _localSha.substring(0, 7);
    el.title = 'Commit SHA: ' + _localSha;
  } else {
    el.textContent = '(sconosciuta)';
    el.title = 'Nessun deploy registrato';
  }
}

function _mostraBadgeAggiornamento() {
  if (_badgeAggiornamentoMostrato) return;
  _badgeAggiornamentoMostrato = true;
  var b = document.getElementById('btnAggiornamentoDisponibile');
  if (b) b.style.display = 'inline-flex';
  // Mostra anche il toast in basso a destra (idempotente, gestisce
  // internamente il caso focus mode → posticipa fino a chiusura).
  if (typeof window._mostraToastUpdate === 'function') {
    try { window._mostraToastUpdate(); } catch(e) {}
  }
}

// Click handler del badge — cleanup totale + hard reload
function _applicaAggiornamentoVersione() {
  // Disabilita doppio click
  var b = document.getElementById('btnAggiornamentoDisponibile');
  if (b) b.disabled = true;

  // Mostra subito l'overlay full-screen "Aggiornamento in corso..."
  // così l'utente non vede il caos visivo del reload (login Google,
  // librerie non caricate, ecc.).
  var ov = document.getElementById('versionUpdateOverlay');
  if (ov) ov.style.display = 'flex';

  // Sequenza cleanup → hard reload
  Promise.resolve()
    .then(function() {
      if (window.caches) {
        return caches.keys().then(function(keys) {
          return Promise.all(keys.map(function(k) { return caches.delete(k); }));
        });
      }
    })
    .catch(function(){})
    .then(function() {
      if ('serviceWorker' in navigator) {
        return navigator.serviceWorker.getRegistrations().then(function(regs) {
          return Promise.all(regs.map(function(r) { return r.unregister(); }));
        });
      }
    })
    .catch(function(){})
    .then(function() {
      // Rimuovi lo SHA salvato: al prossimo load _versionCheckInit
      // registra lo SHA nuovo come baseline.
      try { localStorage.removeItem('_appDeploySha'); } catch(e){}
      // Hard reload con cache-bust via query string
      var url = window.location.pathname + '?_r=' + Date.now() + window.location.hash;
      window.location.replace(url);
    });
}
window._applicaAggiornamentoVersione = _applicaAggiornamentoVersione;


// ══════════════════════════════════════════════════════════════
// PLACEHOLDER per <div contenteditable> vuoti
//
// CSS :empty non basta: il browser auto-inserisce <br>, &nbsp; o
// whitespace dopo focus/cancellazione → :empty smette di matchare
// anche se visivamente il campo è vuoto.
//
// _placeholderCheck applica/rimuove la classe .is-empty in base a
// una logica più tollerante (pulisce <br>, &nbsp;, whitespace prima
// di valutare "vuoto"). La CSS regola .is-empty[data-placeholder]::before
// mostra il placeholder grigio.
// ══════════════════════════════════════════════════════════════
function _placeholderCheck(el) {
  if (!el || !el.classList || !el.hasAttribute('data-placeholder')) return;
  var html = (el.innerHTML || '')
    .replace(/<br\s*\/?\s*>/gi, '')
    .replace(/&nbsp;/gi, '')
    .replace(/<div>\s*<\/div>/gi, '')
    .replace(/\s+/g, '');
  var vuoto = (html === '');
  // Doppia validazione: anche textContent.trim() deve essere vuoto
  if (vuoto) {
    var text = (el.textContent || '').replace(/ /g, '').trim();
    vuoto = (text === '');
  }
  el.classList.toggle('is-empty', vuoto);
}
window._placeholderCheck = _placeholderCheck;

function _aggiornaTuttiPlaceholders(root) {
  var ctx = root || document;
  ctx.querySelectorAll('.editable-area[data-placeholder]').forEach(_placeholderCheck);
}
window._aggiornaTuttiPlaceholders = _aggiornaTuttiPlaceholders;

// Listener globale: ogni input/focus/blur su un contenteditable aggiorna
// la classe .is-empty per il singolo elemento (no scan completo).
document.addEventListener('input', function(e) {
  var t = e.target;
  if (t && t.classList && t.classList.contains('editable-area') &&
      t.hasAttribute && t.hasAttribute('data-placeholder')) {
    _placeholderCheck(t);
  }
}, true);
document.addEventListener('blur', function(e) {
  var t = e.target;
  if (t && t.classList && t.classList.contains('editable-area') &&
      t.hasAttribute && t.hasAttribute('data-placeholder')) {
    _placeholderCheck(t);
  }
}, true);


