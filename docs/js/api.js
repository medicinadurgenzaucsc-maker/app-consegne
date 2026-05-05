// ============================================================
// api.js — Backend Supabase v2
// Sostituisce GAS + Google Sheets con Supabase (Postgres + Realtime)
// ============================================================


// ── CONFIGURAZIONE ────────────────────────────────────────────
var SUPABASE_URL     = 'https://ifmmcvxzhwdkmzhsxcvb.supabase.co';
var SUPABASE_ANON_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImlmbW1jdnh6aHdka216aHN4Y3ZiIiwicm9sZSI6ImFub24iLCJpYXQiOjE3Nzc2Nzk1ODAsImV4cCI6MjA5MzI1NTU4MH0.LH8h4Fivtl3-TuiA050oF8iS4b80xrd2Dn6z8JjCoeA';
var APP_URL          = 'https://medicinadurgenzaucsc-maker.github.io/app-consegne/';
var PRINT_URL        = APP_URL + 'print.html';
var LOCK_TTL_MS      = 300000; // 5 minuti — safety net per crash browser (nessun rinnovo attivo)

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
  'Allergie':            'allergie',
  'CodiceSanitario':     'codice_sanitario',
  'Ossigeno':            'ossigeno',
  'Sesso':               'sesso',
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


// ── RENDERER CARD PRINCIPALE ──────────────────────────────────
function _renderMainCard(p) {
  var letto = String(p.Letto || '');
  var tipo  = String(p.TipologiaLetto || '').trim().toUpperCase();
  var ric   = _parseDataRicovero(p.DataRicovero);
  var nasc  = _parseDataNascita(p.DataNascita);
  var eta   = nasc.hasData ? nasc.eta : (p.Eta || '');
  var colore = tipo ? stringToColor(tipo) : 'transparent';

  var badgeHtml = tipo
    ? '<span class="badge text-white text-uppercase shadow-sm"' +
      ' style="background-color:' + colore + ';cursor:pointer;"' +
      ' id="badge-tipo-' + _aEsc(letto) + '"' +
      ' ondblclick="_apriModalTipologia(\'' + _aEsc(letto) + '\')"' +
      ' title="Doppio click per cambiare tipologia">' + _aEsc(tipo) + '</span>'
    : '<span class="badge bg-light text-secondary border text-uppercase shadow-sm"' +
      ' style="cursor:pointer;"' +
      ' ondblclick="_apriModalTipologia(\'' + _aEsc(letto) + '\')"' +
      ' title="Doppio click per cambiare tipologia">STANDARD</span>';

  return '<div class="patient-card shadow-sm" data-bed="' + _aEsc(letto) + '" data-tipologia="' + _aEsc(tipo) + '">' +
    '<div class="row m-0 header-row">' +
    '<div class="col-2 bed-number-box p-3">' +
    '<button class="focus-pencil-btn" title="Modifica scheda"' +
    ' onclick="_attivaFocusMode(this.closest(\'.patient-card\'))"><i class="bi bi-pencil-fill"></i></button>' +
    '<span id="status-' + _aEsc(letto) + '" class="badge status-badge position-absolute top-0 start-0 m-1 bg-secondary"></span>' +
    '<div class="text-muted fw-bold mt-2">LETTO</div>' +
    '<div class="bed-number-flex">' +
    '<div class="bed-number">' + _aEsc(letto) + '</div>' +
    '<span class="sesso-symbol" data-field="Sesso" data-sesso="' + _aEsc(p.Sesso || '') + '"></span>' +
    '</div>' +
    '<div class="mt-2 text-center tipo-badge-wrap" style="font-size:0.8rem;">' + badgeHtml + '</div>' +
    '</div>' +
    '<div class="col-10 patient-info-box">' +
    '<div class="d-flex justify-content-between mb-2 mt-1 align-items-start">' +
    '<div class="d-flex flex-column w-75 align-items-start pt-1">' +
    '<div class="d-flex w-100 mb-2 align-items-start">' +
    '<span class="text-muted me-2 mt-1">Paziente:</span>' +
    '<div class="editable-area plain-text border-bottom w-100 fw-bold fs-4 text-uppercase"' +
    ' contenteditable="true" data-field="Nome" style="min-height:40px;padding:2px;">' + (p.Nome || '') + '</div>' +
    '</div>' +
    '<div class="d-flex w-100 align-items-start">' +
    '<span class="text-muted me-2 mt-1 fw-normal text-nowrap">Diagnosi / Motivo Ricovero:</span>' +
    '<div class="editable-area plain-text border-bottom w-100 me-5 fw-bold"' +
    ' contenteditable="true" data-field="Diagnosi" style="min-height:30px;padding:2px;">' + (p.Diagnosi || '') + '</div>' +
    '</div>' +
    '</div>' +
    '<div class="d-flex flex-column align-items-end me-3" style="min-width:220px;">' +
    '<div class="d-flex align-items-center mb-1 text-secondary w-100 justify-content-end" style="font-size:0.95rem;">' +
    '<span class="fw-bold me-2">Data di Nascita:</span>' +
    '<input type="text" class="data-nascita-text" placeholder="gg/mm/aaaa"' +
    ' oninput="formattaDataNascita(this)" onblur="formattaDataNascitaBlur(this)" readonly' +
    ' value="' + _aEsc(nasc.vis) + '"' +
    ' style="border:none;border-bottom:1px solid #999;outline:none;background:transparent;padding:2px;width:110px;text-align:right;font-size:0.95rem;">' +
    '</div>' +
    '<div class="d-flex align-items-center mb-1 w-100 justify-content-end">' +
    '<span class="text-muted me-2">Et&agrave;:</span>' +
    '<div class="editable-area plain-text border-bottom text-end"' +
    ' contenteditable="' + (nasc.hasData ? 'false' : 'true') + '" data-field="Eta"' +
    ' style="min-height:25px;padding:2px;width:50px;' + (nasc.hasData ? 'opacity:0.6;cursor:not-allowed;' : '') + '">' + eta + '</div>' +
    '</div>' +
    '<div class="d-flex align-items-center mb-1 text-secondary w-100 justify-content-end" style="font-size:0.95rem;">' +
    '<span class="fw-bold me-2">Data Ricovero:</span>' +
    '<input type="text" class="data-ricovero-text" placeholder="gg/mm/aaaa"' +
    ' oninput="formattaDataRicovero(this)" onblur="formattaDataRicoveroBlur(this)" readonly' +
    ' value="' + _aEsc(ric.vis) + '"' +
    ' style="border:none;border-bottom:1px solid #999;outline:none;background:transparent;padding:2px;width:110px;text-align:right;font-size:0.95rem;">' +
    '</div>' +
    '<div class="d-flex align-items-center text-secondary w-100 justify-content-end" style="font-size:0.95rem;">' +
    '<span class="fw-bold me-2">Giorni di Ricovero:</span>' +
    '<span class="valore-giorni text-danger fw-bold" style="width:20px;text-align:right;">' + ric.giorni + '</span>' +
    '</div>' +
    '<div class="d-flex align-items-center text-secondary w-100 justify-content-end mt-1" style="font-size:0.95rem;">' +
    '<span class="fw-bold me-2">C.S.:</span>' +
    '<div class="editable-area plain-text border-bottom text-end" contenteditable="true" data-field="CodiceSanitario"' +
    ' style="min-height:22px;padding:2px;min-width:100px;">' + (p.CodiceSanitario || '') + '</div>' +
    '</div>' +
    '</div></div></div></div>' +
    '<div class="row m-0 content-row">' +
    '<div class="column-box col-terapia p-0">' +
    '<div class="column-header"><i class="bi bi-exclamation-triangle-fill text-danger me-1" style="font-size:0.75rem;"></i>ALLERGIE</div>' +
    '<div class="editable-area rich-text" contenteditable="true" data-field="Allergie">' + (p.Allergie || '') + '</div>' +
    '<div class="column-header" style="border-top:1px solid #ccc;"><i class="bi bi-wind me-1 text-info" style="font-size:0.75rem;"></i>OSSIGENO</div>' +
    '<div class="editable-area plain-text" contenteditable="true" data-field="Ossigeno" style="padding:4px;">' + (p.Ossigeno || '') + '</div>' +
    '<div class="column-header" style="border-top:2px solid #333;">NOTE E TERAPIA</div>' +
    '<div class="editable-area rich-text" contenteditable="true" data-field="NoteTerapia">' + (p.NoteTerapia || '') + '</div>' +
    '</div>' +
    '<div class="column-box col-diaria p-0">' +
    '<div class="column-header">DIARIA ED EPICRISI</div>' +
    '<div class="editable-area rich-text" contenteditable="true" data-field="Diaria">' + (p.Diaria || '') + '</div>' +
    '</div>' +
    '<div class="column-box col-da-fare p-0">' +
    '<div class="column-header">DA FARE / RICHIESTE</div>' +
    '<div class="editable-area rich-text" contenteditable="true" data-field="DaFare">' + (p.DaFare || '') + '</div>' +
    '</div>' +
    '</div>' +
    '<div class="row m-0 piano-terapeutico-row p-0">' +
    '<div class="column-header-piano">PIANO TERAPEUTICO</div>' +
    '<div class="editable-area editable-area-piano rich-text" contenteditable="true" data-field="PianoTerapeutico">' + (p.PianoTerapeutico || '') + '</div>' +
    '</div>' +
    '</div>';
}


// ── RENDERER CARD ALTERNATIVA ─────────────────────────────────
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
      ' ondblclick="_apriModalTipologia(\'' + _aEsc(letto) + '\')"' +
      ' title="Doppio click per cambiare tipologia">' + _aEsc(tipo) + '</span>'
    : '<span class="badge bg-light text-secondary border text-uppercase shadow-sm"' +
      ' style="cursor:pointer;"' +
      ' ondblclick="_apriModalTipologia(\'' + _aEsc(letto) + '\')"' +
      ' title="Doppio click per cambiare tipologia">STANDARD</span>';

  return '<div class="alt-row patient-card" data-bed="' + _aEsc(letto) + '" data-tipologia="' + _aEsc(tipo) + '">' +
    '<div class="alt-col alt-col-info">' +
    '<button class="focus-pencil-btn" title="Modifica scheda"' +
    ' onclick="_attivaFocusMode(this.closest(\'.patient-card\'))"><i class="bi bi-pencil-fill"></i></button>' +
    '<span id="status-alt-' + _aEsc(letto) + '" class="badge status-badge position-absolute top-0 start-0 m-1 bg-secondary"></span>' +
    '<div class="alt-info-box">' +
    '<div class="bed-number-flex">' +
    '<div class="alt-bed-number">' + _aEsc(letto) + '</div>' +
    '<span class="sesso-symbol" data-field="Sesso" data-sesso="' + _aEsc(p.Sesso || '') + '"></span>' +
    '</div>' +
    '<div class="text-center tipo-badge-wrap" style="font-size:0.75rem;">' + badgeHtml + '</div>' +
    '<div class="editable-area plain-text alt-nome" contenteditable="true" data-field="Nome">' + (p.Nome || '') + '</div>' +
    '<div class="alt-allergie-label"><i class="bi bi-exclamation-triangle-fill"></i> Allergie</div>' +
    '<div class="editable-area rich-text alt-allergie-val" contenteditable="true" data-field="Allergie">' + (p.Allergie || '') + '</div>' +
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
    '<div class="editable-area plain-text alt-info-val" contenteditable="true" data-field="Ossigeno">' + (p.Ossigeno || '') + '</div></div>' +
    '</div></div>' +
    '<div class="alt-col alt-col-diag">' +
    '<div class="alt-diag-top">' +
    '<div class="alt-col-header">Diagnosi / Motivo Ricovero</div>' +
    '<div class="editable-area plain-text alt-editable" contenteditable="true" data-field="Diagnosi">' + (p.Diagnosi || '') + '</div>' +
    '</div>' +
    '<div class="alt-diag-bottom">' +
    '<div class="alt-col-header-split">Piano Terapeutico</div>' +
    '<div class="editable-area rich-text alt-editable" contenteditable="true" data-field="PianoTerapeutico">' + (p.PianoTerapeutico || '') + '</div>' +
    '</div>' +
    '<div class="alt-diag-third">' +
    '<div class="alt-col-header-split">Da Fare / Richieste</div>' +
    '<div class="editable-area rich-text alt-editable" contenteditable="true" data-field="DaFare">' + (p.DaFare || '') + '</div>' +
    '</div>' +
    '</div>' +
    '<div class="alt-col alt-col-diaria">' +
    '<div class="alt-col-header">Diaria ed Epicrisi</div>' +
    '<div class="editable-area rich-text alt-editable" contenteditable="true" data-field="Diaria">' + (p.Diaria || '') + '</div>' +
    '</div>' +
    '<div class="alt-col alt-col-terapia">' +
    '<div class="alt-col-header">Note e Terapia</div>' +
    '<div class="editable-area rich-text alt-editable" contenteditable="true" data-field="NoteTerapia">' + (p.NoteTerapia || '') + '</div>' +
    '</div>' +
    '<div class="alt-row-spacer"></div>' +
    '</div>';
}


// ── RENDERER SCHEDA NOTE (layout principale) ──────────────────
function _renderNoteCardMain(p) {
  return '<div class="patient-card note-card shadow-sm" data-bed="NOTE">' +
    '<div class="row m-0 header-row">' +
    '<div class="col-2 bed-number-box p-3">' +
    '<button class="focus-pencil-btn" title="Modifica note"' +
    ' onclick="_attivaFocusMode(this.closest(\'.patient-card\'))"><i class="bi bi-pencil-fill"></i></button>' +
    '<span id="status-NOTE" class="badge status-badge position-absolute top-0 start-0 m-1 bg-secondary"></span>' +
    '<div class="text-muted fw-bold mt-2" style="font-size:0.75rem;">NOTE</div>' +
    '<div class="bed-number-flex">' +
    '<div class="bed-number" style="font-size:1.5rem;letter-spacing:1px;">NOTE</div>' +
    '</div>' +
    '</div>' +
    '<div class="col-10 patient-info-box d-flex align-items-stretch p-0">' +
    '<div class="editable-area rich-text w-100" contenteditable="true" data-field="Diaria"' +
    ' style="min-height:300px;padding:10px;font-size:0.9rem;">' + (p.Diaria || '') + '</div>' +
    '</div>' +
    '</div>' +
    '</div>';
}

// ── RENDERER SCHEDA NOTE (layout alternativo) ─────────────────
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
    ' style="min-height:300px;padding:10px;font-size:0.9rem;">' + (p.Diaria || '') + '</div>' +
    '</div>' +
    '<div class="alt-row-spacer"></div>' +
    '</div>';
}

// ── RENDERER COMPLETO (entrambe le viste) ─────────────────────
function _renderCardsHtml(pazienti) {
  var main = '<div class="container-fluid" id="cardsContainer">';
  var alt  = '<div class="container-fluid" id="cardsContainerAlt">';
  var noteP = null;
  (pazienti || []).forEach(function(p) {
    if (p.Letto === 'NOTE') { noteP = p; return; }
    main += _renderMainCard(p);
    alt  += _renderAltCard(p);
  });
  // NOTE sempre in fondo
  if (noteP) {
    main += _renderNoteCardMain(noteP);
    alt  += _renderNoteCardAlt(noteP);
  }
  main += '<div id="noLettiMsg" class="container text-center mt-5"' +
          ' style="display:none;"><i class="bi bi-inboxes text-muted" style="font-size:4rem;"></i>' +
          '<h4 class="mt-3 text-secondary">Nessun Letto configurato.</h4></div></div>';
  alt  += '<div id="noLettiMsgAlt" class="container text-center mt-5"' +
          ' style="display:none;"><i class="bi bi-inboxes text-muted" style="font-size:4rem;"></i>' +
          '<h4 class="mt-3 text-secondary">Nessun Letto configurato.</h4></div></div>';
  return main + alt;
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
  return _q(_sb.from('consegne').update(row).eq('letto', String(letto)))
    .then(function() { return { success: true, ora: row.ultimo_aggiornamento }; });
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

function _sbAggiungiLetto(numeroLetto) {
  return _q(_sb.from('consegne').select('letto').eq('letto', String(numeroLetto)).maybeSingle())
    .then(function(existing) {
      if (existing) return { success: false, message: 'Il letto esiste già!' };
      return _q(_sb.from('consegne').insert({ letto: String(numeroLetto), tipologia_letto: 'STANDARD' }))
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
  var campiVuoti = {
    nome:'', eta:'', data_nascita:'', data_ricovero:'', diagnosi:'',
    note_terapia:'', diaria:'', da_fare:'', piano_terapeutico:'',
    allergie:'', codice_sanitario:'', ossigeno:'', sesso:'',
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
        'note_terapia','diaria','da_fare','piano_terapeutico','allergie',
        'codice_sanitario','ossigeno','sesso'];

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
  return _sbGetPazienti().then(function(pazienti) {
    var uomini = 0, donne = 0, indefinito = 0, vuoti = 0;
    var tipologie = {};
    pazienti.forEach(function(p) {
      if (p.Letto === 'NOTE') return; // non contare NOTE come letto
      var hasPaziente = (p.Nome || '').trim() !== '';
      if (!hasPaziente) { vuoti++; return; }
      var sesso = (p.Sesso || '').toUpperCase();
      if (sesso === 'M') uomini++;
      else if (sesso === 'F') donne++;
      else indefinito++;
      var tip = (p.TipologiaLetto || 'STANDARD').toUpperCase();
      tipologie[tip] = (tipologie[tip] || 0) + 1;
    });
    return { uomini: uomini, donne: donne, indefinito: indefinito, tipologie: tipologie, vuoti: vuoti };
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

var BACKUP_INTERVALLO_MS = 6 * 60 * 60 * 1000; // 6 ore

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
  return _q(_sb.from('consegne').update({ tipologia_letto: nuovaTipologia }).eq('letto', String(letto)))
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

// Costruisce HTML per una singola scheda letto — 3 colonne, larghezza 100%.
// Struttura compatibile Google Docs: no nested tables, colspan minimo.
//   Riga 1 (header): C1=Letto (17%) | C2+C3=Nome/Diagnosi/Meta (colspan=2, 83%)
//   Riga 2 (corpo):  C1=Allergie/Ossigeno/Terapia (17%) | C2=Diaria (63%) | C3=DaFare (20%)
//   Riga 3 (piano):  C1+C2+C3=Piano Terapeutico (colspan=3, 100%)
function _driveRenderCard(p) {
  var tipo = (p.TipologiaLetto || '').trim().toUpperCase() || 'STANDARD';

  var meta = '';
  if (p.DataNascita)     meta += 'Nasc.: <b>' + p.DataNascita + '</b> &nbsp; ';
  if (p.Eta)             meta += 'Et&agrave;: <b>' + p.Eta + '</b> &nbsp; ';
  if (p.DataRicovero)    meta += 'Ricovero: <b>' + p.DataRicovero + '</b> &nbsp; ';
  if (p.CodiceSanitario) meta += 'C.S.: <b>' + p.CodiceSanitario + '</b>';

  var B  = '1.5pt solid #333333';
  var Bi = '1pt solid #888888';
  var HB = 'background-color:#e8e6e1;font-weight:bold;font-size:8pt;text-transform:uppercase;padding:3pt 5pt;border-bottom:' + Bi + ';';
  var CB = 'padding:4pt 6pt;font-size:9pt;';

  return (
    '<table width="100%" style="border-collapse:collapse;margin-bottom:10pt;font-family:Arial,sans-serif;font-size:9pt;">' +

    // ── Riga 1: header ──────────────────────────────────────────────────────
    '<tr>' +
      '<td width="17%" style="border:' + B + ';text-align:center;vertical-align:middle;background-color:#f2efe9;padding:6pt;">' +
        '<p style="font-size:22pt;font-weight:bold;margin:0;line-height:1;">' + (p.Letto || '') + '</p>' +
        '<p style="font-size:7pt;font-weight:bold;color:#546e7a;text-transform:uppercase;margin:2pt 0 0;">' + tipo + '</p>' +
      '</td>' +
      '<td colspan="2" width="83%" style="border:' + B + ';vertical-align:top;padding:5pt 8pt;">' +
        '<p style="font-size:12pt;font-weight:bold;text-transform:uppercase;border-bottom:1pt solid #999;padding-bottom:2pt;margin:0 0 4pt;">' + (p.Nome || '') + '</p>' +
        '<p style="font-weight:bold;margin:0 0 3pt;">' + (p.Diagnosi || '') + '</p>' +
        (meta ? '<p style="font-size:8pt;color:#444;margin:3pt 0 0;">' + meta + '</p>' : '') +
      '</td>' +
    '</tr>' +

    // ── Riga 2: corpo ───────────────────────────────────────────────────────
    '<tr>' +
      '<td width="17%" style="border:' + B + ';vertical-align:top;padding:0;">' +
        '<p style="' + HB + '">&#x26A0; Allergie</p>' +
        '<p style="' + CB + '">' + (p.Allergie || '') + '</p>' +
        '<p style="' + HB + 'border-top:' + Bi + ';">Ossigeno</p>' +
        '<p style="' + CB + '">' + (p.Ossigeno || '') + '</p>' +
        '<p style="' + HB + 'border-top:' + B + ';">Note e Terapia</p>' +
        '<p style="' + CB + '">' + (p.NoteTerapia || '') + '</p>' +
      '</td>' +
      '<td width="63%" style="border:' + B + ';vertical-align:top;padding:0;">' +
        '<p style="' + HB + '">Diaria ed Epicrisi</p>' +
        '<p style="' + CB + '">' + (p.Diaria || '') + '</p>' +
      '</td>' +
      '<td width="20%" style="border:' + B + ';vertical-align:top;padding:0;">' +
        '<p style="' + HB + '">Da Fare / Richieste</p>' +
        '<p style="' + CB + '">' + (p.DaFare || '') + '</p>' +
      '</td>' +
    '</tr>' +

    // ── Riga 3: piano terapeutico ────────────────────────────────────────────
    '<tr>' +
      '<td colspan="3" width="100%" style="border:' + B + ';vertical-align:top;padding:0;">' +
        '<p style="' + HB + 'text-align:left;">Piano Terapeutico</p>' +
        '<p style="' + CB + '">' + (p.PianoTerapeutico || '') + '</p>' +
      '</td>' +
    '</tr>' +

    '</table>'
  );
}

function _driveRenderNoteCard(p) {
  var B = '1.5pt solid #546e7a';
  var HB = 'background-color:#eceff1;font-weight:bold;font-size:8pt;text-transform:uppercase;padding:3pt 5pt;border-bottom:1pt solid #90a4ae;';
  var CB = 'padding:4pt 6pt;font-size:9pt;';
  return (
    '<table width="100%" style="border-collapse:collapse;margin-bottom:10pt;font-family:Arial,sans-serif;font-size:9pt;border:' + B + ';">' +
    '<tr>' +
      '<td width="17%" style="border:' + B + ';text-align:center;vertical-align:middle;background-color:#eceff1;padding:6pt;">' +
        '<p style="font-size:1rem;font-weight:bold;letter-spacing:2px;color:#546e7a;margin:0;">NOTE</p>' +
      '</td>' +
      '<td width="83%" style="border:' + B + ';vertical-align:top;padding:0;">' +
        '<p style="' + CB + '">' + (p.Diaria || '') + '</p>' +
      '</td>' +
    '</tr>' +
    '</table>'
  );
}

// Esegue backup su Google Drive come Google Doc con layout standard.
// Le cartelle vengono create la prima volta e l'ID salvato in Supabase.
// Fire-and-forget: non blocca il flusso principale.
function _driveBackupConsegne(pazienti, ts) {
  if (!window._googleDriveToken) {
    console.log('[Drive backup] Nessun token Drive disponibile — solo Supabase.');
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

  // Costruisce HTML con layout standard (tabelle — Drive converte in Google Doc)
  var cardsHtml = ordinati.map(function(p) {
    return p.Letto === 'NOTE' ? _driveRenderNoteCard(p) : _driveRenderCard(p);
  }).join('');

  var html = '<!DOCTYPE html><html><head><meta charset="UTF-8">' +
    '<style>' +
    '@page{size:A4 landscape;margin:15mm 20mm 15mm 20mm;}' +
    'body{font-family:Arial,sans-serif;font-size:9pt;margin:0;padding:0;}' +
    'p{margin:0;padding:0;}' +
    '</style>' +
    '</head><body>' +
    '<table width="100%" style="margin-bottom:12pt;border-bottom:2pt solid #333;">' +
    '<tr>' +
    '<td style="font-size:14pt;font-weight:bold;padding-bottom:4pt;">Consegne Reparto</td>' +
    '<td style="text-align:right;font-size:9pt;color:#555;padding-bottom:4pt;">Backup del ' + dataLabel + '</td>' +
    '</tr></table>' +
    cardsHtml +
    '</body></html>';

  // Cerca o crea la cartella nella root, crea il file, poi pulisce i vecchi
  return _driveGetOrCreateFolder('BACKUP CONSEGNE EMERGENZA')
    .then(function(folderId) {
      return _driveCreaGoogleDoc(nomeFile, html, folderId)
        .then(function(file) {
          console.log('[Drive backup] Google Doc creato:', (file && file.name), (file && file.id));
          // Pulizia file vecchi con la stessa retention di Supabase archivio
          return _sbGetGiorniConservazione().then(function(giorni) {
            var cutoff = Number(ts) - (giorni * 86400000);
            _driveEliminaVecchi(folderId, cutoff); // fire-and-forget
          });
        });
    })
    .catch(function(e) { console.warn('[Drive backup] Errore (non bloccante):', e.message || e); });
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
      function() { _scheduleRealtimeSync(); })
    // Lock: gestiti direttamente in-memory, nessuna query aggiuntiva al DB
    .on('postgres_changes', { event: 'INSERT', schema: 'public', table: 'locks' },
      function(payload) { _onLockChange('INSERT', payload); })
    .on('postgres_changes', { event: 'UPDATE', schema: 'public', table: 'locks' },
      function(payload) { _onLockChange('UPDATE', payload); })
    .on('postgres_changes', { event: 'DELETE', schema: 'public', table: 'locks' },
      function(payload) { _onLockChange('DELETE', payload); })
    .subscribe(function(status) {
      console.log('[Realtime]', status);
    });
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

function _scheduleRealtimeSync() {
  if (_pendingSync) return; // già schedulato
  _pendingSync = true;
  clearTimeout(_realtimeTimer);
  _realtimeTimer = setTimeout(function() {
    _pendingSync = false;

    var ind = document.getElementById('syncIndicator');
    var st  = document.getElementById('syncStatus');
    if (ind) ind.className = 'badge bg-warning text-dark ms-2';
    if (st)  st.innerText  = 'Sync...';

    _sbGetPazienti().then(function(pazienti) {
      if (typeof _applicaAggiornamentoCompleto === 'function') {
        _applicaAggiornamentoCompleto(_renderCardsHtml(pazienti));
      }
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
// EMULAZIONE google.script.run  (interfaccia invariata per app.js / app2.js)
// ══════════════════════════════════════════════════════════════
(function() {

  function Runner(ok, err) {
    this._ok  = ok  || function() {};
    this._err = err || function(e) { console.error('API error:', e); };
  }
  Runner.prototype.withSuccessHandler = function(fn) { return new Runner(fn, this._err); };
  Runner.prototype.withFailureHandler = function(fn) { return new Runner(this._ok, fn); };

  function wrap(promise, ok, err) {
    promise.then(ok).catch(function(e) { err({ message: e.message || String(e) }); });
  }

  // ── Lettura dati ─────────────────────────────────────────────
  Runner.prototype.getDatiPazienti = function() {
    var ok = this._ok, err = this._err;
    wrap(_sbGetPazienti(), ok, err);
  };
  Runner.prototype.getLettiFull = function() {
    var ok = this._ok, err = this._err;
    wrap(_sbGetLettiFull(), ok, err);
  };
  Runner.prototype.getLocks = function() {
    var ok = this._ok, err = this._err;
    wrap(_sbGetLocks(), ok, err);
  };
  Runner.prototype.getRiepilogoLetti = function() {
    var ok = this._ok, err = this._err;
    wrap(_sbGetRiepilogo(), ok, err);
  };
  Runner.prototype.getNewHtml = function() {
    var ok = this._ok, err = this._err;
    wrap(_sbGetPazienti().then(function(p) { return _renderCardsHtml(p); }), ok, err);
  };

  // ── Archivio ─────────────────────────────────────────────────
  Runner.prototype.archiviaGiornoCorrente = function() {
    var ok = this._ok, err = this._err;
    wrap(_sbArchiviaGiornoCorrente(), ok, err);
  };
  Runner.prototype.getGiorniArchivio = function() {
    var ok = this._ok, err = this._err;
    // Ritorna il numero di giorni di conservazione (impostazione)
    wrap(_sbGetGiorniConservazione(), ok, err);
  };
  Runner.prototype.getGiorniArchiviati = function() {
    var ok = this._ok, err = this._err;
    // Ritorna l'elenco delle date archiviate
    wrap(_sbGetGiorniArchivio(), ok, err);
  };
  Runner.prototype.getDatiArchivioGiorno = function(dataStr) {
    var ok = this._ok, err = this._err;
    wrap(_sbGetDatiArchivioGiorno(dataStr), ok, err);
  };
  Runner.prototype.getTimestampGiorno = function(dataStr) {
    var ok = this._ok, err = this._err;
    // Ritorna array di epoch-ms (string) per tutti i backup di quel giorno
    wrap(_sbGetTimestampsGiorno(dataStr), ok, err);
  };
  Runner.prototype.salvaGiorniArchivio = function(val) {
    var ok = this._ok, err = this._err;
    wrap(_sbSalvaGiorniConservazione(val), ok, err);
  };
  Runner.prototype.pulisciArchivioVecchio = function() { /* gestito in archiviaGiornoCorrente */ };
  Runner.prototype.pulisciBackupEmergenzaVecchi = function() { /* no-op */ };

  // ── Backup / status (no-op, non più necessario) ───────────────
  Runner.prototype.checkBackupStatus  = function() { if (this._ok) this._ok({ inCorso: false }); };
  Runner.prototype.rinnovaLockBackup  = function() { /* no-op */ };

  // ── Salvataggio paziente ──────────────────────────────────────
  Runner.prototype.autoSavePazienteCompleto = function(letto, datiPaziente) {
    var ok = this._ok, err = this._err;
    wrap(_sbSalvaPaziente(letto, datiPaziente), ok, err);
  };

  // ── Lock ─────────────────────────────────────────────────────
  Runner.prototype.acquistaLock = function(letto, token) {
    var ok = this._ok, err = this._err;
    wrap(_sbAcquistaLock(letto, token), ok, err);
  };
  Runner.prototype.rilasciaLock = function(letto, token) {
    var ok = this._ok, err = this._err;
    wrap(_sbRilasciaLock(letto, token), ok, err);
  };
  Runner.prototype.acquistaLockMultiplo = function(letti, token) {
    var ok = this._ok, err = this._err;
    wrap(_sbAcquistaLockMultiplo(letti, token), ok, err);
  };
  Runner.prototype.rilasciaLockMultiplo = function(letti, token) {
    var ok = this._ok, err = this._err;
    wrap(_sbRilasciaLockMultiplo(letti, token), ok, err);
  };

  // ── Gestione letti ────────────────────────────────────────────
  Runner.prototype.aggiungiLetto = function(numeroLetto) {
    var ok = this._ok, err = this._err;
    wrap(_sbAggiungiLetto(numeroLetto), ok, err);
  };
  Runner.prototype.eliminaLetto = function(numeroLetto) {
    var ok = this._ok, err = this._err;
    wrap(_sbEliminaLetto(numeroLetto), ok, err);
  };
  Runner.prototype.dimettiLetto = function(numeroLetto) {
    var ok = this._ok, err = this._err;
    wrap(_sbDimettiLetto(numeroLetto), ok, err);
  };
  Runner.prototype.spostaPaziente = function(lettoOrigine, lettoDestinazione) {
    var ok = this._ok, err = this._err;
    wrap(_sbSpostaPaziente(lettoOrigine, lettoDestinazione), ok, err);
  };

  // ── Tipologie ─────────────────────────────────────────────────
  Runner.prototype.getColoriTipologie = function() {
    var ok = this._ok, err = this._err;
    wrap(_sbGetColoriTipologie(), ok, err);
  };
  Runner.prototype.getTipologieConfigurate = function() {
    var ok = this._ok, err = this._err;
    // Ritorna array [{nome, colore}] come si aspetta app2.js
    wrap(_sbGetColoriTipologie().then(function(mappa) {
      return Object.keys(mappa).filter(Boolean).map(function(nome) {
        return { nome: nome, colore: mappa[nome] || stringToColor(nome) };
      });
    }), ok, err);
  };
  Runner.prototype.getDatiLettiConTipologia = function() {
    var ok = this._ok, err = this._err;
    wrap(_sbGetDatiLettiConTipologia(), ok, err);
  };
  Runner.prototype.getTipologieLettiBed = function() {
    var ok = this._ok, err = this._err;
    wrap(_sbGetTipologieLettiBed(), ok, err);
  };
  Runner.prototype.salvaColoriTipologie = function(mappa) {
    var ok = this._ok, err = this._err;
    wrap(_sbSalvaColoriTipologie(mappa), ok, err);
  };
  Runner.prototype.salvaTipologieBatch = function(modifiche) {
    var ok = this._ok, err = this._err;
    wrap(_sbSalvaTipologieBatch(modifiche), ok, err);
  };
  Runner.prototype.eliminaTipologiaConfigurata = function(nome, force) {
    var ok = this._ok, err = this._err;
    wrap(_sbEliminaTipologia(nome, force), ok, err);
  };
  Runner.prototype.cambiaTipologiaALetto = function(letto, nuovaTipologia) {
    var ok = this._ok, err = this._err;
    wrap(_sbCambiaTipologiaALetto(letto, nuovaTipologia), ok, err);
  };
  Runner.prototype.modificaTipologiaLetto = function(letto, nuovaTipologia) {
    var ok = this._ok, err = this._err;
    wrap(_sbCambiaTipologiaALetto(letto, nuovaTipologia), ok, err);
  };

  // ── Link utili ────────────────────────────────────────────────
  Runner.prototype.getLinkUtili = function() {
    var ok = this._ok, err = this._err;
    wrap(_sbGetLinkUtili(), ok, err);
  };
  Runner.prototype.aggiungiLinkUtile = function(nome, url) {
    var ok = this._ok, err = this._err;
    wrap(_sbAggiungiLink(nome, url), ok, err);
  };
  Runner.prototype.modificaLinkUtile = function(indice, nome, url) {
    var ok = this._ok, err = this._err;
    wrap(_sbModificaLink(indice, nome, url), ok, err);
  };
  Runner.prototype.eliminaLinkUtile = function(indice) {
    var ok = this._ok, err = this._err;
    wrap(_sbEliminaLink(indice), ok, err);
  };

  // ── Impostazioni ─────────────────────────────────────────────
  Runner.prototype.ottieniNomeReparto = function() {
    var ok = this._ok, err = this._err;
    wrap(_sbOttieniNomeReparto(), ok, err);
  };
  Runner.prototype.salvaNomeReparto = function(nuovoNome) {
    var ok = this._ok, err = this._err;
    wrap(_sbSalvaNomeReparto(nuovoNome), ok, err);
  };

  // Espone google.script.run globalmente
  window.google = window.google || {};
  window.google.script = window.google.script || {};
  window.google.script.run = new Runner();

})();
