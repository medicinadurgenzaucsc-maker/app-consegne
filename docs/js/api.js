// ============================================================
// api.js — Backend Supabase v2
// Sostituisce GAS + Google Sheets con Supabase (Postgres + Realtime)
// ============================================================


// ── CONFIGURAZIONE ────────────────────────────────────────────
var SUPABASE_URL     = 'https://ifmmcvxzhwdkmzhsxcvb.supabase.co';
var SUPABASE_ANON_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImlmbW1jdnh6aHdka216aHN4Y3ZiIiwicm9sZSI6ImFub24iLCJpYXQiOjE3Nzc2Nzk1ODAsImV4cCI6MjA5MzI1NTU4MH0.LH8h4Fivtl3-TuiA050oF8iS4b80xrd2Dn6z8JjCoeA';
var APP_URL          = 'https://medicinadurgenzaucsc-maker.github.io/app-consegne/';
var PRINT_URL        = APP_URL + 'print.html';
var LOCK_TTL_MS      = 30000; // 30 secondi

// ── CLIENT SUPABASE ───────────────────────────────────────────
var _sb = supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY);


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
    '</div>';
}


// ── RENDERER COMPLETO (entrambe le viste) ─────────────────────
function _renderCardsHtml(pazienti) {
  var main = '<div class="container-fluid" id="cardsContainer">';
  var alt  = '<div class="container-fluid" id="cardsContainerAlt">';
  (pazienti || []).forEach(function(p) {
    main += _renderMainCard(p);
    alt  += _renderAltCard(p);
  });
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
      var nA = parseInt(a.Letto, 10), nB = parseInt(b.Letto, 10);
      return (!isNaN(nA) && !isNaN(nB)) ? nA - nB : String(a.Letto).localeCompare(String(b.Letto));
    });
    return pazienti;
  });
}

function _sbGetLettiFull() {
  return _q(_sb.from('consegne').select('letto,nome,tipologia_letto')).then(function(rows) {
    var letti = (rows || []).map(function(r) {
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
    (rows || []).forEach(function(r) {
      if (now - Number(r.ts) <= LOCK_TTL_MS) {
        locks[r.letto] = { token: r.token, ts: Number(r.ts) };
      }
    });
    return locks;
  });
}

function _sbAcquistaLock(letto, token) {
  var k = String(letto);
  var now = Date.now();
  return _sbGetLocks().then(function(locks) {
    var existing = locks[k];
    if (existing && existing.token !== token) {
      return { success: false, blocked: true, message: 'Scheda in aggiornamento da altro utente.' };
    }
    return _q(_sb.from('locks').upsert({ letto: k, token: token, ts: now }, { onConflict: 'letto' }))
      .then(function() { return { success: true, blocked: false }; });
  }).catch(function() { return { success: false, blocked: false, message: 'Errore server, riprova.' }; });
}

function _sbRilasciaLock(letto, token) {
  return _q(_sb.from('locks').delete().eq('letto', String(letto)).eq('token', token))
    .then(function() { return { success: true }; })
    .catch(function() { return { success: false }; });
}

function _sbAcquistaLockMultiplo(letti, token) {
  return _sbGetLocks().then(function(locks) {
    var bloccati = letti.filter(function(l) {
      var ex = locks[String(l)];
      return ex && ex.token !== token;
    });
    if (bloccati.length > 0) {
      return { success: false, bloccati: bloccati, message: 'Letti bloccati da altro utente.' };
    }
    var now = Date.now();
    var upserts = letti.map(function(l) { return { letto: String(l), token: token, ts: now }; });
    return _q(_sb.from('locks').upsert(upserts, { onConflict: 'letto' }))
      .then(function() { return { success: true }; });
  }).catch(function() { return { success: false, bloccati: [], message: 'Errore server.' }; });
}

function _sbRilasciaLockMultiplo(letti, token) {
  return _q(_sb.from('locks').delete().in('letto', letti.map(String)).eq('token', token))
    .then(function() { return { success: true }; })
    .catch(function() { return { success: false }; });
}


// ══════════════════════════════════════════════════════════════
// ARCHIVIO GIORNALIERO
// ══════════════════════════════════════════════════════════════

function _sbArchiviaGiornoCorrente() {
  var dataStr = _oggiStr();
  return _q(_sb.from('archivio').select('id').eq('data_str', dataStr).maybeSingle())
    .then(function(existing) {
      if (existing) return { inCorso: false }; // già archiviato oggi
      return _sbGetPazienti().then(function(pazienti) {
        return _q(_sb.from('archivio').insert({
          data_str: dataStr,
          ts: Date.now(),
          dati: pazienti
        })).then(function() { return { inCorso: false }; });
      });
    })
    .then(function() {
      // Pulizia archivio oltre 90 giorni
      var limit = Date.now() - (90 * 86400000);
      _q(_sb.from('archivio').delete().lt('ts', limit)).catch(function() {});
      return { inCorso: false };
    })
    .catch(function() { return { inCorso: false }; });
}

function _sbGetGiorniArchivio() {
  return _q(_sb.from('archivio').select('data_str,ts').order('ts', { ascending: false }))
    .then(function(rows) { return (rows || []).map(function(r) { return r.data_str; }); });
}

function _sbGetDatiArchivioGiorno(dataStr) {
  return _q(_sb.from('archivio').select('dati,ts').eq('data_str', dataStr).maybeSingle())
    .then(function(row) {
      if (!row) return null;
      return { pazienti: row.dati || [], timestamp: row.ts };
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
    if (m.azione === 'rinomina') {
      return Promise.all([
        _q(_sb.from('tipologie').update({ nome: m.nuovoNome, colore: m.colore || '' }).eq('nome', m.vecchioNome)),
        _q(_sb.from('consegne').update({ tipologia_letto: m.nuovoNome })
          .eq('tipologia_letto', m.vecchioNome))
      ]);
    } else if (m.azione === 'colore') {
      return _q(_sb.from('tipologie').upsert({ nome: m.nome, colore: m.colore }, { onConflict: 'nome' }));
    }
    return Promise.resolve();
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
  return _q(_sb.from('consegne').select('letto,tipologia_letto')).then(function(rows) {
    var mappa = {};
    (rows || []).forEach(function(r) {
      mappa[r.letto] = (r.tipologia_letto || 'STANDARD').toUpperCase();
    });
    return mappa;
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


// ══════════════════════════════════════════════════════════════
// REALTIME — aggiornamenti istantanei senza polling
// ══════════════════════════════════════════════════════════════

var _realtimeTimer = null;
var _realtimeChannel = null;

function _inizializzaRealtime() {
  if (_realtimeChannel) { _sb.removeChannel(_realtimeChannel); }

  _realtimeChannel = _sb.channel('consegne-live')
    .on('postgres_changes', { event: '*', schema: 'public', table: 'consegne' },
      function() { _scheduleRealtimeSync('consegne'); })
    .on('postgres_changes', { event: '*', schema: 'public', table: 'locks' },
      function() { _scheduleRealtimeSync('locks'); })
    .subscribe(function(status) {
      console.log('[Realtime]', status);
    });
}

var _pendingSync = { consegne: false, locks: false };

function _scheduleRealtimeSync(tipo) {
  _pendingSync[tipo] = true;
  clearTimeout(_realtimeTimer);
  _realtimeTimer = setTimeout(function() {
    var syncConsegne = _pendingSync.consegne;
    var syncLocks    = _pendingSync.locks;
    _pendingSync = { consegne: false, locks: false };

    var ind = document.getElementById('syncIndicator');
    var st  = document.getElementById('syncStatus');
    if (ind) ind.className = 'badge bg-warning text-dark ms-2';
    if (st)  st.innerText  = 'Sync...';

    var promesse = [];
    if (syncConsegne) {
      promesse.push(_sbGetPazienti().then(function(pazienti) {
        if (typeof _applicaAggiornamentoCompleto === 'function') {
          _applicaAggiornamentoCompleto(_renderCardsHtml(pazienti));
        }
      }));
    }
    if (syncLocks) {
      promesse.push(_sbGetLocks().then(function(locks) {
        if (typeof _applicaLocks === 'function') { _applicaLocks(locks); }
      }));
    }
    Promise.all(promesse)
      .then(function() {
        if (ind) ind.className = 'badge bg-success ms-2';
        if (st)  st.innerText  = new Date().toLocaleTimeString('it-IT');
      })
      .catch(function(e) {
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
    wrap(_sbGetGiorniArchivio().then(function(g) { return { giorni: g }; }), ok, err);
  };
  Runner.prototype.getGiorniArchiviati = function() {
    var ok = this._ok, err = this._err;
    wrap(_sbGetGiorniArchivio(), ok, err);
  };
  Runner.prototype.getDatiArchivioGiorno = function(dataStr) {
    var ok = this._ok, err = this._err;
    wrap(_sbGetDatiArchivioGiorno(dataStr), ok, err);
  };
  Runner.prototype.getTimestampGiorno = function(dataStr) {
    var ok = this._ok, err = this._err;
    wrap(_sbGetDatiArchivioGiorno(dataStr).then(function(r) {
      return r ? { ts: r.timestamp } : { ts: null };
    }), ok, err);
  };
  Runner.prototype.salvaGiorniArchivio = function() {
    // no-op: l'archivio ora è su Supabase
    if (this._ok) this._ok({ ok: true });
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
    wrap(_sbGetTipologieConfigurate(), ok, err);
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
