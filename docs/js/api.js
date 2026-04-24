// ============================================================
// api.js — Livello di comunicazione con il backend GAS
// Sostituisce google.script.run con chiamate fetch() REST
// ============================================================

// ── CONFIGURAZIONE ────────────────────────────────────────────
// GAS_URL: URL del tuo deployment Google Apps Script
// (Trova in GAS: Deploy → Gestisci deployment → copia URL)
var GAS_URL = 'https://script.google.com/macros/s/AKfycbz9c6PiBHwjGJcajbXsBMAt1Slrehp7GIuOEXqiC8perArb2QkxxBSBPox_f0P_owyl/exec';

// APP_URL: URL di questa app su GitHub Pages (per i reload)
var APP_URL = 'https://medicinadurgenzaucsc-maker.github.io/app-consegne/';

// PRINT_URL: URL della pagina di stampa GitHub Pages
var PRINT_URL = APP_URL + 'print.html';

// API_TOKEN: token segreto da impostare anche in GAS Script Properties
// (GAS: Impostazioni progetto → Proprietà script → chiave "API_TOKEN")
// Se lasciato vuoto, il backend accetta tutte le richieste.
var API_TOKEN = '';


// ── HELPERS FETCH ─────────────────────────────────────────────
function _apiGet(action, params) {
  var url = GAS_URL + '?action=' + encodeURIComponent(action);
  if (API_TOKEN) url += '&token=' + encodeURIComponent(API_TOKEN);
  if (params) {
    Object.keys(params).forEach(function(k) {
      if (params[k] !== undefined && params[k] !== null) {
        url += '&' + encodeURIComponent(k) + '=' + encodeURIComponent(params[k]);
      }
    });
  }
  return fetch(url).then(function(r) {
    if (!r.ok) throw new Error('HTTP ' + r.status);
    return r.json();
  });
}

function _apiPost(action, body) {
  var url = GAS_URL + '?action=' + encodeURIComponent(action);
  if (API_TOKEN) url += '&token=' + encodeURIComponent(API_TOKEN);
  var payload = Object.assign({}, body || {});
  return fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=UTF-8' },
    body: JSON.stringify(payload)
  }).then(function(r) {
    if (!r.ok) throw new Error('HTTP ' + r.status);
    return r.json();
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


// ── COLORE DA STRINGA (identico a stringToColor in Index.html) ─
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
// Chiamato da index.html dopo il caricamento della pagina
function _caricaCardIniziali(container, onDone) {
  _apiGet('getDatiPazienti', {}).then(function(pazienti) {
    if (container) container.innerHTML = _renderCardsHtml(pazienti);
    if (typeof onDone === 'function') onDone(pazienti);
  }).catch(function(e) {
    console.error('Errore caricamento card:', e);
    if (typeof onDone === 'function') onDone([]);
  });
}


// ── LETTURA TOAST DA URL (sostituisce template GAS) ───────────
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


// ── EMULAZIONE google.script.run ──────────────────────────────
(function() {

  function Runner(ok, err) {
    this._ok  = ok  || function() {};
    this._err = err || function(e) { console.error('API error:', e); };
  }

  Runner.prototype.withSuccessHandler = function(fn) {
    return new Runner(fn, this._err);
  };
  Runner.prototype.withFailureHandler = function(fn) {
    return new Runner(this._ok, fn);
  };

  // ── GET senza parametri ──────────────────────────────────────
  ['getLocks','getLettiFull','getRiepilogoLetti','getGiorniArchiviati',
   'getGiorniArchivio','checkBackupStatus','getColoriTipologie',
   'getTipologieConfigurate','getDatiLettiConTipologia',
   'getTipologieLettiBed','getLinkUtili','ottieniNomeReparto',
   'getDatiPazienti'
  ].forEach(function(name) {
    Runner.prototype[name] = function() {
      var ok = this._ok, err = this._err;
      _apiGet(name, {})
        .then(ok)
        .catch(function(e) { err({ message: e.message }); });
    };
  });

  // ── GET con parametri ────────────────────────────────────────
  Runner.prototype.getTimestampGiorno = function(dataStr) {
    var ok = this._ok, err = this._err;
    _apiGet('getTimestampGiorno', { dataStr: dataStr })
      .then(ok).catch(function(e) { err({ message: e.message }); });
  };
  Runner.prototype.getDatiArchivioGiorno = function(dataStr) {
    var ok = this._ok, err = this._err;
    _apiGet('getDatiArchivioGiorno', { dataStr: dataStr })
      .then(ok).catch(function(e) { err({ message: e.message }); });
  };

  // ── getNewHtml: chiama getDatiPazienti e renderizza client-side
  Runner.prototype.getNewHtml = function() {
    var ok = this._ok, err = this._err;
    _apiGet('getDatiPazienti', {}).then(function(pazienti) {
      ok(_renderCardsHtml(pazienti));
    }).catch(function(e) { err({ message: e.message }); });
  };

  // ── POST calls ───────────────────────────────────────────────
  Runner.prototype.autoSavePazienteCompleto = function(letto, datiPaziente, token) {
    var ok = this._ok, err = this._err;
    _apiPost('autoSavePazienteCompleto', { letto: letto, datiPaziente: datiPaziente, token: token })
      .then(ok).catch(function(e) { err({ message: e.message }); });
  };
  Runner.prototype.acquistaLock = function(letto, token) {
    var ok = this._ok, err = this._err;
    _apiPost('acquistaLock', { letto: letto, token: token })
      .then(ok).catch(function(e) { err({ message: e.message }); });
  };
  Runner.prototype.rilasciaLock = function(letto, token) {
    var ok = this._ok, err = this._err;
    _apiPost('rilasciaLock', { letto: letto, token: token })
      .then(ok).catch(function(e) { err({ message: e.message }); });
  };
  Runner.prototype.acquistaLockMultiplo = function(letti, token) {
    var ok = this._ok, err = this._err;
    _apiPost('acquistaLockMultiplo', { letti: letti, token: token })
      .then(ok).catch(function(e) { err({ message: e.message }); });
  };
  Runner.prototype.rilasciaLockMultiplo = function(letti, token) {
    var ok = this._ok, err = this._err;
    _apiPost('rilasciaLockMultiplo', { letti: letti, token: token })
      .then(ok).catch(function(e) { err({ message: e.message }); });
  };
  Runner.prototype.aggiungiLetto = function(numeroLetto) {
    var ok = this._ok, err = this._err;
    _apiPost('aggiungiLetto', { numeroLetto: numeroLetto })
      .then(ok).catch(function(e) { err({ message: e.message }); });
  };
  Runner.prototype.eliminaLetto = function(numeroLetto) {
    var ok = this._ok, err = this._err;
    _apiPost('eliminaLetto', { numeroLetto: numeroLetto })
      .then(ok).catch(function(e) { err({ message: e.message }); });
  };
  Runner.prototype.dimettiLetto = function(numeroLetto) {
    var ok = this._ok, err = this._err;
    _apiPost('dimettiLetto', { numeroLetto: numeroLetto })
      .then(ok).catch(function(e) { err({ message: e.message }); });
  };
  Runner.prototype.spostaPaziente = function(lettoOrigine, lettoDestinazione) {
    var ok = this._ok, err = this._err;
    _apiPost('spostaPaziente', { lettoOrigine: lettoOrigine, lettoDestinazione: lettoDestinazione })
      .then(ok).catch(function(e) { err({ message: e.message }); });
  };
  Runner.prototype.modificaTipologiaLetto = function(letto, nuovaTipologia) {
    var ok = this._ok, err = this._err;
    _apiPost('modificaTipologiaLetto', { letto: letto, nuovaTipologia: nuovaTipologia })
      .then(ok).catch(function(e) { err({ message: e.message }); });
  };
  Runner.prototype.rinnovaLockBackup = function() {
    _apiPost('rinnovaLockBackup', {}).catch(function() {});
  };
  Runner.prototype.archiviaGiornoCorrente = function() {
    var ok = this._ok, err = this._err;
    _apiPost('archiviaGiornoCorrente', {})
      .then(ok).catch(function(e) { err({ message: e.message }); });
  };
  Runner.prototype.salvaGiorniArchivio = function(giorni) {
    var ok = this._ok, err = this._err;
    _apiPost('salvaGiorniArchivio', { giorni: giorni })
      .then(ok).catch(function(e) { err({ message: e.message }); });
  };
  Runner.prototype.pulisciArchivioVecchio = function() {
    _apiPost('pulisciArchivioVecchio', {}).catch(function() {});
  };
  Runner.prototype.pulisciBackupEmergenzaVecchi = function() {
    _apiPost('pulisciBackupEmergenzaVecchi', {}).catch(function() {});
  };
  Runner.prototype.salvaColoriTipologie = function(mappa) {
    var ok = this._ok, err = this._err;
    _apiPost('salvaColoriTipologie', { mappa: mappa })
      .then(ok).catch(function(e) { err({ message: e.message }); });
  };
  Runner.prototype.salvaTipologieBatch = function(modifiche) {
    var ok = this._ok, err = this._err;
    _apiPost('salvaTipologieBatch', { modifiche: modifiche })
      .then(ok).catch(function(e) { err({ message: e.message }); });
  };
  Runner.prototype.eliminaTipologiaConfigurata = function(nome, force) {
    var ok = this._ok, err = this._err;
    _apiPost('eliminaTipologiaConfigurata', { nome: nome, force: force })
      .then(ok).catch(function(e) { err({ message: e.message }); });
  };
  Runner.prototype.cambiaTipologiaALetto = function(letto, nuovaTipologia) {
    var ok = this._ok, err = this._err;
    _apiPost('cambiaTipologiaALetto', { letto: letto, nuovaTipologia: nuovaTipologia })
      .then(ok).catch(function(e) { err({ message: e.message }); });
  };
  Runner.prototype.aggiungiLinkUtile = function(nome, url) {
    var ok = this._ok, err = this._err;
    _apiPost('aggiungiLinkUtile', { nome: nome, url: url })
      .then(ok).catch(function(e) { err({ message: e.message }); });
  };
  Runner.prototype.modificaLinkUtile = function(indice, nome, url) {
    var ok = this._ok, err = this._err;
    _apiPost('modificaLinkUtile', { indice: indice, nome: nome, url: url })
      .then(ok).catch(function(e) { err({ message: e.message }); });
  };
  Runner.prototype.eliminaLinkUtile = function(indice) {
    var ok = this._ok, err = this._err;
    _apiPost('eliminaLinkUtile', { indice: indice })
      .then(ok).catch(function(e) { err({ message: e.message }); });
  };
  Runner.prototype.salvaNomeReparto = function(nuovoNome) {
    var ok = this._ok, err = this._err;
    _apiPost('salvaNomeReparto', { nuovoNome: nuovoNome })
      .then(ok).catch(function(e) { err({ message: e.message }); });
  };

  // Espone google.script.run globalmente
  window.google = window.google || {};
  window.google.script = window.google.script || {};
  window.google.script.run = new Runner();

})();
