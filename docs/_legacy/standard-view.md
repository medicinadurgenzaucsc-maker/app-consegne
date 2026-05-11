# Vista Standard — Codice salvato (commit pre-rimozione)

Questa è la vista "standard" delle card pazienti, rimossa dall'app
il 2026-05-11 in favore della sola vista alternativa (più compatta e
informativa). Il codice è conservato qui per eventuale ri-integrazione
futura.

Rimosso in: commit con messaggio "Solo vista alternativa: rimozione
codice standard, alleggerimento bundle"

## Quando riattivare

Per ri-integrare:
1. Reincollare `_renderMainCard` e `_renderNoteCardMain` in `docs/js/api.js`
2. Ripristinare la logica `_renderCardsHtml` originale (vedi sotto)
3. Reincollare `_inizializzaView` / `toggleViewAlt` / `var _viewAltAttiva` in `docs/js/app.js`
4. Re-aggiungere la voce menu `<li id="btnToggleViewAlt">` in `docs/index.html`
5. Verificare riferimenti `cardsContainer` vs `cardsContainerAlt` nei file modificati
6. Re-aggiungere campo `EsamiColturali` anche in `_renderMainCard` (NON presente nel backup originale, era stato aggiunto solo alla vista alt)

---

## `_renderMainCard` (in `docs/js/api.js`, ex righe 168-268)

```javascript
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
```

## `_renderNoteCardMain` (in `docs/js/api.js`, ex righe 354-372)

```javascript
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
```

## `_renderCardsHtml` originale (in `docs/js/api.js`, ex righe 394-415)

```javascript
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
```

## `_inizializzaView` + `toggleViewAlt` + `_viewAltAttiva` (in `docs/js/app.js`, ex righe 1307-1333)

```javascript
var _viewAltAttiva = false;

function _inizializzaView() {
  var saved = localStorage.getItem('viewAlt');
  // Default: visualizzazione alternativa (null = prima volta, non ancora salvata)
  _viewAltAttiva = saved === null ? true : (saved === '1');
  var main = document.getElementById('cardsContainer');
  var alt  = document.getElementById('cardsContainerAlt');
  var btn  = document.querySelector('#btnToggleViewAlt');
  if (_viewAltAttiva) {
    if (main) main.classList.add('nascosto');
    if (alt)  alt.classList.add('attiva');
    if (btn)  { btn.classList.add('attiva'); btn.style.setProperty('background-color', '#212529', 'important'); btn.style.setProperty('color', '#fff', 'important'); btn.style.setProperty('border-radius', '4px', 'important'); }
  } else {
    if (main) main.classList.remove('nascosto');
    if (alt)  alt.classList.remove('attiva');
    if (btn)  { btn.classList.remove('attiva'); btn.style.removeProperty('background-color'); btn.style.removeProperty('color'); btn.style.removeProperty('border-radius'); }
  }
  applicaOrdinamentoSalvato();
}

// GitHub Pages: toggle diretto senza reload
function toggleViewAlt() {
  var nuovoStato = localStorage.getItem('viewAlt') === '1' ? '0' : '1';
  localStorage.setItem('viewAlt', nuovoStato);
  _inizializzaView();
}
```

## Voce menu HTML (in `docs/index.html`)

```html
<li><a class="dropdown-item" id="btnToggleViewAlt" href="javascript:void(0)" onclick="toggleViewAlt()"><i class="bi bi-eye me-2"></i>Visualizzazione Alternativa</a></li>
```

## CSS specifiche (in `docs/css/styles.css`)

```css
#btnToggleViewAlt.attiva {
  background-color: #212529 !important;
  color: #fff !important;
  border-radius: 4px !important;
}
#btnToggleViewAlt.attiva:hover,
#btnToggleViewAlt.attiva:focus { background-color: #343a40 !important; color: #fff !important; }
#btnToggleViewAlt.attiva:hover i,
#btnToggleViewAlt.attiva:focus i { color: #fff !important; }
```

## Note di funzionamento

- La vista standard occupava più spazio (card piene larghe) ma aveva un layout
  diverso da quella alternativa (alt-row = riga compatta con 4 colonne).
- Le card della vista standard usavano `class="patient-card shadow-sm"` (senza `.alt-row`).
- Le card NOTE main avevano `class="patient-card note-card shadow-sm"` (senza `.alt-row`).
- L'ID dei badge status era `status-{letto}` (standard) vs `status-alt-{letto}` (alt).
- L'ID dei badge tipo era `badge-tipo-{letto}` (standard) vs `badge-tipo-alt-{letto}` (alt).
- `_renderCardsHtml` generava SEMPRE entrambe le viste e usava CSS classes
  `nascosto`/`attiva` per mostrarne una alla volta.
- `_inizializzaView` leggeva da `localStorage.viewAlt` (default true).
- Riferimenti a `cardsContainer` rimangono nel codice per i selettori DOM
  globali (l'ID viene riusato per la singola vista alt).
