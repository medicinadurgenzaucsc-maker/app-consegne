// ============================================================
// sanifica.js — l'HTML e il testo che arrivano da fuori
// ============================================================
// Tutto ciò che non è scritto nel codice dell'app (righe del database,
// payload Realtime, backup, memoria locale, appunti, parametri dell'URL)
// può contenere HTML ostile: chi possiede una sessione può salvare in un
// campo qualunque stringa. Questo file è l'UNICO punto in cui quel materiale
// viene reso innocuo prima di finire in pagina.
//
// Va caricato DOPO DOMPurify e PRIMA di api.js, in index.html e in print.html.
//
//   _pulisciHtml(html)    campi a testo ricco delle schede. Resta solo ciò che
//                         l'app sa produrre: elenco CHIUSO di tag, attributi,
//                         classi e proprietà di stile (ricavato dalle schede
//                         vere con collaudo/strumenti/inventario-markup.js).
//   _pulisciHtmlConfig(h) testi scritti da chi gestisce l'app (definizioni delle
//                         scale): formattazione ed elenchi sì, nulla di attivo.
//   _testoHtml(s)         testo semplice dentro una stringa HTML (anche dentro
//                         un attributo, con virgolette o con apici).
//   _urlSicuro(u)         solo http, https, mailto, tel: il resto diventa ''.
//   _coloreSicuro(c, r)   un colore CSS valido, altrimenti il ripiego r.
//   _cssVal(v)            un valore dentro un selettore: [data-bed="…"].
//   _lettoValido(nome)    il nome di un letto è un'etichetta semplice?
//   _sanificaPronta()     false se DOMPurify non si è caricato: in quel caso
//                         l'app NON disegna e NON salva le schede (api.js).
//
// REGOLA PER CHI SCRIVE CODICE NUOVO: una stringa che non nasce nel codice non
// si concatena mai dentro innerHTML / Swal html / title così com'è. O è testo
// (-> _testoHtml) o è HTML di un campo (-> _pulisciHtml). Mai dentro un
// gestore scritto in linea (onclick="f('…')"): lì la codifica HTML non
// protegge. Se una funzione dell'app comincia a scrivere dentro i campi un
// tag, una classe o uno stile nuovi, vanno aggiunti agli elenchi qui sotto,
// altrimenti spariscono (e la prova «formattazione che l'app produce» di
// collaudo/pagina/prove.js lo segnala).
(function () {
  'use strict';

  // ── Elenchi di ciò che può restare dentro un campo ─────────────────────
  var TAG_AMMESSI = [
    // testo e blocchi
    'div', 'span', 'br', 'p', 'b', 'strong', 'i', 'em', 'u', 's', 'strike', 'del', 'ins',
    'sub', 'sup', 'small', 'big', 'mark', 'font', 'ul', 'ol', 'li', 'blockquote', 'pre',
    'code', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'hr',
    // tabelle (incollate da altri programmi)
    'table', 'thead', 'tbody', 'tfoot', 'tr', 'td', 'th', 'caption', 'colgroup', 'col',
    // bottoni dei blocchi gestiti dall'app (problemi attivi, laboratorio)
    'button',
    // icone disegnate dall'app (alambicco, virus, cuore, spunte)
    'svg', 'g', 'path', 'circle', 'rect', 'line', 'polyline', 'polygon', 'ellipse'
  ];
  var ATTRIBUTI_AMMESSI = [
    'class', 'style', 'title', 'contenteditable', 'type', 'role', 'aria-hidden', 'aria-checked',
    // dati dei blocchi gestiti dall'app
    'data-ord', 'data-preg', 'data-ts', 'data-scala', 'data-isolamento', 'data-rianimazione',
    // <font> e tabelle
    'color', 'face', 'size', 'colspan', 'rowspan',
    // icone
    'viewbox', 'width', 'height', 'd', 'fill', 'stroke', 'stroke-width', 'stroke-linecap',
    'stroke-linejoin', 'cx', 'cy', 'r', 'x', 'y', 'rx', 'ry', 'x1', 'y1', 'x2', 'y2', 'points'
  ];
  // Attributi che contengono testo libero scritto dall'utente (il tipo di
  // isolamento, la nota di rianimazione…): DOMPurify scarterebbe un valore
  // come «da contatto: KPC» scambiandolo per un indirizzo con schema ignoto.
  // Nessuno di questi viene mai usato come indirizzo.
  var ATTRIBUTI_DI_TESTO = ['data-isolamento', 'data-rianimazione', 'data-scala', 'data-ts', 'data-ord', 'data-preg', 'face'];
  // Le sole classi che l'app scrive dentro i campi. Tutte le altre si tolgono:
  // con le classi di Bootstrap (position-fixed, w-100, modal…) un contenuto
  // ostile potrebbe coprire la pagina o travestirsi da finestra dell'app, e
  // con «editable-area» fingersi un altro campo della scheda.
  var CLASSE_AMMESSA = new RegExp('^(?:' + [
    'bi', 'bi-[a-z0-9-]+',
    'pa-(?:add|add-btn|item|txt|az|mod|del)',
    'lab-(?:box|apri-btn|apri-txt|remind|remind-off)',
    'scale-box', 'scala-risultato',
    'terapia-(?:box|box-head|placeholder)', 'pregressi-(?:box|head)', 'note-hdr',
    'diaria-status-box', 'isolamento-banner', 'rianimazione-banner',
    'cl-(?:item|check|esito|esito-arrow|esito-sbarrato|sbarrato|separator)'
  ].join('|') + ')$');
  // Solo aspetto del testo: niente position, z-index, dimensioni, trasformazioni.
  var STILE_AMMESSO = {
    'color': 1, 'background-color': 1,
    'font-size': 1, 'font-weight': 1, 'font-style': 1, 'font-family': 1, 'font-variant': 1,
    'text-decoration': 1, 'text-decoration-line': 1, 'text-decoration-color': 1,
    'text-decoration-style': 1, 'text-decoration-thickness': 1,
    'text-align': 1, 'text-indent': 1, 'text-transform': 1, 'letter-spacing': 1,
    'word-spacing': 1, 'line-height': 1, 'vertical-align': 1, 'white-space': 1,
    'padding-left': 1, 'padding-right': 1, 'padding-top': 1, 'padding-bottom': 1,
    'margin-left': 1, 'margin-right': 1, 'margin-top': 1, 'margin-bottom': 1,
    '-webkit-tap-highlight-color': 1, '-webkit-text-size-adjust': 1, 'text-size-adjust': 1
  };
  // Dentro il valore di una proprietà: niente indirizzi, funzioni attive, escape.
  var VALORE_VIETATO = /url\s*\(|expression\s*\(|image\s*\(|image-set|cross-fade|element\s*\(|javascript\s*:|behaviou?r\s*:|-moz-binding|@|\\|<|>|\/\*/i;
  // Colore di riempimento/tratto delle icone e del tag <font>.
  var COLORE = /^(?:none|currentcolor|transparent|#[0-9a-f]{3,8}|[a-z]{3,30}|(?:rgb|rgba|hsl|hsla)\(\s*[0-9.,%\s\/a-z-]{1,60}\))$/i;

  // ── Testo, indirizzi, colori, selettori, nomi dei letti ────────────────
  var ENTITA = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' };
  function testoHtml(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) { return ENTITA[c]; });
  }
  function urlSicuro(u) {
    var s = String(u == null ? '' : u).trim();
    if (!s) return '';
    try {
      var p = new URL(s, window.location.href);
      return (/^(https?|mailto|tel):$/i.test(p.protocol)) ? p.href : '';
    } catch (e) { return ''; }
  }
  function coloreSicuro(c, ripiego) {
    var s = String(c == null ? '' : c).trim();
    if (/^#(?:[0-9a-f]{3,4}|[0-9a-f]{6}|[0-9a-f]{8})$/i.test(s)) return s;
    if (/^(?:rgb|rgba|hsl|hsla)\(\s*[0-9.,%\s\/a-z-]{1,60}\)$/i.test(s)) return s;
    if (/^[a-z]{3,30}$/i.test(s)) return s;
    return ripiego == null ? 'transparent' : ripiego;
  }
  // Valore da mettere fra virgolette dentro un selettore: [data-bed="…"].
  function cssVal(v) {
    var s = String(v == null ? '' : v);
    return (window.CSS && typeof window.CSS.escape === 'function') ? window.CSS.escape(s) : s.replace(/["\\\n\r\f]/g, '\\$&');
  }
  // Il nome di un letto finisce in selettori, attributi e chiavi di oggetti:
  // deve restare un'etichetta semplice (lettere maiuscole, cifre, spazio e
  // . _ / + -), al massimo 30 caratteri. È la regola di «Aggiungi letto».
  var LETTO_VALIDO = /^[A-Z0-9À-ÖØ-Þ][A-Z0-9À-ÖØ-Þ ._\/+-]{0,29}$/;
  function lettoValido(nome) { return LETTO_VALIDO.test(String(nome == null ? '' : nome)); }

  // ── Filtro dell'HTML dei campi ─────────────────────────────────────────
  var pulitore = null, tolti = null;
  try {
    if (window.DOMPurify && window.DOMPurify.isSupported) {
      // Istanza a parte: ganci e configurazione non toccano altri usi della libreria.
      pulitore = window.DOMPurify(window);
      pulitore.addHook('afterSanitizeAttributes', function (nodo) {
        if (!nodo || nodo.nodeType !== 1 || !nodo.getAttribute) return;
        // classi: resta solo l'elenco dell'app
        var classi = nodo.getAttribute('class');
        if (classi != null) {
          var lista = classi.split(/\s+/).filter(Boolean);
          var buone = lista.filter(function (c) { return CLASSE_AMMESSA.test(c); });
          if (buone.length !== lista.length) {
            if (tolti) lista.forEach(function (c) { if (!CLASSE_AMMESSA.test(c)) tolti.classi[c.slice(0, 40)] = 1; });
            if (buone.length) nodo.setAttribute('class', buone.join(' ')); else nodo.removeAttribute('class');
          }
        }
        // stile: resta solo l'aspetto del testo. Il testo dell'attributo si
        // riscrive SOLO se va tolto qualcosa: un contenuto pulito resta identico.
        var stile = nodo.getAttribute('style');
        if (stile != null) {
          var dichiarazioni = stile.split(';'), tenute = [], scartata = false;
          dichiarazioni.forEach(function (d) {
            if (!d.trim()) return;
            var i = d.indexOf(':');
            var nome = i > 0 ? d.slice(0, i).trim().toLowerCase() : '';
            var valore = i > 0 ? d.slice(i + 1).trim() : '';
            if (STILE_AMMESSO[nome] === 1 && valore && !VALORE_VIETATO.test(valore)) tenute.push(d.trim());
            else { scartata = true; if (tolti) tolti.stili[(nome || '?').slice(0, 40)] = 1; }
          });
          if (scartata) {
            if (tenute.length) nodo.setAttribute('style', tenute.join('; ') + ';'); else nodo.removeAttribute('style');
          }
        }
        // contenteditable: solo acceso o spento
        var ce = nodo.getAttribute('contenteditable');
        if (ce != null && ce !== 'false' && ce !== 'true' && ce !== '') nodo.removeAttribute('contenteditable');
        // colori di icone e <font>: un colore, non un riferimento a una risorsa
        ['fill', 'stroke', 'color'].forEach(function (a) {
          var v = nodo.getAttribute(a);
          if (v != null && !COLORE.test(v.trim())) { nodo.removeAttribute(a); if (tolti) tolti.stili[a] = 1; }
        });
      });
    }
  } catch (e) { pulitore = null; }

  var CONFIGURAZIONE = {
    ALLOWED_TAGS: TAG_AMMESSI,
    ALLOWED_ATTR: ATTRIBUTI_AMMESSI,
    ADD_URI_SAFE_ATTR: ATTRIBUTI_DI_TESTO,
    ALLOW_DATA_ATTR: false,
    ALLOW_ARIA_ATTR: false,
    ALLOW_UNKNOWN_PROTOCOLS: false,
    KEEP_CONTENT: true
  };

  // I campi si ripuliscono a ogni sincronizzazione: la stessa stringa non si
  // rifiltra due volte.
  var memoria = new Map(), MEMORIA_MAX = 400;
  var segnalati = {}, nSegnalati = 0;

  function pronta() { return !!pulitore; }

  // Se il filtro ha tolto qualcosa lo si annota nel registro dell'app, con i
  // soli NOMI di ciò che è stato tolto (mai il contenuto): serve ad accorgersi
  // sia di un tentativo ostile sia di una formattazione lecita che l'elenco
  // qui sopra non prevede. Al massimo 20 segnalazioni per pagina aperta.
  function segnala(riepilogo) {
    try {
      if (nSegnalati >= 20 || segnalati[riepilogo]) return;
      segnalati[riepilogo] = 1; nSegnalati++;
      if (typeof window._log === 'function') window._log('warning', 'html-ripulito', 'Tolto dal contenuto di un campo: ' + riepilogo, null, null);
    } catch (e) {}
  }

  function pulisciHtml(html) {
    var s = String(html == null ? '' : html);
    if (s === '') return '';
    // Senza segni di tag né di entità non c'è nulla da interpretare: è già
    // testo che il browser mostrerà così com'è.
    if (s.indexOf('<') < 0 && s.indexOf('&') < 0) return s;
    if (!pulitore) return testoHtml(s);          // senza libreria: tutto a testo
    var gia = memoria.get(s);
    if (gia !== undefined) return gia;
    tolti = { classi: {}, stili: {} };
    var pulito;
    try {
      pulito = String(pulitore.sanitize(s, CONFIGURAZIONE));
      var parti = [];
      (pulitore.removed || []).forEach(function (r) {
        if (r && r.element && r.element.nodeName) {
          var tag = String(r.element.nodeName).toLowerCase();
          // <body> è il contenitore di lavoro della libreria: compare sempre
          // fra i «tolti» e non dice nulla sul contenuto del campo.
          if (tag !== 'body' && tag !== 'html' && tag !== 'head') parti.push('<' + tag + '>');
        }
        else if (r && r.attribute && r.attribute.name) parti.push(String(r.attribute.name).toLowerCase() + '=');
      });
      Object.keys(tolti.classi).forEach(function (c) { parti.push('classe ' + c); });
      Object.keys(tolti.stili).forEach(function (c) { parti.push('stile ' + c); });
      if (parti.length) {
        var unici = parti.filter(function (p, i) { return parti.indexOf(p) === i; }).sort();
        segnala(unici.slice(0, 25).join(' '));
      }
    } catch (e) {
      pulito = testoHtml(s);
    }
    tolti = null;
    if (pulito === s) pulito = s;                // una sola copia in memoria
    if (memoria.size >= MEMORIA_MAX) memoria.delete(memoria.keys().next().value);
    memoria.set(s, pulito);
    return pulito;
  }

  // ── Testi scritti da chi gestisce l'app (definizioni delle scale) ──────
  // Non sono modificabili dagli utenti, ma restano una copia nella memoria del
  // browser: formattazione, elenchi, tabelle e classi sì; niente di attivo.
  var pulitoreConfig = null;
  try { if (pulitore) pulitoreConfig = window.DOMPurify(window); } catch (e) { pulitoreConfig = null; }
  var CONFIGURAZIONE_TESTI = {
    ALLOWED_TAGS: ['div', 'span', 'p', 'br', 'b', 'strong', 'i', 'em', 'u', 'small', 'sub', 'sup',
      'ul', 'ol', 'li', 'table', 'thead', 'tbody', 'tr', 'td', 'th', 'h4', 'h5', 'h6', 'hr'],
    ALLOWED_ATTR: ['class', 'colspan', 'rowspan', 'title'],
    ALLOW_DATA_ATTR: false,
    ALLOW_ARIA_ATTR: false,
    KEEP_CONTENT: true
  };
  function pulisciHtmlConfig(html) {
    var s = String(html == null ? '' : html);
    if (s.indexOf('<') < 0 && s.indexOf('&') < 0) return s;
    if (!pulitoreConfig) return testoHtml(s);
    try { return String(pulitoreConfig.sanitize(s, CONFIGURAZIONE_TESTI)); } catch (e) { return testoHtml(s); }
  }

  window._pulisciHtml = pulisciHtml;
  window._pulisciHtmlConfig = pulisciHtmlConfig;
  window._testoHtml = testoHtml;
  window._urlSicuro = urlSicuro;
  window._coloreSicuro = coloreSicuro;
  window._cssVal = cssVal;
  window._lettoValido = lettoValido;
  window._sanificaPronta = pronta;
})();
