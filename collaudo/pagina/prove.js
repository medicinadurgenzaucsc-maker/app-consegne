// Prove automatiche dell'app. Girano DENTRO l'app vera, nel banco di prova
// locale (http://localhost:8765/) collegato al database di COLLAUDO, a reparto
// finto appena generato (generatore-pazienti.js). Usano le funzioni e
// l'interfaccia dell'app, i segnalibri veri e il finto TrakCare.
//
//   await window.__prove()                         tutte le sezioni
//   await window.__prove({ sezioni: ['trak'] })    solo alcune
//
// Ogni sezione rimette i dati com'erano. Esito: { prove, superate, fallite, elenco }.
(function () {
  if (typeof AMBIENTE === 'undefined' || AMBIENTE !== 'collaudo') { console.error('prove: consentite solo nel collaudo'); return; }
  var REF_PRODUZIONE = 'ifmmcvxzhwdkmzhsxcvb';
  var esiti = [], C = null;

  // ── attrezzi ─────────────────────────────────────────────────────────
  var attendi = function (ms) { return new Promise(function (r) { setTimeout(r, ms); }); };
  async function finche(cond, ms, passo) {
    var t0 = Date.now();
    while (Date.now() - t0 < ms) {
      var v = false;
      try { v = await cond(); } catch (e) {}
      if (v) return v;
      await attendi(passo || 150);
    }
    return false;
  }
  // una prova è superata se la funzione restituisce true (o nulla); una stringa è il motivo del fallimento
  async function prova(sezione, nome, fn) {
    var t0 = Date.now(), esito;
    try { esito = await fn(); } catch (e) { esito = 'eccezione: ' + ((e && e.message) || e); }
    var ok = esito === true || esito === undefined;
    esiti.push({ sezione: sezione, nome: nome, ok: ok, dettaglio: ok ? '' : String(esito).slice(0, 600), ms: Date.now() - t0 });
    return ok;
  }
  function carica(src) {
    return new Promise(function (ok, ko) {
      var s = document.createElement('script');
      s.src = src; s.onload = ok; s.onerror = function () { ko(new Error('non riesco a caricare ' + src)); };
      document.head.appendChild(s);
    });
  }
  function scheda(letto) { return document.querySelector('.patient-card[data-bed="' + (window.CSS && CSS.escape ? CSS.escape(letto) : letto) + '"]'); }
  var uguali = function (a, b) { return JSON.stringify(a) === JSON.stringify(b); };
  var copia = function (o) { return JSON.parse(JSON.stringify(o)); };
  function primaDifferenza(a, b) {
    var sa = JSON.stringify(a), sb = JSON.stringify(b), i = 0;
    while (i < sa.length && sa[i] === sb[i]) i++;
    return 'atteso «…' + sb.slice(Math.max(0, i - 40), i + 60) + '…» trovato «…' + sa.slice(Math.max(0, i - 40), i + 60) + '…»';
  }

  // ── finto TrakCare dentro una cornice, segnalibri veri ───────────────
  function cornice(url) {
    return new Promise(function (ok, ko) {
      var casa = document.getElementById('proveCornici');
      if (!casa) { casa = document.createElement('div'); casa.id = 'proveCornici'; casa.style.cssText = 'position:fixed;left:-9000px;top:0;width:1300px;height:900px'; document.body.appendChild(casa); }
      var f = document.createElement('iframe');
      f.style.cssText = 'width:1280px;height:860px;border:0;display:block';
      f.onload = function () { ok(f); };
      f.onerror = function () { ko(new Error('pagina non caricata: ' + url)); };
      f.src = url;
      casa.appendChild(f);
    });
  }
  function avvisoDi(win) {
    var b = win.document.getElementById('bmLabBanner');
    if (b) return b.textContent;
    var u = win.document.body.lastElementChild;
    return (u && /position:\s*fixed/.test(u.getAttribute('style') || '')) ? u.textContent : '';
  }
  // Esegue il segnalibro nella finestra indicata, come farebbe il clic sulla
  // barra dei preferiti; restituisce ciò che avrebbe copiato negli appunti
  // ({copiato}) oppure l'avviso mostrato ({avviso}).
  function premi(win, segnalibro, ms) {
    return new Promise(function (ok) {
      var chiuso = false, t0 = Date.now(), iv = null;
      function fine(v) { if (chiuso) return; chiuso = true; if (iv) clearInterval(iv); ok(v); }
      // via gli avvisi lasciati da una pressione precedente: non vanno scambiati per quello nuovo
      try {
        [].slice.call(win.document.body.children).forEach(function (el) {
          if (el.id === 'bmLabBanner' || /z-index:\s*2147483647/.test(el.getAttribute('style') || '')) el.remove();
        });
      } catch (e) {}
      try { win.navigator.clipboard.writeText = function (t) { fine({ copiato: String(t) }); return Promise.resolve(); }; } catch (e) { return fine({ errore: 'appunti non sostituibili: ' + e.message }); }
      try { win.eval(segnalibro.replace(/^javascript:/, '')); } catch (e) { return fine({ errore: e.message }); }
      if (chiuso) return;
      iv = setInterval(function () {
        var t = avvisoDi(win);
        if (/⚠️/.test(t)) return fine({ avviso: t });
        if (Date.now() - t0 > ms) fine({ errore: 'tempo scaduto', avviso: t });
      }, 200);
    });
  }
  function leggiPayload(esito, marcatore) {
    if (!esito || !esito.copiato) return null;
    var t = esito.copiato.trim();
    if (t.indexOf(marcatore) !== 0) return null;
    try { return JSON.parse(t.slice(marcatore.length)); } catch (e) { return null; }
  }
  function esamiDi(pagine) { var out = []; (pagine || []).forEach(function (pg) { out.push.apply(out, _parseLabRaw(pg)); }); return out; }

  // ═════════════════════════════════════════════════════════════════════
  // AMBIENTE: la pagina sta davvero lavorando sul collaudo, e solo lì
  // ═════════════════════════════════════════════════════════════════════
  async function sezioneAmbiente() {
    var S = 'ambiente';
    await prova(S, 'l\'app si dichiara «collaudo» e punta al progetto di collaudo', function () {
      return (AMBIENTE === 'collaudo' && SUPABASE_URL.indexOf(REF_PRODUZIONE) < 0 && /rqvohwpthhumydpbwktq/.test(SUPABASE_URL)) || ('AMBIENTE=' + AMBIENTE + ' url=' + SUPABASE_URL);
    });
    await prova(S, 'cornice «COLLAUDO» visibile e titolo marcato', function () {
      var f = document.getElementById('fasciaCollaudo');
      return (!!f && getComputedStyle(f).display !== 'none' && /^\[COLLAUDO\]/.test(document.title)) || 'cornice o titolo assenti';
    });
    await prova(S, 'il database si dichiara «collaudo» e nessun blocco d\'ambiente è scattato', async function () {
      var r = await _q(_sb.from('impostazioni').select('valore').eq('chiave', 'AMBIENTE').maybeSingle());
      if (!r || r.valore !== 'collaudo') return 'riga AMBIENTE = ' + JSON.stringify(r);
      return !document.getElementById('bloccoAmbiente') || 'è comparso il blocco d\'ambiente';
    });
    await prova(S, 'nessuna richiesta di rete verso il progetto di produzione', function () {
      var verso = performance.getEntriesByType('resource').map(function (e) { return e.name; }).filter(function (u) { return u.indexOf(REF_PRODUZIONE) >= 0; });
      return !verso.length || ('richieste verso la produzione: ' + verso.slice(0, 3).join(' , '));
    });
    await prova(S, 'reparto finto al completo: 28 schede, 24 pazienti', function () {
      var tutte = document.querySelectorAll('.patient-card').length;
      var piene = [].filter.call(document.querySelectorAll('.patient-card [data-field="Nome"]'), function (n) { return (n.innerText || '').trim(); }).length;
      return (tutte === 28 && piene === 24) || ('schede ' + tutte + ', con paziente ' + piene);
    });
    await prova(S, 'nessun errore JavaScript in pagina', function () {
      return !(window.__erroriBanco || []).length || window.__erroriBanco.join(' | ');
    });
  }

  // ═════════════════════════════════════════════════════════════════════
  // FINTO TRAKCARE: i segnalibri veri leggono ciò che il catalogo descrive,
  // e reimportarlo nelle schede appena generate non cambia nulla
  // ═════════════════════════════════════════════════════════════════════
  async function sezioneTrak(opz) {
    var S = 'importazione da TrakCare';
    var piani = C.pianifica();
    await prova(S, 'i segnalibri non contengono sequenze %XX (il browser le decodificherebbe)', function () {
      var m = (_BM_TERAPIA + _BM_LAB).match(/%[0-9a-fA-F]{2}/);
      return !m || ('trovata la sequenza ' + m[0]);
    });

    var f = await cornice('/trak-finto/index.html?cb=' + Date.now());
    var win = f.contentWindow, doc = win.document;
    function mostra(indice, vista, giorno) {
      var sp = doc.getElementById('scegliPaziente'), sg = doc.getElementById('scegliGiorno');
      sp.value = String(indice); sp.dispatchEvent(new win.Event('change'));
      sg.value = String(giorno); sg.dispatchEvent(new win.Event('change'));
      doc.querySelector('nav button[data-vista="' + vista + '"]').click();
    }
    async function terapiaLetta(indice, giorno) {
      mostra(indice, 'terapia', giorno);
      return leggiPayload(await premi(win, _BM_TERAPIA, 5000), '⟦TERAPIA-TRAK⟧');
    }

    // ── terapia, giorno delle schede: tutti i pazienti ──
    var lette = [], ko = { payload: [], righe: [], nome: [], reimporto: [] };
    for (var i = 0; i < piani.length; i++) {
      var p = piani[i], o = await terapiaLetta(i, 0);
      lette.push(o);
      if (!o || o.v !== _BM_VERSIONE || !o.sezioni) { ko.payload.push(p.paziente.Letto); continue; }
      if (!uguali(o.sezioni, [{ t: 'Pannello terapia', righe: C.terapiaDi(p, 0) }])) ko.righe.push(p.paziente.Letto + ': ' + primaDifferenza(o.sezioni, [{ t: 'Pannello terapia', righe: C.terapiaDi(p, 0) }]));
      var card = scheda(p.paziente.Letto), nomeScheda = card ? (card.querySelector('[data-field="Nome"]').innerText || '').trim() : '';
      if (o.paz !== p.nome || o.cs !== p.paziente.CodiceSanitario || !_nomiCombaciano(o.paz, nomeScheda)) ko.nome.push(p.paziente.Letto + ' («' + o.paz + '» / scheda «' + nomeScheda + '»)');
      // reimporto su una copia del campo: il contenuto (timbro a parte) non deve cambiare
      try {
        var finta = document.createElement('div'); finta.className = 'patient-card';
        var campo = card.querySelector('[data-field="NoteTerapia"]').cloneNode(true);
        finta.appendChild(campo);
        var senzaTimbro = function (el) { var c = el.cloneNode(true), h = c.querySelector('.terapia-box-head'); if (h) h.remove(); return c.innerHTML; };
        var prima = senzaTimbro(campo), box = campo.querySelector('.terapia-box');
        var r = _parseTerapiaTrak(o.sezioni);
        _terapiaInserisci(finta, r.corso, o.paz, Number(box && box.getAttribute('data-ts')) || Date.now(), r.pregressi);
        if (senzaTimbro(campo) !== prima) ko.reimporto.push(p.paziente.Letto + ': ' + primaDifferenza(senzaTimbro(campo), prima));
      } catch (e) { ko.reimporto.push(p.paziente.Letto + ': eccezione ' + e.message); }
    }
    await prova(S, 'terapia: il segnalibro copia un payload valido per tutti i 24 pazienti', function () { return !ko.payload.length || ('letti senza payload: ' + ko.payload.join(', ')); });
    await prova(S, 'terapia: le righe lette dalla pagina sono quelle del catalogo', function () { return !ko.righe.length || ko.righe.slice(0, 2).join(' ‖ '); });
    await prova(S, 'terapia: nome e codice del paziente letti dall\'intestazione combaciano con la scheda', function () { return !ko.nome.length || ko.nome.join(', '); });
    await prova(S, 'terapia: reimportare il giorno delle schede non cambia il campo', function () { return !ko.reimporto.length || ko.reimporto.slice(0, 2).join(' ‖ '); });
    await prova(S, 'terapia: le righe da scartare (ordine di reparto, errore di inserimento) non entrano', function () {
      var dentro = [];
      piani.forEach(function (p, i) {
        if (!lette[i]) return;
        var r = _parseTerapiaTrak(lette[i].sezioni), tutto = (r.corso + ' ' + JSON.stringify(r.pregressi)).toLowerCase();
        if (/contramal|tramadolo|monitoraggio diuresi|parafarmaco/.test(tutto)) dentro.push(p.paziente.Letto);
      });
      var presentiNelTrak = piani.filter(function (p, i) { return lette[i] && /PARAFARMACO|Errore inserimento/.test(JSON.stringify(lette[i].sezioni)); }).length;
      if (!presentiNelTrak) return 'il finto TrakCare non contiene righe da scartare: la prova non proverebbe nulla';
      return !dentro.length || ('importate nei letti ' + dentro.join(', '));
    });

    // ── terapia, giorni simulati ──
    var koGiorni = [];
    for (var g = 1; g <= 2; g++) {
      for (var j = 0; j < piani.length; j++) {
        var og = await terapiaLetta(j, g);
        if (!og || !uguali(og.sezioni, [{ t: 'Pannello terapia', righe: C.terapiaDi(piani[j], g) }])) koGiorni.push(piani[j].paziente.Letto + '+' + g);
      }
    }
    await prova(S, 'terapia: anche nei giorni simulati (+1, +2) la pagina e il catalogo coincidono', function () { return !koGiorni.length || koGiorni.join(', '); });
    await prova(S, 'terapia: il giorno dopo compaiono una sospensione e una dose immediata; due giorni dopo un farmaco nuovo', async function () {
      var r1 = _parseTerapiaTrak((await terapiaLetta(0, 1)).sezioni), r2 = _parseTerapiaTrak((await terapiaLetta(0, 2)).sezioni);
      var domani = C.dmy(C.giorno(1)).slice(0, 5), dopodomani = C.dmy(C.giorno(2)).slice(0, 5);
      var sosp = r1.pregressi.filter(function (x) { return /isolyte/i.test(x.nome); })[0], imm = r1.pregressi.filter(function (x) { return /plasil/i.test(x.nome); })[0];
      if (/isolyte/i.test(r1.corso)) return 'Isolyte è ancora nella terapia in corso';
      if (!sosp || sosp.date.join(' ').indexOf('sospeso il ' + domani) < 0) return 'sospensione non riconosciuta: ' + JSON.stringify(sosp);
      if (!imm || imm.date.join(' ').indexOf(domani) < 0) return 'dose immediata non riconosciuta: ' + JSON.stringify(imm);
      return new RegExp('Laevolac[^\\n]*dal ' + dopodomani).test(r2.corso) || ('farmaco nuovo assente: ' + r2.corso.slice(0, 300));
    });

    // ── terapia: esiti attesi del parser (fotografia da confrontare) ──
    var foto = {};
    for (var gg = 0; gg <= 2; gg++) piani.forEach(function (p) { foto[p.paziente.Letto + '+' + gg] = _parseTerapiaTrak([{ t: 'Pannello terapia', righe: C.terapiaDi(p, gg) }]); });
    // { fotografa: true } rifà la fotografia: va fatto SOLO dopo aver riletto le
    // differenze e deciso che il comportamento nuovo del parser è quello voluto.
    if (opz.fotografa) {
      var salvata = await fetch('/pagina/attesi-terapia.json', { method: 'PUT', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ epoca: C.EPOCA, esiti: foto }, null, 1) });
      await prova(S, 'fotografia degli esiti del parser salvata', function () { return salvata.ok || ('il banco ha risposto ' + salvata.status); });
    }
    await prova(S, 'terapia: il parser produce gli esiti attesi (fotografia in attesi-terapia.json)', async function () {
      var res = await fetch('/pagina/attesi-terapia.json?cb=' + Date.now());
      if (!res.ok) return 'manca attesi-terapia.json: va creata con __prove({ fotografa: true })';
      var attesi = await res.json();
      if (attesi.epoca !== C.EPOCA) return 'la fotografia è dell\'epoca ' + attesi.epoca + ', il catalogo è all\'epoca ' + C.EPOCA + ': va rifatta';
      var diversi = Object.keys(foto).filter(function (k) { return !uguali(foto[k], attesi.esiti[k]); });
      return !diversi.length || (diversi.length + ' esiti diversi (letto+giorno): ' + diversi.slice(0, 6).join(', ') + ' — ' + primaDifferenza(foto[diversi[0]], attesi.esiti[diversi[0]]));
    });

    // ── segnalibro sbagliato per la pagina ──
    await prova(S, 'il segnalibro degli esami premuto sulla Terapia avvisa e non copia', async function () {
      mostra(0, 'terapia', 0);
      var e = await premi(win, _BM_LAB, 8000);
      return (!e.copiato && /Non sei nella pagina «Laboratorio»/.test(e.avviso || '') && /Terapia/.test(e.avviso || '')) || JSON.stringify(e).slice(0, 200);
    });
    await prova(S, 'il segnalibro della terapia premuto sul Laboratorio avvisa e non copia', async function () {
      mostra(0, 'laboratorio', 0);
      var e = await premi(win, _BM_TERAPIA, 5000);
      return (!e.copiato && /Non trovo la terapia/.test(e.avviso || '')) || JSON.stringify(e).slice(0, 200);
    });
    f.remove();

    // ── esami: ogni paziente in una cornice propria (il segnalibro sfoglia le pagine) ──
    async function esamiLetti(p, giorno) {
      var fr = await cornice('/trak-finto/index.html?cb=' + Date.now() + Math.random() + '#letto=' + encodeURIComponent(p.paziente.Letto) + '&vista=laboratorio&giorno=' + giorno);
      try {
        await finche(function () { return fr.contentDocument.getElementById('tOEOrdItem_ListLabCummEMRList_0'); }, 8000);
        return await premi(fr.contentWindow, _BM_LAB, 120000);
      } finally { fr.remove(); }
    }
    async function aGruppi(elenco, quanti, fn) {
      var out = [];
      for (var k = 0; k < elenco.length; k += quanti) out.push.apply(out, await Promise.all(elenco.slice(k, k + quanti).map(fn)));
      return out;
    }
    var esitiLab = await aGruppi(piani, 12, function (p) { return esamiLetti(p, 0); });
    var koLab = { payload: [], pagine: [], esami: [], unione: [], assenti: [], vuoto: [] };
    for (var n = 0; n < piani.length; n++) {
      var pz = piani[n], el = esitiLab[n], letto = pz.paziente.Letto;
      if (pz.senzaEsami) { if (el.copiato || !/sembra vuota/.test(el.avviso || '')) koLab.vuoto.push(letto + ': ' + JSON.stringify(el).slice(0, 120)); continue; }
      var pl = leggiPayload(el, '⟦LAB-TRAK⟧');
      if (!pl || pl.v !== _LAB_BM_VERSIONE || pl.paz !== pz.nome) { koLab.payload.push(letto + ': ' + JSON.stringify(el).slice(0, 120)); continue; }
      var attese = C.labDi(pz, 0);
      if ((pl.pagine || []).length !== attese.length) koLab.pagine.push(letto + ' (' + (pl.pagine || []).length + ' invece di ' + attese.length + ')');
      var letti = esamiDi(pl.pagine);
      if (!uguali(letti, esamiDi(attese))) koLab.esami.push(letto + ': ' + primaDifferenza(letti, esamiDi(attese)));
      var riga = await _labCarica(letto);
      if (pz.labNellaScheda) {
        if (!riga || !riga.dati) { koLab.unione.push(letto + ': nel database non ci sono esami'); continue; }
        var m = _labMerge(copia(riga.dati), letti);
        if (m.nuovi !== 0 || m.valori !== 0 || !uguali(m.doc, riga.dati)) koLab.unione.push(letto + ': ' + m.nuovi + ' esami nuovi, ' + m.valori + ' valori aggiunti');
      } else if (riga) koLab.assenti.push(letto);
    }
    await prova(S, 'esami: il segnalibro sfoglia le pagine e copia un payload valido per i 23 pazienti con prelievi', function () { return !koLab.payload.length || koLab.payload.slice(0, 3).join(' ‖ '); });
    await prova(S, 'esami: tutte le pagine dello storico vengono lette', function () { return !koLab.pagine.length || koLab.pagine.join(', '); });
    await prova(S, 'esami: i valori letti dalla pagina sono quelli del catalogo', function () { return !koLab.esami.length || koLab.esami.slice(0, 2).join(' ‖ '); });
    await prova(S, 'esami: reimportare il giorno delle schede non aggiunge né cambia nulla', function () { return !koLab.unione.length || koLab.unione.join(' ‖ '); });
    await prova(S, 'esami: i pazienti «mai importati» non hanno esami nel database', function () { return !koLab.assenti.length || ('hanno già esami: ' + koLab.assenti.join(', ')); });
    await prova(S, 'esami: sulla griglia vuota il segnalibro avvisa e non copia', function () { return !koLab.vuoto.length || koLab.vuoto.join(' ‖ '); });

    // ── esami, il giorno dopo: solo valori nuovi, il passato resta com'è ──
    var conEsami = piani.filter(function (p) { return p.labNellaScheda; }).slice(0, 6);
    var domaniLab = await aGruppi(conEsami, 6, function (p) { return esamiLetti(p, 1); });
    var koDomani = [];
    for (var q = 0; q < conEsami.length; q++) {
      var pd = conEsami[q], pld = leggiPayload(domaniLab[q], '⟦LAB-TRAK⟧'), rd = await _labCarica(pd.paziente.Letto);
      if (!pld || !rd) { koDomani.push(pd.paziente.Letto + ': lettura mancata'); continue; }
      var md = _labMerge(copia(rd.dati), esamiDi(pld.pagine)), cambiati = 0, nuovaData = C.dmy(C.giorno(1)) + ' 06:00', conNuovaData = 0;
      Object.keys(rd.dati.esami).forEach(function (k) { Object.keys(rd.dati.esami[k].v).forEach(function (d) { if (md.doc.esami[k].v[d] !== rd.dati.esami[k].v[d]) cambiati++; }); });
      Object.keys(md.doc.esami).forEach(function (k) { if (md.doc.esami[k].v[nuovaData] != null) conNuovaData++; });
      if (cambiati || !md.valori || md.valori !== conNuovaData) koDomani.push(pd.paziente.Letto + ': ' + cambiati + ' valori passati cambiati, ' + md.valori + ' aggiunti, ' + conNuovaData + ' con la data nuova');
    }
    await prova(S, 'esami: il giorno dopo si aggiungono solo i valori del prelievo nuovo', function () { return !koDomani.length || koDomani.join(' ‖ '); });
    return { lette: lette, esitiLab: esitiLab };
  }

  // ═════════════════════════════════════════════════════════════════════
  // XSS: ciò che arriva da fuori (indirizzo della pagina, righe del database,
  // backup, memoria del browser) non deve MAI diventare codice o markup attivo,
  // e i contenuti leciti non devono cambiare di una virgola.
  //
  // Ogni contenuto ostile porta un <img src=x onerror=…>: se viene eseguito
  // lascia una traccia in window.__xss, se finisce in pagina come elemento lo
  // trova esche(). La sezione scrive sul letto libero «5», crea e cancella
  // righe di prova (un letto, una tipologia, due backup, due link) e alla fine
  // rimette tutto com'era.
  // ═════════════════════════════════════════════════════════════════════
  var LETTO_PROVA = '5';
  function traccia(id) { return "(window.__xss=window.__xss||[]).push('" + id + "')"; }
  // dentro una cornice figlia la traccia va lasciata alla pagina che la contiene
  function tracciaDaCornice(id) { return "(parent.__xss=parent.__xss||[]).push('" + id + "')"; }
  function eseguiti(win) { try { return ((win || window).__xss || []).slice(); } catch (e) { return ['finestra non leggibile']; } }
  function esca(id) { return '<img src=x onerror="' + traccia(id) + '">'; }
  // elementi che un contenuto ostile lascerebbe in pagina se venisse interpretato
  function esche(el) { return el ? el.querySelectorAll('img[src="x" i], [onerror], [onload], [onmouseover], [onfocus], [ontoggle], iframe, script').length : 0; }
  async function finestraSwal(ms) {
    return finche(function () { var p = document.querySelector('.swal2-popup'); return (p && Swal.isVisible()) ? p : false; }, ms || 8000);
  }
  async function chiudiSwal() { try { Swal.close(); } catch (e) {} await attendi(450); }
  // annulla il salvataggio ritardato che una prova ha fatto partire su un letto
  function annullaSalvataggio(letto) {
    try {
      clearTimeout(timerSalvataggioLetto[letto]); clearInterval(_countdownIntervallo[letto]);
      _dirtyLetti.delete(letto); _letti_salvataggioAttivi.delete(letto); _mostraMatite();
      _getBadges(letto).forEach(function (b) { b.classList.remove('visible'); b.style.opacity = '0'; });
    } catch (e) {}
  }
  // controlla una finestra Swal aperta: nessun elemento iniettato, nessuna esecuzione
  async function controllaSwal(dove, testoAtteso) {
    var p = await finestraSwal();
    if (!p) return dove + ': la finestra non si è aperta';
    await attendi(350);
    var guai = [];
    if (esche(p)) guai.push(dove + ': contenuto interpretato come HTML');
    if (testoAtteso && p.textContent.indexOf(testoAtteso) < 0) guai.push(dove + ': manca il testo atteso «' + testoAtteso + '»');
    if (eseguiti().length) guai.push(dove + ': codice ESEGUITO ' + eseguiti().join(', '));
    return guai.length ? guai.join(' ‖ ') : true;
  }

  // Vettori per il filtro: [nome, HTML ostile]. Coprono gestori di evento,
  // elementi attivi, indirizzi javascript:, fogli di stile, cambi di spazio dei
  // nomi (svg/math) e i casi noti di «mutazione» fra un'analisi e la successiva.
  function vettori() {
    var t = traccia;
    return [
      ['img onerror', '<img src=x onerror="' + t('v-img') + '">'],
      ['img senza virgolette', '<img src=x onerror=' + t('v-img2') + '>'],
      ['svg onload', '<svg onload="' + t('v-svg') + '"></svg>'],
      ['svg con script', '<svg><script>' + t('v-svgscript') + '</' + 'script></svg>'],
      ['svg animate', '<svg><animate onbegin="' + t('v-animate') + '" attributeName=x dur=1s></animate></svg>'],
      ['svg set', '<svg><set attributeName="onmouseover" to="' + t('v-set') + '"></set></svg>'],
      ['svg foreignObject', '<svg><foreignObject><iframe srcdoc="&lt;script&gt;' + tracciaDaCornice('v-fo') + '&lt;/script&gt;"></iframe></foreignObject></svg>'],
      ['svg use', '<svg><use href="data:image/svg+xml,&lt;svg id=x xmlns=http://www.w3.org/2000/svg&gt;&lt;script&gt;' + t('v-use') + '&lt;/script&gt;&lt;/svg&gt;#x"></use></svg>'],
      ['svg a', '<svg><a xlink:href="javascript:' + t('v-svga') + '"><circle r="40"></circle></a></svg>'],
      ['svg riempimento con indirizzo', '<svg><circle r="9" fill="url(https://example.invalid/x.svg#p)" stroke="url(javascript:' + t('v-fill') + ')"></circle></svg>'],
      ['svg style', '<svg><style><img src=x onerror="' + t('v-svgstyle') + '"></style></svg>'],
      ['svg title', '<svg><title><img src=x onerror="' + t('v-svgtitle') + '"></title></svg>'],
      ['script', '<script>' + t('v-script') + '</' + 'script>'],
      ['iframe srcdoc', '<iframe srcdoc="<script>' + tracciaDaCornice('v-iframe') + '</' + 'script>"></iframe>'],
      ['iframe javascript', '<iframe src="javascript:' + tracciaDaCornice('v-iframe2') + '"></iframe>'],
      ['link javascript', '<a href="javascript:' + t('v-a') + '">apri</a>'],
      ['link javascript mascherato', '<a href="jav&#x09;ascript:' + t('v-a2') + '">apri</a>'],
      ['gestore su tag lecito', '<b onmouseover="' + t('v-b') + '" onclick="' + t('v-b2') + '">testo</b>'],
      ['gestore su div', '<div onpointerenter="' + t('v-div') + '" onfocus="' + t('v-div2') + '" tabindex=0>testo</div>'],
      ['gestore su bottone lecito', '<button type="button" class="pa-del" onclick="' + t('v-btn') + '" formaction="javascript:' + t('v-btn2') + '">x</button>'],
      ['details ontoggle', '<details open ontoggle="' + t('v-details') + '">x</details>'],
      ['input autofocus', '<input autofocus onfocus="' + t('v-input') + '">'],
      ['video source', '<video><source onerror="' + t('v-video') + '"></video>'],
      ['audio onerror', '<audio src=x onerror="' + t('v-audio') + '"></audio>'],
      ['body onload', '<body onload="' + t('v-body') + '">'],
      ['object', '<object data="javascript:' + t('v-object') + '"></object>'],
      ['embed', '<embed src="javascript:' + t('v-embed') + '">'],
      ['form formaction', '<form><button formaction="javascript:' + t('v-form') + '">invia</button></form>'],
      ['meta refresh', '<meta http-equiv="refresh" content="0;url=javascript:' + t('v-meta') + '">'],
      ['base href', '<base href="javascript:' + t('v-base') + '//">'],
      ['style con import', '<style>@import "javascript:' + t('v-style') + '";body{background:red}</style>'],
      ['link stylesheet', '<link rel="stylesheet" href="https://example.invalid/x.css">'],
      ['stile con url', '<div style="background:url(javascript:' + t('v-css') + ')">x</div>'],
      ['stile con expression', '<div style="width:expression(' + t('v-expr') + ')">x</div>'],
      ['stile con colore e indirizzo', '<span style="color:red;background-color:url(https://example.invalid/x);font-size:1em">x</span>'],
      ['strato a tutta pagina', '<div style="position:fixed;top:0;left:0;width:9px;height:9px;z-index:99999;background:red" class="position-fixed top-0 start-0 w-100 h-100 modal fade show">sopra</div>'],
      ['campo finto', '<div class="editable-area rich-text patient-card focus-mode" data-field="Allergie" data-bed="1" contenteditable="true">NESSUNA ALLERGIA</div>'],
      ['math mtext', '<math><mtext><table><mglyph><style><!--</style><img title="--&gt;&lt;img src=x onerror=' + t('v-math') + '&gt;">'],
      ['math annotation', '<math><annotation-xml encoding="text/html"><img src=x onerror="' + t('v-math2') + '"></annotation-xml></math>'],
      ['noscript', '<noscript><p title="</noscript><img src=x onerror=' + t('v-noscript') + '>">'],
      ['template', '<template><img src=x onerror="' + t('v-template') + '"></template>'],
      ['textarea', '<textarea><img src=x onerror="' + t('v-textarea') + '"></textarea>'],
      ['title', '<title><img src=x onerror="' + t('v-title') + '"></title>'],
      ['xmp', '<xmp><img src=x onerror="' + t('v-xmp') + '"></xmp>'],
      ['commento', '<!--><img src=x onerror="' + t('v-comm') + '">-->'],
      ['select', '<select><option><img src=x onerror="' + t('v-select') + '"></option></select>'],
      ['tabella fuori posto', '<table><tr><td><img src=x onerror="' + t('v-table') + '"></td></tr><img src=x onerror="' + t('v-table2') + '"></table>'],
      ['bottone con dati Bootstrap', '<button data-bs-toggle="modal" data-bs-target="#modalRipristina" data-cmd="terapiatrak" id="loginOverlay" name="body">apri</button>'],
      ['marquee', '<marquee onstart="' + t('v-marquee') + '">x</marquee>'],
      ['isindex', '<isindex type=image src=1 onerror="' + t('v-isindex') + '">'],
      ['attributo con a capo', '<img src=x\nonerror\n=\n"' + t('v-nl') + '">'],
      ['maiuscole miste', '<ImG sRc=x OnErRoR="' + t('v-case') + '">'],
      ['entità nel nome', '<img src=x on&#x65;rror="' + t('v-ent') + '">'],
      ['attributo di testo con codice', '<div class="isolamento-banner" data-isolamento="&quot;&gt;&lt;img src=x onerror=' + t('v-attr') + '&gt;" title="&quot; onmouseover=&quot;' + t('v-attr2') + '">x</div>']
    ];
  }
  // Cosa NON deve esserci nell'HTML ripulito, una volta messo in pagina.
  function ispeziona(radice) {
    var guai = [];
    [].forEach.call(radice.querySelectorAll('*'), function (n) {
      var tag = n.tagName.toLowerCase();
      if (/^(script|iframe|object|embed|img|a|style|link|meta|base|form|input|textarea|select|option|video|audio|source|details|math|foreignobject|use|animate|set|template|noscript|title|xmp|marquee|isindex|body|html|head)$/.test(tag)) guai.push('<' + tag + '>');
      [].forEach.call(n.attributes, function (a) {
        if (/^on/i.test(a.name)) guai.push(tag + ' ' + a.name);
        if (/^(href|src|action|formaction|xlink:href|srcdoc|data|id|name|tabindex|data-field|data-bed)$/i.test(a.name)) guai.push(tag + ' ' + a.name);
        if (/^data-(bs-|cmd)/i.test(a.name)) guai.push(tag + ' ' + a.name);
        if (a.name === 'style' && /url|expression|position|z-index|javascript/i.test(a.value)) guai.push(tag + ' style «' + a.value.slice(0, 40) + '»');
        if (a.name === 'class' && /position-fixed|w-100|modal|fade|show|top-0|editable-area|patient-card|focus-mode/.test(a.value)) guai.push(tag + ' class «' + a.value.slice(0, 40) + '»');
        if (/^(fill|stroke)$/.test(a.name) && /url|javascript/i.test(a.value)) guai.push(tag + ' ' + a.name + ' «' + a.value.slice(0, 40) + '»');
      });
    });
    return guai;
  }

  async function sezioneXss(opz) {
    var S = 'XSS';
    window.__xss = [];
    var erroriPrima = (window.__erroriBanco || []).length;
    var inizioSezione = Date.now();
    var haFiltro = typeof window._pulisciHtml === 'function';

    // ── 1. l'avviso letto dall'indirizzo (XSS riflesso, anche senza accesso) ──
    async function provaToast(indirizzo, controlla) {
      var fr = await cornice(indirizzo + '&cb=' + Date.now());
      try {
        var w = fr.contentWindow;
        var toast = await finche(function () { return w.document.querySelector('#toastContainer .toast'); }, 7000);
        await attendi(1200);   // tempo a un eventuale gestore di scattare
        if (eseguiti(w).length) return 'codice ESEGUITO nella pagina: ' + eseguiti(w).join(', ');
        if (!toast) return 'l\'avviso non è comparso';
        if (esche(toast) || toast.querySelector('.toast-body b, .toast-body a')) return 'il parametro è finito in pagina come HTML';
        if (w.location.search) return 'l\'indirizzo non è stato ripulito: ' + w.location.search;
        return controlla(toast);
      } finally { fr.remove(); }
    }
    // (La stessa prova a utente NON collegato si fa a mano, su una pagina a sé:
    //  vedi «Prova di fumo» nel README. Qui non si può: per scollegare la cornice
    //  bisognerebbe togliere la sessione anche a questa pagina, che è la stessa origine.)
    var ostileToast = '<img src=x onerror="' + traccia('toast') + '"><b>grassetto</b>';
    await prova(S, 'indirizzo ?toast=info&msg=<codice>: nulla viene eseguito e il messaggio compare come testo', function () {
      return provaToast('/?toast=info&msg=' + encodeURIComponent(ostileToast), function (t) {
        var corpo = t.querySelector('.toast-body span');
        return (corpo && corpo.textContent === ostileToast) || ('testo mostrato: «' + (corpo && corpo.textContent) + '»');
      });
    });
    await prova(S, 'un avviso lecito (?toast=success&msg=…) si vede come prima: titolo «Successo» e testo', function () {
      var msg = 'Letto 5 aggiunto con successo.';
      return provaToast('/?toast=success&msg=' + encodeURIComponent(msg), function (t) {
        var titolo = t.querySelector('.toast-body strong'), corpo = t.querySelector('.toast-body span');
        if (!titolo || titolo.textContent !== 'Successo') return 'titolo: «' + (titolo && titolo.textContent) + '»';
        if (!corpo || corpo.textContent !== msg) return 'testo: «' + (corpo && corpo.textContent) + '»';
        return t.classList.contains('bg-success') || 'manca il colore verde';
      });
    });
    await prova(S, 'il messaggio dell\'indirizzo viene accorciato (200 caratteri) e il tipo è solo «success» o «info»', function () {
      var lungo = new Array(51).join('0123456789');
      return provaToast('/?toast=' + encodeURIComponent('danger" onmouseover="x') + '&msg=' + lungo, function (t) {
        var corpo = t.querySelector('.toast-body span');
        if (!corpo || corpo.textContent.length !== 200) return 'lunghezza mostrata: ' + (corpo && corpo.textContent.length);
        return (!t.classList.contains('bg-danger') && t.querySelector('.toast-body strong').textContent === 'Informazione') || 'il tipo non previsto è stato accettato';
      });
    });
    await prova(S, 'showToast scrive testo, mai HTML', async function () {
      var prima = eseguiti().length;
      window.showToast('<img src=x onerror="' + traccia('toast-titolo') + '">', '<img src=x onerror="' + traccia('toast-msg') + '"><i>x</i>', 'info');
      await attendi(700);
      var c = document.getElementById('toastContainer'), iniettati = c ? (esche(c) + c.querySelectorAll('.toast-body i:not(.bi)').length) : 0;
      [].forEach.call(c ? c.querySelectorAll('.toast') : [], function (t) { t.remove(); });
      if (eseguiti().length > prima) return 'codice ESEGUITO: ' + eseguiti().slice(prima).join(', ');
      return !iniettati || ('titolo o messaggio interpretati come HTML (' + iniettati + ' elementi)');
    });

    // Fin qui l'app è stata avviata più volte (nelle cornici) col solo reparto
    // finto, cioè con contenuti leciti: il filtro non deve aver tolto nulla.
    await prova(S, 'avvii dell\'app con le sole schede lecite: il filtro non segnala di aver tolto nulla', async function () {
      await attendi(1500);   // le segnalazioni partono senza attendere risposta
      var righe = await _q(_sb.from('logs').select('messaggio').eq('tipo', 'html-ripulito').gte('ts', inizioSezione));
      return !righe.length || (righe.length + ' segnalazioni: ' + righe.slice(0, 3).map(function (r) { return r.messaggio; }).join(' ‖ '));
    });

    // ── 2. gli attrezzi: codifica, selettori, colori, indirizzi, nomi dei letti ──
    await prova(S, 'il filtro dell\'HTML esiste ed è pronto', function () {
      if (!haFiltro) return 'window._pulisciHtml non esiste: i campi vengono scritti in pagina così come sono';
      return (typeof window._sanificaPronta === 'function' && window._sanificaPronta() && _filtroPronto()) || 'la libreria del filtro non si è caricata';
    });
    if (!haFiltro) return;
    await prova(S, 'codifica del testo, selettori, colori, indirizzi e nomi dei letti', function () {
      var k = [];
      function atteso(nome, ottenuto, voluto) { if (ottenuto !== voluto) k.push(nome + ': «' + ottenuto + '» invece di «' + voluto + '»'); }
      atteso('_testoHtml', _testoHtml('<a href="x" title=\'y\'>&'), '&lt;a href=&quot;x&quot; title=&#39;y&#39;&gt;&amp;');
      atteso('_testoHtml(0)', _testoHtml(0), '0');
      atteso('_testoHtml(null)', _testoHtml(null), '');
      atteso('_aEsc', _aEsc('a\'b"c<d>&'), 'a&#39;b&quot;c&lt;d&gt;&amp;');
      ['5', 'APPOGGIO 5M', 'a"b', 'a\'b', 'a\\b', 'x"] , body [y="', '1 2', '__proto__'].forEach(function (v) {
        var d = document.createElement('div'); d.setAttribute('data-bed', v); d.className = 'prova-sel'; document.body.appendChild(d);
        var t; try { t = document.querySelector('.prova-sel[data-bed="' + _cssVal(v) + '"]'); } catch (e) { t = 'eccezione'; }
        d.remove();
        if (t !== d) k.push('_cssVal non ritrova «' + v + '»');
      });
      [['#c62828', true], ['#FFF', true], ['rgb(1, 2, 3)', true], ['hsl(210,45%,40%)', true], ['red', true], ['red;position:fixed', false],
       ['#fff" onmouseover="x', false], ['url(x)', false], ['expression(1)', false], ['', false]].forEach(function (c) {
        var r = _coloreSicuro(c[0], 'RIPIEGO');
        if ((r !== 'RIPIEGO') !== c[1]) k.push('_coloreSicuro(«' + c[0] + '») = ' + r);
      });
      [['https://example.org/a', true], ['http://example.org', true], ['mailto:a@example.org', true], ['javascript:alert(1)', false],
       ['JaVaScRiPt:alert(1)', false], [' javascript:alert(1)', false], ['java\tscript:alert(1)', false],
       ['data:text/html,<b>x</b>', false], ['vbscript:x', false], ['', false]].forEach(function (c) {
        if (!!_urlSicuro(c[0]) !== c[1]) k.push('_urlSicuro(«' + c[0] + '») = «' + _urlSicuro(c[0]) + '»');
      });
      [['5', true], ['APPOGGIO 5M', true], ['APPOGGIO AL 9E', true], ['NOTE', true], ['12/A', true], ['5"', false], ['5\'', false], ['<B>', false],
       ['__proto__', false], ['constructor', false], ['', false], [' 5', false], [new Array(32).join('A'), false]].forEach(function (c) {
        if (_lettoValido(c[0]) !== c[1]) k.push('_lettoValido(«' + c[0] + '») = ' + _lettoValido(c[0]));
      });
      [['__proto__', true], ['constructor', true], ['toString', true], ['', true], ['5', false], ['NOTE', false], ['APPOGGIO 5M', false]].forEach(function (c) {
        if (_lettoRiservato(c[0]) !== c[1]) k.push('_lettoRiservato(«' + c[0] + '») = ' + _lettoRiservato(c[0]));
      });
      return !k.length || k.join(' ‖ ');
    });

    // ── 3. il filtro, vettore per vettore ──
    var banco = document.createElement('div');
    banco.style.cssText = 'position:fixed;left:-9000px;top:0;width:600px';
    document.body.appendChild(banco);
    var koVettori = [], prima2 = eseguiti().length;
    vettori().forEach(function (v) {
      var d = document.createElement('div');
      d.innerHTML = window._pulisciHtml(v[1]);
      banco.appendChild(d);
      var g = ispeziona(d);
      if (g.length) koVettori.push(v[0] + ' → ' + g.join(', '));
      var fisso = [].filter.call(d.querySelectorAll('*'), function (n) { return /fixed|absolute|sticky/.test(getComputedStyle(n).position); });
      if (fisso.length) koVettori.push(v[0] + ' → elemento fuori dal flusso (' + getComputedStyle(fisso[0]).position + ')');
    });
    await attendi(1200);   // tempo agli eventuali gestori (onerror, onload) di scattare
    var scattati = eseguiti().slice(prima2);
    banco.remove();
    await prova(S, vettori().length + ' vettori d\'attacco: dopo il filtro non resta nulla di attivo', function () { return !koVettori.length || koVettori.slice(0, 4).join(' ‖ '); });
    await prova(S, vettori().length + ' vettori d\'attacco: messi in pagina, nessuno viene eseguito', function () { return !scattati.length || ('ESEGUITI: ' + scattati.join(', ')); });
    await prova(S, 'il testo dei contenuti ostili resta leggibile (si tolgono i tag, non le parole)', function () {
      var p = window._pulisciHtml('<b onmouseover="x()">grassetto</b> e <a href="javascript:x()">collegamento</a> e <span class="position-fixed lab-box">riquadro</span>');
      var d = document.createElement('div'); d.innerHTML = p;
      return (d.textContent === 'grassetto e collegamento e riquadro' && !!d.querySelector('b') && d.querySelector('span').className === 'lab-box') || ('ottenuto: ' + p);
    });
    await prova(S, 'ripulire due volte dà lo stesso risultato', function () {
      var diversi = vettori().filter(function (v) { var a = window._pulisciHtml(v[1]); return window._pulisciHtml(a) !== a; });
      return !diversi.length || ('cambia alla seconda passata: ' + diversi.map(function (v) { return v[0]; }).join(', '));
    });
    await prova(S, 'testi delle scale (scritti da chi gestisce l\'app): restano elenchi e classi, sparisce ciò che è attivo', function () {
      var lecito = '<div class="risk-alto"><b>Rischio alto</b><ul><li>neurochirurgia</li></ul></div>';
      if (_pulisciHtmlConfig(lecito) !== lecito) return 'il testo lecito cambia: ' + _pulisciHtmlConfig(lecito);
      var d = document.createElement('div');
      d.innerHTML = _pulisciHtmlConfig(lecito + esca('scale-config') + '<script>' + traccia('scale-script') + '</' + 'script><a href="javascript:' + traccia('scale-a') + '">x</a>');
      return (!esche(d) && !d.querySelector('a, script') && !!d.querySelector('.risk-alto li')) || ('ottenuto: ' + d.innerHTML.slice(0, 200));
    });

    // ── 4. i contenuti leciti restano identici ──
    await prova(S, 'le schede del reparto finto passano dal filtro senza cambiare di una virgola', async function () {
      var righe = await _q(_sb.from('consegne').select('letto,note_terapia,diaria,da_fare,piano_terapeutico,esami_colturali,allergie'));
      var cambiati = [], n = 0;
      (righe || []).forEach(function (r) {
        if (r.letto === LETTO_PROVA) return;
        ['note_terapia', 'diaria', 'da_fare', 'piano_terapeutico', 'esami_colturali', 'allergie'].forEach(function (c) {
          var v = r[c] || ''; if (!v) return; n++;
          var p = window._pulisciHtml(v);
          if (p !== v) cambiati.push(r.letto + '.' + c + ': ' + primaDifferenza(p, v));
        });
      });
      if (!n) return 'nessun campo da controllare: il reparto finto è vuoto';
      return !cambiati.length || (cambiati.length + ' campi su ' + n + ' cambiano — ' + cambiati.slice(0, 2).join(' ‖ '));
    });
    await prova(S, 'le schede in pagina mostrano esattamente ciò che è nel database (campi ricchi e di solo testo)', async function () {
      var righe = await _q(_sb.from('consegne').select('*'));
      var diversi = [];
      (righe || []).forEach(function (r) {
        if (r.letto === LETTO_PROVA || r.letto === 'NOTE') return;
        var c = scheda(r.letto); if (!c) { diversi.push(r.letto + ': scheda assente'); return; }
        [['Nome', 'nome'], ['Diagnosi', 'diagnosi'], ['CodiceSanitario', 'codice_sanitario'], ['Ossigeno', 'ossigeno'], ['Vitto', 'vitto']].forEach(function (f) {
          var el = c.querySelector('[data-field="' + f[0] + '"]');
          var atteso = String(r[f[1]] || '').replace(/\s+/g, ' ').trim(), visto = (el.textContent || '').replace(/\s+/g, ' ').trim();
          if (el.children.length || visto !== atteso) diversi.push(r.letto + '.' + f[0]);
        });
        // nei campi ricchi il testo del database deve esserci tutto (i blocchi che l'app aggiunge a video non contano)
        [['Diaria', 'diaria'], ['NoteTerapia', 'note_terapia'], ['DaFare', 'da_fare'], ['Allergie', 'allergie']].forEach(function (f) {
          var el = c.querySelector('[data-field="' + f[0] + '"]');
          var d = document.createElement('div'); d.innerHTML = window._pulisciHtml(r[f[1]] || '');
          var atteso = (d.textContent || '').replace(/\s+/g, ' ').trim(), visto = (el.textContent || '').replace(/\s+/g, ' ').trim();
          if (visto !== atteso) diversi.push(r.letto + '.' + f[0]);
        });
      });
      return !diversi.length || ('diversi dal database: ' + diversi.slice(0, 8).join(', '));
    });
    await prova(S, 'ciò che l\'app e la barra degli strumenti producono passa dal filtro identico', function () {
      var campioni = [];
      function serializza(html) { var d = document.createElement('div'); d.innerHTML = html; return d.innerHTML; }
      // scritti a mano: modelli dei campi, spunte, tabelle, caratteri speciali
      [_TEMPLATE_DAFARE, _TEMPLATE_NOTETERAPIA, _TEMPLATE_NOTETERAPIA_VUOTO,
        '<div><span style="background-color: rgb(255, 235, 59);">evidenziato</span> <u>sottolineato</u> <i>corsivo</i> <strong>forte</strong></div>',
        '<div><span style="font-size: 18px;">grande</span> <font color="#c62828" face="Arial" size="3">colorato</font></div>',
        '<div class="terapia-box" data-ts="1791132126786"><div class="terapia-box-head" style="font-size:0.72rem;font-weight:700;color:#00796b;margin-bottom:3px;text-transform:uppercase;letter-spacing:0.03em;">Aggiornata da TrakCare il 04/10 18:02</div><div style="font-weight:700;color:#00695c;margin-top:5px;">EV</div><div style="padding-left:26px;text-indent:-22px;">- Ceftriaxone 2g ×1 (dal 28/09)</div><div class="pregressi-head">Farmaci pregressi</div><div data-preg="1" data-ord="3" style="padding-left:26px;text-indent:-22px;">- Meropenem (28/09–02/10)</div></div><div class="note-hdr" contenteditable="false">NOTE</div><div><br></div>',
        '<div class="cl-item"><strong>- <span class="cl-check" contenteditable="false" role="button" aria-checked="false">' + serializza('<svg width="16" height="16" viewBox="0 0 24 24" aria-hidden="true" style="vertical-align:-3px;"><rect x="3" y="3" width="18" height="18" rx="3" fill="#fff" stroke="#90a4ae" stroke-width="2"/></svg>') + '</span> 04/10:</strong>&nbsp;Emocolture<div class="cl-esito cl-esito-sbarrato"><strong><span contenteditable="false" class="cl-esito-arrow">↳&nbsp;</span>05/10:</strong>&nbsp;negative</div></div>',
        '<div>PA &lt; 90 e K &gt; 5, FE 35% &amp; BNP — «virgolette» e apostrofo d\'uso</div>',
        '<table><tbody><tr><td colspan="2">cella</td></tr></tbody></table><ul><li>voce</li></ul>'
      ].forEach(function (h) { campioni.push(['modello', serializza(h)]); });
      // prodotti dalle funzioni dell'app su una scheda di servizio: non è una
      // «patient-card» e non ha letto, quindi non fa partire alcun salvataggio
      var finta = document.createElement('div');
      finta.innerHTML = '<div class="alt-col-info"></div><div data-field="Diaria"></div><div data-field="EsamiColturali"></div><div data-field="PianoTerapeutico"></div>';
      finta.style.cssText = 'position:fixed;left:-9000px;top:0;width:600px';
      document.body.appendChild(finta);
      try {
        _isolamentoApplica(finta, 'da contatto: KPC "ospite"');
        _rianimazioneApplica(finta, 'DNR: sì, <concordato>');
        finta.querySelector('[data-field="Diaria"]').insertAdjacentHTML('beforeend', '<div><strong>-04/10:</strong>&nbsp;stabile</div>');
        campioni.push(['banner di isolamento e rianimazione', finta.querySelector('[data-field="Diaria"]').innerHTML]);
        _inserisciRisultatoScala(finta, 'has-bled', 'HAS-BLED del 04/10: 3 — rischio alto');
        var esami = finta.querySelector('[data-field="EsamiColturali"]');
        esami.insertAdjacentHTML('afterbegin', '<div class="lab-box" contenteditable="false">' + _labBottoneHtml() + '<div class="lab-remind lab-remind-off" title="Ultimo prelievo: 01/10/2026 — clicca per segnarlo come letto">⏰ Richiedere gli esami per domani: gli ultimi sono più vecchi di 3 giorni.</div></div>');
        campioni.push(['riquadro laboratorio e scale', esami.innerHTML]);
        var pa = finta.querySelector('[data-field="PianoTerapeutico"]');
        pa.appendChild(_paBarretta()); pa.appendChild(_paRiga('Polmonite <dx> & scompenso'));
        campioni.push(['problemi attivi', pa.innerHTML]);
        // barra degli strumenti: i comandi veri del browser su un testo selezionato
        var ed = document.createElement('div');
        ed.setAttribute('contenteditable', 'true');
        ed.innerHTML = '<div>prima riga da formattare</div><div>seconda riga</div>';
        finta.appendChild(ed);
        ed.focus();
        function seleziona(nodo, da, a) { var r = document.createRange(); r.setStart(nodo, da); r.setEnd(nodo, a); var s = getSelection(); s.removeAllRanges(); s.addRange(r); }
        seleziona(ed.firstChild.firstChild, 0, 5); document.execCommand('bold');
        seleziona(ed.lastChild.firstChild, 0, 7); document.execCommand('italic'); document.execCommand('underline');
        seleziona(ed.lastChild.lastChild, 1, 5);
        if (!document.execCommand('hiliteColor', false, '#ffff00')) document.execCommand('backColor', false, '#ffff00');
        getSelection().selectAllChildren(ed.firstChild);
        _tbSavedRange = null; _tbApplyFontMultiplier(1.5);
        campioni.push(['barra degli strumenti (grassetto, corsivo, sottolineato, evidenziatore, dimensione)', ed.innerHTML]);
        ed.blur();
      } finally { finta.remove(); }
      var cambiati = campioni.filter(function (c) { return window._pulisciHtml(c[1]) !== c[1]; });
      if (campioni.length !== 13) return 'campioni raccolti: ' + campioni.length + ' invece di 13';
      return !cambiati.length || ('cambiano — ' + cambiati.map(function (c) { return c[0] + ': ' + primaDifferenza(window._pulisciHtml(c[1]), c[1]); }).slice(0, 2).join(' ‖ '));
    });

    // ── 5. senza filtro l'app si ferma ──
    await prova(S, 'se il filtro non c\'è l\'app non disegna e non salva', async function () {
      var vero = window._sanificaPronta, k = [];
      window._sanificaPronta = function () { return false; };
      try {
        if (_filtroPronto()) k.push('_filtroPronto resta vero');
        if (_renderCardsHtml([{ Letto: '1', Nome: 'X' }]) !== '') k.push('_renderCardsHtml disegna comunque');
        var lanciato = false; try { _toDb({ Diaria: '<b>x</b>' }); } catch (e) { lanciato = true; }
        if (!lanciato) k.push('_toDb non rifiuta');
        var rifiutato = false; await _sbSalvaPaziente('LETTO-INESISTENTE', { Diaria: 'x' }).catch(function () { rifiutato = true; });
        if (!rifiutato) k.push('_sbSalvaPaziente non rifiuta');
        var c = scheda('1'), el = c.querySelector('[data-field="Diaria"]'), primaHtml = el.innerHTML;
        _aggiornaCardDaPaziente(c, { Diaria: '<b>cambiata</b>' });
        if (el.innerHTML !== primaHtml) { k.push('_aggiornaCardDaPaziente scrive comunque'); el.innerHTML = primaHtml; }
      } finally { window._sanificaPronta = vero; }
      return !k.length || k.join(' ‖ ');
    });
    await prova(S, 'se la libreria del filtro non si carica compare l\'avviso a tutto schermo e nessuna scheda', async function () {
      var fr = await cornice('/?senzaFiltro=1&cb=' + Date.now());
      try {
        var w = fr.contentWindow;
        var blocco = await finche(function () { return w.document.getElementById('bloccoFiltro'); }, 8000);
        if (!blocco) return 'l\'avviso non è comparso';
        await attendi(5000);   // tempo all'app di provare a disegnare
        var n = w.document.querySelectorAll('.patient-card').length;
        return !n || ('sono state disegnate ' + n + ' schede senza filtro');
      } finally { fr.remove(); }
    });

    // ── 6. contenuti ostili salvati in una scheda: ogni strada che li porta in pagina ──
    var card = scheda(LETTO_PROVA);
    var rigaOriginale = null;
    await prova(S, 'il letto di prova «' + LETTO_PROVA + '» esiste ed è libero', async function () {
      if (!card) return 'scheda del letto ' + LETTO_PROVA + ' non trovata';
      rigaOriginale = await _q(_sb.from('consegne').select('nome,diagnosi,tipologia_letto').eq('letto', LETTO_PROVA).maybeSingle());
      return (rigaOriginale && !(rigaOriginale.nome || '').trim() && !(rigaOriginale.diagnosi || '').trim()) || 'il letto non è libero: la prova non lo tocca';
    });
    var libero = esiti[esiti.length - 1].ok;
    if (!libero) return;

    var t = traccia;
    var TIPO_OSTILE = '"><img src=x onerror="' + t('c-tipologia') + '">';
    var ostile = {
      nome: 'ROSSI <img src=x onerror="' + t('c-nome') + '">MARIO',
      diagnosi: 'Polmonite <svg onload="' + t('c-diagnosi') + '"></svg> FE < 35%',
      eta: '70<img src=x onerror="' + t('c-eta') + '">',
      codice_sanitario: '123"><img src=x onerror="' + t('c-cs') + '">',
      ossigeno: 'AA<img src=x onerror="' + t('c-o2') + '">',
      vitto: 'Libero<img src=x onerror="' + t('c-vitto') + '">',
      sesso: 'M"><img src=x onerror="' + t('c-sesso') + '">',
      dimissibile: '05/10/2026"><img src=x onerror="' + t('c-dim') + '">',
      data_nascita: '"><img src=x onerror="' + t('c-nascita') + '">',
      data_ricovero: '"><img src=x onerror="' + t('c-ricovero') + '">',
      tipologia_letto: TIPO_OSTILE,
      allergie: 'Penicillina<img src=x onerror="' + t('c-allergie') + '">',
      diaria: '<div><b>04/10:</b> testo lecito</div><img src=x onerror="' + t('c-diaria-img') + '"><svg><g onload="' + t('c-diaria-svg') + '"></g></svg>'
        + '<a href="javascript:' + t('c-diaria-a') + '">collegamento</a><details open ontoggle="' + t('c-diaria-details') + '">x</details>'
        + '<div class="isolamento-banner" contenteditable="false" data-isolamento="X&quot;&gt;&lt;img src=x onerror=&quot;' + t('c-diaria-attr') + '&quot;&gt;" title="t">ISOLAMENTO</div>'
        + '<div style="position:fixed;top:0;left:0;width:9px;height:9px;z-index:99999;background:red" class="position-fixed">sopra</div>'
        + '<div class="editable-area" data-field="Allergie">CAMPO FINTO</div>',
      note_terapia: '<div class="terapia-box" data-ts="1"><div>- Farmaco</div></div><img src=x onerror="' + t('c-terapia') + '"><div class="note-hdr" contenteditable="false">NOTE</div><div><br></div>',
      da_fare: '<b>DA FARE:</b><br><img src=x onerror="' + t('c-dafare') + '"><input autofocus onfocus="' + t('c-dafare-input') + '">',
      piano_terapeutico: '<div class="pa-item" contenteditable="false"><span class="pa-txt">- Problema <img src=x onerror="' + t('c-pa') + '"></span></div>',
      esami_colturali: '<div class="scale-box" contenteditable="false"><div class="scala-risultato" contenteditable="false" data-scala="&lt;img src=x onerror=&quot;' + t('c-scala') + '&quot;&gt;">Scala finta: 3</div></div>'
        + '<div>Emocolture negative</div><iframe srcdoc="<script>' + tracciaDaCornice('c-esami-iframe') + '</' + 'script>"></iframe><img src=x onerror="' + t('c-esami') + '">',
      updated_at: new Date().toISOString()
    };
    var campiRicchi = ['Allergie', 'Diaria', 'NoteTerapia', 'DaFare', 'PianoTerapeutico', 'EsamiColturali'];
    var campiSemplici = { Nome: ostile.nome, Diagnosi: ostile.diagnosi, CodiceSanitario: ostile.codice_sanitario, Ossigeno: ostile.ossigeno, Vitto: ostile.vitto };
    function controllaScheda(c, dove) {
      var guai = [];
      campiRicchi.forEach(function (f) {
        var el = c.querySelector('[data-field="' + f + '"]'); if (!el) return;
        ispeziona(el).forEach(function (g) { guai.push(f + ': ' + g); });
        [].forEach.call(el.querySelectorAll('*'), function (n) { if (/fixed|absolute|sticky/.test(getComputedStyle(n).position) && !n.closest('.pa-az')) guai.push(f + ': elemento fuori dal flusso'); });
      });
      Object.keys(campiSemplici).forEach(function (f) {
        var el = c.querySelector('[data-field="' + f + '"]'); if (!el) return;
        if (el.children.length) guai.push(f + ': contiene elementi (' + el.children[0].tagName.toLowerCase() + ')');
        else if ((el.textContent || '') !== campiSemplici[f]) guai.push(f + ': il testo non è quello salvato');
      });
      if (esche(c)) guai.push('nella scheda ci sono ' + esche(c) + ' elementi iniettati');
      if (c.querySelectorAll('[data-field="Allergie"]').length !== 1) guai.push('campo Allergie duplicato da un campo finto');
      return guai.length ? (dove + ' — ' + guai.slice(0, 5).join(' ‖ ')) : true;
    }
    async function ridisegna() {
      _applicaAggiornamentoCompleto(_renderCardsHtml(await _sbGetPazienti()));
      await attendi(1000);
    }
    var daPulire = [];   // funzioni che rimettono a posto, eseguite comunque alla fine
    var scritto = false;
    try {
      window.__xss = [];
      await _q(_sb.from('consegne').update(ostile).eq('letto', LETTO_PROVA));
      scritto = true;

      await prova(S, 'aggiornamento in tempo reale di una scheda con contenuti ostili', async function () {
        var arrivato = await finche(function () { var n = scheda(LETTO_PROVA).querySelector('[data-field="Nome"]'); return n && /ROSSI/.test(n.textContent); }, 9000);
        if (!arrivato) return 'l\'aggiornamento non è arrivato alla pagina entro 9 secondi';
        await attendi(1000);
        if (eseguiti().length) return 'codice ESEGUITO: ' + eseguiti().join(', ');
        return controllaScheda(scheda(LETTO_PROVA), 'dopo il Realtime');
      });
      await prova(S, 'ridisegno completo delle schede (come all\'avvio e a ogni risincronizzazione)', async function () {
        window.__xss = [];
        await ridisegna();
        if (eseguiti().length) return 'codice ESEGUITO: ' + eseguiti().join(', ');
        var c = scheda(LETTO_PROVA), badge = c.querySelector('[id^="badge-tipo-"]');
        if (!badge || badge.textContent.trim() !== TIPO_OSTILE.toUpperCase()) return 'il nome della tipologia non è mostrato come testo: «' + (badge && badge.textContent) + '»';
        return controllaScheda(c, 'dopo il ridisegno');
      });
      await prova(S, 'apertura della scheda in modifica (rilegge il letto dal database)', async function () {
        window.__xss = [];
        var c = scheda(LETTO_PROVA);
        _attivaFocusMode(c);
        var aperta = await finche(function () { return c.querySelector('.focus-toolbar'); }, 12000);
        await attendi(900);
        var esito = eseguiti().length ? ('codice ESEGUITO: ' + eseguiti().join(', ')) : (aperta ? controllaScheda(c, 'in modifica') : 'la scheda non si è aperta in modifica');
        // clic sulla riga della scala salvata con un identificativo ostile: il titolo della finestra è testo
        if (esito === true && aperta) {
          var riga = c.querySelector('.scala-risultato[data-scala]');
          if (!riga) esito = 'la riga della scala non è arrivata in pagina';
          else {
            _apriModalRiepilogoScala(riga);
            esito = await controllaSwal('riepilogo della scala', '<img');
            await chiudiSwal();
          }
        }
        // si esce SENZA salvare: la scheda non è stata toccata
        try { if (typeof _dirtyLetti !== 'undefined') _dirtyLetti.delete(LETTO_PROVA); } catch (e) {}
        _disattivaFocusMode(false);
        await attendi(1200);
        return esito;
      });
      await prova(S, 'avvio dell\'app da zero con la scheda ostile già nel database', async function () {
        var fr = await cornice('/?cb=' + Date.now());
        try {
          var w = fr.contentWindow;
          var pronta = await finche(function () { return w.document.querySelector('.patient-card[data-bed="' + LETTO_PROVA + '"] [data-field="Nome"]'); }, 25000);
          if (!pronta) return 'l\'app nella cornice non ha disegnato le schede';
          await attendi(1500);
          if (eseguiti(w).length) return 'codice ESEGUITO: ' + eseguiti(w).join(', ');
          var c = w.document.querySelector('.patient-card[data-bed="' + LETTO_PROVA + '"]');
          var guai = [];
          campiRicchi.forEach(function (f) { var el = c.querySelector('[data-field="' + f + '"]'); if (el) ispeziona(el).forEach(function (g) { guai.push(f + ': ' + g); }); });
          Object.keys(campiSemplici).forEach(function (f) { var el = c.querySelector('[data-field="' + f + '"]'); if (el && el.children.length) guai.push(f + ': contiene elementi'); });
          var tutte = w.document.getElementById('cardsContainer');
          if (esche(tutte)) guai.push('elementi iniettati fra le schede: ' + esche(tutte));
          return !guai.length || guai.slice(0, 5).join(' ‖ ');
        } finally { fr.remove(); }
      });
      await prova(S, 'pagina di stampa (entrambe le impaginazioni) con la scheda ostile', async function () {
        var guai = [];
        for (var i = 0; i < 2; i++) {
          var fr = await cornice('/print.html?senzaStampa=1&layout=' + (i ? 'alt' : 'main') + '&cb=' + Date.now());
          try {
            var w = fr.contentWindow;
            var pronta = await finche(function () { return w.document.querySelector('#printCards .patient-card, #printCards .alt-row'); }, 20000);
            if (!pronta) { guai.push('impaginazione ' + (i ? 'alt' : 'main') + ': la stampa non ha disegnato nulla'); continue; }
            await attendi(1200);
            if (eseguiti(w).length) guai.push('impaginazione ' + (i ? 'alt' : 'main') + ': codice ESEGUITO ' + eseguiti(w).join(', '));
            var cattivi = w.document.querySelectorAll('#printCards img, #printCards iframe, #printCards a[href], #printCards input, #printCards details, #printCards [onerror], #printCards [onload], #printCards [onmouseover], #printCards [onfocus], #printCards [ontoggle]');
            if (cattivi.length) guai.push('impaginazione ' + (i ? 'alt' : 'main') + ': ' + cattivi.length + ' elementi attivi in pagina (' + cattivi[0].tagName.toLowerCase() + ')');
            if (w.document.getElementById('printCards').textContent.indexOf('ROSSI <img') < 0) guai.push('impaginazione ' + (i ? 'alt' : 'main') + ': il nome non compare come testo');
          } finally { fr.remove(); }
        }
        return !guai.length || guai.join(' ‖ ');
      });
      await prova(S, 'pagina di stampa: i parametri dell\'indirizzo non entrano nel foglio di stile né scelgono un\'impaginazione imprevista', async function () {
        var fr = await cornice('/print.html?senzaStampa=1&layout=' + encodeURIComponent('<b>x') + '&orientamento=' + encodeURIComponent('portrait;}body{background:url(https://example.invalid/x)}@page{') + '&cb=' + Date.now());
        try {
          await attendi(1500);
          var st = [].map.call(fr.contentDocument.querySelectorAll('style'), function (s) { return s.textContent; }).join(' ');
          if (/example\.invalid/.test(st)) return 'il parametro «orientamento» è finito nel foglio di stile';
          return /size: A4 portrait;/.test(st) || 'manca la regola @page attesa';
        } finally { fr.remove(); }
      });

      // ── 7. liste e finestre che citano letto, nome, tipologia ──
      await prova(S, 'menu «Lista Pazienti»', async function () {
        window.__xss = [];
        _popolaDropdownPazienti();
        await attendi(700);
        var menu = document.getElementById('menuListaPazienti'), guai = [];
        if (esche(menu)) guai.push('il menu contiene elementi iniettati');
        // (il menu rilegge il nome dalla scheda, dove il foglio di stile lo mostra in maiuscolo)
        if (menu.textContent.toUpperCase().indexOf('ROSSI <IMG') < 0) guai.push('il nome non compare come testo');
        if (eseguiti().length) guai.push('codice ESEGUITO: ' + eseguiti().join(', '));
        return !guai.length || guai.join(' ‖ ');
      });
      await prova(S, 'riepilogo dei letti (nome della tipologia)', async function () {
        window.__xss = [];
        _aggiornaPannelloRiepilogoLetti();
        var corpo = document.getElementById('lettiRiepilogoBody');
        var pronto = await finche(function () { return corpo.querySelector('.badge'); }, 8000);
        await attendi(600);
        if (!pronto) return 'il riepilogo non si è disegnato';
        if (esche(corpo)) return 'il riepilogo contiene elementi iniettati';
        if (corpo.textContent.toUpperCase().indexOf('<IMG SRC=X') < 0) return 'il nome della tipologia non compare come testo';
        return !eseguiti().length || ('codice ESEGUITO: ' + eseguiti().join(', '));
      });
      await prova(S, 'finestra delle opzioni di stampa (elenco delle tipologie)', async function () {
        window.__xss = [];
        _apriModalScala(function () {});
        var lista = document.getElementById('tipologieStampaLista');
        var pronto = await finche(function () { return lista.querySelector('.tipologia-check'); }, 8000);
        await attendi(500);
        var esito = !pronto ? 'l\'elenco non si è disegnato'
          : esche(lista) ? 'l\'elenco contiene elementi iniettati'
          : ![].some.call(lista.querySelectorAll('.tipologia-check'), function (c) { return c.value === TIPO_OSTILE.toUpperCase(); }) ? 'la tipologia non è fra le voci (come valore di testo)'
          : eseguiti().length ? ('codice ESEGUITO: ' + eseguiti().join(', ')) : true;
        var m = bootstrap.Modal.getInstance(document.getElementById('modalScalaStampa')); if (m) m.hide();
        await attendi(600);
        return esito;
      });
      await prova(S, 'finestra della check list odierna', async function () {
        window.__xss = [];
        _stampaCheckListOdierna();
        await finche(function () { return document.querySelector('.cl-letto-cb'); }, 9000);
        var esito = await controllaSwal('check list', 'ROSSI <img');
        await chiudiSwal();
        return esito;
      });
      await prova(S, 'finestra della mail dimissioni (mittente ed elenco dei pazienti) — senza inviare nulla', async function () {
        window.__xss = [];
        // I permessi di invio li «conferma» un server simulato: la prova non deve
        // dipendere dalla cassaforte del collaudo. Il mittente che dichiara è ostile.
        var mittenteOstile = 'ROSSI ' + esca('mittente');
        fingiStatoMail({ configurato: true, autorizzato: true, verificato: true, email: 'x@example.com', mittente: mittenteOstile });
        try {
          apriModalMailDimissioni();
          await finche(function () { return document.querySelector('.mail-dim-cb'); }, 9000);
        } finally { statoMailVero(); }
        var esito = await controllaSwal('mail dimissioni', 'ROSSI <img');
        var campo = document.getElementById('mailDimMittente');
        if (esito === true && !(campo && campo.value === mittenteOstile)) esito = 'mail dimissioni: il mittente non è mostrato come testo';
        await chiudiSwal();
        return esito;
      });
      await prova(S, 'conferma di spostamento su un letto occupato (nome e tipologia del paziente)', async function () {
        window.__xss = [];
        var so = document.getElementById('selectSpostaOrigine'), sd = document.getElementById('selectSpostaDestinazione');
        if (!so || !sd) return 'menu a tendina dello spostamento non trovati';
        function voce(sel, v) { var o = document.createElement('option'); o.value = v; o.textContent = v; o.setAttribute('data-prova', '1'); sel.appendChild(o); sel.value = v; }
        voce(so, '1'); voce(sd, LETTO_PROVA);
        eseguiSpostaPaziente();
        await finche(function () { var p = document.querySelector('.swal2-popup'); return p && Swal.isVisible() && /Letto Occupato/.test(p.textContent); }, 9000);
        var esito = await controllaSwal('spostamento', 'ROSSI <img');
        try { Swal.clickCancel(); } catch (e) {}
        await attendi(500);
        [].forEach.call(document.querySelectorAll('option[data-prova]'), function (o) { o.remove(); });
        return esito;
      });
      await prova(S, 'avvisi di scheda in uso (il nome del letto è testo)', async function () {
        window.__xss = [];
        var nomeLetto = esca('lock-letto');
        _mostraAvvisoLock(nomeLetto);
        var a = await controllaSwal('avviso di blocco', '<img');
        await chiudiSwal();
        if (a !== true) return a;
        _lockState[nomeLetto] = { token: 'altro', ts: Date.now() };
        var chiamata = false;
        try { _verificaLockEProcedi([nomeLetto], function () { chiamata = true; }); } finally { delete _lockState[nomeLetto]; }
        var b = await controllaSwal('letto in uso', '<img');
        await chiudiSwal();
        return chiamata ? 'l\'operazione è proseguita nonostante il blocco' : b;
      });

      // ── 8. salvataggio: ciò che parte verso il database è già ripulito ──
      await prova(S, 'salvataggio: l\'HTML ostile presente in pagina non arriva al database, e un campo finto non sovrascrive quello vero', async function () {
        window.__xss = [];
        var c = scheda(LETTO_PROVA), diaria = c.querySelector('[data-field="Diaria"]');
        // come se qualcuno avesse scritto nel DOM con gli strumenti del browser
        var img = document.createElement('img'); img.setAttribute('src', 'x'); img.setAttribute('onerror', 'void 0');
        var finto = document.createElement('div'); finto.className = 'editable-area'; finto.setAttribute('data-field', 'Allergie'); finto.textContent = 'ALLERGIE FALSE';
        diaria.appendChild(img); diaria.appendChild(finto);
        // l'HTML si giudica da ciò che DIVENTA (elementi e attributi), non dal testo:
        // un valore come «<img onerror=…>» dentro un attributo di testo è solo testo
        function guaiHtml(html) { var m = document.createElement('template'); m.innerHTML = html; return ispeziona(m.content); }
        var riga = _toDb({ Diaria: diaria.innerHTML, Nome: '<b>testo</b>' });
        if (guaiHtml(riga.diaria).length) return '_toDb lascia passare: ' + guaiHtml(riga.diaria).join(', ');
        if (riga.nome !== '<b>testo</b>') return 'un campo di solo testo è stato alterato: ' + riga.nome;
        eseguiSalvataggioLettoCompleto(LETTO_PROVA, c);
        await attendi(2500);
        var r = await _q(_sb.from('consegne').select('diaria,allergie,nome,note_terapia,da_fare,piano_terapeutico,esami_colturali').eq('letto', LETTO_PROVA).maybeSingle());
        img.remove(); finto.remove();
        annullaSalvataggio(LETTO_PROVA);
        var sporchi = ['diaria', 'allergie', 'note_terapia', 'da_fare', 'piano_terapeutico', 'esami_colturali'].filter(function (k) { return guaiHtml(r[k] || '').length; });
        if (sporchi.length) return 'nel database è arrivato HTML non ripulito in: ' + sporchi.join(', ') + ' (' + guaiHtml(r[sporchi[0]]).slice(0, 3).join(', ') + ')';
        if (r.diaria.indexOf('testo lecito') < 0) return 'il contenuto lecito della diaria non è stato salvato';
        if (/ALLERGIE FALSE/.test(r.allergie)) return 'il campo finto ha sovrascritto le allergie';
        // (i campi di solo testo si salvano come si vedono: il nome è mostrato in maiuscolo)
        return (r.nome.toUpperCase() === ostile.nome.toUpperCase()) || 'il nome (solo testo) è cambiato nel salvataggio: ' + r.nome;
      });

      // ── 9. istantanee locali (memoria del browser) ──
      await prova(S, 'recupero da un\'istantanea locale manomessa', async function () {
        window.__xss = [];
        var c = scheda(LETTO_PROVA), primaIst = window._snapshots[LETTO_PROVA];
        window._snapshots[LETTO_PROVA] = { current: null, storico: [{ tsOpen: Date.now() - 5000, tsUpdate: Date.now() - 4000, tsClose: Date.now() - 3000, tipo: esca('ist-tipo'),
          dati: { Diaria: esca('ist-diaria') + '<b>recuperata</b>', Nome: 'BIANCHI <b>x</b>' + esca('ist-nome'), PianoTerapeutico: '<div class="editable-area" data-field="Allergie">FINTO</div>',
                  'Diaria"],body [data-x="': 'selettore', Dimissibile: '06/10"><img src=x onerror="' + t('ist-dim') + '">' } }] };
        var esito;
        try {
          _snapshotApriModale(LETTO_PROVA);
          esito = await controllaSwal('anteprima istantanea', '<img');
          if (esito === true) {
            Swal.clickConfirm(); await attendi(700);
            esito = await controllaSwal('conferma ripristino', '<img');
            if (esito === true) {
              Swal.clickConfirm(); await attendi(900);
              var d = c.querySelector('[data-field="Diaria"]'), n = c.querySelector('[data-field="Nome"]');
              if (esche(c)) esito = 'dopo il ripristino la scheda contiene elementi iniettati';
              else if (!d.querySelector('b') || d.textContent.indexOf('recuperata') < 0) esito = 'il contenuto lecito dell\'istantanea non è stato ripristinato';
              else if (n.children.length || n.textContent.indexOf('BIANCHI <b>x</b>') !== 0) esito = 'il nome non è stato ripristinato come testo';
              else if (c.querySelectorAll('[data-field="Allergie"]').length !== 1) esito = 'campo finto entrato nella scheda';
              else if (eseguiti().length) esito = 'codice ESEGUITO: ' + eseguiti().join(', ');
            }
          }
        } finally {
          await chiudiSwal();
          annullaSalvataggio(LETTO_PROVA);   // niente salvataggio dell'istantanea di prova
          if (primaIst) window._snapshots[LETTO_PROVA] = primaIst; else delete window._snapshots[LETTO_PROVA];
          try { localStorage.setItem('_snapshots_v1', JSON.stringify(window._snapshots)); } catch (e) {}
        }
        return esito;
      });

      // ── 10. grafico degli esami: il fumetto ──
      await prova(S, 'grafico degli esami: nome e unità di misura ostili restano testo (tabella e fumetto)', async function () {
        window.__xss = [];
        var corpo = document.createElement('div');
        corpo.style.cssText = 'position:fixed;left:-9000px;top:0;width:900px';
        document.body.appendChild(corpo);
        try {
          var doc = { ord: ['k'], esami: { k: { g: 'EMOCROMO' + esca('lab-g'), n: 'Emoglobina' + esca('lab-n'), range: '12-16', um: 'g/dL' + esca('lab-um'),
            v: { '01/10/2026 08:00': '10.5', '02/10/2026 08:00': '11.2 L' } } } };
          var ui = { doc: doc, body: corpo, ov: null, avvisi: document.createElement('div'), multi: false, letto: LETTO_PROVA };
          _labRender(ui, doc);
          if (esche(corpo)) return 'la tabella degli esami contiene elementi iniettati';
          _labMostraGrafico(ui, ['k']);
          var punto = corpo.querySelector('circle.lab-pt');
          if (!punto) return 'il grafico non ha punti';
          punto.dispatchEvent(new MouseEvent('mouseenter', { clientX: 10, clientY: 10 }));
          await attendi(500);
          var tip = corpo.querySelector('#labTip');
          if (esche(corpo)) return 'grafico o fumetto contengono elementi iniettati';
          if (!tip || tip.textContent.indexOf('Emoglobina<img') < 0) return 'il fumetto non mostra il nome come testo: «' + (tip && tip.textContent) + '»';
          return !eseguiti().length || ('codice ESEGUITO: ' + eseguiti().join(', '));
        } finally { corpo.remove(); }
      });

      // ── 11. tipologie: nome e colore ──
      await prova(S, 'tipologia con colore e nome ostili: liste e «Gestisci tipologie»', async function () {
        window.__xss = [];
        var NOME = 'XSS-PROVA' + esca('tip-nome');
        var COLORE = 'red;position:fixed" onmouseover="' + t('tip-colore') + '"><img src=x onerror="' + t('tip-colore2') + '">';
        await _q(_sb.from('tipologie').upsert({ nome: NOME, colore: COLORE }, { onConflict: 'nome' }));
        daPulire.push(function () { return _q(_sb.from('tipologie').delete().eq('nome', NOME)); });
        await _q(_sb.from('consegne').update({ tipologia_letto: NOME, updated_at: new Date().toISOString() }).eq('letto', LETTO_PROVA));
        await new Promise(function (ok) { _caricaColoriTipologie(ok); });
        await ridisegna();
        var guai = [], colore = _getColoreTipo(NOME);
        if (_coloreSicuro(colore, '') !== colore) guai.push('_getColoreTipo restituisce «' + String(colore).slice(0, 40) + '»');
        _popolaDropdownPazienti(); await attendi(500);
        if (esche(document.getElementById('menuListaPazienti'))) guai.push('menu dei pazienti');
        _aggiornaPannelloRiepilogoLetti();
        await finche(function () { return document.querySelector('#lettiRiepilogoBody .badge'); }, 8000); await attendi(500);
        if (esche(document.getElementById('lettiRiepilogoBody'))) guai.push('riepilogo dei letti');
        _gtApri();
        var lista = document.getElementById('gtListContainer');
        await finche(function () { return lista.querySelector('.gt-riga'); }, 8000); await attendi(400);
        if (esche(lista)) guai.push('«Gestisci tipologie»');
        if (![].some.call(lista.querySelectorAll('.gt-nome-input'), function (i) { return i.value === NOME; })) guai.push('«Gestisci tipologie»: il nome non è nel campo di testo');
        [].forEach.call(lista.querySelectorAll('.gt-swatch, .gt-badge-preview'), function (s) { if (/fixed/.test(getComputedStyle(s).position) || s.hasAttribute('onmouseover')) guai.push('«Gestisci tipologie»: stile iniettato'); });
        if (esche(scheda(LETTO_PROVA).closest('#cardsContainer'))) guai.push('schede');
        if (eseguiti().length) guai.push('codice ESEGUITO: ' + eseguiti().join(', '));
        return !guai.length || ('elementi iniettati in: ' + guai.join(' ‖ '));
      });

      // ── 12. un letto dal nome ostile, e uno dal nome «riservato» ──
      await prova(S, 'letto con un nome ostile scritto direttamente nel database: si disegna come testo, i selettori reggono, il nome non diventa codice', async function () {
        window.__xss = [];
        var NOME = 'X\');' + t('letto-js') + ';//"] ,body [x="' + esca('letto-html');
        await _q(_sb.from('consegne').insert({ letto: NOME, tipologia_letto: 'STANDARD' }));
        daPulire.push(function () { return _q(_sb.from('consegne').delete().eq('letto', NOME)); });
        await _q(_sb.from('consegne').insert({ letto: '__proto__', tipologia_letto: 'STANDARD', nome: 'RISERVATO' }));
        daPulire.push(function () { return _q(_sb.from('consegne').delete().eq('letto', '__proto__')); });
        var c = await finche(function () { return scheda(NOME); }, 12000);
        if (!c) { await ridisegna(); c = scheda(NOME); }
        if (!c) return 'la scheda del letto non è stata disegnata';
        await attendi(800);
        var guai = [];
        if (esche(c.closest('#cardsContainer'))) guai.push('elementi iniettati fra le schede');
        if (c.querySelector('.alt-bed-number').textContent !== NOME) guai.push('il nome non è mostrato come testo');
        if ([].some.call(document.querySelectorAll('.patient-card'), function (x) { return x.getAttribute('data-bed') === '__proto__'; })) guai.push('il letto «__proto__» è stato disegnato');
        // il clic sul cartellino della tipologia eseguiva il nome del letto come codice
        var cartellino = c.querySelector('.tipo-badge-wrap .badge');
        if (/__xss/.test(cartellino.getAttribute('onclick') || '')) guai.push('il nome del letto è ancora dentro il gestore del cartellino');
        cartellino.click();
        try { _applicaLocks((function () { var o = {}; o[NOME] = { token: 'altro', ts: Date.now() }; return o; })()); if (!c.classList.contains('card-locked')) guai.push('il blocco non trova la scheda'); _applicaLocks({}); }
        catch (e) { guai.push('_applicaLocks: ' + e.message); }
        try { _aggiornaBadgePrincipali(); ordinaLetti('numero', true); applicaOrdinamentoSalvato(); _popolaDropdownPazienti(); } catch (e) { guai.push('liste: ' + e.message); }
        await attendi(500);
        var menu = document.getElementById('menuListaPazienti');
        if (esche(menu)) guai.push('menu dei pazienti');
        var voce = [].filter.call(menu.querySelectorAll('[data-scroll-letto]'), function (a) { return a.getAttribute('data-scroll-letto') === NOME; })[0];
        if (!voce) guai.push('il letto non è nel menu dei pazienti'); else { try { voce.click(); } catch (e) { guai.push('clic sul menu: ' + e.message); } }
        var r = await _sbAggiungiLetto('5\'"<b>');
        if (!r || r.success !== false) guai.push('«Aggiungi letto» accetta un nome non valido');
        await attendi(800);
        if (eseguiti().length) guai.push('codice ESEGUITO: ' + eseguiti().join(', '));
        // cancellazione: la scheda deve sparire (selettore col nome ostile)
        await _q(_sb.from('consegne').delete().eq('letto', NOME));
        var sparita = await finche(function () { return !scheda(NOME); }, 9000);
        if (!sparita) guai.push('la scheda non sparisce alla cancellazione del letto');
        return !guai.length || guai.join(' ‖ ');
      });

      // ── 13. backup: elenco dei giorni, griglia di ripristino, stampa di un backup ──
      await prova(S, 'backup con contenuti ostili: giorni, orari, griglia di ripristino, conferma e stampa', async function () {
        window.__xss = [];
        var TS = 1577840000000 + Math.floor(Math.random() * 1000000), TS2 = TS + 1;
        var GIORNO_OSTILE = '2020-01-02\');' + t('arch-giorno') + ';//';
        var LETTO_OSTILE = 'L' + esca('arch-letto');
        await _q(_sb.from('archivio').insert([
          { data_str: '2020-01-01', ts: TS, lab: null, dati: [
            { Letto: '5', Nome: 'ARCHIVIO ' + esca('arch-nome'), Diagnosi: 'dx ' + esca('arch-dx'), TipologiaLetto: '"><img src=x onerror="' + t('arch-tipo') + '">',
              Diaria: esca('arch-diaria') + '<b>dal backup</b>', Allergie: esca('arch-all'), DataNascita: '"><img src=x onerror="' + t('arch-nasc') + '">', DataRicovero: esca('arch-ric'),
              Eta: esca('arch-eta'), CodiceSanitario: esca('arch-cs'), Ossigeno: esca('arch-o2'), Vitto: esca('arch-vitto'), Sesso: 'M', Dimissibile: esca('arch-dim') },
            { Letto: LETTO_OSTILE, Nome: 'SECONDO', Diaria: { oggetto: 'non una stringa' } },
            'non una scheda', null, 42 ] },
          { data_str: GIORNO_OSTILE, ts: TS2, lab: null, dati: [] }
        ]));
        daPulire.push(function () { return _q(_sb.from('archivio').delete().in('ts', [TS, TS2])); });
        var guai = [];
        var giorni = await _sbGetGiorniArchivio();
        if (giorni.indexOf(GIORNO_OSTILE) >= 0) guai.push('il giorno dal nome ostile è nell\'elenco');
        if (giorni.indexOf('2020-01-01') < 0) guai.push('il giorno di prova non è nell\'elenco');
        _ripristinaGoGiorni();
        var lg = document.getElementById('ripristinaListaGiorni');
        await finche(function () { return lg.querySelector('a'); }, 8000); await attendi(300);
        if (lg.querySelector('[onclick*="__xss"]') || esche(lg)) guai.push('elenco dei giorni: contenuto ostile in pagina');
        _ripristinaSelGiorno('2020-01-01', '01/01/2020');
        var lo = document.getElementById('ripristinaListaOrari');
        await finche(function () { return lo.querySelector('a'); }, 8000); await attendi(300);
        if (esche(lo) || esche(document.getElementById('ripristinaBreadcrumb'))) guai.push('elenco degli orari');
        _ripristinaSelTimestamp(String(TS), '00:00:00');
        var gr = document.getElementById('ripristinaGridLetti');
        await finche(function () { return gr.querySelector('.ripristina-letto-card'); }, 8000); await attendi(600);
        if (esche(gr)) guai.push('griglia di ripristino: elementi iniettati');
        var schede = gr.querySelectorAll('.ripristina-letto-card');
        if (schede.length !== 2) guai.push('griglia di ripristino: attese 2 schede, trovate ' + schede.length);
        if (gr.textContent.indexOf('ARCHIVIO <IMG') < 0) guai.push('griglia di ripristino: il nome non compare come testo');
        // conferma di ripristino sul letto dal nome ostile: si ANNULLA, nulla viene ripristinato
        var ostileCard = [].filter.call(schede, function (x) { return x.getAttribute('data-letto') === LETTO_OSTILE; })[0];
        if (!ostileCard) guai.push('griglia: manca la scheda del letto ostile');
        else {
          ostileCard.click();
          _eseguiRipristina();
          var cs = await controllaSwal('conferma di ripristino', '<img');
          if (cs !== true) guai.push(cs);
          try { Swal.clickCancel(); } catch (e) {}
          await attendi(500);
        }
        try { _chiudiRipristina(); } catch (e) {}
        // stampa del backup (pagina a sé, dati presi dall'archivio)
        for (var i = 0; i < 2; i++) {
          var fr = await cornice('/print.html?senzaStampa=1&layout=' + (i ? 'alt' : 'main') + '&dataArchivio=' + TS + '&cb=' + Date.now());
          try {
            var w = fr.contentWindow;
            var pronta = await finche(function () { return w.document.querySelector('#printCards .patient-card, #printCards .alt-row'); }, 20000);
            if (!pronta) { guai.push('stampa del backup (' + (i ? 'alt' : 'main') + '): nulla disegnato'); continue; }
            await attendi(1200);
            if (eseguiti(w).length) guai.push('stampa del backup: codice ESEGUITO ' + eseguiti(w).join(', '));
            if (esche(w.document.getElementById('printCards'))) guai.push('stampa del backup (' + (i ? 'alt' : 'main') + '): elementi iniettati');
            if (w.document.getElementById('printCards').textContent.indexOf('dal backup') < 0) guai.push('stampa del backup: manca il contenuto lecito');
          } finally { fr.remove(); }
        }
        if (eseguiti().length) guai.push('codice ESEGUITO: ' + eseguiti().join(', '));
        return !guai.length || guai.slice(0, 5).join(' ‖ ');
      });

      // ── 14. link utili ──
      await prova(S, 'link utili: nome e indirizzo ostili restano testo, «javascript:» non diventa un collegamento, si modifica e si elimina il link indicato', async function () {
        window.__xss = [];
        var guai = [];
        var prima = await _q(_sb.from('link_utili').select('*').order('id'));
        var a = await _q(_sb.from('link_utili').insert({ nome: 'x\');' + t('lu-nome') + ';//' + esca('lu-nome2'), url: 'javascript:' + t('lu-url') }).select().single());
        var b = await _q(_sb.from('link_utili').insert({ nome: 'PROVA B', url: 'https://example.org/b' }).select().single());
        daPulire.push(function () { return _q(_sb.from('link_utili').delete().in('id', [a.id, b.id])); });
        var contenitore = document.getElementById('linkUtiliList');
        _renderLinkUtili(await _sbGetLinkUtili());
        if (esche(contenitore)) guai.push('elenco: elementi iniettati');
        if (contenitore.querySelector('a[href^="javascript" i], [onclick]')) guai.push('elenco: collegamento javascript: o gestore in linea');
        var rigaA = contenitore.querySelector('[data-lu-riga="' + a.id + '"]');
        if (!rigaA || rigaA.querySelector('a')) guai.push('il link con indirizzo non valido è un collegamento');
        if (!contenitore.querySelector('[data-lu-riga="' + b.id + '"] a[href="https://example.org/b"]')) guai.push('il link lecito non è un collegamento');
        rigaA.querySelector('[data-lu-modifica]').click();
        await attendi(300);
        var campoNome = rigaA.querySelector('[data-lu-campo="nome"]'), campoUrl = rigaA.querySelector('[data-lu-campo="url"]');
        if (!campoNome || campoNome.value !== a.nome || campoUrl.value !== a.url) guai.push('la modifica non mostra i valori salvati');
        // salvataggio con indirizzo non valido: rifiutato
        rigaA.querySelector('[data-lu-salva]').click();
        var avviso = await finestraSwal(4000);
        if (!avviso || avviso.textContent.indexOf('Indirizzo non valido') < 0) guai.push('il salvataggio di un indirizzo javascript: non viene rifiutato');
        await chiudiSwal();
        // modifica ed eliminazione per identificativo: cambia SOLO il link indicato
        var m = await _sbModificaLink(b.id, 'PROVA B2', 'https://example.org/b2');
        var dopo = await _q(_sb.from('link_utili').select('*').order('id'));
        var cambiati = dopo.filter(function (r) { var p = prima.concat([a, b]).filter(function (x) { return x.id === r.id; })[0]; return !p || p.nome !== r.nome || p.url !== r.url; });
        if (!m.success || cambiati.length !== 1 || cambiati[0].id !== b.id || cambiati[0].nome !== 'PROVA B2') guai.push('la modifica ha toccato: ' + cambiati.map(function (r) { return r.id; }).join(',') + ' (atteso solo ' + b.id + ')');
        await _sbEliminaLink(a.id); await _sbEliminaLink(b.id);
        var fine = await _q(_sb.from('link_utili').select('*').order('id'));
        if (JSON.stringify(fine) !== JSON.stringify(prima)) guai.push('dopo l\'eliminazione dei due link di prova l\'elenco non è quello di partenza');
        var nulla = await _sbModificaLink(-12345, 'x', 'https://example.org');
        if (nulla.success) guai.push('la modifica di un link inesistente risulta riuscita');
        _renderLinkUtili(fine);
        if (eseguiti().length) guai.push('codice ESEGUITO: ' + eseguiti().join(', '));
        return !guai.length || guai.join(' ‖ ');
      });
    } finally {
      for (var i = daPulire.length - 1; i >= 0; i--) { try { await daPulire[i](); } catch (e) {} }
      if (scritto) {
        try { await _sbDimettiLetto(LETTO_PROVA); } catch (e) {}
        try { await _q(_sb.from('consegne').update({ tipologia_letto: (rigaOriginale && rigaOriginale.tipologia_letto) || 'STANDARD', updated_at: new Date().toISOString() }).eq('letto', LETTO_PROVA)); } catch (e) {}
        await attendi(1500);
        try { await new Promise(function (ok) { _caricaColoriTipologie(ok); }); } catch (e) {}
        try { await ridisegna(); } catch (e) {}
      }
    }
    await prova(S, 'tutto rimesso a posto: letto di prova libero, nessuna riga di prova rimasta', async function () {
      var r = await _q(_sb.from('consegne').select('nome,diagnosi,diaria,sesso,dimissibile,tipologia_letto').eq('letto', LETTO_PROVA).maybeSingle());
      if (!(r && !r.nome && !r.diagnosi && !r.diaria && !r.sesso && !r.dimissibile)) return 'contenuto rimasto nel letto di prova: ' + JSON.stringify(r).slice(0, 120);
      if (r.tipologia_letto !== ((rigaOriginale && rigaOriginale.tipologia_letto) || 'STANDARD')) return 'tipologia del letto di prova non ripristinata: ' + r.tipologia_letto;
      var letti = await _q(_sb.from('consegne').select('letto'));
      if (letti.length !== 28) return 'i letti sono ' + letti.length + ' invece di 28';
      var tip = await _q(_sb.from('tipologie').select('nome'));
      if (tip.some(function (x) { return /XSS-PROVA/.test(x.nome); })) return 'tipologia di prova rimasta';
      var arch = await _q(_sb.from('archivio').select('ts').lt('ts', 1600000000000));
      if (arch.length) return 'backup di prova rimasti: ' + arch.length;
      var lu = await _q(_sb.from('link_utili').select('id,nome'));
      return !lu.some(function (x) { return /PROVA B|__xss/.test(x.nome || ''); }) || 'link di prova rimasti';
    });
    await prova(S, 'il registro annota ciò che il filtro ha tolto, coi soli nomi di tag e attributi (mai il contenuto)', async function () {
      await attendi(1500);
      var righe = await _q(_sb.from('logs').select('messaggio,descrizione').eq('tipo', 'html-ripulito').gte('ts', inizioSezione));
      // le righe di questa prova non restano nel registro del collaudo
      try { await _q(_sb.from('logs').delete().eq('tipo', 'html-ripulito').gte('ts', inizioSezione)); } catch (e) {}
      if (!righe.some(function (r) { return /<img>/.test(r.messaggio); })) return 'nessuna segnalazione per i contenuti ostili (' + righe.length + ' righe)';
      var conContenuto = righe.filter(function (r) { return /__xss|ROSSI|onerror="|javascript:|src=x/i.test(r.messaggio + ' ' + (r.descrizione || '')); });
      return !conContenuto.length || ('una segnalazione riporta il contenuto: ' + conContenuto[0].messaggio.slice(0, 140));
    });
    await prova(S, 'nessun errore JavaScript durante le prove', function () {
      var nuovi = (window.__erroriBanco || []).slice(erroriPrima);
      return !nuovi.length || nuovi.slice(0, 4).join(' | ');
    });
  }

  // ══════════════════════════════════════════════════════════════════════
  // MAIL DIMISSIONI: il mittente scelto in «Impostazioni email» e i permessi
  // di invio, verificati PRIMA di mostrare la procedura.
  // La cassaforte del collaudo passa alle credenziali finte (il banco mette da
  // parte il consenso vero e lo rimette alla fine): nessuna mail parte davvero.
  // Al posto della finestra di Google c'è un codice che il finto Google capisce.
  // ══════════════════════════════════════════════════════════════════════

  // Il server della mail risponde alle richieste di «stato» ciò che decide la
  // prova (tutte le altre chiamate passano): serve a mettere la pagina davanti
  // a risposte che il collaudo non sa produrre a comando.
  var _fetchVero = null;
  function fingiStatoMail(risposta) {
    if (!_fetchVero) _fetchVero = window.fetch;
    var vero = _fetchVero;
    window.fetch = function (u, o) {
      if (String(u).indexOf('/functions/v1/google-token') >= 0 && o && /"azione":"stato"/.test(String(o.body || ''))) {
        return Promise.resolve(new Response(JSON.stringify(risposta), { status: 200, headers: { 'Content-Type': 'application/json' } }));
      }
      return vero.apply(window, arguments);
    };
  }
  function statoMailVero() { if (_fetchVero) { window.fetch = _fetchVero; _fetchVero = null; } }
  // lo stesso per un'azione qualunque («config», «scambia», …); si toglie con statoMailVero()
  function fingiMail(azione, risposta) {
    if (!_fetchVero) _fetchVero = window.fetch;
    var vero = _fetchVero;
    window.fetch = function (u, o) {
      if (String(u).indexOf('/functions/v1/google-token') >= 0 && o && String(o.body || '').indexOf('"azione":"' + azione + '"') >= 0) {
        return Promise.resolve(new Response(JSON.stringify(risposta), { status: 200, headers: { 'Content-Type': 'application/json' } }));
      }
      return vero.apply(window, arguments);
    };
  }

  async function sezioneMail() {
    var S = 'mail';
    var ALTRO = 'mittente-prova@example.com', TERZO = 'terzo-mittente@example.com';
    var CHIAVI = ['MAIL_DIMISSIONI_MITTENTE', 'MAIL_DIMISSIONI_DESTINATARI', 'MAIL_DIMISSIONI_OGGETTO', 'MAIL_DIMISSIONI_CORPO', 'MAIL_DIMISSIONI_CHIUSURA'];
    var segreto = null, clientFinto = '', prima = {}, letto = false, gisVero = null, richieste = [], prossimoCodice = null, reparto = '';
    var richiestePrima = 0, richiesteDopo = function () { return richieste.length - richiestePrima; };
    var erroriPrima = (window.__erroriBanco || []).length, inizio = Date.now();
    var imp = async function (chiave) {
      var r = await _sb.from('impostazioni').select('valore').eq('chiave', chiave).maybeSingle();
      if (r.error) throw new Error(r.error.message);
      return r.data ? r.data.valore : null;
    };
    var scriviImp = async function (chiave, valore) {
      var r = (valore === null || valore === undefined)
        ? await _sb.from('impostazioni').delete().eq('chiave', chiave)
        : await _sb.from('impostazioni').upsert([{ chiave: chiave, valore: valore }], { onConflict: 'chiave' });
      if (r.error) throw new Error(r.error.message);
    };
    var come = function (email) { return segreto + '.come.' + btoa(email).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, ''); };
    // La finestra Swal aperta ADESSO. Una finestra chiusa può restare nel DOM
    // (classe swal2-hide) finché il browser non consegna la fine dell'animazione,
    // cosa che a pannello nascosto non fa: per l'app è chiusa, e così per le prove.
    var aperta = function () { var p = Swal.getPopup(); return (p && Swal.isVisible() && !p.classList.contains('swal2-hide')) ? p : null; };
    var titolo = function () { var p = aperta(), t = p ? p.querySelector('.swal2-title') : null; return t ? t.textContent.trim() : ''; };
    var testoFinestra = function () { var p = aperta(); return p ? p.textContent.replace(/\s+/g, ' ').trim() : ''; };
    var attendiTitolo = function (re, ms) { return finche(function () { return re.test(titolo()) ? titolo() : false; }, ms || 15000); };
    var bottoneGiallo = function () { var li = document.getElementById('navItemCassaforte'); return !!li && getComputedStyle(li).display !== 'none'; };
    var procedura = function () { var p = aperta(); return !!(p && (p.querySelector('#mailDimDest') || p.querySelector('#mailDimAnteprima'))); };
    var postaDi = async function (oggetto) { var r = await fetch('/banco/posta?oggetto=' + encodeURIComponent(oggetto)); return r.ok ? await r.json() : { errore: 'HTTP ' + r.status }; };
    // riempie la procedura e arriva alla conferma; restituisce '' oppure il motivo per cui non ci riesce
    var compila = async function (oggetto) {
      if (!document.getElementById('mailDimDest')) return 'procedura non aperta: «' + titolo() + '»';
      document.getElementById('mailDimDest').value = 'destinatario1@example.com';
      document.getElementById('mailDimOggetto').value = oggetto;
      if (!document.querySelector('.mail-dim-cb:checked')) {
        var cb = document.querySelector('.mail-dim-cb'); if (!cb) return 'nessun letto in elenco';
        cb.checked = true; cb.dispatchEvent(new Event('change', { bubbles: true })); await attendi(250);
      }
      Swal.getConfirmButton().click();
      if (!(await attendiTitolo(/Confermi l.invio/))) return 'conferma non comparsa: «' + titolo() + '» ' + testoFinestra().slice(0, 160);
      return '';
    };

    try {
      var rs = await fetch('/banco/cassaforte?modo=finta', { method: 'POST' });
      if (rs.ok) { var jr = await rs.json(); segreto = jr.segreto; clientFinto = jr.clientFinto || ''; }
    } catch (e) {}
    await prova(S, 'il banco passa la cassaforte del collaudo alle credenziali finte', function () { return (!!segreto && !!clientFinto) || 'credenziali finte non disponibili (banco.js aggiornato e riavviato? funzioni-collaudo.js configura?)'; });
    if (!segreto || !clientFinto) return;

    try {
      for (var i = 0; i < CHIAVI.length; i++) prima[CHIAVI[i]] = await imp(CHIAVI[i]);
      letto = true;
      reparto = String((await imp('ACCOUNT_LOGIN')) || '').toLowerCase().trim();
      await scriviImp('MAIL_DIMISSIONI_MITTENTE', null);          // si parte dal caso «mai impostato»

      var haGoogle = typeof google !== 'undefined' && !!google.accounts && !!google.accounts.oauth2;
      await prova(S, 'la libreria di Google è caricata (serve alla finestra del consenso)', function () { return haGoogle || 'accounts.google.com non raggiungibile: le prove del consenso non possono girare'; });
      if (!haGoogle) return;
      gisVero = google.accounts.oauth2.initCodeClient;
      google.accounts.oauth2.initCodeClient = function (cfg) {   // la finestra di Google: risponde ciò che decide la prova
        richieste.push({ login_hint: cfg.login_hint, scope: cfg.scope, client_id: cfg.client_id });
        return { requestCode: function () {
          var c = prossimoCodice; prossimoCodice = null;
          setTimeout(function () { if (c) cfg.callback({ code: c }); else cfg.error_callback({ type: 'popup_closed' }); }, 60);
        } };
      };

      await prova(S, 'menu della rotellina: voce «Impostazioni email»', function () {
        var a = document.querySelector('a.dropdown-item[onclick="_apriImpostazioniEmail()"]');
        return (!!a && /Impostazioni email/.test(a.textContent)) || 'voce assente';
      });

      await prova(S, 'Impostazioni email: precompilata con l\'indirizzo del reparto, permessi già concessi', async function () {
        window._apriImpostazioniEmail();
        var campo = await finche(function () { return document.getElementById('impMailMittente'); }, 15000);
        if (!campo) return 'la finestra non si è aperta: «' + titolo() + '»';
        var p = document.getElementById('impMailPermessi').textContent;
        if (!reparto || campo.value !== reparto) return 'campo «' + campo.value + '» invece di «' + reparto + '»';
        return (/: concessi/.test(p) && p.indexOf(reparto) >= 0) || 'stato dei permessi: ' + p;
      });

      await prova(S, 'un indirizzo non valido non viene salvato', async function () {
        var campo = document.getElementById('impMailMittente'); if (!campo) return 'finestra chiusa';
        campo.value = 'non un indirizzo';
        Swal.clickConfirm(); await attendi(600);
        var v = document.querySelector('.swal2-validation-message');
        if (!(v && getComputedStyle(v).display !== 'none')) return 'nessun messaggio di errore';
        return (!!document.getElementById('impMailMittente') && (await imp('MAIL_DIMISSIONI_MITTENTE')) === null) || 'salvato lo stesso';
      });

      await prova(S, 'nuovo mittente: salvato in minuscolo, i permessi risultano da concedere e compare il bottone giallo', async function () {
        var campo = document.getElementById('impMailMittente'); if (!campo) return 'finestra chiusa';
        campo.value = '  Mittente-Prova@Example.com ';
        Swal.clickConfirm();
        if (!(await attendiTitolo(/Servono i permessi di invio/))) return 'finestra attesa non comparsa: «' + titolo() + '»';
        var salvato = await imp('MAIL_DIMISSIONI_MITTENTE'), st = await window._cassaforteStato(true);
        if (salvato !== ALTRO) return 'salvato «' + salvato + '»';
        if (st.autorizzato !== false || st.mittente !== ALTRO || st.verificato !== true) return 'stato del server: ' + JSON.stringify(st);
        return bottoneGiallo() || 'il bottone giallo non è comparso';
      });

      await prova(S, '«Più tardi»: nessuna finestra di Google', async function () {
        Swal.clickCancel(); await attendi(700);
        return (richieste.length === 0 && !aperta()) || 'richieste a Google: ' + richieste.length + ', finestra «' + titolo() + '»';
      });

      await prova(S, 'il cambio di mittente resta nel registro', async function () {
        return (await finche(async function () {
          var q = await _sb.from('logs').select('descrizione').eq('tipo', 'mail-mittente').gte('ts', inizio).order('ts', { ascending: false }).limit(1);
          return !!(q.data && q.data[0] && String(q.data[0].descrizione).indexOf('A: ' + ALTRO) >= 0 && String(q.data[0].descrizione).indexOf('Da: ' + reparto) >= 0);
        }, 8000, 500)) || 'riga non trovata';
      });

      await prova(S, 'Invia mail dimissioni senza permessi: li chiede e NON mostra la procedura', async function () {
        window.apriModalMailDimissioni();
        if (!(await attendiTitolo(/Servono i permessi di invio/))) return 'finestra attesa non comparsa: «' + titolo() + '»';
        if (procedura()) return 'la procedura è visibile';
        return testoFinestra().indexOf(ALTRO) >= 0 || 'non dice per quale indirizzo: ' + testoFinestra().slice(0, 160);
      });

      await prova(S, '«Cambia mittente» dalla richiesta dei permessi apre Impostazioni email', async function () {
        var b = Swal.getDenyButton();
        if (!b || getComputedStyle(b).display === 'none' || !/Cambia mittente/.test(b.textContent)) return 'bottone assente';
        Swal.clickDeny();
        var campo = await finche(function () { return document.getElementById('impMailMittente'); }, 15000);
        if (!campo) return 'Impostazioni email non si è aperta: «' + titolo() + '»';
        if (campo.value !== ALTRO || richieste.length !== 0) return 'campo «' + campo.value + '», richieste a Google ' + richieste.length;
        await chiudiSwal();
        window.apriModalMailDimissioni();                // si torna alla richiesta dei permessi per le prove che seguono
        return !!(await attendiTitolo(/Servono i permessi di invio/)) || 'la richiesta dei permessi non ricompare: «' + titolo() + '»';
      });

      await prova(S, 'consenso dato con un altro account: rifiutato, procedura ancora nascosta', async function () {
        prossimoCodice = segreto;                       // il finto Google risponde: ha acconsentito l'account del reparto
        Swal.clickConfirm();
        if (!(await attendiTitolo(/Permessi non registrati/))) return 'finestra attesa non comparsa: «' + titolo() + '»';
        var testo = testoFinestra();
        if (richieste.length !== 1 || richieste[0].login_hint !== ALTRO) return 'a Google non è stato suggerito il mittente: ' + JSON.stringify(richieste);
        if (!/gmail\.send/.test(richieste[0].scope)) return 'permesso richiesto: ' + richieste[0].scope;
        if (procedura()) return 'la procedura è visibile';
        return (testo.indexOf('account non ammesso') >= 0 && testo.indexOf(reparto) >= 0) || 'messaggio: ' + testo.slice(0, 220);
      });

      await prova(S, 'finestra di Google chiusa senza consenso: niente procedura', async function () {
        await chiudiSwal();
        window.apriModalMailDimissioni();
        if (!(await attendiTitolo(/Servono i permessi di invio/))) return 'finestra attesa non comparsa: «' + titolo() + '»';
        prossimoCodice = null;                          // l'utente chiude la finestra di Google
        Swal.clickConfirm();
        if (!(await attendiTitolo(/Permessi non concessi/))) return 'esito: «' + titolo() + '»';
        return !procedura() || 'la procedura è visibile';
      });

      await prova(S, 'consenso dato con l\'account del mittente: la procedura compare, col mittente non modificabile', async function () {
        await chiudiSwal();
        window.apriModalMailDimissioni();
        if (!(await attendiTitolo(/Servono i permessi di invio/))) return 'finestra attesa non comparsa: «' + titolo() + '»';
        prossimoCodice = come(ALTRO);
        Swal.clickConfirm();
        await finche(function () { return /Permessi concessi/.test(titolo()) || !!document.getElementById('mailDimMittente'); }, 20000);
        if (/Permessi concessi/.test(titolo())) Swal.clickConfirm();      // senza aspettare che si chiuda da sola
        var campo = await finche(function () { return document.getElementById('mailDimMittente'); }, 25000);
        if (!campo) return 'la procedura non è comparsa: «' + titolo() + '» ' + testoFinestra().slice(0, 160);
        if (campo.value !== ALTRO) return 'mittente mostrato «' + campo.value + '»';
        if (!(campo.readOnly && campo.disabled)) return 'il campo del mittente è modificabile';
        return (procedura() && !bottoneGiallo()) || 'manca il resto della procedura, o il bottone giallo è rimasto';
      });

      await prova(S, 'annullare la conferma riporta alla procedura com\'era (destinatari, oggetto, selezione, anteprima ritoccata)', async function () {
        var caselle = document.querySelectorAll('.mail-dim-cb');
        if (caselle.length < 3) return 'servono almeno tre letti in elenco';
        // selezione scelta a mano: solo il secondo e il terzo letto
        for (var n = 0; n < caselle.length; n++) {
          var voluta = (n === 1 || n === 2);
          if (caselle[n].checked !== voluta) { caselle[n].checked = voluta; caselle[n].dispatchEvent(new Event('change', { bubbles: true })); }
        }
        await attendi(250);
        var ante = document.getElementById('mailDimAnteprima');
        ante.insertAdjacentHTML('beforeend', '<p>Riga aggiunta a mano dalla prova</p>');
        ante.dispatchEvent(new Event('input', { bubbles: true }));
        var corpoPrima = ante.innerHTML, lettiPrima = [caselle[1].getAttribute('data-letto'), caselle[2].getAttribute('data-letto')].join(',');
        document.getElementById('mailDimDest').value = 'primo@example.com; secondo@example.com';
        document.getElementById('mailDimOggetto').value = 'Oggetto della prova di ritorno';
        Swal.getConfirmButton().click();
        if (!(await attendiTitolo(/Confermi l.invio/))) return 'conferma non comparsa: «' + titolo() + '» ' + testoFinestra().slice(0, 160);
        Swal.clickCancel();
        var campo = await finche(function () { return procedura() ? document.getElementById('mailDimDest') : false; }, 8000);
        if (!campo) return 'la procedura non è tornata: «' + titolo() + '»';
        await attendi(200);
        var lettiDopo = Array.prototype.map.call(document.querySelectorAll('.mail-dim-cb:checked'), function (c) { return c.getAttribute('data-letto'); }).join(',');
        var guai = [];
        if (campo.value !== 'primo@example.com; secondo@example.com') guai.push('destinatari «' + campo.value + '»');
        if (document.getElementById('mailDimOggetto').value !== 'Oggetto della prova di ritorno') guai.push('oggetto «' + document.getElementById('mailDimOggetto').value + '»');
        if (lettiDopo !== lettiPrima) guai.push('selezione ' + lettiDopo + ' invece di ' + lettiPrima);
        if (document.getElementById('mailDimAnteprima').innerHTML !== corpoPrima) guai.push('anteprima diversa');
        if (document.getElementById('mailDimMittente').value !== ALTRO) guai.push('mittente «' + document.getElementById('mailDimMittente').value + '»');
        // l'anteprima era stata ritoccata: cambiare la selezione deve chiedere prima di rigenerarla
        var prima0 = document.querySelectorAll('.mail-dim-cb')[0];
        prima0.checked = true; prima0.dispatchEvent(new Event('change', { bubbles: true }));
        if (!(await attendiTitolo(/Rigenerare l.anteprima/, 4000))) guai.push('nessuna domanda prima di rigenerare un\'anteprima ritoccata');
        return guai.length ? guai.join(' ‖ ') : true;
      });

      await prova(S, 'mittente cambiato mentre la finestra è aperta: la mail non parte a nome di un altro', async function () {
        var oggetto = 'Prova mittente cambiato ' + Date.now();
        // la prova precedente ha lasciato aperta una domanda: si riparte dalla procedura
        await chiudiSwal();
        window.apriModalMailDimissioni();
        if (!(await finche(function () { return procedura(); }, 25000))) return 'la procedura non si riapre: «' + titolo() + '»';
        await attendi(300);
        await scriviImp('MAIL_DIMISSIONI_MITTENTE', TERZO);
        try {
          var guaio = await compila(oggetto); if (guaio) return guaio;
          Swal.clickConfirm();
          if (!(await attendiTitolo(/Errore invio/, 25000))) return 'esito: «' + titolo() + '» ' + testoFinestra().slice(0, 200);
          var testo = testoFinestra();
          if (testo.indexOf('il mittente della mail è cambiato') < 0 || testo.indexOf(TERZO) < 0) return 'messaggio: ' + testo.slice(0, 220);
          return (await postaDi(oggetto)) === null || 'la mail è partita lo stesso';
        } finally { await scriviImp('MAIL_DIMISSIONI_MITTENTE', ALTRO); }
      });

      await prova(S, 'conferma e invio: il mittente è dichiarato e la mail parte a suo nome (simulata)', async function () {
        var oggetto = 'Prova del mittente ' + Date.now();
        await chiudiSwal();
        window.apriModalMailDimissioni();
        if (!(await finche(function () { return document.getElementById('mailDimMittente'); }, 25000))) return 'la procedura non si riapre: «' + titolo() + '» ' + testoFinestra().slice(0, 160);
        await attendi(300);
        var guaio = await compila(oggetto); if (guaio) return guaio;
        if (testoFinestra().indexOf('Mittente: ' + ALTRO) < 0) return 'la conferma non dichiara il mittente: ' + testoFinestra().slice(0, 200);
        Swal.clickConfirm();
        if (!(await attendiTitolo(/Mail inviata/, 25000))) return 'esito: «' + titolo() + '» ' + testoFinestra().slice(0, 200);
        var esito = testoFinestra(), m = await postaDi(oggetto);
        if (!m || m.errore) return 'il finto Google non ha ricevuto la mail';
        if (m.mittente !== ALTRO || m.reale !== false || m.destinatari !== 'destinatario1@example.com') return 'spedita da «' + m.mittente + '» a «' + m.destinatari + '», reale: ' + m.reale;
        return esito.indexOf(ALTRO) >= 0 || 'l\'esito non nomina il mittente: ' + esito.slice(0, 160);
      });

      await prova(S, 'Google non conferma i permessi: avviso, e la procedura non compare', async function () {
        await chiudiSwal();
        fingiStatoMail({ configurato: true, autorizzato: true, verificato: false, problema: 'refresh rifiutato: Guasto simulato di Google', email: ALTRO, mittente: ALTRO });
        try {
          window.apriModalMailDimissioni();
          if (!(await attendiTitolo(/Permessi di invio non verificabili/))) return 'cancello: «' + titolo() + '»';
          var testo = testoFinestra();
          if (procedura()) return 'la procedura è visibile';
          if (testo.indexOf('Google non conferma') < 0 || testo.indexOf('Guasto simulato') < 0) return 'messaggio: ' + testo.slice(0, 200);
          await chiudiSwal();
          window._apriImpostazioniEmail();
          var riga = await finche(function () { return document.getElementById('impMailPermessi'); }, 15000);
          return (!!riga && /non verificabili in questo momento/.test(riga.textContent)) || 'Impostazioni email: ' + (riga ? riga.textContent : 'non aperta');
        } finally { statoMailVero(); }
      });

      await prova(S, 'il server della mail risponde con un errore: avviso, mai la richiesta del client secret', async function () {
        await chiudiSwal();
        fingiStatoMail({ code: 'BOOT_ERROR', message: 'Function failed to start (prova)' });
        try {
          window.apriModalMailDimissioni();
          if (!(await attendiTitolo(/Permessi di invio non verificabili/))) return 'cancello: «' + titolo() + '»';
          if (procedura() || testoFinestra().indexOf('Function failed to start') < 0) return 'cancello: ' + testoFinestra().slice(0, 200);
          await chiudiSwal();
          var esito = window._cassaforteAutorizza();          // il bottone giallo
          if (!(await attendiTitolo(/Server della mail non raggiungibile/))) return 'bottone giallo: «' + titolo() + '»';
          if (document.getElementById('cassaSecret')) return 'viene chiesto il client secret';
          if ((await esito) !== false) return 'la concessione risulta riuscita';
          await chiudiSwal();
          window._apriImpostazioniEmail();
          return !!(await attendiTitolo(/Impostazioni email non disponibili/)) || 'Impostazioni email: «' + titolo() + '»';
        } finally { statoMailVero(); }
      });

      await prova(S, 'server della mail non ancora aggiornato: la procedura si apre col mittente della cassaforte', async function () {
        await chiudiSwal();
        fingiStatoMail({ configurato: true, autorizzato: true, email: 'vecchio-server@example.com' });
        try {
          window.apriModalMailDimissioni();
          var campo = await finche(function () { return document.getElementById('mailDimMittente'); }, 25000);
          if (!campo) return 'la procedura non si apre: «' + titolo() + '»';
          if (campo.value !== 'vecchio-server@example.com') return 'mittente mostrato «' + campo.value + '»';
          await chiudiSwal();
          window._apriImpostazioniEmail();
          return !!(await finche(function () { return /non è ancora aggiornato/.test(testoFinestra()); }, 15000)) || 'Impostazioni email: ' + testoFinestra().slice(0, 160);
        } finally { statoMailVero(); }
      });

      // ── il client secret: si scrive, non si legge ──
      var scriviSegreto = async function (valore) {       // scrive nel campo e preme «Verifica e salva»
        var campo = await finche(function () { var p = aperta(); return p ? p.querySelector('#cassaSecret') : null; }, 15000);
        if (!campo) return 'la finestra del client secret non è aperta: «' + titolo() + '»';
        campo.value = valore;
        Swal.clickConfirm();
        return '';
      };
      var avvisoCampo = function () { var p = aperta(), v = p ? p.querySelector('.swal2-validation-message') : null; return (v && getComputedStyle(v).display !== 'none') ? v.textContent.trim() : ''; };

      await prova(S, 'Impostazioni email: il client secret risulta registrato e non compare da nessuna parte', async function () {
        await chiudiSwal();
        window._apriImpostazioniEmail();
        var riga = await finche(function () { return document.getElementById('impMailSegretoStato'); }, 15000);
        if (!riga) return 'riquadro del client secret assente: «' + titolo() + '»';
        var b = document.getElementById('impMailSegreto');
        if (!/: registrato/.test(riga.textContent) || !b || !/Cambia/.test(b.textContent)) return 'stato «' + riga.textContent + '», bottone «' + (b ? b.textContent : 'assente') + '»';
        return document.documentElement.innerHTML.indexOf(clientFinto) < 0 || 'il client secret custodito è nella pagina';
      });

      await prova(S, 'cambio del client secret: il campo è vuoto, non è un campo password e non viene mai precompilato', async function () {
        document.getElementById('impMailSegreto').click();
        var campo = await finche(function () { var p = aperta(); return p ? p.querySelector('#cassaSecret') : null; }, 15000);
        if (!campo) return 'la finestra del client secret non si è aperta: «' + titolo() + '»';
        if (campo.value !== '' || campo.getAttribute('value')) return 'campo precompilato';
        if (campo.type !== 'text' || getComputedStyle(campo).webkitTextSecurity !== 'disc') return 'tipo ' + campo.type + ', mascheratura ' + getComputedStyle(campo).webkitTextSecurity;
        return testoFinestra().indexOf('Stai sostituendo') >= 0 || 'non dice che si sta sostituendo quello custodito';
      });

      await prova(S, 'un ID client incollato al posto del secret viene riconosciuto, e il campo si svuota', async function () {
        var guaio = await scriviSegreto(GOOGLE_CLIENT_ID); if (guaio) return guaio;
        if (!(await finche(function () { return /ID client/.test(avvisoCampo()); }, 6000))) return 'avviso: «' + avvisoCampo() + '»';
        var campo = document.getElementById('cassaSecret');
        return (!!campo && campo.value === '' && !!aperta()) || 'la finestra si è chiusa o il campo non è stato svuotato';
      });

      await prova(S, 'un secret che Google non riconosce: errore, richiesto di nuovo, e quello custodito resta al suo posto', async function () {
        var guaio = await scriviSegreto(clientFinto + '-sbagliato'); if (guaio) return guaio;
        if (!(await finche(function () { return /non riconosce/.test(avvisoCampo()); }, 15000))) return 'avviso: «' + avvisoCampo() + '» titolo «' + titolo() + '»';
        var campo = document.getElementById('cassaSecret');
        if (!(campo && campo.value === '' && /Client secret di Google/.test(titolo()))) return 'la finestra non è rimasta a richiederlo';
        var st = await window._cassaforteStato(true);
        return (st.autorizzato === true && st.configurato === true && !st.segreto_errato) || 'il secret custodito è cambiato: ' + JSON.stringify(st);
      });

      await prova(S, 'un secret valido: salvato dopo la verifica, i permessi restano attivi, nulla resta in pagina', async function () {
        var guaio = await scriviSegreto(clientFinto + '-bis'); if (guaio) return guaio;
        if (!(await attendiTitolo(/Client secret registrato/, 15000))) return 'esito: «' + titolo() + '» ' + avvisoCampo();
        if (testoFinestra().indexOf('permessi di invio sono attivi') < 0) return 'messaggio: ' + testoFinestra().slice(0, 160);
        var st = await window._cassaforteStato(true);
        if (!(st.autorizzato === true && st.verificato === true)) return 'stato dopo il cambio: ' + JSON.stringify(st);
        if ('client_secret' in st || JSON.stringify(st).indexOf(clientFinto) >= 0) return 'lo stato del server contiene il secret';
        return document.documentElement.innerHTML.indexOf(clientFinto) < 0 || 'il secret scritto è rimasto nella pagina';
      });

      await prova(S, 'il messaggio d\'errore del server sul secret è mostrato come testo', async function () {
        await chiudiSwal();
        window.__xss = [];
        window._apriImpostazioniEmail();
        var b = await finche(function () { return document.getElementById('impMailSegreto'); }, 15000);
        if (!b) return 'Impostazioni email non si apre';
        b.click();
        fingiMail('config', { errore: 'ROSSI ' + esca('secret') });
        try {
          var guaio = await scriviSegreto('una-stringa-qualunque-lunga-abbastanza'); if (guaio) return guaio;
          if (!(await finche(function () { return /ROSSI/.test(avvisoCampo()); }, 8000))) return 'avviso: «' + avvisoCampo() + '»';
          await attendi(300);
          return (avvisoCampo().indexOf('ROSSI <img') >= 0 && !esche(aperta()) && !eseguiti().length) || 'il messaggio è diventato HTML';
        } finally { statoMailVero(); }
      });

      await prova(S, 'annullare il cambio del secret riporta a Impostazioni email', async function () {
        Swal.clickCancel();
        return !!(await finche(function () { var p = aperta(); return p ? p.querySelector('#impMailMittente') : null; }, 15000)) || 'non si torna alle impostazioni: «' + titolo() + '»';
      });

      await prova(S, 'il secret custodito smette di valere: Invia mail lo dice, lo fa reinserire e riparte senza rifare il consenso', async function () {
        await chiudiSwal();
        var r1 = await fetch('/banco/cassaforte?modo=segreto-sbagliato', { method: 'POST' });
        if (!r1.ok) return 'il banco non ha cambiato il secret: HTTP ' + r1.status;
        richiestePrima = richieste.length;
        window.apriModalMailDimissioni();
        if (!(await attendiTitolo(/Client secret non più valido/))) return 'cancello: «' + titolo() + '»';
        if (procedura()) return 'la procedura è visibile';
        Swal.clickConfirm();
        var guaio = await scriviSegreto(clientFinto); if (guaio) return guaio;
        await finche(function () { return /Client secret registrato/.test(titolo()) || procedura(); }, 20000);
        if (/Client secret registrato/.test(titolo())) { if (testoFinestra().indexOf('già attivi') < 0) return 'chiede di rifare il consenso: ' + testoFinestra().slice(0, 160); Swal.clickConfirm(); }
        var campo = await finche(function () { return procedura() ? document.getElementById('mailDimMittente') : null; }, 25000);
        return (!!campo && campo.value === ALTRO && richiesteDopo() === 0) || 'la procedura non è ripartita: «' + titolo() + '», richieste a Google ' + richiesteDopo();
      });

      await prova(S, 'primo avvio: prima il client secret (uno sbagliato viene respinto), poi il consenso di Google, poi la procedura', async function () {
        await chiudiSwal();
        var r2 = await fetch('/banco/cassaforte?modo=senza-segreto', { method: 'POST' });
        if (!r2.ok) return 'il banco non ha tolto il secret: HTTP ' + r2.status;
        var prima0 = richieste.length;
        window.apriModalMailDimissioni();
        if (!(await attendiTitolo(/Servono i permessi di invio/))) return 'cancello: «' + titolo() + '»';
        Swal.clickConfirm();
        var campo = await finche(function () { var p = aperta(); return p ? p.querySelector('#cassaSecret') : null; }, 15000);
        if (!campo) return 'il client secret non viene chiesto: «' + titolo() + '»';
        if (testoFinestra().indexOf('Primo avvio') < 0) return 'non spiega perché lo chiede';
        if (richieste.length !== prima0) return 'la finestra di Google si è aperta prima del client secret';
        var guaio = await scriviSegreto(clientFinto + '-sbagliato'); if (guaio) return guaio;
        if (!(await finche(function () { return /non riconosce/.test(avvisoCampo()); }, 15000))) return 'secret sbagliato accettato: «' + titolo() + '» ' + avvisoCampo();
        if (richieste.length !== prima0) return 'con un secret sbagliato si è aperta la finestra di Google';
        guaio = await scriviSegreto(clientFinto); if (guaio) return guaio;
        if (!(await attendiTitolo(/Client secret registrato/, 15000))) return 'esito: «' + titolo() + '» ' + avvisoCampo();
        if (testoFinestra().indexOf('Ora manca il permesso') < 0) return 'messaggio: ' + testoFinestra().slice(0, 160);
        prossimoCodice = come(ALTRO);
        Swal.clickConfirm();
        await finche(function () { return /Permessi concessi/.test(titolo()) || procedura(); }, 20000);
        if (/Permessi concessi/.test(titolo())) Swal.clickConfirm();
        var mitt = await finche(function () { return procedura() ? document.getElementById('mailDimMittente') : null; }, 25000);
        return (!!mitt && mitt.value === ALTRO && richieste.length === prima0 + 1) || 'la procedura non è comparsa: «' + titolo() + '», richieste a Google ' + (richieste.length - prima0);
      });

      await prova(S, 'un mittente ostile scritto nel database non diventa HTML e non viene usato', async function () {
        await chiudiSwal();
        window.__xss = [];
        // senza spazi: passerebbe un controllo «qualcosa@qualcosa.xx» fatto alla buona
        await scriviImp('MAIL_DIMISSIONI_MITTENTE', '"><svg/onload=' + traccia('mittente') + '>@example.com');
        var st = await window._cassaforteStato(true);
        if (st.mittente !== null || st.autorizzato !== false) return 'il server lo accetta: ' + JSON.stringify(st);
        window.apriModalMailDimissioni();
        if (!(await attendiTitolo(/Mittente non disponibile/))) return 'finestra attesa non comparsa: «' + titolo() + '»';
        if (procedura() || esche(aperta())) return 'la procedura è visibile, o il valore è finito in pagina';
        await chiudiSwal();
        window._apriImpostazioniEmail();
        var campo = await finche(function () { return document.getElementById('impMailMittente'); }, 15000);
        if (!campo) return 'Impostazioni email non si apre: «' + titolo() + '»';
        await attendi(300);
        var p = Swal.getPopup();
        return (campo.value === reparto && !esche(p) && !p.querySelector('svg') && !eseguiti().length) || 'campo «' + campo.value + '», elementi iniettati ' + esche(p) + ', eseguito ' + eseguiti().join(',');
      });
    } finally {
      await chiudiSwal();
      statoMailVero();
      if (gisVero) google.accounts.oauth2.initCodeClient = gisVero;
      if (letto) { for (var k = 0; k < CHIAVI.length; k++) { try { await scriviImp(CHIAVI[k], prima[CHIAVI[k]]); } catch (e) {} } }
      try { await _sb.from('logs').delete().gte('ts', inizio).in('tipo', ['mail-mittente', 'mail-dimissioni', 'google-cassaforte']); } catch (e) {}
      var dopo = null;
      try { var rv = await fetch('/banco/cassaforte?modo=vera', { method: 'POST' }); dopo = rv.ok ? await rv.json() : null; } catch (e) {}
      if (typeof window._cassaforteAvvio === 'function') window._cassaforteAvvio();
      await prova(S, 'impostazioni e cassaforte del collaudo rimesse com\'erano', async function () {
        if (!letto) return 'le impostazioni di partenza non erano state lette';
        for (var j = 0; j < CHIAVI.length; j++) { if ((await imp(CHIAVI[j])) !== prima[CHIAVI[j]]) return CHIAVI[j] + ' non ripristinata'; }
        return (!!dopo && dopo.modo !== 'finta' && !dopo.veraInAttesa) || 'cassaforte: ' + JSON.stringify(dopo);
      });
      await prova(S, 'nessun errore JavaScript durante le prove della mail', function () {
        var nuovi = (window.__erroriBanco || []).slice(erroriPrima);
        return !nuovi.length || nuovi.slice(0, 4).join(' | ');
      });
    }
  }

  var SEZIONI = { ambiente: sezioneAmbiente, trak: sezioneTrak, xss: sezioneXss, mail: sezioneMail };

  window.__prove = async function (opz) {
    opz = opz || {};
    esiti = [];
    if (!window.CatalogoFinto) await carica('/trak-finto/catalogo.js');
    C = window.CatalogoFinto;
    var nomi = opz.sezioni || Object.keys(SEZIONI);
    for (var i = 0; i < nomi.length; i++) {
      if (!SEZIONI[nomi[i]]) { esiti.push({ sezione: nomi[i], nome: 'sezione sconosciuta', ok: false, dettaglio: '', ms: 0 }); continue; }
      try { await SEZIONI[nomi[i]](opz); }
      catch (e) { esiti.push({ sezione: nomi[i], nome: 'la sezione si è interrotta', ok: false, dettaglio: String((e && e.stack) || e).slice(0, 600), ms: 0 }); }
    }
    var casa = document.getElementById('proveCornici'); if (casa) casa.remove();
    var falliti = esiti.filter(function (e) { return !e.ok; });
    return { prove: esiti.length, superate: esiti.length - falliti.length, fallite: falliti.length, elenco: esiti };
  };
  console.log('[collaudo] prove pronte: await window.__prove()');
})();
