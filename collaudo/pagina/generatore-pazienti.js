// Generatore dei PAZIENTI FITTIZI del collaudo. Gira DENTRO l'app vera (banco
// di prova locale, collegato al database di collaudo) e riempie le schede
// passando dalle stesse funzioni che usano i colleghi: import terapia (parser
// del Pannello Terapia), problemi attivi, scale, isolamento/rianimazione,
// laboratorio e il salvataggio standard. Così l'HTML salvato è quello che
// scriverebbe l'app.
//
// I pazienti vengono dal catalogo condiviso col finto TrakCare
// (collaudo/trak-finto/catalogo.js), caricato da qui se manca.
//
// Uso, dalla pagina dell'app nel banco (http://localhost:8765/):
//   await window.__generaPazienti()                  riscrive i 24 pazienti
//   await window.__generaPazienti({ azzera: true })  prima riporta il reparto
//       allo stato iniziale: letti del catalogo, tutti vuoti, nessun esame
(function () {
  if (typeof AMBIENTE === 'undefined' || AMBIENTE !== 'collaudo') { console.error('generatore: consentito solo nel collaudo'); return; }

  function carica(src) {
    return new Promise(function (ok, ko) {
      var s = document.createElement('script');
      s.src = src; s.onload = ok; s.onerror = function () { ko(new Error('non riesco a caricare ' + src)); };
      document.head.appendChild(s);
    });
  }
  function scheda(letto) { return document.querySelector('.patient-card[data-bed="' + (window.CSS && CSS.escape ? CSS.escape(letto) : letto) + '"]'); }
  function salvaCard(letto, card) {
    return new Promise(function (resolve, reject) {
      eseguiSalvataggioLettoCompleto(letto, card);
      var t0 = Date.now();
      (function controlla() {
        if (!_saveInFlight[letto]) return resolve();
        if (Date.now() - t0 > 20000) return reject(new Error('salvataggio letto ' + letto + ' non concluso'));
        setTimeout(controlla, 120);
      })();
    });
  }
  function ridisegna() { return _sbGetPazienti().then(function (paz) { _applicaAggiornamentoCompleto(_renderCardsHtml(paz)); }); }

  // Reparto allo stato iniziale, con le funzioni dell'app: i letti del catalogo
  // (creati se mancano, con la loro tipologia) svuotati come dopo una
  // dimissione; i letti aggiunti durante le prove eliminati; nessun esame.
  async function azzera(C, rapporto) {
    var presenti = {};
    ((await _q(_sb.from('consegne').select('letto,tipologia_letto'))) || []).forEach(function (r) { presenti[r.letto] = r.tipologia_letto; });
    for (var i = 0; i < C.LETTI.length; i++) {
      var letto = C.LETTI[i][0], tipologia = C.LETTI[i][1];
      if (!(letto in presenti)) { await _sbAggiungiLetto(letto); presenti[letto] = 'STANDARD'; rapporto.creati++; }
      if (presenti[letto] !== tipologia) await _sbCambiaTipologiaALetto(letto, tipologia);
      await _sbDimettiLetto(letto);
      rapporto.svuotati++;
    }
    var canonici = C.LETTI.map(function (l) { return l[0]; });
    var estranei = Object.keys(presenti).filter(function (l) { return canonici.indexOf(l) < 0; });
    for (var k = 0; k < estranei.length; k++) { await _sbDimettiLetto(estranei[k]); await _sbEliminaLetto(estranei[k]); rapporto.eliminati++; }
    // le pulizie degli esami lanciate dalle dimissioni non sono attese dall'app:
    // si lascia che arrivino, poi si toglie ciò che resta
    await C.attendi(2500);
    await _q(_sb.from('lab_esami').delete().neq('letto', ''));
    window._labInfoLetti = {};
    await ridisegna();
    await _labBottoniAggiorna(true);
  }

  window.__generaPazienti = async function (opz) {
    opz = opz || {};
    if (!window.CatalogoFinto) await carica('/trak-finto/catalogo.js');
    var C = window.CatalogoFinto;
    var rapporto = { creati: 0, svuotati: 0, eliminati: 0, occupati: 0, lab: 0, errori: [] };
    if (opz.azzera) await azzera(C, rapporto);
    var scale = await _sbGetScaleValutazione();
    var piani = C.pianifica();

    // 1. laboratorio (prima: l'alambicco sulla scheda deriva da qui). Le pagine
    //    si leggono nell'ordine in cui le sfoglia il segnalibro.
    for (var i = 0; i < piani.length; i++) {
      var pl = piani[i];
      if (!pl.labNellaScheda) continue;
      try {
        var esami = [];
        C.labDi(pl, 0).forEach(function (pg) { esami.push.apply(esami, _parseLabRaw(pg)); });
        var m = _labMerge(null, esami);
        await _labSalva(pl.paziente.Letto, m.doc, pl.nome);
        rapporto.lab++;
      } catch (e) { rapporto.errori.push('lab ' + pl.paziente.Letto + ': ' + e.message); }
    }
    await _labBottoniAggiorna(true);

    // 2. schede
    for (var k = 0; k < piani.length; k++) {
      var p = piani[k], letto = p.paziente.Letto;
      try {
        var card = scheda(letto);
        if (!card) throw new Error('scheda non trovata');
        _aggiornaCardDaPaziente(card, p.paziente);
        // problemi attivi
        var pa = card.querySelector('[data-field="PianoTerapeutico"]');
        p.quadro.problemi.forEach(function (t) { pa.appendChild(_paRiga(t)); });
        _paApplica(pa);
        // terapia: righe del Pannello → parser vero → inserimento vero
        var r = _parseTerapiaTrak([{ t: 'Pannello terapia', righe: C.terapiaDi(p, 0) }]);
        _terapiaInserisci(card, r.corso, p.nome, Date.now() - (1 + (k % 5)) * 3600 * 1000, r.pregressi);
        if (k % 3 === 0) {
          var nt = card.querySelector('[data-field="NoteTerapia"]'), hdr = nt.querySelector('.note-hdr');
          if (hdr && hdr.nextElementSibling) hdr.nextElementSibling.innerHTML = 'Rivalutare terapia antalgica. Familiari riferiscono intolleranza a tramadolo.';
        }
        // scale: si usano i casi di verifica della scala stessa come dati d'ingresso
        p.quadro.scale.forEach(function (idScala, j) {
          var sc = scale.filter(function (s) { return s.id === idScala; })[0];
          if (!sc || !sc.definizione || !(sc.definizione.casi_verita || []).length) return;
          var casi = sc.definizione.casi_verita, caso = casi[(k + j) % casi.length];
          var ris = _calcolaScala(sc.definizione, caso.input);
          if (ris) _inserisciRisultatoScala(card, sc.id, _scalaFormattaRisultato(sc, ris));
        });
        if (p.quadro.isolamento) _isolamentoApplica(card, p.quadro.isolamento);
        if (p.quadro.rianimazione && typeof _rianimazioneApplica === 'function') _rianimazioneApplica(card, p.quadro.rianimazione);
        if (typeof window._labBottoniApplica === 'function') window._labBottoniApplica(card);
        await salvaCard(letto, card);
        rapporto.occupati++;
      } catch (e) { rapporto.errori.push('letto ' + letto + ': ' + e.message); }
      await C.attendi(150);
    }
    return rapporto;
  };
  console.log('[collaudo] generatore pazienti pronto: await window.__generaPazienti({ azzera: true })');
})();
