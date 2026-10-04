// Catalogo dei PAZIENTI FITTIZI del collaudo: anagrafiche, quadri clinici,
// terapie nel formato del «Pannello Terapia» di TrakCare e pannello degli esami.
// Lo usano sia il generatore che riempie le schede (collaudo/pagina/
// generatore-pazienti.js) sia il finto TrakCare (index.html in questa
// cartella): stessi semi, quindi stessi pazienti, stesse terapie, stessi esami.
//
// Il tempo del reparto finto è fermo a EPOCA: le schede generate descrivono i
// pazienti a quel giorno, qualunque sia il giorno in cui gira il generatore.
// Il finto TrakCare può «andare avanti» di qualche giorno (giorno simulato):
// compaiono prelievi nuovi e la terapia cambia, ma ciò che era già successo
// resta identico. Tutto inventato.
window.CatalogoFinto = (function () {
  var EPOCA = '2026-10-04';
  var SEME = 20261004;
  var GIORNI_AVANTI = 5;   // quanti giorni dopo l'epoca sa simulare il finto TrakCare

  // ── utilità ──────────────────────────────────────────────────────────
  function prng(seme) { var a = seme >>> 0; return function () { a |= 0; a = a + 0x6D2B79F5 | 0; var t = Math.imul(a ^ a >>> 15, 1 | a); t = t + Math.imul(t ^ t >>> 7, 61 | t) ^ t; return ((t ^ t >>> 14) >>> 0) / 4294967296; }; }
  // seme ricavato da più numeri: lo stesso valore per lo stesso (paziente, esame, giorno)
  function mescola(a, b, c, d) { return (Math.imul(a | 0, 73856093) ^ Math.imul(b | 0, 19349663) ^ Math.imul(c | 0, 83492791) ^ Math.imul(d | 0, 40503)) >>> 0; }
  var due = function (n) { return ('0' + n).slice(-2); };
  // giorno(0) è l'epoca; i giorni simulati hanno delta positivo
  var giorno = function (delta) { var p = EPOCA.split('-'); var d = new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]), 12, 0, 0, 0); d.setDate(d.getDate() + delta); return d; };
  var dmy = function (d) { return due(d.getDate()) + '/' + due(d.getMonth() + 1) + '/' + d.getFullYear(); };
  var dm = function (d) { return due(d.getDate()) + '/' + due(d.getMonth() + 1); };
  var iso = function (d) { return d.getFullYear() + '-' + due(d.getMonth() + 1) + '-' + due(d.getDate()); };
  var attendi = function (ms) { return new Promise(function (r) { setTimeout(r, ms); }); };
  var esc = function (s) { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); };

  // ── il reparto: gli stessi 28 letti della produzione ─────────────────
  var LETTI = [['1', 'APPOGGIO']];
  (function () {
    var i;
    for (i = 2; i <= 13; i++) LETTI.push([String(i), 'MEDICINA']);
    for (i = 14; i <= 19; i++) LETTI.push([String(i), 'SUB-INTENSIVA']);
    LETTI.push(['20', 'APPOGGIO'], ['21', 'SUB-INTENSIVA'], ['22', 'SUB-INTENSIVA'], ['23', 'APPOGGIO'], ['24', 'SUB-INTENSIVA'], ['25', 'SUB-INTENSIVA'],
      ['APPOGGIO 5M', 'STANDARD'], ['APPOGGIO AL 9E', 'STANDARD'], ['NOTE', 'STANDARD']);
  })();
  var LIBERI = ['5', '14', 'APPOGGIO 5M', 'NOTE'];   // 24 occupati, 3 liberi e la scheda NOTE

  // ── anagrafica inventata ─────────────────────────────────────────────
  var COGNOMI = ['Bianchi', 'Russo', 'Ferrari', 'Esposito', 'Romano', 'Colombo', 'Ricci', 'Marino', 'Greco', 'Bruno', 'Gallo', 'Conti', 'De Luca', 'Mancini', 'Costa', 'Giordano', 'Rizzo', 'Lombardi', 'Moretti', 'Barbieri', 'Fontana', 'Santoro', 'Mariani', 'Rinaldi', 'Caruso', 'Ferrara', 'Galli', 'Martini', 'Leone', 'Longo', "D'Angelo", 'De Santis'];
  var NOMI_M = ['Mario', 'Giuseppe', 'Antonio', 'Franco', 'Luigi', 'Paolo', 'Carlo', 'Sergio', 'Aldo', 'Bruno', 'Enzo', 'Pietro', 'Gianni', 'Renato'];
  var NOMI_F = ['Maria', 'Anna', 'Rosa', 'Giovanna', 'Teresa', 'Lucia', 'Carla', 'Elena', 'Angela', 'Franca', 'Paola', 'Silvana', 'Maria Grazia', 'Anna Rita'];
  var ALLERGIE = ['Nega allergie', 'Nega allergie', 'Nessuna nota', 'Penicillina (orticaria)', 'ASA', 'MdC iodato', 'Nega allergie', 'Lattice', 'Cefalosporine', 'Nega allergie'];
  var MEDICI = ['Dott. Verdi L.', 'Dott.ssa Neri P.', 'Dott. Gialli M.', 'Dott.ssa Rosa E.'];

  // ── terapia: righe nel formato del «Pannello Terapia» di TrakCare ─────
  var F = {
    piptazo: 'PIPERACILLINA SODICA/TAZOBACTAM SODICO — PIPERACILLINA TA AU*EV 10FL 4G+0,5G — 4G+0,5G polvere per soluzione per infusione DOSE 4.5 Grammi — In 30 min — ENDOVENOSA Infusione non Continua — 6-12-18-24 — Continuativo',
    ceftriaxone: 'CEFTRIAXONE SODICO — CEFTRIAXONE FIDIA*EV 1FL 2G — 2G polvere per soluzione per infusione DOSE 2 Grammi — ENDOVENOSA Infusione non Continua — 1 volta al giorno — Continuativo',
    meropenem: 'MEROPENEM TRIIDRATO — MEROPENEM AU*EV 10FL 1G — 1G polvere per soluzione iniettabile DOSE 1 Grammi — In 3 hrs — ENDOVENOSA Infusione non Continua — 8-16-24 — Continuativo',
    vancomicina: 'VANCOMICINA CLORIDRATO — VANCOMICINA HIK*EV 1FL 1G — 1G polvere per soluzione per infusione DOSE 1 Grammi — In 2 hrs — ENDOVENOSA Infusione non Continua — 8-20 — Continuativo',
    pantorc: 'PANTOPRAZOLO SODICO — PANTORC*EV 1FL 40MG — 40MG polvere per soluzione iniettabile DOSE 40 Milligrammi — ENDOVENOSA Bolo — 8 — Continuativo',
    pantorc2: 'PANTOPRAZOLO SODICO — PANTORC*EV 1FL 40MG — 40MG polvere per soluzione iniettabile DOSE 40 Milligrammi — ENDOVENOSA Bolo — 8-20 — Continuativo',
    lasix: 'FUROSEMIDE — LASIX*5F 20MG 2ML — 20MG/2ML soluzione iniettabile DOSE 20 Milligrammi — ENDOVENOSA Bolo — 8-16 — Continuativo',
    lasix40: 'FUROSEMIDE — LASIX*5F 20MG 2ML — 20MG/2ML soluzione iniettabile DOSE 40 Milligrammi — ENDOVENOSA Bolo — 8-14-20 — Continuativo',
    perfalgan: 'PARACETAMOLO — PERFALGAN*EV 12FL 100ML 10MG/ML — 10MG/ML soluzione per infusione DOSE 1 Grammi — ENDOVENOSA Infusione non Continua — 2 volte al giorno Al bisogno (Dose massima 3 Grammi in 24 ore) — Continuativo',
    isolyte: 'ELETTROLITI — ISOLYTE*SOL INF 500ML — soluzione per infusione DOSE 1000 Millilitri — In 24 hrs — ENDOVENOSA Infusione Continua — 1 volta al giorno — Continuativo',
    fisiologica: 'SODIO CLORURO — SODIO CLORURO BBU*0,9% 500ML — 0,9% soluzione per infusione DOSE 500 Millilitri — In 12 hrs — ENDOVENOSA Infusione Continua — 2 volte al giorno — Continuativo',
    urbason: 'METILPREDNISOLONE SODIO SUCCINATO — URBASON*EV 1F 40MG — 40MG polvere e solvente per soluzione iniettabile DOSE 40 Milligrammi — ENDOVENOSA Bolo — 8-20 — Continuativo',
    cordarone: 'AMIODARONE CLORIDRATO — CORDARONE*EV 5F 150MG 3ML — 150MG/3ML soluzione iniettabile DOSE 900 Milligrammi — In 24 hrs — ENDOVENOSA Infusione Continua — 1 volta al giorno — Continuativo',
    congescor: 'BISOPROLOLO FUMARATO — CONGESCOR*28CPR 2,5MG — 2,5MG compresse DOSE 1 Compressa — ORALE — 8 — Continuativo',
    eliquis: 'APIXABAN — ELIQUIS*60CPR RIV 5MG — 5MG compresse rivestite con film DOSE 1 Compressa — ORALE — 8-20 — Continuativo',
    torvast: 'ATORVASTATINA CALCICA — TORVAST*30CPR RIV 20MG — 20MG compresse rivestite con film DOSE 1 Compressa — ORALE — 22:00 — Continuativo',
    eutirox: 'LEVOTIROXINA SODICA — EUTIROX*50CPR 75MCG — 75MCG compresse DOSE 1 Compressa — ORALE — 7 — Continuativo',
    norvasc: 'AMLODIPINA BESILATO — NORVASC*28CPR 5MG — 5MG compresse DOSE 1 Compressa — ORALE — 8 — Continuativo',
    triatec: 'RAMIPRIL — TRIATEC*28CPR DIV 5MG — 5MG compresse divisibili DOSE 1 Compressa — ORALE — 8 — Continuativo',
    cardioasa: 'ACIDO ACETILSALICILICO — CARDIOASPIRIN*30CPR GASTR 100MG — 100MG compresse gastroresistenti DOSE 1 Compressa — ORALE — 13 — Continuativo',
    zyloric: 'ALLOPURINOLO — ZYLORIC*50CPR 100MG — 100MG compresse DOSE 1 Compressa — ORALE — 20 — Continuativo',
    metformina: 'METFORMINA CLORIDRATO — GLUCOPHAGE*30CPR RIV 500MG — 500MG compresse rivestite con film DOSE 1 Compressa — ORALE — 8-20 — Continuativo',
    coumadin: 'WARFARIN SODICO — COUMADIN*30CPR 5MG — 5MG compresse DOSE Conferma medica — ORALE — 18 — Continuativo',
    laevolac: 'LATTULOSIO — LAEVOLAC*SCIR 180ML 66,7% — 66,7G/100ML sciroppo DOSE 15 Millilitri — ORALE — 8-20 — Continuativo',
    seroquel: 'QUETIAPINA FUMARATO — SEROQUEL*30CPR RIV 25MG — 25MG compresse rivestite con film DOSE 1 Compressa — ORALE — 22:00 Al bisogno — Continuativo',
    deltacortene: 'PREDNISONE — DELTACORTENE*10CPR 25MG — 25MG compresse DOSE 1 Compressa — ORALE — 8 — Per 5 Giorni',
    clexane: 'ENOXAPARINA SODICA — CLEXANE*6SIR 4000UI 0,4ML — 4000UI/0,4ML soluzione iniettabile DOSE 4000 Unita — SOTTOCUTE — 20 — Continuativo',
    clexane2: 'ENOXAPARINA SODICA — CLEXANE*6SIR 6000UI 0,6ML — 6000UI/0,6ML soluzione iniettabile DOSE 6000 Unita — SOTTOCUTE — 8-20 — Continuativo',
    humalog: 'INSULINA LISPRO — HUMALOG*KWIKPEN 100UI/ML — 100UI/ML soluzione iniettabile DOSE 6 Unita — SOTTOCUTE — 8-12-18 Al bisogno (Dose massima 30 Unita in 24 ore) — Continuativo',
    lantus: 'INSULINA GLARGINE — LANTUS*SOLOSTAR 100UI/ML — 100UI/ML soluzione iniettabile DOSE 14 Unita — SOTTOCUTE — 22:00 — Continuativo',
    breva: 'IPRATROPIO BROMURO/SALBUTAMOLO SOLFATO — BREVA*SOL NEBUL 15ML — 0,375%+0,075% soluzione da nebulizzare DOSE 10 GOCCE — AEROSOLICA — 8-14-20 — Continuativo',
    aircort: 'BUDESONIDE — AIRCORT*20FL 2ML 0,5MG/ML — 0,5MG/ML sospensione da nebulizzare DOSE 1 Milligrammi — AEROSOLICA — 8-20 — Continuativo',
    tranex: 'ACIDO TRANEXAMICO — TRANEX*6F 500MG 5ML — 500MG/5ML soluzione iniettabile DOSE 500 Milligrammi — ENDOVENOSA Bolo — Immediata',
    plasil: 'METOCLOPRAMIDE CLORIDRATO — PLASIL*5F 10MG 2ML — 10MG/2ML soluzione iniettabile DOSE 10 Milligrammi — ENDOVENOSA Bolo — Immediata',
    lasixSubito: 'FUROSEMIDE — LASIX*5F 20MG 2ML — 20MG/2ML soluzione iniettabile DOSE 20 Milligrammi — ENDOVENOSA Bolo — Immediata',
    // righe che l'importazione deve SCARTARE: un ordine di reparto scritto come
    // parafarmaco e una prescrizione chiusa per errore di inserimento
    ordineReparto: 'PARAFARMACO — PARAFARMACO*GENERICO — DOSE 1 Unita — ORALE — 1 volta al giorno — Continuativo monitoraggio diuresi delle 24 ore',
    perErrore: 'TRAMADOLO CLORIDRATO — CONTRAMAL*OS GTT 10ML 100MG/ML — 100MG/ML gocce orali DOSE 20 GOCCE — ORALE — 8-20 — Continuativo',
  };
  var FARMACI_DI_CASA = ['zyloric', 'eutirox', 'torvast', 'norvasc', 'laevolac', 'cardioasa', 'deltacortene'];
  // evoluzione nei giorni simulati: dal giorno +1 si sospende il primo di questi
  // che il paziente ha in corso e si dà una dose immediata; dal +2 se ne aggiunge uno
  var SOSPENDIBILI = ['isolyte', 'fisiologica', 'perfalgan', 'pantorc', 'pantorc2'];
  var AGGIUNGIBILI = ['laevolac', 'seroquel', 'zyloric', 'norvasc'];
  var FINESTRA = 3;   // il Pannello mostra da 3 giorni prima a 3 giorni dopo
  function finestra(s) { var gg = []; for (var g = (s || 0) - FINESTRA; g <= (s || 0) + FINESTRA; g++) gg.push(dmy(giorno(g))); return gg; }

  // Gli attrezzi con cui i quadri clinici descrivono la terapia vista nel
  // giorno simulato «s». Le date sono in giorni dall'epoca (negative = prima).
  function attrezzi(p, s, sospese) {
    function date(da, a) { var gg = []; for (var i = Math.max(da, s - FINESTRA); i <= Math.min(a, s + FINESTRA); i++) gg.push(dmy(giorno(i))); return gg.join(' '); }
    function medico(chiave) { var h = 0; for (var i = 0; i < chiave.length; i++) h = (h * 31 + chiave.charCodeAt(i)) >>> 0; return MEDICI[(h + p.indice) % MEDICI.length]; }
    function riga(chiave, testo, giorni, info, stato, statoInfo) {
      return { chiave: chiave, Prestazione: testo, Gruppo: /ENDOVENOSA/.test(testo) ? 'ENDOVENA' : 'TERAPIA', Giorni: giorni, Info: info, Stato: stato, StatoInfo: statoInfo, _r: testo };
    }
    var T = {
      // in corso dal giorno «da»: la griglia mostra anche i giorni programmati
      inCorso: function (chiave, da, nota) {
        if (da > s) return null;                                   // non ancora prescritto
        if (sospese[chiave] != null && sospese[chiave] <= s) return T.chiuso(chiave, da, sospese[chiave], 'Sospeso');
        var t = F[chiave] + (nota ? ' ' + nota : '');
        return riga(chiave, t, date(da, s + FINESTRA), 'Data inizio ' + dmy(giorno(da)) + ' ' + medico(chiave), '', '');
      },
      // dal giorno «da» al giorno «a»: finché «a» non è passato è ancora in corso
      chiuso: function (chiave, da, a, stato, motivo) {
        if (da > s) return null;
        if (a > s) return T.inCorso(chiave, da);
        var m = medico(chiave);
        return riga(chiave, F[chiave] + ' ' + stato, date(da, a),
          'Data inizio ' + dmy(giorno(da)) + (stato === 'Completato' ? ' Data fine ' + dmy(giorno(a)) : '') + ' ' + m, stato,
          stato === 'Completato' ? 'Prescrizione eseguita completamente' : 'Prescrizione sospesa da ' + m + ' il ' + dmy(giorno(a)) + ' 10:30' + (motivo ? ' Reason: ' + motivo : ''));
      },
      // dosi una tantum nei giorni indicati
      immediata: function (chiave, quando) {
        var fatte = quando.filter(function (g) { return g <= s; });
        if (!fatte.length) return null;
        return riga(chiave, F[chiave], fatte.filter(function (g) { return g >= s - FINESTRA; }).map(function (g) { return dmy(giorno(g)); }).join(' '), 'Data inizio ' + dmy(giorno(fatte[0])), '', '');
      },
    };
    return T;
  }

  // ── laboratorio: pannello base + scostamenti per quadro clinico ───────
  var EMO = 'ESAME EMOCROMOCIT. E MORFOLOGICO COMPLETO (con FORMULA) -> Sangue';
  var EGA = 'EMOGASANALISI ARTERIOSA -> Sangue', URI = 'ESAME CHIMICO FISICO E MICROSCOPICO URINE -> Urine';
  // [gruppo, esame, minimo, massimo, unità, valore tipico, decimali, (esiti a parole)]
  var PANNELLO = [
    [EMO, 'Leucociti', 4.0, 10.0, '10^9/L', 7.2, 2], [EMO, 'Eritrociti', 4.2, 5.8, '10^12/L', 4.6, 2], [EMO, 'Emoglobina', 12.0, 16.0, 'g/dL', 13.4, 1],
    [EMO, 'Ematocrito', 36, 48, '%', 40.5, 1], [EMO, 'MCV', 80, 98, 'fL', 88.0, 1], [EMO, 'Piastrine', 150, 450, '10^9/L', 245, 0],
    [EMO, 'Neutrofili %', 40, 75, '%', 62.0, 1], [EMO, 'Linfociti %', 20, 45, '%', 27.0, 1],
    ['SODIO -> Sangue', 'Sodio', 136, 145, 'mmol/L', 140, 0], ['POTASSIO -> Sangue', 'Potassio', 3.5, 5.1, 'mmol/L', 4.2, 1],
    ['CLORURO -> Sangue', 'Cloruro', 98, 107, 'mmol/L', 102, 0], ['CALCIO TOTALE -> Sangue', 'Calcio', 8.6, 10.2, 'mg/dL', 9.2, 1],
    ['TEMPO DI PROTROMBINA (PT) -> Sangue', 'INR', 0.8, 1.2, '', 1.05, 2], ['TEMPO DI TROMBOPLASTINA PARZIALE (aPTT) -> Sangue', 'aPTT', 25, 38, 'sec', 31.0, 1],
    ['FIBRINOGENO -> Sangue', 'Fibrinogeno', 200, 400, 'mg/dL', 330, 0],
    ['CREATININA -> Sangue', 'Creatinina', 0.6, 1.2, 'mg/dL', 0.9, 2], ['UREA -> Sangue', 'Azotemia', 17, 49, 'mg/dL', 36, 0],
    ['BILIRUBINA TOTALE E FRAZIONATA -> Sangue', 'Bilirubina totale', 0.2, 1.2, 'mg/dL', 0.7, 2],
    ['ASPARTATO AMINOTRANSFERASI (AST) (GOT) -> Sangue', 'AST', 5, 40, 'U/L', 24, 0], ['ALANINA AMINOTRANSFERASI (ALT) (GPT) -> Sangue', 'ALT', 5, 41, 'U/L', 22, 0],
    ['GAMMA GLUTAMIL TRANSPEPTIDASI -> Sangue', 'Gamma GT', 8, 61, 'U/L', 30, 0],
    ['SIDEREMIA -> Sangue', 'Ferro', 65, 175, 'gamma/dL', 80, 0], ['FERRITINA -> Sangue', 'Ferritina', 30, 400, 'ng/mL', 150, 0],
    ['GLUCOSIO -> Sangue', 'Glucosio', 70, 110, 'mg/dL', 96, 0], ['PROTEINA C REATTIVA -> Sangue', 'Proteina C reattiva', 0, 5, 'mg/L', 3.0, 1],
    ['PROCALCITONINA -> Sangue', 'Procalcitonina', 0, 0.5, 'ng/mL', 0.08, 2], ['ALBUMINA -> Sangue', 'Albumina', 3.5, 5.2, 'g/dL', 3.9, 1],
    [EMO, 'MCH', 27, 32, 'pg', 29.4, 1], [EMO, 'MCHC', 32, 36, 'g/dL', 33.6, 1], [EMO, 'RDW', 11.5, 14.5, '%', 13.4, 1], [EMO, 'MPV', 7.5, 11.5, 'fL', 9.2, 1],
    [EMO, 'Monociti %', 2, 10, '%', 7.5, 1], [EMO, 'Eosinofili %', 0, 6, '%', 2.4, 1], [EMO, 'Basofili %', 0, 1.5, '%', 0.5, 1],
    [EMO, 'Neutrofili #', 1.8, 7.5, '10^9/L', 4.5, 2], [EMO, 'Linfociti #', 1.0, 4.0, '10^9/L', 1.9, 2], [EMO, 'Monociti #', 0.2, 1.0, '10^9/L', 0.55, 2],
    [EMO, 'Eosinofili #', 0, 0.5, '10^9/L', 0.17, 2], [EMO, 'Basofili #', 0, 0.1, '10^9/L', 0.03, 2],
    ['MAGNESIO -> Sangue', 'Magnesio', 1.6, 2.6, 'mg/dL', 2.0, 1], ['FOSFORO -> Sangue', 'Fosforo', 2.5, 4.5, 'mg/dL', 3.4, 1],
    ['TEMPO DI PROTROMBINA (PT) -> Sangue', 'PT %', 70, 120, '%', 92, 0], ['TEMPO DI PROTROMBINA (PT) -> Sangue', 'PT sec', 10, 14, 'sec', 11.8, 1],
    ['TEMPO DI TROMBOPLASTINA PARZIALE (aPTT) -> Sangue', 'aPTT ratio', 0.8, 1.2, '', 1.02, 2], ['D-DIMERO -> Sangue', 'D-Dimero', 0, 500, 'ng/mL', 310, 0],
    ['ANTITROMBINA -> Sangue', 'Antitrombina', 80, 120, '%', 96, 0],
    ['CREATININA -> Sangue', 'eGFR (CKD-EPI)', 60, 140, 'mL/min', 82, 0], ['ACIDO URICO -> Sangue', 'Acido urico', 2.4, 7.0, 'mg/dL', 5.1, 1],
    ['BILIRUBINA TOTALE E FRAZIONATA -> Sangue', 'Bilirubina diretta', 0, 0.3, 'mg/dL', 0.18, 2], ['BILIRUBINA TOTALE E FRAZIONATA -> Sangue', 'Bilirubina indiretta', 0.1, 0.9, 'mg/dL', 0.5, 2],
    ['FOSFATASI ALCALINA -> Sangue', 'Fosfatasi alcalina', 40, 130, 'U/L', 78, 0], ['LATTATO DEIDROGENASI (LDH) -> Sangue', 'LDH', 135, 225, 'U/L', 188, 0],
    ['AMILASI -> Sangue', 'Amilasi', 28, 100, 'U/L', 61, 0], ['LIPASI -> Sangue', 'Lipasi', 13, 60, 'U/L', 34, 0],
    ['CREATINCHINASI (CPK) -> Sangue', 'CPK', 30, 200, 'U/L', 96, 0], ['COLINESTERASI -> Sangue', 'Colinesterasi', 5300, 12900, 'U/L', 7800, 0],
    ['TRANSFERRINA -> Sangue', 'Transferrina', 200, 360, 'mg/dL', 248, 0], ['TRANSFERRINA -> Sangue', 'Saturazione transferrina', 20, 50, '%', 26, 0],
    ['VITAMINA B12 -> Sangue', 'Vitamina B12', 191, 663, 'pg/mL', 402, 0], ['FOLATO -> Sangue', 'Folati', 3.9, 26.8, 'ng/mL', 8.2, 1],
    ['PROTEINE TOTALI -> Sangue', 'Proteine totali', 6.4, 8.3, 'g/dL', 6.8, 1],
    ['ELETTROFORESI DELLE SIEROPROTEINE -> Sangue', 'Albumina %', 55.8, 66.1, '%', 58.4, 1], ['ELETTROFORESI DELLE SIEROPROTEINE -> Sangue', 'Alfa 1 %', 2.9, 4.9, '%', 4.1, 1],
    ['ELETTROFORESI DELLE SIEROPROTEINE -> Sangue', 'Alfa 2 %', 7.1, 11.8, '%', 10.2, 1], ['ELETTROFORESI DELLE SIEROPROTEINE -> Sangue', 'Beta 1 %', 4.7, 7.2, '%', 6.0, 1],
    ['ELETTROFORESI DELLE SIEROPROTEINE -> Sangue', 'Beta 2 %', 3.2, 6.5, '%', 4.8, 1], ['ELETTROFORESI DELLE SIEROPROTEINE -> Sangue', 'Gamma %', 11.1, 18.8, '%', 16.5, 1],
    ['TROPONINA I ALTA SENSIBILITA -> Sangue', 'Troponina I hs', 0, 34, 'ng/L', 12, 0], ['PEPTIDE NATRIURETICO (NT-proBNP) -> Sangue', 'NT-proBNP', 0, 300, 'pg/mL', 210, 0],
    ['CK-MB MASSA -> Sangue', 'CK-MB massa', 0, 5, 'ng/mL', 1.8, 1], ['MIOGLOBINA -> Sangue', 'Mioglobina', 25, 72, 'ng/mL', 41, 0],
    ['EMOGLOBINA GLICATA -> Sangue', 'Emoglobina glicata', 20, 42, 'mmol/mol', 39, 0], ['COLESTEROLO TOTALE -> Sangue', 'Colesterolo totale', 0, 200, 'mg/dL', 176, 0],
    ['COLESTEROLO HDL -> Sangue', 'Colesterolo HDL', 40, 100, 'mg/dL', 48, 0], ['COLESTEROLO LDL -> Sangue', 'Colesterolo LDL', 0, 116, 'mg/dL', 104, 0],
    ['TRIGLICERIDI -> Sangue', 'Trigliceridi', 0, 150, 'mg/dL', 128, 0],
    ['TIREOTROPINA (TSH) -> Sangue', 'TSH', 0.35, 4.94, 'mcUI/mL', 2.1, 2], ['TIROXINA LIBERA (FT4) -> Sangue', 'FT4', 9.0, 19.0, 'pmol/L', 13.2, 1], ['TRIIODOTIRONINA LIBERA (FT3) -> Sangue', 'FT3', 2.6, 5.7, 'pmol/L', 4.0, 1],
    ['VELOCITA DI ERITROSEDIMENTAZIONE (VES) -> Sangue', 'VES', 0, 20, 'mm/h', 16, 0], ['LATTATO -> Sangue', 'Lattato', 0.5, 2.2, 'mmol/L', 1.3, 1], ['AMMONIO -> Sangue', 'Ammonio', 16, 60, 'mcmol/L', 38, 0],
    [EGA, 'pH', 7.35, 7.45, '', 7.41, 2], [EGA, 'pCO2', 35, 45, 'mmHg', 39, 0], [EGA, 'pO2', 80, 100, 'mmHg', 86, 0], [EGA, 'HCO3-', 22, 26, 'mmol/L', 24.1, 1],
    [EGA, 'BE', -2, 2, 'mmol/L', 0.4, 1], [EGA, 'Saturazione O2', 95, 99, '%', 96.5, 1], [EGA, 'Lattati (EGA)', 0.5, 1.6, 'mmol/L', 1.1, 1], [EGA, 'Glucosio (EGA)', 70, 110, 'mg/dL', 104, 0],
    [URI, 'pH urinario', 5.0, 7.5, '', 6.0, 1], [URI, 'Peso specifico', 1.005, 1.030, '', 1.018, 3],
    [URI, 'Proteine urinarie', 0, 0, '', 0, 0, ['Assenti', 'Assenti', 'Tracce', '30 mg/dL']], [URI, 'Glucosio urinario', 0, 0, '', 0, 0, ['Assente', 'Assente', 'Presente']],
    [URI, 'Emoglobina urinaria', 0, 0, '', 0, 0, ['Assente', 'Tracce', 'Presente']], [URI, 'Esterasi leucocitaria', 0, 0, '', 0, 0, ['Negativa', 'Negativa', 'Positiva']],
    [URI, 'Nitriti', 0, 0, '', 0, 0, ['Negativi', 'Negativi', 'Positivi']], [URI, 'Corpi chetonici', 0, 0, '', 0, 0, ['Assenti', 'Assenti', 'Presenti']],
    [URI, 'Urobilinogeno', 0, 0, '', 0, 0, ['Nei limiti della norma']], [URI, 'Sedimento', 0, 0, '', 0, 0, ['Nei limiti della norma', 'Rari leucociti', 'Numerosi leucociti e batteri']],
  ];
  // esami «di approfondimento»: non a ogni prelievo
  var RARO = PANNELLO.map(function (e) { return e[0] !== EMO && !/^(Sodio|Potassio|Cloruro|Calcio|Creatinina|Azotemia|eGFR|Glucosio$|Proteina C reattiva|INR|aPTT$|AST|ALT|Bilirubina totale)/.test(e[1]); });
  // quadro: { esame: [valore all'ingresso, valore il giorno dell'epoca] }
  var QUADRI = {
    sepsi: { Leucociti: [19.4, 11.2], 'Neutrofili %': [88, 76], 'Linfociti %': [6, 14], 'Proteina C reattiva': [246, 71], Procalcitonina: [14.6, 1.2], Creatinina: [1.9, 1.2], Azotemia: [92, 55], Piastrine: [128, 190], Albumina: [2.8, 3.0] },
    scompenso: { Creatinina: [1.5, 1.3], Azotemia: [68, 58], Sodio: [132, 135], Potassio: [4.9, 3.6], Emoglobina: [11.2, 11.4], Albumina: [3.2, 3.3] },
    anemia: { Emoglobina: [6.9, 8.8], Eritrociti: [2.6, 3.2], Ematocrito: [21.5, 27.0], MCV: [74, 76], Ferro: [18, 22], Ferritina: [8, 12], Azotemia: [78, 44] },
    bpco: { Leucociti: [13.1, 9.4], 'Proteina C reattiva': [58, 14], Emoglobina: [15.9, 15.6], Ematocrito: [49.5, 48.6] },
    aki: { Creatinina: [4.3, 2.1], Azotemia: [168, 96], Potassio: [6.1, 4.8], Sodio: [131, 137], Calcio: [8.1, 8.5], Cloruro: [96, 101] },
    epatico: { 'Bilirubina totale': [5.8, 3.1], AST: [212, 96], ALT: [164, 88], 'Gamma GT': [488, 301], INR: [1.6, 1.4], Albumina: [2.6, 2.7], Piastrine: [86, 92], Sodio: [130, 133] },
    tep: { 'Proteina C reattiva': [34, 12], Leucociti: [11.4, 8.8] },
    diabete: { Glucosio: [486, 162], Sodio: [129, 138], Potassio: [5.6, 4.1], Creatinina: [1.6, 1.0], Azotemia: [74, 42] },
    polmonite: { Leucociti: [16.8, 9.9], 'Neutrofili %': [84, 70], 'Proteina C reattiva': [188, 42], Procalcitonina: [3.4, 0.4], Sodio: [133, 137] },
    normale: {},
  };
  var PER_PAGINA = 5, MAX_PRELIEVI = 12;

  // ── quadri clinici (tutto inventato) ─────────────────────────────────
  // Ogni quadro: diagnosi, problemi attivi, terapia, laboratorio, diaria, da fare.
  // terapia(T, da): «da» è il giorno del ricovero, in giorni dall'epoca.
  var b = function (t) { return '<b>' + t + '</b>'; };
  var QUADRI_CLINICI = [
    { id: 'polmonite', diagnosi: 'Polmonite comunitaria basale dx con insufficienza respiratoria tipo 1', lab: 'polmonite', o2: ['O2 2 lt/min CN', 'VMK 40%'], vitto: 'Leggero',
      problemi: ['Polmonite comunitaria basale dx', 'Insufficienza respiratoria ipossiemica', 'Ipertensione arteriosa'],
      terapia: function (T, da) { return [T.inCorso('ceftriaxone', da), T.inCorso('perfalgan', da), T.inCorso('pantorc', da), T.inCorso('isolyte', da), T.inCorso('norvasc', da), T.inCorso('clexane', da), T.inCorso('breva', da)]; },
      scale: ['curb-65', 'news2'], colturali: ['Emocolture ' + dm(giorno(-3)) + ': negative a 72h', 'Ag urinari Legionella/Pneumococco: negativi'],
      apr: 'Ipertensione arteriosa in terapia. Ex fumatore (30 p/y).', app: 'Da 4 giorni febbre fino a 39°C, tosse produttiva e dispnea ingravescente. In PS SpO2 88% in AA, Rx torace: addensamento basale dx.',
      diario: ['Apiretico da 24h, eupnoico in O2 a bassi flussi. MV ridotto alla base dx con crepitii. Indici di flogosi in riduzione.', 'Stabile. Si scala O2. Prosegue terapia antibiotica.'],
      dafare: ['Rx torace di controllo', 'Scalare O2 se SpO2 > 94%'], attesa: ['Consulenza pneumologica'], eseguiti: ['Rx torace all\'ingresso', 'EGA'] },
    { id: 'scompenso', diagnosi: 'Scompenso cardiaco acuto in cardiopatia ischemica cronica (FE 35%)', lab: 'scompenso', o2: ['O2 3 lt/min CN', 'AA'], vitto: 'Iposodico',
      problemi: ['Scompenso cardiaco acuto', 'Cardiopatia ischemica cronica', 'IRC stadio 3a', 'Diabete mellito tipo 2'],
      terapia: function (T, da) { return [T.inCorso('lasix40', da), T.inCorso('pantorc', da), T.inCorso('congescor', da), T.inCorso('triatec', da), T.inCorso('cardioasa', da), T.inCorso('torvast', da), T.inCorso('clexane', da), T.inCorso('humalog', da), T.inCorso('lantus', da), T.chiuso('metformina', da, da + 1, 'Sospeso')]; },
      scale: ['padua'], colturali: [],
      apr: 'Cardiopatia ischemica (PTCA+DES su IVA nel 2019), DM2, IRC stadio 3a, dislipidemia.', app: 'Dispnea ingravescente da una settimana, ortopnea, edemi declivi. In PS: PA 160/95, rantoli bibasali, NT-proBNP elevato.',
      diario: ['Bilancio idrico negativo (-1400 ml). Edemi in riduzione, eupnoico a riposo. Peso -1,8 kg.', 'Diuresi valida. K 3,6: si integra per os. Creatinina stabile.'],
      dafare: ['Peso quotidiano e bilancio idrico', 'Controllo elettroliti domattina'], attesa: ['Ecocardiogramma', 'Consulenza cardiologica'], eseguiti: ['ECG', 'Rx torace'] },
    { id: 'emorragia', diagnosi: 'Emorragia digestiva superiore da ulcera duodenale (Forrest IIb), anemizzazione', lab: 'anemia', o2: ['AA'], vitto: 'Digiuno',
      problemi: ['Melena da ulcera duodenale', 'Anemia acuta post-emorragica', 'FA permanente (TAO sospesa)'],
      terapia: function (T, da) { return [T.inCorso('pantorc2', da), T.inCorso('isolyte', da), T.inCorso('congescor', da), T.chiuso('eliquis', da - 400, da, 'Sospeso'), T.immediata('tranex', [da]), T.immediata('plasil', [da, da + 1])]; },
      scale: ['has-bled', 'cha2ds2-vasc'], colturali: [],
      apr: 'FA permanente in apixaban, ipertensione arteriosa, artrosi (FANS al bisogno).', app: 'Melena da 2 giorni, astenia marcata, un episodio lipotimico. In PS Hb 6,9 g/dL: trasfuse 2 UEC. EGDS: ulcera duodenale Forrest IIb trattata con clip.',
      diario: ['Non ulteriori episodi di melena. Hb stabile dopo trasfusione. Emodinamica stabile.', 'Alvo aperto a feci normocromiche. Si riprende alimentazione semiliquida da domani se Hb stabile.'],
      dafare: ['Emocromo ogni 12h', 'Rivalutare ripresa anticoagulante'], attesa: ['EGDS di controllo'], eseguiti: ['EGDS con clip', 'Trasfusione 2 UEC'] },
    { id: 'bpco', diagnosi: 'Riacutizzazione di BPCO con insufficienza respiratoria ipercapnica', lab: 'bpco', o2: ['VMK 28%', 'O2 1 lt/min CN'], vitto: 'Libero',
      problemi: ['BPCO riacutizzata', 'Insufficienza respiratoria tipo 2', 'Tabagismo attivo'],
      terapia: function (T, da) { return [T.inCorso('urbason', da), T.inCorso('breva', da), T.inCorso('aircort', da), T.inCorso('pantorc', da), T.inCorso('clexane', da), T.inCorso('seroquel', da), T.chiuso('ceftriaxone', da, da + 4, 'Completato')]; },
      scale: ['news2'], colturali: ['Escreato ' + dm(giorno(-2)) + ': flora mista'],
      apr: 'BPCO GOLD 3 in LABA/LAMA, fumatore attivo (50 p/y), OSAS non trattata.', app: 'Peggioramento della dispnea da 3 giorni con aumento dell\'espettorato. EGA all\'ingresso: pH 7,31, pCO2 62.',
      diario: ['EGA di controllo: pH 7,38, pCO2 54. Meno broncospasmo. Target SpO2 88-92%.', 'In miglioramento. Si passa a steroide per os.'],
      dafare: ['EGA di controllo', 'Counselling antitabagico'], attesa: ['Spirometria alla dimissione'], eseguiti: ['Rx torace', 'EGA seriati'] },
    { id: 'urosepsi', diagnosi: 'Sepsi a partenza urinaria da E. coli ESBL+', lab: 'sepsi', o2: ['AA', 'O2 2 lt/min CN'], vitto: 'Leggero', isolamento: 'Contatto — E. coli ESBL+',
      problemi: ['Sepsi urinaria (E. coli ESBL+)', 'Insufficienza renale acuta su cronica', 'Demenza vascolare'],
      terapia: function (T, da) { return [T.inCorso('meropenem', da + 1), T.inCorso('isolyte', da), T.inCorso('perfalgan', da), T.inCorso('pantorc', da), T.inCorso('clexane', da), T.inCorso('seroquel', da), T.chiuso('piptazo', da, da + 1, 'Sospeso')]; },
      scale: ['qsofa', 'news2'], colturali: ['Urinocoltura ' + dm(giorno(-4)) + ': E. coli ESBL+ (S a meropenem, R a pip/tazo)', 'Emocolture ' + dm(giorno(-4)) + ': E. coli ESBL+ in 2/2 set'],
      apr: 'Demenza vascolare, IRC stadio 3b, portatrice di catetere vescicale a permanenza.', app: 'Febbre con brivido, stato confusionale acuto e ipotensione. In PS lattati 3,8, avviata fluidoterapia e antibiotico empirico, poi ottimizzato su antibiogramma.',
      diario: ['Apiretica da 48h. PA 125/70 senza supporto. Diuresi 1600 ml/24h. Creatinina in riduzione.', 'Vigile, orientata nella persona. Sostituito catetere vescicale.'],
      dafare: ['Emocolture di controllo', 'Durata meropenem: 7 giorni'], attesa: ['Ecografia renale'], eseguiti: ['Urinocoltura', 'Emocolture x2', 'Sostituzione CV'] },
    { id: 'fa', diagnosi: 'Fibrillazione atriale ad alta risposta ventricolare di nuova diagnosi', lab: 'normale', o2: ['AA'], vitto: 'Libero',
      problemi: ['FA ad alta risposta ventricolare', 'Ipertiroidismo subclinico da approfondire'],
      terapia: function (T, da) { return [T.inCorso('cordarone', da), T.inCorso('congescor', da), T.inCorso('clexane2', da), T.inCorso('pantorc', da), T.inCorso('triatec', da)]; },
      scale: ['cha2ds2-vasc', 'has-bled'], colturali: [],
      apr: 'Ipertensione arteriosa. Nessun precedente cardiologico.', app: 'Cardiopalmo e dispnea da sforzo da 24h. ECG: FA a FC 150/min. Avviato controllo della frequenza.',
      diario: ['FC 85-95/min, asintomatico. Ecocardiogramma: FE conservata, atrio sx lievemente dilatato.', 'Ritmo sinusale dopo carico di amiodarone. Si imposta anticoagulante orale.'],
      dafare: ['ECG quotidiano', 'TSH, fT3, fT4'], attesa: ['Ecocardiogramma'], eseguiti: ['ECG', 'Troponina seriata negativa'] },
    { id: 'aki', diagnosi: 'Insufficienza renale acuta prerenale con iperkaliemia in corso di gastroenterite', lab: 'aki', o2: ['AA'], vitto: 'Leggero',
      problemi: ['IRA prerenale (KDIGO 3)', 'Iperkaliemia', 'Disidratazione'],
      terapia: function (T, da) { return [T.inCorso('fisiologica', da), T.inCorso('isolyte', da), T.inCorso('pantorc', da), T.inCorso('clexane', da), T.chiuso('triatec', da - 300, da, 'Sospeso'), T.chiuso('lasix', da - 300, da, 'Sospeso')]; },
      scale: [], colturali: ['Coprocoltura ' + dm(giorno(-2)) + ': in corso'],
      apr: 'Ipertensione arteriosa in ACE-inibitore e diuretico, DM2.', app: 'Diarrea profusa e vomito da 5 giorni con ridotto introito. In PS creatinina 4,3 e K 6,1 con alterazioni ECG: trattata con insulina-glucosio e calcio gluconato.',
      diario: ['Diuresi ripresa (2100 ml/24h). Creatinina in rapido calo, K normalizzato. Alvo: 2 scariche.', 'Si riduce idratazione ev. Non necessità dialitica.'],
      dafare: ['Elettroliti e funzione renale ogni 12h', 'Bilancio idrico'], attesa: ['Ecografia reni e vie urinarie'], eseguiti: ['ECG', 'EGA venoso'] },
    { id: 'cirrosi', diagnosi: 'Cirrosi epatica alcol-correlata scompensata (ascite ed encefalopatia grado 2)', lab: 'epatico', o2: ['AA'], vitto: 'Iposodico',
      problemi: ['Cirrosi scompensata (Child C)', 'Encefalopatia epatica', 'Ascite tesa', 'Piastrinopenia'],
      terapia: function (T, da) { return [T.inCorso('laevolac', da), T.inCorso('lasix', da), T.inCorso('pantorc', da), T.inCorso('isolyte', da), T.inCorso('ceftriaxone', da)]; },
      scale: ['gcs'], colturali: ['Liquido ascitico ' + dm(giorno(-3)) + ': PMN 120/mm3, coltura negativa'],
      apr: 'Potus attivo, cirrosi nota da 2 anni, varici esofagee F2 in legatura.', app: 'Stato confusionale e aumento della circonferenza addominale. Paracentesi evacuativa di 5 litri con albumina.',
      diario: ['Più vigile, flapping assente. 3 evacuazioni/die con lattulosio. Addome trattabile.', 'Stabile. Si programma EGDS di controllo.'],
      dafare: ['Bilancio idrico e peso', 'Rivalutazione stato neurologico'], attesa: ['Consulenza epatologica', 'EGDS'], eseguiti: ['Paracentesi 5 L', 'Ecografia addome'] },
    { id: 'tep', diagnosi: 'Embolia polmonare a rischio intermedio-basso con TVP femoro-poplitea sx', lab: 'tep', o2: ['O2 2 lt/min CN', 'AA'], vitto: 'Libero',
      problemi: ['Embolia polmonare bilaterale', 'TVP femoro-poplitea sx'],
      terapia: function (T, da) { return [T.inCorso('clexane2', da), T.inCorso('perfalgan', da), T.inCorso('pantorc', da), T.inCorso('eutirox', da)]; },
      scale: ['wells-pe', 'pesi'], colturali: [],
      apr: 'Ipotiroidismo in terapia sostitutiva. Recente intervento ortopedico (protesi d\'anca 3 settimane fa).', app: 'Dispnea improvvisa e dolore toracico pleuritico. AngioTC: difetti di riempimento bilaterali. Troponina negativa, VD non dilatato.',
      diario: ['Eupnoica, SpO2 96% in AA. Non segni di sovraccarico destro. Si mobilizza.', 'Si programma passaggio ad anticoagulante orale diretto.'],
      dafare: ['Passaggio a DOAC', 'Calze elastiche'], attesa: ['Ecocardiogramma'], eseguiti: ['AngioTC torace', 'Ecocolordoppler arti inferiori'] },
    { id: 'dka', diagnosi: 'Chetoacidosi diabetica in diabete mellito tipo 1 (scarsa aderenza)', lab: 'diabete', o2: ['AA'], vitto: 'Diabetico',
      problemi: ['Chetoacidosi diabetica risolta', 'DM1 scompensato'],
      terapia: function (T, da) { return [T.inCorso('lantus', da), T.inCorso('humalog', da), T.inCorso('isolyte', da), T.inCorso('pantorc', da), T.inCorso('clexane', da)]; },
      scale: [], colturali: [],
      apr: 'DM1 dall\'età di 14 anni, scarsa aderenza alla terapia insulinica.', app: 'Poliuria, polidipsia, vomito e dolore addominale. In PS glicemia 486, pH 7,12, chetoni 5,8: avviati idratazione e insulina ev.',
      diario: ['Gap anionico chiuso, chetoni negativi. Passaggio a schema basal-bolus. Glicemie 140-190.', 'Si alimenta. Educazione terapeutica con il team diabetologico.'],
      dafare: ['Profilo glicemico', 'Educazione terapeutica'], attesa: ['Consulenza diabetologica'], eseguiti: ['EGA seriati', 'Chetonemia'] },
    { id: 'osservazione', diagnosi: 'Sincope di natura da determinare, in osservazione', lab: 'normale', o2: ['AA'], vitto: 'Libero',
      problemi: ['Sincope in accertamento'],
      terapia: function (T, da) { return [T.inCorso('cardioasa', da), T.inCorso('torvast', da), T.inCorso('pantorc', da)]; },
      scale: [], colturali: [],
      apr: 'Dislipidemia, pregresso TIA nel 2021.', app: 'Episodio sincopale senza prodromi, con trauma cranico minore. TC encefalo negativa, ECG nella norma.',
      diario: ['Asintomatico. Telemetria: non aritmie. Prova di ortostatismo negativa.', 'In attesa di Holter ed ecocardiogramma.'],
      dafare: ['Telemetria 24h'], attesa: ['Holter ECG', 'Ecocardiogramma'], eseguiti: ['TC encefalo', 'ECG'] },
    { id: 'polmonite-ab', diagnosi: 'Polmonite ab ingestis in esiti di ictus ischemico', lab: 'sepsi', o2: ['VMK 35%', 'O2 4 lt/min CN'], vitto: 'Semiliquido', rianimazione: 'Non indicazione a manovre invasive (condiviso con i familiari)',
      problemi: ['Polmonite ab ingestis', 'Disfagia post-ictale', 'Esiti di ictus ischemico', 'Lesione da pressione sacrale stadio 2'],
      terapia: function (T, da) { return [T.inCorso('piptazo', da), T.inCorso('vancomicina', da + 1), T.inCorso('isolyte', da), T.inCorso('pantorc', da), T.inCorso('clexane', da), T.inCorso('perfalgan', da), T.inCorso('breva', da)]; },
      scale: ['gcs', 'news2'], colturali: ['Broncoaspirato ' + dm(giorno(-2)) + ': S. aureus MRSA', 'Emocolture ' + dm(giorno(-3)) + ': negative'], isolamento: 'Contatto — MRSA',
      apr: 'Ictus ischemico 2 anni fa con emiparesi dx e disfagia, allettamento, demenza.', app: 'Febbre, desaturazione e ingombro secretivo dopo un pasto. Rx torace: addensamenti basali bilaterali.',
      diario: ['Febbricola. Secrezioni abbondanti, broncoaspirazione al bisogno. SpO2 93% in VMK.', 'Colloquio con i familiari: condivisa la non indicazione a manovre invasive.'],
      dafare: ['Valutazione logopedica', 'Medicazione lesione sacrale'], attesa: ['Consulenza nutrizionale'], eseguiti: ['Rx torace', 'Broncoaspirato'] },
  ];

  // ── la scheda di un paziente, com'è il giorno dell'epoca ──────────────
  function costruisci(letto, tipologia, indice, rnd) {
    var q = QUADRI_CLINICI[indice % QUADRI_CLINICI.length];
    var femmina = rnd() < 0.5;
    var cognome = COGNOMI[(indice * 7 + 3) % COGNOMI.length], nome = (femmina ? NOMI_F : NOMI_M)[(indice * 5 + 1) % 14];
    var anni = 44 + Math.floor(rnd() * 48);
    var oggi = giorno(0);
    var nascita = giorno(0); nascita.setFullYear(nascita.getFullYear() - anni); nascita.setMonth(Math.floor(rnd() * 12)); nascita.setDate(1 + Math.floor(rnd() * 27));
    // età vera alla data dell'epoca (il compleanno può non essere ancora passato)
    var eta = oggi.getFullYear() - nascita.getFullYear() - ((oggi.getMonth() < nascita.getMonth() || (oggi.getMonth() === nascita.getMonth() && oggi.getDate() < nascita.getDate())) ? 1 : 0);
    var giorniRicovero = 2 + Math.floor(rnd() * (tipologia === 'SUB-INTENSIVA' ? 9 : 17));
    var da = -giorniRicovero;
    var nomeCompleto = (cognome + ' ' + nome).toUpperCase();
    var scegli = function (a) { return a[Math.floor(rnd() * a.length)]; };
    var parametri = function () { return 'PA ' + (105 + Math.floor(rnd() * 45)) + '/' + (60 + Math.floor(rnd() * 25)) + ' mmHg, FC ' + (62 + Math.floor(rnd() * 40)) + '/min, SpO2 ' + (91 + Math.floor(rnd() * 8)) + '% (' + q.o2[0] + '), TC ' + (36 + Math.floor(rnd() * 2)) + ',' + Math.floor(rnd() * 10) + ' °C. Diuresi ' + (9 + Math.floor(rnd() * 14)) + '00 ml/24h.'; };
    var EO = ['Vigile, orientato e collaborante. Eupnoico a riposo. Toni cardiaci validi, pause libere. Addome trattabile, non dolente, peristalsi presente. Non edemi declivi.',
      'Paziente vigile, collaborante. Cute e mucose normoidratate. MV presente su tutto l\'ambito, non rumori aggiunti di rilievo. Addome piano, trattabile. Polsi periferici presenti e simmetrici.',
      'Condizioni generali discrete. Lieve disorientamento temporale nelle ore serali. Decubito indifferente. Alvo aperto a feci formate. Si alimenta parzialmente.',
      'Notte tranquilla, riposo conservato. Non dolore. Deambula con assistenza. Catetere venoso periferico in sede, non segni di flebite.'];
    var PROGRAMMA = ['Prosegue terapia in atto. Monitoraggio dei parametri ogni 8 ore.', 'Si richiedono esami ematici di controllo per domattina.', 'Rivalutazione clinica nel pomeriggio. Informati i familiari.',
      'Si ottimizza la terapia come da prescrizione. Bilancio idrico.', 'Mobilizzazione in poltrona. Si contatta il servizio sociale per la dimissione protetta.'];
    var REFERTI = ['<div><u>Rx torace</u>: <i>non versamento pleurico; ombra cardiaca nei limiti; non lesioni pleuroparenchimali a focolaio oltre a quanto già noto.</i></div>',
      '<div><u>ECG</u>: ritmo sinusale a FC 78/min, PR 160 ms, QRS stretto, non alterazioni acute della ripolarizzazione.</div>',
      '<div><strong>Ecografia addome</strong>: fegato di dimensioni regolari a ecostruttura omogenea; colecisti alitiasica; reni in sede, non idronefrosi; milza nei limiti; non versamento libero.</div>',
      '<div><u>TC encefalo senza mdc</u>: <i>non lesioni emorragiche né aree di alterata densità di significato acuto; sistema ventricolare in asse.</i></div>',
      '<div><strong>Ecocardiogramma</strong>: ventricolo sinistro di normali dimensioni, FE 55%; non valvulopatie di rilievo; VD nei limiti; non versamento pericardico.</div>',
      '<div><font color="#c62828"><b>Attenzione:</b> rivalutare la funzione renale prima di eventuale mdc.</font></div>'];
    var voci = [];
    for (var gI = da + 1; gI <= 0; gI++) {
      var testoBase = q.diario[Math.min(q.diario.length - 1, Math.floor((gI - da - 1) * q.diario.length / Math.max(1, giorniRicovero)))];
      voci.push('<div>' + b(dm(giorno(gI)) + ':') + ' ' + esc(parametri()) + '</div><div>' + esc(scegli(EO)) + ' ' + esc(testoBase) + '</div><div>' + esc(scegli(PROGRAMMA)) + '</div>'
        + (rnd() < 0.45 ? scegli(REFERTI) : '')
        + '<div><i>Pomeriggio:</i> ' + esc(scegli(EO)) + ' ' + esc(scegli(PROGRAMMA)) + '</div><div><br></div>');
    }
    var diaria = '<div>' + b('APR:') + ' ' + esc(q.apr) + '</div><div>' + b('APP:') + ' ' + esc(q.app) + '</div><div><br></div>'
      + '<div>' + b('All\'ingresso:') + ' ' + esc(parametri()) + ' ' + esc(EO[0]) + '</div>' + scegli(REFERTI) + '<div><br></div>'
      + voci.join('')
      + (indice % 4 === 0 ? '<div><span style="background-color: rgb(255, 235, 59);">Parlato con i familiari, aggiornati sulle condizioni cliniche.</span></div>' : '')
      + '<div><br></div>';
    var dafare = b('DA FARE:') + '<br>' + q.dafare.map(function (t) { return '- ' + esc(t) + '<br>'; }).join('') + '<br>'
      + b('RICHIESTI/IN ATTESA:') + '<br>' + q.attesa.map(function (t) { return '- ' + esc(t) + '<br>'; }).join('') + '<br>'
      + b('ESEGUITI:') + '<br>' + q.eseguiti.map(function (t) { return '- ' + esc(t) + '<br>'; }).join('') + '<br>'
      + b('NOTE:') + '<br>' + (indice % 2 === 0 ? '<div>Familiare di riferimento: figlia (contatto in cartella).</div><div><u>Dimissione</u>: da programmare con il servizio sociale, valutare ADI.</div>' : '<div><span style="background-color: rgb(255, 235, 59);">Ricontrollare esami in sospeso prima della dimissione.</span></div>');
    // Un paziente non ha esami nemmeno sul TrakCare (griglia vuota); un terzo
    // li ha sul TrakCare ma nessuno li ha ancora importati nella scheda.
    var senzaEsami = indice === 10;
    return {
      quadro: q, da: da, nome: nomeCompleto, indice: indice,
      senzaEsami: senzaEsami, labNellaScheda: !senzaEsami && indice % 3 !== 2,
      paziente: {
        Letto: letto, TipologiaLetto: tipologia, Nome: nomeCompleto, Sesso: femmina ? 'F' : 'M', Eta: String(eta), DataNascita: dmy(nascita), DataRicovero: iso(giorno(da)),
        Diagnosi: q.diagnosi, Allergie: ALLERGIE[indice % ALLERGIE.length], CodiceSanitario: String(90000000 + Math.floor(rnd() * 9999999)),
        Ossigeno: q.o2[indice % q.o2.length], Vitto: q.vitto, Dimissibile: (indice % 9 === 4) ? dmy(giorno(0)) : '',
        Diaria: diaria, DaFare: dafare, PianoTerapeutico: '', EsamiColturali: q.colturali.map(function (t) { return '<div>' + esc(t) + '</div>'; }).join('') || '<div><br></div>',
        NoteTerapia: window._TEMPLATE_NOTETERAPIA_VUOTO, UltimoAggiornamento: '',
      },
    };
  }

  // I piani dei letti occupati, nell'ordine del reparto.
  function pianifica() {
    var rnd = prng(SEME), piani = [];
    LETTI.forEach(function (l) {
      if (LIBERI.indexOf(l[0]) >= 0) return;
      piani.push(costruisci(l[0], l[1], piani.length, rnd));
    });
    return piani;
  }

  // Le righe del Pannello Terapia del paziente nel giorno simulato «s» (0 =
  // epoca), nell'ordine in cui la pagina le mostra: prima il gruppo ENDOVENA.
  // Sono esattamente ciò che il segnalibro «Prendi Terapia Trak» legge.
  function terapiaDi(p, s) {
    s = s || 0;
    function componi(giornoSimulato, sospese) {
      var T = attrezzi(p, giornoSimulato, sospese);
      var righe = p.quadro.terapia(T, p.da).filter(Boolean);
      var presenti = righe.map(function (x) { return x.Prestazione.slice(0, 18); });
      FARMACI_DI_CASA.forEach(function (c, ci) {
        if ((p.indice + ci) % 3 !== 0 || presenti.indexOf(F[c].slice(0, 18)) >= 0) return;
        var r = ci % 4 === 3 ? T.chiuso(c, p.da, p.da + 2, 'Sospeso') : T.inCorso(c, p.da, ci === 4 ? 'se alvo chiuso da 48 ore' : '');
        if (r) { righe.push(r); presenti.push(F[c].slice(0, 18)); }
      });
      if (p.indice % 4 === 1) righe.push(T.inCorso('ordineReparto', p.da));
      if (p.indice % 5 === 2) righe.push(T.chiuso('perErrore', p.da, p.da, 'Sospeso', 'Errore inserimento'));
      return { T: T, righe: righe, presenti: presenti };
    }
    var sospese = {};
    if (s >= 1) {
      var allEpoca = componi(0, {}).righe.filter(function (r) { return !r.Stato; }).map(function (r) { return r.chiave; });
      var daSospendere = SOSPENDIBILI.filter(function (c) { return allEpoca.indexOf(c) >= 0; })[0];
      if (daSospendere) sospese[daSospendere] = 1;
    }
    var c = componi(s, sospese);
    if (s >= 1) c.righe.push(c.T.immediata(p.indice % 2 === 0 ? 'plasil' : 'lasixSubito', [1]));
    if (s >= 2) {
      var nuovo = AGGIUNGIBILI.filter(function (k) { return c.presenti.indexOf(F[k].slice(0, 18)) < 0; })[0];
      if (nuovo) c.righe.push(c.T.inCorso(nuovo, 2));
    }
    var pulita = function (r) { return { Prestazione: r.Prestazione, Gruppo: r.Gruppo, Giorni: r.Giorni, Info: r.Info, Stato: r.Stato, StatoInfo: r.StatoInfo, _r: r._r }; };
    var tutte = c.righe.filter(Boolean);
    return tutte.filter(function (r) { return r.Gruppo === 'ENDOVENA'; }).concat(tutte.filter(function (r) { return r.Gruppo !== 'ENDOVENA'; })).map(pulita);
  }

  // Le pagine della griglia Laboratorio nel giorno simulato «s», dalla più
  // recente: fino a 12 prelievi (uno al giorno dal ricovero), 5 per pagina.
  // Il valore di un esame in un dato giorno non dipende da «s».
  function labDi(p, s) {
    s = s || 0;
    if (p.senzaEsami) return [];
    var mod = QUADRI[p.quadro.lab] || {};
    var giorni = []; for (var d = Math.max(p.da, s - MAX_PRELIEVI + 1); d <= s; d++) giorni.push(d);
    function eseguito(j, d) {
      if (!RARO[j]) return true;
      var degenza = d - p.da;
      return degenza === 0 || degenza % 3 === 0 || (!!mod[PANNELLO[j][1]] && degenza % 2 === 0);
    }
    function valore(j, d) {
      var e = PANNELLO[j], rnd = prng(mescola(SEME, p.indice + 1, j + 1, d - p.da + 1));
      if (e[7]) return e[7][Math.floor(rnd() * e[7].length)];
      var m = mod[e[1]], base = e[5];
      if (m) base = d <= 0 ? m[0] + (m[1] - m[0]) * ((d - p.da) / -p.da) : m[1] + (e[5] - m[1]) * Math.min(1, 0.25 * d);
      // oscillazione proporzionata all'ampiezza dell'intervallo, non al valore:
      // pH e peso specifico non devono ballare come una glicemia
      var v = base + (rnd() - 0.5) * (e[3] - e[2]) * 0.3;
      if (e[2] >= 0 && v < 0) v = 0;
      var txt = v.toFixed(e[6]), n = Number(txt);
      return txt + (n > e[3] ? ' H' : (n < e[2] ? ' L' : ''));
    }
    function etichetta(d) { return 'LAB00' + (81000000 + Math.floor(prng(mescola(SEME, p.indice + 1, 7777, d - p.da + 1))() * 900000)) + ' ' + dmy(giorno(d)) + ' 06:00'; }
    var pagine = [];
    for (var fine = giorni.length; fine > 0; fine -= PER_PAGINA) {
      var colonne = giorni.slice(Math.max(0, fine - PER_PAGINA), fine), righe = [], gruppo = '';
      PANNELLO.forEach(function (e, j) {
        var vals = colonne.map(function (d) { return eseguito(j, d) ? valore(j, d) : ''; });
        if (vals.join('') === '') return;
        if (e[0] !== gruppo) { gruppo = e[0]; righe.push({ c: [gruppo, '', ''].concat(colonne.map(function () { return ''; })), b: 1 }); }
        righe.push({ c: [e[1], e[7] ? '' : (e[2] + ' - ' + e[3]), e[4]].concat(vals), b: 0 });
      });
      pagine.push({ int: ['Esame', 'Intervallo di riferimento', 'Unità'].concat(colonne.map(etichetta)), righe: righe });
    }
    return pagine;
  }

  return { EPOCA: EPOCA, SEME: SEME, GIORNI_AVANTI: GIORNI_AVANTI, LETTI: LETTI, LIBERI: LIBERI, attendi: attendi, giorno: giorno, dmy: dmy, finestra: finestra, pianifica: pianifica, terapiaDi: terapiaDi, labDi: labDi };
})();
