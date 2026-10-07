// Batteria di prove della funzione decorso-clinico nel progetto di COLLAUDO.
// Scrive contenuti noti sul letto libero di prova, chiama la funzione come
// utente autorizzato (e come estraneo) e controlla ciò che torna indietro:
// chi può chiamare, limiti, HTML dei campi → testo, blocchi derivati tolti,
// terapia importata separata, esami ordinati, referti separati e in ordine
// cronologico, diaria intera. Alla fine rimette il letto com'era.
//   node collaudo/strumenti/prova-funzione-decorso.js
const { collaudo } = require('./sb.js');
const { genera } = require('./gettone-prova.js');
const amb = require('./ambiente.json');

const LETTO = '5';   // il letto libero usato dalle prove automatiche
const esiti = [];
function prova(nome, ok, dettaglio) {
  esiti.push(ok);
  console.log((ok ? 'OK  ' : 'KO  ') + nome + (dettaglio ? ' — ' + dettaglio : ''));
}
const dollaro = (t) => { const d = '$p' + Math.random().toString(36).slice(2, 8) + '$'; if (t.includes(d)) throw new Error('delimitatore'); return d + t + d; };

const TS_TERAPIA = 1791124930674;   // 04/10/2026 16:42 (ora italiana)
const SCHEDA = {
  nome: 'PROVA DECORSO <img src=x onerror="x()">',
  data_nascita: '01/01/1950', eta: '76', sesso: 'M', data_ricovero: '2026-10-01', tipologia_letto: 'STANDARD', codice_sanitario: 'CS000',
  diagnosi: 'Polmonite <b>comunitaria</b> basale dx',
  allergie: '<b>Nessuna</b> allergia&nbsp;nota',
  diaria: '<div><b>01/10:</b> ingresso in reparto.&nbsp;PA 140/80 &amp; FC 90</div><div>02/10: migliora<br>apiretico</div><div data-isolamento="contatto">Isolamento da contatto</div><ul><li>Primo punto</li><li>Secondo punto</li></ul>',
  piano_terapeutico: '<ul><li>Problema 1</li><li>Problema 2</li></ul>',
  da_fare: '<div>Rx torace</div><div>EGA</div>',
  esami_colturali: '<div class="lab-box"><div class="lab-apri-btn">Leucociti 13.0 H</div><div>copia degli esami</div></div><div class="scale-box"><div class="scala-risultato" data-scala="news2">NEWS2 = 5 (rischio medio)</div></div><div>Urinocoltura negativa</div>',
  note_terapia: '<div class="terapia-box" data-ts="' + TS_TERAPIA + '"><div class="terapia-box-head">Aggiornata da TrakCare il 04/10 16:42</div><div style="font-weight:700">Terapia EV</div><div>- Ceftriaxone 2g ×1 (dal 25/09)</div></div><div>Note a mano: sospendere il 10/10</div>',
  ossigeno: 'O2 2 lt/min', vitto: 'Leggero', dimissibile: '',
};
const ESAMI = {
  esami: {
    a: { g: 'EMOCROMO -> Sangue', n: 'Leucociti', um: '10^9/L', range: '4 - 10', v: { '02/10/2026 06:00': '11.2 H', '01/10/2026 06:00': '13.0 H' } },
    b: { g: 'ELETTROLITI -> Sangue', n: 'Sodio', um: 'mmol/L', range: '135 - 145', v: { '01/10/2026 06:00': '138' } },
  },
  ord: ['a', 'b'],
};
const REFERTI = [
  'Nota iniziale senza titolo né data',
  'TC cranio del 03/10/2026',
  'Non lesioni acute. Confronto con esame del 20/09/2026,',
  'sovrapponibile.',
  '',
  'Consulenza cardiologica del 02/10/2026',
  'Ritmo sinusale, nessuna indicazione.',
  '',
  'Ecografia addome 1 ottobre 2026',
  'Fegato nei limiti.',
  'Controllo <b>tra</b> 7 giorni',
].join('\n');
const DIARIA = ['01/10/2026 ore 10:00 Ingresso. Paziente vigile.', 'Obiettività: nulla da segnalare.', '', '02/10/2026 ore 09:00 Apiretico, migliora.', 'Prosegue terapia.   ', '03/10/2026 ore 08:30 Fine diaria.'].join('\r\n');

(async () => {
  const c = collaudo();
  const { token: utente } = await genera(1);
  const prima = await c.query("select * from public.consegne where letto = '" + LETTO + "'");
  if (!prima.length) throw new Error('il letto di prova ' + LETTO + ' non esiste nel collaudo');
  if ((prima[0].nome || '').trim()) throw new Error('il letto di prova ' + LETTO + ' non è libero');
  const labPrima = await c.query("select * from public.lab_esami where letto = '" + LETTO + "'");
  const impPrima = await c.query("select count(*)::int as n from public.impostazioni");

  const chiama = async (bearer, corpo, grezzo) => {
    const r = await fetch(amb.url + '/functions/v1/decorso-clinico', {
      method: 'POST', headers: { 'Content-Type': 'application/json', apikey: amb.anon, Authorization: 'Bearer ' + bearer },
      body: grezzo !== undefined ? grezzo : JSON.stringify(corpo),
    });
    const testo = await r.text();
    const tipo = r.headers.get('content-type') || '';
    if (!/x-ndjson/.test(tipo)) { let j = null; try { j = JSON.parse(testo); } catch (e) {} return { http: r.status, tipo, j: j || {}, eventi: [] }; }
    const eventi = testo.split('\n').filter((l) => l.trim()).map((l) => JSON.parse(l));
    const fine = eventi.find((e) => e.passo === 'fine') || null;
    return { http: r.status, tipo, j: {}, eventi, fine };
  };

  try {
    // ── il letto di prova con contenuti noti ──
    const set = Object.keys(SCHEDA).map((k) => k + ' = ' + dollaro(SCHEDA[k])).join(', ');
    await c.query("update public.consegne set " + set + ", updated_at = now() where letto = '" + LETTO + "'");
    await c.query("insert into public.lab_esami (letto, paziente, dati, n_esami, ultimo_esame) values ('" + LETTO + "', 'PROVA DECORSO', " + dollaro(JSON.stringify(ESAMI)) + "::jsonb, 2, '2026-10-02') "
      + "on conflict (letto) do update set paziente = excluded.paziente, dati = excluded.dati, n_esami = excluded.n_esami, ultimo_esame = excluded.ultimo_esame");

    // ── chi può chiamare ──
    let r = await chiama(amb.anon, { letto: LETTO, referti: '', diaria: '' });
    prova('estraneo (sola chiave pubblica): rifiutato prima di leggere qualunque cosa', r.http === 403 && r.j.non_autorizzato === true, 'HTTP ' + r.http);
    r = await chiama(utente, null, '{non json');
    prova('JSON non valido: rifiutato', r.http === 400);
    r = await chiama(utente, { referti: '', diaria: '' });
    prova('letto mancante: rifiutato', r.http === 400 && /letto/.test(r.j.errore || ''), r.j.errore);
    r = await chiama(utente, { letto: LETTO, referti: 'x'.repeat(200001), diaria: '' });
    prova('referti oltre il limite: rifiutati prima di tutto', r.http === 400 && /troppo lunghi/.test(r.j.errore || ''), r.j.errore);
    r = await chiama(utente, { letto: LETTO, referti: '', diaria: 'x'.repeat(300001) });
    prova('diaria oltre il limite: rifiutata', r.http === 400 && /troppo lunga/.test(r.j.errore || ''), r.j.errore);
    r = await chiama(utente, { letto: 'ZZZ-inesistente', referti: '', diaria: '' });
    prova('letto inesistente: il passo «consegne» fallisce e la fine lo dice', r.http === 200 && r.eventi.some((e) => e.passo === 'consegne' && e.stato === 'errore') && r.fine && r.fine.stato === 'errore' && /non esiste/.test(r.fine.errore || ''), r.fine && r.fine.errore);

    // ── la chiamata buona ──
    r = await chiama(utente, { letto: LETTO, referti: REFERTI, diaria: DIARIA, impostazioni: { righe: 20, dettaglio: 9 } });
    const passi = r.eventi.filter((e) => e.stato === 'ok' || e.stato === 'salto').map((e) => e.passo);
    prova('risposta a flusso, un evento per passo, nell\'ordine giusto', r.http === 200 && /x-ndjson/.test(r.tipo) && passi.join(',') === 'ricezione,consegne,terapia,laboratorio,referti,diaria,fascicolo,fine', passi.join(','));
    prova('ogni passo annuncia l\'inizio prima dell\'esito', ['consegne', 'terapia', 'laboratorio', 'referti', 'diaria', 'fascicolo'].every((p) => { const i = r.eventi.findIndex((e) => e.passo === p && e.stato === 'inizio'); const j = r.eventi.findIndex((e) => e.passo === p && e.stato !== 'inizio'); return i >= 0 && j > i; }));
    const f = (r.fine && r.fine.fascicolo) || {};
    prova('fase di prova dichiarata, riepilogo presente', r.fine && r.fine.fase === 'prova' && r.fine.riepilogo && r.fine.riepilogo.byte > 1000, JSON.stringify(r.fine && r.fine.riepilogo));
    prova('impostazioni: 20 righe, dettaglio fuori scala riportato a 4', f.impostazioni && f.impostazioni.righe === 20 && f.impostazioni.dettaglio === 4 && f.impostazioni.dettaglioTesto === 'estremamente dettagliato', JSON.stringify(f.impostazioni));

    const id = f.identita || {}, cons = f.consegne || {};
    prova('identità: i tag nel nome spariscono, il resto è testo', id.nome === 'PROVA DECORSO' && id.letto === LETTO && id.eta === '76' && id.data_nascita === '01/01/1950', JSON.stringify(id));
    prova('diagnosi e allergie: HTML → testo, entità decodificate', cons.diagnosi === 'Polmonite comunitaria basale dx' && cons.allergie === 'Nessuna allergia nota', JSON.stringify([cons.diagnosi, cons.allergie]));
    prova('diaria delle consegne: righe al posto giusto, «&amp;» decodificato, elenco puntato', cons.diaria === '01/10: ingresso in reparto. PA 140/80 & FC 90\n02/10: migliora\napiretico\nIsolamento da contatto\n• Primo punto\n• Secondo punto', JSON.stringify(cons.diaria));
    prova('problemi attivi e da fare', cons.problemi_attivi === '• Problema 1\n• Problema 2' && cons.da_fare === 'Rx torace\nEGA', JSON.stringify([cons.problemi_attivi, cons.da_fare]));
    prova('esami colturali: il riquadro del laboratorio (copia degli esami) è tolto, la scala e il testo restano', cons.esami_colturali_e_scale === 'NEWS2 = 5 (rischio medio)\nUrinocoltura negativa', JSON.stringify(cons.esami_colturali_e_scale));
    prova('nessun tag HTML in ciò che viene dai campi della scheda (il testo incollato invece resta com\'è)', !/<[a-z!\/]/i.test(JSON.stringify({ a: f.identita, b: f.ricovero, c: f.consegne, d: f.terapia })));

    const t = f.terapia || {};
    prova('terapia: riquadro importato da TrakCare separato dalle note, con la data dell\'import', t.importata_da_trakcare === true && t.importata_il === new Date(TS_TERAPIA).toISOString() && /Terapia EV\n- Ceftriaxone 2g ×1 \(dal 25\/09\)$/.test(t.testo) && t.note === 'Note a mano: sospendere il 10/10', JSON.stringify(t));

    const lab = f.laboratorio || {};
    const leu = (lab.esami || [])[0] || {};
    prova('laboratorio: esami nell\'ordine salvato, valori in ordine di data', lab.presenti === true && lab.n_esami === 2 && lab.n_valori === 3 && leu.nome === 'Leucociti' && leu.um === '10^9/L' && leu.range === '4 - 10' && leu.valori.map((v) => v.data).join('|') === '01/10/2026 06:00|02/10/2026 06:00' && lab.dal === '01/10/2026' && lab.al === '02/10/2026', JSON.stringify(lab).slice(0, 300));

    const ref = f.referti || {};
    const titoli = (ref.elenco || []).map((x) => (x.titolo || '(senza titolo)') + ' @ ' + (x.data || '-'));
    prova('referti: separati dalle righe con titolo e data e messi in ordine cronologico, senza data in coda',
      titoli.join(' | ') === 'Ecografia addome 1 ottobre 2026 @ 2026-10-01 | Consulenza cardiologica del 02/10/2026 @ 2026-10-02 | TC cranio del 03/10/2026 @ 2026-10-03 | (senza titolo) @ -' && ref.senza_data === 1, titoli.join(' | '));
    const tc = (ref.elenco || []).find((x) => /^TC cranio/.test(x.titolo)) || {};
    prova('una riga con una data che finisce con la virgola non spezza il referto', tc.testo === 'Non lesioni acute. Confronto con esame del 20/09/2026,\nsovrapponibile.', JSON.stringify(tc.testo));
    const eco = (ref.elenco || []).find((x) => /^Ecografia/.test(x.titolo)) || {};
    prova('il testo dei referti resta com\'è (è testo incollato, non HTML)', eco.testo === 'Fegato nei limiti.\nControllo <b>tra</b> 7 giorni' && (ref.elenco || [])[3].testo === 'Nota iniziale senza titolo né data', JSON.stringify(eco.testo));

    const d = f.diaria || {};
    prova('diaria: intera, fine riga uniformati, spazi finali tolti, inizio e fine riportati', d.righe === 6 && d.caratteri === DIARIA.replace(/\r\n/g, '\n').replace(/[ \t]+$/gm, '').length && d.inizio === '01/10/2026 ore 10:00 Ingresso. Paziente vigile.' && d.fine === '03/10/2026 ore 08:30 Fine diaria.' && d.date_trovate === 3 && d.testo.indexOf('Prosegue terapia.\n') >= 0, JSON.stringify({ righe: d.righe, caratteri: d.caratteri, inizio: d.inizio, fine: d.fine, date: d.date_trovate }));
    prova('riepilogo coerente con il fascicolo', r.fine.riepilogo.referti === 4 && r.fine.riepilogo.referti_senza_data === 1 && r.fine.riepilogo.righe_diaria === 6 && r.fine.riepilogo.esami === 2 && r.fine.riepilogo.campi_consegne === 8, JSON.stringify(r.fine.riepilogo));

    // ── senza referti, senza diaria, senza esami: passi «saltati», non errori ──
    await c.query("delete from public.lab_esami where letto = '" + LETTO + "'");
    r = await chiama(utente, { letto: LETTO, referti: '   ', diaria: '' });
    const salti = r.eventi.filter((e) => e.stato === 'salto').map((e) => e.passo);
    prova('senza esami, referti e diaria: passi saltati con spiegazione, fine ok, impostazioni predefinite', r.fine && r.fine.stato === 'ok' && salti.join(',') === 'laboratorio,referti,diaria' && r.fine.fascicolo.impostazioni.righe === 15 && r.fine.fascicolo.impostazioni.dettaglio === 2 && r.fine.fascicolo.laboratorio.presenti === false, salti.join(','));
    await c.query("update public.consegne set note_terapia = '' where letto = '" + LETTO + "'");
    r = await chiama(utente, { letto: LETTO, referti: '', diaria: '' });
    prova('senza terapia: passo saltato', r.fine && r.fine.stato === 'ok' && r.eventi.some((e) => e.passo === 'terapia' && e.stato === 'salto'));

    // ── nulla resta sul server ──
    const impDopo = await c.query("select count(*)::int as n from public.impostazioni");
    prova('il server non ha scritto nulla nelle impostazioni', impDopo[0].n === impPrima[0].n);
  } finally {
    // il letto di prova torna com'era
    const colonne = Object.keys(prima[0]).filter((k) => k !== 'letto');
    await c.query("update public.consegne set " + colonne.map((k) => k + ' = ' + (prima[0][k] === null ? 'null' : dollaro(String(prima[0][k])))).join(', ') + " where letto = '" + LETTO + "'");
    await c.query("delete from public.lab_esami where letto = '" + LETTO + "'");
    if (labPrima.length) {
      const l = labPrima[0];
      await c.query("insert into public.lab_esami (letto, paziente, dati, n_esami, ultimo_esame, allarme_giorni, allarme_visto) values ('" + LETTO + "', " + dollaro(String(l.paziente || '')) + ", " + dollaro(JSON.stringify(l.dati)) + "::jsonb, " + (l.n_esami || 0) + ", " + (l.ultimo_esame ? "'" + l.ultimo_esame + "'" : 'null') + ", " + (l.allarme_giorni == null ? 'null' : l.allarme_giorni) + ", " + (l.allarme_visto ? "'" + l.allarme_visto + "'" : 'null') + ")");
    }
    const dopo = await c.query("select (nome = '' or nome is null) as libero from public.consegne where letto = '" + LETTO + "'");
    console.log('\nletto di prova rimesso com\'era: ' + (dopo[0] && dopo[0].libero ? 'sì' : 'DA GUARDARE'));
  }

  const n = esiti.filter(Boolean).length;
  console.log('\nESITO: ' + n + '/' + esiti.length + (n === esiti.length ? ' — tutte superate' : ' — CI SONO PROVE FALLITE'));
  if (n !== esiti.length) process.exitCode = 1;
})().catch((e) => { console.log('ERRORE: ' + e.message); process.exit(1); });
