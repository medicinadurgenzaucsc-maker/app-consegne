// Copia l'elenco della rubrica telefonica dal suo vecchio database (progetto
// Supabase «rubrica-gemelli», account del reparto) nelle tabelle rubrica_ del
// database delle consegne.
//
//   node collaudo/strumenti/copia-rubrica.js collaudo               svuota la rubrica del collaudo e la riempie di nuovo
//   node collaudo/strumenti/copia-rubrica.js collaudo confronta     dice solo se le due copie coincidono
//   node collaudo/strumenti/copia-rubrica.js collaudo prima-copia   come in produzione: copia SOLO se la destinazione è vuota
//   node collaudo/strumenti/copia-rubrica.js produzione confronta   sola lettura
//   node collaudo/strumenti/copia-rubrica.js produzione prima-copia --confermo-produzione
//
// La sorgente si legge soltanto (transazione di sola lettura imposta dal
// server). La destinazione viene riempita da UNA sola istruzione, che alla
// fine ricalcola conteggi e impronta e li confronta con quelli della
// sorgente: se non coincidono l'istruzione fallisce e la destinazione resta
// com'era. O entra tutto, con gli stessi identificativi, o non cambia nulla.
//
// In PRODUZIONE la rubrica non si svuota mai: esiste solo la «prima copia»,
// che si rifiuta se nella destinazione c'è già anche una sola riga. Dal
// momento in cui il reparto usa la rubrica dentro l'app, le modifiche dei
// colleghi stanno lì e nessuno strumento deve poterle sovrascrivere. La
// prima copia vuole l'ok di chi gestisce l'app e `--confermo-produzione`.
//
// Nomi e numeri non vengono mai stampati: escono solo conteggi e impronte, e
// di un errore del database solo lo stato, il codice e i nomi di tabelle,
// colonne e vincoli.
// Nel collaudo ci sono i numeri VERI per decisione di Stefano dell'08/10/2026:
// sono numeri di telefono dell'ospedale, non dati di pazienti, e servono a
// provare la rubrica com'è davvero.
const fs = require('fs');
const os = require('os');
const path = require('path');
const { REF_COLLAUDO, REF_PRODUZIONE, maschera } = require('./sb.js');

const REF_RUBRICA = 'nbbekxuvuarxkuvvvgbi';
const CONFERMA = '--confermo-produzione';
const DESTINAZIONI = {
  collaudo: { voce: 'supabase-collaudo', ref: REF_COLLAUDO, valore: 'collaudo', nome: 'del collaudo' },
  produzione: { voce: 'supabase', ref: REF_PRODUZIONE, valore: 'produzione', nome: 'di produzione' },
};
// L'impronta di un elenco: ogni riga diventa un elenco JSON dei suoi campi (così
// un campo vuoto, un campo assente e un separatore dentro il testo restano
// distinti), se ne fa l'impronta, e poi l'impronta delle impronte in ordine di id.
const IMPRONTA_CONTATTI = "coalesce(md5(string_agg(md5(jsonb_build_array(id, nome, categoria, numeri, note)::text), '' order by id)), '')";
const IMPRONTA_CATEGORIE = "coalesce(md5(string_agg(md5(jsonb_build_array(id, nome, ordine)::text), '' order by id)), '')";
const firma = (impronta, tabella) => 'select count(*)::int as n, ' + impronta + ' as impronta from public.' + tabella;

function token(voce, refAtteso) {
  let j = null;
  try { j = JSON.parse(fs.readFileSync(path.join(os.homedir(), '.claude.json'), 'utf8')); }
  catch (e) { throw new Error('la configurazione locale (~/.claude.json) non si legge'); }
  const a = ((j.mcpServers || {})[voce] || {}).args || [];
  if (refAtteso && !a.includes('--project-ref=' + refAtteso)) throw new Error('la voce «' + voce + '» non punta al progetto atteso');
  const t = a.find((x) => x.startsWith('--access-token='));
  if (!t) throw new Error('chiave assente nella voce «' + voce + '»');
  return t.slice('--access-token='.length);
}
// Di un errore del database si riporta solo ciò che non può contenere dati: stato,
// codice, nomi di vincoli, tabelle e colonne, e i nostri messaggi fissi.
function erroreSenzaDati(ref, stato, testo) {
  const codice = (/ERROR:\s+([0-9A-Z]{5})\b/.exec(testo) || /"code"\s*:\s*"([0-9A-Z]{5})"/.exec(testo) || [])[1] || 'senza codice';
  const nomi = [];
  String(testo).replace(/(constraint|relation|column|function) \\?"([A-Za-z0-9_.]+)/g, (tutto, tipo, nome) => { nomi.push(tipo + ' ' + nome); return tutto; });
  const nostro = (/(COPIA NON FEDELE|DESTINAZIONE NON VUOTA|la destinazione non si dichiara (?:collaudo|produzione))/.exec(testo) || [])[1];
  return new Error('database ' + ref.slice(0, 6) + '… (HTTP ' + stato + ', codice ' + codice + ')' + (nostro ? ': ' + nostro : '') + (nomi.length ? ' [' + nomi.slice(0, 4).join(', ') + ']' : ''));
}
async function interroga(ref, tok, sql, solaLettura) {
  const r = await fetch('https://api.supabase.com/v1/projects/' + ref + '/database/query', {
    method: 'POST',
    headers: { Authorization: 'Bearer ' + tok, 'Content-Type': 'application/json' },
    body: JSON.stringify(solaLettura ? { query: sql, read_only: true } : { query: sql }),
  });
  const t = await r.text();
  if (r.status >= 300) throw erroreSenzaDati(ref, r.status, t);
  let righe = null;
  try { righe = JSON.parse(t); } catch (e) { righe = null; }
  // una risposta che non è un elenco non vale «zero righe»
  if (!Array.isArray(righe)) throw new Error('database ' + ref.slice(0, 6) + '…: risposta che non si legge (HTTP ' + r.status + ')');
  return righe;
}

(async () => {
  const argomenti = process.argv.slice(2);
  // anche il codice di un reparto: «gemelli-11n» è la produzione
  const ambiente = require('./reparti.js').nomeStorico(argomenti[0]) || argomenti[0], azione = argomenti[1] || '', resto = argomenti.slice(2);
  const D = Object.prototype.hasOwnProperty.call(DESTINAZIONI, String(ambiente)) ? DESTINAZIONI[ambiente] : null;
  const ammesso = D && (
    (ambiente === 'collaudo' && ['', 'confronta', 'prima-copia'].includes(azione) && resto.length === 0)
    || (ambiente === 'produzione' && azione === 'confronta' && resto.length === 0)
    || (ambiente === 'produzione' && azione === 'prima-copia' && resto.length === 1 && resto[0] === CONFERMA));
  if (!ammesso) {
    console.log('uso: copia-rubrica.js collaudo [confronta|prima-copia]');
    console.log('     copia-rubrica.js produzione confronta');
    console.log('     copia-rubrica.js produzione prima-copia ' + CONFERMA + '   (solo con l\'ok di chi gestisce l\'app)');
    process.exit(1);
  }
  // nel collaudo, senza azione, la rubrica si svuota e si riempie; in produzione non si svuota MAI
  const svuota = ambiente === 'collaudo' && azione === '';
  // la chiave dell'account del reparto vede anche il progetto della rubrica; lì si legge soltanto
  const tokSorgente = token('supabase', null);
  const tokDest = token(D.voce, D.ref);
  const sorgente = (sql) => interroga(REF_RUBRICA, tokSorgente, sql, true);
  const dest = (sql, lettura) => interroga(D.ref, tokDest, sql, !!lettura);

  const amb = await dest("select valore from public.impostazioni where chiave = 'AMBIENTE'", true);
  if (!amb[0] || amb[0].valore !== D.valore) throw new Error('il database di destinazione non si dichiara «' + D.valore + '»: non tocco nulla');

  const firme = async () => ({
    sc: (await sorgente(firma(IMPRONTA_CONTATTI, 'contatti')))[0], sk: (await sorgente(firma(IMPRONTA_CATEGORIE, 'categorie')))[0],
    dc: (await dest(firma(IMPRONTA_CONTATTI, 'rubrica_contatti'), true))[0], dk: (await dest(firma(IMPRONTA_CATEGORIE, 'rubrica_categorie'), true))[0],
  });
  const mostra = (f) => {
    console.log('  sorgente:     ' + f.sc.n + ' contatti (' + f.sc.impronta.slice(0, 12) + '), ' + f.sk.n + ' categorie (' + f.sk.impronta.slice(0, 12) + ')');
    console.log('  destinazione: ' + f.dc.n + ' contatti (' + f.dc.impronta.slice(0, 12) + '), ' + f.dk.n + ' categorie (' + f.dk.impronta.slice(0, 12) + ')');
    return f.sc.n === f.dc.n && f.sk.n === f.dk.n && f.sc.impronta === f.dc.impronta && f.sk.impronta === f.dk.impronta;
  };
  console.log('destinazione: database ' + D.nome + ' (' + D.ref.slice(0, 6) + '…)');
  if (azione === 'confronta') {
    const uguali = mostra(await firme());
    console.log(uguali ? 'LE DUE COPIE COINCIDONO' : 'LE DUE COPIE SONO DIVERSE');
    process.exit(uguali ? 0 : 1);
  }
  if (!svuota) {
    const gia = (await dest('select (select count(*) from public.rubrica_contatti)::int as contatti, (select count(*) from public.rubrica_categorie)::int as categorie', true))[0];
    if (gia.contatti || gia.categorie) throw new Error('la destinazione non è vuota (' + gia.contatti + ' contatti, ' + gia.categorie + ' categorie): la prima copia non sovrascrive nulla, mi fermo');
  }

  // 1. si legge la sorgente (in memoria: nulla viene stampato né scritto su disco) e, DOPO, la sua impronta
  const contatti = await sorgente('select id, nome, categoria, numeri, note from public.contatti order by id');
  const categorie = await sorgente('select id, nome, ordine from public.categorie order by id');
  const fc = (await sorgente(firma(IMPRONTA_CONTATTI, 'contatti')))[0], fk = (await sorgente(firma(IMPRONTA_CATEGORIE, 'categorie')))[0];
  if (!contatti.length || !categorie.length) throw new Error('la sorgente risulta vuota (' + contatti.length + ' contatti, ' + categorie.length + ' categorie): non tocco la destinazione');
  if (contatti.length !== fc.n || categorie.length !== fk.n) throw new Error('la sorgente è cambiata durante la lettura: rilanciare');
  if (!/^[0-9a-f]{32}$/.test(fc.impronta) || !/^[0-9a-f]{32}$/.test(fk.impronta)) throw new Error('impronta della sorgente non valida');
  console.log('letti dalla sorgente: ' + contatti.length + ' contatti, ' + categorie.length + ' categorie');

  // 2. UNA sola istruzione nella destinazione: (svuota oppure pretende che sia vuota), riempie,
  //    ricalcola l'impronta e la confronta con quella della sorgente; se non coincide fallisce e
  //    non resta nulla. I dati viaggiano come testo JSON dentro una stringa fra dollari con un
  //    nome che nei dati non compare.
  const dati = JSON.stringify({ contatti, categorie });
  const delimitatore = (base) => { for (let i = 0; i < 50; i++) { const s = '$' + base + '_' + Math.random().toString(36).slice(2, 10) + '$'; if (dati.indexOf(s) < 0) return s; } throw new Error('non trovo un delimitatore sicuro per i dati'); };
  const segnoDati = delimitatore('dati'), segnoCorpo = delimitatore('corpo');
  const sql =
    'do ' + segnoCorpo + '\n' +
    'declare j jsonb := ' + segnoDati + dati + segnoDati + '::jsonb; n1 int; n2 int; f1 text; f2 text;\n' +
    'begin\n' +
    "  if (select valore from public.impostazioni where chiave = 'AMBIENTE') is distinct from '" + D.valore + "' then raise exception 'la destinazione non si dichiara " + D.valore + "'; end if;\n" +
    (svuota
      ? '  truncate public.rubrica_contatti, public.rubrica_categorie restart identity;\n'
      : "  if exists (select 1 from public.rubrica_contatti) or exists (select 1 from public.rubrica_categorie) then raise exception 'DESTINAZIONE NON VUOTA'; end if;\n") +
    "  insert into public.rubrica_categorie (id, nome, ordine) select x.id, x.nome, x.ordine from jsonb_to_recordset(j->'categorie') as x(id int, nome text, ordine int);\n" +
    "  insert into public.rubrica_contatti (id, nome, categoria, numeri, note) select x.id, x.nome, x.categoria, x.numeri, x.note from jsonb_to_recordset(j->'contatti') as x(id int, nome text, categoria text, numeri text, note text);\n" +
    "  perform setval(pg_get_serial_sequence('public.rubrica_categorie', 'id'), greatest((select coalesce(max(id), 0) from public.rubrica_categorie), 1));\n" +
    "  perform setval(pg_get_serial_sequence('public.rubrica_contatti', 'id'), greatest((select coalesce(max(id), 0) from public.rubrica_contatti), 1));\n" +
    '  select count(*), ' + IMPRONTA_CONTATTI + ' into n1, f1 from public.rubrica_contatti;\n' +
    '  select count(*), ' + IMPRONTA_CATEGORIE + ' into n2, f2 from public.rubrica_categorie;\n' +
    '  if n1 <> ' + fc.n + " or f1 <> '" + fc.impronta + "' or n2 <> " + fk.n + " or f2 <> '" + fk.impronta + "' then raise exception 'COPIA NON FEDELE'; end if;\n" +
    'end ' + segnoCorpo + ';';
  await dest(sql, false);
  console.log('destinazione ' + (svuota ? 'svuotata e riempita' : 'riempita (era vuota)') + ' con una sola istruzione, impronta verificata prima di confermare');

  // 3. confronto finale, da fuori
  const uguali = mostra(await firme());
  if (!uguali) { console.log('ERRORE: dopo la copia le due impronte sono diverse (la sorgente è cambiata nel frattempo? rilanciare il confronto)'); process.exit(1); }
  console.log('FATTO: la rubrica ' + D.nome + ' è identica a quella vera, con gli stessi identificativi');
})().catch((e) => { console.log('ERRORE: ' + maschera(String(e && e.message || e))); process.exit(1); });
