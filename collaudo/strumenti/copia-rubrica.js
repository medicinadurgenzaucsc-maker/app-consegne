// Copia l'elenco della rubrica telefonica dal suo vecchio database (progetto
// Supabase «rubrica-gemelli», account del reparto) nelle tabelle rubrica_ del
// database delle consegne.
//
//   node collaudo/strumenti/copia-rubrica.js collaudo            copia nel collaudo
//   node collaudo/strumenti/copia-rubrica.js collaudo confronta  dice solo se le due copie coincidono
//
// La sorgente si legge soltanto (transazione di sola lettura imposta dal
// server). La destinazione viene svuotata e riempita da UNA sola istruzione,
// che alla fine ricalcola conteggi e impronta e li confronta con quelli della
// sorgente: se non coincidono l'istruzione fallisce e la destinazione resta
// com'era. O entra tutto, con gli stessi identificativi, o non cambia nulla.
//
// Nomi e numeri non vengono mai stampati: escono solo conteggi e impronte, e
// di un errore del database solo lo stato, il codice e i nomi di tabelle,
// colonne e vincoli.
// Nel collaudo ci sono i numeri VERI per decisione di Stefano dell'08/10/2026:
// sono numeri di telefono dell'ospedale, non dati di pazienti, e servono a
// provare la rubrica com'è davvero.
//
// La produzione non è prevista qui: quando la rubrica entrerà nel database
// delle consegne del reparto la copia si farà con l'ok di chi gestisce l'app,
// e questo strumento andrà esteso allora.
const fs = require('fs');
const os = require('os');
const path = require('path');
const { REF_COLLAUDO, maschera } = require('./sb.js');

const REF_RUBRICA = 'nbbekxuvuarxkuvvvgbi';
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
// codice, nomi di vincoli, tabelle e colonne, e i nostri due messaggi fissi.
function erroreSenzaDati(ref, stato, testo) {
  const codice = (/ERROR:\s+([0-9A-Z]{5})\b/.exec(testo) || /"code"\s*:\s*"([0-9A-Z]{5})"/.exec(testo) || [])[1] || 'senza codice';
  const nomi = [];
  String(testo).replace(/(constraint|relation|column|function) \\?"([A-Za-z0-9_.]+)/g, (tutto, tipo, nome) => { nomi.push(tipo + ' ' + nome); return tutto; });
  const nostro = (/(COPIA NON FEDELE|la destinazione non si dichiara collaudo)/.exec(testo) || [])[1];
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
  const [ambiente, azione] = process.argv.slice(2);
  if (ambiente !== 'collaudo' || (azione && azione !== 'confronta') || process.argv.length > 4) {
    console.log('uso: copia-rubrica.js collaudo [confronta]');
    process.exit(1);
  }
  // la chiave dell'account del reparto vede anche il progetto della rubrica; lì si legge soltanto
  const tokSorgente = token('supabase', null);
  const tokDest = token('supabase-collaudo', REF_COLLAUDO);
  const sorgente = (sql) => interroga(REF_RUBRICA, tokSorgente, sql, true);
  const dest = (sql, lettura) => interroga(REF_COLLAUDO, tokDest, sql, !!lettura);

  const amb = await dest("select valore from public.impostazioni where chiave = 'AMBIENTE'", true);
  if (!amb[0] || amb[0].valore !== 'collaudo') throw new Error('il database di destinazione non si dichiara «collaudo»: non tocco nulla');

  const firme = async () => ({
    sc: (await sorgente(firma(IMPRONTA_CONTATTI, 'contatti')))[0], sk: (await sorgente(firma(IMPRONTA_CATEGORIE, 'categorie')))[0],
    dc: (await dest(firma(IMPRONTA_CONTATTI, 'rubrica_contatti'), true))[0], dk: (await dest(firma(IMPRONTA_CATEGORIE, 'rubrica_categorie'), true))[0],
  });
  const mostra = (f) => {
    console.log('  sorgente:     ' + f.sc.n + ' contatti (' + f.sc.impronta.slice(0, 12) + '), ' + f.sk.n + ' categorie (' + f.sk.impronta.slice(0, 12) + ')');
    console.log('  destinazione: ' + f.dc.n + ' contatti (' + f.dc.impronta.slice(0, 12) + '), ' + f.dk.n + ' categorie (' + f.dk.impronta.slice(0, 12) + ')');
    return f.sc.n === f.dc.n && f.sk.n === f.dk.n && f.sc.impronta === f.dc.impronta && f.sk.impronta === f.dk.impronta;
  };
  if (azione === 'confronta') {
    const uguali = mostra(await firme());
    console.log(uguali ? 'LE DUE COPIE COINCIDONO' : 'LE DUE COPIE SONO DIVERSE');
    process.exit(uguali ? 0 : 1);
  }

  // 1. si legge la sorgente (in memoria: nulla viene stampato né scritto su disco) e, DOPO, la sua impronta
  const contatti = await sorgente('select id, nome, categoria, numeri, note from public.contatti order by id');
  const categorie = await sorgente('select id, nome, ordine from public.categorie order by id');
  const fc = (await sorgente(firma(IMPRONTA_CONTATTI, 'contatti')))[0], fk = (await sorgente(firma(IMPRONTA_CATEGORIE, 'categorie')))[0];
  if (!contatti.length || !categorie.length) throw new Error('la sorgente risulta vuota (' + contatti.length + ' contatti, ' + categorie.length + ' categorie): non svuoto la destinazione');
  if (contatti.length !== fc.n || categorie.length !== fk.n) throw new Error('la sorgente è cambiata durante la lettura: rilanciare');
  if (!/^[0-9a-f]{32}$/.test(fc.impronta) || !/^[0-9a-f]{32}$/.test(fk.impronta)) throw new Error('impronta della sorgente non valida');
  console.log('letti dalla sorgente: ' + contatti.length + ' contatti, ' + categorie.length + ' categorie');

  // 2. UNA sola istruzione nella destinazione: svuota, riempie, ricalcola l'impronta e la
  //    confronta con quella della sorgente; se non coincide fallisce e non resta nulla.
  //    I dati viaggiano come testo JSON dentro una stringa fra dollari con un nome che nei dati non compare.
  const dati = JSON.stringify({ contatti, categorie });
  const delimitatore = (base) => { for (let i = 0; i < 50; i++) { const s = '$' + base + '_' + Math.random().toString(36).slice(2, 10) + '$'; if (dati.indexOf(s) < 0) return s; } throw new Error('non trovo un delimitatore sicuro per i dati'); };
  const segnoDati = delimitatore('dati'), segnoCorpo = delimitatore('corpo');
  const sql =
    'do ' + segnoCorpo + '\n' +
    'declare j jsonb := ' + segnoDati + dati + segnoDati + '::jsonb; n1 int; n2 int; f1 text; f2 text;\n' +
    'begin\n' +
    "  if (select valore from public.impostazioni where chiave = 'AMBIENTE') is distinct from 'collaudo' then raise exception 'la destinazione non si dichiara collaudo'; end if;\n" +
    '  truncate public.rubrica_contatti, public.rubrica_categorie restart identity;\n' +
    "  insert into public.rubrica_categorie (id, nome, ordine) select x.id, x.nome, x.ordine from jsonb_to_recordset(j->'categorie') as x(id int, nome text, ordine int);\n" +
    "  insert into public.rubrica_contatti (id, nome, categoria, numeri, note) select x.id, x.nome, x.categoria, x.numeri, x.note from jsonb_to_recordset(j->'contatti') as x(id int, nome text, categoria text, numeri text, note text);\n" +
    "  perform setval(pg_get_serial_sequence('public.rubrica_categorie', 'id'), greatest((select coalesce(max(id), 0) from public.rubrica_categorie), 1));\n" +
    "  perform setval(pg_get_serial_sequence('public.rubrica_contatti', 'id'), greatest((select coalesce(max(id), 0) from public.rubrica_contatti), 1));\n" +
    '  select count(*), ' + IMPRONTA_CONTATTI + ' into n1, f1 from public.rubrica_contatti;\n' +
    '  select count(*), ' + IMPRONTA_CATEGORIE + ' into n2, f2 from public.rubrica_categorie;\n' +
    '  if n1 <> ' + fc.n + " or f1 <> '" + fc.impronta + "' or n2 <> " + fk.n + " or f2 <> '" + fk.impronta + "' then raise exception 'COPIA NON FEDELE'; end if;\n" +
    'end ' + segnoCorpo + ';';
  await dest(sql, false);
  console.log('destinazione svuotata e riempita con una sola istruzione, impronta verificata prima di confermare');

  // 3. confronto finale, da fuori
  const uguali = mostra(await firme());
  if (!uguali) { console.log('ERRORE: dopo la copia le due impronte sono diverse (la sorgente è cambiata nel frattempo? rilanciare)'); process.exit(1); }
  console.log('FATTO: la rubrica del collaudo è identica a quella vera, con gli stessi identificativi');
})().catch((e) => { console.log('ERRORE: ' + maschera(String(e && e.message || e))); process.exit(1); });
