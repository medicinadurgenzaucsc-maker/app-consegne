// Ciò che c'è OGGI nelle schede resterebbe intatto passando dal filtro dell'HTML?
//
// Il filtro (docs/js/sanifica.js + DOMPurify) lascia nei campi solo un elenco
// chiuso di tag, attributi, classi e stili. Prima di pubblicare una versione
// che lo introduce o lo cambia va controllato che i contenuti veri non usino
// nulla che verrebbe tolto al primo salvataggio.
//
// Dai database escono SOLO nomi e conteggi, mai il testo: l'analisi gira
// dentro la query (in produzione in sola lettura) e i valori degli attributi
// vengono giudicati lì, con le stesse regole della libreria. Le regole sui
// nomi invece sono quelle VERE: questo script carica sanifica.js e interroga
// il suo gancio, così non c'è un secondo elenco da tenere allineato.
//
//   node collaudo/strumenti/confronta-filtro.js [produzione|collaudo] [--backup N]
//
// --backup N: quanti backup recenti di «archivio» esaminare oltre alle schede
// (predefinito 30; 0 = nessuno). Esce con codice 1 se qualcosa verrebbe tolto.
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { collaudo, produzione } = require('./sb.js');

const RICCHI = { note_terapia: 'NoteTerapia', diaria: 'Diaria', da_fare: 'DaFare', piano_terapeutico: 'PianoTerapeutico', esami_colturali: 'EsamiColturali', allergie: 'Allergie' };
const SEMPLICI = ['nome', 'diagnosi', 'eta', 'codice_sanitario', 'ossigeno', 'vitto', 'tipologia_letto', 'letto'];
// Attributi il cui valore la libreria non tratta mai come indirizzo (elenco
// interno di DOMPurify 3.4.16), a cui sanifica.js aggiunge i suoi.
const NON_INDIRIZZI_DI_SERIE = ['alt', 'class', 'for', 'id', 'label', 'name', 'pattern', 'placeholder', 'role', 'summary', 'title', 'value', 'style', 'xmlns'];

// ── Le regole vere del filtro ────────────────────────────────────────────
function regoleDelFiltro() {
  const sorgente = fs.readFileSync(path.join(__dirname, '../../docs/js/sanifica.js'), 'utf8');
  let gancio = null, cfg = null;
  const libreriaFinta = function () {
    return { addHook: (nome, fn) => { if (nome === 'afterSanitizeAttributes') gancio = fn; }, sanitize: (s, c) => { cfg = c; return s; }, removed: [] };
  };
  libreriaFinta.isSupported = true;
  const window = { DOMPurify: libreriaFinta, location: { href: 'http://localhost/' } };
  vm.runInNewContext(sorgente, { window, URL, Map, console });
  window._pulisciHtml('<b>x</b>'); // fa arrivare la configurazione alla libreria finta
  if (!gancio || !cfg) throw new Error('sanifica.js non ha esposto le sue regole: lo script va adeguato');
  const nodo = (attr) => { const m = new Map(Object.entries(attr)); return { nodeType: 1, getAttribute: (n) => (m.has(n) ? m.get(n) : null), setAttribute: (n, v) => m.set(n, String(v)), removeAttribute: (n) => m.delete(n), m }; };
  const dopo = (attr) => { const n = nodo(attr); gancio(n); return n.m; };
  return {
    tag: new Set(cfg.ALLOWED_TAGS.map((t) => t.toLowerCase())),
    attributi: new Set(cfg.ALLOWED_ATTR.map((t) => t.toLowerCase())),
    nonIndirizzi: new Set(NON_INDIRIZZI_DI_SERIE.concat(cfg.ADD_URI_SAFE_ATTR || []).map((t) => t.toLowerCase())),
    classeResta: (c) => dopo({ class: c }).get('class') === c,
    stileResta: (p) => dopo({ style: p + ': 1' }).has('style'),
  };
}

// ── L'analisi, dentro il database ────────────────────────────────────────
// VALORE_VIETATO e COLORE ricalcano le due espressioni di sanifica.js: se
// cambiano lì vanno cambiate qui (sono le uniche regole duplicate).
function sqlAnalisi(sorgente) {
  const spazi = String.raw`'[\u0001-   ᠎ -  　]'`;
  return String.raw`
with src as (select distinct campo, html from (${sorgente}) u where html <> ''),
tag as (
  select lower(m[1]) as nome, coalesce(m[2], '') as attr
  from src s, regexp_matches(s.html, '<([a-zA-Z][a-zA-Z0-9:-]*)((?:\s[^<>]*)?)>', 'g') m
),
att as (
  select lower(a[1]) as nome,
         replace(replace(replace(replace(replace(replace(a[2], '&quot;', '"'), '&#39;', ''''), '&lt;', '<'), '&gt;', '>'), '&nbsp;', U&'\00A0'), '&amp;', '&') as valore
  from tag t, regexp_matches(t.attr, '([a-zA-Z_:][-a-zA-Z0-9_:.]*)\s*=\s*"([^"]*)"', 'g') a
),
dich as (
  select btrim(d) as d from att, regexp_split_to_table(att.valore, ';') d where att.nome = 'style' and btrim(d) <> ''
),
prop as (
  select case when position(':' in d) > 1 then lower(btrim(split_part(d, ':', 1))) else '' end as nome,
         case when position(':' in d) > 1 then btrim(substr(d, position(':' in d) + 1)) else '' end as valore
  from dich
)
select jsonb_build_object(
  'campi', (select count(*) from src),
  'tag', (select coalesce(jsonb_object_agg(nome, n), '{}'::jsonb) from (select nome, count(*) n from tag group by 1) x),
  'attributi', (select coalesce(jsonb_object_agg(nome, n), '{}'::jsonb) from (select nome, count(*) n from att where nome ~ '^[a-z_:][-a-z0-9_:.]{0,40}$' group by 1) x),
  'classi', (select coalesce(jsonb_object_agg(k, n), '{}'::jsonb) from (select k, count(*) n from att, regexp_split_to_table(btrim(att.valore), '\s+') k where att.nome = 'class' and k <> '' and k ~ '^[-a-zA-Z0-9_]{1,60}$' group by 1) x),
  'classi_strane', (select count(*) from att, regexp_split_to_table(btrim(att.valore), '\s+') k where att.nome = 'class' and k <> '' and k !~ '^[-a-zA-Z0-9_]{1,60}$'),
  'stili', (select coalesce(jsonb_object_agg(nome, n), '{}'::jsonb) from (select nome, count(*) n from prop where nome ~ '^[-a-z]{1,40}$' group by 1) x),
  'stili_senza_nome', (select count(*) from prop where nome !~ '^[-a-z]{1,40}$'),
  'stili_valore_vuoto', (select coalesce(jsonb_object_agg(nome, n), '{}'::jsonb) from (select nome, count(*) n from prop where nome ~ '^[-a-z]{1,40}$' and valore = '' group by 1) x),
  'stili_valore_vietato', (select coalesce(jsonb_object_agg(nome, n), '{}'::jsonb) from (select nome, count(*) n from prop where nome ~ '^[-a-z]{1,40}$' and valore ~* 'url\s*\(|expression\s*\(|image\s*\(|image-set|cross-fade|element\s*\(|javascript\s*:|behaviou?r\s*:|-moz-binding|@|\\|<|>|/\*' group by 1) x),
  'valori_schema_ignoto', (select coalesce(jsonb_object_agg(nome, n), '{}'::jsonb) from (select nome, count(*) n from att
      where regexp_replace(valore, ${spazi}, '', 'g') ~* '^[a-z+.\-]+:'
        and regexp_replace(valore, ${spazi}, '', 'g') !~* '^(?:(?:f|ht)tps?|mailto|tel|callto|sms|cid|xmpp|matrix):' group by 1) x),
  'valori_con_chiusure', (select coalesce(jsonb_object_agg(nome, n), '{}'::jsonb) from (select nome, count(*) n from att
      where valore ~* '((--!?|\])>)|</(style|script|title|xmp|textarea|noscript|iframe|noembed|noframes)' group by 1) x),
  'contenteditable_altri', (select count(*) from att where nome = 'contenteditable' and valore not in ('false', 'true', '')),
  'colori_non_validi', (select coalesce(jsonb_object_agg(nome, n), '{}'::jsonb) from (select nome, count(*) n from att
      where nome in ('fill', 'stroke', 'color') and btrim(valore) !~* '^(?:none|currentcolor|transparent|#[0-9a-f]{3,8}|[a-z]{3,30}|(?:rgb|rgba|hsl|hsla)\(\s*[0-9.,%\s/a-z-]{1,60}\))$' group by 1) x),
  'note_html', (select count(*) from src s, regexp_matches(s.html, '<!--', 'g') m),
  'tag_virgolette_dispari', (select count(*) from tag where (length(attr) - length(replace(attr, '"', ''))) % 2 = 1),
  'attributi_senza_virgolette', (select count(*) from tag t, regexp_matches(regexp_replace(t.attr, '"[^"]*"', '""', 'g'), '=\s*[^"\s]', 'g') a)
) as esito`;
}
const sorgenteSchede = Object.keys(RICCHI).map((c) => "select '" + c + "'::text as campo, coalesce(" + c + ", '') as html from public.consegne").join('\n  union all ');
const sorgenteBackup = (n) => Object.values(RICCHI).map((k) => "select '" + k + "'::text as campo, coalesce(p->>'" + k + "', '') as html from (select dati from public.archivio order by ts desc limit " + Number(n) + ") a, jsonb_array_elements(a.dati) p").join('\n  union all ');
const sqlSemplici = 'select ' + SEMPLICI.map((c) => "(select count(*) from public.consegne where coalesce(" + c + "::text, '') ~ '<[a-zA-Z!/]') as tag_" + c + ", (select count(*) from public.consegne where coalesce(" + c + "::text, '') ~ '&(#[0-9]+|#x[0-9a-fA-F]+|[a-zA-Z]+);') as ent_" + c).join(',\n  ');

// ── Confronto ────────────────────────────────────────────────────────────
function giudica(nome, e, R) {
  const problemi = [];
  const elenco = (mappa, chiavi) => chiavi.sort().map((x) => x + ' ×' + mappa[x]).join(', ');
  const fuori = (titolo, mappa, resta) => { const k = Object.keys(mappa || {}).filter((x) => !resta(x)); if (k.length) problemi.push(titolo + ': ' + elenco(mappa, k)); };
  const tutti = (titolo, mappa, filtro) => { const k = Object.keys(mappa || {}).filter(filtro || (() => true)); if (k.length) problemi.push(titolo + ': ' + elenco(mappa, k)); };
  fuori('tag che il filtro toglierebbe (il testo dentro resta)', e.tag, (t) => R.tag.has(t));
  fuori('attributi che il filtro toglierebbe', e.attributi, (a) => R.attributi.has(a));
  fuori('classi che il filtro toglierebbe', e.classi, R.classeResta);
  fuori('proprietà di stile che il filtro toglierebbe', e.stili, R.stileResta);
  tutti('stili senza valore (tolti)', e.stili_valore_vuoto);
  tutti('stili con un valore vietato (tolti)', e.stili_valore_vietato, (p) => R.stileResta(p));
  tutti('attributi il cui valore verrebbe scambiato per un indirizzo e tolto', e.valori_schema_ignoto, (a) => R.attributi.has(a) && !R.nonIndirizzi.has(a));
  tutti('attributi tolti perché il valore contiene «-->», «]>» o una chiusura di tag', e.valori_con_chiusure, (a) => R.attributi.has(a));
  tutti('colori non validi (attributo tolto)', e.colori_non_validi);
  if (Number(e.contenteditable_altri)) problemi.push('contenteditable con valori imprevisti: ×' + e.contenteditable_altri);
  if (Number(e.classi_strane)) problemi.push('classi con caratteri imprevisti (tolte): ×' + e.classi_strane);
  if (Number(e.stili_senza_nome)) problemi.push('dichiarazioni di stile senza nome leggibile (tolte): ×' + e.stili_senza_nome);
  if (Number(e.note_html)) problemi.push('commenti HTML (tolti, invisibili): ×' + e.note_html);
  const incerti = [];
  if (Number(e.tag_virgolette_dispari)) incerti.push('tag con virgolette non appaiate: ×' + e.tag_virgolette_dispari);
  if (Number(e.attributi_senza_virgolette)) incerti.push('attributi senza virgolette doppie: ×' + e.attributi_senza_virgolette);
  const n = (m) => Object.values(m || {}).reduce((s, x) => s + Number(x), 0);
  console.log('\n=== ' + nome + ' — ' + e.campi + ' campi con contenuto, ' + n(e.tag) + ' tag, ' + n(e.attributi) + ' attributi, ' + n(e.classi) + ' classi, ' + n(e.stili) + ' dichiarazioni di stile');
  console.log('    in uso: tag ' + Object.keys(e.tag || {}).length + ' diversi · attributi ' + Object.keys(e.attributi || {}).length + ' · classi ' + Object.keys(e.classi || {}).length + ' · proprietà di stile ' + Object.keys(e.stili || {}).length);
  if (problemi.length) problemi.forEach((p) => console.log('    KO  ' + p)); else console.log('    ok  il filtro non toglierebbe nulla');
  incerti.forEach((p) => console.log('    ??  ' + p + ' (analisi non affidabile su quei tag)'));
  return problemi.length;
}

// ── Lo strumento vede davvero ciò che deve vedere? ───────────────────────
// Un contenuto costruito apposta, analizzato dal database di collaudo: ogni
// regola deve scattare sul suo pezzo, e nessuna su ciò che è lecito.
async function autoprova(R) {
  const html = '<div class="position-fixed pa-item" style="position: fixed; color: red; background: url(x)" data-foo="a"'
    + ' data-isolamento="da contatto: KPC" data-ts="12:30" title="peggiora --&gt; UTI" onclick="x()" contenteditable="plaintext-only">'
    + '<iframe src="x"></iframe><font color="url(x)" face="nota: x">t</font>'
    + '<svg viewBox="0 0 1 1"><path d="M1 1" fill="red" stroke="da: x"></path></svg><!-- nota --></div>';
  const e = (await collaudo().query(sqlAnalisi("select 'prova'::text as campo, $prova$" + html + "$prova$::text as html")))[0].esito;
  const attesi = {
    'tag iframe': !R.tag.has('iframe') && e.tag.iframe === 1,
    'tag div, font, svg, path ammessi': ['div', 'font', 'svg', 'path'].every((t) => R.tag.has(t) && e.tag[t] === 1),
    'attributi data-foo, onclick, src da togliere': ['data-foo', 'onclick', 'src'].every((a) => !R.attributi.has(a) && e.attributi[a] === 1),
    'attributi leciti riconosciuti': ['class', 'style', 'data-isolamento', 'data-ts', 'title', 'viewbox', 'd', 'fill'].every((a) => R.attributi.has(a) && e.attributi[a] >= 1),
    'classe position-fixed da togliere, pa-item no': !R.classeResta('position-fixed') && R.classeResta('pa-item') && e.classi['position-fixed'] === 1 && e.classi['pa-item'] === 1,
    'stile position e background da togliere, color no': !R.stileResta('position') && !R.stileResta('background') && R.stileResta('color') && e.stili.position === 1 && e.stili.color === 1,
    'testo libero negli attributi dell\'app non scambiato per indirizzo': e.valori_schema_ignoto['data-isolamento'] === 1 && R.nonIndirizzi.has('data-isolamento') && R.nonIndirizzi.has('face'),
    'valore «da: x» in un attributo qualunque scambiato per indirizzo': e.valori_schema_ignoto.stroke === 1 && !R.nonIndirizzi.has('stroke'),
    'orario 12:30 non scambiato per indirizzo': !('data-ts' in e.valori_schema_ignoto),
    'freccia «-->» dentro un attributo vista': e.valori_con_chiusure.title === 1,
    'contenteditable imprevisto visto': Number(e.contenteditable_altri) === 1,
    'colore non valido visto (font color e stroke), fill rosso no': e.colori_non_validi.color === 1 && e.colori_non_validi.stroke === 1 && !('fill' in e.colori_non_validi),
    'commento HTML visto': Number(e.note_html) === 1,
    'nessun tag giudicato illeggibile': Number(e.tag_virgolette_dispari) === 0 && Number(e.attributi_senza_virgolette) === 0,
  };
  let ko = 0;
  Object.keys(attesi).forEach((k) => { if (!attesi[k]) ko++; console.log((attesi[k] ? 'ok  ' : 'KO  ') + k); });
  console.log(ko ? '\n' + ko + ' controlli dell\'autoprova falliti: lo strumento NON è affidabile' : '\nautoprova superata: lo strumento vede ciò che deve vedere');
  return ko;
}

(async () => {
  if (process.argv.indexOf('--autoprova') >= 0) { process.exitCode = (await autoprova(regoleDelFiltro())) ? 1 : 0; return; }
  const quale = (process.argv[2] || 'collaudo').toLowerCase();
  const iB = process.argv.indexOf('--backup');
  const nBackup = iB >= 0 ? Number(process.argv[iB + 1]) : 30;
  const leggi = quale === 'produzione' ? produzione().leggi : collaudo().query;
  const R = regoleDelFiltro();
  console.log('regole lette da docs/js/sanifica.js: ' + R.tag.size + ' tag, ' + R.attributi.size + ' attributi ammessi · database: ' + quale.toUpperCase());
  let ko = 0;
  ko += giudica('schede di oggi', (await leggi(sqlAnalisi(sorgenteSchede)))[0].esito, R);
  if (nBackup > 0) ko += giudica('ultimi ' + nBackup + ' backup', (await leggi(sqlAnalisi(sorgenteBackup(nBackup))))[0].esito, R);
  const s = (await leggi(sqlSemplici))[0];
  const conTag = SEMPLICI.filter((c) => Number(s['tag_' + c])).map((c) => c + ' ×' + s['tag_' + c]);
  const conEnt = SEMPLICI.filter((c) => Number(s['ent_' + c])).map((c) => c + ' ×' + s['ent_' + c]);
  console.log('\n=== campi di solo testo (ora mostrati lettera per lettera, senza interpretare nulla)');
  console.log('    ' + (conTag.length ? '!!  contengono qualcosa che somiglia a un tag: ' + conTag.join(', ') : 'ok  nessuno contiene tag'));
  console.log('    ' + (conEnt.length ? '!!  contengono entità (es. &amp;) che si vedrebbero scritte così: ' + conEnt.join(', ') : 'ok  nessuno contiene entità HTML'));
  console.log('\n' + (ko ? ko + ' DIFFERENZE: il filtro cambierebbe i contenuti' : 'nessuna differenza: i contenuti passano dal filtro intatti'));
  process.exitCode = ko ? 1 : 0;
})().catch((e) => { console.log('ERRORE: ' + String(e.message).slice(0, 700)); process.exit(2); });
