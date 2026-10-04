// Inventario del MARKUP presente nei campi HTML delle schede: nomi di tag,
// nomi di attributi, classi, proprietà di stile, schemi dei link. Serve a
// tarare (e a ricontrollare) il filtro dell'HTML: tutto ciò che le schede vere
// usano deve restare, tutto il resto si può togliere.
// Dalla PRODUZIONE escono solo nomi e conteggi, mai il testo: i valori degli
// attributi vengono svuotati dentro la query, prima di estrarre i nomi.
//
//   node collaudo/strumenti/inventario-markup.js [produzione|collaudo|entrambi] [--json file]
const fs = require('fs');
const { collaudo, produzione } = require('./sb.js');

const RICCHI = ['note_terapia', 'diaria', 'da_fare', 'piano_terapeutico', 'esami_colturali', 'allergie'];
const SEMPLICI = ['nome', 'diagnosi', 'eta', 'codice_sanitario', 'ossigeno', 'vitto', 'dimissibile', 'sesso', 'data_nascita', 'data_ricovero', 'tipologia_letto', 'letto'];
// proprietà di stile di cui interessa anche il VALORE (sono parole chiave CSS, non testo)
const VALORI_STILE = ['position', 'display', 'float', 'z-index', 'visibility', 'overflow', 'top', 'left', 'right', 'bottom', 'inset'];

function sqlCampo(campo) {
  const col = "coalesce(c." + campo + ", '')";
  return `
    with tag as (
      select lower(m[1]) as nome, coalesce(m[2], '') as attr
      from public.consegne c, regexp_matches(${col}, '<([a-zA-Z][a-zA-Z0-9:-]*)((?:\\s[^<>]*)?)>', 'g') m
    ), att as (
      select t.nome as tag, lower(n[1]) as nome
      from tag t,
           regexp_matches(regexp_replace(regexp_replace(t.attr, '"[^"]*"', '""', 'g'), '''[^'']*''', '''''', 'g'),
                          '([a-zA-Z_:][-a-zA-Z0-9_:.]*)\\s*=', 'g') n
      where n[1] ~ '^[a-zA-Z_:][-a-zA-Z0-9_:.]{0,30}$'
    ), stile as (
      select s[1] as valore from public.consegne c, regexp_matches(${col}, 'style="([^"]*)"', 'g') s
    ), prop as (
      select lower(trim(split_part(d, ':', 1))) as nome, lower(trim(split_part(d, ':', 2))) as valore
      from stile, regexp_split_to_table(stile.valore, ';') d
      where position(':' in d) > 0
    )
    select
      (select count(*) from public.consegne c where ${col} <> '') as non_vuoti,
      (select coalesce(jsonb_object_agg(nome, n), '{}'::jsonb) from (select nome, count(*) as n from tag group by 1) x) as tag,
      (select coalesce(jsonb_object_agg(k, n), '{}'::jsonb) from (select tag || ' ' || nome as k, count(*) as n from att group by 1) x) as attributi,
      (select coalesce(jsonb_object_agg(k, n), '{}'::jsonb) from (
          select k, count(*) as n from public.consegne c, regexp_matches(${col}, 'class="([^"]*)"', 'g') m, regexp_split_to_table(trim(m[1]), '\\s+') k where k <> '' group by 1) x) as classi,
      (select coalesce(jsonb_object_agg(nome, n), '{}'::jsonb) from (select nome, count(*) as n from prop where nome ~ '^[-a-z]{1,40}$' group by 1) x) as stili,
      (select coalesce(jsonb_object_agg(k, n), '{}'::jsonb) from (
          select nome || ': ' || left(valore, 40) as k, count(*) as n from prop where nome in (${VALORI_STILE.map((v) => "'" + v + "'").join(', ')}) group by 1) x) as stili_di_posizione,
      (select count(*) from stile where valore ~* 'url\\s*\\(') as stili_con_url,
      (select count(*) from stile where valore ~* 'expression\\s*\\(|javascript:|behavior\\s*:|-moz-binding') as stili_sospetti,
      (select coalesce(jsonb_object_agg(k, n), '{}'::jsonb) from (
          select lower(m[1]) as k, count(*) as n from public.consegne c, regexp_matches(${col}, '(?:href|src|action|formaction|data|xlink:href)\\s*=\\s*["'']?\\s*([a-zA-Z][a-zA-Z0-9+.-]{0,15}):', 'g') m group by 1) x) as schemi_url,
      (select count(*) from public.consegne c, regexp_matches(${col}, '<!--', 'g') m) as commenti,
      (select count(*) from public.consegne c, regexp_matches(${col}, '\\son[a-zA-Z]+\\s*=', 'g') m) as attributi_evento`;
}
const sqlSemplici = `
  select ${SEMPLICI.map((c) => "(select count(*) from public.consegne c where coalesce(c." + c + "::text, '') ~ '<[a-zA-Z!/]') as " + c).join(',\n         ')}`;
const sqlLab = `
  select
    (select count(*) from public.lab_esami l where l.dati::text ~ '<[a-zA-Z!/]') as documenti_con_tag,
    (select count(*) from public.lab_esami l where coalesce(l.paziente, '') ~ '<[a-zA-Z!/]') as paziente_con_tag`;

async function inventario(leggi) {
  const out = { campi: {} };
  for (const campo of RICCHI) out.campi[campo] = (await leggi(sqlCampo(campo)))[0];
  out.campi_semplici_con_tag = (await leggi(sqlSemplici))[0];
  out.laboratorio = (await leggi(sqlLab))[0];
  return out;
}
function somma(inv, chiave) {
  const tot = {};
  Object.values(inv.campi).forEach((c) => Object.entries(c[chiave] || {}).forEach(([k, n]) => { tot[k] = (tot[k] || 0) + Number(n); }));
  return tot;
}
function stampa(nome, inv) {
  const riga = (titolo, o) => console.log('  ' + titolo.padEnd(22) + (Object.keys(o).length ? Object.keys(o).sort((a, b) => o[b] - o[a] || a.localeCompare(b)).map((k) => k + ' ×' + o[k]).join('  ·  ') : '(nessuno)'));
  console.log('=== ' + nome.toUpperCase());
  riga('tag', somma(inv, 'tag'));
  riga('attributi (tag nome)', somma(inv, 'attributi'));
  riga('classi', somma(inv, 'classi'));
  riga('proprietà di stile', somma(inv, 'stili'));
  riga('stili di posizione', somma(inv, 'stili_di_posizione'));
  riga('schemi negli URL', somma(inv, 'schemi_url'));
  const n = (k) => Object.values(inv.campi).reduce((s, c) => s + Number(c[k] || 0), 0);
  console.log('  stili con url(): ' + n('stili_con_url') + ' · stili sospetti: ' + n('stili_sospetti') + ' · commenti HTML: ' + n('commenti') + ' · attributi evento (on…=): ' + n('attributi_evento'));
  console.log('  campi di solo testo che contengono tag: ' + JSON.stringify(inv.campi_semplici_con_tag));
  console.log('  laboratorio: ' + JSON.stringify(inv.laboratorio));
}

(async () => {
  const quale = (process.argv[2] || 'entrambi').toLowerCase();
  const fileJson = process.argv.indexOf('--json') >= 0 ? process.argv[process.argv.indexOf('--json') + 1] : null;
  const esito = {};
  if (quale !== 'collaudo') { esito.produzione = await inventario(produzione().leggi); stampa('produzione', esito.produzione); }
  if (quale !== 'produzione') { esito.collaudo = await inventario(collaudo().query); stampa('collaudo', esito.collaudo); }
  if (fileJson) { fs.writeFileSync(fileJson, JSON.stringify(esito, null, 1)); console.log('scritto ' + fileJson); }
})().catch((e) => { console.log('ERRORE: ' + e.message); process.exit(1); });
