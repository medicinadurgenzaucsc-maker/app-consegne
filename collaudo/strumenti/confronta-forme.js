// Confronta la FORMA dei contenuti (non i contenuti) fra produzione e collaudo:
// per ogni campo HTML delle schede, l'insieme dei tag e delle classi usate e le
// chiavi/forma dei documenti di laboratorio. Dalla produzione si leggono solo
// nomi di tag, nomi di classi e conteggi: mai il testo.
const { collaudo, produzione } = require('./sb.js');

const CAMPI = ['note_terapia', 'diaria', 'da_fare', 'piano_terapeutico', 'esami_colturali'];

function sqlForme(campo) {
  return `
    select
      (select coalesce(jsonb_agg(distinct m[1] order by m[1]), '[]'::jsonb)
         from public.consegne c, regexp_matches(coalesce(c.${campo}, ''), 'class="([^"]*)"', 'g') m) as classi,
      (select coalesce(jsonb_agg(distinct lower(m[1]) order by lower(m[1])), '[]'::jsonb)
         from public.consegne c, regexp_matches(coalesce(c.${campo}, ''), '<([a-zA-Z][a-zA-Z0-9]*)', 'g') m) as tag,
      (select coalesce(jsonb_agg(distinct m[1] order by m[1]), '[]'::jsonb)
         from public.consegne c, regexp_matches(coalesce(c.${campo}, ''), '\\s(data-[a-z-]+)=', 'g') m) as attributi_data,
      (select count(*) from public.consegne c where coalesce(c.${campo}, '') <> '') as non_vuoti,
      (select round(avg(length(c.${campo}))) from public.consegne c where coalesce(c.nome, '') <> '') as lunghezza_media`;
}
const sqlLab = `
  select
    (select count(*) from public.lab_esami) as righe,
    (select coalesce(jsonb_agg(distinct k order by k), '[]'::jsonb) from public.lab_esami l, jsonb_object_keys(l.dati) k) as chiavi_documento,
    (select coalesce(jsonb_agg(distinct k order by k), '[]'::jsonb) from public.lab_esami l, jsonb_each(l.dati->'esami') e, jsonb_object_keys(e.value) k) as chiavi_esame,
    (select round(avg(n_esami)) from public.lab_esami) as esami_medi,
    (select round(avg(octet_length(dati::text))) from public.lab_esami) as byte_medi,
    (select count(*) from public.lab_esami where ultimo_esame is not null) as con_ultimo_esame,
    (select string_agg(distinct regexp_replace(d, '[0-9]', '9', 'g'), ' | ')
       from public.lab_esami l, jsonb_each(l.dati->'esami') e, jsonb_object_keys(e.value->'v') d) as formato_date`;

(async () => {
  const p = produzione(), c = collaudo();
  let differenze = 0;
  for (const campo of CAMPI) {
    const a = (await p.leggi(sqlForme(campo)))[0], b = (await c.query(sqlForme(campo)))[0];
    console.log('=== ' + campo + ' — non vuoti: produzione ' + a.non_vuoti + ', collaudo ' + b.non_vuoti + ' | lunghezza media: ' + a.lunghezza_media + ' / ' + b.lunghezza_media);
    for (const k of ['classi', 'tag', 'attributi_data']) {
      const soloP = a[k].filter((x) => !b[k].includes(x)), soloC = b[k].filter((x) => !a[k].includes(x));
      console.log('  ' + k.padEnd(15) + 'comuni: ' + a[k].filter((x) => b[k].includes(x)).join(', '));
      if (soloP.length) console.log('  ' + ''.padEnd(15) + 'solo in PRODUZIONE: ' + soloP.join(', '));
      if (soloC.length) { console.log('  ' + ''.padEnd(15) + 'solo in COLLAUDO:   ' + soloC.join(', ')); differenze += soloC.length; }
    }
  }
  const la = (await p.leggi(sqlLab))[0], lb = (await c.query(sqlLab))[0];
  console.log('=== lab_esami');
  for (const k of Object.keys(la)) console.log('  ' + k.padEnd(18) + 'produzione: ' + JSON.stringify(la[k]) + '  |  collaudo: ' + JSON.stringify(lb[k]));
  console.log(differenze ? ('ATTENZIONE: ' + differenze + ' elementi di forma presenti solo nel collaudo') : 'FORMA: nel collaudo non c\'è nulla che la produzione non usi già');
})().catch((e) => { console.log('ERRORE: ' + e.message); process.exit(1); });
