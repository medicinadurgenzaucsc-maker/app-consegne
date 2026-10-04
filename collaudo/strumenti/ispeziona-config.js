// Guarda la CONFIGURAZIONE della produzione senza mostrarne i contenuti:
// per ogni chiave di impostazioni solo lunghezza e natura del valore.
const { produzione } = require('./sb.js');

(async () => {
  const p = produzione();
  const imp = await p.leggi(`
    select chiave, length(coalesce(valore, '')) as lunghezza,
           (valore ~ '<[a-zA-Z]') as html,
           (valore ~ '^\\s*[\\[{]') as json,
           (select count(*) from regexp_matches(coalesce(valore, ''), '[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\\.[A-Za-z]{2,}', 'g')) as email_dentro,
           (valore ~ '^[0-9]+$') as numero,
           (valore ~* '^https?://') as url
    from public.impostazioni order by chiave`);
  console.log('=== impostazioni: ' + imp.length + ' chiavi ===');
  imp.forEach((r) => console.log('  ' + r.chiave.padEnd(28) + ' lunghezza ' + String(r.lunghezza).padStart(6)
    + (r.html ? ' | HTML' : '') + (r.json ? ' | JSON' : '') + (r.numero ? ' | numero' : '') + (r.url ? ' | URL' : '')
    + (Number(r.email_dentro) ? ' | ' + r.email_dentro + ' indirizzi email' : '')));

  const tip = await p.leggi('select nome, colore from public.tipologie order by nome');
  console.log('=== tipologie: ' + tip.map((t) => t.nome + ' (' + (t.colore || '-') + ')').join(', '));

  const link = await p.leggi(`
    select id, nome, substring(url from '^https?://([^/]+)') as host,
           length(url) as lunghezza_url,
           (url ~* '(password|pwd|token|user=|login=)') as sospetto
    from public.link_utili order by id`);
  console.log('=== link_utili: ' + link.length + ' ===');
  link.forEach((l) => console.log('  ' + String(l.id).padStart(3) + '  ' + String(l.nome).slice(0, 44).padEnd(44) + ' → ' + (l.host || '(non http)') + (l.sospetto ? '  ⚠ parametri sospetti' : '')));

  const letti = await p.leggi(`
    select letto, tipologia_letto, (nome <> '') as occupato,
           sesso, (data_nascita <> '') as ha_nascita, (data_ricovero <> '') as ha_ricovero,
           (codice_sanitario <> '') as ha_cs, ossigeno <> '' as ha_o2, vitto <> '' as ha_vitto, dimissibile
    from public.consegne order by nullif(regexp_replace(letto, '\\D', '', 'g'), '')::int nulls last, letto`);
  console.log('=== letti: ' + letti.length + ' ===');
  console.log('  ' + letti.map((l) => l.letto + ':' + l.tipologia_letto + (l.occupato ? '' : '(vuoto)')).join('  '));
  const conta = (f) => letti.filter(f).length;
  console.log('  occupati ' + conta((l) => l.occupato) + ' | con data nascita ' + conta((l) => l.ha_nascita) + ' | con data ricovero ' + conta((l) => l.ha_ricovero)
    + ' | con codice sanitario ' + conta((l) => l.ha_cs) + ' | con ossigeno ' + conta((l) => l.ha_o2) + ' | con vitto ' + conta((l) => l.ha_vitto));
  console.log('  valori di «sesso»: ' + [...new Set(letti.map((l) => JSON.stringify(l.sesso)))].join(' ') + ' | valori di «dimissibile»: ' + [...new Set(letti.map((l) => JSON.stringify(l.dimissibile)))].join(' '));

  const forme = await p.leggi(`
    select
      (select string_agg(distinct regexp_replace(eta, '[0-9]', '9', 'g'), ' | ') from public.consegne where eta <> '') as eta,
      (select string_agg(distinct regexp_replace(data_nascita, '[0-9]', '9', 'g'), ' | ') from public.consegne where data_nascita <> '') as data_nascita,
      (select string_agg(distinct regexp_replace(data_ricovero, '[0-9]', '9', 'g'), ' | ') from public.consegne where data_ricovero <> '') as data_ricovero,
      (select string_agg(distinct regexp_replace(codice_sanitario, '[0-9]', '9', 'g'), ' | ') from public.consegne where codice_sanitario <> '') as codice_sanitario,
      (select string_agg(distinct regexp_replace(ultimo_aggiornamento, '[0-9]', '9', 'g'), ' | ') from public.consegne where ultimo_aggiornamento <> '') as ultimo_aggiornamento,
      (select string_agg(distinct regexp_replace(regexp_replace(ossigeno, '[0-9]', '9', 'g'), '[A-Za-zÀ-ÿ]', 'x', 'g'), ' | ') from public.consegne where ossigeno <> '') as ossigeno,
      (select string_agg(distinct regexp_replace(regexp_replace(vitto, '[0-9]', '9', 'g'), '[A-Za-zÀ-ÿ]', 'x', 'g'), ' | ') from public.consegne where vitto <> '') as vitto`);
  console.log('=== formati dei campi brevi (cifre → 9, lettere → x) ===');
  Object.entries(forme[0]).forEach(([k, v]) => console.log('  ' + k.padEnd(22) + String(v).slice(0, 300)));
})().catch((e) => { console.log('ERRORE: ' + e.message); process.exit(1); });
