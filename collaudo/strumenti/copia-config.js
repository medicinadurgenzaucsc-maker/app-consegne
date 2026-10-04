// Copia nel COLLAUDO la sola configurazione non personale della produzione.
// I contenuti passano da un database all'altro dentro questo script, senza
// essere stampati, e come TESTO: il JSON non viene mai riletto da JavaScript,
// che altererebbe la scrittura dei numeri (1.0 diventerebbe 1).
// Ripetibile: le tabelle di configurazione del collaudo vengono svuotate e
// riempite di nuovo.
//
//  - tipologie, scale_valutazione: copiate identiche
//  - link_utili: nomi identici; i link a documenti privati (Google Docs/Drive,
//    Teams) sono ridotti al solo sito, senza il percorso del documento
//  - impostazioni: stesse chiavi; i destinatari della mail dimissioni diventano
//    indirizzi fittizi (example.com), i riferimenti Drive restano vuoti
//  - utenti_autorizzati: gli stessi account della produzione
const crypto = require('crypto');
const { collaudo, produzione } = require('./sb.js');

function dollaro(testo) {
  const t = '$cfg_' + crypto.randomBytes(6).toString('hex') + '$';
  if (testo.includes(t)) throw new Error('collisione di delimitatore');
  return t + testo + t;
}

(async () => {
  const p = produzione(), c = collaudo();
  const testo = async (sql) => String((await p.leggi(sql))[0].j);

  const tipologie = await testo("select coalesce(jsonb_agg(to_jsonb(t)), '[]'::jsonb)::text as j from public.tipologie t");
  const scale = await testo("select coalesce(jsonb_agg(to_jsonb(s) order by s.ordine, s.id), '[]'::jsonb)::text as j from public.scale_valutazione s");
  const link = await testo(`
    select coalesce(jsonb_agg(jsonb_build_object('id', l.id, 'nome', l.nome, 'url',
      case when substring(l.url from '^https?://([^/]+)') in ('docs.google.com', 'drive.google.com', 'teams.microsoft.com')
           then substring(l.url from '^(https?://[^/]+)') || '/'
           else l.url end) order by l.id), '[]'::jsonb)::text as j
    from public.link_utili l`);
  const impostazioni = await testo(`
    select coalesce(jsonb_agg(jsonb_build_object('chiave', i.chiave, 'valore',
      case
        when i.chiave = 'MAIL_DIMISSIONI_DESTINATARI' then (
          select string_agg(case when x.pezzo ~ '@' then 'destinatario' || x.n_mail || '@example.com' else x.pezzo end, '' order by x.ord)
          from (
            select m.ord, m.pezzo, count(*) filter (where m.pezzo ~ '@') over (order by m.ord) as n_mail
            from regexp_split_to_table(regexp_replace(coalesce(i.valore, ''), '([A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\\.[A-Za-z]{2,})', E'\\x01\\\\1\\x01', 'g'), E'\\x01') with ordinality as m(pezzo, ord)
          ) x)
        when i.chiave in ('DRIVE_FOLDER_BACKUP', 'DRIVE_FOLDER_ROOT') then ''
        when i.chiave = 'ULTIMO_BACKUP' then '0'
        else i.valore
      end) order by i.chiave), '[]'::jsonb)::text as j
    from public.impostazioni i`);
  const utenti = await testo("select coalesce(jsonb_agg(to_jsonb(u) order by u.email), '[]'::jsonb)::text as j from public.utenti_autorizzati u");
  const interruttore = (await p.leggi("select richiedi_utente from public.google_oauth where id = 'reparto'"))[0];

  const sql = `
    truncate public.tipologie, public.scale_valutazione, public.link_utili, public.impostazioni,
             public.utenti_autorizzati, public.app_version, public.keepalive, public.google_oauth;
    insert into public.tipologie select * from jsonb_populate_recordset(null::public.tipologie, ${dollaro(tipologie)}::jsonb);
    insert into public.scale_valutazione select * from jsonb_populate_recordset(null::public.scale_valutazione, ${dollaro(scale)}::jsonb);
    insert into public.link_utili select * from jsonb_populate_recordset(null::public.link_utili, ${dollaro(link)}::jsonb);
    select setval('public.link_utili_id_seq', (select coalesce(max(id), 1) from public.link_utili));
    insert into public.impostazioni select * from jsonb_populate_recordset(null::public.impostazioni, ${dollaro(impostazioni)}::jsonb);
    insert into public.utenti_autorizzati select * from jsonb_populate_recordset(null::public.utenti_autorizzati, ${dollaro(utenti)}::jsonb);
    insert into public.app_version (id, sha, deployed_at, message) values (1, '', 0, 'collaudo');
    insert into public.keepalive (id) values (1);
    insert into public.google_oauth (id, richiedi_utente) values ('reparto', ${interruttore && interruttore.richiedi_utente ? 'true' : 'false'});
    select jsonb_build_object(
      'tipologie', (select count(*) from public.tipologie),
      'scale', (select count(*) from public.scale_valutazione),
      'scale_byte', (select sum(octet_length(definizione::text)) from public.scale_valutazione),
      'link', (select count(*) from public.link_utili),
      'impostazioni', (select count(*) from public.impostazioni),
      'utenti_autorizzati', (select count(*) from public.utenti_autorizzati)
    ) as esito;`;
  const r = await c.query(sql);
  console.log('collaudo dopo la copia: ' + JSON.stringify(r[0].esito));

  // Controprova: le scale del collaudo sono identiche a quelle della produzione?
  const firma = "select md5(string_agg(md5(to_jsonb(s)::text), '' order by s.id)) as f, count(*)::int as n from public.scale_valutazione s";
  const fp = (await p.leggi(firma))[0], fc = (await c.query(firma))[0];
  console.log('scale — produzione ' + fp.n + ' (' + fp.f.slice(0, 10) + '…) | collaudo ' + fc.n + ' (' + fc.f.slice(0, 10) + '…) → ' + (fp.f === fc.f ? 'IDENTICHE' : 'DIVERSE'));
  const ft = "select md5(string_agg(nome || '|' || coalesce(colore, ''), ',' order by nome)) as f from public.tipologie";
  console.log('tipologie → ' + ((await p.leggi(ft))[0].f === (await c.query(ft))[0].f ? 'IDENTICHE' : 'DIVERSE'));

  // Riepilogo del collaudo, senza contenuti sensibili.
  const imp = await c.query(`select chiave, length(coalesce(valore, '')) as n,
      case when chiave in ('MAIL_DIMISSIONI_DESTINATARI', 'ULTIMO_BACKUP', 'GIORNI_ARCHIVIO') then valore else null end as mostra
    from public.impostazioni order by chiave`);
  imp.forEach((x) => console.log('  ' + x.chiave.padEnd(28) + ' lunghezza ' + String(x.n).padStart(4) + (x.mostra !== null ? ' → ' + x.mostra : '')));
  const lk = await c.query("select count(*) filter (where url ~ '^https?://[^/]+/$') as ridotti, count(*) as tot from public.link_utili");
  console.log('link utili: ' + lk[0].tot + ' (di cui ' + lk[0].ridotti + ' ridotti al solo sito)');
})().catch((e) => { console.log('ERRORE: ' + e.message); process.exit(1); });
