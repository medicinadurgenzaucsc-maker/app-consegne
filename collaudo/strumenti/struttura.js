// Descrizione normalizzata della struttura dello schema «public» di un database:
// tabelle, colonne, vincoli, indici, sequenze, funzioni, trigger, RLS, policy,
// privilegi, pubblicazione realtime. Serve a due scopi: generare il DDL del
// collaudo a partire dalla produzione e poi dimostrare che le due strutture
// coincidono (stessa descrizione = stesso database, dati a parte).
// Ogni blocco è UNA select che restituisce una riga con una colonna jsonb.

const RUOLI = "('anon','authenticated','service_role')";

const BLOCCHI = {
  tabelle: `
    select coalesce(jsonb_agg(jsonb_build_object(
      'tabella', c.relname,
      'proprietario', c.relowner::regrole::text,
      'rls', c.relrowsecurity,
      'rls_forzata', c.relforcerowsecurity,
      'replica_identity', c.relreplident::text,
      'opzioni', coalesce(array_to_string(c.reloptions, ','), ''),
      'commento', coalesce(obj_description(c.oid, 'pg_class'), '')
    ) order by c.relname), '[]'::jsonb) as j
    from pg_class c join pg_namespace n on n.oid = c.relnamespace
    where n.nspname = 'public' and c.relkind = 'r'`,

  colonne: `
    select coalesce(jsonb_agg(jsonb_build_object(
      'tabella', c.relname,
      'pos', a.attnum,
      'colonna', a.attname,
      'tipo', format_type(a.atttypid, a.atttypmod),
      'not_null', a.attnotnull,
      'default', coalesce(pg_get_expr(d.adbin, d.adrelid), ''),
      'identity', a.attidentity::text,
      'generata', a.attgenerated::text,
      'collazione', coalesce((select collname from pg_collation co where co.oid = a.attcollation and co.collname <> 'default'), ''),
      'commento', coalesce(col_description(c.oid, a.attnum), '')
    ) order by c.relname, a.attnum), '[]'::jsonb) as j
    from pg_attribute a
    join pg_class c on c.oid = a.attrelid
    join pg_namespace n on n.oid = c.relnamespace
    left join pg_attrdef d on d.adrelid = a.attrelid and d.adnum = a.attnum
    where n.nspname = 'public' and c.relkind = 'r' and a.attnum > 0 and not a.attisdropped`,

  vincoli: `
    select coalesce(jsonb_agg(jsonb_build_object(
      'tabella', c.relname,
      'vincolo', k.conname,
      'tipo', k.contype::text,
      'definizione', pg_get_constraintdef(k.oid)
    ) order by c.relname, k.conname), '[]'::jsonb) as j
    from pg_constraint k
    join pg_class c on c.oid = k.conrelid
    join pg_namespace n on n.oid = c.relnamespace
    where n.nspname = 'public' and c.relkind = 'r'`,

  indici: `
    select coalesce(jsonb_agg(jsonb_build_object(
      'tabella', t.relname,
      'indice', i.relname,
      'definizione', pg_get_indexdef(x.indexrelid),
      'da_vincolo', exists (select 1 from pg_constraint k where k.conindid = x.indexrelid)
    ) order by t.relname, i.relname), '[]'::jsonb) as j
    from pg_index x
    join pg_class i on i.oid = x.indexrelid
    join pg_class t on t.oid = x.indrelid
    join pg_namespace n on n.oid = t.relnamespace
    where n.nspname = 'public' and t.relkind = 'r'`,

  sequenze: `
    select coalesce(jsonb_agg(jsonb_build_object(
      'sequenza', s.sequencename,
      'tipo', s.data_type::text,
      'inizio', s.start_value::text, 'min', s.min_value::text, 'max', s.max_value::text,
      'passo', s.increment_by::text, 'ciclo', s.cycle, 'cache', s.cache_size::text,
      'di', coalesce((select dc.relname || '.' || da.attname
                      from pg_depend dp
                      join pg_class dc on dc.oid = dp.refobjid
                      join pg_attribute da on da.attrelid = dp.refobjid and da.attnum = dp.refobjsubid
                      where dp.objid = (quote_ident(s.schemaname) || '.' || quote_ident(s.sequencename))::regclass
                        and dp.classid = 'pg_class'::regclass and dp.deptype in ('a', 'i') limit 1), ''),
      'interna_identity', exists (select 1 from pg_depend dp
                      where dp.objid = (quote_ident(s.schemaname) || '.' || quote_ident(s.sequencename))::regclass
                        and dp.classid = 'pg_class'::regclass and dp.deptype = 'i')
    ) order by s.sequencename), '[]'::jsonb) as j
    from pg_sequences s where s.schemaname = 'public'`,

  funzioni: `
    select coalesce(jsonb_agg(jsonb_build_object(
      'funzione', p.proname,
      'argomenti', pg_get_function_identity_arguments(p.oid),
      'proprietario', p.proowner::regrole::text,
      'definizione', pg_get_functiondef(p.oid)
    ) order by p.proname, pg_get_function_identity_arguments(p.oid)), '[]'::jsonb) as j
    from pg_proc p join pg_namespace n on n.oid = p.pronamespace
    where n.nspname = 'public' and p.prokind in ('f', 'p')`,

  trigger: `
    select coalesce(jsonb_agg(jsonb_build_object(
      'tabella', c.relname,
      'trigger', t.tgname,
      'stato', t.tgenabled::text,
      'definizione', pg_get_triggerdef(t.oid)
    ) order by c.relname, t.tgname), '[]'::jsonb) as j
    from pg_trigger t
    join pg_class c on c.oid = t.tgrelid
    join pg_namespace n on n.oid = c.relnamespace
    where n.nspname = 'public' and not t.tgisinternal`,

  policy: `
    select coalesce(jsonb_agg(jsonb_build_object(
      'tabella', c.relname,
      'policy', p.polname,
      'comando', p.polcmd::text,
      'permissiva', p.polpermissive,
      'ruoli', (select coalesce(jsonb_agg(case when r = 0 then 'public' else r::regrole::text end order by 1), '[]'::jsonb) from unnest(p.polroles) r),
      'using', coalesce(pg_get_expr(p.polqual, p.polrelid), ''),
      'check', coalesce(pg_get_expr(p.polwithcheck, p.polrelid), '')
    ) order by c.relname, p.polname), '[]'::jsonb) as j
    from pg_policy p
    join pg_class c on c.oid = p.polrelid
    join pg_namespace n on n.oid = c.relnamespace
    where n.nspname = 'public'`,

  privilegi_tabelle: `
    select coalesce(jsonb_agg(jsonb_build_object(
      'tabella', x.relname, 'ruolo', x.ruolo, 'privilegi', x.privilegi
    ) order by x.relname, x.ruolo), '[]'::jsonb) as j
    from (
      select c.relname,
             case when a.grantee = 0 then 'PUBLIC' else a.grantee::regrole::text end as ruolo,
             string_agg(a.privilege_type || case when a.is_grantable then '*' else '' end, ',' order by a.privilege_type) as privilegi
      from pg_class c
      join pg_namespace n on n.oid = c.relnamespace
      cross join lateral aclexplode(coalesce(c.relacl, acldefault('r', c.relowner))) a
      where n.nspname = 'public' and c.relkind = 'r'
        and (a.grantee = 0 or a.grantee::regrole::text in ${RUOLI})
      group by c.relname, a.grantee
    ) x`,

  privilegi_colonne: `
    select coalesce(jsonb_agg(jsonb_build_object(
      'tabella', table_name, 'colonna', column_name, 'ruolo', grantee, 'privilegio', privilege_type
    ) order by table_name, column_name, grantee, privilege_type), '[]'::jsonb) as j
    from information_schema.column_privileges cp
    where table_schema = 'public' and grantee in ${RUOLI}
      and not exists (select 1 from information_schema.role_table_grants g
                      where g.table_schema = cp.table_schema and g.table_name = cp.table_name
                        and g.grantee = cp.grantee and g.privilege_type = cp.privilege_type)`,

  privilegi_sequenze: `
    select coalesce(jsonb_agg(jsonb_build_object(
      'sequenza', x.relname, 'ruolo', x.ruolo, 'privilegi', x.privilegi
    ) order by x.relname, x.ruolo), '[]'::jsonb) as j
    from (
      select c.relname,
             case when a.grantee = 0 then 'PUBLIC' else a.grantee::regrole::text end as ruolo,
             string_agg(a.privilege_type, ',' order by a.privilege_type) as privilegi
      from pg_class c
      join pg_namespace n on n.oid = c.relnamespace
      cross join lateral aclexplode(coalesce(c.relacl, acldefault('S', c.relowner))) a
      where n.nspname = 'public' and c.relkind = 'S'
        and (a.grantee = 0 or a.grantee::regrole::text in ${RUOLI})
      group by c.relname, a.grantee
    ) x`,

  privilegi_funzioni: `
    select coalesce(jsonb_agg(jsonb_build_object(
      'funzione', x.firma, 'ruolo', x.ruolo, 'privilegi', x.privilegi
    ) order by x.firma, x.ruolo), '[]'::jsonb) as j
    from (
      select p.proname || '(' || pg_get_function_identity_arguments(p.oid) || ')' as firma,
             case when a.grantee = 0 then 'PUBLIC' else a.grantee::regrole::text end as ruolo,
             string_agg(a.privilege_type, ',' order by a.privilege_type) as privilegi
      from pg_proc p
      join pg_namespace n on n.oid = p.pronamespace
      cross join lateral aclexplode(coalesce(p.proacl, acldefault('f', p.proowner))) a
      where n.nspname = 'public'
        and (a.grantee = 0 or a.grantee::regrole::text in ${RUOLI})
      group by p.proname, p.oid, a.grantee
    ) x`,

  pubblicazione: `
    select coalesce(jsonb_agg(jsonb_build_object(
      'pubblicazione', p.pubname,
      'tabella', c.relname,
      'colonne', coalesce((select string_agg(a.attname, ',' order by a.attnum) from pg_attribute a where a.attrelid = c.oid and a.attnum = any (r.prattrs::int2[])), ''),
      'filtro', coalesce(pg_get_expr(r.prqual, r.prrelid), '')
    ) order by p.pubname, c.relname), '[]'::jsonb) as j
    from pg_publication_rel r
    join pg_publication p on p.oid = r.prpubid
    join pg_class c on c.oid = r.prrelid
    join pg_namespace n on n.oid = c.relnamespace
    where n.nspname = 'public'`,

  altri_oggetti: `
    select jsonb_build_object(
      'viste', (select coalesce(jsonb_agg(c.relname order by c.relname), '[]'::jsonb) from pg_class c join pg_namespace n on n.oid = c.relnamespace where n.nspname = 'public' and c.relkind in ('v', 'm')),
      'tipi', (select coalesce(jsonb_agg(t.typname order by t.typname), '[]'::jsonb) from pg_type t join pg_namespace n on n.oid = t.typnamespace where n.nspname = 'public' and t.typtype in ('e', 'd', 'c') and not exists (select 1 from pg_class c where c.reltype = t.oid)),
      'tabelle_esterne_o_partizioni', (select coalesce(jsonb_agg(c.relname order by c.relname), '[]'::jsonb) from pg_class c join pg_namespace n on n.oid = c.relnamespace where n.nspname = 'public' and c.relkind in ('f', 'p'))
    ) as j`,
};

// esegui(sql) -> array di righe; restituisce { blocco: valore }
async function descrivi(esegui) {
  const out = {};
  for (const [nome, sql] of Object.entries(BLOCCHI)) {
    const righe = await esegui(sql.trim());
    out[nome] = righe[0].j;
  }
  return out;
}

module.exports = { descrivi, BLOCCHI };
