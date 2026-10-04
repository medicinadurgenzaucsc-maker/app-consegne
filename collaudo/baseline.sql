-- Baseline del collaudo: struttura dello schema public identica alla produzione.
-- Generata da clona-struttura.js: non modificare a mano.

-- 1. Sequenze
create sequence public."archivio_id_seq" as integer increment by 1 minvalue 1 maxvalue 2147483647 start with 1 cache 1 no cycle;
create sequence public."link_utili_id_seq" as integer increment by 1 minvalue 1 maxvalue 2147483647 start with 1 cache 1 no cycle;
create sequence public."logs_id_seq" as bigint increment by 1 minvalue 1 maxvalue 9223372036854775807 start with 1 cache 1 no cycle;
create sequence public."scale_valutazione_versioni_id_seq" as bigint increment by 1 minvalue 1 maxvalue 9223372036854775807 start with 1 cache 1 no cycle;

-- 2. Tabelle
create table public."app_version" (
  "id" integer default 1 not null,
  "sha" text default ''::text not null,
  "deployed_at" bigint default 0 not null,
  "message" text default ''::text
);
create table public."archivio" (
  "id" integer default nextval('archivio_id_seq'::regclass) not null,
  "data_str" text not null,
  "ts" bigint not null,
  "dati" jsonb not null,
  "created_at" timestamp with time zone default now(),
  "lab" jsonb
);
create table public."consegne" (
  "letto" text not null,
  "nome" text default ''::text,
  "eta" text default ''::text,
  "data_nascita" text default ''::text,
  "data_ricovero" text default ''::text,
  "diagnosi" text default ''::text,
  "note_terapia" text default ''::text,
  "diaria" text default ''::text,
  "da_fare" text default ''::text,
  "tipologia_letto" text default 'STANDARD'::text,
  "piano_terapeutico" text default ''::text,
  "allergie" text default ''::text,
  "codice_sanitario" text default ''::text,
  "ossigeno" text default ''::text,
  "sesso" text default ''::text,
  "ultimo_aggiornamento" text default ''::text,
  "updated_at" timestamp with time zone default now(),
  "esami_colturali" text default ''::text not null,
  "dimissibile" text default ''::text not null,
  "vitto" text default ''::text not null
);
create table public."google_oauth" (
  "id" text default 'reparto'::text not null,
  "client_secret" text,
  "refresh_token" text,
  "email" text,
  "scopes" text,
  "access_token" text,
  "access_scad" timestamp with time zone,
  "updated_at" timestamp with time zone default now(),
  "richiedi_utente" boolean default false not null
);
create table public."impostazioni" (
  "chiave" text not null,
  "valore" text default ''::text
);
create table public."keepalive" (
  "id" integer not null,
  "updated_at" timestamp with time zone default now() not null
);
create table public."lab_esami" (
  "letto" text not null,
  "paziente" text default ''::text not null,
  "dati" jsonb default '{}'::jsonb not null,
  "updated_at" timestamp with time zone default now() not null,
  "allarme_giorni" integer,
  "ultimo_esame" date,
  "n_esami" integer default 0 not null,
  "allarme_visto" date
);
create table public."link_utili" (
  "id" integer default nextval('link_utili_id_seq'::regclass) not null,
  "nome" text not null,
  "url" text not null
);
create table public."locks" (
  "letto" text not null,
  "token" text not null,
  "ts" bigint not null
);
create table public."logs" (
  "id" bigint default nextval('logs_id_seq'::regclass) not null,
  "ts" bigint not null,
  "livello" text default 'info'::text not null,
  "tipo" text not null,
  "messaggio" text not null,
  "descrizione" text,
  "device_id" text,
  "device_name" text,
  "user_agent" text,
  "url" text,
  "created_at" timestamp with time zone default now()
);
create table public."scale_valutazione" (
  "id" text not null,
  "nome" text not null,
  "descrizione" text,
  "categoria" text,
  "ordine" integer default 100,
  "attiva" boolean default true,
  "versione" integer default 1,
  "fonte_url" text,
  "riferimento" text,
  "data_revisione" date default now(),
  "definizione" jsonb not null,
  "updated_at" timestamp with time zone default now()
);
create table public."scale_valutazione_versioni" (
  "id" bigint default nextval('scale_valutazione_versioni_id_seq'::regclass) not null,
  "scala_id" text not null,
  "versione" integer not null,
  "nome" text,
  "fonte_url" text,
  "riferimento" text,
  "definizione" jsonb not null,
  "sostituita_il" timestamp with time zone default now() not null,
  "motivo" text,
  "casi_esito" text
);
create table public."tipologie" (
  "nome" text not null,
  "colore" text default ''::text
);
create table public."utenti_autorizzati" (
  "email" text not null,
  "note" text,
  "creato_il" timestamp with time zone default now() not null
);

-- 3. Proprietà delle sequenze
alter sequence public."archivio_id_seq" owned by public."archivio"."id";
alter sequence public."link_utili_id_seq" owned by public."link_utili"."id";
alter sequence public."logs_id_seq" owned by public."logs"."id";
alter sequence public."scale_valutazione_versioni_id_seq" owned by public."scale_valutazione_versioni"."id";

-- 4. Vincoli (chiavi primarie, univoci e check prima delle chiavi esterne)
alter table public."app_version" add constraint "app_version_pkey" PRIMARY KEY (id);
alter table public."archivio" add constraint "archivio_pkey" PRIMARY KEY (id);
alter table public."consegne" add constraint "consegne_pkey" PRIMARY KEY (letto);
alter table public."google_oauth" add constraint "google_oauth_pkey" PRIMARY KEY (id);
alter table public."impostazioni" add constraint "impostazioni_pkey" PRIMARY KEY (chiave);
alter table public."keepalive" add constraint "keepalive_pkey" PRIMARY KEY (id);
alter table public."lab_esami" add constraint "lab_esami_pkey" PRIMARY KEY (letto);
alter table public."link_utili" add constraint "link_utili_pkey" PRIMARY KEY (id);
alter table public."locks" add constraint "locks_pkey" PRIMARY KEY (letto);
alter table public."logs" add constraint "logs_pkey" PRIMARY KEY (id);
alter table public."scale_valutazione" add constraint "scale_valutazione_pkey" PRIMARY KEY (id);
alter table public."scale_valutazione_versioni" add constraint "scale_valutazione_versioni_pkey" PRIMARY KEY (id);
alter table public."tipologie" add constraint "tipologie_pkey" PRIMARY KEY (nome);
alter table public."utenti_autorizzati" add constraint "utenti_autorizzati_pkey" PRIMARY KEY (email);

-- 5. Indici non legati a un vincolo
CREATE INDEX idx_logs_livello ON public.logs USING btree (livello);
CREATE INDEX idx_logs_ts ON public.logs USING btree (ts DESC);
CREATE INDEX idx_svv_scala ON public.scale_valutazione_versioni USING btree (scala_id, versione DESC);

-- 6. Funzioni
CREATE OR REPLACE FUNCTION public.consegne_touch_updated_at()
 RETURNS trigger
 LANGUAGE plpgsql
 SET search_path TO 'public'
AS $function$
BEGIN
  NEW.updated_at := now();
  RETURN NEW;
END;
$function$;
CREATE OR REPLACE FUNCTION public.is_autorizzato()
 RETURNS boolean
 LANGUAGE sql
 STABLE SECURITY DEFINER
 SET search_path TO 'public'
AS $function$
  SELECT COALESCE(
    (auth.jwt() ->> 'email') IN (SELECT email FROM public.utenti_autorizzati),
    false
  );
$function$;
CREATE OR REPLACE FUNCTION public.ping()
 RETURNS jsonb
 LANGUAGE plpgsql
 SECURITY DEFINER
 SET search_path TO 'public'
AS $function$
begin
  update public.keepalive set updated_at = now() where id = 1;
  return jsonb_build_object('pong', now());
end $function$;
CREATE OR REPLACE FUNCTION public.scale_touch_updated_at()
 RETURNS trigger
 LANGUAGE plpgsql
 SET search_path TO 'public'
AS $function$
BEGIN NEW.updated_at := now(); RETURN NEW; END; $function$;

-- 7. Trigger
CREATE TRIGGER trg_consegne_touch_updated_at BEFORE UPDATE ON public.consegne FOR EACH ROW EXECUTE FUNCTION consegne_touch_updated_at();
CREATE TRIGGER trg_scale_touch BEFORE UPDATE ON public.scale_valutazione FOR EACH ROW EXECUTE FUNCTION scale_touch_updated_at();

-- 8. Row Level Security
alter table public."app_version" enable row level security;
alter table public."archivio" enable row level security;
alter table public."consegne" enable row level security;
alter table public."google_oauth" enable row level security;
alter table public."impostazioni" enable row level security;
alter table public."keepalive" enable row level security;
alter table public."lab_esami" enable row level security;
alter table public."link_utili" enable row level security;
alter table public."locks" enable row level security;
alter table public."logs" enable row level security;
alter table public."scale_valutazione" enable row level security;
alter table public."scale_valutazione_versioni" enable row level security;
alter table public."tipologie" enable row level security;
alter table public."utenti_autorizzati" enable row level security;
create policy "auth_read" on public."app_version" as permissive for select to "authenticated" using (is_autorizzato());
create policy "auth_rw" on public."archivio" as permissive for all to "authenticated" using (is_autorizzato()) with check (is_autorizzato());
create policy "auth_rw" on public."consegne" as permissive for all to "authenticated" using (is_autorizzato()) with check (is_autorizzato());
create policy "auth_rw" on public."impostazioni" as permissive for all to "authenticated" using (is_autorizzato()) with check (is_autorizzato());
create policy "auth_rw" on public."lab_esami" as permissive for all to "authenticated" using (is_autorizzato()) with check (is_autorizzato());
create policy "auth_rw" on public."link_utili" as permissive for all to "authenticated" using (is_autorizzato()) with check (is_autorizzato());
create policy "auth_rw" on public."locks" as permissive for all to "authenticated" using (is_autorizzato()) with check (is_autorizzato());
create policy "auth_rw" on public."logs" as permissive for all to "authenticated" using (is_autorizzato()) with check (is_autorizzato());
create policy "auth_read" on public."scale_valutazione" as permissive for select to "authenticated" using (is_autorizzato());
create policy "auth_read" on public."scale_valutazione_versioni" as permissive for select to "authenticated" using (is_autorizzato());
create policy "auth_rw" on public."tipologie" as permissive for all to "authenticated" using (is_autorizzato()) with check (is_autorizzato());

-- 9. Privilegi: si azzera ciò che il progetto concede in automatico e si
--    rimette esattamente ciò che ha la produzione.
revoke all on table public."app_version" from PUBLIC, "anon", "authenticated", "service_role";
revoke all on table public."archivio" from PUBLIC, "anon", "authenticated", "service_role";
revoke all on table public."consegne" from PUBLIC, "anon", "authenticated", "service_role";
revoke all on table public."google_oauth" from PUBLIC, "anon", "authenticated", "service_role";
revoke all on table public."impostazioni" from PUBLIC, "anon", "authenticated", "service_role";
revoke all on table public."keepalive" from PUBLIC, "anon", "authenticated", "service_role";
revoke all on table public."lab_esami" from PUBLIC, "anon", "authenticated", "service_role";
revoke all on table public."link_utili" from PUBLIC, "anon", "authenticated", "service_role";
revoke all on table public."locks" from PUBLIC, "anon", "authenticated", "service_role";
revoke all on table public."logs" from PUBLIC, "anon", "authenticated", "service_role";
revoke all on table public."scale_valutazione" from PUBLIC, "anon", "authenticated", "service_role";
revoke all on table public."scale_valutazione_versioni" from PUBLIC, "anon", "authenticated", "service_role";
revoke all on table public."tipologie" from PUBLIC, "anon", "authenticated", "service_role";
revoke all on table public."utenti_autorizzati" from PUBLIC, "anon", "authenticated", "service_role";
grant MAINTAIN, SELECT on table public."app_version" to "authenticated";
grant DELETE, INSERT, MAINTAIN, REFERENCES, SELECT, TRIGGER, TRUNCATE, UPDATE on table public."app_version" to "service_role";
grant DELETE, INSERT, MAINTAIN, SELECT, UPDATE on table public."archivio" to "authenticated";
grant DELETE, INSERT, MAINTAIN, REFERENCES, SELECT, TRIGGER, TRUNCATE, UPDATE on table public."archivio" to "service_role";
grant DELETE, INSERT, MAINTAIN, SELECT, UPDATE on table public."consegne" to "authenticated";
grant DELETE, INSERT, MAINTAIN, REFERENCES, SELECT, TRIGGER, TRUNCATE, UPDATE on table public."consegne" to "service_role";
grant DELETE, INSERT, MAINTAIN, REFERENCES, SELECT, TRIGGER, TRUNCATE, UPDATE on table public."google_oauth" to "service_role";
grant SELECT on table public."impostazioni" to "anon";
grant DELETE, INSERT, MAINTAIN, SELECT, UPDATE on table public."impostazioni" to "authenticated";
grant DELETE, INSERT, MAINTAIN, REFERENCES, SELECT, TRIGGER, TRUNCATE, UPDATE on table public."impostazioni" to "service_role";
grant DELETE, INSERT, MAINTAIN, REFERENCES, SELECT, TRIGGER, TRUNCATE, UPDATE on table public."keepalive" to "service_role";
grant DELETE, INSERT, MAINTAIN, SELECT, UPDATE on table public."lab_esami" to "authenticated";
grant DELETE, INSERT, MAINTAIN, REFERENCES, SELECT, TRIGGER, TRUNCATE, UPDATE on table public."lab_esami" to "service_role";
grant DELETE, INSERT, MAINTAIN, SELECT, UPDATE on table public."link_utili" to "authenticated";
grant DELETE, INSERT, MAINTAIN, REFERENCES, SELECT, TRIGGER, TRUNCATE, UPDATE on table public."link_utili" to "service_role";
grant DELETE, INSERT, MAINTAIN, SELECT, UPDATE on table public."locks" to "authenticated";
grant DELETE, INSERT, MAINTAIN, REFERENCES, SELECT, TRIGGER, TRUNCATE, UPDATE on table public."locks" to "service_role";
grant DELETE, INSERT, MAINTAIN, SELECT, UPDATE on table public."logs" to "authenticated";
grant DELETE, INSERT, MAINTAIN, REFERENCES, SELECT, TRIGGER, TRUNCATE, UPDATE on table public."logs" to "service_role";
grant MAINTAIN, SELECT on table public."scale_valutazione" to "authenticated";
grant DELETE, INSERT, MAINTAIN, REFERENCES, SELECT, TRIGGER, TRUNCATE, UPDATE on table public."scale_valutazione" to "service_role";
grant MAINTAIN, SELECT on table public."scale_valutazione_versioni" to "authenticated";
grant DELETE, INSERT, MAINTAIN, REFERENCES, SELECT, TRIGGER, TRUNCATE, UPDATE on table public."scale_valutazione_versioni" to "service_role";
grant DELETE, INSERT, MAINTAIN, SELECT, UPDATE on table public."tipologie" to "authenticated";
grant DELETE, INSERT, MAINTAIN, REFERENCES, SELECT, TRIGGER, TRUNCATE, UPDATE on table public."tipologie" to "service_role";
grant DELETE, INSERT, MAINTAIN, REFERENCES, SELECT, TRIGGER, TRUNCATE, UPDATE on table public."utenti_autorizzati" to "service_role";
revoke all on sequence public."archivio_id_seq" from PUBLIC, "anon", "authenticated", "service_role";
revoke all on sequence public."link_utili_id_seq" from PUBLIC, "anon", "authenticated", "service_role";
revoke all on sequence public."logs_id_seq" from PUBLIC, "anon", "authenticated", "service_role";
revoke all on sequence public."scale_valutazione_versioni_id_seq" from PUBLIC, "anon", "authenticated", "service_role";
grant SELECT, UPDATE, USAGE on sequence public."archivio_id_seq" to "authenticated";
grant SELECT, UPDATE, USAGE on sequence public."archivio_id_seq" to "service_role";
grant SELECT, UPDATE, USAGE on sequence public."link_utili_id_seq" to "authenticated";
grant SELECT, UPDATE, USAGE on sequence public."link_utili_id_seq" to "service_role";
grant SELECT, UPDATE, USAGE on sequence public."logs_id_seq" to "authenticated";
grant SELECT, UPDATE, USAGE on sequence public."logs_id_seq" to "service_role";
grant SELECT, UPDATE, USAGE on sequence public."scale_valutazione_versioni_id_seq" to "authenticated";
grant SELECT, UPDATE, USAGE on sequence public."scale_valutazione_versioni_id_seq" to "service_role";
revoke all on function public."consegne_touch_updated_at"() from PUBLIC, "anon", "authenticated", "service_role";
revoke all on function public."is_autorizzato"() from PUBLIC, "anon", "authenticated", "service_role";
revoke all on function public."ping"() from PUBLIC, "anon", "authenticated", "service_role";
revoke all on function public."scale_touch_updated_at"() from PUBLIC, "anon", "authenticated", "service_role";
grant EXECUTE on function public."consegne_touch_updated_at"() to "anon";
grant EXECUTE on function public."consegne_touch_updated_at"() to "authenticated";
grant EXECUTE on function public."consegne_touch_updated_at"() to PUBLIC;
grant EXECUTE on function public."consegne_touch_updated_at"() to "service_role";
grant EXECUTE on function public."is_autorizzato"() to "anon";
grant EXECUTE on function public."is_autorizzato"() to "authenticated";
grant EXECUTE on function public."is_autorizzato"() to PUBLIC;
grant EXECUTE on function public."is_autorizzato"() to "service_role";
grant EXECUTE on function public."ping"() to "anon";
grant EXECUTE on function public."ping"() to "authenticated";
grant EXECUTE on function public."ping"() to "service_role";
grant EXECUTE on function public."scale_touch_updated_at"() to "anon";
grant EXECUTE on function public."scale_touch_updated_at"() to "authenticated";
grant EXECUTE on function public."scale_touch_updated_at"() to PUBLIC;
grant EXECUTE on function public."scale_touch_updated_at"() to "service_role";

-- 10. Realtime
alter table public."app_version" replica identity full;
alter publication "supabase_realtime" add table public."app_version";
alter publication "supabase_realtime" add table public."consegne";
alter publication "supabase_realtime" add table public."locks";

-- 11. Commenti
comment on table public."scale_valutazione_versioni" is 'Snapshot delle versioni precedenti di scale_valutazione (rollback auto-update).';
comment on column public."archivio"."lab" is 'Snapshot di lab_esami al momento del backup (array di righe). Azzerato sui backup piu vecchi di 7 giorni per non gonfiare l archivio: gli esami restano comunque su TrakCare.';
comment on column public."lab_esami"."allarme_giorni" is 'Ogni quanti giorni vanno richiesti gli esami. NULL = usa il default della tipologia letto (SUB-INTENSIVA 1, altre 3).';
comment on column public."lab_esami"."ultimo_esame" is 'Data del prelievo piu recente, calcolata all''import: serve al promemoria senza scaricare i dati.';
comment on column public."lab_esami"."n_esami" is 'Numero di esami nel documento: se 0 la scheda non mostra il bottone Laboratorio.';
comment on column public."lab_esami"."allarme_visto" is 'Data in cui un collega ha spuntato il promemoria: per quel giorno non viene piu'' mostrato.';
