# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Regole di collaborazione

- **Rispondere sempre in italiano** ed **elencare i file modificati alla fine di ogni risposta**.
- **Questa è produzione con pazienti veri** (reparto Medicina d'Urgenza 11N, Policlinico Gemelli). I test vanno fatti su letti liberi o d'appoggio, ripristinando poi lo stato originale; mai inventare o lasciare dati clinici finti. Da TrakCare si **legge soltanto** (mai toccare Ordini, checkbox «Interrompi», «Dimissione finale»); i login Google/TrakCare li fa l'utente.
- A fine lavoro, se sono cambiati i flussi dati, **simulare l'impatto sui limiti Supabase** (free tier 5 GB egress/mese; consumo storico stimato ~0,5–0,9 GB/mese). Vincolo di progetto: gli esami di laboratorio completi non viaggiano MAI nel full-sync/Realtime — solo le colonne di riepilogo.

## Cos'è

PWA «Sistema Consegne Reparto»: schede pazienti per letto, modificate in tempo reale da ~6 PC di reparto. Vanilla HTML/JS (nessun framework, nessuna build), Bootstrap 5 + SweetAlert2 + bootstrap-icons, servita da **GitHub Pages dalla cartella `docs/`**. Backend **Supabase** (progetto `ifmmcvxzhwdkmzhsxcvb`, raggiungibile via MCP `supabase`): PostgREST + Realtime + Edge Functions. I file nella **root** (`Codice.gs`, `Scripts.html`, `Index.html`…) sono il legacy Google Apps Script: non si toccano, l'app è solo `docs/`.

- URL pubblico: `https://medicinadurgenzaucsc-maker.github.io/app-consegne/` — **sottocartella**: la radice del dominio risponde 404 e sembra un deploy mancato.
- File principali: `docs/index.html` (~10k righe: login, focus mode, modal, bookmarklet, tutta la logica inline), `docs/js/api.js` (client Supabase, template card, ciclo dati, backup, laboratorio, problemi attivi), `docs/js/app.js` (avvio, sync, operazioni sui letti, token Google), `docs/js/app2.js`, `docs/print.html` (stampa autonoma: rilegge i dati da Supabase e ha CSS proprio — le regole di stile vanno replicate lì), `docs/sw.js` (service worker cache-first).

## Rilascio (nessuna build, nessun test runner)

1. Modifica i file in `docs/`.
2. **Bump obbligatorio** di `CACHE_NAME` in `docs/sw.js` (`consegne-vNNN` → vNNN+1) a ogni modifica servita: senza bump i PC restano sulla versione vecchia. Un «bug» segnalato subito dopo un deploy è spesso solo cache non aggiornata.
3. Verifica sintassi: `node --check docs/js/*.js` + compilazione degli script inline di `index.html` (estrarli con regex `<script>…</script>` e compilarli con `new Function`) + i due bookmarklet (`_bmTerapiaSorgente`, `_bmLabSorgente`) devono compilare **anche collassati su una riga**.
4. `git fetch` prima di pushare: il remoto può essere avanti (keepalive, altre sessioni) → `git pull --rebase`.
5. `git push origin master` → GitHub Pages pubblica in ~30–120 s. Verificare la propagazione con `curl` su `…/app-consegne/sw.js?cb=$RANDOM` finché compare la versione nuova.
6. Il workflow `.github/workflows/notify-deploy.yml` aggiorna `app_version` su Supabase a build concluso → i client aperti mostrano il badge «Update».
7. Database: migrazioni e query via MCP (`apply_migration`/`execute_sql`). Dal 30/10/2026 le tabelle nuove richiedono **GRANT espliciti**; abilitare sempre RLS subito (una tabella privata = RLS senza policy + `revoke` ad anon/authenticated, come `google_oauth`).

## Architettura dati

- `consegne`: una riga per letto; i campi scheda sono **HTML rich-text** (mappa campo↔colonna in `_CAMPI_MAP`, api.js). Svuotamento/spostamento letti: `_sbDimettiLetto`/`_sbSpostaPaziente` (api.js), che trascinano anche gli esami (`_labPulisci`/`_labScambia`).
- `lab_esami`: una riga per letto, `dati` jsonb con TUTTI gli esami (`{esami:{chiave:{g,n,range,um,v:{"gg/mm/aaaa hh:mm":"valore"}}}, ord:[…]}`) + colonne di riepilogo (`n_esami`, `ultimo_esame`, `allarme_giorni`, `allarme_visto`) che bastano a disegnare alambicco e promemoria sulle card con una query da ~1 KB (cache 5 min in `window._labInfoLetti` — **va aggiornata in modo sincrono** da chi cambia paziente/letto, non basta scrivere sul DB).
- `archivio`: backup orario automatico (CAS anti-doppione tra client) — `dati` = array pazienti completo, `lab` = snapshot esami (azzerato dopo 7 giorni per peso). L'elenco archivio seleziona solo `data_str`/`ts`: mai aggiungere select pesanti lì.
- `impostazioni` (chiave/valore): `account_login`, template mail, `ULTIMO_BACKUP`…
- `google_oauth` (privata) + Edge Function **`google-token`** (azioni `stato`/`config`/`scambia`/`token`): la «cassaforte di reparto» — refresh token Google sul server, autorizzazione concessa una volta per tutti i PC. Client: `_cassaforteToken()` è il tentativo 0 per Drive/Docs/Gmail; riserve: iframe `oauth-silent.html`, poi popup GIS.
- Sincronizzazione: full-sync al boot (`_sincronizzaEPoiFai`), Realtime per gli aggiornamenti, lock per-letto, poll di sicurezza 5 min, modalità emergenza auto-attivabile dalla diagnostica (che gira in background: l'avvio non si blocca).
- Backup su Google Drive: Google Doc orario con tabelle label→valore (`_driveRenderCard`), **round-trip** re-importabile (parser per label in app.js).

## Pattern critici (violarli = bug già vissuti)

- **Focus mode**: i campi si modificano solo lì. Il listener `focusout` (index.html) chiude il focus mode quando il fuoco esce dalla card, con guardie per: apertura in corso, Swal (`.swal2-container` + grace `_swalChiusoTs`), modal Bootstrap, modal Laboratorio (`labModalOverlay` + `_labModalChiusoTs`). **Qualunque nuova finestra sopra il focus mode** deve usare Swal (pattern del checkbox Dimissibile: click delegato in capture + `stopImmediatePropagation` + guard focus-mode) oppure un overlay che blocca la propagazione come quello del Laboratorio — mai un overlay nuovo senza guardia.
- **Regola del fuoco**: dopo un'azione programmatica riportare il fuoco su un nodo stabile della card (`returnFocus:false` sui Swal — il ritorno automatico può puntare a nodi distrutti); `_aggiornaCardDaPaziente` rifocalizza il campo che aveva il fuoco anche senza caret. Senza, il sync butta il fuoco su `body` e il focus mode si chiude da solo.
- **Stato derivato dentro i campi** (blocchi Laboratorio, scale, barretta/bottoni dei Problemi Attivi): viene salvato nell'HTML del campo ma è ricostruito/normalizzato a ogni giro. La ricostruzione vive **dentro `_aggiornaCardDaPaziente`** perché ogni riscrittura di campo passa di lì, compreso il fresh-fetch all'apertura del focus mode: agganciarla solo all'apertura non basta (il fetch arriva dopo e cancella).
- **Problemi Attivi** (campo `PianoTerapeutico`): elenco puntato gestito SOLO dai bottoni; scrittura manuale bloccata (`beforeinput`/`paste`/`drop`). Scrivere nel campo esclusivamente via `_paRiga`/`_paApplica`; nel Doc di Drive va il testo pulito (`_paDriveTesto`).
- **Bookmarklet TrakCare** (import terapia e laboratorio): le funzioni sorgente vengono collassate su una riga (`toString().replace(/\n\s*/g,' ')`) → **mai commenti `//` al loro interno**. Ogni cambio di comportamento richiede bump di `_BM_VERSIONE`/`_LAB_BM_VERSIONE` (payload compreso): l'import si blocca e obbliga a ritrascinare il segnalibro, altrimenti i colleghi usano il vecchio senza saperlo. Il match del nome paziente (`_nomiCombaciano`) non si bypassa mai.
- **Modifiche al sorgente**: `index.html` contiene NBSP letterali e i file sono CRLF → le sostituzioni esatte possono fallire. Il metodo collaudato: script node con ancore uniche (errore se assenti o doppie), normalizzazione CRLF→LF, patch, riconversione.
- **Egress**: non aggiungere select con colonne jsonb pesanti in percorsi frequenti (sync, render card); le feature nuove devono riusare il salvataggio standard (`attivaSalvataggioRitardato(letto, card)`) invece di scritture proprie.
- **Sicurezza (debito noto)**: anon key in chiaro + RLS `USING(true)` su quasi tutto («Path A» di chiusura ancora da fare). Non peggiorare la superficie; i segreti veri vanno in tabelle private lette solo dalle Edge Function (service role).

## Verifiche

Non esiste una suite: i controlli si scrivono ad hoc nello scratchpad e si eseguono prima del deploy — compilazione script inline + bookmarklet collassati, e un «banco di prova» HTML servito da un mini-server node locale che carica CSS reale + funzioni estratte dal sorgente per asserire comportamento e computed style (il Browser pane non esegue script nei file aperti via `file://`). Per i test funzionali sull'app vera serve l'accesso dell'utente (login Google) — chiederlo invece di aggirarlo.
