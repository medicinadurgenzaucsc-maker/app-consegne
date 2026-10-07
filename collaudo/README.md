# Ambiente di collaudo

Copia dell'app con **dati di pazienti inventati**, dove si provano modifiche e
correzioni prima di portarle al reparto. Stesso codice della produzione: cambia
solo il database a cui il sito si collega.

|                    | Produzione                                              | Collaudo                                     |
|--------------------|---------------------------------------------------------|----------------------------------------------|
| Sito               | `medicinadurgenzaucsc-maker.github.io/app-consegne/`    | `gistech2026.github.io/app-consegne/`        |
| Repository         | `medicinadurgenzaucsc-maker/app-consegne` (`origin`)    | `gistech2026/app-consegne` (`collaudo`)      |
| Progetto Supabase  | `ifmmcvxzhwdkmzhsxcvb`                                  | `rqvohwpthhumydpbwktq`                       |
| Mail dimissioni    | parte davvero (Gmail del reparto)                       | registrata in `posta_simulata`, mai spedita  |

Il sito sceglie il database dal nome host (`_AMBIENTI` in `docs/js/api.js`): è
collaudo solo su `gistech2026.github.io` e su `localhost`; qualunque altro
indirizzo è produzione. Una cornice arancione con «COLLAUDO — DATI FITTIZI»
segnala l'ambiente, anche in stampa. Ogni database dichiara chi è nella riga
`AMBIENTE` di `impostazioni`: se sito e database non concordano l'app si ferma.

## Come si lavora

1. Le modifiche si fanno sul ramo `collaudo`.
2. `git push collaudo collaudo:master` pubblica il sito di collaudo; poi
   `node collaudo/strumenti/pubblica-versione.js` aggiorna `app_version` nel
   database di collaudo (in produzione lo fa il workflow `notify-deploy`).
3. Si prova lì: a mano, con la batteria di prove dentro l'app (vedi «Banco di
   prova») e con le batterie in `collaudo/strumenti/`.
4. Con l'ok di chi gestisce l'app: `git checkout master && git merge --ff-only
   collaudo && git push origin master`. In produzione arriva lo **stesso
   commit** collaudato.

Le modifiche al database e alle Edge Function seguono la stessa strada: prima
sul progetto di collaudo, poi identiche in produzione.

## Cosa c'è nel database di collaudo

- **Struttura**: identica alla produzione (tabelle, vincoli, indici, funzioni,
  trigger, RLS, privilegi, realtime). `clona-struttura.js confronta` lo
  dimostra confrontando le due descrizioni; `baseline.sql` è lo script generato
  dalla produzione il 04/10/2026.
- **Configurazione**: tipologie e scale di valutazione copiate identiche; link
  utili con i documenti privati ridotti al solo sito; destinatari della mail
  sostituiti con indirizzi `example.com`.
- **Pazienti**: 28 letti come il reparto, 24 occupati da pazienti inventati,
  generati facendo girare le funzioni vere dell'app (`generatore-pazienti.js`).
  Nessuna riga di pazienti veri è mai stata copiata.
- **In più rispetto alla produzione**, per scelta: la funzione di RLS
  automatica (`rls_auto_enable`), la tabella `posta_simulata` e la funzione
  `google-finto`, che fa la parte di Google.

## Strumenti (`collaudo/strumenti/`)

I token si leggono dalla configurazione locale di Claude Code (voci MCP
`supabase` e `supabase-collaudo`) e dall'archivio credenziali di Git: nel
repository non c'è alcun segreto.

| Script | A cosa serve |
|---|---|
| `sb.js` | Accesso ai due database: `collaudo()` legge e scrive solo sul progetto di collaudo, `produzione()` è in **sola lettura**. |
| `clona-struttura.js ddl\|applica\|confronta` | Ricostruisce la struttura della produzione nel collaudo e ne verifica l'identità. |
| `copia-config.js` | Ricopia la configurazione non personale. |
| `confronta-forme.js` | Confronta la forma dei contenuti (tag e classi, mai il testo) fra i due ambienti. |
| `gettone-prova.js [ore] [cartella]` | Token di sessione per l'utente fittizio delle prove automatiche (vale solo nel collaudo). |
| `banco.js [porta]` | Banco di prova locale: serve l'app così com'è nella cartella di lavoro, collegata al collaudo e già «dentro» con l'utente fittizio. |
| `inventario-markup.js [produzione\|collaudo]` | Elenco di tag, attributi, classi e proprietà di stile presenti nei campi delle schede (solo nomi e conteggi, mai il testo): serve a tarare e a ricontrollare il filtro dell'HTML (`docs/js/sanifica.js`). |
| `pubblica-sito.js` | Pubblica il sito di servizio `gistech2026.github.io/collaudo/` (finto TrakCare e informativa). |
| `controlli-rilascio.js [riferimento]` | Controlli statici prima di ogni pubblicazione: sintassi, segnalibri collassati e loro versione, tag script al completo, impronta della libreria del filtro, service worker, uscita solo locale (`signOut` con `scope: 'local'`), versione nel menu uguale a `CACHE_NAME`. Confronta con `origin/master` (la produzione) se non si indica altro. |
| `confronta-filtro.js [produzione\|collaudo] [--backup N]` | Prima di pubblicare una versione che introduce o cambia il filtro dell'HTML: controlla che le schede (e gli ultimi N backup) non usino tag, attributi, classi o stili che il filtro toglierebbe. L'analisi gira dentro il database: escono solo nomi e conteggi, mai il testo. `--autoprova` verifica lo strumento stesso su un contenuto costruito apposta. |
| `prova-uscita.js` | Dimostra con sessioni vere che «Esci» su un dispositivo non scollega gli altri: utente provvisorio, tre sessioni, uscita `local` dalla prima, le altre due restano; controprova con `global`. L'utente viene eliminato alla fine. |
| `funzioni-collaudo.js pubblica\|configura` | Pubblica `google-token` e `google-finto` nel collaudo e ne imposta le variabili. |
| `prova-funzione-mail.js` | 25 prove della funzione mail contro il finto Google. |
| `supabase-accesso.js`, `github-accesso.js` | Autorizzazione «a codice» dei due account di collaudo, senza far passare chiavi dalla chat. |

## Provare in locale da utente collegato

Il login Google lo fa solo una persona. Per le prove automatiche si usa un
utente fittizio, `collaudo-automatico@example.com`, presente nella lista
autorizzati del solo collaudo: `gettone-prova.js` gli firma un token a scadenza
con la chiave del progetto di collaudo. Lo stesso token in produzione è
rifiutato.

## Banco di prova e prove automatiche

```
node collaudo/strumenti/banco.js 8765
```

| Indirizzo | Cosa mostra |
|---|---|
| `http://localhost:8765/` | l'app (cartella `docs/`), collegata al collaudo con l'utente fittizio |
| `http://localhost:8765/?sloggato` | la stessa, senza sessione: è la **prova di fumo** |
| `http://localhost:8765/?senzaFiltro` | la stessa, come se la libreria del filtro dell'HTML non si fosse caricata: deve comparire l'avviso a tutto schermo e nessuna scheda |
| `http://localhost:8765/trak-finto/` | il finto TrakCare |

Nel banco il service worker è spento e nessun file viene copiato o modificato.
Dentro la pagina dell'app si caricano, dalla console, gli script di
`collaudo/pagina/` (serviti sotto `/pagina/`):

- `generatore-pazienti.js` — `await window.__generaPazienti({ azzera: true })`
  riempie i 28 letti con 24 pazienti inventati usando le funzioni vere dell'app.
- `prove.js` — `await window.__prove()` esegue le prove e restituisce
  `{ prove, superate, fallite, elenco }`. Le sezioni si possono scegliere:
  `window.__prove({ sezioni: ['xss'] })`.

| Sezione | Cosa controlla |
|---|---|
| `ambiente` | la pagina lavora sul collaudo e solo lì; reparto finto al completo; nessun errore |
| `trak` | i due segnalibri veri, eseguiti sul finto TrakCare, leggono ciò che il catalogo descrive; reimportare non cambia nulla; giorni simulati |
| `xss` | ciò che arriva da fuori (indirizzo, righe del database, backup, memoria del browser) non diventa mai codice né markup attivo; i contenuti leciti restano identici; senza filtro l'app si ferma |

La sezione `xss` scrive sul letto libero «5», crea e cancella righe di prova
(un letto, una tipologia, due backup, due link) e alla fine rimette tutto
com'era. Per essere sicuri che le prove misurino davvero, le si rilancia dopo
aver neutralizzato a mano una protezione (`window._testoHtml = String` oppure
`window._pulisciHtml = String`): devono fallire.

**Prova di fumo, da fare a mano prima di ogni rilascio** (anche sul sito di
collaudo pubblicato, dove il service worker è attivo):

1. `…/?sloggato` (sul sito pubblicato: una finestra anonima): console pulita,
   bottone «Accedi con Google» visibile, nessun avviso a tutto schermo.
2. `…/?toast=info&msg=%3Cimg%20src%3Dx%20onerror%3D%22alert(document.domain)%22%3E`:
   non deve aprirsi alcuna finestra; a utente collegato il messaggio compare
   nel toast come testo, tale e quale.

## Finto TrakCare e sito di servizio

`collaudo/trak-finto/` riproduce le due pagine di TrakCare da cui i segnalibri
leggono (Pannello Terapia e Laboratorio) per i 24 pazienti inventati:
`catalogo.js` descrive terapie ed esami, con una data fissa e alcuni «giorni
dopo» per simulare ciò che cambia da un giorno all'altro. Sta fuori da `docs/`
perché ciò che è in `docs/` arriva anche in produzione; `pubblica-sito.js` lo
pubblica, con l'informativa sull'accesso, su `gistech2026.github.io/collaudo/`.

## Regole

- Mai copiare nel collaudo righe di pazienti veri, nemmeno anonimizzate: il
  testo libero identifica comunque.
- Dagli strumenti la produzione si legge soltanto.
- Il repository di collaudo non ha segreti: i due workflow sono attivi solo nel
  repository di produzione.
- Una modifica al database che crea tabelle scrive sempre la RLS in modo
  esplicito: nel collaudo verrebbe accesa comunque in automatico, in produzione no.
