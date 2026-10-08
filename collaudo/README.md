# Ambiente di collaudo

Copia dell'app con **dati di pazienti inventati**, dove si provano modifiche e
correzioni prima di portarle al reparto. Stesso codice della produzione: cambia
solo il database a cui il sito si collega.

|                    | Produzione                                              | Collaudo                                     |
|--------------------|---------------------------------------------------------|----------------------------------------------|
| Sito               | `medicinadurgenzaucsc-maker.github.io/app-consegne/`    | `consegne-collaudo.pages.dev`, dietro l'accesso con Google |
| Repository         | `medicinadurgenzaucsc-maker/app-consegne` (`origin`)    | `gistech2026/app-consegne-sorgente`, privato (`collaudo`) |
| Progetto Supabase  | `ifmmcvxzhwdkmzhsxcvb`                                  | `rqvohwpthhumydpbwktq`                       |
| Mail dimissioni    | parte davvero (dal mittente scelto in «Impostazioni email») | parte davvero solo fra indirizzi di prova; simulata nelle prove automatiche |

Il sito sceglie il database dal nome host (`_AMBIENTI` in `docs/js/api.js`):
ogni ambiente elenca i propri indirizzi. Sono di collaudo `gistech2026.github.io`,
`localhost` e il sito su Cloudflare `consegne-collaudo.pages.dev` con le sue
anteprime. Un indirizzo che non sta in nessun elenco non è un ambiente: lì l'app
non parte e mostra «Indirizzo non riconosciuto», quindi un sito nuovo va
aggiunto all'elenco giusto prima di essere usato. Una cornice arancione con «COLLAUDO — DATI FITTIZI»
segnala l'ambiente, anche in stampa. Ogni database dichiara chi è nella riga
`AMBIENTE` di `impostazioni`: se sito e database non concordano l'app si ferma.

Dall'08/10/2026 il collaudo vive su Cloudflare: il vecchio indirizzo
`gistech2026.github.io/app-consegne/` mostra solo la pagina che rimanda a quello
nuovo, e il sorgente sta in un repository privato (vedi «Pubblicazione su
Cloudflare»).

## Come si lavora

1. Le modifiche si fanno sul ramo `collaudo`.
2. `git push collaudo collaudo:master` manda il commit al repository privato
   di collaudo: se tocca il sito, il flusso «Pubblica su Cloudflare» lo
   pubblica e avvisa le pagine aperte (aggiorna `app_version`).
   `node collaudo/strumenti/pubblicazione.js` segue il flusso fino alla fine e
   controlla che l'avviso sia arrivato. In produzione, finché il sito sta su
   GitHub Pages, pubblica GitHub e avvisa il workflow `notify-deploy`.
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
  utili con i documenti privati ridotti al solo sito. Mittente e destinatari
  della mail sono **indirizzi di prova di chi gestisce l'app** (impostazioni
  `MAIL_DIMISSIONI_MITTENTE` e `MAIL_DIMISSIONI_DESTINATARI`): `copia-config.js`
  non li tocca, così gli indirizzi del reparto non arrivano mai nel collaudo.
- **Pazienti**: 28 letti come il reparto, 24 occupati da pazienti inventati,
  generati facendo girare le funzioni vere dell'app (`generatore-pazienti.js`).
  Nessuna riga di pazienti veri è mai stata copiata.
- **In più rispetto alla produzione**, per scelta: la funzione di RLS
  automatica (`rls_auto_enable`), la tabella `posta_simulata`, la funzione
  `google-finto` che sta fra `google-token` e Google, le variabili `GOOGLE_*`
  che la fanno chiamare, e nella cassaforte `google_oauth` le righe `prova` e
  `reparto_vero` (vedi «La mail nel collaudo»).

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
| `banco-sw.js [porta]` | Banco del service worker: il sito servito come farebbero Cloudflare Pages o GitHub Pages, col service worker acceso e i guasti simulati (vedi «Service worker: banco e prove»). |
| `prepara-sito.js [cartella]` | Elenca (o copia in una cartella) i soli file del sito, quelli che vengono pubblicati su Cloudflare. |
| `prepara-rinvio.js <nuovo indirizzo> [cartella]` | Genera i file della pagina di rinvio (ciò che resta al vecchio indirizzo quando il sito si sposta), dai modelli in `rinvio/`. |
| `scambio-repository.js stato\|prepara\|scambia\|annulla` | Prova generale dello scambio dei repository nel collaudo: sorgente in un repository privato, al vecchio nome solo la pagina di rinvio (vedi «Pubblicazione su Cloudflare»). |
| `pubblicazione.js [avvia]` | Segue il flusso «Pubblica su Cloudflare» del repository di collaudo fino alla fine, ne mostra i passi e controlla che l'avviso ai PC sia arrivato in `app_version`; con `avvia` lo fa ripartire a mano sullo stesso commit. Non scrive nel database. |
| `verifica-chiusura.js collaudo\|produzione [indirizzo] [--archivi]` | Prova **da estraneo**, senza credenziali, che il sorgente non si scarichi più da nessuna strada, che al vecchio indirizzo resti solo la pagina di rinvio e che il sito nuovo chieda l'accesso: 28 prove, più 4 sugli archivi pubblici di terzi. Dove lo scambio non è stato fatto deve fallire: è la sua controprova. |
| `banco.js [porta]` | Banco di prova locale: serve l'app così com'è nella cartella di lavoro, collegata al collaudo e già «dentro» con l'utente fittizio. |
| `inventario-markup.js [produzione\|collaudo]` | Elenco di tag, attributi, classi e proprietà di stile presenti nei campi delle schede (solo nomi e conteggi, mai il testo): serve a tarare e a ricontrollare il filtro dell'HTML (`docs/js/sanifica.js`). |
| `pubblica-sito.js` | Pubblica il sito di servizio `gistech2026.github.io/collaudo/` (finto TrakCare e informativa). |
| `controlli-rilascio.js [riferimento]` | Controlli statici prima di ogni pubblicazione: sintassi, segnalibri collassati e loro versione, tag script al completo, impronta della libreria del filtro, service worker, uscita solo locale (`signOut` con `scope: 'local'`), versione nel menu uguale a `CACHE_NAME`. Confronta con `origin/master` (la produzione) se non si indica altro. |
| `confronta-filtro.js [produzione\|collaudo] [--backup N]` | Prima di pubblicare una versione che introduce o cambia il filtro dell'HTML: controlla che le schede (e gli ultimi N backup) non usino tag, attributi, classi o stili che il filtro toglierebbe. L'analisi gira dentro il database: escono solo nomi e conteggi, mai il testo. `--autoprova` verifica lo strumento stesso su un contenuto costruito apposta. |
| `prova-uscita.js` | Dimostra con sessioni vere che «Esci» su un dispositivo non scollega gli altri: utente provvisorio, tre sessioni, uscita `local` dalla prima, le altre due restano; controprova con `global`. L'utente viene eliminato alla fine. |
| `funzioni-collaudo.js pubblica\|configura` | Pubblica `google-token` e `google-finto` nel collaudo (dopo averne controllato la sintassi) e ne imposta variabili e credenziali finte. |
| `cassaforte-collaudo.js stato\|finta\|vera` | Dice in che stato è la cassaforte della mail del collaudo e la passa dalle credenziali vere a quelle finte e ritorno. Lo scambio avviene dentro il database: i token veri non passano dallo script. Per le prove sa anche togliere il client secret (`senzaSegreto`, come al primo avvio) e metterne uno che il finto Google non riconosce (`segretoSbagliato`), sempre e solo su una cassaforte finta. |
| `prova-funzione-mail.js` | 69 prove della funzione mail contro il finto Google: chi può chiamare (anche il finto Google stesso), mittente scelto, consenso dell'account giusto, verifica dal vivo dei permessi, mittente mostrato diverso da quello in uso, client secret (salvato solo se riconosciuto, mai restituito, ID client respinto, secret custodito non più valido). Mette da parte il consenso vero e lo rimette alla fine. |
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
| `mail` | «Impostazioni email» (mittente precompilato, convalida, registro); senza i permessi del mittente la procedura di invio non compare e vengono chiesti; il consenso di un altro account è rifiutato; il mittente è mostrato e non modificabile; annullare la conferma riporta alla procedura com'era; se il mittente cambia a finestra aperta la mail non parte; risposte anomale del server; il client secret si inserisce e si cambia da «Impostazioni email» a campo sempre vuoto, uno sbagliato viene respinto e richiesto, primo avvio in due passi |

La sezione `xss` scrive sul letto libero «5», crea e cancella righe di prova
(un letto, una tipologia, due backup, due link) e alla fine rimette tutto
com'era. Per essere sicuri che le prove misurino davvero, le si rilancia dopo
aver neutralizzato a mano una protezione (`window._testoHtml = String` oppure
`window._pulisciHtml = String`): devono fallire.

La sezione `mail` gira con le credenziali finte: il banco (`/banco/cassaforte`)
mette da parte il consenso vero e lo rimette alla fine, e al posto della
finestra di Google la prova risponde con un codice che il finto Google capisce.
Nessuna mail parte. A pannello del browser nascosto una finestra già chiusa può
restare nel DOM (il browser non consegna la fine dell'animazione): per questo le
prove guardano la classe `swal2-hide` e non `Swal.isVisible()`.

**Prova di fumo, da fare a mano prima di ogni rilascio** (anche sul sito di
collaudo pubblicato, dove il service worker è attivo):

1. `…/?sloggato` (sul sito pubblicato: una finestra anonima): console pulita,
   bottone «Accedi con Google» visibile, nessun avviso a tutto schermo.
2. `…/?toast=info&msg=%3Cimg%20src%3Dx%20onerror%3D%22alert(document.domain)%22%3E`:
   non deve aprirsi alcuna finestra; a utente collegato il messaggio compare
   nel toast come testo, tale e quale.

## Service worker: banco e prove

Nel banco di prova normale il service worker è spento. Ha un banco suo:

```
node collaudo/strumenti/banco-sw.js 8766
```

Serve il sito sotto `/sito/` come lo servirebbero Cloudflare Pages (`x.html`
rinvia a `x`) o GitHub Pages, con il service worker acceso, e sa fingere i
guasti che il service worker deve reggere: sito che risponde 404 o 500 (i
minuti in cui viene ripubblicato), rete assente, un cancello di accesso che
rinvia ogni richiesta a un altro sito. Sa anche servire la versione oggi in
produzione (`origin/master`), per provare il passaggio da quella alla versione
di lavoro.

Dalla pagina `http://localhost:8766/prove/`: `await window.__proveSw()`, 46
prove nei due stili (`window.__avanzamentoSw()` dice a che punto è). Vanno
rilanciate a ogni modifica di `docs/sw.js` o dell'elenco dei file del sito.

**Controprova**: `await window.__proveSw({ versione: 'produzione' })` esegue le
stesse prove sul service worker della v181, che davanti a un 404 mostra
l'errore e lo salva in cache. Devono fallire (25 su 36): se passano, le prove
non stanno misurando.

## Pubblicazione su Cloudflare

Il sito viene pubblicato anche su Cloudflare Pages, in vista dello spostamento
del sito del reparto: sorgente in un repository privato, sito dietro un
accesso. Si prova tutto qui, con account di collaudo separati, e solo dopo si
ripete in produzione.

- **Cosa viene pubblicato**: i soli file del sito, cioè i file di `docs/`
  registrati in git tranne le cartelle dal nome che comincia con `_`
  (`prepara-sito.js`), più `docs/_headers`. Un file nuovo va quindi registrato
  in git e, se le pagine lo caricano, aggiunto all'elenco di `sw.js`: lo
  verificano i controlli di rilascio.
- **Chi pubblica**: il flusso `.github/workflows/pubblica-cloudflare.yml`, a
  ogni push su `master` che tocca il sito. È lo stesso nei due repository: a
  quale progetto pubblicare lo dicono la variabile `CLOUDFLARE_PROGETTO` e i
  segreti `CLOUDFLARE_API_TOKEN` e `CLOUDFLARE_ACCOUNT_ID` del repository. Le
  chiavi le mette nei segreti di GitHub chi possiede l'account Cloudflare: non
  passano dalla chat né dal PC di sviluppo. Dove mancano, il flusso controlla
  soltanto i file e lo strumento e non pubblica nulla.
- **Lo strumento** (wrangler) ha versione fissa e l'impronta di ogni pacchetto
  in `.github/pubblicazione/package-lock.json`, ed è installato senza
  eseguirne gli script. Per aggiornarlo: nuova versione in `package.json`, poi
  `npm install --package-lock-only --ignore-scripts` in quella cartella (non
  installa nulla, riscrive solo le impronte).
- **L'avviso ai PC**: con la variabile `CLOUDFLARE_AVVISA` a `si` è il flusso
  ad aggiornare `app_version` a pubblicazione fatta. Va accesa solo quando i PC
  usano il sito su Cloudflare; finché usano quello su GitHub l'avviso resta a
  `notify-deploy`. Nel collaudo è accesa dall'08/10/2026, giorno dello scambio
  dei repository: il flusso usa i segreti `SUPABASE_URL` e
  `SUPABASE_SERVICE_ROLE_KEY` del repository privato, che sono quelli del
  progetto di collaudo, messi via API senza mai mostrarli. `pubblicazione.js`
  segue il flusso e controlla l'avviso.
- **Dove**: dall'08/10/2026 il collaudo è pubblicato su
  `https://consegne-collaudo.pages.dev/` (progetto `consegne-collaudo`,
  account Cloudflare di gistech). Dallo scambio dei repository, fatto lo
  stesso giorno, è l'unico sito del collaudo.
- **L'accesso** (Cloudflare Access, dall'08/10/2026): chi apre il sito su
  Cloudflare senza essersi fatto riconoscere non riceve né la pagina né i
  singoli file: viene mandato all'accesso. Com'è configurato:
  - squadra Zero Trust `consegne-collaudo` (piano gratuito): l'accesso
    passa da `consegne-collaudo.cloudflareaccess.com`;
  - **due applicazioni**, perché quella creata dall'interruttore del
    progetto Pages («Accesso in anteprima») copre solo gli indirizzi di
    anteprima `*.consegne-collaudo.pages.dev` e il suo nome host non si
    può cambiare: il sito vero è protetto da una seconda applicazione
    «Self-hosted» sul dominio `consegne-collaudo.pages.dev`, senza
    sottodominio. **Senza la seconda il sito resta aperto a tutti**;
  - chi entra: un criterio con l'elenco degli indirizzi ammessi
    (`autorizzati-collaudo`: chi gestisce il collaudo, la casella del
    reparto, Stefano). Le anteprime tengono il criterio messo da
    Cloudflare, col solo proprietario dell'account e il codice via mail;
  - come si entra: **solo con Google**, con «Applica autenticazione
    immediata»: nessuna pagina di Cloudflare in mezzo, si va dritti a
    Google. Il client Google dell'accesso («Accesso Cloudflare collaudo»)
    è un client a parte nello stesso progetto Google del client dell'app;
    l'indirizzo di ritorno registrato su Google è
    `https://consegne-collaudo.cloudflareaccess.com/cdn-cgi/access/callback`.
    Il segreto di quel client lo incolla in Cloudflare chi possiede
    l'account: non passa dalla chat né dal PC di sviluppo. Provato
    l'08/10/2026 con la casella del reparto, già collegata a Google nel
    browser e già usata per entrare nell'app: l'ingresso è avvenuto senza
    nemmeno un clic (non provato con un account che non ha mai usato
    l'app, né con più account Google collegati: lì Google fa scegliere);
  - ogni quanto: sessione di un mese; cookie non leggibile dagli script
    («Solo HTTP») e con il cookie aggiuntivo di protezione («Abilita
    Binding Cookie»);
  - se Google non dovesse più funzionare si riaccende il codice via mail:
    «Controlli Access» → «Applicazioni» → l'applicazione del sito →
    «Metodi di login» → aggiungere «One-time PIN» (con due metodi
    l'autenticazione immediata si spegne da sé). La console di Cloudflare
    ha un accesso suo, che non dipende da queste impostazioni.

  «Esci» nell'app chiude la sessione dell'app, non quella di Cloudflare:
  da quel browser le pagine restano raggiungibili, i dati no. Dopo ogni
  modifica alle impostazioni di Access si ricontrolla da estraneo, con
  `curl`: il sito, un file (`/js/api.js`) e un indirizzo di anteprima
  devono rimandare tutti a `…cloudflareaccess.com`. Provato l'08/10/2026
  da dentro: sincronizzazione e blocco fra due client, stampa, finestra
  della mail, avviso di aggiornamento, uscita. Resta da provare a mano
  l'importazione dal finto TrakCare (sul nuovo indirizzo il browser chiede
  una volta il permesso per gli appunti).
- **Scambio dei repository** (`scambio-repository.js`): è il passo che toglie
  il codice dal pubblico. Il repository attuale cambia nome in
  `app-consegne-sorgente` e diventa privato (conserva storia, segreti e
  flussi); il nome `app-consegne` passa a un repository pubblico che contiene
  solo la pagina di rinvio, un commit senza storia, chiuso alle scritture.
  `prepara` crea quel repository col nome provvisorio `app-consegne-rinvio` e
  il sito già in linea; `scambia` fa i due cambi di nome, misura per quanti
  secondi il vecchio indirizzo non risponde, rende privato il sorgente e
  sposta il remoto `collaudo` di questa cartella; `annulla` rimette tutto
  com'era. **Dopo lo scambio ogni altra copia della cartella ha il remoto che
  punta al repository pubblico**: per questo è chiuso alle scritture, ma il
  remoto va corretto prima di usarla.

  **Fatto nel collaudo l'08/10/2026.** Tempi: 2 secondi per la prima
  rinomina, 5 per la seconda, 23 perché il vecchio indirizzo passasse alla
  pagina di rinvio, senza risposte «404» nel mezzo: fino a quel momento ha
  continuato a servire l'applicazione. Il sorgente è diventato privato dopo
  29 secondi e GitHub ne ha spento da sé il sito (un account gratuito non
  pubblica da un repository privato). Subito dopo il flusso di Cloudflare ha
  pubblicato dal repository privato e ha avvisato le pagine aperte.
  Lo stesso giorno è stato provato anche il ritorno: `annulla` ha rimesso
  tutto com'era in 39 secondi, di cui circa 30 col vecchio indirizzo in
  «404», e un secondo `scambia` ha richiuso tutto in 28.
  Controlli da rifare **da estraneo, senza credenziali**, dopo ogni scambio
  (li fa tutti `verifica-chiusura.js`, che su un repository ancora pubblico
  deve fallire):
  - `api.github.com/repos/<proprietario>/app-consegne-sorgente`, la pagina
    su `github.com`, un file da `raw.githubusercontent.com`, lo zip del ramo
    e `git ls-remote` devono rispondere 404 o chiedere le credenziali;
  - dal vecchio nome, un file del sorgente chiesto per ramo o per impronta
    del commit deve dare 404, e lo zip deve contenere solo i 7 file della
    pagina di rinvio;
  - `<proprietario>.github.io/app-consegne/js/api.js` deve dare 404 anche
    senza parametri contro la cache, e così
    `<proprietario>.github.io/app-consegne-sorgente/`;
  - prima dello scambio: nessun fork, nessun collaboratore oltre al
    proprietario, nessuna chiave di pubblicazione, nessun aggancio esterno.
    Un fork pubblico resterebbe pubblico anche dopo.

  Lo scambio non richiama le copie già fatte: chi ha scaricato il sorgente
  quando era pubblico lo conserva. In un repository privato di un account
  gratuito le regole («rulesets») non esistono: `stato` lo sa. Per qualche
  minuto dopo lo scambio i file «raw» e lo zip possono rispondere ancora:
  GitHub impiega qualche secondo a propagare il cambio e una risposta presa
  in quel momento resta cinque minuti nella cache del distributore. La
  verifica si lancia quindi due minuti dopo e si ripete dopo dieci.
  Nel repository privato le azioni ammesse nei flussi sono solo quelle
  scritte da GitHub (impostazione del repository, non del codice).
- **Lo schema per la produzione** è in `collaudo/PASSAGGIO-IN-PRODUZIONE.md`:
  che cosa si prepara prima, il passaggio vero, le prove, il ritorno
  indietro e le trappole incontrate.
- **Differenze di Cloudflare Pages**: `print.html` risponde con un rinvio a
  `/print` e `index.html` a `/`; senza `404.html` un indirizzo sconosciuto
  riceverebbe la pagina principale; le intestazioni si decidono in
  `docs/_headers`; il sito sta in radice, non sotto `/app-consegne/`.

## La mail nel collaudo

La funzione `google-token` è la stessa della produzione; nel collaudo le
variabili `GOOGLE_URL_*` le fanno chiamare `google-finto` al posto di Google, e
`GOOGLE_CLIENT_ID` è il client Google del collaudo. `google-finto` gira
richieste a Google e non verifica il JWT: perché non diventi un passaggio aperto
risponde solo sotto un indirizzo che contiene una chiave casuale
(`…/google-finto/k/<chiave>/…`), scritta da `funzioni-collaudo.js configura`
nelle sole variabili delle due funzioni. Decide poi dalle credenziali che riceve:

- **credenziali finte** (client secret e token delle prove) → fa la parte di
  Google e «spedisce» scrivendo in `posta_simulata` (`reale = false`). Conosce
  due client secret finti validi (quello delle prove e lo stesso con `-bis`);
  ogni altro che comincia allo stesso modo è per lui un secret sbagliato, e
  risponde `invalid_client` come farebbe Google;
- **credenziali vere** (il consenso dato davvero dal sito di collaudo) → gira la
  richiesta a Google: la mail **parte davvero**, dal mittente di prova ai
  destinatari di prova, e in `posta_simulata` ne resta una copia (`reale = true`).

La cassaforte `google_oauth` ha quindi tre righe: `reparto` (quella che la
funzione usa), `prova` (le credenziali finte, create da `funzioni-collaudo.js
configura`) e, solo mentre girano le prove, `reparto_vero` (il consenso vero
messo da parte). A riposo la riga `reparto` contiene il consenso vero.

Il **consenso vero** lo dà una persona, una volta, dal sito di collaudo: il
bottone giallo «Autorizza Google» chiede prima il client secret del client OAuth
del collaudo (nella console Google: Credenziali, clic sul nome del client,
«Client secret»; non è l'ID client) e, quando Google lo ha riconosciuto, apre
la finestra di Google, dove va scelto l'account impostato come mittente. Il
secret si cambia in ogni momento da «Impostazioni email». Nel progetto Google Cloud del collaudo devono essere attive le
Gmail API; finché la schermata di consenso è «In fase di test» il mittente deve
essere fra gli utenti di prova e il consenso scade dopo 7 giorni (l'app se ne
accorge e lo richiede prima di mostrare la procedura di invio).

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
