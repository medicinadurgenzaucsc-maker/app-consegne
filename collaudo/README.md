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

Chi è un sito lo dice `js/ambiente.js`, generato dal file del suo reparto
(`reparti/<codice>.json`, vedi `reparti/README.md`): i suoi indirizzi, il suo
database. Sono di collaudo `gistech2026.github.io`, `localhost` e il sito su
Cloudflare `consegne-collaudo.pages.dev` con le sue anteprime. Un indirizzo che
non sta fra quelli del sito non è un ambiente: lì l'app non parte e mostra
«Indirizzo non riconosciuto», quindi **un indirizzo nuovo si aggiunge al campo
`host` del file del reparto** prima di essere usato (per la 11N poi
`node collaudo/strumenti/reparti.js scrivi`). Dalla versione 193 in
`docs/js/api.js` non c'è più la tabella di riserva degli ambienti: senza il
file l'app si ferma e offre «Ricarica».
`docs/` è il sito della 11N: servita così com'è da un altro indirizzo (un
server locale, o GitHub Pages dell'account di collaudo dopo un «annulla» dello
scambio) le pagine arrivano ma l'app si ferma su «Indirizzo non riconosciuto».
Una cornice arancione con «COLLAUDO — DATI FITTIZI»
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
- **Rubrica** («Numeri Telefono»): le tabelle `rubrica_categorie`,
  `rubrica_contatti` e `rubrica_versione`, con le loro due funzioni. Dentro
  c'è **l'elenco vero**, copiato con `copia-rubrica.js` per decisione di
  Stefano dell'08/10/2026: sono numeri di telefono dell'ospedale, non dati di
  pazienti, e servono a provare la rubrica com'è davvero. Il collaudo
  contiene quindi nomi e numeri veri: non va «rimesso a posto» con contatti
  inventati, e non va trattato né mostrato ad altri come se fosse privo di
  dati personali (vedi le Regole in fondo). Finché le stesse tabelle non
  esistono in produzione, `clona-struttura.js confronta` le mostra come
  differenza: è attesa, e non vanno messe fra quelle «solo collaudo».
- **In più rispetto alla produzione**, per scelta: la funzione di RLS
  automatica (`rls_auto_enable`), la tabella `posta_simulata`, la funzione
  `google-finto` che sta fra `google-token` e Google, le variabili `GOOGLE_*`
  che la fanno chiamare, e nella cassaforte `google_oauth` le righe `prova` e
  `reparto_vero` (vedi «La mail nel collaudo»).

## Strumenti (`collaudo/strumenti/`)

I token si leggono dalla configurazione locale di Claude Code (voci MCP
`supabase` e `supabase-collaudo`) e dall'archivio credenziali di Git: nel
repository non c'è alcun segreto.

Dalla v199 (il «ponte» del collaudo) account e repository di GitHub non sono
più costanti scritte negli strumenti: li dice il file del reparto, voce
`github` (`utente`, `sorgente`, `rinvio`, `sitoServizio`; schema in
`reparti/README.md`). Un repository può essere di un'organizzazione diversa
dall'utente, come il sorgente del collaudo dopo il trasferimento
all'organizzazione della società: `gh-api.js` lascia lavorare solo nei
repository che il file nomina, uno per uno e metodo per metodo, crea solo quei
nomi (in un'organizzazione con `/orgs/<nome>/repos`), non segue i rinvii di
GitHub, e da qualunque ambiente rifiuta una scrittura in un repository della
produzione fuori dall'ambiente `produzione`. `pubblica-sito.js` crea il
repository solo con `--crea` e scrive solo in uno appena nato o che contiene
già soltanto i file del sito di servizio. `gh-api.js` tiene una riserva con gli stessi dati, per
quando il file non si legge: chi cambia un account o un repository nel file
cambia anche la riserva (il controllo 14 lo pretende).

| Script | A cosa serve |
|---|---|
| `sb.js` | Accesso ai due database: `collaudo()` legge e scrive solo sul progetto di collaudo, `produzione()` è in **sola lettura**; `per(nome)` dà quello giusto dal codice di un reparto. |
| `rinvio-rubrica.js file\|stato\|pubblica --confermo-produzione` | Il vecchio indirizzo della rubrica (`…github.io/rubrica-gemelli/`, quello delle icone già installate sui telefoni): tiene lì un repository pubblico con la sola pagina di rinvio verso l'app aperta in modo rubrica (`?rubrica=1`) e un service worker che toglie quello di prima con la sua cache. Solo produzione; `stato` è in sola lettura. |
| `reparti.js [mostra <codice>\|controlla\|scrivi]` | I reparti (`reparti/<codice>.json`): li elenca e li convalida, stampa il file `js/ambiente.js` di un reparto, controlla o riscrive `docs/js/ambiente.js` (che è sempre quello della 11N). Il codice di un reparto vale da nome d'ambiente in tutti gli strumenti: `gemelli-11n` è `produzione`. Schema in `reparti/README.md`. |
| `keepalive.js reparto <progetto Cloudflare> <repository>|segnala` | Il flusso Keepalive di ogni reparto (dalla v195): `reparto` trova il reparto dalla variabile `CLOUDFLARE_PROGETTO` del repository (o, senza, dal repository che pubblica `docs/` con GitHub Pages; un repository che non è quello del reparto viene rifiutato) e scrive progetto, indirizzo, chiave pubblica e centrale per i passi del flusso; gli errori vanno sul canale degli errori, non fra le uscite. `segnala` manda a fine giro l'esito (`ok` solo se tutti i passi sono riusciti) e il primo passo fallito alla funzione `keepalive` della centrale. Nessun segreto passa di qui; se la centrale non risponde il giro non diventa rosso. |
| `clona-struttura.js ddl\|applica\|confronta` | Ricostruisce la struttura della produzione nel collaudo e ne verifica l'identità. |
| `copia-config.js` | Ricopia la configurazione non personale. |
| `confronta-forme.js` | Confronta la forma dei contenuti (tag e classi, mai il testo) fra i due ambienti. |
| `gettone-prova.js [ore] [cartella]` | Token di sessione per l'utente fittizio delle prove automatiche (vale solo nel collaudo). |
| `banco-sw.js [porta]` | Banco del service worker: il sito servito come farebbero Cloudflare Pages o GitHub Pages, col service worker acceso e i guasti simulati (vedi «Service worker: banco e prove»). |
| `prepara-sito.js [cartella] [compressa] [--reparto <codice>]` | Elenca (o copia in una cartella) i soli file del sito, **per un reparto**: genera il suo `js/ambiente.js` e controlla che nel sito non ci siano dati di un altro reparto. Senza `--reparto` è il sito di `docs/` così com'è (la 11N). Con `compressa` ne produce la copia compressa, la forma pubblicata su Cloudflare, e la controlla. La cartella viene svuotata: può essere solo `_sito` o una sua sottocartella. |
| `sito-pubblico.js <reparto> stato\|pubblica [--crea]` | Il «livello 0»: pubblica il sito di un reparto, in copia compressa, in un repository **pubblico che contiene solo quei file** (voce `pubblicazione.sitoPubblico` del suo file), servito da GitHub Pages; il sorgente resta privato. A ogni pubblicazione un commit in più con l'intero sito, mai una spinta forzata, poi aspetta che il sito serva gli stessi file. Pubblica solo ciò che è registrato in git, solo in un repository appena nato o che è già il sito di quel reparto, e lo crea solo con `--crea`. `stato` è di sola lettura e dice se la versione servita è quella della cartella. Da riga di comando solo per un reparto di collaudo. Oggi: `gistech2026/collaudo-sito`, sito `https://gistech2026.github.io/collaudo-sito/`; non si aggiorna da solo, va rilanciato dopo `pubblicazione.js`. |
| `prepara-rinvio.js <nuovo indirizzo> [cartella]` | Genera i file della pagina di rinvio (ciò che resta al vecchio indirizzo quando il sito si sposta), dai modelli in `rinvio/`. |
| `scambio-repository.js <ambiente> stato\|prepara\|scambia\|annulla` | Prova generale dello scambio dei repository nel collaudo: sorgente in un repository privato, al vecchio nome solo la pagina di rinvio (vedi «Pubblicazione su Cloudflare»). Dalla v199 guarda ciascun repository nel suo account: se il sorgente sta in un'altra organizzazione lo scambio è per forza già fatto, e «scambia» da capo e «annulla» si rifiutano (un cambio di nome non cambia il proprietario). |
| `pubblicazione.js [ambiente] [avvia]` | Segue il flusso «Pubblica su Cloudflare» del repository di collaudo fino alla fine, ne mostra i passi e controlla che l'avviso ai PC sia arrivato in `app_version`; con `avvia` lo fa ripartire a mano sullo stesso commit. Non scrive nel database. |
| `verifica-chiusura.js collaudo\|produzione [indirizzo] [--archivi]` | Prova **da estraneo**, senza credenziali, che il sorgente non si scarichi più da nessuna strada, che al vecchio indirizzo resti solo la pagina di rinvio e che il sito nuovo chieda l'accesso: 29 prove, più 4 sugli archivi pubblici di terzi. Dove lo scambio non è stato fatto deve fallire: è la sua controprova. Dalla v199: il rinvio deve portare **esattamente** alla pagina d'accesso della squadra che il file del reparto nomina (`pubblicazione.cloudflare.squadra`), per quel nome; se il sorgente sta in un altro account deve mancare a tutti e due i nomi; se il file dichiara un dominio le prove del sito si rifanno lì; un dominio che non risolve vale «NON SO» (lo si chiede anche a un risolutore pubblico), e solo con `--dominio-non-collegato`, se non esiste nemmeno lì, si annota senza contarlo. Finché il dominio del collaudo non è collegato: `verifica-chiusura.js collaudo --dominio-non-collegato`. |
| `copia-rubrica.js collaudo [confronta\|prima-copia]`, `copia-rubrica.js produzione confronta`, `copia-rubrica.js produzione prima-copia --confermo-produzione` | Copia l'elenco della rubrica telefonica dal suo vecchio database nelle tabelle `rubrica_`, in una sola transazione e con gli stessi identificativi; `confronta` dice solo se le due copie coincidono. Nel collaudo, senza azione, svuota e riempie. **In produzione non svuota mai**: esiste solo `prima-copia`, che si rifiuta se la destinazione ha già righe (fatta il 09/10/2026). Non stampa nomi né numeri: solo conteggi e impronte. La sorgente si legge soltanto; quando il vecchio database della rubrica verrà spento lo strumento non avrà più da dove leggere. |
| `banco.js [porta] [compressa]` | Banco di prova locale (con `compressa` serve la copia compressa del sito): serve l'app così com'è nella cartella di lavoro, collegata al collaudo e già «dentro» con l'utente fittizio. |
| `inventario-markup.js [produzione\|collaudo]` | Elenco di tag, attributi, classi e proprietà di stile presenti nei campi delle schede (solo nomi e conteggi, mai il testo): serve a tarare e a ricontrollare il filtro dell'HTML (`docs/js/sanifica.js`). |
| `pubblica-sito.js` | Pubblica il sito di servizio (finto TrakCare e informativa) nel repository che il file del collaudo dice in `github.sitoServizio`, oggi `gistech2026/collaudo` (sito `gistech2026.github.io/collaudo/`); se manca lo crea, anche in un'organizzazione. |
| `controlli-rilascio.js [riferimento]` | Controlli statici prima di ogni pubblicazione: sintassi, segnalibri collassati e loro versione, tag script al completo, impronta della libreria del filtro, service worker, uscita solo locale (`signOut` con `scope: 'local'`), versione nel menu uguale a `CACHE_NAME`. Confronta con `origin/master` (la produzione) se non si indica altro. |
| `confronta-filtro.js [produzione\|collaudo] [--backup N]` | Prima di pubblicare una versione che introduce o cambia il filtro dell'HTML: controlla che le schede (e gli ultimi N backup) non usino tag, attributi, classi o stili che il filtro toglierebbe. L'analisi gira dentro il database: escono solo nomi e conteggi, mai il testo. `--autoprova` verifica lo strumento stesso su un contenuto costruito apposta. |
| `prova-uscita.js` | Dimostra con sessioni vere che «Esci» su un dispositivo non scollega gli altri: utente provvisorio, tre sessioni, uscita `local` dalla prima, le altre due restano; controprova con `global`. L'utente viene eliminato alla fine. |
| `funzioni-collaudo.js pubblica\|configura` | Pubblica `google-token` e `google-finto` nel collaudo (dopo averne controllato la sintassi) e ne imposta variabili e credenziali finte. |
| `cassaforte-collaudo.js stato\|finta\|vera\|verifica` | Dice in che stato è la cassaforte della mail del collaudo e la passa dalle credenziali vere a quelle finte e ritorno. Lo scambio avviene dentro il database: i token veri non passano dallo script. Per le prove sa anche togliere il client secret (`senzaSegreto`, come al primo avvio) e metterne uno che il finto Google non riconosce (`segretoSbagliato`), sempre e solo su una cassaforte finta. |
| `prova-funzione-mail.js` | 69 prove della funzione mail contro il finto Google: chi può chiamare (anche il finto Google stesso), mittente scelto, consenso dell'account giusto, verifica dal vivo dei permessi, mittente mostrato diverso da quello in uso, client secret (salvato solo se riconosciuto, mai restituito, ID client respinto, secret custodito non più valido). Mette da parte il consenso vero e lo rimette alla fine. |
| `supabase-accesso.js`, `github-accesso.js` | Autorizzazione «a codice» dei due account di collaudo, senza far passare chiavi dalla chat. L'account di GitHub è quello di `github.utente` del file del collaudo: un accesso autorizzato da un altro account non viene registrato. |

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
| `http://localhost:8765/rubrica/` | la rubrica aperta **da sola**: deve restare sul messaggio fisso («La rubrica si apre dal menu…») senza chiedere nulla al database |
| `http://localhost:8765/?rubrica=1` | il **modo rubrica**: dopo l'accesso solo la rubrica a tutto schermo, nessun dato dei pazienti scaricato; con `&sloggato` la schermata di accesso intitolata «Numeri Telefono» |
| `http://localhost:8765/trak-finto/` | il finto TrakCare |

Nel banco il service worker è spento e nessun file viene copiato o modificato.
Un file solo non è quello di `docs/`: `js/ambiente.js`. Quello nel repository è
della 11N e non ammette `localhost`; il banco serve al suo posto quello del
reparto di collaudo, generato in memoria da `reparti/collaudo.json`. Per le
prove sa servire anche file d'ambiente **finti** (mai quello di un reparto
vero): `?ambiente=senza` toglie il tag dalla pagina, come una pagina della
versione di prima; `assente`, `sintassi` e `nullo` fingono un file che non
arriva, e l'app deve fermarsi offrendo «Ricarica»; `rotto` e `altro-sito` un
file sbagliato, davanti al quale l'app deve fermarsi.
Con `node collaudo/strumenti/banco.js 8767 compressa` il banco serve invece la
**copia compressa** del sito, quella pubblicata su Cloudflare, preparata
all'avvio in `_sito/banco`: dopo una modifica a `docs/` va riavviato. Le prove
vanno fatte passare anche lì prima di ogni rilascio.
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
| `rubrica` | la rubrica «Numeri Telefono» (`docs/rubrica/`): aperta da sola, o dentro l'app senza accesso, non chiede nulla al database; con la sola chiave pubblica non si legge e non si scrive nulla; contatti e categorie ostili (ma ammessi dai vincoli) compaiono tali e quali in elenco, ricerca, finestre e conferme, senza che nulla venga eseguito; solo un numero fatto di cifre è un collegamento; esportazioni senza formule; errori del database tradotti, e tabelle assenti dette con chiarezza; riapertura del riquadro con una sola lettura; 401 e rinnovo della sessione; richiesta senza risposta interrotta; **scritture fatte dall'interfaccia** (nuovo, modifica, elimina, categorie: aggiunta, rinomina, riordino, eliminazione) sempre e solo sulle righe di prova; due PC sullo stesso contatto; rete che cade durante un salvataggio; uscita con un lavoro a metà; memoria del browser ostile; all'uscita la rubrica viene svuotata |
| `centrale` | il battito verso la centrale (dalla v194) e i moduli (dalla v195; `dimissioni` dalla v196): il file del collaudo la nomina; il battito parte con i soli sette campi, senza chiavi, sessione né cookie, a ogni caricamento ma al più uno ogni 15 secondi, e a pagina aperta una volta all'ora (il giro parte dal primo battito vero, una volta sola, mai dalle prove), e senza rete rimette l'ora di prima; in modo rubrica il nome dice solo browser e sistema; senza la centrale, o con un indirizzo scritto male, non parte nulla. I moduli della risposta spengono e riaccendono le funzioni (classe su `<html>`, memoria del browser), quelli scritti male non cambiano nulla, una chiave sconosciuta si ignora e una chiave nota che manca vale accesa (dalla v196); con i moduli spenti ogni varco mostra l'avviso e non apre nulla, e la mail non chiede nulla al suo server; con `dimissioni` spento spariscono le due voci di menu e il bottoncino giallo; col laboratorio spento spariscono i blocchi dell'alambicco e `lab_esami` non si legge; la pagina in modo rubrica con la rubrica spenta mostra il pannello di testo e non carica la rubrica. Le richieste sono intercettate e nulla viene scritto nella centrale, tranne una GET vera che non scrive; ora del battito e moduli del dispositivo tornano com'erano. Le altre sezioni girano con i cinque moduli accesi. |

La sezione `xss` scrive sul letto libero «5», crea e cancella righe di prova
(un letto, una tipologia, due backup, due link) e alla fine rimette tutto
com'era. Per essere sicuri che le prove misurino davvero, le si rilancia dopo
aver neutralizzato a mano una protezione (`window._testoHtml = String` oppure
`window._pulisciHtml = String`): devono fallire.

La sezione `rubrica` lavora su una rubrica che è **quella vera** (vedi sopra),
quindi ha tre regole, che valgono anche per chi le aggiunge una prova:

1. tocca **solo righe create da lei**, in `rubrica_contatti` e
   `rubrica_categorie`, col nome che comincia con `PROVA-AUTOMATICA`, e alla
   fine le toglie; una riga che c'era già non viene mai modificata né
   cancellata, e i filtri sono sempre stretti (il nome di prova, oppure
   l'identificativo di una riga di prova);
2. in un esito **non finisce mai il testo dell'elenco**: solo le righe di
   prova, conteggi, nomi di attributi, metodo e tabella di una richiesta;
3. le prove che usano l'interfaccia (Salva, Modifica, Elimina, categorie)
   girano sotto il **salvagente**: ogni scrittura che parte dalla rubrica
   viene guardata prima di partire (metodo, tabella, filtro, dati), e se non
   riguarda soltanto righe di prova viene fermata e fa fallire la prova. Così
   un filtro sbagliato in una modifica o in un'eliminazione si vede senza che
   un contatto vero venga toccato.

Apre il riquadro vero e alla fine rimette le preferenze del browser com'erano.
La pagina della rubrica dal banco riceve solo la raccolta degli errori
(`/banco-errori.js`, un file a parte perché la sua regola CSP non ammette script
scritti in pagina), mai la sessione: quella la chiede alla pagina dell'app.
Fra gli errori raccolti c'è anche ciò che la regola CSP ha bloccato. La
controprova, a riquadro aperto:
`document.getElementById('rubricaIframe').contentWindow.esc = String`, poi di
nuovo le prove: devono fallire. La sezione dura qualche minuto: una prova
aspetta apposta i 12 secondi del tempo massimo di una richiesta, un'altra i 6
secondi di un rinnovo della sessione che non risponde.

Due prove dicono qualcosa del **database**, non della pagina. La prova 14
pretende che il numero di `rubrica_versione` **cresca a ogni modifica**, anche
a più modifiche nello stesso secondo: fa cinque modifiche di fila a un
contatto di prova e vuole una salita di almeno cinque. Nel collaudo è così
dall'08/10/2026 (migrazione `rubrica_versione_sempre_crescente`: il trigger
scrive `ts = greatest(ts + 1, secondi dell'orologio)`), e la stessa funzione
va creata così anche in produzione; se la prova fallisce va guardato il
trigger. «Chi usa l'app non può scrivere la versione» fallisce se quella
tabella è scrivibile da chi è collegato, oppure se il rifiuto viene detto
«account non autorizzato» invece che «il database non permette».

`await window.__prove({ sezioni: ['rubrica'], fuoco: true })` aggiunge la prova
del fuoco: scheda in modifica sul letto libero «5», «Numeri Telefono» aperto,
il fuoco nella rubrica non deve chiudere la modifica, e chiudendo il riquadro
deve tornare nella scheda. Vuole la finestra del browser **in primo piano**
(nel pannello automatico il fuoco non è affidabile, e la prova lo dice invece
di passare a vuoto); altrimenti la stessa prova si fa a mano.

Restano da provare **a mano**, nel riquadro, su Chrome o Edge e sull'iPad,
guardando la console: «Esporta tutto in CSV» e «Salva nei contatti del
telefono» (il file deve arrivare davvero, e il CSV aprirsi in colonne),
«Copia numero» incollando altrove, la chiamata da un numero, il
**trascinamento** delle categorie col dito o col mouse (la prova automatica
dice nel suo nome se ha dovuto chiamare la funzione al posto del gesto), e un
giro completo crea / modifica / elimina su un contatto di prova.

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
3. `…/rubrica/` aperta da sola in una scheda: solo il messaggio «La rubrica si
   apre dal menu “Numeri Telefono” dell'app», console pulita, nessuna richiesta
   verso Supabase. Poi dall'app, a utente collegato: «Numeri Telefono» mostra i
   contatti e la console resta pulita (nessun avviso della regola CSP).

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

Dalla pagina `http://localhost:8766/prove/`: `await window.__proveSw()`, 62
prove nei due stili (`window.__avanzamentoSw()` dice a che punto è). Vanno
rilanciate a ogni modifica di `docs/sw.js` o dell'elenco dei file del sito.
Otto riguardano la cornice della rubrica: davanti a un guasto del sito la
pagina di «Numeri Telefono» deve arrivare dalla cache, e davanti a un cancello
di accesso la cornice deve ricevere la copia in cache invece del rinvio (che
in una cornice non porterebbe da nessuna parte).

**Controprova**: `await window.__proveSw({ versione: 'produzione' })` esegue le
stesse prove sul service worker di `origin/master`. Finché la produzione era la
v181, che davanti a un 404 mostrava l'errore e lo salvava in cache, dovevano
fallire (25 su 36, misurate allora). **Dalla v187 in produzione c'è il service
worker nuovo e la controprova passa anche lei**: non dice più se le prove
misurano. Chi ne ha bisogno la rifà contro un commit vecchio.

Il sito del banco è quello del reparto di collaudo: `js/ambiente.js` è generato
in memoria (anche per la versione in produzione, quando ha il file), e ogni
pagina arrivata deve anche riconoscere il proprio indirizzo. Senza, le prove
passerebbero su un'app ferma su «Indirizzo non riconosciuto».

`await window.__proveSw({ versione: 'compressa' })` esegue le 62 prove sulla
copia compressa del sito: è la forma in cui il service worker arriva ai PC da
Cloudflare.

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
- **La copia compressa**: su Cloudflare non vanno i file di `docs/` così come
  sono ma una copia compressa, prodotta dal flusso con
  `.github/compressione/comprimi.js` (terser, versione fissa e impronte,
  installato senza eseguirne gli script). È lo stesso programma, più scomodo
  da leggere e da copiare: via i commenti e gli spazi, nomi **locali**
  accorciati, negli script dei file e in quelli scritti in pagina; dei
  commenti HTML e CSS resta il segno vuoto. Non cambiano mai i nomi globali
  (pagine, gestori scritti negli attributi e file diversi si chiamano fra
  loro per nome), la struttura delle pagine, i file delle librerie. `docs/`
  resta com'è: è lì che si lavora, ed è ciò che serve GitHub Pages.

  La copia si controlla da sé e, se un controllo fallisce, il flusso si ferma
  senza pubblicare: ogni script compila, la struttura delle pagine e i nomi
  globali sono identici, non compare nessun indirizzo di rete che non sia nel
  sorgente, i due segnalibri di TrakCare restano funzioni intere senza
  sequenze `%XX`, il service worker ha la stessa versione e lo stesso elenco.
  Gli stessi controlli girano in locale con `controlli-rilascio.js`, dopo
  aver installato lo strumento una volta (`npm ci --ignore-scripts` dentro
  `.github/compressione`). Il flusso stampa l'impronta di ogni file
  pubblicato.

  **Si prova nella forma in cui si pubblica**: `banco.js 8767 compressa` con
  le prove dentro l'app, e `__proveSw({ versione: 'compressa' })`. Dalla
  v186 (08/10/2026): il sito passa da 1347 a 909 KB.

  Da sapere: sul sito compresso i segnalibri nascono dal codice compresso;
  nel service worker la riga della versione diventa
  `CACHE_NAME="consegne-vNNN"`; gli errori annotati nel registro portano
  posizioni e nomi della copia compressa, non del sorgente.
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
- **Scambio dei repository** (`scambio-repository.js`; i comandi che cambiano
  qualcosa vogliono l'ambiente scritto, per esempio
  `scambio-repository.js collaudo scambia`): è il passo che toglie
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
pubblica, con l'informativa sull'accesso, nel repository di
`github.sitoServizio` del file del collaudo (oggi `gistech2026.github.io/collaudo/`;
il link «privacy» del client Google del collaudo punta lì).

## Regole

- Mai copiare nel collaudo righe di pazienti veri, nemmeno anonimizzate: il
  testo libero identifica comunque.
- La rubrica del collaudo è **quella vera**, per decisione di Stefano
  dell'08/10/2026 (numeri di telefono dell'ospedale, non dati di pazienti): non
  si sostituisce con contatti inventati. In cambio: le prove toccano solo le
  righe che creano loro (`PROVA-AUTOMATICA…`), non modificano mai una riga
  esistente, e non riportano mai nomi o numeri dell'elenco in un esito, in un
  messaggio o in una conversazione: solo conteggi, impronte e righe di prova.
- Dagli strumenti la produzione si legge soltanto.
- Il repository di collaudo non ha segreti: i due workflow sono attivi solo nel
  repository di produzione.
- Una modifica al database che crea tabelle scrive sempre la RLS in modo
  esplicito: nel collaudo verrebbe accesa comunque in automatico, in produzione no.
