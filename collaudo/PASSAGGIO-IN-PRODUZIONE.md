# Passaggio del sito: schema per la produzione

Come si porta il reparto dal sito pubblico su GitHub al sito su Cloudflare con
l'accesso, e il sorgente in un repository privato. È la sequenza provata nel
collaudo l'08/10/2026, riscritta per rifarla in produzione: **prima si prepara
tutto, senza che i PC del reparto se ne accorgano; poi il passaggio vero dura
pochi minuti.**

Stato: nel collaudo è tutto fatto e provato, compreso il ritorno indietro. In
produzione non è stato toccato nulla.

## Regole che valgono per tutto il passaggio

- In produzione ogni passo si fa solo con l'ok di Stefano, uno alla volta.
- Registrazioni, accessi e chiavi li fa Stefano. Le chiavi non passano dalla
  chat né dal PC di sviluppo: le incolla lui dove servono.
- Le voci delle pagine di Cloudflare e di Google si leggono dallo schermo prima
  di nominarle: cambiano spesso e la documentazione è vecchia. Dove si può, le
  impostazioni le cambia Claude nel browser di Stefano e lui guarda.
- Il sito di produzione non si apre mai nel pannello del browser dell'app
  Claude: le prove da collegato si fanno dal browser di Stefano.
- Sui pazienti veri non si scrive nulla di finto: le prove di scrittura si
  fanno su un letto libero e si rimette tutto com'era.

## Com'è fatto il risultato

| | Prima | Dopo |
|---|---|---|
| Sito | `<proprietario>.github.io/app-consegne/`, aperto a tutti | `<progetto>.pages.dev`, dietro l'accesso con Google |
| Sorgente | repository pubblico `app-consegne` | repository **privato** `app-consegne-sorgente` |
| Codice che arriva al browser | i file di `docs/` così come sono | copia compressa: senza commenti, nomi locali accorciati |
| Vecchio indirizzo | l'applicazione | una pagina che rimanda a quello nuovo |
| Chi pubblica | GitHub Pages | il flusso «Pubblica su Cloudflare» del repository privato |
| Avviso «Update» | flusso `notify-deploy` | lo stesso flusso di Cloudflare (`CLOUDFLARE_AVVISA = si`) |

Nel collaudo: proprietario `gistech2026`, progetto `consegne-collaudo`, squadra
di accesso `consegne-collaudo`. In produzione il proprietario è
`medicinadurgenzaucsc-maker`; nome del progetto e della squadra li sceglie
Stefano (fase A1).

## Tempi misurati nel collaudo

| Cosa | Tempo |
|---|---|
| Una pubblicazione su Cloudflare, dal comando alla fine | circa 1 minuto |
| Impostazioni dell'accesso su Cloudflare, fatte nel browser | 15–20 minuti la prima volta |
| Scambio dei repository: il vecchio indirizzo passa alla pagina di rinvio | 23 e 28 secondi nelle due prove, nessun errore nel mezzo |
| Sorgente privato | 29 e 34 secondi dall'inizio |
| Ritorno indietro (`annulla`) | 39 secondi, di cui circa 30 con il vecchio indirizzo che risponde «404» |
| Prove da estraneo (`verifica-chiusura.js`) | 1 minuto, 3 con gli archivi |

## Fase A. Preparazione: nulla cambia per i PC

Si fa nei giorni prima. Finché non si arriva alla fase B il reparto continua a
usare il sito di sempre.

### A1. Decisioni di Stefano

1. **Nome del progetto su Cloudflare**: diventa l'indirizzo che userà il
   reparto, `https://<nome>.pages.dev/`. Deve essere libero su Cloudflare.
2. **Chi può entrare** nel sito: elenco di indirizzi Google. Nel collaudo sono
   tre. In produzione almeno la casella del reparto e quello di Stefano.
3. **La sera del passaggio**: un momento tranquillo, con Stefano presente.

### A2. Account e protezione, a cura di Stefano

1. Account Cloudflare della produzione, intestato alla casella del reparto.
   Se ci si registra con «Google» l'account nasce senza password e la verifica
   in due passaggi non si può accendere: prima ci si dà una password da
   `dash.cloudflare.com/forgot-password`, poi si accende la verifica.
2. Verifica in due passaggi accesa anche sull'account GitHub del reparto e
   sull'account Google del reparto. Dopo il passaggio il lucchetto di pagine e
   dati è quell'account Google.

### A3. Codice e database, a cura di Claude

1. Nel collaudo: aggiungere a `_AMBIENTI.produzione.host`, in
   `docs/js/api.js`, l'indirizzo `'<nome>.pages.dev'`. Senza, il flusso si
   rifiuta di pubblicare e l'app sul sito nuovo si ferma con «Indirizzo non
   riconosciuto». In produzione **niente voce col punto**: le anteprime di
   Cloudflare non devono poter lavorare sui pazienti veri. Versione nuova,
   controlli di rilascio, prove nel banco.
2. Con l'ok: riga `AMBIENTE = produzione` nella tabella `impostazioni` del
   database di produzione. Oggi manca: l'app funziona lo stesso, ma è quella
   riga a fermarla se un sito parla col database sbagliato.
3. Con l'ok: rilascio in produzione sul sito di oggi. Deve arrivare ai PC
   **almeno una notte prima** del passaggio: porta il service worker che regge
   gli errori, quello che rende invisibile il cambio. I PC lo ricevono con
   «Update» o col ricaricamento delle 04:00.
   Attenzione: questo rilascio porta nel repository ancora pubblico anche
   strumenti e documentazione del passaggio. Per ridurre la finestra, rilascio
   e scambio si fanno a un giorno di distanza, non di più.

### A4. Chiave di pubblicazione, a cura di Stefano

1. Su Cloudflare, nei token API del profilo: creare un token partendo da zero
   con il solo permesso «Cloudflare Pages: Edit», che sta nel gruppo delle
   piattaforme per sviluppatori. Le voci si rileggono dallo schermo.
2. L'identificativo dell'account è nell'indirizzo della console:
   `dash.cloudflare.com/<identificativo>/…`.
3. Nel repository su GitHub, nei segreti dei flussi: `CLOUDFLARE_API_TOKEN` e
   `CLOUDFLARE_ACCOUNT_ID`. Nelle variabili: `CLOUDFLARE_PROGETTO` con il nome
   scelto. `CLOUDFLARE_AVVISA` resta spenta fino alla fase B.

### A5. Prima pubblicazione su Cloudflare, a cura di Claude

1. Avvio a mano del flusso «Pubblica su Cloudflare». Il progetto nasce da
   solo, col ramo `master` come ramo di produzione.
2. Subito dopo si passa ad A6: finché l'accesso non è configurato il sito
   nuovo è aperto a chi ne conosce l'indirizzo.

### A6. L'accesso, nel browser di Stefano

Percorsi della console, dove `<id>` è l'identificativo dell'account:
`dash.cloudflare.com/<id>/one/access-controls/apps`, `…/policies`,
`…/settings`, `…/one/integrations/identity-providers`,
`…/one/reusable-components/custom-pages`.

1. Attivare Zero Trust: nome della squadra, piano gratuito. Da qui nasce
   l'indirizzo `https://<squadra>.cloudflareaccess.com`.
2. Nel progetto Pages, «Impostazioni», «Generale», «Accesso in anteprima»,
   «Limita le anteprime». Crea l'applicazione `<nome> - Cloudflare Pages`, che
   copre **solo** gli indirizzi di anteprima `*.<nome>.pages.dev`: il suo nome
   host non si può cambiare. Si lascia com'è: solo il proprietario, codice via
   mail, 24 ore.
3. «Controlli Access», «Applicazioni», «Crea nuova applicazione»,
   «Self-hosted e privata»: «Dominio» `<nome>.pages.dev`, «Sottodominio»
   vuoto. **È questa a proteggere il sito vero: senza, resta aperto.**
4. Criterio dell'applicazione: azione «Allow», «Includi», «Emails», gli
   indirizzi decisi in A1.
5. «Durata sessione»: «1 month».
6. «Impostazioni aggiuntive» dell'applicazione: cookie «Solo HTTP» e «Abilita
   Binding Cookie».
7. «Componenti riutilizzabili», «Pagine personalizzate»: nome
   dell'organizzazione mostrato nella pagina di accesso.
8. Provider Google, vedi A7. Poi nell'applicazione del sito, «Metodi di
   login»: spento «Accetta tutti i provider di identità disponibili», in
   «Scegli i provider di identità disponibili per questa applicazione» resta
   solo «Google - google», acceso «Applica autenticazione immediata». Con un
   solo metodo l'interruttore si accende da sé.

### A7. Google, a cura di Stefano con Claude che prepara

Nella console di Google bisogna prima controllare in alto **account e
progetto**: nel browser di Stefano il primo account è un altro, con un
progetto che non c'entra e che non va toccato. Il progetto giusto è quello che
contiene già il client dell'app di quell'ambiente.

1. **Client dell'app**, quello già esistente: aggiungere a «Origini JavaScript
   autorizzate» l'indirizzo del sito nuovo, `https://<nome>.pages.dev`. Senza,
   il pulsante «Accedi con Google» dell'app non funziona sul sito nuovo. La
   vecchia origine si toglie solo a passaggio concluso.
2. **Client dell'accesso**, nuovo, di tipo «Applicazione web», nome «Accesso
   Cloudflare …»: origine `https://<squadra>.cloudflareaccess.com`, indirizzo
   di reindirizzamento
   `https://<squadra>.cloudflareaccess.com/cdn-cgi/access/callback`. Claude
   compila, Stefano preme «Crea».
3. Su Cloudflare, «Integrazioni», «Provider di identità», aggiungere Google:
   identificativo e segreto del client li incolla Stefano.

Stando nello stesso progetto Google del client dell'app, chi ha già usato
l'app con quell'account non deve dare altri consensi: nel collaudo la casella
del reparto è entrata senza nemmeno un clic.

### A8. Prove sul sito nuovo, prima del passaggio

I due siti convivono sullo stesso database: si può provare con calma.

- Da estraneo, con `curl`: sito, un file e un indirizzo di anteprima devono
  rimandare tutti a `…cloudflareaccess.com`.
- Da dentro, dal browser di Stefano: l'ambiente dichiarato dall'app, la
  versione nel menu della rotellina, la stampa, la finestra della mail senza
  inviare, la sincronizzazione fra due dispositivi su un letto libero,
  un'importazione da TrakCare. Al primo uso sul nuovo indirizzo il browser
  chiede una volta il permesso per gli appunti.
- «Esci» chiude la sessione dell'app, non quella di Cloudflare.
- Le prove automatiche si fanno passare anche sulla copia compressa, che è
  la forma pubblicata su Cloudflare: `banco.js 8767 compressa` e
  `__proveSw({ versione: 'compressa' })`.

### A9. Pagina di rinvio e accesso degli strumenti

1. Accesso degli strumenti all'account GitHub della produzione, col metodo «a
   codice»: Stefano digita un codice sul sito di GitHub, la chiave finisce
   nell'archivio credenziali di Windows e non viene mai mostrata.
2. `scambio-repository.js prepara https://<nome>.pages.dev/`: crea il
   repository pubblico provvisorio `app-consegne-rinvio`, 7 file, un commit
   senza storia, sito già in linea, chiuso alle scritture.

Gli strumenti oggi lavorano solo sull'account di collaudo: vanno estesi alla
produzione prima del giorno, vedi «Da preparare negli strumenti».

## Fase B. Il passaggio: pochi minuti

1. **Controlli di partenza**: cartella di lavoro pulita e allineata al remoto;
   sul repository nessun fork, nessun collaboratore oltre al proprietario,
   nessuna chiave di pubblicazione, nessun aggancio esterno;
   `scambio-repository.js stato`; il sito nuovo risponde.
   Non si interrogano i file «raw» né lo zip nei minuti prima dello scambio.
2. **`scambio-repository.js scambia`**: due cambi di nome, il vecchio indirizzo
   passa alla pagina di rinvio, il sorgente diventa privato, il remoto della
   cartella viene spostato. Mezzo minuto.
3. **Avviso ai PC**: variabile `CLOUDFLARE_AVVISA = si`, poi
   `pubblicazione.js avvia` e controllo che `app_version` porti il commit
   pubblicato. In produzione i due segreti di Supabase ci sono già.
4. **Prove da estraneo**: `verifica-chiusura.js <ambiente>` due minuti dopo lo
   scambio, e di nuovo dopo dieci con `--archivi`. Devono passare tutte.
5. **Azioni ammesse nei flussi** del repository privato: solo quelle scritte
   da GitHub.

Cosa vedono i PC: chi ha l'app aperta continua a lavorare; al primo
ricaricamento il vecchio indirizzo lo porta su quello nuovo, dove rifà **un
accesso con Google** perché la sessione è legata all'indirizzo. Segnalibri e
icona installata vanno aggiornati con calma: il rinvio resta.

## Fase C. Dopo

- Togliere la vecchia origine dal client Google dell'app.
- In un rilascio successivo, quando il ritorno indietro non serve più:
  togliere il vecchio indirizzo dall'elenco dell'ambiente in `docs/js/api.js`.
- Dopo ogni cambio alle impostazioni di sicurezza dell'account Google che ha
  dato il permesso della mail, controllare che il permesso valga ancora. Nel
  collaudo: `cassaforte-collaudo.js verifica`. L'08/10/2026 l'accensione
  della verifica in due passaggi non l'ha toccato.
- Cambiare la chiave di GitHub scritta nel remoto `origin` di questa cartella
  con un accesso «a codice», e revocare quella vecchia.
- Aggiornare `CLAUDE.md` e questo documento con nomi e date veri.
- Ricontrollare dopo qualche giorno `verifica-chiusura.js … --archivi`.

## Ritorno indietro

| Se va storto | Rimedio |
|---|---|
| Il sito nuovo non funziona, prima dello scambio | Niente da fare: i PC usano ancora quello vecchio |
| Google non fa più entrare | Nell'applicazione del sito, «Metodi di login», aggiungere «One-time PIN»: torna il codice via mail |
| Lo scambio lascia il vecchio indirizzo in errore | `scambio-repository.js annulla`: in 39 secondi torna tutto com'era, sorgente di nuovo pubblico |
| Una versione nuova dà problemi dopo lo scambio | Si ripubblica il commit precedente: `git revert`, push, il flusso pubblica e avvisa |

## Trappole viste nel collaudo

- L'interruttore di Pages protegge solo le anteprime: serve la seconda
  applicazione.
- Con due metodi di accesso l'autenticazione immediata si spegne.
- Appena acceso l'accesso immediato, basta aprire il sito in un browser con la
  sessione Google attiva perché il passaggio da Google avvenga da solo.
- Dietro l'accesso il sito non si legge più con `curl`: la versione servita si
  guarda dall'esito del flusso o da un browser già entrato.
- Rendere privato un repository di un account gratuito ne spegne il sito: è
  per questo che il vecchio indirizzo passa a un altro repository.
- Nei repository privati di un account gratuito le regole di protezione dei
  rami non esistono.
- Per qualche minuto dopo lo scambio i file «raw» e lo zip possono rispondere
  ancora dalla cache di GitHub: spariscono entro cinque minuti.
- Dopo lo scambio un push verso il vecchio nome finirebbe nel repository
  pubblico: per questo è chiuso alle scritture. Ogni altra copia della
  cartella ha il remoto da correggere.
- Il pannello del browser dell'app Claude non apre finestre nuove: lì il
  pulsante di Google dell'app non funziona.
- Sul sito compresso gli errori annotati nel registro portano posizioni e
  nomi della copia compressa, non del sorgente.
- Le copie del sorgente fatte quando era pubblico non si richiamano.

## Strumenti

| Strumento | A cosa serve nel passaggio |
|---|---|
| `controlli-rilascio.js` | Controlli statici prima di ogni rilascio, compresi indirizzi degli ambienti e file pubblicati |
| `banco.js`, `banco-sw.js` | Prove dentro l'app e prove del service worker, anche sulla copia compressa |
| `prepara-rinvio.js` | Genera i file della pagina di rinvio |
| `scambio-repository.js stato\|prepara\|scambia\|annulla` | Lo scambio e il suo ritorno |
| `pubblicazione.js [avvia]` | Segue la pubblicazione su Cloudflare e controlla l'avviso |
| `verifica-chiusura.js <ambiente> [--archivi]` | Le prove da estraneo: 28, più 4 sugli archivi |

## Da preparare negli strumenti prima del giorno

Oggi `gh-api.js`, `scambio-repository.js` e `pubblicazione.js` lavorano solo
sull'account di collaudo, per costruzione. Per la produzione vanno estesi, e
provati di nuovo nel collaudo, così:

- l'ambiente si sceglie con un argomento esplicito, e la produzione chiede
  una conferma scritta per ogni comando che cambia qualcosa;
- proprietario, remoto della cartella e database si ricavano dall'ambiente:
  `medicinadurgenzaucsc-maker`, `origin`, sola lettura sul database;
- l'accesso «a codice» a GitHub vale anche per l'account del reparto.
