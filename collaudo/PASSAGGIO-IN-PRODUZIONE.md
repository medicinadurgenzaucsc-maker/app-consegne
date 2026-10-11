# Passaggio del sito: schema per la produzione

Come si porta il reparto dal sito pubblico su GitHub al sito su Cloudflare con
l'accesso, e il sorgente in un repository privato. È la sequenza provata nel
collaudo l'08/10/2026, riscritta per rifarla in produzione, in tre tempi:
**prima si prepara tutto, senza che i PC del reparto se ne accorgano; poi
ogni dispositivo che usa l'app viene portato sul nuovo indirizzo, uno alla
volta, mentre i due siti convivono; solo alla fine lo scambio dei
repository, che dura pochi minuti.**

Stato: nel collaudo è tutto fatto e provato, compreso il ritorno indietro. In
produzione l'08/10/2026 è stata fatta la preparazione che non tocca il sito in
uso, vedi «A che punto è la produzione» in fondo. Il reparto lavora ancora sul
sito di sempre.

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

Si fa nei giorni prima. Fino alla fase A10 il reparto continua a usare il
sito di sempre; è in A10, prima della fase B, che i dispositivi cambiano
indirizzo.

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

1. Nel collaudo: aggiungere l'indirizzo `<nome>.pages.dev` al campo `host` del
   file del reparto in `reparti/` e rigenerare `docs/js/ambiente.js`
   (`node collaudo/strumenti/reparti.js scrivi`). (Fino alla 189 l'elenco stava
   in `_AMBIENTI.produzione.host`, in `docs/js/api.js`: dalla 190 quella
   tabella è solo una riserva e non si tocca.) Senza, il flusso si
   rifiuta di pubblicare e l'app sul sito nuovo si ferma con «Indirizzo non
   riconosciuto». In produzione **niente voce col punto**: le anteprime di
   Cloudflare non devono poter lavorare sui pazienti veri. Versione nuova,
   controlli di rilascio, prove nel banco.
2. Con l'ok: riga `AMBIENTE = produzione` nella tabella `impostazioni` del
   database di produzione. Oggi manca: l'app funziona lo stesso, ma è quella
   riga a fermarla se un sito parla col database sbagliato.
3. Con l'ok: rilascio in produzione sul sito di oggi. Porta il service
   worker che regge gli errori e l'app che riconosce il nuovo indirizzo. I PC
   lo ricevono con «Update» o col ricaricamento delle 04:00. Lo stesso push
   fa partire la prima pubblicazione su Cloudflare.
   Attenzione: questo rilascio porta nel repository ancora pubblico anche
   strumenti e documentazione del passaggio, e lì restano fino allo scambio.
   Lo scambio si fa quando la fase A10 è conclusa, quanti giorni servano:
   la permanenza di quei file nel repository pubblico per quei giorni è
   una scelta da riportare a Stefano. Non contengono chiavi.
4. Prima del push si prepara il ramo di ritorno, vedi «Ritorno indietro».

### A4. Chiave di pubblicazione, a cura di Stefano

1. Su Cloudflare, nei token API del profilo: creare un token partendo da zero
   con il solo permesso «Cloudflare Pages: Edit», che sta nel gruppo delle
   piattaforme per sviluppatori. Le voci si rileggono dallo schermo.
2. L'identificativo dell'account è nell'indirizzo della console:
   `dash.cloudflare.com/<identificativo>/…`.
3. Nel repository su GitHub, nei segreti dei flussi: `CLOUDFLARE_API_TOKEN` e
   `CLOUDFLARE_ACCOUNT_ID`. Nelle variabili: `CLOUDFLARE_PROGETTO` con il nome
   scelto. `CLOUDFLARE_AVVISA` resta spenta fino alla fase B.

### A5. Il progetto su Cloudflare, a cura di Claude

La strada migliore, usata in produzione: creare il progetto **vuoto** dalla
console e configurare l'accesso prima della prima pubblicazione, così il sito
non resta mai aperto.

1. «Workers & Pages», «Create application»: la procedura proposta crea un
   Worker, non un progetto Pages. Serve il collegamento in basso «Continue to
   Pages», poi «Drag and drop your files», il nome del progetto, «Create
   project». Non occorre caricare nulla: il progetto esiste già e il suo
   indirizzo risponde 522 finché non arriva la prima pubblicazione.
2. Un progetto creato dalla console ha come ramo di produzione `main`: va
   cambiato in `master` da «Settings», «General», «Production branch»,
   «Rename». Senza, il flusso si rifiuta di pubblicare.
3. In alternativa il progetto nasce da solo al primo avvio del flusso, già col
   ramo giusto: ma fino alla configurazione dell'accesso il sito è aperto.

### A6. L'accesso, nel browser di Stefano

Percorsi della console, dove `<id>` è l'identificativo dell'account:
`dash.cloudflare.com/<id>/one/access-controls/apps`, `…/policies`,
`…/settings`, `…/one/integrations/identity-providers`,
`…/one/reusable-components/custom-pages`.

1. Attivare Zero Trust, piano gratuito: chiede una carta e l'accettazione
   delle condizioni, quindi lo fa Stefano. Il nome della squadra viene
   assegnato a caso: va cambiato subito da «Settings», «Team domain», «Edit»,
   prima di creare il client Google, perché l'indirizzo
   `https://<squadra>.cloudflareaccess.com` finisce dentro quel client.
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
2. `scambio-repository.js <ambiente> prepara https://<nome>.pages.dev/`,
   in produzione con `--confermo-produzione`: crea il repository pubblico
   provvisorio `app-consegne-rinvio`, 7 file, un commit senza storia, sito
   già in linea, chiuso alle scritture. Lo strumento accetta solo
   l'indirizzo del sito nuovo di quell'ambiente: una pagina di rinvio della
   produzione non può portare al collaudo.

**I comandi che cambiano qualcosa vogliono l'ambiente scritto**: `collaudo`
oppure `produzione`, come primo argomento. Senza, lo strumento si ferma. Un
comando copiato da un documento non può più lavorare sul collaudo al posto
della produzione, o il contrario.

## Fase A10. Portare i PC sul nuovo indirizzo: PRIMA dello scambio

**Lo scambio non è invisibile da solo.** La sessione dell'app è legata
all'indirizzo: sul sito nuovo ogni PC deve fare un accesso. Dopo lo scambio
il vecchio indirizzo rimanda al nuovo a ogni ricaricamento, compreso quello
che ogni PC fa da solo fra le 04:00 e le 05:00, e compreso «Update»: un PC che
non è mai entrato sul sito nuovo si ritrova fermo sulla schermata di accesso,
di notte, finché qualcuno non clicca. Inoltre da una pagina rimasta aperta
sul vecchio indirizzo la stampa non funziona più e «Numeri Telefono», se non
era già aperto, può non aprirsi.

**Chi usa l'app non sono solo i PC del reparto.** Misura del 07/10/2026:
118 sessioni vive, 35 usate negli ultimi 7 giorni, 16 delle quali su
telefoni, tablet o Mac, da 19 reti diverse. Dopo lo scambio ogni browser
non preparato trova prima il cancello di Cloudflare, dove entrano solo gli
indirizzi del criterio, e poi la schermata di accesso dell'app. Quindi
prima di cominciare si fa l'inventario dal registro, in sola lettura:
identificativo, nome e ultima data dei dispositivi che hanno scritto righe
`supa-auth`. Per ognuno Stefano decide: lo si passa, lo si lascia fuori, o
si aggiunge un indirizzo al criterio di accesso. Sessione e identificativo
stanno nella memoria del browser, quindi **ogni browser e ogni profilo
Windows di ogni PC è un caso a sé**.

Mentre i due siti convivono, su **ogni** dispositivo che deve continuare a
usare l'app:

1. aprire il nuovo indirizzo; passa da Google, senza clic o con un clic per
   scegliere l'account. Se nel browser l'account Google del reparto non è
   collegato servono password e secondo passaggio: va scoperto adesso, non
   alle 04:00;
2. premere «Accedi con Google» nell'app;
3. fare una stampa e un'importazione da TrakCare: alla prima il browser
   chiede il permesso per gli appunti;
4. mettere il preferito e l'icona del nuovo indirizzo e chiudere la scheda
   del vecchio. Il preferito vecchio **non si elimina fino allo scambio**:
   lo si rinomina, perché è la strada del ritorno se la versione nuova desse
   problemi.

Se al cancello si sceglie per sbaglio un account Google che non è nel
criterio, Cloudflare rifiuta e ricorda la scelta: si apre
`https://<nome>.pages.dev/cdn-cgi/access/logout` e si rientra con la casella
del reparto. Da provare una volta nel collaudo.

«Numeri Telefono» sul sito nuovo riparte senza preferiti né recenti: la
rubrica li tiene nella memoria del browser, separata per sito. L'08/10/2026
Stefano ha deciso che la rubrica entrerà nell'app: fino ad allora la
perdita è accettata. Come ci entra, e in che ordine, è scritto più sotto
in «La rubrica dentro l'app».

Il primo PC del reparto vale anche come prova dalla **rete dell'ospedale**:
se quella rete non lasciasse passare `pages.dev` o `cloudflareaccess.com`, lo
scambio lascerebbe il reparto senza applicazione. Finché non è provato da
lì, lo scambio non si fa.

**Quando si può dire che tutti sono passati.** Ogni avvio scrive nel
registro una riga `supa-auth` con l'identificativo del dispositivo, che
nasce nella memoria del browser e quindi è diverso per indirizzo: sul sito
nuovo ogni browser ne riceve uno nuovo, mentre un browser rimasto sul
vecchio indirizzo continua a scrivere con quello di prima, anche al
ricaricamento delle 04:00. Criterio: lo scambio si fa quando nessun
identificativo già presente prima del rilascio, fra quelli dei dispositivi
che contano, ha scritto righe `supa-auth` nelle ultime 48 ore. È una
lettura con `sb.js`, solo identificativi e date. Il registro tiene 20
giorni: l'elenco degli identificativi «di prima» si salva la sera del
rilascio.

**Rilasci nei giorni di convivenza.** `CLOUDFLARE_AVVISA` resta spenta
fino allo scambio, quindi «Update» lo dà `notify-deploy` quando ha finito
GitHub Pages, mentre Cloudflare finisce per conto suo: un PC del sito nuovo
che ricarica prima che Cloudflare abbia finito resta sulla versione
vecchia senza altri avvisi fino alle 04:00. Regola: fra il primo
dispositivo passato e lo scambio non si rilascia nulla, salvo urgenze. In
caso di urgenza si aspetta che `pubblicazione.js produzione` dica
«pubblicato» e che il vecchio sito serva la versione nuova, poi i
dispositivi si fanno ricaricare a mano, controllando il numero di versione
nel menu della rotellina.

Da decidere con Stefano prima dello scambio: la sessione di Cloudflare dura
un mese. Alla scadenza il ricaricamento delle 04:00 ripassa da Google: se la
sessione Google di quel browser nel frattempo è caduta, il PC resta sulla
pagina di Google finché non arriva chi ha la password. Le strade sono tre:
tenere il mese, allungare la durata, oppure esentare dall'accesso la rete
dell'ospedale.

## Fase B. Lo scambio: pochi minuti, a PC già passati

1. **Controlli di partenza**: cartella di lavoro pulita e allineata al remoto;
   `scambio-repository.js produzione stato` deve dire: scambio non fatto,
   pagina di rinvio che porta al sito nuovo e repository chiuso alle
   scritture, accessi al sorgente a posto, sito su Cloudflare allineato al
   ramo. Sono gli stessi controlli che «scambia» rifà da sé prima di
   partire; un elenco che GitHub non lascia leggere vale come controllo
   fallito.
   Non si interrogano i file «raw» né lo zip nei minuti prima dello scambio.
2. **`scambio-repository.js produzione scambia --confermo-produzione`**:
   primo cambio di nome, remoto della cartella spostato, secondo cambio di
   nome, sorgente subito privato, poi si aspetta che il vecchio indirizzo
   serva la pagina di rinvio. Mezzo minuto. Se si interrompe a metà si
   rilancia lo stesso comando: riprende dal passo che manca. Se l'ultima
   riga comincia con «ATTENZIONE» qualcosa non è andato fino in fondo, e lo
   strumento esce con errore.
3. **Avviso ai PC**: variabile `CLOUDFLARE_AVVISA = si`, poi
   `pubblicazione.js produzione avvia --confermo-produzione` e controllo
   che `app_version` porti il commit pubblicato. In produzione i due
   segreti di Supabase ci sono già.
4. **Prove da estraneo**: `verifica-chiusura.js produzione` due minuti dopo
   lo scambio: devono passare tutte e 29. Dopo dieci minuti di nuovo con
   `--archivi`: le 4 prove in più dicono se un archivio pubblico di terzi
   tiene una copia del sorgente di quando era pubblico. Quelle copie non si
   richiamano: è un accertamento da riferire a Stefano, non una condizione
   dello scambio, e non giustifica un ritorno indietro.
5. **Azioni ammesse nei flussi** del repository privato: solo quelle scritte
   da GitHub.

Cosa succede a un PC rimasto sul vecchio indirizzo: la pagina aperta continua
a lavorare e a salvare, perché parla solo col database, ma la stampa non
funziona e «Numeri Telefono», se non era già aperto, può non aprirsi; al
primo ricaricamento passa al sito nuovo e deve fare l'accesso.
Per questo la pagina di rinvio è solo una rete di sicurezza per chi è
sfuggito alla fase A10, non il modo di spostare i PC.

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

## La rubrica dentro l'app (versione 189)

«Numeri Telefono» smette di essere un sito a parte: diventa una pagina
dell'app (`docs/rubrica/`), coi dati nel database delle consegne, visibile
solo a chi è collegato. È un rilascio a sé, che non dipende dal passaggio del
sito, ma ha un **ordine obbligato**, e in produzione ogni passo vuole l'ok di
Stefano. Stato al 09/10/2026: passi 1–4 fatti (in produzione la notte del
09/10, fra le 01:09 e le 01:11); resta il passo 5.

**Perché l'ordine conta.** Il codice nuovo legge le tabelle `rubrica_` del
database delle consegne. Se arrivasse in produzione prima di loro, l'app
delle consegne funzionerebbe lo stesso, ma su tutti i PC il riquadro direbbe
«La rubrica non è ancora disponibile su questo sito», e dal menu la rubrica
di prima non si aprirebbe più: il reparto resterebbe senza numeri. L'ordine
giusto, tabelle prima e codice poi, è innocuo: tabelle che nessun codice
legge non danno fastidio.

1. **Nel collaudo, prima di tutto**: la sezione `rubrica` delle prove
   automatiche, sul sorgente e sulla copia compressa; le prove del service
   worker; le prove a mano elencate in `README.md`.
2. **Le tabelle in produzione**: `rubrica_categorie`, `rubrica_contatti`,
   `rubrica_versione` con la sua riga `id = 1`, le due funzioni, i trigger,
   RLS accesa e policy scritte in modo esplicito, GRANT espliciti, nulla ad
   `anon`. Il testo della migrazione **non è nel repository**: va ripreso
   dalla storia delle migrazioni del collaudo, così com'è stato applicato lì.
   Prima di proseguire:
   - `clona-struttura.js confronta` non segnala più le tabelle `rubrica_`
     come differenza;
   - da utente del reparto, col blocco `DO` descritto in `CLAUDE.md`
     («Verifiche»): la riga di `rubrica_versione` si legge, contatti e
     categorie si leggono e si scrivono, le due funzioni si possono chiamare.
     Una riga di versione che manca, o un permesso dimenticato, farebbero
     leggere a tutti i PC «Il database non mostra la rubrica a questo
     account»;
   - da estraneo, con la sola chiave pubblica: non si legge e non si scrive
     nulla;
   - il numero di `rubrica_versione` deve crescere a **ogni** modifica, anche
     a più modifiche nello stesso secondo: nel collaudo è così
     dall'08/10/2026 (migrazione `rubrica_versione_sempre_crescente`: il
     trigger scrive `ts = greatest(ts + 1, secondi dell'orologio)`), e **la
     stessa funzione del trigger va creata così anche in produzione**. Lo
     controlla la prova 14 della sezione `rubrica`: cinque modifiche di fila,
     salita di almeno cinque.
3. **I contatti in produzione**: copia dal vecchio database della rubrica,
   con gli stessi identificativi. `copia-rubrica.js produzione prima-copia --confermo-produzione`
   (dal 09/10/2026): la stessa istruzione unica e lo stesso confronto delle
   impronte, ma in produzione non svuota mai e si rifiuta se la destinazione
   ha già righe. È una
   scrittura nel database del reparto: si decide e si prova prima con
   Stefano. Da quel momento ciò che viene scritto nella rubrica di prima
   **non arriva** in quella nuova: la copia si fa a ridosso del rilascio, e
   `confronta` dice se le due copie coincidono ancora. Chi usa la rubrica di
   prima dal telefono va avvisato che da quel giorno non si aggiorna più lì.
4. **Il codice**: rilascio della versione 189 col percorso di sempre
   (collaudo, controlli, ok, `master`). Subito dopo, dal browser di Stefano:
   «Numeri Telefono» mostra i contatti, la console resta pulita, e un giro
   crea / modifica / elimina su un contatto di prova, poi tolto.
5. **Solo alla fine, a rubrica nuova provata**: spegnere il vecchio sito
   della rubrica e chiudere alle scritture anonime il suo database. È
   l'ultimo passo dello stesso rilascio, non una pulizia facoltativa. Più
   avanti, quando non servono nemmeno per un ritorno, si tolgono il suo
   repository e il suo database. Dalla sera dell'08/10/2026 quel database ha
   già tre vincoli che rifiutano `<`, `>` e, nei numeri, le virgolette.

**Ritorno.** Finché il vecchio sito è acceso, tornare indietro è ripubblicare
la versione di prima: il menu riapre la rubrica di prima. Dopo il passo 5 non
c'è più dove tornare, ed è per questo che viene per ultimo.

**Preferiti.** Dove la rubrica di prima e l'app stavano sullo stesso sito, la
rubrica nuova riprende una volta sola i preferiti e i chiamati di recente (gli
identificativi dei contatti sono gli stessi); il tema riparte chiaro. Su un
indirizzo nuovo ripartono da zero, come scritto nella fase A10.

**Che cosa cambia per chi la usa**, oltre al fatto che serve essere collegati
all'app: un numero ha al massimo 20 cifre e una nota 120 caratteri (una nota
più lunga già salvata resta com'è finché non la si tocca); l'esportazione CSV
usa il «;» come separatore, così Excel la apre già in colonne; la categoria
va scelta, e un contatto che non ne ha non ne riceve una di nascosto; il tema
predefinito è chiaro.

## Ritorno indietro

| Se va storto | Rimedio |
|---|---|
| La versione nuova dà problemi sul sito di oggi, prima dello scambio | Ramo `ritorno-v181`, preparato prima del rilascio: sopra il commit rilasciato, le pagine della 181 col numero 188. `git checkout master`, `git merge --ff-only refs/heads/ritorno-v181`, `git push origin master`. GitHub Pages torna alla 181 e `notify-deploy` avvisa i PC. **Vale solo per chi lavora dal vecchio indirizzo**: il flusso di Cloudflare si ferma al passo «File del sito» (a quelle pagine manca `404.html`: provato) e il sito nuovo resta alla 187. I dispositivi già passati vanno riportati sul vecchio indirizzo, dove la sessione c'è ancora, finché una versione corretta non è pubblicata su entrambi i siti |
| Dopo un ritorno | Il ramo `collaudo` non discende più da `master`: va riallineato prima del rilascio successivo, che porterà il numero 189. `scambio-repository.js produzione stato` dirà «NON PRONTO» finché non c'è una nuova pubblicazione su Cloudflare |
| Il sito nuovo non funziona, prima dello scambio | Niente da fare: i PC usano ancora quello vecchio |
| Google non fa più entrare | Nell'applicazione del sito, «Metodi di login», aggiungere «One-time PIN»: torna il codice via mail |
| Lo scambio si interrompe a metà | Rilanciare `scambio-repository.js produzione scambia --confermo-produzione`, che riprende dal passo mancante, oppure lo stesso comando con `annulla`. `scambio-repository.js produzione stato` dice a che punto si è |
| Lo scambio lascia il vecchio indirizzo in errore | `scambio-repository.js produzione annulla --confermo-produzione`: in 39 secondi torna tutto com'era, sorgente di nuovo pubblico. Si può rilanciare: se i nomi sono già tornati ma il sito no, lo riaccende |
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
- GitHub Pages pubblica al massimo una decina di volte l'ora per sito. Dopo
  cinque scambi e annullamenti di fila, l'08/10/2026 nel collaudo, le
  pubblicazioni sono andate in errore e il vecchio indirizzo è rimasto a 404
  finché l'ora non è passata. In produzione lo scambio è uno solo: ma
  non si fanno prove ripetute sul repository del reparto.
- Un comando scritto senza ambiente vale collaudo solo se legge. `scambia`,
  `annulla`, `prepara` e `avvia` senza ambiente si fermano.
- Il repository della pagina di rinvio è chiuso da **due** regole, una per i
  rami e una per le etichette: con la sola regola dei rami un `git push
  --tags` vi porterebbe tutta la storia del sorgente. `stato` dice «chiuso
  solo in parte» se ne manca una; `scambia` e `prepara` mettono quella che
  manca. Provato nel collaudo: né un ramo né un'etichetta si lasciano creare.
- Subito dopo una rinomina il vecchio indirizzo serve ancora per mezzo minuto
  la copia di prima: per questo `annulla` conta buona l'applicazione solo a
  pubblicazione conclusa, e `scambia` solo la pagina di rinvio che porta al
  sito nuovo.
- `scambia` rilanciato a scambio già fatto controlla comunque il vecchio
  indirizzo e, se serve, richiede di nuovo la pubblicazione a GitHub Pages.
- `pubblicazione.js <ambiente> avvia` ripubblica lo stesso commit: sui PC non compare
  «Update», perché l'avviso scatta solo quando il commit annunciato cambia.
- Il nome `collaudo` è sia un ramo sia un remoto: nei comandi di unione si
  scrive `refs/heads/collaudo`.
- La riga `AMBIENTE` deve valere esattamente `produzione`, in minuscolo: un
  altro valore fa comparire su ogni PC l'avviso di configurazione incoerente.
- Un indirizzo aggiunto a `utenti_autorizzati` va scritto in minuscolo, e dà
  accesso pieno ai dati dei pazienti.
- La pagina di rinvio non toglie dal browser, sul vecchio indirizzo, la
  sessione e le copie di recupero delle schede: restano lì finché qualcuno
  non svuota i dati del sito. Da valutare prima dello scambio in produzione.
- Le copie del sorgente fatte quando era pubblico non si richiamano.

## Strumenti

| Strumento | A cosa serve nel passaggio |
|---|---|
| `controlli-rilascio.js` | Controlli statici prima di ogni rilascio, compresi indirizzi degli ambienti e file pubblicati |
| `banco.js`, `banco-sw.js` | Prove dentro l'app e prove del service worker, anche sulla copia compressa |
| `prepara-rinvio.js` | Genera i file della pagina di rinvio |
| `scambio-repository.js <ambiente> stato\|prepara\|scambia\|annulla` | Lo scambio e il suo ritorno |
| `pubblicazione.js <ambiente> [avvia]` | Segue la pubblicazione su Cloudflare e aspetta l'avviso ai PC, che venga dal flusso o da `notify-deploy` |
| `verifica-chiusura.js <ambiente> [--archivi]` | Le prove da estraneo: 29, più 4 sugli archivi |

## A che punto è la produzione

Fatto l'08/10/2026, senza toccare il sito in uso al reparto:

- account Cloudflare del reparto, progetto Pages `consegne-reparto` vuoto,
  ramo di produzione `master`: l'indirizzo sarà
  `https://consegne-reparto.pages.dev/`;
- Zero Trust attivo, squadra `consegne-reparto`; applicazione
  `consegne-reparto - sito` col criterio `autorizzati-produzione` (casella del
  reparto, gistech, Stefano), sessione di un mese, cookie protetti; anteprime
  chiuse. Da estraneo sito, file e anteprime rimandano già all'accesso;
- ingresso **solo con Google** e autenticazione immediata: provider Google
  aggiunto da Stefano col client «Accesso Cloudflare produzione»; da estraneo
  la pagina di accesso manda dritta a Google, che riconosce client e
  indirizzo di ritorno. Nome mostrato: «Consegne Reparto»;
- nel repository del reparto la variabile `CLOUDFLARE_PROGETTO` e i segreti
  `CLOUDFLARE_ACCOUNT_ID` e `CLOUDFLARE_API_TOKEN` (il token l'ha messo
  Stefano: se è giusto lo dirà la prima pubblicazione);
- repository `app-consegne-rinvio` con la pagina di rinvio, chiuso alle
  scritture nei rami; la regola delle etichette, aggiunta allo strumento la
  sera dell'08/10, la mette `scambia` al momento dello scambio;
- versione 187 provata nel collaudo: riconosce il nuovo indirizzo.

Da fare, nell'ordine:

1. Stefano, quando vuole: «Test» accanto a Google in «Integrations»,
   «Identity providers». È l'unico modo di sapere prima che il segreto del
   client incollato in Cloudflare è quello giusto.
2. La sera, con l'ok a ogni passo:
   - già pronto: il ramo `ritorno-v181`. Verifiche prima del push:
     `git merge-base --is-ancestor <commit da rilasciare> ritorno-v181`
     deve riuscire, e `git diff 1468f36 ritorno-v181 -- docs` deve mostrare
     solo le due righe col numero di versione;
   - prima del rilascio, perché non ne dipendono e così si verificano con
     calma: nel database la riga `AMBIENTE = produzione` in `impostazioni`
     e gistech, in minuscolo, in `utenti_autorizzati`; nel client Google
     dell'app «ConsegneReparto» l'origine
     `https://consegne-reparto.pages.dev`. Le due scritture sono state
     provate nel collaudo e non raggiungono i PC: le due tabelle non sono
     fra quelle in tempo reale e non hanno trigger;
   - si salva l'elenco degli identificativi dei dispositivi che hanno
     scritto nel registro prima del rilascio: serve per la fase A10;
   - rilascio: `git checkout master`, poi
     `git merge --ff-only refs/heads/collaudo` (il nome `collaudo` da solo è
     ambiguo: è anche il nome di un remoto), poi `git push origin master`.
     GitHub Pages serve la 187 e `notify-deploy` avvisa i PC; lo stesso push
     fa partire la pubblicazione su Cloudflare. `pubblicazione.js produzione`
     segue il flusso e poi aspetta che l'avviso di `notify-deploy` arrivi
     nel database. Se il flusso di Cloudflare fallisce, per esempio perché
     la chiave non è giusta, lo strumento lo dice e aspetta comunque
     l'avviso: il sito che il reparto usa non passa da quel flusso;
   - prove dal browser di Stefano: sul **vecchio** indirizzo, che è quello
     che il reparto usa, versione 187 nel menu della rotellina e nessun
     avviso a tutto schermo; poi sul sito nuovo, le prove della fase A8.
3. Nei giorni dopo: inventario dei dispositivi e passaggio di ognuno al
   nuovo indirizzo, fase A10, cominciando da un PC del reparto per provare
   la rete dell'ospedale. In quei giorni nessun rilascio, salvo urgenze.
4. Solo quando tutti i dispositivi che contano sono entrati sul sito nuovo:
   lo scambio, fase B.

## Da preparare negli strumenti prima del giorno

Fatto l'08/10/2026: `gh-api.js`, `scambio-repository.js` e `pubblicazione.js`
accettano l'ambiente come primo argomento.

- `scambio-repository.js produzione stato` e `pubblicazione.js produzione`
  leggono soltanto;
- i comandi che cambiano qualcosa vogliono l'ambiente scritto e, in
  produzione, anche `--confermo-produzione`;
- `scambia` e `annulla` si possono rilanciare da ogni stato a metà, e
  `annulla` riaccende il sito se i nomi sono tornati ma il sito no; i
  controlli di partenza sono stati provati su un caso positivo nel
  collaudo, con un aggancio esterno finto;
- proprietario, remoto della cartella e database si ricavano dall'ambiente:
  `medicinadurgenzaucsc-maker`, `origin`, sola lettura sul database;
- la chiave di GitHub della produzione è quella già scritta nel remoto
  `origin`, che ha i permessi necessari. L'accesso «a codice» anche per
  l'account del reparto resta da fare, insieme al cambio di quella chiave.

## Stato al 09/10/2026

- **Rilascio della 187**: 08/10/2026 alle 20:23, dal vecchio sito.
- **Scambio dei repository**: 09/10/2026 alle 00:35, 27 secondi, nessun
  errore nel mezzo; `verifica-chiusura.js produzione`: 29 prove superate.
  Da allora `origin` è `medicinadurgenzaucsc-maker/app-consegne-sorgente`
  (privato), il sito del reparto è `https://consegne-reparto.pages.dev/`,
  il vecchio indirizzo mostra la pagina di rinvio e l'avviso ai PC lo dà il
  flusso di Cloudflare (`CLOUDFLARE_AVVISA = si`). Lo scambio è stato fatto
  prima che i PC del reparto fossero entrati sul sito nuovo, per decisione
  di Stefano: ogni PC rifà l'accesso al primo ricaricamento.
- **Rubrica dentro l'app**: tabelle e contatti in produzione, versione 189
  pubblicata dal repository privato (prima pubblicazione fatta tutta dal
  flusso di Cloudflare, avviso compreso).
- **Rubrica di prima**: spenta il 09/10/2026 alle 08:50, su ordine di
  Stefano: repository `rubrica-gemelli` privato (sito spento), flusso
  Keepalive spento, database in pausa. Da estraneo sei prove su sei chiuse.
  Nulla è stato cancellato: la cancellazione definitiva la fa Stefano.
- **Notte del passaggio**: quattro PC del reparto hanno ricaricato da soli
  fra le 04:03 e le 04:08 e sono arrivati al sito nuovo dalla rete
  dell'ospedale, passando da soli il cancello di Cloudflare; si sono fermati
  sul pulsante «Accedi con Google» dell'app finché qualcuno non l'ha premuto.
- **Il repository `app-consegne-sorgente` non si cancella**: è privato, e lì
  girano la pubblicazione su Cloudflare, l'avviso ai PC e il Keepalive del
  database del reparto. Anche la pagina di rinvio vive su GitHub.
- **Restano**: il passaggio degli ultimi dispositivi al nuovo indirizzo; le
  voci della fase C; rifare dopo qualche giorno
  `verifica-chiusura.js produzione --archivi`.

## Stato al 10/10/2026

- **Ritorno su GitHub Pages**: il 09/10/2026 alle 13:20 il filtro Fortinet del
  Policlinico ha bloccato `consegne-reparto.pages.dev` (categoria
  «Information Technology»). Alle 13:26, su ordine di Stefano,
  `scambio-repository.js produzione annulla --confermo-produzione`: 40
  secondi. Da allora `origin` è di nuovo
  `medicinadurgenzaucsc-maker/app-consegne`, pubblico; il sito del reparto è
  `https://medicinadurgenzaucsc-maker.github.io/app-consegne/`, servito da
  `docs/`; `CLOUDFLARE_AVVISA = no` e l'avviso ai PC lo dà di nuovo
  `notify-deploy`. Il sito su Cloudflare resta pubblicato, dietro l'accesso,
  ma dall'ospedale non si raggiunge; la pagina di rinvio è parcheggiata in
  `app-consegne-rinvio`. Verificato: i 21 file serviti identici al commit
  `28a7bc3` (v189), mail di prova partita il 10/10 alle 08:21.
- **Lezione**: un filtro ospedaliero può bloccare `pages.dev` e
  `cloudflareaccess.com`. Prima di ogni passaggio si prova la rete da un PC
  del reparto (dalla 190: `rete.html`) e si tiene pronto l'indirizzo di prima.
  Come andare avanti: nel piano, capitolo 3.3 (il piano sta nel ramo `piano`
  del repository privato di collaudo, non in questo repository).

## Rilascio della 190: i reparti (fase 1 del piano)

Per il reparto non cambia nulla di visibile: stesso indirizzo, stesso accesso,
stesse schede. Cambia dove l'app legge chi è: da `js/ambiente.js`, generato da
`reparti/gemelli-11n.json`, invece che dalla tabella in `api.js` (che resta
come riserva). Nessuna modifica al database. In produzione **solo con l'ok di
Stefano, di sera**.

### A che punto è

- **Banco, 10/10/2026**: controlli di rilascio tutti superati; `__prove()` 142
  su 142 nel sorgente e nella copia compressa; `__proveSw()` 56 su 56 nel
  sorgente e nella copia compressa; prova di fumo da scollegato pulita;
  revisione indipendente fatta e correzioni applicate.
- **Sito di collaudo, 10/10/2026 alle 09:46**: pubblicata dal flusso
  (esecuzione n. 9, commit `04b54bb`; l'11/10/2026 la storia non ancora pubblica
  è stata riscritta per tenere privato il piano, e lo stesso contenuto è ora nel
  commit `5a7d141`). Il passo «File del sito» dice «reparto:
  collaudo», «Avvisa i PC aperti» è riuscito (il segreto `SUPABASE_URL` è
  quello giusto), l'impronta di `js/ambiente.js` pubblicato è quella del file
  del collaudo. Dal Chrome di Stefano, già collegato: versione 190,
  `_AMB.origine` uguale a `file`, 28 schede, service worker `consegne-v190`
  con `js/ambiente.js` in cache (la pagina era passata da sola dalla 189 alla
  190), `rete.html` con tutte e cinque le prove riuscite. Da estraneo ogni
  indirizzo, file nuovi compresi, rimanda ancora al cancello.
- **Ritorno, 10/10/2026**: ramo `ritorno-v189` preparato e provato nel banco.
  Gli strumenti della 190 preparano e comprimono il suo sito per tutti e due i
  reparti (il flusso di Cloudflare quindi lo pubblicherebbe). Nel banco del
  service worker, passaggio dalla 190 al ritorno nei due stili: 54 prove su
  56, e le due fallite sono quelle attese (il service worker 191 non tiene in
  cache `js/ambiente.js`). Il ramo va ricostruito sopra il commit che si
  rilascia davvero, se nel frattempo è cambiato.
- **Produzione**: è ancora la 189.

### Sequenza

0. **Prima il sito di collaudo.** Commit sul ramo `collaudo`,
   `git push collaudo collaudo:master`, `node collaudo/strumenti/pubblicazione.js`:
   il passo «File del sito» deve dire «reparto: collaudo» e «Avvisa i PC
   aperti» deve riuscire (da questa versione pretende che il segreto
   `SUPABASE_URL` del repository sia il database del reparto: se fallisse con
   quel messaggio, il segreto va riscritto nella forma
   `https://<progetto>.supabase.co`). Poi dal sito di collaudo: versione 190 nel
   menu della rotellina e, in console, `_AMB.origine` uguale a `file`.
1. Sul commit da rilasciare, `node collaudo/strumenti/controlli-rilascio.js`:
   tutti i controlli superati. Contano soprattutto «docs/js/ambiente.js è
   quello generato per il reparto gemelli-11n» e «riconoscimento degli
   indirizzi»: sono la sola prova possibile del file della 11N prima del
   rilascio, perché su un PC di sviluppo quel file non si serve mai.
2. **Il ritorno si prepara prima del push**, come per le versioni di prima:
   ramo `ritorno-v189`, sopra il commit da rilasciare, con le pagine della 189
   e il numero 191.

   ```
   git checkout -b ritorno-v189 <commit della 190>
   git checkout 28a7bc3 -- docs
   (191 in docs/sw.js e nel menu di docs/index.html, poi il commit)
   ```

   `git checkout 28a7bc3 -- docs` riporta i quattro file cambiati e **lascia al
   loro posto `docs/js/ambiente.js` e `docs/rete.html`**, apposta: le pagine
   della 189 non li chiedono, e con loro gli strumenti e il flusso della 190
   continuano a funzionare, quindi il ritorno si può pubblicare anche nel
   collaudo. Verifiche: `git merge-base --is-ancestor <commit della 190>
   ritorno-v189` riesce; `git diff 28a7bc3 ritorno-v189 -- docs` mostra solo
   le due righe della versione e i due file in più. Su quel commit
   `controlli-rilascio.js` **dà sette KO attesi**, ai punti 3, 5, 8 e 9: le
   pagine della 189 non hanno il tag di `js/ambiente.js`, il loro service
   worker non lo tiene in cache e in `api.js` non c'è il blocco nuovo (a 190
   rilasciata diventano nove: il confronto con `origin/master` vede anche il
   tag sparito dalle due pagine). La verifica di un ritorno è quel confronto. Il passaggio dalla 190 al ritorno si prova nel banco del service
   worker, lanciato dal ramo di ritorno con `BANCO_RIF=<commit della 190>`.
3. **Non chiedere `…/app-consegne/js/ambiente.js` prima che GitHub Pages abbia
   pubblicato**, nemmeno con un parametro in coda: la rete di GitHub non ne
   tiene conto, e il 404 che terrebbe impedirebbe al service worker 190 di
   installarsi.
4. `git merge --ff-only collaudo` su `master`, `git fetch`, `git push origin
   master`.
5. Prima `node collaudo/strumenti/pubblicazione.js produzione`, che aspetta la
   fine della pubblicazione e l'avviso di `notify-deploy`; solo dopo
   `node .claude/stasera/verifica-sito-servito.js <commit> attendi`: ogni file
   servito identico al commit, `js/ambiente.js` e `rete.html` compresi.
6. Da un PC collegato (o dal Chrome di Stefano): rotellina, versione 190; in
   console `_AMB.origine` deve dire `file`.
7. Lo stesso push fa partire «Pubblica su Cloudflare» sul progetto
   `consegne-reparto` (irraggiungibile dall'ospedale, non avvisa i PC): il
   passo «File del sito» deve dire «reparto: gemelli-11n».
8. Nei primi minuti qualche PC può partire dalla tabella di riserva (pagina
   vecchia accanto al codice nuovo): lo si vede nel registro dell'app, senza
   leggere i registri di Supabase:
   `select count(*), to_timestamp(max(ts)/1000) from logs where tipo =
   'ambiente-da-tabella'`. **Nessuna riga è l'esito normale** e non prova che
   ogni PC abbia già ricaricato: dice solo che nessuna pagina è partita dalla
   tabella.

**Se un PC mostra «Indirizzo non riconosciuto» o «Configurazione del sito non
valida»**: avvisare Stefano e aspettare. Se decide il ritorno:
`git checkout master`, `git merge --ff-only refs/heads/ritorno-v189`,
`git push origin master`. GitHub Pages torna alle pagine della 189 e
`notify-deploy` avvisa i PC. Regge in ogni mescolanza di file vecchi e nuovi:
le pagine della 189 non chiedono `js/ambiente.js`, e una pagina della 190
rimasta nel browser lo trova ancora. Dopo un ritorno il ramo `collaudo` non
discende più da `master` e va riallineato; il rilascio successivo porta il
numero 192.

**Dopo**: quando nel registro non compaiono righe `ambiente-da-tabella` da
qualche giorno, una versione a parte toglie la tabella da `api.js` (elenco di
ciò che va adattato in `CLAUDE.md`, «Reparti»).

## Stato all'11/10/2026

- **La 190 è in produzione dalle 09:00** (commit `c953b37`, con l'ok di
  Stefano, «anche adesso»). `pubblicazione.js produzione`: flusso di
  Cloudflare riuscito con «reparto: gemelli-11n», avviso ai PC dato da
  `notify-deploy`. `verifica-sito-servito.js`: i 23 file serviti identici al
  commit. Registro dell'app: il primo PC ripartito alle 09:02, nessun errore e
  nessuna riga `ambiente-da-tabella`. Il ramo `ritorno-v189` (`6c80c11`) resta
  pronto.
- **Il piano è rimasto privato.** Prima del rilascio la storia non ancora
  pubblicata è stata riscritta senza `collaudo/PIANO-REPARTI.md`, che ora vive
  nel ramo `piano` del repository privato di collaudo; la storia di prima è
  nel ramo `archivio-con-piano-2026-10-11` dello stesso repository. È servita
  una sola spinta forzata sul repository privato, autorizzata da Stefano. Per
  questo i commit della 190 hanno cambiato nome (`04b54bb` è diventato
  `5a7d141`). Due guardie: il controllo 15 di `controlli-rilascio.js` e il
  gancio `pre-push` di questa cartella.
- **Il vecchio indirizzo della rubrica risponde di nuovo**, dalle 09:22:
  repository pubblico `medicinadurgenzaucsc-maker/rubrica-gemelli` con la sola
  pagina di rinvio verso `…/app-consegne/?rubrica=1`, chiuso alle scritture
  (`rinvio-rubrica.js stato`). Finché in produzione c'è la 190, che non
  conosce il segno, chi arriva di lì vede l'app intera e apre la rubrica dal
  menu; dalla 192 vede solo la rubrica.

## Rilascio della 192: il modo rubrica

Che cosa cambia: la pagina aperta con `?rubrica=1` mostra, dopo l'accesso,
solo la rubrica a tutto schermo, con l'icona per uscire dall'account e una sua
icona installabile; aperta senza il segno l'app resta com'è. Nessuna modifica
al database. Regole in `CLAUDE.md`, «Modo rubrica».

1. Nel collaudo: `controlli-rilascio.js` (punti 11 bis e 16), `__prove()` nel
   sorgente e nella copia compressa, `__proveSw()` nelle due forme, prova di
   fumo da scollegato con e senza il segno.
2. Sito di collaudo: `git push collaudo collaudo:master`,
   `pubblicazione.js`; poi dal sito, da un telefono: `…/?rubrica=1`, accesso,
   solo la rubrica, icona di uscita, «Aggiungi a schermata Home». Su un
   telefono Android che ha già l'icona delle consegne va guardato se Chrome
   offre lo stesso di installare quella dei numeri: le due app hanno lo stesso
   perimetro.
3. Prima del push in produzione il ramo di ritorno `ritorno-v190`: le pagine
   della 190 col numero 193, lasciando al loro posto i file nuovi di `docs/`.
4. Produzione, solo con l'ok di Stefano: `git merge --ff-only collaudo` su
   `master`, push, `pubblicazione.js produzione`, poi
   `verifica-sito-servito.js <commit> attendi`. I file nuovi (manifesto e
   icone) non si chiedono al sito prima che GitHub Pages abbia finito.
5. Dopo: `rinvio-rubrica.js stato` deve dire che l'app in linea riconosce il
   segno; dal telefono di un medico la vecchia icona porta alla rubrica.
6. Su un PC del reparto può essere rimasta una scheda sull'indirizzo col segno
   dai giorni della 190 (ci si arrivava dal vecchio indirizzo della rubrica,
   anche dalla riga rimasta in «Link utili»). Con «Update» torna da sola
   all'app intera: il ricaricamento lo fa ancora il codice della 190, che il
   segno non lo conserva. Con F5 o col ricaricamento delle 04:00, che tengono
   l'indirizzo, mostra invece solo i numeri: allora va chiusa, e l'app riaperta
   dal collegamento solito. L'icona di uscita, lì, scollega anche le consegne
   di quel PC: la domanda lo dice.

### Com'è andata (11/10/2026)

- Revisione indipendente della 192, poi controllo indipendente delle sole
  correzioni: nessun difetto nel codice servito; ritoccati una prova e due
  documenti.
- Collaudo: controlli statici, 145 prove nell'app (sorgente e copia
  compressa), 62 nel service worker (nelle due forme), prova di fumo da
  scollegato con e senza il segno. Sito di collaudo pubblicato (`8918f49`) e
  guardato nel browser: col segno solo la rubrica, nessuna richiesta verso i
  dati dei pazienti, icona di uscita al suo posto; senza il segno l'app di
  sempre.
- Ramo di ritorno `ritorno-v190` (`b18eeb0`): le pagine della 190 col numero
  193, con manifesto e icone della rubrica lasciati al loro posto. Provato nel
  banco del service worker con `BANCO_RIF=8918f49`: 56 prove su 56 (le 6 sulla
  pagina aperta per la rubrica lì non si applicano).
- Produzione: push alle 10:36:53 (`96de162`), avviso ai PC da `notify-deploy`
  alle 10:37:38, `verifica-sito-servito.js`: i 27 file serviti identici al
  commit. `rinvio-rubrica.js stato`: l'app in linea riconosce il segno, quindi
  la vecchia icona dei telefoni apre solo la rubrica.

## Livello 0: il sito pubblico di prova (11/10/2026)

Che cos'è: il sito di un reparto servito da GitHub Pages da un repository
**pubblico che contiene soltanto i file del sito, già compressi**; il sorgente
(pagine leggibili, strumenti, documenti, storia) resta in un repository
privato. GitHub Pages non si accende su un repository privato di un account
gratuito: per questo il sito sta in un repository a parte.

Nel collaudo, con l'ok di Stefano:

- reparto `collaudo`, voce `pubblicazione.sitoPubblico`: repository
  `gistech2026/collaudo-sito`, sito `https://gistech2026.github.io/collaudo-sito/`;
- `node collaudo/strumenti/sito-pubblico.js collaudo pubblica`: repository
  creato, 28 file (i 27 della copia compressa senza `_headers`, più
  `.nojekyll`), sito pronto dopo 29 secondi con gli stessi file, impronta per
  impronta;
- da estraneo: nel repository ci sono solo quei file, nessuno del sorgente;
  la storia ha due commit (quello iniziale di GitHub e il sito); un indirizzo
  che non esiste riceve la pagina 404 del sito;
- nel browser: ambiente riconosciuto (l'host era già fra quelli del collaudo),
  service worker attivo con ambito `/collaudo-sito/`, 26 file in cache; da
  collegato le schede del collaudo e il canale in tempo reale; col segno
  `?rubrica=1` solo la rubrica, cornice in `/collaudo-sito/rubrica/`.

Che cosa resta diverso dal sito su Cloudflare: nessun cancello davanti alle
pagine (chiunque riceve il codice compresso, come oggi dal sito del reparto);
il sito non si aggiorna da solo, va rilanciato `sito-pubblico.js` dopo ogni
pubblicazione.

Dopo la prima pubblicazione una revisione indipendente ha fatto aggiungere
allo strumento quattro protezioni, perché ciò che arriva in un repository
pubblico non si ritira: pubblica solo ciò che è registrato in git; prima di
scrivere controlla che il repository sia appena nato o sia già il sito di
quel reparto (la pubblicazione ne sostituisce tutto il contenuto); lo crea
solo con `--crea`, e solo se la chiave è dell'account che lo possiede; firma
i commit con un autore neutro. I primi due commit di `collaudo-sito`, fatti
prima, portano come autore l'indirizzo di posta dell'account di collaudo:
restano nella storia, e si tolgono solo rifacendo il repository da capo.

Per portarci la 11N servono, ognuno col suo ok:

1. Lo scambio dei repository con il sito al posto della pagina di rinvio: il
   repository col sorgente diventa privato sotto un altro nome, e al nome di
   sempre (`app-consegne`) va un repository pubblico col solo sito compresso.
   L'indirizzo dei PC non cambia, nessuno rifà l'accesso, e `github.io` la
   rete dell'ospedale lo lascia già passare. `scambio-repository.js` oggi
   mette al vecchio nome la pagina di rinvio: va adattato e provato nel
   collaudo.
2. La pubblicazione dal flusso del repository privato: prepara la copia
   compressa, la pubblica nel repository pubblico, aspetta che GitHub Pages
   serva i file (`attendiServito`) e solo allora avvisa i PC. Oggi l'avviso lo
   dà `notify-deploy`, che vive nel repository pubblicato da GitHub Pages.
3. Una chiave che scriva solo in quel repository pubblico, messa da Stefano
   nei segreti del repository privato.
4. `docs/js/ambiente.js` della 11N non sarà più «il file servito così
   com'è»: lo genererà la preparazione del sito, come per ogni altro reparto.
5. Le copie del sorgente scaricate quando era pubblico non si richiamano.
