# I reparti

Un file per reparto: `reparti/<codice>.json`. Il codice dell'applicazione è uno
solo per tutti; di un reparto cambia soltanto ciò che sta scritto qui, e da qui
nasce il piccolo file servito `js/ambiente.js` che dice all'applicazione chi è,
da quali indirizzi può essere aperta e a quale database si collega.

Lo strumento è `collaudo/strumenti/reparti.js`:

```
node collaudo/strumenti/reparti.js                  elenca i reparti e li convalida
node collaudo/strumenti/reparti.js mostra <codice>  stampa il file js/ambiente.js di quel reparto
node collaudo/strumenti/reparti.js controlla        docs/js/ambiente.js è quello generato?
node collaudo/strumenti/reparti.js scrivi           riscrive docs/js/ambiente.js
```

## Regole

- **Mai un segreto.** In questi file stanno solo dati pubblici per natura:
  indirizzi del sito, indirizzo e chiave pubblica del database, identificativo
  del client Google. Sono gli stessi dati che riceve chiunque apra il sito. La
  convalida rifiuta una chiave che non sia quella pubblica, una chiave di un
  altro progetto e ogni campo che non conosce.
- **Ogni reparto ha i suoi indirizzi, il suo database, il suo progetto.** Due
  reparti non possono avere in comune un nome, un indirizzo, un database o un
  progetto Cloudflare: la convalida lo rifiuta.
- **Un reparto di produzione si apre solo da indirizzi scritti per esteso.**
  Niente voci col punto (le anteprime) e niente `localhost`: sui pazienti veri
  non si lavora da un'anteprima né da un PC di sviluppo.
- **`docs/js/ambiente.js` non si modifica a mano.** `docs/` è il sito del
  reparto che pubblica con GitHub Pages dalla cartella così com'è (oggi la
  11N), quindi nel repository quel file è sempre il suo, generato con
  `reparti.js scrivi`. Per ogni altro reparto il file si genera quando si
  prepara il sito (`prepara-sito.js --reparto <codice>`); i banchi di prova lo
  generano in memoria per il collaudo.
- **I nomi `collaudo` e `produzione` sono riservati.** Gli strumenti li usano da
  sempre e alla parola «produzione» legano le loro barriere (sola lettura,
  conferma esplicita). Un reparto li può portare, come codice o come sinonimo,
  solo se è di quel tipo. `produzione` è un sinonimo di `gemelli-11n`: non va
  tolto né spostato su un altro file.

## I campi

| Campo | Che cosa è |
|---|---|
| `codice` | il nome del file senza estensione: minuscole, cifre e trattini |
| `sinonimi` | altri nomi con cui gli strumenti riconoscono il reparto (facoltativo) |
| `nome`, `ospedale` | per le persone: compaiono negli strumenti, `nome` anche nel file servito |
| `tipo` | `produzione` oppure `collaudo` |
| `ambiente` | il valore che la riga `AMBIENTE` del suo database deve avere; oggi coincide col tipo |
| `hosting` | chi ospita il database: `reparto`, `fornitore`, `condiviso`, `locale` (per ora solo letto) |
| `accesso` | come si entra: `google` |
| `lingua` | `it` |
| `canale` | `pilota` oppure `stabile`: chi riceve per primo una versione (per ora solo letto) |
| `versioneLocale` | deve valere `0`: la versione locale arriva con la fase 5 del piano |
| `livello` | dove sta il sito: `0` GitHub Pages, `1` Cloudflare senza cancello, `2` Cloudflare con cancello |
| `host` | i nomi host da cui il sito può essere aperto, senza `https://` e senza percorso; una voce col punto davanti vale per i sottodomini (solo nel collaudo) |
| `pubblicazione` | dove si pubblica, vedi sotto |
| `supabaseUrl` | `https://<progetto>.supabase.co`, senza barra finale |
| `supabaseAnonKey` | la chiave PUBBLICA di quel progetto |
| `googleClientId` | l'identificativo del client Google dell'accesso |

`pubblicazione` ha una o più di queste voci:

- `cloudflare: { progetto }`: il progetto Pages. `<progetto>.pages.dev` deve
  essere fra gli `host`. Il flusso di pubblicazione riconosce il reparto da
  qui: pubblica per il reparto che dichiara il progetto scritto nella variabile
  `CLOUDFLARE_PROGETTO` del repository.
- `githubPages: { repository, cartella: "docs" }`: vuol dire «la cartella
  `docs/` di questo repository, servita così com'è». Può averla un reparto
  solo. Il repository pubblico col solo sito generato è la voce qui sotto:
  non va descritto con questa.
- `sitoPubblico: { repository }`: il «livello 0». Un repository PUBBLICO che
  contiene soltanto i file che il sito serve, già compressi, servito da GitHub
  Pages; il sorgente resta in un repository privato. Lo pubblica
  `collaudo/strumenti/sito-pubblico.js`. Il sito risponde a
  `<proprietario>.github.io/<nome>/`, e `<proprietario>.github.io` deve essere
  fra gli `host`. Poiché due reparti non possono avere un host in comune (e su
  uno stesso indirizzo condividerebbero anche la sessione nel browser), ogni
  reparto al livello 0 vuole un account GitHub suo. Oggi lo ha solo il
  collaudo, per prova.

Nel file servito finiscono solo `codice`, `nome`, `ambiente`, `lingua`,
`versioneLocale`, `host`, `supabaseUrl`, `supabaseAnonKey`, `googleClientId`.

## Aggiungere un reparto

1. Scrivere `reparti/<codice>.json` e lanciare `node collaudo/strumenti/reparti.js`.
2. `node collaudo/strumenti/prepara-sito.js _sito/<codice> compressa --reparto <codice>`
   prepara il suo sito e controlla che non porti dati di un altro reparto.
3. `node collaudo/strumenti/controlli-rilascio.js` costruisce il sito di ogni
   reparto e pretende che differiscano solo in `js/ambiente.js`.

## La tabella di riserva, tolta nella 193

Le versioni 190, 191 e 192 tenevano in `docs/js/api.js` la tabella degli
ambienti della 189, per i minuti di un aggiornamento in cui il browser accoppia
una pagina vecchia al codice nuovo. La 193 l'ha tolta: da allora chi è un sito
lo dice solo il suo file, senza il file l'app si ferma e offre «Ricarica», e il
controllo «nessun dato di altri reparti nel sito» guarda `api.js` per intero.
Gli indirizzi e il database di un reparto si cambiano nel suo file e basta.

## Riservatezza

Oggi il repository della 11N è pubblico e questa cartella contiene solo dati
già pubblici. Il giorno in cui ci entrano i file di altri reparti (ospedale,
indirizzi, database) deve stare in un repository privato.
