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
3. Si prova lì: a mano, e con le batterie in `collaudo/strumenti/`.
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
| `generatore-pazienti.js` | Si carica nella pagina dell'app (in locale) e riempie le schede: `await window.__generaPazienti()`. |
| `funzioni-collaudo.js pubblica\|configura` | Pubblica `google-token` e `google-finto` nel collaudo e ne imposta le variabili. |
| `prova-funzione-mail.js` | 25 prove della funzione mail contro il finto Google. |
| `supabase-accesso.js`, `github-accesso.js` | Autorizzazione «a codice» dei due account di collaudo, senza far passare chiavi dalla chat. |

## Provare in locale da utente collegato

Il login Google lo fa solo una persona. Per le prove automatiche si usa un
utente fittizio, `collaudo-automatico@example.com`, presente nella lista
autorizzati del solo collaudo: `gettone-prova.js` gli firma un token a scadenza
con la chiave del progetto di collaudo. Lo stesso token in produzione è
rifiutato.

## Regole

- Mai copiare nel collaudo righe di pazienti veri, nemmeno anonimizzate: il
  testo libero identifica comunque.
- Dagli strumenti la produzione si legge soltanto.
- Il repository di collaudo non ha segreti: i due workflow sono attivi solo nel
  repository di produzione.
- Una modifica al database che crea tabelle scrive sempre la RLS in modo
  esplicito: nel collaudo verrebbe accesa comunque in automatico, in produzione no.
