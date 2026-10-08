// Prepara i file della «pagina di rinvio»: ciò che resta al VECCHIO indirizzo
// dell'applicazione quando il sito si sposta altrove. Chi apre il vecchio
// indirizzo (anche la stampa, anche una pagina che non esiste più) viene portato
// al nuovo con gli stessi parametri; il service worker lasciato nei browser dalla
// versione precedente si toglie da solo. I modelli stanno in rinvio/.
//
//   node collaudo/strumenti/prepara-rinvio.js <nuovo indirizzo>             elenca i file
//   node collaudo/strumenti/prepara-rinvio.js <nuovo indirizzo> <cartella>  li scrive nella cartella
//
// Il nuovo indirizzo è la radice del sito nuovo, https e con la barra finale
// (es. https://consegne-collaudo.pages.dev/). Usato da scambio-repository.js.
const fs = require('fs');
const path = require('path');

const RADICE = path.resolve(__dirname, '../..');
const MODELLI = path.join(RADICE, 'rinvio');
const SEGNO = '{{DESTINAZIONE}}';

function destinazione(nuovo) {
  let u;
  try { u = new URL(String(nuovo)); } catch (e) { throw new Error('indirizzo non valido: ' + nuovo); }
  if (u.protocol !== 'https:') throw new Error('il nuovo indirizzo deve cominciare con https://');
  if (u.search || u.hash || u.username || u.password || u.port) throw new Error('il nuovo indirizzo non deve avere parametri, ancora, credenziali o porta');
  if (!/^[a-z0-9.-]+$/.test(u.hostname) || u.hostname.indexOf('.') < 0) throw new Error('nome host non valido: ' + u.hostname);
  if (!/^[/]([a-z0-9_-]+[/])*$/.test(u.pathname)) throw new Error('il percorso deve finire con la barra e contenere solo lettere minuscole, cifre, trattini: ' + u.pathname);
  return u.href;
}

// [percorso nel sito, contenuto] per ogni file della pagina di rinvio
function fileRinvio(nuovo) {
  const base = destinazione(nuovo);
  const modello = fs.readFileSync(path.join(MODELLI, 'pagina.html'), 'utf8').replace(/\r\n/g, '\n');
  if (modello.split(SEGNO).length !== 4) throw new Error('rinvio/pagina.html: attesi 3 punti «' + SEGNO + '»');
  const pagina = (dest) => Buffer.from(modello.split(SEGNO).join(dest));
  const fisso = (nome) => Buffer.from(fs.readFileSync(path.join(MODELLI, nome), 'utf8').replace(/\r\n/g, '\n'));
  return [
    ['index.html', pagina(base)],
    ['print.html', pagina(base + 'print.html')],
    ['404.html', pagina(base)],
    ['sw.js', fisso('sw.js')],
    ['robots.txt', fisso('robots.txt')],
    ['README.md', fisso('README.md')],
    ['.nojekyll', Buffer.from('')],
  ];
}

module.exports = { fileRinvio, destinazione, RADICE };

if (require.main === module) {
  try {
    const file = fileRinvio(process.argv[2]);
    const dove = process.argv[3];
    if (dove) {
      const dest = path.resolve(dove);
      if (!dest.startsWith(RADICE + path.sep)) throw new Error('la cartella deve stare dentro il repository');
      // La cartella viene svuotata: si accetta solo una cartella che non esiste, vuota,
      // o che contiene soltanto i file di un rinvio preparato prima.
      if (fs.existsSync(dest)) {
        if (!fs.statSync(dest).isDirectory()) throw new Error('la destinazione esiste e non è una cartella: indicarne una nuova o vuota');
        const noti = file.map((f) => f[0]);
        const altri = fs.readdirSync(dest).filter((x) => noti.indexOf(x) < 0);
        if (altri.length) throw new Error('la cartella esiste e contiene altro (' + altri.slice(0, 3).join(', ') + '): non la svuoto. Indicarne una nuova o vuota');
      }
      fs.rmSync(dest, { recursive: true, force: true });
      fs.mkdirSync(dest, { recursive: true });
      file.forEach((f) => fs.writeFileSync(path.join(dest, f[0]), f[1]));
    }
    file.forEach((f) => console.log('  ' + f[0] + ' (' + f[1].length + ' byte)'));
    console.log(file.length + ' file per il rinvio a ' + destinazione(process.argv[2]) + (dove ? ', scritti in ' + dove : ''));
  } catch (e) { console.log('ERRORE: ' + e.message); process.exit(1); }
}
