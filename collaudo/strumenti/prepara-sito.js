// Prepara la cartella da pubblicare su Cloudflare Pages: una copia dei SOLI file
// del sito. Sono i file di docs/ registrati in git, tolto tutto ciò che sta in
// una cartella dal nome che comincia con «_» o con «.» (come docs/_legacy): gli
// stessi che pubblica GitHub Pages. In più passano i due file di configurazione
// di Cloudflare, _headers e _redirects, se ci sono. I file sono copiati così
// come sono, byte per byte: ciò che si prova è ciò che si pubblica.
//
//   node collaudo/strumenti/prepara-sito.js <cartella>   prepara la cartella
//   node collaudo/strumenti/prepara-sito.js              stampa solo l'elenco
//
// Usato dal flusso .github/workflows/pubblica-cloudflare.yml, dai controlli di
// rilascio e dal banco del service worker (banco-sw.js).
const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');

const RADICE = path.resolve(__dirname, '../..');
const CONFIGURAZIONE = ['_headers', '_redirects'];

// Percorsi relativi a docs/, con la barra «/».
function elencoSito(riferimento) {
  const args = riferimento ? ['ls-tree', '-r', '-z', '--name-only', riferimento, 'docs'] : ['ls-files', '-z', 'docs'];
  const tutti = execFileSync('git', ['-C', RADICE].concat(args), { encoding: 'utf8', maxBuffer: 16 * 1024 * 1024 })
    .split('\0').filter(Boolean).map((p) => p.replace(/^docs\//, ''));
  return tutti.filter((p) => {
    const pezzi = p.split('/');
    if (pezzi.length === 1 && CONFIGURAZIONE.indexOf(p) >= 0) return true;
    return !pezzi.some((x) => x.charAt(0) === '_' || x.charAt(0) === '.');
  }).sort();
}

function prepara(destinazione) {
  const dest = path.resolve(destinazione);
  if (dest === RADICE || dest === path.join(RADICE, 'docs') || !dest.startsWith(RADICE + path.sep)) {
    throw new Error('la cartella di destinazione deve stare dentro il repository e non essere docs/: ' + dest);
  }
  fs.rmSync(dest, { recursive: true, force: true });
  const elenco = elencoSito();
  let byte = 0;
  elenco.forEach((p) => {
    const da = path.join(RADICE, 'docs', p), a = path.join(dest, p);
    if (!fs.existsSync(da)) throw new Error('file registrato in git ma assente: docs/' + p);
    fs.mkdirSync(path.dirname(a), { recursive: true });
    fs.copyFileSync(da, a);
    byte += fs.statSync(a).size;
  });
  return { elenco, byte };
}

module.exports = { elencoSito, prepara, RADICE };

if (require.main === module) {
  try {
    const dest = process.argv[2];
    if (!dest) { elencoSito().forEach((p) => console.log(p)); process.exit(0); }
    const r = prepara(dest);
    r.elenco.forEach((p) => console.log('  ' + p));
    console.log(r.elenco.length + ' file, ' + r.byte + ' byte in ' + path.relative(RADICE, path.resolve(dest)).split(path.sep).join('/') + '/');
  } catch (e) { console.log('ERRORE: ' + e.message); process.exit(1); }
}
