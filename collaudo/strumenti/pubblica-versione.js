// Dopo una pubblicazione sul sito di COLLAUDO: aspetta che GitHub Pages serva la
// versione nuova e poi aggiorna app_version nel database di collaudo, come fa in
// produzione il workflow notify-deploy. È quell'aggiornamento a far comparire il
// badge «Update» (o il ricaricamento automatico) sulle pagine aperte.
//
//   node pubblica-versione.js            -> usa il commit in cima al ramo corrente
const { execSync } = require('child_process');
const fs = require('fs');
const path = require('path');
const { collaudo } = require('./sb.js');

const SITO = 'https://gistech2026.github.io/app-consegne/';
const RADICE = path.resolve(__dirname, '../..');
const attendi = (ms) => new Promise((r) => setTimeout(r, ms));

(async () => {
  const sha = execSync('git rev-parse HEAD', { cwd: RADICE, encoding: 'utf8' }).trim();
  const attesa = (/CACHE_NAME\s*=\s*'([^']+)'/.exec(fs.readFileSync(path.join(RADICE, 'docs', 'sw.js'), 'utf8')) || [])[1];
  let servita = '';
  for (let i = 0; i < 40; i++) {
    try {
      const t = await (await fetch(SITO + 'sw.js?cb=' + Date.now())).text();
      servita = (/CACHE_NAME\s*=\s*'([^']+)'/.exec(t) || [])[1] || '';
    } catch (e) { servita = ''; }
    if (servita === attesa) break;
    await attendi(8000);
  }
  if (servita !== attesa) { console.log('ERRORE: il sito serve «' + servita + '» invece di «' + attesa + '»'); process.exit(1); }
  const r = await collaudo().query("update public.app_version set sha = '" + sha.replace(/[^0-9a-f]/g, '') + "', deployed_at = " + Date.now() + ", message = 'deploy " + sha.slice(0, 7) + "' where id = 1 returning left(sha, 7) as sha");
  console.log('sito di collaudo: ' + servita + ' | app_version: ' + r[0].sha);
})().catch((e) => { console.log('ERRORE: ' + e.message); process.exit(1); });
