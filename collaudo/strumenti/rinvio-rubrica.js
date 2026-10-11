// Il vecchio indirizzo della rubrica. Fino al 09/10/2026 la rubrica telefonica
// era un sito a parte (…github.io/rubrica-gemelli/), con la sua icona sui
// telefoni dei medici. Da allora è una pagina dell'applicazione delle consegne,
// e il suo repository è stato tolto: chi tocca la vecchia icona trova «pagina
// non trovata». Questo strumento tiene a quell'indirizzo la SOLA pagina di
// rinvio verso l'app (nessun codice, nessun dato, nessuna chiave) e un service
// worker che toglie dai browser quello della rubrica di prima, con la sua cache.
//
//   node collaudo/strumenti/rinvio-rubrica.js file [cartella]   elenca i file; con la cartella li scrive
//                                                               (solo una sottocartella di _sito)
//   node collaudo/strumenti/rinvio-rubrica.js stato             sola lettura: repository, sito e indirizzo
//   node collaudo/strumenti/rinvio-rubrica.js pubblica --confermo-produzione
//        crea (o aggiorna) il repository pubblico, con un solo commit senza storia,
//        accende il suo sito, lo chiude alle scritture e aspetta che risponda
//
// È uno strumento della sola PRODUZIONE: l'indirizzo da tenere vivo è quello
// dell'account del reparto. Cambia qualcosa solo «pubblica», e solo con la
// conferma esplicita. Nessuna chiave viene mai stampata.
//
// ATTENZIONE alla copia di sicurezza della rubrica di prima (fuori da questo
// repository, in D:\Progetti AI\RUBRICA): è una copia a specchio che ha per
// remoto proprio questo nome. Lì dentro non si lancia mai «git fetch» né «git
// push»: il primo sposterebbe la copia sul commit del rinvio, il secondo
// rimetterebbe in pubblico la storia di prima (le regole di sola lettura del
// repository lo rifiutano, ma solo finché ci sono).
const fs = require('fs');
const path = require('path');
const ghApi = require('./gh-api.js');
const reparti = require('./reparti.js');

const RADICE = path.resolve(__dirname, '../..');
const MODELLI = path.join(RADICE, 'rinvio', 'rubrica');
const SEGNO = '{{DESTINAZIONE}}';
const NOME = 'rubrica-gemelli';
// Il segno con cui l'app capisce che la si è aperta per la rubrica.
const SEGNO_RUBRICA = '?rubrica=1';
// Come per la pagina di rinvio dell'app: una regola per i rami e una per le
// etichette, così nel repository pubblico non può arrivare nient'altro.
const REGOLE = [
  { name: 'sola-lettura', target: 'branch', rules: [{ type: 'creation' }, { type: 'update' }, { type: 'deletion' }, { type: 'non_fast_forward' }] },
  { name: 'sola-lettura-etichette', target: 'tag', rules: [{ type: 'creation' }, { type: 'update' }, { type: 'deletion' }] },
];
const attendi = (ms) => new Promise((r) => setTimeout(r, ms));
const pulito = (t) => ghApi.maschera(String(t)).replace(/\/\/[^/@\s]+@/g, '//***@');

// [percorso nel sito, contenuto] per ogni file. «dest» è l'indirizzo intero a cui
// si rimanda: qui non lo si giudica (il banco di prova ne usa uno locale), lo
// giudica destinazioneVera() per ciò che si pubblica.
function fileRubrica(dest) {
  if (typeof dest !== 'string' || !dest || /['"<>\\\s]/.test(dest)) throw new Error('destinazione non valida');
  const leggi = (dove) => fs.readFileSync(dove, 'utf8').replace(/\r\n/g, '\n');
  const modello = leggi(path.join(MODELLI, 'pagina.html'));
  if (modello.split(SEGNO).length !== 4) throw new Error('rinvio/rubrica/pagina.html: attesi 3 punti «' + SEGNO + '»');
  const pagina = Buffer.from(modello.split(SEGNO).join(dest));
  return [
    ['index.html', pagina],
    ['404.html', pagina],
    ['sw.js', Buffer.from(leggi(path.join(MODELLI, 'sw.js')))],
    ['robots.txt', Buffer.from(leggi(path.join(RADICE, 'rinvio', 'robots.txt')))],
    ['README.md', Buffer.from(leggi(path.join(MODELLI, 'README.md')))],
    ['.nojekyll', Buffer.from('')],
  ];
}

// L'indirizzo vero: l'app del reparto che pubblica da docs/ con GitHub Pages,
// sullo stesso account del vecchio indirizzo della rubrica.
function destinazioneVera(utente) {
  const r = reparti.delSitoInDocs();
  const [proprietario, nome] = r.pubblicazione.githubPages.repository.split('/');
  if (proprietario.toLowerCase() !== String(utente).toLowerCase()) throw new Error('il reparto che pubblica da docs/ non è di questo account');
  const host = proprietario.toLowerCase() + '.github.io';
  if (!reparti.hostAmmesso(r, host)) throw new Error('l\'indirizzo ' + host + ' non è fra quelli del reparto «' + r.codice + '»');
  if (!/^[a-z0-9-]+$/.test(nome)) throw new Error('nome del repository dell\'app inatteso');
  return { app: 'https://' + host + '/' + nome + '/', intera: 'https://' + host + '/' + nome + '/' + SEGNO_RUBRICA, vecchio: 'https://' + host + '/' + NOME + '/', reparto: r.codice };
}

module.exports = { fileRubrica, destinazioneVera, NOME, SEGNO_RUBRICA };

if (require.main === module) {
  let aperto = false;   // vero da quando le regole di sola lettura sono state tolte a quando sono state rimesse
  (async () => {
    const argomenti = process.argv.slice(2).filter((x) => x !== ghApi.CONFERMA);
    const comando = argomenti[0];
    if (['file', 'stato', 'pubblica'].indexOf(comando) < 0 || argomenti.length > (comando === 'file' ? 2 : 1)) {
      console.log('uso: rinvio-rubrica.js file [cartella] | stato | pubblica ' + ghApi.CONFERMA);
      process.exitCode = 2; return;
    }
    const AMB = ghApi.per('produzione');
    const gh = AMB.gh, UTENTE = AMB.UTENTE;
    const D = destinazioneVera(UTENTE);
    const base = '/repos/' + UTENTE + '/' + NOME;
    const api = async (metodo, percorso, corpo, ammessi) => {
      const r = await gh(metodo, percorso, corpo);
      if (r.stato >= 300 && !(ammessi || []).includes(r.stato)) throw new Error(metodo + ' ' + percorso + ' → HTTP ' + r.stato + ' ' + JSON.stringify(r.dati).slice(0, 300));
      return r;
    };
    const repo = async () => { const r = await api('GET', base, null, [404]); return (r.stato === 404 || !r.dati || r.dati.name !== NOME) ? null : r.dati; };
    const regolePresenti = async () => {
      const r = await api('GET', base + '/rulesets', null, [404, 403]);
      const nomi = REGOLE.map((x) => x.name);
      return (Array.isArray(r.dati) ? r.dati : []).filter((x) => nomi.indexOf(x.name) >= 0);
    };
    const chiudi = async () => {
      const presenti = (await regolePresenti()).map((x) => x.name);
      for (const x of REGOLE) {
        if (presenti.indexOf(x.name) >= 0) continue;
        await api('POST', base + '/rulesets', { name: x.name, target: x.target, enforcement: 'active', bypass_actors: [], conditions: { ref_name: { include: ['~ALL'], exclude: [] } }, rules: x.rules });
      }
      aperto = false;
    };
    // Che cosa risponde il vecchio indirizzo e, se è la pagina di rinvio, dove porta.
    const cheRisponde = async (file) => {
      try {
        const r = await fetch(D.vecchio + (file || '') + '?cb=' + Date.now(), { redirect: 'manual', cache: 'no-store' });
        const t = await r.text();
        const m = /var NUOVO = '([^']*)'/.exec(t);
        return { stato: r.status, rinvio: m ? m[1] : null, testo: t };
      } catch (e) { return { stato: 0, rinvio: null, testo: '' }; }
    };
    const swBuono = (sw) => sw.stato === 200 && sw.testo.indexOf("n.indexOf('rubrica-') === 0") >= 0 && sw.testo.indexOf('registration.unregister()') >= 0 && sw.testo.indexOf("addEventListener('fetch'") < 0;
    // L'app in linea: deve rispondere; e si guarda se conosce già il segno della rubrica.
    const appInLinea = async () => {
      const a = await fetch(D.app, { cache: 'no-store' }).then(async (r) => ({ stato: r.status, testo: await r.text() }), () => ({ stato: 0, testo: '' }));
      return { risponde: a.stato === 200 && a.testo.indexOf('navVersioneApp') >= 0, stato: a.stato, conosceIlSegno: a.testo.indexOf('rubrica=1') >= 0 || /[?&]rubrica\b|get\('rubrica'\)/.test(a.testo) };
    };

    if (comando === 'file') {
      const file = fileRubrica(D.intera);
      if (argomenti[1]) {
        const dest = path.resolve(argomenti[1]), SITI = path.join(RADICE, '_sito');
        if (!dest.startsWith(SITI + path.sep)) throw new Error('la cartella viene svuotata: può essere solo una sottocartella di _sito');
        fs.rmSync(dest, { recursive: true, force: true });
        fs.mkdirSync(dest, { recursive: true });
        file.forEach((f) => fs.writeFileSync(path.join(dest, f[0]), f[1]));
      }
      file.forEach((f) => console.log('  ' + f[0] + ' (' + f[1].length + ' byte)'));
      console.log(file.length + ' file: dal vecchio indirizzo ' + D.vecchio + ' si viene portati a ' + D.intera + (argomenti[1] ? ' (scritti in ' + argomenti[1] + ')' : ''));
      return;
    }

    if (comando === 'stato') {
      const r = await repo();
      let pagine = '';
      if (r && r.has_pages) { const p = await api('GET', base + '/pages', null, [404]); pagine = p.stato === 200 ? ' (' + String(p.dati.status || 'stato non detto') + ')' : ''; }
      console.log('repository ' + UTENTE + '/' + NOME + ': ' + (r ? ((r.private ? 'PRIVATO' : 'pubblico') + ', sito ' + (r.has_pages ? 'acceso' + pagine : 'spento') + ', regole di sola lettura: ' + (await regolePresenti()).length + ' su ' + REGOLE.length) : 'non esiste'));
      const v = await cheRisponde('');
      console.log('il vecchio indirizzo risponde: HTTP ' + v.stato + (v.rinvio ? ', pagina di rinvio verso ' + v.rinvio + (v.rinvio === D.intera ? ' (giusta)' : ' (NON è quella attesa: ' + D.intera + ')') : ''));
      if (v.rinvio) console.log('service worker di pulizia: ' + (swBuono(await cheRisponde('sw.js')) ? 'al suo posto' : 'NON servito come atteso'));
      const a = await appInLinea();
      console.log('app in linea: ' + (a.risponde ? 'risponde' : 'NON risponde (HTTP ' + a.stato + ')') + '; il segno «' + SEGNO_RUBRICA + '» ' + (a.conosceIlSegno ? 'è riconosciuto: apre la rubrica' : 'NON è ancora riconosciuto: chi arriva dal vecchio indirizzo vede l\'app intera, e la rubrica si apre dal menu'));
      return;
    }

    // ── pubblica ──
    // L'app deve rispondere, altrimenti si rimanderebbe a una pagina che non c'è.
    const a = await appInLinea();
    if (!a.risponde) throw new Error('l\'app ' + D.app + ' non risponde come atteso (HTTP ' + a.stato + ')');
    if (!a.conosceIlSegno) console.log('NOTA: l\'app in linea non riconosce ancora il segno «' + SEGNO_RUBRICA + '». Finché non esce la versione che lo legge, chi arriva dal vecchio indirizzo vede l\'app intera e apre la rubrica dal menu «Numeri Telefono».');
    // La chiave deve essere di questo account: «crea repository» è l'unica chiamata che non lo dice nel percorso.
    const io = await api('GET', '/user');
    if (String((io.dati || {}).login || '').toLowerCase() !== UTENTE.toLowerCase()) throw new Error('la chiave di GitHub trovata non è dell\'account ' + UTENTE);
    let r = await repo();
    if (r && r.private) throw new Error('il repository ' + NOME + ' esiste ed è privato: non lo rendo pubblico da qui');
    if (!r) {
      await api('POST', '/user/repos', { name: NOME, description: "Rinvio alla rubrica dentro l'applicazione delle consegne", private: false, auto_init: true, has_issues: false, has_wiki: false, has_projects: false });
      console.log('repository ' + UTENTE + '/' + NOME + ' creato');
      aperto = true;   // appena nato non ha ancora le regole di sola lettura
      await attendi(3000);
      r = await repo();
      if (!r) throw new Error('il repository appena creato non risponde ancora: rilanciare');
    }
    const ramo = r.default_branch;
    const albero = [];
    for (const [dove, corpo] of fileRubrica(D.intera)) {
      const blob = await api('POST', base + '/git/blobs', { content: corpo.toString('base64'), encoding: 'base64' });
      albero.push({ path: dove, mode: '100644', type: 'blob', sha: blob.dati.sha });
    }
    const tree = await api('POST', base + '/git/trees', { tree: albero });
    const testa = (await api('GET', base + '/git/ref/heads/' + ramo)).dati.object.sha;
    const attuale = (await api('GET', base + '/git/commits/' + testa)).dati;
    let giaGiusta = attuale.tree.sha === tree.dati.sha && !attuale.parents.length;
    if (giaGiusta) console.log('la pagina di rinvio è già quella giusta');
    else {
      // Solo per spostare il ramo le regole di sola lettura vanno tolte un momento; si rimettono subito sotto.
      aperto = true;
      for (const x of await regolePresenti()) await api('DELETE', base + '/rulesets/' + x.id);
      // un solo commit, senza storia: nel repository pubblico non resta altro
      const commit = await api('POST', base + '/git/commits', { message: 'Pagina di rinvio alla rubrica', tree: tree.dati.sha, parents: [] });
      await api('PATCH', base + '/git/refs/heads/' + ramo, { sha: commit.dati.sha, force: true });
      console.log('pagina di rinvio pubblicata (' + albero.length + ' file) verso ' + D.intera);
    }
    // stesso ordine dello strumento gemello, già provato: prima il sito, poi le regole
    const pagine = await api('POST', base + '/pages', { source: { branch: ramo, path: '/' } }, [409]);
    console.log(pagine.stato < 300 ? 'GitHub Pages attivato' : 'GitHub Pages già attivo');
    await chiudi();
    console.log('repository chiuso alle scritture (' + REGOLE.length + ' regole)');
    // Si aspetta che il vecchio indirizzo risponda con la pagina giusta. Se GitHub Pages non parte da solo (la
    // pagina era già quella giusta, oppure la sua pubblicazione è andata in errore) gli si chiede di ripubblicare.
    const t0 = Date.now(), QUANDO = giaGiusta ? [0, 60000, 150000] : [45000, 130000, 220000];
    let v = { stato: 0, rinvio: null }, spinte = 0;
    for (let i = 0; i < 80; i++) {
      v = await cheRisponde('');
      if (v.stato === 200 && v.rinvio === D.intera) break;
      if (spinte < QUANDO.length && Date.now() - t0 >= QUANDO[spinte]) {
        spinte++;
        try { const x = await gh('POST', base + '/pages/builds'); console.log('chiesta a GitHub Pages una pubblicazione (HTTP ' + x.stato + ')'); } catch (e) { console.log('richiesta di pubblicazione a GitHub Pages non riuscita: continuo ad aspettare'); }
      }
      await attendi(5000);
    }
    if (v.stato !== 200 || v.rinvio !== D.intera) throw new Error('dopo ' + Math.round((Date.now() - t0) / 1000) + ' s il vecchio indirizzo risponde ancora HTTP ' + v.stato + (v.rinvio ? ' con un rinvio verso ' + v.rinvio : '') + ': rilanciare «pubblica» fra qualche minuto, che chiede di nuovo la pubblicazione');
    let sw = { stato: 0, testo: '' };
    for (let i = 0; i < 8; i++) { sw = await cheRisponde('sw.js'); if (swBuono(sw)) break; await attendi(8000); }
    console.log('pronto dopo ' + Math.round((Date.now() - t0) / 1000) + ' s: ' + D.vecchio + ' rimanda a ' + D.intera + '; service worker di pulizia ' + (swBuono(sw) ? 'al suo posto' : 'NON ancora servito come atteso (HTTP ' + sw.stato + ')'));
    if (!swBuono(sw)) process.exitCode = 1;
  })().catch((e) => {
    console.log('ERRORE: ' + pulito((e && e.message) || e));
    if (aperto) console.log('ATTENZIONE: il repository può essere rimasto APERTO alle scritture. Lo dice «stato»; rilanciare «pubblica» lo richiude.');
    process.exitCode = 1;
  });
}
