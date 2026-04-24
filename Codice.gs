// Costanti
const NOME_FOGLIO = "Consegne";
const NOME_ARCHIVIO = "Archivio_Dati";
const NOME_TIPOLOGIE = "Tipologie"; // ← NUOVO: foglio dedicato alle tipologie
const LOCK_DURATA_MS = 30000; // 30s: lock scade se non rinnovato

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// Rimuove tag non consentiti dall'HTML dei campi rich-text.
// Tag consentiti: b, i, u, strong, em, span, font, div, br.
// Per i tag non consentiti: se hanno contenuto testuale lo mantiene (es. <a>testo</a> → testo),
// se sono tag vuoti o pericolosi (script, style, iframe, img) li rimuove completamente col contenuto.
function sanitizzaHtmlCampo(html) {
  if (!html || typeof html !== 'string') return html;
  // Rimuove completamente i tag pericolosi (con tutto il loro contenuto)
  var pericolosi = /(<(script|style|iframe|object|embed|form|input|button|select|textarea|meta|link|base)[^>]*>[\s\S]*?<\/\2>|<(script|style|iframe|object|embed|form|input|button|select|textarea|meta|link|base)[^>]*\/?>)/gi;
  html = html.replace(pericolosi, '');
  // Sostituisce tag non consentiti mantenendo il testo interno
  // Gestisce prima i tag con chiusura
  var consentiti = /^(b|i|u|strong|em|span|font|div|br)$/i;
  html = html.replace(/<\/?([\w-]+)([^>]*)>/gi, function(match, tag, attrs) {
    if (consentiti.test(tag)) return match;
    // Tag non consentito: rimuove solo il tag, il testo interno rimane
    return '';
  });
  // Protezione template GAS: evita sequenze che chiuderebbero i tag <? ?>
  html = html.replace(/\?>/g, '? >');
  return html;
}

const CAMPI_RICH_TEXT = ['NoteTerapia', 'Diaria', 'DaFare', 'PianoTerapeutico', 'Allergie'];


function doGet(e) {
  var params = (e && e.parameter) ? e.parameter : {};

  if (params.page === 'print') {
    try {
      var tmpl = HtmlService.createTemplateFromFile('Print');
      tmpl.layout       = params.layout       || 'main';
      tmpl.orientamento = params.orientamento === 'portrait' ? 'portrait' : 'landscape';
      tmpl.ordinamento  = params.ordinamento  || 'tipologia';
      tmpl.saltaVuoti   = params.saltaVuoti   === '1';
      tmpl.dataArchivio = params.dataArchivio || '';
      tmpl.scala        = params.scala ? parseInt(params.scala, 10) : 0;
      tmpl.tipologie  = params.tipologie  || 'all';
      tmpl.appUrl       = ScriptApp.getService().getUrl();
      return tmpl.evaluate()
        .setTitle('Stampa Consegne')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
    } catch(ex) {
      // Print.html non ancora creato: ricade sulla pagina principale
    }
  }

  var indexTmpl = HtmlService.createTemplateFromFile('Index');
  indexTmpl.toastType   = params.toast       || '';
  indexTmpl.toastMsg    = params.msg         || '';
  indexTmpl.initialLoad = true; // skip getDatiPazienti() server-side: le card vengono caricate client-side via google.script.run
  return indexTmpl.evaluate()
    .setTitle('Gestione Consegne')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}


function getScriptUrl() { return ScriptApp.getService().getUrl(); }


function ottieniNomeReparto() {
  return PropertiesService.getScriptProperties().getProperty('NOME_REPARTO') || "Consegne Reparto";
}
function salvaNomeReparto(nuovoNome) {
  PropertiesService.getScriptProperties().setProperty('NOME_REPARTO', nuovoNome);
  return nuovoNome;
}


// ==========================================
// LETTURA DATI
// ==========================================
function getDatiPazienti(cacheBuster) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(NOME_FOGLIO);
  if (!sheet) { sheet = ss.insertSheet(NOME_FOGLIO); sheet.appendRow(["Letto","Nome","Eta","DataRicovero","Diagnosi","NoteTerapia","Diaria","DaFare","TipologiaLetto","PianoTerapeutico","UltimoAggiornamento"]); }
  let data = sheet.getDataRange().getValues(); let headers = data[0];
  if (headers.indexOf("TipologiaLetto") === -1) { sheet.getRange(1, headers.length+1).setValue("TipologiaLetto"); data = sheet.getDataRange().getValues(); headers = data[0]; }
  if (headers.indexOf("PianoTerapeutico") === -1) { sheet.getRange(1, headers.length+1).setValue("PianoTerapeutico"); data = sheet.getDataRange().getValues(); headers = data[0]; }
  if (headers.indexOf("Allergie") === -1) { sheet.getRange(1, headers.length+1).setValue("Allergie"); data = sheet.getDataRange().getValues(); headers = data[0]; }
  if (headers.indexOf("DataNascita") === -1) { sheet.getRange(1, headers.length+1).setValue("DataNascita"); data = sheet.getDataRange().getValues(); headers = data[0]; }
  if (headers.indexOf("CodiceSanitario") === -1) { sheet.getRange(1, headers.length+1).setValue("CodiceSanitario"); data = sheet.getDataRange().getValues(); headers = data[0]; }
  if (headers.indexOf("Ossigeno") === -1) { sheet.getRange(1, headers.length+1).setValue("Ossigeno"); data = sheet.getDataRange().getValues(); headers = data[0]; }
  if (headers.indexOf("Sesso") === -1) { sheet.getRange(1, headers.length+1).setValue("Sesso"); data = sheet.getDataRange().getValues(); headers = data[0]; }
  if (data.length <= 1) return [];
  const pazienti = [];
  for (let i = 1; i < data.length; i++) {
    let row = data[i]; let numeroLetto = String(row[0]).trim();
    if (numeroLetto !== "") {
      let paziente = {};
      for (let j = 0; j < headers.length; j++) {
        let val = row[j];
        if (CAMPI_RICH_TEXT.indexOf(headers[j]) !== -1) val = sanitizzaHtmlCampo(val);
        paziente[headers[j]] = val;
      }
      pazienti.push(paziente);
    }
  }
  pazienti.sort((a,b) => { let A=parseInt(a.Letto,10),B=parseInt(b.Letto,10); if(!isNaN(A)&&!isNaN(B))return A-B; return String(a.Letto).localeCompare(String(b.Letto)); });
  return pazienti;
}


function getNewHtml() {
  return HtmlService.createTemplateFromFile('Cards').evaluate().getContent();
}


// ==========================================
// GESTIONE LOCK
// ==========================================
function _getCacheKey() { return 'ward_locks_v2'; }


function acquistaLock(letto, token) {
  var scriptLock = LockService.getScriptLock();
  try {
    scriptLock.waitLock(3000);
  } catch(e) {
    return { success: false, blocked: false, message: 'Server occupato, riprova.' };
  }
  try {
    var cache = CacheService.getScriptCache();
    var locks = JSON.parse(cache.get(_getCacheKey()) || '{}');
    var now = Date.now();
    var k = String(letto);
    Object.keys(locks).forEach(function(kk) { if (now - locks[kk].ts > LOCK_DURATA_MS) delete locks[kk]; });
    var existing = locks[k];
    if (existing && existing.token !== token) {
      cache.put(_getCacheKey(), JSON.stringify(locks), 120);
      return { success: false, blocked: true, message: 'Scheda in aggiornamento da altro utente.' };
    }
    locks[k] = { token: token, ts: now };
    cache.put(_getCacheKey(), JSON.stringify(locks), 120);
    return { success: true, blocked: false };
  } finally {
    scriptLock.releaseLock();
  }
}


function rilasciaLock(letto, token) {
  var scriptLock = LockService.getScriptLock();
  try { scriptLock.waitLock(2000); } catch(e) { return { success: false }; }
  try {
    var cache = CacheService.getScriptCache();
    var locks = JSON.parse(cache.get(_getCacheKey()) || '{}');
    var k = String(letto);
    if (locks[k] && locks[k].token === token) delete locks[k];
    cache.put(_getCacheKey(), JSON.stringify(locks), 120);
    return { success: true };
  } finally {
    scriptLock.releaseLock();
  }
}


function acquistaLockMultiplo(letti, token) {
  var scriptLock = LockService.getScriptLock();
  try { scriptLock.waitLock(4000); } catch(e) { return { success: false, blocked: false, message: 'Server occupato, riprova.' }; }
  try {
    var cache = CacheService.getScriptCache();
    var locks = JSON.parse(cache.get(_getCacheKey()) || '{}');
    var now = Date.now();
    Object.keys(locks).forEach(function(kk) { if (now - locks[kk].ts > LOCK_DURATA_MS) delete locks[kk]; });
    var bloccati = [];
    letti.forEach(function(l) { var k = String(l); if (locks[k] && locks[k].token !== token) bloccati.push(l); });
    if (bloccati.length > 0) {
      cache.put(_getCacheKey(), JSON.stringify(locks), 120);
      return { success: false, blocked: true, bloccati: bloccati };
    }
    letti.forEach(function(l) { locks[String(l)] = { token: token, ts: now }; });
    cache.put(_getCacheKey(), JSON.stringify(locks), 120);
    return { success: true };
  } finally { scriptLock.releaseLock(); }
}


function rilasciaLockMultiplo(letti, token) {
  var scriptLock = LockService.getScriptLock();
  try { scriptLock.waitLock(2000); } catch(e) { return { success: false }; }
  try {
    var cache = CacheService.getScriptCache();
    var locks = JSON.parse(cache.get(_getCacheKey()) || '{}');
    letti.forEach(function(l) { var k = String(l); if (locks[k] && locks[k].token === token) delete locks[k]; });
    cache.put(_getCacheKey(), JSON.stringify(locks), 120);
    return { success: true };
  } finally { scriptLock.releaseLock(); }
}


function getLocks() {
  try {
    var cache = CacheService.getScriptCache();
    var locks = JSON.parse(cache.get(_getCacheKey()) || '{}');
    var now = Date.now();
    var attivi = {};
    Object.keys(locks).forEach(function(k) { if (now - locks[k].ts <= LOCK_DURATA_MS) attivi[k] = locks[k]; });
    return attivi;
  } catch(e) { return {}; }
}


// ==========================================
// SALVATAGGIO
// ==========================================
function autoSavePazienteCompleto(letto, datiPaziente, token) {
  var locks = getLocks();
  var k = String(letto);
  if (locks[k] && locks[k].token !== token) {
    return { success: false, message: 'Salvataggio rifiutato: la scheda è in aggiornamento da un altro utente.' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(NOME_FOGLIO);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) { if (String(data[i][0]) === String(letto)) { rowIndex = i+1; break; } }
  if (rowIndex === -1) return { success: false, message: 'Letto non trovato.' };

  const CAMPI_TESTO = ['Ossigeno','CodiceSanitario','Allergie','Nome','Diagnosi','Eta','NoteTerapia','Diaria','DaFare','PianoTerapeutico','TipologiaLetto','DataNascita','Sesso'];
  for (let campo in datiPaziente) {
    let colIndex = headers.indexOf(campo);
    if (colIndex !== -1) {
      var valore = datiPaziente[campo];
      if (CAMPI_RICH_TEXT.indexOf(campo) !== -1) valore = sanitizzaHtmlCampo(valore);
      var cella = sheet.getRange(rowIndex, colIndex+1);
      if (CAMPI_TESTO.indexOf(campo) !== -1) cella.setNumberFormat('@');
      cella.setValue(valore);
    }
  }
  const now = new Date();
  const oraFormattata = ('0'+now.getHours()).slice(-2)+":"+('0'+now.getMinutes()).slice(-2)+":"+('0'+now.getSeconds()).slice(-2);
  let tsCol = headers.indexOf("UltimoAggiornamento");
  if (tsCol !== -1) sheet.getRange(rowIndex, tsCol+1).setValue(oraFormattata);
  SpreadsheetApp.flush();
  return { success: true, ora: oraFormattata };
}


// ==========================================
// GESTIONE LETTI
// ==========================================
function aggiungiLetto(numeroLetto) { const ss=SpreadsheetApp.getActiveSpreadsheet(); const sheet=ss.getSheetByName(NOME_FOGLIO); const data=sheet.getDataRange().getValues(); const headers=data[0]; for(let i=1;i<data.length;i++){if(String(data[i][0])===String(numeroLetto))return{success:false,message:"Il letto esiste già!"};} let newRow=new Array(headers.length).fill(""); newRow[0]=numeroLetto; sheet.appendRow(newRow); SpreadsheetApp.flush(); return{success:true,message:"Letto aggiunto."}; }
function eliminaLetto(numeroLetto) { const ss=SpreadsheetApp.getActiveSpreadsheet(); const sheet=ss.getSheetByName(NOME_FOGLIO); const data=sheet.getDataRange().getValues(); for(let i=1;i<data.length;i++){if(String(data[i][0])===String(numeroLetto)){if(String(data[i][1]).trim()!==""||String(data[i][4]).trim()!=="")return{success:false,message:"Il letto non è vuoto."}; sheet.deleteRow(i+1); SpreadsheetApp.flush(); return{success:true,message:"Letto eliminato."};}} return{success:false,message:"Letto non trovato."}; }
function dimettiLetto(numeroLetto) { const ss=SpreadsheetApp.getActiveSpreadsheet(); const sheet=ss.getSheetByName(NOME_FOGLIO); const data=sheet.getDataRange().getValues(); const headers=data[0]; for(let i=1;i<data.length;i++){if(String(data[i][0])===String(numeroLetto)){for(let j=1;j<headers.length;j++){if(headers[j]!=="UltimoAggiornamento"&&headers[j]!=="TipologiaLetto")sheet.getRange(i+1,j+1).setValue("");} const dataColIndex=headers.indexOf("UltimoAggiornamento"); const now=new Date(); if(dataColIndex!==-1)sheet.getRange(i+1,dataColIndex+1).setValue(('0'+now.getHours()).slice(-2)+":"+('0'+now.getMinutes()).slice(-2)+":"+('0'+now.getSeconds()).slice(-2)); SpreadsheetApp.flush(); return{success:true,message:"Letto svuotato."};}} return{success:false,message:"Letto non trovato."}; }
function spostaPaziente(lettoOrigine,lettoDestinazione) { const ss=SpreadsheetApp.getActiveSpreadsheet(); const sheet=ss.getSheetByName(NOME_FOGLIO); const data=sheet.getDataRange().getValues(); const headers=data[0]; let rowOrigine=-1,rowDestinazione=-1; for(let i=1;i<data.length;i++){if(String(data[i][0])===String(lettoOrigine))rowOrigine=i+1; if(String(data[i][0])===String(lettoDestinazione))rowDestinazione=i+1;} if(rowOrigine===-1||rowDestinazione===-1)return{success:false,message:"Letto non trovato."}; const patientCols=["Nome","Eta","DataNascita","CodiceSanitario","DataRicovero","Diagnosi","NoteTerapia","Diaria","DaFare","PianoTerapeutico","Allergie","Ossigeno","Sesso"]; for(let colName of patientCols){let cIdx=headers.indexOf(colName); if(cIdx!==-1){let valO=sheet.getRange(rowOrigine,cIdx+1).getValue(); let valD=sheet.getRange(rowDestinazione,cIdx+1).getValue(); sheet.getRange(rowOrigine,cIdx+1).setValue(valD); sheet.getRange(rowDestinazione,cIdx+1).setValue(valO);}} SpreadsheetApp.flush(); return{success:true,message:"Spostati con successo!"}; }
function getLettiFull() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(NOME_FOGLIO);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const iLetto = headers.indexOf('Letto');
  const iNome  = headers.indexOf('Nome');
  const iTip   = headers.indexOf('TipologiaLetto');
  const result = [];
  for (let i = 1; i < data.length; i++) {
    const letto = String(data[i][iLetto] || '').trim();
    if (!letto) continue;
    const nome = iNome !== -1 ? String(data[i][iNome] || '').trim() : '';
    const tip  = iTip  !== -1 ? String(data[i][iTip]  || '').trim().toUpperCase() : '';
    result.push({ letto: letto, nome: nome, tipologia: tip || 'STANDARD' });
  }
  result.sort(function(a, b) {
    var nA = parseInt(a.letto, 10), nB = parseInt(b.letto, 10);
    return (!isNaN(nA) && !isNaN(nB)) ? nA - nB : a.letto.localeCompare(b.letto);
  });
  return result;
}

function getRiepilogoLetti() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(NOME_FOGLIO);
  if (!sheet) return { uomini:0, donne:0, indefinito:0, tipologie:{}, vuoti:0 };
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const iLetto    = headers.indexOf('Letto');
  const iNome     = headers.indexOf('Nome');
  const iSesso    = headers.indexOf('Sesso');
  const iTipologia = headers.indexOf('TipologiaLetto');
  let uomini = 0, donne = 0, indefinito = 0, vuoti = 0;
  const tipologie = {};
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (iLetto === -1 || String(row[iLetto] || '').trim() === '') continue;
    const nome     = iNome     !== -1 ? String(row[iNome]     || '').trim() : '';
    const sesso    = iSesso    !== -1 ? String(row[iSesso]    || '').trim() : '';
    const tipologia = iTipologia !== -1 ? String(row[iTipologia] || '').trim().toUpperCase() : '';
    const tip = tipologia || 'STANDARD';
    if (nome === '') { vuoti++; continue; }
    if      (sesso === 'M') uomini++;
    else if (sesso === 'F') donne++;
    else                     indefinito++;
    tipologie[tip] = (tipologie[tip] || 0) + 1;
  }
  return { uomini: uomini, donne: donne, indefinito: indefinito, tipologie: tipologie, vuoti: vuoti };
}

function modificaTipologiaLetto(letto,nuovaTipologia) { const ss=SpreadsheetApp.getActiveSpreadsheet(); const sheet=ss.getSheetByName(NOME_FOGLIO); const data=sheet.getDataRange().getValues(); const headers=data[0]; for(let i=1;i<data.length;i++){if(String(data[i][0])===String(letto)){let colIndex=headers.indexOf("TipologiaLetto"); if(colIndex!==-1)sheet.getRange(i+1,colIndex+1).setValue(nuovaTipologia); SpreadsheetApp.flush(); return{success:true,message:"Tipologia aggiornata."};}} return{success:false,message:"Letto non trovato."}; }


// ==========================================
// ARCHIVIO
// ==========================================
// Controlla se un backup è genuinamente in corso (con rilevamento lock stale)
function _isBackupInCorso(props) {
  if (props.getProperty('BACKUP_IN_CORSO') !== '1') return false;
  var ts = props.getProperty('BACKUP_LOCK_TS');
  if (!ts) return false;
  // Lock stale se non rinnovato da più di 20 secondi
  return (new Date().getTime() - parseInt(ts, 10)) < 20000;
}

// Esposta al client per il polling dalla schermata di attesa
function checkBackupStatus() {
  return { inCorso: _isBackupInCorso(PropertiesService.getScriptProperties()) };
}

// Esposta al client: rinnova il timestamp del lock ogni 5s durante il backup
function rinnovaLockBackup() {
  PropertiesService.getScriptProperties().setProperty('BACKUP_LOCK_TS', String(new Date().getTime()));
}

function archiviaGiornoCorrente() {
  var props = PropertiesService.getScriptProperties();
  var ora   = new Date().getTime();

  // ── 1. Controlla se backup già in corso ───────────────────────────
  if (_isBackupInCorso(props)) return { inCorso: true };

  // ── 2. Controlla cooldown 6 ore ───────────────────────────────────
  var ultimoTs = props.getProperty('ULTIMO_ARCHIVIO_TS');
  if (ultimoTs && (ora - parseInt(ultimoTs, 10)) < 6 * 60 * 60 * 1000) {
    return { eseguito: false };
  }

  // ── 3. Acquisisce lock atomico e imposta flag ─────────────────────
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(3000);
  } catch(e) {
    return { inCorso: true }; // qualcun altro ha appena acquisito il lock
  }
  // Doppio controllo post-lock
  if (_isBackupInCorso(props)) { lock.releaseLock(); return { inCorso: true }; }
  ultimoTs = props.getProperty('ULTIMO_ARCHIVIO_TS');
  if (ultimoTs && (ora - parseInt(ultimoTs, 10)) < 6 * 60 * 60 * 1000) {
    lock.releaseLock(); return { eseguito: false };
  }
  // Segna l'inizio: timestamp subito (cooldown vale anche in caso di errore)
  props.setProperties({
    'BACKUP_IN_CORSO':   '1',
    'BACKUP_LOCK_TS':    String(ora),
    'ULTIMO_ARCHIVIO_TS': String(ora)
  });
  lock.releaseLock(); // LockService rilasciato: il flag BACKUP_IN_CORSO è ora il nostro mutex

  // ── 4. Esegui archivio e backup ───────────────────────────────────
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetConsegne = ss.getSheetByName(NOME_FOGLIO);
    if (!sheetConsegne) return { eseguito: false };

    var dataConsegne = sheetConsegne.getDataRange().getValues();
    var headersCons  = dataConsegne[0];
    if (dataConsegne.length <= 1) return { eseguito: false };

    var iLetto = headersCons.indexOf('Letto');

    // Crea foglio archivio se non esiste
    var sheetArchivio = ss.getSheetByName(NOME_ARCHIVIO);
    if (!sheetArchivio) {
      sheetArchivio = ss.insertSheet(NOME_ARCHIVIO);
      sheetArchivio.appendRow(['DataArchivio'].concat(headersCons));
      SpreadsheetApp.flush();
    }

    // Consistenza campi: aggiungi colonne mancanti nell'archivio
    var archLastCol = sheetArchivio.getLastColumn();
    var archHeaders = sheetArchivio.getRange(1, 1, 1, archLastCol).getValues()[0];
    var archFieldSet = {};
    for (var i = 1; i < archHeaders.length; i++) archFieldSet[String(archHeaders[i]).trim()] = true;
    headersCons.forEach(function(campo) {
      campo = String(campo).trim();
      if (campo && !archFieldSet[campo]) {
        var newCol = sheetArchivio.getLastColumn() + 1;
        sheetArchivio.getRange(1, newCol).setValue(campo);
        archFieldSet[campo] = true;
      }
    });
    SpreadsheetApp.flush();

    // Rileggi header archivio dopo eventuali aggiunte
    archLastCol = sheetArchivio.getLastColumn();
    archHeaders = sheetArchivio.getRange(1, 1, 1, archLastCol).getValues()[0];

    // Costruisci righe da appendere
    var now = new Date();
    var rowsToAppend = [];
    for (var r = 1; r < dataConsegne.length; r++) {
      var rowC = dataConsegne[r];
      if (iLetto === -1 || String(rowC[iLetto] || '').trim() === '') continue;
      var obj = {};
      for (var j = 0; j < headersCons.length; j++) obj[String(headersCons[j]).trim()] = rowC[j];
      var row = [];
      for (var k = 0; k < archHeaders.length; k++) {
        var hdr = String(archHeaders[k]).trim();
        row.push(k === 0 ? now : (obj[hdr] !== undefined ? obj[hdr] : ''));
      }
      rowsToAppend.push(row);
    }
    if (rowsToAppend.length > 0) {
      sheetArchivio.getRange(sheetArchivio.getLastRow() + 1, 1, rowsToAppend.length, rowsToAppend[0].length).setValues(rowsToAppend);
      SpreadsheetApp.flush();
    }

    // Costruisci array pazienti per il backup doc
    var pazientiDoc = [];
    for (var rd = 1; rd < dataConsegne.length; rd++) {
      var rowD = dataConsegne[rd];
      if (iLetto === -1 || String(rowD[iLetto] || '').trim() === '') continue;
      var pazDoc = {};
      for (var jd = 0; jd < headersCons.length; jd++) {
        var campoD = String(headersCons[jd]).trim();
        var valD = rowD[jd];
        if (CAMPI_RICH_TEXT.indexOf(campoD) !== -1) valD = sanitizzaHtmlCampo(valD);
        pazDoc[campoD] = valD;
      }
      pazientiDoc.push(pazDoc);
    }
    pazientiDoc.sort(function(a, b) {
      var nA = parseInt(a.Letto, 10), nB = parseInt(b.Letto, 10);
      return (!isNaN(nA) && !isNaN(nB)) ? nA - nB : String(a.Letto).localeCompare(String(b.Letto));
    });

    var backupEseguito = creaBackupEmergenzaDoc(pazientiDoc);
    return { eseguito: true, backupEseguito: backupEseguito };

  } finally {
    props.deleteProperty('BACKUP_IN_CORSO');
  }
}
function getGiorniArchiviati() { const ss=SpreadsheetApp.getActiveSpreadsheet(); const sheetArchivio=ss.getSheetByName(NOME_ARCHIVIO); if(!sheetArchivio)return[]; const data=sheetArchivio.getRange(2,1,Math.max(1,sheetArchivio.getLastRow()-1),1).getValues(); const tz=Session.getScriptTimeZone(); const giorni={}; for(let i=0;i<data.length;i++){if(data[i][0]){let dStr=(data[i][0] instanceof Date)?Utilities.formatDate(data[i][0],tz,"yyyy-MM-dd"):String(data[i][0]); giorni[dStr]=true;}} return Object.keys(giorni); }
function getTimestampGiorno(dataStr) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetArchivio = ss.getSheetByName(NOME_ARCHIVIO);
  if (!sheetArchivio || sheetArchivio.getLastRow() <= 1) return [];
  var numRows = sheetArchivio.getLastRow() - 1;
  var data = sheetArchivio.getRange(2, 1, numRows, 1).getValues();
  var tz = Session.getScriptTimeZone();
  var seen = {}, risultati = [];
  for (var i = 0; i < data.length; i++) {
    var cellVal = data[i][0];
    if (!cellVal) continue;
    var dayStr = (cellVal instanceof Date) ? Utilities.formatDate(cellVal, tz, 'yyyy-MM-dd') : String(cellVal).substring(0, 10);
    if (dayStr !== dataStr) continue;
    var tsStr = (cellVal instanceof Date) ? Utilities.formatDate(cellVal, tz, 'yyyy-MM-dd HH:mm:ss') : String(cellVal);
    if (!seen[tsStr]) { seen[tsStr] = true; risultati.push(tsStr); }
  }
  risultati.sort();
  return risultati;
}
function getDatiArchivioGiorno(dataSelezionataStr) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetArchivio = ss.getSheetByName(NOME_ARCHIVIO);
  if (!sheetArchivio) return [];
  var data = sheetArchivio.getDataRange().getValues();
  var headers = data[0];
  var pazienti = [];
  var tz = Session.getScriptTimeZone();
  var isTimestamp = dataSelezionataStr.indexOf(' ') !== -1;
  var fmt = isTimestamp ? 'yyyy-MM-dd HH:mm:ss' : 'yyyy-MM-dd';
  for (var i = 1; i < data.length; i++) {
    var cellVal = data[i][0];
    var dStr = (cellVal instanceof Date)
      ? Utilities.formatDate(cellVal, tz, fmt)
      : (isTimestamp ? String(cellVal) : String(cellVal).substring(0, 10));
    if (dStr === dataSelezionataStr) {
      var p = {};
      for (var j = 1; j < headers.length; j++) {
        var val = data[i][j];
        if (val instanceof Date) val = Utilities.formatDate(val, tz, 'dd/MM/yyyy');
        if (CAMPI_RICH_TEXT.indexOf(headers[j]) !== -1) val = sanitizzaHtmlCampo(val);
        p[headers[j]] = val;
      }
      pazienti.push(p);
    }
  }
  return pazienti;
}


// ==========================================
// IMPOSTAZIONI ARCHIVIO
// ==========================================
function getGiorniArchivio() { var val=PropertiesService.getScriptProperties().getProperty('GIORNI_ARCHIVIO'); return val?parseInt(val,10):30; }
function salvaGiorniArchivio(giorni) { var n=parseInt(giorni,10); if(isNaN(n)||n<1)return{success:false,message:'Valore non valido.'}; PropertiesService.getScriptProperties().setProperty('GIORNI_ARCHIVIO',String(n)); return{success:true,giorni:n}; }
function pulisciArchivioVecchio() { var giorni=getGiorniArchivio(); var ss=SpreadsheetApp.getActiveSpreadsheet(); var sheetArchivio=ss.getSheetByName(NOME_ARCHIVIO); if(!sheetArchivio||sheetArchivio.getLastRow()<=1)return 0; var oggi=new Date(); oggi.setHours(0,0,0,0); var limiteMs=oggi.getTime()-(giorni*86400000); var data=sheetArchivio.getDataRange().getValues(); var righeEliminate=0; for(var i=data.length-1;i>=1;i--){var cellDate=data[i][0]; var d=(cellDate instanceof Date)?cellDate:new Date(String(cellDate)); if(!isNaN(d.getTime())&&d.getTime()<limiteMs){sheetArchivio.deleteRow(i+1); righeEliminate++;}} if(righeEliminate>0)SpreadsheetApp.flush(); return righeEliminate; }


// ==========================================
// BACKUP EMERGENZA SU GOOGLE DRIVE
// ==========================================

function _stripHtml(html) {
  if (!html) return '';
  return String(html)
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/p>/gi, '\n').replace(/<\/div>/gi, '\n').replace(/<\/li>/gi, '\n')
    .replace(/<li[^>]*>/gi, '• ')
    .replace(/<[^>]+>/g, '')
    .replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&').replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>').replace(/&quot;/g, '"')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

function _getOrCreateBackupFolder() {
  var consegneFolders = DriveApp.getFoldersByName('CONSEGNE APP');
  var consegneFolder = consegneFolders.hasNext()
    ? consegneFolders.next()
    : DriveApp.getRootFolder().createFolder('CONSEGNE APP');
  var backupFolders = consegneFolder.getFoldersByName('BACKUP EMERGENZA');
  return backupFolders.hasNext() ? backupFolders.next() : consegneFolder.createFolder('BACKUP EMERGENZA');
}

function _estraiDataDaNomeBackup(nome) {
  // Formato atteso: Consegne_yyyy-MM-dd_HH-mm
  var m = String(nome).match(/(\d{4})-(\d{2})-(\d{2})_(\d{2})-(\d{2})/);
  if (!m) return null;
  return new Date(parseInt(m[1],10), parseInt(m[2],10)-1, parseInt(m[3],10), parseInt(m[4],10), parseInt(m[5],10));
}

function pulisciBackupEmergenzaVecchi() {
  try {
    var consegneFolders = DriveApp.getFoldersByName('CONSEGNE APP');
    if (!consegneFolders.hasNext()) return 0;
    var backupFolders = consegneFolders.next().getFoldersByName('BACKUP EMERGENZA');
    if (!backupFolders.hasNext()) return 0;
    var folder = backupFolders.next();
    var limite = new Date().getTime() - (getGiorniArchivio() * 86400000);
    var files = folder.getFiles();
    var n = 0;
    while (files.hasNext()) {
      var f = files.next();
      var dataFile = _estraiDataDaNomeBackup(f.getName());
      var ts = dataFile ? dataFile.getTime() : f.getDateCreated().getTime();
      if (ts < limite) { f.setTrashed(true); n++; }
    }
    return n;
  } catch(e) { return 0; }
}

function creaBackupEmergenzaDoc(pazienti) {
  try {
    var backupFolder = _getOrCreateBackupFolder();
    var tz = Session.getScriptTimeZone();
    var now = new Date();
    var nomeFile = 'Consegne_' + Utilities.formatDate(now, tz, 'yyyy-MM-dd_HH-mm');
    var nomeReparto = ottieniNomeReparto();

    var doc = DocumentApp.create(nomeFile);
    var body = doc.getBody();

    // Layout orizzontale A4 (842 x 595 pt)
    var pageAttr = {};
    pageAttr[DocumentApp.Attribute.PAGE_WIDTH]    = 842;
    pageAttr[DocumentApp.Attribute.PAGE_HEIGHT]   = 595;
    pageAttr[DocumentApp.Attribute.MARGIN_TOP]    = 28;
    pageAttr[DocumentApp.Attribute.MARGIN_BOTTOM] = 28;
    pageAttr[DocumentApp.Attribute.MARGIN_LEFT]   = 28;
    pageAttr[DocumentApp.Attribute.MARGIN_RIGHT]  = 28;
    body.setAttributes(pageAttr);

    // Larghezza utile ≈ 842 - 56 = 786 pt
    // Proporzioni come visualizzazione alternativa: Info 12% | Diag+Piano+DaFare 18% | Diaria 58% | Terapia 12%
    var W = 786;
    var COL_W = [Math.round(W * 0.12), Math.round(W * 0.18), Math.round(W * 0.58), Math.round(W * 0.12)];

    // Intestazione
    var titlePar = body.appendParagraph(nomeReparto.toUpperCase() + ' — CONSEGNE DI EMERGENZA');
    titlePar.setHeading(DocumentApp.ParagraphHeading.HEADING1);
    titlePar.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    titlePar.editAsText().setFontSize(13);
    var datePar = body.appendParagraph('Backup del: ' + Utilities.formatDate(now, tz, 'dd/MM/yyyy HH:mm'));
    datePar.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    datePar.editAsText().setFontSize(8).setForegroundColor('#666666');

    pazienti.forEach(function(p) {
      if (!p.Letto) return;

      // ── Colonna 1: anagrafica ──────────────────────────────────────
      var sesso = p.Sesso === 'M' ? ' ♂' : (p.Sesso === 'F' ? ' ♀' : '');
      var nomePaz = p.Nome ? String(p.Nome).toUpperCase() : '(vuoto)';
      var tipologia = (p.TipologiaLetto || 'STANDARD').toUpperCase();
      var col1Lines = ['LETTO ' + p.Letto + sesso, '[' + tipologia + ']', '', nomePaz];
      if (p.DataNascita) col1Lines.push('Nato: ' + String(p.DataNascita));
      if (p.Eta)         col1Lines.push('Età: ' + p.Eta);
      if (p.DataRicovero) col1Lines.push('Ric.: ' + String(p.DataRicovero));
      if (p.CodiceSanitario) col1Lines.push('C.S.: ' + p.CodiceSanitario);
      var allergie = _stripHtml(p.Allergie);
      var ossigeno = _stripHtml(p.Ossigeno);
      if (allergie) col1Lines.push('\nALL.:\n' + allergie);
      if (ossigeno) col1Lines.push('\nO2:\n' + ossigeno);
      var col1 = col1Lines.join('\n');

      // ── Colonna 2: diagnosi + piano + da fare ─────────────────────
      var col2Parts = [];
      var diagnosi = _stripHtml(p.Diagnosi);
      if (diagnosi) col2Parts.push('DIAGNOSI:\n' + diagnosi);
      var piano = _stripHtml(p.PianoTerapeutico);
      if (piano) col2Parts.push('PIANO TERAPEUTICO:\n' + piano);
      var daFare = _stripHtml(p.DaFare);
      if (daFare) col2Parts.push('DA FARE:\n' + daFare);
      var col2 = col2Parts.join('\n\n') || '—';

      // ── Colonna 3: diaria ─────────────────────────────────────────
      var col3 = _stripHtml(p.Diaria) || '—';

      // ── Colonna 4: note e terapia ─────────────────────────────────
      var col4 = _stripHtml(p.NoteTerapia) || '—';

      // ── Tabella 4 colonne ─────────────────────────────────────────
      var tbl = body.appendTable([
        ['ANAGRAFICA', 'DIAGNOSI / PIANO / DA FARE', 'DIARIA ED EPICRISI', 'NOTE E TERAPIA'],
        [col1, col2, col3, col4]
      ]);
      tbl.setBorderWidth(1);

      // Larghezze colonne
      for (var ci = 0; ci < 4; ci++) {
        var wAttr = {};
        wAttr[DocumentApp.Attribute.WIDTH] = COL_W[ci];
        tbl.getCell(0, ci).setAttributes(wAttr);
        tbl.getCell(1, ci).setAttributes(wAttr);
      }

      // Header row: bold + sfondo grigio
      for (var c = 0; c < 4; c++) {
        var hCell = tbl.getCell(0, c);
        hCell.editAsText().setBold(true).setFontSize(8).setForegroundColor('#333333');
        hCell.setBackgroundColor('#e8e6e1');
        hCell.setPaddingTop(3); hCell.setPaddingBottom(3);
        hCell.setPaddingLeft(4); hCell.setPaddingRight(4);
      }

      // Content row: dimensione font + bold etichette
      var labelRegex = /(LETTO \d+[^♂♀\n]*|DIAGNOSI:|PIANO TERAPEUTICO:|DA FARE:|ALL\.:|O2:)/g;
      for (var c2 = 0; c2 < 4; c2++) {
        var cCell = tbl.getCell(1, c2);
        cCell.editAsText().setFontSize(9);
        cCell.setPaddingTop(3); cCell.setPaddingBottom(3);
        cCell.setPaddingLeft(4); cCell.setPaddingRight(4);
        var cellTxt = cCell.getText();
        var m;
        labelRegex.lastIndex = 0;
        while ((m = labelRegex.exec(cellTxt)) !== null) {
          cCell.editAsText().setBold(m.index, m.index + m[0].length - 1, true);
        }
      }
    });

    doc.saveAndClose();
    DriveApp.getFileById(doc.getId()).moveTo(backupFolder);
    return true;
  } catch(e) {
    Logger.log('Errore creaBackupEmergenzaDoc: ' + e.toString());
    return false;
  }
}


// ==========================================
// GESTIONE TIPOLOGIE
// ==========================================

function getColoriTipologie() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(NOME_TIPOLOGIE);
    if (sheet && sheet.getLastRow() > 1) {
      var data = sheet.getDataRange().getValues();
      var mappa = {};
      for (var i = 1; i < data.length; i++) {
        var nome = String(data[i][0]).trim().toUpperCase();
        var colore = data[i].length > 1 ? String(data[i][1]).trim() : '';
        if (nome) mappa[nome] = colore || null;
      }
      return mappa;
    }
    var v = PropertiesService.getScriptProperties().getProperty('TIPOLOGIE_COLORI');
    return v ? JSON.parse(v) : {};
  } catch(e) { return {}; }
}


function salvaColoriTipologie(mappa) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(NOME_TIPOLOGIE);
    if (!sheet) {
      sheet = ss.insertSheet(NOME_TIPOLOGIE);
    }
    sheet.clearContents();
    sheet.getRange(1, 1, 1, 2).setValues([['Nome', 'Colore']]);
    var nomi = Object.keys(mappa || {}).filter(function(k) { return k; }).sort();
    if (nomi.length > 0) {
      var rows = nomi.map(function(nome) { return [nome, mappa[nome] || '']; });
      sheet.getRange(2, 1, rows.length, 2).setValues(rows);
    }
    SpreadsheetApp.flush();
    try {
      PropertiesService.getScriptProperties().setProperty('TIPOLOGIE_COLORI', JSON.stringify(mappa || {}));
    } catch(e2) {}
  } catch(e) {
    PropertiesService.getScriptProperties().setProperty('TIPOLOGIE_COLORI', JSON.stringify(mappa || {}));
  }
}


function getTipologieConfigurate() {
  var mappa = getColoriTipologie();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(NOME_FOGLIO);
  var result = {};
  Object.keys(mappa).forEach(function(k) {
    if (k) result[k] = mappa[k] || null;
  });
  if (sheet) {
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var colIdx = headers.indexOf('TipologiaLetto');
    if (colIdx !== -1) {
      for (var i = 1; i < data.length; i++) {
        var t = String(data[i][colIdx]).trim();
        if (t && !(t in result)) result[t] = null;
      }
    }
  }
  return Object.keys(result).sort().map(function(k) {
    return { nome: k, colore: result[k] };
  });
}


function getDatiLettiConTipologia() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(NOME_FOGLIO);
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var result = [];
  for (var i = 1; i < data.length; i++) {
    var letto = String(data[i][0]).trim();
    if (letto) {
      var obj = {};
      for (var j = 0; j < headers.length; j++) obj[headers[j]] = data[i][j];
      result.push({ letto: letto, nome: String(obj['Nome'] || '').trim(), tipologia: String(obj['TipologiaLetto'] || '').trim() });
    }
  }
  result.sort(function(a,b) {
    var na=parseInt(a.letto,10), nb=parseInt(b.letto,10);
    if (!isNaN(na)&&!isNaN(nb)) return na-nb;
    return String(a.letto).localeCompare(String(b.letto));
  });
  return result;
}


function salvaTipologieBatch(modifiche) {
  if (!modifiche || !modifiche.length) return { success: true };
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(NOME_FOGLIO);
  var mappa = getColoriTipologie();

  modifiche.forEach(function(m) {
    var nomeOld = (m.nomeOld || '').trim().toUpperCase();
    var nomeNew = (m.nomeNew || '').trim().toUpperCase();
    if (!nomeNew) return;
    if (nomeOld && nomeOld !== nomeNew) delete mappa[nomeOld];
    mappa[nomeNew] = m.colore || null;
  });

  if (sheet) {
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var colIdx = headers.indexOf('TipologiaLetto');
    if (colIdx !== -1) {
      modifiche.forEach(function(m) {
        var nomeOld = (m.nomeOld || '').trim().toUpperCase();
        var nomeNew = (m.nomeNew || '').trim().toUpperCase();
        if (!nomeNew || !nomeOld || nomeOld === nomeNew) return;
        for (var i = 1; i < data.length; i++) {
          if (String(data[i][colIdx]).trim() === nomeOld) {
            sheet.getRange(i+1, colIdx+1).setValue(nomeNew);
            data[i][colIdx] = nomeNew;
          }
        }
      });
      SpreadsheetApp.flush();
    }
  }

  salvaColoriTipologie(mappa);
  return { success: true };
}


function eliminaTipologiaConfigurata(nome, force) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(NOME_FOGLIO);
  var count = 0;
  if (sheet) {
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var colIdx = headers.indexOf('TipologiaLetto');
    if (colIdx !== -1) {
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][colIdx]).trim() === nome) count++;
      }
      if (count > 0 && !force) {
        return { success: false, count: count };
      }
      if (force && count > 0) {
        for (var i = 1; i < data.length; i++) {
          if (String(data[i][colIdx]).trim() === nome) {
            sheet.getRange(i+1, colIdx+1).setValue('');
          }
        }
        SpreadsheetApp.flush();
      }
    }
  }
  var mappa = getColoriTipologie();
  delete mappa[nome];
  salvaColoriTipologie(mappa);
  return { success: true };
}


function cambiaTipologiaALetto(letto, nuovaTipologia) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(NOME_FOGLIO);
  if (!sheet) return { success: false, message: 'Foglio non trovato.' };
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var colIdx = headers.indexOf('TipologiaLetto');
  if (colIdx === -1) return { success: false, message: 'Colonna TipologiaLetto non trovata.' };
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(letto).trim()) {
      sheet.getRange(i+1, colIdx+1).setValue(nuovaTipologia || '');
      SpreadsheetApp.flush();
      return { success: true };
    }
  }
  return { success: false, message: 'Letto non trovato.' };
}


// ==========================================
// SETUP PRIMO AVVIO
// ==========================================
//
// ISTRUZIONI:
// 1. Apri appsscript.json nell'editor GAS (Impostazioni progetto → mostra manifest)
// 2. Assicurati che contenga gli oauthScopes corretti (la funzione lo verifica e li logga se mancano)
// 3. Seleziona "primoAvvio" nel menu funzioni e clicca ▶ Esegui
// 4. Accetta il dialogo dei permessi quando appare
// 5. Controlla il log — tutto deve risultare OK
//
function primoAvvio() {
  var log = [];
  var errori = 0;

  // ── Step 1: verifica scope appsscript.json ──────────────────────────
  var SCOPE_RICHIESTI = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/documents',
    'https://www.googleapis.com/auth/script.external_request'
  ];
  try {
    var token = ScriptApp.getOAuthToken();
    // Decode JWT payload per leggere gli scope autorizzati
    var payload = token.split('.')[1];
    // padding base64
    payload = payload.replace(/-/g, '+').replace(/_/g, '/');
    while (payload.length % 4) payload += '=';
    var decoded = JSON.parse(Utilities.newBlob(Utilities.base64Decode(payload)).getDataAsString());
    var scopeAutorizzati = (decoded.scope || '').split(' ');
    var scopeMancanti = SCOPE_RICHIESTI.filter(function(s) { return scopeAutorizzati.indexOf(s) === -1; });
    if (scopeMancanti.length === 0) {
      log.push('[OK] appsscript.json: tutti gli scope autorizzati');
    } else {
      errori++;
      log.push('[!!] SCOPE MANCANTI in appsscript.json: ' + scopeMancanti.join(', '));
      log.push('     Incolla questo in appsscript.json e ri-esegui:\n' +
        JSON.stringify({ oauthScopes: SCOPE_RICHIESTI }, null, 2));
    }
  } catch(e) {
    log.push('[??] Verifica scope non riuscita (non bloccante): ' + e.message);
  }

  // ── Step 2: autorizza DriveApp ──────────────────────────────────────
  try {
    DriveApp.getRootFolder();
    log.push('[OK] DriveApp autorizzato');
  } catch(e) {
    errori++;
    log.push('[ERR] DriveApp: ' + e.message);
  }

  // ── Step 3: autorizza DocumentApp ──────────────────────────────────
  try {
    var doc = DocumentApp.create('_test_autorizzazione_consegne_');
    DriveApp.getFileById(doc.getId()).setTrashed(true);
    log.push('[OK] DocumentApp autorizzato (file di test eliminato)');
  } catch(e) {
    errori++;
    log.push('[ERR] DocumentApp: ' + e.message);
  }

  // ── Step 4: crea cartelle backup su Drive ───────────────────────────
  try {
    var folder = _getOrCreateBackupFolder();
    log.push('[OK] Cartella "CONSEGNE APP/BACKUP EMERGENZA" pronta (id: ' + folder.getId() + ')');
  } catch(e) {
    errori++;
    log.push('[ERR] Creazione cartelle: ' + e.message);
  }

  // ── Step 5: test backup con dati reali ─────────────────────────────
  try {
    var pazienti = getDatiPazienti();
    if (pazienti.length === 0) {
      log.push('[--] Nessun paziente trovato — test backup saltato (aggiungi almeno un letto e ri-esegui)');
    } else {
      var ok = creaBackupEmergenzaDoc(pazienti);
      if (ok) {
        log.push('[OK] Backup di test creato in "CONSEGNE APP/BACKUP EMERGENZA"');
      } else {
        errori++;
        log.push('[ERR] Backup fallito — controlla i log sopra per i dettagli');
      }
    }
  } catch(e) {
    errori++;
    log.push('[ERR] Test backup: ' + e.message);
  }

  // ── Step 6: reset timer archivio ────────────────────────────────────
  try {
    PropertiesService.getScriptProperties().deleteProperty('ULTIMO_ARCHIVIO_TS');
    log.push('[OK] Timer archivio resettato — il prossimo caricamento eseguirà archivio + backup');
  } catch(e) {
    log.push('[??] Reset timer: ' + e.message);
  }

  // ── Riepilogo ───────────────────────────────────────────────────────
  log.push('');
  log.push(errori === 0
    ? '✓ SETUP COMPLETATO — ricarica la pagina web per verificare'
    : '✗ SETUP CON ' + errori + ' ERRORE/I — risolvi i problemi segnalati e ri-esegui');

  Logger.log(log.join('\n'));
}

// Test diretto del backup: utile per verificare layout e funzionamento dopo modifiche.
function testBackup() {
  var pazienti = getDatiPazienti();
  var result = creaBackupEmergenzaDoc(pazienti);
  Logger.log('Risultato backup: ' + result);
}

// Resetta il timer archivio: forza l'esecuzione di archivio + backup al prossimo caricamento.
function resetArchivioTs() {
  PropertiesService.getScriptProperties().deleteProperty('ULTIMO_ARCHIVIO_TS');
  Logger.log('Timer resettato. Ricarica la pagina per eseguire il backup.');
}


// ==========================================
// TIPOLOGIE LETTI ATTIVE (per modal stampa)
// ==========================================
function getTipologieLettiBed() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(NOME_FOGLIO);
  if (!sheet) return ['STANDARD'];
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var iLetto = headers.indexOf('Letto');
  var iTip   = headers.indexOf('TipologiaLetto');
  var found  = {};
  for (var i = 1; i < data.length; i++) {
    if (iLetto === -1 || String(data[i][iLetto] || '').trim() === '') continue;
    var t = iTip !== -1 ? String(data[i][iTip] || '').trim().toUpperCase() : '';
    found[t || 'STANDARD'] = true;
  }
  var result = Object.keys(found).sort();
  return result.length > 0 ? result : ['STANDARD'];
}


// ==========================================
// LINK UTILI
// ==========================================
const NOME_LINK_UTILI = 'Link Utili';

function getLinkUtili() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(NOME_LINK_UTILI);
  if (!sheet) {
    sheet = ss.insertSheet(NOME_LINK_UTILI);
    sheet.getRange(1, 1, 1, 2).setValues([['Nome', 'Collegamento']]);
    SpreadsheetApp.flush();
    return [];
  }
  if (sheet.getLastRow() <= 1) return [];
  var numRows = sheet.getLastRow() - 1;
  var data = sheet.getRange(2, 1, numRows, 2).getValues();
  var result = [];
  for (var i = 0; i < data.length; i++) {
    var nome = String(data[i][0] || '').trim();
    var url  = String(data[i][1] || '').trim();
    if (nome || url) result.push({ id: i, nome: nome, url: url });
  }
  return result;
}

function aggiungiLinkUtile(nome, url) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(NOME_LINK_UTILI);
  if (!sheet) {
    sheet = ss.insertSheet(NOME_LINK_UTILI);
    sheet.getRange(1, 1, 1, 2).setValues([['Nome', 'Collegamento']]);
  }
  sheet.appendRow([nome || '', url || '']);
  SpreadsheetApp.flush();
  return { success: true };
}

function modificaLinkUtile(indice, nome, url) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(NOME_LINK_UTILI);
  if (!sheet || indice < 0) return { success: false };
  var riga = indice + 2;
  if (riga > sheet.getLastRow()) return { success: false };
  sheet.getRange(riga, 1, 1, 2).setValues([[nome || '', url || '']]);
  SpreadsheetApp.flush();
  return { success: true };
}

function eliminaLinkUtile(indice) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(NOME_LINK_UTILI);
  if (!sheet || indice < 0) return { success: false };
  var riga = indice + 2;
  if (riga > sheet.getLastRow()) return { success: false };
  sheet.deleteRow(riga);
  SpreadsheetApp.flush();
  return { success: true };
}
