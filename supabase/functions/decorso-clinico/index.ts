// decorso-clinico v1 — il fascicolo per il decorso clinico si prepara SUL
// SERVER: il browser manda solo ciò che non sta nel database (referti e
// diaria incollati, impostazioni) e riceve la risposta. Il codice che la
// produce resta qui.
//
// FASE ATTUALE (prova delle schermate): la funzione raccoglie, controlla e
// rimanda indietro ciò che ha ricevuto, perche' si veda che cosa arriva. Il
// motore che scrive il decorso prendera' il posto del passo «fascicolo».
// Niente viene salvato sul server e nessun contenuto finisce nei registri:
// solo nomi dei passi, esiti e dimensioni.
//
// Richiesta: POST JSON { letto, referti, diaria, impostazioni: {righe, dettaglio} }
// Risposta:  un flusso di righe JSON (application/x-ndjson), una per evento,
//            cosi' le schermate mostrano i passi mentre avvengono:
//              {"passo":"consegne","stato":"inizio"}
//              {"passo":"consegne","stato":"ok","nota":"12 campi"}
//              …  stato: inizio | ok | salto | errore
//              {"passo":"fine","stato":"ok","fase":"prova","fascicolo":{…},"riepilogo":{…}}
//            Se un passo fallisce arriva {"passo":…,"stato":"errore","errore":…}
//            e poi {"passo":"fine","stato":"errore"}.
//
// CHI PUO' CHIAMARE: solo un utente autorizzato (sessione Supabase di un
// indirizzo in utenti_autorizzati, verificata con il SUO Bearer tramite
// is_autorizzato()). La chiave pubblica da sola riceve 403 prima di qualunque
// lettura. Qui non c'e' interruttore: il controllo e' sempre acceso.
//
// DA DOVE VENGONO I DATI: le consegne del letto e gli esami di laboratorio
// si leggono dal database (chiave di servizio) — sono quelli SALVATI, per
// questo l'app salva la scheda prima di chiamare. I referti e la diaria li
// incolla chi usa l'app.

import { createClient } from 'npm:@supabase/supabase-js@2';

const CORS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'authorization, x-client-info, apikey, content-type',
  'Access-Control-Allow-Methods': 'POST, OPTIONS',
};
const json = (body: unknown, status = 200) =>
  new Response(JSON.stringify(body), { status, headers: { ...CORS, 'Content-Type': 'application/json' } });

const SB_URL = Deno.env.get('SUPABASE_URL')!;
const sb = createClient(SB_URL, Deno.env.get('SUPABASE_SERVICE_ROLE_KEY')!);

// Limiti di ciò che il browser può mandare (caratteri).
const MAX_REFERTI = 200_000;
const MAX_DIARIA = 300_000;
const RIGHE_MIN = 7, RIGHE_MAX = 50, RIGHE_DEFAULT = 15;
const DETTAGLIO_MIN = 1, DETTAGLIO_MAX = 4, DETTAGLIO_DEFAULT = 2;
const DETTAGLI = ['molto generico', 'poco dettagliato', 'dettagliato', 'estremamente dettagliato'];

// Il chiamante e' un utente loggato e in lista? Si chiede al database con il
// SUO Bearer: la chiave pubblica (ruolo anon) e gli account fuori lista danno
// false, come pure qualunque errore.
async function chiamanteAutorizzato(req: Request): Promise<boolean> {
  const auth = req.headers.get('Authorization') || '';
  if (!/^Bearer\s+\S+/i.test(auth)) return false;
  const apikey = Deno.env.get('SUPABASE_ANON_KEY') || req.headers.get('apikey') || '';
  if (!apikey) return false;
  try {
    const c = createClient(SB_URL, apikey, {
      global: { headers: { Authorization: auth } },
      auth: { persistSession: false, autoRefreshToken: false },
    });
    const { data, error } = await c.rpc('is_autorizzato');
    return !error && data === true;
  } catch {
    return false;
  }
}

// ── HTML dei campi → testo ──────────────────────────────────────────
// I campi della scheda sono HTML gia' passato dal filtro dell'app (elenco
// chiuso di tag: niente script, niente attivo). Qui diventano testo con le
// righe al posto giusto. Alcuni blocchi dentro i campi sono stato DERIVATO
// (il riquadro del laboratorio e' una copia degli esami, che arrivano dalla
// loro tabella): si tolgono; il riquadro della terapia importata si separa
// dal resto delle note.
const VUOTI = new Set(['br', 'img', 'hr', 'input', 'meta', 'link', 'wbr', 'col', 'area', 'base', 'source', 'track', 'embed', 'param']);
const ENTITA: Record<string, string> = {
  nbsp: ' ', amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", euro: '€', deg: '°', micro: 'µ',
  agrave: 'à', egrave: 'è', eacute: 'é', igrave: 'ì', ograve: 'ò', ugrave: 'ù', Agrave: 'À', Egrave: 'È', Eacute: 'É',
  laquo: '«', raquo: '»', ndash: '–', mdash: '—', hellip: '…', middot: '·', bull: '•', plusmn: '±', times: '×', frac12: '½',
};
function decodificaEntita(s: string): string {
  return s.replace(/&(#x[0-9a-fA-F]+|#\d+|[a-zA-Z][a-zA-Z0-9]*);/g, (m, c) => {
    if (c[0] === '#') {
      const n = c[1] === 'x' || c[1] === 'X' ? parseInt(c.slice(2), 16) : parseInt(c.slice(1), 10);
      return (n > 0 && n < 0x110000) ? String.fromCodePoint(n) : m;
    }
    return c in ENTITA ? ENTITA[c] : m;
  });
}

// Separa dall'HTML i sottoalberi la cui classe corrisponde a «classe».
// Restituisce il resto e i blocchi tolti (HTML, nell'ordine in cui stavano).
function separaBlocchi(html: string, classe: RegExp): { resto: string; blocchi: string[] } {
  const blocchi: string[] = [];
  let resto = '';
  const re = /<\/?([a-zA-Z][a-zA-Z0-9]*)\b([^>]*)>/g;
  let i = 0, profondita = 0, inizio = -1;
  let m: RegExpExecArray | null;
  while ((m = re.exec(html))) {
    const tag = m[1].toLowerCase();
    const chiude = m[0][1] === '/';
    const vuoto = VUOTI.has(tag) || /\/\s*>$/.test(m[0]);
    if (profondita === 0) {
      resto += html.slice(i, m.index);
      const cls = /\bclass\s*=\s*(?:"([^"]*)"|'([^']*)')/i.exec(m[2]);
      const nomi = cls ? (cls[1] ?? cls[2] ?? '') : '';
      if (!chiude && !vuoto && nomi && classe.test(nomi)) { profondita = 1; inizio = m.index; }
      else resto += m[0];
    } else if (!chiude && !vuoto) {
      profondita++;
    } else if (chiude) {
      profondita--;
      if (profondita === 0) blocchi.push(html.slice(inizio, m.index + m[0].length));
    }
    i = m.index + m[0].length;
  }
  if (profondita === 0) resto += html.slice(i); else blocchi.push(html.slice(inizio));
  return { resto, blocchi };
}

function testoDaHtml(html: string): string {
  let s = String(html || '');
  s = s.replace(/<(script|style)\b[^>]*>[\s\S]*?<\/\1>/gi, ' ');
  s = s.replace(/<br\s*\/?>/gi, '\n');
  // la chiusura di un blocco manda a capo UNA volta anche quando piu' blocchi
  // si chiudono insieme (</div></div>): un segnaposto, poi le sequenze
  // diventano un solo a capo. Una riga vuota vera (<div><br></div>) resta.
  s = s.replace(/<\/(p|div|li|tr|h[1-6]|blockquote|pre|section|article|ul|ol|table|dd|dt)>/gi, '\u0001');
  s = s.replace(/<li\b[^>]*>/gi, '• ');
  s = s.replace(/<\/t[dh]>/gi, ' | ');
  s = s.replace(/<[^>]+>/g, '');
  s = s.replace(/\u0001+/g, '\n');
  s = decodificaEntita(s).replace(/ /g, ' ');
  s = s.split('\n').map((r) => r.replace(/[ \t]+/g, ' ').replace(/(\s*\|)+\s*$/, '').trim()).join('\n');
  return s.replace(/\n{3,}/g, '\n\n').trim();
}
const righe = (t: string) => (t ? t.split('\n').length : 0);

// ── date ────────────────────────────────────────────────────────────
const MESI: Record<string, number> = {
  gennaio: 1, gen: 1, febbraio: 2, feb: 2, marzo: 3, mar: 3, aprile: 4, apr: 4, maggio: 5, mag: 5, giugno: 6, giu: 6,
  luglio: 7, lug: 7, agosto: 8, ago: 8, settembre: 9, set: 9, sett: 9, ottobre: 10, ott: 10, novembre: 11, nov: 11, dicembre: 12, dic: 12,
};
const iso = (a: number, m: number, g: number) => a + '-' + String(m).padStart(2, '0') + '-' + String(g).padStart(2, '0');
function dataValida(a: number, m: number, g: number): boolean {
  if (m < 1 || m > 12 || g < 1 || g > 31 || a < 1900 || a > 2100) return false;
  const d = new Date(Date.UTC(a, m - 1, g));
  return d.getUTCMonth() === m - 1 && d.getUTCDate() === g;
}
// Prima data riconoscibile in una riga: «27/10/26», «27/10/2026», «27-10-2026», «27.10.2026», «27 ottobre 2026», «27 ott 26».
function dataNellaRiga(riga: string): { iso: string; testo: string } | null {
  const n = /(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2,4})(?!\d)/.exec(riga);
  const p = /(\d{1,2})\s+([a-zA-Z]{3,9})\.?\s+(\d{2,4})(?!\d)/.exec(riga);
  const candidati: Array<{ i: number; g: number; m: number; a: number; testo: string }> = [];
  if (n) candidati.push({ i: n.index, g: +n[1], m: +n[2], a: +n[3], testo: n[0] });
  if (p && MESI[p[2].toLowerCase()]) candidati.push({ i: p.index, g: +p[1], m: MESI[p[2].toLowerCase()], a: +p[3], testo: p[0] });
  candidati.sort((x, y) => x.i - y.i);
  for (const c of candidati) {
    const a = c.a < 100 ? 2000 + c.a : c.a;
    if (dataValida(a, c.m, c.g)) return { iso: iso(a, c.m, c.g), testo: c.testo };
  }
  return null;
}
// «gg/mm/aaaa hh:mm» (le chiavi degli esami) → chiave ordinabile «aaaammgghhmm»
function chiaveData(s: string): string {
  const m = /(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2}))?/.exec(s);
  if (!m) return '';
  return m[3] + m[2].padStart(2, '0') + m[1].padStart(2, '0') + (m[4] || '00').padStart(2, '0') + (m[5] || '00');
}

// ── referti: un riquadro solo, separati dalle righe con titolo e data ──
// Una riga «titolo» e' corta, contiene una data e non finisce con una virgola
// o un punto e virgola (una frase spezzata). Tutto cio' che sta sotto, fino
// al titolo successivo, e' il testo del referto. Il testo prima del primo
// titolo e' un referto senza titolo.
type Referto = { titolo: string; data: string | null; dataTesto: string | null; testo: string; righe: number; ordine: number };
function separaReferti(testo: string): { referti: Referto[]; senzaData: number } {
  const linee = String(testo || '').replace(/\r\n?/g, '\n').split('\n');
  const referti: Referto[] = [];
  let corrente: Referto | null = null;
  const chiudi = () => {
    if (!corrente) return;
    corrente.testo = corrente.testo.replace(/\n{3,}/g, '\n\n').trim();
    corrente.righe = righe(corrente.testo);
    if (corrente.titolo || corrente.testo) referti.push(corrente);
    corrente = null;
  };
  linee.forEach((l) => {
    const r = l.trim();
    const d = r.length >= 3 && r.length <= 100 ? dataNellaRiga(r) : null;
    if (d && !/[,;]$/.test(r)) {
      chiudi();
      corrente = { titolo: r, data: d.iso, dataTesto: d.testo, testo: '', righe: 0, ordine: referti.length };
      return;
    }
    if (!corrente) corrente = { titolo: '', data: null, dataTesto: null, testo: '', righe: 0, ordine: 0 };
    corrente.testo += (corrente.testo ? '\n' : '') + l.replace(/[ \t]+$/, '');
  });
  chiudi();
  referti.forEach((x, i) => { x.ordine = i; });
  // cronologico; i referti senza data in coda, nell'ordine in cui erano
  referti.sort((a, b) => {
    if (a.data && b.data) return a.data < b.data ? -1 : a.data > b.data ? 1 : a.ordine - b.ordine;
    if (a.data) return -1;
    if (b.data) return 1;
    return a.ordine - b.ordine;
  });
  return { referti, senzaData: referti.filter((x) => !x.data).length };
}

// ── esami di laboratorio: dal documento salvato a un elenco ordinato ──
type Esame = { nome: string; gruppo: string; um: string; range: string; valori: Array<{ data: string; valore: string }> };
function esamiDaDocumento(doc: any): { esami: Esame[]; dal: string | null; al: string | null; nValori: number } {
  const tutti = (doc && doc.esami) || {};
  const ordine: string[] = Array.isArray(doc?.ord) && doc.ord.length ? doc.ord : Object.keys(tutti);
  const esami: Esame[] = [];
  let min = '', max = '', nValori = 0;
  ordine.forEach((k) => {
    const e = tutti[k];
    if (!e) return;
    const valori = Object.keys(e.v || {})
      .map((d) => ({ chiave: chiaveData(d), data: d, valore: String(e.v[d] ?? '') }))
      .sort((a, b) => a.chiave < b.chiave ? -1 : a.chiave > b.chiave ? 1 : 0);
    valori.forEach((v) => { if (v.chiave) { if (!min || v.chiave < min) min = v.chiave; if (v.chiave > max) max = v.chiave; } });
    nValori += valori.length;
    esami.push({ nome: String(e.n || k), gruppo: String(e.g || ''), um: String(e.um || ''), range: String(e.range || ''), valori: valori.map((v) => ({ data: v.data, valore: v.valore })) });
  });
  const giorno = (c: string) => c ? c.slice(6, 8) + '/' + c.slice(4, 6) + '/' + c.slice(0, 4) : null;
  return { esami, dal: giorno(min), al: giorno(max), nValori };
}

// ── la richiesta ────────────────────────────────────────────────────
type Richiesta = { letto: string; referti: string; diaria: string; impostazioni: { righe: number; dettaglio: number; dettaglioTesto: string } };
function leggiRichiesta(body: Record<string, unknown>): { ok: true; r: Richiesta } | { ok: false; errore: string } {
  const letto = String(body.letto ?? '').trim();
  if (!letto || letto.length > 20) return { ok: false, errore: 'letto mancante' };
  const referti = typeof body.referti === 'string' ? body.referti : '';
  const diaria = typeof body.diaria === 'string' ? body.diaria : '';
  if (referti.length > MAX_REFERTI) return { ok: false, errore: 'referti troppo lunghi (' + referti.length + ' caratteri, massimo ' + MAX_REFERTI + ')' };
  if (diaria.length > MAX_DIARIA) return { ok: false, errore: 'diaria troppo lunga (' + diaria.length + ' caratteri, massimo ' + MAX_DIARIA + ')' };
  const imp = (body.impostazioni && typeof body.impostazioni === 'object') ? body.impostazioni as Record<string, unknown> : {};
  const intero = (v: unknown, min: number, max: number, def: number) => {
    const n = Math.round(Number(v));
    if (!Number.isFinite(n)) return def;
    return Math.min(max, Math.max(min, n));
  };
  const righe_ = intero(imp.righe, RIGHE_MIN, RIGHE_MAX, RIGHE_DEFAULT);
  const dettaglio = intero(imp.dettaglio, DETTAGLIO_MIN, DETTAGLIO_MAX, DETTAGLIO_DEFAULT);
  return { ok: true, r: { letto, referti, diaria, impostazioni: { righe: righe_, dettaglio, dettaglioTesto: DETTAGLI[dettaglio - 1] } } };
}

Deno.serve(async (req) => {
  if (req.method === 'OPTIONS') return new Response('ok', { headers: CORS });
  if (req.method !== 'POST') return json({ errore: 'solo POST' }, 405);
  if (!(await chiamanteAutorizzato(req))) {
    return json({ errore: 'Operazione riservata agli utenti del reparto: aggiorna la pagina (F5), accedi con l\'account autorizzato e riprova.', non_autorizzato: true }, 403);
  }
  let body: Record<string, unknown>;
  try { body = await req.json(); } catch { return json({ errore: 'JSON non valido' }, 400); }
  const letta = leggiRichiesta(body);
  if (!letta.ok) return json({ errore: letta.errore }, 400);
  const r = letta.r;

  const enc = new TextEncoder();
  const stream = new ReadableStream({
    async start(controller) {
      const manda = (ev: Record<string, unknown>) => controller.enqueue(enc.encode(JSON.stringify(ev) + '\n'));
      const registra = (passo: string, stato: string, byte?: number) => console.log(JSON.stringify({ ev: 'decorso', passo, stato, ...(byte !== undefined ? { byte } : {}) }));
      const inizio = Date.now();
      let inCorso = 'ricezione';
      // Esegue un passo: annuncia l'inizio, poi l'esito. Un'eccezione ferma tutto.
      const passo = async <T>(nome: string, fn: () => Promise<{ nota?: string; salto?: string; valore: T }>): Promise<T> => {
        inCorso = nome;
        manda({ passo: nome, stato: 'inizio' });
        const e = await fn();
        manda({ passo: nome, stato: e.salto ? 'salto' : 'ok', nota: e.salto || e.nota || '' });
        registra(nome, e.salto ? 'salto' : 'ok');
        return e.valore;
      };
      try {
        await passo('ricezione', async () => ({
          nota: 'referti ' + r.referti.length + ' caratteri, diaria ' + r.diaria.length + ' caratteri, ' + r.impostazioni.righe + ' righe, ' + r.impostazioni.dettaglioTesto,
          valore: null,
        }));

        // consegne del letto (salvate)
        const scheda = await passo('consegne', async () => {
          const { data, error } = await sb.from('consegne').select('*').eq('letto', r.letto).maybeSingle();
          if (error) throw new Error('lettura delle consegne: ' + error.message);
          if (!data) throw new Error('il letto «' + r.letto + '» non esiste nelle consegne');
          const t = (campo: string) => testoDaHtml(String(data[campo] ?? ''));
          const esamiColturali = separaBlocchi(String(data.esami_colturali ?? ''), /\blab-box\b/);   // il riquadro del laboratorio e' una copia degli esami
          const campi = {
            diagnosi: t('diagnosi'),
            allergie: t('allergie'),
            diaria: t('diaria'),
            problemi_attivi: t('piano_terapeutico'),
            da_fare: t('da_fare'),
            esami_colturali_e_scale: testoDaHtml(esamiColturali.resto),
            ossigeno: t('ossigeno'),
            vitto: t('vitto'),
            dimissibile: t('dimissibile'),
          };
          const pieni = Object.keys(campi).filter((k) => (campi as Record<string, string>)[k]).length;
          return {
            nota: pieni + ' campi con contenuto',
            valore: {
              identita: { letto: r.letto, nome: t('nome'), data_nascita: t('data_nascita'), eta: t('eta'), sesso: t('sesso'), codice_sanitario: t('codice_sanitario') },
              ricovero: { data_ricovero: t('data_ricovero'), tipologia_letto: t('tipologia_letto'), ultimo_aggiornamento: t('ultimo_aggiornamento'), salvata_il: data.updated_at || null },
              campi,
              note_terapia_html: String(data.note_terapia ?? ''),
            },
          };
        });

        // terapia: il riquadro importato da TrakCare (con la sua data) e le note scritte a mano
        const terapia = await passo('terapia', async () => {
          const sep = separaBlocchi(scheda.note_terapia_html, /\bterapia-box\b/);
          const box = sep.blocchi[0] || '';
          const ts = /\bdata-ts\s*=\s*["'](\d{10,16})["']/.exec(box);
          const importata = box ? testoDaHtml(box) : '';
          const note = testoDaHtml(sep.resto);
          const quando = ts ? new Date(Number(ts[1])).toISOString() : null;
          const v = { importata_da_trakcare: !!box, importata_il: quando, testo: importata, note: note };
          if (!importata && !note) return { salto: 'nessuna terapia nelle consegne', valore: v };
          return { nota: box ? ('importata da TrakCare' + (quando ? ' il ' + quando.slice(8, 10) + '/' + quando.slice(5, 7) + '/' + quando.slice(0, 4) : '')) : 'scritta a mano', valore: v };
        });
        delete (scheda as Record<string, unknown>).note_terapia_html;

        // esami di laboratorio (tabella lab_esami, documento completo)
        const laboratorio = await passo('laboratorio', async () => {
          const { data, error } = await sb.from('lab_esami').select('dati,n_esami,ultimo_esame,paziente').eq('letto', r.letto).maybeSingle();
          if (error) throw new Error('lettura degli esami: ' + error.message);
          if (!data || !data.dati) return { salto: 'nessun esame salvato per questo letto', valore: { presenti: false, esami: [], n_esami: 0, n_valori: 0, dal: null, al: null, paziente: '' } };
          const e = esamiDaDocumento(data.dati);
          return {
            nota: e.esami.length + ' esami, ' + e.nValori + ' valori' + (e.dal ? ', dal ' + e.dal + ' al ' + e.al : ''),
            valore: { presenti: true, esami: e.esami, n_esami: e.esami.length, n_valori: e.nValori, dal: e.dal, al: e.al, paziente: String(data.paziente || '') },
          };
        });

        // referti e consulenze incollati
        const referti = await passo('referti', async () => {
          if (!r.referti.trim()) return { salto: 'nessun referto incollato', valore: { elenco: [] as Referto[], senza_data: 0, caratteri: 0 } };
          const s = separaReferti(r.referti);
          return {
            nota: s.referti.length + (s.referti.length === 1 ? ' referto' : ' referti') + (s.senzaData ? ', ' + s.senzaData + ' senza data' : ', tutti con data'),
            valore: { elenco: s.referti, senza_data: s.senzaData, caratteri: r.referti.length },
          };
        });

        // diaria della cartella clinica incollata
        const diaria = await passo('diaria', async () => {
          const t = r.diaria.replace(/\r\n?/g, '\n').replace(/[ \t]+$/gm, '').trim();
          if (!t) return { salto: 'nessuna diaria incollata', valore: { testo: '', righe: 0, caratteri: 0, inizio: '', fine: '', date_trovate: 0 } };
          const linee = t.split('\n').filter((l) => l.trim());
          const dateTrovate = linee.filter((l) => dataNellaRiga(l)).length;
          return {
            nota: righe(t) + ' righe, ' + t.length + ' caratteri',
            valore: { testo: t, righe: righe(t), caratteri: t.length, inizio: linee[0].slice(0, 120), fine: linee[linee.length - 1].slice(0, 120), date_trovate: dateTrovate },
          };
        });

        // il fascicolo: cio' che il motore ricevera'
        const fascicolo = await passo('fascicolo', async () => {
          const f = { identita: scheda.identita, ricovero: scheda.ricovero, consegne: scheda.campi, terapia, laboratorio, referti, diaria, impostazioni: r.impostazioni };
          const byte = enc.encode(JSON.stringify(f)).length;
          return { nota: Math.round(byte / 1024) + ' KB', valore: { f, byte } };
        });

        const riepilogo = {
          fase: 'prova',
          byte: fascicolo.byte,
          durata_ms: Date.now() - inizio,
          campi_consegne: Object.keys(scheda.campi).filter((k) => (scheda.campi as Record<string, string>)[k]).length,
          esami: laboratorio.n_esami,
          valori_esami: laboratorio.n_valori,
          referti: referti.elenco.length,
          referti_senza_data: referti.senza_data,
          righe_diaria: diaria.righe,
        };
        manda({ passo: 'fine', stato: 'ok', fase: 'prova', fascicolo: fascicolo.f, riepilogo });
        registra('fine', 'ok', fascicolo.byte);
      } catch (e) {
        const msg = String((e as Error)?.message || e);
        manda({ passo: inCorso, stato: 'errore', errore: msg });
        manda({ passo: 'fine', stato: 'errore', errore: msg });
        registra(inCorso, 'errore');
      } finally {
        controller.close();
      }
    },
  });
  return new Response(stream, {
    headers: { ...CORS, 'Content-Type': 'application/x-ndjson; charset=utf-8', 'Cache-Control': 'no-store', 'X-Content-Type-Options': 'nosniff' },
  });
});
