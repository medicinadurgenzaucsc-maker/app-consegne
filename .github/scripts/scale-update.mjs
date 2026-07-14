#!/usr/bin/env node
// ─────────────────────────────────────────────────────────────────────────────
//  Auto-aggiornamento SCALE DI VALUTAZIONE (riferimento MDCalc) — GRATIS.
//  Usa Google Gemini (piano gratuito) per riverificare i punteggi/soglie
//  ufficiali delle scale standard. Nessun costo fisso: ~30 letture/mese stanno
//  ampiamente nel free tier. NB: niente grounding Google Search (non disponibile
//  sul free tier → 429); il modello usa la sua conoscenza clinica consolidata,
//  con doppia lettura + casi_verità come rete di sicurezza.
//
//  Per ogni scala attiva nel DB Supabase:
//   1. legge la definizione corrente
//   2. interroga Gemini DUE volte in modo indipendente (temperature 0) per
//      riverificare i punteggi/soglie ufficiali
//   3. confronta le due letture tra loro (anti-allucinazione) e col DB
//   4. se le due letture CONCORDANO e DIFFERISCONO dal DB → pubblica il cambiamento
//        · policy "sempre automatico": pubblica anche se i casi_verità non passano
//          più, ma con log di allerta gialla (livello=warning)
//        · salva sempre lo snapshot precedente in scale_valutazione_versioni
//   5. fail-safe: qualsiasi errore su una scala → la salta, DB intatto, log error
//
//  Env richieste:
//    SUPABASE_URL, SUPABASE_SERVICE_ROLE_KEY, GEMINI_API_KEY
//  Env opzionali:
//    SCALE_UPDATE_MODEL   (default: gemini-2.0-flash)
//    SCALE_UPDATE_DRY_RUN ("true" → non scrive nulla, stampa soltanto)
//    SCALE_UPDATE_ONLY    (id scala singola, per test mirati)
// ─────────────────────────────────────────────────────────────────────────────

const TEST_MODE = process.env.SCALE_UPDATE_TEST === '1';
const SUPABASE_URL   = TEST_MODE ? '' : req('SUPABASE_URL');
const SERVICE_KEY    = TEST_MODE ? '' : req('SUPABASE_SERVICE_ROLE_KEY');
const GEMINI_KEY     = TEST_MODE ? '' : req('GEMINI_API_KEY');
const MODEL          = process.env.SCALE_UPDATE_MODEL || 'gemini-2.0-flash';
const DRY_RUN        = String(process.env.SCALE_UPDATE_DRY_RUN || '').toLowerCase() === 'true';
const ONLY           = (process.env.SCALE_UPDATE_ONLY || '').trim();

function req(name) {
  const v = process.env[name];
  if (!v) { console.error(`✗ Variabile d'ambiente mancante: ${name}`); process.exit(1); }
  return v;
}

const sleep = ms => new Promise(r => setTimeout(r, ms));

// ── Motore di calcolo — copia FEDELE di docs/js/api.js ───────────────────────
const _SCALE_FORMULE = {
  cockcroft_gault(v) {
    const eta = +v.eta, peso = +v.peso, cr = +v.creatinina;
    if (!(eta > 0) || !(peso > 0) || !(cr > 0)) return null;
    let r = ((140 - eta) * peso) / (72 * cr);
    if (v.sesso === 'F') r *= 0.85;
    return r;
  },
  ckd_epi_2021(v) {
    const eta = +v.eta, cr = +v.creatinina;
    if (!(eta > 0) || !(cr > 0)) return null;
    const f = (v.sesso === 'F');
    const k = f ? 0.7 : 0.9, a = f ? -0.241 : -0.302;
    let r = 142 * Math.pow(Math.min(cr / k, 1), a) *
            Math.pow(Math.max(cr / k, 1), -1.200) * Math.pow(0.9938, eta);
    if (f) r *= 1.012;
    return r;
  },
  glasgow_blatchford(v) {
    const urea = +v.urea, hb = +v.hb, pas = +v.pas;
    if (isNaN(urea) || isNaN(hb) || isNaN(pas)) return null;
    let s = 0;
    if (urea >= 25) s += 6; else if (urea >= 10) s += 4; else if (urea >= 8) s += 3; else if (urea >= 6.5) s += 2;
    if (v.sesso === 'F') { if (hb < 10) s += 6; else if (hb < 12) s += 1; }
    else { if (hb < 10) s += 6; else if (hb < 12) s += 3; else if (hb < 13) s += 1; }
    if (pas < 90) s += 3; else if (pas < 100) s += 2; else if (pas < 110) s += 1;
    if (v.fc) s += 1;
    if (v.melena) s += 1;
    if (v.sincope) s += 2;
    if (v.epatopatia) s += 2;
    if (v.scompenso) s += 2;
    return s;
  }
};

function calcolaScala(def, valori) {
  if (!def) return null;
  let punteggio;
  if (def.tipo === 'formula') {
    const fn = _SCALE_FORMULE[def.formula];
    if (typeof fn !== 'function') return null;
    punteggio = fn(valori || {});
    if (punteggio == null || isNaN(punteggio)) return null;
  } else {
    punteggio = 0;
    (def.criteri || []).forEach(function (c) {
      const val = valori ? valori[c.id] : undefined;
      if (c.tipo === 'bool') {
        if (val) punteggio += (c.punti || 0);
      } else if (c.tipo === 'choice') {
        const idx = (val == null) ? 0 : Number(val);
        const op = (c.opzioni || [])[idx];
        if (op && typeof op.punti === 'number') punteggio += op.punti;
      } else if (c.tipo === 'number' && c.contributo === 'valore') {
        if (val != null && val !== '' && !isNaN(val)) punteggio += Number(val);
      }
    });
  }
  const pRound = Math.round(punteggio);
  let interp = '';
  (def.interpretazione || []).forEach(function (f) {
    if (interp) return;
    if (pRound >= f.min && pRound <= f.max) interp = f.testo;
  });
  return { punteggio, interpretazione: interp };
}

// Esegue i casi_verità di una definizione. → { ok, tot, falliti[] }
function verificaCasi(def) {
  const casi = def.casi_verita || [];
  let ok = 0; const falliti = [];
  casi.forEach(function (cv, i) {
    const ris = calcolaScala(def, cv.input);
    if (!ris) { falliti.push(`#${i} → null`); return; }
    const pass = (def.tipo === 'formula')
      ? Math.abs(Math.round(ris.punteggio) - cv.atteso) <= (def.tolleranza || 0)
      : (ris.punteggio === cv.atteso);
    if (pass) ok++; else falliti.push(`#${i} atteso=${cv.atteso} ott=${ris.punteggio}`);
  });
  return { ok, tot: casi.length, falliti };
}

// ── Firma NUMERICA (solo ciò che cambia il punteggio, non le etichette) ──────
function firmaNumerica(def) {
  if (!def) return 'null';
  const criteri = (def.criteri || []).map(function (c) {
    return {
      id: c.id,
      tipo: c.tipo,
      punti: c.punti ?? null,
      contributo: c.contributo ?? null,
      opzioni: (c.opzioni || []).map(function (o) {
        return { punti: o.punti ?? null, valore: o.valore ?? null };
      })
    };
  });
  const interp = (def.interpretazione || []).map(function (f) {
    return { min: f.min, max: f.max };
  });
  return JSON.stringify({ tipo: def.tipo, formula: def.formula ?? null, criteri, interp });
}

// Struttura (id criteri + numero opzioni/fasce): decide se merge sicuro o cambio strutturale
function firmaStruttura(def) {
  if (!def) return 'null';
  const criteri = (def.criteri || []).map(function (c) {
    return { id: c.id, tipo: c.tipo, nOpz: (c.opzioni || []).length };
  });
  return JSON.stringify({ tipo: def.tipo, formula: def.formula ?? null, criteri, nInterp: (def.interpretazione || []).length });
}

// Fonde i valori numerici estratti sulla definizione DB (preserva label, testi, casi_verità)
function mergeValori(defDb, defEstratta) {
  const nuova = JSON.parse(JSON.stringify(defDb));
  const byId = {};
  (defEstratta.criteri || []).forEach(function (c) { byId[c.id] = c; });
  (nuova.criteri || []).forEach(function (c) {
    const e = byId[c.id];
    if (!e) return;
    if (c.tipo === 'bool' && typeof e.punti === 'number') c.punti = e.punti;
    if (c.tipo === 'number' && e.contributo != null) c.contributo = e.contributo;
    if (c.tipo === 'choice') {
      (c.opzioni || []).forEach(function (o, i) {
        const eo = (e.opzioni || [])[i];
        if (!eo) return;
        if (typeof eo.punti === 'number') o.punti = eo.punti;
        if (eo.valore != null) o.valore = eo.valore;
      });
    }
  });
  (nuova.interpretazione || []).forEach(function (f, i) {
    const ef = (defEstratta.interpretazione || [])[i];
    if (ef && typeof ef.min === 'number') f.min = ef.min;
    if (ef && typeof ef.max === 'number') f.max = ef.max;
  });
  // testi clinici (italiano): aggiorna se l'estrazione li fornisce
  ['quando_usare', 'perche_usare', 'insidie'].forEach(function (k) {
    if (defEstratta[k] && String(defEstratta[k]).trim()) nuova[k] = defEstratta[k];
  });
  return nuova;
}

// ── Chiamata Gemini con retry su rate-limit ──────────────────────────────────
// NB: niente tool `google_search` (grounding): sul piano GRATUITO non è
// disponibile e restituisce 429/RESOURCE_EXHAUSTED. Il modello risponde dalla
// sua conoscenza delle scale standard; la doppia lettura + i casi_verità fanno
// da rete di sicurezza.
async function geminiCall(prompt) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${MODEL}:generateContent?key=${GEMINI_KEY}`;
  const body = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: { temperature: 0, maxOutputTokens: 3000 }
  };
  let ultimoErrore = '';
  for (let tentativo = 1; tentativo <= 4; tentativo++) {
    const resp = await fetch(url, {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify(body)
    });
    if (resp.ok) {
      const data = await resp.json();
      const cand = (data.candidates || [])[0];
      const parts = (cand && cand.content && cand.content.parts) || [];
      return parts.map(p => p.text || '').join('\n');
    }
    const t = await resp.text().catch(() => '');
    let msg = t.slice(0, 200);
    try { const j = JSON.parse(t); if (j.error && j.error.message) msg = j.error.message; } catch (e) {}
    ultimoErrore = `HTTP ${resp.status}: ${msg}`;
    // 429 (rate limit) o 5xx → backoff e ritenta; altri errori → lancia subito
    if (resp.status === 429 || resp.status >= 500) {
      const attesa = 2000 * tentativo;
      console.log(`   ⏳ Gemini ${ultimoErrore.slice(0, 120)} — ritento tra ${attesa}ms (${tentativo}/4)`);
      await sleep(attesa);
      continue;
    }
    throw new Error(`Gemini ${ultimoErrore}`);
  }
  throw new Error(`Gemini: troppi tentativi falliti — ${ultimoErrore}`);
}

async function leggiDaMdcalc(scala) {
  const def = scala.definizione || {};
  const schemaEsempio = JSON.stringify({
    tipo: def.tipo,
    formula: def.formula,
    criteri: (def.criteri || []).map(c => ({
      id: c.id, tipo: c.tipo,
      ...(c.tipo === 'bool' ? { punti: c.punti } : {}),
      ...(c.tipo === 'choice' ? { opzioni: (c.opzioni || []).map(o => ({ label: o.label, punti: o.punti, valore: o.valore })) } : {}),
      ...(c.tipo === 'number' ? { contributo: c.contributo } : {})
    })),
    interpretazione: (def.interpretazione || []).map(f => ({ min: f.min, max: f.max, testo: f.testo })),
    descrizione: scala.descrizione || '',
    quando_usare: def.quando_usare || '',
    perche_usare: def.perche_usare || '',
    insidie: def.insidie || ''
  });

  const prompt =
`Sei un assistente clinico esperto. Verifica una scala medica secondo la definizione validata (come pubblicata su MDCalc: ${scala.fonte_url}).

Scala: "${scala.nome}" (id interno: ${scala.id})

Definizione attualmente in uso nel nostro sistema (JSON):
${schemaEsempio}

COMPITO: restituisci la definizione UFFICIALE CORRENTE usando ESATTAMENTE la stessa struttura,
gli stessi "id" e lo stesso ordine di criteri e opzioni.
- Cambia SOLO i numeri (punti, valore, min, max, contributo) se differiscono dalla definizione ufficiale della scala.
- NON aggiungere o togliere criteri/opzioni se la struttura è invariata.
- Se tutto coincide con la definizione ufficiale, restituisci gli stessi identici numeri.
- Basati sulla definizione clinica consolidata della scala; se non sei certo di un valore, mantieni quello attuale.
- Aggiorna i testi clinici IN ITALIANO, concisi (max ~2 frasi ciascuno), tradotti dalle sezioni ufficiali:
  "descrizione" (breve), "quando_usare" (When to Use), "perche_usare" (Why Use), "insidie" (Pitfalls/Advice).

Rispondi SOLTANTO con un blocco di codice \`\`\`json ... \`\`\` contenente la definizione (con criteri, interpretazione, descrizione, quando_usare, perche_usare, insidie), senza altro testo.`;

  const testo = await geminiCall(prompt);
  return estraiJson(testo);
}

function estraiJson(testo) {
  if (!testo) return null;
  let m = testo.match(/```json\s*([\s\S]*?)```/i) || testo.match(/```\s*([\s\S]*?)```/);
  let raw = m ? m[1] : null;
  if (!raw) {
    const s = testo.indexOf('{'), e = testo.lastIndexOf('}');
    if (s !== -1 && e > s) raw = testo.slice(s, e + 1);
  }
  if (!raw) return null;
  try { return JSON.parse(raw); } catch { return null; }
}

// ── Supabase REST helpers (service_role) ─────────────────────────────────────
async function sb(path, opts = {}) {
  const r = await fetch(`${SUPABASE_URL}/rest/v1/${path}`, {
    ...opts,
    headers: {
      apikey: SERVICE_KEY,
      Authorization: `Bearer ${SERVICE_KEY}`,
      'Content-Type': 'application/json',
      ...(opts.headers || {})
    }
  });
  if (!r.ok) throw new Error(`Supabase ${r.status}: ${(await r.text().catch(() => '')).slice(0, 300)}`);
  const ct = r.headers.get('content-type') || '';
  return ct.includes('application/json') ? r.json() : null;
}

async function scriviLog(livello, tipo, messaggio, descrizione) {
  if (DRY_RUN) { console.log(`   [dry-run] log ${livello}: ${messaggio}`); return; }
  try {
    await sb('logs', {
      method: 'POST',
      headers: { Prefer: 'return=minimal' },
      body: JSON.stringify([{
        ts: Date.now(),
        livello, tipo, messaggio,
        descrizione: descrizione || null,
        device_name: 'Auto-aggiornamento scale',
        device_id: 'github-action'
      }])
    });
  } catch (e) { console.error(`   ✗ log non scritto: ${e.message}`); }
}

// ── Loop principale ──────────────────────────────────────────────────────────
async function main() {
  console.log(`\n╔═ Auto-aggiornamento scale ${DRY_RUN ? '(DRY-RUN)' : ''} — modello ${MODEL}`);
  const filtro = ONLY ? `&id=eq.${encodeURIComponent(ONLY)}` : '';
  const scale = await sb(`scale_valutazione?select=*&attiva=eq.true${filtro}&order=ordine`);
  console.log(`║  ${scale.length} scale da verificare\n`);

  const esiti = { invariate: 0, aggiornate: 0, discordanti: 0, errori: 0 };

  for (const scala of scale) {
    const tag = `[${scala.id}]`;
    try {
      if (!scala.fonte_url) { console.log(`${tag} nessuna fonte_url, salto`); continue; }

      // Doppia lettura indipendente (sequenziale, per non superare il rate-limit free)
      const a = await leggiDaMdcalc(scala);
      await sleep(1500);
      const b = await leggiDaMdcalc(scala);
      if (!a || !b) throw new Error('estrazione non parsabile');

      const fA = firmaNumerica(a), fB = firmaNumerica(b), fDb = firmaNumerica(scala.definizione);

      if (fA !== fB) {
        esiti.discordanti++;
        console.log(`${tag} ⚠ letture DISCORDANTI → non aggiorno`);
        await scriviLog('warning', 'scale-update', `Scala "${scala.nome}": letture MDCalc discordanti, nessun aggiornamento`,
          'Le due riletture indipendenti non coincidono: possibile ambiguità della fonte. Definizione lasciata invariata.');
        continue;
      }
      const testiMancano = !scala.definizione.quando_usare || !scala.definizione.perche_usare || !scala.definizione.insidie;
      if (fA === fDb && !testiMancano) {
        esiti.invariate++;
        console.log(`${tag} ✓ invariata`);
        continue;
      }
      const soloTesti = (fA === fDb); // numeri invariati: si aggiornano solo i testi mancanti

      // ── Cambiamento affidabile rilevato (numerico e/o testi) ──
      const stessaStruttura = firmaStruttura(a) === firmaStruttura(scala.definizione);
      let defNuova;
      if (stessaStruttura) {
        defNuova = mergeValori(scala.definizione, a);           // preserva label/casi, aggiorna numeri+testi
      } else {
        defNuova = JSON.parse(JSON.stringify(a));               // cambiamento strutturale
        defNuova.casi_verita = scala.definizione.casi_verita;  // conserva la rete di sicurezza
        if (scala.definizione.unita && !defNuova.unita) defNuova.unita = scala.definizione.unita;
        if (scala.definizione.tolleranza != null && defNuova.tolleranza == null) defNuova.tolleranza = scala.definizione.tolleranza;
        ['quando_usare', 'perche_usare', 'insidie'].forEach(function (k) {
          if (!defNuova[k] && scala.definizione[k]) defNuova[k] = scala.definizione[k];
        });
      }
      const nuovaDescr = (a.descrizione && String(a.descrizione).trim()) ? a.descrizione : scala.descrizione;

      const casi = verificaCasi(defNuova);
      const casiOk = casi.falliti.length === 0;
      const nuovaVer = (scala.versione || 1) + 1;
      const oggi = new Date().toISOString().slice(0, 10);
      const esitoStr = casi.tot ? (casiOk ? `pass (${casi.ok}/${casi.tot})` : `FAIL (${casi.ok}/${casi.tot}: ${casi.falliti.join('; ')})`) : 'n/a';

      console.log(`${tag} ★ ${soloTesti ? 'TESTI' : 'CAMBIAMENTO'} → v${scala.versione}→v${nuovaVer} | struttura ${stessaStruttura ? 'invariata' : 'MODIFICATA'} | casi ${esitoStr}`);

      if (DRY_RUN) {
        console.log(`${tag}   [dry-run] non scrivo. Nuova firma: ${firmaNumerica(defNuova)}`);
        esiti.aggiornate++;
        continue;
      }

      // 1) snapshot della versione precedente (rollback)
      await sb('scale_valutazione_versioni', {
        method: 'POST',
        headers: { Prefer: 'return=minimal' },
        body: JSON.stringify([{
          scala_id: scala.id, versione: scala.versione || 1,
          nome: scala.nome, fonte_url: scala.fonte_url, riferimento: scala.riferimento,
          definizione: scala.definizione, motivo: 'auto-update MDCalc (Gemini)', casi_esito: esitoStr
        }])
      });

      // 2) pubblica la nuova definizione (+ descrizione colonna)
      await sb(`scale_valutazione?id=eq.${encodeURIComponent(scala.id)}`, {
        method: 'PATCH',
        headers: { Prefer: 'return=minimal' },
        body: JSON.stringify({ definizione: defNuova, versione: nuovaVer, data_revisione: oggi, descrizione: nuovaDescr })
      });

      // 3) notifica nel Log dell'app
      await scriviLog(
        casiOk ? 'info' : 'warning',
        'scale-update',
        `Scala "${scala.nome}" ${soloTesti ? 'testi aggiornati' : 'aggiornata'} da MDCalc (v${scala.versione}→v${nuovaVer})` + (casiOk ? '' : ' — VERIFICARE'),
        (casiOk
          ? `Aggiornamento pubblicato automaticamente. Casi di verifica: ${esitoStr}.`
          : `⚠ Attenzione: la nuova definizione da MDCalc NON supera i casi di verifica noti (${esitoStr}). ` +
            `Pubblicata comunque secondo policy automatica: controlla la scala e, se errata, ripristina la versione precedente ` +
            `dalla tabella scale_valutazione_versioni.`) +
          ` Struttura ${stessaStruttura ? 'invariata' : 'modificata'}.`
      );
      esiti.aggiornate++;
    } catch (e) {
      esiti.errori++;
      console.log(`${tag} ✗ errore: ${e.message}`);
      await scriviLog('error', 'scale-update', `Scala "${scala.nome}": errore durante l'aggiornamento`, e.message);
    }
    await sleep(2500); // margine sul rate-limit del piano gratuito Gemini
  }

  console.log(`\n╚═ Fatto — invariate:${esiti.invariate} aggiornate:${esiti.aggiornate} discordanti:${esiti.discordanti} errori:${esiti.errori}\n`);
  if (!DRY_RUN && (esiti.aggiornate || esiti.discordanti || esiti.errori)) {
    await scriviLog('info', 'scale-update', 'Riepilogo auto-aggiornamento scale',
      `Invariate ${esiti.invariate}, aggiornate ${esiti.aggiornate}, discordanti ${esiti.discordanti}, errori ${esiti.errori}.`);
  }
}

// Funzioni pure esportate per i test unitari (vedi scratchpad test-scale-update)
export { calcolaScala, verificaCasi, firmaNumerica, firmaStruttura, mergeValori, estraiJson, _SCALE_FORMULE };

if (!TEST_MODE) {
  main().catch(e => { console.error('Errore fatale:', e); process.exit(1); });
}
