#!/usr/bin/env node
// ─────────────────────────────────────────────────────────────────────────────
//  Auto-aggiornamento SCALE DI VALUTAZIONE dalla fonte unica MDCalc.
//
//  Per ogni scala attiva nel DB Supabase:
//   1. legge la definizione corrente
//   2. interroga Claude DUE volte in modo indipendente (web search ristretta a
//      mdcalc.com, temperature 0) per rileggere i punteggi/soglie ufficiali
//   3. confronta le due letture tra loro (anti-allucinazione) e col DB
//   4. se le due letture CONCORDANO e DIFFERISCONO dal DB → pubblica il cambiamento
//        · policy "sempre automatico": pubblica anche se i casi_verità non passano
//          più, ma con log di allerta gialla (livello=warning)
//        · salva sempre lo snapshot precedente in scale_valutazione_versioni
//   5. fail-safe: qualsiasi errore su una scala → la salta, DB intatto, log error
//
//  Env richieste:
//    SUPABASE_URL, SUPABASE_SERVICE_ROLE_KEY, ANTHROPIC_API_KEY
//  Env opzionali:
//    SCALE_UPDATE_MODEL   (default: claude-sonnet-5)
//    SCALE_UPDATE_DRY_RUN ("true" → non scrive nulla, stampa soltanto)
//    SCALE_UPDATE_ONLY    (id scala singola, per test mirati)
// ─────────────────────────────────────────────────────────────────────────────

const TEST_MODE = process.env.SCALE_UPDATE_TEST === '1';
const SUPABASE_URL   = TEST_MODE ? '' : req('SUPABASE_URL');
const SERVICE_KEY    = TEST_MODE ? '' : req('SUPABASE_SERVICE_ROLE_KEY');
const ANTHROPIC_KEY  = TEST_MODE ? '' : req('ANTHROPIC_API_KEY');
const MODEL          = process.env.SCALE_UPDATE_MODEL || 'claude-sonnet-5';
const DRY_RUN        = String(process.env.SCALE_UPDATE_DRY_RUN || '').toLowerCase() === 'true';
const ONLY           = (process.env.SCALE_UPDATE_ONLY || '').trim();

function req(name) {
  const v = process.env[name];
  if (!v) { console.error(`✗ Variabile d'ambiente mancante: ${name}`); process.exit(1); }
  return v;
}

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
// Confronta punti/valori/soglie ignorando il wording delle label.
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

// Struttura (id criteri + numero opzioni/fasce): decide se si può fare un merge
// valore-su-valore sicuro o se il cambiamento è strutturale.
function firmaStruttura(def) {
  if (!def) return 'null';
  const criteri = (def.criteri || []).map(function (c) {
    return { id: c.id, tipo: c.tipo, nOpz: (c.opzioni || []).length };
  });
  return JSON.stringify({ tipo: def.tipo, formula: def.formula ?? null, criteri, nInterp: (def.interpretazione || []).length });
}

// Fonde i valori numerici estratti sulla definizione DB (preserva label, testi,
// casi_verità curati). Usato quando la struttura coincide.
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
  return nuova;
}

// ── Chiamata Claude con web search ristretta a MDCalc ────────────────────────
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
    interpretazione: (def.interpretazione || []).map(f => ({ min: f.min, max: f.max, testo: f.testo }))
  });

  const prompt =
`Sei un assistente clinico che verifica la definizione di una scala medica sulla fonte ufficiale MDCalc.

Scala: "${scala.nome}" (id interno: ${scala.id})
Pagina ufficiale: ${scala.fonte_url}

Definizione attualmente in uso nel nostro sistema (JSON):
${schemaEsempio}

COMPITO: consulta la pagina MDCalc indicata e verifica se i PUNTEGGI, i VALORI e le SOGLIE numeriche corrispondono ancora.
Restituisci la definizione CORRENTE secondo MDCalc, usando ESATTAMENTE la stessa struttura, gli stessi "id" e lo stesso ordine di criteri e opzioni.
- Cambia SOLO i numeri (punti, valore, min, max, contributo) se su MDCalc sono diversi.
- NON aggiungere o togliere criteri/opzioni se la struttura è invariata.
- Se tutto coincide, restituisci gli stessi identici numeri.

Rispondi SOLTANTO con un blocco di codice \`\`\`json ... \`\`\` contenente la definizione, senza altro testo.`;

  const resp = await fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: {
      'x-api-key': ANTHROPIC_KEY,
      'anthropic-version': '2023-06-01',
      'content-type': 'application/json'
    },
    body: JSON.stringify({
      model: MODEL,
      max_tokens: 3000,
      temperature: 0,
      tools: [{ type: 'web_search_20250305', name: 'web_search', max_uses: 4, allowed_domains: ['mdcalc.com'] }],
      messages: [{ role: 'user', content: prompt }]
    })
  });

  if (!resp.ok) {
    const t = await resp.text().catch(() => '');
    throw new Error(`Anthropic HTTP ${resp.status}: ${t.slice(0, 300)}`);
  }
  const data = await resp.json();
  const testo = (data.content || [])
    .filter(b => b.type === 'text')
    .map(b => b.text)
    .join('\n');
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

      // Doppia lettura indipendente
      const [a, b] = await Promise.all([leggiDaMdcalc(scala), leggiDaMdcalc(scala)]);
      if (!a || !b) throw new Error('estrazione non parsabile');

      const fA = firmaNumerica(a), fB = firmaNumerica(b), fDb = firmaNumerica(scala.definizione);

      if (fA !== fB) {
        esiti.discordanti++;
        console.log(`${tag} ⚠ letture DISCORDANTI → non aggiorno`);
        await scriviLog('warning', 'scale-update', `Scala "${scala.nome}": letture MDCalc discordanti, nessun aggiornamento`,
          'Le due riletture indipendenti da MDCalc non coincidono: possibile ambiguità della fonte. Definizione lasciata invariata.');
        continue;
      }
      if (fA === fDb) {
        esiti.invariate++;
        console.log(`${tag} ✓ invariata`);
        continue;
      }

      // ── Cambiamento affidabile rilevato ──
      const stessaStruttura = firmaStruttura(a) === firmaStruttura(scala.definizione);
      let defNuova;
      if (stessaStruttura) {
        defNuova = mergeValori(scala.definizione, a);           // preserva label/testi/casi
      } else {
        defNuova = JSON.parse(JSON.stringify(a));               // cambiamento strutturale
        defNuova.casi_verita = scala.definizione.casi_verita;  // conserva la rete di sicurezza
        if (scala.definizione.unita && !defNuova.unita) defNuova.unita = scala.definizione.unita;
        if (scala.definizione.tolleranza != null && defNuova.tolleranza == null) defNuova.tolleranza = scala.definizione.tolleranza;
      }

      const casi = verificaCasi(defNuova);
      const casiOk = casi.falliti.length === 0;
      const nuovaVer = (scala.versione || 1) + 1;
      const oggi = new Date().toISOString().slice(0, 10);
      const esitoStr = casi.tot ? (casiOk ? `pass (${casi.ok}/${casi.tot})` : `FAIL (${casi.ok}/${casi.tot}: ${casi.falliti.join('; ')})`) : 'n/a';

      console.log(`${tag} ★ CAMBIAMENTO → v${scala.versione}→v${nuovaVer} | struttura ${stessaStruttura ? 'invariata' : 'MODIFICATA'} | casi ${esitoStr}`);

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
          definizione: scala.definizione, motivo: 'auto-update MDCalc', casi_esito: esitoStr
        }])
      });

      // 2) pubblica la nuova definizione
      await sb(`scale_valutazione?id=eq.${encodeURIComponent(scala.id)}`, {
        method: 'PATCH',
        headers: { Prefer: 'return=minimal' },
        body: JSON.stringify({ definizione: defNuova, versione: nuovaVer, data_revisione: oggi })
      });

      // 3) notifica nel Log dell'app
      await scriviLog(
        casiOk ? 'info' : 'warning',
        'scale-update',
        `Scala "${scala.nome}" aggiornata da MDCalc (v${scala.versione}→v${nuovaVer})` + (casiOk ? '' : ' — VERIFICARE'),
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
