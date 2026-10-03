// ================================================================
// app.js — Logica principale frontend (da Scripts.html)
// Modifiche rispetto alla versione GAS:
//  - ricaricaPagina() usa _sincronizzaEPoiFai invece di reload
//  - eseguiAggiungiLetto / eseguiEliminaLetto: toast inline, no redirect
//  - toggleViewAlt(): toggle locale senza reload pagina
//  - stampaConsegne / apriFinestraStampaSalvata: usa PRINT_URL
//  - salvaNuovoNome: gestisce risposta REST { nome: '...' }
//  - ottieniNomeReparto all'avvio per popolare navbar e modal rinomina
// ================================================================

    // APP_URL e PRINT_URL sono definiti in api.js

    // Avvia il flusso dell'app dopo il login Google completato con successo.
    // Chiamata da _completaLogin() in index.html.
    window._avviaFlussoApp = function() {

      // ── Browser compatibility check ─────────────────────────────────
      if (!window._browserCompatibile) {
        var ov = document.getElementById('backupCheckOverlay');
        if (ov) {
          ov.style.background = '#fff';
          ov.innerHTML = [
            '<div style="text-align:center;background:#fff;border:2px solid #c62828;border-radius:10px;',
            'padding:32px 36px;max-width:480px;margin:20px;font-family:Arial,sans-serif;">',
            '<div style="font-size:3rem;margin-bottom:12px;">⚠️</div>',
            '<h2 style="color:#c62828;margin-bottom:10px;font-size:1.4rem;">Browser non aggiornato</h2>',
            '<p style="color:#37474f;line-height:1.6;margin-bottom:18px;font-size:0.92rem;">',
            'Questo browser non supporta alcune funzionalità richieste.<br>',
            'Aggiorna il browser o usa uno di quelli indicati:</p>',
            '<div style="display:flex;gap:10px;justify-content:center;flex-wrap:wrap;margin-bottom:16px;">',
            '<a href="https://www.google.com/chrome/" target="_top"',
            ' style="background:#4285f4;color:#fff;padding:10px 16px;border-radius:8px;',
            'text-decoration:none;font-weight:bold;font-size:0.85rem;">🌐 Google Chrome</a>',
            '<a href="https://www.microsoft.com/edge" target="_top"',
            ' style="background:#0078d7;color:#fff;padding:10px 16px;border-radius:8px;',
            'text-decoration:none;font-weight:bold;font-size:0.85rem;">🌐 Microsoft Edge</a>',
            '<a href="https://www.apple.com/safari/" target="_top"',
            ' style="background:#1d6fa4;color:#fff;padding:10px 16px;border-radius:8px;',
            'text-decoration:none;font-weight:bold;font-size:0.85rem;">🌐 Safari</a>',
            '</div>',
            '<p style="color:#90a4ae;font-size:0.78rem;">Supportate le ultime 3 versioni di ogni browser.</p>',
            '</div>'
          ].join('');
        }
        return; // stop all initialization
      }

      // FASE 4: avvia init dati — overlay si chiude solo quando le card sono pronte
      function _avviaCaricamentoDati() {
        clearInterval(_lockRenewal);
        clearTimeout(_msgBackupTimer);
        var msg = document.getElementById('backupCheckMsg');
        if (msg) msg.textContent = 'Caricamento dati in corso...';

        // Mostra barra di progresso
        var progressWrap = document.getElementById('backupCheckProgressWrap');
        if (progressWrap) progressWrap.style.display = 'block';

        function _setProgress(pct) {
          var bar = document.getElementById('backupCheckBar');
          var lbl = document.getElementById('backupCheckPct');
          if (bar) bar.style.width = pct + '%';
          if (lbl) lbl.textContent = pct + '%';
        }

        _setProgress(5);
        _caricaColoriTipologie();
        _inizializzaView();
        _inizializzaDataNascita();
        _inizializzaAllergieUppercase();
        _setProgress(10);

        // Carica il nome del reparto in navbar
        _sbOttieniNomeReparto()
          .then(function(res) {
            var nome = (res && res.nome) ? res.nome : (typeof res === 'string' ? res : 'Consegne Reparto');
            var el = document.getElementById('nav-app-name');
            if (el) el.innerText = nome;
          })
          .catch(function(e) { console.error('ottieniNomeReparto error:', e); });

        // FASE 5: l'app si apre APPENA le card sono pronte. La diagnostica
        // non trattiene piu' nessuno: parte subito dopo, in background, con
        // gli stessi automatismi di prima (auto-emergenza solo per rischio
        // salvataggio); il suo esito si legge dal chip PC OK / PC LENTO.
        _sincronizzaEPoiFai(function() {
          _setProgress(100);
          setTimeout(function() {
            var ov = document.getElementById('backupCheckOverlay');
            if (ov) {
              ov.style.transition = 'opacity 0.4s';
              ov.style.opacity = '0';
              setTimeout(function() { ov.style.display = 'none'; }, 420);
            }
            // Se eravamo in version update reload (?_r=), nascondi anche
            // l'overlay versione che copriva tutto il caos visivo del reload.
            var vo = document.getElementById('versionUpdateOverlay');
            if (vo && vo.style.display !== 'none') {
              vo.style.transition = 'opacity 0.4s';
              vo.style.opacity = '0';
              setTimeout(function() { vo.style.display = 'none'; }, 420);
            }
          }, 250);
          // Diagnostica in BACKGROUND: boot-mode, nessun modal; auto-attiva
          // l'emergenza SOLO se 'Rischio bug salvataggio' e' presente.
          // I critici (WebSocket, REST) restano alla diagnostica oraria.
          setTimeout(function() {
            if (typeof window._eseguiDiagnosticaSilent === 'function') {
              window._eseguiDiagnosticaSilent(function() {}, {
                autoAttivaEmergenza: true,
                autoDisattivaEmergenza: true,
                noModal: true,
                soloRischioSalvataggio: true
              });
            }
          }, 1200);
          // Cassaforte Google: controllo invisibile (una fetch); se manca
          // l'autorizzazione di reparto compare solo il bottoncino giallo.
          setTimeout(function() {
            if (typeof window._cassaforteAvvio === 'function') window._cassaforteAvvio();
          }, 2500);
        }, _setProgress);
      }

      // _lockRenewal rimosso: rinnovaLockBackup è un no-op nella versione Supabase
      var _lockRenewal = setInterval(function() {}, 5000);

      // ═══════════════════════════════════════════════════════════════════════
      // BACKUP AUTOMATICO ORARIO
      // ──────────────────────────────────────────────────────────────────────
      // Prima il backup veniva eseguito SOLO al caricamento della pagina:
      // chi apriva al mattino e lasciava aperto tutto il giorno NON faceva
      // mai backup. Se nessun altro client si collegava, i dati restavano
      // senza backup per ore (anche giorni).
      //
      // Adesso, OLTRE al backup al caricamento, c'è un timer che chiama
      // _sbArchiviaGiornoCorrente() periodicamente:
      //
      //   - Tick interno ogni 5 minuti (resiliente al tab throttling)
      //   - _sbArchiviaGiornoCorrente() ha già il guard interno:
      //     se ultimo backup < BACKUP_INTERVALLO_MS (1h) → skip silente
      //   - CAS atomico impedisce double-backup tra client
      //   - visibilitychange: al ritorno tab in foreground forza il check
      //
      // Cosi' anche un PC che resta sempre aperto fa il backup ogni ora.
      // ═══════════════════════════════════════════════════════════════════════
      var BACKUP_CHECK_TICK_MS = 15 * 60 * 1000; // tick ogni 15 min (era 5 min)
      function _backupOrarioAutomatico() {
        // Skip se diagnostica silenziosa in corso (evita interferenze)
        if (typeof _silentDiag !== 'undefined' && _silentDiag) return;
        if (typeof _sbArchiviaGiornoCorrente !== 'function') return;

        // Il backup va solo su Supabase: nessun token da rinnovare prima.
        var pre = Promise.resolve();
        pre.then(function() {
          _sbArchiviaGiornoCorrente().then(function(res) {
            // Logga solo quando un backup è effettivamente partito
            if (res && res.inCorso) {
              console.log('[Backup orario] Backup in corso da questo client');
            }
          }).catch(function(e) {
            console.warn('[Backup orario] Errore:', e && e.message);
          });
        });
      }
      setInterval(_backupOrarioAutomatico, BACKUP_CHECK_TICK_MS);
      document.addEventListener('visibilitychange', function() {
        if (document.visibilityState === 'visible') _backupOrarioAutomatico();
      });

      // ═══════════════════════════════════════════════════════════════════════
      // AUTO-REFRESH NOTTURNO — prevenzione degrado browser
      // ─────────────────────────────────────────────────────────────────────
      // I PC di reparto restano accesi per giorni/settimane: il renderer del
      // browser si degrada progressivamente (memoria, timer, compositor —
      // osservato in lab il 19/07: scroll programmatico congelato su tab
      // molto vecchie). Un reload notturno riparte ogni giorno con un
      // browser "fresco" e con gli asset aggiornati (SW network-first).
      //
      // SICUREZZA (reparto attivo H24 — il reload NON deve mai interrompere
      // il lavoro). Ricarica SOLO se TUTTE le condizioni valgono:
      //   - ora tra le 04:00 e le 04:59 (finestra di minima attività)
      //   - nessuna scheda in focus mode su questo PC
      //   - nessun salvataggio pendente (_dirtyLetti vuoto, no save in volo)
      //   - nessuna attività utente (mouse/tastiera/touch) da >= 30 minuti
      //   - non già eseguito oggi (localStorage, sopravvive al reload)
      // Se le condizioni non si verificano mai nella finestra → nessun
      // refresh quel giorno (riprova la notte successiva).
      // ═══════════════════════════════════════════════════════════════════════
      window._lastUserActivityTs = Date.now();
      ['pointerdown', 'keydown', 'wheel', 'touchstart'].forEach(function(ev) {
        document.addEventListener(ev, function() {
          window._lastUserActivityTs = Date.now();
        }, { passive: true, capture: true });
      });

      function _autoRefreshNotturno() {
        try {
          if (new Date().getHours() !== 4) return;
          var d = new Date();
          var oggi = d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
          if (localStorage.getItem('_ultimoAutoRefreshNotturno') === oggi) return;
          if (typeof _focusCard !== 'undefined' && _focusCard) return;
          if (typeof _dirtyLetti !== 'undefined' && _dirtyLetti && _dirtyLetti.size > 0) return;
          if (typeof _syncPaused !== 'undefined' && _syncPaused) return;
          if (Date.now() - (window._lastUserActivityTs || 0) < 30 * 60 * 1000) return;

          // Tutte le condizioni ok: marca PRIMA del reload (il reload azzera
          // il runtime, il localStorage sopravvive → max 1 refresh/giorno).
          localStorage.setItem('_ultimoAutoRefreshNotturno', oggi);
          if (typeof window._log === 'function') {
            try {
              window._log('info', 'auto-refresh',
                'Riavvio notturno automatico della pagina',
                'Prevenzione del degrado del browser sui PC sempre accesi. ' +
                'Condizioni verificate: nessun editing in corso, nessun salvataggio pendente, ' +
                'inattività >= 30 min, finestra 04:00-04:59.');
            } catch(e) {}
          }
          // Piccolo delay: lascia partire l'insert del log prima del reload
          setTimeout(function() { location.reload(); }, 800);
        } catch(e) {}
      }
      window._autoRefreshNotturno = _autoRefreshNotturno; // esposta per test/diagnostica
      setInterval(_autoRefreshNotturno, 10 * 60 * 1000);

      // ── Cleanup log vecchi (> 20 giorni) — eseguito una volta
      // all'avvio + ogni 6 ore se la pagina resta aperta.
      function _cleanupLogVecchi() {
        if (typeof window._sbPuliciLogVecchi !== 'function') return;
        window._sbPuliciLogVecchi(20).catch(function(e) {
          console.warn('[Log cleanup] Errore:', e && e.message);
        });
      }
      // Avvio (con piccolo delay per non sovraccaricare il boot)
      setTimeout(_cleanupLogVecchi, 10000);
      // Periodico ogni 6 ore
      setInterval(_cleanupLogVecchi, 6 * 60 * 60 * 1000);

      // Dopo 3s senza risposta → questo client sta eseguendo il backup
      var _msgBackupTimer = setTimeout(function() {
        var msg = document.getElementById('backupCheckMsg');
        if (msg) msg.innerHTML =
          'Backup in corso...<br>' +
          '<span style="font-size:0.75rem;color:#78909c;font-weight:normal;">' +
          'Non chiudere la pagina.<br>L\'operazione richiede qualche secondo.' +
          '</span>';
      }, 3000);

      // FASE 3: un altro utente sta facendo il backup → polling 3s
      function _mostraAttesoBackup() {
        clearTimeout(_msgBackupTimer);
        var msg = document.getElementById('backupCheckMsg');
        if (msg) msg.innerHTML =
          'Un altro utente sta eseguendo il backup...<br>' +
          '<span style="font-size:0.75rem;color:#78909c;font-weight:normal;">' +
          'Attendi. Il caricamento partirà automaticamente al termine.' +
          '</span>';
        function poll() {
          Promise.resolve({inCorso:false})
            .then(function(res) {
              if (!res || !res.inCorso) { _avviaCaricamentoDati(); }
              else { setTimeout(poll, 3000); }
            })
            .catch(function() { setTimeout(poll, 3000); });
        }
        setTimeout(poll, 3000);
      }

      // FASE 2: controllo backup
      _sbArchiviaGiornoCorrente()
        .then(function(res) {
          clearTimeout(_msgBackupTimer);
          if (res && res.inCorso) {
            _mostraAttesoBackup();
            return;
          }
          _avviaCaricamentoDati();
        })
        .catch(function() {
          clearTimeout(_msgBackupTimer);
          _avviaCaricamentoDati();
        });

    }; // fine _avviaFlussoApp

    let calDate = new Date();
    let listaGiorniArchiviati = [];
    let archivioHtmlPronto = "";

    function apriArchivio() {
      if (typeof _emergGuard === 'function' && _emergGuard()) return;
      _opStart('Caricamento archivio consegne...');
      _sbGetGiorniArchivio()
        .then(function(giorni) {
          listaGiorniArchiviati = giorni;
          calDate = new Date();
          ripristinaCalendarioView();
          _opEnd();
          let modalEl = document.getElementById('modalCalendario');
          let modalInstance = bootstrap.Modal.getOrCreateInstance(modalEl);
          modalInstance.show();
        })
        .catch(function(e) { _opEnd(); console.error('getGiorniArchiviati error:', e); });
    }

    function cambiaMese(dir) { calDate.setMonth(calDate.getMonth() + dir); disegnaCalendario(); }

    function disegnaCalendario() {
      const mesi = ["Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno", "Luglio", "Agosto", "Settembre", "Ottobre", "Novembre", "Dicembre"];
      let anno = calDate.getFullYear(); let mese = calDate.getMonth();
      document.getElementById('cal-month-year').innerText = mesi[mese] + " " + anno;
      let primoGiornoMese = new Date(anno, mese, 1).getDay(); primoGiornoMese = primoGiornoMese === 0 ? 6 : primoGiornoMese - 1;
      let giorniNelMese = new Date(anno, mese + 1, 0).getDate();
      let html = '<div class="row text-center fw-bold mb-2 text-muted" style="font-size:0.9rem;"><div class="col">LUN</div><div class="col">MAR</div><div class="col">MER</div><div class="col">GIO</div><div class="col">VEN</div><div class="col">SAB</div><div class="col">DOM</div></div><div class="row text-center">';
      for(let i = 0; i < primoGiornoMese; i++) html += '<div class="col p-2 border m-1" style="background:#f8f9fa; opacity:0.3; border-radius: 8px;"></div>';
      for(let g = 1; g <= giorniNelMese; g++) {
         let dataStr = anno + "-" + String(mese + 1).padStart(2, '0') + "-" + String(g).padStart(2, '0');
         let hasArchive = listaGiorniArchiviati.includes(dataStr);
         let bgClass = hasArchive ? 'cal-day-hover shadow-sm' : '';
         let style = hasArchive ? 'background-color: #c8e6c9 !important; color: #155724 !important; cursor: pointer; border: 1px solid #81c784;' : 'background-color: #fff; color: #ccc; border: 1px solid #eee;';
         let onclickAction = hasArchive ? `onclick="stampaArchivio('${dataStr}')"` : '';
         html += `<div class="col p-2 m-1 rounded fw-bold fs-5 d-flex align-items-center justify-content-center ${bgClass}" style="${style} min-height: 65px; transition: all 0.2s ease;" ${onclickAction}>${g}</div>`;
         if((g + primoGiornoMese) % 7 === 0) html += '</div><div class="row text-center">';
      }
      html += '</div>';
      document.getElementById('cal-grid').innerHTML = html;
    }

    function ripristinaCalendarioView() {
      document.getElementById('cal-header-controls').style.display = 'flex';
      document.getElementById('cal-info-text').style.display = 'block';
      disegnaCalendario();
    }

    function stampaArchivio(dataStr) {
      document.getElementById('cal-header-controls').style.display = 'none';
      document.getElementById('cal-info-text').style.display = 'none';
      let d = dataStr.split('-'); let dataIta = `${d[2]}/${d[1]}/${d[0]}`;
      document.getElementById('cal-grid').innerHTML = `
        <div class="text-center py-5">
            <div class="spinner-border text-success mb-3" style="width: 3rem; height: 3rem;" role="status"></div>
            <h4 class="text-success fw-bold">Verifica backup disponibili...</h4>
            <p class="text-muted">Controllo archivio del ${dataIta}.</p>
        </div>`;
      _sbGetTimestampsGiorno(dataStr)
        .then(function(timestamps) {
          if (!timestamps || timestamps.length === 0) {
            document.getElementById('cal-grid').innerHTML = `<div class="text-center py-5"><i class="bi bi-exclamation-triangle-fill text-warning" style="font-size: 4rem;"></i><h4 class="mt-3">Nessun backup trovato</h4><button class="btn btn-outline-secondary mt-3" onclick="ripristinaCalendarioView()"><i class="bi bi-arrow-left me-1"></i> Torna al calendario</button></div>`;
            return;
          }
          // Mostra sempre la lista, anche con un solo backup
          let listaHtml = timestamps.map(function(ts) {
            let ora = /^\d{10,13}$/.test(ts)
              ? new Date(Number(ts)).toLocaleTimeString('it-IT', { hour: '2-digit', minute: '2-digit' })
              : (ts.indexOf(' ') !== -1 ? ts.split(' ')[1].substring(0, 5) : ts);
            return `<button class="btn btn-outline-success btn-lg w-100 mb-2 text-start px-4" onclick="apriFinestraStampaSalvata('${ts}')"><i class="bi bi-clock me-2"></i>${ora}</button>`;
          }).join('');
          document.getElementById('cal-grid').innerHTML = `
            <div class="py-4">
              <h5 class="fw-bold mb-1"><i class="bi bi-archive-fill text-success me-2"></i>Backup del ${dataIta}</h5>
              <p class="text-muted mb-3" style="font-size:0.9rem;">Seleziona il backup da stampare:</p>
              ${listaHtml}
              <button class="btn btn-outline-secondary mt-3" onclick="ripristinaCalendarioView()"><i class="bi bi-arrow-left me-1"></i> Torna al calendario</button>
            </div>`;
        })
        .catch(function(e) {
          document.getElementById('cal-grid').innerHTML = `<div class="text-center py-5"><i class="bi bi-exclamation-triangle-fill text-danger" style="font-size: 4rem;"></i><h4 class="mt-3 text-danger">Errore</h4><p>${e.message}</p><button class="btn btn-outline-secondary mt-3" onclick="ripristinaCalendarioView()">Torna al Calendario</button></div>`;
        });
    }

    function apriFinestraStampaSalvata(tsStr) {
      // Chiude il modal calendario e apre il modal impostazioni stampa
      var mod = bootstrap.Modal.getInstance(document.getElementById('modalCalendario'));
      if (mod) mod.hide();
      // Piccolo delay per lasciar completare l'animazione di chiusura Bootstrap
      setTimeout(function() {
        _apriModalScala(function(saltaVuoti, orientamento, tipologie, scala) {
          // Solo layout 'alt' (vista standard rimossa)
          var url = PRINT_URL + '?layout=alt' +
                    '&saltaVuoti=' + (saltaVuoti ? '1' : '0') +
                    '&orientamento=' + orientamento +
                    '&ordinamento=' + encodeURIComponent(localStorage.getItem('ordinamentoPreferito') || 'tipologia') +
                    '&dataArchivio=' + encodeURIComponent(tsStr) +
                    '&tipologie=' + encodeURIComponent(tipologie || 'all') +
                    '&scala=' + (scala || 100);
          window.open(url, '_blank');
        });
      }, 350);
    }


    function _apriModalScala(callback) {
      _scalaModalCallback = callback;
      var modalEl = document.getElementById('modalScalaStampa');
      var mi = bootstrap.Modal.getOrCreateInstance(modalEl);

      // ── Ripristina sempre scala default 80% ad ogni apertura ──────
      var scala80El = document.getElementById('scala80');
      if (scala80El) scala80El.checked = true;

      // ── Caricamento dinamico tipologie (ad ogni apertura) ──────────
      var tipoLista = document.getElementById('tipologieStampaLista');
      var tipoTutte = document.getElementById('checkTuttiTipologie');
      if (tipoLista) {
        tipoLista.innerHTML = '<span class="text-muted small"><span class="spinner-border spinner-border-sm me-1"' +
          ' style="width:.85rem;height:.85rem;border-width:2px;"></span> Caricamento...</span>';
        if (tipoTutte) tipoTutte.checked = true;
        _sbGetTipologieLettiBed()
          .then(function(tipologie) {
            var html = '';
            (tipologie || []).forEach(function(t) {
              var sid = 'ckTip_' + t.replace(/[^a-zA-Z0-9]/g, '_');
              html += '<div class="form-check mb-1">' +
                '<input class="form-check-input tipologia-check" type="checkbox" id="' + sid + '" value="' + t + '"' +
                ' style="width:1.1rem;height:1.1rem;cursor:pointer;">' +
                '<label class="form-check-label ms-2 small" for="' + sid + '">' + t + '</label></div>';
            });
            tipoLista.innerHTML = html || '';
          })
          .catch(function() { tipoLista.innerHTML = ''; });
      }

      if (!modalEl._stampaBound) {
        modalEl._stampaBound = true;

        // "Tutte" checkbox: deseleziona le singole
        var tutteEl = document.getElementById('checkTuttiTipologie');
        if (tutteEl) {
          tutteEl.addEventListener('change', function() {
            if (!this.checked) return;
            var lista = document.getElementById('tipologieStampaLista');
            if (lista) lista.querySelectorAll('.tipologia-check').forEach(function(c) { c.checked = false; });
          });
        }

        // Delegazione su lista: singola → gestisce "tutte"
        var listaEl = document.getElementById('tipologieStampaLista');
        if (listaEl) {
          listaEl.addEventListener('change', function(e) {
            if (!e.target.classList.contains('tipologia-check')) return;
            var ttEl = document.getElementById('checkTuttiTipologie');
            if (e.target.checked && ttEl) ttEl.checked = false;
            var anyOn = Array.from(this.querySelectorAll('.tipologia-check')).some(function(c) { return c.checked; });
            if (!anyOn && ttEl) ttEl.checked = true;
          });
        }

        // Conferma stampa
        var confBtn = modalEl.querySelector('#btnConfermaStampa');
        if (confBtn) {
          confBtn.addEventListener('click', function() {
            var saltaVuoti = document.getElementById('checkSaltaVuoti').checked;
            var includiNoteEl = document.getElementById('checkIncludiNote');
            var includiNote = includiNoteEl ? includiNoteEl.checked : true;
            var orientamento = 'portrait';
            var tipologie = 'all';
            var scalaEl = document.querySelector('input[name="scalaStampa"]:checked');
            var scala = scalaEl ? parseInt(scalaEl.value, 10) : 100;
            var ttEl = document.getElementById('checkTuttiTipologie');
            if (ttEl && !ttEl.checked) {
              var lEl = document.getElementById('tipologieStampaLista');
              if (lEl) {
                var sel = Array.from(lEl.querySelectorAll('.tipologia-check:checked')).map(function(c) { return c.value; });
                if (sel.length > 0) tipologie = sel.join(',');
              }
            }
            // ⚠ Chiudi PRIMA il modal (teardown del focus-trap Bootstrap con
            // la tab ANCORA in primo piano), POI apri la stampa. Se window.open
            // (dentro il callback) partisse mentre il modal si sta chiudendo,
            // la nuova tab manda in background la principale, l'animazione di
            // chiusura si sospende, 'hidden.bs.modal' non scatta e il focus-trap
            // resta ATTIVO: rende non scrivibili gli input dei modal successivi
            // (es. il campo del modal isolamento). Il fallback timeout copre
            // l'improbabile caso in cui 'hidden' non scatti.
            var _opt = [saltaVuoti, orientamento, tipologie, scala, includiNote];
            var _lanciato = false;
            var _lancia = function() {
              if (_lanciato) return; _lanciato = true;
              if (typeof _scalaModalCallback === 'function') {
                _scalaModalCallback(_opt[0], _opt[1], _opt[2], _opt[3], _opt[4]);
              }
            };
            modalEl.addEventListener('hidden.bs.modal', _lancia, { once: true });
            setTimeout(_lancia, 700);
            mi.hide();
          });
        }
      }
      mi.show();
    }

    function stampaConsegne() {
      if (typeof _emergGuard === 'function' && _emergGuard()) return;
      _apriModalScala(function(saltaVuoti, orientamento, tipologie, scala, includiNote) {
        // Solo layout 'alt' (vista standard rimossa)
        var url = PRINT_URL + '?layout=alt' +
                  '&saltaVuoti=' + (saltaVuoti ? '1' : '0') +
                  '&orientamento=' + orientamento +
                  '&ordinamento=' + encodeURIComponent(localStorage.getItem('ordinamentoPreferito') || 'tipologia') +
                  '&tipologie=' + encodeURIComponent(tipologie || 'all') +
                  '&scala=' + (scala || 100) +
                  '&includiNote=' + (includiNote !== false ? '1' : '0');
        window.open(url, '_blank');
      });
    }

    // ── Link Utili ────────────────────────────────────────────────────
    function apriLinkUtili() {
      var modalEl = document.getElementById('modalLinkUtili');
      var mi = bootstrap.Modal.getOrCreateInstance(modalEl);
      var container = document.getElementById('linkUtiliList');
      container.innerHTML = '<div class="text-center text-muted py-3"><span class="spinner-border spinner-border-sm me-1"></span> Caricamento...</div>';
      _sbGetLinkUtili()
        .then(function(links) { _renderLinkUtili(links); })
        .catch(function() { container.innerHTML = '<p class="text-danger small text-center py-2">Errore nel caricamento.</p>'; });
      mi.show();
    }

    function _luEsc(s) { return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }

    function _renderLinkUtili(links) {
      var container = document.getElementById('linkUtiliList');
      if (!links || !links.length) {
        container.innerHTML = '<p class="text-muted small text-center py-3">Nessun link salvato.<br>Aggiungine uno qui sotto.</p>';
        return;
      }
      var html = '';
      links.forEach(function(l) {
        html += '<div class="d-flex align-items-center gap-2 py-2 border-bottom" id="luRow_' + l.id + '">' +
          '<a href="' + _luEsc(l.url) + '" target="_blank" rel="noopener noreferrer"' +
          ' class="flex-grow-1 text-truncate small fw-medium text-decoration-none" style="min-width:0;">' +
          '<i class="bi bi-arrow-up-right-square me-1 text-primary"></i>' + _luEsc(l.nome || l.url) +
          '</a>' +
          '<button class="btn btn-outline-secondary btn-sm py-0 px-2" title="Modifica"' +
          ' onclick="_luEditInline(' + l.id + ',\'' + _luEsc(l.nome).replace(/'/g,'&#39;') + '\',\'' + _luEsc(l.url).replace(/'/g,'&#39;') + '\')">' +
          '<i class="bi bi-pencil-fill" style="font-size:.72rem;"></i></button>' +
          '</div>';
      });
      container.innerHTML = html;
    }

    function _luEditInline(id, nome, url) {
      var row = document.getElementById('luRow_' + id);
      if (!row) return;
      row.innerHTML =
        '<div class="flex-grow-1">' +
        '<input type="text" class="form-control form-control-sm mb-1" id="luNome_' + id + '"' +
        ' value="' + _luEsc(nome) + '" placeholder="Nome visualizzato">' +
        '<input type="url" class="form-control form-control-sm" id="luUrl_' + id + '"' +
        ' value="' + _luEsc(url) + '" placeholder="https://...">' +
        '</div>' +
        '<div class="d-flex flex-column gap-1 ms-1">' +
        '<button class="btn btn-success btn-sm py-0 px-2" title="Salva" onclick="_luSalva(' + id + ')">' +
        '<i class="bi bi-check-lg"></i></button>' +
        '<button class="btn btn-outline-danger btn-sm py-0 px-2" title="Elimina" onclick="_luElimina(' + id + ')">' +
        '<i class="bi bi-trash3"></i></button>' +
        '</div>';
    }

    function _luSalva(id) {
      var nomeEl = document.getElementById('luNome_' + id);
      var urlEl  = document.getElementById('luUrl_'  + id);
      if (!nomeEl || !urlEl) return;
      var row = document.getElementById('luRow_' + id);
      if (row) row.style.opacity = '0.5';
      _sbModificaLink(id, nomeEl.value.trim(), urlEl.value.trim())
        .then(function(r) {
          if (r && r.success) {
            _sbGetLinkUtili().then(_renderLinkUtili).catch(function(e) { console.error('getLinkUtili error:', e); });
          } else {
            if (row) row.style.opacity = '1';
            Swal.fire({ icon: 'error', title: 'Errore', text: 'Impossibile salvare.' });
          }
        })
        .catch(function() {
          if (row) row.style.opacity = '1';
          Swal.fire({ icon: 'error', title: 'Errore', text: 'Impossibile salvare.' });
        });
    }

    function _luElimina(id) {
      Swal.fire({
        icon: 'question', title: 'Eliminare questo link?',
        showCancelButton: true, confirmButtonText: 'Elimina', cancelButtonText: 'Annulla',
        confirmButtonColor: '#dc3545'
      }).then(function(r) {
        if (!r.isConfirmed) return;
        _sbEliminaLink(id)
          .then(function() {
            _sbGetLinkUtili().then(_renderLinkUtili).catch(function(e) { console.error('getLinkUtili error:', e); });
          })
          .catch(function() { Swal.fire({ icon: 'error', title: 'Errore', text: 'Impossibile eliminare.' }); });
      });
    }

    function _aggiungiLink() {
      var nomeEl = document.getElementById('inputNuovoLinkNome');
      var urlEl  = document.getElementById('inputNuovoLinkUrl');
      var nome = nomeEl ? nomeEl.value.trim() : '';
      var url  = urlEl  ? urlEl.value.trim()  : '';
      if (!url) { Swal.fire({ icon: 'warning', title: 'URL mancante', text: 'Inserisci almeno il link.' }); return; }
      if (nomeEl) nomeEl.value = '';
      if (urlEl)  urlEl.value  = '';
      _sbAggiungiLink(nome, url)
        .then(function() {
          _sbGetLinkUtili().then(_renderLinkUtili).catch(function(e) { console.error('getLinkUtili error:', e); });
        })
        .catch(function() { Swal.fire({ icon: 'error', title: 'Errore', text: 'Impossibile aggiungere il link.' }); });
    }
    // ── Fine Link Utili ───────────────────────────────────────────────

    // ── ricaricaPagina (GitHub Pages: niente reload, solo sync) ──────
    function ricaricaPagina() {
      var loadOv = document.getElementById('loadingOverlay');
      if (loadOv) loadOv.style.display = 'block';
      _sincronizzaEPoiFai(function() {
        _inizializzaView();
        _inizializzaDataNascita();
        _inizializzaAllergieUppercase();
        if (loadOv) loadOv.style.display = 'none';
      });
    }

    function showToast(title, msg, type) { const toastContainer = document.getElementById('toastContainer'); let bgColor = type === "danger" ? "bg-danger" : (type === "success" ? "bg-success" : "bg-warning text-dark"); let icon = type === "danger" ? "bi-exclamation-triangle" : (type === "success" ? "bi-check-circle" : "bi-exclamation-circle"); const toastHTML = `<div class="toast align-items-center text-white border-0 mb-2 ${bgColor}" role="alert" aria-live="assertive" aria-atomic="true"><div class="d-flex"><div class="toast-body"><i class="bi ${icon} me-2 fs-5 align-middle"></i><strong>${title}</strong><br><span class="${type==='warning'?'text-dark':'text-white'}">${msg}</span></div><button type="button" class="btn-close ${type==='warning'?'':'btn-close-white'} me-2 m-auto" data-bs-dismiss="toast"></button></div></div>`; const div = document.createElement('div'); div.innerHTML = toastHTML; const toastElement = div.firstElementChild; toastContainer.appendChild(toastElement); const bsToast = new bootstrap.Toast(toastElement, { delay: 4000 }); bsToast.show(); toastElement.addEventListener('hidden.bs.toast', () => { toastElement.remove(); }); }
    function ordinaLetti(criterio, isSilent = false) { const container = document.getElementById('cardsContainer'); if (!container) return; const cards = Array.from(container.querySelectorAll('.patient-card')); if (cards.length === 0) return; cards.sort((a, b) => { const getLettoVal = (el) => el.getAttribute('data-bed'); if(getLettoVal(a)==='NOTE') return 1; if(getLettoVal(b)==='NOTE') return -1; const cmpLetto = (bedA, bedB) => { let numA = parseInt(bedA, 10); let numB = parseInt(bedB, 10); if(!isNaN(numA) && !isNaN(numB)) return numA - numB; return String(bedA).localeCompare(String(bedB)); }; if (criterio === 'numero') return cmpLetto(getLettoVal(a), getLettoVal(b)); if (criterio === 'nome') { let nA=(a.querySelector('[data-field="Nome"]')?.innerText||'').trim().toLowerCase(), nB=(b.querySelector('[data-field="Nome"]')?.innerText||'').trim().toLowerCase(); if(!nA&&!nB) return cmpLetto(getLettoVal(a), getLettoVal(b)); if(!nA) return 1; if(!nB) return -1; let res = nA.localeCompare(nB); return res!==0 ? res : cmpLetto(getLettoVal(a), getLettoVal(b)); } if (criterio === 'tipologia') { let tA=(a.getAttribute('data-tipologia')||'').trim().toLowerCase(), tB=(b.getAttribute('data-tipologia')||'').trim().toLowerCase(); if(!tA&&!tB) return cmpLetto(getLettoVal(a), getLettoVal(b)); if(!tA) return 1; if(!tB) return -1; let res = tA.localeCompare(tB); return res!==0 ? res : cmpLetto(getLettoVal(a), getLettoVal(b)); } }); cards.forEach(card => container.appendChild(card)); localStorage.setItem('ordinamentoPreferito', criterio); }
    function applicaOrdinamentoSalvato() { const saved = localStorage.getItem('ordinamentoPreferito'); if (saved) ordinaLetti(saved, true); else ordinaLetti('tipologia', true); }

    // Restituisce il badge di stato per un letto.
    // - Se la card è in focus mode su questo letto → #focusStatusBadge
    //   (fixed top-left del viewport, sempre visibile anche scrollando).
    // - Altrimenti → il badge dentro la card (vista alt: id "status-alt-{letto}").
    // La vista standard è stata rimossa, quindi non cerchiamo più "status-{letto}".
    function _getBadges(letto) {
      if (typeof _focusCard !== 'undefined' && _focusCard &&
          _focusCard.getAttribute('data-bed') === letto) {
        var fsb = document.getElementById('focusStatusBadge');
        if (fsb) return [fsb];
      }
      var b = document.getElementById('status-alt-' + letto);
      return b ? [b] : [];
    }
    // Compatibilità: restituisce il badge della vista (usato per lettura stato)
    function _getStatusBadge(letto) {
      return document.getElementById('status-alt-' + letto);
    }

    let timerSalvataggioLetto = {}; const RITARDO_SALVATAGGIO = 60000; let timerDissolvenza = {};
    var _dirtyLetti = new Set();        // letti con modifiche non ancora salvate
    var _ultimoSaveTs = {};             // letto -> timestamp ultimo save completato (per heartbeat 90s)
    var _saveRetryCount = {};           // letto -> numero retry attuali (max 3: 5s, 15s, 45s)
    var _saveRetryTimer = {};           // letto -> handle del timer retry corrente
    var _saveInFlight = {};             // letto -> true se HTTP save attualmente in volo (evita doppi save concorrenti)
    var _savePending = {};              // letto -> true se è stato richiesto un nuovo save mentre uno è in volo (queue)
    var _dirtyVersion = {};             // letto -> contatore incrementato ad ogni input (snapshot prima del save HTTP)
    const HEARTBEAT_MS = 90000;         // save di sicurezza ogni 90s di scrittura ininterrotta
    const RETRY_DELAYS = [5000, 15000, 45000]; // backoff esponenziale
    var _countdownIntervallo = {};      // intervalli per il countdown badge
    document.body.addEventListener('input', function(e) { if (e.target && e.target.classList && e.target.classList.contains('editable-area')) { const card = e.target.closest('.patient-card'); if(!card) return; const letto = card.getAttribute('data-bed'); attivaSalvataggioRitardato(letto, card); } });

    // ── Blocco drag-and-drop fuori focus mode ─────────────────────────────
    function _isAreaFuoriFocus(target) {
      var area = target && target.closest && target.closest('.editable-area');
      if (!area) return false;
      var card = area.closest('.patient-card');
      return !card || !card.classList.contains('focus-mode');
    }
    document.addEventListener('dragstart', function(e) {
      if (_isAreaFuoriFocus(e.target)) e.preventDefault();
    });
    document.addEventListener('dragover', function(e) {
      if (_isAreaFuoriFocus(e.target)) { e.preventDefault(); e.dataTransfer.dropEffect = 'none'; }
    });
    document.addEventListener('drop', function(e) {
      if (_isAreaFuoriFocus(e.target)) e.preventDefault();
    });

    // ── Helper: parsing smart data ────────────────────────────────────────
    function _parseDateInput(raw) {
      if (!raw || !raw.trim()) return { valid: false, error: '' };
      var digits = raw.replace(/\D/g, '');
      var dd, mm, yyyy;
      var threshold = new Date().getFullYear() - 2000;
      if (digits.length === 6) {
        dd = parseInt(digits.substring(0,2),10); mm = parseInt(digits.substring(2,4),10);
        var yy = parseInt(digits.substring(4,6),10);
        yyyy = yy <= threshold ? 2000+yy : 1900+yy;
      } else if (digits.length === 8) {
        dd = parseInt(digits.substring(0,2),10); mm = parseInt(digits.substring(2,4),10);
        yyyy = parseInt(digits.substring(4,8),10);
      } else {
        var parts = raw.trim().split('/');
        if (parts.length !== 3) return { valid: false, error: 'Formato non riconosciuto. Usa gg/mm/aaaa' };
        dd = parseInt(parts[0],10); mm = parseInt(parts[1],10);
        var yrStr = parts[2].trim();
        if (yrStr.length === 2) { var yy2=parseInt(yrStr,10); yyyy = yy2<=threshold ? 2000+yy2 : 1900+yy2; }
        else { yyyy = parseInt(yrStr,10); }
      }
      if (isNaN(dd)||isNaN(mm)||isNaN(yyyy)) return { valid:false, error:'La data contiene caratteri non validi' };
      if (mm<1||mm>12)  return { valid:false, error:'Mese non valido (deve essere tra 1 e 12)' };
      if (dd<1||dd>31)  return { valid:false, error:'Giorno non valido (deve essere tra 1 e 31)' };
      if (yyyy<1900||yyyy>2100) return { valid:false, error:'Anno non valido' };
      var d = new Date(yyyy, mm-1, dd);
      if (d.getFullYear()!==yyyy||d.getMonth()!==mm-1||d.getDate()!==dd)
        return { valid:false, error:'Data non valida (il giorno non esiste in quel mese)' };
      return { valid:true, formatted: String(dd).padStart(2,'0')+'/'+String(mm).padStart(2,'0')+'/'+yyyy };
    }
    function _dateClearError(inputEl) {
      inputEl.classList.remove('date-input-error');
      var container = inputEl.closest('.alt-info-row') || inputEl.closest('.d-flex') || inputEl.parentNode;
      var next = container ? container.nextElementSibling : null;
      if (next && next.classList.contains('date-error-msg')) next.remove();
    }
    function _dateSetError(inputEl, errorMsg) {
      inputEl.classList.add('date-input-error');
      var container = inputEl.closest('.alt-info-row') || inputEl.closest('.d-flex') || inputEl.parentNode;
      if (!container) return;
      var next = container.nextElementSibling;
      if (next && next.classList.contains('date-error-msg')) next.remove();
      var msg = document.createElement('span'); msg.className = 'date-error-msg'; msg.textContent = errorMsg;
      container.parentNode.insertBefore(msg, container.nextSibling);
    }
    // ── Seleziona tutto al focus sui campi data (solo se editabili) ────────
    document.addEventListener('focus', function(e) {
      var el = e.target;
      if (el && (el.classList.contains('data-nascita-text') || el.classList.contains('data-ricovero-text'))) {
        if (!el.readOnly && !el.disabled) setTimeout(function() { el.select(); }, 0);
      }
    }, true);
    // ── Formattazione live (oninput) ───────────────────────────────────────
    function formattaDataNascita(inputEl) {
      _dateClearError(inputEl);
      var raw = inputEl.value.replace(/\D/g, '').substring(0, 8);
      var fmt = raw;
      if (raw.length > 4) { fmt = raw.substring(0,2) + '/' + raw.substring(2,4) + '/' + raw.substring(4); }
      else if (raw.length > 2) { fmt = raw.substring(0,2) + '/' + raw.substring(2); }
      inputEl.value = fmt;
    }
    // ── Parsing + validazione su blur ──────────────────────────────────────
    function formattaDataNascitaBlur(inputEl) {
      _dateClearError(inputEl);
      if (!inputEl.value.trim()) { _aggiornaEtaDaNascita(inputEl); return; }
      var result = _parseDateInput(inputEl.value);
      if (result.valid) { inputEl.value = result.formatted; _aggiornaEtaDaNascita(inputEl); }
      else { _dateSetError(inputEl, result.error); }
    }
    function _aggiornaEtaDaNascita(inputEl) {
      var val = inputEl.value;
      var card = inputEl.closest('.patient-card');
      if (!card) return;
      var letto = card.getAttribute('data-bed');
      var etaFields = card.querySelectorAll('.editable-area[data-field="Eta"]');
      var parti = val.split('/');
      if (parti.length === 3 && parti[2].length === 4) {
        var dN = new Date(parseInt(parti[2],10), parseInt(parti[1],10)-1, parseInt(parti[0],10));
        var oggiN = new Date(); dN.setHours(0,0,0,0); oggiN.setHours(0,0,0,0);
        var anni = oggiN.getFullYear() - dN.getFullYear();
        if (oggiN.getMonth() < dN.getMonth() || (oggiN.getMonth() === dN.getMonth() && oggiN.getDate() < dN.getDate())) anni--;
        etaFields.forEach(function(ef) { ef.innerText = anni; ef.contentEditable = 'false'; ef.style.opacity = '0.6'; ef.style.cursor = 'not-allowed'; });
      } else {
        etaFields.forEach(function(ef) { ef.contentEditable = 'true'; ef.style.opacity = ''; ef.style.cursor = ''; });
      }
      attivaSalvataggioRitardato(letto, card);
    }
    function _inizializzaDataNascita() {
      document.querySelectorAll('.data-nascita-text').forEach(function(inp) {
        if (!inp.value) return;
        var card = inp.closest('.patient-card');
        if (!card) return;
        var parti = inp.value.split('/');
        if (parti.length !== 3 || parti[2].length !== 4) return;
        var dN = new Date(parseInt(parti[2],10), parseInt(parti[1],10)-1, parseInt(parti[0],10));
        var oggiN = new Date(); dN.setHours(0,0,0,0); oggiN.setHours(0,0,0,0);
        var anni = oggiN.getFullYear() - dN.getFullYear();
        if (oggiN.getMonth() < dN.getMonth() || (oggiN.getMonth() === dN.getMonth() && oggiN.getDate() < dN.getDate())) anni--;
        card.querySelectorAll('.editable-area[data-field="Eta"]').forEach(function(ef) {
          ef.innerText = anni; ef.contentEditable = 'false'; ef.style.opacity = '0.6'; ef.style.cursor = 'not-allowed';
        });
      });
    }
    function formattaDataRicovero(inputEl) {
      _dateClearError(inputEl);
      var raw = inputEl.value.replace(/\D/g, '').substring(0, 8);
      var fmt = raw;
      if (raw.length > 4) { fmt = raw.substring(0,2) + '/' + raw.substring(2,4) + '/' + raw.substring(4); }
      else if (raw.length > 2) { fmt = raw.substring(0,2) + '/' + raw.substring(2); }
      inputEl.value = fmt;
    }
    function formattaDataRicoveroBlur(inputEl) {
      _dateClearError(inputEl);
      if (!inputEl.value.trim()) { _aggiornaGiorniDaRicovero(inputEl); return; }
      var result = _parseDateInput(inputEl.value);
      if (result.valid) { inputEl.value = result.formatted; _aggiornaGiorniDaRicovero(inputEl); }
      else { _dateSetError(inputEl, result.error); }
    }
    function _aggiornaGiorniDaRicovero(inputEl) {
      var val = inputEl.value;
      var card = inputEl.closest('.patient-card');
      if (!card) return;
      var letto = card.getAttribute('data-bed');
      var parti = val.split('/');
      if (parti.length === 3 && parti[2].length === 4) {
        var dR = new Date(parseInt(parti[2],10), parseInt(parti[1],10)-1, parseInt(parti[0],10));
        var oggiR = new Date(); dR.setHours(0,0,0,0); oggiR.setHours(0,0,0,0);
        var gg = Math.floor((oggiR - dR) / 86400000);
        card.querySelectorAll('.valore-giorni').forEach(function(gf) { gf.innerText = gg >= 0 ? gg : '-'; });
      } else {
        card.querySelectorAll('.valore-giorni').forEach(function(gf) { gf.innerText = '-'; });
      }
      attivaSalvataggioRitardato(letto, card);
    }
    function _inizializzaAllergieUppercase() {
      document.querySelectorAll('[data-field="Allergie"]').forEach(function(el) {
        var walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
        while (walker.nextNode()) { walker.currentNode.nodeValue = walker.currentNode.nodeValue.toUpperCase(); }
      });
    }
    document.body.addEventListener('beforeinput', function(e) {
      if (!e.target || !e.target.matches('[data-field="Allergie"]')) return;
      var card = e.target.closest('.patient-card');
      if (!card || !card.classList.contains('focus-mode')) return;
      if (e.inputType === 'insertText' && e.data) {
        e.preventDefault();
        document.execCommand('insertText', false, e.data.toUpperCase());
      }
    });
    document.body.addEventListener('paste', function(e) {
      if (!e.target || !e.target.matches('[data-field="Allergie"]')) return;
      var card = e.target.closest('.patient-card');
      if (!card || !card.classList.contains('focus-mode')) { e.preventDefault(); return; }
      e.preventDefault();
      var text = (e.clipboardData || window.clipboardData).getData('text/plain') || '';
      document.execCommand('insertText', false, text.toUpperCase());
    }, true);
    // Incolla sempre come testo semplice (senza formattazione) in tutti i campi editabili
    document.body.addEventListener('paste', function(e) {
      if (!e.target) return;
      if (e.target.matches('[data-field="Allergie"]')) return;
      if (!e.target.isContentEditable) return;
      var card = e.target.closest && e.target.closest('.patient-card');
      if (!card || !card.classList.contains('focus-mode')) { e.preventDefault(); return; }
      e.preventDefault();
      var text = (e.clipboardData || window.clipboardData).getData('text/plain') || '';
      document.execCommand('insertText', false, text);
    });
    function attivaSalvataggioRitardato(letto, card) {
      _dirtyLetti.add(letto);
      // FIX #2: incrementa version → ogni input "marca" la card. Al ritorno
      // di un save HTTP, _dopoSalvataggio confronta la version snapshotata
      // all'avvio del save con quella attuale: se cambiata, dirty resta
      // (ci sono modifiche più nuove non ancora salvate).
      _dirtyVersion[letto] = (_dirtyVersion[letto] || 0) + 1;
      clearTimeout(timerSalvataggioLetto[letto]);
      clearInterval(_countdownIntervallo[letto]);
      _letti_salvataggioAttivi.add(letto);
      _nascondMatite();

      // ── HEARTBEAT 90s ────────────────────────────────────────────────
      // Se il medico scrive ininterrottamente, il debounce 60s viene resettato
      // ad ogni tasto e nessun save parte. Heartbeat: se sono passati >=90s
      // dall'ultimo save effettivo per questo letto → save immediato in background.
      // FIX #1: skip se c'è già un save HTTP in volo (evita doppi save concorrenti
      // su PC lenti dove il save impiega secondi e l'utente continua a digitare).
      var ora = Date.now();
      var lastSave = _ultimoSaveTs[letto] || 0;
      if (lastSave && (ora - lastSave) >= HEARTBEAT_MS && !_saveInFlight[letto]) {
        // Save di sicurezza, non resetta debounce: il timer 60s riparte sotto
        try { eseguiSalvataggioLettoCompleto(letto, card); } catch(_e) {}
      } else if (!lastSave) {
        // primo input dopo apertura: stabilisce baseline temporale
        _ultimoSaveTs[letto] = ora;
      }

      // Aggiorna il badge con il conto alla rovescia
      var _sec = Math.round(RITARDO_SALVATAGGIO / 1000);
      function _aggiornaBadge(s) {
        _getBadges(letto).forEach(function(b) {
          clearTimeout(timerDissolvenza[letto]);
          b.className = 'badge status-badge position-absolute top-0 start-0 m-1 bg-warning text-dark visible';
          b.style.cssText += ';opacity:1;';
          b.innerText = 'Salvataggio automatico in ' + s + 's';
        });
      }
      _aggiornaBadge(_sec);

      // Conto alla rovescia: aggiorna ogni secondo
      _countdownIntervallo[letto] = setInterval(function() {
        _sec--;
        if (_sec <= 0) { clearInterval(_countdownIntervallo[letto]); return; }
        _aggiornaBadge(_sec);
      }, 1000);

      // Timer 60s: parte al primo tasto, si resetta ad ogni tasto
      timerSalvataggioLetto[letto] = setTimeout(function() {
        clearInterval(_countdownIntervallo[letto]);
        eseguiSalvataggioLettoCompleto(letto, card);
      }, RITARDO_SALVATAGGIO);
    }
    function eseguiSalvataggioLettoCompleto(letto, card) {
      // FIX #1: se c'è già un save HTTP in volo per questo letto, evita di lanciarne
      // un secondo. Marca _savePending: dopo il completamento del save corrente,
      // _dopoSalvataggio kickerà un nuovo save (con i dati più aggiornati) se
      // _dirtyLetti è ancora popolato (Fix #2).
      if (_saveInFlight[letto]) {
        _savePending[letto] = true;
        return;
      }
      _saveInFlight[letto] = true;
      // FIX #3: traccia il save anche per path non-debounce (retry, force-save, floppy).
      // Fa sì che le matite siano nascoste durante qualsiasi save in volo, non solo quelli
      // partiti da attivaSalvataggioRitardato.
      _letti_salvataggioAttivi.add(letto);
      _nascondMatite();
      // FIX #2: snapshotta la version corrente. Se _dirtyVersion cambia durante il save
      // (utente digita), al ritorno NON cancelliamo il dirty (mod più nuove pending).
      var versionAtSave = _dirtyVersion[letto] || 0;

      _syncPaused = true;
      const datiPaziente = {};
      const aree = card.querySelectorAll('.editable-area');
      aree.forEach(area => { const campo = area.getAttribute('data-field'); if(campo) datiPaziente[campo] = area.classList.contains('plain-text') ? area.innerText : area.innerHTML; });
      const sessoEl = card.querySelector('.sesso-symbol[data-field="Sesso"]');
      if (sessoEl) datiPaziente['Sesso'] = sessoEl.getAttribute('data-sesso') || '';
      const dimRowEl = card.querySelector('.dim-row[data-field="Dimissibile"]');
      if (dimRowEl) datiPaziente['Dimissibile'] = dimRowEl.getAttribute('data-value') || '';
      const dataRicInput = card.querySelector('.data-ricovero-text');
      if(dataRicInput) { var drP = dataRicInput.value.split('/'); datiPaziente['DataRicovero'] = (drP.length===3&&drP[2].length===4) ? drP[2]+'-'+drP[1].padStart(2,'0')+'-'+drP[0].padStart(2,'0') : ''; }
      const nascitaInput = card.querySelector('.data-nascita-text');
      if(nascitaInput) datiPaziente['DataNascita'] = nascitaInput.value;

      // ── BADGE "Salvataggio in corso..." durante la richiesta HTTP ──────
      // Su rete lenta il medico vede che qualcosa sta succedendo (prima il countdown
      // arriva a 0 e finora il badge restava lì silenzioso).
      clearInterval(_countdownIntervallo[letto]);
      _getBadges(letto).forEach(function(b) {
        b.innerText = 'Salvataggio in corso...';
        b.className = 'badge status-badge position-absolute top-0 start-0 m-1 bg-info text-dark visible';
        b.style.cssText += ';opacity:1;';
      });

      function _pianificaRetry() {
        var n = _saveRetryCount[letto] || 0;
        if (n >= RETRY_DELAYS.length) {
          // Tutti i retry esauriti: badge rosso persistente, letto resta dirty
          // per essere ripreso da beforeunload o da prossima digitazione.
          _getBadges(letto).forEach(function(b) {
            b.innerText = 'Salvataggio fallito — controlla connessione';
            b.className = 'badge status-badge position-absolute top-0 start-0 m-1 bg-danger visible';
            b.style.opacity = '1';
          });
          // No timer dissolvenza: il badge resta visibile come allarme persistente
          _saveRetryCount[letto] = 0; // reset per prossima sessione
          // LOG: tutti i retry esauriti — situazione critica, l'utente vede
          // ancora il badge rosso ma il save non riprova più finche' non
          // digita di nuovo
          if (typeof window._log === 'function') {
            window._log('error', 'save-failed-exhausted',
              'Save letto ' + letto + ': retry esauriti dopo ' + RETRY_DELAYS.length + ' tentativi',
              'Tutti i tentativi di salvataggio (5s, 15s, 45s) sono falliti. Il letto resta nello stato "dirty": il save ripartirà alla prossima digitazione o alla chiusura focus mode. Verificare connessione di rete e proxy.',
              null);
          }
          return;
        }
        var delay = RETRY_DELAYS[n];
        _saveRetryCount[letto] = n + 1;
        if (_saveRetryTimer[letto]) clearTimeout(_saveRetryTimer[letto]);
        _saveRetryTimer[letto] = setTimeout(function() {
          _saveRetryTimer[letto] = null;
          // Riprende dalla card più aggiornata in DOM
          var c2 = document.querySelector('.patient-card[data-bed="' + letto + '"]') || card;
          eseguiSalvataggioLettoCompleto(letto, c2);
        }, delay);
        // Badge intermedio
        _getBadges(letto).forEach(function(b) {
          b.innerText = 'Errore — nuovo tentativo tra ' + Math.round(delay/1000) + 's...';
          b.className = 'badge status-badge position-absolute top-0 start-0 m-1 bg-warning text-dark visible';
          b.style.opacity = '1';
        });
      }

      function _dopoSalvataggio(success) {
        clearInterval(_countdownIntervallo[letto]);
        if (success) {
          // FIX #2: cancella dirty SOLO se non ci sono modifiche più nuove
          // (version invariata dal momento del save HTTP). Altrimenti resta
          // dirty per il prossimo save (debounce 60s, heartbeat o pending kick-off).
          if ((_dirtyVersion[letto] || 0) === versionAtSave) {
            _dirtyLetti.delete(letto);
          }
          _ultimoSaveTs[letto] = Date.now();
          _saveRetryCount[letto] = 0;
          if (_saveRetryTimer[letto]) { clearTimeout(_saveRetryTimer[letto]); _saveRetryTimer[letto] = null; }
        }
        // ── Snapshot recovery: chiudi lo snapshot corrente (save/save-failed)
        // e RIAPRINE uno nuovo solo se siamo ancora in focus mode su questo letto.
        if (typeof window._snapshotChiudi === 'function') {
          try { window._snapshotChiudi(letto, success ? 'save' : 'save-failed'); } catch(e) {}
        }
        if (typeof window._snapshotApri === 'function' &&
            typeof _focusCard !== 'undefined' && _focusCard &&
            _focusCard.getAttribute('data-bed') === letto) {
          try { window._snapshotApri(letto, _focusCard); } catch(e) {}
        }
        // Su fallimento: il letto RESTA in _dirtyLetti per consentire il retry
        // (sia in modalità normale tramite _pianificaRetry, sia in emergenza tramite force-save scan).
        _letti_salvataggioAttivi.delete(letto);
        _mostraMatite();
        _syncPaused = false; // Realtime si occupa di aggiornare gli altri client

        // FIX #1 (queue): segnala fine del save in volo
        _saveInFlight[letto] = false;
        // Se c'è un save accodato (richiesto durante il save corrente) e ci sono
        // ancora modifiche pending → kickoff. Esegue in microtask per dare al
        // chiamante il tempo di stabilizzare lo stato (non re-entrante).
        if (_savePending[letto]) {
          _savePending[letto] = false;
          if (_dirtyLetti.has(letto)) {
            setTimeout(function() {
              var c2 = document.querySelector('.patient-card[data-bed="' + letto + '"]') || card;
              if (c2) eseguiSalvataggioLettoCompleto(letto, c2);
            }, 0);
          }
        }
      }
      _sbSalvaPaziente(letto, datiPaziente)
        .then(function(res) {
          if (!res || !res.success) {
            // "Salvataggio impedito — scheda in uso" è un caso particolare (lock di altri):
            // non ha senso ritentare automaticamente, il letto è bloccato sul server.
            _getBadges(letto).forEach(function(b) { b.innerText = 'Salvataggio impedito — scheda in uso'; b.className = 'badge status-badge position-absolute top-0 start-0 m-1 bg-danger visible'; b.style.opacity='1'; });
            if(timerDissolvenza[letto]) clearTimeout(timerDissolvenza[letto]);
            timerDissolvenza[letto] = setTimeout(function() { _getBadges(letto).forEach(function(b) { b.classList.remove('visible'); b.style.opacity='0'; }); }, 6000);
            // Caso "scheda in uso": NON ritentiamo (il server ha respinto perché il lock
            // è di un altro). Rimuoviamo dal dirty per non triggerare loop di retry.
            _dirtyLetti.delete(letto);
            _saveRetryCount[letto] = 0;
            if (_saveRetryTimer[letto]) { clearTimeout(_saveRetryTimer[letto]); _saveRetryTimer[letto] = null; }
            // FIX #1: cancella anche un eventuale pending (otterrebbe lo stesso errore)
            _savePending[letto] = false;
            // LOG: save fallito (problema critico, può causare perdita dati)
            if (typeof window._log === 'function') {
              window._log('error', 'save-failed',
                'Save letto ' + letto + ' fallito',
                (res && res.message) || 'Risposta non success dal server. Probabile causa: 0 righe aggiornate (letto eliminato), rete corrotta, o lock di un altro utente. Il letto NON è stato salvato.',
                null);
            }
            _dopoSalvataggio(false);
            return;
          }
          _getBadges(letto).forEach(function(b) { b.innerText = 'Salvato alle ' + res.ora; b.className = 'badge status-badge position-absolute top-0 start-0 m-1 bg-success visible'; b.style.opacity='1'; });
          if(timerDissolvenza[letto]) clearTimeout(timerDissolvenza[letto]);
          timerDissolvenza[letto] = setTimeout(function() { _getBadges(letto).forEach(function(b) { b.classList.remove('visible'); b.style.opacity='0'; }); }, 6000);
          _dopoSalvataggio(true);
        })
        .catch(function() {
          // Errore di rete: pianifica retry esponenziale (5s → 15s → 45s).
          _dopoSalvataggio(false);
          _pianificaRetry();
        });
    }

    function setLoadingState(buttonId, text) { const btn = document.getElementById(buttonId); if(!btn.dataset.originalText) btn.dataset.originalText = btn.innerHTML; btn.disabled = true; btn.innerHTML = `<span class="spinner-border spinner-border-sm me-2" role="status" aria-hidden="true"></span>${text}`; }
    function resetLoadingState(buttonId) { const btn = document.getElementById(buttonId); btn.disabled = false; if(btn.dataset.originalText) btn.innerHTML = btn.dataset.originalText; }

    function salvaNuovoNome() {
      if (typeof _emergGuard === 'function' && _emergGuard()) return;
      const nuovoNome = document.getElementById("inputNuovoNome").value.trim();
      if(nuovoNome === "") return;
      setLoadingState('btnSalvaRinomina', 'Elaborazione...');
      _sbSalvaNomeReparto(nuovoNome)
        .then(function(res) {
          // REST API restituisce { nome: '...' }, GAS diretto restituisce stringa
          var nomeSalvato = (res && typeof res === 'object' && res.nome) ? res.nome : (typeof res === 'string' ? res : nuovoNome);
          document.getElementById("nav-app-name").innerText = nomeSalvato;
          resetLoadingState('btnSalvaRinomina');
          let modalEl = document.getElementById('modalRinomina');
          let modalInstance = bootstrap.Modal.getInstance(modalEl);
          if(modalInstance) modalInstance.hide();
          Swal.fire({icon: 'success', title: 'Salvato!', timer: 2000, showConfirmButton: false});
        })
        .catch(function(err) {
          resetLoadingState('btnSalvaRinomina');
          Swal.fire({icon: 'error', title: 'Errore', text: err.message || 'Salvataggio fallito.'});
        });
    }

    function eseguiAggiungiLetto() {
      if (typeof _emergGuard === 'function' && _emergGuard()) return;
      var numLetto = document.getElementById('inputNuovoLetto').value.trim().toUpperCase();
      if (!numLetto) { Swal.fire({ icon: 'error', text: 'Inserisci un numero valido.' }); return; }
      var modalEl = document.getElementById('modalAggiungiLetto');
      var mi = bootstrap.Modal.getInstance(modalEl); if (mi) mi.hide();
      Swal.fire({ title: 'Aggiunta letto in corso...', allowOutsideClick: false, showConfirmButton: false, didOpen: function() { Swal.showLoading(); } });
      _sbAggiungiLetto(numLetto)
        .then(function(res) {
          Swal.close();
          if (res.success) {
            _opServer({ barMsg: 'Aggiornamento dati...', successTitle: 'Letto aggiunto',
              successText: 'Letto ' + numLetto + ' aggiunto con successo.',
              errorTitle: 'Errore',
              serverFn: function(onOk) { onOk(); }
            });
          } else {
            Swal.fire({ icon: 'error', title: 'Impossibile aggiungere', text: res.message || 'Operazione fallita.' });
          }
        })
        .catch(function(err) {
          Swal.close();
          Swal.fire({ icon: 'error', title: 'Errore', text: err.message || 'Operazione fallita.' });
        });
    }

    // Controlla i lock sui letti indicati usando lo stato in-memory (0 query).
    // _lockState è mantenuto aggiornato da Realtime in api.js.
    function _verificaLockEProcedi(letti, callback) {
      // Blocca se c'è un lock attivo su uno qualsiasi dei letti, a prescindere da chi lo detiene
      var bloccati = letti.filter(function(l) {
        return (typeof _lockState !== 'undefined') && !!_lockState[String(l)];
      });
      if (bloccati.length > 0) {
        var msg = bloccati.length === 1
          ? 'Il letto <strong>' + bloccati[0] + '</strong> è attualmente in modifica. Attendi il rilascio prima di procedere.'
          : 'I letti <strong>' + bloccati.join('</strong> e <strong>') + '</strong> sono attualmente in modifica. Attendi il rilascio prima di procedere.';
        Swal.fire({ icon: 'error', title: 'Letto in uso', html: msg, confirmButtonColor: '#d33' });
        return;
      }
      callback();
    }

    function eseguiEliminaLetto() {
      if (typeof _emergGuard === 'function' && _emergGuard()) return;
      var numLetto = document.getElementById('selectEliminaLetto').value;
      if (!numLetto) return;
      var modalEl = document.getElementById('modalEliminaLetto');
      var mi = bootstrap.Modal.getInstance(modalEl); if (mi) mi.hide();
      _verificaLockEProcedi([numLetto], function() {
        Swal.fire({ title: 'Sei sicuro?', text: 'Vuoi eliminare definitivamente il letto ' + numLetto + '?', icon: 'warning', showCancelButton: true, confirmButtonColor: '#d33', confirmButtonText: 'Sì, elimina'
        }).then(function(r) {
          if (!r.isConfirmed) return;
          _opServer({ barMsg: 'Eliminazione letto in corso...', successTitle: 'Letto eliminato',
            successText: 'Letto ' + numLetto + ' eliminato con successo.',
            errorTitle: 'Errore',
            serverFn: function(onOk, onErr) {
              _sbEliminaLetto(numLetto)
                .then(function(res) {
                  if (res.success) onOk();
                  else onErr(res.message || 'Operazione fallita.');
                })
                .catch(function(err) {
                  onErr(err.message || 'Operazione fallita.');
                });
            }
          });
        });
      });
    }

    function eseguiDimettiLetto() {
      if (typeof _emergGuard === 'function' && _emergGuard()) return;
      var numLetto = document.getElementById('selectDimettiLetto').value;
      if (!numLetto) return;
      var modalEl = document.getElementById('modalDimetti');
      var mi = bootstrap.Modal.getInstance(modalEl); if (mi) mi.hide();
      _verificaLockEProcedi([numLetto], function() {
      Swal.fire({ title: 'Sei sicuro?', text: 'Svuotare i dati del letto ' + numLetto + '?', icon: 'warning', showCancelButton: true, confirmButtonColor: '#ffc107', confirmButtonText: 'Sì, svuota'
      }).then(function(r) {
        if (!r.isConfirmed) return;
        _opServer({ barMsg: 'Svuotamento letto in corso...', successTitle: 'Letto svuotato', successText: 'I dati del letto ' + numLetto + ' sono stati cancellati.', errorTitle: 'Errore',
          afterSync: function() {
            document.querySelectorAll('.patient-card[data-bed="' + numLetto + '"]').forEach(function(card) {
              // NB: DaFare e NoteTerapia NON sono nella lista perché vengono
              // ripopolati col rispettivo template sotto.
              ['Nome','Diagnosi','Eta','Diaria','PianoTerapeutico','EsamiColturali','Allergie','CodiceSanitario','Ossigeno','Vitto'].forEach(function(campo) {
                var el = card.querySelector('[data-field="' + campo + '"]'); if (el) el.innerHTML = '';
              });
              // DaFare: ripopola con il template (DA FARE / RICHIESTI / ESEGUITI / NOTE)
              var daFareEl = card.querySelector('[data-field="DaFare"]');
              if (daFareEl) {
                daFareEl.innerHTML = (typeof window._TEMPLATE_DAFARE === 'string') ? window._TEMPLATE_DAFARE : '';
              }
              // NoteTerapia: ripopola con il template (TERAPIA EV / OS / SC / AEROSOL)
              var noteTerEl = card.querySelector('[data-field="NoteTerapia"]');
              if (noteTerEl) {
                noteTerEl.innerHTML = (typeof window._TEMPLATE_NOTETERAPIA === 'string') ? window._TEMPLATE_NOTETERAPIA : '';
              }
              var dateEl = card.querySelector('.data-ricovero-text'); if (dateEl) dateEl.value = '';
              var nascEl = card.querySelector('.data-nascita-text'); if (nascEl) nascEl.value = '';
              var ggEl = card.querySelector('.valore-giorni'); if (ggEl) ggEl.innerText = '-';
              // Reset campo Dimissibile
              var dimRow = card.querySelector('.dim-row[data-field="Dimissibile"]');
              if (dimRow) {
                dimRow.setAttribute('data-value', '');
                var cb = dimRow.querySelector('.dim-checkbox'); if (cb) cb.checked = false;
                var info = dimRow.querySelector('.dim-data-info, .dim-data-info-empty');
                if (info) { info.className = 'dim-data-info-empty text-muted'; info.innerHTML = ''; }
              }
              var dimIcon = card.querySelector('.tipo-badge-wrap .dim-icon-badge');
              if (dimIcon) dimIcon.remove();
              // Rimuove le icone di stato (isolamento/rianimazione) accanto al
              // numero letto: la Diaria è stata svuotata, quindi non ci sono più
              // banner. Le scale spariscono già con EsamiColturali svuotato.
              if (typeof window._isolamentoSyncIcona === 'function') window._isolamentoSyncIcona(card);
            });
          },
          serverFn: function(onOk, onErr) { _sbDimettiLetto(numLetto).then(function(res) { if (res.success) onOk(); else onErr(res.message); }).catch(function(err) { onErr(err.message || 'Operazione fallita.'); }); }
        });
      });
      });
    }

    function eseguiSpostaPaziente() {
      if (typeof _emergGuard === 'function' && _emergGuard()) return;
      var selOrig = document.getElementById('selectSpostaOrigine'); var selDest = document.getElementById('selectSpostaDestinazione');
      var lettoOrig = selOrig ? selOrig.value : ''; var lettoDest = selDest ? selDest.value : '';
      if (!lettoOrig || !lettoDest) { Swal.fire({ icon: 'warning', title: 'Attenzione', text: 'Seleziona sia il letto di origine che quello di destinazione.', confirmButtonColor: '#0d6efd' }); return; }
      if (lettoOrig === lettoDest) { Swal.fire({ icon: 'warning', title: 'Attenzione', text: 'Il letto di origine e destinazione sono lo stesso!', confirmButtonColor: '#0d6efd' }); return; }
      var modalEl = document.getElementById('modalSposta'); var mi = bootstrap.Modal.getInstance(modalEl); if (mi) mi.hide();

      Swal.fire({ title: 'Verifica letti in corso...', allowOutsideClick: false, showConfirmButton: false, didOpen: function() { Swal.showLoading(); } });
      _sbGetLettiFull()
        .then(function(letti) {
          Swal.close();
          var origInfo = null, destInfo = null;
          letti.forEach(function(l) {
            if (l.letto === lettoOrig) origInfo = l;
            if (l.letto === lettoDest) destInfo = l;
          });
          var nomeOrigine    = origInfo ? origInfo.nome     : '';
          var nomeDestinazione = destInfo ? destInfo.nome   : '';
          var tipoDestinazione = destInfo ? destInfo.tipologia : 'STANDARD';

          _verificaLockEProcedi([lettoOrig, lettoDest], function() {
            var conferma;
            if (nomeDestinazione && nomeDestinazione.trim() !== '') {
              conferma = Swal.fire({ icon: 'warning', title: 'Letto Occupato', html: 'Il letto <strong>' + lettoDest + '</strong> è occupato da:<br><br><span class="fs-5 fw-bold text-danger">' + nomeDestinazione + '</span><br><span class="text-muted">[' + tipoDestinazione + ']</span><br><br>Se confermi, <strong>' + (nomeOrigine||'il paziente') + '</strong> e <strong>' + nomeDestinazione + '</strong> verranno <u>scambiati</u>.', showCancelButton: true, confirmButtonText: 'Sì, scambia', cancelButtonText: 'Annulla', confirmButtonColor: '#dc3545', cancelButtonColor: '#6c757d', reverseButtons: true });
            } else {
              conferma = Swal.fire({ icon: 'question', title: 'Conferma Spostamento', html: 'Sposto <strong>' + (nomeOrigine||'il paziente') + '</strong> dal letto <strong>' + lettoOrig + '</strong> al letto <strong>' + lettoDest + '</strong> (vuoto). Confermi?', showCancelButton: true, confirmButtonText: 'Sì, sposta', cancelButtonText: 'Annulla', confirmButtonColor: '#0d6efd', cancelButtonColor: '#6c757d', reverseButtons: true });
            }
            conferma.then(function(r) {
              if (!r.isConfirmed) return;
              Swal.fire({ title: 'Acquisizione blocco letti...', allowOutsideClick: false, showConfirmButton: false, didOpen: function() { Swal.showLoading(); } });
              _sbAcquistaLockMultiplo([lettoOrig, lettoDest], _mioToken)
                .then(function(res) {
                  if (!res.success) {
                    var msgErr = (res.bloccati && res.bloccati.length > 0)
                      ? 'I letti <strong>' + res.bloccati.join('</strong> e <strong>') + '</strong> sono stati modificati da un altro utente nel frattempo. Riprova.'
                      : (res.message || 'Impossibile bloccare i letti. Riprova.');
                    Swal.fire({ icon: 'error', title: 'Letto in uso', html: msgErr, confirmButtonColor: '#d33' });
                    return;
                  }
                  _eseguiSpostatoEffettivo(lettoOrig, lettoDest, nomeOrigine, nomeDestinazione.trim() !== '' ? nomeDestinazione : null);
                })
                .catch(function(err) {
                  Swal.fire({ icon: 'error', title: 'Errore', text: err.message || 'Impossibile bloccare i letti.' });
                });
            });
          });
        })
        .catch(function() {
          Swal.close();
          Swal.fire({ icon: 'error', title: 'Errore', text: 'Impossibile verificare lo stato dei letti. Riprova.' });
        });
    }

    function _eseguiSpostatoEffettivo(lettoOrig, lettoDest, nomeOrig, nomeDest) {
      var testo = nomeDest ? (nomeOrig||'Paziente') + ' e ' + nomeDest + ' scambiati correttamente.' : (nomeOrig||'Paziente') + ' spostato al letto ' + lettoDest + '.';
      _opServer({ barMsg: 'Spostamento paziente in corso...', successTitle: 'Spostamento eseguito', successText: testo, errorTitle: 'Errore spostamento',
        serverFn: function(onOk, onErr) {
          _sbSpostaPaziente(lettoOrig, lettoDest)
            .then(function(res) {
              _sbRilasciaLockMultiplo([lettoOrig, lettoDest], _mioToken).catch(function(){});
              if (res.success) onOk(); else onErr(res.message);
            })
            .catch(function(err) {
              _sbRilasciaLockMultiplo([lettoOrig, lettoDest], _mioToken).catch(function(){});
              onErr(err.message || 'Operazione fallita.');
            });
        }
      });
    }

    document.addEventListener('mousedown', function(e) {
      var tb = e.target.closest ? e.target.closest('.focus-toolbar') : null;
      if (tb) { var sel = window.getSelection(); if (sel && sel.rangeCount > 0) _tbSavedRange = sel.getRangeAt(0).cloneRange(); e.preventDefault(); }
      else { _tbSavedRange = null; }
    }, true);

    document.addEventListener('click', function(e) {
      var btn = e.target.closest ? e.target.closest('[data-cmd],[data-fsize],[data-color],[data-fmult]') : null;
      if (!btn) return;
      var tb = btn.closest('.focus-toolbar');
      if (!tb) return;
      e.stopPropagation();
      var cmd = btn.getAttribute('data-cmd'); var fsize = btn.getAttribute('data-fsize'); var color = btn.getAttribute('data-color'); var fmult = btn.getAttribute('data-fmult');
      if (cmd === 'palette') { var p = tb.querySelector('[data-tb-submenu="palette"]'); if (p) p.classList.toggle('open'); return; }
      if (cmd === 'fsmenu')  { var p = tb.querySelector('[data-tb-submenu="fsmenu"]');  if (p) p.classList.toggle('open'); return; }
      if (cmd === 'apriSnapshot') {
        // Click sull'icona ⟲ nella toolbar Diaria/Epicrisi: apri il modal
        // di recupero snapshot per il letto della card corrente.
        var card = tb.closest('.patient-card');
        var letto = card ? card.getAttribute('data-bed') : null;
        if (letto && typeof window._snapshotApriModale === 'function') {
          window._snapshotApriModale(letto);
        }
        return;
      }
      if (cmd === 'isolamento') {
        // Click sull'icona virus nella toolbar Diaria/Epicrisi: apri il
        // modal isolamento per il paziente della card corrente (solo FM).
        var cardIso = tb.closest('.patient-card');
        if (cardIso && typeof window._apriModalIsolamento === 'function') {
          window._apriModalIsolamento(cardIso);
        }
        return;
      }
      if (cmd === 'rianimazione') {
        // Click sull'icona cuore nella toolbar Diaria/Epicrisi: apri il
        // modal rianimazione per il paziente della card corrente (solo FM).
        var cardRi = tb.closest('.patient-card');
        if (cardRi && typeof window._apriModalRianimazione === 'function') {
          window._apriModalRianimazione(cardRi);
        }
        return;
      }
      if (cmd === 'scale') {
        // Click sul pulsante "Scale" nella toolbar Esami Colturali:
        // apri il modal delle scale di valutazione per la card corrente.
        var cardSc = tb.closest('.patient-card');
        if (cardSc && typeof window._apriModalScale === 'function') {
          window._apriModalScale(cardSc);
        }
        return;
      }
      if (cmd === 'labtrak') {
        // Click sul pulsante "Laboratorio": apre il modal degli esami di
        // laboratorio (storico + Aggiorna Esami da TrakCare).
        var cardLb = tb.closest('.patient-card');
        if (cardLb && typeof window._apriLabModal === 'function') {
          window._apriLabModal(cardLb);
        }
        return;
      }
      if (cmd === 'terapiatrak') {
        // Click sul pulsante "Terapia": importa la terapia copiata da
        // TrakCare (bookmarklet) nel campo Note e Terapia della card.
        var cardTk = tb.closest('.patient-card');
        if (cardTk && typeof window._importaTerapiaTrak === 'function') {
          window._importaTerapiaTrak(cardTk);
        }
        return;
      }
      _tbRestoreSelection();
      if (fmult !== null) { _tbApplyFontMultiplier(parseFloat(fmult)); var p = tb.querySelector('[data-tb-submenu="fsmenu"]'); if (p) p.classList.remove('open'); return; }
      if (cmd) { document.execCommand(cmd, false, null); }
      if (color !== null) {
        if (color === 'transparent') { document.execCommand('removeFormat', false, null); }
        else { if (!document.execCommand('hiliteColor', false, color)) document.execCommand('backColor', false, color); }
        var p = tb.querySelector('[data-tb-submenu="palette"]'); if (p) p.classList.remove('open');
      }
      _tbDispatch();
    });

    function _popolaDropdownPazienti() {
      var menu = document.getElementById('menuListaPazienti');
      if (!menu) return;
      // Vista unica (alt) — container id rimane "cardsContainer"
      var cards = Array.from(document.querySelectorAll('#cardsContainer .patient-card'));
      if (cards.length === 0) { menu.innerHTML = '<li><span class="dropdown-item text-muted fst-italic">Nessun paziente.</span></li>'; return; }
      var html = '';
      cards.forEach(function(card) {
        var letto = card.getAttribute('data-bed') || ''; var tipo = (card.getAttribute('data-tipologia') || '').trim();
        var nomeEl = card.querySelector('[data-field="Nome"]'); var nome = nomeEl ? (nomeEl.innerText || nomeEl.textContent || '').trim() : '';
        var sessoEl = card.querySelector('.sesso-symbol[data-field="Sesso"]');
        var sesso = sessoEl ? (sessoEl.getAttribute('data-sesso') || '').toUpperCase() : '';
        var sessoBg = sesso === 'M' ? '#7ec8e3' : (sesso === 'F' ? '#f4a7c3' : '#212529');
        var nomeTxt = nome ? nome : '<em class="text-muted">Vuoto</em>';
        var tipoColore = tipo ? ((typeof window._getColoreTipo === 'function') ? window._getColoreTipo(tipo) : stringToColor(tipo)) : '';
        var tipoBadge = tipo ? ' <span class="badge text-white" style="font-size:0.6rem;background:' + tipoColore + ';">' + tipo + '</span>' : '';
        // Icona dimissione se il paziente ha valore dimissibile.
        // Compare in DUE posizioni:
        //   1. Sovrapposta nell'angolo in alto a sinistra del badge L.X
        //   2. Accanto al nome del paziente
        var dimRow = card.querySelector('.dim-row');
        var dimVal = dimRow ? (dimRow.getAttribute('data-value') || '').trim() : '';
        var dimSvg = (typeof window._dimIconSvg === 'function') ? window._dimIconSvg(11) : '';
        var dimIconAccantoNome = dimVal
          ? ' <span class="dim-icon-list ms-1" title="In dimissione — ' + dimVal + '">' + dimSvg + '</span>'
          : '';
        var dimIconAngolo = dimVal
          ? '<span class="dim-icon-corner" title="In dimissione — ' + dimVal + '">' + dimSvg + '</span>'
          : '';
        // Icone di stato accanto al nome: isolamento e/o rianimazione, se attivi
        // (fonte di verità: i banner nel campo Diaria della card).
        var bIso = card.querySelector('.isolamento-banner');
        var bRianim = card.querySelector('.rianimazione-banner');
        var statusIcons =
          (bIso && typeof window._virusIconSvg === 'function'
            ? ' <span class="bedstatus-icon-list" title="In isolamento — ' + (bIso.getAttribute('data-isolamento') || '') + '">' + window._virusIconSvg(13) + '</span>' : '') +
          (bRianim && typeof window._rianimIconSvg === 'function'
            ? ' <span class="bedstatus-icon-list" title="In rianimazione — ' + (bRianim.getAttribute('data-rianimazione') || '') + '">' + window._rianimIconSvg(13) + '</span>' : '');
        html += '<li><a class="dropdown-item d-flex align-items-center gap-2 py-1" href="javascript:void(0)" data-scroll-letto="' + letto + '">'+
                '<span class="badge text-white position-relative" style="min-width:36px;font-size:0.8rem;background:' + sessoBg + ';">' + dimIconAngolo + 'L.' + letto + '</span>'+
                '<span>' + nomeTxt + tipoBadge + dimIconAccantoNome + statusIcons + '</span>'+
                '</a></li>';
      });
      menu.innerHTML = html;
      menu.onclick = function(e) {
        var a = e.target.closest('[data-scroll-letto]'); if (!a) return;
        var letto = a.getAttribute('data-scroll-letto');
        var tog = document.getElementById('dropdownListaPazienti'); if (tog) { var dd = bootstrap.Dropdown.getInstance(tog); if (dd) dd.hide(); }
        var card = document.querySelector('#cardsContainer .patient-card[data-bed="' + letto + '"]'); if (!card) return;
        // Scroll ROBUSTO a doppio assestamento. Il vecchio one-shot smooth
        // falliva "ogni tanto": se durante l'animazione arrivava un update
        // Realtime / un reflow (equalizza colonne, contenuti applicati),
        // le posizioni cambiavano e lo scroll atterrava nel punto sbagliato;
        // su PC lenti i 200ms non bastavano nemmeno per la chiusura del
        // dropdown. Ora: attesa 250ms → smooth → a +700ms ri-misura e
        // corregge (istantaneo) → ultima verifica a +1500ms.
        function _scrollAllaCard(behavior) {
          var top = card.getBoundingClientRect().top + window.pageYOffset - 80;
          window.scrollTo({ top: Math.max(0, top), behavior: behavior });
        }
        function _fuoriPosizione() {
          // La card è "a posto" se il suo top è vicino agli 80px dalla navbar
          var delta = Math.abs(card.getBoundingClientRect().top - 80);
          // A inizio/fine pagina non si può scrollare oltre: considera ok
          var maxScroll = document.documentElement.scrollHeight - window.innerHeight;
          var target = card.getBoundingClientRect().top + window.pageYOffset - 80;
          if (target <= 0 || target >= maxScroll) return false;
          return delta > 30;
        }
        setTimeout(function() {
          _scrollAllaCard('smooth');
          card.style.transition = 'box-shadow 0.3s'; card.style.boxShadow = '0 0 0 4px rgba(55,71,79,0.45)';
          setTimeout(function() { if (_fuoriPosizione()) _scrollAllaCard('auto'); }, 700);
          setTimeout(function() {
            if (_fuoriPosizione()) {
              // Ultimo tentativo: scrollIntoView usa un percorso diverso del
              // motore e può funzionare dove scrollTo viene ignorato (renderer
              // sotto carico sui PC lenti). scroll-margin-top CSS compensa la navbar.
              try { card.scrollIntoView({ behavior: 'auto', block: 'start' }); } catch(e) {}
            }
            card.style.boxShadow = ''; card.style.transition = '';
          }, 1500);
        }, 250);
      };
    }

    function _opStart(msg) { var bar = document.getElementById('opBar'); var txt = document.getElementById('opBarMsg'); if (bar) bar.style.display = 'flex'; if (txt) txt.textContent = msg || 'Operazione in corso...'; var ov = document.getElementById('opOverlay'); if (ov) ov.classList.add('attivo'); }
    function _opEnd() { var bar = document.getElementById('opBar'); if (bar) bar.style.display = 'none'; var ov = document.getElementById('opOverlay'); if (ov) ov.classList.remove('attivo'); }

    function _applicaAggiornamentoCompleto(html) {
      if (!html) return;
      var nuovoDoc = new DOMParser().parseFromString(html, 'text/html');

      var lettiNuovi = new Set();
      nuovoDoc.querySelectorAll('.patient-card[data-bed]').forEach(function(c) {
        lettiNuovi.add(c.getAttribute('data-bed'));
      });
      if (lettiNuovi.size === 0) {
        _applicaAggiornamentoDaHtml(html);
        if (typeof window._aggiornaBadgePrincipali === 'function') window._aggiornaBadgePrincipali();
        return;
      }

      var lettiAttuali = new Set();
      document.querySelectorAll('.patient-card[data-bed]').forEach(function(c) {
        lettiAttuali.add(c.getAttribute('data-bed'));
      });

      // ── 1. RIMOZIONI ──────────────────────────────────────────────────────
      // GUARD CRITICO: NON rimuovere la card in focus mode anche se manca
      // da lettiNuovi. Senza questo guard, una risposta parziale di
      // _sbGetPazienti (race condition tipica della modalità emergenza con
      // rete instabile) farebbe sparire la card che il medico sta editando,
      // perdendo tutti i dati locali non ancora propagati al DOM via Realtime.
      document.querySelectorAll('.patient-card[data-bed]').forEach(function(card) {
        var lettoCard = card.getAttribute('data-bed');
        if (typeof _lettoLockAtt !== 'undefined' && _lettoLockAtt && _lettoLockAtt === lettoCard) {
          if (!lettiNuovi.has(lettoCard)) {
            console.warn('[Sync] Letto '+lettoCard+' in focus mode MA mancante dal DB sync — card preservata (probabile risposta parziale di _sbGetPazienti)');
          }
          return;
        }
        if (!lettiNuovi.has(lettoCard)) card.remove();
      });

      // ── 2. AGGIUNTE ───────────────────────────────────────────────────────
      // Solo vista alt (la standard è stata rimossa) — un solo container
      var aggiunti = false;
      var nuovoContainer   = nuovoDoc.getElementById('cardsContainer');
      var attualeContainer = document.querySelector('#cardsContainer');

      lettiNuovi.forEach(function(letto) {
        if (lettiAttuali.has(letto)) return;
        aggiunti = true;

        if (nuovoContainer && attualeContainer) {
          var nuovaCard = nuovoContainer.querySelector('.patient-card[data-bed="' + letto + '"]');
          if (nuovaCard) attualeContainer.appendChild(nuovaCard.cloneNode(true));
        }
        // Badge tipologia: solo "badge-tipo-alt-" (era anche "badge-tipo-")
        var badge = document.getElementById('badge-tipo-alt-' + letto);
        if (badge) {
          var val = (badge.innerText || '').trim();
          if (val && val !== 'STANDARD') {
            badge.style.backgroundColor = (typeof window._getColoreTipo === 'function')
              ? window._getColoreTipo(val)
              : stringToColor(val);
          }
        }
      });

      if (aggiunti) { applicaOrdinamentoSalvato(); _inizializzaDataNascita(); }

      // ── 3. AGGIORNAMENTO DATI ─────────────────────────────────────────────
      _applicaAggiornamentoDaHtml(html);

      // ── 3b. RIAPPLICA COLORI CUSTOM BADGE (il renderer usa stringToColor) ──
      if (typeof window._aggiornaBadgePrincipali === 'function') window._aggiornaBadgePrincipali();

      // ── 3c. SCAN PLACEHOLDER (aggiunge/rimuove .is-empty su tutti i campi) ──
      if (typeof window._aggiornaTuttiPlaceholders === 'function') window._aggiornaTuttiPlaceholders();
      if (typeof window._verificaPianoCuraTutti === 'function') window._verificaPianoCuraTutti();
      if (typeof window._verificaTerapiaObsoletaTutti === 'function') window._verificaTerapiaObsoletaTutti();
      if (typeof window._labBottoniAggiorna === 'function') window._labBottoniAggiorna();
      if (typeof window._normalizzaTerapiaBox === 'function') window._normalizzaTerapiaBox();

      // ── 4. MESSAGGIO "NESSUN LETTO" ───────────────────────────────────────
      var _nessunLetto = document.querySelectorAll('.patient-card[data-bed]').length === 0;
      var noLettiMsg = document.getElementById('noLettiMsg'); if (noLettiMsg) noLettiMsg.style.display = _nessunLetto ? '' : 'none';
    }

    function _sincronizzaEPoiFai(callback, onProgress) {
      var ind = document.getElementById('syncIndicator'); var st = document.getElementById('syncStatus');
      if (ind) ind.className = 'badge bg-warning text-dark ms-2'; if (st) st.innerText = 'Aggiornamento...';
      if (typeof onProgress === 'function') onProgress(20);
      _sbGetPazienti().then(function(p){ return _renderCardsHtml(p); })
        .then(function(html) {
          if (typeof onProgress === 'function') onProgress(65);
          return _sbGetLocks()
            .then(function(locks) {
              if (typeof onProgress === 'function') onProgress(90);
              _applicaAggiornamentoCompleto(html); _applicaLocks(locks || {});
              if (ind) ind.className = 'badge bg-success ms-2'; if (st) st.innerText = new Date().toLocaleTimeString('it-IT');
              _syncPaused = false; _aggiornaVoceSync(); _inizializzaRealtime();
              if (typeof _versionCheckInit === 'function') _versionCheckInit();
              // Polling fallback ogni 5 min — safety net se Realtime perde eventi
              if (typeof window._versionCheckPollStart === 'function') window._versionCheckPollStart();
              if (typeof callback === 'function') callback();
            })
            .catch(function() {
              if (typeof onProgress === 'function') onProgress(90);
              _applicaAggiornamentoCompleto(html);
              if (ind) ind.className = 'badge bg-success ms-2'; if (st) st.innerText = new Date().toLocaleTimeString('it-IT');
              _syncPaused = false; _aggiornaVoceSync(); _inizializzaRealtime();
              if (typeof _versionCheckInit === 'function') _versionCheckInit();
              // Polling fallback ogni 5 min — safety net se Realtime perde eventi
              if (typeof window._versionCheckPollStart === 'function') window._versionCheckPollStart();
              if (typeof callback === 'function') callback();
            });
        })
        .catch(function() {
          if (ind) ind.className = 'badge bg-danger ms-2'; if (st) st.innerText = 'Errore sync';
          _syncPaused = false; _aggiornaVoceSync(); _inizializzaRealtime();
          if (typeof callback === 'function') callback();
        });
    }

    function _riprendiSync() { _syncPaused = false; _aggiornaVoceSync(); _inizializzaRealtime(); }

    function _opServer(cfg) {
      _opStart(cfg.barMsg || 'Operazione in corso...');
      var timer = cfg.successTimer !== undefined ? cfg.successTimer : 2200;
      function onOk() {
        _sincronizzaEPoiFai(function() {
          if (typeof cfg.afterSync === 'function') cfg.afterSync();
          _opEnd();
          Swal.fire({ icon: 'success', title: cfg.successTitle || 'Operazione completata', text: cfg.successText || '', timer: timer, showConfirmButton: false, timerProgressBar: true });
        });
      }
      function onErr(msg) { _sincronizzaEPoiFai(function() { _opEnd(); Swal.fire({ icon: 'error', title: cfg.errorTitle || 'Errore', text: msg || 'Si è verificato un errore.' }); }); }
      cfg.serverFn(onOk, onErr);
    }

    function _autoSalvaOra(letto) {
      var card = document.querySelector('.patient-card[data-bed="' + letto + '"]');
      if (!card) return;
      eseguiSalvataggioLettoCompleto(letto, card);
    }

    // NB: La vista standard è stata rimossa il 2026-05-11 in favore della
    // sola vista alternativa (compatta a riga). Le funzioni _inizializzaView,
    // toggleViewAlt, _sincronizzaOrdineAlt e la variabile _viewAltAttiva
    // sono state rimosse. Codice preservato in docs/_legacy/standard-view.md.
    // Il container resta con id "cardsContainer" per compat selettori.

    // Inizializzazione "view" (no-op, mantenuto per non rompere chiamanti)
    function _inizializzaView() {
      applicaOrdinamentoSalvato();
    }

    function ricaricaImmediata() {
      var eraSyncPaused = _syncPaused; _syncPaused = true; ricaricaPagina();
      setTimeout(function() { if (!eraSyncPaused) { _syncPaused = false; _aggiornaVoceSync(); } }, 3000);
    }
