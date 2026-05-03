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
        google.script.run.pulisciArchivioVecchio();
        google.script.run.pulisciBackupEmergenzaVecchi();
        _caricaColoriTipologie();
        _inizializzaView();
        _inizializzaDataNascita();
        _inizializzaAllergieUppercase();
        _setProgress(10);

        // Carica il nome del reparto in navbar
        google.script.run
          .withSuccessHandler(function(res) {
            var nome = (res && res.nome) ? res.nome : (typeof res === 'string' ? res : 'Consegne Reparto');
            var el = document.getElementById('nav-app-name');
            if (el) el.innerText = nome;
          })
          .ottieniNomeReparto();

        // FASE 5: overlay sparisce solo quando le card sono caricate
        _sincronizzaEPoiFai(function() {
          _setProgress(100);
          setTimeout(function() {
            var ov = document.getElementById('backupCheckOverlay');
            if (ov) {
              ov.style.transition = 'opacity 0.4s';
              ov.style.opacity = '0';
              setTimeout(function() { ov.style.display = 'none'; }, 420);
            }
          }, 300);
        }, _setProgress);
      }

      // Rinnova lock ogni 5s (utile solo se questo client sta eseguendo il backup)
      var _lockRenewal = setInterval(function() {
        google.script.run.rinnovaLockBackup();
      }, 5000);

      // Il backup viene eseguito solo al caricamento pagina (vedi sotto).
      // Non servono setInterval: il CAS in _sbArchiviaGiornoCorrente
      // impedisce che due utenti eseguano il backup contemporaneamente.

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
          google.script.run
            .withSuccessHandler(function(res) {
              if (!res || !res.inCorso) { _avviaCaricamentoDati(); }
              else { setTimeout(poll, 3000); }
            })
            .withFailureHandler(function() { setTimeout(poll, 3000); })
            .checkBackupStatus();
        }
        setTimeout(poll, 3000);
      }

      // FASE 2: controllo backup
      google.script.run
        .withSuccessHandler(function(res) {
          clearTimeout(_msgBackupTimer);
          if (res && res.inCorso) {
            _mostraAttesoBackup();
            return;
          }
          _avviaCaricamentoDati();
        })
        .withFailureHandler(function() {
          clearTimeout(_msgBackupTimer);
          _avviaCaricamentoDati();
        })
        .archiviaGiornoCorrente();

    }; // fine _avviaFlussoApp

    let calDate = new Date();
    let listaGiorniArchiviati = [];
    let archivioHtmlPronto = "";

    function apriArchivio() {
      _opStart('Caricamento archivio consegne...');
      google.script.run.withSuccessHandler((giorni) => {
        listaGiorniArchiviati = giorni;
        calDate = new Date();
        ripristinaCalendarioView();
        _opEnd();
        let modalEl = document.getElementById('modalCalendario');
        let modalInstance = bootstrap.Modal.getOrCreateInstance(modalEl);
        modalInstance.show();
      }).getGiorniArchiviati();
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
      google.script.run
        .withSuccessHandler(function(timestamps) {
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
        .withFailureHandler(function(err) {
          document.getElementById('cal-grid').innerHTML = `<div class="text-center py-5"><i class="bi bi-exclamation-triangle-fill text-danger" style="font-size: 4rem;"></i><h4 class="mt-3 text-danger">Errore</h4><p>${err.message}</p><button class="btn btn-outline-secondary mt-3" onclick="ripristinaCalendarioView()">Torna al Calendario</button></div>`;
        })
        .getTimestampGiorno(dataStr);
    }

    function apriFinestraStampaSalvata(tsStr) {
      // Chiude il modal calendario e apre il modal impostazioni stampa
      var mod = bootstrap.Modal.getInstance(document.getElementById('modalCalendario'));
      if (mod) mod.hide();
      // Piccolo delay per lasciar completare l'animazione di chiusura Bootstrap
      setTimeout(function() {
        _apriModalScala(function(saltaVuoti, orientamento, tipologie, scala) {
          var url = PRINT_URL + '?layout=' + (_viewAltAttiva ? 'alt' : 'main') +
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
        google.script.run
          .withSuccessHandler(function(tipologie) {
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
          .withFailureHandler(function() { tipoLista.innerHTML = ''; })
          .getTipologieLettiBed();
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
            mi.hide();
            if (typeof _scalaModalCallback === 'function') _scalaModalCallback(saltaVuoti, orientamento, tipologie, scala);
          });
        }
      }
      mi.show();
    }

    function stampaConsegne() {
      _apriModalScala(function(saltaVuoti, orientamento, tipologie, scala) {
        var layout = _viewAltAttiva ? 'alt' : 'main';
        var url = PRINT_URL + '?layout=' + layout +
                  '&saltaVuoti=' + (saltaVuoti ? '1' : '0') +
                  '&orientamento=' + orientamento +
                  '&ordinamento=' + encodeURIComponent(localStorage.getItem('ordinamentoPreferito') || 'tipologia') +
                  '&tipologie=' + encodeURIComponent(tipologie || 'all') +
                  '&scala=' + (scala || 100);
        window.open(url, '_blank');
      });
    }

    // ── Link Utili ────────────────────────────────────────────────────
    function apriLinkUtili() {
      var modalEl = document.getElementById('modalLinkUtili');
      var mi = bootstrap.Modal.getOrCreateInstance(modalEl);
      var container = document.getElementById('linkUtiliList');
      container.innerHTML = '<div class="text-center text-muted py-3"><span class="spinner-border spinner-border-sm me-1"></span> Caricamento...</div>';
      google.script.run
        .withSuccessHandler(function(links) { _renderLinkUtili(links); })
        .withFailureHandler(function() { container.innerHTML = '<p class="text-danger small text-center py-2">Errore nel caricamento.</p>'; })
        .getLinkUtili();
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
      google.script.run
        .withSuccessHandler(function(r) {
          if (r && r.success) {
            google.script.run.withSuccessHandler(_renderLinkUtili).getLinkUtili();
          } else {
            if (row) row.style.opacity = '1';
            Swal.fire({ icon: 'error', title: 'Errore', text: 'Impossibile salvare.' });
          }
        })
        .withFailureHandler(function() {
          if (row) row.style.opacity = '1';
          Swal.fire({ icon: 'error', title: 'Errore', text: 'Impossibile salvare.' });
        })
        .modificaLinkUtile(id, nomeEl.value.trim(), urlEl.value.trim());
    }

    function _luElimina(id) {
      Swal.fire({
        icon: 'question', title: 'Eliminare questo link?',
        showCancelButton: true, confirmButtonText: 'Elimina', cancelButtonText: 'Annulla',
        confirmButtonColor: '#dc3545'
      }).then(function(r) {
        if (!r.isConfirmed) return;
        google.script.run
          .withSuccessHandler(function() {
            google.script.run.withSuccessHandler(_renderLinkUtili).getLinkUtili();
          })
          .withFailureHandler(function() { Swal.fire({ icon: 'error', title: 'Errore', text: 'Impossibile eliminare.' }); })
          .eliminaLinkUtile(id);
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
      google.script.run
        .withSuccessHandler(function() {
          google.script.run.withSuccessHandler(_renderLinkUtili).getLinkUtili();
        })
        .withFailureHandler(function() { Swal.fire({ icon: 'error', title: 'Errore', text: 'Impossibile aggiungere il link.' }); })
        .aggiungiLinkUtile(nome, url);
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
    function ordinaLetti(criterio, isSilent = false) { const container = document.getElementById('cardsContainer'); if (!container) return; const cards = Array.from(container.querySelectorAll('.patient-card')); if (cards.length === 0) return; cards.sort((a, b) => { const getLettoVal = (el) => el.getAttribute('data-bed'); const cmpLetto = (bedA, bedB) => { let numA = parseInt(bedA, 10); let numB = parseInt(bedB, 10); if(!isNaN(numA) && !isNaN(numB)) return numA - numB; return String(bedA).localeCompare(String(bedB)); }; if (criterio === 'numero') return cmpLetto(getLettoVal(a), getLettoVal(b)); if (criterio === 'nome') { let nA=(a.querySelector('[data-field="Nome"]').innerText||'').trim().toLowerCase(), nB=(b.querySelector('[data-field="Nome"]').innerText||'').trim().toLowerCase(); if(!nA&&!nB) return cmpLetto(getLettoVal(a), getLettoVal(b)); if(!nA) return 1; if(!nB) return -1; let res = nA.localeCompare(nB); return res!==0 ? res : cmpLetto(getLettoVal(a), getLettoVal(b)); } if (criterio === 'tipologia') { let tA=(a.getAttribute('data-tipologia')||'').trim().toLowerCase(), tB=(b.getAttribute('data-tipologia')||'').trim().toLowerCase(); if(!tA&&!tB) return cmpLetto(getLettoVal(a), getLettoVal(b)); if(!tA) return 1; if(!tB) return -1; let res = tA.localeCompare(tB); return res!==0 ? res : cmpLetto(getLettoVal(a), getLettoVal(b)); } }); cards.forEach(card => container.appendChild(card)); localStorage.setItem('ordinamentoPreferito', criterio); if (typeof _sincronizzaOrdineAlt === 'function') _sincronizzaOrdineAlt(); }
    function applicaOrdinamentoSalvato() { const saved = localStorage.getItem('ordinamentoPreferito'); if (saved) ordinaLetti(saved, true); else ordinaLetti('tipologia', true); }

    // Restituisce entrambi i badge (main + alt) per un letto
    function _getBadges(letto) {
      var b1 = document.getElementById('status-' + letto);
      var b2 = document.getElementById('status-alt-' + letto);
      return [b1, b2].filter(function(b) { return b; });
    }
    // Compatibilità: restituisce il badge della vista attiva (usato per lettura stato)
    function _getStatusBadge(letto) {
      if (_viewAltAttiva) return document.getElementById('status-alt-' + letto);
      return document.getElementById('status-' + letto);
    }

    let timerSalvataggioLetto = {}; const RITARDO_SALVATAGGIO = 60000; let timerDissolvenza = {};
    var _dirtyLetti = new Set();        // letti con modifiche non ancora salvate
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
      clearTimeout(timerSalvataggioLetto[letto]);
      clearInterval(_countdownIntervallo[letto]);
      _letti_salvataggioAttivi.add(letto);
      _nascondMatite();

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
      _syncPaused = true;
      const datiPaziente = {};
      const aree = card.querySelectorAll('.editable-area');
      aree.forEach(area => { const campo = area.getAttribute('data-field'); if(campo) datiPaziente[campo] = area.classList.contains('plain-text') ? area.innerText : area.innerHTML; });
      const sessoEl = card.querySelector('.sesso-symbol[data-field="Sesso"]');
      if (sessoEl) datiPaziente['Sesso'] = sessoEl.getAttribute('data-sesso') || '';
      const dataRicInput = card.querySelector('.data-ricovero-text');
      if(dataRicInput) { var drP = dataRicInput.value.split('/'); datiPaziente['DataRicovero'] = (drP.length===3&&drP[2].length===4) ? drP[2]+'-'+drP[1].padStart(2,'0')+'-'+drP[0].padStart(2,'0') : ''; }
      const nascitaInput = card.querySelector('.data-nascita-text');
      if(nascitaInput) datiPaziente['DataNascita'] = nascitaInput.value;
      function _dopoSalvataggio() {
        clearInterval(_countdownIntervallo[letto]);
        _dirtyLetti.delete(letto);
        _letti_salvataggioAttivi.delete(letto);
        _mostraMatite();
        _syncPaused = false; // Realtime si occupa di aggiornare gli altri client
      }
      google.script.run
        .withSuccessHandler((res) => {
          if (!res || !res.success) {
            _getBadges(letto).forEach(function(b) { b.innerText = 'Salvataggio impedito — scheda in uso'; b.className = 'badge status-badge position-absolute top-0 start-0 m-1 bg-danger visible'; b.style.opacity='1'; });
            if(timerDissolvenza[letto]) clearTimeout(timerDissolvenza[letto]);
            timerDissolvenza[letto] = setTimeout(() => { _getBadges(letto).forEach(function(b) { b.classList.remove('visible'); b.style.opacity='0'; }); }, 6000);
            _dopoSalvataggio(); return;
          }
          _getBadges(letto).forEach(function(b) { b.innerText = 'Salvato alle ' + res.ora; b.className = 'badge status-badge position-absolute top-0 start-0 m-1 bg-success visible'; b.style.opacity='1'; });
          if(timerDissolvenza[letto]) clearTimeout(timerDissolvenza[letto]);
          timerDissolvenza[letto] = setTimeout(() => { _getBadges(letto).forEach(function(b) { b.classList.remove('visible'); b.style.opacity='0'; }); }, 6000);
          _dopoSalvataggio();
        })
        .withFailureHandler((err) => {
          _getBadges(letto).forEach(function(b) { b.innerText = 'Errore di rete — riprova'; b.className = 'badge status-badge position-absolute top-0 start-0 m-1 bg-danger visible'; b.style.opacity='1'; });
          if(timerDissolvenza[letto]) clearTimeout(timerDissolvenza[letto]);
          timerDissolvenza[letto] = setTimeout(() => { _getBadges(letto).forEach(function(b) { b.classList.remove('visible'); b.style.opacity='0'; }); }, 6000);
          _dopoSalvataggio();
        })
        .autoSavePazienteCompleto(letto, datiPaziente, _mioToken);
    }

    function setLoadingState(buttonId, text) { const btn = document.getElementById(buttonId); if(!btn.dataset.originalText) btn.dataset.originalText = btn.innerHTML; btn.disabled = true; btn.innerHTML = `<span class="spinner-border spinner-border-sm me-2" role="status" aria-hidden="true"></span>${text}`; }
    function resetLoadingState(buttonId) { const btn = document.getElementById(buttonId); btn.disabled = false; if(btn.dataset.originalText) btn.innerHTML = btn.dataset.originalText; }

    function salvaNuovoNome() {
      const nuovoNome = document.getElementById("inputNuovoNome").value.trim();
      if(nuovoNome === "") return;
      setLoadingState('btnSalvaRinomina', 'Elaborazione...');
      google.script.run.withSuccessHandler((res) => {
        // REST API restituisce { nome: '...' }, GAS diretto restituisce stringa
        var nomeSalvato = (res && typeof res === 'object' && res.nome) ? res.nome : (typeof res === 'string' ? res : nuovoNome);
        document.getElementById("nav-app-name").innerText = nomeSalvato;
        resetLoadingState('btnSalvaRinomina');
        let modalEl = document.getElementById('modalRinomina');
        let modalInstance = bootstrap.Modal.getInstance(modalEl);
        if(modalInstance) modalInstance.hide();
        Swal.fire({icon: 'success', title: 'Salvato!', timer: 2000, showConfirmButton: false});
      }).withFailureHandler(function(err) {
        resetLoadingState('btnSalvaRinomina');
        Swal.fire({icon: 'error', title: 'Errore', text: err.message || 'Salvataggio fallito.'});
      }).salvaNomeReparto(nuovoNome);
    }

    function eseguiAggiungiLetto() {
      var numLetto = document.getElementById('inputNuovoLetto').value.trim().toUpperCase();
      if (!numLetto) { Swal.fire({ icon: 'error', text: 'Inserisci un numero valido.' }); return; }
      var modalEl = document.getElementById('modalAggiungiLetto');
      var mi = bootstrap.Modal.getInstance(modalEl); if (mi) mi.hide();
      Swal.fire({ title: 'Aggiunta letto in corso...', allowOutsideClick: false, showConfirmButton: false, didOpen: function() { Swal.showLoading(); } });
      google.script.run
        .withSuccessHandler(function(res) {
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
        .withFailureHandler(function(err) {
          Swal.close();
          Swal.fire({ icon: 'error', title: 'Errore', text: err.message || 'Operazione fallita.' });
        })
        .aggiungiLetto(numLetto);
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
              google.script.run
                .withSuccessHandler(function(res) {
                  if (res.success) onOk();
                  else onErr(res.message || 'Operazione fallita.');
                })
                .withFailureHandler(function(err) {
                  onErr(err.message || 'Operazione fallita.');
                })
                .eliminaLetto(numLetto);
            }
          });
        });
      });
    }

    function eseguiDimettiLetto() {
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
              ['Nome','Diagnosi','Eta','NoteTerapia','Diaria','DaFare','PianoTerapeutico','Allergie','CodiceSanitario','Ossigeno'].forEach(function(campo) {
                var el = card.querySelector('[data-field="' + campo + '"]'); if (el) el.innerHTML = '';
              });
              var dateEl = card.querySelector('.data-ricovero-text'); if (dateEl) dateEl.value = '';
              var nascEl = card.querySelector('.data-nascita-text'); if (nascEl) nascEl.value = '';
              var ggEl = card.querySelector('.valore-giorni'); if (ggEl) ggEl.innerText = '-';
            });
          },
          serverFn: function(onOk, onErr) { google.script.run.withSuccessHandler(function(res) { if (res.success) onOk(); else onErr(res.message); }).withFailureHandler(function(err) { onErr(err.message || 'Operazione fallita.'); }).dimettiLetto(numLetto); }
        });
      });
      });
    }

    function eseguiSpostaPaziente() {
      var selOrig = document.getElementById('selectSpostaOrigine'); var selDest = document.getElementById('selectSpostaDestinazione');
      var lettoOrig = selOrig ? selOrig.value : ''; var lettoDest = selDest ? selDest.value : '';
      if (!lettoOrig || !lettoDest) { Swal.fire({ icon: 'warning', title: 'Attenzione', text: 'Seleziona sia il letto di origine che quello di destinazione.', confirmButtonColor: '#0d6efd' }); return; }
      if (lettoOrig === lettoDest) { Swal.fire({ icon: 'warning', title: 'Attenzione', text: 'Il letto di origine e destinazione sono lo stesso!', confirmButtonColor: '#0d6efd' }); return; }
      var modalEl = document.getElementById('modalSposta'); var mi = bootstrap.Modal.getInstance(modalEl); if (mi) mi.hide();

      Swal.fire({ title: 'Verifica letti in corso...', allowOutsideClick: false, showConfirmButton: false, didOpen: function() { Swal.showLoading(); } });
      google.script.run
        .withSuccessHandler(function(letti) {
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
              google.script.run
                .withSuccessHandler(function(res) {
                  if (!res.success) {
                    var msgErr = (res.bloccati && res.bloccati.length > 0)
                      ? 'I letti <strong>' + res.bloccati.join('</strong> e <strong>') + '</strong> sono stati modificati da un altro utente nel frattempo. Riprova.'
                      : (res.message || 'Impossibile bloccare i letti. Riprova.');
                    Swal.fire({ icon: 'error', title: 'Letto in uso', html: msgErr, confirmButtonColor: '#d33' });
                    return;
                  }
                  _eseguiSpostatoEffettivo(lettoOrig, lettoDest, nomeOrigine, nomeDestinazione.trim() !== '' ? nomeDestinazione : null);
                })
                .withFailureHandler(function(err) {
                  Swal.fire({ icon: 'error', title: 'Errore', text: err.message || 'Impossibile bloccare i letti.' });
                })
                .acquistaLockMultiplo([lettoOrig, lettoDest], _mioToken);
            });
          });
        })
        .withFailureHandler(function() {
          Swal.close();
          Swal.fire({ icon: 'error', title: 'Errore', text: 'Impossibile verificare lo stato dei letti. Riprova.' });
        })
        .getLettiFull();
    }

    function _eseguiSpostatoEffettivo(lettoOrig, lettoDest, nomeOrig, nomeDest) {
      var testo = nomeDest ? (nomeOrig||'Paziente') + ' e ' + nomeDest + ' scambiati correttamente.' : (nomeOrig||'Paziente') + ' spostato al letto ' + lettoDest + '.';
      _opServer({ barMsg: 'Spostamento paziente in corso...', successTitle: 'Spostamento eseguito', successText: testo, errorTitle: 'Errore spostamento',
        serverFn: function(onOk, onErr) {
          google.script.run
            .withSuccessHandler(function(res) {
              google.script.run.rilasciaLockMultiplo([lettoOrig, lettoDest], _mioToken);
              if (res.success) onOk(); else onErr(res.message);
            })
            .withFailureHandler(function(err) {
              google.script.run.rilasciaLockMultiplo([lettoOrig, lettoDest], _mioToken);
              onErr(err.message || 'Operazione fallita.');
            })
            .spostaPaziente(lettoOrig, lettoDest);
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
      var containerId = _viewAltAttiva ? 'cardsContainerAlt' : 'cardsContainer';
      var cards = Array.from(document.querySelectorAll('#' + containerId + ' .patient-card'));
      if (cards.length === 0) { menu.innerHTML = '<li><span class="dropdown-item text-muted fst-italic">Nessun paziente.</span></li>'; return; }
      var html = '';
      cards.forEach(function(card) {
        var letto = card.getAttribute('data-bed') || ''; var tipo = (card.getAttribute('data-tipologia') || '').trim();
        var nomeEl = card.querySelector('[data-field="Nome"]'); var nome = nomeEl ? (nomeEl.innerText || nomeEl.textContent || '').trim() : '';
        var nomeTxt = nome ? nome : '<em class="text-muted">Vuoto</em>';
        var tipoColore = tipo ? ((typeof window._getColoreTipo === 'function') ? window._getColoreTipo(tipo) : stringToColor(tipo)) : '';
        var tipoBadge = tipo ? ' <span class="badge text-white" style="font-size:0.6rem;background:' + tipoColore + ';">' + tipo + '</span>' : '';
        html += '<li><a class="dropdown-item d-flex align-items-center gap-2 py-1" href="javascript:void(0)" data-scroll-letto="' + letto + '"><span class="badge bg-dark" style="min-width:36px;font-size:0.8rem;">L.' + letto + '</span><span>' + nomeTxt + tipoBadge + '</span></a></li>';
      });
      menu.innerHTML = html;
      menu.onclick = function(e) {
        var a = e.target.closest('[data-scroll-letto]'); if (!a) return;
        var letto = a.getAttribute('data-scroll-letto');
        var tog = document.getElementById('dropdownListaPazienti'); if (tog) { var dd = bootstrap.Dropdown.getInstance(tog); if (dd) dd.hide(); }
        var containerId = _viewAltAttiva ? 'cardsContainerAlt' : 'cardsContainer';
        var card = document.querySelector('#' + containerId + ' .patient-card[data-bed="' + letto + '"]'); if (!card) return;
        setTimeout(function() {
          var top = card.getBoundingClientRect().top + window.pageYOffset - 80;
          window.scrollTo({ top: Math.max(0, top), behavior: 'smooth' });
          card.style.transition = 'box-shadow 0.3s'; card.style.boxShadow = '0 0 0 4px rgba(55,71,79,0.45)';
          setTimeout(function() { card.style.boxShadow = ''; card.style.transition = ''; }, 1500);
        }, 200);
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
      document.querySelectorAll('.patient-card[data-bed]').forEach(function(card) {
        if (!lettiNuovi.has(card.getAttribute('data-bed'))) card.remove();
      });

      // ── 2. AGGIUNTE ───────────────────────────────────────────────────────
      var aggiunti = false;
      var nuovoContainer    = nuovoDoc.getElementById('cardsContainer');
      var nuovoContainerAlt = nuovoDoc.getElementById('cardsContainerAlt');
      var attualeContainer    = document.querySelector('#cardsContainer');
      var attualeContainerAlt = document.querySelector('#cardsContainerAlt');

      lettiNuovi.forEach(function(letto) {
        if (lettiAttuali.has(letto)) return;
        aggiunti = true;

        if (nuovoContainer && attualeContainer) {
          var nuovaCard = nuovoContainer.querySelector('.patient-card[data-bed="' + letto + '"]');
          if (nuovaCard) attualeContainer.appendChild(nuovaCard.cloneNode(true));
        }
        if (nuovoContainerAlt && attualeContainerAlt) {
          var nuovaRiga = nuovoContainerAlt.querySelector('.patient-card[data-bed="' + letto + '"]');
          if (nuovaRiga) attualeContainerAlt.appendChild(nuovaRiga.cloneNode(true));
        }
        ['badge-tipo-' + letto, 'badge-tipo-alt-' + letto].forEach(function(id) {
          var badge = document.getElementById(id);
          if (badge) { var val = (badge.innerText || '').trim(); if (val && val !== 'STANDARD') badge.style.backgroundColor = (typeof window._getColoreTipo === 'function') ? window._getColoreTipo(val) : stringToColor(val); }
        });
      });

      if (aggiunti) { applicaOrdinamentoSalvato(); _inizializzaDataNascita(); }

      // ── 3. AGGIORNAMENTO DATI ─────────────────────────────────────────────
      _applicaAggiornamentoDaHtml(html);

      // ── 3b. RIAPPLICA COLORI CUSTOM BADGE (il renderer usa stringToColor) ──
      if (typeof window._aggiornaBadgePrincipali === 'function') window._aggiornaBadgePrincipali();

      // ── 4. MESSAGGIO "NESSUN LETTO" ───────────────────────────────────────
      var _nessunLetto = document.querySelectorAll('.patient-card[data-bed]').length === 0;
      var noLettiMsg = document.getElementById('noLettiMsg'); if (noLettiMsg) noLettiMsg.style.display = _nessunLetto ? '' : 'none';
      var noLettiMsgAlt = document.getElementById('noLettiMsgAlt'); if (noLettiMsgAlt) noLettiMsgAlt.style.display = _nessunLetto ? '' : 'none';
    }

    function _sincronizzaEPoiFai(callback, onProgress) {
      var ind = document.getElementById('syncIndicator'); var st = document.getElementById('syncStatus');
      if (ind) ind.className = 'badge bg-warning text-dark ms-2'; if (st) st.innerText = 'Aggiornamento...';
      if (typeof onProgress === 'function') onProgress(20);
      google.script.run
        .withSuccessHandler(function(html) {
          if (typeof onProgress === 'function') onProgress(65);
          google.script.run
            .withSuccessHandler(function(locks) {
              if (typeof onProgress === 'function') onProgress(90);
              _applicaAggiornamentoCompleto(html); _applicaLocks(locks || {});
              if (ind) ind.className = 'badge bg-success ms-2'; if (st) st.innerText = new Date().toLocaleTimeString('it-IT');
              _syncPaused = false; _aggiornaVoceSync(); _inizializzaRealtime();
              if (typeof callback === 'function') callback();
            })
            .withFailureHandler(function() {
              if (typeof onProgress === 'function') onProgress(90);
              _applicaAggiornamentoCompleto(html);
              if (ind) ind.className = 'badge bg-success ms-2'; if (st) st.innerText = new Date().toLocaleTimeString('it-IT');
              _syncPaused = false; _aggiornaVoceSync(); _inizializzaRealtime();
              if (typeof callback === 'function') callback();
            })
            .getLocks();
        })
        .withFailureHandler(function() {
          if (ind) ind.className = 'badge bg-danger ms-2'; if (st) st.innerText = 'Errore sync';
          _syncPaused = false; _aggiornaVoceSync(); _inizializzaRealtime();
          if (typeof callback === 'function') callback();
        })
        .getNewHtml();
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

    // ============================================================
    // TOGGLE VISUALIZZAZIONE ALTERNATIVA — localStorage
    // (GitHub Pages: nessun reload, toggle locale)
    // ============================================================
    var _viewAltAttiva = false;

    function _inizializzaView() {
      var saved = localStorage.getItem('viewAlt');
      // Default: visualizzazione alternativa (null = prima volta, non ancora salvata)
      _viewAltAttiva = saved === null ? true : (saved === '1');
      var main = document.getElementById('cardsContainer');
      var alt  = document.getElementById('cardsContainerAlt');
      var btn  = document.querySelector('#btnToggleViewAlt');
      if (_viewAltAttiva) {
        if (main) main.classList.add('nascosto');
        if (alt)  alt.classList.add('attiva');
        if (btn)  { btn.classList.add('attiva'); btn.style.setProperty('background-color', '#212529', 'important'); btn.style.setProperty('color', '#fff', 'important'); btn.style.setProperty('border-radius', '4px', 'important'); }
      } else {
        if (main) main.classList.remove('nascosto');
        if (alt)  alt.classList.remove('attiva');
        if (btn)  { btn.classList.remove('attiva'); btn.style.removeProperty('background-color'); btn.style.removeProperty('color'); btn.style.removeProperty('border-radius'); }
      }
      applicaOrdinamentoSalvato();
    }

    // GitHub Pages: toggle diretto senza reload
    function toggleViewAlt() {
      var nuovoStato = localStorage.getItem('viewAlt') === '1' ? '0' : '1';
      localStorage.setItem('viewAlt', nuovoStato);
      _inizializzaView();
    }

    // Sincronizza ordine alt con ordine main (usato da ordinaLetti)
    function _sincronizzaOrdineAlt() {
      var main = document.getElementById('cardsContainer');
      var alt  = document.getElementById('cardsContainerAlt');
      if (!main || !alt) return;
      var ordine = Array.from(main.querySelectorAll('.patient-card[data-bed]'))
                        .map(function(c) { return c.getAttribute('data-bed'); });
      ordine.forEach(function(letto) {
        var row = alt.querySelector('.patient-card[data-bed="' + letto + '"]');
        if (row) alt.appendChild(row);
      });
    }

    function ricaricaImmediata() {
      var eraSyncPaused = _syncPaused; _syncPaused = true; ricaricaPagina();
      setTimeout(function() { if (!eraSyncPaused) { _syncPaused = false; _aggiornaVoceSync(); } }, 3000);
    }


    // ══════════════════════════════════════════════════════════════
    // IMPORTA DA FILE CONSEGNE (Google Docs API v1)
    // ══════════════════════════════════════════════════════════════

    var _docsTokenClient = null;
    var _googleDocsToken = null;

    function apriModalImportaConsegne() {
      // Reset UI
      document.getElementById('importaStep1').style.display = '';
      document.getElementById('importaStep2').style.display = 'none';
      document.getElementById('importaStep3').style.display = 'none';
      document.getElementById('importaFooter').style.display = '';
      document.getElementById('inputImportaUrl').value = '';
      document.getElementById('inputImportaPwd').value = '';
      var mi = bootstrap.Modal.getOrCreateInstance(document.getElementById('modalImportaConsegne'));
      mi.show();
    }

    function eseguiImportaConsegne() {
      var url = (document.getElementById('inputImportaUrl').value || '').trim();
      var pwd = (document.getElementById('inputImportaPwd').value || '').trim();

      if (!url) { Swal.fire({ icon: 'warning', title: 'URL mancante', text: 'Inserisci il link al Google Doc.', confirmButtonColor: '#e65100' }); return; }
      if (!pwd) { Swal.fire({ icon: 'warning', title: 'Password mancante', text: 'Inserisci la password di autorizzazione.', confirmButtonColor: '#e65100' }); return; }

      // Estrai document ID dall'URL
      var m = url.match(/\/d\/([a-zA-Z0-9_\-]+)/);
      var docId = m ? m[1] : url;
      if (!docId || docId.length < 10) { Swal.fire({ icon: 'warning', title: 'URL non valido', text: 'Non riesco a estrarre l\'ID documento dall\'URL fornito.', confirmButtonColor: '#e65100' }); return; }

      // Nascondi footer, mostra spinner
      document.getElementById('importaStep1').style.display = 'none';
      document.getElementById('importaFooter').style.display = 'none';
      document.getElementById('importaStep2').style.display = '';
      document.getElementById('importaStepMsg').textContent = 'Verifica password...';

      // ── Verifica password su Supabase prima di procedere ─────────────────────
      _q(_sb.from('impostazioni').select('valore').eq('chiave', 'IMPORT_PWD').maybeSingle())
        .then(function(row) {
          var pwdCorretta = row ? row.valore : null;
          if (!pwdCorretta || pwd !== pwdCorretta) {
            _resetModalImporta();
            Swal.fire({ icon: 'error', title: 'Password errata', text: 'La password inserita non è corretta. Impossibile procedere con l\'importazione.', confirmButtonColor: '#d33' });
            return;
          }
          // Password corretta → richiedi token Docs e procedi
          document.getElementById('importaStep2').style.display = '';
          document.getElementById('importaStepMsg').textContent = 'Richiesta accesso documento...';
          if (_googleDocsToken) {
            _eseguiParsingEImport(docId);
          } else {
            _richiediDocsToken(function(ok) {
              if (!ok) {
                _resetModalImporta();
                Swal.fire({ icon: 'error', title: 'Accesso negato', text: 'Non è stato possibile ottenere l\'accesso al documento Google. Riprova.', confirmButtonColor: '#d33' });
                return;
              }
              _eseguiParsingEImport(docId);
            });
          }
        })
        .catch(function() {
          _resetModalImporta();
          Swal.fire({ icon: 'error', title: 'Errore verifica', text: 'Impossibile verificare la password. Controlla la connessione e riprova.', confirmButtonColor: '#d33' });
        });
    }

    function _richiediDocsToken(callback) {
      if (typeof google === 'undefined' || !google.accounts || !google.accounts.oauth2) { callback(false); return; }
      // Controlla sessionStorage
      try {
        var sess = JSON.parse(sessionStorage.getItem('appSession') || '{}');
        if (sess.docsToken && sess.docsExpiry > Date.now()) {
          _googleDocsToken = sess.docsToken;
          callback(true); return;
        }
      } catch(e) {}

      _docsTokenClient = google.accounts.oauth2.initTokenClient({
        client_id: GOOGLE_CLIENT_ID,
        scope: 'https://www.googleapis.com/auth/documents.readonly',
        callback: function(resp) {
          if (resp.error || !resp.access_token) { callback(false); return; }
          _googleDocsToken = resp.access_token;
          try {
            var sess = JSON.parse(sessionStorage.getItem('appSession') || '{}');
            sess.docsToken  = resp.access_token;
            sess.docsExpiry = Date.now() + 3500000;
            sessionStorage.setItem('appSession', JSON.stringify(sess));
          } catch(e) {}
          callback(true);
        }
      });
      // prompt:'consent' forza sempre la schermata di autorizzazione Google,
      // così l'utente vede esplicitamente che sta concedendo l'accesso ai Docs.
      // Necessario la prima volta; le volte successive sarà già in cache.
      _docsTokenClient.requestAccessToken({ prompt: 'consent' });
    }

    function _resetModalImporta() {
      document.getElementById('importaStep1').style.display = '';
      document.getElementById('importaStep2').style.display = 'none';
      document.getElementById('importaStep3').style.display = 'none';
      document.getElementById('importaFooter').style.display = '';
    }

    function _eseguiParsingEImport(docId) {
      document.getElementById('importaStepMsg').textContent = 'Lettura documento in corso...';

      // Diagnostica: verifica che il token sia presente
      if (!_googleDocsToken) {
        _resetModalImporta();
        Swal.fire({ icon: 'error', title: 'Token mancante', text: 'Non è stato ottenuto il token per leggere il documento. Riprova: al prossimo tentativo potrebbe comparire la schermata di autorizzazione Google.', confirmButtonColor: '#d33' });
        return;
      }

      fetch('https://docs.googleapis.com/v1/documents/' + docId, {
        headers: { 'Authorization': 'Bearer ' + _googleDocsToken }
      })
      .then(function(r) {
        if (!r.ok) {
          // Legge il body dell'errore per mostrare un messaggio preciso
          return r.json().catch(function() { return {}; }).then(function(body) {
            var errMsg = (body.error && body.error.message) ? body.error.message : '';
            if (r.status === 401) throw new Error('Token non valido o scaduto (401). Ricarica la pagina e riprova.');
            if (r.status === 403) throw new Error('Accesso negato al documento (403). ' +
              (errMsg || 'Verifica che il documento sia condiviso con medicinadurgenza.ucsc@gmail.com o accessibile con questo account.'));
            if (r.status === 404) throw new Error('Documento non trovato (404). Controlla che l\'URL sia corretto e che il documento esista.');
            throw new Error('Errore HTTP ' + r.status + (errMsg ? ': ' + errMsg : '') + '. Verifica che il documento sia accessibile con questo account.');
          });
        }
        return r.json();
      })
      .then(function(doc) {
        document.getElementById('importaStepMsg').textContent = 'Parsing tabelle...';
        var schedeLetto = _imp_parseDocumento(doc);
        if (schedeLetto.length === 0) throw new Error('Nessuna scheda letto trovata nel documento. Controlla che il formato sia corretto.');

        document.getElementById('importaStepMsg').textContent = 'Salvataggio su database (' + schedeLetto.length + ' letti)...';
        return _imp_salvaLetti(schedeLetto);
      })
      .then(function(riepilogo) {
        document.getElementById('importaStep2').style.display = 'none';
        document.getElementById('importaStep3').style.display = '';
        var html = '<div class="alert alert-success mb-2"><i class="bi bi-check-circle-fill me-2"></i><strong>' + riepilogo.importati + ' letti importati</strong> con successo.</div>';
        if (riepilogo.saltati.length > 0) {
          html += '<div class="alert alert-warning mb-2" style="font-size:0.82rem;"><strong>Saltati (' + riepilogo.saltati.length + '):</strong> ' + riepilogo.saltati.join(', ') + '</div>';
        }
        if (riepilogo.errori.length > 0) {
          html += '<div class="alert alert-danger mb-2" style="font-size:0.82rem;"><strong>Errori (' + riepilogo.errori.length + '):</strong> ' + riepilogo.errori.join('; ') + '</div>';
        }
        document.getElementById('importaRisultato').innerHTML = html;
        document.getElementById('importaFooter').innerHTML =
          '<button type="button" class="btn btn-secondary" data-bs-dismiss="modal" onclick="_sincronizzaEPoiFai(function(){})">Chiudi e aggiorna</button>';
        document.getElementById('importaFooter').style.display = '';
      })
      .catch(function(e) {
        _resetModalImporta();
        Swal.fire({ icon: 'error', title: 'Errore importazione', text: e.message || String(e), confirmButtonColor: '#d33' });
      });
    }

    // ── Parsing del documento Google Docs API v1 ──────────────────────────────

    function _imp_parseDocumento(doc) {
      var schedeLetto = [];
      if (!doc || !doc.body || !doc.body.content) return schedeLetto;

      var content = doc.body.content;

      // ── 1. Parsing tabelle (schede letto) ────────────────────────────────────
      content.forEach(function(elem) {
        if (!elem.table) return;
        var table = elem.table;
        var rows  = table.tableRows || [];
        if (rows.length < 1) return;

        try {
          var dati = _imp_parseSchedaLetto(rows);
          if (dati && dati.Letto && String(dati.Letto).trim()) {
            schedeLetto.push(dati);
          }
        } catch(e) {
          console.warn('[Import] Tabella saltata:', e.message);
        }
      });

      // ── 2. Sezione "PENDENTI POST-DIMISIONE NON CANCELLARE" ──────────────────
      // Cerca il paragrafo con questa intestazione e raccoglie tutto il contenuto
      // successivo (fino a fine documento) nel letto NOTE → campo Diaria.
      var pendentiLines = [];
      var inPendenti    = false;

      content.forEach(function(elem) {
        if (!elem.paragraph) return; // salta tabelle e altri elementi

        var para = elem.paragraph;
        var testo = _imp_paraGetText(para).trim();

        if (!inPendenti) {
          // Cerca l'intestazione (anche parziale / con variazioni di capitalizzazione)
          if (/PENDENTI\s+POST.{0,6}DIMISS?ION/i.test(testo)) {
            inPendenti = true;
            // L'intestazione stessa NON va inserita nel contenuto
          }
          return;
        }

        // Siamo dentro la sezione: aggiungi riga preservando HTML
        pendentiLines.push(_imp_paraToHtml(para));
      });

      if (pendentiLines.length > 0) {
        // Rimuovi righe vuote iniziali e finali
        while (pendentiLines.length > 0 && pendentiLines[0] === '') pendentiLines.shift();
        while (pendentiLines.length > 0 && pendentiLines[pendentiLines.length - 1] === '') pendentiLines.pop();

        schedeLetto.push({
          Letto:  'NOTE',
          Diaria: pendentiLines.join('<br>')
        });
      }

      return schedeLetto;
    }

    function _imp_parseSchedaLetto(rows) {
      var row0   = rows[0];
      var cells0 = row0.tableCells || [];
      var dati   = {};

      if (cells0.length >= 1) {
        var sin = _imp_parseColonnaSinistra(cells0[0]);
        for (var k in sin) dati[k] = sin[k];
      }
      if (cells0.length >= 2) {
        var cen = _imp_parseColonnaCentrale(cells0[1]);
        for (var k in cen) dati[k] = cen[k];
      }
      if (cells0.length >= 3) {
        dati.DaFare = _imp_cellToHtml(cells0[2]);
      }

      // Riga PIANO DI CURA (seconda riga della tabella)
      if (rows.length >= 2) {
        var cells1 = rows[1].tableCells || [];
        var pianoCell = null;
        for (var c = 0; c < cells1.length; c++) {
          var ct = _imp_cellGetText(cells1[c]).trim().toUpperCase();
          if (ct.indexOf('PIANO') === -1 && ct !== '') { pianoCell = cells1[c]; break; }
        }
        if (!pianoCell && cells1.length >= 2) pianoCell = cells1[1];
        if (pianoCell) dati.PianoTerapeutico = _imp_cellToHtml(pianoCell);
      }
      return dati;
    }

    function _imp_parseColonnaSinistra(cell) {
      var result = { Letto:'', DataRicovero:'', DataNascita:'', CodiceSanitario:'', Allergie:'', Ossigeno:'', NoteTerapia:'' };
      var paras = _imp_getCellParas(cell);
      var noteParts = [];
      var bedFound  = false;

      for (var i = 0; i < paras.length; i++) {
        var text = _imp_paraGetText(paras[i]).trim();
        if (!text) { if (bedFound && noteParts.length > 0) noteParts.push(''); continue; }

        if (!bedFound) {
          result.Letto = text.replace(/^l(?:etto)?\.?\s*/i, '').trim();
          bedFound = true; continue;
        }

        var lower = text.toLowerCase();

        // Sesso → salta
        if (/^[mf]\.?$/i.test(text) || /^(maschio|femmina|uomo|donna)$/i.test(lower)) continue;

        // Ingresso / data ricovero
        if (lower.indexOf('ingresso') !== -1) {
          // 1) Data completa dd/mm/yyyy (o con - o .)
          var dm = text.match(/\b\d{1,2}[\-\/\.]\d{1,2}[\-\/\.]\d{2,4}\b/);
          if (dm) {
            result.DataRicovero = dm[0];
          } else {
            // 2) Data parziale dd/mm senza anno → aggiunge anno corrente
            var dmParz = text.match(/\b(\d{1,2})[\-\/\.](\d{1,2})\b/);
            if (dmParz) {
              result.DataRicovero = dmParz[1] + '/' + dmParz[2] + '/' + new Date().getFullYear();
            } else {
              var after = text.replace(/ingresso\s*[:=]?\s*/i,'').trim();
              if (after) result.DataRicovero = after;
              else if (i+1 < paras.length) {
                var nt = _imp_paraGetText(paras[i+1]).trim();
                if (nt && /\d/.test(nt)) { result.DataRicovero = nt; i++; }
              }
            }
          }
          continue;
        }

        // Codice Sanitario
        if (/^c\.?\s*s\.?\s*[:=]?\s*\S/i.test(text) || /^codice\s+sanitario\s*[:=]?\s*\S/i.test(text)) {
          result.CodiceSanitario = text.replace(/^(c\.?\s*s\.?\s*[:=]?\s*|codice\s+sanitario\s*[:=]?\s*)/i,'').trim(); continue;
        }
        if (/^c\.?\s*s\.?$/i.test(text) || /^codice\s+sanitario$/i.test(text)) {
          if (i+1 < paras.length) { result.CodiceSanitario = _imp_paraGetText(paras[i+1]).trim(); i++; } continue;
        }

        // Allergie — prende solo il testo DOPO "Allergie:", "Allergia:", "Allergie", "Allergia"
        if (/^allergi[ae]/i.test(text)) {
          // Cerca il primo ':' nel testo (fuori dai tag HTML) per separare etichetta da contenuto
          var rawHtml = _imp_paraToHtml(paras[i]);
          // Trova la posizione del primo ':' nel testo (non all'interno di tag HTML)
          var colonPosHtml = -1;
          var inTag = false;
          for (var ci = 0; ci < rawHtml.length; ci++) {
            if (rawHtml[ci] === '<') { inTag = true; continue; }
            if (rawHtml[ci] === '>') { inTag = false; continue; }
            if (!inTag && rawHtml[ci] === ':') { colonPosHtml = ci; break; }
          }
          var avText = text.indexOf(':') !== -1
            ? text.substring(text.indexOf(':') + 1).trim()
            : text.replace(/^allergi[ae]\s*/i, '').trim();
          if (avText) {
            // Prende l'HTML dopo il ':' (preserva formattazione colori/grassetto)
            result.Allergie = colonPosHtml !== -1
              ? rawHtml.substring(colonPosHtml + 1).trim()
              : rawHtml.replace(/^allergi[ae]\s*/i, '').trim();
          } else if (i+1 < paras.length && _imp_paraGetText(paras[i+1]).trim()) {
            result.Allergie = _imp_paraToHtml(paras[i+1]); i++;
          }
          continue;
        }

        // DDN / Data nascita
        if (/^ddn\s*[:=]?\s*/i.test(text)) {
          var ddv = text.replace(/^ddn\s*[:=]?\s*/i,'').trim();
          if (ddv) result.DataNascita = ddv;
          else if (i+1 < paras.length) { result.DataNascita = _imp_paraGetText(paras[i+1]).trim(); i++; }
          continue;
        }

        // Ossigeno
        if (/^ossigeno/i.test(text)) {
          var ov = text.replace(/^ossigeno\s*[:=]?\s*/i,'').trim();
          if (ov) result.Ossigeno = ov;
          else if (i+1 < paras.length) { result.Ossigeno = _imp_paraGetText(paras[i+1]).trim(); i++; }
          continue;
        }

        // Tutto il resto → Note e Terapia
        noteParts.push(_imp_paraToHtml(paras[i]));
      }

      while (noteParts.length > 0 && noteParts[0] === '') noteParts.shift();
      while (noteParts.length > 0 && noteParts[noteParts.length-1] === '') noteParts.pop();
      result.NoteTerapia = noteParts.join('<br>');
      return result;
    }

    function _imp_parseColonnaCentrale(cell) {
      var result = { Nome:'', Diagnosi:'', Diaria:'' };
      var paras  = _imp_getCellParas(cell);
      var phase  = 'name';
      var diagP  = [], diarP = [];

      for (var i = 0; i < paras.length; i++) {
        var text = _imp_paraGetText(paras[i]).trim();

        if (phase === 'name') {
          if (!text) continue;
          var ci = text.indexOf(',');
          result.Nome = (ci !== -1 ? text.substring(0, ci) : text).trim();
          phase = 'diagnosi'; continue;
        }
        if (phase === 'diagnosi') {
          if (!text) { if (diagP.length > 0) phase = 'diaria'; continue; }
          var bold = _imp_isParaBold(paras[i]);
          if (bold) { diagP.push(_imp_paraToHtml(paras[i])); }
          else {
            if (diagP.length === 0) diagP.push(_imp_paraToHtml(paras[i]));
            else { phase = 'diaria'; diarP.push(_imp_paraToHtml(paras[i])); }
          }
          continue;
        }
        if (phase === 'diaria') {
          diarP.push(text ? _imp_paraToHtml(paras[i]) : '');
        }
      }
      while (diarP.length > 0 && diarP[0] === '') diarP.shift();
      while (diarP.length > 0 && diarP[diarP.length-1] === '') diarP.pop();
      result.Diagnosi = diagP.join('<br>');
      result.Diaria   = diarP.join('<br>');
      return result;
    }

    // ── Utility Docs API v1 ───────────────────────────────────────────────────

    function _imp_getCellParas(cell) {
      var paras = [];
      (cell.content || []).forEach(function(elem) {
        if (elem.paragraph) paras.push(elem.paragraph);
      });
      return paras;
    }

    function _imp_paraGetText(para) {
      return (para.elements || []).map(function(el) {
        return (el.textRun && el.textRun.content) ? el.textRun.content : '';
      }).join('').replace(/\n$/, '');
    }

    function _imp_cellGetText(cell) {
      return _imp_getCellParas(cell).map(_imp_paraGetText).join('\n');
    }

    function _imp_cellToHtml(cell) {
      var parts = _imp_getCellParas(cell).map(_imp_paraToHtml);
      while (parts.length > 0 && parts[0] === '') parts.shift();
      while (parts.length > 0 && parts[parts.length-1] === '') parts.pop();
      return parts.join('<br>');
    }

    function _imp_isParaBold(para) {
      var els = para.elements || [];
      for (var i = 0; i < els.length; i++) {
        var tr = els[i].textRun;
        if (tr && tr.content && tr.content.trim()) {
          return !!(tr.textStyle && tr.textStyle.bold);
        }
      }
      return false;
    }

    function _imp_paraToHtml(para) {
      var html = '';
      (para.elements || []).forEach(function(el) {
        if (!el.textRun) return;
        var content = (el.textRun.content || '').replace(/\n$/, '');
        if (!content) return;
        var ts  = el.textRun.textStyle || {};
        var esc = content.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
        // Background color (evidenziazione)
        var bg = ts.backgroundColor && ts.backgroundColor.color && ts.backgroundColor.color.rgbColor;
        if (bg) {
          var r = Math.round((bg.red   || 0) * 255);
          var g = Math.round((bg.green || 0) * 255);
          var b = Math.round((bg.blue  || 0) * 255);
          esc = '<span style="background-color:rgb(' + r + ',' + g + ',' + b + ')">' + esc + '</span>';
        }
        if (ts.underline) esc = '<u>' + esc + '</u>';
        if (ts.italic)    esc = '<i>' + esc + '</i>';
        if (ts.bold)      esc = '<b>' + esc + '</b>';
        html += esc;
      });
      return html;
    }

    // ── Salvataggio su Supabase ───────────────────────────────────────────────

    function _imp_salvaLetti(schedeLetto) {
      var importati = [], saltati = [], errori = [];
      return new Promise(function(resolve) {
        var pending = schedeLetto.length;
        if (pending === 0) { resolve({ importati: 0, saltati: [], errori: [] }); return; }
        schedeLetto.forEach(function(dati) {
          var letto = String(dati.Letto).trim();
          _sbImportaLetto(letto, dati)
            .then(function(ok) {
              if (ok) importati.push(letto);
              else saltati.push(letto + ' (letto non trovato nel DB)');
            })
            .catch(function(e) { errori.push(letto + ': ' + (e.message || 'errore')); })
            .finally(function() {
              pending--;
              if (pending === 0) resolve({ importati: importati.length, saltati: saltati, errori: errori });
            });
        });
      });
    }
