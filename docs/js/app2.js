// ================================================================
// app2.js — Logica secondaria (da Scripts2.html)
// Modifiche: aggiunto handler modalRinomina per popolare
//            #inputNuovoNome via API ottieniNomeReparto
// ================================================================


    // ============================================================
    // AGGIORNAMENTO DINAMICO SELECT NEI MODAL (sempre freschi al DOM corrente)
    // ============================================================
    document.addEventListener('show.bs.modal', function(e) {
      var id = e.target && e.target.id;
      var cards = Array.from(document.querySelectorAll('#cardsContainer .patient-card'));

      if (id === 'modalEliminaLetto') {
        var sel = document.getElementById('selectEliminaLetto');
        if (!sel) return;
        sel.innerHTML = '<option value="">--</option>';
        cards.forEach(function(card) {
          var letto = card.getAttribute('data-bed');
          if (letto === 'NOTE') return; // la scheda NOTE non può essere eliminata
          var opt = document.createElement('option');
          opt.value = letto;
          opt.textContent = 'Letto ' + letto;
          sel.appendChild(opt);
        });
      }

      if (id === 'modalDimetti') {
        var sel = document.getElementById('selectDimettiLetto');
        if (!sel) return;
        sel.innerHTML = '<option value="">Caricamento...</option>';
        sel.disabled = true;
        google.script.run
          .withSuccessHandler(function(letti) {
            sel.innerHTML = '<option value="">-- Seleziona letto --</option>';
            letti.forEach(function(l) {
              var opt = document.createElement('option');
              opt.value = l.letto;
              opt.textContent = l.letto + ' - ' + (l.nome || '(vuoto)') + ' - ' + l.tipologia;
              sel.appendChild(opt);
            });
            sel.disabled = false;
          })
          .withFailureHandler(function() {
            sel.innerHTML = '<option value="">Errore caricamento</option>';
            sel.disabled = false;
          })
          .getLettiFull();
      }

      if (id === 'modalSposta') {
        var selO = document.getElementById('selectSpostaOrigine');
        var selD = document.getElementById('selectSpostaDestinazione');
        if (!selO || !selD) return;
        var prevO = selO.value, prevD = selD.value;
        [selO, selD].forEach(function(s) { s.innerHTML = '<option value="">Caricamento...</option>'; s.disabled = true; });
        google.script.run
          .withSuccessHandler(function(letti) {
            selO.innerHTML = '<option value="">-- Letto origine --</option>';
            selD.innerHTML = '<option value="">-- Letto destinazione --</option>';
            letti.forEach(function(l) {
              var label = l.letto + ' - ' + (l.nome || '(vuoto)') + ' - ' + l.tipologia;
              [selO, selD].forEach(function(s) {
                var opt = document.createElement('option');
                opt.value = l.letto;
                opt.textContent = label;
                s.appendChild(opt);
              });
            });
            selO.disabled = false; selD.disabled = false;
            if (prevO) selO.value = prevO;
            if (prevD) selD.value = prevD;
          })
          .withFailureHandler(function() {
            [selO, selD].forEach(function(s) { s.innerHTML = '<option value="">Errore caricamento</option>'; s.disabled = false; });
          })
          .getLettiFull();
      }

      if (id === 'modalGestisciTipologie') { _gtApri(); }
      if (id === 'modalCambiaTipologiaLetto') { _ctlApri(); }

      // Popola inputNuovoNome con il nome attuale del reparto
      if (id === 'modalRinomina') {
        var inp = document.getElementById('inputNuovoNome');
        if (!inp) return;
        var navName = document.getElementById('nav-app-name');
        if (navName && navName.innerText) {
          inp.value = navName.innerText;
        } else {
          google.script.run
            .withSuccessHandler(function(res) {
              var nome = (res && res.nome) ? res.nome : (typeof res === 'string' ? res : '');
              if (inp) inp.value = nome;
            })
            .ottieniNomeReparto();
        }
      }
    });


    // ============================================================
    // TOGGLE SYNC
    // ============================================================
    function toggleSync() {
      if (!_syncPaused) {
        Swal.fire({
          icon: 'question',
          title: 'Metti in pausa il sync?',
          text: 'Il sync automatico verrà  sospeso. Potrai riattivarlo dal menu Impostazioni.',
          showCancelButton: true,
          confirmButtonColor: '#f0a500',
          confirmButtonText: 'Sì, metti in pausa',
          cancelButtonText: 'Annulla'
        }).then(function(result) {
          if (result.isConfirmed) {
            _syncPaused = true;
            _aggiornaVoceSync();
            var ind = document.getElementById('syncIndicator');
            var st  = document.getElementById('syncStatus');
            if (ind) ind.className = 'badge bg-warning text-dark ms-2';
            if (st)  st.innerText  = 'In pausa';
          }
        });
      } else {
        Swal.fire({
          icon: 'question',
          title: 'Riattivare il sync?',
          text: 'Il sync automatico verrà  riattivato.',
          showCancelButton: true,
          confirmButtonColor: '#198754',
          confirmButtonText: 'Sì, riattiva',
          cancelButtonText: 'Annulla'
        }).then(function(result) {
          if (result.isConfirmed) {
            _syncPaused = false;
            _aggiornaVoceSync();
          }
        });
      }
    }

    function _aggiornaVoceSync() {
      var icon  = document.getElementById('iconToggleSync');
      var label = document.getElementById('labelToggleSync');
      if (_syncPaused) {
        if (icon)  { icon.className  = 'bi bi-play-circle text-success me-2'; }
        if (label) { label.innerText = 'Riattiva il sync'; }
      } else {
        if (icon)  { icon.className  = 'bi bi-pause-circle text-warning me-2'; }
        if (label) { label.innerText = 'Metti in pausa il sync'; }
      }
    }


    function apriImpostazioniArchivio() {
      var modalEl = document.getElementById('modalImpostazioniArchivio');
      if (!modalEl) return;
      var input = document.getElementById('inputGiorniArchivio');
      if (input) input.value = '';
      google.script.run
        .withSuccessHandler(function(giorni) {
          if (input) input.value = giorni || 30;
          bootstrap.Modal.getOrCreateInstance(modalEl).show();
        })
        .withFailureHandler(function() {
          if (input) input.value = 30;
          bootstrap.Modal.getOrCreateInstance(modalEl).show();
        })
        .getGiorniArchivio();
    }

    function salvaImpostazioniArchivio() {
      var input    = document.getElementById('inputGiorniArchivio');
      var feedback = document.getElementById('archivioFeedback');
      var val = parseInt(input ? input.value : '', 10);
      if (!val || val < 1) { if (feedback) feedback.style.display = 'block'; return; }
      if (feedback) feedback.style.display = 'none';
      var modalEl = document.getElementById('modalImpostazioniArchivio');
      var mi = bootstrap.Modal.getInstance(modalEl); if (mi) mi.hide();
      _opStart('Salvataggio impostazioni...');
      google.script.run
        .withSuccessHandler(function(res) {
          _opEnd();
          if (res && res.success) {
            Swal.fire({ icon: 'success', title: 'Impostazioni salvate',
              text: 'Le consegne verranno conservate per ' + res.giorni + ' giorni.',
              timer: 2500, showConfirmButton: false, timerProgressBar: true });
          } else {
            Swal.fire({ icon: 'error', title: 'Errore', text: (res && res.message) || 'Salvataggio fallito.' });
          }
        })
        .withFailureHandler(function(err) {
          _opEnd();
          Swal.fire({ icon: 'error', title: 'Errore', text: err.message || 'Salvataggio fallito.' });
        })
        .salvaGiorniArchivio(val);
    }


    // ============================================================
    // TIPOLOGIE — cache colori
    // ============================================================
    var _coloriTipologie = {};

    function _getColoreTipo(tipo) {
      if (!tipo) return '#adb5bd';
      if (_coloriTipologie[tipo]) return _coloriTipologie[tipo];
      return stringToColor(tipo);
    }
    // Esposto globalmente: app.js lo chiama dopo ogni aggiornamento DOM
    window._getColoreTipo = _getColoreTipo;

    function _aggiornaBadgePrincipali() {
      document.querySelectorAll('[id^="badge-tipo-"]').forEach(function(b) {
        var letto = b.id.replace('badge-tipo-alt-', '').replace('badge-tipo-', '');
        var card = document.querySelector('.patient-card[data-bed="' + letto + '"]');
        if (!card) return;
        var tipo = (card.getAttribute('data-tipologia') || '').trim();
        if (tipo) { b.innerText = tipo; b.style.backgroundColor = _getColoreTipo(tipo); }
      });
    }
    // Esposto globalmente: app.js lo chiama dopo ogni aggiornamento DOM
    window._aggiornaBadgePrincipali = _aggiornaBadgePrincipali;

    function _caricaColoriTipologie(callback) {
      google.script.run
        .withSuccessHandler(function(arr) {
          if (arr) {
            arr.forEach(function(item) {
              _coloriTipologie[item.nome] = item.colore || null;
            });
          }
          _aggiornaBadgePrincipali();
          if (typeof callback === 'function') callback();
        })
        .withFailureHandler(function() {
          if (typeof callback === 'function') callback();
        })
        .getTipologieConfigurate();
    }

    // ============================================================
    // PALETTE COLORI (condivisa)
    // ============================================================
    // Color picker Pickr: si posiziona vicino al swatch, spettro completo
    var _pickrInstance = null;

    function _apriPaletteSuSwatch(swatchEl, onSelect) {
      // Distruggi eventuale istanza precedente
      if (_pickrInstance) {
        try { _pickrInstance.destroyAndRemove(); } catch(e) {}
        _pickrInstance = null;
      }

      // Normalizza il colore iniziale in hex
      var bgColor = swatchEl.style.background || swatchEl.style.backgroundColor || '#78909c';
      if (!bgColor.match(/^#[0-9a-fA-F]{6}$/)) bgColor = '#78909c';

      _pickrInstance = Pickr.create({
        el: swatchEl,
        useAsButton: true,
        theme: 'nano',
        default: bgColor,
        position: 'bottom-start',
        components: {
          preview: true,
          opacity: false,
          hue: true,
          interaction: {
            hex: true,
            input: true,
            save: true
          }
        },
        i18n: {
          'btn:save': 'Conferma',
          'btn:cancel': 'Annulla'
        }
      });

      _pickrInstance.on('change', function(color) {
        var hex = color.toHEXA().toString().slice(0, 7);
        onSelect(hex);
        swatchEl.style.background = hex;
      });

      _pickrInstance.on('save', function(color) {
        var hex = color.toHEXA().toString().slice(0, 7);
        onSelect(hex);
        try { _pickrInstance.destroyAndRemove(); } catch(e) {}
        _pickrInstance = null;
      });

      _pickrInstance.on('cancel', function() {
        try { _pickrInstance.destroyAndRemove(); } catch(e) {}
        _pickrInstance = null;
      });

      _pickrInstance.show();
    }

    // ============================================================
    // MODAL 1 — GESTISCI TIPOLOGIE
    // ============================================================
    var _gtDirty = false;
    var _gtRighe = [];

    function _gtApri() {
      _gtDirty = false;
      _gtRighe = [];
      var container = document.getElementById('gtListContainer');
      if (container) {
        container.innerHTML =
          '<div class="text-center py-4">' +
            '<div class="spinner-border text-primary" style="width:2rem;height:2rem" role="status"></div>' +
            '<p class="text-muted small mt-2 mb-0">Caricamento tipologie...</p>' +
          '</div>';
      }
      _caricaColoriTipologie(function() {
        google.script.run
          .withSuccessHandler(function(arr) {
            _gtRighe = (arr || []).map(function(item) {
              return { nomeOld: item.nome, nomeNew: item.nome,
                       colore: item.colore || stringToColor(item.nome), isNew: false };
            });
            _gtRenderLista();
          })
          .withFailureHandler(function() { _gtRenderLista(); })
          .getTipologieConfigurate();
      });
    }

    function _gtRenderLista() {
      var container = document.getElementById('gtListContainer');
      if (!container) return;
      if (_gtRighe.length === 0) {
        container.innerHTML = '<p class="text-muted small fst-italic text-center py-3 mb-0">' +
          '<i class="bi bi-info-circle me-1"></i>Nessuna tipologia configurata. Il default è <strong>STANDARD</strong>.</p>';
        return;
      }
      var html = '';
      _gtRighe.forEach(function(riga, idx) {
        var color = riga.colore || stringToColor(riga.nomeNew || 'X');
        html += '<div class="d-flex align-items-center gap-2 px-2 py-2 border-bottom gt-riga" data-idx="' + idx + '">' +
          '<div class="gt-swatch" data-idx="' + idx + '" ' +
               'style="width:36px;height:36px;border-radius:8px;background:' + color + ';' +
               'border:2px solid rgba(0,0,0,.15);cursor:pointer;flex-shrink:0;position:relative" ' +
               'title="Clicca per cambiare colore"></div>' +
          '<input type="text" class="form-control form-control-sm text-uppercase gt-nome-input" ' +
                 'data-idx="' + idx + '" value="' + (riga.nomeNew || '').replace(/"/g,'&quot;') + '" ' +
                 'placeholder="Nome tipologia" ' +
                 'style="max-width:220px;font-weight:600">' +
          '<span class="badge text-white ms-1 gt-badge-preview" data-idx="' + idx + '" ' +
                'style="background:' + color + ';font-size:.75rem;min-width:70px">' + (riga.nomeNew || '') + '</span>' +
          '<button type="button" class="btn btn-link btn-sm text-danger ms-auto p-1 gt-del-btn" ' +
                  'data-idx="' + idx + '" title="Elimina tipologia"><i class="bi bi-trash3-fill"></i></button>' +
        '</div>';
      });
      container.innerHTML = html;

      container.querySelectorAll('.gt-swatch').forEach(function(sw) {
        sw.addEventListener('click', function(e) {
          e.stopPropagation();
          var i = parseInt(sw.getAttribute('data-idx'));
          _apriPaletteSuSwatch(sw, function(color) {
            _gtRighe[i].colore = color;
            _gtDirty = true;
            sw.style.background = color;
            var badge = container.querySelector('.gt-badge-preview[data-idx="'+i+'"]');
            if (badge) badge.style.background = color;
          });
        });
      });

      container.querySelectorAll('.gt-nome-input').forEach(function(inp) {
        inp.addEventListener('input', function() {
          var i = parseInt(inp.getAttribute('data-idx'));
          var val = inp.value.toUpperCase();
          inp.value = val;
          _gtRighe[i].nomeNew = val;
          _gtDirty = true;
          var badge = container.querySelector('.gt-badge-preview[data-idx="'+i+'"]');
          if (!_gtRighe[i]._colorManuale && val) {
            var suggerito = stringToColor(val);
            _gtRighe[i].colore = suggerito;
            var sw = container.querySelector('.gt-swatch[data-idx="'+i+'"]');
            if (sw) sw.style.background = suggerito;
            if (badge) badge.style.background = suggerito;
          }
          if (badge) badge.innerText = val || '';
        });
      });

      container.querySelectorAll('.gt-del-btn').forEach(function(btn) {
        btn.addEventListener('click', function() {
          var i = parseInt(btn.getAttribute('data-idx'));
          _gtEliminaRiga(i);
        });
      });
    }

    function _gtAggiungiRiga() {
      _gtRighe.push({ nomeOld: '', nomeNew: '', colore: '#78909c', isNew: true, _colorManuale: false });
      _gtDirty = true;
      _gtRenderLista();
      var container = document.getElementById('gtListContainer');
      if (container) {
        var inputs = container.querySelectorAll('.gt-nome-input');
        if (inputs.length > 0) inputs[inputs.length-1].focus();
      }
    }

    function _gtEliminaRiga(idx) {
      var riga = _gtRighe[idx];
      if (!riga) return;

      if (riga.isNew || !riga.nomeOld) {
        Swal.fire({
          icon: 'question', title: 'Rimuovere?',
          html: 'Vuoi rimuovere la tipologia <strong>' + (riga.nomeNew || 'nuova') + '</strong>?',
          showCancelButton: true, confirmButtonColor: '#dc3545',
          confirmButtonText: 'Sì, rimuovi', cancelButtonText: 'Annulla'
        }).then(function(r) {
          if (!r.isConfirmed) return;
          _gtRighe.splice(idx, 1);
          _gtDirty = true;
          _gtRenderLista();
        });
        return;
      }

      Swal.fire({
        icon: 'question', title: 'Eliminare tipologia?',
        html: 'Vuoi eliminare la tipologia <strong>' + riga.nomeOld + '</strong>?',
        showCancelButton: true, confirmButtonColor: '#dc3545',
        confirmButtonText: 'Sì, elimina', cancelButtonText: 'Annulla'
      }).then(function(confirmed) {
        if (!confirmed.isConfirmed) return;

        var rowEl = document.querySelector('.gt-riga[data-idx="' + idx + '"]');
        if (rowEl) {
          rowEl.innerHTML = '<div class="d-flex align-items-center gap-2 px-2 py-2 text-muted small">' +
            '<div class="spinner-border spinner-border-sm text-danger me-2" role="status"></div>' +
            'Verifica letti assegnati...' +
          '</div>';
        }

        google.script.run
          .withSuccessHandler(function(res) {
            if (res.success) {
              _gtRighe.splice(idx, 1);
              _gtDirty = true;
              _gtRenderLista();
            } else if (res.count > 0) {
              _gtRenderLista();
              Swal.fire({
                icon: 'warning',
                title: 'Tipologia in uso',
                html: 'La tipologia <strong>' + riga.nomeOld + '</strong> è assegnata a ' +
                      '<strong>' + res.count + '</strong> letto/i.<br><br>' +
                      'Se continui, quei letti diventeranno <strong>STANDARD</strong>.',
                showCancelButton: true,
                confirmButtonColor: '#dc3545',
                confirmButtonText: '<i class="bi bi-trash3 me-1"></i>Elimina comunque',
                cancelButtonText: 'Annulla'
              }).then(function(r2) {
                if (!r2.isConfirmed) return;
                var rowEl2 = document.querySelector('.gt-riga[data-idx="' + idx + '"]');
                if (rowEl2) {
                  rowEl2.innerHTML = '<div class="d-flex align-items-center gap-2 px-2 py-2 text-muted small">' +
                    '<div class="spinner-border spinner-border-sm text-danger me-2" role="status"></div>' +
                    'Eliminazione in corso...' +
                  '</div>';
                }
                google.script.run
                  .withSuccessHandler(function() {
                    delete _coloriTipologie[riga.nomeOld];
                    _gtRighe.splice(idx, 1);
                    _gtDirty = true;
                    _gtRenderLista();
                  })
                  .withFailureHandler(function(err) {
                    _gtRenderLista();
                    Swal.fire({ icon:'error', title:'Errore', text: err.message || 'Eliminazione fallita.' });
                  })
                  .eliminaTipologiaConfigurata(riga.nomeOld, true);
              });
            } else {
              _gtRenderLista();
              Swal.fire({ icon:'error', title:'Errore', text:'Impossibile verificare i letti assegnati.' });
            }
          })
          .withFailureHandler(function(err) {
            _gtRenderLista();
            Swal.fire({ icon:'error', title:'Errore', text: err.message || 'Operazione fallita.' });
          })
          .eliminaTipologiaConfigurata(riga.nomeOld, false);
      });
    }

    function _chiudiGestisciTipologie() {
      if (_gtDirty) {
        Swal.fire({
          icon: 'warning',
          title: 'Modifiche non salvate',
          html: 'Hai modificato le tipologie senza salvare.<br>Vuoi <strong>salvare</strong> prima di uscire?',
          showDenyButton: true,
          showCancelButton: true,
          confirmButtonText: '<i class="bi bi-save me-1"></i>Salva ed esci',
          denyButtonText: 'Esci senza salvare',
          cancelButtonText: 'Rimani',
          confirmButtonColor: '#0d6efd',
          denyButtonColor: '#6c757d'
        }).then(function(result) {
          if (result.isConfirmed) {
            _gtSalva(true);
          } else if (result.isDenied) {
            _gtDirty = false;
            bootstrap.Modal.getInstance(document.getElementById('modalGestisciTipologie')).hide();
          }
        });
      } else {
        bootstrap.Modal.getInstance(document.getElementById('modalGestisciTipologie')).hide();
      }
    }

    function _gtSalva(chiudiDopo) {
      var vuoti = _gtRighe.filter(function(r) { return !r.nomeNew || !r.nomeNew.trim(); });
      if (vuoti.length > 0) {
        Swal.fire({ icon:'warning', text:'Ci sono tipologie senza nome. Compilale o rimuovile.', timer:3000, showConfirmButton:false });
        return;
      }
      var nomi = _gtRighe.map(function(r){ return r.nomeNew.trim().toUpperCase(); });
      var unici = new Set(nomi);
      if (unici.size !== nomi.length) {
        Swal.fire({ icon:'warning', text:'Ci sono tipologie con lo stesso nome.', timer:3000, showConfirmButton:false });
        return;
      }

      var modalEl = document.getElementById('modalGestisciTipologie');
      bootstrap.Modal.getInstance(modalEl).hide();

      _opStart('Salvataggio tipologie in corso...');
      var payload = _gtRighe.map(function(r) {
        return { nomeOld: r.nomeOld || r.nomeNew, nomeNew: r.nomeNew.trim().toUpperCase(), colore: r.colore || null };
      });

      google.script.run
        .withSuccessHandler(function() {
          _gtDirty = false;
          _coloriTipologie = {};
          payload.forEach(function(p) { if (p.nomeNew) _coloriTipologie[p.nomeNew] = p.colore; });
          _caricaColoriTipologie(function() { _aggiornaBadgePrincipali(); });
          _sincronizzaEPoiFai(function() {
            _opEnd();
            Swal.fire({ icon:'success', title:'Tipologie salvate', timer:2000, showConfirmButton:false, timerProgressBar:true });
          });
        })
        .withFailureHandler(function(err) {
          _opEnd();
          Swal.fire({ icon:'error', title:'Errore', text: err.message || 'Salvataggio fallito.' });
        })
        .salvaTipologieBatch(payload);
    }

    // ============================================================
    // MODAL 2 — CAMBIA TIPOLOGIA AD UN LETTO
    // ============================================================

    function _ctlApri() {
      var selLetto = document.getElementById('ctlSelectLetto');
      var selTipo  = document.getElementById('ctlSelectTipologia');
      if (!selLetto || !selTipo) return;

      selTipo.innerHTML = '<option value="">STANDARD (nessuna)</option>';
      Object.keys(_coloriTipologie).sort().forEach(function(tipo) {
        var opt = document.createElement('option');
        opt.value = tipo; opt.textContent = tipo;
        selTipo.appendChild(opt);
      });

      selLetto.innerHTML = '<option value="">Caricamento...</option>';
      selLetto.disabled = true;
      google.script.run
        .withSuccessHandler(function(letti) {
          selLetto.innerHTML = '<option value="">-- Seleziona letto --</option>';
          letti.forEach(function(l) {
            var opt = document.createElement('option');
            opt.value = l.letto;
            opt.textContent = l.letto + ' - ' + (l.nome || '(vuoto)') + ' - ' + l.tipologia;
            selLetto.appendChild(opt);
          });
          selLetto.disabled = false;
          if (_ctlPreselezioneLetto) { selLetto.value = _ctlPreselezioneLetto; _ctlPreselezioneLetto = null; }
        })
        .withFailureHandler(function() {
          selLetto.innerHTML = '<option value="">Errore caricamento</option>';
          selLetto.disabled = false;
        })
        .getLettiFull();
    }

    var _ctlPreselezioneLetto = null;
    function _apriModalTipologia(letto) {
      _ctlPreselezioneLetto = letto;
      bootstrap.Modal.getOrCreateInstance(document.getElementById('modalCambiaTipologiaLetto')).show();
    }

    function _ctlSalva() {
      var selLetto = document.getElementById('ctlSelectLetto');
      var selTipo  = document.getElementById('ctlSelectTipologia');
      var letto = selLetto ? selLetto.value : '';
      var nuovaTipo = selTipo ? selTipo.value : '';
      if (!letto) {
        Swal.fire({ icon:'warning', text:'Seleziona un letto.', timer:2000, showConfirmButton:false });
        return;
      }
      var modalEl = document.getElementById('modalCambiaTipologiaLetto');
      bootstrap.Modal.getInstance(modalEl).hide();
      var label = nuovaTipo || 'STANDARD';
      _opServer({ barMsg: 'Aggiornamento tipologia letto in corso...', successTitle: 'Tipologia aggiornata',
        successText: 'Letto ' + letto + ' → ' + label, errorTitle: 'Errore',
        serverFn: function(onOk, onErr) {
          google.script.run
            .withSuccessHandler(function(res) {
              if (res && res.success) {
                onOk();
              } else {
                onErr((res && res.message) || 'Operazione fallita.');
              }
            })
            .withFailureHandler(function(err) { onErr(err.message || 'Operazione fallita.'); })
            .cambiaTipologiaALetto(letto, nuovaTipo);
        }
      });
    }
