'use strict';
// ═══════════════════════════════════════════════════════════════════════════
// FILTROS DE LA GUÍA (GUIA_DISEÑO_GENERAL · §03) — compartido por los dashboards
// ═══════════════════════════════════════════════════════════════════════════
// Operaciones (Todos · GT · SV), Período, Moneda con tasa oficial y Filtros
// combinados tipo Excel. Los estilos viven en assets/css/filtros-guia.css.
//
// createXF — cada columna: { key, label, single?, options:[{value,label,sub?,count?}], selected }
//   selected = null  → sin filtro (todo marcado, como Excel)
//   selected = [...] → valores elegidos
// Columnas "single" (API que no acepta listas): una opción o "(Todas)".
// El panel se ubica dentro del elemento padre del host (la barra de filtros),
// que debe estar posicionado (relative o sticky).
(function () {
  const fmtN = n => Number(n || 0).toLocaleString('es-GT');
  const esc  = v => String(v ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#039;' }[c]));
  const norm = s => String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '');
  const listOf  = s => String(s || '').split(',').map(v => v.trim()).filter(Boolean);
  const selOf   = s => (s ? listOf(s) : null);
  const joinSel = sel => (sel && sel.length ? sel.join(',') : '');

  const ICONS = {
    calendar: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="5" width="18" height="16" rx="2"/><path d="M16 3v4M8 3v4M3 10h18"/></svg>',
    reset   : '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 12a9 9 0 1 0 3-6.7L3 8"/><path d="M3 3v5h5"/></svg>',
    filter  : '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><path d="M4 6h16M7 12h10M10 18h4"/></svg>',
    chevron : '<svg class="xf-chevron" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round"><path d="m6 9 6 6 6-6"/></svg>',
  };
  function paintIcons(root = document) {
    root.querySelectorAll('.period-icon').forEach(el => { el.innerHTML = ICONS.calendar; });
    root.querySelectorAll('.period-reset, .xf-reset').forEach(el => { el.innerHTML = ICONS.reset; });
  }

  // Tipo de cambio oficial (Banco de Guatemala) vía excel-parser en Render.
  // Si no responde se usa la tasa de respaldo y se indica que no es oficial.
  const FX_URL = 'https://excel-parser-m9q8.onrender.com/api/tipo-cambio';
  async function tipoCambioOficial(respaldo = 7.7, timeoutMs = 60000) {
    const ctrl = new AbortController();
    const tid = setTimeout(() => ctrl.abort(), timeoutMs);
    try {
      const r = await fetch(`${FX_URL}?_=${Date.now()}`, { cache: 'no-store', signal: ctrl.signal });
      if (!r.ok) throw new Error(`HTTP ${r.status}`);
      const data = await r.json();
      const rate = Number(data && data.rate_gtq_per_usd);
      if (!(rate > 0)) throw new Error('sin tasa');
      return { rate, oficial: true, fecha: data.date || '', fuente: data.source || 'Banco de Guatemala' };
    } catch (error) {
      console.warn('[Filtros] Tipo de cambio oficial no disponible, se usa respaldo:', error.message);
      return { rate: respaldo, oficial: false, fecha: '', fuente: 'Tasa de referencia' };
    } finally {
      clearTimeout(tid);
    }
  }

  const openXFs = new Set();
  function closeAllXF(except) { openXFs.forEach(xf => { if (xf !== except) xf.close(); }); }
  document.addEventListener('click', e => { if (!e.target.closest('.xf')) closeAllXF(); });
  document.addEventListener('keydown', e => { if (e.key === 'Escape') closeAllXF(); });

  function createXF({ host, chipsHost, onApply, searchAllColumns = false }) {
    let columns = [];
    let draft   = {};          // key → Set (multi) | string|null (single)
    const search = {};
    host.innerHTML =
      `<button class="xf-toggle" type="button" aria-expanded="false">${ICONS.filter}<span>Filtros</span><b class="xf-count" hidden>0</b>${ICONS.chevron}</button>
       <div class="xf-panel" hidden>
         <div class="xf-columns"></div>
         <div class="xf-footer"><span class="xf-hint"></span>
           <button class="xf-btn" type="button" data-xf="clear">Limpiar todo</button>
           <button class="xf-btn xf-apply" type="button" data-xf="apply">Aplicar</button></div>
       </div>`;
    const toggle = host.querySelector('.xf-toggle'), panel = host.querySelector('.xf-panel');
    const colsEl = host.querySelector('.xf-columns'), applyBtn = host.querySelector('[data-xf="apply"]');
    const hint   = host.querySelector('.xf-hint');
    const bar    = host.parentElement;

    const colOf    = key => columns.find(c => c.key === key);
    const labelOf  = (col, v) => (col.options.find(o => o.value === v) || {}).label || v;
    const initDraft = col => col.single
      ? (col.selected && col.selected.length ? col.selected[0] : null)
      : new Set(col.selected ? col.selected : col.options.map(o => o.value));

    function visibles(col) {
      const q = norm(search[col.key] || '');
      return q ? col.options.filter(o => norm(o.label).includes(q) || norm(o.sub).includes(q) || norm(o.value).includes(q)) : col.options;
    }
    // Marcadas entre las opciones disponibles (la selección puede traer valores
    // que otra columna dejó fuera del contexto; esos no cuentan aquí).
    const markedIn = col => col.options.filter(o => draft[col.key].has(o.value)).length;
    function syncColumnHead(col) {
      const el = colsEl.querySelector(`[data-col="${CSS.escape(col.key)}"]`); if (!el) return;
      const d = draft[col.key];
      el.querySelector('.xf-col-count').textContent = col.single
        ? (d ? '1' : 'Todas')
        : `${fmtN(markedIn(col))}/${fmtN(col.options.length)}`;
      const all = el.querySelector('[data-xf-all]');
      if (all && !col.single) {
        const vis = visibles(col);
        const marked = vis.filter(o => d.has(o.value)).length;
        all.checked = vis.length > 0 && marked === vis.length;
        all.indeterminate = marked > 0 && marked < vis.length;
      }
      const empty = columns.filter(c => !c.single && draft[c.key] && c.options.length && markedIn(c) === 0);
      applyBtn.disabled = empty.length > 0;
      hint.textContent = empty.length ? `Elige al menos una opción en ${empty.map(c => c.label).join(', ')}` : '';
    }
    const optionHtml = (type, name, o, checked) =>
      `<label class="xf-option"><input type="${type}"${name ? ` name="${name}"` : ''} value="${esc(o.value)}" ${checked ? 'checked' : ''}><span class="xf-opt-text">${esc(o.label)}${o.sub ? `<small>${esc(o.sub)}</small>` : ''}</span>${o.count != null ? `<em>${fmtN(o.count)}</em>` : ''}</label>`;
    function renderList(col) {
      const el = colsEl.querySelector(`[data-col="${CSS.escape(col.key)}"] .xf-list`); if (!el) return;
      const vis = visibles(col), d = draft[col.key], q = search[col.key];
      let html = '';
      if (col.single) {
        const name = `xf-${col.key}-${Math.random().toString(36).slice(2, 7)}`;
        html += `<label class="xf-option xf-all"><input type="radio" name="${name}" value="" ${d ? '' : 'checked'}> (Todas)</label>`;
        html += vis.map(o => optionHtml('radio', name, o, d === o.value)).join('');
      } else {
        if (vis.length) html += `<label class="xf-option xf-all"><input type="checkbox" data-xf-all> (Seleccionar todo${q ? ' lo visible' : ''})</label>`;
        html += vis.map(o => optionHtml('checkbox', '', o, d.has(o.value))).join('');
      }
      if (!vis.length) html += '<div class="xf-empty">Sin coincidencias</div>';
      el.innerHTML = html;
      syncColumnHead(col);
    }
    function renderPanel() {
      colsEl.innerHTML = columns.map(col => `
        <div class="xf-col" data-col="${esc(col.key)}">
          <h5>${esc(col.label)} <small class="xf-col-count"></small></h5>
          ${searchAllColumns || col.options.length > 6 ? `<input class="xf-search" type="search" placeholder="Buscar…" aria-label="Buscar en ${esc(col.label)}" value="${esc(search[col.key] || '')}">` : ''}
          <div class="xf-list"></div>
        </div>`).join('');
      columns.forEach(renderList);
    }
    colsEl.addEventListener('input', e => {
      if (!e.target.classList.contains('xf-search')) return;
      const col = colOf(e.target.closest('[data-col]').dataset.col);
      search[col.key] = e.target.value;
      renderList(col);
    });
    colsEl.addEventListener('change', e => {
      const input = e.target, col = colOf(input.closest('[data-col]')?.dataset.col);
      if (!col || input.classList.contains('xf-search')) return;
      if (col.single) { draft[col.key] = input.value || null; syncColumnHead(col); return; }
      const d = draft[col.key];
      if (input.hasAttribute('data-xf-all')) {
        visibles(col).forEach(o => (input.checked ? d.add(o.value) : d.delete(o.value)));
        renderList(col);
        return;
      }
      input.checked ? d.add(input.value) : d.delete(input.value);
      syncColumnHead(col);   // no se repinta la lista: conserva el scroll
    });

    function selectionFromDraft() {
      const out = {};
      columns.forEach(col => {
        const d = draft[col.key];
        if (col.single) out[col.key] = d ? [d] : null;
        else out[col.key] = markedIn(col) === col.options.length ? null : [...d];
      });
      return out;
    }
    function currentSelection() {
      const out = {}; columns.forEach(col => { out[col.key] = col.selected; }); return out;
    }

    const api = {
      open() {
        closeAllXF(api);
        columns.forEach(col => { draft[col.key] = initDraft(col); search[col.key] = ''; });
        renderPanel();
        bar?.classList.add('xf-open');
        panel.hidden = false; toggle.setAttribute('aria-expanded', 'true'); openXFs.add(api);
        // Alineado con el botón; se corre a la izquierda solo si no cabe en la barra.
        if (bar) {
          const left = Math.max(0, Math.min(toggle.offsetLeft, bar.clientWidth - panel.offsetWidth));
          panel.style.left = `${left}px`; panel.style.right = 'auto';
        }
      },
      close() {
        if (panel.hidden) return;
        panel.hidden = true; toggle.setAttribute('aria-expanded', 'false'); openXFs.delete(api);
        bar?.classList.remove('xf-open');
      },
      // Nuevas opciones/selección (tras cada consulta). Si el panel está abierto
      // se conserva lo que el usuario está marcando.
      update(cols) {
        columns = cols;
        if (!panel.hidden) {
          columns.forEach(col => { if (draft[col.key] === undefined) draft[col.key] = initDraft(col); });
          renderPanel();
        }
        api.renderChips();
      },
      renderChips() {
        const activos = columns.filter(c => c.selected && c.selected.length);
        const badge = host.querySelector('.xf-count');
        badge.hidden = !activos.length; badge.textContent = activos.length;
        toggle.classList.toggle('has-filters', activos.length > 0);
        if (!chipsHost) return;
        chipsHost.innerHTML = activos.map(col => {
          const sel = col.selected;
          const text = col.single || sel.length <= 2 ? sel.map(v => labelOf(col, v)).join(', ') : `${fmtN(sel.length)} de ${fmtN(col.options.length)}`;
          return `<span class="xf-chip">${esc(col.label)}: <b>${esc(text)}</b><button type="button" data-xf-quitar="${esc(col.key)}" aria-label="Quitar filtro ${esc(col.label)}">×</button></span>`;
        }).join('');
      },
      get isOpen() { return !panel.hidden; },
    };

    toggle.addEventListener('click', () => (panel.hidden ? api.open() : api.close()));
    host.addEventListener('click', e => {
      const b = e.target.closest('[data-xf]'); if (!b) return;
      if (b.dataset.xf === 'clear') {
        columns.forEach(col => { draft[col.key] = col.single ? null : new Set(col.options.map(o => o.value)); search[col.key] = ''; });
        renderPanel();
      }
      if (b.dataset.xf === 'apply' && !b.disabled) { const sel = selectionFromDraft(); api.close(); onApply(sel); }
    });
    chipsHost?.addEventListener('click', e => {
      const key = e.target.dataset?.xfQuitar; if (!key) return;
      onApply({ ...currentSelection(), [key]: null });
    });
    return api;
  }

  window.FiltrosGuia = { createXF, closeAllXF, paintIcons, tipoCambioOficial, ICONS, listOf, selOf, joinSel };
})();
