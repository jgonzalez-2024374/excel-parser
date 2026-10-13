/* ==========================================================================
   INGRESOS · Botón "Descargar PDF" (a la par de Filtros)
   --------------------------------------------------------------------------
   Al hacer click abre una ventana de SELECCIÓN MÚLTIPLE: marcas los módulos
   que quieres imprimir y solo esos viajan al documento. Genera el reporte
   con jsPDF + jspdf-autotable (cargados bajo demanda, igual que en
   RRHH / SAC / Flujo de ventas):

     · Portada con logo RED real (data-URI) y alcance del reporte
     · Chips de alcance (país · período · moneda · movimientos)
     · Franja con los módulos incluidos en el documento
     · 1. Indicadores del período (KPIs)          ← módulo "kpi"
     · 2. Análisis gráfico por banco              ← módulo "charts"
     · 3. Resumen por banco                       ← módulo "banks"
     · 4. Detalle de cuentas bancarias            ← módulo "accounts"
     · 5. Flujo diario (ingresos, egresos, neto)  ← módulo "daily"
     · 6. Movimientos de cheques                  ← módulo "cheques"
     · 7. Informe detallado de movimientos        ← módulo "informe"

   Respeta el país, el período y los filtros de banco activos en pantalla.
   Encabezado compacto y pie con número de página en todas las hojas.
   ========================================================================== */
(function () {
  'use strict';

  const BTN = document.getElementById('ingresosPdfBtn');
  if (!BTN) return;

  const LOGO_IMG = BTN.querySelector('img');

  const JSPDF_URL = 'https://cdn.jsdelivr.net/npm/jspdf@2.5.2/dist/jspdf.umd.min.js';
  const AUTOTABLE_URL = 'https://cdn.jsdelivr.net/npm/jspdf-autotable@3.8.2/dist/jspdf.plugin.autotable.min.js';

  // ─────────────────────────────────────── catálogo de módulos imprimibles
  // El orden aquí define el orden de las secciones en el PDF y su numeración.
  const MODULES = [
    {
      id: 'kpi',
      title: 'Indicadores del período',
      hint: 'KPIs con saldo inicial, saldo final, ingresos, egresos, variación neta y movimientos.',
      icon: '<path d="M3 20h18" /><rect x="5" y="10.5" width="3.6" height="6.5" rx="1" />' +
        '<rect x="10.2" y="6" width="3.6" height="11" rx="1" /><rect x="15.4" y="13" width="3.6" height="4" rx="1" />'
    },
    {
      id: 'charts',
      title: 'Análisis gráfico',
      hint: 'Ingresos vs. egresos y flujo neto (ingresos - egresos) por banco.',
      icon: '<path d="M3.5 20.5h17" /><path d="M6.5 20.5V12.5" /><path d="M11 20.5V6" />' +
        '<path d="M15.5 20.5v-6.5" /><path d="M20 20.5V10" />'
    },
    {
      id: 'banks',
      title: 'Resumen por banco',
      hint: 'Saldos, variación, ingresos/egresos, flujo neto y participación de cada banco.',
      icon: '<path d="M3 10 12 4l9 6" /><path d="M5.5 10v8M10 10v8M14 10v8M18.5 10v8" /><path d="M3 20h18" />'
    },
    {
      id: 'accounts',
      title: 'Detalle de cuentas',
      hint: 'Saldo inicial y final, movimientos y % de variación por cuenta bancaria.',
      icon: '<rect x="2.5" y="5.5" width="19" height="13" rx="2.5" /><path d="M2.5 10h19" /><path d="M6 15h4" />'
    },
    {
      id: 'daily',
      title: 'Flujo diario',
      hint: 'Ingresos, egresos y flujo neto día a día, con los totales del período.',
      icon: '<path d="M3 17.5 8.5 12l3.5 3.5L21 6.5" /><path d="M15 6.5h6v6" />'
    },
    {
      id: 'cheques',
      title: 'Movimientos de cheques',
      hint: 'Cheques emitidos y cobrados en el período, con sus totales.',
      icon: '<path d="M6 3h8l4 4v14H6z" /><path d="M14 3v4h4" /><path d="m9 14 2 2 4-4" />'
    },
    {
      id: 'informe',
      title: 'Informe detallado',
      hint: 'Todos los movimientos con país, banco, cuenta, referencia y descripción.',
      icon: '<rect x="3" y="4" width="18" height="16" rx="2" /><path d="M3 9.5h18" /><path d="M9.5 9.5V20" />' +
        '<path d="M13 13.5h4M13 16.5h4" />'
    }
  ];

  const MODULE_IDS = MODULES.map(module => module.id);

  // A4 horizontal (mm) · márgenes
  const PAGE = { w: 297, h: 210, ml: 12, mr: 12, mt: 24, mb: 18 };
  const X = PAGE.ml;
  const W = PAGE.w - PAGE.ml - PAGE.mr;   // 273 mm de ancho útil

  const C = {
    navy: [21, 28, 46],
    navy2: [28, 36, 56],
    red: [255, 23, 68],
    text: [23, 32, 51],
    muted: [107, 122, 148],
    line: [216, 222, 232],
    rowAlt: [247, 249, 252],
    chip: [241, 244, 249],
    white: [255, 255, 255],
    green: [8, 116, 67],
    rose: [180, 35, 60],
    cyan: [0, 132, 168],
    amber: [166, 104, 10],
    headFoot: [233, 238, 246],
    softText: [169, 185, 207]
  };

  // ────────────────────────────────────────────── utilidades de texto / cifras

  // Helvetica (fuente base del PDF) solo cubre Latin-1: se normaliza el texto.
  const txt = value => String(value == null ? '' : value)
    .replace(/[\u2018\u2019]/g, "'")
    .replace(/[\u201C\u201D]/g, '"')
    .replace(/[\u2013\u2014\u2212]/g, '-')
    .replace(/[\u2022\u00B7]/g, '·')
    .replace(/[\r\n\t]+/g, ' ')
    .replace(/[^\u0000-\u00FF]/g, '')
    .trim();

  const n = value => { const v = Number(value); return Number.isFinite(v) ? v : 0; };
  const num2 = value => n(value).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const num0 = value => Math.round(n(value)).toLocaleString('en-US');

  function money(value) {
    try { return moneyAllViews(n(value)); } catch (e) { /* income.js aún no listo */ }
    try { return `${currencySymbol()} ${num2(value)}`.trim(); } catch (e) { }
    return num2(value);
  }

  function fmtDateX(iso) {
    if (!iso) return '—';
    try { return fmtDate(iso); } catch (e) { }
    const [y, m, d] = String(iso).split('-');
    return (y && m && d) ? `${d}/${m}/${y}` : String(iso);
  }

  function pctX(value, decimals) {
    const v = n(value);
    const dec = decimals == null ? 2 : decimals;
    return `${v >= 0 ? '+' : ''}${v.toFixed(dec)}%`;
  }

  function activeCountry() {
    if (typeof window.DASHBOARD_ACTIVE_COUNTRY === 'string') return window.DASHBOARD_ACTIVE_COUNTRY;
    try { return ACTIVE_COUNTRY; } catch (e) { return 'CONSOLIDADO'; }
  }

  function countryLabelX(country) {
    try { return countryLabel(country); } catch (e) { }
    return ({
      GUATEMALA: 'Guatemala',
      EL_SALVADOR: 'El Salvador',
      CONSOLIDADO: 'Consolidado regional'
    })[country] || '—';
  }

  function currencyLabelX() {
    try { return currencyLabel(); } catch (e) { }
    try { return currencyCode(); } catch (e) { }
    return 'Moneda local';
  }

  function countryOf(transaction) {
    try { return recordCountry(transaction); } catch (e) { }
    const t = transaction || {};
    if (t.country) return t.country;
    return String(t.bankKey || t.key || '').indexOf('GUATEMALA|') === 0 ? 'GUATEMALA' : 'EL_SALVADOR';
  }

  function txAmount(transaction, value) {
    try { return informeAmount(transaction, value); } catch (e) { }
    try { return convertCurrencyForView(amountInActiveBase(transaction, value)); } catch (e) { }
    return n(value);
  }

  const valueOf = id => { const el = document.getElementById(id); return el ? el.value : ''; };

  // ───────────────────────────────────────────────────────── carga de librerías

  function loadScript(url) {
    return new Promise((resolve, reject) => {
      const found = document.querySelector(`script[src="${url}"]`);
      if (found) {
        if (found.dataset.ready === '1') return resolve();
        found.addEventListener('load', () => resolve());
        found.addEventListener('error', () => reject(new Error('Sin conexión para cargar la librería PDF.')));
        return;
      }
      const script = document.createElement('script');
      script.src = url;
      script.async = true;
      script.onload = () => { script.dataset.ready = '1'; resolve(); };
      script.onerror = () => reject(new Error('Sin conexión para cargar la librería PDF.'));
      document.head.appendChild(script);
    });
  }

  async function loadJsPdf() {
    await loadScript(JSPDF_URL);
    if (!window.jspdf || typeof window.jspdf.jsPDF !== 'function') throw new Error('jsPDF no disponible');
    await loadScript(AUTOTABLE_URL);
    return window.jspdf.jsPDF;
  }

  // ───────────────────────────────────────────────────── logo (PNG → dataURL)

  let logoCache;

  // Fuente preferida: el data-URI que inyecta income-pdf-logo.js. Va embebido
  // en el propio JS, así que no hay peticiones de red ni problemas de CORS y
  // el logo real aparece aunque la página se abra con file://.
  function embeddedLogo() {
    const uri = window.INTELFOON_LOGO_DATAURI;
    return (typeof uri === 'string' && uri.indexOf('data:image/png;base64,') === 0) ? uri : null;
  }

  // Respaldo: el <img> del botón de descarga, si la página vuelve a incluirlo.
  function logoFromButton() {
    return new Promise(resolve => {
      if (!LOGO_IMG) return resolve(null);
      Promise.resolve()
        .then(() => {
          if (LOGO_IMG.complete && LOGO_IMG.naturalWidth) return LOGO_IMG;
          return new Promise((res, rej) => {
            LOGO_IMG.addEventListener('load', () => res(LOGO_IMG), { once: true });
            LOGO_IMG.addEventListener('error', rej, { once: true });
            setTimeout(() => rej(new Error('timeout logo')), 5000);
          });
        })
        .then(img => {
          const size = 512;
          const canvas = document.createElement('canvas');
          canvas.width = canvas.height = size;
          canvas.getContext('2d').drawImage(img, 0, 0, size, size);
          resolve(canvas.toDataURL('image/png'));   // puede fallar en file://
        })
        .catch(error => {
          console.warn('[PDF] Logo del botón no embebible:', error.message);
          resolve(null);
        });
    });
  }

  async function logoDataUrl() {
    if (logoCache !== undefined) return logoCache;
    logoCache = embeddedLogo() || await logoFromButton();
    return logoCache;
  }

  // Respaldo si el PNG no se puede convertir (origen restringido).
  function drawVectorLogo(doc, cx, cy, radius) {
    doc.setFillColor(168, 172, 178);
    doc.circle(cx, cy, radius, 'F');
    doc.setFillColor(232, 234, 237);
    doc.circle(cx, cy, radius * 0.92, 'F');
    doc.setFillColor(214, 216, 220);
    doc.circle(cx, cy, radius * 0.8, 'F');
    doc.setFillColor(226, 24, 44);
    doc.circle(cx, cy, radius * 0.74, 'F');
    doc.setFillColor(196, 12, 38);
    doc.circle(cx, cy + radius * 0.35, radius * 0.4, 'F');
    doc.setTextColor(255, 255, 255);
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(Math.max(6, radius * 1.6));
    doc.text('RED', cx, cy + radius * 0.16, { align: 'center' });
  }

  function drawLogo(doc, dataUrl, x, y, size) {
    if (dataUrl) {
      // El PNG mide 256×247: se respeta la proporción y se centra en la caja
      let w = size;
      let h = size;
      try {
        const props = doc.getImageProperties(dataUrl);
        const ratio = props && props.height ? props.width / props.height : 1;
        if (ratio >= 1) h = size / ratio; else w = size * ratio;
      } catch (e) { /* sin proporción: caja cuadrada */ }
      try {
        doc.addImage(dataUrl, 'PNG', x + (size - w) / 2, y + (size - h) / 2, w, h, undefined, 'FAST');
        return;
      } catch (e) { /* cae al vectorial */ }
    }
    drawVectorLogo(doc, x + size / 2, y + size / 2, size / 2);
  }

  // ──────────────────────────────────────────────── modelo de datos a reportar

  function visibleAccounts(all) {
    const select = document.getElementById('bankFilter');
    if (!select) return all;
    const selected = Array.from(select.selectedOptions).map(option => option.value);
    const showAll = !selected.length || selected.indexOf('ALL') !== -1;
    return showAll ? all : all.filter(account => selected.indexOf(account.bankKey) !== -1);
  }

  function chequeRows(from, to) {
    try {
      const regional = activeCountry() === 'CONSOLIDADO';
      const raw = (typeof RAW_TRANSACTIONS !== 'undefined' && RAW_TRANSACTIONS) || [];
      const select = document.getElementById('chequeBankFilter');
      const selected = select ? Array.from(select.selectedOptions).map(option => option.value) : ['ALL'];
      const showAll = !selected.length || selected.indexOf('ALL') !== -1;
      const bankSelection = (typeof SELECTED_BANK_KEYS !== 'undefined' && SELECTED_BANK_KEYS) || null;
      const inSelection = transaction => !bankSelection || !bankSelection.size || bankSelection.has(transaction.bankKey);

      return raw
        .filter(transaction => regional || countryOf(transaction) === activeCountry())
        .filter(inSelection)
        .filter(isChequeTransaction)
        .filter(transaction => showAll || selected.indexOf(transaction.bankKey) !== -1)
        .filter(transaction => transaction.date >= from && transaction.date <= to)
        .sort((a, b) => String(b.date || '').localeCompare(String(a.date || '')));
    } catch (error) {
      console.warn('[PDF] No se pudieron leer los cheques:', error.message);
      return [];
    }
  }

  function informeRowsX() {
    try {
      if (typeof informeRows === 'function') return informeRows();
    } catch (e) { }
    const data = window.DATA || {};
    return (data.transactions || []).slice().sort((a, b) => String(b.date || '').localeCompare(String(a.date || '')));
  }

  function buildModel() {
    const data = window.DATA || {};
    const meta = data.meta || {};
    const totals = data.totals || {};
    const from = meta.filterFrom || valueOf('dateFrom') || '';
    const to = meta.filterTo || valueOf('dateTo') || '';
    const pais = activeCountry();
    const informe = informeRowsX();
    const cheques = chequeRows(from, to);

    let totalDebit = 0;
    let totalCredit = 0;
    informe.forEach(transaction => {
      totalDebit += txAmount(transaction, transaction.debit || 0);
      totalCredit += txAmount(transaction, transaction.credit || 0);
    });

    let chequeDebit = 0;
    let chequeCredit = 0;
    cheques.forEach(transaction => {
      chequeDebit += txAmount(transaction, transaction.debit || 0);
      chequeCredit += txAmount(transaction, transaction.credit || 0);
    });

    return {
      pais,
      from,
      to,
      meta,
      totals,
      countryLabel: countryLabelX(pais),
      currencyLabel: currencyLabelX(),
      banks: (data.banks || []).slice().sort((a, b) => n(b.final) - n(a.final)),
      accounts: visibleAccounts(data.accounts || []),
      daily: (data.daily || []).slice(),
      cheques,
      informe,
      totalDebit,
      totalCredit,
      chequeDebit,
      chequeCredit,
      movementCount: n(meta.movementCount) || informe.length,
      bankCount: n(meta.banks) || (data.banks || []).length,
      accountCount: n(meta.accounts) || (data.accounts || []).length
    };
  }

  function scopeText(model) {
    return `${model.countryLabel} · Período ${fmtDateX(model.from)} al ${fmtDateX(model.to)} · ` +
      `Moneda ${model.currencyLabel}`;
  }

  function fileName(model) {
    const pais = model.pais === 'CONSOLIDADO' ? 'GT-SV' : model.pais === 'GUATEMALA' ? 'GT' : 'SV';
    return `Reporte_ingresos_${pais}_${model.from || 'inicio'}_${model.to || 'fin'}.pdf`;
  }

  // ──────────────────────────────────────────────────────────── piezas gráficas

  function fitFont(doc, text, maxWidth, size, min) {
    let current = size;
    doc.setFontSize(current);
    const value = String(text);
    while (current > (min || 6) && doc.getTextWidth(value) > maxWidth) {
      current -= 0.5;
      doc.setFontSize(current);
    }
    return current;
  }

  function drawCover(doc, logo, model, generado) {
    const h = 26;

    doc.setFillColor(C.navy[0], C.navy[1], C.navy[2]);
    doc.roundedRect(X, 12, W, h, 3, 3, 'F');
    doc.setFillColor(C.red[0], C.red[1], C.red[2]);
    doc.roundedRect(X + 4, 16, 2.4, h - 8, 1.2, 1.2, 'F');

    // Logo real (PNG con fondo transparente) en su pastilla original,
    // separado de la barrita roja del lateral
    drawLogo(doc, logo, X + 10.5, 14.5, 21);

    doc.setTextColor(C.white[0], C.white[1], C.white[2]);
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(8);
    doc.text('DASHBOARD DE INGRESOS · RED INTELFOON', X + 37, 21, { charSpace: 0.7 });
    doc.setFontSize(15.5);
    doc.text('Reporte de saldos bancarios', X + 37, 31);

    doc.setFont('helvetica', 'normal');
    doc.setFontSize(8);
    doc.setTextColor(C.softText[0], C.softText[1], C.softText[2]);
    doc.text(txt(`Generado el ${generado}`), X + W - 6, 21, { align: 'right' });

    doc.setFont('helvetica', 'bold');
    const alcance = txt(scopeText(model));
    fitFont(doc, alcance, W / 2 - 10, 9, 6.5);
    doc.setTextColor(C.white[0], C.white[1], C.white[2]);
    doc.text(alcance, X + W - 6, 31, { align: 'right' });
  }

  function drawChips(doc, model) {
    const y = 43;
    const gap = 4;
    const width = (W - gap * 3) / 4;
    const chips = [
      { label: 'PAÍS / OPERACIÓN', value: model.countryLabel },
      { label: 'PERÍODO', value: `${fmtDateX(model.from)} – ${fmtDateX(model.to)}` },
      { label: 'MONEDA', value: model.currencyLabel },
      { label: 'MOVIMIENTOS', value: `${num0(model.movementCount)} registros` }
    ];

    chips.forEach((chip, index) => {
      const x = X + index * (width + gap);
      doc.setFillColor(C.chip[0], C.chip[1], C.chip[2]);
      doc.roundedRect(x, y, width, 12.5, 2, 2, 'F');
      doc.setFillColor(C.red[0], C.red[1], C.red[2]);
      doc.roundedRect(x, y, 1.6, 12.5, 0.8, 0.8, 'F');

      doc.setFont('helvetica', 'bold');
      doc.setFontSize(6.2);
      doc.setTextColor(C.red[0], C.red[1], C.red[2]);
      doc.text(chip.label, x + 5, y + 5, { charSpace: 0.4 });

      doc.setFontSize(9);
      doc.setTextColor(C.text[0], C.text[1], C.text[2]);
      fitFont(doc, chip.value, width - 8, 9, 6);
      doc.text(txt(chip.value), x + 5, y + 10);
      doc.setFontSize(9);
    });

    return y + 12.5;
  }

  // Franja bajo los chips: lista los módulos que viajan al documento.
  function drawScopeStrip(doc, mods) {
    const y = 58.5;
    const h = 9;

    doc.setFillColor(C.chip[0], C.chip[1], C.chip[2]);
    doc.roundedRect(X, y, W, h, 2, 2, 'F');
    doc.setFillColor(C.red[0], C.red[1], C.red[2]);
    doc.roundedRect(X, y, 1.6, h, 0.8, 0.8, 'F');

    const label = 'MÓDULOS INCLUIDOS';
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(6.2);
    doc.setTextColor(C.red[0], C.red[1], C.red[2]);
    // Se acota a una caja fija: getTextWidth ignora el charSpace del rótulo
    // y la lista se acababa pegando al texto.
    fitFont(doc, label, 30, 6.2, 5);
    doc.text(label, X + 5, y + 5.9, { charSpace: 0.4 });
    const labelWidth = 40;

    const names = mods
      .map(id => (MODULES.find(module => module.id === id) || {}).title)
      .filter(Boolean)
      .map((title, index) => `${index + 1}. ${title}`)
      .join('  ·  ');

    doc.setFont('helvetica', 'normal');
    doc.setFontSize(7.6);
    doc.setTextColor(C.text[0], C.text[1], C.text[2]);
    fitFont(doc, names, W - labelWidth - 6, 7.6, 5.4);
    doc.text(txt(names), X + labelWidth, y + 5.9);

    return y + h;
  }

  // Título de sección: barrita roja + "N. Título" + nota a la derecha.
  // Si el bloque no cabe en la página actual, se salta a una nueva.
  function startSection(doc, y, index, title, note, available) {
    // Si no cabe el título + una fila, se pasa de página
    if (y + (available || 22) > PAGE.h - PAGE.mb) {
      doc.addPage();
      y = PAGE.mt;
    }
    doc.setFillColor(C.red[0], C.red[1], C.red[2]);
    doc.roundedRect(X, y, 2.2, 6, 1, 1, 'F');

    doc.setFont('helvetica', 'bold');
    doc.setFontSize(11);
    doc.setTextColor(C.text[0], C.text[1], C.text[2]);
    const heading = `${index}. ${txt(title)}`;
    doc.text(heading, X + 5, y + 4.8);

    if (note) {
      const headingWidth = doc.getTextWidth(heading);
      doc.setFont('helvetica', 'normal');
      doc.setFontSize(7.5);
      doc.setTextColor(C.muted[0], C.muted[1], C.muted[2]);
      const noteX = X + 5 + headingWidth + 5;
      if (noteX < X + W - 34) doc.text(txt(note), noteX, y + 4.8);
    }
    return y + 9;
  }

  function drawKpis(doc, model, y) {
    const totals = model.totals;
    const net = n(totals.netFlow);
    const items = [
      { label: 'SALDO INICIAL TOTAL', value: money(totals.initial), color: C.navy },
      { label: 'SALDO FINAL TOTAL', value: money(totals.final), color: C.navy },
      { label: 'INGRESOS ACUMULADOS', value: money(totals.credits), color: C.green },
      { label: 'EGRESOS ACUMULADOS', value: money(totals.debits), color: C.rose },
      {
        label: 'VARIACIÓN NETA',
        value: money(net),
        extra: n(totals.debits) ? pctX(net / n(totals.debits) * 100) : '—',
        color: net >= 0 ? C.green : C.rose
      },
      { label: 'MOVIMIENTOS', value: num0(model.movementCount), color: C.cyan }
    ];

    const gap = 4;
    const width = (W - gap * 5) / 6;
    const height = 23;

    items.forEach((item, index) => {
      const x = X + index * (width + gap);

      doc.setFillColor(C.white[0], C.white[1], C.white[2]);
      doc.setDrawColor(C.line[0], C.line[1], C.line[2]);
      doc.setLineWidth(0.2);
      doc.roundedRect(x, y, width, height, 2, 2, 'FD');

      doc.setFillColor(item.color[0], item.color[1], item.color[2]);
      doc.roundedRect(x, y, width, 1.8, 1, 1, 'F');

      doc.setFont('helvetica', 'bold');
      doc.setFontSize(6.2);
      doc.setTextColor(C.muted[0], C.muted[1], C.muted[2]);
      doc.text(item.label, x + 3.5, y + 6.6, { charSpace: 0.3 });

      doc.setFont('helvetica', 'bold');
      doc.setTextColor(item.color[0], item.color[1], item.color[2]);
      fitFont(doc, item.value, width - 7, 12, 7);
      doc.text(item.value, x + 3.5, y + 14);

      doc.setFont('helvetica', 'normal');
      const foot = index < 2 ? 'Al cierre del período seleccionado'
        : index === 2 ? 'Abonos / entradas del período'
          : index === 3 ? 'Cargos / salidas del período'
            : index === 4 ? `Neto vs. egresos: ${item.extra}`
              : 'Registros en el período';
      fitFont(doc, foot, width - 7, 6.4, 5.2);
      doc.setTextColor(C.muted[0], C.muted[1], C.muted[2]);
      doc.text(foot, x + 3.5, y + 19.5);
    });

    return y + height;
  }

  // ───────────────────────────────────────────────────── gráficas por banco

  // Encabezado de cada gráfica: barrita roja + rótulo + subtítulo.
  // Sin tarjeta ni relleno: la gráfica se apoya directo en la página.
  function chartHeading(doc, x, y, w, title, subtitle) {
    doc.setFillColor(C.red[0], C.red[1], C.red[2]);
    doc.roundedRect(x, y + 0.9, 1.8, 4.6, 0.9, 0.9, 'F');

    doc.setFont('helvetica', 'bold');
    doc.setTextColor(C.text[0], C.text[1], C.text[2]);
    fitFont(doc, title, w - 50, 8.6, 6.4);   // deja aire para la leyenda
    doc.text(txt(title), x + 4.4, y + 4.7);

    doc.setFont('helvetica', 'normal');
    doc.setTextColor(C.muted[0], C.muted[1], C.muted[2]);
    fitFont(doc, subtitle, w - 7, 6.6, 5.4);
    doc.text(txt(subtitle), x + 4.4, y + 9.6);
  }

  // Serie compartida: principales bancos por saldo final (top 6 + "Otros").
  function chartSeries(model) {
    const banks = (model.banks || [])
      .filter(bank => n(bank.final) > 0 || n(bank.credits) > 0 || n(bank.debits) > 0)
      .slice()
      .sort((a, b) => n(b.final) - n(a.final));

    const MAX = 7;
    const head = banks.length > MAX ? banks.slice(0, MAX - 1) : banks.slice();
    const tail = banks.slice(head.length);

    const rows = head.map(bank => ({
      name: bank.name,
      final: n(bank.final),
      credits: n(bank.credits),
      debits: n(bank.debits)
    }));

    if (tail.length) {
      rows.push({
        name: `Otros (${tail.length})`,
        final: tail.reduce((sum, bank) => sum + n(bank.final), 0),
        credits: tail.reduce((sum, bank) => sum + n(bank.credits), 0),
        debits: tail.reduce((sum, bank) => sum + n(bank.debits), 0)
      });
    }

    return { rows };
  }

  // 1ª gráfica · barras pareadas de ingresos y egresos por banco.
  function drawFlowPanel(doc, x, y, w, h, model) {
    chartHeading(doc, x, y, w, 'Ingresos vs. egresos por banco',
      'Entradas y salidas de fondos del período');

    const series = chartSeries(model);

    if (!series.rows.length) {
      doc.setFont('helvetica', 'normal');
      doc.setFontSize(8);
      doc.setTextColor(C.muted[0], C.muted[1], C.muted[2]);
      doc.text('Sin datos para la gráfica en la vista activa.', x + 4.4, y + 30);
      return;
    }

    // Leyenda de las series, alineada con el rótulo
    const legendY = y + 2.6;
    const gDot = x + w - 46;
    const rDot = x + w - 23;
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(6.6);
    doc.setTextColor(C.muted[0], C.muted[1], C.muted[2]);
    doc.setFillColor(C.green[0], C.green[1], C.green[2]);
    doc.roundedRect(gDot, legendY, 2.4, 2.4, 0.6, 0.6, 'F');
    doc.text('Ingresos', gDot + 3.6, legendY + 2);
    doc.setFillColor(C.red[0], C.red[1], C.red[2]);
    doc.roundedRect(rDot, legendY, 2.4, 2.4, 0.6, 0.6, 'F');
    doc.text('Egresos', rDot + 3.6, legendY + 2);

    const labelX = x + 4.4;
    const labelW = 34;
    const barX = x + 42;
    const barMax = w - 42 - 6 - 20;          // deja 20 mm para el importe
    const top = y + 16;
    const rowH = (h - 20) / series.rows.length;

    const max = series.rows.reduce(
      (peak, row) => Math.max(peak, Math.abs(row.credits), Math.abs(row.debits)), 0) || 1;

    series.rows.forEach((row, index) => {
      const mid = top + index * rowH + rowH / 2;
      const barH = Math.min(2.2, (rowH - 1.8) / 2);
      const gap = Math.max(0.7, rowH - 2 * barH - 1.4);

      doc.setFont('helvetica', 'bold');
      doc.setFontSize(7);
      doc.setTextColor(C.text[0], C.text[1], C.text[2]);
      fitFont(doc, row.name, labelW, 7, 5.2);
      doc.text(txt(row.name), labelX, mid + 1.3);

      const lines = [
        { value: row.credits, color: C.green, y: mid - gap / 2 - barH },
        { value: row.debits, color: C.red, y: mid + gap / 2 }
      ];

      lines.forEach(line => {
        const length = Math.max(0.6, barMax * (Math.abs(line.value) / max));
        doc.setFillColor(line.color[0], line.color[1], line.color[2]);
        doc.roundedRect(barX, line.y, length, barH, 0.5, 0.5, 'F');

        doc.setFont('helvetica', 'bold');
        doc.setTextColor(line.color[0], line.color[1], line.color[2]);
        const amount = money(line.value);
        fitFont(doc, amount, 19, 6.4, 5);
        doc.text(amount, barX + length + 1.8, line.y + barH - 0.55);
      });
    });
  }

  // 2ª gráfica · lollipop divergente del FLUJO NETO por banco.
  // Tarjeta en fondo oscuro, eje central en cero y tallos con bola en el
  // extremo: verde para los bancos que suman, rojo para los que restan.
  // Paleta del lollipop sobre fondo de página (sin tarjeta).
  const NET_POS = [8, 116, 67];      // verde de la casa (igual que las barras)
  const NET_NEG = [255, 23, 68];     // rojo de la casa
  const NET_AXIS = [148, 163, 184];  // eje neutro
  const NET_DROP = [216, 222, 232];  // guías de caída
  const NET_NAME = [23, 32, 51];     // rótulos de banco

  function netSeries(model) {
    const rows = (model.banks || [])
      .filter(bank => bank && bank.name)
      .map(bank => ({ name: bank.name, net: n(bank.netFlow) }))
      .sort((a, b) => b.net - a.net);   // de mayor a menor flujo

    const MAX = 9;
    if (rows.length <= MAX) return rows;

    const head = rows.slice(0, MAX - 1);
    const tail = rows.slice(MAX - 1);
    head.push({ name: `Otros (${tail.length})`, net: tail.reduce((sum, row) => sum + row.net, 0) });
    return head;
  }

  // Parte el rótulo en, como máximo, dos líneas que quepan en la columna.
  function twoLines(doc, value, maxWidth) {
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(5.8);
    const line = txt(value);
    if (doc.getTextWidth(line) <= maxWidth) return [line];

    const words = line.split(' ');
    if (words.length < 2) return [line];

    let cut = 1;
    let best = Infinity;
    for (let i = 1; i < words.length; i += 1) {
      const diff = Math.abs(doc.getTextWidth(words.slice(0, i).join(' ')) - maxWidth / 2);
      if (diff < best) { best = diff; cut = i; }
    }
    return [words.slice(0, cut).join(' '), words.slice(cut).join(' ')];
  }

  function drawNetPanel(doc, x, y, w, h, model) {
    const rows = netSeries(model);

    // Sin tarjeta ni fondo: solo rótulo, leyenda y gráfica sobre la página
    chartHeading(doc, x, y, w, 'Flujo neto por banco', '(= Ingresos - Egresos)');

    const legendY = y + 2.6;
    const pDot = x + w - 46;
    const nDot = x + w - 23;
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(6.6);
    doc.setTextColor(C.muted[0], C.muted[1], C.muted[2]);
    doc.setFillColor(NET_POS[0], NET_POS[1], NET_POS[2]);
    doc.roundedRect(pDot, legendY, 2.4, 2.4, 0.6, 0.6, 'F');
    doc.text('Positivo', pDot + 3.6, legendY + 2);
    doc.setFillColor(NET_NEG[0], NET_NEG[1], NET_NEG[2]);
    doc.roundedRect(nDot, legendY, 2.4, 2.4, 0.6, 0.6, 'F');
    doc.text('Negativo', nDot + 3.6, legendY + 2);

    if (!rows.length) {
      doc.setFontSize(8);
      doc.text('Sin datos para la gráfica en la vista activa.', x + 4.4, y + 30);
      return;
    }

    const plotL = x + 7;
    const plotR = x + w - 7;
    const top = y + 17;
    const bottom = y + 53;
    const namesTop = y + 55.5;

    const posMax = rows.reduce((peak, row) => Math.max(peak, row.net), 0);
    const negMax = rows.reduce((peak, row) => Math.max(peak, -row.net), 0);
    const scale = (bottom - top - 11) / ((posMax + negMax) || 1);
    const axisY = top + 5.5 + posMax * scale;
    const colW = (plotR - plotL) / rows.length;

    // Guías de caída hasta los rótulos de banco
    doc.setFillColor(NET_DROP[0], NET_DROP[1], NET_DROP[2]);
    rows.forEach((row, index) => {
      const cx = plotL + colW * (index + 0.5);
      doc.rect(cx - 0.1, axisY, 0.2, Math.max(0, namesTop - 1 - axisY), 'F');
    });

    // Eje central (cero)
    doc.setFillColor(NET_AXIS[0], NET_AXIS[1], NET_AXIS[2]);
    doc.rect(plotL, axisY - 0.15, plotR - plotL, 0.3, 'F');

    rows.forEach((row, index) => {
      const cx = plotL + colW * (index + 0.5);
      const positive = row.net >= 0;
      const color = positive ? NET_POS : NET_NEG;
      const stem = Math.abs(row.net) * scale;
      const cy = positive ? axisY - stem : axisY + stem;

      // Tallo + bola
      doc.setFillColor(color[0], color[1], color[2]);
      if (stem > 0.2) doc.rect(cx - 0.35, positive ? cy : axisY, 0.7, stem, 'F');
      doc.circle(cx, cy, 1.7, 'F');

      // Importe fuera del eje
      doc.setFont('helvetica', 'bold');
      doc.setTextColor(color[0], color[1], color[2]);
      const amount = money(row.net);
      fitFont(doc, amount, colW - 1, 6, 4.6);
      doc.text(amount, cx, positive ? cy - 3 : cy + 4.4, { align: 'center' });

      // Nombre del banco al pie de la columna (mismo cuerpo en todas las líneas)
      const lines = twoLines(doc, row.name, colW - 0.8);
      doc.setFont('helvetica', 'bold');
      doc.setFontSize(5.8);
      let size = 5.8;
      while (size > 4.2 && lines.some(line => doc.getTextWidth(line) > colW - 0.8)) {
        size -= 0.2;
        doc.setFontSize(size);
      }
      // Los rótulos de una y dos líneas comparten la primera línea de base,
      // así todos arrancan a la misma altura bajo su guía de caída.
      const base = [namesTop + 3, namesTop + 6.6].slice(0, lines.length);
      doc.setTextColor(NET_NAME[0], NET_NAME[1], NET_NAME[2]);
      lines.forEach((line, i) => doc.text(line, cx, base[i], { align: 'center' }));
    });
  }

  // Los dos paneles en paralelo: resumen ejecutivo bajo los KPI.
  function drawCharts(doc, model, y) {
    const gap = 6;
    const widthLeft = 140;
    const widthRight = W - gap - widthLeft;
    const height = 70;

    drawFlowPanel(doc, X, y, widthLeft, height, model);
    drawNetPanel(doc, X + widthLeft + gap, y, widthRight, height, model);

    return y + height;
  }

  // Tabla genérica con la estética del reporte
  function drawTable(doc, options) {
    const numeric = options.numeric || [];
    const colorByColumn = options.colors || {};
    const signColumns = options.signColors || [];   // verde si ≥ 0, rojo si < 0

    doc.autoTable({
      head: [options.head],
      body: options.body,
      foot: options.foot ? [options.foot] : undefined,
      startY: options.startY,
      margin: { left: X, right: X, top: PAGE.mt, bottom: PAGE.mb },
      theme: 'grid',
      styles: {
        font: 'helvetica',
        fontSize: 7.4,
        cellPadding: { top: 1.5, right: 2, bottom: 1.5, left: 2 },
        textColor: C.text,
        lineColor: C.line,
        lineWidth: 0.1,
        valign: 'middle',
        overflow: 'linebreak'
      },
      headStyles: {
        fillColor: C.navy2,
        textColor: C.white,
        fontStyle: 'bold',
        fontSize: 7.2,
        lineWidth: 0
      },
      footStyles: {
        fillColor: C.headFoot,
        textColor: C.text,
        fontStyle: 'bold',
        fontSize: 7.8,
        lineWidth: 0.1
      },
      alternateRowStyles: { fillColor: C.rowAlt },
      columnStyles: options.columnStyles || {},
      didParseCell: data => {
        const style = (options.columnStyles || {})[data.column.index] || {};
        const align = style.halign || (numeric.indexOf(data.column.index) !== -1 ? 'right' : 'left');

        // Encabezado y pie con la MISMA alineación que las celdas de la columna
        if (data.section === 'head' || data.section === 'foot') {
          data.cell.styles.halign = align;
          return;
        }

        data.cell.styles.halign = align;
        const raw = String(data.cell.raw == null ? '' : data.cell.raw).trim();
        if (!raw || raw === '—' || raw === '-') return;

        // Columnas de variación: color según el signo
        if (signColumns.indexOf(data.column.index) !== -1) {
          data.cell.styles.textColor = raw.indexOf('-') !== -1 ? C.rose : C.green;
          return;
        }

        const color = colorByColumn[data.column.index];
        if (color) data.cell.styles.textColor = color;
      }
    });

    return doc.lastAutoTable.finalY;
  }

  // Encabezado compacto para las páginas 2 en adelante
  function drawMiniHeader(doc, model, logo) {
    drawLogo(doc, logo, X, 11.4, 6.6);
    doc.setFillColor(C.red[0], C.red[1], C.red[2]);
    doc.roundedRect(X + 9, 10.5, 2.2, 6, 1, 1, 'F');
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(9);
    doc.setTextColor(C.text[0], C.text[1], C.text[2]);
    doc.text('REPORTE DE INGRESOS · SALDOS BANCARIOS', X + 13.4, 15.2, { charSpace: 0.5 });
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(7.4);
    doc.setTextColor(C.muted[0], C.muted[1], C.muted[2]);
    doc.text(txt(scopeText(model)), X + W, 15.2, { align: 'right' });
  }

  function drawFooters(doc, model, generado, logo) {
    const total = doc.internal.getNumberOfPages();
    for (let page = 1; page <= total; page++) {
      doc.setPage(page);

      doc.setDrawColor(C.line[0], C.line[1], C.line[2]);
      doc.setLineWidth(0.2);
      doc.line(X, PAGE.h - 13.5, X + W, PAGE.h - 13.5);
      // Acento de marca sobre el filete del pie
      doc.setFillColor(C.red[0], C.red[1], C.red[2]);
      doc.roundedRect(X, PAGE.h - 14.2, 18, 1.4, 0.7, 0.7, 'F');

      doc.setFont('helvetica', 'normal');
      doc.setFontSize(7);
      doc.setTextColor(C.muted[0], C.muted[1], C.muted[2]);
      doc.text('RED Intelfon · Dashboard de Ingresos', X + 21, PAGE.h - 9.5);
      doc.text(txt(`Generado ${generado}`), X + W / 2, PAGE.h - 9.5, { align: 'center' });
      doc.setFont('helvetica', 'bold');
      doc.setTextColor(C.text[0], C.text[1], C.text[2]);
      doc.text(`Página ${page} de ${total}`, X + W, PAGE.h - 9.5, { align: 'right' });

      if (page > 1) drawMiniHeader(doc, model, logo);
    }
  }

  // ───────────────────────────────────────────────────────────── generación

  async function generatePdf(selected) {
    // Módulos marcados en la ventana de selección múltiple (por defecto: todos)
    const mods = Array.isArray(selected) && selected.length
      ? MODULE_IDS.filter(id => selected.indexOf(id) !== -1)
      : MODULE_IDS.slice();
    const on = id => mods.indexOf(id) !== -1;

    const model = buildModel();

    if (!window.DATA || (!model.banks.length && !model.informe.length)) {
      toast('error', 'Sin datos para exportar', 'Espera a que termine de cargar la información.');
      return;
    }

    const jsPDF = await loadJsPdf();
    const logo = await logoDataUrl();
    const ahora = new Date();
    const generado = `${ahora.toLocaleDateString('es-GT')} · ${ahora.toLocaleTimeString('es-GT', { hour: '2-digit', minute: '2-digit' })}`;

    const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4', compress: true });
    doc.setProperties({
      title: `Reporte de ingresos · ${model.countryLabel} · ${fmtDateX(model.from)} al ${fmtDateX(model.to)}`,
      subject: 'Saldos bancarios, movimientos y flujo financiero',
      author: 'RED Intelfon · Dashboard de Ingresos',
      creator: 'Dashboard de Ingresos · Intelfon'
    });

    // ── Portada: encabezado, chips, alcance y módulos incluidos
    drawCover(doc, logo, model, generado);
    drawChips(doc, model);
    drawScopeStrip(doc, mods);

    let y = 71;
    let sec = 0;

    if (on('kpi')) {
      y = startSection(doc, y, ++sec, 'Indicadores del período', 'Cifras en la moneda de visualización activa', 34);
      y = drawKpis(doc, model, y);
    }

    // ── Análisis gráfico (barras ingresos/egresos + lollipop flujo neto)
    if (on('charts')) {
      y = startSection(doc, y + 7, ++sec, 'Análisis gráfico por banco',
        'Ingresos, egresos y flujo neto por entidad', 79);
      y = drawCharts(doc, model, y);
    }

    // ── Resumen por banco
    if (on('banks')) {
      y = startSection(doc, y + 7, ++sec, 'Resumen por banco', `${model.banks.length} banco(s) en la vista activa`, 26);
      y = drawTable(doc, {
        startY: y,
        head: ['#', 'Banco', 'Cuentas', 'Saldo inicial', 'Saldo final', 'Variación', '%', 'Ingresos', 'Egresos', 'Flujo neto', 'Part.'],
        numeric: [3, 4, 5, 6, 7, 8, 9, 10],
        signColors: [5, 6, 9],
        colors: { 7: C.green, 8: C.rose },
        columnStyles: {
          0: { cellWidth: 8, halign: 'center' },
          1: { cellWidth: 52, fontStyle: 'bold' },
          2: { cellWidth: 14, halign: 'center' },
          3: { cellWidth: 30 }, 4: { cellWidth: 30 }, 5: { cellWidth: 27 },
          6: { cellWidth: 18 }, 7: { cellWidth: 27 }, 8: { cellWidth: 27 },
          9: { cellWidth: 27 }, 10: { cellWidth: 13, halign: 'right' }
        },
        body: model.banks.map((bank, index) => [
          index + 1,
          txt(bank.name),
          num0(bank.accounts),
          money(bank.initial),
          money(bank.final),
          money(bank.change),
          bank.changePct === null || bank.changePct === undefined ? '—' : pctX(bank.changePct),
          money(bank.credits),
          money(bank.debits),
          money(bank.netFlow),
          `${n(bank.share).toFixed(1)}%`
        ])
      });
    }

    // ── Detalle de cuentas
    if (on('accounts')) {
      y = startSection(doc, y + 7, ++sec, 'Detalle de cuentas bancarias',
        `${model.accounts.length} cuenta(s) · saldos inicial y final recalculados al período`, 26);
      y = drawTable(doc, {
        startY: y,
        head: ['Banco', 'Cuenta', 'Moneda', 'Saldo inicial', 'Saldo final', 'Variación', '%', 'Ingresos', 'Egresos', 'Movim.'],
        numeric: [3, 4, 5, 6, 7, 8],
        signColors: [5, 6],
        colors: { 7: C.green, 8: C.rose },
        columnStyles: {
          0: { cellWidth: 49, fontStyle: 'bold' },
          1: { cellWidth: 26 },
          2: { cellWidth: 16, halign: 'center' },
          3: { cellWidth: 29 }, 4: { cellWidth: 29 }, 5: { cellWidth: 27 },
          6: { cellWidth: 18 }, 7: { cellWidth: 27 }, 8: { cellWidth: 27 },
          9: { cellWidth: 25, halign: 'center' }
        },
        body: model.accounts.map(account => [
          txt(account.bank),
          txt(account.account),
          txt(account.currencyOriginal || account.currency || '—'),
          money(account.initial),
          money(account.final),
          money(account.change),
          account.changePct === null || account.changePct === undefined ? '—' : pctX(account.changePct),
          money(account.credits),
          money(account.debits),
          num0(account.movements)
        ])
      });
    }

    // ── Flujo diario
    if (on('daily')) {
      y = startSection(doc, y + 7, ++sec, 'Flujo diario', 'Ingresos, egresos y flujo neto por fecha', 26);
      const dailyBody = model.daily.map(day => [
        fmtDateX(day.date),
        num0(day.moves),
        money(day.debits),
        money(day.credits),
        money(day.net)
      ]);
      y = drawTable(doc, {
        startY: y,
        head: ['Fecha', 'Movimientos', 'Egresos', 'Ingresos', 'Flujo neto'],
        numeric: [2, 3, 4],
        signColors: [4],
        colors: { 2: C.rose, 3: C.green },
        columnStyles: {
          0: { cellWidth: 60, fontStyle: 'bold' },
          1: { cellWidth: 45, halign: 'center' },
          2: { cellWidth: 56 }, 3: { cellWidth: 56 }, 4: { cellWidth: 56, fontStyle: 'bold' }
        },
        body: dailyBody.length ? dailyBody : [['—', '—', '—', '—', '—']],
        foot: dailyBody.length ? [
          'TOTALES', num0(model.daily.reduce((sum, day) => sum + n(day.moves), 0)),
          money(model.daily.reduce((sum, day) => sum + n(day.debits), 0)),
          money(model.daily.reduce((sum, day) => sum + n(day.credits), 0)),
          money(model.daily.reduce((sum, day) => sum + n(day.net), 0))
        ] : null
      });
    }

    // ── Movimientos de cheques
    if (on('cheques')) {
      y = startSection(doc, y + 7, ++sec, 'Movimientos de cheques',
        `${model.cheques.length} movimiento(s) con cheque en el período`, 26);
      const chequeBody = model.cheques.map(transaction => {
        const bank = (transaction.bank || transaction.bankName || transaction.bankCanonical || '—');
        // Algunos cheques vienen del Excel como "826476.0": se normaliza a entero
        const check = String(transaction.checkNumber || transaction.chequeNumber || transaction.reference || '—')
          .replace(/\.0+$/, '') || '—';
        let type = '—';
        try { type = chequeType(transaction).type; } catch (e) { type = Number(transaction.debit || 0) > 0 ? 'EGRESO' : 'INGRESO'; }
        return [
          txt(bank),
          fmtDateX(transaction.date),
          txt(check),
          txt(type),
          money(txAmount(transaction, transaction.debit || 0)),
          money(txAmount(transaction, transaction.credit || 0))
        ];
      });
      y = drawTable(doc, {
        startY: y,
        head: ['Banco', 'Fecha', 'No. cheque', 'Tipo', 'Egreso', 'Ingreso'],
        numeric: [4, 5],
        colors: { 4: C.rose, 5: C.green },
        columnStyles: {
          0: { cellWidth: 72, fontStyle: 'bold' },
          1: { cellWidth: 32 },
          2: { cellWidth: 46 },
          3: { cellWidth: 41, fontStyle: 'bold' },
          4: { cellWidth: 41 }, 5: { cellWidth: 41 }
        },
        body: chequeBody.length ? chequeBody : [['—', '—', '—', '—', '—', '—']],
        foot: chequeBody.length ? [
          `TOTALES · ${chequeBody.length} movimiento(s)`, '', '', '',
          money(model.chequeDebit), money(model.chequeCredit)
        ] : null
      });
    }

    // ── Informe detallado (movimientos del período)
    if (on('informe')) {
      y = startSection(doc, y + 7, ++sec, 'Informe detallado de movimientos',
        `${model.informe.length} movimiento(s) · incluye país, banco, cuenta, referencia y descripción`, 26);
      const informeBody = model.informe.map((transaction, index) => {
        const debit = txAmount(transaction, transaction.debit || 0);
        const credit = txAmount(transaction, transaction.credit || 0);
        return [
          index + 1,
          fmtDateX(transaction.date),
          countryLabelX(countryOf(transaction)),
          txt(transaction.bank || '—'),
          txt(transaction.account || '—'),
          txt(transaction.reference || '—'),
          txt(String(transaction.description || '—')).slice(0, 220),
          debit ? money(debit) : '—',
          credit ? money(credit) : '—',
          money(txAmount(transaction, transaction.balance || 0))
        ];
      });
      drawTable(doc, {
        startY: y,
        head: ['#', 'Fecha', 'País', 'Banco', 'Cuenta', 'Referencia', 'Descripción', 'Débito', 'Crédito', 'Saldo'],
        numeric: [7, 8, 9],
        colors: { 7: C.rose, 8: C.green },
        columnStyles: {
          0: { cellWidth: 8, halign: 'center' },
          1: { cellWidth: 22, fontStyle: 'bold' },
          2: { cellWidth: 24 },
          3: { cellWidth: 40 },
          4: { cellWidth: 26 },
          5: { cellWidth: 34 },
          6: { cellWidth: 45 },
          7: { cellWidth: 22 }, 8: { cellWidth: 22 },
          9: { cellWidth: 30, fontStyle: 'bold' }
        },
        body: informeBody.length ? informeBody : [['—', '—', '—', '—', '—', '—', '—', '—', '—', '—']],
        foot: informeBody.length ? [
          // El rótulo va en "Fecha" (columna 0 es "#" y no cabe el texto)
          '', 'TOTALES', '', '', '', '',
          `${informeBody.length} movimiento(s)`,
          money(model.totalDebit),
          money(model.totalCredit),
          ''
        ] : null
      });
    }

    drawFooters(doc, model, generado, logo);

    doc.save(fileName(model));
    toast('ok', 'PDF descargado',
      `${mods.length} módulo(s) · ${model.informe.length} movimientos · ${model.countryLabel} · ` +
      `${fmtDateX(model.from)} al ${fmtDateX(model.to)}`);
  }

  // ─────────────────────────────────────────────────────── estados y avisos

  function toast(kind, title, detail) {
    const old = document.querySelector('.pdf-report-toast');
    if (old) old.remove();

    const el = document.createElement('div');
    el.className = `pdf-report-toast${kind === 'error' ? ' is-error' : ''}`;
    el.setAttribute('role', 'status');
    el.innerHTML = `<i></i><span><b></b><small></small></span>`;
    el.querySelector('b').textContent = title;
    el.querySelector('small').textContent = detail || '';
    document.body.appendChild(el);

    setTimeout(() => {
      el.classList.add('is-out');
      setTimeout(() => el.remove(), 260);
    }, 4200);
  }

  let busy = false;

  // Copia original del botón, para restaurarla al terminar
  const BTN_STRONG = BTN.querySelector('.pdf-report-copy strong');
  const BTN_SMALL = BTN.querySelector('.pdf-report-copy small');
  const BTN_IDLE = {
    strong: BTN_STRONG ? BTN_STRONG.textContent : '',
    small: BTN_SMALL ? BTN_SMALL.textContent : ''
  };

  async function runGeneration(mods) {
    if (busy) return;
    busy = true;

    BTN.classList.add('is-loading');
    BTN.disabled = true;
    if (BTN_STRONG) BTN_STRONG.textContent = 'Generando PDF…';
    if (BTN_SMALL) BTN_SMALL.textContent = `Incluyendo ${mods.length} módulo(s)`;

    try {
      await generatePdf(mods);
    } catch (error) {
      console.error('[PDF] No se pudo generar el reporte:', error);
      toast('error', 'No se pudo generar el PDF', error && error.message ? error.message : 'Error desconocido');
    } finally {
      BTN.classList.remove('is-loading');
      BTN.disabled = false;
      if (BTN_STRONG) BTN_STRONG.textContent = BTN_IDLE.strong;
      if (BTN_SMALL) BTN_SMALL.textContent = BTN_IDLE.small;
      busy = false;
    }
  }

  // ───────────────────────────────── ventana de selección múltiple de módulos

  const MODAL = document.getElementById('pdfModuleModal');
  const OPTIONS = document.getElementById('pdfModuleOptions');
  const COUNT = document.getElementById('pdfModuleCount');
  const GENERATE_BTN = document.getElementById('generatePdfModules');
  const GENERATE_COUNT = document.getElementById('pdfModuleGenerateCount');

  let closeTimer = null;
  let lastFocused = null;

  const selectedIds = () => (OPTIONS
    ? Array.from(OPTIONS.querySelectorAll('input:checked')).map(input => input.value)
    : MODULE_IDS.slice());

  function renderOptions() {
    if (!OPTIONS) return;
    OPTIONS.innerHTML = MODULES.map((module, index) => `
      <label class="pdf-module-option" data-module="${module.id}">
        <input type="checkbox" value="${module.id}" checked>
        <span class="pdf-module-box" aria-hidden="true"></span>
        <span class="pdf-module-icon" aria-hidden="true">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.9"
            stroke-linecap="round" stroke-linejoin="round">${module.icon}</svg>
        </span>
        <span class="pdf-module-copy">
          <strong>${module.title}</strong>
          <span>${module.hint}</span>
        </span>
        <span class="pdf-module-order" aria-hidden="true">${String(index + 1).padStart(2, '0')}</span>
      </label>`).join('');

    OPTIONS.querySelectorAll('input[type="checkbox"]').forEach(input => {
      input.addEventListener('change', refreshCount);
    });
  }

  function refreshCount() {
    if (!OPTIONS) return;
    const ids = selectedIds();

    OPTIONS.querySelectorAll('.pdf-module-option').forEach(option => {
      const input = option.querySelector('input');
      option.classList.toggle('is-checked', !!(input && input.checked));
    });

    if (COUNT) {
      COUNT.textContent = `${ids.length} de ${MODULES.length} módulos seleccionados`;
      COUNT.classList.toggle('is-empty', ids.length === 0);
    }
    if (GENERATE_COUNT) GENERATE_COUNT.textContent = String(ids.length);
    if (GENERATE_BTN) GENERATE_BTN.disabled = ids.length === 0;
  }

  function trapFocus(event) {
    const focusables = Array.from(MODAL.querySelectorAll('button, input, [tabindex]:not([tabindex="-1"])'))
      .filter(el => !el.disabled && el.getClientRects().length);
    if (!focusables.length) return;

    const first = focusables[0];
    const last = focusables[focusables.length - 1];
    if (event.shiftKey && document.activeElement === first) {
      event.preventDefault();
      last.focus();
    } else if (!event.shiftKey && document.activeElement === last) {
      event.preventDefault();
      first.focus();
    }
  }

  function openSelection() {
    // Sin ventana (o sin datos de la interfaz) se genera el reporte completo
    if (!MODAL || !OPTIONS) return runGeneration(MODULE_IDS.slice());
    if (busy) return;

    lastFocused = document.activeElement;
    clearTimeout(closeTimer);
    MODAL.classList.remove('is-closing');
    MODAL.hidden = false;
    document.body.style.overflow = 'hidden';
    refreshCount();

    const first = OPTIONS.querySelector('input');
    if (first) first.focus();
  }

  function closeSelection() {
    if (!MODAL || MODAL.hidden) return;

    MODAL.classList.add('is-closing');
    document.body.style.overflow = '';
    clearTimeout(closeTimer);
    closeTimer = setTimeout(() => {
      MODAL.hidden = true;
      MODAL.classList.remove('is-closing');
    }, 170);

    if (lastFocused && typeof lastFocused.focus === 'function' && lastFocused.isConnected) {
      lastFocused.focus();
    }
  }

  if (MODAL && OPTIONS) {
    renderOptions();

    document.getElementById('closePdfModules')?.addEventListener('click', closeSelection);
    document.getElementById('cancelPdfModules')?.addEventListener('click', closeSelection);

    // Click fuera del panel (sobre el fondo) cierra la ventana
    MODAL.addEventListener('click', event => {
      if (event.target === MODAL) closeSelection();
    });

    MODAL.querySelectorAll('[data-pdf-quick]').forEach(button => {
      button.addEventListener('click', () => {
        const checkAll = button.dataset.pdfQuick === 'all';
        OPTIONS.querySelectorAll('input[type="checkbox"]').forEach(input => {
          input.checked = checkAll;
        });
        refreshCount();
      });
    });

    GENERATE_BTN?.addEventListener('click', () => {
      const ids = selectedIds();
      if (!ids.length) return;
      closeSelection();
      runGeneration(ids);
    });

    document.addEventListener('keydown', event => {
      if (MODAL.hidden) return;
      if (event.key === 'Escape') {
        event.preventDefault();
        closeSelection();
        return;
      }
      if (event.key === 'Tab') trapFocus(event);
    });
  }

  BTN.addEventListener('click', openSelection);
})();
