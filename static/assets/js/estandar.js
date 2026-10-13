/* ══════════════════════════════════════════════════════════════════
   ESTÁNDAR VISUAL — gráficas Chart.js con el aspecto de SIMs
   Solo presentación y solo valores POR DEFECTO de Chart.js (tipografía,
   color de texto y cuadrícula, viñetas de leyenda y tooltip). No modifica
   las opciones de cada gráfica, ni datos, ni colores de las series.
   Se carga al final de cada dashboard, junto con assets/css/estandar.css.
   ══════════════════════════════════════════════════════════════════ */
(function () {
  function aplicar() {
    const C = window.Chart;
    if (!C || !C.defaults) return;
    const d = C.defaults;
    try {
      d.font.family = "'Montserrat', 'Segoe UI', sans-serif";
      d.font.size = 11;
      d.color = '#a9b9cf';
      d.borderColor = 'rgba(255, 255, 255, .08)';
      const lg = d.plugins.legend.labels;
      lg.usePointStyle = true;
      lg.pointStyle = 'rectRounded';
      lg.boxWidth = 9;
      lg.boxHeight = 9;
      const tt = d.plugins.tooltip;
      tt.backgroundColor = '#0b1321';
      tt.borderColor = 'rgba(128, 164, 207, .38)';
      tt.borderWidth = 1;
      tt.titleColor = '#f5f8fc';
      tt.bodyColor = '#a9b9cf';
      tt.padding = 10;
      tt.cornerRadius = 9;
      d.animation = Object.assign(d.animation || {}, { duration: 800, easing: 'easeOutQuart' });
    } catch (e) { /* versión de Chart.js sin alguno de estos valores: se deja como está */ }
  }
  aplicar();
  window.addEventListener('DOMContentLoaded', aplicar);
  // Inventario y Órdenes de compra cargan Chart.js después (ensureChartJs):
  // se aplican los valores en cuanto termina de cargar cualquier script.
  document.addEventListener('load', e => { if (e.target && e.target.tagName === 'SCRIPT') aplicar(); }, true);
})();

/* ══════════════════════════════════════════════════════════════════
   CONSULTA A LA API — igual que SIMs
   Detecta las consultas de datos (fetch, XHR y JSONP) sin tocar la lógica
   del dashboard y muestra el estado como SIMs:
   · pastilla bajo el título: "Consultando la API…" (con spinner) →
     "…puede tardar hasta un minuto" → "Actualizado HH:MM" /
     "Sin conexión con el servidor de datos";
   · secciones resaltadas mientras llegan los datos (.std-busy);
   · body.std-cargando para mostrar el spinner de SIMs en la primera carga.
   Los indicadores propios de cada dashboard se ocultan en estandar.css.
   ══════════════════════════════════════════════════════════════════ */
(function () {
  const ESTATICO = /\.(css|js|mjs|png|jpe?g|gif|svg|webp|ico|woff2?|ttf|otf|map)(\?|#|$)/i;
  const LIBRERIA = /fonts\.(googleapis|gstatic)\.com|cdn\.jsdelivr\.net|cdnjs\.cloudflare\.com|unpkg\.com/i;
  const PANELES = '.card, .chart-panel, .chart-card, .oc-card, .bill-card, .bill-panel, .apx-card, .pc-chart-card, .rrss-card, .content-panel, .kpi, .kpi-card, .kpi-card-exact, .oc-kpi, .bill-kpi, .stat-box, .apx-kpi, .table-wrap';
  let activas = 0, huboError = false, lentoTimer = null, pill = null, primeraCarga = true, huboConsulta = false;
  // Última acción del usuario (clic, cambio de filtro, tecla): solo las consultas que él
  // dispara resaltan las secciones, como "Actualizar datos" en SIMs. Las consultas
  // automáticas o en segundo plano (p. ej. la API en vivo de Ingresos tras el snapshot)
  // solo cambian la pastilla de estado; así el contenido no parpadea mientras responden.
  let ultimaAccion = 0;
  ['click', 'change', 'keydown'].forEach(ev => document.addEventListener(ev, () => { ultimaAccion = Date.now(); }, true));

  // Precarga: hasta que el HTML está listo no se pinta nada. Así ningún KPI, viñeta
  // o gráfica aparece antes que el spinner (antes se veían un instante y luego se ocultaban).
  document.documentElement.classList.add('std-precarga');

  const esDato = url => {
    try {
      const u = new URL(String(url), location.href);
      return /^https?:$/.test(u.protocol) && !ESTATICO.test(u.pathname) && !LIBRERIA.test(u.host);
    } catch (e) { return false; }
  };

  function crearPill() {
    if (pill || !document.body) return pill;
    const titulo = document.querySelector('h1, .cover-title, #pageTitle');
    if (!titulo) return null;
    pill = document.createElement('p');
    pill.className = 'std-sync';
    pill.setAttribute('role', 'status');
    pill.setAttribute('aria-live', 'polite');
    // Debajo del subtítulo (o del título si no hay subtítulo)
    let ancla = titulo;
    const sig = titulo.nextElementSibling;
    if (sig && !sig.matches('button, .btn, div:has(button)') && sig.textContent.trim().length < 200) ancla = sig;
    ancla.insertAdjacentElement('afterend', pill);
    return pill;
  }

  function pintar(estado, texto) {
    const p = crearPill();
    if (!p) return;
    p.classList.toggle('std-sync-busy', estado === 'busy');
    p.classList.toggle('std-sync-error', estado === 'error');
    p.textContent = texto;
  }

  function marcarPaneles(on) {
    document.querySelectorAll(on ? PANELES : '.std-busy').forEach(el => {
      if (!on) { el.classList.remove('std-busy'); return; }
      if (el.offsetParent !== null && !el.closest('.xf-panel, .mso-modal')) el.classList.add('std-busy');
    });
  }

  // Primera carga como SIMs: solo el encabezado y el spinner; todo lo que está debajo
  // del encabezado se oculta ([data-std-gate]) y aparece junto cuando llegan los datos.
  // Es un atributo y no una clase: varios dashboards reescriben className al pintar
  // (p. ej. el aviso de SAC) y con una clase el elemento quedaba visible antes de tiempo.
  function ocultarContenido() {
    const titulo = document.querySelector('h1, .cover-title, #pageTitle');
    if (!titulo) return;
    let cab = titulo.closest('.title-container, .page-header, .oc-header, .bill-header, header, .cab, .topbar, .hero, .header, .apx-header') || titulo.parentElement;
    while (!cab.nextElementSibling && cab.parentElement && cab.parentElement !== document.body) cab = cab.parentElement;
    for (let nodo = cab; nodo && nodo !== document.body; nodo = nodo.parentElement) {
      for (let sig = nodo.nextElementSibling; sig; sig = sig.nextElementSibling) {
        if (!sig.matches('script, style, link, .std-loader, .std-sync')) sig.setAttribute('data-std-gate', '');
      }
    }
  }

  // Primera carga: encabezado + spinner de SIMs; el resto espera a los datos
  function arrancarCarga() {
    if (!primeraCarga || !document.body) return;
    document.documentElement.classList.add('std-cargando');
    if (!document.querySelector('.std-loader')) document.body.insertAdjacentHTML('beforeend', '<div class="std-loader" aria-hidden="true"></div>');
    ocultarContenido();
  }
  // Muestra el contenido (una sola vez): cuando llegan los datos, si el dashboard
  // no consulta APIs, o por tope de tiempo
  function liberar() {
    if (!primeraCarga) return;
    primeraCarga = false;
    document.documentElement.classList.remove('std-cargando', 'std-precarga');
    document.dispatchEvent(new Event('std-listo'));
  }
  document.addEventListener('DOMContentLoaded', () => {
    arrancarCarga();
    document.documentElement.classList.remove('std-precarga');
    // Dashboards que no consultan APIs (cargan archivos): se muestran al terminar de cargar la página
    const sinApi = () => setTimeout(() => { if (!huboConsulta) liberar(); }, 700);
    if (document.readyState === 'complete') sinApi(); else window.addEventListener('load', sinApi, { once: true });
    setTimeout(liberar, 12000);
  });

  function inicio() {
    huboConsulta = true;
    activas++;
    if (activas !== 1) return;
    huboError = false;
    document.documentElement.classList.add('std-consultando');
    if (primeraCarga) {
      arrancarCarga();
      // Tope: si una API tarda demasiado, el contenido se muestra igual (la pastilla sigue consultando)
      setTimeout(liberar, 6000);
    }
    pintar('busy', 'Consultando la API…');
    if (!primeraCarga && Date.now() - ultimaAccion < 1500) marcarPaneles(true);
    clearTimeout(lentoTimer);
    lentoTimer = setTimeout(() => { if (activas) pintar('busy', 'Consultando la API… puede tardar hasta un minuto'); }, 2500);
  }

  function fin(error) {
    if (error) huboError = true;
    activas = Math.max(0, activas - 1);
    if (activas) return;
    clearTimeout(lentoTimer);
    // Pequeña espera: varias consultas seguidas se ven como una sola
    setTimeout(() => {
      if (activas) return;
      document.documentElement.classList.remove('std-consultando');
      liberar();
      marcarPaneles(false);
      const hora = new Date().toLocaleTimeString('es-GT', { hour: '2-digit', minute: '2-digit' });
      if (huboError) pintar('error', 'Sin conexión con el servidor de datos');
      else pintar('ok', 'Actualizado ' + hora);
    }, 150);
  }

  // fetch
  if (window.fetch) {
    const fetchOriginal = window.fetch.bind(window);
    window.fetch = function (recurso, opciones) {
      const url = recurso && recurso.url ? recurso.url : recurso;
      if (!esDato(url)) return fetchOriginal(recurso, opciones);
      inicio();
      return fetchOriginal(recurso, opciones).then(
        r => { fin(false); return r; },
        e => { fin(true); throw e; }
      );
    };
  }

  // XMLHttpRequest
  const abrir = XMLHttpRequest.prototype.open, enviar = XMLHttpRequest.prototype.send;
  XMLHttpRequest.prototype.open = function (metodo, url) { this._stdDato = esDato(url); return abrir.apply(this, arguments); };
  XMLHttpRequest.prototype.send = function () {
    if (this._stdDato) {
      inicio();
      this.addEventListener('loadend', () => fin(this.status === 0), { once: true });
    }
    return enviar.apply(this, arguments);
  };

  // JSONP (p. ej. Google Sheets en GPS): <script src="…callback=…">
  const esJsonp = s => s.tagName === 'SCRIPT' && s.src && /[?&](callback|tqx|jsonp)=|script\.google\.com|docs\.google\.com\/spreadsheets/i.test(s.src);
  new MutationObserver(cambios => cambios.forEach(c => c.addedNodes.forEach(n => {
    if (n.nodeType !== 1 || !esJsonp(n)) return;
    inicio();
    n.addEventListener('load', () => fin(false), { once: true });
    n.addEventListener('error', () => fin(true), { once: true });
  }))).observe(document.documentElement, { childList: true, subtree: true });
  // Los dashboards que no consultan APIs (p. ej. los que cargan archivos) no muestran estado.
})();

/* ══════════════════════════════════════════════════════════════════
   ANIMACIONES — las de SIMs en todos los dashboards
   · Revelado al entrar en pantalla (y se oculta al salir), igual que SIMs:
     fundido + 18px hacia arriba; KPIs en cascada de .08s.
   · Las filas de tabla entran en cascada (estandar.css).
   · Se desactivan las animaciones de entrada propias de cada dashboard.
   ══════════════════════════════════════════════════════════════════ */
(function () {
  if (!('IntersectionObserver' in window) || matchMedia('(prefers-reduced-motion: reduce)').matches) return;
  const KPIS = '.kpi, .kpi-card, .kpi-card-exact, .oc-kpi, .bill-kpi, .stat-box, .apx-kpi';
  const BLOQUES = '.content-panel, .op-card, .sims-fbar, .card, .chart-panel, .chart-card, .oc-card, .bill-card, .bill-panel, .apx-card, .pc-chart-card, .rrss-card, .inventory-table-card, .oc-reveal, '
    + '.inventory-filters-v2, .oc-filters-v2, .apx-filters, #filters.filter-demo, .filters-box-card, #fx-main, .gps-filter-toolbar, .top-toolbar-left';
  // Animaciones de entrada propias (por nombre de @keyframes): se reemplazan por el revelado de SIMs
  const ENTRADAS = new Set(['inventory-row-in', 'aseEntra', 'entra', 'fundido', 'cmpFadeUp', 'fnFadeUp', 'fnFunnelIn',
    'gps-card-in', 'fadeIn', 'catRise', 'flowRise', 'resRise', 'resFade', 'resDetailIn', 'gps-theme-reveal', 'brillo', 'cmpMedalPop', 'cmpTrophyBounce']);

  document.documentElement.classList.add('std-anim');
  const io = new IntersectionObserver(entradas => entradas.forEach(e => {
    const el = e.target;
    if (e.isIntersecting) {
      const entra = !el.classList.contains('std-in');
      el.classList.add('std-in'); el.classList.remove('std-out');
      if (entra && !document.documentElement.classList.contains('std-cargando')) animarGraficas(el);
    }
    else if (el.classList.contains('std-in')) { el.classList.remove('std-in'); el.classList.add('std-out'); }
  }), { threshold: 0.08 });

  // Chart.js: reinicia y vuelve a animar; Ingresos (canvas propio): drawAll()
  let dibujarTodo = null;
  function animarGraficas(raiz) {
    const C = window.Chart;
    raiz.querySelectorAll('canvas').forEach(cv => {
      const ch = C && C.getChart ? C.getChart(cv) : null;
      if (ch) { try { ch.reset(); ch.update(); } catch (e) { /* gráfica destruida */ } }
      else if (typeof window.drawAll === 'function') {
        clearTimeout(dibujarTodo);
        dibujarTodo = setTimeout(() => { try { window.drawAll(); } catch (e) { /* sin datos aún */ } }, 60);
      }
    });
  }
  document.addEventListener('std-listo', () => document.querySelectorAll('.std-rv.std-in').forEach(animarGraficas));

  function registrar(raiz) {
    const lista = raiz.matches && raiz.matches(KPIS + ',' + BLOQUES) ? [raiz] : [];
    lista.push(...(raiz.querySelectorAll ? raiz.querySelectorAll(KPIS + ',' + BLOQUES) : []));
    lista.forEach(el => {
      if (el.dataset.stdRv || el.closest('.xf-panel, .mso-modal, [role="dialog"]')) return;
      // Un bloque dentro de otro bloque ya animado no se anima dos veces
      if (el.parentElement && el.parentElement.closest('.std-rv') && !el.matches(KPIS)) return;
      el.dataset.stdRv = '1';
      if (el.matches(KPIS)) {
        const hermanos = [...el.parentElement.children].filter(h => h.matches(KPIS));
        el.style.transitionDelay = (Math.max(0, hermanos.indexOf(el)) * 0.08).toFixed(2) + 's';
      }
      el.classList.add('std-rv', 'std-out');
      requestAnimationFrame(() => io.observe(el));
    });
  }

  function quitarEntradas(raiz) {
    const lista = [raiz, ...(raiz.querySelectorAll ? raiz.querySelectorAll('*') : [])];
    lista.forEach(el => {
      if (el.nodeType !== 1 || el.tagName === 'TR') return;
      const nombres = getComputedStyle(el).animationName;
      if (!nombres || nombres === 'none' || !nombres.split(',').some(n => ENTRADAS.has(n.trim()))) return;
      el.style.animation = 'none';
      if (getComputedStyle(el).opacity === '0') el.style.opacity = '1';
    });
  }

  let pendientes = new Set(), timer = null;
  const procesar = () => { timer = null; const raices = [...pendientes]; pendientes = new Set(); raices.forEach(r => { if (r.isConnected) { quitarEntradas(r); registrar(r); } }); };
  new MutationObserver(cambios => {
    cambios.forEach(c => c.addedNodes.forEach(n => { if (n.nodeType === 1) pendientes.add(n); }));
    if (pendientes.size && !timer) timer = setTimeout(procesar, 120);
  }).observe(document.documentElement, { childList: true, subtree: true });
  const inicio = () => { quitarEntradas(document.body); registrar(document.body); };
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', inicio); else inicio();
})();
