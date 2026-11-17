/* App de campo COR-GPPT-R006 — funciona sin cuentas y sin señal.
   Los pilotos y las fotos se guardan en el propio dispositivo (IndexedDB).
   El Excel V2 se genera en el dispositivo con excel.js + plantilla.js. */
(function () {
'use strict';
var STD = COR.STD, FIXED8 = COR.FIXED8, CRIT9 = COR.CRIT9;
var CRIT_OPTS = [['Sí', 'Sí', 'si'], ['Parcialmente o requiere ajustes', 'Parcial / ajustes', 'par'], ['No', 'No', 'no'], ['No aplica', 'No aplica', '']];

/* ---------- Formato (claves internas iguales a las del formulario en Claude) ---------- */
var F1 = [
  { k: 'nombre', label: 'Nombre de la iniciativa' },
  { k: 'codigo', label: 'Código o referencia', prompt: 'Código de Wave o número de iniciativa.' },
  { k: 'fechas', label: 'Fecha de la prueba', type: 'range' },
  { k: 'lugar', label: 'Lugar de ejecución' },
  { k: 'responsable', label: 'Dueño de iniciativa' },
  { k: 'encargado', label: 'Encargado técnico' },
  { k: 'objetivo', label: 'Objetivo de la prueba piloto', type: 'area' },
  { k: 'partes', label: 'Partes interesadas que pueden ser o sentirse impactadas por la ejecución de la iniciativa', type: 'area' },
  { k: 'resultado', label: 'Resultado esperado', type: 'area' },
  { k: 'evidenciaCom', label: 'Evidencia de comunicación a las partes interesadas (indique cuál o su referencia)', type: 'area' },
  { k: 'alcance', label: 'Alcance del piloto', type: 'area' }
];
var F2 = [
  { k: 'hipotesis', label: 'Hipótesis de valor o supuesto a validar', type: 'area' },
  { k: 'lineaBase', label: 'Línea base o situación inicial', type: 'area', prompt: 'Estado actual con datos disponibles: producción, tiempo, productividad, calidad, costo u otra variable relevante.' },
  { k: 'metodo', label: 'Método de evaluación', type: 'area', prompt: 'Mediciones, observaciones, análisis, muestras, comparaciones o fuentes de evidencia.' }
];
var F10 = [
  { k: 'decision', label: 'Decisión', type: 'choice', options: ['Implementar', 'Implementar con ajustes', 'Realizar nueva prueba'] },
  { k: 'cronograma', label: 'Cronograma de implementación elaborado', type: 'choice', options: ['Sí', 'No', 'Pendiente'] },
  { k: 'frec', label: 'Frecuencia de reporte de implementación y de KPI' },
  { k: 'kpiDef', label: 'KPI de implementación definidos', type: 'choice', options: ['Sí', 'No', 'Pendiente'] },
  { k: 'resp', label: 'Responsable del seguimiento' }
];
var F11 = [
  { k: 'ordenEjec', label: 'Orden de ejecución de la prueba piloto' },
  { k: 'ordenLiq', label: 'Orden de liquidación de la prueba piloto' },
  { k: 'fechaCierre', label: 'Fecha de cierre', type: 'date' }
];
var ROLES = [['elab', 'Elaborado por'], ['rev', 'Revisado por'], ['aprob', 'Aprobado por']];
var COLS = {
  s3: [{ k: 'etapa', label: 'Etapa o actividad' }, { k: 'desc', label: 'Descripción', type: 'area' }, { k: 'tiempo', label: 'Tiempo' }, { k: 'equipo', label: 'Equipo o herramienta' }],
  s4: [{ k: 'recurso', label: 'Recurso o insumo' }, { k: 'cantidad', label: 'Cantidad', mode: 'decimal' }, { k: 'unidad', label: 'Unidad' }, { k: 'porcentaje', label: 'Porcentaje (opcional)', optional: true }, { k: 'obs', label: 'Observación (opcional)', optional: true }],
  s5: [{ k: 'aspecto', label: 'Etapa o aspecto', plan: true }, { k: 'parametro', label: 'Parámetro o condición', plan: true }, { k: 'valor', label: 'Valor o rango', big: true }, { k: 'obs', label: 'Observación (opcional)', optional: true }],
  s7: [{ k: 'kpi', label: 'KPI o resultado', plan: true }, { k: 'evalua', label: 'Qué evalúa', plan: true }, { k: 'meta', label: 'Meta o criterio (con unidad)', plan: true }, { k: 'resultado', label: 'Resultado obtenido (con unidad)', big: true }, { k: 'cumple', label: 'Cumplimiento', choice: ['Sí', 'Parcial', 'No'] }],
  s13: [{ k: 'anexo', label: 'Anexo o evidencia' }, { k: 'desc', label: 'Descripción o referencia', type: 'area' }]
};
var ROWNAME = { s3: 'Etapa', s4: 'Recurso', s5: 'Parámetro', s7: 'KPI / resultado', s13: 'Anexo' };

var PHASES = {
  antes: { name: 'Antes de la visita', icon: '📋', desc: 'Datos generales, metodología y qué se va a medir', views: [
    { s: 's1', t: '1 · Datos generales' },
    { s: 's2', t: '2 · Metodología de evaluación', note: 'Defina cómo se comprobará el desempeño de la prueba y cuál será el punto de comparación.' },
    { s: 's7', m: 'plan', t: '6 · KPI y resultados: qué se va a medir', note: 'Defina los KPI y resultados con su meta. En campo solo se anotará el valor obtenido.' },
    { s: 's5', m: 'plan', t: '5 · Parámetros a controlar', note: 'Liste las condiciones que influyen en la prueba (temperatura, pH, Brix, presión, tiempo…). En campo solo se anotará el valor.' }
  ]},
  campo: { name: 'En campo', icon: '🏭', desc: 'Procedimiento, insumos, mediciones, resultados, dificultades y fotos', views: [
    { s: 's3', t: '3 · Procedimiento realizado', note: 'Etapas principales en el orden en que se ejecutaron.' },
    { s: 's4', t: '4 · Recursos e insumos', note: 'Materias primas, materiales, reactivos, equipos, servicios. Si hay formulación, inclúyala aquí.' },
    { s: 's5', m: 'campo', t: '5 · Condiciones y parámetros: valores', note: 'Anote el valor o rango medido de cada parámetro.' },
    { s: 's7', m: 'campo', t: '6 · KPI y resultados: valores obtenidos', note: 'Anote el resultado (con unidad) y marque si cumple la meta.' },
    { s: 's8', t: '7 · Riesgos, dificultades y mejoras' },
    { s: 'fotos', t: 'Fotos de la visita', note: 'Displays, hojas de laboratorio, bitácora, equipo. Se agregan solas como anexos y en una hoja del Excel.' }
  ]},
  cierre: { name: 'Cierre', icon: '✅', desc: 'Evaluación, seguimiento, cierre, validación y anexos', views: [
    { s: 's9', t: '8 · Evaluación de la prueba piloto', note: 'Marque una opción por criterio y complemente la conclusión con la evidencia disponible.' },
    { s: 's10', t: '9 · Seguimiento de implementación', note: 'Solo si la prueba es favorable y se autoriza avanzar.' },
    { s: 's11', t: '10 · Cierre administrativo y contable', note: 'Referencias de cierre según los controles internos.' },
    { s: 's12', t: '11 · Validación del registro', note: 'La firma se estampa en el documento impreso.' },
    { s: 's13', t: '12 · Anexos y evidencias', note: 'Hojas de datos, resultados de laboratorio, gráficos, cálculos, cotizaciones, bitácoras. Las fotos de la app se agregan solas.' }
  ]}
};
var PH_ORDER = ['antes', 'campo', 'cierre'];

/* ---------- Utilidades ---------- */
function esc(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); }
function clone(o) { return JSON.parse(JSON.stringify(o || {})); }
function get(o, path) { return path.split('.').reduce(function (a, k) { return a == null ? undefined : a[k]; }, o); }
function set(o, path, val) {
  var ks = path.split('.'), a = o;
  for (var i = 0; i < ks.length - 1; i++) {
    var k = ks[i], nk = ks[i + 1];
    if (a[k] == null || typeof a[k] !== 'object') a[k] = /^\d+$/.test(nk) ? [] : {};
    a = a[k];
  }
  a[ks[ks.length - 1]] = val;
}
function blank(v) { return v == null || String(v).trim() === ''; }
function uid() { return 'p' + Date.now().toString(36) + Math.random().toString(36).slice(2, 7); }
function today() { var d = new Date(); return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0'); }
function fmtTs(iso) { if (!iso) return ''; var d = new Date(iso); return d.toLocaleDateString('es-HN', { day: '2-digit', month: 'short' }) + ' ' + d.toLocaleTimeString('es-HN', { hour: '2-digit', minute: '2-digit' }); }
function toast(msg) {
  var old = document.querySelector('.toast'); if (old) old.remove();
  var t = document.createElement('div'); t.className = 'toast'; t.textContent = msg; document.body.appendChild(t);
  setTimeout(function () { t.remove(); }, 3000);
}
function safeName(s) { return String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^\w\-]+/g, '_').replace(/_+/g, '_').replace(/^_|_$/g, '').slice(0, 60); }

/* ---------- Almacenamiento local (IndexedDB) ---------- */
var DB = null, PIL = {}, URLS = {};
function openDb() {
  return new Promise(function (res) {
    try {
      var rq = indexedDB.open('cor-campo', 1);
      rq.onupgradeneeded = function () {
        var db = rq.result;
        db.createObjectStore('pilotos', { keyPath: 'id' });
        db.createObjectStore('fotos', { keyPath: 'id' });
      };
      rq.onsuccess = function () { res(rq.result); };
      rq.onerror = function () { res(null); };
    } catch (e) { res(null); }
  });
}
function tx(store, mode, fn) {
  return new Promise(function (res, rej) {
    if (!DB) return rej(new Error('sin almacenamiento'));
    var t = DB.transaction(store, mode), st = t.objectStore(store), out;
    out = fn(st);
    t.oncomplete = function () { res(out && out.result !== undefined ? out.result : undefined); };
    t.onerror = function () { rej(t.error); };
  });
}
function loadAll() { return tx('pilotos', 'readonly', function (st) { return st.getAll(); }).then(function (arr) { (arr || []).forEach(function (p) { PIL[p.id] = p; }); }); }
var saveTimers = {};
function save(p, now) {
  p.updatedAt = new Date().toISOString();
  clearTimeout(saveTimers[p.id]);
  var run = function () { return tx('pilotos', 'readwrite', function (st) { st.put(p); }).catch(function () { toast('No se pudo guardar en este dispositivo.'); }); };
  if (now) return run();
  saveTimers[p.id] = setTimeout(run, 350);
  flashSaved();
  return Promise.resolve();
}
function flashSaved() { var el = document.getElementById('savedFlag'); if (el) { el.textContent = 'Guardado en este dispositivo ✓'; } }
function putPhoto(id, pid, blob) { return tx('fotos', 'readwrite', function (st) { st.put({ id: id, pid: pid, blob: blob }); }); }
function getPhoto(id) { return tx('fotos', 'readonly', function (st) { return st.get(id); }).then(function (r) { return r && r.blob; }); }
function delPhoto(id) { if (URLS[id]) { URL.revokeObjectURL(URLS[id]); delete URLS[id]; } return tx('fotos', 'readwrite', function (st) { st.delete(id); }); }
function photoUrl(id) {
  if (URLS[id]) return Promise.resolve(URLS[id]);
  return getPhoto(id).then(function (b) { if (!b) return ''; URLS[id] = URL.createObjectURL(b); return URLS[id]; });
}

/* ---------- Revisión de faltantes ---------- */
function analyze(p) {
  var d = p.data || {}, out = { total: 0, miss: [], byView: {} };
  PH_ORDER.forEach(function (ph) {
    PHASES[ph].views.forEach(function (vw, i) {
      var key = ph + ':' + i, m = [], tot = 0;
      var chk = function (ok, t) { tot++; if (!ok) m.push(t); };
      viewChecks(p, d, vw, chk);
      out.total += tot;
      out.byView[key] = { tot: tot, miss: m.length };
      m.forEach(function (t) { out.miss.push({ ph: ph, i: i, t: t, title: vw.t }); });
    });
  });
  out.pct = out.total ? Math.round((out.total - out.miss.length) / out.total * 100) : 0;
  return out;
}
function viewChecks(p, d, vw, chk) {
  var s = vw.s, sd = d[s] || {};
  var fields = function (list) {
    list.forEach(function (f) {
      if (f.type === 'range') { chk(!blank(sd.inicio), 'Fecha de inicio'); chk(!blank(sd.fin), 'Fecha final'); }
      else chk(!blank(sd[f.k]), f.label);
    });
  };
  if (s === 's1') return fields(F1);
  if (s === 's2') return fields(F2);
  if (s === 's11') return fields(F11);
  if (s === 's10') { if (p.gate === 'no') return; return fields(F10); }
  if (s === 'fotos') return;
  if (COLS[s]) {
    var rows = sd.rows || [];
    if (s === 's13' && (p.fotos || []).length) { if (!rows.length) return; }
    else if (vw.m !== 'campo') chk(sd.na || rows.length > 0, 'Agregue al menos una fila o marque "No aplica"');
    if (sd.na) return;
    rows.forEach(function (r, i) {
      COLS[s].forEach(function (c) {
        if (c.optional) return;
        if (vw.m === 'plan' && !c.plan) return;
        if (vw.m === 'campo' && c.plan) return;
        chk(!blank(r[c.k]), ROWNAME[s] + ' ' + (i + 1) + ': ' + c.label.replace(/ \(.*\)$/, ''));
      });
    });
    return;
  }
  if (s === 's8') {
    var fx = sd.fixed || {};
    FIXED8.forEach(function (a) { var x = fx[a.k] || {}; chk(!blank(x.sit) && !blank(x.accion) && !blank(x.resp), a.label); });
    (sd.otros || []).forEach(function (r, i) { chk(!blank(r.aspecto) && !blank(r.sit) && !blank(r.accion) && !blank(r.resp), 'Otro aspecto ' + (i + 1)); });
    return;
  }
  if (s === 's9') {
    chk(!!p.gate, '¿La prueba es favorable y se autoriza avanzar?');
    CRIT9.forEach(function (q, i) { chk(!blank((sd.crit || [])[i]), q); });
    chk(!blank(sd.grafico), 'Comparación línea base vs. resultados');
    chk(!blank(sd.conclusion), 'Conclusión y recomendación');
    return;
  }
  if (s === 's12') {
    ROLES.forEach(function (r) { chk(!blank(get(sd, r[0] + '.nombre')), r[1] + ': nombre'); chk(!blank(get(sd, r[0] + '.fecha')), r[1] + ': fecha'); });
  }
}
/* Completa todos los vacíos (salvo decisiones que la persona debe tomar) con un valor estándar */
function fillBlanks(p, val) {
  var d = p.data, n = 0;
  var fill = function (o, k) { if (blank(o[k])) { o[k] = val; n++; } };
  [['s1', F1], ['s2', F2], ['s11', F11]].concat(p.gate === 'no' ? [] : [['s10', F10]]).forEach(function (x) {
    var sd = d[x[0]] = d[x[0]] || {};
    x[1].forEach(function (f) {
      if (f.type === 'range') { fill(sd, 'inicio'); fill(sd, 'fin'); }
      else if (f.type === 'choice') { if (blank(sd[f.k]) && (val === 'No aplica' || f.options.indexOf(val) >= 0)) { sd[f.k] = val; n++; } }
      else fill(sd, f.k);
    });
  });
  Object.keys(COLS).forEach(function (s) {
    var sd = d[s] = d[s] || {};
    if (s === 's13' && (p.fotos || []).length && !(sd.rows || []).length) return;
    if (!sd.na && !(sd.rows || []).length) { if (val === 'No aplica') { sd.na = true; n++; } return; }
    if (sd.na) return;
    sd.rows.forEach(function (r) { COLS[s].forEach(function (c) { if (!c.optional && !c.choice) fill(r, c.k); }); });
  });
  var s8 = d.s8 = d.s8 || {}; s8.fixed = s8.fixed || {};
  FIXED8.forEach(function (a) { var x = s8.fixed[a.k] = s8.fixed[a.k] || {}; ['sit', 'accion', 'resp'].forEach(function (c) { fill(x, c); }); });
  (s8.otros || []).forEach(function (r) { ['sit', 'accion', 'resp'].forEach(function (c) { fill(r, c); }); });
  var s9 = d.s9 = d.s9 || {};
  fill(s9, 'grafico'); fill(s9, 'conclusion');
  if (val === 'No aplica') { s9.crit = s9.crit || []; CRIT9.forEach(function (q, i) { if (blank(s9.crit[i])) { s9.crit[i] = 'No aplica'; n++; } }); }
  var s12 = d.s12 = d.s12 || {};
  ROLES.forEach(function (r) { var x = s12[r[0]] = s12[r[0]] || {}; fill(x, 'nombre'); fill(x, 'fecha'); });
  return n;
}

/* ---------- Navegación (con el botón "atrás" del celular) ---------- */
var R = { v: 'home' };
function parseHash() {
  var h = location.hash.replace(/^#/, ''), o = { v: 'home' };
  h.split('&').forEach(function (kv) { var p = kv.split('='); if (p[0]) o[p[0]] = decodeURIComponent(p[1] || ''); });
  if (o.i != null) o.i = +o.i;
  return o;
}
function go(o, replace) {
  var h = Object.keys(o).filter(function (k) { return o[k] != null && o[k] !== ''; }).map(function (k) { return k + '=' + encodeURIComponent(o[k]); }).join('&');
  if (replace) { history.replaceState(null, '', '#' + h); route(); }
  else if (location.hash === '#' + h) route();
  else location.hash = h;
}
function route() {
  R = parseHash();
  if (R.p && !PIL[R.p]) R = { v: 'home' };
  render();
  window.scrollTo(0, 0);
}
window.addEventListener('hashchange', route);

/* ---------- Render ---------- */
var app = document.getElementById('app');
function top(title, sub, backTo, actHtml) {
  return '<header class="top">' + (backTo ? '<button class="back" data-go="' + esc(backTo) + '" aria-label="Volver">←</button>' : '') +
    '<h1>' + esc(title) + (sub ? '<small>' + esc(sub) + '</small>' : '') + '</h1>' + (actHtml || '') + '</header>';
}
function backHash(o) { return Object.keys(o).map(function (k) { return k + '=' + encodeURIComponent(o[k]); }).join('&'); }
function render() {
  var p = R.p && PIL[R.p];
  if (R.v === 'home' || !p) return renderHome();
  if (R.v === 'pilot') return renderPilot(p);
  if (R.v === 'phase') return renderPhase(p, R.ph);
  if (R.v === 'view') return renderView(p, R.ph, R.i);
  if (R.v === 'review') return renderReview(p);
  renderHome();
}

var installEvt = null;
window.addEventListener('beforeinstallprompt', function (e) { e.preventDefault(); installEvt = e; if (R.v === 'home') renderHome(); });

function renderHome() {
  var list = Object.keys(PIL).map(function (k) { return PIL[k]; }).sort(function (a, b) { return (b.updatedAt || '').localeCompare(a.updatedAt || ''); });
  var h = top('Pruebas piloto', 'COR-GPPT-R006 · registro de campo') + '<main class="wrap">';
  if (!DB) h += '<div class="banner warn">Este navegador no permite guardar datos. Abra la app en Chrome (no en modo incógnito).</div>';
  h += '<div class="row" style="margin-bottom:14px"><button class="btn primary" data-act="new">+ Nuevo piloto</button>' +
    '<label class="btn sky filebtn">📥 Abrir piloto recibido<input type="file" accept=".json,application/json" data-import></label></div>';
  if (installEvt) h += '<button class="btn block sky" data-act="install" style="margin-bottom:14px">📲 Instalar en este celular</button>';
  if (!list.length) {
    h += '<div class="empty">Aún no hay pilotos en este celular.<br>Toque <b>+ Nuevo piloto</b> para empezar.</div>';
  } else {
    list.forEach(function (p) {
      var a = analyze(p), d = p.data || {};
      h += '<button class="pilot" data-go="v=pilot&p=' + p.id + '"><b>' + esc(get(d, 's1.nombre') || 'Sin nombre') + '</b>' +
        '<div class="m"><span>' + esc(get(d, 's1.codigo') || '—') + ' · ' + (p.fotos || []).length + ' foto(s)</span><span>' + a.pct + ' % · ' + fmtTs(p.updatedAt) + '</span></div>' +
        '<div class="prog"><i style="width:' + a.pct + '%"></i></div></button>';
    });
  }
  h += '<div class="card" style="margin-top:18px"><h3>Cómo se usa</h3><ol class="steps">' +
    '<li><b>Antes de la visita:</b> datos generales, metodología y lo que se va a medir (KPI con su meta y parámetros).</li>' +
    '<li><b>En campo:</b> procedimiento, insumos, valores medidos, dificultades y fotos. Funciona sin señal.</li>' +
    '<li><b>Cierre:</b> evaluación, seguimiento, cierre y validación.</li>' +
    '<li><b>Revisar y generar Excel:</b> la app muestra lo que falta y genera el formato V2 para compartir por WhatsApp, correo u OneDrive.</li></ol>' +
    '<p class="muted" style="margin:6px 0 0">Los datos se guardan solo en este celular. Para que otra persona continúe un piloto, use <b>Compartir piloto</b> y ella lo abre con <b>Abrir piloto recibido</b>.</p></div>';
  h += '</main>';
  app.innerHTML = h;
}

function renderPilot(p) {
  var d = p.data || {}, a = analyze(p);
  var h = top(get(d, 's1.nombre') || 'Piloto', (get(d, 's1.codigo') || 'Sin código') + ' · ' + a.pct + ' % completo', 'v=home', '<button class="act" data-act="menu">⋯</button>') + '<main class="wrap">';
  PH_ORDER.forEach(function (ph) {
    var P = PHASES[ph], tot = 0, miss = 0;
    P.views.forEach(function (vw, i) { var b = a.byView[ph + ':' + i]; tot += b.tot; miss += b.miss; });
    var badge = !tot ? '' : miss ? '<span class="badge warn">Faltan ' + miss + '</span>' : '<span class="badge ok">Completo</span>';
    if (ph === 'campo') badge += ' <span class="badge n">📷 ' + (p.fotos || []).length + '</span>';
    h += '<button class="phase" data-go="v=phase&p=' + p.id + '&ph=' + ph + '"><span class="ic">' + P.icon + '</span><span class="tx"><b>' + P.name + '</b><span>' + P.desc + '</span><span class="bl">' + badge + '</span></span><span class="chev">›</span></button>';
  });
  h += '<div class="dock"><div class="in-dock"><button class="btn primary" data-go="v=review&p=' + p.id + '">📊 Revisar y generar Excel</button></div></div>';
  h += '</main>';
  app.innerHTML = h;
}

function renderPhase(p, ph) {
  var P = PHASES[ph]; if (!P) return renderPilot(p);
  var a = analyze(p);
  var h = top(P.name, get(p.data, 's1.nombre') || '', 'v=pilot&p=' + p.id) + '<main class="wrap">';
  P.views.forEach(function (vw, i) {
    var b = a.byView[ph + ':' + i], badge;
    if (vw.s === 'fotos') badge = '<span class="badge n">' + (p.fotos || []).length + '</span>';
    else if (vw.s === 's10' && p.gate === 'no') badge = '<span class="badge n">No aplica</span>';
    else badge = b.miss ? '<span class="badge warn">' + b.miss + '</span>' : '<span class="badge ok">✓</span>';
    var parts = vw.t.split(' · ');
    h += '<button class="secitem" data-go="v=view&p=' + p.id + '&ph=' + ph + '&i=' + i + '"><span class="num">' + (parts.length > 1 ? esc(parts[0]) : '📷') + '</span><span class="tx">' + esc(parts.length > 1 ? parts.slice(1).join(' · ') : vw.t) + '</span>' + badge + '</button>';
  });
  h += '</main>';
  app.innerHTML = h;
}

/* --- Controles --- */
function inputHtml(path, val, o) {
  o = o || {};
  val = val == null ? '' : String(val);
  var attrs = ' data-p="' + esc(path) + '" class="in' + (o.big ? ' big' : '') + '"';
  if (o.type === 'area') return '<textarea' + attrs + ' rows="' + (o.rows || 3) + '">' + esc(val) + '</textarea>';
  if (o.type === 'date') {
    var isDate = /^\d{4}-\d{2}-\d{2}$/.test(val);
    return '<input type="date"' + attrs + ' value="' + (isDate ? esc(val) : '') + '">' + (val && !isDate ? '<div class="muted" style="margin-top:4px">Valor: <b>' + esc(val) + '</b></div>' : '');
  }
  return '<input type="text"' + (o.mode ? ' inputmode="' + o.mode + '"' : '') + attrs + ' value="' + esc(val) + '" autocomplete="off">';
}
function chips(path, val) {
  return '<div class="chips">' + STD.map(function (s) { return '<button class="chip' + (val === s ? ' sel' : '') + '" data-setp="' + esc(path) + '" data-val="' + esc(s) + '">' + s + '</button>'; }).join('') + '</div>';
}
function seg(path, options, val, extraStd) {
  var opts = options.map(function (o) { return Array.isArray(o) ? o : [o, o, '']; });
  if (extraStd) STD.forEach(function (s) { if (!opts.some(function (o) { return o[0] === s; })) opts.push([s, s, '']); });
  return '<div class="seg">' + opts.map(function (o) {
    return '<button class="' + (val === o[0] ? 'sel ' + o[2] : '') + '" data-setp="' + esc(path) + '" data-val="' + esc(o[0]) + '">' + esc(o[1]) + '</button>';
  }).join('') + '</div>';
}
function fieldBlock(sk, f, sd) {
  var path = sk + '.' + f.k, h = '<div class="f"><label>' + esc(f.label) + '</label>' + (f.prompt ? '<div class="prompt">' + esc(f.prompt) + '</div>' : '');
  if (f.type === 'range') {
    h += '<div class="row"><div style="flex:1;min-width:140px"><div class="muted">Inicio</div>' + inputHtml(sk + '.inicio', sd.inicio, { type: 'date' }) + '</div>' +
      '<div style="flex:1;min-width:140px"><div class="muted">Final</div>' + inputHtml(sk + '.fin', sd.fin, { type: 'date' }) + '</div></div>';
    return h + '</div>';
  }
  if (f.type === 'choice') return h + seg(path, f.options, sd[f.k], true) + '</div>';
  h += inputHtml(path, sd[f.k], { type: f.type });
  if (f.k !== 'nombre' && f.k !== 'codigo') h += chips(path, sd[f.k]);
  return h + '</div>';
}

function tableView(p, s, mode) {
  var d = p.data, sd = d[s] = d[s] || {}, rows = sd.rows || [], h = '';
  if (sd.na) {
    return '<label class="nacheck"><input type="checkbox" data-na="' + s + '" checked> Esta sección no aplica</label><div class="empty">Sección marcada como "No aplica". En el Excel aparecerá así.</div>';
  }
  if (s === 's13' && (p.fotos || []).length) h += '<div class="banner info">Las ' + p.fotos.length + ' foto(s) de la app se agregan solas como anexos. Aquí solo agregue otras evidencias.</div>';
  if (mode === 'campo' && !rows.length) {
    h += '<div class="empty">No se definieron ' + (s === 's7' ? 'KPI' : 'parámetros') + ' antes de la visita.<br>Puede agregarlos aquí.</div>';
  }
  rows.forEach(function (r, i) {
    var known = mode === 'campo' && COLS[s].filter(function (c) { return c.plan; }).every(function (c) { return !blank(r[c.k]); });
    h += '<div class="rowcard"><div class="rh"><span>' + ROWNAME[s] + ' ' + (i + 1) + '</span><span class="ops">' +
      (s === 's3' && i > 0 ? '<button class="iconb" data-up="' + i + '" data-s="' + s + '" aria-label="Subir">↑</button>' : '') +
      '<button class="iconb" data-rowna="' + i + '" data-s="' + s + '" title="Vacíos → No aplica">NA</button>' +
      '<button class="iconb" data-delrow="' + i + '" data-s="' + s + '" aria-label="Eliminar">🗑</button></span></div><div class="rb">';
    if (known) h += '<div class="given">' + COLS[s].filter(function (c) { return c.plan; }).map(function (c) { return '<b>' + esc(c.label.replace(/ \(.*\)$/, '')) + ':</b> ' + esc(r[c.k]); }).join('<br>') + '</div>';
    COLS[s].forEach(function (c) {
      if (mode === 'plan' && !c.plan) return;
      if (known && c.plan) return;
      var path = s + '.rows.' + i + '.' + c.k;
      h += '<div class="f"><label>' + esc(c.label) + '</label>' + (c.choice ? seg(path, [['Sí', 'Sí', 'si'], ['Parcial', 'Parcial', 'par'], ['No', 'No', 'no']], r[c.k]) : inputHtml(path, r[c.k], { type: c.type, big: c.big && mode === 'campo', mode: c.mode })) + '</div>';
    });
    h += '</div></div>';
  });
  var addLabel = { s3: '+ Agregar etapa', s4: '+ Agregar recurso o insumo', s5: mode === 'campo' ? '+ Agregar parámetro no previsto' : '+ Agregar parámetro', s7: mode === 'campo' ? '+ Agregar resultado no previsto' : '+ Agregar KPI o resultado', s13: '+ Agregar anexo' }[s];
  h += '<button class="btn block sky" data-addrow="' + s + '">' + addLabel + '</button>';
  if (mode !== 'campo') h += '<label class="nacheck" style="margin-top:12px"><input type="checkbox" data-na="' + s + '"> Esta sección no aplica</label>';
  if (s === 's4') h += '<p class="note">Complete el porcentaje solo cuando corresponda y con una misma base de unidades. <button class="chip" data-act="pct">Calcular % automáticamente</button></p>';
  return h;
}

function s8View(p) {
  var sd = p.data.s8 = p.data.s8 || {}, fx = sd.fixed = sd.fixed || {}, h = '';
  FIXED8.forEach(function (a) {
    var x = fx[a.k] || {};
    h += '<div class="rowcard"><div class="rh"><span>' + esc(a.label) + '</span><span class="ops"><button class="iconb" data-fill8="' + a.k + '">NA</button></span></div><div class="rb">' +
      '<div class="f"><label>Situación observada</label>' + inputHtml('s8.fixed.' + a.k + '.sit', x.sit, { type: 'area', rows: 2 }) + '</div>' +
      '<div class="f"><label>Acción o mejora propuesta</label>' + inputHtml('s8.fixed.' + a.k + '.accion', x.accion, { type: 'area', rows: 2 }) + '</div>' +
      '<div class="f"><label>Responsable o estado</label>' + inputHtml('s8.fixed.' + a.k + '.resp', x.resp) + '</div></div></div>';
  });
  (sd.otros || []).forEach(function (r, i) {
    h += '<div class="rowcard"><div class="rh"><span>Otro aspecto ' + (i + 1) + '</span><span class="ops"><button class="iconb" data-delotro="' + i + '">🗑</button></span></div><div class="rb">' +
      '<div class="f"><label>Aspecto</label>' + inputHtml('s8.otros.' + i + '.aspecto', r.aspecto) + '</div>' +
      '<div class="f"><label>Situación observada</label>' + inputHtml('s8.otros.' + i + '.sit', r.sit, { type: 'area', rows: 2 }) + '</div>' +
      '<div class="f"><label>Acción o mejora propuesta</label>' + inputHtml('s8.otros.' + i + '.accion', r.accion, { type: 'area', rows: 2 }) + '</div>' +
      '<div class="f"><label>Responsable o estado</label>' + inputHtml('s8.otros.' + i + '.resp', r.resp) + '</div></div></div>';
  });
  return h + '<button class="btn block sky" data-act="addotro">+ Agregar otro aspecto</button>';
}
function s9View(p) {
  var sd = p.data.s9 = p.data.s9 || {}, h = '';
  h += '<div class="gate"><b>¿La prueba es favorable y se autoriza avanzar hacia la implementación?</b>' +
    '<div class="seg"><button class="' + (p.gate === 'favorable' ? 'sel si' : '') + '" data-gate="favorable">Sí → se llena la secc. 9</button><button class="' + (p.gate === 'no' ? 'sel no' : '') + '" data-gate="no">No → secc. 9 "No aplica"</button></div></div>';
  CRIT9.forEach(function (q, i) { h += '<div class="crit"><div class="q">' + esc(q) + '</div>' + seg('s9.crit.' + i, CRIT_OPTS, (sd.crit || [])[i]) + '</div>'; });
  h += '<div class="f"><label>Comparación entre la línea base y los resultados obtenidos</label><div class="prompt">Descríbala o indique el anexo con el gráfico.</div>' + inputHtml('s9.grafico', sd.grafico, { type: 'area' }) + chips('s9.grafico', sd.grafico) + '</div>';
  h += '<div class="f"><label>Conclusión y recomendación</label>' + inputHtml('s9.conclusion', sd.conclusion, { type: 'area', rows: 4 }) + chips('s9.conclusion', sd.conclusion) + '</div>';
  return h;
}
function s12View(p) {
  var sd = p.data.s12 = p.data.s12 || {}, h = '';
  ROLES.forEach(function (r) {
    var x = sd[r[0]] || {};
    h += '<div class="rowcard"><div class="rh"><span>' + r[1] + '</span></div><div class="rb">' +
      '<div class="f"><label>Nombre</label>' + inputHtml('s12.' + r[0] + '.nombre', x.nombre) + chips('s12.' + r[0] + '.nombre', x.nombre) + '</div>' +
      '<div class="f"><label>Fecha</label>' + inputHtml('s12.' + r[0] + '.fecha', x.fecha, { type: 'date' }) + '</div></div></div>';
  });
  return h;
}
function fotosView(p) {
  var fs = p.fotos || [], h = '<div class="row" style="margin-bottom:12px">' +
    '<label class="btn primary filebtn">📷 Tomar foto<input type="file" accept="image/*" capture="environment" data-addfoto></label>' +
    '<label class="btn sky filebtn">🖼 De la galería<input type="file" accept="image/*" multiple data-addfoto></label></div>';
  if (!fs.length) return h + '<div class="empty">Aún no hay fotos.</div>';
  h += '<div class="photos">' + fs.map(function (f, i) {
    return '<div class="ph"><img data-fid="' + f.id + '" alt="Foto ' + (i + 1) + '"><span class="n">' + (i + 1) + '</span><button class="del" data-delfoto="' + f.id + '" aria-label="Eliminar">🗑</button>' +
      '<textarea data-cap="' + f.id + '" placeholder="Descripción (qué muestra la foto)">' + esc(f.caption || '') + '</textarea></div>';
  }).join('') + '</div>';
  return h;
}

function renderView(p, ph, i) {
  var P = PHASES[ph], vw = P && P.views[i]; if (!vw) return renderPhase(p, ph);
  var s = vw.s, d = p.data, h = top(vw.t, get(d, 's1.nombre') || '', 'v=phase&p=' + p.id + '&ph=' + ph) + '<main class="wrap">';
  if (vw.note) h += '<p class="note">' + esc(vw.note) + '</p>';
  if (s === 's1') { var sd1 = d.s1 = d.s1 || {}; F1.forEach(function (f) { h += fieldBlock('s1', f, sd1); }); }
  else if (s === 's2') { var sd2 = d.s2 = d.s2 || {}; F2.forEach(function (f) { h += fieldBlock('s2', f, sd2); }); }
  else if (s === 's10') {
    if (p.gate === 'no') h += '<div class="banner info">En la evaluación (secc. 8) se indicó que la prueba <b>no es favorable</b>. Esta sección aparecerá como "No aplica" en el Excel.</div>';
    else { if (!p.gate) h += '<div class="banner warn">Primero responda en la secc. 8 si la prueba es favorable.</div>'; var sd10 = d.s10 = d.s10 || {}; F10.forEach(function (f) { h += fieldBlock('s10', f, sd10); }); }
  }
  else if (s === 's11') { var sd11 = d.s11 = d.s11 || {}; F11.forEach(function (f) { h += fieldBlock('s11', f, sd11); }); }
  else if (COLS[s]) h += tableView(p, s, vw.m);
  else if (s === 's8') h += s8View(p);
  else if (s === 's9') h += s9View(p);
  else if (s === 's12') h += s12View(p);
  else if (s === 'fotos') h += fotosView(p);
  h += '<div class="saved" id="savedFlag"></div>';
  var prev = i > 0 ? 'v=view&p=' + p.id + '&ph=' + ph + '&i=' + (i - 1) : null;
  var next = i < P.views.length - 1 ? 'v=view&p=' + p.id + '&ph=' + ph + '&i=' + (i + 1) : 'v=phase&p=' + p.id + '&ph=' + ph;
  h += '<div class="dock"><div class="in-dock">' + (prev ? '<button class="btn" data-go="' + prev + '" data-rep="1">← Anterior</button>' : '') +
    '<button class="btn primary" data-go="' + next + '" data-rep="1">' + (i < P.views.length - 1 ? 'Siguiente →' : 'Listo ✓') + '</button></div></div>';
  h += '</main>';
  app.innerHTML = h;
  if (s === 'fotos') app.querySelectorAll('img[data-fid]').forEach(function (img) { photoUrl(img.dataset.fid).then(function (u) { if (u) img.src = u; }); });
}

function renderReview(p) {
  var a = analyze(p), h = top('Revisar y generar Excel', get(p.data, 's1.nombre') || '', 'v=pilot&p=' + p.id) + '<main class="wrap">';
  if (!a.miss.length) h += '<div class="banner ok"><b>Todo completo.</b> El registro cumple todos los campos del formato.</div>';
  else {
    h += '<div class="banner warn">Faltan <b>' + a.miss.length + '</b> dato(s) del formato. Puede generar el Excel igual: los vacíos quedan en blanco.</div>';
    PH_ORDER.forEach(function (ph) {
      var items = a.miss.filter(function (m) { return m.ph === ph; });
      if (!items.length) return;
      h += '<div class="sub">' + PHASES[ph].name + ' (' + items.length + ')</div><div class="card" style="padding:4px 12px">';
      items.slice(0, 40).forEach(function (m) { h += '<button class="miss" data-go="v=view&p=' + p.id + '&ph=' + ph + '&i=' + m.i + '">' + esc(m.title.split(' · ')[0]) + ' · ' + esc(m.t) + '<span>›</span></button>'; });
      if (items.length > 40) h += '<div class="muted" style="padding:8px 2px">… y ' + (items.length - 40) + ' más</div>';
      h += '</div>';
    });
    h += '<div class="card"><h3>Completar los vacíos de una vez</h3><p class="muted" style="margin:0 0 10px">Úselo solo para lo que de verdad no aplica o no se registró. No marca decisiones ni cumplimiento de KPI.</p><div class="row">' +
      STD.map(function (s) { return '<button class="btn sm" data-fillall="' + s + '">' + s + '</button>'; }).join('') + '</div></div>';
  }
  var nf = (p.fotos || []).length;
  h += '<div class="card"><label style="display:flex;gap:10px;align-items:center;font-size:15px"><input type="checkbox" id="incFotos" style="width:22px;height:22px"' + (nf ? ' checked' : ' disabled') + '> Incluir hoja "Anexo fotográfico" (' + nf + ' foto' + (nf === 1 ? '' : 's') + ')</label></div>';
  h += '<div class="dock"><div class="in-dock"><button class="btn primary" data-act="excel" id="btnExcel">📊 Generar Excel</button></div></div>';
  h += '</main>';
  app.innerHTML = h;
}

/* ---------- Modales ---------- */
function modal(html, onReady) {
  var bg = document.createElement('div'); bg.className = 'modal-bg';
  bg.innerHTML = '<div class="modal">' + html + '</div>';
  document.body.appendChild(bg);
  var close = function () { bg.remove(); };
  bg.addEventListener('click', function (e) { if (e.target === bg || e.target.closest('[data-close]')) close(); });
  if (onReady) onReady(bg, close);
  return close;
}
function confirmBox(title, text, okLabel, fn, danger) {
  modal('<h3>' + esc(title) + '</h3><p>' + esc(text) + '</p><div class="row"><button class="btn" data-close>Cancelar</button><button class="btn ' + (danger ? 'danger' : 'primary') + '" data-ok>' + esc(okLabel) + '</button></div>', function (bg, close) {
    bg.querySelector('[data-ok]').onclick = function () { close(); fn(); };
  });
}
function newPilot() {
  modal('<h3>Nuevo piloto</h3><p>Podrá completar el resto por etapas.</p>' +
    '<div class="f"><label>Nombre de la iniciativa</label><input class="in" id="nNom" autocomplete="off"></div>' +
    '<div class="f"><label>Código o referencia</label><div class="prompt">Código de Wave o número de iniciativa.</div><input class="in" id="nCod" autocomplete="off"></div>' +
    '<div class="row"><button class="btn" data-close>Cancelar</button><button class="btn primary" data-ok>Crear</button></div>', function (bg, close) {
    var n = bg.querySelector('#nNom'); setTimeout(function () { n.focus(); }, 50);
    bg.querySelector('[data-ok]').onclick = function () {
      var nom = n.value.trim(), cod = bg.querySelector('#nCod').value.trim();
      if (!nom) { toast('Escriba el nombre de la iniciativa.'); return; }
      var dup = cod && Object.keys(PIL).some(function (k) { return (get(PIL[k].data, 's1.codigo') || '').toLowerCase() === cod.toLowerCase(); });
      if (dup) { toast('Ya hay un piloto con ese código en este celular.'); return; }
      var p = { id: uid(), createdAt: new Date().toISOString(), data: { __v: 2, s1: { nombre: nom, codigo: cod } }, fotos: [], gate: '' };
      PIL[p.id] = p; save(p, true); close();
      go({ v: 'pilot', p: p.id });
    };
  });
}
function pilotMenu(p) {
  modal('<h3>' + esc(get(p.data, 's1.nombre') || 'Piloto') + '</h3><p>Opciones del piloto</p>' +
    '<button class="btn block sky" data-m="share" style="margin-bottom:10px">📤 Compartir piloto (para que otra persona lo continúe)</button>' +
    '<button class="btn block danger" data-m="del" style="margin-bottom:10px">🗑 Eliminar de este celular</button>' +
    '<button class="btn block" data-close>Cerrar</button>', function (bg, close) {
    bg.querySelector('[data-m="share"]').onclick = function () { close(); exportPilot(p); };
    bg.querySelector('[data-m="del"]').onclick = function () {
      close();
      confirmBox('Eliminar piloto', 'Se borrará "' + (get(p.data, 's1.nombre') || '') + '" y sus fotos de este celular. Si no lo ha compartido, se pierde.', 'Eliminar', function () {
        (p.fotos || []).forEach(function (f) { delPhoto(f.id); });
        delete PIL[p.id];
        tx('pilotos', 'readwrite', function (st) { st.delete(p.id); });
        go({ v: 'home' });
      }, true);
    };
  });
}

/* ---------- Archivos: compartir o descargar ---------- */
function deliver(blob, name, mime, title) {
  var file = null;
  try { file = new File([blob], name, { type: mime }); } catch (e) {}
  var canShare = file && navigator.canShare && navigator.canShare({ files: [file] });
  modal('<h3>' + esc(title) + '</h3><p>' + esc(name) + '</p>' +
    (canShare ? '<button class="btn block primary" data-share style="margin-bottom:10px">📤 Compartir (WhatsApp, correo, OneDrive…)</button>' : '') +
    '<button class="btn block ' + (canShare ? 'sky' : 'primary') + '" data-dl style="margin-bottom:10px">⬇️ Guardar en el celular</button>' +
    '<button class="btn block" data-close>Cerrar</button>', function (bg, close) {
    var sh = bg.querySelector('[data-share]');
    if (sh) sh.onclick = function () {
      navigator.share({ files: [file], title: name }).then(close, function (e) { if (!e || e.name !== 'AbortError') toast('No se pudo compartir. Use "Guardar en el celular".'); });
    };
    bg.querySelector('[data-dl]').onclick = function () {
      var u = URL.createObjectURL(blob), a = document.createElement('a');
      a.href = u; a.download = name; document.body.appendChild(a); a.click();
      setTimeout(function () { URL.revokeObjectURL(u); a.remove(); }, 4000);
      toast('Archivo guardado en Descargas.');
    };
  });
}
function blobToDataUrl(b) { return new Promise(function (res) { var r = new FileReader(); r.onload = function () { res(r.result); }; r.onerror = function () { res(''); }; r.readAsDataURL(b); }); }
function dataUrlToBlob(u) {
  var m = /^data:([^;]+);base64,(.*)$/.exec(u || ''); if (!m) return null;
  var bin = atob(m[2]), arr = new Uint8Array(bin.length);
  for (var i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i);
  return new Blob([arr], { type: m[1] });
}
function exportPilot(p) {
  toast('Preparando el archivo…');
  Promise.all((p.fotos || []).map(function (f) { return getPhoto(f.id).then(function (b) { return b ? blobToDataUrl(b) : ''; }); })).then(function (urls) {
    var out = { formato: 'cor-gppt-r006-piloto', version: 1, exportado: new Date().toISOString(), piloto: clone(p) };
    out.piloto.fotos = (p.fotos || []).map(function (f, i) { var o = clone(f); o.dataUrl = urls[i]; return o; });
    var blob = new Blob([JSON.stringify(out)], { type: 'application/json' });
    var name = 'Piloto_' + (safeName(get(p.data, 's1.codigo')) || safeName(get(p.data, 's1.nombre')) || 'sin_codigo') + '.json';
    deliver(blob, name, 'application/json', 'Compartir piloto');
  });
}
function importPilot(file) {
  var r = new FileReader();
  r.onload = function () {
    var o; try { o = JSON.parse(r.result); } catch (e) { toast('El archivo no es un piloto válido.'); return; }
    if (!o || o.formato !== 'cor-gppt-r006-piloto' || !o.piloto || !o.piloto.data) { toast('El archivo no es un piloto de esta app.'); return; }
    var p = o.piloto, exists = PIL[p.id];
    var doIt = function () {
      var olds = exists ? (exists.fotos || []) : [];
      olds.forEach(function (f) { delPhoto(f.id); });
      var fotos = (p.fotos || []).map(function (f) { var b = dataUrlToBlob(f.dataUrl); var m = clone(f); delete m.dataUrl; return { meta: m, blob: b }; }).filter(function (x) { return x.blob; });
      p.fotos = fotos.map(function (x) { return x.meta; });
      Promise.all(fotos.map(function (x) { return putPhoto(x.meta.id, p.id, x.blob); })).then(function () {
        PIL[p.id] = p; return save(p, true);
      }).then(function () { toast('Piloto abierto: ' + (get(p.data, 's1.nombre') || '')); go({ v: 'pilot', p: p.id }); }, function () { toast('No se pudo guardar el piloto.'); });
    };
    if (exists) confirmBox('Reemplazar piloto', 'Ya tiene "' + (get(exists.data, 's1.nombre') || '') + '" en este celular. ¿Reemplazarlo con el archivo recibido (modificado ' + fmtTs(p.updatedAt) + ')?', 'Reemplazar', doIt);
    else doIt();
  };
  r.readAsText(file);
}

/* ---------- Fotos ---------- */
function compress(file) {
  return new Promise(function (res) {
    var url = URL.createObjectURL(file), img = new Image();
    img.onload = function () {
      var max = 1600, w = img.naturalWidth, h = img.naturalHeight, k = Math.min(1, max / Math.max(w, h));
      var c = document.createElement('canvas'); c.width = Math.round(w * k); c.height = Math.round(h * k);
      c.getContext('2d').drawImage(img, 0, 0, c.width, c.height);
      c.toBlob(function (b) { URL.revokeObjectURL(url); res(b ? { blob: b, w: c.width, h: c.height } : null); }, 'image/jpeg', 0.82);
    };
    img.onerror = function () { URL.revokeObjectURL(url); res(null); };
    img.src = url;
  });
}
function addFotos(p, files) {
  files = Array.prototype.slice.call(files || []);
  if (!files.length) return;
  toast('Guardando ' + files.length + ' foto(s)…');
  var chain = Promise.resolve(), n = 0;
  files.forEach(function (f) {
    chain = chain.then(function () { return compress(f); }).then(function (c) {
      if (!c) return;
      var id = 'f' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
      return putPhoto(id, p.id, c.blob).then(function () { p.fotos = p.fotos || []; p.fotos.push({ id: id, w: c.w, h: c.h, caption: '', at: new Date().toISOString() }); n++; });
    });
  });
  chain.then(function () { save(p, true); render(); toast(n + ' foto(s) guardada(s). Agregue una descripción.'); }, function () { toast('No se pudieron guardar las fotos.'); render(); });
}

/* ---------- Excel ---------- */
function makeExcel(p) {
  var btn = document.getElementById('btnExcel');
  if (btn) { btn.disabled = true; btn.innerHTML = '<span class="spin"></span> Generando…'; }
  var inc = document.getElementById('incFotos'), withPhotos = inc && inc.checked;
  var rec = clone(p.data), fotos = p.fotos || [];
  if (p.gate === 'no') rec.s10 = { decision: 'No aplica', cronograma: 'No aplica', kpiDef: 'No aplica', frec: 'No aplica', resp: 'No aplica' };
  if (fotos.length) {
    rec.s13 = rec.s13 || {};
    var base = rec.s13.na ? [] : (rec.s13.rows || []);
    rec.s13 = { na: false, rows: base.concat(fotos.map(function (f, i) {
      return { anexo: 'Fotografía ' + (i + 1), desc: (f.caption || 'Fotografía de campo') + ' (' + fmtTs(f.at) + ')' + (withPhotos ? ' — ver hoja "Anexo fotográfico"' : '') };
    })) };
  }
  var photosP = withPhotos ? Promise.all(fotos.map(function (f) {
    return getPhoto(f.id).then(function (b) { return b ? b.arrayBuffer() : null; }).then(function (ab) { return ab ? { bytes: new Uint8Array(ab), w: f.w, h: f.h, caption: f.caption } : null; });
  })).then(function (a) { return a.filter(Boolean); }) : Promise.resolve([]);
  photosP.then(function (photos) {
    return COR.buildXlsx(JSZip, rec, self.COR_XLSX_B64, false, { photos: photos });
  }).then(function (blob) {
    var name = 'COR-GPPT-R006_' + (safeName(get(p.data, 's1.codigo')) ? safeName(get(p.data, 's1.codigo')) + '_' : '') + (safeName(get(p.data, 's1.nombre')) || 'piloto') + '.xlsx';
    deliver(blob, name, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'Excel listo');
  }, function () { toast('No se pudo generar el Excel.'); }).then(function () {
    if (btn) { btn.disabled = false; btn.innerHTML = '📊 Generar Excel'; }
  });
}
function calcPct(p) {
  var rows = get(p.data, 's4.rows') || [];
  var num = function (x) { return parseFloat(String(x).replace(',', '.')); };
  var used = rows.filter(function (r) { return !isNaN(num(r.cantidad)); });
  if (!used.length) { toast('No hay cantidades numéricas.'); return; }
  var units = {}; used.forEach(function (r) { units[(r.unidad || '').trim().toLowerCase()] = 1; });
  if (Object.keys(units).length > 1) { toast('Las filas usan unidades distintas: el formato pide una misma base.'); return; }
  var tot = used.reduce(function (a, r) { return a + num(r.cantidad); }, 0); if (!tot) return;
  used.forEach(function (r) { r.porcentaje = (num(r.cantidad) / tot * 100).toFixed(2) + ' %'; });
  save(p); render();
}

/* ---------- Eventos ---------- */
function cur() { return R.p && PIL[R.p]; }
app.addEventListener('input', function (e) {
  var t = e.target, p = cur(); if (!p) return;
  if (t.dataset.p) {
    set(p.data, t.dataset.p, t.value);
    var wrap = t.closest('.f'); if (wrap) wrap.querySelectorAll('.chip.sel').forEach(function (c) { c.classList.remove('sel'); });
    save(p);
  } else if (t.dataset.cap) {
    var f = (p.fotos || []).find(function (x) { return x.id === t.dataset.cap; }); if (f) { f.caption = t.value; save(p); }
  }
});
app.addEventListener('change', function (e) {
  var t = e.target, p = cur();
  if (t.hasAttribute('data-import')) { if (t.files[0]) importPilot(t.files[0]); t.value = ''; return; }
  if (!p) return;
  if (t.hasAttribute('data-addfoto')) { addFotos(p, t.files); t.value = ''; return; }
  if (t.dataset.na) { var sd = p.data[t.dataset.na] = p.data[t.dataset.na] || {}; sd.na = t.checked; save(p); render(); }
});
app.addEventListener('click', function (e) {
  var b = e.target.closest('button, [data-go]'); if (!b) return;
  var ds = b.dataset, p = cur();
  if (ds.go != null) { var o = parseHashStr(ds.go); go(o, ds.rep === '1' && R.v === 'view' && o.v === 'view'); return; }
  if (ds.act === 'new') return newPilot();
  if (ds.act === 'install' && installEvt) { installEvt.prompt(); installEvt.userChoice.then(function () { installEvt = null; render(); }); return; }
  if (!p) return;
  if (ds.act === 'menu') return pilotMenu(p);
  if (ds.act === 'excel') return makeExcel(p);
  if (ds.act === 'pct') return calcPct(p);
  if (ds.setp) {
    var curV = get(p.data, ds.setp);
    set(p.data, ds.setp, curV === ds.val ? '' : ds.val); save(p); keepScroll(render); return;
  }
  if (ds.gate) {
    p.gate = p.gate === ds.gate ? '' : ds.gate;
    if (p.gate === 'no') p.data.s10 = { decision: 'No aplica', cronograma: 'No aplica', kpiDef: 'No aplica', frec: 'No aplica', resp: 'No aplica' };
    else if (get(p.data, 's10.decision') === 'No aplica') p.data.s10 = {};
    save(p); keepScroll(render); return;
  }
  if (ds.addrow) {
    var sd = p.data[ds.addrow] = p.data[ds.addrow] || {}; sd.na = false; sd.rows = sd.rows || []; sd.rows.push({}); save(p);
    keepScroll(render);
    var cards = app.querySelectorAll('.rowcard'); var last = cards[cards.length - 1];
    if (last) { last.scrollIntoView({ block: 'center' }); var inp = last.querySelector('.in'); if (inp) inp.focus(); }
    return;
  }
  if (ds.delrow != null) {
    var s = ds.s, i = +ds.delrow;
    confirmBox('Eliminar fila', 'Se eliminará ' + ROWNAME[s].toLowerCase() + ' ' + (i + 1) + '.', 'Eliminar', function () { p.data[s].rows.splice(i, 1); save(p); keepScroll(render); }, true);
    return;
  }
  if (ds.rowna != null) {
    var r = p.data[ds.s].rows[+ds.rowna], vw = PHASES[R.ph].views[R.i];
    COLS[ds.s].forEach(function (c) { if (c.optional || c.choice) return; if (vw.m === 'plan' && !c.plan) return; if (blank(r[c.k])) r[c.k] = 'No aplica'; });
    save(p); keepScroll(render); return;
  }
  if (ds.up != null) { var rows = p.data[ds.s].rows, k = +ds.up, tmp = rows[k - 1]; rows[k - 1] = rows[k]; rows[k] = tmp; save(p); keepScroll(render); return; }
  if (ds.fill8) { var fx = p.data.s8.fixed[ds.fill8] = p.data.s8.fixed[ds.fill8] || {}; ['sit', 'accion', 'resp'].forEach(function (c) { if (blank(fx[c])) fx[c] = 'No aplica'; }); save(p); keepScroll(render); return; }
  if (ds.act === 'addotro') { var s8 = p.data.s8 = p.data.s8 || {}; s8.otros = s8.otros || []; s8.otros.push({}); save(p); keepScroll(render); return; }
  if (ds.delotro != null) { p.data.s8.otros.splice(+ds.delotro, 1); save(p); keepScroll(render); return; }
  if (ds.delfoto) {
    confirmBox('Eliminar foto', 'La foto se borrará de este celular.', 'Eliminar', function () {
      p.fotos = (p.fotos || []).filter(function (f) { return f.id !== ds.delfoto; }); delPhoto(ds.delfoto); save(p, true); keepScroll(render);
    }, true);
    return;
  }
  if (ds.fillall) {
    confirmBox('Completar vacíos', 'Todos los campos vacíos quedarán como "' + ds.fillall + '". Las decisiones y el cumplimiento de KPI no se tocan.', 'Completar', function () {
      var n = fillBlanks(p, ds.fillall); save(p, true); render(); toast(n + ' campo(s) completado(s).');
    });
  }
});
function parseHashStr(s) { var o = {}; s.split('&').forEach(function (kv) { var q = kv.split('='); if (q[0]) o[q[0]] = decodeURIComponent(q[1] || ''); }); return o; }
function keepScroll(fn) { var y = window.scrollY; fn(); window.scrollTo(0, y); }

/* ---------- Inicio ---------- */
if ('serviceWorker' in navigator && /^https:|^http:\/\/(localhost|127\.)/.test(location.href)) {
  navigator.serviceWorker.register('sw.js').catch(function () {});
}
if (navigator.storage && navigator.storage.persist) navigator.storage.persist().catch(function () {});
openDb().then(function (db) {
  DB = db;
  return DB ? loadAll() : null;
}).then(route, route);
})();
