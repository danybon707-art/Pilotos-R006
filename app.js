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
  s7: [{ k: 'kpi', label: 'KPI o resultado', plan: true }, { k: 'evalua', label: 'Qué evalúa', plan: true }, { k: 'base', label: 'Línea base: valor actual (opcional, para el gráfico)', plan: true, optional: true }, { k: 'meta', label: 'Meta o criterio (con unidad)', plan: true }, { k: 'resultado', label: 'Resultado obtenido (con unidad)', big: true }, { k: 'cumple', label: 'Cumplimiento', choice: ['Sí', 'Parcial', 'No'] }],
  s13: [{ k: 'anexo', label: 'Anexo o evidencia' }, { k: 'desc', label: 'Descripción o referencia', type: 'area' }]
};
var ROWNAME = { s3: 'Etapa', s4: 'Recurso', s5: 'Parámetro', s7: 'KPI / resultado', s13: 'Anexo' };

var PHASES = {
  antes: { name: 'Antes de la visita', icon: '📋', desc: 'Datos generales, metodología y qué se va a medir', views: [
    { s: 's1', t: '1 · Datos generales' },
    { s: 's2', t: '2 · Metodología de evaluación', note: 'Defina cómo se comprobará el desempeño de la prueba y cuál será el punto de comparación.' },
    { s: 's7', m: 'plan', t: '5 · KPI y resultados: qué se va a medir', note: 'Liste los KPI o resultados con su meta. En campo solo se anotará el resultado obtenido.' }
  ]},
  campo: { name: 'En campo', icon: '🏭', desc: 'Procedimiento, insumos, mediciones, resultados, dificultades y fotos', views: [
    { s: 's3', t: '3 · Procedimiento realizado', note: 'Etapas principales en el orden en que se ejecutaron.' },
    { s: 's4', t: '4 · Recursos e insumos', note: 'Materias primas, materiales, reactivos, equipos, servicios. Si hay formulación, inclúyala aquí.' },
    { s: 's7', m: 'campo', t: '5 · KPI y resultados obtenidos', note: 'Anote el resultado de cada KPI (con unidad) y marque si cumple la meta.' },
    { s: 's8', t: '6 · Riesgos, dificultades y mejoras' },
    { s: 'fotos', t: 'Fotos de la visita', note: 'Displays, hojas de laboratorio, bitácora, equipo. Se agregan solas como anexos y en una hoja del Excel.' }
  ]},
  cierre: { name: 'Cierre', icon: '✅', desc: 'Evaluación, seguimiento, cierre, validación y anexos', views: [
    { s: 's9', t: '7 · Evaluación de la prueba piloto', note: 'Marque una opción por criterio y complemente la conclusión con la evidencia disponible.' },
    { s: 's10', t: '8 · Seguimiento de implementación', note: 'Solo si la prueba es favorable y se autoriza avanzar.' },
    { s: 's11', t: '9 · Cierre administrativo y contable', note: 'Referencias de cierre según los controles internos.' },
    { s: 's12', t: '10 · Validación del registro', note: 'La firma se estampa en el documento impreso.' },
    { s: 's13', t: '11 · Anexos y evidencias', note: 'Hojas de datos, resultados de laboratorio, gráficos, cálculos, cotizaciones, bitácoras. Las fotos de la app se agregan solas.' }
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
  if (R.ph === 'campo' && R.v === 'view') stampCampo(p);
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

/* ---------- Valores automáticos ----------
   eff(p) devuelve una copia del piloto con todo lo que la app puede deducir sola.
   Lo que la persona escribe o marca siempre tiene prioridad. */
var NAME_KEY = 'cor-campo-mi-nombre';
function myName() { try { return (localStorage.getItem(NAME_KEY) || '').trim(); } catch (e) { return ''; } }
function setMyName(v) { try { localStorage.setItem(NAME_KEY, v); } catch (e) {} }
function parseNum(str) {
  var m = String(str || '').match(/-?\d[\d.,]*/); if (!m) return null;
  var t = m[0].replace(/[.,]$/, '');
  if (t.indexOf(',') >= 0 && t.indexOf('.') >= 0) t = t.lastIndexOf(',') > t.lastIndexOf('.') ? t.replace(/\./g, '').replace(',', '.') : t.replace(/,/g, '');
  else if (/^-?[1-9]\d{0,2}([.,]\d{3})+$/.test(t)) t = t.replace(/[.,]/g, ''); /* 1.200 o 1,200 = mil doscientos */
  else t = t.replace(',', '.');
  var n = parseFloat(t); return isNaN(n) ? null : n;
}
/* Compara el resultado con la meta: "≤ 0,83", "mínimo 95 %", "entre 7,5 y 8", "> 20"… */
function evalKpi(meta, res) {
  if (!real(meta) || !real(res)) return '';
  var r = parseNum(res); if (r === null) return '';
  var m = String(meta).toLowerCase();
  var nums = (m.match(/-?\d[\d.,]*/g) || []).map(parseNum).filter(function (x) { return x !== null; });
  if (!nums.length) return '';
  var ok;
  if (/≤|<=|=<|m[aá]x|menor o igual|hasta|no (mayor|m[aá]s|superior)|como m[aá]ximo/.test(m)) ok = r <= nums[0];
  else if (/≥|>=|=>|m[ií]n|mayor o igual|al menos|no (menor|inferior)|como m[ií]nimo/.test(m)) ok = r >= nums[0];
  else if (/<|menor|inferior|reducir a menos/.test(m)) ok = r < nums[0];
  else if (/>|mayor|superior|m[aá]s de/.test(m)) ok = r > nums[0];
  else if (nums.length >= 2 && /entre|[–—-]|\sa\s|\sy\s/.test(m)) { var lo = Math.min(nums[0], nums[1]), hi = Math.max(nums[0], nums[1]); ok = r >= lo && r <= hi; }
  else return '';
  return ok ? 'Sí' : 'No';
}
function kpiAgg(rows) {
  var ev = rows.filter(function (r) { return EST[r.cumple]; });
  if (!ev.length) return '';
  if (ev.every(function (r) { return r.cumple === 'Sí'; })) return 'Sí';
  if (ev.every(function (r) { return r.cumple === 'No'; })) return 'No';
  return 'Parcial';
}
function samePct(rows) {
  var used = rows.filter(function (r) { return parseNum(r.cantidad) !== null; });
  if (used.length < 2) return null;
  var u = {}; used.forEach(function (r) { u[(r.unidad || '').trim().toLowerCase()] = 1; });
  if (Object.keys(u).length !== 1) return null;
  var tot = used.reduce(function (a, r) { return a + parseNum(r.cantidad); }, 0); if (!tot) return null;
  var out = {}; rows.forEach(function (r, i) { var n = parseNum(r.cantidad); if (n !== null) out[i] = (n / tot * 100).toFixed(1).replace('.', ',') + ' %'; });
  return out;
}
function eff(p) {
  var E = { id: p.id, fotos: p.fotos || [], campo: p.campo, data: clone(p.data), gate: p.gate || '' };
  var d = E.data;
  var s7 = d.s7 = d.s7 || {};
  if (!s7.na) (s7.rows || []).forEach(function (r) { if (blank(r.cumple)) r.cumple = evalKpi(r.meta, r.resultado); });
  var s4 = d.s4 = d.s4 || {};
  if (!s4.na && s4.rows) { var pc = samePct(s4.rows); if (pc) s4.rows.forEach(function (r, i) { if (blank(r.porcentaje) && pc[i]) r.porcentaje = pc[i]; }); }
  var agg = kpiAgg(s7.na ? [] : (s7.rows || []).filter(function (r) { return real(r.kpi); }));
  var CR = { 'Sí': 'Sí', 'Parcial': 'Parcialmente o requiere ajustes', 'No': 'No' }[agg];
  var s9 = d.s9 = d.s9 || {}; s9.crit = (s9.crit || []).slice();
  if (CR) [1, 2, 3].forEach(function (i) { if (blank(s9.crit[i])) s9.crit[i] = CR; });
  if (!E.gate && agg) E.gate = agg === 'No' ? 'no' : 'favorable';
  var s1 = d.s1 = d.s1 || {};
  if (E.gate === 'no') d.s10 = { decision: 'No aplica', cronograma: 'No aplica', kpiDef: 'No aplica', frec: 'No aplica', resp: 'No aplica' };
  else {
    var s10 = d.s10 = d.s10 || {};
    if (blank(s10.decision) && agg) s10.decision = agg === 'Sí' ? 'Implementar' : agg === 'Parcial' ? 'Implementar con ajustes' : 'Realizar nueva prueba';
    if (blank(s10.resp) && real(s1.responsable)) s10.resp = s1.responsable;
  }
  var s12 = d.s12 = d.s12 || {}, el = s12.elab = s12.elab || {};
  if (blank(el.nombre) && myName()) el.nombre = myName();
  if (blank(el.fecha)) el.fecha = today();
  var cp = p.campo || {};
  if (blank(s1.inicio) && cp.ini) s1.inicio = cp.ini;
  if (blank(s1.fin) && cp.fin) s1.fin = cp.fin;
  return E;
}
/* Datos para los gráficos: un gráfico por KPI con línea base, meta y resultado */
function unitOf(t) {
  return String(t || '').replace(/\([^)]*\)/g, ' ').replace(/[≤≥<>=±]/g, ' ').replace(/-?\d[\d.,]*/g, ' ')
    .replace(/\b(m[aá]x(imo)?|m[ií]n(imo)?|entre|hasta|al menos|como|mayor|menor|o igual|igual|superior|inferior|no|que|de|a|y|aprox\.?)\b/gi, ' ')
    .replace(/[–—\-]/g, ' ').replace(/\s+/g, ' ').trim();
}
var CUMCOL = { 'Sí': '1F8A5B', 'Parcial': 'C98A10', 'No': 'B3362C' };
function chartData(E) {
  var out = [];
  kpiRows(E).forEach(function (r) {
    var res = parseNum(r.resultado); if (res === null) return;
    var cats = [], vals = [], cols = [];
    var b = real(r.base) ? parseNum(r.base) : null;
    if (b !== null) { cats.push('Línea base'); vals.push(b); cols.push('9AA9B8'); }
    if (real(r.meta)) {
      var m = String(r.meta).toLowerCase(), nums = (m.match(/-?\d[\d.,]*/g) || []).map(parseNum).filter(function (x) { return x !== null; });
      var hasOp = /≤|≥|<|>|m[aá]x|m[ií]n|menor|mayor|hasta|al menos|superior|inferior/.test(m);
      if (nums.length >= 2 && !hasOp && /entre|[–—-]|\sa\s|\sy\s/.test(m)) { cats.push('Meta mín.', 'Meta máx.'); vals.push(Math.min(nums[0], nums[1]), Math.max(nums[0], nums[1])); cols.push('1F4E78', '1F4E78'); }
      else if (nums.length) { cats.push('Meta'); vals.push(nums[0]); cols.push('1F4E78'); }
    }
    cats.push('Resultado'); vals.push(res); cols.push(CUMCOL[r.cumple] || '4A90D9');
    if (vals.length < 2) return;
    var unit = unitOf(r.resultado) || unitOf(r.meta) || unitOf(r.base);
    out.push({ title: r.kpi.trim() + (unit ? ' (' + unit + ')' : ''), unit: unit, cats: cats, vals: vals, colors: cols, cumple: r.cumple });
  });
  return out;
}
function chartSvg(ch) {
  var W = 300, H = 120, pad = 18, max = Math.max.apply(null, ch.vals.concat([0])) || 1, n = ch.vals.length, bw = (W - 20) / n * 0.62;
  var bars = ch.vals.map(function (v, i) {
    var x = 10 + (W - 20) / n * i + ((W - 20) / n - bw) / 2, h = Math.max(1, (H - pad * 2) * v / max), y = H - pad - h;
    return '<rect x="' + x.toFixed(1) + '" y="' + y.toFixed(1) + '" width="' + bw.toFixed(1) + '" height="' + h.toFixed(1) + '" rx="3" fill="#' + ch.colors[i] + '"/>' +
      '<text x="' + (x + bw / 2).toFixed(1) + '" y="' + (y - 4).toFixed(1) + '" font-size="10" text-anchor="middle" fill="#1a2735" font-weight="700">' + String(v).replace('.', ',') + '</text>' +
      '<text x="' + (x + bw / 2).toFixed(1) + '" y="' + (H - 4) + '" font-size="9.5" text-anchor="middle" fill="#5a6b7d">' + esc(ch.cats[i]) + '</text>';
  }).join('');
  return '<div class="chartbox"><b>' + esc(ch.title) + '</b><svg viewBox="0 0 ' + W + ' ' + H + '" width="100%" role="img" aria-label="' + esc(ch.title) + '"><line x1="6" x2="' + (W - 6) + '" y1="' + (H - pad) + '" y2="' + (H - pad) + '" stroke="#d6e0eb"/>' + bars + '</svg></div>';
}
function stampCampo(p) {
  var t = today(), c = p.campo = p.campo || {};
  if (!c.ini || t < c.ini) c.ini = t;
  if (!c.fin || t > c.fin) c.fin = t;
}

/* ---------- Revisión de faltantes ---------- */
function analyze(p) {
  var E = eff(p), d = E.data, out = { total: 0, miss: [], byView: {} };
  PH_ORDER.forEach(function (ph) {
    PHASES[ph].views.forEach(function (vw, i) {
      var key = ph + ':' + i, m = [], tot = 0;
      var chk = function (ok, t) { tot++; if (!ok) m.push(t); };
      viewChecks(E, d, vw, chk);
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
  if (s === 'med') { viewChecks(p, d, { s: 's5', m: vw.m }, chk); viewChecks(p, d, { s: 's7', m: vw.m }, chk); return; }
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
    return;
  }
  if (s === 's12') {
    ROLES.forEach(function (r) { chk(!blank(get(sd, r[0] + '.nombre')), r[1] + ': nombre'); chk(!blank(get(sd, r[0] + '.fecha')), r[1] + ': fecha'); });
  }
}
/* Completa todos los vacíos (salvo decisiones que la persona debe tomar) con un valor estándar */
function fillBlanks(p, val) {
  /* Solo completa lo que la app tampoco puede deducir sola (ver eff) */
  var d = p.data, E = eff(p), ed = E.data, n = 0;
  var fill = function (o, k, eo) { if (blank(o[k]) && (!eo || blank(eo[k]))) { o[k] = val; n++; } };
  [['s1', F1], ['s2', F2], ['s11', F11]].concat(E.gate === 'no' ? [] : [['s10', F10]]).forEach(function (x) {
    var sd = d[x[0]] = d[x[0]] || {}, eo = ed[x[0]] || {};
    x[1].forEach(function (f) {
      if (f.type === 'range') { fill(sd, 'inicio', eo); fill(sd, 'fin', eo); }
      else if (f.type === 'choice') { if (blank(sd[f.k]) && blank(eo[f.k]) && (val === 'No aplica' || f.options.indexOf(val) >= 0)) { sd[f.k] = val; n++; } }
      else fill(sd, f.k, eo);
    });
  });
  Object.keys(COLS).forEach(function (s) {
    if (s === 's5') return; /* la secc. 5 no se captura en la app */
    var sd = d[s] = d[s] || {};
    if (s === 's13' && (p.fotos || []).length && !(sd.rows || []).length) return;
    if (!sd.na && !(sd.rows || []).length) { if (val === 'No aplica') { sd.na = true; n++; } return; }
    if (sd.na) return;
    sd.rows.forEach(function (r) { COLS[s].forEach(function (c) { if (!c.optional && !c.choice) fill(r, c.k); }); });
  });
  var s8 = d.s8 = d.s8 || {}; s8.fixed = s8.fixed || {};
  FIXED8.forEach(function (a) { var x = s8.fixed[a.k] = s8.fixed[a.k] || {}; ['sit', 'accion', 'resp'].forEach(function (c) { fill(x, c); }); });
  (s8.otros || []).forEach(function (r) { ['sit', 'accion', 'resp'].forEach(function (c) { fill(r, c); }); });
  var s9 = d.s9 = d.s9 || {}, ec = get(ed, 's9.crit') || [];
  if (val === 'No aplica') { s9.crit = s9.crit || []; CRIT9.forEach(function (q, i) { if (blank(s9.crit[i]) && blank(ec[i])) { s9.crit[i] = 'No aplica'; n++; } }); }
  var s12 = d.s12 = d.s12 || {};
  ROLES.forEach(function (r) { var x = s12[r[0]] = s12[r[0]] || {}, eo = get(ed, 's12.' + r[0]) || {}; fill(x, 'nombre', eo); fill(x, 'fecha', eo); });
  return n;
}

/* ---------- Redacción automática (sin IA: arma el texto con los datos registrados) ---------- */
function real(v) { return !blank(v) && STD.indexOf(String(v).trim()) < 0 && ['Sin novedades', 'Ninguna'].indexOf(String(v).trim()) < 0; }
function kpiRows(p) { var s7 = p.data.s7 || {}; return s7.na ? [] : (s7.rows || []).filter(function (r) { return real(r.kpi); }); }
var EST = { 'Sí': 'cumple', 'Parcial': 'cumple parcialmente', 'No': 'no cumple' };
function kpiResumen(rows) {
  var n = rows.length, si = 0, par = 0, no = 0;
  rows.forEach(function (r) { if (r.cumple === 'Sí') si++; else if (r.cumple === 'Parcial') par++; else if (r.cumple === 'No') no++; });
  if (!n) return '';
  var t = si + ' de ' + n + ' KPI ' + (si === 1 ? 'cumple' : 'cumplen') + ' la meta';
  if (par) t += ', ' + par + ' parcialmente';
  if (no) t += (par ? ' y ' : ', ') + no + (no === 1 ? ' no la cumple' : ' no la cumplen');
  var sin = n - si - par - no; if (sin) t += ' (' + sin + ' sin evaluar)';
  return t + '.';
}
function autoCrit2(p) {
  var rows = kpiRows(p).filter(function (r) { return EST[r.cumple]; });
  if (!rows.length) return '';
  if (rows.every(function (r) { return r.cumple === 'Sí'; })) return 'Sí';
  if (rows.every(function (r) { return r.cumple === 'No'; })) return 'No';
  return 'Parcialmente o requiere ajustes';
}
function autoComparacion(p) {
  var d = p.data, lb = get(d, 's2.lineaBase'), rows = kpiRows(p), out = [];
  if (real(lb)) out.push('Línea base: ' + String(lb).trim().replace(/\.$/, '') + '.');
  if (rows.length) {
    out.push('Resultados obtenidos frente a la meta:');
    rows.forEach(function (r) {
      out.push('• ' + r.kpi.trim() + ': ' + (real(r.base) ? 'línea base ' + r.base.trim() + '; ' : '') + 'meta ' + (real(r.meta) ? r.meta.trim() : 'no definida') + '; resultado ' + (real(r.resultado) ? r.resultado.trim() : 'no registrado') + ' (' + (EST[r.cumple] || 'sin evaluar') + ').');
    });
    out.push(kpiResumen(rows));
  } else out.push('No se registraron KPI para comparar con la línea base.');
  var s5 = d.s5 || {}, par = s5.na ? [] : (s5.rows || []).filter(function (r) { return real(r.parametro) && real(r.valor); });
  if (par.length) out.push('Condiciones registradas: ' + par.map(function (r) { return r.parametro.trim() + ' ' + r.valor.trim(); }).join('; ') + '.');
  return out.join('\n');
}
function autoConclusion(p) {
  var d = p.data, c = (get(d, 's9.crit') || []).slice(), out = [];
  var ej = { 'Sí': 'La prueba se ejecutó según lo previsto', 'Parcialmente o requiere ajustes': 'La prueba se ejecutó con ajustes respecto a lo previsto', 'No': 'La prueba no pudo ejecutarse según lo previsto' }[c[0]];
  var re = { 'Sí': 'se obtuvo el resultado esperado', 'Parcialmente o requiere ajustes': 'el resultado esperado se obtuvo de forma parcial', 'No': 'no se obtuvo el resultado esperado' }[c[1]];
  if (ej && re) out.push(ej + ' y ' + re + '.'); else if (ej) out.push(ej + '.'); else if (re) out.push(re.charAt(0).toUpperCase() + re.slice(1) + '.');
  var kr = kpiResumen(kpiRows(p)); if (kr) out.push('En cuanto a los indicadores, ' + kr.charAt(0).toLowerCase() + kr.slice(1));
  var fx = get(d, 's8.fixed') || {}, prob = [], mej = [];
  FIXED8.forEach(function (a) {
    var x = fx[a.k] || {}; if (!real(x.sit)) return;
    var lbl = { dificultades: 'dificultades', faltante: 'equipo o recurso faltante', riesgo: 'riesgo', mejora: 'mejora al proceso', pendiente: 'información pendiente', oportunidad: 'oportunidad' }[a.k] || a.label.toLowerCase();
    var sit = x.sit.trim().replace(/\.$/, ''); sit = sit.charAt(0).toLowerCase() + sit.slice(1);
    var t = lbl + ': ' + sit;
    if (a.k === 'mejora' || a.k === 'oportunidad') mej.push(t + (real(x.accion) ? ' (propuesta: ' + x.accion.trim().replace(/\.$/, '') + ')' : ''));
    else prob.push(t);
  });
  (get(d, 's8.otros') || []).forEach(function (o) { if (real(o.sit)) prob.push((real(o.aspecto) ? o.aspecto.trim().toLowerCase() + ': ' : '') + o.sit.trim().replace(/\.$/, '')); });
  if (prob.length) out.push('Durante la prueba se identificaron ' + prob.join('; ') + '.');
  if (mej.length) out.push('Como mejoras u oportunidades: ' + mej.join('; ') + '.');
  var rec;
  if (c[3] === 'Sí') rec = 'Se recomienda continuar con la iniciativa' + (p.gate === 'favorable' ? ' y avanzar hacia la implementación' + (real(get(d, 's10.decision')) && get(d, 's10.decision') !== 'Implementar' ? ' (' + get(d, 's10.decision').toLowerCase() + ')' : '') : '') + '.';
  else if (c[3] === 'Parcialmente o requiere ajustes') rec = 'Se recomienda continuar con la iniciativa, aplicando los ajustes identificados antes de implementarla.';
  else if (c[3] === 'No') rec = 'No se recomienda continuar con la iniciativa en las condiciones evaluadas.';
  else if (p.gate === 'no') rec = 'La prueba no se considera favorable; no se recomienda avanzar a la implementación en las condiciones actuales.';
  else if (p.gate === 'favorable') rec = 'La prueba se considera favorable y se recomienda avanzar hacia la implementación.';
  if (rec) out.push(rec);
  if (c[4] === 'Sí') out.push('Se requiere la validación de un experto técnico antes de implementar.');
  return out.join(' ') || 'Complete la evaluación (criterios y KPI) para redactar la conclusión.';
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
  document.querySelectorAll('.modal-bg').forEach(function (m) { m.remove(); });
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
  h += '<div class="card"><div class="f" style="margin:0"><label>Su nombre</label><div class="prompt">Se usa para "Elaborado por" y en el botón 👤 Yo. Se guarda en este celular.</div><input class="in" data-myname value="' + esc(myName()) + '" placeholder="Ej.: Ing. Marta López" autocomplete="off"></div></div>';
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
  return '<input type="text"' + (o.mode ? ' inputmode="' + o.mode + '"' : '') + (o.ph ? ' placeholder="' + esc(o.ph) + '"' : '') + attrs + ' value="' + esc(val) + '" autocomplete="off">';
}
function chips(path, val, withName) {
  var me = withName && myName() ? '<button class="chip' + (val === myName() ? ' sel' : '') + '" data-setp="' + esc(path) + '" data-val="' + esc(myName()) + '">👤 Yo</button>' : '';
  return '<div class="chips">' + me + STD.map(function (s) { return '<button class="chip' + (val === s ? ' sel' : '') + '" data-setp="' + esc(path) + '" data-val="' + esc(s) + '">' + s + '</button>'; }).join('') + '</div>';
}
function seg(path, options, val, extraStd, autoVal, autoNote) {
  var opts = options.map(function (o) { return Array.isArray(o) ? o : [o, o, '']; });
  if (extraStd) STD.forEach(function (s) { if (!opts.some(function (o) { return o[0] === s; })) opts.push([s, s, '']); });
  var useAuto = blank(val) && !blank(autoVal);
  return '<div class="seg">' + opts.map(function (o) {
    var cls = val === o[0] ? 'sel ' + o[2] : (useAuto && autoVal === o[0] ? 'sel auto ' + o[2] : '');
    return '<button class="' + cls + '" data-setp="' + esc(path) + '" data-val="' + esc(o[0]) + '">' + esc(o[1]) + '</button>';
  }).join('') + '</div>' + (useAuto ? '<div class="autonote">⚙️ ' + esc(autoNote || 'Automático') + '. Tóquelo para confirmarlo o elija otra opción.</div>' : '');
}
function fieldBlock(sk, f, sd, E) {
  var path = sk + '.' + f.k, h = '<div class="f"><label>' + esc(f.label) + '</label>' + (f.prompt ? '<div class="prompt">' + esc(f.prompt) + '</div>' : '');
  var ed = E ? (E.data[sk] || {}) : {};
  if (f.type === 'range') {
    if ((blank(sd.inicio) && ed.inicio) || (blank(sd.fin) && ed.fin)) h += '<div class="autonote" style="margin:0 0 6px">⚙️ Si lo deja vacío se usan las fechas del registro de campo: ' + esc(COR.fmtDate(ed.inicio || '') || '—') + ' al ' + esc(COR.fmtDate(ed.fin || '') || '—') + '.</div>';
    h += '<div class="row"><div style="flex:1;min-width:140px"><div class="muted">Inicio</div>' + inputHtml(sk + '.inicio', sd.inicio, { type: 'date' }) + '</div>' +
      '<div style="flex:1;min-width:140px"><div class="muted">Final</div>' + inputHtml(sk + '.fin', sd.fin, { type: 'date' }) + '</div></div>';
    return h + '</div>';
  }
  if (f.type === 'choice') return h + seg(path, f.options, sd[f.k], true, ed[f.k], 'Sugerido según el cumplimiento de los KPI') + '</div>';
  var ph = blank(sd[f.k]) && !blank(ed[f.k]) ? 'Automático: ' + ed[f.k] : '';
  h += inputHtml(path, sd[f.k], { type: f.type, ph: ph });
  if (f.k !== 'nombre' && f.k !== 'codigo') h += chips(path, sd[f.k], ['encargado', 'responsable', 'resp'].indexOf(f.k) >= 0);
  return h + '</div>';
}

function tableView(p, s, mode, E) {
  E = E || eff(p);
  var d = p.data, sd = d[s] = d[s] || {}, rows = sd.rows || [], h = '';
  if (sd.na && mode === 'campo') return '<p class="muted">No aplica (definido antes de la visita).</p>';
  if (sd.na) {
    return '<label class="nacheck"><input type="checkbox" data-na="' + s + '" checked> ' + ({ s5: 'No hay parámetros que controlar', s7: 'No hay KPI ni resultados que medir' }[s] || 'Esta sección no aplica') + '</label><p class="muted">En el Excel aparecerá como "No aplica".</p>';
  }
  if (s === 's13' && (p.fotos || []).length) h += '<div class="banner info">Las ' + p.fotos.length + ' foto(s) de la app se agregan solas como anexos. Aquí solo agregue otras evidencias.</div>';
  if (mode === 'campo' && !rows.length) {
    h += '<div class="empty">No se definieron ' + (s === 's7' ? 'KPI' : 'parámetros') + ' antes de la visita.<br>Puede agregarlos aquí.</div>';
  }
  rows.forEach(function (r, i) {
    var known = mode === 'campo' && COLS[s].filter(function (c) { return c.plan && !c.optional; }).every(function (c) { return !blank(r[c.k]); });
    h += '<div class="rowcard" data-s="' + s + '"><div class="rh"><span>' + ROWNAME[s] + ' ' + (i + 1) + '</span><span class="ops">' +
      (s === 's3' && i > 0 ? '<button class="iconb" data-up="' + i + '" data-s="' + s + '" aria-label="Subir">↑</button>' : '') +
      '<button class="iconb" data-rowna="' + i + '" data-s="' + s + '" title="Vacíos → No aplica">NA</button>' +
      '<button class="iconb" data-delrow="' + i + '" data-s="' + s + '" aria-label="Eliminar">🗑</button></span></div><div class="rb">';
    if (known) h += '<div class="given">' + COLS[s].filter(function (c) { return c.plan && !blank(r[c.k]); }).map(function (c) { return '<b>' + esc(c.label.replace(/ \(.*\)$/, '')) + ':</b> ' + esc(r[c.k]); }).join('<br>') + '</div>';
    COLS[s].forEach(function (c) {
      if (mode === 'plan' && !c.plan) return;
      if (known && c.plan) return;
      var path = s + '.rows.' + i + '.' + c.k;
      var er = (get(E.data, s + '.rows') || [])[i] || {};
      var ctl = c.choice ? '<div data-cum="' + i + '">' + cumSeg(path, r, er) + '</div>'
        : inputHtml(path, r[c.k], { type: c.type, big: c.big && mode === 'campo', mode: c.mode, ph: c.k === 'porcentaje' && blank(r[c.k]) && er[c.k] ? 'Automático: ' + er[c.k] : '' });
      h += '<div class="f"><label>' + esc(c.label) + '</label>' + ctl + '</div>';
    });
    h += '</div></div>';
  });
  var addLabel = { s3: '+ Agregar etapa', s4: '+ Agregar recurso o insumo', s5: mode === 'campo' ? '+ Agregar parámetro no previsto' : '+ Agregar parámetro', s7: mode === 'campo' ? '+ Agregar resultado no previsto' : '+ Agregar KPI o resultado', s13: '+ Agregar anexo' }[s];
  h += '<button class="btn block sky" data-addrow="' + s + '">' + addLabel + '</button>';
  if (mode !== 'campo') h += '<label class="nacheck" style="margin-top:12px"><input type="checkbox" data-na="' + s + '"> ' + ({ s5: 'No hay parámetros que controlar', s7: 'No hay KPI ni resultados que medir' }[s] || 'Esta sección no aplica') + '</label>';
  if (s === 's4') h += '<p class="note">Complete el porcentaje solo cuando corresponda y con una misma base de unidades. <button class="chip" data-act="pct">Calcular % automáticamente</button></p>';
  return h;
}

function cumSeg(path, r, er) {
  var a = blank(r.cumple) ? er.cumple || evalKpi(r.meta, r.resultado) : '';
  return seg(path, [['Sí', 'Sí', 'si'], ['Parcial', 'Parcial', 'par'], ['No', 'No', 'no']], r.cumple, false, a, 'Calculado comparando el resultado con la meta');
}
function s8View(p) {
  var sd = p.data.s8 = p.data.s8 || {}, fx = sd.fixed = sd.fixed || {}, h = '';
  FIXED8.forEach(function (a) {
    var x = fx[a.k] || {};
    h += '<div class="rowcard"><div class="rh"><span>' + esc(a.label) + '</span><span class="ops"><button class="iconb wide" data-sinnov="' + a.k + '">✓ Sin novedades</button><button class="iconb" data-fill8="' + a.k + '">NA</button></span></div><div class="rb">' +
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
function s9View(p, E) {
  E = E || eff(p);
  var sd = p.data.s9 = p.data.s9 || {}, ec = get(E.data, 's9.crit') || [], h = '';
  var g = p.gate || '', ag = !g ? E.gate : '';
  var gcls = function (v, c) { return g === v ? 'sel ' + c : ag === v ? 'sel auto ' + c : ''; };
  h += '<div class="gate"><b>¿La prueba es favorable y se autoriza avanzar hacia la implementación?</b>' +
    '<div class="seg"><button class="' + gcls('favorable', 'si') + '" data-gate="favorable">Sí → se llena la secc. 8</button><button class="' + gcls('no', 'no') + '" data-gate="no">No → secc. 8 "No aplica"</button></div>' +
    (ag ? '<div class="autonote">⚙️ Sugerido según el cumplimiento de los KPI. Tóquelo para confirmarlo o elija la otra opción.</div>' : '') + '</div>';
  CRIT9.forEach(function (q, i) {
    h += '<div class="crit"><div class="q">' + esc(q) + '</div>' + seg('s9.crit.' + i, CRIT_OPTS, (sd.crit || [])[i], false, ec[i], 'Sugerido según el cumplimiento de los KPI') + '</div>';
  });
  h += '<div class="banner info">✨ La comparación y la conclusión se redactan solas con los datos registrados. Toque <b>Redactar</b> para verlas y ajustarlas. Si las deja vacías, se redactan al generar el Excel.</div>';
  h += '<div class="f"><label>Comparación entre la línea base y los resultados obtenidos</label><div class="prompt">Se arma con la línea base, los KPI (meta, resultado, cumplimiento). Si hay gráficos comparativos, en el Excel este apartado muestra solo los gráficos; este texto se usa cuando no hay gráficos.</div>' + inputHtml('s9.grafico', sd.grafico, { type: 'area', rows: 6 }) + '<div style="margin-top:8px"><button class="btn sm sky" data-auto="grafico">✨ Redactar</button></div></div>';
  h += '<div class="f"><label>Conclusión y recomendación</label><div class="prompt">Se arma con los criterios, los KPI, las dificultades y mejoras de la secc. 6 y la decisión.</div>' + inputHtml('s9.conclusion', sd.conclusion, { type: 'area', rows: 7 }) + '<div style="margin-top:8px"><button class="btn sm sky" data-auto="conclusion">✨ Redactar</button></div></div>';
  return h;
}
function s12View(p, E) {
  var sd = p.data.s12 = p.data.s12 || {}, h = '';
  if (!myName()) h += '<div class="banner warn">Configure <b>su nombre</b> en la pantalla de inicio para que "Elaborado por" se llene solo.</div>';
  ROLES.forEach(function (r) {
    var x = sd[r[0]] || {}, el = r[0] === 'elab';
    h += '<div class="rowcard"><div class="rh"><span>' + r[1] + '</span></div><div class="rb">' +
      '<div class="f"><label>Nombre</label>' + inputHtml('s12.' + r[0] + '.nombre', x.nombre, { ph: el && myName() ? 'Automático: ' + myName() : '' }) + chips('s12.' + r[0] + '.nombre', x.nombre, true) + '</div>' +
      '<div class="f"><label>Fecha</label>' + inputHtml('s12.' + r[0] + '.fecha', x.fecha, { type: 'date' }) + (el && blank(x.fecha) ? '<div class="autonote">⚙️ Si la deja vacía se usa la fecha en que se genera el Excel.</div>' : '') + '</div></div></div>';
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
  var s = vw.s, d = p.data, E = eff(p), h = top(vw.t, get(d, 's1.nombre') || '', 'v=phase&p=' + p.id + '&ph=' + ph) + '<main class="wrap">';
  if (vw.note) h += '<p class="note">' + esc(vw.note) + '</p>';
  if (s === 's1') { var sd1 = d.s1 = d.s1 || {}; F1.forEach(function (f) { h += fieldBlock('s1', f, sd1, E); }); }
  else if (s === 's2') { var sd2 = d.s2 = d.s2 || {}; F2.forEach(function (f) { h += fieldBlock('s2', f, sd2); }); }
  else if (s === 's10') {
    if (E.gate === 'no') h += '<div class="banner info">' + (p.gate ? 'En la evaluación (secc. 7) se indicó' : 'Según los KPI registrados (ninguno cumple la meta), se considera') + ' que la prueba <b>no es favorable</b>. Esta sección aparecerá como "No aplica" en el Excel.</div>';
    else { if (!E.gate) h += '<div class="banner warn">Primero responda en la secc. 7 si la prueba es favorable.</div>'; var sd10 = d.s10 = d.s10 || {}; F10.forEach(function (f) { h += fieldBlock('s10', f, sd10, E); }); }
  }
  else if (s === 's11') { var sd11 = d.s11 = d.s11 || {}; F11.forEach(function (f) { h += fieldBlock('s11', f, sd11); }); }
  else if (s === 'med') {
    h += '<div class="sub">5 · Condiciones y parámetros</div>' + tableView(p, 's5', vw.m);
    h += '<div class="sub">6 · KPI y resultados</div>' + tableView(p, 's7', vw.m);
  }
  else if (COLS[s]) h += tableView(p, s, vw.m, E);
  else if (s === 's8') h += s8View(p);
  else if (s === 's9') h += s9View(p, E);
  else if (s === 's12') h += s12View(p, E);
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
  var chs = chartData(eff(p));
  h += '<div class="card"><label style="display:flex;gap:10px;align-items:center;font-size:15px;margin-bottom:10px"><input type="checkbox" id="incGraf" style="width:22px;height:22px"' + (chs.length ? ' checked' : ' disabled') + '> Incluir gráficos comparativos en la secc. 7 (' + chs.length + ' KPI)</label>' +
    '<label style="display:flex;gap:10px;align-items:center;font-size:15px"><input type="checkbox" id="incFotos" style="width:22px;height:22px"' + (nf ? ' checked' : ' disabled') + '> Incluir hoja "Anexo fotográfico" (' + nf + ' foto' + (nf === 1 ? '' : 's') + ')</label>' +
    (chs.length ? '<div class="sub" style="margin-top:14px">Vista previa de los gráficos</div>' + chs.map(chartSvg).join('') :
      '<p class="muted" style="margin:10px 0 0">Para tener gráficos, cada KPI necesita un resultado con número y una meta o línea base con número.</p>') + '</div>';
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
/* Copia lo reutilizable de un piloto (planificación, procedimiento, recursos y KPI con meta), sin resultados */
function templateFrom(src) {
  var d = clone(src.data), o = { __v: 2 };
  var s1 = d.s1 || {}; o.s1 = { lugar: s1.lugar, responsable: s1.responsable, encargado: s1.encargado, objetivo: s1.objetivo, partes: s1.partes, resultado: s1.resultado, alcance: s1.alcance };
  if (d.s2) o.s2 = { hipotesis: d.s2.hipotesis, lineaBase: d.s2.lineaBase, metodo: d.s2.metodo };
  if (d.s3) o.s3 = d.s3;
  if (d.s4) o.s4 = { na: d.s4.na, rows: (d.s4.rows || []).map(function (r) { return { recurso: r.recurso, cantidad: r.cantidad, unidad: r.unidad, obs: r.obs }; }) };
  if (d.s7) o.s7 = { na: d.s7.na, rows: (d.s7.rows || []).map(function (r) { return { kpi: r.kpi, evalua: r.evalua, base: r.base, meta: r.meta }; }) };
  return o;
}
function newPilot(src) {
  modal('<h3>' + (src ? 'Duplicar como plantilla' : 'Nuevo piloto') + '</h3><p>' + (src ? 'Se copian los datos generales, la metodología, el procedimiento, los recursos y los KPI con su meta de "' + esc(get(src.data, 's1.nombre') || '') + '". No se copian resultados, evaluación, fotos ni fechas.' : 'Podrá completar el resto por etapas.') + '</p>' +
    '<div class="f"><label>Nombre de la iniciativa</label><input class="in" id="nNom" autocomplete="off"></div>' +
    '<div class="f"><label>Código o referencia</label><div class="prompt">Código de Wave o número de iniciativa.</div><input class="in" id="nCod" autocomplete="off"></div>' +
    '<div class="row"><button class="btn" data-close>Cancelar</button><button class="btn primary" data-ok>Crear</button></div>', function (bg, close) {
    var n = bg.querySelector('#nNom'); setTimeout(function () { n.focus(); }, 50);
    bg.querySelector('[data-ok]').onclick = function () {
      var nom = n.value.trim(), cod = bg.querySelector('#nCod').value.trim();
      if (!nom) { toast('Escriba el nombre de la iniciativa.'); return; }
      var dup = cod && Object.keys(PIL).some(function (k) { return (get(PIL[k].data, 's1.codigo') || '').toLowerCase() === cod.toLowerCase(); });
      if (dup) { toast('Ya hay un piloto con ese código en este celular.'); return; }
      var data = src ? templateFrom(src) : { __v: 2, s1: {} };
      data.s1 = data.s1 || {}; data.s1.nombre = nom; data.s1.codigo = cod;
      var p = { id: uid(), createdAt: new Date().toISOString(), data: data, fotos: [], gate: '' };
      PIL[p.id] = p; save(p, true); close();
      go({ v: 'pilot', p: p.id });
    };
  });
}
function pilotMenu(p) {
  modal('<h3>' + esc(get(p.data, 's1.nombre') || 'Piloto') + '</h3><p>Opciones del piloto</p>' +
    '<button class="btn block sky" data-m="share" style="margin-bottom:10px">📤 Compartir piloto (para que otra persona lo continúe)</button>' +
    '<button class="btn block sky" data-m="dup" style="margin-bottom:10px">📄 Duplicar como plantilla para otro piloto</button>' +
    '<button class="btn block danger" data-m="del" style="margin-bottom:10px">🗑 Eliminar de este celular</button>' +
    '<button class="btn block" data-close>Cerrar</button>', function (bg, close) {
    bg.querySelector('[data-m="share"]').onclick = function () { close(); exportPilot(p); };
    bg.querySelector('[data-m="dup"]').onclick = function () { close(); newPilot(p); };
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
    var download = function (msg) {
      var u = URL.createObjectURL(blob), a = document.createElement('a');
      a.href = u; a.download = name; document.body.appendChild(a); a.click();
      setTimeout(function () { URL.revokeObjectURL(u); a.remove(); }, 4000);
      if (msg) showSaved(); else toast('Archivo guardado en Descargas.');
    };
    var showSaved = function () {
      var box = bg.querySelector('.modal');
      box.innerHTML = '<h3>Guardado en Descargas</h3><p>Este celular no permite compartir archivos de Excel directamente desde el navegador, así que se guardó en <b>Descargas</b>:<br><b>' + esc(name) + '</b></p>' +
        '<ol class="steps"><li>Abra la app <b>Archivos</b> (o la notificación de descarga).</li><li>Entre a <b>Descargas</b> y mantenga presionado el archivo.</li><li>Toque <b>Compartir</b> y elija WhatsApp, correo u OneDrive.</li></ol>' +
        '<button class="btn block primary" data-close style="margin-top:8px">Entendido</button>';
    };
    var sh = bg.querySelector('[data-share]');
    if (sh) sh.onclick = function () {
      navigator.share({ files: [file], title: name }).then(close, function (e) {
        if (e && e.name === 'AbortError') return;
        download(true); /* el celular no deja compartir este tipo de archivo: se guarda y se explica cómo enviarlo */
      });
    };
    bg.querySelector('[data-dl]').onclick = function () { download(false); };
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
      return putPhoto(id, p.id, c.blob).then(function () { stampCampo(p); p.fotos = p.fotos || []; p.fotos.push({ id: id, w: c.w, h: c.h, caption: '', at: new Date().toISOString() }); n++; });
    });
  });
  chain.then(function () { save(p, true); render(); toast(n + ' foto(s) guardada(s). Agregue una descripción.'); }, function () { toast('No se pudieron guardar las fotos.'); render(); });
}

/* ---------- Excel ---------- */
function makeExcel(p) {
  var btn = document.getElementById('btnExcel');
  if (btn) { btn.disabled = true; btn.innerHTML = '<span class="spin"></span> Generando…'; }
  var inc = document.getElementById('incFotos'), withPhotos = inc && inc.checked;
  var E = eff(p), rec = E.data, fotos = p.fotos || [];
  var ig = document.getElementById('incGraf'), charts = ig && ig.checked ? chartData(E) : [];
  if (blank(rec.s9.grafico)) rec.s9.grafico = autoComparacion(E);
  if (blank(rec.s9.conclusion)) rec.s9.conclusion = autoConclusion(E);
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
    return COR.buildXlsx(JSZip, rec, self.COR_XLSX_B64, false, { photos: photos, charts: charts });
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
  var t = e.target;
  if (t.hasAttribute('data-myname')) { setMyName(t.value.trim()); return; }
  var p = cur(); if (!p) return;
  if (t.dataset.p) {
    set(p.data, t.dataset.p, t.value);
    var wrap = t.closest('.f'); if (wrap) wrap.querySelectorAll('.chip.sel').forEach(function (c) { c.classList.remove('sel'); });
    save(p);
    var mk = /^s7\.rows\.(\d+)\.(resultado|meta)$/.exec(t.dataset.p);
    if (mk) { var box = app.querySelector('[data-cum="' + mk[1] + '"]'); if (box) box.innerHTML = cumSeg('s7.rows.' + mk[1] + '.cumple', p.data.s7.rows[+mk[1]], {}); }
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
  if (ds.auto) {
    var EE = eff(p), key = ds.auto, txt = key === 'grafico' ? autoComparacion(EE) : autoConclusion(EE);
    var apply = function () { set(p.data, 's9.' + key, txt); save(p); keepScroll(render); toast('Texto redactado. Revíselo y ajústelo si hace falta.'); };
    if (!blank(get(p.data, 's9.' + key))) confirmBox('Reemplazar texto', 'Ya hay texto en este campo. ¿Reemplazarlo por la redacción automática?', 'Reemplazar', apply); else apply();
    return;
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
    var cards = app.querySelectorAll('.rowcard[data-s="' + ds.addrow + '"]'); var last = cards[cards.length - 1];
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
  if (ds.sinnov) { var fs8 = p.data.s8.fixed[ds.sinnov] = p.data.s8.fixed[ds.sinnov] || {}; if (blank(fs8.sit)) fs8.sit = 'Sin novedades'; if (blank(fs8.accion)) fs8.accion = 'Ninguna'; if (blank(fs8.resp)) fs8.resp = 'No aplica'; save(p); keepScroll(render); return; }
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
