/* Generador del Excel COR-GPPT-R006 (V2 simplificada) a partir de un registro del formulario.
   Llena la plantilla original celda por celda: conserva logos, estilos, casillas y saltos de página.
   Funciona en el navegador (window.JSZip) y en Node (require('jszip')) para pruebas. */
(function (g) {
  'use strict';
  var STD = ['No aplica', 'No registrado', 'Pendiente'];
  var FIXED8 = [
    { k: 'dificultades', label: 'Dificultades durante la prueba' },
    { k: 'faltante', label: 'Equipo o recurso faltante' },
    { k: 'riesgo', label: 'Riesgo identificado' },
    { k: 'mejora', label: 'Mejora al proceso' },
    { k: 'pendiente', label: 'Información pendiente' },
    { k: 'oportunidad', label: 'Oportunidad identificada' }
  ];
  var CRIT9 = [
    '¿Fue posible ejecutar la prueba según lo previsto?',
    '¿Se obtuvo el resultado esperado?',
    '¿Se cumplieron los KPI o criterios definidos?',
    '¿Se recomienda continuar con la iniciativa?',
    '¿Requiere validación de un experto técnico?'
  ];
  var VALUE_FONT_ID = 9; /* Aptos 8,5 de la plantilla */

  function fmtDate(s) {
    if (!s) return '';
    if (STD.indexOf(s) >= 0) return s;
    var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s);
    return m ? m[3] + '/' + m[2] + '/' + m[1] : s;
  }
  function v(x) { return x == null ? '' : String(x).trim(); }
  function xesc(s) {
    return String(s).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  /* ---------- Referencias de celda ---------- */
  function colNum(c) { var n = 0; for (var i = 0; i < c.length; i++) n = n * 26 + (c.charCodeAt(i) - 64); return n; }
  function colName(n) { var s = ''; while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }
  function splitRef(ref) { var m = /^\$?([A-Z]+)\$?(\d+)$/.exec(ref); return { c: m[1], r: +m[2] }; }

  /* ---------- Modelo mínimo de la hoja ---------- */
  function parseSheet(xml) {
    var a = xml.indexOf('<sheetData>'), b = xml.indexOf('</sheetData>');
    var head = xml.slice(0, a), body = xml.slice(a + 11, b), tail = xml.slice(b + 12);
    var rows = {};
    var re = /<row\b([^>]*?)(\/>|>([\s\S]*?)<\/row>)/g, m;
    while ((m = re.exec(body))) {
      var attrs = m[1], r = +/\br="(\d+)"/.exec(attrs)[1], cells = {};
      if (m[3]) {
        var rc = /<c\b([^>]*?)(\/>|>([\s\S]*?)<\/c>)/g, k;
        while ((k = rc.exec(m[3]))) {
          var ca = k[1], ref = /\br="([A-Z]+)\d+"/.exec(ca)[1];
          cells[ref] = { attrs: ca.replace(/\s*\br="[A-Z]+\d+"/, ''), inner: k[3] == null ? null : k[3] };
        }
      }
      rows[r] = { attrs: attrs.replace(/\s*\br="\d+"/, ''), cells: cells };
    }
    var merges = [];
    tail = tail.replace(/<mergeCells[^>]*>([\s\S]*?)<\/mergeCells>/, function (_, inner) {
      inner.replace(/<mergeCell ref="([^"]+)"\/>/g, function (__, ref) {
        var p = ref.split(':'), s = splitRef(p[0]), e = splitRef(p[1]);
        merges.push({ c1: colNum(s.c), r1: s.r, c2: colNum(e.c), r2: e.r });
      });
      return '@@MERGES@@';
    });
    return { head: head, tail: tail, rows: rows, merges: merges };
  }
  function serializeSheet(S) {
    var nums = Object.keys(S.rows).map(Number).sort(function (a, b) { return a - b; });
    var body = nums.map(function (r) {
      var row = S.rows[r];
      var cols = Object.keys(row.cells).sort(function (a, b) { return colNum(a) - colNum(b); });
      var attrs = ' r="' + r + '"' + row.attrs;
      if (!cols.length) return '<row' + attrs + '/>';
      return '<row' + attrs + '>' + cols.map(function (c) {
        var cell = row.cells[c];
        return '<c r="' + c + r + '"' + cell.attrs + (cell.inner == null ? '/>' : '>' + cell.inner + '</c>');
      }).join('') + '</row>';
    }).join('');
    var mc = '<mergeCells count="' + S.merges.length + '">' + S.merges.map(function (m) {
      return '<mergeCell ref="' + colName(m.c1) + m.r1 + ':' + colName(m.c2) + m.r2 + '"/>';
    }).join('') + '</mergeCells>';
    var last = nums.length ? nums[nums.length - 1] : 1;
    var head = S.head.replace(/<dimension ref="[^"]*"\/>/, '<dimension ref="B2:AO' + last + '"/>');
    return head + '<sheetData>' + body + '</sheetData>' + S.tail.replace('@@MERGES@@', mc);
  }

  /* Inserta `times` copias (tal como están en la plantilla) del bloque de filas [r1..r2] justo debajo de r2. */
  function insertBlock(S, wb, r1, r2, times) {
    if (times <= 0) return;
    var h = r2 - r1 + 1, add = h * times;
    var sh = function (r) { return r > r2 ? r + add : r; };
    var newRows = {};
    Object.keys(S.rows).forEach(function (k) { newRows[sh(+k)] = S.rows[k]; });
    for (var t = 1; t <= times; t++) {
      for (var r = r1; r <= r2; r++) {
        var src = S.rows[r], cells = {};
        if (!src) continue;
        Object.keys(src.cells).forEach(function (c) {
          var cell = src.cells[c];
          cells[c] = { attrs: cell.attrs, inner: cell.inner };
        });
        newRows[r + h * t] = { attrs: src.attrs, cells: cells };
      }
    }
    S.rows = newRows;
    var extra = [];
    S.merges.forEach(function (m) {
      if (m.r1 >= r1 && m.r2 <= r2) for (var t = 1; t <= times; t++) extra.push({ c1: m.c1, c2: m.c2, r1: m.r1 + h * t, r2: m.r2 + h * t });
      m.r1 = sh(m.r1); m.r2 = sh(m.r2);
    });
    S.merges = S.merges.concat(extra);
    var shiftRange = function (rng) {
      return rng.split(' ').map(function (part) {
        var p = part.split(':'), a = splitRef(p[0]), b = p[1] ? splitRef(p[1]) : a;
        var out = [a.c + sh(a.r) + (p[1] ? ':' + b.c + sh(b.r) : '')];
        if (a.r >= r1 && b.r <= r2) for (var t = 1; t <= times; t++) out.push(a.c + (a.r + h * t) + (p[1] ? ':' + b.c + (b.r + h * t) : ''));
        return out.join(' ');
      }).join(' ');
    };
    S.tail = S.tail.replace(/sqref="([^"]+)"/g, function (_, rng) { return 'sqref="' + shiftRange(rng) + '"'; });
    S.tail = S.tail.replace(/<brk id="(\d+)"/g, function (_, id) { return '<brk id="' + sh(+id) + '"'; });
    wb.printEnd = sh(wb.printEnd);
  }

  /* ---------- Estilos: copia del estilo de la celda con ajuste de texto ---------- */
  function StyleBook(xml) {
    this.xml = xml; this.map = {};
    var m = /<cellXfs count="(\d+)">([\s\S]*?)<\/cellXfs>/.exec(xml);
    this.count = +m[1];
    this.xfs = m[2].match(/<xf\b[^>]*?(?:\/>|>[\s\S]*?<\/xf>)/g);
    this.added = [];
  }
  StyleBook.prototype.value = function (s, align) {
    var key = s + '|' + (align || 'left');
    if (this.map[key] != null) return this.map[key];
    var base = this.xfs[+s] || this.xfs[0];
    var open = /^<xf\b[^>]*?(?=\/?>)/.exec(base)[0]
      .replace(/\s*\bfontId="\d+"/, '').replace(/\s*\bapplyFont="\d"/, '').replace(/\s*\bapplyAlignment="\d"/, '');
    var xf = open + ' fontId="' + VALUE_FONT_ID + '" applyFont="1" applyAlignment="1"><alignment horizontal="' + (align || 'left') + '" vertical="center" wrapText="1"/></xf>';
    this.added.push(xf);
    return (this.map[key] = this.count + this.added.length - 1);
  };
  StyleBook.prototype.toXml = function () {
    var self = this;
    return this.xml.replace(/<cellXfs count="\d+">([\s\S]*?)<\/cellXfs>/, function (_, inner) {
      return '<cellXfs count="' + (self.count + self.added.length) + '">' + inner + self.added.join('') + '</cellXfs>';
    });
  };

  /* ---------- Llenado ---------- */
  function Writer(S, styles) { this.S = S; this.st = styles; this.filled = []; }
  Writer.prototype.cell = function (ref) {
    var p = splitRef(ref), row = this.S.rows[p.r];
    if (!row) row = this.S.rows[p.r] = { attrs: ' spans="2:41" ht="15" customHeight="1"', cells: {} };
    return row.cells[p.c] || (row.cells[p.c] = { attrs: '', inner: null });
  };
  /* Escribe texto en la celda superior izquierda de un área combinada (o sola). */
  Writer.prototype.text = function (ref, text, align) {
    text = v(text);
    var c = this.cell(ref), sm = /\bs="(\d+)"/.exec(c.attrs), s = sm ? sm[1] : '0';
    var ns = align === 'keep' ? s : this.st.value(s, align);
    c.attrs = ' s="' + ns + '"' + (text ? ' t="inlineStr"' : '');
    c.inner = text ? '<is><t xml:space="preserve">' + xesc(text) + '</t></is>' : null;
    if (text) this.filled.push({ ref: ref, text: text });
  };
  /* Cambia un ☐ existente por ☒ (o lo deja en ☐). */
  Writer.prototype.box = function (ref, on) {
    var c = this.cell(ref);
    c.attrs = c.attrs.replace(/\s*\bt="[^"]*"/, '') + ' t="inlineStr"';
    c.inner = '<is><t>' + (on ? '☒' : '☐') + '</t></is>';
  };

  /* Ajusta la altura de las filas para que el texto quepa (Excel no lo hace en celdas combinadas). */
  function fitHeights(S, filled) {
    var COLPT = 18.4, CHARPT = 4.55, LINEPT = 11.6, PAD = 5;
    var need = {};
    filled.forEach(function (f) {
      var p = splitRef(f.ref), c = colNum(p.c);
      var m = S.merges.find(function (x) { return x.r1 === p.r && x.c1 === c; }) || { c1: c, c2: c, r1: p.r, r2: p.r };
      var cpl = Math.max(4, Math.floor(((m.c2 - m.c1 + 1) * COLPT - 6) / CHARPT));
      var lines = f.text.split('\n').reduce(function (a, ln) { return a + Math.max(1, Math.ceil(ln.length * 1.08 / cpl)); }, 0);
      var h = lines * LINEPT + PAD, n = m.r2 - m.r1 + 1, cur = 0;
      for (var r = m.r1; r <= m.r2; r++) cur += rowHeight(S, r);
      if (h <= cur) return;
      var per = Math.ceil(h / n * 4) / 4;
      for (var r2 = m.r1; r2 <= m.r2; r2++) need[r2] = Math.max(need[r2] || 0, per);
    });
    Object.keys(need).forEach(function (r) {
      if (need[r] > rowHeight(S, +r)) setRowHeight(S, +r, Math.min(need[r], 400));
    });
  }
  function rowHeight(S, r) {
    var row = S.rows[r]; if (!row) return 15;
    var m = /\bht="([\d.]+)"/.exec(row.attrs); return m ? +m[1] : 15;
  }
  function setRowHeight(S, r, h) {
    var row = S.rows[r];
    if (!row) row = S.rows[r] = { attrs: ' spans="2:41"', cells: {} };
    var rest = row.attrs.replace(/\s*\bht="[^"]*"/, '').replace(/\s*\bcustomHeight="[^"]*"/, '').trim();
    row.attrs = ' ht="' + h + '" customHeight="1"' + (rest ? ' ' + rest : '');
  }

  /* ---------- Mapeo registro → plantilla V2 ---------- */
  function tableRows(sec) {
    sec = sec || {};
    if (sec.na) return { na: true, rows: [] };
    return { na: false, rows: sec.rows || [] };
  }

  function fill(S, wb, styles, d) {
    var W = new Writer(S, styles), off = 0;
    var s1 = d.s1 || {}, s2 = d.s2 || {}, s8 = d.s8 || {}, s9 = d.s9 || {}, s10 = d.s10 || {}, s11 = d.s11 || {}, s12 = d.s12 || {};
    var R = function (r) { return r + off; };

    /* 1. Datos generales */
    W.text('J14', s1.nombre); W.text('AD14', s1.codigo);
    W.text('J16', 'Inicio: ' + fmtDate(s1.inicio) + '\n\nFinal: ' + fmtDate(s1.fin));
    W.text('AD16', s1.lugar);
    W.text('J19', s1.responsable); W.text('AD19', s1.encargado);
    W.text('J21', s1.objetivo); W.text('AD21', s1.partes);
    W.text('J24', s1.resultado); W.text('AD24', s1.evidenciaCom);
    W.text('J27', s1.alcance);
    /* 2. Metodología */
    W.text('O36', s2.hipotesis); W.text('O38', s2.lineaBase); W.text('O43', s2.metodo);

    /* Tablas dinámicas: plantilla con N filas; si hay más, se insertan copias de la última. */
    function dyn(first, cap, h, list, write) {
      var extra = Math.max(0, list.length - cap);
      var lastStart = R(first) + (cap - 1) * h;
      insertBlock(S, wb, lastStart, lastStart + h - 1, extra);
      list.forEach(function (row, i) { write(R(first) + i * h, row, i); });
      off += extra * h;
    }
    function naList(t, n) { return t.na ? [{ __na: true }] : t.rows; }

    /* 3. Proceso (7 filas: 53–59) */
    var t3 = tableRows(d.s3);
    dyn(53, 7, 1, naList(t3), function (r, x, i) {
      W.text('B' + r, i + 1, 'keep');
      if (x.__na) { W.text('D' + r, 'No aplica'); return; }
      W.text('D' + r, x.etapa); W.text('M' + r, x.desc); W.text('AB' + r, x.tiempo, 'center'); W.text('AF' + r, x.equipo);
    });
    /* 4. Recursos (6 filas: 66–71) */
    dyn(66, 6, 1, naList(tableRows(d.s4)), function (r, x) {
      if (x.__na) { W.text('B' + r, 'No aplica'); return; }
      W.text('B' + r, x.recurso); W.text('N' + r, x.cantidad, 'center'); W.text('T' + r, x.unidad, 'center'); W.text('Y' + r, x.porcentaje, 'center'); W.text('AD' + r, x.obs);
    });
    /* 5. Condiciones (5 filas: 80–84) */
    dyn(80, 5, 1, naList(tableRows(d.s5)), function (r, x) {
      if (x.__na) { W.text('B' + r, 'No aplica'); return; }
      W.text('B' + r, x.aspecto); W.text('L' + r, x.parametro); W.text('X' + r, x.valor, 'center'); W.text('AE' + r, x.obs);
    });
    /* 6. KPI y resultados (6 bloques de 2 filas: 90–101) */
    dyn(90, 6, 2, naList(tableRows(d.s7)), function (r, x) {
      if (x.__na) { W.text('B' + r, 'No aplica'); return; }
      W.text('B' + r, x.kpi); W.text('J' + r, x.evalua); W.text('U' + r, x.meta); W.text('AC' + r, x.resultado);
      W.box('AJ' + r, x.cumple === 'Sí'); W.box('AL' + r, x.cumple === 'Parcial'); W.box('AN' + r, x.cumple === 'No');
    });
    /* 7. Riesgos (filas fijas 106,108,110–113; "Otro" 114–117) */
    var fx = s8.fixed || {}, fixedRows = [106, 108, 110, 111, 112, 113];
    FIXED8.forEach(function (a, i) {
      var x = fx[a.k] || {}, r = R(fixedRows[i]);
      W.text('K' + r, x.sit); W.text('V' + r, x.accion); W.text('AI' + r, x.resp);
    });
    var otros = (s8.otros || []).filter(function (o) { return v(o.aspecto) || v(o.sit) || v(o.accion) || v(o.resp); });
    dyn(114, 4, 1, otros, function (r, x) {
      W.text('B' + r, 'Otro: ' + v(x.aspecto), 'keep'); W.text('K' + r, x.sit); W.text('V' + r, x.accion); W.text('AI' + r, x.resp);
    });
    /* 8. Evaluación (criterios 124–128) */
    var opts = ['Sí', 'Parcialmente o requiere ajustes', 'No', 'No aplica'], colsC = ['U', 'X', 'AH', 'AL'];
    CRIT9.forEach(function (q, i) {
      var val = (s9.crit || [])[i];
      colsC.forEach(function (c, j) { W.box(c + R(124 + i), val === opts[j]); });
    });
    W.text('B' + R(131), s9.grafico); W.text('B' + R(137), s9.conclusion);
    /* 9. Seguimiento (148–152) */
    var trio = function (r, options, val) {
      ['Q', 'Y', 'AG'].forEach(function (c, j) { W.box(c + R(r), val === options[j]); });
    };
    if (s10.decision === 'No aplica') {
      W.text('Q' + R(150), 'No aplica'); W.text('Q' + R(152), 'No aplica');
      trio(148, [], ''); trio(149, [], ''); trio(151, [], '');
    } else {
      trio(148, ['Implementar', 'Implementar con ajustes', 'Realizar nueva prueba'], s10.decision);
      trio(149, ['Sí', 'No', 'Pendiente'], s10.cronograma);
      W.text('Q' + R(150), s10.frec);
      trio(151, ['Sí', 'No', 'Pendiente'], s10.kpiDef);
      W.text('Q' + R(152), s10.resp);
    }
    /* 10. Cierre (158–160) */
    W.text('R' + R(158), s11.ordenEjec); W.text('R' + R(159), s11.ordenLiq); W.text('R' + R(160), fmtDate(s11.fechaCierre));
    /* 11. Validación (165–167) */
    [['elab', 165], ['rev', 166], ['aprob', 167]].forEach(function (p) {
      var x = s12[p[0]] || {};
      W.text('K' + R(p[1]), x.nombre); W.text('AI' + R(p[1]), fmtDate(x.fecha), 'center');
      setRowHeight(S, R(p[1]), Math.max(rowHeight(S, R(p[1])), 30));
    });
    /* 12. Anexos (5 filas: 174–178) */
    var t13 = tableRows(d.s13);
    dyn(174, 5, 1, naList(t13), function (r, x, i) {
      W.text('B' + r, i + 1, 'keep');
      if (x.__na) { W.text('E' + r, 'No aplica'); return; }
      W.text('E' + r, x.anexo); W.text('S' + r, x.desc);
    });
    return W;
  }

  /* ---------- Hoja "Anexo fotográfico" (opcional) ----------
     photos: [{ bytes: Uint8Array (JPEG), w, h, caption }] */
  function addPhotoSheet(zip, styles, d, photos) {
    var MAXW = 640, MAXH = 350, ROWPX = 20, EMU = 9525;
    var s1 = d.s1 || {};
    var title = 'Anexo fotográfico — ' + (v(s1.nombre) || 'Prueba piloto') + (v(s1.codigo) ? ' (' + v(s1.codigo) + ')' : '');
    var sTitle = 6, sHead = 25, sCap = styles.value('22', 'left');
    var rows = [], merges = [], anchors = [], breaks = [], rels = [];
    var cellRow = function (r, ht, cells) { rows.push('<row r="' + r + '"' + (ht ? ' ht="' + ht + '" customHeight="1"' : '') + '>' + cells + '</row>'); };
    var istr = function (ref, s, t) { return '<c r="' + ref + '" s="' + s + '" t="inlineStr"><is><t xml:space="preserve">' + xesc(t) + '</t></is></c>'; };
    var blankRun = function (r, s) { var o = ''; for (var c = 3; c <= 11; c++) o += '<c r="' + colName(c) + r + '" s="' + s + '"/>'; return o; };
    cellRow(2, 22, istr('B2', sTitle, title)); merges.push('B2:K2');
    var r = 4;
    photos.forEach(function (ph, i) {
      if (i > 0 && i % 2 === 0) breaks.push(r - 1);
      var k = Math.min(1, MAXW / ph.w, MAXH / ph.h), w = Math.round(ph.w * k), h = Math.round(ph.h * k);
      cellRow(r, 18, istr('B' + r, sHead, 'Foto ' + (i + 1)) + blankRun(r, sHead)); merges.push('B' + r + ':K' + r);
      var nrows = Math.ceil((h + 10) / ROWPX);
      anchors.push('<xdr:oneCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>' + (5 * EMU) + '</xdr:colOff><xdr:row>' + r + '</xdr:row><xdr:rowOff>' + (5 * EMU) + '</xdr:rowOff></xdr:from>' +
        '<xdr:ext cx="' + w * EMU + '" cy="' + h * EMU + '"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="' + (i + 2) + '" name="Foto ' + (i + 1) + '"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>' +
        '<xdr:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rIdF' + (i + 1) + '"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>' +
        '<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="' + w * EMU + '" cy="' + h * EMU + '"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>');
      rels.push('<Relationship Id="rIdF' + (i + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/foto' + (i + 1) + '.jpeg"/>');
      zip.file('xl/media/foto' + (i + 1) + '.jpeg', ph.bytes);
      r += nrows + 1;
      var cap = v(ph.caption) || 'Sin descripción';
      cellRow(r, Math.max(18, Math.ceil(cap.length / 130) * 12 + 6), istr('B' + r, sCap, cap) + blankRun(r, sCap)); merges.push('B' + r + ':K' + r);
      r += 2;
    });
    var sheet = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr><dimension ref="B2:K' + r + '"/><sheetViews><sheetView showGridLines="0" workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/>' +
      '<cols><col min="1" max="1" width="2.77734375" customWidth="1"/><col min="2" max="11" width="9.6640625" customWidth="1"/></cols>' +
      '<sheetData>' + rows.join('') + '</sheetData><mergeCells count="' + merges.length + '">' + merges.map(function (m) { return '<mergeCell ref="' + m + '"/>'; }).join('') + '</mergeCells>' +
      '<printOptions horizontalCentered="1"/><pageMargins left="0.65" right="0.65" top="0.6" bottom="0.6" header="0.3" footer="0.3"/><pageSetup fitToHeight="0" orientation="portrait"/>' +
      (breaks.length ? '<rowBreaks count="' + breaks.length + '" manualBreakCount="' + breaks.length + '">' + breaks.map(function (b) { return '<brk id="' + b + '" max="16383" man="1"/>'; }).join('') + '</rowBreaks>' : '') +
      '<drawing r:id="rId1"/></worksheet>';
    zip.file('xl/worksheets/sheet2.xml', sheet);
    zip.file('xl/worksheets/_rels/sheet2.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing2.xml"/></Relationships>');
    zip.file('xl/drawings/drawing2.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' + anchors.join('') + '</xdr:wsDr>');
    zip.file('xl/drawings/_rels/drawing2.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + rels.join('') + '</Relationships>');
    return Promise.all([zip.file('[Content_Types].xml').async('string'), zip.file('xl/_rels/workbook.xml.rels').async('string')]).then(function (x) {
      var ct = x[0];
      if (ct.indexOf('Extension="jpeg"') < 0) ct = ct.replace('<Default Extension="xml"', '<Default Extension="jpeg" ContentType="image/jpeg"/><Default Extension="xml"');
      ct = ct.replace('</Types>', '<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/drawings/drawing2.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>');
      zip.file('[Content_Types].xml', ct);
      zip.file('xl/_rels/workbook.xml.rels', x[1].replace('</Relationships>', '<Relationship Id="rIdFotos" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/></Relationships>'));
    });
  }

  function build(JSZip, rec, templateB64, asBuffer, opts) {
    var d = rec || {}, photos = (opts && opts.photos) || [];
    return JSZip.loadAsync(templateB64, typeof templateB64 === 'string' ? { base64: true } : {}).then(function (zip) {
      return Promise.all([
        zip.file('xl/worksheets/sheet1.xml').async('string'),
        zip.file('xl/styles.xml').async('string'),
        zip.file('xl/workbook.xml').async('string')
      ]).then(function (x) {
        var S = parseSheet(x[0]), styles = new StyleBook(x[1]);
        var wb = { printEnd: 178 };
        var W = fill(S, wb, styles, d);
        fitHeights(S, W.filled);
        zip.file('xl/worksheets/sheet1.xml', serializeSheet(S));
        var wbx = x[2].replace(/\$B\$2:\$AO\$\d+/, '$B$2:$AO$' + wb.printEnd);
        if (photos.length) wbx = wbx.replace('</sheets>', '<sheet name="Anexo fotográfico" sheetId="2" r:id="rIdFotos"/></sheets>');
        zip.file('xl/workbook.xml', wbx);
        zip.remove('xl/calcChain.xml');
        return (photos.length ? addPhotoSheet(zip, styles, d, photos) : Promise.resolve()).then(function () {
          zip.file('xl/styles.xml', styles.toXml());
          return zip.generateAsync({ type: asBuffer ? 'nodebuffer' : 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', compression: 'DEFLATE' });
        });
      });
    });
  }

  g.COR = { buildXlsx: build, FIXED8: FIXED8, CRIT9: CRIT9, STD: STD, fmtDate: fmtDate };
})(typeof self !== 'undefined' ? self : globalThis);
