/* Перевод заполненного письма Word в PDF внутри устройства.
   Поддерживает то, что есть в бланке: абзацы, выравнивание, отступы, интервалы, жирный,
   таблицы с рамками, нумерацию, картинки (логотип, печать с подписью), колонтитулы. */
(function (root) {
  'use strict';
  var W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  var WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
  var A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
  var R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  var PR = 'http://schemas.openxmlformats.org/package/2006/relationships';
  var TW = 1 / 20, EMU = 1 / 12700;
  var P = root.SBPDF;

  function parse(s) { return new DOMParser().parseFromString(s, 'application/xml'); }
  function kids(n, name, ns) {
    var out = []; if (!n) return out;
    for (var c = n.firstChild; c; c = c.nextSibling) if (c.nodeType === 1 && c.namespaceURI === (ns || W) && (!name || c.localName === name)) out.push(c);
    return out;
  }
  function kid(n, name, ns) { return kids(n, name, ns)[0] || null; }
  function wa(el, name) { if (!el) return null; var v = el.getAttributeNS(W, name); return v === '' ? null : v; }
  function num(v) { return v == null ? undefined : parseFloat(v); }
  function onOff(el) { if (!el) return undefined; var v = wa(el, 'val'); return !(v === '0' || v === 'false' || v === 'off'); }
  function def(x) { return x !== undefined && x !== null; }

  /* ---------- свойства ---------- */
  function readPPr(el) {
    var p = { ind: {}, spacing: {} };
    if (!el) return p;
    var jc = kid(el, 'jc'); if (jc) p.jc = wa(jc, 'val');
    var ind = kid(el, 'ind');
    if (ind) {
      var l = num(wa(ind, 'left') || wa(ind, 'start')), r = num(wa(ind, 'right') || wa(ind, 'end'));
      if (def(l)) p.ind.left = l; if (def(r)) p.ind.right = r;
      if (def(wa(ind, 'hanging'))) { p.ind.hanging = num(wa(ind, 'hanging')); p.ind.firstLine = 0; }
      else if (def(wa(ind, 'firstLine'))) { p.ind.firstLine = num(wa(ind, 'firstLine')); p.ind.hanging = 0; }
    }
    var sp = kid(el, 'spacing');
    if (sp) ['before', 'after', 'line'].forEach(function (k) { if (def(wa(sp, k))) p.spacing[k] = num(wa(sp, k)); });
    if (sp && wa(sp, 'lineRule')) p.spacing.lineRule = wa(sp, 'lineRule');
    var np = kid(el, 'numPr');
    if (np) { var id = kid(np, 'numId'), lv = kid(np, 'ilvl'); p.numPr = { numId: id ? wa(id, 'val') : null, ilvl: lv ? +wa(lv, 'val') : 0 }; }
    var ps = kid(el, 'pStyle'); if (ps) p.style = wa(ps, 'val');
    var tabs = kid(el, 'tabs');
    if (tabs) p.tabs = kids(tabs, 'tab').filter(function (t) { return wa(t, 'val') !== 'clear'; }).map(function (t) { return num(wa(t, 'pos')); });
    return p;
  }
  function readRPr(el) {
    var r = {};
    if (!el) return r;
    var b = kid(el, 'b'); if (b) r.b = onOff(b);
    var i = kid(el, 'i'); if (i) r.i = onOff(i);
    var sz = kid(el, 'sz'); if (sz) r.sz = num(wa(sz, 'val'));
    var sp = kid(el, 'spacing'); if (sp) r.cs = num(wa(sp, 'val'));
    var rs = kid(el, 'rStyle'); if (rs) r.style = wa(rs, 'val');
    var v = kid(el, 'vanish'); if (v) r.hidden = onOff(v);
    return r;
  }
  function mergeP(a, b) {
    var o = { ind: Object.assign({}, a.ind, b.ind), spacing: Object.assign({}, a.spacing, b.spacing) };
    o.jc = def(b.jc) ? b.jc : a.jc;
    o.numPr = b.numPr || a.numPr;
    o.tabs = b.tabs || a.tabs;
    return o;
  }
  function mergeR(a, b) { var o = Object.assign({}, a); for (var k in b) if (def(b[k]) && k !== 'style') o[k] = b[k]; return o; }
  function readBorders(el) {
    var o = {}; if (!el) return o;
    ['top', 'left', 'bottom', 'right', 'insideH', 'insideV', 'start', 'end'].forEach(function (k) {
      var e = kid(el, k); if (!e) return;
      var v = wa(e, 'val');
      var key = k === 'start' ? 'left' : k === 'end' ? 'right' : k;
      o[key] = (v === 'none' || v === 'nil') ? null : { sz: num(wa(e, 'sz')) || 4 };
    });
    return o;
  }
  function readMar(el) {
    var o = {}; if (!el) return o;
    ['top', 'left', 'bottom', 'right', 'start', 'end'].forEach(function (k) {
      var e = kid(el, k); if (!e) return;
      o[k === 'start' ? 'left' : k === 'end' ? 'right' : k] = num(wa(e, 'w'));
    });
    return o;
  }

  /* ---------- стили и нумерация ---------- */
  function Styles(doc) {
    this.byId = {}; this.defP = { ind: {}, spacing: {} }; this.defR = {}; this.defaultPara = null; this.defaultTable = null;
    if (!doc) return;
    var dd = doc.getElementsByTagNameNS(W, 'docDefaults')[0];
    if (dd) {
      var rd = kid(kid(dd, 'rPrDefault'), 'rPr'), pd = kid(kid(dd, 'pPrDefault'), 'pPr');
      this.defR = readRPr(rd); this.defP = readPPr(pd);
    }
    var self = this;
    Array.prototype.forEach.call(doc.getElementsByTagNameNS(W, 'style'), function (s) {
      var id = wa(s, 'styleId'), type = wa(s, 'type');
      var tblPr = kid(s, 'tblPr');
      self.byId[id] = {
        type: type, basedOn: kid(s, 'basedOn') ? wa(kid(s, 'basedOn'), 'val') : null,
        pPr: readPPr(kid(s, 'pPr')), rPr: readRPr(kid(s, 'rPr')),
        borders: readBorders(tblPr && kid(tblPr, 'tblBorders')), cellMar: readMar(tblPr && kid(tblPr, 'tblCellMar'))
      };
      if (wa(s, 'default') === '1' && type === 'paragraph') self.defaultPara = id;
      if (wa(s, 'default') === '1' && type === 'table') self.defaultTable = id;
    });
  }
  Styles.prototype.chain = function (id) {
    var out = [], seen = {};
    while (id && this.byId[id] && !seen[id]) { seen[id] = 1; out.unshift(this.byId[id]); id = this.byId[id].basedOn; }
    return out;
  };
  Styles.prototype.pStyle = function (id) {
    var ch = this.chain(id || this.defaultPara), p = { ind: {}, spacing: {} }, r = {};
    ch.forEach(function (s) { p = mergeP(p, s.pPr); r = mergeR(r, s.rPr); });
    return { p: p, r: r };
  };
  Styles.prototype.rStyle = function (id) {
    var r = {}; this.chain(id).forEach(function (s) { r = mergeR(r, s.rPr); }); return r;
  };
  Styles.prototype.tStyle = function (id) {
    var ch = this.chain(id || this.defaultTable), o = { borders: {}, cellMar: {}, p: { ind: {}, spacing: {} }, r: {} };
    ch.forEach(function (s) {
      o.borders = Object.assign({}, o.borders, s.borders); o.cellMar = Object.assign({}, o.cellMar, s.cellMar);
      o.p = mergeP(o.p, s.pPr); o.r = mergeR(o.r, s.rPr);
    });
    return o;
  };

  function Numbering(doc) {
    this.nums = {}; this.abs = {}; this.counters = {};
    if (!doc) return;
    var self = this;
    Array.prototype.forEach.call(doc.getElementsByTagNameNS(W, 'abstractNum'), function (a) {
      var lv = {};
      kids(a, 'lvl').forEach(function (l) {
        lv[+wa(l, 'ilvl')] = {
          start: kid(l, 'start') ? +wa(kid(l, 'start'), 'val') : 1,
          fmt: kid(l, 'numFmt') ? wa(kid(l, 'numFmt'), 'val') : 'decimal',
          text: kid(l, 'lvlText') ? wa(kid(l, 'lvlText'), 'val') : '%1.',
          pPr: readPPr(kid(l, 'pPr')), rPr: readRPr(kid(l, 'rPr'))
        };
      });
      self.abs[wa(a, 'abstractNumId')] = lv;
    });
    Array.prototype.forEach.call(doc.getElementsByTagNameNS(W, 'num'), function (n) {
      self.nums[wa(n, 'numId')] = wa(kid(n, 'abstractNumId'), 'val');
    });
  }
  Numbering.prototype.level = function (numPr) {
    if (!numPr || !numPr.numId || numPr.numId === '0') return null;
    var a = this.abs[this.nums[numPr.numId]]; return a ? a[numPr.ilvl || 0] || null : null;
  };
  Numbering.prototype.next = function (numPr) {
    var lv = this.level(numPr); if (!lv) return '';
    var key = numPr.numId + ':' + (numPr.ilvl || 0);
    this.counters[key] = (this.counters[key] || (lv.start - 1)) + 1;
    var n = this.counters[key];
    if (lv.fmt === 'bullet') return '•';
    if (lv.fmt === 'none') return '';
    return lv.text.replace(/%\d/g, String(n));
  };

  /* ---------- метрики шрифта ---------- */
  function fontName(rp) { return rp.b ? 'serifBold' : 'serif'; }
  function metrics(rp) {
    var f = P.font(fontName(rp)), size = (rp.sz || 22) / 2;
    return { f: f, size: size, a: f.asc / f.upm * size, d: -f.desc / f.upm * size, g: f.gap / f.upm * size };
  }

  /* ---------- абзац ---------- */
  async function layoutPara(p, width, ctx, tstyle) {
    var st = ctx.styles;
    var direct = readPPr(kid(p, 'pPr'));
    var ps = st.pStyle(direct.style);
    var base = mergeP(st.defP, tstyle ? tstyle.p : { ind: {}, spacing: {} });
    var pp = mergeP(mergeP(base, ps.p), direct);
    var lvl = ctx.numbering.level(pp.numPr);
    if (lvl) pp = mergeP(mergeP(mergeP(base, ps.p), lvl.pPr), direct);
    var baseR = mergeR(mergeR(st.defR, tstyle ? tstyle.r : {}), ps.r);
    var markR = mergeR(baseR, readRPr(kid(kid(p, 'pPr'), 'rPr')));

    var L = (pp.ind.left || 0) * TW, Rt = (pp.ind.right || 0) * TW;
    var F = ((pp.ind.firstLine || 0) - (pp.ind.hanging || 0)) * TW;

    // токены
    var tokens = [], anchors = [];
    var cur = null;
    function pushChar(ch, rp) {
      if (ch === ' ') {
        cur = null;
        var last = tokens[tokens.length - 1];
        var w = P.textWidth(P.font(fontName(rp)), ' ', (rp.sz || 22) / 2, (rp.cs || 0) * TW);
        if (last && last.type === 'space') { last.w += w; last.n++; last.rp = rp; }
        else tokens.push({ type: 'space', w: w, n: 1, rp: rp });
        return;
      }
      if (!cur) { cur = { type: 'word', frags: [], w: 0 }; tokens.push(cur); }
      var fr = cur.frags[cur.frags.length - 1];
      if (!fr || fr.rp !== rp) { fr = { text: '', rp: rp }; cur.frags.push(fr); }
      fr.text += ch;
    }
    async function drawingToken(dr) {
      var inline = kid(dr, 'inline', WP), anchor = kid(dr, 'anchor', WP);
      var node = inline || anchor; if (!node) return;
      var ext = kid(node, 'extent', WP);
      var w = num(ext.getAttribute('cx')) * EMU, h = num(ext.getAttribute('cy')) * EMU;
      var blip = node.getElementsByTagNameNS(A, 'blip')[0];
      var rid = blip && (blip.getAttributeNS(R, 'embed'));
      var img = rid ? await ctx.image(rid) : null;
      if (!img) return;
      if (inline) { cur = null; tokens.push({ type: 'img', w: w, h: h, img: img }); }
      else {
        var ph = kid(node, 'positionH', WP), pv = kid(node, 'positionV', WP);
        var ox = ph && kid(ph, 'posOffset', WP) ? num(kid(ph, 'posOffset', WP).textContent) * EMU : 0;
        var oy = pv && kid(pv, 'posOffset', WP) ? num(kid(pv, 'posOffset', WP).textContent) * EMU : 0;
        anchors.push({ img: img, w: w, h: h, ox: ox, oy: oy,
          rh: ph ? ph.getAttribute('relativeFrom') : 'column', rv: pv ? pv.getAttribute('relativeFrom') : 'paragraph' });
      }
    }
    async function walkRuns(node) {
      for (var c = node.firstChild; c; c = c.nextSibling) {
        if (c.nodeType !== 1 || c.namespaceURI !== W) continue;
        var ln = c.localName;
        if (ln === 'r') {
          var dr = readRPr(kid(c, 'rPr'));
          var rp = mergeR(mergeR(baseR, dr.style ? st.rStyle(dr.style) : {}), dr);
          if (rp.hidden) continue;
          for (var e = c.firstChild; e; e = e.nextSibling) {
            if (e.nodeType !== 1) continue;
            var en = e.localName;
            if (en === 't') { var s = e.textContent; for (var i = 0; i < s.length; i++) pushChar(s[i], rp); }
            else if (en === 'tab') { cur = null; tokens.push({ type: 'tab', rp: rp, w: 0 }); }
            else if (en === 'br' || en === 'cr') { cur = null; tokens.push({ type: 'br' }); }
            else if (en === 'noBreakHyphen') pushChar('-', rp);
            else if (en === 'drawing') await drawingToken(e);
          }
        } else if (ln === 'hyperlink' || ln === 'smartTag' || ln === 'ins' || ln === 'customXml' || ln === 'sdt' || ln === 'sdtContent' || ln === 'fldSimple') {
          await walkRuns(c);
        }
      }
    }
    await walkRuns(p);
    tokens.forEach(function (t) {
      if (t.type !== 'word') return;
      t.w = 0;
      t.frags.forEach(function (fr) {
        fr.w = P.textWidth(P.font(fontName(fr.rp)), fr.text, (fr.rp.sz || 22) / 2, (fr.rp.cs || 0) * TW);
        t.w += fr.w;
      });
    });

    // нумерация
    var label = lvl ? ctx.numbering.next(pp.numPr) : '';
    var labelR = lvl ? mergeR(markR, lvl.rPr) : null;

    // разбиение на строки
    var lines = [];
    var line = null;
    function newLine() {
      var first = lines.length === 0;
      line = { items: [], start: first ? L + F : L, x: 0, first: first };
      line.x = line.start;
      if (first && label) {
        var lw = P.textWidth(P.font(fontName(labelR)), label, (labelR.sz || 22) / 2, 0);
        line.items.push({ type: 'word', frags: [{ text: label, rp: labelR, w: lw }], w: lw, x: line.x, label: true });
        var endL = line.x + lw, stop = L > endL + 0.01 ? L : Math.ceil((endL + 0.01) / (ctx.tabStop)) * ctx.tabStop;
        line.x = stop;
      }
      lines.push(line);
    }
    function nextTab(x) {
      var stops = (pp.tabs || []).map(function (t) { return t * TW; }).sort(function (a, b) { return a - b; });
      for (var i = 0; i < stops.length; i++) if (stops[i] > x + 0.01) return stops[i];
      return Math.ceil((x + 0.01) / ctx.tabStop) * ctx.tabStop;
    }
    var right = width - Rt;
    newLine();
    tokens.forEach(function (t) {
      if (t.type === 'br') { line.br = true; newLine(); return; }
      if (t.type === 'space') { line.items.push({ type: 'space', w: t.w, x: line.x, rp: t.rp }); line.x += t.w; return; }
      if (t.type === 'tab') { var nx = nextTab(line.x); line.items.push({ type: 'tab', w: nx - line.x, x: line.x, rp: t.rp }); line.x = nx; return; }
      var hasContent = line.items.some(function (i) { return i.type === 'word' || i.type === 'img'; });
      // Word 2013+ при выравнивании по ширине может сжать пробелы (примерно до 75%)
      var shrink = 0;
      if (pp.jc === 'both' && hasContent) {
        var seen = false;
        line.items.forEach(function (it) { if (it.type === 'word' || it.type === 'img') seen = true; else if (it.type === 'space' && seen) shrink += it.w * 0.25; });
      }
      if (line.x + t.w - shrink > right + 0.01 && hasContent) { newLine(); hasContent = false; }
      if (t.type === 'word' && !hasContent && line.x + t.w > right + 0.01) {
        // слово длиннее строки — режем по буквам
        t.frags.forEach(function (fr) {
          for (var i = 0; i < fr.text.length; i++) {
            var ch = fr.text[i], cw = P.textWidth(P.font(fontName(fr.rp)), ch, (fr.rp.sz || 22) / 2, (fr.rp.cs || 0) * TW);
            if (line.x + cw > right + 0.01 && line.items.length) newLine();
            line.items.push({ type: 'word', frags: [{ text: ch, rp: fr.rp, w: cw }], w: cw, x: line.x });
            line.x += cw;
          }
        });
        return;
      }
      line.items.push({ type: t.type, frags: t.frags, w: t.w, h: t.h, img: t.img, x: line.x });
      line.x += t.w;
    });

    // выравнивание и высота строк
    var sp = pp.spacing, rule = sp.lineRule || 'auto', lineVal = def(sp.line) ? sp.line : 240;
    lines.forEach(function (ln, idx) {
      var content = ln.items.filter(function (i) { return i.type === 'word' || i.type === 'img'; });
      var lastC = content[content.length - 1];
      var end = lastC ? lastC.x + lastC.w : ln.start;
      var free = right - end;
      var isLast = idx === lines.length - 1;
      var shift = 0;
      if (pp.jc === 'center') shift = free / 2;
      else if (pp.jc === 'right' || pp.jc === 'end') shift = free;
      else if (pp.jc === 'both' && ((!isLast && !ln.br) || free < 0)) {
        var firstIdx = ln.items.indexOf(content[0]), lastIdx = ln.items.indexOf(lastC);
        var spaces = ln.items.filter(function (it, k) { return it.type === 'space' && k > firstIdx && k < lastIdx; });
        var add = spaces.length ? free / spaces.length : 0, acc = 0;
        ln.items.forEach(function (it, k) {
          it.x += acc;
          if (it.type === 'space' && k > firstIdx && k < lastIdx) acc += add;
        });
      }
      if (shift) ln.items.forEach(function (it) { if (!it.label || pp.jc !== 'both') it.x += shift; });
      // высота
      var maxA = 0, maxD = 0, maxNat = 0, imgH = 0;
      var rps = [];
      ln.items.forEach(function (it) {
        if (it.type === 'word') it.frags.forEach(function (fr) { rps.push(fr.rp); });
        else if (it.type === 'img') imgH = Math.max(imgH, it.h);
        else if (it.rp && !content.length) rps.push(it.rp);
      });
      if (!rps.length) rps.push(markR);
      rps.forEach(function (rp) {
        var m = metrics(rp);
        maxA = Math.max(maxA, m.a + m.g); maxD = Math.max(maxD, m.d); maxNat = Math.max(maxNat, m.a + m.d + m.g);
      });
      var h;
      if (rule === 'exact') h = Math.abs(lineVal) * TW;
      else if (rule === 'atLeast') h = Math.max(maxNat, lineVal * TW);
      else h = maxNat * lineVal / 240;
      if (imgH) h = Math.max(h, imgH + maxD);
      ln.h = h; ln.baseline = h - maxD;
    });
    return {
      before: (sp.before || 0) * TW, after: (sp.after || 0) * TW, lines: lines, anchors: anchors,
      height: (sp.before || 0) * TW + (sp.after || 0) * TW + lines.reduce(function (a, l) { return a + l.h; }, 0)
    };
  }

  function drawLine(doc, page, ln, x, y) {
    ln.items.forEach(function (it) {
      if (it.type === 'word') {
        var xx = x + it.x;
        it.frags.forEach(function (fr) {
          var size = (fr.rp.sz || 22) / 2;
          doc.text(page, fontName(fr.rp), size, xx, y + ln.baseline, fr.text, (fr.rp.cs || 0) * TW);
          xx += fr.w;
        });
      } else if (it.type === 'img') {
        doc.image(page, it.img, x + it.x, y + ln.baseline - it.h, it.w, it.h);
      }
    });
  }

  /* ---------- таблица ---------- */
  async function layoutTable(tbl, colLeft, ctx) {
    var tp = kid(tbl, 'tblPr');
    var ts = ctx.styles.tStyle(tp && kid(tp, 'tblStyle') ? wa(kid(tp, 'tblStyle'), 'val') : null);
    var borders = Object.assign({}, ts.borders, readBorders(tp && kid(tp, 'tblBorders')));
    var cellMar = Object.assign({ left: 108, right: 108, top: 0, bottom: 0 }, ts.cellMar, readMar(tp && kid(tp, 'tblCellMar')));
    var ind = tp && kid(tp, 'tblInd') ? num(wa(kid(tp, 'tblInd'), 'w')) * TW : 0;
    var grid = kids(kid(tbl, 'tblGrid'), 'gridCol').map(function (g) { return num(wa(g, 'w')) * TW; });
    var x0 = colLeft + ind;
    var trs = kids(tbl, 'tr');
    var rows = [];
    for (var ri = 0; ri < trs.length; ri++) {
      var tr = trs[ri], trPr = kid(tr, 'trPr');
      var trH = trPr && kid(trPr, 'trHeight') ? num(wa(kid(trPr, 'trHeight'), 'val')) * TW : 0;
      var hRule = trPr && kid(trPr, 'trHeight') ? wa(kid(trPr, 'trHeight'), 'hRule') : null;
      var col = 0, cells = [];
      var tcs = kids(tr, 'tc');
      for (var ci = 0; ci < tcs.length; ci++) {
        var tc = tcs[ci], tcPr = kid(tc, 'tcPr');
        var span = tcPr && kid(tcPr, 'gridSpan') ? +wa(kid(tcPr, 'gridSpan'), 'val') : 1;
        var cx = x0; for (var g = 0; g < col; g++) cx += grid[g] || 0;
        var cw = 0; for (g = col; g < col + span; g++) cw += grid[g] || 0;
        var mar = Object.assign({}, cellMar, readMar(tcPr && kid(tcPr, 'tcMar')));
        var cb = readBorders(tcPr && kid(tcPr, 'tcBorders'));
        var firstCol = col === 0, lastCol = col + span >= grid.length;
        var eb = {
          top: 'top' in cb ? cb.top : (ri === 0 ? borders.top : borders.insideH),
          bottom: 'bottom' in cb ? cb.bottom : (ri === trs.length - 1 ? borders.bottom : borders.insideH),
          left: 'left' in cb ? cb.left : (firstCol ? borders.left : borders.insideV),
          right: 'right' in cb ? cb.right : (lastCol ? borders.right : borders.insideV)
        };
        var innerW = cw - (mar.left + mar.right) * TW;
        var paras = [];
        var ps = kids(tc, 'p');
        for (var pi = 0; pi < ps.length; pi++) paras.push(await layoutPara(ps[pi], innerW, ctx, ts));
        var contentH = paras.reduce(function (a, p) { return a + p.height; }, 0);
        var va = tcPr && kid(tcPr, 'vAlign') ? wa(kid(tcPr, 'vAlign'), 'val') : 'top';
        cells.push({ x: cx, w: cw, mar: mar, paras: paras, contentH: contentH, va: va, borders: eb });
        col += span;
      }
      var need = cells.reduce(function (m, c) { return Math.max(m, c.contentH + (c.mar.top + c.mar.bottom) * TW); }, 0);
      var h = hRule === 'exact' ? trH : Math.max(trH, need);
      rows.push({ h: h, cells: cells });
    }
    return rows;
  }
  function drawRow(doc, page, row, y) {
    row.cells.forEach(function (c) {
      var inner = row.h - (c.mar.top + c.mar.bottom) * TW;
      var dy = c.va === 'center' ? (inner - c.contentH) / 2 : c.va === 'bottom' ? inner - c.contentH : 0;
      var yy = y + c.mar.top * TW + Math.max(0, dy);
      var cx = c.x + c.mar.left * TW;
      c.paras.forEach(function (p) {
        yy += p.before;
        var top = yy;
        p.lines.forEach(function (ln) { drawLine(doc, page, ln, cx, yy); yy += ln.h; });
        p.anchors.forEach(function (a) { drawAnchor(doc, page, a, cx, top, null); });
        yy += p.after;
      });
      var b = c.borders;
      if (b.top) doc.line(page, c.x, y, c.x + c.w, y, b.top.sz / 8);
      if (b.bottom) doc.line(page, c.x, y + row.h, c.x + c.w, y + row.h, b.bottom.sz / 8);
      if (b.left) doc.line(page, c.x, y, c.x, y + row.h, b.left.sz / 8);
      if (b.right) doc.line(page, c.x + c.w, y, c.x + c.w, y + row.h, b.right.sz / 8);
    });
  }
  function drawAnchor(doc, page, a, colLeft, paraTop, geo) {
    var x = a.rh === 'page' ? a.ox : (a.rh === 'margin' && geo ? geo.left + a.ox : colLeft + a.ox);
    var y = a.rv === 'page' ? a.oy : (a.rv === 'margin' && geo ? geo.top + a.oy : paraTop + a.oy);
    doc.image(page, a.img, x, y, a.w, a.h);
  }

  /* ---------- документ целиком ---------- */
  async function render(zip) {
    var docXml = parse(await zip.file('word/document.xml').async('string'));
    var stylesF = zip.file('word/styles.xml'), numF = zip.file('word/numbering.xml'), setF = zip.file('word/settings.xml');
    var styles = new Styles(stylesF ? parse(await stylesF.async('string')) : null);
    var numbering = new Numbering(numF ? parse(await numF.async('string')) : null);
    var tabStop = 708 * TW;
    if (setF) {
      var sd = parse(await setF.async('string')), dts = sd.getElementsByTagNameNS(W, 'defaultTabStop')[0];
      if (dts) tabStop = num(wa(dts, 'val')) * TW;
    }
    var imgCache = {};
    async function relMap(part) {
      var rp = part.replace(/([^/]+)$/, '_rels/$1.rels'), f = zip.file(rp), m = {};
      if (!f) return m;
      var rd = parse(await f.async('string'));
      Array.prototype.forEach.call(rd.getElementsByTagNameNS(PR, 'Relationship'), function (r) {
        m[r.getAttribute('Id')] = { target: r.getAttribute('Target'), ext: r.getAttribute('TargetMode') === 'External' };
      });
      return m;
    }
    function makeCtx(rels) {
      return {
        styles: styles, numbering: numbering, tabStop: tabStop,
        image: async function (rid) {
          var rel = rels[rid]; if (!rel || rel.ext) return null;
          var path = 'word/' + rel.target.replace(/^\/?word\//, '').replace(/^\.\//, '');
          if (imgCache[path]) return imgCache[path];
          var f = zip.file(path); if (!f) return null;
          var bytes = await f.async('uint8array');
          var mime = /\.jpe?g$/i.test(path) ? 'image/jpeg' : 'image/png';
          imgCache[path] = await P.toJpeg(bytes, mime, 1000);
          return imgCache[path];
        }
      };
    }
    var ctx = makeCtx(await relMap('word/document.xml'));
    var body = docXml.getElementsByTagNameNS(W, 'body')[0];
    var sect = kid(body, 'sectPr');
    var pgSz = kid(sect, 'pgSz'), pgMar = kid(sect, 'pgMar');
    var PW = num(wa(pgSz, 'w')) * TW, PH = num(wa(pgSz, 'h')) * TW;
    var M = {}; ['top', 'right', 'bottom', 'left', 'header', 'footer'].forEach(function (k) { M[k] = Math.abs(num(wa(pgMar, k)) || 0) * TW; });
    var bodyW = PW - M.left - M.right;

    // колонтитулы (только основной)
    async function hf(kind) {
      var ref = kids(sect, kind + 'Reference').filter(function (r) { return wa(r, 'type') === 'default'; })[0];
      if (!ref) return null;
      var rels = await relMap('word/document.xml');
      var rel = rels[ref.getAttributeNS(R, 'id')]; if (!rel) return null;
      var path = 'word/' + rel.target;
      var f = zip.file(path); if (!f) return null;
      var d = parse(await f.async('string'));
      var hctx = makeCtx(await relMap(path));
      var paras = [];
      for (var p of kids(d.documentElement, 'p')) paras.push(await layoutPara(p, bodyW, hctx, null));
      return { paras: paras, h: paras.reduce(function (a, p) { return a + p.height; }, 0) };
    }
    var header = await hf('header'), footer = await hf('footer');
    var top = Math.max(M.top, header ? M.header + header.h : 0);
    var bottom = PH - Math.max(M.bottom, footer ? M.footer + footer.h : 0);
    var geo = { left: M.left, top: top };

    var doc = new P.Doc();
    var page = null, y = 0;
    function drawHF(hfo, y0) {
      if (!hfo) return;
      var yy = y0;
      hfo.paras.forEach(function (p) { yy += p.before; var t = yy; p.lines.forEach(function (ln) { drawLine(doc, page, ln, M.left, yy); yy += ln.h; }); p.anchors.forEach(function (a) { drawAnchor(doc, page, a, M.left, t, geo); }); yy += p.after; });
    }
    function newPage() {
      page = doc.addPage(PW, PH); y = top;
      drawHF(header, M.header);
      if (footer) drawHF(footer, PH - M.footer - footer.h);
    }
    function ensure(h) { if (y + h > bottom + 0.01 && y > top + 0.01) newPage(); }
    newPage();

    var children = kids(body);
    for (var i = 0; i < children.length; i++) {
      var el = children[i];
      if (el.localName === 'p') {
        var pl = await layoutPara(el, bodyW, ctx, null);
        y += pl.before;
        var paraTop = y, paraPage = page;
        pl.lines.forEach(function (ln, k) {
          ensure(ln.h);
          if (k === 0) { paraTop = y; paraPage = page; }
          drawLine(doc, page, ln, M.left, y); y += ln.h;
        });
        pl.anchors.forEach(function (a) { drawAnchor(doc, paraPage, a, M.left, paraTop, geo); });
        y += pl.after;
      } else if (el.localName === 'tbl') {
        var rows = await layoutTable(el, M.left, ctx);
        rows.forEach(function (row) { ensure(row.h); drawRow(doc, page, row, y); y += row.h; });
      }
    }
    return doc.save();
  }

  root.SBDocxPdf = { render: render };
})(window);
