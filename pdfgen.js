/* Минимальная сборка PDF внутри устройства: шрифты TrueType (Identity-H), картинки JPEG, линии.
   Ничего не загружает из сети. */
(function (root) {
  'use strict';

  /* ---------- TrueType ---------- */
  function b64ToBytes(b64) {
    var bin = atob(b64), u8 = new Uint8Array(bin.length);
    for (var i = 0; i < bin.length; i++) u8[i] = bin.charCodeAt(i);
    return u8;
  }
  function parseTTF(u8) {
    var dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
    var n = dv.getUint16(4), t = {};
    for (var i = 0; i < n; i++) {
      var o = 12 + i * 16;
      var tag = String.fromCharCode(u8[o], u8[o + 1], u8[o + 2], u8[o + 3]);
      t[tag] = { off: dv.getUint32(o + 8), len: dv.getUint32(o + 12) };
    }
    var head = t.head.off, hhea = t.hhea.off;
    var f = {
      bytes: u8,
      upm: dv.getUint16(head + 18),
      bbox: [dv.getInt16(head + 36), dv.getInt16(head + 38), dv.getInt16(head + 40), dv.getInt16(head + 42)],
      asc: dv.getInt16(hhea + 4), desc: dv.getInt16(hhea + 6), gap: dv.getInt16(hhea + 8),
      nH: dv.getUint16(hhea + 34), numGlyphs: dv.getUint16(t.maxp.off + 4)
    };
    f.capHeight = f.asc * 0.66;
    if (t['OS/2'] && dv.getUint16(t['OS/2'].off) >= 2) f.capHeight = dv.getInt16(t['OS/2'].off + 88);
    var adv = new Array(f.numGlyphs), last = 0;
    for (var g = 0; g < f.numGlyphs; g++) {
      if (g < f.nH) last = dv.getUint16(t.hmtx.off + g * 4);
      adv[g] = last;
    }
    f.adv = adv;
    // cmap format 4 (3,1)
    var cm = t.cmap.off, nsub = dv.getUint16(cm + 2), sub = -1;
    for (var s = 0; s < nsub; s++) {
      var pid = dv.getUint16(cm + 4 + s * 8), eid = dv.getUint16(cm + 6 + s * 8), so = dv.getUint32(cm + 8 + s * 8);
      if (pid === 3 && eid === 1 && dv.getUint16(cm + so) === 4) sub = cm + so;
    }
    var map = {};
    if (sub >= 0) {
      var segX2 = dv.getUint16(sub + 6), segs = segX2 / 2;
      var ends = sub + 14, starts = ends + segX2 + 2, deltas = starts + segX2, ros = deltas + segX2;
      for (var k = 0; k < segs; k++) {
        var e = dv.getUint16(ends + k * 2), st = dv.getUint16(starts + k * 2);
        var d = dv.getInt16(deltas + k * 2), ro = dv.getUint16(ros + k * 2);
        for (var c = st; c <= e && c !== 0xFFFF; c++) {
          var gid;
          if (ro === 0) gid = (c + d) & 0xFFFF;
          else {
            var gi = dv.getUint16(ros + k * 2 + ro + (c - st) * 2);
            gid = gi ? (gi + d) & 0xFFFF : 0;
          }
          if (gid) map[c] = gid;
        }
      }
    }
    f.map = map;
    return f;
  }
  var FONT_CACHE = {};
  function font(name) {
    if (!FONT_CACHE[name]) {
      var f = parseTTF(b64ToBytes(root.SB_FONTS[name]));
      f.name = name;
      FONT_CACHE[name] = f;
    }
    return FONT_CACHE[name];
  }
  function glyph(f, ch) {
    var c = ch.charCodeAt(0);
    if (c === 0xA0) c = 32;
    return f.map[c] || f.map[0x3F] || 0;
  }
  // ширина строки в пунктах
  function textWidth(f, s, size, charSpacing) {
    var w = 0;
    for (var i = 0; i < s.length; i++) w += f.adv[glyph(f, s[i])] * size / f.upm + (charSpacing || 0);
    return w;
  }

  /* ---------- сжатие ---------- */
  async function deflate(u8) {
    if (typeof CompressionStream === 'undefined') return null;
    var cs = new CompressionStream('deflate');
    var stream = new Blob([u8]).stream().pipeThrough(cs);
    return new Uint8Array(await new Response(stream).arrayBuffer());
  }

  /* ---------- документ ---------- */
  function num(x) { var s = (Math.round(x * 1000) / 1000).toString(); return s === '-0' ? '0' : s; }
  function hex4(n) { var h = n.toString(16).toUpperCase(); while (h.length < 4) h = '0' + h; return h; }

  function Doc() {
    this.pages = [];           // {w,h,ops:[]}
    this.fontsUsed = {};       // name -> {f, glyphs:{gid:unicode}}
    this.images = [];          // {jpeg,w,h}
  }
  Doc.prototype.addPage = function (w, h) { var p = { w: w, h: h, ops: [] }; this.pages.push(p); return p; };
  Doc.prototype.useFont = function (name) {
    if (!this.fontsUsed[name]) this.fontsUsed[name] = { f: font(name), glyphs: {}, id: 'F' + (Object.keys(this.fontsUsed).length + 1) };
    return this.fontsUsed[name];
  };
  // y — положение базовой линии от верха страницы
  Doc.prototype.text = function (page, fontName, size, x, y, s, charSpacing) {
    if (!s) return;
    var fu = this.useFont(fontName), f = fu.f, hex = '';
    for (var i = 0; i < s.length; i++) {
      var g = glyph(f, s[i]);
      fu.glyphs[g] = s.charCodeAt(i) === 0xA0 ? 32 : s.charCodeAt(i);
      hex += hex4(g);
    }
    page.ops.push('BT /' + fu.id + ' ' + num(size) + ' Tf ' + num(charSpacing || 0) + ' Tc ' + num(x) + ' ' + num(page.h - y) + ' Td <' + hex + '> Tj ET');
  };
  Doc.prototype.line = function (page, x1, y1, x2, y2, width) {
    page.ops.push(num(width) + ' w 0 G ' + num(x1) + ' ' + num(page.h - y1) + ' m ' + num(x2) + ' ' + num(page.h - y2) + ' l S');
  };
  Doc.prototype.image = function (page, img, x, y, w, h) {
    var idx = this.images.indexOf(img);
    if (idx < 0) { this.images.push(img); idx = this.images.length - 1; }
    page.ops.push('q ' + num(w) + ' 0 0 ' + num(h) + ' ' + num(x) + ' ' + num(page.h - y - h) + ' cm /Im' + (idx + 1) + ' Do Q');
  };

  Doc.prototype.save = async function () {
    var enc = new TextEncoder();
    var chunks = [], offsets = [], pos = 0, objs = 0;
    function push(u8) { chunks.push(u8); pos += u8.length; }
    function str(s) { push(enc.encode(s)); }
    function obj(n, body) { offsets[n] = pos; str(n + ' 0 obj\n' + body + '\nendobj\n'); }
    async function stream(n, dict, data, noCompress) {
      var u8 = typeof data === 'string' ? enc.encode(data) : data;
      var z = noCompress ? null : await deflate(u8);
      offsets[n] = pos;
      str(n + ' 0 obj\n<<' + dict + (z ? ' /Filter /FlateDecode' : '') + ' /Length ' + (z || u8).length + '>>\nstream\n');
      push(z || u8);
      str('\nendstream\nendobj\n');
    }
    function alloc() { return ++objs; }

    push(new Uint8Array([0x25, 0x50, 0x44, 0x46, 0x2D, 0x31, 0x2E, 0x37, 0x0A, 0x25, 0xE2, 0xE3, 0xCF, 0xD3, 0x0A])); // %PDF-1.7 + бинарная метка

    var catalog = alloc(), pagesObj = alloc();
    var fontRefs = {}, self = this;
    // шрифты
    for (var name in this.fontsUsed) {
      var fu = this.fontsUsed[name], f = fu.f;
      var type0 = alloc(), cid = alloc(), desc = alloc(), file = alloc(), tou = alloc();
      fontRefs[fu.id] = type0;
      var sc = 1000 / f.upm;
      var gids = Object.keys(fu.glyphs).map(Number).sort(function (a, b) { return a - b; });
      var W = gids.map(function (g) { return g + ' [' + Math.round(f.adv[g] * sc) + ']'; }).join(' ');
      var baseName = (name === 'serifBold' ? 'AAAAAB+LiberationSerif-Bold' : 'AAAAAA+LiberationSerif');
      obj(type0, '<< /Type /Font /Subtype /Type0 /BaseFont /' + baseName + ' /Encoding /Identity-H /DescendantFonts [' + cid + ' 0 R] /ToUnicode ' + tou + ' 0 R >>');
      obj(cid, '<< /Type /Font /Subtype /CIDFontType2 /BaseFont /' + baseName + ' /CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >> /FontDescriptor ' + desc + ' 0 R /CIDToGIDMap /Identity /DW 500 /W [' + W + '] >>');
      obj(desc, '<< /Type /FontDescriptor /FontName /' + baseName + ' /Flags 34 /FontBBox [' + f.bbox.map(function (v) { return Math.round(v * sc); }).join(' ') +
        '] /ItalicAngle 0 /Ascent ' + Math.round(f.asc * sc) + ' /Descent ' + Math.round(f.desc * sc) + ' /CapHeight ' + Math.round(f.capHeight * sc) +
        ' /StemV ' + (name === 'serifBold' ? 140 : 80) + ' /FontFile2 ' + file + ' 0 R >>');
      await stream(file, ' /Length1 ' + f.bytes.length, f.bytes);
      var lines = [];
      for (var i = 0; i < gids.length; i += 100) {
        var part = gids.slice(i, i + 100);
        lines.push(part.length + ' beginbfchar');
        part.forEach(function (g) { lines.push('<' + hex4(g) + '> <' + hex4(fu.glyphs[g]) + '>'); });
        lines.push('endbfchar');
      }
      await stream(tou, '', '/CIDInit /ProcSet findresource begin 12 dict begin begincmap /CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def /CMapName /Adobe-Identity-UCS def /CMapType 2 def 1 begincodespacerange <0000> <FFFF> endcodespacerange\n' + lines.join('\n') + '\nendcmap CMapName currentdict /CMap defineresource pop end end');
    }
    // картинки
    var imgRefs = [];
    for (var k = 0; k < this.images.length; k++) {
      var im = this.images[k], n = alloc(); imgRefs.push(n);
      await stream(n, ' /Type /XObject /Subtype /Image /Width ' + im.w + ' /Height ' + im.h + ' /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode', im.jpeg, true);
    }
    var fontDict = Object.keys(fontRefs).map(function (id) { return '/' + id + ' ' + fontRefs[id] + ' 0 R'; }).join(' ');
    var xobj = imgRefs.map(function (r, i) { return '/Im' + (i + 1) + ' ' + r + ' 0 R'; }).join(' ');
    var pageRefs = [];
    for (var p = 0; p < this.pages.length; p++) {
      var pg = this.pages[p], pn = alloc(), cn = alloc(); pageRefs.push(pn);
      await stream(cn, '', pg.ops.join('\n'));
      obj(pn, '<< /Type /Page /Parent ' + pagesObj + ' 0 R /MediaBox [0 0 ' + num(pg.w) + ' ' + num(pg.h) + '] /Resources << /Font << ' + fontDict + ' >> /XObject << ' + xobj + ' >> >> /Contents ' + cn + ' 0 R >>');
    }
    obj(pagesObj, '<< /Type /Pages /Kids [' + pageRefs.map(function (r) { return r + ' 0 R'; }).join(' ') + '] /Count ' + pageRefs.length + ' >>');
    obj(catalog, '<< /Type /Catalog /Pages ' + pagesObj + ' 0 R >>');
    var info = alloc();
    obj(info, '<< /Producer (Dostup SB offline) >>');
    var xref = pos;
    var x = 'xref\n0 ' + (objs + 1) + '\n0000000000 65535 f \n';
    for (var j = 1; j <= objs; j++) x += String(offsets[j]).padStart(10, '0') + ' 00000 n \n';
    str(x + 'trailer\n<< /Size ' + (objs + 1) + ' /Root ' + catalog + ' 0 R /Info ' + info + ' 0 R >>\nstartxref\n' + xref + '\n%%EOF\n');
    return new Blob(chunks, { type: 'application/pdf' });
  };

  /* ---------- картинка → JPEG (через canvas, внутри устройства) ---------- */
  async function toJpeg(bytes, mime, maxSide) {
    var blob = new Blob([bytes], { type: mime || 'image/png' });
    var bmp;
    if (typeof createImageBitmap === 'function') bmp = await createImageBitmap(blob);
    else {
      bmp = await new Promise(function (res, rej) {
        var url = URL.createObjectURL(blob), im = new Image();
        im.onload = function () { URL.revokeObjectURL(url); res(im); }; im.onerror = rej; im.src = url;
      });
    }
    var w = bmp.width, h = bmp.height, sc = Math.min(1, (maxSide || 1600) / Math.max(w, h));
    w = Math.max(1, Math.round(w * sc)); h = Math.max(1, Math.round(h * sc));
    var c = document.createElement('canvas'); c.width = w; c.height = h;
    var ctx = c.getContext('2d'); ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, w, h); ctx.drawImage(bmp, 0, 0, w, h);
    var jb = await new Promise(function (res) { c.toBlob(res, 'image/jpeg', 0.95); });
    return { jpeg: new Uint8Array(await jb.arrayBuffer()), w: w, h: h };
  }

  root.SBPDF = { Doc: Doc, font: font, textWidth: textWidth, glyph: glyph, toJpeg: toJpeg };
})(window);
