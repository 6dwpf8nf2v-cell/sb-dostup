/* Заполнение шаблонов Word и Excel. Шаблоны хранятся только на устройстве.
   Поля ищутся по тексту шаблона, поэтому подойдёт и обновлённый бланк того же вида. */
(function (root) {
  'use strict';
  var W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  var R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  var S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  var PR = 'http://schemas.openxmlformats.org/package/2006/relationships';
  var XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';

  function parse(s) { return new DOMParser().parseFromString(s, 'application/xml'); }
  function ser(d) { return new XMLSerializer().serializeToString(d); }
  function arr(nl) { return Array.prototype.slice.call(nl); }
  function kids(n, ns, name) { return arr(n.childNodes).filter(function (c) { return c.nodeType === 1 && c.localName === name && c.namespaceURI === ns; }); }
  function relsPath(part) { var i = part.lastIndexOf('/'); return part.slice(0, i + 1) + '_rels/' + part.slice(i + 1) + '.rels'; }
  function resolve(base, target) {
    if (target.charAt(0) === '/') return target.slice(1);
    var parts = base.split('/'); parts.pop();
    target.split('/').forEach(function (p) { if (p === '..') parts.pop(); else if (p !== '.') parts.push(p); });
    return parts.join('/');
  }

  /* ---------- удаление внешних картинок (счётчики открытия) ---------- */
  async function stripExternalImages(zip, isWord) {
    var removed = 0;
    var relFiles = Object.keys(zip.files).filter(function (p) { return /_rels\/[^/]+\.rels$/.test(p) && !/^_rels\//.test(p); });
    for (var i = 0; i < relFiles.length; i++) {
      var rp = relFiles[i];
      var rd = parse(await zip.file(rp).async('string'));
      var ext = arr(rd.getElementsByTagNameNS(PR, 'Relationship')).filter(function (r) {
        return r.getAttribute('TargetMode') === 'External' && /\/image$/.test(r.getAttribute('Type') || '');
      });
      if (!ext.length) continue;
      var part = rp.replace(/_rels\/([^/]+)\.rels$/, '$1');
      var pf = zip.file(part); if (!pf) continue;
      var pd = parse(await pf.async('string'));
      ext.forEach(function (rel) {
        var id = rel.getAttribute('Id');
        var hits = arr(pd.getElementsByTagName('*')).filter(function (e) {
          return e.getAttributeNS(R, 'link') === id || e.getAttributeNS(R, 'embed') === id || e.getAttributeNS(R, 'id') === id;
        });
        hits.forEach(function (h) {
          var n = h;
          while (n && !(isWord ? (n.localName === 'drawing' || n.localName === 'pict') : /Anchor$/.test(n.localName))) n = n.parentNode;
          if (n && n.parentNode) { n.parentNode.removeChild(n); removed++; }
        });
        rel.parentNode.removeChild(rel);
      });
      zip.file(part, ser(pd));
      zip.file(rp, ser(rd));
    }
    return removed;
  }

  /* ---------- Word ---------- */
  function paraSegs(p) {
    var segs = [];
    (function walk(n) {
      for (var c = n.firstChild; c; c = c.nextSibling) {
        if (c.nodeType !== 1) continue;
        if (c.namespaceURI === W && (c.localName === 'p' || c.localName === 'txbxContent' || c.localName === 'drawing')) continue;
        if (c.namespaceURI === W && c.localName === 't') { segs.push(c); continue; }
        walk(c);
      }
    })(p);
    return segs;
  }
  function paraText(p) { return paraSegs(p).map(function (t) { return t.textContent; }).join(''); }
  function setRange(segs, s, e, txt) {
    var pos = 0, placed = false;
    segs.forEach(function (t) {
      var v = t.textContent, a = pos, b = pos + v.length;
      pos = b;
      if (b <= s || a >= e) return;
      var from = Math.max(s, a) - a, to = Math.min(e, b) - a;
      var nv = v.slice(0, from) + (placed ? '' : txt) + v.slice(to);
      placed = true;
      t.textContent = nv;
      t.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
    });
    return placed;
  }
  // regex с тремя группами: (до)(что заменить)(после)
  function replaceFirst(paras, re, value) {
    for (var i = 0; i < paras.length; i++) {
      var segs = paraSegs(paras[i]);
      var full = segs.map(function (t) { return t.textContent; }).join('');
      var m = re.exec(full);
      if (!m) continue;
      var s = m.index + m[1].length, e = s + m[2].length;
      return setRange(segs, s, e, value);
    }
    return false;
  }
  function cellText(tc) { return kids(tc, W, 'p').map(paraText).join(' '); }
  function setCellText(tc, text) {
    var ps = kids(tc, W, 'p');
    var p = ps[0];
    ps.slice(1).forEach(function (x) { tc.removeChild(x); });
    var runs = arr(p.getElementsByTagNameNS(W, 'r')).filter(function (r) { return r.getElementsByTagNameNS(W, 't').length; });
    if (runs.length) {
      var ts = runs[0].getElementsByTagNameNS(W, 't');
      ts[0].textContent = text; ts[0].setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
      for (var k = ts.length - 1; k > 0; k--) ts[k].parentNode.removeChild(ts[k]);
      runs.slice(1).forEach(function (r) { r.parentNode.removeChild(r); });
    } else {
      var doc = p.ownerDocument, r = doc.createElementNS(W, 'w:r');
      var ppr = kids(p, W, 'pPr')[0], mark = ppr && kids(ppr, W, 'rPr')[0];
      if (mark) r.appendChild(mark.cloneNode(true));
      var t = doc.createElementNS(W, 'w:t'); t.textContent = text;
      t.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
      r.appendChild(t); p.appendChild(r);
    }
  }

  async function fillDocx(buf, d) {
    var zip = await JSZip.loadAsync(buf);
    var doc = parse(await zip.file('word/document.xml').async('string'));
    var body = doc.getElementsByTagNameNS(W, 'body')[0];
    var paras = arr(body.getElementsByTagNameNS(W, 'p'));
    var found = {};
    found.outNo = replaceFirst(paras, /(Исх\.?\s*:?\s*№?\s*)(\d+\/\d{2}\/\d{2})()/, d.outNo);
    found.date = replaceFirst(paras, /(^\s*от\s+)(\d{2}\.\d{2}\.\d{4})(\s*г)/, d.date);
    found.contract = replaceFirst(paras, /(договора\s+)(№\s*\d+\s+от\s+\d{2}\.\d{2}\.\d{2,4})(\s*г)/, d.contract);
    found.place = replaceFirst(paras, /(во\s+все\s+ВСП\s+)(.+?банка\s+ГОСБ\s*№\s*\d+\s+.+?)(\s+ПАО\s+«?Сбербанк)/, d.place);
    var pEnd = replaceFirst(paras, /(предоставить\s+с\s+\d{2}\.\d{2}\.\d{4}\s+по\s+)(\d{2}\.\d{2}\.\d{4})()/, d.periodEnd);
    var pStart = replaceFirst(paras, /(предоставить\s+с\s+)(\d{2}\.\d{2}\.\d{4})()/, d.periodStart);
    found.period = pEnd && pStart;

    // таблица сотрудников
    found.table = false;
    var tbl = arr(body.getElementsByTagNameNS(W, 'tbl')).filter(function (t) {
      var tx = arr(t.getElementsByTagNameNS(W, 't')).map(function (x) { return x.textContent; }).join('');
      return /Ф\.?\s*И\.?\s*О/.test(tx) && /паспорт/i.test(tx);
    })[0];
    if (tbl) {
      var rows = kids(tbl, W, 'tr');
      var head = rows[0], cols = {};
      kids(head, W, 'tc').forEach(function (tc, i) {
        var tx = cellText(tc);
        if (/Ф\.?\s*И\.?\s*О/.test(tx)) cols.fio = i;
        else if (/рожд/i.test(tx)) cols.birth = i;
        else if (/паспорт/i.test(tx)) cols.passport = i;
      });
      if (rows.length >= 2 && cols.fio != null && cols.birth != null && cols.passport != null) {
        var tpl = rows[1];
        rows.slice(1).forEach(function (r) { tbl.removeChild(r); });
        d.people.forEach(function (p) {
          var r = tpl.cloneNode(true), tcs = kids(r, W, 'tc');
          setCellText(tcs[cols.fio], p.fio);
          setCellText(tcs[cols.birth], p.birth);
          setCellText(tcs[cols.passport], p.passport);
          tbl.appendChild(r);
        });
        found.table = true;
      }
    }
    zip.file('word/document.xml', ser(doc));
    found.stripped = await stripExternalImages(zip, true);
    var blob = await zip.generateAsync({ type: 'blob', compression: 'DEFLATE', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
    return { blob: blob, zip: zip, found: found };
  }

  /* ---------- Excel ---------- */
  var MONTHS_GEN = ['января', 'февраля', 'марта', 'апреля', 'мая', 'июня', 'июля', 'августа', 'сентября', 'октября', 'ноября', 'декабря'];
  function colOf(ref) { return /^([A-Z]+)/.exec(ref)[1]; }
  function excelSerial(dt) { return Math.round((Date.UTC(dt.getFullYear(), dt.getMonth(), dt.getDate()) - Date.UTC(1899, 11, 30)) / 86400000); }

  async function fillXlsx(buf, d) {
    var zip = await JSZip.loadAsync(buf);
    var wb = parse(await zip.file('xl/workbook.xml').async('string'));
    var wbr = parse(await zip.file('xl/_rels/workbook.xml.rels').async('string'));
    var sheetEl = wb.getElementsByTagNameNS(S, 'sheet')[0];
    var rid = sheetEl.getAttributeNS(R, 'id');
    var target = arr(wbr.getElementsByTagNameNS(PR, 'Relationship')).filter(function (r) { return r.getAttribute('Id') === rid; })[0].getAttribute('Target');
    var sheetPath = resolve('xl/workbook.xml', target);
    var sh = parse(await zip.file(sheetPath).async('string'));

    // общие строки
    var sstPath = 'xl/sharedStrings.xml';
    var sst = zip.file(sstPath) ? parse(await zip.file(sstPath).async('string')) : null;
    var strings = sst ? arr(sst.getElementsByTagNameNS(S, 'si')).map(function (si) {
      return arr(si.getElementsByTagNameNS(S, 't')).map(function (t) { return t.textContent; }).join('');
    }) : [];
    function addString(s) {
      var i = strings.indexOf(s);
      if (i >= 0) return i;
      var si = sst.createElementNS(S, 'si'), t = sst.createElementNS(S, 't');
      t.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
      t.textContent = s; si.appendChild(t); sst.documentElement.appendChild(si);
      strings.push(s); return strings.length - 1;
    }
    function cellValue(c) {
      var v = kids(c, S, 'v')[0];
      if (c.getAttribute('t') === 's' && v) return strings[+v.textContent] || '';
      if (c.getAttribute('t') === 'inlineStr') return c.textContent;
      return v ? v.textContent : '';
    }

    // стили: убираем апостроф (quotePrefix) у ячеек паспорта
    var stylesPath = 'xl/styles.xml';
    var st = parse(await zip.file(stylesPath).async('string'));
    var cellXfs = st.getElementsByTagNameNS(S, 'cellXfs')[0];
    var xfs = kids(cellXfs, S, 'xf');
    var noQuoteCache = {};
    function noQuote(sIdx) {
      var xf = xfs[sIdx];
      if (!xf || xf.getAttribute('quotePrefix') !== '1') return sIdx;
      if (noQuoteCache[sIdx] != null) return noQuoteCache[sIdx];
      var cl = xf.cloneNode(true); cl.removeAttribute('quotePrefix');
      cellXfs.appendChild(cl); xfs.push(cl);
      cellXfs.setAttribute('count', String(xfs.length));
      noQuoteCache[sIdx] = xfs.length - 1;
      return noQuoteCache[sIdx];
    }

    var sheetData = sh.getElementsByTagNameNS(S, 'sheetData')[0];
    var rows = kids(sheetData, S, 'row');
    var found = {};
    // шапка: первая строка с «Фамилия»
    var headRow = rows.filter(function (r) { return kids(r, S, 'c').some(function (c) { return /фамил/i.test(cellValue(c)); }); })[0];
    var map = {};
    if (headRow) kids(headRow, S, 'c').forEach(function (c) {
      var v = cellValue(c).toLowerCase(), col = colOf(c.getAttribute('r'));
      if (/фамил/.test(v)) map[col] = 'surname';
      else if (/^\s*имя/.test(v)) map[col] = 'name';
      else if (/отчеств/.test(v)) map[col] = 'patr';
      else if (/рожд/.test(v)) map[col] = 'birth';
      else if (/паспорт.*сер/.test(v)) map[col] = 'series';
      else if (/паспорт.*ном/.test(v)) map[col] = 'number';
      else if (/договор/.test(v)) map[col] = 'contract';
      else if (/телефон/.test(v)) map[col] = 'phone';
    });
    var need = ['surname', 'name', 'patr', 'birth', 'series', 'number', 'phone', 'contract'];
    found.columns = need.every(function (k) { return Object.keys(map).some(function (c) { return map[c] === k; }); });
    var headR = headRow ? +headRow.getAttribute('r') : 1;
    var dataRow = rows.filter(function (r) { return +r.getAttribute('r') === headR + 1; })[0];
    var refRow = rows.filter(function (r) { return kids(r, S, 'c').some(function (c) { return /^\s*№\s*\d+\/\d{2}\/\d{2}/.test(cellValue(c)); }); })[0];
    found.ref = !!refRow;
    found.dataRow = !!dataRow;
    if (!dataRow || !found.columns) {
      return { blob: null, found: found };
    }
    var dataR = +dataRow.getAttribute('r'), n = d.people.length, shift = n - 1;

    function setRowNum(row, r) {
      row.setAttribute('r', String(r));
      kids(row, S, 'c').forEach(function (c) { c.setAttribute('r', colOf(c.getAttribute('r')) + r); });
    }
    function setStr(c, s) {
      kids(c, S, 'v').concat(kids(c, S, 'is'), kids(c, S, 'f')).forEach(function (x) { c.removeChild(x); });
      c.setAttribute('t', 's');
      var v = sh.createElementNS(S, 'v'); v.textContent = String(addString(s)); c.appendChild(v);
    }
    function setNum(c, num) {
      kids(c, S, 'v').concat(kids(c, S, 'is'), kids(c, S, 'f')).forEach(function (x) { c.removeChild(x); });
      c.removeAttribute('t');
      var v = sh.createElementNS(S, 'v'); v.textContent = String(num); c.appendChild(v);
    }

    // сдвигаем строки ниже шаблонной
    rows.forEach(function (r) { var rn = +r.getAttribute('r'); if (rn > dataR) setRowNum(r, rn + shift); });
    // строки людей
    var anchor = dataRow.nextSibling;
    sheetData.removeChild(dataRow);
    d.people.forEach(function (p, i) {
      var row = dataRow.cloneNode(true);
      setRowNum(row, dataR + i);
      kids(row, S, 'c').forEach(function (c) {
        var key = map[colOf(c.getAttribute('r'))];
        if (!key) return;
        if (key === 'birth') setNum(c, excelSerial(p.birthDate));
        else if (key === 'phone') setNum(c, p.phone);
        else if (key === 'series' || key === 'number') { setStr(c, p[key]); c.setAttribute('s', String(noQuote(+(c.getAttribute('s') || 0)))); }
        else if (key === 'contract') setStr(c, d.contractDigits);
        else setStr(c, p[key] || '');
      });
      sheetData.insertBefore(row, anchor);
    });
    // строка со ссылкой на письмо
    if (refRow) kids(refRow, S, 'c').forEach(function (c) {
      if (/^\s*№\s*\d+\/\d{2}\/\d{2}/.test(cellValue(c))) setStr(c, d.refText);
    });
    // размеры и проверки данных
    var dim = sh.getElementsByTagNameNS(S, 'dimension')[0];
    if (dim) {
      var last = kids(sheetData, S, 'row').reduce(function (m, r) { return Math.max(m, +r.getAttribute('r')); }, 1);
      dim.setAttribute('ref', dim.getAttribute('ref').replace(/(:[A-Z]+)\d+$/, '$1' + last));
    }
    arr(sh.getElementsByTagNameNS(S, 'dataValidation')).forEach(function (dv) {
      var sq = (dv.getAttribute('sqref') || '').split(/\s+/).map(function (ref) {
        var m = /^([A-Z]+)(\d+)$/.exec(ref);
        if (m && +m[2] === dataR) return n > 1 ? ref + ':' + m[1] + (dataR + n - 1) : ref;
        var m2 = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(ref);
        if (m2 && +m2[2] === dataR) return m2[1] + m2[2] + ':' + m2[3] + (Math.max(+m2[4], dataR) + shift);
        return ref;
      });
      dv.setAttribute('sqref', sq.join(' '));
    });
    // выделение — на первую ячейку
    arr(sh.getElementsByTagNameNS(S, 'selection')).forEach(function (s) { s.setAttribute('activeCell', 'A1'); s.setAttribute('sqref', 'A1'); });

    if (sst) {
      sst.documentElement.setAttribute('uniqueCount', String(strings.length));
      sst.documentElement.setAttribute('count', String(strings.length + 64));
      zip.file(sstPath, ser(sst));
    }
    zip.file(stylesPath, ser(st));
    zip.file(sheetPath, ser(sh));
    zip.remove('xl/calcChain.xml');
    found.stripped = await stripExternalImages(zip, false);
    // пустой рисунок после удаления счётчика — убираем ссылку на него
    await dropEmptyDrawings(zip, sheetPath);
    var blob = await zip.generateAsync({ type: 'blob', compression: 'DEFLATE', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    return { blob: blob, found: found };
  }

  async function dropEmptyDrawings(zip, sheetPath) {
    var rp = relsPath(sheetPath);
    if (!zip.file(rp)) return;
    var rd = parse(await zip.file(rp).async('string'));
    var sh = parse(await zip.file(sheetPath).async('string'));
    var changed = false;
    for (var rel of arr(rd.getElementsByTagNameNS(PR, 'Relationship'))) {
      if (!/\/drawing$/.test(rel.getAttribute('Type') || '')) continue;
      var dp = resolve(sheetPath, rel.getAttribute('Target'));
      var f = zip.file(dp); if (!f) continue;
      var dd = parse(await f.async('string'));
      var anchors = arr(dd.documentElement.childNodes).filter(function (c) { return c.nodeType === 1; });
      if (anchors.length) continue;
      var id = rel.getAttribute('Id');
      arr(sh.getElementsByTagNameNS(S, 'drawing')).forEach(function (e) { if (e.getAttributeNS(R, 'id') === id) e.parentNode.removeChild(e); });
      rel.parentNode.removeChild(rel);
      zip.remove(dp); zip.remove(relsPath(dp));
      var dir = dp.replace(/[^/]+$/, '');
      if (!Object.keys(zip.files).some(function (k) { return k.indexOf(dir) === 0 && !zip.files[k].dir; })) zip.remove(dir.replace(/\/$/, ''));
      var ct = parse(await zip.file('[Content_Types].xml').async('string'));
      arr(ct.getElementsByTagName('Override')).forEach(function (o) { if (o.getAttribute('PartName') === '/' + dp) o.parentNode.removeChild(o); });
      zip.file('[Content_Types].xml', ser(ct));
      changed = true;
    }
    if (changed) { zip.file(sheetPath, ser(sh)); zip.file(rp, ser(rd)); }
  }

  function refText(outNo, date) {
    return '№ ' + outNo + ', от «' + SBCore.pad2(date.getDate()) + '» ' + MONTHS_GEN[date.getMonth()] + ' ' + date.getFullYear() + ' г.';
  }

  root.SBFill = { fillDocx: fillDocx, fillXlsx: fillXlsx, refText: refText };
})(window);
