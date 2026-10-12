/* Чтение файлов (Excel, Word, PDF, текст) и сборка Excel — всё внутри устройства. */
(function (root) {
  'use strict';

  var NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  var NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  var NS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

  function xml(str) { return new DOMParser().parseFromString(str, 'application/xml'); }
  function kids(node, ns, name) {
    var out = [];
    for (var c = node.firstChild; c; c = c.nextSibling) if (c.nodeType === 1 && c.localName === name && (!ns || c.namespaceURI === ns)) out.push(c);
    return out;
  }
  function esc(s) {
    return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  /* ---------- Excel → текст ---------- */
  function colIndex(ref) {
    var m = /^([A-Z]+)/.exec(ref || ''); if (!m) return -1;
    var n = 0; for (var i = 0; i < m[1].length; i++) n = n * 26 + (m[1].charCodeAt(i) - 64);
    return n - 1;
  }
  function excelDate(serial) {
    var ms = Math.round((serial - 25569) * 86400000);
    var d = new Date(ms);
    return new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate());
  }

  async function readXlsx(buf) {
    var zip = await JSZip.loadAsync(buf);
    var shared = [];
    var ssF = zip.file('xl/sharedStrings.xml');
    if (ssF) {
      var ss = xml(await ssF.async('string'));
      Array.prototype.forEach.call(ss.getElementsByTagNameNS(NS_MAIN, 'si'), function (si) {
        var t = '';
        Array.prototype.forEach.call(si.getElementsByTagNameNS(NS_MAIN, 't'), function (x) {
          // пропускаем фонетические подсказки (rPh)
          if (x.parentNode && x.parentNode.localName === 'rPh') return;
          t += x.textContent;
        });
        shared.push(t);
      });
    }
    // какие стили — даты
    var dateStyles = {};
    var stF = zip.file('xl/styles.xml');
    if (stF) {
      var st = xml(await stF.async('string'));
      var custom = {};
      Array.prototype.forEach.call(st.getElementsByTagNameNS(NS_MAIN, 'numFmt'), function (f) {
        custom[f.getAttribute('numFmtId')] = f.getAttribute('formatCode') || '';
      });
      var xfsParent = st.getElementsByTagNameNS(NS_MAIN, 'cellXfs')[0];
      if (xfsParent) kids(xfsParent, NS_MAIN, 'xf').forEach(function (xf, i) {
        var id = +xf.getAttribute('numFmtId');
        var code = custom[id] || '';
        var isDate = (id >= 14 && id <= 22) || (id >= 45 && id <= 47) ||
          (code && /[dmy]/i.test(code.replace(/\[[^\]]*\]|"[^"]*"/g, '')) && !/^general$/i.test(code));
        if (isDate) dateStyles[i] = 1;
      });
    }
    var wb = xml(await zip.file('xl/workbook.xml').async('string'));
    var rels = xml(await zip.file('xl/_rels/workbook.xml.rels').async('string'));
    var relMap = {};
    Array.prototype.forEach.call(rels.getElementsByTagName('Relationship'), function (r) { relMap[r.getAttribute('Id')] = r.getAttribute('Target'); });
    var lines = [];
    var sheets = wb.getElementsByTagNameNS(NS_MAIN, 'sheet');
    for (var si2 = 0; si2 < sheets.length; si2++) {
      var rid = sheets[si2].getAttributeNS(NS_REL, 'id');
      var target = relMap[rid]; if (!target) continue;
      var path = target.charAt(0) === '/' ? target.slice(1) : 'xl/' + target.replace(/^\.\//, '');
      var f = zip.file(path); if (!f) continue;
      var sh = xml(await f.async('string'));
      var heads = null;
      Array.prototype.forEach.call(sh.getElementsByTagNameNS(NS_MAIN, 'row'), function (row) {
        var cells = [], numeric = {};
        kids(row, NS_MAIN, 'c').forEach(function (c) {
          var idx = colIndex(c.getAttribute('r'));
          if (idx < 0) idx = cells.length;
          var t = c.getAttribute('t'), s = +(c.getAttribute('s') || 0);
          var v = kids(c, NS_MAIN, 'v')[0], val = '';
          if (t === 's') val = v ? (shared[+v.textContent] || '') : '';
          else if (t === 'inlineStr') { var is = kids(c, NS_MAIN, 'is')[0]; val = is ? is.textContent : ''; }
          else if (t === 'str' || t === 'e') val = v ? v.textContent : '';
          else if (t === 'b') val = v ? v.textContent : '';
          else if (v) {
            var num = Number(v.textContent);
            if (dateStyles[s] && num > 0 && num < 80000) val = SBCore.fmtDate(excelDate(num));
            else if (isFinite(num) && Math.floor(num) === num) { val = num.toFixed(0); numeric[idx] = 1; }
            else val = v.textContent;
          }
          cells[idx] = String(val).replace(/[\t\n]+/g, ' ');
        });
        for (var k = 0; k < cells.length; k++) if (cells[k] === undefined) cells[k] = '';
        if (!cells.join('').trim()) return;
        var low = cells.map(function (x) { return String(x).toLowerCase(); });
        if (!heads && low.some(function (x) { return /фио|фамили|паспорт|серия|рожд/.test(x); })) heads = low;  // строка заголовков
        else if (heads) {
          for (var c2 = 0; c2 < cells.length; c2++) {
            var h = heads[c2] || '';
            // ИНН, номер договора, телефон — не паспорт: в разбор людей не идут
            if (/инн|огрн|кпп|снилс|договор|телефон|(^|[^а-яё])тел([^а-яё]|$)|организаци/.test(h)) { cells[c2] = ''; continue; }
            if (!numeric[c2]) continue;      // чиним только числа: у текста Excel нули не теряет
            var v2 = cells[c2], n2 = +v2;
            if (/рожд/.test(h) && n2 > 3000 && n2 < 60000) cells[c2] = SBCore.fmtDate(excelDate(n2));   // дата числом
            else if (/серия/.test(h) && !/номер/.test(h) && v2.length < 4) cells[c2] = ('0000' + v2).slice(-4);  // 0312 → 312
            else if (/номер/.test(h) && !/серия|телеф/.test(h) && v2.length < 6) cells[c2] = ('000000' + v2).slice(-6);
            else if (/паспорт/.test(h) && v2.length >= 7 && v2.length < 10) cells[c2] = ('0000000000' + v2).slice(-10);
          }
        }
        lines.push(cells.join('\t'));
      });
      lines.push('');
    }
    return lines.join('\n');
  }

  /* ---------- Word → текст ---------- */
  function wText(node) {
    var s = '';
    (function walk(n) {
      for (var c = n.firstChild; c; c = c.nextSibling) {
        if (c.nodeType !== 1) continue;
        if (c.namespaceURI === NS_W) {
          if (c.localName === 't') { s += c.textContent; continue; }
          if (c.localName === 'tab') { s += '\t'; continue; }
          if (c.localName === 'br' || c.localName === 'cr') { s += '\n'; continue; }
          if (c.localName === 'p' && s && !/\n$/.test(s)) s += '\n';
          if (c.localName === 'tbl') continue;
        }
        walk(c);
      }
    })(node);
    return s;
  }
  async function readDocx(buf) {
    var zip = await JSZip.loadAsync(buf);
    var doc = xml(await zip.file('word/document.xml').async('string'));
    var body = doc.getElementsByTagNameNS(NS_W, 'body')[0];
    var lines = [];
    (function walkBody(n) {
      for (var c = n.firstChild; c; c = c.nextSibling) {
        if (c.nodeType !== 1) continue;
        if (c.namespaceURI === NS_W && c.localName === 'p') lines.push(wText(c).replace(/\n+$/, ''));
        else if (c.namespaceURI === NS_W && c.localName === 'tbl') {
          kids(c, NS_W, 'tr').forEach(function (tr) {
            var cells = kids(tr, NS_W, 'tc').map(function (tc) { return wText(tc).replace(/\s*\n\s*/g, ' ').trim(); });
            lines.push(cells.join('\t'));
          });
        } else walkBody(c);
      }
    })(body);
    return lines.join('\n');
  }

  /* ---------- PDF → текст (pdf.js работает внутри устройства) ---------- */
  var pdfjsPromise = null;
  async function readPdf(buf) {
    if (!pdfjsPromise) pdfjsPromise = import('./vendor/pdf.min.mjs').then(function (m) {
      m.GlobalWorkerOptions.workerSrc = new URL('./vendor/pdf.worker.min.mjs', document.baseURI).href;
      return m;
    });
    var pdfjs = await pdfjsPromise;
    var doc = await pdfjs.getDocument({ data: new Uint8Array(buf), isEvalSupported: false, disableFontFace: true,
      useSystemFonts: false, stopAtErrors: false }).promise;
    var out = [];
    for (var p = 1; p <= doc.numPages; p++) {
      var page = await doc.getPage(p);
      var tc = await page.getTextContent();
      var rows = [];
      tc.items.forEach(function (it) {
        if (!it.str) return;
        var y = Math.round(it.transform[5]), x = it.transform[4];
        var row = rows.filter(function (r) { return Math.abs(r.y - y) <= 3; })[0];
        if (!row) { row = { y: y, items: [] }; rows.push(row); }
        row.items.push({ x: x, s: it.str, w: it.width || 0 });
      });
      rows.sort(function (a, b) { return b.y - a.y; });
      rows.forEach(function (r) {
        r.items.sort(function (a, b) { return a.x - b.x; });
        var line = '', lastEnd = null;
        r.items.forEach(function (it) {
          if (lastEnd !== null) line += (it.x - lastEnd > 12) ? '\t' : (it.x - lastEnd > 1 ? ' ' : '');
          line += it.s; lastEnd = it.x + it.w;
        });
        out.push(line);
      });
      out.push('');
    }
    await doc.destroy();
    return out.join('\n');
  }

  async function readText(buf) {
    var u8 = new Uint8Array(buf);
    var t = new TextDecoder('utf-8').decode(u8);
    if ((t.match(/�/g) || []).length > 3) {
      try { t = new TextDecoder('windows-1251').decode(u8); } catch (e) {}
    }
    return t.replace(/;/g, '\t');
  }

  async function readFile(file) {
    var name = (file.name || '').toLowerCase();
    var buf = await file.arrayBuffer();
    if (/\.(xlsx|xlsm)$/.test(name)) return readXlsx(buf);
    if (/\.docx$/.test(name)) return readDocx(buf);
    if (/\.pdf$/.test(name)) return readPdf(buf);
    if (/\.(xls|doc)$/.test(name)) throw new Error('Старый формат .' + name.split('.').pop() + ' не читается. Пересохраните файл как .' + name.split('.').pop() + 'x или скопируйте текст.');
    return readText(buf);
  }

  /* ---------- сборка Excel (временный макет до шаблона) ---------- */
  function colName(i) { var s = ''; i++; while (i > 0) { var m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); } return s; }
  function cell(ref, value, style) {
    return '<c r="' + ref + '" t="inlineStr" s="' + style + '"><is><t xml:space="preserve">' + esc(value) + '</t></is></c>';
  }

  async function buildXlsx(people, meta) {
    var headers = ['№ п/п', 'Фамилия', 'Имя', 'Отчество', 'Серия и номер паспорта', 'Дата рождения', 'Телефон'];
    var widths = [7, 18, 14, 18, 22, 15, 15];
    var rows = [];
    rows.push('<row r="1">' + cell('A1', 'Приложение к письму № ' + meta.number + ' от ' + meta.date, 1) + '</row>');
    rows.push('<row r="3">' + headers.map(function (h, i) { return cell(colName(i) + '3', h, 2); }).join('') + '</row>');
    people.forEach(function (p, k) {
      var r = k + 4;
      var vals = [String(k + 1), p.surname, p.name, p.patr, p.passport, p.birth, p.phone];
      rows.push('<row r="' + r + '">' + vals.map(function (v, i) { return cell(colName(i) + r, v, 3); }).join('') + '</row>');
    });
    var sheet = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<worksheet xmlns="' + NS_MAIN + '" xmlns:r="' + NS_REL + '">' +
      '<cols>' + widths.map(function (w, i) { return '<col min="' + (i + 1) + '" max="' + (i + 1) + '" width="' + w + '" customWidth="1"/>'; }).join('') + '</cols>' +
      '<sheetData>' + rows.join('') + '</sheetData>' +
      '<pageMargins left="0.5" right="0.5" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>' +
      '<pageSetup orientation="landscape" paperSize="9"/></worksheet>';
    var styles = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<styleSheet xmlns="' + NS_MAIN + '">' +
      '<fonts count="2"><font><sz val="11"/><name val="Times New Roman"/></font><font><b/><sz val="11"/><name val="Times New Roman"/></font></fonts>' +
      '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
      '<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border>' +
      '<border><left style="thin"/><right style="thin"/><top style="thin"/><bottom style="thin"/><diagonal/></border></borders>' +
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
      '<cellXfs count="4">' +
      '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
      '<xf numFmtId="49" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' +
      '<xf numFmtId="49" fontId="1" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>' +
      '<xf numFmtId="49" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>' +
      '</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>';
    var zip = new JSZip();
    zip.file('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
      '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
      '</Types>');
    zip.file('_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
      '</Relationships>');
    zip.file('xl/workbook.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<workbook xmlns="' + NS_MAIN + '" xmlns:r="' + NS_REL + '"><sheets><sheet name="Список" sheetId="1" r:id="rId1"/></sheets></workbook>');
    zip.file('xl/_rels/workbook.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
      '</Relationships>');
    zip.file('xl/worksheets/sheet1.xml', sheet);
    zip.file('xl/styles.xml', styles);
    return zip.generateAsync({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', compression: 'DEFLATE' });
  }

  root.SBIO = { readFile: readFile, readXlsx: readXlsx, readDocx: readDocx, readPdf: readPdf, buildXlsx: buildXlsx };
})(window);
