/* Интерфейс. Люди хранятся только в переменной state.people. */
(function () {
  'use strict';
  var VERSION = '0.2 · бланк, PDF, Excel по шаблону';
  var C = window.SBCore, G = window.SB_GOSB, IO = window.SBIO;
  var $ = function (id) { return document.getElementById(id); };

  var state = { gosb: null, people: [], usedPhones: {} };

  /* ---------- мелочи ---------- */
  function toast(msg) {
    var t = $('toast'); t.textContent = msg; t.hidden = false;
    clearTimeout(toast._t); toast._t = setTimeout(function () { t.hidden = true; }, 2200);
  }
  function el(tag, cls, text) { var e = document.createElement(tag); if (cls) e.className = cls; if (text != null) e.textContent = text; return e; }
  function copy(text) {
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(text).then(function () { toast('Скопировано'); }, function () { toast('Не удалось скопировать'); });
    }
  }
  function lsGet(k) { try { return localStorage.getItem(k); } catch (e) { return null; } }
  function lsSet(k, v) { try { localStorage.setItem(k, v); } catch (e) {} }
  function todayISO() { var d = new Date(); return d.getFullYear() + '-' + C.pad2(d.getMonth() + 1) + '-' + C.pad2(d.getDate()); }
  function letterDate() { var v = $('letterDate').value; var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(v); return m ? new Date(+m[1], +m[2] - 1, +m[3]) : new Date(); }

  /* ---------- защита: проверка, что сеть закрыта ---------- */
  function checkNetworkBlocked() {
    var violated = false;
    document.addEventListener('securitypolicyviolation', function (e) {
      if (/^connect-src/.test(e.violatedDirective || e.effectiveDirective || '')) violated = true;
    });
    var pill = $('pillNet');
    var probe = './__network_probe__?' + Date.now();
    var done = function (fetchBlocked) {
      setTimeout(function () {
        if (fetchBlocked && violated) { pill.className = 'pill ok'; pill.textContent = 'Отправка данных заблокирована'; }
        else if (fetchBlocked) { pill.className = 'pill ok'; pill.textContent = 'Отправка данных заблокирована'; }
        else { pill.className = 'pill bad'; pill.textContent = 'Защита не работает — не вводите данные'; }
      }, 250);
    };
    try { fetch(probe, { cache: 'no-store' }).then(function () { done(false); }, function () { done(true); }); }
    catch (e) { done(true); }
  }
  function showOnline() {
    var p = $('pillOnline');
    p.className = 'pill';
    p.textContent = navigator.onLine ? 'Интернет на телефоне включён' : 'Авиарежим / без интернета';
  }

  /* ---------- 1. ГОСБ ---------- */
  function renderGosbList() {
    var q = $('gosbQ').value, ul = $('gosbList');
    ul.textContent = '';
    var res = C.searchGosb(G.LIST, q).slice(0, 12);
    if (!q.trim()) { ul.hidden = true; return; }
    if (!res.length) { var li0 = el('li', null, 'Ничего не нашлось'); ul.appendChild(li0); ul.hidden = false; return; }
    res.forEach(function (g) {
      var li = el('li'); li.tabIndex = 0;
      li.appendChild(el('span', null, (g.gosb ? 'ГОСБ ' + g.gosb + ' · ' : '') + g.region));
      li.appendChild(el('span', 'sub', g.bank + ' банк' + (g.contract ? '' : ' · нет договора')));
      var choose = function () { chooseGosb(g); };
      li.addEventListener('click', choose);
      li.addEventListener('keydown', function (e) { if (e.key === 'Enter') choose(); });
      ul.appendChild(li);
    });
    ul.hidden = false;
  }
  function chooseGosb(g) {
    state.gosb = g;
    $('gosbList').hidden = true;
    $('gosbQ').value = (g.gosb ? g.gosb + ' ' : '') + g.region;
    var box = $('gosbChosen'); box.textContent = ''; box.hidden = false;
    box.className = 'chosen' + (g.contract && g.gosb ? '' : ' no');
    box.appendChild(el('b', null, (g.gosb ? 'ГОСБ ' + g.gosb + ' · ' : '') + g.region));
    box.appendChild(el('span', 'sub', g.bank + ' банк · ' + (g.contract ? 'договор ' + g.contract : 'нет договора — письмо не собирается')));
    if (!g.gosb) box.appendChild(el('span', 'sub', ' · номер ГОСБ не указан в справочнике'));
    if (g.bankCode !== 'ЮЗБ') box.appendChild(el('span', 'sub', ' · форма «' + g.phrase + '» не проверена на живом письме'));
    refresh();
  }

  /* ---------- 2. письмо ---------- */
  function outNoValue() {
    var n = ($('outNo').value || '').replace(/\D/g, '');
    return n ? C.outNumber(+n, letterDate()) : '';
  }
  function renderLetterMeta() {
    var d = letterDate();
    $('outNoSuffix').textContent = '/' + C.pad2(d.getMonth() + 1) + '/' + C.pad2(d.getFullYear() % 100);
    var next = document.querySelector('input[name=start]:checked').value === 'next';
    var p = C.accessPeriod(d, next);
    $('period').textContent = 'Срок допуска 6 месяцев: с ' + C.fmtDate(p.start) + ' по ' + C.fmtDate(p.end) + '.';
  }

  /* ---------- 3. люди ---------- */
  function ensurePhones() {
    state.people.forEach(function (p) {
      if (!p.phone) p.phone = C.genPhone(state.usedPhones);
    });
  }
  function revalidate() {
    C.validate(state.people, letterDate());
  }
  var FIELDS = [
    ['surname', 'Фамилия', 'text'], ['name', 'Имя', 'text'], ['patr', 'Отчество', 'text'],
    ['passport', 'Серия и номер', 'numeric'], ['birth', 'Дата рождения', 'text']
  ];
  function renderPeople() {
    var box = $('people'); box.textContent = '';
    state.people.forEach(function (p, i) {
      var card = el('div', 'person');
      var head = el('div', 'person-head');
      head.appendChild(el('span', 'idx', (i + 1) + '.'));
      var del = el('button', 'btn small ghost', 'Убрать'); del.type = 'button';
      del.addEventListener('click', function () {
        delete state.usedPhones[p.phone];
        state.people.splice(i, 1); revalidate(); renderPeople(); refresh();
      });
      head.appendChild(del);
      card.appendChild(head);
      var grid = el('div', 'pgrid');
      FIELDS.forEach(function (f) {
        var lab = el('label'); lab.appendChild(document.createTextNode(f[1]));
        var inp = el('input'); inp.type = 'text'; inp.value = p[f[0]] || '';
        inp.setAttribute('autocomplete', 'off'); inp.setAttribute('autocorrect', 'off');
        inp.setAttribute('autocapitalize', f[2] === 'text' && f[0] !== 'birth' ? 'words' : 'off');
        inp.spellcheck = false;
        if (f[2] === 'numeric' || f[0] === 'birth') inp.inputMode = f[0] === 'birth' ? 'decimal' : 'numeric';
        if (f[0] === 'birth') inp.placeholder = 'ДД.ММ.ГГГГ';
        inp.addEventListener('input', function () { p[f[0]] = inp.value; p.flags = []; revalidate(); updateStatuses(); refresh(); });
        inp.addEventListener('change', function () { revalidate(); if (f[0] === 'birth' || f[0] === 'passport') inp.value = p[f[0]]; updateStatuses(); });
        lab.appendChild(inp); grid.appendChild(lab);
      });
      var lab2 = el('label', 'phone-lab'); lab2.appendChild(document.createTextNode('Телефон (сгенерирован)'));
      var ph = el('div', 'phone');
      var pin = el('input'); pin.type = 'text'; pin.value = p.phone; pin.readOnly = true;
      var regen = el('button', 'btn small', '↻'); regen.type = 'button'; regen.title = 'Другой номер';
      regen.addEventListener('click', function () { delete state.usedPhones[p.phone]; p.phone = C.genPhone(state.usedPhones); pin.value = p.phone; });
      ph.appendChild(pin); ph.appendChild(regen); lab2.appendChild(ph); grid.appendChild(lab2);
      card.appendChild(grid);
      card.appendChild(el('div', 'status'));
      box.appendChild(card);
    });
    updateStatuses();
  }
  function updateStatuses() {
    var cards = $('people').children;
    var nErr = 0, nWarn = 0;
    state.people.forEach(function (p, i) {
      var card = cards[i]; if (!card) return;
      var st = card.querySelector('.status'); st.textContent = '';
      card.className = 'person' + (p.errors.length ? ' has-err' : (p.warnings.length ? ' has-warn' : ''));
      p.errors.forEach(function (e) { st.appendChild(el('span', 'e', '✕ ' + e)); });
      p.warnings.forEach(function (w) { st.appendChild(el('span', 'w', '! ' + w)); });
      if (!p.errors.length && !p.warnings.length) st.appendChild(el('span', 'ok', '✓ всё на месте'));
      if (p.errors.length) nErr++; else if (p.warnings.length) nWarn++;
    });
    var s = $('summary');
    if (!state.people.length) { s.hidden = true; return; }
    s.hidden = false;
    s.textContent = 'Людей: ' + state.people.length + (nErr ? ' · с ошибками: ' + nErr : '') + (nWarn ? ' · проверить: ' + nWarn : '') + (!nErr && !nWarn ? ' · все данные в порядке' : '');
  }
  function doParse() {
    var text = $('src').value;
    var msg = $('parseMsg');
    if (!text.trim()) { msg.hidden = false; msg.className = 'msg warn'; msg.textContent = 'Нечего разбирать: вставьте текст или откройте файл.'; return; }
    var found = C.parsePeople(text, { refDate: letterDate() });
    if (!found.length) { msg.hidden = false; msg.className = 'msg warn'; msg.textContent = 'Не нашёл ни одного ФИО. Проверьте текст или добавьте людей вручную.'; return; }
    state.usedPhones = {};
    state.people = found;
    ensurePhones();
    revalidate();
    $('src').value = '';
    msg.hidden = false; msg.className = 'msg'; msg.textContent = 'Найдено: ' + found.length + '. Исходный текст убран из поля. Проверьте карточки ниже.';
    renderPeople(); refresh();
  }
  function addManual() {
    state.people.push({ surname: '', name: '', patr: '', passport: '', birth: '', flags: [] });
    ensurePhones(); revalidate(); renderPeople(); refresh();
    var cards = $('people').children; var last = cards[cards.length - 1];
    if (last) { var i = last.querySelector('input'); if (i) i.focus(); }
  }

  /* ---------- шаблоны: хранятся только в этом браузере на этом устройстве ---------- */
  var tpl = { docx: null, xlsx: null };
  function idb() {
    return new Promise(function (res, rej) {
      var r = indexedDB.open('sb-dostup', 1);
      r.onupgradeneeded = function () { r.result.createObjectStore('files'); };
      r.onsuccess = function () { res(r.result); }; r.onerror = function () { rej(r.error); };
    });
  }
  function idbDo(mode, fn) {
    return idb().then(function (db) {
      return new Promise(function (res, rej) {
        var tx = db.transaction('files', mode), st = tx.objectStore('files'), req = fn(st);
        tx.oncomplete = function () { res(req && req.result); db.close(); };
        tx.onerror = function () { rej(tx.error); db.close(); };
      });
    });
  }
  function loadTemplates() {
    return Promise.all(['docx', 'xlsx'].map(function (k) {
      return idbDo('readonly', function (st) { return st.get(k); }).then(function (v) { tpl[k] = v || null; }).catch(function () { tpl[k] = null; });
    })).then(renderTemplates);
  }
  var TEST_PERSON = { fio: 'Тестов Тест Тестович', birth: '01.02.1990', passport: '0000 000000', surname: 'Тестов', name: 'Тест', patr: 'Тестович',
    birthDate: new Date(1990, 1, 1), series: '0000', number: '000000', phone: '79000000000' };
  function checkTemplate(kind, bytes) {
    if (kind === 'docx') {
      return SBFill.fillDocx(bytes, { outNo: '1/01/26', date: '01.01.2026', contract: '№ 1 от 01.01.26', place: 'Тестового банка ГОСБ №0000 в тестовой области',
        periodStart: '01.01.2026', periodEnd: '01.07.2026', people: [TEST_PERSON] }).then(function (r) {
        var f = r.found, miss = [];
        if (!f.table) throw new Error('не нашёл таблицу сотрудников с колонками «Ф.И.О.», «Дата рождения», «Паспорт»');
        if (!f.outNo) miss.push('номер исходящего «Исх.: …»');
        if (!f.date) miss.push('дату письма «от ДД.ММ.ГГГГ г.»');
        if (!f.contract) miss.push('договор «договора № … от …»');
        if (!f.place) miss.push('место «во все ВСП … ГОСБ № …»');
        if (!f.period) miss.push('срок «предоставить с … по …»');
        return SBDocxPdf.render(r.zip).then(function () { return miss; });
      });
    }
    return SBFill.fillXlsx(bytes, { people: [TEST_PERSON], contractDigits: '1', refText: '№ 1/01/26, от «01» января 2026 г.' }).then(function (r) {
      if (!r.found.columns) throw new Error('не нашёл колонки Фамилия, Имя, Отчество, Дата рождения, Паспорт (серия), Паспорт (номер), телефон, договор');
      return r.found.ref ? [] : ['строку со ссылкой на письмо «№ …/…/…, от …»'];
    });
  }
  function onTemplateFile(kind, input) {
    var f = input.files && input.files[0]; if (!f) return;
    input.value = '';
    var msg = $('tplMsg'); msg.hidden = false; msg.className = 'msg'; msg.textContent = 'Проверяю шаблон «' + f.name + '»…';
    if (!new RegExp('\\.' + kind + '$', 'i').test(f.name)) { msg.className = 'msg err'; msg.textContent = 'Нужен файл .' + kind; return; }
    f.arrayBuffer().then(function (buf) {
      return checkTemplate(kind, buf).then(function (miss) {
        var rec = { name: f.name, bytes: buf, savedAt: Date.now(), miss: miss };
        return idbDo('readwrite', function (st) { return st.put(rec, kind); }).then(function () {
          tpl[kind] = rec;
          if (navigator.storage && navigator.storage.persist) navigator.storage.persist().catch(function () {});
          msg.className = miss.length ? 'msg warn' : 'msg';
          msg.textContent = 'Шаблон сохранён на этом устройстве.' + (miss.length ? ' Не нашёл в нём: ' + miss.join('; ') + ' — эти места останутся как в шаблоне.' : ' Все поля найдены.');
          renderTemplates(); refresh();
        });
      });
    }).catch(function (e) { msg.className = 'msg err'; msg.textContent = 'Шаблон не подходит: ' + (e && e.message ? e.message : e); });
  }
  function renderTemplates() {
    ['docx', 'xlsx'].forEach(function (k) {
      var t = tpl[k], box = $(k === 'docx' ? 'tplDocx' : 'tplXlsx');
      box.textContent = t ? '✓ ' + t.name + ' · загружен ' + C.fmtDate(new Date(t.savedAt)) + (t.miss && t.miss.length ? ' · есть замечания' : '') : 'не загружен';
      box.className = t ? 'ok' : 'no';
    });
    $('btnTplDel').hidden = !(tpl.docx || tpl.xlsx);
    $('tplCard').classList.toggle('done', !!(tpl.docx && tpl.xlsx));
  }
  function deleteTemplates() {
    idbDo('readwrite', function (st) { st.delete('docx'); return st.delete('xlsx'); }).then(function () {
      tpl.docx = tpl.xlsx = null; renderTemplates(); refresh(); toast('Шаблоны удалены с этого устройства');
    });
  }

  /* ---------- 4. файлы ---------- */
  function blockers() {
    var b = [];
    if (!tpl.docx || !tpl.xlsx) b.push('загрузите шаблоны Word и Excel (раздел «Бланк» ниже)');
    if (!state.gosb) b.push('выберите ГОСБ');
    else if (!state.gosb.contract) b.push('для этого банка нет договора — письмо не собирается');
    else if (!state.gosb.gosb) b.push('в справочнике нет номера ГОСБ');
    if (!outNoValue()) b.push('укажите номер исходящего');
    if (!state.people.length) b.push('добавьте людей');
    var bad = state.people.filter(function (p) { return p.errors.length; }).length;
    if (bad) b.push('исправьте ошибки в карточках: ' + bad);
    return b;
  }
  function fileDate() { return C.fmtDate(letterDate()); }
  function xlsxName() { return 'Приложение Список сотрудников ' + fileDate() + '.xlsx'; }
  function letterName(ext) { return 'Письмо на допуск ' + fileDate() + '.' + ext; }
  function subject() { return state.gosb ? 'Согласование доступа / ГОСБ ' + state.gosb.gosb + ' ' + C.subjectRegion(state.gosb) : ''; }
  var BODY = 'Добрый день, коллеги.\n\nПросьба, пожалуйста, принять в работу документы из вложения.\n\nСпасибо.';

  function period() {
    var next = document.querySelector('input[name=start]:checked').value === 'next';
    return C.accessPeriod(letterDate(), next);
  }
  function buildAll() {
    var b = blockers();
    if (b.length) return Promise.reject(new Error(b.join('; ')));
    var g = state.gosb, d = letterDate(), per = period(), no = outNoValue();
    var people = state.people.map(function (p) {
      return {
        fio: [p.surname, p.name, p.patr].filter(Boolean).join(' '), birth: p.birth,
        passport: p.passport.slice(0, 4) + ' ' + p.passport.slice(4),
        surname: p.surname, name: p.name, patr: p.patr, birthDate: C.parseDateStr(p.birth),
        series: p.passport.slice(0, 4), number: p.passport.slice(4), phone: p.phone
      };
    });
    var dd = { outNo: no, date: C.fmtDate(d), contract: g.contract, place: g.bankGen + ' ГОСБ №' + g.gosb + ' ' + g.phrase,
      periodStart: C.fmtDate(per.start), periodEnd: C.fmtDate(per.end), people: people };
    var contractDigits = (/№\s*(\d+)/.exec(g.contract) || [])[1] || '';
    return Promise.all([
      SBFill.fillDocx(tpl.docx.bytes, dd).then(function (w) {
        return SBDocxPdf.render(w.zip).then(function (pdf) {
          return { docx: new File([w.blob], letterName('docx'), { type: w.blob.type }), pdf: new File([pdf], letterName('pdf'), { type: 'application/pdf' }) };
        });
      }),
      SBFill.fillXlsx(tpl.xlsx.bytes, { people: people, contractDigits: contractDigits, refText: SBFill.refText(no, d) })
    ]).then(function (r) {
      if (!r[1].blob) throw new Error('не удалось заполнить Excel по шаблону');
      lsSet('sb.lastNo', ($('outNo').value || '').replace(/\D/g, ''));
      return { pdf: r[0].pdf, docx: r[0].docx, xlsx: new File([r[1].blob], xlsxName(), { type: r[1].blob.type }) };
    });
  }
  function saveFile(file) {
    var url = URL.createObjectURL(file);
    var a = document.createElement('a'); a.href = url; a.download = file.name;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(function () { URL.revokeObjectURL(url); }, 60000);
  }
  function busy(on) {
    ['btnShare', 'btnPdf', 'btnXlsx', 'btnDocx'].forEach(function (id) { $(id).disabled = on || !!blockers().length; });
    $('buildMsg').hidden = !on; $('buildMsg').textContent = on ? 'Собираю файлы…' : '';
  }
  function fail(e) { busy(false); toast('Не получилось: ' + (e && e.message ? e.message : e)); }
  function onShare() {
    busy(true);
    buildAll().then(function (f) {
      busy(false);
      var files = [f.pdf, f.xlsx];
      if (navigator.canShare && navigator.canShare({ files: files })) {
        return navigator.share({ files: files, title: subject() }).catch(function (e) {
          if (e && e.name === 'AbortError') return;
          files.forEach(saveFile);
        });
      }
      files.forEach(saveFile); toast('Отправка из этого браузера недоступна — файлы сохранены');
    }).catch(fail);
  }
  function onSave(kind) { busy(true); buildAll().then(function (f) { busy(false); saveFile(f[kind]); }).catch(fail); }

  function renderLetterFields() {
    var dl = $('letterFields'); dl.textContent = '';
    var rows = [
      ['Тема письма', subject()],
      ['Текст письма', BODY],
      ['Файлы в банк', letterName('pdf') + '\n' + xlsxName()]
    ];
    rows.forEach(function (r) {
      if (!r[1]) return;
      var wrap = el('div');
      wrap.appendChild(el('dt', null, r[0]));
      wrap.appendChild(el('dd', null, r[1]));
      if (r[0] !== 'Файлы в банк') {
        var b = el('button', 'btn small', 'Копировать'); b.type = 'button';
        b.addEventListener('click', function () { copy(r[1]); });
        wrap.appendChild(b);
      }
      dl.appendChild(wrap);
    });
  }

  function refresh() {
    renderLetterMeta();
    renderLetterFields();
    var b = blockers();
    var box = $('blockers');
    box.hidden = !b.length;
    box.textContent = b.length ? 'Чтобы собрать файлы: ' + b.join('; ') + '.' : '';
    ['btnShare', 'btnPdf', 'btnXlsx', 'btnDocx'].forEach(function (id) { $(id).disabled = !!b.length; });
  }

  function clearAll(silent) {
    state.people = []; state.usedPhones = {};
    $('src').value = ''; $('fileIn').value = '';
    $('parseMsg').hidden = true;
    renderPeople(); refresh();
    if (!silent) toast('Список очищен');
  }

  /* ---------- запуск ---------- */
  function init() {
    $('ver').textContent = 'Версия ' + VERSION;
    checkNetworkBlocked(); showOnline();
    window.addEventListener('online', showOnline); window.addEventListener('offline', showOnline);

    $('letterDate').value = todayISO();
    var last = parseInt(lsGet('sb.lastNo') || '60', 10);
    $('outNo').value = String((isFinite(last) ? last : 60) + 1);

    $('gosbQ').addEventListener('input', renderGosbList);
    $('gosbQ').addEventListener('focus', renderGosbList);
    ['outNo', 'letterDate'].forEach(function (id) { $(id).addEventListener('input', function () { revalidate(); updateStatuses(); refresh(); }); });
    Array.prototype.forEach.call(document.querySelectorAll('input[name=start]'), function (r) { r.addEventListener('change', refresh); });

    $('btnPaste').addEventListener('click', function () {
      if (navigator.clipboard && navigator.clipboard.readText) {
        navigator.clipboard.readText().then(function (t) { $('src').value = t; doParse(); }, function () { $('src').focus(); toast('Вставьте текст в поле'); });
      } else { $('src').focus(); toast('Вставьте текст в поле'); }
    });
    $('fileIn').addEventListener('change', function () {
      var f = this.files && this.files[0]; if (!f) return;
      var msg = $('parseMsg'); msg.hidden = false; msg.className = 'msg'; msg.textContent = 'Читаю «' + f.name + '»…';
      IO.readFile(f).then(function (text) {
        $('fileIn').value = '';
        $('src').value = text; doParse();
      }).catch(function (e) {
        $('fileIn').value = '';
        msg.className = 'msg err'; msg.textContent = 'Не удалось прочитать файл: ' + (e && e.message ? e.message : e) + '. Можно скопировать текст и вставить.';
      });
    });
    $('btnParse').addEventListener('click', doParse);
    $('btnAdd').addEventListener('click', addManual);
    $('btnShare').addEventListener('click', onShare);
    $('btnPdf').addEventListener('click', function () { onSave('pdf'); });
    $('btnXlsx').addEventListener('click', function () { onSave('xlsx'); });
    $('btnDocx').addEventListener('click', function () { onSave('docx'); });
    $('tplDocxIn').addEventListener('change', function () { onTemplateFile('docx', this); });
    $('tplXlsxIn').addEventListener('change', function () { onTemplateFile('xlsx', this); });
    $('btnTplDel').addEventListener('click', deleteTemplates);
    $('btnClear').addEventListener('click', function () { clearAll(false); });

    // автоочистка, если инструмент свёрнут дольше 15 минут
    var hiddenAt = 0;
    document.addEventListener('visibilitychange', function () {
      if (document.hidden) hiddenAt = Date.now();
      else if (hiddenAt && Date.now() - hiddenAt > 15 * 60 * 1000 && state.people.length) { clearAll(true); toast('Список очищен: инструмент был свёрнут больше 15 минут'); }
    });

    refresh();
    loadTemplates().then(refresh);

    if ('serviceWorker' in navigator && (location.protocol === 'https:' || location.hostname === 'localhost' || location.hostname === '127.0.0.1')) {
      navigator.serviceWorker.register('sw.js').catch(function () {});
    }
  }
  init();
})();
