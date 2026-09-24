/* Ядро: разбор списка людей, проверки, телефоны, номер исходящего, даты.
   Работает только в памяти, ничего не сохраняет и не отправляет. */
(function (root) {
  'use strict';

  var MONTHS = {
    'января': 1, 'февраля': 2, 'марта': 3, 'апреля': 4, 'мая': 5, 'июня': 6,
    'июля': 7, 'августа': 8, 'сентября': 9, 'октября': 10, 'ноября': 11, 'декабря': 12
  };

  // Слова, которые пишутся с заглавной, но не бывают частью ФИО
  var STOP = ('паспорт паспорта паспортные серия серии номер дата даты рождения рожд выдан выдана выдано выдачи код ' +
    'подразделения гу гуву мвд россии рф уфмс оуфмс уфмс овд умвд ровд тп оуфмс отдел отделом отделения отделением отделение ' +
    'управление управлением управления главным главное федеральной миграционной службы службой внутренних дел ' +
    'по г город города городе гор края краю край области обл области республики республике республика район района ' +
    'районе округа округе автономного административного муниципального центрального северного южного западного ' +
    'восточного тел телефон мобильный фио фамилия имя отчество сотрудник сотрудники сотрудника список списка приложение ' +
    'итого должность монтажник инженер техник водитель мастер бригадир прораб электромонтажник ооо ао пао зао ип ' +
    'бизнес трейд бт телеком сбербанк сбер банк гражданство гражданин гражданка адрес регистрация место прописка ' +
    'дополнительно примечание всего добрый день коллеги спасибо просьба пожалуйста уважаемые прошу согласовать доступ ' +
    'допуск снилс инн пол муж жен январь февраль март апрель май июнь июль август сентябрь октябрь ноябрь декабрь ' +
    'краснодарскому ростовской ставропольскому москве москва санкт петербург петербурге'
  ).split(/\s+/).reduce(function (m, w) { if (w) m[w] = 1; return m; }, {});

  var PATR = /(вич|вна|ич|ична|оглы|кызы|угли|кизи)$/i;
  var WORD_RE = '(?:[А-ЯЁ][а-яё]+(?:-[А-ЯЁ][а-яё]+)?|[А-ЯЁ]{2,}(?:-[А-ЯЁ]{2,})?)';
  var TAIL_RE = '(?:[ \\t]+(?:оглы|кызы|угли|кизи))?';
  var LETTER = /[А-Яа-яЁёA-Za-z]/;

  function pad2(n) { return (n < 10 ? '0' : '') + n; }
  function fmtDate(d) { return pad2(d.getDate()) + '.' + pad2(d.getMonth() + 1) + '.' + d.getFullYear(); }
  function mkDate(y, m, d) {
    var dt = new Date(y, m - 1, d);
    if (dt.getFullYear() !== y || dt.getMonth() !== m - 1 || dt.getDate() !== d) return null;
    return dt;
  }
  function fullYear(yy) {
    var y = parseInt(yy, 10);
    if (yy.length === 4) return y;
    var cur = new Date().getFullYear() % 100;
    return y > cur ? 1900 + y : 2000 + y;
  }
  function parseDateStr(s) {
    s = String(s || '').trim();
    var m = s.match(/^(\d{1,2})[.\/\-](\d{1,2})[.\/\-](\d{2}|\d{4})$/);
    if (m) return mkDate(fullYear(m[3]), +m[2], +m[1]);
    m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (m) return mkDate(+m[1], +m[2], +m[3]);
    m = s.toLowerCase().match(/^(\d{1,2})\s+([а-я]+)\s+(\d{4})/);
    if (m && MONTHS[m[2]]) return mkDate(+m[3], MONTHS[m[2]], +m[1]);
    return null;
  }
  function addDays(d, n) { var x = new Date(d.getFullYear(), d.getMonth(), d.getDate() + n); return x; }
  function addMonths(d, n) {
    var y = d.getFullYear(), m = d.getMonth() + n, day = d.getDate();
    var last = new Date(y, m + 1, 0).getDate();
    return new Date(y, m, Math.min(day, last));
  }
  function ageOn(birth, on) {
    var a = on.getFullYear() - birth.getFullYear();
    if (on.getMonth() < birth.getMonth() || (on.getMonth() === birth.getMonth() && on.getDate() < birth.getDate())) a--;
    return a;
  }

  function mask(text, start, end) {
    return text.slice(0, start) + new Array(end - start + 1).join('·') + text.slice(end);
  }

  function titleCase(w) {
    return w.split('-').map(function (p) {
      if (!p) return p;
      return p.charAt(0).toUpperCase() + p.slice(1).toLowerCase();
    }).join('-');
  }

  function lineStartOk(text, pos) {
    var ls = text.lastIndexOf('\n', pos - 1) + 1;
    return /^[\s\d.)|№\-·]*$/.test(text.slice(ls, pos));
  }
  function lineIndexAt(lineStarts, pos) {
    var lo = 0, hi = lineStarts.length - 1;
    while (lo < hi) { var mid = (lo + hi + 1) >> 1; if (lineStarts[mid] <= pos) lo = mid; else hi = mid - 1; }
    return lo;
  }

  /* ---------- извлечение сущностей ---------- */

  function findDates(text) {
    var out = [];
    var res = [
      /(^|[^\d])(\d{1,2})[.\/](\d{1,2})[.\/](\d{4}|\d{2})(?!\d)/g,
      /(^|[^\d])(\d{1,2})-(\d{1,2})-(\d{4})(?!\d)/g,
      /(^|[^\d])(\d{4})-(\d{2})-(\d{2})(?!\d)/g,
      /(^|[^\dА-Яа-яЁё])(\d{1,2})\s+(января|февраля|марта|апреля|мая|июня|июля|августа|сентября|октября|ноября|декабря)\s+(\d{4})(?:\s*г\.?(?:ода)?)?/gi
    ];
    res.forEach(function (re, idx) {
      var m;
      while ((m = re.exec(text))) {
        var start = m.index + m[1].length, end = m.index + m[0].length;
        var d;
        if (idx === 2) d = mkDate(+m[2], +m[3], +m[4]);
        else if (idx === 3) d = mkDate(+m[4], MONTHS[m[3].toLowerCase()], +m[2]);
        else d = mkDate(fullYear(m[4]), +m[3], +m[2]);
        if (!d) continue;
        if (out.some(function (o) { return start < o.end && end > o.start; })) continue;
        // метка: ищем ближайшее слово слева в пределах 40 символов на той же строке
        var ls = text.lastIndexOf('\n', start - 1) + 1;
        var back = text.slice(Math.max(ls, start - 40), start).toLowerCase();
        var label = '';
        var bi = Math.max(back.lastIndexOf('рожд'), back.search(/д\.\s?р|г\.\s?р|родил/) );
        var ii = Math.max(back.lastIndexOf('выда'), back.lastIndexOf('выд.'));
        if (bi >= 0 || ii >= 0) label = bi > ii ? 'birth' : 'issue';
        out.push({ type: 'date', start: start, end: end, date: d, label: label });
      }
    });
    return out;
  }

  function maskNoise(text) {
    var res = [
      /()(\d{3}-\d{3}-\d{3}[ \-]\d{2})/g,                                     // СНИЛС
      /()(\+\s?7[\s(\-]*\d{3}[\s)\-]*\d{3}[\s\-]*\d{2}[\s\-]*\d{2})/g,         // +7 телефоны
      /(^|[^\d])(8[\s(\-]*9\d{2}[\s)\-]*\d{3}[\s\-]*\d{2}[\s\-]*\d{2})(?!\d)/g, // 8 9xx телефоны
      /(^|[^\d])([78]9\d{9})(?!\d)/g,                                          // 79xxxxxxxxx
      /(^|[^\d])(\d{3}-\d{3})(?![\d\-])/g                                      // код подразделения
    ];
    res.forEach(function (re) {
      text = text.replace(re, function (all, pre, body) {
        return pre + new Array(body.length + 1).join('·');
      });
    });
    return text;
  }

  function findPassports(text) {
    // убираем слова-связки между серией и номером, сохраняя позиции
    var t = text.replace(/(серия|серии|сер\.|номер|№|No\.?|N)(?=[\s\d])/gi, function (m) { return new Array(m.length + 1).join(' '); });
    var out = [];
    var re = /\d(?:[ \t|\-]*\d)*/g, m;
    while ((m = re.exec(t))) {
      var cluster = m[0];
      var groups = [], gre = /\d+/g, g;
      while ((g = gre.exec(cluster))) groups.push({ s: g[0], at: m.index + g.index });
      var total = groups.reduce(function (a, x) { return a + x.s.length; }, 0);
      if (total < 7) continue;
      var found = null;
      if (total === 10) found = [0, groups.length - 1];
      else if (total > 10) {
        // ищем подряд идущие группы на 10 цифр, предпочитая окно в конце
        for (var i = groups.length - 1; i >= 0 && !found; i--) {
          var sum = 0;
          for (var j = i; j >= 0; j--) {
            sum += groups[j].s.length;
            if (sum === 10) { found = [j, i]; break; }
            if (sum > 10) break;
          }
        }
      }
      if (found) {
        var digits = groups.slice(found[0], found[1] + 1).map(function (x) { return x.s; }).join('');
        var st = groups[found[0]].at, en = groups[found[1]].at + groups[found[1]].s.length;
        out.push({ type: 'passport', start: st, end: en, digits: digits, ok: true });
      } else {
        out.push({ type: 'passport', start: m.index, end: m.index + cluster.length,
          digits: groups.map(function (x) { return x.s; }).join(''), ok: false });
      }
    }
    return out;
  }

  function findNames(text) {
    var out = [];
    // инициалы: «Иванов И.И.» и «И.И. Иванов»
    var ini1 = new RegExp('(^|[^А-Яа-яЁё])(' + WORD_RE + ')[ \\t]+([А-ЯЁ])\\.[ \\t]?([А-ЯЁ])\\.?', 'g');
    var ini2 = new RegExp('(^|[^А-Яа-яЁё])([А-ЯЁ])\\.[ \\t]?([А-ЯЁ])\\.[ \\t]?(' + WORD_RE + ')', 'g');
    var m;
    while ((m = ini1.exec(text))) {
      if (STOP[m[2].toLowerCase()]) continue;
      var s = m.index + m[1].length;
      out.push({ type: 'fio', start: s, end: m.index + m[0].length, initials: true,
        surname: titleCase(m[2]), name: m[3] + '.', patr: m[4] + '.' });
    }
    while ((m = ini2.exec(text))) {
      if (STOP[m[4].toLowerCase()]) continue;
      var s2 = m.index + m[1].length;
      if (out.some(function (o) { return s2 < o.end && s2 + m[0].length > o.start; })) continue;
      out.push({ type: 'fio', start: s2, end: m.index + m[0].length, initials: true,
        surname: titleCase(m[4]), name: m[2] + '.', patr: m[3] + '.' });
    }
    var masked = text;
    out.forEach(function (o) { masked = mask(masked, o.start, o.end); });

    var runRe = new RegExp('(^|[^А-Яа-яЁёA-Za-z])(' + WORD_RE + TAIL_RE + '(?:[ \\t|]+' + WORD_RE + TAIL_RE + ')*)', 'g');
    while ((m = runRe.exec(masked))) {
      var runStart = m.index + m[1].length;
      var toks = [], tre = new RegExp(WORD_RE + TAIL_RE, 'g'), t;
      while ((t = tre.exec(m[2]))) toks.push({ w: t[0].replace(/[ \t]+/g, ' '), at: runStart + t.index, len: t[0].length });
      // режем по стоп-словам
      var chunks = [[]];
      toks.forEach(function (tk) {
        var base = tk.w.split(' ')[0].toLowerCase();
        if (STOP[base]) chunks.push([]); else chunks[chunks.length - 1].push(tk);
      });
      chunks.forEach(function (ch) {
        while (ch.length >= 2) {
          var take, fio;
          if (ch.length >= 3 && PATR.test(ch[2].w.split(' ')[0]) ) {
            take = 3; fio = { surname: ch[0].w, name: ch[1].w, patr: ch[2].w, strong: true };
          } else if (ch.length >= 3 && PATR.test(ch[1].w.split(' ')[0]) && !PATR.test(ch[0].w)) {
            take = 3; fio = { surname: ch[2].w, name: ch[0].w, patr: ch[1].w, strong: true };
          } else if (ch.length >= 3 && / (оглы|кызы|угли|кизи)$/i.test(ch[2].w)) {
            take = 3; fio = { surname: ch[0].w, name: ch[1].w, patr: ch[2].w, strong: true };
          } else if (ch.length === 2 || ch.length === 3) {
            take = ch.length; fio = { surname: ch[0].w, name: ch[1].w, patr: ch[2] ? ch[2].w : '', strong: false };
          } else {
            // 4+ слов без отчества на позиции 3 — берём первые два как слабое ФИО
            take = 2; fio = { surname: ch[0].w, name: ch[1].w, patr: '', strong: false };
          }
          var first = ch[0], last = ch[take - 1];
          fio.type = 'fio';
          fio.start = first.at; fio.end = last.at + last.len;
          fio.surname = titleCase(fio.surname); fio.name = titleCase(fio.name);
          fio.patr = fio.patr.split(' ').map(function (p, i) { return i === 0 ? titleCase(p) : p.toLowerCase(); }).join(' ');
          out.push(fio);
          ch = ch.slice(take);
        }
      });
    }
    var hasStrong = out.some(function (o) { return o.strong || o.initials; });
    return out.filter(function (o) {
      if (o.initials || o.strong) return true;
      return !hasStrong || lineStartOk(text, o.start);
    });
  }

  /* ---------- сборка записей ---------- */

  function parsePeople(input, opts) {
    opts = opts || {};
    var refDate = opts.refDate || new Date();
    var text = String(input || '').replace(/\r\n?/g, '\n').replace(/[   ]/g, ' ').replace(/[–—]/g, '-');

    var dates = findDates(text);
    var t1 = text;
    dates.forEach(function (d) { t1 = mask(t1, d.start, d.end); });
    t1 = maskNoise(t1);
    var names = findNames(t1);
    var t2 = t1;
    names.forEach(function (n) { t2 = mask(t2, n.start, n.end); });
    var passports = findPassports(t2);

    var ents = [].concat(names, dates, passports).sort(function (a, b) { return a.start - b.start; });

    // построчная группировка для таблиц: в каждой строке ровно одно ФИО и есть паспорт или дата
    var lineStarts = [0];
    for (var i = 0; i < text.length; i++) if (text[i] === '\n') lineStarts.push(i + 1);
    var byLine = {};
    ents.forEach(function (e) { var li = lineIndexAt(lineStarts, e.start); (byLine[li] = byLine[li] || []).push(e); });
    var lineKeys = Object.keys(byLine);
    var tableLike = names.length > 0 && lineKeys.every(function (k) {
      var arr = byLine[k], f = arr.filter(function (e) { return e.type === 'fio'; }).length;
      return f === 0 || (f === 1 && arr.length > 1);
    }) && lineKeys.filter(function (k) { return byLine[k].some(function (e) { return e.type === 'fio'; }); })
      .every(function (k) { return byLine[k].some(function (e) { return e.type !== 'fio'; }); });

    var groups = [];
    if (tableLike) {
      lineKeys.forEach(function (k) {
        var arr = byLine[k];
        if (arr.some(function (e) { return e.type === 'fio'; })) groups.push(arr);
      });
    } else {
      var pending = [], cur = null;
      ents.forEach(function (e) {
        if (e.type === 'fio') { cur = [e]; groups.push(cur); }
        else if (cur) cur.push(e);
        else pending.push(e);
      });
      if (groups.length && pending.length) groups[0] = pending.concat(groups[0]);
    }

    var people = groups.map(function (g) {
      var fio = g.filter(function (e) { return e.type === 'fio'; })[0];
      var pps = g.filter(function (e) { return e.type === 'passport'; });
      var ds = g.filter(function (e) { return e.type === 'date'; });
      var good = pps.filter(function (p) { return p.ok; });
      var p = {
        surname: fio.surname, name: fio.name, patr: fio.patr || '',
        passport: good.length ? good[0].digits : (pps.length ? pps[0].digits : ''),
        birth: '', flags: []
      };
      if (fio.initials) p.flags.push('initials');
      if (good.length > 1 && good[0].digits !== good[1].digits) p.flags.push('twoPassports');
      var birth = ds.filter(function (d) { return d.label === 'birth'; })[0];
      if (!birth) {
        var cands = ds.filter(function (d) { return d.label !== 'issue'; });
        if (!cands.length) cands = [];
        cands.sort(function (a, b) { return a.date - b.date; });
        birth = cands[0];
      }
      if (birth) p.birth = fmtDate(birth.date);
      return p;
    });
    return validate(people, refDate);
  }

  function validate(people, refDate) {
    refDate = refDate || new Date();
    var seen = {};
    people.forEach(function (p) {
      var errors = [], warnings = [];
      p.surname = (p.surname || '').trim(); p.name = (p.name || '').trim(); p.patr = (p.patr || '').trim();
      p.passport = String(p.passport || '').replace(/\D/g, '');
      p.birth = (p.birth || '').trim();
      var flags = p.flags || [];
      if (flags.indexOf('initials') >= 0 || /\.$/.test(p.name) || /\.$/.test(p.patr) || p.name.length === 1)
        errors.push('инициалы вместо полного ФИО');
      if (!p.surname || !p.name) errors.push('нет фамилии или имени');
      else if (!p.patr) warnings.push('нет отчества — проверьте, что его действительно нет');
      if (!p.passport) errors.push('нет паспорта');
      else if (p.passport.length !== 10) errors.push('паспорт: ' + p.passport.length + ' цифр вместо 10' +
        (p.passport.length === 9 ? ' (возможно, потерян 0 в начале)' : ''));
      if (flags.indexOf('twoPassports') >= 0) warnings.push('в записи два разных номера паспорта — проверьте');
      var bd = parseDateStr(p.birth);
      if (!p.birth) errors.push('нет даты рождения');
      else if (!bd) errors.push('дата рождения не распознана');
      else {
        var age = ageOn(bd, refDate);
        if (age < 18 || age > 85) errors.push('дата рождения вызывает сомнения (возраст ' + age + ')');
        p.birth = fmtDate(bd);
      }
      if (p.passport.length === 10) {
        if (seen[p.passport]) errors.push('паспорт повторяется в списке');
        seen[p.passport] = 1;
      }
      p.errors = errors; p.warnings = warnings;
    });
    return people;
  }

  /* ---------- телефоны, номер, срок ---------- */

  function randInt(max) {
    if (root.crypto && root.crypto.getRandomValues) {
      var a = new Uint32Array(1); root.crypto.getRandomValues(a); return a[0] % max;
    }
    return Math.floor(Math.random() * max);
  }
  function genPhone(used) {
    for (var k = 0; k < 1000; k++) {
      var s = '79';
      for (var i = 0; i < 9; i++) s += randInt(10);
      if (!used || !used[s]) { if (used) used[s] = 1; return s; }
    }
    throw new Error('phone generation failed');
  }
  function outNumber(n, date) {
    return n + '/' + pad2(date.getMonth() + 1) + '/' + pad2(date.getFullYear() % 100);
  }
  function accessPeriod(letterDate, nextDay) {
    var start = nextDay ? addDays(letterDate, 1) : new Date(letterDate.getFullYear(), letterDate.getMonth(), letterDate.getDate());
    var end = addMonths(start, 6);
    return { start: start, end: end };
  }

  /* ---------- справочник ---------- */

  function norm(s) { return String(s || '').toLowerCase().replace(/ё/g, 'е').replace(/[^a-zа-я0-9]+/g, ' ').trim(); }
  function searchGosb(list, q) {
    var nq = norm(q);
    if (!nq) return [];
    var parts = nq.split(' ');
    var scored = [];
    list.forEach(function (g) {
      var hay = norm([g.gosb, g.region, g.phrase, g.aliases, g.bank, g.bankCode].join(' '));
      var ok = parts.every(function (p) { return hay.indexOf(p) >= 0; });
      if (!ok) return;
      var score = 0;
      if (g.gosb && g.gosb === nq) score += 100;
      else if (g.gosb && g.gosb.indexOf(nq) === 0) score += 50;
      if (norm(g.region).indexOf(nq) === 0) score += 20;
      scored.push({ g: g, s: score });
    });
    scored.sort(function (a, b) { return b.s - a.s; });
    return scored.map(function (x) { return x.g; });
  }
  function subjectRegion(g) {
    return g.region
      .replace(/\s*\((Наро-Фоминск|Альметьевск|Казань|Набережные Челны|Новый Уренгой|Салехард|Сургут|Ханты-Мансийск)\)/, '')
      .replace('Ямало-Ненецкий АО', 'Ямало-Ненецкий автономный округ')
      .replace('Ханты-Мансийский АО', 'Ханты-Мансийский автономный округ');
  }
  function placePhrase(g) {
    return 'во все ВСП ' + g.bankGen + ' ГОСБ №' + g.gosb + ' ' + g.phrase + ' ПАО «Сбербанк»';
  }

  var API = {
    parsePeople: parsePeople, validate: validate, parseDateStr: parseDateStr, fmtDate: fmtDate,
    genPhone: genPhone, outNumber: outNumber, accessPeriod: accessPeriod, addMonths: addMonths, addDays: addDays,
    searchGosb: searchGosb, subjectRegion: subjectRegion, placePhrase: placePhrase, pad2: pad2
  };
  root.SBCore = API;
  if (typeof module !== 'undefined') module.exports = API;
})(typeof window !== 'undefined' ? window : globalThis);
