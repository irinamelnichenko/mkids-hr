/* v7.599 «📄 Звіт для батьків (PDF)» — модуль «Розвиток» картки дитини (clients.html, вкладка 🌱).
   Дані — лише те, що внесено: c.development = {'РРРР-ММ': {criteria:{'1.1':'+'|'-'}, by, at}} (точки Вересень / Січень / Червень),
   шаблон критеріїв — DEVELOPMENT_TEMPLATES за групою. Коментарів і фото в «Розвитку» немає — у звіті їх теж немає.
   PDF — нова вкладка зі звітом → window.print() «Зберегти як PDF» (v7.600; логотип вбудовано). Бекенд не потрібен. */
(function(root){
  'use strict';
  var MON = ['січень', 'лютий', 'березень', 'квітень', 'травень', 'червень', 'липень', 'серпень', 'вересень', 'жовтень', 'листопад', 'грудень'];
  var MON_G = ['січня', 'лютого', 'березня', 'квітня', 'травня', 'червня', 'липня', 'серпня', 'вересня', 'жовтня', 'листопада', 'грудня'];
  var POINT = {'09': 'Вересень', '01': 'Січень', '06': 'Червень'};
  var BRAND = {orange: '#F26E33', orangeSoft: '#FFF1E8', green: '#22963F', greenSoft: '#EAF6EC', ink: '#3B3A40', muted: '#7C7A85', line: '#EFE7E1', amber: '#E9A23B', amberSoft: '#FDF3E1'};
  function esc(s){ return String(s == null ? '' : s).replace(/[&<>"']/g, function(c){ return {'&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'}[c]; }); }
  function pad(n){ return ('0' + n).slice(-2); }
  function ymOf(d){ return d.getFullYear() + '-' + pad(d.getMonth() + 1); }
  function mark(v){ v = String(v == null ? '' : v).trim(); return v === '+' ? '+' : (v === '-' || v === '−' || v === '–') ? '-' : ''; }
  function sectionIcon(title){
    var t = String(title || '').toLowerCase();
    var M = [[/емоці/, '💛'], [/соціал/, '👫'], [/конфлікт/, '🕊️'], [/самостій/, '🧭'], [/мовлен/, '💬'], [/сенсор|пізнав/, '🔍'], [/ігров/, '🧸'], [/предметно|практич/, '✂️'], [/художн|естетич/, '🎨'], [/рух|фізич/, '🤸']];
    for (var i = 0; i < M.length; i++) if (M[i][0].test(t)) return M[i][1];
    return '🌱';
  }
  function level(pct){
    if (pct >= 70) return {key: 'good', label: 'Упевнено', color: BRAND.green, soft: BRAND.greenSoft};
    if (pct >= 40) return {key: 'grow', label: 'Активно розвивається', color: BRAND.amber, soft: BRAND.amberSoft};
    return {key: 'learn', label: 'Вчиться, ми поруч', color: BRAND.orange, soft: BRAND.orangeSoft};
  }
  // Період: {kind:'month'|'quarter'|'range', from:'РРРР-ММ', to:'РРРР-ММ'}
  function periodOf(kind, anchorYm, toYm){
    var y = +String(anchorYm).slice(0, 4), m = +String(anchorYm).slice(5, 7);
    if (kind === 'month') return {kind: kind, from: anchorYm, to: anchorYm};
    if (kind === 'quarter'){ var q0 = Math.floor((m - 1) / 3) * 3 + 1; return {kind: kind, from: y + '-' + pad(q0), to: y + '-' + pad(q0 + 2)}; }
    var a = String(anchorYm), b = String(toYm || anchorYm); return a <= b ? {kind: 'range', from: a, to: b} : {kind: 'range', from: b, to: a};
  }
  function periodLabel(p){
    var f = function(ym){ return MON[+ym.slice(5, 7) - 1] + ' ' + ym.slice(0, 4); };
    if (p.from === p.to) return f(p.from);
    if (p.kind === 'quarter') return Math.ceil(+p.from.slice(5, 7) / 3) + '-й квартал ' + p.from.slice(0, 4) + ' (' + MON[+p.from.slice(5, 7) - 1] + ' – ' + MON[+p.to.slice(5, 7) - 1] + ')';
    return p.from.slice(0, 4) === p.to.slice(0, 4) ? MON[+p.from.slice(5, 7) - 1] + ' – ' + f(p.to) : f(p.from) + ' – ' + f(p.to);
  }
  // Спостереження за період: точки з хоч однією позначкою, від старшої до новішої
  function observations(dev, p){
    dev = dev && typeof dev === 'object' ? dev : {};
    return Object.keys(dev).filter(function(k){ return /^\d{4}-\d{2}$/.test(k) && k >= p.from && k <= p.to; }).sort()
      .map(function(k){ var cr = (dev[k] || {}).criteria || {}, c = {}; Object.keys(cr).forEach(function(id){ var v = mark(cr[id]); if (v) c[id] = v; });
        return {key: k, label: (POINT[k.slice(5, 7)] || MON[+k.slice(5, 7) - 1]) + ' ' + k.slice(0, 4), criteria: c, by: (dev[k] || {}).by || '', at: (dev[k] || {}).at || ''}; })
      .filter(function(o){ return Object.keys(o.criteria).length > 0; });
  }
  function hasData(dev, p){ return observations(dev, p).length > 0; }
  // Права: CFO (і власник) — усі; директор — свої локації; вихователь — свої локації (облікові записи вихователів — на локацію, без групи),
  // а якщо в записі є група(и) — лише свої групи
  function canReport(user, card){
    if (!user || !card) return false;
    var r = String(user.role || '').toLowerCase();
    if (r === 'cfo' || r === 'founder' || Number(user.id) === 1) return true;
    if (r !== 'director' && r !== 'vyhovatel') return false;
    var locs = (Array.isArray(user.locs) && user.locs.length ? user.locs : String(user.loc || '').split(',')).map(function(x){ return String(x).trim(); }).filter(Boolean);
    var loc = String(card.loc || card['Локація'] || '').trim();
    if (locs.indexOf(loc) < 0) return false;
    if (r === 'vyhovatel'){
      var gs = [].concat(user.groups || (user.group ? String(user.group).split(',') : [])).map(function(x){ return String(x).trim().toLowerCase(); }).filter(Boolean);
      if (gs.length && gs.indexOf(String(card.group || card['Група'] || '').trim().toLowerCase()) < 0) return false;
    }
    return true;
  }
  // Дані звіту (без вигадок): розділи з позначками останнього спостереження, динаміка між першим і останнім спостереженням періоду
  function buildData(card, template, p, now){
    var obs = observations(card.development, p); if (!obs.length) return null;
    var last = obs[obs.length - 1], first = obs[0], multi = obs.length > 1;
    var sections = [];
    Object.keys(template || {}).forEach(function(sid){
      var S = template[sid], items = S.items || [];
      var marked = items.filter(function(it){ return last.criteria[it.id]; });
      if (!marked.length && !(multi && items.some(function(it){ return obs.some(function(o){ return o.criteria[it.id]; }); }))) return;
      var pct = function(o){ var n = 0; items.forEach(function(it){ if (o.criteria[it.id] === '+') n++; }); return items.length ? Math.round(n / items.length * 100) : 0; };
      var any = function(o){ return items.some(function(it){ return o.criteria[it.id]; }); };
      var trend = multi ? obs.filter(any).map(function(o){ return {label: o.label, short: MON[+o.key.slice(5, 7) - 1].slice(0, 3) + ' ' + o.key.slice(2, 4), pct: pct(o)}; }) : [];
      var newSkills = multi && any(first) ? items.filter(function(it){ return first.criteria[it.id] !== '+' && last.criteria[it.id] === '+'; }).map(function(it){ return it.text; }) : [];
      sections.push({id: sid, title: S.title, icon: sectionIcon(S.title), pct: pct(last), level: level(pct(last)), total: items.length,
        can: marked.filter(function(it){ return last.criteria[it.id] === '+'; }).map(function(it){ return it.text; }),
        learn: marked.filter(function(it){ return last.criteria[it.id] === '-'; }).map(function(it){ return it.text; }),
        trend: trend.length > 1 ? trend : [], newSkills: newSkills});
    });
    if (!sections.length) return null;
    var d = now || new Date();
    return {child: String(card.name || card['ПІБ дитини'] || '').trim(), loc: String(card.loc || card['Локація'] || '').trim(), group: String(card.group || card['Група'] || '').trim(),
      teacher: String(card.teacher || card['Вихователь'] || '').trim().replace(/^\d+[\s.)-]+/, ''),   // «2 Наталія» у картці → «Наталія»
      period: periodLabel(p), periodKey: p, observations: obs.map(function(o){ return o.label; }),
      lastLabel: last.label, multi: multi, sections: sections, generated: d.getDate() + ' ' + MON_G[d.getMonth()] + ' ' + d.getFullYear()};
  }
  function bar(pct, color){
    return '<div style="height:10px;border-radius:6px;background:#F3EEEA;overflow:hidden"><div style="height:10px;width:' + Math.max(pct, 3) + '%;background:' + color + ';border-radius:6px"></div></div>';
  }
  function list(arr, dot, color){
    return arr.map(function(t){ return '<div style="display:flex;gap:6px;align-items:flex-start;margin:2px 0;font-size:11px;line-height:1.35"><span style="flex:0 0 auto;width:13px;height:13px;border-radius:50%;background:' + color + ';color:#fff;font-size:8.5px;font-weight:700;display:inline-flex;align-items:center;justify-content:center;margin-top:1px">' + dot + '</span><span>' + esc(t) + '</span></div>'; }).join('');
  }
  // HTML звіту (A4, 794px — так само рендериться і в html2pdf, і в «друк у PDF»)
  function renderHtml(D, logoSrc){
    var B = BRAND, h = '';
    h += '<div class="mkr" style="width:794px;box-sizing:border-box;padding:26px 38px 22px;font-family:\'Nunito\',\'Segoe UI\',\'Helvetica Neue\',Arial,sans-serif;color:' + B.ink + ';background:#fff">';
    h += '<div style="display:flex;align-items:center;gap:18px;padding-bottom:18px;border-bottom:3px solid ' + B.orange + '">' +
      (logoSrc ? '<img src="' + esc(logoSrc) + '" style="width:78px;height:78px;border-radius:50%">' : '') +
      '<div style="flex:1"><div style="font-size:12px;letter-spacing:.08em;text-transform:uppercase;color:' + B.orange + ';font-weight:800">m.kids · ' + esc(D.loc) + '</div>' +
      '<div style="font-size:27px;font-weight:800;margin:4px 0 2px;line-height:1.15">' + esc(D.child) + '</div>' +
      '<div style="font-size:13px;color:' + B.muted + '">Звіт про розвиток для родини</div></div>' +
      '<div style="text-align:right;font-size:12px;line-height:1.7;color:' + B.muted + '">' +
      (D.group ? '<div>Група: <b style="color:' + B.ink + '">' + esc(D.group) + '</b></div>' : '') +
      '<div>Період: <b style="color:' + B.ink + '">' + esc(D.period) + '</b></div>' +
      (D.teacher ? '<div>Вихователь: <b style="color:' + B.ink + '">' + esc(D.teacher) + '</b></div>' : '') + '</div></div>';
    h += '<div style="margin:18px 0 6px;padding:14px 18px;background:' + B.orangeSoft + ';border-radius:14px;font-size:13px;line-height:1.55">' +
      'Ми уважно спостерігаємо, як дитина грає, спілкується, творить і пізнає світ. Нижче — що вже виходить упевнено і над чим працюємо разом. ' +
      (D.multi ? 'За цей період у нас ' + D.observations.length + (D.observations.length < 5 ? ' спостереження (' : ' спостережень (') + esc(D.observations.join(' → ')) + '), тож видно і динаміку.' : 'Спостереження: ' + esc(D.lastLabel) + '.') + '</div>';
    h += '<div style="display:flex;gap:16px;margin:10px 0 14px;font-size:11.5px;color:' + B.muted + '">' +
      [level(80), level(50), level(10)].map(function(l){ return '<span style="display:inline-flex;align-items:center;gap:6px"><span style="width:12px;height:12px;border-radius:50%;background:' + l.color + '"></span>' + l.label + '</span>'; }).join('') + '</div>';
    // v7.600: розділи у дві колонки — повний звіт (9 розділів × 5 пунктів) вміщується в 1–2 аркуші A4
    h += '<div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;align-items:start">';
    D.sections.forEach(function(s){
      h += '<div class="mkr-sec" style="page-break-inside:avoid;break-inside:avoid;border:1px solid ' + B.line + ';border-radius:14px;padding:11px 13px 9px;background:#fff">' +
        '<div style="display:flex;align-items:center;gap:9px;margin-bottom:6px"><span style="flex:0 0 auto;font-size:18px;width:30px;height:30px;border-radius:9px;background:' + s.level.soft + ';display:inline-flex;align-items:center;justify-content:center">' + s.icon + '</span>' +
        '<div style="font-size:13.5px;font-weight:800;line-height:1.2">' + esc(s.title) + '</div></div>' +
        '<div style="display:flex;align-items:center;gap:8px;margin:0 0 5px"><div style="flex:1">' + bar(s.pct, s.level.color).replace(/height:10px/g, 'height:8px') + '</div>' +
        '<span style="font-size:11px;font-weight:700;color:' + s.level.color + ';white-space:nowrap">' + s.level.label + '</span></div>';
      if (s.trend.length) h += '<div style="font-size:10.5px;margin:0 0 5px;color:' + B.muted + '">Було → стало: ' + s.trend.map(function(t){ return '<b style="color:' + level(t.pct).color + '">' + esc(t.short) + '</b>&nbsp;<span style="display:inline-block;width:30px;vertical-align:middle">' + bar(t.pct, level(t.pct).color).replace(/height:10px/g, 'height:5px') + '</span>'; }).join(' <span style="color:' + B.orange + '">→</span> ') + '</div>';
      if (s.newSkills.length) h += '<div style="margin:0 0 5px;padding:6px 9px;background:' + B.greenSoft + ';border-radius:8px;font-size:11px;line-height:1.35"><b style="color:' + B.green + '">✨ Нове за період:</b> ' + esc(s.newSkills.map(function(t){ return String(t).replace(/[.\s]+$/, ''); }).join('; ')) + '.' + '</div>';
      h += (s.can.length ? '<div style="font-size:10px;font-weight:800;color:' + B.green + ';margin:3px 0 1px;text-transform:uppercase;letter-spacing:.05em">Уже вміє</div>' + list(s.can, '✓', B.green) : '') +
        (s.learn.length ? '<div style="font-size:10px;font-weight:800;color:' + B.orange + ';margin:5px 0 1px;text-transform:uppercase;letter-spacing:.05em">Вчимося разом</div>' + list(s.learn, '•', B.orange) : '') + '</div>';
    });
    h += '</div>';
    h += '<div style="page-break-inside:avoid;break-inside:avoid;margin-top:14px;padding-top:12px;border-top:2px dashed ' + B.line + ';display:flex;justify-content:space-between;align-items:flex-end;font-size:12px;color:' + B.muted + '">' +
      '<div>Звіт сформовано ' + esc(D.generated) + '<br>за спостереженнями вихователів m.kids</div>' +
      '<div style="text-align:right;font-size:15px;color:' + B.orange + ';font-weight:800">з турботою,<br>команда m.kids 🧡</div></div>';
    return h + '</div>';
  }
  // v7.600: окрема сторінка звіту (нова вкладка) → повна промальовка (шрифти, логотип) → window.print() «Зберегти як PDF» — текст у PDF виділяється.
  // Раніше html2pdf знімав картинку елемента поза екраном: при прокрученій сторінці (картка дитини відкрита нижче) знімок виходив порожнім.
  function pageHtml(D, logoSrc){
    var title = fileName(D).replace(/\.pdf$/, '');
    return '<!doctype html><html lang="uk"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>' + esc(title) + '</title><style>' +
      '@page{size:A4;margin:8mm 0 10mm}html,body{margin:0}body{background:#F6F1EC;-webkit-print-color-adjust:exact;print-color-adjust:exact}' +
      '#mkrBar{position:sticky;top:0;z-index:5;display:flex;gap:10px;align-items:center;justify-content:center;flex-wrap:wrap;padding:10px 12px;background:#fff;border-bottom:1px solid #EFE7E1;font-family:-apple-system,Segoe UI,Arial,sans-serif;font-size:14px;color:#7C7A85}' +
      '#mkrBar button{font:inherit;font-weight:700;border:0;border-radius:10px;padding:9px 16px;background:#F26E33;color:#fff;cursor:pointer}#mkrBar button:disabled{opacity:.5}' +
      '#mkrPage{width:794px;margin:16px auto;background:#fff;box-shadow:0 4px 24px rgba(0,0,0,.08);border-radius:6px;overflow:hidden}' +
      '@media print{#mkrBar{display:none}body{background:#fff}#mkrPage{margin:0;box-shadow:none;border-radius:0;overflow:visible;zoom:1!important}}' +
      '</style></head><body><div id="mkrBar"><span id="mkrSt">⏳ Готую звіт…</span><button id="mkrPrint" disabled>🖨️ Зберегти як PDF</button></div>' +
      '<div id="mkrPage">' + renderHtml(D, logoSrc) + '</div><script>(' + pageScript.toString() + ')();<' + '/script></body></html>';
  }
  // Виконується в новій вкладці: чекаємо шрифти й зображення, далі — друк (на телефоні — кнопка «Поділитися / Зберегти PDF»)
  function pageScript(){
    var mob = /Android|iPhone|iPad|iPod|Mobile/i.test(navigator.userAgent) || (navigator.maxTouchPoints > 1 && /Mac/.test(navigator.platform) && !/Chrome/.test(navigator.userAgent));
    var pg = document.getElementById('mkrPage'), bt = document.getElementById('mkrPrint'), st = document.getElementById('mkrSt');
    function fit(){ var w = window.innerWidth - 16; pg.style.zoom = w < 794 ? String(Math.max(w, 280) / 794) : ''; }
    fit(); window.addEventListener('resize', fit);
    if (mob) bt.textContent = '📤 Поділитися / Зберегти PDF';
    var imgs = Array.prototype.slice.call(document.images).map(function(im){ return im.decode ? im.decode().catch(function(){}) : (im.complete ? 0 : new Promise(function(ok){ im.onload = im.onerror = ok; })); });
    var fonts = document.fonts && document.fonts.ready ? document.fonts.ready : Promise.resolve();
    // шрифти й логотип; запобіжник 4 с (фонова вкладка може не віддавати кадри / fonts.ready) — потім ще один такт на розкладку
    var all = Promise.all([fonts].concat(imgs)), cap = new Promise(function(ok){ setTimeout(ok, 4000); });
    Promise.race([all, cap]).then(function(){ return new Promise(function(ok){ setTimeout(ok, 60); }); }).then(function(){
      document.documentElement.setAttribute('data-ready', '1');
      bt.disabled = false; st.textContent = mob ? 'Звіт готовий — збережи або надішли батькам' : 'Звіт готовий. У вікні друку обери «Зберегти як PDF»';
      bt.onclick = function(){ window.print(); };
      if (!mob && !/HeadlessChrome/.test(navigator.userAgent)) setTimeout(function(){ window.print(); }, 150);
    });
  }
  var LOGO = 'data:image/png;base64,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';
  function fileName(D){ return ('Звіт_' + D.child + '_' + D.periodKey.from + (D.periodKey.to !== D.periodKey.from ? '_' + D.periodKey.to : '')).replace(/[\\\/:*?"<>|]+/g, '').replace(/\s+/g, '_') + '.pdf'; }

  var API = {pageHtml: pageHtml, LOGO: LOGO, observations: observations, hasData: hasData, canReport: canReport, buildData: buildData, renderHtml: renderHtml, periodOf: periodOf, periodLabel: periodLabel, level: level, sectionIcon: sectionIcon, fileName: fileName, mark: mark};
  if (typeof module !== 'undefined' && module.exports) module.exports = API;
  root.DevReport = API;
  if (typeof document === 'undefined') return;

  // ── UI (clients.html): кнопка у вкладці «Розвиток» → вибір періоду → PDF ──
  function cardFor(id){ var c = typeof root.getCRMCard === 'function' ? root.getCRMCard(id) : null; return c; }
  function tplFor(c){ var T = root.DEVELOPMENT_TEMPLATES || {}; return T[typeof root.devTemplateKey === 'function' ? root.devTemplateKey(c && c.group || '') : 'study'] || T.study || {}; }
  function latestYm(c){ var ks = Object.keys((c && c.development) || {}).filter(function(k){ return /^\d{4}-\d{2}$/.test(k) && hasData(c.development, {from: k, to: k}); }).sort(); return ks.length ? ks[ks.length - 1] : ymOf(new Date()); }
  // кнопка в шапці вкладки
  root.devReportButton = function(id){
    var c = cardFor(id); if (!c || !canReport(root.S && root.S.user, c)) return '';
    return '<button type="button" class="btn btn-secondary btn-sm" onclick="devReportOpen(' + JSON.stringify(String(id)).replace(/"/g, '&quot;') + ')" title="PDF для батьків за обраний період">📄 Звіт для батьків (PDF)</button>';
  };
  root.devReportOpen = function(id){
    var c = cardFor(id); if (!c) return;
    if (!canReport(root.S && root.S.user, c)){ if (root.showFloatMsg) root.showFloatMsg('🚫 Звіт доступний вихователю групи, директору локації та CFO'); return; }
    var old = document.getElementById('devRepBox'); if (old) old.remove();
    var a = latestYm(c), box = document.createElement('div'); box.id = 'devRepBox';
    box.style.cssText = 'position:fixed;inset:0;background:rgba(20,20,30,.45);z-index:10050;display:flex;align-items:center;justify-content:center;padding:16px';
    var inp = 'padding:7px 9px;border:1px solid #cbd5e1;border-radius:8px;font-size:14px';
    box.innerHTML = '<div style="background:#fff;border-radius:14px;max-width:420px;width:100%;padding:18px 20px;box-shadow:0 10px 40px rgba(0,0,0,.2)">' +
      '<div style="font-weight:800;font-size:16px;margin-bottom:4px">📄 Звіт для батьків</div><div style="font-size:13px;color:#7f8c8d;margin-bottom:12px">' + esc(c.name || '') + '</div>' +
      '<div style="display:flex;gap:6px;margin-bottom:10px">' + [['month', 'Місяць'], ['quarter', 'Квартал'], ['range', 'Свій діапазон']].map(function(k, i){ return '<label style="flex:1;font-size:13px;border:1px solid #e5e7eb;border-radius:8px;padding:6px 8px;cursor:pointer"><input type="radio" name="devRepK" value="' + k[0] + '"' + (k[0] === 'quarter' ? ' checked' : '') + '> ' + k[1] + '</label>'; }).join('') + '</div>' +
      '<div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap"><input type="month" id="devRepA" value="' + a + '" style="' + inp + '"><span id="devRepToW" style="display:none">—&nbsp;<input type="month" id="devRepB" value="' + a + '" style="' + inp + '"></span></div>' +
      '<div id="devRepP" style="font-size:12.5px;color:#7f8c8d;margin:10px 0 14px"></div>' +
      '<div style="display:flex;gap:8px;justify-content:flex-end"><button type="button" class="btn btn-secondary btn-sm" id="devRepX">Скасувати</button><button type="button" class="btn btn-primary btn-sm" id="devRepGo">📄 Відкрити звіт (PDF)</button></div></div>';
    document.body.appendChild(box);
    var per = function(){ var k = (box.querySelector('input[name=devRepK]:checked') || {}).value || 'quarter'; document.getElementById('devRepToW').style.display = k === 'range' ? '' : 'none';
      return periodOf(k, document.getElementById('devRepA').value || a, document.getElementById('devRepB').value); };
    var upd = function(){ var p = per(), ok = hasData(c.development, p), go = document.getElementById('devRepGo'), obs = observations(c.development, p);
      go.disabled = !ok; go.title = ok ? '' : 'немає записів за період'; go.style.opacity = ok ? '' : '.5';
      document.getElementById('devRepP').innerHTML = esc(periodLabel(p)) + ' · ' + (ok ? 'спостереження: ' + esc(obs.map(function(o){ return o.label; }).join(', ')) : '<b style="color:#c0392b">немає записів за період</b>'); };
    Array.prototype.forEach.call(box.querySelectorAll('input'), function(e){ e.onchange = upd; e.oninput = upd; });
    document.getElementById('devRepX').onclick = function(){ box.remove(); };
    box.onclick = function(e){ if (e.target === box) box.remove(); };
    document.getElementById('devRepGo').onclick = function(){
      var p = per(), D = buildData({name: c.name, loc: c.loc, group: c.group, teacher: c.teacher || c.vyhovatel || '', development: c.development}, tplFor(c), p);
      if (!D){ upd(); return; }
      // вкладку відкриваємо одразу в обробнику натискання — інакше браузер блокує спливаюче вікно
      var w = root.open('', '_blank');
      if (!w){ if (root.showFloatMsg) root.showFloatMsg('⚠️ Браузер заблокував нову вкладку — дозволь спливаючі вікна для цього сайту'); return; }
      w.document.open(); w.document.write(pageHtml(D, LOGO)); w.document.close();
      box.remove();
    };
    upd();
  };
})(typeof window !== 'undefined' ? window : this);
