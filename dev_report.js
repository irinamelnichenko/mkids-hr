/* v7.599 «📄 Звіт для батьків (PDF)» — модуль «Розвиток» картки дитини (clients.html, вкладка 🌱).
   Дані — лише те, що внесено: c.development = {'РРРР-ММ': {criteria:{'1.1':'+'|'-'}, by, at}} (точки Вересень / Січень / Червень),
   шаблон критеріїв — DEVELOPMENT_TEMPLATES за групою. Коментарів і фото в «Розвитку» немає — у звіті їх теж немає.
   PDF — html2pdf.bundle.min.js (локально, вантажиться лише при натисканні). Бекенд не потрібен. */
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
    var M = [[/емоці/, '💛'], [/соціал/, '🤝'], [/конфлікт/, '🕊️'], [/самостій/, '🧭'], [/мовлен/, '💬'], [/сенсор|пізнав/, '🔍'], [/ігров/, '🧸'], [/предметно|практич/, '✂️'], [/художн|естетич/, '🎨'], [/рух|фізич/, '🤸']];
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
      teacher: String(card.teacher || card['Вихователь'] || '').trim(), period: periodLabel(p), periodKey: p, observations: obs.map(function(o){ return o.label; }),
      lastLabel: last.label, multi: multi, sections: sections, generated: d.getDate() + ' ' + MON_G[d.getMonth()] + ' ' + d.getFullYear()};
  }
  function bar(pct, color){
    return '<div style="height:10px;border-radius:6px;background:#F3EEEA;overflow:hidden"><div style="height:10px;width:' + Math.max(pct, 3) + '%;background:' + color + ';border-radius:6px"></div></div>';
  }
  function list(arr, dot, color){
    return arr.map(function(t){ return '<div style="display:flex;gap:8px;align-items:flex-start;margin:3px 0;font-size:12.5px;line-height:1.45"><span style="flex:0 0 auto;width:16px;height:16px;border-radius:50%;background:' + color + ';color:#fff;font-size:10px;font-weight:700;display:inline-flex;align-items:center;justify-content:center;margin-top:1px">' + dot + '</span><span>' + esc(t) + '</span></div>'; }).join('');
  }
  // HTML звіту (A4, 794px — так само рендериться і в html2pdf, і в «друк у PDF»)
  function renderHtml(D, logoSrc){
    var B = BRAND, h = '';
    h += '<div class="mkr" style="width:794px;box-sizing:border-box;padding:34px 46px 30px;font-family:\'Nunito\',\'Segoe UI\',\'Helvetica Neue\',Arial,sans-serif;color:' + B.ink + ';background:#fff">';
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
    D.sections.forEach(function(s){
      h += '<div class="mkr-sec" style="page-break-inside:avoid;break-inside:avoid;border:1px solid ' + B.line + ';border-radius:16px;padding:14px 18px 12px;margin:0 0 12px;background:#fff">' +
        '<div style="display:flex;align-items:center;gap:12px;margin-bottom:8px"><span style="font-size:24px;width:40px;height:40px;border-radius:12px;background:' + s.level.soft + ';display:inline-flex;align-items:center;justify-content:center">' + s.icon + '</span>' +
        '<div style="flex:1"><div style="font-size:15.5px;font-weight:800">' + esc(s.title) + '</div>' +
        '<div style="display:flex;align-items:center;gap:10px;margin-top:5px"><div style="flex:1">' + bar(s.pct, s.level.color) + '</div>' +
        '<span style="font-size:12px;font-weight:700;color:' + s.level.color + ';white-space:nowrap">' + s.level.label + '</span></div></div></div>';
      if (s.trend.length) h += '<div style="font-size:12px;margin:2px 0 8px 52px;color:' + B.muted + '">Було → стало: ' + s.trend.map(function(t){ return '<b style="color:' + level(t.pct).color + '">' + esc(t.short) + '</b> ' + '<span style="display:inline-block;width:46px;vertical-align:middle">' + bar(t.pct, level(t.pct).color).replace(/height:10px/g, 'height:6px') + '</span>'; }).join(' <span style="color:' + B.orange + '">→</span> ') + '</div>';
      if (s.newSkills.length) h += '<div style="margin:0 0 6px 52px;padding:8px 12px;background:' + B.greenSoft + ';border-radius:10px;font-size:12.5px"><b style="color:' + B.green + '">✨ Нове за період:</b> ' + esc(s.newSkills.map(function(t){ return String(t).replace(/[.\s]+$/, ''); }).join('; ')) + '.' + '</div>';
      h += '<div style="margin-left:52px">' + (s.can.length ? '<div style="font-size:11.5px;font-weight:800;color:' + B.green + ';margin:4px 0 2px;text-transform:uppercase;letter-spacing:.05em">Уже вміє</div>' + list(s.can, '✓', B.green) : '') +
        (s.learn.length ? '<div style="font-size:11.5px;font-weight:800;color:' + B.orange + ';margin:8px 0 2px;text-transform:uppercase;letter-spacing:.05em">Вчимося разом</div>' + list(s.learn, '•', B.orange) : '') + '</div></div>';
    });
    h += '<div style="page-break-inside:avoid;break-inside:avoid;margin-top:18px;padding-top:14px;border-top:2px dashed ' + B.line + ';display:flex;justify-content:space-between;align-items:flex-end;font-size:12px;color:' + B.muted + '">' +
      '<div>Звіт сформовано ' + esc(D.generated) + '<br>за спостереженнями вихователів m.kids</div>' +
      '<div style="text-align:right;font-size:15px;color:' + B.orange + ';font-weight:800">з турботою,<br>команда m.kids 🧡</div></div>';
    return h + '</div>';
  }
  function fileName(D){ return ('Звіт_' + D.child + '_' + D.periodKey.from + (D.periodKey.to !== D.periodKey.from ? '_' + D.periodKey.to : '')).replace(/[\\\/:*?"<>|]+/g, '').replace(/\s+/g, '_') + '.pdf'; }

  var API = {observations: observations, hasData: hasData, canReport: canReport, buildData: buildData, renderHtml: renderHtml, periodOf: periodOf, periodLabel: periodLabel, level: level, sectionIcon: sectionIcon, fileName: fileName, mark: mark};
  if (typeof module !== 'undefined' && module.exports) module.exports = API;
  root.DevReport = API;
  if (typeof document === 'undefined') return;

  // ── UI (clients.html): кнопка у вкладці «Розвиток» → вибір періоду → PDF ──
  var LIB = 'html2pdf.bundle.min.js';
  function loadLib(){ return new Promise(function(ok, bad){ if (root.html2pdf) return ok(); var s = document.createElement('script'); s.src = LIB; s.onload = function(){ ok(); }; s.onerror = function(){ bad(new Error('не вдалося завантажити ' + LIB)); }; document.head.appendChild(s); }); }
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
      '<div style="display:flex;gap:8px;justify-content:flex-end"><button type="button" class="btn btn-secondary btn-sm" id="devRepX">Скасувати</button><button type="button" class="btn btn-primary btn-sm" id="devRepGo">⬇️ Завантажити PDF</button></div></div>';
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
      var btn = this; btn.disabled = true; btn.textContent = '⏳ Формую…';
      var host = document.createElement('div'); host.style.cssText = 'position:fixed;left:-10000px;top:0'; host.innerHTML = renderHtml(D, 'mkids-logo.png'); document.body.appendChild(host);
      loadLib().then(function(){
        return root.html2pdf().set({margin: [6, 0, 8, 0], filename: fileName(D), image: {type: 'jpeg', quality: 0.95}, html2canvas: {scale: 2, useCORS: true, backgroundColor: '#ffffff'},
          jsPDF: {unit: 'mm', format: 'a4', orientation: 'portrait'}, pagebreak: {mode: ['css', 'legacy'], avoid: '.mkr-sec'}}).from(host.firstChild).save();
      }).then(function(){ box.remove(); if (root.showFloatMsg) root.showFloatMsg('✅ Звіт завантажено'); })
        .catch(function(e){ btn.disabled = false; btn.textContent = '⬇️ Завантажити PDF'; if (root.showFloatMsg) root.showFloatMsg('⚠️ ' + (e && e.message || e)); })
        .then(function(){ host.remove(); });
    };
    upd();
  };
})(typeof window !== 'undefined' ? window : this);
