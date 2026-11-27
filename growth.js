// v7.602 «📏 Вага і зріст» дитини: вкладка картки (історія + внесення), список «потрібно виміряти» для медсестри,
// «💬 Надіслати батькам» (Viber deep link на номер з картки + «Копіювати текст»), позначка «надіслано» (дата, хто).
// Вносять: медсестра (свої локації), директор (своя), CFO. Бачать усі, хто бачить картку. Жодних норм чи оцінок — лише внесені цифри.
(function(root){
  'use strict';
  var EDIT_ROLES = ['nurse', 'director', 'cfo'];
  var LIVE = ['active', 'adaptation'];
  function esc(s){ return String(s == null ? '' : s).replace(/[&<>"']/g, function(c){ return {'&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'}[c]; }); }
  function pad(n){ return ('0' + n).slice(-2); }
  function isoOf(d){ return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate()); }
  function num(v){ var s = String(v == null ? '' : v).trim().replace(',', '.'); if (!/^\d{1,3}(\.\d+)?$/.test(s)) return NaN; return Math.round(parseFloat(s) * 10) / 10; }
  // 15.9 → «15,9»; вага завжди з 1 знаком, зріст — без «,0»
  function fmtW(n){ return (n == null || isNaN(n)) ? '—' : Number(n).toFixed(1).replace('.', ','); }
  function fmtH(n){ return (n == null || isNaN(n)) ? '—' : (Math.round(n) === n ? String(n) : Number(n).toFixed(1).replace('.', ',')); }
  function fmtDate(iso){ var m = String(iso || '').match(/^(\d{4})-(\d{2})-(\d{2})/); return m ? m[3] + '.' + m[2] + '.' + m[1] : String(iso || ''); }
  function validate(o, todayIso){
    var d = String((o && o.date) || '').slice(0, 10), w = num(o && o.weight), h = num(o && o.height);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(d)) return {error: 'Вкажіть дату виміру'};
    if (todayIso && d > todayIso) return {error: 'Дата виміру — не з майбутнього'};
    if (!(w >= 1 && w < 100)) return {error: 'Вага — у кг, напр. 15,9'};
    if (!(h >= 30 && h < 200)) return {error: 'Зріст — у см, напр. 96'};
    return {date: d, weight: w, height: h};
  }
  // нові виміри + старі квартальні заміри — одна історія, найновіші зверху
  function merge(items, legacy){
    return (items || []).concat(legacy || []).filter(function(r){ return r && r.date; })
      .sort(function(a, b){ return a.date < b.date ? 1 : a.date > b.date ? -1 : String(b.at || '').localeCompare(String(a.at || '')); });
  }
  function last(list){ var l = merge(list); return l.length ? l[0] : null; }
  function minus3m(iso){ var p = String(iso).split('-').map(Number), d = new Date(p[0], p[1] - 1 - 3, p[2]); if (d.getDate() !== p[2]) d = new Date(p[0], p[1] - 3, 0); return isoOf(d); }
  // «потрібно виміряти»: виміру немає або останній старший за 3 місяці
  function isDue(list, todayIso){ var l = last(list); return !l || l.date < minus3m(todayIso); }
  function messageText(w, h){
    return 'Вітаю, рада сповістити, що малеча фізично розвивається та зростає відповідно графіку на сьогодні:\nВага: ' + fmtW(w) + ' кг\nЗріст: ' + fmtH(h) + ' см';
  }
  function normPhone(raw){
    var d = String(raw || '').replace(/\D/g, '');
    if (d.length === 10 && d.charAt(0) === '0') d = '38' + d; else if (d.length === 9) d = '380' + d; else if (d.length === 11 && d.slice(0, 2) === '80') d = '3' + d;
    return d.length >= 11 ? d : '';
  }
  // номер батьків з картки: підписант договору, інакше другий з батьків
  function parentPhone(c){
    c = c || {}; var dad = c.signerParent === 'dad', a = dad ? c.dadPhone : c.momPhone, b = dad ? c.momPhone : c.dadPhone;
    var pa = normPhone(a); if (pa) return {intl: pa, raw: String(a), who: dad ? 'тато' : 'мама'};
    var pb = normPhone(b); return pb ? {intl: pb, raw: String(b), who: dad ? 'мама' : 'тато'} : null;
  }
  function viberUrl(intl, text){ return 'viber://chat?number=%2B' + intl + '&text=' + encodeURIComponent(text); }
  function role(u){ return String((u && u.role) || '').toLowerCase().trim(); }
  function locs(u){ u = u || {}; var a = Array.isArray(u.locs) && u.locs.length ? u.locs : String(u.loc || '').split(/\s*[,;]\s*/); return a.map(function(x){ return String(x || '').trim(); }).filter(Boolean); }
  function sameLoc(a, b){ return String(a || '').trim().toLowerCase().replace(/[’ʼ`]/g, "'") === String(b || '').trim().toLowerCase().replace(/[’ʼ`]/g, "'"); }
  function canEdit(u, loc){
    var r = role(u); if (EDIT_ROLES.indexOf(r) < 0) return false;
    if (r === 'cfo') return true;
    return locs(u).some(function(l){ return sameLoc(l, loc); });
  }
  function canList(u){ return EDIT_ROLES.indexOf(role(u)) >= 0; }
  // останній вимір за період звіту (місяці 'YYYY-MM' включно) — для «📄 Звіт для батьків»
  function inPeriod(list, p){ var l = merge(list).filter(function(r){ var m = r.date.slice(0, 7); return m >= p.from && m <= p.to && (r.weight > 0 || r.height > 0); }); return l.length ? l[0] : null; }

  // v7.603: рядок для журналу «було → стало» (як у бекенді _growthStr)
  function describe(it){ return it ? fmtDate(it.date) + ' · ' + fmtW(it.weight > 0 ? it.weight : null) + ' кг · ' + fmtH(it.height > 0 ? it.height : null) + ' см' : ''; }
  var API = {describe: describe, num: num, fmtW: fmtW, fmtH: fmtH, fmtDate: fmtDate, validate: validate, merge: merge, last: last, minus3m: minus3m, isDue: isDue,
             messageText: messageText, normPhone: normPhone, parentPhone: parentPhone, viberUrl: viberUrl, canEdit: canEdit, canList: canList, inPeriod: inPeriod, LIVE: LIVE};
  if (typeof module !== 'undefined' && module.exports) module.exports = API;
  root.MkGrowth = API;
  if (typeof document === 'undefined') return;

  // ── UI ──────────────────────────────────────────────────────────────────────
  var CACHE = {};   // childId → злитий список
  var LOG = {};     // v7.603: childId → журнал змін
  var EDITING = ''; // v7.603: id виміру, що зараз редагується у вкладці
  function S(){ return root.S || {}; }
  function user(){ return S().user || null; }
  function todayIso(){ return isoOf(new Date()); }
  function actorName(){ var u = user() || {}; return String(u.name || u.fullName || u.role || ''); }
  function card(id){ return typeof root.getCRMCard === 'function' ? root.getCRMCard(id) : ((S().crm || {})[id] || null); }
  function toast(m){ if (root.showFloatMsg) root.showFloatMsg(m); }
  function apiGet(params){ return root.apiGetFull(params); }
  function apiPost(body){ return root.apiPost(body); }

  API.load = function(childId, force){
    if (!force && CACHE[childId]) return Promise.resolve(CACHE[childId]);
    return apiGet({action: 'getGrowth', childId: childId}).then(function(j){
      CACHE[childId] = merge(j.items, (j.legacy || {})[childId]); LOG[childId] = j.log || []; return CACHE[childId];
    });
  };
  API.cached = function(childId){ return CACHE[childId] || null; };

  function save(childId, c, vals){
    var v = validate(vals, todayIso()); if (v.error) return Promise.reject(new Error(v.error));
    var u = user() || {};
    return apiPost({action: 'addGrowth', childId: childId, name: c.name || '', loc: c.loc || '', date: v.date, weight: v.weight, height: v.height,
                    actorId: u.id, actorName: actorName()}).then(function(j){
      if (!j) throw new Error('немає зв\'язку з хмарою');
      if (j.ok === false) throw new Error(j.error || 'Помилка сервера');
      var it = j.item || {id: '', childId: childId, date: v.date, weight: v.weight, height: v.height, by: actorName(), sent: []};
      CACHE[childId] = merge([it], (CACHE[childId] || []).filter(function(r){ return r.id !== it.id; }));
      return it;
    });
  }

  // ── модалка «💬 Надіслати батькам» ──
  API.sendOpen = function(childId, itemId){
    var c = card(childId) || {}, list = CACHE[childId] || [], it = list.filter(function(r){ return r.id === itemId; })[0] || (itemId ? null : last(list));
    if (!it){ toast('⚠️ Вимір не знайдено'); return; }
    var text = messageText(it.weight, it.height), ph = parentPhone(c);
    var old = document.getElementById('grSendBox'); if (old) old.remove();
    var box = document.createElement('div'); box.id = 'grSendBox';
    box.style.cssText = 'position:fixed;inset:0;background:rgba(20,20,30,.45);z-index:10060;display:flex;align-items:center;justify-content:center;padding:16px';
    var btn = 'flex:1;text-align:center;padding:10px 12px;border-radius:9px;border:1px solid #d0d0d0;background:#fff;cursor:pointer;font-size:14px;font-weight:600;text-decoration:none;color:#222;display:inline-block';
    box.innerHTML = '<div style="background:#fff;border-radius:14px;max-width:480px;width:100%;padding:18px 20px;box-shadow:0 10px 40px rgba(0,0,0,.2)">' +
      '<div style="font-weight:800;font-size:16px;display:flex;justify-content:space-between">💬 Надіслати батькам<span id="grSendX" style="cursor:pointer;color:#999;font-size:22px;line-height:1">×</span></div>' +
      '<div style="font-size:13px;color:#7f8c8d;margin:4px 0 10px">' + esc(c.name || '') + (ph ? ' · ' + esc(ph.who) + ' ' + esc(ph.raw) : ' · <b style="color:#c0392b">номера батьків у картці немає — лише копіювання</b>') + '</div>' +
      '<textarea readonly style="width:100%;box-sizing:border-box;height:110px;font:inherit;font-size:13px;line-height:1.45;border:1px solid #ddd;border-radius:8px;padding:10px;resize:vertical">' + esc(text) + '</textarea>' +
      '<div style="display:flex;gap:8px;margin-top:10px"><button type="button" id="grCopy" style="' + btn + '">📋 Копіювати текст</button>' +
      (ph ? '<a id="grViber" href="' + esc(viberUrl(ph.intl, text)) + '" style="' + btn + ';background:#7360f2;color:#fff;border-color:#7360f2">💬 Надіслати у Viber</a>' : '') + '</div>' +
      '<div style="display:flex;gap:8px;justify-content:flex-end;margin-top:10px"><button type="button" id="grMark" style="' + btn + ';flex:0 0 auto;background:#1faa59;color:#fff;border-color:#1faa59">✓ Позначити надісланим</button></div>' +
      '<div style="font-size:11.5px;color:#7f8c8d;margin-top:8px">На телефоні текст підставиться у Viber сам; на комп\'ютері — скопіюй і встав. Натискання «Надіслати у Viber» теж ставить позначку.</div></div>';
    document.body.appendChild(box);
    var close = function(){ box.remove(); };
    document.getElementById('grSendX').onclick = close; box.onclick = function(e){ if (e.target === box) close(); };
    document.getElementById('grCopy').onclick = function(){
      var b = this, done = function(){ b.textContent = '✓ Скопійовано'; setTimeout(function(){ b.textContent = '📋 Копіювати текст'; }, 1800); };
      if (navigator.clipboard && navigator.clipboard.writeText) navigator.clipboard.writeText(text).then(done, function(){ fallback(); });
      else fallback();
      function fallback(){ var ta = box.querySelector('textarea'); ta.removeAttribute('readonly'); ta.select(); try { document.execCommand('copy'); done(); } catch(e){} ta.setAttribute('readonly', ''); }
    };
    var marked = false;
    var mark = function(){
      if (marked) return; marked = true;
      var rec = {at: 'щойно', by: actorName()}; it.sent = (it.sent || []).concat([rec]);
      if (it.id) apiPost({action: 'markGrowthSent', id: it.id, actorId: (user() || {}).id, actorName: actorName()}).then(function(j){ if (j && j.sent) it.sent = j.sent; rerender(childId); });
      toast('✓ Позначено надісланим'); rerender(childId);
    };
    document.getElementById('grMark').onclick = function(){ mark(); close(); };
    var vb = document.getElementById('grViber'); if (vb) vb.addEventListener('click', function(){ mark(); });
  };
  function rerender(childId){
    if (S().editId === childId && document.getElementById('tpGrowth') && document.getElementById('tpGrowth').classList.contains('active')) API.renderTab();
    var row = document.querySelector('[data-gr-row="' + cssId(childId) + '"]'); if (row && row._refresh) row._refresh();
  }
  function cssId(id){ return String(id).replace(/["\\]/g, '_'); }
  function sentLine(it){ var s = it.sent || []; if (!s.length) return ''; var x = s[s.length - 1]; return '✓ ' + esc(x.at === 'щойно' ? 'щойно' : String(x.at || '').slice(0, 16)) + (x.by ? ' · ' + esc(x.by) : '') + (s.length > 1 ? ' (×' + s.length + ')' : ''); }
  function inputs(pref, d){
    var inp = 'padding:7px 9px;border:1px solid #cbd5e1;border-radius:8px;font-size:14px';
    return '<input type="date" id="' + pref + 'D" value="' + esc(d || todayIso()) + '" max="' + todayIso() + '" style="' + inp + '">' +
      '<input id="' + pref + 'W" inputmode="decimal" placeholder="вага, кг" style="' + inp + ';width:92px">' +
      '<input id="' + pref + 'H" inputmode="decimal" placeholder="зріст, см" style="' + inp + ';width:92px">';
  }
  function readInputs(pref){ var g = function(k){ var e = document.getElementById(pref + k); return e ? e.value : ''; }; return {date: g('D'), weight: g('W'), height: g('H')}; }

  // ── вкладка картки «📏 Вага і зріст» ──
  API.renderTab = function(force){
    var id = S().editId, tp = document.getElementById('tpGrowth'); if (!id || !tp) return;
    if (force === true) delete CACHE[id];   // відкриття вкладки — свіжі дані (виміри з інших пристроїв)
    var c = card(id) || {}, list = CACHE[id];
    if (!list){
      tp.innerHTML = '<div class="loading"><div class="spinner"></div><p>Завантаження…</p></div>';
      API.load(id).then(function(){ if (S().editId === id) API.renderTab(); }, function(e){ tp.innerHTML = '<div class="empty"><span class="ei">⚠️</span><h4>Не вдалось завантажити</h4><p>' + esc(e && e.message || e) + '</p><button class="btn btn-secondary btn-sm" onclick="MkGrowth.renderTab()">🔄 Ще раз</button></div>'; });
      return;
    }
    var edit = canEdit(user(), c.loc), lst = last(list), due = isDue(list, todayIso());
    var h = '<div class="card" style="margin-bottom:12px"><div class="card-hd" style="gap:10px;flex-wrap:wrap"><h3>📏 Вага і зріст</h3>' +
      '<span class="badge ' + (due ? 'bo' : 'bg') + '" style="margin-left:auto">' + (lst ? 'Останній вимір ' + fmtDate(lst.date) : 'Вимірів ще немає') + (due ? ' · потрібно виміряти' : '') + '</span></div>';
    if (edit) h += '<div style="padding:12px 14px;display:flex;gap:8px;flex-wrap:wrap;align-items:center">' + inputs('grT') +
      '<button class="btn btn-primary btn-sm" id="grTSave">💾 Зберегти</button><span id="grTMsg" style="font-size:13px"></span></div>';
    h += '</div>';
    if (!list.length) h += '<p style="color:var(--gray);font-size:13px">Записів немає</p>';
    else {
      h += '<div class="tbl-wrap"><table><thead><tr><th>Дата</th><th>Вага, кг</th><th>Зріст, см</th><th>Хто вніс</th><th>Надіслано батькам</th>' + (edit ? '<th></th>' : '') + '</tr></thead><tbody>';
      list.forEach(function(r){
        var own = edit && !r.legacy && r.id;
        if (own && EDITING === r.id){   // v7.603: рядок у режимі редагування
          h += '<tr style="background:#fffaf3"><td colspan="5"><span style="display:inline-flex;gap:6px;flex-wrap:wrap;align-items:center">' + inputs('grE', r.date) +
            '<button class="btn btn-primary btn-xs" id="grESave">💾 Зберегти зміни</button><button class="btn btn-secondary btn-xs" id="grECancel">Скасувати</button><span id="grEMsg" style="font-size:12px;color:#c0392b"></span></span></td><td></td></tr>';
          return;
        }
        h += '<tr><td>' + fmtDate(r.date) + (r.legacy ? ' <span style="font-size:11px;color:var(--gray)">квартальні заміри</span>' : '') + '</td><td>' + fmtW(r.weight) + '</td><td>' + fmtH(r.height) + '</td>' +
          '<td style="font-size:12px;color:var(--gray)">' + esc(r.by || '') + '</td><td style="font-size:12px;color:var(--green,#27ae60)">' + sentLine(r) + '</td>' +
          (edit ? '<td style="white-space:nowrap">' + (!r.legacy && r.weight > 0 && r.height > 0 ? '<button class="btn btn-secondary btn-xs" onclick="MkGrowth.sendOpen(' + esc(JSON.stringify(String(id))) + ',' + esc(JSON.stringify(String(r.id))) + ')">💬 Надіслати батькам</button>' : '') +
            (own ? ' <button class="btn btn-secondary btn-xs" title="Редагувати" onclick="MkGrowth.editStart(' + esc(JSON.stringify(String(r.id))) + ')">✏️</button> <button class="btn btn-secondary btn-xs" title="Видалити" onclick="MkGrowth.remove(' + esc(JSON.stringify(String(r.id))) + ')">🗑</button>' : '') + '</td>' : '') + '</tr>';
      });
      h += '</tbody></table></div>';
      if (list.some(function(r){ return r.legacy; }) && edit) h += '<div style="font-size:11.5px;color:var(--gray);margin:4px 2px">Квартальні заміри змінюються у вкладці «🩺 Здоров\'я».</div>';
    }
    var lg = LOG[id] || [];
    if (lg.length) h += '<details style="margin-top:12px"><summary style="cursor:pointer;font-size:13px;font-weight:700">🕓 Журнал змін (' + lg.length + ')</summary><div style="font-size:12.5px;margin-top:6px">' +
      lg.slice().reverse().map(function(x){ return '<div style="padding:4px 0;border-bottom:1px solid #f0f0f0">' + esc(String(x.at || '').slice(0, 16)) + ' · <b>' + esc(x.by || '') + '</b> · ' + esc(x.action || '') + ': ' + esc(x.was || '') + (x.now ? ' → ' + esc(x.now) : '') + '</div>'; }).join('') + '</div></details>';
    tp.innerHTML = h;
    if (EDITING){ var er = list.filter(function(r){ return r.id === EDITING; })[0];
      if (er){ document.getElementById('grEW').value = fmtW(er.weight); document.getElementById('grEH').value = fmtH(er.height);
        document.getElementById('grECancel').onclick = function(){ EDITING = ''; API.renderTab(); };
        document.getElementById('grESave').onclick = function(){
          var b = this, m = document.getElementById('grEMsg'), v = validate(readInputs('grE'), todayIso());
          if (v.error){ m.textContent = '⚠️ ' + v.error; return; }
          b.disabled = true; m.textContent = '⏳';
          var u = user() || {};
          apiPost({action: 'editGrowth', id: er.id, loc: er.loc || c.loc || '', date: v.date, weight: v.weight, height: v.height, actorId: u.id, actorName: actorName()}).then(function(j){
            if (!j || j.ok === false){ b.disabled = false; m.textContent = '⚠️ ' + ((j && j.error) || 'немає зв\'язку з хмарою'); return; }
            EDITING = ''; toast(j.unchanged ? 'Без змін' : '✅ Вимір змінено'); API.load(id, true).then(function(){ API.renderTab(); }, function(){ API.renderTab(); });
          }, function(e){ b.disabled = false; m.textContent = '⚠️ ' + (e && e.message || e); });
        }; } else EDITING = ''; }
    var sb = document.getElementById('grTSave');
    if (sb) sb.onclick = function(){
      var msg = document.getElementById('grTMsg'); sb.disabled = true; msg.textContent = '⏳ Зберігаю…'; msg.style.color = '';
      save(id, c, readInputs('grT')).then(function(it){
        API.renderTab();
        var m2 = document.getElementById('grTMsg');
        if (m2){ m2.innerHTML = '✅ Збережено ' + fmtW(it.weight) + ' кг · ' + fmtH(it.height) + ' см <button class="btn btn-primary btn-xs" id="grTSend">💬 Надіслати батькам</button>'; document.getElementById('grTSend').onclick = function(){ API.sendOpen(id, it.id); }; }
        if (!it.id) API.load(id, true).then(function(){ });
      }, function(e){ sb.disabled = false; msg.textContent = '⚠️ ' + (e && e.message || e); msg.style.color = '#c0392b'; });
    };
  };

  API.editStart = function(itemId){ EDITING = itemId; API.renderTab(); };
  API.remove = function(itemId){
    var id = S().editId, list = CACHE[id] || [], it = list.filter(function(r){ return r.id === itemId; })[0]; if (!it) return;
    if (!root.confirm('Видалити вимір ' + describe(it) + '?\nЗапис залишиться в журналі змін.')) return;
    var u = user() || {}, c = card(id) || {};
    apiPost({action: 'deleteGrowth', id: it.id, loc: it.loc || c.loc || '', actorId: u.id, actorName: actorName()}).then(function(j){
      if (!j || j.ok === false){ toast('⚠️ ' + ((j && j.error) || 'немає зв\'язку з хмарою')); return; }
      toast('🗑 Вимір видалено'); API.load(id, true).then(function(){ API.renderTab(); }, function(){ API.renderTab(); });
    }, function(e){ toast('⚠️ ' + (e && e.message || e)); });
  };

  // ── розділ «📏 Вага і зріст»: діти локації, яких потрібно виміряти ──
  var LIST = {loc: '', all: false, data: null, err: ''};
  function listLocs(){
    var u = user(); if (role(u) === 'cfo') return ((S().locs || []).map(function(l){ return l.loc; })).filter(Boolean);
    return locs(u);
  }
  function kidsOf(loc){
    var crm = S().crm || {};
    return Object.keys(crm).filter(function(id){ var c = crm[id]; return c && c.name && sameLoc(c.loc, loc) && LIVE.indexOf(String(c.status || 'active')) >= 0; })
      .map(function(id){ return {id: id, c: crm[id]}; })
      .sort(function(a, b){ return String(a.c.group || '').localeCompare(String(b.c.group || ''), 'uk') || String(a.c.name).localeCompare(String(b.c.name), 'uk'); });
  }
  API.renderList = function(){
    var main = document.getElementById('main'); if (!main) return;
    var u = user(); if (!canList(u)){ main.innerHTML = '<div class="empty"><span class="ei">🔒</span><h4>Немає доступу</h4></div>'; return; }
    var L = listLocs(); if (!LIST.loc || L.indexOf(LIST.loc) < 0) LIST.loc = (u && L.indexOf(u.loc) >= 0) ? u.loc : (L[0] || '');   // активна локація (locswitch) або перша
    var loc = LIST.loc;
    var h = '<div style="max-width:1000px"><p style="color:var(--gray);font-size:13.5px;max-width:80ch;margin:0 0 12px">Діти локації, у яких останній вимір старший за 3 місяці або його немає. Внеси дату, вагу (кг, напр. 15,9) і зріст (см) — «💾» або Enter, курсор перейде до наступної дитини. Після збереження — «💬 Надіслати батькам».</p>';
    h += '<div style="display:flex;gap:10px;align-items:center;flex-wrap:wrap;margin-bottom:12px">' +
      (L.length > 1 ? '<select id="grLoc" style="padding:6px 10px;border:1.5px solid #e5ded7;border-radius:8px">' + L.map(function(l){ return '<option' + (l === loc ? ' selected' : '') + '>' + esc(l) + '</option>'; }).join('') + '</select>' : '<b>' + esc(loc) + '</b>') +
      '<label style="font-size:13px"><input type="checkbox" id="grAll"' + (LIST.all ? ' checked' : '') + '> показати всіх дітей</label><span id="grCnt" style="font-size:13px;color:var(--gray)"></span></div><div id="grBody"></div></div>';
    main.innerHTML = h;
    var sel = document.getElementById('grLoc'); if (sel) sel.onchange = function(){ LIST.loc = this.value; LIST.data = null; API.renderList(); };
    document.getElementById('grAll').onchange = function(){ LIST.all = this.checked; drawRows(); };
    if (!loc){ document.getElementById('grBody').innerHTML = '<p>Немає локації</p>'; return; }
    if (LIST.data && LIST.data.loc === loc) return drawRows();
    document.getElementById('grBody').innerHTML = '<div class="loading"><div class="spinner"></div><p>Завантаження вимірів…</p></div>';
    apiGet({action: 'getGrowth', loc: loc}).then(function(j){
      var by = {}; (j.items || []).forEach(function(r){ (by[r.childId] = by[r.childId] || []).push(r); });
      Object.keys(j.legacy || {}).forEach(function(id){ by[id] = (by[id] || []).concat(j.legacy[id]); });
      Object.keys(by).forEach(function(id){ CACHE[id] = merge(by[id]); });
      LIST.data = {loc: loc, by: by}; drawRows();
    }, function(e){ document.getElementById('grBody').innerHTML = '<div class="empty"><span class="ei">⚠️</span><h4>Не вдалось завантажити</h4><p>' + esc(e && e.message || e) + '</p></div>'; });
  };
  function drawRows(){
    var body = document.getElementById('grBody'); if (!body || !LIST.data) return;
    var t = todayIso(), kids = kidsOf(LIST.loc), due = kids.filter(function(k){ return isDue(CACHE[k.id] || [], t); });
    var show = LIST.all ? kids : due;
    document.getElementById('grCnt').textContent = 'потрібно виміряти: ' + due.length + ' з ' + kids.length;
    if (!show.length){ body.innerHTML = '<div class="empty"><span class="ei">✅</span><h4>Усі діти виміряні за останні 3 місяці</h4></div>'; return; }
    body.innerHTML = '<div class="tbl-wrap"><table><thead><tr><th>Дитина</th><th>Група</th><th>Останній вимір</th><th>Новий вимір</th></tr></thead><tbody>' + show.map(function(k, i){
      return '<tr data-gr-row="' + esc(cssId(k.id)) + '" data-i="' + i + '"><td><span class="child-link" data-open="1">' + esc(k.c.name) + '</span></td><td style="font-size:12.5px">' + esc(k.c.group || '') + '</td><td class="grLast" style="font-size:12.5px"></td>' +
        '<td class="grNew" style="white-space:nowrap"></td></tr>'; }).join('') + '</tbody></table></div>';
    Array.prototype.forEach.call(body.querySelectorAll('tr[data-gr-row]'), function(tr){
      var k = show[+tr.getAttribute('data-i')], pref = 'grR' + tr.getAttribute('data-i');
      var ed = canEdit(user(), k.c.loc), lastCell = tr.querySelector('.grLast'), newCell = tr.querySelector('.grNew'), saved = null;
      tr.querySelector('[data-open]').onclick = function(){ if (typeof root.openChild === 'function') root.openChild(k.id); };
      tr._refresh = function(){
        var l = last(CACHE[k.id] || []);
        lastCell.innerHTML = l ? fmtDate(l.date) + ' · ' + fmtW(l.weight) + ' кг · ' + fmtH(l.height) + ' см' : '<span style="color:#c0392b">немає</span>';
        if (saved) newCell.innerHTML = '<span style="color:#1e8449;font-weight:600">✅ ' + fmtW(saved.weight) + ' кг · ' + fmtH(saved.height) + ' см</span> <button class="btn btn-primary btn-xs grSend">💬 Надіслати батькам</button> <span style="font-size:11.5px;color:#1e8449">' + sentLine(saved) + '</span>';
        var sb = newCell.querySelector('.grSend'); if (sb) sb.onclick = function(){ API.sendOpen(k.id, saved.id); };
      };
      if (ed){
        newCell.innerHTML = '<span style="display:inline-flex;gap:6px;align-items:center">' + inputs(pref) + '<button class="btn btn-primary btn-sm grSave">💾</button><span class="grMsg" style="font-size:12px;color:#c0392b"></span></span>';
        var go = function(){
          var b = newCell.querySelector('.grSave'), msg = newCell.querySelector('.grMsg'); b.disabled = true; msg.textContent = '⏳';
          save(k.id, k.c, readInputs(pref)).then(function(it){
            saved = it; tr._refresh(); if (!it.id) API.load(k.id, true).then(function(l){ saved = last(l); tr._refresh(); });
            var nx = tr.nextElementSibling, ni = nx && nx.querySelector('input[id$="W"]'); if (ni) ni.focus();
          }, function(e){ b.disabled = false; msg.textContent = '⚠️ ' + (e && e.message || e); });
        };
        newCell.querySelector('.grSave').onclick = go;
        Array.prototype.forEach.call(newCell.querySelectorAll('input'), function(inp){ inp.onkeydown = function(e){ if (e.key === 'Enter'){ e.preventDefault(); go(); } }; });
      } else newCell.innerHTML = '<span style="font-size:12px;color:var(--gray)">перегляд</span>';
      tr._refresh();
    });
  }
  // пункт меню — медсестра, директор, CFO
  API.revealNav = function(){ var el = document.getElementById('navGrowth'); if (el) el.style.display = canList(user()) ? '' : 'none'; };
})(typeof window !== 'undefined' ? window : this);
