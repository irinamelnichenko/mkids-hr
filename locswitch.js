// v7.516: КІЛЬКА ЛОКАЦІЙ КОРИСТУВАЧА (директор / медсестра «Кар'єрна, Школа Кар'єрна») — активна локація + перемикач.
//   «Користувачі» → «Локація» через кому. У сесії: user.locs — усі, user.loc — АКТИВНА (з нею працюють усі сторінки,
//   як з однією локацією). Вибір пам'ятається в mkids_active_loc (за id користувача) і переживає перезахід.
//   Підключати в <head> ДО скриптів сторінки: сесію нормалізуємо до того, як сторінка її прочитає.
(function(){
  var SK = 'mkids_session', AK = 'mkids_active_loc';
  var NET = ['hr','cfo','ceo','cco','rnd_director','hr_trainer','legal','cmo','founder','coo'];
  function rd(k){ try { return JSON.parse(localStorage.getItem(k) || 'null'); } catch(e){ return null; } }
  function wr(k, v){ try { localStorage.setItem(k, JSON.stringify(v)); } catch(e){} }
  function split(s){ return String(s || '').split(/\s*[,;]\s*/).map(function(x){ return x.trim(); }).filter(Boolean); }
  function locsOf(u){
    if (!u || NET.indexOf(String(u.role || '').toLowerCase()) >= 0) return [];
    var a = Array.isArray(u.locs) && u.locs.length ? u.locs.slice() : split(u.loc);
    return a.length > 1 ? a : [];
  }
  function activeFor(u, locs){
    var a = rd(AK), want = a && String(a.id) === String(u.id) ? a.loc : '';
    if (locs.indexOf(want) >= 0) return want;
    return locs.indexOf(u.loc) >= 0 ? u.loc : locs[0];
  }
  // Нормалізує об'єкт користувача (локації через кому → locs, loc = активна). Повертає true, якщо щось змінив.
  function norm(u){
    var locs = locsOf(u); if (!locs.length) return false;
    var act = activeFor(u, locs), ch = !Array.isArray(u.locs) || u.locs.join('|') !== locs.join('|') || u.loc !== act;
    u.locs = locs; u.loc = act; return ch;
  }
  var s = rd(SK);
  if (s && s.user && norm(s.user)) wr(SK, s);

  window.mkidsNormUserLocs = norm;
  window.mkidsUserLocs = function(){ var x = rd(SK); return x && x.user ? locsOf(x.user) : []; };
  window.mkidsSetActiveLoc = function(loc){
    var x = rd(SK); if (!x || !x.user) return;
    var locs = locsOf(x.user); if (locs.indexOf(loc) < 0) return;
    wr(AK, {id:x.user.id, loc:loc}); x.user.locs = locs; x.user.loc = loc; wr(SK, x);
    location.reload();
  };

  function ui(){
    var x = rd(SK), locs = x && x.user ? locsOf(x.user) : [];
    var host = document.getElementById('uLoc');
    var old = document.getElementById('mkLocSw');
    if (!locs.length || !host){ if (old) old.remove(); return; }
    if (old) return;
    var sel = document.createElement('select');
    sel.id = 'mkLocSw'; sel.title = 'Ваші локації — перемкнути';
    sel.style.cssText = 'max-width:100%;margin-top:2px;padding:2px 4px;border:1px solid rgba(127,127,127,.45);border-radius:6px;font:inherit;font-size:12px;background:transparent;color:inherit;cursor:pointer';
    locs.forEach(function(l){ var o = document.createElement('option'); o.value = l; o.textContent = '📍 ' + l; if (l === x.user.loc) o.selected = true; sel.appendChild(o); });
    ['click','mousedown','touchstart'].forEach(function(ev){ sel.addEventListener(ev, function(e){ e.stopPropagation(); }); });
    sel.addEventListener('change', function(){ window.mkidsSetActiveLoc(sel.value); });
    host.style.display = 'none';
    host.parentNode.insertBefore(sel, host.nextSibling);
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', ui); else ui();
  // сторінка може показати шапку пізніше (після входу) — перевіряємо ще кілька разів
  var n = 0, t = setInterval(function(){ ui(); if (++n > 20) clearInterval(t); }, 1500);
  window.mkidsLocSwitchUi = ui;
})();
