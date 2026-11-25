/* v7.589 Польща (лише CFO): виписка PKO BP → перегляд → «Внести» доданками; «Готівка Федора»; звірка BALANCE з банком.
   Підключається на reconcile.html; українська частина сторінки не змінюється. Бекенд: plPreview / plApply / plBalance / plRowsList / plCash. */
(function(){
  'use strict';
  var PL = {ops: [], periods: {}, items: [], options: {}, bank: {}, overrides: {}, remember: {}, showAll: false, busy: false};
  var LOCS = ['LUBLIN', 'PRUZSKOW', 'KONSTANCIN', 'PRIMROSE'];
  var KIND = {payment: 'Payment', salary: 'Salary', opex: 'OPEX'};
  var MUA = {jan: 'січень', feb: 'лютий', mar: 'березень', apr: 'квітень', may: 'травень', jun: 'червень', jul: 'липень', aug: 'серпень', sep: 'вересень', oct: 'жовтень', nov: 'листопад', dec: 'грудень'};
  var ST = {ok: ['✅', 'буде внесено'], decide: ['❓', 'обери рядок'], dup: ['↩️', 'вже внесено'], skip: ['⏭', 'не вносимо'], closed: ['🔒', 'місяць закритий']};
  function $(id){ return document.getElementById(id); }
  function esc(s){ return String(s == null ? '' : s).replace(/[&<>"']/g, function(c){ return {'&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'}[c]; }); }
  function money(n){ return (Number(n) || 0).toLocaleString('uk-UA', {minimumFractionDigits: 2, maximumFractionDigits: 2}); }
  function toast(m, t, ms){ if (typeof window.showToast === 'function') window.showToast(m, t || 'info', ms); }
  function actor(){ return (window.S && S.user && S.user.id) || ''; }
  function isCfo(){ return !!(window.S && S.user && String(S.user.role || '').toLowerCase() === 'cfo'); }
  async function post(body){
    body.actorId = actor();
    var r = await fetch(CFG.SCRIPT_URL, {method: 'POST', redirect: 'follow', headers: {'Content-Type': 'text/plain;charset=utf-8'}, body: JSON.stringify(body)});
    var t = await r.text(); try { return JSON.parse(t); } catch(e){ return {ok: false, error: 'Відповідь не JSON: ' + t.slice(0, 120)}; }
  }

  // ── розмітка панелі ──
  function ensureBox(){
    if ($('plBox')) return;
    var box = document.createElement('div'); box.id = 'plBox'; box.className = 'hidden';
    box.innerHTML =
      '<div class="inv-card" style="margin-bottom:14px">' +
      '  <div style="display:flex;gap:10px;flex-wrap:wrap;align-items:center">' +
      '    <input type="file" id="plFile" accept=".xls,.xlsx" multiple style="display:none">' +
      '    <button class="btn btn-primary" id="plPick">📂 Виписки PKO BP (.xls)</button>' +
      '    <button class="btn" id="plBal">⚖️ Звірка BALANCE з банком</button>' +
      '    <button class="btn" id="plCashBtn">💵 Готівка Федора</button>' +
      '    <label style="font-size:13px;color:var(--gray)"><input type="checkbox" id="plAll"> показати вже внесене й пропущене</label>' +
      '  </div>' +
      '  <div id="plMeta" style="margin-top:8px;font-size:13px;color:var(--gray)">Можна обрати кілька файлів — по одному на рахунок. Рахунок → локація визначається автоматично. Повторне завантаження тієї ж виписки не дублює.</div>' +
      '</div>' +
      '<div id="plCash" class="inv-card hidden" style="margin-bottom:14px"></div>' +
      '<div id="plSummary" class="fc-chips"></div>' +
      '<div class="inv-card"><div id="plWrap"><div class="inv-empty">📂 Обери виписки PKO BP — система розкладе кожну операцію у файл, вкладку й рядок.</div></div></div>' +
      '<div id="plApplyBar" style="margin-top:14px;display:none"><button class="btn btn-primary" id="plApplyBtn">✅ Внести</button>' +
      '  <span id="plApplyHint" style="margin-left:10px;color:var(--gray);font-size:12px"></span></div>' +
      '<div id="plResult" style="margin-top:14px"></div>';
    var c = $('content'); c.parentNode.insertBefore(box, c.nextSibling);
    $('plPick').onclick = function(){ $('plFile').click(); };
    $('plFile').onchange = function(){ onFiles(this.files); this.value = ''; };
    $('plBal').onclick = function(){ showBalance(); };
    $('plCashBtn').onclick = function(){ cashForm(); };
    $('plAll').onchange = function(){ PL.showAll = this.checked; render(); };
    $('plApplyBtn').onclick = function(){ apply(); };
  }
  window.plMode = function(on){
    if (!isCfo()) return;
    ensureBox();
    var pl = on == null ? $('plBox').classList.contains('hidden') : !!on;
    $('plBox').classList.toggle('hidden', !pl);
    ['content', 'rcBankBox'].forEach(function(id){ var e = $(id); if (e) e.style.display = pl ? 'none' : ''; });
    var pick = document.querySelector('.hd-ctrl > .btn.btn-primary'); if (pick) pick.style.display = pl ? 'none' : '';
    if ($('meta')) $('meta').textContent = pl ? 'Польща: виписка PKO BP → фінфайли 2026 (лише CFO)' : 'завантаж виписку ПриватБанку (.xls / .xlsx / .csv)';
    if ($('plModeBtn')) $('plModeBtn').textContent = pl ? '🇺🇦 Україна' : '🇵🇱 Польща';
    if (!pl && typeof window.rcBankInit === 'function') window.rcBankInit();
  };

  // ── читання виписки PKO (.xls): «Lista transakcji» + «Kryteria wyszukiwania» ──
  function isoDate(v){
    if (typeof v === 'number' && typeof XLSX !== 'undefined' && XLSX.SSF){ var d = XLSX.SSF.parse_date_code(v); if (d) return d.y + '-' + ('0' + d.m).slice(-2) + '-' + ('0' + d.d).slice(-2); }
    var s = String(v || '').trim(), m = /^(\d{4})-(\d{2})-(\d{2})/.exec(s); if (m) return m[0];
    m = /^(\d{2})\.(\d{2})\.(\d{4})/.exec(s); return m ? m[3] + '-' + m[2] + '-' + m[1] : '';
  }
  function num(v){ if (typeof v === 'number') return v; var n = Number(String(v || '').replace(/[\s ]/g, '').replace(',', '.')); return isNaN(n) ? 0 : n; }
  function parseBook(buf){
    var wb = XLSX.read(new Uint8Array(buf), {type: 'array'});
    var crit = wb.Sheets['Kryteria wyszukiwania'] || wb.Sheets[wb.SheetNames[1]], tr = wb.Sheets['Lista transakcji'] || wb.Sheets[wb.SheetNames[0]];
    var acc = '', from = '', to = '';
    (crit ? XLSX.utils.sheet_to_json(crit, {header: 1, raw: true, defval: ''}) : []).forEach(function(r){
      var k = String(r[0] || ''); if (/^Numer rachunku/i.test(k)) acc = String(r[1] || '').replace(/\s/g, ''); if (/^Od dnia/i.test(k)) from = isoDate(r[1]); if (/^Do dnia/i.test(k)) to = isoDate(r[1]); });
    var rows = XLSX.utils.sheet_to_json(tr, {header: 1, raw: true, defval: ''}), ops = [];
    rows.slice(1).forEach(function(v){ var d = isoDate(v[0]); if (!d) return;
      ops.push({acc: acc, date: d, type: String(v[2] || ''), amount: num(v[3]), saldo: num(v[5]), from_acc: String(v[6] || ''), from: String(v[7] || '').trim(),
        to_acc: String(v[9] || ''), to: String(v[10] || '').trim(), desc: v.slice(12).map(function(x){ return String(x).trim(); }).filter(Boolean).join(' | ')}); });
    return {acc: acc, from: from, to: to, ops: ops};
  }
  function onFiles(files){
    if (!files || !files.length) return;
    if (typeof XLSX === 'undefined'){ toast('❌ Бібліотека XLSX не завантажена', 'error'); return; }
    var list = Array.prototype.slice.call(files), done = 0, books = [];
    list.forEach(function(f){ var rd = new FileReader();
      rd.onload = function(ev){ try { books.push(parseBook(ev.target.result)); } catch(e){ toast('❌ ' + f.name + ': ' + e.message, 'error'); }
        if (++done === list.length) loaded(books); };
      rd.readAsArrayBuffer(f); });
  }
  function loaded(books){
    PL.ops = []; PL.periods = {}; PL.overrides = {}; PL.remember = {};
    var seen = {};
    books.forEach(function(b){ if (!b.acc) return; PL.periods[b.acc.slice(-4)] = {from: b.from, to: b.to};
      b.ops.forEach(function(o){ var k = [o.acc, o.date, o.amount, o.saldo, o.desc].join('|'); if (seen[k]) return; seen[k] = 1; PL.ops.push(o); }); });
    $('plMeta').textContent = books.length + ' файл(и), ' + PL.ops.length + ' операцій · ' + books.map(function(b){ return '…' + b.acc.slice(-4) + ' ' + b.from + '…' + b.to; }).join(', ');
    preview();
  }
  async function preview(){
    if (!PL.ops.length) return;
    $('plWrap').innerHTML = '<div class="inv-loading">⏳ Розкладаю операції по файлах…</div>';
    var j = await post({action: 'plPreview', ops: PL.ops, overrides: PL.overrides});
    if (!j.ok){ $('plWrap').innerHTML = '<div class="err">❌ ' + esc(j.error || 'помилка') + '</div>'; return; }
    PL.items = j.items || []; PL.options = j.options || {}; PL.bank = j.bank || {};
    render();
  }

  // ── перегляд ──
  function target(it){
    if (!it.cells.length && !it.cf.length) return '<span style="color:var(--gray)">—</span>';
    var a = it.cells.map(function(c){ return '<b>' + esc(c.loc) + '</b> · ' + esc(c.tab) + ' · ' + esc(c.a1) + ' <span style="color:var(--gray)">' + esc(c.name || '') + '</span>' + (it.cells.length > 1 ? ' = ' + money(c.val) : ''); });
    it.cf.forEach(function(c){ a.push('CF · ' + esc(c.a1) + ' <span style="color:var(--gray)">' + esc(c.name || '') + '</span>'); });
    return a.join('<br>');
  }
  function picker(it){
    if (['ok', 'decide'].indexOf(it.status) < 0 || it.cells.length > 1 || it.journal || it.type === 'Opłata') return '';
    var o = PL.options[it.acc + '|' + it.mon]; if (!o) return '';
    var cur = it.cells[0] ? it.cells[0].kind + '|' + it.cells[0].row : '';
    var h = '<select data-k="' + esc(it.key) + '" class="plPick" style="max-width:260px;padding:4px;border:1px solid #cbd5e1;border-radius:6px"><option value="">' + (it.status === 'decide' ? '— обери рядок —' : 'інший рядок…') + '</option>';
    Object.keys(KIND).forEach(function(k){ h += '<optgroup label="' + KIND[k] + '">' + (o[k] || []).map(function(x){ var v = k + '|' + x.r;
      return '<option value="' + v + '"' + (v === cur && it.status !== 'decide' ? ' selected' : '') + '>' + x.r + ' · ' + esc(String(x.a).slice(0, 50)) + '</option>'; }).join('') + '</optgroup>'; });
    return h + '</select>';
  }
  function render(){
    var c = {}; PL.items.forEach(function(it){ c[it.status] = (c[it.status] || 0) + 1; });
    $('plSummary').innerHTML = Object.keys(ST).filter(function(k){ return c[k]; }).map(function(k){ return '<span class="fc-chip">' + ST[k][0] + ' ' + ST[k][1] + ': <b>' + c[k] + '</b></span>'; }).join('');
    var h = '';
    LOCS.concat(['OVERHEAD', '?']).forEach(function(l){
      var arr = PL.items.filter(function(it){ return it.acc === l && (PL.showAll || ['ok', 'decide'].indexOf(it.status) >= 0); });
      if (!arr.length) return;
      arr.sort(function(a, b){ return (a.status === 'decide' ? 0 : 1) - (b.status === 'decide' ? 0 : 1) || String(a.date).localeCompare(String(b.date)); });
      h += '<h3 style="margin:16px 0 8px">' + esc(l) + ' <span style="font-weight:400;color:var(--gray);font-size:13px">' + arr.length + ' операцій</span></h3>' +
        '<div style="overflow-x:auto"><table class="inv-table" style="width:100%;border-collapse:collapse;font-size:13px"><thead><tr>' +
        '<th>Дата</th><th>Контрагент</th><th style="text-align:right">Сума</th><th>Куди: файл · вкладка · клітинка · рядок</th><th>Місяць</th><th>Підстава</th><th>Змінити рядок</th></tr></thead><tbody>' +
        arr.map(function(it){ var s = ST[it.status] || ['', it.status];
          return '<tr style="border-top:1px solid #e5e7eb' + (it.status === 'decide' ? ';background:#fff7ed' : (it.status === 'ok' ? '' : ';color:var(--gray)')) + '">' +
            '<td style="white-space:nowrap">' + s[0] + ' ' + esc(it.date) + '</td>' +
            '<td>' + esc(String(it.party).slice(0, 48)) + '<div style="color:var(--gray);font-size:11px">' + esc(String(it.title || '').slice(0, 70)) + '</div></td>' +
            '<td style="text-align:right;white-space:nowrap;' + (it.amount < 0 ? 'color:#b91c1c' : 'color:#047857') + '">' + money(it.amount) + '</td>' +
            '<td>' + target(it) + '</td><td>' + esc(MUA[it.mon] || it.mon) + '</td>' +
            '<td style="font-size:12px">' + esc(it.why) + (it.journal ? ' <b>· журнал ' + esc(it.journal) + '</b>' : '') + '</td>' +
            '<td>' + picker(it) + '</td></tr>'; }).join('') + '</tbody></table></div>';
    });
    $('plWrap').innerHTML = h || '<div class="inv-empty">Немає операцій для внесення' + (PL.items.length ? ' — усе вже внесено або не вноситься (познач «показати вже внесене»)' : '') + '.</div>';
    Array.prototype.forEach.call(document.querySelectorAll('.plPick'), function(s){ s.onchange = function(){ pick(this.getAttribute('data-k'), this.value); }; });
    var ok = c.ok || 0, dec = c.decide || 0;
    $('plApplyBar').style.display = ok ? '' : 'none';
    $('plApplyHint').textContent = ok + ' операцій буде внесено доданками з примітками' + (dec ? ' · ' + dec + ' без рядка — не вносяться, поки не обереш' : '') + '.';
  }
  function pick(key, v){
    var it = PL.items.filter(function(x){ return x.key === key; })[0]; if (!it) return;
    if (!v){ delete PL.overrides[key]; delete PL.remember[it.acc + '|' + it.partyKey]; preview(); return; }
    var q = v.split('|'), o = PL.options[it.acc + '|' + it.mon] || {}, nm = ((o[q[0]] || []).filter(function(x){ return String(x.r) === q[1]; })[0] || {}).a || '';
    PL.overrides[key] = {kind: q[0], row: Number(q[1])};
    if (it.partyKey) PL.remember[it.acc + '|' + it.partyKey] = {loc: it.acc, party: it.partyKey, kind: q[0], row: Number(q[1]), name: nm};
    preview();
  }

  // ── «Внести» ──
  async function apply(){
    if (PL.busy) return; PL.busy = true;
    var accs = LOCS.filter(function(l){ return PL.items.some(function(it){ return it.acc === l && it.status === 'ok'; }); });
    var out = [], rem = Object.keys(PL.remember).map(function(k){ return PL.remember[k]; });
    $('plApplyBtn').disabled = true;
    try {
      for (var i = 0; i < accs.length; i++){
        $('plApplyHint').textContent = '⏳ Вношу ' + accs[i] + ' (' + (i + 1) + ' з ' + accs.length + ')…';
        var j = await post({action: 'plApply', ops: PL.ops, overrides: PL.overrides, remember: i === 0 ? rem : [], only: accs[i], periods: PL.periods});
        out.push({loc: accs[i], r: j});
      }
      PL.remember = {};
      var hdr = out.map(function(x){ var r = x.r;
        if (!r.ok) return '<div class="err">❌ ' + esc(x.loc) + ': ' + esc(r.error || 'помилка') + '</div>';
        var bad = (r.unsafe || []).concat(r.failed || []);
        return '<div style="margin:4px 0">' + (bad.length || (r.notDone || []).length ? '⚠️' : '✅') + ' <b>' + esc(x.loc) + '</b>: внесено операцій ' + r.done + ', клітинок змінено ' + r.cells +
          (r.dup ? ', уже були ' + r.dup : '') + (bad.length ? ' · <b>не записано клітинок: ' + bad.length + '</b> — ' + esc(bad.map(function(b){ return b.k.split('|').slice(1).join(' ') + ' (' + (b.why || b.err) + ')'; }).join('; ')) : '') + '</div>'; }).join('');
      $('plResult').innerHTML = '<div class="inv-card">' + hdr + '</div>';
      toast('Внесено', 'info', 4000);
      var mons = {}; PL.items.forEach(function(it){ if (it.status === 'ok') mons[it.mon] = 1; });
      await showBalance(Object.keys(mons));
      await preview();
    } catch(e){ toast('❌ ' + (e && e.message || e), 'error'); }
    finally { PL.busy = false; $('plApplyBtn').disabled = false; }
  }

  // ── звірка BALANCE з банком ──
  async function showBalance(months){
    ensureBox();
    var box = document.createElement('div'); box.className = 'inv-card'; box.style.marginTop = '14px'; box.innerHTML = '<div class="inv-loading">⏳ Звіряю BALANCE з банком…</div>';
    var old = $('plBalBox'); if (old) old.remove(); box.id = 'plBalBox'; $('plResult').appendChild(box);
    var j = await post({action: 'plBalance', months: months || []});
    if (!j.ok){ box.innerHTML = '<div class="err">❌ ' + esc(j.error || 'помилка') + '</div>'; return; }
    box.innerHTML = '<h3 style="margin:0 0 8px">⚖️ BALANCE у файлі проти банку</h3><div style="overflow-x:auto"><table style="width:100%;font-size:13px;border-collapse:collapse"><thead><tr>' +
      '<th>Локація</th><th>Місяць</th><th style="text-align:right">Opening</th><th style="text-align:right">BALANCE у файлі</th><th style="text-align:right">Банк на кінець (за датами)</th><th style="text-align:right">Різниця</th><th style="text-align:right">Банк за проведенням</th></tr></thead><tbody>' +
      (j.rows || []).map(function(r){ var okk = r.diff != null && Math.abs(r.diff) < 0.01;
        return '<tr style="border-top:1px solid #e5e7eb"><td><b>' + esc(r.loc) + '</b></td><td>' + esc(MUA[r.mon] || r.mon) + '</td><td style="text-align:right">' + money(r.opening) + '</td><td style="text-align:right">' + money(r.balance) + '</td>' +
          '<td style="text-align:right">' + (r.bankByDate == null ? '<span style="color:var(--gray)">виписки немає</span>' : money(r.bankByDate)) + '</td>' +
          '<td style="text-align:right;font-weight:700;color:' + (r.diff == null ? 'inherit' : okk ? '#047857' : '#b91c1c') + '">' + (r.diff == null ? '—' : (okk ? '✓ 0' : money(r.diff))) + '</td>' +
          '<td style="text-align:right;color:var(--gray)">' + (r.bankClose == null ? '' : money(r.bankClose) + (r.gap ? ' (розрив ' + money(r.gap) + ')' : '')) + '</td></tr>'; }).join('') +
      '</tbody></table></div><div style="font-size:12px;color:var(--gray);margin-top:6px">«Розрив» — операції, які банк провів у цьому місяці, але датовані іншим (у виписку за датами не потрапили).</div>';
  }

  // ── «Готівка Федора» ──
  function cashForm(){
    var b = $('plCash'); b.classList.toggle('hidden'); if (b.classList.contains('hidden')) return;
    var today = new Date().toISOString().slice(0, 10);
    b.innerHTML = '<h3 style="margin:0 0 10px">💵 Готівка Федора</h3><div style="font-size:12px;color:var(--gray);margin-bottom:8px">Витрата — доданком у свій рядок Salary / OPEX, та сама сума — у CF «наличка Федор» як надходження (залишок банку не зменшує).</div>' +
      '<div style="display:flex;gap:8px;flex-wrap:wrap;align-items:center">' +
      '<input type="date" id="plcD" value="' + today + '" style="padding:6px;border:1px solid #cbd5e1;border-radius:8px">' +
      '<input type="number" step="0.01" min="0" id="plcA" placeholder="Сума, zł" style="width:120px;padding:6px;border:1px solid #cbd5e1;border-radius:8px">' +
      '<select id="plcL" style="padding:6px;border:1px solid #cbd5e1;border-radius:8px">' + LOCS.map(function(l){ return '<option>' + l + '</option>'; }).join('') + '</select>' +
      '<select id="plcK" style="padding:6px;border:1px solid #cbd5e1;border-radius:8px"><option value="salary">Salary</option><option value="opex">OPEX</option></select>' +
      '<select id="plcR" style="max-width:280px;padding:6px;border:1px solid #cbd5e1;border-radius:8px"><option value="">завантажую рядки…</option></select>' +
      '<input id="plcN" placeholder="Примітка (за що)" style="min-width:200px;padding:6px;border:1px solid #cbd5e1;border-radius:8px">' +
      '<button class="btn btn-primary" id="plcGo">Внести готівку</button></div><div id="plcMsg" style="margin-top:8px;font-size:13px"></div>';
    var rows = {};
    async function load(){ var l = $('plcL').value, d = $('plcD').value, mi = Number(d.slice(5, 7)) - 1, m = ['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec'][mi];
      $('plcR').innerHTML = '<option value="">завантажую рядки…</option>';
      var j = await post({action: 'plRowsList', loc: l, month: m}); rows = (j && j.options) || {}; fill(); }
    function fill(){ var k = $('plcK').value; $('plcR').innerHTML = '<option value="">— рядок —</option>' + (rows[k] || []).map(function(x){ return '<option value="' + x.r + '">' + x.r + ' · ' + esc(String(x.a).slice(0, 50)) + '</option>'; }).join(''); }
    $('plcL').onchange = load; $('plcD').onchange = load; $('plcK').onchange = fill;
    $('plcGo').onclick = async function(){
      var body = {action: 'plCash', id: Date.now() + '-' + Math.random().toString(36).slice(2, 8), date: $('plcD').value, amount: Number($('plcA').value), loc: $('plcL').value, kind: $('plcK').value, row: Number($('plcR').value), note: $('plcN').value};
      if (!(body.amount > 0) || !body.row){ $('plcMsg').textContent = 'Вкажи суму й рядок'; return; }
      this.disabled = true; $('plcMsg').textContent = '⏳ Вношу…';
      var j = await post(body); this.disabled = false;
      $('plcMsg').innerHTML = j.ok ? '✅ Внесено: ' + esc(body.loc) + ' · ' + esc(j.name || '') + ' · ' + money(body.amount) + ' zł, і стільки ж у CF «наличка Федор»' : '❌ ' + esc(j.error || 'помилка');
      if (j.ok){ $('plcA').value = ''; $('plcN').value = ''; }
    };
    load();
  }
})();
