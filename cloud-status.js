// ═══════════════════════════════════════════════════════════════════════════
// m.kids — спільний хмарний шар: cloudFetch (таймаут+ретрай) + індикатор статусу.
// Підключається у index.html та clients.html. Глобали: cloudFetch, setCloudStatus,
// runCloudLoaders. v7.94.
// ═══════════════════════════════════════════════════════════════════════════
(function(global){
  'use strict';

  // ── 1) cloudFetch(url, opts) ──────────────────────────────────────────────
  //   AbortController-таймаут 45с; 2 повтори (паузи 1с і 3с) ЛИШЕ на МЕРЕЖЕВИЙ збій
  //   (fetch throw / abort). HTTP-статус (навіть 4xx/5xx) НЕ повторюємо — повертаємо
  //   response як є (нехай викликач сам вирішує по r.ok). no-cors POST резолвиться
  //   (opaque) → повертається без повтору.
  //   opts: { timeoutMs?:45000, retries?:[1000,3000], fetch?:{...fetch-init} }
  var DEFAULT_RETRIES = [1000, 3000];
  function _sleep(ms){ return new Promise(function(res){ setTimeout(res, ms); }); }
  async function cloudFetch(url, opts){
    opts = opts || {};
    var timeoutMs = opts.timeoutMs || 45000;
    var delays = opts.retries || DEFAULT_RETRIES;   // 2 повтори за замовч.
    var fetchInit = Object.assign({ redirect: 'follow' }, opts.fetch || {});
    var attempts = delays.length + 1;
    var lastErr;
    for (var i = 0; i < attempts; i++){
      var ctrl = new AbortController();
      var timer = setTimeout(function(){ try { ctrl.abort(); } catch(_){} }, timeoutMs);
      try {
        var resp = await fetch(url, Object.assign({}, fetchInit, { signal: ctrl.signal }));
        clearTimeout(timer);
        return resp;   // HTTP-статус не ретраїмо
      } catch(e){
        clearTimeout(timer);
        lastErr = e;                       // мережевий збій / таймаут(abort)
        if (i < attempts - 1){ await _sleep(delays[i]); continue; }
        throw lastErr;                      // ретраї вичерпано
      }
    }
    throw lastErr;
  }

  // ── 2) setCloudStatus(state, elId) — єдиний індикатор (за замовч. #sheetsStatus) ──
  var LABELS = {
    sync:  { t: '☁️ Синхронізація…', c: '#e67e22' },
    ok:    { t: '☁️ Синхронізовано', c: '#27ae60' },
    err:   { t: '⚠️ Помилка хмари',  c: '#e74c3c' },
    local: { t: '💾 Лише локально',  c: '#7f8c8d' },
    none:  { t: '',                  c: '#7f8c8d' }
  };
  function setCloudStatus(state, elId){
    var el = document.getElementById(elId || 'sheetsStatus'); if (!el) return;
    var s = LABELS[state] || LABELS.none;
    el.textContent = s.t; el.style.color = s.c;
  }

  // ── 3) runCloudLoaders(loaders) — критичні/некритичні ─────────────────────
  //   loaders: [{name, crit:bool, run:()=>Promise}]. Виконує паралельно (allSettled).
  //   Повертає {critFail, results}: critFail=true лише коли впав КРИТИЧНИЙ лоадер
  //   (некритичні падіння логуються, але індикатор не червоніє). Дзеркало логіки,
  //   що вже є в clients.html (init/refreshAll).
  async function runCloudLoaders(loaders){
    loaders = loaders || [];
    var results = await Promise.allSettled(loaders.map(function(l){
      return Promise.resolve().then(function(){ return l.run(); });
    }));
    var critFail = false;
    results.forEach(function(r, i){
      if (r.status === 'rejected'){
        var l = loaders[i] || {};
        console.error('[cloud] лоадер "' + l.name + '" ВПАВ' + (l.crit ? ' (КРИТИЧНИЙ)' : ' (некритичний)') + ':', r.reason);
        if (l.crit) critFail = true;
      }
    });
    return { critFail: critFail, results: results };
  }

  global.cloudFetch = cloudFetch;
  global.setCloudStatus = setCloudStatus;
  global.runCloudLoaders = runCloudLoaders;
})(window);
