// m.kids PWA Service Worker (cache v7.401, libs v1)
// Cache: static shell (HTML + manifest + icons + xlsx)
// Strategy:
//   • POST → завжди network (ніколи не кешуємо)
//   • GET до /macros/ (Apps Script) → network-only, fallback на cache
//   • GET до інших static URLs → cache-first з фоновим оновленням

var CACHE = 'mkids-cache-v7.403';
// v7.401: бібліотеки (xlsx 952 КБ, pdf.js + worker 1,3 МБ, chart.js, Leaflet) — в ОКРЕМОМУ кеші, який НЕ
// стирається з кожним релізом і не перекачується щоразу у фоні. Нова версія бібліотеки → змінити LIBS.
var LIBS = 'mkids-libs-v1';
var LIB_RE = /(xlsx\.full\.min\.js|pdf\.min\.js|pdf\.worker\.min\.js|chart\.umd\.min\.js|leaflet|markercluster)/i;
var LIB_FILES = ['xlsx.full.min.js', 'pdf.min.js', 'pdf.worker.min.js'];
var SHELL = [
  './',
  'activities.html',
  'leads.html',
  'meals.html',
  'install.html',
  'invoice_report.html',
  'control.html',
  'reconcile.html',
  'salary_reconcile.html',
  'map.html',            // v7.371 карта клієнтів
  'history.html',        // v7.372 історія внесень
  'cash.html',           // v7.373 готівка
  'manifest.json',
  'icon-192.png',
  'icon-512.png'
];

self.addEventListener('install', function(ev){
  self.skipWaiting();
  ev.waitUntil(Promise.all([
    caches.open(CACHE).then(function(c){
      // Кешуємо по одному — щоб одна 404 не зривала весь install
      return Promise.all(SHELL.map(function(url){
        return c.add(url).catch(function(e){
          console.warn('[sw] skip cache:', url, e && e.message);
        });
      }));
    }),
    caches.open(LIBS).then(function(c){   // v7.401: бібліотеки — лише якщо їх ще немає
      return Promise.all(LIB_FILES.map(function(url){
        return c.match(url).then(function(hit){ return hit || c.add(url).catch(function(){}); });
      }));
    })
  ]));
});

self.addEventListener('activate', function(ev){
  ev.waitUntil(
    caches.keys().then(function(keys){
      return Promise.all(keys.map(function(k){
        if (k !== CACHE && k !== LIBS) return caches.delete(k);   // v7.401: бібліотеки не чіпаємо
      }));
    }).then(function(){ return self.clients.claim(); })
  );
});

self.addEventListener('fetch', function(ev){
  var req = ev.request;
  if (req.method !== 'GET') return;        // POST/PUT/DELETE — pass-through

  var url = new URL(req.url);
  if (url.hostname === 'api.geoapify.com') return;   // v7.368: підказки адрес — завжди мережа, без кешу
  if (/(^|\.)tile\.openstreetmap\.org$|basemaps\.cartocdn\.com$/.test(url.hostname)) return;   // v7.401: тайли карти — повз SW

  // v7.401: бібліотеки — cache-first з окремого кешу, без фонового перекачування
  if (LIB_RE.test(url.pathname)){
    ev.respondWith(caches.open(LIBS).then(function(c){
      return c.match(req).then(function(hit){
        return hit || fetch(req).then(function(resp){
          if (resp && (resp.status === 200 || resp.type === 'opaque')) c.put(req, resp.clone());
          return resp;
        });
      });
    }));
    return;
  }
  var isApi = url.hostname === 'script.google.com' ||
              url.hostname === 'script.googleusercontent.com';

  if (isApi){
    // Network-first для API; fallback на кеш якщо офлайн (краще ніж error).
    ev.respondWith(
      fetch(req).then(function(resp){
        // Не кешуємо API (дані змінюються)
        return resp;
      }).catch(function(){
        return caches.match(req).then(function(c){
          return c || new Response(JSON.stringify({ok:false, error:'offline'}),
            {status:503, headers:{'Content-Type':'application/json'}});
        });
      })
    );
    return;
  }

  // v6.16: HTML — NETWORK-FIRST щоб оновлення підхоплювались одразу
  // (cache-first для HTML призводив до того що користувач бачив стару версію
  // навіть після push на GitHub Pages; SW оновлював у фоні і тільки на
  // наступне відкриття показував свіже)
  var isHtml = url.pathname.endsWith('.html') ||
               url.pathname === '/' ||
               url.pathname === '/mkids-hr/' ||
               url.pathname.endsWith('/');
  if (isHtml){
    ev.respondWith(
      fetch(req).then(function(resp){
        if (resp && resp.status === 200 && resp.type === 'basic'){
          var clone = resp.clone();
          caches.open(CACHE).then(function(c){ c.put(req, clone); });
        }
        return resp;
      }).catch(function(){
        return caches.match(req).then(function(c){
          return c || new Response('Offline', {status: 503});
        });
      })
    );
    return;
  }

  // Static (images, manifest, JS-libs) — cache-first з оновленням у фоні
  ev.respondWith(
    caches.match(req).then(function(cached){
      var fresh = fetch(req).then(function(resp){
        if (resp && resp.status === 200 && resp.type === 'basic'){
          var clone = resp.clone();
          caches.open(CACHE).then(function(c){ c.put(req, clone); });
        }
        return resp;
      }).catch(function(){ return cached; });
      return cached || fresh;
    })
  );
});
