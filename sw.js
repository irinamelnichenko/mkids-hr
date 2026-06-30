// m.kids PWA Service Worker (v6.51-reconcile / cache v6.60)
// Cache: static shell (HTML + manifest + icons + xlsx)
// Strategy:
//   • POST → завжди network (ніколи не кешуємо)
//   • GET до /macros/ (Apps Script) → network-only, fallback на cache
//   • GET до інших static URLs → cache-first з фоновим оновленням

var CACHE = 'mkids-cache-v6.99';
var SHELL = [
  './',
  'activities.html',
  'install.html',
  'invoice_report.html',
  'control.html',
  'reconcile.html',
  'manifest.json',
  'icon-192.png',
  'icon-512.png',
  'xlsx.full.min.js'
];

self.addEventListener('install', function(ev){
  self.skipWaiting();
  ev.waitUntil(
    caches.open(CACHE).then(function(c){
      // Кешуємо по одному — щоб одна 404 не зривала весь install
      return Promise.all(SHELL.map(function(url){
        return c.add(url).catch(function(e){
          console.warn('[sw] skip cache:', url, e && e.message);
        });
      }));
    })
  );
});

self.addEventListener('activate', function(ev){
  ev.waitUntil(
    caches.keys().then(function(keys){
      return Promise.all(keys.map(function(k){
        if (k !== CACHE) return caches.delete(k);
      }));
    }).then(function(){ return self.clients.claim(); })
  );
});

self.addEventListener('fetch', function(ev){
  var req = ev.request;
  if (req.method !== 'GET') return;        // POST/PUT/DELETE — pass-through

  var url = new URL(req.url);
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
