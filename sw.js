/* ════════════════════════════════════════════════════════════════════
   Service Worker — Padrón Prestadores
   Objetivo: que al REFRESCAR no se re-descargue el archivo (2.5 MB) ni las
   librerías. Los DATOS del repo (.csv.gz) siempre van a la red, así un
   período nuevo se agarra solo sin tocar nada.

   ▶ Subir este archivo UNA vez a la raíz del repo, junto a index.html.
   ════════════════════════════════════════════════════════════════════ */

const CACHE = 'padron-v1';

/* Librerías estáticas que conviene cachear (no cambian) */
const PRECACHE = [
  './',
  './index.html',
  'https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js',
  'https://cdn.jsdelivr.net/npm/d3@7.9.0/dist/d3.min.js',
  'https://cdn.jsdelivr.net/npm/pako@2.1.0/dist/pako.min.js',
  'https://cdn.jsdelivr.net/npm/papaparse@5.4.1/papaparse.min.js',
  'https://cdn.jsdelivr.net/npm/xlsx-js-style@1.2.0/dist/xlsx.bundle.js',
  'https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.0/chart.umd.min.js',
  'https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js'
];

self.addEventListener('install', function(e){
  self.skipWaiting();
  e.waitUntil(
    caches.open(CACHE).then(function(c){
      // cachear de a uno, ignorando los que fallen (CDN caído, etc.)
      return Promise.all(PRECACHE.map(function(url){
        return c.add(url).catch(function(){ /* ignorar fallo individual */ });
      }));
    })
  );
});

self.addEventListener('activate', function(e){
  e.waitUntil(
    caches.keys().then(function(keys){
      return Promise.all(keys.map(function(k){
        if(k !== CACHE) return caches.delete(k);
      }));
    }).then(function(){ return self.clients.claim(); })
  );
});

self.addEventListener('fetch', function(e){
  const req = e.request;
  if(req.method !== 'GET'){ return; }

  let url;
  try { url = new URL(req.url); } catch(_) { return; }

  /* 1) DATOS del repo (.csv.gz de raw.githubusercontent) → SIEMPRE a la red.
        Así un período nuevo se baja sin quedar pegado a una versión vieja. */
  if(url.hostname.indexOf('raw.githubusercontent.com') !== -1 ||
     /\.csv\.gz($|\?)/i.test(url.pathname)){
    return; // sin responder → el navegador hace su fetch normal a la red
  }

  /* 2) NAVEGACIÓN (abrir/refrescar la página) → stale-while-revalidate:
        sirve el HTML cacheado al instante y, en paralelo, busca uno nuevo
        para la próxima vez. Refrescar deja de re-descargar 2.5 MB. */
  const isHTML = req.mode === 'navigate' ||
                 (req.headers.get('accept')||'').indexOf('text/html') !== -1;
  if(isHTML){
    e.respondWith(
      caches.open(CACHE).then(function(c){
        return c.match('./index.html').then(function(cached){
          const net = fetch(req).then(function(res){
            if(res && res.status === 200){ c.put('./index.html', res.clone()); }
            return res;
          }).catch(function(){ return cached; });
          return cached || net;
        });
      })
    );
    return;
  }

  /* 3) LIBRERÍAS / otros estáticos → cache-first (rápido), con respaldo a red. */
  e.respondWith(
    caches.match(req).then(function(cached){
      if(cached) return cached;
      return fetch(req).then(function(res){
        if(res && res.status === 200 && (url.protocol === 'https:')){
          const copy = res.clone();
          caches.open(CACHE).then(function(c){ c.put(req, copy); });
        }
        return res;
      });
    })
  );
});
