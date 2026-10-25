/* sw.js — Permite abrir la página aunque falle Internet.
   Siempre intenta descargar la versión más nueva; si no hay conexión,
   usa la última copia guardada. Nunca guarda las ventas (eso va a
   Google Sheets o a la lista de pendientes). */
var CACHE = 'sistema-ventas-v1';
var ARCHIVOS = ['./', 'index.html', 'styles.css', 'logica.js', 'script.js'];

self.addEventListener('install', function (e) {
  e.waitUntil(caches.open(CACHE).then(function (c) { return c.addAll(ARCHIVOS); }));
  self.skipWaiting();
});

self.addEventListener('activate', function (e) {
  e.waitUntil(caches.keys().then(function (claves) {
    return Promise.all(claves.filter(function (k) { return k !== CACHE; }).map(function (k) { return caches.delete(k); }));
  }));
  self.clients.claim();
});

self.addEventListener('fetch', function (e) {
  var req = e.request;
  if (req.method !== 'GET' || new URL(req.url).origin !== self.location.origin) return;
  e.respondWith(
    fetch(req).then(function (resp) {
      var copia = resp.clone();
      if (resp.ok) caches.open(CACHE).then(function (c) { c.put(req, copia); });
      return resp;
    }).catch(function () {
      return caches.match(req, { ignoreSearch: true });
    })
  );
});
