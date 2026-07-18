const CACHE_NAME = 'aaa-auditoria-v2';
const APP_SHELL = ['./index.html', './manifest.json', './icon-192.png', './icon-512.png'];

self.addEventListener('install', function (event) {
  event.waitUntil(
    caches.open(CACHE_NAME).then(function (cache) {
      return cache.addAll(APP_SHELL);
    })
  );
  self.skipWaiting();
});

self.addEventListener('activate', function (event) {
  event.waitUntil(
    caches.keys().then(function (keys) {
      return Promise.all(
        keys.filter(function (k) { return k !== CACHE_NAME; }).map(function (k) { return caches.delete(k); })
      );
    })
  );
  self.clients.claim();
});

self.addEventListener('fetch', function (event) {
  const url = event.request.url;
  // Los datos en vivo (Google Sheets) siempre van a la red, nunca a la caché.
  if (url.indexOf('docs.google.com') !== -1 || url.indexOf('googleusercontent.com') !== -1) {
    event.respondWith(
      fetch(event.request).catch(function () { return caches.match(event.request); })
    );
    return;
  }
  // El resto del "app shell" usa estrategia cache-first para funcionar sin conexión.
  event.respondWith(
    caches.match(event.request).then(function (cached) {
      return cached || fetch(event.request);
    })
  );
});
