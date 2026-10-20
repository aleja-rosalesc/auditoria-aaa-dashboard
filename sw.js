const CACHE_NAME = "aaa-entrada-v1";
const APP_SHELL = ["./index.html", "./manifest.json", "./icon-192.png", "./icon-512.png", "./logo.png"];

self.addEventListener("install", function (event) {
  event.waitUntil(caches.open(CACHE_NAME).then(function (cache) { return cache.addAll(APP_SHELL); }));
  self.skipWaiting();
});

self.addEventListener("activate", function (event) {
  event.waitUntil(
    caches.keys().then(function (keys) {
      return Promise.all(keys.filter(function (k) { return k !== CACHE_NAME; }).map(function (k) { return caches.delete(k); }));
    })
  );
  self.clients.claim();
});

self.addEventListener("fetch", function (event) {
  if (event.request.method !== "GET") return;
  // Solo controla la pantalla de entrada; no intercepta pedidos hacia /critica/ ni /pregeofono/,
  // esos los maneja el service worker propio de cada tablero.
  const url = new URL(event.request.url);
  if (url.pathname.indexOf("/critica/") !== -1 || url.pathname.indexOf("/pregeofono/") !== -1 || url.pathname.indexOf("/auditoria/") !== -1) return;
  event.respondWith(
    fetch(event.request.url, { cache: "no-store" }).then(function (response) {
      const copy = response.clone();
      caches.open(CACHE_NAME).then(function (cache) { cache.put(event.request, copy); });
      return response;
    }).catch(function () { return caches.match(event.request); })
  );
});
