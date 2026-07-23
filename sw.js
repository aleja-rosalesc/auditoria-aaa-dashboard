const CACHE_NAME = "aaa-auditoria-v5";

const APP_SHELL = [ "./index.html", "./manifest.json", "./icon-192.png", "./icon-512.png", "./logo.png" ];

self.addEventListener("install", function(event) {
  event.waitUntil(caches.open(CACHE_NAME).then(function(cache) {
    return cache.addAll(APP_SHELL);
  }));
  self.skipWaiting();
});

self.addEventListener("activate", function(event) {
  event.waitUntil(caches.keys().then(function(keys) {
    return Promise.all(keys.filter(function(k) {
      return k !== CACHE_NAME;
    }).map(function(k) {
      return caches.delete(k);
    }));
  }));
  self.clients.claim();
});

self.addEventListener("fetch", function(event) {
  if (event.request.method !== "GET") return;
  event.respondWith(fetch(event.request.url, {
    cache: "no-store"
  }).then(function(response) {
    const copy = response.clone();
    caches.open(CACHE_NAME).then(function(cache) {
      cache.put(event.request, copy);
    });
    return response;
  }).catch(function() {
    return caches.match(event.request);
  }));
});