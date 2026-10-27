const CACHE_NAME = "aaa-auditoria-v2";
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
  event.respondWith(
    fetch(event.request.url, { cache: "no-store" }).then(function (response) {
      const copy = response.clone();
      caches.open(CACHE_NAME).then(function (cache) { cache.put(event.request, copy); });
      return response;
    }).catch(function () { return caches.match(event.request); })
  );
});

// Notificaciones push (Firebase Cloud Messaging): llegan con la app cerrada
self.addEventListener("push", function (event) {
  let p = {};
  try { p = event.data ? event.data.json() : {}; } catch (e) {}
  const d = p.data || p.notification || p;
  const titulo = d.titulo || d.title || "Auditoría AAA", cuerpo = d.cuerpo || d.body || "";
  event.waitUntil(self.clients.matchAll({ type: "window", includeUncontrolled: true }).then(function (cs) {
    cs.forEach(function (c) { c.postMessage({ tipo: "push", titulo: titulo, cuerpo: cuerpo }); });
    if (cs.some(function (c) { return c.visibilityState === "visible"; })) return;   // con la app a la vista, la propia app avisa
    return self.registration.showNotification(titulo, { body: cuerpo, icon: "icon-192.png", badge: "icon-192.png", tag: d.tag || "aud", renotify: true, data: { url: "./" } });
  }));
});
self.addEventListener("notificationclick", function (event) {
  event.notification.close();
  event.waitUntil(self.clients.matchAll({ type: "window", includeUncontrolled: true }).then(function (cs) {
    for (const c of cs) { if ("focus" in c) return c.focus(); }
    return self.clients.openWindow("./");
  }));
});
