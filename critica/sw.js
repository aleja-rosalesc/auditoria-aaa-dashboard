importScripts("https://www.gstatic.com/firebasejs/10.14.1/firebase-app-compat.js");
importScripts("https://www.gstatic.com/firebasejs/10.14.1/firebase-messaging-compat.js");

firebase.initializeApp({
  apiKey: "AIzaSyALAAWgglH-BPJpNOwTEaLz1N3EgTRV4UU",
  authDomain: "inspecciones-aaa.firebaseapp.com",
  projectId: "inspecciones-aaa",
  storageBucket: "inspecciones-aaa.firebasestorage.app",
  messagingSenderId: "417714382892",
  appId: "1:417714382892:web:7d8eac84e582823b06c989"
});
const messaging = firebase.messaging();
messaging.onBackgroundMessage(function(payload) {
  const title = payload.notification && payload.notification.title || "Crítica";
  const body = payload.notification && payload.notification.body || "";
  self.registration.showNotification(title, { body: body, icon: "./icon-192.png", badge: "./icon-192.png" });
});

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