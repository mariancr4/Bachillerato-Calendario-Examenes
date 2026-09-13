const CACHE_NAME = "examenes-mauxiliadora-v3";
const APP_SHELL = ["index.html", "manifest.json", "icon-192.png", "icon-512.png"];

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(APP_SHELL))
  );
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k)))
    )
  );
  self.clients.claim();
});

// Solo se cachea el "cascarón" estático de la web (mismo origen, GET).
// Las llamadas a la API de Apps Script y a Google Calendar viajan siempre
// directas a la red, nunca se sirven desde caché.
//
// Estrategia "red primero": se intenta siempre traer la versión más
// reciente del servidor; la copia en caché solo se usa como respaldo si
// no hay conexión. Así, al publicar cambios en la web, se ven de
// inmediato en el siguiente recarga en vez de quedarse con una versión
// vieja guardada en el dispositivo.
self.addEventListener("fetch", (event) => {
  const req = event.request;
  if (req.method !== "GET" || new URL(req.url).origin !== self.location.origin) return;

  event.respondWith(
    fetch(req)
      .then((res) => {
        const clone = res.clone();
        caches.open(CACHE_NAME).then((cache) => cache.put(req, clone));
        return res;
      })
      .catch(() => caches.match(req))
  );
});
