// Service Worker — desinstala versiones anteriores y no cachea nada
self.addEventListener('install', () => self.skipWaiting());
self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys => Promise.all(keys.map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});
// Sin caché — todo va directo a la red
self.addEventListener('fetch', e => {
  e.respondWith(fetch(e.request, { cache: 'no-store' }).catch(() => caches.match(e.request)));
});
