// Service Worker v20260610-006 — SE AUTODESTRUYE al activarse
// Limpia todos los cachés, se desregistra y fuerza recarga limpia
self.addEventListener('install', () => self.skipWaiting());
self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys()
      .then(keys => Promise.all(keys.map(k => caches.delete(k))))
      .then(() => self.registration.unregister())
      .then(() => self.clients.matchAll({ type: 'window' }))
      .then(clients => clients.forEach(c => c.navigate(c.url)))
  );
});
// No intercepta nada — el browser va directo a la red
self.addEventListener('fetch', () => {});
