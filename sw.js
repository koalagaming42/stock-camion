// Moteur hors ligne de l'appli Stock Camion.
// - L'appli est d'abord chargée depuis Internet (pour recevoir les mises à jour),
//   et depuis la copie enregistrée sur le téléphone quand il n'y a pas de réseau.
// - Les données (stock, historique, réglages) ne passent JAMAIS par ici :
//   elles restent uniquement sur le téléphone.
const CACHE = 'stock-camion-v1';
const FILES = ['./', './index.html', './manifest.webmanifest',
  './icon-192.png', './icon-512.png', './icon-maskable-512.png', './apple-touch-icon.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(FILES)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET' || new URL(req.url).origin !== location.origin) return;

  // Pages : réseau d'abord (max 4 s), sinon la copie hors ligne
  if (req.mode === 'navigate') {
    e.respondWith((async () => {
      try {
        const res = await Promise.race([
          fetch(req, {cache: 'no-store'}),
          new Promise((_, rej) => setTimeout(() => rej(new Error('timeout')), 4000))
        ]);
        if (res && res.ok) { const c = await caches.open(CACHE); c.put('./index.html', res.clone()); }
        return res;
      } catch (err) {
        return (await caches.match('./index.html')) || (await caches.match('./')) || Response.error();
      }
    })());
    return;
  }

  // Icônes, manifeste : copie locale d'abord, mise à jour en arrière-plan
  e.respondWith(
    caches.match(req).then(cached => {
      const net = fetch(req).then(res => {
        if (res && res.ok) caches.open(CACHE).then(c => c.put(req, res.clone()));
        return res;
      }).catch(() => cached);
      return cached || net;
    })
  );
});
