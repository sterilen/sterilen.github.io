// Gör att sidan fungerar utan täckning. Hämtar alltid senaste versionen när det finns nät,
// och använder annars den senast sparade kopian.
const CACHE = 'lagersok-v2';
const FILER = ['./', 'index.html', 'data.js', 'manifest.webmanifest', 'icon-192.png', 'icon-512.png', 'apple-touch-icon.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(FILER.map(f => new Request(f, { cache: 'reload' })))).then(() => self.skipWaiting()));
});

self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(n => Promise.all(n.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});

self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET' || new URL(req.url).origin !== self.location.origin) return;
  e.respondWith((async () => {
    try {
      const svar = await fetch(req, { cache: 'no-cache' });
      if (svar && svar.ok) (await caches.open(CACHE)).put(req, svar.clone());
      return svar;
    } catch (_) {
      const sparad = await caches.match(req, { ignoreSearch: true });
      if (sparad) return sparad;
      if (req.mode === 'navigate') return caches.match('index.html');
      return Response.error();
    }
  })());
});
