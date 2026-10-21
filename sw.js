// Offline-stöd och snabb start. Sidan och databasen visas direkt från den sparade kopian,
// och hämtas samtidigt på nytt i bakgrunden. Är något nytt laddas sidan om automatiskt
// (om man inte håller på att söka), annars gäller det nya nästa gång appen öppnas.
const CACHE = 'lagersok-v3';
const FILER = ['./', 'index.html', 'data.js', 'manifest.webmanifest', 'icon-192.png', 'icon-512.png', 'apple-touch-icon.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(FILER.map(f => new Request(f, { cache: 'reload' })))).then(() => self.skipWaiting()));
});

self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(n => Promise.all(n.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});

async function uppdateraIBakgrunden(req) {
  try {
    const svar = await fetch(req, { cache: 'no-cache' });
    if (!svar || !svar.ok) return;
    const cache = await caches.open(CACHE);
    const gammal = await cache.match(req, { ignoreSearch: true });
    let nytt = true;
    if (gammal) {
      try { nytt = (await gammal.clone().text()) !== (await svar.clone().text()); } catch (_) {}
    }
    await cache.put(req, svar.clone());
    if (nytt && gammal) {
      const klienter = await self.clients.matchAll({ type: 'window' });
      klienter.forEach(k => k.postMessage('ny-version'));
    }
  } catch (_) { /* inget nät - den sparade kopian används */ }
}

self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET' || new URL(req.url).origin !== self.location.origin) return;
  e.respondWith((async () => {
    const sparad = await caches.match(req, { ignoreSearch: true });
    if (sparad) {
      e.waitUntil(uppdateraIBakgrunden(req));   // hämta nytt i bakgrunden
      return sparad;                            // visa direkt
    }
    try {
      const svar = await fetch(req, { cache: 'no-cache' });
      if (svar && svar.ok) (await caches.open(CACHE)).put(req, svar.clone());
      return svar;
    } catch (_) {
      if (req.mode === 'navigate') return caches.match('index.html');
      return Response.error();
    }
  })());
});
