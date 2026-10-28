// sw.js v10 – hämtar alltid färsk index.html och data.js när det finns nät, och använder den sparade kopian utan nät.
const CACHE = 'sterilsok-v10';
const ASSETS = ['./', 'index.html', 'data.js', 'manifest.webmanifest', 'icon-192.png', 'icon-512.png', 'apple-touch-icon.png'];

self.addEventListener('install', e => {
  self.skipWaiting();
  // En fil i taget: saknas en av dem (t.ex. en ikon) ska inte hela cachen bli tom.
  e.waitUntil(caches.open(CACHE).then(c =>
    Promise.all(ASSETS.map(f => c.add(new Request(f, { cache: 'reload' })).catch(() => {})))
  ));
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', e => {
  const req = e.request;
  const url = new URL(req.url);
  // Rör inte statistikanrop (POST) eller andra adresser än den egna sidan.
  if (req.method !== 'GET' || url.origin !== self.location.origin) return;

  const farsk = url.pathname.endsWith('index.html') || url.pathname.endsWith('/') || url.pathname.endsWith('data.js');
  if (farsk) {
    e.respondWith(
      fetch(req, { cache: 'no-store' }).then(r => {
        if (r && r.ok) { const kopia = r.clone(); caches.open(CACHE).then(c => c.put(req, kopia)).catch(() => {}); }
        return r;
      }).catch(async () => (await caches.match(req, { ignoreSearch: true })) || (await caches.match('index.html')) || Response.error())
    );
    return;
  }
  e.respondWith(
    caches.match(req, { ignoreSearch: true }).then(r => r || fetch(req).then(svar => {
      if (svar && svar.ok) { const kopia = svar.clone(); caches.open(CACHE).then(c => c.put(req, kopia)).catch(() => {}); }
      return svar;
    }).catch(async () => req.mode === 'navigate' ? ((await caches.match('index.html')) || Response.error()) : Response.error()))
  );
});

self.addEventListener('message', e => {
  if (e.data === 'skipWaiting') self.skipWaiting();
});
