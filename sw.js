// sw.js v4 – tvingar bort gammal cache för historik-fixen
const CACHE = 'sterilsok-v4';
const ASSETS = ['./', './index.html', './data.js', './manifest.webmanifest'];

self.addEventListener('install', e => {
  self.skipWaiting();
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(ASSETS)).catch(()=>{}));
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(()=> self.clients.claim())
  );
});

self.addEventListener('fetch', e => {
  const url = new URL(e.request.url);
  // Hämta alltid färsk index.html och data.js
  if (url.pathname.endsWith('index.html') || url.pathname.endsWith('/') || url.pathname.endsWith('data.js')) {
    e.respondWith(fetch(e.request, {cache:'no-store'}).then(r => {
      const clone = r.clone();
      caches.open(CACHE).then(c => c.put(e.request, clone)).catch(()=>{});
      return r;
    }).catch(()=> caches.match(e.request)));
    return;
  }
  e.respondWith(caches.match(e.request).then(r => r || fetch(e.request).catch(()=> caches.match('./index.html'))));
});

self.addEventListener('message', e => {
  if (e.data === 'skipWaiting') self.skipWaiting();
});
