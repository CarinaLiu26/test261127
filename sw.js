// Service worker: app shell available offline; CDN libraries and photos cached as they are used.
const VERSION = 'hojian-v1';
const SHELL = ['./', './index.html', './manifest.webmanifest', './icons/icon-192.png', './icons/icon-512.png', './icons/apple-touch-icon.png', './icons/favicon-48.png'];
const RUNTIME_HOSTS = ['cdnjs.cloudflare.com', 'unpkg.com', 'cdn.tailwindcss.com', 'images.unsplash.com', 'i.postimg.cc'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(VERSION).then(c => c.addAll(SHELL)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', e => {
  e.waitUntil(caches.keys()
    .then(keys => Promise.all(keys.filter(k => k !== VERSION).map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
});

self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);

  // Pages: try the network first so updates show up, fall back to the cached copy offline.
  if (req.mode === 'navigate') {
    e.respondWith(fetch(req).then(res => {
      const copy = res.clone();
      caches.open(VERSION).then(c => c.put('./index.html', copy));
      return res;
    }).catch(() => caches.match(req).then(r => r || caches.match('./index.html'))));
    return;
  }

  // Own files and known CDNs: serve from cache immediately, refresh in the background.
  if (url.origin === self.location.origin || RUNTIME_HOSTS.includes(url.hostname)) {
    e.respondWith(caches.open(VERSION).then(async cache => {
      const cached = await cache.match(req);
      const fresh = fetch(req).then(res => {
        if (res.ok || res.type === 'opaque') cache.put(req, res.clone());
        return res;
      }).catch(() => cached);
      return cached || fresh;
    }));
  }
});
