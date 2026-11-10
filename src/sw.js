// Service worker: stores the whole app on first visit so it then works offline.
// It only ever caches the app's own files. Documents are never sent anywhere.
const VERSION = '__VERSION__';
const CACHE = 'doc-to-md-' + VERSION;
const FILES = __PRECACHE__;

self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(FILES.map((f) => new Request(f, { cache: 'reload' })))));
});

self.addEventListener('activate', (e) => {
  e.waitUntil((async () => {
    for (const k of await caches.keys()) if (k.startsWith('doc-to-md-') && k !== CACHE) await caches.delete(k);
    await self.clients.claim();
  })());
});

self.addEventListener('message', (e) => { if (e.data === 'skipWaiting') self.skipWaiting(); });

self.addEventListener('fetch', (e) => {
  const req = e.request;
  if (req.method !== 'GET' || new URL(req.url).origin !== self.location.origin) return;
  e.respondWith((async () => {
    const cache = await caches.open(CACHE);
    const hit = await cache.match(req, { ignoreSearch: true }) || (req.mode === 'navigate' ? await cache.match('./') || await cache.match('index.html') : null);
    if (hit) return hit;
    try { return await fetch(req); } catch { return new Response('Offline and not cached.', { status: 503 }); }
  })());
});
