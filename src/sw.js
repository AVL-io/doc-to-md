// Service worker: stores the whole app on first visit so it then works offline.
// It only ever caches the app's own files. Documents are never sent anywhere.
const VERSION = '__VERSION__';
const CACHE = 'doc-to-md-' + VERSION;
const FILES = __PRECACHE__;
const OCR_CACHE = 'doc-to-md-ocr-eng-1';   // OCR engine files: cached on first use, kept across app updates

self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(FILES.map((f) => new Request(f, { cache: 'reload' })))));
});

self.addEventListener('activate', (e) => {
  e.waitUntil((async () => {
    for (const k of await caches.keys()) if (k.startsWith('doc-to-md-') && k !== CACHE && k !== OCR_CACHE) await caches.delete(k);
    await self.clients.claim();
  })());
});

self.addEventListener('message', (e) => { if (e.data === 'skipWaiting') self.skipWaiting(); });

self.addEventListener('fetch', (e) => {
  const req = e.request;
  if (req.method !== 'GET' || new URL(req.url).origin !== self.location.origin) return;
  if (new URL(req.url).pathname.includes('/ocr/')) {
    e.respondWith((async () => {
      const cache = await caches.open(OCR_CACHE);
      const hit = await cache.match(req);
      if (hit) return hit;
      const res = await fetch(req);
      if (res.ok) await cache.put(req, res.clone());
      return res;
    })());
    return;
  }
  e.respondWith((async () => {
    const cache = await caches.open(CACHE);
    const hit = await cache.match(req, { ignoreSearch: true }) || (req.mode === 'navigate' ? await cache.match('./') || await cache.match('index.html') : null);
    if (hit) return hit;
    try { return await fetch(req); } catch { return new Response('Offline and not cached.', { status: 503 }); }
  })());
});
