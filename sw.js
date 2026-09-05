// מטמון מהיר לקבצי הממשק בלבד. נתוני העסק ממשיכים להגיע תמיד מהשרת.
const STATIC_CACHE = 'amn-static-v4';

self.addEventListener('install', event => {
  self.skipWaiting();
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys()
      .then(keys => Promise.all(keys.filter(key => key !== STATIC_CACHE).map(key => caches.delete(key))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', event => {
  const request = event.request;
  if (request.method !== 'GET') return;

  const url = new URL(request.url);
  const sameOrigin = url.origin === self.location.origin;
  const isPage = request.mode === 'navigate';
  const isStatic = sameOrigin && /\.(?:png|jpg|jpeg|svg|webp|ico|json|css|js)$/i.test(url.pathname);

  // דפי האפליקציה: רשת תחילה, ובמצב ניתוק עותק אחרון.
  if (sameOrigin && isPage) {
    event.respondWith(
      fetch(request).then(response => {
        const copy = response.clone();
        caches.open(STATIC_CACHE).then(cache => cache.put(request, copy));
        return response;
      }).catch(() => caches.match(request))
    );
    return;
  }

  // תמונות וקבצי ממשק: הצגה מיידית מהמטמון ועדכון ברקע.
  if (isStatic) {
    event.respondWith(
      caches.open(STATIC_CACHE).then(async cache => {
        const cached = await cache.match(request);
        const network = fetch(request).then(response => {
          if (response.ok) cache.put(request, response.clone());
          return response;
        }).catch(() => cached);
        return cached || network;
      })
    );
  }
});
