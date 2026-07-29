// Service Worker מינימלי - נדרש כדי שדפדפני אנדרואיד יאפשרו "הוספה למסך הבית"
// כמו אפליקציה אמיתית. לא שומר מידע במטמון בכוונה - הנתונים מהגיליון
// תמיד צריכים להיות עדכניים, לכן פשוט מעביר כל בקשה ישירות לרשת.
self.addEventListener('install', (event) => {
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  self.clients.claim();
});

self.addEventListener('fetch', (event) => {
  event.respondWith(fetch(event.request));
});
