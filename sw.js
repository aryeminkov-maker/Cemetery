// עובד רשת (Service Worker) לאתר בית העלמין בבית אל.
// שומר בקאש את קובץ האתר ואת תמונות האוויר של החלקות, כדי שחיפוש, מפה ודפי
// קברים בודדים ימשיכו לעבוד גם עם קליטה חלשה או בלי אינטרנט בכלל (על בסיס
// נתוני הגיבוי המוטמעים בקובץ). בקשות לגיליון החי (Google Apps Script) תמיד
// יוצאות לרשת בלבד ולא נשמרות בקאש, כדי לא להציג נתונים ישנים בטעות.

const CACHE_NAME = 'beit-el-cemetery-v1';
const CACHE_FILES = [
  './',
  './beit-el-cemetery.html',
  './photos/block-a.jpg',
  './photos/block-b.jpg',
  './photos/block-c.jpg'
];

self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then(cache => cache.addAll(CACHE_FILES))
      .catch(() => {}) // אם קובץ מסוים חסר (למשל תמונה שעוד לא הועלתה) - לא נכשיל את כל ההתקנה
  );
  self.skipWaiting();
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', event => {
  const url = new URL(event.request.url);

  // בקשות לגיליון החי - תמיד רשת, אף פעם לא קאש (נתונים חייבים להיות עדכניים)
  if (url.hostname.indexOf('script.google.com') !== -1) {
    return;
  }
  // בקשות לדומיינים חיצוניים אחרים (Google Sign-In וכו') - משאירים לדפדפן, לא מתערבים
  if (url.origin !== self.location.origin) {
    return;
  }

  event.respondWith(
    caches.match(event.request).then(cached => {
      const network = fetch(event.request).then(res => {
        if (res && res.ok) {
          const resClone = res.clone();
          caches.open(CACHE_NAME).then(cache => cache.put(event.request, resClone)).catch(() => {});
        }
        return res;
      }).catch(() => cached);
      return cached || network;
    })
  );
});
