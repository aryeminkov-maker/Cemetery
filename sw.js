// עובד רשת (Service Worker) לאתר בית העלמין בבית אל.
// שומר בקאש את קובץ האתר ואת תמונות האוויר של החלקות, כדי שחיפוש, מפה ודפי
// קברים בודדים ימשיכו לעבוד גם עם קליטה חלשה או בלי אינטרנט בכלל (על בסיס
// נתוני הגיבוי המוטמעים בקובץ). בקשות לגיליון החי (Google Apps Script) תמיד
// יוצאות לרשת בלבד ולא נשמרות בקאש, כדי לא להציג נתונים ישנים בטעות.

const CACHE_NAME = 'beit-el-cemetery-v3';
// רק קבצי הליבה נשמרים מראש בהתקנה. תצלומי האוויר הגדולים של החלקות לא נטענים
// כאן (הם היו מתחרים בביקור הראשון בתמונת הרקע ומאטים אותה) - הם נשמרים בקאש
// אוטומטית בפעם הראשונה שנכנסים לחלקה, ומשם הם זמינים גם בלי אינטרנט.
const CACHE_FILES = [
  './',
  './beit-el-cemetery.html',
  './photos/hero.webp',
  './manifest.json',
  './photos/icon-192.png'
];

self.addEventListener('install', event => {
  // כל קובץ נשמר בנפרד - קובץ חסר (למשל תמונה שעוד לא הועלתה, או שם קובץ שונה
  // כמו index.html) לא מפיל את שמירת שאר הקבצים בקאש.
  event.waitUntil(
    caches.open(CACHE_NAME).then(cache =>
      Promise.all(CACHE_FILES.map(f => cache.add(f).catch(() => {})))
    )
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

  const isAppShell = event.request.mode === 'navigate' || url.pathname.endsWith('.html');

  if (isAppShell) {
    // קובץ ה-HTML עצמו - קודם רשת (network-first), כדי שעדכון שהועלה יגיע מיד
    // למבקרים, ורק אם אין בכלל אינטרנט נופלים לגרסה השמורה בקאש.
    event.respondWith(
      fetch(event.request).then(res => {
        if (res && res.ok) {
          const resClone = res.clone();
          caches.open(CACHE_NAME).then(cache => cache.put(event.request, resClone)).catch(() => {});
        }
        return res;
      }).catch(() => caches.match(event.request))
    );
    return;
  }

  // קבצים אחרים (בעיקר תמונות) - קודם קאש (מהיר), עם עדכון ברקע לפעם הבאה
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
