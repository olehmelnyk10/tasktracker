// TaskTracker PWA Service Worker v5
// Рамедас Україна

var CACHE = 'tt-v6';
var ASSETS = [
  '/',
  '/index.html',
  'https://fonts.googleapis.com/css2?family=Manrope:wght@400;500;600;700;800;900&family=JetBrains+Mono:wght@400;600&display=swap'
];

// Встановлення — кешуємо основні файли
self.addEventListener('install', function(e) {
  e.waitUntil(
    caches.open(CACHE).then(function(c) {
      return c.addAll(ASSETS).catch(function(err) {
        console.log('SW cache failed:', err);
      });
    })
  );
  self.skipWaiting();
});

// Активація — видаляємо старі кеші
self.addEventListener('activate', function(e) {
  e.waitUntil(
    caches.keys().then(function(keys) {
      return Promise.all(
        keys.filter(function(k) { return k !== CACHE; })
            .map(function(k) { return caches.delete(k); })
      );
    })
  );
  self.clients.claim();
});

// Fetch — API запити завжди через мережу, решта — кеш або мережа
self.addEventListener('fetch', function(e) {
  var url = e.request.url;

  // GAS API запити — завжди мережа (не кешуємо)
  if (url.indexOf('script.google.com') !== -1) {
    return;
  }

  // Google Fonts — мережа спочатку, потім кеш
  if (url.indexOf('fonts.googleapis.com') !== -1 || url.indexOf('fonts.gstatic.com') !== -1) {
    e.respondWith(
      fetch(e.request).then(function(r) {
        var copy = r.clone();
        caches.open(CACHE).then(function(c) { c.put(e.request, copy); });
        return r;
      }).catch(function() {
        return caches.match(e.request);
      })
    );
    return;
  }

  // Основний додаток — кеш спочатку, при помилці fallback на index.html
  e.respondWith(
    caches.match(e.request).then(function(cached) {
      if (cached) return cached;
      return fetch(e.request).then(function(r) {
        // Кешуємо статичні файли
        if (e.request.method === 'GET' && r.status === 200) {
          var copy = r.clone();
          caches.open(CACHE).then(function(c) { c.put(e.request, copy); });
        }
        return r;
      }).catch(function() {
        // Офлайн fallback
        return caches.match('/index.html');
      });
    })
  );
});
