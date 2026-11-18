// «Ұлан digital» – қосымша (PWA) ретінде жұмыс істеуге арналған Service Worker.
// Беттің өзі әрқашан желіден алынады (жаңа нұсқа бірден көрінеді),
// желі болмаса – соңғы сақталған нұсқасы ашылады.
const CACHE = 'ulan-v2';
const SHELL = [
  '/',
  '/balaqai.png',
  '/balaqai_face.png',
  '/ulan_logo.png',
  '/kaspi_logo.png',
  '/icons/icon-192.png',
  '/icons/icon-512.png',
  '/manifest.webmanifest'
];

self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(SHELL)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', (e) => {
  e.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (e) => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  if (url.origin !== self.location.origin) return; // Google Script, Drive т.б. – тікелей желіден
  if (req.headers.has('range') || /\.(mp3|mp4|m4a|webm)$/i.test(url.pathname)) return; // аудио/бейне – браузердің өзі (Safari үшін маңызды)

  // Бет (HTML): алдымен желі, болмаса – кэш
  if (req.mode === 'navigate') {
    e.respondWith(
      fetch(req)
        .then((res) => {
          const copy = res.clone();
          if (res.ok && url.pathname === '/') caches.open(CACHE).then((c) => c.put('/', copy));
          return res;
        })
        .catch(() => caches.match('/'))
    );
    return;
  }

  // Суреттер мен файлдар: кэштен бірден, фонда жаңарту
  e.respondWith(
    caches.match(req).then((cached) => {
      const fresh = fetch(req)
        .then((res) => {
          if (res.ok) {
            const copy = res.clone();
            caches.open(CACHE).then((c) => c.put(req, copy));
          }
          return res;
        })
        .catch(() => cached);
      return cached || fresh;
    })
  );
});
