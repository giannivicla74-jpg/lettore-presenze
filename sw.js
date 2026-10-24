// Service worker del Lettore Presenze AI: conserva i file dell'app per aprirla anche con rete lenta o assente.
// Strategia "prima la rete": quando c'è internet si scarica sempre la versione aggiornata dell'app,
// altrimenti si usa la copia conservata. Le chiamate a Gemini e i file Excel non passano di qui.
const CACHE = 'lettore-presenze-v9';
const FILES = [
  './', 'index.html', 'style.css', 'app.js', 'excel.js', 'cartella.js', 'manifest.webmanifest',
  'libs/pdf.min.js', 'libs/pdf.worker.min.js', 'libs/jszip.min.js',
  'icons/icon-192.png', 'icons/icon-512.png', 'icons/icon-maskable-512.png', 'icons/logo.png'
];

self.addEventListener('install', event => {
  event.waitUntil(caches.open(CACHE).then(cache => cache.addAll(FILES)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys()
      .then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', event => {
  const url = new URL(event.request.url);
  if (event.request.method !== 'GET' || url.origin !== self.location.origin) return; // Gemini e altri siti: diretti
  event.respondWith(
    fetch(event.request)
      .then(response => {
        const copy = response.clone();
        caches.open(CACHE).then(cache => cache.put(event.request, copy));
        return response;
      })
      .catch(() => caches.match(event.request, { ignoreSearch: true }))
  );
});
