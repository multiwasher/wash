/* Service Worker - Gerador de Relatórios MultiWasher (PWA)
 * Estratégias:
 *  - navegação (HTML): network-first com fallback para o index.html em cache
 *  - estáticos da própria origem: cache-first
 *  - CDNs (Tailwind, Font Awesome, Google Fonts, Cloudinary): stale-while-revalidate
 * Ao publicar uma nova versão do app, incrementar APP_VERSION.
 */
const APP_VERSION = 'v1';
const PRECACHE = `mw-precache-${APP_VERSION}`;
const RUNTIME = `mw-runtime-${APP_VERSION}`;

const PRECACHE_URLS = [
  './',
  './index.html',
  './manifest.webmanifest',
  './icon-192.png',
  './icon-512.png',
  './icon-maskable-512.png',
  './apple-touch-icon.png',
  './favicon-32.png',
  './favicon-16.png',
];

const RUNTIME_HOSTS = [
  'cdn.tailwindcss.com',
  'cdnjs.cloudflare.com',
  'fonts.googleapis.com',
  'fonts.gstatic.com',
  'res.cloudinary.com',
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(PRECACHE)
      .then((cache) => cache.addAll(PRECACHE_URLS))
      .then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(
        keys.filter((k) => k !== PRECACHE && k !== RUNTIME).map((k) => caches.delete(k))
      ))
      .then(() => self.clients.claim())
  );
});

// Permite forçar a ativação a partir da página
self.addEventListener('message', (event) => {
  if (event.data === 'SKIP_WAITING') self.skipWaiting();
});

self.addEventListener('fetch', (event) => {
  const { request } = event;
  if (request.method !== 'GET') return;

  const url = new URL(request.url);

  // Navegações / documentos HTML -> network-first
  if (request.mode === 'navigate' || (request.destination === 'document')) {
    event.respondWith(
      fetch(request)
        .then((res) => {
          const copy = res.clone();
          caches.open(PRECACHE).then((c) => c.put('./index.html', copy)).catch(() => {});
          return res;
        })
        .catch(() => caches.match('./index.html', { ignoreSearch: true })
          .then((cached) => cached || caches.match('./')))
    );
    return;
  }

  // Mesma origem -> cache-first
  if (url.origin === self.location.origin) {
    event.respondWith(
      caches.match(request).then((cached) => cached || fetch(request).then((res) => {
        const copy = res.clone();
        caches.open(RUNTIME).then((c) => c.put(request, copy)).catch(() => {});
        return res;
      }))
    );
    return;
  }

  // CDNs conhecidos -> stale-while-revalidate
  if (RUNTIME_HOSTS.includes(url.hostname)) {
    event.respondWith(
      caches.open(RUNTIME).then((cache) => cache.match(request).then((cached) => {
        const network = fetch(request)
          .then((res) => {
            if (res && (res.ok || res.type === 'opaque')) {
              cache.put(request, res.clone()).catch(() => {});
            }
            return res;
          })
          .catch(() => cached);
        return cached || network;
      }))
    );
  }
});
