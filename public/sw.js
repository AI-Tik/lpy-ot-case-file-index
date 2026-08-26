const CACHE_NAME = 'lpy-case-index-v16';
const BASE_PATH = self.location.pathname.replace(/\/sw\.js$/, '');
const appUrl = (path) => `${BASE_PATH}${path}`;
const ROOT_URL = appUrl('/');
const CORE_FILES = [
  ROOT_URL,
  appUrl('/offline.html'),
  appUrl('/manifest.webmanifest'),
  appUrl('/icon-192.png'),
  appUrl('/icon-512.png'),
  appUrl('/apple-touch-icon.png'),
  appUrl('/ot-case-file-index-template.xlsx'),
];

async function store(request, response) {
  if (!response || !response.ok) return response;
  const cache = await caches.open(CACHE_NAME);
  await cache.put(request, response.clone());
  return response;
}

async function fetchWithTimeout(request, timeout = 4000) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeout);
  try {
    return await fetch(request, { signal: controller.signal });
  } finally {
    clearTimeout(timer);
  }
}

async function precacheApp() {
  const cache = await caches.open(CACHE_NAME);
  await Promise.allSettled(CORE_FILES.slice(1).map(async (url) => {
    const response = await fetch(url, { cache: 'reload' });
    if (response.ok) await cache.put(url, response);
  }));

  const page = await fetch(ROOT_URL, { cache: 'reload' });
  if (!page.ok) return;
  await cache.put(ROOT_URL, page.clone());
  const html = await page.text();
  const urls = [...html.matchAll(/(?:src|href)=["'](\/[^"']+)["']/g)]
    .map((match) => match[1])
    .filter((url) => !url.startsWith('/__debug'));
  await Promise.allSettled([...new Set(urls)].map(async (url) => {
    const response = await fetch(url, { cache: 'reload' });
    if (response.ok) await cache.put(url, response);
  }));
}

self.addEventListener('install', (event) => {
  event.waitUntil(precacheApp().then(() => self.skipWaiting()));
});

self.addEventListener('activate', (event) => {
  event.waitUntil((async () => {
    const keys = await caches.keys();
    await Promise.all(keys.filter((key) => key.startsWith('lpy-case-index-') && key !== CACHE_NAME)
      .map((key) => caches.delete(key)));
    await self.clients.claim();
  })());
});

self.addEventListener('fetch', (event) => {
  const { request } = event;
  if (request.method !== 'GET') return;
  const url = new URL(request.url);
  if (url.origin !== self.location.origin || url.pathname.startsWith('/__debug')) return;

  if (request.mode === 'navigate') {
    event.respondWith((async () => {
      const cached = (await caches.match(request)) || (await caches.match(ROOT_URL));
      if (cached) {
        event.waitUntil(fetchWithTimeout(request)
          .then((response) => store(request, response))
          .catch(() => undefined));
        return cached;
      }
      try {
        return await store(request, await fetchWithTimeout(request));
      } catch {
        return await caches.match(appUrl('/offline.html'));
      }
    })());
    return;
  }

  event.respondWith((async () => {
    const cached = await caches.match(request);
    if (cached) return cached;
    try {
      return await store(request, await fetch(request));
    } catch {
      return new Response('', { status: 504, statusText: 'Offline' });
    }
  })());
});
