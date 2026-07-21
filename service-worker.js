const CACHE_NAME = 'trimension-version-1.1.5';
const LOCAL_ASSETS = [
  './',
  './index.html',
  './style.css',
  './app.js',
  './manifest.json',
  './images/titleLogo.png',
  './images/panelLogo.png',
  './images/favicon-cube.svg',
  './images/trimensionIcon-512.png'
];

const CDN_ASSETS = [
  'https://cdn.jsdelivr.net/npm/three@0.160.0/build/three.module.js',
  'https://cdn.jsdelivr.net/npm/three@0.160.0/examples/jsm/controls/OrbitControls.js',
  'https://cdn.jsdelivr.net/npm/three@0.160.0/examples/jsm/lines/Line2.js',
  'https://cdn.jsdelivr.net/npm/three@0.160.0/examples/jsm/lines/LineGeometry.js',
  'https://cdn.jsdelivr.net/npm/three@0.160.0/examples/jsm/lines/LineMaterial.js'
];

const NAVIGATION_NETWORK_TIMEOUT_MS = 1800;
const ASSET_NETWORK_TIMEOUT_MS = 4000;
const BACKGROUND_REFRESH_DELAY_MS = 8000;

function fetchWithTimeout(request, timeoutMs) {
  const controller = new AbortController();
  const timeoutId = setTimeout(() => controller.abort(), timeoutMs);

  return fetch(request, { signal: controller.signal })
    .finally(() => clearTimeout(timeoutId));
}

async function cacheFirstWithBackgroundRefresh(request, options = {}, networkTimeoutMs = ASSET_NETWORK_TIMEOUT_MS) {
  const cache = await caches.open(CACHE_NAME);
  const cached = await cache.match(request, options);

  const networkPromise = fetchWithTimeout(request, networkTimeoutMs)
    .then((response) => {
      if (response && response.ok) {
        cache.put(request, response.clone());
      }
      return response;
    })
    .catch(() => null);

  if (cached) {
    return { response: cached, background: networkPromise };
  }

  const networkResponse = await networkPromise;
  if (networkResponse) {
    return { response: networkResponse, background: null };
  }

  return { response: null, background: null };
}

async function updateCacheInBackground(request, timeoutMs) {
  try {
    const cache = await caches.open(CACHE_NAME);
    const response = await fetchWithTimeout(request, timeoutMs);
    if (response && response.ok) {
      await cache.put(request, response.clone());
    }
  } catch {
    // Ignore background refresh errors.
  }
}

function toScopeUrl(path) {
  return new URL(path, self.registration.scope).href;
}

async function getCachedAppShell() {
  const cache = await caches.open(CACHE_NAME);
  const candidates = [
    toScopeUrl('./index.html'),
    toScopeUrl('./')
  ];

  for (const candidate of candidates) {
    const match = await cache.match(candidate, { ignoreSearch: true });
    if (match) {
      return match;
    }
  }

  return null;
}

self.addEventListener('install', (event) => {
  event.waitUntil((async () => {
    const cache = await caches.open(CACHE_NAME);

    // Local shell is mandatory so launch is fast even with poor connectivity.
    await cache.addAll(LOCAL_ASSETS.map((asset) => toScopeUrl(asset)));

    // External CDN files are best-effort: don't block install when network is weak.
    await Promise.allSettled(CDN_ASSETS.map((url) => cache.add(url)));
  })());

  self.skipWaiting();
});

self.addEventListener('message', (event) => {
  if (event.data && event.data.type === 'SKIP_WAITING') {
    self.skipWaiting();
  }
});

self.addEventListener('fetch', (event) => {
  const { request } = event;

  if (request.method !== 'GET') {
    return;
  }

  const url = new URL(request.url);
  const isSameOrigin = url.origin === self.location.origin;
  const isNavigation = request.mode === 'navigate';
  const isStaticAsset = ['script', 'style', 'image', 'font'].includes(request.destination);
  const isJsDelivr = url.origin === 'https://cdn.jsdelivr.net';

  // For same-origin navigations, prefer network so HTML updates are picked up quickly.
  if (isNavigation && isSameOrigin) {
    event.respondWith((async () => {
      const cache = await caches.open(CACHE_NAME);
      try {
        const networkResponse = await fetchWithTimeout(request, NAVIGATION_NETWORK_TIMEOUT_MS);
        if (networkResponse && networkResponse.ok) {
          cache.put(request, networkResponse.clone());
        }
        return networkResponse;
      } catch {
        const cachedNavigation = await cache.match(request, { ignoreSearch: true });
        if (cachedNavigation) {
          return cachedNavigation;
        }

        const shell = await getCachedAppShell();
        if (shell) {
          return shell;
        }

        throw new Error('Offline and no cached app shell available');
      }
    })());

    return;
  }

  // Cache-first for local static files and CDN modules, then refresh in background.
  if ((isSameOrigin && isStaticAsset) || isJsDelivr) {
    event.respondWith((async () => {
      const cache = await caches.open(CACHE_NAME);
      const cachedAsset = await cache.match(request, { ignoreSearch: true });
      if (cachedAsset) {
        // CDN assets are version-pinned and never change, so skip background refresh to avoid
        // wasting bandwidth on slow connections. Only refresh local assets, and delay to avoid
        // competing with the initial page load.
        if (isSameOrigin) {
          setTimeout(() => updateCacheInBackground(request, ASSET_NETWORK_TIMEOUT_MS), BACKGROUND_REFRESH_DELAY_MS);
        }
        return cachedAsset;
      }

      const networkResponse = await fetchWithTimeout(request, ASSET_NETWORK_TIMEOUT_MS);
      if (networkResponse && networkResponse.ok) {
        cache.put(request, networkResponse.clone());
      }
      return networkResponse;
    })());

    return;
  }

  // Default behavior for other GET requests: network first with cache fallback.
  event.respondWith(
    fetchWithTimeout(request, ASSET_NETWORK_TIMEOUT_MS)
      .then((response) => {
        if (response && response.ok && isSameOrigin) {
          const responseClone = response.clone();
          caches.open(CACHE_NAME).then((cache) => cache.put(request, responseClone));
        }

        return response;
      })
      .catch(async () => {
        const cached = await caches.match(request, { ignoreSearch: true });
        if (cached) {
          return cached;
        }

        if (isNavigation) {
          const shell = await getCachedAppShell();
          if (shell) {
            return shell;
          }
        }

        throw new Error('Request failed and no cache fallback found');
      })
  );
});

self.addEventListener('activate', (event) => {
  event.waitUntil((async () => {
    const cacheNames = await caches.keys();
    await Promise.all(
      cacheNames
        .filter((cacheName) => cacheName !== CACHE_NAME)
        .map((cacheName) => caches.delete(cacheName))
    );
  })());

  self.clients.claim();
});
