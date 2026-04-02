const CACHE_NAME = 'maizuru-tour-v2';
const TILE_CACHE = 'maizuru-tiles-v2';

// Core resources to cache immediately
const CORE_RESOURCES = [
  './',
  './index.html',
  './manifest.json',
  'https://unpkg.com/leaflet@1.9.4/dist/leaflet.css',
  'https://unpkg.com/leaflet@1.9.4/dist/leaflet.js'
];

// Pre-calculate tile URLs for Maizuru area
// Bounds: lat 35.41-35.53, lon 135.27-135.43
function getTileUrls() {
  const tiles = [];
  const zoomLevels = [
    { z: 12, latRange: [35.38, 35.55], lonRange: [135.24, 135.45] },
    { z: 13, latRange: [35.40, 35.53], lonRange: [135.26, 135.43] },
    { z: 14, latRange: [35.41, 35.53], lonRange: [135.27, 135.43] },
    { z: 15, latRange: [35.42, 35.52], lonRange: [135.28, 135.42] },
    { z: 16, latRange: [35.43, 35.50], lonRange: [135.30, 135.40] }
  ];

  for (const { z, latRange, lonRange } of zoomLevels) {
    const n = Math.pow(2, z);
    const xMin = Math.floor((lonRange[0] + 180) / 360 * n);
    const xMax = Math.floor((lonRange[1] + 180) / 360 * n);
    const yMin = Math.floor((1 - Math.log(Math.tan(latRange[1] * Math.PI / 180) + 1 / Math.cos(latRange[1] * Math.PI / 180)) / Math.PI) / 2 * n);
    const yMax = Math.floor((1 - Math.log(Math.tan(latRange[0] * Math.PI / 180) + 1 / Math.cos(latRange[0] * Math.PI / 180)) / Math.PI) / 2 * n);

    for (let x = xMin; x <= xMax; x++) {
      for (let y = yMin; y <= yMax; y++) {
        // Japanese tiles (OpenStreetMap)
        tiles.push(`https://tile.openstreetmap.org/${z}/${x}/${y}.png`);
        // English tiles (CARTO Voyager)
        const s = ['a', 'b', 'c', 'd'][Math.floor(Math.random() * 4)];
        tiles.push(`https://${s}.basemaps.cartocdn.com/rastertiles/voyager/${z}/${x}/${y}.png`);
      }
    }
  }
  return tiles;
}

// Install: cache core resources and pre-fetch map tiles
self.addEventListener('install', event => {
  event.waitUntil(
    (async () => {
      // Cache core resources
      const coreCache = await caches.open(CACHE_NAME);
      await coreCache.addAll(CORE_RESOURCES);
      console.log('[SW] Core resources cached');

      // Cache map tiles in batches
      const tileCache = await caches.open(TILE_CACHE);
      const tileUrls = getTileUrls();
      console.log(`[SW] Pre-caching ${tileUrls.length} map tiles...`);

      // Fetch tiles in batches of 10 to avoid overwhelming the network
      const batchSize = 10;
      for (let i = 0; i < tileUrls.length; i += batchSize) {
        const batch = tileUrls.slice(i, i + batchSize);
        const results = await Promise.allSettled(
          batch.map(url =>
            fetch(url).then(resp => {
              if (resp.ok) return tileCache.put(url, resp);
            }).catch(() => {})
          )
        );
      }
      console.log('[SW] Map tiles cached');

      self.skipWaiting();
    })()
  );
});

// Activate: clean old caches
self.addEventListener('activate', event => {
  event.waitUntil(
    (async () => {
      const keys = await caches.keys();
      await Promise.all(
        keys.filter(key => key !== CACHE_NAME && key !== TILE_CACHE)
          .map(key => caches.delete(key))
      );
      self.clients.claim();
      console.log('[SW] Activated and claimed clients');
    })()
  );
});

// Fetch: serve from cache first (offline-first strategy)
self.addEventListener('fetch', event => {
  const url = event.request.url;

  // For tile requests, use cache-first
  if (url.includes('tile.openstreetmap.org') || url.includes('basemaps.cartocdn.com')) {
    event.respondWith(
      (async () => {
        const cached = await caches.match(event.request);
        if (cached) return cached;

        try {
          const response = await fetch(event.request);
          if (response.ok) {
            const tileCache = await caches.open(TILE_CACHE);
            tileCache.put(event.request, response.clone());
          }
          return response;
        } catch {
          // Return a transparent 1x1 PNG for missing tiles when offline
          return new Response(
            Uint8Array.from(atob('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mN88P/BfwAJhAPkv6JaagAAAABJRU5ErkJggg=='), c => c.charCodeAt(0)),
            { headers: { 'Content-Type': 'image/png' } }
          );
        }
      })()
    );
    return;
  }

  // For other requests, try cache first then network
  event.respondWith(
    (async () => {
      const cached = await caches.match(event.request);
      if (cached) return cached;

      try {
        const response = await fetch(event.request);
        if (response.ok) {
          const cache = await caches.open(CACHE_NAME);
          cache.put(event.request, response.clone());
        }
        return response;
      } catch {
        // Offline fallback
        if (event.request.destination === 'document') {
          return caches.match('./index.html');
        }
        return new Response('Offline', { status: 503 });
      }
    })()
  );
});
