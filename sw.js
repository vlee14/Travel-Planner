const CACHE_NAME = 'travel-planner-v4';
const STATIC_ASSETS = [
    './',
    'index.html',
    'styles.css',
    'app.js',
    'https://unpkg.com/leaflet@1.9.4/dist/leaflet.css',
    'https://unpkg.com/leaflet@1.9.4/dist/leaflet.js',
    'https://cdn.jsdelivr.net/npm/marked/marked.min.js',
    'https://fonts.googleapis.com/css2?family=DM+Sans:ital,opsz,wght@0,9..40,400;0,9..40,500;0,9..40,600;0,9..40,700&display=swap'
];

self.addEventListener('install', (event) => {
    self.skipWaiting();
    event.waitUntil(
        caches.open(CACHE_NAME).then((cache) => cache.addAll(STATIC_ASSETS))
    );
});

self.addEventListener('activate', (event) => {
    // Delete old versions of the cache
    event.waitUntil(
        caches.keys().then((cacheNames) => {
            return Promise.all(
                cacheNames.filter((name) => name !== CACHE_NAME)
                    .map((name) => caches.delete(name))
            );
        }).then(() => {
            return clients.claim();
        })
    );
});

self.addEventListener('fetch', (event) => {
    // Skip requests with unsupported schemes (like chrome-extension:// or data:)
    if (!event.request.url.startsWith('http')) {
        return;
    }

    const url = new URL(event.request.url);

    // Exclude API calls from Service Worker cache
    // (The app handles these with localStorage logic for freshness/expiration)
    if (url.hostname.includes('accounts.google.com')) return;
    if (url.hostname.includes('apis.google.com')) return;
    if (url.hostname.includes('googleapis.com') && url.pathname.includes('generateContent')) return;
    if (url.hostname.includes('visualcrossing.com')) return;
    if (url.hostname.includes('open-meteo.com')) return;
    if (url.hostname.includes('tinyurl.com')) return;

    event.respondWith(
        (async () => {
            const cache = await caches.open(CACHE_NAME);

            // Network-first for navigation requests (the main HTML page)
            // This ensures users get the latest UI if they are online.
            if (event.request.mode === 'navigate') {
                try {
                    const networkResponse = await fetch(event.request);
                    cache.put(event.request, networkResponse.clone());
                    return networkResponse;
                } catch (e) {
                    return cache.match(event.request);
                }
            }

            // Stale-while-revalidate for assets (CSS, JS, etc.)
            const cachedResponse = await cache.match(event.request);
            const fetchPromise = fetch(event.request).then((networkResponse) => {
                if (networkResponse && networkResponse.status === 200 && networkResponse.type === 'basic') {
                    cache.put(event.request, networkResponse.clone());
                }
                return networkResponse;
            }).catch(() => {
                // Fail silently, the cachedResponse will be returned
            });

            return cachedResponse || fetchPromise;
        })()
    );
});