const CACHE_NAME = 'travel-planner-v2';
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
    event.waitUntil(clients.claim());
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
        caches.match(event.request).then((cachedResponse) => {
            if (cachedResponse) {
                return cachedResponse;
            }

            return fetch(event.request).then((response) => {
                // Cache valid responses (200) or opaque responses (0) for cross-origin assets like weather icons
                const isSuccess = response && response.status === 200;
                const isOpaque = response && response.type === 'opaque' && response.status === 0;
                const isAllowedType = response && (response.type === 'basic' || response.type === 'cors' || response.type === 'opaque');

                if (!response || (!isSuccess && !isOpaque) || !isAllowedType) {
                    return response;
                }

                const responseToCache = response.clone();
                caches.open(CACHE_NAME).then((cache) => {
                    cache.put(event.request, responseToCache);
                }).catch(() => {
                    // Ignore cache put errors (e.g. quota exceeded)
                });
                return response;
            }).catch((error) => {
                // Re-throw the error so the browser recognizes the network failure 
                // instead of reporting the request as "Aborted" by the Service Worker.
                throw error;
            });
        })
    );
});