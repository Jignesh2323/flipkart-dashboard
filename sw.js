// Minimal service worker so the dashboard installs as its own app (scope /flipkart-dashboard/).
// Always goes to the network so the daily data is never served stale.
self.addEventListener('install', () => self.skipWaiting());
self.addEventListener('activate', e => e.waitUntil(self.clients.claim()));
self.addEventListener('fetch', e => e.respondWith(fetch(e.request)));
