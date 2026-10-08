// Minimal service worker so the dashboard installs as its own app (scope /flipkart-dashboard/).
// v3: page + data scripts bypass the HTTP cache, so the daily data is never served stale.
self.addEventListener('install', () => self.skipWaiting());
self.addEventListener('activate', e => e.waitUntil(self.clients.claim()));
self.addEventListener('fetch', e => {
  const u = new URL(e.request.url);
  const fresh = u.origin === location.origin && /(\.js|\.html|\/)$/.test(u.pathname);
  e.respondWith(fresh ? fetch(e.request, { cache: 'no-store' }).catch(() => fetch(e.request)) : fetch(e.request));
});
