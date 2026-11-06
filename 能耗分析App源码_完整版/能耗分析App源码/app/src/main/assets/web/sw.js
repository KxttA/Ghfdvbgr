const CACHE = 'energy-m-v2';
const CORE = ['/', '/app.css', '/app.js', '/manifest.webmanifest', '/icon-192.png', '/icon-512.png'];
self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(CORE)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const r = e.request;
  if (r.method !== 'GET' || !r.url.startsWith(self.location.origin)) return;
  if (r.url.includes('/export') || r.url.includes('/download') || r.url.includes('/backup')) return;
  if (new URL(r.url).pathname.startsWith('/api/')) {
    // 接口：网络优先，离线时回缓存（可看上次数据）
    e.respondWith(fetch(r).then(res => {
      const cl = res.clone(); caches.open(CACHE).then(c => c.put(r, cl)); return res;
    }).catch(() => caches.match(r)));
    return;
  }
  e.respondWith(caches.match(r).then(hit => {
    const net = fetch(r).then(res => { const cl = res.clone(); caches.open(CACHE).then(c => c.put(r, cl)); return res; }).catch(() => hit);
    return hit || net;
  }));
});
