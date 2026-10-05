// Service Worker v2：缓存优先（stale-while-revalidate）——二次打开秒开，后台自动更新
const CACHE = 'wps-station-v2';
const CORE = ['./index.html', './app.css', './app.js', './manifest.json', './icon-192.png', './icon-512.png', './icon-180.png'];
const CDN = 'https://cdn.jsdelivr.net/npm/';

self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(CORE)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', (e) => {
  e.waitUntil(caches.keys().then((ks) => Promise.all(ks.filter((k) => k !== CACHE).map((k) => caches.delete(k)))).then(() => self.clients.claim()));
});

self.addEventListener('fetch', (e) => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);

  // 同源静态资源 + CDN 库：缓存优先，后台更新（离线/二次打开秒开）
  if (url.origin === self.location.origin || url.href.startsWith(CDN)) {
    e.respondWith(
      caches.match(req).then((hit) => {
        // 后台刷新缓存
        const refresh = fetch(req).then((res) => {
          if (res && (res.ok || res.type === 'opaque')) {
            const copy = res.clone();
            caches.open(CACHE).then((c) => c.put(req, copy)).catch(() => {});
          }
          return res;
        }).catch(() => null);
        if (hit) {
          // 返回缓存，同时触发后台更新
          e.waitUntil(refresh.then(() => {}));
          return hit;
        }
        return refresh.then((res) => res || caches.match('./index.html'));
      })
    );
    return;
  }

  // 其他请求（高德 API 等）：网络优先，失败回退缓存
  e.respondWith(
    fetch(req).then((res) => {
      const copy = res.clone();
      caches.open(CACHE).then((c) => c.put(req, copy)).catch(() => {});
      return res;
    }).catch(() => caches.match(req).then((m) => m || caches.match('./index.html')))
  );
});
