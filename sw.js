/* FilaCost 3D列印成本系統 — Service Worker(離線快取)
 * 更新網站檔案後,把 VERSION 改成新值,使用者下次開啟就會換成新版快取。
 */
const VERSION = 'filacost-v6';
const SHELL_CACHE = VERSION + '-shell';
const RUNTIME_CACHE = VERSION + '-runtime';

const APP_SHELL = [
  './',
  './index.html',
  './support.js',
  './manifest.webmanifest',
  './icons/icon-192.png',
  './icons/icon-512.png',
  './icons/icon-maskable-512.png',
  './icons/apple-touch-icon.png',
];

// 字型與 3D 預覽元件:可快取供離線使用
const RUNTIME_HOSTS = ['fonts.googleapis.com', 'fonts.gstatic.com', 'ajax.googleapis.com'];
// Google Sheets(Apps Script)與 3D 模型資產:一律走網路,避免資料不同步
const NEVER_CACHE_HOSTS = ['script.google.com', 'script.googleusercontent.com', 'res.cloudinary.com'];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(SHELL_CACHE)
      .then((cache) => cache.addAll(APP_SHELL))
      .then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys
        .filter((k) => k.startsWith('filacost-') && !k.startsWith(VERSION + '-'))
        .map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (event) => {
  const req = event.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  if (NEVER_CACHE_HOSTS.indexOf(url.hostname) >= 0) return;

  if (url.origin === self.location.origin) {
    // 頁面本身(含 support.js 執行時重抓的 location.href):網路優先,離線時用快取
    const isPage = req.mode === 'navigate' || url.pathname.endsWith('/') || url.pathname.endsWith('.html');
    event.respondWith(isPage ? networkFirst(req) : staleWhileRevalidate(req, SHELL_CACHE));
    return;
  }
  if (RUNTIME_HOSTS.indexOf(url.hostname) >= 0) {
    event.respondWith(staleWhileRevalidate(req, RUNTIME_CACHE));
  }
});

async function networkFirst(req) {
  const cache = await caches.open(SHELL_CACHE);
  try {
    const res = await fetch(req);
    if (res && res.ok) cache.put(req, res.clone());
    return res;
  } catch (err) {
    return (await cache.match(req, { ignoreSearch: true }))
      || (await cache.match('./index.html'))
      || Response.error();
  }
}

async function staleWhileRevalidate(req, cacheName) {
  const cache = await caches.open(cacheName);
  const cached = await cache.match(req, { ignoreSearch: true });
  const network = fetch(req)
    .then((res) => {
      if (res && (res.ok || res.type === 'opaque')) cache.put(req, res.clone());
      return res;
    })
    .catch(() => cached);
  return cached || network;
}
