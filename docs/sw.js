// サイト本体をキャッシュして、一度開いた端末なら圏外でも表示できるようにする。
// キャッシュがあれば即それを返し、裏で最新版を取ってキャッシュを差し替える（次回表示から反映）。
// データ（GAS）は別オリジンなので触らない。データのオフライン表示は index.html 側の localStorage で行う。
const CACHE = 'gl-v1';
const FILES = ['./', 'index.html', 'manifest.json'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(FILES)));
  self.skipWaiting();
});

self.addEventListener('activate', e => {
  e.waitUntil(caches.keys()
    .then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
});

self.addEventListener('fetch', e => {
  const r = e.request;
  if (r.method !== 'GET' || new URL(r.url).origin !== location.origin) return;
  e.respondWith(caches.open(CACHE).then(async c => {
    const hit = await c.match(r, { ignoreSearch: true });
    const net = fetch(r.url, { cache: 'no-cache' }).then(res => {
      if (res.ok) c.put(r.url, res.clone());
      return res;
    });
    if (!hit) return net;
    e.waitUntil(net.catch(() => {}));
    return hit;
  }));
});
