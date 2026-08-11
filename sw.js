/* つくる手帖 ── キャッシュはオリジン単位で共有されます。
   github.io は全リポジトリが同じオリジンなので、
   自分の名前空間（TT_NS）のものだけを消します。 */
/* sw.js — オフライン用 Service Worker
 * アプリシェル + ライブラリをキャッシュ。バージョンを上げると更新が反映される。 */
const TT_NS = 'tt:CRUD4Forguncy:';
const TT_OLD = 'crud-v4';   /* 旧名。次の更新のときに消して構いません */
const CACHE = TT_NS + 'v4';
const ASSETS = [
  './',
  './index.html',
  './app.js',
  './crud-core.js',
  './worker-core.js',
  './manifest.webmanifest',
  './vendor/jszip.min.js',
  './vendor/exceljs.min.js',
  './icons/icon-192.png',
  './icons/icon-512.png',
  './icons/icon-maskable-512.png'
];

self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(ASSETS)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', (e) => {
  e.waitUntil(
    caches.keys().then((keys) => Promise.all(keys.filter(k => (k.startsWith(TT_NS) || k === TT_OLD) && k !== CACHE).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

// キャッシュ優先（オフラインで確実に動く）。無ければネットワーク。
self.addEventListener('fetch', (e) => {
  if (e.request.method !== 'GET') return;
  e.respondWith(
    caches.match(e.request).then((hit) => hit || fetch(e.request).then((res) => {
      const copy = res.clone();
      caches.open(CACHE).then((c) => c.put(e.request, copy)).catch(() => {});
      return res;
    }).catch(() => caches.match('./index.html')))
  );
});
