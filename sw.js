/**
 * 戰勝21點 · 離線快取
 *
 * 賭場裡常常收不到訊號，練習不該因此連開都開不起來。
 *
 * 策略分兩種：
 * - 頁面本身「先連網、連不到才用快取」—— 這樣線上的人永遠拿到最新版，
 *   推 main 之後不會有人卡在舊版；離線的人則照樣打得開。
 * - 圖示等不會變的檔案直接吃快取。
 *
 * 後端（Apps Script）完全不攔截：讓它自己失敗，前端本來就會把記錄排進佇列，
 * 等有網路再補傳。攔下來反而會把失敗包裝成看起來成功。
 */
var CACHE = 'bj21-v2';   // 換了快取清單就要換版號，否則舊的快取不會更新
var PAGE = './index.html';
var ASSETS = [
  './', PAGE, './manifest.json', './favicon.svg', './favicon.png',
  './icon-192.png', './icon-512.png', './apple-touch-icon.png'
];

self.addEventListener('install', function (e) {
  e.waitUntil(
    caches.open(CACHE)
      .then(function (c) { return c.addAll(ASSETS); })
      .then(function () { return self.skipWaiting(); })
  );
});

self.addEventListener('activate', function (e) {
  e.waitUntil(
    caches.keys().then(function (keys) {
      return Promise.all(keys.map(function (k) {
        return k === CACHE ? null : caches.delete(k);
      }));
    }).then(function () { return self.clients.claim(); })
  );
});

self.addEventListener('fetch', function (e) {
  var req = e.request;
  if (req.method !== 'GET') return;

  var url = new URL(req.url);
  if (url.origin !== location.origin) return;   // 後端、外部資源一律不碰

  var isPage = req.mode === 'navigate' ||
    url.pathname.endsWith('/') || url.pathname.endsWith('index.html');

  if (isPage) {
    e.respondWith(
      fetch(req).then(function (res) {
        var copy = res.clone();
        caches.open(CACHE).then(function (c) { c.put(PAGE, copy); });
        return res;
      }).catch(function () {
        return caches.match(PAGE).then(function (m) { return m || caches.match('./'); });
      })
    );
    return;
  }

  e.respondWith(
    caches.match(req).then(function (m) {
      if (m) return m;
      return fetch(req).then(function (res) {
        if (res && res.ok) {
          var copy = res.clone();
          caches.open(CACHE).then(function (c) { c.put(req, copy); });
        }
        return res;
      });
    })
  );
});
