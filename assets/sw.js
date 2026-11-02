/* Toolkit service worker: pages network-first (offline copy as fallback), site assets and download libraries cache-first.
   build.py fills in the version and the precache list. */
var V = 'tk-__VERSION__', PAGES = 'tk-pages', LIBS = 'tk-libs';
var PRE = __PRECACHE__;
self.addEventListener('install', function (e) {
  e.waitUntil(caches.open(V).then(function (c) { return c.addAll(PRE); }).then(function () { return self.skipWaiting(); }));
});
self.addEventListener('activate', function (e) {
  e.waitUntil(caches.keys().then(function (ks) {
    return Promise.all(ks.filter(function (k) { return k !== V && k !== PAGES && k !== LIBS; }).map(function (k) { return caches.delete(k); }));
  }).then(function () { return self.clients.claim(); }));
});
self.addEventListener('fetch', function (e) {
  var r = e.request, u = new URL(r.url);
  if (r.method !== 'GET') return;
  if (r.mode === 'navigate') {
    e.respondWith(fetch(r).then(function (res) {
      var copy = res.clone(); caches.open(PAGES).then(function (c) { c.put(r, copy); }); return res;
    }).catch(function () { return caches.match(r).then(function (m) { return m || caches.match('/offline.html'); }); }));
    return;
  }
  var lib = /cdnjs\.cloudflare\.com|cdn\.jsdelivr\.net|fonts\.(googleapis|gstatic)\.com|share\.groupchaton\.com\/promos|googleusercontent|assets\.(infinity\.)?groupchaton\.com\/assets\/brand/.test(u.href);
  if (u.origin === location.origin && /^\/(assets|brand|data)\//.test(u.pathname) || lib) {
    var name = lib ? LIBS : V;
    e.respondWith(caches.match(r).then(function (m) {
      var net = fetch(r).then(function (res) { if (res.ok || res.type === 'opaque') { var c2 = res.clone(); caches.open(name).then(function (c) { c.put(r, c2); }); } return res; });
      return m ? (u.href.indexOf('promos.json') >= 0 ? net.catch(function () { return m; }) : m) : net;
    }));
  }
});
