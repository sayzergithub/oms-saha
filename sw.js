/* OMS Saha service worker: offline, cache-first, versioned.
   VERSION is replaced by tools/publish_field.ps1 with a date-time stamp; every publish = new cache, old ones are dropped.
   All paths are relative to this file, so the page works under any sub-path (/repo-name/ on GitHub Pages, /saha/ on the PC). */
const VERSION = '20261023-142919';
const CACHE = 'oms-saha-' + VERSION;
const ASSETS = ['./', './index.html', './app.js', './envelope.js', './labelparse.js', './zxing.min.js',
  './manifest.webmanifest', './icon.svg', './icon-192.png', './icon-512.png',
  /* V6 modules: camera settings, stock catalog, direct send (MSAL v5 + redirect bridge) */
  './camsettings.js', './catalog.js', './sendmail.js', './msal-browser.min.js',
  /* V7: cloud folder send (share sheet / File System Access / download) */
  './cloudsend.js',
  /* V8: direct OneDrive (MSAL v5 above), stock list tab, serial scan logic (copied from Web/ like camsettings.js) */
  './onedrive.js', './stockview.js', './serialscan.js',
  /* V10: direct send over the local network */
  './lansend.js'];

self.addEventListener('install', function (e) {
  e.waitUntil(caches.open(CACHE).then(function (c) {
    /* one by one: a missing optional file must not abort the install; bypass the HTTP cache so a new VERSION gets new files */
    return Promise.all(ASSETS.map(function (u) {
      return fetch(new Request(u, { cache: 'reload' })).then(function (r) { if (r.ok) return c.put(u, r); }).catch(function () { });
    }));
  }).then(function () {
    /* first install: take over at once; updates wait for the page's "Şimdi yenile" (no reload under the user's hands) */
    return self.registration.active ? null : self.skipWaiting();
  }));
});

self.addEventListener('activate', function (e) {
  e.waitUntil(caches.keys().then(function (ks) {
    return Promise.all(ks.map(function (k) { if (k.indexOf('oms-saha-') === 0 && k !== CACHE) return caches.delete(k); }));
  }).then(function () { return self.clients.claim(); }));
});

self.addEventListener('message', function (e) {
  if (e.data && e.data.t === 'skip') self.skipWaiting();
});

self.addEventListener('fetch', function (e) {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  const scope = new URL(self.registration.scope);
  if (url.origin !== scope.origin || url.pathname.indexOf(scope.pathname) !== 0) return;   /* not ours (e.g. the LAN page next to /saha/) */
  if (url.pathname.indexOf('/api/') >= 0) return;   /* PC API under /saha/api/ (catalog): always network, never cached */
  if (req.mode === 'navigate') {
    /* any page load inside the scope -> the app shell; the #OMSP1 hash never reaches the network anyway */
    e.respondWith(caches.open(CACHE).then(function (c) {
      return c.match('./index.html').then(function (hit) {
        return hit || c.match('./').then(function (h2) { return h2 || fetch(req); });
      });
    }).catch(function () { return fetch(req); }));
    return;
  }
  e.respondWith(caches.open(CACHE).then(function (c) {
    return c.match(req, { ignoreSearch: true }).then(function (hit) {
      if (hit) return hit;
      return fetch(req).then(function (r) {
        if (r && r.ok && r.type === 'basic') c.put(req, r.clone());
        return r;
      });
    });
  }));
});
