/* Offline cache for the map viewer.
 *
 * The map is one big data file and fifteen hundred pictures, all of which sit
 * on the server unchanged - a picture is never edited under the same name, it
 * is replaced by one with a new hash upstream. So a returning visitor should
 * not pay for them twice: icons are answered from the cache and fetched once,
 * the data file is shown from the cache while a fresh copy is pulled in the
 * background, and the page shell itself is fetched first so a new deployment
 * is picked up immediately.
 *
 * Bump CACHE when the way things are stored changes; the old cache is then
 * dropped on activate.
 */
var CACHE = 'entropia-maps-v1';

/* The shell is small and is what makes the page work offline at all; the data
 * file is half a megabyte and is worth having before the first pan. */
var SHELL = ['./', 'src/app.js', 'src/data.js'];

self.addEventListener('install', function (e) {
  e.waitUntil(
    caches.open(CACHE).then(function (c) {
      return Promise.all(SHELL.map(function (u) {
        return c.add(new Request(u, { cache: 'reload' })).catch(function () {});
      }));
    }).then(function () { return self.skipWaiting(); })
  );
});

self.addEventListener('activate', function (e) {
  e.waitUntil(
    caches.keys().then(function (keys) {
      return Promise.all(keys.map(function (k) {
        if (k !== CACHE) { return caches.delete(k); }
      }));
    }).then(function () { return self.clients.claim(); })
  );
});

function isIcon(path) { return path.indexOf('/icons/') !== -1; }
function isData(path) { return /\/src\/data\.js$/.test(path); }

/* A picture: the cached copy is the answer, and the network is only asked when
 * there is none.  Stored as it arrives so the next view of the same world is
 * instant. */
function fromCacheFirst(req) {
  return caches.match(req).then(function (hit) {
    if (hit) { return hit; }
    return fetch(req).then(function (res) {
      if (res && res.ok) {
        var copy = res.clone();
        caches.open(CACHE).then(function (c) { c.put(req, copy); });
      }
      return res;
    });
  });
}

/* The data file: hand over what is stored and refresh it behind the page, so a
 * new set of worlds lands on the next load without a wait. */
function fromCacheThenNetwork(req) {
  return caches.open(CACHE).then(function (c) {
    return c.match(req).then(function (hit) {
      var net = fetch(req).then(function (res) {
        if (res && res.ok) { c.put(req, res.clone()); }
        return res;
      }).catch(function () { return hit; });
      return hit || net;
    });
  });
}

/* The shell: the network decides, because a new deployment has to win, and
 * the cache is only the answer when the network cannot be reached. */
function fromNetworkThenCache(req) {
  return fetch(req).then(function (res) {
    if (res && res.ok) {
      var copy = res.clone();
      caches.open(CACHE).then(function (c) { c.put(req, copy); });
    }
    return res;
  }).catch(function () { return caches.match(req); });
}

self.addEventListener('fetch', function (e) {
  var req = e.request;
  if (req.method !== 'GET') { return; }
  var url;
  try { url = new URL(req.url); } catch (err) { return; }
  if (url.origin !== self.location.origin) { return; }
  if (isIcon(url.pathname)) { e.respondWith(fromCacheFirst(req)); return; }
  if (isData(url.pathname)) { e.respondWith(fromCacheThenNetwork(req)); return; }
  e.respondWith(fromNetworkThenCache(req));
});
