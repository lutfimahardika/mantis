/**
 * Mantis - service worker for the PWA app shell.
 *
 * Only caches the static shell (this file's own origin: index.html, manifest,
 * icons). Every request to a different origin - Google Sheets gviz reads,
 * the Apps Script write/upload endpoints, Google Fonts, the Tailwind CDN -
 * is left completely untouched and goes straight to the network, since that
 * data has to stay live and is never something to serve stale from cache.
 */
const CACHE_NAME = 'mantis-shell-v1';
const APP_SHELL = [
  './',
  './index.html',
  './manifest.json',
  './favicon.svg',
  './icons/icon-192.png',
  './icons/icon-512.png'
];

self.addEventListener('install', function(event){
  event.waitUntil(
    caches.open(CACHE_NAME).then(function(cache){ return cache.addAll(APP_SHELL); })
  );
  self.skipWaiting();
});

self.addEventListener('activate', function(event){
  event.waitUntil(
    caches.keys().then(function(keys){
      return Promise.all(keys.filter(function(key){ return key !== CACHE_NAME; }).map(function(key){ return caches.delete(key); }));
    })
  );
  self.clients.claim();
});

/* Network-first for the app shell: whoever's online always gets the latest
   deploy immediately (no "stuck on an old cached version" surprise), and
   offline visits fall back to whatever shell was cached on the last visit. */
self.addEventListener('fetch', function(event){
  const req = event.request;
  if(req.method !== 'GET') return;

  const url = new URL(req.url);
  if(url.origin !== self.location.origin) return;

  event.respondWith(
    fetch(req).then(function(res){
      const resClone = res.clone();
      caches.open(CACHE_NAME).then(function(cache){ cache.put(req, resClone); });
      return res;
    }).catch(function(){
      return caches.match(req).then(function(cached){ return cached || caches.match('./index.html'); });
    })
  );
});
