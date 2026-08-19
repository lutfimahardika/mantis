/**
 * Mantis - PWA conversion reverted.
 *
 * This version replaces the old caching service worker on any device that
 * already registered it (from commit e283763), purely so it can clean up
 * after itself: it wipes the app-shell cache it created and unregisters,
 * then reloads any open tab so it goes back to being a normal, un-cached
 * page. index.html still has to keep calling navigator.serviceWorker.register('sw.js')
 * for this file to ever reach those devices - once this has had time to
 * propagate, both this file and that registration call can be deleted for
 * good (see index.html's PWA section).
 */
self.addEventListener('install', function(event){
  self.skipWaiting();
});

self.addEventListener('activate', function(event){
  event.waitUntil(
    caches.keys().then(function(keys){
      return Promise.all(keys.map(function(key){ return caches.delete(key); }));
    }).then(function(){
      return self.registration.unregister();
    }).then(function(){
      return self.clients.matchAll({ type: 'window' });
    }).then(function(clients){
      clients.forEach(function(client){ client.navigate(client.url); });
    })
  );
});
