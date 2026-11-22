/* Service worker: halaman selalu ambil versi terbaru dari internet, cadangan offline dari cache. */
var C='approval-tf-20261122180109';
var CORE=['./','index.html','config.js','manifest.webmanifest','favicon.svg','icons/icon-192.png','icons/icon-512.png','icons/apple-touch-icon.png'];
self.addEventListener('install',function(e){ self.skipWaiting(); e.waitUntil(caches.open(C).then(function(c){ return c.addAll(CORE) }).catch(function(){})) });
self.addEventListener('activate',function(e){ e.waitUntil(caches.keys().then(function(k){ return Promise.all(k.filter(function(x){ return x!==C }).map(function(x){ return caches.delete(x) })) }).then(function(){ return self.clients.claim() })) });
self.addEventListener('fetch',function(e){
  var u=new URL(e.request.url);
  if(e.request.method!=='GET'||u.origin!==location.origin) return;
  e.respondWith(fetch(e.request).then(function(r){ var cp=r.clone(); caches.open(C).then(function(c){ c.put(e.request,cp) }); return r })
    .catch(function(){ return caches.match(e.request).then(function(m){ return m||caches.match('index.html') }) }));
});
