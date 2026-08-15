const CACHE_NAME = "shahada-cache-v4";
const ASSETS = [
  "./",
  "./index.html",
  "./manifest.json",
  "./assets/style.css",
  "./assets/app.js",
  "./assets/data.js",
  "./assets/xlsx.full.min.js",
  "./assets/icon-192.png",
  "./assets/icon-512.png",
  "./assets/apple-touch-icon.png"
];

self.addEventListener("install", (event)=>{
  event.waitUntil(
    caches.open(CACHE_NAME).then(cache => cache.addAll(ASSETS)).catch(()=>{})
  );
  self.skipWaiting();
});

self.addEventListener("activate", (event)=>{
  event.waitUntil(
    caches.keys().then(keys => Promise.all(
      keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k))
    ))
  );
  self.clients.claim();
});

self.addEventListener("fetch", (event)=>{
  if(event.request.method !== "GET") return;
  event.respondWith(
    caches.match(event.request).then(cached=>{
      const network = fetch(event.request).then(res=>{
        if(res && res.status === 200 && res.type === "basic"){
          const clone = res.clone();
          caches.open(CACHE_NAME).then(cache => cache.put(event.request, clone));
        }
        return res;
      }).catch(()=> cached);
      return cached || network;
    })
  );
});
