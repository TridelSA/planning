// Service worker : permet l'installation et le fonctionnement hors ligne.
// Réseau d'abord (les données et le code restent toujours à jour) ;
// la copie en cache ne sert que si le réseau est indisponible.
const CACHE='planning-shell-v1';
const SHELL=['./','index.html','style.css','app.js','manifest.json','icon-192-v5.png'];
self.addEventListener('install',e=>{
  e.waitUntil(caches.open(CACHE).then(c=>c.addAll(SHELL)).catch(()=>{}).then(()=>self.skipWaiting()));
});
self.addEventListener('activate',e=>{
  e.waitUntil(caches.keys().then(ks=>Promise.all(ks.filter(k=>k!==CACHE).map(k=>caches.delete(k)))).then(()=>self.clients.claim()));
});
self.addEventListener('fetch',e=>{
  const r=e.request;
  if(r.method!=='GET')return;
  const u=new URL(r.url);
  if(u.origin!==location.origin)return;
  e.respondWith(
    fetch(r).then(res=>{
      if(res&&res.ok){ const copy=res.clone(); caches.open(CACHE).then(c=>c.put(u.pathname+(u.search.startsWith('?v=')?'':u.search),copy)).catch(()=>{}); }
      return res;
    }).catch(()=>caches.match(u.pathname+(u.search.startsWith('?v=')?'':u.search)).then(m=>m||caches.match(r,{ignoreSearch:true})))
  );
});
