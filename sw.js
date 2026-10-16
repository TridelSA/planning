// Service worker minimal : requis par certains navigateurs pour proposer
// l'installation. Il ne met rien en cache, les données restent toujours à jour.
self.addEventListener('install',()=>self.skipWaiting());
self.addEventListener('activate',e=>e.waitUntil(self.clients.claim()));
self.addEventListener('fetch',()=>{});
