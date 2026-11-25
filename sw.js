// Sube V cuando cambien index.html / core.js / styles.css. PERM guarda la librería y los modelos (pesados) y NO se borra al actualizar.
const V='ac-v11-7',PERM='ac-assets-1';
const SHELL=['./','index.html','admin.html','reporte.html','aviso.html','styles.css','core.js','manifest.webmanifest','icon-192.png','icon-512.png'];
const HEAVY=['face-api.js'].concat(['tiny_face_detector_model','face_landmark_68_model','face_recognition_model'].flatMap(m=>['models/'+m+'-weights_manifest.json','models/'+m+'.bin']));
self.addEventListener('install',e=>{e.waitUntil((async()=>{
  const c=await caches.open(V); await Promise.all(SHELL.map(u=>c.add(u).catch(()=>{})));
  const p=await caches.open(PERM); await Promise.all(HEAVY.map(async u=>{if(!await p.match(u)) await p.add(u).catch(()=>{});}));
  self.skipWaiting();
})());});
self.addEventListener('activate',e=>{e.waitUntil(caches.keys().then(k=>Promise.all(k.filter(x=>x!==V&&x!==PERM).map(x=>caches.delete(x)))).then(()=>self.clients.claim()));});
self.addEventListener('fetch',e=>{
  const u=new URL(e.request.url); if(e.request.method!=='GET'||u.origin!==location.origin)return;
  const heavy=HEAVY.some(h=>u.pathname.endsWith('/'+h));
  e.respondWith((async()=>{
    const c=await caches.open(heavy?PERM:V), hit=await c.match(e.request,{ignoreSearch:true});
    if(heavy&&hit) return hit; // librería y modelos: siempre del caché
    const net=fetch(e.request).then(r=>{if(r.ok)c.put(e.request,r.clone());return r;}).catch(()=>null);
    return hit||await net||(e.request.mode==='navigate'?await caches.match('index.html'):Response.error());
  })());
});
