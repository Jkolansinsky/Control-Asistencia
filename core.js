// Núcleo compartido: configuración, API, almacenamiento local, cola offline, rostro, GPS.
export const CFG={
  API:'https://script.google.com/macros/s/AKfycby9PepTsDIn30G09B5bgPgStGskrej0pLuXFc8B4Tqiw_5T0gBYMjAslYubaY-aqRDylw/exec',
  MODELS:'model/', // mejor: copia /model a tu repo y usa 'models/'
  UMBRAL:0.5, MARGEN:0.04, ACC_MAX:150
};
// ---- IndexedDB mínima ----
const db=new Promise((ok,ko)=>{const r=indexedDB.open('ac11',1);r.onupgradeneeded=()=>{r.result.createObjectStore('kv');r.result.createObjectStore('q',{keyPath:'opId'});};r.onsuccess=()=>ok(r.result);r.onerror=()=>ko(r.error);});
const tx=async(s,m,f)=>{const d=await db;return new Promise((ok,ko)=>{const t=d.transaction(s,m);const q=f(t.objectStore(s));t.oncomplete=()=>ok(q&&q.result);t.onerror=()=>ko(t.error);});};
export const kv={get:k=>tx('kv','readonly',s=>s.get(k)),set:(k,v)=>tx('kv','readwrite',s=>s.put(v,k))};
const queue={add:o=>tx('q','readwrite',s=>s.put(o)),del:i=>tx('q','readwrite',s=>s.delete(i)),all:()=>tx('q','readonly',s=>s.getAll())};
export const hoyKey=()=>new Date().toLocaleDateString('sv');
export async function deviceId(){let d=localStorage.getItem('ac_dev')||await kv.get('dev');if(!d)d=crypto.randomUUID();localStorage.setItem('ac_dev',d);await kv.set('dev',d);return d;}
// ---- API (text/plain evita preflight CORS en Apps Script) ----
async function api(action,data={}){
  const c=new AbortController(),t=setTimeout(()=>c.abort(),12000);
  try{const r=await fetch(CFG.API,{method:'POST',headers:{'Content-Type':'text/plain;charset=utf-8'},body:JSON.stringify({action,...data}),signal:c.signal});return await r.json();}
  catch{return {ok:false,code:'NET',msg:'Sin conexión'};}
  finally{clearTimeout(t);}
}
let sess=null;
export async function login(id){const r=await api('activar',{idLicencia:id,deviceId:await deviceId()});if(r.ok){sess=r;localStorage.setItem('ac_lic',id);await kv.set('sess',r);}return r;}
export async function call(action,data={}){
  if(!sess)sess=await kv.get('sess');
  if(!sess||sess.exp-Date.now()<36e5){const id=localStorage.getItem('ac_lic');if(id){const r=await login(id);if(!r.ok&&r.code!=='NET')return r;}}
  if(!sess||sess.exp<Date.now())return {ok:false,code:'NET',msg:'Sin sesión'};
  let r=await api(action,{token:sess.token,...data});
  if(r.code==='AUTH'||r.code==='AUTH_EXPIRED'){const id=localStorage.getItem('ac_lic');const l=id&&await login(id);if(l&&l.ok)r=await api(action,{token:sess.token,...data});}
  return r;
}
export async function sync(){const c=await kv.get('data');const r=await call('bootstrap',{version:c&&c.version});if(r.ok&&!r.unchanged){await kv.set('data',r);return r;}return c;}
// ---- Cola offline: se guarda ANTES de enviar; el opId evita duplicados ----
const DEF=['YA_REGISTRADO','NO_EMP','BAJA','BAD','GPS'];
export async function enviar(op){await queue.add(op);const r=await call('registrar',op);if(r.ok||DEF.includes(r.code))await queue.del(op.opId);return r;}
export async function vaciarCola(){for(const op of await queue.all()){const r=await call('registrar',{...op,origen:'diferido'});if(r.ok||DEF.includes(r.code))await queue.del(op.opId);else if(r.code==='NET')break;}}
export const pendientes=async()=>(await queue.all()).length;
// ---- Rostro (solo descriptores numéricos, nunca fotos) ----
export const loadFace=()=>Promise.all([faceapi.nets.tinyFaceDetector,faceapi.nets.faceLandmark68Net,faceapi.nets.faceRecognitionNet].map(n=>n.loadFromUri(CFG.MODELS)));
export const describe=el=>faceapi.detectSingleFace(el,new faceapi.TinyFaceDetectorOptions({inputSize:320,scoreThreshold:0.5})).withFaceLandmarks().withFaceDescriptor();
export const dist=(a,b)=>{let s=0;for(let i=0;i<128;i++){const d=a[i]-b[i];s+=d*d;}return Math.sqrt(s);};
export function match(d,emps){
  let best=null,bd=9,sd=9;
  for(const e of emps){let m=9;for(const x of e.d)m=Math.min(m,dist(d,x));if(m<bd){sd=bd;bd=m;best=e;}else if(m<sd)sd=m;}
  return best&&bd<=CFG.UMBRAL&&sd-bd>=CFG.MARGEN?{emp:best,dist:bd}:null;
}
// ---- GPS ----
export const hav=(a,b,c,d)=>{const R=6371000,r=Math.PI/180,x=(c-a)*r,y=(d-b)*r,h=Math.sin(x/2)**2+Math.cos(a*r)*Math.cos(c*r)*Math.sin(y/2)**2;return 2*R*Math.asin(Math.sqrt(h));};
export function watchGPS(cb){if(!navigator.geolocation)return;navigator.geolocation.watchPosition(p=>cb({lat:p.coords.latitude,lng:p.coords.longitude,acc:p.coords.accuracy,t:Date.now()}),()=>{},{enableHighAccuracy:true,maximumAge:15000,timeout:20000});}
