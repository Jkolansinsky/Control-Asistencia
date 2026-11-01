// ============ ASISTENCIA v11 — Apps Script (sin fotos en Sheets) ============
// Implementar como: Aplicación web · Ejecutar como: Yo · Acceso: Cualquier persona
const CFG={
  CLIENT_SS:'1HKY4xozQB_hhgYukdbMORpKSbX8HuOjjzOx1icgVVCI',
  LIC_SS:'1I0OXIUGXxUE3K7ibVjel4hW8-efrA13ldFjVtdCf6og',
  HORARIOS:['Hora Entrada','Salida Comida','Entrada Comida','Hora Salida'],
  RADIO_M:200, ACC_MAX:150, BLOQUEAR_FUERA:false, RECIENTES:1500, TOKEN_H:24, TURNO_H:16 // TURNO_H: horas máximas que dura un turno abierto (tus turnos son de 8 a 12 h)
};
class ApiError extends Error{constructor(c,m,x){super(m);this.code=c;this.extra=x||{};}}
const ss_=()=>SpreadsheetApp.openById(CFG.CLIENT_SS), tz_=()=>Session.getScriptTimeZone();
const out_=o=>ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
const hh_=v=>v instanceof Date?Utilities.formatDate(v,tz_(),'HH:mm:ss'):String(v||'');

function doGet(){return out_({ok:true,app:'asistencia',v:'11'});}
function doPost(e){
  try{
    const r=JSON.parse(e.postData.contents), a=r.action;
    if(a==='activar') return out_(Object.assign({ok:true},activar_(r)));
    // Reportes para gerentes/administradores desde su celular: se autentican con usuario y contraseña (sin licencia de dispositivo)
    if(a==='reporteLogin') return out_(Object.assign({ok:true},reporteLogin_(r)));
    if(a==='reportePdf') return out_(Object.assign({ok:true},reportePdf_(r,tok_(r.token,'rep'))));
    const lic=tok_(r.token,'dev'); let res;
    switch(a){
      case 'bootstrap': res=bootstrap_(r); break;
      case 'registrar': res=registrar_(r,lic); break;
      case 'adminLogin': res=adminLogin_(r,lic); break;
      case 'guardarEmpleado': case 'legacyLote': {
        const ad=tok_(r.adminToken,'adm'); if(ad.lic!==lic.lic) throw new ApiError('AUTH','Sesión no válida');
        res=a==='guardarEmpleado'?guardar_(r):legacy_(r); break; }
      default: throw new ApiError('BAD','Acción no válida');
    }
    return out_(Object.assign({ok:true},res));
  }catch(err){
    if(err instanceof ApiError) return out_(Object.assign({ok:false,code:err.code,msg:err.message},err.extra));
    console.error(err); return out_({ok:false,code:'SERVER',msg:'Error interno'});
  }
}

// ---------- Tokens firmados (HMAC) ----------
function sec_(){const p=PropertiesService.getScriptProperties();let s=p.getProperty('SECRET');if(!s){s=Utilities.getUuid()+Utilities.getUuid();p.setProperty('SECRET',s);}return s;}
const sig_=b=>Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(b,sec_()));
function mk_(p){const b=Utilities.base64EncodeWebSafe(JSON.stringify(p));return b+'.'+sig_(b);}
function tok_(t,tipo){
  const [b,s]=String(t||'').split('.');
  if(!b||!s||sig_(b)!==s) throw new ApiError('AUTH','Sesión no válida');
  const p=JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(b)).getDataAsString());
  if(p.t!==tipo) throw new ApiError('AUTH','Sesión no válida');
  if(p.exp<Date.now()) throw new ApiError('AUTH_EXPIRED','Sesión expirada');
  return p;
}

// ---------- Licencia: vencimiento + dispositivo (UUID se enlaza en el primer uso) ----------
function activar_(r){
  const id=String(r.idLicencia||'').trim().toUpperCase(), dev=String(r.deviceId||'').trim();
  if(!id||dev.length<16) throw new ApiError('BAD','Faltan datos');
  const lock=LockService.getScriptLock(); lock.waitLock(15000);
  try{
    const sh=SpreadsheetApp.openById(CFG.LIC_SS).getSheetByName('Licencias'), d=sh.getDataRange().getValues();
    for(let i=1;i<d.length;i++){
      if(String(d[i][0]).trim().toUpperCase()!==id) continue;
      if(String(d[i][5]).trim()!=='Activa') throw new ApiError('LIC','Licencia inactiva.');
      const ex=d[i][4], hoy=new Date(); hoy.setHours(0,0,0,0);
      if(ex instanceof Date && ex<hoy) throw new ApiError('LIC','Licencia vencida.');
      const bound=String(d[i][6]||'').trim();
      if(!bound) sh.getRange(i+1,7).setNumberFormat('@').setValue(dev);
      else if(bound!==dev) throw new ApiError('LIC','Dispositivo no autorizado.');
      const exp=Date.now()+CFG.TOKEN_H*36e5;
      return {token:mk_({t:'dev',lic:id,exp}),exp,cliente:d[i][1],dias:ex instanceof Date?Math.ceil((ex-hoy)/864e5):null};
    }
    throw new ApiError('LIC','Licencia no encontrada.');
  }finally{lock.releaseLock();}
}

// ---------- Datos (versionados + caché) ----------
const ver_=()=>PropertiesService.getScriptProperties().getProperty('VER')||'0';
function bump_(){PropertiesService.getScriptProperties().setProperty('VER',String(Date.now()));const c=CacheService.getScriptCache();c.remove('EMP');c.remove('SUC');}
function onEdit(e){const n=e.range.getSheet().getName(); if(n==='BD'||n==='Sucursales') bump_();}
function bootstrap_(r){
  const v=ver_(); if(r.version&&r.version===v) return {unchanged:true,version:v};
  return {version:v,sucursales:sucursales_(),empleados:empleadosConRostro_()};
}
function sucursales_(){
  const c=CacheService.getScriptCache(), h=c.get('SUC'); if(h) return JSON.parse(h);
  const d=ss_().getSheetByName('Sucursales').getDataRange().getValues(), o=[];
  for(let i=1;i<d.length;i++){
    const n=String(d[i][0]||'').trim(); if(!n) continue;
    const m=String(d[i][1]||'').match(/-?\d+(?:\.\d+)?/g);
    if(m&&m.length>=2) o.push({nombre:n,lat:+m[0],lng:+m[1],radio:+d[i][4]||CFG.RADIO_M});
  }
  try{c.put('SUC',JSON.stringify(o),600);}catch(e){}
  return o;
}
function empleadosConRostro_(){
  const sh=ss_().getSheetByName('BD'), n=sh.getLastRow()-1; if(n<1) return [];
  const b=sh.getRange(2,1,n,7).getValues(), dd=sh.getRange(2,9,n,1).getValues(), o=[]; // no lee la col. H (fotos)
  b.forEach((x,i)=>{const id=String(x[0]).trim().toUpperCase(), d=String(dd[i][0]||'');
    if(id&&d&&String(x[5]).trim().toLowerCase()==='alta') o.push({id,nombre:String(x[4]).trim(),sucursal:String(x[1]).trim(),d});});
  return o;
}
function empMap_(){
  const c=CacheService.getScriptCache(), h=c.get('EMP'); if(h) return JSON.parse(h);
  const sh=ss_().getSheetByName('BD'), n=sh.getLastRow()-1, m={};
  if(n>0) sh.getRange(2,1,n,7).getValues().forEach(x=>{const id=String(x[0]).trim().toUpperCase(); if(id) m[id]={n:String(x[4]).trim(),s:String(x[1]).trim(),a:String(x[5]).trim().toLowerCase()==='alta'};});
  try{c.put('EMP',JSON.stringify(m),600);}catch(e){}
  return m;
}
function hav_(a,b,c,d){const R=6371000,r=Math.PI/180,x=(c-a)*r,y=(d-b)*r,h=Math.sin(x/2)**2+Math.cos(a*r)*Math.cos(c*r)*Math.sin(y/2)**2;return 2*R*Math.asin(Math.sqrt(h));}
function ts_(x){ // marca de tiempo de una fila de Registros
  if(x[1] instanceof Date) return x[1].getTime();
  try{return Utilities.parseDate(fk_(x[6])+' '+hh_(x[7]),tz_(),'yyyy-MM-dd HH:mm:ss').getTime();}catch(e){return 0;}
}
function fk_(v){
  if(v instanceof Date) return Utilities.formatDate(v,tz_(),'yyyy-MM-dd');
  const s=String(v||'').trim(), m=s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  return m?m[3]+'-'+('0'+m[2]).slice(-2)+'-'+('0'+m[1]).slice(-2):s.slice(0,10);
}

// ---------- Registrar (idempotente, con candado, sucursal calculada en servidor) ----------
function registrar_(r,lic){
  const id=String(r.id||'').trim().toUpperCase(), hor=String(r.horario||''), op=String(r.opId||'');
  if(!id||op.length<8||CFG.HORARIOS.indexOf(hor)<0) throw new ApiError('BAD','Datos incompletos');
  const emp=empMap_()[id];
  if(!emp) throw new ApiError('NO_EMP','El ID '+id+' no existe.');
  if(!emp.a) throw new ApiError('BAJA','Empleado dado de baja.');
  const ahora=new Date(); let cuando=ahora, origen='online';
  if(r.origen==='diferido'){const t=new Date(r.ts); if(!isNaN(t)&&t<=ahora&&ahora-t<6048e5){cuando=t;origen='diferido';}}
  const tieneGps=r.lat!=null&&isFinite(+r.lat)&&isFinite(+r.lng), lat=+r.lat, lng=+r.lng, acc=Math.round(+r.acc||0);
  let sg='', dist='', est='SIN_GPS';
  if(tieneGps){
    let b=null,bd=Infinity;
    sucursales_().forEach(s=>{const d=hav_(lat,lng,s.lat,s.lng); if(d<bd){bd=d;b=s;}});
    if(b){dist=Math.round(bd); if(bd<=b.radio){sg=b.nombre; est=acc>CFG.ACC_MAX?'GPS_IMPRECISO':'OK';} else est='FUERA_DE_RANGO';}
  }
  if(CFG.BLOQUEAR_FUERA&&(est==='SIN_GPS'||est==='FUERA_DE_RANGO')) throw new ApiError('GPS','Ubicación no válida para registrar.');
  const fecha=fk_(cuando), hora=Utilities.formatDate(cuando,tz_(),'HH:mm:ss');
  const lock=LockService.getScriptLock(); lock.waitLock(20000);
  try{
    const sh=ss_().getSheetByName('Registros');
    if(sh.getMaxColumns()<20) sh.insertColumnsAfter(sh.getMaxColumns(),20-sh.getMaxColumns());
    const last=sh.getLastRow(), st=Math.max(2,last-CFG.RECIENTES+1);
    const rows=last>=2?sh.getRange(st,1,last-st+1,20).getValues():[];
    // ---- Lógica por TURNO (no por fecha): un turno va de una entrada a su salida, aunque cruce la medianoche ----
    const T=cuando.getTime(), mine=[]; let rep=false;
    rows.forEach(x=>{
      if(String(x[12])===op) rep=true;
      if(String(x[0]).trim().toUpperCase()===id){const t=ts_(x); if(t&&t<=T&&T-t<CFG.TURNO_H*36e5) mine.push({t,h:String(x[3]),hr:hh_(x[7])});}
    });
    mine.sort((p,q)=>p.t-q.t);
    let i0=0; mine.forEach((m,i)=>{if(m.h==='Hora Salida') i0=i+1;}); // lo anterior a la última salida ya es otro turno
    const turno=mine.slice(i0), hoy=turno.map(m=>m.h);
    if(rep) return {repetido:true,nombre:emp.n,hoy,hora};
    const dupR=turno.find(m=>m.h===hor);
    if(dupR) throw new ApiError('YA_REGISTRADO','Ya registraste "'+hor+'" en este turno a las '+dupR.hr+'.',{hora:dupR.hr,hoy});
    const ent=turno.find(m=>m.h==='Hora Entrada');
    const jornada=hor==='Hora Entrada'?fecha:(ent?fk_(new Date(ent.t)):fecha); // la salida de las 00:01 cuenta en la jornada del día que entró
    const nota=hor!=='Hora Entrada'&&!ent?'SIN_ENTRADA':'';
    const sr=sg||emp.s;
    const f=[id,cuando,sr,hor,tieneGps?lat.toFixed(6)+', '+lng.toFixed(6):'Sin GPS',emp.n,Utilities.parseDate(fecha,tz_(),'yyyy-MM-dd'),hora,sg,dist,acc||'',est,op,lic.lic,origen,+r.confianza||'',emp.s,ahora,Utilities.parseDate(jornada,tz_(),'yyyy-MM-dd'),nota];
    sh.getRange(last+1,1,1,20)
      .setNumberFormats([['@','dd/MM/yyyy HH:mm:ss','@','@','@','@','dd/MM/yyyy','@','@','0','0','@','@','@','@','0.000','@','dd/MM/yyyy HH:mm:ss','dd/MM/yyyy','@']])
      .setValues([f]);
    SpreadsheetApp.flush();
    return {nombre:emp.n,sucursal:sr,estatus:est,hora,hoy:hor==='Hora Salida'?[]:hoy.concat(hor)};
  }finally{lock.releaseLock();}
}

// ---------- Admin: contraseñas con hash + límite de intentos ----------
function hash_(salt,p){let h=salt+':'+p;for(let i=0;i<1500;i++)h=Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,h));return h;}
function loginUsuario_(r){ // valida usuario y clave (hash); devuelve {usuario,rol,suc}
  const u=String(r.usuario||'').trim().toLowerCase(), p=String(r.password||'');
  const c=CacheService.getScriptCache(), k='RL:'+u, n=+(c.get(k)||0);
  if(n>=5) throw new ApiError('LOCK','Demasiados intentos. Espera 15 minutos.');
  const sh=ss_().getSheetByName('admin'), d=sh.getDataRange().getValues();
  for(let i=1;i<d.length;i++){
    if(String(d[i][0]).trim().toLowerCase()!==u||String(d[i][3]).trim().toLowerCase()!=='activo') continue;
    let ok=false;
    if(d[i][4]) ok=hash_(String(d[i][5]),p)===String(d[i][4]);
    else if(p&&String(d[i][1]).trim()===p){ok=true;const s=Utilities.getUuid(); sh.getRange(i+1,2).setValue(''); sh.getRange(i+1,5,1,2).setValues([[hash_(s,p),s]]);} // migra clave en texto plano a hash
    if(ok){c.remove(k);return {usuario:String(d[i][0]).trim(),rol:String(d[i][6]||'admin').trim().toLowerCase(),suc:String(d[i][7]||'').trim()};}
  }
  c.put(k,String(n+1),900); throw new ApiError('AUTH_ADMIN','Usuario o contraseña incorrectos.');
}
function adminLogin_(r,lic){
  const x=loginUsuario_(r);
  if(x.rol!=='admin') throw new ApiError('AUTH_ADMIN','Este usuario no tiene permiso de administración.');
  return {adminToken:mk_({t:'adm',u:x.usuario,lic:lic.lic,exp:Date.now()+18e5}),usuario:x.usuario};
}
function reporteLogin_(r){
  const x=loginUsuario_(r);
  if(x.rol!=='admin'&&x.rol!=='gerente') throw new ApiError('AUTH_ADMIN','Usuario sin permiso para reportes.');
  if(x.rol==='gerente'&&!x.suc) throw new ApiError('AUTH_ADMIN','Usuario sin sucursal asignada.');
  return {token:mk_({t:'rep',u:x.usuario,rol:x.rol,suc:x.suc,exp:Date.now()+8*36e5}),usuario:x.usuario,rol:x.rol,
    sucursales:x.rol==='admin'?['TODAS'].concat(sucursales_().map(s=>s.nombre)):[x.suc]};
}
// PDF idéntico a la hoja: se dibuja en un archivo temporal (no choca con otros usuarios), se exporta y se elimina
function reportePdf_(r,p){
  const c=CacheService.getScriptCache(), kl='PDF:'+p.u;
  if(c.get(kl)) throw new ApiError('WAIT','Espera unos segundos antes de generar otro reporte.');
  c.put(kl,'1',20);
  const iso=/^\d{4}-\d{2}-\d{2}$/; if(!iso.test(String(r.ini))||!iso.test(String(r.fin))) throw new ApiError('BAD','Fechas inválidas');
  const admin=p.rol==='admin', sucSel=admin?String(r.sucursal||'TODAS').trim():p.suc; // el gerente SIEMPRE queda limitado a su sucursal
  const modo=admin&&r.modo===MODOS[1]?MODOS[1]:MODOS[0];
  const tmp=SpreadsheetApp.create('tmp-reporte-'+Date.now());
  try{
    const sh=tmp.getSheets()[0]; sh.setName('Reporte');
    const res=dibujarReporte_(sh,{k0:r.ini,k1:r.fin,sucSel,modo,pdf:true});
    if(!res.ok) throw new ApiError('EMPTY',res.msg.replace('⚠️ ',''));
    SpreadsheetApp.flush();
    const url='https://docs.google.com/spreadsheets/d/'+tmp.getId()+'/export?format=pdf&size=letter&portrait=false&fitw=true&gridlines=false&sheetnames=false&printtitle=false&pagenumbers=true&top_margin=0.3&bottom_margin=0.3&left_margin=0.3&right_margin=0.3&gid='+sh.getSheetId();
    let resp, n=0; // reintenta si Google responde con límite de uso o error temporal
    do{ resp=UrlFetchApp.fetch(url,{headers:{Authorization:'Bearer '+ScriptApp.getOAuthToken()},muteHttpExceptions:true}); if(resp.getResponseCode()===200) break; Utilities.sleep(1500*(++n)); }while(n<3);
    if(resp.getResponseCode()!==200) throw new Error('Google respondió '+resp.getResponseCode()+' al exportar');
    return {pdf:Utilities.base64Encode(resp.getBlob().getBytes()),nombre:'Asistencia_'+sucSel.replace(/\W+/g,'')+'_'+r.ini+'_a_'+r.fin+'.pdf'};
  }catch(e){
    if(e instanceof ApiError) throw e;
    console.error(e); throw new ApiError('PDF','No se pudo generar el PDF: '+e.message); // muestra la causa real en la página
  }finally{try{DriveApp.getFileById(tmp.getId()).setTrashed(true);}catch(e){}}
}
// Ejecutar desde el editor para crear un gerente de franquicia (cambia los 3 valores y borra la clave del código después)
function crearGerente(){
  const U='gerente.tamulte', CLAVE='CAMBIA_ESTA_CLAVE', SUC='TAMULTE';
  if(CLAVE.length<8||CLAVE==='CAMBIA_ESTA_CLAVE') throw new Error('Pon una clave de al menos 8 caracteres');
  const s=Utilities.getUuid();
  ss_().getSheetByName('admin').appendRow([U,'',new Date(),'activo',hash_(s,CLAVE),s,'gerente',SUC]);
  Logger.log('Gerente creado: '+U+' → '+SUC);
}

// ---------- Admin: alta/edición (solo descriptor facial, nunca foto) ----------
function guardar_(r){
  const id=String(r.id||'').trim().toUpperCase();
  if(!/^[A-Z0-9]{3,12}$/.test(id)) throw new ApiError('BAD','ID inválido (3-12 letras o números).');
  let ds=null;
  if(r.descriptors){
    if(r.consent!==true) throw new ApiError('BAD','Falta confirmar el consentimiento expreso por escrito del empleado.');
    if(!Array.isArray(r.descriptors)||!r.descriptors.length||r.descriptors.length>5||r.descriptors.some(a=>!Array.isArray(a)||a.length!==128||a.some(x=>!isFinite(x)))) throw new ApiError('BAD','Descriptor inválido');
    ds=JSON.stringify(r.descriptors.map(a=>a.map(x=>Math.round(x*1e4)/1e4)));
  }
  const lock=LockService.getScriptLock(); lock.waitLock(20000);
  try{
    const sh=ss_().getSheetByName('BD'), n=sh.getLastRow()-1, ids=n>0?sh.getRange(2,1,n,1).getValues():[];
    let row=0; ids.forEach((x,i)=>{if(String(x[0]).trim().toUpperCase()===id) row=i+2;});
    const now=Utilities.formatDate(new Date(),tz_(),'dd/MM/yyyy HH:mm:ss');
    if(!row){
      if(!r.nombre||!r.sucursal||!ds) throw new ApiError('BAD','Para un empleado nuevo captura nombre, sucursal y rostro.');
      row=sh.getLastRow()+1; sh.getRange(row,1,1,3).setNumberFormat('@').setValues([[id,r.sucursal,now]]);
    }
    const set=(c,v)=>sh.getRange(row,c).setValue(v);
    if(r.sucursal) set(2,r.sucursal); if(r.nombre) set(5,r.nombre); if(r.status) set(6,r.status); if(r.puesto) set(7,r.puesto);
    if(ds){set(9,ds); set(8,''); set(10,now); set(11,now);} // set(11)=fecha del consentimiento; borra la foto antigua
    if(String(r.status||'').toLowerCase()==='baja') set(9,''); // al dar de baja se elimina el dato biométrico
    bump_(); return {id};
  }finally{lock.releaseLock();}
}
function legacy_(r){ // entrega fotos antiguas (de 4 en 4) para convertirlas a descriptor en el navegador
  const sh=ss_().getSheetByName('BD'), n=sh.getLastRow()-1; if(n<1) return {lote:[]};
  const a=sh.getRange(2,1,n,1).getValues(), d9=sh.getRange(2,9,n,1).getValues(), skip=r.skip||[], lote=[];
  for(let i=0;i<n&&lote.length<4;i++){
    const id=String(a[i][0]).trim().toUpperCase(); if(!id||d9[i][0]||skip.indexOf(id)>=0) continue;
    const f=String(sh.getRange(i+2,8).getValue()||''); if(f.length<100) continue;
    lote.push({id,foto:f.indexOf('data:')===0?f:'data:image/jpeg;base64,'+f});
  }
  return {lote};
}

// ---------- Ejecutar UNA vez desde el editor ----------
function configuracionInicial(){
  const ss=ss_(), reg=ss.getSheetByName('Registros');
  if(reg.getMaxColumns()<20) reg.insertColumnsAfter(reg.getMaxColumns(),20-reg.getMaxColumns());
  reg.getRange(1,1,1,20).setValues([['ID','FECHAS','SUCURSAL','HORARIO','COORDENADAS','EMPLEADO','FECHA','HORA','SUCURSAL GPS','DIST (m)','PRECISION (m)','ESTATUS GPS','ID OPERACION','DISPOSITIVO','ORIGEN','CONFIANZA','SUCURSAL ASIGNADA','RECIBIDO','JORNADA','NOTA']]);
  reg.getRange('A2:A').setNumberFormat('@');
  ss.getSheetByName('BD').getRange(1,9,1,3).setValues([['descriptor','actualizado','consentimiento']]);
  ss.getSheetByName('admin').getRange(1,5,1,4).setValues([['Hash','Salt','Rol','Sucursal']]);
  instalarDisparadores(); if(!ss.getSheetByName('Reporte')) prepararReporte();
  bump_(); Logger.log('Listo. Ahora implementa la web app.');
}
