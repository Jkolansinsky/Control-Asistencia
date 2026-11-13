// ============ ARCHIVO MENSUAL DE REGISTROS (mantiene ligera la hoja "Registros") ============
// Cada mes mueve los registros de meses anteriores a un archivo "Registros AAAA" (una pestaña por mes: AAAA-MM),
// en la misma carpeta de Drive que tu hoja principal. El reporte lee ese archivo cuando pides fechas antiguas.
const ENC_REG=['ID','FECHAS','SUCURSAL','HORARIO','COORDENADAS','EMPLEADO','FECHA','HORA','SUCURSAL GPS','DIST (m)','PRECISION (m)','ESTATUS GPS','ID OPERACION','DISPOSITIVO','ORIGEN','CONFIANZA','SUCURSAL ASIGNADA','RECIBIDO','JORNADA','NOTA'];
const FMT_REG=['@','dd/MM/yyyy HH:mm:ss','@','@','@','@','dd/MM/yyyy','@','@','0','0','@','@','@','@','0.000','@','dd/MM/yyyy HH:mm:ss','dd/MM/yyyy','@'];
const clave_=x=>String(x[12]||'')||(String(x[0]).toUpperCase()+'|'+(x[1] instanceof Date?x[1].getTime():x[1])+'|'+x[3]);
const pad20_=r=>{const y=r.slice(0,20);while(y.length<20)y.push('');return y;};

function archivoAnio_(y,crear){
  const p=PropertiesService.getScriptProperties(), k='ARCH:'+y, id=p.getProperty(k);
  if(id){try{return SpreadsheetApp.openById(id);}catch(e){}}
  const nombre='Registros '+y, pa=DriveApp.getFileById(CFG.CLIENT_SS).getParents(), carpeta=pa.hasNext()?pa.next():null;
  if(carpeta){const it=carpeta.getFilesByName(nombre); if(it.hasNext()){const f=it.next(); p.setProperty(k,f.getId()); return SpreadsheetApp.openById(f.getId());}}
  if(!crear) return null;
  const s=SpreadsheetApp.create(nombre); if(carpeta) DriveApp.getFileById(s.getId()).moveTo(carpeta);
  p.setProperty(k,s.getId()); return s;
}
function pestanaMes_(arch,mes){
  let t=arch.getSheetByName(mes);
  if(!t){
    t=arch.insertSheet(mes); t.getRange(1,1,1,20).setValues([ENC_REG]).setFontWeight('bold'); t.setFrozenRows(1);
    arch.getSheets().forEach(s=>{if(s.getName()!==mes&&/^(Hoja ?1|Sheet ?1)$/i.test(s.getName())&&s.getLastRow()<=1&&arch.getSheets().length>1) arch.deleteSheet(s);});
  }
  return t;
}

// Se ejecuta sola el día 3 de cada mes (así ya cerraron los turnos que cruzan la medianoche del día 1). También puedes correrla a mano.
function archivarRegistros(){
  const lock=LockService.getScriptLock(); lock.waitLock(30000);
  try{
    const sh=ss_().getSheetByName('Registros'), last=sh.getLastRow(), tz=tz_();
    if(last<2) return 'Sin registros';
    const corte=Utilities.formatDate(new Date(),tz,'yyyy-MM')+'-01', nc=Math.min(20,sh.getMaxColumns());
    const data=sh.getRange(2,1,last-1,nc).getValues(), porMes={}, idx=[];
    data.forEach((x,i)=>{const k=fk_(x[6])||fk_(x[1]); if(/^\d{4}-\d{2}-\d{2}$/.test(k)&&k<corte){(porMes[k.slice(0,7)]=porMes[k.slice(0,7)]||[]).push(pad20_(x)); idx.push(i);}});
    if(!idx.length) return 'Nada que archivar (se conserva desde '+corte+')';
    Object.keys(porMes).sort().forEach(mes=>{ // 1) copiar y verificar
      const t=pestanaMes_(archivoAnio_(+mes.slice(0,4),true),mes), ya=new Set(), lt=t.getLastRow();
      if(lt>=2) t.getRange(2,1,lt-1,20).getValues().forEach(x=>ya.add(clave_(x)));
      const nuevas=porMes[mes].filter(x=>!ya.has(clave_(x)));
      if(nuevas.length){
        const ini=lt<1?2:lt+1;
        t.getRange(ini,1,nuevas.length,20).setNumberFormats(nuevas.map(()=>FMT_REG)).setValues(nuevas);
        SpreadsheetApp.flush();
        if(t.getLastRow()!==ini-1+nuevas.length) throw new Error('Falló la verificación al archivar '+mes+'; no se borró nada');
      }
    });
    if(last-1-idx.length<1) sh.insertRowsAfter(sh.getMaxRows(),5); // la hoja principal nunca queda sin filas
    const bloques=[]; idx.forEach(i=>{const f=i+2,b=bloques[bloques.length-1]; if(b&&f===b[1]+1) b[1]=f; else bloques.push([f,f]);});
    for(let j=bloques.length-1;j>=0;j--) sh.deleteRows(bloques[j][0],bloques[j][1]-bloques[j][0]+1); // 2) borrar solo lo ya copiado
    PropertiesService.getScriptProperties().setProperty('ARCH_ULTIMO',new Date().toISOString());
    const msg=idx.length+' registros archivados ('+Object.keys(porMes).sort().join(', ')+')'; Logger.log(msg); return msg;
  }finally{lock.releaseLock();}
}
function instalarArchivoMensual(){
  ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()==='archivarRegistros').forEach(t=>ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('archivarRegistros').timeBased().onMonthDay(3).atHour(3).create();
  Logger.log('Archivo mensual programado: día 3 de cada mes, 3:00 a.m.');
}

// Lee registros con FECHA entre m0 y m1 (yyyy-MM-dd): hoja principal + archivo histórico. Quita duplicados por seguridad.
function leerRegistros_(m0,m1){
  const reg=ss_().getSheetByName('Registros'), last=reg.getLastRow(), out=[], vistos=new Set();
  const add=x=>{const k=clave_(x); if(!vistos.has(k)){vistos.add(k); out.push(x);}};
  if(last>=2){
    const g=reg.getRange(2,7,last-1,1).getValues(); let a=-1,b=-1;
    g.forEach((v,i)=>{const k=fk_(v[0]); if(k>=m0&&k<=m1){if(a<0)a=i;b=i;}});
    if(a>=0) reg.getRange(a+2,1,b-a+1,Math.min(20,reg.getMaxColumns())).getValues().forEach(x=>add(pad20_(x)));
  }
  const mesHoy=Utilities.formatDate(new Date(),tz_(),'yyyy-MM'), cache={};
  for(let y=+m0.slice(0,4),m=+m0.slice(5,7);;){
    const mes=y+'-'+('0'+m).slice(-2); if(mes>m1.slice(0,7)) break;
    if(mes<mesHoy){ // solo los meses anteriores al actual pueden estar archivados
      if(!(y in cache)) cache[y]=archivoAnio_(y,false);
      const t=cache[y]&&cache[y].getSheetByName(mes), lt=t?t.getLastRow():0;
      if(lt>=2) t.getRange(2,1,lt-1,20).getValues().forEach(x=>{const k=fk_(x[6]); if(k>=m0&&k<=m1) add(x);});
    }
    if(++m>12){m=1;y++;}
  }
  return out;
}
