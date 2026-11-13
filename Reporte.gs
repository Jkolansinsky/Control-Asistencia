// ============ REPORTE SEMANAL DE ASISTENCIA (hoja "Reporte") ============
const EMPRESA='impresto!'; // texto que sustituye al logo en el PDF
const LOGO_TEXTO_EN_HOJA=false; // false = en la hoja no se escribe el texto (usa tus imágenes); true = también aparece en la hoja
const MODOS=['Personal asignado a la sucursal','Registros hechos en la sucursal'];
const MES=['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre'];
const DSEM=['domingo','lunes','martes','miércoles','jueves','viernes','sábado'];

function onOpen(){
  try{SpreadsheetApp.getUi().createMenu('📋 Asistencia').addItem('Generar reporte','generarReporte').addItem('Preparar plantilla de reporte','prepararReporte').addToUi();}catch(e){}
}
function alEditar_(e){ // disparador instalable: cambios en BD/Sucursales y casilla "Generar" (I2)
  try{
    const sh=e.range.getSheet(), n=sh.getName();
    if(n==='BD'||n==='Sucursales') bump_();
    else if(n==='Reporte'&&e.range.getA1Notation()==='I2'&&e.range.getValue()===true){ try{generarReporte();}finally{e.range.setValue(false);} }
  }catch(err){console.error(err);}
}
function instalarDisparadores(){
  ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()==='alEditar_').forEach(t=>ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('alEditar_').forSpreadsheet(CFG.CLIENT_SS).onEdit().create();
}

// Panel de control (filas 1-2). El reporte se dibuja de la fila 3 hacia abajo.
function prepararReporte(){
  const ss=ss_(), sh=ss.getSheetByName('Reporte')||ss.insertSheet('Reporte',0), hoy=new Date();
  sh.getRange(1,1,sh.getMaxRows(),sh.getMaxColumns()).breakApart().clear().clearDataValidations();
  sh.getRange('B1').setValue('Fecha inicio'); sh.getRange('C1').setValue('Fecha fin'); sh.getRange('D1').setValue('Sucursal');
  sh.getRange('F1').setValue('Mostrar'); sh.getRange('I1').setValue('☑ Generar');
  sh.getRange('D1:E1').merge(); sh.getRange('F1:H1').merge(); sh.getRange('D2:E2').merge(); sh.getRange('F2:H2').merge();
  sh.getRange('A1:J1').setFontWeight('bold').setFontColor('#555555').setFontSize(9);
  sh.getRange('B2:C2').setValues([[new Date(hoy.getFullYear(),hoy.getMonth(),hoy.getDate()-6),hoy]]).setNumberFormat('dd/MM/yyyy')
    .setDataValidation(SpreadsheetApp.newDataValidation().requireDate().build());
  sh.getRange('D2').setValue('TODAS').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['TODAS'].concat(sucursales_().map(s=>s.nombre)),true).build());
  sh.getRange('F2').setValue(MODOS[0]).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(MODOS,true).build());
  sh.getRange('I2').insertCheckboxes();
  sh.getRange('B2:I2').setBackground('#FDECEE').setHorizontalAlignment('center').setFontWeight('bold');
  sh.getRange('J2').setValue('Elige fechas (máx. 31 días), sucursal y marca ☑ Generar.').setFontColor('#555555');
  sh.setColumnWidths(1,1,34); sh.setColumnWidths(2,1,200); sh.setColumnWidths(3,1,165);
}

function generarReporte(){
  const sh=ss_().getSheetByName('Reporte');
  if(!sh){prepararReporte();return;}
  const st=t=>{sh.getRange('J2').setValue(t);SpreadsheetApp.flush();};
  const ini=sh.getRange('B2').getValue(), fin=sh.getRange('C2').getValue();
  if(!(ini instanceof Date)||!(fin instanceof Date)) return st('⚠️ Captura fechas válidas en B2 y C2');
  st('⏳ Generando reporte…');
  let r; try{r=dibujarReporte_(sh,{k0:fk_(ini),k1:fk_(fin),sucSel:String(sh.getRange('D2').getValue()||'TODAS').trim(),modo:String(sh.getRange('F2').getValue())});}
  catch(err){console.error(err); return st('❌ Error: '+err.message);} // el error queda visible en J2 y en Ejecuciones
  st(r.ok?'✔ '+r.n+' empleados · '+r.nd+' días · generado '+Utilities.formatDate(new Date(),tz_(),'dd/MM HH:mm'):r.msg);
}

// Dibuja el reporte en la hoja "sh" (la pestaña Reporte, o un archivo temporal para generar el PDF)
function dibujarReporte_(sh,o){
  const ss=ss_(), k0=o.k0, k1=o.k1, sucSel=o.sucSel||'TODAS', suc=sucSel.toUpperCase(), modo=o.modo;
  if(k0>k1) return {ok:false,msg:'⚠️ La fecha de inicio es mayor que la final'};
  const tz=tz_(), dia=(k,n)=>fk_(new Date(Utilities.parseDate(k,tz,'yyyy-MM-dd').getTime()+n*864e5));
  const dias=[]; for(let k=k0,i=0;k<=k1&&i<40;k=dia(k,1),i++) dias.push(k);
  if(dias.length>31) return {ok:false,msg:'⚠️ Máximo 31 días por reporte'};
  const H=CFG.HORARIOS, nd=dias.length, TC=7+nd, hoyK=fk_(new Date());
  // ---- 1. leer registros del rango (±1 día por turnos que cruzan la medianoche): hoja principal + archivo histórico ----
  const rows=leerRegistros_(dia(k0,-1),dia(k1,1));
  // ---- 2. agrupar por empleado y JORNADA ----
  const bd=empMap_(), E={};
  const get=(id,n)=>E[id]||(E[id]={id,nombre:n||(bd[id]?bd[id].n:id),asig:bd[id]?bd[id].s:'',d:{}});
  rows.forEach(x=>{
    const id=String(x[0]).trim().toUpperCase(); if(!id) return;
    const j=x[18]?fk_(x[18]):fk_(x[6]); if(j<k0||j>k1) return;
    const o=get(id,String(x[5]||'')), a=String(x[16]||'').trim(); if(!o.asig&&a) o.asig=a;
    const dd=o.d[j]||(o.d[j]={slots:{}}), h=String(x[3]), t=ts_(x);
    if(!dd.slots[h]||t<dd.slots[h].t) dd.slots[h]={t,s:String(x[2]||'').trim()};
  });
  Object.keys(bd).forEach(id=>{if(bd[id].a) get(id);}); // plantilla activa: así se ven también las faltas
  const list=Object.keys(E).map(k=>E[k]).filter(o=>{
    if(suc==='TODAS') return true;
    if(modo===MODOS[1]) return Object.keys(o.d).some(j=>Object.keys(o.d[j].slots).some(h=>o.d[j].slots[h].s.toUpperCase()===suc));
    return o.asig.toUpperCase()===suc;
  }).sort((p,q)=>p.nombre.localeCompare(q.nombre));
  // ---- 3. limpiar y dibujar encabezado ----
  sh.setFrozenRows(0); sh.setFrozenColumns(0); // quita inmovilizados de versiones anteriores: impiden combinar celdas
  const z=sh.getRange(3,1,sh.getMaxRows()-2,sh.getMaxColumns()); z.breakApart(); z.clear();
  if(!list.length) return {ok:false,msg:'⚠️ No hay empleados para esos filtros'};
  if(sh.getMaxColumns()<TC) sh.insertColumnsAfter(sh.getMaxColumns(),TC-sh.getMaxColumns());
  const R0=8, NR=list.length*7; if(sh.getMaxRows()<R0+NR) sh.insertRowsAfter(sh.getMaxRows(),R0+NR-sh.getMaxRows());
  const pk=k=>{const p=k.split('-').map(Number);return {y:p[0],m:p[1],d:p[2]};}, a0=pk(k0), b0=pk(k1);
  const tit='Reporte de Asistencia\n'+(a0.y===b0.y&&a0.m===b0.m?'del '+a0.d+' al '+b0.d+' de '+MES[b0.m-1]+' del '+b0.y
    :'del '+a0.d+' de '+MES[a0.m-1]+(a0.y!==b0.y?' del '+a0.y:'')+' al '+b0.d+' de '+MES[b0.m-1]+' del '+b0.y);
  sh.setColumnWidths(4,nd,88); sh.setColumnWidths(4+nd,4,62);
  const hd=[0,1,2].map(()=>Array(TC).fill(''));
  hd[2][0]='';hd[0][1]='Sucursal:';hd[1][1]=sucSel;
  dias.forEach((k,c)=>{const p=pk(k);hd[1][3+c]=DSEM[new Date(p.y,p.m-1,p.d).getDay()];hd[2][3+c]=p.d;});
  ['Asist. Correcta','Sin Asist.','Asist. incompleta','Hras. Trabajadas'].forEach((t,i)=>hd[0][3+nd+i]=t);
  sh.getRange(5,1,3,TC).setValues(hd).setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.getRange(3,4).setValue(tit);
  if(EMPRESA&&(o.pdf||LOGO_TEXTO_EN_HOJA)){sh.getRange(3,1).setValue(EMPRESA); sh.getRange(3,4+nd).setValue(EMPRESA);}
  sh.getRange(3,1,2,3).merge().setFontSize(28).setFontWeight('bold').setFontStyle('italic').setFontColor('#e60012').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.getRange(3,4,2,nd).merge().setFontSize(16).setFontWeight('bold').setWrap(true).setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.getRange(3,4+nd,2,4).merge().setFontSize(18).setFontWeight('bold').setFontStyle('italic').setFontColor('#e60012').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.getRange(5,2).setFontSize(14); sh.getRange(6,2).setFontSize(14).setFontColor('#e60012');
  sh.getRange(6,4,1,nd).setBackground('#1F2937').setFontColor('#ffffff'); sh.getRange(7,4,1,nd).setBackground('#E5E7EB');
  for(let c=0,s=0;c<=nd;c++){ // encabezado de mes agrupado
    const p=c<nd?pk(dias[c]):null, q=pk(dias[s]);
    if(c===nd||p.m!==q.m||p.y!==q.y){const r=sh.getRange(5,4+s,1,c-s); r.setValue(MES[q.m-1]+' de '+q.y); if(c-s>1) r.merge(); s=c;}
  }
  for(let i=0;i<4;i++) sh.getRange(5,4+nd+i,3,1).merge().setBackground('#ffff00').setTextRotation(90).setWrap(true);
  sh.setRowHeights(3,2,30); sh.setRowHeights(5,1,44); sh.setRowHeights(6,2,28);
  // ---- 4. un bloque de 7 filas por empleado ----
  const LAB=['🟢 Hora Entrada','🟡 Salida Comida','🔵 Entrada Comida','🔴 Hora Salida','Suc. provisional :','Tiempo de Comida (min)','Horas Trabajadas x día'];
  const RB=['#ffffff','#ffffff','#ffffff','#ffffff','#eeeeee','#fce4d6','#ffffff'];
  const fmtS=ms=>{const s=Math.round(ms/1000);return Math.floor(s/3600)+':'+('0'+Math.floor(s%3600/60)).slice(-2)+':'+('0'+s%60).slice(-2);};
  const hms=(t,j)=>{const d=new Date(t),s=Utilities.formatDate(d,tz,'H:mm:ss');return fk_(d)!==j?s+' (+1)':s;};
  const V=[],B=[],F=[];
  list.forEach((o,n)=>{
    const v=LAB.map(()=>Array(TC).fill('')), b=RB.map(c=>Array(TC).fill('#ffffff').map((w,i)=>i>=2&&i<3+nd?c:w)), f=LAB.map(()=>Array(TC).fill('#000000'));
    v[0][0]=n+1; v[0][1]=o.nombre; LAB.forEach((l,r)=>{v[r][2]=l;}); f[6].fill('#c00000');
    let ok=0,sin=0,inc=0,tot=0;
    dias.forEach((k,c)=>{
      const col=3+c, dd=o.d[k]; if(k>hoyK) return;
      if(!dd||!Object.keys(dd.slots).length){sin++;for(let r=0;r<7;r++){b[r][col]='#e7e6e6';} for(let r=0;r<4;r++){v[r][col]='---';f[r][col]='#999999';} return;}
      const T=h=>dd.slots[h]?dd.slots[h].t:0, e=T(H[0]), sc=T(H[1]), ec=T(H[2]), sa=T(H[3]);
      H.forEach((h,r)=>{if(dd.slots[h]) v[r][col]=hms(dd.slots[h].t,k); else{v[r][col]='---';f[r][col]='#999999'; if(r===0||r===3) b[r][col]='#fde2e1';}});
      const sucs=[]; Object.keys(dd.slots).map(h=>dd.slots[h]).sort((p,q)=>p.t-q.t).forEach(s=>{if(s.s&&sucs.indexOf(s.s)<0) sucs.push(s.s);});
      const otras=o.asig?sucs.filter(s=>s.toUpperCase()!==o.asig.toUpperCase()):[]; // solo sucursales distintas a la asignada en BD
      if(otras.length){v[4][col]=otras.join(' → ');b[4][col]='#fff200';}
      if(sc&&ec&&ec>sc) v[5][col]=Math.round((ec-sc)/6e4)+' min'; else if(sc) v[5][col]='Sin regreso';
      if(e&&sa&&sa>e){let ms=sa-e; if(sc&&ec&&ec>sc) ms-=ec-sc; v[6][col]=fmtS(ms); tot+=ms; ok++;} else{v[6][col]='---';f[6][col]='#999999'; inc++;}
    });
    v[0][3+nd]=ok; v[0][4+nd]=sin; v[0][5+nd]=inc; v[0][6+nd]=fmtS(tot);
    v.forEach(x=>V.push(x)); b.forEach(x=>B.push(x)); f.forEach(x=>F.push(x));
  });
  sh.getRange(R0,4,NR,nd).setNumberFormat('@'); sh.getRange(R0,4+nd+1,NR,1).setNumberFormat('0;-0;;@'); sh.getRange(R0,7+nd,NR,1).setNumberFormat('@');
  const body=sh.getRange(R0,1,NR,TC);
  body.setValues(V).setBackgrounds(B).setFontColors(F).setFontSize(10).setVerticalAlignment('middle').setHorizontalAlignment('center');
  sh.getRange(R0,3,NR,1).setHorizontalAlignment('left');
  sh.setRowHeights(R0,NR,20);
  sh.getRange(5,1,3+NR,TC).setBorder(true,true,true,true,true,true,'#bfbfbf',SpreadsheetApp.BorderStyle.SOLID);
  list.forEach((o,n)=>{
    const r=R0+n*7;
    [1,2,4+nd,5+nd,6+nd,7+nd].forEach(c=>sh.getRange(r,c,7,1).merge());
    sh.getRange(r,2).setFontWeight('bold').setWrap(true); sh.getRange(r,4+nd,1,4).setFontWeight('bold');
    sh.getRange(r+6,3,1,1+nd).setFontStyle('italic');
    sh.getRange(r,1,7,TC).setBorder(null,null,true,null,null,null,'#ff0000',SpreadsheetApp.BorderStyle.SOLID_THICK);
  });
  // sh.setFrozenRows(7); sh.setFrozenColumns(3);
  return {ok:true,n:list.length,nd};
}
