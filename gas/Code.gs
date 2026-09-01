// TESORERIA 2 Basico A - Google Apps Script
// Version con soporte JSONP para acceso desde archivo local (file://)

const HOJAS = {
  ALUMNOS:'ALUMNOS', CONCEPTOS:'CONCEPTOS', PAGOS:'PAGOS',
  GASTOS:'GASTOS', CONFIGURACION:'CONFIGURACION', USUARIOS:'USUARIOS'
};

function doGet(e){
  const callback = e && e.parameter && e.parameter.callback;
  const bodyParam = e && e.parameter && e.parameter.body;

  // Si viene con ?body=... es un POST enviado via JSONP desde file://
  if(bodyParam){
    try{
      const body = JSON.parse(decodeURIComponent(bodyParam));
      const result = procesarAccion(body);
      return responderJsonp(result, callback);
    }catch(err){
      return responderJsonp({error:err.toString()}, callback);
    }
  }

  // GET normal: devolver todos los datos
  try{
    const data = obtenerTodo();
    return responderJsonp(data, callback);
  }catch(err){
    return responderJsonp({error:err.toString()}, callback);
  }
}

function doPost(e){
  try{
    const body = JSON.parse(e.postData.contents);
    const result = procesarAccion(body);
    return responder(result);
  }catch(err){
    return responder({error:err.toString(), stack:err.stack});
  }
}

// Funcion central que procesa cualquier accion
function procesarAccion(body){
  let result;
  switch(body.action){
    case 'login':             return verificarLogin(body.usuario, body.password);
    case 'loginApoderado':    return loginApoderado(body.rut);
    case 'uploadComprobante': return subirComprobante(body.data);
    case 'configurarDrive':   return inicializarEstructuraDrive(body.conceptos, body.alumnos);
    case 'saveAlumno':
      result=insertarFila(HOJAS.ALUMNOS,[
        body.data.id, body.data.apellido||'', body.data.nombre||'',
        body.data.rut||'', body.data.apoderado||'',
        body.data.email_apoderado||'', body.data.telefono||'', body.data.direccion||''
      ]); break;
    case 'deleteAlumno':      result=eliminarFila(HOJAS.ALUMNOS,body.id);    break;
    case 'savePago':
      result=insertarFila(HOJAS.PAGOS,[
        body.data.id,body.data.alumno_id,body.data.concepto_id,
        body.data.mes||'',body.data.ano||2026,
        body.data.monto,body.data.fecha,body.data.nota||''
      ]); break;
    case 'deletePago':        result=eliminarFila(HOJAS.PAGOS,body.id);      break;
    case 'saveGasto':
      result=insertarFila(HOJAS.GASTOS,[
        body.data.id,body.data.concepto,body.data.detalle||'',
        body.data.proveedor||'',body.data.monto,body.data.fecha,body.data.nota||''
      ]); break;
    case 'deleteGasto':       result=eliminarFila(HOJAS.GASTOS,body.id);     break;
    case 'saveConcepto':
      if(body.isEdit) eliminarFila(HOJAS.CONCEPTOS,body.data.id);
      result=insertarFila(HOJAS.CONCEPTOS,[
        body.data.id,body.data.nombre,body.data.tipo,body.data.monto||0
      ]); break;
    case 'deleteConcepto':    result=eliminarFila(HOJAS.CONCEPTOS,body.id);  break;
    case 'saveConfig':        result=actualizarConfig(body.clave,body.valor); break;
    default:                  result={error:'Accion no reconocida: '+body.action};
  }
  return {ok:true, result};
}

// Respuesta normal JSON
function responder(data){
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// Respuesta JSONP (para llamadas desde file://)
function responderJsonp(data, callback){
  const json = JSON.stringify(data);
  if(callback){
    // JSONP: callback(data)
    return ContentService.createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

// --- HELPERS ---
function getHoja(nombre){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  let h=ss.getSheetByName(nombre);
  if(!h){
    // getSheetByName es sensible a mayusculas/minusculas; buscamos por si la
    // hoja real esta nombrada distinto (ej. 'Alumnos' en vez de 'ALUMNOS')
    const nombreLower=String(nombre).toLowerCase();
    h=ss.getSheets().find(s=>s.getName().toLowerCase()===nombreLower);
  }
  if(!h) throw new Error('Hoja no encontrada: '+nombre);
  return h;
}
function normalizar(str){
  return str.toLowerCase()
    .replace(/n/g,'n').replace(/a/g,'a').replace(/e/g,'e')
    .replace(/i/g,'i').replace(/o/g,'o').replace(/u/g,'u')
    .replace(/\s+/g,'_').trim();
}
// Normalizar conservando tildes para busqueda de RUT
function normalizarStr(str){
  return str.toLowerCase()
    .replace(/\u00f1/g,'n')
    .replace(/[\u00e0-\u00e5]/g,'a')
    .replace(/[\u00e8-\u00eb]/g,'e')
    .replace(/[\u00ec-\u00ef]/g,'i')
    .replace(/[\u00f2-\u00f6]/g,'o')
    .replace(/[\u00f9-\u00fc]/g,'u')
    .replace(/\s+/g,'_').trim();
}
function normalizarRut(rut){
  return String(rut||'').toLowerCase().replace(/\./g,'').replace(/\s/g,'').trim();
}
function insertarFila(nombreHoja,datos){
  const h=getHoja(nombreHoja);
  h.appendRow(datos);
  const uf=h.getLastRow();
  h.getRange(uf,1,1,datos.length)
    .setBackground(uf%2===0?'#F6F5F0':'#FFFFFF')
    .setFontFamily('Arial').setFontSize(10);
  return{ok:true,fila:uf};
}
function eliminarFila(nombreHoja,id){
  const h=getHoja(nombreHoja);
  const v=h.getDataRange().getValues();
  for(let i=1;i<v.length;i++){
    if(String(v[i][0]).trim()===String(id).trim()){
      h.deleteRow(i+1);return{ok:true,fila:i+1};
    }
  }
  return{error:'ID no encontrado: '+id};
}
function hojaAObjetos(nombreHoja){
  const h=getHoja(nombreHoja);
  const v=h.getDataRange().getValues();
  if(v.length<2)return[];
  const cab=v[0].map(h=>normalizarStr(String(h)));
  return v.slice(1).filter(f=>{
    if(!f[0])return false;
    if(nombreHoja===HOJAS.ALUMNOS){const n=Number(f[0]);return!isNaN(n)&&n>0&&n===Math.floor(n);}
    return true;
  }).map(f=>{
    const obj={};
    cab.forEach((h,i)=>{
      let v=f[i];
      if(v instanceof Date)v=Utilities.formatDate(v,Session.getScriptTimeZone(),'yyyy-MM-dd');
      obj[h]=(v!=null)?v:'';
    });
    return obj;
  });
}
function obtenerTodo(){
  return{
    alumnos:hojaAObjetos(HOJAS.ALUMNOS),
    conceptos:hojaAObjetos(HOJAS.CONCEPTOS),
    pagos:hojaAObjetos(HOJAS.PAGOS),
    gastos:hojaAObjetos(HOJAS.GASTOS),
    config:obtenerConfig(),
    timestamp:new Date().toISOString()
  };
}
function obtenerConfig(){
  try{
    const v=getHoja(HOJAS.CONFIGURACION).getDataRange().getValues();
    const c={};
    v.slice(1).forEach(f=>{if(f[0])c[String(f[0]).toLowerCase()]=f[1];});
    return c;
  }catch{return{saldo_2025:20462};}
}
function actualizarConfig(clave,valor){
  const h=getHoja(HOJAS.CONFIGURACION);
  const v=h.getDataRange().getValues();
  for(let i=1;i<v.length;i++){
    if(String(v[i][0]).toLowerCase()===String(clave).toLowerCase()){
      h.getRange(i+1,2).setValue(valor);return{ok:true};
    }
  }
  h.appendRow([clave,valor,'']);return{ok:true,nuevo:true};
}
function verificarLogin(usuario,password){
  try{
    const h=getHoja(HOJAS.USUARIOS);
    const v=h.getDataRange().getValues();
    if(v.length<2)return{ok:false,error:'Sin usuarios configurados'};
    const cab=v[0].map(h=>normalizarStr(String(h)));
    for(let i=1;i<v.length;i++){
      const o={};cab.forEach((h,j)=>{o[h]=v[i][j];});
      if(String(o.usuario||'').toLowerCase().trim()===String(usuario||'').toLowerCase().trim()&&
         String(o.password||'').trim()===String(password||'').trim()){
        try{const c=cab.indexOf('ultimo_acceso');if(c>=0)h.getRange(i+1,c+1).setValue(new Date());}catch{}
        return{ok:true,user:{
          nombre:String(o.nombre||usuario),
          rol:String(o.rol||'tesorero'),
          usuario:String(o.usuario||usuario),
          email:String(o.email||'')
        }};
      }
    }
    Utilities.sleep(500);
    return{ok:false,error:'Usuario o contrasena incorrectos'};
  }catch(err){return{ok:false,error:'Error: '+err.message};}
}
function loginApoderado(rutIngresado){
  try{
    if(!rutIngresado)return{ok:false,error:'Ingresa el RUT'};
    const rn=normalizarRut(rutIngresado);
    const hA=getHoja(HOJAS.ALUMNOS);
    const vA=hA.getDataRange().getValues();
    if(vA.length<2)return{ok:false,error:'Sin datos'};
    const cab=vA[0].map(h=>normalizarStr(String(h)));
    let al=null;
    for(let i=1;i<vA.length;i++){
      const f=vA[i];if(!f[0])continue;
      const o={};cab.forEach((h,j)=>{o[h]=f[j];});
      if(normalizarRut(String(o.rut||''))===rn){
        al={id:Number(o.id),apellido:String(o.apellido||''),nombre:String(o.nombre||''),
          rut:String(o.rut||''),apoderado:String(o.apoderado||''),
          email:String(o.email_apoderado||''),telefono:String(o.telefono||''),
          direccion:String(o.direccion||'')};
        break;
      }
    }
    if(!al){Utilities.sleep(300);return{ok:false,error:'RUT no encontrado.'};}
    const hP=getHoja(HOJAS.PAGOS);
    const vP=hP.getDataRange().getValues();
    const cabP=vP[0].map(h=>normalizarStr(String(h)));
    const pagos=[];
    for(let i=1;i<vP.length;i++){
      const f=vP[i];if(!f[0])continue;
      const o={};cabP.forEach((h,j)=>{o[h]=f[j];});
      if(Number(o.alumno_id)===al.id){
        let fv=o.fecha;
        if(fv instanceof Date)fv=Utilities.formatDate(fv,Session.getScriptTimeZone(),'yyyy-MM-dd');
        pagos.push({id:String(o.id),concepto_id:String(o.concepto_id),
          mes:o.mes!==''?Number(o.mes):null,ano:Number(o.ano||2026),
          monto:parseFloat(o.monto)||0,fecha:String(fv||''),nota:String(o.nota||'')});
      }
    }
    const conceptos=hojaAObjetos(HOJAS.CONCEPTOS).map(c=>({
      id:String(c.id),nombre:String(c.nombre),tipo:String(c.tipo),monto:parseFloat(c.monto)||0
    }));
    return{ok:true,alumno:al,pagos,conceptos,config:obtenerConfig()};
  }catch(err){return{ok:false,error:'Error: '+err.message};}
}

// --- DRIVE ---
const DRIVE_ROOT='2 Basico A (2026)';
function getOrCreate(p,n){
  n=String(n||'x').replace(/[\/\\:*?"<>|]/g,'').replace(/\s+/g,' ').trim().substring(0,100);
  const it=p.getFoldersByName(n);return it.hasNext()?it.next():p.createFolder(n);
}
function subirComprobante(data){
  try{
    if(!data||!data.base64)return{ok:false,error:'Sin datos'};
    const root=getOrCreate(DriveApp.getRootFolder(),DRIVE_ROOT);
    let tf;
    if(data.tipo==='gasto'){tf=getOrCreate(getOrCreate(root,'GASTOS'),data.conceptoNombre||'Otros');}
    else{const ing=getOrCreate(root,'INGRESOS');const con=getOrCreate(ing,data.conceptoNombre||'Pagos');tf=data.alumnoNombre?getOrCreate(con,data.alumnoNombre):con;}
    const file=tf.createFile(Utilities.newBlob(Utilities.base64Decode(data.base64),data.mimeType||'application/octet-stream',data.filename||'comprobante'));
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK,DriveApp.Permission.VIEW);
    return{ok:true,url:file.getUrl(),id:file.getId(),nombre:file.getName()};
  }catch(err){return{ok:false,error:'Error Drive: '+err.message};}
}
function inicializarEstructuraDrive(conceptos,alumnos){
  try{
    const root=getOrCreate(DriveApp.getRootFolder(),DRIVE_ROOT);
    const ing=getOrCreate(root,'INGRESOS');const gas=getOrCreate(root,'GASTOS');
    let n=0;
    if(Array.isArray(conceptos)){conceptos.forEach(c=>{const cc=getOrCreate(ing,c.nombre||c.id);if(c.tipo!=='free'&&Array.isArray(alumnos)){alumnos.forEach(a=>{const nm=`${a.apellido||''}, ${a.nombre||''}`.trim().replace(/^,\s*/,'');if(nm){getOrCreate(cc,nm);n++;}});}});}
    ['Utiles','Actividades','Alimentacion','Transporte','Otros'].forEach(c=>getOrCreate(gas,c));
    return{ok:true,rootUrl:root.getUrl(),carpetasCreadas:n};
  }catch(err){return{ok:false,error:err.message};}
}
// Ejecuta esta funcion manualmente desde el editor de Apps Script (Ejecutar > listarHojas)
// para ver en el Registro (Ver > Registro) los nombres reales de las hojas de tu planilla.
function listarHojas(){
  const nombres=SpreadsheetApp.getActiveSpreadsheet().getSheets().map(s=>s.getName());
  Logger.log('Hojas encontradas: '+JSON.stringify(nombres));
  return nombres;
}
function testAPI(){
  const d=obtenerTodo();
  Logger.log('Alumnos:'+d.alumnos.length+' Pagos:'+d.pagos.length+' Gastos:'+d.gastos.length);
  Logger.log('Config:'+JSON.stringify(d.config));
}
function forzarPermisoDrive(){
  SpreadsheetApp.getActiveSpreadsheet();
  const f=getOrCreate(DriveApp.getRootFolder(),DRIVE_ROOT);
  Logger.log('Drive OK: '+f.getUrl());
}