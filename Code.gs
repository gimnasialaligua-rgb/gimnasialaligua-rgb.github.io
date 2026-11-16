// Servidor de la Copa La Ligua. Instalar en una planilla privada de Google Sheets.
const COLS_ = ['id','bib','name','team','level','band','salto','viga','suelo','barras','archived'];
const ROLES_ = ['admin'];
const EVENTS_ = ['salto','viga','suelo','barras'];
const TEAMS_ = ['Pesasus','Quillota','Calera','M. Unamuno','Sport House','La Ligua','Power Flex','Legend Gymnastics'];
const LEVELS_ = ['Mini','Pre nivel','Nivel 1','Nivel 2','Nivel 3','Nivel 4','Intermedio','Novato','Máster'];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index').setTitle('Copa La Ligua · Puntajes')
    .addMetaTag('viewport','width=device-width, initial-scale=1');
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Copa La Ligua')
    .addItem('Preparar planilla','prepararCopa_')
    .addItem('Renovar todos los accesos','renovarAccesos_').addToUi();
}

function prepararCopa_() {
  const lock=acquireWriteLock_();
  try {
    const book = SpreadsheetApp.getActiveSpreadsheet();
    const props = PropertiesService.getScriptProperties();
    props.setProperty('SHEET_ID',book.getId());
    for (const [name,headers] of [['Nomina',['Dorsal','Nombre completo','Equipo','Nivel','Franja']],['Detalle de puntajes',['Dorsal','Gimnasta','Equipo','Nivel','Franja','Salto','Viga','Suelo','Barras','All Around','Estado','ID']],['Puntajes',COLS_],['Historial',['fecha','responsable','accion','gimnasta_id','campo','antes','despues']],['Accesos',['responsable','codigo_privado']]]) {
      const sheet = book.getSheetByName(name)||book.insertSheet(name);
      if (!sheet.getLastRow()) { sheet.appendRow(headers); sheet.setFrozenRows(1); }
    }
    if (!props.getProperty('AUTH_VERSION')) crearAccesos_(book,props);
    syncDetail_(book);
    SpreadsheetApp.flush();
  } finally { lock.releaseLock(); }
}

// Se ejecuta solamente desde el menú de la planilla por su propietario.
function renovarAccesos_() {
  const ui = SpreadsheetApp.getUi();
  if (ui.alert('Renovar accesos','Se cerrarán las sesiones de todas las personas. ¿Continuar?',ui.ButtonSet.YES_NO)!==ui.Button.YES) return;
  const lock=acquireWriteLock_();
  try { crearAccesos_(SpreadsheetApp.getActiveSpreadsheet(),PropertiesService.getScriptProperties()); }
  finally { lock.releaseLock(); }
}

function crearAccesos_(book,props) {
  Object.keys(props.getProperties()).filter(k=>k.indexOf('ACCESS_')===0).forEach(k=>props.deleteProperty(k));
  props.setProperty('AUTH_VERSION',Utilities.getUuid());
  const rows=[['responsable','codigo_privado']];
  for (const role of ROLES_) {
    const code=Utilities.getUuid().replace(/-/g,'')+Utilities.getUuid().replace(/-/g,'');
    props.setProperty('ACCESS_'+hash_(code),JSON.stringify({id:Utilities.getUuid(),name:role,role,enabled:true}));
    rows.push([role,code]);
  }
  const sheet=book.getSheetByName('Accesos'); sheet.clearContents();
  sheet.getRange(1,1,rows.length,2).setValues(rows);sheet.setFrozenRows(1);
}

function hash_(text) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,text,Utilities.Charset.UTF_8)
    .map(b=>('0'+((b+256)%256).toString(16)).slice(-2)).join('');
}
function book_() {
  const id=PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if(!id)throw Error('La organización todavía debe preparar la planilla.');
  return SpreadsheetApp.openById(id);
}
function auth_(token) {
  if(typeof token!=='string'||token.length>150)throw Error('Inicia sesión con tu código de acceso.');
  const text=CacheService.getScriptCache().get('session:'+hash_(token));
  if(!text)throw Error('Tu sesión terminó. Vuelve a ingresar el código de acceso.');
  const actor=JSON.parse(text);
  if(actor.version!==PropertiesService.getScriptProperties().getProperty('AUTH_VERSION'))throw Error('Los accesos fueron renovados. Solicita tu nuevo código.');
  const access=access_(actor.accessHash);
  if(!access||!access.enabled)throw Error('Este acceso fue revocado. Contacta a la administración.');
  if(!ROLES_.includes(access.role))throw Error('Este acceso fue revocado. Contacta a la administración.');
  return {...actor,...access};
}
function access_(hash) {
  if(!hash)return null;
  const raw=PropertiesService.getScriptProperties().getProperty('ACCESS_'+hash);
  if(!raw)return null;
  return raw[0]==='{'?JSON.parse(raw):{id:hash,name:raw,role:raw,enabled:true};
}
function requireAdmin_(token) {
  const actor=auth_(token);if(actor.role!=='admin')throw Error('Solo administración tiene permiso para esta acción.');return actor;
}
function normalizedName_(name) {return String(name||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').trim().replace(/\s+/g,' ').toLowerCase();}
function login(code,name) {
  if(typeof code!=='string'||code.length<32||code.length>128)throw Error('Código de acceso no válido.');
  const props=PropertiesService.getScriptProperties();
  const accessHash=hash_(code.trim()),access=access_(accessHash);
  if(!access||!access.enabled||!ROLES_.includes(access.role))throw Error('Código de acceso no válido o revocado.');
  if(typeof name!=='string'||name.length>80||!normalizedName_(name)||normalizedName_(name)!==normalizedName_(access.name))throw Error('El nombre y el código no corresponden a un acceso autorizado.');
  const token=Utilities.getUuid()+Utilities.getUuid();
  CacheService.getScriptCache().put('session:'+hash_(token),JSON.stringify({accessHash,version:props.getProperty('AUTH_VERSION')}),21600);
  return {token,role:access.role,name:access.name,id:access.id};
}
function logout(token) { if(typeof token==='string')CacheService.getScriptCache().remove('session:'+hash_(token)); return true; }
function records_(sheet) {
  if(sheet.getLastRow()<2)return [];
  return sheet.getRange(2,1,sheet.getLastRow()-1,COLS_.length).getValues().map((r,i)=>({row:i+2,data:Object.fromEntries(COLS_.map((k,j)=>[k,r[j]]))}));
}
function athlete_(record) {
  const a={};for(const key of COLS_.slice(1,-1))a[key]=EVENTS_.includes(key)?(record[key]===''?null:record[key]):String(record[key]);return a;
}
function readCup(token) {
  // Las lecturas no toman el candado de escritura: una importación no impide ingresar.
  auth_(token);
  const out={};
  for(const r of records_(book_().getSheetByName('Puntajes')))if(r.data.archived!==true){
    const a=athlete_(r.data);
    out[r.data.id]=a;
  }
  return out;
}
function appendRows_(sheet,rows){if(rows.length)sheet.getRange(sheet.getLastRow()+1,1,rows.length,rows[0].length).setValues(rows);}
function acquireWriteLock_(){const lock=LockService.getScriptLock();if(!lock.tryLock(10000))throw Error('Otra operación está guardando datos. Espera unos segundos antes de volver a guardar o importar.');return lock;}
function validate_(a) {
  if(typeof a.bib!=='string'||!a.bib.trim()||a.bib.length>12||typeof a.name!=='string'||!a.name.trim()||a.name.length>100)throw Error('Nombre o dorsal no válido.');
  if(!TEAMS_.includes(a.team)||!LEVELS_.includes(a.level))throw Error('Equipo o nivel no válido.');
  const max=(a.level==='Mini'||a.level==='Intermedio')?2:3;
  if(!(a.level==='Novato'&&a.band==='')&&(!['1','2','3'].includes(String(a.band))||Number(a.band)>max))throw Error('Franja no válida.');
  for(const e of EVENTS_)if(a[e]!==null&&(typeof a[e]!=='number'||!Number.isFinite(a[e])||a[e]<0||a[e]>20||Math.abs(a[e]*1000-Math.round(a[e]*1000))>0.000001))throw Error('Puntaje no válido.');
  // Impide interpretar nombres o dorsales como fórmulas de planilla.
  for(const k of ['bib','name'])if(/^[\s]*[=+@-]/.test(a[k]))throw Error('El nombre y dorsal no pueden comenzar con =, +, @ o -.');
}
function writeCup(token,id,patch,remove,expected,requestId) {
  requireAdmin_(token);
  if(typeof id!=='string'||!/^[a-zA-Z0-9_-]{1,80}$/.test(id))throw Error('Identificador inválido.');
  if(typeof requestId!=='string'||!/^[a-zA-Z0-9_-]{1,80}$/.test(requestId))throw Error('Solicitud inválida.');
  const lock=acquireWriteLock_();
  try {
    const cache=CacheService.getScriptCache(), dedupe='write:'+hash_(token+requestId);
    const cached=cache.get(dedupe);if(cached)return JSON.parse(cached);
    const book=book_(),sheet=book.getSheetByName('Puntajes'),records=records_(sheet),found=records.find(r=>r.data.id===id);
    const actor=requireAdmin_(token);
    let result,detailRecord;
    if(remove) {
      if(!found||found.data.archived===true)throw Error('Esta gimnasta ya no está activa.');
      sheet.getRange(found.row,11).setValue(true);
      book.getSheetByName('Historial').appendRow([new Date().toISOString(),actor.name+' ('+actor.role+')','archivar',id,'registro',JSON.stringify(athlete_(found.data)),'archivado']);
      detailRecord={row:found.row,data:{...found.data,archived:true}};
      result={removed:true};
    } else {
      if(!patch||typeof patch!=='object'||Array.isArray(patch))throw Error('Datos no válidos.');
      const keys=Object.keys(patch);
      if(!keys.length||keys.some(k=>!COLS_.slice(1,-1).includes(k)))throw Error('Campo no permitido.');
      if(found&&found.data.archived===true)throw Error('Esta gimnasta fue retirada de la nómina.');
      const before=found?athlete_(found.data):null;
      const next={...(before||{}),...patch};validate_(next);
      if(records.some(r=>r.data.id!==id&&r.data.archived!==true&&String(r.data.bib).toLowerCase()===next.bib.toLowerCase()))throw Error('Ese dorsal ya está registrado.');
      if(found)for(const k of keys)if(!expected||expected[k]!==before[k])throw Error('Otro dispositivo modificó este dato. Actualiza los resultados antes de volver a guardarlo.');
      const row=COLS_.map(k=>k==='id'?id:k==='archived'?false:(next[k]===null?'':next[k]));
      const targetRow=found?found.row:sheet.getLastRow()+1;
      sheet.getRange(targetRow,1,1,COLS_.length).setValues([row]);
      detailRecord={row:targetRow,data:{id,...next,archived:false}};
      appendRows_(book.getSheetByName('Historial'),keys.map(k=>[new Date().toISOString(),actor.name+' ('+actor.role+')',found?'puntaje':'alta',id,k,JSON.stringify(before?before[k]:null),JSON.stringify(next[k])]));
      result={athlete:next};
    }
    syncDetailRow_(book,detailRecord);SpreadsheetApp.flush();cache.put(dedupe,JSON.stringify(result),600);return result;
  } finally {lock.releaseLock();}
}

function listAccess(token) {
  const actor=requireAdmin_(token),props=PropertiesService.getScriptProperties().getProperties();
  return Object.keys(props).filter(k=>k.indexOf('ACCESS_')===0).map(k=>{
    const a=access_(k.slice(7));return {id:a.id,name:a.name,role:a.role,enabled:a.enabled,self:a.id===actor.id};
  }).filter(a=>ROLES_.includes(a.role));
}
function createAccess(token,name,role) {
  const actor=requireAdmin_(token);
  if(typeof name!=='string'||!name.trim()||name.length>80||/^[\s]*[=+@-]/.test(name))throw Error('Ingresa un nombre válido.');
  if(!ROLES_.includes(role))throw Error('Permiso no válido.');
  const lock=acquireWriteLock_();
  try {
    const all=PropertiesService.getScriptProperties().getProperties();
    if(Object.keys(all).filter(k=>k.indexOf('ACCESS_')===0).some(k=>{const a=access_(k.slice(7));return a.enabled&&normalizedName_(a.name)===normalizedName_(name);}))throw Error('Ya existe una persona activa con ese nombre. Usa el nombre completo para distinguirla.');
    const code=Utilities.getUuid().replace(/-/g,'')+Utilities.getUuid().replace(/-/g,'');
    const access={id:Utilities.getUuid(),name:name.trim(),role,enabled:true};
    PropertiesService.getScriptProperties().setProperty('ACCESS_'+hash_(code),JSON.stringify(access));
    book_().getSheetByName('Historial').appendRow([new Date().toISOString(),actor.name+' ('+actor.role+')','crear_acceso',access.id,'permiso','',role]);
    return {...access,code};
  }finally{lock.releaseLock();}
}
function revokeAccess(token,id) {
  const actor=requireAdmin_(token);if(id===actor.id)throw Error('No puedes revocar tu propio acceso.');
  const lock=acquireWriteLock_();
  try {
    const props=PropertiesService.getScriptProperties(),all=props.getProperties();
    const key=Object.keys(all).filter(k=>k.indexOf('ACCESS_')===0).find(k=>access_(k.slice(7)).id===id);
    if(!key)throw Error('Acceso no encontrado.');
    const access=access_(key.slice(7));access.enabled=false;props.setProperty(key,JSON.stringify(access));
    book_().getSheetByName('Historial').appendRow([new Date().toISOString(),actor.name+' ('+actor.role+')','revocar_acceso',id,'permiso',access.role,'revocado']);
    return true;
  }finally{lock.releaseLock();}
}
function adminReport(token) {
  requireAdmin_(token);const lock=acquireWriteLock_();
  try {
    const book=book_(),athletes={};
    for(const r of records_(book.getSheetByName('Puntajes')))athletes[r.data.id]={...athlete_(r.data),archived:r.data.archived===true};
    const history=book.getSheetByName('Historial');
    return {athletes,history:history.getLastRow()>1?history.getRange(2,1,history.getLastRow()-1,7).getValues().map(row=>row.map(v=>v instanceof Date?v.toISOString():v)):[],exportedAt:new Date().toISOString()};
  }finally{lock.releaseLock();}
}

function syncDetail_(book) {
  const sheet=book.getSheetByName('Detalle de puntajes')||book.insertSheet('Detalle de puntajes');
  sheet.clearContents();
  const headers=['Dorsal','Gimnasta','Equipo','Nivel','Franja','Salto','Viga','Suelo','Barras','All Around','Estado','ID'];
  sheet.getRange(1,1,1,12).setValues([headers]);sheet.setFrozenRows(1);
  const records=records_(book.getSheetByName('Puntajes'));
  if(records.length){
    const rows=records.map(r=>{const a=athlete_(r.data);return [a.bib,a.name,a.team,a.level,a.band,a.salto??'',a.viga??'',a.suelo??'',a.barras??'','',r.data.archived?'Archivada':'Activa',r.data.id];});
    sheet.getRange(2,1,rows.length,12).setValues(rows);
    sheet.getRange(2,10,rows.length,1).setFormulas(rows.map((_,i)=>[`=IF(COUNT(F${i+2}:I${i+2})=4,SUM(F${i+2}:I${i+2}),"")`]));
    // Google Sheets puede importar tablas con formato de columna tipado.
    // Aplicar el formato a una sola columna por llamada.
    for(let column=6;column<=10;column++){
      sheet.getRange(2,column,rows.length,1).setNumberFormat('0.000');
    }
  }
  sheet.getRange(1,1,1,12).setBackground('#491b83').setFontColor('#ffffff').setFontWeight('bold');
  for(let column=1;column<=12;column++){
    sheet.setColumnWidth(column,column===2?240:column===12?290:120);
  }
}

function importNomina(token) {
  const actor=requireAdmin_(token),lock=acquireWriteLock_();
  try {
    const book=book_(),input=book.getSheetByName('Nomina');
    if(!input||input.getLastRow()<2)throw Error('Completa primero la hoja Nomina en Google Sheets.');
    const current=records_(book.getSheetByName('Puntajes')),seen=new Set();
    const incoming=input.getRange(2,1,input.getLastRow()-1,5).getValues().filter(r=>r.some(v=>v!==''));
    const staged=[];
    incoming.forEach((row,index)=>{
      const a={bib:String(row[0]).trim(),name:String(row[1]).trim(),team:String(row[2]).trim(),level:String(row[3]).trim(),band:String(row[4]).trim(),salto:null,viga:null,suelo:null,barras:null};
      try{validate_(a);}catch(e){throw Error('Nómina, registro '+(index+1)+': '+e.message);}
      const bib=a.bib.toLowerCase();if(seen.has(bib))throw Error('Dorsal repetido en la nómina: '+a.bib);seen.add(bib);
      const existing=current.find(r=>r.data.archived!==true&&String(r.data.bib).toLowerCase()===bib);
      if(existing){if(['name','team','level','band'].some(k=>String(existing.data[k])!==String(a[k])))throw Error('El dorsal '+a.bib+' ya existe con otros datos. Revisa la nómina antes de importar.');}
      else staged.push({id:Utilities.getUuid(),...a});
    });
    // Dos escrituras en bloque en vez de dos llamadas por gimnasta.
    appendRows_(book.getSheetByName('Puntajes'),staged.map(a=>COLS_.map(k=>k==='archived'?false:(a[k]??''))));
    appendRows_(book.getSheetByName('Historial'),staged.map(a=>[new Date().toISOString(),actor.name+' (admin)','importar_nomina',a.id,'registro','',JSON.stringify(a)]));
    syncDetail_(book);SpreadsheetApp.flush();return {added:staged.length,total:current.filter(r=>r.data.archived!==true).length+staged.length};
  }finally{lock.releaseLock();}
}

function syncDetailRow_(book,record){
  const sheet=book.getSheetByName('Detalle de puntajes');
  if(!sheet)return syncDetail_(book);
  const a=athlete_(record.data),row=record.row;
  sheet.getRange(row,1,1,12).setValues([[a.bib,a.name,a.team,a.level,a.band,a.salto??'',a.viga??'',a.suelo??'',a.barras??'','',record.data.archived?'Archivada':'Activa',record.data.id]]);
  sheet.getRange(row,10,1,1).setFormulas([[`=IF(COUNT(F${row}:I${row})=4,SUM(F${row}:I${row}),"")`]]);
}
