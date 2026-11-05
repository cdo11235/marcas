/**
 * MI REGISTRO DE JORNADA — Google Apps Script
 * 1) Crear una Hoja de cálculo de Google.
 * 2) Extensiones > Apps Script; reemplazar Código.gs con este contenido.
 * 3) Crear archivo HTML llamado Index, con el archivo Index.html entregado.
 * 4) Cambiar CORREO_DESTINO y ejecutar instalarSistema() UNA VEZ.
 * 5) Implementar > Nueva implementación > Aplicación web;
 *    ejecutar como: Yo; acceso: Solo yo.
 */
const CORREO_DESTINO = 'CAMBIAR_POR_TU_CORREO@ejemplo.com';
const ZONA_HORARIA = 'America/Santiago';
const NOMBRE_HOJA = 'Marcaciones';
const MAX_REGISTROS = 10000;
const CABECERAS = ['ID', 'Tipo', 'Fecha ISO', 'Comentario', 'Última modificación ISO', 'Turno'];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Mi registro de jornada');
}

/** Instalar una sola vez desde el editor de Apps Script. */
function instalarSistema() {
  verificarCorreo_();
  const activa = SpreadsheetApp.getActiveSpreadsheet();
  if (!activa) throw new Error('Abre este script desde Extensiones > Apps Script de una hoja de cálculo.');
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', activa.getId());
  const hoja = obtenerHoja_();
  // Se usa un activador diario para poder reintentar en caso de fallo del día 1.
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'enviarResumenSiCorresponde') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('enviarResumenSiCorresponde').timeBased().everyDays(1).atHour(8).create();
  return 'Configurado. Datos: ' + hoja.getName() + '. Correo: ' + CORREO_DESTINO +
    '. Envío: día 1 de cada mes, alrededor de las 08:00 (hora de Santiago), con reintentos hasta el día 7.';
}

function obtenerHoja_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) throw new Error('Primero ejecuta instalarSistema() desde el editor de Apps Script.');
  const libro = SpreadsheetApp.openById(id);
  let hoja = libro.getSheetByName(NOMBRE_HOJA);
  if (!hoja) hoja = libro.insertSheet(NOMBRE_HOJA);
  if (hoja.getLastRow() === 0) {
    hoja.getRange(1, 1, 1, CABECERAS.length).setValues([CABECERAS]);
    hoja.getRange(1, 1, 1, CABECERAS.length).setFontWeight('bold');
    hoja.setFrozenRows(1);
    hoja.getRange('A:F').setNumberFormat('@'); // Preservar ISO y comentarios como texto.
  }
  if (hoja.getRange(1, 6).getValue() !== 'Turno') hoja.getRange(1, 6).setValue('Turno');
  return hoja;
}

function leerRegistros_() {
  const hoja = obtenerHoja_();
  const n = hoja.getLastRow();
  if (n < 2) return [];
  return hoja.getRange(2, 1, n - 1, 6).getDisplayValues()
    .map(r => ({ id: r[0], type: r[1], iso: r[2], note: r[3] || '', shift: r[5] || '' }))
    .filter(r => r.id && (r.type === 'entrada' || r.type === 'salida') && !isNaN(Date.parse(r.iso)));
}

function listarRegistros() { return leerRegistros_(); }
function obtenerEstado() {
  const props = PropertiesService.getScriptProperties();
  return { correo: CORREO_DESTINO, ultimoInforme: props.getProperty('ULTIMO_ENVIO_MES') || '' };
}
function validarTipo_(tipo) {
  if (tipo !== 'entrada' && tipo !== 'salida') throw new Error('Tipo de marcación inválido.');
}
function validarNota_(nota) {
  const texto = String(nota || '').trim().slice(0, 2000);
  // Evitar que un comentario introducido por el usuario se ejecute como fórmula de Sheets.
  return /^[=+\-@]/.test(texto) ? "'" + texto : texto;
}
function validarISO_(iso) {
  const txt = String(iso || '');
  const d = new Date(txt);
  if (!/^\d{4}-\d\d-\d\dT\d\d:\d\d:\d\d(?:\.\d{3})?Z$/.test(txt) ||
      isNaN(d.getTime()) || d.getUTCFullYear() < 2000 || d.getUTCFullYear() > 2100)
    throw new Error('Fecha u hora inválida.');
  return d.toISOString();
}
function validarTurno_(v) { if (v !== '' && v !== 'AM' && v !== 'PM') throw new Error('Turno inválido.'); return v || ''; }
function conBloqueo_(f) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try { return f(); } finally { lock.releaseLock(); }
}
function anexar_(registro) {
  const sh = obtenerHoja_();
  if (sh.getLastRow() >= MAX_REGISTROS + 1) throw new Error('Se alcanzó el límite de registros.');
  sh.appendRow([registro.id, registro.type, registro.iso, registro.note, new Date().toISOString(), registro.shift || '']);
}

/** Marca con la hora del servidor, no con la hora modificable del teléfono. */
function registrarMarcacion(tipo, nota, turno) {
  validarTipo_(tipo);
  return conBloqueo_(() => {
    const prev = leerRegistros_().sort((a,b) => Date.parse(a.iso)-Date.parse(b.iso));
    const ultima = prev.length ? prev[prev.length - 1] : null;
    const abierta = ultima && ultima.type === 'entrada';
    if ((tipo === 'entrada' && abierta) || (tipo === 'salida' && !abierta))
      throw new Error('La secuencia de entrada/salida no permite esta marcación. Corrige el historial si es necesario.');
    anexar_({id: Utilities.getUuid(), type:tipo, iso:new Date().toISOString(), note:validarNota_(nota),shift:validarTurno_(turno || '')});
    return leerRegistros_();
  });
}

function agregarMarcacionManual(tipo, iso, nota, turno) {
  validarTipo_(tipo);
  const fecha = validarISO_(iso);
  return conBloqueo_(() => {
    anexar_({id:Utilities.getUuid(),type:tipo,iso:fecha,note:validarNota_(nota),shift:validarTurno_(turno || '')});
    return leerRegistros_();
  });
}

function actualizarMarcacion(id, tipo, iso, nota, turno) {
  validarTipo_(tipo);
  const fecha = validarISO_(iso);
  return conBloqueo_(() => {
    const sh = obtenerHoja_(), n = sh.getLastRow();
    if (n < 2) throw new Error('No hay marcaciones para editar.');
    const ids = sh.getRange(2,1,n-1,1).getDisplayValues();
    const pos = ids.findIndex(r => r[0] === id);
    if (pos < 0) throw new Error('El registro ya no existe; recarga la aplicación.');
    sh.getRange(pos+2,1,1,6).setValues([[id,tipo,fecha,validarNota_(nota),new Date().toISOString(),validarTurno_(turno || '')]]);
    return leerRegistros_();
  });
}

function eliminarMarcacion(id) {
  return conBloqueo_(() => {
    const sh = obtenerHoja_(), n = sh.getLastRow();
    if (n < 2) throw new Error('No hay marcaciones para eliminar.');
    const ids = sh.getRange(2,1,n-1,1).getDisplayValues();
    const pos = ids.findIndex(r => r[0] === id);
    if (pos < 0) throw new Error('El registro ya no existe; recarga la aplicación.');
    sh.deleteRow(pos+2);
    return leerRegistros_();
  });
}

/** Solo importar una copia propia; sustituye completamente la tabla. */
function reemplazarRegistros(registros) {
  if (!Array.isArray(registros) || registros.length > MAX_REGISTROS)
    throw new Error('El respaldo excede el máximo de ' + MAX_REGISTROS + ' registros.');
  const ids = new Set();
  const clean = registros.map(r => {
    validarTipo_(r.type);
    const id = String(r.id || Utilities.getUuid()).slice(0,120);
    if (ids.has(id)) throw new Error('IDs duplicados en el archivo.');
    ids.add(id);
    return {id:id,type:r.type,iso:validarISO_(r.iso),note:validarNota_(r.note),shift:validarTurno_(r.shift || '')};
  });
  return conBloqueo_(() => {
    const sh = obtenerHoja_();
    const n = sh.getLastRow();
    if (n > 1) sh.getRange(2,1,n-1,6).clearContent();
    if (clean.length) {
      const rows = clean.map(r=>[r.id,r.type,r.iso,r.note,new Date().toISOString(),r.shift || '']);
      sh.getRange(2,1,rows.length,5).setValues(rows);
    }
    return leerRegistros_();
  });
}

function verificarCorreo_() {
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(CORREO_DESTINO) || CORREO_DESTINO.includes('CAMBIAR_'))
    throw new Error('Configura CORREO_DESTINO en Code.gs antes de instalar o enviar correos.');
}
function mesAnterior_(ahora) {
  const local = Utilities.formatDate(ahora, ZONA_HORARIA, 'yyyy-MM');
  const partes = local.split('-').map(Number);
  const y = partes[1] === 1 ? partes[0]-1 : partes[0];
  const m = partes[1] === 1 ? 12 : partes[1]-1;
  return {anio:y, mes:m, clave:y+'-'+('0'+m).slice(-2)};
}
function generarPares_(data) {
  const ordered = data.slice().sort((a,b) => {
    const da=Date.parse(a.iso), db=Date.parse(b.iso);
    return da-db || (a.type === b.type ? 0 : (a.type === 'entrada' ? -1 : 1));
  });
  let abierta=null; const pares=[]; const duraciones={};
  ordered.forEach(r=>{
    if (r.type === 'entrada') abierta=r;
    else if (abierta) {
      const a=Date.parse(abierta.iso),b=Date.parse(r.iso);
      if (b>=a) { pares.push({ini:a,fin:b,salidaId:r.id});duraciones[r.id]=b-a; }
      abierta=null;
    }
  });
  return {pares:pares,duraciones:duraciones};
}
/** Calcula el comienzo local de un mes, independiente de la zona del proyecto. */
function inicioMes_(anio,mes) {
  // Mediodía UTC corresponde a mañana local en Santiago y permite leer el desfase.
  const mediodiaUtc = new Date(Date.UTC(anio, mes - 1, 1, 12));
  const z = Utilities.formatDate(mediodiaUtc, ZONA_HORARIA, 'Z');
  const signo = z[0] === '-' ? -1 : 1;
  const desfase = signo * (Number(z.slice(1,3)) * 60 + Number(z.slice(3,5)));
  return Date.UTC(anio, mes - 1, 1) - desfase * 60000;
}
function duracion_(ms) {
  const min = Math.floor(Math.max(0,ms)/60000);
  return ('0'+Math.floor(min/60)).slice(-2)+':'+('0'+(min%60)).slice(-2);
}

/** Reglas: lunes-martes 17:30/18:00; miércoles-viernes 16:30/17:00. */
function calcularJornada_(entrada, salida) {
  const di = new Date(entrada.iso);
  const fecha = Utilities.formatDate(di,ZONA_HORARIA,'yyyy-MM-dd');
  const dia = new Date(fecha+'T12:00:00Z').getUTCDay();
  const hi = Utilities.formatDate(di,ZONA_HORARIA,'HH:mm').split(':').map(Number);
  const he = hi[0]*60+hi[1];
  const turno = entrada.shift || (he < 8*60+15 ? 'AM' : 'PM');
  if (dia === 0 || dia === 6) return {turno:turno,atraso:null,extra:null,horario:'Sin horario de fin de semana'};
  const inicio = turno==='AM'?480:510;
  const fin = (dia <= 2 ? 1050:990) + (turno==='PM'?30:0);
  const salidaFecha = Utilities.formatDate(new Date(salida.iso),ZONA_HORARIA,'yyyy-MM-dd');
  if (salidaFecha !== fecha) return {turno:turno,atraso:Math.max(0,he-inicio-5),extra:null,horario:'Salida en otra fecha'};
  const hs = Utilities.formatDate(new Date(salida.iso),ZONA_HORARIA,'HH:mm').split(':').map(Number);
  const se=hs[0]*60+hs[1];
  return {turno:turno,atraso:Math.max(0,he-inicio-5),extra:se-fin>=30?se-fin:0,horario:(turno==='AM'?'08:00':'08:30')+'–'+('0'+Math.floor(fin/60)).slice(-2)+':'+('0'+fin%60).slice(-2)};
}

function csvCelda_(valor) {
  let str=String(valor == null ? '' : valor);
  // Evita fórmulas de Excel cuando se abre el CSV.
  if (/^[\s]*[=+\-@]/.test(str)) str="'"+str;
  return '"'+str.replace(/"/g,'""')+'"';
}
function crearInforme_(mes) {
  const data = leerRegistros_();
  const items = data.filter(r=>Utilities.formatDate(new Date(r.iso),ZONA_HORARIA,'yyyy-MM')===mes.clave)
    .sort((a,b)=>Date.parse(a.iso)-Date.parse(b.iso));
  const par = generarPares_(data);
  // Inicio y fin del mes según la hora de Santiago, incluso si cambia el huso horario.
  const comienzo = inicioMes_(mes.anio,mes.mes);
  const final = inicioMes_(mes.anio,mes.mes+1);
  const ms = par.pares.reduce((sum,p)=>sum+Math.max(0,Math.min(p.fin,final)-Math.max(p.ini,comienzo)),0);
  const jornadas = par.pares.filter(p=>p.fin>=comienzo&&p.fin<final).length;
  const paresTurnos = {}; const ordenados=data.slice().sort((a,b)=>Date.parse(a.iso)-Date.parse(b.iso)); let activa=null;
  ordenados.forEach(r=>{if(r.type==='entrada')activa=r;else if(activa){const j=calcularJornada_(activa,r);paresTurnos[activa.id]=j;paresTurnos[r.id]=j;activa=null;}});
  const filas = [['Fecha','Hora','Marcación','Turno','Atraso (min)','Horas extra (min)','Comentario','Duración (hh:mm)']];
  items.forEach(r => {
    const d=new Date(r.iso);
    filas.push([
      Utilities.formatDate(d,ZONA_HORARIA,'dd/MM/yyyy'),
      Utilities.formatDate(d,ZONA_HORARIA,'HH:mm:ss'),
      r.type==='entrada'?'Entrada':'Salida',
      (paresTurnos[r.id]||{}).turno || r.shift || '',
      r.type==='entrada'?((paresTurnos[r.id]||{}).atraso ?? ''):'',
      r.type==='salida'?((paresTurnos[r.id]||{}).extra ?? ''):'',
      r.note || '',
      par.duraciones[r.id] == null?'':duracion_(par.duraciones[r.id])
    ]);
  });
  const csv='\ufeff'+filas.map(row=>row.map(csvCelda_).join(';')).join('\r\n')+'\r\n';
  const resumen = 'Informe de jornada: '+mes.clave+'\nMarcaciones: '+items.length+
    '\nJornadas completas (con salida en el mes): '+jornadas+
    '\nHoras cerradas registradas en el período: '+duracion_(ms)+'\n\nSe adjunta CSV compatible con Excel.'+
    '\nLas horas se calculan emparejando entradas y salidas; no incluyen turnos abiertos ni descansos descontados automáticamente.';
  return {csv:csv,resumen:resumen,cantidad:items.length,jornadas:jornadas,horas:duracion_(ms)};
}
function enviarInforme_(mes,esPrueba) {
  verificarCorreo_();
  const informe=crearInforme_(mes);
  const prefijo=esPrueba?'[PRUEBA] ':'';
  MailApp.sendEmail({
    to:CORREO_DESTINO,
    subject:prefijo+'Registro de jornada — '+mes.clave,
    body:informe.resumen,
    attachments:[Utilities.newBlob(informe.csv,'text/csv','registro_jornada_'+mes.clave+'.csv')],
    name:'Mi registro de jornada'
  });
  return 'Informe de '+mes.clave+' enviado a '+CORREO_DESTINO+' ('+informe.cantidad+' marcaciones).';
}
/** Activador automático diario: envía SOLO una vez por mes en los primeros 7 días. */
function enviarResumenSiCorresponde() {
  const ahora=new Date();
  const dia=Number(Utilities.formatDate(ahora,ZONA_HORARIA,'d'));
  if (dia<1||dia>7) return 'Fuera de ventana: no se envía.';
  return conBloqueo_(() => {
    const mes=mesAnterior_(ahora);
    const p=PropertiesService.getScriptProperties();
    if(p.getProperty('ULTIMO_ENVIO_MES')===mes.clave) return 'Informe ya enviado: '+mes.clave;
    const mensaje=enviarInforme_(mes,false);
    p.setProperty('ULTIMO_ENVIO_MES',mes.clave);
    return mensaje;
  });
}
/** Enviar el informe del mes anterior como prueba, sin alterar el programado. */
function enviarInformePrueba() {
  return enviarInforme_(mesAnterior_(new Date()),true);
}
