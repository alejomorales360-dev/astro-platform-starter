// COLACIONES SEMANALES - Google Apps Script backend
// Controla la inscripcion semanal al menu de colaciones de los trabajadores:
// evita que alguien quede sin anotarse y evita anotaciones duplicadas.
//
// Hojas requeridas en la planilla de Google Sheets ligada a este script:
//   Trabajadores: RUT | Nombre | Tipo | Activo | FechaInicio | FechaFin | Notas
//     - Tipo: "Fijo" o "Spot" (visitas/clientes temporales)
//     - Activo: "Si" o "No"
//     - FechaInicio / FechaFin: solo se usan para Tipo=Spot (rango de vigencia,
//       formato AAAA-MM-DD). Si estan vacias, se considera vigente siempre.
//   Menus: Semana | Dia | Opcion | Descripcion | Activo | Especial
//     - Semana: fecha del lunes de esa semana, formato AAAA-MM-DD
//     - Dia: Lunes, Martes, Miercoles, Jueves, Viernes
//     - Opcion: A, B, C...
//     - Especial: "Si" o "No" - marca un dia como almuerzo mejorado (para
//       destacarlo y llevar registro de quienes se anotaron ese dia)
//   Pedidos: ID | Semana | RUT | Nombre | Dia | Opcion | Timestamp
//     - Un trabajador solo puede tener UNA fila por (Semana, RUT, Dia): al
//       guardar un pedido para un dia ya elegido, se reemplaza la opcion en
//       vez de crear una fila nueva (evita anotaciones dobles).
//   Config: Clave | Valor
//     - admin_password: clave de acceso del panel de administracion

const HOJAS_COL = {
  TRABAJADORES: 'Trabajadores',
  MENUS: 'Menus',
  PEDIDOS: 'Pedidos',
  CONFIG: 'Config'
};

function doGet(e) {
  const callback = e && e.parameter && e.parameter.callback;
  const bodyParam = e && e.parameter && e.parameter.body;

  if (bodyParam) {
    try {
      const body = JSON.parse(decodeURIComponent(bodyParam));
      return responderJsonpCol(procesarAccionCol(body), callback);
    } catch (err) {
      return responderJsonpCol({ ok: false, error: err.toString() }, callback);
    }
  }

  try {
    return responderJsonpCol(obtenerTodoCol(), callback);
  } catch (err) {
    return responderJsonpCol({ ok: false, error: err.toString() }, callback);
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    return responderCol(procesarAccionCol(body));
  } catch (err) {
    return responderCol({ ok: false, error: err.toString(), stack: err.stack });
  }
}

function procesarAccionCol(body) {
  switch (body.action) {
    case 'loginAdmin':
      return verificarLoginAdminCol(body.password);
    case 'loginTrabajador':
      return loginTrabajadorCol(body.rut);
    case 'guardarTrabajador':
      return guardarTrabajadorCol(body.data, body.isEdit);
    case 'eliminarTrabajador':
      return eliminarTrabajadorCol(body.rut);
    case 'guardarMenu':
      return guardarMenuCol(body.data);
    case 'eliminarMenu':
      return eliminarMenuCol(body.semana, body.dia, body.opcion);
    case 'marcarDiaEspecial':
      return marcarDiaEspecialCol(body.semana, body.dia, body.especial);
    case 'copiarMenuSemana':
      return copiarMenuSemanaCol(body.semanaOrigen, body.semanaDestino);
    case 'guardarPedido':
      return guardarPedidoCol(body.data);
    case 'eliminarPedido':
      return eliminarPedidoCol(body.semana, body.rut, body.dia);
    default:
      return { ok: false, error: 'Accion no reconocida: ' + body.action };
  }
}

// --- RESPUESTAS ---
function responderCol(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}
function responderJsonpCol(data, callback) {
  const json = JSON.stringify(data);
  if (callback) {
    return ContentService.createTextOutput(callback + '(' + json + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

// --- HELPERS ---
function getHojaCol(nombre) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let h = ss.getSheetByName(nombre);
  if (!h) {
    const nombreLower = String(nombre).toLowerCase();
    h = ss.getSheets().find(s => s.getName().toLowerCase() === nombreLower);
  }
  if (!h) throw new Error('Hoja no encontrada: ' + nombre);
  return h;
}
function normalizarRutCol(rut) {
  return String(rut || '').toLowerCase().replace(/\./g, '').replace(/-/g, '').replace(/\s/g, '').trim();
}
function formatearFechaCol(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(v || '').trim();
}
// Compara una celda (que Sheets puede haber convertido a Date) contra un
// texto plano tipo 'AAAA-MM-DD', normalizando ambos lados igual.
function mismaFechaCol(celda, texto) {
  return formatearFechaCol(celda) === formatearFechaCol(texto);
}
function hojaAObjetosCol(nombreHoja) {
  const h = getHojaCol(nombreHoja);
  const v = h.getDataRange().getValues();
  if (v.length < 2) return [];
  const cab = v[0].map(c => String(c).trim().toLowerCase());
  return v.slice(1)
    .filter(f => f.some(c => String(c).trim() !== ''))
    .map((f, idx) => {
      const obj = { _fila: idx + 2 };
      cab.forEach((c, i) => {
        let val = f[i];
        if (val instanceof Date) val = formatearFechaCol(val);
        obj[c] = val != null ? val : '';
      });
      return obj;
    });
}
function obtenerTodoCol() {
  return {
    ok: true,
    trabajadores: hojaAObjetosCol(HOJAS_COL.TRABAJADORES),
    menus: hojaAObjetosCol(HOJAS_COL.MENUS),
    pedidos: hojaAObjetosCol(HOJAS_COL.PEDIDOS),
    config: obtenerConfigCol(),
    timestamp: new Date().toISOString()
  };
}
function obtenerConfigCol() {
  try {
    const v = getHojaCol(HOJAS_COL.CONFIG).getDataRange().getValues();
    const c = {};
    v.slice(1).forEach(f => { if (f[0]) c[String(f[0]).toLowerCase().trim()] = f[1]; });
    return c;
  } catch (err) {
    return {};
  }
}

// --- ADMIN LOGIN ---
function verificarLoginAdminCol(password) {
  const cfg = obtenerConfigCol();
  const clave = String(cfg.admin_password || '').trim();
  if (!clave) return { ok: false, error: 'No hay clave de administrador configurada (revisa la hoja Config).' };
  if (String(password || '').trim() === clave) return { ok: true };
  Utilities.sleep(400);
  return { ok: false, error: 'Clave incorrecta' };
}

// --- TRABAJADORES ---
function guardarTrabajadorCol(data, isEdit) {
  if (!data || !data.rut || !data.nombre) return { ok: false, error: 'Falta RUT o nombre' };
  const h = getHojaCol(HOJAS_COL.TRABAJADORES);
  const v = h.getDataRange().getValues();
  const rn = normalizarRutCol(data.rut);
  const fila = [
    String(data.rut).trim(),
    String(data.nombre).trim(),
    data.tipo === 'Spot' ? 'Spot' : 'Fijo',
    data.activo === false || data.activo === 'No' ? 'No' : 'Si',
    data.fechaInicio || '',
    data.fechaFin || '',
    data.notas || ''
  ];
  for (let i = 1; i < v.length; i++) {
    if (normalizarRutCol(v[i][0]) === rn) {
      h.getRange(i + 1, 1, 1, fila.length).setValues([fila]);
      return { ok: true, actualizado: true };
    }
  }
  if (isEdit) return { ok: false, error: 'Trabajador no encontrado para editar' };
  h.appendRow(fila);
  return { ok: true, creado: true };
}
function eliminarTrabajadorCol(rut) {
  const h = getHojaCol(HOJAS_COL.TRABAJADORES);
  const v = h.getDataRange().getValues();
  const rn = normalizarRutCol(rut);
  for (let i = 1; i < v.length; i++) {
    if (normalizarRutCol(v[i][0]) === rn) {
      h.deleteRow(i + 1);
      return { ok: true };
    }
  }
  return { ok: false, error: 'RUT no encontrado: ' + rut };
}

// --- MENUS ---
function guardarMenuCol(data) {
  if (!data || !data.semana || !data.dia || !data.opcion) return { ok: false, error: 'Faltan datos del menu' };
  const h = getHojaCol(HOJAS_COL.MENUS);
  const v = h.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) {
    if (mismaFechaCol(v[i][0], data.semana) &&
        String(v[i][1]).trim() === String(data.dia).trim() &&
        String(v[i][2]).trim() === String(data.opcion).trim()) {
      const fila = [
        data.semana, data.dia, data.opcion,
        data.descripcion || '',
        data.activo === false ? 'No' : 'Si'
      ];
      h.getRange(i + 1, 1, 1, fila.length).setValues([fila]);
      return { ok: true, actualizado: true };
    }
  }
  // Una opcion nueva hereda el estado "especial" que ya tenga ese dia
  // (si otras opciones del mismo dia estan marcadas como almuerzo mejorado).
  let especialDelDia = 'No';
  for (let i = 1; i < v.length; i++) {
    if (mismaFechaCol(v[i][0], data.semana) && String(v[i][1]).trim() === String(data.dia).trim() && String(v[i][5]).trim() === 'Si') {
      especialDelDia = 'Si';
      break;
    }
  }
  h.appendRow([
    data.semana, data.dia, data.opcion,
    data.descripcion || '',
    data.activo === false ? 'No' : 'Si',
    especialDelDia
  ]);
  return { ok: true, creado: true };
}
function marcarDiaEspecialCol(semana, dia, especial) {
  if (!semana || !dia) return { ok: false, error: 'Faltan datos del dia' };
  const h = getHojaCol(HOJAS_COL.MENUS);
  const v = h.getDataRange().getValues();
  const valor = especial ? 'Si' : 'No';
  let actualizadas = 0;
  for (let i = 1; i < v.length; i++) {
    if (mismaFechaCol(v[i][0], semana) && String(v[i][1]).trim() === String(dia).trim()) {
      h.getRange(i + 1, 6).setValue(valor);
      actualizadas++;
    }
  }
  if (!actualizadas) return { ok: false, error: 'Primero define al menos una opcion para ese dia.' };
  return { ok: true, actualizadas: actualizadas };
}
function eliminarMenuCol(semana, dia, opcion) {
  const h = getHojaCol(HOJAS_COL.MENUS);
  const v = h.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) {
    if (mismaFechaCol(v[i][0], semana) &&
        String(v[i][1]).trim() === String(dia).trim() &&
        String(v[i][2]).trim() === String(opcion).trim()) {
      h.deleteRow(i + 1);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Opcion de menu no encontrada' };
}
function copiarMenuSemanaCol(semanaOrigen, semanaDestino) {
  if (!semanaOrigen || !semanaDestino) return { ok: false, error: 'Faltan semanas' };
  const h = getHojaCol(HOJAS_COL.MENUS);
  const v = h.getDataRange().getValues();
  let copiadas = 0;
  const nuevasFilas = [];
  for (let i = 1; i < v.length; i++) {
    if (mismaFechaCol(v[i][0], semanaOrigen)) {
      nuevasFilas.push([semanaDestino, v[i][1], v[i][2], v[i][3], v[i][4], v[i][5] || 'No']);
      copiadas++;
    }
  }
  nuevasFilas.forEach(f => h.appendRow(f));
  return { ok: true, copiadas: copiadas };
}

// --- PEDIDOS ---
// Un trabajador solo puede tener UNA opcion elegida por dia dentro de la
// misma semana: si ya existe una fila para (semana, rut, dia) se actualiza
// en vez de duplicar. Esto es lo que impide las anotaciones dobles.
function guardarPedidoCol(data) {
  if (!data || !data.semana || !data.rut || !data.dia || !data.opcion) {
    return { ok: false, error: 'Faltan datos del pedido' };
  }
  const h = getHojaCol(HOJAS_COL.PEDIDOS);
  const v = h.getDataRange().getValues();
  const rn = normalizarRutCol(data.rut);
  const ahora = new Date().toISOString();
  for (let i = 1; i < v.length; i++) {
    if (mismaFechaCol(v[i][1], data.semana) &&
        normalizarRutCol(v[i][2]) === rn &&
        String(v[i][4]).trim() === String(data.dia).trim()) {
      h.getRange(i + 1, 6, 1, 2).setValues([[data.opcion, ahora]]);
      return { ok: true, actualizado: true };
    }
  }
  const id = Utilities.getUuid();
  h.appendRow([id, data.semana, data.rut, data.nombre || '', data.dia, data.opcion, ahora]);
  return { ok: true, creado: true, id: id };
}
function eliminarPedidoCol(semana, rut, dia) {
  const h = getHojaCol(HOJAS_COL.PEDIDOS);
  const v = h.getDataRange().getValues();
  const rn = normalizarRutCol(rut);
  for (let i = 1; i < v.length; i++) {
    if (mismaFechaCol(v[i][1], semana) &&
        normalizarRutCol(v[i][2]) === rn &&
        String(v[i][4]).trim() === String(dia).trim()) {
      h.deleteRow(i + 1);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Pedido no encontrado' };
}

// --- TRABAJADOR: login + su semana ---
function loginTrabajadorCol(rutIngresado) {
  if (!rutIngresado) return { ok: false, error: 'Ingresa tu RUT' };
  const rn = normalizarRutCol(rutIngresado);
  const trabajadores = hojaAObjetosCol(HOJAS_COL.TRABAJADORES);
  const t = trabajadores.find(x => normalizarRutCol(x.rut) === rn);
  if (!t) { Utilities.sleep(300); return { ok: false, error: 'RUT no encontrado. Consulta con administracion.' }; }
  if (String(t.activo).toLowerCase() !== 'si') {
    return { ok: false, error: 'Tu registro esta inactivo. Consulta con administracion.' };
  }
  const misPedidos = hojaAObjetosCol(HOJAS_COL.PEDIDOS).filter(p => normalizarRutCol(p.rut) === rn);
  return {
    ok: true,
    trabajador: { rut: String(t.rut), nombre: String(t.nombre), tipo: String(t.tipo || 'Fijo') },
    menus: hojaAObjetosCol(HOJAS_COL.MENUS),
    pedidos: misPedidos
  };
}

// --- UTILIDADES DE DIAGNOSTICO (ejecutar manualmente desde el editor) ---
function listarHojasCol() {
  const nombres = SpreadsheetApp.getActiveSpreadsheet().getSheets().map(s => s.getName());
  Logger.log('Hojas encontradas: ' + JSON.stringify(nombres));
  return nombres;
}
function testAPICol() {
  const d = obtenerTodoCol();
  Logger.log('Trabajadores:' + d.trabajadores.length + ' Menus:' + d.menus.length + ' Pedidos:' + d.pedidos.length);
}
function crearHojasIniciales() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const specs = {
    Trabajadores: ['RUT', 'Nombre', 'Tipo', 'Activo', 'FechaInicio', 'FechaFin', 'Notas'],
    Menus: ['Semana', 'Dia', 'Opcion', 'Descripcion', 'Activo', 'Especial'],
    Pedidos: ['ID', 'Semana', 'RUT', 'Nombre', 'Dia', 'Opcion', 'Timestamp'],
    Config: ['Clave', 'Valor']
  };
  Object.keys(specs).forEach(nombre => {
    let h = ss.getSheetByName(nombre);
    if (!h) h = ss.insertSheet(nombre);
    if (h.getLastRow() === 0) h.appendRow(specs[nombre]);
  });
  const cfg = ss.getSheetByName('Config');
  const v = cfg.getDataRange().getValues();
  const yaTiene = v.slice(1).some(f => String(f[0]).toLowerCase().trim() === 'admin_password');
  if (!yaTiene) cfg.appendRow(['admin_password', 'cambiar123']);
  Logger.log('Hojas listas. Recuerda cambiar la clave admin_password en la hoja Config.');
}
// Ejecutar UNA VEZ si tu hoja "Menus" ya existia antes de que se agregara la
// columna "Especial" (almuerzo mejorado). Agrega el encabezado si falta.
function agregarColumnaEspecial() {
  const h = getHojaCol(HOJAS_COL.MENUS);
  const encabezado = h.getRange(1, 1, 1, Math.max(6, h.getLastColumn())).getValues()[0];
  if (String(encabezado[5] || '').trim().toLowerCase() === 'especial') {
    Logger.log('La columna Especial ya existe.');
    return;
  }
  h.getRange(1, 6).setValue('Especial');
  Logger.log('Columna "Especial" agregada en Menus!F1.');
}
