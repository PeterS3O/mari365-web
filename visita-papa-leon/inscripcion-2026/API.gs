const CONFIG = {
  SPREADSHEET_ID: '12b2T7yNNkBGIbUe2XmWG4oQ1coLeKaXMHzikOsIKYPs',
  HOJA_INSCRIPCIONES: 'Inscripciones',
  ESTADO_INICIAL: 'inscripto',
  ORIGEN_DEFAULT: 'web'
};

const HEADERS = [
  'ID',
  'Timestamp',
  'Nombre',
  'Apellido',
  'DNI',
  'Fecha de nacimiento',
  'Email',
  'Celular',
  'Parroquia',
  'Comunidad',
  'Restricciones alimenticias',
  'Otra restriccion',
  'Estado',
  'Origen'
];

function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  try {
    const payload = getPayload(e);
    const action = payload.action || 'health';
    if (action === 'health') return json({ ok: true, service: 'visita-papa-leon-inscripcion-2026' });
    if (action === 'registrarInscripcion') return json(registrarInscripcion(payload));
    return json({ ok: false, error: 'Accion no reconocida.' });
  } catch (err) {
    return json({ ok: false, error: err.message || 'Error interno.' });
  }
}

function autorizarBackend() {
  const sheet = getSheet();
  return 'Backend autorizado. Filas actuales: ' + sheet.getLastRow();
}

function getPayload(e) {
  if (!e) return {};
  if (e.postData && e.postData.contents) {
    try {
      return JSON.parse(e.postData.contents);
    } catch (err) {
      throw new Error('Solicitud invalida.');
    }
  }
  return e.parameter || {};
}

function registrarInscripcion(payload) {
  const data = normalizarInscripcion(payload);
  const sheet = getSheet();
  asegurarEncabezados(sheet);

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    if (dniYaInscripto(sheet, data.dni)) {
      return {
        ok: false,
        code: 'DNI_DUPLICADO',
        error: 'Ya estas registrado. Por favor comunicate con tu catequista.'
      };
    }
    const id = Utilities.getUuid();
    sheet.appendRow([
      id,
      new Date(),
      data.nombre,
      data.apellido,
      data.dni,
      data.fechaNacimiento,
      data.email,
      textForSheet(data.celular),
      data.parroquia,
      data.comunidad,
      data.restricciones,
      data.otraRestriccion,
      CONFIG.ESTADO_INICIAL,
      data.origen
    ]);
    return { ok: true, id };
  } finally {
    lock.releaseLock();
  }
}

function normalizarInscripcion(payload) {
  const data = {
    nombre: cleanText(payload.nombre),
    apellido: cleanText(payload.apellido),
    dni: cleanDni(payload.dni),
    fechaNacimiento: cleanText(payload.fechaNacimiento),
    email: cleanEmail(payload.email),
    celular: normalizePhone(payload.celular),
    parroquia: cleanText(payload.parroquia),
    comunidad: cleanText(payload.comunidad),
    restricciones: cleanText(payload.restricciones),
    otraRestriccion: cleanText(payload.otraRestriccion),
    origen: cleanText(payload.origen) || CONFIG.ORIGEN_DEFAULT
  };

  if (!data.nombre) throw new Error('Ingresá el nombre.');
  if (!data.apellido) throw new Error('Ingresá el apellido.');
  if (!/^\d{7,9}$/.test(data.dni)) throw new Error('Ingresá un DNI valido.');
  if (!data.fechaNacimiento) throw new Error('Ingresá la fecha de nacimiento.');
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(data.email)) throw new Error('Ingresá un email valido.');
  if (!/^\+\d{10,15}$/.test(data.celular)) {
    throw new Error('Ingresá el celular completo con codigo de pais. Ejemplo: +54 9 379 4000000.');
  }
  if (!data.parroquia) throw new Error('Seleccioná una parroquia.');
  if (!data.comunidad) throw new Error('Seleccioná una comunidad.');
  if (!data.restricciones) throw new Error('Seleccioná una restriccion alimenticia.');
  if (data.restricciones === 'Otros' && !data.otraRestriccion) {
    throw new Error('Detallá la restriccion alimenticia.');
  }
  return data;
}

function getSheet() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.HOJA_INSCRIPCIONES);
  if (!sheet) throw new Error('No existe la hoja de inscripciones.');
  return sheet;
}

function asegurarEncabezados(sheet) {
  const current = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  const ok = HEADERS.every((header, index) => current[index] === header);
  if (!ok) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
  }
}

function dniYaInscripto(sheet, dni) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;
  const values = sheet.getRange(2, 5, lastRow - 1, 1).getValues();
  return values.some(row => cleanDni(row[0]) === dni);
}

function cleanEmail(value) {
  return String(value || '').trim().toLowerCase();
}

function normalizePhone(value) {
  let phone = String(value || '').trim().replace(/[^\d+]/g, '');
  if (phone.indexOf('+') > 0) phone = phone.replace(/\+/g, '');
  if (phone.startsWith('00')) phone = '+' + phone.slice(2);
  if (!phone.startsWith('+')) phone = '+' + phone;
  return phone;
}

function textForSheet(value) {
  return "'" + String(value || '');
}

function cleanText(value) {
  return String(value || '').trim().replace(/\s+/g, ' ');
}

function cleanDni(value) {
  return String(value || '').replace(/\D/g, '');
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
