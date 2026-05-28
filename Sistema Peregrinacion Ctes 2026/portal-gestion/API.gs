// ============================================================
// SISTEMA DE PAGOS PEREGRINACION CNC CORRIENTES 2026 - BACKEND
// Google Apps Script vinculado al Google Sheet de respuestas.
// ============================================================

const CONFIG = {
  HOJA_INSCRIPTOS: 'Respuestas de formulario 1',
  HOJA_PAGOS: 'Pagos',
  HOJA_CAMBIOS_PERFIL: 'Cambios Perfil',
  HOJA_REPORTES_ERROR: 'Reportes Error',
  EMAIL_ORGANIZADOR: 'TU_EMAIL@gmail.com',
  ALIAS_TRANSFERENCIA: 'alias.mercadopago',
  NOMBRE_EVENTO: 'Peregrinacion de Jovenes - CNC Corrientes 2026',
  CAPACIDAD: 100,
  ADMIN_PASSWORD: 'TU_CLAVE_ADMIN',
  ADMIN_VIEW_PASSWORD: 'TU_CLAVE_VISTA',
  PORTAL_URL: 'https://peters3o.github.io/peregrinacion-corrientes-2026/',
};

// Plan provisorio. Cuando definamos montos y vencimientos reales,
// este bloque debe coincidir con PLAN_PAGOS en index.html.
const CUOTAS = [
  { key: 'cuota1', label: '1ra cuota', monto: 35000, vencimiento: '30/06/2026' },
  { key: 'cuota2', label: '2da cuota', monto: 35000, vencimiento: '31/07/2026' },
  { key: 'cuota3', label: '3ra cuota', monto: 35000, vencimiento: '31/08/2026' },
  { key: 'cuota4', label: '4ta cuota', monto: 35000, vencimiento: '30/09/2026' },
];
const TOTAL = CUOTAS.reduce((a, c) => a + c.monto, 0);

const SYSTEM_COLUMNS = [
  'Aviso WhatsApp',
  'Fecha aviso WhatsApp',
  'Estado cupo',
  'Orden inscripcion',
  'Nota baja',
  'Reemplazo',
  'Vencimiento DNI',
];

const PAGOS_HEADERS = [
  'ID', 'Timestamp', 'DNI', 'Nombre', 'Cuotas', 'Monto', 'Comprobante',
  'Estado', 'Revisado Por', 'Fecha Revision', 'Email Enviado',
  'Motivo Rechazo', 'Nombre Archivo', 'Tipo Archivo',
  'WhatsApp Enviado', 'Fecha WhatsApp', 'Tipo WhatsApp', 'Medio Pago', 'Responsable Efectivo',
];

const CAMBIOS_PERFIL_HEADERS = [
  'ID', 'Timestamp', 'DNI', 'Nombre', 'Cambios JSON', 'Estado',
  'Revisado Por', 'Fecha Revision', 'Motivo Rechazo',
];

const REPORTES_ERROR_HEADERS = [
  'ID', 'Timestamp', 'DNI', 'Nombre', 'Email', 'Mensaje', 'Estado',
  'Revisado Por', 'Fecha Revision',
];

function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  const p = e.parameter || {};
  const action = p.action || '';
  try {
    if (action === 'getEstado') return json(getEstado(p));
    if (action === 'registrarPago') return json(registrarPago(p));
    if (action === 'solicitarCambioPerfil') return json(solicitarCambioPerfil(p));
    if (action === 'reportarError') return json(reportarError(p));
    if (action === 'loginAdmin') return json(loginAdmin(p));
    if (action === 'getAdmin') return json(getAdmin(p));
    if (action === 'confirmarPago') return json(confirmarPago(p));
    if (action === 'resolverCambioPerfil') return json(resolverCambioPerfil(p));
    if (action === 'getInscriptos') return json(getInscriptos(p));
    if (action === 'marcarBaja') return json(marcarBaja(p));
    if (action === 'liberarCupo') return json(liberarCupo(p));
    if (action === 'marcarAvisoWA') return json(marcarAvisoWA(p));
    if (action === 'marcarAvisoWAPago') return json(marcarAvisoWAPago(p));
    if (action === 'exportar') return json(exportar(p));
    return json({ ok: false, error: 'Accion desconocida.' });
  } catch (err) {
    return json({ ok: false, error: err.message });
  }
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function loginAdmin(p) {
  const pass = (p.adminPass || '').trim();
  let role = '';
  if (pass === CONFIG.ADMIN_PASSWORD) role = 'edit';
  else if (pass === CONFIG.ADMIN_VIEW_PASSWORD) role = 'view';
  else return { ok: false, error: 'Clave incorrecta.' };

  const token = Utilities.getUuid();
  CacheService.getScriptCache().put('admin_' + token, role, 21600);
  return { ok: true, token, role };
}

function obtenerRolAdmin(p) {
  const token = (p.adminToken || '').trim();
  return token ? (CacheService.getScriptCache().get('admin_' + token) || '') : '';
}

function validarAdmin(p) {
  const role = obtenerRolAdmin(p);
  return role === 'edit' || role === 'view';
}

function validarAdminEdicion(p) {
  return obtenerRolAdmin(p) === 'edit';
}

function getSheet(nombre) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(nombre);
  if (!s) s = ss.insertSheet(nombre);
  return s;
}

function getPagosSheet() {
  const s = getSheet(CONFIG.HOJA_PAGOS);
  if (s.getLastRow() === 0) {
    s.appendRow(PAGOS_HEADERS);
    s.getRange(1, 1, 1, PAGOS_HEADERS.length)
      .setFontWeight('bold')
      .setBackground('#185FA5')
      .setFontColor('white');
  } else {
    const last = Math.max(s.getLastColumn(), 1);
    const headers = s.getRange(1, 1, 1, last).getValues()[0].map(String);
    PAGOS_HEADERS.forEach(h => {
      if (!headers.some(x => normHeader(x) === normHeader(h))) {
        s.getRange(1, s.getLastColumn() + 1).setValue(h);
      }
    });
  }
  return s;
}

function getCambiosPerfilSheet() {
  const s = getSheet(CONFIG.HOJA_CAMBIOS_PERFIL);
  if (s.getLastRow() === 0) {
    s.appendRow(CAMBIOS_PERFIL_HEADERS);
    s.getRange(1, 1, 1, CAMBIOS_PERFIL_HEADERS.length)
      .setFontWeight('bold')
      .setBackground('#185FA5')
      .setFontColor('white');
  } else {
    const last = Math.max(s.getLastColumn(), 1);
    const headers = s.getRange(1, 1, 1, last).getValues()[0].map(String);
    CAMBIOS_PERFIL_HEADERS.forEach(h => {
      if (!headers.some(x => normHeader(x) === normHeader(h))) {
        s.getRange(1, s.getLastColumn() + 1).setValue(h);
      }
    });
  }
  return s;
}

function getReportesErrorSheet() {
  const s = getSheet(CONFIG.HOJA_REPORTES_ERROR);
  if (s.getLastRow() === 0) {
    s.appendRow(REPORTES_ERROR_HEADERS);
    s.getRange(1, 1, 1, REPORTES_ERROR_HEADERS.length)
      .setFontWeight('bold')
      .setBackground('#185FA5')
      .setFontColor('white');
  } else {
    const last = Math.max(s.getLastColumn(), 1);
    const headers = s.getRange(1, 1, 1, last).getValues()[0].map(String);
    REPORTES_ERROR_HEADERS.forEach(h => {
      if (!headers.some(x => normHeader(x) === normHeader(h))) {
        s.getRange(1, s.getLastColumn() + 1).setValue(h);
      }
    });
  }
  return s;
}

function normHeader(v) {
  return (v || '').toString()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .trim();
}

function normDni(d) {
  return (d || '').toString().replace(/\D/g, '').trim();
}

function normEmail(e) {
  return (e || '').toString().trim().toLowerCase();
}

function normTel(tel) {
  let n = (tel || '').toString().replace(/\D/g, '');
  if (!n) return '';
  n = n.replace(/^00/, '');
  if (n.startsWith('549') && n.length >= 12) return n;
  if (n.startsWith('54') && !n.startsWith('549')) {
    return '549' + n.slice(2).replace(/^0/, '').replace(/^15/, '');
  }
  return '549' + n.replace(/^0/, '').replace(/^15/, '');
}

function idxByTokens(headers, required, optional) {
  const req = required.map(normHeader);
  const opt = (optional || []).map(normHeader);
  for (let i = 0; i < headers.length; i++) {
    const h = normHeader(headers[i]);
    if (req.every(t => h.includes(t)) && (!opt.length || opt.some(t => h.includes(t)))) return i;
  }
  return -1;
}

function idxExact(headers, name) {
  const target = normHeader(name);
  return headers.findIndex(h => normHeader(h) === target);
}

function firstIndex(indices) {
  return indices.find(i => i >= 0) ?? -1;
}

function ensureSystemColumns(sheet) {
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  let headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  SYSTEM_COLUMNS.forEach(name => {
    if (idxExact(headers, name) < 0) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue(name);
      headers.push(name);
    }
  });
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
}

function columnasInscriptos() {
  const s = getSheet(CONFIG.HOJA_INSCRIPTOS);
  const headers = ensureSystemColumns(s);
  const cols = {
    timestamp: idxByTokens(headers, ['marca', 'temporal']),
    nombre: idxByTokens(headers, ['nombre']),
    apellido: idxByTokens(headers, ['apellido']),
    fechaNacimiento: idxByTokens(headers, ['fecha', 'nacimiento']),
    dni: idxByTokens(headers, ['dni']),
    emailPortal: idxByTokens(headers, ['email'], ['portal', 'pagos']),
    celular: idxByTokens(headers, ['celular']),
    emergencia: idxByTokens(headers, ['emergencia']),
    parroquia: idxByTokens(headers, ['parroquia']),
    comunidad: idxByTokens(headers, ['comunidad']),
    restricciones: idxByTokens(headers, ['restricciones']),
    avisoWhatsapp: idxExact(headers, 'Aviso WhatsApp'),
    fechaAvisoWhatsapp: idxExact(headers, 'Fecha aviso WhatsApp'),
    estadoCupo: idxExact(headers, 'Estado cupo'),
    ordenInscripcion: idxExact(headers, 'Orden inscripcion'),
    notaBaja: idxExact(headers, 'Nota baja'),
    reemplazo: idxExact(headers, 'Reemplazo'),
    vencimientoDni: firstIndex([
      idxExact(headers, 'Fecha de vencimiento del DNI'),
      idxByTokens(headers, ['fecha', 'vencimiento', 'dni']),
      idxExact(headers, 'Vencimiento DNI'),
    ]),
  };
  const faltan = [];
  ['nombre', 'apellido', 'dni', 'emailPortal', 'celular', 'parroquia', 'comunidad'].forEach(k => {
    if (cols[k] < 0) faltan.push(k);
  });
  if (faltan.length) throw new Error('No pude detectar columnas requeridas: ' + faltan.join(', '));
  return { sheet: s, headers, cols };
}

function inicializarCupos() {
  const info = columnasInscriptos();
  const s = info.sheet;
  const cols = info.cols;
  const data = s.getDataRange().getValues();
  let orden = 0;
  for (let i = 1; i < data.length; i++) {
    if (!normDni(data[i][cols.dni])) continue;
    orden += 1;
    if (!data[i][cols.ordenInscripcion]) {
      s.getRange(i + 1, cols.ordenInscripcion + 1).setValue(orden);
    }
    if (!data[i][cols.estadoCupo]) {
      s.getRange(i + 1, cols.estadoCupo + 1).setValue(orden <= CONFIG.CAPACIDAD ? 'con_cupo' : 'lista_espera');
    }
  }
  return true;
}

function buscarInscriptoPorDniEmail(dni, email) {
  inicializarCupos();
  const info = columnasInscriptos();
  const data = info.sheet.getDataRange().getValues();
  const dn = normDni(dni);
  const em = normEmail(email);
  let encontroDni = false;
  for (let i = 1; i < data.length; i++) {
    if (normDni(data[i][info.cols.dni]) === dn) {
      encontroDni = true;
      if (normEmail(data[i][info.cols.emailPortal]) === em) {
        return { ok: true, row: i + 1, d: data[i], cols: info.cols };
      }
    }
  }
  return { ok: false, error: encontroDni ? 'EMAIL_NO_COINCIDE' : 'DNI_NO_ENCONTRADO' };
}

function buscarInscripto(dni) {
  inicializarCupos();
  const info = columnasInscriptos();
  const data = info.sheet.getDataRange().getValues();
  const dn = normDni(dni);
  for (let i = 1; i < data.length; i++) {
    if (normDni(data[i][info.cols.dni]) === dn) return { row: i + 1, d: data[i], cols: info.cols };
  }
  return null;
}

function nombreCompleto(d, cols) {
  return ((d[cols.nombre] || '') + ' ' + (d[cols.apellido] || '')).toString().trim();
}

function fechaAR(value) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return value.toLocaleDateString('es-AR');
  }
  const text = value.toString().trim();
  const match = text.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
  if (match) {
    const y = Number(match[3].length === 2 ? '20' + match[3] : match[3]);
    return Utilities.formatDate(new Date(y, Number(match[2]) - 1, Number(match[1])), Session.getScriptTimeZone(), 'dd/MM/yyyy');
  }
  const parsed = new Date(value);
  return isNaN(parsed.getTime()) ? text : parsed.toLocaleDateString('es-AR');
}

function horaAR(value) {
  return value ? new Date(value).toLocaleTimeString('es-AR', { hour: '2-digit', minute: '2-digit' }) : '';
}

function parseFechaDDMMYYYY(v) {
  const p = (v || '').split('/');
  return new Date(Number(p[2]), Number(p[1]) - 1, Number(p[0]));
}

function cuotaVencida(c) {
  const v = parseFechaDDMMYYYY(c.vencimiento);
  v.setHours(23, 59, 59, 999);
  return new Date() > v;
}

function getPagos(dni) {
  const s = getPagosSheet();
  const dn = normDni(dni);
  return s.getDataRange().getValues().slice(1)
    .filter(r => normDni(r[2]) === dn)
    .map(r => ({
      id: r[0],
      cuotas: r[4] || '',
      monto: Number(r[5] || 0),
      comprobanteUrl: r[6] || '',
      estado: r[7] || '',
      fecha: fechaAR(r[1]),
      motivoRechazo: r[11] || '',
      comprobanteNombre: r[12] || '',
      whatsappEnviado: r[14] || '',
      fechaWhatsapp: fechaAR(r[15]),
      tipoWhatsapp: r[16] || '',
      medioPago: r[17] || '',
      responsableEfectivo: r[18] || '',
    }));
}

function resumenPagos(pagos) {
  const keysUnicas = lista => Array.from(new Set(
    lista.flatMap(p => (p.cuotas || '').split(',').map(c => c.trim()).filter(Boolean))
  ));
  const confirmadas = keysUnicas(pagos.filter(p => p.estado === 'confirmado'));
  const pendientes = keysUnicas(pagos.filter(p => p.estado === 'pendiente'));
  const rechazadas = keysUnicas(pagos.filter(p => p.estado === 'rechazado'));
  const totalPagado = pagos.filter(p => p.estado === 'confirmado').reduce((a, p) => a + Number(p.monto || 0), 0);
  const totalPendConf = pagos.filter(p => p.estado === 'pendiente').reduce((a, p) => a + Number(p.monto || 0), 0);
  return { confirmadas, pendientes, rechazadas, totalPagado, totalPendConf };
}

function estadoInscripto(d, cols, pagos) {
  const r = resumenPagos(pagos);
  const estadoCupo = (d[cols.estadoCupo] || 'con_cupo').toString().trim() || 'con_cupo';
  if (estadoCupo === 'lista_espera') return 'lista_espera';
  if (estadoCupo === 'baja') return 'baja';
  if (r.confirmadas.length === CUOTAS.length) return 'pago_completo';
  if (r.rechazadas.length > 0 && r.pendientes.length === 0) return 'comprobante_rechazado';
  if (r.pendientes.length > 0) return 'en_revision';

  const hoy = new Date();
  hoy.setHours(0, 0, 0, 0);
  const vencidasSinPago = CUOTAS.filter(c => {
    return cuotaVencida(c) && !r.confirmadas.includes(c.key) && !r.pendientes.includes(c.key);
  });
  if (vencidasSinPago.length > 0) return 'debe_vencidas';
  if (r.confirmadas.length > 0) return 'al_dia';
  return 'sin_pagos';
}

function getEstado(p) {
  const v = buscarInscriptoPorDniEmail(p.dni, p.email);
  if (!v.ok) {
    return {
      ok: false,
      error: v.error === 'EMAIL_NO_COINCIDE'
        ? 'El email no coincide con el usado en la inscripcion.'
        : 'No encontramos una inscripcion con ese DNI.',
    };
  }

  const d = v.d;
  const cols = v.cols;
  const pagos = getPagos(d[cols.dni]);
  const resumen = resumenPagos(pagos);
  return {
    ok: true,
    nombre: nombreCompleto(d, cols),
    dni: d[cols.dni],
    parroquia: d[cols.parroquia] || '',
    comunidad: d[cols.comunidad] || '',
    celular: d[cols.celular] || '',
    emergencia: d[cols.emergencia] || '',
    restricciones: d[cols.restricciones] || '',
    vencimientoDni: fechaAR(d[cols.vencimientoDni]),
    estadoCupo: (d[cols.estadoCupo] || 'con_cupo').toString().trim() || 'con_cupo',
    pagos,
    resumen,
    cuotas: CUOTAS,
    total: TOTAL,
  };
}

function normalizarCambiosPerfil(p) {
  const cambios = {};
  const allowed = {
    celular: 'Celular',
    emergencia: 'Telefono emergencia',
    parroquia: 'Parroquia',
    comunidad: 'Comunidad',
    restricciones: 'Restricciones alimentarias',
  };
  Object.keys(allowed).forEach(key => {
    if (p[key] !== undefined) {
      const value = (p[key] || '').toString().trim();
      if (value) cambios[key] = value;
    }
  });
  return cambios;
}

function solicitarCambioPerfil(p) {
  const v = buscarInscriptoPorDniEmail(p.dni, p.email);
  if (!v.ok) return { ok: false, error: 'Datos no validos.' };
  const cambios = normalizarCambiosPerfil(p);
  if (!Object.keys(cambios).length) return { ok: false, error: 'No indicaste ningun dato para actualizar.' };
  const id = 'PER-' + Date.now();
  getCambiosPerfilSheet().appendRow([
    id, new Date(), v.d[v.cols.dni], nombreCompleto(v.d, v.cols),
    JSON.stringify(cambios), 'pendiente', '', '', '',
  ]);
  return { ok: true, mensaje: 'Tus cambios quedaron pendientes de aprobacion por un catequista.' };
}

function reportarError(p) {
  const v = buscarInscriptoPorDniEmail(p.dni, p.email);
  if (!v.ok) return { ok: false, error: 'Datos no validos.' };
  const mensaje = (p.mensaje || '').toString().trim();
  if (!mensaje) return { ok: false, error: 'Escribi que problema encontraste.' };
  const id = 'ERR-' + Date.now();
  const nombre = nombreCompleto(v.d, v.cols);
  const email = v.d[v.cols.emailPortal] || '';
  getReportesErrorSheet().appendRow([
    id, new Date(), v.d[v.cols.dni], nombre, email, mensaje, 'pendiente', '', '',
  ]);
  MailApp.sendEmail({
    to: CONFIG.EMAIL_ORGANIZADOR,
    subject: 'Reporte de error del portal - ' + nombre,
    body: 'Persona: ' + nombre + '\nDNI: ' + v.d[v.cols.dni] + '\nEmail: ' + email + '\n\nMensaje:\n' + mensaje,
  });
  return { ok: true, mensaje: 'Gracias. Enviamos tu reporte a los catequistas.' };
}

function cambiosPerfilPendientes() {
  const s = getCambiosPerfilSheet();
  return s.getDataRange().getValues().slice(1)
    .filter(r => (r[5] || '').toString() === 'pendiente')
    .map(r => ({
      id: r[0],
      fecha: fechaAR(r[1]),
      dni: r[2],
      nombre: r[3],
      cambios: parseJsonSafe(r[4]),
      estado: r[5],
    }));
}

function parseJsonSafe(txt) {
  try { return JSON.parse(txt || '{}'); } catch (e) { return {}; }
}

function resolverCambioPerfil(p) {
  if (!validarAdminEdicion(p)) return { ok: false, error: 'No autorizado.' };
  const id = (p.cambioId || '').toString().trim();
  const decision = (p.decision || '').toString().trim();
  if (!id || !['aprobado', 'rechazado'].includes(decision)) return { ok: false, error: 'Decision invalida.' };
  const s = getCambiosPerfilSheet();
  const rows = s.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if ((rows[i][0] || '').toString() === id) {
      if ((rows[i][5] || '').toString() !== 'pendiente') return { ok: false, error: 'Este cambio ya fue revisado.' };
      const dni = rows[i][2];
      const cambios = parseJsonSafe(rows[i][4]);
      if (decision === 'aprobado') aplicarCambiosPerfil(dni, cambios);
      s.getRange(i + 1, 6).setValue(decision);
      s.getRange(i + 1, 7).setValue('admin');
      s.getRange(i + 1, 8).setValue(new Date());
      s.getRange(i + 1, 9).setValue(p.motivo || '');
      return { ok: true };
    }
  }
  return { ok: false, error: 'Solicitud no encontrada.' };
}

function aplicarCambiosPerfil(dni, cambios) {
  const ins = buscarInscripto(dni);
  if (!ins) throw new Error('Inscripto no encontrado.');
  const s = getSheet(CONFIG.HOJA_INSCRIPTOS);
  const map = {
    celular: ins.cols.celular,
    emergencia: ins.cols.emergencia,
    parroquia: ins.cols.parroquia,
    comunidad: ins.cols.comunidad,
    restricciones: ins.cols.restricciones,
  };
  Object.keys(map).forEach(key => {
    if (cambios[key] !== undefined && map[key] >= 0) {
      s.getRange(ins.row, map[key] + 1).setValue(cambios[key]);
    }
  });
}

function registrarPago(p) {
  const v = buscarInscriptoPorDniEmail(p.dni, p.email);
  if (!v.ok) return { ok: false, error: 'Datos no validos.' };
  const d = v.d;
  const cols = v.cols;
  const estadoCupo = (d[cols.estadoCupo] || 'con_cupo').toString().trim() || 'con_cupo';
  if (estadoCupo === 'lista_espera') return { ok: false, error: 'Estas en lista de espera. Todavia no cargues pagos.' };
  if (estadoCupo === 'baja') return { ok: false, error: 'Tu inscripcion fue dada de baja.' };

  const cuotasKeys = (p.cuotas || '').split(',').map(c => c.trim()).filter(Boolean);
  if (!cuotasKeys.length) return { ok: false, error: 'Selecciona al menos una cuota.' };
  if (!p.comprobanteUrl) return { ok: false, error: 'No se recibio el comprobante.' };

  const keysValidas = CUOTAS.map(c => c.key);
  const invalidas = cuotasKeys.filter(k => !keysValidas.includes(k));
  if (invalidas.length) return { ok: false, error: 'Cuotas invalidas: ' + invalidas.join(', ') };

  const existentes = getPagos(d[cols.dni]);
  const previasFaltantes = cuotasPreviasFaltantes(cuotasKeys, existentes);
  if (previasFaltantes.length) {
    return { ok: false, error: 'Para cargar esa cuota, primero tenes que cargar el comprobante de: ' + previasFaltantes.join(', ') };
  }
  for (const key of cuotasKeys) {
    const ya = existentes.find(x => (x.cuotas || '').split(',').map(c => c.trim()).includes(key) && x.estado !== 'rechazado');
    if (ya) return { ok: false, error: 'La cuota "' + key + '" ya tiene un comprobante en estado: ' + ya.estado };
  }

  const cuotaObjs = CUOTAS.filter(c => cuotasKeys.includes(c.key));
  const monto = cuotaObjs.reduce((a, c) => a + c.monto, 0);
  const label = cuotaObjs.map(c => c.label).join(' + ');
  const nombre = nombreCompleto(d, cols);
  const id = 'PAG-' + Date.now();
  const medioPago = (p.medioPago || '').toString().trim();
  const responsableEfectivo = (p.responsableEfectivo || '').toString().trim();
  if (medioPago.toLowerCase().includes('efectivo') && !responsableEfectivo) {
    return { ok: false, error: 'Indica nombre y apellido del responsable que recibio el efectivo.' };
  }
  getPagosSheet().appendRow([
    id, new Date(), d[cols.dni], nombre, cuotasKeys.join(','), monto, p.comprobanteUrl,
    'pendiente', '', '', 'no', '', p.comprobanteNombre || '', p.comprobanteTipo || '',
    'no', '', '', medioPago, responsableEfectivo,
  ]);

  MailApp.sendEmail({
    to: CONFIG.EMAIL_ORGANIZADOR,
    subject: 'Nuevo comprobante: ' + nombre + ' - ' + label,
    body: 'Persona: ' + nombre + '\nDNI: ' + d[cols.dni] + '\nCuotas: ' + label +
      ' ($' + monto + ')\nMedio: ' + (medioPago || 'No indicado') +
      (responsableEfectivo ? '\nResponsable efectivo: ' + responsableEfectivo : '') +
      '\nID: ' + id + '\nComprobante: ' + p.comprobanteUrl,
  });
  return { ok: true, mensaje: 'Tu comprobante quedo registrado. Cuando se revise, vas a ver el estado actualizado.' };
}

function cuotasPreviasFaltantes(cuotasKeys, existentes) {
  const activas = resumenPagos(existentes.filter(p => p.estado !== 'rechazado'));
  const cargadas = Array.from(new Set([].concat(activas.confirmadas, activas.pendientes)));
  const faltan = [];
  cuotasKeys.forEach(key => {
    const idx = CUOTAS.findIndex(c => c.key === key);
    for (let i = 0; i < idx; i++) {
      if (!cargadas.includes(CUOTAS[i].key) && !cuotasKeys.includes(CUOTAS[i].key)) {
        faltan.push(CUOTAS[i].label);
      }
    }
  });
  return Array.from(new Set(faltan));
}

function confirmarPago(p) {
  if (!validarAdminEdicion(p)) return { ok: false, error: 'No autorizado.' };
  const nuevoEstado = (p.nuevoEstado || '').trim();
  if (!['confirmado', 'rechazado', 'pendiente'].includes(nuevoEstado)) return { ok: false, error: 'Estado invalido.' };

  const sheet = getPagosSheet();
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if ((rows[i][0] || '').toString() === (p.pagoId || '').toString()) {
      sheet.getRange(i + 1, 8).setValue(nuevoEstado);
      sheet.getRange(i + 1, 9).setValue('admin');
      sheet.getRange(i + 1, 10).setValue(new Date());
      if (nuevoEstado === 'rechazado') sheet.getRange(i + 1, 12).setValue(p.motivo || '');
      if (nuevoEstado === 'pendiente') {
        sheet.getRange(i + 1, 11).setValue('no');
        sheet.getRange(i + 1, 12).setValue('');
        sheet.getRange(i + 1, 15).setValue('no');
        sheet.getRange(i + 1, 16).setValue('');
        sheet.getRange(i + 1, 17).setValue('');
      }
      if (nuevoEstado === 'confirmado' && rows[i][10] !== 'si') {
        const ins = buscarInscripto(rows[i][2]);
        if (ins) {
          enviarEmailConfirmacion(ins.d, ins.cols, getPagos(rows[i][2]), rows[i][4]);
          sheet.getRange(i + 1, 11).setValue('si');
        }
      }
      return { ok: true };
    }
  }
  return { ok: false, error: 'Pago no encontrado.' };
}

function contactoPorDni() {
  inicializarCupos();
  const info = columnasInscriptos();
  const data = info.sheet.getDataRange().getValues();
  const map = {};
  data.slice(1).forEach(d => {
    const dni = normDni(d[info.cols.dni]);
    if (!dni) return;
    map[dni] = {
      nombre: nombreCompleto(d, info.cols),
      email: d[info.cols.emailPortal] || '',
      celular: d[info.cols.celular] || '',
      whatsapp: normTel(d[info.cols.celular]),
      fechaNacimiento: fechaAR(d[info.cols.fechaNacimiento]),
      parroquia: d[info.cols.parroquia] || '',
      comunidad: d[info.cols.comunidad] || '',
      vencimientoDni: fechaAR(d[info.cols.vencimientoDni]),
      fechaAvisoWhatsapp: fechaAR(d[info.cols.fechaAvisoWhatsapp]),
      estadoCupo: (d[info.cols.estadoCupo] || 'con_cupo').toString().trim() || 'con_cupo',
    };
  });
  return map;
}

function getAdmin(p) {
  if (!validarAdmin(p)) return { ok: false, error: 'No autorizado.' };
  const contactos = contactoPorDni();
  const pagos = getPagosSheet().getDataRange().getValues().slice(1).map(r => {
    const dni = normDni(r[2]);
    const c = contactos[dni] || {};
    return {
      id: r[0],
      dni: dni || r[2],
      nombre: r[3] || c.nombre || '',
      cuotas: r[4] || '',
      monto: Number(r[5] || 0),
      comprobanteUrl: r[6] || '',
      estado: r[7] || '',
      fecha: fechaAR(r[1]),
      motivoRechazo: r[11] || '',
      comprobanteNombre: r[12] || '',
      whatsappEnviado: r[14] || '',
      fechaWhatsapp: fechaAR(r[15]),
      tipoWhatsapp: r[16] || '',
      medioPago: r[17] || '',
      responsableEfectivo: r[18] || '',
      email: c.email || '',
      celular: c.celular || '',
      whatsapp: c.whatsapp || '',
      fechaNacimiento: c.fechaNacimiento || '',
      parroquia: c.parroquia || '',
      comunidad: c.comunidad || '',
      fechaAvisoWhatsapp: c.fechaAvisoWhatsapp || '',
      estadoCupo: c.estadoCupo || '',
      vencimientoDni: c.vencimientoDni || '',
    };
  });
  const totalInscriptos = Object.keys(contactos).length;
  const totalActivos = Object.values(contactos).filter(c => c.estadoCupo !== 'baja').length;
  const totalBajas = Object.values(contactos).filter(c => c.estadoCupo === 'baja').length;
  const dineroADevolver = Object.keys(contactos).reduce((total, dni) => {
    if (contactos[dni].estadoCupo !== 'baja') return total;
    return total + pagos
      .filter(p => normDni(p.dni) === dni && p.estado === 'confirmado')
      .reduce((a, p) => a + Number(p.monto || 0), 0);
  }, 0);
  return { ok: true, totalInscriptos, totalActivos, totalBajas, dineroADevolver, pagos, cambiosPerfil: cambiosPerfilPendientes() };
}

function getInscriptos(p) {
  if (!validarAdmin(p)) return { ok: false, error: 'No autorizado.' };
  inicializarCupos();
  const info = columnasInscriptos();
  const data = info.sheet.getDataRange().getValues();
  const pagosData = getPagosSheet().getDataRange().getValues().slice(1);
  const inscriptos = data.slice(1)
    .filter(d => normDni(d[info.cols.dni]))
    .map((d, idx) => {
      const dni = d[info.cols.dni];
      const pagosPersona = pagosData.filter(r => normDni(r[2]) === normDni(dni)).map(r => ({
        id: r[0],
        cuotas: r[4] || '',
        monto: Number(r[5] || 0),
        comprobanteUrl: r[6] || '',
        estado: r[7] || '',
        fecha: fechaAR(r[1]),
        medioPago: r[17] || '',
      }));
      const resumen = resumenPagos(pagosPersona);
      const estadoCupo = (d[info.cols.estadoCupo] || 'con_cupo').toString().trim() || 'con_cupo';
      return {
        orden: d[info.cols.ordenInscripcion] || (idx + 1),
        nombre: nombreCompleto(d, info.cols),
        dni,
        email: d[info.cols.emailPortal] || '',
        celular: d[info.cols.celular] || '',
        emergencia: d[info.cols.emergencia] || '',
        fechaNacimiento: fechaAR(d[info.cols.fechaNacimiento]),
        parroquia: d[info.cols.parroquia] || '',
        comunidad: d[info.cols.comunidad] || '',
        restricciones: d[info.cols.restricciones] || '',
        vencimientoDni: fechaAR(d[info.cols.vencimientoDni]),
        whatsapp: normTel(d[info.cols.celular]),
        fechaInscripcion: fechaAR(d[info.cols.timestamp]),
        horaInscripcion: horaAR(d[info.cols.timestamp]),
        estadoCupo,
        estado: estadoInscripto(d, info.cols, pagosPersona),
        fechaAvisoSena: fechaAR(d[info.cols.fechaAvisoWhatsapp]),
        notaBaja: d[info.cols.notaBaja] || '',
        reemplazo: d[info.cols.reemplazo] || '',
        resumen,
        pagos: pagosPersona,
      };
    });
  return { ok: true, inscriptos };
}

function marcarBaja(p) {
  if (!validarAdminEdicion(p)) return { ok: false, error: 'No autorizado.' };
  const ins = buscarInscripto(p.dni);
  if (!ins) return { ok: false, error: 'Inscripto no encontrado.' };
  const s = getSheet(CONFIG.HOJA_INSCRIPTOS);
  s.getRange(ins.row, ins.cols.estadoCupo + 1).setValue('baja');
  s.getRange(ins.row, ins.cols.notaBaja + 1).setValue(p.nota || '');
  s.getRange(ins.row, ins.cols.reemplazo + 1).setValue(p.reemplazo || '');
  return { ok: true, dni: normDni(p.dni), nombre: nombreCompleto(ins.d, ins.cols), estadoCupo: 'baja' };
}

function liberarCupo(p) {
  if (!validarAdminEdicion(p)) return { ok: false, error: 'No autorizado.' };
  inicializarCupos();
  const info = columnasInscriptos();
  const data = info.sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (normDni(data[i][info.cols.dni]) === normDni(p.dni)) {
      info.sheet.getRange(i + 1, info.cols.estadoCupo + 1).setValue('baja');
      break;
    }
  }
  for (let i = 1; i < data.length; i++) {
    if ((data[i][info.cols.estadoCupo] || '').toString().trim() === 'lista_espera') {
      info.sheet.getRange(i + 1, info.cols.estadoCupo + 1).setValue('cupo_liberado');
      const nombre = nombreCompleto(data[i], info.cols);
      MailApp.sendEmail({
        to: CONFIG.EMAIL_ORGANIZADOR,
        subject: 'Nuevo cupo disponible: ' + nombre,
        body: nombre + ' (DNI ' + data[i][info.cols.dni] + ') paso de lista de espera a cupo disponible.',
      });
      return { ok: true, nuevoCupo: nombre };
    }
  }
  return { ok: true, nuevoCupo: null };
}

function marcarAvisoWA(p) {
  if (!validarAdminEdicion(p)) return { ok: false, error: 'No autorizado.' };
  const ins = buscarInscripto(p.dni);
  if (!ins) return { ok: false, error: 'No encontrado.' };
  const s = getSheet(CONFIG.HOJA_INSCRIPTOS);
  s.getRange(ins.row, ins.cols.avisoWhatsapp + 1).setValue('enviado');
  s.getRange(ins.row, ins.cols.fechaAvisoWhatsapp + 1).setValue(new Date());
  return { ok: true };
}

function marcarAvisoWAPago(p) {
  if (!validarAdminEdicion(p)) return { ok: false, error: 'No autorizado.' };
  const id = (p.pagoId || '').toString().trim();
  if (!id) return { ok: false, error: 'Pago no indicado.' };
  const tipo = (p.tipo || '').toString().trim() || 'recordatorio';
  const s = getPagosSheet();
  const rows = s.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if ((rows[i][0] || '').toString().trim() === id) {
      s.getRange(i + 1, 15).setValue('si');
      s.getRange(i + 1, 16).setValue(new Date());
      s.getRange(i + 1, 17).setValue(tipo);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Pago no encontrado.' };
}

function exportar(p) {
  if (!validarAdmin(p)) return { ok: false, error: 'No autorizado.' };
  const result = getInscriptos(p);
  if (!result.ok) return result;
  return {
    ok: true,
    data: result.inscriptos.map(ins => ({
      Orden: ins.orden,
      Nombre: ins.nombre,
      DNI: ins.dni,
      Email: ins.email,
      Celular: ins.celular,
      Parroquia: ins.parroquia,
      Comunidad: ins.comunidad,
      Restricciones: ins.restricciones,
      'Fecha inscripcion': ins.fechaInscripcion + ' ' + ins.horaInscripcion,
      'Estado cupo': ins.estadoCupo,
      'Estado pagos': ins.estado,
      'Total confirmado': ins.resumen.totalPagado,
      'En revision': ins.resumen.totalPendConf,
      'Falta abonar': TOTAL - ins.resumen.totalPagado - ins.resumen.totalPendConf,
    })),
  };
}

function escapeHtmlEmail(txt) {
  return (txt || '').toString()
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function estadoCuotaEmail(c, resumen) {
  if (resumen.confirmadas.includes(c.key)) return 'Confirmada';
  if (resumen.pendientes.includes(c.key)) return 'En revision';
  if (resumen.rechazadas.includes(c.key)) return 'Comprobante no validado';
  return 'Pendiente. Vence el ' + c.vencimiento;
}

function lineasEstadoCuotasTexto(resumen) {
  return CUOTAS.map(c => c.label + ' - $' + c.monto + ' - ' + estadoCuotaEmail(c, resumen)).join('\n');
}

function tablaEstadoCuotasHtml(resumen) {
  return CUOTAS.map(c => {
    const estado = estadoCuotaEmail(c, resumen);
    const bg = estado === 'Confirmada' ? '#e8f5ef' : estado === 'En revision' ? '#fef3dc' : estado === 'Comprobante no validado' ? '#fdecea' : '#f7f8f7';
    const color = estado === 'Confirmada' ? '#1a7a56' : estado === 'En revision' ? '#d4820a' : estado === 'Comprobante no validado' ? '#b83232' : '#5a635a';
    return '<tr><td style="padding:10px 12px;border-bottom:1px solid #efefef;font-weight:600">' +
      escapeHtmlEmail(c.label) + '</td><td style="padding:10px 12px;border-bottom:1px solid #efefef">$' +
      c.monto + '</td><td style="padding:10px 12px;border-bottom:1px solid #efefef"><span style="background:' +
      bg + ';color:' + color + ';border-radius:999px;padding:4px 10px;font-size:12px;font-weight:700">' +
      escapeHtmlEmail(estado) + '</span></td></tr>';
  }).join('');
}

function enviarEmailConfirmacion(d, cols, pagos, cuotasStr) {
  const email = d[cols.emailPortal];
  if (!email) return;
  const resumen = resumenPagos(pagos);
  const nombre = (d[cols.nombre] || '').toString().trim();
  const cuotasKeys = (cuotasStr || '').split(',').map(c => c.trim()).filter(Boolean);
  const cuotasLabel = CUOTAS.filter(c => cuotasKeys.includes(c.key)).map(c => c.label).join(' y ');
  const falta = Math.max(0, TOTAL - resumen.totalPagado);
  const body = 'Hola, ' + nombre + '.\n\nConfirmamos el pago de tu ' + cuotasLabel +
    ' para ' + CONFIG.NOMBRE_EVENTO + '.\n\nEstado de cuotas:\n\n' +
    lineasEstadoCuotasTexto(resumen) + '\n\nTotal confirmado: $' + resumen.totalPagado +
    '\nFalta abonar: $' + falta + '\n\nPortal de pagos:\n' + CONFIG.PORTAL_URL;
  const html = '<div style="margin:0;padding:24px;background:#f7f8f7;font-family:Arial,sans-serif"><div style="max-width:620px;margin:0 auto;background:#fff;border:1px solid #efefef;border-radius:16px;overflow:hidden"><div style="padding:22px 24px;border-bottom:1px solid #efefef"><div style="font-size:12px;text-transform:uppercase;color:#1a7a56;font-weight:700;margin-bottom:8px">Pago confirmado</div><h1 style="font-size:22px;margin:0">' +
    escapeHtmlEmail(CONFIG.NOMBRE_EVENTO) + '</h1></div><div style="padding:22px 24px"><p>Hola, ' +
    escapeHtmlEmail(nombre) + '.</p><p>Confirmamos el pago de tu <strong>' +
    escapeHtmlEmail(cuotasLabel) + '</strong>.</p><table style="width:100%;border-collapse:collapse;margin:0 0 20px;font-size:14px"><tbody>' +
    tablaEstadoCuotasHtml(resumen) + '</tbody></table><p><strong>Total confirmado:</strong> $' +
    resumen.totalPagado + '<br><strong>Falta abonar:</strong> $' + falta +
    '</p><p>Portal de pagos: <a href="' + CONFIG.PORTAL_URL + '">' + CONFIG.PORTAL_URL + '</a></p></div></div></div>';
  MailApp.sendEmail({ to: email, subject: 'Pago confirmado - ' + CONFIG.NOMBRE_EVENTO, body, htmlBody: html });
}

function valorNamed(resp, nombres) {
  for (const nombre of nombres) {
    if (resp[nombre] && resp[nombre][0]) return resp[nombre][0];
  }
  const normalizedTargets = nombres.map(normHeader);
  for (const key in resp) {
    const nk = normHeader(key);
    if (normalizedTargets.some(t => nk.includes(t)) && resp[key] && resp[key][0]) return resp[key][0];
  }
  return '';
}

function enviarEmailInscripcion(email, nombre) {
  if (!email) return;
  const body = 'Hola, ' + nombre + '.\n\nRecibimos tu inscripcion para ' + CONFIG.NOMBRE_EVENTO +
    '.\n\nPara consultar cuotas y subir comprobantes, ingresa al portal con tu DNI y el email que cargaste en el formulario.\n\nPortal:\n' +
    CONFIG.PORTAL_URL + '\n\nPlan de pagos:\n' +
    CUOTAS.map(c => '- ' + c.label + ': $' + c.monto + ' - vence ' + c.vencimiento).join('\n') +
    '\n\nTotal: $' + TOTAL;
  MailApp.sendEmail({ to: email, subject: 'Inscripcion recibida - ' + CONFIG.NOMBRE_EVENTO, body });
}

function onFormSubmit(e) {
  try {
    const resp = e.namedValues || {};
    const email = valorNamed(resp, [
      'Email que vas a usar para el portal de pagos*',
      'Email que vas a usar para el portal de pagos',
      'Email',
      'Direccion de correo electronico',
    ]);
    const nombre = valorNamed(resp, ['Nombre/s', 'Nombre']);
    enviarEmailInscripcion(email, nombre);
    inicializarCupos();
  } catch (err) {
    Logger.log('Error onFormSubmit: ' + err.message);
  }
}
