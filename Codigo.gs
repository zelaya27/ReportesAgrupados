/******************************************************
 * APP REPORTES AGRUPADOS UTCD
 * Backend Google Apps Script
 * Google Sheet: Reportes_Agrupados
 * Hojas: usuarios, reportepadre, agrupados
 ******************************************************/

const CONFIG_APP = {
  SPREADSHEET_ID: '1Ur66Z6adcsAFult1pixdAGbp_Mnofk9HtFb4wubv53c',
  HOJA_USUARIOS: 'usuarios',
  HOJA_PADRES: 'reportepadre',
  HOJA_AGRUPADOS: 'agrupados',
  ESTADO_PENDIENTE: 'PENDIENTE',
  ESTADO_EJECUTADO: 'EJECUTADO'
};

function doGet(e) {
  try {
    const p = e.parameter || {};
    const action = p.action || '';
    let result;

    if (action === 'login') result = loginUsuario_(p.usuario, p.password);
    else if (action === 'obtenerReportes') result = obtenerReportes_(p.usuario, p.sector, p.tipoUsuario);
    else if (action === 'obtenerReporte') result = obtenerReporte_(p.id);
    else if (action === 'guardarReporte') result = guardarReporte_(parsePayload_(p.payload));
    else if (action === 'cambiarEstado') result = cambiarEstado_(p.id, p.estado, p.usuario, p.tipoUsuario);
    else result = { ok: false, error: 'Acción no reconocida: ' + action };

    return responder_(result, p.callback);
  } catch (err) {
    return responder_({ ok: false, error: err.message || String(err) }, (e.parameter || {}).callback);
  }
}

function doPost(e) {
  return doGet(e);
}

function parsePayload_(payload) {
  if (!payload) return {};
  try {
    return JSON.parse(payload);
  } catch (err) {
    throw new Error('Payload inválido.');
  }
}

function responder_(obj, callback) {
  const json = JSON.stringify(obj);
  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function ss_() {
  return SpreadsheetApp.openById(CONFIG_APP.SPREADSHEET_ID);
}

function hoja_(nombre) {
  const sh = ss_().getSheetByName(nombre);
  if (!sh) throw new Error('No existe la hoja: ' + nombre);
  return sh;
}

function normalizar_(v) {
  return String(v == null ? '' : v).trim();
}

function upper_(v) {
  return normalizar_(v).toUpperCase();
}

function fechaDDMMYYYY_(v) {
  if (!v) return '';
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v)) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  }
  const s = normalizar_(v);
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
    const [y, m, d] = s.split('-');
    return `${d}/${m}/${y}`;
  }
  return s;
}

function validar7Digitos_(valor, nombreCampo) {
  const s = normalizar_(valor);
  if (!/^\d{7}$/.test(s)) throw new Error(nombreCampo + ' debe ser un número entero de 7 dígitos.');
  return s;
}

function validarGrupo_(padre, asociados) {
  const padreStr = validar7Digitos_(padre, 'Reporte padre');
  if (!Array.isArray(asociados) || asociados.length === 0) {
    throw new Error('Debe ingresar al menos un reporte asociado.');
  }

  const limpios = [];
  const repetidos = {};
  asociados.forEach((a, i) => {
    const asociado = validar7Digitos_(a, 'Reporte asociado #' + (i + 1));
    if (Number(asociado) < Number(padreStr)) {
      throw new Error('El Reporte Padre debe ser el número menor. No puede ser mayor que un reporte asociado.');
    }
    if (repetidos[asociado]) throw new Error('Hay un reporte asociado repetido: ' + asociado);
    repetidos[asociado] = true;
    limpios.push(asociado);
  });
  return { padre: padreStr, asociados: limpios };
}

function loginUsuario_(usuario, password) {
  usuario = upper_(usuario);
  password = normalizar_(password);
  if (!usuario || !password) throw new Error('Ingrese usuario y contraseña.');

  const sh = hoja_(CONFIG_APP.HOJA_USUARIOS);
  const data = sh.getDataRange().getValues();
  if (data.length < 2) throw new Error('No hay usuarios registrados.');

  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    const u = upper_(r[0]);
    const pass = normalizar_(r[1]);
    if (u === usuario && pass === password) {
      return {
        ok: true,
        usuario: u,
        tipoUsuario: normalizar_(r[3]) || '3',
        tipoNombre: normalizar_(r[2]),
        sector: upper_(r[4]),
        cuadrilla: normalizar_(r[5])
      };
    }
  }
  throw new Error('Usuario o contraseña incorrectos.');
}

function obtenerReportes_(usuario, sector, tipoUsuario) {
  usuario = upper_(usuario);
  sector = upper_(sector);
  tipoUsuario = normalizar_(tipoUsuario || '3');

  const shPadres = hoja_(CONFIG_APP.HOJA_PADRES);
  const shAgr = hoja_(CONFIG_APP.HOJA_AGRUPADOS);
  const padres = shPadres.getDataRange().getValues();
  const agr = shAgr.getDataRange().getValues();

  const conteos = {};
  for (let i = 1; i < agr.length; i++) {
    const padre = normalizar_(agr[i][0]);
    const asociado = normalizar_(agr[i][1]);
    if (!padre || !asociado) continue;
    conteos[padre] = (conteos[padre] || 0) + 1;
  }

  const lista = [];
  for (let i = 1; i < padres.length; i++) {
    const r = padres[i];
    const id = normalizar_(r[0]);
    const reportePadre = normalizar_(r[2]);
    const userRow = upper_(r[3]);
    const sectorRow = upper_(r[4]);
    if (!id || !reportePadre) continue;

    if (tipoUsuario === '3') {
      if (sectorRow !== sector || userRow !== usuario) continue;
    } else {
      if (sectorRow !== sector && sector !== 'GENERAL') continue;
    }

    lista.push({
      id: id,
      fecha: fechaDDMMYYYY_(r[1]),
      reportePadre: reportePadre,
      usuario: userRow,
      sector: sectorRow,
      estado: upper_(r[5]) || CONFIG_APP.ESTADO_PENDIENTE,
      asociados: conteos[reportePadre] || 0
    });
  }

  lista.reverse();
  return { ok: true, reportes: lista };
}

function obtenerReporte_(id) {
  id = normalizar_(id);
  if (!id) throw new Error('ID requerido.');

  const shPadres = hoja_(CONFIG_APP.HOJA_PADRES);
  const data = shPadres.getDataRange().getValues();
  let rowIndex = -1;
  let row = null;

  for (let i = 1; i < data.length; i++) {
    if (normalizar_(data[i][0]) === id) {
      rowIndex = i + 1;
      row = data[i];
      break;
    }
  }
  if (!row) throw new Error('No se encontró el reporte.');

  const padre = normalizar_(row[2]);
  const shAgr = hoja_(CONFIG_APP.HOJA_AGRUPADOS);
  const agr = shAgr.getDataRange().getValues();
  const asociados = [];
  for (let j = 1; j < agr.length; j++) {
    if (normalizar_(agr[j][0]) === padre && normalizar_(agr[j][1])) {
      asociados.push(normalizar_(agr[j][1]));
    }
  }

  return {
    ok: true,
    reporte: {
      id: normalizar_(row[0]),
      fecha: fechaDDMMYYYY_(row[1]),
      reportePadre: padre,
      usuario: upper_(row[3]),
      sector: upper_(row[4]),
      estado: upper_(row[5]) || CONFIG_APP.ESTADO_PENDIENTE,
      asociados: asociados,
      rowIndex: rowIndex
    }
  };
}

function guardarReporte_(payload) {
  payload = payload || {};
  const modo = normalizar_(payload.modo || 'nuevo');
  const id = normalizar_(payload.id);
  const usuario = upper_(payload.usuario);
  const sector = upper_(payload.sector);
  const fecha = fechaDDMMYYYY_(payload.fecha);
  const validado = validarGrupo_(payload.reportePadre, payload.asociados || []);

  if (!usuario || !sector) throw new Error('Sesión inválida. Vuelva a ingresar.');
  if (!fecha) throw new Error('Ingrese la fecha.');

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);

  try {
    const shPadres = hoja_(CONFIG_APP.HOJA_PADRES);
    const shAgr = hoja_(CONFIG_APP.HOJA_AGRUPADOS);
    asegurarEncabezados_();

    const padres = shPadres.getDataRange().getValues();

    if (modo === 'editar') {
      if (!id) throw new Error('ID requerido para editar.');
      let rowIndex = -1;
      let oldPadre = '';
      let estadoActual = CONFIG_APP.ESTADO_PENDIENTE;

      for (let i = 1; i < padres.length; i++) {
        if (normalizar_(padres[i][0]) === id) {
          rowIndex = i + 1;
          oldPadre = normalizar_(padres[i][2]);
          estadoActual = upper_(padres[i][5]) || CONFIG_APP.ESTADO_PENDIENTE;
          break;
        }
      }
      if (rowIndex === -1) throw new Error('No se encontró el registro a editar.');
      if (estadoActual === CONFIG_APP.ESTADO_EJECUTADO) throw new Error('Este registro está EJECUTADO y no se puede editar.');

      validarPadreUnico_(validado.padre, id);
      shPadres.getRange(rowIndex, 1, 1, 8).setValues([[
        id,
        fecha,
        validado.padre,
        usuario,
        sector,
        estadoActual,
        '',
        ''
      ]]);
      reemplazarAsociados_(shAgr, oldPadre, validado.padre, validado.asociados);
      return { ok: true, message: 'Registro actualizado correctamente.', id: id };
    }

    validarPadreUnico_(validado.padre, '');
    const nuevoId = generarNuevoId_(padres);
    shPadres.appendRow([
      nuevoId,
      fecha,
      validado.padre,
      usuario,
      sector,
      CONFIG_APP.ESTADO_PENDIENTE,
      '',
      ''
    ]);
    agregarAsociados_(shAgr, validado.padre, validado.asociados);
    return { ok: true, message: 'Registro guardado correctamente.', id: nuevoId };
  } finally {
    lock.releaseLock();
  }
}

function cambiarEstado_(id, estado, usuario, tipoUsuario) {
  id = normalizar_(id);
  estado = upper_(estado);
  tipoUsuario = normalizar_(tipoUsuario || '3');

  if (!id) throw new Error('ID requerido.');
  if (tipoUsuario === '3') throw new Error('No tiene permisos para cambiar estado.');
  if (![CONFIG_APP.ESTADO_PENDIENTE, CONFIG_APP.ESTADO_EJECUTADO].includes(estado)) {
    throw new Error('Estado inválido.');
  }

  const sh = hoja_(CONFIG_APP.HOJA_PADRES);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (normalizar_(data[i][0]) === id) {
      sh.getRange(i + 1, 6).setValue(estado);
      return { ok: true, message: 'Estado actualizado.', id: id, estado: estado };
    }
  }
  throw new Error('No se encontró el registro.');
}

function generarNuevoId_(padres) {
  let max = 0;
  for (let i = 1; i < padres.length; i++) {
    const id = normalizar_(padres[i][0]);
    const m = id.match(/^RA-(\d{6})$/i);
    if (m) max = Math.max(max, Number(m[1]));
  }
  return 'RA-' + String(max + 1).padStart(6, '0');
}

function validarPadreUnico_(padre, idPermitido) {
  const sh = hoja_(CONFIG_APP.HOJA_PADRES);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const id = normalizar_(data[i][0]);
    const p = normalizar_(data[i][2]);
    if (p === padre && id !== idPermitido) {
      throw new Error('El reporte padre ' + padre + ' ya existe en otro registro.');
    }
  }
}

function agregarAsociados_(shAgr, padre, asociados) {
  const rows = asociados.map(a => [padre, a]);
  if (rows.length) shAgr.getRange(shAgr.getLastRow() + 1, 1, rows.length, 2).setValues(rows);
}

function reemplazarAsociados_(shAgr, oldPadre, newPadre, asociados) {
  const data = shAgr.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (normalizar_(data[i][0]) === oldPadre) {
      shAgr.deleteRow(i + 1);
    }
  }
  agregarAsociados_(shAgr, newPadre, asociados);
}

function asegurarEncabezados_() {
  const shPadres = hoja_(CONFIG_APP.HOJA_PADRES);
  const shAgr = hoja_(CONFIG_APP.HOJA_AGRUPADOS);

  if (shPadres.getLastRow() === 0 || !normalizar_(shPadres.getRange(1, 1).getValue())) {
    shPadres.getRange(1, 1, 1, 8).setValues([['ID','FECHA','REPORTE PADRE','USUARIO','SECTOR','ESTADO','OPERADOR COD','OBSERVACIONES']]);
  }

  if (shAgr.getLastRow() === 0 || !normalizar_(shAgr.getRange(1, 1).getValue())) {
    shAgr.getRange(1, 1, 1, 2).setValues([['REPORTE PADRE','REPORTES ASOCIADOS']]);
  }
}
