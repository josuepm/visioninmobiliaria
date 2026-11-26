/**
 * ============================================================================
 * VISIÓN INMOBILIARIA — Google Apps Script Backend (Web App)
 * ============================================================================
 * Conecta directamente con la hoja de cálculo:
 * https://docs.google.com/spreadsheets/d/1Ku65cMPFa9kBIp_No8KsVog5_4dLi2lLITdhTFkhyEU/edit
 *
 * INSTRUCCIONES DE INSTALACIÓN (1 minuto):
 * 1. Abre tu hoja de cálculo en Google Sheets.
 * 2. Ve al menú: Extensiones -> Apps Script.
 * 3. Borra el contenido inicial y pega todo este código. Guarda (Ctrl+S).
 * 4. Haz clic en "Implementar" -> "Nueva implementación".
 * 5. Tipo: "Aplicación web" (Web app).
 *    - Ejecutar como: "Yo" (tu cuenta de Google).
 *    - Quién tiene acceso: "Cualquier usuario" (Anyone).
 * 6. Autoriza los permisos y copia la URL generada (termina en /exec).
 * 7. Pega esa URL (y el ADMIN_PIN) en Panel de Control.html -> "⚙ Base de Datos".
 *    (Opcional) Pega la misma URL en DEFAULT_GAS_URL dentro de index.html y
 *    Panel de Control.html para que funcione automáticamente en cualquier dispositivo.
 * ============================================================================
 */

const SPREADSHEET_ID = '1Ku65cMPFa9kBIp_No8KsVog5_4dLi2lLITdhTFkhyEU';
const ADMIN_PIN = '2026'; // Cambia este PIN por uno seguro si lo deseas

const SHEET_PROSPECTOS = 'Prospectos Web';
const SHEET_PROPIEDADES = 'Propiedades Publicadas';
const HEADER_ROW = 10;
const DATA_START_ROW = 11;

function getSpreadsheet_() {
  try {
    const active = SpreadsheetApp.getActiveSpreadsheet();
    if (active && active.getId() === SPREADSHEET_ID) return active;
  } catch (e) {}
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function jsonResponse_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function normalizeText_(val) {
  return String(val === null || val === undefined ? '' : val).trim();
}

function parsePriceNumber_(val) {
  if (typeof val === 'number') return val;
  const cleaned = String(val || '').replace(/[^0-9.]/g, '');
  const num = parseFloat(cleaned);
  return isNaN(num) ? 0 : num;
}

/**
 * Encuentra la primera fila disponible a partir de DATA_START_ROW donde la
 * columna clave (col A) esté vacía. Evita escribir hasta la fila 101 cuando
 * las filas 11..100 tienen formato o validaciones preaplicadas.
 */
function getNextAvailableRow_(sheet, startRow, keyCol) {
  const maxRows = Math.max(sheet.getMaxRows(), startRow);
  const numRows = maxRows - startRow + 1;
  const keys = sheet.getRange(startRow, keyCol, numRows, 1).getValues();
  for (let i = 0; i < keys.length; i++) {
    if (normalizeText_(keys[i][0]) === '') {
      return startRow + i;
    }
  }
  return maxRows + 1;
}

function readProspectos_(ss) {
  const sheet = ss.getSheetByName(SHEET_PROSPECTOS);
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) return [];

  const numRows = lastRow - DATA_START_ROW + 1;
  const displayRows = sheet.getRange(DATA_START_ROW, 1, numRows, 12).getDisplayValues();
  const rawRows = sheet.getRange(DATA_START_ROW, 1, numRows, 12).getValues();

  const list = [];
  for (let i = 0; i < displayRows.length; i++) {
    const row = displayRows[i];
    const folio = normalizeText_(row[0]);
    const nombre = normalizeText_(row[2]);
    if (!folio && !nombre) continue;

    list.push({
      rowIndex: DATA_START_ROW + i,
      folio: folio,
      fechaHora: normalizeText_(row[1]),
      nombre: nombre,
      telefono: normalizeText_(row[3]),
      origenLinea: normalizeText_(row[4]),
      objetivo: normalizeText_(row[5]),
      tipoPropiedad: normalizeText_(row[6]),
      ciudad: normalizeText_(row[7]),
      presupuesto: normalizeText_(row[8]),
      presupuestoRaw: rawRows[i][8],
      detalles: normalizeText_(row[9]),
      estado: normalizeText_(row[10]) || 'Nuevo',
      asesor: normalizeText_(row[11]) || 'Josué Peña'
    });
  }
  return list.reverse();
}

function readPropiedades_(ss, onlyActive) {
  const sheet = ss.getSheetByName(SHEET_PROPIEDADES);
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) return [];

  const numRows = lastRow - DATA_START_ROW + 1;
  const displayRows = sheet.getRange(DATA_START_ROW, 1, numRows, 11).getDisplayValues();
  const rawRows = sheet.getRange(DATA_START_ROW, 1, numRows, 11).getValues();

  const list = [];
  for (let i = 0; i < displayRows.length; i++) {
    const row = displayRows[i];
    const idPropiedad = normalizeText_(row[0]);
    const titulo = normalizeText_(row[2]);
    if (!idPropiedad && !titulo) continue;

    const estadoPub = normalizeText_(row[8]) || 'Activa';
    if (onlyActive && estadoPub.toLowerCase() !== 'activa' && estadoPub.toLowerCase() !== 'disponible') {
      continue;
    }

    list.push({
      rowIndex: DATA_START_ROW + i,
      idPropiedad: idPropiedad,
      fechaPublicacion: normalizeText_(row[1]),
      titulo: titulo,
      operacion: normalizeText_(row[3]) || 'Venta',
      tipoPropiedad: normalizeText_(row[4]) || 'Casa',
      ciudad: normalizeText_(row[5]),
      precio: parsePriceNumber_(rawRows[i][6] || row[6]),
      precioTexto: normalizeText_(row[6]),
      caracteristicas: normalizeText_(row[7]),
      estadoPublicacion: estadoPub,
      agente: normalizeText_(row[9]) || 'Josué Peña',
      urlImagen: normalizeText_(row[10])
    });
  }
  return list.reverse();
}

function findRowByColA_(sheet, value) {
  const target = normalizeText_(value);
  if (!target) return -1;
  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) return -1;
  const colValues = sheet.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, 1).getDisplayValues();
  for (let i = 0; i < colValues.length; i++) {
    if (normalizeText_(colValues[i][0]) === target) {
      return DATA_START_ROW + i;
    }
  }
  return -1;
}

function doGet(e) {
  try {
    const params = (e && e.parameter) || {};
    const action = normalizeText_(params.action || 'getPublicProperties');
    const ss = getSpreadsheet_();

    if (action === 'getPublicProperties') {
      const properties = readPropiedades_(ss, true);
      return jsonResponse_({
        ok: true,
        properties: properties,
        updatedAt: new Date().toISOString()
      });
    }

    if (action === 'getAdminData') {
      if (normalizeText_(params.pin) !== ADMIN_PIN) {
        return jsonResponse_({ ok: false, error: 'PIN administrativo incorrecto.' });
      }
      const leads = readProspectos_(ss);
      const properties = readPropiedades_(ss, false);
      return jsonResponse_({
        ok: true,
        leads: leads,
        properties: properties,
        updatedAt: new Date().toISOString()
      });
    }

    return jsonResponse_({ ok: false, error: 'Acción GET no reconocida: ' + action });
  } catch (err) {
    return jsonResponse_({ ok: false, error: err.message || String(err) });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const action = normalizeText_(body.action);
    const ss = getSpreadsheet_();

    // 1) Alta pública de prospecto desde index.html o desde el panel
    if (action === 'createLead') {
      const lead = body.lead || body;
      const nombre = normalizeText_(lead.nombre || lead.Nombre_Cliente);
      const telefono = normalizeText_(lead.telefono || lead.Telefono_WhatsApp);
      if (!nombre || !telefono) {
        return jsonResponse_({ ok: false, error: 'Nombre y teléfono son obligatorios.' });
      }

      const sheet = ss.getSheetByName(SHEET_PROSPECTOS);
      if (!sheet) return jsonResponse_({ ok: false, error: 'No existe la pestaña Prospectos Web.' });

      const tz = ss.getSpreadsheetTimeZone() || 'America/Mexico_City';
      const now = new Date();
      const fechaHora = normalizeText_(lead.fechaHora) || Utilities.formatDate(now, tz, 'yyyy-MM-dd HH:mm');
      const folio = normalizeText_(lead.folio || lead.id) ||
        ('VI-' + Utilities.formatDate(now, tz, 'yyMMdd') + '-' + Math.floor(100 + Math.random() * 900));

      let origenLinea = normalizeText_(lead.origenLinea || lead.linea);
      if (origenLinea === 'usa' || origenLinea === 'EE. UU.') {
        origenLinea = 'EE. UU. (+1 915 478 7644)';
      } else if (origenLinea === 'mx' || origenLinea === 'México') {
        origenLinea = 'México (+52 656 206 6331)';
      }

      const rowValues = [
        folio,
        fechaHora,
        nombre,
        telefono,
        origenLinea || 'EE. UU. (+1 915 478 7644)',
        normalizeText_(lead.objetivo || lead.intencion),
        normalizeText_(lead.tipoPropiedad),
        normalizeText_(lead.ciudad),
        normalizeText_(lead.presupuesto),
        normalizeText_(lead.detalles || lead.notas),
        normalizeText_(lead.estado) || 'Nuevo',
        normalizeText_(lead.asesor) || 'Josué Peña'
      ];

      const targetRow = getNextAvailableRow_(sheet, DATA_START_ROW, 1);
      sheet.getRange(targetRow, 1, 1, 12).setValues([rowValues]);

      return jsonResponse_({
        ok: true,
        folio: folio,
        rowIndex: targetRow
      });
    }

    // Acciones administrativas protegidas por PIN
    if (normalizeText_(body.pin) !== ADMIN_PIN) {
      return jsonResponse_({ ok: false, error: 'PIN administrativo incorrecto.' });
    }

    // 2) Actualizar estado, detalles o asesor de un prospecto
    if (action === 'updateLead') {
      const sheet = ss.getSheetByName(SHEET_PROSPECTOS);
      const rowIndex = findRowByColA_(sheet, body.folio);
      if (rowIndex < DATA_START_ROW) {
        return jsonResponse_({ ok: false, error: 'Folio no encontrado en Prospectos Web.' });
      }
      if (body.detalles !== undefined) {
        sheet.getRange(rowIndex, 10).setValue(normalizeText_(body.detalles));
      }
      if (body.estado !== undefined) {
        sheet.getRange(rowIndex, 11).setValue(normalizeText_(body.estado));
      }
      if (body.asesor !== undefined) {
        sheet.getRange(rowIndex, 12).setValue(normalizeText_(body.asesor));
      }
      return jsonResponse_({ ok: true, rowIndex: rowIndex });
    }

    // 3) Eliminar prospecto
    if (action === 'deleteLead') {
      const sheet = ss.getSheetByName(SHEET_PROSPECTOS);
      const rowIndex = findRowByColA_(sheet, body.folio);
      if (rowIndex >= DATA_START_ROW) {
        sheet.deleteRow(rowIndex);
      }
      return jsonResponse_({ ok: true });
    }

    // 4) Crear o actualizar propiedad en "Propiedades Publicadas"
    if (action === 'saveProperty') {
      const prop = body.property || {};
      const sheet = ss.getSheetByName(SHEET_PROPIEDADES);
      if (!sheet) return jsonResponse_({ ok: false, error: 'No existe la pestaña Propiedades Publicadas.' });

      const tz = ss.getSpreadsheetTimeZone() || 'America/Mexico_City';
      const now = new Date();
      const idProp = normalizeText_(prop.idPropiedad) ||
        ('PROP-' + Utilities.formatDate(now, tz, 'yyyy') + '-' + Math.floor(100 + Math.random() * 900));
      const fechaPub = normalizeText_(prop.fechaPublicacion) || Utilities.formatDate(now, tz, 'yyyy-MM-dd');

      const precioNum = parsePriceNumber_(prop.precio);
      const rowValues = [
        idProp,
        fechaPub,
        normalizeText_(prop.titulo),
        normalizeText_(prop.operacion) || 'Venta',
        normalizeText_(prop.tipoPropiedad) || 'Casa',
        normalizeText_(prop.ciudad),
        precioNum > 0 ? precioNum : normalizeText_(prop.precioTexto || prop.precio),
        normalizeText_(prop.caracteristicas),
        normalizeText_(prop.estadoPublicacion) || 'Activa',
        normalizeText_(prop.agente) || 'Josué Peña',
        normalizeText_(prop.urlImagen)
      ];

      const existingRow = findRowByColA_(sheet, idProp);
      const targetRow = existingRow >= DATA_START_ROW
        ? existingRow
        : getNextAvailableRow_(sheet, DATA_START_ROW, 1);

      sheet.getRange(targetRow, 1, 1, 11).setValues([rowValues]);

      return jsonResponse_({
        ok: true,
        idPropiedad: idProp,
        rowIndex: targetRow
      });
    }

    // 5) Cambiar estado de publicación de una propiedad
    if (action === 'updatePropertyStatus') {
      const sheet = ss.getSheetByName(SHEET_PROPIEDADES);
      const rowIndex = findRowByColA_(sheet, body.idPropiedad);
      if (rowIndex < DATA_START_ROW) {
        return jsonResponse_({ ok: false, error: 'Propiedad no encontrada en la hoja.' });
      }
      sheet.getRange(rowIndex, 9).setValue(normalizeText_(body.estadoPublicacion));
      return jsonResponse_({ ok: true, rowIndex: rowIndex });
    }

    // 6) Eliminar propiedad
    if (action === 'deleteProperty') {
      const sheet = ss.getSheetByName(SHEET_PROPIEDADES);
      const rowIndex = findRowByColA_(sheet, body.idPropiedad);
      if (rowIndex >= DATA_START_ROW) {
        sheet.deleteRow(rowIndex);
      }
      return jsonResponse_({ ok: true });
    }

    return jsonResponse_({ ok: false, error: 'Acción POST no reconocida: ' + action });
  } catch (err) {
    return jsonResponse_({ ok: false, error: err.message || String(err) });
  } finally {
    lock.releaseLock();
  }
}
