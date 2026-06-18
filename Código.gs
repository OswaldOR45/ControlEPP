// ============================================================
// CONTROL DE EPP - BACKEND (Code.gs)
// Google Apps Script
// ============================================================

// SUSTITUYE por tu ID real
const SPREADSHEET_ID = "";

function getInitialData() {
  try {
    // Forzamos la apertura por ID para evitar el 'undefined'
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    return {
      ok: true,
      areas: AREAS, // El arreglo que ya tienes definido arriba
      eppItems: EPP_ITEMS
    };
  } catch (e) {
    console.error("Error en getInitialData: " + e.message);
    return { ok: false, mensaje: e.message };
  }
}
const AREAS = [
  "Control de calidad",
  "Limpieza",
  "Administración de producción",
  "Controles",
  "Chalanes de controles",
  "Extrusión",
  "Chalanes de extrusión",
  "Calderas",
  "Envasado",
  "Chalanes de envasado",
  "Personal General",
  "Muestras y etiquetas",
  "Encargado de cargadores",
  "RSI",
  "PERSONAL EXT.",
  "Cargadores"
];

const COL = {
  TIMESTAMP: 1, // A
  NOMBRE:    2, // B
  FALTA1:    3, // C
  FALTA2:    4, // D
  FALTA3:    5, // E
  FIRMA1:    6, // F
  FIRMA2:    7, // G
  FIRMA3:    8  // H
};

const EPP_ITEMS = [
  "Casco de seguridad",
  "Lentes de seguridad",
  "Calzado de seguridad",
  "Cofia",
  "Uniforme",
  "Objetos / Pertenencias no autorizados",
  "Barba",
  "Limpieza personal",
  "Uñas"
];

// ============================================================
// PUNTO DE ENTRADA WEB
// ============================================================
function doGet() {
  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("Control de EPP")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ============================================================
// INICIALIZACIÓN DE HOJAS
// Crea encabezados si no existen
// ============================================================
function initSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  AREAS.forEach(area => {
    let sheet = ss.getSheetByName(area);
    if (!sheet) {
      sheet = ss.insertSheet(area);
    }
    // Encabezados en fila 1
    const headers = sheet.getRange(1, 1, 1, 8).getValues()[0];
    if (!headers[0]) {
      sheet.getRange(1, 1, 1, 8).setValues([[
        "TIMESTAMP", "NOMBRE", "FALTA 1", "FALTA 2", "FALTA 3",
        "FIRMA 1", "FIRMA 2", "FIRMA 3"
      ]]);
      sheet.getRange(1, 1, 1, 8).setFontWeight("bold");
      sheet.getRange(1, 1, 1, 8).setBackground("#1a73e8");
      sheet.getRange(1, 1, 1, 8).setFontColor("white");
    }
  });
  return { ok: true };
}

// ============================================================
// OBTENER LISTA DE ÁREAS
// ============================================================
function getAreas() {
  return AREAS;
}

// ============================================================
// OBTENER PERSONAL DE UN ÁREA
// El nombre de cada empleado empieza en fila 3, col B
// ============================================================
function getPersonal(area) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(area);
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return [];

  // Leer nombres desde fila 3 hasta el final en col B
  const nombres = sheet.getRange(3, COL.NOMBRE, lastRow - 2, 1).getValues();
  return nombres
    .map(r => r[0])
    .filter(n => n && String(n).trim() !== "");
}

// ============================================================
// OBTENER ÍTEMS DE EPP
// ============================================================
function getEppItems() {
  return EPP_ITEMS;
}

// ============================================================
// REGISTRAR INCUMPLIMIENTO
// Busca la fila del empleado y agrega las faltas + firma en
// la primera columna disponible (FALTA 1, 2 o 3)
// ============================================================
function registrarIncumplimiento(payload) {
  /*
    payload = {
      area: string,
      nombre: string,
      faltas: [string],      // EPP incumplidos
      firmaBase64: string    // imagen en base64
    }
  */
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(payload.area);
    if (!sheet) throw new Error("Área no encontrada: " + payload.area);

    const lastRow = sheet.getLastRow();
    if (lastRow < 3) throw new Error("No hay personal registrado en esta área.");

    // Buscar la fila del empleado en col B desde fila 3
    const nombres = sheet.getRange(3, COL.NOMBRE, lastRow - 2, 1).getValues();
    let filaEmpleado = -1;
    for (let i = 0; i < nombres.length; i++) {
      if (String(nombres[i][0]).trim() === String(payload.nombre).trim()) {
        filaEmpleado = i + 3; // +3 porque empieza en fila 3
        break;
      }
    }
    if (filaEmpleado === -1) throw new Error("Empleado no encontrado: " + payload.nombre);

    // Leer el estado actual de la fila
    const rowData = sheet.getRange(filaEmpleado, 1, 1, 8).getValues()[0];
    const timestamp = rowData[COL.TIMESTAMP - 1];
    const falta1   = rowData[COL.FALTA1 - 1];
    const falta2   = rowData[COL.FALTA2 - 1];
    const falta3   = rowData[COL.FALTA3 - 1];
    const firma1   = rowData[COL.FIRMA1 - 1];
    const firma2   = rowData[COL.FIRMA2 - 1];
    const firma3   = rowData[COL.FIRMA3 - 1];

    // Contar cuántas faltas ya tiene
    const faltasExistentes = [falta1, falta2, falta3].filter(f => f && f !== "").length;

    if (faltasExistentes >= 3) {
      return {
        ok: false,
        mensaje: `${payload.nombre} ya tiene 3 incumplimientos registrados. Comuníquese con el supervisor.`,
        alerta3: true,
        acumulados: 3
      };
    }

    // Determinar qué columna usar
    const now = new Date();
    const faltasStr = payload.faltas.join(", ");

    if (faltasExistentes === 0) {
      // Primera falta
      sheet.getRange(filaEmpleado, COL.TIMESTAMP).setValue(now);
      sheet.getRange(filaEmpleado, COL.FALTA1).setValue(faltasStr);
      if (payload.firmaBase64) {
        _guardarFirmaComoImagen(sheet, filaEmpleado, COL.FIRMA1, payload.firmaBase64, payload.nombre);
      }
    } else if (faltasExistentes === 1) {
      // Segunda falta
      sheet.getRange(filaEmpleado, COL.FALTA2).setValue(faltasStr);
      if (payload.firmaBase64) {
        _guardarFirmaComoImagen(sheet, filaEmpleado, COL.FIRMA2, payload.firmaBase64, payload.nombre);
      }
    } else if (faltasExistentes === 2) {
      // Tercera falta
      sheet.getRange(filaEmpleado, COL.FALTA3).setValue(faltasStr);
      if (payload.firmaBase64) {
        _guardarFirmaComoImagen(sheet, filaEmpleado, COL.FIRMA3, payload.firmaBase64, payload.nombre);
      }
    }

    // Colorear fila si tiene 3 faltas
    const totalFaltas = faltasExistentes + 1;
    if (totalFaltas === 3) {
      sheet.getRange(filaEmpleado, 1, 1, 8).setBackground("#FFCCCC");
      return {
        ok: true,
        mensaje: `Incumplimiento #3 registrado para ${payload.nombre}. ¡ALERTA: 3 incumplimientos acumulados!`,
        alerta3: true,
        acumulados: 3
      };
    }

    return {
      ok: true,
      mensaje: `Incumplimiento registrado correctamente para ${payload.nombre}. Faltas acumuladas: ${totalFaltas}.`,
      alerta3: false,
      acumulados: totalFaltas
    };

  } catch (e) {
    return { ok: false, mensaje: e.message };
  }
}

// ============================================================
// GUARDAR FIRMA COMO IMAGEN EN DRIVE Y URL EN CELDA
// ============================================================
function _guardarFirmaComoImagen(sheet, fila, col, base64Data, nombreEmpleado) {
  try {
    // Remover prefijo data:image/png;base64,
    const base64 = base64Data.replace(/^data:image\/\w+;base64,/, "");
    const blob = Utilities.newBlob(
      Utilities.base64Decode(base64),
      "image/png",
      `Firma_${nombreEmpleado}_${Date.now()}.png`
    );

    // Guardar en Drive (carpeta raíz, puedes cambiar a una carpeta específica)
    const file = DriveApp.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const url = file.getUrl();

    sheet.getRange(fila, col).setValue(url);
  } catch (e) {
    Logger.log("Error guardando firma: " + e.message);
    sheet.getRange(fila, col).setValue("[Firma registrada - error al guardar imagen]");
  }
}

// ============================================================
// OBTENER INCUMPLIMIENTOS DE UN EMPLEADO
// ============================================================
function getIncumplimientosEmpleado(area, nombre) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(area);
    if (!sheet) return { ok: false, data: [] };

    const lastRow = sheet.getLastRow();
    if (lastRow < 3) return { ok: true, data: [] };

    const nombres = sheet.getRange(3, COL.NOMBRE, lastRow - 2, 1).getValues();
    for (let i = 0; i < nombres.length; i++) {
      if (String(nombres[i][0]).trim() === String(nombre).trim()) {
        const fila = i + 3;
        const row = sheet.getRange(fila, 1, 1, 8).getValues()[0];
        return {
          ok: true,
          data: {
            timestamp: row[COL.TIMESTAMP - 1] ? new Date(row[COL.TIMESTAMP - 1]).toLocaleString("es-MX") : "",
            falta1: row[COL.FALTA1 - 1] || "",
            falta2: row[COL.FALTA2 - 1] || "",
            falta3: row[COL.FALTA3 - 1] || "",
            firma1: row[COL.FIRMA1 - 1] || "",
            firma2: row[COL.FIRMA2 - 1] || "",
            firma3: row[COL.FIRMA3 - 1] || "",
            total: [row[COL.FALTA1-1], row[COL.FALTA2-1], row[COL.FALTA3-1]].filter(f => f && f !== "").length
          }
        };
      }
    }
    return { ok: true, data: null };
  } catch (e) {
    return { ok: false, mensaje: e.message };
  }
}

// ============================================================
// REPORTE DIARIO
// Devuelve todos los incumplimientos del día actual
// ============================================================
function getReporteDiario() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const hoy = new Date();
    hoy.setHours(0, 0, 0, 0);
    const manana = new Date(hoy);
    manana.setDate(manana.getDate() + 1);

    const reporte = [];

    AREAS.forEach(area => {
      const sheet = ss.getSheetByName(area);
      if (!sheet) return;

      const lastRow = sheet.getLastRow();
      if (lastRow < 3) return;

      const datos = sheet.getRange(3, 1, lastRow - 2, 8).getValues();
      datos.forEach(row => {
        const ts = row[COL.TIMESTAMP - 1];
        const nombre = row[COL.NOMBRE - 1];
        if (!ts || !nombre) return;

        const fecha = new Date(ts);
        if (fecha >= hoy && fecha < manana) {
          const faltas = [row[COL.FALTA1-1], row[COL.FALTA2-1], row[COL.FALTA3-1]].filter(f => f && f !== "");
          if (faltas.length > 0) {
            reporte.push({
              area,
              nombre: String(nombre),
              faltas,
              timestamp: fecha.toLocaleString("es-MX"),
              totalFaltas: faltas.length
            });
          }
        }
      });
    });

    return { ok: true, data: reporte, fecha: hoy.toLocaleDateString("es-MX") };
  } catch (e) {
    return { ok: false, mensaje: e.message };
  }
}

// ============================================================
// REPORTE MENSUAL
// Devuelve todos los incumplimientos del mes indicado
// ============================================================
function getReporteMensual(anio, mes) {
  // mes: 0-based (0=Enero, 11=Diciembre)
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const inicio = new Date(anio, mes, 1);
    const fin    = new Date(anio, mes + 1, 0, 23, 59, 59);

    const reporte = [];
    let totalIncumplimientos = 0;
    const porArea = {};
    const empleadosConAlerta = [];

    AREAS.forEach(area => {
      const sheet = ss.getSheetByName(area);
      if (!sheet) return;

      const lastRow = sheet.getLastRow();
      if (lastRow < 3) return;

      const datos = sheet.getRange(3, 1, lastRow - 2, 8).getValues();
      datos.forEach(row => {
        const ts = row[COL.TIMESTAMP - 1];
        const nombre = row[COL.NOMBRE - 1];
        if (!ts || !nombre) return;

        const fecha = new Date(ts);
        if (fecha >= inicio && fecha <= fin) {
          const faltas = [row[COL.FALTA1-1], row[COL.FALTA2-1], row[COL.FALTA3-1]].filter(f => f && f !== "");
          if (faltas.length > 0) {
            totalIncumplimientos += faltas.length;
            porArea[area] = (porArea[area] || 0) + faltas.length;

            if (faltas.length >= 3) {
              empleadosConAlerta.push({ nombre: String(nombre), area });
            }

            reporte.push({
              area,
              nombre: String(nombre),
              faltas,
              timestamp: fecha.toLocaleString("es-MX"),
              totalFaltas: faltas.length
            });
          }
        }
      });
    });

    const meses = ["Enero","Febrero","Marzo","Abril","Mayo","Junio",
                   "Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"];

    return {
      ok: true,
      data: reporte,
      resumen: {
        mes: `${meses[mes]} ${anio}`,
        totalIncumplimientos,
        totalEmpleados: reporte.length,
        porArea,
        empleadosConAlerta
      }
    };
  } catch (e) {
    return { ok: false, mensaje: e.message };
  }
}

// ============================================================
// AGREGAR EMPLEADO A UN ÁREA
// ============================================================
function agregarEmpleado(area, nombre) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(area);
    if (!sheet) return { ok: false, mensaje: "Área no encontrada." };

    const lastRow = sheet.getLastRow();
    const nuevaFila = Math.max(lastRow + 1, 3);
    sheet.getRange(nuevaFila, COL.NOMBRE).setValue(nombre.trim());
    return { ok: true, mensaje: `Empleado "${nombre}" agregado al área "${area}".` };
  } catch (e) {
    return { ok: false, mensaje: e.message };
  }
}

// ============================================================
// VERIFICAR ALERTA DE 3 INCUMPLIMIENTOS
// ============================================================
function verificarAlerta(area, nombre) {
  const result = getIncumplimientosEmpleado(area, nombre);
  if (!result.ok || !result.data) return { alerta: false };
  return {
    alerta: result.data.total >= 3,
    acumulados: result.data.total
  };
}