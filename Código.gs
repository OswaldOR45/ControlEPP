// ============================================================
// CONTROL DE EPP - BACKEND (Código.gs)
// Google Apps Script
// ------------------------------------------------------------
// Modelo de EVENTOS: cada falta es un renglón con su propia
// fecha en la hoja "_Registro". Esto permite:
//   - Ventana móvil de 30 días (las faltas "caducan" solas).
//   - Histórico permanente (nada se sobrescribe).
//   - Análisis por área / EPP / reincidentes.
// Las 16 hojas de área siguen siendo el PADRÓN (nombres en
// columna B desde la fila 3). No cambia el alta de personal.
// ============================================================

// SUSTITUYE por tu ID real
const SPREADSHEET_ID = ""

// ---------- Configuración del negocio ----------
const VENTANA_DIAS  = 30;   // ventana móvil para contar faltas
const LIMITE_FALTAS = 3;    // faltas que disparan el escalamiento

// ---------- Hojas auxiliares ----------
const SHEET_REGISTRO  = "_Registro";
const SHEET_HISTORICO = "_Historico";

// ---------- Carpeta privada de firmas ----------
const FIRMAS_FOLDER_NAME = "Firmas EPP (privado)";

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

// Columnas del PADRÓN (hojas de área)
const COL = {
  NOMBRE: 2 // B
};

// Columnas de la hoja _Registro (1 falta por renglón)
const REG = {
  FECHA:  1, // A
  AREA:   2, // B
  NOMBRE: 3, // C
  EPP:    4, // D
  FIRMA:  5  // E (URL en Drive)
};

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
// DATOS INICIALES
// ============================================================
function getInitialData() {
  try {
    SpreadsheetApp.openById(SPREADSHEET_ID); // valida acceso
    return {
      ok: true,
      areas: AREAS,
      eppItems: EPP_ITEMS,
      config: { ventanaDias: VENTANA_DIAS, limiteFaltas: LIMITE_FALTAS }
    };
  } catch (e) {
    return { ok: false, mensaje: e.message };
  }
}

function getAreas()   { return AREAS; }
function getEppItems(){ return EPP_ITEMS; }

// ============================================================
// HELPERS DE HOJA / FECHA
// ============================================================
function _ss() { return SpreadsheetApp.openById(SPREADSHEET_ID); }

function _tz(ss) { return ss.getSpreadsheetTimeZone() || "America/Mexico_City"; }

function _fmt(date, ss) {
  if (!date) return "";
  return Utilities.formatDate(new Date(date), _tz(ss), "dd/MM/yyyy HH:mm");
}
function _fmtFecha(date, ss) {
  if (!date) return "";
  return Utilities.formatDate(new Date(date), _tz(ss), "dd/MM/yyyy");
}

// Fecha de corte: hace `dias` días a partir de ahora (ventana móvil)
function _corte(dias) {
  return new Date(Date.now() - dias * 24 * 60 * 60 * 1000);
}

function _getRegistroSheet(ss) {
  let sh = ss.getSheetByName(SHEET_REGISTRO);
  if (!sh) {
    sh = ss.insertSheet(SHEET_REGISTRO);
    sh.getRange(1, 1, 1, 5).setValues([["FECHA", "AREA", "NOMBRE", "EPP", "FIRMA"]]);
    sh.getRange(1, 1, 1, 5).setFontWeight("bold").setBackground("#1a73e8").setFontColor("white");
    sh.setFrozenRows(1);
  }
  return sh;
}

function _getHistoricoSheet(ss) {
  let sh = ss.getSheetByName(SHEET_HISTORICO);
  if (!sh) {
    sh = ss.insertSheet(SHEET_HISTORICO);
    sh.getRange(1, 1, 1, 6).setValues([[
      "FECHA ESCALAMIENTO", "AREA", "NOMBRE", "FALTAS EN VENTANA", "DETALLE", "ESTADO"
    ]]);
    sh.getRange(1, 1, 1, 6).setFontWeight("bold").setBackground("#CC2A1A").setFontColor("white");
    sh.setFrozenRows(1);
  }
  return sh;
}

// Carpeta privada para las firmas (NO se comparte públicamente)
function _getFirmasFolder() {
  const it = DriveApp.getFoldersByName(FIRMAS_FOLDER_NAME);
  if (it.hasNext()) return it.next();
  return DriveApp.createFolder(FIRMAS_FOLDER_NAME);
}

// ============================================================
// INICIALIZACIÓN DE HOJAS
// ============================================================
function initSheets() {
  const ss = _ss();
  AREAS.forEach(area => {
    let sheet = ss.getSheetByName(area);
    if (!sheet) sheet = ss.insertSheet(area);
    const headers = sheet.getRange(1, 1, 1, 2).getValues()[0];
    if (!headers[0]) {
      sheet.getRange(1, 1, 1, 2).setValues([["ÁREA", "NOMBRE"]]);
      sheet.getRange(1, 1, 1, 2).setFontWeight("bold").setBackground("#1a73e8").setFontColor("white");
      sheet.getRange(2, 1).setValue(area);
    }
  });
  _getRegistroSheet(ss);
  _getHistoricoSheet(ss);
  _getFirmasFolder();
  return { ok: true, mensaje: "Hojas y carpeta de firmas listas." };
}

// ============================================================
// PADRÓN: PERSONAL DE UN ÁREA (sin cambios de UX)
// ============================================================
function getPersonal(area) {
  const ss = _ss();
  const sheet = ss.getSheetByName(area);
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return [];
  const nombres = sheet.getRange(3, COL.NOMBRE, lastRow - 2, 1).getValues();
  return nombres.map(r => r[0]).filter(n => n && String(n).trim() !== "").map(String);
}

function agregarEmpleado(area, nombre) {
  try {
    const ss = _ss();
    const sheet = ss.getSheetByName(area);
    if (!sheet) return { ok: false, mensaje: "Área no encontrada." };
    const lastRow = sheet.getLastRow();
    const nuevaFila = Math.max(lastRow + 1, 3);
    sheet.getRange(nuevaFila, COL.NOMBRE).setValue(String(nombre).trim());
    return { ok: true, mensaje: `Empleado "${nombre}" agregado al área "${area}".` };
  } catch (e) {
    return { ok: false, mensaje: e.message };
  }
}

// ============================================================
// LECTURA DEL REGISTRO DE EVENTOS
// ============================================================
function _leerRegistro(ss) {
  const sh = _getRegistroSheet(ss);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const vals = sh.getRange(2, 1, last - 1, 5).getValues();
  const out = [];
  vals.forEach(r => {
    if (!r[REG.FECHA - 1] || !r[REG.NOMBRE - 1]) return;
    out.push({
      fecha:  new Date(r[REG.FECHA - 1]),
      area:   String(r[REG.AREA - 1]),
      nombre: String(r[REG.NOMBRE - 1]).trim(),
      epp:    String(r[REG.EPP - 1] || ""),
      firma:  String(r[REG.FIRMA - 1] || "")
    });
  });
  return out;
}

// Faltas de un empleado evaluadas en ventanas fijas de 30 días desde la primera falta
function _getEstadoFaltasFijas(area, nombre, todosLosRegistros) {
  const n = String(nombre).trim();

  // 1. Filtrar solo las de este empleado y ordenar cronológicamente
  const faltasEmpleado = todosLosRegistros
    .filter(r => r.area === area && r.nombre === n)
    .sort((a, b) => a.fecha - b.fecha);

  if (faltasEmpleado.length === 0) {
    return { cantidad: 0, faltas: [], diasRestantes: 0 };
  }

  const hoy = new Date();
  let i = 0;

  // 2. Simular las ventanas
  while (i < faltasEmpleado.length) {
    let inicio = new Date(faltasEmpleado[i].fecha);
    let fin = new Date(inicio.getTime() + (VENTANA_DIAS * 24 * 60 * 60 * 1000));
    let faltasEnVentana = [];

    // Agrupar las que caen dentro de estos 30 días
    while (i < faltasEmpleado.length && faltasEmpleado[i].fecha <= fin) {
      faltasEnVentana.push(faltasEmpleado[i]);
      i++;
    }

    // Si el día de hoy cae dentro de esta ventana, esta es la ventana activa
    if (hoy <= fin) {
      return {
        cantidad: faltasEnVentana.length,
        faltas: faltasEnVentana.reverse(), // Mostrar la más reciente primero
        diasRestantes: Math.ceil((fin - hoy) / (1000 * 60 * 60 * 24)),
        inicio: inicio,
        fin: fin
      };
    }
  }

  // Si todas sus ventanas ya expiraron
  return { cantidad: 0, faltas: [], diasRestantes: 0 };
}

// ============================================================
// GUARDAR FIRMA EN CARPETA PRIVADA (sin enlace público)
// ============================================================
function _guardarFirma(base64Data, nombreEmpleado) {
  try {
    const base64 = base64Data.replace(/^data:image\/\w+;base64,/, "");
    const blob = Utilities.newBlob(
      Utilities.base64Decode(base64),
      "image/png",
      `Firma_${nombreEmpleado}_${Date.now()}.png`
    );
    const folder = _getFirmasFolder();
    const file = folder.createFile(blob);
    // IMPORTANTE: NO se llama setSharing(ANYONE_WITH_LINK).
    // El archivo queda privado para el dueño; el acceso de los
    // supervisores se da compartiéndoles la carpeta una sola vez.
    return file.getUrl();
  } catch (e) {
    Logger.log("Error guardando firma: " + e.message);
    return "[Firma registrada - error al guardar imagen]";
  }
}

// ============================================================
// REGISTRAR INCUMPLIMIENTO (modelo de eventos)
// payload = { area, nombre, faltas:[], firmaBase64, confirmado:bool }
// ============================================================
function registrarIncumplimiento(payload) {
  try {
    const ss = _ss();
    if (!payload || !payload.area || !payload.nombre)
      throw new Error("Faltan datos del registro.");
    if (!payload.faltas || payload.faltas.length === 0)
      throw new Error("Selecciona al menos un EPP incumplido.");

    // Conteo ANTES de registrar
    let todosLosRegistros = _leerRegistro(ss);
    const estadoPrevio = _getEstadoFaltasFijas(payload.area, payload.nombre, todosLosRegistros);
    const conteoPrevio = estadoPrevio.cantidad;

    if (conteoPrevio >= LIMITE_FALTAS && !payload.confirmado) {
      return {
        ok: false,
        requiereConfirmacion: true,
        acumulados: conteoPrevio,
        mensaje: `${payload.nombre} ya tiene ${conteoPrevio} falta(s) en su periodo actual de ${VENTANA_DIAS} días. ¿Registrar otra de todos modos?`
      };
    }

    // Guardar firma (privada)
    let firmaUrl = "";
    if (payload.firmaBase64) {
      firmaUrl = _guardarFirma(payload.firmaBase64, payload.nombre);
    }

    // Agregar el evento a la hoja
    const sh = _getRegistroSheet(ss);
    const ahora = new Date();
    const eppStr = payload.faltas.join(", ");
    sh.appendRow([ahora, payload.area, String(payload.nombre).trim(), eppStr, firmaUrl]);

    // Conteo DESPUÉS de registrar (volvemos a leer la base de datos actualizada)
    todosLosRegistros = _leerRegistro(ss);
    const estadoActual = _getEstadoFaltasFijas(payload.area, payload.nombre, todosLosRegistros);
    const conteoActual = estadoActual.cantidad;
    let escalado = false;

    // Escalamiento (Solo al cruzar el límite)
    if (conteoActual === LIMITE_FALTAS) {
      escalado = true;
      const detalle = estadoActual.faltas
        .slice(0, LIMITE_FALTAS)
        .map((e, i) => `${i + 1}) ${_fmt(e.fecha, ss)} — ${e.epp}`)
        .join("  |  ");
      _getHistoricoSheet(ss).appendRow([
        ahora, payload.area, String(payload.nombre).trim(),
        conteoActual, detalle, "Pendiente"
      ]);
    }

    let mensaje;
    if (conteoActual >= LIMITE_FALTAS) {
      mensaje = `Falta #${conteoActual} registrada para ${payload.nombre}. ` +
                (escalado
                  ? `¡ALERTA: alcanzó ${LIMITE_FALTAS} faltas! Se guardó en el histórico para el supervisor.`
                  : `Lleva ${conteoActual} faltas en su periodo actual.`);
    } else {
      mensaje = `Falta registrada para ${payload.nombre}. Acumuladas: ${conteoActual}. (Su historial se limpiará en ${estadoActual.diasRestantes} días).`;
    }

    return {
      ok: true,
      acumulados: conteoActual,
      alerta: conteoActual >= LIMITE_FALTAS,
      escalado: escalado,
      mensaje: mensaje
    };

  } catch (e) {
    return { ok: false, mensaje: e.message };
  }
}

// ============================================================
// HISTORIAL DE UN EMPLEADO (dentro de la ventana)
// ============================================================
function getIncumplimientosEmpleado(area, nombre) {
  try {
    const ss = _ss();
    const todosLosRegistros = _leerRegistro(ss);
    const estado = _getEstadoFaltasFijas(area, nombre, todosLosRegistros);
    return {
      ok: true,
      data: {
        total: estado.cantidad,
        ventanaDias: VENTANA_DIAS,
        faltas: estado.faltas.map(f => ({ fecha: _fmt(f.fecha, ss), epp: f.epp }))
      }
    };
  } catch (e) {
    return { ok: false, mensaje: e.message };
  }
}

function verificarAlerta(area, nombre) {
  const ss = _ss();
  const todosLosRegistros = _leerRegistro(ss);
  const estado = _getEstadoFaltasFijas(area, nombre, todosLosRegistros);
  return { alerta: estado.cantidad >= LIMITE_FALTAS, acumulados: estado.cantidad };
}

// ============================================================
// REPORTE DIARIO (eventos de hoy)
// ============================================================
function getReporteDiario() {
  try {
    const ss = _ss();
    const hoy = new Date(); hoy.setHours(0, 0, 0, 0);
    const manana = new Date(hoy); manana.setDate(manana.getDate() + 1);

    const registro = _leerRegistro(ss);
    // conteo en ventana por persona (para marcar alerta)
    const corte = _corte(VENTANA_DIAS);
    const ventana = {};
    registro.forEach(e => {
      if (e.fecha >= corte) {
        const k = e.area + "||" + e.nombre;
        ventana[k] = (ventana[k] || 0) + 1;
      }
    });

    const reporte = [];
    registro.forEach(e => {
      if (e.fecha >= hoy && e.fecha < manana) {
        const k = e.area + "||" + e.nombre;
        reporte.push({
          area: e.area,
          nombre: e.nombre,
          faltas: e.epp ? e.epp.split(", ") : [],
          timestamp: _fmt(e.fecha, ss),
          totalFaltas: ventana[k] || 1 // acumulado en ventana
        });
      }
    });

    return { ok: true, data: reporte, fecha: _fmtFecha(hoy, ss) };
  } catch (e) {
    return { ok: false, mensaje: e.message };
  }
}

// ============================================================
// REPORTE MENSUAL (mes calendario) — agregado por empleado
// ============================================================
function getReporteMensual(anio, mes) {
  try {
    const ss = _ss();
    const inicio = new Date(anio, mes, 1);
    const fin = new Date(anio, mes + 1, 0, 23, 59, 59);

    const registro = _leerRegistro(ss).filter(e => e.fecha >= inicio && e.fecha <= fin);

    const porArea = {};
    const porEmpleado = {}; // area||nombre -> {area, nombre, eppSet, count, ultima}
    registro.forEach(e => {
      porArea[e.area] = (porArea[e.area] || 0) + 1;
      const k = e.area + "||" + e.nombre;
      if (!porEmpleado[k]) porEmpleado[k] = { area: e.area, nombre: e.nombre, epp: {}, count: 0, ultima: e.fecha };
      porEmpleado[k].count++;
      if (e.fecha > porEmpleado[k].ultima) porEmpleado[k].ultima = e.fecha;
      (e.epp ? e.epp.split(", ") : []).forEach(x => { if (x) porEmpleado[k].epp[x] = true; });
    });

    const data = [];
    const empleadosConAlerta = [];
    Object.values(porEmpleado).forEach(p => {
      data.push({
        area: p.area, nombre: p.nombre,
        faltas: Object.keys(p.epp),
        timestamp: _fmt(p.ultima, ss),
        totalFaltas: p.count
      });
      if (p.count >= LIMITE_FALTAS) empleadosConAlerta.push({ nombre: p.nombre, area: p.area, total: p.count });
    });

    const meses = ["Enero","Febrero","Marzo","Abril","Mayo","Junio",
                   "Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"];
    return {
      ok: true,
      data: data,
      resumen: {
        mes: `${meses[mes]} ${anio}`,
        totalIncumplimientos: registro.length,
        totalEmpleados: data.length,
        porArea: porArea,
        empleadosConAlerta: empleadosConAlerta
      }
    };
  } catch (e) {
    return { ok: false, mensaje: e.message };
  }
}

// ============================================================
// ANÁLISIS: ÁREAS MÁS PROBLEMÁTICAS (ventana configurable)
// ============================================================
function getAnalisis(dias) {
  try {
    const ss = _ss();
    const d = dias && dias > 0 ? dias : VENTANA_DIAS;
    const corte = _corte(d);

    const registro = _leerRegistro(ss).filter(e => e.fecha >= corte);

    const porArea = {};
    const porEpp = {};
    const porEmpleado = {}; // area||nombre -> {area, nombre, count}

    registro.forEach(e => {
      porArea[e.area] = (porArea[e.area] || 0) + 1;
      (e.epp ? e.epp.split(", ") : []).forEach(x => { if (x) porEpp[x] = (porEpp[x] || 0) + 1; });
      const k = e.area + "||" + e.nombre;
      if (!porEmpleado[k]) porEmpleado[k] = { area: e.area, nombre: e.nombre, count: 0 };
      porEmpleado[k].count++;
    });

    const reincidentes = Object.values(porEmpleado)
      .sort((a, b) => b.count - a.count)
      .slice(0, 10);

    // Escalamientos (histórico) dentro de la ventana
    const hsh = _getHistoricoSheet(ss);
    const escalamientos = [];
    const hlast = hsh.getLastRow();
    if (hlast >= 2) {
      hsh.getRange(2, 1, hlast - 1, 6).getValues().forEach(r => {
        if (!r[0]) return;
        const f = new Date(r[0]);
        if (f >= corte) {
          escalamientos.push({
            fecha: _fmt(f, ss), area: String(r[1]), nombre: String(r[2]),
            faltas: r[3], detalle: String(r[4]), estado: String(r[5])
          });
        }
      });
    }
    escalamientos.reverse(); // más reciente primero

    const areasAfectadas = Object.keys(porArea).length;
    const empleadosUnicos = Object.keys(porEmpleado).length;

    return {
      ok: true,
      dias: d,
      desde: _fmtFecha(corte, ss),
      hasta: _fmtFecha(new Date(), ss),
      totalFaltas: registro.length,
      areasAfectadas: areasAfectadas,
      empleadosUnicos: empleadosUnicos,
      porArea: porArea,
      porEpp: porEpp,
      reincidentes: reincidentes,
      escalamientos: escalamientos
    };
  } catch (e) {
    return { ok: false, mensaje: e.message };
  }
}

// ============================================================
// MIGRACIÓN: del modelo viejo (3 columnas por fila) al
// registro de eventos. Idempotente: limpia A y C:H tras migrar,
// así que correrla otra vez no duplica nada.
// ============================================================
function migrarDatos() {
  try {
    const ss = _ss();
    _getRegistroSheet(ss);
    _getHistoricoSheet(ss);
    _getFirmasFolder();

    let migradas = 0;
    const areasTocadas = [];

    AREAS.forEach(area => {
      const sheet = ss.getSheetByName(area);
      if (!sheet) return;
      const lastRow = sheet.getLastRow();
      if (lastRow < 3) return;

      // Modelo viejo: A=timestamp, B=nombre, C/D/E=falta1/2/3, F/G/H=firma1/2/3
      const datos = sheet.getRange(3, 1, lastRow - 2, 8).getValues();
      const filasAppend = [];

      datos.forEach((row, idx) => {
        const ts = row[0];
        const nombre = row[1];
        if (!nombre) return;
        const faltas = [row[2], row[3], row[4]];
        const firmas = [row[5], row[6], row[7]];
        let migroEsta = false;
        for (let i = 0; i < 3; i++) {
          const epp = faltas[i];
          if (epp && String(epp).trim() !== "") {
            // No hay fecha individual en el modelo viejo: usamos el
            // timestamp de la fila (o "ahora" si está vacío).
            const fecha = ts ? new Date(ts) : new Date();
            filasAppend.push([fecha, area, String(nombre).trim(), String(epp), String(firmas[i] || "")]);
            migradas++;
            migroEsta = true;
          }
        }
        if (migroEsta) {
          // Limpiar A y C:H, conservar el nombre (B) -> idempotente
          const fila = idx + 3;
          sheet.getRange(fila, 1).clearContent();          // A
          sheet.getRange(fila, 3, 1, 6).clearContent();    // C:H
        }
      });

      if (filasAppend.length > 0) {
        const reg = _getRegistroSheet(ss);
        reg.getRange(reg.getLastRow() + 1, 1, filasAppend.length, 5).setValues(filasAppend);
        areasTocadas.push(area);
      }
    });

    return {
      ok: true,
      mensaje: `Migración completa: ${migradas} falta(s) movidas a "${SHEET_REGISTRO}" desde ${areasTocadas.length} área(s). ` +
               `Nota: las faltas antiguas comparten la fecha de la fila original (el modelo viejo no guardaba fecha por falta).`
    };
  } catch (e) {
    return { ok: false, mensaje: e.message };
  }
}