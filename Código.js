// =============================================================================
// CEVAZ — SISTEMA INTEGRAL DE GESTIÓN DE CASOS
// Archivo: Código.gs — Lógica del Servidor (Google Apps Script)
// =============================================================================

// --- CONFIGURACIÓN GLOBAL ---
const CONFIG = {
  SHEET_NAME: "Respuestas de formulario 1", // Ajusta si el nombre difiere

  // Mapeo de columnas a índice base-0
  COLS: {
    MARCA_TEMPORAL:       0,  // A
    EMAIL_GOOGLE:         1,  // B
    NOMBRE_ALUMNO:        2,  // C
    ID_ALUMNO:            3,  // D
    EMAIL_CONTACTO:       4,  // E
    TIPO_PLANTEAMIENTO:   5,  // F
    AREA_PROGRAMA:        6,  // G
    NIVEL:                7,  // H
    CATEGORIA_CURSO:      8,  // I
    SEDE:                 9,  // J
    PROFESOR:            10,  // K
    HORARIO:             11,  // L
    TELEFONO:            12,  // M
    DESCRIPCION:         13,  // N
    ACCION_SOLICITADA:   14,  // O
    
    ID_CASO:             15,  // P
    DEPARTAMENTO:        16,  // Q
    ESTADO:              17,  // R
    PERSONAL:            18,  // S
    FECHA_RESOLUCION:    19,  // T
    TIEMPO_RESPUESTA:    20,  // U
    COMENTARIOS:         21,  // V
  },

  // Áreas que pertenecen al departamento "Academia"
  AREAS_ACADEMIA: [
    "Cursos de Inglés",
    "Profesores",
    "Plataforma Virtual"
  ],

  ESTADOS: {
    PENDIENTE:   "Pendiente",
    EN_PROCESO:  "En Proceso",
    RESUELTO:    "Resuelto"
  },

  USUARIOS_SHEET: "Usuarios",
  USUARIOS_COLS: {
    NOMBRE:       0,  // A
    CORREO:       1,  // B
    TELEFONO:     2,  // C
    SEDE:         3,  // D
    DEPARTAMENTO: 4,  // E
    ROL:          5   // F
  }
};

Object.freeze(CONFIG);

// =============================================================================
// CONFIGURACIÓN INICIAL — Ejecutar una vez desde el editor
// =============================================================================

/**
 * Configura el SPREADSHEET_ID en Properties Service.
 * Ejecutar una vez: Editor > Ejecutar > configurarSpreadsheetId
 */
function configurarSpreadsheetId() {
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', '1Q9HRlVl8-Ryn1PJeOxBAhx2DMf6qIO2LgqxLpKGI7HI');
  Logger.log("SPREADSHEET_ID configurado correctamente.");
}

/**
 * Retorna la URL completa de la hoja de cálculo para abrir desde el cliente.
 */
function getUrlHojaCalculo() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  return "https://docs.google.com/spreadsheets/d/" + id + "/edit";
}

// =============================================================================
// PUNTO DE ENTRADA — doGet()
// =============================================================================

/**
 * Sirve el dashboard HTML como WebApp.
 * URL de despliegue: Implementar > Nueva implementación > Aplicación web
 */
function doGet(e) {
  const template = HtmlService.createTemplateFromFile("Index");
  return template.evaluate()
    .setTitle("CEVAZ — Gestión de Casos")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag("viewport", "width=device-width, initial-scale=1.0");
}

/**
 * Utility para incluir archivos HTML (CSS y JS) dentro del template principal.
 * Uso en Index.html: <?!= include('Estilos') ?>
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// =============================================================================
// ACCESO A LA HOJA DE CÁLCULO
// =============================================================================

/**
 * Retorna la hoja de cálculo activa o por ID configurado.
 */
function getSheet_() {
  const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  const ss = spreadsheetId
    ? SpreadsheetApp.openById(spreadsheetId)
    : SpreadsheetApp.getActiveSpreadsheet();

  if (!ss) {
    throw new Error(
      "No se pudo acceder al spreadsheet. " +
      "Establece CONFIG.SPREADSHEET_ID con el ID de tu sheet " +
      "(lo encuentras en la URL: docs.google.com/spreadsheets/d/***ID***/edit)."
    );
  }

  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  if (!sheet) {
    throw new Error(
      `No se encontró la hoja "${CONFIG.SHEET_NAME}". ` +
      `Verifica el nombre exacto en CONFIG.SHEET_NAME.`
    );
  }
  return sheet;
}

// =============================================================================
// LÓGICA DE NEGOCIO — ASIGNACIÓN DE DEPARTAMENTOS
// =============================================================================

/**
 * Determina el departamento según el área/programa del caso.
 * @param {string} area - Valor de la columna G.
 * @returns {string} "Academia" | "SAC"
 */
function asignarDepartamento_(area) {
  if (!area) return "SAC";
  const areaLimpia = area.trim();

  if (CONFIG.AREAS_ACADEMIA.some(a => areaLimpia.toLowerCase().includes(a.toLowerCase()))) {
    return "Academia";
  }
  return "SAC";
}

/**
 * Genera un ID de caso legible basado en timestamp y número de fila.
 * Formato: CEVAZ-YYYYMMDD-NNNN
 * @param {Date} fecha - Marca temporal del formulario.
 * @param {number} fila - Número de fila en el sheet (base 1).
 * @returns {string}
 */
function generarIdCaso_(fecha, fila) {
  const f = fecha instanceof Date ? fecha : new Date(fecha);
  const anio  = f.getFullYear();
  const mes   = String(f.getMonth() + 1).padStart(2, "0");
  const dia   = String(f.getDate()).padStart(2, "0");
  const num   = String(fila).padStart(4, "0");
  return `CEVAZ-${anio}${mes}${dia}-${num}`;
}

// =============================================================================
// INICIALIZACIÓN — Rellenar columnas de control en filas nuevas
// =============================================================================

/**
 * Rellena las columnas de control (P–V) en todas las filas que aún no las tengan.
 * Ejecutar manualmente una vez, o asociar al trigger onFormSubmit.
 */
function inicializarColumnasControl() {
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return "Sin datos para inicializar.";

  // Leer datos y bloque de control existente en una sola llamada cada uno
  const data = sheet.getRange(2, 1, lastRow - 1, 22).getValues();
  const controlData = sheet.getRange(2, 16, lastRow - 1, 7).getValues();

  let contador = 0;

  data.forEach((fila, i) => {
    // Solo inicializar si la columna P (ID_CASO) está vacía
    if (!controlData[i][0]) {
      const marcaTemporal = fila[CONFIG.COLS.MARCA_TEMPORAL];
      const area          = fila[CONFIG.COLS.AREA_PROGRAMA];
      const idCaso        = generarIdCaso_(marcaTemporal, i + 2);
      const departamento  = asignarDepartamento_(area);

      controlData[i] = [
        idCaso,
        departamento,
        CONFIG.ESTADOS.PENDIENTE,
        "", "", "", ""
      ];
      contador++;
    }
  });

  // Una sola escritura en lote
  sheet.getRange(2, 16, lastRow - 1, 7).setValues(controlData);

  return `${contador} filas inicializadas.`;
}

/**
 * Trigger: Se ejecuta automáticamente al recibir una nueva respuesta del formulario.
 * Configurar en: Triggers > onFormSubmit > Hoja de cálculo
 */
function onFormSubmit(e) {
  try {
    const sheet   = e.range.getSheet();
    const filaNum = e.range.getRow();
    const fila    = sheet.getRange(filaNum, 1, 1, 15).getValues()[0];

    // Verificar si ya se procesó (evita duplicados por trigger)
    const idExistente = sheet.getRange(filaNum, CONFIG.COLS.ID_CASO + 1).getValue();
    if (idExistente) {
      Logger.log(`Fila ${filaNum} ya procesada (${idExistente}), saltando.`);
      return;
    }

    const marcaTemporal = fila[CONFIG.COLS.MARCA_TEMPORAL];
    const area          = fila[CONFIG.COLS.AREA_PROGRAMA];
    const idCaso        = generarIdCaso_(marcaTemporal, filaNum);
    const departamento  = asignarDepartamento_(area);

    sheet.getRange(filaNum, CONFIG.COLS.ID_CASO + 1, 1, 7).setValues([
      [idCaso, departamento, CONFIG.ESTADOS.PENDIENTE, "", "", "", ""]
    ]);

    Logger.log(`Caso creado: ${idCaso} → ${departamento}`);

    // Enviar correo si es QUEJA
    const tipoPlanteamiento = fila[CONFIG.COLS.TIPO_PLANTEAMIENTO] || "";
    if (tipoPlanteamiento.includes("QUEJA")) {
      enviarCorreoQueja({
        idCaso,
        marcaTemporal,
        nombreAlumno:    fila[CONFIG.COLS.NOMBRE_ALUMNO]    || "",
        idAlumno:        fila[CONFIG.COLS.ID_ALUMNO]        || "",
        sede:            fila[CONFIG.COLS.SEDE]              || "",
        departamento,
        areaPrograma:    area,
        profesor:        fila[CONFIG.COLS.PROFESOR]         || "",
        horario:         fila[CONFIG.COLS.HORARIO]          || "",
        telefono:        fila[CONFIG.COLS.TELEFONO]         || "",
        descripcion:     fila[CONFIG.COLS.DESCRIPCION]      || "",
        accionSolicitada: fila[CONFIG.COLS.ACCION_SOLICITADA] || ""
      });
    }
  } catch (err) {
    Logger.log(`Error en onFormSubmit: ${err.message}`);
  }
}

// =============================================================================
// LECTURA DE DATOS — getCasos()
// =============================================================================

/**
 * Lee todos los casos del sheet y los retorna como array de objetos JSON.
 * Llamado desde el cliente via google.script.run.getCasos()
 *
 * @param {Object} filtros - Opcional: { sede, departamento, estado }
 * @returns {Array<Object>} Array de objetos de caso.
 */
function getCasos(filtros) {
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 22).getValues();
  const C = CONFIG.COLS;
  const f = filtros || {};

  const casos = data
    .map((fila, i) => {
      // Ignorar filas completamente vacías
      if (!fila[C.MARCA_TEMPORAL] && !fila[C.NOMBRE_ALUMNO]) return null;

      const marcaTemporal    = fila[C.MARCA_TEMPORAL];
      const fechaRegistro    = marcaTemporal instanceof Date
        ? Utilities.formatDate(marcaTemporal, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm")
        : String(marcaTemporal);

      const fechaResolucion  = fila[C.FECHA_RESOLUCION];
      const fechaResStr      = fechaResolucion instanceof Date && fechaResolucion.getTime() > 0
        ? Utilities.formatDate(fechaResolucion, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm")
        : "";

      return {
        fila:              i + 2,                              // Número de fila en sheet
        idCaso:            fila[C.ID_CASO]            || "",
        fechaRegistro:     fechaRegistro,
        emailGoogle:       fila[C.EMAIL_GOOGLE]       || "",
        nombreAlumno:      fila[C.NOMBRE_ALUMNO]      || "",
        idAlumno:          fila[C.ID_ALUMNO]          || "",
        emailContacto:     fila[C.EMAIL_CONTACTO]     || "",
        tipoPlanteamiento: fila[C.TIPO_PLANTEAMIENTO] || "",
        areaPrograma:      fila[C.AREA_PROGRAMA]      || "",
        nivel:             fila[C.NIVEL]               || "",
        categoriaCurso:    fila[C.CATEGORIA_CURSO]    || "",
        sede:              fila[C.SEDE]               || "",
        profesor:          fila[C.PROFESOR]           || "",
        horario:           fila[C.HORARIO]            || "",
        telefono:          fila[C.TELEFONO]           || "",
        descripcion:       fila[C.DESCRIPCION]        || "",
        accionSolicitada:  fila[C.ACCION_SOLICITADA]  || "",
        departamento:      fila[C.DEPARTAMENTO]       || asignarDepartamento_(fila[C.AREA_PROGRAMA]),
        estado:            fila[C.ESTADO]             || CONFIG.ESTADOS.PENDIENTE,
        personal:          fila[C.PERSONAL]           || "",
        fechaResolucion:   fechaResStr,
        tiempoRespuesta:   fila[C.TIEMPO_RESPUESTA]   || "",
        comentarios:       fila[C.COMENTARIOS]        || ""
      };
    })
    .filter(caso => caso !== null);

  // Aplicar filtros si se proporcionaron
  return casos.filter(caso => {
    if (f.sede        && f.sede        !== "Todos" && caso.sede        !== f.sede)        return false;
    if (f.departamento && f.departamento !== "Todos" && caso.departamento !== f.departamento) return false;
    if (f.estado      && f.estado      !== "Todos" && caso.estado      !== f.estado)      return false;

    // Filtro por fecha
    if (f.fechaDesde || f.fechaHasta) {
      const fechaCaso = parsearFechaRegistro_(caso.fechaRegistro);
      if (f.fechaDesde) {
        const desde = new Date(f.fechaDesde + "T00:00:00");
        if (fechaCaso < desde) return false;
      }
      if (f.fechaHasta) {
        const hasta = new Date(f.fechaHasta + "T23:59:59");
        if (fechaCaso > hasta) return false;
      }
    }

    return true;
  });
}

// =============================================================================
// ACCIÓN — Tomar un Caso (Pendiente → En Proceso)
// =============================================================================

/**
 * Registra que un agente tomó el caso. Cambia estado a "En Proceso".
 * Llamado desde cliente: google.script.run.tomarCaso(filaNum, nombrePersonal)
 *
 * @param {number} filaNum     - Número de fila en el sheet.
 * @param {string} nombrePersonal - Nombre del agente que toma el caso.
 * @returns {Object} { exito: boolean, mensaje: string }
 */
function tomarCaso(filaNum, nombrePersonal, rolUsuario) {
  try {
    validarEntrada_({ filaNum, nombrePersonal });

    const sheet = getSheet_();
    const C = CONFIG.COLS;

    // Leer estado y departamento actual
    const estadoActual = sheet.getRange(filaNum, C.ESTADO + 1).getValue();
    const departamento = sheet.getRange(filaNum, C.DEPARTAMENTO + 1).getValue();

    // Validar permisos por departamento
    if (rolUsuario !== "Administrador" && departamento !== rolUsuario) {
      return { exito: false, mensaje: "No tienes permiso para tomar casos de este departamento." };
    }

    if (estadoActual === CONFIG.ESTADOS.RESUELTO) {
      return { exito: false, mensaje: "Este caso ya fue resuelto y no puede modificarse." };
    }
    if (estadoActual === CONFIG.ESTADOS.EN_PROCESO) {
      const personalActual = sheet.getRange(filaNum, C.PERSONAL + 1).getValue();
      return {
        exito: false,
        mensaje: `Este caso ya está siendo atendido por: ${personalActual || "un agente"}.`
      };
    }

    // Actualizar columnas R (Estado) y S (Personal) en una sola operación
    sheet.getRange(filaNum, C.ESTADO + 1, 1, 2).setValues([
      [CONFIG.ESTADOS.EN_PROCESO, nombrePersonal.trim()]
    ]);

    SpreadsheetApp.flush(); // Forzar escritura inmediata

    Logger.log(`Caso fila ${filaNum} tomado por: ${nombrePersonal}`);
    return {
      exito:   true,
      mensaje: `Caso asignado correctamente a ${nombrePersonal}.`,
      estado:  CONFIG.ESTADOS.EN_PROCESO,
      personal: nombrePersonal.trim()
    };

  } catch (err) {
    Logger.log(`Error en tomarCaso: ${err.message}`);
    return { exito: false, mensaje: `Error del servidor: ${err.message}` };
  }
}

// =============================================================================
// ACCIÓN — Resolver un Caso (En Proceso → Resuelto)
// =============================================================================

/**
 * Cierra el caso con resolución. Actualiza Estado, Personal, Fecha, Tiempo y Comentarios.
 * Llamado desde cliente: google.script.run.resolverCaso(filaNum, nombrePersonal, comentarios)
 *
 * @param {number} filaNum        - Número de fila en el sheet.
 * @param {string} nombrePersonal - Nombre del agente que resolvió.
 * @param {string} comentarios    - Descripción de la solución aplicada.
 * @returns {Object} { exito: boolean, mensaje: string, tiempoRespuesta: string }
 */
function resolverCaso(filaNum, nombrePersonal, comentarios, rolUsuario) {
  try {
    validarEntrada_({ filaNum, nombrePersonal, comentarios });

    const sheet = getSheet_();
    const C = CONFIG.COLS;

    // Leer fila completa para verificar estado, departamento y obtener marca temporal
    const filaData = sheet.getRange(filaNum, 1, 1, 22).getValues()[0];
    const estadoActual = filaData[C.ESTADO];
    const departamento = filaData[C.DEPARTAMENTO];

    // Validar permisos por departamento
    if (rolUsuario !== "Administrador" && departamento !== rolUsuario) {
      return { exito: false, mensaje: "No tienes permiso para resolver casos de este departamento." };
    }

    if (estadoActual === CONFIG.ESTADOS.RESUELTO) {
      return { exito: false, mensaje: "Este caso ya fue resuelto anteriormente." };
    }

    // Calcular tiempo de respuesta
    const marcaTemporal  = filaData[C.MARCA_TEMPORAL];
    const ahora          = new Date();
    const tiempoRespuesta = calcularTiempoRespuesta_(marcaTemporal, ahora);

    // Actualizar columnas S (Personal), T (Fecha Res), U (Tiempo), V (Comentarios), R (Estado)
    // Columna R = Estado (índice 17, col 18)
    // Columnas S-V = índices 18-21, cols 19-22
    sheet.getRange(filaNum, C.ESTADO + 1, 1, 5).setValues([[
      CONFIG.ESTADOS.RESUELTO,
      nombrePersonal.trim(),
      ahora,
      tiempoRespuesta,
      comentarios.trim()
    ]]);

    SpreadsheetApp.flush();

    const fechaStr = Utilities.formatDate(ahora, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
    Logger.log(`Caso fila ${filaNum} RESUELTO por ${nombrePersonal}. Tiempo: ${tiempoRespuesta}`);

    return {
      exito:           true,
      mensaje:         `Caso resuelto exitosamente por ${nombrePersonal}.`,
      estado:          CONFIG.ESTADOS.RESUELTO,
      fechaResolucion: fechaStr,
      tiempoRespuesta: tiempoRespuesta,
      personal:        nombrePersonal.trim()
    };

  } catch (err) {
    Logger.log(`Error en resolverCaso: ${err.message}`);
    return { exito: false, mensaje: `Error del servidor: ${err.message}` };
  }
}

// =============================================================================
// DATOS PARA ANALYTICS — getEstadisticas()
// =============================================================================

/**
 * Calcula y retorna estadísticas agregadas para los gráficos del dashboard.
 * Llamado desde cliente: google.script.run.getEstadisticas()
 *
 * @returns {Object} Objeto con métricas desglosadas.
 */
function getEstadisticas() {
  try {
    const casos = getCasos(); // Sin filtros: todos los casos

    // Inicializar acumuladores
    const stats = {
      totalCasos:     casos.length,
      porEstado: {
        [CONFIG.ESTADOS.PENDIENTE]:  0,
        [CONFIG.ESTADOS.EN_PROCESO]: 0,
        [CONFIG.ESTADOS.RESUELTO]:   0
      },
      porDepartamento: { SAC: 0, Academia: 0 },
      porSede:         {},
      porNivel:        {},
      porArea:         {},
      tiempoPromedio:  { SAC: [], Academia: [] }, // Arrays temporales para calcular promedio
    };

    casos.forEach(caso => {
      // Por estado
      if (stats.porEstado.hasOwnProperty(caso.estado)) {
        stats.porEstado[caso.estado]++;
      }

      // Por departamento
      const dep = caso.departamento || "SAC";
      stats.porDepartamento[dep] = (stats.porDepartamento[dep] || 0) + 1;

      // Por sede
      const sede = caso.sede || "Sin Sede";
      stats.porSede[sede] = (stats.porSede[sede] || 0) + 1;

      // Por nivel académico
      const nivel = caso.nivel || "No aplica";
      stats.porNivel[nivel] = (stats.porNivel[nivel] || 0) + 1;

      // Por área/programa
      const area = caso.areaPrograma || "Otros";
      stats.porArea[area] = (stats.porArea[area] || 0) + 1;

      // Tiempo de respuesta para casos resueltos
      if (caso.estado === CONFIG.ESTADOS.RESUELTO && caso.tiempoRespuesta) {
        const horas = parsearHorasDeTexto_(caso.tiempoRespuesta);
        if (horas !== null && dep in stats.tiempoPromedio) {
          stats.tiempoPromedio[dep].push(horas);
        }
      }
    });

    // Calcular promedios finales
    const promedioFinal = {};
    Object.keys(stats.tiempoPromedio).forEach(dep => {
      const arr = stats.tiempoPromedio[dep];
      promedioFinal[dep] = arr.length > 0
        ? Math.round(arr.reduce((a, b) => a + b, 0) / arr.length * 10) / 10
        : 0;
    });

    return {
      totalCasos:     stats.totalCasos,
      porEstado:      stats.porEstado,
      porDepartamento: stats.porDepartamento,
      porSede:        stats.porSede,
      porNivel:       stats.porNivel,
      porArea:        stats.porArea,
      tiempoPromedioHoras: promedioFinal
    };

  } catch (err) {
    Logger.log(`Error en getEstadisticas: ${err.message}`);
    return { error: err.message };
  }
}

// =============================================================================
// HELPERS PRIVADOS
// =============================================================================

/**
 * Calcula el tiempo de respuesta entre dos fechas.
 * Retorna string legible: "2d 4h" o "3h 15m"
 */
function calcularTiempoRespuesta_(fechaInicio, fechaFin) {
  try {
    const inicio = fechaInicio instanceof Date ? fechaInicio : new Date(fechaInicio);
    const fin    = fechaFin    instanceof Date ? fechaFin    : new Date(fechaFin);
    const diffMs = fin.getTime() - inicio.getTime();

    if (diffMs < 0) return "N/A";

    const totalMinutos = Math.floor(diffMs / 60000);
    const dias    = Math.floor(totalMinutos / 1440);
    const horas   = Math.floor((totalMinutos % 1440) / 60);
    const minutos = totalMinutos % 60;

    if (dias > 0)  return `${dias}d ${horas}h`;
    if (horas > 0) return `${horas}h ${minutos}m`;
    return `${minutos}m`;
  } catch (e) {
    Logger.log(`Error en calcularTiempoRespuesta_: ${e.message}`);
    return "N/A";
  }
}

/**
 * Convierte texto de tiempo ("2d 4h", "3h 15m") a número de horas para promedios.
 */
function parsearHorasDeTexto_(texto) {
  if (!texto || texto === "N/A") return null;
  let horas = 0;
  const diasMatch  = texto.match(/(\d+)d/);
  const horasMatch = texto.match(/(\d+)h/);
  const minsMatch  = texto.match(/(\d+)m/);
  if (diasMatch)  horas += parseInt(diasMatch[1])  * 24;
  if (horasMatch) horas += parseInt(horasMatch[1]);
  if (minsMatch)  horas += parseInt(minsMatch[1]) / 60;
  return horas;
}

/**
 * Valida entradas básicas de las funciones de acción.
 * Lanza Error si algún campo requerido falta.
 */
function validarEntrada_(params) {
  if (!params.filaNum || typeof params.filaNum !== "number" || params.filaNum < 2) {
    throw new Error("Número de fila inválido.");
  }
  if (params.nombrePersonal !== undefined && (!params.nombrePersonal || !params.nombrePersonal.trim())) {
    throw new Error("El nombre del personal es obligatorio.");
  }
  if (params.comentarios !== undefined && (!params.comentarios || !params.comentarios.trim())) {
    throw new Error("Los comentarios de solución son obligatorios.");
  }
}

/**
 * Parsea una fecha en formato "dd/MM/yyyy HH:mm" a objeto Date.
 * @param {string} fechaStr - Fecha en formato "dd/MM/yyyy HH:mm"
 * @returns {Date}
 */
function parsearFechaRegistro_(fechaStr) {
  if (!fechaStr) return new Date(0);
  try {
    const partes = fechaStr.split(" ");
    const fechaPart = partes[0].split("/");
    const horaPart = partes[1] || "00:00";
    const [hora, minuto] = horaPart.split(":");
    return new Date(
      parseInt(fechaPart[2]),
      parseInt(fechaPart[1]) - 1,
      parseInt(fechaPart[0]),
      parseInt(hora),
      parseInt(minuto)
    );
  } catch {
    return new Date(0);
  }
}

// =============================================================================
// GESTIÓN DE USUARIOS — CRUD sobre pestaña "Usuarios"
// =============================================================================

/**
 * Obtiene la pestaña "Usuarios" (la crea si no existe).
 */
function getUsuariosSheet_() {
  const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  const ss = spreadsheetId
    ? SpreadsheetApp.openById(spreadsheetId)
    : SpreadsheetApp.getActiveSpreadsheet();

  let sheet = ss.getSheetByName(CONFIG.USUARIOS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.USUARIOS_SHEET);
    sheet.appendRow(["Nombre", "Correo", "Teléfono", "Sede", "Departamento", "Rol"]);
    sheet.getRange(1, 1, 1, 6).setFontWeight("bold").setBackground("#f0f0f0");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * Lee todos los usuarios y retorna un array de objetos.
 * Agrega el número de fila (base-2) para poder editar/eliminar.
 */
function getUsuarios() {
  try {
    const sheet = getUsuariosSheet_();
    const data  = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    return data.slice(1).map((row, i) => ({
      fila:         i + 2,
      nombre:       row[CONFIG.USUARIOS_COLS.NOMBRE]      || "",
      correo:       row[CONFIG.USUARIOS_COLS.CORREO]      || "",
      telefono:     row[CONFIG.USUARIOS_COLS.TELEFONO]    || "",
      sede:         row[CONFIG.USUARIOS_COLS.SEDE]        || "",
      departamento: row[CONFIG.USUARIOS_COLS.DEPARTAMENTO] || "",
      rol:          row[CONFIG.USUARIOS_COLS.ROL]          || ""
    }));
  } catch (e) {
    Logger.log("Error getUsuarios: " + e.message);
    return [];
  }
}

/**
 * Crea un usuario nuevo. Recibe un objeto con nombre, correo, telefono, sede, departamento, rol.
 */
function crearUsuario(usuario) {
  try {
    const sheet = getUsuariosSheet_();
    sheet.appendRow([
      usuario.nombre      || "",
      usuario.correo      || "",
      usuario.telefono    || "",
      usuario.sede        || "",
      usuario.departamento || "",
      usuario.rol          || ""
    ]);
    return { exito: true, mensaje: "Usuario creado correctamente." };
  } catch (e) {
    return { exito: false, mensaje: "Error al crear usuario: " + e.message };
  }
}

/**
 * Edita un usuario existente por número de fila (base-2).
 */
function editarUsuario(filaNum, usuario) {
  try {
    const sheet = getUsuariosSheet_();
    const lastRow = sheet.getLastRow();
    if (filaNum < 2 || filaNum > lastRow) {
      return { exito: false, mensaje: "Número de fila inválido." };
    }
    sheet.getRange(filaNum, 1).setValue(usuario.nombre      || "");
    sheet.getRange(filaNum, 2).setValue(usuario.correo      || "");
    sheet.getRange(filaNum, 3).setValue(usuario.telefono    || "");
    sheet.getRange(filaNum, 4).setValue(usuario.sede        || "");
    sheet.getRange(filaNum, 5).setValue(usuario.departamento || "");
    sheet.getRange(filaNum, 6).setValue(usuario.rol          || "");
    return { exito: true, mensaje: "Usuario actualizado correctamente." };
  } catch (e) {
    return { exito: false, mensaje: "Error al editar usuario: " + e.message };
  }
}

/**
 * Elimina un usuario por número de fila (base-2).
 */
function eliminarUsuario(filaNum) {
  try {
    const sheet = getUsuariosSheet_();
    const lastRow = sheet.getLastRow();
    if (filaNum < 2 || filaNum > lastRow) {
      return { exito: false, mensaje: "Número de fila inválido." };
    }
    sheet.deleteRow(filaNum);
    return { exito: true, mensaje: "Usuario eliminado correctamente." };
  } catch (e) {
    return { exito: false, mensaje: "Error al eliminar usuario: " + e.message };
  }
}

// =============================================================================
// ENVÍO DE CORREO — QUEJAS
// =============================================================================

/**
 * Escapa HTML para usar en templates de correo (server-side).
 */
function escHtml_(str) {
  if (!str) return "";
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

/**
 * Busca los destinatarios para correo de QUEJA.
 * Retorna { admin: [...], supervisores: [...] }
 */
function buscarDestinatariosQueja_(departamento, sede) {
  const sheet = getUsuariosSheet_();
  const data  = sheet.getDataRange().getValues();
  const C     = CONFIG.USUARIOS_COLS;
  const admin = [];
  const supervisores = [];

  for (let i = 1; i < data.length; i++) {
    const row  = data[i];
    const rol  = (row[C.ROL] || "").toLowerCase();
    const correo = row[C.CORREO] || "";
    const sedeUsuario = row[C.SEDE] || "";

    if (!correo) continue;

    if (rol.includes("administrador")) {
      admin.push(correo);
    } else if (rol.includes("supervisor") &&
               rol.includes(departamento.toLowerCase()) &&
               sedeUsuario === sede) {
      supervisores.push(correo);
    }
  }
  return { admin, supervisores };
}

/**
 * Genera el template HTML para el correo de QUEJA.
 */
function generarTemplateQueja_(datos) {
  const d = datos;
  return `
  <!DOCTYPE html>
  <html>
  <head><meta charset="UTF-8"></head>
  <body style="margin:0;padding:0;background:#f4f4f4;font-family:Arial,sans-serif;">
    <div style="max-width:600px;margin:20px auto;background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,.1);">
      <div style="background:#C0392B;color:#fff;padding:20px 24px;">
        <h1 style="margin:0;font-size:18px;">NUEVA QUEJA REGISTRADA</h1>
        <p style="margin:6px 0 0;opacity:.9;font-size:14px;">${escHtml_(d.idCaso)}</p>
      </div>
      <div style="padding:20px 24px;">
        <table style="width:100%;border-collapse:collapse;font-size:14px;">
          <tr><td style="padding:8px 0;color:#666;width:140px;"><strong>Alumno</strong></td><td style="padding:8px 0;">${escHtml_(d.nombreAlumno)}</td></tr>
          <tr><td style="padding:8px 0;color:#666;"><strong>ID Alumno</strong></td><td style="padding:8px 0;">${escHtml_(d.idAlumno)}</td></tr>
          <tr><td style="padding:8px 0;color:#666;"><strong>Sede</strong></td><td style="padding:8px 0;">${escHtml_(d.sede)}</td></tr>
          <tr><td style="padding:8px 0;color:#666;"><strong>Departamento</strong></td><td style="padding:8px 0;">${escHtml_(d.departamento)}</td></tr>
          <tr><td style="padding:8px 0;color:#666;"><strong>Área</strong></td><td style="padding:8px 0;">${escHtml_(d.areaPrograma)}</td></tr>
          <tr><td style="padding:8px 0;color:#666;"><strong>Profesor</strong></td><td style="padding:8px 0;">${escHtml_(d.profesor)}</td></tr>
          <tr><td style="padding:8px 0;color:#666;"><strong>Horario</strong></td><td style="padding:8px 0;">${escHtml_(d.horario)}</td></tr>
          <tr><td style="padding:8px 0;color:#666;"><strong>Teléfono</strong></td><td style="padding:8px 0;">${escHtml_(d.telefono)}</td></tr>
          <tr><td colspan="2" style="padding:8px 0;border-top:1px solid #eee;"></td></tr>
          <tr><td style="padding:8px 0;color:#666;vertical-align:top;"><strong>Descripción</strong></td><td style="padding:8px 0;">${escHtml_(d.descripcion)}</td></tr>
          ${d.accionSolicitada ? `<tr><td style="padding:8px 0;color:#666;vertical-align:top;"><strong>Acción solicitada</strong></td><td style="padding:8px 0;">${escHtml_(d.accionSolicitada)}</td></tr>` : ""}
        </table>
      </div>
      <div style="background:#f9f9f9;padding:16px 24px;text-align:center;border-top:1px solid #eee;">
        <p style="margin:0;font-size:12px;color:#999;">Generado automáticamente por CEVAZ — Sistema de Control de Casos</p>
      </div>
    </div>
  </body>
  </html>`;
}

/**
 * Envía correo cuando se registra una QUEJA.
 * Admin recibe como principal. Supervisores del depto/sede reciben en copia (CC).
 */
function enviarCorreoQueja(datos) {
  try {
    const { admin, supervisores } = buscarDestinatariosQueja_(datos.departamento, datos.sede);
    if (admin.length === 0 && supervisores.length === 0) {
      Logger.log("No se encontraron destinatarios para queja " + datos.idCaso);
      return;
    }

    const asunto   = `QUEJA — ${datos.idCaso} — ${datos.nombreAlumno}`;
    const htmlBody = generarTemplateQueja_(datos);
    const to       = admin.join(",");
    const cc       = supervisores.join(",");

    const opciones = {
      htmlBody: htmlBody,
      name: "CEVAZ — Sistema de Casos"
    };
    if (cc) opciones.cc = cc;

    GmailApp.sendEmail(to, asunto, "", opciones);

    Logger.log(`Correo enviado: ${datos.idCaso} → To: ${to}` + (cc ? ` | CC: ${cc}` : ""));
  } catch (e) {
    Logger.log("Error enviando correo de queja: " + e.message);
  }
}

// =============================================================================
// FUNCIÓN DE PRUEBA — Ejecutar manualmente desde el editor
// =============================================================================

/**
 * Prueba rápida: Inicializa columnas y lee los primeros 3 casos.
 * Ejecutar en: Editor > Ejecutar > testSistema
 */
function testSistema() {
  Logger.log("=== INICIANDO TEST DEL SISTEMA CEVAZ ===");

  // 1. Inicializar columnas de control
  const resultadoInit = inicializarColumnasControl();
  Logger.log("Inicialización: " + resultadoInit);

  // 2. Leer casos sin filtro
  const casos = getCasos();
  Logger.log(`Total de casos leídos: ${casos.length}`);
  if (casos.length > 0) {
    Logger.log("Primer caso: " + JSON.stringify(casos[0], null, 2));
  }

  // 3. Obtener estadísticas
  const stats = getEstadisticas();
  Logger.log("Estadísticas: " + JSON.stringify(stats, null, 2));

}
