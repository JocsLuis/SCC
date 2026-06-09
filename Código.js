// =============================================================================
// CEVAZ — SISTEMA INTEGRAL DE GESTIÓN DE CASOS
// Archivo: Código.gs — Lógica del Servidor (Google Apps Script)
// =============================================================================

// --- CONFIGURACIÓN GLOBAL ---
const CONFIG = {
  SHEET_NAME: "Respuestas de formulario 1", // Ajusta si el nombre difiere
  SPREADSHEET_ID: "", // Dejar vacío para usar el Spreadsheet activo (recomendado)

  // Mapeo de columnas a índice base-0 pruebaaaa
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
    "Plataforma Virtual",
    "Eventos Culturales"
  ],

  // Áreas que pertenecen al departamento "SAC"
  AREAS_SAC: [
    "Servicio al Estudiante",
    "Inscripciones",
    "Instalaciones",
    "Administración",
    "Seguridad",
    "Cafetería",
    "Atención al Cliente",
    "Otros"
  ],

  ESTADOS: {
    PENDIENTE:   "Pendiente",
    EN_PROCESO:  "En Proceso",
    RESUELTO:    "Resuelto"
  }
};

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
  const ss = CONFIG.SPREADSHEET_ID
    ? SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();

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
  if (lastRow < 2) return; // Sin datos

  const TOTAL_COLS_CONTROL = 7; // P hasta V = columnas 16 a 22 (base-0: 15–21)
  const dataRange = sheet.getRange(2, 1, lastRow - 1, 22); // Filas de datos, 22 columnas
  const data = dataRange.getValues();

  const updates = [];
  const rowsToUpdate = [];

  data.forEach((fila, i) => {
    const sheetRow = i + 2; // Fila real en el sheet (encabezado en fila 1)
    const idCasoExistente = fila[CONFIG.COLS.ID_CASO];

    // Solo inicializar si la columna P está vacía
    if (!idCasoExistente) {
      const marcaTemporal = fila[CONFIG.COLS.MARCA_TEMPORAL];
      const area          = fila[CONFIG.COLS.AREA_PROGRAMA];
      const idCaso        = generarIdCaso_(marcaTemporal, sheetRow);
      const departamento  = asignarDepartamento_(area);

      // Columnas P-V: [ID_CASO, DEPARTAMENTO, ESTADO, PERSONAL, FECHA_RES, TIEMPO, COMENTARIOS]
      const controlRow = [
        idCaso,
        departamento,
        CONFIG.ESTADOS.PENDIENTE,
        "",   // Personal (vacío hasta que alguien tome el caso)
        "",   // Fecha Resolución
        "",   // Tiempo Respuesta
        ""    // Comentarios
      ];
      updates.push(controlRow);
      rowsToUpdate.push(sheetRow);
    }
  });

  // Escribir en lote para eficiencia
  rowsToUpdate.forEach((sheetRow, idx) => {
    sheet.getRange(sheetRow, 16, 1, 7).setValues([updates[idx]]);
  });

  return `${rowsToUpdate.length} filas inicializadas.`;
}

/**
 * Trigger: Se ejecuta automáticamente al recibir una nueva respuesta del formulario.
 * Configurar en: Triggers > onFormSubmit > Hoja de cálculo
 */
function onFormSubmit(e) {
  try {
    const sheet   = getSheet_();
    const lastRow = sheet.getLastRow();
    const fila    = sheet.getRange(lastRow, 1, 1, 15).getValues()[0];

    const marcaTemporal = fila[CONFIG.COLS.MARCA_TEMPORAL];
    const area          = fila[CONFIG.COLS.AREA_PROGRAMA];
    const idCaso        = generarIdCaso_(marcaTemporal, lastRow);
    const departamento  = asignarDepartamento_(area);

    const controlRow = [
      [idCaso, departamento, CONFIG.ESTADOS.PENDIENTE, "", "", "", ""]
    ];

    sheet.getRange(lastRow, 16, 1, 7).setValues(controlRow);

    Logger.log(`Caso creado: ${idCaso} → ${departamento}`);
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
function tomarCaso(filaNum, nombrePersonal) {
  try {
    validarEntrada_({ filaNum, nombrePersonal });

    const sheet = getSheet_();
    const C = CONFIG.COLS;

    // Leer estado actual (columna R = col 18 en base-0 → col 18 en base-1)
    const estadoActual = sheet.getRange(filaNum, C.ESTADO + 1).getValue();

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
function resolverCaso(filaNum, nombrePersonal, comentarios) {
  try {
    validarEntrada_({ filaNum, nombrePersonal, comentarios });

    const sheet = getSheet_();
    const C = CONFIG.COLS;

    // Leer fila completa para verificar estado y obtener marca temporal
    const filaData = sheet.getRange(filaNum, 1, 1, 22).getValues()[0];
    const estadoActual = filaData[C.ESTADO];

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
    sheet.getRange(filaNum, C.ESTADO + 1).setValue(CONFIG.ESTADOS.RESUELTO);
    sheet.getRange(filaNum, C.PERSONAL + 1).setValue(nombrePersonal.trim());
    sheet.getRange(filaNum, C.FECHA_RESOLUCION + 1).setValue(ahora);
    sheet.getRange(filaNum, C.TIEMPO_RESPUESTA + 1).setValue(tiempoRespuesta);
    sheet.getRange(filaNum, C.COMENTARIOS + 1).setValue(comentarios.trim());

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
  } catch {
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

  Logger.log("=== TEST COMPLETADO ===");
}
