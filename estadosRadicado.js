/**
 * ============================================================================
 * DISTRIBUCIÓN AUTOMÁTICA DE SOLICITUDES POR IMPUESTO CON ESTADOS
 * ============================================================================
 */

// 👇 ÍNDICES DE COLUMNAS (0 = columna A, 1 = B, 2 = C...)
const IDX_RADICADO  = 1; // Columna B
const IDX_DOCUMENTO = 6; // Columna G

function distribuirSolicitudes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const HOJA_IMPORT = "IMPORT";

  const CONFIG = [
    { texto: "Retención de ICA", hoja: "Reintegro de retencion de ICA" },
    { texto: "Retención de Renta", hoja: "Reintegro de retencion de renta" },
    { texto: "Retención de IVA", hoja: "Reintegro de retencion de IVA" },
    { texto: "Impuesto de IVA", hoja: "Reintegro de impuesto IVA" }
  ];

  const MOTIVOS_VALIDOS = ["Marcación", "Reintegro", "Ambas"];
  const ESTADOS_DISPONIBLES = ["RECIBIDO EN PROCESO", "APROBADO", "RECHAZADO", "REQUERIDO"];
  const ESTADO_POR_DEFECTO = "RECIBIDO EN PROCESO";

  const hojaImport = ss.getSheetByName(HOJA_IMPORT);
  if (!hojaImport) throw new Error(`No existe la hoja '${HOJA_IMPORT}'`);

  const datos = hojaImport.getDataRange().getValues();
  if (datos.length <= 1) {
    Logger.log("No hay datos para procesar.");
    return;
  }

  const encabezadosOriginales = datos[0];
  const filas = datos.slice(1);

  // Crear nuevos encabezados (R = col 18, S = 19, T = 20)
  let nuevosEncabezados = [...encabezadosOriginales];
  while (nuevosEncabezados.length < 18) nuevosEncabezados.push("");
  nuevosEncabezados[18] = "ESTADO";        // Columna S
  nuevosEncabezados[19] = "OBSERVACIONES"; // Columna T

  const encabezadosNormalizados = encabezadosOriginales.map(h => normalizarTexto(h));
  const idxImpuestos = buscarIndice(encabezadosNormalizados, ["impuestos", "tipoimpuesto", "tipodeimpuesto"]);
  const idxMotivo = buscarIndice(encabezadosNormalizados, ["motivo", "motivodelasolicitud", "tipodesolicitud"]);

  if (idxImpuestos === -1 || idxMotivo === -1) {
    throw new Error("No se encontraron las columnas requeridas (Impuestos o Motivo).");
  }

  // Preparar contenedores
  const resultados = {};
  CONFIG.forEach(cfg => {
    resultados[cfg.hoja] = [];
  });

  // Procesar filas
  filas.forEach(fila => {
    const motivo = (fila[idxMotivo] || "").toString().trim();
    if (!MOTIVOS_VALIDOS.includes(motivo)) return;

    const impuestosTexto = (fila[idxImpuestos] || "").toString();

    // Completar la fila hasta la columna T (índice 19)
    let filaCompleta = [...fila];
    while (filaCompleta.length < 18) filaCompleta.push("");

    filaCompleta[18] = ESTADO_POR_DEFECTO; // Columna S
    filaCompleta[19] = "";                 // Columna T

    CONFIG.forEach(cfg => {
      if (impuestosTexto.includes(cfg.texto)) {
        resultados[cfg.hoja].push(filaCompleta);
      }
    });
  });

  // Escribir hojas destino
  CONFIG.forEach(cfg => {
    let hojaDestino = ss.getSheetByName(cfg.hoja);
    if (!hojaDestino) hojaDestino = ss.insertSheet(cfg.hoja);

    hojaDestino.clearContents();
    hojaDestino.clearFormats();

    const dataFinal = [nuevosEncabezados, ...resultados[cfg.hoja]];

    if (dataFinal.length > 1) {
      hojaDestino.getRange(1, 1, dataFinal.length, dataFinal[0].length).setValues(dataFinal);

      // Chips desplegables en la columna S
      const rangoChips = hojaDestino.getRange(2, 19, dataFinal.length - 1, 1);
      const reglaEx = SpreadsheetApp.newDataValidation()
        .requireValueInList(ESTADOS_DISPONIBLES)
        .setAllowInvalid(false)
        .build();
      rangoChips.setDataValidation(reglaEx);
      rangoChips.setBackground("#fff2cc");
    }
  });

  Logger.log("Distribución y configuración de estados completada.");
}

function normalizarTexto(texto) {
  return (texto || "").toString().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^a-z0-9]/g, "");
}

function buscarIndice(encabezados, opciones) {
  for (let i = 0; i < encabezados.length; i++) {
    if (opciones.includes(encabezados[i])) return i;
  }
  return -1;
}

// =========================================================================
// INTERFAZ DE CONSULTA (WEB APP)
// =========================================================================

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Interfaz')
      .setTitle('Consulta de Estado de Solicitudes')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Normaliza un valor para comparación:
 * - Quita espacios
 * - Quita ceros a la izquierda (0013 = 13)
 * - Pasa a minúsculas
 */
function normalizarBusqueda(valor) {
  if (valor === null || valor === undefined) return "";
  let s = valor.toString().trim().toLowerCase();
  s = s.replace(/^0+(?=\d)/, "");
  return s;
}

/**
 * Busca por Radicado o por Cédula/NIT en todas las hojas destino.
 * Devuelve todas las coincidencias.
 */
function buscarRadicado(textoBusqueda) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojasAConsultar = [
    "Reintegro de retencion de ICA",
    "Reintegro de retencion de renta",
    "Reintegro de retencion de IVA",
    "Reintegro de impuesto IVA"
  ];

  if (!textoBusqueda || !textoBusqueda.toString().trim()) {
    return { exito: false, mensaje: "Por favor ingresa un radicado o documento válido." };
  }

  const busqueda = normalizarBusqueda(textoBusqueda);
  const coincidencias = [];

  for (let nombreHoja of hojasAConsultar) {
    const hoja = ss.getSheetByName(nombreHoja);
    if (!hoja) continue;

    const datos = hoja.getDataRange().getValues();
    if (datos.length <= 1) continue;

    for (let i = 1; i < datos.length; i++) {
      const fila = datos[i];
      const radicadoCelda  = normalizarBusqueda(fila[IDX_RADICADO]);
      const documentoCelda = normalizarBusqueda(fila[IDX_DOCUMENTO]);

      if (radicadoCelda === busqueda || documentoCelda === busqueda) {
        coincidencias.push({
          radicado: fila[IDX_RADICADO],
          documento: fila[IDX_DOCUMENTO],
          impuesto: nombreHoja.replace("Reintegro de ", ""),
          estado: fila[18] || "RECIBIDO EN PROCESO",
          observaciones: fila[19] || "Sin observaciones registradas."
        });
      }
    }
  }

  if (coincidencias.length === 0) {
    return { exito: false, mensaje: "No se encontró ningún trámite con ese radicado o documento." };
  }

  return { exito: true, resultados: coincidencias };
}