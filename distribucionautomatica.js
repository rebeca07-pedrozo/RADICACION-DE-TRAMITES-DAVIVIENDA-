/**
 * ============================================================================
 * DISTRIBUCIÓN AUTOMÁTICA DE SOLICITUDES POR IMPUESTO
 * ============================================================================
 *
 * Hoja origen:
 *   - IMPORT
 *
 * Hojas destino:
 *   - Reintegro de retencion de ICA
 *   - Reintegro de retencion de renta
 *   - Reintegro de retencion de IVA
 *   - Reintegro de impuesto IVA
 *
 * El script:
 * - NO modifica la hoja IMPORT
 * - Limpia y reconstruye las hojas destino
 * - Soporta múltiples impuestos por fila
 * - Filtra solo motivos:
 *      Marcación
 *      Reintegro
 *      Ambas
 * ============================================================================
 */

function distribuirSolicitudes() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const HOJA_IMPORT = "IMPORT";

  const CONFIG = [
    {
      texto: "Retención de ICA",
      hoja: "Reintegro de retencion de ICA"
    },
    {
      texto: "Retención de Renta",
      hoja: "Reintegro de retencion de renta"
    },
    {
      texto: "Retención de IVA",
      hoja: "Reintegro de retencion de IVA"
    },
    {
      texto: "Impuesto de IVA",
      hoja: "Reintegro de impuesto IVA"
    }
  ];

  const MOTIVOS_VALIDOS = [
    "Marcación",
    "Reintegro",
    "Ambas"
  ];

  // =========================================================================
  // LEER IMPORT
  // =========================================================================

  const hojaImport = ss.getSheetByName(HOJA_IMPORT);

  if (!hojaImport) {
    throw new Error(`No existe la hoja '${HOJA_IMPORT}'`);
  }

  const datos = hojaImport.getDataRange().getValues();

  if (datos.length <= 1) {
    Logger.log("No hay datos para procesar.");
    return;
  }

  const encabezados = datos[0];
  const filas = datos.slice(1);

  // =========================================================================
  // DETECTAR COLUMNAS
  // =========================================================================

  const encabezadosNormalizados = encabezados.map(h =>
    normalizarTexto(h)
  );

  const idxImpuestos = buscarIndice(encabezadosNormalizados, [
    "impuestos",
    "tipoimpuesto",
    "tipodeimpuesto"
  ]);

  const idxMotivo = buscarIndice(encabezadosNormalizados, [
    "motivo",
    "motivodelasolicitud",
    "tipodesolicitud"
  ]);

  if (idxImpuestos === -1) {
    throw new Error("No se encontró la columna de impuestos.");
  }

  if (idxMotivo === -1) {
    throw new Error("No se encontró la columna de motivo.");
  }

  // =========================================================================
  // PREPARAR CONTENEDORES
  // =========================================================================

  const resultados = {};

  CONFIG.forEach(cfg => {
    resultados[cfg.hoja] = [encabezados];
  });

  // =========================================================================
  // PROCESAR FILAS
  // =========================================================================

  filas.forEach(fila => {

    const motivo = (fila[idxMotivo] || "").toString().trim();

    if (!MOTIVOS_VALIDOS.includes(motivo)) {
      return;
    }

    const impuestosTexto = (fila[idxImpuestos] || "").toString();

    CONFIG.forEach(cfg => {

      if (impuestosTexto.includes(cfg.texto)) {
        resultados[cfg.hoja].push(fila);
      }

    });

  });

  // =========================================================================
  // ESCRIBIR HOJAS DESTINO
  // =========================================================================

  CONFIG.forEach(cfg => {

    let hojaDestino = ss.getSheetByName(cfg.hoja);

    if (!hojaDestino) {
      hojaDestino = ss.insertSheet(cfg.hoja);
    }

    hojaDestino.clearContents();

    const data = resultados[cfg.hoja];

    if (data.length > 0) {

      hojaDestino
        .getRange(1, 1, data.length, data[0].length)
        .setValues(data);

    }

  });

  Logger.log("Distribución completada correctamente.");
}

/**
 * ============================================================================
 * NORMALIZAR TEXTO
 * ============================================================================
 */

function normalizarTexto(texto) {

  return (texto || "")
    .toString()
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9]/g, "");

}

/**
 * ============================================================================
 * BUSCAR ÍNDICE DE COLUMNA
 * ============================================================================
 */

function buscarIndice(encabezados, opciones) {

  for (let i = 0; i < encabezados.length; i++) {

    if (opciones.includes(encabezados[i])) {
      return i;
    }

  }

  return -1;
}