/**
 * ============================================================================
 * DISTRIBUCIÓN AUTOMÁTICA DE SOLICITUDES POR IMPUESTO CON ESTADOS
 * ============================================================================
 */

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

  // Crear nuevos encabezados para las hojas destino (R es col 18, S es 19, T es 20)
  // Aseguramos que los encabezados tengan longitud suficiente hasta la columna T
  let nuevosEncabezados = [...encabezadosOriginales];
  while (nuevosEncabezados.length < 18) nuevosEncabezados.push(""); // Rellenar si tiene menos de R
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

    // Completar la fila original hasta la columna R (índice 17) para que mantenga estructura
    let filaCompleta = [...fila];
    while (filaCompleta.length < 18) filaCompleta.push("");
    
    // Añadir Estado inicial y Observación vacía
    filaCompleta[18] = ESTADO_POR_DEFECTO; // Columna S
    filaCompleta[19] = "";                 // Columna T

    CONFIG.forEach(cfg => {
      if (impuestosTexto.includes(cfg.texto)) {
        resultados[cfg.hoja].push(filaCompleta);
      }
    });
  });

  // Escribir hojas destino y aplicar chips descriptivos
  CONFIG.forEach(cfg => {
    let hojaDestino = ss.getSheetByName(cfg.hoja);
    if (!hojaDestino) hojaDestino = ss.insertSheet(cfg.hoja);

    hojaDestino.clearContents();
    hojaDestino.clearFormats();

    const dataFinal = [nuevosEncabezados, ...resultados[cfg.hoja]];

    if (dataFinal.length > 1) {
      hojaDestino.getRange(1, 1, dataFinal.length, dataFinal[0].length).setValues(dataFinal);

      // Crear regla de validación para los Chips Desplegables en la Columna S (desde fila 2)
      const rangoChips = hojaDestino.getRange(2, 19, dataFinal.length - 1, 1);
      const reglaEx = SpreadsheetApp.newDataValidation()
        .requireValueInList(ESTADOS_DISPONIBLES)
        .setAllowInvalid(false)
        .build();
      rangoChips.setDataValidation(reglaEx);
      
      // Aplicar un color sutil (opcional)
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
// INTERFAZ DE CONSULTA DE RADICADOS (WEB APP)
// =========================================================================

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Interfaz')
      .setTitle('Consulta de Estado de Solicitudes')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Función que busca el radicado en todas las hojas destino
 */
function buscarRadicado(numeroRadicado) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojasAConsultar = [
    "Reintegro de retencion de ICA",
    "Reintegro de retencion de renta",
    "Reintegro de retencion de IVA",
    "Reintegro de impuesto IVA"
  ];
  
  if (!numeroRadicado) return { exito: false, mensaje: "Por favor ingresa un radicado válido." };
  
  numeroRadicado = numeroRadicado.toString().trim().toLowerCase();

  for (let nombreHoja of hojasAConsultar) {
    const hoja = ss.getSheetByName(nombreHoja);
    if (!hoja) continue;
    
    const datos = hoja.getDataRange().getValues();
    if (datos.length <= 1) continue;
    
    // Suponiendo que el RADICADO está en la Columna A (índice 0)
    for (let i = 1; i < datos.length; i++) {
      let radicadoCelda = (datos[i][0] || "").toString().trim().toLowerCase();
      
      if (radicadoCelda === numeroRadicado) {
        return {
          exito: true,
          radicado: datos[i][0],
          impuesto: nombreHoja.replace("Reintegro de ", ""),
          estado: datos[i][18] || "RECIBIDO EN PROCESO",
          observaciones: datos[i][19] || "Sin observaciones registradas."
        };
      }
    }
  }
  
  return { exito: false, mensaje: "No se encontró ningún trámite con el número de radicado ingresado." };
}