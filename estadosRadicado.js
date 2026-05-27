/**
 * ============================================================================
 * DISTRIBUCIÓN AUTOMÁTICA DE SOLICITUDES POR IMPUESTO CON ESTADOS
 * - Preserva estado y observaciones de filas existentes
 * - Agrega solo filas nuevas
 * - Elimina filas que ya no están en IMPORT
 * ============================================================================
 */

// 👇 ÍNDICES DE COLUMNAS (0 = A, 1 = B, 2 = C...)
const IDX_RADICADO  = 1; // Columna B (identificador único)
const IDX_DOCUMENTO = 6; // Columna G (Cédula/NIT)

function distribuirSolicitudes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const HOJA_IMPORT = "IMPORT";

  // Cada hoja destino tiene una lista de textos que la activan.
  // Si el campo "Impuestos" de IMPORT contiene cualquiera de esos textos, la fila va a esa hoja.
  const CONFIG = [
    {
      hoja: "Reintegro de retencion de ICA",
      textos: ["Retención de ICA"]
    },
    {
      hoja: "Reintegro de retencion de renta",
      textos: [
        "Retención de Renta",
        "JELPIT",
        "Propiedad horizontal",   // captura "Propiedad horizontal (rete ICA y retefuente)"
        "Régimen simple"          // captura "Régimen simple (rete ICA y retefuente)"
      ]
    },
    {
      hoja: "Reintegro de retencion de IVA",
      textos: ["Retención de IVA"]
    },
    {
      hoja: "Reintegro de impuesto IVA",
      textos: ["Impuesto de IVA"]
    }
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

  // Construir encabezados finales (con ESTADO en S y OBSERVACIONES en T)
  let nuevosEncabezados = [...encabezadosOriginales];
  while (nuevosEncabezados.length < 18) nuevosEncabezados.push("");
  nuevosEncabezados[18] = "ESTADO";        // Columna S (índice 18)
  nuevosEncabezados[19] = "OBSERVACIONES"; // Columna T (índice 19)

  const encabezadosNormalizados = encabezadosOriginales.map(h => normalizarTexto(h));
  const idxImpuestos = buscarIndice(encabezadosNormalizados, ["impuestos", "tipoimpuesto", "tipodeimpuesto"]);
  const idxMotivo = buscarIndice(encabezadosNormalizados, ["motivo", "motivodelasolicitud", "tipodesolicitud"]);

  if (idxImpuestos === -1 || idxMotivo === -1) {
    throw new Error("No se encontraron las columnas requeridas (Impuestos o Motivo).");
  }

  // ========== PROCESAR CADA HOJA DESTINO ==========
  CONFIG.forEach(cfg => {

    // 1. Leer la hoja destino actual y guardar estado/observaciones por radicado
    let hojaDestino = ss.getSheetByName(cfg.hoja);
    const estadosGuardados = {}; // { radicadoNormalizado: { estado, observaciones } }

    if (hojaDestino) {
      const datosDestino = hojaDestino.getDataRange().getValues();
      for (let i = 1; i < datosDestino.length; i++) {
        const rad = normalizarClave(datosDestino[i][IDX_RADICADO]);
        if (rad) {
          estadosGuardados[rad] = {
            estado: datosDestino[i][18] || ESTADO_POR_DEFECTO,
            observaciones: datosDestino[i][19] || ""
          };
        }
      }
    } else {
      hojaDestino = ss.insertSheet(cfg.hoja);
    }

    // 2. Construir las filas que SÍ deben estar en esta hoja según IMPORT actual
    const filasNuevas = [];

    filas.forEach(fila => {
      const motivo = (fila[idxMotivo] || "").toString().trim();
      if (!MOTIVOS_VALIDOS.includes(motivo)) return;

      const impuestosTexto = (fila[idxImpuestos] || "").toString();

      // ¿Esta fila pertenece a esta hoja?
      const perteneceAEstaHoja = cfg.textos.some(t => impuestosTexto.includes(t));
      if (!perteneceAEstaHoja) return;

      // Completar fila hasta columna T
      let filaCompleta = [...fila];
      while (filaCompleta.length < 18) filaCompleta.push("");

      // Buscar si esta fila ya existía → conservar estado y observaciones
      const radClave = normalizarClave(fila[IDX_RADICADO]);
      const guardado = estadosGuardados[radClave];

      if (guardado) {
        filaCompleta[18] = guardado.estado;
        filaCompleta[19] = guardado.observaciones;
      } else {
        filaCompleta[18] = ESTADO_POR_DEFECTO;
        filaCompleta[19] = "";
      }

      filasNuevas.push(filaCompleta);
    });

    // 3. Escribir la hoja desde cero (con datos preservados donde corresponde)
    hojaDestino.clearContents();
    hojaDestino.clearFormats();

    const dataFinal = [nuevosEncabezados, ...filasNuevas];

    if (dataFinal.length > 0) {
      hojaDestino.getRange(1, 1, dataFinal.length, dataFinal[0].length).setValues(dataFinal);

      // Formato encabezado
      const rangoHeader = hojaDestino.getRange(1, 1, 1, dataFinal[0].length);
      rangoHeader.setFontWeight("bold")
                 .setBackground("#ED1C27")
                 .setFontColor("white")
                 .setHorizontalAlignment("center");

      // Validación de datos en columna S (Estado) si hay filas
      if (filasNuevas.length > 0) {
        const rangoChips = hojaDestino.getRange(2, 19, filasNuevas.length, 1);
        const reglaEx = SpreadsheetApp.newDataValidation()
          .requireValueInList(ESTADOS_DISPONIBLES)
          .setAllowInvalid(false)
          .build();
        rangoChips.setDataValidation(reglaEx);

        // Colores por estado en columna S
        const reglas = [
          SpreadsheetApp.newConditionalFormatRule()
            .whenTextEqualTo("RECIBIDO EN PROCESO")
            .setBackground("#fff2cc").setFontColor("#7f6000")
            .setRanges([rangoChips]).build(),
          SpreadsheetApp.newConditionalFormatRule()
            .whenTextEqualTo("APROBADO")
            .setBackground("#d9ead3").setFontColor("#274e13")
            .setRanges([rangoChips]).build(),
          SpreadsheetApp.newConditionalFormatRule()
            .whenTextEqualTo("RECHAZADO")
            .setBackground("#f4cccc").setFontColor("#990000")
            .setRanges([rangoChips]).build(),
          SpreadsheetApp.newConditionalFormatRule()
            .whenTextEqualTo("REQUERIDO")
            .setBackground("#d9d2e9").setFontColor("#20124d")
            .setRanges([rangoChips]).build()
        ];
        hojaDestino.setConditionalFormatRules(reglas);

        // Columna T (Observaciones) con fondo neutro
        hojaDestino.getRange(2, 20, filasNuevas.length, 1).setBackground("#f9f9f9");
      }

      // Congelar fila de encabezados
      hojaDestino.setFrozenRows(1);
    }
  });

  Logger.log("Distribución completada preservando estados existentes.");
}

/**
 * Normaliza una clave (radicado) para comparación: quita espacios y ceros a la izquierda.
 */
function normalizarClave(valor) {
  if (valor === null || valor === undefined) return "";
  let s = valor.toString().trim().toLowerCase();
  s = s.replace(/^0+(?=\d)/, "");
  return s;
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

function normalizarBusqueda(valor) {
  if (valor === null || valor === undefined) return "";
  let s = valor.toString().trim().toLowerCase();
  s = s.replace(/^0+(?=\d)/, "");
  return s;
}

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