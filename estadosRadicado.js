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

  const CONFIG = [
    {
      hoja: "Marcación y Reintegro de retencion de ICA",
      textos: ["Retención de ICA"]
    },
    {
      hoja: "Marcación y Reintegro de retencion de renta",
      textos: [
        "Retención de Renta",
        "JELPIT",
        "Propiedad horizontal",
        "Régimen simple"
      ]
    },
    {
      hoja: "Marcación y Reintegro de retencion de IVA",
      textos: ["Retención de IVA"]
    },
    {
      hoja: "Marcación y Reintegro de impuesto IVA",
      textos: ["Impuesto de IVA"]
    }
  ];

  const MOTIVOS_VALIDOS = [
    "Marcación",
    "Reintegro",
    "Ambas",
    "Desmarcación",
    "Certif. Régimen Simple",
    "Desistimiento"
  ];
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

  let nuevosEncabezados = [...encabezadosOriginales];
  while (nuevosEncabezados.length < 18) nuevosEncabezados.push("");
  nuevosEncabezados[18] = "ESTADO";
  nuevosEncabezados[19] = "OBSERVACIONES";

  const encabezadosNormalizados = encabezadosOriginales.map(h => normalizarTexto(h));
  const idxImpuestos = buscarIndice(encabezadosNormalizados, ["impuestos", "tipoimpuesto", "tipodeimpuesto"]);
  const idxMotivo = buscarIndice(encabezadosNormalizados, ["motivo", "motivodelasolicitud", "tipodesolicitud"]);

  if (idxImpuestos === -1 || idxMotivo === -1) {
    throw new Error("No se encontraron las columnas requeridas (Impuestos o Motivo).");
  }

  CONFIG.forEach(cfg => {

    let hojaDestino = ss.getSheetByName(cfg.hoja);
    const estadosGuardados = {};

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

    const filasNuevas = [];

    filas.forEach(fila => {
      const motivo = (fila[idxMotivo] || "").toString().trim();
      if (!MOTIVOS_VALIDOS.includes(motivo)) return;

      const impuestosTexto = (fila[idxImpuestos] || "").toString();
      const perteneceAEstaHoja = cfg.textos.some(t => impuestosTexto.includes(t));
      if (!perteneceAEstaHoja) return;

      let filaCompleta = [...fila];
      while (filaCompleta.length < 18) filaCompleta.push("");

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

    hojaDestino.clearContents();
    hojaDestino.clearFormats();

    const dataFinal = [nuevosEncabezados, ...filasNuevas];

    if (dataFinal.length > 0) {
      hojaDestino.getRange(1, 1, dataFinal.length, dataFinal[0].length).setValues(dataFinal);

      const rangoHeader = hojaDestino.getRange(1, 1, 1, dataFinal[0].length);
      rangoHeader.setFontWeight("bold")
                 .setBackground("#ED1C27")
                 .setFontColor("white")
                 .setHorizontalAlignment("center");

      if (filasNuevas.length > 0) {
        const rangoChips = hojaDestino.getRange(2, 19, filasNuevas.length, 1);
        const reglaEx = SpreadsheetApp.newDataValidation()
          .requireValueInList(ESTADOS_DISPONIBLES)
          .setAllowInvalid(false)
          .build();
        rangoChips.setDataValidation(reglaEx);

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

        hojaDestino.getRange(2, 20, filasNuevas.length, 1).setBackground("#f9f9f9");
      }

      hojaDestino.setFrozenRows(1);
    }
  });

  Logger.log("Distribución completada preservando estados existentes.");
}

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
  return HtmlService.createHtmlOutputFromFile('InterfazEstado')
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
    "Marcación y Reintegro de retencion de ICA",
    "Marcación y Reintegro de retencion de renta",
    "Marcación y Reintegro de retencion de IVA",
    "Marcación y Reintegro de impuesto IVA"
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
          impuesto: nombreHoja.replace("Marcación y Reintegro de ", ""),
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