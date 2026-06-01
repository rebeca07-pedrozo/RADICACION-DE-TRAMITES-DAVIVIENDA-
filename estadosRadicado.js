/**
 * ============================================================================
 * DISTRIBUCIÓN AUTOMÁTICA DE SOLICITUDES + NOTIFICACIONES POR CORREO
 * ============================================================================
 */

// 👇 ÍNDICES DE COLUMNAS (0 = A, 1 = B, 2 = C...)
const IDX_RADICADO  = 1; // Columna B
const IDX_EMAIL     = 2; // Columna C (email del cliente)
const IDX_DOCUMENTO = 6; // Columna G (Cédula/NIT)

// Columnas que agregamos:
const COL_ESTADO        = 19; // Columna S (índice 18 en array, columna 19 en getRange)
const COL_OBSERVACIONES = 20; // Columna T
const COL_NOTIFICAR     = 21; // Columna U  ← NUEVA

// Hojas destino (las usamos en varios lugares)
const HOJAS_DESTINO = [
  "Marcación y Reintegro de retencion de ICA",
  "Marcación y Reintegro de retencion de renta",
  "Marcación y Reintegro de retencion de IVA",
  "Marcación y Reintegro de impuesto IVA"
];

const ESTADOS_DISPONIBLES = ["RECIBIDO EN PROCESO", "APROBADO", "RECHAZADO", "REQUERIDO"];
const NOTIFICAR_OPCIONES  = ["NO ENVIADO", "ENVIAR CORREO"];
const ESTADO_POR_DEFECTO  = "RECIBIDO EN PROCESO";
const NOTIFICAR_DEFECTO   = "NO ENVIADO";


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
    "Marcación", "Reintegro", "Ambas",
    "Desmarcación", "Certif. Régimen Simple", "Desistimiento"
  ];

  const hojaImport = ss.getSheetByName(HOJA_IMPORT);
  if (!hojaImport) throw new Error(`No existe la hoja '${HOJA_IMPORT}'`);

  const datos = hojaImport.getDataRange().getValues();
  if (datos.length <= 1) {
    Logger.log("No hay datos para procesar.");
    return;
  }

  const encabezadosOriginales = datos[0];
  const filas = datos.slice(1);

  // Construir encabezados finales (ESTADO, OBSERVACIONES, NOTIFICAR)
  let nuevosEncabezados = [...encabezadosOriginales];
  while (nuevosEncabezados.length < 18) nuevosEncabezados.push("");
  nuevosEncabezados[18] = "ESTADO";         // S
  nuevosEncabezados[19] = "OBSERVACIONES";  // T
  nuevosEncabezados[20] = "NOTIFICAR";      // U

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
            observaciones: datosDestino[i][19] || "",
            notificar: datosDestino[i][20] || NOTIFICAR_DEFECTO
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
        filaCompleta[20] = guardado.notificar;
      } else {
        filaCompleta[18] = ESTADO_POR_DEFECTO;
        filaCompleta[19] = "";
        filaCompleta[20] = NOTIFICAR_DEFECTO;
      }

      filasNuevas.push(filaCompleta);
    });

    hojaDestino.clearContents();
    hojaDestino.clearFormats();

    const dataFinal = [nuevosEncabezados, ...filasNuevas];

    if (dataFinal.length > 0) {
      hojaDestino.getRange(1, 1, dataFinal.length, dataFinal[0].length).setValues(dataFinal);

      // Encabezado rojo
      const rangoHeader = hojaDestino.getRange(1, 1, 1, dataFinal[0].length);
      rangoHeader.setFontWeight("bold")
                 .setBackground("#ED1C27")
                 .setFontColor("white")
                 .setHorizontalAlignment("center");

      if (filasNuevas.length > 0) {
        // ===== Columna S: ESTADO =====
        const rangoEstado = hojaDestino.getRange(2, COL_ESTADO, filasNuevas.length, 1);
        const reglaEstado = SpreadsheetApp.newDataValidation()
          .requireValueInList(ESTADOS_DISPONIBLES)
          .setAllowInvalid(false)
          .build();
        rangoEstado.setDataValidation(reglaEstado);

        // ===== Columna U: NOTIFICAR =====
        const rangoNotificar = hojaDestino.getRange(2, COL_NOTIFICAR, filasNuevas.length, 1);
        const reglaNotif = SpreadsheetApp.newDataValidation()
          .requireValueInList(NOTIFICAR_OPCIONES)
          .setAllowInvalid(true) // permitimos "ENVIADO ..." que el script escribe
          .build();
        rangoNotificar.setDataValidation(reglaNotif);

        // Formato condicional para ambas columnas
        const reglas = [
          // Estado
          SpreadsheetApp.newConditionalFormatRule()
            .whenTextEqualTo("RECIBIDO EN PROCESO")
            .setBackground("#fff2cc").setFontColor("#7f6000")
            .setRanges([rangoEstado]).build(),
          SpreadsheetApp.newConditionalFormatRule()
            .whenTextEqualTo("APROBADO")
            .setBackground("#d9ead3").setFontColor("#274e13")
            .setRanges([rangoEstado]).build(),
          SpreadsheetApp.newConditionalFormatRule()
            .whenTextEqualTo("RECHAZADO")
            .setBackground("#f4cccc").setFontColor("#990000")
            .setRanges([rangoEstado]).build(),
          SpreadsheetApp.newConditionalFormatRule()
            .whenTextEqualTo("REQUERIDO")
            .setBackground("#d9d2e9").setFontColor("#20124d")
            .setRanges([rangoEstado]).build(),

          // Notificar
          SpreadsheetApp.newConditionalFormatRule()
            .whenTextEqualTo("NO ENVIADO")
            .setBackground("#f4cccc").setFontColor("#990000")
            .setRanges([rangoNotificar]).build(),
          SpreadsheetApp.newConditionalFormatRule()
            .whenTextEqualTo("ENVIAR CORREO")
            .setBackground("#fce5cd").setFontColor("#783f04").setBold(true)
            .setRanges([rangoNotificar]).build(),
          SpreadsheetApp.newConditionalFormatRule()
            .whenTextContains("ENVIADO ")
            .setBackground("#d9ead3").setFontColor("#274e13")
            .setRanges([rangoNotificar]).build()
        ];
        hojaDestino.setConditionalFormatRules(reglas);

        // Observaciones fondo neutro
        hojaDestino.getRange(2, COL_OBSERVACIONES, filasNuevas.length, 1).setBackground("#f9f9f9");
      }

      hojaDestino.setFrozenRows(1);
    }
  });

  Logger.log("Distribución completada preservando estados existentes.");
}

// ===========================================================================
// TRIGGER ONEDIT: detecta cuando alguien selecciona "ENVIAR CORREO"
// ===========================================================================

function onEdit(e) {
  try {
    if (!e || !e.range) return;

    const sheet = e.range.getSheet();
    const nombreHoja = sheet.getName();

    // Solo nos interesan las 4 hojas destino
    if (HOJAS_DESTINO.indexOf(nombreHoja) === -1) return;

    const col = e.range.getColumn();
    const fila = e.range.getRow();

    // Solo si editan la columna U (NOTIFICAR) y no es el encabezado
    if (col !== COL_NOTIFICAR || fila < 2) return;

    const valorNuevo = (e.value || "").toString().trim();
    if (valorNuevo !== "ENVIAR CORREO") return;

    // Leer datos de la fila
    const datosFila = sheet.getRange(fila, 1, 1, COL_NOTIFICAR).getValues()[0];

    const radicado      = datosFila[IDX_RADICADO];
    const email         = (datosFila[IDX_EMAIL] || "").toString().trim();
    const documento     = datosFila[IDX_DOCUMENTO];
    const estado        = datosFila[18];
    const observaciones = datosFila[19];

    if (!email || email.indexOf("@") === -1) {
      sheet.getRange(fila, COL_NOTIFICAR).setValue("⚠ SIN EMAIL VÁLIDO");
      return;
    }

    // Determinar nombre del impuesto a partir del nombre de la hoja
    const impuesto = nombreHoja.replace("Marcación y Reintegro de ", "");

    // Enviar correo
    enviarCorreoNotificacion({
      destinatario: email,
      radicado: radicado,
      documento: documento,
      impuesto: impuesto,
      estado: estado,
      observaciones: observaciones
    });

    // Marcar como enviado con fecha y hora
    const ahora = new Date();
    const fecha = Utilities.formatDate(ahora, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
    sheet.getRange(fila, COL_NOTIFICAR).setValue(`✅ ENVIADO ${fecha}`);

  } catch (err) {
    Logger.log("Error en onEdit: " + err.message);
    try {
      e.range.setNote("Error al enviar correo: " + err.message);
    } catch (_) {}
  }
}

/**
 * Envía el correo de notificación al cliente.
 */
function enviarCorreoNotificacion(d) {
  const asunto = `Davivienda - Actualización de su trámite ${d.impuesto} - Radicado ${d.radicado}`;

  const colorEstado = {
    "APROBADO":  "#27ae60",
    "RECHAZADO": "#e74c3c",
    "REQUERIDO": "#8e44ad",
    "RECIBIDO EN PROCESO": "#f39c12"
  }[d.estado] || "#7f8c8d";

  const observacionesHtml = d.observaciones
    ? d.observaciones.toString().replace(/\n/g, "<br>")
    : "<em>Sin observaciones adicionales.</em>";

  const html = `
    <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; border: 1px solid #e0e0e0; border-radius: 10px; overflow: hidden;">
      <div style="background:#ED1C27; color:white; padding:25px; text-align:center;">
        <h1 style="margin:0; font-size:22px;">Davivienda</h1>
        <p style="margin:5px 0 0; font-size:14px;">Actualización de su trámite</p>
      </div>

      <div style="padding:30px; background:white;">
        <p style="font-size:15px; color:#2c3e50;">Estimado(a) cliente,</p>
        <p style="font-size:14px; color:#34495e; line-height:1.6;">
          Le informamos que su trámite de <strong>${d.impuesto}</strong> ha sido actualizado.
        </p>

        <table style="width:100%; border-collapse:collapse; margin:20px 0; font-size:14px;">
          <tr>
            <td style="padding:10px; border-bottom:1px solid #ecf0f1; color:#7f8c8d; width:40%;">Radicado</td>
            <td style="padding:10px; border-bottom:1px solid #ecf0f1; color:#2c3e50; font-weight:600;">${d.radicado}</td>
          </tr>
          <tr>
            <td style="padding:10px; border-bottom:1px solid #ecf0f1; color:#7f8c8d;">Documento</td>
            <td style="padding:10px; border-bottom:1px solid #ecf0f1; color:#2c3e50; font-weight:600;">${d.documento}</td>
          </tr>
          <tr>
            <td style="padding:10px; border-bottom:1px solid #ecf0f1; color:#7f8c8d;">Impuesto</td>
            <td style="padding:10px; border-bottom:1px solid #ecf0f1; color:#2c3e50; font-weight:600;">${d.impuesto}</td>
          </tr>
          <tr>
            <td style="padding:10px; border-bottom:1px solid #ecf0f1; color:#7f8c8d;">Estado actual</td>
            <td style="padding:10px; border-bottom:1px solid #ecf0f1;">
              <span style="background:${colorEstado}; color:white; padding:5px 12px; border-radius:4px; font-weight:600; font-size:12px;">${d.estado}</span>
            </td>
          </tr>
        </table>

        <div style="background:#f8f9fa; border-left:4px solid #ED1C27; padding:15px 20px; margin:20px 0; border-radius:4px;">
          <p style="margin:0 0 8px; font-size:12px; color:#7f8c8d; text-transform:uppercase; font-weight:600;">Observaciones</p>
          <p style="margin:0; font-size:14px; color:#2c3e50; line-height:1.5;">${observacionesHtml}</p>
        </div>

        <p style="font-size:13px; color:#7f8c8d; line-height:1.6; margin-top:25px;">
          Si tiene alguna duda sobre su trámite, puede consultar el estado en cualquier momento con su número de radicado o documento.
        </p>
      </div>

      <div style="background:#f4f7f6; padding:15px; text-align:center; font-size:11px; color:#95a5a6;">
        Este es un mensaje automático. Por favor no responda a este correo.
      </div>
    </div>
  `;

  MailApp.sendEmail({
    to: d.destinatario,
    subject: asunto,
    htmlBody: html
  });
}

// ===========================================================================
// FUNCIONES AUXILIARES
// ===========================================================================

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

// ===========================================================================
// WEB APP DE CONSULTA
// ===========================================================================

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

  if (!textoBusqueda || !textoBusqueda.toString().trim()) {
    return { exito: false, mensaje: "Por favor ingresa un radicado o documento válido." };
  }

  const busqueda = normalizarBusqueda(textoBusqueda);
  const coincidencias = [];

  for (let nombreHoja of HOJAS_DESTINO) {
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