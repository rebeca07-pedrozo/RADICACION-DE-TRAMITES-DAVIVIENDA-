/**
 * ============================================================================
 * DISTRIBUCIÓN AUTOMÁTICA + NOTIFICACIONES (CLIENTE + ÁREA REMITIDA)
 * ============================================================================
 */

// ============================================================================
// ÍNDICES DE COLUMNAS (0 = A, 1 = B, 2 = C...)
// ============================================================================

// Datos del formulario (columnas A-R, vienen del IMPORTRANGE)
const IDX_TIMESTAMP = 0;  // A
const IDX_RADICADO  = 1;  // B
const IDX_ENTIDAD   = 2;  // C  ← NUEVA
const IDX_EMAIL     = 3;  // D  ← (antes era C)
const IDX_DOCUMENTO = 7;  // H  ← (antes era G)

// Columnas operativas para CLIENTE: S, T, U
const COL_ESTADO        = 19; // S
const COL_OBSERVACIONES = 20; // T
const COL_NOTIFICAR     = 21; // U

// Columnas operativas para ÁREA REMITIDA: V, W, X (nuevas)
const COL_CORREO_AREA    = 22; // V
const COL_OBSERVAC_AREA  = 23; // W
const COL_NOTIF_AREA     = 24; // X

// Índices en array (col-1) para lectura
const IDX_ESTADO         = COL_ESTADO - 1;        // 18
const IDX_OBSERVACIONES  = COL_OBSERVACIONES - 1; // 19
const IDX_NOTIFICAR      = COL_NOTIFICAR - 1;     // 20
const IDX_CORREO_AREA    = COL_CORREO_AREA - 1;   // 21
const IDX_OBSERVAC_AREA  = COL_OBSERVAC_AREA - 1; // 22
const IDX_NOTIF_AREA     = COL_NOTIF_AREA - 1;    // 23

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


// ============================================================================
// DISTRIBUIR SOLICITUDES
// ============================================================================

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
      textos: ["Retención de Renta", "JELPIT", "Propiedad horizontal", "Régimen simple"]
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
  if (datos.length <= 1) return;

  const encabezadosOriginales = datos[0];
  const filas = datos.slice(1);

  // Construir encabezados finales
  let nuevosEncabezados = [...encabezadosOriginales];
  // Rellenar hasta el índice 18 (columna S) por si IMPORT trae menos columnas
  while (nuevosEncabezados.length < IDX_ESTADO) nuevosEncabezados.push("");
  nuevosEncabezados[IDX_ESTADO]        = "ESTADO";
  nuevosEncabezados[IDX_OBSERVACIONES] = "OBSERVACIONES";
  nuevosEncabezados[IDX_NOTIFICAR]     = "NOTIFICAR";
  nuevosEncabezados[IDX_CORREO_AREA]   = "CORREO_AREA";
  nuevosEncabezados[IDX_OBSERVAC_AREA] = "OBSERVACIONES_AREA";
  nuevosEncabezados[IDX_NOTIF_AREA]    = "NOTIFICAR_AREA";

  const encabezadosNormalizados = encabezadosOriginales.map(h => normalizarTexto(h));
  const idxImpuestos = buscarIndice(encabezadosNormalizados, ["impuestos", "tipoimpuesto", "tipodeimpuesto"]);
  const idxMotivo    = buscarIndice(encabezadosNormalizados, ["motivo", "motivodelasolicitud", "tipodesolicitud"]);

  if (idxImpuestos === -1 || idxMotivo === -1) {
    throw new Error("No se encontraron las columnas 'Impuestos' o 'Motivo' en IMPORT.");
  }

  CONFIG.forEach(cfg => {
    let hojaDestino = ss.getSheetByName(cfg.hoja);
    const guardados = {};

    // Preservar todo lo escrito a mano en S-X
    if (hojaDestino) {
      const datosDestino = hojaDestino.getDataRange().getValues();
      for (let i = 1; i < datosDestino.length; i++) {
        const rad = normalizarClave(datosDestino[i][IDX_RADICADO]);
        if (rad) {
          guardados[rad] = {
            estado:        datosDestino[i][IDX_ESTADO]        || ESTADO_POR_DEFECTO,
            observaciones: datosDestino[i][IDX_OBSERVACIONES] || "",
            notificar:     datosDestino[i][IDX_NOTIFICAR]     || NOTIFICAR_DEFECTO,
            correoArea:    datosDestino[i][IDX_CORREO_AREA]    || "",
            observAreas:   datosDestino[i][IDX_OBSERVAC_AREA]  || "",
            notifArea:     datosDestino[i][IDX_NOTIF_AREA]     || NOTIFICAR_DEFECTO
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
      if (!cfg.textos.some(t => impuestosTexto.includes(t))) return;

      let f = [...fila];
      while (f.length < IDX_ESTADO) f.push("");

      const radClave = normalizarClave(fila[IDX_RADICADO]);
      const g = guardados[radClave];

      if (g) {
        f[IDX_ESTADO]        = g.estado;
        f[IDX_OBSERVACIONES] = g.observaciones;
        f[IDX_NOTIFICAR]     = g.notificar;
        f[IDX_CORREO_AREA]   = g.correoArea;
        f[IDX_OBSERVAC_AREA] = g.observAreas;
        f[IDX_NOTIF_AREA]    = g.notifArea;
      } else {
        f[IDX_ESTADO]        = ESTADO_POR_DEFECTO;
        f[IDX_OBSERVACIONES] = "";
        f[IDX_NOTIFICAR]     = NOTIFICAR_DEFECTO;
        f[IDX_CORREO_AREA]   = "";
        f[IDX_OBSERVAC_AREA] = "";
        f[IDX_NOTIF_AREA]    = NOTIFICAR_DEFECTO;
      }

      filasNuevas.push(f);
    });

    hojaDestino.clearContents();
    hojaDestino.clearFormats();

    const dataFinal = [nuevosEncabezados, ...filasNuevas];

    if (dataFinal.length > 0) {
      hojaDestino.getRange(1, 1, dataFinal.length, dataFinal[0].length).setValues(dataFinal);

      // Encabezado
      hojaDestino.getRange(1, 1, 1, dataFinal[0].length)
        .setFontWeight("bold").setBackground("#ED1C27").setFontColor("white")
        .setHorizontalAlignment("center");

      if (filasNuevas.length > 0) {
        aplicarValidacionesYFormatos(hojaDestino, filasNuevas.length);
      }

      hojaDestino.setFrozenRows(1);
    }
  });

  Logger.log("Distribución completada.");
}

function aplicarValidacionesYFormatos(hoja, numFilas) {
  // Validación columna S (Estado)
  const rEstado = hoja.getRange(2, COL_ESTADO, numFilas, 1);
  rEstado.setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(ESTADOS_DISPONIBLES).setAllowInvalid(false).build()
  );

  // Validación columna U (Notificar cliente)
  const rNotif = hoja.getRange(2, COL_NOTIFICAR, numFilas, 1);
  rNotif.setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(NOTIFICAR_OPCIONES).setAllowInvalid(true).build()
  );

  // Validación columna X (Notificar área)
  const rNotifArea = hoja.getRange(2, COL_NOTIF_AREA, numFilas, 1);
  rNotifArea.setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(NOTIFICAR_OPCIONES).setAllowInvalid(true).build()
  );

  // Formato condicional
  const reglas = [
    // ESTADO
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("RECIBIDO EN PROCESO")
      .setBackground("#fff2cc").setFontColor("#7f6000").setRanges([rEstado]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("APROBADO")
      .setBackground("#d9ead3").setFontColor("#274e13").setRanges([rEstado]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("RECHAZADO")
      .setBackground("#f4cccc").setFontColor("#990000").setRanges([rEstado]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("REQUERIDO")
      .setBackground("#d9d2e9").setFontColor("#20124d").setRanges([rEstado]).build(),

    // NOTIFICAR CLIENTE
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("NO ENVIADO")
      .setBackground("#f4cccc").setFontColor("#990000").setRanges([rNotif]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("ENVIAR CORREO")
      .setBackground("#fce5cd").setFontColor("#783f04").setBold(true).setRanges([rNotif]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextContains("ENVIADO ")
      .setBackground("#d9ead3").setFontColor("#274e13").setRanges([rNotif]).build(),

    // NOTIFICAR ÁREA
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("NO ENVIADO")
      .setBackground("#f4cccc").setFontColor("#990000").setRanges([rNotifArea]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("ENVIAR CORREO")
      .setBackground("#cfe2f3").setFontColor("#073763").setBold(true).setRanges([rNotifArea]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextContains("ENVIADO ")
      .setBackground("#d9ead3").setFontColor("#274e13").setRanges([rNotifArea]).build()
  ];
  hoja.setConditionalFormatRules(reglas);

  // Fondos suaves para columnas de observaciones y correo área
  hoja.getRange(2, COL_OBSERVACIONES, numFilas, 1).setBackground("#f9f9f9");
  hoja.getRange(2, COL_CORREO_AREA,    numFilas, 1).setBackground("#e3f0fb");
  hoja.getRange(2, COL_OBSERVAC_AREA,  numFilas, 1).setBackground("#f0f8ff");
}


// ============================================================================
// TRIGGER ONEDIT: detecta cambios en U (cliente) o X (área)
// ============================================================================

function alEditarHoja(e) {
  try {
    if (!e || !e.range) return;

    const sheet = e.range.getSheet();
    if (HOJAS_DESTINO.indexOf(sheet.getName()) === -1) return;

    const col  = e.range.getColumn();
    const fila = e.range.getRow();
    if (fila < 2) return;

    const valorNuevo = (e.value || "").toString().trim();
    if (valorNuevo !== "ENVIAR CORREO") return;

    // Si editaron U → notificar al cliente
    if (col === COL_NOTIFICAR) {
      enviarCorreoCliente(sheet, fila);
      return;
    }

    // Si editaron X → notificar al área
    if (col === COL_NOTIF_AREA) {
      enviarCorreoArea(sheet, fila);
      return;
    }

  } catch (err) {
    Logger.log("Error en alEditarHoja: " + err.message);
    try { e.range.setNote("Error: " + err.message); } catch (_) {}
  }
}


// ============================================================================
// ENVÍO AL CLIENTE
// ============================================================================

function enviarCorreoCliente(sheet, fila) {
  const datosFila = sheet.getRange(fila, 1, 1, COL_NOTIF_AREA).getValues()[0];
  const nombreHoja = sheet.getName();

  const email = (datosFila[IDX_EMAIL] || "").toString().trim();
  if (!email || email.indexOf("@") === -1) {
    sheet.getRange(fila, COL_NOTIFICAR).setValue("⚠ SIN EMAIL VÁLIDO");
    return;
  }

  // Leer encabezados de IMPORT para asociar nombre↔valor
  const datosCorreo = construirDatosFormulario(datosFila, nombreHoja);
  datosCorreo.estado        = datosFila[IDX_ESTADO];
  datosCorreo.observaciones = datosFila[IDX_OBSERVACIONES];

  const html = construirCorreoClienteHTML(datosCorreo);

  MailApp.sendEmail({
    to: email,
    subject: `Davivienda — Actualización trámite ${datosCorreo.impuesto} — Radicado ${datosCorreo.radicado}`,
    htmlBody: html
  });

  marcarEnviado(sheet, fila, COL_NOTIFICAR);
}


// ============================================================================
// ENVÍO AL ÁREA REMITIDA
// ============================================================================

function enviarCorreoArea(sheet, fila) {
  const datosFila = sheet.getRange(fila, 1, 1, COL_NOTIF_AREA).getValues()[0];
  const nombreHoja = sheet.getName();

  const correoAreaRaw = (datosFila[IDX_CORREO_AREA] || "").toString().trim();
  if (!correoAreaRaw) {
    sheet.getRange(fila, COL_NOTIF_AREA).setValue("⚠ SIN CORREO ÁREA");
    return;
  }

  // Parsear múltiples correos separados por coma o punto y coma
  const correos = correoAreaRaw
    .split(/[,;]/)
    .map(c => c.trim())
    .filter(c => c.indexOf("@") !== -1);

  if (correos.length === 0) {
    sheet.getRange(fila, COL_NOTIF_AREA).setValue("⚠ CORREO INVÁLIDO");
    return;
  }

  const datosCorreo = construirDatosFormulario(datosFila, nombreHoja);
  datosCorreo.estado            = datosFila[IDX_ESTADO];
  datosCorreo.observaciones     = datosFila[IDX_OBSERVACIONES];
  datosCorreo.observacionesArea = datosFila[IDX_OBSERVAC_AREA];

  const html = construirCorreoAreaHTML(datosCorreo);

  MailApp.sendEmail({
    to: correos.join(","),
    subject: `[Remisión Interna] Trámite ${datosCorreo.impuesto} — Radicado ${datosCorreo.radicado}`,
    htmlBody: html
  });

  marcarEnviado(sheet, fila, COL_NOTIF_AREA);
}


// ============================================================================
// HELPER: construir objeto con datos del formulario para el correo
// ============================================================================

function construirDatosFormulario(datosFila, nombreHoja) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaImport = ss.getSheetByName("IMPORT");
  let encabezados = [];
  if (hojaImport && hojaImport.getLastColumn() > 0) {
    encabezados = hojaImport.getRange(1, 1, 1, hojaImport.getLastColumn()).getValues()[0];
  }

  const obj = {};
  encabezados.forEach((h, i) => {
    obj[normalizarTexto(h)] = datosFila[i];
  });

  // Atajos cómodos para la plantilla
  obj.radicado    = datosFila[IDX_RADICADO];
  obj.entidad     = datosFila[IDX_ENTIDAD];
  obj.documento   = datosFila[IDX_DOCUMENTO];
  obj.impuesto    = nombreHoja.replace("Marcación y Reintegro de ", "");
  obj.timestampStr = datosFila[IDX_TIMESTAMP] instanceof Date
    ? Utilities.formatDate(datosFila[IDX_TIMESTAMP], "America/Bogota", "dd/MM/yyyy HH:mm")
    : (datosFila[IDX_TIMESTAMP] || "").toString();

  return obj;
}


// ============================================================================
// HELPER: marcar como enviado (fecha y hora)
// ============================================================================

function marcarEnviado(sheet, fila, columna) {
  const fecha = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
  sheet.getRange(fila, columna).setValue(`ENVIADO ${fecha}`);
}


// ============================================================================
// PLANTILLA HTML — CORREO AL CLIENTE (estilo Davivienda completo)
// ============================================================================

function construirCorreoClienteHTML(d) {
  const colorEstado = {
    "APROBADO":  "#27ae60",
    "RECHAZADO": "#e74c3c",
    "REQUERIDO": "#8e44ad",
    "RECIBIDO EN PROCESO": "#f39c12"
  }[d.estado] || "#7f8c8d";

  const fila = (etiqueta, valor) => {
    if (!valor && valor !== 0) return "";
    return `<tr>
      <td style="padding:8px 12px;background:#f9fafb;color:#6b7280;font-size:13px;width:42%;border-bottom:1px solid #e5e7eb;"><strong>${etiqueta}</strong></td>
      <td style="padding:8px 12px;color:#1f2937;font-size:13px;border-bottom:1px solid #e5e7eb;">${valor}</td>
    </tr>`;
  };

  const obsHtml = d.observaciones
    ? d.observaciones.toString().replace(/\n/g, "<br>")
    : "<em>Sin observaciones adicionales.</em>";

  return `<!DOCTYPE html>
<html><body style="margin:0;padding:0;background:#f5f6fa;font-family:Arial,sans-serif;">
<table width="100%" cellpadding="0" cellspacing="0" style="background:#f5f6fa;padding:20px 0;">
  <tr><td align="center">
    <table width="600" cellpadding="0" cellspacing="0" style="background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,0.08);">

      <!-- Cabecera roja -->
      <tr><td style="background:#E1251B;padding:24px 32px;color:#fff;">
        <div style="font-size:12px;letter-spacing:1px;opacity:0.9;">BANCO DAVIVIENDA</div>
        <div style="font-size:20px;font-weight:bold;margin-top:4px;">Actualización de su solicitud</div>
      </td></tr>

      <!-- Número de radicado destacado -->
      <tr><td style="padding:24px 32px 8px;">
        <div style="text-align:center;padding:20px;background:#FCEBEA;border-radius:8px;border:2px dashed #E1251B;">
          <div style="font-size:13px;color:#B81E15;letter-spacing:1px;">RADICADO</div>
          <div style="font-size:32px;font-weight:bold;color:#B81E15;margin-top:8px;letter-spacing:2px;">#${d.radicado}</div>
          <div style="margin-top:14px;">
            <span style="background:${colorEstado};color:white;padding:6px 16px;border-radius:4px;font-weight:600;font-size:13px;letter-spacing:0.5px;">${d.estado || ""}</span>
          </div>
        </div>
      </td></tr>

      <!-- Saludo -->
      <tr><td style="padding:16px 32px;">
        <p style="margin:0;font-size:14px;color:#1f2937;line-height:1.6;">
        Estimado(a) <strong>${d.razonsocial || "solicitante"}</strong>,<br><br>
        Le informamos que su trámite de <strong>${d.impuesto}</strong> ha sido actualizado.
        A continuación encontrará el detalle.
        </p>
      </td></tr>

      <!-- Tabla resumen -->
      <tr><td style="padding:0 32px 16px;">
        <div style="font-size:13px;font-weight:bold;color:#E1251B;text-transform:uppercase;margin-bottom:8px;letter-spacing:0.5px;">Resumen de la solicitud</div>
        <table width="100%" cellpadding="0" cellspacing="0" style="border:1px solid #e5e7eb;border-radius:6px;overflow:hidden;">
          ${fila("Entidad", d.entidad)}
          ${fila("Fecha solicitud", d.timestampStr)}
          ${fila("Oficina solicitante", d.nombreoficina)}
          ${fila("Código de oficina", d.codigooficina)}
          ${fila("Razón social", d.razonsocial)}
          ${fila("Cédula/NIT", (d.cedula || "") + (d.dv ? " - DV " + d.dv : ""))}
          ${fila("Motivo", d.motivo)}
          ${fila("Tipo de producto", d.tipoproducto)}
          ${fila("N° de producto", d.numeroproducto)}
          ${fila("Impuestos", d.impuestos)}
          ${fila("Año a reintegrar", d.periodo)}
          ${fila("Valor a reintegrar", d.valor)}
          ${fila("Ciudad", d.ciudad)}
        </table>
      </td></tr>

      <!-- Observaciones -->
      <tr><td style="padding:0 32px 16px;">
        <div style="font-size:13px;font-weight:bold;color:#E1251B;text-transform:uppercase;margin-bottom:8px;letter-spacing:0.5px;">Observaciones del trámite</div>
        <div style="padding:14px 16px;background:#f8f9fa;border-left:4px solid #E1251B;border-radius:4px;font-size:13.5px;color:#1f2937;line-height:1.6;">
          ${obsHtml}
        </div>
      </td></tr>

      <!-- Nota azul -->
      <tr><td style="padding:8px 32px 24px;">
        <div style="padding:14px;background:#eff6ff;border-left:3px solid #2563eb;border-radius:4px;font-size:12.5px;color:#1e3a8a;line-height:1.6;">
          Si tiene alguna duda sobre su trámite, comuníquese con su oficina indicando el número de radicado <strong>#${d.radicado || ""}</strong>.
        </div>
      </td></tr>

      <!-- Pie -->
      <tr><td style="padding:16px 32px;background:#f9fafb;border-top:1px solid #e5e7eb;font-size:11px;color:#6b7280;text-align:center;">
        Este correo se generó automáticamente. Por favor no responda a este mensaje.<br>
        <strong style="color:#E1251B;">Banco Davivienda S.A.</strong>
      </td></tr>

    </table>
  </td></tr>
</table>
</body></html>`;
}



// ============================================================================
// PLANTILLA HTML — CORREO AL ÁREA REMITIDA (formal/técnico)
// ============================================================================

function construirCorreoAreaHTML(d) {
  const fila = (etiqueta, valor) => {
    if (!valor && valor !== 0) return "";
    return `<tr>
      <td style="padding:8px 12px;background:#f1f5f9;color:#475569;font-size:13px;width:38%;border-bottom:1px solid #e2e8f0;font-weight:600;">${etiqueta}</td>
      <td style="padding:8px 12px;color:#0f172a;font-size:13px;border-bottom:1px solid #e2e8f0;">${valor}</td>
    </tr>`;
  };

  const obsAreaHtml = d.observacionesArea
    ? d.observacionesArea.toString().replace(/\n/g, "<br>")
    : "<em>Sin instrucciones específicas para el área remitida.</em>";

  const obsHtml = d.observaciones
    ? d.observaciones.toString().replace(/\n/g, "<br>")
    : "<em>Sin observaciones registradas.</em>";

  return `<!DOCTYPE html>
<html><body style="margin:0;padding:0;background:#eef2f7;font-family:Arial,sans-serif;">
<table width="100%" cellpadding="0" cellspacing="0" style="background:#eef2f7;padding:20px 0;">
  <tr><td align="center">
    <table width="640" cellpadding="0" cellspacing="0" style="background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 2px 10px rgba(0,0,0,0.08);">

      <!-- Cabecera azul oscuro -->
      <tr><td style="background:#E1251B;padding:22px 32px;color:#fff;">
        <div style="font-size:11px;letter-spacing:2px;opacity:0.85;">REMISIÓN INTERNA — BANCO DAVIVIENDA</div>
        <div style="font-size:19px;font-weight:bold;margin-top:6px;">Trámite remitido para gestión</div>
      </td></tr>

      <!-- Aviso -->
      <tr><td style="padding:20px 32px 10px;">
        <div style="padding:14px 18px;background:#fff7ed;border-left:4px solid #ea580c;border-radius:4px;font-size:13px;color:#7c2d12;line-height:1.6;">
          <strong>Acción requerida:</strong> Se ha remitido el siguiente trámite a su área para gestión. Por favor revise el detalle y las instrucciones al pie del correo.
        </div>
      </td></tr>

      <!-- Datos del trámite -->
      <tr><td style="padding:6px 32px 14px;">
        <div style="font-size:12px;font-weight:bold;color:#0f3a6b;text-transform:uppercase;margin-bottom:8px;letter-spacing:1px;">Identificación del trámite</div>
        <table width="100%" cellpadding="0" cellspacing="0" style="border:1px solid #e2e8f0;border-radius:6px;overflow:hidden;">
          ${fila("Radicado", "<strong>#" + (d.radicado || "") + "</strong>")}
          ${fila("Estado actual", d.estado)}
          ${fila("Tipo de impuesto", d.impuesto)}
          ${fila("Entidad", d.entidad)}
          ${fila("Fecha solicitud", d.timestampStr)}
        </table>
      </td></tr>

      <!-- Datos del solicitante -->
      <tr><td style="padding:6px 32px 14px;">
        <div style="font-size:12px;font-weight:bold;color:#0f3a6b;text-transform:uppercase;margin-bottom:8px;letter-spacing:1px;">Datos del solicitante</div>
        <table width="100%" cellpadding="0" cellspacing="0" style="border:1px solid #e2e8f0;border-radius:6px;overflow:hidden;">
          ${fila("Razón social", d.razonsocial)}
          ${fila("Cédula/NIT", (d.cedula || "") + (d.dv ? " - DV " + d.dv : ""))}
          ${fila("Oficina solicitante", d.nombreoficina)}
          ${fila("Código de oficina", d.codigooficina)}
          ${fila("Correo del solicitante", d.emailaddress)}
        </table>
      </td></tr>

      <!-- Detalle del caso -->
      <tr><td style="padding:6px 32px 14px;">
        <div style="font-size:12px;font-weight:bold;color:#0f3a6b;text-transform:uppercase;margin-bottom:8px;letter-spacing:1px;">Detalle del caso</div>
        <table width="100%" cellpadding="0" cellspacing="0" style="border:1px solid #e2e8f0;border-radius:6px;overflow:hidden;">
          ${fila("Motivo", d.motivo)}
          ${fila("Tipo de producto", d.tipoproducto)}
          ${fila("N° de producto", d.numeroproducto)}
          ${fila("Impuestos solicitados", d.impuestos)}
          ${fila("Año a reintegrar", d.periodo)}
          ${fila("Valor a reintegrar", d.valor)}
          ${fila("Ciudad", d.ciudad)}
          ${fila("Descripción del motivo", d.descripcionmotivo)}
        </table>
      </td></tr>

      <!-- Observaciones del compañero al área -->
      <tr><td style="padding:6px 32px 14px;">
        <div style="font-size:12px;font-weight:bold;color:#0f3a6b;text-transform:uppercase;margin-bottom:8px;letter-spacing:1px;">Instrucciones para el área remitida</div>
        <div style="padding:14px 18px;background:#eff6ff;border-left:4px solid #2563eb;border-radius:4px;font-size:13.5px;color:#0f172a;line-height:1.6;">
          ${obsAreaHtml}
        </div>
      </td></tr>

      <!-- Observaciones del trámite (contexto adicional) -->
      <tr><td style="padding:6px 32px 18px;">
        <div style="font-size:12px;font-weight:bold;color:#0f3a6b;text-transform:uppercase;margin-bottom:8px;letter-spacing:1px;">Observaciones generales del trámite</div>
        <div style="padding:14px 18px;background:#f8fafc;border:1px solid #e2e8f0;border-radius:4px;font-size:13px;color:#334155;line-height:1.6;">
          ${obsHtml}
        </div>
      </td></tr>

      <!-- Pie -->
      <tr><td style="padding:16px 32px;background:#f8fafc;border-top:1px solid #e2e8f0;font-size:11px;color:#64748b;text-align:center;">
        Comunicación interna generada automáticamente desde el sistema de gestión de marcaciones y reintegros.<br>
        <strong style="color:#0f3a6b;">Banco Davivienda S.A.</strong>
      </td></tr>

    </table>
  </td></tr>
</table>
</body></html>`;
}

// ============================================================================
// FUNCIONES AUXILIARES
// ============================================================================

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


// ============================================================================
// WEB APP DE CONSULTA
// ============================================================================

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
          estado: fila[IDX_ESTADO] || "RECIBIDO EN PROCESO",
          observaciones: fila[IDX_OBSERVACIONES] || "Sin observaciones registradas."
        });
      }
    }
  }

  if (coincidencias.length === 0) {
    return { exito: false, mensaje: "No se encontró ningún trámite con ese radicado o documento." };
  }

  return { exito: true, resultados: coincidencias };
}