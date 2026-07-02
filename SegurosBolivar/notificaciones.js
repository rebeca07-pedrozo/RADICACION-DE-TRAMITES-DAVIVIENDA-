const IDX_RAD_SB            = 2;  //C 
const IDX_COD_EXO_SB        = 3;  // D
const IDX_MUNI_SB           = 4;  // E
const IDX_TIPONOV_SB        = 5;  // F
const IDX_REIDESDE_SB       = 6;  // G
const IDX_REIHASTA_SB       = 7;  // H
const IDX_DPTO_SB           = 8;  // I
const IDX_ANIO_SB           = 9;  // J
const IDX_MOTIVO_SB         = 10; // K
const IDX_IDCLI_SB          = 11; // L
const IDX_DV_SB             = 12; // M
const IDX_NOMBRE_SB         = 13; // N
const IDX_CONCEPTO_SB       = 14; // O
const IDX_CORREO_SB         = 15; // P


function alRecibirRespuestaSB(e) {
  try {
    const sheet = e.range.getSheet();
    const fila = e.range.getRow();
    const numCols = sheet.getLastColumn();
    const datosFila = sheet.getRange(fila, 1, 1, numCols).getValues()[0];

    // 1. Radicado automático (si viene vacío)
    let radicado = (datosFila[IDX_RAD_SB] || "").toString().trim();
    if (!radicado) {
      radicado = generarRadicadoSB();
      sheet.getRange(fila, IDX_RAD_SB + 1).setValue(radicado);
    }

    // 2. Enviar correo de confirmación
    const correo = (datosFila[IDX_CORREO_SB] || "").toString().trim();
    if (!correo || correo.indexOf("@") === -1) {
      Logger.log("Radicado " + radicado + ": sin correo válido, no se envía confirmación.");
      return;
    }

    const datos = {
      radicado:        radicado,
      codigoExonerar:  datosFila[IDX_COD_EXO_SB],
      municipio:       datosFila[IDX_MUNI_SB],
      tipoNovedad:     datosFila[IDX_TIPONOV_SB],
      reintegrarDesde: formatearFechaSB(datosFila[IDX_REIDESDE_SB]),
      reintegrarHasta: formatearFechaSB(datosFila[IDX_REIHASTA_SB]),
      dpto:            datosFila[IDX_DPTO_SB],
      anioReintegro:   datosFila[IDX_ANIO_SB],
      motivo:          datosFila[IDX_MOTIVO_SB],
      idCliente:       datosFila[IDX_IDCLI_SB],
      dv:              datosFila[IDX_DV_SB],
      nombreCliente:   datosFila[IDX_NOMBRE_SB],
      concepto:        datosFila[IDX_CONCEPTO_SB]
    };

    const html = construirCorreoSegurosBolivarHTML(datos);

    MailApp.sendEmail({
      to: correo,
      subject: "Davivienda — Confirmación de radicación " + radicado + " (Seguros Bolívar)",
      htmlBody: html,
      name: "Marcaciones y Reintegros - Davivienda",
      noReply: true
    });

    Logger.log("Radicado " + radicado + " generado y correo enviado a " + correo);

  } catch (err) {
    Logger.log("Error en alRecibirRespuestaSB: " + err.message);
  }
}

function generarRadicadoSB() {
  const props = PropertiesService.getScriptProperties();
  const actual = parseInt(props.getProperty("ultimoRadicadoSB") || "0", 10);
  const siguiente = actual + 1;
  props.setProperty("ultimoRadicadoSB", siguiente.toString());
  return "RADICADO-" + siguiente;
}

function inicializarContadorSB() {
  PropertiesService.getScriptProperties().setProperty("ultimoRadicadoSB", "336");
  Logger.log("Contador inicializado. El próximo radicado será RADICADO-337.");
}

function formatearFechaSB(valor) {
  if (valor instanceof Date) {
    return Utilities.formatDate(valor, "America/Bogota", "dd/MM/yyyy");
  }
  return (valor || "").toString();
}

function construirCorreoSegurosBolivarHTML(d) {
  const fila = (etiqueta, valor) => {
    if (!valor && valor !== 0) return "";
    return `<tr>
      <td style="padding:8px 12px;background:#f9fafb;color:#6b7280;font-size:13px;width:42%;border-bottom:1px solid #e5e7eb;"><strong>${etiqueta}</strong></td>
      <td style="padding:8px 12px;color:#1f2937;font-size:13px;border-bottom:1px solid #e5e7eb;">${valor}</td>
    </tr>`;
  };

  return `<!DOCTYPE html>
<html><body style="margin:0;padding:0;background:#f5f6fa;font-family:Arial,sans-serif;">
<table width="100%" cellpadding="0" cellspacing="0" style="background:#f5f6fa;padding:20px 0;">
  <tr><td align="center">
    <table width="600" cellpadding="0" cellspacing="0" style="background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,0.08);">

      <tr><td style="background:#E1251B;padding:24px 32px;color:#fff;">
        <div style="font-size:12px;letter-spacing:1px;opacity:0.9;">BANCO DAVIVIENDA</div>
        <div style="font-size:20px;font-weight:bold;margin-top:4px;">Solicitud radicada exitosamente</div>
      </td></tr>

      <tr><td style="padding:24px 32px 8px;">
        <div style="text-align:center;padding:20px;background:#FCEBEA;border-radius:8px;border:2px dashed #E1251B;">
          <div style="font-size:13px;color:#B81E15;letter-spacing:1px;">SU NÚMERO DE RADICADO</div>
          <div style="font-size:30px;font-weight:bold;color:#B81E15;margin-top:8px;letter-spacing:1px;">${d.radicado}</div>
          <div style="font-size:12px;color:#6b7280;margin-top:8px;">Guarde este número para consultar el estado de su solicitud.</div>
        </div>
      </td></tr>

      <tr><td style="padding:16px 32px;">
        <p style="margin:0;font-size:14px;color:#1f2937;line-height:1.6;">
        Estimado(a) <strong>${d.nombreCliente || "solicitante"}</strong>,<br><br>
        Hemos recibido su solicitud. A continuación el resumen de su radicación.
        </p>
      </td></tr>

      <tr><td style="padding:0 32px 16px;">
        <div style="font-size:13px;font-weight:bold;color:#E1251B;text-transform:uppercase;margin-bottom:8px;letter-spacing:0.5px;">Resumen de la solicitud</div>
        <table width="100%" cellpadding="0" cellspacing="0" style="border:1px solid #e5e7eb;border-radius:6px;overflow:hidden;">
          ${fila("ID Cliente", (d.idCliente || "") + (d.dv ? " - DV " + d.dv : ""))}
          ${fila("Motivo", d.motivo)}
          ${fila("Tipo de novedad", d.tipoNovedad)}
          ${fila("Concepto", d.concepto)}
          ${fila("Código único a exonerar", d.codigoExonerar)}
          ${fila("Municipio", d.municipio)}
          ${fila("Departamento (ReteICA)", d.dpto)}
          ${fila("Año a reintegrar", d.anioReintegro)}
          ${fila("Reintegrar desde", d.reintegrarDesde)}
          ${fila("Reintegrar hasta", d.reintegrarHasta)}
        </table>
      </td></tr>

      <tr><td style="padding:16px 32px 24px;">
        <div style="padding:14px;background:#eff6ff;border-left:3px solid #2563eb;border-radius:4px;font-size:12.5px;color:#1e3a8a;line-height:1.6;">
          Para cualquier consulta sobre el estado de su solicitud, indique el número de radicado <strong>${d.radicado}</strong>.
        </div>
      </td></tr>

      <tr><td style="padding:16px 32px;background:#f9fafb;border-top:1px solid #e5e7eb;font-size:11px;color:#6b7280;text-align:center;">
        Este correo se generó automáticamente. Por favor no responda a este mensaje.<br>
        <strong style="color:#E1251B;">Banco Davivienda S.A.</strong>
      </td></tr>

    </table>
  </td></tr>
</table>
</body></html>`;
}