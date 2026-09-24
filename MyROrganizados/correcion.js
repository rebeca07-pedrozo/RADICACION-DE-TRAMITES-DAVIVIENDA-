//  EDITA AQUÍ
// Es lo que está entre /d/ y /view en el link de compartir de la imagen
const ID_LOGO_CORREOS = "PEGA_AQUI_EL_ID_CORTO_DEL_LOGO";

//enviarCorreoCliente (reemplázala completa)
function enviarCorreoCliente(sheet, fila) {
  const datosFila = sheet.getRange(fila, 1, 1, COL_NOTIF_AREA).getValues()[0];
  const nombreHoja = sheet.getName();

  const email = (datosFila[IDX_EMAIL] || "").toString().trim();
  if (!email || email.indexOf("@") === -1) {
    sheet.getRange(fila, COL_NOTIFICAR).setValue("⚠ SIN EMAIL VÁLIDO");
    return;
  }

  const datosCorreo = construirDatosFormulario(datosFila, nombreHoja);
  datosCorreo.estado        = datosFila[IDX_ESTADO];
  datosCorreo.observaciones = datosFila[IDX_OBSERVACIONES];

  const html = construirCorreoClienteHTML(datosCorreo);
  const logoBlob = DriveApp.getFileById(ID_LOGO_CORREOS).getBlob().setName("logoDavivienda");

  GmailApp.sendEmail(
    email,
    `Davivienda — Actualización trámite ${datosCorreo.impuesto} — Radicado ${datosCorreo.radicado}`,
    "Actualización de su trámite. Abra este correo en un cliente compatible con HTML.",
    {
      htmlBody: html,
      name: "Marcaciones y Reintegros - Davivienda",
      noReply: true,
      inlineImages: { logoDavivienda: logoBlob }
    }
  );

  marcarEnviado(sheet, fila, COL_NOTIFICAR);
}

//enviarCorreoArea 
function enviarCorreoArea(sheet, fila) {
  const datosFila = sheet.getRange(fila, 1, 1, COL_NOTIF_AREA).getValues()[0];
  const nombreHoja = sheet.getName();

  const correoAreaRaw = (datosFila[IDX_CORREO_AREA] || "").toString().trim();
  if (!correoAreaRaw) {
    sheet.getRange(fila, COL_NOTIF_AREA).setValue("⚠ SIN CORREO ÁREA");
    return;
  }

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
  const logoBlob = DriveApp.getFileById(ID_LOGO_CORREOS).getBlob().setName("logoDavivienda");

  GmailApp.sendEmail(
    correos.join(","),
    `[Remisión Interna] Trámite ${datosCorreo.impuesto} — Radicado ${datosCorreo.radicado}`,
    "Trámite remitido para gestión. Abra este correo en un cliente compatible con HTML.",
    {
      htmlBody: html,
      name: "Marcaciones y Reintegros - Davivienda",
      noReply: true,
      inlineImages: { logoDavivienda: logoBlob }
    }
  );

  marcarEnviado(sheet, fila, COL_NOTIF_AREA);
}

//construirCorreoClienteHTML
function construirCorreoClienteHTML(d) {
  const fechaEnvioStr = Utilities.formatDate(new Date(), "America/Bogota", "dd/MM/yyyy HH:mm");

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

      <tr><td style="background:#E1251B;background:linear-gradient(135deg, #E1251B 0%, #B81E15 100%);padding:18px 28px;">
        <table width="100%" cellpadding="0" cellspacing="0"><tr>
          <td align="left" valign="middle">
            <table cellpadding="0" cellspacing="0"><tr>
              <td valign="middle" style="padding-right:10px;"><img src="cid:logoDavivienda" alt="Davivienda" width="26" height="26" style="display:block;"></td>
              <td valign="middle"><span style="color:#fff;font-size:15px;font-weight:700;letter-spacing:0.5px;">DAVIVIENDA</span></td>
            </tr></table>
          </td>
          <td align="right" valign="middle"><span style="color:#ffffffcc;font-size:12px;">📅 ${fechaEnvioStr}</span></td>
        </tr></table>
      </td></tr>
      <tr><td style="background:#E1251B;background:linear-gradient(135deg, #E1251B 0%, #B81E15 100%);padding:0 28px 20px;color:#fff;">
        <div style="font-size:20px;font-weight:bold;">Actualización de su solicitud</div>
      </td></tr>

      <tr><td style="padding:24px 32px 8px;">
        <div style="text-align:center;padding:20px;background:#FCEBEA;border-radius:8px;border:2px dashed #E1251B;">
          <div style="font-size:13px;color:#B81E15;letter-spacing:1px;">RADICADO</div>
          <div style="font-size:32px;font-weight:bold;color:#B81E15;margin-top:8px;letter-spacing:2px;">#${d.radicado}</div>
          <div style="margin-top:14px;">
            <span style="background:${colorEstado};color:white;padding:6px 16px;border-radius:4px;font-weight:600;font-size:13px;letter-spacing:0.5px;">${d.estado || ""}</span>
          </div>
        </div>
      </td></tr>

      <tr><td style="padding:16px 32px;">
        <p style="margin:0;font-size:14px;color:#1f2937;line-height:1.6;">
        Estimado(a) <strong>${d.razonsocial || "solicitante"}</strong>,<br><br>
        Le informamos que su trámite de <strong>${d.impuesto}</strong> ha sido actualizado.
        A continuación encontrará el detalle.
        </p>
      </td></tr>

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

      <tr><td style="padding:0 32px 16px;">
        <div style="font-size:13px;font-weight:bold;color:#E1251B;text-transform:uppercase;margin-bottom:8px;letter-spacing:0.5px;">Observaciones del trámite</div>
        <div style="padding:14px 16px;background:#f8f9fa;border-left:4px solid #E1251B;border-radius:4px;font-size:13.5px;color:#1f2937;line-height:1.6;">
          ${obsHtml}
        </div>
      </td></tr>

      <tr><td style="padding:8px 32px 24px;">
        <div style="padding:14px;background:#eff6ff;border-left:3px solid #2563eb;border-radius:4px;font-size:12.5px;color:#1e3a8a;line-height:1.6;">
          Si tiene alguna duda sobre su trámite, comuníquese con su oficina indicando el número de radicado <strong>#${d.radicado || ""}</strong>.
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
//construirCorreoAreaHTML
function construirCorreoAreaHTML(d) {
  const fechaEnvioStr = Utilities.formatDate(new Date(), "America/Bogota", "dd/MM/yyyy HH:mm");

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

      <tr><td style="background:#E1251B;background:linear-gradient(135deg, #E1251B 0%, #B81E15 100%);padding:18px 28px;">
        <table width="100%" cellpadding="0" cellspacing="0"><tr>
          <td align="left" valign="middle">
            <table cellpadding="0" cellspacing="0"><tr>
              <td valign="middle" style="padding-right:10px;"><img src="cid:logoDavivienda" alt="Davivienda" width="26" height="26" style="display:block;"></td>
              <td valign="middle"><span style="color:#fff;font-size:15px;font-weight:700;letter-spacing:0.5px;">DAVIVIENDA</span></td>
            </tr></table>
          </td>
          <td align="right" valign="middle"><span style="color:#ffffffcc;font-size:12px;">📅 ${fechaEnvioStr}</span></td>
        </tr></table>
      </td></tr>
      <tr><td style="background:#E1251B;background:linear-gradient(135deg, #E1251B 0%, #B81E15 100%);padding:0 28px 20px;color:#fff;">
        <div style="font-size:11px;letter-spacing:2px;opacity:0.85;">REMISIÓN INTERNA</div>
        <div style="font-size:19px;font-weight:bold;margin-top:4px;">Trámite remitido para gestión</div>
      </td></tr>

      <tr><td style="padding:20px 32px 10px;">
        <div style="padding:14px 18px;background:#fff7ed;border-left:4px solid #ea580c;border-radius:4px;font-size:13px;color:#7c2d12;line-height:1.6;">
          <strong>Acción requerida:</strong> Se ha remitido el siguiente trámite a su área para gestión. Por favor revise el detalle y las instrucciones al pie del correo.
        </div>
      </td></tr>

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

      <tr><td style="padding:6px 32px 14px;">
        <div style="font-size:12px;font-weight:bold;color:#0f3a6b;text-transform:uppercase;margin-bottom:8px;letter-spacing:1px;">Instrucciones para el área remitida</div>
        <div style="padding:14px 18px;background:#eff6ff;border-left:4px solid #2563eb;border-radius:4px;font-size:13.5px;color:#0f172a;line-height:1.6;">
          ${obsAreaHtml}
        </div>
      </td></tr>

      <tr><td style="padding:6px 32px 18px;">
        <div style="font-size:12px;font-weight:bold;color:#0f3a6b;text-transform:uppercase;margin-bottom:8px;letter-spacing:1px;">Observaciones generales del trámite</div>
        <div style="padding:14px 18px;background:#f8fafc;border:1px solid #e2e8f0;border-radius:4px;font-size:13px;color:#334155;line-height:1.6;">
          ${obsHtml}
        </div>
      </td></tr>

      <tr><td style="padding:16px 32px;background:#f8fafc;border-top:1px solid #e2e8f0;font-size:11px;color:#64748b;text-align:center;">
        Comunicación interna generada automáticamente desde el sistema de gestión de marcaciones y reintegros.<br>
        <strong style="color:#0f3a6b;">Banco Davivienda S.A.</strong>
      </td></tr>

    </table>
  </td></tr>
</table>
</body></html>`;
}