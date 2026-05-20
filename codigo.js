/**

 * ============================================================================

 *  COMBINACIÓN DE CORRESPONDENCIA - VERSIÓN REPARADA

 * ============================================================================

 *  Triggers necesarios:

 *  - onOpen        : Automático (función reservada de Apps Script)

 *  - completarFormato : Trigger instalable, "Al enviar formulario"

 *

 *  Cambios respecto a la versión original:

 *  - Se reemplaza CigoApp.enviarCorreo por MailApp.sendEmail (sin librerías).

 *  - obtenerParametros() ahora usa un objeto/Map y no toString().split(",").

 *  - obtenerNumeroColumna devuelve número (antes devolvía string).

 *  - completarFormato: se elimina el check de columnStart==columnEnd que

 *    abortaba el proceso indebidamente.

 *  - instalar(): solo verifica si YA existe el trigger onFormSubmit.

 *  - Nuevas funciones: doGet, getCamposFormulario, procesarFormularioWeb

 *    para servir el formulario HTML (modal y web app).

 * ============================================================================

 */



var parametros;



// ====================================================================

//  MENÚ

// ====================================================================



/**

 * Crea el menú en la hoja de cálculo al abrir el archivo.

 */

function onOpen() {

  SpreadsheetApp.getUi()

    .createMenu("Marcaciones y Reintegros")

    .addItem(" Radicar nueva solicitud", "abrirFormulario")

    .addSeparator()

    .addItem(" Generar PDFs pendientes", "generarPDFs")

    .addItem(" Instalar activadores", "instalar")

    .addToUi();

}



/**

 * Abre el formulario HTML como diálogo modal dentro de la hoja.

 */

function abrirFormulario() {

  var html = HtmlService.createTemplateFromFile("Formulario")

    .evaluate()

    .setWidth(760)

    .setHeight(680)

    .setTitle("Radicar solicitud");

  SpreadsheetApp.getUi().showModalDialog(html, "Radicar solicitud");

}



/**

 * Sirve el formulario como Web App pública (URL independiente).

 * Para activarlo: Implementar > Nueva implementación > Tipo: Aplicación web.

 */

function doGet() {

  return HtmlService.createTemplateFromFile("Formulario")

    .evaluate()

    .setTitle("Radicar trámite")

    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

}



/**

 * Permite incluir otros archivos HTML (CSS, JS) dentro del HTML principal.

 */

function include(filename) {

  return HtmlService.createHtmlOutputFromFile(filename).getContent();

}



// ====================================================================

//  INSTALACIÓN DE TRIGGERS

// ====================================================================



/**

 * Instala el trigger de envío de formulario si aún no existe.

 */

function instalar() {

  var triggers = ScriptApp.getProjectTriggers();

  var yaInstalado = triggers.some(function (t) {

    return t.getHandlerFunction() === "completarFormato" &&

           t.getEventType() === ScriptApp.EventType.ON_FORM_SUBMIT;

  });



  if (yaInstalado) {

    Browser.msgBox(

      "Advertencia",

      "Ya tienes instalado el activador necesario.",

      Browser.Buttons.OK

    );

    return;

  }



  ScriptApp.newTrigger("completarFormato")

    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())

    .onFormSubmit()

    .create();



  Browser.msgBox(

    "Instalación Exitosa",

    "Los activadores fueron instalados exitosamente.",

    Browser.Buttons.OK

  );

}



// ====================================================================

//  FLUJO DEL FORMULARIO DE GOOGLE FORMS (original, reparado)

// ====================================================================



/**

 * Se ejecuta al recibir una respuesta del Google Form.

 * Copia los datos a la hoja Consolidado, genera el PDF y notifica por correo.

 * @param {Event} e Evento del envío del formulario

 */

function completarFormato(e) {

  parametros = obtenerParametros();



  var sps = SpreadsheetApp.getActiveSpreadsheet();

  var hojaRespuestas = sps.getActiveSheet();

  var fila = hojaRespuestas.getActiveRange().getRowIndex();

  var respuestasFormulario = hojaRespuestas

    .getRange(fila, 1, 1, hojaRespuestas.getLastColumn())

    .getValues()[0];



  var hojaConsolidado = sps.getSheetByName(obtenerParametro("nombreHojaConsolidado"));

  hojaConsolidado.appendRow(respuestasFormulario);



  // Copia fórmulas a la nueva fila si la hoja Parametros lo indica

  var columnasFormulasRaw = obtenerParametro("columnasFormulas");

  var columnasFormulas = columnasFormulasRaw ? columnasFormulasRaw.toString().split(",") : [];

  var ultimaFila = hojaConsolidado.getLastRow();

  var ultimaColumna = hojaConsolidado.getLastColumn();



  for (var i = 0; i < columnasFormulas.length; i++) {

    var col = columnasFormulas[i].trim();

    if (!col) continue;

    var rangoFormula = hojaConsolidado.getRange(col + "2");

    rangoFormula.copyTo(hojaConsolidado.getRange(col + ultimaFila));

  }



  var respuestasConsolidado = hojaConsolidado

    .getRange(ultimaFila, 1, 1, ultimaColumna)

    .getValues()[0];



  var archivoPDF = combinarCampos(respuestasConsolidado);



  // Primero escribe la URL para no perderla si falla el envío de correo

  hojaConsolidado.getRange(ultimaFila, ultimaColumna).setValue(archivoPDF.getUrl());



  try {

    enviarNotificacion(respuestasConsolidado, archivoPDF);

  } catch (err) {

    Logger.log("Error enviando correo: " + err);

  }

}



// ====================================================================

//  GENERACIÓN MASIVA

// ====================================================================



/**

 * Genera PDFs para todas las filas del consolidado que aún no lo tengan.

 */

function generarPDFs() {

  parametros = obtenerParametros();

  var hojaConsolidado = SpreadsheetApp.getActiveSpreadsheet()

    .getSheetByName(obtenerParametro("nombreHojaConsolidado"));



  var datos = hojaConsolidado.getDataRange().getValues();

  var columnaPDF = hojaConsolidado.getLastColumn();



  var generados = 0;

  for (var fila = 1; fila < datos.length; fila++) {

    if (datos[fila][0] === "") break;

    if (datos[fila][columnaPDF - 1] !== "") continue;



    var archivoPDF = combinarCampos(datos[fila]);

    hojaConsolidado.getRange(fila + 1, columnaPDF).setValue(archivoPDF.getUrl());



    try {

      enviarNotificacion(datos[fila], archivoPDF);

    } catch (err) {

      Logger.log("Error enviando correo fila " + (fila + 1) + ": " + err);

    }

    generados++;

  }



  Browser.msgBox(

    "PDFs Generados",

    "Se generaron " + generados + " PDF(s) exitosamente.",

    Browser.Buttons.OK

  );

}



// ====================================================================

//  COMBINACIÓN DE CAMPOS

// ====================================================================



/**

 * Crea el PDF a partir de la plantilla, reemplazando los tags por los datos.

 * @param {Array} registro Fila de datos del consolidado

 * @return {File} Archivo PDF generado

 */

function combinarCampos(registro) {

  var idPlantilla;



  if (obtenerParametro("multiplesPlantillas")) {

    var opcion = registro[obtenerNumeroColumna(obtenerParametro("columnaOpcionPlantilla")) - 1];

    var datosPlantillas = SpreadsheetApp.getActiveSpreadsheet()

      .getSheetByName(obtenerParametro("nombreHojaPlantillas"))

      .getDataRange()

      .getValues();

    idPlantilla = buscarPorLlave(opcion, datosPlantillas, 1, 2);

    if (!idPlantilla) idPlantilla = obtenerParametro("idPlantilla");

  } else {

    idPlantilla = obtenerParametro("idPlantilla");

  }



  // Nombre del archivo basado en el radicado (columna 'radicado' en el consolidado)

  var idNombreArchivo = registro[obtenerNumeroColumna(obtenerParametro("columnaIdArchivo")) - 1];

  var nombreArchivo = obtenerParametro("nombreArchivo").replace("##", idNombreArchivo);



  // La copia de trabajo se crea en la carpeta general (temporal)

  var carpetaTrabajo = DriveApp.getFolderById(obtenerParametro("idCarpetaRepositorio"));

  // El PDF final se guarda en la carpeta dedicada de PDFs

  var carpetaPDFs = DriveApp.getFolderById(obtenerParametro("idCarpetaPDFs"));



  var archivoPlantilla = DriveApp.getFileById(idPlantilla)

    .makeCopy(nombreArchivo, carpetaTrabajo);



  var datosCampos = SpreadsheetApp.getActiveSpreadsheet()

    .getSheetByName(obtenerParametro("nombreHojaCampos"))

    .getDataRange()

    .getValues();



  var documentoPlantilla = DocumentApp.openById(archivoPlantilla.getId());

  var body = documentoPlantilla.getBody();



  for (var i = 1; i < datosCampos.length; i++) {

    var numeroColumna = obtenerNumeroColumna(datosCampos[i][0]);

    var valorColumna = registro[numeroColumna - 1];



    if (valorColumna instanceof Date) {

      valorColumna = Utilities.formatDate(valorColumna, "America/Bogota", "dd/MM/yyyy HH:mm");

    }

    if (valorColumna === null || valorColumna === undefined) {

      valorColumna = "";

    }



    var tag = datosCampos[i][1];

    body.replaceText(tag, valorColumna.toString());

  }



  documentoPlantilla.saveAndClose();



  var blobPDF = DriveApp.getFileById(archivoPlantilla.getId()).getAs(MimeType.PDF);

  var archivoPDF = carpetaPDFs.createFile(blobPDF);

  archivoPDF.setName(nombreArchivo);

  archivoPlantilla.setTrashed(true);



  return archivoPDF;

}



// ====================================================================

//  GENERADOR DE NÚMERO DE RADICADO CONSECUTIVO

// ====================================================================



/**

 * Genera el siguiente número de radicado consecutivo (formato 00001, 00002, ...).

 * Usa PropertiesService para persistir el contador entre ejecuciones.

 * @return {String} Número de radicado con padding de ceros (5 dígitos)

 */

function generarRadicado() {

  var props = PropertiesService.getScriptProperties();

  var actual = parseInt(props.getProperty("ultimoRadicado") || "0", 10);

  var siguiente = actual + 1;

  props.setProperty("ultimoRadicado", siguiente.toString());



  // Padding con ceros a la izquierda hasta 5 dígitos

  var str = siguiente.toString();

  while (str.length < 5) str = "0" + str;

  return str;

}



/**

 * Función opcional para reiniciar o cambiar el contador.

 * Se ejecuta manualmente desde el editor si se necesita.

 */

function reiniciarRadicado() {

  var props = PropertiesService.getScriptProperties();

  props.setProperty("ultimoRadicado", "0");

  Browser.msgBox("Contador reiniciado",

    "El próximo radicado será 00001.",

    Browser.Buttons.OK);

}



// ====================================================================

//  NOTIFICACIÓN POR CORREO

// ====================================================================



/**

 * Envía un correo con el PDF adjunto.

 * Reemplaza CigoApp por MailApp (nativo, sin dependencias externas).

 */

function enviarNotificacion(respuestas, archivoPDF) {

  var columnaCorreo = obtenerParametro("columnaEnvioCorreo");

  if (!columnaCorreo || columnaCorreo === 0) return;



  var correo = respuestas[obtenerNumeroColumna(columnaCorreo) - 1];

  if (!correo) return;



  // Construir un objeto con todos los datos por nombre de columna

  var sps = SpreadsheetApp.getActiveSpreadsheet();

  var hojaConsolidado = sps.getSheetByName(obtenerParametro("nombreHojaConsolidado"));

  var encabezados = hojaConsolidado

    .getRange(1, 1, 1, hojaConsolidado.getLastColumn())

    .getValues()[0];



  var datos = {};

  for (var i = 0; i < encabezados.length; i++) {

    var clave = normalizarClave(encabezados[i]);

    var valor = respuestas[i];

    if (valor instanceof Date) {

      valor = Utilities.formatDate(valor, "America/Bogota", "dd/MM/yyyy HH:mm");

    }

    datos[clave] = valor || "";

  }



  var radicado = datos.radicado || "(sin asignar)";

  var asunto = "Solicitud radicada #" + radicado + " - " + (datos.razonsocial || "Davivienda");



  // Construir el cuerpo del correo en HTML

  var htmlBody = construirCorreoHTML(datos);

  var textoPlano = construirCorreoTextoPlano(datos);



  MailApp.sendEmail({

    to: correo,

    subject: asunto,

    body: textoPlano,

    htmlBody: htmlBody,

    noReply: true,

    attachments: [archivoPDF.getAs(MimeType.PDF)]

  });

}



/**

 * Construye el cuerpo HTML del correo de confirmación.

 */

function construirCorreoHTML(datos) {

  var fila = function(etiqueta, valor) {

    if (!valor) return "";

    return '<tr>' +

      '<td style="padding:8px 12px;background:#f9fafb;color:#6b7280;font-size:13px;width:40%;border-bottom:1px solid #e5e7eb;"><strong>' + etiqueta + '</strong></td>' +

      '<td style="padding:8px 12px;color:#1f2937;font-size:13px;border-bottom:1px solid #e5e7eb;">' + valor + '</td>' +

    '</tr>';

  };



  var html =

'<!DOCTYPE html><html><body style="margin:0;padding:0;background:#f5f6fa;font-family:Arial,sans-serif;">' +

  '<table width="100%" cellpadding="0" cellspacing="0" style="background:#f5f6fa;padding:20px 0;">' +

    '<tr><td align="center">' +

      '<table width="600" cellpadding="0" cellspacing="0" style="background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,0.08);">' +



        // Cabecera roja

        '<tr><td style="background:#E1251B;padding:24px 32px;color:#fff;">' +

          '<div style="font-size:12px;letter-spacing:1px;opacity:0.9;">BANCO DAVIVIENDA</div>' +

          '<div style="font-size:20px;font-weight:bold;margin-top:4px;">Solicitud radicada exitosamente</div>' +

        '</td></tr>' +



        // Número de radicado destacado

        '<tr><td style="padding:24px 32px 8px;">' +

          '<div style="text-align:center;padding:20px;background:#FCEBEA;border-radius:8px;border:2px dashed #E1251B;">' +

            '<div style="font-size:13px;color:#B81E15;letter-spacing:1px;">SU NÚMERO DE RADICADO</div>' +

            '<div style="font-size:36px;font-weight:bold;color:#B81E15;margin-top:8px;letter-spacing:2px;">#' + (datos.radicado || "--") + '</div>' +

            '<div style="font-size:12px;color:#6b7280;margin-top:8px;">Guarde este número para consultar el estado de su solicitud.</div>' +

          '</div>' +

        '</td></tr>' +



        // Saludo

        '<tr><td style="padding:16px 32px;">' +

          '<p style="margin:0;font-size:14px;color:#1f2937;line-height:1.6;">' +

          'Estimado(a) <strong>' + (datos.razonsocial || "solicitante") + '</strong>,<br><br>' +

          'Hemos recibido su solicitud de <strong>' + (datos.motivo || "trámite") + '</strong>. ' +

          'A continuación encontrará el resumen y, adjunto a este correo, el PDF con su radicación.' +

          '</p>' +

        '</td></tr>' +



        // Tabla resumen

        '<tr><td style="padding:0 32px 16px;">' +

          '<div style="font-size:13px;font-weight:bold;color:#E1251B;text-transform:uppercase;margin-bottom:8px;letter-spacing:0.5px;">Resumen de la solicitud</div>' +

          '<table width="100%" cellpadding="0" cellspacing="0" style="border:1px solid #e5e7eb;border-radius:6px;overflow:hidden;">' +

            fila("Fecha", datos.timestamp) +

            fila("Oficina solicitante", datos.nombreoficina) +

            fila("Código de oficina", datos.codigooficina) +

            fila("Razón social", datos.razonsocial) +

            fila("Cédula/NIT", datos.cedula + (datos.dv ? " - DV " + datos.dv : "")) +

            fila("Motivo", datos.motivo) +

            fila("Tipo de producto", datos.tipoproducto) +

            fila("N° de producto", datos.numeroproducto) +

            fila("Impuestos", datos.impuestos) +

            fila("Año a reintegrar", datos.periodo) +

            fila("Valor a reintegrar", datos.valor) +

            fila("Ciudad", datos.ciudad) +

          '</table>' +

        '</td></tr>' +



        // Descripción del motivo

        (datos.descripcionmotivo ?

          '<tr><td style="padding:0 32px 16px;">' +

            '<div style="font-size:13px;font-weight:bold;color:#E1251B;text-transform:uppercase;margin-bottom:8px;letter-spacing:0.5px;">Descripción del motivo</div>' +

            '<div style="padding:12px 14px;background:#f9fafb;border-radius:6px;font-size:13px;color:#1f2937;line-height:1.5;border:1px solid #e5e7eb;">' +

              datos.descripcionmotivo +

            '</div>' +

          '</td></tr>' : ''

        ) +



        // Nota

        '<tr><td style="padding:16px 32px 24px;">' +

          '<div style="padding:14px;background:#eff6ff;border-left:3px solid #2563eb;border-radius:4px;font-size:12.5px;color:#1e3a8a;line-height:1.6;">' +

            '<strong> Documento adjunto:</strong> Encuentre el PDF de su radicación adjunto a este correo. ' +

            'Para cualquier consulta sobre el estado de la solicitud, comuníquese con su oficina indicando el número de radicado <strong>#' + (datos.radicado || "") + '</strong>.' +

          '</div>' +

        '</td></tr>' +



        // Pie

        '<tr><td style="padding:16px 32px;background:#f9fafb;border-top:1px solid #e5e7eb;font-size:11px;color:#6b7280;text-align:center;">' +

          'Este correo se generó automáticamente. Por favor no responda a este mensaje.<br>' +

          '<strong style="color:#E1251B;">Banco Davivienda S.A.</strong>' +

        '</td></tr>' +



      '</table>' +

    '</td></tr>' +

  '</table>' +

'</body></html>';



  return html;

}



/**

 * Versión en texto plano del correo (para clientes que no muestran HTML).

 */

function construirCorreoTextoPlano(datos) {

  return [

    "BANCO DAVIVIENDA",

    "Solicitud radicada exitosamente",

    "",

    "═══════════════════════════════════════",

    "NÚMERO DE RADICADO: #" + (datos.radicado || "--"),

    "═══════════════════════════════════════",

    "",

    "Estimado(a) " + (datos.razonsocial || "solicitante") + ",",

    "",

    "Hemos recibido su solicitud de " + (datos.motivo || "trámite") + ".",

    "Encontrará adjunto el PDF con su radicación.",

    "",

    "RESUMEN DE LA SOLICITUD",

    "───────────────────────────────────────",

    "Fecha:               " + (datos.timestamp || ""),

    "Oficina:             " + (datos.nombreoficina || ""),

    "Código oficina:      " + (datos.codigooficina || ""),

    "Razón social:        " + (datos.razonsocial || ""),

    "Cédula/NIT:          " + (datos.cedula || "") + (datos.dv ? " - DV " + datos.dv : ""),

    "Motivo:              " + (datos.motivo || ""),

    "Tipo producto:       " + (datos.tipoproducto || ""),

    "N° producto:         " + (datos.numeroproducto || ""),

    datos.impuestos   ? "Impuestos:           " + datos.impuestos : "",

    datos.periodo     ? "Año a reintegrar:    " + datos.periodo : "",

    datos.valor       ? "Valor:               " + datos.valor : "",

    datos.ciudad      ? "Ciudad:              " + datos.ciudad : "",

    "",

    datos.descripcionmotivo ? "Descripción:\n" + datos.descripcionmotivo : "",

    "",

    "───────────────────────────────────────",

    "Para consultar el estado, comuníquese con su oficina",

    "indicando el número de radicado #" + (datos.radicado || ""),

    "",

    "Este correo se generó automáticamente.",

    "Banco Davivienda S.A."

  ].filter(function(l) { return l !== ""; }).join("\n");

}



// ====================================================================

//  FLUJO DEL FORMULARIO WEB (modal + web app)

// ====================================================================



/**

 * Recibe los datos del formulario HTML, los inserta en el consolidado,

 * sube los archivos adjuntos a Drive y dispara la generación del PDF.

 *

 * @param {Object} datos Objeto con los campos del formulario

 * @param {Array} archivos Lista de objetos {nombre, tipo, base64} adjuntados

 * @return {Object} {ok, urlPDF, mensaje}

 */

function procesarFormularioWeb(datos, archivos) {

  try {

    parametros = obtenerParametros();

    var sps = SpreadsheetApp.getActiveSpreadsheet();

    var hojaConsolidado = sps.getSheetByName(obtenerParametro("nombreHojaConsolidado"));



    // 1. Generar el número de radicado consecutivo (00001, 00002, ...)

    var radicado = generarRadicado();

    datos.radicado = radicado;



    // 2. Subir adjuntos a una subcarpeta nombrada con el radicado

    var urlsAdjuntos = subirAdjuntos(archivos, radicado);



    // 3. Construir la fila respetando el orden de columnas del consolidado

    var encabezados = hojaConsolidado

      .getRange(1, 1, 1, hojaConsolidado.getLastColumn())

      .getValues()[0];



    var fila = encabezados.map(function (h) {

      var clave = normalizarClave(h);

      if (clave === "soportes" || clave === "adjuntos") return urlsAdjuntos;

      if (clave === "timestamp" || clave === "marcatemporal") return new Date();

      if (clave === "radicado") return radicado;

      return datos[clave] || "";

    });



    hojaConsolidado.appendRow(fila);



    var ultimaFila = hojaConsolidado.getLastRow();

    var ultimaColumna = hojaConsolidado.getLastColumn();

    var respuestasConsolidado = hojaConsolidado

      .getRange(ultimaFila, 1, 1, ultimaColumna)

      .getValues()[0];



    // 4. Generar el PDF

    var archivoPDF = combinarCampos(respuestasConsolidado);

    hojaConsolidado.getRange(ultimaFila, ultimaColumna).setValue(archivoPDF.getUrl());



    // 5. Enviar correo de confirmación al solicitante

    try {

      enviarNotificacion(respuestasConsolidado, archivoPDF);

    } catch (err) {

      Logger.log("Error correo: " + err);

    }



    return {

      ok: true,

      radicado: radicado,

      urlPDF: archivoPDF.getUrl(),

      mensaje: "Solicitud radicada correctamente con el número " + radicado + "."

    };

  } catch (err) {

    Logger.log("Error procesarFormularioWeb: " + err);

    return { ok: false, mensaje: err.toString() };

  }

}



/**

 * Sube los archivos adjuntos del formulario a una subcarpeta en Drive.

 * @return {String} URLs separadas por coma

 */

function subirAdjuntos(archivos, identificador) {

  if (!archivos || archivos.length === 0) return "";



  var carpetaRaiz = DriveApp.getFolderById(obtenerParametro("idCarpetaRepositorio"));

  var nombreSubcarpeta = "Soportes_Radicacion_" + identificador;

  var subcarpeta = carpetaRaiz.createFolder(nombreSubcarpeta);



  var urls = [];

  for (var i = 0; i < archivos.length; i++) {

    var a = archivos[i];

    var blob = Utilities.newBlob(Utilities.base64Decode(a.base64), a.tipo, a.nombre);

    var file = subcarpeta.createFile(blob);

    urls.push(file.getUrl());

  }

  return urls.join(", ");

}



/**

 * Normaliza un encabezado para usarlo como clave (sin tildes, minúsculas, sin espacios).

 */

function normalizarClave(texto) {

  return texto

    .toString()

    .toLowerCase()

    .normalize("NFD").replace(/[\u0300-\u036f]/g, "") // quita tildes

    .replace(/[^a-z0-9]/g, ""); // solo letras y números

}



// ====================================================================

//  UTILIDADES

// ====================================================================



/**

 * Convierte una letra de columna ("A", "B", "AA") a su número (1, 2, 27).

 */

function obtenerNumeroColumna(columna) {

  if (typeof columna === "string") {

    return SpreadsheetApp.getActiveSheet().getRange(columna + "1").getColumn();

  }

  if (Array.isArray(columna)) {

    return columna.map(obtenerNumeroColumna);

  }

  return columna;

}



/**

 * Busca un valor en una columna de un rango y devuelve el valor de otra columna.

 */

function buscarPorLlave(llave, datos, colBus, colRes) {

  for (var i = 0; i < datos.length; i++) {

    if (datos[i][colBus - 1] === llave) {

      if (colRes === -1) return datos[i];

      return datos[i][colRes - 1];

    }

  }

  return null;

}



/**

 * Carga los parámetros desde la hoja "Parametros".

 * Reparado: ahora usa un objeto en lugar de toString().split(",").

 */

function obtenerParametros() {

  var hojaParametros = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Parametros");

  if (!hojaParametros) {

    throw new Error("No se encuentra la hoja 'Parametros'.");

  }



  var datos = hojaParametros.getDataRange().getValues();

  var mapa = {};

  for (var i = 0; i < datos.length; i++) {

    var clave = datos[i][0];

    if (clave) mapa[clave] = datos[i][1];

  }

  return mapa;

}



/**

 * Devuelve el valor de un parámetro.

 */

function obtenerParametro(parametro) {

  if (!(parametro in parametros)) {

    throw new Error("El parámetro '" + parametro + "' no se encuentra en la hoja de parámetros.");

  }

  return parametros[parametro];

}

