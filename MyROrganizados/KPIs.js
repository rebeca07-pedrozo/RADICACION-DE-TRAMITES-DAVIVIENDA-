/**
 * ============================================================================
 * KPIs — MARCACIONES Y REINTEGROS (Formulario + Organizados + Uso)
 * Registra un snapshot cada 10 minutos en la hoja "KPIs" para armar un
 * tablero histórico (Looker Studio puede conectarse directo a esta hoja).
 * ============================================================================
 */

const HOJA_KPIS = "KPIs";

const ENCABEZADOS_KPIS = [
  "Fecha y Hora",

  // ===== FORMULARIO (intake) =====
  "FORM_Total_Historico",
  "FORM_Radicados_Hoy",
  "FORM_Radicados_Semana",
  "FORM_Radicados_Mes",
  "FORM_Motivo_Marcacion",
  "FORM_Motivo_Reintegro",
  "FORM_Motivo_Ambas",
  "FORM_Motivo_Desmarcacion",
  "FORM_Motivo_CertifRegimenSimple",
  "FORM_Motivo_Desistimiento",

  // ===== ORGANIZADOS (operación) =====
  "ORG_Activos_ICA",
  "ORG_Activos_Renta",
  "ORG_Activos_IVA",
  "ORG_Activos_ImpuestoIVA",
  "ORG_Total_Activos",
  "ORG_Estado_RecibidoEnProceso",
  "ORG_Estado_Aprobado",
  "ORG_Estado_Rechazado",
  "ORG_Estado_Requerido",
  "ORG_PorcentajeAprobacion",
  "ORG_PendientesNotificarCliente",
  "ORG_PendientesNotificarArea",
  "ORG_Total_Historico",
  "ORG_Total_General",

  // ===== USO DE HERRAMIENTAS =====
  "USO_ConsultasPortal_Total",
  "USO_ConsultasPortal_Hoy"
];

/**
 * Función principal — corre cada 10 minutos con el trigger.
 */
function calcularYRegistrarKPIs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let hojaKpis = ss.getSheetByName(HOJA_KPIS);
  if (!hojaKpis) {
    hojaKpis = ss.insertSheet(HOJA_KPIS);
    hojaKpis.getRange(1, 1, 1, ENCABEZADOS_KPIS.length).setValues([ENCABEZADOS_KPIS]);
    hojaKpis.getRange(1, 1, 1, ENCABEZADOS_KPIS.length)
      .setFontWeight("bold").setBackground("#E1251B").setFontColor("white")
      .setHorizontalAlignment("center");
    hojaKpis.setFrozenRows(1);
  }

  const kpisFormulario = calcularKpisFormulario(ss);
  const kpisOrganizados = calcularKpisOrganizados(ss);
  const kpisUso = obtenerContadoresConsultas();

  const fila = [
    new Date(),
    kpisFormulario.totalHistorico,
    kpisFormulario.radicadosHoy,
    kpisFormulario.radicadosSemana,
    kpisFormulario.radicadosMes,
    kpisFormulario.porMotivo["marcacion"] || 0,
    kpisFormulario.porMotivo["reintegro"] || 0,
    kpisFormulario.porMotivo["ambas"] || 0,
    kpisFormulario.porMotivo["desmarcacion"] || 0,
    (kpisFormulario.porMotivo["certifregimensimple"] || 0) + (kpisFormulario.porMotivo["certificacionregimensimple"] || 0),
    kpisFormulario.porMotivo["desistimiento"] || 0,

    kpisOrganizados.activosPorHoja["ICA"] || 0,
    kpisOrganizados.activosPorHoja["Renta"] || 0,
    kpisOrganizados.activosPorHoja["IVA"] || 0,
    kpisOrganizados.activosPorHoja["ImpuestoIVA"] || 0,
    kpisOrganizados.totalActivos,
    kpisOrganizados.porEstado["RECIBIDO EN PROCESO"] || 0,
    kpisOrganizados.porEstado["APROBADO"] || 0,
    kpisOrganizados.porEstado["RECHAZADO"] || 0,
    kpisOrganizados.porEstado["REQUERIDO"] || 0,
    kpisOrganizados.porcentajeAprobacion,
    kpisOrganizados.pendientesNotificarCliente,
    kpisOrganizados.pendientesNotificarArea,
    kpisOrganizados.totalHistorico,
    kpisOrganizados.totalGeneral,

    kpisUso.total,
    kpisUso.hoy
  ];

  hojaKpis.appendRow(fila);
  Logger.log("KPIs registrados: " + new Date());
}

/**
 * Calcula KPIs del formulario (intake) a partir de la hoja IMPORT.
 */
function calcularKpisFormulario(ss) {
  const hojaImport = ss.getSheetByName("IMPORT");
  const resultado = {
    totalHistorico: 0,
    radicadosHoy: 0,
    radicadosSemana: 0,
    radicadosMes: 0,
    porMotivo: {}
  };

  if (!hojaImport) return resultado;

  const datos = hojaImport.getDataRange().getValues();
  if (datos.length <= 1) return resultado;

  const encabezados = datos[0].map(h => normalizarTexto(h));
  const idxMotivo = buscarIndice(encabezados, ["motivo", "motivodelasolicitud", "tipodesolicitud"]);

  const ahora = new Date();
  const inicioHoy = new Date(ahora.getFullYear(), ahora.getMonth(), ahora.getDate());
  const inicioSemana = new Date(inicioHoy);
  inicioSemana.setDate(inicioSemana.getDate() - inicioSemana.getDay());
  const inicioMes = new Date(ahora.getFullYear(), ahora.getMonth(), 1);

  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];
    const radicado = (fila[1] || "").toString().trim();
    if (!radicado) continue;

    resultado.totalHistorico++;

    const ts = fila[0];
    if (ts instanceof Date) {
      if (ts >= inicioHoy) resultado.radicadosHoy++;
      if (ts >= inicioSemana) resultado.radicadosSemana++;
      if (ts >= inicioMes) resultado.radicadosMes++;
    }

    if (idxMotivo !== -1) {
      const motivo = normalizarTexto(fila[idxMotivo]);
      if (motivo) {
        resultado.porMotivo[motivo] = (resultado.porMotivo[motivo] || 0) + 1;
      }
    }
  }

  return resultado;
}

/**
 * Calcula KPIs operativos a partir de las hojas de Marcaciones y el Histórico.
 */
function calcularKpisOrganizados(ss) {
  const resultado = {
    activosPorHoja: {},
    totalActivos: 0,
    porEstado: {
      "RECIBIDO EN PROCESO": 0,
      "APROBADO": 0,
      "RECHAZADO": 0,
      "REQUERIDO": 0
    },
    porcentajeAprobacion: 0,
    pendientesNotificarCliente: 0,
    pendientesNotificarArea: 0,
    totalHistorico: 0,
    totalGeneral: 0
  };

  const nombresCortos = {
    "Marcación y Reintegro de retencion de ICA": "ICA",
    "Marcación y Reintegro de retencion de renta": "Renta",
    "Marcación y Reintegro de retencion de IVA": "IVA",
    "Marcación y Reintegro de impuesto IVA": "ImpuestoIVA"
  };

  HOJAS_DESTINO.forEach(nombreHoja => {
    const hoja = ss.getSheetByName(nombreHoja);
    if (!hoja) return;

    const datos = hoja.getDataRange().getValues();
    if (datos.length <= 1) return;

    let contador = 0;
    for (let i = 1; i < datos.length; i++) {
      const fila = datos[i];
      const radicado = (fila[IDX_RADICADO] || "").toString().trim();
      if (!radicado) continue;

      contador++;

      const estado = (fila[IDX_ESTADO] || "").toString().trim().toUpperCase();
      if (resultado.porEstado.hasOwnProperty(estado)) {
        resultado.porEstado[estado]++;
      }

      const notif = (fila[IDX_NOTIFICAR] || "").toString().trim();
      if (notif === "NO ENVIADO") resultado.pendientesNotificarCliente++;

      const notifArea = (fila[IDX_NOTIF_AREA] || "").toString().trim();
      const correoArea = (fila[IDX_CORREO_AREA] || "").toString().trim();
      if (correoArea && notifArea === "NO ENVIADO") resultado.pendientesNotificarArea++;
    }

    const corto = nombresCortos[nombreHoja] || nombreHoja;
    resultado.activosPorHoja[corto] = contador;
    resultado.totalActivos += contador;
  });

  const hojaHist = ss.getSheetByName(HOJA_HISTORICO);
  if (hojaHist) {
    const datosHist = hojaHist.getDataRange().getValues();
    for (let i = 1; i < datosHist.length; i++) {
      const radicado = (datosHist[i][IDX_RADICADO] || "").toString().trim();
      if (radicado) resultado.totalHistorico++;
    }
  }

  const totalDecididos = resultado.porEstado["APROBADO"] + resultado.porEstado["RECHAZADO"];
  resultado.porcentajeAprobacion = totalDecididos > 0
    ? Math.round((resultado.porEstado["APROBADO"] / totalDecididos) * 1000) / 10
    : 0;

  resultado.totalGeneral = resultado.totalActivos + resultado.totalHistorico;

  return resultado;
}

// ============================================================================
// CONTADOR DE USO DEL PORTAL DE CONSULTA
// ============================================================================

/**
 * Incrementa el contador de consultas (total + hoy).
 * Se llama desde buscarRadicado() en el Código.gs principal.
 */
function incrementarContadorConsultas() {
  const props = PropertiesService.getScriptProperties();

  const totalActual = parseInt(props.getProperty("consultasTotal") || "0", 10);
  props.setProperty("consultasTotal", (totalActual + 1).toString());

  const hoyStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
  const fechaGuardada = props.getProperty("consultasHoyFecha");

  if (fechaGuardada !== hoyStr) {
    props.setProperty("consultasHoyFecha", hoyStr);
    props.setProperty("consultasHoyCount", "1");
  } else {
    const countActual = parseInt(props.getProperty("consultasHoyCount") || "0", 10);
    props.setProperty("consultasHoyCount", (countActual + 1).toString());
  }
}

/**
 * Devuelve los contadores actuales de consultas (para el snapshot de KPIs).
 */
function obtenerContadoresConsultas() {
  const props = PropertiesService.getScriptProperties();

  const total = parseInt(props.getProperty("consultasTotal") || "0", 10);

  const hoyStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
  const fechaGuardada = props.getProperty("consultasHoyFecha");
  const hoy = (fechaGuardada === hoyStr)
    ? parseInt(props.getProperty("consultasHoyCount") || "0", 10)
    : 0;

  return { total, hoy };
}