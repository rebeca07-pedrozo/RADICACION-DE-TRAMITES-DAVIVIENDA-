/**
 * FUNCIÓN DE DIAGNÓSTICO TEMPORAL
 * Muestra exactamente qué ve el script para cada fila de IMPORT
 */
function diagnosticarRadicado() {
  const RADICADO_A_BUSCAR = "43"; // ← Cambia este número por el radicado que estás probando
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaImport = ss.getSheetByName("IMPORT");
  
  if (!hojaImport) {
    Logger.log("❌ No existe la hoja IMPORT");
    return;
  }
  
  const datos = hojaImport.getDataRange().getValues();
  Logger.log("📊 Total filas en IMPORT (incluido header): " + datos.length);
  
  const encabezadosOriginales = datos[0];
  Logger.log("📋 Encabezados de IMPORT:");
  encabezadosOriginales.forEach((h, i) => {
    Logger.log("  [" + i + "] = '" + h + "'");
  });
  
  const encabezadosNormalizados = encabezadosOriginales.map(h => normalizarTexto(h));
  const idxImpuestos = buscarIndice(encabezadosNormalizados, ["impuestos", "tipoimpuesto", "tipodeimpuesto"]);
  const idxMotivo    = buscarIndice(encabezadosNormalizados, ["motivo", "motivodelasolicitud", "tipodesolicitud"]);
  
  Logger.log("🔍 idxImpuestos = " + idxImpuestos + " | idxMotivo = " + idxMotivo);
  
  // Buscar la fila del radicado
  let filaEncontrada = null;
  for (let i = 1; i < datos.length; i++) {
    if ((datos[i][IDX_RADICADO] || "").toString().trim() === RADICADO_A_BUSCAR) {
      filaEncontrada = datos[i];
      Logger.log("✅ Fila encontrada en posición " + (i + 1));
      break;
    }
  }
  
  if (!filaEncontrada) {
    Logger.log("❌ No se encontró el radicado " + RADICADO_A_BUSCAR + " en IMPORT");
    return;
  }
  
  Logger.log("📄 Contenido completo de la fila:");
  filaEncontrada.forEach((celda, i) => {
    Logger.log("  [" + i + "] = '" + celda + "'");
  });
  
  const motivo = (filaEncontrada[idxMotivo] || "").toString().trim();
  const impuestosTexto = (filaEncontrada[idxImpuestos] || "").toString();
  
  Logger.log("🎯 Motivo leído: '" + motivo + "'");
  Logger.log("🎯 Impuestos leído: '" + impuestosTexto + "'");
  
  const MOTIVOS_VALIDOS = [
    "Marcación", "Reintegro", "Ambas",
    "Desmarcación", "Certif. Régimen Simple", "Certificación Régimen Simple", "Desistimiento"
  ];
  
  Logger.log("✓ ¿Motivo está en lista válida? " + MOTIVOS_VALIDOS.includes(motivo));
  
  // Verificar coincidencia con cada hoja
  const CONFIG = [
    { hoja: "ICA",   textos: ["Retención de ICA"] },
    { hoja: "Renta", textos: ["Retención de Renta", "JELPIT", "Propiedad horizontal", "Régimen simple"] },
    { hoja: "IVA",   textos: ["Retención de IVA"] },
    { hoja: "ImpIVA",textos: ["Impuesto de IVA"] }
  ];
  
  CONFIG.forEach(cfg => {
    const coincide = cfg.textos.some(t => impuestosTexto.includes(t));
    Logger.log("   → " + cfg.hoja + ": " + (coincide ? "✅ DEBE ENTRAR" : "❌ no entra"));
    if (coincide) {
      cfg.textos.forEach(t => {
        if (impuestosTexto.includes(t)) {
          Logger.log("      ↳ matchea con: '" + t + "'");
        }
      });
    }
  });
}