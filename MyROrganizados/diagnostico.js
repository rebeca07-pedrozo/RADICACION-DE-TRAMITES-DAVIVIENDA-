/**
 * ⚠️ SOLO USAR ANTES DEL LANZAMIENTO A PRODUCCIÓN.
 * Limpia las 4 hojas operativas y el histórico, dejando solo encabezados.
 */
function limpiarTodoParaProduccion() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // Confirmación de seguridad
  const respuesta = ui.alert(
    "⚠️ CONFIRMACIÓN",
    "Esto va a BORRAR todos los datos de las 4 hojas operativas y del Histórico. ¿Estás segura?",
    ui.ButtonSet.YES_NO
  );

  if (respuesta !== ui.Button.YES) {
    ui.alert("Operación cancelada. No se borró nada.");
    return;
  }

  let totalLimpiadas = 0;

  // Limpiar las 4 hojas operativas
  HOJAS_DESTINO.forEach(nombre => {
    const h = ss.getSheetByName(nombre);
    if (!h) return;
    const ultimaFila = h.getLastRow();
    if (ultimaFila > 1) {
      h.deleteRows(2, ultimaFila - 1);
      totalLimpiadas++;
    }
  });

  // Limpiar histórico si existe
  const hojaHist = ss.getSheetByName(HOJA_HISTORICO);
  if (hojaHist && hojaHist.getLastRow() > 1) {
    hojaHist.deleteRows(2, hojaHist.getLastRow() - 1);
    totalLimpiadas++;
  }

  ui.alert(
    "✅ Limpieza completada",
    `Se limpiaron ${totalLimpiadas} hoja(s).\n\nEl sistema está listo para producción.`,
    ui.ButtonSet.OK
  );
}