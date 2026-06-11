// ===========================================================================
//  ARCHIVADO AUTOMÁTICO DE CASOS CERRADOS
//  Mueve casos APROBADO/RECHAZADO con más de 3 meses a una hoja histórica.
// ===========================================================================

const MESES_ARCHIVO = 3;
const HOJA_HISTORICO = "Histórico Casos Cerrados";

function archivarCasosCerrados() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ahora = new Date();
  const limite = new Date(ahora.getFullYear(), ahora.getMonth() - MESES_ARCHIVO, ahora.getDate());

  let hojaHist = ss.getSheetByName(HOJA_HISTORICO);
  if (!hojaHist) {
    hojaHist = ss.insertSheet(HOJA_HISTORICO);
  }

  let totalArchivados = 0;

  HOJAS_DESTINO.forEach(nombreHoja => {
    const hoja = ss.getSheetByName(nombreHoja);
    if (!hoja) return;

    const datos = hoja.getDataRange().getValues();
    if (datos.length <= 1) return;

    const encabezados = datos[0];

    if (hojaHist.getLastRow() === 0) {
      const headerHist = [...encabezados, "Origen", "Fecha Archivado"];
      hojaHist.getRange(1, 1, 1, headerHist.length).setValues([headerHist]);
      hojaHist.getRange(1, 1, 1, headerHist.length)
        .setFontWeight("bold").setBackground("#34495e").setFontColor("white");
      hojaHist.setFrozenRows(1);
    }

    const filasAEliminar = [];
    const filasParaArchivar = [];

    for (let i = 1; i < datos.length; i++) {
      const fila = datos[i];
      const estado = (fila[18] || "").toString().trim().toUpperCase();
      const notificar = (fila[20] || "").toString();

      if (estado !== "APROBADO" && estado !== "RECHAZADO") continue;

      let fechaReferencia = null;
      const matchFecha = notificar.match(/(\d{2})\/(\d{2})\/(\d{4})/);
      if (matchFecha) {
        fechaReferencia = new Date(matchFecha[3], parseInt(matchFecha[2], 10) - 1, matchFecha[1]);
      } else if (fila[0] instanceof Date) {
        fechaReferencia = fila[0];
      }

      if (fechaReferencia && fechaReferencia < limite) {
        filasParaArchivar.push([...fila, nombreHoja, new Date()]);
        filasAEliminar.push(i + 1);
      }
    }

    if (filasParaArchivar.length > 0) {
      const inicio = hojaHist.getLastRow() + 1;
      hojaHist.getRange(inicio, 1, filasParaArchivar.length, filasParaArchivar[0].length)
              .setValues(filasParaArchivar);
    }

    filasAEliminar.sort((a, b) => b - a).forEach(num => hoja.deleteRow(num));
    totalArchivados += filasParaArchivar.length;
  });

  Logger.log(`Archivado completado: ${totalArchivados} casos movidos al histórico.`);
  return totalArchivados;
}

function archivarAhora() {
  const total = archivarCasosCerrados();
  SpreadsheetApp.getUi().alert(`Archivado completado: ${total} caso(s) movido(s) al histórico.`);
}