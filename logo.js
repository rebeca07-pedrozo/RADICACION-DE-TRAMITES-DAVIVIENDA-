function generarBase64LogoCorreo() {
  const ID_IMAGEN = "PEGA_AQUI_EL_ID_DEL_LOGO";

  const archivo = DriveApp.getFileById(ID_IMAGEN);
  const blob = archivo.getBlob();
  const base64 = Utilities.base64Encode(blob.getBytes());
  const mimeType = blob.getContentType();
  const dataUri = "data:" + mimeType + ";base64," + base64;

  const doc = DocumentApp.create("LOGO_BASE64_TEMPORAL_" + new Date().getTime());
  doc.getBody().setText(dataUri);
  doc.saveAndClose();

  Logger.log("Ábrelo y copia TODO (Ctrl+A, Ctrl+C): " + doc.getUrl());
}