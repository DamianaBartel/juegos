/**
 * CONECTOR: Torre de Hanói  ->  Google Sheets
 *
 * Este código va pegado en una planilla de Google (Extensiones > Apps Script).
 * Recibe cada partida ganada y la agrega como una fila nueva.
 *
 * PASOS (una sola vez):
 *  1. Creá una planilla nueva en https://sheets.new
 *  2. Menú  Extensiones > Apps Script
 *  3. Borrá todo lo que haya y pegá ESTE archivo completo
 *  4. Botón  Implementar > Nueva implementación
 *       - Tipo: Aplicación web
 *       - Ejecutar como: Yo
 *       - Quién tiene acceso: CUALQUIER PERSONA   <-- importante
 *  5. Copiá la URL que termina en /exec y pasásela a Claude
 */

const NOMBRE_HOJA = "Partidas";

/** Recibe una partida desde el juego y la guarda. */
function doPost(e) {
  try {
    const datos = JSON.parse(e.postData.contents);
    const hoja = obtenerHoja();
    hoja.appendRow([
      new Date(),                        // fecha del servidor (confiable)
      datos.name || "Anónimo",
      datos.disks,
      datos.moves,
      datos.min,
      datos.perfect ? "Sí" : "No",
      datos.secs,
      datos.date || ""                   // fecha según la compu del alumno
    ]);
    return responder({ ok: true });
  } catch (err) {
    return responder({ ok: false, error: String(err) });
  }
}

/** Abrir la URL en el navegador muestra esto: sirve para probar que quedó bien. */
function doGet() {
  const hoja = obtenerHoja();
  const partidas = Math.max(0, hoja.getLastRow() - 1);
  return ContentService.createTextOutput(
    "OK - El conector funciona. Partidas registradas: " + partidas
  );
}

/** Devuelve la hoja, creándola con encabezados si todavía no existe. */
function obtenerHoja() {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = libro.getSheetByName(NOMBRE_HOJA) || libro.insertSheet(NOMBRE_HOJA);
  if (hoja.getLastRow() === 0) {
    hoja.appendRow([
      "Fecha", "Nombre", "Discos", "Movimientos",
      "Mínimos", "Perfecto", "Tiempo (s)", "Fecha del alumno"
    ]);
    hoja.getRange(1, 1, 1, 8).setFontWeight("bold");
    hoja.setFrozenRows(1);
  }
  return hoja;
}

function responder(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
