/**
 * CONECTOR: Torre de Hanói  ->  Google Sheets
 *
 * Guarda tres cosas en tres pestañas distintas de la misma planilla:
 *   Partidas  -> cada partida ganada
 *   Fórmulas  -> cada intento de fórmula (cuál escribió y si acertó)
 *   Acertijo  -> la respuesta a la pregunta de la leyenda
 *
 * ---------------------------------------------------------------
 * CÓMO ACTUALIZARLO (si ya lo tenías funcionando):
 *   1. Abrí tu planilla > Extensiones > Apps Script
 *   2. Borrá todo y pegá este archivo completo. Guardá (💾)
 *   3. Implementar > ADMINISTRAR IMPLEMENTACIONES
 *   4. Clic en el lápiz (editar) de la implementación que ya existe
 *   5. En "Versión" elegí NUEVA VERSIÓN y apretá Implementar
 *
 *   ⚠ Hacelo así (editando la que ya existe) para que la URL NO cambie.
 *      Si creás una implementación nueva, la URL cambia y hay que
 *      actualizarla también en el juego.
 * ---------------------------------------------------------------
 */

/** Recibe los datos del juego y los guarda donde corresponde. */
function doPost(e) {
  try {
    const datos = JSON.parse(e.postData.contents);
    const tipo = datos.tipo || "partida";

    if (tipo === "formula") {
      hoja("Fórmulas", ["Fecha", "Nombre", "Intento N°", "Fórmula que escribió", "¿Acertó?"])
        .appendRow([
          new Date(),
          datos.name || "Anónimo",
          datos.intento || "",
          datos.formula || "",
          datos.correcta
        ]);

    } else if (tipo === "acertijo") {
      hoja("Acertijo", ["Fecha", "Nombre", "Respuesta a la leyenda"])
        .appendRow([
          new Date(),
          datos.name || "Anónimo",
          datos.respuesta || ""
        ]);

    } else {
      hoja("Partidas", ["Fecha", "Nombre", "Discos", "Movimientos",
                        "Mínimos", "Perfecto", "Tiempo (s)", "Fecha del alumno"])
        .appendRow([
          new Date(),
          datos.name || "Anónimo",
          datos.disks,
          datos.moves,
          datos.min,
          datos.perfect ? "Sí" : "No",
          datos.secs,
          datos.date || ""
        ]);
    }
    return responder({ ok: true });
  } catch (err) {
    return responder({ ok: false, error: String(err) });
  }
}

/** Abrir la URL en el navegador muestra esto: sirve para probar que quedó bien. */
function doGet() {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  const contar = function (nombre) {
    const h = libro.getSheetByName(nombre);
    return h ? Math.max(0, h.getLastRow() - 1) : 0;
  };
  return ContentService.createTextOutput(
    "OK - El conector funciona." +
    " Partidas: " + contar("Partidas") +
    " | Fórmulas: " + contar("Fórmulas") +
    " | Acertijo: " + contar("Acertijo")
  );
}

/** Devuelve la pestaña pedida, creándola con sus encabezados si no existe. */
function hoja(nombre, encabezados) {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  let h = libro.getSheetByName(nombre) || libro.insertSheet(nombre);
  if (h.getLastRow() === 0) {
    h.appendRow(encabezados);
    h.getRange(1, 1, 1, encabezados.length).setFontWeight("bold");
    h.setFrozenRows(1);
  }
  return h;
}

function responder(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
