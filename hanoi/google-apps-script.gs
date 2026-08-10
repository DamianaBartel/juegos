/**
 * CONECTOR: Torre de Hanói  ->  Google Sheets   (v4)
 *
 * Usa las 11 columnas que ya tiene la planilla:
 *  A Fecha | B Nombre | C Discos | D Movimientos | E Mínimos | F Perfecto |
 *  G Tiempo (s) | H Intentos de fórmula | I Fórmulas que probó |
 *  J ¿Acertó la fórmula? | K Respuesta del acertijo
 *
 * Novedad de esta versión: "¿Acertó la fórmula?" queda en "Sí" apenas
 * acierta una vez, aunque después pruebe otras fórmulas equivocadas.
 */

const HOJA = "Partidas";
const ENCABEZADOS = [
  "Fecha", "Nombre", "Discos", "Movimientos", "Mínimos", "Perfecto",
  "Tiempo (s)", "Intentos de fórmula", "Fórmulas que probó",
  "¿Acertó la fórmula?", "Respuesta del acertijo"
];

const COL_NOMBRE = 2, COL_INTENTOS = 8, COL_FORMULAS = 9,
      COL_ACERTO = 10, COL_ACERTIJO = 11;

function doPost(e) {
  try {
    const datos = JSON.parse(e.postData.contents);
    const tipo = datos.tipo || "partida";
    const h = hoja();
    const nombre = (datos.name || "Anónimo").toString().trim();

    if (tipo === "formula") {
      const fila = filaDelAlumno(h, nombre);
      const previas = String(h.getRange(fila, COL_FORMULAS).getValue() || "");
      const nueva = datos.formula + " (" + datos.correcta + ")";
      const texto = previas ? previas + "  |  " + nueva : nueva;

      h.getRange(fila, COL_FORMULAS).setValue(texto);
      h.getRange(fila, COL_INTENTOS).setValue(texto.split("|").length).setNumberFormat("0");

      // Si en TODA la lista de fórmulas hay alguna con "(Sí)", entonces acertó.
      // Así no se pisa aunque después pruebe otras equivocadas.
      if (texto.indexOf("(Sí)") !== -1) {
        h.getRange(fila, COL_ACERTO).setValue("Sí");
      } else {
        h.getRange(fila, COL_ACERTO).setValue(datos.correcta);
      }

    } else if (tipo === "acertijo") {
      const fila = filaDelAlumno(h, nombre);
      h.getRange(fila, COL_ACERTIJO).setValue(datos.respuesta || "");

    } else {
      h.appendRow([
        new Date(), nombre, datos.disks, datos.moves, datos.min,
        datos.perfect ? "Sí" : "No", datos.secs, "", "", "", ""
      ]);
    }
    return responder({ ok: true });
  } catch (err) {
    return responder({ ok: false, error: String(err) });
  }
}

/** Abrir la URL en el navegador muestra esto: sirve para probar que quedó bien. */
function doGet() {
  const filas = Math.max(0, hoja().getLastRow() - 1);
  return ContentService.createTextOutput(
    "OK - Conector v4. Registros: " + filas
  );
}

/** Busca la última fila del alumno. Si no tiene ninguna, le crea una. */
function filaDelAlumno(h, nombre) {
  const ultima = h.getLastRow();
  if (ultima >= 2) {
    const nombres = h.getRange(2, COL_NOMBRE, ultima - 1, 1).getValues();
    for (let i = nombres.length - 1; i >= 0; i--) {
      if (String(nombres[i][0]).trim().toLowerCase() === nombre.toLowerCase()) {
        return i + 2;
      }
    }
  }
  h.appendRow([new Date(), nombre, "", "", "", "", "", "", "", "", ""]);
  return h.getLastRow();
}

/** Devuelve la hoja, creándola con encabezados si todavía no existe. */
function hoja() {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  let h = libro.getSheetByName(HOJA) || libro.insertSheet(HOJA);
  if (h.getLastRow() === 0) {
    h.appendRow(ENCABEZADOS);
    h.getRange(1, 1, 1, ENCABEZADOS.length).setFontWeight("bold");
    h.setFrozenRows(1);
  }
  return h;
}

function responder(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
