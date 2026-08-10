const HOJA = "Partidas";
const CAB = ["Fecha", "Nombre", "Discos", "Movimientos", "Minimos", "Perfecto",
             "Tiempo (s)", "Intentos de formula", "Formulas que probo",
             "Acerto la formula", "Respuesta del acertijo"];

function doPost(e) {
  try {
    var d = JSON.parse(e.postData.contents);
    var h = hoja();
    var nombre = String(d.name || "Anonimo").trim();
    var t = d.tipo || "partida";

    if (t === "formula") {
      var f = fila(h, nombre);
      var prev = String(h.getRange(f, 9).getValue() || "");
      var uno = d.formula + " (" + d.correcta + ")";
      var txt = prev ? prev + "  |  " + uno : uno;

      h.getRange(f, 9).setValue(txt);
      h.getRange(f, 8).setValue(txt.split("|").length).setNumberFormat("0");

      var acerto = d.yaAcerto === true || txt.indexOf("(Si") > -1;
      h.getRange(f, 10).setValue(acerto ? "Si" : d.correcta);

    } else if (t === "acertijo") {
      h.getRange(fila(h, nombre), 11).setValue(d.respuesta || "");

    } else {
      h.appendRow([new Date(), nombre, d.disks, d.moves, d.min,
                   d.perfect ? "Si" : "No", d.secs, "", "", "", ""]);
    }
    return ContentService.createTextOutput("ok");
  } catch (err) {
    return ContentService.createTextOutput("error: " + err);
  }
}

function doGet() {
  return ContentService.createTextOutput(
    "Conector v7. Registros: " + Math.max(0, hoja().getLastRow() - 1));
}

function fila(h, nombre) {
  var u = h.getLastRow();
  if (u >= 2) {
    var n = h.getRange(2, 2, u - 1, 1).getValues();
    for (var i = n.length - 1; i >= 0; i--) {
      if (String(n[i][0]).trim().toLowerCase() === nombre.toLowerCase()) return i + 2;
    }
  }
  h.appendRow([new Date(), nombre, "", "", "", "", "", "", "", "", ""]);
  return h.getLastRow();
}

function hoja() {
  var lib = SpreadsheetApp.getActiveSpreadsheet();
  var h = lib.getSheetByName(HOJA) || lib.insertSheet(HOJA);
  if (h.getLastRow() === 0) h.appendRow(CAB);
  return h;
}
