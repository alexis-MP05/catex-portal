// ================================================================
// CATEX Onboarding Portal — Google Apps Script
// Proyecto: ingresosCATEX
// Última actualización: 2026
// ================================================================

var SHEET_ID   = "1WUzD1Jn8fShjfHij7UObjlY1x2kZuE-87VbVs5YCTVc";
var SHEET_NAME = "CATEX Drivers";

var HEADERS = [
  "Fecha ingreso", "Nombre", "RUT", "Email", "Teléfono",
  "Ciudad", "Local", "Patente", "Marca", "Modelo",
  "Alto cm", "Largo cm", "Ancho cm",
  "Chofer", "RUT Chofer", "Cel Chofer",
  "Días disponibles", "Estado", "Beetrack"
];

// ----------------------------------------------------------------
// Obtener o crear la hoja de destino
// ----------------------------------------------------------------
function getSheet() {
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  return sheet;
}

// ----------------------------------------------------------------
// Crear encabezados si la hoja está vacía
// ----------------------------------------------------------------
function ensureHeaders(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    var headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#0F6E56");
    headerRange.setFontColor("#FFFFFF");
    headerRange.setFontFamily("Arial");
    headerRange.setFontSize(10);
    sheet.setFrozenRows(1);

    // Anchos de columna sugeridos
    sheet.setColumnWidth(1, 160);  // Fecha
    sheet.setColumnWidth(2, 180);  // Nombre
    sheet.setColumnWidth(3, 110);  // RUT
    sheet.setColumnWidth(4, 200);  // Email
    sheet.setColumnWidth(5, 130);  // Teléfono
    sheet.setColumnWidth(6, 120);  // Ciudad
    sheet.setColumnWidth(7, 70);   // Local
    sheet.setColumnWidth(8, 80);   // Patente
    sheet.setColumnWidth(9, 110);  // Marca
    sheet.setColumnWidth(10, 130); // Modelo
    sheet.setColumnWidth(11, 70);  // Alto
    sheet.setColumnWidth(12, 70);  // Largo
    sheet.setColumnWidth(13, 70);  // Ancho
    sheet.setColumnWidth(14, 160); // Chofer
    sheet.setColumnWidth(15, 110); // RUT Chofer
    sheet.setColumnWidth(16, 130); // Cel Chofer
    sheet.setColumnWidth(17, 160); // Días
    sheet.setColumnWidth(18, 90);  // Estado
    sheet.setColumnWidth(19, 90);  // Beetrack
  }
}

// ----------------------------------------------------------------
// Guardar una fila con los datos del driver
// ----------------------------------------------------------------
function guardarDatos(data) {
  var sheet = getSheet();
  ensureHeaders(sheet);

  var fila = [
    new Date().toLocaleString("es-CL"),
    data.nombre   || "",
    data.rut      || "",
    data.email    || "",
    data.tel      || "",
    data.ciudad   || "",
    data.local    || "",
    data.patente  || "",
    data.marca    || "",
    data.modelo   || "",
    data.alto     || "",
    data.largo    || "",
    data.ancho    || "",
    data.chofer      || "",
    data.chofer_rut  || "",
    data.chofer_cel  || "",
    Array.isArray(data.dias) ? data.dias.join(" - ") : (data.dias || ""),
    data.estado   || "completo",
    data.beetrack ? "Sí" : "No"
  ];

  sheet.appendRow(fila);

  // Alternar color de fila para legibilidad
  var lastRow = sheet.getLastRow();
  if (lastRow % 2 === 0) {
    sheet.getRange(lastRow, 1, 1, HEADERS.length).setBackground("#F0FAF6");
  }

  return true;
}

// ----------------------------------------------------------------
// doGet — maneja GET (con o sin parámetro data)
// ----------------------------------------------------------------
function doGet(e) {
  // CORS headers
  var output;

  // Si llega el parámetro "data", guardar registro
  if (e && e.parameter && e.parameter.data) {
    try {
      var data = JSON.parse(e.parameter.data);
      guardarDatos(data);
      output = ContentService
        .createTextOutput(JSON.stringify({ ok: true, message: "Registro guardado correctamente." }))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      output = ContentService
        .createTextOutput(JSON.stringify({ ok: false, error: err.toString() }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    return output;
  }

  // Sin parámetros: respuesta de estado (health check)
  return ContentService
    .createTextOutput(JSON.stringify({ status: "CATEX API activa" }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ----------------------------------------------------------------
// doPost — maneja POST con JSON en el body
// ----------------------------------------------------------------
function doPost(e) {
  try {
    var data;

    // Intentar parsear body como JSON
    if (e && e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
    }
    // También aceptar parámetro "data" en POST
    else if (e && e.parameter && e.parameter.data) {
      data = JSON.parse(e.parameter.data);
    }
    else {
      return ContentService
        .createTextOutput(JSON.stringify({ ok: false, error: "No se recibieron datos." }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    guardarDatos(data);

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, message: "Registro guardado correctamente." }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ----------------------------------------------------------------
// TEST — ejecutar manualmente desde el editor para probar
// ----------------------------------------------------------------
function testRegistro() {
  var dataPrueba = {
    nombre:     "Juan Pérez González",
    rut:        "12.345.678-9",
    email:      "juan.perez@test.com",
    tel:        "+56912345678",
    ciudad:     "Santiago",
    local:      "RM",
    patente:    "ABCD12",
    marca:      "Mercedes-Benz",
    modelo:     "Sprinter 315",
    alto:       "180",
    largo:      "350",
    ancho:      "200",
    chofer:     "Carlos Pérez",
    chofer_rut: "11.111.111-1",
    chofer_cel: "+56911111111",
    dias:       ["Lun", "Mar", "Mié", "Jue", "Vie"],
    estado:     "completo",
    beetrack:   false
  };

  var sheet = getSheet();
  ensureHeaders(sheet);
  guardarDatos(dataPrueba);
  Logger.log("✅ Registro de prueba guardado en hoja: " + SHEET_NAME);
}
