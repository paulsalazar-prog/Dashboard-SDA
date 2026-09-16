function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Dashboard Gestión de Recursos')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getDatosConsolidados() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetConsolidado = ss.getSheetByName("consolidado");
  var sheetSolicitud = ss.getSheetByName("Solicitudes"); // Hoja 'Solicitudes'
  
  if (!sheetConsolidado) {
    throw new Error("No se encontró la hoja 'consolidado'.");
  }

  // 1. Crear un mapa de Solicitante (Columna H = índice 7) indexado por el Código SR (Columna A = índice 0)
  var mapaSolicitantes = {};
  if (sheetSolicitud) {
    var dataSolicitud = sheetSolicitud.getDataRange().getValues();
    if (dataSolicitud.length > 1) {
      for (var i = 1; i < dataSolicitud.length; i++) {
        var cod = String(dataSolicitud[i][0]).trim(); // Columna A (Código SR)
        var nom = String(dataSolicitud[i][7]).trim(); // Columna H (Solicitante)
        if (cod) {
          mapaSolicitantes[cod] = nom;
        }
      }
    }
  }

  // 2. Leer la hoja 'consolidado'
  var data = sheetConsolidado.getDataRange().getValues();
  if (data.length <= 1) {
    return { status: "empty", data: [], timestamp: "" };
  }

  var headers = data[0].map(function(h) { return String(h).toLowerCase().trim(); });

  var idxId = headers.indexOf("código sr");
  if (idxId === -1) idxId = headers.indexOf("codigo sr");
  
  var idxProyecto = headers.indexOf("proyecto");
  var idxTipo = headers.indexOf("tipo");
  var idxClase = headers.indexOf("clase");
  var idxCentro = headers.indexOf("centro");
  var idxEstado = headers.indexOf("estado");
  var idxFAprob = headers.indexOf("fecha aprobación");
  var idxFTerm = headers.indexOf("fecha término");
  var idxFEnt = headers.indexOf("fecha entrega");

  var solicitudesMap = {};
  var itemsList = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var idSR = idxId !== -1 ? String(row[idxId]).trim() : "";
    if (!idSR) continue;

    var fAprob = idxFAprob !== -1 ? parseFechaJS(row[idxFAprob]) : null;
    var fTerm = idxFTerm !== -1 ? parseFechaJS(row[idxFTerm]) : null;
    var fEnt = idxFEnt !== -1 ? parseFechaJS(row[idxFEnt]) : null;

    if (!solicitudesMap[idSR]) {
      // Cruza el nombre del solicitante extraído de la Columna H de 'Solicitudes'
      var nombreSolicitante = mapaSolicitantes[idSR] || "";

      solicitudesMap[idSR] = {
        id: idSR,
        proyecto: idxProyecto !== -1 ? String(row[idxProyecto]) : "",
        solicitante: nombreSolicitante,
        centro: idxCentro !== -1 ? String(row[idxCentro]) : "",
        estado: idxEstado !== -1 ? String(row[idxEstado]) : "",
        fechaAprob: fAprob ? fAprob.getTime() : null,
        fechaTermino: fTerm ? fTerm.getTime() : null
      };
    }

    itemsList.push({
      idSR: idSR,
      tipo: idxTipo !== -1 ? String(row[idxTipo]) : "",
      clase: idxClase !== -1 ? String(row[idxClase]) : "",
      centro: idxCentro !== -1 ? String(row[idxCentro]) : "",
      estado: idxEstado !== -1 ? String(row[idxEstado]) : "",
      fechaAprob: fAprob ? fAprob.getTime() : null,
      fechaTermino: fTerm ? fTerm.getTime() : null,
      fechaEntrega: fEnt ? fEnt.getTime() : null
    });
  }

  var timestamp = "";
  try {
    var valZ1 = sheetConsolidado.getRange("Z1").getValue();
    if (valZ1) {
      if (valZ1 instanceof Date) {
        timestamp = Utilities.formatDate(valZ1, ss.getSpreadsheetTimeZone(), "dd-MM-yyyy HH:mm");
      } else {
        timestamp = String(valZ1);
      }
    }
  } catch(e) {}

  return {
    status: "success",
    solicitudes: Object.values(solicitudesMap),
    items: itemsList,
    timestamp: timestamp
  };
}

function actualizarDatosSDA() {
  var url = "https://api-sda.cydingenieria.com/v1/getDataExcel";
  
  var hoy = new Date();
  var hace6Meses = new Date();
  hace6Meses.setMonth(hoy.getMonth() - 6);

  Logger.log("Rango de consulta: Desde " + hace6Meses.toISOString().substring(0,10) + " Hasta " + hoy.toISOString().substring(0,10));

  var payload = {
    "desde": hace6Meses.toISOString(),
    "hasta": hoy.toISOString(),
    "centro": [1, 3, 2],
    "proyecto": [],
    "gerencia": [],
    "estado": [],
    "tipoRecurso": [],
    "usuario": "6657fb1aab835b00115961c4"
  };

  var opciones = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  try {
    Logger.log("1. Solicitando datos a la API del SDA...");
    var respuesta = UrlFetchApp.fetch(url, opciones);
    var code = respuesta.getResponseCode();

    if (code !== 200 && code !== 201) {
      throw new Error("Error en la API del SDA (HTTP " + code + ")");
    }

    var textoRespuesta = respuesta.getContentText();
    Logger.log("2. Respuesta recibida. Procesando JSON...");

    var objetoJSON = JSON.parse(textoRespuesta);
    var destSs = SpreadsheetApp.getActiveSpreadsheet();

    // 1. PROCESAR SOLICITUDES (dataSR)
    var matrizSolicitudes = objetoJSON.dataSR || objetoJSON.solicitudes || [];
    if (matrizSolicitudes.length > 0) {
      Logger.log("3. Escribiendo pestaña 'Solicitudes' (" + matrizSolicitudes.length + " filas)...");
      var hojaSol = destSs.getSheetByName("Solicitudes");
      if (hojaSol) {
        escribirConEncabezadosYFormato(hojaSol, matrizSolicitudes);
        Logger.log("¡Pestaña 'Solicitudes' actualizada correctamente!");
      }
    }

    // 2. PROCESAR ÍTEMS (dataItems / dataIT)
    var matrizItems = objetoJSON.dataItems || objetoJSON.dataIT || objetoJSON.items || [];
    if (matrizItems.length > 0) {
      Logger.log("4. Escribiendo pestaña 'Ítems' (" + matrizItems.length + " filas)...");
      var hojaItem = destSs.getSheetByName("Ítems") || destSs.getSheetByName("Items");
      if (hojaItem) {
        escribirConEncabezadosYFormato(hojaItem, matrizItems);
        Logger.log("¡Pestaña 'Ítems' actualizada correctamente!");
      }
    }

    // Guardar fecha y hora de actualización en la celda Z1
    var hojaSol = destSs.getSheetByName("Solicitudes");
    if (hojaSol) {
      var ahora = Utilities.formatDate(new Date(), destSs.getSpreadsheetTimeZone(), "dd-MM-yyyy HH:mm");
      hojaSol.getRange("Z1").setValue(ahora);
    }

    SpreadsheetApp.flush();
    Logger.log("\n5. PROCESO GLOBAL COMPLETADO CON ÉXITO.");

  } catch (error) {
    Logger.log("Error durante el procesamiento: " + error.toString());
  }
}

/**
 * Mantiene la Fila 1 con encabezados (en Negrita) y escribe
 * los datos de la Fila 2 en adelante con texto normal (sin Negrita).
 */
function escribirConEncabezadosYFormato(sheet, matrizDatos) {
  if (matrizDatos.length === 0) return;

  var tieneEncabezadoEnMatriz = isNaN(matrizDatos[0][0]) && typeof matrizDatos[0][0] === 'string' && !matrizDatos[0][0].includes("GIN");

  var datosAEscribir = matrizDatos;
  var encabezadosExistentes = [];

  if (sheet.getLastRow() > 0 && sheet.getLastColumn() > 0) {
    encabezadosExistentes = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  }

  sheet.clearContents();

  if (!tieneEncabezadoEnMatriz && encabezadosExistentes.length > 0) {
    // 1. Escribir Encabezado en Fila 1 con Negrita
    var rangoEncabezado = sheet.getRange(1, 1, 1, encabezadosExistentes.length);
    rangoEncabezado.setValues([encabezadosExistentes]);
    rangoEncabezado.setFontWeight("bold");

    // 2. Escribir Datos desde Fila 2 con Texto Normal (sin negrita)
    var rangoDatos = sheet.getRange(2, 1, datosAEscribir.length, datosAEscribir[0].length);
    rangoDatos.setValues(datosAEscribir);
    rangoDatos.setFontWeight("normal");
  } else {
    var rangoCompleto = sheet.getRange(1, 1, datosAEscribir.length, datosAEscribir[0].length);
    rangoCompleto.setValues(datosAEscribir);

    // Formatear Fila 1 con negrita y el resto normal
    sheet.getRange(1, 1, 1, datosAEscribir[0].length).setFontWeight("bold");
    if (datosAEscribir.length > 1) {
      sheet.getRange(2, 1, datosAEscribir.length - 1, datosAEscribir[0].length).setFontWeight("normal");
    }
  }
}