function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Dashboard Gestión de Recursos')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getDataConsolidada() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaSolicitudes = ss.getSheetByName("Solicitudes");
  const hojaItems = ss.getSheetByName("Ítems");

  if (!hojaSolicitudes || !hojaItems) {
    throw new Error("No se encontraron las hojas 'Solicitudes' e 'Ítems'. Verifique nombres.");
  }

  // Leer la última fecha de actualización guardada en la celda Z1
  let fechaAct = hojaSolicitudes.getRange("Z1").getDisplayValue() || "";

  const dataS = hojaSolicitudes.getDataRange().getValues();
  const dataI = hojaItems.getDataRange().getValues();

  let solicitudes = [];
  let items = [];

  // Convertidor universal de fechas a Timestamp
  function parseFechaUniversal(val) {
    if (!val) return null;
    if (val instanceof Date) return isNaN(val.getTime()) ? null : val.getTime();
    
    // Si la fecha viene como Texto "DD/MM/YYYY" o "DD-MM-YYYY"
    if (typeof val === 'string') {
      let partes = val.trim().split(/[\/\-]/);
      if (partes.length === 3) {
        let dia = parseInt(partes[0], 10);
        let mes = parseInt(partes[1], 10) - 1;
        let anio = parseInt(partes[2], 10);
        if (anio < 100) anio += 2000;
        let d = new Date(anio, mes, dia);
        if (!isNaN(d.getTime())) return d.getTime();
      }
    }
    
    let d = new Date(val);
    return !isNaN(d.getTime()) ? d.getTime() : null;
  }

  // 1. LECTURA DE SOLICITUDES
  for (let i = 1; i < dataS.length; i++) {
    let id = dataS[i][0]; // Columna A
    if (!id) continue;

    let centroRaw = String(dataS[i][3] || ""); // Columna D
    let centro = "Sin Centro";
    if (centroRaw.includes("Antofagasta")) centro = "Antofagasta";
    else if (centroRaw.includes("Santiago")) centro = "Santiago";
    else if (centroRaw.includes("Concepción") || centroRaw.includes("Concepcion")) centro = "Concepción";

    solicitudes.push({
      id: id,
      tipo: String(dataS[i][1] || "N/A").trim(),  // Col B
      clase: String(dataS[i][2] || "N/A").trim(), // Col C
      centro: centro,                             // Col D
      estado: String(dataS[i][4] || "").trim(),   // Col E
      proyecto: dataS[i][5] || "N/A",             // Col F
      fechaSol: parseFechaUniversal(dataS[i][8]),   // Col I (Solicitud)
      fechaAprob: parseFechaUniversal(dataS[i][9]), // Col J (Aprobación)
      fechaTermino: parseFechaUniversal(dataS[i][10]) // Col K (Término)
    });
  }

  // 2. LECTURA DE ÍTEMS
  for (let j = 1; j < dataI.length; j++) {
    let idSol = dataI[j][1]; // Col B en Ítems
    if (!idSol) continue;

    items.push({
      idSolicitud: idSol,
      estadoItem: String(dataI[j][3] || "").trim(),          // Col D (Estado Ítem)
      colK: String(dataI[j][10] || "").trim().toLowerCase(), // Col K (Condición "no")
      fechaCotizacion: parseFechaUniversal(dataI[j][12]),    // Col M
      fechaOC: parseFechaUniversal(dataI[j][13]),            // Col N
      fechaSisRecep: parseFechaUniversal(dataI[j][19]),      // Col T
      fechaIngRecep: parseFechaUniversal(dataI[j][20]),      // Col U
      fechaSisEntrega: parseFechaUniversal(dataI[j][27])     // Col AB
    });
  }

  // Retornar los datos junto con la variable fechaActualizacion
  return {
    solicitudes: solicitudes,
    items: items,
    fechaActualizacion: fechaAct
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