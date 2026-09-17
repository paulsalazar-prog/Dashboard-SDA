function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Dashboard SDA')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Carga y consolida los datos de 'Solicitudes' e 'Ítems'
 */
function getDataConsolidada() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Lectura de la pestaña Solicitudes
  var sheetSol = ss.getSheetByName("Solicitudes");
  var dataSol = sheetSol ? sheetSol.getDataRange().getValues() : [];
  var solicitudes = [];

  for (var i = 1; i < dataSol.length; i++) {
    if (!dataSol[i][0]) continue; // ID Solicitud
    
    solicitudes.push({
      id: String(dataSol[i][0]),
      fechaSol: parseFechaMs(dataSol[i][1]),
      fechaAprob: parseFechaMs(dataSol[i][2]),
      fechaEntregado: parseFechaMs(dataSol[i][3]),
      estado: String(dataSol[i][4] || ''),
      centro: String(dataSol[i][5] || ''),
      solicitante: String(dataSol[i][7] || '')
    });
  }

  // 2. Lectura de la pestaña Ítems
  var sheetItem = ss.getSheetByName("Ítems");
  var dataItem = sheetItem ? sheetItem.getDataRange().getValues() : [];
  var items = [];

  for (var j = 1; j < dataItem.length; j++) {
    if (!dataItem[j][0]) continue;

    items.push({
      idItem: String(dataItem[j][0]),
      idSolicitud: String(dataItem[j][1]),
      tipo: String(dataItem[j][2] || ''),
      fechaCompra: parseFechaMs(dataItem[j][3]),
      fechaProv: parseFechaMs(dataItem[j][4]),
      fechaEntrega: parseFechaMs(dataItem[j][5])
    });
  }

  return {
    solicitudes: solicitudes,
    items: items
  };
}

/**
 * Parsea fechas a milisegundos para filtrado dinámico
 */
function parseFechaMs(val) {
  if (!val) return null;
  if (val instanceof Date) return val.getTime();
  var d = new Date(val);
  return isNaN(d.getTime()) ? null : d.getTime();
}

/**
 * NAVEGACIÓN Y HISTÓRICO: Guarda captura plana semanal en 'Historico_SDA'
 */
function guardarFotoSemanalPlana() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojaHistorico = ss.getSheetByName("Historico_SDA");
  
  if (!hojaHistorico) {
    hojaHistorico = ss.insertSheet("Historico_SDA");
    hojaHistorico.appendRow(["Fecha_Snapshot", "ID_Solicitud", "Centro", "Tipo", "Estado"]);
    hojaHistorico.getRange("1:1").setFontWeight("bold");
  }

  var datos = getDataConsolidada();
  var solicitudes = datos.solicitudes || [];
  var items = datos.items || [];
  
  var mapaTipos = {};
  items.forEach(function(item) {
    if (item.idSolicitud && !mapaTipos[item.idSolicitud]) {
      mapaTipos[item.idSolicitud] = item.tipo;
    }
  });

  var fechaHoy = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "dd/MM/yyyy");
  var filasEscribir = [];

  solicitudes.forEach(function(s) {
    filasEscribir.push([
      fechaHoy,
      s.id,
      s.centro || "Sin Centro",
      mapaTipos[s.id] || "Sin Tipo",
      s.estado || "Sin Estado"
    ]);
  });

  if (filasEscribir.length > 0) {
    hojaHistorico.getRange(hojaHistorico.getLastRow() + 1, 1, filasEscribir.length, 5).setValues(filasEscribir);
  }
}

/**
 * Lee la pestaña 'Historico_SDA' y la envía al Frontend para Chart.js
 */
function getDatosHistoricos() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName("Historico_SDA");
    if (!hoja || hoja.getLastRow() <= 1) return [];

    var data = hoja.getDataRange().getValues();
    var historico = [];

    for (var i = 1; i < data.length; i++) {
      historico.push({
        fecha: data[i][0] instanceof Date ? Utilities.formatDate(data[i][0], ss.getSpreadsheetTimeZone(), "dd/MM/yyyy") : String(data[i][0]),
        id: String(data[i][1]),
        centro: String(data[i][2]),
        tipo: String(data[i][3]),
        estado: String(data[i][4])
      });
    }
    return historico;
  } catch (e) {
    Logger.log("Error en getDatosHistoricos: " + e.toString());
    return [];
  }
}