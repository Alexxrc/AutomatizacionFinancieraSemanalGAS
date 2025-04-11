function obtenerFilasPorTipoDia(tipo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaBarra = ss.getSheetByName("RAW BARRA");
  const totalFilas = hojaBarra.getLastRow();

  if (tipo === "viernes" || tipo === "sabado") {
    return {
      stockInicial: totalFilas - 1,
      stockFinal: totalFilas
    };
  }

  throw new Error("Tipo de día inválido. Usa 'viernes' o 'sabado'.");
}

function guardarInformeEnCarpeta(blob, folderLink) {
  const folderId = folderLink.split("/")[5];
  const folder = DriveApp.getFolderById(folderId);
  return folder.createFile(blob);
}

function generarInformeSabado() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaPedidos = ss.getSheetByName("RAW PEDIDOS");
  const hojaBarra = ss.getSheetByName("RAW BARRA");
  const hojaAlmacen = ss.getSheetByName("RAW ALMACÉN");

  const headers = hojaPedidos.getRange(1, 1, 1, hojaPedidos.getLastColumn()).getValues()[0];
  const { stockInicial, stockFinal } = obtenerFilasPorTipoDia("sabado");

  const barraInicial = hojaBarra.getRange(stockInicial, 1, 1, headers.length).getValues()[0];
  const almacenInicial = hojaAlmacen.getRange(stockInicial, 1, 1, headers.length).getValues()[0];
  const barraFinal = hojaBarra.getRange(stockFinal, 1, 1, headers.length).getValues()[0];
  const almacenFinal = hojaAlmacen.getRange(stockFinal, 1, 1, headers.length).getValues()[0];

  const ssFecha = Utilities.formatDate(new Date(), "Europe/Madrid", "yyyy-MM-dd");
  const hojaInforme = ss.insertSheet("Informe Sábado " + ssFecha);

  hojaInforme.getRange("A1:E1").merge().setValue("INFORME DE VENTAS - SÁBADO " + ssFecha)
    .setFontSize(16).setFontWeight("bold").setFontColor("white")
    .setBackground("#455A64").setHorizontalAlignment("center");

  hojaInforme.getRange("A2:E2").setValues([[
    "Producto", "Stock Inicial", "Barra Final", "Almacén Final", "Vendidos"
  ]]).setFontWeight("bold").setFontColor("white").setBackground("#607D8B").setHorizontalAlignment("center");

  const datosHistorico = [];

  for (let i = 0; i < headers.length; i++) {
    const producto = headers[i];
    if (producto === "Marca temporal") continue;

    const inicial = (Number(barraInicial[i]) || 0) + (Number(almacenInicial[i]) || 0);
    const final = (Number(barraFinal[i]) || 0) + (Number(almacenFinal[i]) || 0);
    const vendidos = Math.max(0, inicial - final);

    const fila = hojaInforme.getLastRow() + 1;
    hojaInforme.appendRow([
      producto, inicial, barraFinal[i], almacenFinal[i], vendidos
    ]);

    hojaInforme.getRange(fila, 1).setFontWeight("bold").setBackground("#CFD8DC").setFontSize(12);
    hojaInforme.getRange(fila, 5).setFontColor("#D84315").setFontWeight("bold").setFontSize(14);

    datosHistorico.push({ producto, vendidos });
  }

  hojaInforme.autoResizeColumns(1, 5);
  hojaInforme.getRange(3, 1, hojaInforme.getLastRow() - 2, 5).setWrap(true);

  registrarVentasEnHistorico(datosHistorico, new Date(), "Sábado");

  const url = ss.getUrl().replace(/edit$/, '') + 'export?format=pdf&gid=' + hojaInforme.getSheetId();
  const opciones = { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } };
  const blob = UrlFetchApp.fetch(url, opciones).getBlob().setName('Ventas_Sabado_' + ssFecha + '.pdf');

  const folderLink = 'https://drive.google.com/drive/folders/FOLDER_ID_AQUI';
  guardarInformeEnCarpeta(blob, folderLink);

  MailApp.sendEmail({
    to: "CORREO_DESTINO_AQUI",
    subject: 'Ventas del Sábado',
    body: 'Adjunto el informe de ventas del sábado.',
    attachments: [blob]
  });

  ss.deleteSheet(hojaInforme);
}
