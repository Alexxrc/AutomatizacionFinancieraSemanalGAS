function generarAnalisisSemanal() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaPedidos = ss.getSheetByName("RAW PEDIDOS");
  const hojaBarra = ss.getSheetByName("RAW BARRA");
  const hojaAlmacen = ss.getSheetByName("RAW ALMACÉN");

  const headers = hojaPedidos.getRange(1, 1, 1, hojaPedidos.getLastColumn()).getValues()[0];
  const filaPedido = hojaPedidos.getLastRow();
  const pedido = hojaPedidos.getRange(filaPedido, 1, 1, headers.length).getValues()[0];

  const totalFilas = hojaBarra.getLastRow();

  const barraInicial = hojaBarra.getRange(totalFilas - 2, 1, 1, headers.length).getValues()[0];
  const almacenInicial = hojaAlmacen.getRange(totalFilas - 2, 1, 1, headers.length).getValues()[0];
  const barraFinal = hojaBarra.getRange(totalFilas, 1, 1, headers.length).getValues()[0];
  const almacenFinal = hojaAlmacen.getRange(totalFilas, 1, 1, headers.length).getValues()[0];

  const fecha = new Date();
  const semana = Math.ceil(fecha.getDate() / 7);
  const mes = Utilities.formatDate(fecha, "Europe/Madrid", "MMMM").toUpperCase();
  const anio = Utilities.formatDate(fecha, "Europe/Madrid", "yyyy");
  const titulo = `ANÁLISIS SEMANAL - SEMANA ${semana} ${mes} ${anio}`;

  const hojaInforme = ss.insertSheet("Análisis Semanal " + semana + " " + mes + " " + anio);

  hojaInforme.getRange("A1:F1").merge().setValue(titulo)
    .setFontSize(16).setFontWeight("bold").setFontColor("white")
    .setBackground("#455A64").setHorizontalAlignment("center");

  hojaInforme.getRange("A2:F2").setValues([[
    "Producto", "Stock Inicial", "Pedido", "Barra Final", "Almacén Final", "Vendidos Est."
  ]]).setFontWeight("bold").setFontColor("white").setBackground("#607D8B").setHorizontalAlignment("center");

  for (let i = 0; i < headers.length; i++) {
    const producto = headers[i];
    if (producto === "Marca temporal") continue;

    const pedidoSem = Number(pedido[i]) || 0;
    const stockInicial = (Number(barraInicial[i]) || 0) + (Number(almacenInicial[i]) || 0);
    const barra = Number(barraFinal[i]) || 0;
    const almacen = Number(almacenFinal[i]) || 0;
    const stockFinal = barra + almacen;
    const vendidos = (stockInicial + pedidoSem) - stockFinal;

    const fila = hojaInforme.getLastRow() + 1;
    hojaInforme.appendRow([
      producto, stockInicial, pedidoSem, barra, almacen, vendidos
    ]);

    hojaInforme.getRange(fila, 1).setFontWeight("bold").setBackground("#CFD8DC").setFontSize(12);
    hojaInforme.getRange(fila, 6).setFontColor("#D84315").setFontWeight("bold").setFontSize(14);
  }

  hojaInforme.autoResizeColumns(1, 6);
  hojaInforme.getRange(3, 1, hojaInforme.getLastRow() - 2, 6).setWrap(true);

  const url = ss.getUrl().replace(/edit$/, '') + 'export?format=pdf&gid=' + hojaInforme.getSheetId();
  const opciones = { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } };
  const blob = UrlFetchApp.fetch(url, opciones).getBlob().setName('Analisis_Semanal_' + semana + "_" + mes + "_" + anio + '.pdf');

  MailApp.sendEmail({
    to: "CORREO_DESTINO_AQUI",
    subject: 'Análisis Semanal de Ventas',
    body: 'Adjunto el análisis semanal de ventas e inventario.',
    attachments: [blob]
  });

  ss.deleteSheet(hojaInforme);
}
