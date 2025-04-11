function registrarVentasEnHistorico(datos, fecha) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = ss.getSheetByName("HISTÓRICO VENTAS");

  if (!hoja) {
    hoja = ss.insertSheet("HISTÓRICO VENTAS");
    hoja.appendRow(["Día", "Producto", "Vendidos", "Fecha", "Año", "Semana", "Mes"]);
  }

  const fechaVenta = new Date(fecha);
  fechaVenta.setDate(fechaVenta.getDate() - 1);

  const diaNombre = Utilities.formatDate(fechaVenta, "Europe/Madrid", "EEEE");
  const anio = fechaVenta.getFullYear();
  const semana = Math.ceil(fechaVenta.getDate() / 7);
  const fechaStr = Utilities.formatDate(fechaVenta, "Europe/Madrid", "yyyy-MM-dd");
  const mes = Utilities.formatDate(fechaVenta, "Europe/Madrid", "MMMM").toUpperCase();

  const registros = datos.map(item => [
    diaNombre,
    item.producto,
    item.vendidos,
    fechaStr,
    anio,
    semana,
    mes
  ]);

  hoja.getRange(hoja.getLastRow() + 1, 1, registros.length, registros[0].length).setValues(registros);
}
