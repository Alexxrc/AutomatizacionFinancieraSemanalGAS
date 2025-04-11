function generarInformeFinancieroDesdeFacturacion() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetFacturacion = ss.getSheetByName("HISTÓRICO FACTURACIÓN");
  const sheetVentas = ss.getSheetByName("HISTÓRICO VENTAS");
  const sheetPrecios = ss.getSheetByName("BASE DATOS PRECIOS");
  const folderId = "FOLDER_ID_AQUI";

  const lastRow = sheetFacturacion.getLastRow();
  const datosFact = sheetFacturacion.getRange(lastRow, 1, 1, sheetFacturacion.getLastColumn()).getValues()[0];

  const fechaFactDate = new Date(datosFact[0]);
  const fechaFacturacion = Utilities.formatDate(fechaFactDate, "Europe/Madrid", "dd/MM/yyyy");

  const totalFacturado = parseFloat(datosFact[9]) || 0;
  const gastosSeguridad = parseFloat(datosFact[10]) || 0;
  const gastosMusica = parseFloat(datosFact[11]) || 0;
  const ebita = parseFloat(datosFact[12]) || 0;

  let ventas = sheetVentas.getDataRange().getValues();
  const cabecera = ventas[0].join("").toLowerCase();
  if (cabecera.includes("producto") || cabecera.includes("cantidad") || cabecera.includes("fecha")) {
    ventas = ventas.slice(1);
  }

  const ventasDelDia = ventas.filter(row => {
    const fechaVenta = new Date(row[3]);
    const fechaVentaStr = Utilities.formatDate(fechaVenta, "Europe/Madrid", "dd/MM/yyyy");
    return fechaVentaStr === fechaFacturacion;
  });

  if (ventasDelDia.length === 0) {
    Logger.log("No hay ventas para la fecha: " + fechaFacturacion);
    return;
  }

  const precios = sheetPrecios.getDataRange().getValues();
  const mapaPrecios = {};
  precios.forEach(row => {
    const producto = row[0].toString().trim().toUpperCase();
    const costo = parseFloat(row[1]) || 0;
    mapaPrecios[producto] = costo;
  });

  const resumen = {};
  let gastoProductos = 0;

  ventasDelDia.forEach(row => {
    const producto = row[1].toString().trim().toUpperCase();
    const cantidad = parseFloat(row[2]) || 0;
    const costoUnitario = mapaPrecios[producto] || 0;
    const gasto = cantidad * costoUnitario;

    if (!resumen[producto]) {
      resumen[producto] = { cantidad: 0, costoUnitario, totalGasto: 0 };
    }

    resumen[producto].cantidad += cantidad;
    resumen[producto].totalGasto += gasto;
    gastoProductos += gasto;
  });

  const ivaSobreProductos = gastoProductos * 0.21;
  const gastosGenerales = gastosSeguridad + gastosMusica;
  const gastoTotal = gastoProductos + ivaSobreProductos + gastosGenerales;
  const margenBruto = totalFacturado - gastoProductos;
  const margenNeto = totalFacturado - gastoTotal;

  let topProducto = "";
  let maxCantidad = 0;
  for (const prod in resumen) {
    if (resumen[prod].cantidad > maxCantidad) {
      maxCantidad = resumen[prod].cantidad;
      topProducto = prod;
    }
  }

  const doc = DocumentApp.create(`Informe_Financiero_${fechaFacturacion.replace(/\//g, "-")}`);
  const body = doc.getBody();

  body.appendParagraph(`Informe Financiero - Fecha: ${fechaFacturacion}`).setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph(`Total Facturado: ${totalFacturado.toFixed(2)} €`);
  body.appendParagraph(`Producto más vendido: ${topProducto} (${maxCantidad} uds)`);
  body.appendParagraph("");
  body.appendParagraph("Detalle por producto:").setHeading(DocumentApp.ParagraphHeading.HEADING2);

  const tableData = [["Producto", "Cantidad", "Costo Unitario (€)", "Total Gasto (€)"]];
  for (const producto in resumen) {
    const r = resumen[producto];
    tableData.push([
      producto,
      r.cantidad,
      r.costoUnitario.toFixed(2),
      r.totalGasto.toFixed(2)
    ]);
  }
  body.appendTable(tableData);

  body.appendParagraph("");
  body.appendParagraph("Resumen Financiero").setHeading(DocumentApp.ParagraphHeading.HEADING2);

  const tablaResumen = [
    ["Concepto", "Importe (€)"],
    ["Gasto por productos", gastoProductos.toFixed(2)],
    ["IVA sobre productos (21%)", ivaSobreProductos.toFixed(2)],
    ["Gastos generales (seguridad + música)", gastosGenerales.toFixed(2)],
    ["Gasto total real estimado", gastoTotal.toFixed(2)],
    ["", ""],
    ["Margen neto (facturado - gasto total)", margenNeto.toFixed(2)]
  ];
  body.appendTable(tablaResumen);

  doc.saveAndClose();

  const pdf = DriveApp.getFileById(doc.getId()).getAs("application/pdf");
  const folder = DriveApp.getFolderById(folderId);
  folder.createFile(pdf).setName(`Informe_Financiero_${fechaFacturacion.replace(/\//g, "-")}.pdf`);
  DriveApp.getFileById(doc.getId()).setTrashed(true);
}
