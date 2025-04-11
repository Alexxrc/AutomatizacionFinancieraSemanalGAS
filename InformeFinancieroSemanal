function generarInformeFinancieroDesdeFacturacion() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetFacturacion = ss.getSheetByName("HISTÓRICO FACTURACIÓN");
  const sheetVentas = ss.getSheetByName("HISTÓRICO VENTAS");
  const sheetPrecios = ss.getSheetByName("BASE DATOS PRECIOS");
  const folderId = "FOLDER_ID_AQUI";

  const lastRow = sheetFacturacion.getLastRow();
  const datosFact = sheetFacturacion.getRange(lastRow - 1, 1, 2, sheetFacturacion.getLastColumn()).getValues();
  const fechasSemana = datosFact.map(row => Utilities.formatDate(new Date(row[0]), "Europe/Madrid", "dd/MM/yyyy"));
  const [datosViernes, datosSabado] = datosFact;

  const totalFacturadoViernes = parseFloat(datosViernes[9]) || 0;
  const totalFacturadoSabado = parseFloat(datosSabado[9]) || 0;

  const gastosSeguridadViernes = parseFloat(datosViernes[10]) || 0;
  const gastosMusicaViernes = parseFloat(datosViernes[11]) || 0;
  const gastosSeguridadSabado = parseFloat(datosSabado[10]) || 0;
  const gastosMusicaSabado = parseFloat(datosSabado[11]) || 0;

  const totalEBITA = (parseFloat(datosViernes[12]) || 0) + (parseFloat(datosSabado[12]) || 0);

  let ventas = sheetVentas.getDataRange().getValues();
  const cabecera = ventas[0].join("").toLowerCase();
  if (cabecera.includes("producto") || cabecera.includes("cantidad") || cabecera.includes("fecha")) ventas = ventas.slice(1);

  const ventasFiltradas = ventas.filter(row => {
    const fechaVenta = Utilities.formatDate(new Date(row[3]), "Europe/Madrid", "dd/MM/yyyy");
    return fechasSemana.includes(fechaVenta);
  });

  const precios = sheetPrecios.getDataRange().getValues();
  const mapaPrecios = {};
  precios.forEach(row => {
    const producto = row[0].toString().trim().toUpperCase();
    const costo = parseFloat(row[1]) || 0;
    mapaPrecios[producto] = costo;
  });

  const resumen = {};
  let gastoProductos = 0;

  ventasFiltradas.forEach(row => {
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

  const iva = gastoProductos * 0.21;
  const gastosGenerales = gastosSeguridadViernes + gastosMusicaViernes + gastosSeguridadSabado + gastosMusicaSabado;
  const gastoTotal = gastoProductos + iva + gastosGenerales;
  const totalFacturado = totalFacturadoViernes + totalFacturadoSabado;
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

  const doc = DocumentApp.create(`Informe_Semanal_${fechasSemana[0].replace(/\//g, "-")}_a_${fechasSemana[1].replace(/\//g, "-")}`);
  const body = doc.getBody();
  body.appendParagraph(`Informe Financiero Semanal`).setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph(`Fechas: ${fechasSemana[0]} (viernes) y ${fechasSemana[1]} (sábado)`);
  body.appendParagraph(`Producto más vendido: ${topProducto} (${maxCantidad} uds)`);
  body.appendParagraph("");
  body.appendParagraph("Detalle por producto:").setHeading(DocumentApp.ParagraphHeading.HEADING2);

  const tablaProductos = [["Producto", "Cantidad", "Costo Unitario (€)", "Total Gasto (€)"]];
  for (const producto in resumen) {
    const r = resumen[producto];
    tablaProductos.push([producto, r.cantidad, r.costoUnitario.toFixed(2), r.totalGasto.toFixed(2)]);
  }
  body.appendTable(tablaProductos);

  body.appendParagraph("");
  body.appendParagraph("Resumen Financiero").setHeading(DocumentApp.ParagraphHeading.HEADING2);
  const tablaResumen = [
    ["Concepto", "Importe (€)"],
    ["Facturado viernes", totalFacturadoViernes.toFixed(2)],
    ["Facturado sábado", totalFacturadoSabado.toFixed(2)],
    ["Gastos seguridad viernes", gastosSeguridadViernes.toFixed(2)],
    ["Gastos música viernes", gastosMusicaViernes.toFixed(2)],
    ["Gastos seguridad sábado", gastosSeguridadSabado.toFixed(2)],
    ["Gastos música sábado", gastosMusicaSabado.toFixed(2)],
    ["", ""],
    ["Gasto por productos (total)", gastoProductos.toFixed(2)],
    ["IVA sobre productos (21%)", iva.toFixed(2)],
    ["Gastos generales (seguridad + música)", gastosGenerales.toFixed(2)],
    ["Gasto total real estimado", gastoTotal.toFixed(2)],
    ["", ""],
    ["Margen bruto (facturado - gasto productos)", margenBruto.toFixed(2)],
    ["Margen neto (facturado - gasto total)", margenNeto.toFixed(2)],
  ];
  body.appendTable(tablaResumen);

  doc.saveAndClose();
  const pdf = DriveApp.getFileById(doc.getId()).getAs("application/pdf");
  const folder = DriveApp.getFolderById(folderId);
  folder.createFile(pdf).setName(`Informe_Semanal_${fechasSemana[0].replace(/\//g, "-")}_a_${fechasSemana[1].replace(/\//g, "-")}.pdf`);
  DriveApp.getFileById(doc.getId()).setTrashed(true);
}
