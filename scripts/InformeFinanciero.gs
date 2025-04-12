function generarInformeFinancieroDiario() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetFacturacion = ss.getSheetByName("HISTÓRICO FACTURACIÓN");
  const sheetVentas = ss.getSheetByName("HISTÓRICO VENTAS");
  const sheetPrecios = ss.getSheetByName("BASE DATOS PRECIOS");

  const lastRow = sheetFacturacion.getLastRow();
  const datosFact = sheetFacturacion.getRange(lastRow, 1, 1, sheetFacturacion.getLastColumn()).getValues()[0];

  const fechaFactDate = new Date(datosFact[0]);
  const fechaFacturacion = Utilities.formatDate(fechaFactDate, "Europe/Madrid", "dd/MM/yyyy");

  const totalFacturado = parseFloat(datosFact[9]) || 0;
  const gastosSeguridad = parseFloat(datosFact[10]) || 0;
  const gastosMusica = parseFloat(datosFact[11]) || 0;

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
  const gastoTotal = gastoProductos + ivaSobreProductos + gastosSeguridad + gastosMusica;
  const margenBruto = totalFacturado - gastoProductos;

  const doc = DocumentApp.create(`Informe_Financiero_${fechaFacturacion.replace(/\//g, "-")}`);
  const body = doc.getBody();

  body.appendParagraph(`Informe Financiero - Fecha: ${fechaFacturacion}`).setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph(`Total Facturado: ${totalFacturado.toFixed(2)} €`);
  body.appendParagraph("");

  // Resumen financiero
  const tablaResumen = [
  ["Concepto", "Importe (€)"],
  ["Total Facturado", totalFacturado.toFixed(2)],
  ["Gasto por productos", gastoProductos.toFixed(2)],
  ["IVA sobre productos (21%)", ivaSobreProductos.toFixed(2)],
  ["Gastos seguridad", gastosSeguridad.toFixed(2)],
  ["Gastos música", gastosMusica.toFixed(2)],
  ["Gasto total real estimado", gastoTotal.toFixed(2)],
  ["", ""],
  ["Margen bruto", margenBruto.toFixed(2)]
];

  body.appendTable(tablaResumen);

  body.appendParagraph(""); // Espacio visual
  body.appendParagraph("Detalle por producto:").setHeading(DocumentApp.ParagraphHeading.HEADING2);

  const tablaProductos = [["Producto", "Cantidad", "Costo Unitario (€)", "Total Gasto (€)"]];
for (const producto in resumen) {
  const r = resumen[producto];
  tablaProductos.push([
    producto,
    r.cantidad,
    r.costoUnitario.toFixed(2),
    r.totalGasto.toFixed(2)
  ]);
}

// Añadir fila de total
tablaProductos.push([
  "TOTAL",
  "",
  "",
  gastoProductos.toFixed(2)
]);

body.appendTable(tablaProductos);
// === GUARDAR EN HISTÓRICO GANANCIAS ===
const hojaGanancias = ss.getSheetByName("HISTÓRICO GANANCIAS") || ss.insertSheet("HISTÓRICO GANANCIAS");

if (hojaGanancias.getLastRow() === 0) {
  hojaGanancias.appendRow([
    "Fecha", "Día", "Semana", "Mes", "Año",
    "Total Facturado (€)", "Gasto Productos (€)", "Gasto Seguridad (€)",
    "Gasto Música (€)", "Gasto Misc (€)", "Margen Bruto (€)"
  ]);
}

const diaNombre = Utilities.formatDate(fechaFactDate, "Europe/Madrid", "EEEE");
const semana = Math.ceil(fechaFactDate.getDate() / 7);
const mes = Utilities.formatDate(fechaFactDate, "Europe/Madrid", "MMMM").toUpperCase();
const anio = fechaFactDate.getFullYear();

hojaGanancias.appendRow([
  fechaFacturacion,
  diaNombre,
  semana,
  mes,
  anio,
  totalFacturado.toFixed(2),
  gastoProductos.toFixed(2),
  gastosSeguridad.toFixed(2),
  gastosMusica.toFixed(2),
  0, // Gasto misceláneo a futuro
  margenBruto.toFixed(2)
]);


  doc.saveAndClose();
  const pdf = DriveApp.getFileById(doc.getId()).getAs("application/pdf");
  const folderId = "1Aw5oMvUz0r6cnvtdBGm0G8rpEoS4Dh2O"; // Reemplaza por tu carpeta real
  const folder = DriveApp.getFolderById(folderId);
  folder.createFile(pdf).setName(`Informe_Financiero_${fechaFacturacion.replace(/\//g, "-")}.pdf`);
  DriveApp.getFileById(doc.getId()).setTrashed(true);
}
