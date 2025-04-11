# Sistema de Automatización Financiera con Google Apps Script

Este proyecto implementa un sistema completo de automatización financiera semanal para un negocio que opera ciertos días de la semana (por ejemplo, viernes y sábado). Utilizando Google Apps Script, Google Sheets y Google Drive, automatiza el registro de ventas y facturación diaria, calcula márgenes financieros, gastos e impuestos, y genera informes consolidados en formato PDF que se almacenan automáticamente en una carpeta designada de Google Drive.

El objetivo principal es reducir la carga administrativa y mejorar la precisión del seguimiento financiero, todo sin necesidad de herramientas externas ni software de pago.

---

## Funcionalidades clave

- Registro automático de ventas por producto
- Registro de facturación diaria diferenciando efectivo y tarjeta
- Cálculo de:
  - Gastos por productos
  - IVA sobre productos (21%)
  - Gastos fijos diarios
  - Margen bruto y neto
  - Comparativa con EBITA registrado
- Generación de informes diarios y semanales en PDF
- Almacenamiento automático en Google Drive

---

## Estructura del repositorio

Este repositorio contiene:

- Scripts de Google Apps Script (`.gs`)
- Documentación de cada función
- Instrucciones de integración
- (Opcional) Plantillas de hojas de cálculo si se desea replicar el entorno

---

## Funciones disponibles
A continuación se listan todas las funciones del sistema con su descripción, parámetros de entrada y comportamiento esperado.

## Función: generarInformeViernes

Esta función genera un informe detallado de ventas para el día viernes. Utiliza los datos de pedidos, stock de barra y almacén, y calcula la cantidad de productos vendidos. El resultado se estructura en una hoja de cálculo temporal, se convierte en PDF, y se guarda automáticamente en una carpeta de Google Drive. Además, el informe se envía por correo y las ventas se registran en una base de datos histórica.

### Fuentes de datos
- `RAW PEDIDOS`: contiene el pedido más reciente de productos.
- `RAW BARRA`: contiene el stock de productos en barra.
- `RAW ALMACÉN`: contiene el stock de productos en almacén.

### Proceso
1. Se detecta la última fila de `RAW PEDIDOS` como el pedido actual.
2. Se calcula el stock inicial y final de barra y almacén, para el viernes.
3. Se genera una nueva hoja con el informe de ventas del viernes, incluyendo:
   - Producto
   - Stock inicial
   - Pedido
   - Stock final (barra y almacén)
   - Unidades vendidas
4. El informe se exporta como PDF.
5. El archivo se guarda en una carpeta de Drive.
6. Se registra la venta en una hoja histórica mediante `registrarVentasEnHistorico`.
7. Se elimina la hoja temporal de informe.

### Funciones auxiliares
- `obtenerFilasPorTipoDia(tipo)`: determina las filas de stock según el día (viernes o sábado).
- `guardarInformeEnCarpeta(blob, folderLink)`: guarda un archivo en una carpeta de Google Drive.

## Función: generarInformeSabado

Esta función automatiza la generación de un informe de ventas para el sábado. Utiliza los datos de stock registrados en `RAW BARRA` y `RAW ALMACÉN` para calcular el total de unidades vendidas por producto. El informe se presenta en una hoja temporal, exportada como PDF, guardada en Google Drive y enviada por correo. Las ventas se registran en una base de datos histórica.

### Fuentes de datos
- `RAW BARRA`: stock en barra antes y después del servicio.
- `RAW ALMACÉN`: stock en almacén antes y después del servicio.
- `RAW PEDIDOS`: contiene los encabezados con la lista de productos.

### Proceso
1. Obtiene las filas correspondientes al stock del sábado.
2. Calcula la diferencia entre stock inicial y final.
3. Genera una hoja con los datos consolidados (producto, stock, vendidos).
4. Exporta la hoja como PDF.
5. Guarda el archivo en una carpeta de Drive.
6. Envía el informe por correo electrónico.
7. Registra las ventas con la función `registrarVentasEnHistorico`.
8. Elimina la hoja temporal del informe.

### Funciones auxiliares utilizadas
- `obtenerFilasPorTipoDia(tipo)`: identifica las filas de stock según el día.
- `guardarInformeEnCarpeta(blob, folderLink)`: guarda el informe generado en una carpeta de Drive.

## Función: generarAnalisisSemanal

Esta función genera un informe de análisis de inventario y ventas semanales. Calcula los productos vendidos durante la semana utilizando datos de stock inicial, pedidos realizados y stock final. El informe se genera en una hoja temporal, se convierte en PDF y se envía por correo electrónico.

### Fuentes de datos
- `RAW PEDIDOS`: contiene los pedidos semanales realizados.
- `RAW BARRA` y `RAW ALMACÉN`: contienen los niveles de stock al inicio y al final de la semana.

### Proceso
1. Obtiene los datos de las dos últimas filas de stock para calcular stock inicial y final.
2. Calcula por producto:
   - Stock inicial (barra + almacén)
   - Pedido realizado
   - Stock final (barra + almacén)
   - Unidades vendidas estimadas
3. Genera una hoja con el resumen semanal.
4. Exporta la hoja como PDF.
5. Envía el informe por correo electrónico.
6. Elimina la hoja temporal del análisis.

### Formato del informe
El título incluye la semana, el mes y el año en que se genera el análisis. Los datos se presentan en una tabla con totales por producto.

## Función: registrarVentasEnHistorico

Registra automáticamente en una hoja de cálculo las unidades vendidas por producto en una fecha determinada. La función ajusta la fecha real de venta restando un día a la fecha proporcionada, ya que los registros se hacen normalmente al día siguiente de la jornada operativa.

### Proceso
1. Si no existe la hoja `HISTÓRICO VENTAS`, se crea con su encabezado.
2. A la fecha proporcionada se le resta un día para reflejar la venta real.
3. Se calcula:
   - Día de la semana (nombre)
   - Año
   - Semana del mes
   - Mes (en texto)
4. Se formatean los datos y se añaden al final de la hoja.

### Entrada
- `datos`: array de objetos con las propiedades `{ producto, vendidos }`
- `fecha`: fecha de referencia (día posterior al real)

### Salida
Inserta una fila por producto vendido en la hoja `HISTÓRICO VENTAS`.
## Función: generarInformeFinancieroDesdeFacturacion

Genera un informe financiero en PDF basado en los datos de facturación y ventas de un único día. Compara las unidades vendidas con el coste por producto, calcula los márgenes, gastos generales e IVA. El informe se guarda automáticamente en Google Drive como archivo PDF.

### Fuentes de datos
- `HISTÓRICO FACTURACIÓN`: contiene la facturación diaria (efectivo, tarjeta, gastos, EBITA).
- `HISTÓRICO VENTAS`: contiene los productos vendidos por día.
- `BASE DATOS PRECIOS`: contiene el coste unitario de cada producto.

### Proceso
1. Se toma la última fila de facturación como la referencia.
2. Se filtran las ventas correspondientes a esa misma fecha.
3. Se cruzan los productos vendidos con su coste unitario.
4. Se calcula:
   - Gasto total en productos
   - IVA aplicado (21%)
   - Gastos generales (seguridad + música)
   - Margen bruto y margen neto
   - Producto más vendido
5. Se genera un documento en Google Docs con tablas detalladas y resumen financiero.
6. Se convierte el documento en PDF y se guarda en Google Drive.
7. Se elimina el documento original tras exportar el PDF.

### Salida
PDF con informe financiero completo y datos por producto.
