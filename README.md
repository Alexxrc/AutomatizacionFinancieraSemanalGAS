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
