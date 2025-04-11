# Sistema de Automatización Financiera con Google Apps Script

Este proyecto implementa un sistema completo de automatización financiera semanal para un negocio que opera ciertos días de la semana (por ejemplo, viernes y sábado). Utilizando Google Apps Script, Google Sheets y Google Drive, automatiza el registro de ventas y facturación diaria, calcula márgenes financieros, gastos e impuestos, y genera informes consolidados en formato PDF que se almacenan automáticamente en una carpeta de Google Drive.

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
- Documentación detallada de cada función
- (Opcional) Plantillas de hojas de cálculo para replicar el entorno

---

## Funciones disponibles

### `generarInformeViernes`

Genera un informe detallado de ventas del viernes utilizando los datos de pedidos, stock de barra y almacén. Calcula las unidades vendidas por producto, genera una hoja temporal y la convierte en PDF.

**Fuentes de datos:**
- `RAW PEDIDOS`, `RAW BARRA`, `RAW ALMACÉN`

**Proceso:**
1. Detecta el último pedido registrado
2. Calcula stock inicial y final
3. Estima productos vendidos
4. Exporta la hoja como PDF
5. Guarda el archivo en Drive
6. Envía el informe por correo
7. Registra ventas en `HISTÓRICO VENTAS`
8. Elimina la hoja temporal

**Funciones auxiliares:**  
- `obtenerFilasPorTipoDia(tipo)`  
- `guardarInformeEnCarpeta(blob, folderLink)`

---

### `generarInformeSabado`

Genera el informe de ventas del sábado de forma análoga al viernes, usando el stock inicial y final para calcular las unidades vendidas.

**Fuentes de datos:**
- `RAW BARRA`, `RAW ALMACÉN`, `RAW PEDIDOS`

**Proceso:**
1. Obtiene filas correspondientes al sábado
2. Calcula ventas estimadas
3. Crea hoja resumen
4. Exporta como PDF y guarda en Drive
5. Envía por correo
6. Registra las ventas
7. Elimina hoja temporal

---

### `generarAnalisisSemanal`

Crea un análisis general del inventario y ventas de la semana. Compara el pedido semanal con los movimientos de stock para estimar unidades vendidas.

**Fuentes de datos:**
- `RAW PEDIDOS`, `RAW BARRA`, `RAW ALMACÉN`

**Proceso:**
1. Obtiene stock inicial y final de semana
2. Calcula unidades vendidas por producto
3. Genera hoja con resumen por producto
4. Exporta y envía por correo como PDF
5. Elimina hoja temporal

---

### `registrarVentasEnHistorico`

Registra en una hoja de cálculo las unidades vendidas por producto. Ajusta automáticamente la fecha para reflejar el día real de venta (un día antes del registro).

**Entradas:**
- `datos`: array de `{ producto, vendidos }`
- `fecha`: fecha de registro (día posterior al real)

**Proceso:**
1. Crea la hoja `HISTÓRICO VENTAS` si no existe
2. Calcula día, semana, mes y año
3. Añade fila por producto con información detallada

---

### `generarInformeFinancieroDesdeFacturacion` (Informe diario)

Genera un informe financiero en PDF basado en los datos de facturación y ventas de un único día. Calcula márgenes, IVA, gastos y el producto más vendido.

**Fuentes de datos:**
- `HISTÓRICO FACTURACIÓN`, `HISTÓRICO VENTAS`, `BASE DATOS PRECIOS`

**Proceso:**
1. Extrae la última fila de facturación
2. Filtra ventas del mismo día
3. Cruza con precios unitarios
4. Calcula:
   - Gasto en productos
   - IVA (21%)
   - Gastos generales
   - Margen neto y bruto
5. Genera documento PDF y lo guarda en Drive

---

### `generarInformeFinancieroDesdeFacturacion` (Informe semanal)

Esta variante genera un informe financiero **semanal** a partir de las dos últimas entradas (viernes y sábado) en `HISTÓRICO FACTURACIÓN`. 

**Fuentes de datos:**
- `HISTÓRICO FACTURACIÓN`, `HISTÓRICO VENTAS`, `BASE DATOS PRECIOS`

**Proceso:**
1. Toma las dos últimas fechas de facturación
2. Filtra ventas de esas fechas
3. Cruza con precios unitarios
4. Calcula:
   - Gasto total en productos
   - IVA (21%)
   - Gastos por día
   - Margen bruto y neto
5. Identifica el producto más vendido
6. Genera un documento PDF con el resumen semanal
7. Guarda el archivo en Drive

**Salida:**  
Informe PDF con:
- Detalle por producto
- Tabla de ingresos y gastos por día
- Márgenes y totales consolidados

---
