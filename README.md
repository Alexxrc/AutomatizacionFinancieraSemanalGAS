# AutomatizacionFinancieraSemanalGAS
# Automatización Financiera Semanal

Este proyecto automatiza la recopilación y análisis de datos de facturación y ventas de un negocio que opera dos días a la semana. El sistema genera un informe financiero semanal en PDF utilizando datos extraídos de varias hojas de cálculo de Google Sheets.

---

## Funcionalidades

- Registro automatizado de facturación y ventas por día.
- Comparación de ventas por producto con su coste unitario.
- Cálculo de márgenes, gastos, IVA y EBITA.
- Generación de un informe financiero semanal en PDF.
- Almacenamiento automático del informe en Google Drive.

---

## Estructura de hojas de cálculo

### `HISTÓRICO FACTURACIÓN`
Contiene:
- Fecha de facturación
- Total facturado (efectivo, tarjeta)
- Gastos por día (seguridad, música)
- EBITA diario

### `HISTÓRICO VENTAS`
Contiene:
- Una fila por producto vendido y fecha
- Producto, cantidad, fecha

### `BASE DATOS PRECIOS`
Contiene:
- Nombre del producto (en mayúsculas)
- Coste unitario

---

## Contenido del informe semanal

- Fechas analizadas (viernes y sábado)
- Producto más vendido
- Tabla por producto con:
  - Cantidad total vendida
  - Costo unitario
  - Gasto total en productos

### Resumen financiero

| Concepto                     | Descripción                                      |
|-----------------------------|--------------------------------------------------|
| Facturado por día           | Totales individuales por jornada                |
| Gastos por día              | Seguridad y música, detallados por jornada      |
| Gasto en productos          | Suma total del coste de productos vendidos      |
| IVA aplicado (21%)          | Sobre el total de productos                     |
| Gasto total estimado        | Productos + IVA + otros gastos                  |
| Margen bruto                | Facturado - coste de productos                  |
| Margen neto                 | Facturado - gastos totales                      |
| EBITA registrado            | Según datos del registro de facturación        |

---

## Cómo funciona

El script principal:
1. Toma las dos últimas fechas registradas en `HISTÓRICO FACTURACIÓN`
2. Filtra las ventas correspondientes en `HISTÓRICO VENTAS`
3. Cruza los productos con sus precios desde `BASE DATOS PRECIOS`
4. Calcula los totales y genera el informe
5. Exporta un archivo PDF y lo guarda automáticamente en Google Drive

---

## Requisitos

- Proyecto de Google Apps Script vinculado a un documento de Google Sheets.
- Las hojas de cálculo mencionadas correctamente configuradas.
- Acceso a Google Drive para almacenamiento de archivos.

---

## Mejoras futuras sugeridas

- Envío automático del informe PDF por correo electrónico.
- Generación de gráficos y dashboards con Google Charts.
- Análisis mensual y alertas automáticas.
- Registro histórico de informes semanales en una hoja separada.

---

## Licencia

Este proyecto está pensado para uso educativo y de automatización interna. Puede modificarse y adaptarse libremente para otros contextos.


