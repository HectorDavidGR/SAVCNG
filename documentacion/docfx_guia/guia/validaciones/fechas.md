# Validar una fecha (día, mes y año)

La opción **Fechas** valida una fecha dividida en tres campos continuos y comprueba que el día exista para el mes y año capturados.

## Preparar las celdas

Selecciona conjuntamente tres campos en este orden: **Día, Mes, Año**. Pueden ser celdas combinadas, pero deben ser contiguas. No captures solo el campo del año: la regla necesita los tres componentes.

## Pasos

1. Selecciona los campos de día, mes y año correspondientes a las filas que quieres validar y pulsa **Capturar Rango**.
2. Marca **Fechas** y pulsa **Aplicar**.
3. Escribe el año mínimo y después el año máximo permitidos.
4. Selecciona una celda en una columna auxiliar libre (AF en adelante) para iniciar las fórmulas de revisión; deja espacio libre hacia abajo para las filas del rango.
5. Si SAVCNG encuentra reglas previas, decide si conservas los formatos y bloqueos. La regla de captura de Excel se reemplaza.
6. Cuando termine, revisa los campos resaltados en rojo. Si lo deseas, responde **Sí** a la pregunta para agregar un mensaje y selecciona el rango auxiliar y el destino del texto.

## Qué comprueba

- Mes entre 1 y 12.
- Día válido para el mes y el año; por ejemplo, febrero admite 29 días solo cuando el calendario lo permite.
- Año dentro de los límites indicados.
- Permite `NS` y `NA` en los campos, y permite celdas vacías.

Las fórmulas auxiliares producen una marca por fila (1 = hay error; 0 = sin error). Mantén esa zona auxiliar y el destino del mensaje sin datos que deban conservarse: SAVCNG escribe fórmulas allí.
