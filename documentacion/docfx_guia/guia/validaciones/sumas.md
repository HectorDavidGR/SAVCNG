# Comprobar sumas por fila

La validación compara, en cada fila, una columna de **Total** con la suma de las columnas de **desagregados**. También muestra un aviso con la cantidad de inconsistencias y el primer numeral donde aparece alguna.

## Preparar la tabla

Selecciona las filas de datos con la opción **Sumas**; excluye encabezados y la fila destinada a sumas verticales. Los rangos de total y desagregados deben corresponder a esas mismas filas. Reserva una columna auxiliar libre desde AF en adelante y espacio libre hacia abajo para las fórmulas.

## Pasos

1. Selecciona las filas de la tabla que se van a revisar y pulsa **Capturar Rango**.
2. Marca **Sumas** y pulsa **Aplicar**.

![Rango de la tabla capturado y opción Sumas seleccionada](../imagenes/capturas/validacion-sumas/Sumas%200.png)

<p align="center"><i><strong>Imagen 3.</strong> Rango capturado y opción Sumas seleccionada antes de aplicar la validación.</i></p>

3. Selecciona la columna **Total** (o pivote) correspondiente a las filas capturadas.

![Selección de la columna Total para la validación](../imagenes/capturas/validacion-sumas/Sumas%201.png)

<p align="center"><strong>Imagen 4.</strong> Selección de la columna Total para las filas incluidas en la captura.</p>

4. Selecciona las columnas de desagregados. Si no son contiguas, usa **Ctrl** para seleccionar bloques separados.

![Selección de columnas desagregadas en Excel](../imagenes/capturas/validacion-sumas/Sumas%202.png)

<p align="center"><strong>Imagen 5.</strong> Selección de las columnas desagregadas; en el ejemplo, Hombres y Mujeres.</p>

5. Selecciona el rango donde aparecerá el aviso de inconsistencias.

![Selección del destino para el aviso de error](../imagenes/capturas/validacion-sumas/Sumas%203.png)

<p align="center"><strong>Imagen 6.</strong> Selección del destino donde aparecerá el aviso de inconsistencias.</p>

6. Selecciona una celda en una columna libre (por ejemplo AF) para las fórmulas auxiliares. Deja espacio para todas las filas capturadas.

![Selección de la celda auxiliar](../imagenes/capturas/validacion-sumas/Sumas%204.png)

<p align="center"><strong>Imagen 7.</strong> Selección de una celda auxiliar en una columna libre.</p>

7. Cuando pregunte por la sumatoria vertical (Σ), elige **Sí** para seleccionar la fila de totales o **No** para omitirla.

![Pregunta para agregar la sumatoria vertical](../imagenes/capturas/validacion-sumas/Sumas%205.png)

<p align="center"><strong>Imagen 8.</strong> Confirmación opcional para agregar la sumatoria vertical.</p>

8. Si elegiste **Sí**, selecciona las celdas de la fila donde quieres las sumas, para las columnas de Total y desagregados.

![Selección de la fila de sumatoria vertical](../imagenes/capturas/validacion-sumas/Sumas%206.png)

<p align="center"><strong>Imagen 9.</strong> Selección de las celdas de la fila Σ para el Total y los desagregados.</p>

9. Revisa los errores reportados y corrige los valores de la tabla.

![Confirmación de que la validación de sumas se aplicó](../imagenes/capturas/validacion-sumas/Sumas%207.png)

<p align="center"><strong>Imagen 10.</strong> Confirmación de la validación de sumas y, si se eligió, de la sumatoria vertical.</p>

## Tratamiento de NS y NA

- Si el Total es numérico, las sumas se comparan; los `NS` en los desagregados se omiten bajo las condiciones previstas por la regla.
- `NA` debe aparecer de manera consistente en el total y todos los desagregados; una mezcla parcial genera inconsistencia.
- `NS` en el Total también se revisa según los datos de los desagregados, por lo que no debe asumirse que toda fila con `NS` queda ignorada.

## Alerta y sumatoria vertical

El mensaje indica cuántos errores hay y el primer numeral a revisar. Si habilitas Σ, SAVCNG escribe fórmulas para sumar cada columna y conserva códigos `NS`, `NA` o cero bajo las condiciones de la fórmula. Las celdas de alerta y la columna auxiliar se modifican: usa zonas vacías.
