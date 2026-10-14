# Aplicar un bloqueo dinámico

El bloqueo permite capturar en una celda solo cuando otra celda o rango cumple una condición. Por ejemplo: habilitar “Especifique” cuando la respuesta sea `X`.

## Pasos

1. Selecciona las celdas que quieres habilitar o bloquear y pulsa **Capturar Rango**.
2. Marca **Bloqueo** y pulsa **Aplicar**.
3. Cuando se solicite, selecciona la celda o el rango que controla la condición.
   - Una celda se evalúa en cada fila.
   - Un rango puede evaluarse buscando el valor en todo el rango; si tiene el mismo número de filas que el destino, también puedes elegir evaluación fila por fila.
4. Escribe el operador que habilita la captura (`=`, `<>`, `>`, `<`, `>=` o `<=`).
5. Escribe el valor que completa la condición, por ejemplo `Sí`, `X` o `6`.
6. Indica si una condición vacía deja la celda habilitada o la bloquea.
7. Decide si quieres resaltar una celda vacía cuando ya quedó habilitada.
8. Si SAVCNG detecta formatos condicionales previos, elige conservarlos o reemplazarlos.
9. Comprueba ambas situaciones: una fila que cumple la condición y otra que no.

## Cómo se ve

Las celdas bloqueadas se muestran con trama gris. Si activaste el aviso de obligatoriedad, una celda habilitada pero vacía se resalta en azul. Excel rechaza el intento de captura cuando la condición no permite escribir.

## Importante

La opción de celda vacía y la evaluación fila por fila cambian el alcance de la regla. Antes de aplicarla a muchas filas, confirma que el rango de condición corresponde a las filas del destino.
