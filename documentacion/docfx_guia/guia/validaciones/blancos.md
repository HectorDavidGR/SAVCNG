# Detectar campos vacíos obligatorios

Esta regla avisa cuando una fila ya contiene alguna respuesta pero dejó vacías otras celdas del rango. **No impide escribir ni bloquea la captura**; resalta lo pendiente en azul y presenta un aviso.

## Pasos

1. Selecciona la matriz de respuestas que quieras revisar, incluyendo las columnas que deben quedar completas por fila, y pulsa **Capturar Rango**.
2. Marca **Blancos** y pulsa **Aplicar**.
3. Indica valores de excepción que no deben activar la regla, separados por coma o punto y coma. El cuadro propone `2, 9` inicialmente; bórralos si no quieres excepciones o reemplázalos por tus valores.
4. Selecciona el rango donde aparecerá el mensaje. SAVCNG lo combinará y aplicará formato azul.
5. Si hay formatos condicionales previos, decide si conservarlos y apilar esta revisión o eliminarlos.
6. Prueba una fila vacía y otra parcialmente contestada; el aviso debe aparecer solo para la fila con respuestas y campos pendientes.

## Cómo interpreta las excepciones

La regla se omite cuando la fila contiene alguno de los valores indicados como excepción. Revisa la lista antes de confirmar, porque puede dejar fuera de revisión esas filas.

El mensaje se construye automáticamente; reserva el rango de destino para este aviso y no escribas allí información que deba conservarse.
