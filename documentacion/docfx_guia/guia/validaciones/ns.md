# Revisar el código NS

Esta opción busca el texto `NS` en las celdas capturadas y las resalta en color amarillo/dorado. También muestra un aviso cuando encuentra ese código. **No restringe los demás valores**: es una alerta de revisión, no una validación que limite la captura.

## Pasos

1. Selecciona las celdas que deben revisarse y pulsa **Capturar Rango**.
2. Marca **NS** y pulsa **Aplicar**.
3. Escribe el texto del aviso, por ejemplo: `Alerta: justificar el uso de NS.`
4. Si SAVCNG detecta reglas previas, elige si las conserva y apila esta alerta o si las elimina.
5. Para cada pregunta detectada, selecciona dónde aparecerá el aviso. Si capturaste varios bloques de la misma pregunta, podrás elegir una sola celda de aviso para todos o una celda por bloque.
6. Prueba escribiendo `NS` en una celda y confirma el resaltado y el mensaje.

## Recomendaciones

- Reserva espacio visible para el aviso. SAVCNG puede combinar las celdas seleccionadas como destino.
- El resaltado reconoce `NS` ignorando espacios exteriores y mayúsculas/minúsculas. Para que el mensaje lo detecte, captura exactamente `NS`.
- Las celdas capturadas pueden contener otros valores; esta función únicamente resalta y comunica la presencia de `NS`.
