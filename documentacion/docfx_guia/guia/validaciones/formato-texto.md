# Validar el formato de texto

Esta regla restringe las respuestas de texto para mantenerlas uniformes. Acepta letras mayúsculas, números y espacios simples; también acepta celdas vacías.

## Antes de aplicarla

Reserva una columna auxiliar vacía, preferentemente desde **AF** en adelante, con espacio libre hacia abajo para las filas seleccionadas. SAVCNG escribe allí fórmulas de apoyo. Después podrás ocultar esa columna, pero no borres sus fórmulas mientras la regla siga en uso.

## Pasos

1. Selecciona las celdas de texto y pulsa **Capturar Rango**.
2. Marca **Formato Texto** y pulsa **Aplicar**.
3. Cuando Excel solicite la ubicación auxiliar, selecciona una celda dentro de la columna libre donde iniciará el motor.
4. Si hay reglas previas, decide si conservar los formatos condicionales o limpiarlos.
5. Prueba una respuesta en mayúsculas y otra con minúsculas o puntuación para confirmar el resultado.

## Reglas aplicadas

- El texto debe estar en mayúsculas. Se admiten letras de la A a la Z y la Ñ.
- Se admiten números y espacios sencillos entre palabras.
- No se permiten espacios dobles ni espacios al inicio o al final.
- No se permiten signos de puntuación, comillas, paréntesis ni otros caracteres especiales.

La regla bloquea el ingreso inválido y muestra un mensaje con los requisitos. El motor auxiliar usa una fila por cada fila del rango capturado.
