# Validar números enteros o decimales

Esta regla controla qué tipo de número se puede capturar en las celdas seleccionadas.

## Cuándo usarla

- **Enteros:** conteos o claves numéricas sin fracción.
- **Decimales:** mediciones que admiten hasta 10 posiciones decimales.

En ambos casos acepta valores no negativos y también los códigos `NS` y `NA`. No acepta números negativos.

## Pasos

1. Selecciona las celdas de respuesta y pulsa **Capturar Rango**.
2. Marca **Decimales** y pulsa **Aplicar**.
3. En la ventana de configuración, escribe `1` para enteros o `2` para decimales.
4. Si aparecen reglas anteriores, elige **Sí** para conservar formatos y bloqueos, **No** para borrarlos, o **Cancelar** para salir sin aplicar la regla.
5. Prueba un valor permitido y otro no permitido. Los datos inválidos se rechazan y pueden resaltarse en rojo si se modifican por otra vía.

## Qué considera inválido

- Enteros: texto distinto de `NS` o `NA`, fracciones y valores negativos.
- Decimales: texto distinto de `NS` o `NA`, valores negativos y más de 10 posiciones decimales.
- Las celdas vacías se permiten.

## Nota sobre reglas existentes

La validación de captura previa se reemplaza porque Excel solo admite una regla de validación por celda. La pregunta sobre conservar reglas mantiene o elimina los **formatos condicionales** previos; no conserva una validación de captura anterior.
