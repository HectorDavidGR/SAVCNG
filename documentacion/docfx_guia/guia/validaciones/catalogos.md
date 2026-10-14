# Aplicar un catálogo desplegable

Un catálogo agrega una lista desplegable a las celdas seleccionadas. Sirve para estandarizar respuestas, por ejemplo, permitir únicamente `1`, `2` o `9`.

## Antes de empezar

- Ten abierto el libro de censo correcto.
- Identifica las celdas de respuesta que recibirán la lista.
- Decide si escribirás las opciones directamente o si usarás una lista que ya está en la hoja.

## Pasos

1. En Excel, selecciona las celdas de respuesta.
2. En la ventana de SAVCNG, pulsa **Capturar Rango** y comprueba la dirección mostrada.
3. Marca **Catálogos**.
4. Pulsa **Aplicar**.
5. Cuando SAVCNG pregunte el origen del catálogo, elige **Sí** para escribir opciones manualmente o **No** para seleccionar celdas de Excel.
6. Si elegiste la entrada manual, escribe los valores separados por comas, por ejemplo `1,2,9`, y confirma.
7. Si elegiste celdas, selecciona la celda o el rango que contiene las opciones y confirma. El programa solicitará autorización para extraer únicamente los números encontrados en ese origen.
8. Si Excel avisa que ya existe una validación en las celdas destino, confirma el reemplazo solo si quieres sustituirla.
9. Haz clic en una celda destino y comprueba que aparezca la flecha desplegable con las opciones esperadas.

## Ejemplo: catálogo manual

Para permitir las claves `1`, `2` y `9`, captura primero las celdas de respuesta y escribe `1,2,9` cuando SAVCNG lo solicite.

## Si el catálogo está en la hoja

Selecciona el rango que contiene las opciones. En este modo, la implementación actual extrae **solo los dígitos** de cada celda seleccionada. Por ejemplo, una celda con `1. Sí` aporta `1`, no la etiqueta `Sí`. Si necesitas conservar textos como `Sí` o `No`, usa la entrada manual cuando sea apropiado y verifica el resultado en la lista desplegable.

## Qué ocurre al aplicarlo

Excel muestra una lista dentro de cada celda destino y restringe la entrada a las opciones configuradas. La validación ignora las celdas vacías. Si las celdas destino ya tenían otra regla de validación, SAVCNG pide confirmación antes de reemplazarla.

![Flujo para aplicar un catálogo](../imagenes/flujo-catalogo.svg)

<p align="center"><strong>Imagen 2.</strong> Flujo general para aplicar un catálogo desplegable en SAVCNG.</p>

## Problemas frecuentes

- **No aparece la flecha:** confirma que capturaste el rango de respuesta y que la aplicación terminó correctamente.
- **Faltan opciones de texto al usar un rango de origen:** el modo de selección de celdas extrae números; vuelve a aplicar el catálogo con la entrada manual si corresponde.
- **Se canceló el proceso:** vuelve a seleccionar el rango destino, captúralo y comienza de nuevo.

## De dónde se obtuvo esta guía

Los pasos se contrastaron con `Validaciones/ValidacionCatalogos.cs` y el flujo de aplicación de `FrmValidaciones.cs`. La selección de rango es el comportamiento implementado actualmente; las capturas de pantalla reales se pueden añadir cuando se registren los pasos en Excel.
