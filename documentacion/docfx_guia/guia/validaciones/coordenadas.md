# Validar coordenadas geográficas

La opción **Coordenadas** aplica una regla de latitud o longitud a las celdas capturadas.

## Pasos

1. Selecciona las celdas de una sola variable (latitud o longitud) y pulsa **Capturar Rango**.
2. Marca **Coordenadas** y pulsa **Aplicar**.
3. Escribe `1` para latitud o `2` para longitud.
4. Si el rango ya tiene formatos condicionales, elige `1` para eliminarlos o `2` para conservarlos y apilar las alertas geográficas.
5. Prueba una coordenada válida y una fuera de los límites.

## Límites

| Tipo | Rango aceptado | Máximo de dígitos |
|---|---:|---:|
| Latitud | 11 a 33 | 8, sin contar el punto decimal |
| Longitud | -123 a -83 | 9, sin contar el punto decimal ni el signo negativo |

Las celdas vacías se permiten. Valores fuera del rango, no numéricos o con demasiados dígitos se rechazan al capturarlos y se marcan visualmente en rojo si la regla puede detectarlos.
