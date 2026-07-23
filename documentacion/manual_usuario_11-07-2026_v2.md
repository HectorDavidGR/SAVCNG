# SAVCNG - Manual de Usuario
## Guía Paso a Paso para Aplicar Validaciones en Excel

**Versión**: 2.4.0  
**Fecha**: 11 de julio de 2026  
**Dirigido a**: Usuarios finales (sin conocimientos técnicos)

---

## ¿Qué es SAVCNG?

**SAVCNG** es una herramienta que se instala en tu Excel y te ayuda a proteger los datos de los archivos de censo. Funciona como un asistente: te hace preguntas paso a paso y tú solo tienes que seleccionar celdas y responder con Sí o No.

No necesitas saber programación ni fórmulas. La herramienta hace todo el trabajo por ti.

---

## Tabla de Contenidos

1. [¿Cómo abrir SAVCNG?](#1-cómo-abrir-savcng)
2. [Paso previo obligatorio: Capturar el Rango](#2-paso-previo-obligatorio-capturar-el-rango)
3. [Validación 1 — Solo Números Enteros](#3-validación-1--solo-números-enteros)
4. [Validación 2 — Lista de Opciones (Catálogo)](#4-validación-2--lista-de-opciones-catálogo)
5. [Validación 3 — Campo NS (No Sabe / No Aplica)](#5-validación-3--campo-ns-no-sabe--no-aplica)
6. [Validación 4 — Formato de Texto](#6-validación-4--formato-de-texto)
7. [Validación 5 — Bloqueo Dinámico](#7-validación-5--bloqueo-dinámico)
8. [Validación 6 — Campos Vacíos Obligatorios](#8-validación-6--campos-vacíos-obligatorios)
9. [Validación 7 — Revisar el campo "Especifique"](#9-validación-7--revisar-el-campo-especifique)
10. [Validación 8 — Validación de Años](#10-validación-8--validación-de-años)
11. [Validación 9 — Verificación de Sumas](#11-validación-9--verificación-de-sumas)
12. [Preguntas Frecuentes](#12-preguntas-frecuentes)

---

## 1. ¿Cómo abrir SAVCNG?

Cuando abres Excel, verás una pestaña nueva en la parte superior llamada **"SAVCNG - Censo"**.

**Pasos:**

1. Abre Excel normalmente.
2. Busca la pestaña **"SAVCNG - Censo"** en la cinta de opciones (arriba).
3. Haz clic en el botón **"Cargar Censo"**.
4. Se abrirá una ventana para buscar tu archivo. Navega hasta él y haz clic en **Abrir**.
5. Aparecerá una ventana flotante (la ventana de SAVCNG) que dirá el nombre del archivo cargado.

> 💡 **Consejo**: La ventana de SAVCNG siempre estará encima de Excel para que puedas usarla mientras trabajas en la hoja.

---

## 2. Paso Previo Obligatorio: Capturar el Rango

> ⚠️ **Este paso es obligatorio antes de aplicar CUALQUIER validación.**

Antes de aplicar cualquier validación, debes decirle a SAVCNG en qué celdas vas a trabajar. A esto se le llama **"capturar el rango"**.

**Pasos:**

1. En tu archivo de Excel, **selecciona con el mouse las celdas** donde quieres aplicar la validación.
   - Puede ser una sola celda, una columna, varias columnas o filas.
   - Para seleccionar varias áreas separadas, mantén presionada la tecla **Ctrl** mientras haces clic.

2. Sin deseleccionar las celdas, ve a la ventana de SAVCNG y haz clic en el botón **"Capturar Rango"**.

3. Verás un mensaje que dice **"Se capturó correctamente el rango"** y la ventana mostrará qué celdas quedaron guardadas.

> ✅ **Listo.** Ahora puedes aplicar cualquiera de las validaciones de abajo.

> 🔁 **Recuerda**: Cada vez que quieras validar un grupo diferente de celdas, debes repetir este paso.

---

## 3. Validación 1 — Solo Números Enteros

**¿Para qué sirve?**  
Impide que alguien escriba números con decimales (como 1.5 o 3.7) o letras en las celdas. Solo acepta números enteros como 1, 15, 100.

**Pasos:**

1. Selecciona las celdas en Excel y haz clic en **"Capturar Rango"**.
2. Marca la casilla **"Decimales"**.
3. Haz clic en **"Aplicar"** y confirma el mensaje.

**¿Qué pasará cuando alguien intente escribir mal?**

> *"El formato de esta celda no admite números con decimales ni texto. Por favor, introduce únicamente un número entero (Ej: 1, 15, 100)."*

---

## 4. Validación 2 — Lista de Opciones (Catálogo)

**¿Para qué sirve?**  
Crea una lista desplegable en las celdas para que el informante solo pueda elegir opciones predefinidas, como "1, 2, 3" o "Sí, No, N/A".

**Pasos:**

1. Selecciona las celdas en Excel y haz clic en **"Capturar Rango"**.
2. Marca la casilla **"Catálogos"** y haz clic en **"Aplicar"**.
3. La herramienta te preguntará de dónde vienen las opciones (manual o desde celdas de Excel).

---

## 5. Validación 3 — Campo NS (No Sabe / No Aplica)

**¿Para qué sirve?**  
Permite que la celda acepte únicamente un número >= 0 o el texto **"NS"**, y muestra un mensaje de alerta cuando alguien escribe "NS".

**Pasos:**

1. Selecciona las celdas en Excel y haz clic en **"Capturar Rango"**.
2. Marca la casilla **"NS"** y haz clic en **"Aplicar"**.
3. Elige la celda donde aparecerá el mensaje de alerta.
4. Confirma o personaliza el texto del mensaje.

---

## 6. Validación 4 — Formato de Texto

**¿Para qué sirve?**  
Obliga a que el texto cumpla un formato estricto: solo MAYÚSCULAS, sin dobles espacios y sin caracteres especiales.

**Pasos:**

1. Selecciona las celdas en Excel y haz clic en **"Capturar Rango"**.
2. Marca la casilla **"Formato Texto"** y haz clic en **"Aplicar"**.
3. Selecciona una columna vacía lejana para alojar el motor de validación invisible.

> 💡 Puedes ocultar esa columna auxiliar sin afectar la validación (clic derecho en la letra de columna → Ocultar).

---

## 7. Validación 5 — Bloqueo Dinámico

**¿Para qué sirve?**  
Bloquea celdas de captura para que **no se pueda escribir en ellas** a menos que otra celda cumpla una condición específica.

**Pasos:**

1. Selecciona las celdas que quieres bloquear/desbloquear y haz clic en **"Capturar Rango"**.
2. Marca la casilla **"Bloqueo"** y haz clic en **"Aplicar"**.
3. Sigue los 4 pasos guiados: celda de condición → operador → valor → comportamiento con celda vacía → resaltado azul opcional.

---

## 8. Validación 6 — Campos Vacíos Obligatorios

**¿Para qué sirve?**  
Detecta y resalta en azul las celdas que están vacías dentro de una fila que ya tiene datos en otras columnas.

> ℹ️ Esta validación **no bloquea** la captura, solo resalta visualmente lo que falta.

**Pasos:**

1. Selecciona el rango de la matriz de captura y haz clic en **"Capturar Rango"**.
2. Marca la casilla **"Blancos"** y haz clic en **"Aplicar"**.
3. Elige la celda del mensaje de alerta y personaliza el texto.
4. Si ya hay otras validaciones activas, decide si conservarlas o reemplazarlas.

---

## 9. Validación 7 — Revisar el campo "Especifique"

**¿Para qué sirve?**  
Detecta si el texto libre escrito en un campo "Especifique" coincide con alguna opción del catálogo.

> ℹ️ Esta validación **no bloquea** la captura, solo alerta visualmente.

**Pasos:**

1. Selecciona la celda del "Especifique" y haz clic en **"Capturar Rango"**.
2. Marca la casilla **"Esp. Clave"** y haz clic en **"Aplicar"**.
3. Sigue los 3 pasos guiados: opciones del catálogo → celda del mensaje de alerta → celda vacía para el motor.

---

## 10. Validación 8 — Validación de Años

**¿Para qué sirve?**  
Restringe la celda para que solo acepte un año numérico dentro de un rango definido (por ejemplo, entre 1821 y 2026), o el código "NS".

**Pasos:**

1. Selecciona las celdas en Excel y haz clic en **"Capturar Rango"**.
2. Marca la casilla **"Años"** y haz clic en **"Aplicar"**.
3. Ingresa el límite inferior (año más antiguo) y el límite superior (año más reciente).

---

## 11. Validación 9 — Verificación de Sumas ⭐ (Nuevo en v2.4.0)

**¿Para qué sirve?**  
Verifica automáticamente que la suma de los valores de las columnas de **desagregados** coincida con la columna de **total** en cada fila de la tabla. Además, calcula e inyecta las fórmulas de **sumatoria vertical** al pie de la tabla.

**Ejemplo práctico:**

| Total | Hombres | Mujeres | Otro |
|-------|---------|---------|------|
| 100   | 30      | 40      | 20   | ← ❌ Se resalta en ROJO: 30+40+20=90 ≠ 100 |
| 90    | 30      | 40      | 20   | ← ✅ Correcto |
| **90** | **60** | **80** | **40** | ← Fórmulas de suma vertical inyectadas automáticamente |

**¿Cuándo usarla?**  
En cualquier tabla de captura donde exista una columna de total que debe ser igual a la suma de sus partes (desagregados por sexo, por categoría, por municipio, etc.).

> ℹ️ Esta validación **no bloquea** la captura. Resalta en rojo las inconsistencias para que el capturista las corrija.

---

**Pasos:**

### Paso previo — Capturar el Rango de la tabla

1. En Excel, selecciona **todas las filas de datos de la tabla** (sin incluir la fila de encabezado ni la fila de totales verticales).
   - Por ejemplo, si tu tabla de datos va de la fila 5 a la fila 54, selecciona desde la primera celda de datos hasta la última.
2. Haz clic en **"Capturar Rango"** en la ventana de SAVCNG.

---

### Paso 1 — ¿Cuál es la columna del Total?

Aparecerá una ventana con el mensaje:

> *"Selecciona la COLUMNA del TOTAL (debe coincidir con las filas del rango capturado)"*

1. En Excel, **haz clic en la columna** que contiene los totales de cada fila (por ejemplo, la columna "Total general").
   - Selecciona solo la columna o las celdas de esa columna que corresponden a las filas de datos.
2. Haz clic en **Aceptar**.

> ⚠️ **Importante:** La columna que selecciones debe tener exactamente las mismas filas que el rango que capturaste antes.

---

### Paso 2 — ¿Cuáles son las columnas de los Desagregados?

Aparecerá una ventana con el mensaje:

> *"Selecciona las COLUMNAS de los DESAGREGADOS (Puedes usar CTRL para seleccionar varias separadas)"*

1. Selecciona en Excel todas las columnas que **se suman para dar el total**.
   - Si las columnas están juntas: haz clic en la primera y arrastra hasta la última.
   - Si las columnas están separadas: mantén presionada la tecla **Ctrl** y haz clic en cada columna.
2. Haz clic en **Aceptar**.

---

### Paso 3 — ¿Dónde va la fila de sumatoria vertical?

Aparecerá una ventana con el mensaje:

> *"Selecciona la FILA o CELDAS destino para la Sumatoria Vertical (Σ) al final de la tabla"*

1. En Excel, selecciona **la celda o celdas donde quieres que aparezcan las sumas totales** al pie de la tabla.
   - Por ejemplo, si la fila de totales está en la fila 55, selecciona las celdas de esa fila que corresponden a la columna de Total y a las columnas de Desagregados.
2. Haz clic en **Aceptar**.

---

### Resultado

Después de los 3 pasos:

- **Las celdas donde la suma no coincide** quedarán resaltadas en **rojo** para indicar el error.
- **Las celdas correctas** no tendrán color.
- **Las celdas con "NS"** no se verificarán matemáticamente (se omiten automáticamente para no generar alertas incorrectas).
- **La fila de totales verticales** tendrá las fórmulas de suma inyectadas automáticamente.

Aparecerá el mensaje:

> *"Validación de consistencia horizontal (Sumas) y fórmulas verticales inyectadas con éxito."*

Haz clic en **Aceptar**.

---

### ¿Qué pasa si ya había reglas previas de sumas?

Si ya habías aplicado esta validación antes en las mismas celdas, la herramienta te preguntará:

> *"Se detectaron reglas de validación de sumas previas. ¿Deseas CONSERVARLAS e integrar esta nueva capa?"*

| Respuesta | Resultado |
|-----------|-----------|
| **SÍ** | Se agregan las nuevas reglas sin borrar las anteriores. Recomendado para tablas con subtotales y totales generales. |
| **NO** | Se borran las reglas anteriores y se aplica solo la nueva. |
| **Cancelar** | No se hace ningún cambio. |

---

### Caso especial: Tabla con subtotales

Si tu tabla tiene una estructura jerárquica (por ejemplo, subtotales por región y un total general), puedes aplicar la validación **dos veces** sobre el mismo rango:

1. **Primera aplicación**: selecciona las columnas de los subtotales como desagregados del total general → elige **SÍ** cuando te pregunte si conservar reglas.
2. **Segunda aplicación**: selecciona otro nivel de columnas como desagregados de un subtotal → elige **SÍ** de nuevo.

Así tendrás ambas capas de verificación activas al mismo tiempo.

---

## 12. Preguntas Frecuentes

---

**¿Puedo aplicar más de una validación a las mismas celdas?**

Sí. Por ejemplo, puedes aplicar primero el **Bloqueo Dinámico** y luego los **Campos Vacíos** sobre el mismo rango. Cuando apliques la segunda validación, la herramienta te preguntará si deseas conservar las reglas anteriores. Responde **SÍ** para no perder lo que ya aplicaste.

---

**¿Puedo apilar dos capas de Sumas en la misma tabla?**

Sí. Si tu tabla tiene subtotales intermedios y un total general, aplica la validación de Sumas dos veces sobre el mismo rango. En la segunda aplicación, elige **SÍ** cuando te pregunte si conservar las reglas previas. Ambas capas coexistirán.

---

**¿Qué hago si me equivoqué y quiero quitar una validación?**

Selecciona las celdas en Excel, ve a la pestaña **Datos** en la cinta de opciones, haz clic en **Validación de Datos** y luego en **Borrar todo**. Para el resaltado de colores: **Inicio → Formato condicional → Borrar reglas**.

---

**¿Por qué la herramienta pone cosas en columnas lejanas de mi hoja?**

Algunas validaciones (Formato Texto y Especifique por Palabras Clave) necesitan columnas auxiliares para funcionar. **No borres esas columnas**, ya que son parte de la validación. Puedes ocultarlas si no quieres verlas (clic derecho en la letra de la columna → Ocultar).

---

**¿Qué significa cuando una celda se pone de color azul?**

Significa que esa celda **está vacía y debería llenarse**. Lo aplican las validaciones de Blancos y Bloqueo Dinámico.

---

**¿Qué significa cuando una celda se pone de color rojo?**

Significa que la suma de los desagregados **no coincide con el total** en esa fila. Lo aplica la validación de Sumas. Revisa los valores de esa fila y corrige el error.

---

**¿Qué significa cuando una celda tiene un patrón gris cruzado?**

Significa que esa celda está **bloqueada**: no se puede escribir en ella porque la condición de desbloqueo no se ha cumplido. Lo aplica la validación de Bloqueo Dinámico.

---

**Una fila tiene "NS" en alguna columna pero no se resalta en rojo, ¿es correcto?**

Sí, es el comportamiento esperado. Cuando cualquier celda de la fila (total o desagregado) contiene "NS", la herramienta omite la verificación matemática de esa fila para evitar alertas incorrectas.

---

**¿Puedo cambiar el texto de los mensajes de alerta?**

Sí, en las validaciones **NS** y **Campos Vacíos** la herramienta te pide que escribas o confirmes el texto del mensaje antes de aplicarlo.

---

**¿Puedo usar SAVCNG en cualquier hoja del archivo?**

Sí. Solo necesitas capturar el rango en la hoja donde quieres aplicar la validación antes de hacer clic en "Aplicar".

---

**¿Qué pasa si cierro la ventana de SAVCNG?**

Las validaciones ya aplicadas **permanecen en el archivo de Excel**. Solo se borran si tú las eliminas manualmente. Puedes volver a abrir SAVCNG en cualquier momento desde la pestaña **"SAVCNG - Censo"**.

---

## Resumen Rápido de Validaciones

| # | Validación | Checkbox | ¿Para qué sirve? | ¿Bloquea la captura? |
|---|-----------|----------|-----------------|----------------------|
| 1 | Solo Enteros | Decimales | Impide decimales y texto | ✅ Sí |
| 2 | Lista de Opciones | Catálogos | Crea un menú desplegable | ✅ Sí |
| 3 | NS / No Aplica | NS | Permite números ≥0 o "NS" con alerta | ✅ Sí |
| 4 | Formato de Texto | Formato Texto | Obliga MAYÚSCULAS y sin caracteres especiales | ✅ Sí |
| 5 | Bloqueo Dinámico | Bloqueo | Bloquea celdas hasta que se cumpla una condición | ✅ Sí |
| 6 | Campos Vacíos | Blancos | Resalta en azul los campos vacíos obligatorios | ❌ Solo resalta |
| 7 | Especifique | Esp. Clave | Alerta si el texto libre coincide con el catálogo | ❌ Solo alerta |
| 8 | Años | Años / Fechas | Acepta solo años dentro de un rango definido | ✅ Sí |
| 9 | Sumas | Sumas | Resalta en rojo inconsistencias de suma + totales verticales | ❌ Solo resalta |

---

## Flujo de Trabajo Recomendado

```
1. Abrir Excel
   └─→ Pestaña "SAVCNG - Censo" → Cargar Censo

2. Seleccionar las celdas en Excel

3. Clic en "Capturar Rango" (SIEMPRE PRIMERO)

4. Marcar la casilla de la validación deseada

5. Clic en "Aplicar"

6. Seguir las instrucciones paso a paso que aparecen en pantalla

7. Confirmar el mensaje de éxito

8. Repetir desde el paso 2 para el siguiente grupo de celdas
```

---

**Manual de Usuario — SAVCNG ExcelDNA**  
**Versión**: 2.4.0 | **Fecha**: 11 de julio de 2026
