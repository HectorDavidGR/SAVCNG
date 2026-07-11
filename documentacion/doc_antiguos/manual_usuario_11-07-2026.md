# SAVCNG - Manual de Usuario
## Guía Paso a Paso para Aplicar Validaciones en Excel

**Versión**: 2.3.0  
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
11. [Preguntas Frecuentes](#11-preguntas-frecuentes)

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

**¿Cuándo usarla?**  
Cuando la celda solo debe recibir cantidades enteras, por ejemplo: número de personas, número de años, conteos.

---

**Pasos:**

1. Selecciona las celdas en Excel y haz clic en **"Capturar Rango"** (ver sección 2).

2. En la ventana de SAVCNG, marca la casilla **"Decimales"** (o "Solo Enteros").

3. Haz clic en el botón **"Aplicar"**.

4. Aparecerá un mensaje confirmando que la validación fue aplicada. Haz clic en **Aceptar**.

**¿Qué pasará cuando alguien intente escribir mal?**  
Si el informante escribe un decimal o texto, Excel mostrará automáticamente este mensaje:

> *"El formato de esta celda no admite números con decimales ni texto. Por favor, introduce únicamente un número entero (Ej: 1, 15, 100)."*

---

## 4. Validación 2 — Lista de Opciones (Catálogo)

**¿Para qué sirve?**  
Crea una lista desplegable en las celdas para que el informante solo pueda elegir opciones predefinidas, como "1, 2, 3" o "Sí, No, N/A".

**¿Cuándo usarla?**  
Cuando la celda solo debe aceptar ciertos valores específicos y quieres evitar errores de escritura.

---

**Pasos:**

1. Selecciona las celdas en Excel y haz clic en **"Capturar Rango"**.

2. Marca la casilla **"Catálogos"** y haz clic en **"Aplicar"**.

3. La herramienta te preguntará de dónde vienen las opciones:

---

### Opción A — Escribir las opciones manualmente

Elige **"SÍ"** cuando te pregunte si deseas escribirlas manualmente.

- Se abrirá un cuadro donde puedes escribir las opciones separadas por comas.
- Ejemplo: escribe `1,2,3,9` o `X,N/A`
- Haz clic en **Aceptar**.

---

### Opción B — Seleccionar las opciones desde una celda de Excel

Elige **"NO"** para seleccionar celdas que ya tienen las opciones escritas.

- Se abrirá una ventana pidiéndote que selecciones el rango de opciones en Excel.
- Selecciona las celdas con las opciones y haz clic en **Aceptar**.

Luego la herramienta te preguntará cómo extraer las opciones:

| Respuesta | ¿Qué hace? |
|-----------|-----------|
| **SÍ** | Extrae solo los números del texto (ignora letras) |
| **NO** | Toma el texto exactamente como está escrito |
| **CANCELAR** | Usa la celda completa como referencia directa |

4. Una vez elegida la opción, la herramienta aplica la lista desplegable. Haz clic en **Aceptar** en el mensaje de confirmación.

**¿Qué verá el informante?**  
Las celdas tendrán una pequeña flecha. Al hacer clic en ella, aparecerá la lista de opciones para elegir.

---

## 5. Validación 3 — Campo NS (No Sabe / No Aplica)

**¿Para qué sirve?**  
Permite que la celda acepte únicamente:
- Un **número igual o mayor a cero** (0, 1, 5, 100…), o
- El texto **"NS"** (No Sabe / No Aplica).

Además, cuando alguien escribe "NS", automáticamente aparece un mensaje de alerta que pide una justificación.

**¿Cuándo usarla?**  
En preguntas numéricas donde existe la posibilidad de que el informante no sepa o no aplique el dato.

---

**Pasos:**

1. Selecciona las celdas en Excel y haz clic en **"Capturar Rango"**.

2. Marca la casilla **"NS"** y haz clic en **"Aplicar"**.

3. La herramienta te pedirá que **elijas dónde aparecerá el mensaje de alerta**:
   - Ve al archivo de Excel y selecciona la celda o celdas donde quieres que aparezca el aviso.
   - Haz clic en **Aceptar**.

4. La herramienta te mostrará un texto de alerta sugerido. Puedes:
   - Dejarlo tal como está y hacer clic en **Aceptar**, o
   - Borrarlo y escribir tu propio mensaje personalizado, luego haz clic en **Aceptar**.

5. Aparecerá un mensaje de confirmación. Haz clic en **Aceptar**.

**¿Qué pasará?**

- Si alguien escribe un número negativo o texto distinto a "NS", Excel mostrará un error.
- Si alguien escribe "NS", el mensaje de alerta aparecerá automáticamente en el lugar que elegiste, diciendo que se necesita una justificación en los comentarios.

---

## 6. Validación 4 — Formato de Texto

**¿Para qué sirve?**  
Obliga a que el texto cumpla un formato estricto:
- ✅ Solo letras en **MAYÚSCULAS** (incluyendo Ñ y acentos: Á, É, Í, Ó, Ú)
- ✅ Solo números
- ✅ Sin dobles espacios
- ✅ Sin espacios al inicio ni al final
- ✅ Sin caracteres especiales (puntos, comas, signos, etc.)

**¿Cuándo usarla?**  
En campos de nombre de cargo, nombre de funcionario, nombre de institución, u otros campos donde se exige texto estandarizado.

---

**Pasos:**

1. Selecciona las celdas en Excel y haz clic en **"Capturar Rango"**.

2. Marca la casilla **"Formato Texto"** y haz clic en **"Aplicar"**.

3. La herramienta te pedirá que **selecciones una columna vacía** (lejos de tus datos) para alojar un motor de validación invisible.
   - Por ejemplo, selecciona una celda en la columna **CW** o cualquier columna muy a la derecha que esté vacía.
   - Haz clic en **Aceptar**.

4. Aparecerá un mensaje de confirmación indicando en qué columna quedó alojado el motor. Haz clic en **Aceptar**.

> 💡 **¿Qué es esa columna auxiliar?** Es una columna donde la herramienta guarda las fórmulas matemáticas de validación. No la borres ni la muevas, ya que es parte del funcionamiento. Puedes ocultarla si lo deseas (clic derecho en la letra de la columna → Ocultar).

**¿Qué verá el informante si escribe mal?**

> *"El texto capturado debe cumplir estrictamente las siguientes reglas:*  
> *• Todo en MAYÚSCULAS.*  
> *• Sin dobles espacios ni espacios a las orillas.*  
> *• SOLO LETRAS Y NÚMEROS. No se permiten caracteres especiales."*

---

## 7. Validación 5 — Bloqueo Dinámico

**¿Para qué sirve?**  
Bloquea celdas de captura para que **no se pueda escribir en ellas** a menos que otra celda cumpla una condición específica.

**Ejemplo práctico:**  
La celda "Especifique" solo se activa si en la pregunta anterior se eligió la opción "Otra". Si se elige cualquier otra opción, la celda queda bloqueada (sombreada en gris).

**¿Cuándo usarla?**  
Cuando una celda solo debe llenarse si se cumple una condición en otra parte del formulario.

---

**Pasos:**

1. Selecciona las celdas que quieres **bloquear/desbloquear** y haz clic en **"Capturar Rango"**.

2. Marca la casilla **"Bloqueo"** y haz clic en **"Aplicar"**.

3. La herramienta te irá haciendo preguntas. Solo sigue los pasos:

---

**Paso 1 — ¿Cuál celda controla el bloqueo?**

- Se abre una ventana. Selecciona en Excel la celda (o celdas) cuyo valor decide si se desbloquea la captura.
- Haz clic en **Aceptar**.

> Si seleccionaste un rango del mismo tamaño que el rango capturado, la herramienta te preguntará si quieres que se evalúe **fila por fila** (SÍ) o de forma **global** (NO). Elige según tu necesidad.

---

**Paso 2 — ¿Qué operador usar?**

- Se abre una ventana con los operadores disponibles:

| Operador | Significa |
|----------|-----------|
| `=`  | Igual a |
| `<>` | Diferente de |
| `>`  | Mayor que |
| `<`  | Menor que |
| `>=` | Mayor o igual a |
| `<=` | Menor o igual a |

- Escribe el operador que necesitas y haz clic en **Aceptar**.

---

**Paso 3 — ¿Cuál es el valor?**

- Escribe el valor con el que se compara. Ejemplos: `Otra`, `6`, `X`
- Haz clic en **Aceptar**.

---

**Paso 3.5 — ¿Qué pasa si la celda de condición está vacía?**

- **SÍ**: Si la celda de condición está en blanco, la celda de captura se desbloquea (se puede escribir).
- **NO**: Si la celda de condición está en blanco, la celda de captura permanece bloqueada.

Elige según tu caso y haz clic en la opción correspondiente.

---

**Paso 4 — ¿Resaltar en azul cuando está desbloqueada pero vacía?**

- **SÍ**: La celda se pintará de azul para indicar que es obligatoria pero falta llenarse.
- **NO**: No se aplica ese resaltado.

---

4. Aparecerá un mensaje confirmando el modo aplicado. Haz clic en **Aceptar**.

**¿Qué verá el informante?**

- **Celda bloqueada**: aparece sombreada con un patrón gris. Si intenta escribir, Excel muestra un mensaje de error.
- **Celda desbloqueada y vacía**: aparece en azul (si elegiste esa opción), indicando que debe llenarse.
- **Celda desbloqueada y llena**: se ve normal.

---

## 8. Validación 6 — Campos Vacíos Obligatorios

**¿Para qué sirve?**  
Detecta y resalta en azul las celdas que están vacías dentro de una fila que ya tiene datos en otras columnas. Además muestra un mensaje de alerta visible.

**Ejemplo práctico:**  
Si en una fila ya se llenaron algunas columnas, pero otras quedaron vacías, esas celdas vacías se resaltan en azul para que el capturista las complete.

**¿Cuándo usarla?**  
En matrices de captura donde todas las columnas de una fila son obligatorias si la fila tiene al menos un dato.

> ℹ️ Esta validación **no bloquea** la captura, solo resalta visualmente lo que falta.

---

**Pasos:**

1. Selecciona el rango de la **matriz de captura** (todas las celdas de la tabla que quieres vigilar) y haz clic en **"Capturar Rango"**.

2. Marca la casilla **"Blancos"** y haz clic en **"Aplicar"**.

3. La herramienta te pedirá que **elijas dónde aparecerá el mensaje de alerta**:
   - Selecciona una celda o rango en Excel para el mensaje azul.
   - Si seleccionas varias celdas, se combinarán automáticamente.
   - Haz clic en **Aceptar**.

4. La herramienta te mostrará un texto de alerta sugerido. Puedes modificarlo o dejarlo como está.
   - Haz clic en **Aceptar**.

5. Si ya existían otras validaciones en esas celdas, la herramienta te preguntará:
   - **SÍ**: Conservar los formatos previos (recomendado si ya tienes bloqueos aplicados).
   - **NO**: Borrar formatos previos y aplicar solo los de blancos.

6. Aparecerá un mensaje de confirmación. Haz clic en **Aceptar**.

**¿Qué verá el informante?**

- Las celdas vacías dentro de filas con datos se pintarán de **azul**.
- En la celda del mensaje aparecerá el aviso (ej: *"Favor de revisar la información faltante en las celdas sombreadas"*).
- Las filas completamente vacías **no se resaltan** (solo se exige completar las filas que ya tienen algún dato).

---

## 9. Validación 7 — Revisar el campo "Especifique"

**¿Para qué sirve?**  
Detecta si el texto que alguien escribió en un campo "Especifique" ya existe (o es muy similar) a alguna de las opciones del catálogo. Esto ayuda a evitar duplicados o inconsistencias.

**Ejemplo práctico:**  
El catálogo tiene la opción "Educación y formación". Si el informante escribe en el Especifique "educacion basica", la herramienta detecta que ya existe una opción similar, resalta esa opción del catálogo en amarillo y muestra un mensaje de alerta.

**¿Cuándo usarla?**  
Cuando hay un campo abierto de "Especifique" junto a un catálogo de opciones y quieres avisar si el texto libre coincide con alguna opción ya existente.

> ℹ️ Esta validación **no bloquea** la captura, solo alerta visualmente.

---

**Pasos:**

1. Selecciona la **celda del campo "Especifique"** (donde el informante escribe libremente) y haz clic en **"Capturar Rango"**.

2. Marca la casilla **"Esp. Clave"** y haz clic en **"Aplicar"**.

3. La herramienta te hará tres preguntas:

---

**Pregunta 1 — ¿Cuáles son las opciones del catálogo?**

- Selecciona en Excel las celdas que contienen las opciones del catálogo (las que tienen texto como "Educación y formación", "Salud pública", etc.).
- Haz clic en **Aceptar**.

---

**Pregunta 2 — ¿Dónde mostrará el mensaje de alerta?**

- Selecciona la celda donde aparecerá el mensaje cuando se detecte una coincidencia.
- Haz clic en **Aceptar**.

---

**Pregunta 3 — ¿Dónde alojar el motor de búsqueda?**

- La herramienta necesita espacio vacío en la hoja para construir su motor interno.
- Selecciona una **celda vacía lejos de tus datos** (columna AF en adelante).
- Necesita aproximadamente **2 columnas de ancho** y tantas filas como opciones tenga tu catálogo.
- Haz clic en **Aceptar**.

---

4. La herramienta construirá el motor automáticamente. Cuando termine, aparecerá el mensaje de confirmación. Haz clic en **Aceptar**.

**¿Qué verá el informante?**

- Si lo que escribió en "Especifique" coincide con alguna opción del catálogo, esa opción se resaltará en **amarillo**.
- Aparecerá el mensaje: *"Alerta: Revise el texto ingresado en el Especifique ya que podría existir en las opciones resaltadas en amarillo"*.
- Si no hay coincidencia, nada cambia.

---

## 10. Validación 8 — Validación de Años

**¿Para qué sirve?**  
Restringe la celda para que solo acepte un año numérico dentro de un rango definido por ti (por ejemplo, entre 1821 y 2026), o el código "NS".

**¿Cuándo usarla?**  
En campos donde se captura un año (año de nacimiento, año de fundación, año de inicio de funciones, etc.).

---

**Pasos:**

1. Selecciona las celdas en Excel y haz clic en **"Capturar Rango"**.

2. Marca la casilla **"Años"** (o "Fechas") y haz clic en **"Aplicar"**.

3. La herramienta te pedirá el **límite inferior** (el año más antiguo permitido):
   - Verás un valor sugerido de `1821`. Puedes cambiarlo si necesitas otro año.
   - Haz clic en **Aceptar**.

4. Luego te pedirá el **límite superior** (el año más reciente permitido):
   - Verás un valor sugerido de `2026`. Cámbialo si es necesario.
   - Haz clic en **Aceptar**.

5. Aparecerá un mensaje de confirmación. Haz clic en **Aceptar**.

**¿Qué pasará cuando alguien escriba mal?**

Si el informante escribe un año fuera del rango o texto no permitido, aparecerá este mensaje:

> *"El valor ingresado debe ser un año válido de 4 dígitos entre [límite inferior] y [límite superior], o el código 'NS'."*

**Valores aceptados:**
- ✅ Cualquier número entero dentro del rango definido (ej: 1995, 2010, 2024)
- ✅ El código `NS` (sin espacios, en mayúsculas)
- ❌ Decimales, letras, años fuera del rango

---

## 11. Preguntas Frecuentes

---

**¿Puedo aplicar más de una validación a las mismas celdas?**

Sí. Por ejemplo, puedes aplicar primero el **Bloqueo Dinámico** y luego los **Campos Vacíos** sobre el mismo rango. Cuando apliques la segunda validación, la herramienta te preguntará si deseas conservar las reglas anteriores. Responde **SÍ** para no perder lo que ya aplicaste.

---

**¿Qué hago si me equivoqué y quiero quitar una validación?**

Selecciona las celdas en Excel, ve a la pestaña **Datos** en la cinta de opciones de Excel, haz clic en **Validación de Datos** y luego en **Borrar todo**. También puedes ir a **Inicio → Formato condicional → Borrar reglas**.

---

**¿Por qué la herramienta pone cosas en columnas lejanas de mi hoja?**

Algunas validaciones (como Formato Texto y Especifique por Palabras Clave) necesitan columnas auxiliares para funcionar. La herramienta te pregunta dónde colocarlas. **No borres esas columnas**, ya que son parte de la validación. Puedes ocultarlas si no quieres verlas (clic derecho en la letra de la columna → Ocultar).

---

**¿Qué significa cuando una celda se pone de color azul?**

Significa que esa celda **está vacía y debería llenarse**. Lo aplican las validaciones de Blancos y Bloqueo Dinámico para guiar visualmente al capturista.

---

**¿Qué significa cuando una celda tiene un patrón gris cruzado?**

Significa que esa celda está **bloqueada**: no se puede escribir en ella porque la condición de desbloqueo no se ha cumplido. Lo aplica la validación de Bloqueo Dinámico.

---

**¿Puedo cambiar el texto de los mensajes de alerta?**

Sí, en las validaciones **NS** y **Campos Vacíos** la herramienta te pide que escribas o confirmes el texto del mensaje antes de aplicarlo. Puedes personalizarlo libremente.

---

**El mensaje de error aparece pero el dato ya estaba correcto, ¿por qué?**

Puede ocurrir en la validación de **Formato Texto** si la celda tiene espacios invisibles al inicio o al final, o si tiene algún carácter especial que no se ve a simple vista. Intenta borrar el contenido de la celda y escribirlo de nuevo.

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
**Versión**: 2.3.0 | **Fecha**: 11 de julio de 2026
