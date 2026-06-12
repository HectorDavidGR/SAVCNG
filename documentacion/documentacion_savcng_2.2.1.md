# SAVCNG ExcelDNA - Documentación Técnica Completa

**Versión**: 2.2.1.0  
**Última Actualización**: 11 de junio de 2026 (Revisión)  
**Estado**: En Desarrollo Continuo  

## Tabla de Contenidos
1. [Descripción General](#descripción-general)
2. [Arquitectura del Proyecto](#arquitectura-del-proyecto)
3. [Componentes Principales](#componentes-principales)
4. [Sistema de Validaciones Detallado](#sistema-de-validaciones-detallado)
   - [4.1 Validación de Decimales (Enteros)](#41-validación-de-decimales-enteros)
   - [4.2 Validación de Catálogos (Listas Desplegables)](#42-validación-de-catálogos-listas-desplegables)
   - [4.3 Validación NS (No Aplica)](#43-validación-ns-no-aplica)
   - [4.4 Validación de Formato Texto](#44-validación-de-formato-texto)
   - [4.5 Validación de Bloqueo Dinámico](#45-validación-de-bloqueo-dinámico-⭐-nuevo)
5. [FrmValidaciones.cs - Análisis Profundo](#frmvalidacionescs---análisis-profundo)
6. [Flujo de Trabajo](#flujo-de-trabajo)
7. [Especificaciones Técnicas](#especificaciones-técnicas)
8. [Guía de Desarrollo](#guía-de-desarrollo)
9. [Patrones de Código y Mejores Prácticas](#patrones-de-código-y-mejores-prácticas)
10. [Troubleshooting y Casos Especiales](#troubleshooting-y-casos-especiales)

---

## Descripción General

**SAVCNG ExcelDNA** es un complemento (Add-in) de Microsoft Excel desarrollado en C# (.NET Framework 4.8) que proporciona un sistema avanzado de validación de datos para archivos de censo de gobierno. 

El proyecto utiliza **ExcelDNA** como framework de integración con Excel, permitiendo crear un menú personalizado en la cinta de opciones (Ribbon) de Excel desde el cual se pueden cargar archivos de censo y aplicar múltiples validaciones de integridad de datos.

### Propósito Principal
- Garantizar la calidad e integridad de los datos en archivos de censo
- Proporcionar una interfaz gráfica intuitiva para validación de datos
- Detectar y resaltar errores de formato y contenido
- Facilitar listas desplegables (Data Validation) personalizadas
- Validar presencia de respuestas "No Aplica" (NS) con alertas contextuales
- Validar formato de texto: mayúsculas, sin espacios extra, caracteres válidos
- Crear reglas de bloqueo dinámico condicionadas a valores de otras celdas
- Soporte automático para matrices y evaluaciones fila por fila

### Tecnologías Utilizadas
- **Lenguaje**: C# 7.3
- **Framework**: .NET Framework 4.8
- **Librería de Interop**: Microsoft.Office.Interop.Excel
- **Plugin Framework**: ExcelDNA
- **UI**: Windows Forms (.NET Framework)

---

## Arquitectura del Proyecto

```
SAVCNG_ExcelDNA/
├── MenuCenso.cs                    (Punto de entrada y menú)
├── FrmValidaciones.cs              (Lógica principal de validaciones)
├── FrmValidaciones.Designer.cs     (Diseño de formulario)
├── Properties/
│   └── AssemblyInfo.cs             (Metadatos del ensamblado)
├── bin/                            (Ejecutables compilados)
├── obj/                            (Artefactos compilados)
└── documentacion/
	└── documentacion_savcng_YYYY-MM-DD.md    (Documentación técnica)
```

### Patrón Arquitectónico
- **Model**: Las colecciones (List, Dictionary) almacenan datos de validación y rangos
- **View**: Windows Forms (FrmValidaciones) para la interfaz de usuario
- **Controller**: Métodos de eventos que procesan interacciones del usuario
- **Service**: Métodos estáticos y funciones auxiliares para lógica compartida

---

## Componentes Principales

### 1. **MenuCenso.cs** - Punto de Entrada

Responsable de la integración con Excel y la interfaz de usuario del ribbon.

#### Características Principales:
```csharp
public class MenuCenso : ExcelRibbon
{
	// Define un menú personalizado en la cinta de Excel
	public override string GetCustomUI(string RibbonID)
}
```

#### Funcionalidades:
- **XML Ribbon**: Define la estructura del menú en la pestaña "SAVCNG - Censo"
- **Diálogo de Archivo**: Permite seleccionar archivos de censo (.xlsx, .xlsm, .xlsb)
- **Gestión de Workbook**: Abre el archivo y lo pasa a la ventana de validaciones
- **Manejo de Errores**: Captura excepciones durante la carga de archivos

#### Flujo de Ejecución:
1. Usuario hace clic en "Cargar Censo" en el menú de Excel
2. Se abre un diálogo `OpenFileDialog`
3. El archivo seleccionado se abre en Excel mediante `excelApp.Workbooks.Open()`
4. Se instancia `FrmValidaciones` pasando el workbook
5. La ventana se muestra flotante sobre Excel

---

### 2. **FrmValidaciones.cs** - Motor de Validaciones

Este es el corazón del aplicativo. Implementa todas las validaciones de datos y maneja la interacción del usuario.

#### 2.1 Variables de Estado Principales

```csharp
private Excel.Workbook _libroCenso;
```
- **Propósito**: Referencia al archivo de Excel cargado
- **Tipo**: `Excel.Workbook` (objeto COM de Office)
- **Inicialización**: En el constructor, recibida desde `MenuCenso.cs`
- **Ciclo de vida**: Se mantiene mientras la ventana esté abierta
- **Notas**: Es un objeto COM, requiere manejo cuidadoso de memoria

```csharp
private Excel.Range _rangoCapturado;
```
- **Propósito**: Almacena el rango de celdas seleccionado por el usuario
- **Tipo**: `Excel.Range` (objeto COM de Office)
- **Inicialización**: Se asigna en `btnCapturarRango_Click()`
- **Validación**: Se verifica que sea un rango (no imagen, gráfico, etc.)
- **Uso**: Base para aplicar todas las validaciones
- **Notas**: Puede contener múltiples áreas (rangos no contiguos)

---

## Sistema de Validaciones Detallado

### 4.1 Validación de Decimales (Enteros)

**Objetivo**: Detectar y resaltar celdas que contienen números con decimales cuando solo se permiten números enteros.

**Características**:
- Formato condicional visual (amarillo fondo, rojo texto)
- Detección automática de decimales
- No es restrictiva (solo visual)

---

### 4.2 Validación de Catálogos (Listas Desplegables)

**Objetivo**: Crear listas desplegables (Data Validation) con opciones predefinidas.

**Características Avanzadas**:
- Soporte para celda única, rango de celdas y celdas combinadas
- Opción de extraer números de texto
- Limpieza automática de caracteres innecesarios

---

### 4.3 Validación NS (No Aplica) - Caso Especializado

**Objetivo**: Permitir números >= 0 o valor "NS", con alerta automática cuando se usa NS.

**Contexto**: En censos, "NS" significa "No Sabe" o "No Aplica" y requiere justificación.

**Características**:
- Validación restrictiva: solo permite números >= 0 o "NS"
- Alerta automática cuando hay registros NS
- Mensaje combinado y formateado en fila siguiente
- Manejo de rangos no contiguos (Areas)

---

### 4.4 Validación de Formato Texto

**Objetivo**: Garantizar que el texto cumpla formato específico: mayúsculas, sin espacios extra, caracteres válidos.

**Reglas Implementadas**:
1. ✓ Solo mayúsculas (A-Z, Ñ)
2. ✓ Espacios simples (sin dobles espacios)
3. ✓ Sin espacios al inicio o final

**Características**:
- Data Validation restrictiva (el usuario NO puede violar las reglas)
- Mensaje de error personalizado
- Compatible con celdas vacías (IgnoreBlank = true)

---

### 4.5 Validación de Bloqueo Dinámico ⭐ MEJORADO (v2.2.1)

**Objetivo**: Crear reglas de bloqueo condicional basadas en valores de otras celdas o rangos, con inteligencia automática para detectar matrices y evaluaciones fila por fila.

**Contexto de Uso**: En formularios de censo:
- Campo "Especifique" que solo se activa si respuesta anterior es "Otra"
- Campos de justificación que se habilitan si hay respuesta "NS"
- Matrices paralelas donde cada fila tiene su propia condición
- Validaciones cascada donde un campo depende de múltiples condiciones

**Casos de Uso Prácticos**:
```
Ejemplo 1: Campo "Especifique Otra" (Simple)
├─ Si [P1] = "Otra" → Se desbloquea [P2 Especifique]
├─ Si [P1] ≠ "Otra" → Se bloquea [P2] (gris)
└─ Si [P1] = "Otra" pero [P2] vacío → Se resalta en rojo

Ejemplo 2: Matriz Paralela (Inteligencia Automática) ⭐ NUEVO
├─ Si selecciona rango de 10 filas en P1 y rango de 10 filas en P1.1
├─ Sistema detecta automáticamente: "Mismo tamaño, ¿fila por fila?"
├─ Si responde SÍ: Fila 1 bloquea fila 1, fila 2 bloquea fila 2, etc.
└─ Si responde NO: Cualquier valor en P1:P10 desbloquea todo P1.1:P1.10

Ejemplo 3: Mensaje de Alerta Dinámico ⭐ NUEVO
├─ Especifique el texto del mensaje (ej: "Especifique el nombre")
├─ Seleccione dónde aparecerá el mensaje
└─ Sistema muestra automáticamente cuando se cumple condición + campo vacío
```

**Mejoras en v2.2.1**:
1. ✨ **Detección Automática de Matrices**: Si rangos miden igual, pregunta si es fila por fila
2. ✨ **Direcciones Inteligentes**: `dirCondicionLocal` para fila por fila, `dirCondicionGlobal` para búsqueda global
3. ✨ **Mensaje Dinámico en Celda**: Permite crear alertas personalizadas que aparecen automáticamente
4. ✨ **Fórmula Universal**: Usa sintaxis INGLÉS para máxima compatibilidad multiidioma
5. ✨ **Modo de Operación Mostrado**: Mensaje final indica qué modo se aplicó (Fila por Fila vs Búsqueda Global)

**Características Avanzadas**:
1. **Motor de Búsqueda CONTAR.SI**: Busca valores en rango/celda
2. **Soporte de Operadores**: =, <>, >, <, >=, <=
3. **Cinco Visuales Dinámicos**:
   - Verde (desbloqueado con valor)
   - Rojo (desbloqueado pero vacío - obligatorio)
   - Gris (bloqueado por condición no cumplida)
   - Amarillo (instrucción relacionada activa)
   - Rojo dinámico (mensaje de alerta personalizado)
4. **Inteligencia Fila-por-Fila**: Detección automática de matrices paralelas
5. **Manejo de Rangos No Contiguos**: Funciona con múltiples áreas

**Flujo del Usuario - 6 Pasos** (v2.2.1):

1. **Seleccionar Condición** (Rango/Celda de Referencia)
   - Una celda → Automáticamente fila por fila
   - Rango igual tamaño → Sistema pregunta si es fila por fila
   - Rango diferente tamaño → Automáticamente búsqueda global

2. **Elegir Operador** (=, <>, >, <, >=, <=)

3. **Ingresar Valor** (Número, texto, etc.)

4. **¿Resalte Rojo?** (Opcional)
   - Sí: Cuando se desbloquee y esté vacío, resalta en rojo
   - No: Solo sombreado gris cuando bloqueado

5. **¿Resalte de Instrucción?** (Opcional)
   - Sí: Seleccionar celda con instrucción que se resaltará en amarillo
   - No: Continuar

6. **¿Mensaje de Alerta Dinámico?** (Nuevo en v2.2.1)
   - Sí: Ingresar texto → Seleccionar ubicación → Aparece automáticamente
   - No: Finalizar

**Implementación Técnica - Generalización v2.2.1**:

```csharp
// PASO 1: Detección Automática de Tipo de Rango
bool esFilaPorFila = false;

if (rangoCondicion.Count == 1)
{
	// Una sola celda siempre es fila por fila
	esFilaPorFila = true;
}
else if (_rangoCapturado.Rows.Count == rangoCondicion.Rows.Count)
{
	// ¡Magia! Si miden igual, preguntar al usuario
	DialogResult respFila = MessageBox.Show(
		"He detectado que el rango de condición tiene el MISMO número de filas...\n\n" +
		"¿Fila por Fila (recomendado para matrices)?",
		"Evaluación Inteligente",
		MessageBoxButtons.YesNo,
		MessageBoxIcon.Question);

	esFilaPorFila = (respFila == DialogResult.Yes);
}
// Si no cumple ninguna condición, se mantiene en búsqueda global

// Direcciones Inteligentes
string dirCondicionLocal = esFilaPorFila
	? rangoCondicion.Cells[1, 1].Address.Replace("$", "")  // A1 (relativo)
	: rangoCondicion.Address;                               // $A$1:$A$21 (fijo)

string dirCondicionGlobal = rangoCondicion.Address;  // Siempre fijo para instrucción

// PASO 6: Mensaje Dinámico (Nueva característica)
if (respuestaAlerta == DialogResult.Yes)
{
	// Solicitar texto del mensaje
	string textoAlerta = excelApp.InputBox("Escribe el mensaje...");

	// Solicitar ubicación
	Excel.Range rangoAlerta = (Excel.Range)excelApp.InputBox("Selecciona ubicación...");

	// Combinar si es múltiple
	if (rangoAlerta.Count > 1)
		rangoAlerta.Merge();

	// Fórmula Universal (INGLÉS): Muestra si condición se cumple Y campo vacío
	string formulaAlerta = $"=IF(COUNTIF({dirCondicionGlobal},{criterioContarSi})>COUNTA({dirCapturadaGlobal}),\"{textoAlerta}\",\"\")";

	rangoAlerta.Formula = formulaAlerta;  // Usa .Formula, no .FormulaLocal
}
```

#### 4.5.1 Estados Visuales Completos

```
Estado 1: BLOQUEADO
├─ Condición: NO se cumple
├─ Visual: Sombreado gris con patrón CrissCross
├─ Data Validation: RECHAZA entrada
└─ Usuario: No puede escribir

Estado 2: DESBLOQUEADO - CON VALOR
├─ Condición: Se cumple Y celda tiene contenido
├─ Visual: Normal (sin resalte)
├─ Data Validation: PERMITE entrada
└─ Usuario: Puede ingresar datos

Estado 3: DESBLOQUEADO - VACÍO (Resalte Rojo)
├─ Condición: Se cumple pero celda está vacía
├─ Visual: Rojo suave (#FFC7CE) con letra roja
├─ Data Validation: PERMITE (obligatorio)
└─ Indicador: "Este campo es obligatorio"

Estado 4: INSTRUCCIÓN RELACIONADA
├─ Condición: Se cumple EN ALGUNA FILA
├─ Visual: Amarillo con negritas (resalta)
├─ Uso: Guía visual del capturista
└─ Nota: Usa dirCondicionGlobal (busca en todo el bloque)

Estado 5: MENSAJE PERSONALIZADO ⭐ NUEVO
├─ Condición: Se cumple pero campo vacío
├─ Visual: Texto rojo, centrado, negritas
├─ Contenido: Personalizado por usuario (ej: "Especifique el nombre")
└─ Dinámica: Aparece/desaparece según condición
```

#### 4.5.2 Ejemplo Práctico Mejorado: Matriz de Incisos

**Escenario**:
```
Pregunta 5: Indique el grado de educación (Seleccione una por inciso):
────────────────────────────────────────────
		 Inciso A | Inciso B | Inciso C
Primaria    □    |    □     |    □
Secundaria  □    |    □     |    □
Terciaria   □    |    □     |    □
────────────────────────────────────────────
Especifique: [BLOQUEADO hasta seleccionar "Terciaria"]
```

**Configuración del Bloqueo (v2.2.1)**:
1. Capturar rango: E5:G7 (matriz de 3 filas × 3 columnas)
2. Paso 1: Seleccionar condición = D5:D7 (misma altura, 3 filas)
   - Sistema detecta: "¿Fila por fila?" → Sí
3. Paso 2: Operador = "="
4. Paso 3: Valor = "Terciaria"
5. Paso 4: ¿Resalte rojo? = Sí
6. Paso 5: ¿Resalte instrucción? = No
7. Paso 6: ¿Mensaje de alerta? = Sí
   - Texto: "Especifique nivel de educación alcanzado"
   - Ubicación: E8 (debajo de la matriz)

**Comportamiento Resultante**:
```
Fila 1 (Primaria):
  └─ E5:G5 → Gris (bloqueado)

Fila 2 (Secundaria):
  └─ E6:G6 → Gris (bloqueado)

Fila 3 (Terciaria):
  └─ Si selecciona en D7="Terciaria"
	 ├─ E7:G7 → Normal (desbloqueado)
	 ├─ Si vacío → Rojo (obligatorio)
	 └─ E8 → Muestra mensaje en rojo

Resultado: INDEPENDENCIA POR FILA
- Cada fila se maneja por su propia condición
- No interfiere entre filas
- Ideal para datos paralelos
```

#### 4.5.3 Comparativa v2.2.0 vs v2.2.1

| Aspecto | v2.2.0 | v2.2.1 |
|---------|--------|--------|
| Detección de tipo | Manual | ✨ Automática |
| Soporte matriz fila-por-fila | No | ✨ Sí (Inteligente) |
| Direcciones | Simples | ✨ Local + Global |
| Mensaje personalizado | No | ✨ Sí (Dinámico) |
| Fórmula | Local | ✨ Universal (INGLÉS) |
| Modo mostrado | No | ✨ Mostrado en mensaje |
| Pasos | 5 | ✨ 6 |
| Casos de uso | Simples | ✨ Complejos |

#### 4.5.4 Desglose de Fórmulas v2.2.1

**Fórmula de Validación (Fila por Fila)**:
```
=CONTAR.SI(A1;\"=\"&6)>0
├─ A1: Celda relativa (se adapta a cada fila)
├─ CONTAR.SI(...;\"=\"&6): Busca exactamente "=6"
└─ >0: Verdadero si encuentra coincidencia
```

**Fórmula de Validación (Búsqueda Global)**:
```
=CONTAR.SI($A$1:$A$10;\"=\"&6)>0
├─ $A$1:$A$10: Rango fijo (siempre busca en todo)
├─ CONTAR.SI(...): Cuenta coincidencias en toda la lista
└─ >0: Verdadero si encuentra AL MENOS una
```

**Fórmula de Mensaje Dinámico (v2.2.1)**:
```
=IF(COUNTIF($A$1:$A$10,\"=\"&6)>COUNTA($B$1:$B$10),\"Especifique el nombre\",\"\")
├─ COUNTIF($A$1:$A$10,\"=\"&6): Cuenta coincidencias en condición
├─ COUNTA($B$1:$B$10): Cuenta celdas con contenido en captura
├─ >: Si coincidencias > contenido → Hay campos obligatorios vacíos
└─ IF(...): Si verdadero muestra mensaje, si no muestra vacío
```

---

## FrmValidaciones.cs - Análisis Profundo

### 5.1 Método: btnAplicar_Click - Validación de Bloqueo (v2.2.1)

**Responsabilidad**: Orquestar la validación de bloqueo dinámico con detección automática de matrices.

**Estructura de Decisión Inteligente**:

```csharp
// PASO 1: Detectar automáticamente tipo de rango
bool esFilaPorFila = false;

// Caso 1: Una sola celda
if (rangoCondicion.Count == 1)
{
	esFilaPorFila = true;  // Automático
}
// Caso 2: Múltiples celdas del mismo alto
else if (_rangoCapturado.Rows.Count == rangoCondicion.Rows.Count)
{
	// Preguntar al usuario (es matriz potencial)
	DialogResult respuesta = MessageBox.Show(
		"¿Deseas evaluación FILA POR FILA?",
		"Matriz Detectada",
		MessageBoxButtons.YesNo);

	esFilaPorFila = (respuesta == DialogResult.Yes);
}
// Caso 3: Otros casos (implícitamente búsqueda global)

// PASO 2: Generar direcciones según tipo detectado
string dirCondicionLocal = esFilaPorFila
	? rangoCondicion.Cells[1, 1].Address.Replace("$", "")  // Relativa
	: rangoCondicion.Address;                               // Absoluta

string dirCondicionGlobal = rangoCondicion.Address;  // Siempre absoluta

// PASO 6: Mensaje de alerta personalizado
if (respuestaAlerta == DialogResult.Yes)
{
	// Obtener texto y ubicación
	string textoAlerta = excelApp.InputBox(...);
	Excel.Range rangoAlerta = (Excel.Range)excelApp.InputBox(...);

	// Fórmula compatible multiidioma
	string formulaAlerta = 
		$"=IF(COUNTIF({dirCondicionGlobal},{criterioContarSi})>COUNTA({dirCapturadaGlobal})," +
		$"\"{textoAlerta}\",\"\")";

	rangoAlerta.Formula = formulaAlerta;  // Usa .Formula (no .FormulaLocal)
}
```

---

## Flujo de Trabajo Completo v2.2.1

```
Validación Bloqueo Dinámico:
  │
  ├─→ PASO 1: Usuario selecciona rango de condición
  │     ├─ Celda única → esFilaPorFila = TRUE
  │     ├─ Rango igual alto → Preguntar usuario
  │     └─ Rango diferente → esFilaPorFila = FALSE
  │
  ├─→ PASO 2: Usuario elige operador
  │
  ├─→ PASO 3: Usuario ingresa valor de criterio
  │
  ├─→ PASO 4: ¿Resalte rojo? (Sí/No)
  │     └─ Sí → Formato rojo suave cuando desbloqueado + vacío
  │
  ├─→ PASO 5: ¿Resalte instrucción? (Sí/No)
  │     └─ Sí → Seleccionar celda para amarillo
  │
  ├─→ PASO 6: ¿Mensaje personalizado? (Nuevo en v2.2.1)
  │     ├─ Sí → Ingresar texto del mensaje
  │     │     → Seleccionar ubicación
  │     │     → Sistema aplica fórmula dinámica
  │     └─ No → Ir a aplicación
  │
  └─→ APLICACIÓN DE REGLAS:
	  ├─ Data Validation (CONTAR.SI con dirCondicionLocal)
	  ├─ Formato Gris (bloqueado)
	  ├─ Formato Rojo (opcional, desbloqueado + vacío)
	  ├─ Formato Amarillo (opcional, instrucción)
	  └─ Mensaje dinámico (opcional, alerta personalizada)
```

---

## Especificaciones Técnicas

### Requisitos del Sistema

| Aspecto | Requisito |
|--------|-----------|
| **Versión .NET** | .NET Framework 4.8 |
| **Versión C#** | 7.3 |
| **Versión Excel** | 2013 o superior |
| **SO** | Windows 7/8/10/11 (32 o 64 bits) |
| **Memoria RAM Mínima** | 4 GB |
| **Espacio en Disco** | ~100 MB |
| **Bitness** | Mismo que Office instalado (32 o 64) |

### Rendimiento v2.2.1

| Operación | Tiempo Estimado | Variables |
|-----------|-----------------|-----------|
| Detección automática de matriz | Inmediato | Comparación de alturas |
| Aplicación de validación bloqueo | 2-4 seg | Complejidad de fórmula |
| Mensaje dinámico | < 1 seg | Longitud del texto |
| Total (completo con 6 pasos) | 10-15 seg | Interacción usuario |

### Límites Conocidos

- **Máximo criterio CONTAR.SI**: ~256 caracteres
- **Máximo texto de mensaje**: 255 caracteres (limitación Excel)
- **Máximo rango**: 1,048,576 filas × 16,384 columnas
- **Máximo rangos no contiguos**: Sin límite teórico

---

## Patrones de Código v2.2.1

### Pattern 1: Detección Automática de Tipo de Rango

```csharp
bool esFilaPorFila = false;

if (rangoCondicion.Count == 1)
{
	esFilaPorFila = true;
}
else if (_rangoCapturado.Rows.Count == rangoCondicion.Rows.Count)
{
	// Pregunta solo si es potencial matriz
	DialogResult resp = MessageBox.Show("¿Fila por fila?", ...);
	esFilaPorFila = (resp == DialogResult.Yes);
}
```

### Pattern 2: Direcciones Dinámicas

```csharp
string dirLocal = esFilaPorFila
	? rangoCondicion.Cells[1, 1].Address.Replace("$", "")
	: rangoCondicion.Address;

string dirGlobal = rangoCondicion.Address;
```

### Pattern 3: Fórmula Universal

```csharp
// IMPORTANTE: Usar sintaxis INGLÉS para máxima compatibilidad
string formula = $"=IF(COUNTIF({dir},{criterio})>0,\"mensaje\",\"\")";
rangoAlerta.Formula = formula;  // NUNCA .FormulaLocal para IF/COUNTIF
```

---

## Troubleshooting v2.2.1

### Problema 1: Detección de matriz no funciona

**Síntoma**: Sistema no pregunta por fila por fila aunque rangos midan igual.

**Causa**: El rango de captura tiene distinto número de filas que el de condición.

**Solución**:
```csharp
// Verificar que ambos rangos tengan mismo alto
int filasCaptura = _rangoCapturado.Rows.Count;
int filasCondicion = rangoCondicion.Rows.Count;

System.Diagnostics.Debug.WriteLine($"Captura: {filasCaptura}, Condición: {filasCondicion}");
// Deben ser iguales para activar la pregunta
```

### Problema 2: Mensaje dinámico no aparece

**Síntoma**: Fórmula está pero no muestra el texto.

**Causa**: COUNTA está contando más celdas de las esperadas.

**Solución**:
```csharp
// Asegurarse que dirCapturadaGlobal apunta a rango correcto
string dirCapturada = _rangoCapturado.Address;  // Debe ser el rango a validar

// Verificar en Excel: =COUNTA(B1:B10) debe contar solo celdas con datos
```

### Problema 3: Fórmula en español no funciona

**Síntoma**: Mensaje personalizado no aparece en sistema español.

**Causa**: Se usó .FormulaLocal con IF/COUNTIF.

**Solución**:
```csharp
// CORRECTO: Usar .Formula (sintaxis inglés universal)
rangoAlerta.Formula = "=IF(COUNTIF(...),\"mensaje\",\"\")";

// INCORRECTO: No usar .FormulaLocal
// rangoAlerta.FormulaLocal = "=SI(CONTAR.SI(...),\"mensaje\",\"\")";
```

---

## Resumen de Cambios v2.2.1

### ✨ Nuevas Características

1. **Detección Automática de Matrices**
   - Identifica si rangos tienen igual altura
   - Pregunta inteligentemente si es fila por fila
   - Elimina necesidad de configuración manual

2. **Direcciones Dinámicas**
   - `dirCondicionLocal`: Relativa para fila-por-fila (A1)
   - `dirCondicionGlobal`: Absoluta para búsqueda global ($A$1:$A$10)
   - Aplicación correcta según contexto

3. **Mensaje de Alerta Dinámico**
   - Usuario ingresa texto personalizado
   - Selecciona ubicación en formulario
   - Fórmula universal (INGLÉS) para compatibilidad
   - Aparece automáticamente cuando se cumple condición + campo vacío

4. **Modo Operacional Mostrado**
   - Mensaje final indica: "Fila por Fila" o "Búsqueda Global"
   - Feedback claro al usuario sobre configuración aplicada

### 🔧 Mejoras Técnicas

1. Mejor generalización del código de bloqueo
2. Eliminación de duplicación lógica
3. Fórmulas más robustas con sintaxis universal
4. Mejor separación de responsabilidades (Local vs Global)

### 📊 Matriz de Validaciones Disponibles (v2.2.1)

| Validación | Tipo | Motor | Pasos | Versión |
|-----------|------|-------|-------|---------|
| Decimales | FormatCondition | ESNUMERO + TRUNCAR | 1 | v1.0 |
| Catálogos | Data Validation (List) | Rango directo | 2-3 | v1.0 |
| NS | Data Validation (Custom) | O(Y(ESNUMERO, >=0), "NS") | 1 | v2.0 |
| Formato Texto | Data Validation (Custom) | IGUAL + ESPACIOS | 1 | v2.1 |
| **Bloqueo Dinámico** | Data Validation + Format | CONTAR.SI | **6** | **v2.2.1 ⭐** |

---

## Conclusión

SAVCNG ExcelDNA en versión 2.2.1 introduce mejoras significativas a la validación de Bloqueo Dinámico:

**Mejoras clave de v2.2.1**:
✓ Detección automática e inteligente de matrices paralelas
✓ Direcciones dinámicas (Local vs Global)
✓ Mensaje de alerta personalizado por usuario
✓ Fórmula universal para máxima compatibilidad multiidioma
✓ Flujo mejorado de 6 pasos iteractivo
✓ Mejor feedback visual del modo operacional

**Fortalezas acumulativas (v2.0 → v2.2.1)**:
✓ Interfaz intuitiva y accesible
✓ Soporte multiidioma mediante funciones localizadas
✓ Validaciones versátiles con diferentes enfoques
✓ Detección automática de contexto
✓ Restricción de entrada en tiempo real
✓ Lógica condicional compleja pero inteligente
✓ **Soporte automático para matrices y datos paralelos**
✓ **Mensajes personalizados y dinámicos**

**Versión**: 2.2.1.0  
**Última Actualización**: 11 de junio de 2026 (Revisión)  
**Estado**: En Desarrollo Continuo  
**Rama de Git**: Desarrollo

---

## Apéndice: Cambios Detallados desde v2.2.0 a v2.2.1

### Archivo: FrmValidaciones.cs

**Adiciones**:
- Lógica de detección automática de matriz (líneas 448-474)
- Variables `dirCondicionLocal` y `dirCondicionGlobal` (líneas 477-482)
- Paso 6 de mensaje de alerta dinámico (líneas 621-672)
- Fórmula universal en INGLÉS con .Formula (línea 669)
- Modo operacional mostrado en mensaje final (línea 674)

**Cambios**:
- Generalización de direcciones según tipo de rango
- Reemplazo de `get_Address()` por `.Address` (cumplimiento COM)
- Fórmula universal para máxima compatibilidad

**Líneas Modificadas**: ~60 líneas
**Líneas Agregadas**: ~45 líneas (Paso 6)
**Complejidad Ciclomática**: +2 ramas

### Consideraciones de Compatibilidad

- ✓ Compatible con Excel 2013+
- ✓ Compatible con .NET Framework 4.8
- ✓ Compatible con todos los idiomas localizados (español, inglés, francés, etc.)
- ✓ Retrocompatible con validaciones v2.0.x y v2.1.x
- ✓ Archivos generados anteriormente funcionan sin cambios

---

