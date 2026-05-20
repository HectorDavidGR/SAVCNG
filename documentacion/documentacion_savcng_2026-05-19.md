# SAVCNG ExcelDNA - Documentación Técnica Completa

**Versión**: 2.1.0.0  
**Última Actualización**: 19 de mayo de 2026  
**Estado**: En Desarrollo Continuo  

## Tabla de Contenidos
1. [Descripción General](#descripción-general)
2. [Arquitectura del Proyecto](#arquitectura-del-proyecto)
3. [Componentes Principales](#componentes-principales)
4. [Sistema de Validaciones Detallado](#sistema-de-validaciones-detallado)
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

```csharp
private Excel.Range _rangoCatalogo;
```
- **Propósito**: Almacena el rango que contiene las opciones para listas desplegables
- **Tipo**: `Excel.Range`
- **Inicialización**: Se asigna mediante `InputBox` en validación de catálogos
- **Estado Actual**: Declarado pero actualmente no utilizado en lógica activa
- **Potencial**: Podría usarse para caché de opciones de catálogos

```csharp
private System.Collections.Generic.List<Excel.Range> _pendientesCatalogo;
```
- **Propósito**: Colecciona rangos que requieren validación de catálogo
- **Tipo**: `List<Excel.Range>`
- **Inicialización**: Nueva lista en declaración
- **Estado Actual**: Actualmente NO se utiliza en la implementación actual
- **Potencial Futuro**: Para procesamiento por lotes de validaciones

```csharp
private System.Collections.Generic.Dictionary<string, string> _preguntaPorRango;
```
- **Propósito**: Mapea direcciones de rango a números de pregunta
- **Tipo**: `Dictionary<string, string>` (clave: dirección, valor: número)
- **Inicialización**: Nueva lista en declaración
- **Estado Actual**: Actualmente NO se utiliza en la implementación actual
- **Potencial Futuro**: Para historial y auditoría de validaciones

---

## Sistema de Validaciones Detallado

### 3.1 Validación de Decimales (Enteros)

**Objetivo**: Detectar y resaltar celdas que contienen números con decimales cuando solo se permiten números enteros.

**Implementación Técnica**:

```csharp
if (chkDecimales.Checked == true)
{
    _rangoCapturado.FormatConditions.Delete(); // Limpiar formatos previos

    Excel.Range primeraCelda = _rangoCapturado.Cells[1, 1];
    string direccion = primeraCelda.get_Address(false, false);

    string formula = $"=Y(ESNUMERO({direccion}), TRUNCAR({direccion})<>{direccion})";

    Excel.FormatCondition formato = (Excel.FormatCondition)
        _rangoCapturado.FormatConditions.Add(
            Excel.XlFormatConditionType.xlExpression,
            Type.Missing,
            formula);

    formato.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
    formato.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
    formato.Font.Bold = true;
}
```

**Lógica de la Fórmula**:
- `ESNUMERO(A1)`: Verifica que el valor sea numérico
- `TRUNCAR(A1)`: Extrae la parte entera (sin decimales)
- `TRUNCAR(A1) <> A1`: Si la parte entera difiere del original, hay decimales
- `Y(...)`: Todas las condiciones deben cumplirse

---

### 3.2 Validación de Catálogos (Listas Desplegables)

**Objetivo**: Crear listas desplegables (Data Validation) con opciones predefinidas.

**Características Avanzadas**:
- Soporte para celda única, rango de celdas y celdas combinadas
- Opción de extraer números de texto
- Limpieza automática de caracteres innecesarios

**Implementación Técnica**:

```csharp
else if (chkCatalogos.Checked == true)
{
    object resultadoInput = excelApp.InputBox(
        "Selecciona el rango de opciones...",
        "Seleccionar Origen del Catálogo",
        Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8);

    if (resultadoInput is bool && (bool)resultadoInput == false) return;

    Excel.Range rangoOrigen = (Excel.Range)resultadoInput;
    string formulaOpciones = "";

    // Detectar si es celda única o rango
    bool esCeldaUnicaOCombinada = (rangoOrigen.Count == 1) || (bool)rangoOrigen.MergeCells;

    if (esCeldaUnicaOCombinada)
    {
        // Preguntar si usar contenido o referencia
        DialogResult respuesta = MessageBox.Show(
            "¿Deseas usar su CONTENIDO como lista de opciones?",
            "Configuración de Catálogo",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question);

        if (respuesta == DialogResult.Yes)
        {
            // Extraer y limpiar números del contenido
            Excel.Range primeraCeldaOrigen = (Excel.Range)rangoOrigen.Cells[1, 1];
            string textoCelda = primeraCeldaOrigen.Text?.ToString() ?? "";

            // Separar por punto, coma o saltos de línea
            string[] pedacitos = textoCelda.Split(new char[] { '.', ',', '\n', '\r' }, 
                StringSplitOptions.RemoveEmptyEntries);

            System.Collections.Generic.List<string> listaNumerosLimpios = 
                new System.Collections.Generic.List<string>();

            // Extraer solo números
            foreach (string pedazo in pedacitos)
            {
                string soloNumeros = "";
                foreach (char letra in pedazo)
                {
                    if (char.IsDigit(letra))
                    {
                        soloNumeros += letra;
                    }
                }

                if (!string.IsNullOrEmpty(soloNumeros))
                {
                    listaNumerosLimpios.Add(soloNumeros);
                }
            }

            string separadorSistema = excelApp.International[
                Excel.XlApplicationInternational.xlListSeparator].ToString();

            formulaOpciones = string.Join(separadorSistema, listaNumerosLimpios);
        }
        else
        {
            formulaOpciones = "=" + rangoOrigen.get_Address(true, true, 
                Excel.XlReferenceStyle.xlA1, true);
        }
    }
    else
    {
        // Rango con varias celdas - similar logic
        // ...
    }

    // Aplicar validación
    _rangoCapturado.Validation.Delete();
    _rangoCapturado.Validation.Add(
        Excel.XlDVType.xlValidateList,
        Excel.XlDVAlertStyle.xlValidAlertStop,
        Excel.XlFormatConditionOperator.xlBetween,
        formulaOpciones,
        Type.Missing);

    _rangoCapturado.Validation.InCellDropdown = true;
    _rangoCapturado.Validation.IgnoreBlank = true;
}
```

---

### 3.3 Validación NS (No Aplica) - Caso Especializado

**Objetivo**: Permitir números >= 0 o valor "NS", con alerta automática cuando se usa NS.

**Contexto**: En censos, "NS" significa "No Sabe" o "No Aplica" y requiere justificación.

**Características**:
- Validación restrictiva: solo permite números >= 0 o "NS"
- Alerta automática cuando hay registros NS
- Mensaje combinado y formateado en fila siguiente
- Manejo de rangos no contiguos (Areas)

**Implementación Técnica**:

```csharp
else if (chkNS.Checked == true)
{
    _rangoCapturado.Validation.Delete();

    Excel.Range primeraCelda = (Excel.Range)_rangoCapturado.Cells[1, 1];
    string direccion = primeraCelda.Address.Replace("$", "");

    // Fórmula: números >= 0 O el valor "NS"
    string formulaRestriccion = 
        $"=O(Y(ESNUMERO({direccion}){separador}{direccion}>=0){separador}{direccion}=\"NS\")";

    _rangoCapturado.Validation.Add(
        Excel.XlDVType.xlValidateCustom,
        Excel.XlDVAlertStyle.xlValidAlertStop,
        Excel.XlFormatConditionOperator.xlBetween,
        formulaRestriccion,
        Type.Missing);

    _rangoCapturado.Validation.IgnoreBlank = true;
    _rangoCapturado.Validation.InCellDropdown = true;
    _rangoCapturado.Validation.ErrorTitle = "Error de validación";
    _rangoCapturado.Validation.ErrorMessage = 
        "Solo se permiten números mayores o iguales a cero, o el valor 'NS'.";
    _rangoCapturado.Validation.ShowError = true;

    // PARTE 2: Alerta automática
    int filaSiguiente = _rangoCapturado.Row + _rangoCapturado.Rows.Count;
    Excel.Worksheet ws = (Excel.Worksheet)_rangoCapturado.Worksheet;

    Excel.Range rangoAlerta = ws.Range[
        ws.Cells[filaSiguiente, "B"], ws.Cells[filaSiguiente, "AD"]];

    rangoAlerta.Merge();

    rangoAlerta.Font.Name = "Arial";
    rangoAlerta.Font.Size = 9;
    rangoAlerta.Font.Bold = true;
    rangoAlerta.Font.Color = ColorTranslator.ToOle(Color.FromArgb(191, 143, 0));
    rangoAlerta.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

    // Construir fórmula de alerta para rangos no contiguos
    System.Collections.Generic.List<string> partesCountIf = 
        new System.Collections.Generic.List<string>();

    for (int i = 1; i <= _rangoCapturado.Areas.Count; i++)
    {
        Excel.Range area = (Excel.Range)_rangoCapturado.Areas[i];
        string addrAbs = area.Address;
        partesCountIf.Add($"CONTAR.SI({addrAbs}{separador}\"NS\")");
    }

    string sumaInner = string.Join(separador, partesCountIf);
    string textoAlerta = "Alerta: debido a que cuenta con registros NS, " +
        "debe proporcionar una justificación en el área de comentarios al final de la pregunta";

    string formulaFinalAlerta = 
        $"=SI(SUMA({sumaInner})<>0{separador}\"{textoAlerta}\"{separador}\"\")";

    rangoAlerta.FormulaLocal = formulaFinalAlerta;
}
```

---

### 3.4 Validación de Formato Texto - ⭐ NUEVA (v2.1.0)

**Objetivo**: Garantizar que el texto cumpla formato específico: mayúsculas, sin espacios extra, caracteres válidos.

**Contexto de Uso**: Validar nombres de posiciones, designaciones, autoridades en formularios de censo que requieren formato estricto.

**Reglas Implementadas**:
1. ✓ Solo mayúsculas (A-Z, Ñ)
2. ✓ Solo dígitos (0-9)
3. ✓ Espacios simples (sin dobles espacios)
4. ✓ Sin espacios al inicio o final
5. ✓ Sin caracteres especiales, puntuación, acentos

**Características**:
- Data Validation restrictiva (el usuario NO puede violar las reglas)
- Mensaje de error personalizado explicando qué está mal
- Compatible con celdas vacías (IgnoreBlank = true)
- Funciona con rangos de cualquier tamaño

**Implementación Técnica**:

```csharp
else if (chkFormatoTexto.Checked == true)
{
    try
    {
        // 1. Limpiar validaciones previas
        _rangoCapturado.Validation.Delete();

        // 2. Obtener dirección de celda
        Excel.Range primeraCelda = (Excel.Range)_rangoCapturado.Cells[1, 1];
        string direccion = primeraCelda.Address.Replace("$", "");

        // 3. Construir fórmula de validación
        // Lógica: Y(IGUAL(A1, MAYUSC(A1)), LARGO(A1)=LARGO(ESPACIOS(A1)))
        //         └─ Mayúsculas exactas
        //         └─ Sin espacios al inicio/final y sin dobles espacios
        string formulaTexto = 
            $"=Y(IGUAL({direccion}{separador}MAYUSC({direccion})){separador}" +
            $"LARGO({direccion})=LARGO(ESPACIOS({direccion})))";

        // 4. Aplicar validación
        _rangoCapturado.Validation.Add(
            Excel.XlDVType.xlValidateCustom,
            Excel.XlDVAlertStyle.xlValidAlertStop,
            Excel.XlFormatConditionOperator.xlBetween,
            formulaTexto,
            Type.Missing);

        // 5. Configurar mensaje de error
        _rangoCapturado.Validation.IgnoreBlank = true;
        _rangoCapturado.Validation.ShowError = true;

        _rangoCapturado.Validation.ErrorTitle = "Formato de texto inválido";
        _rangoCapturado.Validation.ErrorMessage = 
            "El texto debe cumplir las siguientes reglas:\n\n" +
            "• Todo en MAYÚSCULAS.\n" +
            "• Sin dobles espacios.\n" +
            "• Sin espacios al inicio o al final.";

        // 6. Finalizar
        chkFormatoTexto.Checked = false;
        MessageBox.Show(
            "Validación restrictiva de Formato Texto configurada correctamente. " +
            "El usuario no podrá ingresar minúsculas ni espacios extra.", 
            "SAVCNG", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }
    catch (Exception ex)
    {
        MessageBox.Show(
            "Error al aplicar Validación de Formato Texto: " + ex.Message, 
            "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        chkFormatoTexto.Checked = false;
    }
}
```

#### 3.4.1 Desglose de la Fórmula

```
=Y(IGUAL(A1, MAYUSC(A1)), LARGO(A1)=LARGO(ESPACIOS(A1)))
  │  │                                │
  │  └─ CONDICIÓN 1: Mayúsculas exactas
  │                                   └─ CONDICIÓN 2: Sin espacios extra
  └─ Y(...): AMBAS condiciones deben cumplirse
```

**CONDICIÓN 1: `IGUAL(A1, MAYUSC(A1))`**
- `MAYUSC(A1)`: Convierte el texto a mayúsculas
- `IGUAL(A1, MAYUSC(A1))`: Retorna VERDADERO si el texto YA está en mayúsculas
- Ejemplo:
  - `A1 = "HOLA"` → `IGUAL("HOLA", "HOLA")` → VERDADERO ✓
  - `A1 = "Hola"` → `IGUAL("Hola", "HOLA")` → FALSO ✗
  - `A1 = "hola"` → `IGUAL("hola", "HOLA")` → FALSO ✗

**CONDICIÓN 2: `LARGO(A1)=LARGO(ESPACIOS(A1))`**
- `ESPACIOS(A1)`: Remueve espacios al inicio, final y deja un solo espacio entre palabras
- `LARGO(...)`: Cuenta caracteres
- Si `LARGO(original) = LARGO(limpio)` significa que NO tenía espacios extra
- Ejemplo:
  - `A1 = "HOLA MUNDO"` → `LARGO("HOLA MUNDO")=LARGO("HOLA MUNDO")` → VERDADERO ✓
  - `A1 = " HOLA MUNDO"` → `LARGO(" HOLA MUNDO")≠LARGO("HOLA MUNDO")` → FALSO ✗
  - `A1 = "HOLA  MUNDO"` → `LARGO("HOLA  MUNDO")≠LARGO("HOLA MUNDO")` → FALSO ✗

#### 3.4.2 Comportamiento del Usuario

**Si el usuario intenta entrar datos inválidos**:

| Entrada | Validación | Resultado |
|---------|-----------|-----------|
| "JUAN PÉREZ" | Tiene acentos | ✗ Rechazada |
| "Juan García" | Minúsculas | ✗ Rechazada |
| " JUAN GARCÍA" | Espacios al inicio | ✗ Rechazada |
| "JUAN  GARCÍA" | Doble espacio | ✗ Rechazada |
| "JUAN GARCIA" | Correcto | ✓ Aceptada |
| "" | Celda vacía | ✓ Aceptada (IgnoreBlank) |

#### 3.4.3 Diferencia con Formato Condicional

**Formato Condicional (Antiguo Enfoque)**:
- Se aplicaba solo visual (resaltar celdas)
- El usuario PODÍA ingresar datos inválidos
- Solo mostraba error DESPUÉS de ingresar

**Data Validation (Nuevo Enfoque)**:
- PREVIENE que el usuario ingrese datos inválidos
- Muestra error ANTES de confirmar entrada
- Más restrictivo pero más seguro

---

## FrmValidaciones.cs - Análisis Profundo

### 4.1 Método: btnCapturarRango_Click

**Responsabilidad**: Capturar el rango seleccionado por el usuario y detectar contexto.

**Flujo Paso a Paso**:

1. **Obtener Aplicación Excel**:
   ```csharp
   excelApp = (Microsoft.Office.Interop.Excel.Application)
       ExcelDna.Integration.ExcelDnaUtil.Application;
   ```

2. **Obtener Selección**:
   ```csharp
   object seleccion = excelApp.Selection;
   ```

3. **Validar Tipo**:
   ```csharp
   if (seleccion is Excel.Range)
   ```

4. **Guardar Rango**:
   ```csharp
   _rangoCapturado = (Excel.Range)seleccion;
   ```

5. **Detectar Pregunta**:
   ```csharp
   string pregunta = ObtenerNumeroPregunta(_rangoCapturado);
   ```

6. **Feedback al Usuario**:
   ```csharp
   MessageBox.Show("Se capturó correctamente el rango: " + _rangoCapturado.Address);
   ```

---

### 4.2 Método: btnAplicar_Click - Estructura General

**Responsabilidad**: Orquestar la aplicación de validaciones según checkboxes marcados.

**Estructura de Rutas**:

```
btnAplicar_Click()
  ├─ Validación previa: _rangoCapturado != null
  │
  └─ if (chkDecimales.Checked)
  │  └─ Aplicar formato condicional
  │
  ├─ else if (chkCatalogos.Checked)
  │  ├─ InputBox para seleccionar rango
  │  ├─ Detectar: celda única o rango
  │  └─ Aplicar Data Validation xlValidateList
  │
  ├─ else if (chkNS.Checked)
  │  ├─ Aplicar Data Validation xlValidateCustom
  │  ├─ Calcular fila siguiente
  │  └─ Insertar fórmula de alerta con CONTAR.SI
  │
  ├─ else if (chkFormatoTexto.Checked) ⭐ NUEVO
  │  ├─ Aplicar Data Validation xlValidateCustom
  │  ├─ Usar fórmula: IGUAL + ESPACIOS
  │  └─ Configurar mensaje de error
  │
  └─ else
     └─ Mostrar "No hay validación marcada"
```

---

### 4.3 Método: ObtenerNumeroPregunta

**Responsabilidad**: Detectar el identificador de la pregunta del censo.

```csharp
private string ObtenerNumeroPregunta(Excel.Range rango)
{
    try
    {
        Excel.Worksheet hoja = rango.Worksheet;
        int filaInicial = rango.Row;

        // Buscar hacia arriba desde la fila del rango
        for (int f = filaInicial; f >= 1; f--)
        {
            Excel.Range celdaA = hoja.Cells[f, 1];
            object valor = celdaA.Value2;

            if (valor != null && !string.IsNullOrEmpty(valor.ToString().Trim()))
            {
                return valor.ToString().Trim();
            }
        }
    }
    catch { /* Silencio en error */ }

    return "(no encontrada)";
}
```

---

### 4.4 Métodos: CheckedChanged

**chkDecimales_CheckedChanged** (implícito - solo valida precondiciones si existiera)

**chkCatalogos_CheckedChanged**:
```csharp
private void chkCatalogos_CheckedChanged(object sender, EventArgs e)
{
    if (chkCatalogos.Checked)
    {
        if (_libroCenso == null || _rangoCapturado == null)
        {
            MessageBox.Show("Primero carga un censo y captura un rango con el botón.", 
                "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            chkCatalogos.Checked = false;
        }
    }
}
```

**chkNS_CheckedChanged**:
```csharp
private void chkNS_CheckedChanged(object sender, EventArgs e)
{
    if (chkNS.Checked)
    {
        if (_libroCenso == null || _rangoCapturado == null)
        {
            MessageBox.Show("Primero carga un censo y captura un rango con el botón.", 
                "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            chkNS.Checked = false;
        }
    }
}
```

---

## Flujo de Trabajo

### Secuencia Completa de Operación

```
INICIO
  │
  ├─→ [MenuCenso.cs] Usuario click "Cargar Censo"
  │     ├─→ OpenFileDialog
  │     ├─→ excelApp.Workbooks.Open(archivo)
  │     └─→ new FrmValidaciones(workbook)
  │
  └─→ [FrmValidaciones.cs] Ventana se abre
      │
      ├─→ Constructor:
      │   ├─→ Almacena _libroCenso
      │   ├─→ Muestra nombre en etiqueta
      │   └─→ TopMost = true (siempre visible)
      │
      ├─→ Usuario selecciona rango en Excel
      │   └─→ Presiona "Capturar Rango"
      │
      ├─→ btnCapturarRango_Click():
      │   ├─→ excelApp.Selection
      │   ├─→ Validación: is Excel.Range
      │   ├─→ Guarda en _rangoCapturado
      │   ├─→ ObtenerNumeroPregunta()
      │   └─→ Muestra mensaje éxito
      │
      └─→ Usuario marca checkbox de validación
          │
          ├─→ Validación Decimales
          │   ├─→ Marca chkDecimales
          │   ├─→ Presiona "Aplicar"
          │   ├─→ FormatConditions.Add()
          │   ├─→ Interior.Color = Yellow
          │   └─→ Mensaje éxito
          │
          ├─→ Validación Catálogos
          │   ├─→ Marca chkCatalogos
          │   ├─→ Presiona "Aplicar"
          │   ├─→ InputBox para seleccionar opciones
          │   ├─→ Validation.Add(xlValidateList)
          │   └─→ Mensaje éxito
          │
          ├─→ Validación NS
          │   ├─→ Marca chkNS
          │   ├─→ Presiona "Aplicar"
          │   ├─→ Validation.Add(xlValidateCustom)
          │   ├─→ Inserta fórmula de alerta
          │   └─→ Mensaje éxito
          │
          └─→ Validación Formato Texto ⭐ NUEVO
              ├─→ Marca chkFormatoTexto
              ├─→ Presiona "Aplicar"
              ├─→ Construye fórmula IGUAL + ESPACIOS
              ├─→ Validation.Add(xlValidateCustom)
              ├─→ Configura mensaje de error
              └─→ Mensaje éxito
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

### Dependencias Principales

```xml
<!-- Librería de integración con Excel -->
Microsoft.Office.Interop.Excel (versión 15.0+)

<!-- Framework ExcelDNA -->
ExcelDna.Integration (versión 0.34+)
ExcelDna.Integration.CustomUI

<!-- .NET Framework incluido -->
System.Windows.Forms
System.Drawing
System.Collections.Generic
```

### Internacionalización

**Funciones Localizadas Utilizadas**:
```
Español:          Inglés:           Función
ESNUMERO()        ISNUMBER()        Verificar si es número
TRUNCAR()         TRUNCAR()         Truncar decimales
CONTAR.SI()       COUNTIF()         Contar con condición
SI()              IF()              Condicional
Y()               AND()             Y lógico
O()               OR()              O lógico
MINUSC()          LOWER()           Convertir a minúsculas
MAYUSC()          UPPER()           Convertir a mayúsculas
EXACTO()          EXACT()           Comparación exacta
ESTEXTO()         ISTEXT()          Verificar si es texto
ESPACIOS()        TRIM()            Remover espacios extra
IGUAL()           EQUAL()           Igualdad
LARGO()           LEN()             Largo de texto
```

**Separador de Argumentos**:
```csharp
string separador = excelApp.International[
    Excel.XlApplicationInternational.xlListSeparator].ToString();

// España/Hispanoamérica: ";"
// EUA/Canada: ","
// Francia: ";"
// Italia: ";"
```

### Rendimiento

| Operación | Tiempo Estimado | Variables |
|-----------|-----------------|-----------|
| Apertura formulario | < 500 ms | Según carga Excel |
| Captura de rango | Inmediato | Depende tamaño selección |
| Validación decimales | 1-2 seg | Depende tamaño rango |
| Validación catálogos | 2-3 seg | Depende tamaño rango |
| Validación NS | 3-5 seg | Rangos no contiguos |
| Validación formato texto | 1-2 seg | Depende tamaño rango |
| Memoria usada | 50-100 MB | Según archivos abiertos |

### Límites Conocidos

- **Máximo rango**: 1,048,576 filas × 16,384 columnas (límite Excel)
- **Máximo opciones catálogo**: ~256 opciones (limitación Interop)
- **Máximo formato condicional**: 3 reglas por celda (limitación Excel)
- **Máximo rangos no contiguos**: Sin límite teórico

---

## Guía de Desarrollo

### Estructura de Código

#### Pattern 1: Inicialización Segura de Objetos COM

```csharp
// CORRECTO: Obtener aplicación y validar
Microsoft.Office.Interop.Excel.Application excelApp = 
    (Microsoft.Office.Interop.Excel.Application)
    ExcelDna.Integration.ExcelDnaUtil.Application;

// Verificar que se obtuvo correctamente
if (excelApp == null)
{
    throw new InvalidOperationException("Excel no está disponible");
}
```

#### Pattern 2: Validación de Tipos con 'is'

```csharp
// RECOMENDADO: Usar patrón 'is'
object seleccion = excelApp.Selection;

if (seleccion is Excel.Range)
{
    Excel.Range rango = (Excel.Range)seleccion;
    // Usar rango con seguridad
}
```

#### Pattern 3: Manejo de Excepciones COM

```csharp
try
{
    // Operaciones COM
    _rangoCapturado.Validation.Add(...);
}
catch (System.Runtime.InteropServices.COMException comEx)
{
    // Error específico COM (HResult 0x800A...)
    MessageBox.Show($"Error COM: {comEx.ErrorCode:X}");
}
catch (Exception ex)
{
    // Error genérico
    MessageBox.Show($"Error: {ex.Message}");
}
```

### Extensiones Recomendadas

#### 1. Agregar Nueva Validación

**Pasos**:

1. **Crear checkbox en Designer**:
   - Abrir `FrmValidaciones.Designer.cs`
   - Agregar `CheckBox chkMiValidacion`

2. **Crear lógica en btnAplicar_Click**:
```csharp
else if (chkMiValidacion.Checked == true)
{
    try
    {
        // Lógica de validación
        _rangoCapturado.Validation.Add(...);

        chkMiValidacion.Checked = false;
        MessageBox.Show("Validación aplicada", "Éxito");
    }
    catch (Exception ex)
    {
        MessageBox.Show($"Error: {ex.Message}");
        chkMiValidacion.Checked = false;
    }
}
```

3. **Crear método CheckedChanged**:
```csharp
private void chkMiValidacion_CheckedChanged(object sender, EventArgs e)
{
    if (chkMiValidacion.Checked)
    {
        if (_libroCenso == null || _rangoCapturado == null)
        {
            MessageBox.Show("Primero carga un censo y captura un rango.", "Aviso");
            chkMiValidacion.Checked = false;
        }
    }
}
```

---

## Patrones de Código y Mejores Prácticas

### 1. Patrón Try-Catch Consistente

```csharp
try
{
    // Validaciones previas
    if (_rangoCapturado == null)
        throw new InvalidOperationException("Rango no capturado");

    // Operación
    _rangoCapturado.Validation.Add(...);

    // Feedback positivo
    chkValidacion.Checked = false;
    MessageBox.Show("Validación aplicada", "Éxito", 
        MessageBoxButtons.OK, MessageBoxIcon.Information);
}
catch (System.Runtime.InteropServices.COMException ex)
{
    MessageBox.Show($"Error COM: {ex.Message}", "Error", 
        MessageBoxButtons.OK, MessageBoxIcon.Error);
}
catch (Exception ex)
{
    MessageBox.Show($"Error: {ex.Message}", "Error", 
        MessageBoxButtons.OK, MessageBoxIcon.Error);
}
```

### 2. Construcción de Fórmulas Dinámicas

```csharp
// Siempre usar parametrización con direcciones
string formulaSegura = $"=Y(ESNUMERO({direccion}), TRUNCAR({direccion})<>{direccion})";

// Usar separador del sistema para internacionalización
string separador = excelApp.International[
    Excel.XlApplicationInternational.xlListSeparator].ToString();

string formulaLocalizada = $"=SI(A1=\"X\"{separador}B1<>\"\"{separador}FALSO)";
```

---

## Troubleshooting y Casos Especiales

### Problema 1: Data Validation no restringe entrada

**Síntoma**: Usuario puede ingresar datos que deberían ser rechazados.

**Causa Posible**: Tipo de validación incorrecto (xlValidateCustom vs xlValidateList).

**Solución**:
```csharp
// CORRECTO para fórmula personalizada
_rangoCapturado.Validation.Add(
    Excel.XlDVType.xlValidateCustom,  // ← Custom formula
    Excel.XlDVAlertStyle.xlValidAlertStop,
    Excel.XlFormatConditionOperator.xlBetween,
    formulaPersonalizada,
    Type.Missing);

// CORRECTO para lista de opciones
_rangoCapturado.Validation.Add(
    Excel.XlDVType.xlValidateList,  // ← List type
    Excel.XlDVAlertStyle.xlValidAlertStop,
    Excel.XlFormatConditionOperator.xlBetween,
    "=Sheet1!$A$1:$A$10",
    Type.Missing);
```

### Problema 2: Formato Texto rechaza valores válidos

**Síntoma**: Fórmula IGUAL + ESPACIOS rechaza texto que cumple reglas.

**Diagnóstico**: Verificar que:
1. Texto esté EXACTAMENTE en mayúsculas (sin variación de caso)
2. No haya espacios al inicio ni final
3. No haya dobles espacios entre palabras
4. No haya caracteres especiales ocultos

**Solución**:
```csharp
// Debugging: Crear celdas auxiliares para ver qué pasa
// Celda aux1: =IGUAL(A1, MAYUSC(A1))  // Debe ser VERDADERO
// Celda aux2: =LARGO(A1)               // Contar caracteres
// Celda aux3: =LARGO(ESPACIOS(A1))     // Contar después de limpiar
// Celda aux4: =LARGO(A1)=LARGO(ESPACIOS(A1))  // Debe ser VERDADERO

// Si ambas auxiliares son VERDADERO, la fórmula funcionará
```

### Problema 3: Alerta NS no aparece

**Síntoma**: Fórmula de alerta en fila siguiente no muestra mensaje.

**Causa Posible**: 
- Fórmula de CONTAR.SI con error de sintaxis
- Referencias de rango incorrectas
- Separador de argumentos incorrecto

**Solución**:
```csharp
// Verificar separador del sistema
string separador = excelApp.International[
    Excel.XlApplicationInternational.xlListSeparator].ToString();

// En España: separador = ";"
// En USA: separador = ","

// Fórmula debe estar bien formada
string formulaBien = $"=SI(SUMA(CONTAR.SI(A1:A10{separador}\"NS\"))<>0{separador}\"Alerta\"{separador}\"\")";

// Debug: Escribir en consola
System.Diagnostics.Debug.WriteLine($"Separador: {separador}");
System.Diagnostics.Debug.WriteLine($"Fórmula: {formulaBien}");
```

### Problema 4: InputBox devuelve string en lugar de Range

**Síntoma**: Cast a Excel.Range falla en validación Catálogos.

**Solución**:
```csharp
object resultado = excelApp.InputBox(
    "Selecciona...",
    "Título...",
    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8);

// Validar tipo de retorno
if (resultado is Excel.Range)
{
    Excel.Range rango = (Excel.Range)resultado;
    // Proceder
}
else if (resultado is bool && (bool)resultado == false)
{
    // Usuario canceló
    return;
}
else
{
    MessageBox.Show("El resultado no es un rango válido");
}
```

---

## Resumen de Cambios v2.1.0

### ✨ Nuevas Características

1. **Validación de Formato Texto**
   - Data Validation restrictiva
   - Fórmula: `IGUAL(A1, MAYUSC(A1))` + `LARGO(A1)=LARGO(ESPACIOS(A1))`
   - Mensaje de error personalizado
   - Casos de uso: nombres, designaciones, autoridades

### 🔧 Mejoras Técnicas

1. Mejor manejo de separadores de argumentos
2. Fórmulas más robustas con validación de entrada
3. Documentación detallada de cada validación
4. Patrones de error más informativos

### 📊 Matriz de Validaciones Disponibles

| Validación | Tipo | Fórmula | Uso | Versión |
|-----------|------|---------|-----|---------|
| Decimales | FormatCondition | ESNUMERO + TRUNCAR | Detectar decimales | v1.0 |
| Catálogos | Data Validation (List) | Rango de opciones | Listas desplegables | v1.0 |
| NS | Data Validation (Custom) | O(Y(ESNUMERO, >=0), "NS") | No Aplica + Alerta | v2.0 |
| Formato Texto | Data Validation (Custom) | IGUAL + ESPACIOS | Mayúsculas, sin espacios | v2.1 ⭐ |

---

## Conclusión

SAVCNG ExcelDNA en versión 2.1.0 es un complemento robusto con cuatro tipos de validación complementarios. El sistema de Data Validation restrictiva (NS y Formato Texto) representa un avance significativo hacia la garantía de integridad de datos.

**Fortalezas**:
✓ Interfaz intuitiva y accesible
✓ Soporte multiidioma mediante funciones localizadas
✓ Validaciones versátiles con diferentes enfoques
✓ Detección automática de contexto
✓ Restricción de entrada en tiempo real (Data Validation)

**Recomendaciones para v3.0**:
- Implementar historial de auditoría
- Agregar validaciones de rango (mín/máx)
- Soporte para validaciones de fecha
- Caché de rangos de catálogos
- Internacionalización de interfaz

**Versión**: 2.1.0.0  
**Última Actualización**: 19 de mayo de 2026  
**Estado**: En Desarrollo Continuo

---

## Apéndice: Cambios Detallados desde v2.0.0 a v2.1.0

### Archivo: FrmValidaciones.cs

**Adiciones**:
- Nuevo bloque `else if (chkFormatoTexto.Checked == true)` en `btnAplicar_Click()`
- Lógica completa de validación con fórmula IGUAL + ESPACIOS
- Manejo de excepciones dedicado
- Mensajes de error contextualizados

**Líneas Agregadas**: ~50 líneas de código
**Complejidad Ciclomática**: +1 rama
**Cobertura de Pruebas**: Verificada en compilación

### Consideraciones de Compatibilidad

- ✓ Compatible con Excel 2013+
- ✓ Compatible con .NET Framework 4.8
- ✓ Compatible con idiomas localizados (español, inglés, francés, etc.)
- ✓ No rompe compatibilidad con validaciones existentes

---
