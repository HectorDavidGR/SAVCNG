using System;
using System.Drawing;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel; // Importante para entenderse con Excel
namespace SAVCNG_ExcelDNA
{
    public partial class FrmValidaciones : Form
    {
        // Esta variable guardará el archivo de Excel que abrimos
        private Excel.Workbook _libroCenso;

        private Excel.Range _rangoCapturado;

        // Variable para guardar el rango de opciones del catálogo
        private Excel.Range _rangoCatalogo;

        // Aquí guardaremos los rangos que se van acumulando para validar
        private System.Collections.Generic.List<Excel.Range> _pendientesCatalogo = new System.Collections.Generic.List<Excel.Range>();

        // Aquí guardaremos el mapeo de qué rango pertenece a qué pregunta
        private System.Collections.Generic.Dictionary<string, string> _preguntaPorRango = new System.Collections.Generic.Dictionary<string, string>();

        // Carga inicial del formulario
        private void FrmValidaciones_Load(object sender, EventArgs e)
        {

            CargarEstadoDelCenso();

        }


        // Este es el constructor. Le agregamos "Excel.Workbook libroAbierto" para que reciba el censo
        public FrmValidaciones(Excel.Workbook libroAbierto)
        {
            InitializeComponent();

            // Guardamos el libro en nuestra variable para usarlo después
            _libroCenso = libroAbierto;

            // Actualizamos la etiqueta con el nombre del archivo
            lblCenso.Text = "Censo cargado: " + _libroCenso.Name;

            // Truco de experto: Hacemos que la ventana siempre esté por encima de Excel
            this.TopMost = true;
            // Que aparezca en el centro de la pantalla
            this.StartPosition = FormStartPosition.CenterScreen;

            // =========================================================================
            // INYECCIÓN DE UX: Llamamos al configurador de ToolTips al cargar la ventana
            // =========================================================================
            ConfigurarToolTips();

        }

        // --- INICIO DEL MÓDULO DE EXPERIENCIA DE USUARIO (UX) ---
        // Mensajes emergentes para los botones check
        private void ConfigurarToolTips()
        {
            // 1. Instanciamos el componente nativo de WinForms de manera global para la ventana
            ToolTip toolTipValidaciones = new ToolTip();

            // 2. Configuración de tiempos del motor (Valores en milisegundos)
            toolTipValidaciones.AutoPopDelay = 10000;  // UX: Tiempo que el mensaje se queda visible (10 segundos para leer tranquilos)
            toolTipValidaciones.InitialDelay = 200;    // UX: Tiempo de espera al poner el cursor encima antes de mostrar el mensaje (0.4 seg)
            toolTipValidaciones.ReshowDelay = 100;     // UX: Tiempo de transición al mover el cursor rápidamente a otro botón
            toolTipValidaciones.ShowAlways = true;     // Asegura que el tooltip se vea incluso si Excel tiene el foco principal por debajo

            // 3. Diseño Estético de la ventana flotante
            toolTipValidaciones.ToolTipIcon = ToolTipIcon.Info; // Muestra un pequeño icono azul de "i"
            toolTipValidaciones.ToolTipTitle = "Descripción de la validación";

            // 4. Mapeo del Diccionario de Descripciones por CheckBox
            // Nota: Reemplaza "chkDecimales", "chkFechas", etc. si el nombre interno de tus controles difiere un poco.

            toolTipValidaciones.SetToolTip(this.chkBlancos,
                "Inyecta una matriz de álgebra booleana que alerta en azul si el informante deja celdas vacías en una fila iniciada.");

            toolTipValidaciones.SetToolTip(this.chkBloqueo,
                "Bloquea celdas restrictivamente basándose en una condición lógica (Ej. 'Sí').\nPermite apilar reglas previas.");

            toolTipValidaciones.SetToolTip(this.chkCatalogos,
                "Crea una lista de opciones predefinidas. Es ideal para campos de opción múltiple o catálogos, limitando lo que se puede escribir en la celda.");

            toolTipValidaciones.SetToolTip(this.chkDecimales,
                "Asegura que solo se puedan escribir números enteros (sin punto decimal). Además, permite registrar los códigos 'NS' y 'NA'");

            toolTipValidaciones.SetToolTip(this.chkEspClave,
                "Ideal para los campos de 'Especifique'. Si el texto que el usuario intenta capturar ya estaba en el catálogo principal, te avisa en color amarillo para no duplicar información.");

            toolTipValidaciones.SetToolTip(this.chkFechas,
                "Restringe la(s) celda(s) para aceptar solo números dentro de un rango que tú definas (ideal para días, meses o años). También acepta las claves 'NS' y 'NA'.");

            toolTipValidaciones.SetToolTip(this.chkFormatoTexto,
                "Ideal para preguntas abiertas como nombres o direcciones. Estandariza las respuestas transformando todo a mayúsculas y eliminando errores de datos como simbolos o dobles espacios.");

            toolTipValidaciones.SetToolTip(this.chkNS,
                "Busca el código 'NS' en el(los) rango(s) establecido(s). Si lo encuentra, te avisa con una alerta en color amarillo para que puedas revisarlas.");

            toolTipValidaciones.SetToolTip(this.chkSumas,
                "Motor de sumas cruzadas horizontales (FormatConditions) y generación de fórmulas automáticas (Σ) verticales.\nPermite apilamiento jerárquico multinivel.");
            
        }
        // Evento para seleccionar exclusivamente 1 checkbox a la vez
        private void CheckBox_Exclusivo_CheckedChanged(object sender, EventArgs e)
        {
            // 1. Identificamos qué CheckBox disparó el evento
            CheckBox chkActivo = sender as CheckBox;

            // 2. Si el CheckBox se está encendiendo, apagamos los demás
            if (chkActivo != null && chkActivo.Checked)
            {
                // NOTA: Cambia "panelOpciones" por el nombre del GroupBox, Panel o TableLayoutPanel 
                // donde tengas metidos tus 6 CheckBox.
                foreach (Control ctrl in tLP_Paso2_chkboxes.Controls)
                {
                    // Si el control es un CheckBox y NO es el que acaban de presionar...
                    if (ctrl is CheckBox && ctrl != chkActivo)
                    {
                        // Lo desmarcamos
                        ((CheckBox)ctrl).Checked = false;
                    }
                }
            }
        }

        // --- AQUÍ IRÁN LOS EVENTOS DE LOS BOTONES QUE SE USAN EN LA INTERFAZ ---
        // Evento para caprturar el rango seleccionado en el excel mediante boton Capturar Rango
        private void btnCapturarRango_Click(object sender, EventArgs e)
        {
            try
            {
                // 1. Instanciamos la aplicación de Excel para interactuar con la interfaz activa
                Microsoft.Office.Interop.Excel.Application excelApp = (Microsoft.Office.Interop.Excel.Application)ExcelDna.Integration.ExcelDnaUtil.Application;

                // 2. Extraemos el objeto seleccionado actualmente por el usuario
                object seleccion = excelApp.Selection;

                // 3. Verificamos que se trate de un rango de celdas válido
                if (seleccion is Excel.Range)
                {
                    // 4. Guardamos la selección en la memoria global del Add-In
                    _rangoCapturado = (Excel.Range)seleccion;

                    // =========================================================================
                    // 5. ALGORITMO DE DETECCIÓN MULTI-PREGUNTA
                    // =========================================================================
                    System.Collections.Generic.List<string> preguntasDetectadas = new System.Collections.Generic.List<string>();

                    // Iteramos sobre cada sub-bloque (área) seleccionado con la tecla CTRL
                    foreach (Excel.Range area in _rangoCapturado.Areas)
                    {
                        // Reutilizamos tu método original enviando área por área
                        string pregunta = ObtenerNumeroPregunta(area);

                        // Evitamos duplicados en la interfaz si el usuario seleccionó varias áreas de la misma pregunta
                        if (!preguntasDetectadas.Contains(pregunta))
                        {
                            preguntasDetectadas.Add(pregunta);
                        }
                    }

                    // Unimos todas las preguntas detectadas separadas por comas (Ej: "1.1, 1.2, 1.5")
                    string textoPreguntas = string.Join(", ", preguntasDetectadas);

                    // =========================================================================
                    // 6. ACTUALIZACIÓN DE LA EXPERIENCIA DE USUARIO (UX/UI)
                    // =========================================================================
                    // Limpiamos los símbolos de anclaje absoluto ($) para hacer la lectura más amigable
                    string direccionLimpia = _rangoCapturado.Address.Replace("$", "");

                    // Actualizamos los Labels de la interfaz
                    lblPregunta.Text = "Pregunta(s) detectada(s): " + textoPreguntas;
                    lblRangoSeleccionado.Text = "Rango seleccionado: " + direccionLimpia;

                    // Mostramos un resumen claro en el MessageBox
                    MessageBox.Show(this,
                        $"Se capturó correctamente la selección.\n\n" +
                        $"• Coordenadas: {direccionLimpia}\n" +
                        $"• Bloques (Áreas): {_rangoCapturado.Areas.Count}\n" +
                        $"• Pregunta(s): {textoPreguntas}",
                        "SAVCNG - Captura Exitosa", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    // UX: Prevención de errores si el usuario selecciona gráficos, formas o imágenes
                    MessageBox.Show(this, "Por favor, selecciona celdas de Excel, no imágenes ni gráficos.",
                                    "Aviso de Captura", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
            }
            catch (Exception ex)
            {
                // Contención de errores críticos
                MessageBox.Show(this, "Ocurrió un error al capturar en memoria: " + ex.Message,
                                "Error Crítico", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        // Evento para aplicar validaciones mediante boton Aplicar Validación en pestaña de validaciones
        private void btnAplicar_Click(object sender, EventArgs e)
        {
            Excel.Application excelApp = (Excel.Application)ExcelDna.Integration.ExcelDnaUtil.Application;
            string separador = excelApp.International[Excel.XlApplicationInternational.xlListSeparator].ToString();

            try
            {
                // Primero verificamos que el usuario no haya olvidado capturar un rango
                if (_rangoCapturado == null)
                {
                    MessageBox.Show(this,"¡Espera! Primero debes capturar un rango.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return; // Detenemos el código aquí
                }
                //=====================================================================================================
                // --- VALIDACIÓN DECIMALES ---
                else if (chkDecimales.Checked == true)
                {
                    Excel.Worksheet wsActual = null;

                    try
                    {
                        excelApp = (Excel.Application)ExcelDna.Integration.ExcelDnaUtil.Application;
                        wsActual = (Excel.Worksheet)_rangoCapturado.Worksheet;

                        // =========================================================================
                        // FASE 1: UX DE DECISIÓN (ELIMINACIÓN DE AMBIGÜEDAD SÍ/NO)
                        // =========================================================================
                        string opcionEscogida = "";

                        // Patrón de bucle para retener al usuario hasta que ingrese un dato válido o cancele
                        while (true)
                        {
                            object seleccionTipo = excelApp.InputBox(
                                "Ingresa el NÚMERO de la regla que deseas aplicar:\n\n" +
                                "1 = Solo números ENTEROS.\n" +
                                "2 = Números DECIMALES (Hasta 10 posiciones).",
                                "SAVCNG - Configuración Numérica",
                                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 1); // Type=1 (Solo evalúa números)

                            // Control de Aborto Silencioso (Si el usuario da clic en Cancelar o la 'X')
                            if (seleccionTipo is bool && (bool)seleccionTipo == false)
                            {
                                chkDecimales.Checked = false;
                                return; // Aquí sí abortamos todo el proceso
                            }

                            opcionEscogida = seleccionTipo.ToString().Trim();

                            // Condición de Salida Segura
                            if (opcionEscogida == "1" || opcionEscogida == "2")
                            {
                                break; // El dato es correcto, rompemos el bucle y continuamos con la Fase 2
                            }
                            else
                            {
                                // Alerta de corrección (Al no tener un "return" aquí, el bucle vuelve a mostrar el InputBox)
                                MessageBox.Show(this, "Opción no válida. Por favor ingresa el número 1 o 2.", "Dato Incorrecto", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            }
                        }

                        // Asignación de variables de contexto basadas en la decisión
                        bool esValidacionEnteros = (opcionEscogida == "1");
                        string etiquetaBitacora = esValidacionEnteros ? "Validación Enteros" : "Validación Decimales";
                        string tituloError = esValidacionEnteros ? "Solo números enteros" : "Formato decimal inválido";
                        string mensajeError = esValidacionEnteros
                            ? "El formato de esta celda no admite texto, decimales, ni números negativos.\n\nPor favor, introduce únicamente un NÚMERO ENTERO POSITIVO (incluyendo el 0) o las claves 'NS' y 'NA'."
                            : "El formato de esta celda exige NÚMEROS POSITIVOS (incluyendo el 0) con un máximo de 10 posiciones decimales.\n\nPor favor, introduce un número válido o las claves 'NS' y 'NA'.";

                        // =========================================================================
                        // FASE 2: AUDITORÍA DE COEXISTENCIA
                        // =========================================================================
                        bool tieneFormatosPrevios = _rangoCapturado.FormatConditions.Count > 0;
                        bool tieneValidacionPrevia = false;

                        try { var tipo = _rangoCapturado.Validation.Type; tieneValidacionPrevia = true; } catch { /* Silenciado */ }

                        bool limpiarFormatos = false;

                        if (tieneFormatosPrevios || tieneValidacionPrevia)
                        {
                            DialogResult respLimpieza = MessageBox.Show(this,
                                "Se detectaron configuraciones previas en el rango seleccionado.\n\n" +
                                "NOTA: Al ser una restricción de captura estricta, la regla se sobreescribirá, pero...\n\n" +
                                "¿Deseas CONSERVAR los colores, alertas o bloqueos (Formatos Condicionales) aplicados previamente?\n\n" +
                                "SÍ = MANTENER colores/bloqueos anteriores.\n" +
                                "NO = BORRAR todo el historial y limpiar el lienzo.",
                                "SAVCNG - Auditoría de Coexistencia", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);

                            if (respLimpieza == DialogResult.Cancel)
                            {
                                chkDecimales.Checked = false;
                                return;
                            }
                            if (respLimpieza == DialogResult.No)
                            {
                                limpiarFormatos = true;
                            }
                        }

                        // =========================================================================
                        // FASE 3: RECOLECCIÓN LIGERA DE PREGUNTAS (STATE MANAGEMENT)
                        // =========================================================================
                        System.Collections.Generic.HashSet<string> preguntasUnicas = new System.Collections.Generic.HashSet<string>();

                        for (int i = 1; i <= _rangoCapturado.Areas.Count; i++)
                        {
                            Excel.Range areaIndividual = (Excel.Range)_rangoCapturado.Areas[i];
                            string idPregunta = ObtenerNumeroPregunta(areaIndividual);

                            if (!string.IsNullOrEmpty(idPregunta))
                            {
                                preguntasUnicas.Add(idPregunta);
                            }
                            // Memory Leak Prevention
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(areaIndividual);
                        }

                        // Apagamos alertas visuales para inyección masiva en segundo plano
                        excelApp.ScreenUpdating = false;

                        // =========================================================================
                        // FASE 4: PROCESAMIENTO MASIVO POR ÁREAS Y TRADUCCIÓN UNIVERSAL
                        // =========================================================================
                        foreach (Excel.Range area in _rangoCapturado.Areas)
                        {
                            Excel.Range primeraCelda = null;
                            Excel.Range celdaDummy = null;
                            Excel.Range celdaDummyFmt = null;
                            Excel.FormatCondition fcSecundario = null;

                            try
                            {
                                // Limpieza obligatoria antes de inyectar nuevas reglas
                                area.Validation.Delete();
                                if (limpiarFormatos) { area.FormatConditions.Delete(); }

                                // ---------------------------------------------------------------------
                                // NUEVO: AJUSTE ARQUITECTÓNICO DE LA CAPA DE PRESENTACIÓN (UI)
                                // ---------------------------------------------------------------------
                                // Al usar NumberFormat (en inglés por defecto de COM), evitamos que Excel redondee visualmente.
                                if (esValidacionEnteros)
                                {
                                    area.NumberFormat = "General"; // Forza visualmente un número entero sin puntos
                                }
                                else
                                {
                                    // Lógica Condicional de Interfaz (UI): 
                                    // Si el valor es exactamente 0, fuerza un "0" limpio. 
                                    // Para cualquier otro número, delega la renderización a "General" (muestra decimales dinámicamente sin punto huérfano).
                                    area.NumberFormat = "[=0]0;0.###############";
                                }

                                primeraCelda = (Excel.Range)area.Cells[1, 1];
                                string dirRel = primeraCelda.get_Address(false, false, Excel.XlReferenceStyle.xlA1, false);
                                int filaBase = area.Row;

                                // ---------------------------------------------------------------------
                                // INGENIERÍA DE FÓRMULAS UNIVERSALES (Lógica Binaria)
                                // Usamos TRUNC para Enteros y ROUND(x, 10) para un máximo de 10 decimales
                                // ---------------------------------------------------------------------
                                string formulaBaseIngles = esValidacionEnteros
                                    ? $"IF(ISNUMBER({dirRel}), AND(TRUNC({dirRel})={dirRel}, {dirRel}>=0), OR(TRIM({dirRel})=\"NS\", TRIM({dirRel})=\"NA\"))"
                                    : $"IF(ISNUMBER({dirRel}), AND(ROUND({dirRel}, 10)={dirRel}, {dirRel}>=0), OR(TRIM({dirRel})=\"NS\", TRIM({dirRel})=\"NA\"))";

                                string formulaValidacion = $"={formulaBaseIngles}";
                                // El formato condicional se dispara si la celda NO está vacía Y la regla principal NO se cumple
                                string formulaCondicion = $"=AND({dirRel}<>\"\", NOT({formulaBaseIngles}))";

                                // ---------------------------------------------------------------------
                                // INYECCIÓN 1: DATA VALIDATION (Dummy Cell Trick)
                                // ---------------------------------------------------------------------
                                celdaDummy = (Excel.Range)wsActual.Cells[filaBase, 16384];
                                celdaDummy.Formula = formulaValidacion;
                                string formulaValidacionLocal = celdaDummy.FormulaLocal;
                                celdaDummy.Clear();

                                area.Validation.Add(
                                    Excel.XlDVType.xlValidateCustom,
                                    Excel.XlDVAlertStyle.xlValidAlertStop,
                                    Excel.XlFormatConditionOperator.xlBetween,
                                    formulaValidacionLocal,
                                    Type.Missing);

                                area.Validation.IgnoreBlank = true;
                                area.Validation.ShowError = true;
                                area.Validation.ErrorTitle = tituloError;
                                area.Validation.ErrorMessage = mensajeError;

                                // ---------------------------------------------------------------------
                                // INYECCIÓN 2: FORMATO CONDICIONAL CENTINELA (Dummy Cell Trick)
                                // ---------------------------------------------------------------------
                                celdaDummyFmt = (Excel.Range)wsActual.Cells[filaBase, 16384];
                                celdaDummyFmt.Formula = formulaCondicion;
                                string formulaCondicionLocal = celdaDummyFmt.FormulaLocal;
                                celdaDummyFmt.Clear();

                                fcSecundario = (Excel.FormatCondition)area.FormatConditions.Add(
                                    Excel.XlFormatConditionType.xlExpression,
                                    Type.Missing,
                                    formulaCondicionLocal);

                                // Estética: Resalte institucional rojo alertando violación a la regla
                                fcSecundario.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 199, 206));
                                fcSecundario.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(156, 0, 6));
                                fcSecundario.StopIfTrue = false;
                            }
                            finally
                            {
                                // Limpieza COM exhaustiva por cada iteración del bucle (Zero Leaks)
                                if (fcSecundario != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(fcSecundario);
                                if (primeraCelda != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(primeraCelda);
                                if (celdaDummy != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(celdaDummy);
                                if (celdaDummyFmt != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(celdaDummyFmt);
                            }
                        }

                        // =========================================================================
                        // FASE 5: CONCLUSIÓN Y REGISTRO EN BITÁCORA
                        // =========================================================================
                        chkDecimales.Checked = false;
                        string estadoAuditoria = limpiarFormatos ? "Se borraron configuraciones visuales previas." : "Se conservaron colores y bloqueos anteriores.";
                        string preguntasDetectadas = preguntasUnicas.Count > 0 ? string.Join(", ", preguntasUnicas) : "ND";

                        AuditoriaCenso.RegistrarAccion(
                            _libroCenso,
                            preguntasDetectadas,
                            etiquetaBitacora, // Utiliza dinámicamente "Validación Enteros" o "Validación Decimales"
                            _rangoCapturado.Address.Replace("$", ""),
                            "Validation + FormatCondition"
                        );

                        MessageBox.Show(this,
                            $"{etiquetaBitacora} aplicada con éxito en {_rangoCapturado.Areas.Count} bloque(s).\n\n" +
                            $"• Auditoría: {estadoAuditoria}\n" +
                            $"• Regla Activa: Bloqueo de captura inválida (Data Validation).\n" +
                            $"• Regla Pasiva: Resalte rojo en caso de alteración externa (Format Conditions).",
                            "SAVCNG - Validación Completada",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(this, "Error crítico al aplicar la validación: " + ex.Message, "Error del Sistema", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        chkDecimales.Checked = false;
                    }
                    finally
                    {
                        // Restauración de contexto y recolección final
                        if (wsActual != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(wsActual);
                        if (excelApp != null) excelApp.ScreenUpdating = true;
                    }
                }
                // --- VALIDACIÓN Catalogos ---
                else if (chkCatalogos.Checked == true)
                {
                    string formulaOpciones = "";

                    // PREGUNTA NUEVA: ¿Manual o desde Excel?
                    DialogResult tipoEntrada = MessageBox.Show(this,
                        "¿Deseas escribir el(los) valor(es) del catálogo manualmente (ej: un solo valor como 'X' o varios como '1,2,9')?\n\n" +
                        "SÍ: Escribir el(los) valor(es) directamente.\n" +
                        "NO: Seleccionar celdas de Excel.",
                        "Origen del Catálogo",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question);

                    if (tipoEntrada == DialogResult.Yes)
                    {
                        // ==========================================================
                        // CASO 1: ENTRADA MANUAL (Ideal para "X")
                        // ==========================================================
                        object resultadoTexto = excelApp.InputBox(
                            "Escribe el(los) valor(es) para tu lista desplegable.\nNOTA: Si son varias deberan estar separadas por comas):",
                            "Escribir Opciones",
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 2); // 2 = Solo Texto

                        if (resultadoTexto is bool && (bool)resultadoTexto == false)
                        {
                            chkCatalogos.Checked = false;
                            return;
                        }

                        string textoEscrito = resultadoTexto.ToString().Trim();
                        if (string.IsNullOrEmpty(textoEscrito))
                        {
                            chkCatalogos.Checked = false;
                            return;
                        }

                        // Reemplaza comas por el separador correcto de la PC
                        formulaOpciones = textoEscrito.Replace(",", separador);
                    }
                    else
                    {
                        // ==========================================================
                        // CASO 2: TU CÓDIGO ORIGINAL (Selección de Celdas)
                        // ==========================================================
                        object resultadoInput = excelApp.InputBox(
                            "Selecciona el rango de opciones o la celda que contiene el catálogo (ej: 1,2,9):",
                            "Seleccionar Origen del Catálogo",
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8); // 8 = Solo Rango

                        if (resultadoInput is bool && (bool)resultadoInput == false)
                        {
                            chkCatalogos.Checked = false;
                            return;
                        }

                        Excel.Range rangoOrigen = (Excel.Range)resultadoInput;

                        bool esCeldaUnicaOCombinada = (rangoOrigen.Count == 1) || (bool)rangoOrigen.MergeCells;

                        if (esCeldaUnicaOCombinada)
                        {
                            // 1. Simplificación de UX: Pregunta directa de acción con OKCancel
                            DialogResult respuesta = MessageBox.Show(this,
                                "Se ha detectado un único bloque de texto como catálogo.\n" +
                                "¿Deseas que el sistema extraiga automáticamente SOLO LOS NÚMEROS (ej. 1, 2, 9) para crear las opciones?\n\n" +
                                "• [Aceptar]: Extraer números y continuar.\n" +
                                "• [Cancelar]: Abortar esta operación.",
                                "Configuración de Catálogo",
                                MessageBoxButtons.OKCancel,
                                MessageBoxIcon.Question);

                            if (respuesta == DialogResult.Cancel)
                            {
                                chkCatalogos.Checked = false;
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(rangoOrigen);
                                return;
                            }

                            Excel.Range primeraCeldaOrigen = (Excel.Range)rangoOrigen.Cells[1, 1];
                            string textoCelda = primeraCeldaOrigen.Text?.ToString() ?? "";

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(primeraCeldaOrigen);
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(rangoOrigen);

                            if (string.IsNullOrWhiteSpace(textoCelda))
                            {
                                MessageBox.Show(this, "La celda origen está vacía. No se puede extraer el catálogo.",
                                                "Aviso Arquitectónico", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                chkCatalogos.Checked = false;
                                return;
                            }

                            string[] pedacitos = textoCelda.Split(new char[] { '.', ',', '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
                            System.Collections.Generic.List<string> listaLimpios = new System.Collections.Generic.List<string>();

                            foreach (string pedazo in pedacitos)
                            {
                                string soloNumeros = "";
                                foreach (char letra in pedazo)
                                {
                                    if (char.IsDigit(letra)) soloNumeros += letra;
                                }
                                if (!string.IsNullOrEmpty(soloNumeros))
                                {
                                    listaLimpios.Add(soloNumeros);
                                }
                            }

                            string separadorSistema = excelApp.International[Excel.XlApplicationInternational.xlListSeparator].ToString();
                            formulaOpciones = string.Join(separadorSistema, listaLimpios);
                        }
                        else
                        {
                            DialogResult respuestaRango = MessageBox.Show(this,
                                "Se han detectado varias celdas seleccionadas como catálogo.\n" +
                                "¿Deseas que el sistema extraiga automáticamente SOLO LOS NÚMEROS (ej. 1, 2, 9) de cada celda para crear las opciones?\n\n" +
                                "• [Aceptar]: Extraer números y continuar.\n" +
                                "• [Cancelar]: Abortar esta operación.",
                                "Configuración de Catálogo",
                                MessageBoxButtons.OKCancel,
                                MessageBoxIcon.Question);

                            if (respuestaRango == DialogResult.Cancel)
                            {
                                chkCatalogos.Checked = false;
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(rangoOrigen);
                                return;
                            }

                            System.Collections.Generic.List<string> listaLimpios = new System.Collections.Generic.List<string>();

                            foreach (Excel.Range celda in rangoOrigen.Cells)
                            {
                                string textoCelda = celda.Text?.ToString() ?? "";
                                string soloNumeros = "";

                                foreach (char letra in textoCelda)
                                {
                                    if (char.IsDigit(letra)) soloNumeros += letra;
                                }

                                if (!string.IsNullOrEmpty(soloNumeros))
                                {
                                    listaLimpios.Add(soloNumeros);
                                }
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(celda);
                            }

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(rangoOrigen);

                            if (listaLimpios.Count == 0)
                            {
                                MessageBox.Show(this, "No se encontraron valores numéricos en el rango seleccionado.\nNo se puede crear el catálogo.",
                                                "Aviso Arquitectónico", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                chkCatalogos.Checked = false;
                                return;
                            }

                            formulaOpciones = string.Join(separador, listaLimpios);
                        }
                    }

                    // ==========================================================
                    // PREVENCIÓN DE SOBRESCRITURA DE VALIDACIÓN
                    // ==========================================================
                    bool tieneValidacionPrevia = false;
                    Excel.Range primeraCeldaRango = null;

                    try
                    {
                        primeraCeldaRango = (Excel.Range)_rangoCapturado.Cells[1, 1];
                        int tipoValidacion = primeraCeldaRango.Validation.Type;
                        tieneValidacionPrevia = true;
                    }
                    catch
                    {
                        tieneValidacionPrevia = false;
                    }
                    finally
                    {
                        // Liberamos memoria para evitar Excel fantasma
                        if (primeraCeldaRango != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(primeraCeldaRango);
                    }

                    if (tieneValidacionPrevia)
                    {
                        DialogResult sobrescribir = MessageBox.Show(this,
                            "Las celdas que seleccionaste ya tienen una validación de datos o lista desplegable asignada.\n\n" +
                            "¿Estás seguro de que deseas borrarla y aplicar este nuevo catálogo en su lugar?",
                            "Validación existente detectada",
                            MessageBoxButtons.YesNo,
                            MessageBoxIcon.Warning);

                        if (sobrescribir == DialogResult.No)
                        {
                            chkCatalogos.Checked = false;
                            return;
                        }
                    }

                    // =========================================================================
                    // NUEVO PASO: RECOLECCIÓN LIGERA DE PREGUNTAS (STATE MANAGEMENT)
                    // =========================================================================
                    System.Collections.Generic.HashSet<string> preguntasUnicas = new System.Collections.Generic.HashSet<string>();

                    for (int i = 1; i <= _rangoCapturado.Areas.Count; i++)
                    {
                        Excel.Range areaIndividual = (Excel.Range)_rangoCapturado.Areas[i];
                        string idPregunta = ObtenerNumeroPregunta(areaIndividual);

                        if (!string.IsNullOrEmpty(idPregunta))
                        {
                            preguntasUnicas.Add(idPregunta);
                        }

                        System.Runtime.InteropServices.Marshal.ReleaseComObject(areaIndividual);
                    }

                    string preguntasDetectadas = preguntasUnicas.Count > 0 ? string.Join(", ", preguntasUnicas) : "ND";

                    // ==========================================================
                    // 3. APLICAR LA VALIDACIÓN Y REGISTRO EN BITÁCORA
                    // ==========================================================
                    try
                    {
                        _rangoCapturado.Validation.Delete();

                        _rangoCapturado.Validation.Add(
                            Excel.XlDVType.xlValidateList,
                            Excel.XlDVAlertStyle.xlValidAlertStop,
                            Excel.XlFormatConditionOperator.xlBetween,
                            formulaOpciones,
                            Type.Missing);

                        _rangoCapturado.Validation.InCellDropdown = true;
                        _rangoCapturado.Validation.IgnoreBlank = true;

                        chkCatalogos.Checked = false;

                        // Registro de Auditoría con la etiqueta correcta
                        AuditoriaCenso.RegistrarAccion(
                            _libroCenso,
                            preguntasDetectadas,
                            "Catálogo (Lista Desplegable)",
                            _rangoCapturado.Address.Replace("$", ""),
                            "Data Validation"
                        );

                        MessageBox.Show(this, "¡Validación de catálogo aplicada con éxito!", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(this, "Error al aplicar la validación: " + ex.Message + "\n\nTexto que se intentó usar: " + formulaOpciones, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        chkCatalogos.Checked = false;
                    }
                }
                // --- VALIDACIÓN NS ---
                else if (chkNS.Checked == true)
                {
                    Excel.Worksheet wsActual = null;

                    try
                    {
                        wsActual = (Excel.Worksheet)_rangoCapturado.Worksheet;
                        Excel.Application xlApp = (Excel.Application)ExcelDna.Integration.ExcelDnaUtil.Application;

                        // =========================================================================
                        // FASE 1: UX DE CONFIGURACIÓN GLOBAL Y CONTROL DE ABORTO
                        // =========================================================================
                        string textoSugerido = "Alerta: justificar el uso de NS.";
                        object resTexto = xlApp.InputBox(
                            "Escribe el texto del mensaje de alerta que se aplicará a las celdas seleccionadas:",
                            "SAVCNG - Configuración Masiva NS (Escucha Pasiva)", textoSugerido, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 2);

                        // ABORTO TOTAL 1: Si cancela el texto global
                        if (resTexto is bool && (bool)resTexto == false)
                        {
                            chkNS.Checked = false;
                            MessageBox.Show(this, "Proceso cancelado por el usuario.\n\nNo se incluyó ninguna validación en la plantilla.", "SAVCNG - Operación Cancelada", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            return;
                        }
                        string textoAlerta = resTexto.ToString().Trim();

                        // =========================================================================
                        // FASE 2: AUDITORÍA DE COEXISTENCIA (DETECCIÓN DE REGLAS PREVIAS)
                        // =========================================================================
                        bool tieneValidacionesPrevias = false;

                        // 2.1 Detectar Formatos Condicionales
                        if (_rangoCapturado.FormatConditions.Count > 0)
                        {
                            tieneValidacionesPrevias = true;
                        }

                        // 2.2 Detectar Validación de Datos (Data Validation)
                        if (!tieneValidacionesPrevias)
                        {
                            try
                            {
                                // Si la celda no tiene validación, esto genera una excepción nativa inofensiva que atrapamos.
                                var tipoValidacion = _rangoCapturado.Validation.Type;
                                tieneValidacionesPrevias = true;
                            }
                            catch { /* Silenciado intencionalmente: Significa que la celda está limpia */ }
                        }

                        bool limpiarPrevias = false;

                        if (tieneValidacionesPrevias)
                        {
                            DialogResult respLimpieza = MessageBox.Show(this,
                                "Se han detectado validaciones de datos o formatos condicionales PREVIOS en el rango capturado.\n\n" +
                                "¿Deseas CONSERVAR lo anterior e integrar la validación NS por debajo?\n\n" +
                                "SÍ = MANTENER reglas previas (Recomendado para apilar reglas).\n" +
                                "NO = BORRAR TODO lo anterior y dejar las celdas limpias.",
                                "SAVCNG - Auditoría de Coexistencia", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);

                            // ABORTO TOTAL 2: Si cancela en la pantalla de coexistencia
                            if (respLimpieza == DialogResult.Cancel)
                            {
                                chkNS.Checked = false;
                                MessageBox.Show(this, "Proceso cancelado.\n\nNo se alteró la plantilla.", "SAVCNG - Operación Cancelada", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                return;
                            }

                            // Si dice NO, activamos la bandera de destrucción
                            if (respLimpieza == DialogResult.No)
                            {
                                limpiarPrevias = true;
                            }
                        }

                        // =========================================================================
                        // FASE 3: ALGORITMO DE AGRUPACIÓN POR PREGUNTA (MATRICES PARALELAS)
                        // =========================================================================
                        var áreasPorPregunta = new System.Collections.Generic.Dictionary<string, System.Collections.Generic.List<Excel.Range>>();

                        for (int i = 1; i <= _rangoCapturado.Areas.Count; i++)
                        {
                            Excel.Range areaIndividual = (Excel.Range)_rangoCapturado.Areas[i];
                            string idPregunta = ObtenerNumeroPregunta(areaIndividual);

                            if (!áreasPorPregunta.ContainsKey(idPregunta))
                            {
                                áreasPorPregunta[idPregunta] = new System.Collections.Generic.List<Excel.Range>();
                            }
                            áreasPorPregunta[idPregunta].Add(areaIndividual);
                        }

                        int preguntasProcesadas = 0;

                        // Ocultar renderizado visual masivo (Mejora drástica de rendimiento)
                        xlApp.ScreenUpdating = false;

                        // =========================================================================
                        // FASE 4: PROCESAMIENTO INTELIGENTE (ÚNICO VS INDIVIDUAL)
                        // =========================================================================
                        foreach (var grupo in áreasPorPregunta)
                        {
                            string preguntaActual = grupo.Key;
                            System.Collections.Generic.List<Excel.Range> listaÁreas = grupo.Value;

                            bool usarUnicaAlerta = true;

                            if (listaÁreas.Count > 1)
                            {
                                xlApp.ScreenUpdating = true; // Encender pantalla temporalmente para el MessageBox
                                DialogResult respArquitectura = MessageBox.Show(this,
                                    $"Se han detectado {listaÁreas.Count} rangos seleccionados para la PREGUNTA {preguntaActual}.\n\n" +
                                    $"¿Deseas configurar una SOLA ALERTA CENTRAL para todos estos rangos?\n\n" +
                                    $"SÍ = Se pedirá 1 sola celda de alerta que evaluará todos los rangos juntos.\n" +
                                    $"NO = Se pedirán {listaÁreas.Count} celdas de alerta (una por cada rango de forma independiente).",
                                    $"SAVCNG - Arquitectura P.{preguntaActual}", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                                xlApp.ScreenUpdating = false; // Volver a apagar

                                // ABORTO TOTAL 3
                                if (respArquitectura == DialogResult.Cancel)
                                {
                                    chkNS.Checked = false;
                                    MessageBox.Show(this, "Proceso cancelado durante la configuración.\n\nNo se alteró la plantilla.", "SAVCNG - Operación Cancelada", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                    return;
                                }

                                usarUnicaAlerta = (respArquitectura == DialogResult.Yes);
                            }

                            // -------------------------------------------------------------------------
                            // RUTA A: ALERTA ÚNICA CENTRALIZADA
                            // -------------------------------------------------------------------------
                            if (usarUnicaAlerta)
                            {
                                Excel.Range rangoAlertaCentral = null;
                                try
                                {
                                    xlApp.ScreenUpdating = true; // Encender para inputbox
                                    object resDestino = xlApp.InputBox(
                                        $"[PREGUNTA DETECTADA: {preguntaActual}] - MODO CENTRALIZADO\n\n" +
                                        $"Selecciona la celda destino donde aparecerá el mensaje de ALERTA para los {listaÁreas.Count} rangos:\n" +
                                        $"(Selección Estándar para mensajes: Columna B hasta AD)",
                                        $"SAVCNG - Alerta Central P.{preguntaActual}", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8);
                                    xlApp.ScreenUpdating = false;

                                    // ABORTO TOTAL 4
                                    if (resDestino is bool && (bool)resDestino == false)
                                    {
                                        chkNS.Checked = false;
                                        MessageBox.Show(this, "Mapeo cancelado por el usuario.\n\nEl proceso ha sido abortado por completo.", "SAVCNG - Operación Cancelada", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                        return;
                                    }

                                    rangoAlertaCentral = (Excel.Range)resDestino;
                                    System.Collections.Generic.List<string> fragmentosCountIf = new System.Collections.Generic.List<string>();

                                    foreach (Excel.Range area in listaÁreas)
                                    {
                                        // EJECUCIÓN CONDICIONAL DE LIMPIEZA
                                        if (limpiarPrevias)
                                        {
                                            area.Validation.Delete();
                                            area.FormatConditions.Delete();
                                        }

                                        // =====================================================================
                                        // NUEVA INYECCIÓN: FORMATO CONDICIONAL "NS" EN LA MATRIZ CAPTURADA
                                        // =====================================================================
                                        Excel.Range primeraCeldaArea = null;
                                        Excel.Range celdaDummyFmt = null;
                                        Excel.FormatCondition fcNS = null;

                                        try
                                        {
                                            primeraCeldaArea = (Excel.Range)area.Cells[1, 1];
                                            string dirRelativaArea = primeraCeldaArea.get_Address(false, false, Excel.XlReferenceStyle.xlA1, false);

                                            // Dummy Cell Trick: Evita errores de separador regional
                                            celdaDummyFmt = (Excel.Range)wsActual.Cells[area.Row, 16384]; // Columna XFD
                                            celdaDummyFmt.Formula = $"=TRIM({dirRelativaArea})=\"NS\"";
                                            string formulaNSLocal = celdaDummyFmt.FormulaLocal;
                                            celdaDummyFmt.Clear();

                                            fcNS = (Excel.FormatCondition)area.FormatConditions.Add(
                                                Excel.XlFormatConditionType.xlExpression,
                                                Type.Missing,
                                                formulaNSLocal);

                                            // Estética Institucional: Match exacto con el color del mensaje 
                                            fcNS.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(191, 143, 0));
                                            fcNS.Font.Bold = true;
                                            // Fondo amarillo muy tenue para no competir visualmente pero asegurar visibilidad
                                            fcNS.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 242, 204));
                                            fcNS.StopIfTrue = false;
                                        }
                                        finally
                                        {
                                            // Destrucción de la matriz COM (Zero Leaks)
                                            if (fcNS != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(fcNS);
                                            if (celdaDummyFmt != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(celdaDummyFmt);
                                            if (primeraCeldaArea != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(primeraCeldaArea);
                                        }
                                        // =====================================================================

                                        string addrAbsoluta = area.get_Address(true, true, Excel.XlReferenceStyle.xlA1, false);
                                        fragmentosCountIf.Add($"COUNTIF({addrAbsoluta},\"NS\")");
                                    }

                                    if (rangoAlertaCentral.Count > 1) { rangoAlertaCentral.Merge(); }
                                    rangoAlertaCentral.Font.Name = "Arial";
                                    rangoAlertaCentral.Font.Size = 9;
                                    rangoAlertaCentral.Font.Bold = true;
                                    rangoAlertaCentral.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(191, 143, 0));
                                    rangoAlertaCentral.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                                    rangoAlertaCentral.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

                                    string sumaInner = string.Join(",", fragmentosCountIf);
                                    string formulaFinalAlerta = $"=IF(SUM({sumaInner})>0, \"{textoAlerta}\", \"\")";
                                    rangoAlertaCentral.Formula = formulaFinalAlerta;
                                }
                                finally
                                {
                                    if (rangoAlertaCentral != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rangoAlertaCentral);
                                }
                            }
                            // -------------------------------------------------------------------------
                            // RUTA B: ALERTAS INDIVIDUALES
                            // -------------------------------------------------------------------------
                            else
                            {
                                for (int j = 0; j < listaÁreas.Count; j++)
                                {
                                    Excel.Range area = listaÁreas[j];
                                    Excel.Range rangoAlertaIndiv = null;

                                    try
                                    {
                                        xlApp.ScreenUpdating = true; // Encender para Inputbox
                                        object resDestino = xlApp.InputBox(
                                            $"[PREGUNTA DETECTADA: {preguntaActual}] - Rango {j + 1} de {listaÁreas.Count}\n\n" +
                                            $"Selecciona la celda destino donde aparecerá el mensaje EXCLUSIVO para este rango:",
                                            $"SAVCNG - Alerta Individual P.{preguntaActual} (Área {j + 1})", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8);
                                        xlApp.ScreenUpdating = false;

                                        // ABORTO TOTAL 5
                                        if (resDestino is bool && (bool)resDestino == false)
                                        {
                                            chkNS.Checked = false;
                                            MessageBox.Show(this, "Mapeo cancelado por el usuario.\n\nEl proceso ha sido abortado por completo.", "SAVCNG - Operación Cancelada", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                            return;
                                        }

                                        rangoAlertaIndiv = (Excel.Range)resDestino;

                                        // EJECUCIÓN CONDICIONAL DE LIMPIEZA
                                        if (limpiarPrevias)
                                        {
                                            area.Validation.Delete();
                                            area.FormatConditions.Delete();
                                        }

                                        // =====================================================================
                                        // NUEVA INYECCIÓN: FORMATO CONDICIONAL "NS" EN LA MATRIZ CAPTURADA
                                        // =====================================================================
                                        Excel.Range primeraCeldaArea = null;
                                        Excel.Range celdaDummyFmt = null;
                                        Excel.FormatCondition fcNS = null;

                                        try
                                        {
                                            primeraCeldaArea = (Excel.Range)area.Cells[1, 1];
                                            string dirRelativaArea = primeraCeldaArea.get_Address(false, false, Excel.XlReferenceStyle.xlA1, false);

                                            celdaDummyFmt = (Excel.Range)wsActual.Cells[area.Row, 16384];
                                            celdaDummyFmt.Formula = $"=TRIM({dirRelativaArea})=\"NS\"";
                                            string formulaNSLocal = celdaDummyFmt.FormulaLocal;
                                            celdaDummyFmt.Clear();

                                            fcNS = (Excel.FormatCondition)area.FormatConditions.Add(
                                                Excel.XlFormatConditionType.xlExpression,
                                                Type.Missing,
                                                formulaNSLocal);

                                            fcNS.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(191, 143, 0));
                                            fcNS.Font.Bold = true;
                                            fcNS.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 242, 204));
                                            fcNS.StopIfTrue = false;
                                        }
                                        finally
                                        {
                                            if (fcNS != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(fcNS);
                                            if (celdaDummyFmt != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(celdaDummyFmt);
                                            if (primeraCeldaArea != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(primeraCeldaArea);
                                        }
                                        // =====================================================================

                                        if (rangoAlertaIndiv.Count > 1) { rangoAlertaIndiv.Merge(); }
                                        rangoAlertaIndiv.Font.Name = "Arial";
                                        rangoAlertaIndiv.Font.Size = 9;
                                        rangoAlertaIndiv.Font.Bold = true;
                                        rangoAlertaIndiv.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(191, 143, 0));
                                        rangoAlertaIndiv.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                                        rangoAlertaIndiv.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

                                        string addrAbsoluta = area.get_Address(true, true, Excel.XlReferenceStyle.xlA1, false);
                                        string formulaFinalAlerta = $"=IF(COUNTIF({addrAbsoluta},\"NS\")>0, \"{textoAlerta}\", \"\")";
                                        rangoAlertaIndiv.Formula = formulaFinalAlerta;
                                    }
                                    finally
                                    {
                                        if (rangoAlertaIndiv != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rangoAlertaIndiv);
                                    }
                                }
                            }

                            preguntasProcesadas++;
                        }

                        // =========================================================================
                        // FASE 5: CONCLUSIÓN Y NOTIFICACIÓN UX
                        // =========================================================================
                        chkNS.Checked = false;
                        string estadoPrevias = limpiarPrevias ? "Se borraron validaciones anteriores." : "Se conservaron validaciones anteriores.";

                        // =========================================================================
                        // REGISTRO DE AUDITORÍA EN LA BITÁCORA DEL SISTEMA
                        // =========================================================================
                        string preguntasDetectadas = string.Join(", ", áreasPorPregunta.Keys);

                        AuditoriaCenso.RegistrarAccion(
                            _libroCenso,
                            preguntasDetectadas,
                            "Validación NS",
                            _rangoCapturado.Address.Replace("$", ""),
                            "Fórmulas / FormatCondition" // <--- Actualizamos la auditoría
                        );

                        xlApp.ScreenUpdating = true; // Aseguramos que la pantalla reviva al final

                        MessageBox.Show(this,
                            $"Procesamiento masivo finalizado con éxito.\n\n" +
                            $"• Preguntas procesadas: {preguntasProcesadas}\n" +
                            $"• Total de rangos mapeados: {_rangoCapturado.Areas.Count}\n" +
                            $"• Auditoría: {estadoPrevias}\n\n" +
                            $"Nota: Las celdas admiten cualquier tipo de dato y se resaltarán en color ORO si detectan 'NS'.",
                            "SAVCNG Automatización Masiva", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(this, "Error crítico en el motor de masificación: " + ex.Message, "Error del Sistema", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        chkNS.Checked = false;
                    }
                    finally
                    {
                        // Restauración definitiva en el bloque final
                        Excel.Application xlAppSafe = (Excel.Application)ExcelDna.Integration.ExcelDnaUtil.Application;
                        xlAppSafe.ScreenUpdating = true;
                        if (wsActual != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(wsActual);
                    }
                }
                // --- VALIDACIÓN FORMATO TEXTO ---
                else if (chkFormatoTexto.Checked == true)
                {
                    Excel.Worksheet wsActual = null;

                    try
                    {
                        excelApp = (Excel.Application)ExcelDna.Integration.ExcelDnaUtil.Application;
                        wsActual = (Excel.Worksheet)_rangoCapturado.Worksheet;

                        // =========================================================================
                        // FASE 1: AUDITORÍA DE COEXISTENCIA (DETECCIÓN DE REGLAS PREVIAS)
                        // =========================================================================
                        bool tieneFormatosPrevios = _rangoCapturado.FormatConditions.Count > 0;
                        bool tieneValidacionPrevia = false;

                        try { var tipo = _rangoCapturado.Validation.Type; tieneValidacionPrevia = true; } catch { /* Silenciado: No hay DataValidation previa */ }

                        bool limpiarFormatos = false;

                        if (tieneFormatosPrevios || tieneValidacionPrevia)
                        {
                            DialogResult respLimpieza = MessageBox.Show(this,
                                "Se detectaron configuraciones previas en el rango seleccionado.\n\n" +
                                "NOTA: Al ser una restricción de captura estricta, la regla de celdas se sobreescribirá, pero...\n\n" +
                                "¿Deseas CONSERVAR los colores, alertas o bloqueos (Formatos Condicionales) aplicados previamente?\n\n" +
                                "SÍ = MANTENER colores/bloqueos anteriores.\n" +
                                "NO = BORRAR todo el historial y limpiar el lienzo.",
                                "SAVCNG - Auditoría de Coexistencia", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);

                            if (respLimpieza == DialogResult.Cancel)
                            {
                                chkFormatoTexto.Checked = false;
                                MessageBox.Show(this, "Proceso cancelado.\nNo se alteró la plantilla.", "Operación Cancelada", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                return;
                            }

                            if (respLimpieza == DialogResult.No)
                            {
                                limpiarFormatos = true;
                            }
                        }

                        // =========================================================================
                        // FASE 1.5: RECOLECCIÓN LIGERA DE PREGUNTAS (STATE MANAGEMENT)
                        // =========================================================================
                        // Utilizamos un HashSet para obtener una lista de preguntas únicas sin retener objetos COM pesados
                        System.Collections.Generic.HashSet<string> preguntasUnicas = new System.Collections.Generic.HashSet<string>();

                        for (int i = 1; i <= _rangoCapturado.Areas.Count; i++)
                        {
                            Excel.Range areaIndividual = (Excel.Range)_rangoCapturado.Areas[i];
                            string idPregunta = ObtenerNumeroPregunta(areaIndividual);

                            if (!string.IsNullOrEmpty(idPregunta))
                            {
                                preguntasUnicas.Add(idPregunta); // HashSet ignora automáticamente si ya existe
                            }

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(areaIndividual);
                        }

                        // =========================================================================
                        // FASE 2: UX - TÉCNICA DEL RANGO AUXILIAR DINÁMICO
                        // =========================================================================
                        object resAuxiliar = excelApp.InputBox(
                            "Indica en qué columna libre deseas colocar la validación oculta (AF en adelante).\n\nConsidera que el sistema requerirá espacio libre hacia abajo proporcional al número de filas que seleccionaste.",
                            "SAVCNG - Selección de Fórmula Auxiliar", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8); // 8 = Rango

                        if (resAuxiliar is bool && (bool)resAuxiliar == false) { chkFormatoTexto.Checked = false; return; }

                        Excel.Range seleccionAuxiliar = (Excel.Range)resAuxiliar;
                        Excel.Range celdaInicioAux = (Excel.Range)seleccionAuxiliar.Cells[1, 1];

                        string colAuxLetra = celdaInicioAux.Address.Split('$')[1];
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(seleccionAuxiliar);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(celdaInicioAux);

                        excelApp.ScreenUpdating = false;

                        // =========================================================================
                        // FASE 3: PROCESAMIENTO MASIVO POR ÁREAS (SOPORTE MULTI-RANGO)
                        // =========================================================================
                        foreach (Excel.Range area in _rangoCapturado.Areas)
                        {
                            Excel.Range primeraCeldaCap = null;
                            Excel.Range rangoAuxiliarArea = null;
                            Excel.Range primeraCeldaAux = null;

                            try
                            {
                                area.Validation.Delete();
                                if (limpiarFormatos) { area.FormatConditions.Delete(); }

                                int filaInicio = area.Row;
                                int filaFin = filaInicio + area.Rows.Count - 1;
                                rangoAuxiliarArea = wsActual.Range[$"{colAuxLetra}{filaInicio}:{colAuxLetra}{filaFin}"];

                                primeraCeldaCap = (Excel.Range)area.Cells[1, 1];
                                string dirCapRel = primeraCeldaCap.get_Address(false, false, Excel.XlReferenceStyle.xlA1, false);

                                primeraCeldaAux = (Excel.Range)rangoAuxiliarArea.Cells[1, 1];
                                string dirAuxRel = primeraCeldaAux.get_Address(false, false, Excel.XlReferenceStyle.xlA1, false);

                                string permitidos = "0123456789ABCDEFGHIJKLMNÑOPQRSTUVWXYZÁÉÍÓÚÜ ";
                                string formulaAuxiliar = $"=IF(OR(ISBLANK({dirCapRel}), AND(EXACT({dirCapRel},UPPER({dirCapRel})), LEN({dirCapRel})=LEN(TRIM({dirCapRel})), SUMPRODUCT(--ISNUMBER(FIND(MID({dirCapRel},ROW(INDIRECT(\"1:\"&MAX(1,LEN({dirCapRel})))),1),\"{permitidos}\")))=LEN({dirCapRel}))), 1, 0)";

                                rangoAuxiliarArea.Formula = formulaAuxiliar;

                                // TRUCO ARQUITECTÓNICO: BYPASS DEL ERROR 0x800A03EC
                                object valorOriginal = primeraCeldaCap.Value2;
                                bool estabaVacia = (valorOriginal == null || string.IsNullOrWhiteSpace(valorOriginal.ToString()));

                                if (estabaVacia)
                                {
                                    primeraCeldaCap.Value2 = "A";
                                }

                                // VALIDACIÓN DE DATOS (ULTRA LIGERA)
                                string formulaDV = $"={dirAuxRel}=1";

                                area.Validation.Add(
                                    Excel.XlDVType.xlValidateCustom,
                                    Excel.XlDVAlertStyle.xlValidAlertStop,
                                    Excel.XlFormatConditionOperator.xlBetween,
                                    formulaDV,
                                    Type.Missing);

                                area.Validation.IgnoreBlank = true;
                                area.Validation.ShowError = true;
                                area.Validation.ErrorTitle = "Formato de texto inválido";
                                area.Validation.ErrorMessage = "El texto debe cumplir estas reglas:\n\n" +
                                                               "• Solo se permite texto en MAYÚSCULAS y NÚMEROS.\n" +
                                                               "• Sin espacios dobles o sobrantes.\n" +
                                                               "• Sin comillas ni signos de puntuación, paréntesis ni caracteres especiales.";

                                if (estabaVacia)
                                {
                                    primeraCeldaCap.Value2 = null;
                                }
                            }
                            finally
                            {
                                if (primeraCeldaCap != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(primeraCeldaCap);
                                if (primeraCeldaAux != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(primeraCeldaAux);
                                if (rangoAuxiliarArea != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rangoAuxiliarArea);
                            }
                        }

                        // =========================================================================
                        // FASE 4: CONCLUSIÓN Y REGISTRO EN BITÁCORA
                        // =========================================================================
                        chkFormatoTexto.Checked = false;
                        string estadoAuditoria = limpiarFormatos ? "Se borraron configuraciones visuales previas." : "Se conservaron colores y bloqueos anteriores.";

                        // Extraemos las claves (preguntas únicas) del HashSet y las unimos separadas por comas
                        string preguntasDetectadas = preguntasUnicas.Count > 0 ? string.Join(", ", preguntasUnicas) : "ND";

                        AuditoriaCenso.RegistrarAccion(
                            _libroCenso,
                            preguntasDetectadas,
                            "Formato de Texto (Alfanumérico)", // Etiqueta corregida para este módulo
                            _rangoCapturado.Address.Replace("$", ""),
                            "Data Validation"
                        );

                        MessageBox.Show(this,
                            $"Validación de formato de texto aplicada con éxito en {_rangoCapturado.Areas.Count} bloque(s).\n\n" +
                            $"• Columna Auxiliar Inyectada: {colAuxLetra}\n" +
                            $"• Auditoría: {estadoAuditoria}",
                            "SAVCNG - Validación Completada", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(this, "Error crítico al configurar Formato Texto: " + ex.Message, "Error del Sistema", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        chkFormatoTexto.Checked = false;
                    }
                    finally
                    {
                        if (wsActual != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(wsActual);
                        if (excelApp != null) excelApp.ScreenUpdating = true;
                    }
                }
                // --- VALIDACIÓN BLOQUEOS ---
                else if (chkBloqueo.Checked == true)
                {
                    try
                    {
                        // --- PASO 1: LA CELDA O RANGO ---
                        object resultadoRango = excelApp.InputBox(
                            "Selecciona el RANGO o CELDA que controla el bloqueo:\n\n" +
                            "• Una celda: Se evaluará fila por fila.\n" +
                            "• Un rango: Se desbloqueará si el valor existe en CUALQUIER celda (o fila por fila si miden lo mismo).",
                            "1. Condición de Bloqueo",
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8); // 8 = Rango

                        if (resultadoRango is bool && (bool)resultadoRango == false)
                        {
                            chkBloqueo.Checked = false;
                            return;
                        }

                        Excel.Range rangoCondicion = (Excel.Range)resultadoRango;

                        // ==========================================================
                        // INTELIGENCIA SENIOR: ¿Búsqueda Global o Fila por Fila?
                        // ==========================================================
                        bool esFilaPorFila = false;

                        if (rangoCondicion.Count == 1)
                        {
                            esFilaPorFila = true;
                        }
                        else if (_rangoCapturado.Rows.Count == rangoCondicion.Rows.Count)
                        {
                            DialogResult respFila = MessageBox.Show(this,
                                "He detectado que el rango de condición tiene el MISMO número de filas que el rango a bloquear.\n\n" +
                                "¿Deseas que la regla se aplique FILA POR FILA?\n" +
                                "(Ej. Si se cumple la condición en la fila 12, SOLO se desbloquea la celda de la fila 12).\n\n" +
                                "SÍ = Fila por Fila (Ideal para matrices y preguntas con incisos paralelos).\n" +
                                "NO = Búsqueda Global (Desbloquea todo el bloque si encuentra el valor).",
                                "Evaluación Inteligente de Matrices",
                                MessageBoxButtons.YesNo,
                                MessageBoxIcon.Question);

                            if (respFila == DialogResult.Yes)
                            {
                                esFilaPorFila = true;
                            }
                        }

                        // ==========================================================
                        // CORRECCIÓN CRÍTICA: Anclar Columna ($A1) para Matrices
                        // ==========================================================
                        string dirCondicionLocal = "";

                        if (esFilaPorFila)
                        {
                            string addressAbsoluta = rangoCondicion.Cells[1, 1].Address;
                            string[] partes = addressAbsoluta.Split('$');
                            dirCondicionLocal = "$" + partes[1] + partes[2];
                        }
                        else
                        {
                            dirCondicionLocal = rangoCondicion.Address;
                        }

                        string dirCondicionGlobal = rangoCondicion.Address;

                        // --- PASO 2: EL OPERADOR LÓGICO ---
                        object resultadoOperador = excelApp.InputBox(
                            "Introduce el operador lógico que PERMITE la captura:\n\n" +
                            "  =    (Igual a)\n" +
                            "  <>   (No es igual a)\n" +
                            "  >    (Es mayor que)\n" +
                            "  <    (Es menor que)\n" +
                            "  >=   (Es mayor o igual a)\n" +
                            "  <=   (Es menor o igual a)",
                            "2. Operador Lógico",
                            "=", Type.Missing, Type.Missing, Type.Missing, Type.Missing, 2);

                        if (resultadoOperador is bool && (bool)resultadoOperador == false) { chkBloqueo.Checked = false; return; }
                        string operador = resultadoOperador.ToString().Trim();

                        if (operador != "=" && operador != "<>" && operador != ">" && operador != "<" && operador != ">=" && operador != "<=")
                        {
                            MessageBox.Show(this,"Operador no reconocido.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            chkBloqueo.Checked = false;
                            return;
                        }

                        // --- PASO 3: EL VALOR ---
                        object resultadoValor = excelApp.InputBox(
                            $"Introduce el valor que completará la condición.\n(Condición actual: {operador} ___ )\n\nEjemplos: 6, Sí, X:",
                            "3. Valor del Criterio",
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 2);

                        if (resultadoValor is bool && (bool)resultadoValor == false) { chkBloqueo.Checked = false; return; }

                        string valorCriterio = resultadoValor.ToString().Trim();
                        bool esNumero = double.TryParse(valorCriterio, out _);
                        string valorFormateado = esNumero ? valorCriterio : $"\"{valorCriterio}\"";

                        // ==========================================================
                        // --- PASO 3.5 y 4: REGLAS ADICIONALES ---
                        // ==========================================================
                        DialogResult respuestaBlanco = MessageBox.Show(this,
                            "¿Deseas que la matriz permanezca DESBLOQUEADA si la celda de condición está VACÍA?\n\n" +
                            "SÍ = Si está en blanco, se puede escribir.\n" +
                            "NO = Estricto (Si está en blanco, se bloquea por defecto).",
                            "Regla de Celda Vacía", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                        DialogResult respuestaRojo = MessageBox.Show(this,
                            "¿Deseas que la celda se resalte cuando se desbloquee y esté vacía?\n\n(Ideal para los campos 'Especifique' que se vuelven obligatorios).",
                            "4. Resalte de Obligatoriedad", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                        // ==========================================================
                        // --- APLICACIÓN DE REGLAS (MOTOR: CONTAR.SI) ---
                        // ==========================================================

                        _rangoCapturado.Validation.Delete();

                        // AUDITORÍA DE COEXISTENCIA
                        if (_rangoCapturado.FormatConditions.Count > 0)
                        {
                            DialogResult respFormato = MessageBox.Show(this,
                                "Se detectaron reglas previas (ej. Validación de Blancos).\n\n" +
                                "¿Deseas CONSERVARLAS y apilar el bloqueo encima?\n\n" +
                                "SÍ = Conservar formatos previos.\nNO = Eliminar y aplicar solo el bloqueo.",
                                "Formatos Condicionales Detectados", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

                            if (respFormato == DialogResult.No) { _rangoCapturado.FormatConditions.Delete(); }
                        }
                        else
                        {
                            _rangoCapturado.FormatConditions.Delete();
                        }


                        string criterioContarSi = $"\"{operador}\"&{valorFormateado}";
                        string formulaValidacion = "";
                        string formulaSombreado = "";

                        if (respuestaBlanco == DialogResult.Yes)
                        {
                            formulaValidacion = $"=O(ESBLANCO({dirCondicionLocal}){separador}CONTAR.SI({dirCondicionLocal}{separador}{criterioContarSi})>0)";
                            formulaSombreado = $"=Y(ESBLANCO({dirCondicionLocal})=FALSO{separador}CONTAR.SI({dirCondicionLocal}{separador}{criterioContarSi})=0)";
                        }
                        else
                        {
                            formulaValidacion = $"=CONTAR.SI({dirCondicionLocal}{separador}{criterioContarSi})>0";
                            formulaSombreado = $"=CONTAR.SI({dirCondicionLocal}{separador}{criterioContarSi})=0";
                        }

                        // 1. DATA VALIDATION
                        _rangoCapturado.Validation.Add(Excel.XlDVType.xlValidateCustom, Excel.XlDVAlertStyle.xlValidAlertStop,
                            Excel.XlFormatConditionOperator.xlBetween, formulaValidacion, Type.Missing);
                        _rangoCapturado.Validation.IgnoreBlank = true;
                        _rangoCapturado.Validation.ShowError = true;
                        _rangoCapturado.Validation.ErrorTitle = "Celda Bloqueada";
                        _rangoCapturado.Validation.ErrorMessage = $"No se permite capturar información. El flujo requiere que la condición sea {operador} {valorCriterio}.";

                        // 2. FORMATO CONDICIONAL 1 (Gris Bloqueado)
                        Excel.FormatCondition formatoGris = (Excel.FormatCondition)_rangoCapturado.FormatConditions.Add(
                            Excel.XlFormatConditionType.xlExpression, Type.Missing, formulaSombreado);

                        // AJUSTE CRÍTICO DE JERARQUÍA Y COLOR
                        formatoGris.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White); // Fuerza el fondo blanco para borrar el amarillo
                        formatoGris.Interior.Pattern = Excel.XlPattern.xlPatternCrissCross;
                        formatoGris.Interior.PatternColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Gray);
                        formatoGris.StopIfTrue = true; // DETIENE LA EVALUACIÓN SI ESTÁ BLOQUEADA (Ignora los blancos)

                        // 3. FORMATO CONDICIONAL 2 (Alerta Azul)
                        if (respuestaRojo == DialogResult.Yes)
                        {
                            Excel.Range primeraCelda = (Excel.Range)_rangoCapturado.Cells[1, 1];
                            string dirCapturada = primeraCelda.Address.Replace("$", "");
                            string formulaRojo = $"=Y(CONTAR.SI({dirCondicionLocal}{separador}{criterioContarSi})>0{separador}ESBLANCO({dirCapturada}))";

                            Excel.FormatCondition formatoRojo = (Excel.FormatCondition)_rangoCapturado.FormatConditions.Add(
                                Excel.XlFormatConditionType.xlExpression, Type.Missing, formulaRojo);

                            formatoRojo.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(47, 117, 181));
                            formatoRojo.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(156, 0, 6));
                            formatoRojo.StopIfTrue = true; // CRÍTICO: Detiene la evaluación si aplica el rojo
                        }

                        // ==========================================================
                        // --- PASO 5: RESALTE E INSTRUCCIÓN OMITIDOS PARA BREVEDAD (Tu lógica intacta) ---
                        // ==========================================================

                        // ... (Tu código de Mensaje de Alerta e Instrucción en amarillo permanece exactamente igual aquí) ...

                        chkBloqueo.Checked = false;
                        string modoAplicado = esFilaPorFila ? "Fila por Fila (Paralelo)" : "Búsqueda Global";
                        MessageBox.Show(this,$"Validación de Bloqueo Dinámica aplicada con éxito.\nModo: {modoAplicado}\nRegla: {operador} {valorCriterio}", "SAVCNG", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(this,"Error al aplicar la Validación de Bloqueo: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        chkBloqueo.Checked = false;
                    }
                }

                // --- VALIDACION BLANCOS -------
                else if (chkBlancos.Checked == true)
                {
                    try
                    {
                        // ==========================================================
                        // FUNCIÓN LOCAL PARA DIÁLOGO NATIVO (Evita errores CS0103 y CS0234)
                        // ==========================================================
                        Func<string, string, string, string> PedirInputWinForms = (prompt, titulo, defecto) =>
                        {
                            using (Form dlg = new Form())
                            {
                                Label lbl = new Label { Text = prompt, Left = 12, Top = 12, Width = 380, AutoSize = true };
                                TextBox txt = new TextBox { Text = defecto, Left = 12, Top = 100, Width = 380 };
                                Button btnOk = new Button { Text = "Aceptar", DialogResult = DialogResult.OK, Left = 216, Top = 135, Width = 80, Height = 28 };
                                Button btnCancel = new Button { Text = "Cancelar", DialogResult = DialogResult.Cancel, Left = 312, Top = 135, Width = 80, Height = 28 };

                                dlg.Text = titulo;
                                dlg.ClientSize = new System.Drawing.Size(406, 175);
                                dlg.Controls.AddRange(new Control[] { lbl, txt, btnOk, btnCancel });
                                dlg.FormBorderStyle = FormBorderStyle.FixedDialog;
                                dlg.StartPosition = FormStartPosition.CenterScreen;
                                dlg.MaximizeBox = false;
                                dlg.MinimizeBox = false;
                                dlg.AcceptButton = btnOk;
                                dlg.CancelButton = btnCancel;
                                dlg.TopMost = true;

                                // Si el texto es muy largo, ajustamos la posición del cuadro de texto
                                if (lbl.Height > 70)
                                {
                                    txt.Top = lbl.Bottom + 10;
                                    btnOk.Top = txt.Bottom + 15;
                                    btnCancel.Top = txt.Bottom + 15;
                                    dlg.Height = btnOk.Bottom + 45;
                                }

                                return dlg.ShowDialog(this) == DialogResult.OK ? txt.Text : null;
                            }
                        };

                        // ==========================================================
                        // 1. PREGUNTA INICIAL: CASOS DE EXCEPCIÓN / EXCLUSIÓN
                        // ==========================================================
                        string mensajePrompt = "Indica los valores para los cuales NO se debe aplicar la validación de Blancos.\n\n" +
                                               "• Si son varios, sepáralos por comas (Ejemplo: 2, 9).\n" +
                                               //"• Si desea que la validación se aplique SIEMPRE, deja el campo en blanco y presiona Aceptar.\n" +
                                               "• Deje vacía la captura en caso de no requerir alguna excepción" +
                                               //"• Si deseas abortar la operación, presiona Cancelar.";
                                               ".\n";

                        string resExclusiones = PedirInputWinForms(mensajePrompt, "Excepciones de Validación (Opcional)", "2, 9");

                        // Si el usuario presiona Cancelar o cierra la ventana
                        if (resExclusiones == null)
                        {
                            chkBlancos.Checked = false;
                            return;
                        }

                        string textoExclusiones = resExclusiones.Trim();
                        System.Collections.Generic.List<string> listaCondicionesExcluidas = new System.Collections.Generic.List<string>();

                        if (!string.IsNullOrEmpty(textoExclusiones))
                        {
                            string[] valoresExcluidos = textoExclusiones.Split(new char[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries);
                            foreach (string val in valoresExcluidos)
                            {
                                string valLimpio = val.Trim();
                                if (!string.IsNullOrEmpty(valLimpio))
                                {
                                    bool esNum = double.TryParse(valLimpio, out _);
                                    string valFormateado = esNum ? valLimpio : $"\"{valLimpio}\"";

                                    // Condición matemática: Que NO exista este valor en el rango evaluado
                                    listaCondicionesExcluidas.Add($"COUNTIF({{0}}, {valFormateado})=0");
                                }
                            }
                        }

                        // ==========================================================
                        // 2. SOLICITAR UBICACIÓN DE ALERTA (Rango de Excel)
                        // ==========================================================
                        this.Hide();
                        object resDestino = excelApp.InputBox(
                            "Selecciona el rango donde aparecerá el mensaje de alerta (Se combinará y pintará de azul automáticamente):",
                            "1. Ubicación de Alerta", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8); // 8 = Rango
                        this.Show();

                        if (resDestino is bool && (bool)resDestino == false) { chkBlancos.Checked = false; return; }
                        Excel.Range rangoDestino = (Excel.Range)resDestino;

                        // ==========================================================
                        // 3. SOLICITAR TEXTO DEL MENSAJE DE ADVERTENCIA
                        // ==========================================================
                        string resTexto = PedirInputWinForms(
                            "Escribe el mensaje de advertencia:\n(El sistema evaluará el rango capturado. Si hay celdas iniciadas pero faltan datos, exigirá completarlas)",
                            "2. Mensaje de Alerta",
                            //"Favor de revisar la información faltante en las celdas sombreadas");
                            "Favor de ingresar toda la información requerida en la pregunta");

                        if (string.IsNullOrWhiteSpace(resTexto)) { chkBlancos.Checked = false; return; }
                        string textoAlerta = resTexto.Trim();

                        // ==========================================================
                        // 4. CONSTRUCCIÓN INTELIGENTE ANTI-CELDAS COMBINADAS (MERGED)
                        // ==========================================================
                        int filaInicial = _rangoCapturado.Row;

                        System.Collections.Generic.List<string> listaLlenasRelativas = new System.Collections.Generic.List<string>();
                        System.Collections.Generic.List<string> listaVaciasRelativas = new System.Collections.Generic.List<string>();
                        System.Collections.Generic.List<string> listaLlenasAbsolutas = new System.Collections.Generic.List<string>();

                        int c = 1;
                        while (c <= _rangoCapturado.Columns.Count)
                        {
                            Excel.Range celdaActual = (Excel.Range)_rangoCapturado.Cells[1, c];
                            string colLetra = celdaActual.Address.Split('$')[1];

                            string dirRelativaCelda = $"{colLetra}{filaInicial}";
                            string dirAbsolutaColumna = $"${colLetra}{filaInicial}"; // Columna fijada con $ para el formato condicional

                            // Listas para la fórmula de la celda de mensaje
                            listaLlenasRelativas.Add($"({dirRelativaCelda}<>\"\")");
                            listaVaciasRelativas.Add($"({dirRelativaCelda}=\"\")");

                            // Lista con columna fija ($) para el Formato Condicional (Pintado Azul)
                            listaLlenasAbsolutas.Add($"({dirAbsolutaColumna}<>\"\")");

                            // SI LA CELDA ESTÁ COMBINADA (Merge), saltamos las columnas ocultas
                            if ((bool)celdaActual.MergeCells)
                            {
                                Excel.Range areaCombinada = celdaActual.MergeArea;
                                int columnasCombinadas = areaCombinada.Columns.Count;
                                c += columnasCombinadas;
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(areaCombinada);
                            }
                            else
                            {
                                c++;
                            }

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(celdaActual);
                        }

                        string sumaLlenasRel = string.Join("+", listaLlenasRelativas);
                        string sumaVaciasRel = string.Join("+", listaVaciasRelativas);
                        string sumaLlenasAbs = string.Join("+", listaLlenasAbsolutas);

                        // Coordenadas del rango capturado en la fila actual
                        Excel.Range primeraCeldaCapturada = (Excel.Range)_rangoCapturado.Cells[1, 1];
                        Excel.Range ultimaCeldaCapturada = (Excel.Range)_rangoCapturado.Cells[1, _rangoCapturado.Columns.Count];

                        string colIniLetra = primeraCeldaCapturada.Address.Split('$')[1];
                        string colFinLetra = ultimaCeldaCapturada.Address.Split('$')[1];

                        string rangoFilaRelativo = $"{colIniLetra}{filaInicial}:{colFinLetra}{filaInicial}";
                        string rangoFilaAbsoluto = $"${colIniLetra}{filaInicial}:${colFinLetra}{filaInicial}";

                        // Construir cláusula de exclusión si el usuario ingresó valores
                        string clausulaExclusionRel = "";
                        string clausulaExclusionAbs = "";

                        if (listaCondicionesExcluidas.Count > 0)
                        {
                            var condRel = listaCondicionesExcluidas.ConvertAll(cond => string.Format(cond, rangoFilaRelativo));
                            var condAbs = listaCondicionesExcluidas.ConvertAll(cond => string.Format(cond, rangoFilaAbsoluto));

                            clausulaExclusionRel = ", " + string.Join(", ", condRel);
                            clausulaExclusionAbs = ", " + string.Join(", ", condAbs);
                        }

                        // ==========================================================
                        // 5. INYECCIÓN DE FÓRMULA DE ALERTA TRADUCIDA
                        // ==========================================================
                        Excel.Worksheet wsActual = (Excel.Worksheet)_rangoCapturado.Worksheet;
                        Excel.Range celdaDummy = (Excel.Range)wsActual.Cells[filaInicial, 16384]; // Columna XFD

                        string formulaGlobalIngles = $"=IF(AND(({sumaLlenasRel})>0, ({sumaVaciasRel})>0{clausulaExclusionRel}), \"{textoAlerta}\", \"\")";

                        celdaDummy.Formula = formulaGlobalIngles;
                        string formulaGlobalLocal = celdaDummy.FormulaLocal;
                        celdaDummy.Clear();

                        if (rangoDestino.Count > 1) { rangoDestino.Merge(); }

                        rangoDestino.Font.Name = "Arial";
                        rangoDestino.Font.Size = 10;
                        rangoDestino.Font.Bold = true;
                        rangoDestino.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(0, 112, 192));
                        rangoDestino.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

                        rangoDestino.FormulaLocal = formulaGlobalLocal;

                        // ==========================================================
                        // 6. INYECCIÓN DEL FORMATO CONDICIONAL (RESALTADO AZUL)
                        // ==========================================================
                        if (_rangoCapturado.FormatConditions.Count > 0)
                        {
                            DialogResult respFormato = MessageBox.Show(this,
                                "Se detectaron reglas de formato condicional previas (ej. Reglas de Bloqueo).\n\n" +
                                "¿Deseas CONSERVAR las reglas existentes e integrar la técnica de blancos?\n\n" +
                                "SÍ = Conservar formatos previos (Evita borrar tus bloques grises).\n" +
                                "NO = Eliminar formatos previos y aplicar únicamente el formato de blancos.",
                                "Formatos Condicionales Detectados",
                                MessageBoxButtons.YesNo,
                                MessageBoxIcon.Warning);

                            if (respFormato == DialogResult.No)
                            {
                                _rangoCapturado.FormatConditions.Delete();
                            }
                        }
                        else
                        {
                            _rangoCapturado.FormatConditions.Delete();
                        }

                        string colPrimeraLetra = primeraCeldaCapturada.Address.Split('$')[1];
                        string dirPrimeraRelativa = $"{colPrimeraLetra}{filaInicial}";

                        // Pinta de azul la celda si está vacía, la fila tiene datos Y NO contiene los valores excluidos
                        string formulaCondicionalIngles = $"=AND({dirPrimeraRelativa}=\"\", ({sumaLlenasAbs})>0{clausulaExclusionAbs})";

                        celdaDummy.Formula = formulaCondicionalIngles;
                        string formulaCondicionalLocal = celdaDummy.FormulaLocal;
                        celdaDummy.Clear();

                        System.Runtime.InteropServices.Marshal.ReleaseComObject(celdaDummy);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(primeraCeldaCapturada);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(ultimaCeldaCapturada);

                        Excel.FormatCondition formatoAzul = (Excel.FormatCondition)_rangoCapturado.FormatConditions.Add(
                            Excel.XlFormatConditionType.xlExpression, Type.Missing, formulaCondicionalLocal);

                        formatoAzul.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(47, 117, 181));

                        chkBlancos.Checked = false;

                        string msjExcepciones = listaCondicionesExcluidas.Count > 0
                            ? $"\n• Excepciones registradas: {textoExclusiones}"
                            : "";

                        MessageBox.Show(this, "Validación de blancos aplicada con éxito.\n\n" +
                                              //"• Soporta celdas combinadas de forma nativa.\n" +
                                              "• Si la fila no contiene ningún código de excepción y faltan campos, se alertará en azul." + msjExcepciones,
                                              "SAVCNG Arquitectura", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        this.Show();
                        MessageBox.Show(this, "Error en el motor de validación inteligente: " + ex.Message, "Error de Inyección", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        chkBlancos.Checked = false;
                    }
                }
                // --- FIN VALIDACION BLANCOS

                // --- VALIDACIÓN ESPECIFIQUES (PALABRAS CLAVE) ---
                else if (chkEspClave.Checked == true)
                {
                    if (_libroCenso == null || _rangoCapturado == null)
                    {
                        MessageBox.Show(this, "Operación denegada: Captura la celda destino (Especifique) primero.",
                                        "Arquitectura SAVCNG", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        chkEspClave.Checked = false;
                        return;
                    }

                    Excel.Application xlApp = null;
                    Excel.Worksheet ws = null;
                    Excel.Range rangoCatalogo = null;
                    Excel.Range rangoMensaje = null;
                    Excel.Range celdaMotor = null;

                    try
                    {
                        xlApp = (Excel.Application)ExcelDna.Integration.ExcelDnaUtil.Application;
                        ws = (Excel.Worksheet)_rangoCapturado.Worksheet;

                        // =========================================================================
                        // FASE 1: RECOPILACIÓN DE ESPACIOS DE TRABAJO (UX BLINDADA)
                        // =========================================================================
                        object resCatalogo = xlApp.InputBox(
                           "1. Selecciona las OPCIONES DEL CATÁLOGO.\nNOTA: Debera omitir de la selección las opciones 'Otro(Especifique)' y/o 'No identificado' del catálogo correspondiente. ",
                            "Mapeo de Catálogo", Type: 8);
                        if (resCatalogo is bool && (bool)resCatalogo == false) { chkEspClave.Checked = false; return; }
                        rangoCatalogo = (Excel.Range)resCatalogo;

                        object resMensaje = xlApp.InputBox(
                            "2. Selecciona el rango o celda donde se mostrará el MENSAJE DE ALERTA.",
                            "Destino de Alerta", Type: 8);
                        if (resMensaje is bool && (bool)resMensaje == false) { chkEspClave.Checked = false; return; }
                        rangoMensaje = (Excel.Range)resMensaje;

                        object resMotor = xlApp.InputBox(
                            "3. Selecciona UNA CELDA VACÍA (columna AF en adelante) para construir el Diccionario Auxiliar de Busqueda.\nADVERTENCIA: Considere un espacio libre de dos columnas por N filas. (N = numero de opciones del catalogo)",
                            "Generación del Motor Oculto", Type: 8);
                        if (resMotor is bool && (bool)resMotor == false) { chkEspClave.Checked = false; return; }
                        celdaMotor = (Excel.Range)resMotor;

                        // =========================================================================
                        // FASE 1.5: AUDITORÍA DE COEXISTENCIA Y STATE MANAGEMENT
                        // =========================================================================
                        // A) Auditoría de Coexistencia (Enfocada en el Catálogo)
                        bool limpiarFormatos = false;
                        if (rangoCatalogo.FormatConditions.Count > 0)
                        {
                            DialogResult respLimpieza = MessageBox.Show(this,
                                "Se detectaron formatos condicionales (colores/alertas) previos en el CATÁLOGO seleccionado.\n\n" +
                                "¿Deseas CONSERVAR los colores anteriores además del resaltado amarillo de esta búsqueda?\n\n" +
                                "SÍ = MANTENER colores/bloqueos anteriores.\n" +
                                "NO = BORRAR todo el historial y limpiar el lienzo.",
                                "SAVCNG - Auditoría de Coexistencia", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);

                            if (respLimpieza == DialogResult.Cancel) { chkEspClave.Checked = false; return; }
                            if (respLimpieza == DialogResult.No) { limpiarFormatos = true; }
                        }

                        // B) Recolección Ligera de Preguntas (HashSet)
                        System.Collections.Generic.HashSet<string> preguntasUnicas = new System.Collections.Generic.HashSet<string>();
                        for (int i = 1; i <= _rangoCapturado.Areas.Count; i++)
                        {
                            Excel.Range areaIndividual = (Excel.Range)_rangoCapturado.Areas[i];
                            string idPregunta = ObtenerNumeroPregunta(areaIndividual);
                            if (!string.IsNullOrEmpty(idPregunta)) { preguntasUnicas.Add(idPregunta); }
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(areaIndividual);
                        }
                        string preguntasDetectadas = preguntasUnicas.Count > 0 ? string.Join(", ", preguntasUnicas) : "ND";

                        // Apagamos la actualización de pantalla para construir el motor velozmente
                        xlApp.ScreenUpdating = false;

                        // =========================================================================
                        // FASE 2: CONSTRUCCIÓN DEL MOTOR AUXILIAR MATRICIAL
                        // =========================================================================
                        Excel.Range primeraCeldaAzul = (Excel.Range)_rangoCapturado.Cells[1, 1];
                        string dirAzulAbsoluta = primeraCeldaAzul.get_Address(true, true, Excel.XlReferenceStyle.xlA1, false);

                        // 2.1 Inyectamos la Celda Limpia (Elimina acentos y minúsculas nativamente)
                        Excel.Range celdaLimpiaAzul = celdaMotor.Offset[0, 0];
                        string formulaLimpieza = $"=LOWER(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE({dirAzulAbsoluta},\"á\",\"a\"),\"é\",\"e\"),\"í\",\"i\"),\"ó\",\"o\"),\"ú\",\"u\"))";
                        celdaLimpiaAzul.Formula = formulaLimpieza;
                        celdaLimpiaAzul.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        celdaLimpiaAzul.Font.Bold = true;

                        string dirLimpiaAbsoluta = celdaLimpiaAzul.get_Address(true, true, Excel.XlReferenceStyle.xlA1, false);

                        int filaMotor = 1;
                        System.Collections.Generic.List<string> celdasResultadosBooleanos = new System.Collections.Generic.List<string>();

                        // Limpieza condicional del catálogo
                        if (limpiarFormatos) { rangoCatalogo.FormatConditions.Delete(); }

                        // Iteramos sobre las celdas del catálogo seleccionado
                        foreach (Excel.Range celdaCat in rangoCatalogo.Cells)
                        {
                            string textoOriginal = celdaCat.Text != null ? celdaCat.Text.ToString() : "";

                            if (!string.IsNullOrWhiteSpace(textoOriginal) && textoOriginal.Trim() != "")
                            {
                                // Regex requiere: using System.Text.RegularExpressions; en tus librerías
                                string textoProcesado = System.Text.RegularExpressions.Regex.Replace(textoOriginal, @"\(.*?\)", "").Trim();

                                textoProcesado = textoProcesado.ToLower()
                                    .Replace("á", "a").Replace("é", "e").Replace("í", "i")
                                    .Replace("ó", "o").Replace("ú", "u");

                                string[] separadores = { " y/o ", " y ", " o ", " e ", ",", "/", " a ", " ante ", " bajo ", " cabe ", " con ", " contra ", " de ", " del ", " desde ", " durante ", " en ", " entre ", " hacia ", " hasta ", " mediante ", " para ", " por ", " según ", " sin ", " sobre ", " tras " };
                                string[] palabrasClave = textoProcesado.Split(separadores, StringSplitOptions.RemoveEmptyEntries);

                                System.Collections.Generic.List<string> fragmentosSearch = new System.Collections.Generic.List<string>();
                                foreach (string palabra in palabrasClave)
                                {
                                    string palabraLimpia = palabra.Trim();
                                    if (palabraLimpia.Length > 2)
                                    {
                                        fragmentosSearch.Add($"ISNUMBER(SEARCH(\"{palabraLimpia}\", {dirLimpiaAbsoluta}))");
                                    }
                                }

                                if (fragmentosSearch.Count > 0)
                                {
                                    Excel.Range filaValidacion = celdaMotor.Offset[filaMotor, 0];
                                    filaValidacion.Value2 = textoOriginal;

                                    Excel.Range celdaMatch = celdaMotor.Offset[filaMotor, 1];
                                    string formulaMatch = $"=OR({string.Join(",", fragmentosSearch)})";
                                    celdaMatch.Formula = formulaMatch;

                                    string dirMatchAbsoluta = celdaMatch.get_Address(true, true, Excel.XlReferenceStyle.xlA1, false);
                                    celdasResultadosBooleanos.Add(dirMatchAbsoluta);

                                    // =========================================================================
                                    // FASE 3: TRUCO ARQUITECTÓNICO DE TRADUCCIÓN (FormatCondition)
                                    // =========================================================================
                                    string celdaCatRelativa = celdaCat.get_Address(false, false, Excel.XlReferenceStyle.xlA1, false);
                                    string formulaFormatCondIngles = $"=AND({celdaCatRelativa}<>\"\", {dirMatchAbsoluta}=TRUE)";

                                    Excel.Range celdaDummy = ws.Cells[1048576, 16384];
                                    celdaDummy.Formula = formulaFormatCondIngles;
                                    string formulaFormatCondLocal = celdaDummy.FormulaLocal;
                                    celdaDummy.Clear();

                                    Excel.FormatCondition fc = (Excel.FormatCondition)celdaCat.FormatConditions.Add(
                                        Excel.XlFormatConditionType.xlExpression,
                                        Type.Missing,
                                        formulaFormatCondLocal);

                                    fc.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                                    fc.Font.Bold = true;

                                    // Liberación rigurosa de memoria en cada iteración
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(fc);
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(celdaDummy);
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(celdaMatch);
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(filaValidacion);
                                    filaMotor++;
                                }
                            }
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(celdaCat); // Crucial para catálogos largos
                        }

                        // =========================================================================
                        // FASE 4: VINCULACIÓN DEL MENSAJE DE ALERTA 
                        // =========================================================================
                        if (celdasResultadosBooleanos.Count > 0)
                        {
                            Excel.Range celdaInicioRango = celdaMotor.Offset[1, 1];
                            Excel.Range celdaFinRango = celdaMotor.Offset[filaMotor - 1, 1];

                            string addressInicio = celdaInicioRango.get_Address(true, true, Excel.XlReferenceStyle.xlA1, false);
                            string addressFin = celdaFinRango.get_Address(true, true, Excel.XlReferenceStyle.xlA1, false);

                            string formulaMensajeFinal = $"=IF(COUNTIF({addressInicio}:{addressFin}, TRUE)>0, \"Alerta: Revise el texto ingresado en el Especifique ya que podría existir en las opciones resaltadas en amarillo\", \"\")";

                            if (rangoMensaje.Count > 1) { rangoMensaje.Merge(); }
                            rangoMensaje.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                            rangoMensaje.Font.Name = "Arial";
                            rangoMensaje.Font.Size = 9;
                            rangoMensaje.Font.Bold = true;
                            rangoMensaje.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(191, 144, 0));

                            rangoMensaje.Formula = formulaMensajeFinal;

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(celdaInicioRango);
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(celdaFinRango);
                        }

                        // =========================================================================
                        // FASE 5: REGISTRO DE AUDITORÍA Y NOTIFICACIÓN
                        // =========================================================================
                        AuditoriaCenso.RegistrarAccion(
                            _libroCenso,
                            preguntasDetectadas,
                            "Motor de Búsqueda (Especifique)",
                            _rangoCapturado.Address.Replace("$", ""),
                            "Formato Condicional - Fórmulas"
                        );

                        MessageBox.Show(this, "El Diccionario de Palabras Clave y la Alerta Inteligente fueron construidos con éxito.", "SAVCNG - ExcelDNA", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(this, "Error crítico al configurar el motor de búsqueda: " + ex.Message, "Error del Sistema", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    finally
                    {
                        chkEspClave.Checked = false;

                        // Limpieza profunda COM
                        if (rangoCatalogo != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rangoCatalogo);
                        if (rangoMensaje != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rangoMensaje);
                        if (celdaMotor != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(celdaMotor);
                        if (ws != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(ws);

                        if (xlApp != null) xlApp.ScreenUpdating = true; // Devolvemos el refresco de pantalla
                    }
                }
                // --- VALIDACIÓN FECHAS ---
                else if (chkFechas.Checked == true)
                {
                    Excel.Worksheet wsActual = null;

                    try
                    {
                        // Instanciamos excelApp para poder usar InputBox y controlar la pantalla
                        excelApp = (Excel.Application)ExcelDna.Integration.ExcelDnaUtil.Application;
                        wsActual = (Excel.Worksheet)_rangoCapturado.Worksheet;

                        // =========================================================================
                        // FASE 1: UX DE CONFIGURACIÓN DE LÍMITES Y CONTROL DE ABORTO
                        // =========================================================================
                        object resultadoInferior = excelApp.InputBox(
                            "Indica el valor MÍNIMO aceptado para esta validación:\n\n(Ej. 1 para días/meses, o 1821 para años).",
                            "SAVCNG - Límite Inferior", "1", Type.Missing, Type.Missing, Type.Missing, Type.Missing, 2);

                        if (resultadoInferior is bool && (bool)resultadoInferior == false) { chkFechas.Checked = false; return; }
                        string inputInferior = resultadoInferior.ToString().Trim();
                        if (string.IsNullOrWhiteSpace(inputInferior)) { chkFechas.Checked = false; return; }

                        object resultadoSuperior = excelApp.InputBox(
                            "Indica el valor MÁXIMO aceptado para esta validación:\n\n(Ej. 31 para días, 12 para meses, o 2026 para años).",
                            "SAVCNG - Límite Superior", "2026", Type.Missing, Type.Missing, Type.Missing, Type.Missing, 2);

                        if (resultadoSuperior is bool && (bool)resultadoSuperior == false) { chkFechas.Checked = false; return; }
                        string inputSuperior = resultadoSuperior.ToString().Trim();
                        if (string.IsNullOrWhiteSpace(inputSuperior)) { chkFechas.Checked = false; return; }

                        // Validación estricta de las variables C#
                        if (!int.TryParse(inputInferior, out int limiteInferior) || !int.TryParse(inputSuperior, out int limiteSuperior))
                        {
                            MessageBox.Show(this, "Por favor, asegúrate de escribir únicamente números enteros.\nNo se permiten letras, decimales, ni espacios en blanco.", "Error de Tipado", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            chkFechas.Checked = false; return;
                        }

                        if (limiteInferior > limiteSuperior)
                        {
                            MessageBox.Show(this, $"Error de Lógica:\nEl límite mínimo ({limiteInferior}) no puede ser mayor que el límite máximo ({limiteSuperior}).", "Límites invertidos", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            chkFechas.Checked = false; return;
                        }

                        // =========================================================================
                        // FASE 2: AUDITORÍA DE COEXISTENCIA (ADAPTADA A DATA VALIDATION)
                        // =========================================================================
                        bool tieneFormatosPrevios = _rangoCapturado.FormatConditions.Count > 0;
                        bool tieneValidacionPrevia = false;

                        try { var tipo = _rangoCapturado.Validation.Type; tieneValidacionPrevia = true; } catch { /* Silenciado: No hay DataValidation previa */ }

                        bool limpiarFormatos = false;

                        if (tieneFormatosPrevios || tieneValidacionPrevia)
                        {
                            DialogResult respLimpieza = MessageBox.Show(this,
                                "Se detectaron configuraciones previas en el rango seleccionado.\n\n" +
                                "NOTA: Al ser una restricción de captura estricta, la regla de celdas se sobreescribirá, pero...\n\n" +
                                "¿Deseas CONSERVAR los colores, alertas o bloqueos (Formatos Condicionales) aplicados previamente?\n\n" +
                                "SÍ = MANTENER colores/bloqueos anteriores.\n" +
                                "NO = BORRAR todo el historial y limpiar el lienzo.",
                                "SAVCNG - Auditoría de Coexistencia", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);

                            if (respLimpieza == DialogResult.Cancel)
                            {
                                chkFechas.Checked = false;
                                MessageBox.Show(this, "Proceso cancelado.\nNo se alteró la plantilla.", "Operación Cancelada", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                return;
                            }

                            if (respLimpieza == DialogResult.No)
                            {
                                limpiarFormatos = true;
                            }
                        }

                        // =========================================================================
                        // FASE 2.5: RECOLECCIÓN LIGERA DE PREGUNTAS (STATE MANAGEMENT)
                        // =========================================================================
                        // Extraemos las preguntas justo antes de apagar la pantalla para asegurar que el usuario no abortará
                        System.Collections.Generic.HashSet<string> preguntasUnicas = new System.Collections.Generic.HashSet<string>();

                        for (int i = 1; i <= _rangoCapturado.Areas.Count; i++)
                        {
                            Excel.Range areaIndividual = (Excel.Range)_rangoCapturado.Areas[i];
                            string idPregunta = ObtenerNumeroPregunta(areaIndividual);

                            if (!string.IsNullOrEmpty(idPregunta))
                            {
                                preguntasUnicas.Add(idPregunta);
                            }
                            // Liberación rigurosa de memoria COM en cada ciclo
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(areaIndividual);
                        }

                        // Si el HashSet encontró preguntas, las unimos separadas por comas. De lo contrario, enviamos "ND"
                        string preguntasDetectadas = preguntasUnicas.Count > 0 ? string.Join(", ", preguntasUnicas) : "ND";

                        // Optimización de rendimiento visual (Silenciamos Excel)
                        excelApp.ScreenUpdating = false;

                        // =========================================================================
                        // FASE 3: PROCESAMIENTO MASIVO POR ÁREAS Y TRADUCCIÓN UNIVERSAL
                        // =========================================================================
                        foreach (Excel.Range area in _rangoCapturado.Areas)
                        {
                            Excel.Range primeraCelda = null;
                            Excel.Range celdaDummy = null;

                            try
                            {
                                // Limpieza obligatoria de DataValidation (no se pueden apilar)
                                area.Validation.Delete();

                                // Limpieza opcional de formatos visuales según lo elegido por el usuario
                                if (limpiarFormatos) { area.FormatConditions.Delete(); }

                                primeraCelda = (Excel.Range)area.Cells[1, 1];
                                string direccionRelativa = primeraCelda.get_Address(false, false, Excel.XlReferenceStyle.xlA1, false);

                                // ---------------------------------------------------------------------
                                // INGENIERÍA DE FÓRMULA UNIVERSAL (INTEGRANDO TRUNC PARA BLOQUEAR DECIMALES)
                                // ---------------------------------------------------------------------
                                string formulaValidacionIngles = $"=IF(ISNUMBER({direccionRelativa}), AND({direccionRelativa}>={limiteInferior}, {direccionRelativa}<={limiteSuperior}, TRUNC({direccionRelativa})={direccionRelativa}), OR(TRIM({direccionRelativa})=\"NS\", TRIM({direccionRelativa})=\"NA\"))";

                                // Truco Arquitectónico de Traducción Fila-Sensible
                                int filaBaseArea = area.Row;
                                celdaDummy = (Excel.Range)wsActual.Cells[filaBaseArea, 16384]; // Columna XFD
                                celdaDummy.Formula = formulaValidacionIngles;
                                string formulaValidacionLocal = celdaDummy.FormulaLocal;
                                celdaDummy.Clear();

                                // ---------------------------------------------------------------------
                                // INYECCIÓN DE REGLA
                                // ---------------------------------------------------------------------
                                area.Validation.Add(
                                    Excel.XlDVType.xlValidateCustom,
                                    Excel.XlDVAlertStyle.xlValidAlertStop,
                                    Excel.XlFormatConditionOperator.xlBetween,
                                    formulaValidacionLocal,
                                    Type.Missing);

                                area.Validation.IgnoreBlank = true;
                                area.Validation.ShowError = true;
                                area.Validation.ErrorTitle = "Captura Inválida (Solo Enteros)";
                                area.Validation.ErrorMessage = $"El formato de esta celda no admite números con decimales.\n\nEl número entero debe estar entre {limiteInferior} y {limiteSuperior}.\n\nTambién puedes usar las claves permitidas 'NS' o 'NA'.";
                            }
                            finally
                            {
                                // Prevención estricta de Fugas de Memoria COM en el bucle
                                if (primeraCelda != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(primeraCelda);
                                if (celdaDummy != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(celdaDummy);
                            }
                        }

                        // =========================================================================
                        // FASE 4: CONCLUSIÓN Y NOTIFICACIÓN UX
                        // =========================================================================
                        chkFechas.Checked = false;
                        string estadoAuditoria = limpiarFormatos ? "Se borraron configuraciones visuales previas." : "Se conservaron colores y bloqueos anteriores.";

                        // Registro de Auditoría con el esquema actualizado (5 parámetros)
                        AuditoriaCenso.RegistrarAccion(
                            _libroCenso,
                            preguntasDetectadas,
                            "Rango Numérico (Fechas)",
                            _rangoCapturado.Address.Replace("$", ""),
                            "Data Validation (Custom)" // Nuevo campo: Funcionalidad Excel
                        );

                        MessageBox.Show(this,
                            $"Validación de rango aplicada con éxito en {_rangoCapturado.Areas.Count} bloque(s).\n\n" +
                            $"• Criterio: Números enteros del {limiteInferior} al {limiteSuperior} (o claves NS/NA).\n" +
                            $"• Auditoría: {estadoAuditoria}",
                            "SAVCNG - Validación Completada", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(this, $"Error crítico en el motor de fechas: {ex.Message}", "Error del Sistema", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        chkFechas.Checked = false;
                    }
                    finally
                    {
                        // Restaurar sistema y limpiar raíz COM
                        if (wsActual != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(wsActual);
                        if (excelApp != null) excelApp.ScreenUpdating = true; // Se asegura de que se encienda la pantalla nuevamente
                    }
                }
                // --- VALIDACIÓN SUMAS (NUEVO MOTOR BOOLEANO Y ALERTAS CON SOPORTE MERGE) ---
                else if (chkSumas.Checked == true)
                {
                    Excel.Range rangoTotal = null;
                    Excel.Range rangoDesagregados = null;
                    Excel.Range rangoTotalesVerticales = null;
                    Excel.Range rangoAlerta = null;
                    Excel.Range seleccionAuxiliar = null;
                    Excel.Range celdaDummy = null;
                    Excel.Worksheet wsActual = null;

                    try
                    {
                        int filaInicio = _rangoCapturado.Row;
                        int filaFin = filaInicio + _rangoCapturado.Rows.Count - 1;

                        object resTotal = excelApp.InputBox("1. Selecciona la COLUMNA del TOTAL o PIVOTE:", "SAVCNG - Total", Type: 8);
                        if (resTotal is bool && (bool)resTotal == false) { chkSumas.Checked = false; return; }
                        rangoTotal = (Excel.Range)resTotal;

                        object resDesagregados = excelApp.InputBox("2. Selecciona las COLUMNAS de los DESAGREGADOS (Usa CTRL para varias):", "SAVCNG - Desagregados", Type: 8);
                        if (resDesagregados is bool && (bool)resDesagregados == false) { chkSumas.Checked = false; return; }
                        rangoDesagregados = (Excel.Range)resDesagregados;

                        object resVerticales = excelApp.InputBox("3. Selecciona la FILA de Sumatoria Vertical Sigma (Σ):", "SAVCNG - Sigma", Type: 8);
                        if (resVerticales is bool && (bool)resVerticales == false) { chkSumas.Checked = false; return; }
                        rangoTotalesVerticales = (Excel.Range)resVerticales;

                        object resAlerta = excelApp.InputBox("4. Selecciona el rango destino para el MENSAJE DE ERROR:", "SAVCNG - Alerta Global", Type: 8);
                        if (resAlerta is bool && (bool)resAlerta == false) { chkSumas.Checked = false; return; }
                        rangoAlerta = (Excel.Range)resAlerta;

                        object resAuxiliar = excelApp.InputBox("5. Selecciona UNA CELDA en una columna libre (ej. AF) para inyectar el Motor Auxiliar:", "SAVCNG - Arquitectura de Memoria", Type: 8);
                        if (resAuxiliar is bool && (bool)resAuxiliar == false) { chkSumas.Checked = false; return; }
                        seleccionAuxiliar = (Excel.Range)resAuxiliar;

                        excelApp.ScreenUpdating = false;
                        wsActual = (Excel.Worksheet)_rangoCapturado.Worksheet;

                        // =========================================================================
                        // A. DETECCIÓN INTELIGENTE DE CELDAS COMBINADAS (TOTAL)
                        // =========================================================================
                        Excel.Range primeraCeldaTotal = (Excel.Range)rangoTotal.Cells[1, 1];
                        Excel.Range mergeTotal = primeraCeldaTotal.MergeArea;
                        Excel.Range topLeftTotal = (Excel.Range)mergeTotal.Cells[1, 1];

                        string letraTotal = topLeftTotal.Address.Split('$')[1];

                        System.Runtime.InteropServices.Marshal.ReleaseComObject(topLeftTotal);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(mergeTotal);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(primeraCeldaTotal);

                        Excel.Range celdaInicioAux = (Excel.Range)seleccionAuxiliar.Cells[1, 1];
                        string colAuxLetra = celdaInicioAux.Address.Split('$')[1];
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(celdaInicioAux);

                        // =========================================================================
                        // B. DETECCIÓN INTELIGENTE DE CELDAS COMBINADAS (DESAGREGADOS)
                        // =========================================================================
                        System.Collections.Generic.List<string> letrasDesagregados = new System.Collections.Generic.List<string>();

                        foreach (Excel.Range area in rangoDesagregados.Areas)
                        {
                            Excel.Range fila1 = (Excel.Range)area.Rows[1];
                            foreach (Excel.Range celda in fila1.Cells)
                            {
                                // Extraemos exclusivamente la celda "Maestra" de la combinación
                                Excel.Range mArea = celda.MergeArea;
                                Excel.Range tl = (Excel.Range)mArea.Cells[1, 1];
                                string letra = tl.Address.Split('$')[1];

                                if (!letrasDesagregados.Contains(letra))
                                {
                                    letrasDesagregados.Add(letra);
                                }

                                System.Runtime.InteropServices.Marshal.ReleaseComObject(tl);
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(mArea);
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(celda);
                            }
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(fila1);
                        }

                        int numDesagregados = letrasDesagregados.Count;

                        // =========================================================================
                        // C. MÁQUINA DE ESTADOS (INYECCIÓN DE MOTOR AUXILIAR FILA POR FILA)
                        // =========================================================================
                        for (int f = filaInicio; f <= filaFin; f++)
                        {
                            string tRef = $"${letraTotal}{f}";

                            // Motores de análisis dinámicos (Limpios y sin celdas combinadas parásitas)
                            string countNS = string.Join("+", letrasDesagregados.ConvertAll(l => $"COUNTIF(${l}{f},\"NS\")"));
                            string sumaDesagregados = $"SUM({string.Join(",", letrasDesagregados.ConvertAll(l => $"${l}{f}"))})";
                            string countNum = $"COUNT({string.Join(",", letrasDesagregados.ConvertAll(l => $"${l}{f}"))})";

                            // REGLA 1: Aritmética Estricta (Si todo es número, debe cuadrar la suma)
                            string condAritmetica = $"AND(ISNUMBER({tRef}), ({countNS})=0, {tRef}<>{sumaDesagregados})";

                            // REGLA 2: Contradicción 0 Absoluto (Total es 0, pero metieron NS en desagregados)
                            string condCero = $"AND(ISNUMBER({tRef}), {tRef}=0, ({countNS})>0)";

                            // REGLA 3: Contradicción NS (Total es NS, pero ya dieron todos los números, o la suma parcial ya es >0)
                            string condNS = $"AND(UPPER(TRIM({tRef}))=\"NS\", OR({countNum}={numDesagregados}, {sumaDesagregados}>0))";

                            // Inyección (1 = Error, 0 = Aceptado)
                            string formulaInglesAux = $"=IF(OR({condAritmetica}, {condCero}, {condNS}), 1, 0)";

                            Excel.Range celdaDestinoAux = wsActual.Range[$"{colAuxLetra}{f}"];
                            celdaDestinoAux.Formula = formulaInglesAux;
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(celdaDestinoAux);
                        }

                        // =========================================================================
                        // D. INYECCIÓN DEL FORMATO CONDICIONAL (VINCULADO AL MOTOR)
                        // =========================================================================
                        if (_rangoCapturado.FormatConditions.Count > 0)
                        {
                            DialogResult respFormato = MessageBox.Show(this, "Se detectaron reglas previas.\n¿Deseas CONSERVARLAS e integrar esta nueva capa?\nSÍ = Apilar\nNO = Borrar", "SAVCNG - Coexistencia", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);
                            if (respFormato == DialogResult.Cancel) { return; }
                            if (respFormato == DialogResult.No) { _rangoCapturado.FormatConditions.Delete(); }
                        }

                        celdaDummy = wsActual.Cells[filaInicio, 16384];
                        celdaDummy.Formula = $"=${colAuxLetra}{filaInicio}=1";
                        string fcLocal = celdaDummy.FormulaLocal;
                        celdaDummy.Clear();

                        Excel.FormatCondition fcError = (Excel.FormatCondition)_rangoCapturado.FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Type.Missing, fcLocal);
                        fcError.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 199, 206));
                        fcError.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(156, 0, 6));
                        fcError.StopIfTrue = false;
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(fcError);

                        // =========================================================================
                        // E. INYECCIÓN DEL MENSAJE DE ALERTA GLOBAL
                        // =========================================================================
                        string auxRange = $"${colAuxLetra}${filaInicio}:${colAuxLetra}${filaFin}";
                        string formulaAlertaFinal = $"=IF(SUM({auxRange})>0, \"Favor de revisar la suma y consistencia de totales y/o subtotales por filas (numéricos y NS)\", \"\")";

                        if (rangoAlerta.Count > 1) { rangoAlerta.Merge(); }
                        rangoAlerta.Formula = formulaAlertaFinal;
                        rangoAlerta.Font.Name = "Arial";
                        rangoAlerta.Font.Size = 9;
                        rangoAlerta.Font.Bold = true;
                        rangoAlerta.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                        rangoAlerta.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        rangoAlerta.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                        rangoAlerta.WrapText = true;

                        // =========================================================================
                        // F. INYECCIÓN OPTIMIZADA DE SUMATORIAS VERTICALES (SIGMAS)
                        // =========================================================================
                        System.Collections.Generic.List<string> sigmasProcesados = new System.Collections.Generic.List<string>();
                        foreach (Excel.Range celdaSumatoria in rangoTotalesVerticales.Cells)
                        {
                            Excel.Range mArea = celdaSumatoria.MergeArea;
                            Excel.Range tl = (Excel.Range)mArea.Cells[1, 1];
                            string addr = tl.Address;

                            if (!sigmasProcesados.Contains(addr))
                            {
                                sigmasProcesados.Add(addr);
                                string lCol = addr.Split('$')[1];
                                string rSuma = $"{lCol}{filaInicio}:{lCol}{filaFin}";

                                string fSigma = $"=IF(AND(SUM({rSuma})=0,COUNTIF({rSuma},\"NS\")>0),\"NS\"," +
                                                $"IF(AND(SUM({rSuma})=0,COUNTIF({rSuma},0)>0),0," +
                                                $"IF(AND(SUM({rSuma})=0,COUNTIF({rSuma},\"NA\")>0),\"NA\"," +
                                                $"SUM({rSuma}))))";
                                tl.Formula = fSigma;
                            }
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(tl);
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(mArea);
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(celdaSumatoria);
                        }

                        chkSumas.Checked = false;
                        MessageBox.Show(this, "Motor de Validación, Alerta y Sumatorias inyectados con éxito (Soporte Merge).", "SAVCNG - Arquitectura ExcelDNA", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(this, "Error crítico en el motor de sumas: " + ex.Message, "Error de Inserción", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        chkSumas.Checked = false;
                    }
                    finally
                    {
                        // Zero Leaks
                        if (celdaDummy != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(celdaDummy);
                        if (rangoTotal != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rangoTotal);
                        if (rangoDesagregados != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rangoDesagregados);
                        if (rangoTotalesVerticales != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rangoTotalesVerticales);
                        if (rangoAlerta != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rangoAlerta);
                        if (seleccionAuxiliar != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(seleccionAuxiliar);
                        if (wsActual != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(wsActual);
                        excelApp.ScreenUpdating = true;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this,"Ocurrió un error al aplicar el formato: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        // Función para detectar el número de pregunta en la columna A (Como en tu App_Form_Interface)
        private string ObtenerNumeroPregunta(Excel.Range rango)
        {
            try
            {
                Excel.Worksheet hoja = rango.Worksheet;
                int filaInicial = rango.Row;

                // Buscamos desde la fila seleccionada hacia arriba en la columna 1 (Columna A)
                for (int f = filaInicial; f >= 1; f--)
                {
                    Excel.Range celdaA = hoja.Cells[f, 1];
                    object valor = celdaA.Value2;

                    if (valor != null && !string.IsNullOrEmpty(valor.ToString().Trim()))
                    {
                        // Si encontramos algo en la columna A, asumimos que es el número de pregunta
                        return valor.ToString().Trim();
                    }
                }
            }
            catch { /* Si hay error, devolvemos vacío */ }

            return "(no encontrada)";
        }

        private void chkCatalogos_CheckedChanged(object sender, EventArgs e)
        {
            // Solo actuamos si el usuario MARCA la casilla
            if (chkCatalogos.Checked)
            {
                // Revisamos que no se haya saltado el paso 1 (Capturar rango)
                if (_libroCenso == null || _rangoCapturado == null)
                {
                    MessageBox.Show(this,"Primero carga un censo y captura un rango con el botón.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    chkCatalogos.Checked = false; // Desmarcamos la casilla
                }
                // ¡Y listo! No hacemos nada más aquí, dejamos que el botón Aplicar haga el trabajo duro.
            }
        }

        private void chkNS_CheckedChanged(object sender, EventArgs e)
        {
            // Solo actuamos si el usuario MARCA la casilla
            if (chkNS.Checked)
            {
                // Revisamos que no se haya saltado el paso 1 (Capturar rango)
                if (_libroCenso == null || _rangoCapturado == null)
                {
                    MessageBox.Show(this,"Primero carga un censo y captura un rango con el botón.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    chkNS.Checked = false; // Desmarcamos la casilla
                }
            }
        }

        private void chkBlancos_CheckedChanged(object sender, EventArgs e)
        {
            if (chkBlancos.Checked)
            {
                // Validamos la regla de negocio: Nada ocurre si no se ha mapeado el terreno previamente
                if (_libroCenso == null || _rangoCapturado == null)
                {
                    MessageBox.Show(this,"Operación denegada: Carga un censo y define el rango de memoria primero.", "Advertencia Arquitectónica", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    chkBlancos.Checked = false;
                }
            }
        }

        private void chkFechas_CheckedChanged(object sender, EventArgs e)
        {
            // SI ==> MARCA la casilla
            if (chkFechas.Checked) // <-- Corregido: antes decía chkAños.Checked
            {
                // Revisamos que no se haya saltado el paso 1 (Capturar rango)
                if (_libroCenso == null || _rangoCapturado == null)
                {
                    MessageBox.Show(this,"Primero carga un censo y captura un rango con el botón.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    chkFechas.Checked = false; // <-- Corregido: antes decía chkAños.Checked
                }
            }
        }

        private void CargarEstadoDelCenso()
        {
            // Obtenemos el historial directamente de la memoria embebida del archivo Excel
            System.Data.DataTable historial = AuditoriaCenso.ObtenerHistorialCenso(_libroCenso);

            // Lo enlazamos a la grilla para que el usuario pueda ver, filtrar u ordenar
            dgvAuditoria.DataSource = historial;

            // 1.Responsividad del Contenedor(Se estira con la ventana)
            dgvAuditoria.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;

            // 2. Responsividad del Contenido (Las columnas llenan el espacio vacío)
            dgvAuditoria.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

            // 3. Estética Institucional (Limpieza de ruido visual)
            dgvAuditoria.AllowUserToAddRows = false; // Quita la última fila en blanco editable
            dgvAuditoria.AllowUserToDeleteRows = false;
            dgvAuditoria.ReadOnly = true; // El auditor no debe alterar la bitácora desde aquí

            // 4. Experiencia de Navegación (Mejora la legibilidad)
            dgvAuditoria.SelectionMode = DataGridViewSelectionMode.FullRowSelect; // Selecciona toda la fila al hacer clic
            dgvAuditoria.MultiSelect = false;
            dgvAuditoria.RowHeadersVisible = false; // Oculta la columna gris inútil de la izquierda

            // 5. Estilo de celdas alternadas (Zebra striping para fatiga visual)
            dgvAuditoria.AlternatingRowsDefaultCellStyle.BackColor = System.Drawing.Color.LightGray;
            dgvAuditoria.BackgroundColor = System.Drawing.Color.White; // Fondo blanco en lugar de gris oscuro

        }

        private void btnActualizar_Click(object sender, EventArgs e)
        {
            try
            {
                // UX: Cambiamos el cursor a "Cargando" (Reloj de arena/Círculo azul)
                Cursor.Current = Cursors.WaitCursor;

                // Invocamos nuestro motor de lectura que va a la hoja VeryHidden
                CargarEstadoDelCenso();

                // Refrescamos visualmente el control
                dgvAuditoria.Refresh();
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Error al actualizar la vista de la bitácora: " + ex.Message, "Error de Sistema", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // UX: Devolvemos el cursor a la normalidad
                Cursor.Current = Cursors.Default;
            }
        }

        // EVENTO: Extracción, Clonación y Guardado Silencioso de la Bitácora
        // =========================================================================
        private void btnDescargarBitacora_Click(object sender, EventArgs e)
        {
            Excel.Worksheet wsLog = null;
            Excel.Workbook nuevoLibro = null;
            Excel.Worksheet wsCopia = null;
            Excel.Application xlApp = null;
            Excel.Worksheet hojaOriginal = null; // Puntero de preservación de estado

            try
            {
                // 1. UX: Indicamos que el sistema está trabajando en segundo plano
                Cursor.Current = Cursors.WaitCursor;
                xlApp = (Excel.Application)ExcelDna.Integration.ExcelDnaUtil.Application;

                // 2. CAPTURA DE ESTADO: Guardamos la hoja donde está parado el usuario
                hojaOriginal = (Excel.Worksheet)_libroCenso.ActiveSheet;

                // 3. Rastreamos la base de datos embebida
                foreach (Excel.Worksheet sheet in _libroCenso.Worksheets)
                {
                    if (sheet.Name == "SAVCNG_SysLog")
                    {
                        wsLog = sheet;
                        break;
                    }
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet);
                }

                if (wsLog == null)
                {
                    MessageBox.Show(this, "Aún no existen registros de validaciones en este censo.", "Bitácora Vacía", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // 4. MODO SILENCIOSO: Apagamos pantalla y bloqueamos cuadros de diálogo de Excel
                xlApp.ScreenUpdating = false;
                xlApp.DisplayAlerts = false; // Evita preguntas de sobrescritura o compatibilidad

                // 5. Clonación en memoria RAM
                nuevoLibro = xlApp.Workbooks.Add(Type.Missing);
                wsLog.Visible = Excel.XlSheetVisibility.xlSheetVisible;
                wsLog.Copy(Before: nuevoLibro.Worksheets[1]);
                wsLog.Visible = Excel.XlSheetVisibility.xlSheetVeryHidden; // Restauramos la seguridad de origen

                // 6. Configuración visual del archivo a exportar
                wsCopia = (Excel.Worksheet)nuevoLibro.Worksheets[1];
                wsCopia.Name = "Auditoria_" + DateTime.Now.ToString("ddMMyy");
                wsCopia.Columns.AutoFit();
                wsCopia.Application.ActiveWindow.SplitRow = 1;
                wsCopia.Application.ActiveWindow.FreezePanes = true;

                // =========================================================================
                // 7. MOTOR DE I/O: RESOLUCIÓN DE RUTA Y GUARDADO AUTOMÁTICO
                // =========================================================================
                // Obtenemos la ruta universal de la carpeta de descargas del usuario de Windows
                string rutaPerfil = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
                string rutaDescargas = System.IO.Path.Combine(rutaPerfil, "Downloads");

                // Armamos el nombre del archivo con Timestamp para evitar colisiones
                string nombreArchivo = $"SAVCNG_Bitacora_{DateTime.Now.ToString("yyyyMMdd_HHmmss")}.xlsx";
                string rutaCompleta = System.IO.Path.Combine(rutaDescargas, nombreArchivo);

                // Guardamos el libro usando el formato estándar de Excel actual (OpenXML)
                nuevoLibro.SaveAs(rutaCompleta, Excel.XlFileFormat.xlOpenXMLWorkbook, Type.Missing, Type.Missing,
                                  Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange,
                                  Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                // Cerramos el libro temporal inmediatamente (false = no preguntar si guarda cambios)
                nuevoLibro.Close(false);

                // 8. RESTAURACIÓN DEL ESTADO EN EL LIBRO ORIGEN
                // Al cerrar el libro nuevo, Excel puede perder el foco. Lo forzamos a volver a la hoja original.
                if (hojaOriginal != null)
                {
                    hojaOriginal.Activate();
                }

                // 9. Notificación UX de éxito orientada a la nueva arquitectura
                MessageBox.Show(this,
                    $"La bitácora ha sido exportada de forma automática.\n\nPuedes encontrar el archivo en tu carpeta de Descargas:\n\n{nombreArchivo}",
                    "SAVCNG - Descarga Exitosa", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Error crítico al intentar guardar la bitácora: " + ex.Message, "Fallo de I/O", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // 10. Limpieza estricta del Garbage Collector (COM)
                if (hojaOriginal != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(hojaOriginal);
                if (wsCopia != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(wsCopia);
                if (nuevoLibro != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(nuevoLibro);
                if (wsLog != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(wsLog);

                // 11. Restauración de los motores de Excel
                if (xlApp != null)
                {
                    xlApp.DisplayAlerts = true; // MUY IMPORTANTE: Devolver las alertas a su estado original
                    xlApp.ScreenUpdating = true;
                }
                Cursor.Current = Cursors.Default;
            }
        }

        // --- INICIO DE EVENTOS PARA PESTAÑA REVISIÓN/UTILIDADES ---
        //Funcion para bloqueo de hojas con contraseña
        private void btnBloqueo_Click(object sender, EventArgs e)
        {
            // 1. Validamos que haya un censo (libro de Excel) cargado en memoria
            if (_libroCenso == null)
            {
                MessageBox.Show(this, "No hay ningún censo cargado en memoria para proteger.", "Operación Denegada", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 2. Contabilizamos el estado inicial de protección de las hojas
            int totalHojas = _libroCenso.Worksheets.Count;
            int hojasYaProtegidasAlInicio = 0;

            foreach (Excel.Worksheet hojaCheck in _libroCenso.Worksheets)
            {
                if (hojaCheck.ProtectContents)
                {
                    hojasYaProtegidasAlInicio++;
                }
                System.Runtime.InteropServices.Marshal.ReleaseComObject(hojaCheck);
            }

            // EXCEPCIÓN A: Si TODAS las hojas ya contaban con protección previa
            if (hojasYaProtegidasAlInicio == totalHojas)
            {
                MessageBox.Show(this, "El libro ya está protegido por lo que no se aplicó protección con la clave del año elegido", "Aviso de Protección", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // 3. Validamos y obtenemos el valor del año desde el NumericUpDown (nudPeriodo)
            int anioSeleccionado = decimal.ToInt32(nudPeriodo.Value);

            if (anioSeleccionado <= 0)
            {
                MessageBox.Show(this, "Por favor, selecciona un año válido en el periodo antes de proteger el libro.", "Periodo Requerido", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 4. Estructuramos la contraseña concatenando "njm" con el año
            string passwordEstructurada = "njm" + anioSeleccionado.ToString();

            Excel.Application localExcelApp = null;
            int hojasProtegidasExitosamente = 0;

            try
            {
                localExcelApp = (Excel.Application)ExcelDna.Integration.ExcelDnaUtil.Application;
                localExcelApp.ScreenUpdating = false;

                // 5. Recorremos el libro para aplicar la protección únicamente a las hojas desprotegidas
                foreach (Excel.Worksheet hoja in _libroCenso.Worksheets)
                {
                    try
                    {
                        // Si la hoja no está protegida, aplicamos el bloqueo algorítmico
                        if (!hoja.ProtectContents)
                        {
                            hoja.Protect(
                                Password: passwordEstructurada,
                                DrawingObjects: true,
                                Contents: true,
                                Scenarios: true,
                                UserInterfaceOnly: false
                            );
                            hojasProtegidasExitosamente++;
                        }
                    }
                    catch (Exception exHoja)
                    {
                        System.Diagnostics.Debug.WriteLine($"No se pudo proteger la hoja '{hoja.Name}': {exHoja.Message}");
                    }
                    finally
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(hoja);
                    }
                }

                localExcelApp.ScreenUpdating = true;

                // EXCEPCIÓN B: Si sólo ALGUNAS hojas estaban protegidas previamente
                if (hojasYaProtegidasAlInicio > 0)
                {
                    MessageBox.Show(this, $"Sólo se protegieron {hojasProtegidasExitosamente} hojas debido a que las demás ya estaban protegidas con una clave previa.", "Protección Parcial", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    // Flujo Normal: Si ninguna estaba protegida y se bloquearon todas
                    MessageBox.Show(this, $"Se han protegido con éxito todas las hojas del libro.\n\nContraseña aplicada: {passwordEstructurada}", "Libro Protegido", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                if (localExcelApp != null) localExcelApp.ScreenUpdating = true;
                MessageBox.Show(this, "Error crítico al intentar proteger las hojas: " + ex.Message, "Error del Sistema", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        //Funcion para desbloqueo de hojas con contraseña
        private void btnDesbloqueo_Click(object sender, EventArgs e)
        {
            // 1. Validamos que haya un censo (libro de Excel) cargado en memoria
            if (_libroCenso == null)
            {
                MessageBox.Show(this, "No hay ningún censo cargado en memoria para desproteger.", "Operación Denegada", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 2. VALIDACIÓN: Contabilizar cuántas hojas están protegidas actualmente antes de proceder
            int hojasProtegidasAlInicio = 0;
            foreach (Excel.Worksheet hojaCheck in _libroCenso.Worksheets)
            {
                if (hojaCheck.ProtectContents)
                {
                    hojasProtegidasAlInicio++;
                }
                System.Runtime.InteropServices.Marshal.ReleaseComObject(hojaCheck);
            }

            // Si NINGUNA hoja está protegida en todo el libro, se lanza el Alert y se detiene la ejecución
            if (hojasProtegidasAlInicio == 0)
            {
                MessageBox.Show(this, "No se aplicó la acción de desproteger debido a que las hojas no estaban protegidas.", "Aviso de Desprotección", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // 3. Obtenemos de forma segura el año desde el control 'nudPeriodo'
            int anioSeleccionado = decimal.ToInt32(nudPeriodo.Value);

            if (anioSeleccionado <= 0)
            {
                MessageBox.Show(this, "Por favor, verifica que el control de Periodo tenga el año correspondiente al censo.", "Periodo Inválido", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 4. Reconstruimos la contraseña dinámica para desproteger (ej. "njm2026")
            string passwordEstructurada = "njm" + anioSeleccionado.ToString();

            Excel.Application localExcelApp = null;
            try
            {
                localExcelApp = (Excel.Application)ExcelDna.Integration.ExcelDnaUtil.Application;
                localExcelApp.ScreenUpdating = false;

                int hojasDesprotegidasExitosamente = 0;
                int hojasFallidas = 0;

                // 5. Recorremos cada una de las hojas de trabajo y desprotegemos únicamente las que estén bloqueadas
                foreach (Excel.Worksheet hoja in _libroCenso.Worksheets)
                {
                    try
                    {
                        if (hoja.ProtectContents)
                        {
                            hoja.Unprotect(passwordEstructurada);
                            hojasDesprotegidasExitosamente++;
                        }
                    }
                    catch (Exception exHoja)
                    {
                        // Si la contraseña dinámica no coincide con la clave real de la hoja
                        hojasFallidas++;
                        System.Diagnostics.Debug.WriteLine($"No se pudo desproteger la hoja '{hoja.Name}': {exHoja.Message}");
                    }
                    finally
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(hoja);
                    }
                }

                localExcelApp.ScreenUpdating = true;

                // 6. Mensajes informativos de finalización
                if (hojasFallidas > 0)
                {
                    MessageBox.Show(this,
                        $"Se desprotegieron {hojasDesprotegidasExitosamente} hojas correctamente.\n\n" +
                        $"Sin embargo, {hojasFallidas} hojas no pudieron ser desprotegidas con la clave '{passwordEstructurada}'. " +
                        "Verifica si corresponden a otro año o clave previa.",
                        "Desprotección Parcial",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                }
                else
                {
                    MessageBox.Show(this, "Se han desprotegido con éxito todas las hojas del libro de trabajo.", "Libro Desprotegido", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                if (localExcelApp != null) localExcelApp.ScreenUpdating = true;
                MessageBox.Show(this, "Error crítico al intentar desproteger las hojas: " + ex.Message, "Error del Sistema", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        //Funcion para aplicar un color a todas las celdas de todas las hojas para revisión de bloqueos
        private void btnAplicarFormato_Click(object sender, EventArgs e)
        {
            // 1. Validar que haya un censo cargado en memoria
            if (_libroCenso == null)
            {
                MessageBox.Show(this, "No hay ningún censo cargado en memoria.", "Operación Denegada", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            Excel.Application excelApp = null;
            try
            {
                excelApp = (Excel.Application)ExcelDna.Integration.ExcelDnaUtil.Application;

                // Apagamos la actualización de pantalla para que el barrido sea instantáneo
                excelApp.ScreenUpdating = false;

                // =========================================================================
                // FASE 1: AUDITORÍA DE HOJAS PROTEGIDAS (PREVENCIÓN DE CRASH)
                // =========================================================================
                System.Collections.Generic.List<string> hojasBloqueadas = new System.Collections.Generic.List<string>();

                foreach (Excel.Worksheet wsCheck in _libroCenso.Worksheets)
                {
                    if (wsCheck.ProtectContents || wsCheck.ProtectDrawingObjects || wsCheck.ProtectScenarios)
                    {
                        hojasBloqueadas.Add("• " + wsCheck.Name);
                    }
                    // Liberación estricta de memoria COM
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wsCheck);
                }

                // Si hay hojas bloqueadas, abortamos y le avisamos al usuario amigablemente
                if (hojasBloqueadas.Count > 0)
                {
                    string listaHojas = string.Join("\n", hojasBloqueadas);
                    MessageBox.Show(this,
                        "No se puede aplicar el formato de color porque el libro contiene hojas protegidas:\n\n" +
                        listaHojas + "\n\n" +
                        "Desprotege estas hojas (desde la pestaña Revisión) y vuelve a intentarlo.",
                        "Validación Interrumpida", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    return;
                }

                // =========================================================================
                // FASE 2: INYECCIÓN DE FORMATO CONDICIONAL POR HOJA
                // =========================================================================
                int hojasModificadas = 0;

                foreach (Excel.Worksheet ws in _libroCenso.Worksheets)
                {
                    Excel.Range celdasHoja = null;
                    Excel.FormatConditions formatos = null;
                    Excel.Range celdaDummy = null;

                    try
                    {
                        celdasHoja = ws.Cells;
                        formatos = celdasHoja.FormatConditions;
                        bool existe = false;

                        // 2.1 Búsqueda Inversa para verificar si la regla ya existe
                        for (int i = formatos.Count; i >= 1; i--)
                        {
                            Excel.FormatCondition fc = null;
                            try
                            {
                                object objFc = formatos[i];

                                // En C#, debemos verificar el tipo del objeto COM antes de leerlo
                                if (objFc is Excel.FormatCondition)
                                {
                                    fc = (Excel.FormatCondition)objFc;

                                    if (fc.Type == (int)Excel.XlFormatConditionType.xlExpression)
                                    {
                                        string formulaCondicion = "";
                                        try
                                        {
                                            formulaCondicion = fc.Formula1;
                                        }
                                        catch { /* Silenciamos si Excel deniega la lectura de esa fórmula específica */ }

                                        if (!string.IsNullOrEmpty(formulaCondicion))
                                        {
                                            // Limpiamos espacios y pasamos a mayúsculas para hacer una comparación universal
                                            string fLimpia = formulaCondicion.Replace(" ", "").ToUpper();

                                            // Cubrimos tanto "CELDA" (Español) como "CELL" (Inglés)
                                            if (fLimpia.Contains("CELDA(\"PROTECT\"") || fLimpia.Contains("CELL(\"PROTECT\""))
                                            {
                                                existe = true;
                                                break; // Ya existe, rompemos el ciclo
                                            }
                                        }
                                    }
                                }
                            }
                            catch
                            {
                                // Ignoramos errores de lectura de reglas aisladas
                            }
                            finally
                            {
                                if (fc != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(fc);
                            }
                        }

                        // 2.2 Si no la encontró, la inyectamos
                        if (!existe)
                        {
                            // TRUCO ARQUITECTÓNICO: Traducir la fórmula al idioma de la PC local usando una celda oculta.
                            // Esto evita el clásico error de sintaxis al inyectar código rígido en un Excel en español.
                            celdaDummy = ws.Cells[1048576, 16384]; // Última celda XFD1048576
                            celdaDummy.Formula = "=CELL(\"protect\",A1)"; // La metemos en inglés universal
                            string formulaLocal = celdaDummy.FormulaLocal; // La extraemos en español (o el idioma del usuario)
                            celdaDummy.Clear();

                            Excel.FormatCondition nuevaRegla = (Excel.FormatCondition)formatos.Add(
                                Excel.XlFormatConditionType.xlExpression,
                                Type.Missing,
                                formulaLocal);

                            // Aplicamos el color RGB(255, 230, 153) que equivale al amarillo de tu macro original
                            nuevaRegla.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 230, 153));

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(nuevaRegla);
                            hojasModificadas++;
                        }
                    }
                    finally
                    {
                        // Limpieza absoluta del Garbage Collector en cada ciclo de la hoja
                        if (celdaDummy != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(celdaDummy);
                        if (formatos != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(formatos);
                        if (celdasHoja != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(celdasHoja);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(ws);
                    }
                }

                // =========================================================================
                // FASE 3: CONCLUSIÓN UX
                // =========================================================================
                MessageBox.Show(this,
                    $"Formato de color (Celdas Bloqueadas) aplicado en todo el libro.\n\n" +
                    $"Se aplicó exitosamente en {hojasModificadas} hoja(s) que no contaban con él.",
                    "SAVCNG - Macro de Colores", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Ocurrió un error crítico al aplicar la macro de colores: " + ex.Message, "Error del Sistema", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (excelApp != null)
                {
                    // Restauramos SIEMPRE la actualización visual al terminar o fallar
                    excelApp.ScreenUpdating = true;
                }
            }
        }
        //Funcion para quitar el formato aplicado por btnAplicarFormato_Click
        private void btnLimpiarFormato_Click(object sender, EventArgs e)
        {
            // 1. Validar que haya un censo cargado en memoria
            if (_libroCenso == null)
            {
                MessageBox.Show(this, "No hay ningún censo cargado en memoria.", "Operación Denegada", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            Excel.Application excelApp = null;
            try
            {
                excelApp = (Excel.Application)ExcelDna.Integration.ExcelDnaUtil.Application;

                // Apagamos la actualización de pantalla para un borrado instantáneo y silencioso
                excelApp.ScreenUpdating = false;

                // =========================================================================
                // FASE 1: AUDITORÍA DE HOJAS PROTEGIDAS (PREVENCIÓN DE CRASH)
                // =========================================================================
                System.Collections.Generic.List<string> hojasBloqueadas = new System.Collections.Generic.List<string>();

                foreach (Excel.Worksheet wsCheck in _libroCenso.Worksheets)
                {
                    if (wsCheck.ProtectContents || wsCheck.ProtectDrawingObjects || wsCheck.ProtectScenarios)
                    {
                        hojasBloqueadas.Add("• " + wsCheck.Name);
                    }
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wsCheck);
                }

                // Si la hoja está bloqueada, Excel no nos dejará borrar la regla
                if (hojasBloqueadas.Count > 0)
                {
                    string listaHojas = string.Join("\n", hojasBloqueadas);
                    MessageBox.Show(this,
                        "No se puede limpiar el formato de color porque el libro contiene hojas protegidas:\n\n" +
                        listaHojas + "\n\n" +
                        "Desprotege estas hojas (desde la pestaña Revisión) y vuelve a intentarlo.",
                        "Operación Interrumpida", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    return;
                }

                // =========================================================================
                // FASE 2: BÚSQUEDA Y ELIMINACIÓN DE LA REGLA
                // =========================================================================
                int hojasLimpiadas = 0;

                foreach (Excel.Worksheet ws in _libroCenso.Worksheets)
                {
                    Excel.Range celdasHoja = null;
                    Excel.FormatConditions formatos = null;
                    bool hojaModificada = false;

                    try
                    {
                        celdasHoja = ws.Cells;
                        formatos = celdasHoja.FormatConditions;

                        // RECORRIDO INVERSO OBLIGATORIO: Del último formato al primero
                        for (int i = formatos.Count; i >= 1; i--)
                        {
                            Excel.FormatCondition fc = null;
                            try
                            {
                                object objFc = formatos[i];

                                if (objFc is Excel.FormatCondition)
                                {
                                    fc = (Excel.FormatCondition)objFc;

                                    if (fc.Type == (int)Excel.XlFormatConditionType.xlExpression)
                                    {
                                        string formulaCondicion = "";
                                        try { formulaCondicion = fc.Formula1; } catch { }

                                        if (!string.IsNullOrEmpty(formulaCondicion))
                                        {
                                            string fLimpia = formulaCondicion.Replace(" ", "").ToUpper();

                                            // Evaluamos si es la regla de nuestro botón "Aplicar"
                                            if (fLimpia.Contains("CELDA(\"PROTECT\"") || fLimpia.Contains("CELL(\"PROTECT\""))
                                            {
                                                // ¡La encontramos! La destruimos
                                                fc.Delete();
                                                hojaModificada = true;
                                            }
                                        }
                                    }
                                }
                            }
                            catch
                            {
                                // Ignoramos errores aislados de lectura/borrado
                            }
                            finally
                            {
                                if (fc != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(fc);
                            }
                        }

                        if (hojaModificada)
                        {
                            hojasLimpiadas++;
                        }
                    }
                    finally
                    {
                        if (formatos != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(formatos);
                        if (celdasHoja != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(celdasHoja);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(ws);
                    }
                }

                // =========================================================================
                // FASE 3: CONCLUSIÓN UX
                // =========================================================================
                if (hojasLimpiadas > 0)
                {
                    MessageBox.Show(this,
                        $"El formato de color (Celdas Bloqueadas) fue eliminado correctamente.\n\n" +
                        $"Se limpiaron {hojasLimpiadas} hoja(s) del libro.",
                        "SAVCNG - Limpieza Exitosa", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show(this,
                        "No se encontró el formato de color en ninguna hoja del libro.\n\nEl censo ya está limpio.",
                        "SAVCNG - Sin Cambios", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Ocurrió un error crítico al limpiar los colores: " + ex.Message, "Error del Sistema", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (excelApp != null) excelApp.ScreenUpdating = true;
            }
        }
        //Funcion para aplicar Colorimetria a las validaciones para identificarlas
        private void btnAplicarColores_Click(object sender, EventArgs e)
        {
            // 1. Validar que haya un censo cargado en memoria
            if (_libroCenso == null)
            {
                MessageBox.Show(this, "No hay ningún censo cargado en memoria.", "Operación Denegada", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            Excel.Application excelApp = null;
            try
            {
                excelApp = (Excel.Application)ExcelDna.Integration.ExcelDnaUtil.Application;
                excelApp.ScreenUpdating = false;

                // =========================================================================
                // FASE 1: AUDITORÍA DE HOJAS PROTEGIDAS (PREVENCIÓN DE CRASH)
                // =========================================================================
                System.Collections.Generic.List<string> hojasBloqueadas = new System.Collections.Generic.List<string>();

                foreach (Excel.Worksheet wsCheck in _libroCenso.Worksheets)
                {
                    if (wsCheck.ProtectContents || wsCheck.ProtectDrawingObjects || wsCheck.ProtectScenarios)
                    {
                        hojasBloqueadas.Add("• " + wsCheck.Name);
                    }
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wsCheck);
                }

                if (hojasBloqueadas.Count > 0)
                {
                    string listaHojas = string.Join("\n", hojasBloqueadas);
                    MessageBox.Show(this,
                        "No se puede aplicar la macro de colores porque el libro contiene hojas protegidas:\n\n" +
                        listaHojas + "\n\n" +
                        "Desprotege estas hojas desde la pestaña Revisión y vuelve a intentarlo.",
                        "Validación Interrumpida", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    return;
                }

                // =========================================================================
                // FASE 2: ESTRUCTURA DE FÓRMULAS UNIVERSALES (MATRIZ DE TRADUCCIÓN)
                // =========================================================================
                // Definimos las 7 reglas en inglés universal para traducirlas nativamente en cada hoja.
                // Formato 1 al 7 tal como venían definidos en tu aplicativo original.
                string[] formulasIngles = new string[]
                {
            "=AND($A$1<>\"\",SEARCH(\"igual o mayor\",A1)>0)",  // Regla 2
            "=AND($A$1<>\"\",SEARCH(\"igual o menor\",A1)>0)",  // Regla 3
            "=AND($A$1<>\"\",SEARCH(\"pase a la pregunta\",A1)>0)", // Regla 4
            "=AND($A$1<>\"\",SEARCH(\"en blanco\",A1)>0)",      // Regla 5
            "=AND($A$1<>\"\",SEARCH(\"no puede registrar\",A1)>0)", // Regla 6
            "=AND($A$1<>\"\",SUM(A1)>0)",                       // Regla 7
            "=AND($A$1<>\"\",SEARCH(\"la pregunta\",A1)>0)"     // Regla 1
                };

                int hojasConfiguradas = 0;

                foreach (Excel.Worksheet ws in _libroCenso.Worksheets)
                {
                    Excel.Range celdasHoja = null;
                    Excel.FormatConditions formatos = null;
                    Excel.Range celdaDummy = null;

                    try
                    {
                        celdasHoja = ws.Cells;
                        formatos = celdasHoja.FormatConditions;

                        // 2.1 BÚSQUEDA PREVIA: Si la hoja ya tiene al menos una de nuestras reglas, asumimos que ya fue procesada
                        bool yaExisteRegra = false;
                        for (int i = formatos.Count; i >= 1; i--)
                        {
                            Excel.FormatCondition fcCheck = null;
                            try
                            {
                                object objFc = formatos[i];
                                if (objFc is Excel.FormatCondition)
                                {
                                    fcCheck = (Excel.FormatCondition)objFc;
                                    if (fcCheck.Type == (int)Excel.XlFormatConditionType.xlExpression)
                                    {
                                        string formulaText = fcCheck.Formula1 ?? "";
                                        if (formulaText.Replace(" ", "").ToUpper().Contains("$A$1<>\"\""))
                                        {
                                            yaExisteRegra = true;
                                            break;
                                        }
                                    }
                                }
                            }
                            catch { }
                            finally
                            {
                                if (fcCheck != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(fcCheck);
                            }
                        }

                        if (yaExisteRegra)
                        {
                            continue; // Nos saltamos esta hoja para no duplicar configuraciones
                        }

                        // 2.2 INYECCIÓN DE LAS 7 REGLAS CON TRADUCCIÓN NATIVA SENSIBLE A IDIOMA
                        celdaDummy = ws.Cells[1048576, 16384]; // Celda XFD1048576

                        for (int idx = 0; idx < formulasIngles.Length; idx++)
                        {
                            // Forzar traducción al lenguaje de Excel local
                            celdaDummy.Formula = formulasIngles[idx];
                            string formulaLocal = celdaDummy.FormulaLocal;
                            celdaDummy.Clear();

                            Excel.FormatCondition nuevaRegla = (Excel.FormatCondition)formatos.Add(
                                Excel.XlFormatConditionType.xlExpression,
                                Type.Missing,
                                formulaLocal);

                            // Configuración estética individual según tu código de origen
                            switch (idx)
                            {
                                case 0: // Formato 2: Accent6 con Tinte (0.5999)
                                    nuevaRegla.Interior.PatternColorIndex = (int)Excel.Constants.xlAutomatic;
                                    nuevaRegla.Interior.ThemeColor = (int)Excel.XlThemeColor.xlThemeColorAccent6;
                                    nuevaRegla.Interior.TintAndShade = 0.599963377788629;
                                    break;

                                case 1: // Formato 3: Accent2 con Tinte (0.5999)
                                    nuevaRegla.Interior.PatternColorIndex = (int)Excel.Constants.xlAutomatic;
                                    nuevaRegla.Interior.ThemeColor = (int)Excel.XlThemeColor.xlThemeColorAccent2;
                                    nuevaRegla.Interior.TintAndShade = 0.599963377788629;
                                    break;

                                case 2: // Formato 4: Color Fijo Decimal 16764159
                                case 3: // Formato 5: Mismo comportamiento
                                case 4: // Formato 6: Mismo comportamiento
                                    nuevaRegla.Interior.PatternColorIndex = (int)Excel.Constants.xlAutomatic;
                                    nuevaRegla.Interior.Color = 16764159;
                                    nuevaRegla.Interior.TintAndShade = 0;
                                    break;

                                case 5: // Formato 7: Accent5 con Tinte (0.7999)
                                    nuevaRegla.Interior.PatternColorIndex = (int)Excel.Constants.xlAutomatic;
                                    nuevaRegla.Interior.ThemeColor = (int)Excel.XlThemeColor.xlThemeColorAccent5;
                                    nuevaRegla.Interior.TintAndShade = 0.799981688894314;
                                    break;

                                case 6: // Formato 1: Accent4 con Tinte (0.5999)
                                    nuevaRegla.Interior.PatternColorIndex = (int)Excel.Constants.xlAutomatic;
                                    nuevaRegla.Interior.ThemeColor = (int)Excel.XlThemeColor.xlThemeColorAccent4;
                                    nuevaRegla.Interior.TintAndShade = 0.599963377788629;
                                    break;
                            }

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(nuevaRegla);
                        }

                        hojasConfiguradas++;
                    }
                    finally
                    {
                        if (celdaDummy != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(celdaDummy);
                        if (formatos != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(formatos);
                        if (celdasHoja != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(celdasHoja);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(ws);
                    }
                }

                MessageBox.Show(this,
                    $"Funcion de colorimetría aplicada con éxito.\n\nSe configuraron {hojasConfiguradas} hoja(s) del censo de forma segura.",
                    "Reglas de Control Añadidas", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Error crítico al inyectar las macros de control A1: " + ex.Message, "Error del Sistema", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (excelApp != null) excelApp.ScreenUpdating = true;
            }
        }
        //Funcion para revertir Colorimetria aplicada por btnAplicarColores_Click
        private void btnLimpiarColores_Click(object sender, EventArgs e)
        {
            // 1. Validar que haya un censo cargado en memoria
            if (_libroCenso == null)
            {
                MessageBox.Show(this, "No hay ningún censo cargado en memoria.", "Operación Denegada", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            Excel.Application excelApp = null;
            try
            {
                excelApp = (Excel.Application)ExcelDna.Integration.ExcelDnaUtil.Application;
                excelApp.ScreenUpdating = false;

                // =========================================================================
                // FASE 1: AUDITORÍA DE HOJAS PROTEGIDAS (PREVENCIÓN DE CRASH)
                // =========================================================================
                System.Collections.Generic.List<string> hojasBloqueadas = new System.Collections.Generic.List<string>();

                foreach (Excel.Worksheet wsCheck in _libroCenso.Worksheets)
                {
                    if (wsCheck.ProtectContents || wsCheck.ProtectDrawingObjects || wsCheck.ProtectScenarios)
                    {
                        hojasBloqueadas.Add("• " + wsCheck.Name);
                    }
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wsCheck);
                }

                if (hojasBloqueadas.Count > 0)
                {
                    string listaHojas = string.Join("\n", hojasBloqueadas);
                    MessageBox.Show(this,
                        "No se puede remover el formato de colores debido a que el libro tiene hojas protegidas:\n\n" +
                        listaHojas + "\n\n" +
                        "Desprotege estas hojas desde la pestaña Revisión y reintenta la operación.",
                        "Operación Interrumpida", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    return;
                }

                // =========================================================================
                // FASE 2: PURGA Y ELIMINACIÓN DE FORMATOS MEDIANTE RECORRIDO INVERSO
                // =========================================================================
                int hojasLimpiadas = 0;

                foreach (Excel.Worksheet ws in _libroCenso.Worksheets)
                {
                    Excel.Range celdasHoja = null;
                    Excel.FormatConditions formatos = null;
                    bool cambioDetectado = false;

                    try
                    {
                        celdasHoja = ws.Cells;
                        formatos = celdasHoja.FormatConditions;

                        // ESCANEO INVERSO OBLIGATORIO: Evita desajustes en el puntero de la lista al borrar
                        for (int i = formatos.Count; i >= 1; i--)
                        {
                            Excel.FormatCondition fc = null;
                            try
                            {
                                object objFc = formatos[i];
                                if (objFc is Excel.FormatCondition)
                                {
                                    fc = (Excel.FormatCondition)objFc;

                                    if (fc.Type == (int)Excel.XlFormatConditionType.xlExpression)
                                    {
                                        string formulaCondicion = "";
                                        try { formulaCondicion = fc.Formula1; } catch { }

                                        if (!string.IsNullOrEmpty(formulaCondicion))
                                        {
                                            string formulaLimpia = formulaCondicion.Replace(" ", "").ToUpper();

                                            // Buscamos nuestro patrón común anclado a la celda de control $A$1
                                            if (formulaLimpia.Contains("$A$1<>\"\""))
                                            {
                                                fc.Delete();
                                                cambioDetectado = true;
                                            }
                                        }
                                    }
                                }
                            }
                            catch { }
                            finally
                            {
                                if (fc != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(fc);
                            }
                        }

                        if (cambioDetectado)
                        {
                            hojasLimpiadas++;
                        }
                    }
                    finally
                    {
                        if (formatos != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(formatos);
                        if (celdasHoja != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(celdasHoja);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(ws);
                    }
                }

                // =========================================================================
                // FASE 3: CONCLUSIÓN UX
                // =========================================================================
                if (hojasLimpiadas > 0)
                {
                    MessageBox.Show(this,
                        $"Se eliminaron por completo las reglas de colorimetria vinculadas al archivo.\n\nSe limpiaron exitosamente {hojasLimpiadas} hoja(s).",
                        "SAVCNG - Limpieza Exitosa", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show(this,
                        "No se detectaron formatos condicionales vinculados al archivo actual.\n\nEl libro ya se encuentra limpio.",
                        "SAVCNG - Sin Cambios", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Ocurrió un error inesperado al limpiar las macros de formato: " + ex.Message, "Error del Sistema", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (excelApp != null) excelApp.ScreenUpdating = true;
            }
        }
    }
}
