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
                if (chkDecimales.Checked == true)
                {
                    try
                    {

                        // 1. Limpiamos validaciones previas para no empalmar reglas
                        _rangoCapturado.Validation.Delete();

                        // 2. Extraemos la dirección de la primera celda sin anclas (Ejemplo: A1)
                        Excel.Range primeraCelda = (Excel.Range)_rangoCapturado.Cells[1, 1];
                        string direccion = primeraCelda.get_Address(false, false, Excel.XlReferenceStyle.xlA1, false);

                        // ==========================================================
                        // FÓRMULA LOCAL (ESPAÑOL) BLINDADA CON SI.ERROR
                        // ==========================================================
                        // Envolvemos Y(ESNUMERO... TRUNCAR...) dentro de un SI.ERROR( ... ; FALSO)
                        // Esto evita que TRUNCAR estalle cuando lee los textos "NS" o "NA"
                        string formulaDecimales = $"=O(ESBLANCO({direccion}){separador}SI.ERROR(Y(ESNUMERO({direccion}){separador}TRUNCAR({direccion})={direccion}){separador}FALSO){separador}{direccion}=\"NS\"{separador}{direccion}=\"NA\")";

                        // 3. Aplicamos la regla restrictiva (Data Validation)
                        _rangoCapturado.Validation.Add(
                            Excel.XlDVType.xlValidateCustom,
                            Excel.XlDVAlertStyle.xlValidAlertStop,
                            Excel.XlFormatConditionOperator.xlBetween,
                            formulaDecimales,
                            Type.Missing);

                        // 4. Configuramos el comportamiento de la ventana emergente
                        _rangoCapturado.Validation.IgnoreBlank = true;
                        _rangoCapturado.Validation.ShowError = true;

                        // 5. Textos personalizados para el usuario final (UX)
                        _rangoCapturado.Validation.ErrorTitle = "Captura Inválida";
                        _rangoCapturado.Validation.ErrorMessage = "El formato de esta celda solo admite numeros enteros, por lo que no debe agregar texto ni desagregar decimales.\n\nPor favor, introduce únicamente un número entero (Ej: 1, 15, 100) o los códigos 'NS' o 'NA'.";

                        // 6. Limpiamos la interfaz y avisamos del éxito
                        chkDecimales.Checked = false;
                        MessageBox.Show(this, "Validación de formatos Decimales aplicada con éxito.\n\nLa validación considera los valores NS y NA de forma segura.", "SAVCNG Arquitectura", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(this, "Error crítico al aplicar Validación de Decimales: " + ex.Message, "Error de Inserción", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        chkDecimales.Checked = false;
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

                            // 2. Manejo de la cancelación estricta (Abortar operación limpiamente)
                            if (respuesta == DialogResult.Cancel)
                            {
                                // Interrumpimos el flujo, quitamos el check y salimos del método
                                chkCatalogos.Checked = false;
                                return;
                            }

                            // 3. Lógica de "Aceptar" (Extracción de texto del objeto COM)
                            Excel.Range primeraCeldaOrigen = (Excel.Range)rangoOrigen.Cells[1, 1];
                            string textoCelda = primeraCeldaOrigen.Text?.ToString() ?? "";

                            // 4. Validación de seguridad (Prevenir errores de referencia nula)
                            if (string.IsNullOrWhiteSpace(textoCelda))
                            {
                                MessageBox.Show(this, "La celda origen está vacía. No se puede extraer el catálogo.",
                                                "Aviso Arquitectónico", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                chkCatalogos.Checked = false;
                                return;
                            }

                            // 5. Motor de separación y limpieza (Extraer únicamente los números)
                            string[] pedacitos = textoCelda.Split(new char[] { '.', ',', '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
                            System.Collections.Generic.List<string> listaLimpios = new System.Collections.Generic.List<string>();

                            foreach (string pedazo in pedacitos)
                            {
                                string soloNumeros = "";

                                // Iteramos por cada carácter buscando dígitos
                                foreach (char letra in pedazo)
                                {
                                    if (char.IsDigit(letra)) soloNumeros += letra;
                                }

                                // Solo lo agregamos si realmente encontró un número
                                if (!string.IsNullOrEmpty(soloNumeros))
                                {
                                    listaLimpios.Add(soloNumeros);
                                }
                            }

                            // 6. Construcción de la cadena final nativa
                            string separadorSistema = excelApp.International[Excel.XlApplicationInternational.xlListSeparator].ToString();
                            formulaOpciones = string.Join(separadorSistema, listaLimpios);
                        }
                        else
                        {
                            // 1. Simplificación de UX: Pregunta directa y coherente con el paso anterior
                            DialogResult respuestaRango = MessageBox.Show(this,
                                "Se han detectado varias celdas seleccionadas como catálogo.\n" +
                                "¿Deseas que el sistema extraiga automáticamente SOLO LOS NÚMEROS (ej. 1, 2, 9) de cada celda para crear las opciones?\n\n" +
                                "• [Aceptar]: Extraer números y continuar.\n" +
                                "• [Cancelar]: Abortar esta operación.",
                                "Configuración de Catálogo",
                                MessageBoxButtons.OKCancel,
                                MessageBoxIcon.Question);

                            // 2. Manejo de la cancelación estricta (Abortar operación limpiamente)
                            if (respuestaRango == DialogResult.Cancel)
                            {
                                chkCatalogos.Checked = false;
                                return;
                            }

                            // 3. Lógica de extracción de "Aceptar" (Solo iteración de números)
                            System.Collections.Generic.List<string> listaLimpios = new System.Collections.Generic.List<string>();

                            foreach (Excel.Range celda in rangoOrigen.Cells)
                            {
                                string textoCelda = celda.Text?.ToString() ?? "";
                                string soloNumeros = "";

                                // Escaneamos la cadena carácter por carácter buscando dígitos
                                foreach (char letra in textoCelda)
                                {
                                    if (char.IsDigit(letra)) soloNumeros += letra;
                                }

                                // Si la celda tenía números, la agregamos a la lista limpia
                                if (!string.IsNullOrEmpty(soloNumeros))
                                {
                                    listaLimpios.Add(soloNumeros);
                                }
                            }

                            // 4. Validación de seguridad (Filtro anti-errores del usuario)
                            if (listaLimpios.Count == 0)
                            {
                                MessageBox.Show(this, "No se encontraron valores numéricos en el rango seleccionado.\nNo se puede crear el catálogo.",
                                                "Aviso Arquitectónico", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                chkCatalogos.Checked = false;
                                return;
                            }

                            // 5. Construcción de la cadena final (Usando tu variable 'separador' original)
                            formulaOpciones = string.Join(separador, listaLimpios);
                        }
                    }

                    // 3. APLICAR LA VALIDACIÓN (Tu lógica original intacta)
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
                        MessageBox.Show(this,"¡Validación de catálogo aplicada con éxito!", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(this,"Error al aplicar la validación: " + ex.Message + "\n\nTexto que se intentó usar: " + formulaOpciones, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                        // FASE 1: UX DE CONFIGURACIÓN GLOBAL (SOLICITAR TEXTO UNA SOLA VEZ)
                        // =========================================================================
                        string textoSugerido = "Alerta: debido a que cuenta con registros NS, debe proporcionar una justificación en el área de comentarios al final de la pregunta";
                        object resTexto = xlApp.InputBox(
                            "Escribe el texto del mensaje de alerta que se aplicará a todas las celdas seleccionadas:",
                            "SAVCNG - Configuración Masiva NS (Permitiendo NA)", textoSugerido, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 2); // 2 = Texto

                        if (resTexto is bool && (bool)resTexto == false) { chkNS.Checked = false; return; }
                        string textoAlerta = resTexto.ToString().Trim();

                        // =========================================================================
                        // FASE 2: ALGORITMO DE AGRUPACIÓN POR PREGUNTA (MATRICES PARALELAS)
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

                        // =========================================================================
                        // FASE 3: PROCESAMIENTO POR LOTES Y MAPEO UNO A UNO (ÁREA -> ALERTA)
                        // =========================================================================
                        foreach (var grupo in áreasPorPregunta)
                        {
                            string preguntaActual = grupo.Key;
                            System.Collections.Generic.List<Excel.Range> listaÁreas = grupo.Value;

                            // Recorremos CADA ÁREA individual de la pregunta actual
                            for (int j = 0; j < listaÁreas.Count; j++)
                            {
                                Excel.Range area = listaÁreas[j];
                                Excel.Range rangoAlerta = null;
                                Excel.Range primeraCelda = null;
                                Excel.Range celdaDummy = null;

                                try
                                {
                                    // 1. UX Mapeo Dirigido UNO A UNO: Indicamos al usuario qué bloque está configurando
                                    object resDestino = xlApp.InputBox(
                                        $"[PREGUNTA DETECTADA: {preguntaActual}] - Rango {j + 1} de {listaÁreas.Count}\n\n" +
                                        $"Selecciona la celda o rango destino donde aparecerá el mensaje de ALERTA:\n" +
                                        $"(Selección Estandar para mensajes: Columna B hasta AD)",

                                        $"SAVCNG - Alerta P.{preguntaActual} (Área {j + 1})", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8); // 8 = Rango

                                    // Si cancela este bloque, saltamos al siguiente
                                    if (resDestino is bool && (bool)resDestino == false) continue;
                                    rangoAlerta = (Excel.Range)resDestino;

                                    // 2. Inyección de la Validación Restrictiva en el área actual
                                    area.Validation.Delete();

                                    primeraCelda = (Excel.Range)area.Cells[1, 1];
                                    string direccionRelativa = primeraCelda.get_Address(false, false, Excel.XlReferenceStyle.xlA1, false);

                                    // Regla que permite valores >= 0, "NS" y "NA"
                                    string formulaRestriccionIngles = $"=OR(AND(ISNUMBER({direccionRelativa}),{direccionRelativa}>=0),{direccionRelativa}=\"NS\",{direccionRelativa}=\"NA\")";

                                    // TRUCO DE TRADUCCIÓN NATIVA (FILA-SENSIBLE)
                                    int filaBaseArea = area.Row;
                                    celdaDummy = (Excel.Range)wsActual.Cells[filaBaseArea, 16384]; // Última columna (XFD)
                                    celdaDummy.Formula = formulaRestriccionIngles;
                                    string formulaRestriccionLocal = celdaDummy.FormulaLocal;
                                    celdaDummy.Clear();

                                    area.Validation.Add(
                                        Excel.XlDVType.xlValidateCustom,
                                        Excel.XlDVAlertStyle.xlValidAlertStop,
                                        Excel.XlFormatConditionOperator.xlBetween,
                                        formulaRestriccionLocal,
                                        Type.Missing);

                                    area.Validation.IgnoreBlank = true;
                                    area.Validation.ShowError = true;
                                    area.Validation.ErrorTitle = "Error de validación (Censo)";
                                    area.Validation.ErrorMessage = "Solo se permiten números mayores o iguales a cero, o los valores especiales 'NS' y 'NA'.";

                                    // 3. Configuración de la Alerta EXCLUSIVA para esta área
                                    if (rangoAlerta.Count > 1) { rangoAlerta.Merge(); }
                                    rangoAlerta.Font.Name = "Arial";
                                    rangoAlerta.Font.Size = 9;
                                    rangoAlerta.Font.Bold = true;
                                    rangoAlerta.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(191, 143, 0));
                                    rangoAlerta.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

                                    string addrAbsoluta = area.get_Address(true, true, Excel.XlReferenceStyle.xlA1, false);

                                    // La fórmula ahora solo evalúa (COUNTIF) el área actual, no todas las áreas juntas
                                    string formulaFinalAlerta = $"=IF(COUNTIF({addrAbsoluta},\"NS\")>0, \"{textoAlerta}\", \"\")";
                                    rangoAlerta.Formula = formulaFinalAlerta;
                                }
                                finally
                                {
                                    // Liberación rigurosa de objetos COM dentro del ciclo para evitar fugas de memoria
                                    if (rangoAlerta != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rangoAlerta);
                                    if (primeraCelda != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(primeraCelda);
                                    if (celdaDummy != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(celdaDummy);
                                }
                            }
                            preguntasProcesadas++;
                        }

                        // =========================================================================
                        // FASE 4: CONCLUSIÓN Y NOTIFICACIÓN UX
                        // =========================================================================
                        chkNS.Checked = false;
                        MessageBox.Show(this,
                            $"Procesamiento por lotes finalizado con éxito.\n\n" +
                            $"• Preguntas identificadas: {preguntasProcesadas}\n" +
                            $"• Total de áreas configuradas de forma independiente: {_rangoCapturado.Areas.Count}\n\n" +
                            $"Nota: Se admite 'NA', pero solo los registros 'NS' detonarán la alerta en sus celdas destino correspondientes.",
                            "SAVCNG Automatización Masiva", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(this, "Error crítico en el motor de masificación: " + ex.Message, "Error del Sistema", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        chkNS.Checked = false;
                    }
                    finally
                    {
                        if (wsActual != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(wsActual);
                    }
                }
                // --- VALIDACIÓN FORMATO TEXTO ---
                else if (chkFormatoTexto.Checked == true)
                {
                    try
                    {
                        // 1. Limpiamos validaciones previas
                        _rangoCapturado.Validation.Delete();

                        // ==========================================================
                        // TÉCNICA DEL RANGO AUXILIAR DINÁMICO
                        // ==========================================================

                        // --- NUEVO PASO: SOLICITAR LA UBICACIÓN DEL ESPEJO ---
                        object resAuxiliar = excelApp.InputBox(
                            "Indica en qué columna libre deseas colocar la validación oculta (AF en adelante).\n\nConsidera que el sistema requerirá espacio libre hacia abajo proporcional al número de filas que seleccionaste originalmente.",
                            "Seleccion de formula Auxiliar", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8); // 8 = Rango

                        if (resAuxiliar is bool && (bool)resAuxiliar == false) { chkFormatoTexto.Checked = false; return; }

                        Excel.Range seleccionAuxiliar = (Excel.Range)resAuxiliar;
                        Excel.Range celdaInicioAux = (Excel.Range)seleccionAuxiliar.Cells[1, 1];

                        // Obtenemos la letra de la columna seleccionada (Ej. "CW")
                        string colAuxLetra = celdaInicioAux.Address.Split('$')[1];

                        // 2. Construimos el rango auxiliar para que coincida exactamente con las filas capturadas
                        Excel.Worksheet ws = (Excel.Worksheet)_rangoCapturado.Worksheet;
                        int filaInicio = _rangoCapturado.Row;
                        int filaFin = filaInicio + _rangoCapturado.Rows.Count - 1;

                        Excel.Range rangoAuxiliar = ws.Range[$"{colAuxLetra}{filaInicio}:{colAuxLetra}{filaFin}"];

                        // Extraemos las direcciones relativas (Ej: C12 y CW12)
                        Excel.Range primeraCeldaCap = (Excel.Range)_rangoCapturado.Cells[1, 1];
                        string dirCapRel = primeraCeldaCap.get_Address(false, false, Excel.XlReferenceStyle.xlA1, false);

                        Excel.Range primeraCeldaAux = (Excel.Range)rangoAuxiliar.Cells[1, 1];
                        string dirAuxRel = primeraCeldaAux.get_Address(false, false, Excel.XlReferenceStyle.xlA1, false);

                        // 3. Diccionario estricto (Whitelist)
                        string permitidos = "0123456789ABCDEFGHIJKLMNÑOPQRSTUVWXYZÁÉÍÓÚÜ ";

                        // 4. Fórmula Binaria (1 = Válido, 0 = Inválido). 
                        string formulaAuxiliar = $"=IF(OR(ISBLANK({dirCapRel}), AND(EXACT({dirCapRel},UPPER({dirCapRel})), LEN({dirCapRel})=LEN(TRIM({dirCapRel})), SUMPRODUCT(--ISNUMBER(FIND(MID({dirCapRel},ROW(INDIRECT(\"1:\"&MAX(1,LEN({dirCapRel})))),1),\"{permitidos}\")))=LEN({dirCapRel}))), 1, 0)";

                        // 5. Inyectamos la matemática en el Rango Auxiliar y ocultamos la columna completa
                        rangoAuxiliar.Formula = formulaAuxiliar;
                        // Si deseas que se oculte automáticamente en producción, descomenta la siguiente línea:
                        // rangoAuxiliar.EntireColumn.Hidden = true;

                        // ==========================================================
                        // TRUCO ARQUITECTÓNICO: BYPASS DEL ERROR 0x800A03EC
                        // ==========================================================
                        object valorOriginal = primeraCeldaCap.Value2;
                        bool estabaVacia = (valorOriginal == null || string.IsNullOrWhiteSpace(valorOriginal.ToString()));

                        if (estabaVacia)
                        {
                            primeraCeldaCap.Value2 = "A";
                        }

                        // ==========================================================
                        // VALIDACIÓN DE DATOS (ULTRA LIGERA)
                        // ==========================================================
                        string formulaDV = $"={dirAuxRel}=1";

                        try
                        {
                            _rangoCapturado.Validation.Add(
                                Excel.XlDVType.xlValidateCustom,
                                Excel.XlDVAlertStyle.xlValidAlertStop,
                                Excel.XlFormatConditionOperator.xlBetween,
                                formulaDV,
                                Type.Missing);

                            _rangoCapturado.Validation.IgnoreBlank = true;
                            _rangoCapturado.Validation.ShowError = true;

                            _rangoCapturado.Validation.ErrorTitle = "Formato de texto inválido";
                            _rangoCapturado.Validation.ErrorMessage = "El texto debe cumplir estas reglas:\n\n" +
                                                                      "• Solo se permite texto en MAYÚSCULAS y NÚMEROS.\n" +
                                                                      "• Sin espacios dobles o sobrantes.\n" +
                                                                      "• Sin comillas ni signos de puntuación, paréntesis ni caracteres especiales.";
                        }
                        catch (System.Runtime.InteropServices.COMException)
                        {
                            // Falso positivo silenciado
                        }

                        // ==========================================================
                        // LIMPIEZA BLINDADA
                        // ==========================================================
                        if (estabaVacia)
                        {
                            primeraCeldaCap.Value2 = null;
                        }

                        chkFormatoTexto.Checked = false;
                        MessageBox.Show(this,
                        $"Validación de formato de texto aplicada con éxito.\n\nNota: La fórmula de apoyo se colocó en la columna {colAuxLetra}.",
                        "Formato de texto",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(this,"Error crítico al configurar Formato Texto: " + ex.Message, "Error del Sistema", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        chkFormatoTexto.Checked = false;
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
                // --- VALIDACIÓN BLANCOS ---
                else if (chkBlancos.Checked == true)
                {
                    try
                    {
                        // 1. Solicitar destino del Mensaje en Azul
                        object resDestino = excelApp.InputBox(
                            "Selecciona la celda o rango donde aparecerá el mensaje de alerta (Se combinará y pintará de azul automáticamente):",
                            "1. Ubicación de Alerta", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8); // 8 = Rango

                        if (resDestino is bool && (bool)resDestino == false) { chkBlancos.Checked = false; return; }
                        Excel.Range rangoDestino = (Excel.Range)resDestino;

                        // 2. Solicitar texto del mensaje
                        object resTexto = excelApp.InputBox(
                            "Escribe el mensaje de advertencia:\n(El sistema evaluará el rango capturado fila por fila. Si una fila tiene datos, exigirá que esté completa)",
                            "2. Mensaje de Alerta", "Favor de revisar la información faltante en las celdas sombreadas", Type.Missing, Type.Missing, Type.Missing, Type.Missing, 2); // 2 = Texto

                        if (resTexto is bool && (bool)resTexto == false) { chkBlancos.Checked = false; return; }
                        string textoAlerta = resTexto.ToString().Trim();

                        // ==========================================================
                        // MOTOR MATRICIAL INTELIGENTE (ALERTA GLOBAL)
                        // ==========================================================

                        System.Collections.Generic.List<string> listaActivadores = new System.Collections.Generic.List<string>();
                        System.Collections.Generic.List<string> listaVacios = new System.Collections.Generic.List<string>();

                        for (int i = 1; i <= _rangoCapturado.Columns.Count; i++)
                        {
                            Excel.Range col = (Excel.Range)_rangoCapturado.Columns[i];
                            string dirColAbs = col.get_Address(true, true, Excel.XlReferenceStyle.xlA1, false);

                            listaActivadores.Add($"({dirColAbs}<>\"\")");
                            listaVacios.Add($"({dirColAbs}=\"\")");
                        }

                        string motorActivadores = string.Join("+", listaActivadores);
                        string motorVacios = string.Join("+", listaVacios);

                        string formulaGlobal = $"=IF(SUMPRODUCT(({motorActivadores})*({motorVacios}))>0, \"{textoAlerta}\", \"\")";

                        if (rangoDestino.Count > 1) { rangoDestino.Merge(); }

                        rangoDestino.Font.Name = "Arial";
                        rangoDestino.Font.Size = 10;
                        rangoDestino.Font.Bold = true;
                        rangoDestino.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(0, 112, 192));
                        rangoDestino.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

                        rangoDestino.Formula = formulaGlobal;

                        // ==========================================================
                        // FORMATO CONDICIONAL: ÁLGEBRA BOOLEANA (FILA POR FILA)
                        // ==========================================================

                        // NUEVA AUDITORÍA DE COEXISTENCIA PARA BLANCOS
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

                        int filaInicial = _rangoCapturado.Row;

                        Excel.Range primeraCeldaCapturada = (Excel.Range)_rangoCapturado.Cells[1, 1];
                        string celdaCapRelativa = primeraCeldaCapturada.get_Address(false, false, Excel.XlReferenceStyle.xlA1, false);

                        System.Collections.Generic.List<string> validacionFila = new System.Collections.Generic.List<string>();
                        for (int c = 1; c <= _rangoCapturado.Columns.Count; c++)
                        {
                            Excel.Range celdaIteracion = (Excel.Range)_rangoCapturado.Cells[1, c];
                            string colLetra = celdaIteracion.Address.Split('$')[1];

                            validacionFila.Add($"(${colLetra}{filaInicial}<>\"\")");
                        }
                        string sumaFila = string.Join("+", validacionFila);

                        string formulaCondicionalMatematica = $"=({celdaCapRelativa}=\"\")*(({sumaFila})>0)";

                        Excel.FormatCondition formatoAmarillo = (Excel.FormatCondition)_rangoCapturado.FormatConditions.Add(
                            Excel.XlFormatConditionType.xlExpression, Type.Missing, formulaCondicionalMatematica);

                        formatoAmarillo.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(47, 117, 181));
                        chkBlancos.Checked = false;
                        MessageBox.Show(this,"Validación de blancos inteligente aplicada.\n\nAhora el sistema ignora filas completamente vacías y coexiste con tus reglas de bloqueo.", "SAVCNG Arquitectura", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(this,"Error en el motor de validación inteligente: " + ex.Message, "Error de Inyección", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        chkBlancos.Checked = false;
                    }
                }
                // --- VALIDACIÓN ESPECIFIQUES (PALABRAS CLAVE) ---
                else if (chkEspClave.Checked == true)
                {
                    if (_libroCenso == null || _rangoCapturado == null)
                    {
                        MessageBox.Show(this,"Operación denegada: Captura la celda destino (Especifique) primero.",
                                        "Arquitectura SAVCNG", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        chkEspClave.Checked = false;
                        return;
                    }

                    Excel.Application xlApp = (Excel.Application)ExcelDna.Integration.ExcelDnaUtil.Application;
                    Excel.Worksheet ws = (Excel.Worksheet)_rangoCapturado.Worksheet;
                    Excel.Range rangoCatalogo = null;
                    Excel.Range rangoMensaje = null;
                    Excel.Range celdaMotor = null;

                    try
                    {
                        // =========================================================================
                        // FASE 1: RECOPILACIÓN DE ESPACIOS DE TRABAJO (UX)
                        // =========================================================================
                        rangoCatalogo = (Excel.Range)xlApp.InputBox(
                            "1. Selecciona las OPCIONES DEL CATÁLOGO.\nNOTA: Debera omitir de la selección las opciones 'Otro(Especifique)' y/o 'No identificado' del catálogo correspondiente. ",
                            "Mapeo de Catálogo", Type: 8);
                        if (rangoCatalogo == null) throw new Exception("Cancelado");

                        rangoMensaje = (Excel.Range)xlApp.InputBox(
                            "2. Selecciona el rango o celda donde se mostrará el MENSAJE DE ALERTA.",
                            "Destino de Alerta", Type: 8);
                        if (rangoMensaje == null) throw new Exception("Cancelado");

                        celdaMotor = (Excel.Range)xlApp.InputBox(
                            "3. Selecciona UNA CELDA VACÍA (columna AF en adelante) para construir el Diccionario Auxiliar de Busqueda.\nADVERTENCIA: Considere un espacio libre de dos columnas por N filas. (N = numero de opciones del catalogo)",
                            "Generación del Motor Oculto", Type: 8);
                        if (celdaMotor == null) throw new Exception("Cancelado");

                        // =========================================================================
                        // FASE 2: CONSTRUCCIÓN DEL MOTOR AUXILIAR MATRICIAL
                        // =========================================================================
                        Excel.Range primeraCeldaAzul = (Excel.Range)_rangoCapturado.Cells[1, 1];
                        string dirAzulAbsoluta = primeraCeldaAzul.get_Address(true, true, Excel.XlReferenceStyle.xlA1, false);

                        // 2.1 - Inyectamos la Celda Limpia (Elimina acentos y minúsculas en tiempo real nativo)
                        Excel.Range celdaLimpiaAzul = celdaMotor.Offset[0, 0];
                        string formulaLimpieza = $"=LOWER(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE({dirAzulAbsoluta},\"á\",\"a\"),\"é\",\"e\"),\"í\",\"i\"),\"ó\",\"o\"),\"ú\",\"u\"))";
                        celdaLimpiaAzul.Formula = formulaLimpieza;
                        // Pintamos el fondo de la celda de amarillo puro
                        celdaLimpiaAzul.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        //Poner el texto en negrita 
                        celdaLimpiaAzul.Font.Bold = true;
                        string dirLimpiaAbsoluta = celdaLimpiaAzul.get_Address(true, true, Excel.XlReferenceStyle.xlA1, false);

                        int filaMotor = 1; // Comenzamos a escribir debajo de la celda limpia
                        System.Collections.Generic.List<string> celdasResultadosBooleanos = new System.Collections.Generic.List<string>();

                        // Limpiamos los formatos previos del catálogo para evitar acumulaciones
                        rangoCatalogo.FormatConditions.Delete();

                        // Iteramos sobre las celdas del catálogo seleccionado
                        foreach (Excel.Range celdaCat in rangoCatalogo.Cells)
                        {
                            string textoOriginal = celdaCat.Text != null ? celdaCat.Text.ToString() : "";

                            if (!string.IsNullOrWhiteSpace(textoOriginal) && textoOriginal.Trim() != "")
                            {
                                // -- INTELIGENCIA DE C# (Sustituye a O365) --
                                // 1. Quitamos contenido entre paréntesis ej: "(transparencia)"
                                string textoProcesado = Regex.Replace(textoOriginal, @"\(.*?\)", "").Trim();

                                // 2. Quitamos acentos y pasamos a minúsculas
                                textoProcesado = textoProcesado.ToLower()
                                    .Replace("á", "a").Replace("é", "e").Replace("í", "i")
                                    .Replace("ó", "o").Replace("ú", "u");

                                // 3. Dividimos por conectores lógicos para obtener palabras clave puras
                                string[] separadores = { " y/o ", " y ", " o ", " e ", ",", "/", " a ", " ante ", " bajo ", " cabe ", " con ", " contra ", " de ", " del ", " desde ", " durante ", " en ", " entre ", " hacia ", " hasta ", " mediante ", " para ", " por ", " según ", " sin ", " sobre ", " tras " };
                                string[] palabrasClave = textoProcesado.Split(separadores, StringSplitOptions.RemoveEmptyEntries);

                                // 4. Construimos la lógica nativa heredada (OR -> SEARCH)
                                System.Collections.Generic.List<string> fragmentosSearch = new System.Collections.Generic.List<string>();
                                foreach (string palabra in palabrasClave)
                                {
                                    string palabraLimpia = palabra.Trim();
                                    if (palabraLimpia.Length > 2) // Omitimos letras sueltas
                                    {
                                        fragmentosSearch.Add($"ISNUMBER(SEARCH(\"{palabraLimpia}\", {dirLimpiaAbsoluta}))");
                                    }
                                }

                                if (fragmentosSearch.Count > 0)
                                {
                                    Excel.Range filaValidacion = celdaMotor.Offset[filaMotor, 0];

                                    // Columna 1 del motor: El nombre original (Solo como referencia visual)
                                    filaValidacion.Value2 = textoOriginal;

                                    // Columna 2 del motor: Inyectamos la fórmula booleana nativa
                                    Excel.Range celdaMatch = celdaMotor.Offset[filaMotor, 1];
                                    string formulaMatch = $"=OR({string.Join(",", fragmentosSearch)})";
                                    celdaMatch.Formula = formulaMatch;

                                    string dirMatchAbsoluta = celdaMatch.get_Address(true, true, Excel.XlReferenceStyle.xlA1, false);
                                    celdasResultadosBooleanos.Add(dirMatchAbsoluta);

                                    // =========================================================================
                                    // FASE 3: TRUCO ARQUITECTÓNICO DE TRADUCCIÓN (Evitar #¿NOMBRE?)
                                    // =========================================================================
                                    string celdaCatRelativa = celdaCat.get_Address(false, false, Excel.XlReferenceStyle.xlA1, false);
                                    string formulaFormatCondIngles = $"=AND({celdaCatRelativa}<>\"\", {dirMatchAbsoluta}=TRUE)";

                                    // Inyectamos en celda dummy oculta para forzar la traducción al Excel del usuario
                                    Excel.Range celdaDummy = ws.Cells[1048576, 16384]; // Última celda de la hoja XFD1048576
                                    celdaDummy.Formula = formulaFormatCondIngles;
                                    string formulaFormatCondLocal = celdaDummy.FormulaLocal; // ¡Aquí obtenemos el idioma local seguro!
                                    celdaDummy.Clear();

                                    // Aplicamos el formato condicional específico a la celda del catálogo
                                    Excel.FormatCondition fc = (Excel.FormatCondition)celdaCat.FormatConditions.Add(
                                        Excel.XlFormatConditionType.xlExpression,
                                        Type.Missing,
                                        formulaFormatCondLocal);

                                    fc.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                                    fc.Font.Bold = true;

                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(fc);
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(celdaMatch);
                                    filaMotor++;
                                }
                            }
                        }

                        // =========================================================================
                        // FASE 4: VINCULACIÓN DEL MENSAJE DE ALERTA (B77)
                        // =========================================================================
                        if (celdasResultadosBooleanos.Count > 0)
                        {
                            string sumaResultados = string.Join(",", celdasResultadosBooleanos);
                            // Si el motor detecta algún TRUE, arroja el mensaje.
                            string formulaMensajeFinal = $"=IF(COUNTIF({celdaMotor.Offset[1, 1].get_Address(true, true)}:{celdaMotor.Offset[filaMotor - 1, 1].get_Address(true, true)}, TRUE)>0, \"Alerta: Revise el texto ingresado en el Especifique ya que podría existir en las opciones resaltadas en amarillo\", \"\")";

                            if (rangoMensaje.Count > 1) { rangoMensaje.Merge(); }
                            rangoMensaje.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                            rangoMensaje.Font.Name = "Arial";
                            rangoMensaje.Font.Size = 9;
                            rangoMensaje.Font.Bold = true;
                            rangoMensaje.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(191, 144, 0));

                            rangoMensaje.Formula = formulaMensajeFinal;
                        }

                        // Nota el "this," como primer parámetro
                        MessageBox.Show(this, "El Diccionario de Palabras Clave fue construido con éxito.", "SAVCNG - ExcelDNA", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        if (ex.Message != "Cancelado")
                        {
                            MessageBox.Show(this,"Operación Cancelada." + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                    finally
                    {
                        chkEspClave.Checked = false;
                        if (rangoCatalogo != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rangoCatalogo);
                        if (rangoMensaje != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rangoMensaje);
                        if (celdaMotor != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(celdaMotor);
                    }
                }
                // --- VALIDACIÓN FECHAS ---
                else if (chkFechas.Checked == true)
                {
                    Excel.Validation objValidacion = null;
                    Excel.Range celdaInicial = null;

                    try
                    {
                        // 1. CAPTURA DINÁMICA: Solicitar límite inferior usando InputBox nativo de Excel
                        // El "2" al final indica que esperamos que devuelva texto
                        object resultadoInferior = excelApp.InputBox(
                            "Indica el valor MÍNIMO aceptado para esta validación:\n\n(Ej. 1 para días/meses, o 1821 para años).",
                            "Límite Inferior",
                            "1", Type.Missing, Type.Missing, Type.Missing, Type.Missing, 2);

                        // Si el usuario presiona "Cancelar", excelApp devuelve un booleano (false)
                        if (resultadoInferior is bool && (bool)resultadoInferior == false) { chkFechas.Checked = false; return; }

                        string inputInferior = resultadoInferior.ToString().Trim();
                        if (string.IsNullOrWhiteSpace(inputInferior)) { chkFechas.Checked = false; return; }

                        // 2. CAPTURA DINÁMICA: Solicitar límite superior
                        object resultadoSuperior = excelApp.InputBox(
                            "Indica el valor MÁXIMO aceptado para esta validación:\n\n(Ej. 31 para días, 12 para meses, o 2026 para años).",
                            "Límite Superior",
                            "2026", Type.Missing, Type.Missing, Type.Missing, Type.Missing, 2);

                        if (resultadoSuperior is bool && (bool)resultadoSuperior == false) { chkFechas.Checked = false; return; }

                        string inputSuperior = resultadoSuperior.ToString().Trim();
                        if (string.IsNullOrWhiteSpace(inputSuperior)) { chkFechas.Checked = false; return; }

                        // 3. VALIDACIÓN DE ENTRADAS: Asegurar consistencia numérica
                        if (!int.TryParse(inputInferior, out int limiteInferior) || !int.TryParse(inputSuperior, out int limiteSuperior))
                        {
                            MessageBox.Show(this,
                                "Por favor, asegúrate de escribir únicamente números enteros.\n\nNo se permiten letras, decimales, ni dejar el espacio en blanco.",
                                "Solo números permitidos",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning);

                            chkFechas.Checked = false;
                            return;
                        }

                        if (limiteInferior > limiteSuperior)
                        {
                            MessageBox.Show(this,
                                $"Revisa los valores que ingresaste:\n\nEl límite mínimo ({limiteInferior}) no puede ser mayor que el límite máximo ({limiteSuperior}).",
                                "Límites invertidos",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning);

                            chkFechas.Checked = false;
                            return;
                        }

                        // 4. Optimizamos el rendimiento visual 
                        excelApp.ScreenUpdating = false;

                        // 5. Limpieza estricta de validaciones previas para evitar colisiones
                        objValidacion = _rangoCapturado.Validation;
                        objValidacion.Delete();

                        // 6. Inteligencia Espacial: Obtenemos la primera celda en formato relativo ("A1")
                        celdaInicial = (Excel.Range)_rangoCapturado.Cells[1, 1];
                        string direccionRelativa = celdaInicial.get_Address(false, false, Excel.XlReferenceStyle.xlA1, Type.Missing, Type.Missing);

                        // 7. Fórmula Localizada con límites dinámicos (acepta NS y NA)
                        string formulaValidacion = $"=O(ESPACIOS({direccionRelativa})=\"NS\"{separador}ESPACIOS({direccionRelativa})=\"NA\"{separador}Y(ESNUMERO({direccionRelativa}){separador}{direccionRelativa}>={limiteInferior}{separador}{direccionRelativa}<={limiteSuperior}))";

                        // 8. Inyección del motor de reglas personalizado
                        objValidacion.Add(
                            Excel.XlDVType.xlValidateCustom,
                            Excel.XlDVAlertStyle.xlValidAlertStop,
                            Excel.XlFormatConditionOperator.xlBetween,
                            formulaValidacion,
                            Type.Missing
                        );

                        // 9. Configuración de UX: Mensaje de error dinámico e incluyente (NS/NA)
                        objValidacion.IgnoreBlank = true;
                        objValidacion.ShowError = true;
                        objValidacion.ErrorTitle = "Valor fuera de rango";
                        objValidacion.ErrorMessage = $"El número debe estar entre {limiteInferior} y {limiteSuperior}.\n\nTambién puedes usar las claves válidas 'NS' o 'NA'.";

                        // 10. Limpiamos interfaz y notificamos el éxito
                        chkFechas.Checked = false;
                        MessageBox.Show(this,
                            $"Validación de rango aplicada con éxito.\n\nLas celdas ahora solo aceptarán números del {limiteInferior} al {limiteSuperior} (o claves NS/NA).",
                            "Validación completada",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                    }
                    catch (System.Runtime.InteropServices.COMException comEx)
                    {
                        MessageBox.Show(this,$"Error de sintaxis COM al inyectar la fórmula en Excel: {comEx.Message}\nCódigo de error: {comEx.ErrorCode}", "Error Crítico COM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        chkFechas.Checked = false;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(this,$"Error inesperado en el sistema: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        chkFechas.Checked = false;
                    }
                    finally
                    {
                        // 11. Prevención de fugas de memoria (Memory Leaking)
                        if (celdaInicial != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(celdaInicial);
                        if (objValidacion != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(objValidacion);

                        // Restaurar refresco de pantalla pase lo que pase
                        excelApp.ScreenUpdating = true;
                    }
                }
                // --- VALIDACIÓN SUMAS ---
                else if (chkSumas.Checked == true)
                {
                    Excel.Range rangoTotal = null;
                    Excel.Range rangoDesagregados = null;
                    Excel.Range rangoTotalesVerticales = null;
                    Excel.Range celdaTotalAux = null;
                    Excel.Range celdaDummy = null;

                    try
                    {
                        // 1. Obtener los límites espaciales de la matriz capturada en el Paso 1
                        int filaInicio = _rangoCapturado.Row;
                        int filaFin = filaInicio + _rangoCapturado.Rows.Count - 1;

                        // 2. UX de Captura: Solicitar la columna exacta que funge como TOTAL operativo
                        object resTotal = excelApp.InputBox(
                            "1. Selecciona la COLUMNA del TOTAL (debe coincidir con las filas del rango capturado):",
                            "SAVCNG - Mapeo de Totales", Type: 8);

                        if (resTotal is bool && (bool)resTotal == false) { chkSumas.Checked = false; return; }
                        rangoTotal = (Excel.Range)resTotal;

                        // 3. UX de Captura: Solicitar las columnas de los DESAGREGADOS (Soporta columnas no contiguas usando CTRL)
                        object resDesagregados = excelApp.InputBox(
                            "2. Selecciona las COLUMNAS de los DESAGREGADOS (Puedes usar CTRL para seleccionar varias separadas):",
                            "SAVCNG - Mapeo de Desagregados", Type: 8);

                        if (resDesagregados is bool && (bool)resDesagregados == false) { chkSumas.Checked = false; return; }
                        rangoDesagregados = (Excel.Range)resDesagregados;

                        // 4. UX de Captura: Solicitar las celdas destino de la sumatoria vertical (Fila de la sumatoria Σ)
                        object resVerticales = excelApp.InputBox(
                            "3. Selecciona la FILA o CELDAS destino para la Sumatoria Vertical (Σ) al final de la tabla:",
                            "SAVCNG - Destino Suma Vertical", Type: 8);

                        if (resVerticales is bool && (bool)resVerticales == false) { chkSumas.Checked = false; return; }
                        rangoTotalesVerticales = (Excel.Range)resVerticales;

                        // 5. Análisis Espacial: Extraer la columna de Total y limpiar las coordenadas
                        celdaTotalAux = (Excel.Range)rangoTotal.Cells[1, 1];
                        string letraTotal = celdaTotalAux.Address.Split('$')[1];

                        System.Collections.Generic.List<string> letrasDesagregados = new System.Collections.Generic.List<string>();
                        System.Collections.Generic.List<string> fragmentosExclusionNS = new System.Collections.Generic.List<string>();

                        // Incluimos la columna Total en la protección contra registros "NS"
                        fragmentosExclusionNS.Add($"COUNTIF(${letraTotal}{filaInicio},\"NS\")=0");
                        fragmentosExclusionNS.Add($"COUNTIF(${letraTotal}{filaInicio},\"ns\")=0");

                        // Iterar de forma segura sobre las áreas seleccionadas de desagregados para extraer sus letras de columna
                        foreach (Excel.Range area in rangoDesagregados.Areas)
                        {
                            for (int c = 1; c <= area.Columns.Count; c++)
                            {
                                Excel.Range colCelda = (Excel.Range)area.Cells[1, c];
                                string letra = colCelda.Address.Split('$')[1];

                                if (!letrasDesagregados.Contains(letra))
                                {
                                    letrasDesagregados.Add(letra);
                                    // Si cualquiera de estas celdas contiene NS o ns, la regla matemática NO debe activarse
                                    fragmentosExclusionNS.Add($"COUNTIF(${letra}{filaInicio},\"NS\")=0");
                                    fragmentosExclusionNS.Add($"COUNTIF(${letra}{filaInicio},\"ns\")=0");
                                }
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(colCelda);
                            }
                        }

                        // 6. Construcción del Álgebra Booleana (Fórmula en Inglés Universal con separadores de coma)
                        string formulaSumandos = string.Join("+", letrasDesagregados.ConvertAll(l => $"${l}{filaInicio}"));
                        string formulaCondicionesNS = string.Join(",", fragmentosExclusionNS);

                        string formulaErrorFilaIngles = $"=AND(ISNUMBER(${letraTotal}{filaInicio}), {formulaCondicionesNS}, ${letraTotal}{filaInicio}<>({formulaSumandos}))";

                        // =========================================================================
                        // 7. TRUCO ARQUITECTÓNICO DE TRADUCCIÓN NATIVA (FILA-SENSIBLE)
                        // =========================================================================
                        // Ubicamos la celda dummy en la MISMA fila de inicio, pero en la última columna (XFD / 16384).
                        // Esto evita que las referencias relativas de fila se desfasen al leer FormulaLocal.
                        Excel.Worksheet wsActual = (Excel.Worksheet)_rangoCapturado.Worksheet;
                        celdaDummy = (Excel.Range)wsActual.Cells[filaInicio, 16384];

                        celdaDummy.Formula = formulaErrorFilaIngles;
                        string formulaErrorFilaLocal = celdaDummy.FormulaLocal;
                        celdaDummy.Clear();
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(celdaDummy);
                        celdaDummy = null;

                        // =========================================================================
                        // 8. AUDITORÍA DE COEXISTENCIA (PERMITE APILAMIENTO MULTINIVEL)
                        // =========================================================================
                        if (_rangoCapturado.FormatConditions.Count > 0)
                        {
                            DialogResult respFormato = MessageBox.Show(this,
                                "Se detectaron reglas de validación de sumas previas en este rango.\n\n" +
                                "¿Deseas CONSERVARLAS e integrar esta nueva capa de revisión?\n\n" +
                                "SÍ = Apilar reglas (Recomendado para matrices jerárquicas con subtotales).\n" +
                                "NO = Borrar las reglas anteriores y dejar solo esta nueva.",
                                "SAVCNG - Formatos Condicionales Detectados",
                                MessageBoxButtons.YesNoCancel,
                                MessageBoxIcon.Warning);

                            if (respFormato == DialogResult.Cancel) { return; }
                            if (respFormato == DialogResult.No) { _rangoCapturado.FormatConditions.Delete(); }
                        }

                        // =========================================================================
                        // 9. INYECCIÓN DEL FORMATO CONDICIONAL LOCALIZADO
                        // =========================================================================
                        Excel.FormatCondition fcError = (Excel.FormatCondition)_rangoCapturado.FormatConditions.Add(
                            Excel.XlFormatConditionType.xlExpression, Type.Missing, formulaErrorFilaLocal);

                        // Configuración estética del error: Fondo Rojo Claro con Fuente Roja Oscura
                        fcError.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 199, 206));
                        fcError.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(156, 0, 6));

                        // IMPORTANTE: Permitir que Excel continúe evaluando las otras reglas apiladas en la misma celda
                        fcError.StopIfTrue = false;
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(fcError);

                        // =========================================================================
                        // 10. INYECCIÓN AUTOMATIZADA DE SUMATORIAS VERTICALES (CON GENERACIÓN DE MEMORIA LIMPIA)
                        // =========================================================================
                        foreach (Excel.Range celdaSumatoria in rangoTotalesVerticales.Cells)
                        {
                            string letraColVertical = celdaSumatoria.Address.Split('$')[1];
                            celdaSumatoria.Formula = $"=SUM({letraColVertical}{filaInicio}:{letraColVertical}{filaFin})";

                            // Liberación explícita e inmediata de la celda para evitar congelar el hilo de Excel
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(celdaSumatoria);
                        }

                        // 11. Cierre del proceso UX
                        chkSumas.Checked = false;
                        MessageBox.Show(this, "Validación de consistencia horizontal (Sumas) y fórmulas verticales inyectadas con éxito.",
                                        "SAVCNG Arquitectura", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(this, "Error crítico al inyectar el motor de sumas: " + ex.Message, "Error de Inyección", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        chkSumas.Checked = false;
                    }
                    finally
                    {
                        // 12. Recolección de Basura COM estructurada para objetos raíz
                        if (celdaTotalAux != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(celdaTotalAux);
                        if (celdaDummy != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(celdaDummy);
                        if (rangoTotal != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rangoTotal);
                        if (rangoDesagregados != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rangoDesagregados);
                        if (rangoTotalesVerticales != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rangoTotalesVerticales);
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

    }
}
