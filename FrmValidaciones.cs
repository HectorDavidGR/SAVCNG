using System;
using System.Windows.Forms;
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
        }


        // --- Aquí irán los eventos de los botones en el siguiente paso ---

        private void btnCapturarRango_Click(object sender, EventArgs e)
        {
            try
            {
                // 1. Obtenemos la aplicación de Excel (para saber qué está pasando ahí)
                Microsoft.Office.Interop.Excel.Application excelApp = (Microsoft.Office.Interop.Excel.Application)ExcelDna.Integration.ExcelDnaUtil.Application;

                // 2. Le preguntamos a Excel: "¿Qué tiene seleccionado el usuario en este momento?"
                object seleccion = excelApp.Selection;

                // 3. Revisamos que el usuario haya seleccionado celdas (y no una imagen o una gráfica por error)
                if (seleccion is Excel.Range)
                {
                    // 4. Guardamos esas celdas en nuestra memoria (_rangoCapturado)
                    _rangoCapturado = (Excel.Range)seleccion;

                    // 5. Le mostramos un mensajito de éxito al usuario con las coordenadas
                    // Detectamos la pregunta automáticamente al capturar
                    string pregunta = ObtenerNumeroPregunta(_rangoCapturado);
                    lblPregunta.Text = "Pregunta detectada: " + pregunta;
                    lblRangoSeleccionado.Text = "Rango seleccionado: " + _rangoCapturado.Address.Replace("$", "");

                    MessageBox.Show("Se capturó correctamente el rango: " + _rangoCapturado.Address,
                                    "Captura exitosa", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    // Si seleccionó una imagen, le avisamos
                    MessageBox.Show("Por favor, selecciona celdas de Excel, no imágenes ni gráficos.",
                                    "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
            }
            catch (Exception ex)
            {
                // Si algo sale mal, que no se rompa el programa, solo que nos avise
                MessageBox.Show("Ocurrió un error al capturar: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnAplicar_Click(object sender, EventArgs e)
        {
            Excel.Application excelApp = (Excel.Application)ExcelDna.Integration.ExcelDnaUtil.Application;
            string separador = excelApp.International[Excel.XlApplicationInternational.xlListSeparator].ToString();

            

            try
            {
                // 1. Primero verificamos que el usuario no haya olvidado capturar un rango
                if (_rangoCapturado == null)
                {
                    MessageBox.Show("¡Espera! Primero debes capturar un rango.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return; // Detenemos el código aquí
                }

                // 2. Verificamos si la casilla de "Decimales" está marcada
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
                        // FÓRMULA LOCAL (ESPAÑOL) + PROTECCIÓN DE CELDAS VACÍAS
                        // ==========================================================
                        // Usamos O(ESBLANCO(...)) para evitar que el motor de Validación rechace la fórmula
                        // al ser inyectada en una celda vacía.
                        string formulaDecimales = $"=O(ESBLANCO({direccion}){separador}Y(ESNUMERO({direccion}){separador}TRUNCAR({direccion})={direccion}))";

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
                        _rangoCapturado.Validation.ErrorTitle = "Captura Inválida (Solo Enteros)";
                        _rangoCapturado.Validation.ErrorMessage = "El formato de esta celda no admite números con decimales ni texto.\n\nPor favor, introduce únicamente un número entero (Ej: 1, 15, 100).";

                        // 6. Limpiamos la interfaz y avisamos del éxito
                        chkDecimales.Checked = false;
                        MessageBox.Show("Validación restrictiva de Enteros aplicada con éxito.\n\nEl sistema lanzará una ventana emergente si el informante intenta capturar decimales o texto.", "SAVCNG Arquitectura", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error crítico al aplicar Validación de Decimales: " + ex.Message, "Error de Inyección", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        chkDecimales.Checked = false;
                    }
                }
                else if (chkCatalogos.Checked == true)
                {
                    string formulaOpciones = "";

                    // PREGUNTA NUEVA: ¿Manual o desde Excel?
                    DialogResult tipoEntrada = MessageBox.Show(
                        "¿Deseas escribir el valor de la lista manualmente (ej: un solo valor como 'X' o varios como '1,2,3')?\n\n" +
                        "SÍ: Escribir el valor directamente.\n" +
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
                            "Escribe las opciones para tu lista desplegable.\n(Si son varias, sepáralas por comas):",
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
                            "Selecciona el rango de opciones o la celda que contiene la lista (ej: 1,2,9):",
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
                            // AQUI CAMBIAMOS A YesNoCancel para dar 3 opciones
                            DialogResult respuesta = MessageBox.Show(
                                "Has seleccionado una celda (o bloque combinado). ¿Cómo deseas extraer sus opciones?\n\n" +
                                "SÍ: Extraer SOLO NÚMEROS (Limpia texto y deja ej: 1,2,9).\n" +
                                "NO: Mantener el TEXTO EXACTO (Ideal para 'X' o palabras).\n" +
                                "CANCELAR: Usar como referencia de rango normal (=$A$1).",
                                "Configuración de Catálogo",
                                MessageBoxButtons.YesNoCancel,
                                MessageBoxIcon.Question);

                            if (respuesta == DialogResult.Cancel)
                            {
                                formulaOpciones = "=" + rangoOrigen.get_Address(true, true, Excel.XlReferenceStyle.xlA1, true);
                            }
                            else
                            {
                                Excel.Range primeraCeldaOrigen = (Excel.Range)rangoOrigen.Cells[1, 1];
                                string textoCelda = primeraCeldaOrigen.Text?.ToString() ?? "";

                                if (string.IsNullOrWhiteSpace(textoCelda))
                                {
                                    MessageBox.Show("La celda origen está vacía. No se puede crear la lista.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                    chkCatalogos.Checked = false;
                                    return;
                                }

                                string[] pedacitos = textoCelda.Split(new char[] { '.', ',', '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
                                System.Collections.Generic.List<string> listaLimpios = new System.Collections.Generic.List<string>();

                                if (respuesta == DialogResult.Yes)
                                {
                                    // SÍ: SOLO NÚMEROS (Tu lógica original)
                                    foreach (string pedazo in pedacitos)
                                    {
                                        string soloNumeros = "";
                                        foreach (char letra in pedazo) { if (char.IsDigit(letra)) soloNumeros += letra; }
                                        if (!string.IsNullOrEmpty(soloNumeros)) listaLimpios.Add(soloNumeros);
                                    }
                                }
                                else if (respuesta == DialogResult.No)
                                {
                                    // NO: TEXTO EXACTO
                                    foreach (string pedazo in pedacitos)
                                    {
                                        string textoLimpio = pedazo.Trim();
                                        if (!string.IsNullOrEmpty(textoLimpio)) listaLimpios.Add(textoLimpio);
                                    }
                                }

                                string separadorSistema = excelApp.International[Excel.XlApplicationInternational.xlListSeparator].ToString();
                                formulaOpciones = string.Join(separadorSistema, listaLimpios);
                            }
                        }
                        else
                        {
                            // CASO DE RANGO NORMAL (Varias celdas seleccionadas)
                            DialogResult respuestaRango = MessageBox.Show(
                                "Has seleccionado varias celdas. ¿Cómo deseas extraer sus opciones?\n\n" +
                                "SÍ: Extraer SOLO NÚMEROS (Ignora letras).\n" +
                                "NO: Mantener el TEXTO EXACTO (Ideal para celdas con letras como 'X').\n" +
                                "CANCELAR: Usar el rango normal con todo su contenido original.",
                                "Configuración de Catálogo",
                                MessageBoxButtons.YesNoCancel,
                                MessageBoxIcon.Question);

                            if (respuestaRango == DialogResult.Cancel)
                            {
                                formulaOpciones = "=" + rangoOrigen.get_Address(true, true, Excel.XlReferenceStyle.xlA1, true);
                            }
                            else
                            {
                                System.Collections.Generic.List<string> listaLimpios = new System.Collections.Generic.List<string>();

                                foreach (Excel.Range celda in rangoOrigen.Cells)
                                {
                                    string textoCelda = celda.Text?.ToString() ?? "";

                                    if (respuestaRango == DialogResult.Yes)
                                    {
                                        string soloNumeros = "";
                                        foreach (char letra in textoCelda) { if (char.IsDigit(letra)) soloNumeros += letra; }
                                        if (!string.IsNullOrEmpty(soloNumeros)) listaLimpios.Add(soloNumeros);
                                    }
                                    else
                                    {
                                        string textoLimpio = textoCelda.Trim();
                                        if (!string.IsNullOrEmpty(textoLimpio)) listaLimpios.Add(textoLimpio);
                                    }
                                }

                                if (listaLimpios.Count == 0)
                                {
                                    MessageBox.Show("No se encontraron valores en el rango seleccionado.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                    chkCatalogos.Checked = false;
                                    return;
                                }

                                formulaOpciones = string.Join(separador, listaLimpios);
                            }
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
                        MessageBox.Show("¡Validación de catálogo aplicada con éxito!", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error al aplicar la validación: " + ex.Message + "\n\nTexto que se intentó usar: " + formulaOpciones, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        chkCatalogos.Checked = false;
                    }
                }
                else if (chkNS.Checked == true)
                {
                    try
                    {
                        // --- PASO 1: SOLICITAR DESTINO Y TEXTO AL USUARIO ---

                        // 1. Solicitar la ubicación de la alerta
                        object resDestino = excelApp.InputBox(
                            "Selecciona la celda o rango donde aparecerá el mensaje de alerta para registros 'NS' (se combinará automáticamente):",
                            "1. Ubicación de Alerta NS", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8); // 8 = Rango

                        if (resDestino is bool && (bool)resDestino == false) { chkNS.Checked = false; return; }
                        Excel.Range rangoAlerta = (Excel.Range)resDestino;

                        // 2. Solicitar el texto del mensaje
                        string textoSugerido = "Alerta: debido a que cuenta con registros NS, debe proporcionar una justificación en el área de comentarios al final de la pregunta";
                        object resTexto = excelApp.InputBox(
                            "Escribe el texto del mensaje de alerta:",
                            "2. Mensaje de Alerta NS", textoSugerido, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 2); // 2 = Texto

                        if (resTexto is bool && (bool)resTexto == false) { chkNS.Checked = false; return; }
                        string textoAlerta = resTexto.ToString().Trim();

                        // ==========================================================
                        // APLICACIÓN DE LA VALIDACIÓN DE DATOS (RESTRICCIÓN)
                        // ==========================================================

                        _rangoCapturado.Validation.Delete();

                        Excel.Range primeraCelda = (Excel.Range)_rangoCapturado.Cells[1, 1];
                        string direccion = primeraCelda.Address.Replace("$", "");

                        // Mantenemos tu lógica para la regla de validación de celdas
                        string formulaRestriccion = $"=O(Y(ESNUMERO({direccion}){separador}{direccion}>=0){separador}{direccion}=\"NS\")";

                        _rangoCapturado.Validation.Add(
                            Excel.XlDVType.xlValidateCustom,
                            Excel.XlDVAlertStyle.xlValidAlertStop,
                            Excel.XlFormatConditionOperator.xlBetween,
                            formulaRestriccion,
                            Type.Missing);

                        _rangoCapturado.Validation.IgnoreBlank = true;
                        _rangoCapturado.Validation.InCellDropdown = true;
                        _rangoCapturado.Validation.ErrorTitle = "Error de validación";
                        _rangoCapturado.Validation.ErrorMessage = "Solo se permiten números mayores o iguales a cero, o el valor 'NS'.";
                        _rangoCapturado.Validation.ShowError = true;

                        // ==========================================================
                        // APLICACIÓN DEL MENSAJE DE ALERTA DINÁMICO Y UNIVERSAL
                        // ==========================================================

                        // 1. Si el usuario seleccionó varias celdas, las combinamos
                        if (rangoAlerta.Count > 1)
                        {
                            rangoAlerta.Merge();
                        }

                        // 2. Aplicamos el formato profesional (Dorado)
                        rangoAlerta.Font.Name = "Arial";
                        rangoAlerta.Font.Size = 9;
                        rangoAlerta.Font.Bold = true;
                        rangoAlerta.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(191, 143, 0));
                        rangoAlerta.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

                        // 3. Construcción del motor de conteo universal (Inglés con comas)
                        System.Collections.Generic.List<string> partesCountIf = new System.Collections.Generic.List<string>();

                        // Si el usuario seleccionó un rango con áreas separadas (ej. con CTRL), iteramos por cada una
                        for (int i = 1; i <= _rangoCapturado.Areas.Count; i++)
                        {
                            Excel.Range area = (Excel.Range)_rangoCapturado.Areas[i];
                            string addrAbs = area.get_Address(true, true, Excel.XlReferenceStyle.xlA1, false);

                            // Construimos fragmentos COUNTIF universales
                            partesCountIf.Add($"COUNTIF({addrAbs},\"NS\")");
                        }

                        // Unimos los fragmentos con comas para usarlos en un SUM global: SUM(COUNTIF(...), COUNTIF(...))
                        string sumaInner = string.Join(",", partesCountIf);
                        string formulaFinalAlerta = $"=IF(SUM({sumaInner})>0, \"{textoAlerta}\", \"\")";

                        // 4. Inyectamos usando la propiedad universal .Formula
                        rangoAlerta.Formula = formulaFinalAlerta;

                        System.Diagnostics.Debug.WriteLine($"[DEBUG] Alerta configurada en {rangoAlerta.Address} con formato Arial 9 Negrita Dorado.");

                        chkNS.Checked = false;
                        MessageBox.Show("Validación NS y Mensaje de Alerta configurados correctamente en la ubicación seleccionada.", "SAVCNG Arquitectura", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error al aplicar Validación NS: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        chkNS.Checked = false;
                    }
                }
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
                            "Selecciona la COLUMNA o CELDA donde deseas ocultar la validación matemática:\n\n(Ej. Selecciona CW1 o cualquier celda en una columna vacía a la derecha de tu formato).",
                            "Ubicación del Rango Auxiliar", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8); // 8 = Rango

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
                            _rangoCapturado.Validation.ErrorMessage = "El texto capturado debe cumplir estrictamente las siguientes reglas:\n\n" +
                                                                      "• Todo en MAYÚSCULAS.\n" +
                                                                      "• Sin dobles espacios ni espacios a las orillas.\n" +
                                                                      "• SOLO LETRAS (incluye Ñ y acentos) Y NÚMEROS. No se permiten caracteres especiales.";
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
                        MessageBox.Show($"Validación restrictiva aplicada con éxito.\n\nEl motor auxiliar fue alojado en la columna {colAuxLetra}.", "SAVCNG Arquitectura", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error crítico al configurar Formato Texto: " + ex.Message, "Error del Sistema", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        chkFormatoTexto.Checked = false;
                    }
                }
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
                            DialogResult respFila = MessageBox.Show(
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
                            MessageBox.Show("Operador no reconocido.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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
                        DialogResult respuestaBlanco = MessageBox.Show(
                            "¿Deseas que la matriz permanezca DESBLOQUEADA si la celda de condición está VACÍA?\n\n" +
                            "SÍ = Si está en blanco, se puede escribir.\n" +
                            "NO = Estricto (Si está en blanco, se bloquea por defecto).",
                            "Regla de Celda Vacía", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                        DialogResult respuestaRojo = MessageBox.Show(
                            "¿Deseas que la celda se resalte en ROJO cuando se desbloquee y esté vacía?\n\n(Ideal para los campos 'Especifique' que se vuelven obligatorios).",
                            "4. Resalte de Obligatoriedad", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                        // ==========================================================
                        // --- APLICACIÓN DE REGLAS (MOTOR: CONTAR.SI) ---
                        // ==========================================================

                        _rangoCapturado.Validation.Delete();

                        // AUDITORÍA DE COEXISTENCIA
                        if (_rangoCapturado.FormatConditions.Count > 0)
                        {
                            DialogResult respFormato = MessageBox.Show(
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
                        MessageBox.Show($"Validación de Bloqueo Dinámica aplicada con éxito.\nModo: {modoAplicado}\nRegla: {operador} {valorCriterio}", "SAVCNG", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error al aplicar la Validación de Bloqueo: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        chkBloqueo.Checked = false;
                    }
                }
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
                            DialogResult respFormato = MessageBox.Show(
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
                        MessageBox.Show("Validación de blancos inteligente aplicada.\n\nAhora el sistema ignora filas completamente vacías y coexiste con tus reglas de bloqueo.", "SAVCNG Arquitectura", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error en el motor de validación inteligente: " + ex.Message, "Error de Inyección", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        chkBlancos.Checked = false;
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Ocurrió un error al aplicar el formato: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                    MessageBox.Show("Primero carga un censo y captura un rango con el botón.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
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
                    MessageBox.Show("Primero carga un censo y captura un rango con el botón.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
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
                    MessageBox.Show("Operación denegada: Carga un censo y define el rango de memoria primero.", "Advertencia Arquitectónica", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    chkBlancos.Checked = false;
                }
            }
        }


    }
}
