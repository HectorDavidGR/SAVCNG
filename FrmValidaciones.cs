using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel; // Importante para entenderse con Excel
using System.Drawing;

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
                    // Limpiamos formatos condicionales anteriores para no encimar reglas
                    _rangoCapturado.FormatConditions.Delete();

                    // 3. Obtenemos la dirección de la primera celda del rango (Ejemplo: "A1")
                    // false, false significa que nos dará "A1" en lugar de "$A$1"
                    Excel.Range primeraCelda = _rangoCapturado.Cells[1, 1];
                    string direccion = primeraCelda.get_Address(false, false);

                    // 4. Creamos la fórmula matemática de Excel en INGLÉS (Interop siempre usa inglés internamente)
                    // La fórmula dice: "Si es un número Y además el número truncado es diferente al original, entonces tiene decimales"
                    // En español sería =Y(ESNUMERO(A1), TRUNCAR(A1)<>A1)
                    string formula = $"=Y(ESNUMERO({direccion}), TRUNCAR({direccion})<>{direccion})";

                    // 5. Aplicamos la regla de formato condicional a todo el rango capturado
                    Excel.FormatCondition formato = (Excel.FormatCondition)_rangoCapturado.FormatConditions.Add(
                        Excel.XlFormatConditionType.xlExpression,
                        Type.Missing,
                        formula);

                    // 6. Si detecta el error (un decimal), le decimos que pinte la celda (Fondo Amarillo, Letra Roja)
                    formato.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    formato.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                    formato.Font.Bold = true;

                    // 7. Quitamos la marca de la casilla para dejarla lista para la siguiente vez
                    chkDecimales.Checked = false;

                    // 8. Avisamos que todo salió bien
                    MessageBox.Show("¡Validación de Decimales (Enteros) aplicada con éxito!", "Listo", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
                        // 1. Limpiamos validaciones previas
                        _rangoCapturado.Validation.Delete();

                        // 2. Extraemos la primera celda y obligamos a C# a verla como un Rango de Excel
                        Excel.Range primeraCelda = (Excel.Range)_rangoCapturado.Cells[1, 1];

                        // TRUCO A PRUEBA DE BALAS: .Address devuelve "$A$1". Con Replace le quitamos los "$" para que quede "A1".
                        string direccion = primeraCelda.Address.Replace("$", "");

                        // Fórmula de restricción
                        string formulaRestriccion = $"=O(Y(ESNUMERO({direccion}){separador}{direccion}>=0){separador}{direccion}=\"NS\")";

                        // 3. Aplicamos la validación
                        _rangoCapturado.Validation.Add(
                            Excel.XlDVType.xlValidateCustom,
                            Excel.XlDVAlertStyle.xlValidAlertStop,
                            Excel.XlFormatConditionOperator.xlBetween,
                            formulaRestriccion,
                            Type.Missing);

                        // 4. Mensajes de Error
                        _rangoCapturado.Validation.IgnoreBlank = true;
                        _rangoCapturado.Validation.InCellDropdown = true;
                        _rangoCapturado.Validation.ErrorTitle = "Error de validación";
                        _rangoCapturado.Validation.ErrorMessage = "Solo se permiten números mayores o iguales a cero, o el valor 'NS'.";
                        _rangoCapturado.Validation.ShowError = true;

                        // --- PARTE 2: EL MENSAJE DE ALERTA (Con Formato Profesional) ---

                        // 1. Calculamos la fila siguiente
                        int filaSiguiente = _rangoCapturado.Row + _rangoCapturado.Rows.Count;
                        Excel.Worksheet ws = (Excel.Worksheet)_rangoCapturado.Worksheet;

                        // 2. Definimos el RANGO DE ALERTA (Desde B hasta AD en la fila siguiente)
                        // Usamos ws.get_Range para marcar el bloque que vamos a combinar
                        Excel.Range rangoAlerta = ws.Range[ws.Cells[filaSiguiente, "B"], ws.Cells[filaSiguiente, "AD"]];

                        // 3. APLICAMOS EL MERGE (Combinar celdas)
                        rangoAlerta.Merge();

                        // 4. APLICAMOS EL FORMATO SOLICITADO
                        rangoAlerta.Font.Name = "Arial";
                        rangoAlerta.Font.Size = 9;
                        rangoAlerta.Font.Bold = true;

                        // Color #BF8F00 (R: 191, G: 143, B: 0)
                        rangoAlerta.Font.Color = ColorTranslator.ToOle(Color.FromArgb(191, 143, 0));

                        // 5. Opcional: Alineación a la izquierda para que el texto se vea ordenado
                        rangoAlerta.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

                        // --- LÓGICA DE FÓRMULAS (Se mantiene igual) ---
                        System.Collections.Generic.List<string> partesCountIf = new System.Collections.Generic.List<string>();

                        for (int i = 1; i <= _rangoCapturado.Areas.Count; i++)
                        {
                            Excel.Range area = (Excel.Range)_rangoCapturado.Areas[i];
                            string addrAbs = area.Address;
                            partesCountIf.Add($"CONTAR.SI({addrAbs}{separador}\"NS\")");
                        }

                        string sumaInner = string.Join(separador, partesCountIf);
                        string textoAlerta = "Alerta: debido a que cuenta con registros NS, debe proporcionar una justificación en el área de comentarios al final de la pregunta";
                        string formulaFinalAlerta = $"=SI(SUMA({sumaInner})<>0{separador}\"{textoAlerta}\"{separador}\"\")";

                        // 6. Pegamos la fórmula en el rango combinado
                        rangoAlerta.FormulaLocal = formulaFinalAlerta;

                        // LOG PARA CONSOLA
                        System.Diagnostics.Debug.WriteLine($"[DEBUG] Alerta configurada en {rangoAlerta.Address} con formato Arial 9 Negrita Dorado.");



                        chkNS.Checked = false;
                        MessageBox.Show("Validación NS y Mensaje de Alerta configurados correctamente.", "SAVCNG", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error al aplicar Validación NS: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else if (chkFormatoTexto.Checked == true)
                {
                    try
                    {
                        // 1. Limpiamos validaciones previas para no empalmar
                        _rangoCapturado.Validation.Delete();
                        
                        // 2. Extraemos la dirección de la celda (Ej: A1) para armar la fórmula
                        Excel.Range primeraCelda = (Excel.Range)_rangoCapturado.Cells[1, 1];
                        string direccion = primeraCelda.Address.Replace("$", "");

                        // 3. Construimos la fórmula de Excel en ESPAÑOL
                        // Lógica: =Y(IGUAL(A1, MAYUSC(A1)), LARGO(A1)=LARGO(ESPACIOS(A1)))
                        // - IGUAL(A1, MAYUSC(A1)) -> Obliga a que sea exactamente MAYÚSCULAS
                        // - LARGO(A1)=LARGO(ESPACIOS(A1)) -> Obliga a no tener espacios dobles, ni al inicio/final
                        string formulaTexto = $"=Y(IGUAL({direccion}{separador}MAYUSC({direccion})){separador}LARGO({direccion})=LARGO(ESPACIOS({direccion})))";

                        // 4. Aplicamos la regla de Data Validation (¡Esto es lo que saca el mensaje de alerta!)
                        _rangoCapturado.Validation.Add(
                            Excel.XlDVType.xlValidateCustom,
                            Excel.XlDVAlertStyle.xlValidAlertStop,
                            Excel.XlFormatConditionOperator.xlBetween,
                            formulaTexto,
                            Type.Missing);

                        // 5. Configuramos el mensaje de error que verá el usuario
                        _rangoCapturado.Validation.IgnoreBlank = true;
                        _rangoCapturado.Validation.ShowError = true;

                        _rangoCapturado.Validation.ErrorTitle = "Formato de texto inválido";
                        _rangoCapturado.Validation.ErrorMessage = "El texto debe cumplir las siguientes reglas:\n\n" +
                                                                  "• Todo en MAYÚSCULAS.\n" +
                                                                  "• Sin dobles espacios.\n" +
                                                                  "• Sin espacios al inicio o al final.";

                        // 6. Limpiamos interfaz y avisamos éxito
                        chkFormatoTexto.Checked = false;
                        MessageBox.Show("Validación restrictiva de Formato Texto configurada correctamente. El usuario no podrá ingresar minúsculas ni espacios extra.", "SAVCNG", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error al aplicar Validación de Formato Texto: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                            esFilaPorFila = true; // Solo una celda seleccionada, asumimos que baja fila por fila
                        }
                        else if (_rangoCapturado.Rows.Count == rangoCondicion.Rows.Count)
                        {
                            // ¡Magia! Los rangos miden exactamente lo mismo de alto.
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

                        // DIRECCIÓN LOCAL (Para bloquear/desbloquear las celdas capturadas)
                        string dirCondicionLocal = esFilaPorFila
                            ? rangoCondicion.Cells[1, 1].Address.Replace("$", "")  // Ej. A1 (Sin anclas)
                            : rangoCondicion.Address;             // Ej. $A$1:$A$21 (Fija todo)

                        // DIRECCIÓN GLOBAL (Para la instrucción amarilla, siempre debe monitorear todo el bloque)
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

                        if (resultadoOperador is bool && (bool)resultadoOperador == false)
                        {
                            chkBloqueo.Checked = false;
                            return;
                        }

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

                        if (resultadoValor is bool && (bool)resultadoValor == false)
                        {
                            chkBloqueo.Checked = false;
                            return;
                        }

                        string valorCriterio = resultadoValor.ToString().Trim();
                        bool esNumero = double.TryParse(valorCriterio, out _);
                        string valorFormateado = esNumero ? valorCriterio : $"\"{valorCriterio}\"";

                        // --- PASO 4: LA ALERTA ROJA ---
                        DialogResult respuestaRojo = MessageBox.Show(
                            "¿Deseas que la celda se resalte en ROJO cuando se desbloquee y esté vacía?\n\n(Ideal para los campos 'Especifique' que se vuelven obligatorios).",
                            "4. Resalte de Obligatoriedad",
                            MessageBoxButtons.YesNo,
                            MessageBoxIcon.Question);


                        // ==========================================================
                        // --- APLICACIÓN DE REGLAS (MOTOR: CONTAR.SI) ---
                        // ==========================================================
                        _rangoCapturado.Validation.Delete();
                        _rangoCapturado.FormatConditions.Delete();

                        string criterioContarSi = $"\"{operador}\"&{valorFormateado}";

                        // 1. DATA VALIDATION (Usamos dirCondicionLocal para evaluar fila por fila o global)
                        string formulaValidacion = $"=CONTAR.SI({dirCondicionLocal}{separador}{criterioContarSi})>0";

                        _rangoCapturado.Validation.Add(
                            Excel.XlDVType.xlValidateCustom,
                            Excel.XlDVAlertStyle.xlValidAlertStop,
                            Excel.XlFormatConditionOperator.xlBetween,
                            formulaValidacion,
                            Type.Missing);

                        _rangoCapturado.Validation.IgnoreBlank = true;
                        _rangoCapturado.Validation.ShowError = true;
                        _rangoCapturado.Validation.ErrorTitle = "Celda Bloqueada";
                        _rangoCapturado.Validation.ErrorMessage = $"No se permite capturar información. El flujo requiere encontrar una celda que sea {operador} {valorCriterio} en la referencia.";

                        // 2. FORMATO CONDICIONAL 1 (Gris Bloqueado)
                        string formulaSombreado = $"=CONTAR.SI({dirCondicionLocal}{separador}{criterioContarSi})=0";

                        Excel.FormatCondition formatoGris = (Excel.FormatCondition)_rangoCapturado.FormatConditions.Add(
                            Excel.XlFormatConditionType.xlExpression,
                            Type.Missing,
                            formulaSombreado);

                        formatoGris.Interior.Pattern = Excel.XlPattern.xlPatternCrissCross;
                        formatoGris.Interior.PatternColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Gray);

                        // 3. FORMATO CONDICIONAL 2 (Alerta Roja)
                        if (respuestaRojo == DialogResult.Yes)
                        {
                            string dirCapturada = _rangoCapturado.Cells[1, 1].Address.Replace("$", "");
                            string formulaRojo = $"=Y(CONTAR.SI({dirCondicionLocal}{separador}{criterioContarSi})>0{separador}ESBLANCO({dirCapturada}))";

                            Excel.FormatCondition formatoRojo = (Excel.FormatCondition)_rangoCapturado.FormatConditions.Add(
                                Excel.XlFormatConditionType.xlExpression,
                                Type.Missing,
                                formulaRojo);

                            formatoRojo.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 0, 0));
                            formatoRojo.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(156, 0, 6));
                        }

                        // ==========================================================
                        // --- PASO 5: RESALTE DE INSTRUCCIÓN ---
                        // ==========================================================
                        DialogResult respuestaInstruccion = MessageBox.Show(
                            "¿Deseas resaltar en AMARILLO alguna instrucción asociada a este bloqueo?\n\n(Esto guía visualmente al capturista para entender la alerta).",
                            "5. Resalte de Instrucción",
                            MessageBoxButtons.YesNo,
                            MessageBoxIcon.Question);

                        if (respuestaInstruccion == DialogResult.Yes)
                        {
                            object resultadoInstruccion = excelApp.InputBox(
                                "Selecciona la celda o rango que contiene la instrucción:",
                                "Seleccionar Instrucción",
                                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8);

                            if (!(resultadoInstruccion is bool && (bool)resultadoInstruccion == false))
                            {
                                Excel.Range rangoInstruccion = (Excel.Range)resultadoInstruccion;
                                rangoInstruccion.FormatConditions.Delete();

                                // Aquí usamos dirCondicionGlobal para que escanee todo el bloque buscando si ALGUNA fila se activó
                                string formulaInstruccion = $"=CONTAR.SI({dirCondicionGlobal}{separador}{criterioContarSi})>0";

                                Excel.FormatCondition formatoInstruccion = (Excel.FormatCondition)rangoInstruccion.FormatConditions.Add(
                                    Excel.XlFormatConditionType.xlExpression,
                                    Type.Missing,
                                    formulaInstruccion);

                                formatoInstruccion.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                                formatoInstruccion.Font.Bold = true;
                            }
                        }
                        // ==========================================================
                        // --- PASO 5: MENSAJE DE ALERTA DINÁMICO EN CELDA ---
                        // ==========================================================
                        DialogResult respuestaAlerta = MessageBox.Show(
                            "¿Deseas agregar un mensaje de alerta en alguna celda específica?\n\n(Aparecerá en rojo cuando se cumpla la condición y la captura siga vacía. Si seleccionas varias celdas, se combinarán automáticamente).",
                            "5. Mensaje de Alerta Especial",
                            MessageBoxButtons.YesNo,
                            MessageBoxIcon.Question);

                        if (respuestaAlerta == DialogResult.Yes)
                        {
                            // 1. Pedimos el texto del mensaje
                            object resultadoTexto = excelApp.InputBox(
                                "Escribe el texto del mensaje de alerta (Ej: 'Especifique el nombre de la otra clasificación'):",
                                "Texto del Mensaje",
                                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 2); // 2 = Texto

                            if (!(resultadoTexto is bool && (bool)resultadoTexto == false))
                            {
                                string textoAlerta = resultadoTexto.ToString().Trim();

                                // 2. Pedimos el lugar donde va a aparecer
                                object resultadoRangoAlerta = excelApp.InputBox(
                                    "Selecciona la celda o rango donde aparecerá este mensaje:",
                                    "Ubicación del Mensaje",
                                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8); // 8 = Rango

                                if (!(resultadoRangoAlerta is bool && (bool)resultadoRangoAlerta == false))
                                {
                                    Excel.Range rangoAlerta = (Excel.Range)resultadoRangoAlerta;

                                    // 3. ¡Magia! Si seleccionó más de una celda, las combinamos (Merge)
                                    if (rangoAlerta.Count > 1)
                                    {
                                        rangoAlerta.Merge();
                                    }

                                    // 4. Aplicamos el formato rojo para que grite "¡Alerta!"
                                    rangoAlerta.Font.Bold = true;
                                    rangoAlerta.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                                    rangoAlerta.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter; // Centrado horizontal
                                    rangoAlerta.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;   // Centrado vertical

                                    // 5. Construimos la fórmula lógica en INGLÉS UNIVERSAL
                                    string dirCapturadaGlobal = _rangoCapturado.Address;

                                    // Usamos IF, COUNTIF, COUNTA y la coma estándar de programación (,). 
                                    // Excel lo traducirá al idioma y símbolos locales mágicamente.
                                    string formulaAlerta = $"=IF(COUNTIF({dirCondicionGlobal},{criterioContarSi})>COUNTA({dirCapturadaGlobal}),\"{textoAlerta}\",\"\")";

                                    // 6. Inyectamos la fórmula usando .Formula (NUNCA .FormulaLocal)
                                    rangoAlerta.Formula = formulaAlerta;
                                }
                            }
                        }
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

                else
                {
                    MessageBox.Show("No has marcado ninguna validación para aplicar.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information);
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


    }
}
