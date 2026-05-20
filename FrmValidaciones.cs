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
                  

                    // 1. Mostrar el InputBox para seleccionar el origen
                    object resultadoInput = excelApp.InputBox(
                        "Selecciona el rango de opciones o la celda que contiene la lista (ej: 1,2,9):",
                        "Seleccionar Origen del Catálogo",
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8);

                    if (resultadoInput is bool && (bool)resultadoInput == false) return;

                    Excel.Range rangoOrigen = (Excel.Range)resultadoInput;
                    string formulaOpciones = "";

                    

                    // 2. LÓGICA DE ELECCIÓN: Detectamos celda única o combinada
                    bool esCeldaUnicaOCombinada = (rangoOrigen.Count == 1) || (bool)rangoOrigen.MergeCells;

                    if (esCeldaUnicaOCombinada)
                    {
                        DialogResult respuesta = MessageBox.Show(
                            "Has seleccionado una celda (o bloque combinado). ¿Deseas usar su CONTENIDO como lista de opciones (ej: 1,2,9)?\n\n" +
                            "SÍ: Extrae el texto dentro de la celda.\n" +
                            "NO: Usa la celda como una referencia de rango normal.",
                            "Configuración de Catálogo",
                            MessageBoxButtons.YesNo,
                            MessageBoxIcon.Question);

                        if (respuesta == DialogResult.Yes)
                        {
                            Excel.Range primeraCeldaOrigen = (Excel.Range)rangoOrigen.Cells[1, 1];
                            string textoCelda = primeraCeldaOrigen.Text?.ToString() ?? "";

                            if (string.IsNullOrWhiteSpace(textoCelda))
                            {
                                MessageBox.Show("La celda origen está vacía. No se puede crear la lista.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                chkCatalogos.Checked = false;
                                return;
                            }

                            // 1. PRIMERA LIMPIEZA: Separamos por punto, coma o saltos de línea
                            string[] pedacitos = textoCelda.Split(new char[] { '.', ',', '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);

                            System.Collections.Generic.List<string> listaNumerosLimpios = new System.Collections.Generic.List<string>();

                            // 2. SEGUNDA LIMPIEZA: Extraemos solo números
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

                            // 3. Obtenemos el separador oficial de tu Excel (coma o punto y coma)
                            string separadorSistema = excelApp.International[Excel.XlApplicationInternational.xlListSeparator].ToString();

                            // 4. UNIÓN FINAL: Usamos el separador normal, no el \0
                            formulaOpciones = string.Join(separadorSistema, listaNumerosLimpios);

                            // --- IMPRIMIR EN CONSOLA PARA REVISIÓN ---
                            System.Diagnostics.Debug.WriteLine("=========================================");
                            System.Diagnostics.Debug.WriteLine($"[DEBUG CATÁLOGOS] Texto original de la celda: '{textoCelda}'");
                            System.Diagnostics.Debug.WriteLine($"[DEBUG CATÁLOGOS] Opciones limpias a insertar : '{formulaOpciones}'");
                            System.Diagnostics.Debug.WriteLine("=========================================");
                        }
                        else
                        {
                            formulaOpciones = "=" + rangoOrigen.get_Address(true, true, Excel.XlReferenceStyle.xlA1, true);
                        }
                    }
                    else
                    {
                        // CASO DE RANGO NORMAL (Varias celdas seleccionadas)
                        DialogResult respuestaRango = MessageBox.Show(
                            "Has seleccionado varias celdas. ¿Deseas LIMPIARLAS y usar solo sus NÚMEROS como opciones (ej: 1, 2, 3)?\n\n" +
                            "SÍ: Extrae solo los números ignorando el texto.\n" +
                            "NO: Usa el rango normal con todo su contenido original.",
                            "Configuración de Catálogo",
                            MessageBoxButtons.YesNo,
                            MessageBoxIcon.Question);

                        if (respuestaRango == DialogResult.Yes)
                        {
                            System.Collections.Generic.List<string> listaNumerosLimpios = new System.Collections.Generic.List<string>();

                            // Recorremos cada celda dentro del rango que seleccionaste
                            foreach (Excel.Range celda in rangoOrigen.Cells)
                            {
                                string textoCelda = celda.Text?.ToString() ?? "";
                                string soloNumeros = "";

                                // Extraemos solo los dígitos de esta celda específica
                                foreach (char letra in textoCelda)
                                {
                                    if (char.IsDigit(letra))
                                    {
                                        soloNumeros += letra;
                                    }
                                }

                                // Si encontramos un número, lo guardamos a la lista general
                                if (!string.IsNullOrEmpty(soloNumeros))
                                {
                                    listaNumerosLimpios.Add(soloNumeros);
                                }
                            }

                            // Verificamos que sí hayamos encontrado números
                            if (listaNumerosLimpios.Count == 0)
                            {
                                MessageBox.Show("No se encontraron números en el rango seleccionado.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                chkCatalogos.Checked = false;
                                return;
                            }

                            // Unimos la lista usando el separador que ya definimos arriba en tu botón
                            formulaOpciones = string.Join(separador, listaNumerosLimpios);

                            // --- IMPRIMIR EN CONSOLA ---
                            System.Diagnostics.Debug.WriteLine("=========================================");
                            System.Diagnostics.Debug.WriteLine($"[DEBUG CATÁLOGOS RANGO] Opciones limpias: '{formulaOpciones}'");
                            System.Diagnostics.Debug.WriteLine("=========================================");
                        }
                        else
                        {
                            // Comportamiento normal de Excel si el usuario elige "NO"
                            formulaOpciones = "=" + rangoOrigen.get_Address(true, true, Excel.XlReferenceStyle.xlA1, true);
                        }
                    }

                    // 3. APLICAR LA VALIDACIÓN
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
                        // Si vuelve a fallar, este mensaje nos dirá EXACTAMENTE qué texto intentó poner
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
