using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using SAVCNG_ExcelDNA.Core; // Consumimos nuestra Fachada

namespace SAVCNG_ExcelDNA.Validaciones
{
    public class ValidacionBloqueos : IValidacionExcel
    {
        public void Ejecutar(Excel.Application excelApp, Excel.Workbook libroCenso, Excel.Range rangoCapturado)
        {
            Excel.Worksheet wsActual = null;
            Excel.Range rangoCondicion = null;

            try
            {
                wsActual = (Excel.Worksheet)rangoCapturado.Worksheet;

                // ==========================================================
                // FASE 1: LA CELDA O RANGO DE CONDICIÓN
                // ==========================================================
                object resultadoRango = excelApp.InputBox(
                    "Selecciona el RANGO o CELDA que controla el bloqueo:\n\n" +
                    "• Una celda: Se evaluará fila por fila.\n" +
                    "• Un rango: Se desbloqueará si el valor existe en CUALQUIER celda (o fila por fila si miden lo mismo).",
                    "1. Condición de Bloqueo",
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8); // 8 = Rango

                if (resultadoRango is bool && (bool)resultadoRango == false) return;

                rangoCondicion = (Excel.Range)resultadoRango;

                // ==========================================================
                // INTELIGENCIA SENIOR: ¿Búsqueda Global o Fila por Fila?
                // ==========================================================
                bool esFilaPorFila = false;

                if (rangoCondicion.Count == 1)
                {
                    esFilaPorFila = true;
                }
                else if (rangoCapturado.Rows.Count == rangoCondicion.Rows.Count)
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
                Excel.Range primeraCeldaCondicion = null;

                try
                {
                    primeraCeldaCondicion = (Excel.Range)rangoCondicion.Cells[1, 1];
                    if (esFilaPorFila)
                    {
                        string addressAbsoluta = primeraCeldaCondicion.Address;
                        string[] partes = addressAbsoluta.Split('$');
                        dirCondicionLocal = "$" + partes[1] + partes[2]; // Ej: $A1
                    }
                    else
                    {
                        dirCondicionLocal = rangoCondicion.Address; // Ej: $A$1:$A$10
                    }
                }
                finally
                {
                    ExcelHelper.LiberarCom(primeraCeldaCondicion);
                }

                // ==========================================================
                // FASE 2 Y 3: EL OPERADOR LÓGICO Y EL VALOR
                // ==========================================================
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

                if (resultadoOperador is bool && (bool)resultadoOperador == false) return;
                string operador = resultadoOperador.ToString().Trim();

                if (operador != "=" && operador != "<>" && operador != ">" && operador != "<" && operador != ">=" && operador != "<=")
                {
                    MessageBox.Show("Operador no reconocido.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                object resultadoValor = excelApp.InputBox(
                    $"Introduce el valor que completará la condición.\n(Condición actual: {operador} ___ )\n\nEjemplos: 6, Sí, X:",
                    "3. Valor del Criterio",
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 2);

                if (resultadoValor is bool && (bool)resultadoValor == false) return;

                string valorCriterio = resultadoValor.ToString().Trim();
                bool esNumero = double.TryParse(valorCriterio, out _);
                string valorFormateado = esNumero ? valorCriterio : $"\"{valorCriterio}\"";

                // ==========================================================
                // FASE 3.5 Y 4: REGLAS ADICIONALES (UX)
                // ==========================================================
                DialogResult respuestaBlanco = MessageBox.Show(
                    "¿Deseas que la matriz permanezca DESBLOQUEADA si la celda de condición está VACÍA?\n\n" +
                    "SÍ = Si está en blanco, se puede escribir.\n" +
                    "NO = Estricto (Si está en blanco, se bloquea por defecto).",
                    "Regla de Celda Vacía", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                DialogResult respuestaRojo = MessageBox.Show(
                    "¿Deseas que la celda se resalte cuando se desbloquee y esté vacía?\n\n(Ideal para los campos 'Especifique' que se vuelven obligatorios).",
                    "4. Resalte de Obligatoriedad", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                // ==========================================================
                // APLICACIÓN DE REGLAS (MOTOR: COUNTIF EN INGLÉS UNIVERSAL)
                // ==========================================================
                rangoCapturado.Validation.Delete();

                if (rangoCapturado.FormatConditions.Count > 0)
                {
                    DialogResult respFormato = MessageBox.Show(
                        "Se detectaron reglas previas (ej. Validación de Blancos).\n\n" +
                        "¿Deseas CONSERVARLAS y apilar el bloqueo encima?\n\n" +
                        "SÍ = Conservar formatos previos.\nNO = Eliminar y aplicar solo el bloqueo.",
                        "Formatos Condicionales Detectados", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

                    if (respFormato == DialogResult.No) { rangoCapturado.FormatConditions.Delete(); }
                }

                string criterioContarSi = $"\"{operador}\"&{valorFormateado}";
                string fValidacionIngles = "";
                string fSombreadoIngles = "";

                // Usamos separador de comas nativo de C#/Inglés. ExcelHelper lo convertirá automáticamente a local.
                if (respuestaBlanco == DialogResult.Yes)
                {
                    fValidacionIngles = $"=OR(ISBLANK({dirCondicionLocal}), COUNTIF({dirCondicionLocal}, {criterioContarSi})>0)";
                    fSombreadoIngles = $"=AND(NOT(ISBLANK({dirCondicionLocal})), COUNTIF({dirCondicionLocal}, {criterioContarSi})=0)";
                }
                else
                {
                    fValidacionIngles = $"=COUNTIF({dirCondicionLocal}, {criterioContarSi})>0";
                    fSombreadoIngles = $"=COUNTIF({dirCondicionLocal}, {criterioContarSi})=0";
                }

                int filaMuestra = rangoCapturado.Row;

                // TRADUCCIÓN NATÍVA (FAÇADE)
                string formulaValidacionLocal = ExcelHelper.TraducirFormulaLocal(wsActual, fValidacionIngles, filaMuestra);
                string formulaSombreadoLocal = ExcelHelper.TraducirFormulaLocal(wsActual, fSombreadoIngles, filaMuestra);

                // 1. DATA VALIDATION
                rangoCapturado.Validation.Add(Excel.XlDVType.xlValidateCustom, Excel.XlDVAlertStyle.xlValidAlertStop,
                    Excel.XlFormatConditionOperator.xlBetween, formulaValidacionLocal, Type.Missing);
                rangoCapturado.Validation.IgnoreBlank = true;
                rangoCapturado.Validation.ShowError = true;
                rangoCapturado.Validation.ErrorTitle = "Celda Bloqueada";
                rangoCapturado.Validation.ErrorMessage = $"No se permite capturar información. El flujo requiere que la condición sea {operador} {valorCriterio}.";

                // 2. FORMATO CONDICIONAL 1 (Gris Bloqueado)
                Excel.FormatCondition formatoGris = null;
                try
                {
                    formatoGris = (Excel.FormatCondition)rangoCapturado.FormatConditions.Add(
                        Excel.XlFormatConditionType.xlExpression, Type.Missing, formulaSombreadoLocal);

                    formatoGris.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                    formatoGris.Interior.Pattern = Excel.XlPattern.xlPatternCrissCross;
                    formatoGris.Interior.PatternColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Gray);
                    formatoGris.StopIfTrue = true;
                }
                finally
                {
                    ExcelHelper.LiberarCom(formatoGris);
                }

                // 3. FORMATO CONDICIONAL 2 (Alerta Azul)
                if (respuestaRojo == DialogResult.Yes)
                {
                    Excel.Range primeraCelda = null;
                    Excel.FormatCondition formatoRojo = null;
                    try
                    {
                        primeraCelda = (Excel.Range)rangoCapturado.Cells[1, 1];
                        string dirCapturada = primeraCelda.Address.Replace("$", "");

                        string fRojoIngles = $"=AND(COUNTIF({dirCondicionLocal}, {criterioContarSi})>0, ISBLANK({dirCapturada}))";
                        string formulaRojoLocal = ExcelHelper.TraducirFormulaLocal(wsActual, fRojoIngles, filaMuestra);

                        formatoRojo = (Excel.FormatCondition)rangoCapturado.FormatConditions.Add(
                            Excel.XlFormatConditionType.xlExpression, Type.Missing, formulaRojoLocal);

                        formatoRojo.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(47, 117, 181));
                        formatoRojo.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(156, 0, 6));
                        formatoRojo.StopIfTrue = true;
                    }
                    finally
                    {
                        ExcelHelper.LiberarCom(primeraCelda);
                        ExcelHelper.LiberarCom(formatoRojo);
                    }
                }

                // ==========================================================
                // --- PASO 5: RESALTE E INSTRUCCIÓN AMARILLO ---
                // ==========================================================
                // -> PEGA AQUÍ TU CÓDIGO ORIGINAL QUE OMITISTE PARA EL MENSAJE DE ALERTA E INSTRUCCIÓN <-

                string modoAplicado = esFilaPorFila ? "Fila por Fila (Paralelo)" : "Búsqueda Global";
                MessageBox.Show($"Validación de Bloqueo Dinámica aplicada con éxito.\nModo: {modoAplicado}\nRegla: {operador} {valorCriterio}", "SAVCNG", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al aplicar la Validación de Bloqueo: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // Limpieza absoluta de memoria base
                ExcelHelper.LiberarCom(rangoCondicion);
                ExcelHelper.LiberarCom(wsActual);
            }
        }
    }
}