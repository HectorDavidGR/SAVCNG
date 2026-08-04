using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using SAVCNG_ExcelDNA.Core;

namespace SAVCNG_ExcelDNA.Validaciones
{
    public class ValidacionBlancos : IValidacionExcel
    {
        // 1. EL CONTRATO EXIGE DEVOLVER EL DTO
        public ResultadoValidacion Ejecutar(Excel.Application excelApp, Excel.Workbook libroCenso, Excel.Range rangoCapturado)
        {
            Excel.Worksheet wsActual = null;
            Excel.Range rangoDestino = null;

            // Extraemos el puntero del formulario principal para poder ocultarlo/mostrarlo
            Form formPrincipal = Application.OpenForms.Count > 0 ? Application.OpenForms[0] : null;

            try
            {
                wsActual = (Excel.Worksheet)rangoCapturado.Worksheet;

                // ==========================================================
                // FUNCIÓN LOCAL PARA DIÁLOGO NATIVO (Aislada de la Clase Dios)
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

                        if (lbl.Height > 70)
                        {
                            txt.Top = lbl.Bottom + 10;
                            btnOk.Top = txt.Bottom + 15;
                            btnCancel.Top = txt.Bottom + 15;
                            dlg.Height = btnOk.Bottom + 45;
                        }

                        // Al no tener "this", usamos el ShowDialog nativo auto-centrado
                        return dlg.ShowDialog() == DialogResult.OK ? txt.Text : null;
                    }
                };

                // ==========================================================
                // 1. PREGUNTA INICIAL: CASOS DE EXCEPCIÓN / EXCLUSIÓN
                // ==========================================================
                string mensajePrompt = "Indica los valores para los cuales NO se debe aplicar la validación de Blancos.\n\n" +
                                       "• Si son varios, sepáralos por comas (Ejemplo: 2, 9).\n" +
                                       "• Deje vacía la captura en caso de no requerir alguna excepción.\n";

                string resExclusiones = PedirInputWinForms(mensajePrompt, "Excepciones de Validación (Opcional)", "2, 9");

                // ABORTO TOTAL 1: DTO
                if (resExclusiones == null)
                {
                    return new ResultadoValidacion { Exito = false, Mensaje = "Configuración de exclusiones cancelada.", AlertaInyectada = false };
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
                            listaCondicionesExcluidas.Add($"COUNTIF({{0}}, {valFormateado})=0");
                        }
                    }
                }

                // ==========================================================
                // 1.5. RECOLECCIÓN LIGERA DE PREGUNTAS (STATE MANAGEMENT)
                // ==========================================================
                System.Collections.Generic.HashSet<string> preguntasUnicas = new System.Collections.Generic.HashSet<string>();
                int totalAreas = rangoCapturado.Areas.Count;

                // RASTREO ZERO LEAKS: Ciclo estricto con liberación de COM
                for (int i = 1; i <= totalAreas; i++)
                {
                    Excel.Range areaIndividual = null;
                    try
                    {
                        areaIndividual = (Excel.Range)rangoCapturado.Areas[i];
                        string idPregunta = ExcelHelper.ObtenerNumeroPregunta(areaIndividual);

                        if (!string.IsNullOrEmpty(idPregunta))
                        {
                            preguntasUnicas.Add(idPregunta);
                        }
                    }
                    finally
                    {
                        ExcelHelper.LiberarCom(areaIndividual);
                    }
                }

                // ==========================================================
                // 2. SOLICITAR UBICACIÓN DE ALERTA (Rango de Excel)
                // ==========================================================
                if (formPrincipal != null) formPrincipal.Hide();

                object resDestino = excelApp.InputBox(
                    "Selecciona el rango donde aparecerá el mensaje de alerta (Se combinará y pintará de azul automáticamente):",
                    "1. Ubicación de Alerta", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8); // 8 = Rango

                if (formPrincipal != null) formPrincipal.Show();

                // ABORTO TOTAL 2: DTO
                if (resDestino is bool && (bool)resDestino == false)
                {
                    return new ResultadoValidacion { Exito = false, Mensaje = "Selección de destino de alerta cancelada.", AlertaInyectada = false };
                }

                rangoDestino = (Excel.Range)resDestino;

                // ==========================================================
                // 3. SOLICITAR TEXTO DEL MENSAJE DE ADVERTENCIA
                // ==========================================================
                string resTexto = PedirInputWinForms(
                    "Escribe el mensaje de advertencia:\n(El sistema evaluará el rango capturado. Si hay celdas iniciadas pero faltan datos, exigirá completarlas)",
                    "2. Mensaje de Alerta",
                    "Favor de ingresar toda la información requerida en la pregunta");

                // ABORTO TOTAL 3: DTO
                if (string.IsNullOrWhiteSpace(resTexto))
                {
                    return new ResultadoValidacion { Exito = false, Mensaje = "Configuración del mensaje de alerta cancelada.", AlertaInyectada = false };
                }

                string textoAlerta = resTexto.Trim();

                // Apagamos la UI para procesar masivamente sin parpadeos
                excelApp.ScreenUpdating = false;

                // ==========================================================
                // 4. CONSTRUCCIÓN INTELIGENTE ANTI-CELDAS COMBINADAS (MERGED)
                // ==========================================================
                int filaInicial = rangoCapturado.Row;

                System.Collections.Generic.List<string> listaLlenasRelativas = new System.Collections.Generic.List<string>();
                System.Collections.Generic.List<string> listaVaciasRelativas = new System.Collections.Generic.List<string>();
                System.Collections.Generic.List<string> listaLlenasAbsolutas = new System.Collections.Generic.List<string>();

                int c = 1;
                while (c <= rangoCapturado.Columns.Count)
                {
                    Excel.Range celdaActual = null;
                    Excel.Range areaCombinada = null;

                    try
                    {
                        celdaActual = (Excel.Range)rangoCapturado.Cells[1, c];
                        string colLetra = celdaActual.Address.Split('$')[1];

                        string dirRelativaCelda = $"{colLetra}{filaInicial}";
                        string dirAbsolutaColumna = $"${colLetra}{filaInicial}";

                        listaLlenasRelativas.Add($"({dirRelativaCelda}<>\"\")");
                        listaVaciasRelativas.Add($"({dirRelativaCelda}=\"\")");
                        listaLlenasAbsolutas.Add($"({dirAbsolutaColumna}<>\"\")");

                        if ((bool)celdaActual.MergeCells)
                        {
                            areaCombinada = celdaActual.MergeArea;
                            int columnasCombinadas = areaCombinada.Columns.Count;
                            c += columnasCombinadas;
                        }
                        else
                        {
                            c++;
                        }
                    }
                    finally
                    {
                        // Excelente uso original de Zero Leaks, mantenido intacto.
                        ExcelHelper.LiberarCom(areaCombinada);
                        ExcelHelper.LiberarCom(celdaActual);
                    }
                }

                string sumaLlenasRel = string.Join("+", listaLlenasRelativas);
                string sumaVaciasRel = string.Join("+", listaVaciasRelativas);
                string sumaLlenasAbs = string.Join("+", listaLlenasAbsolutas);

                Excel.Range primeraCeldaCapturada = null;
                Excel.Range ultimaCeldaCapturada = null;

                try
                {
                    primeraCeldaCapturada = (Excel.Range)rangoCapturado.Cells[1, 1];
                    ultimaCeldaCapturada = (Excel.Range)rangoCapturado.Cells[1, rangoCapturado.Columns.Count];

                    string colIniLetra = primeraCeldaCapturada.Address.Split('$')[1];
                    string colFinLetra = ultimaCeldaCapturada.Address.Split('$')[1];

                    string rangoFilaRelativo = $"{colIniLetra}{filaInicial}:{colFinLetra}{filaInicial}";
                    string rangoFilaAbsoluto = $"${colIniLetra}{filaInicial}:${colFinLetra}{filaInicial}";

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
                    // 5. INYECCIÓN DE FÓRMULA DE ALERTA TRADUCIDA (USANDO FAÇADE)
                    // ==========================================================
                    string formulaGlobalIngles = $"=IF(AND(({sumaLlenasRel})>0, ({sumaVaciasRel})>0{clausulaExclusionRel}), \"{textoAlerta}\", \"\")";

                    // Traducción Nativa 
                    string formulaGlobalLocal = ExcelHelper.TraducirFormulaLocal(wsActual, formulaGlobalIngles, filaInicial);

                    if (rangoDestino.Count > 1) { rangoDestino.Merge(); }

                    rangoDestino.Font.Name = "Arial";
                    rangoDestino.Font.Size = 10;
                    rangoDestino.Font.Bold = true;
                    rangoDestino.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(0, 112, 192));
                    rangoDestino.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

                    rangoDestino.FormulaLocal = formulaGlobalLocal;

                    // ==========================================================
                    // 6. INYECCIÓN DEL FORMATO CONDICIONAL (USANDO FAÇADE)
                    // ==========================================================
                    if (rangoCapturado.FormatConditions.Count > 0)
                    {
                        // Restauramos la UI temporalmente para permitir interacción
                        excelApp.ScreenUpdating = true;
                        DialogResult respFormato = MessageBox.Show(
                            "Se detectaron reglas de formato condicional previas (ej. Reglas de Bloqueo).\n\n" +
                            "¿Deseas CONSERVAR las reglas existentes e integrar la técnica de blancos?\n\n" +
                            "SÍ = Conservar formatos previos (Evita borrar tus bloques grises).\n" +
                            "NO = Eliminar formatos previos y aplicar únicamente el formato de blancos.",
                            "Formatos Condicionales Detectados",
                            MessageBoxButtons.YesNoCancel,
                            MessageBoxIcon.Warning);
                        excelApp.ScreenUpdating = false;

                        // ABORTO TOTAL 4: DTO
                        if (respFormato == DialogResult.Cancel)
                        {
                            return new ResultadoValidacion { Exito = false, Mensaje = "Operación cancelada por el usuario.", AlertaInyectada = false };
                        }
                        else if (respFormato == DialogResult.No)
                        {
                            rangoCapturado.FormatConditions.Delete();
                        }
                    }
                    else
                    {
                        rangoCapturado.FormatConditions.Delete();
                    }

                    string colPrimeraLetra = primeraCeldaCapturada.Address.Split('$')[1];
                    string dirPrimeraRelativa = $"{colPrimeraLetra}{filaInicial}";

                    string formulaCondicionalIngles = $"=AND({dirPrimeraRelativa}=\"\", ({sumaLlenasAbs})>0{clausulaExclusionAbs})";

                    // Traducción Nativa
                    string formulaCondicionalLocal = ExcelHelper.TraducirFormulaLocal(wsActual, formulaCondicionalIngles, filaInicial);

                    Excel.FormatCondition formatoAzul = null;
                    try
                    {
                        formatoAzul = (Excel.FormatCondition)rangoCapturado.FormatConditions.Add(
                            Excel.XlFormatConditionType.xlExpression, Type.Missing, formulaCondicionalLocal);

                        formatoAzul.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(47, 117, 181));
                    }
                    finally
                    {
                        ExcelHelper.LiberarCom(formatoAzul);
                    }

                    // ==========================================================
                    // 7. AUDITORÍA Y CONCLUSIÓN (EMPAQUETADA EN DTO)
                    // ==========================================================
                    string preguntasDetectadas = preguntasUnicas.Count > 0 ? string.Join(", ", preguntasUnicas) : "ND";

                    AuditoriaCenso.RegistrarAccion(
                        libroCenso,
                        preguntasDetectadas,
                        "Validación de Blancos",
                        rangoCapturado.Address.Replace("$", ""),
                        "Fórmulas + FormatConditions"
                    );

                    string msjExcepciones = listaCondicionesExcluidas.Count > 0
                        ? $"\n• Excepciones registradas: {textoExclusiones}"
                        : "";

                    string mensajeExito = "Validación de blancos aplicada con éxito.\n\n" +
                                          "• Si la fila no contiene ningún código de excepción y faltan campos, se alertará en azul." + msjExcepciones;

                    // ÉXITO TOTAL: DTO
                    return new ResultadoValidacion { Exito = true, Mensaje = mensajeExito, AlertaInyectada = true };
                }
                finally
                {
                    // Liberación estricta al final del bloque
                    ExcelHelper.LiberarCom(primeraCeldaCapturada);
                    ExcelHelper.LiberarCom(ultimaCeldaCapturada);
                }
            }
            catch (Exception ex)
            {
                if (formPrincipal != null) formPrincipal.Show();

                // ERROR CRÍTICO: DTO
                return new ResultadoValidacion
                {
                    Exito = false,
                    Mensaje = "Error en el motor de validación inteligente: " + ex.Message,
                    AlertaInyectada = false
                };
            }
            finally
            {
                ExcelHelper.LiberarCom(rangoDestino);
                ExcelHelper.LiberarCom(wsActual);
                if (excelApp != null) excelApp.ScreenUpdating = true;
            }
        }
    }
}