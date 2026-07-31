using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using SAVCNG_ExcelDNA.Core;

namespace SAVCNG_ExcelDNA.Validaciones
{
    public class ValidacionBlancos : IValidacionExcel
    {
        public void Ejecutar(Excel.Application excelApp, Excel.Workbook libroCenso, Excel.Range rangoCapturado)
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

                if (resExclusiones == null) return; // Aborto silencioso

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
                // 2. SOLICITAR UBICACIÓN DE ALERTA (Rango de Excel)
                // ==========================================================
                if (formPrincipal != null) formPrincipal.Hide();

                object resDestino = excelApp.InputBox(
                    "Selecciona el rango donde aparecerá el mensaje de alerta (Se combinará y pintará de azul automáticamente):",
                    "1. Ubicación de Alerta", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8); // 8 = Rango

                if (formPrincipal != null) formPrincipal.Show();

                if (resDestino is bool && (bool)resDestino == false) return;
                rangoDestino = (Excel.Range)resDestino;

                // ==========================================================
                // 3. SOLICITAR TEXTO DEL MENSAJE DE ADVERTENCIA
                // ==========================================================
                string resTexto = PedirInputWinForms(
                    "Escribe el mensaje de advertencia:\n(El sistema evaluará el rango capturado. Si hay celdas iniciadas pero faltan datos, exigirá completarlas)",
                    "2. Mensaje de Alerta",
                    "Favor de ingresar toda la información requerida en la pregunta");

                if (string.IsNullOrWhiteSpace(resTexto)) return;
                string textoAlerta = resTexto.Trim();

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

                    string msjExcepciones = listaCondicionesExcluidas.Count > 0
                        ? $"\n• Excepciones registradas: {textoExclusiones}"
                        : "";

                    MessageBox.Show("Validación de blancos aplicada con éxito.\n\n" +
                                    "• Si la fila no contiene ningún código de excepción y faltan campos, se alertará en azul." + msjExcepciones,
                                    "SAVCNG Arquitectura", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
                MessageBox.Show("Error en el motor de validación inteligente: " + ex.Message, "Error de Inyección", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                ExcelHelper.LiberarCom(rangoDestino);
                ExcelHelper.LiberarCom(wsActual);
            }
        }
    }
}