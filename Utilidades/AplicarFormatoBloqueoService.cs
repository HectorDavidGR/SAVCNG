using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using SAVCNG_ExcelDNA.Core; // Uso obligatorio de la Fachada

namespace SAVCNG_ExcelDNA.Utilidades
{
    public class AplicarFormatoBloqueoService : IOperacionLibro
    {
        public void Ejecutar(Excel.Application excelApp, Excel.Workbook libroCenso)
        {
            try
            {
                // Apagamos la actualización de pantalla para que el barrido sea instantáneo
                excelApp.ScreenUpdating = false;

                // =========================================================================
                // FASE 1: AUDITORÍA DE HOJAS PROTEGIDAS (PREVENCIÓN DE CRASH)
                // =========================================================================
                System.Collections.Generic.List<string> hojasBloqueadas = new System.Collections.Generic.List<string>();
                int totalHojas = libroCenso.Worksheets.Count;

                // RASTREO ZERO LEAKS: Usamos ciclo 'for'
                for (int i = 1; i <= totalHojas; i++)
                {
                    Excel.Worksheet wsCheck = null;
                    try
                    {
                        wsCheck = (Excel.Worksheet)libroCenso.Worksheets[i];
                        if (wsCheck.ProtectContents || wsCheck.ProtectDrawingObjects || wsCheck.ProtectScenarios)
                        {
                            hojasBloqueadas.Add("• " + wsCheck.Name);
                        }
                    }
                    finally
                    {
                        ExcelHelper.LiberarCom(wsCheck);
                    }
                }

                // Si hay hojas bloqueadas, abortamos y le avisamos al usuario amigablemente
                if (hojasBloqueadas.Count > 0)
                {
                    string listaHojas = string.Join("\n", hojasBloqueadas);
                    MessageBox.Show(
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

                for (int j = 1; j <= totalHojas; j++)
                {
                    Excel.Worksheet ws = null;
                    Excel.Range celdasHoja = null;
                    Excel.FormatConditions formatos = null;

                    try
                    {
                        ws = (Excel.Worksheet)libroCenso.Worksheets[j];
                        celdasHoja = ws.Cells;
                        formatos = celdasHoja.FormatConditions;
                        bool existe = false;

                        // 2.1 Búsqueda Inversa para verificar si la regla ya existe
                        int totalFormatos = formatos.Count;
                        for (int k = totalFormatos; k >= 1; k--)
                        {
                            Excel.FormatCondition fc = null;
                            try
                            {
                                object objFc = formatos[k];
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
                            catch { /* Ignoramos errores de lectura de reglas aisladas */ }
                            finally
                            {
                                ExcelHelper.LiberarCom(fc);
                            }
                        }

                        // 2.2 Si no la encontró, la inyectamos usando Façade
                        if (!existe)
                        {
                            // DELEGACIÓN A FAÇADE: Traduce la fórmula automáticamente y limpia la memoria
                            string formulaIngles = "=CELL(\"protect\",A1)";
                            string formulaLocal = ExcelHelper.TraducirFormulaLocal(ws, formulaIngles, 1);

                            Excel.FormatCondition nuevaRegla = null;
                            try
                            {
                                nuevaRegla = (Excel.FormatCondition)formatos.Add(
                                    Excel.XlFormatConditionType.xlExpression,
                                    Type.Missing,
                                    formulaLocal);

                                // Aplicamos el color RGB(255, 230, 153)
                                nuevaRegla.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 230, 153));
                                hojasModificadas++;
                            }
                            finally
                            {
                                ExcelHelper.LiberarCom(nuevaRegla);
                            }
                        }
                    }
                    finally
                    {
                        // Limpieza absoluta del Garbage Collector en cada ciclo de la hoja
                        ExcelHelper.LiberarCom(formatos);
                        ExcelHelper.LiberarCom(celdasHoja);
                        ExcelHelper.LiberarCom(ws);
                    }
                }

                // =========================================================================
                // FASE 3: CONCLUSIÓN UX
                // =========================================================================
                MessageBox.Show(
                    $"Formato de color (Celdas Bloqueadas) aplicado en todo el libro.\n\n" +
                    $"Se aplicó exitosamente en {hojasModificadas} hoja(s) que no contaban con él.",
                    "SAVCNG - Macro de Colores", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ocurrió un error crítico al aplicar la macro de colores: " + ex.Message, "Error del Sistema", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (excelApp != null) excelApp.ScreenUpdating = true;
            }
        }
    }
}