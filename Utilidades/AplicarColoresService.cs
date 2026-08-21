using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using SAVCNG_ExcelDNA.Core; // Uso obligatorio de la Fachada

namespace SAVCNG_ExcelDNA.Utilidades
{
    public class AplicarColoresService : IOperacionLibro
    {
        public void Ejecutar(Excel.Application excelApp, Excel.Workbook libroCenso)
        {
            try
            {
                excelApp.ScreenUpdating = false;

                // =========================================================================
                // FASE 1: AUDITORÍA DE HOJAS PROTEGIDAS (PREVENCIÓN DE CRASH)
                // =========================================================================
                System.Collections.Generic.List<string> hojasBloqueadas = new System.Collections.Generic.List<string>();
                int totalHojas = libroCenso.Worksheets.Count;

                // RASTREO ZERO LEAKS: Búsqueda segura
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

                if (hojasBloqueadas.Count > 0)
                {
                    string listaHojas = string.Join("\n", hojasBloqueadas);
                    MessageBox.Show(
                        "No se puede aplicar la macro de colores porque el libro contiene hojas protegidas:\n\n" +
                        listaHojas + "\n\n" +
                        "Desprotege estas hojas desde la pestaña Revisión y vuelve a intentarlo.",
                        "Validación Interrumpida", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    return;
                }

                // =========================================================================
                // FASE 2: ESTRUCTURA DE FÓRMULAS UNIVERSALES (MATRIZ DE TRADUCCIÓN)
                // =========================================================================
                string[] formulasIngles = new string[]
                {
                    "=AND($A$1<>\"\",SEARCH(\"igual o mayor\",A1)>0)",
                    "=AND($A$1<>\"\",SEARCH(\"igual o menor\",A1)>0)",
                    "=AND($A$1<>\"\",SEARCH(\"pase a la pregunta\",A1)>0)",
                    "=AND($A$1<>\"\",SEARCH(\"en blanco\",A1)>0)",
                    "=AND($A$1<>\"\",SEARCH(\"no puede registrar\",A1)>0)",
                    "=AND($A$1<>\"\",SUM(A1)>0)",
                    "=AND($A$1<>\"\",SEARCH(\"la pregunta\",A1)>0)"
                };

                int hojasConfiguradas = 0;

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

                        // 2.1 BÚSQUEDA PREVIA
                        bool yaExisteRegra = false;
                        int totalFormatos = formatos.Count;

                        for (int k = totalFormatos; k >= 1; k--)
                        {
                            Excel.FormatCondition fcCheck = null;
                            try
                            {
                                object objFc = formatos[k];
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
                            catch { /* Silenciado si Excel bloquea lectura de fórmula específica */ }
                            finally
                            {
                                ExcelHelper.LiberarCom(fcCheck);
                            }
                        }

                        if (yaExisteRegra) continue; // Saltamos la hoja para no duplicar

                        // 2.2 INYECCIÓN DE LAS 7 REGLAS VÍA FAÇADE
                        for (int idx = 0; idx < formulasIngles.Length; idx++)
                        {
                            // Utilizamos nuestro Façade para traducir sin instanciar la celda dummy aquí
                            string formulaLocal = ExcelHelper.TraducirFormulaLocal(ws, formulasIngles[idx], 1);

                            Excel.FormatCondition nuevaRegla = null;
                            try
                            {
                                nuevaRegla = (Excel.FormatCondition)formatos.Add(
                                    Excel.XlFormatConditionType.xlExpression,
                                    Type.Missing,
                                    formulaLocal);

                                // Configuración estética según índice
                                switch (idx)
                                {
                                    case 0:
                                        nuevaRegla.Interior.PatternColorIndex = (int)Excel.Constants.xlAutomatic;
                                        nuevaRegla.Interior.ThemeColor = (int)Excel.XlThemeColor.xlThemeColorAccent6;
                                        nuevaRegla.Interior.TintAndShade = 0.599963377788629;
                                        break;
                                    case 1:
                                        nuevaRegla.Interior.PatternColorIndex = (int)Excel.Constants.xlAutomatic;
                                        nuevaRegla.Interior.ThemeColor = (int)Excel.XlThemeColor.xlThemeColorAccent2;
                                        nuevaRegla.Interior.TintAndShade = 0.599963377788629;
                                        break;
                                    case 2:
                                    case 3:
                                    case 4:
                                        nuevaRegla.Interior.PatternColorIndex = (int)Excel.Constants.xlAutomatic;
                                        nuevaRegla.Interior.Color = 16764159;
                                        nuevaRegla.Interior.TintAndShade = 0;
                                        break;
                                    case 5:
                                        nuevaRegla.Interior.PatternColorIndex = (int)Excel.Constants.xlAutomatic;
                                        nuevaRegla.Interior.ThemeColor = (int)Excel.XlThemeColor.xlThemeColorAccent5;
                                        nuevaRegla.Interior.TintAndShade = 0.799981688894314;
                                        break;
                                    case 6:
                                        nuevaRegla.Interior.PatternColorIndex = (int)Excel.Constants.xlAutomatic;
                                        nuevaRegla.Interior.ThemeColor = (int)Excel.XlThemeColor.xlThemeColorAccent4;
                                        nuevaRegla.Interior.TintAndShade = 0.599963377788629;
                                        break;
                                }
                            }
                            finally
                            {
                                ExcelHelper.LiberarCom(nuevaRegla);
                            }
                        }
                        hojasConfiguradas++;
                    }
                    finally
                    {
                        // Limpieza estricta COM por hoja procesada
                        ExcelHelper.LiberarCom(formatos);
                        ExcelHelper.LiberarCom(celdasHoja);
                        ExcelHelper.LiberarCom(ws);
                    }
                }

                MessageBox.Show(
                    $"Funcion de colorimetría aplicada con éxito.\n\nSe configuraron {hojasConfiguradas} hoja(s) del censo de forma segura.",
                    "Reglas de Control Añadidas", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error crítico al inyectar las macros de control A1: " + ex.Message, "Error del Sistema", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (excelApp != null) excelApp.ScreenUpdating = true;
            }
        }
    }
}