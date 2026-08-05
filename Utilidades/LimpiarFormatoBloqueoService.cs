using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using SAVCNG_ExcelDNA.Core; // Uso obligatorio de la Fachada

namespace SAVCNG_ExcelDNA.Utilidades
{
    public class LimpiarFormatoBloqueoService : IOperacionLibro
    {
        public void Ejecutar(Excel.Application excelApp, Excel.Workbook libroCenso)
        {
            try
            {
                // Apagamos la actualización de pantalla para un borrado instantáneo y silencioso
                excelApp.ScreenUpdating = false;

                // =========================================================================
                // FASE 1: AUDITORÍA DE HOJAS PROTEGIDAS (PREVENCIÓN DE CRASH)
                // =========================================================================
                System.Collections.Generic.List<string> hojasBloqueadas = new System.Collections.Generic.List<string>();
                int totalHojas = libroCenso.Worksheets.Count;

                // RASTREO ZERO LEAKS: Ciclo 'for' para evitar enumeradores COM huérfanos
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

                // Si la hoja está bloqueada, Excel no nos dejará borrar la regla
                if (hojasBloqueadas.Count > 0)
                {
                    string listaHojas = string.Join("\n", hojasBloqueadas);
                    MessageBox.Show(
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

                for (int j = 1; j <= totalHojas; j++)
                {
                    Excel.Worksheet ws = null;
                    Excel.Range celdasHoja = null;
                    Excel.FormatConditions formatos = null;
                    bool hojaModificada = false;

                    try
                    {
                        ws = (Excel.Worksheet)libroCenso.Worksheets[j];
                        celdasHoja = ws.Cells;
                        formatos = celdasHoja.FormatConditions;

                        int totalFormatos = formatos.Count;

                        // RECORRIDO INVERSO OBLIGATORIO: Del último formato al primero
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
                            catch { /* Ignoramos errores aislados de lectura/borrado */ }
                            finally
                            {
                                ExcelHelper.LiberarCom(fc);
                            }
                        }

                        if (hojaModificada)
                        {
                            hojasLimpiadas++;
                        }
                    }
                    finally
                    {
                        // Limpieza COM por cada hoja
                        ExcelHelper.LiberarCom(formatos);
                        ExcelHelper.LiberarCom(celdasHoja);
                        ExcelHelper.LiberarCom(ws);
                    }
                }

                // =========================================================================
                // FASE 3: CONCLUSIÓN UX
                // =========================================================================
                if (hojasLimpiadas > 0)
                {
                    MessageBox.Show(
                        $"El formato de color (Celdas Bloqueadas) fue eliminado correctamente.\n\n" +
                        $"Se limpiaron {hojasLimpiadas} hoja(s) del libro.",
                        "SAVCNG - Limpieza Exitosa", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show(
                        "No se encontró el formato de color en ninguna hoja del libro.\n\nEl censo ya está limpio.",
                        "SAVCNG - Sin Cambios", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ocurrió un error crítico al limpiar los colores: " + ex.Message, "Error del Sistema", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (excelApp != null) excelApp.ScreenUpdating = true;
            }
        }
    }
}