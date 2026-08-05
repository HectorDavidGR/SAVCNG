using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using SAVCNG_ExcelDNA.Core; // Uso obligatorio de la Fachada

namespace SAVCNG_ExcelDNA.Utilidades
{
    public class LimpiarColoresService : IOperacionLibro
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

                // RASTREO ZERO LEAKS: Búsqueda segura con ciclo 'for'
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
                        "No se puede remover el formato de colores debido a que el libro tiene hojas protegidas:\n\n" +
                        listaHojas + "\n\n" +
                        "Desprotege estas hojas desde la pestaña Revisión y reintenta la operación.",
                        "Operación Interrumpida", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    return;
                }

                // =========================================================================
                // FASE 2: PURGA Y ELIMINACIÓN DE FORMATOS MEDIANTE RECORRIDO INVERSO
                // =========================================================================
                int hojasLimpiadas = 0;

                for (int j = 1; j <= totalHojas; j++)
                {
                    Excel.Worksheet ws = null;
                    Excel.Range celdasHoja = null;
                    Excel.FormatConditions formatos = null;
                    bool cambioDetectado = false;

                    try
                    {
                        ws = (Excel.Worksheet)libroCenso.Worksheets[j];
                        celdasHoja = ws.Cells;
                        formatos = celdasHoja.FormatConditions;

                        int totalFormatos = formatos.Count;

                        // ESCANEO INVERSO OBLIGATORIO: Evita desajustes en el puntero de la lista al borrar
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
                                            string formulaLimpia = formulaCondicion.Replace(" ", "").ToUpper();

                                            // Buscamos nuestro patrón común anclado a la celda de control $A$1
                                            if (formulaLimpia.Contains("$A$1<>\"\""))
                                            {
                                                fc.Delete();
                                                cambioDetectado = true;
                                            }
                                        }
                                    }
                                }
                            }
                            catch { /* Silenciado si Excel deniega lectura en la purga */ }
                            finally
                            {
                                ExcelHelper.LiberarCom(fc);
                            }
                        }

                        if (cambioDetectado)
                        {
                            hojasLimpiadas++;
                        }
                    }
                    finally
                    {
                        // Liberación estricta COM por hoja
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
                        $"Se eliminaron por completo las reglas de colorimetria vinculadas al archivo.\n\nSe limpiaron exitosamente {hojasLimpiadas} hoja(s).",
                        "SAVCNG - Limpieza Exitosa", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show(
                        "No se detectaron formatos condicionales vinculados al archivo actual.\n\nEl libro ya se encuentra limpio.",
                        "SAVCNG - Sin Cambios", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ocurrió un error inesperado al limpiar las macros de formato: " + ex.Message, "Error del Sistema", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (excelApp != null) excelApp.ScreenUpdating = true;
            }
        }
    }
}