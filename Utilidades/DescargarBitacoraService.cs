using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using SAVCNG_ExcelDNA.Core; // Uso obligatorio de la Fachada

namespace SAVCNG_ExcelDNA.Utilidades
{
    public class DescargarBitacoraService : IOperacionLibro
    {
        public void Ejecutar(Excel.Application excelApp, Excel.Workbook libroCenso)
        {
            Excel.Worksheet wsLog = null;
            Excel.Workbook nuevoLibro = null;
            Excel.Worksheet wsCopia = null;
            Excel.Worksheet hojaOriginal = null;

            try
            {
                // 1. UX: Indicamos que el sistema está trabajando en segundo plano
                Cursor.Current = Cursors.WaitCursor;

                // 2. CAPTURA DE ESTADO: Guardamos la hoja donde está parado el usuario
                hojaOriginal = (Excel.Worksheet)libroCenso.ActiveSheet;

                // 3. RASTREO ZERO LEAKS: Búsqueda segura en la colección de hojas
                for (int i = 1; i <= libroCenso.Worksheets.Count; i++)
                {
                    Excel.Worksheet sheetTemp = null;
                    try
                    {
                        sheetTemp = (Excel.Worksheet)libroCenso.Worksheets[i];
                        if (sheetTemp.Name == "SAVCNG_SysLog")
                        {
                            wsLog = sheetTemp;
                            break; // Encontramos la hoja, rompemos el ciclo
                        }
                    }
                    finally
                    {
                        // Si la hoja iterada no es nuestra bitácora, la destruimos inmediatamente
                        if (wsLog != sheetTemp)
                        {
                            ExcelHelper.LiberarCom(sheetTemp);
                        }
                    }
                }

                if (wsLog == null)
                {
                    MessageBox.Show("Aún no existen registros de validaciones en este censo.", "Bitácora Vacía", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // 4. MODO SILENCIOSO: Apagamos pantalla y bloqueamos cuadros de diálogo
                excelApp.ScreenUpdating = false;
                excelApp.DisplayAlerts = false;

                // 5. Clonación en memoria RAM
                nuevoLibro = excelApp.Workbooks.Add(Type.Missing);
                wsLog.Visible = Excel.XlSheetVisibility.xlSheetVisible;
                wsLog.Copy(Before: nuevoLibro.Worksheets[1]);
                wsLog.Visible = Excel.XlSheetVisibility.xlSheetVeryHidden; // Restauramos seguridad

                // 6. Configuración visual del archivo a exportar
                wsCopia = (Excel.Worksheet)nuevoLibro.Worksheets[1];
                wsCopia.Name = "Auditoria_" + DateTime.Now.ToString("ddMMyy");
                wsCopia.Columns.AutoFit();
                wsCopia.Application.ActiveWindow.SplitRow = 1;
                wsCopia.Application.ActiveWindow.FreezePanes = true;

                // =========================================================================
                // 7. MOTOR DE I/O: RESOLUCIÓN DE RUTA Y GUARDADO AUTOMÁTICO
                // =========================================================================
                string rutaPerfil = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
                string rutaDescargas = System.IO.Path.Combine(rutaPerfil, "Downloads");

                string nombreArchivo = $"SAVCNG_Bitacora_{DateTime.Now.ToString("yyyyMMdd_HHmmss")}.xlsx";
                string rutaCompleta = System.IO.Path.Combine(rutaDescargas, nombreArchivo);

                nuevoLibro.SaveAs(rutaCompleta, Excel.XlFileFormat.xlOpenXMLWorkbook, Type.Missing, Type.Missing,
                                  Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange,
                                  Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                // Cerramos el libro temporal inmediatamente
                nuevoLibro.Close(false);

                // 8. RESTAURACIÓN DEL ESTADO EN EL LIBRO ORIGEN
                if (hojaOriginal != null)
                {
                    hojaOriginal.Activate();
                }

                // 9. Notificación UX de éxito (Sin depender de 'this')
                MessageBox.Show(
                    $"La bitácora ha sido exportada de forma automática.\n\nPuedes encontrar el archivo en tu carpeta de Descargas:\n\n{nombreArchivo}",
                    "SAVCNG - Descarga Exitosa", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error crítico al intentar guardar la bitácora: " + ex.Message, "Fallo de I/O", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // 10. Limpieza estricta del Garbage Collector (Vía Façade)
                ExcelHelper.LiberarCom(wsCopia);
                ExcelHelper.LiberarCom(nuevoLibro);
                ExcelHelper.LiberarCom(wsLog);
                ExcelHelper.LiberarCom(hojaOriginal);

                // 11. Restauración de los motores de Excel y Windows Forms
                if (excelApp != null)
                {
                    excelApp.DisplayAlerts = true;
                    excelApp.ScreenUpdating = true;
                }
                Cursor.Current = Cursors.Default;
            }
        }
    }
}