using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using SAVCNG_ExcelDNA.Core; // Uso obligatorio de la Fachada

namespace SAVCNG_ExcelDNA.Utilidades
{
    public class DesprotegerLibroService : IOperacionLibro
    {
        // Dependencia inyectada
        private readonly int _anioSeleccionado;

        // Inyección por Constructor
        public DesprotegerLibroService(int anioSeleccionado)
        {
            _anioSeleccionado = anioSeleccionado;
        }

        public void Ejecutar(Excel.Application excelApp, Excel.Workbook libroCenso)
        {
            try
            {
                // 1. VALIDACIÓN: Contabilizar cuántas hojas están protegidas actualmente
                int totalHojas = libroCenso.Worksheets.Count;
                int hojasProtegidasAlInicio = 0;

                // RASTREO ZERO LEAKS: Ciclo 'for' para evitar enumeradores COM
                for (int i = 1; i <= totalHojas; i++)
                {
                    Excel.Worksheet hojaCheck = null;
                    try
                    {
                        hojaCheck = (Excel.Worksheet)libroCenso.Worksheets[i];
                        if (hojaCheck.ProtectContents)
                        {
                            hojasProtegidasAlInicio++;
                        }
                    }
                    finally
                    {
                        ExcelHelper.LiberarCom(hojaCheck);
                    }
                }

                // EXCEPCIÓN A: Si NINGUNA hoja está protegida
                if (hojasProtegidasAlInicio == 0)
                {
                    MessageBox.Show("No se aplicó la acción de desproteger debido a que las hojas no estaban protegidas.", "Aviso de Desprotección", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // 2. Reconstruimos la contraseña dinámica para desproteger
                string passwordEstructurada = "njm" + _anioSeleccionado.ToString();

                int hojasDesprotegidasExitosamente = 0;
                int hojasFallidas = 0;

                excelApp.ScreenUpdating = false;

                // 3. Recorremos el libro para aplicar la desprotección
                for (int j = 1; j <= totalHojas; j++)
                {
                    Excel.Worksheet hoja = null;
                    try
                    {
                        hoja = (Excel.Worksheet)libroCenso.Worksheets[j];

                        if (hoja.ProtectContents)
                        {
                            hoja.Unprotect(passwordEstructurada);
                            hojasDesprotegidasExitosamente++;
                        }
                    }
                    catch (Exception exHoja)
                    {
                        // Si la contraseña dinámica no coincide con la clave real de la hoja
                        hojasFallidas++;
                        System.Diagnostics.Debug.WriteLine($"No se pudo desproteger la hoja '{hoja?.Name}': {exHoja.Message}");
                    }
                    finally
                    {
                        ExcelHelper.LiberarCom(hoja);
                    }
                }

                excelApp.ScreenUpdating = true;

                // 4. Mensajes informativos de finalización
                if (hojasFallidas > 0)
                {
                    MessageBox.Show(
                        $"Se desprotegieron {hojasDesprotegidasExitosamente} hojas correctamente.\n\n" +
                        $"Sin embargo, {hojasFallidas} hojas no pudieron ser desprotegidas con la clave '{passwordEstructurada}'. " +
                        "Verifica si corresponden a otro año o clave previa.",
                        "Desprotección Parcial",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                }
                else
                {
                    MessageBox.Show("Se han desprotegido con éxito todas las hojas del libro de trabajo.", "Libro Desprotegido", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error crítico al intentar desproteger las hojas: " + ex.Message, "Error del Sistema", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (excelApp != null) excelApp.ScreenUpdating = true;
            }
        }
    }
}