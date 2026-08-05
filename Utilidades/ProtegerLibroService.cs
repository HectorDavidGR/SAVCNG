using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using SAVCNG_ExcelDNA.Core; // Uso obligatorio de la Fachada

namespace SAVCNG_ExcelDNA.Utilidades
{
    public class ProtegerLibroService : IOperacionLibro
    {
        // Variable privada para almacenar la dependencia inyectada
        private readonly int _anioSeleccionado;

        // Inyección de Dependencias por Constructor
        public ProtegerLibroService(int anioSeleccionado)
        {
            _anioSeleccionado = anioSeleccionado;
        }

        public void Ejecutar(Excel.Application excelApp, Excel.Workbook libroCenso)
        {
            try
            {
                // 1. Contabilizamos el estado inicial de protección de las hojas
                int totalHojas = libroCenso.Worksheets.Count;
                int hojasYaProtegidasAlInicio = 0;

                // RASTREO ZERO LEAKS: Usamos ciclo 'for' para evitar enumeradores COM huérfanos
                for (int i = 1; i <= totalHojas; i++)
                {
                    Excel.Worksheet hojaCheck = null;
                    try
                    {
                        hojaCheck = (Excel.Worksheet)libroCenso.Worksheets[i];
                        if (hojaCheck.ProtectContents)
                        {
                            hojasYaProtegidasAlInicio++;
                        }
                    }
                    finally
                    {
                        ExcelHelper.LiberarCom(hojaCheck);
                    }
                }

                // EXCEPCIÓN A: Si TODAS las hojas ya contaban con protección previa
                if (hojasYaProtegidasAlInicio == totalHojas)
                {
                    MessageBox.Show("El libro ya está protegido por lo que no se aplicó protección con la clave del año elegido", "Aviso de Protección", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // 2. Estructuramos la contraseña concatenando "njm" con el año inyectado
                string passwordEstructurada = "njm" + _anioSeleccionado.ToString();
                int hojasProtegidasExitosamente = 0;

                excelApp.ScreenUpdating = false;

                // 3. Recorremos el libro para aplicar la protección únicamente a las hojas desprotegidas
                for (int j = 1; j <= totalHojas; j++)
                {
                    Excel.Worksheet hoja = null;
                    try
                    {
                        hoja = (Excel.Worksheet)libroCenso.Worksheets[j];

                        // Si la hoja no está protegida, aplicamos el bloqueo algorítmico
                        if (!hoja.ProtectContents)
                        {
                            hoja.Protect(
                                Password: passwordEstructurada,
                                DrawingObjects: true,
                                Contents: true,
                                Scenarios: true,
                                UserInterfaceOnly: false
                            );
                            hojasProtegidasExitosamente++;
                        }
                    }
                    catch (Exception exHoja)
                    {
                        System.Diagnostics.Debug.WriteLine($"No se pudo proteger la hoja '{hoja.Name}': {exHoja.Message}");
                    }
                    finally
                    {
                        ExcelHelper.LiberarCom(hoja);
                    }
                }

                excelApp.ScreenUpdating = true;

                // EXCEPCIÓN B: Si sólo ALGUNAS hojas estaban protegidas previamente
                if (hojasYaProtegidasAlInicio > 0)
                {
                    MessageBox.Show($"Sólo se protegieron {hojasProtegidasExitosamente} hojas debido a que las demás ya estaban protegidas con una clave previa.", "Protección Parcial", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    // Flujo Normal: Si ninguna estaba protegida y se bloquearon todas
                    MessageBox.Show($"Se han protegido con éxito todas las hojas del libro.\n\nContraseña aplicada: {passwordEstructurada}", "Libro Protegido", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error crítico al intentar proteger las hojas: " + ex.Message, "Error del Sistema", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (excelApp != null) excelApp.ScreenUpdating = true;
            }
        }
    }
}