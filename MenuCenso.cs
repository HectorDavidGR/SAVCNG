using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using Microsoft.Office.Interop.Excel;

namespace SAVCNG_ExcelDNA
{
    // [ComVisible(true)] es obligatorio para que Excel pueda ver nuestro menú
    [ComVisible(true)]
    public class MenuCenso : ExcelRibbon
    {
        // 1. ESTA PARTE DIBUJA EL MENÚ EN EXCEL USANDO CÓDIGO XML
        public override string GetCustomUI(string RibbonID)
        {
            return @"
            <customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
              <ribbon>
                <tabs>
                  <tab id='TabCenso' label='SAVCNG - Censo'>
                    
                    <group id='Grupo1' label='Archivos'>
                      <button id='btnCargar' label='Cargar Censo' size='large' onAction='BotonCargar_Clic' imageMso='FileOpen' />
                    </group>

                  </tab>
                </tabs>
              </ribbon>
            </customUI>";
        }

        // 2. ESTA PARTE ES LA QUE SE EJECUTA CUANDO LE DAS CLIC AL BOTÓN
        public void BotonCargar_Clic(IRibbonControl control)
        {
            // Creamos la ventana para buscar el archivo
            OpenFileDialog fd = new OpenFileDialog();
            fd.Filter = "Archivos de Excel|*.xlsx;*.xlsm;*.xlsb";
            fd.Title = "Seleccione el archivo del censo de gobierno";

            // Si el usuario elige un archivo y le da a "Abrir"...
            if (fd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    // Obtenemos la aplicación de Excel
                    Microsoft.Office.Interop.Excel.Application excelApp = (Microsoft.Office.Interop.Excel.Application)ExcelDnaUtil.Application;

                    excelApp.ScreenUpdating = false;

                    // ABRIMOS EL LIBRO Y LO GUARDAMOS EN UNA VARIABLE LLAMADA 'libroAbierto'
                    Workbook libroAbierto = excelApp.Workbooks.Open(fd.FileName, IgnoreReadOnlyRecommended: true);

                    excelApp.ScreenUpdating = true;

                    // === AQUÍ ESTÁ LA MAGIA NUEVA ===
                    // Instanciamos nuestra ventana y le pasamos el libro que acabamos de abrir
                    FrmValidaciones miVentana = new FrmValidaciones(libroAbierto);

                    // Usamos Show() en lugar de ShowDialog() para que la ventana sea flotante 
                    // y el usuario pueda seguir haciendo clic en las celdas de Excel por detrás.
                    miVentana.Show();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ocurrió un error al cargar el archivo: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
    }
}
