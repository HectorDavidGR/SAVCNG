using SAVCNG_ExcelDNA.Core;
using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace SAVCNG_ExcelDNA.Validaciones
{
    // =========================================================================
    // CONTRATO ARQUITECTÓNICO: PATRÓN STRATEGY
    // =========================================================================
    // Mentoría: Una "Interfaz" es como un molde o un contrato.
    // Obliga a que cualquier clase nueva de validación que creemos en el futuro 
    // tenga exactamente el mismo método de entrada, garantizando que el 
    // botón "Aplicar" siempre sepa cómo ejecutarlas sin importar de qué validación se trate.
    public interface IValidacionExcel
    {
        // Único punto de entrada. Recibe las 3 cosas vitales que necesita una validación:
        // 1. El motor de Excel (Para InputBoxes o ScreenUpdating)
        // 2. El Libro (Para la bitácora)
        // 3. El Rango (Las celdas seleccionadas por el usuario)
        ResultadoValidacion Ejecutar(Excel.Application excelApp, Excel.Workbook libroCenso, Excel.Range rangoCapturado);
    }
}