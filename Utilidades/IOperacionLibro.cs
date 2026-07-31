using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace SAVCNG_ExcelDNA.Utilidades
{
    // =========================================================================
    // CONTRATO ARQUITECTÓNICO: OPERACIONES GLOBALES DEL LIBRO (STRATEGY)
    // =========================================================================
    // Mentoría: Este contrato es exclusivo para acciones de mantenimiento, 
    // exportación y seguridad. No requiere un "rangoCapturado" porque su 
    // contexto de acción es el archivo (Workbook) en su totalidad.
    public interface IOperacionLibro
    {
        void Ejecutar(Excel.Application excelApp, Excel.Workbook libroCenso);
    }
}