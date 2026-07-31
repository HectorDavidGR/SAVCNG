using System;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace SAVCNG_ExcelDNA.Core
{
    // =========================================================================
    // FACHADA DE SERVICIOS EXCEL (PATRÓN FAÇADE)
    // =========================================================================
    // Mentoría: Al ser "public static", podemos llamar a estas herramientas desde 
    // cualquier lugar del proyecto sin tener que instanciarlas con "new".
    public static class ExcelHelper
    {
        // ---------------------------------------------------------------------
        // FASE 1: TRUCO DEL DUMMY CELL CENTRALIZADO (UNIVERSALIDAD)
        // ---------------------------------------------------------------------
        // Este método toma una fórmula en inglés universal, busca dinámicamente
        // la última columna de la hoja (para evitar el error 0x800A03EC en plantillas .xls),
        // inyecta, traduce y limpia sin dejar rastro.
        public static string TraducirFormulaLocal(Excel.Worksheet ws, string formulaIngles, int filaBase = 1)
        {
            Excel.Range columnasHoja = null;
            Excel.Range celdaDummy = null;
            string formulaTraducida = "";

            try
            {
                columnasHoja = ws.Columns;
                int ultimaColumna = columnasHoja.Count;

                celdaDummy = (Excel.Range)ws.Cells[filaBase, ultimaColumna];
                celdaDummy.Formula = formulaIngles;
                formulaTraducida = celdaDummy.FormulaLocal;
                celdaDummy.Clear();
            }
            catch (Exception ex)
            {
                throw new Exception("Fallo en ExcelHelper al traducir fórmula nativa: " + ex.Message);
            }
            finally
            {
                // Limpieza COM rigurosa local
                LiberarCom(columnasHoja);
                LiberarCom(celdaDummy);
            }

            return formulaTraducida;
        }

        // ---------------------------------------------------------------------
        // FASE 2: GESTOR CENTRAL DE MEMORIA (ZERO LEAKS)
        // ---------------------------------------------------------------------
        // Este método atrapa cualquier objeto COM y lo destruye de manera segura.
        // Si el objeto ya estaba destruido o es nulo, el bloque try-catch lo 
        // silencia elegantemente, evitando crasheos en la interfaz (UI).
        public static void LiberarCom(object obj)
        {
            if (obj != null)
            {
                try
                {
                    Marshal.ReleaseComObject(obj);
                }
                catch { /* Silenciado: Prevención de doble liberación */ }
                finally
                {
                    obj = null;
                }
            }
        }
        // ---------------------------------------------------------------------
        // FASE 3: LECTOR DE PREGUNTAS (STATE MANAGEMENT UNIVERSAL)
        // ---------------------------------------------------------------------
        public static string ObtenerNumeroPregunta(Excel.Range rango)
        {
            Excel.Worksheet hoja = null;
            try
            {
                hoja = rango.Worksheet;
                int filaInicial = rango.Row;

                // Buscamos desde la fila seleccionada hacia arriba en la columna 1 (Columna A)
                for (int f = filaInicial; f >= 1; f--)
                {
                    Excel.Range celdaA = null;
                    try
                    {
                        celdaA = hoja.Cells[f, 1];
                        object valor = celdaA.Value2;

                        if (valor != null && !string.IsNullOrEmpty(valor.ToString().Trim()))
                        {
                            return valor.ToString().Trim();
                        }
                    }
                    finally
                    {
                        // PARCHE CRÍTICO ZERO LEAKS: Destruir el puntero en cada paso
                        LiberarCom(celdaA);
                    }
                }
            }
            catch { /* Si hay error, devolvemos vacío */ }
            finally
            {
                // Cerramos la referencia de la hoja
                LiberarCom(hoja);
            }

            return "(no encontrada)";
        }
    }
}