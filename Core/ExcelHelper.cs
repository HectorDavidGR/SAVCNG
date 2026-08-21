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

                // Límite de seguridad para COM Interop: No buscar más allá de 500 filas hacia arriba
                int limiteSuperior = Math.Max(1, filaInicial - 500);

                // Buscamos desde la fila seleccionada hacia arriba en las columnas A (1) y B (2)
                for (int f = filaInicial; f >= limiteSuperior; f--)
                {
                    Excel.Range celdaA = null;
                    Excel.Range celdaB = null;
                    try
                    {
                        celdaA = hoja.Cells[f, 1]; // Columna A (Para Preguntas Clásicas)
                        celdaB = hoja.Cells[f, 2]; // Columna B (Para Complementos / Módulos)

                        object valorObjA = celdaA.Value2;
                        object valorObjB = celdaB.Value2;

                        // ==========================================================
                        // 1. REGLA PARA MÓDULOS O COMPLEMENTOS (Escaneo en Columna B)
                        // ==========================================================
                        if (valorObjB != null)
                        {
                            string valorB = valorObjB.ToString().Trim();
                            if (!string.IsNullOrEmpty(valorB))
                            {
                                string valorBUpper = valorB.ToUpper();
                                if (valorBUpper.StartsWith("COMPLEMENTO") || valorBUpper.StartsWith("MÓDULO") || valorBUpper.StartsWith("MODULO"))
                                {
                                    // Truncamos el texto para la bitácora (Ej: "Complemento 1. Título..." -> "Complemento 1")
                                    int indexPunto = valorB.IndexOf('.');
                                    if (indexPunto > 0)
                                    {
                                        return valorB.Substring(0, indexPunto).Trim();
                                    }
                                    return valorB;
                                }
                            }
                        }

                        // ==========================================================
                        // 2. REGLA PARA PREGUNTAS CLÁSICAS (Escaneo en Columna A)
                        // ==========================================================
                        if (valorObjA != null)
                        {
                            string valorA = valorObjA.ToString().Trim();
                            if (!string.IsNullOrEmpty(valorA))
                            {
                                // Si el texto empieza con un dígito y contiene punto o guion, es una pregunta INEGI
                                if (char.IsDigit(valorA[0]) && (valorA.Contains(".") || valorA.Contains("-")))
                                {
                                    return valorA;
                                }
                            }
                        }

                        // Si llegamos aquí y hay texto basura en ambas columnas, el escáner lo IGNORA y sigue subiendo.
                    }
                    finally
                    {
                        // PARCHE CRÍTICO ZERO LEAKS: Destruir los punteros COM en cada iteración
                        LiberarCom(celdaB);
                        LiberarCom(celdaA);
                    }
                }
            }
            catch { /* Si hay error, devolvemos un valor por defecto seguro */ }
            finally
            {
                // Cerramos la referencia de la hoja
                LiberarCom(hoja);
            }

            return "No Asignado / Origen Desconocido";
        }
    }
}