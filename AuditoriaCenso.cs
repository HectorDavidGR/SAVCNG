using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace SAVCNG_ExcelDNA
{
    public static class AuditoriaCenso
    {
        private const string NOMBRE_HOJA_LOG = "SAVCNG_SysLog";

        // =========================================================================
        // 1. MOTOR DE ESCRITURA: Registra una nueva validación en el archivo
        // =========================================================================
        // =========================================================================
        // 1. MOTOR DE ESCRITURA: Registra una nueva validación en el archivo
        // =========================================================================
        public static void RegistrarAccion(Excel.Workbook libro, string pregunta, string tipoValidacion, string direccionRango)
        {
            Excel.Worksheet wsLog = null;
            Excel.Range rangoInsertar = null;
            Excel.Worksheet hojaOriginal = null; // NUEVO: Puntero para preservación de estado

            try
            {
                // 1. CAPTURA DE ESTADO (UX): Guardamos en memoria cuál es la hoja activa actual
                hojaOriginal = (Excel.Worksheet)libro.ActiveSheet;

                // 2. Buscamos si la hoja ya existe
                foreach (Excel.Worksheet sheet in libro.Worksheets)
                {
                    if (sheet.Name == NOMBRE_HOJA_LOG)
                    {
                        wsLog = sheet;
                        break;
                    }
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet);
                }

                // 3. Si no existe, la creamos y la blindamos
                if (wsLog == null)
                {
                    // Al ejecutar .Add(), Excel roba el foco nativamente hacia esta nueva hoja
                    wsLog = (Excel.Worksheet)libro.Worksheets.Add(After: libro.Worksheets[libro.Worksheets.Count]);
                    wsLog.Name = NOMBRE_HOJA_LOG;

                    // UX/Seguridad: La hacemos "VeryHidden" para protegerla del usuario
                    wsLog.Visible = Excel.XlSheetVisibility.xlSheetVeryHidden;

                    // Creamos los encabezados de nuestra "Base de Datos"
                    wsLog.Cells[1, 1] = "FECHA_HORA";
                    wsLog.Cells[1, 2] = "PREGUNTA";
                    wsLog.Cells[1, 3] = "TIPO_VALIDACION";
                    wsLog.Cells[1, 4] = "RANGO_AFECTADO";
                    wsLog.Cells[1, 5] = "USUARIO_RED";
                }

                // 4. Inteligencia Espacial: Encontramos la primera fila vacía
                Excel.Range ultimaCelda = wsLog.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                int ultimaFila = ultimaCelda.Row + 1;
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ultimaCelda);

                // 5. Escribimos la transacción (Registro de auditoría)
                wsLog.Cells[ultimaFila, 1] = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                wsLog.Cells[ultimaFila, 2] = pregunta;
                wsLog.Cells[ultimaFila, 3] = tipoValidacion;
                wsLog.Cells[ultimaFila, 4] = direccionRango;
                wsLog.Cells[ultimaFila, 5] = Environment.UserName;

                // =========================================================================
                // 6. RESTAURACIÓN DE ESTADO (Bypass del Robo de Foco)
                // =========================================================================
                if (hojaOriginal != null)
                {
                    // Devolvemos silenciosamente la pantalla a donde el censista estaba trabajando
                    hojaOriginal.Activate();
                }
            }
            catch (Exception ex)
            {
                // Falla silenciosa: Evitamos romper el flujo del usuario si falla la bitácora
                System.Diagnostics.Debug.WriteLine("Error en bitácora: " + ex.Message);
            }
            finally
            {
                // Limpieza estricta de memoria COM (Previene memory leaks y Excel fantasma)
                if (hojaOriginal != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(hojaOriginal);
                if (rangoInsertar != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rangoInsertar);
                if (wsLog != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(wsLog);
            }
        }

        // =========================================================================
        // 2. MOTOR DE LECTURA: Extrae el estado actual para mostrarlo en tu Interfaz
        // =========================================================================
        public static System.Data.DataTable ObtenerHistorialCenso(Excel.Workbook libro)
        {
            System.Data.DataTable tabla = new System.Data.DataTable();
            tabla.Columns.Add("Pregunta", typeof(string));
            tabla.Columns.Add("Validación Aplicada", typeof(string));
            tabla.Columns.Add("Fecha de Aplicación", typeof(string));

            Excel.Worksheet wsLog = null;
            try
            {
                foreach (Excel.Worksheet sheet in libro.Worksheets)
                {
                    if (sheet.Name == NOMBRE_HOJA_LOG)
                    {
                        wsLog = sheet;
                        break;
                    }
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet);
                }

                if (wsLog != null)
                {
                    // Extraemos todo el rango de datos a una matriz de C# (Velocidad O(1))
                    Excel.Range rangoUsado = wsLog.UsedRange;
                    object[,] valores = (object[,])rangoUsado.Value2;

                    if (valores != null && valores.GetLength(0) > 1)
                    {
                        for (int f = 2; f <= valores.GetLength(0); f++)
                        {
                            // -------------------------------------------------------------
                            // TRUCO DE INTEROP: DECODIFICACIÓN OADATE
                            // -------------------------------------------------------------
                            // Si Excel nos devuelve un Double (Serial de fecha), lo convertimos 
                            // nativamente a DateTime de C#. Si no, lo leemos como texto.
                            string fechaStr = "";
                            if (valores[f, 1] is double fechaSerial)
                            {
                                fechaStr = DateTime.FromOADate(fechaSerial).ToString("dd/MM/yyyy HH:mm:ss");
                            }
                            else
                            {
                                fechaStr = valores[f, 1]?.ToString() ?? "";
                            }

                            string preg = valores[f, 2]?.ToString() ?? "";
                            string val = valores[f, 3]?.ToString() ?? "";

                            tabla.Rows.Add(preg, val, fechaStr);
                        }
                    }
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(rangoUsado);
                }
            }
            finally
            {
                if (wsLog != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(wsLog);
            }
            return tabla;
        }
    }
}
