using System;
using Excel = Microsoft.Office.Interop.Excel;

public static class AuditoriaCenso
{
    private const string NOMBRE_HOJA_LOG = "SAVCNG_SysLog";

    // =========================================================================
    // 1. MOTOR DE ESCRITURA: Registra una nueva validación en el archivo
    // =========================================================================
    // NUEVO: Agregamos el parámetro 'funcionalidadExcel'
    public static void RegistrarAccion(Excel.Workbook libro, string pregunta, string tipoValidacion, string direccionRango, string funcionalidadExcel)
    {
        Excel.Worksheet wsLog = null;
        Excel.Worksheet hojaOriginal = null;

        try
        {
            // 1. CAPTURA DE ESTADO (Preservación del Foco)
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

            // 3. MIGRACIÓN O CREACIÓN DEL ESQUEMA
            if (wsLog == null)
            {
                // A) CREACIÓN NUEVA (Esquema v2)
                wsLog = (Excel.Worksheet)libro.Worksheets.Add(After: libro.Worksheets[libro.Worksheets.Count]);
                wsLog.Name = NOMBRE_HOJA_LOG;
                wsLog.Visible = Excel.XlSheetVisibility.xlSheetVeryHidden;

                wsLog.Cells[1, 1] = "FECHA_HORA";
                wsLog.Cells[1, 2] = "PREGUNTA";
                wsLog.Cells[1, 3] = "TIPO_VALIDACION";
                wsLog.Cells[1, 4] = "RANGO_AFECTADO";
                wsLog.Cells[1, 5] = "FUNCIONALIDAD_EXCEL"; // <--- NUEVA COLUMNA
                wsLog.Cells[1, 6] = "USUARIO_RED";         // <--- DESPLAZADO
            }
            else
            {
                // B) ACTUALIZACIÓN DE ARCHIVOS VIEJOS (Schema Migration)
                Excel.Range celdaHeader5 = (Excel.Range)wsLog.Cells[1, 5];
                string headerActual = celdaHeader5.Value2?.ToString() ?? "";

                if (headerActual == "USUARIO_RED")
                {
                    // Si el archivo es viejo, insertamos una columna para desplazar a USUARIO_RED a la col 6
                    Excel.Range columna5 = (Excel.Range)wsLog.Columns[5];
                    columna5.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, Type.Missing);
                    wsLog.Cells[1, 5] = "FUNCIONALIDAD_EXCEL";

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(columna5);
                }
                System.Runtime.InteropServices.Marshal.ReleaseComObject(celdaHeader5);
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
            wsLog.Cells[ultimaFila, 5] = funcionalidadExcel;   // <--- INYECCIÓN DEL NUEVO DATO
            wsLog.Cells[ultimaFila, 6] = Environment.UserName;

            // 6. RESTAURACIÓN DE ESTADO
            if (hojaOriginal != null)
            {
                hojaOriginal.Activate();
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine("Error en bitácora: " + ex.Message);
        }
        finally
        {
            if (hojaOriginal != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(hojaOriginal);
            if (wsLog != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(wsLog);
        }
    }

    // =========================================================================
    // 2. MOTOR DE LECTURA: Extrae el estado actual para mostrarlo en tu Interfaz
    // =========================================================================
    public static System.Data.DataTable ObtenerHistorialCenso(Excel.Workbook libro)
    {
        System.Data.DataTable tabla = new System.Data.DataTable();

        // 1. DEFINICIÓN DEL ESQUEMA COMPLETO (6 Columnas Visibles)
        tabla.Columns.Add("Fecha/Hora", typeof(string));
        tabla.Columns.Add("Pregunta", typeof(string));
        tabla.Columns.Add("Validación Aplicada", typeof(string));
        tabla.Columns.Add("Rango Afectado", typeof(string));
        tabla.Columns.Add("Mecanismo Nativo", typeof(string));
        tabla.Columns.Add("Usuario", typeof(string));

        Excel.Worksheet wsLog = null;
        try
        {
            // 2. RASTREO DE LA BITÁCORA
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
                // 3. EXTRACCIÓN MASIVA A RAM (O(1) en Velocidad)
                Excel.Range rangoUsado = wsLog.UsedRange;
                object[,] valores = (object[,])rangoUsado.Value2;

                if (valores != null && valores.GetLength(0) > 1)
                {
                    // Obtenemos el total de columnas reales detectadas por Excel
                    int totalColumnasExcel = valores.GetLength(1);

                    for (int f = 2; f <= valores.GetLength(0); f++)
                    {
                        // 4. MAPEO Y PROTECCIÓN DE DATOS (Bounds Checking)

                        // Col 1: Fecha (Con decodificación OADate de Interop)
                        string fechaStr = "";
                        if (valores[f, 1] is double fechaSerial)
                        {
                            fechaStr = DateTime.FromOADate(fechaSerial).ToString("dd/MM/yyyy HH:mm:ss");
                        }
                        else
                        {
                            fechaStr = valores[f, 1]?.ToString() ?? "";
                        }

                        // Col 2: Pregunta
                        string preg = totalColumnasExcel >= 2 ? (valores[f, 2]?.ToString() ?? "") : "";

                        // Col 3: Tipo de Validación
                        string val = totalColumnasExcel >= 3 ? (valores[f, 3]?.ToString() ?? "") : "";

                        // Col 4: Rango Físico
                        string rango = totalColumnasExcel >= 4 ? (valores[f, 4]?.ToString() ?? "") : "";

                        // Col 5: Funcionalidad Excel (Con fallback por si es un archivo viejo)
                        string mecanismo = totalColumnasExcel >= 5 ? (valores[f, 5]?.ToString() ?? "N/A") : "N/A";

                        // Col 6: Usuario de Red
                        string usuario = totalColumnasExcel >= 6 ? (valores[f, 6]?.ToString() ?? "") : "";

                        // 5. INSERCIÓN EN EL DATATABLE
                        // El orden aquí debe coincidir exactamente con la definición del esquema superior
                        tabla.Rows.Add(fechaStr, preg, val, rango, mecanismo, usuario);
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