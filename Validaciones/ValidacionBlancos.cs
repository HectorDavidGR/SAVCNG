using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;
using SAVCNG_ExcelDNA.Core;

namespace SAVCNG_ExcelDNA.Validaciones
{
    public class ValidacionBlancos : IValidacionExcel
    {
        public ResultadoValidacion Ejecutar(Excel.Application excelApp, Excel.Workbook libroCenso, Excel.Range rangoCapturado)
        {
            try
            {
                // =========================================================================
                // 1. VALIDACIÓN Y VERIFICACIÓN DE MULTI-RANGOS EN TABLA DE BLANCOS
                // =========================================================================
                int numAreasBlancos = rangoCapturado.Areas.Count;
                int filasPrimerRangoBlancos = 0;
                int filaInicial = 0;

                Excel.Range primeraAreaBlancos = null;
                try
                {
                    primeraAreaBlancos = (Excel.Range)rangoCapturado.Areas[1];
                    filasPrimerRangoBlancos = primeraAreaBlancos.Rows.Count;
                    filaInicial = primeraAreaBlancos.Row;
                }
                finally
                {
                    ExcelHelper.LiberarCom(primeraAreaBlancos);
                }

                for (int i = 2; i <= numAreasBlancos; i++)
                {
                    Excel.Range areaBlancos = null;
                    try
                    {
                        areaBlancos = (Excel.Range)rangoCapturado.Areas[i];
                        if (areaBlancos.Rows.Count != filasPrimerRangoBlancos)
                        {
                            return new ResultadoValidacion
                            {
                                Exito = false,
                                Mensaje = $"Error de coincidencia en Blancos: El 1er rango seleccionado tiene {filasPrimerRangoBlancos} fila(s), pero el rango #{i} tiene {areaBlancos.Rows.Count} fila(s).",
                                AlertaInyectada = false
                            };
                        }
                    }
                    finally
                    {
                        ExcelHelper.LiberarCom(areaBlancos);
                    }
                }

                // =========================================================================
                // 2. SELECTOR DE MODALIDAD DE PROCESO
                // =========================================================================
                object resOpcion = excelApp.InputBox(
                    "Selecciona el número del proceso de Blancos que deseas ejecutar:\n\n" +
                    "1 = Validación de Blancos Tipo 1 2 9\n" +
                    "2 = Validación de Blancos Sin Disparador",
                    "Selección de Modalidad - Validaciones de Blancos",
                    1, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 1);

                if (resOpcion is bool && (bool)resOpcion == false)
                {
                    return new ResultadoValidacion { Exito = false, Mensaje = "Operación cancelada por el usuario.", AlertaInyectada = false };
                }

                int opcion = 0;
                if (!int.TryParse(resOpcion.ToString(), out opcion) || (opcion != 1 && opcion != 2))
                {
                    return new ResultadoValidacion { Exito = false, Mensaje = "Opción no válida. Debes ingresar 1 o 2.", AlertaInyectada = false };
                }

                // =========================================================================
                // 3. ENRUTAMIENTO DINÁMICO
                // =========================================================================
                if (opcion == 1)
                {
                    if (filasPrimerRangoBlancos == 1)
                        return EjecutarBlancosMonofilaConDisparador(excelApp, libroCenso, rangoCapturado);
                    else
                        return EjecutarBlancosMultifilaConDisparador(excelApp, libroCenso, rangoCapturado, filasPrimerRangoBlancos, numAreasBlancos, filaInicial);
                }
                else
                {
                    if (filasPrimerRangoBlancos == 1)
                        return EjecutarBlancosMonofilaSinDisparador(excelApp, libroCenso, rangoCapturado);
                    else
                        return EjecutarBlancosMultifilaSinDisparador(excelApp, libroCenso, rangoCapturado, filasPrimerRangoBlancos, numAreasBlancos, filaInicial);
                }
            }
            catch (Exception ex)
            {
                return new ResultadoValidacion { Exito = false, Mensaje = "Error crítico en el selector de blancos: " + ex.Message, AlertaInyectada = false };
            }
        }

        #region PROCESO 1: BLANCOS TIPO 1 2 9 (MONOFILA)
        private ResultadoValidacion EjecutarBlancosMonofilaConDisparador(Excel.Application excelApp, Excel.Workbook libroCenso, Excel.Range rangoCapturado)
        {
            Excel.Worksheet wsActual = null;
            Excel.Range rangoDisparador = null;
            Excel.Range rangoDestino = null;

            try
            {
                wsActual = (Excel.Worksheet)rangoCapturado.Worksheet;

                object resInclusiones = excelApp.InputBox("Indica los valores de la CELDA DISPARADORA Tal que SÍ se aplica la validación de Blancos.\n\n• Ejemplo: 1, 5", "1. Inclusiones", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 2);
                if (resInclusiones is bool && (bool)resInclusiones == false) return new ResultadoValidacion { Exito = false, Mensaje = "Cancelado.", AlertaInyectada = false };
                string textoInclusiones = resInclusiones != null ? resInclusiones.ToString().Trim() : "";

                object resDisparador = excelApp.InputBox("Selecciona el rango de la CELDA DISPARADORA:", "2. Celda Disparadora", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8);
                if (resDisparador is bool && (bool)resDisparador == false) return new ResultadoValidacion { Exito = false, Mensaje = "Cancelado.", AlertaInyectada = false };
                rangoDisparador = (Excel.Range)resDisparador;

                object resDestino = excelApp.InputBox("Selecciona el rango para el mensaje de alerta:", "3. Ubicación Alerta", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8);
                if (resDestino is bool && (bool)resDestino == false) return new ResultadoValidacion { Exito = false, Mensaje = "Cancelado.", AlertaInyectada = false };
                rangoDestino = (Excel.Range)resDestino;

                object resTexto = excelApp.InputBox("Mensaje de advertencia:", "4. Mensaje", "Favor de ingresar toda la información requerida en la pregunta", Type.Missing, Type.Missing, Type.Missing, Type.Missing, 2);
                if (resTexto is bool && (bool)resTexto == false) return new ResultadoValidacion { Exito = false, Mensaje = "Cancelado.", AlertaInyectada = false };
                string textoAlerta = resTexto != null ? resTexto.ToString().Trim() : "Favor de ingresar toda la información requerida en la pregunta";

                excelApp.ScreenUpdating = false;
                int filaInicial = rangoCapturado.Row;
                string dirDisparador = rangoDisparador.Address;
                string clausulaCondicion = "TRUE";

                if (!string.IsNullOrEmpty(textoInclusiones))
                {
                    List<string> frag = new List<string> { $"COUNTIF({dirDisparador},\"\")" };
                    string[] vals = textoInclusiones.Split(new char[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries);
                    foreach (var v in vals)
                    {
                        string valL = v.Trim();
                        frag.Add($"COUNTIF({dirDisparador},{(double.TryParse(valL, out _) ? valL : $"\"{valL}\"")})");
                    }
                    clausulaCondicion = $"({string.Join("+", frag)})>0";
                }

                List<string> listaLlenas = new List<string>();
                List<string> listaVacias = new List<string>();

                for (int a = 1; a <= rangoCapturado.Areas.Count; a++)
                {
                    Excel.Range area = null;
                    try
                    {
                        area = (Excel.Range)rangoCapturado.Areas[a];
                        int cols = area.Columns.Count;
                        int c = 1;
                        while (c <= cols)
                        {
                            Excel.Range celda = null;
                            Excel.Range areaComb = null;
                            try
                            {
                                celda = (Excel.Range)area.Cells[1, c];
                                string dir = $"{celda.Address.Split('$')[1]}{filaInicial}";
                                listaLlenas.Add($"({dir}<>\"\")");
                                listaVacias.Add($"({dir}=\"\")");

                                if (celda.MergeCells is bool && (bool)celda.MergeCells)
                                {
                                    areaComb = celda.MergeArea;
                                    c += areaComb.Columns.Count;
                                }
                                else { c++; }
                            }
                            finally
                            {
                                ExcelHelper.LiberarCom(areaComb);
                                ExcelHelper.LiberarCom(celda);
                            }
                        }
                    }
                    finally { ExcelHelper.LiberarCom(area); }
                }

                string formula = $"=IF(AND(({string.Join("+", listaLlenas)})>0, ({string.Join("+", listaVacias)})>0, {clausulaCondicion}), \"{textoAlerta}\", \"\")";

                if (rangoDestino.Count > 1 && !(rangoDestino.MergeCells is bool && (bool)rangoDestino.MergeCells)) rangoDestino.Merge();
                rangoDestino.Font.Name = "Arial";
                rangoDestino.Font.Size = 10;
                rangoDestino.Font.Bold = true;
                rangoDestino.Font.Color = ColorTranslator.ToOle(Color.FromArgb(0, 112, 192));
                rangoDestino.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                rangoDestino.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                rangoDestino.FormulaLocal = ExcelHelper.TraducirFormulaLocal(wsActual, formula, filaInicial);

                string numeroPregunta = ExcelHelper.ObtenerNumeroPregunta(rangoCapturado);
                AuditoriaCenso.RegistrarAccion(libroCenso, numeroPregunta, "Validación Blancos Tipo 1 2 9 (Monofila)", rangoCapturado.Address.Replace("$", ""), "Alerta Visual");
                return new ResultadoValidacion { Exito = true, Mensaje = "Validación monofila aplicada.", AlertaInyectada = true };
            }
            catch (Exception ex)
            {
                return new ResultadoValidacion { Exito = false, Mensaje = "Error: " + ex.Message, AlertaInyectada = false };
            }
            finally
            {
                ExcelHelper.LiberarCom(rangoDestino);
                ExcelHelper.LiberarCom(rangoDisparador);
                ExcelHelper.LiberarCom(wsActual);
                if (excelApp != null) excelApp.ScreenUpdating = true;
            }
        }
        #endregion

        #region PROCESO 1: BLANCOS TIPO 1 2 9 (MULTIFILA)
        private ResultadoValidacion EjecutarBlancosMultifilaConDisparador(Excel.Application excelApp, Excel.Workbook libroCenso, Excel.Range rangoCapturado, int filasTotales, int numAreas, int filaInicial)
        {
            Excel.Worksheet wsActual = null;
            Excel.Range rangoDisparador = null;
            Excel.Range rangoNumerales = null;
            Excel.Range rangoAux = null;
            Excel.Range rangoAlerta = null;

            try
            {
                wsActual = (Excel.Worksheet)rangoCapturado.Worksheet;

                object resInclusiones = excelApp.InputBox("Indica los valores de la CELDA DISPARADORA Tal que SÍ se aplica la validación de Blancos.\n\n• Ejemplo: 1, 5", "1. Inclusiones", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 2);
                if (resInclusiones is bool && (bool)resInclusiones == false) return new ResultadoValidacion { Exito = false, Mensaje = "Cancelado.", AlertaInyectada = false };
                string txtInclusiones = resInclusiones != null ? resInclusiones.ToString().Trim() : "";

                object resDisparador = excelApp.InputBox("Selecciona los rangos de las CELDAS DISPARADORAS:", "2. Disparador", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8);
                if (resDisparador is bool && (bool)resDisparador == false) return new ResultadoValidacion { Exito = false, Mensaje = "Cancelado.", AlertaInyectada = false };
                rangoDisparador = (Excel.Range)resDisparador;

                // =========================================================================
                // NUEVA REGLA: Validación de Filas (Celda Disparadora)
                // =========================================================================
                int filasDisparador = 0;
                int countAreasDisparador = rangoDisparador.Areas.Count;
                for (int i = 1; i <= countAreasDisparador; i++)
                {
                    Excel.Range areaDisp = null;
                    try
                    {
                        areaDisp = (Excel.Range)rangoDisparador.Areas[i];
                        filasDisparador += areaDisp.Rows.Count;
                    }
                    finally { ExcelHelper.LiberarCom(areaDisp); }
                }

                if (filasDisparador != filasTotales)
                {
                    return new ResultadoValidacion { Exito = false, Mensaje = $"Proceso cancelado: La CELDA DISPARADORA seleccionada tiene {filasDisparador} fila(s), pero el primer rango de Blancos tiene {filasTotales} fila(s).", AlertaInyectada = false };
                }

                object resNumerales = excelApp.InputBox("Selecciona los rangos de los NUMERALES:", "3. Numerales", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8);
                if (resNumerales is bool && (bool)resNumerales == false) return new ResultadoValidacion { Exito = false, Mensaje = "Cancelado.", AlertaInyectada = false };
                rangoNumerales = (Excel.Range)resNumerales;

                // =========================================================================
                // NUEVA REGLA: Validación de Filas (Numerales)
                // =========================================================================
                int filasNumerales = 0;
                int countAreasNumerales = rangoNumerales.Areas.Count;
                for (int i = 1; i <= countAreasNumerales; i++)
                {
                    Excel.Range areaNum = null;
                    try
                    {
                        areaNum = (Excel.Range)rangoNumerales.Areas[i];
                        filasNumerales += areaNum.Rows.Count;
                    }
                    finally { ExcelHelper.LiberarCom(areaNum); }
                }

                if (filasNumerales != filasTotales)
                {
                    return new ResultadoValidacion { Exito = false, Mensaje = $"Proceso cancelado: El rango de NUMERALES tiene {filasNumerales} fila(s), pero el primer rango de Blancos tiene {filasTotales} fila(s).", AlertaInyectada = false };
                }

                object resAux = excelApp.InputBox("Selecciona la celda para las columnas AUXILIARES:", "4. Auxiliares", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8);
                if (resAux is bool && (bool)resAux == false) return new ResultadoValidacion { Exito = false, Mensaje = "Cancelado.", AlertaInyectada = false };
                rangoAux = (Excel.Range)resAux;

                object resAlerta = excelApp.InputBox("Selecciona el rango para el mensaje de ALERTA:", "5. Alerta", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8);
                if (resAlerta is bool && (bool)resAlerta == false) return new ResultadoValidacion { Exito = false, Mensaje = "Cancelado.", AlertaInyectada = false };
                rangoAlerta = (Excel.Range)resAlerta;

                excelApp.ScreenUpdating = false;

                // 1. Condición Tabla Activa
                List<string> fragActiva = new List<string>();
                for (int a = 1; a <= rangoDisparador.Areas.Count; a++)
                {
                    Excel.Range area = null;
                    try { area = (Excel.Range)rangoDisparador.Areas[a]; fragActiva.Add($"COUNTA({area.Address})"); }
                    finally { ExcelHelper.LiberarCom(area); }
                }
                for (int a = 1; a <= numAreas; a++)
                {
                    Excel.Range area = null;
                    try { area = (Excel.Range)rangoCapturado.Areas[a]; fragActiva.Add($"COUNTA({area.Address})"); }
                    finally { ExcelHelper.LiberarCom(area); }
                }
                string condActiva = $"({string.Join("+", fragActiva)})>0";

                // 2. Extraer Direcciones Disparador
                List<string> dirsDisp = new List<string>();
                for (int a = 1; a <= rangoDisparador.Areas.Count; a++)
                {
                    Excel.Range area = null;
                    try
                    {
                        area = (Excel.Range)rangoDisparador.Areas[a];
                        for (int r = 1; r <= area.Rows.Count; r++)
                        {
                            Excel.Range celda = null;
                            try { celda = (Excel.Range)area.Cells[r, 1]; dirsDisp.Add(celda.Address); }
                            finally { ExcelHelper.LiberarCom(celda); }
                        }
                    }
                    finally { ExcelHelper.LiberarCom(area); }
                }

                // 3. Extraer Columnas Blancos
                List<int> iniciosAreas = new List<int>();
                List<List<string>> colsPorArea = new List<List<string>>();
                for (int a = 1; a <= numAreas; a++)
                {
                    Excel.Range area = null;
                    try
                    {
                        area = (Excel.Range)rangoCapturado.Areas[a];
                        iniciosAreas.Add(area.Row);
                        List<string> letras = new List<string>();
                        int c = 1;
                        while (c <= area.Columns.Count)
                        {
                            Excel.Range celda = null, areaComb = null;
                            try
                            {
                                celda = (Excel.Range)area.Cells[1, c];
                                letras.Add(celda.Address.Split('$')[1]);
                                if (celda.MergeCells is bool && (bool)celda.MergeCells) { areaComb = celda.MergeArea; c += areaComb.Columns.Count; }
                                else c++;
                            }
                            finally { ExcelHelper.LiberarCom(areaComb); ExcelHelper.LiberarCom(celda); }
                        }
                        colsPorArea.Add(letras);
                    }
                    finally { ExcelHelper.LiberarCom(area); }
                }

                // 4. Extraer Numerales
                List<string> nums = new List<string>();
                for (int a = 1; a <= rangoNumerales.Areas.Count; a++)
                {
                    Excel.Range area = null;
                    try
                    {
                        area = (Excel.Range)rangoNumerales.Areas[a];
                        for (int r = 1; r <= area.Rows.Count; r++)
                        {
                            Excel.Range celda = null;
                            try
                            {
                                celda = (Excel.Range)area.Cells[r, 1];
                                string val = celda.Value2?.ToString().Trim() ?? "";
                                val = Regex.Replace(val, @"[^\w]+$", "");
                                nums.Add(string.IsNullOrEmpty(val) ? (nums.Count + 1).ToString() : val);
                            }
                            finally { ExcelHelper.LiberarCom(celda); }
                        }
                    }
                    finally { ExcelHelper.LiberarCom(area); }
                }

                // 5. INYECCIÓN DEL ENCABEZADO Y PRESENTACIÓN VISUAL
                int c1 = rangoAux.Column;
                int c2 = c1 + 1;

                if (filaInicial > 1)
                {
                    Excel.Range sub1 = null;
                    Excel.Range sub2 = null;
                    Excel.Range mainCell1 = null;
                    Excel.Range mainCell2 = null;
                    Excel.Range mainRng = null;

                    try
                    {
                        sub1 = (Excel.Range)wsActual.Cells[filaInicial - 1, c1];
                        sub1.Value2 = "Bandera";
                        sub1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        sub1.Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml("#00CC99"));
                        sub1.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                        sub2 = (Excel.Range)wsActual.Cells[filaInicial - 1, c2];
                        sub2.Value2 = "numeral(es)";
                        sub2.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        sub2.Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml("#33CCCC"));
                        sub2.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                        if (filaInicial > 2)
                        {
                            mainCell1 = (Excel.Range)wsActual.Cells[filaInicial - 2, c1];
                            mainCell2 = (Excel.Range)wsActual.Cells[filaInicial - 2, c2];
                            mainRng = wsActual.Range[mainCell1, mainCell2];

                            mainRng.Merge();
                            mainRng.Value2 = "SAV_MENSAJE BLANCOS";
                            mainRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                            mainRng.Font.Bold = true;
                            mainRng.Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml("#FFD966"));
                            mainRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        }
                    }
                    finally
                    {
                        ExcelHelper.LiberarCom(sub1);
                        ExcelHelper.LiberarCom(sub2);
                        ExcelHelper.LiberarCom(mainCell1);
                        ExcelHelper.LiberarCom(mainCell2);
                        ExcelHelper.LiberarCom(mainRng);
                    }
                }

                Excel.Range aux1Init = (Excel.Range)wsActual.Cells[filaInicial, c1];
                Excel.Range aux2Init = (Excel.Range)wsActual.Cells[filaInicial, c2];
                string letraC1 = aux1Init.Address.Split('$')[1];
                string letraC2 = aux2Init.Address.Split('$')[1];
                ExcelHelper.LiberarCom(aux1Init); ExcelHelper.LiberarCom(aux2Init);

                for (int r = 1; r <= filasTotales; r++)
                {
                    int filaActual = filaInicial + r - 1;
                    string dirDisp = dirsDisp[r - 1];

                    string condFila = "TRUE";
                    if (!string.IsNullOrEmpty(txtInclusiones))
                    {
                        List<string> orFrag = new List<string> { $"{dirDisp}=\"\"" };
                        foreach (string v in txtInclusiones.Split(new char[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries))
                        {
                            string val = v.Trim();
                            orFrag.Add($"{dirDisp}={(double.TryParse(val, out _) ? val : $"\"{val}\"")}");
                        }
                        condFila = $"OR({string.Join(",", orFrag)})";
                    }

                    List<string> vacias = new List<string>();
                    for (int a = 0; a < numAreas; a++)
                    {
                        int fArea = iniciosAreas[a] + r - 1;
                        foreach (string l in colsPorArea[a]) vacias.Add($"({l}{fArea}=\"\")");
                    }

                    string form1 = $"=IF(AND({condActiva}, {condFila}, ({string.Join("+", vacias)})>0), \"{nums[r - 1]}\", \"\")";
                    Excel.Range celda1 = null;
                    try
                    {
                        celda1 = (Excel.Range)wsActual.Cells[filaActual, c1];
                        celda1.FormulaLocal = ExcelHelper.TraducirFormulaLocal(wsActual, form1, filaActual);
                        celda1.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    }
                    finally { ExcelHelper.LiberarCom(celda1); }

                    string dir1 = $"{letraC1}{filaActual}";
                    string form2 = r == 1 ? $"={dir1}" : $"=IF({dir1}=\"\", {letraC2}{filaActual - 1}, IF({letraC2}{filaActual - 1}=\"\", {dir1}, {letraC2}{filaActual - 1} & \", \" & {dir1}))";
                    Excel.Range celda2 = null;
                    try
                    {
                        celda2 = (Excel.Range)wsActual.Cells[filaActual, c2];
                        celda2.FormulaLocal = ExcelHelper.TraducirFormulaLocal(wsActual, form2, filaActual);
                        celda2.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    }
                    finally { ExcelHelper.LiberarCom(celda2); }
                }

                // 6. INYECCIÓN ALERTA FINAL
                string ultimaCelda = $"{letraC2}{filaInicial + filasTotales - 1}";
                string formAlerta = $"=IF({ultimaCelda}=\"\", \"\", IF(ISNUMBER(SEARCH(\",\", {ultimaCelda})), \"Favor de revisar la información capturada en los numerales: \" & SUBSTITUTE({ultimaCelda}, \", \", \" y \", LEN({ultimaCelda})-LEN(SUBSTITUTE({ultimaCelda}, \",\", \"\"))), \"Favor de revisar la información capturada en el numeral: \" & {ultimaCelda}))";

                if (rangoAlerta.Count > 1 && !(rangoAlerta.MergeCells is bool && (bool)rangoAlerta.MergeCells)) rangoAlerta.Merge();
                rangoAlerta.Font.Name = "Arial";
                rangoAlerta.Font.Size = 10;
                rangoAlerta.Font.Bold = true;
                rangoAlerta.Font.Color = ColorTranslator.ToOle(Color.FromArgb(0, 112, 192));
                rangoAlerta.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                rangoAlerta.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                rangoAlerta.FormulaLocal = ExcelHelper.TraducirFormulaLocal(wsActual, formAlerta, filaInicial);

                string numeroPregunta = ExcelHelper.ObtenerNumeroPregunta(rangoCapturado);
                AuditoriaCenso.RegistrarAccion(libroCenso, numeroPregunta, "Validación Blancos Tipo 1 2 9 (Multifila)", rangoCapturado.Address.Replace("$", ""), "Auxiliares + Alerta");
                return new ResultadoValidacion { Exito = true, Mensaje = "Proceso multifila con disparador aplicado.", AlertaInyectada = true };
            }
            catch (Exception ex)
            {
                return new ResultadoValidacion { Exito = false, Mensaje = "Error: " + ex.Message, AlertaInyectada = false };
            }
            finally
            {
                ExcelHelper.LiberarCom(rangoAlerta); ExcelHelper.LiberarCom(rangoAux);
                ExcelHelper.LiberarCom(rangoNumerales); ExcelHelper.LiberarCom(rangoDisparador);
                ExcelHelper.LiberarCom(wsActual); if (excelApp != null) excelApp.ScreenUpdating = true;
            }
        }
        #endregion

        #region PROCESO 2: BLANCOS SIN CÓDIGO (MONOFILA)
        private ResultadoValidacion EjecutarBlancosMonofilaSinDisparador(Excel.Application excelApp, Excel.Workbook libroCenso, Excel.Range rangoCapturado)
        {
            Excel.Worksheet wsActual = null;
            Excel.Range rangoDestino = null;

            try
            {
                wsActual = (Excel.Worksheet)rangoCapturado.Worksheet;

                object resDestino = excelApp.InputBox("Selecciona el rango del mensaje de ALERTA:", "1. Alerta", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8);
                if (resDestino is bool && (bool)resDestino == false) return new ResultadoValidacion { Exito = false, Mensaje = "Cancelado.", AlertaInyectada = false };
                rangoDestino = (Excel.Range)resDestino;

                object resTexto = excelApp.InputBox("Escribe el mensaje:", "2. Mensaje", "Favor de ingresar toda la información requerida en la pregunta", Type.Missing, Type.Missing, Type.Missing, Type.Missing, 2);
                if (resTexto is bool && (bool)resTexto == false) return new ResultadoValidacion { Exito = false, Mensaje = "Cancelado.", AlertaInyectada = false };
                string textoAlerta = resTexto != null ? resTexto.ToString().Trim() : "Favor de ingresar toda la información requerida en la pregunta";

                excelApp.ScreenUpdating = false;
                int filaInicial = rangoCapturado.Row;

                List<string> llenas = new List<string>();
                List<string> vacias = new List<string>();

                for (int a = 1; a <= rangoCapturado.Areas.Count; a++)
                {
                    Excel.Range area = null;
                    try
                    {
                        area = (Excel.Range)rangoCapturado.Areas[a];
                        int c = 1;
                        while (c <= area.Columns.Count)
                        {
                            Excel.Range celda = null, areaComb = null;
                            try
                            {
                                celda = (Excel.Range)area.Cells[1, c];
                                string dir = $"{celda.Address.Split('$')[1]}{filaInicial}";
                                llenas.Add($"({dir}<>\"\")");
                                vacias.Add($"({dir}=\"\")");

                                if (celda.MergeCells is bool && (bool)celda.MergeCells) { areaComb = celda.MergeArea; c += areaComb.Columns.Count; }
                                else c++;
                            }
                            finally { ExcelHelper.LiberarCom(areaComb); ExcelHelper.LiberarCom(celda); }
                        }
                    }
                    finally { ExcelHelper.LiberarCom(area); }
                }

                string formula = $"=IF(AND(({string.Join("+", llenas)})>0, ({string.Join("+", vacias)})>0), \"{textoAlerta}\", \"\")";

                if (rangoDestino.Count > 1 && !(rangoDestino.MergeCells is bool && (bool)rangoDestino.MergeCells)) rangoDestino.Merge();
                rangoDestino.Font.Name = "Arial";
                rangoDestino.Font.Size = 10;
                rangoDestino.Font.Bold = true;
                rangoDestino.Font.Color = ColorTranslator.ToOle(Color.FromArgb(0, 112, 192));
                rangoDestino.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                rangoDestino.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                rangoDestino.FormulaLocal = ExcelHelper.TraducirFormulaLocal(wsActual, formula, filaInicial);

                string numeroPregunta = ExcelHelper.ObtenerNumeroPregunta(rangoCapturado);
                AuditoriaCenso.RegistrarAccion(libroCenso, numeroPregunta, "Validación Blancos Sin Código (Monofila)", rangoCapturado.Address.Replace("$", ""), "Alerta Visual");
                return new ResultadoValidacion { Exito = true, Mensaje = "Validación monofila aplicada.", AlertaInyectada = true };
            }
            catch (Exception ex)
            {
                return new ResultadoValidacion { Exito = false, Mensaje = "Error: " + ex.Message, AlertaInyectada = false };
            }
            finally
            {
                ExcelHelper.LiberarCom(rangoDestino);
                ExcelHelper.LiberarCom(wsActual);
                if (excelApp != null) excelApp.ScreenUpdating = true;
            }
        }
        #endregion

        #region PROCESO 2: BLANCOS SIN CÓDIGO (MULTIFILA)
        private ResultadoValidacion EjecutarBlancosMultifilaSinDisparador(Excel.Application excelApp, Excel.Workbook libroCenso, Excel.Range rangoCapturado, int filasTotales, int numAreas, int filaInicial)
        {
            Excel.Worksheet wsActual = null;
            Excel.Range rangoNumerales = null;
            Excel.Range rangoAux = null;
            Excel.Range rangoAlerta = null;

            try
            {
                wsActual = (Excel.Worksheet)rangoCapturado.Worksheet;

                object resNumerales = excelApp.InputBox("Selecciona los rangos de los NUMERALES:", "1. Numerales", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8);
                if (resNumerales is bool && (bool)resNumerales == false) return new ResultadoValidacion { Exito = false, Mensaje = "Cancelado.", AlertaInyectada = false };
                rangoNumerales = (Excel.Range)resNumerales;

                // =========================================================================
                // NUEVA REGLA: Validación de Filas (Numerales)
                // =========================================================================
                int filasNumerales = 0;
                int countAreasNumerales = rangoNumerales.Areas.Count;
                for (int i = 1; i <= countAreasNumerales; i++)
                {
                    Excel.Range areaNum = null;
                    try
                    {
                        areaNum = (Excel.Range)rangoNumerales.Areas[i];
                        filasNumerales += areaNum.Rows.Count;
                    }
                    finally { ExcelHelper.LiberarCom(areaNum); }
                }

                if (filasNumerales != filasTotales)
                {
                    return new ResultadoValidacion { Exito = false, Mensaje = $"Proceso cancelado: El rango de NUMERALES tiene {filasNumerales} fila(s), pero el primer rango de Blancos tiene {filasTotales} fila(s).", AlertaInyectada = false };
                }

                object resAux = excelApp.InputBox("Selecciona la celda para las columnas AUXILIARES:", "2. Auxiliares", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8);
                if (resAux is bool && (bool)resAux == false) return new ResultadoValidacion { Exito = false, Mensaje = "Cancelado.", AlertaInyectada = false };
                rangoAux = (Excel.Range)resAux;

                object resAlerta = excelApp.InputBox("Selecciona la celda para el mensaje de ALERTA:", "3. Alerta", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8);
                if (resAlerta is bool && (bool)resAlerta == false) return new ResultadoValidacion { Exito = false, Mensaje = "Cancelado.", AlertaInyectada = false };
                rangoAlerta = (Excel.Range)resAlerta;

                excelApp.ScreenUpdating = false;

                // 1. Extraer Columnas Blancos
                List<int> iniciosAreas = new List<int>();
                List<List<string>> colsPorArea = new List<List<string>>();
                for (int a = 1; a <= numAreas; a++)
                {
                    Excel.Range area = null;
                    try
                    {
                        area = (Excel.Range)rangoCapturado.Areas[a];
                        iniciosAreas.Add(area.Row);
                        List<string> letras = new List<string>();
                        int c = 1;
                        while (c <= area.Columns.Count)
                        {
                            Excel.Range celda = null, areaComb = null;
                            try
                            {
                                celda = (Excel.Range)area.Cells[1, c];
                                letras.Add(celda.Address.Split('$')[1]);
                                if (celda.MergeCells is bool && (bool)celda.MergeCells) { areaComb = celda.MergeArea; c += areaComb.Columns.Count; }
                                else c++;
                            }
                            finally { ExcelHelper.LiberarCom(areaComb); ExcelHelper.LiberarCom(celda); }
                        }
                        colsPorArea.Add(letras);
                    }
                    finally { ExcelHelper.LiberarCom(area); }
                }

                // 2. Extraer Numerales
                List<string> nums = new List<string>();
                for (int a = 1; a <= rangoNumerales.Areas.Count; a++)
                {
                    Excel.Range area = null;
                    try
                    {
                        area = (Excel.Range)rangoNumerales.Areas[a];
                        for (int r = 1; r <= area.Rows.Count; r++)
                        {
                            Excel.Range celda = null;
                            try
                            {
                                celda = (Excel.Range)area.Cells[r, 1];
                                string val = celda.Value2?.ToString().Trim() ?? "";
                                val = Regex.Replace(val, @"[^\w]+$", "");
                                nums.Add(string.IsNullOrEmpty(val) ? (nums.Count + 1).ToString() : val);
                            }
                            finally { ExcelHelper.LiberarCom(celda); }
                        }
                    }
                    finally { ExcelHelper.LiberarCom(area); }
                }

                // 3. INYECCIÓN DEL ENCABEZADO Y PRESENTACIÓN VISUAL
                int c1 = rangoAux.Column;
                int c2 = c1 + 1;

                if (filaInicial > 1)
                {
                    Excel.Range sub1 = null;
                    Excel.Range sub2 = null;
                    Excel.Range mainCell1 = null;
                    Excel.Range mainCell2 = null;
                    Excel.Range mainRng = null;

                    try
                    {
                        sub1 = (Excel.Range)wsActual.Cells[filaInicial - 1, c1];
                        sub1.Value2 = "Bandera";
                        sub1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        sub1.Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml("#00CC99"));
                        sub1.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                        sub2 = (Excel.Range)wsActual.Cells[filaInicial - 1, c2];
                        sub2.Value2 = "numeral(es)";
                        sub2.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        sub2.Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml("#33CCCC"));
                        sub2.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                        if (filaInicial > 2)
                        {
                            mainCell1 = (Excel.Range)wsActual.Cells[filaInicial - 2, c1];
                            mainCell2 = (Excel.Range)wsActual.Cells[filaInicial - 2, c2];
                            mainRng = wsActual.Range[mainCell1, mainCell2];

                            mainRng.Merge();
                            mainRng.Value2 = "SAV_MENSAJE BLANCOS";
                            mainRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                            mainRng.Font.Bold = true;
                            mainRng.Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml("#FFD966"));
                            mainRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        }
                    }
                    finally
                    {
                        ExcelHelper.LiberarCom(sub1);
                        ExcelHelper.LiberarCom(sub2);
                        ExcelHelper.LiberarCom(mainCell1);
                        ExcelHelper.LiberarCom(mainCell2);
                        ExcelHelper.LiberarCom(mainRng);
                    }
                }

                Excel.Range aux1Init = (Excel.Range)wsActual.Cells[filaInicial, c1];
                Excel.Range aux2Init = (Excel.Range)wsActual.Cells[filaInicial, c2];
                string letraC1 = aux1Init.Address.Split('$')[1];
                string letraC2 = aux2Init.Address.Split('$')[1];
                ExcelHelper.LiberarCom(aux1Init); ExcelHelper.LiberarCom(aux2Init);

                for (int r = 1; r <= filasTotales; r++)
                {
                    int filaActual = filaInicial + r - 1;
                    List<string> llenas = new List<string>();
                    List<string> vacias = new List<string>();

                    for (int a = 0; a < numAreas; a++)
                    {
                        int fArea = iniciosAreas[a] + r - 1;
                        foreach (string l in colsPorArea[a])
                        {
                            llenas.Add($"({l}{fArea}<>\"\")");
                            vacias.Add($"({l}{fArea}=\"\")");
                        }
                    }

                    string form1 = $"=IF(AND(({string.Join("+", llenas)})>0, ({string.Join("+", vacias)})>0), \"{nums[r - 1]}\", \"\")";
                    Excel.Range celda1 = null;
                    try
                    {
                        celda1 = (Excel.Range)wsActual.Cells[filaActual, c1];
                        celda1.FormulaLocal = ExcelHelper.TraducirFormulaLocal(wsActual, form1, filaActual);
                        celda1.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    }
                    finally { ExcelHelper.LiberarCom(celda1); }

                    string dir1 = $"{letraC1}{filaActual}";
                    string form2 = r == 1 ? $"={dir1}" : $"=IF({dir1}=\"\", {letraC2}{filaActual - 1}, IF({letraC2}{filaActual - 1}=\"\", {dir1}, {letraC2}{filaActual - 1} & \", \" & {dir1}))";
                    Excel.Range celda2 = null;
                    try
                    {
                        celda2 = (Excel.Range)wsActual.Cells[filaActual, c2];
                        celda2.FormulaLocal = ExcelHelper.TraducirFormulaLocal(wsActual, form2, filaActual);
                        celda2.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    }
                    finally { ExcelHelper.LiberarCom(celda2); }
                }

                // 4. INYECCIÓN ALERTA FINAL
                string ultimaCelda = $"{letraC2}{filaInicial + filasTotales - 1}";
                string formAlerta = $"=IF({ultimaCelda}=\"\", \"\", IF(ISNUMBER(SEARCH(\",\", {ultimaCelda})), \"Favor de revisar la información capturada en los numerales: \" & SUBSTITUTE({ultimaCelda}, \", \", \" y \", LEN({ultimaCelda})-LEN(SUBSTITUTE({ultimaCelda}, \",\", \"\"))), \"Favor de revisar la información capturada en el numeral: \" & {ultimaCelda}))";

                if (rangoAlerta.Count > 1 && !(rangoAlerta.MergeCells is bool && (bool)rangoAlerta.MergeCells)) rangoAlerta.Merge();
                rangoAlerta.Font.Name = "Arial";
                rangoAlerta.Font.Size = 10;
                rangoAlerta.Font.Bold = true;
                rangoAlerta.Font.Color = ColorTranslator.ToOle(Color.FromArgb(0, 112, 192));
                rangoAlerta.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                rangoAlerta.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                rangoAlerta.FormulaLocal = ExcelHelper.TraducirFormulaLocal(wsActual, formAlerta, filaInicial);

                string numeroPregunta = ExcelHelper.ObtenerNumeroPregunta(rangoCapturado);
                AuditoriaCenso.RegistrarAccion(libroCenso, numeroPregunta, "Validación Blancos Sin Código (Multifila)", rangoCapturado.Address.Replace("$", ""), "Auxiliares + Alerta Dinámica");
                return new ResultadoValidacion { Exito = true, Mensaje = "Proceso multifila aplicado con éxito.", AlertaInyectada = true };
            }
            catch (Exception ex)
            {
                return new ResultadoValidacion { Exito = false, Mensaje = "Error: " + ex.Message, AlertaInyectada = false };
            }
            finally
            {
                ExcelHelper.LiberarCom(rangoAlerta); ExcelHelper.LiberarCom(rangoAux);
                ExcelHelper.LiberarCom(rangoNumerales); ExcelHelper.LiberarCom(wsActual);
                if (excelApp != null) excelApp.ScreenUpdating = true;
            }
        }
        #endregion
    }
}