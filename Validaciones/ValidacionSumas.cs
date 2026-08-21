using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using SAVCNG_ExcelDNA.Core; // Consumo estricto de Fachada y DTO

namespace SAVCNG_ExcelDNA.Validaciones
{
    public class ValidacionSumas : IValidacionExcel
    {
        // 1. EL CONTRATO EXIGE DEVOLVER EL DTO (DUMB VIEW PATTERN)
        public ResultadoValidacion Ejecutar(Excel.Application excelApp, Excel.Workbook libroCenso, Excel.Range rangoCapturado)
        {
            Excel.Range rangoTotal = null;
            Excel.Range rangoDesagregados = null;
            Excel.Range rangoAlerta = null;
            Excel.Range seleccionAuxiliar = null;
            Excel.Range rangoTotalesVerticales = null;
            Excel.Worksheet wsActual = null;

            try
            {
                int filaInicio = rangoCapturado.Row;
                int filaFin = filaInicio + rangoCapturado.Rows.Count - 1;

                // =========================================================================
                // FASE 1: RECOPILACIÓN DE RANGOS DE TRABAJO (DTO BLINDADO)
                // =========================================================================
                object resTotal = excelApp.InputBox("1. Selecciona la COLUMNA del TOTAL o PIVOTE:", "SAVCNG - Total", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8);
                if (resTotal is bool && (bool)resTotal == false) return new ResultadoValidacion { Exito = false, Mensaje = "Selección del Total cancelada.", AlertaInyectada = false };
                rangoTotal = (Excel.Range)resTotal;

                // NUEVO: Aplicamos formato de miles nativo. No afecta textos como "NS" o "NA".
                rangoTotal.NumberFormat = "#,##0";

                object resDesagregados = excelApp.InputBox("2. Selecciona las COLUMNAS de los DESAGREGADOS (Usa CTRL para varias por separado):", "SAVCNG - Desagregados", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8);
                if (resDesagregados is bool && (bool)resDesagregados == false) return new ResultadoValidacion { Exito = false, Mensaje = "Selección de Desagregados cancelada.", AlertaInyectada = false };
                rangoDesagregados = (Excel.Range)resDesagregados;

                // NUEVO: Aplicamos formato de miles a los desagregados
                rangoDesagregados.NumberFormat = "#,##0";

                object resAlerta = excelApp.InputBox("3. Selecciona el rango destino para el MENSAJE DE ERROR", "SAVCNG - Alerta Global", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8);
                if (resAlerta is bool && (bool)resAlerta == false) return new ResultadoValidacion { Exito = false, Mensaje = "Selección de rango de alerta cancelada.", AlertaInyectada = false };
                rangoAlerta = (Excel.Range)resAlerta;

                object resAuxiliar = excelApp.InputBox("4. Selecciona UNA CELDA en una columna libre (ej. AF) para insertar formulas auxiliares.\nADVERTENCIA: Considere un espacio libre N filas. (N = numerales de la pregunta o rango seleccionado)", "SAVCNG - Arquitectura de Memoria", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8);
                if (resAuxiliar is bool && (bool)resAuxiliar == false) return new ResultadoValidacion { Exito = false, Mensaje = "Selección de celda auxiliar cancelada.", AlertaInyectada = false };
                seleccionAuxiliar = (Excel.Range)resAuxiliar;

                // =========================================================================
                // FASE 1.5: RECOLECCIÓN LIGERA DE PREGUNTAS (STATE MANAGEMENT)
                // =========================================================================
                System.Collections.Generic.HashSet<string> preguntasUnicas = new System.Collections.Generic.HashSet<string>();
                int totalAreasCap = rangoCapturado.Areas.Count;

                // RASTREO ZERO LEAKS
                for (int i = 1; i <= totalAreasCap; i++)
                {
                    Excel.Range areaInd = null;
                    try
                    {
                        areaInd = (Excel.Range)rangoCapturado.Areas[i];
                        string idPregunta = ExcelHelper.ObtenerNumeroPregunta(areaInd);
                        if (!string.IsNullOrEmpty(idPregunta)) { preguntasUnicas.Add(idPregunta); }
                    }
                    finally { ExcelHelper.LiberarCom(areaInd); }
                }
                string preguntasDetectadas = preguntasUnicas.Count > 0 ? string.Join(", ", preguntasUnicas) : "ND";

                excelApp.ScreenUpdating = false;
                wsActual = (Excel.Worksheet)rangoCapturado.Worksheet;

                // =========================================================================
                // A. DETECCIÓN INTELIGENTE DE LETRAS (TOTAL)
                // =========================================================================
                Excel.Range primeraCeldaTotal = null;
                Excel.Range mergeTotal = null;
                Excel.Range topLeftTotal = null;
                string letraTotal = "";

                try
                {
                    primeraCeldaTotal = (Excel.Range)rangoTotal.Cells[1, 1];
                    mergeTotal = primeraCeldaTotal.MergeArea;
                    topLeftTotal = (Excel.Range)mergeTotal.Cells[1, 1];
                    letraTotal = topLeftTotal.Address.Split('$')[1];
                }
                finally
                {
                    ExcelHelper.LiberarCom(topLeftTotal);
                    ExcelHelper.LiberarCom(mergeTotal);
                    ExcelHelper.LiberarCom(primeraCeldaTotal);
                }

                Excel.Range celdaInicioAux = null;
                string colAuxLetra = "";
                try
                {
                    celdaInicioAux = (Excel.Range)seleccionAuxiliar.Cells[1, 1];
                    colAuxLetra = celdaInicioAux.Address.Split('$')[1];
                }
                finally
                {
                    ExcelHelper.LiberarCom(celdaInicioAux);
                }

                // =========================================================================
                // B. DETECCIÓN INTELIGENTE DE CELDAS COMBINADAS (DESAGREGADOS)
                // =========================================================================
                System.Collections.Generic.List<string> letrasDesagregados = new System.Collections.Generic.List<string>();
                int totalAreasDesagregados = rangoDesagregados.Areas.Count;

                for (int i = 1; i <= totalAreasDesagregados; i++)
                {
                    Excel.Range areaDesagregado = null;
                    Excel.Range fila1 = null;
                    try
                    {
                        areaDesagregado = (Excel.Range)rangoDesagregados.Areas[i];
                        fila1 = (Excel.Range)areaDesagregado.Rows[1];
                        int totalCeldasFila = fila1.Cells.Count;

                        for (int j = 1; j <= totalCeldasFila; j++)
                        {
                            Excel.Range celda = null;
                            Excel.Range mArea = null;
                            Excel.Range tl = null;
                            try
                            {
                                celda = (Excel.Range)fila1.Cells[j];
                                mArea = celda.MergeArea;
                                tl = (Excel.Range)mArea.Cells[1, 1];
                                string letra = tl.Address.Split('$')[1];

                                if (!letrasDesagregados.Contains(letra))
                                {
                                    letrasDesagregados.Add(letra);
                                }
                            }
                            finally
                            {
                                ExcelHelper.LiberarCom(tl);
                                ExcelHelper.LiberarCom(mArea);
                                ExcelHelper.LiberarCom(celda);
                            }
                        }
                    }
                    finally
                    {
                        ExcelHelper.LiberarCom(fila1);
                        ExcelHelper.LiberarCom(areaDesagregado);
                    }
                }

                int numDesagregados = letrasDesagregados.Count;

                // =========================================================================
                // C. MÁQUINA DE ESTADOS (INSERCIÓN DE MOTOR AUXILIAR XP READY)
                // =========================================================================
                for (int f = filaInicio; f <= filaFin; f++)
                {
                    string tRef = $"${letraTotal}{f}";

                    string countNS = string.Join("+", letrasDesagregados.ConvertAll(l => $"COUNTIF(${l}{f},\"NS\")"));
                    string countNA = string.Join("+", letrasDesagregados.ConvertAll(l => $"COUNTIF(${l}{f},\"NA\")"));

                    string sumaDesagregados = $"SUM({string.Join(",", letrasDesagregados.ConvertAll(l => $"${l}{f}"))})";
                    string countNum = $"COUNT({string.Join(",", letrasDesagregados.ConvertAll(l => $"${l}{f}"))})";

                    string condAritmetica = $"AND(ISNUMBER({tRef}), OR(({countNS})=0, {countNum}>0), {tRef}<>{sumaDesagregados})";
                    string condCero = $"AND(ISNUMBER({tRef}), {tRef}=0, ({countNS})>0)";
                    string condNS = $"AND(UPPER(TRIM({tRef}))=\"NS\", OR({countNum}={numDesagregados}, {sumaDesagregados}>0))";

                    // Validación estricta "Todo o Nada" para NA
                    string isTotalNA = $"IF(UPPER(TRIM({tRef}))=\"NA\", 1, 0)";
                    string totalNAs = $"({isTotalNA} + {countNA})";
                    string condNA = $"AND({totalNAs} > 0, {totalNAs} <> {numDesagregados + 1})";

                    // Ensamblaje final del motor evaluador
                    string formulaInglesAux = $"=IF(OR({condAritmetica}, {condCero}, {condNS}, {condNA}), 1, 0)";

                    Excel.Range celdaDestinoAux = null;
                    try
                    {
                        celdaDestinoAux = wsActual.Range[$"{colAuxLetra}{f}"];
                        celdaDestinoAux.Formula = formulaInglesAux;
                    }
                    finally
                    {
                        ExcelHelper.LiberarCom(celdaDestinoAux);
                    }
                }

                // =========================================================================
                // E. INSERCIÓN DEL MENSAJE DE ALERTA GLOBAL (RADAR AUTOMÁTICO XP READY)
                // =========================================================================
                string auxRange = $"${colAuxLetra}${filaInicio}:${colAuxLetra}${filaFin}";

                // FÓRMULA RADAR GRAMATICAL: Discrimina entre Singular y Plural de forma nativa
                string formulaAlertaFinal = $"=IF(SUM({auxRange})=0, \"\",IF(SUM({auxRange})=1, \"Se detectó 1 error\", \"Se detectaron \" & SUM({auxRange}) & \" errores\") & \" de inconsistencia en sumas. Revisar a partir del numeral \" & MATCH(1, {auxRange}, 0))";

                if (rangoAlerta.Count > 1) { rangoAlerta.Merge(); }
                rangoAlerta.Formula = formulaAlertaFinal;
                rangoAlerta.Font.Name = "Arial";
                rangoAlerta.Font.Size = 9;
                rangoAlerta.Font.Bold = true;
                rangoAlerta.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                rangoAlerta.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                rangoAlerta.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                rangoAlerta.WrapText = true;

                // =========================================================================
                // F. INSERCIÓN OPTIMIZADA DE SUMATORIAS VERTICALES (SIGMAS) - ¡PASO FINAL!
                // =========================================================================
                excelApp.ScreenUpdating = true;

                DialogResult respuestaSigma = MessageBox.Show(
                    "¿Deseas agregar la validación de sumatoria vertical (sigma)?",
                    "SAVCNG - Sumatorias Verticales",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question
                );

                string mensajeExitoFinal = "";
                string funcionalidadAudit = "Fórmulas (Auxiliar + Radar Dinámico)";

                if (respuestaSigma == DialogResult.Yes)
                {
                    object resVerticales = excelApp.InputBox("5. Selecciona la FILA de Sumatoria Vertical Sigma (Σ):", "SAVCNG - Sigma", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8);

                    if (resVerticales is bool && (bool)resVerticales == false)
                    {
                        return new ResultadoValidacion { Exito = false, Mensaje = "Selección de fila Sigma cancelada. La validación base fue aplicada, pero el proceso fue interrumpido.", AlertaInyectada = false };
                    }

                    rangoTotalesVerticales = (Excel.Range)resVerticales;
                    excelApp.ScreenUpdating = false;

                    System.Collections.Generic.List<string> sigmasProcesados = new System.Collections.Generic.List<string>();
                    int totalSigmas = rangoTotalesVerticales.Cells.Count;

                    for (int k = 1; k <= totalSigmas; k++)
                    {
                        Excel.Range celdaSumatoria = null;
                        Excel.Range mArea = null;
                        Excel.Range tl = null;

                        try
                        {
                            celdaSumatoria = (Excel.Range)rangoTotalesVerticales.Cells[k];
                            mArea = celdaSumatoria.MergeArea;
                            tl = (Excel.Range)mArea.Cells[1, 1];
                            string addr = tl.Address;

                            if (!sigmasProcesados.Contains(addr))
                            {
                                sigmasProcesados.Add(addr);
                                string lCol = addr.Split('$')[1];
                                string rSuma = $"{lCol}{filaInicio}:{lCol}{filaFin}";

                                string fSigma = $"=IF(AND(SUM({rSuma})=0,COUNTIF({rSuma},\"NS\")>0),\"NS\"," +
                                                $"IF(AND(SUM({rSuma})=0,COUNTIF({rSuma},0)>0),0," +
                                                $"IF(AND(SUM({rSuma})=0,COUNTIF({rSuma},\"NA\")>0),\"NA\"," +
                                                $"SUM({rSuma}))))";
                                tl.Formula = fSigma;

                                // REQUERIMIENTO CUMPLIDO: Formato Negritas (Bold) y separación por miles
                                tl.Font.Bold = true;
                                tl.NumberFormat = "#,##0";
                            }
                        }
                        finally
                        {
                            ExcelHelper.LiberarCom(tl);
                            ExcelHelper.LiberarCom(mArea);
                            ExcelHelper.LiberarCom(celdaSumatoria);
                        }
                    }

                    mensajeExitoFinal = "Validación de Sumas y Sumatorias Verticales (Sigma) insertadas con éxito.";
                    funcionalidadAudit += " + Sigma Vertical";
                }
                else
                {
                    mensajeExitoFinal = "Validaciones de sumas insertadas con éxito (Se omitió la Sumatoria Vertical).";
                }

                // =========================================================================
                // FASE FINAL: REGISTRO EN BITÁCORA
                // =========================================================================
                AuditoriaCenso.RegistrarAccion(
                    libroCenso,
                    preguntasDetectadas,
                    "Consistencia Aritmética (Sumas)",
                    rangoCapturado.Address.Replace("$", ""),
                    funcionalidadAudit
                );

                return new ResultadoValidacion
                {
                    Exito = true,
                    Mensaje = mensajeExitoFinal,
                    AlertaInyectada = true
                };
            }
            catch (Exception ex)
            {
                return new ResultadoValidacion
                {
                    Exito = false,
                    Mensaje = "Error crítico en la validación de sumas: " + ex.Message,
                    AlertaInyectada = false
                };
            }
            finally
            {
                // Limpieza absoluta de referencias COM base (Zero Leaks Estricto)
                ExcelHelper.LiberarCom(rangoTotal);
                ExcelHelper.LiberarCom(rangoDesagregados);
                ExcelHelper.LiberarCom(rangoTotalesVerticales);
                ExcelHelper.LiberarCom(rangoAlerta);
                ExcelHelper.LiberarCom(seleccionAuxiliar);
                ExcelHelper.LiberarCom(wsActual);

                if (excelApp != null) excelApp.ScreenUpdating = true;
            }
        }
    }
}