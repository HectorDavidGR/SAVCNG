using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using SAVCNG_ExcelDNA.Core; // Consumimos nuestra Fachada y el DTO

namespace SAVCNG_ExcelDNA.Validaciones
{
    public class ValidacionSumas : IValidacionExcel
    {
        // 1. EL CONTRATO EXIGE DEVOLVER EL DTO
        public ResultadoValidacion Ejecutar(Excel.Application excelApp, Excel.Workbook libroCenso, Excel.Range rangoCapturado)
        {
            Excel.Range rangoTotal = null;
            Excel.Range rangoDesagregados = null;
            Excel.Range rangoTotalesVerticales = null;
            Excel.Range rangoAlerta = null;
            Excel.Range seleccionAuxiliar = null;
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

                object resDesagregados = excelApp.InputBox("2. Selecciona las COLUMNAS de los DESAGREGADOS (Usa CTRL para varias):", "SAVCNG - Desagregados", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8);
                if (resDesagregados is bool && (bool)resDesagregados == false) return new ResultadoValidacion { Exito = false, Mensaje = "Selección de Desagregados cancelada.", AlertaInyectada = false };
                rangoDesagregados = (Excel.Range)resDesagregados;

                object resVerticales = excelApp.InputBox("3. Selecciona la FILA de Sumatoria Vertical Sigma (Σ):", "SAVCNG - Sigma", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8);
                if (resVerticales is bool && (bool)resVerticales == false) return new ResultadoValidacion { Exito = false, Mensaje = "Selección de fila Sigma cancelada.", AlertaInyectada = false };
                rangoTotalesVerticales = (Excel.Range)resVerticales;

                object resAlerta = excelApp.InputBox("4. Selecciona el rango destino para el MENSAJE DE ERROR:", "SAVCNG - Alerta Global", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8);
                if (resAlerta is bool && (bool)resAlerta == false) return new ResultadoValidacion { Exito = false, Mensaje = "Selección de rango de alerta cancelada.", AlertaInyectada = false };
                rangoAlerta = (Excel.Range)resAlerta;

                object resAuxiliar = excelApp.InputBox("5. Selecciona UNA CELDA en una columna libre (ej. AF) para inyectar el Motor Auxiliar:", "SAVCNG - Arquitectura de Memoria", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8);
                if (resAuxiliar is bool && (bool)resAuxiliar == false) return new ResultadoValidacion { Exito = false, Mensaje = "Selección de celda auxiliar cancelada.", AlertaInyectada = false };
                seleccionAuxiliar = (Excel.Range)resAuxiliar;

                // =========================================================================
                // FASE 1.5: RECOLECCIÓN LIGERA DE PREGUNTAS (STATE MANAGEMENT)
                // =========================================================================
                System.Collections.Generic.HashSet<string> preguntasUnicas = new System.Collections.Generic.HashSet<string>();
                int totalAreasCap = rangoCapturado.Areas.Count;

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
                // A. DETECCIÓN INTELIGENTE DE CELDAS COMBINADAS (TOTAL)
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

                // RASTREO ZERO LEAKS: Reemplazo de los dos foreach anidados
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
                // C. MÁQUINA DE ESTADOS (INYECCIÓN DE MOTOR AUXILIAR FILA POR FILA)
                // =========================================================================
                for (int f = filaInicio; f <= filaFin; f++)
                {
                    string tRef = $"${letraTotal}{f}";

                    // Motores de análisis dinámicos
                    string countNS = string.Join("+", letrasDesagregados.ConvertAll(l => $"COUNTIF(${l}{f},\"NS\")"));
                    string sumaDesagregados = $"SUM({string.Join(",", letrasDesagregados.ConvertAll(l => $"${l}{f}"))})";
                    string countNum = $"COUNT({string.Join(",", letrasDesagregados.ConvertAll(l => $"${l}{f}"))})";

                    string condAritmetica = $"AND(ISNUMBER({tRef}), ({countNS})=0, {tRef}<>{sumaDesagregados})";
                    string condCero = $"AND(ISNUMBER({tRef}), {tRef}=0, ({countNS})>0)";
                    string condNS = $"AND(UPPER(TRIM({tRef}))=\"NS\", OR({countNum}={numDesagregados}, {sumaDesagregados}>0))";

                    string formulaInglesAux = $"=IF(OR({condAritmetica}, {condCero}, {condNS}), 1, 0)";

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
                // D. INYECCIÓN DEL FORMATO CONDICIONAL (VINCULADO AL MOTOR)
                // =========================================================================
                if (rangoCapturado.FormatConditions.Count > 0)
                {
                    excelApp.ScreenUpdating = true;
                    DialogResult respFormato = MessageBox.Show("Se detectaron reglas previas.\n¿Deseas CONSERVARLAS e integrar esta nueva capa?\nSÍ = Apilar\nNO = Borrar", "SAVCNG - Coexistencia", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);
                    excelApp.ScreenUpdating = false;

                    // ABORTO TOTAL 6: DTO
                    if (respFormato == DialogResult.Cancel) return new ResultadoValidacion { Exito = false, Mensaje = "Operación cancelada en la evaluación de formatos previos.", AlertaInyectada = false };

                    if (respFormato == DialogResult.No) { rangoCapturado.FormatConditions.Delete(); }
                }

                // Traducción de Formato Condicional vía Façade
                string fSombreadoIngles = $"=${colAuxLetra}{filaInicio}=1";
                string fcLocal = ExcelHelper.TraducirFormulaLocal(wsActual, fSombreadoIngles, filaInicio);

                Excel.FormatCondition fcError = null;
                try
                {
                    fcError = (Excel.FormatCondition)rangoCapturado.FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Type.Missing, fcLocal);
                    fcError.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 199, 206));
                    fcError.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(156, 0, 6));
                    fcError.StopIfTrue = false;
                }
                finally
                {
                    ExcelHelper.LiberarCom(fcError);
                }

                // =========================================================================
                // E. INYECCIÓN DEL MENSAJE DE ALERTA GLOBAL
                // =========================================================================
                string auxRange = $"${colAuxLetra}${filaInicio}:${colAuxLetra}${filaFin}";
                string formulaAlertaFinal = $"=IF(SUM({auxRange})>0, \"Favor de revisar la suma y consistencia de totales y/o subtotales por filas (numéricos y NS)\", \"\")";

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
                // F. INYECCIÓN OPTIMIZADA DE SUMATORIAS VERTICALES (SIGMAS)
                // =========================================================================
                System.Collections.Generic.List<string> sigmasProcesados = new System.Collections.Generic.List<string>();
                int totalSigmas = rangoTotalesVerticales.Cells.Count;

                // RASTREO ZERO LEAKS: Reemplazo de foreach
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
                            tl.Formula = fSigma; // Las sumatorias básicas se inyectan en inglés directo sin problemas
                        }
                    }
                    finally
                    {
                        ExcelHelper.LiberarCom(tl);
                        ExcelHelper.LiberarCom(mArea);
                        ExcelHelper.LiberarCom(celdaSumatoria);
                    }
                }

                // =========================================================================
                // FASE FINAL: AUDITORÍA Y NOTIFICACIÓN (DTO)
                // =========================================================================
                AuditoriaCenso.RegistrarAccion(
                    libroCenso,
                    preguntasDetectadas,
                    "Consistencia Aritmética (Sumas)",
                    rangoCapturado.Address.Replace("$", ""),
                    "Fórmulas + FormatConditions"
                );

                // ÉXITO TOTAL: DTO
                return new ResultadoValidacion
                {
                    Exito = true,
                    Mensaje = "Motor de Validación, Alerta y Sumatorias inyectados con éxito (Soporte Merge).",
                    AlertaInyectada = true
                };
            }
            catch (Exception ex)
            {
                // ERROR CRÍTICO: DTO
                return new ResultadoValidacion
                {
                    Exito = false,
                    Mensaje = "Error crítico en el motor de sumas: " + ex.Message,
                    AlertaInyectada = false
                };
            }
            finally
            {
                // Limpieza absoluta de referencias COM base
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